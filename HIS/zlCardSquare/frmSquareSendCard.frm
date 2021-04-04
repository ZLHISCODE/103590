VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareSendCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "消费卡发放"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9840
   Icon            =   "frmSquareSendCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   4
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9735
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   5190
      Width           =   9765
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   903
         Left            =   4665
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   904
         Left            =   1140
         TabIndex        =   39
         Top             =   540
         Width           =   2265
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   906
         Left            =   7920
         PasswordChar    =   "*"
         TabIndex        =   43
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   905
         Left            =   4665
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   540
         Width           =   1725
      End
      Begin zlIDKind.IDKindNew IDKind补卡 
         Height          =   300
         Left            =   1140
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   529
         Appearance      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   902
         Left            =   1710
         TabIndex        =   35
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox cbo原卡卡号 
         Height          =   300
         Left            =   4665
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   901
         Left            =   1140
         TabIndex        =   30
         Top             =   120
         Width           =   2265
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "新卡号(&C)"
         Height          =   180
         Index           =   905
         Left            =   300
         TabIndex        =   38
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码(&W)"
         Height          =   180
         Index           =   906
         Left            =   4020
         TabIndex        =   40
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "确认密码(&V)"
         Height          =   180
         Index           =   907
         Left            =   6870
         TabIndex        =   42
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "原卡号(&O)"
         Height          =   180
         Index           =   901
         Left            =   300
         TabIndex        =   29
         Top             =   180
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&P)"
         Height          =   180
         Index           =   902
         Left            =   480
         TabIndex        =   33
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "原卡密码(&P)"
         Height          =   180
         Index           =   903
         Left            =   3660
         TabIndex        =   31
         Top             =   180
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "原卡号(&O)"
         Height          =   180
         Index           =   904
         Left            =   3840
         TabIndex        =   36
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Index           =   4
      Left            =   0
      TabIndex        =   60
      Top             =   6090
      Width           =   9825
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   2
         Left            =   2670
         ScaleHeight     =   1425
         ScaleWidth      =   7095
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   60
         Width           =   7125
         Begin VB.TextBox txt 
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
            Index           =   403
            Left            =   720
            TabIndex        =   48
            Top             =   578
            Width           =   2445
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            IMEMode         =   3  'DISABLE
            Index           =   405
            Left            =   720
            TabIndex        =   50
            Tag             =   "1"
            Top             =   1020
            Width           =   2445
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            Index           =   406
            Left            =   4290
            TabIndex        =   51
            Tag             =   "1"
            Top             =   1020
            Width           =   2760
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            Index           =   404
            Left            =   4290
            TabIndex        =   49
            Tag             =   "1"
            Top             =   600
            Width           =   2760
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   401
            Left            =   1770
            TabIndex        =   46
            Top             =   150
            Width           =   1410
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   402
            Left            =   4290
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   150
            Width           =   2760
         End
         Begin VB.ComboBox cbo支付方式 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   150
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "缴款人"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   404
            Left            =   60
            TabIndex        =   100
            Top             =   630
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "帐  号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   406
            Left            =   60
            TabIndex        =   67
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算号码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   407
            Left            =   3420
            TabIndex        =   66
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开户行"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   405
            Left            =   3420
            TabIndex        =   65
            Top             =   630
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "收 款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   402
            Left            =   60
            TabIndex        =   64
            Top             =   195
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "找 补"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   403
            Left            =   3690
            TabIndex        =   63
            Top             =   195
            Width           =   570
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   1
         Left            =   30
         ScaleHeight     =   1425
         ScaleWidth      =   2535
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   60
         Width           =   2565
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   401
            Left            =   1665
            TabIndex        =   62
            Top             =   720
            Width           =   780
         End
         Begin XtremeSuiteControls.ShortcutCaption lbl缴款 
            Height          =   375
            Left            =   10
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   10
            Width           =   2535
            _Version        =   589884
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   6
            Caption         =   "收款合计"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.99
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Index           =   1
      Left            =   0
      TabIndex        =   61
      Top             =   7590
      Width           =   10095
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   210
         TabIndex        =   54
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   7290
         TabIndex        =   52
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8460
         TabIndex        =   53
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "本次误差：0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   5010
         TabIndex        =   101
         Top             =   300
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6045
      Index           =   3
      Left            =   30
      ScaleHeight     =   6015
      ScaleWidth      =   9735
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   30
      Width           =   9765
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Index           =   5
         Left            =   7440
         TabIndex        =   95
         Top             =   1290
         Width           =   2265
         Begin VB.PictureBox pic 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "frmSquareSendCard.frx":000C
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   57
            Top             =   3390
            Width           =   495
         End
         Begin VB.Frame fra 
            Caption         =   "    余额"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1185
            Index           =   52
            Left            =   0
            TabIndex        =   97
            Top             =   3525
            Width           =   2250
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   601
               Left            =   1470
               TabIndex        =   98
               Top             =   540
               Width           =   660
            End
         End
         Begin VB.Frame fra 
            Caption         =   "限制类别"
            Height          =   3330
            Index           =   51
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   2250
            Begin MSComctlLib.ListView lvw限制类别 
               Height          =   3015
               Left            =   60
               TabIndex        =   44
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   5318
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Key             =   "类型"
                  Object.Tag             =   "类型"
                  Text            =   "类型"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1245
         Index           =   2
         Left            =   10
         TabIndex        =   68
         Top             =   10
         Width           =   9705
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   201
            Left            =   3165
            TabIndex        =   1
            Top             =   90
            Width           =   2055
         End
         Begin VB.ComboBox cbo卡类型 
            Height          =   300
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   202
            Left            =   5490
            TabIndex        =   2
            Top             =   90
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   300
            Left            =   7650
            Picture         =   "frmSquareSendCard.frx":198E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   90
            Width           =   300
         End
         Begin VB.CommandButton cmdDelete 
            Enabled         =   0   'False
            Height          =   300
            Left            =   8025
            Picture         =   "frmSquareSendCard.frx":81E0
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   90
            Width           =   300
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfCardNo 
            Height          =   735
            Left            =   30
            TabIndex        =   5
            Top             =   450
            Width           =   9675
            _cx             =   17066
            _cy             =   1296
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483643
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   0
            Cols            =   0
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
         Begin VB.Line lineSplit 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9735
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "卡号(&N)"
            Height          =   180
            Index           =   202
            Left            =   2490
            TabIndex        =   73
            Top             =   150
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "卡类型(&T)"
            Height          =   180
            Index           =   201
            Left            =   60
            TabIndex        =   72
            Top             =   150
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Index           =   203
            Left            =   5280
            TabIndex        =   71
            Top             =   150
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "300"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   205
            Left            =   8880
            TabIndex        =   69
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "共发    张"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   204
            Left            =   8430
            TabIndex        =   70
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Frame fra 
         Caption         =   "卡面值"
         Height          =   705
         Index           =   7
         Left            =   60
         TabIndex        =   92
         Top             =   4080
         Width           =   7275
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   702
            Left            =   4965
            TabIndex        =   24
            Top             =   270
            Width           =   2265
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   701
            Left            =   1080
            TabIndex        =   23
            Top             =   270
            Width           =   2265
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "实际销售额(&J)"
            Height          =   180
            Index           =   702
            Left            =   3720
            TabIndex        =   94
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "卡面额(&M)"
            Height          =   180
            Index           =   701
            Left            =   240
            TabIndex        =   93
            Top             =   330
            Width           =   810
         End
      End
      Begin VB.Frame fra 
         Caption         =   "充值信息"
         Height          =   1170
         Index           =   8
         Left            =   60
         TabIndex        =   86
         Top             =   4830
         Width           =   7275
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   803
            Left            =   6090
            TabIndex        =   27
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   802
            Left            =   3510
            TabIndex        =   26
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   801
            Left            =   1080
            TabIndex        =   25
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   315
            Index           =   804
            Left            =   1080
            TabIndex        =   28
            Top             =   720
            Width           =   6090
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "本次充值(&B)"
            Height          =   180
            Index           =   803
            Left            =   2490
            TabIndex        =   91
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "充值扣率(&K)"
            Height          =   180
            Index           =   801
            Left            =   45
            TabIndex        =   90
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "实际充值缴款(&I)"
            Height          =   180
            Index           =   804
            Left            =   4710
            TabIndex        =   89
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   802
            Left            =   2220
            TabIndex        =   88
            Top             =   345
            Width           =   120
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "充值说明(&Z)"
            Height          =   180
            Index           =   805
            Left            =   45
            TabIndex        =   87
            Top             =   795
            Width           =   990
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "发卡基本信息"
         Height          =   2835
         Index           =   3
         Left            =   10
         TabIndex        =   74
         Top             =   1320
         Width           =   7425
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   300
            Left            =   1095
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1245
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            Appearance      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "宋体"
            IDKind          =   -1
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   301
            Left            =   1095
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   60
            Width           =   2295
         End
         Begin VB.CheckBox chk充值 
            Caption         =   "允许充值"
            Height          =   180
            Left            =   1095
            TabIndex        =   8
            Top             =   510
            Width           =   1200
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   311
            Left            =   5040
            TabIndex        =   22
            Top             =   2460
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   310
            Left            =   1095
            TabIndex        =   21
            Top             =   2460
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   308
            Left            =   1095
            TabIndex        =   19
            Top             =   2055
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   307
            Left            =   1095
            TabIndex        =   18
            Top             =   1650
            Width           =   6225
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   270
            Index           =   2
            Left            =   7035
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "领卡部门"
            Top             =   1230
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   270
            Index           =   0
            Left            =   7035
            TabIndex        =   12
            TabStop         =   0   'False
            Tag             =   "发卡原因"
            Top             =   825
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   270
            Index           =   1
            Left            =   3090
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "领卡人"
            Top             =   1260
            Width           =   285
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   304
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   11
            Top             =   825
            Width           =   6225
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   306
            Left            =   5040
            TabIndex        =   16
            Top             =   1230
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   309
            Left            =   5040
            TabIndex        =   20
            Top             =   2055
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtp卡有效日期 
            Height          =   300
            Left            =   5040
            TabIndex        =   10
            Top             =   450
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   186318851
            CurrentDate     =   40156.0854282407
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   303
            Left            =   5040
            TabIndex        =   9
            Top             =   450
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   305
            Left            =   1650
            TabIndex        =   14
            Top             =   1245
            Width           =   1725
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   302
            Left            =   5040
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "确认密码(&E)"
            Height          =   180
            Index           =   302
            Left            =   3990
            TabIndex        =   85
            Top             =   120
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "密码(&W)"
            Height          =   180
            Index           =   301
            Left            =   450
            TabIndex        =   84
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "回收日期"
            Height          =   180
            Index           =   311
            Left            =   4230
            TabIndex        =   83
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "回收人"
            Height          =   180
            Index           =   310
            Left            =   540
            TabIndex        =   82
            Top             =   2520
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "领卡部门(&M)"
            Height          =   180
            Index           =   306
            Left            =   3990
            TabIndex        =   81
            Top             =   1290
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "发卡日期"
            Height          =   180
            Index           =   309
            Left            =   4260
            TabIndex        =   80
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "发卡人"
            Height          =   180
            Index           =   308
            Left            =   540
            TabIndex        =   79
            Top             =   2115
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "备注(&S)"
            Height          =   180
            Index           =   307
            Left            =   450
            TabIndex        =   78
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "卡有效日期(&D)"
            Height          =   180
            Index           =   303
            Left            =   3810
            TabIndex        =   77
            Top             =   510
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "领卡人(&D)"
            Height          =   180
            Index           =   305
            Left            =   270
            TabIndex        =   76
            Top             =   1305
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "发卡原因(&Y)"
            Height          =   180
            Index           =   304
            Left            =   90
            TabIndex        =   75
            Top             =   870
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frmSquareSendCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Private mfrmMain As Form, mlngModule As Long, mstrPrivs As String
Private mEditType As gCardEditType
Private mlng卡类别 As Long, mlng卡ID As Long
Private mlng充值ID As Long

'控件索引枚举值
Private Enum mFrameIndex
    fra_卡号 = 2
    fra_卡信息 = 3
    fra_卡面值 = 7
    fra_充值信息 = 8
    fra_限制类别 = 51
    fra_类别余额 = 5
    fra_余额 = 52
    fra_缴款面板 = 4
    fra_按钮面板 = 1
End Enum
Private Enum mLableIndex
    lbl_卡类型 = 201
    lbl_卡号 = 202
    lbl_卡号至 = 203
    lbl_卡张数 = 204
    lbl_卡张数2 = 205
    lbl_密码 = 301
    lbl_确认密码 = 302
    lbl_卡有效期 = 303
    lbl_发卡原因 = 304
    lbl_领卡人 = 305
    lbl_领卡部门 = 306
    lbl_备注 = 307
    lbl_发卡人 = 308
    lbl_发卡日期 = 309
    lbl_回收人 = 310
    lbl_回收日期 = 311
    lbl_卡面额 = 701
    lbl_销售额 = 702
    
    lbl_缴款合计 = 401
    lbl_支付方式 = 402
    lbl_找补 = 403
    
    lbl_误差 = 0
    
    lbl_原卡卡号 = 901
    lbl_补卡姓名 = 902
    lbl_原卡密码 = 903
    lbl_补卡原卡 = 904
    lbl_新卡卡号 = 905
    lbl_新卡密码 = 906
    lbl_新卡密码确认 = 907
    lbl_卡余额 = 601
End Enum
Private Enum mTextIndex
    txt_开始卡号 = 201
    txt_结束卡号 = 202
    txt_密码 = 301
    txt_确认密码 = 302
    txt_卡有效期 = 303
    txt_发卡原因 = 304
    txt_领卡人 = 305
    txt_领卡部门 = 306
    txt_备注 = 307
    txt_发卡人 = 308
    txt_发卡日期 = 309
    txt_回收人 = 310
    txt_回收日期 = 311
    txt_卡面额 = 701
    txt_销售额 = 702
    
    txt_充值扣率 = 801
    txt_本次充值 = 802
    txt_充值缴款 = 803
    txt_充值说明 = 804
    
    txt_缴款 = 401
    txt_找补 = 402
    txt_缴款人 = 403
    txt_开户行 = 404
    txt_帐号 = 405
    txt_结算号码 = 406
    
    txt_原卡卡号 = 901
    txt_补卡姓名 = 902
    txt_原卡密码 = 903
    txt_新卡卡号 = 904
    txt_新卡密码 = 905
    txt_新卡确认密码 = 906
End Enum
Private Enum mCommandButtonIndex
    cmd_发卡原因 = 0
    cmd_领卡人 = 1
    cmd_领卡部门 = 2
End Enum
Private Enum mPictureIndex
    pic_余额 = 0
    pic_缴款合计 = 1
    pic_缴款信息 = 2
    pic_卡信息 = 3
    pic_补换卡 = 4
End Enum

'模块变量
Private mblnFirst As Boolean, mintSucces As Integer
Private mblnNotClick As Boolean, mblnChange As Boolean

Private mobjKeyboard As Object
Attribute mobjKeyboard.VB_VarHelpID = -1
Private Type Ty_Para
    bln缴款单打印 As Boolean
    bln连续充值 As Boolean
End Type
Private mTy_MoudlePara As Ty_Para

Private Type Ty_CardType
    str卡名称 As String
    str卡号前缀 As String
    lng卡号长度 As Long
    bln卡号密文 As Boolean
    int密码长度 As Integer
    int密码长度限制 As Integer
    byt密码规则 As Byte
    bln严格控制 As Boolean
    str限制类别 As String
    bln特定病人 As Boolean
    lng共用批次 As Long
    lng领用ID As Long
End Type
Private mCardType As Ty_CardType
Private mrs卡类型 As ADODB.Recordset
Private mobjCard As clsSquareCard '当前卡信息
Private mcllCard As Collection '卡片信息集合，回收时使用
Private mdbl实收合计 As Double '收款合计
Private mdbl本次误差 As Double

'支付相关
Private mobjPayCards As Cards
Private mlngPre支付方式 As Long
Private Type TY_PayMoney
    str结算方式  As String
    lng卡类别ID As Long
    byt结算性质 As Byte
    str刷卡卡号 As String
    str刷卡密码 As String
    str交易流水号 As String
    str交易说明 As String
    lng原结算序号 As Long
End Type
Private mCurCardPay As TY_PayMoney '本次卡支付
Private mrsBalance As ADODB.Recordset '原结算情况
Private mBytMoney As Byte '分币处理规则

Public Function zlShowCard(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As gCardEditType, ByVal lng卡类别 As Long, Optional ByVal lng卡ID As Long, _
    Optional lng充值ID As Long) As Boolean
    '程序入口,查看已发卡或增加发卡或修改发卡信息
    '入参：
    '   frmMain - 父窗口
    '   lngModule - 模块号
    '   strPrivs - 权限串
    '   EditType - 操作类型
    '   lng卡类别 As Long - 消费卡类别
    '   lng卡ID - 当前操作消费卡ID
    '   lng充值ID - 充值回退时传入，充值记录的ID
    '返回：操作成功返回True,否则返回False
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs:
    mlng卡类别 = lng卡类别: mEditType = EditType
    mlng卡ID = lng卡ID
    mlng充值ID = lng充值ID
    
    mintSucces = 0
    On Error Resume Next
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
End Function

Private Function CardIsValid(ByVal lng卡ID As Long) As Boolean
    '检查卡信息
    Dim strInfo As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str卡号 As String
    Dim msgBoxStyle As VbMsgBoxStyle
    
    On Error GoTo ErrHandler
    If lng卡ID = 0 Then CardIsValid = True: Exit Function
    If mEditType = gEd_发卡 Then CardIsValid = True: Exit Function
    
    strSQL = _
        "Select ID, 卡类型, 可否充值, 卡号, 序号, " & vbNewLine & _
        "       (Select Max(序号) From 消费卡信息 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号) As 最大序号," & vbNewLine & _
        "       To_Char(回收时间, 'yyyy-mm-dd hh24:mi:ss') As 回收时间, To_Char(有效期, 'yyyy-mm-dd hh24:mi:ss') As 有效期, " & vbNewLine & _
        "       To_Char(停用日期, 'yyyy-mm-dd hh24:mi:ss') As 停用日期, 当前状态, 发卡人, 余额" & vbNewLine & _
        "From 消费卡信息 A" & vbNewLine & _
        "Where a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID)
    
    If rsTemp.EOF Then
        strInfo = Switch(mEditType = gEd_修改, "修改卡信息", mEditType = gEd_充值, "充值", _
                         mEditType = gEd_退卡, "退卡", mEditType = gEd_取消退卡, "取消退卡", _
                         mEditType = gEd_换卡, "换卡", mEditType = gEd_补卡, "补卡", _
                         mEditType = gEd_充值回退, "充值回退", _
                         mEditType = gEd_回收, "回收", mEditType = gEd_取消回收, "取消回收", _
                         True, "继续操作")
        ShowMsgbox mCardType.str卡名称 & "可能已经被他人删除，不能" & strInfo & "！"
        Exit Function
    End If
    str卡号 = NVL(rsTemp!卡号)
    
    '检查卡号是否合法
    Select Case mEditType
    Case gEd_修改
        If Val(NVL(rsTemp!序号)) < Val(NVL(rsTemp!最大序号)) Then
            ShowMsgbox "不能修改历史卡号信息(卡号为:" & str卡号 & ") ！"
            Exit Function
        End If
    Case gEd_充值, gEd_充值回退, gEd_退卡, gEd_取消退卡, gEd_回收, gEd_取消回收, gEd_换卡, gEd_补卡
        strInfo = Switch(mEditType = gEd_充值, "充值", mEditType = gEd_充值回退, "充值回退", _
                         mEditType = gEd_退卡, "退卡", mEditType = gEd_取消回收, "取消回收", _
                         mEditType = gEd_换卡, "换卡", mEditType = gEd_补卡, "补卡", _
                         mEditType = gEd_回收, "回收", _
                         mEditType = gEd_取消退卡, "取消退卡", True, "继续操作")
        
        If Val(NVL(rsTemp!序号)) < Val(NVL(rsTemp!最大序号)) Then
            ShowMsgbox "不能对历史卡号进行" & strInfo & "(卡号为:" & str卡号 & ")！"
            Exit Function
        End If
        If mEditType = gEd_取消回收 Then
            '换卡回收的卡不能取消回收
            If Val(NVL(rsTemp!当前状态)) = 4 Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "为换卡回收的卡，不能取消回收！"
                Exit Function
            End If
            If NVL(rsTemp!回收时间, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "可能被他人取消回收(在用)，不能再取消回收！"
                Exit Function
            End If
        ElseIf mEditType = gEd_取消退卡 Then
            If NVL(rsTemp!回收时间, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "可能被他人取消退卡(在用)，不能再取消退卡！"
                Exit Function
            End If
        ElseIf mEditType = gEd_退卡 Then
            If NVL(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已被退卡，不能再退卡！"
                Exit Function
            End If
        Else
            If NVL(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已被回收或退卡，不能再" & strInfo & "！"
                Exit Function
            End If
        End If
        If Not (mEditType = gEd_回收 Or mEditType = gEd_取消回收) Then
            '停用的也可以回收和取消回收
            If NVL(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已经停止使用，不能再" & strInfo & "！"
                Exit Function
            End If
        End If
        
        Select Case mEditType
        Case gEd_回收
            If Val(NVL(rsTemp!余额)) > 0 Then
                If NVL(rsTemp!有效期, "3000-01-01 00:00:00") > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
                    msgBoxStyle = vbQuestion + vbYesNo + vbDefaultButton2
                Else '已失效，默认回收
                    msgBoxStyle = vbQuestion + vbYesNo + vbDefaultButton1
                End If
                If MsgBox("卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "当前还有余额(" & _
                    FormatEx(Val(NVL(rsTemp!余额)), 6, , , 2) & ")，你确定要回收吗？", msgBoxStyle) = vbNo Then Exit Function
            End If
        Case gEd_充值
            If Val(NVL(rsTemp!可否充值)) <> 1 Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "不是充值卡，不能充值！"
                Exit Function
            End If
        Case gEd_充值回退
            strSQL = "Select 1" & vbNewLine & _
                "From 病人卡结算记录 A, 消费卡信息 B" & vbNewLine & _
                "Where a.消费卡id = b.Id And a.Id = [1] And (Nvl(b.余额, 0) - Nvl(a.应收金额, 0)) >= 0"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng充值ID)
            If rsTemp.EOF Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "的余额不足，不能充值回退！"
                Exit Function
            End If
            
            '只有升级以后的才进行检查 And b.交易序号 > 0
            strSQL = "Select 1" & vbNewLine & _
                "From 病人卡结算记录 A, 帐户缴款余额 B" & vbNewLine & _
                "Where a.交易序号 = b.交易序号 And a.消费卡id = b.消费卡id And (a.应收金额 = b.余额 Or b.交易序号 <= 0) And a.Id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng充值ID)
            If rsTemp.EOF Then
                ShowMsgbox "该笔充值也被消费使用，不能充值回退！"
                Exit Function
            End If
        Case gEd_退卡
            If NVL(rsTemp!发卡人) <> UserInfo.姓名 Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "不是你发放的卡，不能退卡！"
                Exit Function
            End If
            
            strSQL = "Select 1 From 病人卡结算记录 Where 消费卡id = [1] And 记录性质 = 4 And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID)
            If rsTemp.EOF = False Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已经发生消费，不能再退卡处理，只能回收卡片！"
                Exit Function
            End If
            
            strSQL = "Select 1 From 病人卡结算记录　where 消费卡id = [1] And 记录性质 = 2 And 记录状态 = 1 Having Count(1) > 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID)
            If rsTemp.EOF = False Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已经被多次充值，不能再退卡，只能回收卡片！"
                Exit Function
            End If
            
            strSQL = "Select 1 From 病人卡结算记录　where 消费卡id = [1] And 记录性质 = 3 And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID)
            If rsTemp.EOF = False Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已进行了余额退款，不能再退卡，只能回收卡片！"
                Exit Function
            End If
        Case gEd_取消退卡
            If NVL(rsTemp!发卡人) <> UserInfo.姓名 Then
                ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "不是你发放的卡，不能取消退卡！"
                Exit Function
            End If
        End Select
    End Select
    CardIsValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitFace() As Boolean
    '初始化界面
    
    On Error GoTo ErrHandler
    Call SetCtlVisible
    Call FormResize
    
    pic(pic_卡信息).AutoRedraw = True: zlControl.PicShowFlat pic(pic_卡信息)
    pic(pic_缴款合计).AutoRedraw = True: zlControl.PicShowFlat pic(pic_缴款合计)
    pic(pic_缴款信息).AutoRedraw = True: zlControl.PicShowFlat pic(pic_缴款信息)
    pic(pic_补换卡).AutoRedraw = True: zlControl.PicShowFlat pic(pic_补换卡)
    cbo.SetListWidth cbo支付方式, cbo支付方式.Width * 2
    
    Call SetCtlEnable
    Call SetEnabledBackColor(Me)
    lvw限制类别.BackColor = IIf(lvw限制类别.Enabled, &H80000005, Me.BackColor)
    txt(txt_找补).BackColor = Me.BackColor

    Me.Caption = Switch(mEditType = gEd_发卡, "发卡", mEditType = gEd_修改, "信息修改", _
                        mEditType = gEd_换卡, "换卡", mEditType = gEd_补卡, "补卡", _
                        mEditType = gEd_查询, "信息查询", _
                        mEditType = gEd_充值, "充值管理", mEditType = gEd_充值回退, "充值回退", _
                        mEditType = gEd_回收, "回收管理", mEditType = gEd_取消回收, "取消回收", _
                        mEditType = gEd_退卡, "退卡", mEditType = gEd_取消退卡, "取消退卡", _
                        True, "发卡") & " - " & mCardType.str卡名称
    
    If mEditType = gEd_查询 Then
        cmdOK.Visible = False
        cmdCancel.Caption = cmdOK.Caption
    End If
    
    InitFace = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetCtlVisible()
    '设置控件的可见状态
    Dim blnVisible As Boolean
    
    On Error GoTo ErrHandler
    '卡类别启用了"特定病人"，则不能批量发卡
    If mEditType <> gEd_发卡 Or mEditType = gEd_发卡 And mCardType.bln特定病人 Then
        lbl(lbl_卡号至).Visible = False: txt(txt_结束卡号).Visible = False
        If mEditType <> gEd_回收 Then
            cmdAdd.Visible = False: cmdDelete.Visible = False
            lbl(lbl_卡张数).Visible = False: lbl(lbl_卡张数2).Visible = False
            vsfCardNo.Visible = False: lineSplit.Visible = False
        End If
        Call FrameResize(fra_卡号)
    End If
       
    blnVisible = (mEditType = gEd_发卡)
    lbl(lbl_密码).Visible = blnVisible: txt(txt_密码).Visible = blnVisible
    lbl(lbl_确认密码).Visible = blnVisible: txt(txt_确认密码).Visible = blnVisible
    
    blnVisible = (mEditType = gEd_发卡 Or mEditType = gEd_修改)
    dtp卡有效日期.Visible = blnVisible: txt(txt_卡有效期).Visible = Not blnVisible
    cmdSel(cmd_发卡原因).Visible = blnVisible
    IDKind.Visible = blnVisible And mCardType.bln特定病人
    cmdSel(cmd_领卡人).Visible = blnVisible And mCardType.bln特定病人 = False
    cmdSel(cmd_领卡部门).Visible = blnVisible And mCardType.bln特定病人 = False
    
    blnVisible = (mEditType <> gEd_发卡)
    lbl(lbl_发卡人).Visible = blnVisible: txt(txt_发卡人).Visible = blnVisible
    lbl(lbl_发卡日期).Visible = blnVisible: txt(txt_发卡日期).Visible = blnVisible
    lbl(lbl_回收人).Visible = blnVisible: txt(txt_回收人).Visible = blnVisible
    lbl(lbl_回收日期).Visible = blnVisible: txt(txt_回收日期).Visible = blnVisible
    Call FrameResize(fra_卡信息)
    
    blnVisible = (mEditType = gEd_发卡 Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡 Or mEditType = gEd_充值 Or mEditType = gEd_充值回退)
    fra(fra_充值信息).Visible = blnVisible
    
    blnVisible = (mEditType = gEd_发卡 Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡 _
                    Or mEditType = gEd_充值 Or mEditType = gEd_充值回退)
    fra(fra_缴款面板).Visible = blnVisible
    
    blnVisible = (mEditType = gEd_换卡 Or mEditType = gEd_补卡)
    pic(pic_补换卡).Visible = blnVisible
    If blnVisible Then
        blnVisible = (mEditType = gEd_换卡)
        lbl(lbl_原卡卡号).Visible = blnVisible
        txt(txt_原卡卡号).Visible = blnVisible
        lbl(lbl_原卡密码).Visible = blnVisible
        txt(txt_原卡密码).Visible = blnVisible
        
        lbl(lbl_补卡姓名).Visible = Not blnVisible
        txt(txt_补卡姓名).Visible = Not blnVisible
        IDKind补卡.Visible = Not blnVisible
        lbl(lbl_补卡原卡).Visible = Not blnVisible
        cbo原卡卡号.Visible = Not blnVisible
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCtlEnable()
    '设置控件的可用状态
    Dim blnEnable As Boolean
    
    On Error GoTo ErrHandler
    blnEnable = (mEditType = gEd_发卡 Or mEditType = gEd_修改)
    cbo卡类型.Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_发卡 Or mEditType = gEd_回收 Or mEditType = gEd_充值)
    txt(txt_开始卡号).Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_发卡 Or mEditType = gEd_修改)
    chk充值.Enabled = blnEnable And mobjCard.已充值 = False
    txt(txt_发卡原因).Enabled = blnEnable
    txt(txt_领卡人).Enabled = blnEnable
    txt(txt_领卡部门).Enabled = blnEnable And mCardType.bln特定病人 = False
    txt(txt_备注).Enabled = blnEnable
    lvw限制类别.Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_发卡)
    txt(txt_卡面额).Enabled = blnEnable And zlStr.IsHavePrivs(mstrPrivs, "允许更改卡面额")
    txt(txt_销售额).Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_发卡 Or mEditType = gEd_充值) And chk充值.value = vbChecked
    txt(txt_充值扣率).Enabled = blnEnable
    txt(txt_本次充值).Enabled = blnEnable
    txt(txt_充值缴款).Enabled = blnEnable
    txt(txt_充值说明).Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_发卡 Or mEditType = gEd_充值 Or mEditType = gEd_取消退卡)
    txt(txt_缴款人).Enabled = blnEnable
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FrameResize(ByVal Index As Integer)
    Dim sngTop As Single '下一行控件的Top值
    Dim sngSplit As Single '控件的行间距
    Dim sngDiff As Single '标签控件的Top与文本框控件Top的差距
    
    On Error Resume Next
    sngDiff = 60
    sngTop = IIf(mEditType = gEd_发卡, 100, 50)
    sngSplit = IIf(mEditType = gEd_发卡, 80, 160)
    Select Case Index
    Case fra_卡号
        If mEditType = gEd_发卡 And mCardType.bln特定病人 = False Then Exit Sub
        If mEditType = gEd_回收 Then
            cmdAdd.Left = txt(txt_开始卡号).Left + txt(txt_开始卡号).Width + 100
            cmdDelete.Left = cmdAdd.Left + cmdAdd.Width + 50
            lbl(lbl_卡张数).Left = cmdDelete.Left + cmdDelete.Width + 150
            Call SetLblCaption(lbl_卡张数, True)
            sngTop = sngTop + txt(txt_开始卡号).Height + 100
            
            vsfCardNo.Top = sngTop: vsfCardNo.Height = 1000
            sngTop = sngTop + vsfCardNo.Height + 100
            
            fra(fra_卡号).Height = sngTop
            lineSplit.Y1 = fra(fra_卡号).Height - 10: lineSplit.Y2 = lineSplit.Y1
        Else
            cbo卡类型.Top = sngTop: cbo卡类型.Left = txt(txt_密码).Left
            cbo卡类型.Width = txt(txt_密码).Width
            lbl(lbl_卡类型).Top = cbo卡类型.Top + sngDiff: lbl(lbl_卡类型).Left = cbo卡类型.Left - lbl(lbl_卡类型).Width - 20
            txt(txt_开始卡号).Top = cbo卡类型.Top: txt(txt_开始卡号).Left = txt(txt_确认密码).Left
            txt(txt_开始卡号).Width = txt(txt_确认密码).Width
            lbl(lbl_卡号).Top = lbl(lbl_卡类型).Top: lbl(lbl_卡号).Left = txt(txt_开始卡号).Left - lbl(lbl_卡号).Width - 60
            sngTop = sngTop + txt(txt_开始卡号).Height
            
            fra(fra_卡号).Height = sngTop
            fra(fra_卡号).Width = fra(fra_卡信息).Width
        End If
    Case fra_卡信息
        If mEditType = gEd_发卡 Then
            txt(txt_密码).Top = sngTop: lbl(lbl_密码).Top = txt(txt_密码).Top + sngDiff
            txt(txt_确认密码).Top = txt(txt_密码).Top: lbl(lbl_确认密码).Top = lbl(lbl_密码).Top
            sngTop = sngTop + txt(txt_密码).Height + sngSplit
        End If
        
        txt(txt_卡有效期).Top = sngTop: lbl(lbl_卡有效期).Top = sngTop + sngDiff
        dtp卡有效日期.Top = txt(txt_卡有效期).Top
        chk充值.Top = lbl(lbl_卡有效期).Top
        sngTop = sngTop + txt(txt_卡有效期).Height + sngSplit
        
        txt(txt_发卡原因).Top = sngTop: lbl(lbl_发卡原因).Top = sngTop + sngDiff
        cmdSel(cmd_发卡原因).Top = txt(txt_发卡原因).Top
        sngTop = sngTop + txt(txt_发卡原因).Height + sngSplit
        
        txt(txt_领卡人).Top = sngTop: lbl(lbl_领卡人).Top = sngTop + sngDiff
        IDKind.Top = txt(txt_领卡人).Top: cmdSel(cmd_领卡人).Top = txt(txt_领卡人).Top
        txt(txt_领卡部门).Top = sngTop: lbl(lbl_领卡部门).Top = sngTop + sngDiff
        cmdSel(cmd_领卡部门).Top = txt(txt_领卡部门).Top
        sngTop = sngTop + txt(txt_领卡人).Height + sngSplit
        If Not ((mEditType = gEd_发卡 Or mEditType = gEd_修改) And mCardType.bln特定病人) Then
            txt(txt_领卡人).Left = txt(txt_发卡人).Left: txt(txt_领卡人).Width = txt(txt_发卡人).Width
        End If
        
        txt(txt_备注).Top = sngTop: lbl(lbl_备注).Top = sngTop + sngDiff
        sngTop = sngTop + txt(txt_领卡人).Height + sngSplit
        
        If mEditType <> gEd_发卡 Then
            txt(txt_发卡人).Top = sngTop: lbl(lbl_发卡人).Top = txt(txt_发卡人).Top + sngDiff
            txt(txt_发卡日期).Top = txt(txt_发卡人).Top: lbl(lbl_发卡日期).Top = lbl(lbl_发卡人).Top
            sngTop = sngTop + txt(txt_发卡人).Height + sngSplit
            
            txt(txt_回收人).Top = sngTop: lbl(lbl_回收人).Top = txt(txt_回收人).Top + sngDiff
            txt(txt_回收日期).Top = txt(txt_回收人).Top: lbl(lbl_回收日期).Top = lbl(lbl_回收人).Top
            sngTop = sngTop + txt(txt_回收人).Height + sngSplit
        End If
        
        fra(Index).Height = sngTop
    Case fra_类别余额
        fra(fra_余额).Top = fra(Index).Height - fra(fra_余额).Height + 10
        pic(pic_余额).Top = fra(fra_余额).Top - pic(pic_余额).Height / 2 + 100
        fra(fra_限制类别).Top = 50
        fra(fra_限制类别).Height = pic(pic_余额).Top - fra(fra_限制类别).Top
        lvw限制类别.Height = fra(fra_限制类别).Height - lvw限制类别.Top - sngSplit
    End Select
End Sub

Private Sub FormResize()
    Dim sngTop As Single '下一行控件的Top值
    Dim sngSplit As Single '控件的行间距
    
    On Error Resume Next
    sngTop = 10: sngSplit = 80
    fra(fra_卡号).Top = sngTop
    sngTop = sngTop + fra(fra_卡号).Height + sngSplit
    
    fra(fra_卡信息).Top = sngTop
    sngTop = sngTop + fra(fra_卡信息).Height
    
    fra(fra_卡面值).Top = sngTop
    sngTop = sngTop + fra(fra_卡面值).Height + sngSplit
    
    If mEditType = gEd_发卡 Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡 Or mEditType = gEd_充值 Or mEditType = gEd_充值回退 Then
        fra(fra_充值信息).Top = sngTop
        sngTop = sngTop + fra(fra_充值信息).Height + sngSplit
    End If
    
    If mEditType = gEd_发卡 And mCardType.bln特定病人 = False Or mEditType = gEd_回收 Then
        fra(fra_类别余额).Top = fra(fra_卡信息).Top
    Else
        fra(fra_类别余额).Top = fra(fra_卡号).Top
    End If
    fra(fra_类别余额).Height = sngTop - fra(fra_类别余额).Top - sngSplit
    Call FrameResize(fra_类别余额)
    
    '卡基本信息
    pic(pic_卡信息).Top = sngSplit
    pic(pic_卡信息).Height = sngTop
    sngTop = pic(pic_卡信息).Top + pic(pic_卡信息).Height + sngSplit
    
    '换卡补卡面板
    If mEditType = gEd_换卡 Or mEditType = gEd_补卡 Then
        pic(pic_补换卡).Top = sngTop
        sngTop = sngTop + pic(pic_补换卡).Height + sngSplit
    End If
    
    '缴款面板
    If mEditType = gEd_发卡 Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡 Or mEditType = gEd_充值 Or mEditType = gEd_充值回退 Then
        fra(fra_缴款面板).Top = sngTop
        sngTop = sngTop + fra(fra_缴款面板).Height + sngSplit
    End If
    
    sngTop = sngTop - sngSplit
    fra(fra_按钮面板).Top = sngTop
    sngTop = sngTop + fra(fra_按钮面板).Height
    
    Me.Height = sngTop + 480
End Sub

Private Sub cbo原卡卡号_Click()
    On Error GoTo ErrHandler
    If Val(cbo原卡卡号.Tag) = cbo原卡卡号.ItemData(cbo原卡卡号.ListIndex) Then Exit Sub
    cbo原卡卡号.Tag = cbo原卡卡号.ItemData(cbo原卡卡号.ListIndex)
    
    LoadCardData 1, cbo原卡卡号.Text
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo原卡卡号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo卡类型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo支付方式_Click()
    Dim objCard As Card
    Dim ty_Temp As TY_PayMoney
    Dim intSelectIndex As Integer
    
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex) Then Exit Sub
    
    If (mEditType = gEd_充值回退 Or mEditType = gEd_退卡) And mlngPre支付方式 > 0 Then
        '如果不在原缴款结算方式中就不用检查，主要针对支持“转帐及代扣”的
        Set objCard = mobjPayCards(mlngPre支付方式)
        mrsBalance.Filter = "结算方式='" & objCard.结算方式 & "'"
        
        If Not mrsBalance.EOF Then
            mblnNotClick = True
            intSelectIndex = cbo支付方式.ListIndex
            cbo支付方式.ListIndex = cbo.FindIndex(cbo支付方式, mlngPre支付方式)
            If CheckThreeBalanceToCash(objCard) = False Then mblnNotClick = False: Exit Sub
            cbo支付方式.ListIndex = intSelectIndex
            mblnNotClick = False
        End If
    End If
    
    mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    
    mCurCardPay = ty_Temp '自定义Type初始化
    With mCurCardPay
        .str结算方式 = objCard.结算方式
        .lng卡类别ID = IIf(objCard.接口序号 > 0, objCard.接口序号, 0)
        .byt结算性质 = objCard.结算性质
    End With
    
    txt(txt_缴款).Text = ""
    Call SetControlProperty
    Exit Sub
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlProperty(Optional ByVal blnLoadDefault As Boolean)
    '设置控件属性
    '入参:
    '   blnLoadDefault-是否加载缺省值
    Dim blnDel As Boolean, objCard As Card
    Dim blnEnabled As Boolean
    Dim dblTemp As Double, dblErrMoney As Double
    
    On Error GoTo ErrHandler
    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    
    mdbl本次误差 = 0
    If mEditType = gEd_发卡 Or mEditType = gEd_充值 Then
        If objCard.结算性质 = 1 Then
            txt(txt_销售额).Text = Format(CentMoney(Val(txt(txt_销售额).Text), mBytMoney), "#0.00;-#0.00")
            txt(txt_充值缴款).Text = Format(CentMoney(Val(txt(txt_充值缴款).Text), mBytMoney), "#0.00;-#0.00")
            If Val(txt(txt_充值缴款).Text) > Val(txt(txt_本次充值).Text) Then
                txt(txt_本次充值).Text = Format(Val(txt(txt_充值缴款).Text), "#0.00;-#0.00")
            End If
            If Val(txt(txt_本次充值).Text) <> 0 Then
                txt(txt_充值扣率).Text = Format((Round(Val(txt(txt_充值缴款).Text) / Val(txt(txt_本次充值).Text), 6)) * 100, "0.00")
            End If
            Call Calc实收合计(True)
        End If
    End If
    
    blnDel = (mdbl实收合计 < 0)
    If blnDel Then
        lbl缴款.Caption = "退款合计"
        lbl(lbl_支付方式).Caption = "退 款"
        lbl(lbl_缴款合计).ForeColor = vbRed
        lbl(lbl_支付方式).ForeColor = vbRed
    Else
        lbl缴款.Caption = "收款合计"
        lbl(lbl_支付方式).Caption = "收 款"
        lbl(lbl_缴款合计).ForeColor = vbBlue
        lbl(lbl_支付方式).ForeColor = vbBlack
    End If
    
    '支票、一卡通和老版一卡通允许输入缴款单位
    '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
    blnEnabled = InStr(",2,7,8,", "," & objCard.结算性质 & ",") > 0
    txt(txt_开户行).Enabled = objCard.结算性质 <> 1
    txt(txt_帐号).Enabled = objCard.结算性质 <> 1
    txt(txt_结算号码).Enabled = objCard.结算性质 <> 1
    If objCard.结算性质 = 1 Then
        txt(txt_开户行).Text = ""
        txt(txt_帐号).Text = ""
        txt(txt_结算号码).Text = ""
        mdbl本次误差 = mdbl实收合计 - CentMoney(mdbl实收合计, mBytMoney)
    End If
    Call zl_SetCtlBackColor(Array(txt(txt_开户行), txt(txt_帐号), txt(txt_结算号码)), Me)
    
    lbl(lbl_缴款合计).Caption = Format(Abs(mdbl实收合计 - mdbl本次误差), "0.00")
                
    '缺省金额的设置
    txt(txt_缴款).Locked = False
    If objCard.接口序号 > 0 Then '三方结算
        txt(txt_缴款).Text = Format(Abs(mdbl实收合计), "0.00")
        txt(txt_缴款).Locked = True
    ElseIf objCard.结算性质 = 1 Then '现金处理
        txt(txt_缴款).Text = IIf(blnDel, Format(Abs(mdbl实收合计 - mdbl本次误差), "0.00"), "")
    Else
        txt(txt_缴款).Text = Format(Abs(mdbl实收合计), "0.00")
        txt(txt_缴款).Locked = True
    End If
    lbl(lbl_误差).Caption = FormatEx(IIf(blnDel, -1, 1) * mdbl本次误差, 6, , , 2)
    lbl(lbl_误差).Visible = Val(lbl(lbl_误差).Caption) <> 0
    lbl(lbl_误差).Caption = "本次误差：" & lbl(lbl_误差).Caption
    
    '计算找补
    Call SetLblCaption(lbl_找补)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk充值_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdAdd_Click()
    Dim strCardNoRange As String
    Dim lng卡张数 As Long, strCardNos As String, lng卡ID As Long
    Dim objListItem As ListItem
    
    On Error GoTo ErrHandler
    If CheckInput卡号(False, lng卡张数, strCardNos, lng卡ID) = False Then Exit Sub
    
    strCardNoRange = Trim(txt(txt_开始卡号).Text)
    If Trim(txt(txt_结束卡号).Text) <> "" Then
        strCardNoRange = strCardNoRange & "～" & Trim(txt(txt_结束卡号).Text)
    End If
    If mEditType = gEd_回收 Then strCardNos = strCardNoRange
    If ZL_vsGrid_AddCell(vsfCardNo, strCardNoRange, Array(lng卡张数, strCardNos, lng卡ID)) = False Then Exit Sub
    If mEditType = gEd_回收 Then
        mcllCard.Add mobjCard, "K" & mlng卡ID
        Call FindDataInGrid(strCardNos, True)
    End If
    With vsfCardNo
        .Redraw = flexRDNone
        Call ZL_vsGrid_AutoSetGridRowAndCol(vsfCardNo)
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .Redraw = flexRDBuffered
    End With
    
    txt(txt_开始卡号).Text = "": txt(txt_结束卡号).Text = ""
    
    '显示计算当前卡张数
    Call SetLblCaption(lbl_卡张数, mEditType = gEd_回收)
    Call Calc实收合计
    
    zlControl.ControlSetFocus txt(txt_开始卡号)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Dim blnYes As Boolean
    
    On Error GoTo ErrHandler
    If mblnChange Then
        If mEditType = gEd_发卡 Or mEditType = gEd_充值 And mTy_MoudlePara.bln连续充值 Then
            ShowMsgbox "确实要清除当前已录入的内容吗？", True, blnYes
            If blnYes = False Then Exit Sub
            Call ClearCtlData
            mlng卡ID = 0
            Call zlControl.ControlSetFocus(txt(txt_开始卡号))
            Exit Sub
        End If
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    Dim lng卡张数 As Long
    Dim strCardNoRange As String, strCardNos As String
    Dim blnYes As Boolean
    Dim varData As Variant
    
    On Error GoTo ErrHandler
    If ZL_vsGrid_CurrCellHaveData(vsfCardNo) = False Then Exit Sub
    
    strCardNoRange = vsfCardNo.TextMatrix(vsfCardNo.Row, vsfCardNo.Col)
    Call ShowMsgbox("你确定要从" & IIf(mEditType = gEd_回收, "回收", "发卡") & "列表中移除 " & strCardNoRange & " 吗？", True, blnYes)
    If blnYes = False Then zlControl.ControlSetFocus cmdDelete: Exit Sub
    
    varData = vsfCardNo.Cell(flexcpData, vsfCardNo.Row, vsfCardNo.Col) 'Array(卡张数,分解卡号,消费卡ID)
    lng卡张数 = varData(0)
    
    If ZL_vsGrid_RemoveCell(vsfCardNo) = False Then Exit Sub
    If mEditType = gEd_回收 And Not mcllCard Is Nothing Then
        If CollExitsValue(mcllCard, "K" & mlng卡ID) Then
            mcllCard.Remove "K" & mlng卡ID
        End If
    End If
    With vsfCardNo
        .Redraw = flexRDNone
        Call ZL_vsGrid_AutoSetGridRowAndCol(vsfCardNo)
        .Redraw = flexRDBuffered
    End With
    
    '显示计算当前卡张数
    Call SetLblCaption(lbl_卡张数, mEditType = gEd_回收)
    Call Calc实收合计
    
    zlControl.ControlSetFocus txt(txt_开始卡号)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mblnChange = False: mintSucces = mintSucces + 1
    
    If mEditType = gEd_发卡 Or mEditType = gEd_充值 And mTy_MoudlePara.bln连续充值 Then
        Call ClearCtlData
        mlng卡ID = 0
        Call zlControl.ControlSetFocus(txt(txt_开始卡号))
        Exit Sub
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Dim lngID As Long, str编码 As String, str名称 As String
    
    On Error GoTo ErrHandler
    Select Case Index
    Case cmd_领卡人
        '选择人员
        lngID = Val(txt(txt_领卡部门).Tag)
        If Select人员选择器(Me, txt(txt_领卡人), "", lngID, , True) = False Then Exit Sub
        
        If mEditType = gEd_发卡 Or mEditType = gEd_修改 Then
            '领卡人就是缴款人
            txt(txt_缴款人).Text = txt(txt_领卡人).Text
            txt(txt_缴款人).Tag = txt(txt_领卡人).Tag
        End If
        '需要读取缺省部门:
        If zl_From人员获取缺省部门(Val(txt(txt_领卡人).Tag), str编码, str名称, lngID) Then
            txt(txt_领卡部门).Text = str编码 & "-" & str名称
            txt(txt_领卡部门).Tag = lngID
        End If
    Case cmd_领卡部门
        '选择缺省部门
        lngID = Val(txt(txt_领卡人).Tag)
        If Select部门选择器(Me, txt(txt_领卡部门), "", "", IIf(lngID = 0, False, True), "", 0, _
            "部门选择器", , , , , lngID) = False Then Exit Sub
    Case cmd_发卡原因
        If zl_SelectAndNotAddItem(Me, txt(txt_发卡原因), "", "常用发卡原因", "常用发卡原因选择", True, True) = False Then Exit Sub
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtp卡有效日期_Change()
    mblnChange = True
End Sub

Private Sub dtp卡有效日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub ClearCtlData()
    '清除控件数据
    Dim ctl As Control
    
    On Error GoTo ErrHandler
    mdbl实收合计 = 0
    mdbl本次误差 = 0
    Set mcllCard = New Collection
    Set mobjCard = New clsSquareCard
    mobjCard.有效期 = "3000-01-01"
    vsfCardNo.BackColorSel = &HC0C0C0
    
    For Each ctl In Me.Controls
        If UCase(TypeName(ctl)) = "TEXTBOX" Then
            ctl.Text = "": ctl.Tag = ""
        End If
    Next
    Call SetDefaultValue
    
    vsfCardNo.Clear 1
    vsfCardNo.Rows = 0
    vsfCardNo.Cols = 0
    cmdDelete.Enabled = False
    
    '显示计算当前卡张数
    Call SetLblCaption(lbl_卡张数, mEditType = gEd_回收, True)
    If mEditType = gEd_发卡 Then Call Calc实收合计
    
    chk充值.value = vbUnchecked
    
    lbl(lbl_卡余额).Caption = "0.00": lbl(lbl_卡余额).Tag = ""
    lbl缴款.Caption = "收款合计"
    lbl(lbl_缴款合计).Caption = "0.00"
    Call SetLblCaption(lbl_找补, mEditType = gEd_回收)
    
    dtp卡有效日期.value = "3000-01-01"
    dtp卡有效日期.value = Null
    txt(txt_充值扣率).Text = "0.00"
    
    mblnChange = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetLblCaption(ByVal Index As Integer, Optional ByVal bln回收 As Boolean, _
    Optional blnClear As Boolean)
    '设置共发消费卡张数页签的文本显示
    '入参：
    '   blnClear 当前是否为清除界面数据时调用
    Dim lngSpace As Long
    Dim dbl找补 As Double
    Dim blnDel As Boolean, lngCount As Long
    
    On Error GoTo ErrHandler
    Select Case Index
    Case lbl_卡张数
        lngCount = GetCardsCount()
        lbl(lbl_卡张数2).Caption = CStr(lngCount)
        
        lngSpace = Len(lbl(lbl_卡张数2).Caption) + 1
        If mEditType = gEd_回收 Then
            lbl(lbl_卡张数).Caption = "共回收" & Space(lngSpace) & "张"
            lbl(lbl_卡张数2).Left = lbl(lbl_卡张数).Left + 680
            '没有卡时清除界面卡信息
            If Val(lbl(lbl_卡张数2).Caption) = 0 And blnClear = False Then
                Call ClearCtlData
            End If
        Else
            lbl(lbl_卡张数).Caption = "共发" & Space(lngSpace) & "张"
            lbl(lbl_卡张数2).Left = lbl(lbl_卡张数).Left + 470
        End If
    Case lbl_找补
        '设置找补的标题
        blnDel = mdbl实收合计 - mdbl本次误差 < 0
        dbl找补 = IIf(blnDel, -1, 1) * Val(txt(txt_缴款).Text) - (mdbl实收合计 - mdbl本次误差)
        txt(txt_找补).Tag = dbl找补
        txt(txt_找补).Text = Format(dbl找补, "0.00")
        lbl(lbl_找补).ForeColor = IIf(dbl找补 >= 0, vbBlack, vbRed)
        txt(txt_找补).ForeColor = IIf(dbl找补 >= 0, vbBlack, vbRed)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    '初始化界面数据
    Dim strValue As String
    Dim rs收费类别 As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Call ClearCtlData
    
    '消费卡分币处理方式
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 4, 1)))
    
    If mEditType = gEd_发卡 Or mEditType = gEd_修改 Then
        mrs卡类型.Filter = ""
        Do While Not mrs卡类型.EOF
            cbo卡类型.AddItem NVL(mrs卡类型!编码) & "-" & NVL(mrs卡类型!名称)
            mrs卡类型.MoveNext
        Loop
        If cbo卡类型.ListCount > 0 Then cbo卡类型.ListIndex = 0
    End If
    
    lvw限制类别.ListItems.Clear
    If mEditType = gEd_发卡 Or mEditType = gEd_修改 Then
        Set rs收费类别 = zlGet收费类别
        rs收费类别.Filter = 0
        Do While Not rs收费类别.EOF
            lvw限制类别.ListItems.Add , NVL(rs收费类别!名称), NVL(rs收费类别!编码) & "-" & NVL(rs收费类别!名称)
            rs收费类别.MoveNext
        Loop
        
        Call Load限制类别(mCardType.str限制类别)
    End If
    
    If mEditType = gEd_发卡 Or mEditType = gEd_取消退卡 Or mEditType = gEd_充值 Then
        If Load支付方式() = False Then Exit Function
    ElseIf mEditType = gEd_退卡 Or mEditType = gEd_充值回退 Then
        If Load支付方式(True) = False Then Exit Function
    End If
    
    If (mEditType = gEd_发卡 Or mEditType = gEd_修改) And mCardType.bln特定病人 Then
        Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser)
    End If
    
    If mEditType = gEd_补卡 Then
        Call IDKind补卡.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser)
    End If
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadCardData(ByVal bytMode As Byte, _
    Optional ByVal strNO As String, Optional ByVal lng卡ID As Long) As Boolean
    '加载数据到控件
    '入参：
    '   bytMode 0-按消费卡ID加载，1-按消费卡卡号加载
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String, lng卡序号 As Long
    
    On Error GoTo ErrHandler
    If bytMode = 1 Then
        strWhere = " And a.卡号 = [2] And a.接口编号 = [3]" & vbNewLine & _
                   " And 序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = a.卡号 And 接口编号 = a.接口编号)"
    Else
        strWhere = " And a.Id = [1]"
    End If
    strSQL = _
        "Select a.Id, a.卡类型, a.卡号, a.序号, a.可否充值, a.有效期, a.发卡原因, a.密码," & vbNewLine & _
        "       a.发卡人, a.领卡人, a.病人id, a.发卡时间, a.回收人, a.回收时间," & vbNewLine & _
        "       Mod(a.当前状态, 10) As 当前状态, a.备注, a.卡面金额, a.销售金额, a.充值折扣率," & vbNewLine & _
        "       a.余额, a.停用人, a.停用日期, a.领卡部门id, a.限制类别," & vbNewLine & _
        "       Decode(b.编码, Null, '', b.编码 || '-' || b.名称) As 领卡部门," & vbNewLine & _
        "       Nvl((Select 1 From 病人卡结算记录 Where 消费卡ID = a.ID And 记录性质 = 2 And Rownum < 2), 0) As 已充值" & vbNewLine & _
        "From 消费卡信息 A, 部门表 B" & vbNewLine & _
        "Where a.领卡部门id = b.Id(+)" & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID, strNO, mlng卡类别)

    If rsTemp.EOF Then
        ShowMsgbox "未找到相关的" & mCardType.str卡名称 & "信息，可能已经被他人删除！"
        Exit Function
    End If

    If bytMode = 1 Then
        '根据卡号读取出卡信息后检查该卡号
        If CardIsValid(Val(NVL(rsTemp!id))) = False Then Exit Function
    End If
    
    mlng卡ID = Val(NVL(rsTemp!id))
    lng卡序号 = Val(NVL(rsTemp!序号))
    Set mobjCard = New clsSquareCard
    With mobjCard
        .卡类型 = NVL(rsTemp!卡类型)
        .卡号 = NVL(rsTemp!卡号)
        
        .充值卡 = Val(NVL(rsTemp!可否充值)) = 1
        .有效期 = Format(NVL(rsTemp!有效期), "yyyy-MM-dd")
        .发卡原因 = NVL(rsTemp!发卡原因)
        .发卡人 = NVL(rsTemp!发卡人)
        .领卡人 = NVL(rsTemp!领卡人)
        .病人ID = Val(NVL(rsTemp!病人ID))
        .发卡时间 = Format(NVL(rsTemp!发卡时间), "yyyy-MM-dd HH:mm:ss")
        .领卡部门id = Val(NVL(rsTemp!领卡部门id))
        .领卡部门 = NVL(rsTemp!领卡部门)
        .备注 = NVL(rsTemp!备注)
        
        .卡面值 = Val(NVL(rsTemp!卡面金额))
        .实际销售 = Val(NVL(rsTemp!销售金额))
        .充值折扣率 = Val(NVL(rsTemp!充值折扣率))
        .回收人 = NVL(rsTemp!回收人)
        .回收时间 = Format(NVL(rsTemp!回收时间), "yyyy-MM-dd HH:mm:ss")
        .停用人 = NVL(rsTemp!停用人)
        .停用日期 = Format(NVL(rsTemp!停用日期), "yyyy-MM-dd HH:mm:ss")
        .限制类别 = NVL(rsTemp!限制类别)
        
        .当前状态 = Val(NVL(rsTemp!当前状态))
        .卡余额 = Val(NVL(rsTemp!余额))
        .已充值 = Val(NVL(rsTemp!已充值)) = 1
        .原密码 = NVL(rsTemp!密码)
    End With
    
    Call ShowCardInfo(mobjCard)
    
    '读取充值信息
    Select Case mEditType
    Case gEd_充值回退
        strSQL = _
            "Select a.结算方式, a.实收金额, a.扣率, a.应收金额, a.备注, a.结算序号," & vbNewLine & _
            "       a.卡类别id, a.结算卡号, a.交易流水号, a.交易说明" & vbNewLine & _
            "From 病人卡结算记录 A" & vbNewLine & _
            "Where a.记录性质 = 2 And 记录状态 = 1 And a.Id = [1]"
        Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng充值ID)
        If mrsBalance.EOF Then
            ShowMsgbox "未找到原充值记录，不能继续！"
            Exit Function
        End If
        
        txt(txt_充值扣率).Text = Format(Val(NVL(mrsBalance!扣率)), "#0.00;-#0.00")
        txt(txt_本次充值).Text = Format(Val(NVL(mrsBalance!应收金额)), "#0.00;-#0.00")
        txt(txt_充值缴款).Text = Format(Val(NVL(mrsBalance!实收金额)), "#0.00;-#0.00")
        txt(txt_充值说明).Text = NVL(mrsBalance!备注)
        Call Load支付方式(True)
    Case gEd_退卡, gEd_取消退卡
        strSQL = _
            "Select a.记录性质, a.结算方式, a.实收金额, a.扣率, a.应收金额, a.备注, a.结算序号," & vbNewLine & _
            "       a.卡类别id, a.结算卡号, a.交易流水号, a.交易说明" & vbNewLine & _
            "From 病人卡结算记录 A, 病人卡结算记录 B" & vbNewLine & _
            "Where a.消费卡id = b.消费卡id And a.结算序号 = b.结算序号" & vbNewLine & _
            "      And a.记录状态 = [2] And a.记录性质 In (1, 2)" & vbNewLine & _
            "      And b.记录性质 = 1 And b.消费卡id = [1]" & vbNewLine & _
            "      And b.序号 = (Select Max(序号) From 病人卡结算记录 Where 消费卡id = b.消费卡id And 记录性质 = 1)"
        Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng卡ID, IIf(mEditType = gEd_退卡, 1, 3))
        mrsBalance.Filter = "记录性质=2"
        If Not mrsBalance.EOF Then
            txt(txt_充值扣率).Text = Format(Val(NVL(mrsBalance!扣率)), "#0.00;-#0.00")
            txt(txt_本次充值).Text = Format(Val(NVL(mrsBalance!应收金额)), "#0.00;-#0.00")
            txt(txt_充值缴款).Text = Format(Val(NVL(mrsBalance!实收金额)), "#0.00;-#0.00")
            txt(txt_充值说明).Text = NVL(mrsBalance!备注)
        End If
        If mEditType = gEd_退卡 Then
            Call Load支付方式(True)
        Else
            Call Calc余额
        End If
    End Select
    
    Call Calc实收合计

    LoadCardData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowCardInfo(objCard As clsSquareCard, Optional ByVal blnSelectNoList As Boolean)
    '显示当前卡信息
    '入参：
    '   blnSelectNoList 是否选择卡号列表显示卡信息
    On Error GoTo ErrHandler

    With objCard
        mblnNotClick = True
        cbo.SeekIndex cbo卡类型, .卡类型
        If cbo卡类型.ListIndex = -1 Then
            cbo卡类型.AddItem .卡类型
            cbo卡类型.ListIndex = cbo卡类型.NewIndex
        End If
        If blnSelectNoList = False Then
            '选择卡列表时不显示卡号到卡号文本框
            txt(txt_开始卡号).Text = .卡号
        End If
    
        chk充值.value = IIf(.充值卡, vbChecked, vbUnchecked)
        txt(txt_卡有效期).Text = Format(.有效期, "yyyy-MM-DD")
        If txt(txt_卡有效期).Text <> "" Then
            If CDate(txt(txt_卡有效期).Text) >= CDate("3000-01-01") Then txt(txt_卡有效期).Text = ""
        End If
        If txt(txt_卡有效期).Text <> "" Then
            dtp卡有效日期.value = CDate(txt(txt_卡有效期).Text)
        Else
            dtp卡有效日期.value = Null
        End If
        
        txt(txt_发卡原因).Text = .发卡原因
        txt(txt_领卡人).Text = .领卡人
        txt(txt_领卡人).Tag = .病人ID
        txt(txt_领卡部门).Text = .领卡部门
        txt(txt_领卡部门).Tag = .领卡部门id
        txt(txt_备注).Text = .备注
        
        txt(txt_发卡人).Text = .发卡人
        txt(txt_发卡日期).Text = .发卡时间
        txt(txt_回收人).Text = .回收人
        txt(txt_回收日期).Text = .回收时间
        If txt(txt_回收日期).Text <> "" Then
            If CDate(txt(txt_回收日期).Text) >= CDate("3000-01-01") Then txt(txt_回收日期).Text = ""
        End If
        
        
        txt(txt_卡面额).Text = Format(.卡面值, "#0.00;-#0.00")
        txt(txt_销售额).Text = Format(.实际销售, "#0.00;-#0.00")
        txt(txt_销售额).Tag = .实际销售
        
        txt(txt_充值扣率).Text = Format(.充值折扣率, "#0.00;-#0.00")
        txt(txt_本次充值).Text = ""
        txt(txt_充值缴款).Text = ""
        txt(txt_充值说明).Text = ""
    
        lbl(lbl_卡余额).Caption = Format(.卡余额, "#0.00;-#0.00")
        
        If mEditType = gEd_换卡 Then txt(txt_原卡卡号).Text = .卡号
        If mEditType = gEd_换卡 Or mEditType = gEd_补卡 Then
            '新卡密码缺省为原卡密码
            txt(txt_新卡密码).Text = .原密码
            txt(txt_新卡确认密码).Text = .原密码
        End If
        
        Call Load限制类别(.限制类别)
    End With
    mblnNotClick = False
    Exit Sub
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then 'Chr(22):Ctrl+V
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    If InitModulePara() = False Then Unload Me: Exit Sub
    If InitData() = False Then Unload Me: Exit Sub
    
    If mlng卡ID <> 0 Then
        If CardIsValid(mlng卡ID) = False Then Unload Me: Exit Sub
        If LoadCardData(0, , mlng卡ID) = False Then Unload Me: Exit Sub
        
        '显示计算当前卡张数
        If mEditType = gEd_回收 Then
            txt(txt_开始卡号).Tag = 1
            Call SetLblCaption(lbl_卡张数, True)
        End If
    End If
    If InitFace() = False Then Unload Me: Exit Sub
    
    Call CreateObjectKeyboard
    mblnChange = False
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    On Error Resume Next
    '焦点定位
    Select Case mEditType
    Case gEd_发卡, gEd_修改
        zlControl.ControlSetFocus cbo卡类型
    Case gEd_回收
        zlControl.ControlSetFocus txt(txt_开始卡号)
    Case gEd_充值
        If txt(txt_开始卡号).Text = "" Then
            zlControl.ControlSetFocus txt(txt_开始卡号)
        Else
            zlControl.ControlSetFocus txt(txt_本次充值)
        End If
    Case gEd_换卡
        If txt(txt_原卡卡号).Text = "" Then
            zlControl.ControlSetFocus txt(txt_原卡卡号)
        Else
            zlControl.ControlSetFocus txt(txt_原卡密码)
        End If
    Case gEd_补卡
        zlControl.ControlSetFocus txt(txt_原卡卡号)
    Case Else
        zlControl.ControlSetFocus txt(txt_缴款)
    End Select
End Sub

Private Function Load限制类别(ByVal str限制类别 As String) As Boolean
    '加载限制类别
    '入参：
    '   str限制类别 - 格式：西成药,中成药,...
    Dim i As Long, j As Long
    Dim varTemp As Variant, blnFind As Boolean
    Dim objItem As ListItem
    
    On Error GoTo ErrHandler
    For j = 1 To lvw限制类别.ListItems.count
        lvw限制类别.ListItems(j).Checked = False
    Next
    
    If str限制类别 = "" Then Load限制类别 = True: Exit Function
     
    varTemp = Split(str限制类别, ",")
    For i = 0 To UBound(varTemp)
        blnFind = False
        For j = 1 To lvw限制类别.ListItems.count
            If varTemp(i) = lvw限制类别.ListItems(j).Key Then
                lvw限制类别.ListItems(j).Checked = True
                blnFind = True: Exit For
            End If
        Next
        
        If blnFind = False Then
            Set objItem = lvw限制类别.ListItems.Add(, varTemp(i), varTemp(i))
            objItem.Checked = True
        End If
    Next
    Load限制类别 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitModulePara() As Boolean
    '初始化模块变量
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim str共用批次 As String, varData As Variant, varTemp As Variant
    Dim i As Integer
    Dim ty_Temp As Ty_CardType
    
    On Error GoTo ErrHandler
    Set rsTemp = zlGet消费卡接口()
    rsTemp.Filter = "编号=" & mlng卡类别
    If rsTemp.EOF Then
        ShowMsgbox "未发现卡类别信息，不能继续！"
        Exit Function
    End If
    
    mCardType = ty_Temp '自定义Type初始化
    With mCardType
        .str卡名称 = NVL(rsTemp!名称)
        .str卡号前缀 = NVL(rsTemp!前缀文本)
        .lng卡号长度 = Val(NVL(rsTemp!卡号长度))
        .bln卡号密文 = Val(NVL(rsTemp!是否密文)) = 1
        .int密码长度 = Val(NVL(rsTemp!密码长度))
        .int密码长度限制 = Val(NVL(rsTemp!密码长度限制))
        .bln严格控制 = Val(NVL(rsTemp!是否严格控制)) = 1
        .byt密码规则 = Val(NVL(rsTemp!密码规则))
        .str限制类别 = NVL(rsTemp!限制类别)
        .bln特定病人 = Val(NVL(rsTemp!是否特定病人)) = 1
    End With
    
    strSQL = "Select 编码, 名称, 缺省面额, 缺省折扣, 缺省标志 From 消费卡类型"
    Set mrs卡类型 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrs卡类型.RecordCount = 0 Then
        ShowMsgbox "没有设置相关的消费卡类型，请在[字典管理]中设置！"
        Exit Function
    End If

    With mTy_MoudlePara
        .bln缴款单打印 = Val(zlDatabase.GetPara("缴款单打印", glngSys, mlngModule)) = 1
        .bln连续充值 = Val(zlDatabase.GetPara("连续充值", glngSys, mlngModule)) = 1
    End With
    
    str共用批次 = zlDatabase.GetPara("共用消费卡批次", glngSys, mlngModule)
    '领用ID,卡类别ID|...
    varData = Split(str共用批次, "|")
    For i = 0 To UBound(varData)
         varTemp = Split(varData(i), ",")
         If Val(varTemp(0)) <> 0 Then
            If Val(varTemp(1)) = mlng卡类别 Then
                mCardType.lng共用批次 = Val(varTemp(0)): Exit For
            End If
         End If
    Next
    
    If mEditType = gEd_发卡 Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡 _
        Or mEditType = gEd_充值 Or mEditType = gEd_充值回退 Then
        If Init支付方式() = False Then Exit Function
    End If
    
    InitModulePara = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnFirst Or mblnChange = False Then Exit Sub
    If mEditType = gEd_发卡 Or mEditType = gEd_修改 Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
    If Not mrsBalance Is Nothing Then Set mrsBalance = Nothing
End Sub

Private Sub IDKind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub IDKind补卡_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt_Change(Index As Integer)
    Dim lng卡张数 As Long
    
    If mblnNotClick Then Exit Sub
    mblnChange = True
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_开始卡号
        mlng卡ID = 0
        Set mobjCard = New clsSquareCard
        
        '显示计算当前卡张数
        If (mEditType = gEd_发卡 Or mEditType = gEd_回收) And txt(txt_开始卡号).Tag <> "" Then
            txt(txt_开始卡号).Tag = ""
            Call SetLblCaption(lbl_卡张数, mEditType = gEd_回收)
            If mEditType = gEd_发卡 Then Call Calc实收合计
        End If
        
        txt(txt_结束卡号).Text = ""
        txt(txt_结束卡号).Enabled = txt(txt_开始卡号) <> ""
        Call zl_SetCtlBackColor(txt(txt_结束卡号), Me)
    Case txt_结束卡号
        '显示计算当前卡张数
        If mEditType = gEd_发卡 Or mEditType = gEd_回收 Then
            lng卡张数 = Val(txt(txt_开始卡号).Tag)
            If lng卡张数 <> 0 And Val(txt(txt_结束卡号).Tag) <> 0 Then
                txt(txt_开始卡号).Tag = "1"
                Call SetLblCaption(lbl_卡张数, mEditType = gEd_回收)
                If mEditType = gEd_发卡 Then Call Calc实收合计
            End If
            txt(txt_结束卡号).Tag = ""
        End If
    Case txt_缴款
        Call SetLblCaption(lbl_找补)
    Case txt_领卡人
        txt(Index).Tag = ""
        If (mEditType = gEd_发卡 Or mEditType = gEd_修改) And mCardType.bln特定病人 Then
            IDKind.SetAutoReadCard txt(Index).Text = ""
        End If
    Case txt_领卡部门
        txt(Index).Tag = ""
    Case txt_补卡姓名
        txt(Index).Tag = ""
        If mEditType = gEd_补卡 Then
            IDKind补卡.SetAutoReadCard txt(Index).Text = ""
        End If
    Case txt_密码
        txt(txt_确认密码) = ""
    Case txt_新卡密码
        txt(txt_新卡确认密码) = ""
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set可否充值()
    '设置可否充值
    Dim blnEnabled As Boolean
    
    On Error GoTo ErrHandler
    blnEnabled = (chk充值.value = vbChecked)
    blnEnabled = blnEnabled And IIf(mEditType = gEd_发卡, zlStr.IsHavePrivs(mstrPrivs, "充值"), mEditType = gEd_充值)
    
    txt(txt_充值扣率).Enabled = blnEnabled
    txt(txt_本次充值).Enabled = blnEnabled
    txt(txt_充值缴款).Enabled = blnEnabled
    txt(txt_充值说明).Enabled = blnEnabled
    
    Call zl_SetCtlBackColor(Array(txt(txt_充值扣率), txt(txt_本次充值), txt(txt_充值缴款), txt(txt_充值说明)), Me)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    If txt(Index).Enabled = False Or txt(Index).Locked Then Exit Sub
    Select Case Index
    Case txt_开始卡号, txt_结束卡号, txt_原卡卡号, txt_新卡卡号
        
    Case txt_密码, txt_原卡密码, txt_新卡密码
        Call OpenPassKeyboard(txt(Index), False)
    Case txt_确认密码, txt_新卡确认密码
        Call OpenPassKeyboard(txt(Index), True)
    Case txt_领卡人
        If (mEditType = gEd_发卡 Or mEditType = gEd_修改) And mCardType.bln特定病人 Then
            IDKind.SetAutoReadCard txt(Index).Text = ""
        End If
    Case txt_补卡姓名
        If mEditType = gEd_补卡 Then
            IDKind补卡.SetAutoReadCard txt(Index).Text = ""
        End If
    End Select
    zlControl.TxtSelAll txt(Index)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str编码 As String, str名称 As String, lngID As Long
    Dim strCardNo As String, intIndexTemp As Integer
    
    On Error GoTo ErrHandler
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case txt_领卡人
        If txt(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txt(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        
        If mCardType.bln特定病人 = False Then
            '选择人员
            lngID = Val(txt(txt_领卡部门).Tag)
            If Select人员选择器(Me, txt(Index), Trim(txt(Index).Text), lngID, , True, , , , , , , "") = False Then
                zlCommFun.PressKey vbKeyTab
            End If
            If mEditType = gEd_发卡 Then
                '领卡人就是缴款人
                txt(txt_缴款人).Text = txt(txt_领卡人).Text
                txt(txt_缴款人).Tag = txt(txt_领卡人).Tag
            End If
    
            '需要读取缺省部门:
            If zl_From人员获取缺省部门(Val(txt(txt_领卡人).Tag), str编码, str名称, lngID) Then
                txt(txt_领卡部门).Text = str编码 & "-" & str名称
                txt(txt_领卡部门).Tag = lngID
            End If
        End If
    Case txt_领卡部门
        '选择部门
        If txt(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txt(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        '选择缺省部门
        lngID = Val(txt(txt_领卡人).Tag)
        If Select部门选择器(Me, txt(Index), Trim(txt(Index).Text), "", IIf(lngID = 0, False, True), "", 0, _
            "部门选择器", , , , , lngID) = False Then Exit Sub
    Case txt_发卡原因
        '选择发卡原因
        If Trim(txt(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        
        If zl_SelectAndNotAddItem(Me, txt(Index), Trim(txt(Index).Text), "常用发卡原因", _
            "常用发卡原因选择", True, True, , , , True) = False Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    Case txt_密码, txt_新卡密码
        If CheckPassword(txt(Index), , True) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_确认密码, txt_新卡确认密码
        intIndexTemp = IIf(Index = txt_确认密码, txt_密码, txt_新卡密码)
        If CheckPassword(txt(Index), txt(intIndexTemp)) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_卡面额
        If CheckInput卡面额 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_销售额
        If CheckInput实际销售额 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_充值缴款
        If CheckInput实际充值缴款 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_本次充值
        If CheckInput本次充值 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_充值扣率
        If CheckInput充值扣率 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_开始卡号, txt_结束卡号, txt_原卡卡号, txt_新卡卡号
        
    Case txt_领卡人, txt_补卡姓名
        
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '创建密码键盘
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    If mobjKeyboard Is Nothing Then Exit Function
    CreateObjectKeyboard = True
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '打开密码键盘输入
    On Error GoTo ErrHandler
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '关闭密码键盘输入
    On Error GoTo ErrHandler
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput卡面额() As Boolean
    '检查卡面额
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_卡面额).Text), 16, True, False, txt(txt_卡面额).hWnd, "卡面额") = False Then
        zlControl.TxtSelAll txt(txt_卡面额): Exit Function
    End If

    If Val(txt(txt_卡面额).Text) < Val(txt(txt_销售额).Text) Then
        txt(txt_销售额).Text = txt(txt_卡面额).Text
    End If
    CheckInput卡面额 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput实际销售额() As Boolean
    '检查实际销售额
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_销售额).Text), 16, True, False, txt(txt_销售额).hWnd, "实际销售") = False Then
        zlControl.TxtSelAll txt(txt_销售额): Exit Function
    End If
    If Val(txt(txt_卡面额).Text) < Val(txt(txt_销售额).Text) Then
        ShowMsgbox "实际销售额不能大于卡面额，请检查！"
        zlControl.ControlSetFocus txt(txt_销售额)
        zlControl.TxtSelAll txt(txt_销售额)
        Exit Function
    End If
    CheckInput实际销售额 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput充值扣率() As Boolean
    '检查充值扣率是否合法
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_充值扣率).Text), 3, True, False, txt(txt_充值扣率).hWnd, "充值扣率") = False Then
        zlControl.TxtSelAll txt(txt_充值扣率): Exit Function
    End If
    If Val(txt(txt_充值扣率).Text) > 100 Then
        ShowMsgbox "充值扣率不能大于100%，请检查！"
        zlControl.ControlSetFocus txt(txt_充值扣率)
        zlControl.TxtSelAll txt(txt_充值扣率): Exit Function
    End If
    If Val(txt(txt_充值扣率).Text) < 0 Then
        ShowMsgbox "充值扣率不能小于0，请检查！"
        zlControl.ControlSetFocus txt(txt_充值扣率)
        zlControl.TxtSelAll txt(txt_充值扣率): Exit Function
    End If
    CheckInput充值扣率 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput本次充值() As Boolean
    '检查本次充值
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_本次充值).Text), 16, True, False, txt(txt_本次充值).hWnd, "本次充值") = False Then
        zlControl.TxtSelAll txt(txt_本次充值): Exit Function
    End If
    If Val(txt(txt_本次充值).Text) < Val(txt(txt_充值缴款).Text) Then
        txt(txt_充值缴款).Text = txt(txt_本次充值).Text
    End If
    CheckInput本次充值 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput实际充值缴款() As Boolean
    '检查本次充值
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_充值缴款).Text), 16, True, False, txt(txt_充值缴款).hWnd, "实际充值缴款") = False Then
        zlControl.TxtSelAll txt(txt_充值缴款): Exit Function
    End If
    If Val(txt(txt_本次充值).Text) < Val(txt(txt_充值缴款).Text) Then
        ShowMsgbox "实际充值缴款不能大于本次充值，请检查！"
        zlControl.ControlSetFocus txt(txt_充值缴款)
        zlControl.TxtSelAll txt(txt_充值缴款): Exit Function
    End If
    CheckInput实际充值缴款 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCard As Boolean, blnPass As Boolean
    Dim str姓名 As String, lng病人ID As Long
    Dim objIDKind As IDKindNew
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_充值扣率
        '只允许输入整数
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    Case txt_卡面额, txt_销售额, txt_本次充值, txt_充值缴款, txt_缴款
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m金额式)
    Case txt_开始卡号, txt_结束卡号, txt_原卡卡号, txt_新卡卡号
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m文本式)
        If InStr(1, "'~～|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        
        If Index = txt_结束卡号 Then
            If KeyAscii = 13 And Trim(txt(Index)) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
        
        '小写字母转换为大写
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        Call BrushCard(txt(Index), KeyAscii)
    Case txt_领卡人, txt_补卡姓名
        If mCardType.bln特定病人 Then
            If Index = txt_领卡人 Then
                Set objIDKind = IDKind
            Else
                Set objIDKind = IDKind补卡
            End If
            
            If IsCardType(objIDKind, "姓名") Then
                '105567:李南春,2017/5/25,卡号加密导致第一个汉字拼音不能触发输入法
                blnPass = txt(Index).PasswordChar <> ""
                If Not (InStr("-+*", Left(txt(Index).Text, 1)) > 0 And IsNumeric(Mid(txt(Index).Text, 2))) Then
                    blnCard = zlCommFun.InputIsCard(txt(Index), KeyAscii, objIDKind.ShowPassText)
                End If
                txt(Index).IMEMode = 0
                blnPass = blnPass And txt(Index).PasswordChar = ""
                If blnPass Then
                    If txt(Index).SelLength = Len(txt(Index).Text) Then
                        txt(Index).Text = ""
                    End If
                    SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
                End If
            ElseIf IsCardType(objIDKind, "门诊号") Or IsCardType(objIDKind, "住院号") Or IsCardType(objIDKind, "手机号") Then
                If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
                End If
            Else
                txt(Index).PasswordChar = IIf(objIDKind.ShowPassText, "*", "")
                '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
                txt(Index).IMEMode = 0
            End If
        
            If blnCard And Len(txt(Index).Text) = objIDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
                Or KeyAscii = 13 And Trim(txt(Index).Text) <> "" Then
                If KeyAscii <> 13 Then
                    txt(Index).Text = txt(Index).Text & Chr(KeyAscii): txt(Index).SelStart = Len(txt(Index).Text)
                End If
                KeyAscii = 0
                
                strTemp = txt(Index).Text
                If Not GetPatient(objIDKind, txt(Index), txt(Index).Text, blnCard, str姓名, lng病人ID) Then
                    '清除数据
                    If LoadPatientCard(mlng卡类别, lng病人ID) = False Then
                        txt(Index).Text = strTemp
                        zlControl.TxtSelAll txt(Index)
                        Exit Sub
                    End If
                Else
                    If Index = txt_补卡姓名 Then
                        '加载病人有效卡
                        If LoadPatientCard(mlng卡类别, lng病人ID) = False Then
                            txt(Index).Text = strTemp
                            zlControl.TxtSelAll txt(Index)
                            Exit Sub
                        End If
                    End If
                    txt(Index).Text = str姓名
                    txt(Index).Tag = lng病人ID
                    txt(Index).PasswordChar = ""
                    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
                    txt(Index).IMEMode = 0
                    zlCommFun.PressKey vbKeyTab: Exit Sub
                End If
            End If
        End If
    Case txt_密码, txt_确认密码, txt_新卡密码, txt_新卡确认密码
        Call CheckInputPassWord(KeyAscii, mCardType.byt密码规则 = 1)
    Case Else
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m文本式)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatient(ByVal objIDKind As IDKindNew, txtEdit As TextBox, _
    ByVal strInput As String, ByVal blnCard As Boolean, _
    ByRef str姓名 As String, ByRef lng病人ID As Long) As Boolean
    '获取病人信息
    '入参:blnCard=表示是否就诊卡刷卡
    '出参:
    '返回:病人读取成功,返回true,否则返回False
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String
    Dim lng卡类别ID As Long, blnCancel As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim blnIsMobileNO As Boolean
    
    On Error GoTo errH
    str姓名 = "": lng病人ID = 0
    blnIsMobileNO = IDKind.IsMobileNO(strInput)
    If blnCard And IsCardType(objIDKind, "姓名") And (InStr("-+*", Left(strInput, 1)) = 0 And IsNumeric(Mid(strInput, 2))) Then  '刷卡或缺省的卡
        If objIDKind.Cards.按缺省卡查找 And Not objIDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = objIDKind.GetfaultCard.接口序号
        ElseIf objIDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = objIDKind.GetCurCard.接口序号
        Else
            lng卡类别ID = -1
        End If
        
        '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If GetPatiID(lng卡类别ID, strInput, True, lng病人ID, strPassWord, strErrMsg) = False Then
            If blnIsMobileNO Then
                '手机号查找
                If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            Else
                GoTo NotFoundPati:
            End If
        End If
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strWhere = strWhere & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strWhere = strWhere & " And A.住院号=[1]"
    Else
        Select Case objIDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                strPati = _
                    "Select a.病人id As ID, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号," & vbNewLine & _
                    "       a.出生日期, a.身份证号, a.家庭地址, a.工作单位" & vbNewLine & _
                    "From 病人信息 A" & vbNewLine & _
                    "Where a.停用时间 Is Null And a.姓名 Like [1] And Rownum < 101" & vbNewLine & _
                    "Order By 姓名"
                
                vRect = zlControl.GetControlRect(txtEdit.hWnd)
                Set rsTemp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人选择", 1, "", "请选择病人", False, False, True, _
                    vRect.Left, vRect.Top, txtEdit.Height, blnCancel, False, True, strInput & "%")
                If blnCancel Then Exit Function
                If rsTemp Is Nothing Then GoTo NotFoundPati:
                If rsTemp.State <> 1 Then GoTo NotFoundPati:
                If rsTemp.RecordCount = 0 Then GoTo NotFoundPati:
                If Val(NVL(rsTemp!病人ID)) = 0 Then GoTo NotFoundPati:
                
                strInput = "-" & rsTemp!病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.医保号=[2]"
             Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                '问题号:54197
                 If GetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , False) = False Then lng病人ID = 0
                 strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If GetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case "手机号", "手机"
                If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case Else
                '其他类别的号码
                If Val(objIDKind.GetCurCard.接口序号) >= 0 Then
                    lng卡类别ID = objIDKind.GetCurCard.接口序号
                    If GetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If GetPatiID(objIDKind.GetCurCard.名称, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
        End Select
    End If
    
    '读取病人信息
    strSQL = "Select A.病人id,A.姓名 From 病人信息 A Where A.停用时间 is NULL" & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If rsTemp.EOF Then GoTo NotFoundPati:
    
    str姓名 = NVL(rsTemp!姓名)
    lng病人ID = Val(NVL(rsTemp!病人ID))
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
NotFoundPati:
    If blnCard Then
        MsgBox "不能确定病人信息，请检查是否正确刷卡！    ", vbInformation + vbOKOnly, gstrSysName
    Else
        MsgBox "病人信息未找到，请检查是否输入正确！", vbInformation + vbOKOnly, gstrSysName
    End If
End Function

Private Function LoadPatientCard(ByVal lng卡类别 As Long, ByVal lng病人ID As Long) As Boolean
    '功能：加载病人当前有效消费卡
    '入参：
    '   lng卡类别 消费卡类别编号
    '   lng病人ID 病人ID
    '出参：
    '返回：获取到有效消费卡则返回TRUE,否则返回FALSE
    '说明：
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mEditType <> gEd_补卡 Then LoadPatientCard = True: Exit Function
    Call ClearCtlData
    mlng卡ID = 0
    cbo原卡卡号.Clear
    cbo原卡卡号.Tag = ""
    
    If lng病人ID = 0 Then Exit Function
    strSQL = _
        "Select a.Id, a.卡号" & vbNewLine & _
        "From 消费卡信息 A" & vbNewLine & _
        "Where a.序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = a.卡号 And 接口编号 = a.接口编号)" & vbNewLine & _
        "      And a.当前状态 = 1 And a.接口编号 = [1] And a.病人id = [2]" & vbNewLine & _
        "Order By a.发卡时间 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡类别, lng病人ID)
        If rsTemp.EOF Then
        ShowMsgbox "未查找到属于该病人的有效的" & mCardType.str卡名称 & "卡片信息！"
        Exit Function
    End If
    
    Do While Not rsTemp.EOF
        cbo原卡卡号.AddItem NVL(rsTemp!卡号)
        cbo原卡卡号.ItemData(cbo原卡卡号.NewIndex) = Val(NVL(rsTemp!id))
        rsTemp.MoveNext
    Loop
    cbo原卡卡号.ListIndex = 0
    
    LoadPatientCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '刷卡
    Dim blnCard As Boolean
    Dim lng卡张数 As Long
    
    On Error GoTo ErrHandler
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If (mEditType = gEd_发卡 Or (mEditType = gEd_换卡 Or mEditType = gEd_补卡) And objEdit.Index = txt_新卡卡号) _
            And Len(objEdit.Text) = mCardType.lng卡号长度 - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(objEdit.Text) <> "" Then '新卡达到卡号长度或其它回车查找卡信息
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        
        If mEditType <> gEd_发卡 _
            And Not ((mEditType = gEd_换卡 Or mEditType = gEd_补卡) And objEdit.Index = txt_新卡卡号) Then
            If LoadCardData(1, objEdit.Text) = False Then zlControl.TxtSelAll objEdit: Exit Sub
            If mEditType = gEd_充值 Then
                txt(txt_充值扣率).Enabled = (chk充值.value = vbChecked)
                txt(txt_本次充值).Enabled = (chk充值.value = vbChecked)
                txt(txt_充值缴款).Enabled = (chk充值.value = vbChecked)
                txt(txt_充值说明).Enabled = (chk充值.value = vbChecked)
                Call zl_SetCtlBackColor(Array(txt(txt_充值扣率), txt(txt_本次充值), txt(txt_充值缴款), txt(txt_充值说明)), Me)
            End If
        Else
            If CheckInput卡号(False, lng卡张数) = False Then zlControl.TxtSelAll objEdit: Exit Sub
        End If
        
        '显示计算当前卡张数
        If mEditType = gEd_发卡 Then
            txt(txt_开始卡号).Tag = lng卡张数
            txt(txt_结束卡号).Tag = IIf(lng卡张数 > 1, "1", "") '标记是否批量发卡
            Call SetLblCaption(lbl_卡张数)
            Call Calc实收合计
        ElseIf mEditType = gEd_回收 Then
            txt(txt_开始卡号).Tag = 1
            Call SetLblCaption(lbl_卡张数, True)
        End If
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = 13 And Trim(objEdit.Text) = "" Then
        zlCommFun.PressKey vbKeyTab
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Dim intIndexTemp As Integer
    Dim strPassWord As String
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_开始卡号, txt_结束卡号, txt_新卡卡号
        If Not (mEditType <> gEd_发卡 And Index = txt_开始卡号) Then
            If Trim(txt(Index).Text) <> "" And Len(Trim(txt(Index).Text)) <> mCardType.lng卡号长度 Then
                ShowMsgbox "卡号长度应为" & mCardType.lng卡号长度 & "位，请检查！"
                zlControl.ControlSetFocus txt(Index)
                zlControl.TxtSelAll txt(Index)
            End If
        End If
    Case txt_卡面额
        If CheckInput卡面额 = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        Call Calc余额
        Calc实收合计
    Case txt_销售额
        If CheckInput实际销售额 = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        Call Calc余额
        Calc实收合计
    Case txt_本次充值
        If CheckInput本次充值 = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        txt(txt_充值缴款).Text = Format(Val(txt(Index).Text) * (Round(Val(txt(txt_充值扣率)) / 100, 6)), "0.00")
        Call Calc余额
        Calc实收合计
    Case txt_充值扣率
        If CheckInput充值扣率 = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        txt(txt_充值缴款).Text = Format(Val(txt(txt_本次充值).Text) * (Round(Val(txt(txt_充值扣率)) / 100, 4)), "0.00")
        Call Calc余额
        Calc实收合计
    Case txt_充值缴款
        If CheckInput实际充值缴款 = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        If Val(txt(txt_本次充值).Text) <> 0 Then
            txt(txt_充值扣率).Text = Format((Round(Val(txt(txt_充值缴款).Text) / Val(txt(txt_本次充值).Text), 6)) * 100, "0.00")
        Else
             txt(txt_本次充值).Text = txt(txt_充值缴款).Text
        End If
        Call Calc余额
        Calc实收合计
    Case txt_密码, txt_新卡密码
        If CheckPassword(txt(Index), , True) = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
    Case txt_确认密码, txt_新卡确认密码
        intIndexTemp = IIf(Index = txt_确认密码, txt_密码, txt_新卡密码)
        If CheckPassword(txt(Index), txt(intIndexTemp)) = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
    Case txt_原卡密码
        If Trim(txt(Index).Text) <> "" Then
            strPassWord = zlCommFun.zlStringEncode(txt(Index).Text)  '密码加密
            If mobjCard.原密码 <> strPassWord Then
                ShowMsgbox "原卡密码输入错误，请重新输入！"
                zlControl.ControlSetFocus txt(Index)
                zlControl.TxtSelAll txt(Index)
            End If
        End If
    Case txt_领卡人
        If (mEditType = gEd_发卡 Or mEditType = gEd_修改) And mCardType.bln特定病人 Then
            IDKind.SetAutoReadCard False
        End If
    Case txt_领卡部门
        If txt(Index).Tag = "" Then txt(Index).Text = ""
    Case txt_密码, txt_确认密码, txt_原卡密码, txt_新卡密码, txt_新卡确认密码
        Call ClosePassKeyboard(txt(Index))
    Case txt_补卡姓名
        If mEditType = gEd_补卡 Then
            IDKind补卡.SetAutoReadCard False
        End If
    End Select
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub Calc实收合计(Optional blnOnlyCalc As Boolean)
    '计算实收合计
    Dim dbl实收合计 As Double
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If Not (mEditType = gEd_发卡 _
        Or mEditType = gEd_充值 Or mEditType = gEd_充值回退 _
        Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡) Then Exit Sub
    
    If mEditType = gEd_发卡 Or mEditType = gEd_退卡 Or mEditType = gEd_取消退卡 Then
        dbl实收合计 = Val(txt(txt_销售额).Text)
    End If
    If chk充值.value = vbChecked Then
        dbl实收合计 = dbl实收合计 + Val(txt(txt_充值缴款).Text)
    End If
    
    If mEditType = gEd_发卡 Then
        dbl实收合计 = dbl实收合计 * Val(lbl(lbl_卡张数2).Caption)
    End If
    
    If mEditType = gEd_充值回退 Or mEditType = gEd_退卡 Then
        dbl实收合计 = -1 * dbl实收合计 '退款
    End If
    
    mdbl实收合计 = dbl实收合计
    If blnOnlyCalc Then Exit Sub
    
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Calc余额()
    '计算余额
    Dim dbl余额 As Double
    
    On Error GoTo ErrHandler
    If Not (mEditType = gEd_发卡 Or mEditType = gEd_取消退卡 Or mEditType = gEd_充值) Then Exit Sub
    If mEditType = gEd_发卡 Or mEditType = gEd_取消退卡 Then
        dbl余额 = Val(txt(txt_卡面额).Text)
    Else
        dbl余额 = mobjCard.卡余额
    End If
    dbl余额 = dbl余额 + IIf(chk充值.value = 0, 0, Val(txt(txt_本次充值).Text))
    lbl(lbl_卡余额).Caption = Format(dbl余额, "0.00")
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo卡类型_Click()
    If mblnNotClick Then Exit Sub
    mblnChange = True
    
    On Error GoTo ErrHandler
    '重新设置缺省值
    If mEditType <> gEd_发卡 Then Exit Sub
    Call SetDefaultValue
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk充值_Click()
    If mblnNotClick Then Exit Sub
    
    On Error GoTo ErrHandler
    mblnChange = True
    Call Set可否充值
    Call Calc余额
    Call Calc实收合计
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDefaultValue()
    '设置缺省值
    On Error GoTo ErrHandler
    mrs卡类型.Filter = "编码='" & zlStr.NeedCode(cbo卡类型.Text) & "'"
    If mrs卡类型.EOF Then Exit Sub
    
    txt(txt_充值扣率).Text = Format(Val(NVL(mrs卡类型!缺省折扣, 100)), "0.00")
    txt(txt_充值扣率).Tag = txt(txt_充值扣率).Text
    txt(txt_充值缴款).Text = Format(Val(NVL(mrs卡类型!缺省折扣, 100)) * Val(txt(txt_本次充值).Text) / 100, "0.00")
    
    txt(txt_卡面额).Text = Format(Val(NVL(mrs卡类型!缺省面额)), "0.00")
    txt(txt_销售额).Text = Format(Val(NVL(mrs卡类型!缺省面额)) * (txt(txt_充值扣率).Text / 100), "0.00")
    
    Call Calc余额
    Call Calc实收合计
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Init支付方式() As Boolean
    '初始化支付方式
    '说明：
    '   只加入现金、支票和三方卡的结算方式
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim i As Long, objCards As Cards, objCard As Card
    Dim lngKey As Long
    
    On Error GoTo ErrHandler
    Set mobjPayCards = New Cards
    
    Set rsTemp = Get结算方式("消费卡")
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
            '入参:bytType-  0-所有医疗卡;
        '                        1-启用的医疗卡,
        '                        2-所有存在三方账户的三方卡
        '                        3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.count
                If objCards(i).结算方式 = NVL(rsTemp!名称) Then blnFind = True: Exit For
            Next
            If Not blnFind Then
                If (Val(NVL(rsTemp!性质)) = 1 Or Val(NVL(rsTemp!性质)) = 2) _
                    And Val(NVL(rsTemp!应付款)) = 0 Then
                    Set objCard = New Card
                    objCard.短名 = Mid(NVL(!名称), 1, 1)
                    objCard.接口编码 = NVL(!编码)
                    objCard.接口程序名 = ""
                    objCard.接口序号 = -1 * lngKey
                    objCard.结算方式 = NVL(!名称)
                    objCard.名称 = NVL(!名称)
                    objCard.启用 = True
                    objCard.缺省标志 = Val(NVL(rsTemp!缺省)) = 1
                    objCard.启用 = True
                    objCard.结算性质 = Val(!性质)
                    
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '加三方卡
    For i = 1 To objCards.count
        rsTemp.Filter = "名称='" & objCards(i).结算方式 & "'"
        If Not rsTemp.EOF And Not objCards(i).消费卡 Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.count = 0 Then
        ShowMsgbox "消费卡场合没有可用的结算方式，请先到【结算方式管理】中设置。"
        Exit Function
    End If
    Init支付方式 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load支付方式(Optional ByVal blnDel As Boolean) As Boolean
    '加载支付方式
    '说明:
    '   缺省结算方式的规则，优先顺序如下：
    '   1.结算方式应用中设置的缺省项
    '   2.性质为"1-现金结算方式"的结算方式
    Dim objCard As Card, i As Long
    Dim str结算方式 As String, blnExists As Boolean
    
    On Error GoTo ErrHandler
    mlngPre支付方式 = 0

    mblnNotClick = True
    With cbo支付方式
        .Clear
        For i = 1 To mobjPayCards.count
            Set objCard = mobjPayCards(i)
            If objCard.启用 And Not objCard.消费卡 _
                And InStr(str结算方式 & "|", "|" & objCard.结算方式 & "|") = 0 Then
                '三方账户的支付方式显示为医疗卡名称，其它显示结算方式
                If objCard.接口序号 > 0 Then
                    If blnDel And objCard.是否转帐及代扣 = False Then
                        If ExitsInBalance(objCard.结算方式) Then
                            .AddItem objCard.名称
                            .ItemData(.NewIndex) = i
                            If Not (objCard.是否退现 And objCard.是否缺省退现) Then .ListIndex = .NewIndex
                        End If
                    Else
                        .AddItem objCard.名称
                        .ItemData(.NewIndex) = i
                    End If
                Else
                    .AddItem objCard.结算方式
                    .ItemData(.NewIndex) = i
                    If blnDel Then
                        If ExitsInBalance(objCard.结算方式) Then .ListIndex = .NewIndex
                    End If
                End If
                
                str结算方式 = str结算方式 & "|" & objCard.结算方式
            End If
            
            '设置缺省值
            If objCard.缺省标志 And .ListIndex < 0 Then .ListIndex = .NewIndex
            If objCard.结算性质 = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
        Next
            
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Call cbo支付方式_Click
    Load支付方式 = True
    Exit Function
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExitsInBalance(ByVal str结算方式 As String) As Boolean
    '判断结算方式是否存在于收款结算方式中
    On Error GoTo ErrHandler
    If mrsBalance Is Nothing Then Exit Function
    With mrsBalance
        .Filter = ""
        Do While Not .EOF
            If NVL(!结算方式) = str结算方式 Then
                ExitsInBalance = True: Exit Function
            End If
            .MoveNext
        Loop
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get限制类别() As String
    '获取限制类别
    Dim strType As String, i As Long
    
    On Error GoTo ErrHandler
    With lvw限制类别
         For i = 1 To .ListItems.count
            If .ListItems.Item(i).Checked Then
                strType = strType & "," & lvw限制类别.ListItems(i).Key
            End If
         Next
         If strType <> "" Then strType = Mid(strType, 2)
    End With
    Get限制类别 = strType
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '获取当前支付卡
    '出参:
    '   objCard-返回当前退款或缴款的卡对象
    '返回:成功,返回卡对象
    Dim intIndex As Integer
    
    On Error GoTo ErrHandler
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    GetCurCard = True
    Exit Function
ErrHandler:
    Set objCard = New Card
End Function
 
'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    On Error GoTo ErrHandler
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case "住院号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "住院号"
     Case "手机号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "手机号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.名称
            Else
                If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
            End If
     End Select
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
    '数据有效性检查
    Dim strPassWord As String
    
    On Error GoTo ErrHandler
    Select Case mEditType
    Case gEd_发卡
        If CheckInput卡号(True) = False Then Exit Function
        If CheckInput() = False Then Exit Function
        If Check缴款情况 = False Then Exit Function
    Case gEd_修改
        If CheckInput卡号(True) = False Then
            zlControl.ControlSetFocus txt(txt_开始卡号)
            Exit Function
        End If
        If CheckInput() = False Then Exit Function
    Case gEd_充值
        If Trim(txt(txt_开始卡号)) = "" Or mlng卡ID = 0 Then
            ShowMsgbox "请刷卡或输入充值卡号！"
            zlControl.ControlSetFocus txt(txt_开始卡号)
            Exit Function
        End If
        If mEditType = gEd_充值 And Val(txt(txt_本次充值).Text) = 0 Then
            ShowMsgbox "充值金额不能为零！"
            zlControl.ControlSetFocus txt(txt_本次充值)
            Exit Function
        End If
        If CheckInput卡号(True) = False Then
            zlControl.ControlSetFocus txt(txt_开始卡号)
            Exit Function
        End If
       If Check缴款情况 = False Then Exit Function
    Case gEd_充值回退
        If CheckInput卡号(True) = False Then
            zlControl.ControlSetFocus txt(txt_开始卡号)
            Exit Function
        End If
       If Check缴款情况 = False Then Exit Function
    Case gEd_退卡
        If CheckInput卡号(True) = False Then
            zlControl.ControlSetFocus txt(txt_开始卡号)
            Exit Function
        End If
        If Check缴款情况 = False Then Exit Function
    Case gEd_取消退卡
        If CheckInput卡号(True) = False Then
            zlControl.ControlSetFocus txt(txt_开始卡号)
            Exit Function
        End If
        If Check缴款情况 = False Then Exit Function
    Case gEd_回收
        If CheckInput卡号(True) = False Then Exit Function
    Case gEd_取消回收
        If CheckInput卡号(True) = False Then Exit Function
    Case gEd_换卡, gEd_补卡
        If mEditType = gEd_换卡 Then
            If Trim(txt(txt_原卡卡号)) = "" Or mlng卡ID = 0 Then
                ShowMsgbox "请刷卡或输入原卡卡号！"
                zlControl.ControlSetFocus txt(txt_原卡卡号)
                Exit Function
            End If
            
            If CheckInput卡号(True) = False Then Exit Function
            
            strPassWord = zlCommFun.zlStringEncode(txt(txt_原卡密码).Text)  '密码加密
            If mobjCard.原密码 <> strPassWord Then
                ShowMsgbox "原卡密码输入错误，请重新输入！"
                zlControl.ControlSetFocus txt(txt_原卡密码)
                Exit Function
            End If
        Else
            If Val(txt(txt_补卡姓名).Tag) = 0 Then
                ShowMsgbox "请先录入病人，再选择需要补卡的原卡卡号！"
                zlControl.ControlSetFocus txt(txt_补卡姓名)
                Exit Function
            End If
            
            If cbo原卡卡号.Text = "" Or mlng卡ID = 0 Then
                ShowMsgbox "请选择原卡卡号！"
                zlControl.ControlSetFocus cbo原卡卡号
                Exit Function
            End If
            
            If CheckInput卡号(True) = False Then Exit Function
        End If
        
        If CheckPassword(txt(txt_新卡密码), txt(txt_新卡确认密码)) = False Then Exit Function
        If mobjCard.原密码 <> Trim(txt(txt_新卡密码).Text) Then
            If zlCommFun.StrIsValid(Trim(txt(txt_新卡密码).Text), 20, txt(txt_新卡密码).hWnd, "密码") = False Then Exit Function
            If zlCommFun.StrIsValid(Trim(txt(txt_新卡确认密码).Text), 20, txt(txt_新卡确认密码).hWnd, "确认密码") = False Then Exit Function
        End If
    End Select
    IsValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPassword(ByVal txtPass As TextBox, Optional ByVal txtVaild As TextBox, _
    Optional blnOnlyCheckOld As Boolean) As Boolean
    '密码有效性检查
    'blnOnlyCheckOld 是否只检查原密码
    
    On Error GoTo ErrHandler
    If txtPass.Text = "" Or txtPass.Visible = False Then CheckPassword = True: Exit Function
    Select Case mCardType.int密码长度限制
    Case 0
    Case 1
        If Len(txtPass.Text) <> mCardType.int密码长度 Then
            ShowMsgbox "密码必须输入" & mCardType.int密码长度 & "位！"
            zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
            Exit Function
         End If
    Case Else
        If Len(txtPass.Text) <= Abs(mCardType.int密码长度限制) Then
            ShowMsgbox "密码必须输入" & Abs(mCardType.int密码长度限制) & "位以上！"
            zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
            Exit Function
         End If
    End Select
    If mCardType.byt密码规则 = 1 Then '密码只允许为数字
        If (txtPass.Index = txt_新卡密码 Or txtPass.Index = txt_新卡确认密码) And txtPass.Text = mobjCard.原密码 Then
            '特殊处理
        ElseIf IsNumeric(txtPass.Text) = False Then
            ShowMsgbox "密码只能包含数字，请重新输入！"
            zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
            Exit Function
        End If
    End If
    If blnOnlyCheckOld Then CheckPassword = True: Exit Function
    
    If txtPass.Text <> txtVaild.Text Then
        ShowMsgbox "两次输入的密码不一致，请检查！"
        zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
        Exit Function
    End If
    CheckPassword = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
                
Private Function SaveData() As Boolean
    '保存数据
    Dim lng发卡序号 As Long, lngID As Long
    
    Select Case mEditType
    Case gEd_发卡
        lng发卡序号 = zlDatabase.GetNextId("消费卡信息")
        If SavePayCard(lng发卡序号) = False Then Exit Function
        SaveData = True
        
        '打印缴款单
        If mTy_MoudlePara.bln缴款单打印 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
                "付款序号=" & lng发卡序号, "缴款=" & Val(txt(txt_缴款).Text), "找补=" & Val(txt(txt_找补).Tag), _
                "充值ID=0", "ReportFormat=1", 2)
        End If
    Case gEd_修改
        If SaveModifyCard = False Then Exit Function
    Case gEd_回收
        If SaveCallBack = False Then Exit Function
    Case gEd_取消回收
        If SaveCallBack(True) = False Then Exit Function
    Case gEd_退卡
        If SaveBackCard(False) = False Then Exit Function
    Case gEd_取消退卡
        If SaveBackCard(True) = False Then Exit Function
    Case gEd_充值
        If SaveInFull(lngID) = False Then Exit Function
        SaveData = True
        
        '打印缴款单
        If mTy_MoudlePara.bln缴款单打印 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
                "充值ID=" & lngID, "缴款=" & Val(txt(txt_缴款).Text), _
                "找补=" & Val(txt(txt_找补).Tag), "付款序号=0", "ReportFormat=2", 2)
        End If
    Case gEd_充值回退
        If SaveCancelInFull() = False Then Exit Function
    Case gEd_换卡
        If SaveChangeCard() = False Then Exit Function
    Case gEd_补卡
        If SaveReissueCard() = False Then Exit Function
    End Select
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveChangeCard() As Boolean
    '换卡
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_消费卡信息_换卡
    strSQL = "Zl_消费卡信息_换卡("
    '  原卡id_In   消费卡信息.Id%Type,
    strSQL = strSQL & "" & mlng卡ID & ","
    '  新卡卡号_In 消费卡信息.卡号%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_新卡卡号).Text) & "',"
    '  密码_In     消费卡信息.密码%Type,
    If txt(txt_新卡密码).Text = mobjCard.原密码 Then
        strSQL = strSQL & "'" & mobjCard.原密码 & "',"
    Else
        strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txt(txt_新卡密码).Text) & "',"
    End If
    '  发卡人_In   消费卡信息.发卡人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  发卡时间_In 消费卡信息.发卡时间%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  领用id_In   消费卡信息.领用id%Type := Null
    strSQL = strSQL & "" & mCardType.lng领用ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveChangeCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveReissueCard() As Boolean
    '补卡
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_消费卡信息_补卡
    strSQL = "Zl_消费卡信息_补卡("
    '  原卡id_In   消费卡信息.Id%Type,
    strSQL = strSQL & "" & mlng卡ID & ","
    '  新卡卡号_In 消费卡信息.卡号%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_新卡卡号).Text) & "',"
    '  密码_In     消费卡信息.密码%Type,
    If txt(txt_新卡密码).Text = mobjCard.原密码 Then
        strSQL = strSQL & "'" & mobjCard.原密码 & "',"
    Else
        strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txt(txt_新卡密码).Text) & "',"
    End If
    '  发卡人_In   消费卡信息.发卡人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  发卡时间_In 消费卡信息.发卡时间%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  领用id_In   消费卡信息.领用id%Type := Null
    strSQL = strSQL & "" & mCardType.lng领用ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveReissueCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveInFull(ByRef lngID As Long) As Boolean
    '保存充值处理
    '出参:lngID-返回本次的充值的ID
    '返回:充值成功,返回True,否则返回False
    Dim strSQL As String, blnTrain As Boolean
    Dim lng结算序号 As Long, objCard As Card
    
    Err = 0: On Error GoTo ErrHandler
    lngID = zlDatabase.GetNextId("病人卡结算记录")
    lng结算序号 = lngID
    'Zl_病人卡结算记录_充值
    strSQL = "Zl_病人卡结算记录_充值("
    '  Id_In         病人卡结算记录.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  消费卡id_In   病人卡结算记录.消费卡id%Type,
    strSQL = strSQL & "" & mlng卡ID & ","
    '  充值金额_In   病人卡结算记录.应收金额%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_本次充值).Text), 4) & ","
    '  充值折扣_In   病人卡结算记录.扣率%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_充值扣率).Text), 4) & ","
    '  缴款金额_In   病人卡结算记录.实收金额%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_充值缴款).Text), 4) & ","
    '  充值时间_In   病人卡结算记录.交易时间%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  操作员编号_In 病人卡结算记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  缴款人_In     病人卡结算记录.缴款人姓名%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_缴款人).Text) & "',"
    '  充值说明_In   病人卡结算记录.备注%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_充值说明).Text) & "',"
    '  结算方式_In     病人卡结算记录.结算方式%Type,
    strSQL = strSQL & "'" & mCurCardPay.str结算方式 & "',"
    '  结算序号_In   病人卡结算记录.Id%Type,
    strSQL = strSQL & "" & lng结算序号 & ","
    '  开户行_In       病人卡结算记录.单位开户行%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_开户行).Text) & "',"
    '  帐号_In         病人卡结算记录.单位帐号%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_帐号).Text) & "',"
    '  结算号码_In     病人卡结算记录.结算号码%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_结算号码).Text) & "',"
    '  卡类别id_In   病人卡结算记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", mCurCardPay.lng卡类别ID) & ","
    '  结算卡号_In   病人卡结算记录.结算卡号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人卡结算记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易流水号 & "'") & ","
    '  交易说明_In   病人卡结算记录.交易说明%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易说明 & "'") & ","
    '  缴款_In         病人卡结算记录.缴款%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_缴款).Text), 4), "NULL") & ","
    '  找补_In         病人卡结算记录.找补%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_找补).Tag), 4), "NULL") & ")"
    
    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '三方卡结算
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.接口序号 > 0 Then
        If ExecuteThreeSwapPay(objCard, lng结算序号, mdbl实收合计) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveInFull = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCancelInFull() As Boolean
    '充值回退
    Dim strSQL As String, blnTrain As Boolean
    Dim lng结算序号 As Long, objCard As Card
    
    Err = 0: On Error GoTo Errhand
    lng结算序号 = zlDatabase.GetNextId("病人卡结算记录")
    'Zl_病人卡结算记录_充值回退
    strSQL = "Zl_病人卡结算记录_充值回退("
    '  Id_In         病人卡结算记录.Id%Type,
    strSQL = strSQL & "" & mlng充值ID & ","
    '  操作员编号_In 病人卡结算记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  结算方式_In     病人卡结算记录.结算方式%Type,
    strSQL = strSQL & "'" & mCurCardPay.str结算方式 & "',"
    '  结算序号_In   病人卡结算记录.Id%Type,
    strSQL = strSQL & "" & lng结算序号 & ","
    '  误差金额_In   病人卡结算记录.实收金额%Type,
    strSQL = strSQL & "" & -1 * mdbl本次误差 & ","
    '  结算号码_In     病人卡结算记录.结算号码%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_结算号码).Text) & "',"
    '  开户行_In       病人卡结算记录.单位开户行%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_开户行).Text) & "',"
    '  帐号_In         病人卡结算记录.单位帐号%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_帐号).Text) & "',"
    '  卡类别id_In   病人卡结算记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", mCurCardPay.lng卡类别ID) & ","
    '  结算卡号_In   病人卡结算记录.结算卡号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人卡结算记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易流水号 & "'") & ","
    '  交易说明_In   病人卡结算记录.交易说明%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易说明 & "'") & ","
    '  缴款_In         病人卡结算记录.缴款%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_缴款).Text), 4), "NULL") & ","
    '  找补_In         病人卡结算记录.找补%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_找补).Tag), 4), "NULL") & ")"
    
    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '三方卡结算
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.接口序号 > 0 Then
        If ExecuteThreeSwapPay(objCard, lng结算序号, mdbl实收合计) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveCancelInFull = True
    Exit Function
Errhand:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckUsedBill(ByVal lng卡类别 As Long, ByVal lng领用ID As Long, _
    Optional ByVal strBill As String) As Long
    '功能：检查当前操作员是否有可用消费卡领用(自用或共用),并返回可用的领用ID
    '参数：
    '      lng卡类别=消费卡接口编号
    '      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
    '      strBill=要检查范围的票据号
    '说明：
    '    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
    '    2.在检查范围时,长度也在检查范围之内。
    '    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
    '返回：
    '      正常：票据领用ID>0
    '      0=失败
    '      -1:没有自用(用完或未领用)、也没有共用(未设置)
    '      -2:设置的共用已用完
    '      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)
    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    '操作员有剩余的自用票据集
    ' And 剩余数量 > 0  消费卡允许重复使用
    strSQL = _
        "Select ID, 前缀文本, 开始卡号, 终止卡号, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
        "From 消费卡领用记录" & vbNewLine & _
        "Where 接口编号 = [1] And 使用方式 = 1 And 领用人 = [2]" & vbNewLine & _
        "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, 开始卡号"
    Set rsSelf = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡类别, UserInfo.姓名)
    If lng领用ID = 0 Then
        '程序中第一次检查,且没有设置本地共用
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '也没有自用票据
        '有自用票据,按优先原则返回
        lngReturn = Val(NVL(rsSelf!id))
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        strSQL = _
            "Select ID, 使用方式, 剩余数量, 前缀文本, 开始卡号, 终止卡号" & vbNewLine & _
            "From 消费卡领用记录" & vbNewLine & _
            "Where 接口编号 = [1] And ID = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡类别, lng领用ID)
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If Val(NVL(rsTmp!使用方式)) = 2 Then '共用,要先看有没有自用
            If Not rsSelf.EOF Then
                '有自用的，优先
                lngReturn = Val(NVL(rsSelf!id))
            Else
                '没有自用取共用
                If Val(NVL(rsTmp!剩余数量)) = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                lngReturn = Val(NVL(rsTmp!id))
                blnTmp = True
            End If
        Else
            '自用票据
            If Val(NVL(rsTmp!剩余数量)) > 0 Then
                '有剩余
                lngReturn = Val(NVL(rsTmp!id))
            Else
                '其它有剩余的自用
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '其它自用也没有剩余
                lngReturn = Val(NVL(rsSelf!id))
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If Left(strBill, Len(NVL(rsTmp!前缀文本))) <> NVL(rsTmp!前缀文本) Then
                lngReturn = -3
            ElseIf Not (strBill >= NVL(rsTmp!开始卡号) And strBill <= NVL(rsTmp!终止卡号) _
                And Len(strBill) = Len(NVL(rsTmp!开始卡号))) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If Left(strBill, Len(NVL(rsSelf!前缀文本))) <> NVL(rsSelf!前缀文本) Then
                blnTmp = True
            ElseIf Not (strBill >= NVL(rsSelf!开始卡号) And strBill <= NVL(rsSelf!终止卡号) _
                And Len(strBill) = Len(NVL(rsSelf!开始卡号))) Then
                blnTmp = True
            End If
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If Left(strBill, Len(NVL(rsSelf!前缀文本))) <> NVL(rsSelf!前缀文本) Then
                        blnTmp = True
                    ElseIf Not (strBill >= NVL(rsSelf!开始卡号) And strBill <= NVL(rsSelf!终止卡号) _
                        And Len(strBill) = Len(NVL(rsSelf!开始卡号))) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!id: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Private Function Check卡号批次(Optional ByVal strCardNo As String) As Boolean
    '功能:检查严格控制卡号是否再有效领用批次内
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If Not mCardType.bln严格控制 Then Check卡号批次 = True: Exit Function
    
    mCardType.lng领用ID = CheckUsedBill(mlng卡类别, _
        IIf(mCardType.lng领用ID > 0, mCardType.lng领用ID, mCardType.lng共用批次), strCardNo)
    If mCardType.lng领用ID <= 0 Then
        Select Case mCardType.lng领用ID
            Case 0 '操作失败
            Case -1
                If strCardNo <> "" Then ShowMsgbox "你已没有自用及共用的" & mCardType.str卡名称 & "，不能发放！" & vbCrLf & _
                    "请先在本地设置共用批次或领用一批新卡! "
                Exit Function
            Case -2
                If strCardNo <> "" Then ShowMsgbox "本地共用的" & mCardType.str卡名称 & "已用完，不能发放！" & vbCrLf & _
                    "请重新设置本地共用卡批次或领用一批新卡！"
                Exit Function
            Case -3
                ShowMsgbox "该张卡片" & IIf(strCardNo = "", "", "【" & strCardNo & "】") & "不在有效范围内，请检查是否正确刷卡！"
                Exit Function
        End Select
    End If
    
    '检查是否也被报损
    strSQL = _
        "Select 1 From 消费卡使用记录" & vbNewLine & _
        "Where 接口编号 = [1] And 领用id = [2] And 性质 = 1 And 原因 = 5 And 卡号 = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng卡类别, mCardType.lng领用ID, strCardNo)
    If rsTemp.EOF = False Then
        ShowMsgbox "该张卡号" & IIf(strCardNo = "", "", "【" & strCardNo & "】") & "已被报损，不能再使用！"
        Exit Function
    End If
    
    Check卡号批次 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput卡号(ByVal blnSaveData As Boolean, _
    Optional ByRef lng卡张数 As Long, Optional ByRef strCardNos As String, _
    Optional ByRef lng卡ID As Long) As Boolean
    '功能:检查输入的卡号是否合法
    '入参:
    '   blnSaveData - 是否保存数据前的检查
    '出参:
    '   lng卡张数 - 本次发卡范围中卡数量
    '   strCardNos - 分解的卡号，多个用逗号“,”分隔
    '   lng卡ID - 回收卡时消费卡ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strTable As String, varPara() As Variant
    Dim i As Long, j As Long, k As Long
    Dim strCardNoTemp As String, strCardNo As String
    Dim strCardNoStart As String, strCardNoEnd As String, strInCardNos As String
    Dim objListItem As ListItem, strInfo As String
    Dim varData As Variant, varListData As Variant
    Dim strFindNo As String
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    lng卡张数 = 0
    If mEditType = gEd_发卡 Then
        strCardNoStart = Trim(txt(txt_开始卡号).Text)
        strCardNoEnd = Trim(txt(txt_结束卡号).Text)
        If strCardNoStart = "" And blnSaveData = False Then
            ShowMsgbox "请刷卡或输入卡号！": GoTo DataInvalid
        End If
        
        If strCardNoStart <> "" And Len(strCardNoStart) <> mCardType.lng卡号长度 Then
            ShowMsgbox "卡号长度应为" & mCardType.lng卡号长度 & "位，请检查！": GoTo DataInvalid
        End If
        
        If strCardNoEnd <> "" Then
            If Len(strCardNoEnd) <> mCardType.lng卡号长度 Then
                ShowMsgbox "卡号长度应为" & mCardType.lng卡号长度 & "位，请检查！"
                zlControl.ControlSetFocus txt(txt_结束卡号)
                Exit Function
            End If
            If strCardNoEnd <= strCardNoStart Then
                ShowMsgbox "结束卡号必须大于开始卡号，请检查！"
                zlControl.ControlSetFocus txt(txt_结束卡号)
                Exit Function
            End If
            
            If Check卡号批次(strCardNoStart) = False Then Exit Function
            If Check卡号批次(strCardNoEnd) = False Then Exit Function
            
            If SplitCardNos(strCardNoStart & "～" & strCardNoEnd, strInCardNos) = False Then Exit Function
        Else
            If strCardNoStart <> "" Then
                If Check卡号批次(strCardNoStart) = False Then Exit Function
                strInCardNos = strCardNoStart
            End If
        End If
        
        '针对每一张卡号进行检查
        varData = Split(strInCardNos, ",")
        lng卡张数 = UBound(varData) + 1
        For k = 0 To UBound(varData)
            strCardNo = varData(k)
            If FindDataInGrid(strCardNo) Then
                ShowMsgbox "卡号为：" & strCardNo & " 的" & mCardType.str卡名称 & "已存在于发卡列表中！": GoTo DataInvalid
            End If
        Next
        
        If strInCardNos <> "" Then strCardNos = strCardNos & "," & strInCardNos
        If blnSaveData Then
            If CheckCardsInGrid() = False Then Exit Function
            strTemp = GetCardsFromGrid()
            If strTemp <> "" Then strCardNos = strCardNos & "," & strTemp
        End If
        If strCardNos <> "" Then strCardNos = Mid(strCardNos, 2)
        If strCardNos = "" Then
            ShowMsgbox "请刷卡或输入卡号！": GoTo DataInvalid
        End If
        
        varPara = Array(mlng卡类别, mCardType.lng领用ID)
        If FromStringListBulidSQL(0, strCardNos, varPara, strTable, "卡号", 3) = False Then Exit Function
        strSQL = _
            "Select a.ID, a.卡类型, a.可否充值, b.卡号, a.序号," & vbNewLine & _
            "       (Select Max(序号) From 消费卡信息 B Where 卡号 = a.卡号 And 接口编号 = a.接口编号) As 最大序号," & vbNewLine & _
            "       To_Char(a.回收时间, 'yyyy-mm-dd hh24:mi:ss') As 回收时间," & vbNewLine & _
            "       To_Char(a.停用日期, 'yyyy-mm-dd hh24:mi:ss') As 停用日期" & vbNewLine & _
            "From 消费卡信息 A, (" & strTable & ") B" & vbNewLine & _
            "Where a.卡号 = b.卡号 And a.接口编号(+) = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
        Do While Not rsTemp.EOF
            If Val(NVL(rsTemp!id)) <> 0 Then
                If NVL(rsTemp!回收时间, "3000-01-01") >= "3000-01-01" Then
                    ShowMsgbox "卡号为:" & NVL(rsTemp!卡号) & " 的" & mCardType.str卡名称 & "正在使用，不能再发卡！"
                    strFindNo = NVL(rsTemp!卡号): GoTo DataInvalid
                End If
                If NVL(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                    ShowMsgbox "卡号为:" & NVL(rsTemp!卡号) & " 的" & mCardType.str卡名称 & "已经停止使用，不能再发卡！"
                    strFindNo = NVL(rsTemp!卡号): GoTo DataInvalid
                End If
            End If
            rsTemp.MoveNext
        Loop
        
        '两张以上卡片时检查是否已被报损，一张及两张的已检查
        If mCardType.lng领用ID > 0 And UBound(Split(strCardNos, ",")) > 1 Then
            strTemp = ""
            strSQL = _
                "Select Distinct a.卡号" & vbNewLine & _
                "From 消费卡使用记录 A, (" & strTable & ") B" & vbNewLine & _
                "Where a.卡号 = b.卡号 And a.接口编号 = [1] And a.领用id = [2] And a.性质 = 1 And a.原因 = 5"
            Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
            Do While Not rsTemp.EOF
                strTemp = strTemp & "," & NVL(rsTemp!卡号)
                rsTemp.MoveNext
            Loop
            If strTemp <> "" Then
                strTemp = Mid(strTemp, 2)
                ShowMsgbox "以下卡片已被报损，不能再使用：" & vbCrLf & strTemp
                Exit Function
            End If
        End If
        CheckInput卡号 = True
        Exit Function
    End If
    
    If mEditType = gEd_回收 Then
        strCardNo = Trim(txt(txt_开始卡号).Text)
        If strCardNo = "" And blnSaveData = False Then
            ShowMsgbox "请刷卡或输入卡号！": GoTo DataInvalid
        End If
        If FindDataInGrid(strCardNo) Then
            ShowMsgbox "卡号为：" & strCardNo & " 的" & mCardType.str卡名称 & "已存在于回收列表中！": GoTo DataInvalid
        End If
        
        If strCardNo <> "" Then
            lng卡张数 = 1
            strCardNos = strCardNos & "," & strCardNo
        End If
        If blnSaveData Then
            strTemp = GetCardsFromGrid()
            If strTemp <> "" Then strCardNos = strCardNos & "," & strTemp
        End If
        If strCardNos <> "" Then strCardNos = Mid(strCardNos, 2)
        If strCardNos = "" Then
            ShowMsgbox "请刷卡或输入卡号！": GoTo DataInvalid
        End If
         
        varPara = Array(mlng卡类别)
        If FromStringListBulidSQL(0, strCardNos, varPara, strTable, "卡号", 2) = False Then Exit Function
        strSQL = _
            "Select a.ID, a.卡类型, a.可否充值, b.卡号, a.序号," & vbNewLine & _
            "       To_Char(a.回收时间, 'yyyy-mm-dd hh24:mi:ss') As 回收时间," & vbNewLine & _
            "       To_Char(a.停用日期, 'yyyy-mm-dd hh24:mi:ss') As 停用日期" & vbNewLine & _
            "From 消费卡信息 A, (" & strTable & ") B" & vbNewLine & _
            "Where a.卡号(+) = b.卡号 And a.接口编号(+) = [1]" & vbNewLine & _
            "      And (a.序号 Is Null Or a.序号 = (Select Max(序号) From 消费卡信息 Where 卡号 = a.卡号 And 接口编号 = a.接口编号))"
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
        If rsTemp.EOF Then
            ShowMsgbox "未查找到卡的相关信息，可能已经被他人删除，不能回收！": GoTo DataInvalid
        Else
            Do While Not rsTemp.EOF
                strFindNo = NVL(rsTemp!卡号)
                If Val(NVL(rsTemp!id)) = 0 Then
                    '卡不存在
                    ShowMsgbox mCardType.str卡名称 & "(卡号为:" & NVL(rsTemp!卡号) & ")可能已经被他人删除，不能回收！": GoTo DataInvalid
                Else
                    If NVL(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
                        ShowMsgbox "卡号为:" & NVL(rsTemp!卡号) & " 的" & mCardType.str卡名称 & "已被回收，不能再回收！": GoTo DataInvalid
                    End If
                    '停用的也可以回收
                    'If Nvl(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                    '    ShowMsgbox "卡号为:" & Nvl(rsTemp!卡号) & " 的" & mCardType.str卡名称 & "已经停止使用，不能回收！": GoTo DataInvalid
                    'End If
                End If
                lng卡ID = Val(NVL(rsTemp!id))
                rsTemp.MoveNext
            Loop
        End If
        CheckInput卡号 = True
        Exit Function
    End If
    
    If CardIsValid(mlng卡ID) = False Then GoTo DataInvalid
    
    If mEditType = gEd_换卡 Or mEditType = gEd_补卡 Then
        If Trim(txt(txt_新卡卡号).Text) = "" Then
            ShowMsgbox "请刷卡或输入新卡卡号！"
            zlControl.ControlSetFocus txt(txt_新卡卡号)
            Exit Function
        End If
        
        If Check卡号批次(txt(txt_新卡卡号).Text) = False Then
            zlControl.ControlSetFocus txt(txt_新卡卡号)
            zlControl.TxtSelAll txt(txt_新卡卡号)
            Exit Function
        End If
        
        strSQL = _
            "Select a.ID, a.卡类型, a.可否充值, a.卡号, a.序号," & vbNewLine & _
            "       (Select Max(序号) From 消费卡信息 B Where 卡号 = a.卡号 And 接口编号 = a.接口编号) As 最大序号," & vbNewLine & _
            "       To_Char(a.回收时间, 'yyyy-mm-dd hh24:mi:ss') As 回收时间," & vbNewLine & _
            "       To_Char(a.停用日期, 'yyyy-mm-dd hh24:mi:ss') As 停用日期" & vbNewLine & _
            "From 消费卡信息 A" & vbNewLine & _
            "Where a.接口编号 = [1] And a.卡号 = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng卡类别, txt(txt_新卡卡号))
        If Not rsTemp.EOF Then
            If NVL(rsTemp!回收时间, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "卡号为:" & NVL(rsTemp!卡号) & " 的" & mCardType.str卡名称 & "正在使用，不能再发卡！"
                zlControl.ControlSetFocus txt(txt_新卡卡号)
                zlControl.TxtSelAll txt(txt_新卡卡号)
                Exit Function
            End If
            If NVL(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "卡号为:" & NVL(rsTemp!卡号) & " 的" & mCardType.str卡名称 & "已经停止使用，不能再发卡！"
                zlControl.ControlSetFocus txt(txt_新卡卡号)
                zlControl.TxtSelAll txt(txt_新卡卡号)
                Exit Function
            End If
        End If
    End If
    
    CheckInput卡号 = True
    Exit Function
DataInvalid:
    If (mEditType = gEd_发卡 Or mEditType = gEd_回收) And strFindNo <> "" Then
        If FindDataInGrid(strFindNo) Then
            zlControl.ControlSetFocus vsfCardNo
        Else
            zlControl.ControlSetFocus txt(txt_开始卡号)
            zlControl.TxtSelAll txt(txt_开始卡号)
        End If
    ElseIf mEditType = gEd_换卡 Then
        zlControl.ControlSetFocus txt(txt_原卡卡号)
        zlControl.TxtSelAll txt(txt_原卡卡号)
    ElseIf mEditType = gEd_补卡 Then
        zlControl.ControlSetFocus cbo原卡卡号
    Else
        zlControl.ControlSetFocus txt(txt_开始卡号)
        zlControl.TxtSelAll txt(txt_开始卡号)
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCardsFromGrid(Optional ByVal blnGetID As Boolean) As String
    '从表格中获取卡号
    '入参：
    '   blnGetID 是否获取消费卡ID,否则获取卡号
    Dim i As Long, j As Long
    Dim varData As Variant
    Dim strCards As String
    
    On Error GoTo ErrHandler
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            If vsfCardNo.TextMatrix(i, j) <> "" Then
                varData = vsfCardNo.Cell(flexcpData, i, j) 'Array(卡张数,分解卡号,消费卡ID)
                strCards = strCards & "," & IIf(blnGetID, varData(2), varData(1))
            End If
        Next
    Next
    If strCards <> "" Then strCards = Mid(strCards, 2)
    GetCardsFromGrid = strCards
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckCardsInGrid() As Boolean
    '检查表格中的卡号有效性
    Dim i As Long, j As Long
    Dim strCardNoStart As String, strCardNoEnd As String
    
    On Error GoTo ErrHandler
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            If vsfCardNo.TextMatrix(i, j) <> "" Then
                strCardNoStart = Split(vsfCardNo.TextMatrix(i, j) & "～", "～")(0)
                strCardNoEnd = Split(vsfCardNo.TextMatrix(i, j) & "～", "～")(1)
                
                If Check卡号批次(strCardNoStart) = False Then Exit Function
                If strCardNoEnd <> "" Then
                    If Check卡号批次(strCardNoEnd) = False Then Exit Function
                End If
            End If
        Next
    Next
    CheckCardsInGrid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCardsCount() As Long
    '从卡号张数
    Dim i As Long, j As Long
    Dim varData As Variant
    Dim lngCount As Long
    
    On Error GoTo ErrHandler
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            If vsfCardNo.TextMatrix(i, j) <> "" Then
                varData = vsfCardNo.Cell(flexcpData, i, j) 'Array(卡张数,分解卡号,消费卡ID)
                lngCount = lngCount + Val(varData(0))
            End If
        Next
    Next
    '加上未加入表格中的数量
    lngCount = lngCount + Val(txt(txt_开始卡号).Tag)
    
    GetCardsCount = lngCount
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Check缴款情况() As Boolean
    '功能:检查缴款情况
    Dim objCard As Card
    Dim strTitle As String, lng卡片张数 As Long
    Dim blnYes As Boolean
    
    On Error GoTo ErrHandler
    If mdbl实收合计 < 0 Then
        strTitle = "退款"
    Else
        strTitle = "收款"
    End If
    
    If GetCurCard(objCard) = False Then
        ShowMsgbox "当前" & strTitle & "方式未选择，请检查！"
        zlControl.ControlSetFocus cbo支付方式
        Exit Function
    End If
    
    If zlDblIsValid(Trim(txt(txt_缴款).Text), 16, True, False, txt(txt_缴款).hWnd, strTitle) = False Then Exit Function
    
    If objCard.结算性质 = 1 Then
        If Val(txt(txt_缴款).Text) = 0 And RoundEx(mdbl实收合计 - mdbl本次误差, 6) <> 0 Then
            ShowMsgbox "你还未输入" & strTitle & "金额，是否继续？", True, blnYes
            If blnYes = False Then
                zlControl.ControlSetFocus txt(txt_缴款)
                Exit Function
            End If
        End If
    Else
        If RoundEx(mdbl实收合计, 6) = 0 Then
            ShowMsgbox "当前" & strTitle & "金额为零，不能使用非现金结算方式！"
            zlControl.ControlSetFocus cbo支付方式
            Exit Function
        End If
        
        If Val(txt(txt_缴款).Text) = 0 Then
            ShowMsgbox "未输入" & strTitle & "金额，请检查！"
            zlControl.ControlSetFocus txt(txt_缴款)
            Exit Function
        End If
    End If
    If Val(txt(txt_缴款).Text) <> 0 Then
        If Val(txt(txt_缴款).Text) < RoundEx(Abs(mdbl实收合计 - mdbl本次误差), 6) Then
            ShowMsgbox strTitle & "金额(" & Format(Val(txt(txt_缴款).Text), "0.00") & ")不足本次未付金额(" & _
                FormatEx(Abs(mdbl实收合计 - mdbl本次误差), 6, , , 2) & ")，请检查！"
            zlControl.ControlSetFocus txt(txt_缴款)
            Exit Function
        End If
        
        If objCard.结算性质 <> 1 And Val(txt(txt_缴款).Text) > Val(Format(Abs(mdbl实收合计 - mdbl本次误差), "0.00")) Then
            ShowMsgbox strTitle & "金额(" & Format(Val(txt(txt_缴款).Text), "0.00") & ")大于了本次未付金额(" & _
                FormatEx(Abs(mdbl实收合计 - mdbl本次误差), 6, , , 2) & ")，请检查！"
            zlControl.ControlSetFocus txt(txt_缴款)
            Exit Function
        End If
    End If
    
    If mEditType = gEd_发卡 And mCardType.bln特定病人 = False Then
        lng卡片张数 = Val(lbl(lbl_卡张数2).Caption)
    Else
        lng卡片张数 = 1
    End If
    If CheckThreeSwapIsValied(objCard, mdbl实收合计, lng卡片张数) = False Then
        zlControl.ControlSetFocus cbo支付方式
        Exit Function
    End If
    
    If zlCommFun.StrIsValid(Trim(txt(txt_缴款人).Text), 20, txt(txt_缴款人).hWnd, "缴款人") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_开户行).Text), 50, txt(txt_开户行).hWnd, "开户行") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_帐号).Text), 20, txt(txt_帐号).hWnd, "帐号") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_结算号码).Text), 30, txt(txt_结算号码).hWnd, "结算号码") = False Then Exit Function
    Check缴款情况 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckInput() As Boolean
    '功能:检查输入项是否合法
    '返回:合法返回true,否则返回False
    On Error GoTo ErrHandler
    If zlCommFun.StrIsValid(Trim(txt(txt_发卡原因).Text), 50, txt(txt_发卡原因).hWnd, "发卡原因") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_备注).Text), 100, txt(txt_备注).hWnd, "备注") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_充值说明).Text), 100, txt(txt_充值说明).hWnd, "充值说明") = False Then Exit Function
    
    If mEditType = gEd_发卡 Then
        If CheckPassword(txt(txt_确认密码), txt(txt_密码)) = False Then Exit Function
        If zlCommFun.StrIsValid(Trim(txt(txt_密码).Text), 20, txt(txt_密码).hWnd, "密码") = False Then Exit Function
        If zlCommFun.StrIsValid(Trim(txt(txt_确认密码).Text), 20, txt(txt_确认密码).hWnd, "确认密码") = False Then Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txt(txt_领卡人).Text), 20, txt(txt_领卡人).hWnd, "领卡人") = False Then Exit Function
    If mCardType.bln特定病人 And Val(txt(txt_领卡人).Tag) = 0 Then
        ShowMsgbox "领卡人无效，请重新输入！注意，领卡人必须是建档病人。"
        zlControl.ControlSetFocus txt(txt_领卡人)
        Exit Function
    End If
    If Trim(txt(txt_领卡部门).Text) <> "" And Val(txt(txt_领卡部门).Tag) = 0 Then
        ShowMsgbox " 你输入的领卡部门有误，请检查！"
        zlControl.ControlSetFocus txt(txt_领卡部门)
        Exit Function
    End If
    If mEditType = gEd_修改 Then CheckInput = True: Exit Function
    
    '金额检查
    If CheckInput卡面额 = False Then Exit Function
    If CheckInput实际销售额 = False Then Exit Function
    If CheckInput实际充值缴款 = False Then Exit Function
    If CheckInput充值扣率 = False Then Exit Function
    If CheckInput本次充值 = False Then Exit Function
    CheckInput = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SavePayCard(ByVal lng发卡序号 As Long) As Boolean
    '功能:保存发卡信息
    '返回:保存成功,返回true,否则返回False
    Dim strCardNoStart As String, strCardNoEnd As String, strInCardNos As String
    Dim cllPro As New Collection, str发卡时间 As String
    Dim strCardNos As String, varCardNos As Variant
    Dim blnTrain As Boolean
    Dim i As Long, objCard As Card
    Dim lng结算序号 As Long, lng记录ID As Long
    Dim lng卡片张数 As Long
    
    On Error GoTo ErrHandler
    '取还未加入表格的卡号
    strCardNoStart = Trim(txt(txt_开始卡号).Text)
    strCardNoEnd = Trim(txt(txt_结束卡号).Text)
    If strCardNoEnd <> "" Then
        If SplitCardNos(strCardNoStart & "～" & strCardNoEnd, strInCardNos) = False Then Exit Function
    Else
        If strCardNoStart <> "" Then strInCardNos = strCardNoStart
    End If
    
    strCardNos = GetCardsFromGrid()
    If strInCardNos <> "" Then strCardNos = strCardNos & IIf(strCardNos = "", "", ",") & strInCardNos
    
    Set cllPro = New Collection
    str发卡时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    lng记录ID = zlDatabase.GetNextId("病人卡结算记录")
    lng结算序号 = lng记录ID '将第一个 病人卡结算记录.ID 作为结算序号
    If zlCommFun.ActualLen(strCardNos) > 4000 Then
        varCardNos = Split(strCardNos, ",")
        strCardNos = ""
        For i = 0 To UBound(varCardNos)
            If zlCommFun.ActualLen(strCardNos & "," & varCardNos(i)) > 4000 Then
                strCardNos = Mid(strCardNos, 2)
                If AddCardDataSQL(lng发卡序号, strCardNos, str发卡时间, cllPro, lng结算序号, lng记录ID) = False Then Exit Function
                lng记录ID = 0 '除第一个ID外，其它的都到过程中去取
                strCardNos = ""
            End If
            strCardNos = strCardNos & "," & varCardNos(i)
        Next
        If strCardNos <> "" Then strCardNos = Mid(strCardNos, 2)
    End If
    If strCardNos <> "" Then
        If AddCardDataSQL(lng发卡序号, strCardNos, str发卡时间, cllPro, lng结算序号, lng记录ID) = False Then Exit Function
    End If
    If cllPro.count = 0 Then
        ShowMsgbox " 你没有录入任何发卡卡号，请检查！"
        Exit Function
    End If
    
    blnTrain = True
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '三方卡结算
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.接口序号 > 0 Then
        If mEditType = gEd_发卡 And mCardType.bln特定病人 = False Then
            lng卡片张数 = Val(lbl(lbl_卡张数2).Caption)
        Else
            lng卡片张数 = 1
        End If
        If ExecuteThreeSwapPay(objCard, lng结算序号, mdbl实收合计, lng卡片张数) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SavePayCard = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AddCardDataSQL(ByVal lng发卡序号 As Long, ByVal strCardNos As String, _
    ByVal str发卡时间 As String, ByRef cllPro As Collection, ByVal lng结算序号 As Long, _
    Optional ByVal lng记录ID As Long) As Boolean
    '功能:获取插入SQL语句
    '入参:lng发卡序号-主要是标明一批发卡时的发卡序号,以便打印
    '出参:
    '返回:
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_消费卡信息_Insert
    strSQL = "Zl_消费卡信息_Insert("
    '  接口编号_In     消费卡信息.接口编号%Type,
    strSQL = strSQL & "" & mlng卡类别 & ","
    '  卡号_In         Varchar2,--卡号_In 多个用逗号,分隔
    strSQL = strSQL & "'" & strCardNos & "',"
    '  卡类型_In       消费卡信息.卡类型%Type,
    strSQL = strSQL & "'" & zlStr.NeedName(cbo卡类型.Text) & "',"
    '  密码_In         消费卡信息.密码%Type,
    strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txt(txt_密码).Text) & "',"
    '  限制类别_In     消费卡信息.限制类别%Type,
    strSQL = strSQL & "'" & Get限制类别() & "',"
    '  可否充值_In     消费卡信息.可否充值%Type,
    strSQL = strSQL & "" & IIf(chk充值.value = vbChecked, 1, 0) & ","
    '  有效期_In       消费卡信息.有效期%Type,
    If IsNull(dtp卡有效日期.value) Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "To_Date('" & Format(dtp卡有效日期.value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
    End If
    '  发卡原因_In     消费卡信息.发卡原因%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_发卡原因).Text) & "',"
    '  发卡人_In       消费卡信息.发卡人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  发卡人编号_In   病人卡结算记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  发卡时间_In     消费卡信息.发卡时间%Type,
    strSQL = strSQL & "To_Date('" & str发卡时间 & "','yyyy-mm-dd hh24:mi:ss'),"
    '  领卡人_In       消费卡信息.领卡人%Type,
    strSQL = strSQL & "'" & txt(txt_领卡人).Text & "',"
    '  病人id_In       消费卡信息.病人id%Type,
    strSQL = strSQL & "" & IIf(mCardType.bln特定病人, Val(txt(txt_领卡人).Tag), "NULL") & ","
    '  领卡部门id_In   消费卡信息.领卡部门id%Type,
    strSQL = strSQL & "" & IIf(txt(txt_领卡部门).Tag = "", "NULL", Val(txt(txt_领卡部门).Tag)) & ","
    '  备注_In         消费卡信息.备注%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_备注).Text) & "',"
    '  卡面金额_In     消费卡信息.卡面金额%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_卡面额).Text), 4) & ","
    '  销售金额_In     消费卡信息.销售金额%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_销售额).Text), 4) & ","
    '  发卡序号_In     消费卡信息.发卡序号%Type,
    strSQL = strSQL & "" & lng发卡序号 & ","
    '  领用id_In       消费卡信息.领用id%Type,
    strSQL = strSQL & "" & IIf(mCardType.lng领用ID = 0, "NULL", mCardType.lng领用ID) & ","
    '  充值折扣率_In   消费卡信息.充值折扣率%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_充值扣率).Text) * IIf(chk充值.value = vbChecked, 1, 0), 4) & ","
    '  充值金额_In     病人卡结算记录.应收金额%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_本次充值).Text) * IIf(chk充值.value = vbChecked, 1, 0), 4) & ","
    '  充值缴款金额_In 病人卡结算记录.实收金额%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_充值缴款).Text) * IIf(chk充值.value = vbChecked, 1, 0), 4) & ","
    '  充值说明_In     病人卡结算记录.备注%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_充值说明).Text) & "',"
    '  缴款人_In       病人卡结算记录.缴款人姓名%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_缴款人).Text) & "',"
    '  结算方式_In     病人卡结算记录.结算方式%Type,
    strSQL = strSQL & "'" & mCurCardPay.str结算方式 & "',"
    '  结算序号_In     病人卡结算记录.结算序号%Type,
    strSQL = strSQL & "" & lng结算序号 & ","
    '  记录id_In       病人卡结算记录.Id%Type := Null,
    strSQL = strSQL & "" & IIf(lng记录ID = 0, "NULL", lng记录ID) & ","
    '  结算号码_In     病人卡结算记录.结算号码%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_结算号码).Text) & "',"
    '  开户行_In       病人卡结算记录.单位开户行%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_开户行).Text) & "',"
    '  帐号_In         病人卡结算记录.单位帐号%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_帐号).Text) & "',"
    '  卡类别id_In   病人卡结算记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", mCurCardPay.lng卡类别ID) & ","
    '  结算卡号_In   病人卡结算记录.结算卡号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人卡结算记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易流水号 & "'") & ","
    '  交易说明_In   病人卡结算记录.交易说明%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易说明 & "'") & ","
    '  缴款_In         病人卡结算记录.缴款%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_缴款).Text), 4), "NULL") & ","
    '  找补_In         病人卡结算记录.找补%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_找补).Tag), 4), "NULL") & ")"
    AddArray cllPro, strSQL
    AddCardDataSQL = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveModifyCard() As Boolean
    '功能:保存卡片修改信息
    '返回:修改成功,返回True,否则返回False
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_消费卡信息_Update
    strSQL = "Zl_消费卡信息_Update("
    '  Id_In         消费卡信息.Id%Type,
    strSQL = strSQL & "" & mlng卡ID & ","
    '  卡类型_In     消费卡信息.卡类型%Type,
    strSQL = strSQL & "'" & zlStr.NeedName(cbo卡类型.Text) & "',"
    '  可否充值_In   消费卡信息.可否充值%Type,
    strSQL = strSQL & "" & IIf(chk充值.value = vbChecked, 1, 0) & ","
    '  有效期_In     消费卡信息.有效期%Type,
    If IsNull(dtp卡有效日期.value) Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "To_Date('" & Format(dtp卡有效日期.value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
    End If
    '  发卡原因_In   消费卡信息.发卡原因%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_发卡原因).Text) & "',"
    '  领卡人_In     消费卡信息.领卡人%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_领卡人).Text) & "',"
    '  病人id_In     消费卡信息.病人id%Type,
    strSQL = strSQL & "" & IIf(mCardType.bln特定病人, Val(txt(txt_领卡人).Tag), "NULL") & ","
    '  领卡部门id_In 消费卡信息.领卡部门id%Type,
    strSQL = strSQL & "" & IIf(Val(txt(txt_领卡部门).Tag) = 0, "NULL", Val(txt(txt_领卡部门).Tag)) & ","
    '  备注_In       消费卡信息.备注%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_备注).Text) & "',"
    '  限制类别_In     消费卡信息.限制类别%Type
    strSQL = strSQL & "'" & Get限制类别() & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    SaveModifyCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCallBack(Optional blnCancelCallBack As Boolean = False) As Boolean
    '功能:回收处理
    '入参:
    '   blnCancelCallBack-是否取消回收
    Dim cllPro As New Collection, strIDs As String, varIDs As Variant
    Dim strSQL As String, blnTrain As Boolean
    Dim strNow As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If blnCancelCallBack Then
        'Zl_消费卡信息_Callback
        strSQL = "Zl_消费卡信息_Callback("
        '  Ids_In       varchar2,
        strSQL = strSQL & "" & mlng卡ID & ","
        '  回收人_In   消费卡信息.回收人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  回收时间_In 消费卡信息.回收时间%Type,
        strSQL = strSQL & "to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),"
        '  取消回收_In Number := 0
        strSQL = strSQL & "" & 1 & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        SaveCallBack = True: Exit Function
    End If
    
    '可能存在批量回收操作,因此要传入ID的集合
    strIDs = GetCardsFromGrid(True)
    If Trim(txt(txt_开始卡号).Text) <> "" Then
        '取还未加入表格的卡号
        strIDs = strIDs & IIf(strIDs = "", "", ",") & mlng卡ID
    End If
    
    Set cllPro = New Collection
    If zlCommFun.ActualLen(strIDs) > 4000 Then
        varIDs = Split(strIDs, ",")
        strIDs = ""
        For i = 0 To UBound(varIDs)
            If zlCommFun.ActualLen(strIDs & "," & varIDs(i)) > 4000 Then
                strIDs = Mid(strIDs, 2)
                'Zl_消费卡信息_Callback
                strSQL = "Zl_消费卡信息_Callback("
                '  Ids_In     varchar2,
                strSQL = strSQL & "'" & strIDs & "',"
                '  回收人_In   消费卡信息.回收人%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  回收时间_In 消费卡信息.回收时间%Type,
                strSQL = strSQL & "to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),"
                '  取消回收_In Number := 0
                strSQL = strSQL & "" & 0 & ")"
                AddArray cllPro, strSQL
                strIDs = ""
            End If
            strIDs = strIDs & "," & varIDs(i)
        Next
        If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    End If
    If strIDs <> "" Then
        'Zl_消费卡信息_Callback
        strSQL = "Zl_消费卡信息_Callback("
        '  Ids_In       varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '  回收人_In   消费卡信息.回收人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  回收时间_In 消费卡信息.回收时间%Type,
        strSQL = strSQL & "to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),"
        '  取消回收_In Number := 0
        strSQL = strSQL & "" & 0 & ")"
        AddArray cllPro, strSQL
    End If
    If cllPro.count = 0 Then
        ShowMsgbox " 你没有刷要回收的" & mCardType.str卡名称 & "，请检查！"
        Exit Function
    End If
    
    blnTrain = True
    ExecuteProcedureArrAy cllPro, Me.Caption
    blnTrain = False
    SaveCallBack = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveBackCard(Optional blnCancelBackCard As Boolean) As Boolean
    '功能：退卡处理
    '入参:
    '   blnCancelBackCard-是否取消退卡
    Dim strSQL As String, blnTrain As Boolean
    Dim lng结算序号 As Long, objCard As Card
    
    Err = 0: On Error GoTo ErrHandler
    lng结算序号 = zlDatabase.GetNextId("病人卡结算记录")
    'Zl_消费卡信息_Backcard
    strSQL = "Zl_消费卡信息_Backcard("
    '  消费卡Id_In         消费卡信息.Id%Type,
    strSQL = strSQL & "" & mlng卡ID & ","
    '  操作员编号_In 病人卡结算记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  退卡时间_In   消费卡信息.回收时间%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  结算方式_In     病人卡结算记录.结算方式%Type,
    strSQL = strSQL & "'" & mCurCardPay.str结算方式 & "',"
    '  结算序号_In   病人卡结算记录.Id%Type,
    strSQL = strSQL & "" & lng结算序号 & ","
    '  误差金额_In   病人卡结算记录.实收金额%Type,
    strSQL = strSQL & "" & IIf(blnCancelBackCard, 1, -1) * mdbl本次误差 & ","
    '  取消退卡_In   Number := 0,
    strSQL = strSQL & "" & IIf(blnCancelBackCard, 1, 0) & ","
    '  结算号码_In     病人卡结算记录.结算号码%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_结算号码).Text) & "',"
    '  开户行_In       病人卡结算记录.单位开户行%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_开户行).Text) & "',"
    '  帐号_In         病人卡结算记录.单位帐号%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_帐号).Text) & "',"
    '  卡类别id_In   病人卡结算记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", mCurCardPay.lng卡类别ID) & ","
    '  结算卡号_In   病人卡结算记录.结算卡号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人卡结算记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易流水号 & "'") & ","
    '  交易说明_In   病人卡结算记录.交易说明%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易说明 & "'") & ","
    '  缴款_In         病人卡结算记录.缴款%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_缴款).Text), 4), "NULL") & ","
    '  找补_In         病人卡结算记录.找补%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt结算性质 = 1, _
        IIf(mdbl实收合计 < 0, -1, 1) * Round(Val(txt(txt_找补).Tag), 4), "NULL") & ")"
    
    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '三方卡结算
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.接口序号 > 0 Then
        If ExecuteThreeSwapPay(objCard, lng结算序号, mdbl实收合计) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveBackCard = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCardNo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim varData As Variant, lng卡ID As Long
    
    On Error GoTo ErrHandler
    cmdDelete.Enabled = ZL_vsGrid_CurrCellHaveData(vsfCardNo, NewRow, NewCol)
    If mEditType = gEd_回收 And NewRow >= 0 And NewCol >= 0 Then
        If vsfCardNo.TextMatrix(NewRow, NewCol) <> "" Then
            varData = vsfCardNo.Cell(flexcpData, NewRow, NewCol)
            lng卡ID = varData(2) 'Array(卡张数,分解卡号,消费卡ID)
            If CollExitsValue(mcllCard, "K" & lng卡ID) Then
                Call ShowCardInfo(mcllCard("K" & lng卡ID), True)
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FindDataInGrid(ByVal strCardNo As String, Optional ByVal blnSetFocus As Boolean = True) As Boolean
    '在单元格查找数据，并定位
    '入参：
    '   strCardNo - 卡号
    '   blnSetFocus - 是否设置选择单元格
    Dim i As Long, j As Long
    Dim varData As Variant, strTemp As String
    
    On Error GoTo ErrHandler
    If strCardNo = "" Then Exit Function
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            strTemp = vsfCardNo.TextMatrix(i, j)
            If strTemp <> "" Then
                strTemp = strTemp & "～" & strTemp '将不是卡号范围的转换为卡号范围
                varData = Split(strTemp, "～")
                If varData(0) <= strCardNo And strCardNo <= varData(1) Then
                    If blnSetFocus Then
                        vsfCardNo.Row = i
                        vsfCardNo.Col = j
                    End If
                    FindDataInGrid = True
                    Exit Function
                End If
            End If
        Next
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCardNo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    If Trim(vsfCardNo.TextMatrix(NewRow, NewCol)) = "" Then Cancel = True
End Sub
 

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, ByVal dblMoney As Double, _
    Optional ByVal lng卡片张数 As Long = 1) As Boolean
    '功能:三方卡刷卡验证
    '入参:objCard-当前卡
    '返回:刷卡成功,返回true,否则返回False
    Dim strXMLExpend As String, strBalanceIDs As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If objCard.接口序号 <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If dblMoney = 0 Then CheckThreeSwapIsValied = True: Exit Function
    
    If dblMoney > 0 Then
        'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:根据指定支付类别,弹出刷卡窗口
        '入参:rsClassMoney:收费类别,金额
        '        lngCardTypeID-为零时,为老一卡通刷卡
        '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
        '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
        '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
        '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
        '       bln预交-是否转预交
        '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
        '       strXmlIn-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
        '出参:str限制类别-限制类别(消费卡返回)
        '        lng卡ID-消费卡信息.ID(消费卡返回)
        '       strCardNO-返回刷卡的卡号
        '       strPassWord-返回刷卡所对应的密码
        '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
        '返回:成功,返回true,否则返回False
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
            objCard.接口序号, False, "", "", "", dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
            False, False, False, True, Nothing, False, True, "<IN><CZLX>0</CZLX></IN>") = False Then Exit Function
        
        '保存前,一些数据检查
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNos As String, _
        Optional ByVal strXMLExpend As String) As Boolean
        '功能:帐户扣款交易检查
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       strCardTypeID-卡类别ID
        '       strCardNo-卡号
        '       dblMoney-支付金额(退款时为负数)
        '       strNos-本次支付所涉及的单据
        '       strXMLExpend-(XML串:验证密码:自助机用)
        '        消费卡收款时，传入XML串：
        '        <IN>
        '            <MZXSJE>面值销售金额</MZXSJE>
        '            <CZJKJE>充值缴款金额</CZJKJE>
        '        </IN>
        '出参:
        '   strXMLExpend-(XML串:错误信息)
        '返回:扣款合法,返回true,否则返回Flase
        strXMLExpend = ""
        strXMLExpend = strXMLExpend & "<IN>"
            strXMLExpend = strXMLExpend & "<MZXSJE>" & IIf(mEditType = gEd_充值, "0", lng卡片张数 * Val(txt(txt_销售额).Text)) & "</MZXSJE>"
            strXMLExpend = strXMLExpend & "<CZJKJE>" & lng卡片张数 * Val(txt(txt_充值缴款).Text) & "</CZJKJE>"
        strXMLExpend = strXMLExpend & "</IN>"
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, objCard.接口序号, _
            False, mCurCardPay.str刷卡卡号, dblMoney, "", strXMLExpend) = False Then Exit Function
    Else
        mrsBalance.Filter = ""
        If mrsBalance.EOF Then
            ShowMsgbox "未查找到" & objCard.名称 & "原支付结算信息，不能退回！"
            Exit Function
        End If
        mCurCardPay.lng原结算序号 = Val(NVL(mrsBalance!结算序号))
        
        If objCard.是否转帐及代扣 Then
            '   zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                objCard.接口序号, False, "", "", "", dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
                False, False, False, True, Nothing, False, True, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
            
            '调用转帐接口
            'zlTransferAccountsCheck 转帐检查接口
            '参数名  参数类型    入/出   备注
            'frmMain Object  In  调用的主窗体
            'lngModule   Long    In  HIS调用模块号
            'lngCardTypeID   Long    In  卡类别ID
            'strCardNo   String  In  卡号
            'dblMoney    Double  In  转帐金额(代扣时为负数)
            'strBalanceID    String  In  原支付结算序号,费用补充记录.结算序号或病人预交记录.结算序号
            'strXMLExpend String In   XML串:
            '                            <IN>
            '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；
            '                                       2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
            '                            </IN>
            '                    Out  XML串:
            '                            <OUT>
            '                               <ERRMSG>错误信息</ERRMSG >
            '                            </OUT>
            '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
            '说明:
            '１. 在三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
            '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
            '构造XML串
            strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
            If gobjSquare.objSquareCard.zlTransferAccountsCheck(Me, mlngModule, objCard.接口序号, _
                mCurCardPay.str刷卡卡号, -1 * dblMoney, mCurCardPay.lng原结算序号, strXMLExpend) = False Then
                Call ShowThreeSwapErrMsg(0, strXMLExpend)
                Exit Function
            End If
        Else
            mrsBalance.Filter = "卡类别ID=" & objCard.接口序号
            If mrsBalance.EOF Then
                ShowMsgbox "未查找到" & objCard.名称 & "原支付结算信息，不能退回！"
                Exit Function
            End If
            mCurCardPay.lng原结算序号 = Val(NVL(mrsBalance!结算序号))
            
            If objCard.是否全退 Then
                strSQL = "Select Nvl(Sum(实收金额), 0) As 缴款合计 From 病人卡结算记录 Where 结算序号 = [1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(mrsBalance!结算序号)))
                If Val(NVL(rsTemp!缴款合计)) <> -1 * dblMoney Then
                    ShowMsgbox objCard.名称 & "不支持部分退，因此不能退回，请选择其它退款方式！" & _
                        "(原支付金额：" & FormatEx(Val(NVL(rsTemp!缴款合计)), 6, , , 2) & _
                        "，现退款金额：" & FormatEx(-1 * dblMoney, 6, , , 2) & ")"
                    Exit Function
                End If
            End If
            'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
                ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
                ByVal strBalanceIDs As String, _
                ByVal dblMoney As Double, ByVal strSwapNo As String, _
                ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '功能:帐户回退交易前的检查
                '入参:frmMain-调用的主窗体
                '       lngModule-调用的模块号
                '       lngCardTypeID-卡类别ID
                '       strCardNo-卡号
                '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
                '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡收款；
                '                                           类型=7时，ID为病人卡结算记录.结算序号
                '       dblMoney-退款金额
                '       strSwapNo-交易流水号(退款时检查)
                '       strSwapMemo-交易说明(退款时传入)
                '       strXMLExpend    XML IN  可选参数(扩展用):
                '        <TFDATA> //退费数据
                '          <YCTF>1</YCTF> //是否异常重退:1-异常重退;0-退费 此节点可能没有
                '          <TFLIST> //退费列表
                '            <NO></NO> // 退费单据
                '            <TFITEM> //退费项
                '              <SerialNum></SerialNum> //序号
                '              …
                '            </TFITEM>
                '          </TFLIST>
                '          ....
                '        </TFDATA >
                '返回:退款合法,返回true,否则返回Flase
            mCurCardPay.str刷卡卡号 = NVL(mrsBalance!结算卡号)
            mCurCardPay.str交易流水号 = NVL(mrsBalance!交易流水号)
            mCurCardPay.str交易说明 = NVL(mrsBalance!交易说明)
            strBalanceIDs = "7|" & mCurCardPay.lng原结算序号
            If gobjSquare.objSquareCard.zlReturncheck(Me, mlngModule, objCard.接口序号, _
                objCard.消费卡, mCurCardPay.str刷卡卡号, strBalanceIDs, -1 * dblMoney, _
                mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strXMLExpend) = False Then Exit Function
        
            If objCard.是否退款验卡 Then
               '弹出刷卡界面
                'zlBrushCard(frmMain As Object, _
                ByVal lngModule As Long, _
                ByVal rsClassMoney As ADODB.Recordset, _
                ByVal lngCardTypeID As Long, _
                ByVal bln消费卡 As Boolean, _
                ByVal strPatiName As String, ByVal strSex As String, _
                ByVal strOld As String, ByVal dbl金额 As Double, _
                Optional ByRef strCardNo As String, _
                Optional ByRef strPassWord As String, _
                Optional ByRef bln退费 As Boolean = False, _
                Optional ByRef blnShowPatiInfor As Boolean = False, _
                Optional ByRef bln退现 As Boolean = False, _
                Optional ByVal bln余额不足禁止 As Boolean = True, _
                Optional ByRef varSquareBalance As Variant, _
                Optional ByVal bln转预交 As Boolean = False, _
                Optional ByVal blnAllPay As Boolean = False, _
                Optional ByVal strXmlIn As String = "") As Boolean
                '       strXmlIn-三方卡调用XML入参,目前格式如下:
                '       <IN>
                '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
                '       </IN>
                If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                    objCard.接口序号, False, "", "", "", -1 * dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
                    True, False, False, True, Nothing, False, True, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
            End If
        End If
    End If
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapPay(ByVal objCard As Card, ByVal lng结算序号 As Long, _
    ByVal dblMoney As Double, Optional ByVal lng卡片张数 As Long = 1) As Boolean
    '功能:一卡通支付(三方接口)
    '入参:
    '   objCard-当前卡
    '   dblMoney-本次支付金额
    '出参:
    '返回:执行成功,返回true,否则返回False
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSwapExtendInfor As String, strTemp As String
    Dim strXMLExpend As String
    
    On Error GoTo ErrHandler
    If objCard.接口序号 <= 0 Then ExecuteThreeSwapPay = True: Exit Function
    If dblMoney = 0 Then ExecuteThreeSwapPay = True: Exit Function
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    If dblMoney > 0 Then
        'zlPaymentMoney(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        byval  strPrepayNos as string , _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, _
        ByRef strSwapMemo As String, _
        Optional ByRef strSwapExtendInfor As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户扣款交易
        '入参:frmMain-调用的主窗体
        '        lngModule-调用模块号
        '        strBalanceIDs-结帐ID,多个用逗号分离;消费卡收款时为病人卡结算记录.结算序号
        '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
        '       strCardNo-卡号
        '       dblMoney-支付金额
        '       strSwapExtendInfor- 消费卡收款时，传入XML串：
        '                            <IN>
        '                                <MZXSJE>面值销售金额</MZXSJE>
        '                                <CZJKJE>充值缴款金额</CZJKJE>
        '                            </IN>
        '出参:strSwapGlideNO-交易流水号
        '       strSwapMemo-交易说明
        '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
        '返回:扣款成功,返回true,否则返回Flase
        '说明:
        '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
        '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
        '---------------------------------------------------------------------------------------------------------------------------------------------
        strSwapExtendInfor = ""
        strSwapExtendInfor = strSwapExtendInfor & "<IN>"
            strSwapExtendInfor = strSwapExtendInfor & "<MZXSJE>" & IIf(mEditType = gEd_充值, "0", lng卡片张数 * Val(txt(txt_销售额).Text)) & "</MZXSJE>"
            strSwapExtendInfor = strSwapExtendInfor & "<CZJKJE>" & lng卡片张数 * Val(txt(txt_充值缴款).Text) & "</CZJKJE>"
        strSwapExtendInfor = strSwapExtendInfor & "</IN>"
        strTemp = strSwapExtendInfor
        If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, _
            mCurCardPay.str刷卡卡号, lng结算序号, "", dblMoney, _
            mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strSwapExtendInfor) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
        
        Call zlAddUpdateSwapSQL(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
            mCurCardPay.str刷卡卡号, mCurCardPay.str交易流水号, mCurCardPay.str交易说明, cllUpdate, 1, 0, 1)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap, 0, 1)
            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        End If
    Else
        If objCard.是否转帐及代扣 Then
            'zlTransferAccountsMoney
            '参数名  参数类型    入/出   备注
            'frmMain Object  In  调用的主窗体
            'lngModule   Long    In  HIS调用模块号
            'lngCardTypeID   Long    In  卡类别ID
            'strCardNo   String  In  卡号
            'strBalanceID    String  In  结算ID 本次支付结算序号,费用补充记录.结算序号或病人预交记录.结算序号或病人卡结算记录.结算序号
            'dblMoney    Double  In  转帐金额
            'strSwapGlideNO  String  Out 交易流水号
            'strSwapMemo String  Out 交易说明
            'strSwapExtendInfor  String  In 退费业务时，传入本次退费的冲销ID:
            '                               格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                               收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡管理收款(ID为结算序号)
            '                           Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
            'strXMLExpend String In   XML串:
            '                            <IN>
            '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；
            '                                       2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
            '                            </IN>
            '                    Out  XML串:
            '                            <OUT>
            '                               <ERRMSG>错误信息</ERRMSG >
            '                            </OUT>
            '    Boolean 函数返回    True:调用成功,False:调用失败
            '说明:
            '１. 在医保补充结算时进行的三方转帐时调用。
            '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
            '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
            '构造XML串
            strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
            strSwapExtendInfor = "7|" & mCurCardPay.lng原结算序号: strTemp = strSwapExtendInfor
            If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.接口序号, _
                mCurCardPay.str刷卡卡号, lng结算序号, -1 * dblMoney, _
                mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strSwapExtendInfor, strXMLExpend) = False Then
                gcnOracle.RollbackTrans: Call ShowThreeSwapErrMsg(1, strXMLExpend)
                Exit Function
            End If
            gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
            
            Call zlAddUpdateSwapSQL(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                mCurCardPay.str刷卡卡号, mCurCardPay.str交易流水号, mCurCardPay.str交易说明, cllUpdate, 1, 0, 1)
            zlExecuteProcedureArrAy cllUpdate, Me.Caption
            If strTemp <> strSwapExtendInfor Then
                Call zlAddThreeSwapSQLToCollection(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                    mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap, 0, 1)
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
            End If
        Else
            'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
                ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
                ByVal dblMoney As Double, _
                ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
                ByRef strSwapExtendInfor As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户扣款回退交易
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID:医疗卡类别.ID
            '       strCardNo-卡号
            '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
            '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡收款
            '       dblMoney-退款金额
            '       strSwapNo-交易流水号(扣款时的交易流水号)
            '       strSwapMemo-交易说明(扣款时的交易说明)
            '       strSwapExtendInfor-出入，本次退费的冲销ID：
            '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡收款
            '       strSwapExtendInfor-传出，交易的扩展信息
            '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
            strSwapExtendInfor = "7|" & lng结算序号
            If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, _
                mCurCardPay.str刷卡卡号, "7|" & mCurCardPay.lng原结算序号, -1 * dblMoney, _
                mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strSwapExtendInfor) = False Then
                gcnOracle.RollbackTrans: Exit Function
            End If
            gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
            
            Call zlAddUpdateSwapSQL(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                mCurCardPay.str刷卡卡号, mCurCardPay.str交易流水号, mCurCardPay.str交易说明, cllUpdate, 1, 0, 1)
            zlExecuteProcedureArrAy cllUpdate, Me.Caption
            If strTemp <> strSwapExtendInfor Then
                Call zlAddThreeSwapSQLToCollection(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                    mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap, 0, 1)
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
            End If
        End If
    End If
    ExecuteThreeSwapPay = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowThreeSwapErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '功能:三方转账检查与代扣业务出错提示
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "交易检查失败！"
        Else
            strValue = vbCrLf & "交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckThreeBalanceToCash(ByVal objCard As Card) As Boolean
    '三方卡退现检查
    Dim str操作员 As String
    
    On Error GoTo errHandle
    If Not (objCard.接口序号 > 0 And Not objCard.消费卡) Then CheckThreeBalanceToCash = True: Exit Function
    If objCard.是否退现 Then CheckThreeBalanceToCash = True: Exit Function
    
    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
        If MsgBox(objCard.名称 & "不支持退现，你确定要将其强制退现吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        str操作员 = zlDatabase.UserIdentifyByUser(Me, objCard.名称 & "强制退现，权限验证：", _
            glngSys, mlngModule, "三方退款强制退现", , True)
        If str操作员 = "" Then Exit Function
    End If
    CheckThreeBalanceToCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
