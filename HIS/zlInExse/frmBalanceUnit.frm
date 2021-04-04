VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceUnit 
   Caption         =   "合约单位病人结帐"
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   375
   ClientWidth     =   11760
   Icon            =   "frmBalanceUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11760
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picRight 
      Height          =   7095
      Left            =   4080
      ScaleHeight     =   7035
      ScaleWidth      =   7575
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   75
      Width           =   7635
      Begin VB.PictureBox picBalance 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4410
         Left            =   0
         ScaleHeight     =   4410
         ScaleWidth      =   7545
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2640
         Width           =   7545
         Begin VB.Frame fraSplit 
            Height          =   75
            Left            =   -45
            TabIndex        =   36
            Top             =   3960
            Visible         =   0   'False
            Width           =   7500
         End
         Begin VB.PictureBox picOwerFee 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   1245
            Left            =   105
            ScaleHeight     =   1215
            ScaleWidth      =   2835
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1455
            Width           =   2865
            Begin VB.Label lbl自付合计 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "123456789.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   510
               Left            =   15
               TabIndex        =   35
               Top             =   495
               Width           =   2760
            End
            Begin XtremeSuiteControls.ShortcutCaption stcTittleTotal 
               Height          =   420
               Left            =   15
               TabIndex        =   34
               Top             =   30
               Width           =   3330
               _Version        =   589884
               _ExtentX        =   5874
               _ExtentY        =   741
               _StockProps     =   6
               Caption         =   "本次结帐合计"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
            End
         End
         Begin VB.PictureBox picNotPayment 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   1245
            Left            =   105
            ScaleHeight     =   1215
            ScaleWidth      =   2835
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   90
            Width           =   2865
            Begin VB.Label lbl剩余自付 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   510
               Left            =   1965
               TabIndex        =   32
               Top             =   525
               Width           =   840
            End
            Begin XtremeSuiteControls.ShortcutCaption stcTittile 
               Height          =   450
               Left            =   15
               TabIndex        =   31
               Top             =   15
               Width           =   3315
               _Version        =   589884
               _ExtentX        =   5847
               _ExtentY        =   794
               _StockProps     =   6
               Caption         =   "当前未付"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "确定(&O)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   4950
            TabIndex        =   14
            Top             =   2205
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "取消(&C)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6240
            TabIndex        =   15
            Top             =   2205
            Width           =   1215
         End
         Begin VB.PictureBox picCurBalance 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2040
            Left            =   3270
            ScaleHeight     =   2040
            ScaleWidth      =   4005
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   105
            Width           =   4005
            Begin VB.TextBox txtBalance 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   1170
               TabIndex        =   13
               Top             =   1470
               Width           =   2805
            End
            Begin VB.TextBox txtBalance 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1170
               TabIndex        =   11
               Top             =   1020
               Width           =   2805
            End
            Begin VB.TextBox txtBalance 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1170
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   585
               Width           =   2805
            End
            Begin VB.TextBox txtBalance 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   2265
               TabIndex        =   7
               Top             =   105
               Width           =   1710
            End
            Begin zlIDKind.IDKindNew IDKindPaymentsType 
               Height          =   375
               Left            =   1170
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   105
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   661
               ShowSortName    =   0   'False
               Appearance      =   2
               IDKindStr       =   "现|现金|0|0|0|0|0|0;支|支票|0|0|0|0|0|"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontSize        =   12
               FontName        =   "宋体"
               IDKind          =   -1
               DefaultCardType =   "0"
               AllowAutoCommCard=   0   'False
               BackColor       =   -2147483633
            End
            Begin VB.Label lblBalance 
               AutoSize        =   -1  'True
               Caption         =   "找    补"
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
               Index           =   1
               Left            =   75
               TabIndex        =   8
               Top             =   645
               Width           =   1050
            End
            Begin VB.Label lblBalance 
               AutoSize        =   -1  'True
               Caption         =   "摘    要"
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
               Left            =   75
               TabIndex        =   12
               Top             =   1545
               Width           =   960
            End
            Begin VB.Label lblBalance 
               AutoSize        =   -1  'True
               Caption         =   "结算号码"
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
               Left            =   75
               TabIndex        =   10
               Top             =   1110
               Width           =   960
            End
            Begin VB.Label lblBalance 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "缴    款"
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
               Index           =   0
               Left            =   75
               TabIndex        =   29
               Top             =   150
               Width           =   1050
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBlance 
            Height          =   1410
            Left            =   105
            TabIndex        =   16
            Top             =   2835
            Width           =   7365
            _cx             =   12991
            _cy             =   2487
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmBalanceUnit.frx":15162
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
            Begin VB.Image imgDel 
               Height          =   240
               Left            =   75
               Picture         =   "frmBalanceUnit.frx":15270
               Top             =   45
               Visible         =   0   'False
               Width           =   240
            End
         End
         Begin VB.Label lbl误差额 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "误差:0.0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   3825
            TabIndex        =   37
            Top             =   2325
            Visible         =   0   'False
            Width           =   1050
         End
      End
      Begin VB.Frame fraLeft 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         TabIndex        =   20
         Top             =   -15
         Width           =   7575
         Begin VB.CommandButton cmd病人 
            Height          =   360
            Left            =   7095
            Picture         =   "frmBalanceUnit.frx":157FA
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "查找(F3)"
            Top             =   1650
            Width           =   375
         End
         Begin VB.Frame fraLine1 
            Height          =   24
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   7455
         End
         Begin VB.ComboBox cboNO 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   6045
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   960
            Width           =   1425
         End
         Begin VB.TextBox txtInvoice 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   960
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   960
            Width           =   1425
         End
         Begin VB.TextBox txtUnit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            MaxLength       =   20
            TabIndex        =   21
            Top             =   1680
            Width           =   6135
         End
         Begin VB.TextBox txt累计金额 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2145
            Width           =   2250
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   5280
            TabIndex        =   2
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label lblFact 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票据号"
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
            Left            =   160
            TabIndex        =   0
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位(&D)"
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
            Left            =   40
            TabIndex        =   26
            Top             =   1740
            Width           =   840
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "合约单位病人结帐"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   24
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   960
            TabIndex        =   25
            Top             =   240
            Width           =   3960
         End
         Begin VB.Label lblFlag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "废"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   465
            Left            =   6975
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl累计 
            AutoSize        =   -1  'True
            Caption         =   "累计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   375
            TabIndex        =   4
            Top             =   2235
            Width           =   510
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   7215
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2884
            MinWidth        =   882
            Picture         =   "frmBalanceUnit.frx":15944
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14711
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   7095
      Left            =   60
      TabIndex        =   18
      Top             =   75
      Width           =   3990
      _cx             =   7038
      _cy             =   12515
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceUnit.frx":161D8
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
Attribute VB_Name = "frmBalanceUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

'入口参数：
Private mbytInState As Byte          '0=结帐状态(默认新增),1=浏览状态
Private mblnViewCancel As Boolean    '是否查看已作废单据
Private mlng结帐ID As Long           '要浏览的单据号结帐ID,当mbytInState=1时有效
Private mblnNOMoved As Boolean       '操作的单据是否在后备数据表中
'------------------------------------------------------------------------------
Private mrsPatients As ADODB.Recordset    '本次结帐的病人ID未结费用记录集
Private mstrDec As String       '本次结帐的费用小数位数
Private mintDefault As Integer  '缺省的结算方式行号
Private mlng领用ID As Long
Private mintError As Integer    '误差费的结算方式行号
Private mintSucces As Integer
Private mobjInvoice As clsInvoice
Private mobjFact As clsFactProperty
Private mlngModul As Long
Private mstrPrivs As String
Private mrsBalance As ADODB.Recordset
Private mstrErrorBalance As String '误差结算方式
Private mblnNotChange As Boolean
Private mblnPrintBill As Boolean '是否打印票据
Private mstr退支票 As String
Private mobjICCard As Object
Private mstr缺省结算方式 As String
Private mrs结算方式 As ADODB.Recordset
Private mblnUnload As Boolean
Private Enum mInput_Idx
    Idx_缴款 = 0
    Idx_找补 = 1
    Idx_结算号码 = 2
    Idx_摘要 = 3
End Enum

'当前结帐数据
Private Type TY_Balance_Infor
    dbl当前结帐 As Double
    dbl已付合计 As Double
    dbl未付合计 As Double
    
    blnSaveBill As Boolean '当前已经保存结帐单
    strNO As String   '当前保存的结帐单
    lng结帐ID As Long '当前保存的结帐ID
    dtBalanceDate As Date '当前结帐时间
    dbl缴款 As Double
    dbl找补 As Double
    dbl退支票 As Double
    dbl误差额 As Double
    dbl现金 As Double
    lng冲销ID As Long
End Type
Private mtyBalanceInfor As TY_Balance_Infor
Private mcllCurSquareBalance As Collection '当前消费卡消费集
'当前刷卡信息
Private Type TY_BrushCard    '刷卡类型
    str卡号 As String
    str密码 As String
    str交易流水号 As String    '交易流水号
    str交易说明  As String     '交易信息
    str扩展信息 As String    '交易的扩展信息
    dbl帐户余额 As Double
    str结算号码 As String
    str结算摘要 As String
    bln转帐 As Boolean '是否当前为转帐交易
End Type
Private Enum mInput_LblIdx
    Idx_lbl缴款 = 0
    Idx_lbl找补 = 1
End Enum

'3.3 模块参数定义
Private Type Ty_ModulePara
    byt缴款输入控制 As Byte  '
End Type
Private mty_ModulePara As Ty_ModulePara
'-----------------------------------------------------------------
'3.4老版一卡通相关
Private Type TY_OneCard
      blnOneCard As Boolean      '是否启用了一卡通接口
      rsOneCard As ADODB.Recordset
      strOneCard As String       '读卡时所选择的一卡通接口对应的结算方式
End Type
Private mOldOneCard As TY_OneCard

'Private Enum BALANCECOL
'    C0姓名 = 0
'    C1性别 = 1
'    C2年龄 = 2
'    C3结帐金额 = 3
'End Enum
'Private Enum PAYCOL
'    C0方式 = 0
'    C1金额 = 1
'    C2号码 = 2
'    C3备注 = 3
'End Enum
Private Const CASHPAY = 1

Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mty_ModulePara
        '问题:43153:0-不进行控制;1-存在收取现金时,必须输入缴款;2-结帐时按单病人累计
        .byt缴款输入控制 = Val(zlDatabase.GetPara("结帐缴款输入控制", glngSys, mlngModul, 0))
    End With
End Sub
Private Sub cmdCancel_Click()
    If Val(txtUnit.Tag) = 0 Then
        Unload frmPatientsSelect
        Unload Me
    Else
        Call NewBalance
    End If
End Sub
Public Function ShowMe(ByVal frmMain As Object, ByVal bytInState As Byte, _
    ByVal lngModul As Long, ByVal strPrivs As String, _
    Optional ByVal lng结帐ID As Long, Optional ByVal blnViewCancel As Boolean, Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:合药单位结帐
    '入参:bytInState -0=结帐状态(默认新增),1=浏览状态
    '    blnViewCancel-是否查看已作废单据
    '    lng结帐ID-要浏览的单据号结帐ID,当bytInState=1时有效
    '    blnNOMoved-操作的单据是否在后备数据表中
    '返回:结帐成功1次以上,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-02-05 11:17:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mbytInState = bytInState: mblnViewCancel = blnViewCancel: mlng结帐ID = lng结帐ID
    mblnNOMoved = blnNOMoved: mintSucces = 0
    mlngModul = lngModul: mstrPrivs = strPrivs
    If Not gfrmMain Is Nothing Then
        Me.Show , frmMain
    Else
        Me.Show 1, frmMain
    End If
    ShowMe = mintSucces >= 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function isValied(ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结帐数据的合法性
    '出参:tyBrushCard当前刷卡信息
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-18 10:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSetFocus As Object, objCard As Card
    Dim intMouse As Integer
    Set objCard = IDKindPaymentsType.GetCurCard
    
    On Error GoTo errHandle
    If Val(txtUnit.Tag) = 0 Then
        MsgBox "请先选择合约单位的结帐病人!", vbInformation, gstrSysName
        Set objSetFocus = txtUnit: GoTo GoExit:
        Exit Function
    End If
    If IsFirstInputBalanceMoney Then
        '第一次输入时,需要检查总费用并提示
        If Val(lbl自付合计.Caption) = 0 Then
            If MsgBox("所选病人实际没有可结费用,要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Set objSetFocus = txtUnit: GoTo GoExit:
            End If
        End If
    End If
    If InStr(txtBalance(Idx_结算号码).Text, "'") > 0 Then
        MsgBox "结算号码含有非法字符单引号,不允许结帐", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_结算号码): GoTo GoExit:
         Exit Function
    End If
    
    If zlCommFun.ActualLen(txtBalance(Idx_结算号码).Text) > 30 Then
        MsgBox "结算号码最多只能输入30个字符或15个汉字,不允许结帐", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_结算号码): GoTo GoExit:
         Exit Function
    End If
    
    If InStr(txtBalance(Idx_摘要).Text, "'") > 0 Then
        MsgBox "摘要含有非法字符单引号,不允许结帐", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_摘要): GoTo GoExit:
         Exit Function
    End If
 
    If zlCommFun.ActualLen(txtBalance(Idx_摘要).Text) > 30 Then
        MsgBox "结算号码最多只能输入50个字符或25个汉字,不允许结帐", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_摘要): GoTo GoExit:
         Exit Function
    End If
    
    '发票检查
    If CheckFactIsValied(objSetFocus) = False Then GoTo GoExit:
    '检查当前输入的卡对象的数据合法性
    If CheckCurBalanceIsValied(tyBrushCard, False, objSetFocus) = False Then GoTo GoExit:
    isValied = True
    Exit Function
    
GoExit:
    If objSetFocus Is Nothing Then Exit Function
    If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMouse
        Resume
    End If
End Function

Private Function CheckCurBalanceIsValied(ByRef tyBrushCard As TY_BrushCard, ByVal bln预交 As Boolean, Optional ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前结帐是否有效
    '出参:tyBrushCard当前刷卡信息
    '     objSetFocus-光标移动对象
    '返回:有效返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 14:57:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng病人ID As Long, varData As Variant
    Dim dblMoney As Double, i As Long, blnFind As Boolean
    Dim cllDeposit As Collection, int性质 As Integer
    Dim intMouse As Integer
    Dim intCount As Integer '多种结算方式(排开医保)
    On Error GoTo errHandle
    
    intMouse = Screen.MousePointer
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前不存在有效的支付方式，请选择有效的支付方式!", vbInformation + vbOKOnly, gstrSysName
        Set objSetFocus = IDKindPaymentsType
        Exit Function
    End If
    
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If bln预交 Then
                If int性质 = 1 Then blnFind = True: Exit For
            End If
            If Not (objCard.消费卡 And objCard.自制卡) Then '消费卡,已经检查,不用再处理
                If .TextMatrix(i, .ColIndex("支付方式")) = objCard.结算方式 Then blnFind = True
            End If
            If InStr("34", int性质) > 0 Then
                MsgBox "不允许使用:" & .TextMatrix(i, .ColIndex("支付方式")) & "进行结帐!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            int性质 = Val(.TextMatrix(i, .ColIndex("结算性质")))
            If InStr(",1,2,", "," & int性质 & ",") > 0 Then intCount = intCount + 1
        Next
        
        If blnFind Then
            Screen.MousePointer = 0
            If bln预交 Then
                MsgBox "已经用预存款支付,只有删除预存款后才能支付!", vbOKOnly, gstrSysName
            Else
                MsgBox objCard.结算方式 & " 已经支付了,不能再用" & objCard.结算方式 & "进行支付", vbOKOnly + vbDefaultButton1, gstrSysName
            End If
            Exit Function
        End If
        
        If InStr("34", objCard.结算性质) > 0 Then
            MsgBox "不允许使用:" & objCard.结算方式 & "进行结帐!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With

    '数据检查接口数(目前只同时支持两种接口(含医保算一种接口)
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    
    
    '1.消费卡检查
    If CheckSquareBalanceValied(objCard, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_缴款)
        Exit Function
    End If
     
    '2.三方帐户检查
    If CheckThreeSwapValied(objCard, dblMoney, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_缴款)
        Exit Function
    End If
    
    '3.一卡通(老版)检查
    If CheckOldOneCardIsValied(objCard, dblMoney, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_缴款)
        Exit Function
    End If
    
    '4.检查现金结算方式
    If CheckCashValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_缴款)
        Exit Function
    End If
    
    '5.检查支票结算方式是否合法
    If CheckChequeValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_缴款)
        Exit Function
    End If
    
    '6.检查其他结算方式
    If CheckOtherValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_缴款)
        Exit Function
    End If
    
    
    CheckCurBalanceIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckOtherValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查其他结算方式(支票等)的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-14 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl当前未付 As Double
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.接口序号 > 0 Or objCard.结算方式 Like "*支票*" Or objCard.结算性质 = 1 Then CheckOtherValied = True: Exit Function
    
    dbl当前未付 = mtyBalanceInfor.dbl未付合计
    strTittle = IIf(dbl当前未付 < 0, "退款", "收款")
    dblMoney = Format(Val(txtBalance(Idx_缴款).Text), "0.00")
  
    If strTittle = "收款" Then
        If FormatEx(dblMoney, 6) = 0 Then
            Screen.MousePointer = 0
            MsgBox "未输入" & strTittle & "金额！", vbInformation, gstrSysName
            Exit Function
        End If
        If dblMoney > FormatEx(dbl当前未付, 2) Then
            Screen.MousePointer = 0
            MsgBox "注意:" & vbCrLf & "    输入的" & strTittle & "金额大于了未支付的金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckOtherValied = True
        Exit Function
    End If
    
    '退款
    If FormatEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "未输入" & strTittle & "金额！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If dblMoney > FormatEx(Abs(dbl当前未付), 2) Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "    输入的退款金额大于了未退金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckChequeValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查支票结算方式的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl当前未付 As Double
    Dim intMousePointer As Integer
    Dim objTempCard As Card
    Dim blnCheck As Boolean
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.结算性质 <> 2 Or Not objCard.结算方式 Like "*支票*" Then CheckChequeValied = True: Exit Function
    
    
    dbl当前未付 = mtyBalanceInfor.dbl未付合计
    
    strTittle = IIf(dbl当前未付 < 0, "退款", "收款")
    dblMoney = Format(Val(txtBalance(Idx_缴款).Text), "0.00")
     
    If strTittle = "收款" Then
    
        If FormatEx(dblMoney, 6) = 0 Then
            Screen.MousePointer = 0
            MsgBox "未输入收款金额！", vbInformation, gstrSysName
            Exit Function
        End If
        If mstr退支票 = "" And blnCheck Then
            Screen.MousePointer = 0
            MsgBox "在结算方式中没有设置应付款的结算方式,不能进行退支票处理", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckChequeValied = True
        Exit Function
    End If
    
    '退款
    If FormatEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "未输入退款金额！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckChequeValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckCashValied(ByVal objCard As Card, Optional ByVal bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查现金结算方式的一些合法情检查
    '入参:objCard－当前支付卡
    '     bln退款
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, strTittle As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer

    
    On Error GoTo errHandle
    If objCard.结算性质 <> 1 Then CheckCashValied = True: Exit Function
    
    dblMoney = Format(Val(txtBalance(Idx_缴款).Text), "0.00")
    If Not bln退款 Then
        If FormatEx(dblMoney, 6) <> 0 Then
            If Val(dblMoney) < Val(lbl剩余自付.Caption) Then
                Screen.MousePointer = 0
                MsgBox "收款金额不足,请补足应收金额！" & vbCrLf & "本次应收:" & lbl剩余自付.Caption & vbCrLf & "当前收款" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
                Exit Function
            End If
        End If
        '43153
        '缴款控制:0-不进行控制;1-存在收取现金时,必须输入缴款.
        If mty_ModulePara.byt缴款输入控制 = 0 Then CheckCashValied = True: Exit Function
        If txtBalance(Idx_缴款).Text = "" Then
            Screen.MousePointer = 0
            MsgBox "你还未输入缴款金额,不能继续", vbExclamation, gstrSysName
            Exit Function
        End If

        CheckCashValied = True
        Exit Function
    End If
    
    '退款处理
    If dblMoney < Abs(Val(lbl剩余自付.Caption)) And FormatEx(dblMoney, 6) <> 0 Then
        Screen.MousePointer = 0
        MsgBox "输入的退款金额不足！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCashValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    
    Call SaveErrLog
End Function

Private Function CheckOldOneCardIsValied(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef tyBrushCard As TY_BrushCard, _
    Optional bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '入参:objCard-当前卡对象
    '     bln退款-是否退款
    '出参:tyBrushCard-返回刷卡信息
    '返回:一卡通验证正确或非一卡通,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 17:19:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl未付金额 As Double, strCardNo As String
    Dim dblTemp As Double, strXmlIn As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.结算性质 <> 7 Then CheckOldOneCardIsValied = True: Exit Function
    
    mOldOneCard.rsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mOldOneCard.rsOneCard.EOF Then
        Screen.MousePointer = 0
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        CheckOldOneCardIsValied = False: Exit Function
    End If

    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "一卡通接口创建失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblMoney = 0 Then dblMoney = Val(txtBalance(Idx_缴款).Text)
     
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "收款") & "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    dbl未付金额 = FormatEx(mtyBalanceInfor.dbl未付合计, 6)
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "收款") & "金额不能大于本次" & IIf(bln退款, "未退", "未收") & "金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Val(lbl剩余自付.Caption) <> dblMoney Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "收款") & "金额不足:" & Format(Abs(Val(lbl剩余自付.Caption)), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
            
            
    If Not bln退款 Then
       
       '弹出刷卡界面
       'zlBrushCard(frmMain As Object, _
       '    ByVal lngModule As Long, _
       '    ByVal rsClassMoney As ADODB.Recordset, _
       '    ByVal lngCardTypeID As Long, _
       '    ByVal bln消费卡 As Boolean, _
       '    ByVal strPatiName As String, ByVal strSex As String, _
       '    ByVal strOld As String, ByVal dbl金额 As Double, _
       '    Optional ByRef strCardNo As String, _
       '    Optional ByRef strPassWord As String, _
       '    Optional ByRef bln退费 As Boolean = False, _
       '    Optional ByRef blnShowPatiInfor As Boolean = False, _
       '    Optional ByRef bln退现 As Boolean = False, _
       '    Optional ByVal bln余额不足禁止 As Boolean = True) As Boolean
       '---------------------------------------------------------------------------------------------------------------------------------------------
       '功能:根据指定支付类别,弹出刷卡窗口
       '入参:rsClassMoney:收费类别,金额
       '        lngCardTypeID-为零时,为老一卡通刷卡
       '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
        
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, 0, False, _
        txtUnit.Text, "", "", dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
        False, True, False, False, Nothing, False, False, strXmlIn) = False Then Exit Function
        
        tyBrushCard.dbl帐户余额 = mobjICCard.GetSpare
        If tyBrushCard.dbl帐户余额 < dblMoney Then
            Screen.MousePointer = 0
            MsgBox "卡余额不够支付,请检查!" & vbCrLf & vbCrLf & _
            "   卡 余  额" & Format(tyBrushCard.dbl帐户余额, "0.00") & vbCrLf & _
            "   本次支付" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
            Exit Function
        End If
        staThis.Panels(2).Text = Format(tyBrushCard.dbl帐户余额, "0.00")
        staThis.Panels(2).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(tyBrushCard.dbl帐户余额, "0.00")
       
        CheckOldOneCardIsValied = True
        Exit Function
    End If
    '退款检查
    If mrsBalance Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    mrsBalance.Filter = "类型=4"
    If mrsBalance.EOF Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        Screen.MousePointer = 0
        MsgBox "一卡通读卡失败,请将IC卡放在读卡器中", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> NVL(mrsBalance!卡号) Then
        Screen.MousePointer = 0
        MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblTemp = Format(Val(NVL(mrsBalance!冲预交)), "0.00")
    If FormatEx(dblMoney, 6) <> Format(dblTemp, "0.00") Then
        Screen.MousePointer = 0
        MsgBox "一卡通结算必须全退,请检查!" & vbCrLf & vbCrLf & _
        "   结算金额" & Format(dblTemp, "0.00") & vbCrLf & _
        "   本次支付" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckOldOneCardIsValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function
  
  
Private Function IsFirstInputBalanceMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否第一次输入缴款金额
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-18 11:29:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then Exit Function
        Next
    End With
    IsFirstInputBalanceMoney = True
End Function

Private Function CheckFactIsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否有效
    '出参:objSetFocus -出错时,光标定位到哪个对象
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-18 11:09:38
    '说明:第一次输入缴款数据时检查
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, i As Long
    On Error GoTo errHandle
    
    If Not IsFirstInputBalanceMoney Then CheckFactIsValied = True: Exit Function
    '第一次输入,就需检查
    '票据号码检查
    mblnPrintBill = False
    If mobjFact.打印方式 = 0 Then CheckFactIsValied = True: Exit Function
 
    mblnPrintBill = True
    If mobjFact.打印方式 = 2 Then
        If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then mblnPrintBill = False
        If mblnPrintBill = False Then CheckFactIsValied = True: Exit Function
    End If
    
    If Not mobjFact.严格控制 Then
        If Len(txtInvoice.Text) <> mobjFact.票号长度 And txtInvoice.Text <> "" Then
            MsgBox "票据号码长度应该为 " & mobjFact.票号长度 & " 位！", vbInformation, gstrSysName
            Set objSetFocus = txtInvoice: Exit Function
        End If
        CheckFactIsValied = True: Exit Function
    End If
    
    '严格票据管理
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
        Set objSetFocus = txtInvoice: Exit Function
    End If
    
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.用户名, mobjFact.票种, mobjFact.使用类别, mlng领用ID, mobjFact.共享批次ID, mlng领用ID, 1, Trim(txtInvoice.Text)) = False Then Exit Function
    
    If mlng领用ID <= 0 Then
        Select Case mlng领用ID
            Case 0 '操作失败
            Case -1
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Case -3
                MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入", vbInformation, gstrSysName
                Set objSetFocus = txtInvoice: Exit Function
        End Select
        Exit Function
    End If
    CheckFactIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LockedScreen(ByVal blnLocked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:锁定屏
    '编制:刘兴洪
    '日期:2015-05-18 15:55:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    cmdOK.Enabled = Not blnLocked
    txtBalance(Idx_缴款).Locked = blnLocked
    txtBalance(Idx_结算号码).Locked = blnLocked
    txtBalance(Idx_摘要).Locked = blnLocked
    txtBalance(Idx_lbl找补).Locked = True
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveBalanceData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结帐数据
    '编制:刘兴洪
    '日期:2015-05-18 10:57:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngBalanceID As Long, i As Long, lngPatientID As Long
    Dim strNO As String, strTmp As String
    Dim tyBrushCard As TY_BrushCard
    Dim dbl未付金额 As Double, dbl剩余金额 As Double, dbl退支票额 As Double
    Dim dblMoney As Double, dblTemp As Double
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim objCard As Card, intSign As Integer
    Dim strCardNo As String, blnTrans As Boolean
    
    
    On Error GoTo errHandle
    Call LockedScreen(True)
    
    Screen.MousePointer = 11
    If isValied(tyBrushCard) = False Then
        Call LockedScreen(False)
        Screen.MousePointer = 0: Exit Function
    End If
    Set objCard = IDKindPaymentsType.GetCurCard

    With mtyBalanceInfor
        .dbl缴款 = 0: .dbl找补 = 0
        .dbl现金 = 0
        dbl未付金额 = FormatEx(.dbl未付合计, 6)
    End With
    
    intSign = IIf(dbl未付金额 < 0, -1, 1)
    If objCard.结算性质 = 1 Then     '现金
        dblMoney = FormatEx(intSign * Val(txtBalance(Idx_缴款).Text), 6)
        If dblMoney <> 0 Then
            mtyBalanceInfor.dbl缴款 = dblMoney
            mtyBalanceInfor.dbl找补 = IIf(lblBalance(Idx_lbl找补).Caption Like "退*", -1, 1) * Val(txtBalance(Idx_找补).Text)
        End If
        dblTemp = dbl未付金额: dbl剩余金额 = 0
        dblMoney = GetCentMoney(dblTemp)
        mtyBalanceInfor.dbl现金 = dblMoney
    ElseIf objCard.名称 Like "*支票" Then
        dblMoney = FormatEx(intSign * (Val(txtBalance(Idx_缴款).Text)), 6)
        dbl剩余金额 = FormatEx(dbl未付金额 - dblMoney, 6)
        If dbl剩余金额 < 0 Then
            mtyBalanceInfor.dbl退支票 = -1 * Val(txtBalance(Idx_找补).Text)
            dbl剩余金额 = 0
        End If
    Else    '其他结算方式支付
        dblMoney = FormatEx(intSign * Val(txtBalance(Idx_缴款).Text), 6)
        dbl剩余金额 = FormatEx(dbl未付金额 - dblMoney, 6)
    End If
    
    Call Show误差金额
    If Abs(mtyBalanceInfor.dbl误差额) > 1.5 Then
        Screen.MousePointer = 0
        Call MsgBox("误差过大,请检查是否正确!", vbInformation + vbOKOnly, gstrSysName)
        Call LockedScreen(False)
        Exit Function
    End If
    If FormatEx(mtyBalanceInfor.dbl误差额, 6) <> 0 Then
        If mstrErrorBalance = "" Then
            Screen.MousePointer = 0
            MsgBox "在应用场合中未设置误差项，请在结算方式中设置!", vbInformation + vbOKOnly, gstrSysName
            Call LockedScreen(False)
            Exit Function
        End If
    End If
    If dbl剩余金额 <> 0 Then
        '还有剩余数据
        With vsBlance
            If objCard.消费卡 Then
                Call AddSquareBalance(objCard)
            Else
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                strCardNo = tyBrushCard.str卡号
                .TextMatrix(1, .ColIndex("是否密文")) = 0
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                If objCard.结算性质 = 7 And objCard.接口序号 < 0 Then
                    .TextMatrix(1, .ColIndex("类型")) = 4
                    .TextMatrix(1, .ColIndex("编辑状态")) = 0   '0-禁止删除;1-允许编辑金额;2-允许删除
                    .TextMatrix(1, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                ElseIf objCard.接口序号 > 0 Then
                    .TextMatrix(1, .ColIndex("类型")) = 3
                    .TextMatrix(1, .ColIndex("卡类别ID")) = objCard.接口序号
                    .TextMatrix(1, .ColIndex("卡类别名称")) = objCard.名称
                    .TextMatrix(1, .ColIndex("编辑状态")) = 0   '0-禁止删除;1-允许编辑金额;2-允许删除
                    .TextMatrix(1, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                    .TextMatrix(1, .ColIndex("是否密文")) = IIf(objCard.卡号密文规则 <> "", 1, 0)
                Else
                    .TextMatrix(1, .ColIndex("类型")) = 0
                    .TextMatrix(1, .ColIndex("编辑状态")) = 2   '0-禁止删除;1-允许编辑金额;2-允许删除
                    .TextMatrix(1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                End If
                .TextMatrix(1, .ColIndex("结算性质")) = objCard.结算性质
                .TextMatrix(1, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(1, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(1, .ColIndex("校对标志")) = 2
                
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                .TextMatrix(1, .ColIndex("金额")) = Format(dblMoney, "0.00")
                .TextMatrix(1, .ColIndex("结算号码")) = IIf(txtBalance(Idx_结算号码).Visible, Trim(txtBalance(Idx_结算号码).Text), "")
                .TextMatrix(1, .ColIndex("备注")) = Trim(txtBalance(Idx_摘要).Text)
                .TextMatrix(1, .ColIndex("卡号")) = IIf(objCard.卡号密文规则 <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = tyBrushCard.str卡号
                .TextMatrix(1, .ColIndex("交易流水号")) = tyBrushCard.str交易流水号
                .TextMatrix(1, .ColIndex("交易说明")) = tyBrushCard.str交易说明
                mtyBalanceInfor.dbl已付合计 = FormatEx(mtyBalanceInfor.dbl已付合计 + dblMoney, 6)
                mtyBalanceInfor.dbl未付合计 = FormatEx(mtyBalanceInfor.dbl未付合计 - dblMoney, 6)
            End If
            For i = 1 To IDKindPaymentsType.ListCount
                '缺省定位在现金上
                 Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
                If objCard.结算性质 = 1 Then IDKindPaymentsType.IDKIND = i: Exit For
            Next
        End With
        Call LockedScreen(False)
        Screen.MousePointer = 0
        If txtBalance(Idx_缴款).Enabled And txtBalance(Idx_缴款).Visible Then txtBalance(Idx_缴款).SetFocus
        txtBalance(Idx_缴款).Text = ""
        Call LoadCurOwnerPayInfor
        SaveBalanceData = True
        Exit Function
    End If
    
    
    Set cllPro = New Collection: Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    '保存数据
    If GetSaveBalanceSQL(tyBrushCard, mtyBalanceInfor, cllPro) = False Then Exit Function
    '执行一卡通(老版)接口
    If ExecuteOldOneCardPayInterface(0, mtyBalanceInfor.lng结帐ID, objCard, dblMoney, tyBrushCard, cllPro) = False Then
        Call LockedScreen(False): Screen.MousePointer = 0
        Exit Function
    End If
    '执持三方帐户交易接口
    If ExecuteThreeSwapPayInterface(0, mtyBalanceInfor.lng结帐ID, objCard, dblMoney, cllPro, tyBrushCard) = False Then
        Call LockedScreen(False): Screen.MousePointer = 0
        Exit Function
    End If
    If cllPro.Count <> 0 Then
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption
        blnTrans = False
    End If
    
    Call LockedScreen(False): Screen.MousePointer = 0
    Call PrintBill '打印票据
    
    txt累计金额.Text = Format(Val(txt累计金额.Text) + Val(lbl自付合计.Caption), gstrDec)
    
    '单据历史记录
    strNO = mtyBalanceInfor.strNO
    strTmp = mtyBalanceInfor.strNO
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next
    Call NewBalance
    mintSucces = mintSucces + 1
    SaveBalanceData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call LockedScreen(False)
End Function

Private Sub PrintBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印票据
    '编制:刘兴洪
    '日期:2015-05-18 15:48:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, i As Long
    
    On Error GoTo errHandle
    '票据打印
    If Not mblnPrintBill Then Exit Sub
    If Not gblnPrintByPatient Then
        '不按病人打印
        Call frmPrint.ReportPrint(1, mtyBalanceInfor.strNO, mtyBalanceInfor.lng结帐ID, mobjFact, txtInvoice.Text, mtyBalanceInfor.dtBalanceDate, CStr(mtyBalanceInfor.dbl缴款), CStr(mtyBalanceInfor.dbl找补), , mobjFact.打印格式)
        Exit Sub
    End If
    '按病人打印
    If mrsPatients.RecordCount > 0 Then mrsPatients.MoveFirst
    For i = 1 To mrsPatients.RecordCount
        lng病人ID = Val(NVL(mrsPatients!病人ID))
        Call frmPrint.ReportPrint(1, mtyBalanceInfor.strNO, mtyBalanceInfor.lng结帐ID, mobjFact, txtInvoice.Text, mtyBalanceInfor.dtBalanceDate, CStr(mtyBalanceInfor.dbl缴款), CStr(mtyBalanceInfor.dbl找补), lng病人ID, mobjFact.打印格式)
        If i < mrsPatients.RecordCount Then Call RefreshFact
        mrsPatients.MoveNext
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Function GetSaveBalanceSQL(ByRef tyBrushCard As TY_BrushCard, _
    ByRef tyBalanceInfor As TY_Balance_Infor, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐数据保存
    '入参:tbBrushCard-当前刷卡信息
    '出参:tyBalanceInfor-返回当前结帐信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-18 14:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng结帐ID As Long, strTmp As String, str误差NO As String, strPatiIDs As String
    Dim lngFirstPatiID As Long   '仅在存在误差时,传入过程中用于生成误差记帐单相关信息
    Dim strSql As String, objCard As Card
    Dim strNO As String, datBalance As Date, dblMoney As Double
    Dim str消费卡结算 As String
    
    
    Err = 0: On Error GoTo ErrHand:
    
    If mrsPatients.RecordCount > 0 Then mrsPatients.MoveFirst
    For i = 1 To mrsPatients.RecordCount
        If i = 1 Then lngFirstPatiID = Val(mrsPatients!病人ID)
        strPatiIDs = strPatiIDs & mrsPatients!病人ID & ","
        mrsPatients.MoveNext
    Next
    Set objCard = IDKindPaymentsType.GetCurCard
    strNO = zlDatabase.GetNextNo(15)
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    datBalance = zlDatabase.Currentdate
    
    With tyBalanceInfor
        .strNO = strNO
        .lng结帐ID = lng结帐ID
        .dtBalanceDate = datBalance
    End With
    '1.病人结帐记录
    'Zl_病人结帐记录_Insert
    strSql = "zl_病人结帐记录_Insert("
    '  Id_In           病人结帐记录.ID%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  单据号_In       病人结帐记录.NO%Type,
    strSql = strSql & "'" & strNO & "',"
    '  病人id_In       病人结帐记录.病人id%Type,
    strSql = strSql & "" & lngFirstPatiID & ","
    '  收费时间_In     病人结帐记录.收费时间%Type,
    strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  开始日期_In     病人结帐记录.开始日期%Type,
    strSql = strSql & "NULL,"
    '  结束日期_In     病人结帐记录.结束日期%Type,
    strSql = strSql & "NULL,"
    '  中途结帐_In     病人结帐记录.中途结帐%Type := 0,
    strSql = strSql & "0,"
    '  多病人结帐_In   Number := 0,
    strSql = strSql & "1,"
    '  最大结帐次数_In Number := 0,
    strSql = strSql & "0,"
    '  备注_In         病人结帐记录.备注%Type := Null,
    strSql = strSql & "NULL,"
    '  来源_In         Number := 1, '  --1.来源_In:1-门诊;2-住院
    strSql = strSql & "1,"
    '  原因_In         病人结帐记录.原因%Type := Null   '存储合约单位ID'问题:35090
    strSql = strSql & "'" & Trim(txtUnit.Text) & "',1)"
    zlAddArray cllPro, strSql
    '2.结帐缴款记录
    With vsBlance
        For i = 1 To .Rows - 1
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If Val(.TextMatrix(i, .ColIndex("类型"))) <> 5 And _
                Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 Then
                'Zl_结帐缴款记录_Insert
                strSql = "zl_结帐缴款记录_Insert("
                '  No_In         病人结帐记录.No%Type,
                strSql = strSql & "'" & strNO & "',"
                '  病人id_In     病人预交记录.病人id%Type,
                strSql = strSql & "NULL,"
                '  主页id_In     病人预交记录.主页id%Type,
                strSql = strSql & "NULL,"
                '  科室id_In     病人预交记录.科室id%Type,
                strSql = strSql & "NULL,"
                '  结算方式_In   病人预交记录.结算方式%Type,
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("支付方式")) & "',"
                '  结算号码_In   病人预交记录.结算号码%Type,
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("结算号码")) & "',"
                '  金额_In       病人预交记录.冲预交%Type,
                strSql = strSql & "" & Val(.TextMatrix(i, .ColIndex("金额"))) & ","
                '  结帐id_In     病人预交记录.结帐id%Type,
                strSql = strSql & "" & lng结帐ID & ","
                '  操作员编号_In 病人预交记录.操作员编号%Type,
                strSql = strSql & "'" & UserInfo.编号 & "',"
                '  操作员姓名_In 病人预交记录.操作员姓名%Type,
                strSql = strSql & "'" & UserInfo.姓名 & "',"
                '  收费时间_In   病人预交记录.收款时间%Type,
                strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  保险类别_In   保险帐户.险类%Type,
                strSql = strSql & "NULL,"
                '  保险帐号_In   保险帐户.医保号%Type,
                strSql = strSql & "NULL,"
                '  保险密码_In   保险帐户.密码%Type,
                strSql = strSql & "NULL,"
                '  缴款_In       病人预交记录.缴款%Type := Null,
                strSql = strSql & "NULL,"
                '  找补_In       病人预交记录.找补%Type := Null,
                strSql = strSql & "NULL,"
                '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                If Val(.TextMatrix(i, .ColIndex("卡类别ID"))) <> 0 Then
                    strSql = strSql & "" & Val(.TextMatrix(i, .ColIndex("卡类别ID"))) & ","
                Else
                    strSql = strSql & "NULL,"
                End If
                '  卡号_In       病人预交记录.卡号%Type := Null,
                If Trim(.TextMatrix(i, .ColIndex("卡号"))) <> "" Then
                    strSql = strSql & "'" & Trim(.TextMatrix(i, .ColIndex("卡号"))) & "',"
                Else
                    strSql = strSql & "NULL,"
                End If
                '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                If Trim(.TextMatrix(i, .ColIndex("交易说明"))) <> "" Then
                    strSql = strSql & "'" & Trim(.TextMatrix(i, .ColIndex("交易流水号"))) & "',"
                Else
                    strSql = strSql & "NULL,"
                End If
                '  交易说明_In   病人预交记录.交易说明%Type := Null
                If Trim(.TextMatrix(i, .ColIndex("交易说明"))) <> "" Then
                    strSql = strSql & "'" & Trim(.TextMatrix(i, .ColIndex("交易说明"))) & "',"
                Else
                    strSql = strSql & "NULL,"
                End If
                '  消费卡结算_In Varchar2 := Null:卡类别ID|卡号|消费卡ID|消费金额||."
                strSql = strSql & "NULL)"
                '退支票额_In   病人预交记录.冲预交%Type := Null
                zlAddArray cllPro, strSql
            ElseIf Val(.TextMatrix(i, .ColIndex("类型"))) = 5 Then
                If (objCard.接口序号 <> Val(.TextMatrix(i, .ColIndex("卡类别ID"))) And objCard.消费卡) Or Not objCard.消费卡 Then
                    '消费卡
                    str消费卡结算 = str消费卡结算 & "||" & Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                    str消费卡结算 = str消费卡结算 & "|" & Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
                    str消费卡结算 = str消费卡结算 & "|" & Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                    str消费卡结算 = str消费卡结算 & "|" & FormatEx(Val(.TextMatrix(i, .ColIndex("金额"))), 6)
                End If
            End If
        Next
    End With
    If objCard.接口序号 > 0 Then
        If objCard.消费卡 Then
            For i = 1 To mcllCurSquareBalance.Count
                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                str消费卡结算 = str消费卡结算 & "||" & Val(mcllCurSquareBalance(i)(0))
                str消费卡结算 = str消费卡结算 & "|" & Trim(mcllCurSquareBalance(i)(3))
                str消费卡结算 = str消费卡结算 & "|" & Val(mcllCurSquareBalance(i)(1))
                str消费卡结算 = str消费卡结算 & "|" & FormatEx(Val(mcllCurSquareBalance(i)(2)), 6)
            Next
        Else
            'Zl_结帐缴款记录_Insert
            strSql = "zl_结帐缴款记录_Insert("
            '  No_In         病人结帐记录.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  病人id_In     病人预交记录.病人id%Type,
            strSql = strSql & "NULL,"
            '  主页id_In     病人预交记录.主页id%Type,
            strSql = strSql & "NULL,"
            '  科室id_In     病人预交记录.科室id%Type,
            strSql = strSql & "NULL,"
            '  结算方式_In   病人预交记录.结算方式%Type,
            strSql = strSql & "'" & objCard.结算方式 & "',"
            '  结算号码_In   病人预交记录.结算号码%Type,
            strSql = strSql & "'" & txtBalance(Idx_结算号码).Text & "',"
            '  金额_In       病人预交记录.冲预交%Type,
            strSql = strSql & "" & txtBalance(Idx_缴款).Text & ","
            '  结帐id_In     病人预交记录.结帐id%Type,
            strSql = strSql & "" & lng结帐ID & ","
            '  操作员编号_In 病人预交记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 病人预交记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  收费时间_In   病人预交记录.收款时间%Type,
            strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  保险类别_In   保险帐户.险类%Type,
            strSql = strSql & "NULL,"
            '  保险帐号_In   保险帐户.医保号%Type,
            strSql = strSql & "NULL,"
            '  保险密码_In   保险帐户.密码%Type,
            strSql = strSql & "NULL,"
            '  缴款_In       病人预交记录.缴款%Type := Null,
            strSql = strSql & "NULL,"
            '  找补_In       病人预交记录.找补%Type := Null,
            strSql = strSql & "NULL,"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & objCard.接口序号 & ","
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & tyBrushCard.str卡号 & "',"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSql = strSql & "'" & tyBrushCard.str交易流水号 & "',"
            '  交易说明_In   病人预交记录.交易说明%Type := Null
            strSql = strSql & "'" & tyBrushCard.str交易说明 & "',"
            '  消费卡结算_In Varchar2 := Null:卡类别ID|卡号|消费卡ID|消费金额||."
            strSql = strSql & "NULL)"
            '退支票额_In   病人预交记录.冲预交%Type := Null
            zlAddArray cllPro, strSql
        End If
    Else
        '普通结算方式
        dblMoney = Format(Val(txtBalance(Idx_缴款).Text), mstrDec)
        If objCard.结算性质 = 1 Then
            dblMoney = mtyBalanceInfor.dbl现金
        End If
        
        'Zl_结帐缴款记录_Insert
        strSql = "zl_结帐缴款记录_Insert("
        '  No_In         病人结帐记录.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  病人id_In     病人预交记录.病人id%Type,
        strSql = strSql & "NULL,"
        '  主页id_In     病人预交记录.主页id%Type,
        strSql = strSql & "NULL,"
        '  科室id_In     病人预交记录.科室id%Type,
        strSql = strSql & "NULL,"
        '  结算方式_In   病人预交记录.结算方式%Type,
        strSql = strSql & "'" & objCard.结算方式 & "',"
        '  结算号码_In   病人预交记录.结算号码%Type,
        strSql = strSql & "'" & txtBalance(Idx_结算号码).Text & "',"
        '  金额_In       病人预交记录.冲预交%Type,
        strSql = strSql & "" & dblMoney & ","
        '  结帐id_In     病人预交记录.结帐id%Type,
        strSql = strSql & "" & lng结帐ID & ","
        '  操作员编号_In 病人预交记录.操作员编号%Type,
        strSql = strSql & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 病人预交记录.操作员姓名%Type,
        strSql = strSql & "'" & UserInfo.姓名 & "',"
        '  收费时间_In   病人预交记录.收款时间%Type,
        strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  保险类别_In   保险帐户.险类%Type,
        strSql = strSql & "NULL,"
        '  保险帐号_In   保险帐户.医保号%Type,
        strSql = strSql & "NULL,"
        '  保险密码_In   保险帐户.密码%Type,
        strSql = strSql & "NULL,"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        strSql = strSql & "NULL,"
        '  找补_In       病人预交记录.找补%Type := Null,
        strSql = strSql & "NULL,"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSql = strSql & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSql = strSql & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSql = strSql & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null
        strSql = strSql & "NULL,"
        '  消费卡结算_In Varchar2 := Null:卡类别ID|卡号|消费卡ID|消费金额||."
        strSql = strSql & "NULL,"
        '退支票额_In   病人预交记录.冲预交%Type := Null
        If objCard.结算方式 Like "*支票" And Val(txtBalance(Idx_lbl找补).Text) <> 0 Then
            strSql = strSql & "" & mtyBalanceInfor.dbl退支票 & ")"
        Else
            strSql = strSql & "NULL)"
        End If
        zlAddArray cllPro, strSql
    End If
    
    '处理消费卡
    If str消费卡结算 <> "" Then
        str消费卡结算 = Mid(str消费卡结算, 3)
       'Zl_结帐缴款记录_Insert
            strSql = "zl_结帐缴款记录_Insert("
            '  No_In         病人结帐记录.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  病人id_In     病人预交记录.病人id%Type,
            strSql = strSql & "NULL,"
            '  主页id_In     病人预交记录.主页id%Type,
            strSql = strSql & "NULL,"
            '  科室id_In     病人预交记录.科室id%Type,
            strSql = strSql & "NULL,"
            '  结算方式_In   病人预交记录.结算方式%Type,
            strSql = strSql & "NULL,"
            '  结算号码_In   病人预交记录.结算号码%Type,
            strSql = strSql & "NULL,"
            '  金额_In       病人预交记录.冲预交%Type,
            strSql = strSql & "NULL,"
            '  结帐id_In     病人预交记录.结帐id%Type,
            strSql = strSql & "" & lng结帐ID & ","
            '  操作员编号_In 病人预交记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 病人预交记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  收费时间_In   病人预交记录.收款时间%Type,
            strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  保险类别_In   保险帐户.险类%Type,
            strSql = strSql & "NULL,"
            '  保险帐号_In   保险帐户.医保号%Type,
            strSql = strSql & "NULL,"
            '  保险密码_In   保险帐户.密码%Type,
            strSql = strSql & "NULL,"
            '  缴款_In       病人预交记录.缴款%Type := Null,
            strSql = strSql & "NULL,"
            '  找补_In       病人预交记录.找补%Type := Null,
            strSql = strSql & "NULL,"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "NULL,"
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSql = strSql & "NULL,"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSql = strSql & "NULL,"
            '  交易说明_In   病人预交记录.交易说明%Type := Null
            strSql = strSql & "NULL,"
            '  消费卡结算_In Varchar2 := Null:卡类别ID|卡号|消费卡ID|消费金额||."
            strSql = strSql & "'" & str消费卡结算 & "')"
            '退支票额_In   病人预交记录.冲预交%Type := Null
            zlAddArray cllPro, strSql
    End If
    If mtyBalanceInfor.dbl误差额 <> 0 Then

        
        'Zl_结帐缴款记录_Insert
        strSql = "zl_结帐缴款记录_Insert("
        '  No_In         病人结帐记录.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  病人id_In     病人预交记录.病人id%Type,
        strSql = strSql & "NULL,"
        '  主页id_In     病人预交记录.主页id%Type,
        strSql = strSql & "NULL,"
        '  科室id_In     病人预交记录.科室id%Type,
        strSql = strSql & "NULL,"
        '  结算方式_In   病人预交记录.结算方式%Type,
        strSql = strSql & "'" & mstrErrorBalance & "',"
        '  结算号码_In   病人预交记录.结算号码%Type,
        strSql = strSql & "'" & txtBalance(Idx_结算号码).Text & "',"
        '  金额_In       病人预交记录.冲预交%Type,
        strSql = strSql & "" & mtyBalanceInfor.dbl误差额 & ","
        '  结帐id_In     病人预交记录.结帐id%Type,
        strSql = strSql & "" & lng结帐ID & ","
        '  操作员编号_In 病人预交记录.操作员编号%Type,
        strSql = strSql & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 病人预交记录.操作员姓名%Type,
        strSql = strSql & "'" & UserInfo.姓名 & "',"
        '  收费时间_In   病人预交记录.收款时间%Type,
        strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  保险类别_In   保险帐户.险类%Type,
        strSql = strSql & "NULL,"
        '  保险帐号_In   保险帐户.医保号%Type,
        strSql = strSql & "NULL,"
        '  保险密码_In   保险帐户.密码%Type,
        strSql = strSql & "NULL,"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        strSql = strSql & "NULL,"
        '  找补_In       病人预交记录.找补%Type := Null,
        strSql = strSql & "NULL,"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSql = strSql & "Null,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSql = strSql & "'" & tyBrushCard.str卡号 & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSql = strSql & "'" & tyBrushCard.str交易流水号 & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null
        strSql = strSql & "'" & tyBrushCard.str交易说明 & "',"
        '  消费卡结算_In Varchar2 := Null:卡类别ID|卡号|消费卡ID|消费金额||."
        strSql = strSql & "NULL,"
        '退支票额_In   病人预交记录.冲预交%Type := Null
        strSql = strSql & "NULL)"
        zlAddArray cllPro, strSql
    End If
    '3.门诊费用记录
    While strPatiIDs <> ""
        i = 0
        If Len(strPatiIDs) > 3998 Then
            i = InStrRev(Mid(strPatiIDs, 1, 3998), ",")
            strTmp = Mid(strPatiIDs, 1, i - 1)
            strPatiIDs = Mid(strPatiIDs, i + 1)
        Else
            strTmp = Mid(strPatiIDs, 1, Len(strPatiIDs) - 1)
            strPatiIDs = ""
        End If
        strSql = "Zl_结帐费用记录_Unit('" & strTmp & "'," & lng结帐ID & "," & IIf(gblnZero, 1, 0) & ")"
        zlAddArray cllPro, strSql
    Wend

    '4.开始票据号
    If mblnPrintBill And Trim(txtInvoice.Text) <> "" Then
        strSql = "Zl_票据起始号_Update('" & strNO & "','" & Trim(txtInvoice.Text) & "',3)"
        zlAddArray cllPro, strSql
    End If
    GetSaveBalanceSQL = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新收费票据号
    '编制:刘兴洪
    '日期:2015-02-05 11:40:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
       
    On Error GoTo errHandle
        
    If mobjFact.打印方式 = 0 Then Exit Sub
    If Not mobjFact.严格控制 Then
        '非严格控制下
        '松散：取下一个号码
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    If zlGetInvoiceGroupUseID(mlng领用ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '严格：取下一个号码
    If mobjInvoice.zlGetNextBill(1137, mlng领用ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
    '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
    '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    txtInvoice.SelStart = Len(txtInvoice.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.姓名, mobjFact.票种, _
        mobjFact.使用类别, lng领用ID, mobjFact.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng领用ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng领用ID
        Case 0 '操作失败
        Case -1
            If Trim(mobjFact.使用类别) = "" Then
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & mobjFact.使用类别 & "』结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFact.使用类别) = "" Then
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & mobjFact.使用类别 & "』结帐票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If mbytInState = 1 Then Unload Me: Exit Sub
    Call SaveBalanceData
End Sub
Private Sub cmd病人_Click()
    With frmPatientsSelect
        .mstrUnitName = txtUnit.Text
        .Show 1, Me
        
        If gblnOK Then
            Call NewBalance
            txtUnit.Tag = .mlngUnitID
            txtUnit.Text = .mstrUnitName
            txtUnit.Enabled = False
            Set mrsPatients = .mrsPatients
            Call ShowBalanceInfo(mrsPatients)
            Call ClearPayInfo
            If vsBlance.Rows > 0 Then vsBlance.SetFocus: vsBlance.Row = 1: vsBlance.Col = vsBlance.ColIndex("金额")
            staThis.Panels(2).Text = "共选择了" & CStr(vsPati.Rows - 1) & "位病人."
            gblnOK = False
            Call LoadCurOwnerPayInfor
            Call LoadDefaultMoney
            If txtBalance(Idx_lbl缴款).Enabled And txtBalance(Idx_lbl缴款).Visible Then
                txtBalance(Idx_lbl缴款).SetFocus
            End If
        ElseIf Val(txtUnit.Tag) = 0 Then
            Unload frmPatientsSelect
            Call txtUnit.SetFocus
        End If
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnUnload Then Unload Me: Exit Sub
    If mbytInState = 0 And txtUnit.Enabled Then
        txtUnit.SetFocus
    ElseIf mbytInState = 1 And cmdCancel.Enabled Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmd病人_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    If mbytInState = 0 Then
        mintDefault = 0: mlng领用ID = 0: mintError = 0
        Call InitFact
        Call NewBalance
        txt累计金额.Text = Format(0, gstrDec)
        txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnStrictCtrl '89302
    Else
        Call InitViewFace
        Call LoadPatiBalanceData
    End If
    If Init结算方式 = False Then Exit Sub
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面信息
    '编制:刘兴洪
    '日期:2015-05-18 10:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlRaisEffect picBalance, Dw_SubKen
    zlRaisEffect picOwerFee, Dw_SubKen
    zlRaisEffect picNotPayment, Dw_SubKen
    Call zlInitModulePara
    Call InitOldOneCardInfor
    
End Sub

Private Sub InitFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化发票信息
    '编制:刘兴洪
    '日期:2015-02-05 11:26:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytInvoiceKind     As Byte, intFormat As Integer
    Dim intPrintMode As Integer, lngShareUseID As Long
    On Error GoTo errHandle
    If mbytInState <> 0 Then Exit Sub
    
    bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, 1137, "0"))

    Set mobjInvoice = New clsInvoice: Set mobjFact = New clsFactProperty
    mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    Call mobjInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, mobjFact, , , 1)
    
    mobjFact.使用类别 = zlDatabase.GetPara("合约单位结帐打印", glngSys, 1137)
    mobjFact.票种 = IIf(bytInvoiceKind = 0, 3, 1)
    Call mobjInvoice.zlGetInvoicePrintFormat(1137, mobjFact.票种, mobjFact.使用类别, intFormat, 1)
    mobjFact.打印格式 = intFormat
    If mobjInvoice.zlGetInvoicePrintMode(1137, mobjFact.票种, mobjFact.使用类别, intPrintMode) = False Then Exit Sub
    mobjFact.打印方式 = intPrintMode
    If mobjInvoice.zlGetInvoiceShareID(1137, mobjFact.票种, mobjFact.使用类别, lngShareUseID) = False Then Exit Sub
    mobjFact.共享批次ID = lngShareUseID
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    
    vsPati.Height = Me.ScaleHeight - staThis.Height - 100
    vsPati.Width = Me.ScaleWidth - picRight.Width - 200
    
    picRight.Top = vsPati.Top
    picRight.Left = vsPati.Left + vsPati.Width + 50
    picRight.Height = vsPati.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjInvoice = Nothing
    Set mobjFact = Nothing
    
    mblnViewCancel = False
    mblnNOMoved = False
    mbytInState = 0
    mlng结帐ID = 0
    zl_vsGrid_Para_Save mlngModul, vsPati, Me.Name, "病人列表"
    zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "结算列表"
    
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub InitViewFace()
    '功能:查看时初始化界面
    
    If mblnViewCancel Then lblFlag.Visible = True
    
    Call InitPatiGrid
    Call InitGrid_PayList
    
    txtInvoice.Locked = True
    cboNO.Locked = True
    txtUnit.Locked = True
    txtUnit.Width = txtUnit.Width + cmd病人.Width + 30
    cmd病人.Visible = False
        
    lbl累计.Visible = False
    txt累计金额.Visible = False
    picCurBalance.Visible = False
    picOwerFee.Visible = False
    picNotPayment.Visible = False
    cmdOK.Visible = False
    cmdCancel.Caption = "退出(&X)"
    
    fraSplit.Visible = True
    
End Sub

Private Sub LoadPatiBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人结帐数据
    '编制:刘兴洪
    '日期:2015-05-15 15:41:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    
    '加载病人结算数据
    Set rsTmp = GetBalanceData(mlng结帐ID, mblnNOMoved)
    Call ShowBalanceInfo(rsTmp)
    Call LoadBalanceInfo(mlng结帐ID, mblnNOMoved)
    Call LoadOtherData(mlng结帐ID, mblnNOMoved)
    staThis.Panels(2).Text = "当前结帐单共有" & CStr(vsBlance.Rows - 1) & "位病人."
End Sub
Private Function LoadOtherData(ByVal lng结帐ID As Long, Optional ByVal blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载其他数据
    '入参:lng结帐ID-结帐ID
    '     blnNOMoved-是否已经转出到历史数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-15 17:08:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    
    Set rsTmp = GetOtherInfo(lng结帐ID, blnNOMoved)
    If Not rsTmp Is Nothing Then
        txtUnit.Text = "" & rsTmp!合约单位
        txtInvoice.Text = "" & rsTmp!实际票号
        cboNO.AddItem rsTmp!NO
        cboNO.ListIndex = cboNO.NewIndex
    End If
    LoadOtherData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
''
''Private Sub LoadErrorData(ByVal dblErrorMoney As Double)
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    '功能:加载误差数据
''    '入参:dblErrorMoney -误差金额
''    '编制:刘兴洪
''    '日期:2015-05-15 17:30:17
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    Dim lngError As Long, i As Long
''
''    On Error GoTo errHandle
''    lngError = GetBalanceExistError
''    If lngError <= 0 Then
''        With vsBlance
''        .Redraw = flexRDBuffered
''        .Rows = 2: .Clear 1
''        .Rows = IIf(mrsBalance.RecordCount + 1 = 1, 2, mrsBalance.RecordCount + 1)
''        If mrsBalance.RecordCount > 0 Then mrsBalance.MoveFirst
''        For i = 1 To mrsBalance.RecordCount
''            '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
''            .TextMatrix(i, .ColIndex("编辑状态")) = ""
''            ''是否已结算:1-已结算;0-未结算
''            .TextMatrix(i, .ColIndex("结算状态")) = 0
''
''            .TextMatrix(i, .ColIndex("卡类别ID")) = Val(nvl(mrsBalance!卡类别id))
''            .TextMatrix(i, .ColIndex("消费卡ID")) = Val(nvl(mrsBalance!消费卡ID))
''            .TextMatrix(i, .ColIndex("结算性质")) = Val(nvl(mrsBalance!结算性质))
''            .TextMatrix(i, .ColIndex("类型")) = Val(nvl(mrsBalance!类型))
''            .TextMatrix(i, .ColIndex("是否退现")) = Val(nvl(mrsBalance!是否退现))
''            .TextMatrix(i, .ColIndex("是否全退")) = Val(nvl(mrsBalance!是否全退))
''            .TextMatrix(i, .ColIndex("校对标志")) = Val(nvl(mrsBalance!校对标志))
''            .TextMatrix(i, .ColIndex("是否密文")) = Val(nvl(mrsBalance!是否密文))
''            .TextMatrix(i, .ColIndex("支付方式")) = nvl(mrsBalance!结算方式)
''            .TextMatrix(i, .ColIndex("金额")) = Format(mrsBalance!结算金额, mstrDec)
''            .TextMatrix(i, .ColIndex("结算号码")) = nvl(mrsBalance!结算号码)
''            .TextMatrix(i, .ColIndex("备注")) = nvl(mrsBalance!备注)
''            .TextMatrix(i, .ColIndex("卡号")) = nvl(mrsBalance!卡号)
''            .TextMatrix(i, .ColIndex("交易流水号")) = nvl(mrsBalance!交易流水号)
''            .TextMatrix(i, .ColIndex("交易说明")) = nvl(mrsBalance!交易说明)
''            .TextMatrix(i, .ColIndex("卡类别名称")) = nvl(mrsBalance!卡类别名称)
''            .Row = i: .Col = .ColIndex("支付方式")
''            .CellBackColor = &HE7CFBA
''            mrsBalance.MoveNext
''        Next
''        .Redraw = flexRDBuffered
''    End With
''
''    End If
''
''    Exit Sub
''errHandle:
''    If ErrCenter() = 1 Then
''        Resume
''    End If
''
''End Sub
Private Function GetBalanceExistError() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取误差行
    '返回:误差行:=-1表示未找到误差项
    '编制:刘兴洪
    '日期:2015-05-15 17:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("结算性质"))) = 9 Then
                GetBalanceExistError = i: Exit Function
            End If
        Next
    End With
    GetBalanceExistError = -1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LoadBalanceInfo(ByVal lng结帐ID As Long, Optional ByVal blnNOMoved As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算信息
    '入参:lng结帐ID-结帐ID
    '     blnNOMoved-是否已经转出到历史数据
    '编制:刘兴洪
    '日期:2015-05-15 15:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    vsBlance.Rows = 2: vsBlance.Clear 1
    If lng结帐ID = 0 Then Exit Sub
    
    '       字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '       类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    
    Set mrsBalance = zlFromIDGetChargeBalance(0, lng结帐ID, blnNOMoved)
    With vsBlance
        .Redraw = flexRDBuffered
        .Rows = 2: .Clear 1
        .Rows = IIf(mrsBalance.RecordCount + 1 = 1, 2, mrsBalance.RecordCount + 1)
        If mrsBalance.RecordCount > 0 Then mrsBalance.MoveFirst
        For i = 1 To mrsBalance.RecordCount
            '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
            .TextMatrix(i, .ColIndex("编辑状态")) = ""
            ''是否已结算:1-已结算;0-未结算
            .TextMatrix(i, .ColIndex("结算状态")) = 0
            
            .TextMatrix(i, .ColIndex("卡类别ID")) = Val(NVL(mrsBalance!卡类别ID))
            .TextMatrix(i, .ColIndex("消费卡ID")) = Val(NVL(mrsBalance!消费卡ID))
            .TextMatrix(i, .ColIndex("结算性质")) = Val(NVL(mrsBalance!结算性质))
            .TextMatrix(i, .ColIndex("类型")) = Val(NVL(mrsBalance!类型))
            .TextMatrix(i, .ColIndex("是否退现")) = Val(NVL(mrsBalance!是否退现))
            .TextMatrix(i, .ColIndex("是否全退")) = Val(NVL(mrsBalance!是否全退))
            .TextMatrix(i, .ColIndex("校对标志")) = Val(NVL(mrsBalance!校对标志))
            .TextMatrix(i, .ColIndex("是否密文")) = Val(NVL(mrsBalance!是否密文))
            .TextMatrix(i, .ColIndex("支付方式")) = NVL(mrsBalance!结算方式)
            .TextMatrix(i, .ColIndex("金额")) = Format(mrsBalance!冲预交, mstrDec)
            .TextMatrix(i, .ColIndex("结算号码")) = NVL(mrsBalance!结算号码)
            .TextMatrix(i, .ColIndex("备注")) = NVL(mrsBalance!摘要)
            .TextMatrix(i, .ColIndex("卡号")) = NVL(mrsBalance!卡号)
            .TextMatrix(i, .ColIndex("交易流水号")) = NVL(mrsBalance!交易流水号)
            .TextMatrix(i, .ColIndex("交易说明")) = NVL(mrsBalance!交易说明)
            .TextMatrix(i, .ColIndex("卡类别名称")) = NVL(mrsBalance!卡类别名称)
            .Row = i: .Col = .ColIndex("支付方式")
            .CellBackColor = &HE7CFBA
            mrsBalance.MoveNext
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsBlance.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub IDKindPaymentsType_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnNotChange = True Then Exit Sub
    Call LoadDefaultMoney
    If txtBalance(Idx_缴款).Enabled And txtBalance(Idx_缴款).Visible _
        And picCurBalance.Enabled And picBalance.Enabled Then txtBalance(Idx_缴款).SetFocus
    mblnNotChange = True
    Call LoadCurOwnerPayInfor
    mblnNotChange = False
End Sub
Private Sub IDKindPaymentsType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub IDKindPaymentsType_KeyPress(KeyAscii As Integer)
    Call MoveIDKindItem(IDKindPaymentsType, KeyAscii)
End Sub
Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBlance.Width = .ScaleWidth - vsBlance.Left * 2
        If mbytInState = 1 Then
            vsBlance.Top = .ScaleTop + 50
            cmdCancel.Top = .ScaleHeight - cmdCancel.Height - 50
            fraSplit.Top = cmdCancel.Top - cmdCancel.Height - 50
            vsBlance.Height = fraSplit.Top - vsBlance.Top - 50
        Else
            vsBlance.Top = cmdOK.Top + cmdOK.Height + 100
            vsBlance.Height = .ScaleHeight - vsBlance.Top - 50
            cmdCancel.Top = cmdOK.Top
        End If
    End With
End Sub

 
Private Sub picRight_Resize()
    Err = 0: On Error Resume Next
    With picRight
        picBalance.Left = .ScaleLeft + 50
        picBalance.Width = .ScaleWidth - picBalance.Left * 2
        If mbytInState = 1 Then
            picBalance.Top = fraLeft.Top + txtUnit.Top + txtUnit.Height + 100
        Else
            picBalance.Top = fraLeft.Top + fraLeft.Height + 50
        End If
        picBalance.Height = .ScaleHeight - picBalance.Top - 100
    End With
    If mbytInState <> 1 Then
        zlRaisEffect picBalance, Dw_SubKen
    End If
End Sub

Private Sub vsBlance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "结算列表"
End Sub
 Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "结算列表"
End Sub
 
Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub
Private Sub vsBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblCurMoney As Double, dbl个人帐户 As Double, dbl医保基金 As Double
    With vsBlance
        Select Case Col
        Case .ColIndex("支付方式")
        Case .ColIndex("金额")
        Case Else
        End Select
    End With
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytInState = 1 Then Cancel = True: Exit Sub
    With vsBlance
        .ComboList = ""
        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
        Select Case Val(.TextMatrix(Row, .ColIndex("编辑状态")))
        Case 0: Cancel = True: Exit Sub
        Case 1
            If Col <> .ColKey("金额") Then Cancel = True: Exit Sub
        Case 2
            If Col = .ColIndex("支付方式") Then
                 .ComboList = "..."
                 .CellButtonPicture = imgDel
            Else
                Cancel = True: Exit Sub
            End If
        End Select
    End With
End Sub
Private Sub vsPati_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsPati, Me.Name, "病人列表"
End Sub
Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If mbytInState = 1 Then Exit Sub
    Call DeletePayInfor(Row)
End Sub

Private Sub vsBlance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If mbytInState = 1 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("金额") Then Exit Sub
    End With
    Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m金额式)
End Sub

Private Sub vsBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyDelete Then Exit Sub
    With vsBlance
        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
        If Val(.TextMatrix(.Row, .ColIndex("编辑状态"))) = 2 Then
            Call DeletePayInfor(.Row)
        End If
    End With
End Sub

Private Sub vsBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim dbl原始金额 As Currency
    Dim dblTotal As Double, arrValue As Variant
    Dim i As Integer, str结算方式 As String
    Dim varData As Variant
     
    If mbytInState = 1 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("金额") Then Exit Sub
        If Row <= 0 Then Exit Sub
        Cancel = True
    End With
End Sub
Private Sub DeletePayInfor(ByVal lngDelRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除支付信息
    '编制:刘兴洪
    '日期:2015-01-28 15:18:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngRow As Long
    On Error GoTo errHandle
    With vsBlance
    
        dblMoney = Val(.TextMatrix(lngDelRow, .ColIndex("金额")))
        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
        If Val(.TextMatrix(lngDelRow, .ColIndex("编辑状态"))) <> 2 Then Exit Sub
        
        lngRow = lngDelRow
        mtyBalanceInfor.dbl未付合计 = FormatEx(mtyBalanceInfor.dbl未付合计 + dblMoney, 6)
        mtyBalanceInfor.dbl已付合计 = FormatEx(mtyBalanceInfor.dbl已付合计 - dblMoney, 6)
        
        Call LoadCurOwnerPayInfor
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            vsBlance.RemoveItem lngDelRow
        End If
        
        If lngRow <= 1 Then
            lngRow = 1
        ElseIf lngRow >= .Rows - 1 Then
            lngRow = .Rows - 1
        Else
            lngRow = lngDelRow + 1
        End If
        If lngRow > .Rows - 1 Or lngRow <= 1 Then lngRow = 1
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then .ShowCell .Row, .Col
        Call LoadCurOwnerPayInfor
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub vsBlance_DblClick()
    If mbytInState = 1 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("金额") Then Exit Sub
        '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
        If Val(.TextMatrix(.Row, .ColIndex("编辑状态"))) <> 1 Then Exit Sub
        .EditCell
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsPati_DblClick()
    Dim lng病人ID As Long
    
    lng病人ID = Val(vsPati.RowData(vsPati.Row))
    If lng病人ID <> 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me, _
            "病人ID=" & lng病人ID, "结帐ID=" & mlng结帐ID, "ReportFormat=" & mbytInState + 1, 0)
    End If
End Sub
    

Private Sub txtUnit_GotFocus()
    Call OpenIme(gstrIme)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If cmd病人.Visible Then Call cmd病人_Click
    End If
End Sub

Private Sub txtUnit_LostFocus()
    Call OpenIme
End Sub
Private Function Init结算方式() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算信息
    '编制:刘兴洪
    '日期:2015-01-08 12:06:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim objCards As Cards, objCard As Card
    Dim objPayCards As Cards, i As Long
    
    On Error GoTo errHandle
    mstrErrorBalance = ""
    If mbytInState = 1 Then Init结算方式 = True: Exit Function
    
    
    Set objCards = New Cards: Set objPayCards = New Cards
    '性质:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项, _
    '     6-费用折扣,7-一卡通结算,8-结算卡结算
 
    Set mrs结算方式 = GetPayKind
    If mrs结算方式.RecordCount = 0 Then
        MsgBox "未设置结帐场合可用的结算方式。", vbInformation, gstrSysName
        mblnUnload = True
        Exit Function
    End If
     
     
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare Is Nothing Then
        '0-所有医疗卡;1-启用的医疗卡,2-所有存在三方账户的三方卡 3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)    '获取有效的三方帐户会付
    End If

    mrs结算方式.Filter = "性质<7 or 性质=9"
    With mrs结算方式
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
        Do While Not .EOF
            If InStr(",3,4,9,", "," & Val(NVL(!性质)) & ",") = 0 Then
                Set objCard = New Card
                objCard.接口序号 = -1 * i
                objCard.接口编码 = !编码
                objCard.名称 = !名称
                objCard.结算方式 = !名称
                objCard.结算性质 = Val(NVL(!性质))
                objCard.启用 = True
                '85565,李南春,2015/7/19:读卡性质
                objCard.是否刷卡 = True
                objCard.缺省标志 = Val(NVL(!缺省)) = 1
                objPayCards.Add objCard
                If objCard.缺省标志 Then
                    mstr缺省结算方式 = objCard.结算方式
                End If
                i = i + 1
            ElseIf Val(NVL(!性质)) = 9 Then
                mstrErrorBalance = NVL(!名称)
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    mrs结算方式.Filter = "性质>=7 and 性质<9" '一卡通结算
    With mrs结算方式
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            For Each objCard In objCards
                If objCard.结算方式 = NVL(!名称) Then
                    '找到了,增加
                    '85565,李南春,2015/7/19:读卡性质
                    objCard.是否刷卡 = True
                    objCard.缺省标志 = Val(NVL(!缺省)) = 1
                    objPayCards.Add objCard
                    If objCard.缺省标志 Then
                        mstr缺省结算方式 = objCard.结算方式
                    End If
                End If
            Next
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    mrs结算方式.Filter = 0
    mblnNotChange = True
    Set IDKindPaymentsType.Cards = objPayCards
    If objPayCards.Count = 0 Then
        mblnNotChange = True
        MsgBox "结帐场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnload = True: Exit Function
    End If
    mblnNotChange = False
    Init结算方式 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub NewBalance()

    Dim tyBalance As TY_Balance_Infor
    mtyBalanceInfor = tyBalance
    
    mstrDec = gstrDec
    Unload frmPatientsSelect
    
    Call InitPatiGrid
    Call InitGrid_PayList
 
    txtUnit.Tag = ""
    txtUnit.Text = ""
    txtUnit.Enabled = True
    Set mrsPatients = Nothing
    
    lbl自付合计.Caption = Format(0, mstrDec)
    lbl剩余自付.Caption = Format(0, mstrDec)
    txtBalance(Idx_lbl缴款).Text = ""
    txtBalance(Idx_lbl找补).Text = ""
    txtBalance(Idx_结算号码).Text = ""
    txtBalance(Idx_摘要).Text = ""
    
    Call ClearPayInfo
    staThis.Panels(2).Text = ""
    

    '票据号与单据号处理
    Call RefreshFact
    If Visible And txtUnit.Enabled Then txtUnit.SetFocus
End Sub
 
Private Sub InitPatiGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算列头信息
    '编制:刘兴洪
    '日期:2015-05-04 17:33:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsPati
        .Redraw = flexRDNone: .Clear
        .Rows = 2: .Cols = 5: i = 0
        .TextMatrix(0, i) = "病人ID": i = i + 1
        .TextMatrix(0, i) = "姓名": i = i + 1
        .TextMatrix(0, i) = "性别": i = i + 1
        .TextMatrix(0, i) = "年龄": i = i + 1
        .TextMatrix(0, i) = "结帐金额": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*金额" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "性别" Or .ColKey(i) = "年龄" Then
                .ColAlignment(i) = flexAlignCenterCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(0) = 0
            End If
        Next
        .ColWidth(.ColIndex("姓名")) = 1000
        .ColWidth(.ColIndex("性别")) = 650
        .ColWidth(.ColIndex("年龄")) = 650
        .ColWidth(.ColIndex("结帐金额")) = 1450
        .RowHeight(0) = 320
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsPati, Me.Name, "病人列表"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 


Private Sub ClearPayInfo()
    Dim i As Long
    With vsBlance
        .Rows = 2
        .Clear 1
    End With
    txtBalance(Idx_缴款).Text = Format(0, mstrDec)
    txtBalance(Idx_找补).Text = Format(0, mstrDec)
End Sub
Private Sub ShowBalanceInfo(ByVal rsTmp As ADODB.Recordset)
   Dim curTotal As Currency, lngMaxLength As Long, lngP As Long, i As Long
    
    Call InitPatiGrid
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.State <> 1 Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    
    lngMaxLength = Len(Mid(gstrDec, 3))
    For i = 1 To rsTmp.RecordCount
        lngP = InStr(1, CStr(rsTmp!结帐金额), ".")
        If lngP > 0 Then
            lngP = Len(Mid(CStr(rsTmp!结帐金额), lngP + 1))
            If lngP > lngMaxLength Then lngMaxLength = lngP
        End If
        rsTmp.MoveNext
    Next
    
    mstrDec = "0." & String(lngMaxLength, "0")
    With vsPati
        .Redraw = flexRDNone
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        If .Rows = 1 Then .Rows = 2
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            .RowData(i) = Val(rsTmp!病人ID)
            .TextMatrix(i, .ColIndex("病人ID")) = Val(NVL(rsTmp!病人ID))
            .TextMatrix(i, .ColIndex("姓名")) = "" & rsTmp!姓名
            .TextMatrix(i, .ColIndex("性别")) = "" & rsTmp!性别
            .TextMatrix(i, .ColIndex("年龄")) = "" & rsTmp!年龄
            .TextMatrix(i, .ColIndex("结帐金额")) = Format(rsTmp!结帐金额, mstrDec)
            .Row = i: .Col = .ColIndex("结帐金额")
            .CellBackColor = 12900351
            curTotal = curTotal + rsTmp!结帐金额
            rsTmp.MoveNext
        Next
        vsPati.Redraw = flexRDBuffered
    End With
    lbl自付合计.Caption = Format(curTotal, mstrDec)
    With mtyBalanceInfor
        .dbl当前结帐 = curTotal
        .dbl未付合计 = curTotal
    End With
End Sub

Private Function GetPayKind() As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "" & _
    "   Select a.编码,a.名称, a.缺省标志 缺省, a.性质, 1 As 位置" & vbNewLine & _
    "   From 结算方式 a, 结算方式应用 b" & vbNewLine & _
    "   Where a.名称 = b.结算方式 And b.应用场合 = '结帐' And a.性质 Not In (3, 4)  " & _
    "       And Nvl(a.应付款, 0) = 0 And Nvl(a.应收款, 0) = 0" & vbNewLine & _
    "   Union " & _
    "   Select 编码,名称, 缺省标志 As 缺省, 性质,0 As 位置" & _
    "   From 结算方式 " & _
    "   Where 性质=9 " & _
    "Order By 位置,性质,编码"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set GetPayKind = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetBalanceData(ByVal lng结帐ID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strWhere As String
    
    strWhere = " And B.记录状态 In(1,3)"
    If mblnViewCancel Then strWhere = " And B.记录状态 =2"
    
    strSql = "" & _
    " Select A.病人ID,nvl(A.姓名,C.姓名) as 姓名, nvl(A.性别,C.性别) as 性别,  " & _
    "       nvl(A.年龄,C.年龄) as 年龄, Nvl(Sum(A.结帐金额),0) 结帐金额" & vbNewLine & _
    " From 门诊费用记录 A,病人结帐记录 B,病人信息 C" & vbNewLine & _
    " Where A.结帐id = [1] And A.结帐id = B.id " & strWhere & _
    "        And A.病人id=C.病人id " & vbNewLine & _
    " Group By A.病人ID,nvl(A.姓名,C.姓名), nvl(A.性别,C.性别), nvl(A.年龄,C.年龄) "
    
    If blnHistory Then strSql = Replace(Replace(strSql, "门诊费用记录", "H门诊费用记录"), "病人结帐记录", "H病人结帐记录")
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    Set GetBalanceData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPayData(ByVal lng结帐ID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strWhere As String
    strWhere = " And B.记录状态 In(1,3)"
    If mblnViewCancel Then strWhere = " And B.记录状态 =2"

    strSql = "" & _
    "   Select A.结算方式, A.冲预交 结算金额, A.结算号码, A.摘要 备注, " & vbNewLine & _
    "           A.卡类别ID,A.结算卡序号,A.卡号,A.交易流水号,A.交易说明" & vbNewLine & _
    "   From 病人预交记录 A,病人结帐记录 B" & vbNewLine & _
    "   Where A.结帐id = [1] And A.结帐id = B.id " & strWhere
    If blnHistory Then strSql = Replace(Replace(strSql, "病人预交记录", "H病人预交记录"), "病人结帐记录", "H病人结帐记录")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    Set GetPayData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetOtherInfo(ByVal lng结帐ID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String, strWhere As String
    Dim strTable As String, strTable1 As String, strTable2 As String
   
    strWhere = " And B.记录状态 In(1,3)"
    If mblnViewCancel Then strWhere = " And B.记录状态 =2"
    
    strTable1 = "" & _
    "   Select Min(D.名称) 合约单位  " & _
    "   From 门诊费用记录 A,病人结帐记录 B, 病人信息 C, 合约单位 D" & vbNewLine & _
    "   Where A.结帐id = [1] And A.结帐id = B.id  " & strWhere & _
    "         And A.病人id = C.病人id And C.合同单位id = D.ID"
    
    strTable2 = "" & _
    "   Select A.NO, A.实际票号 " & _
    "   From 病人结帐记录 A " & _
    "   Where A.id = [1]  " & Replace(strWhere, "B.", "A.")
    
    strSql = "" & _
    " Select   B.合约单位, C.NO, C.实际票号" & vbNewLine & _
    " From (" & strTable1 & ") B," & vbNewLine & _
    "      (" & strTable2 & ") C"
  
    If blnHistory Then strSql = Replace(Replace(strSql, "门诊费用记录", "H门诊费用记录"), "病人结帐记录", "H病人结帐记录")
    If blnHistory Then strSql = Replace(strSql, "病人预交记录", "H病人预交记录")
 
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    
    Set GetOtherInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitGrid_PayList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化支付列表
    '编制:刘兴洪
    '日期:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
        .Clear: .Rows = 2: i = 0: .Cols = 18
        .TextMatrix(0, i) = "卡类别ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "消费卡ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算性质": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "编辑状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "类型": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否全退": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "校对标志": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否密文": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "支付方式": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "金额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "结算号码": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "卡号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易流水号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易说明": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "备注": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "卡类别名称": .ColWidth(i) = 0: i = i + 1
        
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            Select Case .ColKey(i)
            Case "结算性质", "类型", "是否保存", "是否密文", "校对标志", "编辑状态", "是否退现", "是否全退", "卡类别名称", "结算状态", "是否验证"
                .ColHidden(i) = True
            Case "金额"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
        zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "结算列表"
        If mbytInState = 0 Then '结帐操作
            .Editable = flexEDKbdMouse
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
  


Private Sub txtBalance_Change(Index As Integer)
    If mblnNotChange Then Exit Sub
    
    If mbytInState = 1 Then Exit Sub
    Select Case Index
    Case Idx_缴款
         Call Set找补信息
    Case Idx_摘要
    Case Else
    End Select
End Sub

Private Sub txtBalance_GotFocus(Index As Integer)
    Select Case Index
    Case Idx_缴款
      '  Call LedVoiceSpeak(True)
      '  txtBalance(Index).Text = ""
        
    Case Idx_摘要
        zlCommFun.OpenIme True
    End Select
    zlControl.TxtSelAll txtBalance(Index)
End Sub
Private Sub txtBalance_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim dblMoney As Double, blnChargeEnd As Boolean
    Dim objCard As Card, objKind As IDKindNew
 
    If KeyAscii <> 13 Then
       If Index <> Idx_缴款 Then Exit Sub
       Set objKind = IDKindPaymentsType
       Call MoveIDKindItem(objKind, KeyAscii)
       Exit Sub
    End If
    
    KeyAscii = 0
    Select Case Index
    Case Idx_缴款
        dblMoney = FormatEx(Val(txtBalance(Index).Text), 6)
        Set objCard = IDKindPaymentsType.GetCurCard
        If objCard Is Nothing Then zlCommFun.PressKey vbKeyTab: Exit Sub
        Select Case mty_ModulePara.byt缴款输入控制
        Case 2   '按病人缴款累计
            If objCard.结算性质 = 1 Then '现金
                If txtBalance(Index).Text = "" Then
                    cmdOK.SetFocus: Exit Sub
                End If
                If SaveBalanceData = False Then Exit Sub
                
            ElseIf objCard.结算性质 = 2 Then '非医保结算
                If txtBalance(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
            ElseIf objCard.接口序号 > 0 Then
                If dblMoney = 0 Then
                    MsgBox "未输入缴款金额,不能使用『" & objCard.结算方式 & "』进行结算", vbInformation, gstrSysName
                    Exit Sub
                End If
                If SaveBalanceData = False Then Exit Sub
            End If
            Exit Sub
        Case Else '0-不进行缴款控制,1-输入现金时,必须输入缴款金额
            If objCard.结算性质 = 1 Then Call cmdOK_Click: Exit Sub
        End Select
        
        If objCard.接口序号 > 0 Then
            If dblMoney = 0 Then
                MsgBox "未输入缴款金额,不能使用『" & objCard.结算方式 & "』进行结算", vbInformation, gstrSysName
                Exit Sub
            End If
            If SaveBalanceData = False Then Exit Sub
            Exit Sub
        End If
        If dblMoney <> 0 Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtBalance_LostFocus(Index As Integer)
    Select Case Index
    Case Idx_缴款
        txtBalance(Index).Text = Format(Val(txtBalance(Index).Text), "0.00")
    Case Idx_摘要
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtBalance_Validate(Index As Integer, Cancel As Boolean)
    Dim dblMoney As Double, dbl找补 As Double
    Dim intSign As Integer
    Select Case Index
    Case Idx_缴款
    Case Else
    End Select
End Sub
Private Sub Set找补信息()
    Dim dblMoney As Double
    Dim dbl当前未付 As Double
    Dim objCard As Card
    Dim objBackCard As Card
    Dim objCards As Cards
    Dim objTemp As Card
    dblMoney = Val(txtBalance(Idx_缴款).Text)
    Set objCard = IDKindPaymentsType.GetCurCard
    
    dbl当前未付 = mtyBalanceInfor.dbl当前结帐 - mtyBalanceInfor.dbl已付合计
    If dblMoney = 0 Or objCard Is Nothing Then
        txtBalance(Idx_找补).ForeColor = txtBalance(Idx_结算号码).ForeColor
        txtBalance(Idx_找补).Text = "0.00"
        Exit Sub
    End If
        
    If dbl当前未付 < 0 Then
        '当前状态为退款
        dbl当前未付 = FormatEx(Val(lbl剩余自付.Caption), 6)
        If objCard.结算性质 = 1 Then
            '只有现金,才会在退款时多付给病人,比如:退100给病人,要找回50
            If Abs(dbl当前未付) <= dblMoney Then
                txtBalance(Idx_找补).Text = Format(dblMoney - Abs(dbl当前未付), "0.00")
                lblBalance(Idx_lbl找补).Caption = "收    款"
                txtBalance(Idx_找补).ForeColor = vbRed
            Else
                txtBalance(Idx_找补).Text = Format(Abs(dblMoney - Abs(dbl当前未付)), "0.00")
                lblBalance(Idx_lbl找补).Caption = "退    款"
                txtBalance(Idx_找补).ForeColor = txtBalance(Idx_结算号码).ForeColor
            End If
            Exit Sub
        End If
        
        If Abs(dbl当前未付) < dblMoney Then
            '其他结算方式的,只能退剩余未退款,比如:退款是医院开支票给病人,因此不能再找回支票的可能
            lblBalance(Idx_lbl找补).Caption = "收    款"
            txtBalance(Idx_找补).ForeColor = txtBalance(Idx_结算号码).ForeColor
            txtBalance(Idx_找补).Text = "0.00": Exit Sub
        Else
            txtBalance(Idx_找补).Text = Format(Abs(dblMoney - Abs(dbl当前未付)), "0.00")
            lblBalance(Idx_lbl找补).Caption = "退    款"
            txtBalance(Idx_找补).ForeColor = txtBalance(Idx_结算号码).ForeColor
        End If
        Exit Sub
    End If
    
    '当前状态为收款
    dbl当前未付 = FormatEx(Val(lbl剩余自付.Caption), 6)
    If objCard.结算性质 = 1 Then
        '只有现金,才会在退款时多付给病人,比如:退100给病人,要找回50
        If dbl当前未付 >= dblMoney Then
            '还要收取病人钱
            txtBalance(Idx_找补).Text = Format(Abs(dblMoney - Abs(dbl当前未付)), "0.00")
            lblBalance(Idx_lbl找补).Caption = "收    款"
            txtBalance(Idx_找补).ForeColor = vbRed
        Else
            '退款
            txtBalance(Idx_找补).Text = Format(Abs(dblMoney - Abs(dbl当前未付)), "0.00")
            lblBalance(Idx_lbl找补).Caption = "退    款"
            txtBalance(Idx_找补).ForeColor = txtBalance(Idx_结算号码).ForeColor
        End If
        Exit Sub
    End If
    
    If dbl当前未付 >= dblMoney Then
        '要收款
        lblBalance(Idx_lbl找补).Caption = "收    款"
        txtBalance(Idx_找补).Text = Format(Abs(dblMoney - Abs(dbl当前未付)), "0.00")
        txtBalance(Idx_找补).ForeColor = vbRed
        Exit Sub
    Else
        If objCard.结算性质 = 2 And objCard.结算方式 Like "*支票" Then
            lblBalance(Idx_lbl找补).Caption = "退 支 票"
        Else
            lblBalance(Idx_lbl找补).Caption = "退    款"
        End If
        txtBalance(Idx_找补).Text = Format(Abs(dblMoney - Abs(dbl当前未付)), "0.00")
        txtBalance(Idx_找补).ForeColor = txtBalance(Idx_结算号码).ForeColor
    End If
End Sub

Private Function zlCheckMulitInterfaceNumValied(Optional bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是正同时存在两种以上接口(不含两种)
    '返回:不含两种以上接口的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int性质 As Integer, str结算方式 As String
    Dim varData As Variant, strErrMsg As String
    Dim objCard As Card
    Dim intMousePointer As Integer
    On Error GoTo errHandle
    strErrMsg = ""
    intMousePointer = Screen.MousePointer
    Set objCard = IDKindPaymentsType.GetCurCard
        
    If bln预交 Or objCard.接口序号 <= 0 Then zlCheckMulitInterfaceNumValied = True:        Exit Function
    
   With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If InStr("34", int性质) > 0 Then
                intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str结算方式 & ":" & .TextMatrix(i, .ColIndex("金额"))
            End If
        Next
    End With
    If intCount > 1 Then
        Screen.MousePointer = 0
        Call MsgBox("注意:" & vbCrLf & "   本系统目前只支持一种接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function

Private Function CheckSquareBalanceValied(ByVal objCard As Card, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡结算交易检查
    '入参:objCard-三方卡
    '出参:dblMoney-当前刷卡金额
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl帐户余额 As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln退现 As Boolean, dbl未付金额 As Double
    Dim intMousePointer As Integer, strXmlIn As String
    Dim lng消费卡ID As Long, str卡号 As String, str密码 As String
    Dim str限制类别 As String, byt是否密文   As Byte
    Dim cllBushSquare As Collection, i As Long
    
    
    intMousePointer = Screen.MousePointer

    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then CheckSquareBalanceValied = True: Exit Function
    
    On Error GoTo errHandle
    
    tyBrushCard = strBrushCard
    
    dblMoney = Val(txtBalance(Idx_缴款).Text)
    dbl未付金额 = FormatEx(mtyBalanceInfor.dbl未付合计, 6)
    
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "收款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If dblMoney > Format(dbl未付金额, "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox "收款金额不能大于本次未付金额:" & Format(dbl未付金额, "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '先检查对应的接口
    If zlGetClassMoney(0, rsMoney) = False Then Exit Function
    
     '构建消费卡的刷卡信息
     Set cllSquareBalance = New Collection
     Set mcllCurSquareBalance = New Collection
     With vsBlance
        For i = 1 To .Rows - 1
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            '编辑状态:0-禁止删除;1-允许编辑金额;2-允许删除
            '结算状态:是否已结算:1-已结算;0-未结算
            If Val(.TextMatrix(i, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = objCard.接口序号 _
                And Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
  
                dblTemp = FormatEx(Val(.TextMatrix(i, .ColIndex("金额"))), 6)
                lng消费卡ID = Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                str卡号 = Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
                str密码 = Trim(.Cell(flexcpData, i, .ColIndex("消费卡ID")))  '密码
                str限制类别 = Trim(.Cell(flexcpData, i, .ColIndex("卡类别ID")))  '限制类别
                byt是否密文 = Val(.TextMatrix(i, .ColIndex("是否密文")))
                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                cllSquareBalance.Add Array(objCard.接口序号, lng消费卡ID, dblTemp, str卡号, str密码, str限制类别, byt是否密文)
            End If
        Next
     End With
     For i = 1 To cllSquareBalance.Count
        mcllCurSquareBalance.Add cllSquareBalance(i)
     Next
     
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
        Optional ByRef varSquareBalance As Variant) As Boolean
    'varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    strXmlIn = "<IN><CZLX>0</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, _
            objCard.接口序号, objCard.消费卡, _
            "" & txtUnit.Text, "" & "", "" & "", dblMoney, _
            tyBrushCard.str卡号, tyBrushCard.str密码, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
       
        For i = 1 To cllSquareBalance.Count
           mcllCurSquareBalance.Add cllSquareBalance(i)
        Next
         
        '保存前,一些数据检查
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNOs As String, _
        Optional ByVal strXMLExpend As String
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.接口序号, _
            objCard.消费卡, tyBrushCard.str卡号, dblMoney, "", strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
        '入参:frmMain-调用的主窗体
        '        lngModule-模块号
        '        strCardNo-卡号
        '        strExpand-预留，为空,以后扩展
        '出参:dblMoney-返回帐户余额
        'If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
              tyBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡) = False Then Exit Function
        '已经更改了结算金额
        If FormatEx(dblMoney, 6) <> Val(txtBalance(Idx_缴款).Text) Then
            txtBalance(Idx_缴款).Text = Format(dblMoney, "0.00")
        End If
        CheckSquareBalanceValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckThreeSwapValied(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易验证
    '入参:objCard-三方卡
    '     dblMoney-刷卡金额,>=0表示收款;小于零表示退款
    '     bln退款-true,表示当前为退款检查;False表示当前为收款检查
    '出参:tyBrushCard-刷卡信息
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, cllSquareBalance As Collection
    Dim strXMLExpend As String, bln退现 As Boolean
    Dim dbl帐户余额 As Double, dbl未付金额 As Double
    Dim strExpand As String, strXmlIn As String
    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim blnCurInput As Boolean
    
    intMousePointer = Screen.MousePointer
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then CheckThreeSwapValied = True: Exit Function
    
    
    On Error GoTo errHandle
    
    tyBrushCard.bln转帐 = False
    If dblMoney = 0 Then dblMoney = Val(txtBalance(Idx_缴款).Text): blnCurInput = True
     
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "收款") & "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
     
    dbl未付金额 = FormatEx(mtyBalanceInfor.dbl未付合计, 6)
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 And blnCurInput = False Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "刷卡") & "金额不能大于本次" & IIf(bln退款, "未退", "未付") & "金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Val(lbl剩余自付.Caption) <> dblMoney Then
        Screen.MousePointer = 0
        MsgBox IIf(bln退款, "退款", "刷卡") & "金额不足:" & Format(Abs(Val(lbl剩余自付.Caption)), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not bln退款 Then
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
           Optional ByRef varSquareBalance As Variant) As Boolean
           '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            txtUnit.Text, "", "", dblMoney, _
            tyBrushCard.str卡号, tyBrushCard.str密码, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
            '保存前,一些数据检查
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal strCardTypeID As Long, ByVal strCardNo As String, _
            ByVal dblMoney As Double, ByVal strNOs As String, _
            Optional ByVal strXMLExpend As String
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.接口序号, _
            objCard.消费卡, tyBrushCard.str卡号, dblMoney, "", strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
          '入参:frmMain-调用的主窗体
          '        lngModule-模块号
          '        strCardNo-卡号
          '        strExpand-预留，为空,以后扩展
          '出参:dblMoney-返回帐户余额
        If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
              tyBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡) = False Then Exit Function
        
        staThis.Panels(2).Text = Format(dbl帐户余额, "0.00")
        staThis.Panels(2).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
        tyBrushCard.dbl帐户余额 = FormatEx(dbl帐户余额, 2)
        If dbl帐户余额 <> 0 And dbl帐户余额 < dblMoney Then
            Screen.MousePointer = 0
            MsgBox objCard.结算方式 & "的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        CheckThreeSwapValied = True
        Exit Function
    End If
    
    '退款检查
    If mrsBalance Is Nothing Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    

    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    mrsBalance.Filter = "类型=3 And 卡类别ID=" & objCard.接口序号
    If mrsBalance.EOF Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.结算方式 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    dblTemp = 0
    With mrsBalance
        Do While Not .EOF
            dblTemp = dblTemp + Val(NVL(!冲预交))
            .MoveNext
        Loop
        mrsBalance.MoveFirst
        dblTemp = FormatEx(dblTemp, 5)
    End With
    
    If dblTemp = 0 Then
        If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.结算方式 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objCard.是否全退 And Not objCard.是否退现 Then
        If dblTemp <> dblMoney Then
            If objCard.是否转帐及代扣 Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "注意:" & vbCrLf & objCard.名称 & "进行退款时，必须全退！" & vbCrLf & _
            "  剩余未退:" & Format(Abs(dblTemp), "0.00") & vbCrLf & _
            "  当前金额:" & Format(Abs(dblMoney), "0.00"), vbInformation + vbOKOnly, gstrSysName
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
        '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       dblMoney-退款金额
        '       strSwapNo-交易流水号(退款时检查)
        '       strSwapMemo-交易说明(退款时传入)
        '       strXMLExpend    XML IN  可选参数:异常单据重新退费(1)
        '返回:退款合法,返回true,否则返回Flase
        
    strXMLExpend = ""
    tyBrushCard.str卡号 = NVL(mrsBalance!卡号)
    tyBrushCard.str交易流水号 = NVL(mrsBalance!交易流水号)
    tyBrushCard.str交易说明 = NVL(mrsBalance!交易说明)

    strBalanceIDs = "2|" & mtyBalanceInfor.lng结帐ID & IIf(mtyBalanceInfor.lng冲销ID = 0, "", "," & mtyBalanceInfor.lng冲销ID)
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, _
        strBalanceIDs, dblMoney, tyBrushCard.str交易流水号, tyBrushCard.str交易说明, strXMLExpend) = False Then Exit Function
                
    If objCard.是否退款验卡 Then
       '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        'ByVal lngModule As Long, _
        'ByVal rsClassMoney As ADODB.Recordset, _
        'ByVal lngCardTypeID As Long, _
        'ByVal bln消费卡 As Boolean, _
        'ByVal strPatiName As String, ByVal strSex As String, _
        'ByVal strOld As String, ByVal dbl金额 As Double, _
        'Optional ByRef strCardNo As String, _
        'Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean) As Boolean
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.接口序号, _
            objCard.消费卡, txtUnit.Text, "", _
            "", dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
            True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    End If
    CheckThreeSwapValied = True
    Exit Function
    
GoTransferAccount:
        strXmlIn = "<IN><CZLX>1</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.接口序号, _
            objCard.消费卡, txtUnit.Text, "", _
            "", dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
            True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    
    tyBrushCard.bln转帐 = True
    '调用转帐接口
    '    7.1.    zltransferAccountsCheck(转帐检查接口)
    'zlTransferAccountsCheck 转帐检查接口
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  HIS调用模块号
    'lngCardTypeID   Long    In  卡类别ID
    'strCardNo   String  In  卡号
    'dblMoney    Double  In  转帐金额(代扣时为负数)
    'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
    'strXMLExpend String In   XML串:
    '                            <IN>
    '                                <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务
    '                            </IN>
    '                    Out  XML串:
    '                            <OUT>
    '                               <ERRMSG>错误信息</ERRMSG >
    '                            </OUT>
    '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
    '说明:
    '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '构造XML串
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.接口序号, _
        tyBrushCard.str卡号, dblMoney, mtyBalanceInfor.lng结帐ID, strXMLExpend) = False Then
        Screen.MousePointer = 0
        Call zlShowThreeSwapErrInfor(0, strXMLExpend)
        Exit Function
    End If
    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
          tyBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡)
    If dbl帐户余额 <> 0 Then
        staThis.Panels(2).Text = objCard.结算方式 & "帐户余额:" & Format(dbl帐户余额, "0.00")
        staThis.Panels(2).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
    End If
    tyBrushCard.dbl帐户余额 = FormatEx(dbl帐户余额, 2)
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function



Private Function zlGetClassMoney(ByRef lng结帐ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, dblMoney As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    '初始化数据结构
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "金额", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng结帐ID <> 0 Then
        strSql = "" & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 门诊费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 " & _
        "   Union ALL " & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 住院费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 "
        strSql = "Select 收费类别,Sum(金额) as 金额 From (" & strSql & ")  Group by  收费类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "收费类别='" & NVL(!收费类别, "无") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!收费类别 = NVL(!收费类别, "无")
                rsMoney!金额 = Val(NVL(rsMoney!金额)) + Val(NVL(!金额))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        zlGetClassMoney = True
        Exit Function
    End If
 
    strSql = "" & _
    " Select A.收费类别, Nvl(Sum(实收金额), 0) - Nvl(Sum(结帐金额), 0) As 未结金额" & vbNewLine & _
    " From 门诊费用记录 A, 病人信息 B" & vbNewLine & _
    " Where A.病人id = B.病人id   And A.记录状态 <> 0 And A.记帐费用 = 1  " & _
    "       And A.门诊标志 IN(1,4) And B.合同单位id = [1] And B.当前科室id Is Null " & _
    " Group By a.收费类别" & _
    " Having Nvl(Sum(a.实收金额), 0) - Nvl(Sum(a.结帐金额), 0) <> 0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(txtUnit.Tag))
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            rsMoney.Find "收费类别='" & NVL(!收费类别, "无") & "'", , adSearchForward, 1
            If rsMoney.EOF Then rsMoney.AddNew
            rsMoney!收费类别 = NVL(!收费类别, "无")
            rsMoney!金额 = Val(NVL(rsMoney!金额)) + Val(NVL(!未结金额))
            rsMoney.Update
            .MoveNext
        Loop
    End With
    
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitOldOneCardInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化老一卡通信息
    '编制:刘兴洪
    '日期:2015-01-08 12:02:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbytInState = 1 Then Exit Sub
    Set mOldOneCard.rsOneCard = GetOneCard
    With mOldOneCard
        .blnOneCard = .rsOneCard.RecordCount > 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetCentMoney(ByVal dblMoney As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据分币处理规则,返回分币处理后的金额
    '入参:dblMoney-未处理的原始金额
    '返回:返回分币处理后的金额
    '编制:刘兴洪
    '日期:2015-01-26 10:57:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then GetCentMoney = FormatEx(dblMoney, 2): Exit Function
    '非现金的,保留两位小数
    If objCard.结算性质 <> 1 Then GetCentMoney = FormatEx(dblMoney, 2): Exit Function
    GetCentMoney = CentMoney(CCur(dblMoney))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub Show误差金额()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示误差金额
    '编制:刘兴洪
    '日期:2015-01-14 11:33:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl退支付额 As Double
    Dim dbl剩余金额 As Double, dblTemp As Double, dbl未付款 As Double
    Dim intSign As Integer, objCard As Card
    If mbytInState = 1 Then Exit Sub
    
    With mtyBalanceInfor
        .dbl误差额 = 0
        dbl未付款 = .dbl未付合计
        intSign = IIf(dbl未付款 < 0, -1, 1)
    End With
    
    dblMoney = FormatEx(intSign * Val(txtBalance(Idx_缴款).Text), 6)
    
    dbl退支付额 = 0: dbl剩余金额 = FormatEx(dbl未付款 - dblMoney, 6)
     
    Set objCard = IDKindPaymentsType.GetCurCard
    If Not objCard Is Nothing Then
        If objCard.结算性质 = 1 Then
            dblTemp = dbl未付款: dbl剩余金额 = 0
            dblMoney = GetCentMoney(dblTemp)
            mtyBalanceInfor.dbl误差额 = FormatEx(dbl未付款 - dblMoney, 6)
            GoTo Show误差:
        End If
    End If
    mtyBalanceInfor.dbl误差额 = FormatEx(dbl未付款 - FormatEx(dbl未付款, 2), 6): GoTo Show误差:
Show误差:
    lbl误差额.Visible = mtyBalanceInfor.dbl误差额 <> 0 And mbytInState = 0
    lbl误差额.Caption = "误差:" & FormatEx(mtyBalanceInfor.dbl误差额, 6)
End Sub


Private Function ExecuteOldOneCardPayInterface(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal objCard As Card, ByVal dblMoney As Double, tyBrushCardInfor As TY_BrushCard, _
    ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(老版本)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次结算金额
    '     TYBrushCardInfor-当前刷卡信息
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 16:14:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl余额 As Double, str医院编码 As String
    Dim i As Long, strSql As String, str结算方式 As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intCardType As Integer, strSwapNO As String
    '非一卡通支付,直接返回
    If objCard.结算性质 <> 7 Then ExecuteOldOneCardPayInterface = True: Exit Function

    mOldOneCard.rsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mOldOneCard.rsOneCard.EOF Then
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        ExecuteOldOneCardPayInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '一卡通结算
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If Not mobjICCard.PaymentSwap(dblMoney, dbl余额, intCardType, Val("" & mOldOneCard.rsOneCard!医院编码), tyBrushCardInfor.str卡号, tyBrushCardInfor.str交易流水号, lng结帐ID, lng病人ID) Then
        gcnOracle.RollbackTrans
        MsgBox objCard.结算方式 & "结算失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    strSql = "Zl_一卡通结算_Update(" & 0 & ",'" & objCard.结算方式 & "','" & tyBrushCardInfor.str卡号 & "','" & intCardType & "','" & strSwapNO & "'," & dbl余额 & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Set cllBillPro = New Collection
    blnTrans = False
    ExecuteOldOneCardPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
 End Function
Private Function ExecuteThreeSwapPayInterface(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, objCard As Card, ByVal dblMoney As Double, _
    ByRef cllBillPro As Collection, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     tyBrushCard-当前刷卡信息
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strSql As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str结算方式  As String
    
    Err = 0: On Error GoTo errHandle:
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
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
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-结算金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    str结帐IDs = lng结帐ID
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, _
         str结帐IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    tyBrushCard.str交易流水号 = strSwapGlideNO
    tyBrushCard.str交易说明 = strSwapMemo
    If objCard.消费卡 = False Then
        Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
    End If
    Call zlAddThreeSwapSQLToCollection(False, str结帐IDs, objCard.接口序号, objCard.消费卡, tyBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    ExecuteThreeSwapPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddSquareBalance(ByVal objCard As Card)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加消费卡结算方式到结算方式列表
    '编制:刘兴洪
    '日期:2015-01-23 15:09:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As New Collection
    Dim j As Integer, dblMoney As Double, strCardNo As String
    
    With vsBlance
      '先清除原始的消费卡部分,再重新退费
        Call ClearSquareBalance(objCard.接口序号)
        Set cllBalance = mcllCurSquareBalance
        For j = 1 To cllBalance.Count
            If objCard.接口序号 = Val(cllBalance(j)(0)) Then
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                dblMoney = cllBalance(j)(2)
            
                .TextMatrix(1, .ColIndex("类型")) = 5
                .TextMatrix(1, .ColIndex("是否密文")) = Val(cllBalance(j)(6))
                .TextMatrix(1, .ColIndex("结算性质")) = objCard.结算性质
                .TextMatrix(1, .ColIndex("编辑状态")) = 2   '0-禁止删除;1-允许编辑金额;2-允许删除
                .TextMatrix(1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                .TextMatrix(1, .ColIndex("卡类别ID")) = objCard.接口序号
                .TextMatrix(1, .ColIndex("消费卡ID")) = Val(cllBalance(j)(1))
                .Cell(flexcpData, 1, .ColIndex("消费卡ID")) = cllBalance(j)(4)  '密码
                .Cell(flexcpData, 1, .ColIndex("卡类别ID")) = cllBalance(j)(5)  '限制类别
                
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                 strCardNo = Trim(cllBalance(j)(3))
                .TextMatrix(1, .ColIndex("卡号")) = IIf(objCard.卡号密文规则 <> "" And objCard.卡号密文规则 <> "0", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = strCardNo
                .TextMatrix(1, .ColIndex("金额")) = Format(dblMoney, "0.00")
                .Cell(flexcpData, 1, .ColIndex("金额")) = Format(dblMoney, "0.00")
                .TextMatrix(1, .ColIndex("结算号码")) = ""
                .TextMatrix(1, .ColIndex("备注")) = ""
                .TextMatrix(1, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(1, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(1, .ColIndex("卡类别名称")) = objCard.名称
                
                mtyBalanceInfor.dbl已付合计 = FormatEx(mtyBalanceInfor.dbl已付合计 + dblMoney, 6)
                mtyBalanceInfor.dbl未付合计 = FormatEx(mtyBalanceInfor.dbl未付合计 - dblMoney, 6)
            End If
        Next
    End With
End Sub


Private Sub LoadCurOwnerPayInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载当前结算信息
    '编制:刘兴洪
    '日期:2015-01-12 14:14:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngColor As Long, objCard As Card
    Dim dblTemp As Double
    
    If mbytInState = 1 Then Exit Sub
    
    With mtyBalanceInfor
        dblMoney = .dbl未付合计
        lbl自付合计.Caption = Format(.dbl当前结帐, mstrDec)
        dblTemp = GetCentMoney(Abs(dblMoney))
        
        lbl剩余自付.Caption = Format(FormatEx(dblTemp, 6), mstrDec)
        lbl剩余自付.Tag = Format(dblTemp, mstrDec)
        
    
       stcTittile.Caption = IIf(dblMoney < 0, "当前未退", "当前未付")
       lblBalance(Idx_lbl缴款).Caption = IIf(dblMoney < 0, "退    款", "收    款")
       '设置字体显示
        lngColor = IIf(dblMoney < 0, vbRed, vbBlue)
        lbl剩余自付.ForeColor = lngColor
        IDKindPaymentsType.ForeColor = lngColor
        lblBalance(Idx_lbl缴款).ForeColor = lngColor
        txtBalance(Idx_缴款).ForeColor = lngColor
    End With
    Show误差金额
End Sub
Private Sub MoveIDKindItem(ByVal objKind As IDKindNew, ByVal KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移动IDKind项目
    '入参:objKind-移动的IDKind对象
    '     Keyascii-键值
    '编制:刘兴洪
    '日期:2015-01-29 15:22:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If objKind Is Nothing Then Exit Sub
    If Not (KeyAscii = Asc("+") Or KeyAscii = Asc("-")) Then Exit Sub
    If objKind.ListCount = 1 Then Exit Sub
    
    If KeyAscii = Asc("+") Then
        '下移一项
        If objKind.IDKIND + 1 > objKind.ListCount Then
            objKind.IDKIND = 1
        Else
            objKind.IDKIND = objKind.IDKIND + 1
        End If
        Exit Sub
    End If
    If KeyAscii = Asc("-") Then '上移一项
        If objKind.IDKIND - 1 <= 0 Then
            objKind.IDKIND = objKind.ListCount
        Else
            objKind.IDKIND = objKind.IDKIND - 1
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub ClearSquareBalance(ByVal lngCardTypeID As Long, _
    Optional ByVal lng消费卡ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除消费卡结算
    '编制:刘兴洪
    '日期:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("编辑状态"))) = 2 _
                And Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = lngCardTypeID _
                And (lng消费卡ID = 0 Or (lng消费卡ID <> 0 And Val(.TextMatrix(j, .ColIndex("消费卡ID"))) = lng消费卡ID)) Then
                dblMoney = Val(.Cell(flexcpData, j, .ColIndex("金额")))
                mtyBalanceInfor.dbl已付合计 = FormatEx(mtyBalanceInfor.dbl已付合计 - dblMoney, 6)
                mtyBalanceInfor.dbl未付合计 = FormatEx(mtyBalanceInfor.dbl未付合计 + dblMoney, 6)
                If .Rows > 2 Then
                    .RemoveItem j
                Else
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                   .RowData(1) = ""
                   j = 2
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub

Private Sub LoadDefaultMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缺省的缴款或退款金额
    '编制:刘兴洪
    '日期:2015-01-30 17:38:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    If mbytInState = 1 Then Exit Sub
    
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then Exit Sub
    
    If objCard.接口序号 > 0 Then
         If Not objCard.消费卡 Then
             '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
             txtBalance(Idx_缴款).Text = lbl剩余自付.Caption
         Else
             txtBalance(Idx_缴款).Text = lbl剩余自付.Caption
         End If
    ElseIf objCard.结算性质 <> 1 Then
         txtBalance(Idx_缴款).Text = ""
    Else
         txtBalance(Idx_缴款).Text = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
