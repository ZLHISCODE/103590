VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceUnit 
   Caption         =   "��Լ��λ���˽���"
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   375
   ClientWidth     =   11760
   Icon            =   "frmBalanceUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11760
   StartUpPosition =   1  '����������
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
            Begin VB.Label lbl�Ը��ϼ� 
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
               Caption         =   "���ν��ʺϼ�"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
            Begin VB.Label lblʣ���Ը� 
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
               Caption         =   "��ǰδ��"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
            Caption         =   "ȷ��(&O)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ȡ��(&C)"
            BeginProperty Font 
               Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               IDKindStr       =   "��|�ֽ�|0|0|0|0|0|0;֧|֧Ʊ|0|0|0|0|0|"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontSize        =   12
               FontName        =   "����"
               IDKind          =   -1
               DefaultCardType =   "0"
               AllowAutoCommCard=   0   'False
               BackColor       =   -2147483633
            End
            Begin VB.Label lblBalance 
               AutoSize        =   -1  'True
               Caption         =   "��    ��"
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
               Index           =   1
               Left            =   75
               TabIndex        =   8
               Top             =   645
               Width           =   1050
            End
            Begin VB.Label lblBalance 
               AutoSize        =   -1  'True
               Caption         =   "ժ    Ҫ"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "�������"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "��    ��"
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
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���:0.0"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.CommandButton cmd���� 
            Height          =   360
            Left            =   7095
            Picture         =   "frmBalanceUnit.frx":157FA
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "����(F3)"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
         Begin VB.TextBox txt�ۼƽ�� 
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
            Caption         =   "���ݺ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Ʊ�ݺ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��λ(&D)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Լ��λ���˽���"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl�ۼ� 
            AutoSize        =   -1  'True
            Caption         =   "�ۼ�"
            BeginProperty Font 
               Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Name            =   "����"
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
Option Explicit 'Ҫ���������

'��ڲ�����
Private mbytInState As Byte          '0=����״̬(Ĭ������),1=���״̬
Private mblnViewCancel As Boolean    '�Ƿ�鿴�����ϵ���
Private mlng����ID As Long           'Ҫ����ĵ��ݺŽ���ID,��mbytInState=1ʱ��Ч
Private mblnNOMoved As Boolean       '�����ĵ����Ƿ��ں����ݱ���
'------------------------------------------------------------------------------
Private mrsPatients As ADODB.Recordset    '���ν��ʵĲ���IDδ����ü�¼��
Private mstrDec As String       '���ν��ʵķ���С��λ��
Private mintDefault As Integer  'ȱʡ�Ľ��㷽ʽ�к�
Private mlng����ID As Long
Private mintError As Integer    '���ѵĽ��㷽ʽ�к�
Private mintSucces As Integer
Private mobjInvoice As clsInvoice
Private mobjFact As clsFactProperty
Private mlngModul As Long
Private mstrPrivs As String
Private mrsBalance As ADODB.Recordset
Private mstrErrorBalance As String '�����㷽ʽ
Private mblnNotChange As Boolean
Private mblnPrintBill As Boolean '�Ƿ��ӡƱ��
Private mstr��֧Ʊ As String
Private mobjICCard As Object
Private mstrȱʡ���㷽ʽ As String
Private mrs���㷽ʽ As ADODB.Recordset
Private mblnUnload As Boolean
Private Enum mInput_Idx
    Idx_�ɿ� = 0
    Idx_�Ҳ� = 1
    Idx_������� = 2
    Idx_ժҪ = 3
End Enum

'��ǰ��������
Private Type TY_Balance_Infor
    dbl��ǰ���� As Double
    dbl�Ѹ��ϼ� As Double
    dblδ���ϼ� As Double
    
    blnSaveBill As Boolean '��ǰ�Ѿ�������ʵ�
    strNO As String   '��ǰ����Ľ��ʵ�
    lng����ID As Long '��ǰ����Ľ���ID
    dtBalanceDate As Date '��ǰ����ʱ��
    dbl�ɿ� As Double
    dbl�Ҳ� As Double
    dbl��֧Ʊ As Double
    dbl���� As Double
    dbl�ֽ� As Double
    lng����ID As Long
End Type
Private mtyBalanceInfor As TY_Balance_Infor
Private mcllCurSquareBalance As Collection '��ǰ���ѿ����Ѽ�
'��ǰˢ����Ϣ
Private Type TY_BrushCard    'ˢ������
    str���� As String
    str���� As String
    str������ˮ�� As String    '������ˮ��
    str����˵��  As String     '������Ϣ
    str��չ��Ϣ As String    '���׵���չ��Ϣ
    dbl�ʻ���� As Double
    str������� As String
    str����ժҪ As String
    blnת�� As Boolean '�Ƿ�ǰΪת�ʽ���
End Type
Private Enum mInput_LblIdx
    Idx_lbl�ɿ� = 0
    Idx_lbl�Ҳ� = 1
End Enum

'3.3 ģ���������
Private Type Ty_ModulePara
    byt�ɿ�������� As Byte  '
End Type
Private mty_ModulePara As Ty_ModulePara
'-----------------------------------------------------------------
'3.4�ϰ�һ��ͨ���
Private Type TY_OneCard
      blnOneCard As Boolean      '�Ƿ�������һ��ͨ�ӿ�
      rsOneCard As ADODB.Recordset
      strOneCard As String       '����ʱ��ѡ���һ��ͨ�ӿڶ�Ӧ�Ľ��㷽ʽ
End Type
Private mOldOneCard As TY_OneCard

'Private Enum BALANCECOL
'    C0���� = 0
'    C1�Ա� = 1
'    C2���� = 2
'    C3���ʽ�� = 3
'End Enum
'Private Enum PAYCOL
'    C0��ʽ = 0
'    C1��� = 1
'    C2���� = 2
'    C3��ע = 3
'End Enum
Private Const CASHPAY = 1

Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mty_ModulePara
        '����:43153:0-�����п���;1-������ȡ�ֽ�ʱ,��������ɿ�;2-����ʱ���������ۼ�
        .byt�ɿ�������� = Val(zlDatabase.GetPara("���ʽɿ��������", glngSys, mlngModul, 0))
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
    Optional ByVal lng����ID As Long, Optional ByVal blnViewCancel As Boolean, Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ҩ��λ����
    '���:bytInState -0=����״̬(Ĭ������),1=���״̬
    '    blnViewCancel-�Ƿ�鿴�����ϵ���
    '    lng����ID-Ҫ����ĵ��ݺŽ���ID,��bytInState=1ʱ��Ч
    '    blnNOMoved-�����ĵ����Ƿ��ں����ݱ���
    '����:���ʳɹ�1������,����true,���򷵻�False
    '����:���˺�
    '����:2015-02-05 11:17:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mbytInState = bytInState: mblnViewCancel = blnViewCancel: mlng����ID = lng����ID
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
    '����:���������ݵĺϷ���
    '����:tyBrushCard��ǰˢ����Ϣ
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2015-05-18 10:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSetFocus As Object, objCard As Card
    Dim intMouse As Integer
    Set objCard = IDKindPaymentsType.GetCurCard
    
    On Error GoTo errHandle
    If Val(txtUnit.Tag) = 0 Then
        MsgBox "����ѡ���Լ��λ�Ľ��ʲ���!", vbInformation, gstrSysName
        Set objSetFocus = txtUnit: GoTo GoExit:
        Exit Function
    End If
    If IsFirstInputBalanceMoney Then
        '��һ������ʱ,��Ҫ����ܷ��ò���ʾ
        If Val(lbl�Ը��ϼ�.Caption) = 0 Then
            If MsgBox("��ѡ����ʵ��û�пɽ����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Set objSetFocus = txtUnit: GoTo GoExit:
            End If
        End If
    End If
    If InStr(txtBalance(Idx_�������).Text, "'") > 0 Then
        MsgBox "������뺬�зǷ��ַ�������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_�������): GoTo GoExit:
         Exit Function
    End If
    
    If zlCommFun.ActualLen(txtBalance(Idx_�������).Text) > 30 Then
        MsgBox "����������ֻ������30���ַ���15������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_�������): GoTo GoExit:
         Exit Function
    End If
    
    If InStr(txtBalance(Idx_ժҪ).Text, "'") > 0 Then
        MsgBox "ժҪ���зǷ��ַ�������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_ժҪ): GoTo GoExit:
         Exit Function
    End If
 
    If zlCommFun.ActualLen(txtBalance(Idx_ժҪ).Text) > 30 Then
        MsgBox "����������ֻ������50���ַ���25������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_ժҪ): GoTo GoExit:
         Exit Function
    End If
    
    '��Ʊ���
    If CheckFactIsValied(objSetFocus) = False Then GoTo GoExit:
    '��鵱ǰ����Ŀ���������ݺϷ���
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

Private Function CheckCurBalanceIsValied(ByRef tyBrushCard As TY_BrushCard, ByVal blnԤ�� As Boolean, Optional ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ�����Ƿ���Ч
    '����:tyBrushCard��ǰˢ����Ϣ
    '     objSetFocus-����ƶ�����
    '����:��Ч����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 14:57:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng����ID As Long, varData As Variant
    Dim dblMoney As Double, i As Long, blnFind As Boolean
    Dim cllDeposit As Collection, int���� As Integer
    Dim intMouse As Integer
    Dim intCount As Integer '���ֽ��㷽ʽ(�ſ�ҽ��)
    On Error GoTo errHandle
    
    intMouse = Screen.MousePointer
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "��ǰ��������Ч��֧����ʽ����ѡ����Ч��֧����ʽ!", vbInformation + vbOKOnly, gstrSysName
        Set objSetFocus = IDKindPaymentsType
        Exit Function
    End If
    
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If blnԤ�� Then
                If int���� = 1 Then blnFind = True: Exit For
            End If
            If Not (objCard.���ѿ� And objCard.���ƿ�) Then '���ѿ�,�Ѿ����,�����ٴ���
                If .TextMatrix(i, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ Then blnFind = True
            End If
            If InStr("34", int����) > 0 Then
                MsgBox "������ʹ��:" & .TextMatrix(i, .ColIndex("֧����ʽ")) & "���н���!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            int���� = Val(.TextMatrix(i, .ColIndex("��������")))
            If InStr(",1,2,", "," & int���� & ",") > 0 Then intCount = intCount + 1
        Next
        
        If blnFind Then
            Screen.MousePointer = 0
            If blnԤ�� Then
                MsgBox "�Ѿ���Ԥ���֧��,ֻ��ɾ��Ԥ�������֧��!", vbOKOnly, gstrSysName
            Else
                MsgBox objCard.���㷽ʽ & " �Ѿ�֧����,��������" & objCard.���㷽ʽ & "����֧��", vbOKOnly + vbDefaultButton1, gstrSysName
            End If
            Exit Function
        End If
        
        If InStr("34", objCard.��������) > 0 Then
            MsgBox "������ʹ��:" & objCard.���㷽ʽ & "���н���!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With

    '���ݼ��ӿ���(Ŀǰֻͬʱ֧�����ֽӿ�(��ҽ����һ�ֽӿ�)
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    
    
    '1.���ѿ����
    If CheckSquareBalanceValied(objCard, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
     
    '2.�����ʻ����
    If CheckThreeSwapValied(objCard, dblMoney, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '3.һ��ͨ(�ϰ�)���
    If CheckOldOneCardIsValied(objCard, dblMoney, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '4.����ֽ���㷽ʽ
    If CheckCashValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '5.���֧Ʊ���㷽ʽ�Ƿ�Ϸ�
    If CheckChequeValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '6.����������㷽ʽ
    If CheckOtherValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
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
    '����:����������㷽ʽ(֧Ʊ��)��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl��ǰδ�� As Double
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.�ӿ���� > 0 Or objCard.���㷽ʽ Like "*֧Ʊ*" Or objCard.�������� = 1 Then CheckOtherValied = True: Exit Function
    
    dbl��ǰδ�� = mtyBalanceInfor.dblδ���ϼ�
    strTittle = IIf(dbl��ǰδ�� < 0, "�˿�", "�տ�")
    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), "0.00")
  
    If strTittle = "�տ�" Then
        If FormatEx(dblMoney, 6) = 0 Then
            Screen.MousePointer = 0
            MsgBox "δ����" & strTittle & "��", vbInformation, gstrSysName
            Exit Function
        End If
        If dblMoney > FormatEx(dbl��ǰδ��, 2) Then
            Screen.MousePointer = 0
            MsgBox "ע��:" & vbCrLf & "    �����" & strTittle & "��������δ֧���Ľ��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckOtherValied = True
        Exit Function
    End If
    
    '�˿�
    If FormatEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "δ����" & strTittle & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If dblMoney > FormatEx(Abs(dbl��ǰδ��), 2) Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "    ������˿��������δ�˽��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
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
    '����:���֧Ʊ���㷽ʽ��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl��ǰδ�� As Double
    Dim intMousePointer As Integer
    Dim objTempCard As Card
    Dim blnCheck As Boolean
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.�������� <> 2 Or Not objCard.���㷽ʽ Like "*֧Ʊ*" Then CheckChequeValied = True: Exit Function
    
    
    dbl��ǰδ�� = mtyBalanceInfor.dblδ���ϼ�
    
    strTittle = IIf(dbl��ǰδ�� < 0, "�˿�", "�տ�")
    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), "0.00")
     
    If strTittle = "�տ�" Then
    
        If FormatEx(dblMoney, 6) = 0 Then
            Screen.MousePointer = 0
            MsgBox "δ�����տ��", vbInformation, gstrSysName
            Exit Function
        End If
        If mstr��֧Ʊ = "" And blnCheck Then
            Screen.MousePointer = 0
            MsgBox "�ڽ��㷽ʽ��û������Ӧ����Ľ��㷽ʽ,���ܽ�����֧Ʊ����", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckChequeValied = True
        Exit Function
    End If
    
    '�˿�
    If FormatEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "δ�����˿��", vbInformation, gstrSysName
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

Private Function CheckCashValied(ByVal objCard As Card, Optional ByVal bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ���㷽ʽ��һЩ�Ϸ�����
    '���:objCard����ǰ֧����
    '     bln�˿�
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, strTittle As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer

    
    On Error GoTo errHandle
    If objCard.�������� <> 1 Then CheckCashValied = True: Exit Function
    
    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), "0.00")
    If Not bln�˿� Then
        If FormatEx(dblMoney, 6) <> 0 Then
            If Val(dblMoney) < Val(lblʣ���Ը�.Caption) Then
                Screen.MousePointer = 0
                MsgBox "�տ����,�벹��Ӧ�ս�" & vbCrLf & "����Ӧ��:" & lblʣ���Ը�.Caption & vbCrLf & "��ǰ�տ�" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
                Exit Function
            End If
        End If
        '43153
        '�ɿ����:0-�����п���;1-������ȡ�ֽ�ʱ,��������ɿ�.
        If mty_ModulePara.byt�ɿ�������� = 0 Then CheckCashValied = True: Exit Function
        If txtBalance(Idx_�ɿ�).Text = "" Then
            Screen.MousePointer = 0
            MsgBox "�㻹δ����ɿ���,���ܼ���", vbExclamation, gstrSysName
            Exit Function
        End If

        CheckCashValied = True
        Exit Function
    End If
    
    '�˿��
    If dblMoney < Abs(Val(lblʣ���Ը�.Caption)) And FormatEx(dblMoney, 6) <> 0 Then
        Screen.MousePointer = 0
        MsgBox "������˿���㣡", vbInformation, gstrSysName
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
    Optional bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ���ȷ
    '���:objCard-��ǰ������
    '     bln�˿�-�Ƿ��˿�
    '����:tyBrushCard-����ˢ����Ϣ
    '����:һ��ͨ��֤��ȷ���һ��ͨ,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 17:19:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblδ����� As Double, strCardNo As String
    Dim dblTemp As Double, strXmlIn As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.�������� <> 7 Then CheckOldOneCardIsValied = True: Exit Function
    
    mOldOneCard.rsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mOldOneCard.rsOneCard.EOF Then
        Screen.MousePointer = 0
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        CheckOldOneCardIsValied = False: Exit Function
    End If

    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "һ��ͨ�ӿڴ���ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblMoney = 0 Then dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
     
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "���δ����,����!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    dblδ����� = FormatEx(mtyBalanceInfor.dblδ���ϼ�, 6)
    If Abs(dblMoney) > Format(Abs(dblδ�����), "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "���ܴ��ڱ���" & IIf(bln�˿�, "δ��", "δ��") & "���:" & Format(Abs(dblδ�����), "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Val(lblʣ���Ը�.Caption) <> dblMoney Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "����:" & Format(Abs(Val(lblʣ���Ը�.Caption)), "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
            
            
    If Not bln�˿� Then
       
       '����ˢ������
       'zlBrushCard(frmMain As Object, _
       '    ByVal lngModule As Long, _
       '    ByVal rsClassMoney As ADODB.Recordset, _
       '    ByVal lngCardTypeID As Long, _
       '    ByVal bln���ѿ� As Boolean, _
       '    ByVal strPatiName As String, ByVal strSex As String, _
       '    ByVal strOld As String, ByVal dbl��� As Double, _
       '    Optional ByRef strCardNo As String, _
       '    Optional ByRef strPassWord As String, _
       '    Optional ByRef bln�˷� As Boolean = False, _
       '    Optional ByRef blnShowPatiInfor As Boolean = False, _
       '    Optional ByRef bln���� As Boolean = False, _
       '    Optional ByVal bln�����ֹ As Boolean = True) As Boolean
       '---------------------------------------------------------------------------------------------------------------------------------------------
       '����:����ָ��֧�����,����ˢ������
       '���:rsClassMoney:�շ����,���
       '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
       '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
        
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, 0, False, _
        txtUnit.Text, "", "", dblMoney, tyBrushCard.str����, tyBrushCard.str����, _
        False, True, False, False, Nothing, False, False, strXmlIn) = False Then Exit Function
        
        tyBrushCard.dbl�ʻ���� = mobjICCard.GetSpare
        If tyBrushCard.dbl�ʻ���� < dblMoney Then
            Screen.MousePointer = 0
            MsgBox "������֧��,����!" & vbCrLf & vbCrLf & _
            "   �� ��  ��" & Format(tyBrushCard.dbl�ʻ����, "0.00") & vbCrLf & _
            "   ����֧��" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
            Exit Function
        End If
        staThis.Panels(2).Text = Format(tyBrushCard.dbl�ʻ����, "0.00")
        staThis.Panels(2).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(tyBrushCard.dbl�ʻ����, "0.00")
       
        CheckOldOneCardIsValied = True
        Exit Function
    End If
    '�˿���
    If mrsBalance Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    mrsBalance.Filter = "����=4"
    If mrsBalance.EOF Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        Screen.MousePointer = 0
        MsgBox "һ��ͨ����ʧ��,�뽫IC�����ڶ�������", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> NVL(mrsBalance!����) Then
        Screen.MousePointer = 0
        MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblTemp = Format(Val(NVL(mrsBalance!��Ԥ��)), "0.00")
    If FormatEx(dblMoney, 6) <> Format(dblTemp, "0.00") Then
        Screen.MousePointer = 0
        MsgBox "һ��ͨ�������ȫ��,����!" & vbCrLf & vbCrLf & _
        "   ������" & Format(dblTemp, "0.00") & vbCrLf & _
        "   ����֧��" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
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
    '����:�Ƿ��һ������ɿ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-05-18 11:29:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then Exit Function
        Next
    End With
    IsFirstInputBalanceMoney = True
End Function

Private Function CheckFactIsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ���Ч
    '����:objSetFocus -����ʱ,��궨λ���ĸ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-05-18 11:09:38
    '˵��:��һ������ɿ�����ʱ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, i As Long
    On Error GoTo errHandle
    
    If Not IsFirstInputBalanceMoney Then CheckFactIsValied = True: Exit Function
    '��һ������,������
    'Ʊ�ݺ�����
    mblnPrintBill = False
    If mobjFact.��ӡ��ʽ = 0 Then CheckFactIsValied = True: Exit Function
 
    mblnPrintBill = True
    If mobjFact.��ӡ��ʽ = 2 Then
        If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then mblnPrintBill = False
        If mblnPrintBill = False Then CheckFactIsValied = True: Exit Function
    End If
    
    If Not mobjFact.�ϸ���� Then
        If Len(txtInvoice.Text) <> mobjFact.Ʊ�ų��� And txtInvoice.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & mobjFact.Ʊ�ų��� & " λ��", vbInformation, gstrSysName
            Set objSetFocus = txtInvoice: Exit Function
        End If
        CheckFactIsValied = True: Exit Function
    End If
    
    '�ϸ�Ʊ�ݹ���
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
        Set objSetFocus = txtInvoice: Exit Function
    End If
    
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.�û���, mobjFact.Ʊ��, mobjFact.ʹ�����, mlng����ID, mobjFact.��������ID, mlng����ID, 1, Trim(txtInvoice.Text)) = False Then Exit Function
    
    If mlng����ID <= 0 Then
        Select Case mlng����ID
            Case 0 '����ʧ��
            Case -1
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -3
                MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,����������", vbInformation, gstrSysName
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
    '����:������
    '����:���˺�
    '����:2015-05-18 15:55:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    cmdOK.Enabled = Not blnLocked
    txtBalance(Idx_�ɿ�).Locked = blnLocked
    txtBalance(Idx_�������).Locked = blnLocked
    txtBalance(Idx_ժҪ).Locked = blnLocked
    txtBalance(Idx_lbl�Ҳ�).Locked = True
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveBalanceData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2015-05-18 10:57:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngBalanceID As Long, i As Long, lngPatientID As Long
    Dim strNO As String, strTmp As String
    Dim tyBrushCard As TY_BrushCard
    Dim dblδ����� As Double, dblʣ���� As Double, dbl��֧Ʊ�� As Double
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
        .dbl�ɿ� = 0: .dbl�Ҳ� = 0
        .dbl�ֽ� = 0
        dblδ����� = FormatEx(.dblδ���ϼ�, 6)
    End With
    
    intSign = IIf(dblδ����� < 0, -1, 1)
    If objCard.�������� = 1 Then     '�ֽ�
        dblMoney = FormatEx(intSign * Val(txtBalance(Idx_�ɿ�).Text), 6)
        If dblMoney <> 0 Then
            mtyBalanceInfor.dbl�ɿ� = dblMoney
            mtyBalanceInfor.dbl�Ҳ� = IIf(lblBalance(Idx_lbl�Ҳ�).Caption Like "��*", -1, 1) * Val(txtBalance(Idx_�Ҳ�).Text)
        End If
        dblTemp = dblδ�����: dblʣ���� = 0
        dblMoney = GetCentMoney(dblTemp)
        mtyBalanceInfor.dbl�ֽ� = dblMoney
    ElseIf objCard.���� Like "*֧Ʊ" Then
        dblMoney = FormatEx(intSign * (Val(txtBalance(Idx_�ɿ�).Text)), 6)
        dblʣ���� = FormatEx(dblδ����� - dblMoney, 6)
        If dblʣ���� < 0 Then
            mtyBalanceInfor.dbl��֧Ʊ = -1 * Val(txtBalance(Idx_�Ҳ�).Text)
            dblʣ���� = 0
        End If
    Else    '�������㷽ʽ֧��
        dblMoney = FormatEx(intSign * Val(txtBalance(Idx_�ɿ�).Text), 6)
        dblʣ���� = FormatEx(dblδ����� - dblMoney, 6)
    End If
    
    Call Show�����
    If Abs(mtyBalanceInfor.dbl����) > 1.5 Then
        Screen.MousePointer = 0
        Call MsgBox("������,�����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName)
        Call LockedScreen(False)
        Exit Function
    End If
    If FormatEx(mtyBalanceInfor.dbl����, 6) <> 0 Then
        If mstrErrorBalance = "" Then
            Screen.MousePointer = 0
            MsgBox "��Ӧ�ó�����δ�����������ڽ��㷽ʽ������!", vbInformation + vbOKOnly, gstrSysName
            Call LockedScreen(False)
            Exit Function
        End If
    End If
    If dblʣ���� <> 0 Then
        '����ʣ������
        With vsBlance
            If objCard.���ѿ� Then
                Call AddSquareBalance(objCard)
            Else
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                strCardNo = tyBrushCard.str����
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = 0
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                If objCard.�������� = 7 And objCard.�ӿ���� < 0 Then
                    .TextMatrix(1, .ColIndex("����")) = 4
                    .TextMatrix(1, .ColIndex("�༭״̬")) = 0   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                    .TextMatrix(1, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                ElseIf objCard.�ӿ���� > 0 Then
                    .TextMatrix(1, .ColIndex("����")) = 3
                    .TextMatrix(1, .ColIndex("�����ID")) = objCard.�ӿ����
                    .TextMatrix(1, .ColIndex("���������")) = objCard.����
                    .TextMatrix(1, .ColIndex("�༭״̬")) = 0   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                    .TextMatrix(1, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                    .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�������Ĺ��� <> "", 1, 0)
                Else
                    .TextMatrix(1, .ColIndex("����")) = 0
                    .TextMatrix(1, .ColIndex("�༭״̬")) = 2   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                    .TextMatrix(1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                End If
                .TextMatrix(1, .ColIndex("��������")) = objCard.��������
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(1, .ColIndex("У�Ա�־")) = 2
                
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                .TextMatrix(1, .ColIndex("���")) = Format(dblMoney, "0.00")
                .TextMatrix(1, .ColIndex("�������")) = IIf(txtBalance(Idx_�������).Visible, Trim(txtBalance(Idx_�������).Text), "")
                .TextMatrix(1, .ColIndex("��ע")) = Trim(txtBalance(Idx_ժҪ).Text)
                .TextMatrix(1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = tyBrushCard.str����
                .TextMatrix(1, .ColIndex("������ˮ��")) = tyBrushCard.str������ˮ��
                .TextMatrix(1, .ColIndex("����˵��")) = tyBrushCard.str����˵��
                mtyBalanceInfor.dbl�Ѹ��ϼ� = FormatEx(mtyBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
                mtyBalanceInfor.dblδ���ϼ� = FormatEx(mtyBalanceInfor.dblδ���ϼ� - dblMoney, 6)
            End If
            For i = 1 To IDKindPaymentsType.ListCount
                'ȱʡ��λ���ֽ���
                 Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
                If objCard.�������� = 1 Then IDKindPaymentsType.IDKIND = i: Exit For
            Next
        End With
        Call LockedScreen(False)
        Screen.MousePointer = 0
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
        txtBalance(Idx_�ɿ�).Text = ""
        Call LoadCurOwnerPayInfor
        SaveBalanceData = True
        Exit Function
    End If
    
    
    Set cllPro = New Collection: Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    '��������
    If GetSaveBalanceSQL(tyBrushCard, mtyBalanceInfor, cllPro) = False Then Exit Function
    'ִ��һ��ͨ(�ϰ�)�ӿ�
    If ExecuteOldOneCardPayInterface(0, mtyBalanceInfor.lng����ID, objCard, dblMoney, tyBrushCard, cllPro) = False Then
        Call LockedScreen(False): Screen.MousePointer = 0
        Exit Function
    End If
    'ִ�������ʻ����׽ӿ�
    If ExecuteThreeSwapPayInterface(0, mtyBalanceInfor.lng����ID, objCard, dblMoney, cllPro, tyBrushCard) = False Then
        Call LockedScreen(False): Screen.MousePointer = 0
        Exit Function
    End If
    If cllPro.Count <> 0 Then
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption
        blnTrans = False
    End If
    
    Call LockedScreen(False): Screen.MousePointer = 0
    Call PrintBill '��ӡƱ��
    
    txt�ۼƽ��.Text = Format(Val(txt�ۼƽ��.Text) + Val(lbl�Ը��ϼ�.Caption), gstrDec)
    
    '������ʷ��¼
    strNO = mtyBalanceInfor.strNO
    strTmp = mtyBalanceInfor.strNO
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
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
    '����:��ӡƱ��
    '����:���˺�
    '����:2015-05-18 15:48:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, i As Long
    
    On Error GoTo errHandle
    'Ʊ�ݴ�ӡ
    If Not mblnPrintBill Then Exit Sub
    If Not gblnPrintByPatient Then
        '�������˴�ӡ
        Call frmPrint.ReportPrint(1, mtyBalanceInfor.strNO, mtyBalanceInfor.lng����ID, mobjFact, txtInvoice.Text, mtyBalanceInfor.dtBalanceDate, CStr(mtyBalanceInfor.dbl�ɿ�), CStr(mtyBalanceInfor.dbl�Ҳ�), , mobjFact.��ӡ��ʽ)
        Exit Sub
    End If
    '�����˴�ӡ
    If mrsPatients.RecordCount > 0 Then mrsPatients.MoveFirst
    For i = 1 To mrsPatients.RecordCount
        lng����ID = Val(NVL(mrsPatients!����ID))
        Call frmPrint.ReportPrint(1, mtyBalanceInfor.strNO, mtyBalanceInfor.lng����ID, mobjFact, txtInvoice.Text, mtyBalanceInfor.dtBalanceDate, CStr(mtyBalanceInfor.dbl�ɿ�), CStr(mtyBalanceInfor.dbl�Ҳ�), lng����ID, mobjFact.��ӡ��ʽ)
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
    '����:�������ݱ���
    '���:tbBrushCard-��ǰˢ����Ϣ
    '����:tyBalanceInfor-���ص�ǰ������Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-05-18 14:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng����ID As Long, strTmp As String, str���NO As String, strPatiIDs As String
    Dim lngFirstPatiID As Long   '���ڴ������ʱ,����������������������ʵ������Ϣ
    Dim strSql As String, objCard As Card
    Dim strNO As String, datBalance As Date, dblMoney As Double
    Dim str���ѿ����� As String
    
    
    Err = 0: On Error GoTo ErrHand:
    
    If mrsPatients.RecordCount > 0 Then mrsPatients.MoveFirst
    For i = 1 To mrsPatients.RecordCount
        If i = 1 Then lngFirstPatiID = Val(mrsPatients!����ID)
        strPatiIDs = strPatiIDs & mrsPatients!����ID & ","
        mrsPatients.MoveNext
    Next
    Set objCard = IDKindPaymentsType.GetCurCard
    strNO = zlDatabase.GetNextNo(15)
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    datBalance = zlDatabase.Currentdate
    
    With tyBalanceInfor
        .strNO = strNO
        .lng����ID = lng����ID
        .dtBalanceDate = datBalance
    End With
    '1.���˽��ʼ�¼
    'Zl_���˽��ʼ�¼_Insert
    strSql = "zl_���˽��ʼ�¼_Insert("
    '  Id_In           ���˽��ʼ�¼.ID%Type,
    strSql = strSql & "" & lng����ID & ","
    '  ���ݺ�_In       ���˽��ʼ�¼.NO%Type,
    strSql = strSql & "'" & strNO & "',"
    '  ����id_In       ���˽��ʼ�¼.����id%Type,
    strSql = strSql & "" & lngFirstPatiID & ","
    '  �շ�ʱ��_In     ���˽��ʼ�¼.�շ�ʱ��%Type,
    strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ��ʼ����_In     ���˽��ʼ�¼.��ʼ����%Type,
    strSql = strSql & "NULL,"
    '  ��������_In     ���˽��ʼ�¼.��������%Type,
    strSql = strSql & "NULL,"
    '  ��;����_In     ���˽��ʼ�¼.��;����%Type := 0,
    strSql = strSql & "0,"
    '  �ಡ�˽���_In   Number := 0,
    strSql = strSql & "1,"
    '  �����ʴ���_In Number := 0,
    strSql = strSql & "0,"
    '  ��ע_In         ���˽��ʼ�¼.��ע%Type := Null,
    strSql = strSql & "NULL,"
    '  ��Դ_In         Number := 1, '  --1.��Դ_In:1-����;2-סԺ
    strSql = strSql & "1,"
    '  ԭ��_In         ���˽��ʼ�¼.ԭ��%Type := Null   '�洢��Լ��λID'����:35090
    strSql = strSql & "'" & Trim(txtUnit.Text) & "',1)"
    zlAddArray cllPro, strSql
    '2.���ʽɿ��¼
    With vsBlance
        For i = 1 To .Rows - 1
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If Val(.TextMatrix(i, .ColIndex("����"))) <> 5 And _
                Val(.TextMatrix(i, .ColIndex("���"))) <> 0 Then
                'Zl_���ʽɿ��¼_Insert
                strSql = "zl_���ʽɿ��¼_Insert("
                '  No_In         ���˽��ʼ�¼.No%Type,
                strSql = strSql & "'" & strNO & "',"
                '  ����id_In     ����Ԥ����¼.����id%Type,
                strSql = strSql & "NULL,"
                '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
                strSql = strSql & "NULL,"
                '  ����id_In     ����Ԥ����¼.����id%Type,
                strSql = strSql & "NULL,"
                '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("֧����ʽ")) & "',"
                '  �������_In   ����Ԥ����¼.�������%Type,
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("�������")) & "',"
                '  ���_In       ����Ԥ����¼.��Ԥ��%Type,
                strSql = strSql & "" & Val(.TextMatrix(i, .ColIndex("���"))) & ","
                '  ����id_In     ����Ԥ����¼.����id%Type,
                strSql = strSql & "" & lng����ID & ","
                '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
                strSql = strSql & "'" & UserInfo.��� & "',"
                '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
                strSql = strSql & "'" & UserInfo.���� & "',"
                '  �շ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
                strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  �������_In   �����ʻ�.����%Type,
                strSql = strSql & "NULL,"
                '  �����ʺ�_In   �����ʻ�.ҽ����%Type,
                strSql = strSql & "NULL,"
                '  ��������_In   �����ʻ�.����%Type,
                strSql = strSql & "NULL,"
                '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
                strSql = strSql & "NULL,"
                '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
                strSql = strSql & "NULL,"
                '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
                If Val(.TextMatrix(i, .ColIndex("�����ID"))) <> 0 Then
                    strSql = strSql & "" & Val(.TextMatrix(i, .ColIndex("�����ID"))) & ","
                Else
                    strSql = strSql & "NULL,"
                End If
                '  ����_In       ����Ԥ����¼.����%Type := Null,
                If Trim(.TextMatrix(i, .ColIndex("����"))) <> "" Then
                    strSql = strSql & "'" & Trim(.TextMatrix(i, .ColIndex("����"))) & "',"
                Else
                    strSql = strSql & "NULL,"
                End If
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                If Trim(.TextMatrix(i, .ColIndex("����˵��"))) <> "" Then
                    strSql = strSql & "'" & Trim(.TextMatrix(i, .ColIndex("������ˮ��"))) & "',"
                Else
                    strSql = strSql & "NULL,"
                End If
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
                If Trim(.TextMatrix(i, .ColIndex("����˵��"))) <> "" Then
                    strSql = strSql & "'" & Trim(.TextMatrix(i, .ColIndex("����˵��"))) & "',"
                Else
                    strSql = strSql & "NULL,"
                End If
                '  ���ѿ�����_In Varchar2 := Null:�����ID|����|���ѿ�ID|���ѽ��||."
                strSql = strSql & "NULL)"
                '��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null
                zlAddArray cllPro, strSql
            ElseIf Val(.TextMatrix(i, .ColIndex("����"))) = 5 Then
                If (objCard.�ӿ���� <> Val(.TextMatrix(i, .ColIndex("�����ID"))) And objCard.���ѿ�) Or Not objCard.���ѿ� Then
                    '���ѿ�
                    str���ѿ����� = str���ѿ����� & "||" & Val(.TextMatrix(i, .ColIndex("�����ID")))
                    str���ѿ����� = str���ѿ����� & "|" & Trim(.Cell(flexcpData, i, .ColIndex("����")))
                    str���ѿ����� = str���ѿ����� & "|" & Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                    str���ѿ����� = str���ѿ����� & "|" & FormatEx(Val(.TextMatrix(i, .ColIndex("���"))), 6)
                End If
            End If
        Next
    End With
    If objCard.�ӿ���� > 0 Then
        If objCard.���ѿ� Then
            For i = 1 To mcllCurSquareBalance.Count
                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                str���ѿ����� = str���ѿ����� & "||" & Val(mcllCurSquareBalance(i)(0))
                str���ѿ����� = str���ѿ����� & "|" & Trim(mcllCurSquareBalance(i)(3))
                str���ѿ����� = str���ѿ����� & "|" & Val(mcllCurSquareBalance(i)(1))
                str���ѿ����� = str���ѿ����� & "|" & FormatEx(Val(mcllCurSquareBalance(i)(2)), 6)
            Next
        Else
            'Zl_���ʽɿ��¼_Insert
            strSql = "zl_���ʽɿ��¼_Insert("
            '  No_In         ���˽��ʼ�¼.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "NULL,"
            '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
            strSql = strSql & "NULL,"
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "NULL,"
            '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
            strSql = strSql & "'" & objCard.���㷽ʽ & "',"
            '  �������_In   ����Ԥ����¼.�������%Type,
            strSql = strSql & "'" & txtBalance(Idx_�������).Text & "',"
            '  ���_In       ����Ԥ����¼.��Ԥ��%Type,
            strSql = strSql & "" & txtBalance(Idx_�ɿ�).Text & ","
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '  �շ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
            strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  �������_In   �����ʻ�.����%Type,
            strSql = strSql & "NULL,"
            '  �����ʺ�_In   �����ʻ�.ҽ����%Type,
            strSql = strSql & "NULL,"
            '  ��������_In   �����ʻ�.����%Type,
            strSql = strSql & "NULL,"
            '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
            strSql = strSql & "NULL,"
            '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
            strSql = strSql & "NULL,"
            '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSql = strSql & "" & objCard.�ӿ���� & ","
            '  ����_In       ����Ԥ����¼.����%Type := Null,
            strSql = strSql & "'" & tyBrushCard.str���� & "',"
            '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSql = strSql & "'" & tyBrushCard.str������ˮ�� & "',"
            '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
            strSql = strSql & "'" & tyBrushCard.str����˵�� & "',"
            '  ���ѿ�����_In Varchar2 := Null:�����ID|����|���ѿ�ID|���ѽ��||."
            strSql = strSql & "NULL)"
            '��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null
            zlAddArray cllPro, strSql
        End If
    Else
        '��ͨ���㷽ʽ
        dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), mstrDec)
        If objCard.�������� = 1 Then
            dblMoney = mtyBalanceInfor.dbl�ֽ�
        End If
        
        'Zl_���ʽɿ��¼_Insert
        strSql = "zl_���ʽɿ��¼_Insert("
        '  No_In         ���˽��ʼ�¼.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "NULL,"
        '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
        strSql = strSql & "NULL,"
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "NULL,"
        '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
        strSql = strSql & "'" & objCard.���㷽ʽ & "',"
        '  �������_In   ����Ԥ����¼.�������%Type,
        strSql = strSql & "'" & txtBalance(Idx_�������).Text & "',"
        '  ���_In       ����Ԥ����¼.��Ԥ��%Type,
        strSql = strSql & "" & dblMoney & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng����ID & ","
        '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
        strSql = strSql & "'" & UserInfo.��� & "',"
        '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
        strSql = strSql & "'" & UserInfo.���� & "',"
        '  �շ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
        strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  �������_In   �����ʻ�.����%Type,
        strSql = strSql & "NULL,"
        '  �����ʺ�_In   �����ʻ�.ҽ����%Type,
        strSql = strSql & "NULL,"
        '  ��������_In   �����ʻ�.����%Type,
        strSql = strSql & "NULL,"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        strSql = strSql & "NULL,"
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        strSql = strSql & "NULL,"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSql = strSql & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSql = strSql & "NULL,"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSql = strSql & "NULL,"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
        strSql = strSql & "NULL,"
        '  ���ѿ�����_In Varchar2 := Null:�����ID|����|���ѿ�ID|���ѽ��||."
        strSql = strSql & "NULL,"
        '��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null
        If objCard.���㷽ʽ Like "*֧Ʊ" And Val(txtBalance(Idx_lbl�Ҳ�).Text) <> 0 Then
            strSql = strSql & "" & mtyBalanceInfor.dbl��֧Ʊ & ")"
        Else
            strSql = strSql & "NULL)"
        End If
        zlAddArray cllPro, strSql
    End If
    
    '�������ѿ�
    If str���ѿ����� <> "" Then
        str���ѿ����� = Mid(str���ѿ�����, 3)
       'Zl_���ʽɿ��¼_Insert
            strSql = "zl_���ʽɿ��¼_Insert("
            '  No_In         ���˽��ʼ�¼.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "NULL,"
            '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
            strSql = strSql & "NULL,"
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "NULL,"
            '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
            strSql = strSql & "NULL,"
            '  �������_In   ����Ԥ����¼.�������%Type,
            strSql = strSql & "NULL,"
            '  ���_In       ����Ԥ����¼.��Ԥ��%Type,
            strSql = strSql & "NULL,"
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '  �շ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
            strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  �������_In   �����ʻ�.����%Type,
            strSql = strSql & "NULL,"
            '  �����ʺ�_In   �����ʻ�.ҽ����%Type,
            strSql = strSql & "NULL,"
            '  ��������_In   �����ʻ�.����%Type,
            strSql = strSql & "NULL,"
            '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
            strSql = strSql & "NULL,"
            '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
            strSql = strSql & "NULL,"
            '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSql = strSql & "NULL,"
            '  ����_In       ����Ԥ����¼.����%Type := Null,
            strSql = strSql & "NULL,"
            '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSql = strSql & "NULL,"
            '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
            strSql = strSql & "NULL,"
            '  ���ѿ�����_In Varchar2 := Null:�����ID|����|���ѿ�ID|���ѽ��||."
            strSql = strSql & "'" & str���ѿ����� & "')"
            '��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null
            zlAddArray cllPro, strSql
    End If
    If mtyBalanceInfor.dbl���� <> 0 Then

        
        'Zl_���ʽɿ��¼_Insert
        strSql = "zl_���ʽɿ��¼_Insert("
        '  No_In         ���˽��ʼ�¼.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "NULL,"
        '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
        strSql = strSql & "NULL,"
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "NULL,"
        '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
        strSql = strSql & "'" & mstrErrorBalance & "',"
        '  �������_In   ����Ԥ����¼.�������%Type,
        strSql = strSql & "'" & txtBalance(Idx_�������).Text & "',"
        '  ���_In       ����Ԥ����¼.��Ԥ��%Type,
        strSql = strSql & "" & mtyBalanceInfor.dbl���� & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng����ID & ","
        '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
        strSql = strSql & "'" & UserInfo.��� & "',"
        '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
        strSql = strSql & "'" & UserInfo.���� & "',"
        '  �շ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
        strSql = strSql & "To_Date('" & Format(datBalance, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  �������_In   �����ʻ�.����%Type,
        strSql = strSql & "NULL,"
        '  �����ʺ�_In   �����ʻ�.ҽ����%Type,
        strSql = strSql & "NULL,"
        '  ��������_In   �����ʻ�.����%Type,
        strSql = strSql & "NULL,"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        strSql = strSql & "NULL,"
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        strSql = strSql & "NULL,"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSql = strSql & "Null,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSql = strSql & "'" & tyBrushCard.str���� & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSql = strSql & "'" & tyBrushCard.str������ˮ�� & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
        strSql = strSql & "'" & tyBrushCard.str����˵�� & "',"
        '  ���ѿ�����_In Varchar2 := Null:�����ID|����|���ѿ�ID|���ѽ��||."
        strSql = strSql & "NULL,"
        '��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null
        strSql = strSql & "NULL)"
        zlAddArray cllPro, strSql
    End If
    '3.������ü�¼
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
        strSql = "Zl_���ʷ��ü�¼_Unit('" & strTmp & "'," & lng����ID & "," & IIf(gblnZero, 1, 0) & ")"
        zlAddArray cllPro, strSql
    Wend

    '4.��ʼƱ�ݺ�
    If mblnPrintBill And Trim(txtInvoice.Text) <> "" Then
        strSql = "Zl_Ʊ����ʼ��_Update('" & strNO & "','" & Trim(txtInvoice.Text) & "',3)"
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
    '����:ˢ���շ�Ʊ�ݺ�
    '����:���˺�
    '����:2015-02-05 11:40:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
       
    On Error GoTo errHandle
        
    If mobjFact.��ӡ��ʽ = 0 Then Exit Sub
    If Not mobjFact.�ϸ���� Then
        '���ϸ������
        '��ɢ��ȡ��һ������
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    If zlGetInvoiceGroupUseID(mlng����ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '�ϸ�ȡ��һ������
    If mobjInvoice.zlGetNextBill(1137, mlng����ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
    '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
    '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    txtInvoice.SelStart = Len(txtInvoice.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.����, mobjFact.Ʊ��, _
        mobjFact.ʹ�����, lng����ID, mobjFact.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng����ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng����ID
        Case 0 '����ʧ��
        Case -1
            If Trim(mobjFact.ʹ�����) = "" Then
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & mobjFact.ʹ����� & "������Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFact.ʹ�����) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFact.ʹ����� & "������Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
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
Private Sub cmd����_Click()
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
            If vsBlance.Rows > 0 Then vsBlance.SetFocus: vsBlance.Row = 1: vsBlance.Col = vsBlance.ColIndex("���")
            staThis.Panels(2).Text = "��ѡ����" & CStr(vsPati.Rows - 1) & "λ����."
            gblnOK = False
            Call LoadCurOwnerPayInfor
            Call LoadDefaultMoney
            If txtBalance(Idx_lbl�ɿ�).Enabled And txtBalance(Idx_lbl�ɿ�).Visible Then
                txtBalance(Idx_lbl�ɿ�).SetFocus
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
    If KeyCode = vbKeyF3 Then Call cmd����_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    If mbytInState = 0 Then
        mintDefault = 0: mlng����ID = 0: mintError = 0
        Call InitFact
        Call NewBalance
        txt�ۼƽ��.Text = Format(0, gstrDec)
        txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnStrictCtrl '89302
    Else
        Call InitViewFace
        Call LoadPatiBalanceData
    End If
    If Init���㷽ʽ = False Then Exit Sub
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:���˺�
    '����:2015-05-18 10:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlRaisEffect picBalance, Dw_SubKen
    zlRaisEffect picOwerFee, Dw_SubKen
    zlRaisEffect picNotPayment, Dw_SubKen
    Call zlInitModulePara
    Call InitOldOneCardInfor
    
End Sub

Private Sub InitFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ʊ��Ϣ
    '����:���˺�
    '����:2015-02-05 11:26:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytInvoiceKind     As Byte, intFormat As Integer
    Dim intPrintMode As Integer, lngShareUseID As Long
    On Error GoTo errHandle
    If mbytInState <> 0 Then Exit Sub
    
    bytInvoiceKind = Val(zlDatabase.GetPara("�������Ʊ������", glngSys, 1137, "0"))

    Set mobjInvoice = New clsInvoice: Set mobjFact = New clsFactProperty
    mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    Call mobjInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, mobjFact, , , 1)
    
    mobjFact.ʹ����� = zlDatabase.GetPara("��Լ��λ���ʴ�ӡ", glngSys, 1137)
    mobjFact.Ʊ�� = IIf(bytInvoiceKind = 0, 3, 1)
    Call mobjInvoice.zlGetInvoicePrintFormat(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, intFormat, 1)
    mobjFact.��ӡ��ʽ = intFormat
    If mobjInvoice.zlGetInvoicePrintMode(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, intPrintMode) = False Then Exit Sub
    mobjFact.��ӡ��ʽ = intPrintMode
    If mobjInvoice.zlGetInvoiceShareID(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, lngShareUseID) = False Then Exit Sub
    mobjFact.��������ID = lngShareUseID
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
    mlng����ID = 0
    zl_vsGrid_Para_Save mlngModul, vsPati, Me.Name, "�����б�"
    zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "�����б�"
    
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub InitViewFace()
    '����:�鿴ʱ��ʼ������
    
    If mblnViewCancel Then lblFlag.Visible = True
    
    Call InitPatiGrid
    Call InitGrid_PayList
    
    txtInvoice.Locked = True
    cboNO.Locked = True
    txtUnit.Locked = True
    txtUnit.Width = txtUnit.Width + cmd����.Width + 30
    cmd����.Visible = False
        
    lbl�ۼ�.Visible = False
    txt�ۼƽ��.Visible = False
    picCurBalance.Visible = False
    picOwerFee.Visible = False
    picNotPayment.Visible = False
    cmdOK.Visible = False
    cmdCancel.Caption = "�˳�(&X)"
    
    fraSplit.Visible = True
    
End Sub

Private Sub LoadPatiBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˽�������
    '����:���˺�
    '����:2015-05-15 15:41:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    
    '���ز��˽�������
    Set rsTmp = GetBalanceData(mlng����ID, mblnNOMoved)
    Call ShowBalanceInfo(rsTmp)
    Call LoadBalanceInfo(mlng����ID, mblnNOMoved)
    Call LoadOtherData(mlng����ID, mblnNOMoved)
    staThis.Panels(2).Text = "��ǰ���ʵ�����" & CStr(vsBlance.Rows - 1) & "λ����."
End Sub
Private Function LoadOtherData(ByVal lng����ID As Long, Optional ByVal blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '���:lng����ID-����ID
    '     blnNOMoved-�Ƿ��Ѿ�ת������ʷ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-05-15 17:08:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    
    Set rsTmp = GetOtherInfo(lng����ID, blnNOMoved)
    If Not rsTmp Is Nothing Then
        txtUnit.Text = "" & rsTmp!��Լ��λ
        txtInvoice.Text = "" & rsTmp!ʵ��Ʊ��
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
''    '����:�����������
''    '���:dblErrorMoney -�����
''    '����:���˺�
''    '����:2015-05-15 17:30:17
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
''            '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
''            .TextMatrix(i, .ColIndex("�༭״̬")) = ""
''            ''�Ƿ��ѽ���:1-�ѽ���;0-δ����
''            .TextMatrix(i, .ColIndex("����״̬")) = 0
''
''            .TextMatrix(i, .ColIndex("�����ID")) = Val(nvl(mrsBalance!�����id))
''            .TextMatrix(i, .ColIndex("���ѿ�ID")) = Val(nvl(mrsBalance!���ѿ�ID))
''            .TextMatrix(i, .ColIndex("��������")) = Val(nvl(mrsBalance!��������))
''            .TextMatrix(i, .ColIndex("����")) = Val(nvl(mrsBalance!����))
''            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(nvl(mrsBalance!�Ƿ�����))
''            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(nvl(mrsBalance!�Ƿ�ȫ��))
''            .TextMatrix(i, .ColIndex("У�Ա�־")) = Val(nvl(mrsBalance!У�Ա�־))
''            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(nvl(mrsBalance!�Ƿ�����))
''            .TextMatrix(i, .ColIndex("֧����ʽ")) = nvl(mrsBalance!���㷽ʽ)
''            .TextMatrix(i, .ColIndex("���")) = Format(mrsBalance!������, mstrDec)
''            .TextMatrix(i, .ColIndex("�������")) = nvl(mrsBalance!�������)
''            .TextMatrix(i, .ColIndex("��ע")) = nvl(mrsBalance!��ע)
''            .TextMatrix(i, .ColIndex("����")) = nvl(mrsBalance!����)
''            .TextMatrix(i, .ColIndex("������ˮ��")) = nvl(mrsBalance!������ˮ��)
''            .TextMatrix(i, .ColIndex("����˵��")) = nvl(mrsBalance!����˵��)
''            .TextMatrix(i, .ColIndex("���������")) = nvl(mrsBalance!���������)
''            .Row = i: .Col = .ColIndex("֧����ʽ")
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
    '����:��ȡ�����
    '����:�����:=-1��ʾδ�ҵ������
    '����:���˺�
    '����:2015-05-15 17:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��������"))) = 9 Then
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
Private Sub LoadBalanceInfo(ByVal lng����ID As Long, Optional ByVal blnNOMoved As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ�����Ϣ
    '���:lng����ID-����ID
    '     blnNOMoved-�Ƿ��Ѿ�ת������ʷ����
    '����:���˺�
    '����:2015-05-15 15:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    vsBlance.Rows = 2: vsBlance.Clear 1
    If lng����ID = 0 Then Exit Sub
    
    '       �ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '       ����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    
    Set mrsBalance = zlFromIDGetChargeBalance(0, lng����ID, blnNOMoved)
    With vsBlance
        .Redraw = flexRDBuffered
        .Rows = 2: .Clear 1
        .Rows = IIf(mrsBalance.RecordCount + 1 = 1, 2, mrsBalance.RecordCount + 1)
        If mrsBalance.RecordCount > 0 Then mrsBalance.MoveFirst
        For i = 1 To mrsBalance.RecordCount
            '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
            .TextMatrix(i, .ColIndex("�༭״̬")) = ""
            ''�Ƿ��ѽ���:1-�ѽ���;0-δ����
            .TextMatrix(i, .ColIndex("����״̬")) = 0
            
            .TextMatrix(i, .ColIndex("�����ID")) = Val(NVL(mrsBalance!�����ID))
            .TextMatrix(i, .ColIndex("���ѿ�ID")) = Val(NVL(mrsBalance!���ѿ�ID))
            .TextMatrix(i, .ColIndex("��������")) = Val(NVL(mrsBalance!��������))
            .TextMatrix(i, .ColIndex("����")) = Val(NVL(mrsBalance!����))
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsBalance!�Ƿ�����))
            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(mrsBalance!�Ƿ�ȫ��))
            .TextMatrix(i, .ColIndex("У�Ա�־")) = Val(NVL(mrsBalance!У�Ա�־))
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsBalance!�Ƿ�����))
            .TextMatrix(i, .ColIndex("֧����ʽ")) = NVL(mrsBalance!���㷽ʽ)
            .TextMatrix(i, .ColIndex("���")) = Format(mrsBalance!��Ԥ��, mstrDec)
            .TextMatrix(i, .ColIndex("�������")) = NVL(mrsBalance!�������)
            .TextMatrix(i, .ColIndex("��ע")) = NVL(mrsBalance!ժҪ)
            .TextMatrix(i, .ColIndex("����")) = NVL(mrsBalance!����)
            .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(mrsBalance!������ˮ��)
            .TextMatrix(i, .ColIndex("����˵��")) = NVL(mrsBalance!����˵��)
            .TextMatrix(i, .ColIndex("���������")) = NVL(mrsBalance!���������)
            .Row = i: .Col = .ColIndex("֧����ʽ")
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
    If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible _
        And picCurBalance.Enabled And picBalance.Enabled Then txtBalance(Idx_�ɿ�).SetFocus
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
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "�����б�"
End Sub
 Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "�����б�"
End Sub
 
Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub
Private Sub vsBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblCurMoney As Double, dbl�����ʻ� As Double, dblҽ������ As Double
    With vsBlance
        Select Case Col
        Case .ColIndex("֧����ʽ")
        Case .ColIndex("���")
        Case Else
        End Select
    End With
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytInState = 1 Then Cancel = True: Exit Sub
    With vsBlance
        .ComboList = ""
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        Select Case Val(.TextMatrix(Row, .ColIndex("�༭״̬")))
        Case 0: Cancel = True: Exit Sub
        Case 1
            If Col <> .ColKey("���") Then Cancel = True: Exit Sub
        Case 2
            If Col = .ColIndex("֧����ʽ") Then
                 .ComboList = "..."
                 .CellButtonPicture = imgDel
            Else
                Cancel = True: Exit Sub
            End If
        End Select
    End With
End Sub
Private Sub vsPati_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsPati, Me.Name, "�����б�"
End Sub
Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If mbytInState = 1 Then Exit Sub
    Call DeletePayInfor(Row)
End Sub

Private Sub vsBlance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If mbytInState = 1 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("���") Then Exit Sub
    End With
    Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m���ʽ)
End Sub

Private Sub vsBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyDelete Then Exit Sub
    With vsBlance
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(.Row, .ColIndex("�༭״̬"))) = 2 Then
            Call DeletePayInfor(.Row)
        End If
    End With
End Sub

Private Sub vsBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim dblԭʼ��� As Currency
    Dim dblTotal As Double, arrValue As Variant
    Dim i As Integer, str���㷽ʽ As String
    Dim varData As Variant
     
    If mbytInState = 1 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("���") Then Exit Sub
        If Row <= 0 Then Exit Sub
        Cancel = True
    End With
End Sub
Private Sub DeletePayInfor(ByVal lngDelRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��֧����Ϣ
    '����:���˺�
    '����:2015-01-28 15:18:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngRow As Long
    On Error GoTo errHandle
    With vsBlance
    
        dblMoney = Val(.TextMatrix(lngDelRow, .ColIndex("���")))
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(lngDelRow, .ColIndex("�༭״̬"))) <> 2 Then Exit Sub
        
        lngRow = lngDelRow
        mtyBalanceInfor.dblδ���ϼ� = FormatEx(mtyBalanceInfor.dblδ���ϼ� + dblMoney, 6)
        mtyBalanceInfor.dbl�Ѹ��ϼ� = FormatEx(mtyBalanceInfor.dbl�Ѹ��ϼ� - dblMoney, 6)
        
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
        If .Col <> .ColIndex("���") Then Exit Sub
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(.Row, .ColIndex("�༭״̬"))) <> 1 Then Exit Sub
        .EditCell
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsPati_DblClick()
    Dim lng����ID As Long
    
    lng����ID = Val(vsPati.RowData(vsPati.Row))
    If lng����ID <> 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me, _
            "����ID=" & lng����ID, "����ID=" & mlng����ID, "ReportFormat=" & mbytInState + 1, 0)
    End If
End Sub
    

Private Sub txtUnit_GotFocus()
    Call OpenIme(gstrIme)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If cmd����.Visible Then Call cmd����_Click
    End If
End Sub

Private Sub txtUnit_LostFocus()
    Call OpenIme
End Sub
Private Function Init���㷽ʽ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:���˺�
    '����:2015-01-08 12:06:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim objCards As Cards, objCard As Card
    Dim objPayCards As Cards, i As Long
    
    On Error GoTo errHandle
    mstrErrorBalance = ""
    If mbytInState = 1 Then Init���㷽ʽ = True: Exit Function
    
    
    Set objCards = New Cards: Set objPayCards = New Cards
    '����:1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���, _
    '     6-�����ۿ�,7-һ��ͨ����,8-���㿨����
 
    Set mrs���㷽ʽ = GetPayKind
    If mrs���㷽ʽ.RecordCount = 0 Then
        MsgBox "δ���ý��ʳ��Ͽ��õĽ��㷽ʽ��", vbInformation, gstrSysName
        mblnUnload = True
        Exit Function
    End If
     
     
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare Is Nothing Then
        '0-����ҽ�ƿ�;1-���õ�ҽ�ƿ�,2-���д��������˻��������� 3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)    '��ȡ��Ч�������ʻ��Ḷ
    End If

    mrs���㷽ʽ.Filter = "����<7 or ����=9"
    With mrs���㷽ʽ
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
        Do While Not .EOF
            If InStr(",3,4,9,", "," & Val(NVL(!����)) & ",") = 0 Then
                Set objCard = New Card
                objCard.�ӿ���� = -1 * i
                objCard.�ӿڱ��� = !����
                objCard.���� = !����
                objCard.���㷽ʽ = !����
                objCard.�������� = Val(NVL(!����))
                objCard.���� = True
                '85565,���ϴ�,2015/7/19:��������
                objCard.�Ƿ�ˢ�� = True
                objCard.ȱʡ��־ = Val(NVL(!ȱʡ)) = 1
                objPayCards.Add objCard
                If objCard.ȱʡ��־ Then
                    mstrȱʡ���㷽ʽ = objCard.���㷽ʽ
                End If
                i = i + 1
            ElseIf Val(NVL(!����)) = 9 Then
                mstrErrorBalance = NVL(!����)
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    mrs���㷽ʽ.Filter = "����>=7 and ����<9" 'һ��ͨ����
    With mrs���㷽ʽ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            For Each objCard In objCards
                If objCard.���㷽ʽ = NVL(!����) Then
                    '�ҵ���,����
                    '85565,���ϴ�,2015/7/19:��������
                    objCard.�Ƿ�ˢ�� = True
                    objCard.ȱʡ��־ = Val(NVL(!ȱʡ)) = 1
                    objPayCards.Add objCard
                    If objCard.ȱʡ��־ Then
                        mstrȱʡ���㷽ʽ = objCard.���㷽ʽ
                    End If
                End If
            Next
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    mrs���㷽ʽ.Filter = 0
    mblnNotChange = True
    Set IDKindPaymentsType.Cards = objPayCards
    If objPayCards.Count = 0 Then
        mblnNotChange = True
        MsgBox "���ʳ���û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnload = True: Exit Function
    End If
    mblnNotChange = False
    Init���㷽ʽ = True
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
    
    lbl�Ը��ϼ�.Caption = Format(0, mstrDec)
    lblʣ���Ը�.Caption = Format(0, mstrDec)
    txtBalance(Idx_lbl�ɿ�).Text = ""
    txtBalance(Idx_lbl�Ҳ�).Text = ""
    txtBalance(Idx_�������).Text = ""
    txtBalance(Idx_ժҪ).Text = ""
    
    Call ClearPayInfo
    staThis.Panels(2).Text = ""
    

    'Ʊ�ݺ��뵥�ݺŴ���
    Call RefreshFact
    If Visible And txtUnit.Enabled Then txtUnit.SetFocus
End Sub
 
Private Sub InitPatiGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������ͷ��Ϣ
    '����:���˺�
    '����:2015-05-04 17:33:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsPati
        .Redraw = flexRDNone: .Clear
        .Rows = 2: .Cols = 5: i = 0
        .TextMatrix(0, i) = "����ID": i = i + 1
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "�Ա�": i = i + 1
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "���ʽ��": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*���" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "�Ա�" Or .ColKey(i) = "����" Then
                .ColAlignment(i) = flexAlignCenterCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(0) = 0
            End If
        Next
        .ColWidth(.ColIndex("����")) = 1000
        .ColWidth(.ColIndex("�Ա�")) = 650
        .ColWidth(.ColIndex("����")) = 650
        .ColWidth(.ColIndex("���ʽ��")) = 1450
        .RowHeight(0) = 320
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsPati, Me.Name, "�����б�"
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
    txtBalance(Idx_�ɿ�).Text = Format(0, mstrDec)
    txtBalance(Idx_�Ҳ�).Text = Format(0, mstrDec)
End Sub
Private Sub ShowBalanceInfo(ByVal rsTmp As ADODB.Recordset)
   Dim curTotal As Currency, lngMaxLength As Long, lngP As Long, i As Long
    
    Call InitPatiGrid
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.State <> 1 Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    
    lngMaxLength = Len(Mid(gstrDec, 3))
    For i = 1 To rsTmp.RecordCount
        lngP = InStr(1, CStr(rsTmp!���ʽ��), ".")
        If lngP > 0 Then
            lngP = Len(Mid(CStr(rsTmp!���ʽ��), lngP + 1))
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
            .RowData(i) = Val(rsTmp!����ID)
            .TextMatrix(i, .ColIndex("����ID")) = Val(NVL(rsTmp!����ID))
            .TextMatrix(i, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(i, .ColIndex("�Ա�")) = "" & rsTmp!�Ա�
            .TextMatrix(i, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(i, .ColIndex("���ʽ��")) = Format(rsTmp!���ʽ��, mstrDec)
            .Row = i: .Col = .ColIndex("���ʽ��")
            .CellBackColor = 12900351
            curTotal = curTotal + rsTmp!���ʽ��
            rsTmp.MoveNext
        Next
        vsPati.Redraw = flexRDBuffered
    End With
    lbl�Ը��ϼ�.Caption = Format(curTotal, mstrDec)
    With mtyBalanceInfor
        .dbl��ǰ���� = curTotal
        .dblδ���ϼ� = curTotal
    End With
End Sub

Private Function GetPayKind() As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "" & _
    "   Select a.����,a.����, a.ȱʡ��־ ȱʡ, a.����, 1 As λ��" & vbNewLine & _
    "   From ���㷽ʽ a, ���㷽ʽӦ�� b" & vbNewLine & _
    "   Where a.���� = b.���㷽ʽ And b.Ӧ�ó��� = '����' And a.���� Not In (3, 4)  " & _
    "       And Nvl(a.Ӧ����, 0) = 0 And Nvl(a.Ӧ�տ�, 0) = 0" & vbNewLine & _
    "   Union " & _
    "   Select ����,����, ȱʡ��־ As ȱʡ, ����,0 As λ��" & _
    "   From ���㷽ʽ " & _
    "   Where ����=9 " & _
    "Order By λ��,����,����"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set GetPayKind = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetBalanceData(ByVal lng����ID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strWhere As String
    
    strWhere = " And B.��¼״̬ In(1,3)"
    If mblnViewCancel Then strWhere = " And B.��¼״̬ =2"
    
    strSql = "" & _
    " Select A.����ID,nvl(A.����,C.����) as ����, nvl(A.�Ա�,C.�Ա�) as �Ա�,  " & _
    "       nvl(A.����,C.����) as ����, Nvl(Sum(A.���ʽ��),0) ���ʽ��" & vbNewLine & _
    " From ������ü�¼ A,���˽��ʼ�¼ B,������Ϣ C" & vbNewLine & _
    " Where A.����id = [1] And A.����id = B.id " & strWhere & _
    "        And A.����id=C.����id " & vbNewLine & _
    " Group By A.����ID,nvl(A.����,C.����), nvl(A.�Ա�,C.�Ա�), nvl(A.����,C.����) "
    
    If blnHistory Then strSql = Replace(Replace(strSql, "������ü�¼", "H������ü�¼"), "���˽��ʼ�¼", "H���˽��ʼ�¼")
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    Set GetBalanceData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPayData(ByVal lng����ID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strWhere As String
    strWhere = " And B.��¼״̬ In(1,3)"
    If mblnViewCancel Then strWhere = " And B.��¼״̬ =2"

    strSql = "" & _
    "   Select A.���㷽ʽ, A.��Ԥ�� ������, A.�������, A.ժҪ ��ע, " & vbNewLine & _
    "           A.�����ID,A.���㿨���,A.����,A.������ˮ��,A.����˵��" & vbNewLine & _
    "   From ����Ԥ����¼ A,���˽��ʼ�¼ B" & vbNewLine & _
    "   Where A.����id = [1] And A.����id = B.id " & strWhere
    If blnHistory Then strSql = Replace(Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼"), "���˽��ʼ�¼", "H���˽��ʼ�¼")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    Set GetPayData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetOtherInfo(ByVal lng����ID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String, strWhere As String
    Dim strTable As String, strTable1 As String, strTable2 As String
   
    strWhere = " And B.��¼״̬ In(1,3)"
    If mblnViewCancel Then strWhere = " And B.��¼״̬ =2"
    
    strTable1 = "" & _
    "   Select Min(D.����) ��Լ��λ  " & _
    "   From ������ü�¼ A,���˽��ʼ�¼ B, ������Ϣ C, ��Լ��λ D" & vbNewLine & _
    "   Where A.����id = [1] And A.����id = B.id  " & strWhere & _
    "         And A.����id = C.����id And C.��ͬ��λid = D.ID"
    
    strTable2 = "" & _
    "   Select A.NO, A.ʵ��Ʊ�� " & _
    "   From ���˽��ʼ�¼ A " & _
    "   Where A.id = [1]  " & Replace(strWhere, "B.", "A.")
    
    strSql = "" & _
    " Select   B.��Լ��λ, C.NO, C.ʵ��Ʊ��" & vbNewLine & _
    " From (" & strTable1 & ") B," & vbNewLine & _
    "      (" & strTable2 & ") C"
  
    If blnHistory Then strSql = Replace(Replace(strSql, "������ü�¼", "H������ü�¼"), "���˽��ʼ�¼", "H���˽��ʼ�¼")
    If blnHistory Then strSql = Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼")
 
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    
    Set GetOtherInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitGrid_PayList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��֧���б�
    '����:���˺�
    '����:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
        .Clear: .Rows = 2: i = 0: .Cols = 18
        .TextMatrix(0, i) = "�����ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���ѿ�ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "��������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�༭״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "У�Ա�־": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "֧����ʽ": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "���": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "�������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "������ˮ��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����˵��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "��ע": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "���������": .ColWidth(i) = 0: i = i + 1
        
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            Select Case .ColKey(i)
            Case "��������", "����", "�Ƿ񱣴�", "�Ƿ�����", "У�Ա�־", "�༭״̬", "�Ƿ�����", "�Ƿ�ȫ��", "���������", "����״̬", "�Ƿ���֤"
                .ColHidden(i) = True
            Case "���"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
        zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "�����б�"
        If mbytInState = 0 Then '���ʲ���
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
    Case Idx_�ɿ�
         Call Set�Ҳ���Ϣ
    Case Idx_ժҪ
    Case Else
    End Select
End Sub

Private Sub txtBalance_GotFocus(Index As Integer)
    Select Case Index
    Case Idx_�ɿ�
      '  Call LedVoiceSpeak(True)
      '  txtBalance(Index).Text = ""
        
    Case Idx_ժҪ
        zlCommFun.OpenIme True
    End Select
    zlControl.TxtSelAll txtBalance(Index)
End Sub
Private Sub txtBalance_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim dblMoney As Double, blnChargeEnd As Boolean
    Dim objCard As Card, objKind As IDKindNew
 
    If KeyAscii <> 13 Then
       If Index <> Idx_�ɿ� Then Exit Sub
       Set objKind = IDKindPaymentsType
       Call MoveIDKindItem(objKind, KeyAscii)
       Exit Sub
    End If
    
    KeyAscii = 0
    Select Case Index
    Case Idx_�ɿ�
        dblMoney = FormatEx(Val(txtBalance(Index).Text), 6)
        Set objCard = IDKindPaymentsType.GetCurCard
        If objCard Is Nothing Then zlCommFun.PressKey vbKeyTab: Exit Sub
        Select Case mty_ModulePara.byt�ɿ��������
        Case 2   '�����˽ɿ��ۼ�
            If objCard.�������� = 1 Then '�ֽ�
                If txtBalance(Index).Text = "" Then
                    cmdOK.SetFocus: Exit Sub
                End If
                If SaveBalanceData = False Then Exit Sub
                
            ElseIf objCard.�������� = 2 Then '��ҽ������
                If txtBalance(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
            ElseIf objCard.�ӿ���� > 0 Then
                If dblMoney = 0 Then
                    MsgBox "δ����ɿ���,����ʹ�á�" & objCard.���㷽ʽ & "�����н���", vbInformation, gstrSysName
                    Exit Sub
                End If
                If SaveBalanceData = False Then Exit Sub
            End If
            Exit Sub
        Case Else '0-�����нɿ����,1-�����ֽ�ʱ,��������ɿ���
            If objCard.�������� = 1 Then Call cmdOK_Click: Exit Sub
        End Select
        
        If objCard.�ӿ���� > 0 Then
            If dblMoney = 0 Then
                MsgBox "δ����ɿ���,����ʹ�á�" & objCard.���㷽ʽ & "�����н���", vbInformation, gstrSysName
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
    Case Idx_�ɿ�
        txtBalance(Index).Text = Format(Val(txtBalance(Index).Text), "0.00")
    Case Idx_ժҪ
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtBalance_Validate(Index As Integer, Cancel As Boolean)
    Dim dblMoney As Double, dbl�Ҳ� As Double
    Dim intSign As Integer
    Select Case Index
    Case Idx_�ɿ�
    Case Else
    End Select
End Sub
Private Sub Set�Ҳ���Ϣ()
    Dim dblMoney As Double
    Dim dbl��ǰδ�� As Double
    Dim objCard As Card
    Dim objBackCard As Card
    Dim objCards As Cards
    Dim objTemp As Card
    dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
    Set objCard = IDKindPaymentsType.GetCurCard
    
    dbl��ǰδ�� = mtyBalanceInfor.dbl��ǰ���� - mtyBalanceInfor.dbl�Ѹ��ϼ�
    If dblMoney = 0 Or objCard Is Nothing Then
        txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_�������).ForeColor
        txtBalance(Idx_�Ҳ�).Text = "0.00"
        Exit Sub
    End If
        
    If dbl��ǰδ�� < 0 Then
        '��ǰ״̬Ϊ�˿�
        dbl��ǰδ�� = FormatEx(Val(lblʣ���Ը�.Caption), 6)
        If objCard.�������� = 1 Then
            'ֻ���ֽ�,�Ż����˿�ʱ�ึ������,����:��100������,Ҫ�һ�50
            If Abs(dbl��ǰδ��) <= dblMoney Then
                txtBalance(Idx_�Ҳ�).Text = Format(dblMoney - Abs(dbl��ǰδ��), "0.00")
                lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
                txtBalance(Idx_�Ҳ�).ForeColor = vbRed
            Else
                txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
                lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
                txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_�������).ForeColor
            End If
            Exit Sub
        End If
        
        If Abs(dbl��ǰδ��) < dblMoney Then
            '�������㷽ʽ��,ֻ����ʣ��δ�˿�,����:�˿���ҽԺ��֧Ʊ������,��˲������һ�֧Ʊ�Ŀ���
            lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
            txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_�������).ForeColor
            txtBalance(Idx_�Ҳ�).Text = "0.00": Exit Sub
        Else
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
            txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_�������).ForeColor
        End If
        Exit Sub
    End If
    
    '��ǰ״̬Ϊ�տ�
    dbl��ǰδ�� = FormatEx(Val(lblʣ���Ը�.Caption), 6)
    If objCard.�������� = 1 Then
        'ֻ���ֽ�,�Ż����˿�ʱ�ึ������,����:��100������,Ҫ�һ�50
        If dbl��ǰδ�� >= dblMoney Then
            '��Ҫ��ȡ����Ǯ
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
            txtBalance(Idx_�Ҳ�).ForeColor = vbRed
        Else
            '�˿�
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
            txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_�������).ForeColor
        End If
        Exit Sub
    End If
    
    If dbl��ǰδ�� >= dblMoney Then
        'Ҫ�տ�
        lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
        txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
        txtBalance(Idx_�Ҳ�).ForeColor = vbRed
        Exit Sub
    Else
        If objCard.�������� = 2 And objCard.���㷽ʽ Like "*֧Ʊ" Then
            lblBalance(Idx_lbl�Ҳ�).Caption = "�� ֧ Ʊ"
        Else
            lblBalance(Idx_lbl�Ҳ�).Caption = "��    ��"
        End If
        txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
        txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_�������).ForeColor
    End If
End Sub

Private Function zlCheckMulitInterfaceNumValied(Optional blnԤ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͬʱ�����������Ͻӿ�(��������)
    '����:�����������Ͻӿڵ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int���� As Integer, str���㷽ʽ As String
    Dim varData As Variant, strErrMsg As String
    Dim objCard As Card
    Dim intMousePointer As Integer
    On Error GoTo errHandle
    strErrMsg = ""
    intMousePointer = Screen.MousePointer
    Set objCard = IDKindPaymentsType.GetCurCard
        
    If blnԤ�� Or objCard.�ӿ���� <= 0 Then zlCheckMulitInterfaceNumValied = True:        Exit Function
    
   With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If InStr("34", int����) > 0 Then
                intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str���㷽ʽ & ":" & .TextMatrix(i, .ColIndex("���"))
            End If
        Next
    End With
    If intCount > 1 Then
        Screen.MousePointer = 0
        Call MsgBox("ע��:" & vbCrLf & "   ��ϵͳĿǰֻ֧��һ�ֽӿ�,�����Ѿ��������½ӿڽ���:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
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
    '����:���ѿ����㽻�׼��
    '���:objCard-������
    '����:dblMoney-��ǰˢ�����
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ���ӿڵ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl�ʻ���� As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln���� As Boolean, dblδ����� As Double
    Dim intMousePointer As Integer, strXmlIn As String
    Dim lng���ѿ�ID As Long, str���� As String, str���� As String
    Dim str������� As String, byt�Ƿ�����   As Byte
    Dim cllBushSquare As Collection, i As Long
    
    
    intMousePointer = Screen.MousePointer

    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� = False Then CheckSquareBalanceValied = True: Exit Function
    
    On Error GoTo errHandle
    
    tyBrushCard = strBrushCard
    
    dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
    dblδ����� = FormatEx(mtyBalanceInfor.dblδ���ϼ�, 6)
    
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "�տ���δ����,����!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If dblMoney > Format(dblδ�����, "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox "�տ���ܴ��ڱ���δ�����:" & Format(dblδ�����, "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ȼ���Ӧ�Ľӿ�
    If zlGetClassMoney(0, rsMoney) = False Then Exit Function
    
     '�������ѿ���ˢ����Ϣ
     Set cllSquareBalance = New Collection
     Set mcllCurSquareBalance = New Collection
     With vsBlance
        For i = 1 To .Rows - 1
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
            '����״̬:�Ƿ��ѽ���:1-�ѽ���;0-δ����
            If Val(.TextMatrix(i, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(i, .ColIndex("�����ID"))) = objCard.�ӿ���� _
                And Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
  
                dblTemp = FormatEx(Val(.TextMatrix(i, .ColIndex("���"))), 6)
                lng���ѿ�ID = Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                str���� = Trim(.Cell(flexcpData, i, .ColIndex("����")))
                str���� = Trim(.Cell(flexcpData, i, .ColIndex("���ѿ�ID")))  '����
                str������� = Trim(.Cell(flexcpData, i, .ColIndex("�����ID")))  '�������
                byt�Ƿ����� = Val(.TextMatrix(i, .ColIndex("�Ƿ�����")))
                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                cllSquareBalance.Add Array(objCard.�ӿ����, lng���ѿ�ID, dblTemp, str����, str����, str�������, byt�Ƿ�����)
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
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant) As Boolean
    'varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
    strXmlIn = "<IN><CZLX>0</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, _
            objCard.�ӿ����, objCard.���ѿ�, _
            "" & txtUnit.Text, "" & "", "" & "", dblMoney, _
            tyBrushCard.str����, tyBrushCard.str����, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
       
        For i = 1 To cllSquareBalance.Count
           mcllCurSquareBalance.Add cllSquareBalance(i)
        Next
         
        '����ǰ,һЩ���ݼ��
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNOs As String, _
        Optional ByVal strXMLExpend As String
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.�ӿ����, _
            objCard.���ѿ�, tyBrushCard.str����, dblMoney, "", strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
        '���:frmMain-���õ�������
        '        lngModule-ģ���
        '        strCardNo-����
        '        strExpand-Ԥ����Ϊ��,�Ժ���չ
        '����:dblMoney-�����ʻ����
        'If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
              tyBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�) = False Then Exit Function
        '�Ѿ������˽�����
        If FormatEx(dblMoney, 6) <> Val(txtBalance(Idx_�ɿ�).Text) Then
            txtBalance(Idx_�ɿ�).Text = Format(dblMoney, "0.00")
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
    Optional ByVal bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������֤
    '���:objCard-������
    '     dblMoney-ˢ�����,>=0��ʾ�տ�;С�����ʾ�˿�
    '     bln�˿�-true,��ʾ��ǰΪ�˿���;False��ʾ��ǰΪ�տ���
    '����:tyBrushCard-ˢ����Ϣ
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ���ӿڵ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, cllSquareBalance As Collection
    Dim strXMLExpend As String, bln���� As Boolean
    Dim dbl�ʻ���� As Double, dblδ����� As Double
    Dim strExpand As String, strXmlIn As String
    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim blnCurInput As Boolean
    
    intMousePointer = Screen.MousePointer
    
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then CheckThreeSwapValied = True: Exit Function
    
    
    On Error GoTo errHandle
    
    tyBrushCard.blnת�� = False
    If dblMoney = 0 Then dblMoney = Val(txtBalance(Idx_�ɿ�).Text): blnCurInput = True
     
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "���δ����,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
     
    dblδ����� = FormatEx(mtyBalanceInfor.dblδ���ϼ�, 6)
    If Abs(dblMoney) > Format(Abs(dblδ�����), "0.00") And dblMoney <> 0 And blnCurInput = False Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "ˢ��") & "���ܴ��ڱ���" & IIf(bln�˿�, "δ��", "δ��") & "���:" & Format(Abs(dblδ�����), "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Val(lblʣ���Ը�.Caption) <> dblMoney Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "ˢ��") & "����:" & Format(Abs(Val(lblʣ���Ը�.Caption)), "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not bln�˿� Then
        'zlBrushCard(frmMain As Object, _
           ByVal lngModule As Long, _
           ByVal rsClassMoney As ADODB.Recordset, _
           ByVal lngCardTypeID As Long, _
           ByVal bln���ѿ� As Boolean, _
           ByVal strPatiName As String, ByVal strSex As String, _
           ByVal strOld As String, ByRef dbl��� As Double, _
           Optional ByRef strCardNo As String, _
           Optional ByRef strPassWord As String, _
           Optional ByRef bln�˷� As Boolean = False, _
           Optional ByRef blnShowPatiInfor As Boolean = False, _
           Optional ByRef bln���� As Boolean = False, _
           Optional ByVal bln�����ֹ As Boolean = True, _
           Optional ByRef varSquareBalance As Variant) As Boolean
           '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
            objCard.�ӿ����, objCard.���ѿ�, _
            txtUnit.Text, "", "", dblMoney, _
            tyBrushCard.str����, tyBrushCard.str����, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
            '����ǰ,һЩ���ݼ��
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal strCardTypeID As Long, ByVal strCardNo As String, _
            ByVal dblMoney As Double, ByVal strNOs As String, _
            Optional ByVal strXMLExpend As String
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.�ӿ����, _
            objCard.���ѿ�, tyBrushCard.str����, dblMoney, "", strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
          '���:frmMain-���õ�������
          '        lngModule-ģ���
          '        strCardNo-����
          '        strExpand-Ԥ����Ϊ��,�Ժ���չ
          '����:dblMoney-�����ʻ����
        If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
              tyBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�) = False Then Exit Function
        
        staThis.Panels(2).Text = Format(dbl�ʻ����, "0.00")
        staThis.Panels(2).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
        tyBrushCard.dbl�ʻ���� = FormatEx(dbl�ʻ����, 2)
        If dbl�ʻ���� <> 0 And dbl�ʻ���� < dblMoney Then
            Screen.MousePointer = 0
            MsgBox objCard.���㷽ʽ & "���ʻ�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        CheckThreeSwapValied = True
        Exit Function
    End If
    
    '�˿���
    If mrsBalance Is Nothing Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    

    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    mrsBalance.Filter = "����=3 And �����ID=" & objCard.�ӿ����
    If mrsBalance.EOF Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���㷽ʽ & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    dblTemp = 0
    With mrsBalance
        Do While Not .EOF
            dblTemp = dblTemp + Val(NVL(!��Ԥ��))
            .MoveNext
        Loop
        mrsBalance.MoveFirst
        dblTemp = FormatEx(dblTemp, 5)
    End With
    
    If dblTemp = 0 Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & objCard.���㷽ʽ & "�Ѿ����꣬�������ˣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objCard.�Ƿ�ȫ�� And Not objCard.�Ƿ����� Then
        If dblTemp <> dblMoney Then
            If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "ע��:" & vbCrLf & objCard.���� & "�����˿�ʱ������ȫ�ˣ�" & vbCrLf & _
            "  ʣ��δ��:" & Format(Abs(dblTemp), "0.00") & vbCrLf & _
            "  ��ǰ���:" & Format(Abs(dblMoney), "0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln���ѿ� As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:�ʻ����˽���ǰ�ļ��
        '���:frmMain-���õ�������
        '       lngModule-���õ�ģ���
        '       lngCardTypeID-�����ID
        '       strCardNo-����
        '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
        '       dblMoney-�˿���
        '       strSwapNo-������ˮ��(�˿�ʱ���)
        '       strSwapMemo-����˵��(�˿�ʱ����)
        '       strXMLExpend    XML IN  ��ѡ����:�쳣���������˷�(1)
        '����:�˿�Ϸ�,����true,���򷵻�Flase
        
    strXMLExpend = ""
    tyBrushCard.str���� = NVL(mrsBalance!����)
    tyBrushCard.str������ˮ�� = NVL(mrsBalance!������ˮ��)
    tyBrushCard.str����˵�� = NVL(mrsBalance!����˵��)

    strBalanceIDs = "2|" & mtyBalanceInfor.lng����ID & IIf(mtyBalanceInfor.lng����ID = 0, "", "," & mtyBalanceInfor.lng����ID)
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, _
        strBalanceIDs, dblMoney, tyBrushCard.str������ˮ��, tyBrushCard.str����˵��, strXMLExpend) = False Then Exit Function
                
    If objCard.�Ƿ��˿��鿨 Then
       '����ˢ������
        'zlBrushCard(frmMain As Object, _
        'ByVal lngModule As Long, _
        'ByVal rsClassMoney As ADODB.Recordset, _
        'ByVal lngCardTypeID As Long, _
        'ByVal bln���ѿ� As Boolean, _
        'ByVal strPatiName As String, ByVal strSex As String, _
        'ByVal strOld As String, ByVal dbl��� As Double, _
        'Optional ByRef strCardNo As String, _
        'Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean) As Boolean
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.�ӿ����, _
            objCard.���ѿ�, txtUnit.Text, "", _
            "", dblMoney, tyBrushCard.str����, tyBrushCard.str����, _
            True, True, bln����, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    End If
    CheckThreeSwapValied = True
    Exit Function
    
GoTransferAccount:
        strXmlIn = "<IN><CZLX>1</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.�ӿ����, _
            objCard.���ѿ�, txtUnit.Text, "", _
            "", dblMoney, tyBrushCard.str����, tyBrushCard.str����, _
            True, True, bln����, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    
    tyBrushCard.blnת�� = True
    '����ת�ʽӿ�
    '    7.1.    zltransferAccountsCheck(ת�ʼ��ӿ�)
    'zlTransferAccountsCheck ת�ʼ��ӿ�
    '������  ��������    ��/��   ��ע
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  HIS����ģ���
    'lngCardTypeID   Long    In  �����ID
    'strCardNo   String  In  ����
    'dblMoney    Double  In  ת�ʽ��(����ʱΪ����)
    'strBalanceIDs   String  In  ����IDs������ö��ŷ��룬��ʾ���ζ��Ĵ��շ���Ŀ��������ҽ��������
    'strXMLExpend String In   XML��:
    '                            <IN>
    '                                <CZLX >��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��
    '                            </IN>
    '                    Out  XML��:
    '                            <OUT>
    '                               <ERRMSG>������Ϣ</ERRMSG >
    '                            </OUT>
    '    Boolean ��������    �������ݺϷ�,����True:���򷵻�False
    '˵��:
    '��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
    '��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
    '����XML��
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.�ӿ����, _
        tyBrushCard.str����, dblMoney, mtyBalanceInfor.lng����ID, strXMLExpend) = False Then
        Screen.MousePointer = 0
        Call zlShowThreeSwapErrInfor(0, strXMLExpend)
        Exit Function
    End If
    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
          tyBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�)
    If dbl�ʻ���� <> 0 Then
        staThis.Panels(2).Text = objCard.���㷽ʽ & "�ʻ����:" & Format(dbl�ʻ����, "0.00")
        staThis.Panels(2).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
    End If
    tyBrushCard.dbl�ʻ���� = FormatEx(dbl�ʻ����, 2)
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



Private Function zlGetClassMoney(ByRef lng����ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, dblMoney As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    '��ʼ�����ݽṹ
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "���", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng����ID <> 0 Then
        strSql = "" & _
        "   Select  A.�շ����,nvl(sum(A.���ʽ��) ,0) as ���   " & _
        "   From ������ü�¼ A" & _
        "   Where A.����ID=[1] Group by A.�շ���� " & _
        "   Union ALL " & _
        "   Select  A.�շ����,nvl(sum(A.���ʽ��) ,0) as ���   " & _
        "   From סԺ���ü�¼ A" & _
        "   Where A.����ID=[1] Group by A.�շ���� "
        strSql = "Select �շ����,Sum(���) as ��� From (" & strSql & ")  Group by  �շ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "�շ����='" & NVL(!�շ����, "��") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!�շ���� = NVL(!�շ����, "��")
                rsMoney!��� = Val(NVL(rsMoney!���)) + Val(NVL(!���))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        zlGetClassMoney = True
        Exit Function
    End If
 
    strSql = "" & _
    " Select A.�շ����, Nvl(Sum(ʵ�ս��), 0) - Nvl(Sum(���ʽ��), 0) As δ����" & vbNewLine & _
    " From ������ü�¼ A, ������Ϣ B" & vbNewLine & _
    " Where A.����id = B.����id   And A.��¼״̬ <> 0 And A.���ʷ��� = 1  " & _
    "       And A.�����־ IN(1,4) And B.��ͬ��λid = [1] And B.��ǰ����id Is Null " & _
    " Group By a.�շ����" & _
    " Having Nvl(Sum(a.ʵ�ս��), 0) - Nvl(Sum(a.���ʽ��), 0) <> 0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(txtUnit.Tag))
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            rsMoney.Find "�շ����='" & NVL(!�շ����, "��") & "'", , adSearchForward, 1
            If rsMoney.EOF Then rsMoney.AddNew
            rsMoney!�շ���� = NVL(!�շ����, "��")
            rsMoney!��� = Val(NVL(rsMoney!���)) + Val(NVL(!δ����))
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
    '����:��ʼ����һ��ͨ��Ϣ
    '����:���˺�
    '����:2015-01-08 12:02:17
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
    '����:���ݷֱҴ������,���طֱҴ����Ľ��
    '���:dblMoney-δ�����ԭʼ���
    '����:���طֱҴ����Ľ��
    '����:���˺�
    '����:2015-01-26 10:57:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then GetCentMoney = FormatEx(dblMoney, 2): Exit Function
    '���ֽ��,������λС��
    If objCard.�������� <> 1 Then GetCentMoney = FormatEx(dblMoney, 2): Exit Function
    GetCentMoney = CentMoney(CCur(dblMoney))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub Show�����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����
    '����:���˺�
    '����:2015-01-14 11:33:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl��֧���� As Double
    Dim dblʣ���� As Double, dblTemp As Double, dblδ���� As Double
    Dim intSign As Integer, objCard As Card
    If mbytInState = 1 Then Exit Sub
    
    With mtyBalanceInfor
        .dbl���� = 0
        dblδ���� = .dblδ���ϼ�
        intSign = IIf(dblδ���� < 0, -1, 1)
    End With
    
    dblMoney = FormatEx(intSign * Val(txtBalance(Idx_�ɿ�).Text), 6)
    
    dbl��֧���� = 0: dblʣ���� = FormatEx(dblδ���� - dblMoney, 6)
     
    Set objCard = IDKindPaymentsType.GetCurCard
    If Not objCard Is Nothing Then
        If objCard.�������� = 1 Then
            dblTemp = dblδ����: dblʣ���� = 0
            dblMoney = GetCentMoney(dblTemp)
            mtyBalanceInfor.dbl���� = FormatEx(dblδ���� - dblMoney, 6)
            GoTo Show���:
        End If
    End If
    mtyBalanceInfor.dbl���� = FormatEx(dblδ���� - FormatEx(dblδ����, 2), 6): GoTo Show���:
Show���:
    lbl����.Visible = mtyBalanceInfor.dbl���� <> 0 And mbytInState = 0
    lbl����.Caption = "���:" & FormatEx(mtyBalanceInfor.dbl����, 6)
End Sub


Private Function ExecuteOldOneCardPayInterface(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal objCard As Card, ByVal dblMoney As Double, tyBrushCardInfor As TY_BrushCard, _
    ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�ϰ汾)
    '���:lng�������-��������Ž��д���
    '     dblMoney-���ν�����
    '     TYBrushCardInfor-��ǰˢ����Ϣ
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 16:14:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, strҽԺ���� As String
    Dim i As Long, strSql As String, str���㷽ʽ As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intCardType As Integer, strSwapNO As String
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�������� <> 7 Then ExecuteOldOneCardPayInterface = True: Exit Function

    mOldOneCard.rsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mOldOneCard.rsOneCard.EOF Then
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        ExecuteOldOneCardPayInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    'һ��ͨ����
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If Not mobjICCard.PaymentSwap(dblMoney, dbl���, intCardType, Val("" & mOldOneCard.rsOneCard!ҽԺ����), tyBrushCardInfor.str����, tyBrushCardInfor.str������ˮ��, lng����ID, lng����ID) Then
        gcnOracle.RollbackTrans
        MsgBox objCard.���㷽ʽ & "����ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    strSql = "Zl_һ��ͨ����_Update(" & 0 & ",'" & objCard.���㷽ʽ & "','" & tyBrushCardInfor.str���� & "','" & intCardType & "','" & strSwapNO & "'," & dbl��� & ")"
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
Private Function ExecuteThreeSwapPayInterface(ByVal lng����ID As Long, ByVal lng����ID As Long, objCard As Card, ByVal dblMoney As Double, _
    ByRef cllBillPro As Collection, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:lng�������-��������Ž��д���
    '     dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '     tyBrushCard-��ǰˢ����Ϣ
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long, strSql As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str���㷽ʽ  As String
    
    Err = 0: On Error GoTo errHandle:
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-������
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    str����IDs = lng����ID
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, _
         str����IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    tyBrushCard.str������ˮ�� = strSwapGlideNO
    tyBrushCard.str����˵�� = strSwapMemo
    If objCard.���ѿ� = False Then
        Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
    End If
    Call zlAddThreeSwapSQLToCollection(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '��������������Ϣ
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
    '����:�������ѿ����㷽ʽ�����㷽ʽ�б�
    '����:���˺�
    '����:2015-01-23 15:09:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As New Collection
    Dim j As Integer, dblMoney As Double, strCardNo As String
    
    With vsBlance
      '�����ԭʼ�����ѿ�����,�������˷�
        Call ClearSquareBalance(objCard.�ӿ����)
        Set cllBalance = mcllCurSquareBalance
        For j = 1 To cllBalance.Count
            If objCard.�ӿ���� = Val(cllBalance(j)(0)) Then
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                dblMoney = cllBalance(j)(2)
            
                .TextMatrix(1, .ColIndex("����")) = 5
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = Val(cllBalance(j)(6))
                .TextMatrix(1, .ColIndex("��������")) = objCard.��������
                .TextMatrix(1, .ColIndex("�༭״̬")) = 2   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                .TextMatrix(1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .TextMatrix(1, .ColIndex("�����ID")) = objCard.�ӿ����
                .TextMatrix(1, .ColIndex("���ѿ�ID")) = Val(cllBalance(j)(1))
                .Cell(flexcpData, 1, .ColIndex("���ѿ�ID")) = cllBalance(j)(4)  '����
                .Cell(flexcpData, 1, .ColIndex("�����ID")) = cllBalance(j)(5)  '�������
                
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                 strCardNo = Trim(cllBalance(j)(3))
                .TextMatrix(1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "" And objCard.�������Ĺ��� <> "0", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = strCardNo
                .TextMatrix(1, .ColIndex("���")) = Format(dblMoney, "0.00")
                .Cell(flexcpData, 1, .ColIndex("���")) = Format(dblMoney, "0.00")
                .TextMatrix(1, .ColIndex("�������")) = ""
                .TextMatrix(1, .ColIndex("��ע")) = ""
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(1, .ColIndex("���������")) = objCard.����
                
                mtyBalanceInfor.dbl�Ѹ��ϼ� = FormatEx(mtyBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
                mtyBalanceInfor.dblδ���ϼ� = FormatEx(mtyBalanceInfor.dblδ���ϼ� - dblMoney, 6)
            End If
        Next
    End With
End Sub


Private Sub LoadCurOwnerPayInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�ǰ������Ϣ
    '����:���˺�
    '����:2015-01-12 14:14:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngColor As Long, objCard As Card
    Dim dblTemp As Double
    
    If mbytInState = 1 Then Exit Sub
    
    With mtyBalanceInfor
        dblMoney = .dblδ���ϼ�
        lbl�Ը��ϼ�.Caption = Format(.dbl��ǰ����, mstrDec)
        dblTemp = GetCentMoney(Abs(dblMoney))
        
        lblʣ���Ը�.Caption = Format(FormatEx(dblTemp, 6), mstrDec)
        lblʣ���Ը�.Tag = Format(dblTemp, mstrDec)
        
    
       stcTittile.Caption = IIf(dblMoney < 0, "��ǰδ��", "��ǰδ��")
       lblBalance(Idx_lbl�ɿ�).Caption = IIf(dblMoney < 0, "��    ��", "��    ��")
       '����������ʾ
        lngColor = IIf(dblMoney < 0, vbRed, vbBlue)
        lblʣ���Ը�.ForeColor = lngColor
        IDKindPaymentsType.ForeColor = lngColor
        lblBalance(Idx_lbl�ɿ�).ForeColor = lngColor
        txtBalance(Idx_�ɿ�).ForeColor = lngColor
    End With
    Show�����
End Sub
Private Sub MoveIDKindItem(ByVal objKind As IDKindNew, ByVal KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ƶ�IDKind��Ŀ
    '���:objKind-�ƶ���IDKind����
    '     Keyascii-��ֵ
    '����:���˺�
    '����:2015-01-29 15:22:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If objKind Is Nothing Then Exit Sub
    If Not (KeyAscii = Asc("+") Or KeyAscii = Asc("-")) Then Exit Sub
    If objKind.ListCount = 1 Then Exit Sub
    
    If KeyAscii = Asc("+") Then
        '����һ��
        If objKind.IDKIND + 1 > objKind.ListCount Then
            objKind.IDKIND = 1
        Else
            objKind.IDKIND = objKind.IDKIND + 1
        End If
        Exit Sub
    End If
    If KeyAscii = Asc("-") Then '����һ��
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
    Optional ByVal lng���ѿ�ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ѿ�����
    '����:���˺�
    '����:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("�༭״̬"))) = 2 _
                And Val(.TextMatrix(j, .ColIndex("�����ID"))) = lngCardTypeID _
                And (lng���ѿ�ID = 0 Or (lng���ѿ�ID <> 0 And Val(.TextMatrix(j, .ColIndex("���ѿ�ID"))) = lng���ѿ�ID)) Then
                dblMoney = Val(.Cell(flexcpData, j, .ColIndex("���")))
                mtyBalanceInfor.dbl�Ѹ��ϼ� = FormatEx(mtyBalanceInfor.dbl�Ѹ��ϼ� - dblMoney, 6)
                mtyBalanceInfor.dblδ���ϼ� = FormatEx(mtyBalanceInfor.dblδ���ϼ� + dblMoney, 6)
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
    '����:����ȱʡ�Ľɿ���˿���
    '����:���˺�
    '����:2015-01-30 17:38:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    If mbytInState = 1 Then Exit Sub
    
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then Exit Sub
    
    If objCard.�ӿ���� > 0 Then
         If Not objCard.���ѿ� Then
             '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
             txtBalance(Idx_�ɿ�).Text = lblʣ���Ը�.Caption
         Else
             txtBalance(Idx_�ɿ�).Text = lblʣ���Ը�.Caption
         End If
    ElseIf objCard.�������� <> 1 Then
         txtBalance(Idx_�ɿ�).Text = ""
    Else
         txtBalance(Idx_�ɿ�).Text = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
