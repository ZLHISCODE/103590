VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CO8DDC~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmAdviceReprotBrowse 
   Caption         =   "���������"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14805
   Icon            =   "frmAdviceReprotBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   14805
   StartUpPosition =   1  '����������
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
            Caption         =   "���:"
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
            Index           =   1
            Left            =   60
            TabIndex        =   13
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
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
               Caption         =   "���ѡ��"
               BeginProperty Font 
                  Name            =   "����"
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
                  Caption         =   "������"
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
                  Caption         =   "��ϸ������"
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
                  Caption         =   "���²�������"
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
                  Caption         =   "����"
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
                  Caption         =   "����"
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
                     Name            =   "����"
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
                     Name            =   "����"
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
                     Name            =   "����"
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
                  Caption         =   "��������"
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
                  Left            =   60
                  TabIndex        =   49
                  Top             =   1800
                  Width           =   960
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "δ �� ��"
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
                  Left            =   60
                  TabIndex        =   48
                  Top             =   930
                  Width           =   960
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������"
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
                  Left            =   60
                  TabIndex        =   47
                  Top             =   210
                  Width           =   960
               End
            End
            Begin VB.Frame fraOne 
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "����"
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
                     Name            =   "����"
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
                     Name            =   "����"
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
                     Name            =   "����"
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
                  Text            =   "��΢�����"
                  Top             =   270
                  Width           =   3915
               End
               Begin VB.Label lblMicroscopeFinded 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������"
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
                  Left            =   90
                  TabIndex        =   42
                  Top             =   660
                  Width           =   960
               End
               Begin VB.Label lblMicroscopeNOFind 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "δ �� ��"
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
                  Left            =   60
                  TabIndex        =   41
                  Top             =   1290
                  Width           =   960
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ͨ���豸"
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
            Caption         =   "����:"
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
            Index           =   3
            Left            =   6870
            TabIndex        =   7
            Top             =   60
            Width           =   600
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
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
            Index           =   2
            Left            =   4830
            TabIndex        =   6
            Top             =   60
            Width           =   600
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�:"
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
            Index           =   1
            Left            =   2640
            TabIndex        =   5
            Top             =   60
            Width           =   600
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
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
            ToolTipText     =   "˫���鿴��ͼ"
            Top             =   3600
            Width           =   2145
            Begin VB.Label lblLoading 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ToolTipText     =   "����鿴��ͼ"
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
            ToolTipText     =   "˫���鿴��ͼ"
            Top             =   3600
            Width           =   2145
            Begin VB.Label lblLoading 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ToolTipText     =   "����鿴��ͼ"
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
            ToolTipText     =   "˫���鿴��ͼ"
            Top             =   3600
            Width           =   2145
            Begin VB.Label lblLoading 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ToolTipText     =   "����鿴��ͼ"
               Top             =   60
               Width           =   1965
            End
         End
         Begin VB.Label lblAuditingTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4650
            TabIndex        =   32
            Top             =   5730
            Width           =   1080
         End
         Begin VB.Label lblAuditingMan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1230
            TabIndex        =   31
            Top             =   5730
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
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
            Index           =   8
            Left            =   3540
            TabIndex        =   30
            Top             =   5730
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "�� ϸ ��:"
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
            Left            =   3540
            TabIndex        =   26
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��״��Ƥϸ��:"
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
            Index           =   4
            Left            =   300
            TabIndex        =   25
            Top             =   630
            Width           =   1560
         End
         Begin VB.Label lblXJ 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "ϸ��:"
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
            Left            =   300
            TabIndex        =   21
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������̬:"
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
            Index           =   3
            Left            =   3540
            TabIndex        =   20
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��        ״:"
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
'��    ��:������
'����ʱ��:2017/4/25
'ģ�鹦��:һ�����Ʋ����ļ��鱨��鿴ģ��
'---------------------------------------------------------------------------------------

Option Explicit

'��̬�����Ƿ���ʾ����߿�
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const const_PicRectBackColour As Long = &HE0E0E0

Private mobjFrm As Object

Private mlngAdvicID As String
Private mlngSampleID As Long        '�걾��
Private mlngPaintID As Long         '����ID
Private mintVersion As Integer      '�汾�� 10=�ϰ�,25=�°�
Private mstrPrivs As String         'ģ��Ȩ��
Private mblnHaveBoder As Boolean    '�Ƿ���ʾ����߿�Ͱ�ť
Private mblnDoctorShow As Boolean   '�Ƿ���ҽ��վ���
Private mstrSupplementID As String  '���䱨��ָ��ID

Private mlngVsfHeight As Long       'VSF�ĸ߶�
Private mlngElseCrlHeight As Long   '�����ؼ��ĸ߶�

Private mobjFTP As New clsFtp               'FTP����
Private mblnFtp As Boolean                  'FTP�Ƿ����
Private mstrFtpIp As String                 'FTP���ӵ�ַ
Private mstrFtpUser As String               'FTP�û�
Private mstrFtpPwd As String                'FTP����
Private mstrFtpFolder As String             'FTPĿ¼

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/25
'��    ��:�򿪴���
'           strAdvices      ҽ��ID�����á�,���ָ�
'           intType         �Ƿ�ֱ��Ԥ������
'��    ��:
'��    ��:
'��    ��:
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
'��    ��:������
'����ʱ��:2019-07-27
'��    ��:  ͨ���걾ID��ʾ���棨�����ѯģ����ã�
'��    ��:
'           objFrm          ���ô���
'           mblnDoctorShow  �Ƿ���ҽ��վ����
'           lngSampleID     �걾ID
'           lngPaintID      ����ID
'           intVersion      ����汾��25=�°�LIS��10=�ϰ�LIS
'           intSampleType   �Ƿ���΢���ﱨ�棬0=��ͨ���棬1=΢���ﱨ��
'           intPositive     �������ͣ�1=ҩ�����棬3=PDF���棬����=���Ա���
'           strDiagnosis    ���
'           strResult       ��ע
'           intCount        �ϰ�LIS�������
'           strSupplementID ���䱨��ָ��ID
'           strPrivs        ��ԱȨ��
'��    ��:
'           strThirdReport  ��������
'��    ��:
'����Ӱ��:
'����ע��:
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

          '��ѯ�����Ŀ
13        If intVersion = 25 Then
14            strSQL = "Select Distinct a.������Ŀ || '(' || to_char(a.����ʱ��, 'yyyy/mm/dd hh24:mi:Ss') || '��' || a.�걾���� || ')' �������,a.����ʱ��,a.�Ƿ�Ⱦ��,a.������" & vbCrLf & _
                     "   From ���鱨���¼ A" & vbCrLf & _
                     "   Where a.id = [1]"
15            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����Ŀ", lngSampleID)
16        ElseIf intVersion = 10 Then
17            strSQL = "Select f_List2str(Cast(Collect(b.ҽ������) As t_Strlist)) || '(' || to_char(a.����ʱ��, 'yyyy/mm/dd hh24:mi:Ss') || '��' ||" & vbCrLf & _
                     "           a.�걾���� || ')' �������,a.����ʱ��,0 �Ƿ�Ⱦ��,'' ������" & vbCrLf & _
                     "   From ����걾��¼ A, ����ҽ����¼ B" & vbCrLf & _
                     "   Where a.ҽ��id = b.Id(+) and a.id=[1]" & vbCrLf & _
                     "   Group By a.id, a.����ʱ��, a.�걾����,a.����ʱ��"
18            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�����Ŀ", lngSampleID)
19        End If
20        If Not rsTmp.EOF Then
21            lblPatient(0).Caption = rsTmp("�������") & IIf(intVersion = 25, "(�°�)", "(�ϰ�)")
22            If Val(rsTmp("�Ƿ�Ⱦ��") & "") = 1 Then
23                lblPatient(0).Caption = lblPatient(0).Caption & "(���ƴ�Ⱦ��)"
24                lblPatient(0).ForeColor = vbRed
25                If rsTmp("������") & "" = "" Then
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
38            lblPatient(0).Caption = "������Ŀ:"
39        End If
40        lblPatient(0).FontBold = True

41        If intVersion = 10 Then    '�ϰ�LIS
42            Call GetSampleFromOldLis(lngSampleID, intSampleType, intCount)
43        ElseIf intVersion = 25 Then    '�°�LIS
44            Call GetSampleFromNewLis(lngSampleID, intSampleType, intPositive, strDiagnosis, strResult, strThirdReport)
45        End If

46        ShowReportByID = GetFrmHeight(intSampleType)


47        Exit Function
ShowReportByID_Error:
48        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(ShowReportByID)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
49        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-27
'��    ��:  ��ȡ����ĸ߶�
'��    ��:
'           intSampleType       �������ͣ�1=΢���ﱨ��
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
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
'��    ��:������
'����ʱ��:2018/5/25
'��    ��:����API��̬���ô����border
'��    ��:
'           new_Hwnd    ����ľ��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Exit        '�˳�
            Unload Me
        Case ConMenu_Browse_PrintView   'Ԥ��
            Call PrintReport(Me, 1)
        Case ConMenu_Browse_PrintSet    '��ӡ����
            Call PrintReport(Me, 3)
        Case ConMenu_Appfor_ClincHelp   '���Ʋο�
            Call funShowClincHelp(Me, mlngSampleID, mintVersion)
        Case ConMenu_Browse_Print       '��ӡ
            Call PrintReport(Me, 2)
        Case conFun_Sample_Auditing     '����
            Call AuditingSample(1)
        Case conFun_Sample_unAuditing     'ȡ������
            Call AuditingSample(2)
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
            Call ExePlugIn(Control.Parameter, mlngSampleID)
    End Select
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/4
'��    ��:Ԥ������ӡ���ã���ӡ
'��    ��:
'           objfrm          �������
'           byRunMode       1=Ԥ��,2=��ӡ��3=��ӡ����
'��    ��:
'��    ��:
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
        Case ConMenu_Appfor_ClincHelp       '���Ʋο�
            Control.Visible = VerCompare(gSysInfo.VersionHIS, "10.35.120") <> -1
    End Select
End Sub

Private Sub chkMicroscope_Click()
    PicNegative_Resize
End Sub

Private Sub Form_Load()
'���ܴ���������
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
        .IconsWithShadow = True    '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "��ӡ")
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "��ӡ����  ")
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_unPrint, "���ô�ӡ  ")
            cbrControl.Visible = InStr(mstrPrivs, "���������������ӡ����") > 0
        End With
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conFun_Sample_Auditing, "����"): cbrControl.BeginGroup = True
        cbrControl.Visible = Not mblnDoctorShow
        cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conFun_Sample_unAuditing, "ȡ������")
        cbrControl.Visible = Not mblnDoctorShow
        cbrControl.Enabled = False

        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "���Ʋο�")
        cbrControl.BeginGroup = True
        If mblnHaveBoder Then
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "�˳�")
            cbrControl.BeginGroup = True
        End If
    End With

    '���������ť
    Call CreatePlugInButton(cbrToolBar)

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next


    Call intData
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/25
'��    ��:��ʼ������
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub intData()
    Call setVSF
    Call setTabType '���÷�ҳ
    
'    Call GetSampleInformation        '��ȡҽ��ID��Ӧ�ı걾��Ϣ
End Sub

Private Sub setTabType()
    With Me.tabThis
        .PaintManager.Appearance = xtpTabAppearanceStateButtons
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionLeft

        
        .InsertItem 0, "ͿƬ����", MicroorganismSmear.hWnd, 1
        .PaintManager.Layout = xtpTabLayoutAutoSize
        
        
        .InsertItem 1, "����ҩ��", picCJYM.hWnd, 2
        .PaintManager.Layout = xtpTabLayoutAutoSize
        
        .InsertItem 2, "΢���ﱨ��", picCenter.hWnd, 3
        .PaintManager.Layout = xtpTabLayoutAutoSize
        
        .Item(2).Selected = True
    End With
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/25
'��    ��:��ʼ��VSF�б�
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub setVSF()
      'չʾ��ͨ�����VSF

1         On Error GoTo setVSF_Error

2         With Me.vsfGeneral
3             .FixedRows = 0
4             .FixedCols = 0
5             .Rows = 1
6             .Cols = 8
7             .SelectionMode = flexSelectionFree  '����ѡ��
8             .AllowSelection = False    '������ѡ��
9             .BorderStyle = flexBorderNone    '�ޱ߿�
10            .GridLines = flexGridNone    '��������
11            .FontSize = 12  'С��
12            .BackColorBkg = vbWhite    '��ɫ����
13            .SheetBorder = vbWhite  '��ɫ����
14            .BorderStyle = flexBorderNone
15            .ExplorerBar = flexExSortShowAndMove    '����������򣬲���ʾ����ͼ��
16            .AllowUserResizing = flexResizeColumns  '�ɵ����п�
17            .Editable = flexEDNone                  'ֻ��

18            .ColKey(0) = "�걾ID": .TextMatrix(0, .ColIndex("�걾ID")) = "�걾ID": .ColWidth(.ColIndex("�걾ID")) = 0: .ColHidden(.ColIndex("�걾ID")) = True
19            .ColKey(1) = "������Ŀ": .TextMatrix(0, .ColIndex("������Ŀ")) = "������Ŀ": .ColWidth(.ColIndex("������Ŀ")) = 4000: .ColHidden(.ColIndex("������Ŀ")) = False
20            .ColKey(2) = "������": .TextMatrix(0, .ColIndex("������")) = "������": .ColWidth(.ColIndex("������")) = 1100: .ColHidden(.ColIndex("������")) = False
21            .ColKey(3) = "�����λ": .TextMatrix(0, .ColIndex("�����λ")) = "�����λ": .ColWidth(.ColIndex("�����λ")) = 1100: .ColHidden(.ColIndex("�����λ")) = False
22            .ColKey(4) = "��־": .TextMatrix(0, .ColIndex("��־")) = "��־": .ColWidth(.ColIndex("��־")) = 800: .ColHidden(.ColIndex("��־")) = False
23            .ColKey(5) = "����ο�": .TextMatrix(0, .ColIndex("����ο�")) = "����ο�": .ColWidth(.ColIndex("����ο�")) = 2000: .ColHidden(.ColIndex("����ο�")) = False
24            .ColKey(6) = "�ٴ�����": .TextMatrix(0, .ColIndex("�ٴ�����")) = "�ٴ�����": .ColWidth(.ColIndex("�ٴ�����")) = 2000: .ColHidden(.ColIndex("�ٴ�����")) = True
25            .ColKey(7) = "ID": .TextMatrix(0, .ColIndex("ID")) = "ID": .ColWidth(.ColIndex("ID")) = 2000: .ColHidden(.ColIndex("ID")) = True
26            .Cell(flexcpAlignment, 0, .ColIndex("�걾ID"), 0, .ColIndex("����ο�")) = flexAlignLeftCenter  '���⿿�����
27        End With

          '����ҩ������
28        With Me.VSFCJYM
29            .FixedRows = 0
30            .FixedCols = 0
31            .Rows = 1
32            .Cols = 6
33            .SelectionMode = flexSelectionFree  '����ѡ��
34            .AllowSelection = False    '������ѡ��
35            .BorderStyle = flexBorderNone    '�ޱ߿�
36            .GridLines = flexGridNone    '��������
37            .FontSize = 12  'С��
38            .BackColorBkg = vbWhite    '��ɫ����
39            .SheetBorder = vbWhite  '��ɫ����
40            .BorderStyle = flexBorderNone
41            .AllowUserResizing = flexResizeColumns  '�ɵ����п�
42            .Editable = flexEDNone                  'ֻ��
43            .MergeCells = flexMergeRestrictRows     '�������ϲ�
44            .OutlineBar = flexOutlineBarComplete    '���νṹ
45            .OutlineCol = 0    '���νڵ���
46            .SubtotalPosition = flexSTAbove    '���νṹ��ʽ

47            .ColKey(0) = "ϸ����": .ColWidth(.ColIndex("ϸ����")) = 3000: .ColAlignment(.ColIndex("ϸ����")) = flexAlignLeftCenter
48            .ColKey(1) = "������": .ColWidth(.ColIndex("������")) = 1500: .ColAlignment(.ColIndex("������")) = flexAlignLeftCenter
49            .ColKey(2) = "����": .ColWidth(.ColIndex("����")) = 1500: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter
50            .ColKey(3) = "��ҩ����": .ColWidth(.ColIndex("��ҩ����")) = 1500: .ColAlignment(.ColIndex("��ҩ����")) = flexAlignLeftCenter
51            .ColKey(4) = "�ο�����": .ColWidth(.ColIndex("�ο�����")) = 1500: .ColAlignment(.ColIndex("�ο�����")) = flexAlignLeftCenter
52            .ColKey(5) = "Level": .ColWidth(.ColIndex("Level")) = 1500: .ColAlignment(.ColIndex("Level")) = flexAlignLeftCenter: .ColHidden(.ColIndex("Level")) = True

              '�̶���
53            .TextMatrix(0, .ColIndex("ϸ����")) = "ϸ����"
54            .TextMatrix(0, .ColIndex("������")) = "������"
55            .TextMatrix(0, .ColIndex("����")) = "����"
56            .TextMatrix(0, .ColIndex("��ҩ����")) = "��ҩ����"
57            .TextMatrix(0, .ColIndex("�ο�����")) = "�ο�����"
58        End With

          'չʾ΢��������vsf
59        With Me.vsfMicrobePositive
60            .FixedRows = 0
61            .FixedCols = 0
62            .Rows = 1
63            .Cols = 6
64            .SelectionMode = flexSelectionFree  '����ѡ��
65            .AllowSelection = False    '������ѡ��
66            .BorderStyle = flexBorderNone    '�ޱ߿�
67            .GridLines = flexGridNone    '��������
68            .FontSize = 12  'С��
69            .BackColorBkg = vbWhite    '��ɫ����
70            .SheetBorder = vbWhite  '��ɫ����
71            .BorderStyle = flexBorderNone
72            .AllowUserResizing = flexResizeColumns  '�ɵ����п�
73            .Editable = flexEDNone                  'ֻ��
74            .MergeCells = flexMergeRestrictRows     '�������ϲ�
75            .OutlineBar = flexOutlineBarComplete    '���νṹ
76            .OutlineCol = 2    '���νڵ���
77            .SubtotalPosition = flexSTAbove    '���νṹ��ʽ

78            .ColKey(0) = "KEY": .TextMatrix(0, .ColIndex("KEY")) = "KEY": .ColWidth(.ColIndex("KEY")) = 0: .ColHidden(.ColIndex("KEY")) = True
79            .ColKey(1) = "������ID": .TextMatrix(0, .ColIndex("������ID")) = "������ID": .ColWidth(.ColIndex("������ID")) = 0: .ColHidden(.ColIndex("������ID")) = True
80            .ColKey(2) = "����������": .TextMatrix(0, .ColIndex("����������")) = "ϸ����": .ColWidth(.ColIndex("����������")) = 4000: .ColHidden(.ColIndex("����������")) = False
81            .ColKey(3) = "������": .TextMatrix(0, .ColIndex("������")) = "������": .ColWidth(.ColIndex("������")) = 2000: .ColHidden(.ColIndex("������")) = False
82            .ColKey(4) = "�������": .TextMatrix(0, .ColIndex("�������")) = "����": .ColWidth(.ColIndex("�������")) = 1300: .ColHidden(.ColIndex("�������")) = False
83            .ColKey(5) = "ҩ������": .TextMatrix(0, .ColIndex("ҩ������")) = "��ҩ����": .ColWidth(.ColIndex("ҩ������")) = 1300: .ColHidden(.ColIndex("ҩ������")) = False
84            .Cell(flexcpAlignment, 0, .ColIndex("KEY"), 0, .ColIndex("ҩ������")) = flexAlignLeftCenter  '���⿿�����
85        End With

86        Exit Sub
setVSF_Error:
87        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(setVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
88        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/2
'��    ��:��ȡ�걾��Ϣ�������걾�����°��л����ϰ��У��걾�ţ��Ƿ���΢����걾��
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Function GetSampleInformation() As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim lngSampleID As Long           '�걾ID
          Dim intSampleType As Integer      '�걾���ͣ�0=��ͨ�걾,1=΢����걾
          Dim intVersion As Integer         '�汾��25=�°棬10=�ϰ�
          Dim intCount As Integer           '�������
          Dim intPositive As Integer        '�������ͣ�1=ҩ������,3=PDF����
          Dim strDiagnosis As String        '���
          Dim strResult As String           '��ע
          Dim strSQR As String              '������
          Dim intIsDis As Integer           '�Ƿ��Ǵ�Ⱦ��
          
          '�ж϶�Ӧҽ��ID���ϰ����Ƿ���ڱ걾,���걾�Ƿ�Ϊ΢����걾
1         On Error GoTo GetSampleInformation_Error

2         strSQL = "select a.id �걾ID,a.΢����걾,a.������,a.����ID from ����걾��¼ A where a.ҽ��id=[1] and ����� is not null"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�걾��Դ", mlngAdvicID)
4         If rsTmp.RecordCount > 0 Then
5             intVersion = 10 '�걾���ϰ�LIS��
6             lngSampleID = Val(rsTmp("�걾ID") & "") '��ȡ�걾ID
7             mlngPaintID = Val(rsTmp("����ID") & "")
8             intSampleType = Val(rsTmp("΢����걾") & "")  '��ȡ�걾����
9             intCount = Val(rsTmp("������") & "")  '�걾�������
10        Else
              '�ϰ���û�в�ѯ��ҽ����صľ;͵��°�LIS��ȥ����
11            strSQL = "select b.id �걾ID,b.΢����,b.���Ա���,b.���,b.��ע,a.������,b.�Ƿ�Ⱦ�� from ����������� A,���鱨���¼ B" & _
                      " where a.�걾id=b.id and a.����id=[1]  and b.����� is not null"
12            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�걾��Դ", mlngAdvicID)
13            If rsTmp.RecordCount > 0 Then
14                intVersion = 25 '�걾���°�LIS��
15                lngSampleID = Val(rsTmp("�걾ID") & "") '��ȡ�걾ID
16                intSampleType = Val(rsTmp("΢����") & "") '��ȡ�걾����
17                intPositive = Val(rsTmp("���Ա���") & "") '��������
18                strDiagnosis = rsTmp("���") & ""   '���
19                strResult = rsTmp("��ע") & ""   '��ע
20                strSQR = rsTmp("������") & ""
21                intIsDis = Val(rsTmp("�Ƿ�Ⱦ��") & "")
22            Else
23                GetSampleInformation = False
24                Exit Function
25            End If
26        End If
27        mintVersion = intVersion
28        mlngSampleID = lngSampleID
          
          '��鵱ǰ�û��Ƿ��ܹ��鿴��Ⱦ������
29        If strSQR <> gUserInfo.Name And strSQR <> "" And InStr(";" & mstrPrivs & ";", ";�鿴��Ⱦ������;") <= 0 And intIsDis = 1 Then
30            If Me.Tag = "" Then
31                MsgBox "Ȩ�޲��㣬�޷��鿴�˱���", vbInformation, Me.Caption
32                Me.Tag = "True"
33            End If
34            Exit Function
35        End If
              
          '��ѯ������Ϣ
36        If mintVersion <= 0 Or lngSampleID <= 0 Then Exit Function
37        If GetPatient(intVersion, lngSampleID) = False Then Exit Function

          '��ѯ�걾��¼
38        If intVersion = 10 Then '�ϰ�LIS
39            Call GetSampleFromOldLis(lngSampleID, intSampleType, intCount)
40        ElseIf intVersion = 25 Then '�°�LIS
41            Call GetSampleFromNewLis(lngSampleID, intSampleType, intPositive, strDiagnosis, strResult)
42        End If

43        GetSampleInformation = True
          
44        Exit Function
GetSampleInformation_Error:
45        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(GetSampleInformation)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
46        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/2
'��    ��:��ȡ��Ա������Ϣ
'��    ��:
'           intVersion          �汾��
'           lngSampleType       �걾��
'��    ��:
'��    ��:  True=��ѯ��������Ϣ,False=û�в�ѯ��������Ϣ
'---------------------------------------------------------------------------------------
Private Function GetPatient(ByVal intVersion As Integer, ByVal lngSampleID As Long) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo GetPatient_Error

2         If intVersion = 10 Then '�ϰ�
3             strSQL = "select ����,�Ա�,����,���� from ����걾��¼ where ID=[1]"
4             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��Ա��Ϣ", lngSampleID)
5         ElseIf intVersion = 25 Then '�°�
6             strSQL = "select ����,decode(�Ա�,1,'��',2,'Ů','δ��֪') �Ա�,����,���� from ���鱨���¼ where ID=[1]"
7             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��Ա��Ϣ", lngSampleID)
8         End If

          '�����ѯ������������ݣ�û�������򷵻�False���˳�
9         If rsTmp.RecordCount > 0 Then
10            Me.lblPatient(0).Caption = "����:" & rsTmp("����")
11            Me.lblPatient(1).Caption = "�Ա�:" & rsTmp("�Ա�")
12            Me.lblPatient(2).Caption = "����:" & rsTmp("����")
13            Me.lblPatient(3).Caption = "����:" & rsTmp("����")
14        Else
15            GetPatient = False
16            Exit Function
17        End If
          
18        GetPatient = True


19        Exit Function
GetPatient_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(GetPatient)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/2
'��    ��:���ϰ�LIS�в�ѯ�걾��¼
'��    ��:
'           lngSampleID             �걾ID
'           intSampleType           �걾����
'           intCount                �걾�������
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub GetSampleFromOldLis(ByVal lngSampleID As Long, ByVal intSampleType As Integer, ByVal intCount As Integer)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsAntibiotic As ADODB.Recordset
          
          '���鱸ע
1         On Error GoTo GetSampleFromOldLis_Error

2         strSQL = "SELECT A.��ע FROM ����걾��¼ A WHERE A.ID= [1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��������", lngSampleID)
4         If rsTmp.RecordCount > 0 Then
5             Me.txtResult.Text = rsTmp("��ע") & ""
6             rsTmp.MoveNext
7         End If
          
          '���
8         strSQL = "Select b.ҽ��id, b.��Ŀ, b.����, b.���� From ����걾��¼ a, ����ҽ������ b Where a.ҽ��id = b.ҽ��id and a.ID =[1] Order By ҽ��id, ����"
9         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���", lngSampleID)
10        Me.txtDiagnosis.Text = ""
11        If rsTmp.RecordCount > 0 Then
12            Do While Not rsTmp.EOF
13                Me.txtDiagnosis.Text = Me.txtDiagnosis.Text & NVL(rsTmp("��Ŀ")) & ":" & Replace(NVL(rsTmp("����")), vbCrLf, vbCrLf & "    ") & vbCrLf
14                rsTmp.MoveNext
15            Loop
16        End If
          
         
17        If intSampleType = 0 Then
              '��ͨ�걾
18            strSQL = "Select b.������ĿID ID, a.Id As �걾id, c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ, b.������,d.��λ as �����λ," & _
                       "      Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־,b.����ο�,D.�ٴ�����" & _
                      " From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������ĿĿ¼ H, ������ˮ��ָ�� E" & _
                      " Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And b.������Ŀid = h.Id(+) And b.����걾id = e.�걾id(+) And" & _
                      "      b.������Ŀid = e.��Ŀid(+) And b.��¼���� = [1] And a.Id = [2]" & _
                      " Union All" & _
                      " Select b.������ĿID ID, a.Id As �걾id, c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ, b.������,d.��λ as �����λ," & _
                      "       Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־,b.����ο�,D.�ٴ�����" & _
                      " From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������ĿĿ¼ H, ������ˮ��ָ�� E" & _
                      " Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And b.������Ŀid = h.Id(+) And b.��¼���� = [1] And" & _
                      "      b.����걾id = e.�걾id(+) And b.������Ŀid = e.��Ŀid(+) And a.�ϲ�id = [2]"
19            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ͨ���", intCount, lngSampleID)
20            Call SetGeneralData(rsTmp)
21        ElseIf intSampleType = 1 Then
              '΢����걾
              
              '��ѯϸ��
22            strSQL = "Select b.Id, b.������ As ϸ����, a.������ As ������, a.�������� As ����, a.��ҩ����,'' ��������" & _
                       " From ������ͨ��� A, ����ϸ�� B, ����걾��¼ D" & _
                       " Where a.ϸ��id = b.Id And a.��¼���� = [1] And d.Id = a.����걾id And d.Id = [2]" & _
                       " Order By b.����"
23            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ϸ��", intCount, lngSampleID)

              '��ѯҩ��
24            strSQL = "Select c.ϸ��id As Key, b.Id ������ID, b.������ As ����������, a.��� As ������," & _
                       "      Decode(a.�������, 'R', 'R-��ҩ', 'I', 'I-�н�', 'S', 'S-����', a.�������) As �������," & _
                       "      Decode(a.ҩ������, 1, '1-MIC', 2, '2-DISK', 3, '3-K-B', '') As ҩ������" & _
                       " From ����ҩ����� A, �����ÿ����� B, ������ͨ��� C" & _
                       " Where a.������id = b.Id And c.Id = a.ϸ�����id And c.��¼���� = a.��¼���� And c.����걾id = [1] And c.��¼���� = [2]" & _
                       " Order By c.ϸ��id, b.����"
25            Set rsAntibiotic = ComOpenSQL(Sel_His_DB, strSQL, "����ҩ��", lngSampleID, intCount)
26            Call SetMicroorganismData(rsTmp, rsAntibiotic)    '������
27        End If
          
28        Call SetCrlTyep(intSampleType, 1) '���ô�����Ҫ��ʾ��Щ�ؼ�



29        Exit Sub
GetSampleFromOldLis_Error:
30        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(GetSampleFromOldLis)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
31        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/2
'��    ��:���°�LIS�в�ѯ�걾��¼
'��    ��:
'           lngSampleID             �걾ID
'           intSampleType           �걾����
'           intPositive             �������ͣ�1=ҩ�����棬3=PDF����
'           strDiagnosis            �ٴ����
'           strThirdReport          ����PDF����
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Function GetSampleFromNewLis(ByVal lngSampleID As Long, ByVal intSampleType As Integer, _
                                     ByVal intPositive As Integer, ByVal strDiagnosis As String, _
                                     ByVal strResult As String, Optional ByRef strThirdReport As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsAntibiotic As ADODB.Recordset
          Dim lngRow As Long


1         On Error GoTo GetSampleFromNewLis_Error


          '��ע
2         If strResult <> "" Then
3             Me.txtResult.Text = strResult
4         End If

          '���
5         If strDiagnosis <> "" Then
6             Me.txtDiagnosis.Text = strDiagnosis
7         End If

8         picMain.Visible = True
9         picPDF.Visible = False
10        If intSampleType = 0 Then
              '��ͨ�걾
11            If IsTre(lngSampleID) Then
12                strSQL = "select * from (Select Distinct c.id, a.Id �걾id,b.id ������ϸID, c.������ || '(' || c.Ӣ���� || ')' || decode(h.����ʱ��,null,'', '(' || h.����ʱ�� || ')')  ������Ŀ, b.������, c.��λ �����λ," & _
                         "               Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') ��־, b.����ο�,c.�ٴ�����" & _
                         " From ���鱨���¼ A, ���鱨����ϸ B, ����ָ�� C, ���������Ŀ D, ����������� E, ��ˮ�߼���ָ�� F,��������걾 G,��������ʱ�䷽�� H" & _
                         " Where a.Id = b.�걾id And b.��Ŀid = c.Id And b.���id = d.Id(+) And b.�걾id = f.�걾id(+) And b.��Ŀid = f.��Ŀid(+) And" & _
                         "      b.�걾id = e.�걾id And d.Id = e.���id and b.ID=g.������ϸid(+) and g.���ܷ���id=H.id(+) And b.���id Is Not Null And e.���id Is Not Null And a.Id = [1]" & _
                         " Union All" & _
                         " Select Distinct c.id, a.Id �걾id,b.id ������ϸID,  c.������ || '(' || c.Ӣ���� || ')' || decode(h.����ʱ��,null,'', '(' || h.����ʱ�� || ')') ������Ŀ, b.������, c.��λ �����λ," & _
                         "                Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') ��־, b.����ο�,c.�ٴ�����" & _
                         " From ���鱨���¼ A, ���鱨����ϸ B, ����ָ�� C, ���������Ŀ D, ����������� E, ��ˮ�߼���ָ�� F,��������걾 G,��������ʱ�䷽�� H" & _
                         " Where a.Id = b.�걾id And b.��Ŀid = c.Id And b.���id = d.Id(+) And b.�걾id = f.�걾id(+) And b.��Ŀid = f.��Ŀid(+) And" & _
                         "      b.�걾id = e.�걾id and b.ID=g.������ϸid(+) and g.���ܷ���id=H.id(+) And e.���id Is Null And b.���id Is Null And a.Id =[1] ) order by ������ϸID desc"
13            Else
14                strSQL = "select * from (Select Distinct c.id, a.Id �걾id,b.id ������ϸID, c.������ || '(' || c.Ӣ���� || ')'   ������Ŀ, b.������, c.��λ �����λ," & _
                         "               Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') ��־, b.����ο�,c.�ٴ�����" & _
                         " From ���鱨���¼ A, ���鱨����ϸ B, ����ָ�� C, ���������Ŀ D, ����������� E, ��ˮ�߼���ָ�� F" & _
                         " Where a.Id = b.�걾id And b.��Ŀid = c.Id And b.���id = d.Id(+) And b.�걾id = f.�걾id(+) And b.��Ŀid = f.��Ŀid(+) And" & _
                         "      b.�걾id = e.�걾id And d.Id = e.���id and b.���id Is Not Null And e.���id Is Not Null And a.Id = [1]" & _
                         " Union All" & _
                         " Select Distinct c.id, a.Id �걾id,b.id ������ϸID,  c.������ || '(' || c.Ӣ���� || ')'  ������Ŀ, b.������, c.��λ �����λ," & _
                         "                Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') ��־, b.����ο�,c.�ٴ�����" & _
                         " From ���鱨���¼ A, ���鱨����ϸ B, ����ָ�� C, ���������Ŀ D, ����������� E, ��ˮ�߼���ָ�� F" & _
                         " Where a.Id = b.�걾id And b.��Ŀid = c.Id And b.���id = d.Id(+) And b.�걾id = f.�걾id(+) And b.��Ŀid = f.��Ŀid(+) And" & _
                         "      b.�걾id = e.�걾id and e.���id Is Null And b.���id Is Null And a.Id =[1] ) order by ������ϸID desc"
15            End If
16            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "������ϸ", lngSampleID)
17            Call SetGeneralData(rsTmp)

              '��չʾ�걾�����VSF�����ɾ����
18            With vsfGeneral
19                For lngRow = 1 To .Rows - 1
20                    If InStr("," & mstrSupplementID & ",", "," & .TextMatrix(lngRow, .ColIndex("ID")) & ",") > 0 Then
21                        vsfGeneral.Cell(flexcpFontStrikethru, lngRow, 0, lngRow, vsfGeneral.Cols - 1) = True
22                    End If
23                Next
24            End With

25        ElseIf intSampleType = 1 Then
26            If intPositive = 1 Then
                  '΢�������Ա���
27                strSQL = "Select b.Id, b.������ || '(' || b.Ӣ���� || ')' ϸ����, a.������, a.�������� ����, a.��ҩ����,a.��������" & _
                         " From ���鱨��ϸ�� A, ����ϸ����¼ B" & _
                         " Where a.ϸ��id = b.Id(+) And a.�걾id = [1]"
28                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ϸ��", lngSampleID)

                  '��ѯҩ��
29                strSQL = "Select a.ϸ��id Key, c.Id ������id, c.������ || '(' || c.Ӣ���� || ')' ����������, b.��� ������, b.�������, b.ҩ������" & _
                         " From ���鱨��ϸ�� A, ���鱨��ҩ�� B, ����ҩ�� C, ����ҩ������ҩ D" & _
                         " Where a.Id = b.���id And b.ҩ��id = c.Id And b.ҩ��id = d.ҩ��id(+) And b.ҩ����id = d.ҩ����id(+) And a.�걾id = [1]" & _
                         " Order By d.ҩ����id, d.�������"
30                Set rsAntibiotic = ComOpenSQL(Sel_Lis_DB, strSQL, "����ҩ��", lngSampleID)
31                Call SetMicroorganismData(rsTmp, rsAntibiotic)  '������
32            ElseIf intPositive = 3 Then
                  'PDF����
33                picMain.Visible = False
34                picPDF.Visible = True
35                strThirdReport = findThirdReport(lngSampleID, webSub)
36            Else
                  '΢�������Ա���
37                strSQL = "Select  a.������, a.δ���, a.��������, a.���²���, a.��ϸ��," & _
                           "A.�����豸 , A.������, A.����δ���, A.��������,a.�Ƿ񾵼���,a.�������" & _
                         " From ���鱨��ϸ�� A, ����ϸ����¼ B" & _
                         " Where a.ϸ��id = b.Id(+) And a.�걾id = [1]"
38                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���Ա���", lngSampleID)
39                Call SetPositiveData(rsTmp)
40            End If
41            Call GetfrmMicroorganismSmear(lngSampleID)  '��ѯͿƬ����
42            Call GetMicroorganisCJYM(lngSampleID)       '����ҩ������
43        End If

44        Call SetCrlTyep(intSampleType, intPositive)    '���ô�����Ҫ��ʾ��Щ�ؼ�

45        Exit Function
GetSampleFromNewLis_Error:
46        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(GetSampleFromNewLis)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
47        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/2
'��    ��:����ͨ���
'��    ��:
'           rsTmp           ���ݼ�¼��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub SetGeneralData(ByVal rsTmp As ADODB.Recordset)

1         On Error GoTo SetGeneralData_Error

2         If Not rsTmp Is Nothing Then
3           If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            
4           With Me.vsfGeneral
                '������
5               .Rows = 1
6               Do While Not rsTmp.EOF
7                   .Rows = .Rows + 1
8                   .TextMatrix(.Rows - 1, .ColIndex("�걾ID")) = rsTmp("�걾ID") & ""
9                   .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsTmp("������Ŀ") & ""
10                  .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
11                  .TextMatrix(.Rows - 1, .ColIndex("�����λ")) = rsTmp("�����λ") & ""
12                  .TextMatrix(.Rows - 1, .ColIndex("��־")) = rsTmp("��־") & ""
13                  .TextMatrix(.Rows - 1, .ColIndex("����ο�")) = rsTmp("����ο�") & ""
14                  .TextMatrix(.Rows - 1, .ColIndex("�ٴ�����")) = rsTmp("�ٴ�����") & ""
15                  .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
16                  .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�걾ID"), .Rows - 1, .ColIndex("�ٴ�����")) = flexAlignLeftCenter  '���ݿ������
17                  rsTmp.MoveNext
18              Loop
          
19              lbl(0).Caption = "��ע��"
                
                '��ȡVSF�߶�
20              If mblnHaveBoder = False Then
21                  If mlngVsfHeight < (.Rows + 7) * .RowHeight(0) Then
22                      mlngVsfHeight = (.Rows + 7) * .RowHeight(0)
23                  End If
24              End If
25          End With
26        End If


27        Exit Sub
SetGeneralData_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(SetGeneralData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
29        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/2
'��    ��:��΢������
'��    ��:
'           rsBacteria          ϸ����¼��
'           rsAntibiotic        ҩ����¼��
'           intVersion          �汾��10=�ϰ棬25=�°�
'��    ��:
'��    ��:
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

                  '��ϸ������
9                 .TextMatrix(.Rows - 2, .ColIndex("KEY")) = rsBacteria("ID") & ""
10                .TextMatrix(.Rows - 2, .ColIndex("����������")) = rsBacteria("ϸ����") & ""
11                .TextMatrix(.Rows - 2, .ColIndex("������")) = rsBacteria("������") & ""
12                .TextMatrix(.Rows - 2, .ColIndex("�������")) = rsBacteria("����") & ""
13                .TextMatrix(.Rows - 2, .ColIndex("ҩ������")) = rsBacteria("��ҩ����") & ""
14                txtResult.Text = txtResult.Text & rsBacteria("��������") & ""
15                .Cell(flexcpAlignment, .Rows - 2, .ColIndex("KEY"), .Rows - 2, .ColIndex("ҩ������")) = flexAlignLeftCenter  '���ݿ������

                  '����
16                .IsSubtotal(.Rows - 2) = True   '����Ϊ���νڵ�
17                .RowOutlineLevel(.Rows - 2) = 3
                  '��ʾ�߿���
18                .CellBorderRange .Rows - 2, .ColIndex("����������"), .Rows - 2, .ColIndex("ҩ������"), vbBlack, 0, 0, 0, 1, 0, 0

                  '���ÿ����ر�����
19                .TextMatrix(.Rows - 1, .ColIndex("����������")) = "����������"
20                .TextMatrix(.Rows - 1, .ColIndex("������")) = "������"
21                .TextMatrix(.Rows - 1, .ColIndex("�������")) = "�������"
22                .TextMatrix(.Rows - 1, .ColIndex("ҩ������")) = "ҩ������"

                  '����ϸ����ҩ��
23                lngKey = Val(rsBacteria("ID") & "")
24                rsAntibiotic.Filter = "KEY=" & lngKey
25                lngRowBegin = 0
26                lngRowCount = 0
27                Do While Not rsAntibiotic.EOF
28                    .Rows = .Rows + 1
29                    lngRowCount = lngRowCount + 1
30                    If lngRowBegin = 0 Then lngRowBegin = .Rows - 2
31                    .TextMatrix(.Rows - 1, .ColIndex("KEY")) = rsAntibiotic("KEY") & ""
32                    .TextMatrix(.Rows - 1, .ColIndex("������ID")) = rsAntibiotic("������ID") & ""
33                    .TextMatrix(.Rows - 1, .ColIndex("����������")) = rsAntibiotic("����������") & ""
34                    .TextMatrix(.Rows - 1, .ColIndex("������")) = rsAntibiotic("������") & ""
35                    .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsAntibiotic("�������") & ""
36                    .TextMatrix(.Rows - 1, .ColIndex("ҩ������")) = rsAntibiotic("ҩ������") & ""
37                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("KEY"), .Rows - 1, .ColIndex("ҩ������")) = flexAlignLeftCenter  '���ݿ������

38                    rsAntibiotic.MoveNext
39                Loop
40                rsBacteria.MoveNext
41            Loop


              '���νṹĬչ��
42            For lngRow = 0 To .Rows - 1
43                If .IsSubtotal(lngRow) = True Then
44                    .IsCollapsed(lngRow) = flexOutlineExpanded    'չ������
45                End If
46            Next

              '��ȡVSF�߶�
47            If mblnHaveBoder = False Then
48                If mlngVsfHeight < (.Rows + 5) * .RowHeight(0) Then
49                    mlngVsfHeight = (.Rows + 5) * .RowHeight(0)
50                End If
51            End If
52        End With

53        Exit Sub
SetMicroorganismData_Error:
54        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(SetMicroorganismData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
55        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/3
'��    ��:��΢�������Ա���
'��    ��:
'           rsTmp           ���Ա����¼��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub SetPositiveData(ByVal rsTmp As ADODB.Recordset)

1         On Error GoTo SetPositiveData_Error

2         If rsTmp.RecordCount > 0 Then
3             rsTmp.MoveFirst
4             txtNormalMicrobe.Text = rsTmp("������") & ""
5             txtNoFindMicrobe.Text = rsTmp("δ���") & ""
6             txtNormalMicrobes.Text = rsTmp("��������") & ""
7             chkPathopoiesiaGerm.value = IIf(Val(rsTmp("���²���") & "") = 1, 1, 0)
8             chkNoGerm.value = IIf(Val(rsTmp("��ϸ��") & "") = 1, 1, 0)
9             txtMicroscope.Text = rsTmp("�����豸") & ""
10            txtMicroscopeNOFind.Text = rsTmp("����δ���") & ""
11            txtMicroscopeFinded.Text = rsTmp("������") & ""
12          If Val(rsTmp("�Ƿ񾵼���") & "") = 0 Then
13              chkMicroscope.value = 0
14          Else
15              chkMicroscope.value = 1
16          End If
17          If Val(rsTmp("�������") & "") = 0 Then
18              optReport(1).value = True
19          Else
20              optReport(0).value = True
21          End If
22          optReportShow

23            txtResult.Text = rsTmp("��������") & ""
24        End If
          
          '���Ա���߶�
25        mlngElseCrlHeight = 6000

26        Exit Sub
SetPositiveData_Error:
27        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(SetPositiveData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
'��    ��:������
'����ʱ��:2017/5/3
'��    ��:���ÿؼ���״̬
'��    ��:
'           intSampleType       �걾״̬��0=��ͨ�걾,1=΢����걾
'           intPositive         ��������,1=ҩ�����棬3=PDF����
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub SetCrlTyep(ByVal intSampleType As Integer, ByVal intPositive As Integer)
    If intSampleType = 0 Then   '��ͨ����
        Me.picTab.Visible = False
        Me.picGeneral.Visible = True
    ElseIf intSampleType = 1 Then   '΢���ﱨ��
        Me.picTab.Visible = True
        Me.picGeneral.Visible = False
        If intPositive <> 1 Then '���Ա���
            picMicrobePositive.Visible = False
            PicNegative.Visible = True
        ElseIf intPositive = 1 Then     '���Ա���
            picMicrobePositive.Visible = True
            PicNegative.Visible = False
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/8/1
'��    ��:��ȡ΢���ﾵ����
'��    ��:
'           lngSmapleID     �걾ID
'��    ��:
'��    ��:
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

2         Call ConnFtp        '���FTP�Ƿ����

3         strSQL = "Select a.��״, a.������̬, a.��״��Ƥϸ��, a.��ϸ��, a.���ϸ��, a.�����, a.���ʱ��" & vbCrLf & _
                 "       From ΢����ͿƬ���� A, ΢����ͿƬϸ�� B, ����ϸ����¼ C" & vbCrLf & _
                 "       Where a.�걾id = b.�걾id(+) And b.ϸ��id = c.Id(+) And a.�걾id = [1]" & IIf(mblnDoctorShow, " And a.����� Is Not Null", "")
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "΢����ͿƬ����", lngSmapleID)
5         If rsTmp.RecordCount <= 0 Then Exit Sub
6         lblXZ.Caption = rsTmp("��״") & ""
7         lblXT.Caption = rsTmp("������̬") & ""
8         lblLZ.Caption = rsTmp("��״��Ƥϸ��") & ""
9         lblWBC.Caption = rsTmp("��ϸ��") & ""
10        lblXJ.Caption = Replace(Replace(rsTmp("���ϸ��") & "", ",", vbCrLf), "()", "")
11        lblAuditingMan.Caption = rsTmp("�����") & ""
12        lblAuditingTime.Caption = rsTmp("���ʱ��") & ""

          '��ѯ��������
13        imgPicture(0).Tag = ""
14        imgPicture(1).Tag = ""
15        imgPicture(2).Tag = ""
16        strFloder = App.Path & "\MicroorganismPicture"
17        strSQL = "select b.id,���, b.ͼ��λ�� from ΢����ͿƬ���� A,΢���ﾵ����� B where a.id=b.����id and a.�걾ID=[1] order by b.���"
18        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������", lngSmapleID)
19        Do While Not rsTmp.EOF
20            intNO = Val(rsTmp("���") & "")
              '��鱾���Ƿ��л����ļ����л��������ȶ�ȡ�����ļ�
21            DoEvents        '����һ����ʾһ��
22            strImgPath = strFloder & "\" & mlngSampleID & "_" & intloop & ".bmp"
23            lblLoading(intloop).Caption = "loading..."
24            If Not objFSO.FileExists(strImgPath) Then
25                If mblnFtp And rsTmp("ͼ��λ��") & "" <> "" Then
                      '��FTP��ȡ
26                    intReturn = mobjFTP.FuncDownloadFile(rsTmp("ͼ��λ��") & "", strImgPath, mlngSampleID & "_" & intloop & ".bmp")
27                    If intReturn = 1 Then
28                        MsgBox "FTP����ʧ��", vbInformation, gSysInfo.AppName
29                        Exit Sub
30                    ElseIf intReturn = 2 Then
31                        MsgBox "ͼ������ʧ��", vbInformation, gSysInfo.AppName
32                        Exit Sub
33                    End If
34                Else
                      '�����ݿ��ȡ
35                    strImgData = gobjHisComLib.ReadLob(2500, 0, Val(rsTmp("id") & ""), strImgPath, 1, 0)
                      '����ͼ��
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

49        Call DeleteImg  'ɾ������ͼƬ
50        mlngVsfHeight = mlngVsfHeight + MicroorganismSmear.Height

51        Exit Sub
GetfrmMicroorganismSmear_Error:
52        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(GetfrmMicroorganismSmear)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
53        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/8/14
'��    ��:��ȡ����ҩ������
'��    ��:
'           lngSmapleID     �걾ID
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub GetMicroorganisCJYM(ByVal lngSmapleID As Long)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
          '��ѯ����
1         On Error GoTo GetMicroorganisCJYM_Error

2         strSQL = " Select ϸ����, ������, ����, ��ҩ����, �ο�����, Level" & vbCrLf & _
                   " From (Select 0 ID, b.Id �ϼ�id, '��������' ϸ����, '������' ������, 'ҩ������' ��ҩ����, '�������' ����, '�ο�����' �ο�����" & vbCrLf & _
                   "        From ΢����ͿƬ���� A, ΢����ͿƬϸ�� B, ����ϸ����¼ C" & vbCrLf & _
                   "        Where A.�걾id = B.�걾id And B.ϸ��id = C.ID And A.�걾id = [1] And A.����ҩ������� Is Not Null" & vbCrLf & _
                   "        Union all" & vbCrLf & _
                   "        Select b.Id, Null �ϼ�id, '���������' || c.������ || '(' || c.Ӣ���� || ')' ϸ����, b.������, b.��ҩ����, b.����, '' �ο�����" & vbCrLf & _
                   "        From ΢����ͿƬ���� A, ΢����ͿƬϸ�� B, ����ϸ����¼ C" & vbCrLf & _
                   "        Where A.�걾id = B.�걾id And B.ϸ��id = C.ID And A.�걾id = [1] And A.����ҩ������� Is Not Null" & vbCrLf & _
                   "        Union all" & vbCrLf & _
                   "        Select 0 ID, c.���id �ϼ�id, d.������ || '(' || d.Ӣ���� || ')' ϸ����, nvl(c.���,' ') ������, c.ҩ������ ��ҩ����, c.������� ����, c.�ο�����" & vbCrLf & _
                   "        From ΢����ͿƬ���� A, ΢����ͿƬϸ�� B, ΢����ͿƬҩ�� C, ����ҩ�� D" & vbCrLf & _
                   "        Where a.�걾id = b.�걾id And b.Id = c.���id And c.ҩ��id = d.Id And a.�걾id = [1] And a.����ҩ������� Is Not Null)" & vbCrLf & _
                   " Connect By Prior ID = �ϼ�id" & vbCrLf & _
                   " Start With �ϼ�id Is Null"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ҩ�����", lngSmapleID)
          
          '������
4         With Me.VSFCJYM
              '������
5             .Rows = 1
6             Do While Not rsTmp.EOF
7                 .Rows = .Rows + 1
8                 .TextMatrix(.Rows - 1, .ColIndex("ϸ����")) = rsTmp("ϸ����") & ""
9                 .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
10                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
11                .TextMatrix(.Rows - 1, .ColIndex("��ҩ����")) = rsTmp("��ҩ����") & ""
12                .TextMatrix(.Rows - 1, .ColIndex("�ο�����")) = rsTmp("�ο�����") & ""
13                .TextMatrix(.Rows - 1, .ColIndex("Level")) = rsTmp("Level") & ""
                  
                  '���level=1,������Ϊ�����У�ϸ������level=2��ʾ�Ӽ��У������أ�
14                If Val(rsTmp("Level") & "") = 1 Then
                      '����
15                    .IsSubtotal(.Rows - 1) = True   '����Ϊ���νڵ�
16                    .RowOutlineLevel(.Rows - 1) = 3
                      '��ʾ�߿���
17                    .CellBorderRange .Rows - 1, .ColIndex("ϸ����"), .Rows - 1, .ColIndex("�ο�����"), vbBlack, 0, 0, 0, 1, 0, 0
18                End If
                  
19                rsTmp.MoveNext
20            Loop
              
              '��ȡVSF�߶�
21            If mblnHaveBoder = False Then
22                If mlngVsfHeight < (.Rows + 5) * .RowHeight(0) Then
23                    mlngVsfHeight = (.Rows + 5) * .RowHeight(0)
24                End If
25            End If
26        End With


27        Exit Sub
GetMicroorganisCJYM_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(GetMicroorganisCJYM)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
    Call mobjFTP.FuncFtpDisConnect   '�Ͽ�FTP����
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
        If lblPatient(0).Caption = "����:" Then lblPatient(0).Caption = "������Ŀ:"
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
'��    ��:������
'����ʱ��:2019-04-19
'��    ��:  ��ʾ���Ʋο�
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim objAdvice As Object
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemCode As String
          Dim strItemIDs As String

1         On Error GoTo ShowClincHelp_Error

2         If mlngSampleID <> 0 Then
3             If mintVersion = 25 Then    '�°���ȥ��ѯ
4                 strSQL = "Select f_List2str(Cast(Collect(b.���Ʊ��� || '') As t_Strlist)) ����" & vbCrLf & _
                         "   From ����������� A, ���������Ŀ B" & vbCrLf & _
                         "   Where A.���ID = b.id And a.�걾id = [1]"
5                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", mlngSampleID)
6                 If Not rsTmp.EOF Then
7                     strItemCode = rsTmp("����") & ""
8                 End If

9                 If strItemCode <> "" Then
                      'ͨ�����Ʊ����ѯ������ĿID
10                    strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                               "f_List2str(Cast(Collect(a.ID || '') As t_Strlist)) ID" & vbCrLf & _
                             " From ������ĿĿ¼ A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                             " Where A.���� = B.Column_Value"
11                    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strItemCode)
12                    If Not rsTmp.EOF Then strItemIDs = rsTmp("ID") & ""
13                End If
14            ElseIf mintVersion = 10 Then    '�ϰ���ȥ��ѯ
15                strSQL = " select f_List2str(Cast(Collect(b.������ĿID || '') As t_Strlist)) ������ĿID from ����걾��¼ A, ����ҽ����¼ B where a.ҽ��id=b.id and a.id=[1]"
16                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿID", mlngSampleID)
17                If Not rsTmp.EOF Then
18                    strItemIDs = rsTmp("������ĿID") & ""
19                End If
20            End If
21        End If



          '���ýӿ�
22        If Not rsTmp.EOF Then
23            If objAdvice Is Nothing Then
24                Set objAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
25                If Not objAdvice Is Nothing Then
26                    On Error Resume Next
27                    Call objAdvice.ShowClincHelp(1, Me, 0, False, strItemIDs)
28                    If Err.Number = 438 Then
29                        MsgBox "HIS�汾����", vbInformation, gSysInfo.AppName
30                        Exit Sub
31                    End If
32                End If
33            End If
34        End If


35        Exit Sub
ShowClincHelp_Error:
36        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(ShowClincHelp)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
37        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-12
'��    ��:  ���FTP�������Ƿ����
'��    ��:
'��    ��:
'��    ��:  True=���ã�False=������
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Function ConnFtp() As Boolean
          Dim strFTP As String
          Dim intReturn As Integer

1         On Error GoTo ConnFtp_Error

2         strFTP = ComGetPara(Sel_Lis_DB, "FTP����", 2500, 2500, "")
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
                  
                  '����ͼƬ����Ŀ¼
15                intReturn = mobjFTP.FuncFtpMkDir(mstrFtpFolder, "MicroorganismPicture")
16                If intReturn = 1 Then
17                    MsgBox "FTP����ʧ��", vbInformation, gSysInfo.AppName
18                    Exit Function
19                End If
20            End If
21        End If



22        Exit Function
ConnFtp_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(ConnFtp)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/7/24
'��    ��:ɾ������ʱ�䳬��30���ͼƬ�ļ�
'��    ��:
'��    ��:
'��    ��:
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
          '�ж��ļ����Ƿ����
4         If Not objFSO.FolderExists(strFolder) Then Exit Sub
          '�����ļ����µ������ļ����������ʱ����ڵ���30�죬��ɾ�����ļ�
5         Set objFolder = objFSO.GetFolder(strFolder)
6         Set objFiles = objFolder.Files
7         For Each objFile In objFiles
8             If DateDiff("d", objFile.DateCreated, dateNow) >= 30 Then
9                 objFSO.DeleteFile (objFile.Path)
10            End If
11        Next


12        Exit Sub
DeleteImg_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(DeleteImg)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear

End Sub

Private Sub vsfGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfGeneral
        lngRow = .MouseRow
        lngCol = .MouseCol
        If .ColIndex("�ٴ�����") < 0 Then Exit Sub
        If lngRow < 1 Or lngCol < 0 Then
            Call gobjHisComLib.zlCommFun.ShowTipInfo(0, "")
            Exit Sub
        End If
        Call gobjHisComLib.zlCommFun.ShowTipInfo(.hWnd, .TextMatrix(lngRow, .ColIndex("�ٴ�����")), True)
    End With
End Sub

Private Function AuditingSample(ByVal intType As Integer) As Boolean
      '����/ȡ������
      'intType    1=����,2=ȡ������

          Dim strSQL As String

1         On Error GoTo AuditingSample_Error

2         strSQL = "Zl_���鴫Ⱦ������_Edit(" & intType & "," & mlngSampleID & ",'" & gUserInfo.Name & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "��Ⱦ�����渴��")

4         SaveDBLog 18, 6, Val(mlngPaintID), IIf(intType = 1, "����", "ȡ������"), IIf(intType = 1, "����", "ȡ������"), 2500, "�ٴ�ʵ���ҹ���"

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
14        Call WriteErrLog("zlPublicHisCommLis", "frmAdviceReprotBrowse", "ִ��(AuditingSample)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear

End Function




