VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmOutMedRecEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ҳ"
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
   StartUpPosition =   1  '����������
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
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6480
         TabIndex        =   90
         Top             =   60
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5280
         TabIndex        =   89
         ToolTipText     =   "�ȼ���F2"
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
      TabCaption(0)   =   "������Ϣ"
      TabPicture(0)   =   "frmOutMedRecEdit.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "������Ϣ"
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
            Name            =   "����"
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
               Caption         =   " ����ժҪ "
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
               Caption         =   "����"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   74
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "��"
               Height          =   255
               Index           =   23
               Left            =   6705
               TabIndex        =   86
               TabStop         =   0   'False
               ToolTipText     =   "ѡ��(*)"
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
               Caption         =   "��Ⱦ���ϴ�(&U)"
               Height          =   195
               Index           =   1
               Left            =   3600
               TabIndex        =   75
               Top             =   30
               Width           =   1470
            End
            Begin VB.OptionButton optState 
               Caption         =   "����"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   73
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin MSMask.MaskEdBox txt�������� 
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
            Begin MSMask.MaskEdBox txt����ʱ�� 
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
               Caption         =   "����ҽѧ��ʾ"
               Height          =   180
               Index           =   33
               Left            =   45
               TabIndex        =   87
               Top             =   2100
               Width           =   1080
            End
            Begin VB.Label lblEdit 
               Caption         =   "ҽѧ��ʾ"
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
               Caption         =   "ȥ��"
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
               Caption         =   "������ַ"
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
               Caption         =   "����ʱ��"
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
               Caption         =   "����ҩƷĿ¼����(&1)"
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
               Caption         =   "���ݹ���Դ����(&2)"
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
               Caption         =   " ������¼ "
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
               Caption         =   "������ϱ�׼����(&3)"
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
               Caption         =   "���ݼ�����������(&4)"
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
               ToolTipText     =   "����������ɾ���ժҪ(F12)"
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
               Caption         =   " ��ϼ�¼ "
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
            Caption         =   "��"
            Height          =   255
            Index           =   9
            Left            =   6930
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3705
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "��"
            Height          =   255
            Index           =   6
            Left            =   6930
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2985
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "��"
            Height          =   255
            Index           =   5
            Left            =   6930
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2625
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "��"
            Height          =   240
            Index           =   19
            Left            =   6945
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2265
            Width           =   270
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "��"
            Height          =   255
            Index           =   17
            Left            =   6930
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   4413
            Width           =   285
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "��"
            Height          =   240
            Index           =   13
            Left            =   6945
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
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
         Begin MSMask.MaskEdBox txt����ʱ�� 
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
         Begin MSMask.MaskEdBox txt�������� 
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
            Caption         =   "�����"
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
            Caption         =   "����"
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
            Caption         =   "�Ա�"
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
            Caption         =   "����"
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
            Caption         =   "����״��"
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
            Caption         =   "ְҵ"
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
            Caption         =   "����"
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
            Caption         =   "����"
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
            Caption         =   "���֤��"
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
            Caption         =   "��λ����"
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
            Caption         =   "��λ�绰"
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
            Caption         =   "��λ�ʱ�"
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
            Caption         =   "��ͥ��ַ"
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
            Caption         =   "��ͥ�绰"
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
            Caption         =   "��ͥ�ʱ�"
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
            Caption         =   "���ʽ"
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
            Caption         =   "�����ص�"
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
            Caption         =   "��������"
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
            Caption         =   "�໤��"
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
            Caption         =   "����"
            Height          =   180
            Index           =   11
            Left            =   3555
            TabIndex        =   25
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl����֤�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����֤��"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   2295
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڵ�ַ"
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
            Caption         =   "�����ʱ�"
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
            Caption         =   "����"
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
            Caption         =   "�Ļ��̶�"
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
            Caption         =   "����״��"
            Height          =   180
            Index           =   29
            Left            =   120
            TabIndex        =   57
            Top             =   5526
            Width           =   720
         End
         Begin VB.Label lblEdit 
            Caption         =   "Ѫ��"
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
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   -64800
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
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
Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mlng�Һ�ID As Long
Private mlng����ID As Long
Private mint���� As Integer
Private mint���� As Integer
Private mstr������ As String
Private mbln��ҽ As Boolean
Private mstr����ID As String   '���ڱ��漲��ID,��FormClosed�¼��д��ݸ�������
Private mstr���ID As String   '���ڱ������ID,��FormClosed�¼��д��ݸ�������

Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mbln����ҩ��Edit As Boolean

Private mrsXYDiag  As ADODB.Recordset '��ҽ��ϼ�¼��
Private mrsZYDiag  As ADODB.Recordset '��ҽ��ϼ�¼��
Private mblnUseTYT As Boolean 'ʹ��̫Ԫͨ�ӿ�
Private mint����������Դ As Integer 'ҽ��վ�Ĺ���������Դ

Private Enum TXT_ENUM
    txt���� = 0
    txt����� = 1
    txt�໤�� = 2
    txt���� = 3
    txt���֤�� = 4
    txt�����ص� = 5
    txt��λ���� = 6
    txt��λ�绰 = 7
    txt��λ�ʱ� = 8
    txt��ͥ��ַ = 9
    txt��ͥ�绰 = 10
    txt��ͥ�ʱ� = 11
    txt����ժҪ = 12
    txt���� = 13
'    txt��� = 14
'    txt���� = 15
    txt����֤�� = 16
    txt���ڵ�ַ = 17
    txt���ڵ�ַ�ʱ� = 18
    txt���� = 19
    txt������ַ = 20
'    txt����ѹ = 21
'    txt����ѹ = 22
    txtҽѧ��ʾ = 23
    txt����ҽѧ��ʾ = 24
'    txt���� = 25
End Enum

Private Enum CBO_ENUM
    cbo�Ա� = 0
    cbo���� = 1
    cbo���� = 2
    cbo���� = 3
    cbo���� = 4
    cbo���� = 5
    cboְҵ = 6
    cbo�Ļ��̶� = 7
    cbo����״�� = 8
    cboRh = 9
    cboѪ�� = 10
    cboȥ�� = 11
End Enum

Private Enum CHK_ENUM
    chk��Ⱦ���ϴ� = 1
End Enum

Private Enum OPT_ENUM
    opt���� = 0
    opt���� = 1
End Enum

Private Enum COL_ENUM
    col���� = 0
    col���� = 1
    col��� = 2
    col��ҽ֤�� = 3
    col����ʱ�� = 4
    col���� = 5
    col���ID = 6
    col����ID = 7
    col֤��ID = 8
    colҽ��ID = 9
End Enum

Private Enum AllerColS
    AC_����ʱ�� = 0
    AC_����ҩ�� = 1
    AC_������Ӧ = 2
    AC_����Դ���� = 3
End Enum

Private mlngNum As Long
Private mlngSelNum As Long
Private mlngNumBack As Long
Private mstrEmail As String
Private mstrQQ As String

Public Function ShowMe(frmParent As Object, ByVal str�Һŵ� As String, ByVal strPrivs As String, Optional blnDiagnose As Boolean, Optional ByVal blnReadOnly As Boolean, _
Optional ByRef str����ID As String, Optional ByRef str���ID As String) As Boolean

'������blnDiagnose=�Ƿ����������д���
'���أ�blnDiagnose=�Ƿ���д�˲��˵����
    mblnReadOnly = blnReadOnly
    mblnDiagnose = blnDiagnose
    mstr�Һŵ� = str�Һŵ�
    mstrPrivs = strPrivs
    
    mstr����ID = ""
    mstr���ID = ""

    On Error Resume Next
    Me.Show 1, frmParent
    str����ID = mstr����ID
    str���ID = mstr���ID

    On Error GoTo 0
    
    blnDiagnose = mblnDiagnose
    ShowMe = mblnOk
End Function

Private Sub SetFaceEditable(ByVal blnReadOnly As Boolean)
'���ܣ����ݵ�ǰ�Ƿ�ֻ�������ý���Ŀɱ༭����
    Dim objControl As Object

    For Each objControl In Me.Controls
        If InStr("TextBox;MaskEdBox;ComboBox;CheckBox;VSFlexGrid", TypeName(objControl)) > 0 Then
            'TabStop=False��ʾ��ǰȷʵ���ɱ༭��
            If objControl.Container.Name = "fraInfo" And objControl.TabStop = True Then
                If TypeName(objControl) = "TextBox" And objControl.Enabled Then
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                    objControl.Locked = blnReadOnly
                ElseIf TypeName(objControl) = "MaskEdBox" Then
                    'û��Locked����,��Enabledʵ��
                    objControl.Enabled = Not blnReadOnly
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "ComboBox" And objControl.Enabled Then
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                    objControl.Locked = blnReadOnly
                ElseIf TypeName(objControl) = "CheckBox" Then
                    'û��Locked����,��Enabledʵ��
                    objControl.Enabled = Not blnReadOnly
                ElseIf TypeName(objControl) = "VSFlexGrid" Then
                    'ͬʱע��Ҫ�ڼ�������¼��н���һЩ����
                    objControl.Editable = IIf(blnReadOnly, flexEDNone, flexEDKbdMouse)
                    objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                    objControl.BackColorBkg = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                End If
            End If
        End If
    Next
End Sub

Private Function InitMedData() As Boolean
'���ܣ���ʼ���༭�����ͱ�Ҫ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    Call zlControl.CboSetHeight(cboEdit(cbo����), cboEdit(cbo����).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cbo����), cboEdit(cbo����).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cboְҵ), cboEdit(cboְҵ).Height * 16)
    vsDiagXY.MergeCol(0) = True
    vsDiagZY.MergeCol(0) = True
    
    Call SetCboFromList(Array("��", "��", "��"), cboEdit(cbo����), 0)
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From �Ա� Order by ����", Array(cbo�Ա�))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ҽ�Ƹ��ʽ Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ����״�� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ְҵ Order by ����", Array(cboְҵ))
    
    strSQL = "Select ����, ���� From ����ȥ��"
    cboEdit(cboȥ��).Clear
    cboEdit(cboȥ��).AddItem ("")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cboȥ��), rsTmp, False)
    
    Call SetCboFromList(Array("", "9-��ä�Ͱ���ä", "8-Сѧ��������ѧ��", "7-����", "6-����", "4-��ר", "3-��ר", "2-��ѧ", "1-�о���������"), cboEdit(cbo�Ļ��̶�), 0)
    Call SetCboFromList(Array("", "0-δ����", "1-����1̥", "2-����2̥������", "4-����"), cboEdit(cbo����״��), 0)
    Call SetCboFromList(Array("", "A��", "B��", "O��", "AB��", "����"), cboEdit(cboѪ��), 0)
    Call SetCboFromList(Array("", "��", "��", "����", "δ��"), cboEdit(cboRh), 0)
    optInput(0).TabStop = False: optInput(1).TabStop = False 'Ҫǿ�д���ִ��һ��
    
    InitMedData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadMedRec() As Boolean
'���ܣ���ȡ������ҳ�ĸ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long
    
    On Error GoTo errH
    
    '������Ϣ
    strSQL = "Select A.����ID,B.ID as �Һ�ID,B.ժҪ,B.����,a.����," & _
        " Nvl(Nvl(B.�������ID,Decode(B.ת��״̬,1,B.ת�����ID,NULL)),B.ִ�в���ID) as ����ID," & _
        " B.��Ⱦ���ϴ�,B.����ʱ��,B.������ַ,A.����,A.�����,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,A.��������,A.ҽ�Ƹ��ʽ," & _
        " A.����,A.����,A.����״��,A.ְҵ,A.���֤��,A.�����ص�,A.�໤��,A.��ͥ��ַ,A.��ͥ�绰," & _
        " A.����,A.��ͥ��ַ�ʱ�,A.������λ,A.��ͬ��λid,A.��λ�绰,A.��λ�ʱ�,B.����,C.������,A.����֤��,A.���ڵ�ַ,a.���ڵ�ַ�ʱ�,a.qq,a.email" & _
        " From ������Ϣ A,���˹Һż�¼ B,����������Ϣ C" & _
        " Where A.����ID=B.����ID And B.����ID=C.����ID(+) And B.����=C.����(+) And B.NO=[1] And B.��¼����=1 And B.��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    If rsTmp.EOF Then Exit Function
    
    mlng����ID = rsTmp!����ID
    mlng�Һ�ID = rsTmp!�Һ�ID
    mlng����ID = rsTmp!����ID
    mint���� = Nvl(rsTmp!����, 0)
    mint���� = Nvl(rsTmp!����, 0)
    mstr������ = Nvl(rsTmp!������)
    mbln��ҽ = Have��������(rsTmp!����ID, "��ҽ��")
    mstrEmail = Nvl(rsTmp!Email)
    mstrQQ = Nvl(rsTmp!QQ)
        
    txtEdit(txt����).Text = rsTmp!����
    txtEdit(txt����).Tag = rsTmp!���� '��¼ԭʼ������
    Call GetCboIndex(cboEdit(cbo�Ա�), Nvl(rsTmp!�Ա�))
    txtEdit(txt�����).Text = Nvl(rsTmp!�����)
    
    If Not IsNull(rsTmp!��������) Then
        txt��������.Text = Format(rsTmp!��������, "yyyy-MM-dd")
        If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
            txt����ʱ��.Text = Format(rsTmp!��������, "HH:mm")
        End If
    End If
        
    Call LoadOldData("" & rsTmp!����, txtEdit(txt����), cboEdit(cbo����))
    

    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!ҽ�Ƹ��ʽ))
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!����))
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!����))
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!����״��))
    Call GetCboIndex(cboEdit(cboְҵ), Nvl(rsTmp!ְҵ))
    txtEdit(txt����).Text = Nvl(rsTmp!����)
    txtEdit(txt����).Text = Nvl(rsTmp!����)
    txtEdit(txt�໤��).Text = Nvl(rsTmp!�໤��)
    txtEdit(txt���֤��).Text = Nvl(rsTmp!���֤��)
    txtEdit(txt����֤��).Text = Nvl(rsTmp!����֤��)
    txtEdit(txt�����ص�).Text = Nvl(rsTmp!�����ص�)
    txtEdit(txt��λ����).Text = Nvl(rsTmp!������λ)
    txtEdit(txt��λ����).Tag = Val("" & rsTmp!��ͬ��λid)
    If InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") = 0 And Not IsNull(rsTmp!��ͬ��λid) Then
        txtEdit(txt��λ����).Enabled = False
        cmdEdit(txt��λ����).Enabled = False
    End If
    
    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!��λ�绰)
    txtEdit(txt��λ�ʱ�).Text = Nvl(rsTmp!��λ�ʱ�)
    txtEdit(txt��ͥ��ַ).Text = Nvl(rsTmp!��ͥ��ַ)
    txtEdit(txt��ͥ�绰).Text = Nvl(rsTmp!��ͥ�绰)
    txtEdit(txt��ͥ�ʱ�).Text = Nvl(rsTmp!��ͥ��ַ�ʱ�)
    txtEdit(txt���ڵ�ַ).Text = Nvl(rsTmp!���ڵ�ַ)
    txtEdit(txt���ڵ�ַ�ʱ�).Text = Nvl(rsTmp!���ڵ�ַ�ʱ�)
    txtEdit(txt����ժҪ).Text = Nvl(rsTmp!ժҪ)
    If Nvl(rsTmp!����, 0) = 1 Then
        optState(opt����).Value = True
    End If
    chkEdit(chk��Ⱦ���ϴ�).Value = Nvl(rsTmp!��Ⱦ���ϴ�, 0)
    If Not IsNull(rsTmp!����ʱ��) Then
        txt��������.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd")
        txt����ʱ��.Text = Format(rsTmp!����ʱ��, "HH:mm")
        If txt����ʱ��.Text = "00:00" Then txt����ʱ��.Text = "__:__"
    End If
    txtEdit(txt������ַ).Text = Nvl(rsTmp!������ַ)
    
    '������Ϣ
    Call ucPatiVitalSigns.LoadPatiVitalSigns(mlng����ID, mlng�Һ�ID)
    strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And (����ID=[2] Or ����ID is Null) Order by Nvl(����ID,999999999)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    rsTmp.Filter = "��Ϣ��='�Ļ��̶�'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cbo�Ļ��̶�), Nvl(rsTmp!��Ϣֵ))
    rsTmp.Filter = "��Ϣ��='����״��'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cbo����״��), Nvl(rsTmp!��Ϣֵ))
    rsTmp.Filter = "��Ϣ��='ȥ��'"
    If Not rsTmp.EOF Then cboEdit(cboȥ��).Text = Nvl(rsTmp!��Ϣֵ)
    rsTmp.Filter = "��Ϣ��='Ѫ��'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cboѪ��), Nvl(rsTmp!��Ϣֵ))
    rsTmp.Filter = "��Ϣ��='RH'"
    If Not rsTmp.EOF Then Call GetCboIndex(cboEdit(cboRh), Nvl(rsTmp!��Ϣֵ))
    rsTmp.Filter = "��Ϣ��='ҽѧ��ʾ'"
    If Not rsTmp.EOF Then txtEdit(txtҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
    rsTmp.Filter = "��Ϣ��='����ҽѧ��ʾ'"
    If Not rsTmp.EOF Then txtEdit(txt����ҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
    '������Ϣ:���ιҺŵ�,������
    strSQL = "Select ��¼��Դ,NVL(����ʱ��,��¼ʱ��) as ����ʱ��,ҩ��ID,ҩ����,������Ӧ,����Դ���� From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by NVL(����ʱ��,��¼ʱ��),ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 2 '�̶���+����
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , AC_����ҩ��)
                End If
                If lngRow = -1 Then
                    .TextMatrix(i, AC_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, AC_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")  '���ڱ���
                    .TextMatrix(i, AC_����ҩ��) = Nvl(rsTmp!ҩ����)
                    .Cell(flexcpData, i, AC_����ҩ��) = .TextMatrix(i, AC_����ҩ��) '��������ָ�
                    .TextMatrix(i, AC_������Ӧ) = Nvl(rsTmp!������Ӧ)
                    .Cell(flexcpData, i, AC_������Ӧ) = .TextMatrix(i, AC_������Ӧ)   '��������ָ�
                    .TextMatrix(i, AC_����Դ����) = Nvl(rsTmp!����Դ����)
                    .RowData(i) = Val(Nvl(rsTmp!ҩ��ID, 0))
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    vsAller.Row = 1: vsAller.Col = AC_����ҩ��
    vsAller.Tag = "δ�޸�"
    
    '��ȡ�����Ϣ
    Call LoadPatiDiag(False)
    
    If Not mbln��ҽ Then
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
'���ܣ���ȡ����ʾ�������
'������blnLast=�Ƿ񲻶�ȡ���ξ������ϣ�����ȡ���һ�ξ�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    If blnLast Then
        strSQL = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1 " & _
                " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
        strSQL = "Select a.ID,a.��¼��Դ,a.�������,a.����ID,a.���ID,a.֤��ID,a.�������,a.�Ƿ�����, a.��¼����, a.��¼��,b.���� as ��������,c.���� as ��ϱ���,d.���� as ֤�����,A.����ʱ��,A.��ϴ���, " & _
                " (Select f_List2str(Cast(Collect(c.ҽ��id || '') As t_Strlist)) ҽ��id From �������ҽ�� C,����ҽ����¼ D where c.ҽ��id=d.id  and c.���id=a.id And d.ҽ��״̬<>-1 And D.ҽ��״̬<>4) as ҽ��ID  From �������ҽ�� C where c.���ID=A.ID ) ҽ��ID  " & _
                "From ������ϼ�¼  A, ��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
                " Where  a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��ID=d.ID(+) And a.��¼��Դ IN(1,3) And a.������� IN(1,11)" & _
                " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=(" & strSQL & ")" & _
                " Order by a.�������,a.��ϴ���,a.�������"
    Else
        strSQL = "Select a.ID,a.��¼��Դ,a.�������,a.����ID,a.���ID,a.֤��ID,a.�������,a.�Ƿ�����, a.��¼����, a.��¼��,b.���� as ��������,c.���� as ��ϱ���,d.���� as ֤�����,A.����ʱ��,A.��ϴ���, " & _
                " (Select f_List2str(Cast(Collect(c.ҽ��id || '') As t_Strlist)) ҽ��id From �������ҽ�� C,����ҽ����¼ D where c.ҽ��id=d.id and c.���id=a.id And d.ҽ��״̬<>-1 And D.ҽ��״̬<>4) as ҽ��ID " & _
                " From ������ϼ�¼  A, ��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
                " Where  a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��ID=d.ID(+) And a.��¼��Դ IN(1,3) And a.������� IN(1,11)" & _
                " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=[2]" & _
                " Order by a.�������,a.��ϴ���,a.�������"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    If Not rsTmp.EOF Then
        '��ҽ���
        rsTmp.Filter = "�������=1 And ��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "�������=1 And ��¼��Դ<>3": '������Դ����Ϊȱʡ��ʾ
        With vsDiagXY
            Set mrsXYDiag = zlDatabase.CopyNewRec(rsTmp)
            .Rows = rsTmp.RecordCount + 2
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!�������) Then
                    .TextMatrix(i, col����) = ""
                    .TextMatrix(i, col���) = ""
                Else
                    If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                        '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                        If Val(rsTmp!����id & "") <> 0 Then
                            .TextMatrix(i, col����) = Nvl(rsTmp!��������)
                        ElseIf Val(rsTmp!���id & "") <> 0 Then
                            .TextMatrix(i, col����) = Nvl(rsTmp!��ϱ���)
                        Else
                            .TextMatrix(i, col����) = ""
                        End If
                        .TextMatrix(i, col���) = rsTmp!�������
                    Else
                        .TextMatrix(i, col����) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                        .TextMatrix(i, col���) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                    End If
                End If
                If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                    .Cell(flexcpData, i, col���) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                Else
                    .Cell(flexcpData, i, col���) = .TextMatrix(i, col���)
                End If
                .TextMatrix(i, col����) = IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                .Cell(flexcpData, i, col����) = Val(rsTmp!ID & "")
                .TextMatrix(i, col���ID) = Nvl(rsTmp!���id, 0)
                .TextMatrix(i, col����ID) = Nvl(rsTmp!����id, 0)
                .TextMatrix(i, colҽ��ID) = rsTmp!ҽ��ID & ""
                .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                If Val(rsTmp!��ϴ��� & "") = 1 And .TextMatrix(i, col����ʱ��) <> "" Then
                    '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
                    txt��������.BackColor = vbButtonFace
                    txt��������.Enabled = False
                    txt����ʱ��.BackColor = vbButtonFace
                    txt����ʱ��.Enabled = False
                End If
                rsTmp.MoveNext
            Next
            .Cell(flexcpText, .FixedRows, col����, .Rows - 1, col����) = "��ҽ"
            .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
        End With
        '��ҽ���
        If mbln��ҽ Then
            rsTmp.Filter = "�������=11 And ��¼��Դ=3"
            If rsTmp.EOF Then rsTmp.Filter = "�������=11 And ��¼��Դ<>3"
            With vsDiagZY
                Set mrsZYDiag = zlDatabase.CopyNewRec(rsTmp)
                .Rows = rsTmp.RecordCount + 1
                For i = 0 To rsTmp.RecordCount - 1
                    If IsNull(rsTmp!�������) Then
                        .TextMatrix(i, col����) = ""
                        .TextMatrix(i, col���) = ""
                    Else
                        If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                            '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                            If Val(rsTmp!����id & "") <> 0 Then
                                .TextMatrix(i, col����) = Nvl(rsTmp!��������)
                            ElseIf Val(rsTmp!���id & "") <> 0 Then
                                .TextMatrix(i, col����) = Nvl(rsTmp!��ϱ���)
                            Else
                                .TextMatrix(i, col����) = ""
                            End If
                            .TextMatrix(i, col���) = rsTmp!�������
                        Else
                            .TextMatrix(i, col����) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                            .TextMatrix(i, col���) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                        End If
                    End If

                    .TextMatrix(i, col����) = IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                    .Cell(flexcpData, i, col����) = Val(rsTmp!ID & "")
                    .TextMatrix(i, col���ID) = Nvl(rsTmp!���id, 0)
                    .TextMatrix(i, col����ID) = Nvl(rsTmp!����id, 0)
                    .TextMatrix(i, col֤��ID) = Nvl(rsTmp!֤��id, 0)
                    .TextMatrix(i, colҽ��ID) = rsTmp!ҽ��ID & ""
                    .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                    If Val(rsTmp!��ϴ��� & "") = 1 And .TextMatrix(i, col����ʱ��) <> "" Then
                        '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
                        txt��������.BackColor = vbButtonFace
                        txt��������.Enabled = False
                        txt����ʱ��.BackColor = vbButtonFace
                        txt����ʱ��.Enabled = False
                    End If
                    'ȡ֤������
                    If InStr(.TextMatrix(i, col���), "(") > 0 And InStr(.TextMatrix(i, col���), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(i, col���), InStrRev(.TextMatrix(i, col���), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '��ȡ֤��
                        .TextMatrix(i, col��ҽ֤��) = strTmp
                        'ȥ�����������֤��
                        .TextMatrix(i, col���) = Mid(.TextMatrix(i, col���), 1, InStrRev(.TextMatrix(i, col���), "(") - 1)
                    Else
                       .TextMatrix(i, col��ҽ֤��) = ""
                    End If
                    
                    '����¼����ϵ������������Ҫȥ��֤����˴˾�������
                    If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                        .Cell(flexcpData, i, col���) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                    Else
                        .Cell(flexcpData, i, col���) = .TextMatrix(i, col���)
                    End If
                    rsTmp.MoveNext
                Next
                .Cell(flexcpText, .FixedRows, col����, .Rows - 1, col����) = "��ҽ"
                .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
            End With
        End If
    End If
    vsDiagXY.Row = vsDiagXY.FixedRows: vsDiagXY.Col = 0: vsDiagXY.Col = col���
    vsDiagZY.Row = vsDiagZY.FixedRows: vsDiagZY.Col = 0: vsDiagZY.Col = col���
    vsDiagXY.Tag = "δ�޸�"
    vsDiagZY.Tag = "δ�޸�"
    
    LoadPatiDiag = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMedRec(Optional blnDiagnose As Boolean) As Boolean
'���ܣ������ҳ�������ݺϷ���
'���أ�blnDiagnose=�Ƿ���д�����
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str���֤ As String, str�������� As String, lng�Ա� As Long
    Dim str���� As String, i As Long, j As Long, k As Long
    Dim str����IDs As String, str���IDs As String
    
    blnDiagnose = False
    curDate = zlDatabase.Currentdate
    
    '����Ҫ��������ݼ��
    '-----------------------------------------------------------------------------------------
    If InStr(mstrPrivs, "�޸Ļ�����Ϣ") > 0 Then
        arrInfo = Array(cbo����)
        arrName = Array("���ʽ")
        For i = 0 To UBound(arrInfo)
            If cboEdit(arrInfo(i)).Enabled And Not cboEdit(arrInfo(i)).Locked And cboEdit(arrInfo(i)).ListIndex = -1 Then
                Call ShowMessage(cboEdit(arrInfo(i)), "�������벡�˵�" & arrName(i) & "��")
                Exit Function
            End If
        Next
    End If
        
    '��Ŀ����ĳ��ȼ��
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "�������ݹ��������顣(����Ŀ������� " & objTmp.MaxLength & " ���ַ��� " & objTmp.MaxLength \ 2 & " ������)")
                Exit Function
            End If
        End If
    Next

    '�������ݵ���Ч�Լ��
    '-----------------------------------------------------------------------------------------
    '15������ӦΪδ��
    If Not (cboEdit(cbo����).Text = "" Or cboEdit(cbo����).ListIndex = -1) And IsDate(txt��������.Text) Then
        If DateDiff("yyyy", CDate(txt��������.Text), curDate) < 15 Then
            If InStr(cboEdit(cbo����).Text, "�ѻ�") > 0 _
                Or InStr(cboEdit(cbo����).Text, "ɥż") > 0 Or InStr(cboEdit(cbo����).Text, "���") > 0 Then
                Call ShowMessage(cboEdit(cbo����), "�ò�������̫С����ǰ��д�Ļ���״����Ϣ���ʺϡ�")
                Exit Function
            End If
        End If
    End If
            
    '���֤������
    '�����֤�Ž�����֤
    str���֤ = txtEdit(txt���֤��).Text
    If str���֤ <> "" Then
        If Len(str���֤) <> 15 And Len(str���֤) <> 18 Then
            Call ShowMessage(txtEdit(txt���֤��), "���֤����ĳ��Ȳ���ȷ��ӦΪ15λ��18λ��")
            Exit Function
        End If

        If Len(str���֤) = 15 Then
            str�������� = Mid(str���֤, 7, 6)
            str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
            lng�Ա� = Val(Right(str���֤, 1))
        Else
            str�������� = Mid(str���֤, 7, 8)
            str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
            lng�Ա� = Val(Mid(str���֤, 17, 1))
        End If
        If Not IsDate(str��������) Then
            If ShowMessage(txtEdit(txt���֤��), "���֤�����еĳ���������Ϣ����ȷ���Ƿ������", True) = vbNo Then Exit Function
        ElseIf IsDate(txt��������.Text) Then
            If Format(str��������, "yyyy-MM-dd") <> Format(txt��������.Text, "yyyy-MM-dd") Then
                If ShowMessage(txtEdit(txt���֤��), "���֤�����еĳ���������Ϣ�벡�˵ĳ������ڲ������Ƿ������", True) = vbNo Then Exit Function
            End If
        End If
        If (lng�Ա� Mod 2 = 1 And InStr(cboEdit(cbo�Ա�).Text, "Ů") > 0) Or (lng�Ա� Mod 2 = 0 And InStr(cboEdit(cbo�Ա�).Text, "��") > 0) Then
            If ShowMessage(txtEdit(txt���֤��), "���֤�����е��Ա���Ϣ�벡�˵��Ա𲻷����Ƿ������", True) = vbNo Then Exit Function
        End If
    End If
    
    '��ϱ��ļ��
    '-----------------------------------------------------------------------------------------
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col���)) <> "" Then
                If mint���� = 920 Then '����ҽ������Ҫ��
                    If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 82 Then
                        .Row = i: .Col = col���
                        Call ShowMessage(vsDiagXY, "�������̫����ֻ����82���ַ���41�����֡�")
                        Exit Function
                    End If
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 200 Then
                    .Row = i: .Col = col���
                    Call ShowMessage(vsDiagXY, "�������̫����ֻ����200���ַ���100�����֡�")
                    Exit Function
                End If
                If .TextMatrix(i, col����ʱ��) <> "" Then
                    If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col����ʱ��), "YYYY-MM-DD HH:mm") Then
                         .Row = i: .Col = col����ʱ��
                        Call ShowMessage(vsDiagXY, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                        Exit Function
                    End If
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, col���)) <> "" Then
                        If .TextMatrix(j, col���) = .TextMatrix(i, col���) Then
                            .Row = i: .Col = col���
                            Call ShowMessage(vsDiagXY, "���ִ���������ͬ�������Ϣ��")
                            Exit Function
                        ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                            If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                .Row = i: .Col = col���
                                Call ShowMessage(vsDiagXY, "���ִ���������ͬ�������Ϣ��")
                                Exit Function
                            End If
                        ElseIf Val(.TextMatrix(i, col���ID)) <> 0 Then
                            If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, col���ID)) Then
                                .Row = i: .Col = col���
                                Call ShowMessage(vsDiagXY, "���ִ���������ͬ�������Ϣ��")
                                Exit Function
                            End If
                        End If
                    End If
                Next
                
                If Val(.TextMatrix(i, col����ID)) <> 0 Then str����IDs = str����IDs & "," & Val(.TextMatrix(i, col����ID))
                If Val(.TextMatrix(i, col���ID)) <> 0 Then str���IDs = str���IDs & "," & Val(.TextMatrix(i, col���ID))
                
                blnDiagnose = True
            End If
        Next
    End With
        
    If mbln��ҽ Then
        With vsDiagZY
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col���)) <> "" Then
                    If mint���� = 920 Then '����ҽ������Ҫ��
                        If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 82 Then
                            .Row = i: .Col = col���
                            Call ShowMessage(vsDiagZY, "�������̫����ֻ����82���ַ���41�����֡�")
                            Exit Function
                        End If
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 200 Then
                        .Row = i: .Col = col���
                        Call ShowMessage(vsDiagZY, "�������̫����ֻ����200���ַ���100�����֡�")
                        Exit Function
                    End If
                    If .TextMatrix(i, col����ʱ��) <> "" Then
                        If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col����ʱ��), "YYYY-MM-DD HH:mm") Then
                             .Row = i: .Col = col����ʱ��
                            Call ShowMessage(vsDiagZY, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                            Exit Function
                        End If
                    End If
                    For j = i + 1 To .Rows - 1
                        If Trim(.TextMatrix(j, col���)) <> "" Then
                            If .TextMatrix(j, col���) = .TextMatrix(i, col���) Then
                                .Row = i: .Col = col���
                                Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
                                Exit Function
                            ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                                If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                    .Row = i: .Col = col���
                                    Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
                                    Exit Function
                                End If
                            ElseIf Val(.TextMatrix(i, col���ID)) <> 0 Then
                                '����ҽ��ϴ�֤��,�����޶�Ӧ֤��ID,���ID����ͬ
'                                If Val(.TextMatrix(j, col���ID)) & "," & Val(.TextMatrix(j, col֤��ID)) _
'                                    = Val(.TextMatrix(i, col���ID)) & "," & Val(.TextMatrix(i, col֤��ID)) Then
'                                    .Row = i: .Col = col���
'                                    Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
'                                    Exit Function
'                                End If
                            End If
                        End If
                    Next
                     '��ҽ��Ϻ���ҽ��ϵ�����¼��ҽ�����ܴ�����ͬ��
                    If .TextMatrix(i, col����) = "" Then
                        For k = vsDiagXY.FixedRows To vsDiagXY.Rows - 1
                            If Trim(vsDiagXY.TextMatrix(k, col���)) <> "" And vsDiagXY.TextMatrix(k, col����) = "" Then
                                If vsDiagXY.TextMatrix(k, col���) = .TextMatrix(i, col���) Then
                                    .Row = i: .Col = col���
                                    Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
                                    Exit Function
                                End If
                            End If
                        Next
                    End If
                    If Val(.TextMatrix(i, col����ID)) <> 0 Then str����IDs = str����IDs & "," & Val(.TextMatrix(i, col����ID))
                    If Val(.TextMatrix(i, col���ID)) <> 0 Then str���IDs = str���IDs & "," & Val(.TextMatrix(i, col���ID))
                    
                    blnDiagnose = True
                End If
            Next
        End With
        
        
    End If
    
    '����ҩ������
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, 1)) > 60 Then
                    .Row = i: .Col = 1
                    Call ShowMessage(vsAller, "����ҩ����̫����ֻ����60���ַ���30�����֡�")
                    Exit Function
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, 2)) > 100 Then
                    .Row = i: .Col = 2
                    Call ShowMessage(vsAller, "������Ӧ����̫����ֻ����100���ַ���50�����֡�")
                    Exit Function
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, 1)) <> "" Then
                        If .TextMatrix(j, 1) = .TextMatrix(i, 1) Then
                            .Row = i: .Col = 1
                            Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                            Exit Function
                        ElseIf .RowData(i) <> 0 Then
                            If .RowData(j) = .RowData(i) Then
                                .Row = i: .Col = 1
                                Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    '����ʱ����
    If txt��������.Text <> "____-__-__" Then
        If Not IsDate(txt��������.Text) Then
            Call ShowMessage(txt��������, "��������ȷ�ķ������ڡ�")
            Exit Function
        Else
            If txt����ʱ��.Text <> "__:__" Then
                If Not IsDate(txt����ʱ��.Text) Then
                    Call ShowMessage(txt����ʱ��, "��������ȷ�ķ���ʱ�䡣")
                    Exit Function
                End If
            End If
            
            If txt��������.Text & IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text) _
                > Format(curDate, txt��������.Format & IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Format)) Then
                Call ShowMessage(txt��������, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                Exit Function
            End If
        End If
    End If
    
    mstr����ID = Mid(str����IDs, 2)
    mstr���ID = Mid(str���IDs, 2)
    
    CheckMedRec = True
End Function

Private Function SaveMedRec() As Boolean
'���ܣ�����������ҳ�ĸ�����Ϣ
    Dim arrSQL As Variant, i As Integer
    Dim curDate As Date, intIdx As Integer
    Dim str���� As String, str���� As String
    Dim lng��λID As Long
    Dim blnTrans As Boolean
    Dim str����״�� As String
    Dim str�Ļ��̶� As String
    Dim blnDiagChange As Boolean
    Dim strFilter As String, strTmp As String
    Dim str����ҽ��ID As String
    
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    
    If IsDate(txt��������.Text) Then
        If IsDate(txt����ʱ��.Text) Then
            str���� = "To_Date('" & Format(txt��������.Text & " " & txt����ʱ��.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str���� = "To_Date('" & Format(txt��������.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        End If
    Else
       str���� = "NULL"
    End If
    
    If Trim(txtEdit(txt��λ����).Text) <> "" Then
        lng��λID = Val(txtEdit(txt��λ����).Tag)
    End If
    
    '������Ϣ
    str���� = "NULL"
    If IsDate(txt��������.Text) Then
        If IsDate(txt����ʱ��.Text) Then
            str���� = "To_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI')"
        Else
            str���� = "To_Date('" & txt��������.Text & "','YYYY-MM-DD')"
        End If
    End If
    '�Ļ��̶�
    If cboEdit(cbo�Ļ��̶�).ListIndex > 0 Then
        str�Ļ��̶� = Mid(cboEdit(cbo�Ļ��̶�), 1, InStr(cboEdit(cbo�Ļ��̶�), "-") - 1)
    End If
    '����״��
    If cboEdit(cbo����״��).ListIndex > 0 Then
        str����״�� = Mid(cboEdit(cbo����״��), 1, InStr(cboEdit(cbo����״��), "-") - 1)
    End If
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_������Ϣ_��ҳ����(" & _
        mlng����ID & ",'" & txtEdit(txt�����).Text & "','" & txtEdit(txt����).Text & "'," & _
        "'" & NeedName(cboEdit(cbo�Ա�).Text) & "','" & txtEdit(txt����).Text & cboEdit(cbo����).Text & "'," & _
        str���� & ",'" & txtEdit(txt�����ص�).Text & "','" & txtEdit(txt���֤��).Text & "'," & _
        "'" & NeedName(cboEdit(cbo����).Text) & "','" & NeedName(cboEdit(cbo����).Text) & "','" & txtEdit(txt����).Text & "'," & _
        "'" & NeedName(cboEdit(cbo����).Text) & "','" & NeedName(cboEdit(cboְҵ).Text) & "'," & _
        "'" & NeedName(cboEdit(cbo����).Text) & "','" & txtEdit(txt��ͥ��ַ).Text & "'," & _
        "'" & txtEdit(txt��ͥ�绰).Text & "','" & txtEdit(txt��ͥ�ʱ�).Text & "'," & _
        "'" & txtEdit(txt��λ����).Text & "','" & txtEdit(txt��λ�绰).Text & "'," & _
        "'" & txtEdit(txt��λ�ʱ�).Text & "',Null,Null,Null,Null,'" & txtEdit(txt�໤��).Text & "','" & mstr�Һŵ� & "'," & _
        IIf(optState(opt����).Value, 1, 0) & ",'" & txtEdit(txt����ժҪ).Text & "'," & chkEdit(chk��Ⱦ���ϴ�).Value & "," & str���� & ",'" & _
        Trim(txtEdit(txt����֤��).Text) & "'," & ZVal(lng��λID) & ",'" & txtEdit(txt���ڵ�ַ).Text & "','" & txtEdit(txt���ڵ�ַ�ʱ�).Text & "','" & _
        txtEdit(txt����).Text & "','" & mstrEmail & "','" & mstrQQ & "','" & txtEdit(txt������ַ).Text & "')"
    
    '������Ϣ
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = ucPatiVitalSigns.GetSaveSQL(mlng����ID, mlng�Һ�ID)
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'�Ļ��̶�','" & str�Ļ��̶� & "'," & mlng�Һ�ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'����״��','" & str����״�� & "'," & mlng�Һ�ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'ȥ��','" & cboEdit(cboȥ��).Text & "'," & mlng�Һ�ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'RH','" & cboEdit(cboRh).Text & "'," & mlng�Һ�ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'Ѫ��','" & cboEdit(cboѪ��).Text & "'," & mlng�Һ�ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'ҽѧ��ʾ','" & txtEdit(txtҽѧ��ʾ).Text & "'," & mlng�Һ�ID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'����ҽѧ��ʾ','" & txtEdit(txt����ҽѧ��ʾ).Text & "'," & mlng�Һ�ID & ")"
    '����ҩ��
    If vsAller.Tag = "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3)"
        With vsAller
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, AC_����ҩ��)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = _
                        "zl_���˹�����¼_Insert(" & mlng����ID & "," & mlng�Һ�ID & "," & _
                        "3," & ZVal(.RowData(i)) & ",'" & .TextMatrix(i, AC_����ҩ��) & "',1," & _
                        "To_Date('" & .Cell(flexcpData, i, AC_����ʱ��) & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & .TextMatrix(i, AC_������Ӧ) & "','" & .TextMatrix(i, AC_����Դ����) & "')"
                End If
            Next
        End With
    End If
    
    '��ϼ�¼
    If mbln��ҽ And vsDiagZY.Tag = "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,Null,'11')"
    End If
    If vsDiagXY.Tag = "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,Null,'1')"
        With vsDiagXY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col���)) <> "" Then
                    blnDiagChange = True
                    If Val(.Cell(flexcpData, i, col����) & "") > 0 Then
                        strFilter = "�������=1 And ��¼��Դ=3 And ����id=" & ZVal(.TextMatrix(i, col����ID)) & " And ���id=" & ZVal(.TextMatrix(i, col���ID))

                        strTmp = IIf(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���)
                        strFilter = strFilter & " And �������= '" & strTmp & "'"
                        If IsDate(.TextMatrix(i, col����ʱ��)) Then
                            strFilter = strFilter & " And  ����ʱ��= '" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "'"
                        Else
                            strFilter = strFilter & " And  ����ʱ��= Null "
                        End If
                        
                        strFilter = strFilter & " And �Ƿ�����=" & IIf(.TextMatrix(i, col����) = "", 0, 1)
                        mrsXYDiag.Filter = strFilter
                        blnDiagChange = mrsXYDiag.EOF
                    End If
                     str����ҽ��ID = IIf(.TextMatrix(i, colҽ��ID) = "", "Null", "'" & .TextMatrix(i, colҽ��ID) & "'")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    If blnDiagChange Then
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                            " Null,1," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & ",Null," & _
                            "'" & IIf(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & "',Null,Null," & IIf(.TextMatrix(i, col����) = "", 0, 1) & "," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str����ҽ��ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & UserInfo.���� & "')"
                    Else
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                            " Null,1," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & ",Null," & _
                            "'" & IIf(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & "',Null,Null," & IIf(.TextMatrix(i, col����) = "", 0, 1) & "," & _
                            "To_Date('" & Format(CDate(mrsXYDiag!��¼����), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str����ҽ��ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & mrsXYDiag!��¼�� & "')"
                    
                    End If
                End If
            Next
        End With
    End If
    
    If mbln��ҽ And vsDiagZY.Tag = "" Then
        With vsDiagZY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col���)) <> "" Then
                    blnDiagChange = True
                    If Val(.Cell(flexcpData, i, col����) & "") > 0 Then
                        strFilter = "�������=11 And ��¼��Դ=3 And ����id=" & ZVal(.TextMatrix(i, col����ID)) & " And ���id=" & ZVal(.TextMatrix(i, col���ID))

                        strTmp = IIf(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & "(" & .TextMatrix(i, col��ҽ֤��) & ")"
                        strFilter = strFilter & " And �������= '" & strTmp & "'"

                        strFilter = strFilter & " And  ֤��ID= " & ZVal(.TextMatrix(i, col֤��ID))
                        If IsDate(.TextMatrix(i, col����ʱ��)) Then
                            strFilter = strFilter & " And  ����ʱ��= '" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "'"
                        Else
                            strFilter = strFilter & " And  ����ʱ��= Null "
                        End If
                        
                        strFilter = strFilter & " And �Ƿ�����=" & IIf(.TextMatrix(i, col����) = "", 0, 1)
                        mrsXYDiag.Filter = strFilter
                        blnDiagChange = mrsZYDiag.EOF
                    End If
                    
                    str����ҽ��ID = IIf(.TextMatrix(i, colҽ��ID) = "", "Null", "'" & .TextMatrix(i, colҽ��ID) & "'")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    If blnDiagChange Then
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                            "Null,11," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & "," & _
                            ZVal(.TextMatrix(i, col֤��ID)) & ",'" & IIf(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & "(" & .TextMatrix(i, col��ҽ֤��) & ")" & "',Null,Null," & _
                            IIf(.TextMatrix(i, col����) = "", 0, 1) & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str����ҽ��ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & UserInfo.���� & "')"
                    Else
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                            "Null,11," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & "," & _
                            ZVal(.TextMatrix(i, col֤��ID)) & ",'" & IIf(.TextMatrix(i, col����) <> "", "(" & .TextMatrix(i, col����) & ")", "") & .TextMatrix(i, col���) & "(" & .TextMatrix(i, col��ҽ֤��) & ")" & "',Null,Null," & _
                            IIf(.TextMatrix(i, col����) = "", 0, 1) & ",To_Date('" & Format(CDate(mrsZYDiag!��¼����), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & str����ҽ��ID & "," & intIdx & ",Null,Null,To_date('" & Format(.TextMatrix(i, col����ʱ��), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & mrsZYDiag!��¼�� & "')"
                    
                    End If
                End If
            Next
        End With
    End If
    
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    
    '��������ͬ��
    If Not gobjCommunity Is Nothing And mint���� <> 0 Then
        If Not gobjCommunity.UpdateInfo(glngSys, p����ҽ��վ, mint����, mstr������, mlng����ID, mlng�Һ�ID) Then
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
'˵����ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, blnLevel As Boolean
    Dim strResult As String
    
    'ʹ��Lock�ķ�ʽ,������Enabled�ķ�ʽ
    If Not cmdEdit(Index).Enabled Or txtEdit(Index).Locked Then
        If txtEdit(Index).Enabled Then txtEdit(Index).SetFocus
        Exit Sub
    End If
    
    Select Case Index
        Case txt�����ص�, txt��ͥ��ַ, txt���ڵ�ַ
            'ѡ���������
            strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!����
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt��λ����
            'ѡ��λ��Ϣ
            strSQL = "Select ID,�ϼ�ID,ĩ��,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��" & _
                " From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "��Լ��λ", , , , , True, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""��Լ��λ""���ݣ����ȵ���Լ��λ���������á�", vbInformation, gstrSysName
                End If
                txtEdit(Index).Tag = ""
                If txtEdit(Index).Enabled Then txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") > 0 Then txtEdit(Index).Tag = Val(rsTmp!ID)
                If txtEdit(txt��λ�绰).Text = "" Then
                    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
                If txtEdit(Index).Enabled Then txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt����, txt����
            'ѡ����������
            strSQL = "Select Nvl(����,0) as ���� From ���� Group by Nvl(����,0)"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.RecordCount > 1 Then blnLevel = True
            
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            If blnLevel Then
                strSQL = _
                    " Select ID,�ϼ�id,����,����,����,ĩ��" & _
                    " From (Select ���� As ID,RPad(Substr(����,1,Decode(Nvl(����,0),0,0,1,2,4)),6,'0') As �ϼ�id," & _
                    "       ����,����,����,Decode(Nvl(����,0),2,1,3,1,0) as ĩ��" & _
                    "       From ���� Order By ����)" & _
                    " Start With �ϼ�ID Is Null Connect By Prior ID=�ϼ�id"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "����", , , , , , , vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            Else
                strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!����
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txtҽѧ��ʾ
            'ѡ��ҽѧ��ʾ
            On Error GoTo errH
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            strSQL = "Select Rownum ID,����,����,���� From ҽѧ��ʾ Order by ����"
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, True, True)

            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""ҽѧ��ʾ""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                While Not rsTmp.EOF
                    strResult = strResult & "," & rsTmp!����
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
            If .TextMatrix(i, col���) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, col���) & IIf(.TextMatrix(i, col����) <> "", "(��)", "")
            End If
        Next
    End With
    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col���) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, col���) & IIf(.TextMatrix(i, col����) <> "", "(��)", "")
            End If
        Next
    End With
    If strLog <> "" Then
        If txtEdit(txt����ժҪ).SelStart = 0 And txtEdit(txt����ժҪ).SelLength = Len(txtEdit(txt����ժҪ).Text) Then
            txtEdit(txt����ժҪ).SelStart = Len(txtEdit(txt����ժҪ).Text)
        End If
        i = txtEdit(txt����ժҪ).SelStart
        txtEdit(txt����ժҪ).SelText = Mid(strLog, 2)
        txtEdit(txt����ժҪ).SelStart = i
        txtEdit(txt����ժҪ).SelLength = Len(Mid(strLog, 2))
    End If
    
    txtEdit(txt����ժҪ).SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim blnDiagnose As Boolean
    
    If Not CheckMedRec(blnDiagnose) Then Exit Sub
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("���˵������Ϣ��û�����룬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
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
        
    optInput(Val(zlDatabase.GetPara("�����������", glngSys, p����ҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0))).Value = True
    
    '���������Դ
    If gint�����Դ > 1 Then
        optInput(0).Enabled = False
        optInput(1).Enabled = False
        If gint�����Դ = 2 Then
            optInput(0).Value = True
        ElseIf gint�����Դ = 3 Then
            optInput(1).Value = True
        End If
    End If
    
    If Not InitMedData Then Unload Me: Exit Sub
    If Not LoadMedRec Then Unload Me: Exit Sub
    If mblnReadOnly Then
        Call SetFaceEditable(True)
        cmdOK.Visible = False
    Else
        If InStr(mstrPrivs, "�޸Ļ�����Ϣ") = 0 Then

            If cboEdit(cbo����).ListIndex <> -1 Then
                cboEdit(cbo����).BackColor = Me.BackColor: cboEdit(cbo����).Locked = True: cboEdit(cbo����).TabStop = False
            End If
        End If
    End If
    '���˻�����Ϣ���������Ա����䣬�������ڲ������޸�
    txtEdit(txt����).BackColor = Me.BackColor: txtEdit(txt����).Locked = True: txtEdit(txt����).TabStop = False
    cboEdit(cbo�Ա�).BackColor = Me.BackColor: cboEdit(cbo�Ա�).Locked = True: cboEdit(cbo�Ա�).TabStop = False
    txt��������.BackColor = Me.BackColor: txt��������.Enabled = False
    txt����ʱ��.BackColor = Me.BackColor: txt����ʱ��.Enabled = False
    txtEdit(txt����).BackColor = Me.BackColor: txtEdit(txt����).Locked = True: txtEdit(txt����).TabStop = False
    cboEdit(cbo����).BackColor = Me.BackColor: cboEdit(cbo����).Locked = True: cboEdit(cbo����).TabStop = False

    '���ù���������Դ�ؼ�����
    mblnUseTYT = False
    Call SetContolsByAllerPara(p����ҽ��վ, mint����������Դ, optAller(0), optAller(1))

    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("����رմ��壬�������ĸ��Ľ����ᱣ�档Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    Call zlDatabase.SetPara("�����������", IIf(optInput(0).Value, 0, 1), glngSys, p����ҽ��վ, InStr(mstrPrivs, "��������") > 0)
    Call zlDatabase.SetPara("����������Դ", IIf(optAller(0).Value, "0", "1"), glngSys, p����ҽ��վ, optAller(0).Enabled And optAller(0).Visible)
End Sub

Private Sub optAller_Click(Index As Integer)
'�ܴ���CLiCK�¼�˵��gbytPass=3,gint����������Դ=0���ֻ���жϿؼ���ֵ����
    mblnUseTYT = Index = 1
End Sub

Private Sub optInput_LostFocus(Index As Integer)
    optInput(0).TabStop = False: optInput(1).TabStop = False 'Ҫǿ�д���ִ��һ��
End Sub

Private Sub optState_Click(Index As Integer)
    Dim blnDo As Boolean
    
    If Visible Then
        '����������δ¼�����������Զ���ȡ�ϴ����
        If Index = opt���� Then
            If chkEdit(Index).Value = 1 Then
                blnDo = vsDiagXY.Rows = vsDiagXY.FixedRows + 1 And vsDiagZY.Rows = vsDiagZY.FixedRows + 1
                If blnDo Then blnDo = blnDo And vsDiagXY.TextMatrix(vsDiagXY.FixedRows, col���) = "" And vsDiagZY.TextMatrix(vsDiagZY.FixedRows, col���) = ""
                If blnDo Then Call LoadPatiDiag(True)
            End If
        End If
        
        mblnChange = True
    End If
End Sub

Private Sub timThis_Timer()
    Dim lngSelNum As Long
    
    If vsAller.Col = AC_����ʱ�� Then
        lngSelNum = vsAller.EditSelStart
        If lngSelNum <> mlngSelNum And lngSelNum <> 16 And lngSelNum <> 0 Then
            Call Vs_EditSelChange(lngSelNum)
            mlngSelNum = lngSelNum
        End If
    End If
End Sub

Private Sub Vs_EditSelChange(ByVal lngSelNum As Long)
'���û��л�����ʱ�򴥷�
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
    If Index <> txt����ժҪ Then
        Call zlControl.TxtSelAll(txtEdit(Index))
    ElseIf txtEdit(Index).SelLength = 0 Then
        Call zlControl.TxtSelAll(txtEdit(Index))
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Index = txtҽѧ��ʾ Then
            txtEdit(txtҽѧ��ʾ) = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt�����ص� Or Index = txt��ͥ��ַ Or Index = txt���ڵ�ַ) And txtEdit(Index).Text <> "" Then
            '�����������
            strSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!����
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt��λ���� And txtEdit(Index).Text <> "" Then
            '���빤����λ
            strSQL = "Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������λ", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") > 0 Then txtEdit(Index).Tag = Val(rsTmp!ID)
                If txtEdit(txt��λ�绰).Text = "" Then
                    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
            Else
                txtEdit(Index).Tag = ""
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf (Index = txt���� Or Index = txt����) And txtEdit(Index).Text <> "" Then
            '������������
            strSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(Index = txt����, "����", "����"), False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!����
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Index = txtҽѧ��ʾ Then
            txtEdit(txtҽѧ��ʾ).Text = ""
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '�ǿ��ư���
        If Index = txtҽѧ��ʾ Then
            KeyAscii = 0
        End If
        'ѡ���ݼ�
        If KeyAscii = Asc("*") Then
            'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
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
        
        '�������볤��
        If txtEdit(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(Index).Text) > txtEdit(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '������������
        Select Case Index
'            Case txt���� '��������¼����
'                strMask = "1234567890"
            'Case txt�������� 'MaskEdit������
                'strMask = "1234567890-"
            Case txt���֤��
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Case txt��ͥ�绰, txt��λ�绰
                strMask = "1234567890-()"
            Case txt��ͥ�ʱ�, txt��λ�ʱ�, txt���ڵ�ַ�ʱ�
                strMask = "1234567890"
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    Else
        If Index = txtҽѧ��ʾ Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt��������_Change()
    If Visible Then mblnChange = True
    
    If IsDate(txt��������.Text) Then
        txt����ʱ��.Enabled = True
    Else
        txt����ʱ��.Enabled = False
        txt����ʱ��.Text = "__:__"
    End If
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        txt��������.Text = "____-__-__": Cancel = True
    End If
End Sub

Private Sub txt����ʱ��_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = AC_����ҩ�� Then mbln����ҩ��Edit = True '����س���λ����һ����Ԫ�������
    Call vsAller_AfterRowColChange(-1, -1, Row, Col)
    If Col = AC_����ҩ�� Then mbln����ҩ��Edit = False '����س���λ����һ����Ԫ�������
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAller
        If NewCol = AC_����ҩ�� Then
            If Not mbln����ҩ��Edit Then
                .ComboList = "..."
                .FocusRect = flexFocusSolid
            Else '����س���λ����һ����Ԫ�������
                .ComboList = ""
                .FocusRect = flexFocusSolid
            End If
        Else
            .FocusRect = IIf(Trim(vsAller.TextMatrix(NewRow, AC_����ҩ��)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     If Col = AC_����ʱ�� And Trim(vsAller.Cell(flexcpData, Row, AC_����ҩ��)) = "" Then Cancel = True
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int�Ա� As Integer
    
    With vsAller
        If mblnUseTYT Then
            strSQL = gobjPass.inputAllergy()
            If strSQL <> "" Then
                Call SetAllerInput(Row, , strSQL)
                Call AllerEnterNextCell
            End If
        Else
            If cboEdit(cbo�Ա�).Text Like "*��*" Then
                int�Ա� = 1
            ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                int�Ա� = 2
            End If
            
            strSQL = _
                " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
                " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
                " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                " Union All" & _
                " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
                " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                " From ������ĿĿ¼ A,ҩƷ���� B" & _
                " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
                IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
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
            If .TextMatrix(.Row, AC_����ҩ��) <> "" Then
                If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    .Tag = ""
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
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
            If .Col = AC_����ҩ�� And mblnUseTYT Then KeyAscii = 0: Exit Sub
        End If
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AC_����ҩ�� Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
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
        If Col = AC_����ʱ�� Then
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
        ElseIf Col = AC_����ҩ�� Then
            If KeyAscii <> 13 Then
                If mblnUseTYT Then KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = AC_����ҩ�� Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
    ElseIf Col = AC_����ʱ�� Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = 4
        mlngNum = 0
        mlngNumBack = 0
        timThis.Enabled = True
    End If
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AC_������Ӧ And Trim(vsAller.TextMatrix(Row, AC_����ҩ��)) = "" Then Cancel = True
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int�Ա�  As Integer
    Dim curDate As Date
    
    With vsAller
        If Col = AC_����ҩ�� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call AllerEnterNextCell
            Else
                strInput = UCase(.EditText)
                If cboEdit(cbo�Ա�).Text Like "*��*" Then
                    int�Ա� = 1
                ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                    int�Ա� = 2
                End If
                strSQL = _
                    " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                    " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                    " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                    " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                    IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                    Decode(gint����, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " Order by A.����"
                
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%", int�Ա�, gint���� + 1)
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    Call SetAllerInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call AllerEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = AC_����ʱ�� Then
            If Not IsDate(.EditText) And .EditText <> "" Then
                MsgBox "����������ڸ�ʽ����ȷ����ʽ�磺2010-10-10 18:30��"
                Cancel = True
                .EditText = vsAller.TextMatrix(Row, Col)
            Else
                If .EditText <> "" Then
                    curDate = zlDatabase.Currentdate
                    If CDate(.EditText) > curDate Then
                        MsgBox "����������ڲ��ܴ��ڵ�ǰʱ�䡣��ǰʱ�䣺" & curDate & "��"
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
        If Col = col��� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
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
            If NewCol = col��� Then
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
    If Col = col���� Then Cancel = True
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str�Ա� As String
    
    With vsDiagXY
        If optInput(0).Value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", mlng����ID, , True, False)
        Else
            'D-ICD-10��������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D", mlng����ID, cboEdit(cbo�Ա�).Text, True)
        End If
        If rsTmp Is Nothing Then
            If optInput(0).Value Then
                MsgBox "û�м���������ݿ���ѡ��", vbInformation, gstrSysName
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
            If .Col = col��� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col���) <> "" Then
                If .TextMatrix(.Row, colҽ��ID) = "" Then
                    If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call CreatePlugInOK(p����ҽ��վ)
                        'ɾ����/��Ҫ��Ϻ������ҽӿ�
                        If Not gobjPlugIn Is Nothing Then
                            On Error Resume Next
                            Call gobjPlugIn.DiagnosisDeleted(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(.Row, col���ID)), .TextMatrix(.Row, col���))
                            Call zlPlugInErrH(err, "DiagnosisDeleted")
                            err.Clear: On Error GoTo 0
                        End If
                        .RemoveItem .Row
                        mblnChange = True
                        .Tag = ""
                    End If
                Else
                    MsgBox "����϶�Ӧ�Ĵ����ѷ��ͣ�����ɾ����", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagXY)
        ElseIf KeyAscii = 32 And (.Col = col����) Then
            If DiagCellEditable(vsDiagXY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col���� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
            End If
        Else
            If .Col = col��� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
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
    ElseIf Col = col���� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Function GetZYSQL(ByRef strInput As String, ByRef strSQL As String, ByRef str�Ա� As String, Optional ByVal strType As String) As String
'���ܣ���ò�ѯ��ҽ��ϵ�SQL
'������strInput-��ѯ����,strsql--���ص�SQL��str�Ա�--���˵��Ա�  ,strType�����������ࡣ
'���أ�strsql--��ѯ��ҽ��ϵ�SQL
    If optInput(0).Value And strType <> "Z" Then
        '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
        Else
            strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
        End If
        strSQL = _
            " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
            " From �������Ŀ¼ A,������ϱ��� B" & _
            " Where A.ID=B.���ID And A.���=2" & _
            " And B.����=[4] And (" & strSQL & ")" & _
            " Order by A.����"
    Else
        If cboEdit(cbo�Ա�).Text Like "*��*" Then
            str�Ա� = "��"
        ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
            str�Ա� = "Ů"
        End If
        'B-��ҽ��������
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "���� Like [2]" '���뺺��ʱֻƥ������
        Else
            strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gint���� = 0, "����", "�����") & " Like [2]"
        End If
        strSQL = _
            " Select ID,ID as ��ĿID,����,����,����," & IIf(gint���� = 0, "����", "����� as ����") & ",˵��" & _
            " From ��������Ŀ¼" & _
            " Where ���='" & IIf(strType = "", "B", strType) & "' And (" & strSQL & ")" & _
            IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
            " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by ����"
    End If
    GetZYSQL = strSQL
End Function

Private Function GetXYSQL(ByRef strInput As String, ByRef strSQL As String, ByRef str�Ա� As String) As String
'���ܣ���ò�ѯ��ҽ��ϵ�SQL
'������strInput-��ѯ����,strsql--���ص�SQL��str�Ա�--���˵��Ա�
'���أ�strsql--��ѯ��ҽ��ϵ�SQL
    If optInput(0).Value Then
        '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
        Else
            strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
        End If
        strSQL = _
            " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
            " From �������Ŀ¼ A,������ϱ��� B" & _
            " Where A.ID=B.���ID And A.���=1" & _
            " And B.����=[4] And (" & strSQL & ")" & _
            " Order by A.����"
    Else
        If cboEdit(cbo�Ա�).Text Like "*��*" Then
            str�Ա� = "��"
        ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
            str�Ա� = "Ů"
        End If
        'D-ICD-10��������
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = "���� Like [2]" '���뺺��ʱ,ֻƥ������
        Else
            strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gint���� = 0, "����", "�����") & " Like [2]"
        End If
        strSQL = _
            " Select ID,ID as ��ĿID,����,����,����," & IIf(gint���� = 0, "����", "����� as ����") & ",˵��" & _
            " From ��������Ŀ¼ Where ���='D' And (" & strSQL & ")" & _
            IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
            " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by ����"
    End If
    GetXYSQL = strSQL
End Function

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, strInput As String
    Dim vPoint As POINTAPI, int������� As Integer
    
    With vsDiagXY
        If Col = col��� Then
            If .EditText = "" Then
                If .TextMatrix(Row, col����) <> "" Then
                    .EditText = .Cell(flexcpData, Row, Col)
                End If
                If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
            ElseIf .TextMatrix(Row, col����) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '�жϼ���ǰ׺��������Ƿ������������ϱ���
                strInput = UCase(.EditText)
                strSQL = GetXYSQL(strInput, strSQL, str�Ա�)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, _
                         str�Ա�, gint���� + 1)
                If rsTmp.RecordCount <> 1 Then
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, col���) = .EditText
                Else
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                End If
                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                .Tag = ""
                mblnChange = True
            Else
                int������� = Val(Mid(gstr�������, 1, 1))
                If int������� = 0 Then int������� = 1
                
                strInput = UCase(.EditText)
                strSQL = GetXYSQL(strInput, strSQL, str�Ա�)
                If int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", str�Ա�, gint���� + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                    Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).Value, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", str�Ա�, gint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                            If mblnReturn Then Call DiagEnterNextCell(vsDiagXY)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ʱ�� Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                Else
                    MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��"
                    Cancel = True
                End If
            End If
            If .EditText <> .TextMatrix(Row, Col) Then mblnChange = True: vsDiagXY.Tag = ""
            If Row = 1 Then
                If .EditText <> "" Then
                    '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
                    txt��������.BackColor = vbButtonFace
                    txt��������.Enabled = False
                    txt����ʱ��.BackColor = vbButtonFace
                    txt����ʱ��.Enabled = False
                Else
                    If vsDiagZY.TextMatrix(0, col����ʱ��) = "" Then
                        txt��������.BackColor = vbWindowBackground
                        txt��������.Enabled = True
                        txt����ʱ��.BackColor = vbWindowBackground
                        txt����ʱ��.Enabled = True
                        txt��������.Text = "____-__-__"
                        txt����ʱ��.Text = "__:__"
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
        If Col = col��� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
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
            If NewCol = col��� Then
                .ComboList = "..."
            ElseIf NewCol = col��ҽ֤�� Then
                If .TextMatrix(NewRow, col���) = "" Then
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
    If Col = col���� Then Cancel = True
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str�Ա� As String
    Dim blnCancle As Boolean
    
    With vsDiagZY
        If Col = col��� Then
            If optInput(0).Value Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", mlng����ID, , True, False)
            Else
                'B-��ҽ��������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", mlng����ID, cboEdit(cbo�Ա�).Text, True)
            End If
            If rsTmp Is Nothing Then
                If optInput(0).Value Then
                    MsgBox "û�м���������ݿ���ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call ZYSetDiagInput(Row, rsTmp)
                Call DiagEnterNextCell(vsDiagZY)
            End If
        ElseIf Col = col��ҽ֤�� Then
            If optInput(0).Value Then
                '���������:�Ȳ��Ƿ��ж�Ӧ
                If Not Set��ҽ֤��(Row, Val(.TextMatrix(Row, col���ID))) Then
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, cboEdit(cbo�Ա�).Text, True)
                Else
                    Exit Sub
                End If
            Else
                'Z-��ҽ��������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, cboEdit(cbo�Ա�).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call Set��ҽ֤��(Row, 0, rsTmp)
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
            If .Col = col��� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col���) <> "" Then
                If .TextMatrix(.Row, colҽ��ID) = "" Then
                    If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call CreatePlugInOK(p����ҽ��վ)
                        'ɾ����/��Ҫ��Ϻ������ҽӿ�
                        If Not gobjPlugIn Is Nothing Then
                            On Error Resume Next
                            Call gobjPlugIn.DiagnosisDeleted(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(.Row, col���ID)), .TextMatrix(.Row, col���))
                            Call zlPlugInErrH(err, "DiagnosisDeleted")
                            err.Clear: On Error GoTo 0
                        End If
                        .RemoveItem .Row
                        mblnChange = True
                        .Tag = ""
                    End If
                Else
                    MsgBox "����϶�Ӧ�Ĵ����ѷ��ͣ�����ɾ����", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagZY)
        ElseIf KeyAscii = 32 And (.Col = col����) Then
            If DiagCellEditable(vsDiagZY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col���� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
            End If
        Else
            If .Col = col��� Or .Col = col��ҽ֤�� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
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
    ElseIf Col = col���� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Function DiagCellEditable(objGrid As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With objGrid
        '�����в��ɱ༭
        If .ColHidden(lngCol) Then Exit Function
        
        If .TextMatrix(lngRow, colҽ��ID) <> "" Then
            If lngCol = col��� Then
                Exit Function
            End If
        End If
        '�������������
        If .TextMatrix(lngRow, col���) = "" Then
            If lngCol = col���� Or lngCol = col����ʱ�� Then
                Exit Function
            End If
        End If
        If lngCol = col���� Then
            Exit Function
        End If
        '���������������֤��
        If lngCol = col��ҽ֤�� Then
            If .TextMatrix(lngRow, col���) = "" Then Exit Function
        End If
    End With
    DiagCellEditable = True
End Function

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAller
        If .Col = AC_������Ӧ Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AC_����ҩ��
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
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col���) To col����
                If DiagCellEditable(objGrid, i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col���� Then Exit For
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
'���ܣ��������ҩ�������
'������strTYTInput=̫Ԫͨ������ҩ�ӿڷ��ص��ַ���
    Dim strSQL As String, curDate As Date
    Dim arrTmp As Variant
    Dim strAllerOld As String, strAllerNew As String
    
    With vsAller
        
        strAllerOld = .Cell(flexcpData, lngRow, AC_����ҩ��) & ";" & .TextMatrix(lngRow, AC_����Դ����)
        
        If mblnUseTYT Then
            arrTmp = Split(strTYTInput, ";")
            
            If UBound(arrTmp) < 1 Then Exit Sub
            If strAllerOld <> strTYTInput Or Val(.RowData(lngRow) & "") <> 0 Then
                .TextMatrix(lngRow, AC_����ҩ��) = arrTmp(1)
                .TextMatrix(lngRow, AC_����Դ����) = arrTmp(0)
                .RowData(lngRow) = 0
            End If
        Else
            
            If Not rsInput Is Nothing Then
                .RowData(lngRow) = CLng(rsInput!ID)
                .TextMatrix(lngRow, AC_����ҩ��) = Nvl(rsInput!����)
            Else
                .RowData(lngRow) = 0
                .TextMatrix(lngRow, AC_����ҩ��) = .EditText
            End If
            
            strAllerNew = .TextMatrix(lngRow, AC_����ҩ��) & ";" & .TextMatrix(lngRow, AC_����Դ����)
            
            If strAllerOld <> strAllerNew Or Val(.RowData(lngRow) & "") <> 0 Then
                .TextMatrix(lngRow, AC_����Դ����) = ""
            End If
        End If
        
        .Cell(flexcpData, lngRow, AC_����ҩ��) = .TextMatrix(lngRow, AC_����ҩ��)
        If .Cell(flexcpData, lngRow, AC_����ʱ��) = "" Then
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, AC_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, AC_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        
        .Tag = ""
        mblnChange = True
    End With
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngԭ���id As Long '0 ��ʾ����ӵ���ϣ� ��Ϊ0��ʾ�޸���ϣ�lngԭ���id ��ֵ�����޸�ǰ�� ���ID�򼲲�ID
    
    With vsDiagXY
        If Not rsInput Is Nothing Then
            '����Ƿ������޸�
            If .TextMatrix(.Row, colҽ��ID) <> "" Then
                MsgBox "����϶�Ӧ�Ĵ����ѷ��ͣ������޸ġ�", vbInformation, Me.Caption
                Exit Sub
            End If
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                    .TextMatrix(lngRow, col����) = "��ҽ"
                    lngԭ���id = 0
                Else
                    lngԭ���id = Val(.TextMatrix(lngRow, col���ID))
                End If
                
                .TextMatrix(lngRow, col���) = Nvl(rsInput!����)
                .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
                .TextMatrix(lngRow, col����) = IIf(Not IsNull(rsInput!����), rsInput!����, "")
                '�������ȷ������,����ݼ���ȷ�����
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col����ID) = ""
                    strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col���ID) = ""
                    strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If optInput(0).Value Then
                        .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                Call CreatePlugInOK(p����ҽ��վ)
                '������/��Ҫ��Ϻ������ҽӿ�
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If lngRow = .FixedRows Then
                        Call gobjPlugIn.DiagnosisEnter(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(rsInput!��ĿID), .TextMatrix(lngRow, col���), lngԭ���id)
                        Call zlPlugInErrH(err, "DiagnosisEnter")
                    Else
                        Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(rsInput!��ĿID), .TextMatrix(lngRow, col���), lngԭ���id)
                        Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End If
                    err.Clear: On Error GoTo errH
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col���) = .EditText
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
        End If
        
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col����) = "��ҽ"
        End If
        .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
        mblnChange = True
        .Tag = ""
    End With
    
    If optState(opt����).Value = False Then
        If PatiReSeeDoctor Then
            If MsgBox("���˾�����ҡ�ҽ����������ϴ���ͬ��Ҫ���Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                optState(opt����).Value = True
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function PatiReSeeDoctor() As Boolean
'���ܣ��жϲ��˱����Ƿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ҽ�����������ϴ���ͬ��û��ת������
    strSQL1 = "Select ����ID,ִ���� as ҽ��,ִ�в���ID as ����ID From ���˹Һż�¼ Where ID=[2] And ת�����ID Is Null And �������ID Is Null"
    
    strSQL2 = "Select Max(ID) as ID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
            " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
    strSQL2 = "Select ����ID,ִ���� as ҽ��,ִ�в���ID as ����ID From ���˹Һż�¼ Where ID=(" & strSQL2 & ") And ת�����ID Is Null And �������ID Is Null"
    
    strSQL = "Select 1 From (" & strSQL1 & ") A,(" & strSQL2 & ") B Where A.����ID=B.����ID And A.ҽ��=B.ҽ�� And A.����ID=B.����ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng����ID, mlng�Һ�ID)
    If rsTmp.EOF Then Exit Function
    
    '��Ҫ������ϴ���ͬ
    With vsDiagXY
        If .TextMatrix(.FixedRows, col���) <> "" Then
            strSQL = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                    " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
            strSQL = "Select 1 From ������ϼ�¼" & _
                " Where ����ID=[1] And ��ҳID=(" & strSQL & ")" & _
                " And �������=1 And ��¼��Դ IN(1,3) And ��ϴ���=1" & _
                " And (����ID=[3] And ����ID<>0 Or ���ID=[4] And ���ID<>0 Or �������=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng����ID, mlng�Һ�ID, _
                Val(.TextMatrix(.FixedRows, col����ID)), Val(.TextMatrix(.FixedRows, col���ID)), .TextMatrix(.FixedRows, col���))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    With vsDiagZY
        If .TextMatrix(.FixedRows, col���) <> "" Then
            strSQL = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                   " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
            strSQL = "Select 1 From ������ϼ�¼" & _
                " Where ����ID=[1] And ��ҳID=(" & strSQL & ")" & _
                " And �������=11 And ��¼��Դ IN(1,3) And ��ϴ���=1" & _
                " And (����ID=[3] And ����ID<>0 Or ���ID=[4] And ���ID<>0 Or �������=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng����ID, mlng�Һ�ID, _
                Val(.TextMatrix(.FixedRows, col����ID)), Val(.TextMatrix(.FixedRows, col���ID)), .TextMatrix(.FixedRows, col���))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, str���� As String
    Dim i As Long
    Dim strTmp As String
    Dim lngԭ���id As Long '0 ��ʾ����ӵ���ϣ� ��Ϊ0��ʾ�޸���ϣ�lngԭ���id ��ֵ�����޸�ǰ�� ���ID�򼲲�ID
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            '����Ƿ������޸�
            If .TextMatrix(.Row, colҽ��ID) <> "" Then
                MsgBox "����϶�Ӧ�Ĵ����ѷ��ͣ������޸ġ�", vbInformation, Me.Caption
                Exit Sub
            End If
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                    .TextMatrix(lngRow, col����) = "��ҽ"
                    lngԭ���id = 0
                Else
                    lngԭ���id = Val(.TextMatrix(lngRow, col���ID))
                End If
                
                If Not IsNull(rsInput!����) Then
                    str���� = rsInput!����
                End If
                
                If InStr(.TextMatrix(lngRow, col���), "(") > 0 And InStr(.TextMatrix(lngRow, col���), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, col���), InStrRev(.TextMatrix(lngRow, col���), "("))
                End If
                .TextMatrix(lngRow, col���) = Nvl(rsInput!����) & strTmp
                .TextMatrix(lngRow, col����) = IIf(Not IsNull(rsInput!����), rsInput!����, "")
                
                '�������ȷ������,����ݼ���ȷ�����
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col����ID) = ""
                    strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col���ID) = ""
                    strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If optInput(0).Value Then
                        .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                '��ҽ���ݼ�����ϲο�ȡ֤��
                Call Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, col���ID)))
                                 
                Call CreatePlugInOK(p����ҽ��վ)
                '������/��Ҫ��Ϻ������ҽӿ�
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If lngRow = .FixedRows Then
                        Call gobjPlugIn.DiagnosisEnter(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(rsInput!��ĿID), .TextMatrix(lngRow, col���), lngԭ���id)
                        Call zlPlugInErrH(err, "DiagnosisEnter")
                    Else
                        Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(rsInput!��ĿID), .TextMatrix(lngRow, col���), lngԭ���id)
                        Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End If
                    err.Clear: On Error GoTo errH
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col���) = .EditText
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
            .TextMatrix(lngRow, col֤��ID) = ""
        End If
        
        '����ǳ�Ժ���,ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col����) = "��ҽ"
        End If
        .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
        mblnChange = True
        .Tag = ""
    End With
    
    If optState(opt����).Value = False Then
        If PatiReSeeDoctor Then
            If MsgBox("���˾�����ҡ�ҽ����������ϴ���ͬ��Ҫ���Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                optState(opt����).Value = True
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Set��ҽ֤��(ByVal lngRow As Long, ByVal lng���ID As Long, Optional ByVal rsInput As Recordset) As Boolean
'���ܣ���ҽ���ݼ�����ϲο�ȡ֤��
'������rsInput-�����Ϊ�գ������ָ������ҩ֤���¼��
'���أ��Ƿ��ж�Ӧ��ϵ
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    With vsDiagZY
        'ȥ�����е�֤��
        If InStr(.TextMatrix(lngRow, col���), "(") > 0 And InStr(.TextMatrix(lngRow, col���), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, col���), 1, InStrRev(.TextMatrix(lngRow, col���), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, col���)
        End If
        If rsInput Is Nothing Then
            If lng���ID <> 0 Then
                strSQL = "Select Distinct a.֤����� as ID,a.֤��ID,a.֤������,b.���� as ֤�����" & _
                    " From ������ϲο� A,��������Ŀ¼ B" & _
                    " Where a.֤��ID=b.ID(+) And a.���ID=[1] And a.֤������ is Not NULL" & _
                    " Order by a.֤�����"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng���ID)
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, col֤��ID) = Nvl(rsTmp!֤��id)
                    If Not IsNull(rsTmp!֤������) Then
                        .TextMatrix(lngRow, col���) = strTmp
                        .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
                        .TextMatrix(lngRow, col��ҽ֤��) = Nvl(rsTmp!֤������)
                        .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
                        If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
                        mblnChange = True
                        .Tag = ""
                    End If
                    Set��ҽ֤�� = True
                Else
                    If blnCancel Then
                        Set��ҽ֤�� = True
                        If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col��ҽ֤��)
                    Else
                        Set��ҽ֤�� = False
                    End If
                End If
            Else
                Set��ҽ֤�� = False
            End If
        Else
            .TextMatrix(lngRow, col֤��ID) = Nvl(rsInput!��ĿID)
            .TextMatrix(lngRow, col���) = strTmp
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            .TextMatrix(lngRow, col��ҽ֤��) = Nvl(rsInput!����)
            .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
            .Tag = ""
            mblnChange = True
        End If
    End With
End Function

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim str�Ա� As String, int������� As Integer
    
    With vsDiagZY
        If Col = col��� Or Col = col��ҽ֤�� Then
            If .EditText = "" Then
                If .TextMatrix(Row, col����) <> "" And Col = col��� Then
                    .EditText = .Cell(flexcpData, Row, Col)
                Else
                    '��ҽ֢���������������
                    If Col = col��ҽ֤�� Then
                        .Cell(flexcpData, Row, Col) = ""
                    End If
                End If
                If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
            ElseIf Col = col��� And .TextMatrix(Row, col����) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                strSQL = GetZYSQL(strInput, strSQL, str�Ա�)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str�Ա�, gint���� + 1)
                If rsTmp.RecordCount = 1 Then
                    Call ZYSetDiagInput(Row, rsTmp):
                    .EditText = .Text
                Else
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, col���) = .EditText
                End If
                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                .Tag = ""
                mblnChange = True
            Else
                int������� = Val(Mid(gstr�������, 1, 1))
                If int������� = 0 Then int������� = 1
                
                strInput = UCase(.EditText)
                strSQL = GetZYSQL(strInput, strSQL, str�Ա�, IIf(Col = col���, "B", "Z"))
                If Col = col��� Then
                    If int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                        On Error GoTo errH
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", str�Ա�, gint���� + 1)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                        End If
                        Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                        If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
                    Else
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).Value, "�������", "��������"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", str�Ա�, gint���� + 1)
                        If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                            Cancel = True
                        Else
                            '���������뷽ʽ
                            If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                                MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                                Cancel = True
                            Else
                                Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                                If mblnReturn Then Call DiagEnterNextCell(vsDiagZY)
                            End If
                        End If
                    End If
                ElseIf Col = col��ҽ֤�� Then
                    If optInput(0).Value Then
                        '���������:�Ȳ��Ƿ��ж�Ӧ
                        If Set��ҽ֤��(Row, Val(.TextMatrix(Row, col���ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", str�Ա�, gint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call Set��ҽ֤��(Row, 0, rsTmp)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ʱ�� Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                Else
                    MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��"
                    Cancel = True
                End If
            End If
            If .EditText <> .TextMatrix(Row, Col) Then mblnChange = True: vsDiagZY.Tag = ""
            If Row = 0 Then
                If .EditText <> "" Then
                    '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
                    txt��������.BackColor = vbButtonFace
                    txt��������.Enabled = False
                    txt����ʱ��.BackColor = vbButtonFace
                    txt����ʱ��.Enabled = False
                Else
                    If vsDiagXY.TextMatrix(1, col����ʱ��) = "" Then
                        txt��������.BackColor = vbWindowBackground
                        txt��������.Enabled = True
                        txt����ʱ��.BackColor = vbWindowBackground
                        txt����ʱ��.Enabled = True
                        txt��������.Text = "____-__-__"
                        txt����ʱ��.Text = "__:__"
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
'���ܣ���ָ������Դ�е�����װ��ָ��������һ������ComboBox
'������strSQL=����"ID,����,����,ȱʡ��־"�ֶ�
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, j As Long
    
    '���ԭ������
    For i = 0 To UBound(arrCboIdx)
        cboEdit(arrCboIdx(i)).Clear
    Next
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    'װ������
    For i = 1 To rsTmp.RecordCount
        For j = 0 To UBound(arrCboIdx)
            If IsNull(rsTmp!����) Then
                cboEdit(arrCboIdx(j)).AddItem rsTmp!����
            Else
                cboEdit(arrCboIdx(j)).AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����
            End If
            cboEdit(arrCboIdx(j)).ItemData(cboEdit(arrCboIdx(j)).NewIndex) = Nvl(rsTmp!ID, 0)
            If Nvl(rsTmp!ȱʡ��־, 0) = 1 Then
                Call zlControl.CboSetIndex(cboEdit(arrCboIdx(j)).hwnd, cboEdit(arrCboIdx(j)).NewIndex)
            End If
        Next
        rsTmp.MoveNext
    Next
    '��ȱʡʱ,Ϊδѡ��
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
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
