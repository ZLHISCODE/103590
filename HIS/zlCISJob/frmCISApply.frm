VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISApply 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "���Ӳ�����������"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20475
   Icon            =   "frmCISApply.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   20475
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   15015
      TabIndex        =   10
      Top             =   5040
      Width           =   15015
      Begin VB.Frame picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ȩ��Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   12840
         TabIndex        =   17
         Top             =   840
         Width           =   5175
         Begin VSFlex8Ctl.VSFlexGrid vsInfo 
            Height          =   6915
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4750
            _cx             =   1989550266
            _cy             =   1989554085
            Appearance      =   0
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
            MouseIcon       =   "frmCISApply.frx":6852
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   10000
            ColWidthMin     =   4650
            ColWidthMax     =   10000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCISApply.frx":712C
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.Frame fraFillter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ѯ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   17055
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   277
            Width           =   1365
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ѯ(&F)"
            Height          =   375
            Left            =   13080
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ѳ���"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   12150
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Ѿܾ�"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   10920
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   9675
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   8445
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   7200
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   3
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2595
            TabIndex        =   1
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   217448451
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4890
            TabIndex        =   2
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   217448451
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "frmCISApply.frx":71C7
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   4
            Left            =   11850
            Picture         =   "frmCISApply.frx":7751
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   10605
            Picture         =   "frmCISApply.frx":DFA3
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   9375
            Picture         =   "frmCISApply.frx":147F5
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   8130
            Picture         =   "frmCISApply.frx":1B047
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   6900
            Picture         =   "frmCISApply.frx":21899
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            X1              =   4605
            X2              =   4805
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   330
            Width           =   855
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   7275
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   12645
         _cx             =   1989564192
         _cy             =   1989554720
         Appearance      =   0
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
         MouseIcon       =   "frmCISApply.frx":280EB
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISApply.frx":289C5
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   14
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   240
      Top             =   240
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
            Picture         =   "frmCISApply.frx":28A60
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":2F2C2
            Key             =   "boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":35B24
            Key             =   "����ʱ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":360BE
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":36658
            Key             =   "����ҽ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":36BF2
            Key             =   "���ʲ���"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":3718C
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApply.frx":372E6
            Key             =   "unCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMec 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   6480
      ScaleHeight     =   4575
      ScaleWidth      =   9615
      TabIndex        =   11
      Top             =   1200
      Width           =   9615
      Begin VB.PictureBox picVLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9300
         Left            =   5920
         MousePointer    =   9  'Size W E
         ScaleHeight     =   9300
         ScaleWidth      =   30
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   120
         Width           =   30
      End
      Begin VB.Frame fraPatiFilter 
         Appearance      =   0  'Flat
         Caption         =   "���˲���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   9495
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   5895
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   5535
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   5655
            _Version        =   589884
            _ExtentX        =   9975
            _ExtentY        =   9763
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1800
            TabIndex        =   35
            Top             =   1250
            Width           =   2535
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "����������"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   3
            Left            =   4320
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "����ϲ���"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   2
            Left            =   2960
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "����ʶ����"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   1
            Left            =   1600
            TabIndex        =   30
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "�����Ҳ���"
            ForeColor       =   &H000040C0&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1200
            Width           =   2565
         End
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   600
            ScaleHeight     =   240
            ScaleWidth      =   1140
            TabIndex        =   21
            Top             =   1250
            Width           =   1170
            Begin VB.ComboBox cboFind 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   300
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   -30
               Width           =   1215
            End
         End
         Begin VB.ComboBox cboSelectTime 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   750
            Width           =   2565
         End
         Begin VB.CommandButton cmdPatiFind 
            Caption         =   "����(&F)"
            Height          =   375
            Left            =   4440
            TabIndex        =   37
            Top             =   1150
            Width           =   1215
         End
         Begin VB.Label lblDept 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����˿���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   720
            TabIndex        =   23
            Top             =   1290
            Width           =   900
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "ʱ�䷶Χ(&T)"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   720
            TabIndex        =   22
            Top             =   840
            Width           =   990
         End
      End
      Begin VB.PictureBox picMecInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   6120
         ScaleHeight     =   3735
         ScaleWidth      =   3855
         TabIndex        =   19
         Top             =   240
         Width           =   3855
         Begin VB.PictureBox picShow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   8175
            Left            =   0
            ScaleHeight     =   8175
            ScaleWidth      =   11775
            TabIndex        =   25
            Top             =   480
            Width           =   11775
            Begin VB.PictureBox PicNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   4695
               Left            =   1320
               ScaleHeight     =   4665
               ScaleWidth      =   7665
               TabIndex        =   27
               Top             =   2400
               Width           =   7695
               Begin VB.PictureBox picNoUse 
                  BorderStyle     =   0  'None
                  Height          =   2535
                  Left            =   0
                  Picture         =   "frmCISApply.frx":37440
                  ScaleHeight     =   2535
                  ScaleWidth      =   7815
                  TabIndex        =   28
                  Top             =   960
                  Width           =   7815
               End
            End
            Begin XtremeSuiteControls.TabControl tbcMec 
               Height          =   6420
               Left            =   120
               TabIndex        =   26
               Top             =   480
               Width           =   8610
               _Version        =   589884
               _ExtentX        =   15187
               _ExtentY        =   11324
               _StockProps     =   64
            End
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5580
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   8130
      _Version        =   589884
      _ExtentX        =   14340
      _ExtentY        =   9842
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   10560
      Width           =   20475
      _ExtentX        =   36116
      _ExtentY        =   635
      SimpleText      =   $"frmCISApply.frx":3CD45
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCISApply.frx":3CD8C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   31036
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   600
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCISApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdtBegin As Date, mdtEnd As Date
Private mintPreTime As Integer
Private mobjArchiveView As frmArchiveView
Private mrsTmp As ADODB.Recordset     '��ǰҽ����Ϣ�������ݼ�
Private mstrDeptIds As String      '��ǰҽ������IDs

Private Enum colList
    COL_����ID = 1
    COL_�������� = 2
    COL_����ʱ�� = 3
    COL_����ʱ�� = 4
    COL_������ = 5
    COL_������ = 6

    COL_����ʱ�� = 7
    COL_������ʲ��� = 8
    COL_���ʿ�ʼʱ�� = 9
    COL_���ʽ���ʱ�� = 10
    COL_����ԭ�� = 11
    COL_����״̬ = 12
End Enum

Private Enum RowInfo
    Row_���ʲ��˱��� = 0
    Row_���ʲ��� = 1
    Row_����ʱ�ޱ��� = 3
    Row_����ʱ�� = 4
    Row_�������ݱ��� = 6
    Row_�������� = 7
End Enum


Private Enum colPati
    col_ѡ�� = 0
    col_����Id = 1
    col_���� = 2
    col_�Ա� = 3
    col_���� = 4
    COL_��ʶ�� = 5
    col_���� = 6
    COL_��ǰ״̬ = 7
    col_����ID = 8
    col_����ID = 9
End Enum

Private Enum CmdIndex
    Cmd_���п��� = 9999991
    Cmd_������� = 9999992
    Cmd_סԺ���� = 9999993
End Enum


Private Sub cboSelectTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    If cboSelectTime.ListIndex = mintPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtBegin, mdtEnd, cboSelectTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call Cbo.SetIndex(cboSelectTime.hwnd, mintPreTime)
            Exit Sub
        End If
    Else
        mdtEnd = Format(datCurr, "yyyy-MM-dd 23:59:59")
        mdtBegin = datCurr - intDateCount
    End If
    If mdtBegin = CDate(0) Or mdtEnd = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "��Χ��" & Format(mdtBegin, "yyyy-MM-dd") & " �� " & Format(mdtEnd, "yyyy-MM-dd")
    End If
    mintPreTime = cboSelectTime.ListIndex
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate

    dtpTime(0).MaxDate = curDate + 1
    dtpTime(1).MaxDate = curDate + 1

    
    Select Case cboTime.ListIndex
    Case 0 '����
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '�������
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '�������
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '���һ��
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '���һ��
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 'ָ  ��
        If Me.Visible Then
            dtpTime(0).SetFocus
        End If
    End Select
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngApplyID As Long
    
    Select Case Control.ID
        Case conMenu_Edit_ApplyAdd
            If frmCISApplyEdit.ShowEdit(Me, 0, lngApplyID, IIf(tbcSub.Selected.Tag = "���ʵ��Ӳ���", GetPatiRs, Nothing)) Then
                Call LoadList(lngApplyID)
            End If
        Case conMenu_Edit_ApplyEdit
            If Val(vsList.TextMatrix(vsList.Row, COL_����ID)) = 0 Then Exit Sub
            lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_����ID))
            If frmCISApplyEdit.ShowEdit(Me, 1, lngApplyID) Then
                Call LoadList(lngApplyID)
            End If
        Case conMenu_Edit_ApplyBack
            If Val(vsList.TextMatrix(vsList.Row, COL_����ID)) = 0 And vsList.TextMatrix(vsList.Row, COL_����״̬) <> "������" Then Exit Sub
            lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_����ID))
            If ApplyBack(lngApplyID) Then
                Call LoadList(lngApplyID)
            End If
        Case conMenu_View_Refresh
            If tbcSub.Selected.Tag = "�����¼" Then
                Call LoadList
            Else
                Call LoadPati
            End If
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
            
        Case Cmd_���п���, Cmd_�������, Cmd_סԺ����
            lblDept.Tag = Control.Parameter
            lblDept.Caption = Decode(lblDept.Tag, "", "�����˿���", "����", "���������", "סԺ", "��סԺ����")
            Call LoadDept
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Function ApplyBack(lngApplyID As Long) As Boolean
    Dim strSql As String
    Dim curDate As Date
    Dim blnTran As Boolean
    
    On Error GoTo errH
    If MsgBox("ȷ��Ҫ����ѡ�е���Ȩ�����¼��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    curDate = zlDatabase.Currentdate
    strSql = "Zl_���Ӳ�����������_����״̬(" & lngApplyID & ",4,'" & UserInfo.���� & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    ApplyBack = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lblDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, True)
End Sub



Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_ApplyEdit
        If vsList.Row <= 0 And tbcSub.Selected.Tag = "�����¼" Then Control.Enabled = False: Exit Sub
        Control.Visible = tbcSub.Selected.Tag = "�����¼" And vsList.TextMatrix(vsList.Row, COL_����״̬) = "������"
    Case conMenu_Edit_ApplyBack
        If vsList.Row <= 0 And tbcSub.Selected.Tag = "�����¼" Then Control.Enabled = False: Exit Sub
        Control.Visible = tbcSub.Selected.Tag = "�����¼" And vsList.TextMatrix(vsList.Row, COL_����״̬) = "������"
    Case Cmd_���п���, Cmd_�������, Cmd_סԺ����
         Control.Checked = Control.Parameter = lblDept.Tag
    End Select
End Sub

Private Sub chkFilter_Click(Index As Integer)
    Dim i As Long
    Dim blnCheck As Boolean
    
    For i = 0 To 4
        If chkFilter(i).Value = 1 Then
            blnCheck = True
            Exit For
        End If
    Next
    If Not blnCheck Then
        MsgBox "������ѡ��һ�ַ������ڹ��ˡ�", vbInformation, gstrSysName
        chkFilter(Index).Value = 1
        Exit Sub
    End If
End Sub

Private Sub cmdFind_Click()
    Call LoadList
End Sub

Public Function GetRs��������(rsTmp As ADODB.Recordset) As Boolean
    Dim str����ids As String
    Dim arrTmp As Variant
    Dim colPati As Collection
    Dim i As Long, j As Long
    Dim str���� As String, colValue As Collection
    
    If rsTmp Is Nothing Then Exit Function
    If rsTmp.EOF Then Exit Function
    
    
    '���ز�����Ϣ
    str����ids = ""
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
             If rsTmp!����ids & "" <> "" Then
                arrTmp = Split(rsTmp!����ids & "", ",")
                For j = LBound(arrTmp) To UBound(arrTmp)
                    If InStr("," & str����ids & ",", "," & Val(arrTmp(j)) & ",") = 0 Then
                       str����ids = str����ids & "," & Val(arrTmp(j))
                    End If
                Next
             End If
             rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    End If

    
    If str����ids <> "" Then
        str����ids = Mid(str����ids, 2)
        Set colPati = PatiSvrGetpatiinfo(1, 0, 1241, 0, 2, "", "", "", "", str����ids)
        
        If Not colPati Is Nothing Then
            Set rsTmp = zlDatabase.CopyNewRec(rsTmp)
            Do While Not rsTmp.EOF
               If rsTmp!����ids & "" <> "" Then
                    arrTmp = Split(rsTmp!����ids & "", ",")
                    str���� = ""
                    For j = LBound(arrTmp) To UBound(arrTmp)
                        If Val(arrTmp(j)) <> 0 Then
                            Set colValue = GetColObj(colPati, "_" & arrTmp(j))
                            If Not colValue Is Nothing Then
                                If GetColVal(colValue, "_pati_name") <> "" Then
                                    str���� = str���� & "," & GetColVal(colValue, "_pati_name")
                                End If
                            End If
                        End If
                    Next
                End If
                
                If str���� <> "" Then
                    str���� = Mid(str����, 2)
                    rsTmp!�������� = str����
                End If
                
                rsTmp.MoveNext
            Loop
            rsTmp.MoveFirst
        End If
    End If
End Function

Private Sub LoadList(Optional lng����id As Long)
    Dim strSql As String
    Dim strFilter As String
    Dim str�ѳ��� As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date
    
    For i = 0 To 3
        If chkFilter(i).Value = 1 Then strFilter = strFilter & "," & i
    Next
    strFilter = Mid(strFilter, 2)
    
    '�����ѳ�����¼
    If chkFilter(4).Value = 0 Then
        str�ѳ��� = " And A.����ʱ�� is null"
    Else
        If strFilter = "" Then
            str�ѳ��� = " And A.����ʱ�� is not null"
        Else
            str�ѳ��� = " Or A.����ʱ�� is not null)"
        End If
    End If
    
    On Error GoTo errH
    If cboTime.ListIndex <> 5 Then
        curDate = zlDatabase.Currentdate
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSql = "Select a.Id, a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.����ԭ��, a.����״̬, a.������, a.����ʱ��,A.����ʱ��,A.������," & vbNewLine & _
                "       f_List2str(Cast(Collect(b.����id || '') As t_Strlist)) As ����ids,null as ��������" & vbNewLine & _
                "From ���Ӳ����������� A, ���Ӳ���������ʲ��� B" & vbNewLine & _
                "Where a.Id = b.����id  And a.����ʱ�� Between [1] And [2] And a.������ = [3]" & vbNewLine & _
                IIf(strFilter <> "", IIf(chkFilter(4).Value = 1, " And (Instr([4], a.����״̬) > 0", " And Instr([4], a.����״̬) > 0"), "") & str�ѳ��� & vbNewLine & _
                "Group By a.Id, a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.����ԭ��, a.����״̬, a.������, a.����ʱ��,A.����ʱ��,A.������" & vbNewLine & _
                "Order by a.����״̬,A.id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtpTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboTime.ListIndex <> 5, dtpTime(1).Value + 1, dtpTime(1).Value), "yyyy-MM-dd hh:mm")), UserInfo.����, strFilter)
    
    Call GetRs��������(rsTmp)
    With vsList
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '������
                .TextMatrix(i, COL_����ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_����ʱ��) = Val(rsTmp!����ʱ�� & "")
                .TextMatrix(i, COL_����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_������) = rsTmp!������ & ""
                '��ʾ��
                .TextMatrix(i, COL_������) = rsTmp!������ & ""
                .TextMatrix(i, COL_����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_������ʲ���) = rsTmp!�������� & ""
                .TextMatrix(i, COL_���ʿ�ʼʱ��) = Format(rsTmp!���ʿ�ʼʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_���ʽ���ʱ��) = Format(rsTmp!���ʽ���ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_����ԭ��) = rsTmp!����ԭ�� & ""
                
                If rsTmp!����ʱ�� & "" <> "" Then
                    .TextMatrix(i, COL_����״̬) = "�ѳ���"
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(4).Picture
                Else
                    .TextMatrix(i, COL_����״̬) = Decode(Val(rsTmp!����״̬ & ""), 0, "������", 1, "������", 2, "������", 3, "�Ѿܾ�")
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(Val(rsTmp!����״̬ & "")).Picture
                End If

                If Val(rsTmp!ID & "") = lng����id Then
                    .Row = i
                End If
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             If tbcSub.Selected.Tag = "�����¼" Then stbThis.Panels(2).Text = "��ǰ���˲��ҵ� " & rsTmp.RecordCount & " ��������Ϣ"
        Else
            .Rows = .FixedRows + 1
            If tbcSub.Selected.Tag = "�����¼" Then stbThis.Panels(2).Text = "��ǰ����û�в��ҵ�������Ϣ"
        End If
        
        If .Row <= 0 Then .Row = .Rows - 1
        If Me.Tag = "1" And .Visible Then .SetFocus
        .WordWrap = True
        '�Զ������и�
        .AutoSize COL_������ʲ���, COL_����ԭ��
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub cmdPatiFind_Click()
    Call LoadPati
End Sub

Private Sub fraPatiFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call SetLBLFace(lblDept, False)
End Sub

Private Sub optFind_Click(Index As Integer)
    Call SetFindCtl
    If optFind(0).Value Then
        If cboDept.Visible Then cboDept.SetFocus
    Else
        If txtFind.Visible Then txtFind.SetFocus
    End If
End Sub

Private Sub SetFindCtl()
    txtFind.Text = ""
    cboDept.Visible = optFind(0).Value
    txtFind.Visible = Not optFind(0).Value
    lblDept.Visible = Not optFind(1).Value
    picTmp(1).Visible = optFind(1).Value

    lblDept.Caption = IIf(optFind(0).Value, "�����˿���", IIf(optFind(2).Value, "�������(&D)", IIf(optFind(3).Value, "��������(&O)", "")))
    lblDept.Tag = ""
End Sub


Private Sub picVLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picMec_Resize
End Sub


Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    stbThis.Panels(2).Text = IIf(tbcSub.Selected.Tag = "���ʵ��Ӳ���", "���ʵ��Ӳ���", "�鿴�����¼")
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Or NewCol < 0 Then Exit Sub
    If vsList.Col >= vsList.FixedCols Then
        vsList.ForeColorSel = vsList.Cell(flexcpForeColor, NewRow, NewCol)
    End If
    With vsInfo
        If Val(vsList.TextMatrix(NewRow, COL_����ID)) <> 0 Then
            '���ʲ���
            .TextMatrix(Row_���ʲ���, 0) = vsList.TextMatrix(NewRow, COL_������ʲ���) & ""
            
            '����ʱ��
            .TextMatrix(Row_����ʱ��, 0) = "�� " & Format(vsList.TextMatrix(NewRow, COL_���ʿ�ʼʱ��), "yyyy-mm-dd hh:mm") & vbCrLf & "�� " & _
                                        Format(vsList.TextMatrix(NewRow, COL_���ʽ���ʱ��), "yyyy-mm-dd hh:mm") & "�ڼ�" & vbCrLf & "���ʲ���" & Decode(Val(vsList.TextMatrix(NewRow, COL_����ʱ��)), 0, "���в�������", 1, "δ�鵵�Ĳ���", "�ѹ鵵�Ĳ���")
                             
            '��������
            .TextMatrix(Row_��������, 0) = GetXmlInfo(NewRow)
        Else
            .TextMatrix(Row_���ʲ���, 0) = ""
            .TextMatrix(Row_����ʱ��, 0) = ""
            .TextMatrix(Row_��������, 0) = ""
        End If
        .WordWrap = True
        '�Զ������и�
        .AutoSize 0
    End With
End Sub


Private Function GetXmlString(objXML As Object, ByVal strNode As String, ByRef strValue As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strValue = ""
    If objXML.GetMultiNodeRecord(strNode, rsTmp) Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!node_value
                rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
    End If
    GetXmlString = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetXmlInfo(lngRow As Long) As String
    '��ȡ�������ݵ�Xml������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String
    Dim strOut As String
    Dim strTmp As String
    
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function
    If Val(vsList.TextMatrix(lngRow, COL_����ID)) = 0 Then Exit Function
    
    '��ȡ����
    If vsList.TextMatrix(lngRow, COL_��������) <> "" Then GetXmlInfo = vsList.TextMatrix(lngRow, COL_��������): Exit Function
    
    strXML = Sys.ReadXML("���Ӳ�����������", "��������", "ID=[1]", strErr, Val(vsList.TextMatrix(lngRow, COL_����ID)))
    If Err.Number = 0 And strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        Exit Function
    End If
    
    If objXML.OpenXMLDocument(strXML) = False Then Exit Function

    '��������
    strValue = "": Call objXML.GetSingleNodeValue("all_files", strValue, xsNumber)
    If Val(strValue) = 1 Then
        strOut = "�����Ʒ�����������"
    Else
        '������ҳ��ҽ�����ٴ�·��
        strValue = "": Call objXML.GetSingleNodeValue("medical_record", strValue, xsNumber): If Val(strValue) = 1 Then strOut = "������ҳ��" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("advice", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "����ҽ����" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("cispath", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "�ٴ�·����" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("patipeis", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "��챨�桢" & vbCrLf & vbCrLf
        
        '�����¼
        strValue = "": Call objXML.GetSingleNodeValue("nursing_record", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/nursing_all", strValue, xsNumber)
            If Val(strValue) = 1 Then
                strOut = strOut & "�����¼(���л����¼)" & vbCrLf & vbCrLf
            Else
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/thermometer", strValue, xsNumber): If Val(strValue) = 1 Then strTmp = "���µ���"
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/record_file", strValue, xsNumber)
                If Val(strValue) = 1 Then
                    Call GetXmlString(objXML, "nursing_info/file_name", strValue)
                    strValue = Replace(strValue, ",", "��")
                    strTmp = strTmp & strValue
                Else
                    strTmp = Replace(strTmp, "��", "")
                End If
                strOut = strOut & "�����¼" & vbCrLf & "(��¼��Χ��" & strTmp & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '��鱨��
        strValue = "": Call objXML.GetSingleNodeValue("pacs_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("pacs_info/pacs_type", strValue, xsNumber)
            'pacs_type =0���м�鱨�� =1ָ�����͵ļ�鱨��
            If Val(strValue) = 0 Then
                strOut = strOut & "��鱨��(���м�鱨��)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "pacs_info/pacs_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "��鱨��" & vbCrLf & "(���ͷ�Χ��" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '���鱨��
        strValue = "": Call objXML.GetSingleNodeValue("lis_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("lis_info/lis_type", strValue, xsNumber)
            'lis_type =0 ���м��鱨�� =1ָ�����͵ļ��鱨��
            If Val(strValue) = 0 Then
                strOut = strOut & "���鱨��(���м��鱨��)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "lis_info/lis_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "���鱨��" & vbCrLf & "(���ͷ�Χ��" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '���Ӳ���
        strValue = "": Call objXML.GetSingleNodeValue("emr", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("emr_info/emr_type", strValue, xsNumber)
            'emr_type =0 ���е��Ӳ���  =1ָ�����͵ĵ��Ӳ���  =1ָ������ĵ��Ӳ���
            If Val(strValue) = 0 Then
                strOut = strOut & "���Ӳ���(���е��Ӳ���)" & vbCrLf & vbCrLf
            ElseIf Val(strValue) = 1 Then
                Call GetXmlString(objXML, "emr_info/standard_class/class_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "���Ӳ���" & vbCrLf & "(�������ͷ�Χ��" & strValue & ")" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "emr_info/antetype_class/class_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "���Ӳ���" & vbCrLf & "(������Χ��" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
    End If
    
    If Right(strOut, 5) = "��" & vbCrLf & vbCrLf Then strOut = Left(strOut, Len(strOut) - 5)
    
    '������������
    vsList.TextMatrix(lngRow, COL_��������) = strOut
    GetXmlInfo = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Load()
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    Call MainDefCommandBar
    
    '��ʼ���϶���λ
    Me.picVLine.Left = 5895
    
    'tabControl
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
        .InsertItem(0, "�����¼", picApply.hwnd, 0).Tag = "�����¼"
        .InsertItem(1, "���ʵ��Ӳ���", picMec.hwnd, 0).Tag = "���ʵ��Ӳ���"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call InitListTable
    
    '��ʼ��������
    With vsInfo
        '���ʲ���
        .TextMatrix(Row_���ʲ��˱���, 0) = "���ʲ��ˣ�"
        .Cell(flexcpForeColor, Row_���ʲ��˱���, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_���ʲ��˱���, 0) = img16.ListImages("���ʲ���").Picture
        .Cell(flexcpFontBold, Row_���ʲ��˱���, 0) = True
        
        '����ʱ��
        .TextMatrix(Row_����ʱ�ޱ���, 0) = "����ʱ�ޣ�"
        .Cell(flexcpForeColor, Row_����ʱ�ޱ���, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_����ʱ�ޱ���, 0) = img16.ListImages("����ʱ��").Picture
        .Cell(flexcpFontBold, Row_����ʱ�ޱ���, 0) = True

        '��������
        .TextMatrix(Row_�������ݱ���, 0) = "�������ݣ�"
        .Cell(flexcpForeColor, Row_�������ݱ���, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_�������ݱ���, 0) = img16.ListImages("��������").Picture
        .Cell(flexcpFontBold, Row_�������ݱ���, 0) = True

        .WordWrap = True
        '�Զ������и�
        .AutoSize 0
    End With
    
    '---cboTime
    cboTime.AddItem "��    ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ  ��]"
    cboTime.ListIndex = 3
    
    Call InitSelectTime
    
    Call LoadDept
    
    Call InitReportColumn
    
    Call SetFindCtl
    
    'ִ�н�������˵���ʼ��
    cboFind.Clear
    cboFind.AddItem "����"
    cboFind.AddItem "���֤��"
    cboFind.AddItem "�����"
    cboFind.AddItem "סԺ��"
    cboFind.AddItem "����ID"
    cboFind.ListIndex = 0
    
    Call GetFrom
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Call LoadList
    Me.Tag = "1"
End Sub
'


Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "��������(&A)")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "��������(&E)")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyBack, "��������(&Q)")
            objControl.IconId = 5019
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "��������")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "��������")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyBack, "��������")
            objControl.IconId = 5019
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call cbsMain_Resize
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = Me.Height - stbThis.Height - 1500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mobjArchiveView
    Set mobjArchiveView = Nothing
    Set mrsTmp = Nothing
    mstrDeptIds = ""
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    mintPreTime = -1
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtBegin = datCurr
    mdtEnd = mdtBegin - 7
    
    cboSelectTime.Clear '��Ժ
    With cboSelectTime
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 3
    mintPreTime = cboSelectTime.ListIndex
End Sub

Private Sub picApply_Resize()
    On Error Resume Next
    '�̶���ϸ��Ϣ4000����
    picInfo.Width = 5000

    fraFillter.Top = 100: fraFillter.Left = 30
    fraFillter.Width = picApply.Width - 60
    
    vsList.Top = fraFillter.Top + fraFillter.Height + 150: vsList.Height = picApply.Height - fraFillter.Height - 260

    
    vsList.Left = fraFillter.Left
    vsList.Width = fraFillter.Width - 5000 - 30
    
    picInfo.Top = vsList.Top - 70: picInfo.Left = vsList.Left + vsList.Width + 50
    picInfo.Height = vsList.Height + 70
    vsInfo.Height = picInfo.Height - 300
End Sub


Private Sub picVline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picVLine.Left = Me.picVLine.Left + X
    End If
End Sub

Private Sub picMec_Resize()

    Me.picVLine.Top = 0
    Me.picVLine.Height = Me.picMec.Height
    If Me.picVLine.Left < 100 Then Me.picVLine.Left = 100
    If Me.picVLine.Left > Me.picMec.Width - 100 Then Me.picVLine.Left = Me.picMec.Width - 100


    On Error Resume Next
    fraPatiFilter.Top = 100: fraPatiFilter.Left = 30
    fraPatiFilter.Width = Me.picVLine.Left - Me.fraPatiFilter.Left
    rptPati.Width = fraPatiFilter.Width - 200
    fraPatiFilter.Height = picMec.Height - 100
    
    rptPati.Height = fraPatiFilter.Height - rptPati.Top - 100
    
    picMecInfo.Top = 180: picMecInfo.Height = fraPatiFilter.Height - 80
    picMecInfo.Left = fraPatiFilter.Width + 60
    picMecInfo.Width = picMec.Width - picMecInfo.Left - 60
    
    '����tab��ǩ
    picShow.Top = -360: picShow.Left = 0
    picShow.Width = picMecInfo.Width: picShow.Height = picMecInfo.Height + 360
    
    tbcMec.Top = 0: tbcMec.Left = 0
    tbcMec.Width = picShow.Width: tbcMec.Height = picShow.Height
    
End Sub


Private Sub InitListTable()
'���ܣ���ʼ���б��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "����id;��������;����ʱ��;����ʱ��;������;������;" & _
                "����ʱ��,2000,1;������ʲ���,4000,1;���ʿ�ʼʱ��,2000,1;���ʽ���ʱ��,2000,1;����ԭ��,3800,1;����״̬,1050,4"

    arrHead = Split(strHead, ";")
    With vsList
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDNone
    End With
End Sub

Private Sub LoadDept()
'���ز�ѯ����
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long
    Dim strTmp As String
    
    strSql = "Select B.ID,B.����,B.���� From " & _
            " ���ű� B, ��������˵�� C" & vbNewLine & _
            " Where B.Id = C.����id " & _
            "  And C.�������� = '�ٴ�' " & Decode(lblDept.Tag, "", " And C.������� <> 0 ", "����", " And C.������� in (1,3) ", "סԺ", " And C.������� in (2,3) ") & "  And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"
    On Error GoTo errH
    cboDept.Clear
    '���в���
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID & ""
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
    
    '�������Ա����id�ַ���
    
    strSql = "Select b.Id" & vbNewLine & _
        "From ������Ա A, ���ű� B, ��������˵�� C" & vbNewLine & _
        "Where b.Id = c.����id And a.����id = b.Id And a.��Աid = [1] And c.�������� = '�ٴ�' And c.������� <> 0 And" & vbNewLine & _
        "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
        "Order By b.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            strTmp = strTmp & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        strTmp = strTmp & ","
    End If
    mstrDeptIds = strTmp
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        Set objCol = .Columns.Add(col_ѡ��, "", 18, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False)
        Set objCol = .Columns.Add(col_����, "����", 80, True)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(col_����, "����", 30, True)
        Set objCol = .Columns.Add(COL_��ʶ��, "��ʶ��", 100, True)
        Set objCol = .Columns.Add(col_����, "����", 80, True)
        Set objCol = .Columns.Add(COL_��ǰ״̬, "��ǰ״̬", 150, True)
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False)
        
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
            .HighlightBackColor = &HFFEDCA
            .HighlightForeColor = vbBlack
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub


Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call LoadPati
    Else
        If cboFind.Visible Then
        Select Case cboFind.Text
            Case "סԺ��", "�����", "����ID"
                If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "���֤��"
                If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "����"
        End Select
        End If
    End If
End Sub


Private Sub LoadPati()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Dim colPati As Collection, str����ids As String, i As Long
    
    On Error GoTo errH
    
    
    If (optFind(1).Value Or optFind(2).Value Or optFind(3).Value) And txtFind.Text = "" Then Exit Sub
    '�����Ҳ��ҡ�����ʶ����
    If optFind(0).Value = True Or optFind(1).Value = True Then
        If cboFind.Text = "�����" Then
            strSql = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����,g.ID As ����ID, d.��ʶ��,D.����ID,d.��ǰ״̬" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��, a.id As ����ID,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                        "              From ���˹Һż�¼ A" & vbNewLine & _
                        "              Where A.��¼״̬=1 And a.ִ��ʱ�� Between [2] And [3]" & IIf(txtFind.Text = "", "", " And A.�����=[4]") & ") C) D, ���ű� G" & vbNewLine & _
                        "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.����=[1]") & vbNewLine & _
                        "Order By d.����ʱ�� Desc"
        ElseIf cboFind.Text = "סԺ��" Then
            strSql = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����,g.ID As ����ID, d.��ʶ��,D.����ID,d.��ǰ״̬" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��, B.��ҳID As ����ID,decode(B.��Ժ����,null,'��Ժ','��'||B.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                        "              From ������ҳ B" & vbNewLine & _
                        "              Where  b.��Ժ���� Between [2] And [3]" & IIf(txtFind.Text = "", "", " And B.סԺ��=[4]") & ") C) D, ���ű� G" & vbNewLine & _
                        "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.����=[1]") & vbNewLine & _
                        "Order By d.����ʱ�� Desc"
        Else
        
        
            If cboFind.Text = "���֤��" Then
                Set colPati = PatiSvrGetpatiinfo(1, 0, 1240, 0, 2, txtFind.Text)
            End If
        
            If Not colPati Is Nothing Then
                If colPati.Count > 0 Then
                    For i = 1 To colPati.Count
                        If InStr("," & str����ids & ",", "," & Val(GetColVal(colPati(i), "_pati_id")) & ",") = 0 Then
                           str����ids = str����ids & "," & Val(GetColVal(colPati(i), "_pati_id"))
                        End If
                    Next
                End If
            End If
            If str����ids <> "" Then str����ids = Mid(str����ids, 2)
            If (optFind(0).Value = True And lblDept.Tag = "") Or optFind(1).Value = True Then
                strSql = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����,g.ID As ����ID, d.��ʶ��,D.����ID,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��, a.id As ����ID,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                            "              From ���˹Һż�¼ A" & vbNewLine & _
                            "              Where A.��¼״̬=1  And a.ִ��ʱ�� Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "���֤��", " A.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "����ID", "A.����ID =[4]", "����", "A.���� like [4]")) & vbNewLine & _
                            "              Union All" & vbNewLine & _
                            "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��, B.��ҳID As ����ID,decode(B.��Ժ����,null,'��Ժ','��'||B.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                            "              From ������ҳ B" & vbNewLine & _
                            "              Where b.��Ժ���� Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "���֤��", " B.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "����ID", "B.����ID =[4]", "����", "B.���� like [4]")) & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            ElseIf optFind(0).Value = True And lblDept.Tag = "����" Then
                strSql = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����,g.ID As ����ID, d.��ʶ��,D.����ID,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��, a.id As ����ID,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                            "              From ���˹Һż�¼ A" & vbNewLine & _
                            "              Where A.��¼״̬=1  And a.ִ��ʱ�� Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "���֤��", " A.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "����ID", "A.����ID =[4]", "����", "A.���� like [4]")) & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            ElseIf optFind(0).Value = True And lblDept.Tag = "סԺ" Then
                strSql = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����,g.ID As ����ID, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��, B.��ҳID As ����ID,decode(B.��Ժ����,null,'��Ժ','��'||B.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                            "              From ������ҳ B" & vbNewLine & _
                            "              Where b.��Ժ���� Between [2] And [3] " & IIf(txtFind.Text = "", "", " And " & Decode(cboFind.Text, "���֤��", " B.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([5]) As Zltools.t_Strlist))) ", "����ID", "B.����ID =[4]", "����", "B.���� like [4]")) & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Visible = False, "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            End If
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, cboDept.ItemData(cboDept.ListIndex), mdtBegin, mdtEnd, IIf(InStr(",�����,סԺ��,����ID,", cboFind.Text) > 0, Val(txtFind.Text), IIf(cboFind.Text = "����", txtFind.Text & "%", txtFind.Text)), str����ids)
    ElseIf optFind(2).Value = True Then '����ϲ���
        strSql = "Select d.Id, d.����, d.����, d.�Ա�, d.����, g.���� As ����, g.Id As ����id, d.��ʶ��,D.����ID, d.��ǰ״̬" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��, a.id As ����ID, Decode(a.ִ��״̬, 1, '��' || To_Char(a.ִ��ʱ��, 'yyyy-mm-dd') || '���������Ժ', '�������ھ���') As ��ǰ״̬" & vbNewLine & _
                "              From ���˹Һż�¼ A, ������ϼ�¼ M" & vbNewLine & _
                "              Where  a.����id = m.����id And a.Id = m.��ҳid And a.��¼״̬ = 1 And a.ִ��ʱ�� Between [1] And [2] And" & vbNewLine & _
                "                    (Exists" & vbNewLine & _
                "                     (Select 1 From ��������Ŀ¼ N Where n.Id = m.����id And (n.���� Like [3] Or n.���� Like [3] Or UPPER(n.����) Like [3])) Or Exists" & vbNewLine & _
                "                     (Select 1 From �������Ŀ¼ I,������ϱ��� Z  Where i.Id = m.���id AND i.ID=Z.���ID And (i.���� Like [3] Or i.���� Like [3] or UPPER(Z.����) like [3])))" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��, B.��ҳID As ����ID," & vbNewLine & _
                "                     Decode(b.��Ժ����, Null, '��Ժ', '��' || B.��ҳid || '��סԺ��Ժ') As ��ǰ״̬" & vbNewLine & _
                "              From ������ҳ B, ������ϼ�¼ O" & vbNewLine & _
                "              Where  b.����id = o.����id And b.��ҳid = o.��ҳid And b.��Ժ���� Between [1] And [2] And" & vbNewLine & _
                "                    (Exists (Select 1 From ��������Ŀ¼ N Where n.Id = o.����id And (n.���� Like [3] Or n.���� Like [3] Or UPPER(n.����) Like [3])) Or Exists" & vbNewLine & _
                "                     (Select 1 From �������Ŀ¼ I,������ϱ��� Z Where i.Id = O.���id AND i.ID=Z.���ID And (i.���� Like [3] Or i.���� Like [3] or UPPER(Z.����) like [3]))) ) C) D, ���ű� G" & vbNewLine & _
                "Where g.Id = d.���� And d.Top = 1" & vbNewLine & _
                "Order By d.����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mdtBegin, mdtEnd, UCase(txtFind.Text) & "%")
    ElseIf optFind(3).Value = True Then '����������
        strSql = "Select d.Id, d.����, d.����, d.�Ա�, d.����, g.���� As ����, g.Id As ����id, d.��ʶ��,D.����ID, d.��ǰ״̬" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��, a.id As ����ID," & vbNewLine & _
                "                     Decode(a.ִ��״̬, 1, '��' || To_Char(a.ִ��ʱ��, 'yyyy-mm-dd') || '���������Ժ', '�������ھ���') As ��ǰ״̬" & vbNewLine & _
                "              From ���˹Һż�¼ A, ���������¼ M, ��������Ŀ¼ N" & vbNewLine & _
                "              Where m.��������id = n.Id And a.����id = m.����id And a.Id = m.��ҳid And a.��¼״̬ = 1 And" & vbNewLine & _
                "                    a.ִ��ʱ�� Between [1] And [2] And (Upper(n.����) Like [3] Or n.���� Like [3] Or Upper(n.����) Like [3])" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��, B.��ҳID As ����ID," & vbNewLine & _
                "                     Decode(b.��Ժ����, Null, '��Ժ', '��' || b.��ҳid || '��סԺ��Ժ') As ��ǰ״̬" & vbNewLine & _
                "              From ������ҳ B, ���������¼ O, ��������Ŀ¼ V" & vbNewLine & _
                "              Where o.��������id = v.Id And b.����id = o.����id And b.��ҳid = o.��ҳid And b.��Ժ���� Between [1] And [2] And" & vbNewLine & _
                "                    (Upper(v.����) Like [3] Or v.���� Like [3] Or Upper(v.����) Like [3])) C) D, ���ű� G" & vbNewLine & _
                "Where g.Id = d.���� And d.Top = 1" & vbNewLine & _
                "Order By d.����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mdtBegin, mdtEnd, UCase(txtFind.Text) & "%")
    End If

    rptPati.Records.DeleteAll

    With rptPati
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                Set objRecord = .Records.Add()
                Set objItem = objRecord.AddItem("")
                    objItem.Icon = img16.ListImages("unCheck").Index - 1
                Set objItem = objRecord.AddItem(rsTmp!ID & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                    objItem.Icon = img16.ListImages.Item(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Index - 1
                Set objItem = objRecord.AddItem(rsTmp!�Ա� & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                Set objItem = objRecord.AddItem(rsTmp!��ʶ�� & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                Set objItem = objRecord.AddItem(rsTmp!��ǰ״̬ & "")
                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
                Set objItem = objRecord.AddItem(rsTmp!����id & "")
                rsTmp.MoveNext
            Loop
            stbThis.Panels(2).Text = "�ڵ�ǰ���˲��ҵ� " & rsTmp.RecordCount & " λ" & lblDept.Tag & "����"
        End If
        .Populate
    End With
    Exit Sub
errH:
    MsgBox "�ڵ�ǰ����δ���ҵ�����!", vbInformation, gstrSysName
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetFrom()
'���ܣ����õ��Ӳ������Ĺ��ܣ�Ƕ��ʽ��ȡ�������
    Set mobjArchiveView = New frmArchiveView
    mobjArchiveView.BorderStyle = FormBorderStyleConstants.vbBSNone '����Ϊ�ޱ߿�
    mobjArchiveView.Caption = mobjArchiveView.Caption       '�ص�����һ��
    SetParent mobjArchiveView.hwnd, picMecInfo.hwnd
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcMec
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
        .InsertItem(0, "���Ӳ�������", mobjArchiveView.hwnd, 0).Tag = "���Ӳ�������"
        .InsertItem(1, "������Ȩ", PicNo.hwnd, 0).Tag = "������Ȩ"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub


Private Sub PicNo_Resize()
    On Error Resume Next
    picNoUse.Top = PicNo.Height / 2 - picNoUse.Height / 2
    picNoUse.Left = PicNo.Width / 2 - picNoUse.Width / 2
End Sub



Private Function CheckUse(ByVal lng����ID As Long, ByVal lng����ID As Long, ByRef intTime As Integer) As String
    Dim strSql As String
    Dim strTmp As String
    Dim blnALLTime As Boolean
    Dim blnTmp As Boolean
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    
    '�Զ���Ȩ�޼��
    strSql = "Select Zl_Fun_Checkpatimec([1],[2],[3]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Zl_Fun_Checkpatimec", lng����ID, lng����ID, UserInfo.ID)
    
    If Val(rsTmp!��� & "") = 1 Then
        CheckUse = ""
        Exit Function
    End If
    
    If mrsTmp Is Nothing Then
        strSql = "Select a.Id As ��Ȩid, a.��Ȩ����, a.���ʲ���, a.������, a.���ʲ���, a.����ʱ��, f_List2str(Cast(Collect(c.��Ȩ���� || '') As t_Strlist)) As ��Ȩ��Χ" & vbNewLine & _
                " From ���Ӳ���������Ȩ A, ���Ӳ�����Ȩ������Ա B, ���Ӳ�����Ȩ���ʲ��� C" & vbNewLine & _
                " Where a.Id = b.��Ȩid And a.Id = c.��Ȩid(+) And b.��Աid = [1] And a.���ʿ�ʼʱ�� <= Sysdate And a.���ʽ���ʱ�� >= Sysdate And" & vbNewLine & _
                " a.����ʱ�� Is Null" & vbNewLine & _
                " Group By a.Id, a.��Ȩ����, a.���ʲ���, a.������, a.���ʲ���, a.����ʱ��"
        Set mrsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    End If
    If Not mrsTmp Is Nothing Then
        If mrsTmp.RecordCount > 0 Then mrsTmp.MoveFirst
        Do While Not mrsTmp.EOF
            blnTmp = False
            Select Case Val(mrsTmp!���ʲ���) '0-ȫԺ���ˣ�1-���Ʋ��ˣ�2-ָ�����Ҳ��ˣ�3-ָ�����ˣ�4-���Ϊָ�������Ĳ��ˣ�5-ָ�������Ĳ��ˡ�2-4�Ķ�������ͨ���ӱ�洢';
                Case 0 'ȫԺ����
                    strTmp = strTmp & ";" & Val(mrsTmp!��Ȩid & "") & "," & Val(mrsTmp!����ʱ�� & "")
                    blnTmp = True
                Case 1 '���Ʋ���
                    If InStr(mstrDeptIds, ";" & lng����ID & ",") > 0 Then
                        strTmp = strTmp & ";" & Val(mrsTmp!��Ȩid & "") & "," & Val(mrsTmp!����ʱ�� & "")
                        blnTmp = True
                    End If
                Case 2 'ָ�����Ҳ���
                    If InStr("," & mrsTmp!��Ȩ��Χ & ",", "," & lng����ID & ",") > 0 Then
                        strTmp = strTmp & ";" & Val(mrsTmp!��Ȩid & "") & "," & Val(mrsTmp!����ʱ�� & "")
                        blnTmp = True
                    End If
                Case 3 'ָ������
                    If InStr("," & mrsTmp!��Ȩ��Χ & ",", "," & lng����ID & ",") > 0 Then
                        strTmp = strTmp & ";" & Val(mrsTmp!��Ȩid & "") & "," & Val(mrsTmp!����ʱ�� & "")
                        blnTmp = True
                    End If
            End Select
            
            '�����ۺ�ʱ��
            If blnTmp Then
                If Val(mrsTmp!����ʱ�� & "") = 0 Then
                    blnALLTime = True
                Else
                    If intTime <> 0 And intTime <> Val(mrsTmp!����ʱ�� & "") Then
                        blnALLTime = True
                    End If
                End If
                intTime = Val(mrsTmp!����ʱ�� & "")
            End If
            
            mrsTmp.MoveNext
        Loop
        If blnALLTime Then intTime = 0
        CheckUse = Mid(strTmp, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub rptPati_SelectionChanged()
    Dim strIDs As String
    Dim lngApplyID As Long
    Dim intTime As Integer
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '���������

    If Not mobjArchiveView Is Nothing And Val(rptPati.Tag) <> Val(rptPati.SelectedRows(0).Record(col_����Id).Value) Then
    
        strIDs = CheckUse(Val(rptPati.SelectedRows(0).Record(col_����Id).Value), Val(rptPati.SelectedRows(0).Record(col_����ID).Value), intTime)
        If strIDs <> "" Then
            rptPati.Tag = Val(rptPati.SelectedRows(0).Record(col_����Id).Value)
            Me.tbcMec.Item(0).Selected = True
        Else
            rptPati.Tag = Val(rptPati.SelectedRows(0).Record(col_����Id).Value)
            Me.tbcMec.Item(1).Selected = True
            Exit Sub
        End If
    
    
        If Val(rptPati.SelectedRows(0).Record(col_����Id).Value) <> 0 Then
            rptPati.Tag = Val(rptPati.SelectedRows(0).Record(col_����Id).Value)
            Call mobjArchiveView.zlRefresh(Val(rptPati.SelectedRows(0).Record(col_����Id).Value), Val(rptPati.SelectedRows(0).Record(col_����ID).Value), strIDs, intTime)
            If strIDs = "" Then
                Me.tbcMec.Item(1).Selected = True
            End If
        End If
    End If
    rptPati.SetFocus
End Sub



Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptPati.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptPatiCheck(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(col_ѡ��))
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim objHitTest As ReportHitTestInfo
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(col_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(col_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                            rptPati.Records(i).Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(col_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(col_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                            rptPati.Records(i).Tag = "0"
                        Next
                    End If
                End If
            End If
        ElseIf rptPati.HitTest(X, Y).ht = xtpHitTestReportArea Then
            Set objHitTest = rptPati.HitTest(X, Y)
            If Not objHitTest.Column Is Nothing And Not objHitTest.Row Is Nothing Then
                If objHitTest.Column.Index = col_ѡ�� Then
                    If rptPati.SelectedRows.Count > 0 Then
                        Call rptPatiCheck(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(col_ѡ��))
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPatiCheck(Row As XtremeReportControl.IReportRow, Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(col_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(col_ѡ��).Icon = img16.ListImages.Item("AllCheck").Index - 1
        Row.Record.Tag = "1"
    End If
    rptPati.Populate
End Sub


    
Private Function GetPatiRs() As ADODB.Recordset
    '��ȡ��ѡ���˵ļ�¼��
    Dim rsCurr As New ADODB.Recordset
    Dim i As Long
    '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
    rsCurr.Fields.Append "ID", adInteger, , adFldIsNullable
    rsCurr.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "�Ա�", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "��ʶ��", adVarChar, 4000, adFldIsNullable
    rsCurr.Fields.Append "��ǰ״̬", adVarChar, 4000, adFldIsNullable

    rsCurr.CursorLocation = adUseClient
    rsCurr.LockType = adLockOptimistic
    rsCurr.CursorType = adOpenStatic
    rsCurr.Open
    
    For i = 0 To rptPati.Records.Count - 1
        If rptPati.Records(i).Tag = "1" And Val(rptPati.Records(i)(col_����Id).Value) <> 0 Then
            rsCurr.AddNew
            rsCurr!ID = Val(rptPati.Records(i)(col_����Id).Value)
            rsCurr!���� = rptPati.Records(i)(col_����).Value
            rsCurr!�Ա� = rptPati.Records(i)(col_�Ա�).Value
            rsCurr!���� = rptPati.Records(i)(col_����).Value
            rsCurr!���� = rptPati.Records(i)(col_����).Value
            rsCurr!��ʶ�� = rptPati.Records(i)(COL_��ʶ��).Value
            rsCurr!��ǰ״̬ = rptPati.Records(i)(COL_��ǰ״̬).Value
            rsCurr.Update
        End If
    Next
    If (Not rsCurr Is Nothing) And (Not rsCurr.EOF) Then
        rsCurr.MoveFirst
    Else
        Set rsCurr = Nothing
    End If
    Set GetPatiRs = rsCurr
End Function



Private Sub lblDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vRect As RECT, strSql As String
    Dim str��λ As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, Cmd_���п���, "���п���")
        objControl.Parameter = ""
        Set objControl = .Add(xtpControlButton, Cmd_סԺ����, "סԺ����")
        objControl.Parameter = "סԺ"
        Set objControl = .Add(xtpControlButton, Cmd_�������, "�������")
        objControl.Parameter = "����"
    End With
    GetWindowRect fraPatiFilter.hwnd, vRect
    objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + lblDept.Left + lblDept.Width, vRect.Top * Screen.TwipsPerPixelY + lblDept.Top
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub SetLBLFace(ByRef objCtl As Object, ByVal blnOver As Boolean)
    If blnOver Then
        If objCtl.BorderStyle = 0 Then
            objCtl.BorderStyle = 1
            objCtl.BackStyle = 1
        End If
    Else
        If objCtl.BorderStyle = 1 Then
            objCtl.BorderStyle = 0
            objCtl.BackStyle = 0
        End If
    End If
End Sub

