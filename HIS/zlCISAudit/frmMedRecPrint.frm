VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form frmMedRecPrint 
   Caption         =   "���Ӳ�����ӡ"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmMedRecPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15120
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLine 
      Caption         =   "Frame1"
      Height          =   8655
      Left            =   5160
      MousePointer    =   9  'Size W E
      TabIndex        =   39
      Top             =   120
      Width           =   45
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   4200
      Top             =   240
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   5085
      TabIndex        =   9
      Top             =   0
      Width           =   5085
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   4815
         _Version        =   589884
         _ExtentX        =   8493
         _ExtentY        =   2778
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picShow 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   120
         MouseIcon       =   "frmMedRecPrint.frx":6852
         ScaleHeight     =   270
         ScaleWidth      =   4935
         TabIndex        =   51
         Tag             =   "0"
         Top             =   120
         Width           =   4935
         Begin VB.PictureBox picUpOrDown 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   4560
            Picture         =   "frmMedRecPrint.frx":6B5C
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   52
            Top             =   0
            Width           =   270
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "��ʾ��Χ����"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   0
            TabIndex        =   53
            Top             =   45
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   4695
         TabIndex        =   34
         Top             =   6240
         Width           =   4695
         Begin VB.CommandButton cmdSet 
            Caption         =   "��������"
            Height          =   300
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   1100
         End
         Begin VB.ComboBox cboPrinterName 
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Text            =   "Combo1"
            Top             =   120
            Width           =   3375
         End
         Begin VB.CommandButton cmdPreView 
            Caption         =   "Ԥ��(&V)"
            Height          =   300
            Left            =   2160
            TabIndex        =   7
            Top             =   600
            Width           =   1100
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "��ӡ(&P)"
            Height          =   300
            Left            =   3360
            TabIndex        =   8
            Top             =   600
            Width           =   1100
         End
         Begin VB.Label lblPlugIn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��չ���ܡ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   230
            TabIndex        =   64
            Top             =   1020
            Width           =   990
         End
         Begin VB.Label lblPrint 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "����豸"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame fraFind 
         Caption         =   "ֱ�Ӳ���"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��"
         Top             =   2520
         Width           =   4815
         Begin zlIDKind.PatiIdentify PatiIdentifyFind 
            Height          =   300
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��"
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IDKindStr       =   $"frmMedRecPrint.frx":D3AE
            BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IDKindAppearance=   0
            InputAppearance =   0
            ShowSortName    =   -1  'True
            DefaultCardType =   "���￨"
            IDKindWidth     =   555
            BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AllowAutoCommCard=   -1  'True
            NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
         End
      End
      Begin VB.Frame fraScope 
         Caption         =   "��Χ����"
         Height          =   1935
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   4815
         Begin VB.ComboBox cboOutTime 
            Height          =   300
            ItemData        =   "frmMedRecPrint.frx":D445
            Left            =   960
            List            =   "frmMedRecPrint.frx":D447
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3495
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   960
            TabIndex        =   0
            Text            =   "cboDept"
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton cmdFind 
            Appearance      =   0  'Flat
            Caption         =   "����"
            Height          =   300
            Left            =   3720
            Picture         =   "frmMedRecPrint.frx":D449
            TabIndex        =   4
            Top             =   1440
            Width           =   600
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   960
            TabIndex        =   3
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   238419971
            CurrentDate     =   39998
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   960
            TabIndex        =   2
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   238419971
            CurrentDate     =   39998.8757060185
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            Caption         =   "��Ժ����"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            Caption         =   "��Ժʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   780
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   8925
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22490
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   10320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":13C9B
            Key             =   "��ҳ����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":16CD5
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":1D537
            Key             =   "��鱨��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":22FF9
            Key             =   "���鱨��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":28ABB
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":2F31D
            Key             =   "Patient"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":301F7
            Key             =   "unCheckAll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":30791
            Key             =   "CheckAll"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":30D2B
            Key             =   "סԺ����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":33D65
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":36D9F
            Key             =   "����֤��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":39DD9
            Key             =   "��ҳ��ҳһ"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":3ACB3
            Key             =   "�ٴ�·��"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":3DCED
            Key             =   "��ҳ��ҳ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":40D27
            Key             =   "������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":43D61
            Key             =   "סԺҽ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":46D9B
            Key             =   "�����¼"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":49DD5
            Key             =   "֪���ļ�"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4CE0F
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4D449
            Key             =   "CheckFill"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4DA83
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":4E0BD
            Key             =   "��ҳ����"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":510F7
            Key             =   "down"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":57959
            Key             =   "up"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":5E1BB
            Key             =   "סԺ֤"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   5160
      ScaleHeight     =   7815
      ScaleWidth      =   10095
      TabIndex        =   12
      Top             =   720
      Width           =   10095
      Begin VB.PictureBox picItemInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   4080
         ScaleHeight     =   1575
         ScaleWidth      =   2655
         TabIndex        =   37
         Top             =   5640
         Width           =   2655
         Begin VSFlex8Ctl.VSFlexGrid vsItemInfo 
            Bindings        =   "frmMedRecPrint.frx":6164D
            Height          =   555
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   735
            _cx             =   1296
            _cy             =   979
            Appearance      =   2
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
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
      Begin VB.Frame fraPati 
         BackColor       =   &H80000005&
         Caption         =   "������Ϣ"
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   10000
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   2325
            TabIndex        =   62
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����ӱ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   2325
            TabIndex        =   61
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   6840
            TabIndex        =   60
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "סԺҽʦ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   1500
            TabIndex        =   59
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ժ����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   8280
            TabIndex        =   57
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-24"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   9120
            TabIndex        =   56
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   1500
            TabIndex        =   55
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ַ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   6360
            TabIndex        =   54
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   6840
            TabIndex        =   28
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   4560
            TabIndex        =   27
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "500101198810121245"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   4560
            TabIndex        =   26
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "20150101"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   2325
            TabIndex        =   25
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Ů"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   6840
            TabIndex        =   24
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "28��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   9120
            TabIndex        =   23
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����׿��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   4560
            TabIndex        =   22
            Top             =   360
            Width           =   720
         End
         Begin VB.Image imgPatient 
            Height          =   1185
            Left            =   120
            Picture         =   "frmMedRecPrint.frx":61661
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ժ����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   6000
            TabIndex        =   21
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ժ����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   3720
            TabIndex        =   20
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���֤��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   3720
            TabIndex        =   18
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   8640
            TabIndex        =   58
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "סԺ��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   1680
            TabIndex        =   19
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�Ա�:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   6360
            TabIndex        =   17
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   4080
            TabIndex        =   16
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.PictureBox picCenter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   7  'Invert
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   10005
         TabIndex        =   13
         Top             =   1680
         Width           =   10000
         Begin VB.Frame fraCent 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   2655
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   9015
            Begin VB.VScrollBar vsc 
               Height          =   2295
               Left            =   8640
               Max             =   10
               TabIndex        =   41
               Top             =   120
               Width           =   255
            End
            Begin VB.Frame fraIn 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   2295
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   8415
               Begin VB.PictureBox picItem1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FAEADA&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   3
                  Left            =   4320
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   49
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   2
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":6252B
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image Image1 
                     Height          =   720
                     Index           =   5
                     Left            =   240
                     Picture         =   "frmMedRecPrint.frx":62B55
                     Top             =   240
                     Width           =   720
                  End
                  Begin VB.Label lblItem1 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "סԺ����"
                     ForeColor       =   &H80000008&
                     Height          =   180
                     Index           =   3
                     Left            =   240
                     TabIndex        =   50
                     Top             =   1080
                     Width           =   720
                  End
               End
               Begin VB.PictureBox picItem1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FCE8D7&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   2
                  Left            =   2760
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   47
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   1
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":65B7F
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image Image1 
                     Height          =   720
                     Index           =   4
                     Left            =   300
                     Picture         =   "frmMedRecPrint.frx":661A9
                     Top             =   300
                     Width           =   720
                  End
                  Begin VB.Label lblItem1 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "����/���"
                     ForeColor       =   &H80000008&
                     Height          =   180
                     Index           =   2
                     Left            =   240
                     TabIndex        =   48
                     Top             =   1080
                     Width           =   810
                  End
               End
               Begin VB.PictureBox picItem1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FCE8D7&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   1
                  Left            =   1380
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   45
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":691D3
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image Image1 
                     Appearance      =   0  'Flat
                     Height          =   720
                     Index           =   3
                     Left            =   300
                     Picture         =   "frmMedRecPrint.frx":697FD
                     Top             =   300
                     Width           =   720
                  End
                  Begin VB.Label lblItem1 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "סԺҽ��"
                     ForeColor       =   &H80000008&
                     Height          =   180
                     Index           =   1
                     Left            =   240
                     TabIndex        =   46
                     Top             =   1080
                     Width           =   720
                  End
               End
               Begin VB.PictureBox picItem 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FCE8D7&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1320
                  Index           =   0
                  Left            =   0
                  ScaleHeight     =   1320
                  ScaleWidth      =   1320
                  TabIndex        =   43
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1320
                  Begin VB.Image imgCHK 
                     Appearance      =   0  'Flat
                     Height          =   300
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmMedRecPrint.frx":6C827
                     Top             =   0
                     Width           =   300
                  End
                  Begin VB.Image imgItem 
                     Appearance      =   0  'Flat
                     Height          =   720
                     Index           =   0
                     Left            =   300
                     Picture         =   "frmMedRecPrint.frx":6CE51
                     Stretch         =   -1  'True
                     Top             =   300
                     Width           =   720
                  End
                  Begin VB.Label lblItem 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "��ҳ"
                     ForeColor       =   &H80000001&
                     Height          =   180
                     Index           =   0
                     Left            =   480
                     TabIndex        =   44
                     Top             =   1080
                     Width           =   360
                  End
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   3
                  Visible         =   0   'False
                  X1              =   0
                  X2              =   360
                  Y1              =   0
                  Y2              =   360
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   2
                  Visible         =   0   'False
                  X1              =   1560
                  X2              =   1320
                  Y1              =   1680
                  Y2              =   1500
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   1
                  Visible         =   0   'False
                  X1              =   960
                  X2              =   1200
                  Y1              =   1680
                  Y2              =   2040
               End
               Begin VB.Line Lin 
                  BorderColor     =   &H00FF0000&
                  Index           =   0
                  Visible         =   0   'False
                  X1              =   1320
                  X2              =   1680
                  Y1              =   1920
                  Y2              =   2280
               End
            End
         End
         Begin VB.Frame fraSplit 
            BackColor       =   &H80000005&
            Height          =   45
            Left            =   120
            TabIndex        =   14
            Top             =   420
            Width           =   10480
         End
         Begin VB.Line LineB 
            BorderColor     =   &H80000010&
            X1              =   2160
            X2              =   6480
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line LineR 
            BorderColor     =   &H00E0E0E0&
            X1              =   9480
            X2              =   9480
            Y1              =   2760
            Y2              =   480
         End
         Begin VB.Line LineL 
            BorderColor     =   &H00FFC0C0&
            X1              =   0
            X2              =   0
            Y1              =   1320
            Y2              =   3120
         End
         Begin VB.Line lineT 
            BorderColor     =   &H8000000A&
            X1              =   1920
            X2              =   7200
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   600
            TabIndex        =   33
            Top             =   120
            Width           =   90
         End
         Begin VB.Image imgAll 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   120
            Picture         =   "frmMedRecPrint.frx":6DD1B
            Top             =   60
            Width           =   300
         End
      End
      Begin XtremeSuiteControls.TabControl tbcSub 
         Height          =   495
         Left            =   600
         TabIndex        =   36
         Top             =   5880
         Visible         =   0   'False
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   6600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":6E345
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":74BA7
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":7B409
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":7B9A3
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedRecPrint.frx":7BF3D
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMedRecPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'����
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'ö��
Private Enum ENUM_COLOR
    COLOR_HIGH = &HFCE8D7
    COLOR_ITEM = &HFDF3E9
End Enum

Private Enum PATIREPORT_COLUMN
    col_ѡ�� = 0
    col_ͼ�� = 1
    col_��ӡͼ�� = 2
    col_�Ƿ��Ŀ = 3
    col_��Ŀ���� = 4
    col_סԺ�� = 5
    col_���� = 6
    col_�Ա� = 7
    col_���֤�� = 8
    Col_�������� = 9
    col_��Ժ���� = 10
    col_��Ժ���� = 11
    col_��Ժ���� = 12
    coL_סԺҽʦ = 13
    col_��ͥ��ַ = 14
    col_���� = 15
    
    '������
    col_�������� = 16
    col_����Id = col_�������� + 1        '����
    col_��ҳID = col_�������� + 2         '����
    col_��Ժ����ID = col_�������� + 3   '����
    col_��ӡ��¼ = col_�������� + 4
End Enum

Private Enum PATI_INFO
    lbl_���� = 0
    lbl_�Ա� = 1
    lbl_���� = 2
    lbl_���֤�� = 3
    lbl_סԺ�� = 4
    lbl_��Ժ���� = 5
    lbl_��Ժ���� = 6
    lbl_��Ժ���� = 7
    lbl_��ͥ��ַ = 8
    lbl_סԺҽʦ = 9
    lbl_�������� = 10
End Enum

Private Enum TAB_INFO
    tab_סԺ���� = 0
    tab_������ = 1
    tab_�����¼ = 2
    tab_֪���ļ� = 3
    tab_����֤�� = 4
    tab_���鱨�� = 5
    tab_��鱨�� = 6
    tab_סԺ֤ = 7
    tab_�������� = 8
End Enum

Public Enum Enum_Inside_Program
    p���Ӳ������� = 2250
    p�°�סԺ���� = 2252
    p�°����ﲡ�� = 2251
    p����������д = 1249
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p���Ӳ������� = 1259
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    P�°滤ʿվ = 1265
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    p������� = 1132
    pסԺ���� = 1133
    p���ò�ѯ = 1139
    p���������� = 1113
    p�Ŷӽк�����ģ�� = 1160
    p������ҩ��� = 1266
    p������˹��� = 1267
    p���Ӳ������ = 1560
    p��Ѫ��˹��� = 1268
    p����ӿ� = 2425
    p������Ȩ���� = 1080
    p��Һ�������� = 1345
    p���Ӳ�����ӡ = 1566
End Enum
'����
Private Const M_CON_CATE As String = "��ҳ����,��ҳ����,��ҳ��ҳһ,��ҳ��ҳ��,סԺҽ��,���鱨��,��鱨��,סԺ����,������,�����¼,֪���ļ�,����֤��,�ٴ�·��,סԺ֤,��������"
'����
Private mlngCount As Long
Private mbytSelect As Byte    '��¼��ѡ����
Private mintPatiCount As Integer   '��ѡ������Ŀ
Private mblnTag As Boolean    '���ڱ�ʶ����Ƿ�λĳ��������Ŀ
Private mlngPatiID As Long    '��ǰ����ID
Private mlngDeptId As Long    '��Ժ����ID
Private mlngPatiMainID As Long      '��ǰ������ҳID
Private mlngInNO As Long            'סԺ��

Private mstrPatiName As String
Private mstrCardKind As String
Private mbytRows As Byte            '���ڱ�Ƿ�������
Private mblnLoad As Boolean
Private mbytType As Byte           '���ڱ�����һ����ѡ����,����Ԥ����λ
Private mstrPrivs As String
Private mblnLIS As String        '�Ƿ����°�LIS
Private mstr���鱨���ӡ As String        '0-�ϰ�LIS�������;1-�°�LIS����ʽ
Private mstr�����Ӧ���� As String
Private mstr����Ӧ���� As String
Private mcolReport As Collection
'����
Private mbln���Ի� As Boolean         '�Ƿ����ø��Ի����
Private mintMecStandard As Integer    '������ҳ��ʽ 0-��������׼��1-�Ĵ�ʡ��׼��2-����ʡ��׼,3-����ʡ��׼
Private mstrPrintDocIDs As String '�����������ĵ�ֻ��ӡһ��


'����
Private mclsInOutMedRec As zlMedRecPage.clsInOutMedRec
Private mobjSquareCard As Object     'ҽ�ƿ����㲿��
Private mrsMedRec As ADODB.Recordset
Private mobjRichEMR As Object       '�°���Ӳ���Ԥ����ӡ����

'�¼�
Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1

Private Sub InitFace()
'����:
    Dim i As Long
    Dim intRet As Integer
    Dim arrTmp As Variant
    Dim imgTmp As Image
    Dim lblTmp As Label
    Dim lngW As Long, lngH As Long
    Dim lngPos As Long
    Dim strTmp As String, strSelect As String
    Dim blnPath As Boolean
    Dim rsTmp As ADODB.Recordset
    
    strTmp = M_CON_CATE
    '����Ȩ�޿�����ʾ��ӡ��Ŀ
    If mintMecStandard = 0 Or mintMecStandard = 3 Then
        strTmp = Replace(strTmp, "��ҳ��ҳһ,��ҳ��ҳ��,", "")
    End If
    
    If GetPrivFunc(glngSys, p�ٴ�·��Ӧ��) <> "" And mlngDeptId <> 0 Then
        blnPath = gclsPackage.GetHavePath(mlngDeptId)
        If Not blnPath Then
            strTmp = Replace(strTmp, ",�ٴ�·��", "")
        End If
    End If
    
    arrTmp = Split(strTmp, ",")
    mlngCount = UBound(arrTmp)
    imgAll.Tag = "F"
    mbytSelect = 0
    strSelect = "," & GetRegister(˽��ģ��, "��ӡ����", "��ӡ����", "1,2,3,4,5,6,7,8,9,10") & ","
    '��̬���ؿؼ�����PicItem
    For i = Lin.LBound To Lin.UBound
        Set Lin(i).Container = Me
    Next
    imgAll.Tag = "F"
    Set imgAll.Picture = imgList.ListImages("unCheck").Picture
    For i = picItem.LBound To picItem.UBound
        If i > picItem.LBound Then
            Unload lblItem(i)
            Unload imgItem(i)
            Unload imgCHK(i)
            Unload picItem(i)
        End If
    Next
    For i = 0 To mlngCount
        If i = 0 Then
            lngW = 120
            lngH = fraSplit.Top + fraSplit.Height + 120
        Else
            Load picItem(i)
            Load imgCHK(i)
            Load imgItem(i)
            Load lblItem(i)
            Set picItem(i).Container = fraIn
            Set imgCHK(i).Container = picItem(i)
            Set imgItem(i).Container = picItem(i)
            Set lblItem(i).Container = picItem(i)
            lngW = lngW + 1380  '���60�
        End If
        If lngW + 1380 > picCenter.Width Then
            lngW = 120: lngH = lngH + 1500
        End If
        '����ȱʡ����
        picItem(i).Move lngW, lngH, 1320, 1320
        picItem(i).Visible = True
        picItem(i).BackColor = picCenter.BackColor
        picItem(i).BorderStyle = 0 '�ޱ߿�
        picItem(i).Appearance = 0
        
        '���Ͻ�ͼ��
        imgCHK(i).Visible = False
        imgCHK(i).Tag = "F"     '���δѡ��
        imgCHK(i).Move 15, 15, 300, 300
        Set imgCHK(i).Picture = imgList.ListImages("unCheck").Picture
        
        '����ͼ��
        imgItem(i).Visible = True
        imgItem(i).Move 300, 300, 720, 720
        Set imgItem(i).Picture = imgList.ListImages(arrTmp(i)).Picture
        
        '�ײ����ֱ�ʶ
        lblItem(i).Visible = True
        lblItem(i).AutoSize = True
        lblItem(i).BackStyle = 0 '͸��
        lblItem(i).Caption = arrTmp(i)
        '��¼��ҳǩ�±�
        picItem(i).Tag = GetTabIndex(arrTmp(i))

        lblItem(i).Tag = ReturnItemTag(arrTmp(i))
        If InStr(strSelect, "," & lblItem(i).Tag & ",") > 0 Then
            Call SetPicItemBG(4, CInt(i))
            mblnLoad = False
            Call imgCHK_Click(CInt(i))   'ȱʡѡ��
            mblnLoad = True
        End If
        
        If lblItem(i).Caption = "��������" Then
            imgItem(i).ToolTipText = "��������̶���������ID�������� ����ҳID��������"
        End If
        
        If lblItem(i).Width > picItem(i).Width Then
            lngPos = 0
        Else
            lngPos = (picItem(i).Width - lblItem(i).Width) / 2
        End If
        lblItem(i).Move lngPos, 1050
    Next
End Sub

Private Sub cboDept_Click()
    If Not Me.Visible Then Exit Sub
    
    cboDept.Tag = cboDept.ItemData(cboDept.ListIndex)
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strDeptIDs As String
    
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboDept.Text <> "" Then
            Set rsTmp = GetDataToDepts(cboDept.Text)
            If Not rsTmp.EOF Then
                Call cbo.SeekIndex(cboDept, rsTmp!ID)
            Else
                cboDept.ListIndex = Val(cboDept.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
        End If
    End If
End Sub

Private Sub cboOutTime_Click()
    Dim datCurr As Date
    Dim intDateCount As Integer
    
    intDateCount = cboOutTime.ItemData(cboOutTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpBegin.Enabled = False
    dtpEnd.Enabled = False
   
    If intDateCount = -1 Then
        dtpBegin.Enabled = True
        dtpEnd.Enabled = True
    ElseIf intDateCount = 0 Then
        dtpBegin.Value = Format(datCurr, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(datCurr, "yyyy-MM-dd 23:59:59")
    Else
        dtpEnd.Value = datCurr
        dtpBegin.Value = datCurr - intDateCount
    End If
  
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(glngSys, p���Ӳ�����ӡ, Control.Parameter, mlngPatiID, mlngPatiMainID, 0)
            Call zlPlugInErrH(Err, "ExecuteFunc")
            Err.Clear: On Error GoTo 0
        End If
    End Select
End Sub

Private Sub cmdFind_Click()
    Call ReadPati(2)
End Sub

Private Sub cmdPreview_Click()
    Call FuncPrintPreview
End Sub

Private Sub cmdPrint_Click()
    Call FuncPrint
End Sub

Private Sub cmdSet_Click()
    Dim i As Long
    Dim objFrm As New frmParaSet
    Dim lngRow As Long
    
    Call objFrm.ShowMe(Me, glngSys, glngModul, mstrPrivs)
    '���¼���
    Call FuncLoadReport
    
    If mrsMedRec Is Nothing Then Exit Sub
    
    Call GetRsMedRec(mlngPatiID, mlngPatiMainID, mlngDeptId, mrsMedRec)  '�������ݼ��°�LIS�����ܲ���Ӱ�������¼���
    
    'ȱʡ��λ��һ����ʾ��ѡ�е�ҳǩ
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub.Item(i).Visible Then
            tbcSub.Item(i).Selected = True
            Call tbcSub_SelectedChanged(tbcSub.Item(i))
            Exit Sub
        End If
    Next
    
    '��Ŀ����������ٽ�������
    Call picMain_Resize
End Sub

Private Sub Form_Load()
    Dim dteTime As Date
    Dim strPrinterName As String
    Dim intCount As Integer
    Dim objMenu As CommandBarPopup
    
    mblnLoad = False
    mintPatiCount = 0
    
    '��ȡ����
    '������ҳ��׼
    mintMecStandard = Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    mbln���Ի� = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0
    mstr����Ӧ���� = zlDatabase.GetPara("����Ӧ����", glngSys, p���Ӳ�����ӡ)
    mstr�����Ӧ���� = zlDatabase.GetPara("�����Ӧ����", glngSys, p���Ӳ�����ӡ)
    mstr���鱨���ӡ = zlDatabase.GetPara("���鱨���ӡ", glngSys, p���Ӳ�����ӡ)
    mblnLIS = sys.IsSysSetUp(2500)
    
    'Ȩ��
    mstrPrivs = GetPrivFunc(glngSys, p���Ӳ�����ӡ)
    'ҽ�ƿ�����
    mstrCardKind = "ס|סԺ��|0|0|0|0|0|0;��|���￨|0|0|8|0|0|0;��|����|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    Err.Clear: On Error GoTo 0
    If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
        Set mobjSquareCard = Nothing
        MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
    End If
    If Not mobjSquareCard Is Nothing Then Call PatiIdentifyFind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISAudit")
    'RIS�ӿڴ���
    Call CreateXWHIS(True)
    
    If mblnLIS Then Call InitObjLis(True)
    '�°���Ӳ���
    If Not gobjEmr Is Nothing Then
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            Set gobjEmr = Nothing
        Else
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    '-----------------------------------------------------
    
     With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(tab_סԺ����, "סԺ����", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_������, "������", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_�����¼, "�����¼", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_֪���ļ�, "֪���ļ�", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_����֤��, "����֤��", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_���鱨��, "���鱨��", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_��鱨��, "��鱨��", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_סԺ֤, "סԺ֤", picItemInfo.hWnd, 0).Visible = False
        .InsertItem(tab_��������, "��������", picItemInfo.hWnd, 0).Visible = False
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    

    Call InitVSTable
    Call ClearPatiInfo
    Call InitReportColumn
    '���Ҽ���
    Call InitDepts
    '���س�Ժ����
    Call InitOutTime
    Call InitFace
    '���ش�ӡ�豸
    strPrinterName = GetRegister(˽��ģ��, "��ӡ����", "��ӡ��", Printer.DeviceName)
    
    With cboPrinterName
        .Clear
        For intCount = 0 To Printers.count - 1
            .AddItem Printers(intCount).DeviceName
            If Printers(intCount).DeviceName = strPrinterName Then .ListIndex = intCount
        Next
    End With
    
    Call zlControl.CboSetWidth(cboPrinterName.hWnd, 3000)
    Call SetItemInfoTab
    
    '����,�˴�ֻ�����������ձ�������
    cbsMain.ActiveMenuBar.Visible = False
    Call FuncLoadReport
    
    '��Ҳ˵�
    lblPlugIn.Visible = False
    If CreatePlugInOK(p���Ӳ�����ӡ) Then
        CommandBarsGlobalSettings.App = App
        CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
        CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        cbsMain.VisualTheme = xtpThemeOffice2003
        With Me.cbsMain.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '����VisualTheme����Ч
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
            .SetIconSize False, 16, 16
        End With
        cbsMain.EnableCustomization False
        cbsMain.ActiveMenuBar.Visible = False
        Set cbsMain.Icons = zlCommFun.GetPubIcons
    
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����(&K)", 0, False)
        objMenu.ID = conMenu_Tool_PlugIn
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    If mbln���Ի� Then
        picShow.Tag = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Caption, "��Χ����", 1)
    Else
        picShow.Tag = "0"
    End If
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, True)
    
    mblnLoad = True
    
End Sub

Private Sub SetItemInfoTab(Optional ByVal intIndex As Integer)
    Dim i As Long
    Dim idex As Long
    Dim intNum As Integer
    Dim blnTmp As Boolean
        
    'ͳ�Ʋ���¼��ӡ���
    With tbcSub
        For i = picItem.LBound To picItem.UBound
            If lblItem(i).Caption = "סԺ����" Then
                .Item(tab_סԺ����).Visible = imgCHK(i).Tag = "T"
                .Item(tab_סԺ����).Tag = i
            ElseIf lblItem(i).Caption = "������" Then
                .Item(tab_������).Visible = imgCHK(i).Tag = "T"
                .Item(tab_������).Tag = i
            ElseIf lblItem(i).Caption = "�����¼" Then
                .Item(tab_�����¼).Visible = imgCHK(i).Tag = "T"
                .Item(tab_�����¼).Tag = i
            ElseIf lblItem(i).Caption = "֪���ļ�" Then
                .Item(tab_֪���ļ�).Visible = imgCHK(i).Tag = "T"
                .Item(tab_֪���ļ�).Tag = i
            ElseIf lblItem(i).Caption = "����֤��" Then
                .Item(tab_����֤��).Visible = imgCHK(i).Tag = "T"
                .Item(tab_����֤��).Tag = i
            ElseIf lblItem(i).Caption = "���鱨��" Then
                .Item(tab_���鱨��).Visible = imgCHK(i).Tag = "T"
                .Item(tab_���鱨��).Tag = i
            ElseIf lblItem(i).Caption = "��鱨��" Then
                .Item(tab_��鱨��).Visible = imgCHK(i).Tag = "T"
                .Item(tab_��鱨��).Tag = i
            ElseIf lblItem(i).Caption = "סԺ֤" Then
               .Item(tab_סԺ֤).Visible = imgCHK(i).Tag = "T"
               .Item(tab_סԺ֤).Tag = i
            ElseIf lblItem(i).Caption = "��������" Then
               .Item(tab_��������).Visible = imgCHK(i).Tag = "T"
               .Item(tab_��������).Tag = i
            End If
        Next
        
        '���ѡ����Ŀ
        If intIndex > 0 Then
            If InStr(",סԺ����,������,�����¼,֪���ļ�,����֤��,���鱨��,��鱨��,סԺ֤,��������,", "," & lblItem(intIndex).Caption & ",") > 0 And imgCHK(intIndex).Tag = "F" And Not mrsMedRec Is Nothing Then
                Select Case lblItem(intIndex).Caption
                Case "סԺ����"
                    mrsMedRec.Filter = "�ϼ�ID='R2'"
                Case "������"
                    mrsMedRec.Filter = "�ϼ�ID='R3'"
                Case "�����¼"
                    mrsMedRec.Filter = "�ϼ�ID='R4'"
                Case "֪���ļ�"
                    mrsMedRec.Filter = "�ϼ�ID='R8'"
                Case "����֤��"
                    mrsMedRec.Filter = "�ϼ�ID='R7'"
                Case "���鱨��"
                    mrsMedRec.Filter = "�ϼ�ID='R6' And EPRId ='E' "
                Case "��鱨��"
                    mrsMedRec.Filter = "�ϼ�ID='R6' And EPRId ='D' "
                Case "סԺ֤"
                    mrsMedRec.Filter = "�ϼ�ID ='R10'"
                Case "��������"
                    mrsMedRec.Filter = "�ϼ�ID ='R11'"
                End Select
                If mrsMedRec.RecordCount > 0 Then
                    Do While Not mrsMedRec.EOF
                        mrsMedRec!�Ƿ�ѡ�� = 0  'ȱʡ��ѡ��
                        mrsMedRec.MoveNext
                    Loop
                End If
            End If
        End If
        
        For i = 0 To .ItemCount - 1
            If imgCHK(Val(.Item(i).Tag)).Tag = "T" Then
                If blnTmp = False Then blnTmp = True
                intNum = intNum + 1
            End If
        Next
        .Visible = blnTmp
        If intNum = 0 Or intNum = 1 Then picMain_Resize     '���ص�һ��ҳǩ����������ҳǩʱ���ý���
        
        'ȱʡѡ�е�ǰ��ѡ��Ŀ;�����ȡ����ȱʡ��ѡTab��һ��ҳǩ
        If .Visible Then
            idex = CLng(picItem(intIndex).Tag)
            If idex >= 0 Then
                If .Item(idex).Visible Then
                    .Item(idex).Selected = True
                Else
                    For i = 0 To .ItemCount - 1
                        If .Item(i).Visible Then .Item(i).Selected = True: Exit Sub
                    Next
                End If
                Call LoadItemInfo
            End If
        End If
        
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Not mblnLoad Then Exit Sub
    picPati.Move 0, 0, (Me.ScaleWidth / 10) * 3, Me.ScaleHeight - stbThis.Height
    fraLine.Move picPati.Left + picPati.Width, 0, 45, Me.ScaleHeight - stbThis.Height
    PicMain.Move fraLine.Left + fraLine.Width, 0, Me.ScaleWidth - fraLine.Width - fraLine.Left, Me.ScaleHeight - stbThis.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbln���Ի� Then SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Caption, "��Χ����", picShow.Tag
    Call SaveWinState(Me, App.ProductName)
    Set mclsInOutMedRec = Nothing
    Set mrsMedRec = Nothing
    Set mcolReport = Nothing
    Set mobjRichEMR = Nothing
    If Not mclsDockAduits Is Nothing Then Set mclsDockAduits = Nothing
    If Not mobjSquareCard Is Nothing Then Set mobjSquareCard = Nothing
End Sub

Private Sub FraLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If fraLine.Left + x < (Me.ScaleWidth / 10) * 1 Or fraLine.Left + x > (Me.ScaleWidth / 10) * 9 Or Abs(x) < 100 Then Exit Sub
        fraLine.Left = fraLine.Left + x
        picPati.Width = picPati.Width + x
        PicMain.Left = PicMain.Left + x
        PicMain.Width = PicMain.Width - x
    End If
End Sub

Private Sub imgItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tbcSub.Visible Then
            If Val(picItem(Index).Tag) >= 0 Then
                If tbcSub.Item(picItem(Index).Tag).Visible Then
                    tbcSub.Item(picItem(Index).Tag).Selected = True
                End If
            End If
        End If
        Call FuncSetFocus(Index)
    End If
End Sub

Private Sub lblNote_Click()
    Call picShow_Click
End Sub

Private Sub lblPlugIn_Click()
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_Tool_PlugIn)
    If Not objPopup Is Nothing Then
        objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub

Private Sub PatiIdentifyFind_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strName As String
    Dim vRect As RECT
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    strSQL = ""
    strName = Trim(PatiIdentifyFind.Text)
    If objCard.���� Like "*��*��*" And blnCard = False And strName <> "" And InStr("-*+/", Left(strName, 1)) = 0 Then
        strSQL = "Select 1 As ����id, a.����id As ID, b.��ҳid, b.סԺ��, NVL(b.����,a.����) as ����, NVL(b.�Ա�,a.�Ա�) as �Ա�, NVL(b.����,a.����) as ����, a.���֤��, b.��Ժ����, b.��Ժ����, a.��������, a.סԺ����" & vbNewLine & _
                "From ������Ϣ A, ������ҳ B" & vbNewLine & _
                "Where a.����id = b.����id And a.���� Like [1] And b.��Ժ���� Is Not Null" & vbNewLine & _
                "Order By ����id, ����, ��Ժ���� Desc"
    ElseIf (objCard.���� = "סԺ��" Or Left(strName, 1) = "+") And IsNumeric(Mid(strName, 2)) And blnCard = False Then
        strSQL = "Select *" & vbNewLine & _
                "From (Select 1 As ����id, a.����id As ID, b.��ҳid, b.סԺ��, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) As ����," & vbNewLine & _
                "              a.���֤��, b.��Ժ����, b.��Ժ����, a.��������, a.סԺ����" & vbNewLine & _
                "       From ������Ϣ A, ������ҳ B" & vbNewLine & _
                "       Where a.����id = b.����id And b.סԺ�� = [2] And b.��Ժ���� Is Not Null" & vbNewLine & _
                "       Order By ��Ժ���� Desc) A" & vbNewLine & _
                "Where Rownum < 2"

    End If
    
    If strSQL <> "" Then
        vRect = zlControl.GetControlRect(PatiIdentifyFind.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, PatiIdentifyFind.Height, blnCancel, False, True, strName & "%", strName)
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!ID) = 0 Then
                blnCancel = True: Exit Sub
            Else '�Բ���ID��ȡ
                mlngPatiID = NVL(rsTmp!ID)
                Call ReadPati(1)
                blnCancel = True: Exit Sub
            End If
        Else 'ȡ��ѡ��
            If blnCancel = False Then
                MsgBox "û���ҵ����������Ĳ��ˣ�", vbInformation, gstrSysName
            End If
            blnCancel = True: Exit Sub
        End If

    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgAll_Click()
    If imgAll.Tag = "T" Then
        imgAll.Tag = "F"
        Set imgAll.Picture = imgList.ListImages("unCheck").Picture
        Call SetPicItemBG(1)
        Call FuncSetFocus(0)
    Else
        imgAll.Tag = "T"
        Set imgAll.Picture = imgList.ListImages("Check").Picture
        Call SetPicItemBG(2)
    End If
End Sub

Private Sub imgCHK_Click(Index As Integer)
    If imgCHK(Index).Tag = "T" Then
        Set imgCHK(Index).Picture = imgList.ListImages("unCheck").Picture
        imgCHK(Index).Tag = "F"
        mbytSelect = mbytSelect - 1
        If mbytSelect = mlngCount And imgAll.Tag = "T" Then
            imgAll.Tag = "F"
            Set imgAll.Picture = imgList.ListImages("unCheck").Picture
        End If
    Else
        Set imgCHK(Index).Picture = imgList.ListImages("CheckFill").Picture
        imgCHK(Index).Tag = "T"
        mbytSelect = mbytSelect + 1
        If mbytSelect = mlngCount + 1 And imgAll.Tag = "F" Then
            imgAll.Tag = "T"
            Set imgAll.Picture = imgList.ListImages("Check").Picture
        End If
    End If
    If mblnLoad Then Call SetItemInfoTab(Index)
    Call FuncSetFocus(Index)
End Sub

Private Sub imgItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetPicItemBG(3, Index)
    Call SetPicItemBG(4, Index)
End Sub

Private Sub PatiIdentifyFind_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    
    blnCancel = False
    If objHisPati Is Nothing Then blnCancel = True
    If blnCancel = False Then
        If objHisPati.����ID = 0 Then blnCancel = True
    End If
    
    If blnCancel Then
        MsgBox "û���ҵ����������Ĳ��ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    mlngPatiID = objHisPati.����ID
    
    Call ReadPati(1)
End Sub

Private Sub picCenter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'����:ȡ��ѡ����Ŀ
    Call SetPicItemBG(3)
End Sub

Private Sub picCenter_Resize()
    Dim i As Long
    Dim lngW As Long, lngH As Long
    Dim blnScor As Boolean
    
    On Error Resume Next
    fraSplit.Move 120, 420, picCenter.Width - 240, 45
    imgAll.Move 120, 60, 300, 300
    fraCent.Move 120, fraSplit.Top + fraSplit.Height, picCenter.Width - 135, picCenter.Height - (fraSplit.Top + fraSplit.Height) - 15
    fraIn.Move 0, 0, fraCent.Width - 255, fraCent.Height
    
    '������ɫ����
    lineT.X1 = 0: lineT.Y1 = 0
    lineT.X2 = picCenter.Width: lineT.Y2 = 0
    lineT.BorderColor = &H80000010
    
    LineB.X1 = 0: LineB.Y1 = picCenter.Height - 15
    LineB.X2 = picCenter.Width: LineB.Y2 = picCenter.Height - 15
    LineB.BorderColor = lineT.BorderColor
    
    LineL.X1 = 0: LineL.Y1 = 0
    LineL.X2 = 0: LineL.Y2 = picCenter.Height
    LineL.BorderColor = lineT.BorderColor
    
    LineR.X1 = picCenter.Width - 15: LineR.Y1 = 0
    LineR.X2 = picCenter.Width - 15: LineR.Y2 = picCenter.Height
    LineR.BorderColor = lineT.BorderColor
    '�����Ű�
    mbytRows = 1
    For i = picItem.LBound To picItem.UBound
        If i = 0 Then
            lngW = 60
            lngH = 60
        Else
            lngW = lngW + 1380  '���120�
        End If
        If lngW + 1380 > fraIn.Width Then
            lngW = 60: lngH = lngH + 1500
            mbytRows = mbytRows + 1
            '�����߽���ʾ������
            If lngH + picItem(i).Height + 60 > fraIn.Height Then
                fraIn.Height = fraIn.Height + picItem(i).Height + 120
                blnScor = True
            End If
        End If
        '����ȱʡ����
        picItem(i).Move lngW, lngH, 1320, 1320
    Next
    vsc.Visible = blnScor
    If blnScor Then
        vsc.Max = mbytRows
        vsc.Move fraCent.Width - vsc.Width - 15, 0, 255, fraCent.Height
    End If
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call imgItem_MouseDown(Index, Button, Shift, x, y)
    End If
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetPicItemBG(3, Index)
    Call SetPicItemBG(4, Index)
End Sub

Private Sub SetPicItemBG(ByVal bytFunc As Byte, Optional ByVal Index As Integer = -1)
'����:�������Pic����,�Ƴ���ʾ����ɫ
'����:
'   bytFunc-1 ������� ;2-ѡ������;3-���δѡ���κ���Ŀ,��ո���״̬;4-��궨λĳ����Ŀ,����ɫ����
'   Index- bytFunc=4ʱ����
    Dim i As Integer
    
    On Error Resume Next
    If bytFunc = 1 Then
        For i = 0 To mlngCount
            If imgCHK(i).Tag = "T" Then
                picItem(i).BackColor = picCenter.BackColor
                imgCHK(i).Visible = False
                Call imgCHK_Click(i)
            End If
        Next
        mbytSelect = 0
    ElseIf bytFunc = 2 Then
        For i = 0 To mlngCount
            If imgCHK(i).Tag = "F" Then
                picItem(i).BackColor = COLOR_HIGH
                imgCHK(i).Visible = True
                Call imgCHK_Click(i)
            End If
        Next
        mbytSelect = mlngCount + 1
    ElseIf bytFunc = 3 Then
        '���δ��λ����Ŀ
        For i = 0 To mlngCount
            If imgCHK(i).Tag = "F" And imgCHK(i).Visible = True Then
                If i <> Index Then
                    imgCHK(i).Visible = False
                    picItem(i).BackColor = picCenter.BackColor
                End If
            End If
        Next
        mblnTag = False
    ElseIf bytFunc = 4 And Index <> -1 Then
        If imgCHK(Index).Visible = False Then
            imgCHK(Index).Visible = True
            picItem(Index).BackColor = COLOR_HIGH
        End If
        mblnTag = True
    End If
End Sub

Private Sub picItemInfo_Resize()
    On Error Resume Next
    vsItemInfo.Move 0, 0, picItemInfo.Width, picItemInfo.Height
End Sub

Private Sub picMain_Resize()
    Dim lngW As Long
    On Error Resume Next
    
    If Not mblnLoad Then Exit Sub
    lngW = PicMain.Width - 120
    fraPati.Move 60, 60, lngW, 1575
    If tbcSub.Visible Then
        picCenter.Move 60, fraPati.Top + fraPati.Height + 120, lngW, 3650
        picCenter.Height = IIf(mbytRows > 1, 3560, 2200)
        tbcSub.Move 60, picCenter.Top + picCenter.Height, picCenter.Width, PicMain.Height - picCenter.Height - fraPati.Height - 200
    Else
        picCenter.Move 60, fraPati.Top + fraPati.Height + 120, lngW, PicMain.Height - fraPati.Top - fraPati.Height - 120
    End If
End Sub

Private Sub picPati_Resize()
    Dim lngW As Long
    Dim lngPos As Long
    
    On Error Resume Next
    If Not mblnLoad Then Exit Sub
    lngW = picPati.ScaleWidth - 240
    lngW = IIf(lngW < 4000, 4000, lngW)
    picShow.Move 120, 60, lngW, 270
    
    If picShow.Tag = "0" Then
        lblNote.Caption = "���ط�Χ����"
        Set picUpOrDown.Picture = imgList.ListImages("up").Picture
        fraScope.Visible = True
        fraScope.Move 120, picShow.Top + picShow.Height + 60, lngW, 1965
        fraFind.Move 120, fraScope.Top + fraScope.Height + 120, lngW, 975
    Else
        lblNote.Caption = "��ʾ��Χ����"
        Set picUpOrDown.Picture = imgList.ListImages("down").Picture
        fraScope.Visible = False
        fraFind.Move 120, picShow.Top + picShow.Height + 60, lngW, 975
    End If
    
    lngPos = fraFind.Top + fraFind.Height + 120
    rptPati.Move 120, lngPos, picPati.ScaleWidth - 240, picPati.Height - lngPos - 1335
    picPrint.Move 120, rptPati.Top + rptPati.Height, lngW, 1335
    
    lngW = fraScope.Width - 120 - 960
    lngW = IIf(lngW < 2700, 2700, lngW)
    
    dtpBegin.Move 960, 1080, 2175, 300
    dtpEnd.Move 960, dtpBegin.Top + dtpBegin.Height + 120, 2175, 300
    cboDept.Left = 960: cboDept.Top = 360
    cboDept.Width = lngW: cboDept.Height = 300
    cboOutTime.Left = 960: cboOutTime.Top = 720
    cboOutTime.Width = lngW: cboOutTime.Height = 300
    cmdFind.Move fraScope.Width - cmdFind.Width - 140, dtpBegin.Top + dtpBegin.Height + 120, 600, 300
 
    PatiIdentifyFind.Move 120, 360, fraFind.Width - 240, 300
End Sub

Private Sub picPrint_Resize()
    On Error Resume Next
    lblPrint.Move 120, 120, 720, 180
    cboPrinterName.Move lblPrint.Left + lblPrint.Width + 60, 60, picPrint.Width - (lblPrint.Left + lblPrint.Width + 180)
    cmdSet.Move 60, cboPrinterName.Top + cboPrinterName.Height + 200
    cmdPreView.Move picPrint.ScaleWidth - cmdPrint.Width * 2 - 240, cboPrinterName.Top + cboPrinterName.Height + 200
    cmdPrint.Move picPrint.ScaleWidth - cmdPrint.Width - 120, cboPrinterName.Top + cboPrinterName.Height + 200
    lblPlugIn.Move 140, cmdSet.Top + cmdSet.Height + 200
End Sub

Private Sub picShow_Click()
    If picShow.Tag = "1" Then
        picShow.Tag = "0"
    Else
        picShow.Tag = "1"
    End If
    Call picPati_Resize
End Sub

Private Sub picShow_Resize()
    lblNote.Move picShow.Left, 0, lblNote.Width, 270
    With picUpOrDown
        .Width = 270
        .Height = 270
        .Left = picShow.Width - picUpOrDown.Width - 120
        .Top = 0
    End With
End Sub

Private Sub picUpOrDown_Click()
    Call picShow_Click
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Item.Checked = True Then
        mintPatiCount = mintPatiCount + 1
    Else
        mintPatiCount = mintPatiCount - 1
    End If
    If mintPatiCount = rptPati.Records.count Then
        rptPati.Columns(col_ѡ��).Icon = imgPati.ListImages("Check").Index - 1
    Else
        rptPati.Columns(col_ѡ��).Icon = imgPati.ListImages("UnCheck").Index - 1
    End If
End Sub

Private Sub rptPati_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim hitColumn As ReportColumn
    Dim lngHit As Long
    
    If Button = 1 Then
        Set hitColumn = rptPati.HitTest(x, y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.ItemIndex = col_ѡ�� Then
                lngHit = rptPati.HitTest(x, y).ht
                If xtpHitTestHeader = lngHit Then
                    If rptPati.Records.count = 0 Then Exit Sub  '������ʱ��ֹ�л�
                    If hitColumn.Icon = imgPati.ListImages("Check").Index - 1 Then
                        hitColumn.Icon = imgPati.ListImages("UnCheck").Index - 1
                        SelectItems 2
                    Else
                        hitColumn.Icon = imgPati.ListImages("Check").Index - 1
                        SelectItems 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim hitColumn As ReportColumn
    Dim Item As ReportRecordItem
    Dim strTipInfo As String
    Dim vPos As PointAPI
    Dim lngHwnd As Long
    
    On Error Resume Next
    Set hitColumn = rptPati.HitTest(x, y).Column
    If Not hitColumn Is Nothing Then
        If hitColumn.Index = col_��ӡͼ�� Then
            Set Item = rptPati.HitTest(x, y).Item
            If Not Item Is Nothing Then
                If Item.Record(col_��ӡͼ��).Icon <> -1 Then
                    strTipInfo = Item.Record(col_��ӡ��¼).Value
                    If strTipInfo = "" Then '���û�л�ȡ������������ȡ����¼���б���
                        strTipInfo = GetPrintLog(Item.Record(col_����Id).Value, Item.Record(col_��ҳID).Value) '��ȡ��ӡ��¼
                        Item(col_��ӡ��¼).Value = strTipInfo
                    End If
                    GetCursorPos vPos
                    lngHwnd = WindowFromPoint(vPos.x, vPos.y)
                    Call zlCommFun.ShowTipInfo(lngHwnd, strTipInfo, True)
                End If
            End If
        Else
            Call zlCommFun.ShowTipInfo(lngHwnd, "")
        End If
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    Dim lngRow As Long
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    If Not Me.Visible Then Exit Sub
    If rptPati.SelectedRows.count = 0 Then Exit Sub          '���������
    With rptPati.SelectedRows(0)
    
        If .GroupRow Then
            Call ClearPatiInfo
        Else
            If rptPati.Tag = Val(.Record(col_����Id).Value & "") & "_" & Val(.Record(col_��ҳID).Value & "") Then Exit Sub
            '������Ƭ
            If Not ReadPatPricture(Val(.Record(col_����Id).Value & ""), imgPatient) Then
               Set imgPatient.Picture = imgList.ListImages("Patient").Picture
            End If
            lblShow(lbl_����).Caption = .Record(col_����).Value
            lblShow(lbl_����).FontBold = True
            lblShow(lbl_����).ForeColor = IIf(.Record(col_��������).Value = "��ͨ����" Or .Record(col_��������).Value = "", &H0&, vbRed)
            lblShow(lbl_����).Caption = .Record(col_����).Value
            lblShow(lbl_�Ա�).Caption = .Record(col_�Ա�).Value
            lblShow(lbl_���֤��).Caption = .Record(col_���֤��).Value
            lblShow(lbl_סԺ��).Caption = .Record(col_סԺ��).Value
            lblShow(lbl_��Ժ����).Caption = .Record(col_��Ժ����).Value
            lblShow(lbl_��Ժ����).Caption = .Record(col_��Ժ����).Value
            lblShow(lbl_��Ժ����).Caption = .Record(col_��Ժ����).Value
            lblShow(lbl_��������).Caption = .Record(Col_��������).Value
            lblShow(lbl_סԺҽʦ).Caption = .Record(coL_סԺҽʦ).Value
            lblShow(lbl_��ͥ��ַ).Caption = .Record(col_��ͥ��ַ).Value
            
            mlngPatiID = Val(.Record(col_����Id).Value & "")
            mlngPatiMainID = Val(.Record(col_��ҳID).Value & "")
            mlngDeptId = Val(.Record(col_��Ժ����ID).Value & "")
            mlngInNO = Val(.Record(col_סԺ��).Value & "")
            mstrPatiName = Val(.Record(col_����).Value & "")
            
            If lblPlugIn.Visible Then lblPlugIn.Enabled = mlngPatiID <> 0
                
            Call GetRsMedRec(mlngPatiID, mlngPatiMainID, mlngDeptId, mrsMedRec)
            
            rptPati.Tag = mlngPatiID & "_" & mlngPatiMainID
            
            'ȱʡ��λ��һ����ʾ��ѡ�е�ҳǩ
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    tbcSub.Item(i).Selected = True
                    Call tbcSub_SelectedChanged(tbcSub.Item(i))
                    Exit Sub
                End If
            Next
            
            '��Ŀ����������ٽ�������
            Call picMain_Resize
        End If
    End With
End Sub

Private Sub GetRsMedRec(ByVal lngPatiID As Long, ByVal lngPatiMainID As Long, ByVal lngDeptId As Long, ByRef rsMedRec As ADODB.Recordset, Optional ByVal blnMorePati As Boolean)
    Dim i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim strTemp As String
    Dim strErr As String
    Dim arrTmp As Variant
    Dim arrAdvice As Variant
    
    Set rsMedRec = gclsPackage.GetCISStruct(lngPatiID, lngPatiMainID, lngDeptId, False, "סԺ֤") '�л�һ�β��˶�ȡһ��
    Set rsMedRec = zlDatabase.CopyNewRec(rsMedRec, False, "", Array("�Ƿ�ѡ��", adInteger, 2, Empty, "Ԥ��", adInteger, 2, Empty))
    '
    '׷����������
    For i = 1 To mcolReport.count
        rsMedRec.AddNew
        rsMedRec!ID = mcolReport(i)
        rsMedRec!�ϼ�ID = "R11"
        rsMedRec!���� = Split(mcolReport(i), ",")(0)
        rsMedRec!���� = Split(mcolReport(i), ",")(0) & ";" & Split(mcolReport(i), ",")(1) & ";" & Split(mcolReport(i), ",")(2)  '��������,ϵͳ��,������
        rsMedRec.Update
        If i = mcolReport.count Then
            rsMedRec.MoveFirst
        End If
    Next

    If Not gobjLIS Is Nothing And Val(mstr���鱨���ӡ) = 1 Then
        strTemp = gobjLIS.GetPatientAdvice(mlngPatiID, mlngPatiMainID, strErr)  'ҽ��֮����","�ָ�걾֮����";"�ָ� 8362586,8362588;8362590
        If strErr <> "" Then MsgBox "LIS������ȡҽ��IDʧ�ܣ�" & vbCrLf & strErr, vbInformation, Me.Caption
        If strTemp <> "" Then
            arrTmp = Split(strTemp, ";")
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrAdvice = Split(arrTmp(i), ",")
                strTemp = ""
                For j = LBound(arrAdvice) To UBound(arrAdvice)
                    If UBound(arrAdvice) = 0 Then Exit For
                    rsMedRec.Filter = "�ϼ�ID = 'R6' And EPRID='E' And ID LIKE '*," & arrAdvice(j) & ",*'"
                    If j = UBound(arrAdvice) Then
                        If Not rsMedRec.EOF Then rsMedRec!���� = Mid(strTemp, 2) & "," & rsMedRec!����
                        Exit For
                    Else
                        If Not rsMedRec.EOF Then strTemp = strTemp & "," & Split(rsMedRec!���� & "", "��")(0)
                        rsMedRec.Delete
                    End If
                Next
            Next
            rsMedRec.Filter = ""
        End If
    End If
    
    Do While Not rsMedRec.EOF
        If InStr(",R1,R5,R9,", "," & rsMedRec!ID & ",") > 0 Then
            rsMedRec!�Ƿ�ѡ�� = 1  'ȱʡѡ�� ������ҳ,סԺҽ��,�ٴ�·��
        Else
            If blnMorePati Then
                rsMedRec!�Ƿ�ѡ�� = 1  '����ʱĬ��ѡ��
            Else
                rsMedRec!�Ƿ�ѡ�� = 0  'ȱʡ��ѡ��
            End If
        End If
        rsMedRec!Ԥ�� = 0
        rsMedRec.MoveNext
    Loop
    
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gclsPackage.GetEmrLIST(lngPatiID, lngPatiMainID)
    If Not rsTmp Is Nothing Then
        If rsTmp.State = ADODB.adStateOpen Then
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Do Until rsTmp.EOF
                    rsMedRec.AddNew
                    rsMedRec!ID = rsTmp!ID
                    rsMedRec!�ϼ�ID = "R2"
                    rsMedRec!���� = rsTmp!����
                    rsMedRec!���� = rsTmp!ID 'NVL(rsTmp!����) '�ĵ�ID[|���ĵ�ID]
                    If blnMorePati Then
                        rsMedRec!�Ƿ�ѡ�� = 1  'ȱʡѡ��
                    Else
                        rsMedRec!�Ƿ�ѡ�� = 0  'ȱʡ��ѡ��
                    End If
                    rsMedRec!Ԥ�� = 0
                    rsMedRec.Update
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    
    If rsMedRec.RecordCount > 0 Then
        rsMedRec.MoveFirst
    End If
End Sub
Private Sub LoadItemInfo()
'���ܣ�������ϸ��ӡ����
    Dim arrTmp As Variant
    Dim lngRow As Long
    Dim strSplit As String
    
    If mrsMedRec Is Nothing Then Exit Sub
    If tbcSub.Visible = False Then Exit Sub

    vsItemInfo.Rows = 1
    mrsMedRec.Filter = 0
    If mrsMedRec.RecordCount > 0 Then
        With tbcSub
            If .Selected.Caption = "סԺ����" Then
                mrsMedRec.Filter = "�ϼ�ID='R2'"
                strSplit = "��"
            ElseIf .Selected.Caption = "������" Then
                mrsMedRec.Filter = "�ϼ�ID='R3'"
                strSplit = "��"
            ElseIf .Selected.Caption = "�����¼" Then
                mrsMedRec.Filter = "�ϼ�ID='R4'"
                strSplit = "("
            ElseIf .Selected.Caption = "֪���ļ�" Then
                mrsMedRec.Filter = "�ϼ�ID='R8'"
                strSplit = "��"
            ElseIf .Selected.Caption = "����֤��" Then
                mrsMedRec.Filter = "�ϼ�ID='R7'"
                strSplit = "��"
            ElseIf .Selected.Caption = "���鱨��" Then
                mrsMedRec.Filter = "�ϼ�ID='R6' And EPRID = 'E'"
                strSplit = "��"
            ElseIf .Selected.Caption = "��鱨��" Then
                mrsMedRec.Filter = "�ϼ�ID='R6' And EPRID = 'D'"
                strSplit = "��"
            ElseIf .Selected.Caption = "סԺ֤" Then
                mrsMedRec.Filter = "�ϼ�ID ='R10'"
                strSplit = "��"
            ElseIf .Selected.Caption = "��������" Then
                mrsMedRec.Filter = "�ϼ�ID ='R11'"
            End If
            If mrsMedRec.RecordCount > 0 Then
                mrsMedRec.MoveFirst
                With vsItemInfo
                    .Rows = 1
                    Do While Not mrsMedRec.EOF
                        .Rows = .Rows + 1
                        .Cell(flexcpData, .Rows - 1, 1) = mrsMedRec!ID & ""
                        arrTmp = Split(mrsMedRec!���� & "", strSplit)
                        If UBound(arrTmp) = 1 Then
                            .TextMatrix(.Rows - 1, 1) = arrTmp(0)
                            .TextMatrix(.Rows - 1, 2) = strSplit & arrTmp(1)
                        Else
                            .TextMatrix(.Rows - 1, 1) = mrsMedRec!���� & ""
                        End If
                        If NVL(mrsMedRec!�Ƿ�ѡ��, 1) = 1 Then
                            lngRow = lngRow + 1
                            .Cell(flexcpChecked, .Rows - 1, 0) = 1
                        End If
                        If .Rows - 1 = 1 Then
                            mrsMedRec!Ԥ�� = 1     'ȱʡԤ����һ���ļ�
                        End If
                        mrsMedRec.MoveNext
                    Loop
                    .Row = 1
                    If lngRow = .Rows - 1 Then
                        Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("Check").Picture
                        .Cell(flexcpData, 0, 0) = 1
                    Else
                        Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                        .Cell(flexcpData, 0, 0) = 0
                    End If
                End With
            Else
                Set vsItemInfo.Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                vsItemInfo.Cell(flexcpData, 0, 0) = 0
                vsItemInfo.Rows = 2
            End If
        End With
    End If
End Sub

Private Sub ReadPati(ByVal bytFunc As Byte)
'����:��ȡ������Ϣ
'����:bytFunc =1 ����ͨ������ID��ѯ���˳�Ժ��¼
'     bytFunc=2  ����ͨ����Χ���Ҳ��˳�Ժ��¼
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngColor As Long
    
    On Error GoTo ErrH
    If bytFunc = 1 Then
        strSQL = "Select distinct a.����id, a.��ҳid,a.��Ŀ����,a.סԺ��, NVl(a.����,b.����) as ����, NVL(a.�Ա�,b.�Ա�) as �Ա�, a.����,a.��ͥ��ַ, b.���֤��,b.��������, a.��Ժ����, a.��Ժ����,a.סԺҽʦ,a.��Ժ����id,a.��������, c.���� as ��Ժ����,Decode(D.����ID,NULL,0,1) as �Ƿ��ӡ " & vbNewLine & _
            "From ������ҳ A, ������Ϣ B, ���ű� C,������ӡ��¼ D" & vbNewLine & _
            "Where a.����id = b.����id And a.��Ժ����id = c.Id And A.����ID=D.����ID(+) And A.��ҳID=D.��ҳID(+) And a.����id = [1] and a.��Ժ���� is Not NULL "
    ElseIf bytFunc = 2 Then
        strSQL = "Select distinct a.����id, a.��ҳid,a.��Ŀ����,a.סԺ��,NVL(a.����,b.����) as ����, NVL(a.�Ա�,a.�Ա�) as �Ա�, a.����,a.��ͥ��ַ, b.���֤��,b.��������,a.��Ժ����, a.��Ժ����,a.סԺҽʦ,a.��Ժ����id,a.��������, c.���� as ��Ժ����,Decode(D.����ID,NULL,0,1) as �Ƿ��ӡ " & vbNewLine & _
            "From ������ҳ A, ������Ϣ B, ���ű� C,������ӡ��¼ D" & vbNewLine & _
            "Where a.����id = b.����id And a.��Ժ����id = c.Id And A.����ID=D.����ID(+) And A.��ҳID=D.��ҳID(+) And a.��Ժ����id = [2] And a.��Ժ���� between [3] And [4]"
    End If

    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, Val(cboDept.Tag), CDate(dtpBegin.Value), CDate(dtpEnd.Value))
    '���ز����б�
    Call ClearPatiInfo
    rptPati.Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati.Records.Add()
        Set objItem = objRecord.AddItem("")    'ͼ��
        objItem.HasCheckbox = True
        Set objItem = objRecord.AddItem("")  'ͼ��
        If InStr(rsPati!�Ա� & "", "��") > 0 Then
            objItem.Icon = imgPati.ListImages("Boy").Index - 1
        ElseIf InStr(rsPati!�Ա� & "", "Ů") > 0 Then
            objItem.Icon = imgPati.ListImages("Girl").Index - 1
        End If
        Set objItem = objRecord.AddItem("")  'ͼ��
        If Val(rsPati!�Ƿ��ӡ & "") = 1 Then
            objItem.Icon = imgPati.ListImages("print").Index - 1
        End If
        
        objRecord.AddItem IIf(NVL(rsPati!��Ŀ����) <> "", "�ѱ�Ŀ", "δ��Ŀ")
        objRecord.AddItem Format(rsPati!��Ŀ���� & "", "YYYY-MM-dd")
        objRecord.AddItem rsPati!סԺ�� & ""
        objRecord.AddItem rsPati!���� & ""
        objRecord.AddItem rsPati!�Ա� & ""
        objRecord.AddItem rsPati!���֤�� & ""
        objRecord.AddItem Format(rsPati!�������� & "", "YYYY-MM-DD")
        objRecord.AddItem Format(rsPati!��Ժ���� & "", "YYYY-MM-DD")
        objRecord.AddItem Format(rsPati!��Ժ���� & "", "YYYY-MM-DD")
        objRecord.AddItem rsPati!��Ժ���� & ""
        objRecord.AddItem rsPati!סԺҽʦ & ""
        objRecord.AddItem rsPati!��ͥ��ַ & ""
        objRecord.AddItem rsPati!���� & ""
        '������
        objRecord.AddItem rsPati!�������� & ""
        objRecord.AddItem CLng(rsPati!����ID)
        objRecord.AddItem NVL(rsPati!��ҳID)
        objRecord.AddItem rsPati!��Ժ����ID & ""
        
         '��ʾ������ɫ
        lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
        objRecord.Item(col_����).ForeColor = lngColor

        rsPati.MoveNext
    Next
    rptPati.Populate
    '���ز����б����
    If rptPati.Records.count > 0 Then
        rptPati.Rows(0).Selected = True
        rptPati.SetFocus
        Call rptPati_SelectionChanged
    Else
        '���³�ʼ������
        Set mrsMedRec = Nothing
        Call InitFace
        Call SetItemInfoTab
        vsItemInfo.Rows = 1
        vsItemInfo.Rows = 2
        Call picMain_Resize
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With rptPati
        Set objCol = .Columns.Add(col_ѡ��, "", 20, False)
            objCol.Icon = imgPati.ListImages("UnCheck").Index - 1
            objCol.EditOptions.AllowEdit = True
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_ͼ��, "", 20, False)  'ͼ��
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_��ӡͼ��, "", 20, False)  'col_��ӡͼ��
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_�Ƿ��Ŀ, "�Ƿ��Ŀ", 60, True)
        Set objCol = .Columns.Add(col_��Ŀ����, "��Ŀ����", 80, True)
        Set objCol = .Columns.Add(col_סԺ��, "סԺ��", 80, True)
        Set objCol = .Columns.Add(col_����, "����", 80, True)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 45, True)
        Set objCol = .Columns.Add(col_���֤��, "���֤��", 150, True)
        Set objCol = .Columns.Add(Col_��������, "��������", 80, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 80, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 80, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 90, True)
        Set objCol = .Columns.Add(coL_סԺҽʦ, "סԺҽʦ", 80, True)
        Set objCol = .Columns.Add(col_��ͥ��ַ, "��ַ", 150, True)
        Set objCol = .Columns.Add(col_����, "����", 45, True)
        
        '������
        Set objCol = .Columns.Add(col_��������, "��������", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��Ժ����ID, "��Ժ����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ӡ��¼, "��ӡ��¼", 0, False): objCol.Visible = False
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_��Ժ����)
        .SortOrder(0).SortAscending = True
    End With
    
End Sub

Private Sub ClearPatiInfo()
'����:���������Ϣ��ʾ��
    Dim i As Long
    mlngPatiID = 0
    mlngPatiMainID = 0
    mlngDeptId = 0
    rptPati.Tag = ""
    If lblPlugIn.Visible Then lblPlugIn.Enabled = False
    
    For i = lblShow.LBound To lblShow.UBound
        lblShow(i).Caption = ""
    Next
    Set imgPatient.Picture = imgList.ListImages("Patient").Picture
    
    If Me.Visible Then
        mintPatiCount = 0
        rptPati.Columns(col_ѡ��).Icon = imgPati.ListImages("UnCheck").Index - 1
    End If
End Sub

Private Sub InitOutTime()
'���ܣ���ʼ����Ժ����
    cboOutTime.Clear
    With cboOutTime
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "15����"
        .ItemData(.NewIndex) = 15
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
        
        .ListIndex = 0
    End With
End Sub

Private Function InitDepts(Optional ByVal strIn As String) As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strDeptIDs As String, lngPreDept As Long
    
    cboDept.Clear
    On Error GoTo ErrH
    

    Set rsTmp = GetDataToDepts
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call cbo.SetIndex(cboDept.hWnd, 0)
        cboDept.Tag = cboDept.ItemData(0)
    End If
    
    InitDepts = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDataToDepts(Optional ByVal strIn As String = "") As ADODB.Recordset
'���ܣ���ȡ���Ҳ����б����ݼ�¼��
'������strIn ��������
    Dim strSQL As String
    Dim blnYN As Boolean
    Dim strLike As String
    
    If strIn <> "" Then blnYN = True
    strSQL = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where b.����id = a.Id And b.�������� = '�ٴ�' And" & vbNewLine & _
            "      b.������� In (2, 3) And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)" & vbNewLine & _
            IIf(blnYN, " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
            "Order By a.����"

       
    On Error GoTo ErrH
    If blnYN Then
        strLike = IIf(gstrMatchMethod = "0", "%", "")
        Set GetDataToDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strIn) & "%", strLike & UCase(strIn) & "%")
    Else
        Set GetDataToDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = sys.Readlob(glngSys, 27, lng����ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub FuncPrint()
'����:����ѡ����Ŀ��ӡ���Ӳ���
'   ��ҳ����,��ҳ����,��ҳ��ҳһ,��ҳ��ҳ��,סԺҽ��,���鱨��,��鱨��,סԺ����,������,�����¼,֪���ļ�,����֤��,�ٴ�·��
    Dim i As Long
    Dim strRegRange As String
    Dim strRange As String
    Dim strPrinterName As String
    
    'ͳ�Ʋ���¼��ӡ���
    For i = imgCHK.LBound To imgCHK.UBound
        If imgCHK(i).Tag = "T" Then
            strRegRange = strRegRange & "," & lblItem(i).Tag
            If InStr(",5,52,53,54,", "," & lblItem(i).Tag & ",") > 0 Then '��ҳ�������棬���Ͷ���5
                If InStr(strRange, "R5") = 0 Then 'û��
                    strRange = strRange & ",R5"
                End If
            ElseIf lblItem(i).Tag = 6 Then
                If InStr(strRange, "R6") = 0 Then 'û��
                    strRange = strRange & ",R6"
                End If
            Else
                strRange = strRange & ",R" & lblItem(i).Tag
            End If
        End If
    Next
    
    If strRange <> "" Then
        strRange = strRange & ","
        strRegRange = Mid(strRegRange, 2)
    Else
        MsgBox "��ѡ����Ҫ����ĵ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    strPrinterName = cboPrinterName.Text
    
    If strPrinterName = "" Then
        MsgBox "����ѡ������豸���ٽ��д�ӡ������", vbInformation, gstrSysName
        Exit Sub
    End If
    Call SetRegister(˽��ģ��, "��ӡ����", "��ӡ����", strRegRange)
    Call SetRegister(˽��ģ��, "��ӡ����", "��ӡ��", strPrinterName)
    Call PrintDocument(strRegRange, strRange, strPrinterName)
    
End Sub

Private Sub PrintDocument(ByVal strRegRange As String, ByVal strRange As String, ByVal strPrinterName As String)
    Dim i As Integer, lngNo As Long
    Dim clsPath As zlCISPath.clsDockPath, clsTendsNew As zl9TendFile.clsTendFile, objPacsDoc As Object
    Dim varParam As Variant, strReportNO As String, blnNewTends As Boolean, intSel As Integer, strEprName As String
    Dim lngInNo As Long, blnDataMove As Boolean, strName As String
    Dim lngSel As Long
    Dim lngPage As Long
    Dim strMsg As String
    Dim strReport As String
    Dim rsRet As New ADODB.Recordset
    
    On Error GoTo ErrHand

    '�������
    If mclsDockAduits Is Nothing Then
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
    End If
    Set clsPath = New zlCISPath.clsDockPath
    Set clsTendsNew = New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    '���ô�ӡ
    strReport = mstr�����Ӧ���� & ";" & mstr����Ӧ���� & ";" & IIf(mblnLIS, mstr���鱨���ӡ, "0")
    If strReport = ";" Then strReport = ""
    
    With rptPati
        If mintPatiCount > 1 Then
            For i = 0 To .Records.count - 1
                If .Records(i).Item(col_ѡ��).Checked = True Then
                    With .Records(i)
                        Call GetRsMedRec(CLng(.Item(col_����Id).Value), CLng(.Item(col_��ҳID).Value), CLng(.Item(col_��Ժ����ID).Value), rsRet, True)
                        Call gclsPackage.FuncPrintBatch(CLng(.Item(col_����Id).Value), CLng(.Item(col_��ҳID).Value), CLng(.Item(col_��Ժ����ID).Value), _
                            strRange, strRegRange, mclsDockAduits, clsPath, clsTendsNew, False, "", CStr(.Item(col_����).Value), CLng(.Item(col_סԺ��).Value), Me, _
                            lblInfo.Caption, False, strPrinterName, True, mstrPrintDocIDs, rsRet, lngPage, strReport, mobjRichEMR)
                        
                    End With
                End If
            Next
        ElseIf mlngPatiID <> 0 Then
            lngPage = 0: lngSel = FuncShowTipInfo()
            If lngSel = 0 Then
                MsgBox "��ӡʧ�ܣ���δ��ѡ�κ��ļ���", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
            Call gclsPackage.FuncPrintBatch(mlngPatiID, mlngPatiMainID, mlngDeptId, strRange, strRegRange, mclsDockAduits, _
                    clsPath, clsTendsNew, False, "", mstrPatiName, mlngInNO, Me, lblInfo.Caption, False, strPrinterName, True, mstrPrintDocIDs, mrsMedRec, lngPage, strReport, mobjRichEMR)
            strMsg = "��ѡ����" & lngSel & "���ļ���" & vbCrLf & " һ����ӡ�ˣ�" & lngPage & "�ݡ�"
            If strMsg <> "" Then MsgBox strMsg, vbInformation + vbOKOnly, gstrSysName
            
        End If
    End With
    lblInfo.Caption = ""
    Exit Sub
ErrHand:
    zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
    lblInfo.Caption = ""
    mstrPrintDocIDs = ""
End Sub

Private Function GetPrintLog(ByVal lngPatient As Long, ByVal lngPageID As Long) As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "Select ��ӡ���� As ��ӡ��, ��ӡ����, ��ӡ��, ��ӡʱ�� From ������ӡ��¼ Where ����id = [1] And ��ҳid = [2] Order By ��ӡʱ��, ��ӡ���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatient, lngPageID)
    Do Until rs.EOF
        GetPrintLog = GetPrintLog & vbCrLf & Rpad(rs!��ӡ��, 10) & Rpad(Format(rs!��ӡʱ��, "yyyy-mm-dd hh:MM"), 20) & Rpad(rs!��ӡ����, 40)
        rs.MoveNext
    Loop
    GetPrintLog = Rpad("��ӡ��", 10) & Rpad("��ӡʱ��", 20) & Rpad("��ӡ����", 40) & GetPrintLog
    
    Exit Function
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncPrintPreview()
'����:Ԥ��
'   ��ҳ����,��ҳ����,��ҳ��ҳһ,��ҳ��ҳ��,סԺҽ��,���鱨��,��鱨��,סԺ����,������,�����¼,֪���ļ�,����֤��,�ٴ�·����סԺ֤
    Dim i As Long
    Dim strRegRange As String
    Dim strRange As String
    Dim strPrinterName As String
    Dim lngTabIX As Long
    Dim strTabCaption As String
    
    '��λԤ�����
    lngTabIX = CLng(picItem(mbytType).Tag)
    If lngTabIX >= 0 Then
        If Not (tbcSub.Visible And tbcSub.Item(lngTabIX).Selected) And InStr(",���鱨��,��鱨��,סԺ����,������,�����¼,֪���ļ�,����֤��,סԺ֤,", lblItem.Item(mbytType).Caption) > 0 Then
            MsgBox "��δ��ѡ��" & lblItem(mbytType).Caption & "��������Ԥ����", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    strRegRange = strRegRange & "," & lblItem(mbytType).Tag
    If InStr(",5,52,53,54,", "," & lblItem(mbytType).Tag & ",") > 0 Then '��ҳ�������棬���Ͷ���5
        If InStr(strRange, "R5") = 0 Then 'û��
            strRange = strRange & ",R5"
        End If
    Else
        strRange = strRange & ",R" & lblItem(mbytType).Tag
        strTabCaption = lblItem(mbytType).Caption
    End If

    
    If strRange <> "" Then
        strRange = Replace(strRange, ",", "")

        strRegRange = Replace(strRegRange, ",", "")
    Else
        MsgBox "��ѡ����Ҫ����ĵ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strPrinterName = cboPrinterName.Text
    If strPrinterName = "" Then
        MsgBox "����ѡ������豸���ٽ��д�ӡ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call PrintDocumentView(strRegRange, strRange, strPrinterName, strTabCaption)
End Sub

Private Sub PrintDocumentView(ByVal strRegRange As String, ByVal strRange As String, ByVal strPrinterName As String, ByVal strTabCaption As String)
    Dim i As Integer, rs As New ADODB.Recordset, lngNo As Long
    Dim clsPath As zlCISPath.clsDockPath, clsTendsNew As zl9TendFile.clsTendFile, objPacsDoc As Object
    Dim varParam As Variant, strReportNO As String, blnNewTends As Boolean, intSel As Integer, strEprName As String
    Dim strName As String, strMsg As String
    Dim blnMod As Boolean
    
    On Error GoTo ErrHand

    '�������
    If mclsDockAduits Is Nothing Then
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
    End If
    Set clsPath = New zlCISPath.clsDockPath
    Set clsTendsNew = New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    '��ȡ��¼
    Set rs = mrsMedRec
    If InStr(",R5,R9,R1,", "," & strRange & ",") > 0 Then
        rs.Filter = "ID = '" & strRange & "'"
        If rs.RecordCount > 0 Then
            Select Case rs("ID").Value
            Case "R5"               '��ҳ
                Select Case mintMecStandard
                Case 0 '��������׼
                    If Have��������(mlngDeptId, "��ҽ��") Then
                        strReportNO = "ZL1_INSIDE_1261_4"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_1"
                    End If
                Case 1    '�Ĵ�ʡ��׼
                    If Have��������(mlngDeptId, "��ҽ��") Then
                        strReportNO = "ZL1_INSIDE_1261_6"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_5"
                    End If
                Case 2    '����ʡ��׼
                    If Have��������(mlngDeptId, "��ҽ��") Then
                        strReportNO = "ZL1_INSIDE_1261_8"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_7"
                    End If
                Case 3     '����ʡ��׼
                    If Have��������(mlngDeptId, "��ҽ��") Then
                        strReportNO = "ZL1_INSIDE_1261_10"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_9"
                    End If
                Case Else '�����޸�ʱδ����
                    If Have��������(mlngDeptId, "��ҽ��") Then
                        strReportNO = "ZL1_INSIDE_1261_4"
                    Else
                        strReportNO = "ZL1_INSIDE_1261_1"
                    End If
                End Select
          
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & strReportNO, "Printer", strPrinterName)
                If InStr("," & strRegRange & ",", ",5,") > 0 Then '����
                    Call ReportOpen(gcnOracle, ParamInfo.ϵͳ��, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, "ReportFormat=1", 1)
                End If
                
                If InStr("," & strRegRange & ",", ",52,") > 0 Then '����
                    Call ReportOpen(gcnOracle, ParamInfo.ϵͳ��, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, "ReportFormat=2", 1)
                End If
                
                If InStr("," & strRegRange & ",", ",53,") > 0 Then '��һ
                    Call ReportOpen(gcnOracle, ParamInfo.ϵͳ��, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, "ReportFormat=3", 1)
                End If
                
                If InStr("," & strRegRange & ",", ",54,") > 0 Then '����
                    Call ReportOpen(gcnOracle, ParamInfo.ϵͳ��, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, "ReportFormat=4", 1)
                End If
    
            Case "R1"               'ҽ��
                '�ȴ�ӡ����
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & "ZL1_INSIDE_1254_1", "Printer", strPrinterName)
                Call gobjKernel.zlPrintAdvice(Me, mlngPatiID, mlngPatiMainID, 0, 0, strPrinterName, 1)
                '�ٴ�ӡ����
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & "ZL1_INSIDE_1254_2", "Printer", strPrinterName)
                Call gobjKernel.zlPrintAdvice(Me, mlngPatiID, mlngPatiMainID, 0, 1, strPrinterName, 1)
                
            Case "R9"               '�ٴ�·��
                If Not CheckPatiPath() Then
                    MsgBox "��ǰ����û���ٴ�·����,����Ԥ����", vbOKOnly, Me.Caption
                    Exit Sub
                End If
                Call clsPath.zlRefreshReadOnly(mlngPatiID, mlngPatiMainID)
                Call clsPath.zlFuncPathTableOutPut(2, True, "", 0, 0, strPrinterName)  'Ԥ��
            End Select
        End If
    Else
        '����Ŀ
        rs.Filter = "�ϼ�id = '" & strRange & "'" & " And Ԥ��=1 "
        If rs.RecordCount > 0 Then
            If InStr(rs!ID, "R") = 0 And Len(rs!ID) >= 32 Then
                'EMR����Ԥ��
                If Not mobjRichEMR Is Nothing Then
                    If InStr(rs!����, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(rs!����, "|")(0), Split(rs!����, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(rs!����, "")
                    End If
                    Call mobjRichEMR.zlPrintDoc(True)
                End If
            Else
                varParam = Split(rs("����").Value, ";")
                Select Case rs("�ϼ�id").Value
                Case "R2"               'סԺ����
                    rs.Filter = "ID = '" & strRange & "'"
                    strEprName = Split(rs("����").Value, "��")(0)
                    Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                Case "R3"               '������
                    If InStr("," & mstrPrintDocIDs, "," & Val(varParam(0)) & ",") = 0 Then '����û���
                        strEprName = Split(rs("����").Value, "��")(0)
                        Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                    End If
                Case "R4"               '�����¼
                    blnNewTends = Get�°滤��(mlngPatiID, mlngPatiMainID)
                    If blnNewTends = False Then
                        varParam = Split(rs("����").Value, ";")
                        If UBound(varParam) >= 1 Then
                            If Val(varParam(1)) = -1 Then '���µ�
                                Call mclsDockAduits.zlRefreshTendBody(mlngPatiID, mlngPatiMainID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)))
                                Call mclsDockAduits.zlPrintDocument(1, 1, , strPrinterName)
                            Else '�����¼
                                Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), mlngPatiID, mlngPatiMainID, Val(Split(varParam(0), "_")(0)), CStr(varParam(2)), , Val(varParam(4)))
                                Call mclsDockAduits.zlPrintDocument(2, 1, , strPrinterName)
                            End If
                        End If
                    Else
                        varParam = Split(rs("����").Value, ";")
                        If UBound(varParam) >= 1 Then
                            Select Case Val(varParam(1))
                                Case -1 '���µ�
                                    intSel = 1
                                Case 1  '����ͼ
                                    intSel = 3
                                Case Else '��¼��
                                    intSel = 2
                            End Select
                            Call clsTendsNew.zlPrintDocument(mlngPatiID, mlngPatiMainID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), intSel, strPrinterName, False)
                        End If
                    End If
                Case "R6"               '�����鱨��
                    If NVL(rs!Eprid, "") = "E" And mstr�����Ӧ���� <> "" Then
                        strReportNO = Split(mstr�����Ӧ����, ",")(2)
                        varParam = Split(rs("����").Value, ";")  '�ڶ���������ҽ��ID
                        Call ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, "ҽ��ID=" & varParam(1), 1)
                    ElseIf NVL(rs!Eprid, "") = "D" And mstr����Ӧ���� <> "" Then
                        strReportNO = Split(mstr����Ӧ����, ",")(2)
                        varParam = Split(rs("����").Value, ";")  '�ڶ���������ҽ��ID
                        Call ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, "ҽ��ID=" & varParam(1), 1)
                    Else
                        strEprName = Split(rs("����").Value, "��")(0)
                        If UBound(Split(strEprName, ">")) > 0 Then
                            strEprName = Split(strEprName, ">")(1)
                        End If
                        blnMod = False
                        If NVL(rs!Eprid, "") = "E" And mblnLIS And Val(mstr���鱨���ӡ) = 1 Then
                            If InitObjLis(False) Then
                                blnMod = gobjLIS.PrintReport(Me, Val(varParam(1)), mlngPatiID, 1, strMsg)
                            End If
                        End If
                        If Not blnMod Then
                            If Val(varParam(3)) <> 0 Then
                                'RIS
                                If Not gobjXWHIS Is Nothing Then
                                    Call gobjXWHIS.ShowViewReport(Me.hWnd, Val(varParam(1)), True, Val(varParam(3)))
                                End If
                            ElseIf Val(varParam(0)) <> 0 Then
                                Call mclsDockAduits.zlPrintDocument(4, 1, Val(varParam(0)), strPrinterName)
                            Else
                                If objPacsDoc Is Nothing Then
                                    Set objPacsDoc = DynamicCreate("zlPublicPACS.clsPublicPacs", "�°�PACS�༭��", False)
                                    Call objPacsDoc.InitInterface(gcnOracle, gstrDBUser)
                                End If
                                Call objPacsDoc.PrintReport(varParam(2), strPrinterName, True) 'TrueԤ��
                            End If
                        End If
                    End If
                Case "R7"               '����֤��
                    strEprName = Split(rs("����").Value, "��")(0)
                    Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                Case "R8"               '֪���ļ�
                    strEprName = Split(rs("����").Value, "��")(0)
                    Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrinterName)
                Case "R10"     '���Ӳ�����ӡ סԺ֤
                    strEprName = "ZLCISBILL" & Format(rs!Eprid, "00000") & "-1"
                    If UBound(varParam) >= 1 Then
                        Call ReportOpen(gcnOracle, glngSys, strEprName, Me, "NO=" & varParam(0), "����=" & varParam(1), "ҽ��ID=0", 1)
                    End If
                Case "R11"  '��������
                    If UBound(varParam) >= 1 Then
                        strReportNO = varParam(2)
                        Call ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & mlngPatiID, "��ҳid=" & mlngPatiMainID, 1)
                    End If
                End Select
            End If
        Else
            Select Case strRange
            
            Case "R2"
                strMsg = "��ǰ����û��סԺ����������Ԥ����"
            Case "R3"
                strMsg = "��ǰ����û�л�����������Ԥ����"
            Case "R4"
                strMsg = "��ǰ����û�л����¼������Ԥ����"
            Case "R6"
                strMsg = "��ǰ����û��" & strTabCaption & "������Ԥ����"
            Case "R7"
                strMsg = "��ǰ����û�м���֤��������Ԥ����"
            Case "R8"
                strMsg = "��ǰ����û��֪���ļ�������Ԥ����"
            Case "R10"
                strMsg = "��ǰ����û��סԺ֤������Ԥ����"
            Case "R11"
                strMsg = "��ǰ����û��������������Ԥ����"
            End Select
            If strMsg <> "" Then
                MsgBox strMsg, vbOKOnly, Me.Caption
            End If
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ReturnItemTag(ByVal strName As String) As Integer
'����:�����ض����
'1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-�����鱨��;7-����֤��;8-֪���ļ�;9-�ٴ�·��;10-סԺ֤;11-��������
    Dim intRet As Integer
    
    Select Case strName
    
    Case "סԺҽ��"
        intRet = 1
    Case "סԺ����"
        intRet = 2
    Case "������"
        intRet = 3
    Case "�����¼"
        intRet = 4
    Case "��ҳ����"
        intRet = 5
    Case "��ҳ����"
        intRet = 52
    Case "��ҳ��ҳһ"
        intRet = 53
    Case "��ҳ��ҳ��"
        intRet = 54
    Case "���鱨��", "��鱨��"
        intRet = 6
    Case "����֤��"
        intRet = 7
    Case "֪���ļ�"
        intRet = 8
    Case "�ٴ�·��"
        intRet = 9
    Case "סԺ֤"
        intRet = 10
    Case "��������"
        intRet = 11
    End Select
    ReturnItemTag = intRet
End Function

Private Function GetTabIndex(ByVal strName As String) As Integer
    Dim intRet As Integer
    
    Select Case strName
    
    Case "סԺ����"
        intRet = tab_סԺ����
    Case "������"
        intRet = tab_������
    Case "�����¼"
        intRet = tab_�����¼
    Case "���鱨��"
        intRet = tab_���鱨��
    Case "��鱨��"
        intRet = tab_��鱨��
    Case "����֤��"
        intRet = tab_����֤��
    Case "֪���ļ�"
        intRet = tab_֪���ļ�
    Case "סԺ֤"
        intRet = tab_סԺ֤
    Case "��������"
        intRet = tab_��������
    Case Else
        intRet = -1
    End Select
    GetTabIndex = intRet
End Function
Private Function CheckPatiPath() As Boolean
'����:��鵱ǰ�����Ƿ�����ٴ�·����(״̬=2-��������;3-�������)
    Dim strSQL As String
    Dim rsPath As ADODB.Recordset
    
    On Error GoTo ErrH
    strSQL = "Select Count(1) as ��¼�� From �����ٴ�·�� A Where a.����id = [1] And a.��ҳid = [2] And a.״̬ In (2, 3)"
    Set rsPath = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, mlngPatiMainID)
    CheckPatiPath = Val(rsPath!��¼�� & "") > 0
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectItems(ByVal bytFunc As Byte)
'����:
'   bytFunc=1 ȫѡ,=2ȡ��ȫѡ
    Dim i As Long
    
    With rptPati
        For i = 0 To .Records.count - 1
            If bytFunc = 1 Then
                .Records(i).Item(col_ѡ��).Checked = True
            Else
                .Records.Record(i).Item(0).Checked = False
            End If
        Next
        mintPatiCount = IIf(bytFunc = 1, .Records.count, 0)
    End With
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call LoadItemInfo
    If tbcSub.Visible Then Call FuncSetFocus(Val(tbcSub.Selected.Tag))
End Sub

Private Sub vsc_Change()
    fraIn.Top = 0 - (fraIn.Height - fraCent.Height) * (vsc.Value / vsc.Max)
    'ת�ƽ���
    picCenter.SetFocus
End Sub

Private Sub vsItemInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsItemInfo
        If Col = 0 Then
            If mrsMedRec Is Nothing Then Exit Sub
            mrsMedRec.Filter = "ID='" & .Cell(flexcpData, Row, 1) & "'"
            If mrsMedRec.RecordCount > 0 Then
                mrsMedRec.MoveFirst
                mrsMedRec!�Ƿ�ѡ�� = IIf(.Cell(flexcpChecked, Row, Col) = 1, 1, 0)
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, Col) = flexUnchecked Then Exit For
                Next
                If i = .Rows Then
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("Check").Picture
                    .Cell(flexcpData, 0, 0) = 1
                Else
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                    .Cell(flexcpData, 0, 0) = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItemInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub tmrTime_Timer()
    Dim vPos As PointAPI
    Dim vRect As RECT
    
    stbThis.Panels(2).Text = IIf(mintPatiCount = 0, "", "��ѡ��" & mintPatiCount & "�����ˣ�")
    cmdPrint.Enabled = mlngPatiID <> 0
    cmdPreView.Enabled = mlngPatiID <> 0
    
    If mblnTag = False Then Exit Sub
    
    GetCursorPos vPos
    GetWindowRect picCenter.hWnd, vRect
    If Not (Between(vPos.x, vRect.Left, vRect.Right) And Between(vPos.y, vRect.Top, vRect.Bottom)) Then
        Call SetPicItemBG(3)
    End If
End Sub

Private Sub vsItemInfo_Click()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    If mbytType <> CByte(tbcSub.Selected.Tag) Then Call FuncSetFocus(CLng(tbcSub.Selected.Tag))
    
    If mrsMedRec Is Nothing Then Exit Sub
    With vsItemInfo
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow = 0 And lngCol = 0 Then
            If .TextMatrix(.Rows - 1, 1) <> "" Then
                If Val(.Cell(flexcpData, 0, 0) & "") = 0 Then
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("Check").Picture
                    .Cell(flexcpData, 0, 0) = 1
                Else
                    Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
                    .Cell(flexcpData, 0, 0) = 0
                End If
                
                For i = .FixedRows To .Rows - 1
                    .Cell(flexcpChecked, i, 0) = IIf(Val(.Cell(flexcpData, 0, 0) & "") = 0, flexUnchecked, flexChecked)
                    mrsMedRec.Filter = "ID='" & .Cell(flexcpData, i, 1) & "'"
                    If mrsMedRec.RecordCount > 0 Then
                        mrsMedRec!�Ƿ�ѡ�� = IIf(.Cell(flexcpChecked, i, lngCol) = 1, 1, 0)
                    End If
                Next
            End If
        ElseIf lngRow >= 1 And lngRow <= .Rows - 1 Then
            mrsMedRec.Filter = ""
            Do While Not mrsMedRec.EOF
                mrsMedRec!Ԥ�� = 0 '�������Ԥ����
                mrsMedRec.MoveNext
            Loop
            mrsMedRec.Filter = "ID='" & .Cell(flexcpData, lngRow, 1) & "'"
            If mrsMedRec.RecordCount > 0 Then mrsMedRec!Ԥ�� = 1
        End If
    End With
End Sub

Private Sub InitVSTable()

    With vsItemInfo
        .Cols = 3: .ColWidth(0) = 300
        .ColWidth(1) = 3000: .ColWidth(2) = 7500
        .FixedAlignment(1) = flexAlignCenterCenter
        .RowHeightMin = 300
        .Editable = flexEDKbd
        Set .Cell(flexcpPicture, 0, 0) = imgPati.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, 0) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(0, 1) = "�빴ѡ��Ҫ��ӡ�����ݡ�": .ColDataType(0) = flexDTBoolean
    End With
End Sub

Private Sub FuncSetFocus(ByVal bytIndex As Byte)
    Dim i As Long
    
    If Not mblnLoad Then Exit Sub
    
    For i = Lin.LBound To Lin.UBound
        Set Lin(i).Container = picItem(bytIndex)
        Lin(i).BorderColor = &HFF0000
        Lin(i).Visible = True
    Next
     
    With Lin(0)
        .X1 = 0: .X2 = picItem(i).Width
        .Y1 = 0: .Y2 = 0
    End With
    With Lin(1)
        .X1 = picItem(i).Width - 15: .X2 = picItem(i).Width - 15
        .Y1 = 0: .Y2 = picItem(i).Height
    End With
    With Lin(2)
        .X1 = 0: .X2 = picItem(i).Width - 15
        .Y1 = picItem(i).Height - 15: .Y2 = picItem(i).Height - 15
    End With
    With Lin(3)
        .X1 = 0: .X2 = 0
        .Y1 = 0: .Y2 = picItem(i).Height
    End With
    mbytType = bytIndex
End Sub

Private Function FuncShowTipInfo() As Long
    Dim lngCount As Long
    Dim i As Long
    Dim strRange As String
    Dim lngAll As Long
    Dim lngSub As Long
    Dim strMsg As String
    
    If mrsMedRec Is Nothing Then Exit Function
    If mrsMedRec.RecordCount = 0 Then Exit Function

    'ͳ�ƹ�ѡ�ļ����� ��ҳ������򸽼���һ��,ҽ���嵥������\��ʱ����һ��,�ٴ�·������һ��
    For i = imgCHK.LBound To imgCHK.UBound
        If imgCHK(i).Tag = "T" Then
            If InStr(",5,52,53,54,", "," & lblItem(i).Tag & ",") > 0 Then '��ҳ�������棬���Ͷ���5
                If InStr(strRange, "R5") = 0 Then 'û��
                    strRange = strRange & ",R5"
                    lngCount = lngCount + 1
                End If
            ElseIf InStr(",1,9,", "," & lblItem(i).Tag & ",") > 0 Then
                lngCount = lngCount + 1
            Else
                strRange = strRange & ",R" & lblItem(i).Tag
            End If
        End If
    Next
    'ͳ�Ƶ����ļ���Ŀ
    mrsMedRec.Filter = "�Ƿ�ѡ�� =1"
    lngAll = mrsMedRec.RecordCount
    mrsMedRec.Filter = "�ϼ�ID=Null And �Ƿ�ѡ�� =1"
    lngSub = mrsMedRec.RecordCount
    lngCount = lngCount + (lngAll - lngSub)
    
    FuncShowTipInfo = lngCount
    
End Function

Private Sub FuncLoadReport()
    Dim objControl As CommandBarControl
    Dim objPop As Object
    Dim strHide As String
    Dim i As Long
    
    strHide = ",ZL1_INSIDE_1254_1,ZL1_INSIDE_1254_2,ZL1_INSIDE_1261_1,ZL1_INSIDE_1261_4,ZL1_INSIDE_1261_5,ZL1_INSIDE_1261_6,ZL1_INSIDE_1261_7,ZL1_INSIDE_1261_8,ZL1_INSIDE_1261_9,ZL1_INSIDE_1261_10,"
    mstr����Ӧ���� = zlDatabase.GetPara("����Ӧ����", glngSys, p���Ӳ�����ӡ)
    mstr�����Ӧ���� = zlDatabase.GetPara("�����Ӧ����", glngSys, p���Ӳ�����ӡ)
    If mstr����Ӧ���� <> "" Then strHide = strHide & "," & Split(mstr����Ӧ����, ",")(2) & ","
    If mstr�����Ӧ���� <> "" Then strHide = strHide & "," & Split(mstr�����Ӧ����, ",")(2) & ","
    mstr���鱨���ӡ = zlDatabase.GetPara("���鱨���ӡ", glngSys, p���Ӳ�����ӡ)
    '��ջ���
    Set mcolReport = New Collection
    For i = 1 To cbsMain.ActiveMenuBar.Controls.count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "����*" Then
                cbsMain.ActiveMenuBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    Call zlDatabase.ShowReportMenu(cbsMain, glngSys, p���Ӳ�����ӡ, mstrPrivs, strHide)
    
    For i = 1 To cbsMain.ActiveMenuBar.Controls.count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "����*" Then
            Set objControl = cbsMain.ActiveMenuBar.Controls.Item(i)
            Exit For
        End If
    Next
    
    If Not objControl Is Nothing Then
        With objControl.CommandBar.Controls
            For i = 1 To .count
                Set objPop = .Item(i)
                mcolReport.Add Split(objPop.Caption, "(&")(0) & "," & objPop.Parameter, "_" & i     '��������,ϵͳ��,������
            Next
        End With
    End If
End Sub

Public Sub DefCommandPlugInPopup(ByRef objControls As CommandBarControls)
'���ܣ���չ���ܵ����˵�
    Dim strFunc As String, strTmp As String
    Dim arrTmp As Variant
    Dim objControl As CommandBarControl
    
    Dim i As Long
    
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, p���Ӳ�����ӡ)
    Call zlPlugInErrH(Err, "GetFuncNames")
    Err.Clear: On Error GoTo 0
    If strFunc <> "" Then
        arrTmp = Split(strFunc, ",")
        strTmp = Replace(strFunc, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        If objControls.count = 0 Then
            For i = 0 To UBound(arrTmp)
                Set objControl = objControls.Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                objControl.IconId = conMenu_Tool_PlugIn_Item
                objControl.Parameter = arrTmp(i)
            Next
        End If
        lblPlugIn.Visible = True
        lblPlugIn.Enabled = False
    End If
End Sub


