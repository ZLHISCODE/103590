VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmTechnicStation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "ҽ������վ"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmTechnicStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleMode       =   0  'User
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   615
      ScaleHeight     =   720
      ScaleWidth      =   1350
      TabIndex        =   32
      Top             =   6765
      Visible         =   0   'False
      Width           =   1350
      Begin XtremeReportControl.ReportControl rptNotify 
         Height          =   540
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   675
         _Version        =   589884
         _ExtentX        =   1191
         _ExtentY        =   952
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6120
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   27
      Top             =   4320
      Width           =   855
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   4500
      Left            =   3705
      TabIndex        =   1
      Top             =   2865
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   7937
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   7485
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTechnicStation.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "������ɫ"
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
   Begin VB.Frame fraUD_S 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3720
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   2730
      Width           =   3255
   End
   Begin VB.PictureBox picExec 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2040
      Left            =   3720
      ScaleHeight     =   2040
      ScaleWidth      =   7755
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   7755
      Begin VB.PictureBox picBlood 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   75
         ScaleHeight     =   390
         ScaleWidth      =   1980
         TabIndex        =   34
         Top             =   1155
         Width           =   1980
         Begin VB.Timer timBRefresh 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   270
            Top             =   0
         End
         Begin XtremeDockingPane.DockingPane DkpBlood 
            Left            =   0
            Top             =   0
            _Version        =   589884
            _ExtentX        =   450
            _ExtentY        =   423
            _StockProps     =   0
         End
      End
      Begin VB.PictureBox picApplyUD_S 
         Height          =   855
         Left            =   6650
         MousePointer    =   9  'Size W E
         ScaleHeight     =   855
         ScaleWidth      =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox picApplyInfo 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   6700
         ScaleHeight     =   855
         ScaleWidth      =   1005
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1125
         Visible         =   0   'False
         Width           =   1000
         Begin RichTextLib.RichTextBox rtfAppend 
            Height          =   1395
            Left            =   0
            TabIndex        =   24
            Top             =   240
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   2461
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmTechnicStation.frx":0E1C
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
         Begin VB.Label lblApply 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���븽��"
            Height          =   180
            Left            =   45
            TabIndex        =   25
            Top             =   30
            Width           =   720
         End
      End
      Begin VB.Frame fraDiag 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   60
         TabIndex        =   19
         Top             =   15
         Width           =   7605
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ϣ�"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   45
            Width           =   900
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   1005
            TabIndex        =   20
            Top             =   45
            Width           =   90
         End
      End
      Begin VB.Frame fraExec 
         Height          =   795
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   7380
         Begin VB.Label lblRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00008000&
            Height          =   465
            Left            =   6600
            TabIndex        =   28
            Top             =   330
            Width           =   495
         End
         Begin VB.Label lblAdvice 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   540
            Left            =   120
            TabIndex        =   11
            Top             =   195
            Width           =   6480
         End
         Begin VB.Label lblCash 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   6825
            TabIndex        =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsExec 
         Height          =   885
         Left            =   60
         TabIndex        =   12
         Top             =   1125
         Width           =   6645
         _cx             =   11721
         _cy             =   1561
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
         BackColorSel    =   16772055
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
         MergeCells      =   2
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
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6690
      Left            =   120
      ScaleHeight     =   6690
      ScaleWidth      =   3615
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   675
      Width           =   3615
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   4515
         Left            =   45
         TabIndex        =   0
         Top             =   1230
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   7964
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chkFilter 
         Height          =   255
         Left            =   3120
         Picture         =   "frmTechnicStation.frx":0EB9
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "���ղ��������Բ��˽��й�����ʾ"
         Top             =   840
         Width           =   270
      End
      Begin VB.Timer timRefresh 
         Interval        =   1000
         Left            =   270
         Top             =   1155
      End
      Begin VB.Frame fraFilter 
         Caption         =   "ִ��״̬"
         Height          =   1125
         Left            =   60
         TabIndex        =   14
         Top             =   5505
         Width           =   3480
         Begin VB.CheckBox chkִ��״̬ 
            Caption         =   "����ִ���а����Ѿ��˶�(&5)"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   26
            Top             =   846
            Width           =   2565
         End
         Begin VB.CheckBox chkִ��״̬ 
            Caption         =   "�Ѿ�ִ��(&4)"
            Height          =   195
            Index           =   3
            Left            =   1980
            TabIndex        =   18
            Top             =   555
            Width           =   1290
         End
         Begin VB.CheckBox chkִ��״̬ 
            Caption         =   "��δִ��(&2)"
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   16
            Top             =   300
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox chkִ��״̬ 
            Caption         =   "�ܾ�ִ��(&1)"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   300
            Width           =   1290
         End
         Begin VB.CheckBox chkִ��״̬ 
            Caption         =   "����ִ��(&3)"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   17
            Top             =   555
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   4
            Left            =   105
            Picture         =   "frmTechnicStation.frx":770B
            Top             =   846
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   1755
            Picture         =   "frmTechnicStation.frx":7C95
            Top             =   555
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   105
            Picture         =   "frmTechnicStation.frx":821F
            Top             =   555
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   1755
            Picture         =   "frmTechnicStation.frx":87A9
            Top             =   300
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   105
            Picture         =   "frmTechnicStation.frx":8D33
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1170
         TabIndex        =   5
         Top             =   495
         Width           =   2265
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   345
         Top             =   1320
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
               Picture         =   "frmTechnicStation.frx":92BD
               Key             =   "δִ��"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":9857
               Key             =   "��ִ��"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":9DF1
               Key             =   "�ܾ�ִ��"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":A38B
               Key             =   "����ִ��"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":A925
               Key             =   "�ѱ���"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":AEBF
               Key             =   "CheckCol"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":B459
               Key             =   "Path"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":B9F3
               Key             =   "�Ѻ˶�"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1170
         TabIndex        =   3
         Text            =   "cboDept"
         Top             =   120
         Width           =   2265
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   825
         TabIndex        =   29
         Top             =   840
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmTechnicStation.frx":BF8D
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "���￨"
         IDKindWidth     =   555
         FindPatiShowName=   0   'False
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFind 
         Caption         =   "����(F3)"
         Height          =   255
         Left            =   75
         TabIndex        =   30
         Top             =   870
         Width           =   975
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˲���(&U)"
         Height          =   180
         Left            =   135
         TabIndex        =   4
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������(&D)"
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   1185
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C054
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C0B2
            Key             =   "Meet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C110
            Key             =   "MeetFinish"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C16E
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C1CC
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C22A
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C288
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C2E6
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C344
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C3A2
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C400
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C45E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C4BC
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C51A
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C578
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C5D6
            Key             =   "������"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C634
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C692
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C6F0
            Key             =   "������"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C74E
            Key             =   "Out"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmTechnicStation.frx":C7AC
      Left            =   705
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTechnicStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99
Private Enum PATIREPORT_COLUMN
    col_�ۺ�״̬ = 0
    col_ѡ�� = 1
    col_·�� = 2
    col_ִ��״̬ = 3
    col_ͼ�� = 4
    col_��Դ = 5    '�������۲��ˣ�value='סԺ',caption='����'
    col_���ݺ� = 6
    col_���� = 7
    col_���� = 8
    col_���� = 9
    col_���� = 10
    col_���� = 11
    col_��ʶ�� = 12
    col_���� = 13
    col_�ѱ� = 14
    col_Ҫ��ʱ�� = 15
    col_����ʱ�� = 16
    col_ִ�м� = 17
    col_�Ա� = 18
    col_���� = 19
    col_����� = 20
    col_���ʱ�� = 21
    '������
    col_ִ�п��� = 22
    col_����Id = 23
    col_��ҳID = 24
    col_�Һŵ� = 25
    col_�Һ�ID = 26 '
    col_Ӥ�� = 27 '
    col_���￨�� = 28
    col_���֤�� = 29
    col_IC���� = 30
    col_ҽ���� = 31
    col_����id = 32
    col_��Ժ���� = 33
    COL_״̬ = 34
    col_ҽ��ID = 35
    col_���ID = 36
    col_���ͺ� = 37
    COL_������� = 38
    col_ִ�й��� = 39
    col_ִ�а��� = 40
    col_��¼���� = 41
    COL_����ת�� = 42
    col_�ļ�ID = 43
    col_������ = 44
    col_����ID = 45
    col_�������� = 46
    col_������� = 47
    col_������ = 48
    col_��� = 49
    col_���� = 50
    COL_�������� = 51
    COL_�˶��� = 52
    col_��˱�־ = 53
    COL_����ʱ�� = 54
    col_����ģʽ = 55
    COL_������ĿID = 56
    col_��Ч = 57
    COL_ִ�з��� = 58
    COL_��ҳ�Һ�ID = 59  '������ҳ.�Һ�ID
    COL_���ӱ�־ = 60    '����ҽ����¼.���ӱ�־
End Enum

Private Enum NOTIFYREPORT_COLUMN
    c_ͼ�� = 0
    C_����ID = 1
    C_��ҳID = 2
    
    c_���� = 3
    C_״̬ = 4
    
    C_��Ϣ = 5
    C_��� = 6
    C_���� = 7
    C_ҵ�� = 8
    
End Enum

Private mblnShowBed As Boolean

Private Enum PATI_TYPE
    pt�ҵ� = 1
    pt��Ժ = 2
    ptԤ�� = 3
    pt��Ժ = 4
    pt���� = 5
    pt���� = 6
    pt���ת�� = 7
End Enum

Private Enum Msg_Type '��Ϣ�������
    m�������� = 1
    m������ = 2
    mѪ������ = 3
End Enum

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
Private WithEvents mclsInAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsInAdvices.VB_VarHelpID = -1
Private WithEvents mclsOutAdvices As zlPublicAdvice.clsDockOutAdvices
Attribute mclsOutAdvices.VB_VarHelpID = -1
Private WithEvents mclsExpenses As zlPublicExpense.clsDockExpense
Attribute mclsExpenses.VB_VarHelpID = -1
Private WithEvents mclsInEPRs As zlRichEPR.cDockInEPRs
Attribute mclsInEPRs.VB_VarHelpID = -1
Private WithEvents mclsOutEPRs As zlRichEPR.cDockOutEPRs
Attribute mclsOutEPRs.VB_VarHelpID = -1
Private mclsTendEPRs As zlRichEPR.cDockInTendEPRs
Attribute mclsTendEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zlRichEPR.cDockInTends
Attribute mclsTends.VB_VarHelpID = -1
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '�°滤���¼
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mcolSubForm As Collection
Private mfrmActive As Form
Private mobjFrmBloodExe As Object
Private WithEvents mobjIDCard As clsIDCard '���֤����
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object 'IC������
Private mobjAppendBill As Object '������صĶ���
Private mbln���Ѱ�ť As Boolean '����¸��Ѵ��ڵ������ܰ�ť
Private mstr���� As String

'�����������
Private WithEvents mclsEPRReport As zlRichEPR.cEPRDocument
Attribute mclsEPRReport.VB_VarHelpID = -1

'ҽ�ƿ�
Private mobjSquareCard As Object      '���������
Private mstrCardKind As String        '��������󷵻صĿ��õ�ҽ�ƿ�
Private Enum CardProperty
    CP���� = 0
    CPȫ�� = 1
    CP�ɶ��� = 2
    CP�����ID = 3
    CP���ų��� = 4
    CPȱʡ��� = 5
    CP�����ʻ� = 6
    CP����������ʾ = 7
End Enum

Private mintFindType As Integer '1-���￨,2-��ʶ�ţ�����ţ�,3-���ݺ�,4-����,5-�������֤���֤,6-IC��,7-ҽ����
Private mstrFindType As String '�洢��ǰ������������
Private mblnFindTypeEnabled As Boolean
Private mblnFilter As Boolean '�����Ƿ��Թ���ģʽ������ʾ
Private mblnFilterEnabled As Boolean

'������������
Private Type FilterCond
    Begin As Date
    End As Date
    NO As String
    ����ID As Long
    ��Դ As String
    ���� As Boolean
    ��Ч As Integer
    ��ʶ�� As String
    ���￨ As String
    ���� As String
    ���֤ As String
    IC���� As String
    ҽ���� As String
    ������ As String
    ����ID As Long
End Type
Private mvarCond As FilterCond
Private mblnֻ�����շ� As Boolean
Private mblnֻ�����շ�Enabled As Boolean
Private mbln����ִ�� As Boolean
Private mstr״̬ As String  '1-5λ�ֱ��ʾ���ܾ�ִ��,��δִ��,����ִ��,�Ѿ�ִ��,����ִ�����Ѿ��˶�
Private mstr������ As String
Private mstr�������� As String

'���ز�������
Private mblnExeLog As Boolean
Private mblnƤ����֤ As Boolean
Private mintRefresh As Integer
Private mstrRoom As String
Private mstr������� As String
Private mstr������� As String

'�����������
Private mstrPrivs As String
Private mlngModul As Long
Private mlngDept As Long
Private mstrDeptNode As String '��ǰҽ������������վ��
Private mstrPrePati As String
Private mlng����ID As Long, mlng��ҳID As Long, mstr�Һŵ� As String

Private mblnFirstLoad As Boolean '�ж��Ƿ��ǵ�һ�μ���
Private mblnMoved As Boolean
Private mlngFontSize As Long  '�����С
Private mblnReturn As Boolean      'cboDept/cboUnit�Ļس�����
Private mblnѪ͸�� As Boolean
Private mbln���� As Boolean

Private mbyt������˷�ʽ As Byte '49501:������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
Private mstrNotify As String '���ѵ�����
Private mintDay As Integer '���Ѷ������ڵ���Ϣ
Private mintMin As Integer '�����Զ�ˢ�¼��(����)
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean
Private mblnδ�շ���� As Boolean '������δ�շ����
Private mstrBloodControlIDs  As String '��Ѫִ�в˵�ID��

'����
Private mlngType As Long
Private mlngState As TYPE_PATI_State
Private mlngInIndex As Long
Private mlngOutIndex As Long
Private mlngNurIndex As Long
Private mlngNewNurIndex As Long
Private mlngNurEMRIndex As Long
Private mlngNewIndex As Long '�°没��ѡ������δ���Ϊ ��1
Private mblnNewNurRecord As Boolean     'Ѫ͸���Ƿ�ʹ���°滤���¼

Private mblnIsInit As Boolean

Private COLExec As New Collection

Private mbytSize As Byte '�����С 0-С���壨9�ţ���1-�����壨12�ţ�
Private mblnTabTmp As Boolean
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mobjKernel As zlPublicAdvice.clsPublicAdvice         '�ٴ����Ĳ���
Private mclsPExp As zlPublicExpense.clsPublicExpense

Private Function ExchangeAdvice(ByVal blnClinic As Boolean) As Boolean
'���ܣ��л���ʾ����/סԺҽ��ҳ��
'������blnClinic=�Ƿ���ʾ����ҽ��ҳ
'���أ��Ƿ�������л�ѡ��
    Dim blnSel As Boolean
    Dim blnOld As Boolean, intIdx As Integer
    
    If Not tbcSub.Selected Is Nothing Then
        blnSel = tbcSub.Selected.Tag Like "*ҽ��"
    End If
    
    For intIdx = 0 To tbcSub.ItemCount - 1
        If tbcSub(intIdx).Tag = "����ҽ��" Then
            If tbcSub(intIdx).Visible <> blnClinic Then
                tbcSub(intIdx).Visible = blnClinic
                If blnSel And blnClinic Then
                    tbcSub(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        ElseIf tbcSub(intIdx).Tag = "סԺҽ��" Then
            If tbcSub(intIdx).Visible <> Not blnClinic Then
                tbcSub(intIdx).Visible = Not blnClinic
                If blnSel And Not blnClinic Then
                    tbcSub(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        End If
    Next
End Function

Private Sub cboDept_GotFocus()
    Call zlControl.TxtSelAll(cboDept)
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex
    mblnReturn = False
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        '��Դ���Ų�����վ��
        If cboDept.Text <> "" Then
            strSQL = "Select A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� In(1,2,3) And B.�������� IN('���','����','����','����','Ӫ��')" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is Null)" & _
                IIf(mstrDeptNode <> "", " And (A.վ�� = [3] Or A.վ�� is Null)", "") & _
                " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                " Order by A.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(cboDept.Text) & "%", gstrLike & UCase(cboDept.Text) & "%", mstrDeptNode)
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cboDept, rsTmp!ID)
            Else
                cboDept.ListIndex = Val(cboDept.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboUnit_Click()
'���ܣ����¶�ȡ����
    If cboUnit.ListIndex = -1 Then Exit Sub
    mblnReturn = True
    
    If Val(cboUnit.Tag) = cboUnit.ListIndex Then Exit Sub
    cboUnit.Tag = cboUnit.ListIndex
 
    Call LoadPatients
End Sub

Private Sub cboUnit_GotFocus()
    Call zlControl.TxtSelAll(cboUnit)
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboUnit.hwnd, Val(cboUnit.Tag))
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    mblnReturn = False
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        '��Դ���Ų�����վ��
        If cboUnit.Text <> "" Then
            strSQL = "Select A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� In(1,2,3) And B.��������='����'" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is Null)" & _
                IIf(mstrDeptNode <> "", " And (A.վ�� = [3] Or A.վ�� is Null)", "") & _
                " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                " Order by A.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(cboUnit.Text) & "%", gstrLike & UCase(cboUnit.Text) & "%", mstrDeptNode)
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cboUnit, rsTmp!ID)
            Else
                cboUnit.ListIndex = Val(cboUnit.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkFilter_Click()
    mblnFilter = chkFilter.Value = 1
    PatiIdentify.Text = ""
    If PatiIdentify.Visible And PatiIdentify.Enabled Then PatiIdentify.SetFocus
    
    '�л�ʱ�����������ˢ���嵥
    Call ClearPatiCond
    Call LoadPatients
End Sub

Private Sub chkִ��״̬_Click(Index As Integer)
    If Visible Then
        mstr״̬ = zlStr.SetBit(mstr״̬, Index + 1, chkִ��״̬(Index).Value)
        chkִ��״̬(4).Enabled = chkִ��״̬(2).Value = 1
        Call LoadPatients
    End If
End Sub

Private Sub DkpBlood_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        If Not mobjFrmBloodExe Is Nothing Then
            Item.Handle = mobjFrmBloodExe.hwnd
        End If
    End If
End Sub

Private Sub Form_Activate()
    mblnFirstLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '����
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer, i As Long
    Dim strKey As String, intType As Integer
    Dim objControl As CommandBarControl
    Dim arrTmp As Variant, strTmp As String
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, pҽ������վ, GetInsidePrivs(pҽ������վ))
    Call AddMipModule(mclsMipModule)
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    On Error Resume Next
    Set mobjAppendBill = CreateObject("ZlSoft.HIS.Charge.AppendCharge")
    err.Clear: On Error GoTo 0
    mbln���Ѱ�ť = False
    mstr���� = ""
    Set mclsPExp = New zlPublicExpense.clsPublicExpense
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    mblnFirstLoad = True
    picApplyUD_S.Left = 9200
    
    mstrBloodControlIDs = Join(Array(conMenu_Manage_Complete, conMenu_Manage_Undone, conMenu_Manage_ThingAdd, conMenu_Manage_ThingModi, conMenu_Manage_ThingDel, _
        conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit, conMenu_Manage_ThingAudit * 100# + 1, conMenu_Manage_ThingDelAudit * 100# + 1), ",")

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    Call InitExecTable
        
    mstrRoom = zlDatabase.GetPara("ִ�м䷶Χ", glngSys, pҽ������վ)
    '��������
    mbytSize = zlDatabase.GetPara("����", glngSys, pҽ������վ, "0")
    
    '����ˢ������
    mstrNotify = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pҽ������վ, "000")
    mintDay = Val(zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pҽ������վ, 1))
    mintMin = Val(zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, pҽ������վ))
    mblnδ�շ���� = (Val(zlDatabase.GetPara("δ�շ����", glngSys, pҽ������վ)) = 1)
    mbln��Ϣ���� = Val(zlDatabase.GetPara("����������ʾ", glngSys, pҽ������վ)) = 1
    'Ѫ͸���Ƿ�ʹ���°滤���¼
    mblnNewNurRecord = (Val(zlDatabase.GetPara("Ѫ͸����д�°滤���¼", glngSys, pҽ������վ)) = 1)

    '����������ʼ
    '-----------------------------------------------------
    mvarCond.���� = Val(zlDatabase.GetPara("ֻ��ʾ����סԺ��Ŀ", glngSys, pҽ������վ, "1")) = 1
    mstr�������� = IIf(Val(zlDatabase.GetPara("���˹��˷�ʽ", glngSys, pҽ������վ)) = 1, "����ʱ��", "�״�ʱ��")
    
    '������Դ
    strKey = zlDatabase.GetPara("������Դ", glngSys, pҽ������վ, "111")
    mvarCond.��Դ = ""
    If Not (Val(Mid(strKey, 1, 1)) = 1 And Val(Mid(strKey, 2, 1)) = 1 And Val(Mid(strKey, 3, 1)) = 1) Then
        If Val(Mid(strKey, 1, 1)) = 1 Then mvarCond.��Դ = mvarCond.��Դ & ",1"
        If Val(Mid(strKey, 2, 1)) = 1 Then mvarCond.��Դ = mvarCond.��Դ & ",2"
        If Val(Mid(strKey, 3, 1)) = 1 Then mvarCond.��Դ = mvarCond.��Դ & ",4"
        mvarCond.��Դ = Mid(mvarCond.��Դ & ",3", 2)
    End If
    Call SetUnitVisible

    'ҽ����Ч
    strKey = zlDatabase.GetPara("ҽ����Ч", glngSys, pҽ������վ, "11")
    mvarCond.��Ч = 0
    If Not (Val(Mid(strKey, 1, 1)) = 1 And Val(Mid(strKey, 2, 1)) = 1) Then
        If Val(Mid(strKey, 1, 1)) = 1 Then
            mvarCond.��Ч = 1
        ElseIf Val(Mid(strKey, 2, 1)) = 1 Then
            mvarCond.��Ч = 2
        End If
    End If
    
    '����������ʼ
    mvarCond.Begin = CDate(0)
    mvarCond.End = CDate(0)
    mvarCond.����ID = 0
    
    Call ClearPatiCond
    
    '������
    mstr������ = zlDatabase.GetPara("������", glngSys, pҽ������վ, "")
    If mstr������ <> "" Then mvarCond.������ = mstr������
    
    'һ��ͨ������ʼ������tbcSub_SelectedChanged֮ǰ���Ա㴫�ݸ�ҽ������
    'zlGetIDKindStr�л��Զ�����Ϊ����8λ����
    mstrCardKind = "��|���￨|0|0|8|0|0|0;��|��ʶ��|0|0|0|0|0|0;��|���ݺ�|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;�ɣ�|�ɣÿ�|1|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    
    
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
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'DockingPane
    '-----------------------------------------------------
    Me.DkpMain.SetCommandBars Me.cbsMain
    Me.DkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.DkpMain.Options.ThemedFloatingFrames = True
    Me.DkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.DkpMain.CreatePane(1, IIf(mbytSize = 0, 280, 300), 400, DockLeftOf, Nothing)
    objPane.Title = "ִ�в����б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.DkpMain.CreatePane(2, 310, 100, DockBottomOf, objPane)
    objPane.Title = "��Ϣ����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    '����ѪҺ����
    If gblnѪ��ϵͳ = True Then
        With DkpBlood
            .Options.UseSplitterTracker = False 'ʵʱ�϶�
            .Options.ThemedFloatingFrames = True
            .Options.AlphaDockingContext = True
            .Options.HideClient = True
            
            Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing)
            objPane.Title = "��Ѫִ�еǼ�"
            objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
        End With
        If InitObjBlood = True Then
            Set mobjFrmBloodExe = gobjPublicBlood.zlGetBloodExec
            mobjFrmBloodExe.IsShowExec = True
        End If
    End If
    
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p�°����ﲡ��, True) <> "" Or GetInsidePrivs(p�°�סԺ����, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "���Ӳ���")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    Set mclsOutAdvices = New zlPublicAdvice.clsDockOutAdvices
    Set mclsInAdvices = New zlPublicAdvice.clsDockInAdvices
    Set mclsExpenses = New zlPublicExpense.clsDockExpense
    Set mclsInEPRs = New zlRichEPR.cDockInEPRs
    Set mclsOutEPRs = New zlRichEPR.cDockOutEPRs
    Set mclsTends = New zlRichEPR.cDockInTends
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Set mclsTendEPRs = New zlRichEPR.cDockInTendEPRs
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    If Not mclsExpenses Is Nothing Then
        Call mclsExpenses.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    End If
    
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    mcolSubForm.Add mclsExpenses.zlGetForm, "_ҽ������"
    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_����ҽ��"
    mcolSubForm.Add mclsInAdvices.zlGetForm, "_סԺҽ��"
    mcolSubForm.Add mclsInEPRs.zlGetForm, "_סԺ����"
    mcolSubForm.Add mclsOutEPRs.zlGetForm, "_���ﲡ��"
    mcolSubForm.Add mclsTends.zlGetForm, "_����"
    mcolSubForm.Add mclsTendsNew.zlGetForm, "_�°滤��"
    mcolSubForm.Add mclsTendEPRs.zlGetForm, "_������"
    
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
        If GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
            If mobjAppendBill Is Nothing Then
                .InsertItem(intIdx, "ҽ�����ӷ���", picTmp.hwnd, 0).Tag = "ҽ������": intIdx = intIdx + 1
            Else
                mbln���Ѱ�ť = True
            End If
        End If
        If GetInsidePrivs(p����ҽ���´�, True) <> "" Then
            .InsertItem(intIdx, "����ҽ��", picTmp.hwnd, 0).Tag = "����ҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Then
            .InsertItem(intIdx, "סԺҽ��", picTmp.hwnd, 0).Tag = "סԺҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺ��������, True) <> "" Then
            .InsertItem(intIdx, "סԺ����", picTmp.hwnd, 0).Tag = "סԺ����": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngInIndex = intIdx - 1
        End If
        mlngNewIndex = -1
        If (GetInsidePrivs(p�°����ﲡ��, True) <> "" Or GetInsidePrivs(p�°�סԺ����, True) <> "") And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0).Tag = "�²���": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNewIndex = intIdx - 1
        End If
        If GetInsidePrivs(p���ﲡ������, True) <> "" Then
            .InsertItem(intIdx, "���ﲡ��", picTmp.hwnd, 0).Tag = "���ﲡ��": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngOutIndex = intIdx - 1
        End If
        
        If GetInsidePrivs(p�����¼����, True) <> "" Then
            .InsertItem(intIdx, "�����¼", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNurIndex = intIdx - 1
        End If
        
        If GetInsidePrivs(p�����¼����, True) <> "" Then
            .InsertItem(intIdx, "�����¼", picTmp.hwnd, 0).Tag = "�°滤��": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNewNurIndex = intIdx - 1
            .InsertItem(intIdx, "������", picTmp.hwnd, 0).Tag = "������": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNurEMRIndex = intIdx - 1
        End If
        
        '��Ҳ����п�Ƭ
        Call CreatePlugInOK(pҽ������վ)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, pҽ������վ)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, pҽ������վ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ��ҽ������վ��Ȩ��(�����Ƿ�߱�:ҽ�����ӷ���,����ҽ���´�,סԺҽ���´��Ȩ��֮һ)��", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = zlDatabase.GetPara("ҽ������", glngSys, pҽ������վ)
        If mvarCond.��Դ = "2,3" Then
            Call ExchangeAdvice(False) 'ȱʡ��ʾסԺ��
        Else
            Call ExchangeAdvice(True) 'ȱʡ��ʾ�����
        End If
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '���⼤���¼�
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
        End If
        'ֻ����ѡ����Ӵ���
        Call tbcSub_SelectedChanged(.Selected)
    End With
    
     '������������
    Call InitReportColumn
    picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picExec.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '��ȡ��������
    '-----------------------------------------------------
    mlngDept = -1
    mblnѪ͸�� = False
    mbln���� = False
    mstrDeptNode = ""
    mstrPrePati = ""
    mstr״̬ = zlDatabase.GetPara("ִ��״̬", glngSys, pҽ������վ, "01101", _
        Array(chkִ��״̬(0), chkִ��״̬(1), chkִ��״̬(2), chkִ��״̬(3), chkִ��״̬(4)), InStr(mstrPrivs, "��������") > 0) 'ȱʡ��ʾδִ�С�����ִ��
    chkִ��״̬(0).Value = Val(Mid(mstr״̬, 1, 1))
    chkִ��״̬(1).Value = Val(Mid(mstr״̬, 2, 1))
    chkִ��״̬(2).Value = Val(Mid(mstr״̬, 3, 1))
    chkִ��״̬(3).Value = Val(Mid(mstr״̬, 4, 1))
    chkִ��״̬(4).Value = Val(Mid(mstr״̬, 5, 1))
    chkִ��״̬(4).Enabled = chkִ��״̬(2).Value = 1
    If Val(gstrҽ���˶�) = 0 Then
        chkִ��״̬(4).Visible = False
        Image1(4).Visible = False
    End If
    
    mintFindType = Val(zlDatabase.GetPara("���˲��ҷ�ʽ", glngSys, pҽ������վ, "1", , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    mblnFilter = Val(zlDatabase.GetPara("������ʾģʽ", glngSys, pҽ������վ, , , , intType)) <> 0
    mblnFilterEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    mblnֻ�����շ� = Val(zlDatabase.GetPara("ֻ��ʾ���շѵĲ���", glngSys, pҽ������վ, , , , intType)) <> 0
    mblnֻ�����շ�Enabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    mbln����ִ�� = False
    mblnExeLog = Val(zlDatabase.GetPara("��¼ִ�����", glngSys, pҽ������վ, "0")) <> 0
    mblnƤ����֤ = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, pҽ������վ)) <> 0
    
    mstr������� = zlDatabase.GetPara("�������", glngSys, pҽ������վ)
    mstr������� = zlDatabase.GetPara("�������", glngSys, pҽ������վ)
    mbyt������˷�ʽ = Val(zlDatabase.GetPara(185, glngSys))
        
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    err.Clear: On Error GoTo 0
    
    
    Call SetTimer '�����Զ�ˢ��
    
    
    'ҽ�����ҳ�ʼ��
    '-----------------------------------------------------
    If Not InitDepts Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "���п���") > 0 Then
            MsgBox "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Else
            MsgBox "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If
    
    '����ָ�:�������ִ��
    '-----------------------------------------------------
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
                strTab = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(DkpMain), DkpMain.Name, "")
        If InStr(strTab, "��Ϣ����") <> 0 Then DkpMain.LoadStateFromString strTab
    End If
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    rptPati.Columns.Find(col_ѡ��).Visible = True '= mblnFilter
    rptPati.Columns.Find(col_����).Visible = mblnShowBed
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("ҽ��ˢ�¼��", glngSys, pҽ������վ))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
End Sub

Private Sub InitExecTable()
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "Ҫ��ʱ��,1520,1;ִ��ʱ��,1520,1;��������,815,1;ִ��ժҪ,1550,1;ִ����,700,1;�Ǽ�ʱ��,1520,1;�Ǽ���,700,1;ִ�н��,815,1;�˶���,750,1;�˶�ʱ��,1530,1;˵��,500,1;��Դ,600,1"
    arrHead = Split(strHead, ";")
    With vsExec
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLExec.Count <> UBound(arrHead) + 1 Then COLExec.Add i, Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        
    End With
End Sub

Private Function FuncExecAuditBatch() As Boolean
'���ܣ������˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim strMsgNameSame As String
    Dim strMsgNoRecord As String
    Dim strMsgHave As String
    Dim strMsgKZ As String
    Dim strMsg As String
    
    bln��ѪƤ�� = False
    
    On err GoTo errH
    arrSQL = Array()
    For i = 0 To rptPati.Rows.Count - 1
        With rptPati.Rows(i)
            If Not .GroupRow Then
                If rptPati.Rows(i).Record(col_ѡ��).Checked Then
                    If (Mid(gstrҽ���˶�, 2, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "1" Or _
                        Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "8" Or _
                        Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "K") Then
                        
                        bln��ѪƤ�� = True
                        If .Record(COL_�˶���).Value & "" = "" Then
                            If Val(.Record(col_ִ��״̬).Value & "") = 3 Then
                                If str�˶��� = "" Then str�˶��� = zlDatabase.UserIdentifyByUser(Me, "�ں˶�ִ�����ǰ�������������û�����������������֤��", glngSys, pҽ������վ, "ִ������Ǽ�", , True)
                                If str�˶��� = "" Then Exit Function
                                
                                '��ȡִ����
                                strSQL = "Select ִ���� From ����ҽ��ִ�� Where ҽ��ID=[1] and ���ͺ�=[2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(col_ҽ��ID).Value & ""), Val(.Record(col_���ͺ�).Value & ""))
                                If rsTmp.RecordCount > 0 Then
                                    If str�˶��� = rsTmp!ִ���� & "" Then
                                        strMsgNameSame = strMsgNameSame & "," & .Record(col_���ݺ�).Value
                                    Else
                                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Insert(" & Val(.Record(col_ҽ��ID).Value) & "," & Val(.Record(col_���ͺ�).Value) & ",'" & str�˶��� & "')"
                                    End If
                                End If
                            Else
                                strMsgNoRecord = strMsgNoRecord & "," & .Record(col_���ݺ�).Value
                            End If
                        Else
                            strMsgHave = strMsgHave & "," & .Record(col_���ݺ�).Value
                        End If
                    Else
                        strMsgKZ = strMsgKZ & "," & .Record(col_���ݺ�).Value
                    End If
                    rptPati.Rows(i).Record(col_ѡ��).Checked = False
                End If
            End If
        End With
    Next
    
    If bln��ѪƤ�� = False Then
        If Val(gstrҽ���˶�) = 1 Then
            strSQL = "�㹴ѡ����Ŀ��û��Ƥ����Ŀ������˶ԡ�"
        ElseIf Val(gstrҽ���˶�) = 10 Then
            strSQL = "�㹴ѡ����Ŀ��û����Ѫ��Ŀ������˶ԡ�"
        Else
            strSQL = "�㹴ѡ����Ŀ��û����Ѫ��Ƥ����Ŀ������˶ԡ�"
        End If
        MsgBox strSQL, vbInformation, "ҽ���˶�"
        '��ʾִ�����
        Call LoadPatients
        Exit Function
    End If

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If strMsgNameSame <> "" Then
        strMsg = strMsg & "���µ��ݵ�����˺�ִ����Ϊͬһ���ˣ�" & vbCrLf & Mid(strMsgNameSame, 2) & "��" & vbCrLf
    End If
    If strMsgNoRecord <> "" Then
        strMsg = strMsg & "���µ��ݻ�δ����ִ������Ǽǣ�" & vbCrLf & Mid(strMsgNoRecord, 2) & "��" & vbCrLf
    End If
    If strMsgHave <> "" Then
        strMsg = strMsg & "���µ����Ѿ������˺˶ԣ�" & vbCrLf & Mid(strMsgHave, 2) & "��" & vbCrLf
    End If
    If strMsgKZ <> "" Then
        strMsg = strMsg & "���µ��ݲ�����Ѫ����Ƥ������Ŀ��" & vbCrLf & Mid(strMsgKZ, 2) & "��" & vbCrLf
    End If
    
    If UBound(arrSQL) < 0 Then
        MsgBox "����ѡ����Ŀδ�˶Գɹ������У�" & vbCrLf & strMsg, vbInformation, "ҽ���˶�"
    Else
        If strMsg <> "" Then
            MsgBox "���˶���" & UBound(arrSQL) + 1 & "����Ŀ�����У�" & vbCrLf & strMsg, vbInformation, "ҽ���˶�"
        End If
    End If

    '��ʾִ�����
    Call LoadPatients
    FuncExecAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingAudit() As Boolean
'���ܣ��˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim i As Long
    
    '�ж��Ƿ�����ִ��ģʽ�������ڵ����������е���
    If rptPati.Columns(col_ѡ��).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_ѡ��).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("Ҫ�Ե�ǰѡ���һ��������Ŀ���к˶���", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecAuditBatch
            End If
            Exit Function
        End If
    End If
    
    With rptPati.SelectedRows(0)
        bln��ѪƤ�� = (Mid(gstrҽ���˶�, 2, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "1" Or _
            Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "8" Or _
            Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "K")
            
        If Not bln��ѪƤ�� Then
            If Val(gstrҽ���˶�) = 1 Then
                MsgBox "ֻ�ܺ˶�Ƥ��ҽ����", vbInformation, gstrSysName
            ElseIf Val(gstrҽ���˶�) = 10 Then
                MsgBox "ֻ�ܺ˶���Ѫҽ����", vbInformation, gstrSysName
            Else
                MsgBox "ֻ�ܺ˶���Ѫ����Ƥ��ҽ����", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        
        If vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) <> "" Then
            MsgBox "��ҽ�����Ѿ��˶ԣ������ٴκ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.FixedRows, vsExec.FixedCols) = "" Then
            MsgBox "��ҽ����δ����ִ������Ǽǣ����ܺ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        str�˶��� = zlDatabase.UserIdentifyByUser(Me, "�ں˶�ִ�����ǰ�������������û�����������������֤��", glngSys, pҽ������վ, "ִ������Ǽ�", , True)
        If str�˶��� = "" Then Exit Function
        
        If str�˶��� = vsExec.TextMatrix(vsExec.FixedRows, COLExec("ִ����")) Then
            MsgBox "ִ���˲��ܺ��������ͬ�����ܺ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    With vsExec
        On Error GoTo errH
        strSQL = "Zl_����ҽ���˶�_Insert(" & Val(rptPati.SelectedRows(0).Record(col_ҽ��ID).Value) & "," & Val(rptPati.SelectedRows(0).Record(col_���ͺ�).Value) & ",'" & str�˶��� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "ҽ���˶�")
        '��ʾִ�����
        Call LoadPatients
        FuncThingAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncExecDelAuditBatch() As Boolean
'���ܣ�����ȡ���˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim strMsg As String
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim blnIsTwo As Boolean   '�ж��Ƿ�������������ϵĺ˶���
    Dim strTmp As String
    Dim bln�˶��� As Boolean
    Dim strMsgNoRecord As String
    Dim strMsgHave As String
    Dim strMsgKZ As String
    Dim strMsgNoExec As String
    Dim datCur As Date
    
    bln��ѪƤ�� = False
    
    On err GoTo errH
    arrSQL = Array()
    datCur = zlDatabase.Currentdate
    For i = 0 To rptPati.Rows.Count - 1
        With rptPati.Rows(i)
            If Not .GroupRow Then
                If rptPati.Rows(i).Record(col_ѡ��).Checked And Not blnIsTwo And Not bln�˶��� Then
                    If (Mid(gstrҽ���˶�, 2, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "1" Or _
                        Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "8" Or _
                        Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "K") Then
                        
                        bln��ѪƤ�� = True
                        If .Record(COL_�˶���).Value & "" <> "" Then
                            If Val(.Record(col_ִ��״̬).Value & "") = 3 Then
                                If strTmp <> "" And strTmp <> .Record(COL_�˶���).Value & "" Then
                                    blnIsTwo = True
                                Else
                                    strTmp = .Record(COL_�˶���).Value & ""
                                End If
                                If CanUnExec(CDate(.Record(COL_����ʱ��).Value & ""), datCur) Then
                                    If .Record(COL_�˶���).Value & "" <> UserInfo.���� Then
                                        If str�˶��� = "" Then str�˶��� = zlDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", glngSys, pҽ������վ, "ִ������Ǽ�", , True)
                                        If str�˶��� = "" Then Exit Function
                                        
                                        If str�˶��� = .Record(COL_�˶���).Value & "" Then
                                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                            arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Delete(" & Val(.Record(col_ҽ��ID).Value) & "," & Val(.Record(col_���ͺ�).Value) & ")"
                                        Else
                                            bln�˶��� = True
                                            str�˶��� = .Record(COL_�˶���).Value & ""
                                        End If
                                    Else
                                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Delete(" & Val(.Record(col_ҽ��ID).Value) & "," & Val(.Record(col_���ͺ�).Value) & ")"
                                    End If
                                Else
                                    strMsgNoExec = strMsgNoExec & "," & .Record(col_���ݺ�).Value
                                End If
                            Else
                                strMsgNoRecord = strMsgNoRecord & "," & .Record(col_���ݺ�).Value
                            End If
                        Else
                            strMsgHave = strMsgHave & "," & .Record(col_���ݺ�).Value
                        End If
                    Else
                        strMsgKZ = strMsgKZ & "," & .Record(col_���ݺ�).Value
                    End If
                    rptPati.Rows(i).Record(col_ѡ��).Checked = False
                End If
            End If
        End With
    Next
    
    If bln��ѪƤ�� = False Then
        If Val(gstrҽ���˶�) = 1 Then
            strSQL = "�㹴ѡ����Ŀ��û��Ƥ����Ŀ������ȡ���˶ԡ�"
        ElseIf Val(gstrҽ���˶�) = 10 Then
            strSQL = "�㹴ѡ����Ŀ��û����Ѫ��Ŀ������ȡ���˶ԡ�"
        Else
            strSQL = "�㹴ѡ����Ŀ��û����Ѫ��Ƥ����Ŀ������ȡ���˶ԡ�"
        End If
        MsgBox strSQL, vbInformation, "ȡ���˶�"
        '��ʾִ�����
        Call LoadPatients
        Exit Function
    End If
    
    If blnIsTwo Then
        MsgBox "����ͬʱȡ������˺˶Ե���Ŀ����ѡ��ͬһ�������˶Ե���Ŀ��", vbInformation, "ȡ���˶�"
        '��ʾִ�����
        Call LoadPatients
        Exit Function
    End If
    
    If bln�˶��� Then
        MsgBox "ֻ��ȡ���Լ��˶Ե�ҽ������ǰѡ���ҽ���˶�����""" & str�˶��� & """��", vbInformation, "ȡ���˶�"
        '��ʾִ�����
        Call LoadPatients
        Exit Function
    End If
    

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "ȡ���˶�")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If strMsgNoRecord <> "" Then
        strMsg = strMsg & "���µ��ݻ�δ����ִ������Ǽǣ�" & vbCrLf & Mid(strMsgNoRecord, 2) & "��" & vbCrLf
    End If
    If strMsgHave <> "" Then
        strMsg = strMsg & "���µ��ݻ�δ���к˶ԣ�" & vbCrLf & Mid(strMsgHave, 2) & "��" & vbCrLf
    End If
    If strMsgKZ <> "" Then
        strMsg = strMsg & "���µ��ݲ�����Ѫ����Ƥ������Ŀ��" & vbCrLf & Mid(strMsgKZ, 2) & "��" & vbCrLf
    End If
    If strMsgNoExec <> "" Then
        strMsg = strMsg & "���µ��ݵĺ˶�ʱ�䳬����ҽ��ִ����Ч������" & vbCrLf & Mid(strMsgNoExec, 2) & "��" & vbCrLf
    End If
    
    If UBound(arrSQL) < 0 Then
        MsgBox "����ѡ����Ŀδȡ���ɹ������У�" & vbCrLf & strMsg, vbInformation, "ȡ���˶�"
    Else
        If strMsg <> "" Then
            MsgBox "��ȡ���˶���" & UBound(arrSQL) + 1 & "����Ŀ,���У�" & vbCrLf & strMsg, vbInformation, "ȡ���˶�"
        End If
    End If
    '��ʾִ�����
    Call LoadPatients
    FuncExecDelAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingDelAudit() As Boolean
'���ܣ�ȡ���˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim i As Long
    
    '�ж��Ƿ�����ִ��ģʽ�������ڵ����������е���
    If rptPati.Columns(col_ѡ��).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_ѡ��).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("Ҫ�Ե�ǰѡ���һ��������Ŀ����ȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecDelAuditBatch
            End If
            Exit Function
        End If
    End If
    
    With rptPati.SelectedRows(0)
        bln��ѪƤ�� = (Mid(gstrҽ���˶�, 2, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "1" Or _
            Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "8" Or _
            Mid(gstrҽ���˶�, 1, 1) = "1" And .Record(COL_�������).Value = "K")
            
        If Not bln��ѪƤ�� Then
            If Val(gstrҽ���˶�) = 1 Then
                MsgBox "ֻ��ȡ���˶�Ƥ��ҽ����", vbInformation, gstrSysName
            ElseIf Val(gstrҽ���˶�) = 10 Then
                MsgBox "ֻ��ȡ���˶���Ѫҽ����", vbInformation, gstrSysName
            Else
                MsgBox "ֻ��ȡ���˶���Ѫ����Ƥ��ҽ����", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        
        If vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) = "" Then
            MsgBox "��ҽ����δ���к˶ԣ�����ȡ����", vbInformation, gstrSysName
            Exit Function
        End If

    End With
    With vsExec
        If vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) <> UserInfo.���� Then
            str�˶��� = zlDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", glngSys, pҽ������վ, "ִ������Ǽ�", , True)
            If str�˶��� = "" Then Exit Function
            If str�˶��� <> vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) Then
                MsgBox "ֻ��ȡ���Լ��˶Ե�ҽ������ǰҽ���˶�����""" & vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) & """", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If MsgBox("��ȷ��Ҫȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
        End If
        On Error GoTo errH
        
        strSQL = "Zl_����ҽ���˶�_Delete(" & Val(rptPati.SelectedRows(0).Record(col_ҽ��ID).Value) & "," & Val(rptPati.SelectedRows(0).Record(col_���ͺ�).Value) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "ȡ��ҽ���˶�")
        Call LoadPatients
        FuncThingDelAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long, strCardNO As String
    Dim lngҽ��ID As Long
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    '��Ѫҽ�����س�����
    If picBlood.Visible = True And InStr(1, "," & mstrBloodControlIDs & ",", "," & Control.ID & ",") <> 0 Then
        If Not mobjFrmBloodExe Is Nothing Then
            Call mobjFrmBloodExe.zlExecuteCommandBars(Control)
        End If
        Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                    objControl.Style = xtpButtonIcon
                Else
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S 'С����
        If mbytSize <> 0 Then
            mbytSize = 0
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '������
        If mbytSize <> 1 Then
            mbytSize = 1
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_Tool_Archive '���Ӳ�������
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).Record(col_��Դ).Value = "סԺ" Then
                Call frmArchiveView.ShowArchive(Me, mlng����ID, mlng��ҳID)
            Else
                Call frmArchiveView.ShowArchive(Me, mlng����ID, Get�Һ�ID(mstr�Һŵ�))
            End If
        End If
    Case conMenu_Tool_Reference_1 '������ϲο�
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '������Ŀ��������
        Call Set������Ŀ��������
         
    Case conMenu_View_Find '����
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '��ʱ��Ҫ��λһ��
            If PatiIdentify.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            PatiIdentify.SetFocus
        End If
    Case conMenu_View_FindNext '������һ��
        If PatiIdentify.Text = "" And mvarCond.���֤ = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True, IIf(PatiIdentify.Text = "", mvarCond.���֤, ""))
        End If
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                rptPati.SelectedRows(0).Expanded = False
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    rptPati.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '���۵���λ��������,�����Զ�������¼�
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        If rptPati.SelectedRows.Count > 0 Then
            rptPati.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '�۵�������
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '���۵���λ��������,�����Զ�������¼�
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_AllExpend 'չ��������
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_PatInfor '������ʾ:ֻ��ʾ���շѲ���
        mblnֻ�����շ� = Not mblnֻ�����շ�
        cbsMain.RecalcLayout: Call LoadPatients
    Case conMenu_View_ShowAll '������ʾ:��ʾ����ִ�еĲ���
        mbln����ִ�� = Not mbln����ִ��
        cbsMain.RecalcLayout: Call LoadPatients
    Case conMenu_View_Show '���˹���
        Call PatientFilter
    Case conMenu_View_Refresh 'ˢ��
        Call LoadPatients
        Call LoadNotify
    
    Case conMenu_File_RoomSet 'ִ�м�����
        With frmTechnicRoom
            .lblDept.Tag = Me.cboDept.ItemData(cboDept.ListIndex)
            .lblDept.Caption = Me.cboDept.Text & "ִ�м�"
            .Show 1, Me
        End With
    Case conMenu_File_Parameter '��������
        Call ParameterSetup
    Case conMenu_Manage_Bespeak 'ʱ�䰲��
        Call FuncExecPlanTime
    Case conMenu_Manage_Plan 'ִ�б���
        Call FuncExecPlan
    Case conMenu_Manage_Logout 'ȡ������
        Call FuncExecErase
    Case conMenu_Manage_Refuse '�ܾ�ִ��
        Call FuncExecRefuse
    Case conMenu_Manage_ReGet 'ȡ���ܾ�
        Call FuncExecRestore
    Case conMenu_Manage_ThingAdd '��¼ִ�����
        Call FuncThingNew
    Case conMenu_Manage_ThingModi '����ִ�����
        Call FuncThingModi
    Case conMenu_Manage_ThingDel 'ɾ��ִ�����
        Call FuncThingDel
    Case conMenu_Manage_ThingAudit '�˶�
        Call FuncThingAudit
    Case conMenu_Manage_ThingDelAudit 'ȡ���˶�
        Call FuncThingDelAudit
    Case conMenu_Manage_Complete 'ִ�����
        Call FuncExecFinish
    Case conMenu_Manage_Undone 'ȡ�����
        Call FuncExecCancel
    Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '��ӡ���Ƶ���
        Call FuncBillPrint(Control)
    Case conMenu_Manage_RequestBatPrint '������ӡ����
        Call FuncBatchPrint
    Case conMenu_Manage_ReportEdit '������д
        Call FuncShowReport(0)
    Case conMenu_Manage_ReportView '�������
        Call FuncShowReport(1)
    Case conMenu_Manage_ReportPrint '�����ӡ
        Call FuncShowReport(2)
    Case conMenu_Manage_ReportPreview '����Ԥ��
        Call FuncShowReport(3)
    Case conMenu_Manage_AppendBill '���Ѱ�ť
        Call FuncAppendBill
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case Else
        timRefresh.Enabled = False
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1263_1" Then 'ҽ����������
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                    "ִ�п���=" & zlCommFun.GetNeedName(cboDept.Text) & "|" & cboDept.ItemData(cboDept.ListIndex))
            Else
                If rptPati.SelectedRows.Count = 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ִ�п���=" & cboDept.ItemData(cboDept.ListIndex))
                Else
                    With rptPati.SelectedRows(0)
                        If .Record(col_��Դ).Value = "סԺ" Then
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                                "ִ�п���=" & cboDept.ItemData(cboDept.ListIndex), "ҽ��ID=" & .Record(col_ҽ��ID).Value, "���ͺ�=" & .Record(col_���ͺ�).Value, _
                                    "NO=" & .Record(col_���ݺ�).Value, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "סԺ��=" & .Record(col_��ʶ��).Value)
                        Else
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                                "ִ�п���=" & cboDept.ItemData(cboDept.ListIndex), "ҽ��ID=" & .Record(col_ҽ��ID).Value, "���ͺ�=" & .Record(col_���ͺ�).Value, _
                                    "NO=" & .Record(col_���ݺ�).Value, "����ID=" & mlng����ID, "�Һŵ�=" & mstr�Һŵ�, "�����=" & .Record(col_��ʶ��).Value)
                        End If
                    End With
                End If
            End If
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "ҽ������"
                Call mclsExpenses.zlExecuteCommandBars(Control)
            Case "����ҽ��"
                Call mclsOutAdvices.zlExecuteCommandBars(Control)
            Case "סԺҽ��"
                Call mclsInAdvices.zlExecuteCommandBars(Control)
            Case "סԺ����"
                Call mclsInEPRs.zlExecuteCommandBars(Control)
            Case "���ﲡ��"
                Call mclsOutEPRs.zlExecuteCommandBars(Control)
            Case "�²���"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "�°滤��"
                Call mclsTendsNew.zlExecuteCommandBars(Control)
            Case "������"
                Call mclsTendEPRs.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If rptPati.SelectedRows.Count <> 0 Then lngҽ��ID = rptPati.SelectedRows(0).Record(col_ҽ��ID).Value
                    Call gobjPlugIn.ExeButtomClick(glngSys, pҽ������վ, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng����ID, mlng��ҳID, mstr�Һŵ�, lngҽ��ID)
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        timRefresh.Enabled = True
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim i As Long, strKinds As String
    Dim arrKind() As String
    Dim objControl As CommandBarControl
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Manage_Report '����
        With CommandBar.Controls
            If .Count = 0 Then
                .Add xtpControlButton, conMenu_Manage_ReportEdit, "��д����(&E)"
                .Add xtpControlButton, conMenu_Manage_ReportView, "���ı���(&W)"
                .Add(xtpControlButton, conMenu_Manage_ReportPrint, "��ӡ����(&P)").BeginGroup = True
                .Add xtpControlButton, conMenu_Manage_ReportPreview, "Ԥ������(&V)"
            End If
        End With
    Case Else
        Select Case tbcSub.Selected.Tag
        Case "ҽ������"
            Call mclsExpenses.zlPopupCommandBars(CommandBar)
        Case "����ҽ��"
            Call mclsOutAdvices.zlPopupCommandBars(CommandBar)
        Case "סԺҽ��"
            Call mclsInAdvices.zlPopupCommandBars(CommandBar)
        Case "סԺ����"
            
        Case "���ﲡ��"
            
        End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnSelect As Boolean
    Dim intִ��״̬ As Integer, intִ�й��� As Integer
    Dim int�ۺ�״̬ As Integer, intִ�а��� As Integer
    Dim objControl As CommandBarControl
    Dim arrType() As String, i As Long
    
    '��ʼ��һ��ͨ����,����activate�¼��ڼ���ʱ������
    If Not mblnIsInit Then
        mblnIsInit = True
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
            Else
                mstrCardKind = mobjSquareCard.zlGetIDKindStr(mstrCardKind)
            End If
            Call PatiIdentify.zlInit(Me, glngSys, pҽ������վ, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
            PatiIdentify.objIDKind.AllowAutoICCard = True
            PatiIdentify.objIDKind.AllowAutoIDCard = True
        
            arrType = Split(mstrCardKind, ";")
            For i = 1 To UBound(arrType) + 1
                If i = mintFindType Then
                    PatiIdentify.objIDKind.IDKind = i
                    mstrFindType = PatiIdentify.objIDKind.Cards(i).����
                    Exit For
                End If
            Next
            chkFilter.Value = IIf(mblnFilter, 1, 0)
            chkFilter.Enabled = mblnFilterEnabled
        End If
    End If
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    '��Ѫҽ�����س�����
    If picBlood.Visible = True And InStr(1, "," & mstrBloodControlIDs & ",", "," & Control.ID & ",") <> 0 Then
        If Not mobjFrmBloodExe Is Nothing Then
            Call mobjFrmBloodExe.zlUpdateCommandBars(Control)
        End If
        Exit Sub
    End If
    
    '�Ƿ�ѡ���˲���
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            blnSelect = True
            '0-δִ��,1-��ִ��,2-�ܾ�ִ��,3-����ִ��
            intִ��״̬ = rptPati.SelectedRows(0).Record(col_ִ��״̬).Value
            int�ۺ�״̬ = rptPati.SelectedRows(0).Record(col_�ۺ�״̬).Value '����Ͷ���ִͬ��״̬
            '0-������;1-�ѱ���;2-�����;3-������;4-��д����;5-��˲���;6-�������
            intִ�й��� = rptPati.SelectedRows(0).Record(col_ִ�й���).Value
            '0-���谲��,1-��Ҫ����
            intִ�а��� = rptPati.SelectedRows(0).Record(col_ִ�а���).Value
        End If
    End If
        
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_FontSize_S 'С����
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_View_FontSize_L '������
        Control.Checked = (mbytSize = 1)
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPati.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = rptPati.SelectedRows(0).Expanded
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '�۵�/չ����
        Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
       
    Case conMenu_View_FindNext '������һ��
        Control.Enabled = Not mblnFilter
    Case conMenu_View_PatInfor '������ʾ:ֻ��ʾ���շѲ���
        Control.Checked = mblnֻ�����շ�
        Control.Enabled = mblnֻ�����շ�Enabled
    Case conMenu_View_ShowAll '������ʾ:��ʾ����ִ�еĲ���
        Control.Checked = mbln����ִ��
    Case conMenu_Tool_Archive '���Ӳ�������
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Manage_Bespeak 'ʱ�䰲��:�����Ƿ��ѱ���
        blnEnabled = blnSelect And int�ۺ�״̬ = 0 And intִ��״̬ = 0 And intִ�а��� = 1 '���з�ɢ״̬�˲�����
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Plan 'ִ�б���
        blnEnabled = blnSelect And int�ۺ�״̬ = 0 And intִ��״̬ = 0 '���з�ɢ״̬�˲�����
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Logout 'ȡ������
        blnEnabled = blnSelect And int�ۺ�״̬ = 0 And intִ��״̬ = 0 And intִ�й��� = 1 '���з�ɢ״̬�˲�����
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Refuse '�ܾ�ִ��
        blnEnabled = blnSelect And int�ۺ�״̬ = 0 And intִ��״̬ = 0 And intִ�й��� = 0 '���з�ɢ״̬�˲�����
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_ReGet 'ȡ���ܾ�
        blnEnabled = blnSelect And intִ��״̬ = 2
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_ThingAdd '��¼ִ�����
        Control.Enabled = blnSelect And (int�ۺ�״̬ = 0 Or int�ۺ�״̬ = 3)
    Case conMenu_Manage_ThingModi, conMenu_Manage_ThingDel '����ִ�����,ɾ��ִ�����
        Control.Enabled = blnSelect And (int�ۺ�״̬ = 0 Or int�ۺ�״̬ = 3) _
            And vsExec.TextMatrix(vsExec.Row, vsExec.FixedCols) <> "" And vsExec.Row = vsExec.FixedRows
    Case conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit
        Control.Enabled = blnSelect And (int�ۺ�״̬ = 0 Or int�ۺ�״̬ = 3)
    Case conMenu_Manage_Complete 'ִ�����
        If Me.Visible Then
            If rptPati.Columns(col_ѡ��).Visible Then
                Control.Enabled = CanBatchFinish(Control.ID)
            Else
                Control.Enabled = blnSelect And (int�ۺ�״̬ = 0 Or int�ۺ�״̬ = 3)
            End If
        End If
    Case conMenu_Manage_Undone 'ȡ�����
        Control.Enabled = blnSelect And int�ۺ�״̬ = 1
    Case conMenu_Manage_Request, conMenu_Manage_Report '���롢����˵�
        If blnSelect Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_RequestPrint '��ӡ���Ƶ���
        Control.Enabled = Control.CommandBar.Controls.Count > 0
    Case conMenu_Manage_RequestBatPrint '������ӡ����
        Control.Enabled = blnSelect
    Case conMenu_Manage_ReportEdit '������д����Ӧ�˲������ݣ������б�����������Ҫ��д
        If blnSelect Then blnEnabled = (intִ��״̬ = 0 Or intִ��״̬ = 3) _
            And rptPati.SelectedRows(0).Record(col_�ļ�ID).Value <> 0 And rptPati.SelectedRows(0).Record(col_������).Value <> 0
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_ReportView, conMenu_Manage_ReportPrint, conMenu_Manage_ReportPreview '�������/��ӡ/Ԥ��
        If blnSelect Then blnEnabled = rptPati.SelectedRows(0).Record(col_����ID).Value <> 0
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '�������������
        Control.Enabled = blnEnabled
    Case conMenu_Manage_AppendBill
        Control.Enabled = blnSelect And (int�ۺ�״̬ = 0 Or int�ۺ�״̬ = 3)
    Case Else
        Select Case tbcSub.Selected.Tag
        Case "ҽ������"
            Call mclsExpenses.zlUpdateCommandBars(Control)
        Case "����ҽ��"
            Call mclsOutAdvices.zlUpdateCommandBars(Control)
        Case "סԺҽ��"
            Call mclsInAdvices.zlUpdateCommandBars(Control)
        Case "סԺ����"
            Call mclsInEPRs.zlUpdateCommandBars(Control)
        Case "���ﲡ��"
            Call mclsOutEPRs.zlUpdateCommandBars(Control)
        Case "�²���"
            Call mclsEMR.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "�°滤��"
            Call mclsTendsNew.zlUpdateCommandBars(Control)
        Case "������"
            Call mclsTendEPRs.zlUpdateCommandBars(Control)
        End Select
    End Select
End Sub

Private Function CanBatchFinish(ByVal lngCmdID As Long) As Boolean
'���ܣ��ж�ָ���������ڵ�ǰѡ��״̬���ܷ�����ִ��
    Dim lngCount As Long, i As Long
    Dim blnEnabled As Boolean
    Dim str����ID As String
    
    With rptPati
        '��ѡ�������£���ѡ�����Ϊ׼
        For i = 0 To .Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(col_ѡ��).Checked Then
                    lngCount = lngCount + 1
                    If InStr(str����ID & ",", "," & .Rows(i).Record(col_����Id).Value & ",") = 0 Then
                        str����ID = str����ID & "," & .Rows(i).Record(col_����Id).Value
                    End If
                    If lngCmdID = conMenu_Manage_Complete Then
                        If Not (.Rows(i).Record(col_�ۺ�״̬).Value = 0 Or .Rows(i).Record(col_�ۺ�״̬).Value = 3) Then
                            Exit Function
                        End If
                        'If UBound(Split(Mid(str����ID, 2), ",")) > 0 Then Exit Function 'ֻ���Զ�һ�������������
                    End If
                End If
            End If
        Next
    
        'һ����û��ѡ�������£��Ե�ǰ��Ϊ׼
        If lngCount = 0 Then
            blnEnabled = False
            If .SelectedRows.Count > 0 Then
                If Not .SelectedRows(0).GroupRow Then
                    If lngCmdID = conMenu_Manage_Complete Then
                        If .SelectedRows(0).Record(col_�ۺ�״̬).Value = 0 Or .SelectedRows(0).Record(col_�ۺ�״̬).Value = 3 Then
                            blnEnabled = True
                        End If
                    End If
                End If
            End If
            If Not blnEnabled Then Exit Function
        End If
    End With
    
    CanBatchFinish = True
End Function

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim str����� As String
    '1.����ͨ��"���ж�"��־����,��Ϊ�Ӵ��廹Ҫ�ж�������
    '2.ֻ��Ҫ�жϵ�������д�����Ȼ�Ѳ��ô����Ҳ������(���Ӵ����е�)
    Control.Visible = True
    Select Case Control.ID
    
    Case conMenu_Manage_FeeItemSet
        If InStr(mstrPrivs, "���ﲡ��") = 0 Then    'û��"������Ŀ��������"Ȩ��ʱ�ɲ鿴
            Control.Visible = False
        End If
            
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(p���Ӳ�������) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_1 '������ϲο�
        If GetInsidePrivs(p������ϲο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 'ҩƷ�����Ʋο�
        If GetInsidePrivs(pҩƷ���Ʋο�) = "" Then Control.Visible = False
    Case conMenu_File_Parameter '��������
        'If InStr(mstrPrivs, "��������") = 0 Then Control.Visible = False
    Case conMenu_View_ShowAll '��ʾ����ִ�еĲ���
        If InStr(mstrPrivs, "ִ��������Ŀ") = 0 Then Control.Visible = False
    Case conMenu_File_RoomSet 'ִ�м�����
        If InStr(mstrPrivs, "ִ�м�����") = 0 Then Control.Visible = False
    Case conMenu_Manage_Bespeak, conMenu_Manage_Plan, conMenu_Manage_Logout 'ʱ�䰲��,ִ�а���,ȡ������
        If InStr(mstrPrivs, "ִ�а���") = 0 Then Control.Visible = False
    Case conMenu_Manage_Refuse, conMenu_Manage_ReGet '�ܾ�ִ��,ȡ���ܾ�
        If InStr(mstrPrivs, "�ܾ�ִ��") = 0 Then Control.Visible = False
    Case conMenu_Manage_ThingAdd, conMenu_Manage_ThingModi, conMenu_Manage_ThingDel '��¼,����,ɾ��ִ�����
        If InStr(mstrPrivs, "ִ������Ǽ�") = 0 Then Control.Visible = False
     Case conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit
        If InStr(GetInsidePrivs(pҽ������վ), "ִ������Ǽ�") = 0 Or Val(gstrҽ���˶�) = 0 Then Control.Visible = False
    Case conMenu_Manage_ThingAudit * 100# + 1, conMenu_Manage_ThingDelAudit * 100# + 1
        If InStr(GetInsidePrivs(pҽ������վ), "ִ������Ǽ�") = 0 Or picBlood.Visible = False Then Control.Visible = False
    Case conMenu_Manage_Complete 'ִ�����
        If InStr(mstrPrivs, "ȷ��ִ�����") = 0 Then Control.Visible = False
    Case conMenu_Manage_Undone 'ȡ�����
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow Then
                str����� = Trim(rptPati.SelectedRows(0).Record(col_�����).Value & "")
            End If
        End If
        If str����� = "" Then
            Control.Visible = False
        Else
            If InStr(mstrPrivs, "ȡ��ִ�����") > 0 And str����� = UserInfo.���� Or InStr(mstrPrivs, "ȡ������ִ�����") > 0 And str����� <> UserInfo.���� Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        End If
    Case conMenu_Manage_ReportEdit '�����
        If InStr(mstrPrivs, "������д") = 0 Then Control.Visible = False
    Case conMenu_Manage_RequestBatPrint '������ӡ����
        If InStr(mstrPrivs, "��ӡ��������") = 0 Then Control.Visible = False
    End Select
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If
    
    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)
        
    Me.Caption = "ҽ������վ - " & objItem.Caption & "(��ǰ�û���" & UserInfo.���� & ")"
    
    If InStr(mstrNotify, "1") > 0 Then
        DkpMain.Panes(2).Closed = False '�����
        DkpMain.Panes(2).Hidden = Val(DkpMain.Panes(2).Tag) = 1
        DkpMain.Panes(2).Title = "��Ϣ����"
    Else
        DkpMain.Panes(2).Tag = IIf(DkpMain.Panes(2).Hidden, 1, 0)
        DkpMain.Panes(2).Close
    End If
    
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '���������¼���
    Call MainDefCommandBar
    
    '�Ӵ������¼���
    Select Case objItem.Tag
    Case "ҽ������"
        Call mclsExpenses.zlDefCommandBars(Me, Me.cbsMain, mobjSquareCard)
    Case "����ҽ��"
        Call mclsOutAdvices.zlDefCommandBars(Me, Me.cbsMain, 2, Nothing, mobjSquareCard)
    Case "סԺҽ��"
        Call mclsInAdvices.zlDefCommandBars(Me, Me.cbsMain, 2, False, mobjSquareCard)
    Case "סԺ����"
        Call mclsInEPRs.zlDefCommandBars(Me.cbsMain)
    Case "���ﲡ��"
        Call mclsOutEPRs.zlDefCommandBars(Me.cbsMain)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "����"
        Call mclsTends.zlDefCommandBars(Me.cbsMain)
    Case "�°滤��"
        Call mclsTendsNew.zlDefCommandBars(Me.cbsMain, True)
    Case "������"
        Call mclsTendEPRs.zlDefCommandBars(Me.cbsMain, True)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, pҽ������վ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(err, "GetButtomName")
            '�����˵�
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            err.Clear: On Error GoTo 0
        End If
    End Select
    
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                objControl.Style = xtpButtonIcon
            Else
                objControl.Style = bytStyle
            End If
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next
    
    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ������ݼ�״̬
    Dim int���� As Integer
    Dim blnEdit As Boolean
    
    If mlng����ID = 0 Or (Me.Visible And Not objItem.Visible) Then '��������Ϊ�쳣����ԭ��
        'Ҫ���Ӵ��尴�����ݴ������
        Select Case objItem.Tag
        Case "ҽ������"
            Call mclsExpenses.zlRefresh(0, "")
        Case "����ҽ��"
            Call mclsOutAdvices.zlRefresh(0, "", False)
        Case "סԺҽ��"
            Call mclsInAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "סԺ����"
            Call mclsInEPRs.zlRefresh(0, 0, 0, False, False)
        Case "���ﲡ��"
            Call mclsOutEPRs.zlRefresh(0, 0, 0, False, False)
        Case "�²���"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 2)
        Case "����"
            Call mclsTends.zlRefresh(0, 0, 0, False, False)
        Case "�°滤��"
            Call mclsTendsNew.zlRefresh(0, 0, 0, False, False)
        Case "������"
            Call mclsTendEPRs.zlRefresh(0, 0, 0, False, False, False)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, pҽ������վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        Select Case objItem.Tag
        Case "ҽ������"
            With rptPati.SelectedRows(0)
                Call mclsExpenses.zlRefresh(cboDept.ItemData(cboDept.ListIndex), .Record(col_ҽ��ID).Value & ":" & _
                    .Record(col_���ͺ�).Value & ":" & IIf(Not .ParentRow.GroupRow, 1, 0), .Record(COL_����ת��).Value = 1)
            End With
        Case "����ҽ��"
            With rptPati.SelectedRows(0)
                Call mclsOutAdvices.zlRefresh(mlng����ID, mstr�Һŵ�, _
                    InStr(",0,3,", .Record(col_ִ��״̬).Value) > 0 And .Record(col_��Դ).Value <> "���", _
                    .Record(COL_����ת��).Value = 1, _
                    .Record(col_ҽ��ID).Value, cboDept.ItemData(cboDept.ListIndex), mclsMipModule)
            End With
        Case "סԺҽ��"
            With rptPati.SelectedRows(0)
                If .Record(col_��Ժ����).Value = "" Then
                    If .Record(COL_״̬).Value = 3 Then
                        int���� = 1 'Ԥ��Ժ
                    ElseIf .Record(COL_״̬).Value = 2 Then
                        int���� = 6 'ת�ƻ�ת��������ס����
                    Else
                        int���� = 0 '��Ժ
                    End If
                Else
                    int���� = 2 '��Ժ
                End If
                Call mclsInAdvices.zlRefresh(mlng����ID, mlng��ҳID, .Record(col_����id).Value, _
                    .Record(col_����).Value, int����, .Record(COL_����ת��).Value = 1, _
                    .Record(col_ҽ��ID).Value, .Record(col_ִ��״̬).Value, cboDept.ItemData(cboDept.ListIndex), , , mclsMipModule)
            End With
         Case "סԺ����"
            blnEdit = True
            With rptPati.SelectedRows(0)
                If mlngType = pt��Ժ Or mlngType = pt���� Then
                    '1-�ȴ����;2-�ܾ����;3-�������;4-��鷴��;5-���鵵
                    If Not (.Record(col_���).Value = 0 Or .Record(col_���).Value = 2) Then
                        '��������Ժ��鷴��״̬����Ժ��δ�ύ���
                        If .Record(col_���).Value = 1 Then
                            blnEdit = False
                        Else
                            If PatiMedRecHaveSubmit(mlng����ID, mlng��ҳID) Then blnEdit = False
                        End If
                    End If
                End If
    
                Call mclsInEPRs.zlRefresh(mlng����ID, mlng��ҳID, _
                    mlngDept, blnEdit, Val(.Record(COL_����ת��).Value), 0, False, .Record(col_����id).Value + 0, mlngState)
            End With
        Case "���ﲡ��"
            With rptPati.SelectedRows(0)
                Call mclsOutEPRs.zlRefresh(mlng����ID, Val(.Record(col_�Һ�ID).Value), mlngDept, mlng����ID <> 0, Val(.Record(COL_����ת��).Value))
            End With
        Case "�²���"
            With rptPati.SelectedRows(0)
                If Val(.Record(col_�Һ�ID).Value) = 0 Then
                    If .Record(col_��Ժ����).Value = "" Then
                        If .Record(COL_״̬).Value = 3 Then
                            int���� = 1 'Ԥ��Ժ
                        ElseIf .Record(COL_״̬).Value = 2 Then
                            int���� = 6 'ת�ƻ�ת��������ס����
                        Else
                            int���� = 0 '��Ժ
                        End If
                    Else
                        int���� = 2 '��Ժ
                    End If
                    Call mclsEMR.zlRefresh(mlng����ID, mlng��ҳID, mlngDept, int����, 2)
                Else
                    Call mclsEMR.zlRefresh(mlng����ID, Val(.Record(col_�Һ�ID).Value), mlngDept, 1, 1)
                End If
            End With
        Case "����"
            blnEdit = True
            With rptPati.SelectedRows(0)
                If mlngType = pt��Ժ Or mlngType = pt���� Then
                    If Not (.Record(col_���).Value = 0 Or .Record(col_���).Value = 2 Or .Record(col_���).Value = 999) Then
                        '��������Ժ��鷴��״̬����Ժ��δ�ύ���
                        If .Record(col_ͼ��).Value = 1 Then blnEdit = False
                    End If
                End If
                Call mclsTends.zlRefresh(mlng����ID, mlng��ҳID, .Record(col_����id).Value + 0, blnEdit, False, .Record(col_����id).Value + 0, mlngState)
            End With
        Case "������"
                Call mclsTendEPRs.zlRefresh(mlng����ID, mlng��ҳID, rptPati.SelectedRows(0).Record(col_����id).Value + 0, True, False, Val(rptPati.SelectedRows(0).Record(COL_����ת��).Value))
        Case "�°滤��"
            blnEdit = True
            With rptPati.SelectedRows(0)
                If mlngType = pt��Ժ Or mlngType = pt���� Then
                    If Not (.Record(col_���).Value = 0 Or .Record(col_���).Value = 2 Or .Record(col_���).Value = 999) Then
                        '��������Ժ��鷴��״̬����Ժ��δ�ύ���
                        If .Record(col_ͼ��).Value = 1 Then blnEdit = False
                    End If
                End If
                Call mclsTendsNew.zlRefresh(mlng����ID, mlng��ҳID, .Record(col_����id).Value + 0, blnEdit, False, .Record(col_����id).Value + 0, mlngState)
            End With
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, pҽ������վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng����ID, mstr�Һŵ�, mlng��ҳID, Val(rptPati.SelectedRows(0).Record(COL_����ת��).Value), 0, cboDept.ItemData(cboDept.ListIndex))
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    End If
    Call SetFontSize(Not Me.Visible)
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��") '����
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_RoomSet, "ִ�м�����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "ִ��(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Request, "����(&R)")
        With objPopup.CommandBar.Controls
            .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��(&J)").BeginGroup = True
            .Add xtpControlButton, conMenu_Manage_RequestBatPrint, "������ӡ����(&B)"
        End With
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Report, "����(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ʱ�䰲��(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "ִ�б���(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ������(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "�ܾ�ִ��(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "ȡ���ܾ�(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "ִ�����(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ�����(&U)")
        '��Ѫҽ��ִ��ǰ�˶�
        If gblnѪ��ϵͳ = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit * 100# + 1, "�˲�(&V)")
            objControl.ToolTipText = "ִ��ǰ�˲�"
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_ThingAudit
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit * 100# + 1, "ȡ���˲�(&Z)")
            objControl.IconId = conMenu_Manage_ThingDelAudit
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "��¼ִ�����(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "����ִ�����(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "ɾ��ִ�����(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "ȡ���˶�")
        If mbln���Ѱ�ť Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_AppendBill, "����")
                objControl.IconId = conMenu_Edit_Append
        End If
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "�����С(&N)") '����
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_FontSize_S, "С����(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_FontSize_L, "������(&L)", -1, False '����
        End With

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "ֻ��ʾ���շѵĲ���(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowAll, "��ʾ����ִ�еĲ���(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "���˹���(&O)")
            objControl.IconId = conMenu_View_Filter
            
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "������Ŀ��������(&C)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With


    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
        
        If gblnѪ��ϵͳ = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit * 100# + 1, "�˲�")
            objControl.IconId = conMenu_Manage_ThingAudit
            objControl.ToolTipText = "ִ��ǰ�˲�"
            objControl.BeginGroup = True
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "��¼")
        If gblnѪ��ϵͳ = False Then objControl.BeginGroup = True
        objControl.ToolTipText = "��¼ִ�����"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "���")
        Set objPopup = .Add(xtpControlPopup, conMenu_Manage_Report, "����")
        objPopup.ID = conMenu_Manage_Report: objPopup.IconId = conMenu_Manage_Report
        If mbln���Ѱ�ť Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_AppendBill, "����")
                objControl.IconId = conMenu_Edit_Append
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "����")
            objControl.BeginGroup = True: objControl.IconId = conMenu_View_Filter: objControl.ToolTipText = "���˹���"
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����") '����
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyW, conMenu_Manage_Bespeak 'ʱ�䰲��
        .Add FCONTROL, vbKeyL, conMenu_Manage_Plan 'ִ�б���
        .Add FCONTROL, vbKeyV, conMenu_Manage_ThingAdd '��¼ִ�����
        .Add FCONTROL, vbKeyI, conMenu_Manage_Complete 'ִ�����
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyF, conMenu_View_Find '���Ҳ���
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        .Add 0, vbKeyF12, conMenu_File_Parameter '��������
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add FCONTROL, vbKeyT, conMenu_View_Show '����
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF6, conMenu_View_Jump '��ת
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
'        .AddHiddenCommand conMenu_File_Excel '�����Excel
'        .AddHiddenCommand conMenu_View_Jump '��ת
    End With
            
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'���ܣ���Ϣ����
    Dim blnRecToLis As Boolean '�Ƿ���ص������б���
    Dim rsMsg As ADODB.Recordset
    
    If cboDept.ListIndex = -1 Then Exit Sub
    
    If Mid(mstrNotify, 1, 1) = "1" And strMsgItemIdentity = "ZLHIS_CHARGE_001" Then
        blnRecToLis = True
    ElseIf Mid(mstrNotify, 2, 1) = "1" And strMsgItemIdentity = "ZLHIS_CIS_004" Then
        blnRecToLis = True
    End If
    
    If blnRecToLis Then
        Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
        If rsMsg Is Nothing Then Exit Sub
        Call AddMsgToLis(rsMsg)
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.����ID
    End If
    
    Call ExecuteFindPati(False, , lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index: mstrFindType = objCard.����
End Sub

Private Sub picApplyInfo_Resize()
    rtfAppend.Width = picApplyInfo.ScaleWidth
    rtfAppend.Height = picApplyInfo.ScaleHeight - lblApply.Height
End Sub

Private Sub picApplyUD_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsExec.Width + X < 800 Or picApplyInfo.Width - X < 800 Then Exit Sub
        picApplyUD_S.Left = picApplyUD_S.Left + X
        vsExec.Width = vsExec.Width + X
        picApplyInfo.Left = picApplyInfo.Left + X
        picApplyInfo.Width = picApplyInfo.Width - X
        picBlood.Width = picBlood.Width + X
        Me.Refresh
    End If
End Sub

Private Sub fraUD_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picExec.Height + Y < 1500 Or tbcSub.Height - Y < 2000 Then Exit Sub
        fraUD_S.Top = fraUD_S.Top + Y
        picExec.Height = picExec.Height + Y
        tbcSub.Top = fraUD_S.Top + fraUD_S.Height
        tbcSub.Height = tbcSub.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub mclsEPRReport_AfterSaved(lngRecordId As Long)
    '��д����֮��ˢ���������
    Dim rsTemp As New ADODB.Recordset, lngҽ��ID As Long, strSQL As String
       
    On Error GoTo ErrHand
    strSQL = "Select ҽ��id,����״̬ From ����ҽ������ Where ����Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
    If rsTemp.RecordCount > 0 Then
        lngҽ��ID = Val("" & rsTemp!ҽ��ID)
        If Val("" & rsTemp!����״̬) = 1 Then
            strSQL = "Zl_������ļ�¼_Cancel(" & lngҽ��ID & "," & lngRecordId & ",Null)"
            Call zlDatabase.ExecuteProcedure(strSQL, "���²���״̬")
        End If
    End If
    
    mstrPrePati = ""
    Call rptPati_SelectionChanged
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mclsInAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    If Not RefreshNotify Then Call LoadPatients 'ע��Ҫ�ж�
End Sub

Private Sub mclsInAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsInAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclsInAdvices_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
'���ܣ����༭��ʽ��ӡ����
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr���Ʊ���, ����ID, Not Preview)
End Sub

Private Sub mclsInAdvices_ViewPACSImage(ByVal ҽ��ID As Long)
'���ܣ�PACS��Ƭ����
    With rptPati.SelectedRows(0)
        If CreateObjectPacs(gobjPublicPacs) Then
            Call gobjPublicPacs.ShowImage(ҽ��ID, Me, .Record(COL_����ת��).Value = 1)
        End If
    End With
End Sub

Private Sub mclsOutAdvices_RequestRefresh()
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    Call LoadPatients
End Sub

Private Sub mclsOutAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsExpenses_RequestRefresh()
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    Call LoadPatients
End Sub

Private Sub mclsExpenses_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
    '���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsOutAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclsOutAdvices_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
'���ܣ����༭��ʽ��ӡ����
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr���Ʊ���, ����ID, Not Preview)
End Sub

Private Sub mclsOutAdvices_ViewPACSImage(ByVal ҽ��ID As Long)
'���ܣ�PACS��Ƭ����
    With rptPati.SelectedRows(0)
        If CreateObjectPacs(gobjPublicPacs) Then
            Call gobjPublicPacs.ShowImage(ҽ��ID, Me, .Record(COL_����ת��).Value = 1)
        End If
    End With
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
'���ܣ����֤ʶ��ɹ��󼤻�
    mvarCond.���֤ = strID
    If mstrFindType = "�������֤" Then
        PatiIdentify.Text = mvarCond.���֤
    Else
        PatiIdentify.Text = "" '�������(Ŀǰ�������������²��ܼ���)��
    End If
    Call ExecuteFindPati(False, mvarCond.���֤)
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'���ܣ��Ķ���Ϣ��ɾ����Ϣ����˫����Ϣ����ѡ����Ϣ���ٰ��س�����
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng����ID As Long, lng��ҳID As Long
    Dim strҵ�� As String, strSQLRead As String
    Dim blnTmp As Boolean
    Dim strNO As String
    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim i As Long
 
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_��Ϣ).Value
                strҵ�� = .Item(C_ҵ��).Value
                lng����ID = Val(.Item(C_����ID).Value)
                lng��ҳID = Val(.Item(C_��ҳID).Value)
                lngIndex = .Index
            End With
            strSQLRead = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "',4,'" & UserInfo.���� & "'," & cboDept.ItemData(cboDept.ListIndex) & ")"
            
            If strNO = "ZLHIS_CHARGE_001" Then
                If Val(strҵ��) = 2 Then '������Դ������סԺ
                    '����������˽ӿڣ�����ڷ��ý�����Ҫ��ˢ�½���
                    blnTmp = mobjKernel.ChargeDelAudit(Me, mlngDept, lng����ID)
                    If tbcSub.Selected.Tag = "ҽ������" Then
                        Call mclsInAdvices_RequestRefresh(blnTmp)
                    End If
                End If
            End If
            If strNO = "ZLHIS_BLOOD_007" And gblnѪ��ϵͳ Then     'δ����ǰ��������Ϊ�Ѷ�
                If gobjPublicBlood Is Nothing And gblnѪ��ϵͳ Then InitObjBlood
                If gobjPublicBlood.zlIsBloodMessageDone(1, lng����ID, lng��ҳID, 4, cboDept.ItemData(cboDept.ListIndex)) Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                    Call rptNotify.Populate
                End If
                Exit Sub
            End If
            Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
            Call rptNotify.Records.RemoveAt(lngIndex)
            Call rptNotify.Populate
        End If
    End If
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim strCurPati As String
    Dim lngIndex As Long
    Dim strҵ��  As String
     
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    With rptNotify.SelectedRows(0)
        lngIndex = rptNotify.FocusedRow.Record.Index
        If rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_CIS_004" Or rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_BLOOD_007" Then
            strҵ�� = Val(rptNotify.Rows(lngIndex).Record(C_ҵ��).Value)
            If rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_BLOOD_007" Then strҵ�� = Split(rptNotify.Rows(lngIndex).Record(C_ҵ��).Value, ":")(1)
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    strCurPati = rptPati.SelectedRows(0).Record.Tag & "_"
                End If
            End If
            
            If InStr(strCurPati, strҵ�� & "_") = 0 Then
                If Not LocatePati(strҵ��) Then
                    Call LoadPatients
                    Call LocatePati(strҵ��)
                End If
            End If
        End If
    End With
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objRecord As ReportRecord
    
    '����Ŀ������Ŀ����ͬʱ��ѡ������������ģʽ
    If Row.Record.Childs.Count > 0 And Item.Checked Then
        For Each objRecord In Row.Record.Childs
            objRecord(col_ѡ��).Checked = False
        Next
    ElseIf Not Row.ParentRow.GroupRow And Item.Checked Then
        Row.ParentRow.Record(col_ѡ��).Checked = False
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Childs.Count > 0 And Not Row.GroupRow Then Row.Expanded = Not Row.Expanded
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "������ɫ" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ˢ���Ӵ�����漰����
'˵����������Ϊ�л����濨Ƭ����
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
    
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "ҽ������"
                Set objItem = tbcSub.InsertItem(Index, "ҽ�����ӷ���", mcolSubForm("_ҽ������").hwnd, 0)
                objItem.Tag = "ҽ������"
            Case "����ҽ��"
                Set objItem = tbcSub.InsertItem(Index, "����ҽ��", mcolSubForm("_����ҽ��").hwnd, 0)
                objItem.Tag = "����ҽ��"
            Case "סԺҽ��"
                Set objItem = tbcSub.InsertItem(Index, "סԺҽ��", mcolSubForm("_סԺҽ��").hwnd, 0)
                objItem.Tag = "סԺҽ��"
            Case "סԺ����"
                Set objItem = tbcSub.InsertItem(Index, "סԺ����", mcolSubForm("_סԺ����").hwnd, 0)
                objItem.Tag = "סԺ����"
            Case "�²���"
                Set objItem = tbcSub.InsertItem(Index, "���Ӳ���", mcolSubForm("_�²���").hwnd, 0)
                objItem.Tag = "�²���"
            Case "���ﲡ��"
                Set objItem = tbcSub.InsertItem(Index, "���ﲡ��", mcolSubForm("_���ﲡ��").hwnd, 0)
                objItem.Tag = "���ﲡ��"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "������"
                Set objItem = tbcSub.InsertItem(Index, "������", mcolSubForm("_������").hwnd, 0)
                objItem.Tag = "������"
            Case "�°滤��"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_�°滤��").hwnd, 0)
                objItem.Tag = "�°滤��"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
    'ˢ���Ӵ����Ӧ��CommandBar
    Call SubWinDefCommandBar(Item)
    
    'ˢ���Ӵ�������
    If Visible Then Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboDept_Click()
'���ܣ�ˢ�½�������
'˵�����Ӹ��¼���ʼ,�᲻�ظ�������ص����ݶ�ȡ
    Dim strDeptNode As String
    
    If cboDept.ListIndex = -1 Then Exit Sub
    cboDept.Tag = cboDept.ListIndex
    mblnReturn = True
    If Val(cboDept.ItemData(cboDept.ListIndex)) = mlngDept Then Exit Sub
    
    mlngDept = Val(cboDept.ItemData(cboDept.ListIndex))
    mblnѪ͸�� = Sys.DeptHaveProperty(mlngDept, "Ѫ͸��")
    mbln���� = Sys.DeptHaveProperty(mlngDept, "����")
    
    '���վ��仯����ı䲡�˲����б�
    strDeptNode = GetDeptNode(mlngDept)
    If strDeptNode <> mstrDeptNode Then
        mstrDeptNode = strDeptNode
        
        'ҽ�������Ƿ������ض�վ��ģ����֮ǰ����������ѡ����������վ��Ĳ��˿���
        If mvarCond.����ID <> 0 And mstrDeptNode <> "" Then
            strDeptNode = GetDeptNode(mvarCond.����ID)
            If strDeptNode <> mstrDeptNode Then mvarCond.����ID = 0
        End If
        
        Call LoadPatiUnit(mstrDeptNode)
        
    ElseIf Me.Visible = False Then
        '����ʱ����(mstrDeptNodeΪ��)
        Call LoadPatiUnit(mstrDeptNode)
    End If
    
    '���¶�ȡ����
    Call LoadPatients
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(col_�ۺ�״̬, "����", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_ѡ��, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = 5
        
        Set objCol = .Columns.Add(col_·��, "·��", 30, True): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_ִ��״̬, "״̬", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_ͼ��, "", 18, False): objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_��Դ, "��Դ", 30, False)
        Set objCol = .Columns.Add(col_���ݺ�, "���ݺ�", 65, True)
        Set objCol = .Columns.Add(col_����, "��", 20, True)
        Set objCol = .Columns.Add(col_����, "����", 55, True)
        Set objCol = .Columns.Add(col_����, "����", 150, True)
        Set objCol = .Columns.Add(col_����, "����", 60, True)
        Set objCol = .Columns.Add(col_����, "����", 65, True)
        Set objCol = .Columns.Add(col_��ʶ��, "��ʶ��", 62, True)
        Set objCol = .Columns.Add(col_����, "����", 35, True)
        Set objCol = .Columns.Add(col_�ѱ�, "�ѱ�", 55, True)
        Set objCol = .Columns.Add(col_Ҫ��ʱ��, "Ҫ��ʱ��", 106, True)
        Set objCol = .Columns.Add(col_����ʱ��, "����ʱ��", 106, True)
        Set objCol = .Columns.Add(col_ִ�м�, "ִ�м�", 65, True)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(col_����, "����", 30, True)
        Set objCol = .Columns.Add(col_�����, "�����", 55, True)
        Set objCol = .Columns.Add(col_���ʱ��, "���ʱ��", 106, True)
        
        '����������
        Set objCol = .Columns.Add(col_ִ�п���, "ִ�п���", 0, False)
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False)
        Set objCol = .Columns.Add(col_��ҳID, "��ҳID", 0, False)
        Set objCol = .Columns.Add(col_�Һŵ�, "�Һŵ�", 0, False)
        Set objCol = .Columns.Add(col_�Һ�ID, "�Һ�ID", 0, False)
        Set objCol = .Columns.Add(col_Ӥ��, "Ӥ��", 0, False)
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_���￨��, "���￨��", 0, False)
        Else
            Set objCol = .Columns.Add(col_���￨��, "���￨��", 70, True)
        End If
        Set objCol = .Columns.Add(col_���֤��, "���֤��", 0, False)
        Set objCol = .Columns.Add(col_IC����, "IC����", 0, False)
        Set objCol = .Columns.Add(col_ҽ����, "ҽ����", 0, False)
        Set objCol = .Columns.Add(col_����id, "����ID", 0, False)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 0, False)
        Set objCol = .Columns.Add(COL_״̬, "״̬", 0, False)
        Set objCol = .Columns.Add(col_ҽ��ID, "ҽ��ID", 0, False)
        Set objCol = .Columns.Add(col_���ID, "���ID", 0, False)
        Set objCol = .Columns.Add(col_���ͺ�, "���ͺ�", 0, False)
        Set objCol = .Columns.Add(COL_�������, "�������", 0, False)
        Set objCol = .Columns.Add(col_ִ�й���, "ִ�й���", 0, False)
        Set objCol = .Columns.Add(col_ִ�а���, "ִ�а���", 0, False)
        Set objCol = .Columns.Add(col_��¼����, "��¼����", 0, False)
        Set objCol = .Columns.Add(COL_����ת��, "����ת��", 0, False)
        Set objCol = .Columns.Add(col_�ļ�ID, "�ļ�ID", 0, False)
        Set objCol = .Columns.Add(col_������, "������", 0, False)
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(col_��������, "��������", 80, True)
        Set objCol = .Columns.Add(col_�������, "�������", 0, False)
        
        Set objCol = .Columns.Add(col_������, "������", 65, True)
        Set objCol = .Columns.Add(col_���, "���", 16, False)
        objCol.TreeColumn = True: objCol.Visible = False
        Set objCol = .Columns.Add(col_����, "����", 16, False)
        objCol.TreeColumn = True: objCol.Visible = False
        Set objCol = .Columns.Add(COL_��������, "��������", 0, False)
        Set objCol = .Columns.Add(COL_�˶���, "�˶���", 0, False)
        Set objCol = .Columns.Add(col_��˱�־, "��˱�־", 0, False)
        Set objCol = .Columns.Add(COL_����ʱ��, "����ʱ��", 0, False)
        Set objCol = .Columns.Add(col_����ģʽ, "����ģʽ", 0, False)
        Set objCol = .Columns.Add(COL_������ĿID, "������ĿID", 0, False)
        Set objCol = .Columns.Add(col_��Ч, "��Ч", 0, False)
        Set objCol = .Columns.Add(COL_ִ�з���, "ִ�з���", 0, False)
        Set objCol = .Columns.Add(COL_��ҳ�Һ�ID, "��ҳ�Һ�ID", 0, False)
        Set objCol = .Columns.Add(COL_���ӱ�־, "���ӱ�־", 0, False)
        For Each objCol In .Columns
            If objCol.Index <> col_ѡ�� Then objCol.Editable = False
            objCol.Groupable = objCol.Index = col_�ۺ�״̬
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(col_�ۺ�״̬)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_����ʱ��)
        .SortOrder(0).SortAscending = False
    End With
    
    
    With rptNotify
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
         
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_��� Or objCol.Index <> C_���� Then objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û����������..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        
        '���� ����
        .SortOrder.Add .Columns(C_���)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns(C_����)
        .SortOrder(1).SortAscending = False
    End With
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picMsg.hwnd
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    With Me.picExec
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = fraUD_S.Top - lngTop
    End With
    With Me.fraUD_S
        .Left = lngLeft
         If Not mblnFirstLoad Then .Top = lngTop + picExec.Height
        .Width = lngRight - lngLeft
    End With
    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = fraUD_S.Top + fraUD_S.Height: .Height = lngBottom - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim blnSetup As Boolean
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ������", tbcSub.Selected.Tag, glngSys, pҽ������վ, blnSetup)
    End If
    Call zlDatabase.SetPara("ִ��״̬", mstr״̬, glngSys, pҽ������վ, blnSetup)
    Call zlDatabase.SetPara("ֻ��ʾ���շѵĲ���", IIf(mblnֻ�����շ�, 1, 0), glngSys, pҽ������վ, blnSetup)
    Call zlDatabase.SetPara("���˲��ҷ�ʽ", mintFindType, glngSys, pҽ������վ, blnSetup)
    Call zlDatabase.SetPara("������ʾģʽ", IIf(mblnFilter, 1, 0), glngSys, pҽ������վ, blnSetup)
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(DkpMain), DkpMain.Name, DkpMain.SaveStateToString)
    End If
    Call zlDatabase.SetPara("����", mbytSize, glngSys, pҽ������վ, blnSetup)
    If Me.Visible Then
        '���������̶�����һ���ؼ�����ʽ���棬����վ���������һ���Ǵ�ӡ����̶���ͼ����ʽ,������ָ�Ϊ������ť����ʽ
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If
    mblnIsInit = False
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing
    
    Unload frmTechnicFilter
    
    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsExpenses = Nothing
    Set mclsEMR = Nothing
    Set mfrmActive = Nothing
    Set mclsInEPRs = Nothing
    Set mclsOutEPRs = Nothing
    Set mclsEPRReport = Nothing
    Set gobjPublicPacs = Nothing
    
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Set mobjKernel = Nothing
    Set mclsTends = Nothing
    Set mclsTendsNew = Nothing
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
    Set mclsPExp = Nothing
    Set mobjAppendBill = Nothing
    If Not mobjFrmBloodExe Is Nothing Then
        Unload mobjFrmBloodExe
        Set mobjFrmBloodExe = Nothing
    End If
End Sub

Private Sub picExec_Resize()
    On Error Resume Next
    
    fraDiag.Left = 0
    fraDiag.Top = 0
    fraDiag.Width = picExec.ScaleWidth

    fraExec.Left = 0
    fraExec.Width = picExec.ScaleWidth
    fraExec.Top = fraDiag.Top + fraDiag.Height - Screen.TwipsPerPixelY * 6
    
    lblAdvice.Width = fraExec.Width - lblCash.Width - lblRec.Width - lblAdvice.Left - Screen.TwipsPerPixelX * 4
    lblCash.Left = fraExec.Width - lblCash.Width - Screen.TwipsPerPixelX * 4
    
    lblRec.Top = lblCash.Top
    lblRec.Left = lblCash.Left - lblRec.Width - 10
    
    With Me.picApplyUD_S
        .Top = fraExec.Top + fraExec.Height
        .Height = picExec.ScaleHeight - vsExec.Top
    End With
    
    vsExec.Left = 0
    vsExec.Top = fraExec.Top + fraExec.Height
    vsExec.Width = IIf(Me.picApplyInfo.Visible, picApplyUD_S.Left, picExec.Width)
    vsExec.Height = picExec.ScaleHeight - vsExec.Top
    
    picBlood.Left = 0
    picBlood.Top = vsExec.Top
    picBlood.Width = vsExec.Width
    picBlood.Height = vsExec.Height
    If picBlood.Tag = "�ɼ�" Then
        vsExec.Visible = False
        picBlood.Visible = True
    Else
        vsExec.Visible = True
        picBlood.Visible = False
    End If
    
    picApplyInfo.Move vsExec.Width + picApplyUD_S.Width, vsExec.Top, _
        picExec.ScaleWidth - vsExec.Width - picApplyUD_S.Width, vsExec.Height
    picApplyUD_S.Visible = picApplyInfo.Visible
    
    fraUD_S.Top = picExec.Top + picExec.Height
    
    tbcSub.Top = fraUD_S.Top + fraUD_S.Height
    
End Sub

Private Sub picPati_GotFocus()
    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    cboDept.Top = 30
    lblDept.Top = (cboDept.Height - lblDept.Height) / 2 + cboDept.Top
    lblDept.Left = lblDept.Top
    cboDept.Left = lblDept.Left + lblDept.Width + 30
    cboDept.Width = picPati.ScaleWidth - cboDept.Left - lblDept.Left
    
    If cboUnit.Visible Then
        cboUnit.Top = cboDept.Top + cboDept.Height + 45
        lblUnit.Top = (cboUnit.Height - lblUnit.Height) / 2 + cboUnit.Top
        lblUnit.Left = lblDept.Left
        cboUnit.Left = lblUnit.Left + lblUnit.Width + 30
        cboUnit.Width = cboDept.Width
    End If
    
    If Val(gstrҽ���˶�) = 0 Then
        fraFilter.Height = chkִ��״̬(3).Top + chkִ��״̬(3).Height + IIf(mbytSize = 0, 90, 250)
    Else
        fraFilter.Height = chkִ��״̬(4).Top + chkִ��״̬(4).Height + IIf(mbytSize = 0, 90, 250)
    End If
    
    lblFind.Top = IIf(cboUnit.Visible, cboUnit.Top + cboUnit.Height, cboDept.Top + cboDept.Height) + 100
    PatiIdentify.Top = lblFind.Top - 50
    PatiIdentify.Width = cboDept.Left + cboDept.Width - PatiIdentify.Left - chkFilter.Width - 60
    chkFilter.Left = PatiIdentify.Left + PatiIdentify.Width + 30
    chkFilter.Top = PatiIdentify.Top
    
    rptPati.Left = 0
    rptPati.Top = PatiIdentify.Top + PatiIdentify.Height + 30
    rptPati.Width = picPati.ScaleWidth
    rptPati.Height = picPati.ScaleHeight - rptPati.Top - fraFilter.Height
    
    fraFilter.Left = 30
    fraFilter.Top = rptPati.Top + rptPati.Height
    fraFilter.Width = rptPati.Width - 45
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objColumn As ReportColumn
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        'ȫѡ����
        Set objColumn = rptPati.Columns.Find(col_ѡ��)
        If objColumn.Visible Then
            objColumn.Caption = "1"
            Call SelectALLPati(True)
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        'ȫ�岡��
        Set objColumn = rptPati.Columns.Find(col_ѡ��)
        If objColumn.Visible Then
            objColumn.Caption = ""
            Call SelectALLPati(False)
        End If
    ElseIf KeyCode = vbKeyTab Then
        'Panne�е�Report�ؼ���Ҫǿ�д�����˳��
        '������ʱ���ܲ���vbKeyTab
        If Shift = vbShiftMask Then
            If cboDept.Enabled Then cboDept.SetFocus
        Else
            If vsExec.Enabled Then vsExec.SetFocus
        End If
    ElseIf KeyCode = vbKeySpace Then
        '����ѡ
        If rptPati.Columns(col_ѡ��).Visible Then
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    rptPati.SelectedRows(0).Record(col_ѡ��).Checked = Not rptPati.SelectedRows(0).Record(col_ѡ��).Checked
                    Call rptPati_ItemCheck(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record(col_ѡ��))
                    rptPati.Redraw
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim objColumn As ReportColumn
        
    If Button = 2 Then
        Set objHitTest = rptPati.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
                With objPopup.Controls
                    .Add(xtpControlSplitButtonPopup, conMenu_View_Show, "���˹���(&O)").IconId = conMenu_View_Filter
                    .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)").BeginGroup = True
                    .Add xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)"
                    .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)").BeginGroup = True
                    .Add xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)"
                End With
            Else
                Set objPopup = cbsMain.ActiveMenuBar.Controls(2).CommandBar
            End If
        End If
        
        rptPati.SetFocus
        If Not objPopup Is Nothing Then objPopup.ShowPopup
    ElseIf Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        Call SelectALLPati(True)
                    Else
                        objColumn.Caption = ""
                        Call SelectALLPati(False)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SelectALLPati(ByVal blnSelect As Boolean)
    Dim objParent As ReportRow, i As Long
    
    If rptPati.Columns(col_ѡ��).Visible And rptPati.SelectedRows.Count > 0 Then
        '��������м�¼��ѡ��״̬
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(col_ѡ��).Checked = False
        Next
        
        '��ǰ����
        Set objParent = rptPati.SelectedRows(0)
        If Not objParent.GroupRow Then
            If objParent.ParentRow.GroupRow Then
                Set objParent = objParent.ParentRow
            Else
                Set objParent = objParent.ParentRow.ParentRow
            End If
        End If
        
        'ֻ��Կɼ���Ч�н��д���
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If Not rptPati.Rows(i).ParentRow.GroupRow Then
                    '����������ȫѡ/ȫ��ʱ���ֲ�ѡ״̬
                    rptPati.Rows(i).Record(col_ѡ��).Checked = False
                ElseIf rptPati.Rows(i).ParentRow Is objParent Then
                    'ͬһ����ĲŴ���
                    rptPati.Rows(i).Record(col_ѡ��).Checked = blnSelect
                End If
            End If
        Next
        rptPati.Redraw
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    Dim rsTmp As New ADODB.Recordset, bln�� As Boolean
    Dim strCurPati As String, blnChange As Boolean
    Dim blnUseBlood As Boolean
    If rptPati.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    With rptPati.SelectedRows(0)
        picBlood.Tag = ""
        If Not .GroupRow Then strCurPati = .Record.Tag
        If strCurPati = mstrPrePati Then Exit Sub
        Me.stbThis.Panels(2).Text = ""
        mstrPrePati = strCurPati
        
        If Not .GroupRow Then
            mlng����ID = .Record(col_����Id).Value
            mlng��ҳID = .Record(col_��ҳID).Value
            mstr�Һŵ� = .Record(col_�Һŵ�).Value
            '����
            mlngType = rptPati.SelectedRows(0).Record(col_����).Value
            If (.Record(col_��Դ).Value & "" = "סԺ") Then
                '3-סԺ����
                If GetInsidePrivs(pסԺ��������, True) <> "" Then
                    Me.tbcSub.Item(mlngInIndex).Visible = True
                End If
                If GetInsidePrivs(p���ﲡ������, True) <> "" Then
                    Me.tbcSub.Item(mlngOutIndex).Visible = False
                End If
                
                If mlngNewIndex <> -1 Then
                    If GetInsidePrivs(p�°�סԺ����, True) <> "" Then
                        Me.tbcSub.Item(mlngNewIndex).Visible = True
                    Else
                        Me.tbcSub.Item(mlngNewIndex).Visible = False
                    End If
                End If
                
                If GetInsidePrivs(p�����¼����, True) <> "" Then
                    If mblnѪ͸�� Or mbln���� Then
                        If mblnNewNurRecord Then
                            Me.tbcSub.Item(mlngNewNurIndex).Visible = True
                            Me.tbcSub.Item(mlngNurEMRIndex).Visible = True
                        Else
                            Me.tbcSub.Item(mlngNurIndex).Visible = True
                        End If
                    Else
                        Me.tbcSub.Item(mlngNurIndex).Visible = False
                        Me.tbcSub.Item(mlngNewNurIndex).Visible = False
                        Me.tbcSub.Item(mlngNurEMRIndex).Visible = False
                        '���غ�Ҫ�ٶ�λ
                        If tbcSub.Selected.Tag = "����" Or tbcSub.Selected.Tag = "�°滤��" Then
                            tbcSub.Item(0).Selected = True
                        End If
                    End If
                End If
                
            Else
                '4-���ﲡ��
                If GetInsidePrivs(p���ﲡ������, True) <> "" Then
                    Me.tbcSub.Item(mlngOutIndex).Visible = True
                End If
                If GetInsidePrivs(pסԺ��������, True) <> "" Then
                    Me.tbcSub.Item(mlngInIndex).Visible = False
                End If
                If mlngNewIndex <> -1 Then
                    If GetInsidePrivs(p�°����ﲡ��, True) <> "" Then
                        Me.tbcSub.Item(mlngNewIndex).Visible = True
                    Else
                        Me.tbcSub.Item(mlngNewIndex).Visible = False
                    End If
                End If
                If GetInsidePrivs(p�����¼����, True) <> "" Then
                    Me.tbcSub.Item(mlngNurIndex).Visible = False
                     Me.tbcSub.Item(mlngNewNurIndex).Visible = False
                    '���غ�Ҫ�ٶ�λ
                    If tbcSub.Selected.Tag = "����" Or tbcSub.Selected.Tag = "�°滤��" Then
                        tbcSub.Item(0).Selected = True
                    End If
                End If
            End If
            mlngState = IIf(IsNull(.Record(col_��Ժ����).Value), IIf(.Record(COL_״̬).Value + 0 = 3, psԤ��, ps��Ժ), mlngType)
            
            '��ȡ����ID
            If .Record(col_����ID).Value = 0 Then
                Call ReadMoreInfo
            End If
            
            '��ʾҽ������
            lblAdvice.Caption = Getִ������(.Record(col_���ͺ�).Value, .Record(col_ҽ��ID).Value, .Record(col_���ID).Value, .Record(COL_�������).Value, rptPati.SelectedRows(0))
            
            '��ʾ�������
            If .Record(col_��Դ).Value <> "סԺ" Then
                lblDiag(1).Caption = GetPatiDiagnose(.Record(col_����Id).Value, .Record(col_�Һ�ID).Value, 1)
            Else
                lblDiag(1).Caption = GetPatiDiagnose(.Record(col_����Id).Value, .Record(col_��ҳID).Value, 2)
            End If
            
            '�Ƿ����շ�
            If .Record(col_����).Value <> "" And Val(.Record(col_��¼����).Value) = 1 Then
                'ҽ������
                bln�� = Set�շѱ��(IIf(.Record(col_��Դ).Value = "סԺ", 2, 1), Not .ParentRow.GroupRow, _
                    .Record(col_ҽ��ID).Value, .Record(col_���ID).Value, .Record(col_���ͺ�).Value, .Record(COL_�������).Value, _
                    .Record(col_���ݺ�).Value, .Record(col_��¼����).Value, .Record(col_�������).Value, .Record(COL_����ת��).Value = 1, .Record(col_����ʱ��).Value)
            Else
                bln�� = ItemHaveCash(IIf(.Record(col_��Դ).Value = "סԺ", 2, 1), Not .ParentRow.GroupRow, _
                    .Record(col_ҽ��ID).Value, .Record(col_���ID).Value, .Record(col_���ͺ�).Value, .Record(COL_�������).Value, _
                    .Record(col_���ݺ�).Value, .Record(col_��¼����).Value, .Record(col_�������).Value, 1, .Record(COL_����ת��).Value = 1, .Record(col_����ʱ��).Value)
            End If
            lblCash.Visible = bln��
            lblRec.Visible = Val(.Record(col_����ģʽ).Value) = 1
            
             '�˴��ж��Ƿ�����Ѫҽ��
            blnUseBlood = False
            If gblnѪ��ϵͳ = True And .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "8" And .Record(COL_ִ�з���).Value = "1" Then
                blnUseBlood = True
            End If
            
            If blnUseBlood = False Then
                picBlood.Tag = ""
                '��ʾִ�����
                Call LoadExecList(.Record(col_ҽ��ID).Value, .Record(col_���ͺ�).Value)
            Else
                picBlood.Tag = "�ɼ�"
                If Not mobjFrmBloodExe Is Nothing Then
                    Call mobjFrmBloodExe.zlRefresh(Me, glngSys, pҽ������վ, Val(.Record(col_ҽ��ID).Value), mlngDept, GetInsidePrivs(pҽ������վ), 1, mlngDept, .Record(COL_����ת��).Value = 1, IIf(mbytSize = 0, 9, 12))
                End If
            End If
            
            '�л���ʾ��ͬ��ҽ���Ӵ���
            blnChange = ExchangeAdvice(.Record(col_��Դ).Value <> "סԺ")
        Else
            Call ClearPatiInfo
        End If
        
        '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        Call ShowBillList(cbsMain.FindControl(, conMenu_Manage_RequestPrint, , True))
        
        'ˢ���Ӵ�������
        If Not blnChange Then
            Call SubWinRefreshData(tbcSub.Selected)
        End If
        If Not rptPati.SelectedRows(0).GroupRow Then
            Call ShowBillAppend(1, True)
            picApplyInfo.Visible = Not (rtfAppend.Text = "")
            picApplyUD_S.Visible = picApplyInfo.Visible
            Call picExec_Resize
        Else
            picApplyInfo.Visible = False
            rtfAppend.Text = ""
            picApplyUD_S.Visible = False
            Call picExec_Resize
        End If
    End With
End Sub

Private Sub ReadMoreInfo()
'���ܣ���ȡ������صĸ�����Ϣ��ΪЧ�ʲ�����SQL�ж�ȡ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngҽ��ID As Long
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        '��ȡ�Һ�ID
        .Record(col_�Һ�ID).Value = 0
        If .Record(col_�Һŵ�).Value <> "" Then
            strSQL = "Select ID,���ӱ�־ From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(.Record(col_�Һŵ�).Value))
            If Not rsTmp.EOF Then
                .Record(col_�Һ�ID).Value = rsTmp!ID
                .Record(COL_���ӱ�־).Value = rsTmp!���ӱ�־ & ""
            End If
        ElseIf .Record(COL_��ҳ�Һ�ID).Value <> "" Then
            strSQL = "Select ID,���ӱ�־ From ���˹Һż�¼ Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(COL_��ҳ�Һ�ID).Value))
            If Not rsTmp.EOF Then
                .Record(COL_���ӱ�־).Value = rsTmp!���ӱ�־ & ""
            End If
        End If
        
        '��ȡ����ID
        .Record(col_����ID).Value = 0
        If (.Record(COL_�������).Value = "C" Or .Record(COL_�������).Value = "D") And .Record(col_���ID).Value <> 0 Then
            lngҽ��ID = .Record(col_���ID).Value '�������ȡ���ID
        Else
            lngҽ��ID = .Record(col_ҽ��ID).Value
        End If
    
        strSQL = "Select ����ID From ����ҽ������ Where ҽ��ID=[1]"
        If .Record(COL_����ת��).Value = 1 Then
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
        If Not rsTmp.EOF Then '�����ж����Ŀǰֻ��һ��
            .Record(col_����ID).Value = Val(rsTmp!����id & "")
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    On Error GoTo errH
    
    '��������/סԺҽ������
    str��Դ = "3"
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "1,2,3"
    ElseIf InStr(mstrPrivs, "���ﲡ��") > 0 Then
        str��Դ = "1,3"
    ElseIf InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "2,3"
    End If
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(" & str��Դ & ") And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(" & str��Դ & ") And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    End If
    
    cboDept.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    str����IDs = GetUser����IDs
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.����ID Then
            Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex) 'ֱ����������
        End If
        If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then
            Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
        End If
        
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
        
    If cboDept.ListIndex <> -1 Then
        Call cboDept_Click  'ͬʱ��mstrDeptNode��ֵ
    End If
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadPatiUnit(ByVal strDeptNode As String)
'���ܣ���ȡ�����ز��˲���
'   strDeptNode=��ǰҽ������������վ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngUnit As Long

    On Error GoTo errH
    If cboUnit.ListIndex > 0 Then lngUnit = Val(cboUnit.ItemData(cboUnit.ListIndex))
    
    cboUnit.Clear
    cboUnit.AddItem "���в���"
    Call Cbo.SetIndex(cboUnit.hwnd, 0)
    
    '��Դ���Ÿ��ݵ�ǰҽ�����ҵ�վ��������վ��
    strSQL = "Select A.ID,A.����,A.���� From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(1,2,3) And B.��������='����'" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        IIf(strDeptNode <> "", " And (A.վ�� = [1] Or A.վ�� is Null)", "") & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDeptNode)
    Do While Not rsTmp.EOF
        cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
        cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngUnit Then Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
        rsTmp.MoveNext
    Loop
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetDeptNode(ByVal lngDept As Long) As String
'���ܣ���ȡָ������������վ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select վ�� From ���ű� Where ID = [1] And վ�� is not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡվ��", lngDept)
    If rsTmp.RecordCount > 0 Then GetDeptNode = rsTmp!վ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function LoadPatients() As Boolean
'���ܣ���ȡ�����б�
    Dim rsPati As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Dim objRow As ReportRow, objPreRecord As ReportRecord
    Dim strPatiRow As String, lngPatiRow As Long, strExpend As String
    
    Dim strSQL As String, strSQL1 As String
    Dim str������ As String, str���� As String
    Dim blnDateMoved As Boolean, str��Դ As String, str������Դ As String
    Dim datBegin As Date, datEnd As Date
    Dim curDate As Date, lng����ID As Long
    Dim blnDo As Boolean, blnSub As Boolean, blnPath As Boolean
    Dim lngColor As Long, i As Long, j As Long
    Dim blnNoFilter As Boolean
    Dim strWhere�˶� As String, strBloodWhere As String
    Dim str�շ��ж� As String
    
    '����ģʽʱ���޲��˲����������������
    If mblnFilter Then
        If mvarCond.IC���� = "" And mvarCond.NO = "" And mvarCond.��ʶ�� = "" _
            And mvarCond.���￨ = "" And mvarCond.���֤ = "" And mvarCond.���� = "" And mvarCond.ҽ���� = "" And mvarCond.����ID = 0 Then
            rptPati.Records.DeleteAll
            rptPati.Populate
            Call ClearPatiInfo
            Call SubWinRefreshData(tbcSub.Selected)
            LoadPatients = True: Exit Function
        End If
    End If
    
    If Not mblnFilter And Val(PatiIdentify.Text) = 0 And mvarCond.������ = "" And mstr������ <> "" Then
        blnNoFilter = True
        mvarCond.������ = mstr������
    End If
    
    '��ҳ����������գ�F5ˢ�£�Ӧ�ûָ���һ����ֵ
    If cboUnit.ListIndex = -1 Then Call Cbo.SetIndex(cboUnit.hwnd, Val(cboUnit.Tag))
    If cboDept.ListIndex = -1 Then Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
            
    mblnShowBed = False
            
    '��ѯʱ���
    curDate = zlDatabase.Currentdate
    If mvarCond.Begin = CDate(0) Then
        datBegin = Int(curDate - 1)
    Else
        datBegin = mvarCond.Begin
    End If
    If mvarCond.End = CDate(0) Then
        datEnd = Format(curDate, "yyyy-MM-dd 23:59")
    Else
        datEnd = mvarCond.End
    End If
    blnDateMoved = zlDatabase.DateMoved(datBegin) '��ʱ�俴�Ƿ������ת��
    
    '������ԴȨ��:(1-����,2-סԺ,3-����,4-���),��첡��Ϊ���ﲡ��
    '���������۷��͵�סԺʱ��ҽ����¼�еĲ�����Դ��Ϊ2
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "1,2,3,4"
        str������Դ = ""
    ElseIf InStr(mstrPrivs, "���ﲡ��") > 0 Then
        str��Դ = "1,4"
        str������Դ = " And (Instr([2],','||A.������Դ||',')>0 Or f.�������� = 1 And a.������Դ = 2)"
    ElseIf InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "2"
        str������Դ = " And Instr([2],','||A.������Դ||',')>0 And f.�������� <> 1"
    Else
        str��Դ = "3"
        str������Դ = " And Instr([2],','||A.������Դ||',')>0"
    End If
    
    If mvarCond.��Դ <> "" Then
        If InStr(mvarCond.��Դ, "1") > 0 And InStr(mvarCond.��Դ, "2") = 0 Then
            str������Դ = str������Դ & " And (Instr([12],','||a.������Դ||',')>0 Or f.�������� = 1 And a.������Դ = 2)"
        ElseIf InStr(mvarCond.��Դ, "1") = 0 And InStr(mvarCond.��Դ, "2") > 0 Then
            str������Դ = str������Դ & " And Instr([12],','||a.������Դ||',')>0 And f.�������� <> 1"
        Else
            str������Դ = str������Դ & " And Instr([12],','||a.������Դ||',')>0"
        End If
    End If
    
    If cboUnit.ListIndex <> -1 Then
        lng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    End If
    
    '���¹����뷢��ʱ��NO�ֺŶ�Ӧ��
    '���������:������Ŀֻ��ʾһ��,�ɼ�������ʾһ��
    '��ҩ�巨,�÷����Զ�����ʾһ��
    '��������,��鲿λִ�п��Ҽ�ʱ��������Ŀ��ͬ,����ʾ
    '��������ִ�п���Ϊ��������Ҫ��ʾ
    '��Ѫ��Ŀ����Ѫ;���ֱ�ִ��
    '����ҽ������ʾ(��Ȼִ�п���һ�㲻��Ϊҽ������)
    If mbln����ִ�� Then
        If gblnѪ��ϵͳ = True Then
            '��Ѫҽ����ҽ��ʱ��Ĭ��3���ڵĶ�������ʾ
            strBloodWhere = " And ((Nvl(a.ִ��״̬, 0) = 0 And Trunc(a.����ʱ��) = Trunc(Sysdate)) Or (Exists" & vbNewLine & _
                                       "              (Select 1 From ������ĿĿ¼ Where b.������Ŀid = Id And ��� || �������� || ִ�з��� = 'E81') And" & vbNewLine & _
                                       "              (Nvl(a.ִ��״̬, 0) = 0 Or Exists" & vbNewLine & _
                                       "               (Select 1" & vbNewLine & _
                                       "                 From ѪҺ���ͼ�¼ a, ѪҺִ�м�¼ b, ѪҺ��Ѫ��¼ c" & vbNewLine & _
                                       "                 Where a.�շ�id = b.�շ�id(+) And a.�䷢id = c.Id��and c.����id = b.���id��and(Nvl(a.ִ��״̬, 0) = 0 Or b.ִ�п���id = [1])))))"
        Else
            strBloodWhere = " And Nvl(A.ִ��״̬,0)=0 And Trunc(A.����ʱ��)=Trunc(Sysdate) "
        End If
        
        str���� = "" & _
            " And (A.ִ�в���ID+0=[1] Or A.ִ�в���ID+0<>[1]  " & strBloodWhere & _
            " And Exists(Select 1 From ������ĿĿ¼ C,����ִ�п��� D Where C.ID=B.������ĿID And C.ִ�п���=4 And C.ID=D.������ĿID And D.ִ�п���ID=[1])" & _
            " And Exists(Select 1 From ���ű� C,��������˵�� D Where C.ID=A.ִ�в���ID And C.ID=D.����ID" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) And D.�������� IN('���','����','����','����','Ӫ��'))" & _
            " And Exists(Select 1 From ��������˵�� C Where C.����ID=[1] And C.�������� IN('���','����','����','����','Ӫ��')" & _
            " And (C.�������=Decode(B.������Դ,4,1,3,1,B.������Դ) Or C.�������=3))  )"
    Else
        str���� = " And A.ִ�в���ID+0=[1]"
    End If
    If mstr������� <> "" Then
        str���� = str���� & " And Instr('" & mstr������� & "',B.�������)>0"
    End If
    If mstr������� <> "" Then
        str���� = str���� & " And (Not B.�������='E' Or B.�������='E' And Exists(Select 1 From ������ĿĿ¼ C Where C.ID=B.������ĿID And Instr('" & mstr������� & "',C.��������)>0))"
    End If
    
'    str������ = "Select 1 From ����ҽ����¼ C,����ҽ������ D" & _
'        " Where ((C.�������='C' And C.���ID=B.���ID) or (C.�������='D' And (C.���ID=B.���ID or C.ID=B.���ID or C.���ID=B.ID)))" & _
'        " And C.ID=D.ҽ��ID And D.���ͺ�=A.���ͺ�  And D.ҽ��ID=A.ҽ��ID And D.ִ��״̬ IN([3],[4],[5],[6])"
    
    If mblnֻ�����շ� Then
        str�շ��ж� = " And (A.��¼���� <> 1 Or A.��¼���� = 1 And a.�Ʒ�״̬ = 3 Or " & _
        " A.��¼���� = 1 And a.�Ʒ�״̬ in (-1,0) And Exists ( Select 1 From ����ҽ����¼ C,����ҽ������ D" & _
        " Where ((C.�������='C' And C.���ID=B.���ID) or (C.�������='D' And (C.���ID=B.���ID or C.ID=B.���ID or C.���ID=B.ID)))" & _
        " And C.ID=D.ҽ��ID And D.���ͺ�=A.���ͺ� And D.ִ��״̬ IN([3],[4],[5],[6]) And D.�Ʒ�״̬ = 3) )"
    End If

    strSQL = _
        " Select B.����,B.����,B.�Ա�, A.ҽ��ID,A.���ͺ�,B.���ID,B.���,B.�������,b.ҽ����Ч,B.������ĿID,A.����ʱ��,A.NO,a.������," & _
        "       A.����ʱ��,Nvl(A.����ʱ��,Decode(Nvl(B.ҽ����Ч,0),1,B.��ʼִ��ʱ��,A.�״�ʱ��)) as Ҫ��ʱ��," & _
        "       A.��¼����,A.�������,A.ִ��״̬,A.ִ�й���,A.ִ�в���ID,A.�����,A.���ʱ��,A.����ʱ��," & _
        "       B.����ID,B.��ҳID,B.�Һŵ�,B.Ӥ��,B.���˿���ID,B.������Դ,A.ִ�м�,0 as ����ת��,b.ҽ������,b.�걾��λ,b.��鷽��,b.����ҽ��, B.������־" & _
        " From ����ҽ������ A,  ����ҽ����¼ B,�������� C" & _
        " Where A.ҽ��ID=B.ID And B.�շ�ϸĿID=C.����ID(+) And B.������� Not IN('5','6','7')" & _
                str���� & _
        IIf(Mid(mstr״̬, 1, 4) = "1111", "", " And A.ִ��״̬ IN([3],[4],[5],[6]) ") & str�շ��ж� & _
        IIf(mstrRoom <> "", " And Instr([7],'|'||A.ִ�м�||'|')>0", "") & _
        "       And A." & mstr�������� & " Between [8] And [9]" & _
                IIf(mvarCond.NO <> "", " And A.NO=[10]", "") & _
                IIf(mvarCond.����ID <> 0, " And B.���˿���ID+0=[11]", "") & _
                IIf(mvarCond.��Ч <> 0, " And Nvl(B.ҽ����Ч,0)=[19]", "") & _
                IIf(mvarCond.������ <> "", " And B.����ҽ��=[21]", "") & _
                IIf(mvarCond.����ID <> 0, " And B.����ID=[22]", "")
            
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "0 as ����ת��", "1 as ����ת��")
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    If Not (Mid(mstr״̬, 3, 1) = "0" Or Mid(mstr״̬, 5, 1) = "1" Or Val(gstrҽ���˶�) = 0 Or Mid(mstr״̬, 5, 1) = "") Then
        If Val(gstrҽ���˶�) = 11 Then
            strWhere�˶� = " And (A.ִ��״̬ <> 3 or A.ִ��״̬ = 3 And (Not (C.�������� in ('1','8') And a.�������='E' Or a.�������='K') Or (C.�������� in ('1','8') And a.�������='E' Or a.�������='K') and a.������ is null))"
        ElseIf Mid(gstrҽ���˶�, 2, 1) = "1" Then
            strWhere�˶� = " And (A.ִ��״̬ <> 3 or A.ִ��״̬ = 3 And (Not C.��������='1' And a.������� = 'E' Or C.��������='1' And a.������� = 'E' and a.������ is null))"
        ElseIf Mid(gstrҽ���˶�, 1, 1) = "1" Then
            strWhere�˶� = " And (A.ִ��״̬ <> 3 or A.ִ��״̬ = 3 And (Not (C.��������='8' And a.�������='E' Or a.�������='K') Or (C.��������='8' And a.�������='E' Or a.�������='K') and a.������ is null))"
        End If
    End If
    
    strSQL = _
    " Select /*+ RULE */ DISTINCT" & vbNewLine & _
    "   a.ҽ��id, a.���ͺ�, a.���id, a.���, a.�������,c.��������,c.ִ�з���,a.ҽ����Ч as ��Ч, a.������Ŀid, a.����ʱ��, a.Ҫ��ʱ��, a.����ʱ��, a.No, a.��¼����, a.ִ��״̬, a.ִ�й���, a.ִ�в���id, a.����id,a.������," & vbNewLine & _
    "   a.��ҳid, a.�Һŵ�, a.Ӥ��, a.���˿���id, a.������־ ,e.���� As ����, g.���� As ִ�п���,NVl(NVl(decode(A.������Դ,4,D.����, A.����), F.����),D.����) ���� ,NVl(NVl(decode(A.������Դ,4,D.�Ա�, A.�Ա�), F.�Ա�),D.�Ա�) �Ա�,NVl(NVl(decode(A.������Դ,4,D.����, A.����), F.����),D.����) ���� , d.���￨��, d.���֤��, d.Ic����, d.ҽ����," & vbNewLine & _
    "   Nvl(f.�ѱ�, d.�ѱ�) As �ѱ�, Decode(a.������Դ, 1, d.�����, 2, Decode(f.��������, 1, d.�����, f.סԺ��), 4, d.�����, Null) As ��ʶ��," & vbNewLine & _
    "   f.��Ժ���� As ����,d.����ģʽ, Decode(a.������Դ, 1, '����', 2, 'סԺ', 3, '����', 4, '���') As ��Դ, a.�������, a.����ת��, c.���� As ����, a.�����, a.���ʱ��,A.����ʱ��," & vbNewLine & _
    "   a.ִ�м�, f.��ǰ����id As ����id, f.��Ժ����, f.״̬, c.ִ�а���, Nvl(z.�����ļ�id, 0) As �ļ�id, Nvl(y.ͨ��, 0) As ������, NVL(f.��������,D.��������) As ��������,f.��˱�־, h.���� As ����," & vbNewLine & _
    "   Decode(f.·��״̬, Null, 0, 1) As ·��, a.ҽ������, a.�걾��λ, a.��鷽��, f.��������, a.����ҽ��,F.����״̬, " & vbNewLine & _
    "   Decode(f.��Ժ��ʽ, Null, Decode(f.״̬, 1, 0, 3, 3, 2), Decode(f.��Ժ��ʽ, '����', 5, 4)) As ����,F.�Һ�ID as ��ҳ�Һ�ID" & _
    " From (" & strSQL & ") A,������ĿĿ¼ C,������Ϣ D,������ҳ F,���ű� E,���ű� G,��������Ӧ�� Z,�����ļ��б� Y,������� H" & _
    " Where A.������ĿID=C.ID And A.����ID=D.����ID And A.ִ�в���ID=G.ID " & _
            IIf(mvarCond.��ʶ�� <> "", " And Decode(A.������Դ,1,D.�����,2,Decode(F.��������,1,D.�����,F.סԺ��),3,D.�����,4,D.�����,NULL)=[13]", "") & _
            IIf(mvarCond.���￨ <> "", " And D.���￨��||''=[14]", "") & IIf(mvarCond.���� <> "", " And D.����||''=[15]", "") & _
            IIf(mvarCond.���֤ <> "", " And D.���֤��||''=[16]", "") & IIf(mvarCond.IC���� <> "", " And D.IC����||''=[17]", "") & _
            IIf(mvarCond.ҽ���� <> "", " And D.ҽ����||''=[18]", "") & IIf(mvarCond.����ID <> 0, " And A.����ID=[22]", "") & _
            IIf(mvarCond.����, " And (A.������Դ=2 And A.��ҳID=D.��ҳID Or Nvl(A.������Դ,0)<>2)", "") & str������Դ & strWhere�˶� & _
    "       And A.���˿���ID=E.ID And A.����ID=F.����ID(+) And A.��ҳID=F.��ҳID(+)" & IIf(lng����ID <> 0, " And F.��ǰ����ID+0=[20]", "") & _
    "       And A.������ĿID=Z.������ĿID(+) And A.������Դ=Z.Ӧ�ó���(+) And Z.�����ļ�ID=Y.ID(+) And D.����=H.���(+)" & _
    "       And Not(A.�������='Z' And Nvl(C.��������,'0')<>'0') " & _
    " Order by ����ʱ�� Desc,����ID,���"
        
    Screen.MousePointer = 11
    On Error GoTo errH
    
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), "," & str��Դ & ",", _
        IIf(Mid(mstr״̬, 1, 1) = "1", 2, -1), IIf(Mid(mstr״̬, 2, 1) = "1", 0, -1), IIf(Mid(mstr״̬, 3, 1) = "1", 3, -1), IIf(Mid(mstr״̬, 4, 1) = "1", 1, -1), _
        "|" & mstrRoom & "|", datBegin, datEnd, mvarCond.NO, mvarCond.����ID, "," & mvarCond.��Դ & ",", mvarCond.��ʶ��, mvarCond.���￨, _
        mvarCond.����, mvarCond.���֤, mvarCond.IC����, mvarCond.ҽ����, mvarCond.��Ч - 1, lng����ID, mvarCond.������, mvarCond.����ID)
    
    If blnNoFilter Then mvarCond.������ = ""
    
    '��¼����ѡ�еĲ���
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            lngPatiRow = rptPati.SelectedRows(0).Index '���ڿ������¶�λ
            strPatiRow = rptPati.SelectedRows(0).Record.Tag
            If Not rptPati.SelectedRows(0).ParentRow.GroupRow Then
                '��¼��ǰ������չ���ĸ��չ�����Ӱ��Rows
                strExpend = rptPati.SelectedRows(0).ParentRow.Record.Tag
            End If
        End If
    End If
    rptPati.Records.DeleteAll
    rptPati.Columns.Find(col_����).TreeColumn = False
    
    'ˢ�º�����Զ�չ��
    For i = 1 To rsPati.RecordCount
        '�Ƿ�ֻ��ʾ���շѵĲ���
        '1.ֻ��������,���жϸ��ӷ���
        '2.ֻ���շѻ��۵��ݣ������ʻ��۵���(��Ϊ��Ҫ����ִ�к����)
        '3.����ֱ�Ӱѻ��۵�ɾ�˵����(��������=NULL)
        '4.����Ʒѻ���δ���������õ�Ҳ��ʾ
        blnDo = True: blnSub = False
        '���ҽ��ֻ��ʾһ��(Ϊ�ӿ��ٶȲ���SQL����)
        If Not objPreRecord Is Nothing Then
            If rsPati!������� = "C" And Not IsNull(rsPati!���ID) _
                And objPreRecord(COL_�������).Value = "C" And objPreRecord(col_���ID).Value = Nvl(rsPati!���ID, 0) Then
                objPreRecord(col_����).Value = objPreRecord(col_����).Value & "," & rsPati!����
                blnDo = True: blnSub = True 'һ���ɼ��ļ���������ʾ
                If Not rptPati.Columns.Find(col_����).TreeColumn Then rptPati.Columns.Find(col_����).TreeColumn = True
            ElseIf rsPati!������� = "D" Then
                If Not IsNull(rsPati!���ID) And objPreRecord(col_ҽ��ID).Value = Nvl(rsPati!���ID, 0) Then
                    blnDo = True: blnSub = True '���ͷ�����λ������ʾ
                    If Not rptPati.Columns.Find(col_����).TreeColumn Then rptPati.Columns.Find(col_����).TreeColumn = True
                Else
                    blnDo = True
                End If
            ElseIf rsPati!������� = "F" And Not IsNull(rsPati!���ID) Then
                blnDo = False '�����͸�������
            End If
        End If
                
        If blnDo Then
            If blnSub Then
                '��һ������ϺͶಿλ�����Ŀ����Ϊ������
                If rsPati!������� = "C" Or rsPati!������� = "D" Then
                    If objPreRecord.Childs.Count = 0 Then
                        Set objRecord = objPreRecord.Childs.Add()
                        objRecord.Tag = "Sub_" & objPreRecord.Tag
                        For j = 0 To rptPati.Columns.Count - 1
                            Set objItem = objRecord.AddItem(objPreRecord(j).Value)
                            objItem.Caption = objPreRecord(j).Caption
                            objItem.Icon = objPreRecord(j).Icon
                            objItem.HasCheckbox = objPreRecord(j).HasCheckbox
                            objItem.ForeColor = objPreRecord(j).ForeColor
                            If j = col_���� Then
                                If rsPati!������� = "D" Then
                                    objItem.Value = "" & rsPati!����
                                Else
                                    objItem.Value = Replace(objItem.Value, "," & Nvl(rsPati!����), "")
                                End If
                            End If
                        Next
                    End If
                End If
                Set objRecord = objPreRecord.Childs.Add()
                objRecord.Tag = CStr("Sub_" & rsPati!ҽ��ID & "_" & rsPati!���ͺ�) '���ڲ��˶�λ
            Else
                Set objRecord = Me.rptPati.Records.Add()
                objRecord.Tag = CStr("_" & rsPati!ҽ��ID & "_" & rsPati!���ͺ�) '���ڲ��˶�λ
                If objRecord.Tag = strExpend Then objRecord.Expanded = True
            End If
            
            '����
            Set objItem = objRecord.AddItem(Val(Nvl(rsPati!ִ��״̬, 0))) '������Value��������
            objItem.Caption = Decode(objItem.Value, 0, "δִ��", 1, "��ִ��", 2, "�ܾ�ִ��", 3, "����ִ��")
            
            'ѡ��
            Set objItem = objRecord.AddItem("")
            objItem.HasCheckbox = True
            If mblnFilter And Not blnSub And Val(Nvl(rsPati!ִ��״̬, 0)) = 0 Then
                objItem.Checked = True '������ȱʡѡ��(δִ��)
                rptPati.Columns(col_ѡ��).Caption = "1"
            End If
            
            '·��
            Set objItem = objRecord.AddItem("")
            objItem.Value = Val("" & rsPati!·��)
            objItem.Caption = " "
            If rsPati!·�� = 1 Then
                objItem.Icon = img16.ListImages("Path").Index - 1
                If blnPath = False Then blnPath = True
            End If
            
            'ִ��״̬
            Set objItem = objRecord.AddItem(Val(Nvl(rsPati!ִ��״̬, 0)))
            
            'ͼ��
            Set objItem = objRecord.AddItem("")
            objItem.Icon = Nvl(rsPati!ִ��״̬, 0) 'ImageList�Ǵ�1��ʼ,����ReportControlʱ�Ǵ�0��ʼ
            If Nvl(rsPati!ִ��״̬, 0) = 0 And Nvl(rsPati!ִ�й���, 0) = 1 Then
                objItem.Icon = 5 '�ѱ���
            End If
            
            If objItem.Icon = 3 Then
                '�Ѻ˶Ե�ͼ��ֻ������ִ�е���Ч
                If Val(gstrҽ���˶�) > 0 Then
                    If rsPati!������ & "" <> "" Then
                        If rsPati!�������� & "" = "1" And rsPati!������� & "" = "E" And Mid(gstrҽ���˶�, 2, 1) = "1" Or _
                            rsPati!�������� & "" = "8" And rsPati!������� & "" = "E" And Mid(gstrҽ���˶�, 1, 1) = "1" Or _
                            rsPati!������� & "" = "K" And Mid(gstrҽ���˶�, 1, 1) = "1" Then
                            
                            objRecord(col_ͼ��).Icon = 7
                        End If
                    End If
                End If
            End If
            
            If Nvl(rsPati!��Դ) = "סԺ" And Val("" & rsPati!��������) <> 1 Then mblnShowBed = True
            Set objItem = objRecord.AddItem(CStr(Nvl(rsPati!��Դ)))
            If Nvl(rsPati!��Դ) = "סԺ" And Val("" & rsPati!��������) = 1 Then objItem.Caption = "����"    'value��ΪסԺ
            
            objRecord.AddItem CStr(Nvl(rsPati!NO))
            objRecord.AddItem IIf(Val(rsPati!������־ & "") = 1, "��", "")
            objRecord.AddItem CStr(Nvl(rsPati!����))
'            objRecord.AddItem "��"
            If rsPati!������� = "D" Then
                If IsNull(rsPati!���ID) Then
                    objRecord.AddItem "" & rsPati!ҽ������  '�����а��������еĲ�λ������
                Else
                    If IsNull(rsPati!��鷽��) Then
                        objRecord.AddItem "" & rsPati!�걾��λ
                    Else
                        objRecord.AddItem rsPati!�걾��λ & "(" & rsPati!��鷽�� & ")"
                    End If
                End If
            Else
                objRecord.AddItem "" & rsPati!����
            End If
            
            objRecord.AddItem CStr(Nvl(rsPati!����))
            
            Set objItem = objRecord.AddItem(Val(Nvl(rsPati!���˿���ID, 0)))
            objItem.Caption = Nvl(rsPati!����, " ")
            
            Set objItem = objRecord.AddItem(CStr(Nvl(rsPati!��ʶ��)))
            objItem.Caption = Nvl(rsPati!��ʶ��, " ")
            
            objRecord.AddItem CStr(Nvl(rsPati!����))
            objRecord.AddItem CStr(Nvl(rsPati!�ѱ�))
            
            Set objItem = objRecord.AddItem(Format(rsPati!Ҫ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
            objItem.Caption = Format(rsPati!Ҫ��ʱ��, "yyyy-MM-dd HH:mm")
            If Nvl(rsPati!ִ��״̬, 0) = 0 And Not IsNull(rsPati!����ʱ��) Then
                'δִ�е���Ŀ�����°�����Ҫ��ִ��ʱ���ͻ����ʾ
                objItem.Bold = True
            End If
            
            Set objItem = objRecord.AddItem(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
            objItem.Caption = Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm")
            
            objRecord.AddItem CStr(Nvl(rsPati!ִ�м�))
            objRecord.AddItem CStr(Nvl(rsPati!�Ա�))
            objRecord.AddItem CStr(Nvl(rsPati!����))
            objRecord.AddItem CStr(Nvl(rsPati!�����))
            Set objItem = objRecord.AddItem(Format(rsPati!���ʱ��, "yyyy-MM-dd HH:mm:ss"))
            objItem.Caption = Format(rsPati!���ʱ��, "yyyy-MM-dd HH:mm")
            
            '����������
            Set objItem = objRecord.AddItem(Val(rsPati!ִ�в���ID)): objItem.Caption = rsPati!ִ�п���
            objRecord.AddItem Val(rsPati!����ID)
            objRecord.AddItem Val(Nvl(rsPati!��ҳID, 0))
            objRecord.AddItem CStr(Nvl(rsPati!�Һŵ�))
            objRecord.AddItem 0 '�Һ�ID���ڽ�����ʱ��ȡ
            objRecord.AddItem Val(Nvl(rsPati!Ӥ��, 0))
            objRecord.AddItem CStr(Nvl(rsPati!���￨��))
            objRecord.AddItem CStr(Nvl(rsPati!���֤��))
            objRecord.AddItem CStr(Nvl(rsPati!IC����))
            objRecord.AddItem CStr(Nvl(rsPati!ҽ����))
            objRecord.AddItem Val(Nvl(rsPati!����ID, 0))
            objRecord.AddItem Format(Nvl(rsPati!��Ժ����, ""), "yyyy-MM-dd HH:mm:ss")
            objRecord.AddItem Val(Nvl(rsPati!״̬, 0))
            objRecord.AddItem Val(rsPati!ҽ��ID)
            objRecord.AddItem Val(Nvl(rsPati!���ID, 0))
            objRecord.AddItem Val(rsPati!���ͺ�)
            objRecord.AddItem CStr(rsPati!�������)
            objRecord.AddItem Val(Nvl(rsPati!ִ�й���, 0))
            objRecord.AddItem Val(Nvl(rsPati!ִ�а���, 0))
            objRecord.AddItem Val(Nvl(rsPati!��¼����, 1))
            objRecord.AddItem Val(Nvl(rsPati!����ת��, 0))
            
            objRecord.AddItem Val(Nvl(rsPati!�ļ�ID, 0))
            objRecord.AddItem Val(Nvl(rsPati!������, 0))
            objRecord.AddItem 0 '����ID���ڽ�����ʱ��ȡ
            objRecord.AddItem CStr(Nvl(rsPati!��������))
            objRecord.AddItem Val("" & rsPati!�������)
            
            objRecord.AddItem CStr("" & rsPati!����ҽ��)
            Set objItem = objRecord.AddItem(Val(rsPati!����״̬ & ""))
            objItem.Caption = " "
            objRecord.AddItem (Val(rsPati!����))
            objRecord.AddItem rsPati!�������� & ""
            If InStr(",1,8,", rsPati!�������� & "") > 0 And rsPati!������� & "" = "E" Or rsPati!������� & "" = "K" Then
                objRecord.AddItem rsPati!������ & ""
            Else
                objRecord.AddItem ""
            End If
            objRecord.AddItem rsPati!��˱�־ & ""
            objRecord.AddItem Format(rsPati!����ʱ�� & "", "yyyy-MM-dd HH:mm")
            objRecord.AddItem rsPati!����ģʽ & ""
            objRecord.AddItem rsPati!������ĿID & ""
            objRecord.AddItem rsPati!��Ч & ""
            objRecord.AddItem rsPati!ִ�з��� & ""
	    objRecord.AddItem rsPati!��ҳ�Һ�ID & ""
            objRecord.AddItem ""
            '������ɫ:��ɫ,��ɫ,��ɫ,��ɫ
'            objRecord.Item(0).ForeColor = Decode(Nvl(rsPati!ִ��״̬, 0), 0, 0, 1, &H808080, 2, &H40C0&, 3, &HC00000)
'            For j = 0 To rptPati.Columns.Count - 1
'                objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
'            Next
            If Not IsNull(rsPati!��������) Then
                '���ղ�����ָ��ɫ��ʾ
                lngColor = zlDatabase.GetPatiColor(rsPati!��������)
                objRecord.Item(col_���ݺ�).ForeColor = lngColor
                objRecord.Item(col_��������).ForeColor = lngColor
            ElseIf Not IsNull(rsPati!����) Then
                'δָ���������͵ı��ղ����ú�ɫ��ʾ
                For j = 0 To rptPati.Columns.Count - 1
                    objRecord.Item(j).ForeColor = vbRed
                Next
            End If
            
            '����ִ�еı�ʾ
            If Val(rsPati!ִ�в���ID) <> cboDept.ItemData(cboDept.ListIndex) Then
                objRecord.Item(col_����).Value = objRecord.Item(col_ִ�п���).Caption & "��" & objRecord.Item(col_����).Value
            End If
            If Val(rsPati!������־ & "") = 1 Then
                objRecord.Item(col_����).ForeColor = vbRed
            End If
            If Not blnSub Then Set objPreRecord = objRecord
        End If
        rsPati.MoveNext
    Next
    
    '���û���ٴ�·�����ˣ���������
    rptPati.Columns(col_·��).Visible = blnPath
    
    
    'һ������Ͷಿλ�����ۺϷ���״̬
    If rptPati.Columns.Find(col_����).TreeColumn Then
        For Each objRecord In rptPati.Records
            If objRecord.Childs.Count > 0 Then
                strSQL = ""
                For Each objPreRecord In objRecord.Childs
                    If InStr(strSQL, objPreRecord(col_ִ��״̬).Value) = 0 Then
                        strSQL = strSQL & objPreRecord(col_ִ��״̬).Value
                    End If
                Next
                '�ܾ�ִ��ֻ��������У��������״̬��ʾΪ����ִ��
                '�����Ǹ��游�����ڵķ������
                objRecord(col_�ۺ�״̬).Value = IIf(Len(strSQL) = 1, Val(strSQL), 3)
                objRecord(col_�ۺ�״̬).Caption = Decode(objRecord(col_�ۺ�״̬).Value, 0, "δִ��", 1, "��ִ��", 2, "�ܾ�ִ��", 3, "����ִ��")
                objRecord(col_ͼ��).Icon = objRecord(col_�ۺ�״̬).Value
'                objRecord.Item(0).ForeColor = Decode(objRecord(col_�ۺ�״̬).Value, 0, 0, 1, &H808080, 2, &H40C0&, 3, &HC00000)
'                For j = 0 To rptPati.Columns.Count - 1
'                    objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
'                Next
                If objRecord(col_��������).Value <> "" Then
                    '���ղ����ò�ɫ��ʾ
                    lngColor = zlDatabase.GetPatiColor(objRecord(col_��������).Value)
                    objRecord.Item(col_���ݺ�).ForeColor = lngColor
                    objRecord.Item(col_��������).ForeColor = lngColor
                ElseIf objRecord(col_����).Value <> "" Then
                    'δָ���������͵ı��ղ����ú�ɫ��ʾ
                    For j = 0 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = vbRed
                    Next
                End If
            End If
        Next
    End If
    
    rptPati.Columns.Find(col_����).Visible = mblnShowBed
    rptPati.Populate
    
    '��λ������:��Populate֮��
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Then
        '��������ˢ���Ӵ���
        Call ClearPatiInfo
        Call SubWinRefreshData(tbcSub.Selected)
    Else
        'ȡָ��������
        If strPatiRow <> "" Then
            '�ȿ��ٶ�λ
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '�ٽ��в���
            If objRow Is Nothing Then
                For i = 0 To rptPati.Rows.Count - 1
                    If Not rptPati.Rows(i).GroupRow Then
                        If rptPati.Rows(i).Record.Tag = strPatiRow Then
                            Set objRow = rptPati.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        'ȡ��һ���Ƿ�����
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        If Not objRow.ParentRow.GroupRow Then objRow.ParentRow.Expanded = True
        Set rptPati.FocusedRow = objRow '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
    End If
    
    stbThis.Panels(2).Text = " �� " & rptPati.Records.Count & " ��������Ŀ"
    Screen.MousePointer = 0
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String, i As Long, j As Long
    Dim strTmp As String
    Dim strMsgType As String
    
    On Error GoTo errH
    rptNotify.Records.DeleteAll
    If cboDept.ListIndex = -1 Or cboUnit.ListIndex = -1 Then LoadNotify = True: Exit Function
    If Mid(mstrNotify, m��������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CHARGE_001"
    If Mid(mstrNotify, m������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_004"
    If Mid(mstrNotify, mѪ������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_BLOOD_007"
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
    
    strSQL = "Select b.����id, b.����id as ��ҳid,a.סԺ��,a.����, a.�Ա�, a.����, a.��ǰ���� As ����, Nvl(b.�������id, a.��ǰ����id) As �������id," & _
        " Nvl(b.���ﲡ��id, a.��ǰ����id) As ���ﲡ��id, b.������Դ, b.��Ϣ����, b.���ͱ���, b.ҵ���ʶ, b.���ȳ̶�, b.�Ǽ�ʱ��,a.����" & _
        " From ������Ϣ A, ҵ����Ϣ�嵥 B, ҵ����Ϣ���Ѳ��� C, ҵ����Ϣ������Ա D" & _
        " Where a.����id = b.����id And b.Id = c.��Ϣid And b.Id = d.��Ϣid(+) And b.�Ǽ�ʱ�� >=Trunc(Sysdate-" & (mintDay - 1) & ") and substr(b.���ѳ���,[4],1)='1'" & _
        " And Nvl(b.�Ƿ�����, 0) = 0  And instr(','||[5]||',',','||b.���ͱ���||',')>0 and (c.����id = [1] Or d.������Ա = [3])" & _
        " Order By b.���ȳ̶�, b.�Ǽ�ʱ�� Desc"
    
    Screen.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, cboDept.ItemData(cboDept.ListIndex), , UserInfo.����, 4, strTmp)
    
    If cboDept.ListIndex <> -1 Then
        strTmp = ","
        For i = 1 To rsTmp.RecordCount
            If InStr(strTmp, "," & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & ",") = 0 Then
                strTmp = strTmp & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & ","
                Call AddReportRow(rsTmp!����ID & "," & rsTmp!��ҳID, rsTmp!����ID, rsTmp!��ҳID, Nvl(rsTmp!����), Nvl(rsTmp!��Ϣ����), rsTmp!���ͱ���, _
                        rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!ҵ���ʶ & "")
            End If
            rsTmp.MoveNext
        Next
    End If
    rptNotify.Populate 'ȱʡ��ѡ���κ���
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln��Ϣ���� Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(3)
        End If
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Set mrsMsg = rsTmp
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'���ܣ��������������ص���ʾ��Ϣ
    mlng����ID = 0
    mlng��ҳID = 0
    mstr�Һŵ� = ""
    
    lblAdvice.Caption = ""
    lblDiag(1).Caption = ""
    lblCash.Visible = False
    lblRec.Visible = False
        
    vsExec.Rows = vsExec.FixedRows
    vsExec.Rows = vsExec.FixedRows + 1
    picBlood.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And Not (Me.ActiveControl Is PatiIdentify Or Me.ActiveControl Is rtfAppend Or Me.ActiveControl Is cboDept Or Me.ActiveControl Is cboUnit) And mstrFindType = "���￨" Then
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        PatiIdentify.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
    End If
End Sub

Private Sub timBRefresh_Timer()
    '��Ѫ����Ѫִ�д�����д��ִ�����ݺ�ҽ����Ӧ���ݵ�ˢ��
    Dim intState As Integer
    timBRefresh.Enabled = False
    If Not mobjFrmBloodExe Is Nothing Then
        On Error Resume Next
        intState = mobjFrmBloodExe.AdviceExecState
        If err <> 0 Then
            err.Clear
        Else
            mobjFrmBloodExe.ExecFresh = True
            Select Case intState
                Case 1, 2 '��¼ִ�л����ִ�У�ɾ��ִ��
                    Call LoadPatients 'Ҫ����ִ��״̬
                Case 3, 4 'ִ�����,ȡ�����
                    Call LoadPatients 'Ҫ����ִ��״̬
            End Select
            mobjFrmBloodExe.ExecFresh = False
            mobjFrmBloodExe.AdviceExecState = 0
        End If
    End If
End Sub


Private Sub timRefresh_Timer()
    Static lngSec�����б� As Long
    Static strTim��Ϣ�б� As String
    
    Dim curTime As Date
    
    If mintRefresh <> 0 Then
        lngSec�����б� = lngSec�����б� + 1 '����
        If lngSec�����б� Mod mintRefresh = 0 Then
            lngSec�����б� = 0
            Call LoadPatients
        End If
    End If
    
    If mbln��Ϣ���� Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                timRefresh.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                timRefresh.Enabled = True
            End If
        End If
    End If
    
    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then 'ʹ������Ϣƽ̨�����Զ�ˢ����Ϣ�б�
            Exit Sub
        End If
    End If
    
    If mintMin > 0 And rptNotify.Visible Then
        curTime = Now
        
        If strTim��Ϣ�б� = "" Then
            strTim��Ϣ�б� = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strTim��Ϣ�б�), curTime) > mintMin * CLng(60) Then
            strTim��Ϣ�б� = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call LoadNotify
        End If
    End If
    
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String, Optional ByVal lngPatiID As Long)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
'      strIDCard=����ֵʱ����ʾ�̶������֤�Ų���
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    
    '��������ʽ���Һ��Զ�ˢ���֤�ļ���������ȡ��
    If strIDCard = "" And PatiIdentify.Text <> "" Then mvarCond.���֤ = ""
    
    If Not blnNext And mstrFindType = "���ݺ�" Then
        PatiIdentify.Text = GetFullNO(PatiIdentify.Text, 12)
    End If
    PatiIdentify.SetFocus
            
    '����ģʽʱ����ָ������������ȡ�����嵥
    If mblnFilter Then
        Call ClearPatiCond '�����ڹ��������õ����ʶ���������
        If strIDCard <> "" Then '���֤�Զ�ʶ��ǿ������
            mvarCond.���֤ = strIDCard
        Else
            Select Case mstrFindType
                Case "���￨"
                    mvarCond.���￨ = PatiIdentify.Text '���￨
                Case "��ʶ��"
                    mvarCond.��ʶ�� = PatiIdentify.Text '��ʶ��
                Case "���ݺ�"
                    mvarCond.NO = PatiIdentify.Text '���ݺ�
                Case "����"
                    mvarCond.���� = PatiIdentify.Text '����
                Case "�������֤"
                    mvarCond.���֤ = PatiIdentify.Text '���֤
                Case "IC��"
                    If Not mobjSquareCard Is Nothing Then 'IC��
                        Call mobjSquareCard.zlGetPatiID("IC��", PatiIdentify.Text, , mvarCond.����ID)
                    Else
                        mvarCond.IC���� = PatiIdentify.Text
                    End If
                Case "ҽ����"
                    mvarCond.ҽ���� = PatiIdentify.Text 'ҽ����
                Case Else
                    If Not mobjSquareCard Is Nothing Then
                        Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����), PatiIdentify.Text, , mvarCond.����ID)
                    End If
            End Select
        End If
        Call LoadPatients
        Exit Sub
    End If
    
    '��ʼ������
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then blnHave = True
    End If
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl����������0��ʼ
    Else
        i = rptPati.SelectedRows(0).Index + 1
    End If
    
    '���Ҳ���
    If lngPatiID = 0 And Not mobjSquareCard Is Nothing And mstrFindType <> "���￨" And mstrFindType <> "��ʶ��" And mstrFindType <> "���ݺ�" And mstrFindType <> "����" And mstrFindType <> "�������֤" And mstrFindType <> "ҽ����" Then
        If mstrFindType = "IC��" Then
            Call mobjSquareCard.zlGetPatiID("IC��", PatiIdentify.Text, , lngPatiID)
        Else
            Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����), PatiIdentify.Text, , lngPatiID)
        End If
    End If
    
    With rptPati
        For i = i To .Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).ParentRow.GroupRow Then
                    If strIDCard <> "" Then '���֤�Զ�ʶ��ǿ������
                        If UCase(.Rows(i).Record(col_���֤��).Value) = UCase(strIDCard) Then Exit For
                    Else
                        If Val(.Rows(i).Record(col_����Id).Value) = lngPatiID And lngPatiID <> 0 Then Exit For
                        Select Case mstrFindType
                            Case "���￨"
                                If .Rows(i).Record(col_���￨��).Value = PatiIdentify.Text Then Exit For
                            Case "��ʶ��"
                                If .Rows(i).Record(col_��ʶ��).Value = PatiIdentify.Text Then Exit For
                            Case "���ݺ�"
                                If UCase(.Rows(i).Record(col_���ݺ�).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "����"
                                If .Rows(i).Record(col_����).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                            Case "�������֤"
                                If UCase(.Rows(i).Record(col_���֤��).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "ҽ����"
                                If UCase(.Rows(i).Record(col_ҽ����).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case Else
                                If Val(.Rows(i).Record(col_����Id).Value) = lngPatiID Then Exit For
                        End Select
                    End If
                End If
            End If
        Next
    End With

    If i <= rptPati.Rows.Count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptPati.FocusedRow = rptPati.Rows(i)
        If rptPati.Visible Then rptPati.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub ClearPatiCond()
    mvarCond.IC���� = "": mvarCond.NO = "": mvarCond.��ʶ�� = ""
    mvarCond.���￨ = "": mvarCond.���֤ = "": mvarCond.���� = ""
    mvarCond.ҽ���� = "": mvarCond.������ = "": mvarCond.����ID = 0
End Sub

Private Sub PatientFilter()
    timRefresh.Enabled = False
    frmTechnicFilter.mstrDeptNode = mstrDeptNode
    frmTechnicFilter.mstrPrivs = mstrPrivs
    frmTechnicFilter.mstrCardKind = mstrCardKind
    Set frmTechnicFilter.mobjSquareCard = mobjSquareCard
    frmTechnicFilter.Show 1, Me '��,��������˴��ڵ�Form_Activate�¼�
    If frmTechnicFilter.mblnOK Then
        '���ù��˱���
        With frmTechnicFilter
            '����ʱ��
            mvarCond.Begin = Format(.dtpBegin.Value, "yyyy-MM-dd HH:mm:00")
            If Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
                mvarCond.End = CDate(0) '��ʾȡ��ǰʱ��
            Else
                mvarCond.End = Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm:59")
            End If
            
            Call ClearPatiCond
            If .mlngPatiID <> 0 Then
                '��ȱʡ����ˢ���಻һ��ʱ:ѡ������ʵ����ˢ���￨
                mvarCond.����ID = .mlngPatiID
            Else
                Select Case .mstrFindType
                    Case "���￨"
                        mvarCond.���￨ = .PatiIdentify.Text '���￨
                    Case "��ʶ��"
                        mvarCond.��ʶ�� = .PatiIdentify.Text '��ʶ��
                    Case "���ݺ�"
                        mvarCond.NO = .PatiIdentify.Text '���ݺ�
                    Case "����"
                        mvarCond.���� = .PatiIdentify.Text '����
                    Case "�������֤"
                        mvarCond.���֤ = .PatiIdentify.Text '���֤
                    Case "IC��"
                        mvarCond.IC���� = .PatiIdentify.Text
                    Case "ҽ����"
                        mvarCond.ҽ���� = .PatiIdentify.Text 'ҽ����
                    Case Else
                        If Not mobjSquareCard Is Nothing Then
                            Call mobjSquareCard.zlGetPatiID(Val(.PatiIdentify.objIDKind.GetCurCard.�ӿ����), .PatiIdentify.Text, , mvarCond.����ID)
                        End If
                End Select
            End If
            '���˿���
            If .cboDept.ListIndex <> 0 Then
                mvarCond.����ID = .cboDept.ItemData(.cboDept.ListIndex)
            Else
                mvarCond.����ID = 0
            End If
            
            '������Դ
            mvarCond.��Դ = ""
            If Not (.chk��Դ(0).Value = 1 And .chk��Դ(1).Value = 1 And .chk��Դ(2).Value = 1) Then
                If .chk��Դ(0).Value = 1 Then mvarCond.��Դ = mvarCond.��Դ & ",1"
                If .chk��Դ(1).Value = 1 Then mvarCond.��Դ = mvarCond.��Դ & ",2"
                If .chk��Դ(2).Value = 1 Then mvarCond.��Դ = mvarCond.��Դ & ",4"
                mvarCond.��Դ = Mid(mvarCond.��Դ & ",3", 2)
            End If
            
            '����סԺ
            mvarCond.���� = .chk����סԺ.Value = 1
            
            'ҽ����Ч
            mvarCond.��Ч = 0
            If Not (.chk��Ч(0).Value = 1 And .chk��Ч(1).Value = 1) Then
                If .chk��Ч(0).Value = 1 Then
                    mvarCond.��Ч = 1
                ElseIf .chk��Դ(1).Value = 1 Then
                    mvarCond.��Ч = 2
                End If
            End If
            
            '������
            mvarCond.������ = ""
            If .cboDoctor.Text <> "" And .cboDoctor.ListIndex <> 0 Then
                mvarCond.������ = Split(.cboDoctor.Text, "-")(1)
            End If
        End With
         '����ļ�ʱ�����������
        Me.PatiIdentify.Text = ""
        mstr������ = mvarCond.������
                        
        Call SetUnitVisible
        Call LoadPatients 'ˢ��
                
        'û�в���ʱȱʡ��ʾ��ҽ��ҳ��
        If rptPati.Rows.Count = 0 Then
            If mvarCond.��Դ = "2,3" Then
                Call ExchangeAdvice(False) 'ȱʡ��ʾסԺ��
            Else
                Call ExchangeAdvice(True) 'ȱʡ��ʾ�����
            End If
        End If
    End If
    timRefresh.Enabled = True
End Sub

Private Sub SetUnitVisible()
    If Not (mvarCond.��Դ = "2,3") Then
        If cboUnit.ListCount > 0 Then
            Call Cbo.SetIndex(cboUnit.hwnd, 0)
        End If
    End If
    
    lblUnit.Visible = mvarCond.��Դ = "2,3"
    cboUnit.Visible = mvarCond.��Դ = "2,3"
    Call picPati_Resize
    Me.Refresh
End Sub

Private Sub ParameterSetup()
    Dim strRoom As String, str������� As String, str������� As String
    
    timRefresh.Enabled = False
    frmTechnicSetup.mstrPrivs = mstrPrivs
    frmTechnicSetup.mlng����ID = cboDept.ItemData(cboDept.ListIndex)
    frmTechnicSetup.Show 1, Me
    If frmTechnicSetup.mblnOK Then
        '�ϸ�����¼ִ�е����
        mblnExeLog = Val(zlDatabase.GetPara("��¼ִ�����", glngSys, pҽ������վ, "0")) <> 0
        
        'Ƥ����֤���
        mblnƤ����֤ = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, pҽ������վ)) <> 0
        
        str������� = zlDatabase.GetPara("�������", glngSys, pҽ������վ)
        str������� = zlDatabase.GetPara("�������", glngSys, pҽ������վ)
    
        'ִ�м䷶Χ�ı�
        strRoom = zlDatabase.GetPara("ִ�м䷶Χ", glngSys, pҽ������վ)
        If strRoom <> mstrRoom Or str������� <> mstr������� Or str������� <> mstr������� Then
            mstrRoom = strRoom
            mstr������� = str�������
            mstr������� = str�������
            Call LoadPatients
        End If
        
        '�����Զ�ˢ��
        Call SetTimer
    End If
    timRefresh.Enabled = True
End Sub

Private Function Getִ������(ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal str��� As String, ByVal objRow As ReportRow) As String
'���ܣ�����ָ����ҽ��ID,����ҽ�����ݹ���ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim bln��ҩ;�� As Boolean, i As Integer
    Dim strƤ�Խ�� As String

    On Error GoTo errH
    
    '��ȡҽ������
    If (str��� = "C" And lng���ID <> 0) Or str��� = "D" Then
        strTmp = rptPati.SelectedRows(0).Record(col_����).Value
        
    ElseIf str��� <> "E" Or lng���ID <> 0 Then
        '�䷽�巨,��������,��Ѫ;��,������ҽ��,ֱ����ʾҽ������
        strSQL = "Select ҽ������ From ����ҽ����¼ Where ID=[1]"
        If rptPati.SelectedRows(0).Record(COL_����ת��).Value = 1 Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(str��� = "E", lng���ID, lngҽ��ID))
        If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!ҽ������)
    Else
        '���ΪE,�����ID=0
        strSQL = "Select A.ID,A.���ID,A.�������,A.ҽ������,A.Ƥ�Խ��,A.��������,B.���㵥λ,B.��������,A.ִ��Ƶ��,A.ִ��ʱ�䷽��,B.����" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B" & _
            " Where Not (A.�������='E' And ���ID is Not NULL) And A.������ĿID=B.ID" & _
            " And (A.���ID=[1] Or A.ID=[1])" & _
            " Order by A.���"
        If rptPati.SelectedRows(0).Record(COL_����ת��).Value = 1 Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
        rsTmp.Filter = "���ID=" & lngҽ��ID
        If Not rsTmp.EOF Then bln��ҩ;�� = InStr(",5,6,", rsTmp!�������) > 0
        
        If Not bln��ҩ;�� Then
            'һ��������Ŀ����ҩ�÷�����ɼ�����
            rsTmp.Filter = 0
            If Not rsTmp.EOF Then
                If rsTmp!������� = "E" And rsTmp!�������� = "1" Then
                    strƤ�Խ�� = "��Ƥ�Խ����" & Nvl(rsTmp!Ƥ�Խ��)
                    
                    strSQL = "Select b.������Ӧ, b.����ʱ�� From ����ҽ����¼ A, ���˹�����¼ B, ������ĿĿ¼ C, �����÷����� D" & _
                        " Where a.����id = b.����id And a.������Ŀid = d.�÷�id And d.��Ŀid = c.Id And c.��� In ('5', '6') And d.��Ŀid = b.ҩ��id And" & _
                        " Nvl(d.����, 0) = 0 And b.��¼ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = a.id And �������� = 10) And a.Id = [1] And RowNum<2"

                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                    
                    If Not rsTmp.EOF Then
                        strƤ�Խ�� = strƤ�Խ�� & ",����ʱ�䣺" & Format(rsTmp!����ʱ��, "yyyy-MM-dd") & IIf(Nvl(rsTmp!������Ӧ) = "", "", ",������Ӧ��" & rsTmp!������Ӧ)
                    End If
                End If
            End If
            
            strSQL = "Select ҽ������ From ����ҽ����¼ Where ID=[1]"
            If rptPati.SelectedRows(0).Record(COL_����ת��).Value = 1 Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!ҽ������)
        Else
            '��ҩ;��
            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & "," & rsTmp!ҽ������ & IIf(Not IsNull(rsTmp!��������), " " & FormatEx(rsTmp!��������, 5) & rsTmp!���㵥λ, "")
                rsTmp.MoveNext
            Next
            rsTmp.Filter = "ID=" & lngҽ��ID
            strTmp = rsTmp!���� & "," & rsTmp!ִ��Ƶ�� & "(" & rsTmp!ִ��ʱ�䷽�� & "):ÿ" & rsTmp!���㵥λ & " " & Mid(strTmp, 2)
        End If
    End If
    
    '��ȡ��������
    strSQL = "Select A.��������,Nvl(D.���㵥λ,C.���㵥λ) as ���㵥λ" & _
        " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C,�շ���ĿĿ¼ D" & _
        " Where A.ҽ��ID=[1] And A.���ͺ�=[2]" & _
        " And A.ҽ��ID=B.ID And B.������ĿID=C.ID And B.�շ�ϸĿID=D.ID(+)"
    If rptPati.SelectedRows(0).Record(COL_����ת��).Value = 1 Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    'Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��������) Then
            Getִ������ = " ִ�����ݣ�" & strTmp & strƤ�Խ��
        Else
            Getִ������ = " �������Σ�" & FormatEx(rsTmp!��������, 5) & " " & Nvl(rsTmp!���㵥λ) & "��ִ�����ݣ�" & strTmp & strƤ�Խ��
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadExecList(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long) As Boolean
'���ܣ���ȡָ��ҽ����ִ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strPre As String
    Dim rsѪ�� As ADODB.Recordset
    Dim bln��Ѫ As Boolean
    Dim intѪ���� As Integer
    
    On Error GoTo errH
    
    '������Ŀһ��ִ��ʱ��ִ������Ǽǵ���һ����Ŀ�ϡ���ɢ����ִ��ʱ���Ǽǵ�������Ŀ�ϡ�
    strSQL = "Select A.Ҫ��ʱ��,A.ִ��ʱ��,A.��������,D.���㵥λ,A.ִ��ժҪ,A.ִ����,A.�Ǽ�ʱ��,A.�Ǽ���,DECODE(NVL(A.ִ�н��,1),0,'δִ��',1,'���',2,'�ܾ�',3,'���') As ִ�н��,a.�˶���,a.�˶�ʱ��,d.��������,d.���,a.˵��,a.��¼��Դ as ��Դ" & _
        " From ����ҽ��ִ�� A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D" & _
        " Where A.ҽ��ID=[1] And A.���ͺ�=[2]" & _
        " And A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID And C.������ĿID=D.ID" & _
        " Order by A.�Ǽ�ʱ�� Desc"
    With rptPati.SelectedRows(0)
        If .Record(COL_����ת��).Value = 1 Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            strSQL = Replace(strSQL, "����ҽ��ִ��", "H����ҽ��ִ��")
        End If
    End With
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    With vsExec
        strPre = .Cell(flexcpData, .Row, 0)
        .Redraw = flexRDNone
        .Rows = vsExec.FixedRows
        .Rows = vsExec.FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            '��Ѫҽ���������̱䶯 70823
            If gblnѪ��ϵͳ And Val(rsTmp!�������� & "") = 8 And rsTmp!��� = "E" Then
                strSQL = "select zl_Get_��Ѫִ�д���(���id) as ���� from ����ҽ����¼ where id = [1]"
                Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                If Not rsѪ��.EOF Then intѪ���� = Val(rsѪ��!���� & "")
                bln��Ѫ = True
            End If
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Format(rsTmp!Ҫ��ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 1) = Format(rsTmp!ִ��ʱ��, "yyyy-MM-dd HH:mm")
                If bln��Ѫ Then
                    .TextMatrix(i, 2) = FormatEx(Val(rsTmp!�������� & "") * intѪ����, 0) & " ��"
                Else
                    .TextMatrix(i, 2) = FormatEx(rsTmp!��������, 5) & " " & Nvl(rsTmp!���㵥λ)
                End If
                .TextMatrix(i, 3) = Nvl(rsTmp!ִ��ժҪ)
                .TextMatrix(i, 4) = Nvl(rsTmp!ִ����)
                .TextMatrix(i, 5) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 6) = Nvl(rsTmp!�Ǽ���)
                .TextMatrix(i, 7) = rsTmp!ִ�н�� & ""
                .TextMatrix(i, 8) = Nvl(rsTmp!�˶���)
                .TextMatrix(i, 9) = Format(rsTmp!�˶�ʱ��, "yyyy-MM-dd HH:mm")
		.TextMatrix(i, 10) = NVL(rsTmp!˵��)
                .TextMatrix(i, 11) = IIf(1 = Val(rsTmp!��Դ & ""), "�ƶ���", "PC��")
                .Cell(flexcpData, i, 0) = Format(rsTmp!Ҫ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                .Cell(flexcpData, i, 1) = Format(rsTmp!ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                If .Cell(flexcpData, i, 0) = strPre Then .Row = i
                rsTmp.MoveNext
            Next
            rsTmp.MoveFirst
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
    End With
    LoadExecList = True
    With rptPati
        If .SelectedRows.Count > 0 Then
            If Not .SelectedRows(0).GroupRow Then
            
                If Not (.SelectedRows(0).Record(COL_�������).Value = "E" And .SelectedRows(0).Record(COL_��������).Value = "1" And Mid(gstrҽ���˶�, 2, 1) = "1" Or _
                    .SelectedRows(0).Record(COL_�������).Value = "E" And .SelectedRows(0).Record(COL_��������).Value = "8" And Mid(gstrҽ���˶�, 1, 1) = "1" Or _
                    .SelectedRows(0).Record(COL_�������).Value = "K" And Mid(gstrҽ���˶�, 1, 1) = "1") Then
                    
                    vsExec.ColHidden(8) = True
                    vsExec.ColHidden(9) = True
                Else
                    vsExec.ColHidden(8) = False
                    vsExec.ColHidden(9) = False
                End If
                
            End If
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsExec_GotFocus()
    vsExec.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsExec_LostFocus()
    vsExec.BackColorSel = COLOR_LOST
End Sub

Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
        
    If mlng����ID = 0 Then
        objPopup.CommandBar.Controls.DeleteAll
        ShowBillList = True: Exit Function
    End If
        
    With rptPati.SelectedRows(0)
        '�������ʾ���Ƶ���
        If Not .ParentRow.GroupRow Then
            objPopup.CommandBar.Controls.DeleteAll
            ShowBillList = True: Exit Function
        End If
        
        If .Record(col_ҽ��ID).Value & "_" & .Record(col_���ͺ�).Value = objPopup.Parameter Then
            ShowBillList = True: Exit Function
        Else
            objPopup.Parameter = .Record(col_ҽ��ID).Value & "_" & .Record(col_���ͺ�).Value
            objPopup.CommandBar.Controls.DeleteAll
        End If
    End With
        
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        strSQL = "Select Distinct D.���,D.����,D.˵��" & _
            " From ����ҽ������ A,����ҽ����¼ B,��������Ӧ�� C,�����ļ��б� D" & _
            " Where A.���ͺ�=[1] And A.NO=[2]" & _
            " And A.ҽ��ID=B.ID And B.������ĿID=C.������ĿID" & _
            " And C.Ӧ�ó���=[3] And C.�����ļ�ID=D.ID And D.����=7" & _
            " Order by D.���"
        If .Record(COL_����ת��).Value = 1 Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Record(col_���ͺ�).Value, .Record(col_���ݺ�).Value, .Record(col_��¼����).Value)
    End With
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + i, rsTmp!����)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
                'If i > 1 Then objControl.Enabled = False 'һ����Ŀֻ������һ�����Ƶ���
            End With
            rsTmp.MoveNext
        Next
        
        cbsMain.KeyBindings.Add 0, vbKeyF2, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
    Dim strNO As String, int���� As Integer
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub
    If objControl.Parameter = "" Then '��֣�ֱ�Ӱ�F2ʱ����һ���յ�Control
        Set objControl = cbsMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.Parameter = "" Then Exit Sub
    
    With rptPati.SelectedRows(0)
        If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
            '�Ƿ��ǲɼ���ʽ�������Ҫ���⴦��
            strSQL = "Select A.�������,B.�������� From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Record(col_ҽ��ID).Value)
            If Not rsTmp.EOF Then
                '�ǲɼ���ʽ
                If Nvl(rsTmp(0)) = "E" And Nvl(rsTmp(1)) = "6" Then
                    Print�ɼ���ʽ .Record(col_���ݺ�).Value, .Record(col_��¼����).Value, .Record(col_ҽ��ID).Value, objControl.Parameter
                Else
                    Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & .Record(col_���ݺ�).Value, "����=" & .Record(col_��¼����).Value, "��Ŀ=Untitled", 2)
                End If
            Else
                'Ϊ�˴�ӡ���룬���������ˡ���Ŀ��������By����ͮ��
                Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & .Record(col_���ݺ�).Value, "����=" & .Record(col_��¼����).Value, "��Ŀ=Untitled", 2)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Print�ɼ���ʽ(ByVal strNO As String, ByVal intAttribute As Integer, ByVal lngAdviceID As Long, ByVal strReport As String)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo DataError
    Me.MousePointer = vbHourglass
    
    'ͬһ�걾�ٰ������ֱ��ӡ
'    strSQL = "Select ����ID,�걾,ִ�в���,NO," & _
'        " Trim(����1||' '||����2||' '||����3||' '||����4||' '||����5) As ��Ŀ,����" & _
'        " From" & _
'        " (Select B.����ID,B.�걾��λ As �걾,F.���� As ִ�в���,S.����," & _
'        "  Max(Decode(Mod(Rownum,5),0,B.ҽ������,'')) As ����1," & _
'        "  Max(Decode(Mod(Rownum,5),1,B.ҽ������,'')) As ����2," & _
'        "  Max(Decode(Mod(Rownum,5),2,B.ҽ������,'')) As ����3," & _
'        "  Max(Decode(Mod(Rownum,5),3,B.ҽ������,'')) As ����4," & _
'        "  Max(Decode(Mod(Rownum,5),4,B.ҽ������,'')) As ����5," & _
'        "  Max(S.NO||','||S.��¼����) As NO" & _
'        "  From ����ҽ����¼ B,���ű� F," & _
'        "   (Select DISTINCT ҽ��ID,NO,��¼����,����,������ĿID FROM " & _
'        "    (Select A.ҽ��ID,A.NO,A.��¼����,B.������ĿID,I.������ĿID,MAX(Decode(M.����,NULL,'�ֹ�',M.����)) AS ���� " & _
'        "     From ����ҽ������ A,����ҽ����¼ B,����ҽ����¼ D,������ĿĿ¼ C,���鱨����Ŀ I,����������Ŀ J,�������� M," & _
'        "     (SELECT A.����ID,B.����ʱ��,B.ִ�в���ID FROM ����ҽ����¼ A,����ҽ������ B" & _
'        "      WHERE A.ID=B.ҽ��ID AND B.NO=[1] AND B.��¼����=[2]) N Where a.ҽ��ID+0 = B.ID And B.������ĿID = C.ID" & _
'        "      AND D.���ID = B.ID AND D.������ĿID=I.������ĿID(+) AND I.������ĿID=J.��ĿID(+) AND J.����ID=M.ID(+)" & _
'        "      And C.���='E' And Nvl(C.��������,'0')='6'" & _
'        "      And B.����ID=N.����ID And A.ִ�в���ID+0= N.ִ�в���ID And A.����ʱ�� BETWEEN to_Date(to_Char(N.����ʱ��,'YYYY-MM-DD'),'YYYY-MM-DD HH24:MI:SS') AND to_Date(to_Char(N.����ʱ��,'YYYY-MM-DD')||' 23:59:59','YYYY-MM-DD HH24:MI:SS')" & _
'        "      And Nvl(A.ִ��״̬,0)=0 " & _
'        "     GROUP BY A.ҽ��ID,A.NO,A.��¼����,B.������ĿID,I.������ĿID)" & _
'        "   ) S" & _
'        "  Where B.ִ�п���ID = F.ID And B.���ID = S.ҽ��ID" & _
'        "  Group By B.����ID, B.�걾��λ,F.����,S.����,S.������ĿID)" & _
'        " Order By ����ID"
    'ͬһ�걾ֻ��һ��
    strSQL = "Select ����ID,�걾,ִ�в���," & _
        " Trim(����1||' '||����2||' '||����3||' '||����4||' '||����5) As ��Ŀ" & _
        " From" & _
        " (Select B.����ID,B.�걾��λ As �걾,F.���� As ִ�в���," & _
        "  Max(Decode(Mod(Rownum,5),0,B.ҽ������,'')) As ����1," & _
        "  Max(Decode(Mod(Rownum,5),1,B.ҽ������,'')) As ����2," & _
        "  Max(Decode(Mod(Rownum,5),2,B.ҽ������,'')) As ����3," & _
        "  Max(Decode(Mod(Rownum,5),3,B.ҽ������,'')) As ����4," & _
        "  Max(Decode(Mod(Rownum,5),4,B.ҽ������,'')) As ����5 " & _
        "  From ����ҽ����¼ B,���ű� F" & _
        "  Where B.ִ�п���ID = F.ID And B.���ID = [1]" & _
        "  Group By B.����ID, B.�걾��λ,F.����)" & _
        " Order By ����ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    If rsTmp.EOF Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Call ReportOpen(gcnOracle, glngSys, strReport, Me, "NO=" & strNO, "����=" & intAttribute, "��Ŀ=" & Nvl(rsTmp("��Ŀ")), 2)
    
    Me.MousePointer = vbDefault
    Exit Sub
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
    Me.MousePointer = vbDefault
End Sub

Private Function FuncShowReport(ByVal intOption As Integer) As Boolean
'���ܣ�������д/����/��ӡ
'������intOption=0-��д,1-����,2-��ӡ,3-Ԥ��
    Dim rsTmp As ADODB.Recordset
    Dim lngҽ��ID As Long, strBill As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        If intOption = 0 And .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
    
        If (.Record(COL_�������).Value = "C" Or .Record(COL_�������).Value = "D") And .Record(col_���ID).Value <> 0 Then
            lngҽ��ID = .Record(col_���ID).Value '������ϺͶಿλ���ȡ���ID
        Else
            lngҽ��ID = .Record(col_ҽ��ID).Value
        End If
        
        Set mclsEPRReport = New zlRichEPR.cEPRDocument
        
        If intOption = 0 Then
            If .Record(col_����ID).Value = 0 Then
                Call mclsEPRReport.InitEPRDoc(cprEM_����, cprET_�������༭, .Record(col_�ļ�ID).Value, _
                    Decode(.Record(col_��Դ).Value, "����", cprPF_����, "סԺ", cprPF_סԺ, "���", cprPF_���, "����", cprPF_����), _
                    mlng����ID, IIf(.Record(col_�Һ�ID).Value <> 0, .Record(col_�Һ�ID).Value, .Record(col_��ҳID).Value), _
                    .Record(col_Ӥ��).Value, mlngDept, lngҽ��ID)
            Else
                Call mclsEPRReport.InitEPRDoc(cprEM_�޸�, cprET_�������༭, .Record(col_����ID).Value, _
                    Decode(.Record(col_��Դ).Value, "����", cprPF_����, "סԺ", cprPF_סԺ, "���", cprPF_���, "����", cprPF_����), _
                    mlng����ID, IIf(.Record(col_�Һ�ID).Value <> 0, .Record(col_�Һ�ID).Value, .Record(col_��ҳID).Value), _
                    .Record(col_Ӥ��).Value, mlngDept, lngҽ��ID)
            End If
            Call mclsEPRReport.ShowEPREditor(Me) '�Ƿ�ģ̬��ʾ
        ElseIf intOption = 1 Then
            Call gobjRichEPR.ViewDocument(Me, .Record(col_����ID).Value, True)
        ElseIf intOption = 2 Or intOption = 3 Then
            If .Record(col_������).Value = 1 Then
                '���༭��ʽ��ӡ
                Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr���Ʊ���, .Record(col_����ID).Value, IIf(intOption = 2, True, False))
            ElseIf .Record(col_������).Value = 2 Then
                '�������ʽ��ӡ
                strSQL = "Select ��� From �����ļ��б� Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(col_�ļ�ID).Value))
                strBill = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-2"
                
                If intOption = 2 Then
                    If Not ReportPrintSet(gcnOracle, glngSys, strBill, Me) Then Exit Function
                End If
                Call ReportOpen(gcnOracle, glngSys, strBill, Me, "NO=" & .Record(col_���ݺ�).Value, "����=" & .Record(col_��¼����).Value, "ҽ��ID=" & lngҽ��ID, IIf(intOption = 2, 2, 1))
            End If
        End If
    End With
    
    FuncShowReport = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncBatchPrint()
'���ܣ�������ӡ����
    Dim strPatiSource As String
    
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "סԺ����") > 0 Then
        strPatiSource = "1,2,3"
    ElseIf InStr(mstrPrivs, "���ﲡ��") > 0 Then
        strPatiSource = "1"
    ElseIf InStr(mstrPrivs, "סԺ����") > 0 Then
        strPatiSource = "2"
    Else
        strPatiSource = "3"
    End If
    frmLISBillPrint.ShowMe Me, strPatiSource, cboDept.ItemData(cboDept.ListIndex)
End Sub

Private Sub FuncExecPlanTime()
'���ܣ�ʱ�䰲��
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lngִ�п���ID As Long
    
    With rptPati.SelectedRows(0)
        If .Record(col_ִ�й���).Value > 1 Then
            MsgBox "����Ŀ����Ӱ���鹤��վ�����������ٲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        If .Record(col_ִ�п���).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lngִ�п���ID = cboDept.ItemData(cboDept.ListIndex)
        End If
    End With
    With frmTechnicPlanTime
        If .ShowMe(Me, mclsMipModule, lngҽ��ID, lng���ͺ�, lngִ�п���ID) Then Call LoadPatients
    End With
End Sub

Private Sub FuncExecPlan()
'���ܣ�ִ�б���
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lngִ�п���ID As Long, lng����ID As Long, lng�����ID As Long
    
    With rptPati.SelectedRows(0)
        If .Record(col_ִ�й���).Value > 1 Then
            MsgBox "����Ŀ����Ӱ���鹤��վ�����������ٲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(col_ִ�й���).Value = 1 Then
            If MsgBox("�ò����Ѿ���������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        lng����ID = .Record(col_����Id).Value
        If .Record(col_ִ�п���).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lngִ�п���ID = cboDept.ItemData(cboDept.ListIndex)
        End If
        'ȡ���￨����
        lng�����ID = Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����)
    End With
    With frmTechnicPlan
        If .ShowMe(Me, lngҽ��ID, lng���ͺ�, lngִ�п���ID, lng�����ID, lng����ID, mstrPrivs, mobjSquareCard) Then Call LoadPatients
    End With
End Sub

Private Sub FuncExecErase()
'���ܣ�ȡ������
    Dim lngҽ��ID As Long, lng���ͺ� As Long, strSQL As String
        
    With rptPati.SelectedRows(0)
        If .Record(col_ִ�й���).Value > 1 Then
            MsgBox "����Ŀ����Ӱ���鹤��վ�����������ٲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Record(col_ִ�й���).Value = 0 Then
            MsgBox "�ò��˻�û�б�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪȡ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
    End With
    
    strSQL = "ZL_����ҽ��ִ��_Plan(" & lngҽ��ID & "," & lng���ͺ� & ",0)"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    err.Clear: On Error GoTo 0
    
    Call LoadPatients 'Ҫ����ִ��״̬
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecRefuse()
'���ܣ��ܾ�ִ��
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lngִ�п���ID As Long
    Dim strSQL As String, blnTrans As Boolean
    Dim str��� As String, strTextInput As String
    
    With rptPati.SelectedRows(0)
        '����ִ�л���ִ�в�����ܾ�
        If .Record(col_ִ��״̬).Value = 2 Then
            MsgBox "��ִ����Ŀ��ǰ�Ѿ��ܾ�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Record(col_ִ��״̬).Value = 3 Then
            MsgBox "��ִ����Ŀ��ǰ����ִ�У����ܾܾ���", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Record(col_ִ��״̬).Value = 1 Then
            MsgBox "��ִ����Ŀ��ǰ�Ѿ�ִ�У����ܾܾ���", vbInformation, gstrSysName
            Exit Sub
        End If
        '�ѱ����Ĳ��˲�����ܾ�
        If .Record(col_ִ�й���).Value <> 0 Then
            MsgBox "�ò����Ѿ����������ܾܾ���", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str��� = zlCommFun.ShowMsgBox("�ܾ�ִ��", "�������д�ܾ�ִ�е�ԭ��", _
            "ȷ��(&O),?ȡ��(&C)", Me, vbQuestion, , , , , , _
            "�ܾ�ԭ��(&B)", 50, strTextInput, , True)
        If str��� = "" Then Exit Sub
            
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        If .Record(col_ִ�п���).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lngִ�п���ID = cboDept.ItemData(cboDept.ListIndex)
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
        If lngִ�п���ID <> 0 Then
            strSQL = "Zl_����ҽ������_���ұ��(" & lngҽ��ID & "," & lng���ͺ� & "," & lngִ�п���ID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        strSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & lngҽ��ID & "," & lng���ͺ� & ",NULL,NULL,NULL,'" & strTextInput & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    gcnOracle.CommitTrans: blnTrans = False
    err.Clear: On Error GoTo 0
     
    With rptPati.SelectedRows(0)
        Call ZLHIS_CIS_015(mclsMipModule, .Record(col_����Id).Value, .Record(col_����).Value, .Record(col_��ʶ��).Value, , 2, .Record(col_��ҳID).Value, _
         .Record(col_����id).Value, .Record(col_����).Value, .Record(col_����).Caption, , .Record(col_����).Value, lngҽ��ID, .Record(col_��Ч).Value, _
         .Record(COL_�������).Value, .Record(COL_��������).Value, .Record(COL_������ĿID).Value, .Record(col_����).Value)
    End With
    
    Call LoadPatients 'Ҫ����ִ��״̬
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecRestore()
'���ܣ�ȡ���ܾ�ִ��
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, i As Long
    
    With rptPati.SelectedRows(0)
        '����ִ�л���ִ�в�����ܾ�
        If .Record(col_ִ��״̬).Value <> 2 Then
            MsgBox "��ִ����Ŀû�б��ܾ�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪȡ���ܾ�ִ�и���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
    End With
    
    strSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & lngҽ��ID & "," & lng���ͺ� & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    err.Clear: On Error GoTo 0
    
    Call LoadPatients 'Ҫ����ִ��״̬
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckExecuteLog(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long) As String
'���ܣ�����Ӧ����ҽ����ִ�������¼
'���أ������ִ�������¼���ߴ���δ�ﵽҪ��������򷵻���ʾ��Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.��������,Sum(B.��������) as ��������" & _
        " From ����ҽ������ A,����ҽ��ִ�� B Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And A.ҽ��ID=[1] And A.���ͺ�=[2]" & _
        " Group by A.��������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckExecuteLog", lngҽ��ID, lng���ͺ�)
    If rsTmp.EOF Then
        CheckExecuteLog = "��û�м�¼ִ�����"
    ElseIf Nvl(rsTmp!��������, 0) < Nvl(rsTmp!��������, 0) Then
        CheckExecuteLog = "��ִ������ " & Nvl(rsTmp!��������, 0) & " û�дﵽҪ������� " & Nvl(rsTmp!��������, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncExecFinish()
'���ܣ�ȷ��ִ�����
    Dim rsTmp As New ADODB.Recordset
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long, strNos As String, strҽ��IDs As String
    Dim lng���ID As Long, lng����ID As Long, blnTmp As Boolean
    Dim strSQL As String, strTest As String
    Dim str����Del As String, strִ�� As String
    Dim str��� As String, int��� As Integer, strLabel As String
    Dim cnNew As ADODB.Connection, i As Long
    Dim strUserName As String, strOwner As String, blnTrans As Boolean
    Dim rptRowChild As ReportRow
    Dim blnIsAbnormal As Boolean
    Dim lng�����ID As Long
    Dim dateInput As Date
    Dim strSelect As String
    Dim strSelectInput As String
    Dim strTextInput As String
    Dim dat���ʱ�� As Date
    Dim datTmp As Date
    Dim lng��������� As Long
    
    Dim curMoney As Currency, str��� As String, str����� As String
    
    '�ж��Ƿ�����ִ��ģʽ�������ڵ����������е���
    If rptPati.Columns(col_ѡ��).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_ѡ��).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("Ҫ�Ե�ǰѡ���һ��������Ŀִ�������", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecFinishBatch
            End If
            Exit Sub
        End If
    End If
    
    With rptPati.SelectedRows(0)
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        lng���ID = .Record(col_���ID).Value
        lng����ID = .Record(col_����Id).Value
        
        '���Բ���дִ�����ֱ�����ִ��
        If .Record(col_�ۺ�״̬).Value = 1 Then
            MsgBox "��ִ����Ŀ��ǰ�Ѿ�ִ����ɡ�", vbInformation, gstrSysName
            Exit Sub
        End If
                
        '��鲡���Ƿ�ʼ���
        If Val(.Record(col_��˱�־).Value & "") >= 1 And mbyt������˷�ʽ = 1 Then
            MsgBox "�ò��˵ķ���������˽׶Σ����������ҽ���ͷ��á�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "1" And Mid(gstrҽ���˶�, 2, 1) = "1" Or _
            .Record(COL_�������).Value = "E" And .Record(COL_��������).Value = "8" And Mid(gstrҽ���˶�, 1, 1) = "1" Or _
            .Record(COL_�������).Value = "K" And Mid(gstrҽ���˶�, 1, 1) = "1" Then
            '��Ѫ��Ƥ��ҽ��û�˶Բ��������
            If .Record(COL_�˶���).Value & "" = "" Then
                MsgBox "����Ŀ��" & IIf(.Record(COL_��������).Value = "1", "Ƥ��", "��Ѫ") & "ҽ��������˶��˲�����ɡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
	'�����շ��жϣ����࣬�����ﲡ�ˣ��������ﲡ�� 0-����ǰ�߼���1-��ʾ����ֹ��2-�ɹ���֤ͨ��
        If Val(.Record(COL_���ӱ�־).Value) = 3 Then '�����ﲡ��
            lng��������� = NewOut�շ�(lngҽ��ID)
            If lng��������� = 1 Then Exit Sub
        Else
            lng��������� = 0
        End If
        
        If lng��������� = 0 Then

        blnIsAbnormal = False
        '�Ƿ��������δ�շѲ��˵���Ŀ
        If .Record(col_��¼����).Value = 1 Then
            '���ʻ���,����ִ�к��Զ����
            If Not ItemHaveCash(IIf(.Record(col_��Դ).Value = "סԺ", 2, 1), Not .ParentRow.GroupRow, _
                lngҽ��ID, .Record(col_���ID).Value, .Record(col_���ͺ�).Value, _
                .Record(COL_�������).Value, .Record(col_���ݺ�).Value, .Record(col_��¼����).Value, .Record(col_�������).Value, _
                0, .Record(COL_����ת��).Value = 1, .Record(col_����ʱ��).Value, , , blnIsAbnormal) Then
                
                '�жϵ����Ƿ��쳣
                If blnIsAbnormal Then MsgBox "�ò��˻������쳣���ã����顣", vbInformation, gstrSysName: Exit Sub
                
                If Not mblnδ�շ���� Then
                    If gblnִ��ǰ�Ƚ��� Then
                        '���һ��ִ�л��ߵ���ִ�е�ҽ���ַ���
                        If Not .ParentRow.GroupRow Or (.Childs.Count = 0 And .ParentRow.GroupRow) Then
                            strҽ��IDs = lngҽ��ID
                        Else
                            For Each rptRowChild In .Childs
                                strҽ��IDs = strҽ��IDs & IIf(strҽ��IDs = "" Or rptRowChild.Record(col_ҽ��ID).Value & "" = "", "", ",") & rptRowChild.Record(col_ҽ��ID).Value
                            Next
                        End If
                        'ȡ���￨����
                        lng�����ID = Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����)
                    Else
                    
                        MsgBox "�ò��˻�����δ�շѵķ��ã����顣", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
            '�ж��Ƿ������ʵķѸ�����ʾ
            If Not mblnδ�շ���� Then
                blnTmp = Check���ʷ���(Not .ParentRow.GroupRow, lngҽ��ID, IIf(lng���ID = 0, lngҽ��ID, lng���ID), .Record(COL_�������).Value, IIf(.Record(col_��Դ).Value = "סԺ", 2, 1), .Record(col_���ݺ�).Value)
                If blnTmp Then
                    If MsgBox("�ò��˴������ʻ��˷ѵķ��ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
        End If
        
        
        '��鱨�����д�������Ӧ�˲������ݣ������б�����������Ҫ��д
        If .Record(col_�ļ�ID).Value <> 0 And .Record(col_������).Value <> 0 Then
            i = CheckEPRReport(IIf(.Record(COL_�������).Value = "C" Or (.Record(COL_�������).Value = "D" And lng���ID <> 0), lng���ID, lngҽ��ID), lng����ID, True, .Record(col_�ۺ�״̬).Value)
            If InStr(mstrPrivs, "ֱ��ִ�����") > 0 Then
                If i = 2 Then
                    If MsgBox("����Ŀ�ı�����д�����ݵ���û����ɣ�������ɱ�����ټ�����" & _
                        vbCrLf & vbCrLf & "���߿���ɾ�����÷�δ��ɵı��沢������Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    str����Del = "Zl_���Ӳ�����¼_Delete(" & lng����ID & ")"
                End If
            Else
                If i = 0 Then
                    MsgBox "����Ŀ�ı��滹û����д��������д�����ټ�����", vbInformation, gstrSysName
                    Exit Sub
                ElseIf i = 2 Then
                    MsgBox "����Ŀ�ı�����д�����ݵ���û����ɣ�������ɱ�����ټ�����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        If .Record(col_ִ�п���).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            strִ�� = "Zl_����ҽ������_���ұ��(" & lngҽ��ID & "," & lng���ͺ� & "," & cboDept.ItemData(cboDept.ListIndex) & ")"
        End If
    End With
        
    On Error GoTo errH
    
    '�ж��Ƿ�Ƥ��,����д���
    strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        '�Ѿ���д��Ƥ�Խ��������д
        If rsTmp!������� = "E" And Nvl(rsTmp!��������) = "1" And IsNull(rsTmp!Ƥ�Խ��) Then
            '���������֤
            If mblnƤ����֤ Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "����дƤ�Խ��ǰ�������������û�����������������֤��", glngSys, pҽ������վ, "ȷ��ִ�����", cnNew)
                If strUserName = "" Then Exit Sub
            End If
            '����
            For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(0), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(0), ",")(i) & "|0"
            Next
            '����
            For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(1), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(1), ",")(i) & "|0|2"
            Next
            strSelect = Mid(strSelect, 2)
            
            '��дƤ�Խ��
            str��� = zlCommFun.ShowMsgBox("Ƥ�Խ��", rptPati.SelectedRows(0).Record(col_����).Value & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", _
            "ȷ��(&O),?ȡ��(&C)", Me, vbQuestion, "Ƥ��ʱ��", dateInput, "yyyy-MM-dd HH:mm", "Ƥ�Խ��(&P):" & strSelect, strSelectInput, _
            "������Ӧ(&F)", 50, strTextInput, , True)
            
            If str��� = "" Then Exit Sub
            If strSelectInput = "" Then Exit Sub
            Call GetTestLabel(rsTmp!�걾��λ, strSelectInput, strLabel, int���)
            strTest = "ZL_����ҽ����¼_Ƥ��(" & lngҽ��ID & ",'" & strLabel & "'," & int��� & _
                        ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
        End If
    Else
        MsgBox "��Ӧ��ҽ����¼�����ڣ��޷���ɲ�����", vbInformation, gstrSysName
        Exit Sub
    End If

    '----
    With rptPati.SelectedRows(0)
        'ֻ�����ʷ���
        If .Record(col_��¼����).Value = 2 Then
            curMoney = GetAdviceMoney(IIf(lng���ID = 0, lngҽ��ID, lng���ID), lngҽ��ID, lng���ͺ�, str���, str�����, Not .ParentRow.GroupRow, _
                   IIf(.Record(col_��Դ).Value = "סԺ" And .Record(col_�������).Value = 0, 2, 1))
            If curMoney > 0 Then
                'סԺ��Ժ���˷��ÿ���
                If .Record(col_��Դ).Value = "סԺ" Then
                    If Not PatiCanBilling(.Record(col_����Id).Value, .Record(col_��ҳID).Value, GetInsidePrivs(pҽ�����ѹ���), pҽ�����ѹ���) Then Exit Sub
                End If
                '���ʱ���
                If InitObjPublicExpense Then
                    If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, pҽ������վ, "", .Record(col_���ݺ�).Value, GetInsidePrivs(pҽ�����ѹ���), Val(.Record(col_����id).Value)) = False Then Exit Sub
                End If
                    
                '����һ��ͨ���������֤,ֻ���������ʷ���
                If gdblԤ��������鿨 <> 0 And _
                    (.Record(col_��Դ).Value <> "סԺ" Or .Record(col_��Դ).Value = "סԺ" And .Record(col_�������).Value = 1) Then
                    If Not zlDatabase.PatiIdentify(Me, glngSys, .Record(col_����Id).Value, curMoney, , , , IIf(-1 * gdblԤ��������鿨 >= Val(curMoney), False, True), , , (gdblԤ��������鿨 <> 0), (2 = gdblԤ��������鿨)) Then Exit Sub
                End If
            End If
        End If
    
        '�ϸ�Ҫ���¼ִ�е����
        If mblnExeLog Then
            strSQL = CheckExecuteLog(lngҽ��ID, lng���ͺ�)
            If strSQL <> "" Then
                MsgBox "��ִ����Ŀ" & strSQL & "���������ִ�С�", vbInformation, gstrSysName
                Exit Sub
            Else
                If strTest = "" And str����Del = "" Then
                    If MsgBox("ȷ�ϸ�ִ����Ŀִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        Else
            If strTest = "" And str����Del = "" Then
                If MsgBox("ȷ�ϸ�ִ����Ŀִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        dat���ʱ�� = zlDatabase.Currentdate
        datTmp = dat���ʱ��
        blnTmp = frmSelectTime.ShowMe(Me, dat���ʱ��, datTmp, Me, 1)
        If Not blnTmp Then
            Exit Sub
        End If
        
        
        '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������,�������ݺţ�����ҽ��ID��ȡ����δ�շѵ��ݻ�δ��˵ļ��ʵ�
        If gblnִ��ǰ�Ƚ��� And strҽ��IDs <> "" Then
            If mobjSquareCard.zlSquareAffirm(Me, pҽ������վ, mstrPrivs, lng����ID, lng�����ID, False, , , strҽ��IDs) = False Then
                Exit Sub
            End If
        End If
        
        strSQL = "ZL_����ҽ��ִ��_Finish(" & lngҽ��ID & "," & lng���ͺ� & "," & _
            "Null," & IIf(Not .ParentRow.GroupRow, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngDept & ",0,to_date('" & dat���ʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
   
    End With
    
    If strTest <> "" And Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        cnNew.BeginTrans: blnTrans = True
        
        If strִ�� <> "" Then
            Call SQLTest(App.ProductName, Me.Caption, strִ��)
            cnNew.Execute strOwner & "." & strִ��, , adCmdStoredProc
            Call SQLTest
        End If
        
        Call SQLTest(App.ProductName, Me.Caption, strTest)
        cnNew.Execute strOwner & "." & strTest, , adCmdStoredProc
        Call SQLTest
        
        If str����Del <> "" Then
            Call SQLTest(App.ProductName, Me.Caption, str����Del)
            cnNew.Execute strOwner & "." & str����Del, , adCmdStoredProc
            Call SQLTest
        End If
        
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest
        
        cnNew.CommitTrans: blnTrans = False
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTrans = True
            If strִ�� <> "" Then
                Call zlDatabase.ExecuteProcedure(strִ��, Me.Caption)
            End If
            If strTest <> "" Then
                Call zlDatabase.ExecuteProcedure(strTest, Me.Caption)
            End If
            If str����Del <> "" Then
                Call zlDatabase.ExecuteProcedure(str����Del, Me.Caption)
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    Call LoadPatients 'Ҫ����ִ��״̬
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    If blnTrans Then cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Sub FuncExecFinishBatch()
'���ܣ���һ������ѡ��Ķ����Ŀȷ��ִ�����
    Dim rsTmp As New ADODB.Recordset
    Dim arrSQL As Variant, strSQL As String, i As Long
    Dim str��� As String, int��� As Integer, strLabel As String, strNos As String, strҽ��IDs As String
    Dim lng����ID As Long, intThing As Integer
    Dim blnTmp As Boolean, blnTest As Boolean, blnTrans As Boolean
    
    Dim cnNew As ADODB.Connection
    Dim strUserName As String, strOwner As String
    
    Dim rsPati As ADODB.Recordset
    Dim curMoney As Currency, str��� As String, str����� As String
    Dim strPatiIDs As String, blnIsMany As Boolean         'blnIsMany�Ƿ�ѡ�˶�����˵ĵ���
    Dim blnIsAbnormal As Boolean
    Dim lng�����ID  As Long
    Dim dateInput As Date
    Dim strMsgAduit As String
    Dim strMsg As String
    Dim strSelect As String, j As Long
    Dim strSelectInput As String
    Dim strTextInput As String
    Dim dat���ʱ�� As Date
    Dim datTmp As Date
    Dim lng��������� As Long
    
    On Error GoTo errH
    
    Set rsPati = New ADODB.Recordset
    rsPati.Fields.Append "��Դ", adVarChar, 10
    rsPati.Fields.Append "��¼����", adBigInt
    rsPati.Fields.Append "�������", adBigInt
    rsPati.Fields.Append "����ID", adBigInt
    rsPati.Fields.Append "��ҳID", adBigInt
    rsPati.Fields.Append "����ID", adBigInt
    rsPati.Fields.Append "��ID", adVarChar, 2000
    rsPati.Fields.Append "ҽ��ID", adVarChar, 2000
    rsPati.Fields.Append "���ͺ�", adVarChar, 2000
    rsPati.Fields.Append "NO", adVarChar, 4000
    
    rsPati.CursorLocation = adUseClient
    rsPati.LockType = adLockOptimistic
    rsPati.CursorType = adOpenStatic
    rsPati.Open
    
    arrSQL = Array()
    
    dat���ʱ�� = zlDatabase.Currentdate
    datTmp = dat���ʱ��
    blnTmp = frmSelectTime.ShowMe(Me, dat���ʱ��, datTmp, Me, 1)
    If Not blnTmp Then
        Exit Sub
    End If
    
    'ȡ���￨����
    If gblnִ��ǰ�Ƚ��� Then
        lng�����ID = Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����)
    End If
    
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(col_ѡ��).Checked Then
                With rptPati.Rows(i).Record
                    '���Բ���дִ�����ֱ�����ִ��
                    If .Item(col_�ۺ�״̬).Value = 1 Then
                        MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """��ǰ�Ѿ�ִ����ɡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If .Item(COL_�������).Value = "E" And .Item(COL_��������).Value = "1" And Mid(gstrҽ���˶�, 2, 1) = "1" Or _
                        .Item(COL_�������).Value = "E" And .Item(COL_��������).Value = "8" And Mid(gstrҽ���˶�, 1, 1) = "1" Or _
                        .Item(COL_�������).Value = "K" And Mid(gstrҽ���˶�, 1, 1) = "1" Then
                        '��Ѫ��Ƥ��ҽ��û�˶Բ��������
                        If .Item(COL_�˶���).Value & "" = "" Then
                            strMsgAduit = strMsgAduit & "," & .Item(col_���ݺ�).Value
                        End If
                    End If
                    
                    '��鲡���Ƿ�ʼ���
                    If Val(.Item(col_��˱�־).Value & "") >= 1 And mbyt������˷�ʽ = 1 Then
                        strMsg = strMsg & "," & .Item(col_���ݺ�).Value
                    Else
                        '�ϸ�Ҫ���¼ִ�е����
                        If mblnExeLog Then
                            strSQL = CheckExecuteLog(.Item(col_ҽ��ID).Value, .Item(col_���ͺ�).Value)
                            If strSQL <> "" Then
                                MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """" & strSQL & "���������ִ�С�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                            
                        If .Item(COL_����ת��).Value = 1 Then
                            MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """�������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
			'�����շ��жϣ����࣬�����ﲡ�ˣ��������ﲡ�� 0-����ǰ�߼���1-��ʾ����ֹ��2-�ɹ���֤ͨ��
                        If Val(.Item(COL_���ӱ�־).Value) = 3 Then '�����ﲡ��
                            lng��������� = NewOut�շ�(Val(.Item(col_ҽ��ID).Value))
                            If lng��������� = 1 Then Exit Sub
                        Else
                            lng��������� = 0
                        End If
        
                        If lng��������� = 0 Then

                        If .Item(col_��¼����).Value = 1 Then
                            '���ܼ��ʻ���,����ִ�к��Զ����
                            If Not ItemHaveCash(IIf(.Item(col_��Դ).Value = "סԺ", 2, 1), Not rptPati.Rows(i).ParentRow.GroupRow, _
                                .Item(col_ҽ��ID).Value, .Item(col_���ID).Value, .Item(col_���ͺ�).Value, .Item(COL_�������).Value, _
                                .Item(col_���ݺ�).Value, .Item(col_��¼����).Value, .Item(col_�������).Value, 0, .Item(COL_����ת��).Value = 1, .Item(col_����ʱ��).Value, , , blnIsAbnormal) Then
                                
                                '�жϵ����Ƿ��쳣
                                If blnIsAbnormal Then MsgBox "�ò��˻������쳣���ã����顣", vbInformation, gstrSysName: Exit Sub
                                
                                '�Ƿ��������δ�շѲ��˵���Ŀ
                                If Not mblnδ�շ���� Then
                                    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
                                    If gblnִ��ǰ�Ƚ��� Then
                                        '��ȡ���˵�����ѡ��ҽ��ID�ַ���(�����жϲ���ID��Ϊ��һ�����˶��ŵ���ֻ����һ�νӿڣ�
                                        If InStr("," & strPatiIDs & ",", "," & .Item(col_����Id).Value & ",") = 0 Then
                                            strҽ��IDs = GetSelectAdviceIDs(Val(.Item(col_����Id).Value), blnIsMany)
                                            '����Ƕ�������������㣬������ʾ�����������
                                            If blnIsMany Then
                                                MsgBox "�����漰���ý��㣬������������������ɡ�", vbInformation, Me.Caption
                                                Exit Sub
                                            End If
                                            
                                            strPatiIDs = strPatiIDs & "," & .Item(col_����Id).Value
                                        End If
                                    Else
                                        MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """������δ�շѵķ��ã����顣", vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                            End If
                            '�ж��Ƿ������ʵķѸ�����ʾ
                            If Not mblnδ�շ���� Then
                                blnTmp = Check���ʷ���(Not rptPati.Rows(i).ParentRow.GroupRow, Val(.Item(col_ҽ��ID).Value), _
                                            IIf(Val(.Item(col_���ID).Value) = 0, Val(.Item(col_ҽ��ID).Value), Val(.Item(col_���ID).Value)), _
                                            .Item(COL_�������).Value, IIf(.Item(col_��Դ).Value = "סԺ", 2, 1), .Item(col_���ݺ�).Value)
                                If blnTmp Then
                                    If MsgBox("����""" & .Item(col_����).Value & "��" & .Item(col_���ݺ�).Value & "��" & """����Ŀ""" & .Item(col_����).Value & """�������ʻ��˷ѵķ��ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                                End If
                            End If
                        End If
                              End If                        

                        '��鱨�����д�������Ӧ�˲������ݣ������б�����������Ҫ��д
                        If .Item(col_�ļ�ID).Value <> 0 And .Item(col_������).Value <> 0 Then
                            intThing = CheckEPRReport(IIf(.Item(COL_�������).Value = "C" Or (.Item(COL_�������).Value = "D" And Val(.Item(col_���ID).Value) <> 0), .Item(col_���ID).Value, .Item(col_ҽ��ID).Value), lng����ID, True, .Item(col_�ۺ�״̬).Value)
                            If InStr(mstrPrivs, "ֱ��ִ�����") > 0 Then
                                If intThing = 2 Then
                                    If MsgBox("����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """�ı�����д�����ݵ���û����ɣ�������ɱ�����ټ�����" & _
                                        vbCrLf & vbCrLf & "���߿���ɾ�����÷�δ��ɵı��沢������Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�����¼_Delete(" & lng����ID & ")"
                                End If
                            Else
                                If intThing = 0 Then
                                    MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """�ı��滹û����д��������д�����ټ�����", vbInformation, gstrSysName
                                    Exit Sub
                                ElseIf intThing = 2 Then
                                    MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """�ı�����д�����ݵ���û����ɣ�������ɱ�����ټ�����", vbInformation, gstrSysName
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                        '�ж��Ƿ�Ƥ��,����д���
                        strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Item(col_ҽ��ID).Value)
                        If Not rsTmp.EOF Then
                            '�Ѿ���д��Ƥ�Խ��������д
                            If rsTmp!������� = "E" And Nvl(rsTmp!��������) = "1" And IsNull(rsTmp!Ƥ�Խ��) Then
                                '���������֤
                                If mblnƤ����֤ And cnNew Is Nothing Then
                                    Set cnNew = New ADODB.Connection
                                    strUserName = zlDatabase.UserIdentify(Me, "����дƤ�Խ��ǰ�������������û�����������������֤��", glngSys, pҽ������վ, "ȷ��ִ�����", cnNew)
                                    If strUserName = "" Then Exit Sub
                                End If
                                strSelect = "": strSelectInput = ""
                                '����
                                For j = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(0), ","))
                                    strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(0), ",")(j) & "|0"
                                Next
                                '����
                                For j = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(1), ","))
                                    strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(1), ",")(j) & "|0|2"
                                Next
                                strSelect = Mid(strSelect, 2)
                                '��дƤ�Խ��
                                str��� = zlCommFun.ShowMsgBox("Ƥ�Խ��", .Item(col_����).Value & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", _
                                        "ȷ��(&O),?ȡ��(&C)", Me, vbQuestion, "Ƥ��ʱ��", dateInput, "yyyy-MM-dd HH:mm", "Ƥ�Խ��(&P):" & strSelect, strSelectInput, _
                                        "������Ӧ(&F)", 50, strTextInput, , True)
                                If str��� = "" Then Exit Sub
                                
                                blnTest = True
                                If strSelectInput = "" Then Exit Sub
                                Call GetTestLabel(rsTmp!�걾��λ, strSelectInput, strLabel, int���)
                                
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Ƥ��(" & .Item(col_ҽ��ID).Value & ",'" & strLabel & "'," & int��� & _
                                                    ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
                            End If
                        Else
                            MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """��Ӧ��ҽ����¼�����ڣ��޷���ɲ�����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_����ҽ��ִ��_Finish(" & .Item(col_ҽ��ID).Value & "," & _
                            .Item(col_���ͺ�).Value & ",Null," & IIf(Not rptPati.Rows(i).ParentRow.GroupRow, 1, 0) & "," & _
                            "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngDept & ",0,to_date('" & dat���ʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
                           
                        '�ռ���ͬ�Ĳ�����Ϣ
                        rsPati.Filter = "��Դ='" & .Item(col_��Դ).Value & "' And �������=" & .Item(col_�������).Value & _
                            " And ��¼����=" & .Item(col_��¼����).Value & " And ����ID=" & .Item(col_����Id).Value & _
                            " And ��ҳID=" & .Item(col_��ҳID).Value & " And ����ID=" & .Item(col_����id).Value
                        If rsPati.EOF Then
                            rsPati.AddNew
                            rsPati!��Դ = CStr(.Item(col_��Դ).Value)
                            rsPati!��¼���� = Val(.Item(col_��¼����).Value)
                            rsPati!������� = Val(.Item(col_�������).Value)
                            rsPati!����ID = Val(.Item(col_����Id).Value)
                            rsPati!��ҳID = Val(.Item(col_��ҳID).Value)
                            rsPati!����ID = Val(.Item(col_����id).Value)
                            rsPati.Update
                        End If
                        rsPati!��ID = Nvl(rsPati!��ID) & "," & IIf(.Item(col_���ID).Value = 0, .Item(col_ҽ��ID).Value, .Item(col_���ID).Value)
                        rsPati!ҽ��ID = Nvl(rsPati!ҽ��ID) & "," & .Item(col_ҽ��ID).Value
                        rsPati!���ͺ� = Nvl(rsPati!���ͺ�) & "," & .Item(col_���ͺ�).Value
                        rsPati!NO = Nvl(rsPati!NO) & "," & .Item(col_���ݺ�).Value
                        rsPati.Update
                    End If
                        
                    
                End With
            End If
        End If
    Next
    
    If strMsgAduit <> "" Or strMsg <> "" Then
        If strMsg <> "" And strMsgAduit <> "" Then
            strMsg = "���µ��ݺŵĲ��˷���������˽׶Σ����������ҽ���ͷ��ã�" & vbCrLf & Mid(strMsg, 2) & "��" & vbCrLf & _
                    "���µ��ݺ���������Ѫ����Ƥ����Ŀ������˶Ժ���ִ����ɣ�" & vbCrLf & Mid(strMsgAduit, 2) & "��"
        ElseIf strMsgAduit <> "" Then
            strMsg = "���µ��ݺ���������Ѫ����Ƥ����Ŀ������˶Ժ���ִ����ɣ�" & vbCrLf & Mid(strMsgAduit, 2) & "��"
        Else
            strMsg = "���µ��ݺŵĲ��˷���������˽׶Σ����������ҽ���ͷ��ã�" & vbCrLf & Mid(strMsg, 2) & "��"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    
    If UBound(arrSQL) = -1 Then Exit Sub
    
    '������˵ķ��ü��ͱ���
    rsPati.Filter = "��¼���� = 2"
    Do While Not rsPati.EOF
        curMoney = GetAdviceMoney(Mid(rsPati!��ID, 2), Mid(rsPati!ҽ��ID, 2), Mid(rsPati!���ͺ�, 2), str���, str�����, False, _
                IIf(rsPati!��Դ = "סԺ" And rsPati!������� = 0, 2, 1))
        If curMoney > 0 Then
            'סԺ��Ժ���˷��ÿ���
            If rsPati!��Դ = "סԺ" Then
                If Not PatiCanBilling(rsPati!����ID, rsPati!��ҳID, GetInsidePrivs(pҽ�����ѹ���), pҽ�����ѹ���) Then Exit Sub
            End If
            '���ʱ���
            If InitObjPublicExpense Then
                If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, pҽ������վ, "", Mid(rsPati!NO & "", 2), GetInsidePrivs(pҽ�����ѹ���), Val(rsPati!����ID & "")) = False Then Exit Sub
            End If
            '����һ��ͨ���������֤
            If gdblԤ��������鿨 <> 0 And (rsPati!��Դ <> "סԺ" Or rsPati!��Դ = "סԺ" And rsPati!������� = 1) Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, rsPati!����ID, curMoney, , , , IIf(-1 * gdblԤ��������鿨 >= Val(curMoney), False, True), , , (gdblԤ��������鿨 <> 0), (2 = gdblԤ��������鿨)) Then Exit Sub
            End If
        End If
        rsPati.MoveNext
    Loop
    
    If gblnִ��ǰ�Ƚ��� And strҽ��IDs <> "" Then
        If mobjSquareCard.zlSquareAffirm(Me, pҽ������վ, mstrPrivs, Val(Mid(strPatiIDs, 2)), lng�����ID, False, , , strҽ��IDs) = False Then
            Exit Sub
        End If
    End If

    '�ύSQLִ��
    If blnTest And Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        cnNew.BeginTrans: blnTrans = True
            
        For i = 0 To UBound(arrSQL)
            Call SQLTest(App.ProductName, Me.Caption, arrSQL(i))
            cnNew.Execute strOwner & "." & arrSQL(i), , adCmdStoredProc
            Call SQLTest
        Next
        
        cnNew.CommitTrans: blnTrans = False
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Call LoadPatients 'Ҫ����ִ��״̬
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
errNew:
    cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function GetSelectAdviceIDs(ByVal lngPatiID As Long, ByRef blnIsMany As Boolean) As String
'���ܣ����ݲ��˻������ִ������ҽ��ID�ַ���
    Dim i As Long, strSelectAdvices As String
    Dim rptRowChild As ReportRow
    
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(col_ѡ��).Checked Then
                With rptPati.Rows(i)
                    '�ж��Ƿ�ѡ�ж�����˵ĵ���
                    If lngPatiID <> Val(.Record(col_����Id).Value) Then blnIsMany = True: Exit Function
                    
                    If .Record(col_��¼����).Value = 1 Or .Record(col_��¼����).Value = 2 And (.Record(col_��Դ).Value <> "סԺ" _
                            Or .Record(col_��Դ).Value = "סԺ" And .Record(col_�������).Value = 1) Then
                        If Val(.Record(col_����Id).Value) = lngPatiID Then
                            '���һ��ִ�л��ߵ���ִ�е�ҽ���ַ���
                            If Not .ParentRow.GroupRow Or (.Childs.Count = 0 And .ParentRow.GroupRow) Then
                                strSelectAdvices = strSelectAdvices & IIf(.Record(col_ҽ��ID).Value & "" = "", "", ",") & .Record(col_ҽ��ID).Value
                            Else
                                For Each rptRowChild In .Childs
                                    strSelectAdvices = strSelectAdvices & IIf(rptRowChild.Record(col_ҽ��ID).Value & "" = "", "", ",") & rptRowChild.Record(col_ҽ��ID).Value
                                Next
                            End If
                        End If
                    End If
                End With
            End If
        End If
    Next
    GetSelectAdviceIDs = Mid(strSelectAdvices, 2)
End Function

Private Sub FuncExecCancel()
'���ܣ�ȡ��ִ�����
    Dim lng��ID As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim str������� As String, strSQL As String, byt��Դ As Byte
    Dim strOwner As String, strUserName As String
    Dim cnNew As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim i As Long
    
    '�ж��Ƿ�����ִ��ģʽ�������ڵ����������е���
    If rptPati.Columns(col_ѡ��).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_ѡ��).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("Ҫ�Ե�ǰѡ���һ��������Ŀȡ��ִ�������", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecCancelBatch
            End If
            Exit Sub
        End If
    End If

    With rptPati.SelectedRows(0)
        
        '��������ִ�вſ���ȡ��
        If .Record(col_�ۺ�״̬).Value <> 1 Then
            MsgBox "��ִ����Ŀ��ǰ��������ִ��״̬������ȡ��ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��鲡���Ƿ�ʼ���
        If Val(.Record(col_��˱�־).Value & "") >= 1 And mbyt������˷�ʽ = 1 Then
            MsgBox "�ò��˵ķ���������˽׶Σ����������ҽ���ͷ��á�", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        str������� = .Record(COL_�������).Value
        lng��ID = IIf(.Record(col_���ID).Value = 0, .Record(col_ҽ��ID).Value, .Record(col_���ID).Value)
        
        
        If Val(.Record(col_��¼����).Value) <> 1 Then
            If .Record(col_��Դ).Value = "סԺ" And Val(.Record(col_�������).Value) = 0 Then
                byt��Դ = 2
            Else
                byt��Դ = 1
            End If
            '���ý����ж�
            If Not ItemCanCancel(lngҽ��ID, lng���ͺ�, lng��ID, str�������, Not .ParentRow.GroupRow, .Record(COL_����ת��).Value = 1, byt��Դ) Then Exit Sub
        End If
    End With
            
    If MsgBox("ȷʵҪ����ִ����Ŀȡ��ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '�ж��Ƿ�Ƥ��,����д���
    strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        '�Ѿ���д��Ƥ�Խ��������д
        If rsTmp!������� = "E" And Nvl(rsTmp!��������) = "1" And Not IsNull(rsTmp!Ƥ�Խ��) Then
            '���������֤
            If mblnƤ����֤ Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "��ȡ�����Ƥ��ҽ��ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "Ƥ��ҽ�����", cnNew)
                If strUserName = "" Then Exit Sub
            End If
            strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngҽ��ID & "," & lng���ͺ� & ",1," & _
                IIf(Not rptPati.SelectedRows(0).ParentRow.GroupRow, 1, 0) & "," & mlngDept & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngҽ��ID & "," & lng���ͺ� & ",Null," & _
                IIf(Not rptPati.SelectedRows(0).ParentRow.GroupRow, 1, 0) & "," & mlngDept & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        End If
 
    End If
    If Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)

        On Error GoTo errNew
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End If
    
    Call LoadPatients 'Ҫ����ִ��״̬
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function FuncThingNew(Optional ByVal blnRefresh As Boolean = True) As Boolean
    Dim lng����ID As Long, lngִ�п���ID As Long
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    
    With rptPati.SelectedRows(0)
        If .Record(col_�ۺ�״̬).Value = 1 Then '����Ͷ���ִͬ��״̬
            MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
        
        lng����ID = cboDept.ItemData(cboDept.ListIndex)
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        If .Record(col_ִ�п���).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lngִ�п���ID = cboDept.ItemData(cboDept.ListIndex)
        End If
    
        On Error Resume Next
        If frmTechnicLog.ShowMe(Me, pҽ������վ, lng����ID, lngҽ��ID, lng���ͺ�, Not .ParentRow.GroupRow, , lngִ�п���ID, .Record(col_�����).Value, mstrPrivs) Then
            err.Clear: On Error GoTo 0
            If blnRefresh Then Call LoadPatients '����Ҫ����ִ��״̬
            FuncThingNew = True
        End If
    End With
End Function

Private Sub FuncThingModi()
    Dim lng����ID As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim strִ��ʱ�� As String, lngִ�п���ID As Long
        
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub 'ֻ�ܲ������һ��ִ��
    
    If Val(gstrҽ���˶�) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) <> "" Then
        MsgBox "��ҽ�����Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rptPati.SelectedRows(0)
        If .Record(col_�ۺ�״̬).Value = 1 Then '����Ͷ���ִͬ��״̬
            MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        lng����ID = cboDept.ItemData(cboDept.ListIndex)
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        strִ��ʱ�� = vsExec.Cell(flexcpData, vsExec.Row, 1)
        If .Record(col_ִ�п���).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lngִ�п���ID = cboDept.ItemData(cboDept.ListIndex)
        End If
        
        On Error Resume Next
        If frmTechnicLog.ShowMe(Me, pҽ������վ, lng����ID, lngҽ��ID, lng���ͺ�, Not .ParentRow.GroupRow, strִ��ʱ��, lngִ�п���ID, .Record(col_�����).Value, mstrPrivs) Then
            err.Clear: On Error GoTo 0
            Call LoadExecList(lngҽ��ID, lng���ͺ�)
        End If
    End With
End Sub

Private Sub FuncThingDel()
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strִ��ʱ�� As String, strSQL As String
    
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub 'ֻ�ܲ������һ��ִ��
    
    With rptPati.SelectedRows(0)
        If .Record(col_�ۺ�״̬).Value = 1 Then '����Ͷ���ִͬ��״̬
            MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(gstrҽ���˶�) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) <> "" Then
            MsgBox "��ҽ�����Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, gstrSysName
            Exit Sub
        End If
            
        If .Record(COL_����ת��).Value = 1 Then
            MsgBox "�ò��˵ı���" & Decode(.Record(col_��Դ).Value, "סԺ", "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
            
        If MsgBox("ȷʵҪɾ������ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lngҽ��ID = .Record(col_ҽ��ID).Value
        lng���ͺ� = .Record(col_���ͺ�).Value
        strִ��ʱ�� = vsExec.Cell(flexcpData, vsExec.Row, 1)
    
        strSQL = "ZL_����ҽ��ִ��_Delete(" & lngҽ��ID & "," & lng���ͺ� & "," & _
            "To_Date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')," & IIf(Not .ParentRow.GroupRow, 1, 0) & ",0," & mlngDept & ")"
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        err.Clear: On Error GoTo 0
        
        Call LoadPatients '����Ҫ����ִ��״̬
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsExec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
        
    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Manage_ThingAdd, "��¼ִ�����(&A)"
            .Add xtpControlButton, conMenu_Manage_ThingModi, "����ִ�����(&M)"
            .Add xtpControlButton, conMenu_Manage_ThingDel, "ɾ��ִ�����(&D)"
            .Add xtpControlButton, conMenu_Manage_ThingAudit, "�˶�"
            .Add xtpControlButton, conMenu_Manage_ThingDelAudit, "ȡ���˶�"
        End With
        
        vsExec.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub Set������Ŀ��������()
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallSetClinicCharge(Val(cboDept.ItemData(cboDept.ListIndex)), 1, Me, gcnOracle, glngSys, gstrDBUser, E�������, InStr(mstrPrivs, "������Ŀ��������") = 0)
End Sub

Private Function ShowBillAppend(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'���ܣ���ʾָ����ҽ���ĵ��ݸ�������
'���أ�blnExist=ҽ���Ƿ���ڵ��ݸ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngidx As Long

    blnExist = False
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    strSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order by ����"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(Val(rptPati.SelectedRows(0).Record(col_���ID).Value) = 0, Val(rptPati.SelectedRows(0).Record(col_ҽ��ID).Value), Val(rptPati.SelectedRows(0).Record(col_���ID).Value)))
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIf(.Text = "", "", vbCrLf) & rsTmp!��Ŀ & "��" & Nvl(rsTmp!����)
                lngidx = .Find(rsTmp!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
                If lngidx <> -1 Then
                    .SelStart = lngidx
                    .SelLength = Len(rsTmp!��Ŀ & "��")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            '��궨λ�ڵ�һ�����븽��
            rsTmp.MoveFirst
            lngidx = .Find(rsTmp!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngidx <> -1 Then .SelStart = lngidx + Len(rsTmp!��Ŀ & "��")
            
            Call zlControl.RTFSetFontSize(rtfAppend, IIf(mbytSize = 0, 9, 12))
        End With
        blnExist = True
    End If
    
    ShowBillAppend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'���ܣ����н��������ͳһ����
'������blnSetMainFont  �Ƿ��������������� �����������ӽ����л���
    If blnSetMainFont Then
        Call zlControl.SetPubFontSize(Me, mbytSize, "fraExec")
        Call SetControlPosition
        If Not mobjFrmBloodExe Is Nothing Then
            If mobjFrmBloodExe.Visible = True Then Call mobjFrmBloodExe.SetFontSize(IIf(mbytSize = 0, 9, 12))
        End If
    End If
        
    Select Case tbcSub.Selected.Tag
        Case "ҽ������"
            Call mclsExpenses.SetFontSize(mbytSize)
        Case "����ҽ��"
            Call mclsOutAdvices.SetFontSize(mbytSize)
        Case "סԺҽ��"
            Call mclsInAdvices.SetFontSize(mbytSize)
        Case "סԺ����"
            Call mclsInEPRs.SetFontSize(mbytSize)
        Case "���ﲡ��"
            Call mclsOutEPRs.SetFontSize(mbytSize)
        Case "����"
            Call mclsTends.SetFontSize(mbytSize)
        Case "������"
            Call mclsTendEPRs.SetFontSize(mbytSize)
        Case "�°滤��"
            Call mclsTendsNew.SetFontSize(mbytSize)
                Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
End Sub

Private Sub SetControlPosition()
'���ܣ�����������Ŀռ�λ�ô�С�Լ�����
    Dim lngVcDis As Long
    lngVcDis = IIf(mbytSize = 0, 20, 50)
    lblAdvice.Font.Size = IIf(mbytSize = 0, 9, 12)
    
    fraDiag.Height = Me.TextHeight("��") + 120
    
    fraExec.Top = fraDiag.Top + fraDiag.Height + 10
    lblCash.Height = fraExec.Height - lblCash.Top - 30
    lblRec.Height = lblCash.Height
    
    vsExec.Top = fraExec.Top + fraExec.Height + 20
    '�����ָ�����VSFlexGrid����Ϊ������ʾ
    vsExec.Height = Me.TextHeight("��") * 5
    
    picApplyInfo.Top = vsExec.Top
    picApplyInfo.Height = vsExec.Height
    rtfAppend.Top = lblApply.Top + lblApply.Height + 10
    rtfAppend.Height = picApplyInfo.Height - rtfAppend.Top
    '�����¼�picExec_Resize
    picExec.Height = vsExec.Top + vsExec.Height
    
    Call zlControl.SetPubCtrlPos(False, 0, lblDept, 20, cboDept)
    Call zlControl.SetPubCtrlPos(False, 0, lblFind, 20, PatiIdentify)
    PatiIdentify.Left = IIf(mbytSize = 0, 990, 1050)
    chkFilter.Height = PatiIdentify.Height
    Call zlControl.SetPubCtrlPos(True, 1, chkִ��״̬(0), lngVcDis, chkִ��״̬(2), lngVcDis, chkִ��״̬(4))
    Call zlControl.SetPubCtrlPos(False, 0, Image1(0), 10, chkִ��״̬(0), IIf(mbytSize = 0, 30, 20), Image1(1), 10, chkִ��״̬(1))
    Call zlControl.SetPubCtrlPos(False, 0, Image1(2), 10, chkִ��״̬(2), IIf(mbytSize = 0, 30, 20), Image1(3), 10, chkִ��״̬(3))
    Call zlControl.SetPubCtrlPos(False, 0, Image1(4), 20, chkִ��״̬(4))
    Call picPati_Resize
End Sub

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'���ܣ������յ�����Ϣ���������б���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    If Mid(rsMsg!���ѳ���, 4, 1) <> "1" Then Exit Sub
    
    If InStr("," & rsMsg!����IDs & ",", "," & cboDept.ItemData(cboDept.ListIndex) & ",") > 0 Or _
        InStr("," & rsMsg!������Ա & ",", "," & UserInfo.���� & ",") > 0 Then
        
        '�ж��б��Ƿ��Ѿ���������Ϣ��
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_��Ϣ).Value = rsMsg!���ͱ��� And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!����ID & "," & rsMsg!����id) Then
                    Exit Sub
                End If
            End If
        Next
        strSQL = "Select a.סԺ��, a.����, a.�Ա�, a.����, a.��ǰ���� As ����, a.���� From ������Ϣ A Where a.����id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!����ID))
        
        Call AddReportRow(rsMsg!����ID & "," & rsMsg!����id, rsMsg!����ID, rsMsg!����id, Nvl(rsTmp!����), Nvl(rsMsg!��Ϣ����), rsMsg!���ͱ��� & "", _
                rsMsg!���ȳ̶� & "", Format(rsMsg!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!ҵ���ʶ & "")
        rptNotify.Populate
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ParamArray arrInput() As Variant)
'���ܣ�����Ϣ�����б�������һ��
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim Index As Integer
    
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tagֵ
    Set objItem = objRecord.AddItem(""): objItem.Icon = 3
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '����
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '״̬������
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '��Ϣ���
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1   '���
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '����
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  'ҵ���ʶ
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LocatePati(ByVal strTag As String) As Boolean
'���ܣ�ͨ��reportControl��Record.Tagֵ��λ����
'������strTag ҽ��id_���ͺ�

    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    
    For Each objRow In rptPati.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow Then
            If InStr(objRow.Record.Tag & "_", "_" & strTag & "_") > 0 Then
                blnEnabled = timRefresh.Enabled
                timRefresh.Enabled = False '������������ˢ����������
                Set rptPati.FocusedRow = objRow 'ѡ��,��ʾ,[����Change�¼�]
                timRefresh.Enabled = blnEnabled
                LocatePati = True: Exit Function
            End If
        End If
    Next
End Function

Private Sub FuncExecCancelBatch()
'���ܣ���һ������ѡ��Ķ����Ŀȡ��ִ�����
    Dim lng��ID As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim str������� As String, strSQL As String, byt��Դ As Byte
    Dim strOwner As String, strUserName As String
    Dim cnNew As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL As Variant
    Dim strMsg As String
    Dim blnTrans As Boolean
    Dim blnGroupRow As Boolean
    
    On Error GoTo errH
    
    arrSQL = Array()
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(col_ѡ��).Checked Then
                With rptPati.Rows(i).Record
                    '���Բ���дִ�����ֱ�����ִ��
                    If Val(.Item(col_�ۺ�״̬).Value) <> 1 Then
                        MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """��ǰ��������ִ��״̬������ȡ��ִ�С�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    '��鲡���Ƿ�ʼ���
                    If Val(.Item(col_��˱�־).Value & "") >= 1 And mbyt������˷�ʽ = 1 Then
                        strMsg = strMsg & "," & .Item(col_���ݺ�).Value
                    Else
                        blnGroupRow = Not rptPati.Rows(i).ParentRow.GroupRow
                        '�����Ƿ�ת��
                        If .Item(COL_����ת��).Value = 1 Then
                            MsgBox "����""" & .Item(col_����).Value & """����Ŀ""" & .Item(col_����).Value & """�������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                      
                        lngҽ��ID = .Item(col_ҽ��ID).Value
                        lng���ͺ� = .Item(col_���ͺ�).Value
                        str������� = .Item(COL_�������).Value
                        lng��ID = IIf(.Item(col_���ID).Value = 0, .Item(col_ҽ��ID).Value, .Item(col_���ID).Value)
                        If Val(.Item(col_��¼����).Value) <> 1 Then
                            '���ý����ж�
                            byt��Դ = IIf(.Item(col_��Դ).Value = "סԺ" And Val(.Item(col_�������).Value) = 0, 2, 1)
                            If Not ItemCanCancel(lngҽ��ID, lng���ͺ�, lng��ID, str�������, blnGroupRow, .Item(COL_����ת��).Value = 1, byt��Դ) Then Exit Sub
                        End If
                        
                        
                        '�ж��Ƿ�Ƥ��,����д���
                        strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Item(col_ҽ��ID).Value)
                        If Not rsTmp.EOF Then
                            '�Ѿ���д��Ƥ�Խ��������д
                            If rsTmp!������� = "E" And Nvl(rsTmp!��������) = "1" And Not IsNull(rsTmp!Ƥ�Խ��) Then
                                '���������֤
                                If mblnƤ����֤ Then
                                    Set cnNew = New ADODB.Connection
                                    strUserName = zlDatabase.UserIdentify(Me, "��ȡ�����Ƥ��ҽ��ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "Ƥ��ҽ�����", cnNew)
                                    If strUserName = "" Then Exit Sub
                                End If
                                strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngҽ��ID & "," & lng���ͺ� & "," & IIf(mblnƤ����֤, 1, 0) & "," & IIf(blnGroupRow, 1, 0) & "," & mlngDept & ")"
                            Else
                                strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngҽ��ID & "," & lng���ͺ� & ",Null ," & IIf(blnGroupRow, 1, 0) & "," & mlngDept & ")"
                            End If
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = strSQL
                        End If
                    End If
                End With
            End If
        End If
    Next
    
    If strMsg <> "" Then
        strMsg = "���µ��ݺŵĲ��˷���������˽׶Σ����������ҽ���ͷ��ã�" & vbCrLf & Mid(strMsg, 2) & "��"
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    
    If UBound(arrSQL) = -1 Then Exit Sub
 
    If Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        cnNew.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call SQLTest(App.ProductName, Me.Caption, arrSQL(i))
            cnNew.Execute strOwner & "." & arrSQL(i), , adCmdStoredProc
            Call SQLTest
        Next
        
        cnNew.CommitTrans: blnTrans = False
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
     
    Call LoadPatients
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
errNew:
    cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function Check���ʷ���(ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng��ID As Long, ByVal str������� As String, ByVal int�������� As Integer, ByVal strNO As String) As Boolean
'���ܣ���ȡĳ��ҽ������ĳ��ҽ�����Ƿ�����Ѿ����ʵķ���
'       bln����ִ�� �Ƿ񵥶�ִ�У�����������ڵ��ݵ�ҽ���ĵ���ִ��ĳһ��λ��ĳһ���ּ��
'       lngҽ��ID ����ҽ��ID
'       lng��ID û�и�ҽ�������߸�ҽ��ʱΪҽ��ID,��ҽ��Ϊ���ID
'       str������� ��ҽ�����������
'       int�������� 1-������ã�2-סԺ����'
    Dim rsTmp As ADODB.Recordset, strSQL As String, strTable As String
    strTable = IIf(int�������� = 1, "������ü�¼", "סԺ���ü�¼")
    On Error GoTo errH
    If bln����ִ�� Then
        lng��ID = lngҽ��ID
        strSQL = "Select -1 * Sum(Nvl(a.����, 1) * a.���� / b.����) As ���������" & vbNewLine & _
                "From " & strTable & " A, ����ҽ���Ƽ� B" & vbNewLine & _
                "Where a.ҽ����� = [1] And A.NO=[3] And b.ҽ��id = a.ҽ����� And b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0 And a.��¼״̬ = 2 And a.��¼���� in(1,2,11) And a.�۸񸸺� Is Null And" & vbNewLine & _
                "      a.�շ���� Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
                " (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1)"

    Else
        strSQL = "Select Max(c.������) ���������" & vbNewLine & _
                "From (Select -1 * Sum(Nvl(a.����, 1) * a.���� / b.����) As ������" & vbNewLine & _
                "       From " & strTable & " A, ����ҽ���Ƽ� B" & vbNewLine & _
                "       Where a.ҽ����� In (Select ID From ����ҽ����¼ Where (ID = [1] Or ���id = [1]) And A.NO=[3] And ������� = [2]) And b.ҽ��id = a.ҽ����� And" & vbNewLine & _
                "             b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0 And a.��¼״̬ = 2 And a.��¼���� in(1,2) And a.�۸񸸺� Is Null And a.�շ���� Not In ('5', '6', '7') And" & vbNewLine & _
                "             Not Exists" & vbNewLine & _
                "        (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) " & vbNewLine & _
                "       Group By  a.ҽ�����,a.�շ�ϸĿid) C"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, str�������, strNO)
    If rsTmp.RecordCount <> 0 Then
        Check���ʷ��� = (Val(rsTmp!��������� & "") > 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Set�շѱ��(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lng���ͺ� As Long, _
    ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal blnMove As Boolean, ByVal dat����ʱ�� As Date) As Boolean
'���ܣ��жϽ����ϵ�"��"���Ƿ���ʾ�����÷��ù��������ӿڽ����ж�
    Dim strSQL As String, strTab As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strNos As String
    Dim bln���� As Boolean
    Dim bytState As Byte
    
    On Error GoTo errH
    
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
        bln���� = True
    End If
    
    strSQL = "select a.no from (" & _
        " Select a.no" & _
        " From " & strTab & " A,����ҽ����¼ B" & _
        " Where A.NO=[4] And A.ҽ�����+0=B.ID And MOD(A.��¼����,10)=[5]" & IIf(bln����ִ��, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select A.NO" & _
        " From ����ҽ����¼ C," & strTab & " B,����ҽ������ A" & _
        " Where A.NO=B.NO And A.��¼����=MOD(B.��¼����,10) And A.ҽ��ID=B.ҽ�����+0" & IIf(bln����ִ��, " And A.ҽ��ID=[2]", _
        " And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID=[1] Or ���ID=[1]) And �������=[6])") & _
        " And A.���ͺ�=[3] And A.ҽ��ID=C.ID And A.��¼����=[5]) a group by a.no"
    If blnMove Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, strTab, "H" & strTab)
    ElseIf zlDatabase.DateMoved(dat����ʱ��) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng���ID <> 0, lng���ID, lngҽ��ID), lngҽ��ID, lng���ͺ�, str���ݺ�, int��¼����, str���)
    
    For i = 1 To rsTmp.RecordCount
        strNos = strNos & "," & rsTmp!NO
        rsTmp.MoveNext
    Next
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
        If int��¼���� = 2 Then
            Call mclsPExp.zlGetBalanceStatus(strNos, bytState, bln����)
        Else
            Call mclsPExp.zlGetBillChargeStatus(strNos, bytState)
        End If
    End If
    Set�շѱ�� = (bytState = 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picMsg_Resize()
'
    Dim lngTmp As Long
   
    On Error Resume Next
    
    lngTmp = picMsg.Height
    
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    
    rptNotify.Top = 0
    rptNotify.Left = 0
    rptNotify.Width = picMsg.Width
    rptNotify.Height = lngTmp
End Sub

Private Sub FuncAppendBill()
'���ܣ����������Ѵ���

    Dim rsTmp As ADODB.Recordset
    Dim lngҽ��ID As Long
    Dim strSQL As String
    Dim lngTmp As Long
    
    Dim strPar As String
    Dim str��Դϵͳ As String
    Dim str������Դ As String
    Dim str���˱�ʶ As String
    Dim str�����ʶ As String
    Dim strҽ����� As String
    Dim strҽ�����ͺ� As String
    Dim str��ǰ���ұ�ʶ As String
    Dim str��ǰ���ұ��� As String
    Dim str��ǰ�������� As String
    Dim str����Ա��ʶ As String
    Dim str����Ա���� As String
    Dim str����Ա���� As String
    Dim strԺ������ As String 'վ��
    Dim strԺ������ As String
    Dim str�û��� As String
    Dim str�û����� As String
    
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
    
        lngҽ��ID = .Record(col_ҽ��ID).Value
        
        'Ժ������--ZLHISϵͳ�е�վ����Ϣ
        If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
            strSQL = "Select ���,���� From Zlnodelist Where ���=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
            If Not rsTmp.EOF Then
                strԺ������ = rsTmp!��� & ""
                strԺ������ = rsTmp!���� & ""
            End If
        End If
        
        '���˱�ʶ-- ����ID
        str���˱�ʶ = mlng����ID
        
        '�����ʶ--���ﲡ�� ����,סԺ���� ��ҳID
        If .Record(col_��Դ).Value = "����" Then
            str�����ʶ = ""
        Else
            str�����ʶ = .Record(col_��ҳID).Value
        End If
        
        '��Դϵͳ-- 01 ZLHIS�еĲ��ˣ�02 �²���
        str��Դϵͳ = "01"
        lngTmp = Val(.Record(col_�Һ�ID).Value)
        If lngTmp <> 0 Then
            strSQL = "Select 1 From ���˹Һż�¼ a where a.���ӱ�־=3 and a.id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngTmp)
            If Not rsTmp.EOF Then '�ҵ�����˵���������ﲡ��
                str��Դϵͳ = "02"
                str�����ʶ = ""
            End If
        End If
        
        '������Դ -- 0-����/1-סԺ/2-���
        str������Դ = Decode(.Record(col_��Դ).Value, "����", 0, "סԺ", 1, "���", 2, "����", 3)
 
        '���ͺ�
        strҽ�����ͺ� = .Record(col_���ͺ�).Value
        
        strҽ����� = lngҽ��ID
        
        strSQL = "Select id,��Դid,���� As ��ǰ���ұ���,���� as ��ǰ�������� From ���ű� Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept)
        
        If str��Դϵͳ = "02" Then
            str��ǰ���ұ�ʶ = rsTmp!��Դid & ""
	    str����Ա��ʶ = Sys.RowValue("��Ա��", UserInfo.ID, "��Դid")
        Else
            str��ǰ���ұ�ʶ = rsTmp!ID & ""
	    str����Ա��ʶ = UserInfo.ID
        End If
        str��ǰ���ұ��� = rsTmp!��ǰ���ұ��� & ""
        str��ǰ�������� = rsTmp!��ǰ�������� & ""
          
        str����Ա���� = UserInfo.���
        str����Ա���� = UserInfo.����
    End With
    
    str�û��� = UserInfo.�û��� ' "ZLHIS"
    
    If mstr���� = "" Then
        mstr���� = GetConnPassword
    End If
    str�û����� = mstr����
    
    strPar = _
        "{" & _
            """��Դϵͳ"":""" & str��Դϵͳ & """," & _
            """������Դ"":" & str������Դ & "," & _
            """���˱�ʶ"":""" & str���˱�ʶ & """," & _
            """�����ʶ"":""" & str�����ʶ & """," & _
            """ҽ�����"":""" & strҽ����� & """," & _
            """ҽ�����ͺ�"":""" & strҽ�����ͺ� & """," & _
            """��ǰ���ұ�ʶ"":""" & str��ǰ���ұ�ʶ & """," & _
            """��ǰ���ұ���"":""" & str��ǰ���ұ��� & """," & _
            """��ǰ��������"":""" & str��ǰ�������� & """," & _
            """����Ա��ʶ"":""" & str����Ա��ʶ & """," & _
            """����Ա����"":""" & str����Ա���� & """," & _
            """����Ա����"":""" & str����Ա���� & """," & _
            """Ժ������"":""" & strԺ������ & """," & _
            """Ժ������"":""" & strԺ������ & """," & _
            """�û���"":""" & str�û��� & """," & _
            """�û�����"":""" & str�û����� & """" & _
        "}"
    '���ò���
    Call mobjAppendBill.EditChargeBill(strPar)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetConnPassword()
    '��ȡ��ǰ�û���¼����
    Dim objLogin As Object
    
    On Error Resume Next
    Set objLogin = CreateObject("zlLogin.clsLogin")
    If objLogin Is Nothing Then
        err.Clear
        MsgBox "����zlLogin��������ʧ�ܣ������ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Exit Function
    End If
    
    GetConnPassword = objLogin.InputPwd
End Function

Private Function NewOut�շ�(ByVal lngҽ��ID As Long) As Long
'���ܣ��ж������ﲡ���Ƿ��շѣ�����������ϵͳ�ķ���
'���أ�0-����ǰ�߼���1-��ʾ����ֹ��2-�ɹ���֤ͨ��
    Dim strJsIn As String
    Dim strJsOut As String
    Dim strErr As String
    Dim int���շ� As Integer
    Dim blnTmp As Boolean
    Dim lngRes As Long
    
    Screen.MousePointer = 11
    strJsIn = "{""input"":{""head"":{""bizno"":""RJ001"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""},""apply_id"":" & lngҽ��ID & "}}"
    blnTmp = Sys.NewSystemSvr("������ϵͳ", "�ж�ҽ���Ƿ��շ�", strJsIn, strJsOut, strErr)
    Screen.MousePointer = 0
    If strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        lngRes = 1
        NewOut�շ� = lngRes
        Exit Function
    End If
    
    If blnTmp Then
        If strJsOut <> "" Then
            If Val(zlStr.JSONParse("result", strJsOut) & "") <> 1 Then
                MsgBox zlStr.JSONParse("errmsg", strJsOut) & "", vbInformation, gstrSysName
                lngRes = 1
            End If
            int���շ� = Val(zlStr.JSONParse("kacnt_sign", strJsOut) & "")
        End If
        If int���շ� <> 1 Then
            MsgBox "�ò��˻�����δ�շѵķ��ã����顣", vbInformation, gstrSysName
            lngRes = 1
        Else
            lngRes = 2
        End If
    End If
    NewOut�շ� = lngRes
End Function