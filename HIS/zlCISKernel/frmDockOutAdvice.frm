VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDockOutAdvice 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timHide 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7005
      Top             =   690
   End
   Begin VB.PictureBox PicAdviceDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEFEF&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   135
      ScaleHeight     =   2745
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   2800
      Begin VSFlex8Ctl.VSFlexGrid vsfAdivceDetail 
         Height          =   2475
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2745
         _cx             =   4851
         _cy             =   4366
         Appearance      =   2
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
         BackColor       =   16773103
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16773103
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16773103
         BackColorAlternate=   16773103
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockOutAdvice.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaper       =   "frmDockOutAdvice.frx":003E
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   120
      ScaleHeight     =   4890
      ScaleWidth      =   6570
      TabIndex        =   0
      Top             =   210
      Width           =   6570
      Begin VB.Frame fraMore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5250
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   225
         Begin VB.Image imgMore 
            Height          =   225
            Left            =   0
            Picture         =   "frmDockOutAdvice.frx":12970
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5460
         TabIndex        =   3
         Top             =   255
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockOutAdvice.frx":12D71
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.Frame fraAdviceUD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   4
         Top             =   3720
         Width           =   6975
      End
      Begin VB.Frame fraHide 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   6150
         TabIndex        =   2
         ToolTipText     =   "���ͣ��ʱ,�������������Զ���ʾ"
         Top             =   135
         Visible         =   0   'False
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3630
         Left            =   0
         TabIndex        =   5
         Top             =   45
         Width           =   5265
         _cx             =   9287
         _cy             =   6403
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
         MouseIcon       =   "frmDockOutAdvice.frx":132BF
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
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockOutAdvice.frx":14C51
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
         Begin VB.PictureBox pictmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   6
            Top             =   1320
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin XtremeSuiteControls.TabControl tbcAppend 
         Height          =   1500
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   3285
         _Version        =   589884
         _ExtentX        =   5794
         _ExtentY        =   1482
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   2940
         Left            =   5400
         TabIndex        =   8
         Top             =   675
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   5186
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockOutAdvice.frx":14CEC
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
      Begin XtremeCommandBars.CommandBars cbsSub 
         Left            =   5640
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Bindings        =   "frmDockOutAdvice.frx":14D3A
      Height          =   435
      Left            =   7050
      TabIndex        =   9
      Top             =   60
      Width           =   390
      _Version        =   589884
      _ExtentX        =   688
      _ExtentY        =   767
      _StockProps     =   64
   End
   Begin RichTextLib.RichTextBox rtfInfo 
      Height          =   900
      Left            =   3495
      TabIndex        =   12
      Top             =   5580
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14D4E
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
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   900
      Left            =   3150
      TabIndex        =   13
      Top             =   5580
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14DEB
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
   Begin RichTextLib.RichTextBox rtfSche 
      Height          =   900
      Left            =   3840
      TabIndex        =   14
      Top             =   5580
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14E88
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
   Begin RichTextLib.RichTextBox rtfOther 
      Height          =   900
      Left            =   4185
      TabIndex        =   15
      Top             =   5580
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14F25
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
   Begin VSFlex8Ctl.VSFlexGrid vsAppend 
      Height          =   1155
      Left            =   5130
      TabIndex        =   16
      Top             =   5865
      Width           =   1350
      _cx             =   2381
      _cy             =   2037
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
   Begin MSComctlLib.ImageList img16 
      Left            =   7050
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":14FC2
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":1555C
            Key             =   ""
            Object.Tag             =   "99"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":15AF6
            Key             =   ""
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":16090
            Key             =   ""
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":1662A
            Key             =   ""
            Object.Tag             =   "90003"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":169C4
            Key             =   ""
            Object.Tag             =   "90004"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDockOutAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '���Ѽ���ʱ
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean) 'Ҫ���ӡ����
Public Event ViewPACSImage(ByVal ҽ��ID As Long) 'Ҫ����й�Ƭ
Public Event EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean) '�༭�������
Public Event CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���Id As String, ByRef blnYes As Boolean) '������ϼ���Ƿ���д��Ⱦ�����濨
Public Event VSKeyPress(KeyAscii As Integer)
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrBillPrint As String '��ǰ��ӡ�����Ƶ��ݣ������š�NO����¼����
Private mobjPublicPACS As Object             'PACSҵ���װ��������

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mblnEditable As Boolean
Private mblnCanRevoke As Boolean '�Ƿ��������ҽ��
Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mstr���� As String
Private mstr����� As String
Private mlng�Һ�ID As Long
Private mlngǰ��ID As Long
Private mlng�������ID As Long
Private mlng�Һſ���ID As Long
Private mstrǰ��IDs As String
Private mstr����ҽ�� As String
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�

Private mvRegDate As Date '�Һ�ʱ��,3000-01-01��ʾδ�ҺŵĲ���
Private mblnMoved As Boolean
Private mbln���� As Boolean
Private mbln���� As Boolean
Private mblnָ������ӡ As Boolean
Private mint���� As Integer
Private mblnModalNew As Boolean '�¿������Ƿ�ģ̬

Private mint���� As Integer '���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Private mlng·��״̬ As Long    '-1-δ���룬0-�����ϵ���������1-ִ���У�2-����������3-�������
Private mint�������� As Integer 'pt���� = 0��pt���� = 1��pt���� = 2��ptת�� = 3��ptԤԼ = 4��pt���� = 5
Private mblnNotEvaluete As Boolean  'δ����ʱ�������ҽ��������

Private WithEvents mfrmSend As frmOutAdviceSend
Attribute mfrmSend.VB_VarHelpID = -1
Private WithEvents mfrmEdit As frmOutAdviceEdit
Attribute mfrmEdit.VB_VarHelpID = -1
Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private mcbsMain As Object
Private mMainPrivs As String
Private mblnAppend As Boolean
Private mblnƤ������ As Boolean
Private mblnAutoRead As Boolean
Private mblnAutoReadEnabled As Boolean
Private mrsDefine As ADODB.Recordset    'ҽ�����ݶ���
Private mobjVBA As Object
Private mobjScript As clsScript

Private mlngFontSize As Long  '�����С

Private mblnFirst As Boolean '�Ƿ��״ε���
Private mlngPlugInID As Long '�Զ�ִ�еĲ������ID
Private mrsPlugInBar As ADODB.Recordset '�˵���ʽ
Private mlngPromptRow As Long    '��һ�Σ�������ƶ�ͼ������ʾ����ʾ��Ϣ����
Private mSendControl As CommandBarControl     '���Ͱ�ť
Private mblnSignVisible As Boolean  'ǩ�����ܰ�ť�ɼ���
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mstr�Զ������뵥IDs As String 'ID1,����1|ID2,����2������
Private mrsΣ��ֵ As ADODB.Recordset
Private mblnΣ��ֵ As Boolean '�Ƿ��д���Σ��ֵ��Ȩ��
Private mlngΣ��ֵID As Long '��ǰ�����Σ��ֵ��¼ID
'Pass
Private mobjPassMap As Object  'PASS �������ӳ��
Private mblnPass As Boolean  'PASSȨ��
'�鿴����
Private mblnTag As Boolean  '�Ƿ��ѵ���鿴�ж�
Private mbln����Ԥ�� As Boolean  '�Ƿ��ѵ��Ϊ����Ԥ������
Private mobjFrmBloodList As Object 'ѪҺ��ϸ����
'����ҽ����������
Private Enum CMD_FILTER
    ID_Ӥ�� = 1
    ID_��ֹ = 2
    ID_���� = 5
    ID_��� = 7
    ID_���� = 8
    ID_ȫ�� = 9
    ID_��� = 10
    ID_���� = 11
    ID_���� = 12
    ID_ҽ��ȫ�� = 13
    ID_ҽ������ = 14
    ID_ҽ������ = 15
    ID_δ������ = 16
    ID_�ѳ����� = 17
End Enum

Private Type FilterCond
    Ӥ�� As Integer
    ��ֹ As Boolean     'true ��ʾ����ҽ����false ����ʾ����ҽ��
    ���� As Boolean
    ���� As Integer     '0-ȫ����1����飬2�����飬3������
    ��ʾģʽ As Integer '0-��࣬1-����
    ����ģʽ As Integer '0-ҽ����3������
    ҽ�� As Integer '0-ȫ����1-������2-����
    δ������ As Boolean
    �ѳ����� As Boolean
End Type

Private mvarCond As FilterCond
Private mblnHideFilter As Boolean

Private Enum COLҽ���嵥
    '�̶���
    COL_F��־ = 0
    COL_F���� = 1
    '������
    COL_ID = 2
    COL_���ID = COL_ID + 1
    COL_Ӥ��ID = COL_ID + 2
    COL_ҽ��״̬ = COL_ID + 3
    COL_������� = COL_ID + 4
    COL_�������� = COL_ID + 5
    COL_������� = COL_ID + 6
    COL_��־ = COL_ID + 7
    
    '�ɼ���
    COL_��ʾ = COL_ID + 8 'Pass
    COL_������ = COL_ID + 9
    COL_������ӡ = COL_ID + 10
    COL_����Ԥ�� = COL_ID + 11
    COL_��ʼʱ�� = COL_ID + 12
    COL_�� = COL_ID + 13
    col_ҽ������ = COL_ID + 14
    col_���� = COL_ID + 15
    COL_Ƥ�� = COL_ID + 16
    COL_���� = COL_ID + 17
    COL_���� = COL_ID + 18
    COL_���� = COL_ID + 19
    COL_Ƶ�� = COL_ID + 20
    COL_�÷� = COL_ID + 21
    COL_ҽ������ = COL_ID + 22
    COL_ִ��ʱ�� = COL_ID + 23
    COL_ִ�п��� = COL_ID + 24
    COL_ִ������ = COL_ID + 25
    COL_����ҽ�� = COL_ID + 26
    COL_����ʱ�� = COL_ID + 27
    COL_������ = COL_ID + 28
    col_����ʱ�� = COL_ID + 29
    COL_����˵�� = COL_ID + 30
    COL_����ҩ�� = COL_ID + 31
    COL_����״̬ = COL_ID + 32
    COL_�걾״̬ = COL_ID + 33
    
    '������
    COL_������ĿID = COL_ID + 34
    COL_�Թܱ��� = COL_������ĿID + 1
    COL_ǰ��ID = COL_������ĿID + 2
    COL_ǩ���� = COL_������ĿID + 3
    COL_�ļ�ID = COL_������ĿID + 4
    COL_������ = COL_������ĿID + 5 '0-�ޱ��棬1-�б��沢���༭��ʽ��ӡ��2-�б��沢�������ʽ��ӡ��
    COL_����ID = COL_������ĿID + 6
    COL_���״̬ = COL_������ĿID + 7
    COL_������� = COL_������ĿID + 8
    COL_��ΣҩƷ = COL_������ĿID + 9
    COL_�걾��λ = COL_������ĿID + 10
    COL_�շ�ϸĿID = COL_������ĿID + 11   'Pass
    COL_��������ID = COL_������ĿID + 12
    COL_��ҩĿ�� = COL_������ĿID + 13
    COL_��鱨��ID = COL_������ĿID + 14
    COL_�������״̬ = COL_������ĿID + 15
    COL_��������� = COL_������ĿID + 16
    COL_RISԤԼID = COL_������ĿID + 17
    COL_RIS����ID = COL_������ĿID + 18
    COL_LIS����ID = COL_������ĿID + 19
    COL_RISԤԼ״̬ = COL_������ĿID + 20
    col_������Ŀ���� = COL_������ĿID + 21
    COL_��鷽�� = COL_������ĿID + 22  '��Ѫҽ�������Ǳ�Ѫ������Ѫ
    COL_Σ��ֵID = COL_������ĿID + 23 'ҽ���غ�Σ��ֵ����
    COL_�׵��� = COL_������ĿID + 24 'ҩƷ���׵���
End Enum

Private COLPrice As New Collection
Private COLSend As New Collection
Private COLSign As New Collection

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByVal int���� As Integer, _
                            ByRef objPlugIn As Object, ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
    
    mint���� = int����
    Set mfrmParent = frmParent
        mblnModalNew = blnModalNew
    If Not cbsMain Is Nothing Then

        '��ҳ�������ʼ��
        If Not mblnFirst Then
            mblnFirst = True
            Set mcbsMain = cbsMain
            Set cbsMain.Icons = zlCommFun.GetPubIcons
            Set gobjSquareCard = objSquareCard

            If gobjPlugIn Is Nothing Then
                If Not objPlugIn Is Nothing Then
                    '��ҽ��վ����ʱ�����ⲿ��ʼ�����˴������ٵ���ʼ���ӿ�
                    Set gobjPlugIn = objPlugIn
                Else
                    Call CreatePlugInOK(p����ҽ���´�, mint����)
                End If
            End If
            Call GetPlugInBar(p����ҽ���´�, mint����, mrsPlugInBar)

            'PASS�ӿڳ�ʼ��
            '��Ϊ����ģ�����ͬʱʹ��,�п���gobjPass�����Ѿ�����
            If gobjPass Is Nothing Then
                Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "������ҩ���", True)
                If Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassInit(gcnOracle, glngSys, PM_����ҽ���嵥)
                    If gobjPass.PassType = 0 Then
                        Set gobjPass = Nothing
                    Else
                        mblnPass = True
                    End If
                End If
            End If
        End If
        
        Call zlPASSMap
        If mblnPass Then
            Call gobjPass.zlPassAdviceColHidden(mobjPassMap)
        End If
        
        If mint���� = 0 Then    'ҽ��վ����
            Call DefCommandsOutDoctor(cbsMain)
        ElseIf mint���� = 2 Then    'ҽ��վ����
            Call DefCommandsTechnic(cbsMain)
        End If

        Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
    End If
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object, ByRef rsBar As ADODB.Recordset)
'���ܣ���Ҳ����˵����롣
'˵�����жϹؼ���  Auto  InTool �����˵���ʽ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '������ť
    rsBar.Filter = "IsInTool=1 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                        objControl.IconId = rsBar!ͼ��ID
                        objControl.Parameter = rsBar!������
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '������ť�����ֻ��һ����ť��Ҳ����������ť
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        If Not objMenu Is Nothing Then
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����", , False)
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '��������ť
    Set objBar = cbsMain(2)
    Set objControl = objBar.FindControl(, conMenu_Help_Help)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������, lngTmp + 1)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���, lngTmp + 1)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    '�Զ�ִ�еĹ���
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!����ID
End Sub

Private Sub DefCommandsTechnic(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim intTmp As Integer
    Dim strTmp As String
    Dim strName As String
    Dim lngID As Long
    Dim varArr As Variant
    Dim i As Long
    
    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "ҽ���༭(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)"
        End With
        
        intTmp = Val(Mid(gstrOutUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrOutUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrOutUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrOutUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_OperationApply
                Get�Զ������뵥 1, mstr�Զ������뵥IDs
        If mstr�Զ������뵥IDs <> "" Then
            For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
            Next
        End If
        strTmp = Mid(strTmp, 2)
        
        If strTmp <> "" Then
            If InStr(strTmp, ",") = 0 Then
                strName = Split(strTmp, ":")(0)
                lngID = Val(Split(strTmp, ":")(1))
                Set objControl = .Add(xtpControlButton, lngID, strName)
                    objControl.IconId = conMenu_Manage_Request
                    objControl.ToolTipText = strName
                    objControl.BeginGroup = True
                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
            Else
                varArr = Split(strTmp, ",")
                For i = 0 To UBound(varArr)
                    strTmp = varArr(i)
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    
                    If i = 0 Then
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "�´�����"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "�޸�����")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "ȡ������")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "���ԤԼ")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "ԤԼ(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "����ԤԼ(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "ȡ��ԤԼ(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "�ؼ�ͼ��")
       '2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "��������(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "������ͼ��ͱ���(&Y)")
                objControl.IconId = 237
        End If
        '2017-11-10 ������
        If gblnѪ��ϵͳ Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_BloodInstant, "��Ѫִ�е�")
            objControl.BeginGroup = True
        End If
    End With
    If Not objMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            If mblnָ������ӡ Then
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicIndexBill, "��ӡָ����")
                    objControl.IconId = 103
            End If
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ����")
            objPopup.BeginGroup = True
            Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
        End With
    End If
    
    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    End If
    
    '�鿴�˵�
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '״̬�����
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
    End With
    
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "��ӡԤԼ��")
                objControl.IconId = 103
        End If
        If gbln����ҩ�����հ������������� Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ��ѡ��(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&S)"): objControl.BeginGroup = True
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Call AddToolBarInDoctor
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send 'ҽ������
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '���ı���
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '������ͼ��ͱ���
        
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '���������
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '��������
        .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
    End With

    '���ò���������
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Private Sub DefCommandsOutDoctor(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl, lngIdx As Long
    
    Dim varArr As Variant
    Dim strTmp As String
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    Dim i As Long
    

    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "ҽ���༭(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)"
        End With
     
        intTmp = Val(Mid(gstrOutUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrOutUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrOutUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrOutUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_OperationApply
        Get�Զ������뵥 1, mstr�Զ������뵥IDs
        If mstr�Զ������뵥IDs <> "" Then
            For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
            Next
        End If
        strTmp = Mid(strTmp, 2)
        
        If strTmp <> "" Then
            If InStr(strTmp, ",") = 0 Then
                strName = Split(strTmp, ":")(0)
                lngID = Val(Split(strTmp, ":")(1))
                Set objControl = .Add(xtpControlButton, lngID, strName)
                    objControl.IconId = conMenu_Manage_Request
                    objControl.ToolTipText = strName
                    objControl.BeginGroup = True
                                        If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
            Else
                varArr = Split(strTmp, ",")
                For i = 0 To UBound(varArr)
                    strTmp = varArr(i)
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    
                    If i = 0 Then
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "�´�����"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "�޸�����")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "ȡ������")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "���ԤԼ")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "ԤԼ(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "����ԤԼ(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "ȡ��ԤԼ(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
                
        If gblnѪ��ϵͳ Then Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReaction, "��Ѫ��Ӧ"): objControl.IconId = 4113
        
        If mblnΣ��ֵ Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_CriticalAdvice, "Σ��ֵҽ��")
        End If

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Test, "Ƥ�Խ��(&T)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePay, "���֧��")
            objControl.IconId = conMenu_Edit_Pay
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "�ؼ�ͼ��")
       '2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "��������(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "������ͼ��ͱ���(&Y)")
                objControl.IconId = 237
        End If
            
        Set objControl = .Add(xtpControlButton, conMenu_Manage_RecipeAuditView, "�鿴���������")
        objControl.IconId = 3205
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewDrugExplain, "�鿴ҩƷ˵����")
        objControl.IconId = 3205
        If gbln��ϵͳ Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Refcom, "�ܾ��������")
                objControl.IconId = 3205
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewRefcom, "�������δͨ����Ϣ")
                objControl.IconId = 3205
        End If
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
    End With
    With objMenu.CommandBar.Controls
        If mblnָ������ӡ Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicIndexBill, "��ӡָ����")
                objControl.IconId = 103
        End If
        '���������ǰ��,�������
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ����")
        objPopup.BeginGroup = True
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With
    
    
    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    End If

    '�鿴�˵�
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '״̬�����
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
    End With
    
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "��ӡԤԼ��")
                objControl.IconId = 103
        End If
        If gbln����ҩ�����հ������������� Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ��ѡ��(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&S)"): objControl.BeginGroup = True
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Call AddToolBarInDoctor
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send 'ҽ������
        .Add FCONTROL, vbKeyT, conMenu_Edit_Test 'Ƥ�Խ��
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '���ı���
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '������ͼ��ͱ���
        
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '���������
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '��������
        .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
    End With

    '���ò���������
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim lngҽ��ID As Long
    Dim rsTmp As ADODB.Recordset
        
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_CriticalAdvice
        If mblnΣ��ֵ And Not mrsΣ��ֵ Is Nothing Then
            mrsΣ��ֵ.Filter = 0
            If Not mrsΣ��ֵ.EOF Then
                Set rsTmp = GetCriticalAdvice(lngҽ��ID)
                With CommandBar.Controls
                    .DeleteAll
                    mrsΣ��ֵ.MoveFirst
                    For i = 1 To mrsΣ��ֵ.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_CriticalAdvice * 100# + i, mrsΣ��ֵ!Σ��ֵ���� & "")
                            objControl.Parameter = mrsΣ��ֵ!ID & "," & lngҽ��ID
                        rsTmp.Filter = "Σ��ֵID=" & mrsΣ��ֵ!ID
                        If Not rsTmp.EOF Then
                            objControl.Checked = True
                        End If
                        mrsΣ��ֵ.MoveNext
                    Next
                    mrsΣ��ֵ.MoveFirst
                End With
            End If
            mrsΣ��ֵ.Filter = 0
        End If
    Case conMenu_Edit_Compend '����
        With CommandBar.Controls
            If .Count = 0 Then
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "���ı���(������ʽ)"
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 6, "���ı���(�����ʽ)"
                If gobjExchange Is Nothing Then
                    If mint���� = 1 Then    '��ʿվ
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)"
                    Else
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "��ӡ����(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)"

                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "���Ѳ���(&R)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "�Զ����(&A)"
                    End If
                End If
            End If
        End With
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99
        'PASSҩ�����
        If mblnPass Then
            Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, CommandBar, conMenu_Edit_MediAudit)
        End If
    End Select
End Sub

Private Sub AddToolBarInDoctor()
'���ܣ����ù�������ť����Ӧ��ҽ���˵�����Ĺ������İ�ť���Ƚ���ɾ�������
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngTmp As Long
    Dim objCbs As Object
    Dim lngIdx As Long
    Dim i As Long
    
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    
    Dim blnTwo As Boolean, strInsidePrivs As String
    
    On Error GoTo errH
    
    If mcbsMain Is Nothing Then Exit Sub

    strInsidePrivs = GetInsidePrivs(p����ҽ��վ)
    blnTwo = Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) <> 2
    
    strTmp = "," & conMenu_Edit_NewItem & "," & conMenu_Edit_Apply & "," & conMenu_Edit_ApplyModi & "," & conMenu_Edit_ApplyView & "," & conMenu_Edit_ApplyDel & "," & _
        conMenu_Edit_Blankoff & "," & conMenu_Edit_TraReaction & "," & conMenu_Edit_SendBilling & "," & conMenu_Edit_Send & "," & IIF(blnTwo, conMenu_Edit_Send * 100# + 1 & ",", "") & conMenu_Edit_Untread & "," & _
        conMenu_Edit_Compend & "," & (conMenu_Edit_Compend * 10# + 2) & "," & (conMenu_Edit_Compend * 10# + 3) & "," & conMenu_Edit_MarkMap & "," & conMenu_Edit_MarkKeyMap & "," & conMenu_Edit_MarkKeyMap & "," & conMenu_Manage_ReportLisView & "," & _
        conMenu_Edit_MediAudit & "," & conMenu_Tool_SignNew & "," & conMenu_Edit_Audit & "," & conMenu_Edit_Price & "," & conMenu_Report_ClinicBill & "," & conMenu_Edit_PacsApply & "," & conMenu_Edit_BloodApply & ","
    strTmp = strTmp & "," & conMenu_Edit_PacsApply & "," & (conMenu_Edit_PacsApply * 10# + 1) & "," & conMenu_Edit_LISApply & "," & (conMenu_Edit_LISApply * 10# + 1) & "," & conMenu_Edit_BloodApply & "," & (conMenu_Edit_BloodApply * 10# + 1)
    strTmp = strTmp & "," & conMenu_Edit_OperationApply & "," & (conMenu_Edit_OperationApply * 10# + 1) & "," & conMenu_Edit_ConsultationApply & "," & (conMenu_Edit_ConsultationApply * 10 + 1) & "," & conMenu_Edit_AdvicePay & ","
    
    '���������
    Set objCbs = mcbsMain
    '�ҵ�Ҫ��ӵ�λ��
    lngIdx = 0
    For Each objControl In objCbs(2).Controls '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objCbs(2).Controls(objControl.Index - 1)
            lngIdx = objControl.Index
            Exit For
        End If
    Next
    
    'ɾ����������ť
    For i = objCbs(2).Controls.Count To 1 Step -1
        If InStr(strTmp, "," & objCbs(2).Controls(i).ID & ",") > 0 Then
            objCbs(2).Controls(i).Delete
        End If
    Next i

    With objCbs(2).Controls
        If mvarCond.����ģʽ <> 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "�¿�", lngIdx + 1): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "�¿�")
                    objControl.IconId = conMenu_Edit_NewItem
                .Add xtpControlButton, conMenu_Edit_Modify, "�޸�"
                .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��"
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
        End If
        
        If mint���� = 0 Then 'ֻ������ҽ������վ����ʱ�����⼸����ť
            strTmp = ""
            intTmp = Val(Mid(gstrOutUseApp, 1, 1))
            If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
            intTmp = Val(Mid(gstrOutUseApp, 2, 1))
            If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
            intTmp = Val(Mid(gstrOutUseApp, 3, 1))
            If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
            intTmp = Val(Mid(gstrOutUseApp, 4, 1))
            If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_OperationApply
            Get�Զ������뵥 1, mstr�Զ������뵥IDs
            If mstr�Զ������뵥IDs <> "" Then
                For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                    strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
                Next
            End If
            strTmp = Mid(strTmp, 2)
            
            If strTmp <> "" Then
                If InStr(strTmp, ",") = 0 Then
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    Set objControl = .Add(xtpControlButton, lngID, strName, lngIdx + 1)
                        objControl.IconId = conMenu_Manage_Request
                        objControl.ToolTipText = strName
                        objControl.Style = xtpButtonIconAndCaption
                        objControl.BeginGroup = True
                                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    lngIdx = objControl.Index
                Else
                    varArr = Split(strTmp, ",")
                    For i = 0 To UBound(varArr)
                        strTmp = varArr(i)
                        strName = Split(strTmp, ":")(0)
                        lngID = Val(Split(strTmp, ":")(1))
                        
                        If i = 0 Then
                            Set objPopup = .Add(xtpControlSplitButtonPopup, lngID, strName, lngIdx + 1)
                                objPopup.IconId = conMenu_Manage_Request
                                objPopup.BeginGroup = True
                                objPopup.ToolTipText = strName
                                objPopup.Style = xtpButtonIconAndCaption
                                With objPopup.CommandBar.Controls
                                    Set objControl = .Add(xtpControlButton, lngID * 10# + 1, strName)
                                End With
                        Else
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                        End If
                        If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    Next
                    lngIdx = objPopup.Index
                End If
            End If
        End If
        
        If mvarCond.����ģʽ = 3 And mint���� = 0 Then 'ֻ��סԺҽ������վ����ʱ�����⼸����ť
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "�޸�", lngIdx + 1)
                objControl.IconId = 3002
                objControl.ToolTipText = "�޸�����"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴", objControl.Index + 1)
                objControl.IconId = 102
                objControl.ToolTipText = "�鿴����"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "ȡ��", objControl.Index + 1)
                objControl.IconId = 3004
                objControl.ToolTipText = "ȡ������"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        End If
        
        If blnTwo Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "����", lngIdx + 1)
            objPopup.BeginGroup = True
            objPopup.ToolTipText = "ҽ�����ʹ���"
            objPopup.Style = xtpButtonIconAndCaption
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Edit_Send * 100# + 1, "�Զ���ɷ���"
            End With
            lngIdx = objPopup.Index
        Else
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", lngIdx + 1)
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            objControl.ToolTipText = "ҽ�����ʹ���"
            lngIdx = objControl.Index
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "����", lngIdx + 1)
            objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        lngIdx = objControl.Index
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePay, "���֧��", lngIdx + 1)
            objControl.IconId = conMenu_Edit_Pay
            objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        lngIdx = objControl.Index
        If mvarCond.����ģʽ = 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "����", lngIdx + 1): objPopup.BeginGroup = True
                objPopup.IconId = conMenu_Manage_Report
                objPopup.ToolTipText = "���ı���"
                
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 1, "������ʽ(&B)"): objControl.IconId = 102
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 6, "�����ʽ(&P)"): objControl.IconId = 102
                If gobjExchange Is Nothing And mint���� <> 1 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "���Ѳ���(&R)")
                        objControl.BeginGroup = True
                    .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "�Զ����(&A)"
                End If
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
            If gobjExchange Is Nothing Then
                If mint���� <> 1 Then
                    Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend * 10# + 2, "��ӡ����", lngIdx + 1)
                        objPopup.IconId = 103
                        objPopup.Style = xtpButtonIconAndCaption
                        With objPopup.CommandBar.Controls
                            Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������"): objControl.IconId = 102
                            objControl.Style = xtpButtonIconAndCaption
                        End With
                    lngIdx = objPopup.Index
                Else
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)", lngIdx + 1)
                    objControl.IconId = 102
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
                End If
            End If
    
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "��Ƭ����"
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "�ؼ�ͼ��", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "�ؼ�ͼ��"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���", objControl.Index + 1): objControl.IconId = conMenu_Manage_ReportLisView
                objControl.ToolTipText = "���������"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        Else
            If mint���� = 0 Then
                If mblnPass Then
                    lngIdx = lngIdx + 1
                    Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objCbs(2).Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit, lngIdx)
                End If
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ǩ��", objControl.Index + 1): objControl.BeginGroup = True
                objControl.IconId = conMenu_Tool_Sign
                objControl.Style = xtpButtonIconAndCaption
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ����", objControl.Index + 1)
                objPopup.BeginGroup = True
                objPopup.IconId = conMenu_File_Print
                objPopup.Visible = False
        End If
    End With
    mcbsMain.RecalcLayout
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strErr As String

    Select Case Control.ID
    Case conMenu_File_PrintSet '��ӡ����
        SwitchPrintSet glngSys & "\" & p����ҽ���´�
        Call zlPrintSet
        SwitchPrintSet glngSys & "\" & p����ҽ���´�, True
    Case conMenu_File_Preview 'Ԥ��ҽ���嵥
        Call OutputList(2)
    Case conMenu_File_Print '��ӡҽ���嵥
        Call OutputList(1)
    Case conMenu_File_Excel '���ҽ���嵥
        Call OutputList(3)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_View_Append '��������
        mblnAppend = Not mblnAppend
        tbcAppend.Visible = Not tbcAppend.Visible
        fraAdviceUD.Visible = Not fraAdviceUD.Visible
        Call Form_Resize
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        If vsAppend.Visible And vsAppend.Enabled Then
            vsAppend.SetFocus
        Else
            If vsAdvice.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
        Call cbsSub_Resize
    Case conMenu_View_Hide '�Զ����ع��˹�����
        mblnHideFilter = Not mblnHideFilter
        cbsSub(2).Visible = Not mblnHideFilter
        fraHide.Visible = mblnHideFilter
        cbsSub.RecalcLayout
    Case conMenu_Edit_NewItem, conMenu_Edit_NewItem * 10# + 1 '�¿�ҽ��
        If Control.Parameter <> "" Then
            mlngΣ��ֵID = Val(Control.Parameter)
            Call GetCriticalData
        Else
            mlngΣ��ֵID = 0
        End If
        Call FuncAdviceAdd
    Case conMenu_Edit_Modify '�޸�ҽ��
        Call FuncAdviceModi
    Case conMenu_Edit_Delete, conMenu_Edit_ApplyDel 'ɾ��ҽ��'ȡ����������
        Call FuncAdviceDel
    Case conMenu_Edit_LISApply, conMenu_Edit_LISApply * 10 + 1   '��������
        Call FuncLISApply(0)
    Case conMenu_Edit_PacsApply, conMenu_Edit_PacsApply * 10# + 1 '�������
        Call FuncPacsApply(0, 0)
    Case conMenu_Edit_BloodApply, conMenu_Edit_BloodApply * 10 + 1  '��Ѫ����
        Call FuncBloodApply(0)
    Case conMenu_Edit_OperationApply, conMenu_Edit_OperationApply * 10 + 1 '��������
        Call FuncOperationApply(0)
    Case conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101#
        FuncApplyCustom 0, Control.Parameter
    Case conMenu_Edit_ApplyView '�鿴����
        Call FuncApplyView
    Case conMenu_Edit_ApplyModi '�޸�����
        Call FuncApplyModi
    Case conMenu_Edit_NewRisSch 'RISԤԼ
        Call FuncAdviceRISSch
    Case conMenu_Edit_NewRisDel 'ȡ��ԤԼ
        Call FuncAdviceRISDel
    Case conMenu_Edit_NewRisModi
        Call FuncAdviceRISModi
    Case conMenu_Tool_RisPrint
        Call FuncAdviceRISPrintSch
    Case conMenu_Edit_TraReaction '��Ѫ��Ӧ
        Call FuncTraReaction(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), p����ҽ���´�, mblnMoved)
    Case conMenu_Edit_CriticalAdvice * 100# + 1 To conMenu_Edit_CriticalAdvice * 100# + 99
        Call FuncCriticalAdvice(Control.Parameter, Control.Checked)
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99  '������ҩ���
        If mblnPass Then
            Call zlPASSPati
            Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#)
        End If
    Case conMenu_Edit_Send '����ҽ��
        Call FuncAdviceSend(False)
    Case conMenu_Edit_Send * 100# + 1 '�Զ�����ҽ��
        Call FuncAdviceSend(True)
    Case conMenu_Edit_Blankoff 'ҽ������
        Call FuncAdviceRevoke
    Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
        Call FuncViewDrugExplain(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�շ�ϸĿID)), mfrmParent)
    Case conMenu_Edit_Refcom '�ܾ��������
        Call FuncDrugRefcom 'ҩƷ���ܾ�����
    Case conMenu_Edit_ViewRefcom '�������δͨ����Ϣ
        If Not gobjPass Is Nothing And mlng����ID <> 0 And mlng�Һ�ID <> 0 Then Call gobjPass.ZLPharmReviewResultShow(Me, mlng����ID, mlng�Һ�ID)
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6  '���ġ���ӡ����
        Call FuncEPRReport(Control.ID)
    Case conMenu_Edit_AdvicePay
        Call FuncClinicPay(mfrmParent, mlng����ID, mstr�Һŵ�)
    Case conMenu_Edit_Compend * 10# + 4 '���Ƿ��Ѿ����ĸñ���
        Call FuncExecReportRead(Not Control.Checked)
    Case conMenu_Edit_Compend * 10# + 5 '�Զ���ǲ���״̬
        mblnAutoRead = Not mblnAutoRead
        Call zlDatabase.SetPara("�Զ���Ǳ������״̬", IIF(mblnAutoRead, 1, 0), glngSys, p����ҽ���´�)
    Case conMenu_Edit_MarkMap '��Ƭ����
        RaiseEvent ViewPACSImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    Case conMenu_Edit_MarkKeyMap '�ؼ�ͼ��
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowStaticImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_ViewPacs '������ͼ��ͱ���
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowPatientImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_Test 'Ƥ�Խ��
        Call FuncAdviceTest
    Case conMenu_Tool_SignNew 'ҽ��ǩ��
        Call FuncAdviceSign
    Case conMenu_Tool_SignEarse 'ȡ��ǩ��
        Call FuncAdviceSignErase
    Case conMenu_Tool_SignVerify '��֤ǩ��
        Call FuncAdviceSignVerify
    Case conMenu_Report_ClinicBill * 100# + 1 To conMenu_Report_ClinicBill * 100# + 99 '��ӡ���Ƶ���
        Call FuncBillPrint(Control)
    Case conMenu_Tool_Reference_2 '���ƴ��ϲο�
        Call zlItemRef
    Case conMenu_Tool_Option 'ҽ��ѡ��
        frmOutAdviceSetup.Show 1, mfrmParent
    Case conMenu_Tool_Define '���׷�������
        Call FuncToolScheme
    Case conMenu_Manage_ReportLisView
        Call FuncViewLisRpt
    Case conMenu_Manage_ReportPacsView  '��鱨�����
        Call FuncViewPacsRpt
    Case conMenu_Report_ClinicIndexBill
        Call FuncAdviceIndexBill
    Case conMenu_Manage_RecipeAuditView '�鿴���������
        If InitObjRecipeAudit(p����ҽ���´�) Then
            Call gobjRecipeAudit.ShowResult(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)), mfrmParent)
        End If
    Case conMenu_Report_BloodInstant
        Call PrintBloodReport(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
        If CreatePlugInOK(p����ҽ���´�, mint����) Then
            On Error Resume Next
            If PlugExeNew(Control.Parameter) = False Then
                Call gobjPlugIn.ExecuteFunc(glngSys, p����ҽ���´�, Control.Parameter, _
                    mlng����ID, mstr�Һŵ�, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mlngǰ��ID, mint����)
                Call zlPlugInErrH(err, "ExecuteFunc")
                err.Clear: On Error GoTo 0
            End If
        End If
    End Select
End Sub

Private Function PlugExeNew(ByVal strName As String) As Boolean
'���ܣ����¼�����Ҳ�����ExecuteFunc����
    Dim lngID As Long
    Dim strXML As String
On Error GoTo errH
    If CreatePlugInOK(pסԺҽ���´�, mint����) Then
        With vsAdvice
            lngID = .RowData(.Row)
            strXML = "<ROOT><������Ŀ����>" & .TextMatrix(.Row, col_������Ŀ����) & "</������Ŀ����></ROOT>"
            Call gobjPlugIn.ExecuteFunc(glngSys, p����ҽ���´�, strName, mlng����ID, mstr�Һŵ�, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mlngǰ��ID, mint����, strXML)
        End With
    End If
   PlugExeNew = True
   Exit Function
errH:
    If err.Number = 450 Then
        PlugExeNew = False
        err.Clear
    Else
        PlugExeNew = True
        Call zlPlugInErrH(err, "ExecuteFunc")
        err.Clear: On Error GoTo 0
    End If
End Function


Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnAdvice As Boolean, blnEnabled As Boolean
    Dim i As Integer

    tbcMain.Enabled = mlng����ID <> 0
    For i = 0 To tbcMain.ItemCount - 1
        tbcMain.Item(i).Enabled = mlng����ID <> 0
    Next
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    'Pass
    '����˴������ƣ��� control.Id ������[conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99]������� ʱ,����ҽ���������ֺͰ�ť�ɼ�״̬�л�ı��Pass
    'Enabled����ֵ�������ڶ������������õ�enabled��ֵ���ᱻ���ǡ�
    If Between(Control.ID, conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99) Then
        Control.Visible = IIF(Control.Category <> "", InStr(Control.Category, ";�ɼ�;") > 0, True)
        Control.Enabled = IIF(Control.Category <> "", InStr(Control.Category, ";����;") > 0, True)
        Exit Sub
    End If
    
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
        
    'ҽ����������
    '------------------------------------------------------------------------------
    '�ܵ��ж�:�޲��˻����ﲡ�˲������κβ���
    If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 998) _
        Or Between(Control.ID, conMenu_Edit_NewItem * 10#, (conMenu_Edit_NewItem + 998) * 10# + 9) Then '���������Ӳ˵�
        Control.Enabled = mlng����ID <> 0 _
            And (Control.ID <> conMenu_Edit_Blankoff And mblnEditable Or Control.ID = conMenu_Edit_Blankoff And mblnCanRevoke _
            Or Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Or Control.ID = conMenu_Edit_MarkKeyMap Or Control.ID = conMenu_Edit_Compend _
            Or Between(Control.ID, conMenu_Edit_Compend * 10# + 1, conMenu_Edit_Compend * 10# + 5))
        If Not Control.Enabled Then Exit Sub
    End If

    blnAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
    Select Case Control.ID
    Case conMenu_Edit_NewItem, conMenu_Edit_Apply, conMenu_Edit_LISApply, conMenu_Edit_PacsApply, conMenu_Edit_BloodApply, conMenu_Edit_OperationApply, conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101#  '�¿�ҽ��
        Control.Enabled = (mvRegDate <> CDate("3000-01-01"))
        
        blnEnabled = Control.Enabled
        If blnEnabled And mint���� = 0 Then
            If mstr����ҽ�� <> UserInfo.���� Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Modify, conMenu_Edit_Delete '�޸�ҽ��,ɾ��ҽ��
        blnEnabled = blnAdvice _
            And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 1 _
            And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 0
        
        If blnEnabled And mint���� = 2 Then
            blnEnabled = InStr("," & mstrǰ��IDs & ",", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) & ",") > 0
        ElseIf blnEnabled And mint���� <> 2 Then
            blnEnabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) = 0
        End If
      
        If blnEnabled And mint���� = 0 Then
            If mstr����ҽ�� <> UserInfo.���� Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Manage_RecipeAuditView
        blnEnabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�������״̬)) <> 0
        Control.Enabled = blnEnabled
    '���뵥ȡ��
    Case conMenu_Edit_ApplyDel
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And _
                    (.TextMatrix(.Row, COL_�������) = "D" Or .TextMatrix(.Row, COL_�������) = "F" Or _
                        Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E" Or _
                        .TextMatrix(.Row, COL_�������) = "K")) Then
                    blnEnabled = Val(.TextMatrix(.Row, COL_�������)) <> 0
                End If
                '��Ѫҽ������˲�����ȡ������Ѫ���������ݣ�
                If blnEnabled = True And .TextMatrix(.Row, COL_�������) = "K" And .TextMatrix(.Row, COL_ҽ��״̬) = "1" Then
                    If Val(.TextMatrix(.Row, COL_��鷽��)) = 1 And Val(.TextMatrix(.Row, COL_���״̬)) = 1 Then blnEnabled = False
                End If
            End With
        End If
        If blnEnabled And mint���� = 0 Then
            If mstr����ҽ�� <> UserInfo.���� Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    '��������޸�
    Case conMenu_Edit_ApplyModi
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Val(.TextMatrix(.Row, COL_�������)) <> 0 Then
                    If Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E" Then
                        If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E") Then blnEnabled = False
                    ElseIf .TextMatrix(.Row, COL_�������) = "D" And .TextMatrix(.Row, COL_��������) <> "����" Then
                        If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And .TextMatrix(.Row, COL_�������) = "D") Then blnEnabled = False
                    ElseIf .TextMatrix(.Row, COL_�������) = "K" Then
                        If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And .TextMatrix(.Row, COL_�������) = "K") Then blnEnabled = False
                    ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                        If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And .TextMatrix(.Row, COL_�������) = "F") Then blnEnabled = False
                    Else
                        blnEnabled = Val(.TextMatrix(.Row, COL_�������)) <> 0
                    End If
                Else
                    blnEnabled = False
                End If
            End With
        End If
        If blnEnabled And mint���� = 0 Then
            If mstr����ҽ�� <> UserInfo.���� Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRis
        blnEnabled = False
        With vsAdvice
            If InStr(",D,F,", .TextMatrix(.Row, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_��������))) > 0 And .TextMatrix(.Row, COL_�������) = "E" Then
                blnEnabled = True
            End If
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisSch
        blnEnabled = False
        If gbln����Ӱ����ϢϵͳԤԼ Then
            With vsAdvice
                If (InStr(",D,F,", .TextMatrix(.Row, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_��������))) > 0 And .TextMatrix(.Row, COL_�������) = "E") And Val(.TextMatrix(.Row, COL_RISԤԼID)) = 0 Then
                    blnEnabled = True
                End If
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisDel, conMenu_Tool_RisPrint
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RISԤԼID)) <> 0
    Case conMenu_Edit_NewRisModi
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RISԤԼID)) <> 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 8
    '�鿴����
    Case conMenu_Edit_ApplyView
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (InStr(",K,F,D,", .TextMatrix(.Row, COL_�������)) > 0) Then blnEnabled = Val(.TextMatrix(.Row, COL_�������)) <> 0
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_TraReaction
        With vsAdvice
            blnEnabled = (.TextMatrix(.Row, COL_�������) = "K") And Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 And gblnѪ��ϵͳ
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Blankoff 'ҽ������
        blnEnabled = blnAdvice _
            And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 8 _
            And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 0 Or Not gobjESign Is Nothing)
            
        '��Ȩ���������ϣ�
        '��Ȩ�ޣ�ҽ��վ�Ĺ�����ҽ��ֻ���ڱ����Ҳ������ϣ���ҽ��վ��ҽ��վ����������
        If blnEnabled And mint���� = 2 Then
            blnEnabled = InStr("," & mstrǰ��IDs & ",", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) & ",") > 0 _
            And vsAdvice.TextMatrix(vsAdvice.Row, COL_����ҽ��) = UserInfo.���� Or InStr(GetInsidePrivs(p����ҽ���´�), "��������ҽ��") > 0
        ElseIf blnEnabled And mint���� <> 2 Then
            blnEnabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) = 0 _
            And vsAdvice.TextMatrix(vsAdvice.Row, COL_����ҽ��) = UserInfo.���� Or InStr(GetInsidePrivs(p����ҽ���´�), "��������ҽ��") > 0
        End If
        If blnEnabled And mint���� = 0 Then
            If mstr����ҽ�� <> UserInfo.���� Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Send, conMenu_Edit_Send * 100# + 1 'ҽ������
        blnEnabled = True
        If mint���� = 0 Then
            If mstr����ҽ�� <> UserInfo.���� Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_CriticalAdvice
        blnEnabled = False
        If Not mrsΣ��ֵ Is Nothing Then
            If Not mrsΣ��ֵ.EOF Then
                blnEnabled = True
            End If
        End If
        If blnEnabled Then
            blnEnabled = (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 4 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0)
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
        Control.Enabled = blnAdvice And InStr(",5,6,7,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) > 0
    Case conMenu_Edit_Test 'Ƥ�Խ��
        With vsAdvice
            Control.Enabled = blnAdvice _
                And Val(.TextMatrix(.Row, COL_ǰ��ID)) = 0 _
                And Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 4 _
                And .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "1"
        End With
    Case conMenu_Report_ClinicIndexBill
        blnEnabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 8
        Control.Enabled = blnEnabled
    Case conMenu_Edit_MediAudit 'ҩ�����(��ҩ����ʾ)
        If mblnPass Then
            Call gobjPass.zlPassCommandBarUpdate(mobjPassMap, Control, blnAdvice)
        End If
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6 '���ġ���ӡ����
        If Not gobjExchange Is Nothing Then
            Control.Enabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) <> 0
        Else
            Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_��鱨��ID) <> "" Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS����ID)) <> 0)
        End If
        
        If Control.ID = conMenu_Edit_Compend * 10# + 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        ElseIf Control.ID = conMenu_Edit_Compend * 10# + 6 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 2 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_Compend * 10# + 4 '���Ѿ����ĸñ���
        Control.Checked = Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_����״̬)) = 1
        Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_��鱨��ID) <> "")
    Case conMenu_Edit_Compend * 10# + 5 '�Զ���ǲ���״̬
        Control.Checked = mblnAutoRead
        Control.Enabled = mblnAutoReadEnabled
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs '��Ƭ����
        blnEnabled = blnAdvice And InStr(",4,5,6,7,8,9,H,M,Z,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) = 0 ' And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)) <> 0
        If blnEnabled Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 8 Then
                blnEnabled = False
            End If
        End If
        Control.Enabled = blnEnabled
    End Select

    Select Case Control.ID
    Case conMenu_Report_ClinicBill '��ӡ���Ƶ���
        Control.Enabled = Control.CommandBar.Controls.Count > 0
    End Select
    
    '����ǩ������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_SignNew 'ҽ��ǩ��
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Tool_SignVerify, conMenu_Tool_SignEarse '��֤ǩ��,ȡ��ǩ��
        blnEnabled = mlng����ID <> 0 And blnAdvice And tbcAppend.Selected.Tag = "ǩ��" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        If blnEnabled Then blnEnabled = vsAppend.RowData(vsAppend.Row) <> 0
        Control.Enabled = blnEnabled
    End Select
    
    '��������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnAdvice
    Case conMenu_View_Append '������Ϣ
        Control.Checked = tbcAppend.Visible
    Case conMenu_View_Hide '�Զ����ع��˹�����
        Control.Checked = mblnHideFilter
    Case conMenu_Manage_ReportLisView, conMenu_Manage_ReportPacsView '��飬������
        Control.Enabled = mlng����ID <> 0
    End Select
    
    '��Ѫִ�е���ӡ
    If Control.ID = conMenu_Report_BloodInstant Then 'ִ�е���ӡ
        Control.Visible = InStr(GetInsidePrivs(9005, , 2200), ";��Ѫִ�д�ӡ;") <> 0
        Control.Enabled = vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "K" And Control.Visible
    End If
    
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strItem As String

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    
    '���Ȩ���ж�
    '------------------------------------------------------------------------------
    If InStr(UserInfo.����, "ҽ��") = 0 Then
        If Control.ID = conMenu_EditPopup Then blnVisible = False
        If Control.ID = conMenu_ReportPopup Then blnVisible = False
        If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 999) Then blnVisible = False
    End If

    'ҽ����������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_LISApply, conMenu_Edit_ApplyModi, conMenu_Edit_ApplyDel
        '�¿�ҽ��,�޸�ҽ��,ɾ��ҽ��,�������롢�޸ġ�ɾ��
        If InStr(GetInsidePrivs(p����ҽ���´�), "ҽ���´�") = 0 Then blnVisible = False
        
    Case conMenu_Edit_Blankoff
        'ҽ������
        If InStr(GetInsidePrivs(p����ҽ���´�), "ҽ������") = 0 Then blnVisible = False
    Case conMenu_Edit_Send, conMenu_Edit_Send * 100# + 1
        'ҽ������
        If mSendControl Is Nothing Then Set mSendControl = Control
        If InStr(GetInsidePrivs(p����ҽ���´�), "ҽ������") = 0 Then
            blnVisible = False
        ElseIf InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ�շѵ�") = 0 And InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ���ʵ�") = 0 Then
            blnVisible = False
        End If
        If Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 0 And InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ�շѵ�") = 0 Or _
           Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 1 And InStr(GetInsidePrivs(p����ҽ���´�), "����Ϊ���ʵ�") = 0 Then
            blnVisible = False
        End If
    Case conMenu_Edit_Test
        'Ƥ��ҽ�����
        If InStr(GetInsidePrivs(p����ҽ���´�), "Ƥ��ҽ�����") = 0 Then blnVisible = False
    Case conMenu_Edit_TraReaction
        '��Ѫ��Ӧ�Ǽ�
        If Not (InStr(GetInsidePrivs(9005, , 2200), "��Ѫ��Ӧ�Ǽ�") <> 0 And gblnѪ��ϵͳ) Then blnVisible = False
    
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1
        '���浯��(����ӡ),���ı���
        If InStr(GetInsidePrivs(p����ҽ���´�), ";�������;") = 0 Then blnVisible = False
    Case conMenu_Edit_Compend * 10# + 2, conMenu_Edit_Compend * 10# + 3
        '��ӡ����
        If InStr(GetInsidePrivs(p����ҽ���´�), ";�����ӡ;") = 0 Then blnVisible = False
    Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
        If gobjDrugExplain Is Nothing Or InStr(GetInsidePrivs(p����ҽ���´�), ";ҩƷ˵����;") = 0 Then blnVisible = False
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs
        '��Ƭ����
        If GetInsidePrivs(pXWPACS��Ƭ) <> "" And InStr(GetInsidePrivs(p����ҽ���´�), ";��Ƭ����;") <> 0 Then
            blnVisible = True
        Else
            If Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Then
                If InStr(GetInsidePrivs(p����ҽ���´�), ";��Ƭ����;") = 0 Or GetInsidePrivs(p��Ƭ���߹���) = "" Then
                    blnVisible = False
                End If
            Else
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99
        '������ҩ���
        strItem = GetInsidePrivs(p����ҽ���´�)
        If InStr(strItem, "������ҩ���") = 0 Then blnVisible = False
    Case conMenu_Edit_AdvicePay
        blnVisible = InStr(GetInsidePrivs(p����ҽ���´�), ";����޿�֧��;") > 0
    End Select
        
    '����ǩ������
    Control.Category = "���ж�"
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_Sign, conMenu_Tool_SignNew '����ǩ��,ҽ��ǩ��
        If InStr(UserInfo.����, "ҽ��") = 0 Or gobjESign Is Nothing _
            Or InStr(GetInsidePrivs(p����ҽ���´�), ";ҽ���´�;") = 0 Then
            blnVisible = False
        ElseIf mblnSignVisible = False Then
            blnVisible = False '��ͬ����û������Ҫʹ��ǩ��
        End If
        Control.Category = ""  'ǩ����ť��̬�жϿɼ���
    End Select

    Control.Enabled = blnVisible
    Control.Visible = blnVisible
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal str�Һŵ� As String, ByVal blnEditable As Boolean, _
        Optional ByVal blnMoved As Boolean, Optional ByVal lngǰ��ID As Long, Optional ByVal lng�������ID As Long, _
    Optional ByRef objMip As Object, Optional ByVal lngǰ�����ID As Long, Optional ByVal lng·��״̬ As Long = -1, _
    Optional ByVal int�������� As Integer)
'���ܣ�ˢ������ҽ������
'������lngǰ��ID=����ҽ��վ����ʱ����
'      blnMoved=�ò��˵������Ƿ���ת��
'      blnEditable=�ɷ�Բ���ҽ�����б༭
'      objMip ��Ϣ����
'      lngǰ�����ID= lngǰ��ID����ҽ����Ӧ��ִ�п���ID����ҽ��վ���������������ʱ��lng�������ID<>lngǰ�����ID  lngǰ�����ID�������봫��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg  As String
    Dim objControl As CommandBarControl
    Dim lng����IDOld As Long, str�Һŵ�Old As String
    Dim lng���˿���ID As Long
    Dim lngOld�������ID As Long

    
    lng����IDOld = mlng����ID
    str�Һŵ�Old = mstr�Һŵ�
    lngOld�������ID = mlng�������ID
    
    mlng����ID = lng����ID
    mstr�Һŵ� = str�Һŵ�
    mlngǰ��ID = lngǰ��ID
    mlng�������ID = lng�������ID
    mblnEditable = blnEditable
    mblnCanRevoke = blnEditable
    mblnMoved = blnMoved
    mlng·��״̬ = lng·��״̬
    mint�������� = int��������
    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng����ID, 0, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
        
    '��ȡ�������Ϣ
    If lng����ID <> 0 And str�Һŵ�Old <> mstr�Һŵ� Then
        strSQL = "Select A.ID,A.ִ�в���ID,A.ִ��״̬,A.�Ǽ�ʱ��,C.����,Nvl(Nvl(A.�������ID,Decode(A.ת��״̬,1,A.ת�����ID,NULL)),A.ִ�в���ID) as ���˿���ID,a.�����,c.����,a.ִ����" & _
            " From ���˹Һż�¼ A,������Ϣ C Where C.����id=A.����id And A.NO=[1] And a.��¼����=1 And a.��¼״̬=1"
        If mblnMoved Then strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mstr�Һŵ�, mlng����ID)
        If Not rsTmp.EOF Then
            mvRegDate = rsTmp!�Ǽ�ʱ��
            
            '����ҽ��������Ϊ����(������Ϊ���һ�ιҺţ�Ҫ����Ϊ�������ͬһ�ζ���ҹҺž���ʱ�����������˵�ҽ��)
            If Not mblnCanRevoke Then
                If NVL(rsTmp!ִ��״̬, 0) = 1 Then
                    mblnCanRevoke = True
                End If
            End If
            mlng�Һ�ID = Val(rsTmp!ID & "")
            mstr���� = rsTmp!���� & ""
            mstr����� = rsTmp!����� & ""
            
            'Ӥ������
            mlng�Һſ���ID = Val("" & rsTmp!ִ�в���ID)
            lng���˿���ID = Val("" & rsTmp!���˿���id)
            mbln���� = DeptIsWoman(rsTmp!ִ�в���ID)
            mint���� = Val("" & rsTmp!����)
            If mbln���� Then
                '��ȡ���ȱʡֵ��-1=����,0=����,1-Ӥ��1
                mvarCond.Ӥ�� = Val(zlDatabase.GetPara("����Ӥ������", glngSys, p����ҽ���´�, "0"))
            End If
            mstr����ҽ�� = "" & rsTmp!ִ����
        Else
            mlng�Һſ���ID = 0
            mvRegDate = CDate("3000-01-01") '��첡��,ҽ��վ��ʱ�ǼǵĲ��ˣ�δ�Һţ�
            mbln���� = False
            mvarCond.Ӥ�� = -1
            mlng�Һ�ID = 0
            mstr����ҽ�� = ""
        End If
        Call GetCriticalData
        On Error GoTo 0
    End If
    
    If (lngOld�������ID <> mlng�������ID Or lng����IDOld <> mlng����ID) And mlngǰ��ID <> 0 Then
        mstrǰ��IDs = Getҽ������ҽ��IDs(mlng����ID, mstr�Һŵ�, IIF(0 = lngǰ�����ID, mlng�������ID, lngǰ�����ID), False, mlngǰ��ID)
    ElseIf mlngǰ��ID = 0 Then
        mstrǰ��IDs = ""
    End If
    
    If Visible And lng���˿���ID <> 0 Then
        mblnSignVisible = True
        If mint���� = 0 Then
            If CheckSign(0, 0, mlng�������ID, lng���˿���ID, 1, False, gobjESign) = False Then
                mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
            End If
        ElseIf mint���� = 2 Then
            If CheckSign(3, 0, mlng�������ID, lng���˿���ID, 1, False, gobjESign) = False Then
                mblnSignVisible = False
            End If
        End If
    End If
    
    If Not grsTube Is Nothing Then
        If grsTube.State = 1 Then grsTube.Close
        Set grsTube = Nothing
    End If
    
    If lng����IDOld <> mlng����ID Then
        If mblnPass Then
            Call zlPASSPati
            On Error Resume Next
            Call gobjPass.zlPassClearLight(mobjPassMap)    '��ʼ��״̬��
            On Error GoTo 0
        End If
    End If
    
    'ˢ������
    Call RefreshData
    
    'ִ���Զ�������ܣ�����ID=0Ҳ���ã���ʵ����رս���
    If mlngPlugInID <> 0 And lng����ID <> 0 And str�Һŵ�Old <> mstr�Һŵ� Then
        Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        If Not objControl Is Nothing Then
            objControl.Execute
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshData()
'���ܣ�ˢ������
    If mlng����ID = 0 Then
        '���ҽ���嵥
        Call ClearAdviceData
        Call ClearAppendData
        RaiseEvent StatusTextUpdate("")
    Else
        '��ʾҽ���嵥
        Call LoadAdvice
        '��ʾҽ�����
        Call ShowTotalMoney
    End If
End Sub

Private Sub Refresh����()
'���ܣ��ڱ���ҳ�治ͬ����֮���л�ʱ�����ˢ�£������¶����ݿ����ñ������غ���ʾ
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lngҽ��ID As Long
    If mvarCond.����ģʽ = 0 Then Exit Sub
    With vsAdvice
    
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))      '��¼��ǰ��������ڵ�ǰ����ˢ��ҽ����Ӧ�ò���
        
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_������)) <> 0 Then
                If mvarCond.���� = 0 Then ' ȫ��
                    blnTmp = True
                ElseIf mvarCond.���� = 1 Then ' ���
                    blnTmp = .TextMatrix(i, COL_�������) = "D"
                ElseIf mvarCond.���� = 2 Then '����
                    blnTmp = (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "C")
                ElseIf mvarCond.���� = 3 Then ' ����
                    blnTmp = Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "D" Or .TextMatrix(i, COL_�������) = "C")
                End If
                
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                
                .RowHidden(i) = Not blnTmp
            Else
                .RowHidden(i) = True: .RowHeight(i) = 0
            End If
            '���ӹ���δ���ı�����ѳ��ı���
            If .RowHidden(i) = False Then
                blnTmp = IIF(.TextMatrix(i, COL_����״̬) = "δ��", mvarCond.δ������, mvarCond.�ѳ�����)
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                .RowHidden(i) = Not blnTmp
            End If
        Next
    End With
    Call LocatedDefaultAdviceRow(lngҽ��ID)
    Call SetAdviceColVisible
End Sub

Private Sub Refresh����()
'���ܣ��ڱ���ҳ�治ͬ����֮���л�ʱ�����ˢ�£������¶����ݿ����ñ������غ���ʾ
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lngҽ��ID As Long

    If mvarCond.����ģʽ = 3 Then Exit Sub
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowHidden(i) Then
                .RowHidden(i) = False
            End If
            .TextMatrix(i, COL_������ӡ) = .Cell(flexcpData, i, COL_������ӡ)
            .TextMatrix(i, COL_������) = .Cell(flexcpData, i, COL_������)
            .TextMatrix(i, COL_����Ԥ��) = .Cell(flexcpData, i, COL_����Ԥ��)
            If Val(.TextMatrix(i, COL_ID)) = 0 Then
                .RemoveItem i
            End If
        Next
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))      '��¼��ǰ��������ڵ�ǰ����ˢ��ҽ����Ӧ�ò���
        For i = .FixedRows To .Rows - 1
            Select Case mvarCond.ҽ��
            Case 0
            Case 1
            Case Else
                If .TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6" Or .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 4 Then
                    If mvarCond.ҽ�� = 2 Then .RowHidden(i) = True
                Else
                    If mvarCond.ҽ�� = 1 Then .RowHidden(i) = True
                End If
            
            End Select
            
            If .TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6" Or .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 4 Then
                If mvarCond.ҽ�� = 2 Then
                    .RowHidden(i) = True
                End If
            Else
                If mvarCond.ҽ�� = 1 Then
                    .RowHidden(i) = True
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    Call LocatedDefaultAdviceRow(lngҽ��ID)
    Call SetAdviceColVisible
End Sub

Private Sub LocatedDefaultAdviceRow(Optional ByVal lngҽ��ID As Long)
'���ܣ�ҽ���嵥��ȱʡ��λ�������ҽ��id����ҽ��id��λ
    'ȱʡ��λ����ǰѡ���ҽ��Ϊ��ʾ����λ������λ�����һ�С�
    Dim i As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        .Row = .Rows - 1
        If lngҽ��ID <> 0 Then
            lngҽ��ID = .FindRow(CStr(lngҽ��ID), , COL_ID)
            If lngҽ��ID <> -1 Then
                If Not .RowHidden(lngҽ��ID) Then .Row = lngҽ��ID
            End If
        End If
        If .RowHidden(.Row) Then    '��λ���������еĴ���
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            For i = .Row - 1 To .FixedRows Step -1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            .AddItem "": .Row = .Rows - 1
        End If
        If .Row < .FixedRows Then
            .AddItem "": .Row = .Rows - 1
        End If
        .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        .Refresh
    End With
End Sub

Public Sub zlItemRef()
'���ܣ��������Ʋο�
    Dim lng������ĿID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_�������) = "E" And (RowIs�䷽��(.Row) Or RowIs������(.Row)) Then
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    
    'ToDo:���ƴ�ʩ�ο�
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
'''''''
    Dim objControl As CommandBarControl
    Dim arrBaby As Variant, i As Long
    Dim strTmp As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case ID_Ӥ��
        strTmp = IIF(mvarCond.����ģʽ = 3, "����", "ҽ��")
        With CommandBar.Controls
            .DeleteAll
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100#, "����" & strTmp)
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 1, "����" & strTmp): objControl.BeginGroup = True
            For i = 0 To 4
                Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + i + 2, "Ӥ�� " & (i + 1) & " " & strTmp)
                If i = 0 Then objControl.BeginGroup = True
            Next
        End With
    Case Else
        Call zlPopupCommandBars(CommandBar)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ActiveHotKey(KeyCode, Shift)
End Sub

Private Sub fraHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timHide.Enabled = True
End Sub

Private Sub mfrmEdit_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���Id As String, ByRef blnNo As Boolean)
'    If InStr(";" & GetPrivFunc(glngSys, p���ﲡ������) & ";", ";������д;") > 0 Then
        RaiseEvent CheckInfectDisease(blnOnChek, str����ID, str���Id, blnNo)
'    End If
End Sub

Private Sub mfrmEdit_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, �Һŵ�, Succeed)
End Sub

Private Sub mfrmEdit_FormUnload(Cancel As Integer)
    If mlngΣ��ֵID <> 0 Then
        Call GetCriticalData
    End If
    mlngΣ��ֵID = 0
    If Not Cancel Then
        If mfrmEdit.mblnOK Then
            RaiseEvent RequestRefresh
            'Call LoadAdvice
            'Call ShowTotalMoney
        End If
        Set mfrmEdit = Nothing
        
        If Me.Visible Then
            Call BringWindowToTop(Me.hwnd)
        End If
    End If
    RaiseEvent Activate
End Sub

Private Sub mfrmSend_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, �Һŵ�, Succeed)
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim strSQL As String
    
    '���뵥�ݴ�ӡ֮��Ĵ���
    If mstrBillPrint <> "" Then
        If Split(mstrBillPrint, ",")(0) = ReportNum Then
            strSQL = "Zl_���Ƶ��ݴ�ӡ_Insert('" & Split(mstrBillPrint, ",")(1) & "'," & Val(Split(mstrBillPrint, ",")(2)) & ",1,'" & UserInfo.���� & "')"
        End If
    End If
    
    On Error GoTo errH
    If strSQL <> "" Then
        zlDatabase.ExecuteProcedure strSQL, Me.Name
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub timHide_Timer()
'���ܣ�������˹��������Զ���ʾ������
    Dim vPos As PointAPI, vRect As RECT
    Static sngBegin As Single
    
    If Not mblnHideFilter Then
        timHide.Enabled = False
        sngBegin = 0: Exit Sub
    End If
    
    If sngBegin = 0 Then sngBegin = Timer
    GetCursorPos vPos
    
    If fraHide.Visible Then
        ScreenToClient Me.hwnd, vPos
        If vPos.X * Screen.TwipsPerPixelX >= fraHide.Left And vPos.X * Screen.TwipsPerPixelX <= fraHide.Left + fraHide.Width _
            And vPos.Y * Screen.TwipsPerPixelY >= fraHide.Top And vPos.Y * Screen.TwipsPerPixelY <= picMain.Top + fraHide.Top + fraHide.Height Then
            fraHide.BackColor = cbsSub.GetSpecialColor(XPCOLOR_SEPARATOR)
            If Timer - sngBegin >= 0.35 Then
                fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
                fraHide.Visible = False: cbsSub(2).Visible = True
                sngBegin = 0: cbsSub.RecalcLayout
            End If
        Else
            fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
            sngBegin = 0: timHide.Enabled = False
        End If
    ElseIf cbsSub(2).Visible Then
        cbsSub(2).GetWindowRect vRect.Left, vRect.Top, vRect.Right, vRect.Bottom
        If Not (vPos.X >= vRect.Left / Screen.TwipsPerPixelX And vPos.X <= vRect.Right / Screen.TwipsPerPixelX _
            And vPos.Y >= vRect.Top / Screen.TwipsPerPixelY And vPos.Y <= vRect.Bottom / Screen.TwipsPerPixelY) Then
            If Timer - sngBegin >= 1 Then
                sngBegin = 0: timHide.Enabled = False
                fraHide.Visible = True: cbsSub(2).Visible = False
                cbsSub.RecalcLayout
            End If
        Else
            sngBegin = 0
        End If
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    If Control.ID <> 0 Then
        If cbsSub.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
        vsColumn.Visible = False
    End If

    Select Case Control.ID
        Case ID_Ӥ�� * 100# '����ҽ��
            If mvarCond.Ӥ�� = -1 Then Exit Sub
            mvarCond.Ӥ�� = -1
            Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, p����ҽ���´�)
        Case ID_Ӥ�� * 100# + 1 To ID_Ӥ�� * 100# + 6 '���ˡ�Ӥ��ҽ��
            If mvarCond.Ӥ�� = Control.ID - ID_Ӥ�� * 100# - 1 Then Exit Sub
            mvarCond.Ӥ�� = Control.ID - ID_Ӥ�� * 100# - 1
            Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, p����ҽ���´�)
        Case ID_ȫ��
            mvarCond.���� = 0
        Case ID_���
            mvarCond.���� = 1
        Case ID_����
            mvarCond.���� = 2
        Case ID_δ������
            If mvarCond.δ������ Then
                If mvarCond.�ѳ����� Then
                    mvarCond.δ������ = Not mvarCond.δ������
                End If
            Else
                mvarCond.δ������ = Not mvarCond.δ������
            End If
        Case ID_�ѳ�����
            If mvarCond.�ѳ����� Then
                If mvarCond.δ������ Then
                    mvarCond.�ѳ����� = Not mvarCond.�ѳ�����
                End If
            Else
                mvarCond.�ѳ����� = Not mvarCond.�ѳ�����
            End If
        Case ID_����
            mvarCond.���� = 3
        Case ID_ҽ��ȫ��
            mvarCond.ҽ�� = 0
        Case ID_ҽ������
            mvarCond.ҽ�� = 1
        Case ID_ҽ������
            mvarCond.ҽ�� = 2
        Case ID_��ֹ
            mvarCond.��ֹ = Not mvarCond.��ֹ
        Case ID_����
            mvarCond.���� = Not mvarCond.����
        Case ID_���
            mvarCond.��ʾģʽ = 0
        Case ID_����
            mvarCond.��ʾģʽ = 1
    End Select
    
    bln���� = InStr("," & ID_δ������ & "," & "," & ID_�ѳ����� & "," & "," & ID_ȫ�� & "," & ID_��� & "," & ID_���� & "," & ID_���� & ",", "," & Control.ID & ",") > 0
    
    bln���� = InStr("," & ID_ҽ��ȫ�� & "," & ID_ҽ������ & "," & ID_ҽ������ & ",", "," & Control.ID & ",") > 0
    
    cbsSub.RecalcLayout
    If bln���� Then
        Call Refresh����
    ElseIf bln���� Then
        Call Refresh����
    Else
        Call RefreshData
    End If
    
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Control.Enabled = mlng����ID <> 0
    If Not Control.Enabled Then Exit Sub
    
    Select Case Control.ID
        Case ID_ȫ��
            Control.Checked = mvarCond.���� = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_���
            Control.Checked = mvarCond.���� = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_����
            Control.Checked = mvarCond.���� = 2
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_����
            Control.Checked = mvarCond.���� = 3
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        
        Case ID_ҽ��ȫ��
            Control.Checked = mvarCond.ҽ�� = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 0
        Case ID_ҽ������
            Control.Checked = mvarCond.ҽ�� = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 0
        Case ID_ҽ������
            Control.Checked = mvarCond.ҽ�� = 2
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 0
        
        Case ID_Ӥ�� 'Ӥ��ҽ������
            If mbln���� Then
                Control.Visible = True
                
                If mvarCond.Ӥ�� = -1 Then
                    Control.Caption = IIF(mvarCond.����ģʽ = 3, "���б���", "����ҽ��")
                ElseIf mvarCond.Ӥ�� = 0 Then
                    Control.Caption = IIF(mvarCond.����ģʽ = 3, "���˱���", "����ҽ��")
                Else
                    Control.Caption = "Ӥ�� " & mvarCond.Ӥ��
                End If
            Else
                If mvarCond.Ӥ�� <> -1 Or Control.Visible Then
                    mvarCond.Ӥ�� = -1
                    Control.Visible = False
                    Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, p����ҽ���´�)
                End If
            End If
        Case ID_Ӥ�� * 100# '����ҽ��
            Control.Checked = mvarCond.Ӥ�� = -1
        Case ID_Ӥ�� * 100# + 1 To ID_Ӥ�� * 100# + 6 '���ˡ�Ӥ��ҽ��
            Control.Checked = mvarCond.Ӥ�� = Control.ID - ID_Ӥ�� * 100# - 1
        Case ID_��ֹ
            Control.Checked = mvarCond.��ֹ
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
        Case ID_����
            If mint���� <> 2 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Checked = mvarCond.����
                Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            End If
        Case ID_���
            Control.Checked = mvarCond.��ʾģʽ = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case ID_����
            Control.Checked = mvarCond.��ʾģʽ = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ <> 3
            
        Case ID_δ������
            Control.Checked = mvarCond.δ������
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_�ѳ�����
            Control.Checked = mvarCond.�ѳ�����
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.����ģʽ = 3
    End Select
End Sub

Private Sub cbsSub_Resize()
    Dim BarHideH As Long, PriceH As Long
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If cbsSub.Count >= 2 Then
        If Not cbsSub(2).Visible Then BarHideH = fraHide.Height
    End If
    
    On Error Resume Next
    If fraMore.Visible Then
        fraMore.Tag = ""
        fraMore.Visible = False
    End If
    
    PriceH = IIF(tbcAppend.Visible, fraAdviceUD.Height + tbcAppend.Height, 0)
    
    fraHide.Left = lngLeft
    fraHide.Top = lngTop
    fraHide.Width = lngRight - lngLeft
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = lngTop + BarHideH
    vsAdvice.Width = lngRight - lngLeft
    vsAdvice.Height = lngBottom - lngTop - PriceH - BarHideH
    
    '��ѡ����
    With vsAdvice
        fraColSel.Left = .Left + (.ColWidth(COL_F��־) + .ColWidth(COL_F����) - fraColSel.Width) / 2 + 30
        fraColSel.Top = .Top + (225 - fraColSel.Height) / 2 + 30
    End With
    
    fraAdviceUD.Left = lngLeft
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = vsAdvice.Width
    
    tbcAppend.Left = lngLeft
    tbcAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tbcAppend.Width = vsAdvice.Width
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim blnTmp As Boolean
    If Not Me.Visible Then Exit Sub
    Select Case Item.Tag
    Case "ҽ��"
        mvarCond.����ģʽ = 0
        mvarCond.ҽ�� = 0
        mvarCond.���� = 0
    Case "����"
        mvarCond.����ģʽ = 3
        mvarCond.ҽ�� = 0
        mvarCond.���� = 0
    End Select

    If Item.Tag <> "" And mlng����ID <> 0 Then
        Call AddToolBarInDoctor
        Call RefreshData
    End If
End Sub

Private Sub Form_Activate()
    If Me.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
    RaiseEvent Activate
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If fraMore.Visible = True Then
        fraMore.Tag = ""
        fraMore.Visible = False
        PicAdviceDetail.Visible = False
    End If
End Sub

Private Sub vsAdvice_DblClick()
    Dim lngҽ��ID As Long
    Dim lngNo As Long
    Dim bln��Ѫ As Boolean
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
    End If
    '˫����ҽ����������뵥��ʽ�´�ĵ����鿴���� ��Ѫ�������������飬����
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_�������))
        
        If lngҽ��ID <> 0 And lngNo <> 0 Then
            If .TextMatrix(.Row, COL_�������) = "K" Then
                bln��Ѫ = Val(.TextMatrix(.Row, COL_��鷽��)) = 1
                '��Ѫ
                If Val(Mid(gstrOutUseApp, 3, 1)) = 1 Then
                    If gblnѪ��ϵͳ = True Then
                        Call frmApplyBloodNew.ShowMe(Me, mlng����ID, 0, 1, 2, lngҽ��ID, mlng�Һſ���ID, , mlng�Һſ���ID, , , mrsDefine, mclsMipModule, 1, mstr�Һŵ�, , , , , mlngǰ��ID, IIF(bln��Ѫ = True, 1, 0))
                    Else
                        Call frmApplyBlood.ShowMe(Me, mlng����ID, 0, 1, 2, lngҽ��ID, mlng�Һſ���ID, , mlng�Һſ���ID, , , mrsDefine, mclsMipModule, 1, mstr�Һŵ�, , , , , mlngǰ��ID)
                    End If
                End If
                                
            ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                '����
                If Val(Mid(gstrOutUseApp, 4, 1)) = 1 Then Call frmApplyOperation.ShowMe(Me, 1, 2, mlng����ID, mstr�Һŵ�, 1, lngҽ��ID)
               
            ElseIf .TextMatrix(.Row, COL_�������) = "D" Then
                '���
                If Val(Mid(gstrOutUseApp, 1, 1)) = 1 Then
                    Call ShowApply���(Me, lngNo)
                End If
            ElseIf .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "6" Then
                '����
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, strPrompt As String
    
    With vsAdvice
        lngRow = .MouseRow
        If Button = 0 And lngRow > 0 Then  '���ģʽ���Ը���
            If .MouseCol = col_���� Then
                If Val(fraMore.Tag) <> lngRow Then
                    fraMore.Visible = False
                    fraMore.Tag = lngRow
                    If lngRow = .Row Then
                        fraMore.BackColor = .BackColorSel
                    Else
                        fraMore.BackColor = .BackColor
                    End If
                    fraMore.Height = .RowHeight(lngRow) - 10
                    If fraMore.Height > 250 Then fraMore.Height = 250
                    
                    fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - fraMore.Height
                    If fraMore.Top + fraMore.Height > .Top + .Height - IIF(Grid.HScrollVisible(vsAdvice), 230, 0) Then Exit Sub
                    
                    fraMore.Left = .Left + .ColPos(col_����) + IIF(.ColWidth(col_����) > .ColWidthMax, .ColWidthMax, .ColWidth(col_����)) - fraMore.Width
                    fraMore.Visible = True
                ElseIf PicAdviceDetail.Visible = True Then
                    fraMore.Tag = ""
                    fraMore.Visible = False
                    PicAdviceDetail.Visible = False
                End If
            Else
                If fraMore.Visible = True Then
                    fraMore.Tag = ""
                    fraMore.Visible = False
                    PicAdviceDetail.Visible = False
                End If
                
                strPrompt = ""
                If .MouseCol = COL_F��־ Then
                    If Val(.TextMatrix(lngRow, COL_��־)) = 1 Then
                        strPrompt = "����ҽ��"
                    ElseIf Val(.TextMatrix(lngRow, COL_��־)) = 2 Then
                        strPrompt = "��¼ҽ��"
                    End If
                    '����п�����ҩ�����Ϣ��������ʾ
                    If Val(.TextMatrix(lngRow, COL_ҽ��״̬)) = 1 Then
                        Select Case Val(.TextMatrix(lngRow, COL_���״̬))
                        Case 1
                            If .TextMatrix(lngRow, COL_�������) = "K" And Val(.TextMatrix(lngRow, COL_��鷽��)) = 1 Then '��Ѫҽ�����
                                strPrompt = "��Ѫҽ�����˶�"
                            Else
                                strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "F", "����", "K", "��Ѫ", "������ҩ") & "�����"
                            End If
                        Case 2
                            If Not (.TextMatrix(lngRow, COL_�������) = "K" And Val(.TextMatrix(lngRow, COL_��鷽��)) = 1) Then
                                strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "K", "��Ѫ", "������ҩ") & "���ͨ��"
                            End If
                        Case 3
                            strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "K", "��Ѫ", "������ҩ") & "���δͨ��:" & GetKSSAuditQuestion(Val(.TextMatrix(lngRow, COL_ID)))
                        Case 4
                            If gblnѪ��ϵͳ = False Then strPrompt = "��Ѫ��Ѫ�����"
                        Case 5
                            If gblnѪ��ϵͳ = False Then strPrompt = "��ѪѪ��������Ѫ"
                        Case 7
                            strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "F", "����", "K", "��Ѫ", "������ҩ") & "��ǩ��"
                        End Select
                    End If
                ElseIf .MouseCol = COL_����״̬ Then
                    If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then strPrompt = "����δ��"
                    If Val(.TextMatrix(lngRow, COL_����ID)) <> 0 Or .TextMatrix(lngRow, COL_��鱨��ID) <> "" Or _
                        Val(.TextMatrix(lngRow, COL_RIS����ID)) <> 0 Or Val(.TextMatrix(lngRow, COL_LIS����ID)) <> 0 Then
                        
                        If Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 0 Then
                            strPrompt = "����δ�ģ�����鿴"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 1 Then
                            strPrompt = "�������ģ�����鿴"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 2 Then
                            strPrompt = "���沿�����ģ�����鿴"
                        End If
                    End If
                ElseIf .MouseCol = COL_F���� Then
                    strPrompt = GetAdviceReportTip(lngRow)
                End If
            End If
     
            If .MouseRow > -1 And .MouseCol > -1 And (mvarCond.����ģʽ = 3 And .MouseCol = COL_����״̬ Or .MouseCol = COL_������ӡ Or .MouseCol = COL_����Ԥ��) Then
                If .Cell(flexcpFontUnderline, .MouseRow, .MouseCol) = True And .TextMatrix(.MouseRow, .MouseCol) <> "" Then
                    .MousePointer = 99
                Else
                    .MousePointer = 0
                End If
            Else
                .MousePointer = 0
            End If
                        
            If strPrompt <> "" Then
                Call zlCommFun.ShowTipInfo(.hwnd, strPrompt)
                mlngPromptRow = lngRow
            ElseIf mlngPromptRow <> 0 And lngRow <> mlngPromptRow Then
            '����֮ǰ����ʾ����
                Call zlCommFun.ShowTipInfo(.hwnd, "")
                mlngPromptRow = 0
            End If
        End If
    End With
End Sub


Private Sub vsfAdivceDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMore.Tag = ""
    fraMore.Visible = False
    PicAdviceDetail.Visible = False
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicAdviceDetail.Visible = False And vsAdvice.MouseRow > 0 Then
        Call LoadAdviceDetail(vsAdvice.MouseRow)
    End If
End Sub

Private Sub LoadAdviceDetail(lngRow As Long)
'���ܣ���ʾĳ��ҽ������ϸ����
    Dim i As Long, j As Long
    
    vsfAdivceDetail.Redraw = flexRDNone
    vsfAdivceDetail.Clear
    vsfAdivceDetail.Rows = vsfAdivceDetail.FixedRows
    vsfAdivceDetail.Cols = 2
    j = 0
    With vsAdvice
        For i = 0 To .Cols - 1
             If .Cell(flexcpData, 0, i) = "Detail" Then
                j = j + 1
                vsfAdivceDetail.Rows = vsfAdivceDetail.FixedRows + j
                vsfAdivceDetail.TextMatrix(j - 1, 0) = .TextMatrix(0, i) & "��"
                vsfAdivceDetail.TextMatrix(j - 1, 1) = .TextMatrix(lngRow, i)
                
                vsfAdivceDetail.Col = 0: vsfAdivceDetail.Row = j - 1
                vsfAdivceDetail.CellForeColor = &H8000000C
             End If
        Next
    End With
    With vsfAdivceDetail
        If .Rows > 0 Then
            .AutoSize 0, 1
            .Height = IIF(.RowHeight(0) < .RowHeightMin, .RowHeightMin, .RowHeight(0)) * .Rows + 100
            .Width = .ColWidth(0) + .ColWidth(1)
            .Row = -1
            
            PicAdviceDetail.Height = .Height
            PicAdviceDetail.Width = .Width
            PicAdviceDetail.Left = fraMore.Left + fraMore.Width
            
            If PicAdviceDetail.Height + fraMore.Top + fraMore.Height > Me.Top + Me.Height Then
                PicAdviceDetail.Top = fraMore.Top + fraMore.Height - PicAdviceDetail.Height - 10
            Else
                PicAdviceDetail.Top = fraMore.Top - 10  '���ⶥ�˺ͱ�����غ�
            End If
            
            Call SetPicAdviceDetailEffect
            If PicAdviceDetail.Visible = False Then PicAdviceDetail.Visible = True
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetPicAdviceDetailEffect()
    Dim lngR As Long
    
    '�߿�API=RoundRect
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, 0)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (0, Screen.TwipsPerPixelY)-(0, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (PicAdviceDetail.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
           
    '��״
    lngR = CreateRoundRectRgn(0, 0, PicAdviceDetail.ScaleX(PicAdviceDetail.Width, PicAdviceDetail.ScaleMode, vbPixels) + 1, PicAdviceDetail.ScaleY(PicAdviceDetail.Height, PicAdviceDetail.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(PicAdviceDetail.hwnd, lngR, False)
    
End Sub

Private Sub vsfAdivceDetail_LostFocus()
    PicAdviceDetail.Visible = False
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or tbcAppend.Height - Y < 500 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        tbcAppend.Top = tbcAppend.Top + Y
        tbcAppend.Height = tbcAppend.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsColumn
            If .Visible Then
                .Visible = False
                vsAdvice.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsAdvice.ColHidden(.RowData(i)) Or vsAdvice.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                    '���ô����ŵ��е���ʾ��ʽ
                    If .TextMatrix(i, 1) = "������" Or .TextMatrix(i, 1) = "��ӡ" Or .TextMatrix(i, 1) = "Ԥ��" Then
                        If mvarCond.ҽ�� = 1 Then
                            .TextMatrix(i, 0) = 1
                        Else
                            .TextMatrix(i, 0) = 0
                        End If
                        .Cell(flexcpForeColor, i, 0, i, 1) = .BackColorFixed
                    End If
                Next
                
                vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 150
                If vsColumn.Top + vsColumn.Height > Me.ScaleHeight Then
                    vsColumn.Height = Me.ScaleHeight - vsColumn.Top
                    vsColumn.Width = 1750
                Else
                    vsColumn.Width = 1470
                End If
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Function CheckWindow() As Boolean
'���ܣ����ҽ���༭�����Ƿ��Ѿ���
    If Not mfrmEdit Is Nothing Then
        '��ǰ���ڴ���
        MsgBox "ҽ���༭�����Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        '��λ����ǰ�Ĵ���
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        mfrmEdit.SetFocus
        Exit Function
    Else
        '�������ڴ���
        If Not CheckAdviceWindow("����ҽ���༭") Then Exit Function
    End If
    CheckWindow = True
End Function

Private Sub FuncBillPrint(Optional objControl As CommandBarControl, Optional ByVal strPar As String, Optional strName As String)
'���ܣ���ӡ���Ƶ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strNO As String, lng��¼���� As Long

    Dim lng���ID As Long
    Dim strParameter As String
    Dim strErr As String
    Dim blnDo As Boolean
    Dim strBillName As String '���Ƶ��ݵ�����  �����ļ��б�.����
    
    If Not objControl Is Nothing Then strPar = objControl.Parameter: strName = objControl.Caption
    If strPar = "" Then Exit Sub

    If InStr(strPar, "|") > 0 Then strParameter = Split(strPar, "|")(0): strNO = Split(strPar, "|")(1)
    
    strBillName = strName
    strBillName = Replace("<Tab>" & strBillName, "<Tab>��ӡ:", "")
    If InStr(strBillName, "(&") > 0 Then
        strBillName = Mid(strBillName, 1, InStr(strBillName, "(&") - 1)
    End If
    
    With vsAdvice
        '��ӡ������ʾ
        On Error GoTo errH
        lng���ID = Decode(Val(.TextMatrix(.Row, COL_���ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_���ID)))
        If .TextMatrix(.Row, COL_�������) = "E" And Val(.TextMatrix(.Row, COL_��������)) = 6 Then
            If Not gobjLIS Is Nothing Then '��ӡ�������뵥��
                blnDo = gobjLIS.CheckAcceptance(CStr(lng���ID), strErr)
                If Not blnDo Then
                   MsgBox "�ñ걾�Ѿ�������ƺ��գ����ܴ�ӡ:" & strBillName & "��", vbInformation, gstrSysName
                   Exit Sub
                End If
            End If
        End If
        If strNO <> "" Then
            strSQL = "Select A.NO,A.��¼���� from ����ҽ������ A,����ҽ����¼ B Where a.ҽ��ID=b.id And a.NO=[2] And (b.ID=[3] Or b.���ID=[3])"
        Else
            strSQL = "Select NO,��¼���� from ����ҽ������ Where ҽ��ID=[1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, COL_ID)), strNO, lng���ID)
        If rsTmp.RecordCount > 0 Then
            strNO = rsTmp!NO & ""
            lng��¼���� = Val(rsTmp!��¼���� & "")
            strSQL = "Select ��ӡ��,��ӡʱ�� From ���Ƶ��ݴ�ӡ Where NO=[1] And ��¼����=[2] And ��ӡ����=1 Order by ��ӡʱ�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strNO, lng��¼����)
            If Not mbln����Ԥ�� Then
                If Not rsTmp.EOF Then
                    If MsgBox("��[" & strBillName & "]�Ѿ���ӡ�� " & rsTmp.RecordCount & " �Σ����һ����""" & _
                        rsTmp!��ӡ�� & """��""" & Format(rsTmp!��ӡʱ��, "yyyy-MM-dd HH:mm") & """��ӡ��" & vbCrLf & vbCrLf & "Ҫ������ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            On Error GoTo 0
            SwitchPrintSet glngSys & "\" & p����ҽ���´�
            '���ô�ӡ
            If mobjReport.ReportPrintSet(gcnOracle, glngSys, strParameter, mfrmParent) Then
                mstrBillPrint = strParameter & "," & strNO & "," & lng��¼����
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strParameter, mfrmParent, "NO=" & strNO, "����=" & lng��¼����, IIF(mbln����Ԥ��, 1, 2))
                mstrBillPrint = ""
            End If
            SwitchPrintSet glngSys & "\" & p����ҽ���´�, True
        End If
    End With
    mbln����Ԥ�� = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSign()
'���ܣ���ҽ�����е���ǩ��
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lngǩ��id As Long, lng֤��ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim ColIDs As Collection, ColSource As Collection
    
    If Not mblnEditable Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡǩ��ҽ��Դ��
    intRule = ReadAdviceSignSource(1, mlng����ID, mstr�Һŵ�, strIDs, 0, mblnMoved, strSource, mstrǰ��IDs, , ColIDs, ColSource)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "�ò���Ŀǰû�п���ǩ����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    For i = 1 To ColIDs.Count
        strSign = gobjESign.Signature(ColSource(i), gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
            strSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & ColIDs(i) & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
        End If
    Next
    If strSign <> "" Then
        Call LoadAdvice 'ˢ�½���
        MsgBox "����ɵ���ǩ����", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignErase()
'���ܣ�ȡ��ҽ���ĵ���ǩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If Not mblnEditable Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "ǩ��" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����ǩ������ȡ��
        If .Cell(flexcpData, .Row, 0) = 4 Then
            MsgBox "����ҽ����ǩ������ȡ����", vbInformation, gstrSysName
            Exit Sub
        End If
        '�¿�ǩ�����������¿�״̬
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 1 Then
                MsgBox "����ҽ���Ѿ����ͻ����ϣ���ǩ������ȡ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        'ֻ��ȡ������ǩ����
        If .TextMatrix(.Row, 2) <> UserInfo.���� Then
            MsgBox "��ǩ���˲����㱾�ˣ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪȡ�����ǩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_ҽ��ǩ����¼_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice 'ˢ�½���
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'���ܣ�У��ҽ���ĵ���ǩ��(�ɶ���ת�Ƶ�����)
    Dim strSource As String
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "ǩ��" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ȡǩ��ҽ��Դ��
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '��֤ǩ��
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub

Private Sub FuncAdviceAdd()
'���ܣ�����ҽ��
    If Not CheckWindow Then Exit Sub
        '���ҺŲ����Ƿ���
    If Not FuncTimeLimitCheck Then Exit Sub
    
    If Not FuncPathAdd() Then Exit Sub
    
    Set mfrmEdit = frmOutAdviceEdit
    Call frmOutAdviceEdit.ShowMe(mfrmParent, mint����, mMainPrivs, mlng����ID, mstr�Һŵ�, mlngǰ��ID, , , mblnModalNew, mlng�������ID, mstrǰ��IDs, mclsMipModule, mlng�Һſ���ID, mblnMoved, , mlngΣ��ֵID)
End Sub

Private Sub FuncAdviceDel()
'ɾ����ɾ����ǰҽ��
'˵������������ɾ��,�Լ�����,�������,��ҩ�䷽,������ɾ��,һ����ҩֻɾ����ǰҩƷ
    Dim strSQL As String, lngҽ��ID As Long
    Dim blnGroup As Boolean, i As Long, blnBat As Boolean, blnTrans As Boolean, arrSQL As Variant
    Dim lngRow As Long, strXML As String, lng������� As Long
    Dim strDelIDs As String, arrDelID() As String
    Dim strDelDrugIDs As String         '��¼ɾ����ҩƷҽ��,���ڴ��������ҩ���
    Dim lng��ID As Long
    Dim blnRISԤԼ As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim bln��������� As Boolean
    Dim bln��Ѫ As Boolean, strErr As String
    
    If Not mblnEditable Then Exit Sub
    
    '�����Һŵ���Ч��������Ϊ�����Һŵ���Ч�����Ĳ��ˣ�����ɾ��δ���͵�ҽ���������ɽ��
    
    With vsAdvice
        '����Ƿ����ɾ��
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ������ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(",5,6,", "," & .TextMatrix(.Row, COL_�������) & ",") > 0 Then
            strDelDrugIDs = "����ҩ��" & lngҽ��ID
        ElseIf .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "4" Then
            strDelDrugIDs = "����ҩ��" & .Cell(flexcpData, .Row, COL_���ID)
        End If
        'ҽ���´��ҽ��
        If mint���� = 2 Then
            If InStr("," & mstrǰ��IDs & ",", "," & .TextMatrix(.Row, COL_ǰ��ID) & ",") = 0 Then
                MsgBox "��ҽ����Ϊ��ǰҽ�������´����ɾ����ҽ����", vbInformation, gstrSysName
                Exit Sub
            ElseIf Val(.TextMatrix(.Row, COL_ǰ��ID)) = 0 Then
                MsgBox "��ҽ������ҽ�������´����ɾ����ҽ����", vbInformation, gstrSysName
                Exit Sub
            End If
        ElseIf Val(.TextMatrix(.Row, COL_ǰ��ID)) <> 0 Then
            MsgBox "��ҽ��Ϊҽ�������´����ɾ����ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ����ͻ����ϣ�����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ������ɾ��
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ��������ɾ��������ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mlng·��״̬ = 1 Then
            If CheckPathAdviceIsExeOut(lngҽ��ID) Then
                MsgBox "��ҽ����Ӧ����Ŀ�Ѿ�ִ�С�" & vbCrLf & "��ȡ��ִ�еǼǺ��ٽ���ɾ��������", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        '����Ѫ��ϵͳ��Ѫҽ��ɾ�����ƣ�����Ѫ����˽׶ε��¿�ҽ������ɾ
        bln��Ѫ = gblnѪ��ϵͳ And .TextMatrix(.Row, COL_�������) = "K"
        If gblnѪ��ϵͳ And .TextMatrix(.Row, COL_�������) = "K" And InStr("5,2", Val(.TextMatrix(.Row, COL_���״̬))) > 0 Then
            MsgBox "����Ѫҽ���ѱ�Ѫ�����" & IIF(Val(.TextMatrix(.Row, COL_���״̬)) = 5, "������Ѫ", "�����������Ѫ") & "������ɾ��������ɾ��������Ѫ����ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        arrSQL = Array()
        
        If InStr(",5,6,", .TextMatrix(.Row, COL_�������)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If blnGroup Then
                lng��ID = Val(.TextMatrix(.Row, COL_���ID))
                If MsgBox("ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """������ҩƷһ����ҩ,ȷʵҪɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            If mblnPass Then
                Call gobjPass.zlPassAdviceDel(mobjPassMap, lngҽ��ID, zlDatabase.Currentdate)
            End If
        
        ElseIf .TextMatrix(.Row, COL_�������) <> "" Then
            If .TextMatrix(.Row, COL_�������) = "K" Then
                If MsgBox("ȷʵҪȡ����Ѫ����""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                If MsgBox("ȷʵҪȡ����������""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("Ҫ��""" & .TextMatrix(.Row, col_ҽ������) & """ͬʱ�����������Ŀһ��ȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnBat = True
                End If
            End If
        Else
            If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        If .TextMatrix(.Row, COL_�������) = "D" Then
            If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
                blnRISԤԼ = True
            End If
        End If
        
        Call CreatePlugInOK(p����ҽ���´�, mint����)
        If blnBat Then
            For i = 1 To .Rows - 1
                lng������� = Val(.TextMatrix(.Row, COL_�������))
                If .TextMatrix(i, COL_ҽ��״̬) = "1" And Val(.TextMatrix(i, COL_�������)) = lng������� Then
                    '����ɾ��ǰ��ҽӿ�
                    On Error Resume Next
                    If Not gobjPlugIn Is Nothing Then
                        If gobjPlugIn.AdviceDeletBefor(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(i, COL_ID)), mint����) = False Then
                            If err.Number = 0 Then Exit Sub
                        End If
                        Call zlPlugInErrH(err, "AdviceDeletBefor")
                    End If
                                        If Not CheckDelAdivceOfPathItem(Val(.TextMatrix(.Row, COL_ID))) Then Exit Sub
                    If err.Number <> 0 Then err.Clear
                    On Error GoTo 0
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & .TextMatrix(i, COL_ID) & ",1)"
                    strDelIDs = strDelIDs & "," & .TextMatrix(i, COL_ID)
                End If
            Next
        Else
            '����ɾ��ǰ��ҽӿ�
            On Error Resume Next
            If Not gobjPlugIn Is Nothing Then
                If gobjPlugIn.AdviceDeletBefor(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, lngҽ��ID, mint����) = False Then
                    If err.Number = 0 Then Exit Sub
                End If
                Call zlPlugInErrH(err, "AdviceDeletBefor")
            End If
                        If Not CheckDelAdivceOfPathItem(Val(.TextMatrix(.Row, COL_ID))) Then Exit Sub
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If InStr(",5,6,E,", .TextMatrix(.Row, COL_�������)) > 0 Then
                If Val(.TextMatrix(.Row, COL_�������״̬)) = 1 Or Val(.TextMatrix(.Row, COL_�������״̬)) = 2 Then
                    bln��������� = True
                Else
                    '��ҩ�䷽�ж�
                    If .TextMatrix(.Row, COL_��������) = "4" Then
                        For i = .Row To 1 Step -1
                            If .TextMatrix(i, COL_�������) = "7" And Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_ID)) Then
                                If Val(.TextMatrix(i, COL_�������״̬)) = 1 Or Val(.TextMatrix(i, COL_�������״̬)) = 2 Then
                                    bln��������� = True
                                End If
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
            If bln��������� Then
                arrSQL(UBound(arrSQL)) = "Zl_����ҽ����¼_�������ɾ��(" & IIF(Val(.TextMatrix(.Row, COL_���ID)) = 0, lngҽ��ID, Val(.TextMatrix(.Row, COL_���ID))) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & lngҽ��ID & ",1)"
            End If
            strDelIDs = strDelIDs & "," & lngҽ��ID
        End If
        strDelIDs = Mid(strDelIDs, 2)
    End With
    
    If blnRISԤԼ Then
        Set rsTmp = GetDataRISԤԼ(strDelIDs)
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!ԤԼid & "")) Then 'ɾ��ҽ��
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISSchedulingEx)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            rsTmp.MoveNext
        Next
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    If bln��Ѫ = True Then
        If InitObjBlood(True) = True Then
            If gobjPublicBlood.AdviceOperation(pסԺҽ���´�, lngҽ��ID, 2, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "Ѫ�⹫����������ʧ�ܣ���ϸ��Ϣ��" & strErr, vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "Ѫ�⹫����������ʧ�ܣ����飡", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    With vsAdvice
        '������ֱ��ɾ��
        .Redraw = False
        
        'ɾ��һ����ҩ��һ��ʱ����ʾ����
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_���ID)) = Val(.TextMatrix(.Row + 1, COL_���ID)) Then
                If .TextMatrix(.Row, COL_��ʼʱ��) <> "" And .TextMatrix(.Row + 1, COL_��ʼʱ��) = "" Then
                    .TextMatrix(.Row + 1, COL_��ʼʱ��) = .TextMatrix(.Row, COL_��ʼʱ��)
                    .TextMatrix(.Row + 1, COL_Ƶ��) = .TextMatrix(.Row, COL_Ƶ��)
                    .TextMatrix(.Row + 1, COL_�÷�) = .TextMatrix(.Row, COL_�÷�)
                End If
            End If
        End If
                
        lngRow = .Row
        If blnBat Then
            For i = .Rows - 1 To 1 Step -1
                If .TextMatrix(i, COL_ҽ��״̬) = "1" And Val(.TextMatrix(i, COL_�������)) = lng������� Then
                    .RemoveItem i
                End If
            Next
        Else
            .RemoveItem .Row
        End If
        
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        If lng��ID <> 0 Then
            i = .FindRow(CStr(lng��ID), , COL_���ID)
            If i <> -1 Then
                .TextMatrix(i, COL_��) = ""
                Call SetTagһ����ҩ(i)
            End If
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = True
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col) '��ɫ���������
        
        '����ɾ������ҽӿ�
        On Error Resume Next
        arrDelID = Split(strDelIDs, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                If Not gobjPlugIn Is Nothing Then
                    Call gobjPlugIn.AdviceDeleted(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(arrDelID(i)), mint����)
                    Call zlPlugInErrH(err, "AdviceDeleted")
                End If
            End If
        Next
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
        'PASSҽ��ɾ�����Զ�������鹦��
        If mblnPass Then
            Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 4, strDelDrugIDs)
        End If
    End With
    Call ShowTotalMoney
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDelAdivceOfPathItem(ByVal lngҽ��ID As Long) As Boolean
'���ܣ����ҽ����Ӧ��·����Ŀ�Ƿ�����ɾ��������Ǳ���ִ�е���Ŀ����Ӧ��ҽ��������Ҫ����ԭ��ѡ�񲢸��±���ԭ��
'       ��ӹ�����ԭ��Ĳ������
'���أ�True-����ɾ����ҽ����false-����ɾ��
'����:lngҽ��ID
    Dim blnCancel As Boolean, blnMust As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, rsAdvice As ADODB.Recordset
    Dim strReason As String
    Dim vPoint As PointAPI
    Dim strTemp As String
    Dim arrTmp As Variant
    Dim arrSQL As Variant
    Dim i As Long

    '1.���·����Ŀ
    strSQL = "Select  c.Id as ִ��Id, c.����,c.����ԭ��,d.ִ�з�ʽ,c.����,c.�׶�ID,c.·����¼ID,c.��ĿID " & _
             " From ��������·��ҽ�� B, ��������·��ִ�� C, ����·����Ŀ D" & vbNewLine & _
             "Where b.����ҽ��Id=[1] And b.·��ִ��id = c.Id And d.Id = c.��Ŀid And d.ִ�з�ʽ in (1,2,4)"

    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���·��ҽ��", lngҽ��ID)

    If rsTmp.RecordCount < 1 Then
        CheckDelAdivceOfPathItem = True
        Exit Function    '�� �������ɵ�·��ҽ��
    End If
    '2.���ҽ���ܷ�ɾ��
    '��·����Ŀ������У�Ե�δ���ϵ�����ҽ������ʾ����ֹɾ��    ҽ��״̬ ��3-��У��
    strSQL = "Select a.����ҽ��ID,b.ҽ��״̬ " & vbNewLine & _
             "From ��������·��ҽ�� A, ����ҽ����¼ B" & vbNewLine & _
             "Where a.·��ִ��id = [1] And a.����ҽ��id = b.Id  And b.ҽ��״̬>1 and b.ҽ��״̬<>4"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "���·��ҽ��", rsTmp!ִ��Id)

    If rsAdvice.RecordCount > 0 Then
        MsgBox "ɾ��ҽ�����ڵ�·����Ŀ�д����ѷ��͵�δ���ϵ�ҽ�����������ϸ�ҽ������ִ�д˲�����", vbInformation, gstrSysName
        CheckDelAdivceOfPathItem = False
        Exit Function
    End If
    

    
    '����ִ�з�ʽ �����Ƿ��б�Ҫ��ӱ���ԭ��
    blnMust = CheckPathItemIsMust(Val(rsTmp!ִ�з�ʽ & ""), Val("" & rsTmp!����), Val("" & rsTmp!·����¼id), Val("" & rsTmp!�׶�id), Val("" & rsTmp!��ĿID), 1)
    If Not blnMust Then CheckDelAdivceOfPathItem = True: Exit Function
    
    '----------------------------
    '3.�������ɵ���Ŀ��д����ԭ��
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ԭ�� & "" = "" Then
            strTemp = strTemp & rsTmp!ִ��Id & "," & rsTmp!���� & ";"
        End If
        rsTmp.MoveNext
    Next
    
    If strTemp = "" Then
        CheckDelAdivceOfPathItem = True
        Exit Function
    Else
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    End If

    strSQL = "Select b.���� as ����,a.���� as ID,a.����,a.����,a.���� From ������쳣��ԭ�� a,������쳣��ԭ�� b" & _
             " Where a.����=1 And a.ĩ��=1 And a.�ϼ�=b.���� And b.ĩ��=0 " & _
             " Order by ����,a.����"
    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)

    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "������쳣��ԭ��", True, , , True, True, True, _
                                      vPoint.X, vPoint.Y, vsAdvice.RowHeight(vsAdvice.Row), blnCancel, False, True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "ϵͳû�г�ʼ������쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
        Exit Function
    Else
        strReason = rsTmp!ID
    End If

    If strReason <> "" Then
        arrSQL = Array()
        For i = 0 To UBound(Split(strTemp, ";"))
            arrTmp = Split(Split(strTemp, ";")(i), ",")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_��������·������_Update(" & arrTmp(0) & ",'" & arrTmp(1) & "',Null ,Null,Null,Null,'" & strReason & "')"
        Next
        '�����������������ԭ�����ʧ�ܣ�ҽ������ɾ�����ٴ�ɾ��ʱ����������ӱ���ԭ���ſ�ɾ����
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        CheckDelAdivceOfPathItem = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncAdviceRevoke()
'ɾ������ǰҽ������(һ��ҽ������)
    Dim strSQL As String, lngҽ��ID As Long
       
    
    Dim strNO As String
    Dim lngType As Long
    Dim rsTmp As ADODB.Recordset
    
    If Not mblnCanRevoke Then Exit Sub
    
    With vsAdvice
        '����Ƿ��������
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        
        If RowInһ����ҩ(.Row, 0, 0) Then
            lngType = 1
        End If
        
        If .TextMatrix(.Row, COL_�������) = "E" Then
            If RowIs������(.Row) Then
                lngType = 2
            End If
        End If
        
        If .TextMatrix(.Row, COL_�������) = "5" Or .TextMatrix(.Row, COL_�������) = "6" Then
            strNO = .Cell(flexcpData, .Row, COL_������)
        End If
        
        If RevokeOutAdvice(mlng����ID, mlng�Һ�ID, mstr�Һŵ�, mstr����, mstr�����, mlng�Һſ���ID, lngҽ��ID, Val(.TextMatrix(.Row, COL_ҽ��״̬)), .TextMatrix(.Row, COL_�������), .TextMatrix(.Row, COL_��������), Val(.TextMatrix(.Row, COL_���״̬)), _
            .Cell(flexcpData, .Row, col_����ʱ��), Val(.TextMatrix(.Row, COL_ǩ����)), lngType, .TextMatrix(.Row, col_ҽ������), mblnMoved, mclsMipModule, mint����) = False Then Exit Sub
        
    End With
    
    Call LoadAdvice 'ˢ�½���
    Call ShowTotalMoney
    
    'PASSҽ�����Ϻ��Զ�������鹦��
    If mblnPass Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 3)
    End If
    
    'ҩƷҽ���÷Ϻ��ж��ǲ���Ҫ�ش�
    If strNO <> "" Then
        strSQL = "Select Distinct D.���,D.����,D.˵��,B.NO,B.��¼����" & _
            " From ����ҽ����¼ A,����ҽ������ B,��������Ӧ�� C,�����ļ��б� D" & _
            " Where C.������ĿID = A.������ĿID And a.ID=b.ҽ��ID " & _
            " And C.Ӧ�ó���=1 And C.�����ļ�ID=D.ID And D.����=7 And b.NO=[1] and a.������� in ('5','6') and a.�Һŵ�||''=[2]" & _
            " Order by D.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strNO, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            If MsgBox("�����ϵ�ҩƷ����ǩ�Ѿ���ӡ���Ƿ��ش�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '���ô�ӡ
                SwitchPrintSet glngSys & "\" & p����ҽ���´�
                If mobjReport.ReportPrintSet(gcnOracle, glngSys, "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1", mfrmParent) Then
                    mstrBillPrint = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" & "," & rsTmp!NO & "," & rsTmp!��¼����
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1", mfrmParent, "NO=" & rsTmp!NO, "����=" & rsTmp!��¼����, 2)
                    mstrBillPrint = ""
                End If
                SwitchPrintSet glngSys & "\" & p����ҽ���´�, True
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRevokeͣ��()
'ɾ������ǰҽ������(һ��ҽ������)
    Dim strSQL As String, lngҽ��ID As Long
    Dim lng֤��ID As Long, lngǩ��id As Long
    Dim strSign As String, intRule As Integer
    Dim strSource As String, strIDs As String, blnDo As Boolean
    Dim strTimeStamp As String, blnTran As Boolean, strErr As String, strTimeStampCode As String
    
    If Not mblnCanRevoke Then Exit Sub
    
    With vsAdvice
        '����Ƿ��������
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ���������ϡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 8 Then
            MsgBox "��ǰѡ���ҽ����δ���ͻ��Ѿ����ϡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '92129:ҽ���ѱ���Ѫ�ƽ������ܽ�������
        If .TextMatrix(.Row, COL_�������) = "K" And gblnѪ��ϵͳ And InStr(1, ",2,5,6,", "," & Val(.TextMatrix(.Row, COL_���״̬)) & ",") <> 0 Then
            MsgBox "�������ϵ���Ѫҽ��" & IIF(Val(.TextMatrix(.Row, COL_���״̬)) = 2, "�Ѿ������Ѫ", "����������Ѫ�׶�") & "������ֱ������ҽ������Ҫ����������Ѫ����ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '���з���ת������������
        If zlDatabase.DateMoved(.Cell(flexcpData, .Row, col_����ʱ��)) Then
            If MovedBySend(lngҽ��ID, 0, 1) Then
                MsgBox "��ҽ���ķ����Ѿ�ȫ���򲿷�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '����ǩ��������ʾ
        If Val(.TextMatrix(.Row, COL_ǩ����)) = "1" Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
                Else
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            If gobjESign.CertificateStoped(UserInfo.����) = False Then strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
        End If
        
        '�������ҽ����Ӧ�ķ��ý������
        If Not CheckAdviceBalanceRevoke(lngҽ��ID) Then Exit Sub
        
        '����˼��ʷ��ü��
        If InStr(GetInsidePrivs(p����ҽ���´�), "��������˼���ҽ��") = 0 Then
            If Not CheckAdviceBillingRevoke(lngҽ��ID) Then
                MsgBox "Ҫ����ҽ���Ķ�Ӧ���ʻ��۷����Ѿ���ˣ��������ϡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If RowInһ����ҩ(.Row, 0, 0) Then
            If MsgBox("����һ����ҩ��ҽ������һ�����ϣ�ȷʵҪ������" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("ȷʵҪ����ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_����ҽ����¼_����(" & lngҽ��ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        
        '����ʱ���е���ǩ��
        If strSign <> "" Then
            If gobjESign.CertificateStoped(UserInfo.����) = False Then
                '��ȡǩ��ҽ��Դ��
                strIDs = lngҽ��ID
                intRule = ReadAdviceSignSource(4, mlng����ID, mstr�Һŵ�, strIDs, 0, mblnMoved, strSource)
                If intRule = 0 Then Exit Sub
                If strSource = "" Then
                    MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
                If strSign <> "" Then
                    If strTimeStamp <> "" Then
                        strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        strTimeStamp = "NULL"
                    End If
                    lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                    strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
    
    'RIS���
    If HaveRIS Then
        If 1 <> gobjRis.ReqInteractive(5, "AppNO", lngҽ��ID) Then
            Exit Sub
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            Call ZLHIS_CIS_024(mclsMipModule, mlng����ID, mstr����, , mstr�����, 1, mlng�Һ�ID, mlng�Һſ���ID, "", lngҽ��ID, _
                vsAdvice.TextMatrix(vsAdvice.Row, COL_�������), vsAdvice.TextMatrix(vsAdvice.Row, COL_��������))
        End If
    End If
    '�������Ϻ���ҽӿ�
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.AdviceRevoked(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, lngҽ��ID, mint����)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    Call InitObjLis(p����ҽ��վ)
    '����LIS�������뵥
    If Not gobjLIS Is Nothing Then
        If gobjLIS.DelLisApplicationForm(CStr(lngҽ��ID), strErr) = False Then
            MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
        End If
    End If
    '�������ݽ���ƽ̨����LIS,PACSȡ�����뵥
    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjExchange Is Nothing Then
        With vsAdvice
            If .TextMatrix(.Row, COL_�������) = "D" Then
                blnDo = True
            ElseIf .TextMatrix(.Row, COL_�������) = "E" Then
                blnDo = RowIs������(.Row)
            End If
            If blnDo Then
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_�������) = "D", 2, 1), "����ID::" & mlng����ID & "||��ҳID::0||ҽ��ID::" & lngҽ��ID & "||��������::0")
            End If
        End With
    End If
    
    Call LoadAdvice 'ˢ�½���
    Call ShowTotalMoney
    
    'PASSҽ�����Ϻ��Զ�������鹦��
    If mblnPass Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 3)
    End If
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceModi()
'���ܣ��޸ĵ�ǰҽ��
    Dim lngҽ��ID As Long
    
    If Not CheckWindow Then Exit Sub
        '���ҺŲ����Ƿ���
    If Not FuncTimeLimitCheck Then Exit Sub
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then Exit Sub
        
        'ҽ���´��ҽ��
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "�����޸ĸ�ҽ��,��ҽ���Ǹ���������ҽ�������ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��У�Ի��ѷ�ֹ
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ����ͻ����ϣ������޸ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ�������޸�
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ���������޸ġ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set mfrmEdit = frmOutAdviceEdit
        Call frmOutAdviceEdit.ShowMe(mfrmParent, mint����, mMainPrivs, mlng����ID, mstr�Һŵ�, mlngǰ��ID, _
            Val(.TextMatrix(.Row, COL_Ӥ��ID)), lngҽ��ID, , mlng�������ID, mstrǰ��IDs, mclsMipModule, mlng�Һſ���ID, mblnMoved, mint��������)
    End With
End Sub

Private Sub FuncAdviceTest()
'���ܣ���дƤ�Խ��
    Dim strSQL As String, str��� As String
    Dim int��� As Integer, strLabel As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTran As Boolean
    Dim dateInput As Date
    Dim strSelect As String, i As Long
    Dim strSelectInput As String
    Dim strTextInput As String
    
    If mlng����ID = 0 Then Exit Sub
    If Not mblnEditable Then Exit Sub
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Exit Sub
    If Not (vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "1") Then
        MsgBox "��ǰҽ�����ݲ��ǹ���������Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) <> 0 Then
        MsgBox "�㲻�ܸ��ù���������д�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 4 Then
        MsgBox "�ù�������ҽ���Ѿ����ϣ�������д�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 1 Then
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) = "����" Then
            If MsgBox("�ù�������ҽ���Ѿ����Ϊ���ԣ�Ҫ������Ա����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strLabel = ""
        Else
            If MsgBox("�ù�������ҽ����δ���ͣ���������д������������" & vbCrLf & vbCrLf & _
                "�����Ա��Ϊ���ԣ�ͬʱ��ҽ�������ᷢ�͡�Ҫ���Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strLabel = "����"
        End If
        int��� = -1 '�������ֳ���
        strSQL = "ZL_����ҽ����¼_Ƥ��(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & strLabel & "',NULL)"
    Else
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) <> "" Then
            If MsgBox("�ù�������ҽ���Ѿ���д�˽����Ҫ������д��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            '����Ӧ��ҽ���Ƿ��Ѿ�����
            If mblnƤ������ Then
                If AdviceSended(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), CDate(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ʱ��))) Then
                    MsgBox "��Ƥ�Զ�Ӧ��ҩƷ�Ѿ����ͣ������ٸ���Ƥ�Խ����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        strSQL = "Select Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
        '����
        For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(0), ","))
            strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(0), ",")(i) & "|0"
        Next
        '����
        For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(1), ","))
            strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(1), ",")(i) & "|0|2"
        Next
        strSelect = Mid(strSelect, 2)
        str��� = zlCommFun.ShowMsgBox("Ƥ�Խ��", vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", _
            "ȷ��(&O),?ȡ��(&C)", Me, vbQuestion, "Ƥ��ʱ��", dateInput, "yyyy-MM-dd HH:mm", "Ƥ�Խ��(&P):" & strSelect, strSelectInput, _
            "������Ӧ(&F)", 50, strTextInput, , True)
        If str��� = "" Then Exit Sub
        If strSelectInput = "" Then Exit Sub
        If Format(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") > dateInput Then
            MsgBox "Ƥ��ʱ�䲻����ҽ����Чʱ����ǰ��������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
        Call GetTestLabel(rsTmp!�걾��λ, strSelectInput, strLabel, int���)
        strSQL = "ZL_����ҽ����¼_Ƥ��(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & strLabel & "'," & int��� & _
                ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
    End If
        
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    
    vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) = strLabel
    If mvarCond.��ʾģʽ = 0 Then
        '����Ǽ��ģʽ��������ҩƷ���Ƥ�Խ����
        If InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(+)") > 0 Or InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(-)") > 0 Then
            vsAdvice.TextMatrix(vsAdvice.Row, col_����) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(+)", strLabel)
            vsAdvice.TextMatrix(vsAdvice.Row, col_����) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(-)", strLabel)
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, col_����) = vsAdvice.TextMatrix(vsAdvice.Row, col_����) & strLabel
        End If
    End If
    
    If int��� = 1 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vbRed
    ElseIf int��� = 0 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vbBlue
    Else
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, col_ҽ������)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSended(ByVal lngҽ��ID As Long, Optional dat����ʱ�� As Date) As Boolean
'���ܣ��ж�Ƥ�Զ�Ӧ��ҽ���Ƿ��Ѿ�����(ֻ�ж�Ƥ��ҽ����ʼʱ��֮���ҽ��77377)
'������lngҽ��ID=Ƥ��ҽ����ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '�����ϵĲ���
    strSQL = "Select ������ĿID From ����ҽ����¼ Where ID=[3]"
    strSQL = "Select A.ID From ����ҽ����¼ A,�����÷����� B" & _
        " Where Rownum<2 And A.������� IN('5','6') And A.ҽ��״̬=8" & _
        " And A.������ĿID=B.��ĿID And B.����=0 And B.�÷�ID=(" & strSQL & ")" & _
        " And A.����ID+0=[1] And A.�Һŵ�=[2] And A.����ʱ��>=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�, lngҽ��ID, dat����ʱ��)
    AdviceSended = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSend(blnAuto As Boolean)
'���ܣ����Ͳ���ҽ��(�������üƼ���Ŀ)

    If mlng����ID = 0 Then Exit Sub
    If Not mblnEditable Then Exit Sub

    If mfrmSend Is Nothing Then Set mfrmSend = New frmOutAdviceSend
    If mfrmSend.ShowMe(mfrmParent, mMainPrivs, mlng����ID, mstr�Һŵ�, mstrǰ��IDs, blnAuto, mlng�������ID, mint����, mclsMipModule) Then
        Call LoadAdvice
        Call ShowTotalMoney
    End If
End Sub


Private Sub FuncToolScheme()
'���ܣ����ó��׷���ά��
    On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������û����ȷ��װ�������޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallClinicScheme(mfrmParent, gcnOracle, glngSys, gstrDBUser, IIF(mint���� = 2, 3, 1))
End Sub

Private Sub FuncEPRReport(ByVal lngMenu As Long)
'���ܣ����ġ���ӡ��Ԥ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strBill As String, strTmp As String
    Dim strNO As String, int���� As Long, i As Long
    Dim lngҽ��ID As Long, lngReportID As Long, blnPrint As Boolean, bln��ӡ As Boolean
    Dim bln������ As Boolean, bln�䷽�� As Boolean, arrRPTPar(19) As String, strFlagString As String
    Dim str��鱨��ID As String
    Dim lngViewMode As Long ' 1-������ʽ��6-�����ʽ
    Dim blnLis�ӿ� As Boolean
    
    On Error GoTo errH
    If mblnMoved Then
        MsgBox "��ǰ���˱���������ת������ͳһ�����Ӳ�������ģ���н��в鿴��", vbInformation, gstrSysName
        Exit Sub
    End If
    '�������ݽ���ƽ̨����LIS,PACS���ı���
    If lngMenu = conMenu_Edit_Compend * 10# + 1 Or lngMenu = conMenu_Edit_Compend * 10# + 6 Or lngMenu = conMenu_Edit_Compend Then
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Then
            lngViewMode = 1
        ElseIf lngMenu = conMenu_Edit_Compend * 10# + 6 Then
            lngViewMode = 6
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 1 Then
                lngViewMode = 1
            Else
                lngViewMode = 6
            End If
        End If
        
        If gobjExchange Is Nothing Then
            On Error Resume Next
            Set gobjExchange = CreateObject("zlExchange.clsExchange")
            If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
            err.Clear: On Error GoTo 0
        End If
        If Not gobjExchange Is Nothing Then
            With vsAdvice
                '�����д���ǲɼ��������������ΪE��������ֻ�жϼ����
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_�������) = "D", 4, 3), "ҽ��ID::" & .TextMatrix(.Row, COL_ID) & "||����Ա����::" & UserInfo.���� & "||����Աȱʡ����::" & UserInfo.������)
            End With
            Exit Sub
        End If
    End If
    
    lngReportID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID))
    lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    str��鱨��ID = vsAdvice.TextMatrix(vsAdvice.Row, COL_��鱨��ID)
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS����ID)) <> 0 Then
        Call FuncLisRptFileView(mfrmParent, lngҽ��ID)   '������LIS�ļ�����
        If lngReportID = 0 And str��鱨��ID = "" Then Exit Sub
    End If
    
    '���ж��Ƿ���Լ�������
    Select Case CheckEPRReport(lngҽ��ID, lngReportID, , , mblnMoved)
    Case 0
        MsgBox "��ҽ���ı���û����д��", vbInformation, gstrSysName
        Exit Sub
    Case 2
        strTmp = ""
        '����ҽ�����߱����ɫͨ����Ŀ���Բ鿴δ��ɵı���
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��־)) = 1 Then
            strTmp = "����鿴δ��ɱ���"
        Else
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "D" Then
                strSQL = "select 1 from Ӱ�����¼ a where a.��ɫͨ��=1 and a.ҽ��id=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                If Not rsTmp.EOF Then
                    strTmp = "����鿴δ��ɱ���"
                End If
            End If
        End If
        If InStr(GetInsidePrivs(p����ҽ���´�), "����δ��ɱ���") > 0 Or strTmp <> "" Then
            MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
        Else
            MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS����ID)) <> 0 Then
        If HaveRIS Then 'RIS�������
            i = gobjRis.ShowViewReport(mfrmParent.hwnd, lngҽ��ID, InStr(GetInsidePrivs(p����ҽ���´�), ";�����ӡ;") > 0)
            If i = 0 Then Exit Sub
        End If
    End If
    
    'ִ�в���
    '�°�PACS���棬ֱ��ǿ��ʹ���°�PACS����༭��
    If str��鱨��ID <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(lngҽ��ID, , mblnAutoRead, mfrmParent)
    Else
        bln��ӡ = InStr(GetInsidePrivs(p����ҽ���´�), ";�����ӡ;") > 0 And mblnEditable
        
        '������ĿӦ�õ���LIS�ӿ�
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" Then
            Call InitObjLis(p����ҽ��վ)
            If Not gobjLIS Is Nothing Then
                blnLis�ӿ� = True
            End If
        End If
        
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Or (lngMenu = conMenu_Edit_Compend And lngViewMode = 1) Then
            '���ı���
            If blnLis�ӿ� Then
                strTmp = ""
                Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lngҽ��ID, 0, strTmp)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                RaiseEvent ViewEPRReport(lngReportID, bln��ӡ)
            End If
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 1 And lngMenu <> conMenu_Edit_Compend * 10# + 6 And Not (lngMenu = conMenu_Edit_Compend And lngViewMode = 6) Then
                '���༭��ʽ��ӡ��Ԥ������
                If blnLis�ӿ� Then
                    strTmp = ""
                    Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lngҽ��ID, 0, strTmp)
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    RaiseEvent PrintEPRReport(lngReportID, lngMenu = conMenu_Edit_Compend * 10# + 3)
                End If
            Else
                bln������ = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E"
                If Not bln������ Then bln�䷽�� = RowIs�䷽��(vsAdvice.Row)
                    
                If bln������ Then
                    If blnLis�ӿ� Then
                        strTmp = ""
                        Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lngҽ��ID, 1, strTmp)
                        If strTmp <> "" Then
                            MsgBox strTmp, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Else
                        '����LisWork��ӡ���鱨��
                        blnPrint = IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, True, False)
                        If Not Open_LIS_Report(Me, lngҽ��ID, mlng����ID, mblnMoved, blnPrint, Not bln��ӡ) Then
                            MsgBox "��ҽ���ı���Ϊ�°�LIS��������ʹ��(���������)���ܣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    '��ȡ���һ�η��͵�NO,����
                    If bln������ Or bln�䷽�� Then
                        '����ҽ��Ӧ�Լ�����Ŀ��NOΪ׼
                        strSQL = "Select ID From ����ҽ����¼ Where ���ID=[1] And Rownum=1"
                        strSQL = "Select ҽ��ID,NO,��¼���� From ����ҽ������ Where ҽ��ID=(" & strSQL & ") Order by ���ͺ� Desc"
                    Else
                        strSQL = "Select ҽ��ID,NO,��¼���� From ����ҽ������ Where ҽ��ID=[1] Order by ���ͺ� Desc"
                    End If
                                        If mblnMoved Then
                        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                    If Not rsTmp.EOF Then
                        strNO = NVL(rsTmp!NO): int���� = NVL(rsTmp!��¼����, 0)
                    End If
                    
                    '�������ʽ��ӡ��Ԥ������
                    strSQL = "Select ��� From �����ļ��б� Where ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�ļ�ID)))
                    If Not rsTmp.EOF Then
                        strBill = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-2"
                    End If
                    
                    If lngMenu = conMenu_Edit_Compend * 10# + 2 Then
                        If Not ReportPrintSet(gcnOracle, glngSys, strBill, Me) Then Exit Sub
                    End If
                    
                    
                    If Not bln������ And Not bln�䷽�� Then
                        strFlagString = GetRPTPicture(mblnMoved, lngReportID, strBill, arrRPTPar)
                    End If
                    
                    If lngMenu <> conMenu_Edit_Compend * 10# + 2 And Not bln��ӡ Then
                        strTmp = "DisabledPrint=1"
                    Else
                        strTmp = "DisabledPrint=0"
                    End If
                    
                    'ҽ��IDΪ�ɼ���ʽ��ID������������ID
                    Call ReportOpen(gcnOracle, glngSys, strBill, Me, "NO=" & strNO, "����=" & int����, _
                        "ҽ��ID=" & lngҽ��ID, _
                        strFlagString, _
                        arrRPTPar(0), arrRPTPar(1), arrRPTPar(2), arrRPTPar(3), arrRPTPar(4), arrRPTPar(5), _
                        arrRPTPar(6), arrRPTPar(7), arrRPTPar(8), arrRPTPar(9), arrRPTPar(10), arrRPTPar(11), _
                        arrRPTPar(12), arrRPTPar(13), arrRPTPar(14), arrRPTPar(15), arrRPTPar(16), arrRPTPar(17), _
                        arrRPTPar(18), arrRPTPar(19), strTmp, _
                        IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, 2, 1))
                End If
            End If
        End If
        
        '�Զ����Ϊ�Ѳ��ģ���ʿ���Ĳ���
        If mblnAutoRead And mint���� <> 1 Then Call FuncExecReportRead(True, True)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecReportRead(ByVal blnRead As Boolean, Optional ByVal blnAuto As Boolean)
'���ܣ����õ�ǰ����Ϊ�Ѳ��ģ�����ȡ����ǰ����Ĳ���״̬
'������blnRead=���Ļ���ȡ���Ķ�״̬
'      blnAuto=����Ϊ����ʱ���Ƿ��Զ������ڵ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strAdvice As String
    Dim strTmp As String
    Dim strErr As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) = 0 Then Exit Sub
        '�°�PACS�༭�����棬ֱ�ӵ��ýӿڱ������
        If .TextMatrix(.Row, COL_��鱨��ID) = "" Then
            If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then Exit Sub
            
            If blnRead Then
                If Not blnAuto Then
                    If Val(.Cell(flexcpData, .Row, COL_����״̬)) = 1 Then Exit Sub '�Զ����ʱ���ƴ���
                    If MsgBox("��ȷ�ϸ���Ŀ�������Ѿ���ϸ�Ķ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                strSQL = "Zl_������ļ�¼_Insert(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_����ID)) & ")"
            Else
                If MsgBox("��ȷʵҪȡ���ñ���Ĳ���״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                strSQL = "Zl_������ļ�¼_Cancel(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_����ID)) & ",'" & UserInfo.���� & "')"
            End If
            Call InitObjLis(p����ҽ��վ)
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, "FuncExecReportRead")
            If Not gobjLIS Is Nothing Then
                '������ñ�ǽӿ�
                strTmp = "Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1] order by ���"
                Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
                Do While Not rsTmp.EOF
                    strAdvice = strAdvice & "," & rsTmp!ID
                    rsTmp.MoveNext
                Loop
                If .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "6" Then
                    gobjLIS.WriteAdvicesLookState Mid(strAdvice, 2), IIF(blnRead, 1, 0)
                End If
            End If
            On Error GoTo 0
        Else
            Call CreateObjectPacs(mobjPublicPACS)
            Call mobjPublicPACS.zlDocViewStateUpdate(blnRead, Val(.TextMatrix(.Row, COL_ID)))
        End If
        '���ý���״̬
        If blnRead Then
            .Cell(flexcpData, .Row, COL_����״̬) = 1 '���Ѳ���
        Else
            On Error GoTo errH
            strSQL = "Select Count(*) as ���� From ������ļ�¼ Where ҽ��ID=[1] And ȡ��ʱ�� Is Null"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
            If NVL(rsTmp!����, 0) = 0 Then
                .Cell(flexcpData, .Row, COL_����״̬) = 0 '��δ����
            End If
        End If
        Call SetAdviceReportIcon(.Row)
        .TextMatrix(.Row, COL_����״̬) = "����"
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcAppend_GotFocus()
    If vsAppend.Visible And vsAppend.Enabled Then
        vsAppend.SetFocus
    ElseIf rtfAppend.Visible And rtfAppend.Enabled Then
        rtfAppend.SetFocus
    End If
End Sub

Private Sub tbcAppend_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim blnDo As Boolean
    
    If Item.Tag = "" Then Exit Sub
    
    If Visible Then
        If Decode(vsAppend.Tag, "�Ƽ�", True, "����", True, "ǩ��", True, False) Then
            Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    vsAppend.Tag = Item.Tag '���ڹ����������ָ��Ի�
    
    If Item.Tag = "�Ƽ�" Then
        Call InitPriceTable
    ElseIf Item.Tag = "����" Then
        Call InitSendTable
    ElseIf Item.Tag = "ǩ��" Then
        Call InitSignTable
    ElseIf Item.Tag = "����" Then
        'NoneCode
    ElseIf Item.Tag = "����" Then
        'NoneCode
    End If
    
    If Visible Then
        If Decode(Item.Tag, "�Ƽ�", True, "����", True, "ǩ��", True, False) Then
            Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then
        If vsAppend.Visible And vsAppend.Enabled Then
            vsAppend.SetFocus
        ElseIf rtfAppend.Visible And rtfAppend.Enabled Then
            rtfAppend.SetFocus
        End If
    End If
End Sub

Private Sub vsAdvice_Click()
'���ܣ����ı���
    Dim lngMouseRow As Long, lngMouseCol As Long
    
    If mblnTag Then Exit Sub '����ѵ�����鿴���棬����ʾ����ǰ�������ڵ���鿴
    'PASS
    If mblnPass And Me.Visible Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap, 1)
    End If
    
    On Error GoTo errH
    
    If mvarCond.����ģʽ <> 3 Then Exit Sub
    With vsAdvice
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                .Redraw = False
                mblnTag = True
                Call FuncEPRReport(conMenu_Edit_Compend)
                .Cell(flexcpForeColor, lngMouseRow, COL_����״̬) = &H80& '����
                mblnTag = False
                .Redraw = True
            End If
        End If
    End With
    Exit Sub
errH:
    mblnTag = False
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnExist As Boolean, blnSel As Boolean, bln��Ѫ As Boolean
    Dim varDraw As RedrawSettings, intIdx As Integer
    
    If NewRow = OldRow Then Exit Sub
    If fraMore.Visible = True Then fraMore.BackColor = vsAdvice.BackColorSel
    
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_��ʼʱ��)
    End If
    
     'PASS
    If mblnPass And Me.Visible Then
        If NewRow <> OldRow Then
            Call gobjPass.zlPassSetDrug(mobjPassMap)
        End If
    End If
    
    Call LoadBillList '��ʾ�ɴ�ӡ�����Ƶ���
    If vsAdvice.Redraw <> flexRDNone Then
        If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
            '��ʾ�����Ƿ������Ķ�
            If Val(vsAdvice.TextMatrix(NewRow, COL_����ID)) <> 0 Or vsAdvice.TextMatrix(NewRow, COL_��鱨��ID) <> "" Then
                On Error GoTo errH
                strSQL = "Select 1 From ������ļ�¼ Where ҽ��ID=[1] And ������=[2] And ȡ��ʱ�� Is NULL"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", _
                    Val(vsAdvice.TextMatrix(NewRow, COL_ID)), UserInfo.����)
                If Not rsTmp.EOF Then
                    If vsAdvice.TextMatrix(NewRow, COL_��鱨��ID) = "" Then
                        vsAdvice.Cell(flexcpData, NewRow, COL_����״̬) = 1
                    Else
                        '���ֲ��ĵ�
                        strSQL = "Select 1 From ����ҽ������ A Where not exists(select 1 from ������ļ�¼ B where B.ҽ��ID=A.ҽ��ID And A.��鱨��ID=B.��鱨��ID And B.������=[2] And B.ȡ��ʱ�� Is NULL) and A.ҽ��ID=[1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", Val(vsAdvice.TextMatrix(NewRow, COL_ID)), UserInfo.����)
                        vsAdvice.Cell(flexcpData, NewRow, COL_����״̬) = IIF(Not rsTmp.EOF, 2, 1)
                    End If
                Else
                    vsAdvice.Cell(flexcpData, NewRow, COL_����״̬) = 0
                End If
                On Error GoTo 0
            End If
        
            '��ʾҽ�����ӱ�������
            If mblnAppend Then
                '�жϵ��ݸ����Ƿ�������
                blnSel = False: blnExist = False
                Call ShowBillAppend(NewRow, blnExist)
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "����" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '�������������ظ�����
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                '�жϸ�����Ϣ����ʾ
                blnSel = False: blnExist = False
                Call ShowAdvicePlan(NewRow, blnExist)
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "����" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '�������������ظ�����
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                '�ж�ҽ���Ƿ����(���ϵ�ҽ������ʾѪҺҳ��)
                blnSel = False: blnExist = False: bln��Ѫ = False
                If gblnѪ��ϵͳ And vsAdvice.TextMatrix(NewRow, COL_�������) = "K" Then
                    bln��Ѫ = True
                    With vsAdvice
                        '��Ѫҽ�����״̬=1��������Ѫ�Ʒ�Ѫ�����Ĵ��˶�ҽ����������Ѫҽ�������״̬=4������ҽ����δ����Ѫ�ּ�����ʱ����ʾΪ�ȴ���Ѫ
                        If Val(.TextMatrix(NewRow, COL_���״̬)) = 1 And Val(.TextMatrix(NewRow, COL_��鷽��)) = 1 Then
                            blnExist = True
                        Else
                            blnExist = InStr(",,2,3,4,5,6,", "," & .TextMatrix(NewRow, COL_���״̬) & ",") > 0 And Not (.TextMatrix(NewRow, COL_ҽ��״̬) = "4")
                        End If
                    End With
                End If
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "ѪҺ" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '�������������ظ�����
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                blnSel = False: blnExist = False
                If bln��Ѫ = False Then
                    With vsAdvice
                        blnExist = InStr(",2,3,4,5,", "," & .TextMatrix(NewRow, COL_���״̬) & ",") > 0
                        '����Ѫҽ��ʱ����Ѫ��ϵͳ�����Ϊ4�����״̬������ҽ����δ����Ѫ�ּ�����ʱ�� ���״̬Ϊ4ʱû����Ӧ�Ĳ�����¼<����ҽ��״̬>
                        If Val(.TextMatrix(NewRow, COL_���״̬)) = 4 And .TextMatrix(NewRow, COL_�������) = "K" Then
                            If Val(.TextMatrix(NewRow, COL_��־)) = 1 Or Not gbln��Ѫ�ּ����� Then blnExist = False
                        End If
                    End With
                End If
                
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "����" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '�������������ظ�����
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                '��ԤԼ��Ϣ����ʾ
                blnSel = False: blnExist = False
                Call ShowAdviceRISSch(NewRow, blnExist)
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "ԤԼ" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '�������������ظ�����
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                If tbcAppend.Selected.Tag = "�Ƽ�" Then
                    Call ShowPrice(NewRow)
                ElseIf tbcAppend.Selected.Tag = "����" Then
                    Call ShowSendList(NewRow)
                ElseIf tbcAppend.Selected.Tag = "ǩ��" Then
                    Call ShowSignList(NewRow)
                ElseIf tbcAppend.Selected.Tag = "����" Then
                    'ǰ���ѹ̶���ȡ
                ElseIf tbcAppend.Selected.Tag = "ԤԼ" Then
                    'ǰ���ѹ̶���ȡ
                ElseIf tbcAppend.Selected.Tag = "����" Then
                    'ǰ���ѹ̶���ȡ
                ElseIf tbcAppend.Selected.Tag = "����" Then
                    Call ShowOtherAppend(NewRow)
                ElseIf tbcAppend.Selected.Tag = "ѪҺ" Then
                    If Not mobjFrmBloodList Is Nothing Then
                        Call mobjFrmBloodList.zlRefresh(Val(vsAdvice.TextMatrix(NewRow, COL_ID)), mlngFontSize, mblnMoved)
                    End If
                End If
            End If
        ElseIf mblnAppend Then
            Call ClearAppendData
            vsAppend.Row = vsAppend.FixedRows
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col_ҽ������ Or Col = col_���� Then
        vsAdvice.AutoSize Col, COL_�÷�
    ElseIf Col = COL_Ƥ�� Then
        If vsAdvice.ColWidth(Col) > 1200 Then vsAdvice.ColWidth(Col) = 1200
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_��ʾ Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        ElseIf Col = COL_������ Or Col = COL_������ӡ Or Col = COL_����Ԥ�� Then
            lngLeft = COL_������: lngRight = COL_����Ԥ��
            If Not RowInSameNo(Row, lngBegin, lngEnd) Then Exit Sub
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                'Ϊ��֧��Ԥ�����
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '����һ����ҩ������еı��߼�����
            lngLeft = COL_��ʼʱ��: lngRight = COL_��ʼʱ��
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_����: lngRight = COL_�÷�
            End If
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_Ƥ��: lngRight = COL_Ƥ��
            End If
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                'Ϊ��֧��Ԥ�����
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
        Dim rsTmp As Recordset
    
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    ElseIf Button = 1 Then
        If mvarCond.����ģʽ = 0 And mvarCond.ҽ�� = 1 Then
            With vsAdvice
                If .MouseRow >= .FixedRows And (.MouseCol = COL_������ӡ Or .MouseCol = COL_����Ԥ��) Then
                    If .TextMatrix(.MouseRow, .MouseCol) = "" Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                mbln����Ԥ�� = (.MouseCol = COL_����Ԥ��)
            End With
            vsAdvice.Redraw = flexRDNone
            If mcbsMain Is Nothing Then
                Set rsTmp = GetBillList
                If rsTmp.RecordCount > 0 Then
                    FuncBillPrint , "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" & "|" & rsTmp!NO, rsTmp!���� '��Ӧ���Զ��屨����
                End If
                Exit Sub
            End If
            Set objControl = mcbsMain.FindControl(, conMenu_Report_ClinicBill * 100# + 1, , True)
            If Not objControl Is Nothing Then
                objControl.Execute
            Else
                MsgBox "��ҩƷû�ж�Ӧ���Ƶ��ݣ�û�п��Դ�ӡ�Ĵ���ǩ��", vbInformation, gstrSysName
            End If
            vsAdvice.Redraw = flexRDDirect
        End If
    End If
End Sub

Private Function GetBillList() As Recordset
    Dim strSQL As String
    With vsAdvice
        strSQL = "Select Distinct D.���,D.����,D.˵��,B.NO" & _
            " From ����ҽ����¼ A,����ҽ������ B,��������Ӧ�� C,�����ļ��б� D" & _
            " Where C.������ĿID = A.������ĿID And a.ID=b.ҽ��ID " & _
            " And C.Ӧ�ó���=1 And C.�����ļ�ID=D.ID And D.����=7 And (a.ID=[1] or A.���ID=[1])" & _
            " Order by D.���"
       
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
         Set GetBillList = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Decode(Val(.TextMatrix(.Row, COL_���ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_���ID))))
    End With
End Function

Private Function GetPatiInfo() As ADODB.Recordset
'���ܣ���ȡ������Ϣ
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ִ�в���(�ű����)�����˿���
    strSQL = "Select A.����,A.�Ա�,A.����,B.�����,B.סԺ��,B.������,a.ID as �Һ�ID," & _
        " B.����,B.��������,C.���� as ִ�в���,A.�Ǽ�ʱ��,B.�ѱ�" & _
        " From ���˹Һż�¼ A,������Ϣ B,���ű� C" & _
        " Where A.NO(+)=[2] And a.��¼����(+)=1 And a.��¼״̬(+)=1 And B.����ID=[1]" & _
        " And A.����ID(+)=B.����ID And A.ִ�в���ID=C.ID(+)"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�)
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String, strInfo As String, rsTmp As ADODB.Recordset
    
    If mlng����ID = 0 Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = "����ҽ���嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set rsTmp = GetPatiInfo
    strInfo = _
        "������" & rsTmp!���� & " �Ա�" & NVL(rsTmp!�Ա�) & _
        " ���䣺" & NVL(rsTmp!����) & " ����ţ�" & NVL(rsTmp!�����) & _
        " �Һţ�" & IIF(IsNull(rsTmp!�Ǽ�ʱ��), "", Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm")) & _
        " ���ң�" & NVL(rsTmp!ִ�в���) & " ���ң�" & NVL(rsTmp!��������)
    Set objRow = New zlTabAppRow
    objRow.Add strInfo
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsAdvice
    
    '���
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
    
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim strTab As String, i As Integer
    Dim intType As Integer
    
    mblnFirst = False
    Set mrsPlugInBar = Nothing
    mlngPromptRow = 0
    mbln����Ԥ�� = False
    
    If Not grsSkinTest Is Nothing Then
        grsSkinTest.Close
        Set grsSkinTest = Nothing
    End If
    
    'ҽ���嵥
    '-----------------------------------------------------
    mlngFontSize = 9
    Call InitAdviceTable
    Call InitColumnSelect '��ʼ����ѡ����
    
    'CommandBars
    '-----------------------------------------------------
    Call GetFilterSetting '���ع��˲���
    Call InitFilterBar
    
    'TabControl
    '-----------------------------------------------------
    With tbcMain
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        .InsertItem(0, " ҽ  �� ", picMain.hwnd, 0).Tag = "ҽ��"
        .InsertItem(1, " ��  �� ", picMain.hwnd, 0).Tag = "����"
    End With
    tbcMain.Item(tbcMain.ItemCount - 1).Selected = True
    i = IIF(mvarCond.����ģʽ = 0, 0, 1)
    tbcMain.Item(i).Selected = True
    
    With tbcAppend
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        .InsertItem(0, "ҽ���Ƽ�����", vsAppend.hwnd, 0).Tag = "�Ƽ�"
        .InsertItem(1, "ҽ�����ͼ�¼", vsAppend.hwnd, 0).Tag = "����"
        If Not gobjESign Is Nothing Then '����ǩ����¼
            .InsertItem(2, "ҽ��ǩ����¼", vsAppend.hwnd, 0).Tag = "ǩ��"
        End If
        .InsertItem(3, "���븽��", rtfAppend.hwnd, 0).Tag = "����"
        .InsertItem(4, "�������", rtfInfo.hwnd, 0).Tag = "����"
        .InsertItem(5, "ԤԼ��Ϣ", rtfSche.hwnd, 0).Tag = "ԤԼ" 'RISԤԼ��Ϣ
        .InsertItem(6, "������Ϣ", rtfOther.hwnd, 0).Tag = "����" '����ҩ�������Ϣ
        If gblnѪ��ϵͳ = True Then
            If InitObjBlood = True Then
                Set mobjFrmBloodList = gobjPublicBlood.zlGetBloodListInfo
                .InsertItem(7, "ѪҺ��Ϣ", mobjFrmBloodList.hwnd, 0).Tag = "ѪҺ"  'ѪҺ��Ѫ��Ϣ
            End If
        End If
        '��Ϊ����ͬ,���Ҫ�л��ص�1��;�����ݲ�Ӱ���ٶ�
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    mblnAppend = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "AppendData", 1)) <> 0
    tbcAppend.Visible = mblnAppend: fraAdviceUD.Visible = mblnAppend
    If mblnAppend Then
        strTab = zlDatabase.GetPara("ҽ�����б�", glngSys, p����ҽ���´�, "")
        If strTab <> "" Then
            For i = 0 To tbcAppend.ItemCount - 1
                If tbcAppend(i).Visible And tbcAppend(i).Tag = strTab Then
                    tbcAppend.Item(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
        
    '�ָ����Ի�����
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    vsAdvice.ColWidth(COL_F��־) = 11 * Screen.TwipsPerPixelX
    vsAdvice.ColWidth(COL_F����) = 11 * Screen.TwipsPerPixelX
    
    '������ʼ��
    '-----------------------------------------------------
    mMainPrivs = gMainPrivs '������ģ��Ȩ��
    Set mfrmEdit = Nothing
    Set mobjReport = New clsReport
    Set mrsDefine = InitAdviceDefine
    
    
    '����ע�������
    Call GetLocalSetting
    mblnAutoRead = Val(zlDatabase.GetPara("�Զ���Ǳ������״̬", glngSys, p����ҽ���´�, "1", , , intType)) = 1
    mblnAutoReadEnabled = Not ((intType = 3 Or intType = 15))
        
    If gblnKSSStrict Then Call CheckKSSPrivilege(2)
    If mint���� = 0 Then Call InitObjLis(p����ҽ��վ)
    On Error Resume Next
    Set gobjExchange = CreateObject("zlExchange.clsExchange")
    If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
    err.Clear: On Error GoTo 0
End Sub

Private Sub InitFilterBar()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    cbsSub.AddImageList img16 '��VB.ImageList��Tag��ID���й���
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    
    Set objBar = cbsSub.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objPopup = .Add(xtpControlPopup, ID_Ӥ��, "����ҽ��")
            objPopup.ID = ID_Ӥ��: objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100#, "����ҽ��")
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 1, "����ҽ��"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 2, "Ӥ�� 1 ҽ��"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 3, "Ӥ�� 2 ҽ��")
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 4, "Ӥ�� 3 ҽ��")
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 5, "Ӥ�� 4 ҽ��")
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 6, "Ӥ�� 5 ҽ��")
        End With
        
        Set objControl = .Add(xtpControlButton, ID_��ֹ, "������")
            objControl.BeginGroup = True
            objControl.ToolTipText = "��ʾ�Ѿ����ϵ�ҽ��"
            
        '----------------����ҳ��
        Set objControl = .Add(xtpControlButton, ID_ȫ��, "ȫ��")
            objControl.BeginGroup = True
            objControl.Checked = True
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.IconId = 1 '��ʼʱ����ͼ��
        Set objControl = .Add(xtpControlButton, ID_����, "����")
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.IconId = 1
        
        Set objControl = .Add(xtpControlButton, ID_δ������, "δ������")
            objControl.ToolTipText = "��ʾδ������"
            objControl.BeginGroup = True
            mvarCond.δ������ = True
            
        Set objControl = .Add(xtpControlButton, ID_�ѳ�����, "�ѳ�����")
            objControl.ToolTipText = "��ʾ�ѳ�����"
            mvarCond.�ѳ����� = True
        
        Set objControl = .Add(xtpControlButton, ID_ҽ��ȫ��, "ȫ��")
            objControl.BeginGroup = True
            objControl.Checked = True
        Set objControl = .Add(xtpControlButton, ID_ҽ������, "����")
            objControl.IconId = 1
        Set objControl = .Add(xtpControlButton, ID_ҽ������, "����")
            objControl.IconId = 1
        '-----------------ҽ��
        
        
        Set objControl = .Add(xtpControlButton, ID_����, "�����´�")
            objControl.BeginGroup = True
            objControl.ToolTipText = "ֻ��ʾҽ�������´��ҽ��"
        
        
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButton, ID_����, "��ϸ")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
            
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsSub.KeyBindings
        .Add FCONTROL, vbKeyB, ID_Ӥ�� * 100#
        .Add FCONTROL, vbKey0, ID_Ӥ�� * 100# + 1
        .Add FCONTROL, vbKey1, ID_Ӥ�� * 100# + 2
        .Add FCONTROL, vbKey2, ID_Ӥ�� * 100# + 3
        .Add FCONTROL, vbKey3, ID_Ӥ�� * 100# + 4
        .Add FCONTROL, vbKey4, ID_Ӥ�� * 100# + 5
        .Add FCONTROL, vbKey5, ID_Ӥ�� * 100# + 6
        .Add FCONTROL, vbKey8, ID_��ֹ
        .Add FCONTROL, vbKeyK, ID_����
    End With
    objBar.Visible = Not mblnHideFilter
    fraHide.Visible = mblnHideFilter
    fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
End Sub

Private Sub mfrmParent_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�����������İ���,���ڴ���ҽ�������ȼ�
'˵����
'1.��ҽ���Ӵ���δ����ʱ,�Ӵ���CommandBar���ȼ���Ч
'2.������CommandBar��KeyDown�¼������˵ļ������ټ�����¼�
    
    If Not Me.Visible Then Exit Sub '�������Ӵ���ʱ�Իἤ��
    If mlng����ID = 0 Then Exit Sub
    
    Call ActiveHotKey(KeyCode, Shift)

End Sub

Private Sub ActiveHotKey(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim lngID As Long
    Dim intTab As Integer
    
    If Not Me.Visible Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    intTab = -1
    
    If Shift = vbCtrlMask And KeyCode >= vbKey0 And KeyCode <= vbKey5 Then
        lngID = ID_Ӥ�� * 100# + KeyCode - vbKey0 + 1
    ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKey6
                intTab = 0
            Case vbKey8
                lngID = ID_��ֹ
            Case vbKey9
                intTab = 1
            Case vbKeyB
                lngID = ID_Ӥ�� * 100#
            Case vbKeyK
                lngID = ID_����
            Case vbKeyX
                lngID = ID_���
            Case vbKeyY
                lngID = ID_����
            Case vbKeyQ
                lngID = ID_����
        End Select
    End If
    If lngID <> 0 Then
        Set objControl = cbsSub.FindControl(, lngID, , True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
    If intTab <> -1 Then tbcMain.Item(intTab).Selected = True
End Sub

Private Sub GetLocalSetting()
'���ܣ���ȡ����ע�������
    mblnƤ������ = Val(zlDatabase.GetPara("ҽ������Ƥ������", glngSys, p����ҽ���´�)) <> 0
    'ִ������
    mbln���� = Val(zlDatabase.GetPara("ҽ��ִ������", glngSys, p����ҽ���´�)) <> 0
    
    mblnָ������ӡ = Val(zlDatabase.GetPara("ָ������ӡ��ʽ", glngSys, p����ҽ���´�)) <> 0
    
    mblnΣ��ֵ = InStr(GetInsidePrivs(p����ҽ��վ), ";Σ��ֵ����;") > 0
End Sub

Private Sub GetFilterSetting()
'���ܣ���ȡҽ��������������
    Dim strPar As String
    
    mvarCond.Ӥ�� = 0
    mvarCond.��ֹ = Val(zlDatabase.GetPara("ҽ����ʾ����", glngSys, p����ҽ���´�, "1")) = 1
    mvarCond.���� = Val(zlDatabase.GetPara("����ҽ������", glngSys, p����ҽ���´�, "1")) <> 0
    mblnHideFilter = Val(zlDatabase.GetPara("���������Զ�����", glngSys, p����ҽ���´�, "0")) <> 0
    
    strPar = Val(zlDatabase.GetPara("����鿴����", glngSys, p����ҽ���´�, "0"))
    If InStr(",0,1,2,3,", "," & strPar & ",") > 0 Then
        mvarCond.���� = Val(strPar)
    Else
        mvarCond.���� = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("��ʾģʽ", glngSys, p����ҽ���´�, "0"))
    mvarCond.��ʾģʽ = IIF(Val(strPar) = 0, 0, 1)
    
    strPar = Val(zlDatabase.GetPara("ҽ����ʾ����", glngSys, p����ҽ���´�, "0"))
    If InStr(",0,1,2,", "," & strPar & ",") > 0 Then
        mvarCond.ҽ�� = Val(strPar)
    Else
        mvarCond.ҽ�� = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("ҽ�����˷�ʽ", glngSys, p����ҽ���´�, "0"))
    mvarCond.����ģʽ = IIF(Val(strPar) = 0, 0, 3)
End Sub

Private Sub SaveFilterSetting()
'���ܣ�����ҽ��������������
    Call zlDatabase.SetPara("����ҽ������", IIF(mvarCond.����, 1, 0), glngSys, p����ҽ���´�)
    Call zlDatabase.SetPara("��ʾģʽ", mvarCond.��ʾģʽ, glngSys, p����ҽ���´�)
    Call zlDatabase.SetPara("����鿴����", mvarCond.����, glngSys, p����ҽ���´�)
    Call zlDatabase.SetPara("���������Զ�����", IIF(mblnHideFilter, 1, 0), glngSys, p����ҽ���´�)
    Call zlDatabase.SetPara("ҽ����ʾ����", mvarCond.ҽ��, glngSys, p����ҽ���´�)
    Call zlDatabase.SetPara("ҽ����ʾ����", IIF(mvarCond.��ֹ, 1, 0), glngSys, p����ҽ���´�)
    Call zlDatabase.SetPara("ҽ�����˷�ʽ", mvarCond.����ģʽ, glngSys, p����ҽ���´�)
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    With Me.tbcMain
        .Left = 0
        .Top = 0
        .Height = Me.Height
        .Width = Me.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmSend Is Nothing Then Unload mfrmSend: Set mfrmSend = Nothing
    If Not mfrmEdit Is Nothing Then Unload mfrmEdit: Set mfrmEdit = Nothing
    Set mobjReport = Nothing
    Set gobjExchange = Nothing
    Set gobjLIS = Nothing
    Set mobjPublicPACS = Nothing
    Set gobjRecipeAudit = Nothing
    
    If Not mobjFrmBloodList Is Nothing Then
        Unload mobjFrmBloodList
        Set mobjFrmBloodList = Nothing
    End If
    Set mrsDefine = Nothing
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mSendControl = Nothing
    If Not gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = Nothing
    End If
        
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "AppendData", IIF(mblnAppend, 1, 0)
    If mblnAppend And Not tbcAppend.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ�����б�", tbcAppend.Selected.Tag, glngSys, p����ҽ���´�)
    End If
    Call SaveFilterSetting
    Call SaveWinState(Me, App.ProductName)
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassClearLight(mobjPassMap, 1)
    End If
    mblnPass = False
    Set mobjPassMap = Nothing
 
    '��ҳ��������ֹ
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        On Error Resume Next
        Call gobjPlugIn.Terminate(glngSys, p����ҽ���´�, mint����)
        Call zlPlugInErrH(err, "Terminate")
        err.Clear: On Error GoTo 0
    End If
    Set mclsMipModule = Nothing
    mbln����Ԥ�� = False
    Set mrsΣ��ֵ = Nothing
    mblnΣ��ֵ = False
    mlngΣ��ֵID = 0
End Sub

Private Sub ClearAppendData()
'���ܣ�������ӱ������븽�������
    Dim blnSel As Boolean, intIdx As Integer
    Dim varDraw As RedrawSettings
    
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
    vsAppend.Row = vsAppend.FixedRows
        
    If rtfAppend.Visible Then rtfAppend.Text = ""
    If rtfInfo.Visible Then rtfInfo.Text = ""
    For intIdx = 0 To tbcAppend.ItemCount - 1
        If InStr("����,����,ԤԼ,����,ѪҺ", tbcAppend(intIdx).Tag) > 0 Then
            If tbcAppend(intIdx).Selected Then blnSel = True
            tbcAppend(intIdx).Visible = False
        End If
    Next
    If blnSel Then
        varDraw = vsAdvice.Redraw '�������������ظ�����
        vsAdvice.Redraw = flexRDNone
        tbcAppend.Item(0).Selected = True
        vsAdvice.Redraw = varDraw
    End If
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2500,1;��λ,500,4;�Ƽ�����,850,1;����,900,7;ִ�п���,1000,1;��������,800,1;����,450,4;�շѷ�ʽ,1500,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLPrice.Count <> UBound(arrHead) + 1 Then COLPrice.Add i, Split(arrHead(i), ",")(0)
            .MergeCol(i) = False
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitSendTable()
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "���ͺ�;����ʱ��,1530,1;���ݺ�,850,1;����ҽ��,1800,1;�շ���Ŀ,1800,1;��������,850,1;�Ʒ�״̬,850,1;ִ��״̬,850,1;״̬˵��,1800,1;ִ�п���,1000,1;ִ����,800,1;ִ��ʱ��,1530,1;ִ��˵��,1800,1;������,800,1;��¼����"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLSend.Count <> UBound(arrHead) + 1 Then COLSend.Add i, Split(arrHead(i), ",")(0)
            .MergeCol(i) = False
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .Redraw = flexRDDirect
        
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitSignTable()
'���ܣ���ʼ��ǩ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "ǩ������,1150,1;ǩ��ʱ��,1900,1;ǩ����,800,1;ʱ���,1900,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLSign.Count <> UBound(arrHead) + 1 Then COLSign.Add i, Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeNever
    End With
End Sub

Private Sub ClearAdviceData()
'���ܣ����ҽ���嵥����
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitColumnSelect()
'���ܣ�����ҽ���嵥ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vsAdvice
        For i = .FixedCols To .Cols - 1
            If Not (.ColHidden(i) Or .ColWidth(i) = 0) Then
                If .TextMatrix(0, i) <> "" And Not (i = COL_����״̬ Or i = COL_�걾״̬) Then '�����,Ƥ��
                    vsColumn.Rows = vsColumn.Rows + 1
                    lngRow = vsColumn.Rows - 1
                    vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                    vsColumn.RowData(lngRow) = i
                    
                    '�̶���ʾ��
                    If InStr(",��ʼʱ��,ҽ������,����ҽ��,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsColumn.TextMatrix(lngRow, 0) = 1
                        vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                    End If
                End If
            End If
        Next
    End With
    If vsColumn.Rows > 1 Then vsColumn.Row = 1
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "ID;���ID;Ӥ��ID;ҽ��״̬;�������;��������;�������;��־;" & _
        ",240,4;������,1000,4;��ӡ,800,4;Ԥ��,800,4;��Чʱ��,1530,1;,200,7;ҽ������,3000,1;����,4000,1;,375,1;����,850,1;����,850,1;����,450,1;Ƶ��,1000,1;" & _
        "�÷�,1000,1;ҽ������,1000,1;ִ��ʱ��,1000,1;ִ�п���,1000,1;ִ������,850,1;����ҽ��,850,1;����ʱ��,1530,1;������,850,1;����ʱ��,1530,1;����˵��,1000,1;����ҩ��,850,1;����״̬,700,4;�걾״̬,850,1;������ĿID;�Թܱ���;" & _
        "ǰ��ID;ǩ����;�ļ�ID;������;����ID;���״̬;�������;��ΣҩƷ;�걾��λ;�շ�ϸĿID;��������ID;��ҩĿ��;��鱨��ID;�������״̬;���������;RISԤԼID;RIS����ID;LIS����ID;RISԤԼ״̬;������Ŀ����;��鷽��;Σ��ֵID;�׵���"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        'δ���ú�����ҩʱ�����в��ɼ�������������̫Ԫͨʱ������gbytPass=1 or 3 ʱ �ɼ�
        vsAdvice.ColHidden(COL_��ʾ) = True
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(COL_F��־) = 11 * Screen.TwipsPerPixelX
        .ColWidth(COL_F����) = 11 * Screen.TwipsPerPixelX
        .MergeCells = flexMergeFree
        .MergeCol(COL_������) = True
        .MergeCol(COL_������ӡ) = True
        .MergeCol(COL_����Ԥ��) = True
    End With
End Sub

Private Sub SetRTFFont(bytKind As Byte)
    If bytKind = 0 Or bytKind = 1 Then
        With rtfAppend
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 2 Then
        With rtfInfo
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 3 Then
        With rtfOther
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 4 Then
        With rtfSche
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
End Sub

Private Function LoadAdvice() As Boolean
'���ܣ����ݵ�ǰ�������ö�ȡ����ʾҽ���嵥
    Dim rsTmp As ADODB.Recordset
    Dim rsѪ�� As ADODB.Recordset
    Dim strSQL As String
    Dim strFormat As String, strTmp As String, blnDo As Boolean
    Dim bln��ҩ;�� As Boolean, bln��ҩ�÷� As Boolean
    Dim bln�ɼ����� As Boolean, bln��Ѫ;�� As Boolean
    Dim blnFirst As Boolean, lngҽ��ID As Long
    Dim strBill As String, i As Long, j As Long
    Dim strWhere As String, strҽ��״̬ As String
    Dim strSameDay As String 'ͬһ��
    Dim datCur As Date, strGroupBy As String
 
    If mlng����ID = 0 Then Exit Function

    Screen.MousePointer = 11

    On Error GoTo errH
    
    lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))    '��¼��ǰ��������ڵ�ǰ����ˢ��ҽ����Ӧ�ò���
    
    If mvarCond.Ӥ�� <> -1 Then
        strWhere = strWhere & " And Nvl(A.Ӥ��,0)=[4]"
    End If
    
    If Not mvarCond.��ֹ Then
        strWhere = strWhere & " And Instr(',1,8,',','||Nvl(A.ҽ��״̬,0)||',')>0"
    End If
    
    'ҽ��վ  �����´�
    If mlngǰ��ID <> 0 And mvarCond.���� Then
        strWhere = strWhere & " And Nvl(A.ǰ��ID,0)<>0 and (A.ǰ��ID in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) X) or a.��������ID=[5])"
    End If
    
    'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨
    strSQL = _
    "Select /*+ RULE */ A.ID,A.���ID," & _
             " Nvl(A.Ӥ��,0) as Ӥ��ID,A.ҽ��״̬,A.�������,B.��������,C.�������,A.������־ as ��־," & _
             " A.�����,k.No as ������,Decode(k.no,null,null,'��ӡ') as ��ӡ,Decode(k.no,null,null,'Ԥ��') as Ԥ��,To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,Null as ��,A.ҽ������,Null as ����,A.Ƥ�Խ�� as Ƥ��," & _
             " Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ),'4',A.�ܸ�����||G.���㵥λ,'5',Round(A.�ܸ�����/D.�����װ,5)||D.���ﵥλ,'6',Round(A.�ܸ�����/D.�����װ,5)||D.���ﵥλ,A.�ܸ�����||B.���㵥λ)) as ����," & _
             " Decode(A.��������,NULL,NULL,A.��������||Decode(A.�������,'4',G.���㵥λ,B.���㵥λ)) as ����,A.����," & _
             " A.ִ��Ƶ�� as Ƶ��,Decode(A.�������,'E',Decode(Instr('2468',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�," & _
             " A.ҽ������,A.ִ��ʱ�䷽�� as ִ��ʱ��,Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���," & _
             " Decode(Instr('567E',A.�������),0,NULL,A.ִ������) as ִ������,A.����ҽ��,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��," & _
             " A.ͣ��ҽ�� as ������,A.ͣ��ʱ�� as ����ʱ��,a.����˵��,D.����ҩ��,Decode(Max(NVL(y.����״̬,0)),MiN(NVL(y.����״̬,0)),Max(NVL(y.����״̬,0)),2) As ����״̬,null as �걾״̬,A.������ĿID,B.�Թܱ���,A.ǰ��ID,Decode(A.�¿�ǩ��ID,NULL,0,1) as ǩ����," & _
             " M.�����ļ�ID as �ļ�ID,Nvl(N.ͨ��,0) as ������,Max(y.����id) As ����id,A.���״̬,A.�������,d.��ΣҩƷ,A.�걾��λ,A.�շ�ϸĿID,a.��������ID,a.��ҩĿ��," & _
             " Max(y.��鱨��id)||'' As ��鱨��id,J.״̬ as �������״̬,J.����� as ���������,f.ԤԼid As RISԤԼID,Max(y.RISID) As RIS����ID,Max(y.����ID) as LIS����ID,f.�Ƿ���� as RISԤԼ״̬,b.���� as ������Ŀ����,Max(a.��鷽��) as ��鷽��,max(h.Σ��ֵid) as Σ��ֵID,D.�Ƿ���������"
    strSQL = strSQL & _
             " From ����ҽ����¼ A,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B,�շ���ĿĿ¼ G,����ҽ������ Y,��������Ӧ�� M,�����ļ��б� N, ���������ϸ I, ��������¼ J,����ҽ������ K,RIS���ԤԼ f,����Σ��ֵҽ�� H" & _
             " Where A.������ĿID=B.ID And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+) And a.ID = i.ҽ��ID(+) And I.��ID = J.ID(+) and (I.����ύ =1 Or I.��ID is NULL) and Nvl(A.ִ�б��,0)<>-1 " & _
             " And a.id=k.ҽ��id(+) And Nvl(A.ҽ����Ч,0)=1 And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=G.ID(+) And a.Id=f.ҽ��id(+) and a.id=h.ҽ��ID(+)" & _
             " And A.ID=Y.ҽ��ID(+) And (Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL) Or A.�������='E' And B.��������='8')" & _
             " And A.��ʼִ��ʱ�� is Not NULL" & IIF(mint���� = 2, "", " And A.������Դ<>3") & _
             " And A.������ĿID=M.������ĿID(+) And M.Ӧ�ó���(+)=1 And M.�����ļ�ID=N.ID(+) And N.����(+)=7" & _
             " And A.����ID+0=[1] And A.�Һŵ�=[2]" & strWhere
    strGroupBy = " Group By a.Id, a.���id, a.���, a.Ӥ��, a.ҽ��״̬, a.�������, b.��������, c.�������, a.������־, a.�����, a.ҽ����Ч, a.��ʼִ��ʱ��, a.ҽ������, a.Ƥ�Խ��," & vbNewLine & _
            "         a.�ܸ�����, a.�״�����, g.���㵥λ, d.�����װ, d.���ﵥλ, a.��������, a.����, a.ִ��Ƶ��, a.ҽ������, b.����, a.ִ������, a.ִ��ʱ�䷽��, a.ִ����ֹʱ��, e.����," & vbNewLine & _
            "         a.�ϴ�ִ��ʱ��, a.����ʱ��, a.����ҽ��, a.У�Ի�ʿ, a.У��ʱ��, a.ͣ��ҽ��, a.ͣ��ʱ��, a.ȷ��ͣ����ʿ, a.ȷ��ͣ��ʱ��, a.������Ŀid, b.�Թܱ���, a.ִ�б��, a.���δ�ӡ," & vbNewLine & _
            "         a.ǰ��id, a.�¿�ǩ��id, m.�����ļ�id, n.ͨ��, a.�շ�ϸĿid, b.���㵥λ, a.��������id, a.���״̬, a.�������, a.��˱��, d.����ҩ��, d.��ΣҩƷ, a.�걾��λ,J.״̬,J.�����," & vbNewLine & _
            "         a.��ҩĿ��,a.����˵��,k.No,f.ԤԼid,f.�Ƿ����,b.����,D.�Ƿ���������"
     
    strSQL = strSQL & strGroupBy & " Order by Nvl(A.Ӥ��,0),A.���"

    If mblnMoved Then    '�Һŵ���ҽ��ͬ�����ݿ�
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    datCur = zlDatabase.Currentdate
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�, IIF(mstrǰ��IDs = "", "0", mstrǰ��IDs), mvarCond.Ӥ��, mlng�������ID)
    If Not rsTmp.EOF Then
        strSQL = "Select a.ҽ��id,decode(a.��ѪѪ��,1,'A',2,'B',3,'AB',4,'O','') As Ѫ�� From ��Ѫ�����¼ A, ����ҽ����¼ B Where ҽ��id = b.Id And b.�Һŵ� =[1] And a.��ѪѪ��>0 and b.�������='K'"
        Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr�Һŵ�)
    
        With vsAdvice
            .Redraw = False
                
            '��ʱ�����ʱ��FormatString�ָ�һЩȱʡֵ(�̶����������̶��������ּ����ж���,�ߴ�,�ɼ�)
            'FormatString������ʱ��ֵ��Ч
            '���AutoResize=True,�������п���и߱��Զ�����(����AutoSizeMode)
            '���WordWrap=True,���и߻ᱻ�Զ�����
            .WordWrap = False
            strFormat = GetColFormat(vsAdvice)
            Call ClearAdviceData
            .ScrollBars = flexScrollBarNone
            Set .DataSource = rsTmp
            .ScrollBars = flexScrollBarBoth
            If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
                gcnOracle.Errors.Clear    '��,��ʱ�̶��д˴���
            End If
            Call SetColFormat(vsAdvice, strFormat)
            .TextMatrix(0, COL_Ƥ��) = ""
            .TextMatrix(0, COL_��ʾ) = ""    'Pass
            .TextMatrix(0, COL_��ʼʱ��) = "��Чʱ��"
            .TextMatrix(0, COL_��) = ""
            '�Զ������и�
            .WordWrap = True

            '����ÿ��ҽ��
            i = .FixedRows
            Do While i <= .Rows - 1
                .Cell(flexcpData, i, COL_������) = CStr(.TextMatrix(i, COL_������))
                .Cell(flexcpData, i, COL_������ӡ) = CStr(.TextMatrix(i, COL_������ӡ))
                .Cell(flexcpFontUnderline, i, COL_������ӡ, i, COL_������ӡ) = True
                .Cell(flexcpData, i, COL_����Ԥ��) = CStr(.TextMatrix(i, COL_����Ԥ��))
                .Cell(flexcpFontUnderline, i, COL_����Ԥ��, i, COL_����Ԥ��) = True
                .Cell(flexcpData, i, COL_����״̬) = Val(.TextMatrix(i, COL_����״̬)) '�������״ֵ̬
                
                '������ʱ��
                If .TextMatrix(i, col_����ʱ��) <> "" Then
                    .Cell(flexcpData, i, col_����ʱ��) = .TextMatrix(i, col_����ʱ��)
                    .TextMatrix(i, col_����ʱ��) = Format(.TextMatrix(i, col_����ʱ��), "yyyy-MM-dd HH:mm")
                End If
                .Cell(flexcpData, i, COL_��ʼʱ��) = CStr(.TextMatrix(i, COL_��ʼʱ��))
                
                If .TextMatrix(i, COL_�������) = "K" And gblnѪ��ϵͳ Then
                    strSQL = "select zl_Get_��Ѫִ��Ѫ��([1]) as Ѫ�� from dual"
                    Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(i, COL_ID)))
                    If Not rsѪ��.EOF Then
                        If rsѪ��!Ѫ�� & "" <> "" Then .TextMatrix(i, COL_Ƥ��) = "(" & rsѪ��!Ѫ�� & ")"
                    End If
                End If

                '��ҩ����ҩ��һЩ����
                bln��ҩ;�� = False: bln��ҩ�÷� = False: bln�ɼ����� = False: bln��Ѫ;�� = False
                If .TextMatrix(i, COL_�������) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ;�� = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '��ʾ��ҩ�ĸ�ҩ;��+����
                                    .TextMatrix(j, COL_�÷�) = .TextMatrix(i, COL_�÷�) & .TextMatrix(i, COL_ҽ������)

                                    If mvarCond.��ʾģʽ = 0 Then    '�ϲ��÷���:�÷� Ƶ�� ����
                                        strFormat = .TextMatrix(j, COL_�÷�)
                                        strTmp = .TextMatrix(j, COL_Ƶ��)
                                        If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                                        strTmp = .TextMatrix(j, COL_����)
                                        If strTmp <> "" Then
                                            strFormat = strFormat & IIF(strFormat <> "", ",", "") & "��" & strTmp & "��"
                                        End If
                                        .TextMatrix(j, COL_�÷�) = strFormat
                                    End If

                                    '��ʾ��ҩ��ִ������
                                    If Val(.TextMatrix(j, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                        .TextMatrix(j, COL_ִ������) = "�Ա�ҩ"
                                    ElseIf Val(.TextMatrix(j, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                        .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(j, COL_ִ������) = ""
                                    End If
                                    
                                    'Σ��ֵID��ֻ��������ҽ�����ģ����Ƶ�ҩƷ����
                                    .TextMatrix(j, COL_Σ��ֵID) = .TextMatrix(i, COL_Σ��ֵID)

                                    If mvarCond.��ʾģʽ = 0 Then
                                        If .TextMatrix(j, COL_Ƥ��) <> "" Then
                                            If Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1") Then
                                                .TextMatrix(j, col_����) = .TextMatrix(j, col_����) & "," & .TextMatrix(j, COL_Ƥ��)
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ�÷� = .TextMatrix(i - 1, COL_�������) = "7"    '��ҩ�÷���
                            bln�ɼ����� = .TextMatrix(i - 1, COL_�������) = "C"    '�ɼ�������
                            
                            If bln��ҩ�÷� Then
                                .TextMatrix(i, COL_������) = .TextMatrix(i - 1, COL_������)
                                .TextMatrix(i, COL_������ӡ) = .TextMatrix(i - 1, COL_������ӡ)
                                .Cell(flexcpData, i, COL_������) = CStr(.TextMatrix(i - 1, COL_������))
                                .Cell(flexcpData, i, COL_������ӡ) = CStr(.TextMatrix(i - 1, COL_������ӡ))
                                .Cell(flexcpFontUnderline, i, COL_������ӡ, i, COL_������ӡ) = True
                                .Cell(flexcpData, i, COL_����Ԥ��) = CStr(.TextMatrix(i - 1, COL_����Ԥ��))
                                .Cell(flexcpFontUnderline, i, COL_����Ԥ��, i, COL_����Ԥ��) = True
                            End If

                            '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
                            If bln�ɼ����� Then
                                j = .FindRow(.TextMatrix(i, COL_ID), .FixedRows, COL_���ID)
                                If j <> -1 Then
                                    .TextMatrix(i, COL_�Թܱ���) = .TextMatrix(j, COL_�Թܱ���)
                                End If
                                .Cell(flexcpData, i, COL_Ƥ��) = .TextMatrix(i, COL_Ƥ��)
                                .TextMatrix(i, COL_Ƥ��) = "" '���������ʱ��ID�����ϲ���ʾ
                            End If

                            '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                            .TextMatrix(i, COL_ִ�п���) = .TextMatrix(i - 1, COL_ִ�п���)

                            If bln��ҩ�÷� Then
                                '��ʾ��ҩ�䷽ִ������
                                If Val(.TextMatrix(i - 1, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                    .TextMatrix(i, COL_ִ������) = "�Ա�ҩ"
                                ElseIf Val(.TextMatrix(i - 1, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                    .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ"
                                Else
                                    .TextMatrix(i, COL_ִ������) = ""
                                End If
                            Else
                                .TextMatrix(i, COL_ִ������) = ""
                            End If

                            'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ;ͬʱ�жϼ�������
                            strTmp = ""
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    .TextMatrix(i, COL_������) = .TextMatrix(j, COL_������)    '���顢�䷽������ҽ��Ϊ׼
                                    .TextMatrix(i, COL_�ļ�ID) = .TextMatrix(j, COL_�ļ�ID)
                                    If bln��ҩ�÷� Then  '��ζ��ҩ��ID��¼������������ҩɾ��ʹ��
                                        strTmp = strTmp & IIF(strTmp = "", .TextMatrix(j, COL_ID), "," & .TextMatrix(j, COL_ID))
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    If bln��ҩ�÷� Then
                                        .Cell(flexcpData, i, COL_���ID) = strTmp
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf .TextMatrix(i - 1, COL_�������) = "K" And Val(.TextMatrix(i - 1, COL_ID)) = Val(.TextMatrix(i, COL_���ID)) Then
                        bln��Ѫ;�� = True
                        '��ʾ��Ѫ;��
                        .TextMatrix(i - 1, COL_�÷�) = .TextMatrix(i, COL_�÷�) & .TextMatrix(i, COL_ҽ������)
                    Else
                        .TextMatrix(i, COL_ִ������) = ""
                    End If
                End If

                '����ɼ��еĵ�һЩ��ʶ:�ſ����ɼ�����ʱδɾ������
                If Not (bln��ҩ;�� Or bln��Ѫ;��) And .TextMatrix(i, COL_�������) <> "7" Then
                    
                    '�иߣ�Ϊ��֧��zl9PrintMode:Resize֮��,ȡRowHeight����С��RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    'ֻ��ʾ��ı����ҽ��
                    If mvarCond.����ģʽ = 3 And Val(.TextMatrix(i, COL_������)) = 0 Then
                        .RowHidden(i) = True: .RowHeight(i) = 0
                    End If
                    
                    '��ʾ���ֱ����ҽ��
                    If mvarCond.����ģʽ = 3 Then
                        If mvarCond.���� = 1 Then ' ���
                            If Not .TextMatrix(i, COL_�������) = "D" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.���� = 2 Then '����
                            If Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "C") Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.���� = 3 Then ' ����
                            If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "D" Or .TextMatrix(i, COL_�������) = "C" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    End If

                    '����С��������,��δ�뵽�취
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    
                    '�����м���ӡ״̬��ʶ
                    Call SetAdviceReportIcon(i)

                    'ҽ����ɫ
                    If Val(.TextMatrix(i, COL_ҽ��״̬)) = 4 Then
                        '������(���ͺ�����)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '��ɫ
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 8 Then
                        '�ѷ���(���ͺ��Զ�ֹͣ)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000    '����
                    End If

                    '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                    If .TextMatrix(i, COL_�������) <> "" Then
                        If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(i, COL_�������)) > 0 Then
                            .Cell(flexcpFontBold, i, col_ҽ������) = True
                            .Cell(flexcpFontBold, i, col_����) = True
                        End If
                    End If

                    'Ƥ�Խ����ʶ
                    If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1" And .TextMatrix(i, COL_Ƥ��) <> "" Then
                        j = GetSkinTestResult(Val(.TextMatrix(i, COL_������ĿID)), .TextMatrix(i, COL_Ƥ��))
                        .Cell(flexcpForeColor, i, COL_Ƥ��) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, COL_Ƥ��))
                    End If

                    '������־:һ����ҩֻ��ʾ�ڵ�һ��
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_��־)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("����").Picture
                        ElseIf Val(.TextMatrix(i, COL_��־)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("��¼").Picture
                        End If

                        'һ����ҩ�ģ�ÿ�е����״̬��������
                        If Val(.TextMatrix(i, COL_ҽ��״̬)) < 2 Then   '�¿����ݴ��ҽ��
                            Select Case Val(.TextMatrix(i, COL_���״̬))
                                '0-������ˣ�1-����ˣ�2-���ͨ����3-���δͨ��
                            Case 1
                                If .TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_��鷽��)) = 1 Then
                                    '��Ѫҽ�����ͼ�굥����ʾ(��������ҽ���˶�)
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�˶�").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                                End If
                            Case 2
                                If Not (.TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_��鷽��)) = 1) Then
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���ͨ��").Picture
                                End If
                            Case 3
                                Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                            Case 4, 5
                                If gblnѪ��ϵͳ = False Then Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                            Case 7
                                Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("��ǩ��").Picture
                            Case Else
                            End Select
                            .Cell(flexcpPictureAlignment, i, COL_F��־) = 4
                        End If
                        '�������ϵͳ
                        If .TextMatrix(i, COL_�������״̬) = "0" Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                        ElseIf .TextMatrix(i, COL_�������״̬) = "2" Or .TextMatrix(i, COL_���������) = "1" Then
                            '��ʱ�������ϸ���
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���ͨ��").Picture
                        ElseIf .TextMatrix(i, COL_���������) = "2" Then
                            ' ���ϸ�
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                        End If
                    End If


                    'Pass:�����������ʾ��ʾ��
                    '
                    If mblnPass Then
                        If .TextMatrix(i, COL_��ʾ) <> "" Then
                            Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, Val(.TextMatrix(i, COL_��ʾ)))
                        End If
                    End If
                    .TextMatrix(i, COL_��ʾ) = ""  '�����ʾֵ
                End If


                If bln��ҩ;�� Or bln��Ѫ;�� Then
                    .RemoveItem i
                Else
                    '���ģʽ�����ҽ������
                    If mvarCond.��ʾģʽ = 0 Then
                        strFormat = .TextMatrix(i, col_ҽ������)
                        If .TextMatrix(i, COL_�������) <> "Z" And Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                            'ҽ�����ݶ����а����������ʱ�������ظ����
                            mrsDefine.Filter = "�������='" & .TextMatrix(i, COL_�������) & "'"
                            If Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1") Then
                                strFormat = strFormat & .TextMatrix(i, COL_Ƥ��)
                            End If

                            If Not (InStr("5,6,7", .TextMatrix(i, COL_�������)) = 0 And .TextMatrix(i, COL_Ƶ��) = "һ����") Then
                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_����)
                                    If strTmp <> "" Then strFormat = strFormat & ",��" & strTmp
                                End If

                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_����)
                                    If strTmp <> "" Then strFormat = strFormat & ",ÿ��" & strTmp
                                End If
                            End If
                        End If
                        .TextMatrix(i, col_����) = strFormat

                        '�ϲ��÷���:�÷� Ƶ�� ����(һ����ҩ����ǰ���Ѵ���)
                        
                        '���ģʽ�³�ҩƷ��������Ŀ��������ҽ������ʾ�÷�
                        If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                            InStr(",5,6,7,", "," & .TextMatrix(i, COL_�������) & ",") > 0 And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            strFormat = .TextMatrix(i, COL_�÷�)
                        Else
                            strFormat = ""
                        End If
                        
                        '���� '��� '��Ѫ '����   ���ģʽ�²���ʾƵ��
                        If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 6 Or _
                            .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                            .TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                            .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            strTmp = ""
                        Else
                            strTmp = .TextMatrix(i, COL_Ƶ��)
                        End If
                         
                        If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                        strTmp = .TextMatrix(i, COL_����)
                        If strTmp <> "" Then
                            strFormat = strFormat & IIF(strFormat <> "", ",", "") & "��" & strTmp & "��"
                        End If

                        .TextMatrix(i, COL_�÷�) = strFormat
                    End If
                    
                    If mvarCond.����ģʽ = 3 Then
                        '����Ǳ���ҳǩ�£����� �� ����Ϊ�գ����¸�ֵ
                        .TextMatrix(i, col_����) = .TextMatrix(i, col_ҽ������)
                        
                        If Val(.TextMatrix(i, COL_����ID)) = 0 And .TextMatrix(i, COL_��鱨��ID) = "" And Val(.TextMatrix(i, COL_RIS����ID)) = 0 And Val(.TextMatrix(i, COL_LIS����ID)) = 0 Then
                            .TextMatrix(i, COL_����״̬) = "δ��"
                        Else
                            .TextMatrix(i, COL_����״̬) = "����"
                            If Val(.Cell(flexcpData, i, COL_����״̬)) = 0 Then  'δ��
                                .Cell(flexcpForeColor, i, COL_����״̬, i, COL_����״̬) = &HFF0000     '��ɫ
                            ElseIf Val(.Cell(flexcpData, i, COL_����״̬)) = 2 Then  '�����Ѷ�
                                .Cell(flexcpForeColor, i, COL_����״̬, i, COL_����״̬) = &HFF00FF     '��ɫ
                            Else
                                .Cell(flexcpForeColor, i, COL_����״̬, i, COL_����״̬) = &H80&     '����
                            End If
                            .Cell(flexcpFontUnderline, i, COL_����״̬, i, COL_����״̬) = True
                        End If
                        '���ӹ���δ���ı�����ѳ��ı���
                        If .RowHidden(i) = False Then
                            If Not IIF(.TextMatrix(i, COL_����״̬) = "δ��", mvarCond.δ������, mvarCond.�ѳ�����) Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    End If
                    i = i + 1
                End If
            Loop
            
            '����ҽ�����ݵ�Ԫ���ͼ��
            For i = 1 To .Rows - 1
                Call SetAdviceIcon(i)
            Next
            
            '�Զ������и�
            If mvarCond.��ʾģʽ = 0 And mvarCond.����ģʽ <> 3 Then
                If InStr("2505,3345,1005,1335", .ColWidth(COL_�÷�)) > 0 Then .ColWidth(COL_�÷�) = IIF(mlngFontSize = 9, 2505, 3345)   '�û�δ�ĸ��п�ʱ������
                .AutoSize col_����, COL_�÷�
                .ColWidth(COL_��ʼʱ��) = IIF(mlngFontSize = 9, 1130, 1510)
            Else
                If InStr("2505,3345,1005,1335", .ColWidth(COL_�÷�)) > 0 Then .ColWidth(COL_�÷�) = IIF(mlngFontSize = 9, 1005, 1335)
                .AutoSize col_ҽ������, COL_�÷�
                .ColWidth(COL_��ʼʱ��) = IIF(mlngFontSize = 9, 1530, 2040)
            End If

            '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '����ǩ��ͼ�����
            .Cell(flexcpPictureAlignment, .FixedRows, col_ҽ������, .Rows - 1, col_ҽ������) = 0
            Call SetTagһ����ҩ
            Call Set�걾״̬
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
    imgColSel.Visible = (mvarCond.��ʾģʽ = 1 And mvarCond.����ģʽ = 0)
    
    If mvarCond.����ģʽ <> 0 Then
        Call Refresh����
    Else
        Call Refresh����
    End If
 
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���ҩ�䷽��
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='7' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs�䷽�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���������
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='C' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPrice(ByVal lngRow As Long) As Boolean
'���ܣ���ȡָ��ҽ���ļƼ�,�����ݵ�ǰ�������շ� ��ϵ���и���
    Dim rs������Ŀ As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strҽ��IDs As String, str�շ�ϸĿIDs As String, str�����շ� As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln�䷽�� As Boolean, bln������ As Boolean, blnLoad As Boolean
    Dim lng���˿���ID As Long, lngִ�п���ID As Long
    Dim dblPrice As Double, lng����ID As Long
    Dim lngҽ��ID As Long, lng���ID As Long
    Dim strPriceType As String
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
        
        lngҽ��ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
        lng���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
                                    
        blnLoad = True
        
        'ҩƷ�����ĵļƼ�
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "4" Then
            '���ĵļƼ�
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������,0 as �շѷ�ʽ," & _
                " A.�շ�ϸĿID,1 as �����װ,C.���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,Nvl(B.����,D.ȱʡ�۸�),D.�ּ�) as ����,A.ִ�п���ID,0 as ����" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=B.ҽ��ID(+) And A.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "4", "5", "6") & _
                " And C.������� IN(1,3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
                
                blnLoad = False
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            '��,����ҩ:���ܰ������ҽ��,���������װ�ĵ���
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������,0 as �շѷ�ʽ," & _
                " C.ID as �շ�ϸĿID,B.�����װ,B.���ﵥλ as ���㵥λ,A.�ܸ����� as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.�����װ as ����," & _
                " A.ִ�п���ID,0 as ����" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID And Nvl(A.ִ������,0) Not IN(0,5)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "4", "5", "6") & _
                " And (A.�շ�ϸĿID is NULL Or A.�շ�ϸĿID=B.ҩƷID)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(1,3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
                
                '��һ����ҩ(�����)�ĵ�һ��ҩ�в���ʾ��ҩ;���ļƼ�
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_���ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
        ElseIf bln�䷽�� Then
            '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������,0 as �շѷ�ʽ," & _
                " C.ID as �շ�ϸĿID,B.�����װ,B.���ﵥλ as ���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.�����װ as ����," & _
                " A.ִ�п���ID,0 as ����" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.�������='7' And A.���ID=[1]" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID And C.������� IN(1,3)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "4", "5", "6") & _
                " And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        End If
        
        '��ȡ���мƼ�(ȡ���¼۸�)����ҩƷ��������ļƼ�,�������ҽ���Ƽ�
        '���Ƽ�,�ֹ��Ƽ۵�ҽ������ȡ
        '��Union��ʽ������������
        If blnLoad Then
            '�����¿���ҽ�������ݲ���ҽ���Ƽ���ȡ
            If InStr(",1,2,-1,", vsAdvice.TextMatrix(lngRow, COL_ҽ��״̬)) = 0 Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0) as ��������,Nvl(B.�շѷ�ʽ,0) as �շѷ�ʽ," & _
                    " B.�շ�ϸĿID,1 as �����װ,C.���㵥λ,B.����,Decode(C.�Ƿ���,1,B.����,Sum(D.�ּ�)) as ����," & _
                    " Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,Nvl(B.����,0) as ����" & _
                    " From ����ҽ����¼ A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                    " Where A.������� Not IN('4','5','6','7') And A.ID=B.ҽ��ID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "4", "5", "6") & _
                    " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ�ϸĿID=C.ID And B.�շ�ϸĿID=D.�շ�ϸĿID" & _
                    " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                    " And (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1])" & _
                    " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0),Nvl(B.�շѷ�ʽ,0)," & _
                    " B.�շ�ϸĿID,C.���㵥λ,B.����,C.�Ƿ���,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID),Nvl(B.����,0)"
            Else
                '�¿���ҽ�������������շ� ��ϵ��ȡ(��ҩ�����ʾΪ0)
                '���ֶ�Ӧ�ļƼۣ�
                '   1.���յķ��ã�ֻ������Ŀ������գ�Ŀǰֻ�д��Ի������������
                '   2.�����ķ��ã����Ǿ���ļ�鲿λ�ͼ�鷽����
                '   3.�����ķ��ã��Ǽ�鲿λ�ͷ�����(ע�����걾��д�ڱ걾��λ��)
                lng����ID = 0 '�����Թܷ���,ֻ��ȡ�Թܶ�Ӧ�����ķ���
                If vsAdvice.TextMatrix(lngRow, COL_�Թܱ���) <> "" Then
                    lng����ID = GetTubeMaterial(vsAdvice.TextMatrix(lngRow, COL_�Թܱ���))
                End If
                
                str�����շ� = "Select * From (" & _
                    "Select C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,C.���ÿ���id" & _
                    " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                    " From �����շѹ�ϵ C,����ҽ����¼ A Where (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1]) And A.������ĿID+0=C.������ĿID" & _
                    "   And (a.���id Is Null And a.ִ�б�� In (1, 2) And c.�������� = 1 Or" & vbNewLine & _
                    "   a.�걾��λ = c.��鲿λ And a.��鷽�� = c.��鷽�� And Nvl(c.��������, 0) = 0 Or" & vbNewLine & _
                    "   (a.��鷽�� Is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(c.��������, 0) = 0 And c.��鲿λ Is Null And c.��鷽�� Is Null)" & _
                    "      And (C.���ÿ���ID is Null or C.���ÿ���ID = A.ִ�п���ID And C.������Դ = 1)" & _
                    " ) Where Nvl(���ÿ���id, 0) = Top"
                
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0) as ��������,Nvl(B.�շѷ�ʽ,0) as �շѷ�ʽ," & _
                    " B.�շ���ĿID as �շ�ϸĿID,1 as �����װ,C.���㵥λ,B.�շ����� as ����,Decode(C.�Ƿ���,1,Sum(D.ȱʡ�۸�),Sum(D.�ּ�)) as ����," & _
                    " A.ִ�п���ID,Nvl(B.������Ŀ,0) as ����" & _
                    " From ����ҽ����¼ A,(" & str�����շ� & ") B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                    " Where A.������� Not IN('4','5','6','7') And A.ҽ��״̬ IN(-1,1,2) And A.������ĿID+0=B.������ĿID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "4", "5", "6") & _
                    " And (A.���ID is Null And A.ִ�б�� IN(1,2) And B.��������=1" & _
                    "       Or A.�걾��λ=B.��鲿λ And A.��鷽��=B.��鷽�� And Nvl(B.��������,0)=0" & _
                    "       Or (A.��鷽�� is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(B.��������,0)=0 And B.��鲿λ is Null And B.��鷽�� is Null)" & _
                    " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
                    " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                    " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And C.������� IN(1,3)" & _
                    " And (Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And B.�շ���ĿID=[3] Or Not(Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And [3]<>0))" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) And (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1])" & _
                    " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0),Nvl(B.�շѷ�ʽ,0)," & _
                    " B.�շ���ĿID,C.���㵥λ,B.�շ�����,C.�Ƿ���,A.ִ�п���ID,Nvl(B.������Ŀ,0)"
            End If
        End If
        strSQL = strSQL & " Order by ���,��������,����"
        
        If mblnMoved Then '�Һŵ���ҽ����ͬ�����ݿ�
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ���Ƽ�", "H����ҽ���Ƽ�")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID, lng���ID, lng����ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        
        '��ʾ�Ƽ�����
        If Not rsTmp.EOF Then
            'ȷ����ʾ����
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '��ȡ������Ŀ,�շ�ϸĿ��Ϣ
            For i = 1 To rsTmp.RecordCount
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ID & ",") = 0 Then strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
                If InStr("," & str�շ�ϸĿIDs & ",", "," & rsTmp!�շ�ϸĿID & ",") = 0 Then str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & rsTmp!�շ�ϸĿID
                rsTmp.MoveNext
            Next
            strҽ��IDs = Mid(strҽ��IDs, 2)
            str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 2)
                        
            strSQL = "Select/*+ Rule*/ B.ID,B.���,C.���� as �������,B.����,B.�걾��λ" & _
                " From ����ҽ����¼ A,������ĿĿ¼ B,������Ŀ��� C,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) D" & _
                " Where A.ID = D.Column_Value And A.������ĿID=B.ID And B.���=C.����"
                
            If mblnMoved Then '�Һŵ���ҽ����ͬ�����ݿ�
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strҽ��IDs) 'In
            
            strSQL = "Select A.ID,A.���,B.���� as �������,A.����," & _
                " A.����,A.���,A.����,A.��������,A.�Ƿ���" & _
                " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) D" & _
                " Where A.���=B.���� And A.ID = D.Column_Value"
            strSQL = "Select/*+ Rule*/ A.ID,A.���,A.�������,A.����,Nvl(B.����,A.����) as ����," & _
                " A.���,A.����,A.��������,A.�Ƿ���,C.��������" & _
                " From (" & strSQL & ") A,�շ���Ŀ���� B,�������� C" & _
                " Where A.ID=C.����ID(+) And A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[2]"
            Set rs�շ�ϸĿ = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str�շ�ϸĿIDs, IIF(gbytҩƷ������ʾ = 0, 1, 3))
            
            '��ʾÿ������
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs������Ŀ.Filter = "ID=" & rsTmp!������ĿID
                rs�շ�ϸĿ.Filter = "ID=" & rsTmp!�շ�ϸĿID
                
                '�Ƽ�ҽ��
                If rsTmp!������� = "4" Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��������-" & rs������Ŀ!����
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "ҩƷҽ��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��ҩ;��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��Ѫ;��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "�ɼ�����-" & rs������Ŀ!����
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��ҩ�巨-" & rs������Ŀ!����
                    Else
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��ҩ�÷�-" & rs������Ŀ!����
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "������Ŀ-" & rs������Ŀ!����
                    ElseIf rsTmp!������� = "D" Then
                        '��λ������
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��鲿λ-" & NVL(rsTmp!�걾��λ) & "(" & NVL(rsTmp!��鷽��) & ")"
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��������-" & rs������Ŀ!����
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "������Ŀ-" & rs������Ŀ!����
                    End If
                Else
                    If NVL(rsTmp!��������, 0) = 1 Then
                        '���Ի����м��շ���
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!���� & "(" & Decode(NVL(rsTmp!ִ�б��, 0), 1, "����", 2, "����", "") & "����)"
                    Else
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!����
                    End If
                End If
                
                '���
                .TextMatrix(i, COLPrice("���")) = rs�շ�ϸĿ!�������
                '�շ���Ŀ:���/����
                .TextMatrix(i, COLPrice("�շ���Ŀ")) = rs�շ�ϸĿ!����
                If Not IsNull(rs�շ�ϸĿ!����) Then
                    .TextMatrix(i, COLPrice("�շ���Ŀ")) = .TextMatrix(i, COLPrice("�շ���Ŀ")) & "(" & rs�շ�ϸĿ!���� & ")"
                End If
                If Not IsNull(rs�շ�ϸĿ!���) Then
                    .TextMatrix(i, COLPrice("�շ���Ŀ")) = .TextMatrix(i, COLPrice("�շ���Ŀ")) & " " & rs�շ�ϸĿ!���
                End If
                
                '���㵥λ:ҩ��ҩƷΪ���ﵥλ,��ҩ��ҩƷΪ�ۼ۵�λ
                .TextMatrix(i, COLPrice("��λ")) = NVL(rsTmp!���㵥λ)
                '�Ƽ�����:ҩ��ҩƷΪ1,��ҩ��ҩƷΪ��Ӧ�ۼ���
                If InStr(",5,6,7,", rs������Ŀ!���) > 0 Then
                    .TextMatrix(i, COLPrice("�Ƽ�����")) = 1
                Else
                    .TextMatrix(i, COLPrice("�Ƽ�����")) = FormatEx(rsTmp!����, 5)
                End If
                'ִ�п���
                lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                If rs�շ�ϸĿ!��� = "4" And NVL(rs�շ�ϸĿ!��������, 0) = 1 _
                    Or InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 And InStr(",5,6,7,", rs������Ŀ!���) = 0 Then
                    lng���˿���ID = mlng�Һſ���ID
                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rs�շ�ϸĿ!���, rs�շ�ϸĿ!ID, 4, lng���˿���ID, 0, 1, lngִ�п���ID)
                End If
                
                '���۴���
                If InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 Then
                    If NVL(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                        '��ҩƷʱ��
                        If InStr(",5,6,7,", rs������Ŀ!���) > 0 Then
                            'ҩ��ҩƷ���������װ������ʱ��
                            .TextMatrix(i, COLPrice("����")) = CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rsTmp!����, 1), , , 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            .TextMatrix(i, COLPrice("����")) = Format(Val(.TextMatrix(i, COLPrice("����"))) * NVL(rsTmp!�����װ, 0), gstrDecPrice)
                        Else
                            '��ҩ��ҩƷ��������ۼ��������ۼ�ʵ��
                            .TextMatrix(i, COLPrice("����")) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rsTmp!����, 0), , , 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        End If
                    Else
                        'ҩ��ҩƷΪ���ﵥ��,��ҩҩƷΪ�ۼ�
                        .TextMatrix(i, COLPrice("����")) = Format(NVL(rsTmp!����), gstrDecPrice)
                    End If
                ElseIf rs�շ�ϸĿ!��� = "4" And NVL(rs�շ�ϸĿ!��������, 0) = 1 And NVL(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                    'ʱ�����ĵĵ��ۺ�ҩƷһ������
                    .TextMatrix(i, COLPrice("����")) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rsTmp!����, 0), , , 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                Else
                    .TextMatrix(i, COLPrice("����")) = Format(NVL(rsTmp!����), gstrDecPrice)
                End If
                
                'ִ�п���
                If lngִ�п���ID <> 0 Then
                    .TextMatrix(i, COLPrice("ִ�п���")) = Sys.RowValue("���ű�", lngִ�п���ID, "����")
                End If
                
                '��ʾҽ����������
                If Val(rsTmp!�շ�ϸĿID & "") <> 0 Then
                    strPriceType = GetPriceType(mlng����ID, Val(rsTmp!�շ�ϸĿID & ""), mint����, True)
                End If
                '��������
                If strPriceType = "" Then
                    .TextMatrix(i, COLPrice("��������")) = NVL(rs�շ�ϸĿ!��������)
                Else
                    .TextMatrix(i, COLPrice("��������")) = strPriceType
                End If
                
                '������Ŀ
                .TextMatrix(i, COLPrice("����")) = IIF(NVL(rsTmp!����, 0) = 0, "", "��")
                
                '�շѷ�ʽ
                .TextMatrix(i, COLPrice("�շѷ�ʽ")) = getChargeMode(Val(NVL(rsTmp!�շѷ�ʽ, 0)))
                
                dblPrice = dblPrice + Format(Val(.TextMatrix(i, COLPrice("�Ƽ�����"))) * Val(.TextMatrix(i, COLPrice("����"))), "0.00000")
                
                rsTmp.MoveNext
            Next
        End If
        
        '�ϼ���
        If .Rows > 2 Then
            .MergeCol(COLPrice("�Ƽ�ҽ��")) = True
            .MergeCol(COLPrice("���")) = True
            
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, COLPrice("�Ƽ�ҽ��"), .Rows - 1, COLPrice("��λ")) = "�ϼ�"
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("�Ƽ�ҽ��"), .Rows - 1, COLPrice("��λ")) = 4
            .Cell(flexcpText, .Rows - 1, COLPrice("�Ƽ�����"), .Rows - 1, COLPrice("����")) = Format(dblPrice, gstrDecPrice)
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("�Ƽ�����"), .Rows - 1, COLPrice("����")) = 7
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
            
        End If
        
        .Row = 1: .Col = 0
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    ShowPrice = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ���ķ��ͼ�¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    Dim bln״̬˵�� As Boolean
    Dim lng��Ѫ As Long
    Dim j As Long
    
    On Error GoTo errH
    lng��Ѫ = -1
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��')"
        strExe2 = "Decode(Nvl(B.ִ��״̬,0),0,'δִ��',1,'ִ�����',2,'�ܾ�ִ��',3,'����ִ��')"
        strState = "Decode(A.ִ��״̬,9,'�շ��쳣',Decode(A.��¼����,1,Decode(A.��¼״̬,0,'�շѻ���',1,'���շ�',3,'���˷�'),2,Decode(A.��¼״̬,0,'���ʻ���',1,'�Ѽ���',3,'������'),'δ�Ʒ�'))"
        
        'ҩ����Ӧ��ҩƷ�Ƽ۰������װ��ʾ,��ҩ����Ӧ��ҩƷ�Ƽ۰����۵�λ��ʾ
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            If Not RowInһ����ҩ(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '��ҩ����:��д�˷��ͼ�¼,�������޶�Ӧ����(���Ա�ҩ,��ҽ���й��)
            strSub = "Select a.ҽ�����,MIN(a.��¼����) AS ��¼���� ,A.NO,A.ִ��״̬,Min(A.��¼״̬) as ��¼״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,B.�����װ,B.���ﵥλ" & _
                " From ������ü�¼ A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL And A.�շ���� IN('5','6','7')" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.ҽ�����=[1] Group By a.ҽ�����,A.NO,A.ִ��״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,b.�����װ, b.���ﵥλ"
            If mblnMoved Then
                strSub = Replace(strSub, "������ü�¼", "H������ü�¼")
            ElseIf zlDatabase.DateMoved(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "������ü�¼", "H������ü�¼")
            End If
            
            strSQL = _
                " Select C.���ID,C.�걾��λ,C.��鷽��,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                " Nvl(A.���ﵥλ,D.���ﵥλ) as ��λ," & _
                " Nvl(A.����/Nvl(A.�����װ,1),B.��������/Nvl(D.����ϵ��,1)/Nvl(D.�����װ,1)) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬,B.�״�ʱ��,B.ĩ��ʱ��," & _
                " Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�'," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.״̬˵��,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������,B.���ʱ��,B.�����,B.ִ��˵��" & _
                " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,ҩƷ��� D" & _
                " Where B.ҽ��ID=C.ID And C.�շ�ϸĿID=D.ҩƷID And C.ID=[1]" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And A.ҽ�����(+)=B.ҽ��ID"
            
            '��һ����ҩ�����в���ʾ��ҩ;���ķ���
            If lngRow = lngBegin Then
                '��ҩ;������:��д�˷��ͼ�¼(������),����һ���з���
                strSub = "Select a.ҽ�����,MIN(a.��¼����) AS ��¼���� ,A.NO,A.ִ��״̬,Min(A.��¼״̬) as ��¼״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,B.�����װ,B.���ﵥλ" & _
                    " From ������ü�¼ A,ҩƷ��� B" & _
                    " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                    " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[2] Group By a.ҽ�����,A.NO,A.ִ��״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,b.�����װ, b.���ﵥλ"
                If mblnMoved Then
                    strSub = Replace(strSub, "������ü�¼", "H������ü�¼")
                ElseIf zlDatabase.DateMoved(mvRegDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, "������ü�¼", "H������ü�¼")
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select C.���ID,C.�걾��λ,C.��鷽��,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                    " Decode(Nvl(Instr('567',A.�շ����),0),0,Decode(A.�շ����,'4',F.���㵥λ,D.���㵥λ),Nvl(A.���ﵥλ,E.���ﵥλ)) as ��λ," & _
                    " Nvl(A.����/Nvl(A.�����װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.�����װ,1)) as ��������," & _
                    " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID," & _
                    " Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬,B.�״�ʱ��," & _
                    " B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�'," & strState & ") as �Ʒ�״̬," & _
                    " B.������,B.״̬˵��,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������,B.���ʱ��,B.�����,B.ִ��˵��" & _
                    " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D,ҩƷ��� E,�շ���ĿĿ¼ F" & _
                    " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+) And C.�շ�ϸĿID=F.ID(+)" & _
                    " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID And C.ID=[2]"
            End If
            
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        Else
            '����ҽ��(�������ġ��䷽����飬����һ��ҽ��):��д�˷��ͼ�¼(������),����һ���з���
            '��ҩ�Ա�ҩҲ���޶�Ӧ����(��ҽ���й��)
            strSub = _
                " Select a.ҽ�����,MIN(a.��¼����) AS ��¼���� ,A.NO,A.ִ��״̬,Min(A.��¼״̬) as ��¼״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,B.�����װ,B.���ﵥλ" & _
                " From ������ü�¼ A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[1]  Group By a.ҽ�����,A.NO,A.ִ��״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,b.�����װ, b.���ﵥλ"
            strSub = strSub & " Union ALL " & _
                " Select a.ҽ�����,MIN(a.��¼����) AS ��¼���� ,A.NO,A.ִ��״̬,Min(A.��¼״̬) as ��¼״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,B.�����װ,B.���ﵥλ" & _
                " From ������ü�¼ A,ҩƷ��� B,����ҽ����¼ C" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=C.ID" & _
                " And C.���ID=[1]  Group By a.ҽ�����,A.NO,A.ִ��״̬,A.���,A.ִ�в���ID,a.�շ����,A.����,A.����, a.�շ�ϸĿid,b.�����װ, b.���ﵥλ"
            If mblnMoved Then
                strSub = Replace(strSub, "������ü�¼", "H������ü�¼")
            ElseIf zlDatabase.DateMoved(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "������ü�¼", "H������ü�¼")
            End If
            
            strSQL = _
                " Select * From ����ҽ����¼ Where ID=[1]" & _
                " Union ALL " & _
                " Select * From ����ҽ����¼ Where ���ID=[1]"
            strSQL = _
                " Select C.���ID,C.�걾��λ,C.��鷽��,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                " Decode(Nvl(Instr('567',A.�շ����),0),0,Decode(A.�շ����,'4',F.���㵥λ,D.���㵥λ),Nvl(A.���ﵥλ,E.���ﵥλ)) as ��λ," & _
                " Nvl(Nvl(A.����,1)*A.����/Nvl(A.�����װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.�����װ,1)) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬,B.�״�ʱ��,B.ĩ��ʱ��," & _
                " Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�'," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.״̬˵��,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������,B.���ʱ��,B.�����,B.ִ��˵��" & _
                " From (" & strSub & ") A,����ҽ������ B,(" & strSQL & ") C,������ĿĿ¼ D,ҩƷ��� E,�շ���ĿĿ¼ F" & _
                " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+) And C.�շ�ϸĿID=F.ID(+)" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        End If
        
        strSQL = "Select /*+ RULE */ A.�������,A.�������," & _
            " A.���ID,A.�������,F.���� as �������,D.���� as ������Ŀ,A.�걾��λ,A.��鷽��,A.����ʱ��,A.NO,A.��¼����," & _
            " Nvl(G.����,B.����)||Decode(B.����,NULL,NULL,'('||B.����||')')||Decode(B.���,NULL,NULL,' '||B.���) as �շ���Ŀ," & _
            " A.��λ,A.�������� as ����,C.���� as ִ�п���,A.ִ��״̬,A.�״�ʱ��,A.ĩ��ʱ��,A.�Ʒ�״̬,A.������,A.״̬˵��,A.���ͺ�,A.���ʱ��,A.�����,A.ִ��˵��" & _
            " From (" & strSQL & ") A,�շ���ĿĿ¼ B,���ű� C,������ĿĿ¼ D,������Ŀ��� F,�շ���Ŀ���� G" & _
            " Where A.�շ�ϸĿID=B.ID(+) And A.ִ�в���ID=C.ID(+)" & _
            " And A.������ĿID=D.ID And A.�������=F.����" & _
            " And A.�շ�ϸĿID=G.�շ�ϸĿID(+) And G.����(+)=1 And G.����(+)=" & IIF(gbytҩƷ������ʾ = 0, 1, 3) & _
            " Order by A.���ͺ� Desc,A.�������,A.�������,A.�������"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, COLSend("���ͺ�")) = NVL(rsTmp!���ͺ�, 0)
                .TextMatrix(i, COLSend("����ʱ��")) = Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
                
                '����ҽ��
                If rsTmp!������� = "4" Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "��������-" & rsTmp!������Ŀ
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "ҩƷҽ��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "��ҩ;��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "��Ѫ;��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "�ɼ�����-" & rsTmp!������Ŀ
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "��ҩ�巨-" & rsTmp!������Ŀ
                    Else
                        .TextMatrix(i, COLSend("����ҽ��")) = "��ҩ�÷�-" & rsTmp!������Ŀ
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "������Ŀ-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "D" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "��鲿λ-" & NVL(rsTmp!�걾��λ) & "(" & NVL(rsTmp!��鷽��) & ")"
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "��������-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "������Ŀ-" & rsTmp!������Ŀ
                    End If
                Else
                    .TextMatrix(i, COLSend("����ҽ��")) = rsTmp!������� & "ҽ��-" & rsTmp!������Ŀ
                End If
               
                .TextMatrix(i, COLSend("���ݺ�")) = NVL(rsTmp!NO)
                .TextMatrix(i, COLSend("�շ���Ŀ")) = NVL(rsTmp!�շ���Ŀ)
                .TextMatrix(i, COLSend("��������")) = FormatEx(NVL(rsTmp!����), 5) & NVL(rsTmp!��λ)
                .TextMatrix(i, COLSend("�Ʒ�״̬")) = NVL(rsTmp!�Ʒ�״̬)
                If rsTmp!״̬˵�� & "" <> "" Then
                    bln״̬˵�� = True
                End If
                .TextMatrix(i, COLSend("ִ��״̬")) = NVL(rsTmp!ִ��״̬)
                .TextMatrix(i, COLSend("ִ�п���")) = NVL(rsTmp!ִ�п���)
                .TextMatrix(i, COLSend("������")) = NVL(rsTmp!������)
                .TextMatrix(i, COLSend("״̬˵��")) = NVL(rsTmp!״̬˵��)
                .TextMatrix(i, COLSend("��¼����")) = NVL(rsTmp!��¼����)
                .TextMatrix(i, COLSend("ִ��ʱ��")) = Format(NVL(rsTmp!���ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("ִ����")) = NVL(rsTmp!�����)
                .TextMatrix(i, COLSend("ִ��˵��")) = NVL(rsTmp!ִ��˵��)
                
                '���շѵĻ��۵�ͻ����ʾ
                If .TextMatrix(i, COLSend("�Ʒ�״̬")) = "�ѽɷ�" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '����
                ElseIf .TextMatrix(i, COLSend("�Ʒ�״̬")) = "���˷�" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H808080 '��ɫ
                End If
                If vsAdvice.TextMatrix(lngRow, COL_�������) = "K" And rsTmp!������� & "" = "K" Then
                    If gblnѪ��ϵͳ Then
                        lng��Ѫ = i
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        
        If lng��Ѫ <> -1 Then
            '��Ѫҽ������������Ŀ����Ϣ
            strSQL = "select b.���� as ������Ŀ,a.������ as ����,b.���㵥λ as ��λ,a.������Ŀid from ��Ѫ������Ŀ a,������ĿĿ¼ b where a.������Ŀid=b.id and a.ҽ��id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!������ĿID & "") <> Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID)) Then
                    .AddItem ""
                    For j = .FixedCols To .Cols - 1
                        .TextMatrix(.Rows - 1, j) = .TextMatrix(lng��Ѫ, j)
                    Next
                    .TextMatrix(.Rows - 1, COLSend("����ҽ��")) = "��Ѫҽ��-" & rsTmp!������Ŀ
                    .TextMatrix(.Rows - 1, COLSend("��������")) = FormatEx(NVL(rsTmp!����), 5) & NVL(rsTmp!��λ)
                    .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lng��Ѫ, .FixedCols)
                Else
                    .TextMatrix(lng��Ѫ, COLSend("��������")) = FormatEx(NVL(rsTmp!����), 5) & NVL(rsTmp!��λ)
                End If
                rsTmp.MoveNext
            Next
        End If
        
        .MergeCells = flexMergeFree
        .MergeCol(COLSend("���ͺ�")) = True
        .MergeCol(COLSend("����ʱ��")) = True
        .MergeCol(COLSend("���ݺ�")) = True
        .MergeCol(COLSend("����ҽ��")) = True
        .MergeCol(COLSend("�շ���Ŀ")) = True
        .MergeCol(COLSend("ִ��ʱ��")) = True
        .MergeCol(COLSend("ִ��˵��")) = True
        .MergeCol(COLSend("������")) = True
        .MergeCol(COLSend("״̬˵��")) = True
        
        .ColHidden(COLSend("״̬˵��")) = Not bln״̬˵��
        .Row = 1: .Col = COLSend("����ҽ��")
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ����ǩ����¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.ǩ��ID,A.��������,B.ǩ��ʱ��,B.ǩ����,B.ʱ���," & _
            " Decode(A.��������,1,'�¿�ҽ��',4,'����ҽ��','��������') as ǩ������" & _
            " From ����ҽ��״̬ A,ҽ��ǩ����¼ B Where A.ҽ��ID=[1] And A.ǩ��ID=B.ID Order by B.ǩ��ʱ��"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!ǩ��ID)
                .TextMatrix(i, 0) = rsTmp!ǩ������
                .Cell(flexcpData, i, 0) = Val(rsTmp!��������)
                .TextMatrix(i, 1) = Format(rsTmp!ǩ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!ǩ����
                .TextMatrix(i, 3) = Format(rsTmp!ʱ���, "yyyy-MM-dd HH:mm:ss")
                Set .Cell(flexcpPicture, i, 0) = frmIcons.imgSign.ListImages("ǩ��").Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBillAppend(ByVal lngRow As Long, Optional blnExist As Boolean) As Boolean
'���ܣ���ʾָ����ҽ���ĵ��ݸ�������
'���أ�blnExist=ҽ���Ƿ���ڵ��ݸ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    
    blnExist = False
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    strSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order by ����"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp!��Ŀ & "��" & NVL(rsTmp!����)
                lngIdx = .Find(rsTmp!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp!��Ŀ & "��")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            '��궨λ�ڵ�һ�����븽��
            rsTmp.MoveFirst
            lngIdx = .Find(rsTmp!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp!��Ŀ & "��")
            
            Call SetRTFFont(1)
        End With
        blnExist = True
    End If
    
    ShowBillAppend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowAdvicePlan(ByVal lngRow As Long, Optional blnExist As Boolean) As Boolean
'���ܣ���ʾָ����ҽ����ִ�а�����Ϣ
'���أ�blnExist=ҽ���Ƿ����ִ�а�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    blnExist = False
    rtfInfo.Text = "": rtfInfo.SelStart = 0
    
    On Error GoTo errH
    
    With vsAdvice
        If InStr("D,F,G,", .TextMatrix(lngRow, COL_�������)) > 0 Or _
            .TextMatrix(lngRow, COL_�������) = "E" And InStr(",0,6,", "," & .TextMatrix(lngRow, COL_��������) & ",") > 0 Then
            
            If .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_��������)) = 6 Then
                strSQL = "Select a.����ʱ��,a.ִ�м�,a.ִ��˵�� From ����ҽ������ a,����ҽ����¼ b " & _
                        "Where a.ҽ��ID = b.ID And b.���ID=[1] And (a.ִ��˵�� is Not Null Or a.����ʱ�� is Not Null) And Rownum=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            Else
                strSQL = "Select ����ʱ��,ִ�м�,ִ��˵�� From ����ҽ������ Where ҽ��ID=[1] And (ִ��˵�� is Not Null Or ����ʱ�� is Not Null)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            End If
            
            If Not rsTmp.EOF Then
                strSQL = ""
                
                If Not IsNull(rsTmp!����ʱ��) Then
                    strSQL = strSQL & vbCrLf & "����ʱ�䣺" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                End If
                If Not IsNull(rsTmp!ִ�м�) Then
                    strSQL = strSQL & vbCrLf & "ִ�м䣺" & rsTmp!ִ�м�
                End If
                strSQL = strSQL & vbCrLf & NVL(rsTmp!ִ��˵��)
                
                rtfInfo.Text = Mid(strSQL, 3)
                
                Call SetRTFFont(2)
                blnExist = True
            End If
        End If
    End With
    ShowAdvicePlan = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowOtherAppend(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ���������Ϣ
'˵����ֻ������״̬ͨ����δͨ����ҽ��
'���أ��Ƿ���������Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim int���� As Integer
    Dim str����Ա As String
    Dim strʱ�� As String
    
    strSQL = "Select ������Ա,����ʱ�� From ����ҽ��״̬ Where ҽ��id = [1] And �������� = [2]"
    
    str����Ա = "����ˣ�": strʱ�� = "���ʱ�䣺"
    Select Case vsAdvice.TextMatrix(lngRow, COL_���״̬)
        Case 2
            If gblnѪ��ϵͳ And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then '��Ѫҽ���������̱䶯 70823
                int���� = 15 'Ѫ�����ͨ��
                str����Ա = "Ѫ������ˣ�"
                strʱ�� = "Ѫ�����ʱ�䣺"
            Else
                int���� = 11
            End If
        Case 3
            int���� = 12
        Case 4
            int���� = 11
        Case 5
            int���� = 14
            str����Ա = "Ѫ������ˣ�"
            strʱ�� = "Ѫ�����ʱ�䣺"
    End Select
    rtfOther.Text = ""
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), int����)
    If Not rsTmp.EOF Then
        strSQL = ""
        Do While Not rsTmp.EOF
            strSQL = IIF(strSQL = "", "", strSQL & vbCrLf) & str����Ա & rsTmp!������Ա & vbCrLf & _
                strʱ�� & Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
            rsTmp.MoveNext
        Loop
        rtfOther.Text = strSQL
        Call SetRTFFont(3)
        ShowOtherAppend = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadBillList() As Boolean
'���ܣ���ʾָ���е�ҽ�����Ϳ��Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objpopup1 As CommandBarPopup
    
    If mcbsMain Is Nothing Then LoadBillList = True: Exit Function
    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True)
    If objPopup Is Nothing Then LoadBillList = True: Exit Function
    Set objBar = mcbsMain(2)
    If objBar Is Nothing Then LoadBillList = True: Exit Function
    objPopup.Visible = True
    
    objPopup.CommandBar.Controls.DeleteAll
    
    If mcbsMain Is Nothing Then LoadBillList = True: Exit Function
    Set objMenu = mcbsMain.FindControl(, conMenu_EditPopup, False, True)
    If objMenu Is Nothing Then LoadBillList = True: Exit Function
    Set objpopup1 = objMenu.CommandBar.FindControl(, conMenu_Report_ClinicBill)
    objpopup1.Visible = True
    For i = objMenu.CommandBar.Controls.Count To 1 Step -1
        If objMenu.CommandBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objMenu.CommandBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objMenu.CommandBar.Controls(i).Delete
        End If
    Next
    For i = objBar.Controls.Count To 1 Step -1
        If objBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objBar.Controls(i).Delete
        End If
    Next
    For i = objpopup1.CommandBar.Controls.Count To 1 Step -1
        If objpopup1.CommandBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objpopup1.CommandBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objpopup1.CommandBar.Controls(i).Delete
        End If
    Next
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 _
         Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 8 Then
        LoadBillList = True: Exit Function
    End If
        
    On Error GoTo errH
    
    Set rsTmp = GetBillList

    '���ֻ��һ�����Ƶ��ݿ��ã���ֱ�Ӽ��뵽ҽ���˵���
    If rsTmp.RecordCount = 1 Then
        objPopup.Visible = False
        objPopup.Category = "���ж�"
        objpopup1.Visible = False
        objpopup1.Category = "���ж�"
        Set objPopup = objMenu
    End If
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, IIF(rsTmp.RecordCount = 1, "��ӡ:", "") & rsTmp!����)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                '��ҩ�ļ巨�÷����ݺź���ҩ��һ������������ʾ����ҩ�÷������԰ѵ��ݵ�NOƴ��ȥ
                objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" & "|" & rsTmp!NO  '��Ӧ���Զ��屨����
                'If i > 1 Then objControl.Enabled = False 'һ����Ŀֻ������һ�����Ƶ���
            End With
            '�˵��͹�����Ҫ�ֿ���
            If rsTmp.RecordCount > 1 Then
                With objpopup1.CommandBar.Controls
                    Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, rsTmp!����)
                    If i <= 10 Then
                        objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                    ElseIf i <= 36 Then
                        objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                    End If
                    objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" & "|" & rsTmp!NO  '��Ӧ���Զ��屨����
                End With
            End If
            If rsTmp.RecordCount = 1 Then
                With objBar.Controls
                    Set objControl = .Find(, conMenu_Report_ClinicBill)
                    If Not objControl Is Nothing Then
                        Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, "��ӡ����", objControl.Index + 1)
                        If i <= 10 Then
                            objControl.Caption = objControl.Caption
                        ElseIf i <= 36 Then
                            objControl.Caption = objControl.Caption
                        End If
                        objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" & "|" & rsTmp!NO  '��Ӧ���Զ��屨����
                        objControl.IconId = conMenu_File_Print
                        objControl.Style = xtpButtonIconAndCaption
                        'If i > 1 Then objControl.Enabled = False 'һ����Ŀֻ������һ�����Ƶ���
                    End If
                End With
            End If
            rsTmp.MoveNext
        Next
    End If
    
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsAppend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    
End Sub

Private Sub vsAppend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    With vsAppend
        If Button = 2 And tbcAppend.Selected.Tag = "ǩ��" Then
            If mcbsMain Is Nothing Then Exit Sub
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    Set objPopup = mcbsMain.FindControl(, conMenu_Tool_Sign, False, True) '���ܹ���������
                    If Not objPopup Is Nothing Then
                        If objPopup.CommandBar.Controls.Count > 0 Then
                            'ShowPopup���ᴥ��InitCommandsPopup�¼�
                            objPopup.CommandBar.ShowPopup
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
    
    '��Ϊ����ͬ,��ȡ����ʱ�ᶪʧ��,Resize��ָ�
    tbcAppend.Height = tbcAppend.Height + 30
    tbcAppend.Height = tbcAppend.Height - 30
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Function RowInSameNo(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�����ͬһ�����ݺŷ�Χ�У�����Ƿ����кŷ�Χ
    Dim i As Long
 
    With vsAdvice
        lngBegin = lngRow
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6" Or .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 4 Then
                If .Cell(flexcpData, i, COL_������) = .Cell(flexcpData, lngRow, COL_������) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        lngEnd = lngRow
        For i = lngRow + 1 To .Rows - 1
            If .TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6" Or .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 4 Then
                If .Cell(flexcpData, i, COL_������) = .Cell(flexcpData, lngRow, COL_������) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
    If lngEnd <> lngRow Or lngBegin <> lngRow Then
        RowInSameNo = True
    End If
End Function

Private Sub ShowTotalMoney()
'���ܣ�ҽ���ܽ�����ʾ
'˵��������ҩƷʱ�ۣ��͸�ҩ;������ҩ�巨�÷��ȣ��¿�ҽ����һ��׼ȷ
    Dim rsMoney As New ADODB.Recordset, strSQL As String, str�����շ� As String
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim curҩƷӦ�� As Currency, curҩƷʵ�� As Currency
    Dim cur�¿� As Currency, curҩƷ�¿� As Currency
    Dim curԤ�� As Currency, curTmp As Currency
    Dim strSQLTmp As String
    Dim strTmp As String
    
    '������ҩƷ��ʱ�ۣ�ƴ�ӵ���ѯ�����
    strSQLTmp = "Zl_Calcdrugprice(a.ִ�п���id, s.ҩƷid, a.�ܸ�����," & gbytMediOutMode & "," & Len(gstrDecPrice) - 2 & "," & Len(gstrDecPrice) - 2 & ")"
 
    On Error GoTo errH
    
    strSQL = _
        " Select /*+ RULE */ Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
        " Sum(Decode(Instr('567',A.�շ����),0,0,A.Ӧ�ս��)) as ҩƷӦ��," & _
        " Sum(Decode(Instr('567',A.�շ����),0,0,A.ʵ�ս��)) as ҩƷʵ��" & _
        " From ������ü�¼ A,����ҽ������ B,����ҽ����¼ C" & _
        " Where A.ҽ�����=B.ҽ��ID And B.ҽ��ID=C.ID" & _
        " And C.����ID+0=[1] And C.�Һŵ�=[2]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
    ElseIf zlDatabase.DateMoved(mvRegDate) Then
        strTmp = strSQL
        strTmp = Replace(strTmp, "����ҽ����¼", "H����ҽ����¼")
        strTmp = Replace(strTmp, "����ҽ������", "H����ҽ������")
        strTmp = Replace(strTmp, "������ü�¼", "H������ü�¼")
        strSQL = strSQL & " Union ALL " & strTmp

        strSQL = "Select Sum(Ӧ�ս��) as Ӧ�ս��,Sum(ʵ�ս��) as ʵ�ս��," & _
            " Sum(ҩƷӦ��) as ҩƷӦ��,Sum(ҩƷʵ��) as ҩƷʵ�� From (" & strSQL & ")"
    End If
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�)
    If Not rsMoney.EOF Then
        curӦ�� = NVL(rsMoney!Ӧ�ս��, 0)
        curʵ�� = NVL(rsMoney!ʵ�ս��, 0)
        curҩƷӦ�� = NVL(rsMoney!ҩƷӦ��, 0)
        curҩƷʵ�� = NVL(rsMoney!ҩƷʵ��, 0)
    End If
    
    str�����շ� = "Select * From (" & _
        "Select Distinct C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,C.���ÿ���id" & _
        " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
        " From �����շѹ�ϵ C,����ҽ����¼ A Where A.����ID+0=[1] And A.�Һŵ�=[2] And A.������ĿID+0=C.������ĿID" & _
        "   And (a.���id Is Null And a.ִ�б�� In (1, 2) And c.�������� = 1 Or" & vbNewLine & _
        "   a.�걾��λ = c.��鲿λ And a.��鷽�� = c.��鷽�� And Nvl(c.��������, 0) = 0 Or" & vbNewLine & _
        "   (a.��鷽�� Is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(c.��������, 0) = 0 And c.��鲿λ Is Null And c.��鷽�� Is Null)" & _
        "      And (C.���ÿ���ID is Null or C.���ÿ���ID = A.ִ�п���ID And C.������Դ = 1)" & _
        " ) Where Nvl(���ÿ���id, 0) = Top"
        
    'ʱ��ҩƷȡ"ָ�����ۼ�"
    strSQL = _
        "Select Sum(Round(���," & gbytDec & ")) As ���,Sum(Round(ҩƷ���," & gbytDec & ")) As ҩƷ���" & _
        " From (Select A.�ܸ�����*Decode(I.�Ƿ���,1," & strSQLTmp & ",P.�ּ�) As ���," & _
        "              A.�ܸ�����*Decode(I.�Ƿ���,1," & strSQLTmp & ",P.�ּ�) As ҩƷ���" & _
        "       From ����ҽ����¼ A,�շ���ĿĿ¼ I,�շѼ�Ŀ P,ҩƷ��� S" & _
        "       Where A.�շ�ϸĿID=I.ID And I.ID=P.�շ�ϸĿID And I.ID=S.ҩƷID" & _
        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "I", "P", "3", "4", "5") & _
        "             And (Sysdate Between P.ִ������ And P.��ֹ���� Or Sysdate>=P.ִ������ And P.��ֹ���� is Null)" & _
        "             And A.ҽ��״̬=1 And A.������� In ('5','6')" & _
        "             And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        "       Union All" & _
        "       Select A.�ܸ�����*A.��������/S.����ϵ��*Decode(I.�Ƿ���,1," & strSQLTmp & ",P.�ּ�) As ���," & _
        "              A.�ܸ�����*A.��������/S.����ϵ��*Decode(I.�Ƿ���,1," & strSQLTmp & ",P.�ּ�) As ҩƷ���" & _
        "       From ����ҽ����¼ A,�շ���ĿĿ¼ I,�շѼ�Ŀ P,ҩƷ��� S" & _
        "       Where A.�շ�ϸĿID=I.ID And I.ID=P.�շ�ϸĿID And I.ID=S.ҩƷID" & _
        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "I", "P", "3", "4", "5") & _
        "             And (Sysdate Between P.ִ������ And P.��ֹ���� Or Sysdate>=P.ִ������ And P.��ֹ���� is Null)" & _
        "             And A.ҽ��״̬=1 And A.�������='7'" & _
        "             And A.����ID+0=[1] And A.�Һŵ�=[2]"
        strSQL = strSQL & "  Union All" & _
        "       Select Nvl(A.�ܸ�����,A.Ƶ�ʴ���)*R.�շ�����*Decode(I.�Ƿ���,1,P.ȱʡ�۸�,P.�ּ�) As ���,0 as ҩƷ���" & _
        "       From ����ҽ����¼ A,(" & str�����շ� & ") R,�շ���ĿĿ¼ I,�շѼ�Ŀ P" & _
        "       Where A.������ĿID+0=R.������ĿID And I.ID=R.�շ���ĿID And I.ID=P.�շ�ϸĿID" & _
        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "I", "P", "3", "4", "5") & _
        "             And (I.վ��='" & gstrNodeNo & "' Or I.վ�� is Null)" & _
        "             And (Sysdate Between P.ִ������ And P.��ֹ���� Or Sysdate>=P.ִ������ And P.��ֹ���� is Null)" & _
        "             And Nvl(A.�Ƽ�����,0)=0 And A.ҽ��״̬=1 And A.������� Not In ('5','6','7')" & _
        "             And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        "             And (a.���id Is Null And a.ִ�б�� In (1, 2) And r.�������� = 1 Or" & _
        "                 a.�걾��λ = r.��鲿λ And a.��鷽�� = r.��鷽�� And Nvl(r.��������, 0) = 0 Or" & _
        "                 a.��鷽�� Is Null And Nvl(r.��������, 0) = 0 And r.��鲿λ Is Null And r.��鷽�� Is Null)) A"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
    If Not rsMoney.EOF Then
        cur�¿� = NVL(rsMoney!���, 0)
        curҩƷ�¿� = NVL(rsMoney!ҩƷ���, 0)
    End If
    
    strSQL = "Select Nvl(Ԥ�����,0)-Nvl(�������,0) as ��� From ������� Where ����=1 And ���� = 1 And ����ID=[1]"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID)
    If Not rsMoney.EOF Then curԤ�� = NVL(rsMoney!���, 0)
    
    strSQL = _
        "ҽ���ѷ���Ӧ��:" & FormatEx(curӦ��, gbytDec) & "(ҩ" & FormatEx(curҩƷӦ��, gbytDec) & ")," & _
        "ʵ��:" & FormatEx(curʵ��, gbytDec) & "(ҩ" & FormatEx(curҩƷʵ��, gbytDec) & ")" & _
        "  �¿�Լ:" & FormatEx(cur�¿�, gbytDec) & "(ҩ" & FormatEx(curҩƷ�¿�, gbytDec) & ")" & _
        IIF(curԤ�� = 0, "", "  Ԥ��:" & FormatEx(curԤ��, 2))
    If curԤ�� <> 0 And cur�¿� > curԤ�� Then
        curTmp = cur�¿� - curԤ��
        strSQL = strSQL & "  �貹��:" & FormatEx(curTmp, gbytDec)
    End If
    RaiseEvent StatusTextUpdate(strSQL)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            vsAdvice.ColWidth(lngCol) = vsAdvice.ColData(lngCol)
            vsAdvice.ColHidden(lngCol) = False
        Else
            vsAdvice.ColWidth(lngCol) = 0
            vsAdvice.ColHidden(lngCol) = True
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub FuncBloodApply(ByVal intType As Long)
'���ܣ���Ѫ���뵥
'������intType=0 ������=1�޸ģ�=2�鿴��=4�˶�
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long
    Dim bln��Ѫ As Boolean
    Dim blnApply As Boolean
    
    If intType <> 2 Then
        '���ҺŲ����Ƿ���
        If Not FuncTimeLimitCheck Then Exit Sub
        '����Ƿ������м�����רҵ����ְ��
        If gbln��Ѫ�����м����� Then
            If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
                MsgBox "��������Ѫ�ּ��������Ѫҽ��ֻ���м�������רҵ����ְ��ҽʦ�����´", vbInformation, "��Ѫ���뵥"
                Exit Sub
            End If
        End If
        '�޸�ʱ����Ƿ����
        If intType = 1 Then
            If Not CanEditBloodAdvice(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��־)) = 1, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��鷽��)) = 1) Then Exit Sub
        End If
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         bln��Ѫ = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��鷽��)) = 1
         lngRow = vsAdvice.Row
    End If
    
    If gblnѪ��ϵͳ = True Then
        blnApply = frmApplyBloodNew.ShowMe(Me, mlng����ID, 0, 1, intType, lngUpdateAdvice, mlng�Һſ���ID, , mlng�Һſ���ID, , , mrsDefine, mclsMipModule, 1, mstr�Һŵ�, , , , , mlngǰ��ID, IIF(bln��Ѫ = True, 1, 0))
    Else
        blnApply = frmApplyBlood.ShowMe(Me, mlng����ID, 0, 1, intType, lngUpdateAdvice, mlng�Һſ���ID, , mlng�Һſ���ID, , , mrsDefine, mclsMipModule, 1, mstr�Һŵ�, , , , , mlngǰ��ID)
    End If
    
    If blnApply = True Then
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
    End If
End Sub

Private Sub FuncOperationApply(ByVal intType As Long)
'���ܣ��������뵥
'������intType=0 ������=1�޸ģ�=2�鿴
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long, strDefine As String
    
    If intType <> 2 Then
        '���ҺŲ����Ƿ���
        If Not FuncTimeLimitCheck Then Exit Sub
        '�޸�ʱ����Ƿ����
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 2 Then
                MsgBox "���뵥�Ѿ���ˣ����������޸ġ�", vbInformation, "�������뵥"
                intType = 2
            End If
        End If
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
    End If
     
    If Not mrsDefine Is Nothing Then
        mrsDefine.Filter = "�������='F'"
        If Not mrsDefine.EOF Then strDefine = Trim(NVL(mrsDefine!ҽ������))
    End If

    If frmApplyOperation.ShowMe(Me, 1, intType, mlng����ID, mstr�Һŵ�, 1, lngUpdateAdvice, mlng�Һſ���ID, mlng�Һſ���ID, strDefine, , , , 0, mclsMipModule, , , mlngǰ��ID) Then
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
    End If
End Sub

Private Sub FuncLISApply(ByVal lng������� As Long)
'���ܣ����ü�������������뵥�ͼ���ҽ��
'������lng�������=�޸����뵥ʱ���������
    Dim arrTmp As Variant, arrSQL As Variant, i As Long, blnTrans As Boolean, strSQL As String
    Dim strResult As String, strDiag As String, strDept As String, strErr As String
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset 'ע��˱�����Ҫ����,��LisInfoTrans��������ֵ
    Dim rsTemp As ADODB.Recordset
    Dim lngҽ��ID As Long, lng���ID As Long, lng��� As Long
    Dim lngִ�п���ID As Long, lng�ɼ�����ID As Long, lng������ĿID As Long, lng�ɼ���ĿID As Long
    Dim str����Ƽ����� As String, str�ɼ��Ƽ����� As String, str����ִ������ As String, str�ɼ�ִ������ As String
    Dim str������Ŀ As String, str�ɼ����� As String, str�걾 As String, str���� As String
    Dim strCurDate As String, str��ʼִ��ʱ�� As String, strҽ������ As String, strҽ��IDs As String, blnCancel As Boolean
    Dim strDelIDs As String, arrDelID() As String
    Dim Y As Long, j As Long, str������Ŀ��� As String
    Dim str���� As String, str���� As String
    Dim arrAppend As Variant
    Dim lng������� As Long
    Dim str��� As String
    Dim lng��ҽ��ID As Long '����ҽ��ID����ֵ���˷ѣ�������ύ����ʱ�������ҽ��ID
    Dim strҽ��ID As String, str���ID As String
    Dim varID As Variant
    Dim strTmp As String
    Dim bln���Ѷ��� As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim strItems As String
    Dim strTabAdvice As String
    Dim blnCheckItem As Boolean 'ҽ���ܿؼ��
    Dim rsPrice As ADODB.Recordset
    Dim strժҪ As String, strMsg As String
    Dim rsLISInfo As ADODB.Recordset
    Dim lng������� As Long
    Dim dat��ʼִ��ʱ�� As Date
    Dim dat��ǰʱ�� As Date
    
    '���ҺŲ����Ƿ���
    If Not FuncTimeLimitCheck Then Exit Sub
    
    Set rsPati = GetPatiInfo()
    If rsPati.RecordCount = 0 Then
        MsgBox "δ����ȷ��ȡ������Ϣ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If lng������� <> 0 Then
        strDiag = GetAdviceDiag(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    End If
    
    strDept = Sys.RowValue("���ű�", mlng�Һſ���ID, "����")
    Call InitObjLis(p����ҽ��վ)
    If gobjLIS Is Nothing Then Exit Sub
    Call CreatePlugInOK(p����ҽ���´�, mint����)
    
    On Error GoTo errH
 
    '������ѡ��ļ�����Ŀ��ʽ����: �������ID1,ִ�п���ID1,����ʱ��1,������Ŀ����1,�걾1,����ҽ��1,�ɼ���ʽ������ĿID 1;�������ID2,ִ�п���ID2,����ʱ��2,������Ŀ����2,�걾2,����ҽ��2,�ɼ���ʽ������ĿID 2;.....
    strResult = gobjLIS.ShowLisApplicationForm(mfrmParent, lng�������, mlng����ID, 0, Val("" & rsPati!�Һ�ID), rsPati!����, "" & rsPati!�Ա�, "" & rsPati!����, 1, _
        Val("" & rsPati!�����), Val("" & rsPati!סԺ��), Val("" & rsPati!������), strDiag, UserInfo.����, UserInfo.����ID, UserInfo.������, mlng�Һſ���ID, strDept, blnCancel, strErr)
 
    If strErr <> "" Then
        MsgBox "����ӿ��ڲ�����" & strErr, vbInformation, gstrSysName
    ElseIf blnCancel Then
        Exit Sub    'ȡ�����˳�
    Else
        arrSQL = Array()
        '�޸����뵥ʱ����ɾ���ɵ�ҽ��
        If lng������� <> 0 Then
            strҽ��IDs = GetAdivceBy�������(lng�������)
            For i = 0 To UBound(Split(strҽ��IDs, ","))
                '����ɾ��ǰ��ҽӿ�
                On Error Resume Next
                If Not gobjPlugIn Is Nothing Then
                    If gobjPlugIn.AdviceDeletBefor(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(Split(strҽ��IDs, ",")(i)), mint����) = False Then
                        If err.Number = 0 Then Exit Sub
                    End If
                    Call zlPlugInErrH(err, "AdviceDeletBefor")
                End If
                If err.Number <> 0 Then err.Clear
                On Error GoTo errH
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & Split(strҽ��IDs, ",")(i) & ",1)"
                strDelIDs = strDelIDs & "," & Split(strҽ��IDs, ",")(i)
            Next
            strDelIDs = Mid(strDelIDs, 2)
        End If
        
        If strResult <> "" Then
            If strDiag = "" Then
                'ȡһ�����������Ĭ�Ϲ���
                strSQL = "Select a.ID,a.������� From ������ϼ�¼ A Where a.����id=[1] and a.��ҳid =[2] and a.��¼��Դ = 3 order by a.�������,a.��ϴ���"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
                If Not rsTemp.EOF Then
                    strDiag = rsTemp!ID & ""
                    str��� = "���뵥���<Split2>0<Split2><Split2>" & rsTemp!�������
                End If
            Else
                str��� = GetDiag�������(strDiag)
                If str��� <> "" Then
                    str��� = "���뵥���<Split2>0<Split2><Split2>" & str���
                End If
            End If
              
            bln���Ѷ��� = True
            
            If mint���� <> 0 Then
                If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                    blnCheckItem = True
                End If
            End If
            dat��ǰʱ�� = zlDatabase.Currentdate()
            strCurDate = "To_Date('" & Format(dat��ǰʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            lng��� = GetMaxAdviceNO(mlng����ID, 0, 0)
            lng������� = -1
            '�ڸ÷����ж�rsLISInfo, rsTmp��ֵ
            Call LisInfoTrans(strResult, rsLISInfo, rsTmp)
                        
            'ֻ��������
            For i = 1 To rsLISInfo.RecordCount
                If lng������� <> Val(rsLISInfo!��� & "") Then
                    lng������� = Val(rsLISInfo!��� & "")
                    lng������� = Get�������
                End If
                lng��ҽ��ID = lng��ҽ��ID + 1
                str���ID = "<FAKEID>" & lng��ҽ��ID & "</FAKEID>"
                lng���ID = lng��ҽ��ID
                lng�ɼ�����ID = Val(rsLISInfo!�ɼ�����ID & "")
                lngִ�п���ID = Val(rsLISInfo!ִ�п���ID & "")
                str��ʼִ��ʱ�� = rsLISInfo!��ʼִ��ʱ�� & ""
                str�걾 = rsLISInfo!�걾 & ""
                str���� = rsLISInfo!���� & ""
                str���� = rsLISInfo!���� & ""
                str���� = rsLISInfo!���� & ""
                lng�ɼ���ĿID = Val(rsLISInfo!�ɼ���ĿID & "")
                lng������ĿID = Val(rsLISInfo!������ĿID & "")
                
                dat��ʼִ��ʱ�� = CDate(str��ʼִ��ʱ��)
                str��ʼִ��ʱ�� = "To_Date('" & Format(dat��ʼִ��ʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
                
                'a.�Ȳ�������ҽ�� ���뵥�������ĵļ���ҽ��ֻ��һ��������ĿID
                rsTmp.Filter = "ID=" & lng������ĿID
                str������Ŀ = rsTmp!���� & ""
                str����Ƽ����� = Val("" & rsTmp!�Ƽ�����)
                str����ִ������ = IIF("" & rsTmp!ִ�п��� = "", "NULL", "" & rsTmp!ִ�п���)
                strҽ������ = str������Ŀ & IIF("" = rsLISInfo!ʱ������ & "", "", "(" & rsLISInfo!ʱ������ & ")")
                lng��� = lng��� + 1
                strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(lng������ĿID) & "||1")
                blnCancel = CheckLISAppAdvice(1, mlng����ID, mlng�Һ�ID, mint����, "C", lng������ĿID, mlng�Һſ���ID, UserInfo.����, lngִ�п���ID, Val(rsTmp!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
                
                If Not blnCancel Then Exit Sub
                
                lng��ҽ��ID = lng��ҽ��ID + 1
                strҽ��ID = "<FAKEID>" & lng��ҽ��ID & "</FAKEID>"
                lngҽ��ID = lng��ҽ��ID
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & _
                    strҽ��ID & "," & str���ID & "," & lng��� & ",1," & mlng����ID & "," & _
                    "Null,0,1,1,'C'," & _
                    lng������ĿID & ",Null,Null,Null,1," & _
                    "'" & strҽ������ & "',Null," & "'" & str�걾 & "','һ����',Null," & _
                    "Null,Null,Null," & str����Ƽ����� & "," & lngִ�п���ID & _
                    "," & str����ִ������ & "," & str���� & "," & str��ʼִ��ʱ�� & ",Null," & mlng�Һſ���ID & "," & _
                    mlng�Һſ���ID & ",'" & UserInfo.���� & "'," & strCurDate & ",'" & mstr�Һŵ� & "'," & ZVal(mlngǰ��ID) & "," & _
                    "NULL,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                    ",Null,Null,Null,Null," & lng������� & ",null,null,null,null,null,'" & rsLISInfo!ʱ��ID & "')"
                    
                strItems = strItems & "," & lng������ĿID & ":" & lngִ�п���ID
                
                If blnCheckItem Then
                    strTabAdvice = _
                        "select " & lngҽ��ID & " as ID," & lng��� & " as ���," & lng���ID & " as ���ID,'C' as �������," & lng������ĿID & " as ������ĿID," & _
                        lng������ĿID & " as ������ĿID,-null as �շ�ϸĿID, 1 As ����, 0 As ����,'" & str�걾 & "' as �걾��λ,'' As ��鷽��," & _
                        "0 as ִ�б��," & Val("" & rsTmp!�Ƽ�����) & " as �Ƽ�����, 0 As ��������," & Val("" & rsTmp!ִ�п���) & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"
                End If
            

                'b.�ٲ����ɼ�����ҽ��
                rsTmp.Filter = "ID=" & lng�ɼ���ĿID
                str�ɼ����� = rsTmp!���� & ""
                str�ɼ��Ƽ����� = Val("" & rsTmp!�Ƽ�����)
                str�ɼ�ִ������ = "" & rsTmp!ִ�п���
                strҽ������ = AdviceMakeText(str������Ŀ, str�ɼ�����, str�걾)
                If "" <> rsLISInfo!ʱ������ & "" Then strҽ������ = strҽ������ & "(" & rsLISInfo!ʱ������ & ")"
                lng��� = lng��� + 1
                strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(lng�ɼ���ĿID) & "||1")
                blnCancel = CheckLISAppAdvice(1, mlng����ID, mlng�Һ�ID, mint����, "E", lng�ɼ���ĿID, mlng�Һſ���ID, UserInfo.����, lng�ɼ�����ID, Val(rsTmp!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
                If Not blnCancel Then Exit Sub
                    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & _
                    str���ID & ",Null," & lng��� & ",1," & mlng����ID & "," & _
                    "Null,0,1,1,'E'," & _
                    lng�ɼ���ĿID & ",Null,Null,Null,1," & _
                    "'" & strҽ������ & "','" & str���� & "'," & "'" & str�걾 & "','һ����',Null," & _
                    "Null,Null,Null," & str�ɼ��Ƽ����� & "," & lng�ɼ�����ID & _
                    "," & str�ɼ�ִ������ & "," & str���� & "," & str��ʼִ��ʱ�� & ",Null," & mlng�Һſ���ID & "," & _
                    mlng�Һſ���ID & ",'" & UserInfo.���� & "'," & strCurDate & ",'" & mstr�Һŵ� & "'," & ZVal(mlngǰ��ID) & "," & _
                    "NULL,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                    ",Null,Null,Null,Null," & lng������� & ",null,null,null,null,null,'" & rsLISInfo!ʱ��ID & "')"
                    
                strItems = strItems & "," & lng�ɼ���ĿID & ":" & lng�ɼ�����ID
                
                If blnCheckItem Then
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & lng���ID & " as ID," & lng��� & " as ���,-null as ���ID,'E' as �������," & lng������ĿID & " as ������ĿID," & _
                        lng�ɼ���ĿID & " as ������ĿID,-null as �շ�ϸĿID, 1 As ����, 0 As ����,'" & str�걾 & "' as �걾��λ,'' As ��鷽��," & _
                        "0 as ִ�б��," & Val("" & rsTmp!�Ƽ�����) & " as �Ƽ�����, 0 As ��������," & Val("" & rsTmp!ִ�п���) & " As ִ������," & lng�ɼ�����ID & " as ִ�п���id from dual"
                End If
                
                'ҽ��������
                If gintҽ������ = 2 Then bln���Ѷ��� = True
                strMsg = CheckAdviceInsure(mint����, bln���Ѷ���, mlng����ID, 1, "", Mid(strItems, 2), Left(strҽ������, 50))
                If strMsg <> "" Then
                    If gintҽ������ = 1 Then
                        vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", Me)
                        If vMsg = vbNo Or vMsg = vbCancel Then Exit Sub
                        If vMsg = vbIgnore Then bln���Ѷ��� = False
                    ElseIf gintҽ������ = 2 Then
                        MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strMsg = ""
                End If
                
                'ҽ���ܿ�ʵʱ��⣺�״�����(����)���߸���ʱ���
                If blnCheckItem Then
                    If MakePriceRecord���뵥("11", mlng����ID, mlng�Һ�ID, strTabAdvice, strItems, rsPati!�ѱ� & "", mlng�Һſ���ID, rsPrice) Then
                        If Not gclsInsure.CheckItem(mint����, 0, 0, rsPrice) Then
                            MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´��LIS���뵥���ܱ��档", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                strItems = ""
                
                If str���� <> "" And str��� <> "" Then
                    str���� = str��� & "<Split1>" & str����
                ElseIf str���� = "" And str��� <> "" Then
                    str���� = str���
                End If
                
                '�������븽�������������Ȳ���ҽ��
                If str���� <> "" Then
                    arrAppend = Split(str����, "<Split1>")
                    For j = 0 To UBound(arrAppend)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & str���ID & "," & _
                            "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                            j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                            IIF(j = 0, ",1", "") & ")"
                        lng������� = j + 1
                    Next
                End If
                
                If strDiag <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(" & str���ID & ",'" & strDiag & "')"
                End If
                rsLISInfo.MoveNext
            Next
        End If
        
        '�����в�����ʵ��ҽ��ID
        If lng��ҽ��ID > 0 Then
            For j = 1 To lng��ҽ��ID
                Y = zlDatabase.GetNextID("����ҽ����¼")
                If j = 1 Then
                    strҽ��IDs = ""
                    strҽ��IDs = Y
                Else
                    strҽ��IDs = strҽ��IDs & "," & Y
                End If
            Next
            varID = Split(strҽ��IDs, ",")
            
            For i = 0 To UBound(arrSQL)
                strTmp = arrSQL(i)
                If InStr(strTmp, "<FAKEID>") > 0 Then
                    j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                    strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    
                    If InStr(strTmp, "<FAKEID>") > 0 Then '����滻����
                        j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                        strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    End If
                    arrSQL(i) = strTmp
                End If
            Next
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        'ˢ��ҽ��
        Call RefreshData
        
    End If
    '����ɾ������ҽӿ�
    On Error Resume Next
    arrDelID = Split(strDelIDs, ",")
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            If Not gobjPlugIn Is Nothing Then
                Call gobjPlugIn.AdviceDeleted(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, Val(arrDelID(i)), mint����)
            End If
            Call zlPlugInErrH(err, "AdviceDeleted")
        End If
    Next
    If err.Number <> 0 Then err.Clear
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAdivceBy�������(ByVal lng������� As Long) As String
'���ܣ�����������Ż�ȡ���м��ҽ��ID�����ɼ�ҽ��ID��
    Dim i As Long, strTmp As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_�������)) = lng������� Then
                If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                    strTmp = strTmp & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
        GetAdivceBy������� = Mid(strTmp, 2)
    End With
End Function

Private Function AdviceMakeText(ByVal str���� As String, ByVal str�ɼ� As String, ByVal str�걾 As String) As String
'���ܣ���������ҽ����ҽ������
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
               
    'ȷ���Ƿ���
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "�������='C'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(NVL(mrsDefine!ҽ������)) = "" Then
            blnDefine = False
        End If
    End If
    
    If Not blnDefine Then
        strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
    Else
        strText = mrsDefine!ҽ������
        If InStr(strText, "[������Ŀ]") > 0 Then
            strField = str����
            strText = Replace(strText, "[������Ŀ]", """" & strField & """")
        End If
        If InStr(strText, "[����걾]") > 0 Then
            strField = str�걾
            strText = Replace(strText, "[����걾]", """" & strField & """")
        End If
        If InStr(strText, "[�ɼ�����]") > 0 Then
            strField = str�ɼ�
            strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
        End If
        
        '����ҽ������
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
        End If
        err.Clear: On Error GoTo 0
    End If
        
    AdviceMakeText = strText
End Function


Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    mlngFontSize = IIF(bytSize = 0, 9, 12)
    '����vsFlexGrid�ؼ���ʹ�ø��Ի�����ʱ��Ӵ��п�����ڴ�����μ����ǲ���������,��Ҫ��getForm��������
    If Not Me.Visible Then
        vsAdvice.FontSize = mlngFontSize
        vsAppend.FontSize = mlngFontSize
        vsfAdivceDetail.FontSize = mlngFontSize
    End If
    If mvarCond.��ʾģʽ = 0 Then
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_����)
    Else
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_ҽ������)
    End If
    
    Call Grid.SetFontSize(vsAppend, mlngFontSize)
    
    Call Grid.SetFontSize(vsfAdivceDetail, mlngFontSize)
    
    'ѪҺִ�к�ѪҺ��ϸ����
    If Not mobjFrmBloodList Is Nothing Then
        If mobjFrmBloodList.Visible = True Then Call mobjFrmBloodList.SetFontSize(mlngFontSize)
    End If
    
    Call SetRTFFont(0)
End Sub

Private Sub FuncPacsApply(ByVal lngҽ��ID As Long, ByRef lng������� As Long)
'���ܣ����ü�����뵥
'������lngҽ��ID=�޸����뵥ʱ��ǰ�е�ҽ��ID,lng������� =��ǰ�޸��е��������
    Dim lngNo As Long
    
    '���ҺŲ����Ƿ���
    If Not FuncTimeLimitCheck Then Exit Sub
    
    lngNo = ApplyOutPacs(Me, lng�������, mlng����ID, mstr�Һŵ�, lngҽ��ID, mlng�Һſ���ID, mobjVBA, mobjScript, mrsDefine, mblnMoved, , mlngǰ��ID)
    
    If lngNo <> 0 Then Call LoadAdvice

End Sub

Private Sub FuncApplyModi()
'���ܣ��޸����뵥
    Dim strSQL As String, rsTmp As Recordset
    With vsAdvice
        '���ж��Ƿ����Զ������뵥
        strSQL = "Select �ļ�ID From ҽ�����뵥�ļ� Where ҽ��ID=[1] And RowNum<2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_���ID)) = 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_���ID))))
        If rsTmp.RecordCount > 0 Then
            FuncApplyCustom 1, Val(rsTmp!�ļ�ID)
        Else
                        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 1 Then
                MsgBox "�������޸��ѷ��͵����롣", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E" Then
                Call FuncLISApply(Val(.TextMatrix(.Row, COL_�������)))
            ElseIf .TextMatrix(.Row, COL_�������) = "D" Then
                Call FuncPacsApply(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_�������)))
            ElseIf .TextMatrix(.Row, COL_�������) = "K" Then
                If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 1 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��鷽��)) = 1 Then
                    Call FuncBloodApply(4)
                Else
                    Call FuncBloodApply(1)
                End If
            ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                Call FuncOperationApply(1)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncApplyView()
'���ܣ��鿴���뵥
    Dim lngҽ��ID As Long
    Dim lngNo As Long
    Dim strSQL As String, rsTmp As Recordset
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_�������))
        
        If lngҽ��ID <> 0 And lngNo <> 0 Then
            strSQL = "Select �ļ�ID From ҽ�����뵥�ļ� Where ҽ��ID=[1] And RowNum<2"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_���ID)) = 0, lngҽ��ID, Val(.TextMatrix(.Row, COL_���ID))))
            If rsTmp.RecordCount > 0 Then
                FuncApplyCustom 2, Val(rsTmp!�ļ�ID)
            Else
                If .TextMatrix(.Row, COL_�������) = "K" Then
                    Call FuncBloodApply(2)
                ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                    Call FuncOperationApply(2)
                ElseIf .TextMatrix(.Row, COL_�������) = "D" Then
                    '���
                    If Val(Mid(gstrOutUseApp, 1, 1)) = 1 Then
                        Call ShowApply���(Me, lngNo)
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlPASSMap()
'����:����Pass VsAdvie����ӳ��
'ע��:ɾ�����޸�������������ʱ�����������ҩ�����еĹ�������
    Dim blnTmp As Boolean
    
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "������ҩ���", True)
    End If
    
    If gobjPass Is Nothing Then  '83970
        blnTmp = False
    Else
        blnTmp = gobjPass.PassType <> UNPASS
    End If
    
    mblnPass = Not mobjPassMap Is Nothing And blnTmp
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_����ҽ���嵥
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .VSCOL = .GetVSCOL(COL_ID, COL_���ID, COL_�������, COL_������ĿID, COL_�շ�ϸĿID, col_ҽ������, , COL_����, , COL_�÷�, COL_����, , COL_����ʱ��, COL_����ҽ��, _
                        COL_��ʼʱ��, COL_��������ID, , COL_Ƶ��, , , , COL_��ʾ, , COL_ҽ��״̬, , , , , COL_ִ������, COL_�걾��λ, , , , , , COL_����, , COL_ҽ������, COL_��ҩĿ��, COL_��������)
            Set .PassPati = .GetPatient(mlng����ID, mlng�Һ�ID)
            mblnPass = gobjPass.zlPassCheck(mobjPassMap)
        End With
    End If
End Sub

Private Sub zlPASSPati()
'����:���ò�����Ϣ
    If Not mobjPassMap Is Nothing Then
        With mobjPassMap.PassPati
            .intӤ�� = -1 'ȱʡ��������Ϊ0
            .dbl��ʶ�� = -1
            .Dat�������� = -1
            .lng����ID = mlng����ID
            .lng��ҳID = -1
            .str�Һŵ� = mstr�Һŵ�
            .str���� = ""
            .str�Ա� = ""
            .str���� = ""
        End With
    End If
End Sub

Private Sub SetAdviceColVisible()
'���ܣ�����ҽ������еĿɼ��Ժͱ�ͷ����
    Dim i As Long
    
    '������ʾģʽ������ʾ��
    With vsAdvice
        .ColHidden(col_ҽ������) = mvarCond.��ʾģʽ = 0
        .ColHidden(col_����) = mvarCond.��ʾģʽ = 1
        .ColHidden(COL_Ƥ��) = False
        .ColHidden(COL_����) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_����) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_����) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_Ƶ��) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_ִ��ʱ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ִ��ʱ��) = "Detail"
        .ColHidden(COL_ִ������) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ִ������) = "Detail"
        .ColHidden(COL_����ʱ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_����ʱ��) = "Detail"
        .ColHidden(COL_����ҩ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_����ҩ��) = "Detail"
        .ColHidden(COL_��ΣҩƷ) = True
        .ColHidden(COL_�걾��λ) = True
        .ColHidden(COL_�շ�ϸĿID) = True
        .ColHidden(COL_��鱨��ID) = True
        .ColHidden(COL_�������״̬) = True
        .ColHidden(COL_���������) = True
        .ColHidden(COL_������) = True
        .ColHidden(COL_������ӡ) = True
        .ColHidden(COL_����Ԥ��) = True
        .ColHidden(COL_��) = True
        .ColHidden(COL_�걾״̬) = True
        If mvarCond.ҽ�� = 1 Then
            .ColHidden(COL_������) = False
            .ColHidden(COL_������ӡ) = False
            .ColHidden(COL_����Ԥ��) = False
        End If
        If mvarCond.����ģʽ = 0 And (mvarCond.ҽ�� = 0 Or mvarCond.ҽ�� = 1) Then
            .ColHidden(COL_��) = False
        End If
        If mvarCond.����ģʽ = 3 Then '���Ǳ��濨Ƭ�Ȳ�����ʾ
            For i = COL_��ʼʱ�� + 1 To COL_�걾��λ
                .ColHidden(i) = True
            Next
            .ColHidden(COL_��ʼʱ��) = False
            .ColHidden(col_����) = False
            .ColHidden(COL_ִ�п���) = False
            .ColHidden(COL_����ҽ��) = False
            .TextMatrix(0, COL_����ҽ��) = "����ҽ��"
            .ColHidden(COL_����״̬) = mfrmParent Is Nothing     '���Ӳ�������δ����������,��ֹ��ʾ����״̬
            .ColWidth(COL_����״̬) = 700
            .TextMatrix(0, COL_����״̬) = "����"
            .ColHidden(COL_�걾״̬) = False
            .ColWidth(COL_�걾״̬) = 850
        Else
            .TextMatrix(0, COL_����ҽ��) = "����ҽ��"
            .ColHidden(COL_�÷�) = False
            .ColHidden(COL_ҽ������) = False
            .ColHidden(COL_����״̬) = True
            .TextMatrix(0, COL_����״̬) = "����״̬"
        End If
        If mvarCond.��ʾģʽ = 1 Then .ColHidden(COL_����) = Not mbln����
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    RaiseEvent VSKeyPress(KeyAscii)
End Sub

Private Sub SetTagһ����ҩ(Optional ByVal lngRow As Long)
'���ܣ���һ����ҩ��ҽ��ǰ�ӱ�־
    Dim i As Long
    Dim lngBg As Long, lngEd As Long
    Dim j As Long
    Dim lngStart As Long, lngEnd As Long

    If mvarCond.����ģʽ = 3 Then Exit Sub

    With vsAdvice
        If lngRow = 0 Then
            lngStart = .FixedRows
            lngEnd = .Rows - 1
        Else
            lngStart = lngRow
            lngEnd = lngRow
        End If
        For i = lngStart To lngEnd
             lngBg = -1: lngEd = -1
             If RowInһ����ҩ(i, lngBg, lngEd) Then
                For j = lngBg To lngEd
                    If j = lngBg Then
                        .TextMatrix(j, COL_��) = "��"
                    ElseIf j = lngEd Then
                        .TextMatrix(j, COL_��) = "��"
                    Else
                        .TextMatrix(j, COL_��) = "��"
                    End If
                Next
                If lngEd <> -1 Then
                   i = lngEd + 1
                End If
            Else
                .TextMatrix(i, COL_��) = ""
            End If
        Next
    End With
End Sub

Private Function FuncTimeLimitCheck() As Boolean
'�ҺŲ��˳������޼�飬true��δ���ڣ�false������/mlng����ID = 0/mblnEditable=false
    If mlng����ID = 0 Then FuncTimeLimitCheck = False: Exit Function
    If Not mblnEditable Then FuncTimeLimitCheck = False: Exit Function
    
    If mint���� = 0 Then
        '����ѡ��:0-����Ϊ�շѵ�,1-����Ϊ���ʵ�,2-�ֹ�ѡ��
        If Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�)) = 0 Then
            If BillExpend(mstr�Һŵ�) Then
                MsgBox "�ò��˹Һ��ѳ�����Ч�������������´�ҽ����", vbInformation, gstrSysName
                FuncTimeLimitCheck = False
                Exit Function
            End If
        End If
    End If
    FuncTimeLimitCheck = True
End Function

Private Function ShowAdviceRISSch(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'���ܣ���ʾָ���е�ԤԼ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    Dim i As Long
    
    blnExist = False
    rtfSche.Text = "": rtfSche.SelStart = 0
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_RISԤԼID)) = 0 Then Exit Function
    
    strSQL = "select ����豸����,To_Char(ԤԼ����,'YYYY-MM-DD') as ԤԼ����," & vbNewLine & _
        "To_Char(ԤԼ��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') as ԤԼ��ʼʱ��," & vbNewLine & _
        "To_Char(ԤԼ����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ԤԼ����ʱ��," & vbNewLine & _
        "To_Char(ԤԼ��ʼʱ���,'YYYY-MM-DD HH24:MI:SS') as ԤԼ��ʼʱ���," & vbNewLine & _
        "To_Char(ԤԼ����ʱ���,'YYYY-MM-DD HH24:MI:SS') as ԤԼ����ʱ���,DECODE(�Ƿ����,1,'�Ѿ�ԤԼ����','�Ѿ�ԤԼ') as ԤԼ״̬" & vbNewLine & _
        "from RIS���ԤԼ Where ҽ��ID=[1]"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfSche
            For i = 0 To rsTmp.Fields.Count - 1
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp.Fields(i).Name & "��" & NVL(rsTmp.Fields(i).value)
                lngIdx = .Find(rsTmp.Fields(i).Name & "��", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp.Fields(i).Name & "��")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
            Next
            '��궨λ�ڵ�һ��
            lngIdx = .Find(rsTmp.Fields(0).Name & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp.Fields(0).Name & "��")
            Call SetRTFFont(4)
        End With
        blnExist = True
    End If
    ShowAdviceRISSch = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceRISSch()
'���ܣ�RISҽ��ԤԼ
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    lngResult = -1
    If HaveRIS Then
        With vsAdvice
            If InStr(",1,8,", "," & .TextMatrix(.Row, COL_ҽ��״̬) & ",") >= 0 Then
                lngResult = gobjRis.HISScheduling(1, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_������ĿID)))
                If lngResult = 0 Then
                    '�ɹ�ԤԼ�����״̬
                    strSQL = "select min(ԤԼID) as ID from RIS���ԤԼ where ҽ��id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                    .TextMatrix(.Row, COL_RISԤԼID) = rsTmp!ID & ""
                    Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                End If
            Else
                MsgBox "ҽ��״̬Ϊ�¿����ѷ���ʱ������ԤԼ��", vbInformation, gstrSysName
            End If
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRISDel()
'���ܣ�RISҽ��ȡ��ԤԼ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngResult As Long
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_RISԤԼID)) <> 0 Then
            If Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 Then
                strSQL = "Select Max(b.ִ��״̬) As ��� From ����ҽ����¼ A, ����ҽ������ B Where a.Id = b.ҽ��id And (a.Id =[1] Or a.���id=[1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                If Not rsTmp.EOF Then
                    If Val(rsTmp!��� & "") = 0 Then
                        blnDo = True
                    Else
                        MsgBox "��ҽ���Ѿ���ִ�л��߲���ִ�в���ȡ��ԤԼ��", vbInformation, gstrSysName
                    End If
                End If
            Else
                blnDo = True
            End If
        End If
        If blnDo Then
            If HaveRIS Then
                lngResult = gobjRis.HISSchedulingEx(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_RISԤԼID)))
                If lngResult = 0 Then
                    '�ɹ���ȡ������״̬
                    .TextMatrix(.Row, COL_RISԤԼID) = ""
                    Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetAdviceReportIcon(ByVal lngRow As Long)
'���ܣ����ݵ�ǰ�е���������ҽ�������е�ͼ���ʶ
'˵����ע���ǵ������ã�����һ������

    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_����ID)) <> 0 Or _
            .TextMatrix(lngRow, COL_��鱨��ID) <> "" Or _
            Val(.TextMatrix(lngRow, COL_RIS����ID)) <> 0 Or _
            Val(.TextMatrix(lngRow, COL_LIS����ID)) <> 0 Then
            
            
            If Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 0 Then
                Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("����").Picture
            ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 1 Then
                Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("��������").Picture
            ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 2 Then
                Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("���沿����").Picture
            End If
        Else
            If Val(.TextMatrix(lngRow, COL_RISԤԼID)) <> 0 Then
                Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("ԤԼ").Picture
            End If
        End If
    End With
End Sub

Private Sub FuncAdviceRISPrintSch()
'���ܣ�RISҽ��ԤԼ����ӡ
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    lngResult = -1
    If HaveRIS Then
        With vsAdvice
            If Not .TextMatrix(.Row, COL_�������) = "D" Then
                MsgBox "��ǰҽ������Ӱ������Ŀ��", vbInformation, gstrSysName
                Exit Sub
            End If
            If .TextMatrix(.Row, COL_RISԤԼID) = 0 Then
                MsgBox "��ǰӰ����ҽ��û�б�ԤԼ�����ܴ�ӡ��", vbInformation, gstrSysName
                Exit Sub
            End If
            lngResult = gobjRis.HISPrintOneRisScheduleRpt(Val(.TextMatrix(.Row, COL_ID)))
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAdviceReportTip(ByVal lngRow As Long) As String
'���ܣ���ȡ���������ʾ��
    Dim strTmp As String
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_RIS����ID)) <> 0 Then
            strTmp = "(RIS����)"
        ElseIf Val(.TextMatrix(lngRow, COL_����ID)) <> 0 Then
            strTmp = "(HIS����)"
        ElseIf .TextMatrix(lngRow, COL_��鱨��ID) <> "" Then
            strTmp = "(רҵ��PACS����)"
        ElseIf Val(.TextMatrix(lngRow, COL_LIS����ID)) <> 0 Then
            strTmp = "(����LIS����)"
        Else
            If Val(.TextMatrix(lngRow, COL_RISԤԼID)) <> 0 Then
                If Val(.TextMatrix(lngRow, COL_RISԤԼ״̬)) = 0 Then
                    strTmp = "�Ѿ�ԤԼ"
                Else
                    strTmp = "�Ѿ�ԤԼ����"
                End If
            End If
        End If
        If strTmp <> "" And Val(.TextMatrix(lngRow, COL_RISԤԼID)) = 0 Then
            If Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 0 Then
                strTmp = "����δ��" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 1 Then
                strTmp = "��������" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 2 Then
                strTmp = "���沿������" & strTmp
            End If
        End If
    End With
    GetAdviceReportTip = strTmp
End Function

Private Sub FuncApplyCustom(ByVal intType As Long, ByVal lng�ļ�ID As Long)
'���ܣ��Զ������뵥
'������intType=0 ������=1�޸ģ�=2�鿴
    Dim lng������� As Long
    Dim datTurn As Date
    Dim lngRow As Long
    Dim lng��������ID As Long
    Dim lngNo As Long
    Dim objApplyCustom As New frmApplyCustom
    
    If intType <> 2 Then
        '���ҺŲ����Ƿ���
        If Not FuncTimeLimitCheck Then Exit Sub
        '�޸�ʱ����Ƿ����
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 2 Then
                MsgBox "���뵥�Ѿ���ˣ����������޸ġ�", vbInformation, "���뵥"
                intType = 2
            End If
        End If
    End If
    
    If intType <> 0 Then
         lng������� = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�������))
         lngRow = vsAdvice.Row
    End If
    
    If objApplyCustom.ShowMe(mfrmParent, 1, intType, mlng����ID, mstr�Һŵ�, 1, lng�ļ�ID, lng�������, mlng�Һſ���ID, IIF(mlng�������ID = 0, mlng�Һſ���ID, mlng�������ID), , mrsDefine, , , 0, mclsMipModule, mlngǰ��ID, , mint����) Then
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
    End If
End Sub

Private Sub FuncAdviceRISModi()
'���ܣ�����RISԤԼ
    Dim lngҽ��ID As Long
    Dim lngԤԼID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        lngԤԼID = Val(.TextMatrix(.Row, COL_RISԤԼID))
    End With
    
    strSQL = "select 1 from ����ҽ������ a where a.ҽ��id=[1] and nvl(a.ִ��״̬,0) in (0,3) and nvl(a.ִ�й���,0)<=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        If HaveRIS(False) Then
            Call gobjRis.HISReSchedule(lngҽ��ID, lngԤԼID)
        End If
    Else
        MsgBox "����Ŀ�Ѿ�ִ�У�����������������", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceIndexBill()
'���ܣ���ӡָ����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng���ͺ� As Long
    
    On Error GoTo errH

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 Then
            strSQL = "select a.���ͺ� from ����ҽ������ a where a.ҽ��id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
            If Not rsTmp.EOF Then
                lng���ͺ� = Val(rsTmp!���ͺ� & "")
            End If
        End If
    End With
    '��ӡָ����
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "���ͺ�=" & lng���ͺ�, "����ID=" & mlng����ID, "�Һŵ�=" & mstr�Һŵ�, "PrintEmpty=0", 2)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintBloodReport(ByVal lngAdviceID As Long, objFrm As Object)
    '��Ѫִ�е���ӡ
    If InitObjBlood(True) = True Then
        Call gobjPublicBlood.ShowBloodInstantRptPrint(objFrm, lngAdviceID)
    End If
End Sub

Private Sub SetAdviceIcon(ByVal lngRow As Long)
'���ܣ����ݵ�ǰ�е���������ҽ�����ݵ�ͼ���ʶ
'˵����ע���ǵ������ã�����һ������
    Dim intͼ���� As Integer 'ҽ�����������ͼ�����
    
    With vsAdvice
        '����ǩ����ʶ
        If Val(vsAdvice.TextMatrix(lngRow, COL_ǩ����)) = 1 Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgSign.ListImages("ǩ��").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgSign.ListImages("ǩ��").Picture
            intͼ���� = 1
        End If
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_��ΣҩƷ)) > 0 Then
            If vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) Is Nothing Then
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture
                intͼ���� = 1
            Else
                If vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) <> frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture Then
                    pictmp.Cls
                    pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                    Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                    Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                    intͼ���� = 2
                End If
            End If
        End If
        
        'Σ��ֵͼ��
        If Val(vsAdvice.TextMatrix(lngRow, COL_Σ��ֵID)) > 0 Then
            If intͼ���� = 0 Then
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture
            ElseIf intͼ���� = 1 Then
                pictmp.Cls
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                intͼ���� = 2
            ElseIf intͼ���� = 2 Then
                pictmp.Cls
                pictmp.Width = 720
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, 480, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture, 480, 0, 240, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                pictmp.Width = 480
                intͼ���� = 3
            End If
        End If
        
        '�׵���ͼ��
        If Val(vsAdvice.TextMatrix(lngRow, COL_�׵���)) > 0 Then
            If intͼ���� = 0 Then
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgQuestion.ListImages("�׵���").Picture
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgQuestion.ListImages("�׵���").Picture
            ElseIf intͼ���� = 1 Then
                pictmp.Cls
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                intͼ���� = 2
            ElseIf intͼ���� = 2 Then
                pictmp.Cls
                pictmp.Width = 720
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, 480, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, 480, 0, 240, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                pictmp.Width = 480
                intͼ���� = 3
            ElseIf intͼ���� = 3 Then
                pictmp.Cls
                pictmp.Width = 960
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, 720, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, 720, 0, 240, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                pictmp.Width = 480
                intͼ���� = 4
            End If
        End If
    End With
End Sub

Private Sub FuncCriticalAdvice(ByVal strPar As String, ByVal blnCheck As Boolean)
'���ܣ����ã�����/ȡ����Σֵҽ������
'������strPar-��ʽ��Σ��ֵID,ҽ��ID(��ҽ��ID)
'      blnCheck-true ȡ����ϵ��false ���ù�ϵ
    Dim lngΣ��ֵID As Long
    Dim lngҽ��ID As Long
    Dim lng���� As Long
    Dim strSQL As String
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim i As Long
    Dim lngOtherΣ��ֵID As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    lng���� = IIF(blnCheck, 2, 1)
    lngΣ��ֵID = Split(strPar, ",")(0)
    lngҽ��ID = Split(strPar, ",")(1)
    strSQL = "Zl_����Σ��ֵҽ��_Update(" & lng���� & "," & lngΣ��ֵID & "," & lngҽ��ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If blnCheck Then
        'ͬһ��ҽ���ɹ������Σ��ֵ��ȡ��ʱҪ��һ���ж��Ƿ��й���
        strSQL = "select a.Σ��ֵID,a.ҽ��ID from ����Σ��ֵҽ�� a where a.ҽ��ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
        If Not rsTmp.EOF Then
            lngOtherΣ��ֵID = rsTmp!Σ��ֵID & ""
        End If
    End If
    
    
    If RowInһ����ҩ(vsAdvice.Row, lngBegin, lngEnd) Then
        For i = lngBegin To lngEnd
            Set vsAdvice.Cell(flexcpPicture, i, col_ҽ������) = Nothing
            Set vsAdvice.Cell(flexcpPicture, i, col_����) = Nothing
            If blnCheck Then
                vsAdvice.TextMatrix(i, COL_Σ��ֵID) = lngOtherΣ��ֵID
            Else
                vsAdvice.TextMatrix(i, COL_Σ��ֵID) = lngΣ��ֵID
            End If
            Call SetAdviceIcon(i)
        Next
    Else
        '���½�����ͼ��
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_ҽ������) = Nothing
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_����) = Nothing
        If blnCheck Then
            vsAdvice.TextMatrix(vsAdvice.Row, COL_Σ��ֵID) = lngOtherΣ��ֵID
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, COL_Σ��ֵID) = lngΣ��ֵID
        End If
        Call SetAdviceIcon(vsAdvice.Row)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCriticalAdvice(ByRef lngҽ��ID As Long) As ADODB.Recordset
'���ܣ����ݵ�ǰѡ���е�ҽ����ѯ����֮������Σ��ֵ��¼
'���������� lngҽ��ID ����ǰ������ѡ��ҽ������ҽ��ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
    End With
    
    strSQL = "select a.Σ��ֵID,a.ҽ��ID from ����Σ��ֵҽ�� a where a.ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    Set GetCriticalAdvice = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCriticalData()
'���ܣ���ȡΣ��ֵ��¼
    Dim strSQL As String
    On Error GoTo errH
    If mblnΣ��ֵ Then
        strSQL = "select a.id,a.Σ��ֵ���� from ����Σ��ֵ��¼ a where a.�Һŵ�=[1] order by a.����ʱ�� desc"
        Set mrsΣ��ֵ = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mstr�Һŵ�)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FuncPathAdd() As Boolean
    Dim strSQL As String
    Dim str��ǰ���� As String
    Dim i As Long
    Dim lng����ID As Long, lng���ID As Long
    Dim bln��ҽ As Boolean
    Dim blnDo As Boolean, blnIsCancel As Boolean
    Dim blnIsSend As Boolean, blnYes As Boolean
    Dim rsTmp As ADODB.Recordset, rsPath As ADODB.Recordset
    Dim objDiagEdit As zlMedRecPage.clsDiagEdit
    
    
    '·���еĲ��ˣ�����û������·����Ŀ�����ȵ�������
    If mlng·��״̬ = 1 And mvarCond.Ӥ�� <= 0 Then
        blnDo = True
        If mint���� = 2 Then
            blnDo = zlDatabase.GetPara("ҽ��ҽ����·������", glngSys, P����·��Ӧ��, 0) = 0
        End If
        'δ����ʱ�������ҽ��������
        mblnNotEvaluete = Val(zlDatabase.GetPara("δ����ʱ�������ҽ��������", glngSys, P����·��Ӧ��, 1)) = 1
        If blnDo Then
            If CheckPathNotEvalueteOut(mlng�Һ�ID, blnIsSend, str��ǰ����) = False Then
                If gobjPathOut Is Nothing Then
                    MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ������¿�ҽ����", vbInformation, gstrSysName
                ElseIf InStr(GetInsidePrivs(P����·��Ӧ��), ";����·��;") = 0 Then
                    MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ���û������·����Ȩ�ޣ������¿�ҽ����", vbInformation, gstrSysName
                Else
                    '֮ǰ����û�н���·��ҳ�棬��Ҫ��ͨ��ˢ�½ӿڶ�ȡ��ʼ����
                    Call gobjPathOut.zlRefresh(mlng����ID, mlng�Һ�ID, mstr�Һŵ�, mlng�Һſ���ID, mint��������, mblnMoved, True)
                    Call gobjPathOut.zlExecPathSend(blnIsCancel)
                    Call LoadAdvice
                End If
                If Not blnIsCancel Then Exit Function
             Else
                If Not blnIsSend Then
                    If gobjPathOut Is Nothing Then
                        MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ������¿�ҽ����", vbInformation, gstrSysName
                        Exit Function
                    ElseIf InStr(GetInsidePrivs(P����·��Ӧ��), ";����·��;") = 0 Then
                        MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ���û������·����Ȩ�ޣ������¿�ҽ����", vbInformation, gstrSysName
                        Exit Function
                    Else
                        '��������˲�����δ����ʱ�������ҽ�������죬����ʾ������ֱ�ӽ����������ɲ���
                        If mblnNotEvaluete Then
                            blnYes = MsgBox("��Ҫ���·������Ŀ��''" & str��ǰ���� & "'?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
                        End If
                        '���ѡ���������������ɲ�����ѡ�����������¿�·������Ŀ�� ��ǰ����
                        If blnYes = False Then
                            '֮ǰ����û�н���·��ҳ�棬��Ҫ��ͨ��ˢ�½ӿڶ�ȡ��ʼ����
                            Call gobjPathOut.zlRefresh(mlng����ID, mlng�Һ�ID, mstr�Һŵ�, mlng�Һſ���ID, mint��������, mblnMoved, True)
                            'û�����ɣ��򷵻�false��ֹ�¿�����
                            If Not gobjPathOut.zlExecPathSend Then
                                Call LoadAdvice
                                Exit Function
                            End If
                            Call LoadAdvice
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    FuncPathAdd = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncViewLisRpt()
'���ܣ�������鱨��
'˵����������ģʽ�����жϱ��ξ����Ƿ���PDF����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If mblnMoved Then
        strSQL = "select 1 from H����ҽ����¼ a,H����ҽ������ b,Hҽ���������� c where a.id=b.ҽ��id and b.����id=c.id and c.����  in (0,2) and a.�Һŵ�=[1]"
    Else
        strSQL = "select 1 from ����ҽ����¼ a,����ҽ������ b,ҽ���������� c where a.id=b.ҽ��id and b.����id=c.id and c.����  in (0,2) and a.�Һŵ�=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    
    If Not rsTmp.EOF Then
        '����ҳǩ��ʾ
        Call frmLisALL.ShowMe(mfrmParent, mlng����ID, mlng�Һ�ID, mlng�Һſ���ID, 0, p����ҽ���´�, mMainPrivs)
    Else
        '��ǰ����ģʽ
        Call InitObjLis(p����ҽ��վ)
        If Not gobjLIS Is Nothing Then
            gobjLIS.PatientSampleBrowse mfrmParent, mlng����ID, mMainPrivs, mlng�Һſ���ID, 0, 1
        Else
            frmLisView.ShowMe mlng����ID, p����ҽ���´�, mfrmParent
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDrugRefcom()
'���ܣ�������д�ܾ�������ɴ��ڵ��ú�����ҩ�����ӿ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strAdviceIDs As String
    Dim strErr As String
    
    On Error GoTo errH
    
    strSQL = "select 1 from ����ҽ����¼ a where a.�Һŵ�=[1] and a.ҽ��״̬=1 and a.������� in ('5','6') and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    If Not rsTmp.EOF Then
        '���¿���ҩƷҽ��
        Call gobjPass.ZLPharmReviewResultOut(mfrmParent, mlng����ID, mlng�Һ�ID, mstr�Һŵ�, "", rsTmp, strErr)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set�걾״̬()
'���ܣ��Լ���ҽ�����ñ걾״̬�У������LIS�����з���
    Dim i As Long, strҽ��IDs As String, strMsg As String
    Dim rsAdvice As ADODB.Recordset
    Dim strIDAndRow As String, strTmp As String
    Dim lngRow As Long
    
    On Error GoTo errH
    
    If mvarCond.����ģʽ <> 3 Then Exit Sub
    Call InitObjLis(p����ҽ��վ)
    If gobjLIS Is Nothing Then Exit Sub
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" And Val(.TextMatrix(i, COL_���ID)) = 0 And Val(.TextMatrix(i, COL_ҽ��״̬)) = 8 Then
                strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                strIDAndRow = strIDAndRow & "," & Val(.TextMatrix(i, COL_ID)) & ";" & i & "<Tab>"
            End If
        Next
        If strҽ��IDs <> "" Then
            Set rsAdvice = gobjLIS.GetSampleType(Mid(strҽ��IDs, 2), strMsg)
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
            If Not rsAdvice Is Nothing Then
                rsAdvice.Filter = 0
                For i = 1 To rsAdvice.RecordCount
                    If InStr(strIDAndRow, "," & rsAdvice!ҽ��ID & ";") > 0 Then
                        strTmp = Split(strIDAndRow, "," & rsAdvice!ҽ��ID & ";")(1)
                        lngRow = Val(Split(strTmp, "<Tab>")(0))
                        .TextMatrix(lngRow, COL_�걾״̬) = rsAdvice!ҽ��״̬ & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewPacsRpt()
'���ܣ�������鱨��
'˵����δ�����Ķ����
    Dim blnAutoRead As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngҽ��ID As Long
    
    On Error GoTo errH
    Call CreateObjectPacs(mobjPublicPACS)
    If Not mobjPublicPACS Is Nothing Then
        strSQL = "select max(b.id) as ҽ��ID  from ����ҽ������ a,����ҽ����¼ b " & _
                " Where a.��鱨��ID Is Not Null And a.ҽ��ID = b.ID And b.�Һŵ� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
        lngҽ��ID = Val(rsTmp!ҽ��ID & "")
        Call mobjPublicPACS.zlDocShowReport(lngҽ��ID, , blnAutoRead, mfrmParent)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
