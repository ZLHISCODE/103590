VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLabMain 
   Caption         =   "���鼼ʦ����վ"
   ClientHeight    =   6750
   ClientLeft      =   1515
   ClientTop       =   675
   ClientWidth     =   10995
   Icon            =   "frmLabMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmLabMain.frx":058A
   ScaleHeight     =   6750
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicWindows 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   7500
      ScaleHeight     =   585
      ScaleWidth      =   795
      TabIndex        =   29
      Top             =   2370
      Width           =   795
   End
   Begin VB.PictureBox picBarCodePrint 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4890
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   26
      Top             =   390
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSWinsockLib.Winsock WinsockC 
      Left            =   750
      Top             =   690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox cboExesItem 
      Height          =   300
      Left            =   3930
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1590
      Width           =   1875
   End
   Begin VB.ComboBox cboUnionItem 
      Height          =   300
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1590
      Width           =   1875
   End
   Begin VB.TextBox TxtGoto 
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1530
      Width           =   1755
   End
   Begin VB.ComboBox cboMachine 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1170
      Width           =   1905
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   1875
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   120
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":6DDC
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":7376
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":7910
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":7EAA
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":8444
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":89DE
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":8D78
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":9112
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":94AC
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":9846
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":100A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":1690A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":1D16C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":239CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":2A230
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":30A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":3102C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   1950
      Width           =   5865
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   765
         Left            =   300
         TabIndex        =   4
         Top             =   2460
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptList1 
         Height          =   765
         Left            =   2160
         TabIndex        =   12
         Top             =   2430
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   300
         Left            =   4350
         TabIndex        =   28
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   247267329
         CurrentDate     =   40049
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   300
         Left            =   2970
         TabIndex        =   27
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   247267329
         CurrentDate     =   40049
      End
      Begin VB.ComboBox cboʱ�� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3600
         Width           =   1275
      End
      Begin VB.PictureBox PicFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4500
         Picture         =   "frmLabMain.frx":3788E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         Top             =   22
         Width           =   240
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   23
         ToolTipText     =   "�����ֱ�ӵǼǱ걾"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "סԺ"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   825
         TabIndex        =   22
         ToolTipText     =   "סԺ�걾"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         ToolTipText     =   "û�в�����Ϣ�ı걾"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   2310
         TabIndex        =   20
         ToolTipText     =   "����˱걾"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "δ��"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   3045
         TabIndex        =   19
         ToolTipText     =   "δ��˱걾"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   3780
         TabIndex        =   18
         ToolTipText     =   "δ��˱걾"
         Top             =   30
         Width           =   735
      End
      Begin XtremeSuiteControls.TabControl TabList 
         Height          =   1575
         Left            =   90
         TabIndex        =   11
         Top             =   300
         Width           =   3525
         _Version        =   589884
         _ExtentX        =   6218
         _ExtentY        =   2778
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3990
      ScaleHeight     =   315
      ScaleWidth      =   4245
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   4245
      Begin XtremeCommandBars.CommandBars cbrChild 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox PicTab 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   6840
      ScaleHeight     =   1785
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   330
      Width           =   3675
      Begin XtremeSuiteControls.TabControl TabCtlWindow 
         Bindings        =   "frmLabMain.frx":3E0E0
         Height          =   1575
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   3525
         _Version        =   589884
         _ExtentX        =   6218
         _ExtentY        =   2778
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   6390
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabMain.frx":3E0F4
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14314
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
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin RichTextLib.RichTextBox RtfTxt 
      Height          =   885
      Left            =   2190
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1561
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmLabMain.frx":3E988
   End
   Begin VB.PictureBox PicImage 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   6540
      ScaleHeight     =   2595
      ScaleWidth      =   1935
      TabIndex        =   15
      Top             =   3450
      Width           =   1935
      Begin VB.VScrollBar VScroll 
         Height          =   1245
         Left            =   1620
         Max             =   0
         TabIndex        =   17
         Top             =   150
         Width           =   225
      End
      Begin C1Chart2D8.Chart2D ChartThis 
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   885
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1561
         _ExtentY        =   1296
         _StockProps     =   0
         ControlProperties=   "frmLabMain.frx":3EA17
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   1440
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmLabMain.frx":3EF9A
      Left            =   810
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Dkp_ID_List As Integer = 1                            '�����嵥����
Private Const Dkp_ID_Locate As Integer = 2                          '��λ���Ҵ���
Private Const Dkp_ID_Request As Integer = 3                         '�˶ԵǼǴ���
Private Const Dkp_ID_Append As Integer = 4                          '���桢���١����õȸ��Ӵ���
Private Const Dkp_ID_Image As Integer = 5                           '��ʾ����ͼ��
Private Const pҽ�����ѹ��� As Integer = 1257                       '���˷���ģ����Ȩ
Private Const p����ҽ���´� As Integer = 1252                       '����ҽ���´�
Private Const pסԺҽ���´� As Integer = 1253                       'סԺҽ���´�
Private Const p���ﲡ������ As Integer = 1250                       '���ﲡ��
Private Const pסԺ�������� As Integer = 1251
Private Const p�°没������ As Integer = 2250                       '�°没��

Private Const ID_MENU_MOUSE = 90                                    '�Ҽ��˵�
Private Const con_������ɸѡ_������ As String = "���ﲡ��;סԺ����;�����걾;����걾;δ��걾;��첡��;����ҽ��;�����걾;�ʿر걾;�����ͨ��;���δͨ��;δ����;������;�������ͨ��;�������δͨ��"
Private Const con_������ɸѡ_������ As String = "���ﲡ��;סԺ����;��첡��"
'-----------------���������Ĵ���--------------------
Private WithEvents mfrmRequest As frmLabRequest                     '���յǼǴ���
Attribute mfrmRequest.VB_VarHelpID = -1
Private WithEvents mfrmWrite As frmLisStationWrite                  '������д����
Attribute mfrmWrite.VB_VarHelpID = -1
Private WithEvents mfrmWrite2 As frmLisStationWrite2                '��д΢����
Attribute mfrmWrite2.VB_VarHelpID = -1
Private WithEvents mfrmLabMainSampleUnion  As frmLabMainSampleUnion '�걾�ϲ�
Attribute mfrmLabMainSampleUnion.VB_VarHelpID = -1
Private WithEvents mclsInAdvices As zlCISKernel.clsDockInAdvices    'סԺҽ��
Attribute mclsInAdvices.VB_VarHelpID = -1
Private WithEvents mclsOutAdvices As zlCISKernel.clsDockOutAdvices  '����ҽ��
Attribute mclsOutAdvices.VB_VarHelpID = -1
Private mclsInEPRs As zlRichEPR.cDockInEPRs               'סԺ����
Attribute mclsInEPRs.VB_VarHelpID = -1
Private mclsOutEPRs As zlRichEPR.cDockOutEPRs             'סԺ����
Attribute mclsOutEPRs.VB_VarHelpID = -1
Private mfrmTrack As frmLabTrack                                    '���ζԱ�
Private WithEvents mfrmLabMicrobe3Report As frmLabMicrobe3Report    '��������
Attribute mfrmLabMicrobe3Report.VB_VarHelpID = -1

'Private mfrmLabMainImage As frmLabMainImage                        '����ͼ����ʾ
Private WithEvents mclsExpenses As zlPublicExpense.clsDockExpense        '�µķ���\
Attribute mclsExpenses.VB_VarHelpID = -1
Private mclspublicExpenses As zlPublicExpense.clsPublicExpense        '�µķ��ò���

Private mcolSubForm As Collection                                   'ж���Ӵ���
Private mblnCompelRefresh As Boolean                                'ǿ��ˢ��
Private mintUnion As Integer                                        '�Ƿ���������������ʾ 0=������ 1=����
Private mSendReport As Integer                                      '��˺��Ƿ��Զ����ͱ��� 0=���� 1=������

'-----------------------------------------------------
'-----------------------------------------------------
Private mlngDeptID As Long                                          '����ID
Private mlngKey As Long                                             '����걾ID
Private mintEditState As Integer                                    '��ǰ�༭״̬��0-�Ǳ༭��1-�������գ�2-�����Ǽǣ�4-����ˣ�3-���º��գ�5-����༭;6-�걾�ϲ�;7-��������
Private mintHandleState As Integer                                  '��ǰ����״̬:1 = ������Ϣ = 2���浥���� 3= ��������
Private mintContinue As Integer                                     'Ŀǰ�Ƿ����������յǼ�״̬


Private mstrPrivs As String                                         'Ȩ��
'Private objLISComm As Object                                       'ͨѶ�ӿ�



'-----------------------------------------------------
'----------------------�������ñ���-------------------
Dim blnChecking As Boolean                                          '�Ƿ����ڽ��в���
Dim blnAutoRefresh As Boolean                                       '�Ƿ����յ���������ʱ�Զ�ˢ��
Dim blnComm As Boolean                                              '�Ƿ�����˫��ͨ��
Dim blnAutoPrint As Boolean                                         '��˺��Զ���ӡ
'-----------------------------------------------------
'---------------------���ʱ�ж�----------------------
Private mintAuditing As Integer                                     '�Ƿ������Ȩ�� 0=û��Ȩ�� 1=��Ȩ��
                                                                    '-1��-24=��Чʱ�����ʱȡ����ֵ
Private mDataAuditing As Date                                       '��ʱ���޶���,��¼ʱ��
Private mstrAuditingMan As String                                   '�����,���ʱ�������
Private mstrAuditingManID As String                                 '����˵�½��(ǩ����)
Private mblnCancel As Boolean                                       'ȡ��ˢ��
Private mUserDept As String                                         '�û����������ִ�
Private mblnVerifying(15) As Boolean                                '������ɸѡ״̬
Private mblnWaitVerify(2) As Boolean                                '�ȴ�����

Private mMakeNoRule As String                                       '�걾������ɵ����ڹ���
Private mstrMachines As String                                      '��¼�в���Ȩ�޵�����
Private mstrMachineALL As String                                    '��¼������ʾ������ID�ִ�


Private mbln�ֹ����ͱ��� As Boolean                                 '�ֹ�����
Private mbln�����ֱ����� As Boolean                               '��¼�Ƿ񱣴���Զ����
Private mstrPrintDepts As String                                    '���Դ�ӡ�Ŀ���
Private mblnAout As Boolean                                         '�Ƿ��Զ�����һ������˵ı걾
Private mlngLastShow As Long                                        '�����ʾ�ı걾��ID
Private mTodayQCPrivs As String                                     '�����ʿ�Ȩ��
Private mHistoryPrivs As String                                     '��ʷ�ʿ�Ȩ��
Private mTableRefresh As Boolean                                    'Table�Ƿ�ˢ��
Private mintLoadShow As Integer                                     '��ʾ��ֵ����1

Const mcontIntRowHeight As Integer = 230                            '��¼�и�

Public mblnSendComplete As Boolean                                  '�Ƿ���ɴ���
Public mstrMachineGroup As String                                   '��������
Public mlngMachineID As Long                                        'ѡ������ID -1=�ֹ� 0=�������� >0=����ID
Private mstrMachineID  As String                                    '����΢��������ID��

Dim mclsEMR As Object                                               '�°���Ӳ���


Dim mblnTabList1 As Boolean                                         '�ؼ���ҳ�Ƿ�ʹ�ù� 0=δʹ�� 1=��ʹ��


'==�������б�
Private Enum mCol
    ID = 0
    ����
    ����ҽ��
    ִ��״̬
    �������
    �걾����
    �걾��
    ����
    �Ա�
    ����
    ������Ŀ
    ��ʶ��
    ����
    �������
    ҽ��id
    ����id
    ת��
    ����ID
    �걾ʱ��
    ����ʱ��
    ΢����걾
    �շѵ�
    �Һŵ�
    ������
    �����
    ��������
    Ӥ��
    ���˿���
    ���ͺ�
    ������
    ��ҳID
    ��������ID
    ������
    ��������
    ���䵥λ
    ����
    ������
    �걾��̬
    ������
    ����ʱ��
    ����걾
    NO
    ������
    ����ʱ��
    ���ʱ��
    ����id
    ��������
    ��λ
    ִ�п���ID
    �걾���
    ҽ������
    �걾����
    �������
    ��������
    ����
    ����״̬
    ���淢��
    ���˿���ID
    ������
    ����ʱ��
    ��λ
    ������
    ���δͨ��
    ������Դ
    �����
    סԺ��
    ���Ϊ��
    �ٴ�·������
    �������
End Enum

'==�������б�
Private Enum mRCol
    ����ID
    ����
    ��Դ
    ����
    �Ա�
    ����
    ���˿���
    ��ʶ��
    ����
    ҽ������
    ����ҽ��
    ����ʱ��
    ������ĿID
    ҽ��id
    ִ��״̬
    ��λ
    �Һŵ�
    ǩ��ʱ��
End Enum

'==�걾����
Private Enum mActS
    ���� = 0: �Ǽ�: ���º���: �����
    �޸�������
    �����޸�������
    ɾ�������걾
    ��������
    �������͵�����
    ��Ϊ�ʿ�
    ��Ϊ�Ա�
    ״̬�ع�
    ��������
    ����
    ��Ϊ����
    �ϲ��걾
    �ϲ��걾����
    �޸Ĳ�����Ϣ
End Enum
'==�������
Private Enum mActR
    ������������ = 0
    ��˱���
    ���ͱ���
    ������˱���
    ���ȡ��
    �������
    ȡ������
    ��д����
    д�벡��
    ��֤ǩ��
    ���������
    ��д��������
End Enum
'���桢���ա�����
Private Enum mFileS
    ����
    ����
End Enum
Private Enum mFilter                        '��������
    ���� = 0
    �Ա�
    ����
    ���䵥λ
    ��ʶ��
    �걾��
    ���ݺ�
    �������
    ������
    ������Ŀ
    ����ʱ��
    �ͼ����
    �ͼ���
    ��������
    ϸ��
    ������
    ҩ�����
    �Ƿ�ʹ�ø߼�
    �߼�
    ����ID
End Enum
Private Enum mSWork                         '���ڼ��̿�ݲ���
    Key_PageUP
    Key_PageDown
    Key_Home
    Key_End
End Enum
Private Const conMenu_IDkind_Change  As Integer = 12345
Private int��촦��ʽ As Integer  '1-��ʾ��2-������3-������
Private int���ﴦ��ʽ As Integer  '1-��ʾ��2-������3-������
Private intסԺ����ʽ As Integer  '1-��ʾ��2-������3-������
Private intԺ�⴦��ʽ As Integer  '1-��ʾ��2-������3-������

'--------------------------------------------
'�����ض���
Implements zl9LisQuery_Def.clsLisQueryHost
Private clsPluginLoader  As PlugInLoader
Private mobjPlugin()   As zl9LisQuery_Def.clsLisQuery
'--------------------------------------------
Private Sub CreateCbs()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False
    

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&T)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "������ӡ(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintBedCard, "�ش�����(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "����(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "�嵥��ӡ(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Privacy, "����˵�½(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&O)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "ø����(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&Y)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "����(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Ǽ�(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "��������(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "������������(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "�����(&A)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "���º���(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸�������(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "�����޸�������(&D)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBound, "�޸Ĳ�����Ϣ(&P)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "�������͵�����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_BathSend, "�������͵�����(&S)"): cbrControl.BeginGroup = True

        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "��Ϊ�ʿ�(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "��Ϊ�ȶ�(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "�鿴�ȶ�(&B)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCRes, "�鿴�����ʿ�(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, comMenu_LIS_TodayQC, "�����ʿ�(&T)")
        Set cbrControl = .Add(xtpControlButton, comMenu_LIS_History, "��ʷ�ʿ�(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_LJAverage, "��ֵLJ�ʿ�(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "״̬�ع�(&Z)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "ɾ������(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "����ɾ������(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "ȡ������(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Insert, "�걾�ϲ�(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "�������ϲ�(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "����(&J)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Statistics_PositiveResults, "���Խ������(&Y)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Statistics_Feedback, "���������ѯ(&Z)")
    End With
    'conMenu_EditPopup
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&E)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Report, "������д(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "��������(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Dilute, "�걾ϡ��(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SendReport, "���󱨸�(&S)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Audit, "�������(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ȡ�����(&U)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Seat_Set, "���������(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "��������(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ������(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "�����ѯ(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportFromXML, "���������ռ�(&G)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SignVerify, "��֤ǩ��(&S)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "�Զ�����(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "��������(&L)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SaveSample, "�걾����(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_DropSample, "�걾����(&H)")
    
    End With
'    '�Ҽ��˵�
'    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_LIS_RightMenu, "�Ҽ��˵�", -1, False)
'    cbrMenuBar.ID = conMenu_LIS_RightMenu
'    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������(&A)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ȡ�����(&U)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "״̬�ع�(&Z)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "��������(&D)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ������(&E)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "�����ѯ(&P)"): cbrControl.BeginGroup = True
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "��������(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "��Ϊ�ʿ�(&Q)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "��Ϊ�ȶ�(&Y)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "�鿴�ȶ�(&B)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "�������ϲ�(&E)")
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸�������(&M)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "ɾ������(&D)")
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "����(&J)"): cbrControl.BeginGroup = True
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "����(&C)")
'    End With
'    cbrMenuBar.Visible = False

'    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&C)", -1, False)
'    cbrMenuBar.ID = conMenu_EditPopup
'    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Price, "��������(&P)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "���ӷѻ���(&I)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "���ӷѼ���(&S)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "��Ѽ�¼(&Z)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "�޸ĸ��ӷ�(&M)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "ɾ�����ӷ�(&D)")
'    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "ǰһ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "��һ��(&L)")
        '----------------------------------------------------------------------------------------------------------------
        '���ڿ�ݷ�ʽ����(PageUP,PageDown,Home,End)
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Reference_1, "PageUP"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Reference_2, "PageDown"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_MeetFinish, "Home"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_MeetCancel, "End"): cbrControl.Visible = False
        '----------------------------------------------------------------------------------------------------------------
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "��δ�շ�(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_HideList, "�����б�(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, comMenu_LIS_ShowListHead, "ѡ����ʾ�б�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "��ʾ������"): 'cbrControl.BeginGroup = True
        If zlDatabase.GetPara("��ʾ������", 100, 1208, "False") = "True" Then
            cbrControl.Checked = True
        End If

        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_LeaveMedi, "���ؼ���ͼ��"): 'cbrControl.BeginGroup = True
        
        If zlDatabase.GetPara("���ؼ���ͼ��", 100, 1208, "True") = "True" Then
            cbrControl.Checked = True
        End If
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "�б�ѡ��(&O)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_PatientInfo, "������Ϣ(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "��λ(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "���ٹ���(&K)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "��ϲ�ѯ(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_FindNext, "�������μ���(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportEdit, "�걾������־(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)"): cbrControl.BeginGroup = True
    End With
'-------------------------------------------------------------------------------------------------------------------------------------
    '�ۺϲ�ѯ����˵�
    Dim i           As Long
    ReDim mobjPlugin(0) As zl9LisQuery_Def.clsLisQuery
    If Not clsPluginLoader Is Nothing Then
        clsPluginLoader.FindPlugins
        If clsPluginLoader.PluginCount > 0 Then
            ReDim mobjPlugin(clsPluginLoader.PluginCount) As zl9LisQuery_Def.clsLisQuery
            Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PlugPopup, "��ӳ���(&A)", -1, False)
            cbrMenuBar.ID = conMenu_PlugPopup
            With cbrMenuBar.CommandBar.Controls
                
                For i = 0 To clsPluginLoader.PluginCount - 1
                    Set mobjPlugin(i) = clsPluginLoader.CreatePlugin(i)
                    If Not mobjPlugin(i) Is Nothing Then
                        mobjPlugin(i).Index = i
                        Set cbrControl = .Add(xtpControlButton, conMenu_PlugPopup * 1000# + 100 + i, mobjPlugin(i).Name)
                    End If
                    If i = 0 Then cbrControl.BeginGroup = True
                Next
            End With
        End If
    End If
'-------------------------------------------------------------------------------------------------------------------------------------

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    Set cbrControl = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "�������")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_DrugQuery, "�������")
    cbrCustom.ShortcutText = "����"
    cbrCustom.Handle = Me.cboDept.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Edit_NoPrint, "����С��")
    cbrMenuBar.ID = conMenu_Edit_NoPrint
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Owe, "����С��")
    End With
    cbrMenuBar.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_Reports, "��������")
    cbrCustom.ShortcutText = "��������"
    cbrCustom.Handle = Me.cboMachine.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption

    Set cbrControl = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "�����Ŀ")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_WorkLog, "��������")
    cbrCustom.ShortcutText = "��������"
    cbrCustom.Handle = Me.cboUnionItem.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption

    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Save
        .Add 0, VK_ESCAPE, conMenu_LIS_Cancel
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F4, conMenu_Manage_Plan
        .Add 0, VK_F8, conMenu_Manage_Regist
        .Add FCONTROL, Asc("T"), conMenu_Tool_Apply
        .Add FCONTROL, Asc("Z"), conMenu_Edit_SendBack
        .Add FCONTROL, VK_DELETE, conMenu_Manage_ClearUp
        .Add 0, VK_F7, conMenu_Manage_Report
        .Add 0, VK_F6, conMenu_Edit_Audit
        .Add FCONTROL, VK_LEFT, conMenu_View_Backward
        .Add FCONTROL, VK_RIGHT, conMenu_View_Forward
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add FCONTROL, Asc("F"), conMenu_Manage_Transfer_Force
        .Add 0, VK_F3, conMenu_View_Filter
        .Add 0, VK_HOME, conMenu_Tool_MeetFinish
        .Add 0, VK_END, conMenu_Tool_MeetCancel
        .Add 0, VK_PAGEUP, conMenu_Tool_Reference_1
        .Add 0, VK_PAGEDOWN, conMenu_Tool_Reference_2
        .Add FCONTROL, Asc("H"), conMenu_View_FindNext
        .Add 0, VK_F9, conMenu_Edit_QCRes
        .Add 0, VK_F11, conMenu_Manage_Logout
        
        .Add 0, VK_F10, conMenu_IDkind_Change
    End With

    '���ò����ò˵�
'    With Me.cbrthis.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'        .AddHiddenCommand conMenu_File_Excel
'        .AddHiddenCommand conMenu_View_Jump
'        .AddHiddenCommand conMenu_View_Refresh
'    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Ǽ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "�ع�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Insert, "�ϲ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Report, "���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SendReport, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�󱨸�")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Price, "������"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next


End Sub

Private Sub cboDept_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim intLoop As Integer
    Dim blnSelect As Boolean                    '�Ƿ�ѡ��
    Dim strCoding As String                     'С�����
    On Error GoTo errH
    mstrMachines = ""
    mstrMachineALL = ""
    rptList.Tag = ""
    If cboDept.ListCount > 0 Then
        'д��ˢ��
        mlngDeptID = Val(cboDept.ItemData(cboDept.ListIndex))
        
        gstrSql = "Select Distinct A.����, A.����" & vbNewLine & _
                "From ����С�� A, ����С������ B, �������� C" & vbNewLine & _
                "Where A.ID = B.С��id And B.����id = C.ID And C.ʹ��С��id = [1] order by a.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
        
        Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_NoPrint, True, True)
        With cbrMenuBar.CommandBar.Controls
            .DeleteAll
            .Add xtpControlButton, conMenu_View_Owe, "����С��"
             If Not cbrMenuBar Is Nothing Then
                 Do Until rsTmp.EOF
                    Set cbrControl = .Add(xtpControlButton, conMenu_View_Owe, Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����")))
                    cbrControl.Checked = (Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����")) = mstrMachineGroup)
                    If cbrControl.Checked = True Then
                        blnSelect = True
                    End If
                    rsTmp.MoveNext
                 Loop
             End If
            If blnSelect = False Then
                cbrMenuBar.CommandBar.Controls(1).Checked = True
                mstrMachineGroup = "����С��"
            End If
        End With
        
        
'        objLISComm.DeptID = mlngDeptID
        
        cboMachine.Clear
        
        If cboDept.ListCount > 0 Then
        
            cboMachine.AddItem "<��������>": cboMachine.ItemData(cboMachine.NewIndex) = 0
            cboMachine.AddItem "<�ֹ�>": cboMachine.ItemData(cboMachine.NewIndex) = -1
            If InStr(mstrMachineGroup, "-") > 0 Then
                strCoding = Mid(mstrMachineGroup, 1, InStr(mstrMachineGroup, "-") - 1)
            End If
            If InStr(mstrPrivs, "���п���") > 0 Then
                strSQL = "Select Distinct A.����, A.ID, 1 As ����,c.���� ,A.΢����" & vbNewLine & _
                        "From �������� A, ����С������ B, ����С�� C" & vbNewLine & _
                        "Where A.ID = B.����id And A.ʹ��С��id = [1] And B.С��id = C.ID "
            Else
                strSQL = "Select Distinct D.ID, D.����, C.����,b.����,D.΢���� " & vbNewLine & _
                        " From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                        " Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [2] And C.����id = D.ID And D.ʹ��С��id = [1] "
            End If
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, UserInfo.ID, strCoding)
            
            If mstrMachineGroup <> "����С��" Then
                rsTmp.filter = "���� = '" & strCoding & "'"
            End If
            '���΢��������ID��
            mstrMachineID = ""
            Do Until rsTmp.EOF
                cboMachine.AddItem rsTmp("����")
                cboMachine.ItemData(cboMachine.NewIndex) = rsTmp("Id")
                If rsTmp("΢����") = 1 Then
                    mstrMachineID = mstrMachineID & rsTmp("id") & ","
                End If
                If rsTmp("id") = mlngMachineID Then
                    cboMachine.ListIndex = cboMachine.NewIndex
                End If
                
                rsTmp.MoveNext
            Loop
            If cboMachine.ListCount > 0 And Trim(cboMachine.Text) = "" Then
                cboMachine.ListIndex = 0
                mlngMachineID = cboMachine.ItemData(cboMachine.ListIndex)
            End If
            
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If rsTmp.EOF = False Then
                    rsTmp.filter = ""
                    rsTmp.MoveFirst
                    Do Until rsTmp.EOF
                        If Val(Nvl(rsTmp("����"))) = 1 Then
                            mstrMachines = mstrMachines & ";" & rsTmp("ID")
                        End If
                        mstrMachineALL = mstrMachineALL & "," & rsTmp("ID")
                        rsTmp.MoveNext
                    Loop
                End If
            End If
            If mstrMachines <> "" Then mstrMachines = mstrMachines & ";"
        Else
            mlngMachineID = 0
            mstrMachines = ""
            mstrMachineALL = ""
        End If
    Else
        mlngDeptID = 0
        mstrMachines = ""
        mstrMachineALL = ""
    End If
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
'    RefreshData
    
'    '����ˢ�º�λ��ָ����¼
'    On Error Resume Next
'    Me.dkpMain.FindPane(Dkp_ID_List).Select
'    Me.rptList.SetFocus
End Sub

Private Sub cboExesItem_Click()
    Dim lngAdvice As Long
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean                                             '�Ƿ�ת��
    On Error GoTo errH
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            blnCurrMoved = (.Record(mCol.ת��).Value = "��")
        End With
    End If
    
    strSQL = "select a.id as ҽ��ID, b.���ͺ� from ����ҽ����¼ a,����ҽ������ b " & vbCrLf & _
        " Where a.ID = b.ҽ��id And a.���id = [1] "
        
    If Me.cboExesItem.ListIndex <> -1 Then lngAdvice = Me.cboExesItem.ItemData(Me.cboExesItem.ListIndex)
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngAdvice)
    If rsTmp.EOF = False Then
        mclsExpenses.zlRefresh mlngDeptID, rsTmp(0) & ":" & rsTmp(1), blnCurrMoved
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboMachine_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    
    If cboMachine.ListCount > 0 Then
        'д��ˢ��
        mlngMachineID = cboMachine.ItemData(cboMachine.ListIndex)
    Else
        mlngMachineID = 0
    End If
    
'    strsql = "select distinct a.id,a.����,a.���� from ������ĿĿ¼ a,����ִ�п��� b,���鱨����Ŀ c , ����������Ŀ D " & _
             " where a.��� = 'C' and (a.�����Ŀ = 1 or a.����Ӧ�� = 1 ) " & _
             " and a.id = b.������Ŀid and a.id = c.������Ŀid and c.������Ŀid = d.��ĿID(+) "
    strSQL = "select distinct a.id,a.����,a.���� from ������ĿĿ¼ a,����ִ�п��� b,���鱨����Ŀ c , ����������Ŀ D " & _
             " where a.��� = 'C' and (a.�����Ŀ = 1 or a.����Ӧ�� = 1 ) " & _
             " and a.id = b.������Ŀid and a.id = c.������Ŀid(+) and c.������Ŀid = d.��ĿID(+) " & _
             " and (a.����ʱ�� is null or a.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) "
                 
    '�����ֹ��������������
    If cboMachine.ItemData(cboMachine.ListIndex) = -1 Then
        strSQL = strSQL & " And D.����ID is null "
    ElseIf cboMachine.ItemData(cboMachine.ListIndex) > 0 Then
        strSQL = strSQL & " And D.����ID = [1] "
    Else
        If Me.cboUnionItem.ListCount = 0 Then
            Me.cboUnionItem.Clear
            Me.cboUnionItem.AddItem "<���������Ŀ>"
            Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = 0
            Me.cboUnionItem.AddItem "<δ֪��Ŀ>"
            Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = -1
            If Me.cboUnionItem.ListCount > 0 Then Me.cboUnionItem.ListIndex = 0
        End If
        If Not TabCtlWindow(5).Visible Then TabCtlWindow(5).Visible = True
        Exit Sub
    End If
    
    '�������
    If cboDept.ItemData(cboDept.ListIndex) > 0 Then
        strSQL = strSQL & " And B.ִ�п���ID = [2] "
    End If
    
    strSQL = strSQL & " order by a.���� "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngMachineID, mlngDeptID, CDate("3000/1/1"))
    
    Me.cboUnionItem.Clear
    Me.cboUnionItem.AddItem "<���������Ŀ>"
    Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = 0
    Me.cboUnionItem.AddItem "<δ֪��Ŀ>"
    Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = -1
    
    Do Until rsTmp.EOF
        Me.cboUnionItem.AddItem rsTmp("����") & "-" & rsTmp("����")
        Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    '����΢����������ѡ��ʱ�걾�ϲ�ҳǩ����
    If InStr("," & mstrMachineID, "," & mlngMachineID & ",") > 0 Then
        TabCtlWindow(5).Visible = False
        '����ѡ���˸�ҳǩ�л�����һ��ҳǩ
        If TabCtlWindow(5).Selected Then
            TabCtlWindow(1).Selected = True
        End If
    Else
        TabCtlWindow(5).Visible = True
    End If
    If Me.cboUnionItem.ListCount > 0 Then Me.cboUnionItem.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
'    RefreshData
'
'    '����ˢ�º�λ��ָ����¼
'    On Error Resume Next
'    Me.dkpMain.FindPane(Dkp_ID_List).Select
'    Me.rptList.SetFocus
End Sub

Private Sub cboUnionItem_Click()
'    If Me.Visible = True Then
        On Error Resume Next
'        DoEvents
'        '�ȴ�������ʾ
'        Do While Me.Visible = False
'            DoEvents
'            Me.Show
'        Loop
        If mintLoadShow = 0 Then Exit Sub
        If Me.TabList.Item(1).Selected = True And Me.cboUnionItem.ListCount > 0 Then
            Call RefreshData1
        Else
            Call RefreshData
        End If
'    End If
End Sub

Private Sub cboʱ��_Click()
    If Me.Visible = False Then Exit Sub
    If Me.TabList(0).Selected = True Then
        zlDatabase.SetPara "�걾��Χ", cboʱ��.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Me.dtpDate.Visible = (Me.cboʱ��.Text = "�Զ���")
        Me.dtpDateEnd.Visible = (Me.cboʱ��.Text = "�Զ���")
        Call RefreshData
    Else
        zlDatabase.SetPara "�����շ�Χ", cboʱ��.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Me.dtpDate.Visible = (Me.cboʱ��.Text = "�Զ���")
        Me.dtpDateEnd.Visible = (Me.cboʱ��.Text = "�Զ���")
        Call RefreshData1
    End If
    'ˢ��
    
End Sub

Private Sub cbrChild_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrCbo As CommandBarComboBox
    
    On Error GoTo errH
    
    Select Case Control.ID
        Case conMenu_File_RoomSet                                                               '��λ
            If FindPatient(Control.Text) = True Then
                Control.Text = ""
            End If
            Control.SetFocus
            SendKeys "~"
        Case conMenu_View_Forward                                                               'ǰһ��
            BackOrNextPatient 1
        
        Case conMenu_View_Backward                                                              '��һ��
            BackOrNextPatient 2
        
        Case conMenu_Manage_RequestView                                                         'ʹ������ɨ��
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "ʹ������ɨ��", Control.Checked, 100, 1208
            
        Case conMenu_Manage_RequestPrint                                                        '��������
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "��������", Control.Checked, 100, 1208
            '�Ƿ���������
            mintContinue = IIf(Control.Checked, 1, 0)
            If mintContinue = 1 Then
                Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "�����Ǽ�"
                Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "��������"
            Else
                Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "�Ǽ�"
                Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "����"
            End If
            Me.cbrthis.RecalcLayout
        Case conMenu_Manage_RequestBatPrint                                                         '�����ֱ�����
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "�����ֱ�����", Control.Checked, 100, 1208
        Case XTP_ID_WINDOW_LIST '��ʾ��ע����
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "��ʾ���鱸ע", Control.Checked, 100, 1208
            Call mfrmWrite.Resize
            Call mfrmWrite2.Resize
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrChild_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    Left = -120
End Sub


Private Sub cbrChild_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errH
    
    Select Case Control.ID
        
        Case conMenu_Manage_RequestView                                                         'ʹ������ɨ��
            Control.Checked = Control.Checked
        
        Case conMenu_Manage_RequestPrint                                                        '��������
            Control.Checked = Control.Checked
        
        Case conMenu_Manage_RequestBatPrint                                                     '�����ֱ�����
            Control.Checked = Control.Checked
        Case XTP_ID_WINDOW_LIST                                                                 '��ʾ���鱸ע
            Control.Checked = Control.Checked
        
        Case conMenu_View_Forward                                                               'ǰһ��,��һ��
            If mintEditState <> 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.FocusedRow.Index = 0 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_View_Backward
            
            If mintEditState = 4 Or mintEditState = 5 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.Rows.Count - 1 = Me.rptList.FocusedRow.Index Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_File_RoomSet                                                               '��λ
            If mintEditState <> 0 Then
                txtGoto.Enabled = False
            Else
                txtGoto.Enabled = True
            End If
        Case conMenu_Manage_Transfer_Send, conMenu_Edit_UnArchive                               '����
            Control.Visible = (Me.TabCtlWindow.Selected.Index = 4)
    End Select

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim intLoop As Integer
    
    On Error GoTo errH
           
    Select Case Control.ID
        
        '''''''''''''''''''''''''''''''''''''''�ļ�''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_PrintSet                                                      '��ӡ����
             PrintSetup
            
        Case conMenu_File_Preview                                                       '����Ԥ��
            ReportPrint False
        
        Case conMenu_File_Print                                                         '�����ӡ

            If InStr(",7,8,", CStr(Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon)) = 0 Then
                If MsgBox("���ļ��鵥û�����!�Ƿ�ȷ��Ҫ��ӡ!", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            ReportPrint True
        
        Case conMenu_File_BatPrint                                                      '������ӡ
'            Call frmLisStationPrint.ShowEdit(Me, mlngDeptID)
            Call frmBatchAction.ShowMe(Me, 1, mlngMachineID, mstrPrivs, , , , mlngDeptID, mstrAuditingManID)
        Case conMenu_Edit_Save                                                          '����
            Call SaveDisposal(mFileS.����)
            
        Case conMenu_LIS_Cancel                                                         '����
            mintHandleState = 0
            Call SaveDisposal(mFileS.����)
            
        Case conMenu_File_RowPrint                                                      '�嵥��ӡ
            'Call zlRptPrint(1)
            Call frmLabReport.Show
        Case conMenu_Edit_Privacy                                                       '����˵�½
            Call AuditingRegister
            
        Case conMenu_File_Parameter                                                     '��������
            SetParameter
            
        Case conMenu_Tool_Monitor                                                       'ø��������
'            If frmMBSetup.ShowMe(Me) Then objLISComm.InitMBPara
            frmLabMB.ShowMe Me, mlngMachineID
            
        Case conMenu_File_Exit                                                          '�˳�
            Unload Me
        
        ''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Plan                                                        '����
            Call SampleDisposal(mActS.����)
        
        Case conMenu_Manage_Regist                                                      '�Ǽ�
            Call SampleDisposal(mActS.�Ǽ�)
        
        Case conMenu_Edit_NewParent                                                     '��������
            Call SampleDisposal(mActS.��������)
        
        Case conMenu_Manage_Logout                                                      '������������
            frmAddPatient.ShowMe Me, mlngMachineID, mlngDeptID
            
        Case conMenu_Manage_Receive                                                     '�����
            mintHandleState = 1
            Call SampleDisposal(mActS.�����)
        
        Case conMenu_Manage_Transfer                                                    '���º���
            Call SampleDisposal(mActS.���º���)
        
        Case conMenu_Edit_ModifyParent                                                  '�޸�������
            Call SampleDisposal(mActS.�޸�������)
        
        Case conMenu_Manage_Reset                                                       '�����޸�������
            Call SampleDisposal(mActS.�����޸�������)
            
'        Case conMenu_Edit_CardBound                                                     '�޸Ĳ�����Ϣ
'            Call SampleDisposal(mActS.�޸Ĳ�����Ϣ)
            
        Case conMenu_Tool_Apply                                                         '��������
            mbln�ֹ����ͱ��� = True
            Call SampleDisposal(mActS.��������)
            mbln�ֹ����ͱ��� = False
        Case conMenu_Tool_BathSend                                                      '��������
            
            Call SampleDisposal(mActS.�������͵�����)
        Case conMenu_LIS_TOQC                                                           '��Ϊ�ʿ�
            Call SampleDisposal(mActS.��Ϊ�ʿ�)
            
        Case comMenu_LIS_TodayQC                                                        '�����ʿ�Ȩ��
            frmQCTodayList.Show vbModal, Me
        
        Case comMenu_LIS_History                                                        '��ʷ�ʿ�Ȩ��
            frmQCHistory.Show vbModal, Me
            
        Case conMenu_LIS_LJAverage                                                      '��ֵ�ʿ�ͼ
            ShowLJAverage
        
        Case conMenu_Edit_QCRes                                                         '�鿴�����ʿ�
            Call frmLabMainLJ.ShowMe(mlngKey, Me, mlngMachineID)
        
        Case conMenu_Tool_Analyse                                                       '��Ϊ�ȶ�
            Call SampleDisposal(mActS.��Ϊ�Ա�)
            
        Case conMenu_Manage_ReportView                                                  '�鿴�ȶ�
            Call frmQCContrast.ShowMe(Me, mlngMachineID)
        
        Case conMenu_Edit_SendBack                                                      '״̬�ع�
            Call SampleDisposal(mActS.״̬�ع�)
        
        Case conMenu_Manage_ClearUp                                                     'ɾ�������걾
            Call SampleDisposal(mActS.ɾ�������걾)
        
        Case conMenu_Tool_MedRec                                                        'ָ��ɾ������
'            frmLisStationBatch.ShowEdit Me, mlngDeptID
'            Call RefreshData
            frmBatchAction.ShowMe Me, 3, mlngMachineID, , , , , mlngDeptID, mstrAuditingManID
        Case conMenu_Edit_DeleteParent                                                  'ȡ������
            Call SampleDisposal(mActS.��Ϊ����)
            
        Case conMenu_Edit_Insert                                                        '�ϲ��걾
            Call SampleDisposal(mActS.�ϲ��걾����)
        
        Case conMenu_Edit_Surplus                                                       '�������ϲ�
            frmLabBloodSugar.ShowMe Me, mlngMachineID, mlngKey
            Call RefreshData
            
        Case conMenu_Manage_Refuse                                                      '����
            Call SampleDisposal(mActS.����)
        Case conMenu_Statistics_PositiveResults                                         '���Խ������
            ShowPositiveResults (0)
        
        Case conMenu_Statistics_Feedback                                                '���������ѯ
            ShowPositiveResults (1)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Report                                                      '������д
            If Me.TabCtlWindow.Selected.Index = 2 Then
                mintHandleState = 3
                Call ReportDisposal(mActR.��д��������)
            Else
                mintHandleState = 2
                Call ReportDisposal(mActR.��д����)
            End If
            
        Case conMenu_Edit_Adjust                                                        '��������
            Call ReportDisposal(mActR.������������)
        
        Case conMenu_Edit_Dilute                                                        '�걾ϡ��
            frmDiluteSample.ShowMe Me, mlngKey
            mfrmWrite.zlRefresh mlngKey
        
        Case conMenu_Edit_Audit                                                         '�������
            Call ReportDisposal(mActR.��˱���)
            
        Case conMenu_LIS_SendReport                                                     '���淢��
            Call ReportDisposal(mActR.���ͱ���)
        
        Case conMenu_Manage_Audit                                                       '�������
            Call ReportDisposal(mActR.������˱���)
        
        Case conMenu_Edit_ClearUp                                                       'ȡ�����
            Call ReportDisposal(mActR.���ȡ��)
        
        Case conMenu_Edit_Seat_Set                                                      '���������
            Call ReportDisposal(mActR.���������)
        
        Case conMenu_Manage_Redo                                                        '�������
            Call ReportDisposal(mActR.�������)
        
        Case conMenu_Manage_Undone                                                      'ȡ������
            Call ReportDisposal(mActR.ȡ������)
       
        Case conMenu_Manage_Transfer_Force                                              '���˱����ѯ
            If Me.rptList.FocusedRow Is Nothing Then
                frmLabMainFindRePort.ShowMe -1, Me, mstrPrivs
            Else
                frmLabMainFindRePort.ShowMe Val(Me.rptList.FocusedRow.Record(mCol.����ID).Value), Me, mstrPrivs
            End If
'            Me.SetFocus: Me.TabList.SetFocus: Me.rptList.SetFocus
'            Me.dkpMain.FindPane(Dkp_ID_List).Select
                    
        Case conMenu_File_ImportFromXML                                                 '�������ݲɼ�
            frmLabAnalyseData.ShowMe Me, mlngMachineID
                            
        Case conMenu_LIS_SignVerify                                                     '��֤ǩ��
            Call ReportDisposal(mActR.��֤ǩ��)
                            
        Case conMenu_Edit_Import                                                        '�Զ�����
            GetSaveSetup 1
        
        Case conMenu_Edit_ApplyTo                                                       '��������
            GetSaveSetup 2
            
        Case conMenu_LIS_SaveSample                                                     '�걾����(��ţ�
            frmlabONSample.Show vbModal, Me
            
        Case conMenu_LIS_DropSample                                                     '�걾����
            frmlabDropSample.Show vbModal, Me
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''''''''''''''
'        Case conMenu_Edit_Price                                                         '��������
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_Append
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Manage_ThingAdd                                                    '���ӷѻ���
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_NewItem * 10# + 1
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Edit_ModifyParent                                                  '���ӷѼ���
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_NewItem * 10# + 2
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Edit_NewItem                                                       '��Ѽ�¼
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_NewItem * 10# + 3
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Manage_ThingModi                                                   '�޸ĸ��ӷ�
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_Modify
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Manage_ThingDel                                                    'ɾ�����ӷ�
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_Delete
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''''''''''''''�鿴'''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_ToolBar_Button                                                '��׼��ť
            Control.Checked = Not Control.Checked
            Me.cbrthis(2).Visible = Control.Checked
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                                  '�ı���ǩ
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Size                                                  '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_StatusBar                                                     '״̬��
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Control.Checked
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_Forward                                                       'ǰһ��
            BackOrNextPatient 2
        
        Case conMenu_View_Backward                                                      '��һ��
            BackOrNextPatient 1
        
        Case conMenu_Tool_Reference_1                                                   'PAGEUP
            ShortWork mSWork.Key_PageUP
            
        Case conMenu_Tool_Reference_2                                                   'PAGEDOWN
            ShortWork mSWork.Key_PageDown
        
        Case conMenu_Tool_MeetFinish                                                    'HOME
            ShortWork mSWork.Key_Home
        
        Case conMenu_Tool_MeetCancel                                                    'End
            ShortWork mSWork.Key_End
            
        Case conMenu_View_Notify                                                        '��δ�շ�
            Control.Checked = Not Control.Checked
        
        Case conMenu_LIS_HideList                                                       '�����б�
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_List
        
        Case conMenu_Manage_LeaveMedi                                                   '���ؼ���ͼ��
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_Image
            
        Case comMenu_LIS_ShowListHead                                                   'ѡ���б�
            If TabList.Selected.Index = 0 Then
                ShowHideListHead Me.rptList.Columns, frmPublicFieldChooser.ShowMe(Me, Me.rptList.Columns)
            Else
                ShowHideListHead Me.rptList1.Columns, frmPublicFieldChooser.ShowMe(Me, Me.rptList1.Columns)
            End If
            
            
        Case conMenu_Manage_ReGet                                                       '��ʾ������
            Control.Checked = Not Control.Checked
            Me.TabList.Item(1).Visible = Control.Checked
        
        Case conMenu_LIS_PatientInfo                                                    '������Ϣ
            If Not Me.rptList.FocusedRow Is Nothing Then
                frmDegreeCard.ShowInfo Me, Val(Me.rptList.FocusedRow.Record(mCol.����ID).Value)
            End If
        
        Case conMenu_View_Find                                                          '��λ
            If Me.txtGoto.Enabled Then Me.txtGoto.SetFocus
            
        Case conMenu_View_Filter                                                        '���ٹ���
            Call QUFilter
        
        Case conMenu_View_FindNext                                                      '�������μ���
            Call QuickFindPatient
        
        Case conMenu_Manage_ReportEdit                                                  '�鿴��˼�¼
            frmLabAuditingCourse.ShowMe Me, mlngKey
        
        Case conMenu_View_Refresh                                                       'ˢ��
            '��չ�������
            Me.rptList.Tag = ""
            If Me.TabList.Item(0).Selected = True Then
                Call RefreshData
            Else
                Call RefreshData1
            End If
            
        
        ''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Help_Help                                                          '��������
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_Web                                                           'WEB�ϵ�
            Call zlHomePage(hWnd)
        
        Case conMenu_Help_Web_Home                                                      '��ҳ
            Call zlHomePage(Me.hWnd)
        
        Case conMenu_Help_Web_Mail                                                      '���ͷ���
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_Help_About                                                         '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
        ''''''''''''''''''''''''''''''''''�����б�������б�''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Report_DrugQuery                                                   '����ѡ��
            
        Case conMenu_Report_Reports                                                     'ѡ������
        
        Case conMenu_View_Owe                                                           'ѡ�����
            Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_NoPrint, True, True)
            rptList.Tag = ""
            With cbrMenuBar.CommandBar
                For intLoop = 1 To .Controls.Count
                    .Controls(intLoop).Checked = (.Controls(intLoop).Caption = Control.Caption)
                Next
            End With
            mstrMachineGroup = Control.Caption
            Call cboDept_Click
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_IDkind_Change
             mfrmRequest.IdKindChange
        Case conMenu_File_PrintBedCard                                                '�ش�����
            '21436 ����
            Call PrintBarcord
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "�������=" & mlngDeptID, "��������ID=" & mlngMachineID, "�걾ID=" & mlngKey)
            Else
                Select Case Me.TabCtlWindow.Selected.Index
                    Case 4  '����
                        mclsExpenses.zlExecuteCommandBars Control
                    Case 6  '����
                        mclsOutAdvices.zlExecuteCommandBars Control
                    Case 7  'סԺ
                        mclsInAdvices.zlExecuteCommandBars Control
                    Case Else
                        '------------------------------------------------------
                        '������
                        Dim lngCount As Long
                        If Not clsPluginLoader Is Nothing Then
                            If clsPluginLoader.PluginCount > 0 Then
                                lngCount = Control.ID - conMenu_PlugPopup * 1000# - 100
                                If Mid(Control.ID, 1, 1) = conMenu_PlugPopup And lngCount >= 0 Then
                                    If mobjPlugin(lngCount) Is Nothing Then
                                        MsgBox "���ܵ��ò��!", vbExclamation
                                        Exit Sub
                                    End If
                                    ' ������ṩ����״̬
                                    mobjPlugin(lngCount).InitQuery Me
                                    'ִ�в�� ��Ϊ��ģ̬������ʾ
                                    mobjPlugin(lngCount).DoAction Query_ShowModeless
                                End If
                            End If
                        End If
                        '------------------------------------------------------
                End Select
            End If
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowPositiveResults(ByVal intType As Integer) As Boolean
    '����       ���ô�Ⱦ����������,�ʹ�Ⱦ�����洰��
    'intType    0-���Խ������ 1-���������ѯ
    Dim lngKey As Long      '�걾ID
    Dim strSQL As String
    Dim rsPatient As Recordset  '������Ϣ
    Dim rsJYK As Recordset '�����
    Dim rsBuMen As Recordset '������Ҳ���ID
    Dim intSendOk As Integer
    Dim objPublicAdvic As Object
    
    On Error GoTo ErrHand
    
    Set objPublicAdvic = CreateObject("zlPublicAdvice.clsPublicAdvice")
    Call objPublicAdvic.InitDisease(gcnOracle, 100)
    
    If intType = 0 Then
        With Me.rptList
            If .FocusedRow Is Nothing Then
                MsgBox "��ѡ��һ���걾", vbInformation
                Exit Function
            End If
            lngKey = Val(rptList.FocusedRow.Record(mCol.ID).Value)
        End With
        If lngKey = 0 Then
            MsgBox "��ѡ��һ���걾", vbInformation
        Else
            strSQL = "Select Distinct a.Id, a.����id, a.��ҳid, a.�걾���� �걾����,a.�������id �ͼ����ID, " & _
                     "b.�걾�ͳ�ʱ�� �ͼ�ʱ��, b.�ͼ��� �ͼ�ҽ��, a.�������id, a.�Һŵ�, " & _
                     " a.����ʱ��, a.סԺ��, a.ִ�п���id  �Ǽǿ���id,a.ҽ��id ҽ��ID From ����걾��¼ A, ����ҽ������ B " & _
                     "Where a.ҽ��id = b.ҽ��id and a.Id = [1] and a.ҽ��id is not null"
            Set rsPatient = zlDatabase.OpenSQLRecord(strSQL, "��ѯHIS����ID", lngKey)
            
'            strSQL = "select ID �ͼ����ID from ���ű� where ����=[1]"
'            Set rsBuMen = ComOpenSQL(Sel_His_DB, strSQL, "��ѯHIS����ID", IIf(IsNull(rsPatient("�������")), "", rsPatient("�������")))
            
            If rsPatient.RecordCount < 1 Or (IsNull(rsPatient("��ҳID")) And IsNull(rsPatient("�Һŵ�"))) Then
                MsgBox "û�в��ҵ�������ص�ҽ����Ϣ", vbInformation
            Else
                If IsNull(rsPatient("�ͼ�ʱ��")) Then
                    If IsNull(rsPatient("����ʱ��")) Then
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("����ID")), 0, rsPatient("����ID")), _
                                                       IIf(IsNull(rsPatient("��ҳID")), 0, rsPatient("��ҳID")), _
                                                       IIf(IsNull(rsPatient("�Һŵ�")), "", rsPatient("�Һŵ�")), _
                                                       IIf(IsNull(rsPatient("ҽ��ID")), 0, rsPatient("ҽ��ID")), _
                                                       IIf(IsNull(rsPatient("�Ǽǿ���ID")), 0, rsPatient("�Ǽǿ���ID")), , _
                                                       IIf(IsNull(rsPatient("�ͼ����ID")), 0, rsPatient("�ͼ����ID")), _
                                                       IIf(IsNull(rsPatient("�ͼ�ҽ��")), "", rsPatient("�ͼ�ҽ��")), _
                                                       IIf(IsNull(rsPatient("�걾����")), "", rsPatient("�걾����")))
                    Else
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("����ID")), 0, rsPatient("����ID")), _
                                                       IIf(IsNull(rsPatient("��ҳID")), 0, rsPatient("��ҳID")), _
                                                       IIf(IsNull(rsPatient("�Һŵ�")), "", rsPatient("�Һŵ�")), _
                                                       IIf(IsNull(rsPatient("ҽ��ID")), 0, rsPatient("ҽ��ID")), _
                                                       IIf(IsNull(rsPatient("�Ǽǿ���ID")), 0, rsPatient("�Ǽǿ���ID")), , _
                                                       IIf(IsNull(rsPatient("�ͼ����ID")), 0, rsPatient("�ͼ����ID")), _
                                                       IIf(IsNull(rsPatient("�ͼ�ҽ��")), "", rsPatient("�ͼ�ҽ��")), _
                                                       IIf(IsNull(rsPatient("�걾����")), "", rsPatient("�걾����")), , _
                                                       CDate(rsPatient("����ʱ��")))
                    End If
                Else
                    If IsNull(rsPatient("����ʱ��")) Then
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("����ID")), 0, rsPatient("����ID")), _
                                                       IIf(IsNull(rsPatient("��ҳID")), 0, rsPatient("��ҳID")), _
                                                       IIf(IsNull(rsPatient("�Һŵ�")), "", rsPatient("�Һŵ�")), _
                                                       IIf(IsNull(rsPatient("ҽ��ID")), 0, rsPatient("ҽ��ID")), _
                                                       IIf(IsNull(rsPatient("�Ǽǿ���ID")), 0, rsPatient("�Ǽǿ���ID")), _
                                                       CDate(rsPatient("�ͼ�ʱ��")), _
                                                       IIf(IsNull(rsPatient("�ͼ����ID")), 0, rsPatient("�ͼ����ID")), _
                                                       IIf(IsNull(rsPatient("�ͼ�ҽ��")), "", rsPatient("�ͼ�ҽ��")), _
                                                       IIf(IsNull(rsPatient("�걾����")), "", rsPatient("�걾����")))
                    Else
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("����ID")), 0, rsPatient("����ID")), _
                                                       IIf(IsNull(rsPatient("��ҳID")), 0, rsPatient("��ҳID")), _
                                                       IIf(IsNull(rsPatient("�Һŵ�")), "", rsPatient("�Һŵ�")), _
                                                       IIf(IsNull(rsPatient("ҽ��ID")), 0, rsPatient("ҽ��ID")), _
                                                       IIf(IsNull(rsPatient("�Ǽǿ���ID")), 0, rsPatient("�Ǽǿ���ID")), _
                                                       CDate(rsPatient("�ͼ�ʱ��")), _
                                                       IIf(IsNull(rsPatient("�ͼ����ID")), 0, rsPatient("�ͼ����ID")), _
                                                       IIf(IsNull(rsPatient("�ͼ�ҽ��")), "", rsPatient("�ͼ�ҽ��")), _
                                                       IIf(IsNull(rsPatient("�걾����")), "", rsPatient("�걾����")), , _
                                                       CDate(rsPatient("����ʱ��")))
                    End If
                End If
            End If
        End If
    ElseIf intType = 1 Then
        If mlngDeptID <> 0 Then
            Call objPublicAdvic.ShowDisQuery(Val(mlngDeptID))
        End If
    End If
    Set objPublicAdvic = Nothing
    Exit Function
ErrHand:
    MsgBox ("�������Խ�������������!" & vbCrLf & "������Ϣ:" & Err.Description & "(" & Err.Number & ")"), vbQuestion, "������:ShowPositiveResults"
End Function

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    On Error Resume Next
    Select Case Me.TabCtlWindow.Selected.Index
        Case 4
            Select Case CommandBar.Parent.ID
            Case conMenu_Edit_NewItem '����
                With CommandBar.Controls
                    .DeleteAll
                    '��1λ,Ϊ��ʹ�ÿ�ݼ�
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "�շѵ���(&1)")
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 2, "���ʵ���(&2)")
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 3, "��ķ���(&3)")
                    With cbrthis.KeyBindings
                        .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem * 10# + 1
                        .Add FCONTROL, vbKeyB, conMenu_Edit_NewItem * 10# + 2
                    End With
                End With
            End Select
'            Call mclsExpenses.zlPopupCommandBars(CommandBar)
        Case 6
            Select Case CommandBar.Parent.ID
            Case conMenu_Edit_Compend '����
                With CommandBar.Controls
                    If .Count = 0 Then
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "���ı���(&W)"
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "��ӡ����(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)"
                    End If
                End With
            End Select
'            Call mclsOutAdvices.zlPopupCommandBars(CommandBar)
        Case 7
            Select Case CommandBar.Parent.ID
            Case conMenu_Edit_Compend '����
                With CommandBar.Controls
                    If .Count = 0 Then
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "���ı���(&W)"
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "��ӡ����(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)"
                    End If
                End With
            End Select
'            Call mclsInAdvices.zlPopupCommandBars(CommandBar)
    End Select
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRowCount As Long                                   '��ǰ��ʾ������
    Dim intSampleType  As Integer                             '�걾���� = 3 �ʿ�   = 4 �ȶ� = -1 ��ͨ
    Dim strSource As String                                   '������Դ
    Dim strExeState As String                                 'ִ��״̬ =�Ѽ���/������
    Dim blMicrobe As Integer                                  '΢���� =true ��΢����
    Dim intReportCount As Integer                             '��������
    Dim strPatienName As String                               '����
    Dim lngMachineID As Long                                  '����ID
    Dim blWaiteDispose As Boolean                             '�Ƿ��ڵȴ������б�
    Dim lngExecDept As Long                                   'ִ�п���ID
    Dim blnIF As Boolean                                      '�ж�����
    Dim lngAdvice As Long                                     'ҽ��ID
    Dim str������ As String                                   '������
    Dim blnExec As Boolean                                    'ѡ�����п���ʱ���ܽ��в����д����Ȳ���
    Dim lngSendReport As Long                                 '���ͱ��� <>0 �б���
    Dim blnɾ�������걾 As Boolean                            'ɾ�������걾
    
    On Error GoTo errH
        
    'CSBmk_CS <If Me.Visible = False>
    If Me.Visible = False Then Exit Sub
    
    '���뵱ǰ�е���Ϣ(�����ж��Ƿ�Disabled)
    With Me.rptList
        If Not .FocusedRow Is Nothing Then
            lngRowCount = .Rows.Count
            If .Rows.Count = 0 Then Exit Sub
            intSampleType = .FocusedRow.Record(mCol.�걾����).Icon
            lngExecDept = Val(.FocusedRow.Record(mCol.ִ�п���ID).Value)
            intReportCount = Val(.FocusedRow.Record(mCol.�������).Value)
            lngMachineID = Val(.FocusedRow.Record(mCol.����id).Value)
            lngAdvice = Val(.FocusedRow.Record(mCol.ҽ��id).Value)
            blMicrobe = IIf(Val(.FocusedRow.Record(mCol.΢����걾).Value) = 1, True, False)
            lngSendReport = Val(.FocusedRow.Record(mCol.���淢��).Value)
            str������ = .FocusedRow.Record(mCol.������).Value
            If .FocusedRow.Record(mCol.ִ��״̬).Value = "�Ѽ���" Or .FocusedRow.Record(mCol.ִ��״̬).Value = "�Ѵ�ӡ" Then
                strExeState = "�Ѽ���"
            ElseIf .FocusedRow.Record(mCol.ִ��״̬).Value = "����" Then
                strExeState = "����"
            Else
                strExeState = "������"
            End If
            strPatienName = .FocusedRow.Record(mCol.����).Value
            strSource = .FocusedRow.Record(mCol.�������).Value
            If strSource = "" Or strSource = "����" Then
                If strPatienName = "" Then
                    strSource = "����"
                Else
                    strSource = "Ժ��"
                End If
            End If
        End If
    End With
    
    '�����ж��ǲ����в�����ǰ���ұ��浥Ȩ��
    If Me.TabList.Item(0).Selected = True Then
        If InStr(mstrPrivs, "���п���") = 0 Then
            If mlngMachineID > 0 Or lngMachineID > 0 Then
                blnIF = InStr(";" & Replace(mstrMachineALL, ",", ";") & ";", ";" & IIf(lngMachineID = 0, mlngMachineID, lngMachineID) & ";")
            Else
                blnIF = True
            End If
        Else
            blnIF = True
        End If
    Else
        blnIF = True
    End If
    
    If InStr(";" & mstrPrivs & ";", ";ɾ�������걾;") > 0 Then blnɾ�������걾 = True
    blnExec = InStr(mstrMachines, ";" & IIf(lngMachineID <= 0, mlngMachineID, lngMachineID) & ";")
    '�ֹ���Ŀ�����п���
    If blnExec = False Then
        If mlngMachineID = -1 Or lngMachineID = 0 Then blnExec = True
    End If
'    blnIF = False
    blWaiteDispose = Me.TabList.Item(1).Selected
    
    '�����б��TAB�ؼ��ڱ༭ʱ���ܸı�
    If blWaiteDispose = True Then
        If mintEditState = 4 Or mintEditState Then
            Me.PicList.Enabled = False
        Else
            Me.PicList.Enabled = True
        End If
    Else
        'ֻ����������ʱ������ѡ���б�
        Me.PicList.Enabled = (Me.rptList.Tag <> "Continue")
    End If
    If mintEditState = 5 Or mintEditState = 1 Or mintEditState = 2 Or mintEditState = 4 Then
        Me.TabCtlWindow.Item(2).Enabled = False
        Me.TabCtlWindow.Item(3).Enabled = False
    Else
        Me.TabCtlWindow.Item(2).Enabled = True
        Me.TabCtlWindow.Item(3).Enabled = True
    End If
    
    Select Case Control.ID
'        '''''''''''''''''''''''''''''''''''''''''''''''''�ļ�'''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_PrintSet                                                                      '��ӡ����
            Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0)
        Case conMenu_File_RowPrint                                                                      '�嵥��ӡ
            Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0)
        Case conMenu_File_BatPrint                                                                      '������ӡ
            If InStr(1, mstrPrivs, "������ӡ") <= 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0)
            End If
        Case conMenu_File_Preview, conMenu_File_Print                                                   '����Ԥ��,�����ӡ
            If InStr(1, mstrPrivs, "�����ӡ") <= 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If InStr(1, mstrPrivs, "δ��˴�ӡ") > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" Then
                    Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0 _
                    And IIf(strSource = "����", InStr(1, mstrPrivs, "������ӡ") > 0, True))
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Edit_Save, conMenu_LIS_Cancel, conMenu_IDkind_Change                               '����,����,�л�
            Control.Enabled = (mintEditState <> 0 And blnIF = True And blnExec = True)
        Case conMenu_Manage_Refuse                                                                      '����
'            Control.Enabled = (mintEditState = 1 Or mintEditState = 0 And strExeState = "������" _
'            And strSource <> "����" And intSampleType = -1)
            If Me.rptList1.FocusedRow Is Nothing Or blnIF = False Then
                Control.Enabled = False
            Else
                Control.Enabled = (mintEditState = 0 And Me.rptList1.FocusedRow.Record(mRCol.ִ��״̬).Value <> 2)
            End If
'            Control.Enabled = (Not Me.rptList1.FocusedRow Is Nothing And mintEditState = 0)
        Case conMenu_File_Parameter                                                                     '��������
            If InStr(1, mstrPrivs, "��������") <= 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0)
            End If
        Case conMenu_Tool_Monitor                                                                       'ø��������
            Control.Enabled = (mintEditState = 0)
'        ''''''''''''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Plan                                                                        '����
            If InStr(1, mstrPrivs, "���ձ걾") <= 0 Or blnIF = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If blWaiteDispose = False Then
                    Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
                Else
                    Control.Enabled = (Me.rptList1.Rows.Count > 0 And Not Me.rptList1.FocusedRow Is Nothing And (mintEditState = 0) _
                                       And mlngDeptID > 0)
                End If
            End If
        
        Case conMenu_Manage_Regist                                                                          '�Ǽ�
            If InStr(1, mstrPrivs, "ֱ������") <= 0 Or blWaiteDispose = True Or blnIF = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
            End If
        
        Case conMenu_Manage_Logout                                                                          '�����������ӵǼ�
            If InStr(1, mstrPrivs, "ֱ������") <= 0 Or blWaiteDispose = True Or blnIF = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
            End If
        Case conMenu_Edit_NewParent                                                                         '������������
            If InStr(1, mstrPrivs, "ֱ������") <= 0 Or blWaiteDispose = True Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
            End If
                    
        Case conMenu_Manage_Receive                                                                         '�����
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" _
               Or intSampleType <> -1 And strSource <> "����" Or InStr(1, mstrPrivs, "��������") <= 0 _
               Or intSampleType = 3 Or blWaiteDispose = True Or blnIF = False Or mlngDeptID = 0 Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
                    
        Case conMenu_LIS_TOQC                                                                               '��Ϊ�ʿ�
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" _
               Or intSampleType <> -1 Or strSource <> "����" Or InStr(1, mstrPrivs, "��������") <= 0 _
               Or intSampleType = 3 Or blWaiteDispose = True Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Manage_Transfer                                                                        '���º���
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" _
               Or intSampleType <> -1 Or blWaiteDispose = True Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Tool_Apply                                                                             '��������
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_ModifyParent                                                      '�޸�����
            If InStr(1, mstrPrivs, "�޸ı걾��") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or blWaiteDispose = True Or strExeState = "����" Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
'        Case conMenu_Edit_CardBound                                                         '�޸Ĳ�����Ϣ
'            If InStr(1, mstrPrivs, "�޸Ĳ�����Ϣ") <= 0 Or blnIF = False Or blnExec = False Then
'                Control.Visible = False: Control.Enabled = False
'            Else
'                Control.Visible = True
'                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or blWaiteDispose = True Or strExeState = "����" Then
'                    Control.Enabled = False
'                Else
'                    Control.Enabled = True
'                End If
'            End If
        Case conMenu_Tool_MedRec                                                            'ָ��ɾ������
            Control.Enabled = (mintEditState = 0 And blnIF = True And blnExec = True And blnɾ�������걾 = True)
            
        Case conMenu_Manage_Reset                                                           '�����޸�������
            Control.Enabled = (mintEditState = 0 And blnIF = True Or blnExec = True)
        Case conMenu_Edit_QCRes                                                             '�鿴�����ʿ�
            Control.Enabled = (intSampleType = 3 And mintEditState = 0)
            
        Case comMenu_LIS_TodayQC                                                            '�����ʿ�
            Control.Visible = (mTodayQCPrivs <> "")
            
        Case comMenu_LIS_History                                                            '��ʷ�ʿ�
            Control.Visible = (mHistoryPrivs <> "")
            
        Case conMenu_Tool_Analyse                                                           '��Ϊ�ȶ�
            If lngRowCount = 0 Or mintEditState > 0 Or intSampleType <> -1 Or blWaiteDispose = True Or TabList.Item(0).Visible = False _
               Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                '11198 ��Ϊ�ȶ� �˵����ֹ��걾����Ӧ��Ϊ��Ч
                Control.Enabled = InStr(Me.rptList.FocusedRow.Record(mCol.�걾��).Caption, "-") <= 0
            End If
        Case conMenu_Edit_DeleteParent                                                      '��Ϊ����
            If lngRowCount = 0 Or mintEditState > 0 Or intSampleType <> -1 _
               Or strPatienName = "" Or strExeState = "�Ѽ���" Or blWaiteDispose = True _
               Or blnIF = False Or blnExec = False Or strExeState = "����" Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_SendBack                                                          '״̬�ع�
            If InStr(1, mstrPrivs, "���ճ���") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
'                If lngRowCount = 0 Or mintEditState > 0 Or (strSource = "����" And intSampleType < 3) Or _
'                    (strExeState = "�Ѽ���" And intSampleType <> 4) Then
'                    Control.Enabled = False
'                Else
'                    Control.Enabled = True
'                End If
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or blWaiteDispose = True Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_Manage_ClearUp                                                            'ɾ������
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" _
                Or strSource = "סԺ" Or strSource = "����" And strSource <> "����" _
                Or intSampleType <> -1 Or InStr(1, mstrPrivs, "��������") <= 0 Or blWaiteDispose = True _
                Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If

        '''''''''''''''''''''''''''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Report                                                          '������д
            If InStr(1, mstrPrivs, "������д") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or blWaiteDispose = True Or strExeState = "����" Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_Edit_Adjust, conMenu_Edit_Dilute                                          '��������,ϡ�ͱ���
            If lngRowCount = 0 Or mintEditState > 0 Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        
        Case conMenu_Manage_Audit                                                              '�������
            If InStr(1, mstrPrivs, "��˱걾") <= 0 Or mintEditState > 0 Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_Audit                                                             '�������
            If InStr(1, mstrPrivs, "��˱걾") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" _
                    Or strSource = "����" Or intSampleType = 3 Or blWaiteDispose = True Or _
                    (mSendReport = 1 And str������ = "") Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_LIS_SendReport                                                         '���󱨸�
            If strPatienName = "" Or mSendReport = 0 Or mintEditState > 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (str������ = "" And strPatienName <> "" And strExeState = "������")
            End If
        Case conMenu_Edit_ClearUp                                                           'ȡ�����
            If InStr(1, mstrPrivs, "���ȡ��") <= 0 And InStr(1, mstrPrivs, "24Сʱ���ȡ��") <= 0 _
                Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "������" Or blWaiteDispose = True Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_Manage_Redo                                                            '�������
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" Or strExeState = "����" _
                Or blMicrobe = True Or strSource = "����" Or intSampleType = 3 _
                Or InStr(1, mstrPrivs, "��������") <= 0 Or blWaiteDispose = True Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Manage_Undone                                       'ȡ�����

            If lngRowCount = 0 Or mintEditState > 0 Or intReportCount = 0 Or blWaiteDispose = True _
                        Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_Import, conMenu_Edit_ApplyTo                                      '�Զ�����,��������
            Control.Enabled = mintEditState = 0
        Case conMenu_Edit_Insert                                                            '�ϲ�
            If Me.TabCtlWindow.Selected.Index = 5 Then
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And blnIF = True And blnExec = True)
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_Surplus                                                           '�������ϲ�
            Control.Enabled = (mintEditState = 0 And blnIF = True And strExeState <> "�Ѽ���" And blnExec = True And strExeState <> "����")
        Case conMenu_LIS_SignVerify                                                         '��֤ǩ��
            If gobjESign Is Nothing Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (strExeState = "�Ѽ���")
            End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Case conMenu_Edit_Price                                                             '��������
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" _
'               Or strSource = "����" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                    Or InStr(1, mstrPrivs, "��������") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
'        Case conMenu_Manage_ThingAdd                                                        '���ӷѻ���
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" _
'               Or strSource = "����" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                Or InStr(1, mstrPrivs, "���Ѵ���") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
'        Case conMenu_Edit_ModifyParent, conMenu_Edit_NewItem                                '���ӷѼ���,��Ѽ�¼
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" _
'               Or strSource = "����" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                    Or InStr(1, mstrPrivs, "���Ѵ���") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
'        Case conMenu_Manage_ThingModi, conMenu_Manage_ThingDel                              '�޸ĸ��ӷ�,ɾ�����ӷ�
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "�Ѽ���" _
'               Or strSource = "����" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                Or InStr(1, mstrPrivs, "���Ѵ���") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''�鿴''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_Backward                                                         'ǰһ��
            If mintEditState = 4 Or mintEditState = 5 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.FocusedRow.Index = 0 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_View_Forward                                                          '��һ��
            
            If mintEditState = 4 Or mintEditState = 5 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.Rows.Count - 1 = Me.rptList.FocusedRow.Index Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_Tool_Reference_1, conMenu_Tool_Reference_2, conMenu_Tool_MeetFinish, conMenu_Tool_MeetCancel
            Control.Visible = False
        Case conMenu_View_Filter                                                           '����
            If InStr(1, mstrPrivs, "�ۺϲ�ѯ") <= 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If mintEditState > 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_View_Refresh                                                           'ˢ��
            If mintEditState > 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_LIS_PatientInfo                                                        '������Ϣ
            If lngRowCount = 0 Or mintEditState > 0 Or strSource = "����" _
                Or intSampleType = 3 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_View_FindNext                                                          '�������μ���
            If Not Me.rptList.FocusedRow Is Nothing Then
                Control.Enabled = mintEditState = 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Bespeak                                                         'ֻ��ʾ�շ�
            Control.Checked = Control.Checked
        Case conMenu_View_ToolBar_Button                                                    '��ʾ������
            Control.Checked = Me.cbrthis(2).Visible
        Case conMenu_View_ToolBar_Text                                                      '�Ƿ���ʾ����
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size                                                      '�Ƿ���ʾ��ͼ��
            Control.Checked = Me.cbrthis.Options.LargeIcons
        Case conMenu_View_StatusBar                                                         '�Ƿ���ʾ״̬��
            Control.Checked = Me.stbThis.Visible
        Case conMenu_Manage_ReGet                                                           '��ʾ������
            If mintEditState > 0 Then
                Control.Enabled = False
            Else
                Control.Checked = Control.Checked
                Me.TabList.Item(1).Visible = Control.Checked
                If Control.Checked = False Then Me.TabList.Item(0).Selected = True
            End If
'        case
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Report_DrugQuery, conMenu_Report_Reports, conMenu_Report_WorkLog       '����,����
            If mintEditState <> 0 Then
                Me.cboDept.Enabled = False
                Me.cboMachine.Enabled = False
                Me.cboUnionItem.Enabled = False
            Else
                Me.cboDept.Enabled = True
                Me.cboMachine.Enabled = True
                Me.cboUnionItem.Enabled = True
            End If
        Case Else
            On Error Resume Next
            Select Case Me.TabCtlWindow.Selected.Index
                Case 4 '����
                    mclsExpenses.zlUpdateCommandBars Control
                Case 6 '����ҽ��
                    mclsOutAdvices.zlUpdateCommandBars Control
                Case 7 'סԺҽ��
                    mclsInAdvices.zlUpdateCommandBars Control
            End Select
    End Select
    
    
    
    
    On Error Resume Next
    '�����ǰѡ��Ĵ��岻��Ӧ�ûص�����Ĵ����ָ��
    Select Case gintSelectFocus
        Case 1              '�б�
'            Me.dkpMain.FindPane(Dkp_ID_List).Select
'            Me.TabList.SetFocus
'            If Me.TabList.Selected.Index = 0 Then
'                Me.rptList.SetFocus
'                '���ڽ��㲻��ȷ��Ҫ���涯��������
'                SendKeys "{UP}"
''                SendKeys "{Down}"
'            Else
'                Me.rptList1.SetFocus
'            End If
        Case 2              '������Ϣ
            Me.dkpMain.FindPane(Dkp_ID_Request).Select
            mfrmRequest.Show
        Case 3              '������д
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            TabCtlWindow.SetFocus: mfrmWrite.Vsf.SetFocus
        Case 4
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            TabCtlWindow.SetFocus: mfrmWrite2.Vsf.SetFocus
        Case 5
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            TabCtlWindow.SetFocus: mfrmWrite2.vsfDetail.SetFocus
    End Select
    gintSelectFocus = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub chkSoure_Click(Index As Integer)
    Dim astrItem() As String
    Dim strTypeName As String
    If Me.Visible = False Then Exit Sub
    If Me.TabList.Selected.Index = 0 Then
        astrItem = Split(con_������ɸѡ_������, ";")
        strTypeName = "������"
    Else
        astrItem = Split(con_������ɸѡ_������, ";")
        strTypeName = "������"
    End If
    If strTypeName = "������" Then
        If Index = 5 Then
            zlDatabase.SetPara strTypeName & "_" & astrItem(Index - 3), chkSoure(Index).Value, 100, 1208
        Else
            zlDatabase.SetPara strTypeName & "_" & astrItem(Index), chkSoure(Index).Value, 100, 1208
        End If
    Else
        zlDatabase.SetPara strTypeName & "_" & astrItem(Index), chkSoure(Index).Value, 100, 1208
    End If
    If Me.TabList.Selected.Index = 0 Then
        Call GetVerifying
    Else
        Call GetWaitVerify
    End If
    '���˽����б�
    RptListFilter
End Sub

Private Sub chkSoure_GotFocus(Index As Integer)
    On Error Resume Next
    If Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
    Else
'        Me.rptList1.SetFocus
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Me.Visible = False Then Exit Sub
    Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Me.Visible = False Then Exit Sub
    Select Case Item.ID
    Case Dkp_ID_List
        Item.Handle = Me.PicList.hWnd
    Case Dkp_ID_Locate
        Item.Handle = Me.PicInfo.hWnd
    Case Dkp_ID_Request
        Item.Handle = mfrmRequest.hWnd
    Case Dkp_ID_Append
        Item.Handle = Me.picTab.hWnd
    Case Dkp_ID_Image
        Item.Handle = Me.PicImage.hWnd
    End Select
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    If Me.Visible = False Then Exit Sub
    Me.cbrthis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpMain_Resize()
    If Me.Visible = False Then Exit Sub
    Me.cbrthis.RecalcLayout
    ImageTypeSet Me.VScroll.Max
End Sub

Private Sub dtpDate_Change()
    If Me.TabList.Item(1).Selected = True Then
        zlDatabase.SetPara "�����շ�Χ", cboʱ��.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData1
    Else
        zlDatabase.SetPara "�걾��Χ", cboʱ��.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData
    End If
    
End Sub

Private Sub dtpDateEnd_Change()
    If Me.TabList.Item(1).Selected = True Then
        zlDatabase.SetPara "�����շ�Χ", cboʱ��.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData1
    Else
        zlDatabase.SetPara "�걾��Χ", cboʱ��.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData
    End If
End Sub

Private Sub Form_Activate()
    '�յ��������ݺ󣬴������¼�
    On Error Resume Next
'    If objLISComm.DataReceived And Me.Tag <> "Refresh" And Not blnChecking And blnAutoRefresh And mintEditState = 0 Then
'        Me.Tag = "Refresh"
'        Call RefreshData
'        Me.Tag = ""
'    End If
    
    If mintLoadShow = 0 Then
        '=====================================================
        '�������ɾ��ͼ����ܻ������⣬�ȷŵ���ʱ������
        Call DeleteTmpFile
        '=====================================================
        
        mstrPrivs = gstrPrivs                                       '��ʹ��Ȩ��
        
        Call zlDatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs)
        
        LoadAllData
        
        mintLoadShow = mintLoadShow + 1
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If mintEditState > 0 Then Exit Sub
'    If KeyCode = 38 Then
'        BackOrNextPatient 1
'    ElseIf KeyCode = 40 Then
'        BackOrNextPatient 2
'    End If
End Sub

Private Sub Form_Load()
    Dim strPrivs As String
'    Set objLISComm = CreateObject("Zl9LISComm.clsPublic")
'    '�����������ݽ��ճ�ʼ��
'    If objLISComm Is Nothing Then
'        MsgBox "ͨѶ�����ʼ��ʧ��!", vbExclamation, gstrSysName
'
'        Unload Me
'    End If
'    objLISComm.InitLISComm gcnOracle, Me
    '--------------------------------------------
    '������
    Set clsPluginLoader = New PlugInLoader
    
    ' the interface the plugins have to implement
    ' �������ʵ�ֽӿ�
    Set clsPluginLoader.Interface = New zl9LisQuery_Def.clsLisQuery
    '--------------------------------------------
    
    InitinterFace                                               '��ʼ������
    
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
    
    
    '�����������ݽ��ճ�ʼ��
'    objLISComm.InitLISComm gcnOracle, Me
End Sub

Private Sub DeleteTmpFile()
    'ɾ��������������ʱ�ļ�
    On Error Resume Next
    If Dir(App.path & "\*.BMP") <> "" Then
        Kill App.path & "\*.BMP"
    End If
    If Dir(App.path & "\*.JPG") <> "" Then
        Kill App.path & "\*.JPG"
    End If
    If Dir(App.path & "\*.GIF") <> "" Then
        Kill App.path & "\*.GIF"
    End If
    If Dir(App.path & "\*.CHT") <> "" Then
        Kill App.path & "\*.CHT"
    End If
    If Dir(App.path & "\*.ZIP") <> "" Then
        Kill App.path & "\*.ZIP"
    End If
    If gobjFSO.FolderExists(App.path & "\ZLLIS_ZIP") Then gobjFSO.DeleteFolder App.path & "\ZLLIS_ZIP", True
End Sub

Private Sub CreateDockPane()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    
    
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(Dkp_ID_List, 200, 150, DockLeftOf, Nothing)
    Pane1.Title = "�����嵥"
    Pane1.Handle = Me.PicList.hWnd
'    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(Dkp_ID_Locate, 200, 600, DockRightOf, Nothing)
    Pane2.Title = "���˶�λ"
    Pane2.Handle = Me.PicInfo.hWnd
'    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMain.CreatePane(Dkp_ID_Request, 400, 600, DockBottomOf, Pane2)
    Pane3.Title = "���յǼ�"
    Pane3.Handle = mfrmRequest.hWnd
'    Pane3.Options = PaneNoCaption
    
    Set Pane4 = dkpMain.CreatePane(Dkp_ID_Append, 400, 790, DockRightOf, Pane3)
    Pane4.Title = "���Ӵ���"
    Pane4.Handle = Me.picTab.hWnd
'    Pane4.Options = PaneNoCaption
    
    lngPane5Width = zlDatabase.GetPara("ͼ����", 100, 1208, 200)
    Set Pane5 = dkpMain.CreatePane(Dkp_ID_Image, lngPane5Width, 200, DockRightOf, Pane4)
    Pane5.Title = "ͼ����ʾ"
    Pane5.Handle = Me.PicImage.hWnd
'    Pane5.Options = PaneNoCaption
    
    Call ShowRequest(False)
    
    Pane1.Select
    
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    Dim intLoop As Integer
    On Error Resume Next
    
    If Me.Visible = False Then Exit Sub
    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Locate)
    Pane1.MinTrackSize.SetSize 6954 / Screen.TwipsPerPixelX, 380 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize Pane1.MaxTrackSize.Width, 380 / Screen.TwipsPerPixelY
    
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Request)
    Pane1.MinTrackSize.SetSize 3480 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize 3480 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
    
'    Pane1.MinTrackSize.SetSize 0, 2295 / Screen.TwipsPerPixelY
'    Pane1.MaxTrackSize.SetSize Screen.Width, 2295 / Screen.TwipsPerPixelY
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngloop As Long
    Dim frmThis As Form

    If mintEditState <> 0 Then
        If MsgBox("�����ڱ༭����,�Ƿ�ȷ��Ҫ�˳���", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
            Call SaveDisposal(mFileS.����)
        Else
            Cancel = True
            Exit Sub
        End If
    End If
    
    Call SaveWinState(Me, App.ProductName)
    Me.Visible = False
    mstrAuditingManID = ""
    
    zlDatabase.SetPara "ȱʡ����ID", mlngDeptID, 100, 1208
    zlDatabase.SetPara "��������", mlngMachineID, 100, 1208
    zlDatabase.SetPara "����С��", mstrMachineGroup, 100, 1208
    
    '��������б�
    zlDatabase.SetPara "��ʾ������", Me.cbrthis.FindControl(, conMenu_Manage_ReGet, True, True).Checked, 100, 1208
    zlDatabase.SetPara "���ؼ���ͼ��", Me.cbrthis.FindControl(, conMenu_Manage_LeaveMedi, True, True).Checked, 100, 1208
    'ͼ����ʾ�Ĵ�С
    zlDatabase.SetPara "ͼ����", Me.PicImage.Width / Screen.TwipsPerPixelX, 100, 1208
    '���ļ���ȡʱ��Ĭ����Ϊ���죬���´�ʹ��
    Call zlDatabase.SetPara("�ļ���ȡ��Χ", 0, 100, 1208)
    
    
    '���浱ǰDkp�ķ��,���浽����̫���˻��Ǳ�����ע�����
'    zlDatabase.SetPara "DKP����", dkpMain.SaveStateToString, 100, 1208
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    
    With Me.rptList
        For lngloop = 0 To Me.rptList.SortOrder.Count - 1
            If .SortOrder(lngloop).Caption = "�걾��" Then
                zlDatabase.SetPara "�걾������", .SortOrder(lngloop).SortAscending, 100, 1208
            End If
        Next
    End With
    
    mstrAuditingMan = ""
    mintAuditing = 0
    mintLoadShow = 0
    
    mblnTabList1 = False
    
    '--------------------------------------------------------
    '�ͷŲ��
    Dim i As Long
    '�ͷŵ��õ�DLL
    If Not clsPluginLoader Is Nothing Then
        If clsPluginLoader.PluginCount > 0 Then
            For i = 0 To clsPluginLoader.PluginCount - 1
                Call clsPluginLoader.ClosePlugin(i)
            Next
        End If
        Set clsPluginLoader = Nothing
    End If
        
    For i = LBound(mobjPlugin) To UBound(mobjPlugin)
        Set mobjPlugin(i) = Nothing
    Next
    

    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
'    If mcolSubForm Is Nothing Then
'        For lngLoop = 1 To mcolSubForm.Count
'            Unload mcolSubForm(lngLoop)
'        Next
'    End If
    Set mcolSubForm = Nothing
    
    Me.rptList.Records.DeleteAll
    Me.rptList.Populate
    Me.rptList1.Records.DeleteAll
    Me.rptList1.Populate
    

    Set mclsExpenses = Nothing
    Set mclsInAdvices = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInEPRs = Nothing
'    Set mclsOutEPRs = Nothing
    Set mclsEMR = Nothing


    Unload mfrmRequest
    Unload mfrmWrite
    Unload mfrmWrite2
    Unload mfrmTrack
    Unload mfrmLabMainSampleUnion
    Unload mfrmLabMicrobe3Report
    
    If Not gobjEmr Is Nothing Then
        Call gobjEmr.CloseForms
    End If

    Set mfrmRequest = Nothing
    Set mfrmWrite = Nothing
    Set mfrmWrite2 = Nothing
    Set mfrmTrack = Nothing
    Set mfrmLabMicrobe3Report = Nothing
    Set mfrmLabMainSampleUnion = Nothing
    
    Me.TabCtlWindow.RemoveAll
    Me.TabList.RemoveAll
    Me.cbrChild.DeleteAll
    Me.cbrthis.DeleteAll
    Me.dkpMain.DestroyAll
    
    '=====================================================
    '���ɾ��ͼ���ļ�
    Call DeleteTmpFile
    '=====================================================
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblɸѡ_Click()
    Call picFilter_Click
End Sub

Private Sub mfrmLabMicrobe3Report_StartEdit(Cancel As Boolean)
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    On Error GoTo errH:
    If InStr(",7,8,13,", CStr(Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon)) > 0 Then
        '�Ѽ���
        Cancel = True
        mintHandleState = 0
    Else
        If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
            Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
            ReportDisposal mActR.��д��������
            Cancel = False
        Else
            Cancel = True
        End If
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mfrmRequest_ZlAutoSave(ByVal lngSampleID As Long)
    If lngSampleID = 0 Then Exit Sub
    
    On Error GoTo errH:
    If mintContinue = 0 Then
        '����������
        Me.rptList.Tag = ""   '�����������ı��
        mlngKey = lngSampleID
        If mlngMachineID > 0 Or mlngMachineID = -1 Then
            InsertOneRecored mlngKey, True
        Else
            Call RefreshData
        End If
        Call SaveDisposal(mFileS.����)
        '���պ��Ƿ�����������
        Call SampleDisposal(mActS.��������)

    Else
        Select Case mintEditState
            Case 4
                RefreshData
                '���պ��Ƿ�����������
                Call SampleDisposal(mActS.��������)
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = True Then
                        Call ReportDisposal(mActR.��˱���)
                    End If
                End If
                
                If MoveStation(1, 1) = False Then                       '�����ƶ�
                    'û���ҵ���¼ʱ�˳�����
                    mintHandleState = 0
                    mintEditState = 0
                    Call SaveDisposal(mFileS.����)
                Else
                    Call SampleDisposal(mActS.�����)
                End If
                
            Case Else
                If Me.rptList.Tag = "" Then
                    '��һ������ʱ������б�
                    Me.rptList.Records.DeleteAll
                    Me.rptList.Tag = "Continue"
                End If
                '��Ӹ������ļ�¼���б���
                InsertOneRecored lngSampleID
                 '���պ��Ƿ�����������
                Call SampleDisposal(mActS.��������)
                If mintEditState = 1 Then
                    Call SampleDisposal(mActS.����)
                End If
                If mintEditState = 2 Then
                    Call SampleDisposal(mActS.�Ǽ�)
                End If
        End Select
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mfrmWrite2_StartEdit(Cancel As Boolean)
    If Me.rptList.Rows.Count = 0 Then Exit Sub
    On Error GoTo errH
    If InStr(",7,8,13,", CStr(Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon)) > 0 Then
        '�Ѽ���
        Cancel = True
        mintHandleState = 0
    Else
        '���ڽ��еǼǺ��ղ���ʱ�Զ�����
        If mintEditState >= 1 And mintEditState <= 4 Then
            If Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Visible = True Then
                Call SaveDisposal(mFileS.����)
            End If
        End If
        Cancel = False
'        mintHandleState = 2
        If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
            Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
            ReportDisposal mActR.��д����
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picFilter_Click()
    Dim vRect As RECT
    If Me.picFilter.Tag = "" Then
        Me.picFilter.Tag = "True"
    Else
        Me.picFilter.Tag = ""
    End If
    If Me.TabList.Item(0).Selected = True Then
        vRect = GetControlRect(Me.picFilter.hWnd)
        frmLabMainSizer.ShowMe Me, "������", IIf(Me.picFilter.Tag = "", True, False)
        frmLabMainSizer.Left = vRect.Left - 400
        frmLabMainSizer.Top = vRect.Top + 350
        Call GetVerifying
    Else
        vRect = GetControlRect(Me.picFilter.hWnd)
        frmLabMainSizer.ShowMe Me, "������", IIf(Me.picFilter.Tag = "", True, False)
        frmLabMainSizer.Left = vRect.Left - 400
        frmLabMainSizer.Top = vRect.Top + 350

    End If
    Call RptListFilter
End Sub

Private Sub picFilter_LostFocus()
    If Me.TabList.Item(0).Selected = True Then
        frmLabMainSizer.ShowMe Me, "������", True
        Call GetVerifying
    Else
        frmLabMainSizer.ShowMe Me, "������", True
    End If
    Call RptListFilter
End Sub

Private Sub picList_Click()
    If Me.TabList.Item(0).Selected = True Then
        frmLabMainSizer.ShowMe Me, "������", True
        Call GetVerifying
        If Me.picFilter.Tag = "True" Then Call RptListFilter
        Me.picFilter.Tag = ""
    Else
        frmLabMainSizer.ShowMe Me, "������", True
        Call GetWaitVerify
        If Me.picFilter.Tag = "True" Then Call RptListFilter
        Me.picFilter.Tag = ""
    End If
    
End Sub

Private Sub picList_GotFocus()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
'    Me.cboʱ��.SetFocus
    If Me.TabList.Tag = "" And Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
    Else
'        Me.rptList1.SetFocus
    End If
    Me.TabList.Tag = "Show"
End Sub

Private Sub picList_LostFocus()
'    Me.TabList.Tag = ""
End Sub

Private Sub picList_Resize()
    On Error Resume Next
'    Me.rptList.Top = 0
    Me.TabList.Left = 0
    Me.TabList.Width = PicList.ScaleWidth
    Me.TabList.Height = PicList.ScaleHeight - Me.TabList.Top
    If Me.TabList.Selected.Index = 0 Then
        Me.picFilter.Left = Me.chkSoure(5).Left + Me.chkSoure(5).Width + 30
    Else
        Me.picFilter.Left = Me.chkSoure(2).Left + Me.chkSoure(2).Width + 30
    End If
    Me.cboʱ��.Top = Me.TabList.Top + Me.TabList.Height - Me.cboʱ��.Height
    Me.cboʱ��.Left = 2300
    dtpDate.Top = Me.cboʱ��.Top
    dtpDate.Left = Me.cboʱ��.Left + Me.cboʱ��.Width + 10
    dtpDateEnd.Top = Me.cboʱ��.Top
    dtpDateEnd.Left = Me.dtpDate.Left + Me.dtpDate.Width + 10
End Sub

Private Sub picTab_Resize()
    Me.TabCtlWindow.Top = 0
    Me.TabCtlWindow.Left = 0
    Me.TabCtlWindow.Width = Me.picTab.ScaleWidth
    Me.TabCtlWindow.Height = Me.picTab.ScaleHeight
End Sub
Private Sub CreateTableControl()
    Dim Item As TabControlItem
    'Dim ObjchargeWindow As Object
    Dim strPrivs As String
    
    On Error Resume Next
    
    With Me.TabList
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 0, "������", rptList.hWnd, conMenu_Tool_Report
        .InsertItem 1, "������", rptList1.hWnd, conMenu_Tool_Report
        .PaintManager.Position = xtpTabPositionBottom
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        If zlDatabase.GetPara("��ʾ������", 100, 1208, "False") = "True" Then
            .Item(1).Visible = True
        Else
            .Item(1).Visible = False
        End If
        .Item(0).Selected = True
    End With
    
    
    With Me.TabCtlWindow
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem(0, "������", mfrmWrite.hWnd, conMenu_Tool_Report).Tag = "��ͨ������"
        .InsertItem(1, "������", mfrmWrite2.hWnd, conMenu_Tool_Report).Tag = "΢���ﱨ����"
        .InsertItem(2, "��������", mfrmLabMicrobe3Report.hWnd, conMenu_Tool_Report).Tag = "΢������������"
        .InsertItem(3, "���ζԱ�", mfrmTrack.hWnd, conMenu_Edit_Audit).Tag = "���ζԱ�"
        'Set ObjchargeWindow = mclsExpenses.zlGetForm
        strPrivs = GetPrivFunc(glngSys, pҽ�����ѹ���)  'û��ҽ�����ѹ���Ȩ��ʱ����ʾ
        .InsertItem(4, "���ò�ѯ", PicWindows.hWnd, conMenu_Edit_Price).Tag = IIf(strPrivs <> "", "���ò�ѯ", "")
        .Item(4).Visible = IIf(strPrivs <> "", True, False)
        .InsertItem(5, "�걾�ϲ�", mfrmLabMainSampleUnion.hWnd, conMenu_Edit_Archive).Tag = "�걾�ϲ�"
         strPrivs = GetPrivFunc(glngSys, p����ҽ���´�)
        .InsertItem(6, "����ҽ��", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "����ҽ��", "")
        .Item(6).Visible = IIf(strPrivs <> "", True, False)
        strPrivs = GetPrivFunc(glngSys, pסԺҽ���´�)
        .InsertItem(7, "סԺҽ��", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "סԺҽ��", "")
        .Item(7).Visible = IIf(strPrivs <> "", True, False)
        strPrivs = GetPrivFunc(glngSys, p���ﲡ������)
        .InsertItem(8, "���ﲡ��", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "���ﲡ��", "")
        strPrivs = GetPrivFunc(glngSys, pסԺ��������)
        .InsertItem(9, "סԺ����", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "סԺ����", "")
        strPrivs = GetPrivFunc(glngSys, p�°没������)
        .InsertItem(10, "���Ӳ���", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "���Ӳ���", "")
        
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        .Item(1).Visible = False
        .Item(6).Visible = False
        .Item(7).Visible = False
        .Item(8).Visible = False
        .Item(9).Visible = False
        .Item(10).Visible = False
    End With
    
'    If Me.TabList.Item(0).Selected = True Then
'        cboʱ��.Text = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0)
'        Me.DTPDate.Value = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
'        Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
'    Else
'        cboʱ��.Text = zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��")
'    End If
    cboʱ��.Text = "��  ��"
    Me.dtpDate.Visible = (Me.cboʱ��.Text = "�Զ���")
    Me.dtpDateEnd.Visible = (Me.cboʱ��.Text = "�Զ���")
End Sub
Private Function LoadInterFaceCbo() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lngTmp As Long
    Dim ControlcboDept As CommandBarComboBox
    Dim ControlcboMachine As CommandBarComboBox
    Dim strSQL As String
    
    On Error GoTo errH
    mlngDeptID = zlDatabase.GetPara("ȱʡ����ID", 100, 1208, mlngDeptID)
    mlngMachineID = zlDatabase.GetPara("��������", 100, 1208, mlngMachineID)
    
    '2.��ȡ��������
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = "SELECT A.����||'-'||A.���� as ����,A.ID FROM ���ű� A,��������˵�� B WHERE " & _
                  " (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND " & _
                  " A.ID=B.����ID AND B.��������='����' ORDER BY A.����||'-'||A.����"
    Else
        strSQL = "Select A.���� || '-' || A.���� As ����, A.ID" & vbNewLine & _
                "  From ���ű� A, ��������˵�� B" & vbNewLine & _
                "  Where (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.ID = B.����id And B.�������� = '����' And" & vbNewLine & _
                "        A.ID In (Select Distinct D.ʹ��С��id" & vbNewLine & _
                "                 From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                "                 Where A.С��id = B.ID And B.ID = C.С��id��and C.����id = D.ID And ��Աid = [1] and C.�鿴 = 1)" & vbNewLine & _
                "  Order By A.���� || '-' || A.����"
    End If
    
    cboDept.Clear
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    'If InStr(mstrPrivs, "���п���") > 0 Then
    cboDept.AddItem "���п���"
    
    Do Until rsTmp.EOF
        cboDept.AddItem rsTmp("����")
        cboDept.ItemData(cboDept.NewIndex) = rsTmp("ID")
        If rsTmp("id") = IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID) Then
            cboDept.ListIndex = cboDept.NewIndex
            mlngDeptID = IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID)
'            objLISComm.DeptID = mlngDeptID
        End If
        rsTmp.MoveNext
    Loop
    
    If cboDept.ListCount > 0 And Trim(cboDept.Text) = "" Then
        cboDept.ListIndex = 0
        mlngDeptID = cboDept.ItemData(0)
'        objLISComm.DeptID = mlngDeptID
    End If
    
    
    gstrSql = "select ����ID from ������Ա where ��Աid = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, UserInfo.ID)
    Do Until rsTmp.EOF
        mUserDept = mUserDept & ";" & Nvl(rsTmp("����Id"))
        rsTmp.MoveNext
    Loop
    If mUserDept <> "" Then mUserDept = mUserDept & ";"
    Me.MousePointer = 0
'    If cboDept.ListCount > 0 Then
'
'        cboMachine.Clear
'        cboMachine.AddItem "<��������>": cboMachine.ItemData(cboMachine.NewIndex) = 0
'        cboMachine.AddItem "<�ֹ�>": cboMachine.ItemData(cboMachine.NewIndex) = -1
'        strsql = "SELECT a.����,a.ID FROM �������� a where ʹ��С��ID = [1]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngDeptID)
'        Do Until rsTmp.EOF
'            cboMachine.AddItem rsTmp("����")
'            cboMachine.ItemData(cboMachine.NewIndex) = rsTmp("Id")
'            If rsTmp("id") = mlngMachineID Then
'                cboMachine.ListIndex = cboMachine.NewIndex
'            End If
'            rsTmp.MoveNext
'        Loop
'        If cboMachine.ListCount > 0 And Trim(cboMachine.Text) = "" Then
'            cboMachine.ListIndex = 0
'            mlngMachineID = cboMachine.ItemData(0)
'        End If
'    End If
    Exit Function
errH:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitinterFace()
    '�����ʼ��
    On Error GoTo errH
    
    LoadRegistSetup                     '����ע��������ò���ʼ������������
    CreateTableControl                  '����TAB
    CreateCbs                           '����������
    CreateChildCbs
    CreateDockPane                      '������������
    CreaterptListHead                   '�����б�ͷ
    
    On Error Resume Next
    With Me.WinsockC                    '��ʼ���ͽ��ճ����ͨѶ�ӿ�
        .Protocol = sckUDPProtocol
        .RemoteHost = "Localhost"
        .RemotePort = 1000
        .Bind 1001
    End With
    Exit Sub
errH:

    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadAllData()
    
    On Error GoTo errH
    
    '��������
    Call GetVerifying                   '��������й�������
    Call GetWaitVerify                  '��������չ�������
    LoadInterFaceCbo                    '���������Ϳ���
    RefreshData                         'ˢ��
    RptListFilter                       '�����б�ˢ��
    rptList_SelectionChanged            'ˢ��״̬
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Private Function GetQuerySQL(ByVal strCondition As String, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim mlngLoop As Long
     
    On Error Resume Next
    '�����Ǹ��������������ɵ��������
    
    If strCondition = "" Then Exit Function
    
    varTmp = Split(strCondition, "^")
'    If bytMode = 1 Then
'        If Val(varTmp(0)) > 0 Then GetQuerySQL = GetQuerySQL & " AND A.ִ�п���ID + 0 = " & Val(varTmp(0))
'    Else
'        If Val(varTmp(0)) > 0 Then GetQuerySQL = GetQuerySQL & " AND c.ִ�п���ID + 0 = " & Val(varTmp(0))
'    End If
    If Val(varTmp(0)) > 0 Then GetQuerySQL = GetQuerySQL & " AND c.ִ�п���ID + 0 = " & Val(varTmp(0))
    
    If Val(varTmp(1)) > 0 Then GetQuerySQL = GetQuerySQL & " AND c.����ID = " & Val(varTmp(1))
    If Trim(varTmp(2)) <> "��  ��" Then
        Select Case Trim(varTmp(2))
        Case "ָ  ��"
            GetQuerySQL = GetQuerySQL & " AND c.����ʱ�� BETWEEN TO_DATE('" & Format(varTmp(3), "yyyy-mm-dd hh:mm") & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(4), "yyyy-mm-dd hh:mm") & "', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            GetQuerySQL = GetQuerySQL & " AND c.����ʱ�� BETWEEN TO_DATE('" & GetDateTime(varTmp(2), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(2), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
    End If
    varTmp2 = Split(Trim(varTmp(5)), ",")
    strTmp = ""
    For mlngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(mlngLoop), "��") = 0 Then
            strTmp = strTmp & "  OR c.�걾���=" & TransSampleNO(varTmp2(mlngLoop))
        Else
            strTmp = strTmp & "  OR c.�걾��� BETWEEN " & TransSampleNO(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "��") - 1)) & " AND " & TransSampleNO(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "��") + 1))
        End If
    Next
    If strTmp <> "" Then GetQuerySQL = GetQuerySQL & " AND (1=2 " & strTmp & ")"

    If Trim(varTmp(6)) <> "" Then GetQuerySQL = GetQuerySQL & " AND c.������='" & Trim(varTmp(6)) & "'"
    If Trim(varTmp(7)) <> "" Then GetQuerySQL = GetQuerySQL & " AND c.�����='" & Trim(varTmp(7)) & "'"
    
    If Trim(varTmp(8)) <> "��  ��" Then
        
        Select Case Trim(varTmp(8))
        Case "ָ  ��"
            GetQuerySQL = GetQuerySQL & " AND c.���ʱ�� BETWEEN TO_DATE('" & Format(varTmp(9), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(10), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            GetQuerySQL = GetQuerySQL & " AND c.���ʱ�� BETWEEN TO_DATE('" & GetDateTime(varTmp(8), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(8), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
        
    End If
    
    If Val(varTmp(11)) > 0 Then
'        If bytMode = 1 Then
'            GetQuerySQL = GetQuerySQL & " AND F.ִ��״̬ = " & IIf(Val(varTmp(11)) = 1, "3", "1")
'        Else
'            GetQuerySQL = GetQuerySQL & " AND c.����״̬ = " & IIf(Val(varTmp(11)) = 1, "1", "2")
'        End If
        GetQuerySQL = GetQuerySQL & " AND c.����״̬ = " & IIf(Val(varTmp(11)) = 1, "1", "2")
    End If
    
    If Val(varTmp(12)) > 0 Then
        GetQuerySQL = GetQuerySQL & " AND c.ID IN (SELECT G.����걾ID FROM ������ͨ��� G,������Ŀ H WHERE H.������Ŀid=G.������Ŀid AND G.����걾ID=c.ID "
        GetQuerySQL = GetQuerySQL & " AND G.������ĿID=" & Val(varTmp(12))
        
        If Val(varTmp(13)) = 1 Then
            GetQuerySQL = GetQuerySQL & " AND H.�������=1 AND DECODE(H.�������,1,TO_NUMBER(G.������),0)"
            strTmp = Val(varTmp(16))
        Else
            GetQuerySQL = GetQuerySQL & " AND G.������"
            strTmp = "'" & varTmp(16) & "'"
        End If
        
        Select Case varTmp(15)
        Case "����"
            GetQuerySQL = GetQuerySQL & ">" & strTmp
        Case "С��"
            GetQuerySQL = GetQuerySQL & "<" & strTmp
        Case "���ڵ���"
            GetQuerySQL = GetQuerySQL & ">=" & strTmp
        Case "С�ڵ���"
            GetQuerySQL = GetQuerySQL & "<=" & strTmp
        Case "������"
            GetQuerySQL = GetQuerySQL & "<>" & strTmp
        Case "����"
            GetQuerySQL = GetQuerySQL & " LIKE '%" & varTmp(16) & "%'"
        Case "�ڷ�Χ��"
            If Val(varTmp(13)) = 1 Then
                GetQuerySQL = GetQuerySQL & " BETWEEN " & strTmp & " AND " & Val(varTmp(17))
            Else
                GetQuerySQL = GetQuerySQL & " BETWEEN " & strTmp & " AND '" & varTmp(17) & "'"
            End If
        Case Else
            GetQuerySQL = GetQuerySQL & "=" & strTmp
        End Select
        GetQuerySQL = GetQuerySQL & ")"
    End If
    
    If bytMode = 1 Then
        If Trim(varTmp(18)) <> "" Then GetQuerySQL = GetQuerySQL & " AND b.���� Like '" & Trim(varTmp(18)) & "%'"
        If Val(varTmp(19)) > 0 Then GetQuerySQL = GetQuerySQL & " AND A.���˿���ID = " & Val(varTmp(19))
        If Val(varTmp(20)) > 0 Then GetQuerySQL = GetQuerySQL & " AND b.סԺ��=" & varTmp(20)
        If Val(varTmp(21)) > 0 Then GetQuerySQL = GetQuerySQL & " AND b.��ǰ����=" & Val(varTmp(21))
        If Val(varTmp(22)) > 0 Then GetQuerySQL = GetQuerySQL & " AND b.�����=" & varTmp(22)
'        If Trim(varTmp(23)) <> "" Then GetQuerySQL = GetQuerySQL & " AND A.����ҽ��='" & Trim(varTmp(23)) & "'"
'        If Trim(varTmp(UBound(varTmp))) <> "" Then GetQuerySQL = GetQuerySQL & " AND A.����ҽ��='" & Trim(varTmp(UBound(varTmp))) & "'"
        If Val(varTmp(24)) > 0 Then GetQuerySQL = GetQuerySQL & " AND A.��������ID = " & Val(varTmp(24))
        
        
        If Trim(varTmp(25)) <> "��  ��" Then
            Select Case Trim(varTmp(25))
            Case "ָ  ��"
                GetQuerySQL = GetQuerySQL & " AND A.����ʱ�� BETWEEN TO_DATE('" & Format(varTmp(26), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(27), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
            Case Else
                GetQuerySQL = GetQuerySQL & " AND A.����ʱ�� BETWEEN TO_DATE('" & GetDateTime(varTmp(25), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(25), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
            End Select
        End If
    Else
        If Trim(varTmp(18)) <> "" Or Val(varTmp(19)) > 0 Or Val(varTmp(20)) > 0 Or _
                Val(varTmp(21)) > 0 Or _
                Val(varTmp(22)) > 0 Or _
                Trim(varTmp(23)) <> "" Or _
                Val(varTmp(24)) > 0 Or _
                Trim(varTmp(25)) <> "��  ��" Then
                
            GetQuerySQL = GetQuerySQL & " AND 1=1 "
        End If
    End If
    
    If Trim(varTmp(28)) <> "" Then GetQuerySQL = GetQuerySQL & " AND c.������='" & Trim(varTmp(28)) & "'"
    
    If Trim(varTmp(29)) <> "��  ��" Then
        Select Case Trim(varTmp(29))
        Case "ָ  ��"
            GetQuerySQL = GetQuerySQL & " AND c.����ʱ�� BETWEEN TO_DATE('" & Format(varTmp(30), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(31), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            GetQuerySQL = GetQuerySQL & " AND c.����ʱ�� BETWEEN TO_DATE('" & GetDateTime(varTmp(29), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(29), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
    End If
    
    If Trim(varTmp(32)) <> "��������" And InStr(varTmp(32), "-") > 0 Then GetQuerySQL = GetQuerySQL & " AND c.�걾����='" & zlCommFun.GetNeedName(Trim(varTmp(32))) & "'"
    
    'If GetQuerySQL <> "" Then GetQuerySQL = " AND " & GetQuerySQL
    
End Function



Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    frmLabMainSizer.ShowMe Me, "������", True
    Call GetVerifying
    If Me.picFilter.Tag = "True" Then Call RptListFilter
    Me.picFilter.Tag = ""
'    Me.cboʱ��.SetFocus
'    Me.rptList.SetFocus
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    On Error Resume Next
    If Button = 2 Then
        If rptList.Records.Count <= 0 Then Exit Sub
        If Not rptList.SelectedRows(0).GroupRow Then
            Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������(&A)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ȡ�����(&U)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "״̬�ع�(&Z)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "��������(&D)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ������(&E)")
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "�����ѯ(&P)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "�������͵�����(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "��Ϊ�ʿ�(&Q)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "��Ϊ�ȶ�(&Y)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "�鿴�ȶ�(&B)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "�������ϲ�(&E)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸������ź�����(&M)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "ɾ������(&D)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "����(&J)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "����(&C)")
            End With
            objPopup.ShowPopup
        End If
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error GoTo errH
    If Row.Record(mCol.�걾����).Icon = -1 And InStr(",7,8,13,", Row.Record(mCol.ִ��״̬).Icon) = 0 Then
        If Me.TabCtlWindow.Item(5).Selected = False Then
            mintHandleState = 1
            If Me.cbrthis.FindControl(, conMenu_Manage_Receive, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Manage_Receive, , True).Visible = True Then
                Call SampleDisposal(mActS.�����)
            End If
        Else
            If Me.cbrthis.FindControl(, conMenu_Edit_Insert, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Edit_Insert, , True).Visible = True Then
                Call SampleDisposal(mActS.�ϲ��걾)
            End If
        End If
    ElseIf Row.Record(mCol.�걾����).Icon = 3 Then
        Call frmLabMainLJ.ShowMe(mlngKey, Me, mlngMachineID)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptList_SelectionChanged()
    Dim strSampleType As String                     '�걾����-1=��ͨ,3=�ʿ�,4=�ȶ�
    Dim strEmergen                                  '���� -1="��ͨ",1=��
    Dim strState                                    '״̬ 7,8=�Ѽ���
    Dim i As Integer                                '��ʱ����
    Dim str��� As String
    Dim rs As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim lngSampleID As Long
    Dim intRow As Integer
    Dim strPricegrade As String                     '�۸�ȼ�
    
    Dim tmp As Double
    
    On Error GoTo errH
    If Me.Visible = False Then Exit Sub
    
    strSampleType = ""
    strEmergen = ""
    strState = ""
    
    Select Case mintEditState
        Case 1, 2, 4
            If Me.rptList.Tag = "" And mlngKey <> Me.rptList.FocusedRow.Record(mCol.ID).Value Then
                lngSampleID = mfrmRequest.ZlSave()
                mintEditState = 0
                If lngSampleID = 0 Then
                    mfrmRequest.ZlCancel
                Else
                    intRow = Me.rptList.FocusedRow.Index
                    InsertOneRecored lngSampleID, False
                    Me.rptList.FocusedRow = Me.rptList.Rows(intRow)
                End If
                
                
                gintSelectFocus = 1
'                Exit Sub
            Else
'                Me.rptList.SetFocus
                gintSelectFocus = 2
                
            End If
        Case 5
            If TabCtlWindow.Item(0).Selected = True Then
                mfrmWrite.ZlSave
                mfrmWrite.ZlCancel
                mfrmWrite.zlRefresh mlngKey
            Else
                mfrmWrite2.ZlSave
                mfrmWrite2.ZlCancel
                mfrmWrite2.zlRefresh mlngKey
            End If
            mintEditState = 0
    End Select
    
    If Me.rptList.FocusedRow Is Nothing Then
        mlngKey = 0
        strSampleType = ""
    Else
        If Me.rptList.FocusedRow.Record(mCol.ID).Value = mlngKey And mblnCompelRefresh = False Then
            'ͬһIDʱ��ˢ��
            Exit Sub
        End If
        mblnCompelRefresh = False
        mlngKey = Me.rptList.FocusedRow.Record(mCol.ID).Value
        i = Me.rptList.FocusedRow.Record(mCol.�걾����).Icon
        If i = -1 Then
            strSampleType = "��ͨ����"
        ElseIf i = 3 Then
            strSampleType = "�ʿ�����"
        Else
            strSampleType = "�ȶ�����"
        End If
        
        i = Me.rptList.FocusedRow.Record(mCol.����).Icon
        If i = 1 Then
            strEmergen = "����"
        End If
        
        i = Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon
        If i = 7 Or i = 8 Then
            strState = "�Ѽ���"
        Else
            strState = "������"
        End If
        
        
    End If
    
    If mintLoadShow = 0 Then Exit Sub
    
    Call mfrmRequest.zlRefresh(Me.rptList.FocusedRow)
    Call mfrmLabMicrobe3Report.zlRefresh(mlngKey)
    
    
'    If Me.rptList.FocusedRow Is Nothing Then
'        Call mfrmWrite.zlRefresh(mlngKey)
'    ElseIf Val(Me.rptList.FocusedRow.Record(mCol.΢����걾).Value) = 1 Then
'        Call mfrmWrite2.zlRefresh(mlngKey)
'    Else
'        Call mfrmWrite.zlRefresh(mlngKey)
'    End If
'
'    Call mfrmRequest.zlRefresh(mlngkey)
    
    RefreshTableWindow Me.TabCtlWindow.Selected.Index
    If mlngKey <> 0 Then
        ReadImageData mlngKey, False
    End If
    
    
    
    Set cbrControl = Me.cbrChild.FindControl(, conMenu_View_FindType, True, True)
    If Not cbrControl Is Nothing Then
        If mlngKey > 0 Then '��ʾ��Ŀ���ͼ۸�
        
            '��ȡ�۸�ȼ�
            With Me.rptList.FocusedRow
                strPricegrade = GetAdvicePrice(Val(.Record(mCol.����ID).Value), Val(.Record(mCol.��ҳID).Value))
            End With
            
            If Val(Me.rptList.FocusedRow.Record(mCol.΢����걾).Value) = 1 Then
                gstrSql = "Select /*+ rule */" & vbNewLine & _
                        " Sum(Nvl(�շ�����, 0) * Nvl(�ּ�, 0)) As ���" & vbNewLine & _
                        "From ( --- ���ݱ걾��¼�е� ���룬ҽ��id,���id���õ���Ӧ��������Ŀid" & vbNewLine & _
                        "       Select C.������Ŀid" & vbNewLine & _
                        "       From ����ҽ����¼ C, ����ҽ������ B, ����걾��¼ A" & vbNewLine & _
                        "       Where B.ҽ��id = C.ID And A.�������� = B.�������� And A.ID = [1]" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select C.������Ŀid" & vbNewLine & _
                        "       From ����ҽ����¼ C, ����걾��¼ A" & vbNewLine & _
                        "       Where A.ҽ��id = C.ID And A.ID = [1]" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select C.������Ŀid From ����ҽ����¼ C, ����걾��¼ A Where A.ҽ��id = C.���id And A.ID = [1]) A," & vbNewLine & _
                        "     (Select E.������Ŀid, E.�շ�����, F.�ּ�, J.����, J.����" & vbNewLine & _
                        "       From �����շѹ�ϵ E, �շѼ�Ŀ F, �շ���ĿĿ¼ J" & vbNewLine & _
                        "       Where F.�շ�ϸĿid = J.ID And E.�շ���Ŀid = F.�շ�ϸĿid And (F.��ֹ���� Is Null Or F.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) and f.�۸�ȼ�" & strPricegrade & ") B" & vbNewLine & _
                        "Where A.������Ŀid = B.������Ŀid"

            Else
    
                gstrSql = "Select /*+ rule */" & vbNewLine & _
                            " Sum(Nvl(�շ�����, 0) * Nvl(�ּ�, 0)) As ���" & vbNewLine & _
                            "From (Select Distinct ������Ŀid" & vbNewLine & _
                            "       From (Select ������Ŀid" & vbNewLine & _
                            "              From ������ͨ���" & vbNewLine & _
                            "              Where ����걾id = [1] And ������Ŀid Is Not Null" & vbNewLine & _
                            "              Union All" & vbNewLine & _
                            "              Select B.������Ŀid" & vbNewLine & _
                            "              From ������ͨ��� A, ���鱨����Ŀ B, ������ĿĿ¼ C" & vbNewLine & _
                            "              Where A.������Ŀid = B.������Ŀid And B.������Ŀid = C.ID And C.�����Ŀ = 0" & vbNewLine & _
                            "              And A.����걾id = [1] And A.������Ŀid Is Null)) A," & vbNewLine & _
                            "     (Select E.������Ŀid, E.�շ�����, F.�ּ�, J.����, J.����" & vbNewLine & _
                            "       From �����շѹ�ϵ E, �շѼ�Ŀ F, �շ���ĿĿ¼ J" & vbNewLine & _
                            "       Where F.�շ�ϸĿid = J.ID And E.�շ���Ŀid = F.�շ�ϸĿid" & vbNewLine & _
                            "             And (F.��ֹ���� Is Null Or F.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) and f.�۸�ȼ�" & strPricegrade & ") B" & vbNewLine & _
                            "Where A.������Ŀid = B.������Ŀid"
            End If
            Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            str��� = ""
            If rs.RecordCount > 0 Then
                If Val("" & rs.Fields("���")) <> 0 Then
                    str��� = "����Ŀ�Ƽ� " & Format("" & rs.Fields("���"), "0.00")
                End If
            End If

            gstrSql = "Select Count(����걾id) as ��Ŀ�� from ������ͨ��� where ������ is not null And ����걾id =[1] "
            Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rs.RecordCount > 0 Then
                str��� = "����" & rs.Fields("��Ŀ��") & "����" & str��� & "  "

            End If
        End If '--��ʾ��Ŀ���ͼ۸�
        
        cbrControl.Caption = str��� & "   ״̬;" & strEmergen & " " & strState & " " & strSampleType
        Me.cbrChild.RecalcLayout
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptList1_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    frmLabMainSizer.ShowMe Me, "������", True
    Call GetWaitVerify
    If Me.picFilter.Tag = "True" Then Call RptListFilter
    Me.picFilter.Tag = ""
End Sub

Private Sub rptList1_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    
    On Error Resume Next
    If Button = 2 Then
        If rptList1.Records.Count <= 0 Then Exit Sub
        If Not rptList1.SelectedRows(0).GroupRow Then
            Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������(&A)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ȡ�����(&U)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "״̬�ع�(&Z)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "��������(&D)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ������(&E)")
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "�����ѯ(&P)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "��������(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "��Ϊ�ʿ�(&Q)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "��Ϊ�ȶ�(&Y)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "�鿴�ȶ�(&B)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "�������ϲ�(&E)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸�������(&M)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "ɾ������(&D)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "����(&J)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "����(&C)")
            End With
            objPopup.ShowPopup
        End If
    End If
End Sub

Private Sub rptList1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If mintEditState = 0 Then Call SampleDisposal(mActS.����)
End Sub

Private Sub TabCtlWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible = True And mTableRefresh = False Then
        RefreshTableWindow Item.Index
        Me.TabCtlWindow.Tag = Item.Index
    End If
End Sub


Private Sub RefreshTableWindow(Index As Integer)
    Dim blnCurrMoved As Boolean                                             '�Ƿ�ת��
    Dim lngAdviceID As Long                                                 'ҽ��ID
    Dim intReportCount As Integer                                           '��������
    Dim blMicrobe As Boolean                                                '�Ƿ���΢����
    Dim cbrControl As CommandBarControl                                     '�������а�ť����
    Dim strPatientType As String                                            '������Դ
    Dim str�Һŵ� As String                                                 '�Һŵ�
    Dim lngPatientID As Long                                                '����ID
    Dim intHomePage As Integer                                              '��ҳID
    Dim lngPatientDeptID As Long                                            '���˿���ID
    Dim blnShowButtonText As Boolean                                        '��ʾ��ť�ı�
    Dim lngCount As Long
    Dim cbrCustom As CommandBarControlCustom
    Dim lng�Һ�ID As Long
    Dim strPrivs As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            lngAdviceID = Val(.Record(mCol.ҽ��id).Value)
            blnCurrMoved = (.Record(mCol.ת��).Value = "��")
            intReportCount = Val(.Record(mCol.�������).Value)
            blMicrobe = IIf(Val(.Record(mCol.΢����걾).Value) = 1, True, False)
            strPatientType = .Record(mCol.�������).Value
            str�Һŵ� = .Record(mCol.�Һŵ�).Value
            lngPatientID = Val(.Record(mCol.����ID).Value)
            intHomePage = Val(.Record(mCol.��ҳID).Value)
            lngPatientDeptID = Val(.Record(mCol.��������ID).Value)
        End With
    End If
    
'    Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = False
    
    '�����΢����ʱ����ȶ�Ϊ����ʾ
    If blMicrobe = True Then
        Me.TabCtlWindow.Item(3).Visible = False
        Me.TabCtlWindow.Item(2).Visible = True
    Else
        Me.TabCtlWindow.Item(3).Visible = True
        Me.TabCtlWindow.Item(2).Visible = False
    End If
    
    'ɾ�����ɵİ�ť
    DelButton Index
    'ˢ���Ӵ��ڲ˵�
'    Call LockWindowUpdate(Me.Hwnd)
    
'    If Me.TabCtlWindow.Selected.Index <> Val(Me.TabCtlWindow.Tag) Then
'        'ɾ�����ڵĹ������������˵���
'        For lngCount = cbrthis.ActiveMenuBar.Controls.Count To 1 Step -1
'            cbrthis.ActiveMenuBar.Controls(lngCount).Delete
'        Next
'        For lngCount = cbrthis.Count To 2 Step -1
'            cbrthis(lngCount).Delete
'        Next
'        '���´����˵�
'        Call CreateCbs
'    End If
    
    If strPatientType = "סԺ" Then
        
        If Me.TabCtlWindow.Item(7).Tag = "סԺҽ��" Then
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = True
            If Index = 6 Or Index = 7 Then Index = 7: Me.TabCtlWindow.Item(7).Selected = True
        Else
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = False
        End If
        If TabCtlWindow.ItemCount >= 10 Then
            If Me.TabCtlWindow.Item(9).Tag = "סԺ����" Then
                Me.TabCtlWindow.Item(9).Visible = True
                Me.TabCtlWindow.Item(8).Visible = False
                If Index = 8 Or Index = 9 Then Index = 9: Me.TabCtlWindow.Item(9).Selected = True
            Else
                Me.TabCtlWindow.Item(8).Visible = False
                Me.TabCtlWindow.Item(9).Visible = False
            End If
        End If
        If TabCtlWindow.ItemCount >= 11 Then
            '���Ӳ���
            strPrivs = GetPrivFunc(glngSys, 2252)
            Me.TabCtlWindow.Item(10).Visible = IIf(strPrivs <> "", True, False)
        End If
    Else

        If Me.TabCtlWindow.Item(6).Tag = "����ҽ��" Then
            Me.TabCtlWindow.Item(6).Visible = True
            Me.TabCtlWindow.Item(7).Visible = False
            If Index = 6 Or Index = 7 Then Index = 6: Me.TabCtlWindow.Item(6).Selected = True
        Else
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = False
        End If
        If Me.TabCtlWindow.Item(8).Tag = "���ﲡ��" Then
            Me.TabCtlWindow.Item(8).Visible = True
            If TabCtlWindow.ItemCount >= 10 Then
                Me.TabCtlWindow.Item(9).Visible = False
            End If
            If Index = 8 Or Index = 9 Then Index = 8: Me.TabCtlWindow.Item(8).Selected = True
        Else
            Me.TabCtlWindow.Item(8).Visible = False
            If TabCtlWindow.ItemCount >= 10 Then
                Me.TabCtlWindow.Item(9).Visible = False
            End If
        End If
        If TabCtlWindow.ItemCount >= 11 Then
            '���Ӳ���
            strPrivs = GetPrivFunc(glngSys, 2251)
            Me.TabCtlWindow.Item(10).Visible = IIf(strPrivs <> "", True, False)
        End If
    End If
    
    Select Case Index
        Case 0, 1, 2  '��ͨ�����΢������
            If blMicrobe = True Then
                Me.TabCtlWindow.Item(0).Visible = False
                Me.TabCtlWindow.Item(1).Visible = True
                Me.TabCtlWindow.Item(2).Visible = True
                If mintEditState <> 5 Then
                    mfrmWrite2.zlRefresh mlngKey
                End If
                If Index = 0 Then
                    Me.TabCtlWindow.Item(1).Selected = True
                Else
                    Me.TabCtlWindow.Item(Index).Selected = True
                End If
'                Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).HideFlags = xtpHideGeneric
'                Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = False
            Else
                Me.TabCtlWindow.Item(0).Visible = True
                Me.TabCtlWindow.Item(1).Visible = False
                Me.TabCtlWindow.Item(2).Visible = False
                If mintEditState <> 5 Then
                    mfrmWrite.zlRefresh mlngKey
                End If
                Me.TabCtlWindow.Item(0).Selected = True
'                Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = False
'                Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = False
            End If
        Case 3 '��ʷ�ȶ�
'            Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = False
'            Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = False
            zlCommFun.ShowFlash " ���Դ����ڶ���������ʷ����..."
            mfrmTrack.zlRefresh mlngKey
            zlCommFun.StopFlash
        Case 4  '����
            '��һ�δ�ʱ�ټ���
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsExpenses Is Nothing Then
                Set mclspublicExpenses = New zlPublicExpense.clsPublicExpense
                Call mclspublicExpenses.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
                Set mclsExpenses = New zlPublicExpense.clsDockExpense       '���ò���
                mcolSubForm.Add mclsExpenses.zlGetForm, "_����"             '�õ��Ӵ���
            End If
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "���ò�ѯ", mcolSubForm("_����").hWnd, conMenu_Edit_Price).Tag = "���ò�ѯ"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            mclsExpenses.zlDefCommandBars Me, Me.cbrthis
            strSQL = "select a.id as ҽ��ID, b.���ͺ� from ����ҽ����¼ a,����ҽ������ b " & vbCrLf & _
                    " Where a.ID = b.ҽ��id And a.���id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngAdviceID)
            If rsTmp.EOF = False Then
                mclsExpenses.zlRefresh mlngDeptID, rsTmp(0) & ":" & rsTmp(1), blnCurrMoved
            End If
'            DelButton Index  '����ǰ��ɾ����ť
            
            '�Ƿ���ʾ��ť������
            blnShowButtonText = Me.cbrthis.FindControl(, conMenu_View_ToolBar_Text, True, True).Checked
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(blnShowButtonText, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            
            strSQL = "select distinct c.����,b.���id as ҽ��ID from ������Ŀ�ֲ� a , ����ҽ����¼ b , ������ĿĿ¼ c " & _
                     " where a.ҽ��id = b.���ID and b.������ĿID = c.id and  a.�걾ID =[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ò�ѯ", mlngKey)
            
            If cboExesItem.ListCount > 0 And cboExesItem.ListIndex <> -1 Then lngAdviceID = cboExesItem.ItemData(cboExesItem.ListIndex)
            Me.cboExesItem.Clear
            
            Do Until rsTmp.EOF
                Me.cboExesItem.AddItem rsTmp("����")
                Me.cboExesItem.ItemData(Me.cboExesItem.NewIndex) = rsTmp("ҽ��ID")
                If rsTmp("ҽ��ID") = lngAdviceID Then Me.cboExesItem.ListIndex = Me.cboExesItem.NewIndex
                rsTmp.MoveNext
            Loop
            If cboExesItem.ListCount > 0 Then
                If cboExesItem.ListIndex = -1 Then cboExesItem.ListIndex = 0
            End If
'            Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = True
'            Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = True
'            Me.cbrthis.RecalcLayout
'            Me.cbrChild.RecalcLayout
        Case 5  '�ϲ�
'            Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = True
'            Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = True
        Case 6 '����ҽ��
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsOutAdvices Is Nothing Then
                Set mclsOutAdvices = New zlCISKernel.clsDockOutAdvices      '����ҽ��
                mcolSubForm.Add mclsOutAdvices.zlGetForm, "_����ҽ��"
            End If
            '��һ�δ�ʱ�ټ���
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "����ҽ��", mcolSubForm("_����ҽ��").hWnd, 1).Tag = "����ҽ��"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(6).Visible = True Then
'                DelButton Index  '����ǰ��ɾ����ť
                mclsOutAdvices.zlDefCommandBars Me, Me.cbrthis, 2
                '�Ƿ���ʾ��ť������
                blnShowButtonText = Me.cbrthis.FindControl(, conMenu_View_ToolBar_Text, True, True).Checked
                For Each cbrControl In Me.cbrthis(2).Controls
                    cbrControl.Style = IIf(blnShowButtonText, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
'                Me.cbrthis.RecalcLayout
    '            MsgBox "����ID:" & lngPatientID & ";�Һŵ�:" & str�Һŵ�
                mclsOutAdvices.zlRefresh lngPatientID, str�Һŵ�, True, False, lngAdviceID, mlngDeptID
            End If
        Case 7 'סԺҽ��
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsInAdvices Is Nothing Then
                Set mclsInAdvices = New zlCISKernel.clsDockInAdvices
                mcolSubForm.Add mclsInAdvices.zlGetForm, "_סԺҽ��"
            End If
            '��һ�δ�ʱ�ټ���
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "סԺҽ��", mcolSubForm("_סԺҽ��").hWnd, 1).Tag = "סԺҽ��"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(7).Visible = True Then
'                DelButton Index  '����ǰ��ɾ����ť
                mclsInAdvices.zlDefCommandBars Me, Me.cbrthis, 2
                '�Ƿ���ʾ��ť������
                blnShowButtonText = Me.cbrthis.FindControl(, conMenu_View_ToolBar_Text, True, True).Checked
                For Each cbrControl In Me.cbrthis(2).Controls
                    cbrControl.Style = IIf(blnShowButtonText, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
'                Me.cbrthis.RecalcLayout
    '            MsgBox "����ID:" & lngPatientID & ";��ҳID:" & intHomePage & ";����ID:" & lngPatientDeptID & ";���˿���ID;" & lngPatientDeptID
                mclsInAdvices.zlRefresh lngPatientID, intHomePage, lngPatientDeptID, lngPatientDeptID, 0, False, lngAdviceID, 0, mlngDeptID
            End If
        Case 8 '���ﲡ��
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsOutEPRs Is Nothing Then
                Set mclsOutEPRs = New zlRichEPR.cDockOutEPRs                '����ҽ��
                mcolSubForm.Add mclsOutEPRs.zlGetForm, "_���ﲡ��"
            End If
            '��һ�δ�ʱ�ټ���
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "���ﲡ��", mcolSubForm("_���ﲡ��").hWnd, 1).Tag = "���ﲡ��"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(8).Visible = True Then
                gstrSql = "select ID from ���˹Һż�¼ where ��¼״̬=1 and ��¼����=1 and no = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str�Һŵ�)
                If rsTmp.EOF Then
                    lng�Һ�ID = 0
                Else
                    lng�Һ�ID = Nvl(rsTmp("ID"))
                End If
                mclsOutEPRs.zlRefresh lngPatientID, lng�Һ�ID, mlngDeptID, False
            End If
        Case 9 'סԺ����
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsInEPRs Is Nothing Then
                Set mclsInEPRs = New zlRichEPR.cDockInEPRs                  'סԺ����
                mcolSubForm.Add mclsInEPRs.zlGetForm, "_סԺ����"
            End If
            '��һ�δ�ʱ�ټ���
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "סԺ����", mcolSubForm("_סԺ����").hWnd, 1).Tag = "סԺ����"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(9).Visible = True Then
                mclsInEPRs.zlRefresh lngPatientID, intHomePage, lngPatientDeptID
            End If
        Case 10 '���Ӳ���
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsEMR Is Nothing Then
                Set mclsEMR = CreateObject("zlRichEMR.clsDockEMR")
                If Not mclsEMR Is Nothing Then
                    Set gobjEmr = gfrmMain.mobjEMR
                    If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                      Set mclsEMR = Nothing
                    End If
                End If
                mcolSubForm.Add mclsEMR.zlGetForm, "_���Ӳ���"
            End If
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "���Ӳ���", mcolSubForm("_���Ӳ���").hWnd, 1).Tag = "���Ӳ���"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(10).Visible = True And Not mclsEMR Is Nothing And lngPatientID <> 0 Then
                If strPatientType = "סԺ" Then
                    mclsEMR.zlRefresh lngPatientID, intHomePage, lngPatientDeptID, 0, 2
                Else
                    gstrSql = "select ID from ���˹Һż�¼ where ��¼״̬=1 and ��¼����=1 and no = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str�Һŵ�)
                    If rsTmp.EOF Then
                        lng�Һ�ID = 0
                    Else
                        lng�Һ�ID = Nvl(rsTmp("ID"))
                    End If
                    mclsEMR.zlRefresh lngPatientID, lng�Һ�ID, lngPatientDeptID, 0, 1
                End If
            End If
    End Select
    
    
'    Me.cbrthis.FindControl(, conMenu_Edit_Insert).Visible = IIf(Index = 4, True, False)
'    cbrThis.ActiveMenuBar.FindControl(, conMenu_LIS_RightMenu).Visible = False
    
'    Me.cbrthis.RecalcLayout
'    Me.cbrChild.RecalcLayout
    
    '�������RecalcLayout����������
'    Call LockWindowUpdate(0)
    
    
    mTableRefresh = False
    Exit Sub
errH:
    mTableRefresh = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Property Let AutoRefresh(vData As Boolean)
    '
    '����:�Զ�ˢ��
    '
'
End Property
Private Sub QUFilter()
    '����        ���ٲ�ѯ
    Dim strCondition As String
    '�����ѯ
    AutoRefresh = False
    strCondition = rptList.Tag
    frmLabFilter.ShowMe Me, mlngDeptID, mlngMachineID, mstrMachineALL, strCondition
    If strCondition <> "" Then
        rptList.Tag = strCondition & ";" & 0                            '�������ϲ���ID
        zlCommFun.ShowFlash "���ڸ����������Ժ�...", Me
        RefreshData True
        zlCommFun.StopFlash
    End If
    AutoRefresh = True
End Sub


Private Sub GetSaveSetup(Mode As Integer)
    '������ȡ�����¼
    '����              =1��һ =2����
    Dim strFile As String, lngDeviceID As Long, dtStart As Date, dtEnd As Date, strSampleNO As String
    Dim lngMachineID As Long                                '����ID
    Dim strSampltDate As String                             '�걾ʱ��
    Dim strSampltID As String                                '�걾��
    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            lngMachineID = Val(.Record(mCol.����id).Value)
            strSampltDate = .Record(mCol.�걾ʱ��).Value
            strSampltID = .Record(mCol.�걾��).Value
        End With
    End If
    
    Me.MousePointer = vbHourglass

    If Mode = 1 Then
        strFile = zlDatabase.GetPara("���������ļ�", 100, 1208, "")
        lngDeviceID = zlDatabase.GetPara("�ļ���ȡ����", 100, 1208, -1)
       '27693  �Զ��������������־-��ʾ"���Ͳ�ƥ��"
        'ԭ���룺 strSampleNO = strSampltID
        strSampleNO = IIf(Trim(strSampltID) = "", "-1", strSampltID)
        
        dtStart = Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")
        dtEnd = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
        GetResultFromFile WinsockC, WinsockC.LocalIP, strFile, lngDeviceID, strSampleNO, dtStart, dtEnd
    Else
        strFile = zlDatabase.GetPara("���������ļ�", 100, 1208, "")
        lngDeviceID = zlDatabase.GetPara("�ļ���ȡ����", 100, 1208, -1)
        If Val(zlDatabase.GetPara("�ļ���ȡ��Χ", 100, 1208, 0)) = 0 Then '��ȡ����
            dtStart = Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")
            dtEnd = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
        Else
            dtStart = CDate(Format(zlDatabase.GetPara("�ļ���ȡ��ʼ����", 100, 1208, zlDatabase.Currentdate), "yyyy-mm-dd 00:00:00"))
            dtEnd = CDate(Format(zlDatabase.GetPara("�ļ���ȡ��������", 100, 1208, zlDatabase.Currentdate), "yyyy-mm-dd 23:59:59"))
        End If
        GetResultFromFile WinsockC, WinsockC.LocalIP, strFile, lngDeviceID, -1, dtStart, dtEnd
    End If
    Me.MousePointer = vbDefault
    'ˢ��
    RefreshData
End Sub
Private Sub PrintSetup()
    '��ӡ����
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    lngҽ��ID = Val(rptList.FocusedRow.Record(mCol.ҽ��id).Value)
    lng����ID = Val(rptList.FocusedRow.Record(mCol.����ID).Value)
    
    strSQL = "select ���ͺ� from ����ҽ������ a , ����ҽ����¼ b where b.id = a.ҽ��id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngҽ��ID)
    If rsTmp.EOF = False Then
        lng���ͺ� = Nvl(rsTmp(0))
    End If
    
    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        ReportPrintSet gcnOracle, glngSys, strReportCode, Me
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ReportPrint(ByVal blnPrint As Boolean)
    '���������ӡ
    
    Dim strReportCode As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long, lng���˿���ID As Long, str���� As String
    Dim strSQL As String

    Dim intLoop As Integer
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    lngҽ��ID = Val(rptList.FocusedRow.Record(mCol.ҽ��id).Value)
    lng����ID = Val(rptList.FocusedRow.Record(mCol.����ID).Value)
    lng���ͺ� = Val(rptList.FocusedRow.Record(mCol.���ͺ�).Value)
    lng���˿���ID = Val(rptList.FocusedRow.Record(mCol.���˿���ID).Value)
    str���� = rptList.FocusedRow.Record(mCol.����).Value
    
    '���û��ѡ����Ҿ��˳�
    If InStr("," & mstrPrintDepts & ",", "," & lng���˿���ID & ",") <= 0 And str���� <> "" Then
        Exit Sub
    End If
    
    strSQL = "select ���ͺ� from ����ҽ������ a , ����ҽ����¼ b where b.id = a.ҽ��id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngҽ��ID)
    If rsTmp.EOF = False Then
        lng���ͺ� = Nvl(rsTmp(0))
    End If
    
    If lngҽ��ID = 0 And lng���ͺ� = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord("Select ��� From ZlReports Where ��� Like '%-N'", Me.Caption)
        If Not rsTmp.EOF Then
            strReportCode = rsTmp(0)
            Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "ID=" & mlngKey, IIf(blnPrint, 2, 1))
        End If
        Exit Sub
    End If
    
    blnCurrMoved = rptList.SelectedRows(0).Record.Item(mCol.ת��).Value = "��"
    Call Open_LIS_Report(Me, lngҽ��ID, lng���ͺ�, lng����ID, mlngKey, blnCurrMoved, blnPrint)
    
    
    If blnPrint = True And Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Value = "�Ѽ���" Then
        If mintUnion = 1 Then
            gstrSql = " select id from ����걾��¼ where ҽ��id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҽ��ID)
            Do Until rsTmp.EOF
                strSQL = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
                rsTmp.MoveNext
            Loop
        Else
            strSQL = "ZL_����걾��¼_�걾�ʿ�(" & mlngKey & ",'',1)"
            zlDatabase.ExecuteProcedure strSQL, gstrSysName
        End If
        Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Value = "�Ѵ�ӡ"
        Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon = 8
        Me.rptList.Populate
    End If
    

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub SetParameter()
    Dim blnExec As Boolean
    If frmLisStationPara.ShowPara(Me) Then
        AutoRefresh = True
        blnAutoRefresh = Val(zlDatabase.GetPara("�Զ�ˢ��", 100, 1208, 1))
        blnComm = Val(zlDatabase.GetPara("��������˫��", 100, 1208, 0))
        blnAutoPrint = zlDatabase.GetPara("��˴�ӡ", 100, 1208, 0)
        int��촦��ʽ = Val(zlDatabase.GetPara("��첡����Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
        intԺ�⴦��ʽ = Val(zlDatabase.GetPara("Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
        intסԺ����ʽ = Val(zlDatabase.GetPara("סԺ������Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
        int���ﴦ��ʽ = Val(zlDatabase.GetPara("���ﲡ����Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
        
        blnExec = Val(zlDatabase.GetPara("ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���", 100, 1208, 0))
        frmLabRequest.mMakeNoRule = zlDatabase.GetPara("�걾������ɹ���", 100, 1208, "��  ��")
        mMakeNoRule = zlDatabase.GetPara("�걾������ɹ���", 100, 1208, "��  ��")
        mSendReport = zlDatabase.GetPara("ʹ�ö����������", 100, 1208, 0)
        mstrPrintDepts = zlDatabase.GetPara("ֻ��ָ�����ұ��浥", 100, 1208, "")
        mblnAout = zlDatabase.GetPara("��˺�������һ������걾", 100, 1208, mblnAout)

        
        Call ShowRequest(Not blnExec)
        cboʱ��.Text = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0)
        Me.dtpDate.Value = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
        Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
        mfrmRequest.SetPara
        RefreshData
    End If
End Sub

Private Function RefreshData(Optional blWhere As Boolean) As Boolean
    '����               'ˢ�����mcol(����������)
    '����               '�Ƿ�ʹ��������ѯ
    Dim strSQL As String
    Dim strSQLbak As String
    Dim rsItem As New ADODB.Recordset
    Dim blnMoved As Boolean                                         '�Ƿ��Ƴ�
    Dim blnסԺ���� As Boolean                                      'סԺ����
    Dim bln���ﲡ�� As Boolean                                      '���ﲡ��
    Dim bln�����鵥 As Boolean                                      '�������鵥
    Dim strStart As String                                          '���鿪ʼʱ��
    Dim strEnd As String                                            '�������ʱ��
    Dim Record As ReportRecord                                      '�б��¼��
    Dim Item As ReportRecordItem                                    '�б���ÿһ�ж���
    Dim Rerow As ReportRow                                          '�ж���
    Dim intLoop As Integer                                          'ѭ����ʱ����
    Dim lngloop As Long                                             '
    Dim varFilter As Variant                                        '�����ִ�����
    Dim varUnionFilter As Variant                                   '��ϲ�ѯ
    Dim varItem As Variant                                          '��ϲ�ѯ������
    Dim intAgeBeging As Integer                                     '���俪ʼ
    Dim intAgeEnd As Integer                                        '�������
    Dim lngRow As Long                                              'ˢ��ǰ��¼��ǰ�к�
    Dim strSample As String                                         '�걾���
    Dim lngAdvice As Long                                           'ҽ����
    Dim lngSampleID As Long                                         '����걾ID
    Dim strWhere As String                                          'Ҫ���ӵ�����
    Dim strTable As String                                          'Ҫ���ӵı�
    Dim strDeptID As String                                         '����ID
    Dim strUserMachine  As String                                   '��ǰ�û�����ʹ�õ�����ID
    Dim strTmp As String
    Dim strStartNO As String, strEndNO As String                    '��ʼ�ͽ���NO
    Dim lngRowIndex As Long                                         '������
    Dim lngRowID As Long                                            '��ID
    Dim blnPathPatient As Boolean                                   '�ٴ�·������
    Dim lngUnionItem As Long                                        '�����ĿID

    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    mblnCompelRefresh = True    'ˢ��ʱ����ǿ��ˢ��
    zlCommFun.ShowFlash "���ڶ�ȡ����������ȴ�...", Me
'    Me.stbThis.Panels(2).Text = "���ڶ�ȡ����������ȴ�..."
    Me.MousePointer = 11
            
    On Error GoTo errH
    
    If Not Me.rptList.FocusedRow Is Nothing Then
        lngRow = Me.rptList.FocusedRow.Index
    End If
    
    If cboUnionItem.ListCount > 0 Then
        If (Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) > 0 And rptList.Tag = "") Or _
        (Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) = -1 And rptList.Tag = "") Then
            strTable = " ,����������Ŀ C "
            strWhere = " And a.id = c.�걾ID "
        End If
    End If
    If rptList.Tag <> "" Then
        varFilter = Split(rptList.Tag, ";")
        If varFilter(mFilter.�߼�) <> "" And varFilter(mFilter.�Ƿ�ʹ�ø߼�) = 1 Or _
            InStr(1, varFilter(mFilter.������Ŀ), ",") > 1 Or varFilter(mFilter.ϸ��) <> "" Or varFilter(mFilter.������) <> "" _
            Or varFilter(mFilter.ҩ�����) <> "" Then
            
            strTable = strTable & " ,������ͨ��� G "
            strWhere = strWhere & " And a.id = g.����걾ID "
            
            If varFilter(mFilter.������) <> "" Or varFilter(mFilter.ҩ�����) <> "" Then
                strTable = strTable & "  , ����ҩ����� O "
                strWhere = strWhere & " and g.id = O.ϸ�����ID "
            End If
        End If
    End If
    
    strSQL = "Select /*+ rule */  distinct      Decode(a.�Ƿ���, 1, '', '����ʧ��') As ����," & vbNewLine & _
            "       decode(a.�걾���,1,'����',decode(a.����,1,'����', '')) As ����,Decode(a.����״̬, 1, '������', 2, '�Ѽ���') As ִ��״̬," & vbNewLine & _
            "       Decode(A.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���','����') As �������," & vbNewLine & _
            "       Decode(Sign(Nvl(a.�Ƿ��ʿ�Ʒ, 0)), 0, '��ͨ', 1, '�ʿ�', -1, '�ȶ�') As �걾����," & vbNewLine & _
            "       Decode(a.����id, Null," & vbNewLine & _
            "                 To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
            "                 a.�걾���) As �걾����ʾ,a.�걾���, A.�Һŵ� ," & vbNewLine & _
            "       Decode(A.������Դ, 1, to_char(nvl(a.�����,a.��ʶ��)), 2, to_char(nvl(a.סԺ��,a.��ʶ��)), 3, to_char(nvl(a.NO,a.��ʶ��)), 4, to_char(nvl(a.�����,a.��ʶ��)),to_char(a.��ʶ��)) As ��ʶ��,a.����,a.�Ա�,a.����," & vbNewLine & _
            "       Decode(a.������Դ,2,S.��������,b.��������) as ��������," & vbNewLine & _
            "       a.������ As �������,a.ҽ��ID,a.����ID,'' As ת��,a.Id,a.����ʱ�� ,a.��ӡ����,a.����id," & vbNewLine & _
            "       a.����ʱ��,a.΢����걾,a.������,a.�����,To_Char(A.Ӥ��) As Ӥ��,a.��������,a.�������ID As ��������id," & vbNewLine & _
            "       a.��ҳID,a.������,a.��������,a.���䵥λ,a.�����,a.סԺ��,a.��������,a.�Һŵ�,a.������Ŀ,e.���� as  �������,f.���� as ��������, " & vbNewLine & _
            "       a.�������ID as ���˿���ID,a.����,a.������,a.�걾��̬,a.������,a.����ʱ��,a.�걾���� as ����걾,a.NO,a.������,a.����ʱ��, " & vbNewLine & _
            "       abs(nvl(a.�Ƿ��ʿ�Ʒ,0)) as �ȶԴ���,a.���ʱ��,n.���� as ��������,a.ִ�п���ID,nvl(a.�걾���,0) as �걾���, " & vbNewLine & _
            "       nvl(a.����,0) as ҽ������,nvl(a.�걾���,0) as �걾����,decode(a.���˿���,null,M.����,a.���˿���) as ���˿���, " & vbNewLine & _
            "       a.��������,nvl(r.����״̬,0) as ����״̬,nvl(r.����ID,0) as ���淢��,a.������,a.����ʱ��,b.������λ,p.��Ŀ,p.����,b.������, " & vbNewLine & _
            "       a.���δͨ��,a.������Դ,a.���Ϊ��,nvl(s.·��״̬,0) as �ٴ�·������ ,decode(d.�����Ƿ����,1,'�������','����δ���')  as ������� " & vbNewLine & _
            " From ����걾��¼ a ,���ű� E , �������� f , ������Ϣ b , ���ű� N , ���ű� M, ����ҽ������ R,����ҽ������ p,������ҳ S ,������ˮ�߱걾 D " & strTable & vbNewLine & _
            " Where a.�������ID = E.id(+)  and a.����id=f.id(+) and a.����id =b.����id(+) and b.��ǰ����ID = M.id(+) and " & vbNewLine & _
            " b.��ǰ����id = n.id(+) And a.ҽ��id = R.ҽ��ID(+) and a.ҽ��ID = P.ҽ��ID(+) And p.��Ŀ(+)='��������' " & vbNewLine & _
            " and a.����ID = S.����ID(+) and a.��ҳID = s.��ҳID(+) and  a.id=d.�걾id(+)  " & strWhere
                  
                  
                  
    If mlngDeptID > 0 And rptList.Tag = "" Then
        strSQL = strSQL & " And Instr(To_Char([2]), To_Char(a.ִ�п���id)) > 0 "
        strDeptID = mlngDeptID
    Else
        If InStr(mstrPrivs, "���п���") = 0 Or InStr(mstrPrivs, "�鿴�������ұ���") > 0 Then
            For intLoop = 1 To Me.cboDept.ListCount - 1
                strDeptID = strDeptID & "," & Me.cboDept.ItemData(intLoop)
            Next
            strSQL = strSQL & " And Instr(To_Char([2]), To_Char(a.ִ�п���id)) > 0  "
        End If
    End If
    
    If cboUnionItem.ListCount > 0 Then
        If Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) > 0 And rptList.Tag = "" Then
            strSQL = strSQL & " and c.������ĿId = [16] "
        End If
        
        If Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) = -1 And rptList.Tag = "" Then
            strSQL = strSQL & " and c.������Ŀid is null "
        End If
    End If
    'ʹ�ù����е��������в�ѯ
    If rptList.Tag <> "" Then
        If varFilter(mFilter.����ʱ��) <> "," Then
            strStart = Mid(varFilter(mFilter.����ʱ��), 1, InStr(1, varFilter(mFilter.����ʱ��), ",") - 1)
            strEnd = Mid(varFilter(mFilter.����ʱ��), InStr(1, varFilter(mFilter.����ʱ��), ",") + 1)
            strSQL = strSQL & " And a.����ʱ�� Between [3] And [4] " & vbCrLf
            blnMoved = MovedByDate(CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")))
        Else
            strStart = Now
            strEnd = Now
        End If
        
        If varFilter(mFilter.����) <> "" Then
            strSQL = strSQL & " And a.���� like [5] "
        End If
        
        If varFilter(mFilter.�Ա�) <> "" Then
            strSQL = strSQL & " And a.�Ա� = [6] "
        End If
        
        If varFilter(mFilter.����) <> "," Then
            If InStr(1, varFilter(mFilter.����), ",") = Len(varFilter(mFilter.����)) Then
                strSQL = strSQL & " And a.�������� >= [7] And a.���䵥λ = [20] "
                intAgeBeging = Mid(varFilter(mFilter.����), 1, InStr(1, varFilter(mFilter.����), ",") - 1)
                intAgeEnd = 0
            ElseIf InStr(1, varFilter(mFilter.����), ",") = 1 Then
                strSQL = strSQL & " And a.�������� <= [8] and a.���䵥λ = [20] "
                intAgeBeging = 0
                intAgeEnd = Mid(varFilter(mFilter.����), 2)
            Else
                strSQL = strSQL & " And a.�������� between  [7] And  [8] And a.���䵥λ = [20] "
                intAgeBeging = Mid(varFilter(mFilter.����), 1, InStr(1, varFilter(mFilter.����), ",") - 1)
                intAgeEnd = Mid(varFilter(mFilter.����), InStr(1, varFilter(mFilter.����), ",") + 1)
            End If
        ElseIf varFilter(mFilter.���䵥λ) <> "" Then
            strSQL = strSQL & " and a.���䵥λ = [20] "
        End If
        
        If varFilter(mFilter.�걾��) <> "" Then
            If varFilter(mFilter.�걾��) Like "0*-0*" Then
                varFilter(mFilter.�걾��) = TransSampleNO(varFilter(mFilter.�걾��))
                strSQL = strSQL & " And a.�걾��� = [9] and a.����id is null "
                strStartNO = varFilter(mFilter.�걾��)
            Else
                varFilter(mFilter.�걾��) = Replace(Replace(varFilter(mFilter.�걾��), "��", "~"), "-", "~")
                If InStr(varFilter(mFilter.�걾��), "~") > 0 Then
                    strStartNO = Split(varFilter(mFilter.�걾��), "~")(0)
                    strEndNO = Split(varFilter(mFilter.�걾��), "~")(1)
                    strSQL = strSQL & " And  �걾���  between [9] and [25]  and a.����id is not null "
                Else
                    strStartNO = varFilter(mFilter.�걾��)
                    strSQL = strSQL & " And  �걾��� = [9] and a.����id is not null "
                End If
            End If
        End If

        If varFilter(mFilter.��ʶ��) <> "" Then
            If IsNumeric(varFilter(mFilter.��ʶ��)) Then
                strSQL = strSQL & " and (a.סԺ�� = [10] or a.����� = [10]) "
            Else
                strSQL = strSQL & " and a.no = [10] "
            End If
        End If
        
        If varFilter(mFilter.�������) <> "" Then
            strSQL = strSQL & " And a.�������� = [11] "
        End If
        
        If varFilter(mFilter.������) <> "" Then
            strSQL = strSQL & " And a.������ like [12] "
        End If
        
        If InStr(1, varFilter(mFilter.������Ŀ), ",") > 1 Then
            If Mid(varFilter(mFilter.������Ŀ), InStr(1, varFilter(mFilter.������Ŀ), ",") + 1) = "True" Then
                strSQL = strSQL & " And g.������Ŀid = [13] "
            Else
                strSQL = strSQL & " And g.������ĿID = [13] "
            End If
        End If
        
        If varFilter(mFilter.�ͼ����) <> 0 Then
            strSQL = strSQL & " And a.�������ID = [14] "
        End If
        
        If varFilter(mFilter.�ͼ���) <> "" Then
            strSQL = strSQL & " and a.������ = [15] "
        End If
        
        If varFilter(mFilter.��������) <> 0 Then
            strSQL = strSQL & " and a.����ID = [17] "
        Else
            If InStr(mstrPrivs, "���п���") = 0 Then
                strSQL = strSQL & " and a.����ID in (Select /*+cardinality(a,10)*/ * From Table(Cast(f_Num2list([24]) As zlTools.t_Numlist)) A) "
                strUserMachine = mstrMachineALL
            End If
        End If
        
        '��ϲ�ѯ
        If varFilter(mFilter.�߼�) <> "" And varFilter(mFilter.�Ƿ�ʹ�ø߼�) = 1 Then
            varUnionFilter = Split(varFilter(mFilter.�߼�), ",")
            For intLoop = 0 To UBound(varUnionFilter)
                varItem = Split(varUnionFilter(intLoop), "^")
                
                If intLoop = 0 Then
                    strSQL = strSQL & " And ( g.������ĿId = " & varItem(0) & _
                                IIf(IsNumeric(varItem(3)), " and zl_to_number(������) " & varItem(2) & varItem(3), _
                                " and g.������ " & varItem(2) & " '" & varItem(3) & "'")
                Else
                    strSQL = strSQL & " OR  g.������ĿId = " & varItem(0) & _
                        IIf(IsNumeric(varItem(3)), " and zl_to_number(������) " & varItem(2) & varItem(3), _
                            " and g.������ " & varItem(2) & " '" & varItem(3) & "'")
                End If
                
                If varItem(3) <> "" And varItem(4) <> "" Then
                    strSQL = strSQL & " and g.������ĿId = " & varItem(0) & _
                            IIf(IsNumeric(varItem(5)), " and zl_to_number(������) " & varItem(4) & varItem(5), _
                            " and g.������ " & varItem(4) & " '" & varItem(5) & "'")
                End If
            Next
            strSQL = strSQL & " )"
        End If
        
        If Val(varFilter(mFilter.����ID)) <> 0 Then
            strSQL = strSQL & " And a.����id = [18] "
        End If
        
        If Nvl(varFilter(mFilter.���ݺ�)) <> "" Then
            strSQL = strSQL & " And a.no = [19] "
        End If
        
        If varFilter(mFilter.����) <> "" Or varFilter(mFilter.�Ա�) <> "" Or varFilter(mFilter.����) <> "," Or _
           varFilter(mFilter.��ʶ��) <> "" Or varFilter(mFilter.�������) <> "" Or InStr(1, varFilter(mFilter.������Ŀ), ",") <> 1 _
           Or varFilter(mFilter.�ͼ����) <> 0 Or varFilter(mFilter.�ͼ���) <> "" Or Val(varFilter(mFilter.����ID)) <> 0 Or _
           (varFilter(mFilter.�߼�) <> "" And varFilter(mFilter.�Ƿ�ʹ�ø߼�) = 1) Or varFilter(mFilter.���ݺ�) <> "" Then
           strSQL = strSQL & " And a.����ID is not null "
        End If
        
        If Val(varFilter(mFilter.ϸ��)) <> 0 Then
            strSQL = strSQL & " And g.ϸ��ID = [21] "
        End If
        
        If Val(varFilter(mFilter.������)) <> 0 Then
            strSQL = strSQL & " And O.������ID = [22] "
        End If
        
        If varFilter(mFilter.ҩ�����) <> "" Then
            strSQL = strSQL & " And O.�������  = [23] "
        End If
    Else
        '��ʹ�ù�������ʱ��ʱ�䷶Χ
        strStart = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
        strEnd = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 2)
        
        If strStart = "�Զ���" Then
            strStart = Format(Me.dtpDate.Value, "yyyy-mm-dd 00:00:00")
            strEnd = Format(Me.dtpDateEnd.Value, "yyyy-mm-dd 23:59:59")
        Else
            If strStart = "" Then strStart = GetDateTime("��  ��", 1)
            If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        End If
        
        strSQL = strSQL & " And  a.����ʱ�� Between [3] And [4] "
        
        blnMoved = MovedByDate(CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")))
    End If
                  
    If rptList.Tag = "" Then
        strSQL = strSQL & _
              IIf(mlngMachineID <> 0, IIf(mlngMachineID = -1, "And a.����ID Is NULL ", "AND a.����ID = [1] "), "")
    End If
    
    '����ǰ����Ա���Բ���������
    If rptList.Tag = "" Then
        If mlngMachineID = 0 Then
            If mlngDeptID = 0 Then
                strUserMachine = ""
            Else
                For intLoop = 0 To Me.cboMachine.ListCount - 1
                    strUserMachine = strUserMachine & "," & Me.cboMachine.ItemData(intLoop)
                Next
                strUserMachine = Mid(strUserMachine, 2)
            End If
            If strUserMachine <> "" Then
                strSQL = strSQL & " and f.ID in (Select /*+cardinality(a,10)*/ * From Table(Cast(f_Num2list([24]) As zlTools.t_Numlist)) A) "
            End If
    
        End If
    End If

    
    If blnMoved Then
        strSQLbak = strSQL
        strSQLbak = Replace(strSQLbak, "'' As ת��", "'��' As ת��")
        
        strSQLbak = Replace(strSQLbak, "����걾��¼", "H����걾��¼")
        strSQLbak = Replace(strSQLbak, "������ͨ���", "H������ͨ���")
        strSQLbak = Replace(strSQLbak, "����������Ŀ", "H����������Ŀ")
        strSQL = strSQL & " Union ALL " & strSQLbak
    End If
    
    strSQL = strSQL & " ORDER BY �걾��� "
    
    If cboUnionItem.ListCount > 0 Then
        lngUnionItem = Val(Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex))
    End If
    If rptList.Tag <> "" Then
        Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngMachineID, strDeptID, CDate(Format(strStart, "yyyy-MM-dd HH:mm:ss")), _
                     CDate(Format(strEnd, "yyyy-MM-dd HH:mm:ss")), "%" & CStr(varFilter(mFilter.����)) & "%", CStr(varFilter(mFilter.�Ա�)), Val(intAgeBeging), Val(intAgeEnd), _
                     CStr(strStartNO), UCase(varFilter(mFilter.��ʶ��)), CStr(varFilter(mFilter.�������)), CStr(varFilter(mFilter.������)) & "%", _
                     Mid(varFilter(mFilter.������Ŀ), 1, InStr(1, varFilter(mFilter.������Ŀ), ",") - 1), CStr(varFilter(mFilter.�ͼ����)), CStr(varFilter(mFilter.�ͼ���)), _
                     lngUnionItem, CLng(varFilter(mFilter.��������)), CLng(Val(varFilter(mFilter.����ID))), zlCommFun.GetFullNO(CStr(varFilter(mFilter.���ݺ�))), _
                     CStr(varFilter(mFilter.���䵥λ)), Val(varFilter(mFilter.ϸ��)), Val(varFilter(mFilter.������)), CStr(varFilter(mFilter.ҩ�����)), strUserMachine, strEndNO)
    Else
        Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngMachineID, strDeptID, CDate(Format(strStart, "yyyy-MM-dd HH:mm:ss")), _
                     CDate(Format(strEnd, "yyyy-MM-dd HH:mm:ss")), "", "", "", "", "", "", "", "", "", "", "", _
                     lngUnionItem, 0, 0, "", "", 0, 0, "", strUserMachine, strEndNO)
    End If
    
    'ˢ��ǰ��¼һ��λ��
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For intLoop = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(intLoop).Record.Index
                    lngRowID = Me.rptList.Rows(intLoop).Record(mCol.ID).Value
                End If
            Next
        End If
    End If

    '����������ֻ��ѯһ�κ����
'    Me.rptList.Tag = ""
    Me.rptList.Records.DeleteAll
    blnPathPatient = False
    zlCommFun.ShowFlash "������������..."
    Do Until rsItem.EOF
        
        If lngSampleID <> rsItem("ID") Then
'            Me.stbThis.Panels(2).Text = "������������(" & lngLoop & ")"
            Set Record = Me.rptList.Records.Add
            
            For intLoop = 0 To Me.rptList.Columns.Count + 1
                Record.AddItem ""
            Next
            
            'ǰ�漸����Ҫ����ͼ��
            Record.Item(mCol.����).Value = IIf(Nvl(rsItem("�걾����")) = 1, "����", "")
            If Record.Item(mCol.����).Value = "����" Then
                Record.Item(mCol.����).Icon = 1
            Else
                Record.Item(mCol.����).Icon = -1
            End If
            
            Record.Item(mCol.����ҽ��).Value = IIf(Nvl(rsItem("ҽ������")) = 1, "����", "")
            If Record.Item(mCol.����ҽ��).Value = "����" Then
                Record.Item(mCol.����ҽ��).Icon = 14
            Else
                Record.Item(mCol.����ҽ��).Icon = -1
            End If
            
'            If Nvl(rsItem("ִ��״̬")) = "�Ѽ���" Then
'                Record.Item(mCol.ִ��״̬).Value = "�Ѽ���"
'                Record.Item(mCol.ִ��״̬).Icon = 7
'            ElseIf CInt(Nvl(rsItem("��ӡ����"), "0")) > 0 Then
'                Record.Item(mCol.ִ��״̬).Value = "�Ѵ�ӡ"
'                Record.Item(mCol.ִ��״̬).Icon = 8
'            ElseIf Nvl(rsItem("����")) = "" Then
'                Record.Item(mCol.ִ��״̬).Value = "�Ѵ���"
'                Record.Item(mCol.ִ��״̬).Icon = 6
'            End If
            
'            If Nvl(rsItem("������")) <> "" Then
'                Record.Item(mCol.����״̬).Value = "�ѳ���"
'                Record.Item(mCol.����״̬).Icon = 13
'            End If

                    
            If Nvl(rsItem("����״̬")) = 1 Then
                Record.Item(mCol.����״̬).Value = "�Ѳ���"
                Record.Item(mCol.����״̬).Icon = 11
            End If
                            
                            
                            
            If CInt(Nvl(rsItem("��ӡ����"), "0")) > 0 Then
                Record.Item(mCol.ִ��״̬).Value = "�Ѵ�ӡ"
                Record.Item(mCol.ִ��״̬).Icon = 8
            ElseIf Nvl(rsItem("ִ��״̬")) = "�Ѽ���" Then
                Record.Item(mCol.ִ��״̬).Value = "�Ѽ���"
                Record.Item(mCol.ִ��״̬).Icon = 7
            ElseIf Nvl(rsItem("������")) <> "" Then
                Record.Item(mCol.ִ��״̬).Value = "����"
                Record.Item(mCol.ִ��״̬).Icon = 13
            ElseIf Nvl(rsItem("����")) = "" Then
                Record.Item(mCol.ִ��״̬).Value = "�Ѵ���"
                Record.Item(mCol.ִ��״̬).Icon = 6
            Else
                Record.Item(mCol.ִ��״̬).Value = ""
                Record.Item(mCol.ִ��״̬).Icon = -1
            End If
            
            If Val(Nvl(rsItem("�������"))) > 0 Then
                Record.Item(mCol.����).Icon = 10
            End If
            If rsItem("�������") & "" = "�������" Then
                Record.Item(mCol.�������).Value = "��"
            Else
                Record.Item(mCol.�������).Value = "��"
            End If
            If Val(Nvl(rsItem("�ٴ�·������"))) = 1 Then
                Record.Item(mCol.�ٴ�·������).Icon = 15
                blnPathPatient = True
            Else
                Record.Item(mCol.�ٴ�·������).Icon = -1
            End If
            
            Record.Item(mCol.����).Value = Nvl(rsItem("����")) '& IIf(Nvl(rsItem("Ӥ��"), 0) > 0, "(Ӥ��)", "")
            If Nvl(rsItem("�걾����")) = "�ʿ�" Then
                Record.Item(mCol.�걾����).Value = "�ʿ�"
                Record.Item(mCol.�걾����).Icon = 3
                strSQL = "Select A.�걾id, B.����, B.����, B.ˮƽ From �����ʿؼ�¼ A, �����ʿ�Ʒ B Where A.�ʿ�Ʒid = B.ID And A.�걾id=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsItem("ID"))))
                Do Until rsTmp.EOF
                    Record.Item(mCol.����).Value = "" & rsTmp!���� & "," & rsTmp!���� & ",ˮƽ" & rsTmp!ˮƽ
                    rsTmp.MoveNext
                Loop
            ElseIf Nvl(rsItem("�걾����")) = "�ȶ�" Then
                Record.Item(mCol.�걾����).Value = "�ȶ�"
                Record.Item(mCol.�걾����).Icon = 4
                Record.Item(mCol.����).Value = Record.Item(mCol.����).Value & "(" & Nvl(rsItem("�ȶԴ���")) & ")"
            End If
            
            Record.Item(mCol.�걾��).Value = Val(Nvl(rsItem("�걾���")))
            Record.Item(mCol.�걾��).Caption = Trim(rsItem("�걾����ʾ"))

            If Nvl(rsItem("��������")) = "" Then
                
                If Nvl(rsItem("Ӥ��"), 0) = 0 Then
                    If IsNumeric(Nvl(rsItem("����"))) = True Then
                        Record.Item(mCol.����).Caption = Nvl(rsItem("����")) & "��"
                    Else
                        If Nvl(rsItem("����")) <> "��" And Nvl(rsItem("����")) <> "0��" Then
                            Record.Item(mCol.����).Caption = Nvl(rsItem("����"))
                        End If
                    End If
                    If Record.Item(mCol.����).Caption <> "" Then
                        Record.Item(mCol.����).Value = Val(rsItem("����"))
                    End If
                End If
    '            Record.Item(mCol.����).Caption = IIf(Nvl(rsItem("Ӥ��"), 0) > 0, "", _
                                           IIf(Nvl(rsItem("����")) = "��", "", _
                                           IIf(Nvl(rsItem("����")) = "0��", "", IIf(IsNumeric(Nvl(rsItem("����"))) = True, rsItem("����") & "��", rsItem("����")))))
            Else
                Record.Item(mCol.����).Value = Nvl(rsItem("��������"))
                Record.Item(mCol.����).Caption = Nvl(rsItem("����")) '  Nvl(rsItem("��������")) & Nvl(rsItem("���䵥λ"))
            End If
            If Nvl(rsItem("��������")) <> "" Then
                Record.Item(mCol.����).ForeColor = zlDatabase.GetPatiColor(Nvl(rsItem("��������")), False)
            End If
            Record.Item(mCol.�Ա�).Value = Nvl(rsItem("�Ա�"))
            Record.Item(mCol.�������).Value = Nvl(rsItem("�������"))
            Record.Item(mCol.������Ŀ).Value = Trim(Nvl(rsItem("������Ŀ")))
            Record.Item(mCol.��ʶ��).Value = Nvl(rsItem("��ʶ��"))
            
            Record.Item(mCol.�������).Value = Nvl(rsItem("�������"))
            Record.Item(mCol.ҽ��id).Value = Nvl(rsItem("ҽ��ID"))
            Record.Item(mCol.����id).Value = Nvl(rsItem("����ID"))
            Record.Item(mCol.ת��).Value = Nvl(rsItem("ת��"))
            Record.Item(mCol.����ID).Value = Nvl(rsItem("����id"))
            Record.Item(mCol.ID).Value = Nvl(rsItem("ID"))
            Record.Item(mCol.�걾ʱ��).Caption = Format(Nvl(rsItem("����ʱ��")), "MM-dd HH:mm:ss")
            Record.Item(mCol.�걾ʱ��).Value = Format(Nvl(rsItem("����ʱ��")), "YYYY-MM-dd HH:mm:ss")
            Record.Item(mCol.����ʱ��).Caption = Format(Nvl(rsItem("����ʱ��")), "MM-dd HH:mm")
            Record.Item(mCol.����ʱ��).Value = Format(Nvl(rsItem("����ʱ��")), "YYYY-MM-dd HH:mm")
            Record.Item(mCol.΢����걾).Value = Val(Nvl(rsItem("΢����걾")))
    '        Record.Item(mCol.�շѵ�).Value = Nvl(rsItem("�շѵ�"))
            Record.Item(mCol.�Һŵ�).Value = Nvl(rsItem("�Һŵ�"))
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.�����).Value = Nvl(rsItem("�����"))
            Record.Item(mCol.���˿���).Value = Nvl(rsItem("���˿���"))
            Record.Item(mCol.��������).Value = Nvl(rsItem("��������"))
            'Record.Item(mCol.���ͺ�).Value = Nvl(rsItem("���ͺ�"))
            Record.Item(mCol.Ӥ��).Value = Nvl(rsItem("Ӥ��"))
            Record.Item(mCol.������).Value = Nvl(rsItem("��������"))
            Record.Item(mCol.��ҳID).Value = Nvl(rsItem("��ҳID"))
            Record.Item(mCol.��������ID).Value = Nvl(rsItem("��������Id"))
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.��������).Value = Nvl(rsItem("��������"))
            Record.Item(mCol.���䵥λ).Value = Nvl(rsItem("���䵥λ"))
            Record.Item(mCol.����).Value = Nvl(rsItem("����"))
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.�걾��̬).Value = Nvl(rsItem("�걾��̬"))
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.����ʱ��).Value = Nvl(rsItem("����ʱ��"))
            Record.Item(mCol.����걾).Value = Nvl(rsItem("����걾"))
            Record.Item(mCol.NO).Value = Nvl(rsItem("NO"))
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.����ʱ��).Value = Nvl(rsItem("����ʱ��"))
            Record.Item(mCol.���ʱ��).Value = Nvl(rsItem("���ʱ��"))
            Record.Item(mCol.��������).Value = Nvl(rsItem("��������"))
            Record.Item(mCol.ִ�п���ID).Value = Nvl(rsItem("ִ�п���ID"))
            Record.Item(mCol.�걾���).Value = Nvl(rsItem("�걾���"))
            Record.Item(mCol.ҽ������).Value = Nvl(rsItem("ҽ������"))
            Record.Item(mCol.�걾����).Value = Nvl(rsItem("�걾����"))
            Record.Item(mCol.�������).Value = Nvl(rsItem("�������"))
            Record.Item(mCol.��������).Value = Nvl(rsItem("��������"), 0)
            Record.Item(mCol.���淢��).Value = Nvl(rsItem("���淢��"), 0)
            Record.Item(mCol.���˿���ID).Value = Nvl(rsItem("���˿���ID"), 0)
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.����ʱ��).Value = Nvl(rsItem("����ʱ��"))
            Record.Item(mCol.������).Value = Nvl(rsItem("������"))
            Record.Item(mCol.���δͨ��).Value = Nvl(rsItem("���δͨ��"))
            Record.Item(mCol.������Դ).Value = Nvl(rsItem("������Դ"))
            Record.Item(mCol.�����).Value = Nvl(rsItem("�����"))
            Record.Item(mCol.סԺ��).Value = Nvl(rsItem("סԺ��"))
            If Nvl(rsItem("��Ŀ")) = "��������" Then
                Record.Item(mCol.��λ).Value = Nvl(rsItem("����"))
            End If
            Record.Item(mCol.���Ϊ��).Value = Val(Nvl(rsItem("���Ϊ��")))
            
            
'            Record.Item(mCol.����״̬).Value = Nvl(rsItem("����״̬"), 0)

            
            '------��ú����
            For i = 0 To rptList.Columns.Count + 1
                If Val("" & rsItem!΢����걾) = 0 Then
                    If Record.Item(mCol.���Ϊ��).Value > 0 Then
                        Record.Item(i).BackColor = vbWhite
                    Else
                        Record.Item(i).BackColor = &HFDD6C6
                    End If
                Else
                    Record.Item(i).BackColor = vbWhite
                End If
            Next
            
            lngloop = lngloop + 1
            If mintLoadShow > 0 Then
                DoEvents
            End If
        End If
        lngSampleID = rsItem("ID")
        rsItem.MoveNext
        If lngloop = 10000 Then
            MsgBox "��ѡ���������Χ�����ѳ���10000����¼��" & vbCrLf & _
                   " ������ѡ���������в���!", vbQuestion, Me.Caption
            Call SetControlFocus
            gintSelectFocus = 1
            Exit Do
        End If
    Loop
    
    'û���ٴ�·������ʱ����ʾ��
    Me.rptList.Columns(6).Visible = blnPathPatient
    
    zlCommFun.StopFlash
'    If Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
'    Else
'        Me.rptList1.SetFocus
'    End If
    Me.rptList.Populate
    Me.MousePointer = 0
    
    If mintLoadShow = 0 Then
        mintLoadShow = mintLoadShow + 1
        Exit Function
    End If
    
    '���˽����б�
    RptListFilter
    Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList.Rows.Count & "�����ˡ�"
    
    
    
    
    '���¶�λ����ǰ��λ��
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If
    
        
    For Each Rerow In Me.rptList.Rows
        If Rerow.Record(mCol.ID).Value = mlngKey Then
            Rerow.Selected = True
            mlngKey = Rerow.Record(mCol.ID).Value
            Set Me.rptList.FocusedRow = Rerow
            Me.rptList.Populate
            Exit Function
        End If
    Next
    
    If Me.rptList.Rows.Count > 0 Then
        If lngRow <= Me.rptList.Rows.Count And lngRow > 0 Then
            Set Me.rptList.FocusedRow = rptList.Rows(lngRow - 1)
            mlngKey = rptList.Rows(lngRow - 1).Record(mCol.ID).Value
        Else
            Set Me.rptList.FocusedRow = rptList.Rows(0)
            mlngKey = rptList.Rows(0).Record(mCol.ID).Value
        End If
        Me.rptList.Populate
        Exit Function
    Else
        mlngKey = 0
    End If
    
    
    'ˢ���б�
    If Not Me.rptList.FocusedRow Is Nothing Then
        Call mfrmRequest.zlRefresh(Me.rptList.FocusedRow)
        Call RefreshTableWindow(Me.TabCtlWindow.Selected.Index)
        If mlngKey <> 0 Then
            ReadImageData mlngKey, False
        End If
    End If
    Exit Function
errH:
'    Me.rptList.SetFocus
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub CreaterptListHead()
    Dim Column As ReportColumn
    Dim i As Integer
    With Me.rptList1.Columns
        
        rptList1.AllowColumnRemove = False
        rptList1.ShowItemsInGroups = False
        
        With rptList1.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList1.SetImageList imgList
        Set Column = .Add(mCol.����, "����", 18, False)
        Column.Icon = 0
        Set Column = .Add(mRCol.����ID, "����ID", 45, False): Column.Visible = False: Column.ShowInFieldChooser = False
'        column.
        Set Column = .Add(mRCol.��Դ, "��Դ", 55, True)
        Set Column = .Add(mRCol.����, "����", 55, True)
        Set Column = .Add(mRCol.�Ա�, "�Ա�", 55, True)
        Set Column = .Add(mRCol.����, "����", 55, True)
        Set Column = .Add(mRCol.���˿���, "���˿���", 75, True)
        Set Column = .Add(mRCol.��ʶ��, "��ʶ��", 65, True)
        Set Column = .Add(mRCol.����, "����", 65, True)
        Set Column = .Add(mRCol.ҽ������, "ҽ������", 75, True)
        Set Column = .Add(mRCol.����ҽ��, "����ҽ��", 75, True)
        Set Column = .Add(mRCol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mRCol.ǩ��ʱ��, "ǩ��ʱ��", 75, True)
        Set Column = .Add(mRCol.������ĿID, "������ĿID", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mRCol.ִ��״̬, "ִ��״̬", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mRCol.��λ, "��λ", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mRCol.�Һŵ�, "�Һŵ�", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
    End With
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList.SetImageList imgList
        
        Set Column = .Add(mCol.����, "����", 18, False):   Column.Icon = 0
        Set Column = .Add(mCol.����ҽ��, "����ҽ��", 18, False): Column.Icon = 14
        Set Column = .Add(mCol.ִ��״̬, "ִ��״̬", 18, False): Column.Icon = 5
        Set Column = .Add(mCol.�걾����, "�걾����", 18, False): Column.Icon = 2
        Set Column = .Add(mCol.����, "����", 18, False): Column.Icon = 9
        Set Column = .Add(mCol.����״̬, "����״̬", 18, False): Column.Icon = 11
        Set Column = .Add(mCol.�ٴ�·������, "�ٴ�·������", 18, False): Column.Icon = 15
        Set Column = .Add(mCol.�������, "�������", 18, False): Column.Icon = 16
        
        Set Column = .Add(mCol.�걾��, "�걾��", 65, True)
        Column.SortAscending = zlDatabase.GetPara("�걾������", 100, 1208, 0)
        Column.Sortable = True:  Me.rptList.SortOrder.Add Column
        Set Column = .Add(mCol.����, "����", 45, True)
        Set Column = .Add(mCol.�Ա�, "�Ա�", 40, True)
        Set Column = .Add(mCol.����, "����", 40, True)
        Set Column = .Add(mCol.�걾ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mCol.������Ŀ, "������Ŀ", 90, True)
        Set Column = .Add(mCol.�������, "��Դ", 40, False)
        Set Column = .Add(mCol.��ʶ��, "��ʶ��", 55, True)
        Set Column = .Add(mCol.��������, "��������", 75, True)
        Set Column = .Add(mCol.������, "������", 75, True)
        Set Column = .Add(mCol.��λ, "��λ", 75, True)
        Set Column = .Add(mCol.���˿���, "���˿���", 80, True) ': Column.Visible = False
        Set Column = .Add(mCol.�������, "�������", 65, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.ҽ��id, "ҽ��ID", 65, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����id, "����ID", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.ת��, "ת��", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����ID, "����ID", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.ID, "ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.΢����걾, "΢����걾", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�շѵ�, "�շѵ�", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�Һŵ�, "�Һŵ�", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.������, "������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�����, "�����", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���ͺ�, "���ͺ�", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.Ӥ��, "Ӥ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.��������ID, "��������ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.��ҳID, "��ҳID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.������, "������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.��������, "��������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���䵥λ, "���䵥λ", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����, "����", 30, True) ': Column.Visible = False
        Set Column = .Add(mCol.������, "������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�걾��̬, "�걾��̬", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.������, "������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����걾, "����걾", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.NO, "NO", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.������, "������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���ʱ��, "���ʱ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����id, "����ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.��������, "����", 30, True) ': Column.Visible = False
        Set Column = .Add(mCol.��λ, "��λ", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.ִ�п���ID, "ִ�п���ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�걾���, "�걾���", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.ҽ������, "ҽ������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�걾����, "�걾����", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�������, "�������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.��������, "��������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���淢��, "���淢��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���˿���ID, "���˿���ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.������, "������", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���δͨ��, "���δͨ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.������Դ, "������Դ", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.�����, "�����", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.סԺ��, "סԺ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.���Ϊ��, "���Ϊ��", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
    End With
End Sub

Private Sub SampleDisposal(Disposal As Integer)
    '����           ����������ĸ��ֲ���
    '               �������޸ġ�ɾ�������걾
    Dim strSQL As String                                    '��ʱSQL���
    Dim intExeState As Integer                              'ִ��״̬( 7=�Ѽ��� 8=������ӡ 6=�ѷ������ݸ�����
    Dim strSamptleType As String                            '�걾�������(Ժ�ڡ�Ժ�⡢����������)
    Dim strSamptleKind  As Integer                          '�걾����(=3�ʿ� =4�ȶ�)
    Dim strPatienName As String                             '��������
    Dim lngMachineID As Long                                '����ID
    Dim strSampltDate As String                             '�걾ʱ��
    Dim strSampltID As String                               '�걾��
    Dim blEmergent  As Boolean                              '�Ƿ����
    Dim lngAdvice As Long                                   'ҽ��ID
    Dim lngRetuId As Long                                   '����ģ�鷵�ص�ID
    Dim rsTmp As New ADODB.Recordset                        '���ݼ�
    Dim rs As New ADODB.Recordset
    Dim intMicrobe As Integer                               '�Ƿ���΢����
    Dim strVerifyMan As String                              '������
    Dim lngPatientID As Long                                '����ID
    Dim strDevices As String                                '�豸ID��
    Dim strAdviceIDs As String                              'ҽ��ID��
    Dim aDevice() As String                                 '�豸S
    Dim intLoop As Integer
    Dim astrSQL() As String                                 'SQL����
    Dim intEmerge As Integer                                '�Ƿ�ʹ�ü����־
    Dim lngSampleID As Long                                 '�걾ID
    Dim lngBeginDate As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim blnRollBak As Boolean                               '���˱�־
    Dim strEmergency As String                              '�걾���
    Dim str������ As String                                  '������
    Dim strAdviceIDall As String                            'ҽ��id ,�����°�LIS���� ״̬
    
    Dim bln���ͱ��� As Boolean                              '�Ƿ��ͱ���
    Dim intRow As Integer, strIDList() As String              '����걾ID
    Dim strNoSend As String, lngCount As Long
    On Error GoTo errH
    
    intEmerge = Val(zlDatabase.GetPara("����걾", 100, 1208, 0))
    
    If Not rptList.FocusedRow Is Nothing Then                                   'û�н�����ʱ�˳�
        With Me.rptList.FocusedRow
            intExeState = .Record(mCol.ִ��״̬).Icon
            strSamptleType = .Record(mCol.�������).Value
            strPatienName = .Record(mCol.����).Value
            lngMachineID = Val(.Record(mCol.����id).Value)
            strSampltDate = .Record(mCol.�걾ʱ��).Value
            strSampltID = .Record(mCol.�걾��).Value
            blEmergent = IIf(.Record(mCol.�걾���).Value = "1", True, False)
            lngAdvice = Val(.Record(mCol.ҽ��id).Value)
            strSamptleKind = .Record(mCol.�걾����).Icon
            intMicrobe = Val(.Record(mCol.΢����걾).Value)
            strVerifyMan = .Record(mCol.������).Value
            lngSampleID = .Record(mCol.ID).Value
            strEmergency = .Record(mCol.�걾���).Value
            str������ = .Record(mCol.������).Value
        End With
    End If
    
    '�õ�����Id(�ȴ����б�)
    If Me.TabList.Item(1).Selected = True Then
        If Not Me.rptList1.FocusedRow Is Nothing Then
            With Me.rptList1.FocusedRow
                lngPatientID = Val(.Record(mRCol.����ID).Value)
            End With
        End If
    End If
    
    Select Case Disposal
        Case mActS.�޸�������                                                           '�޸�������
            
            Dim strNewNo As String, str�걾��̬ As String, str�걾���� As String, str���� As String
            '�Ѽ�����Ŀ���ܽ����޸������Ų���
            If intExeState = 7 Or intExeState = 8 Then Exit Sub
            
            frmLisStationModifyNo.ShowEdit Me, mlngKey, strNewNo, str�걾��̬, str�걾����, strEmergency, str����
            If strNewNo = "" Then Exit Sub
            '�жϱ걾�Ƿ����
            strStartDate = GetDateTime(mMakeNoRule, 1, strSampltDate)
            strEndDate = GetDateTime(mMakeNoRule, 2, strSampltDate)
            gstrSql = "Select ID" & vbNewLine & _
                    " From ����걾��¼ A" & vbNewLine & _
                    " Where ����ʱ�� Between [1] And [2] And �걾��� = [3] And Nvl(�걾���, 0) = [4] And ID <> [5] " & vbNewLine & _
                    "       And ����ID = [6] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strStartDate), CDate(strEndDate), _
                    TransSampleNO(strNewNo), IIf(strEmergency, 1, 0), mlngKey, mlngMachineID)
                    
            If rsTmp.EOF = False Then
                MsgBox "�걾��<" & strNewNo & ">�Ѵ��ڣ�", vbInformation, Me.Caption
                Call SetControlFocus
                Exit Sub        '�ҵ���ͬʱ�˳�
            End If
            If strNewNo <> "" Then
                strSQL = "ZL_����걾��¼_�걾���(" & _
                         mlngKey & ",'" & strNewNo & "','" & str�걾��̬ & "','" & str�걾���� & "',NULL,NULL,'" & strEmergency & "','" & str���� & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
'            InsertOneRecored mlngKey, False
            RefreshData
            gintSelectFocus = 1
'            RefreshData
        Case mActS.�����޸�������                                                       '�����޸�������
            
            Call frmBatchAction.ShowMe(Me, 4, mlngMachineID, , , , , mlngDeptID, mstrAuditingManID)
            gintSelectFocus = 1
        Case mActS.ɾ�������걾                                                         'ɾ�������걾
            
            If InStr(";" & mstrPrivs & ";", ";ɾ�������걾;") = 0 Then
                MsgBox "��û��ɾ�������걾��Ȩ�ޣ��������ϵͳ!", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If strSamptleType = "����" Then
                If MsgBox("���Ҫɾ�������걾��", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                End If
                strSQL = "ZL_����걾��¼_�걾ɾ��(" & mlngKey & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            If strSamptleType = "Ժ��" Then
                'ȡ������
                If MsgBox("���Ҫɾ����" & strPatienName & "��Ժ��걾��", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                End If
                Call SampleRefuse(mlngKey)                               'ȡ������
                'ɾ������
                strSQL = "ZL_����걾��¼_�걾ɾ��(" & mlngKey & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            intLoop = Me.rptList.FocusedRow.Index
            DelItem lngSampleID
            On Error Resume Next
            If Me.rptList.Rows.Count > 0 Then
                If Me.rptList.Rows.Count < intLoop Then
                    Me.rptList.FocusedRow = Me.rptList.Rows(Me.rptList.Rows.Count)
                Else
                    If intLoop = 0 Then intLoop = 1
                    Me.rptList.FocusedRow = Me.rptList.Rows(intLoop - 1)
                End If
            End If
            gintSelectFocus = 1
'            With Me.rptList.FocusedRow
'                .Record.DeleteAll
'            End With
'            Me.rptList.Populate
'            RefreshData
'            Me.rptList.SetFocus
        Case mActS.��������                                                             '��������
'            If mbln�ֹ����ͱ��� Then
'                'strSQL = "Select nvl(����ʱָ������,0) as ���� From �������� Where id=[1]"
'                'Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngMachineID)
'                   'Do Until rs.EOF
'                     bln���ͱ��� = Val("" & rs!����) = 1
'                     rs.MoveNext
'                    Loop
'                End If
                If blnComm = False Then Exit Sub        '����Ҫʱֱ���˳�
                
282             SendSample WinsockC, WinsockC.LocalIP, lngMachineID, strSampltDate, strSampltID, "", False, IIf(blEmergent And intEmerge = 1, 1, 0)
                        '�ɹ�
284             If blnComm And Not Me.rptList.FocusedRow Is Nothing Then
286                 Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Value = "�Ѵ���"
288                 Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon = 6
290                 Me.rptList.Populate
                End If
        Case mActS.�������͵�����
                
244             'If blnComm = False Then Exit Sub
                Call RefreshData 'ˢ��״̬���ٷ���
                Me.MousePointer = vbHourglass
                    '2013-11-26 ֧�ֳ���1000�����ϵı걾����
246             strNoSend = "": lngCount = 0
248             ReDim strIDList(0) As String
250             intRow = -1
252             If rptList.Rows.Count > 0 Then
254                 For intRow = 0 To rptList.Rows.Count - 1
256                     With rptList
258                         If Not .Rows(intRow).GroupRow Then
                                '7-�Ѽ��� 8-�Ѵ�ӡ 13-����
                            
260                             If InStr(",7,8,13,", CStr(.Rows(intRow).Record(mCol.ִ��״̬).Icon)) <= 0 Then
262                                 strIDList(UBound(strIDList)) = strIDList(UBound(strIDList)) & "," & .Rows(intRow).Record(mCol.ID).Value
264                                 If Len(strIDList(UBound(strIDList))) > 3000 Then ReDim Preserve strIDList(UBound(strIDList) + 1)
                                Else
268                                 strNoSend = strNoSend & vbNewLine & .Rows(intRow).Record(mCol.�걾��).Value & " " & .Rows(intRow).Record(mCol.����).Value & " ״̬=" & CStr(.Rows(intRow).Record(mCol.ִ��״̬).Icon)
270                                 lngCount = lngCount + 1
                                End If
                            End If
                        End With
                    Next
                End If
272             If intRow >= 0 Then Call frmLabMainSendSample.ShowMe(strIDList(), Me)
274             If strNoSend <> "" Then
276                 stbThis.Panels(2).Text = "��" & lngCount & "���걾��������ˣ����β���������"
'278                 WriteToLog "δ�������ͽ���ı걾�У�" & strNoSend
                End If

292             Me.MousePointer = vbDefault
                '��������ʱ��ˢ��
    '            If Me.rptList.Tag <> "Continue" Then
    '                RefreshData
    '            End If
            
        Case mActS.��Ϊ�ʿ�                                                             '��Ϊ�ʿ�
            
            '����
            frmLabMainSetQC.ShowMe Me, mlngKey, strSampltID, lngMachineID, strVerifyMan, 1
            
            InsertOneRecored mlngKey, False
'            If MsgBox("���Ҫ�������걾תΪ�ʿر걾��", _
'                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            strSql = "ZL_����걾��¼_�걾�ʿ�(" & mlngKey & ",1)"
'            zldatabase.ExecuteProcedure strSql, Me.Caption
            
'            RefreshData
            gintSelectFocus = 1
            InsertOneRecored mlngKey, False
            
        Case mActS.��Ϊ�Ա�                                                             '��Ϊ�Ա�
            
            frmLabToCompare.ShowMe Me, mlngKey
            InsertOneRecored mlngKey, False
'            RefreshData
            gintSelectFocus = 1
        Case mActS.״̬�ع�                                                             '״̬�ع�
            '�Ƿ�ֻ�ܻع����ѵı걾
            If InStr(1, mstrPrivs, "�޸����˽��") <= 0 And UserInfo.���� <> strVerifyMan And strPatienName <> "" Then
                MsgBox "�㲻�ܻع����˵ı��浥��", vbInformation, Me.Caption
                Call SetControlFocus
                Exit Sub
            End If
            
            If intExeState = 7 Or intExeState = 8 Or intExeState = 11 Then
                '��˺�״̬
                If strSamptleKind = 4 Then
                    'ȡ���ȶ�
                    If MsgBox("���Ҫ����" & strPatienName & "���ȶԱ걾תΪ��ͨ�걾��", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                    End If
                    strSQL = "ZL_����걾��¼_�걾�ʿ�(" & mlngKey & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    InsertOneRecored mlngKey, False
                Else
                    '�ع����
                    If InStr(1, ";" & mstrPrivs & ";", ";���ȡ��;") > 0 Or InStr(1, ";" & mstrPrivs & ";", ";24Сʱ���ȡ��;") > 0 Then
                        Call ReportDisposal(mActR.���ȡ��)
                    Else
                        MsgBox "��û�����ȡ����Ȩ��!", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            Else
                '���ǰ
                If strSamptleKind = 4 Then
                    'ȡ���ȶ�
                    If MsgBox("���Ҫ����" & strPatienName & "���ȶԱ걾תΪ��ͨ�걾��", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
                    End If
                    strSQL = "ZL_����걾��¼_�걾�ʿ�(" & mlngKey & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    InsertOneRecored mlngKey, False
                ElseIf strSamptleKind = 3 Then
                    'ȡ���ʿ�
                    If MsgBox("���Ҫ���ʿر걾תΪ��ͨ�걾��", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
                    End If
'                    frmLabMainSetQC.ShowMe Me, mlngKey, strSampltID, lngMachineID, strVerifyMan, 3
                    strSQL = "ZL_�����ʿؼ�¼_EDIT(3," & mlngKey & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    Me.rptList.FocusedRow.Record(mCol.����).Value = ""
                    Me.rptList.FocusedRow.Record(mCol.�걾����).Value = ""
                    Me.rptList.FocusedRow.Record(mCol.�걾����).Icon = -1
                    Me.rptList.Populate
                ElseIf mSendReport = 1 And str������ <> "" Then
                    gstrSql = "Zl_����걾��¼_���󱨸�(" & mlngKey & ",2,'" & UserInfo.���� & "')"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    InsertOneRecored mlngKey, False
                Else
                    If strPatienName <> "" Then
                        If MsgBox("�Ƿ�ȷ��Ҫ��Ϊ�����걾?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                        End If
                        strSQL = "Select Distinct ҽ��ID From (Select ҽ��ID From ������Ŀ�ֲ� Where �걾id = [1] " & _
                                "Union All Select ҽ��ID From ����걾��¼ Where ID = [1])"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)

                        intLoop = 0
                        Do While Not rsTmp.EOF
                            If Not IsNull(rsTmp(0)) And Val(Nvl(rsTmp(0))) <> 0 Then
                                '����˫��ͨ��
                                If blnComm Then
                                    strAdviceIDs = strAdviceIDs & "," & rsTmp(0)
                                    gstrSql = "Select Distinct ����ID From ����걾��¼ A,������Ŀ�ֲ� B " & _
                                        " Where B.ҽ��ID=[1] And B.�걾ID+0=A.ID"
                                    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsTmp(0)))

                                    Do While Not rs.EOF
                                        If InStr(strDevices, "," & zlCommFun.Nvl(rs(0), 0)) = 0 Then
                                            strDevices = strDevices & "," & zlCommFun.Nvl(rs(0), 0)
                                        End If
                                        rs.MoveNext
                                    Loop
                                End If
                                If intLoop = 0 Then
                                    ReDim Preserve astrSQL(1 To 1)
                                    astrSQL(1) = "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")"
                                Else
                                    ReDim Preserve astrSQL(1 To UBound(astrSQL) + 1)
                                    astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")"
'                                    aStrSQL(ReDimArray(aStrSQL)) = "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")"
                                End If
                                strAdviceIDall = strAdviceIDall & "," & rsTmp(0)
'                                zldatabase.ExecuteProcedure "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")", gstrSysName
                                intLoop = intLoop + 1
                            End If
                            rsTmp.MoveNext
                        Loop
                        
                        If intLoop > 0 Then
                            '����˫��ͨ��
                            If blnComm Then
                                If Len(strDevices) > 0 Then strDevices = Mid(strDevices, 2)
                                If Len(strAdviceIDs) > 0 Then strAdviceIDs = Mid(strAdviceIDs, 2)
                                aDevice = Split(strDevices, ",")
                                mblnSendComplete = False
                                For intLoop = 0 To UBound(aDevice)
                                    SendSample WinsockC, WinsockC.LocalIP, CLng(Val(aDevice(intLoop))), "", 0, strAdviceIDs, True, IIf(blEmergent And intEmerge = 1, 1, 0)
                                Next
                                lngBeginDate = Timer
                                Do
                                    DoEvents
                                Loop Until mblnSendComplete = True Or (CLng(Timer) - lngBeginDate > 2)
                            End If
                            gcnOracle.BeginTrans
                            blnRollBak = True
                            For intLoop = 1 To UBound(astrSQL)
                                If astrSQL(intLoop) <> "" Then Call zlDatabase.ExecuteProcedure(astrSQL(intLoop), Me.Caption)
                            Next
                            gcnOracle.CommitTrans
                        Else
                            gstrSql = "Zl_����걾��¼_��Ϊ����(" & mlngKey & ")"
                            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                        End If
                        If strAdviceIDall <> "" Then
                            ModifyApplyToLIS strAdviceIDall, 0
                        End If
                        Call RefreshData
'                        InsertOneRecored mlngkey, False
                    Else
                        If InStr(";" & mstrPrivs & ";", ";ɾ�������걾;") = 0 Then
                            MsgBox "��û��ɾ�������걾��Ȩ�ޣ��������ϵͳ!", vbInformation, Me.Caption
                            Exit Sub
                        End If
                        
                        If MsgBox("�Ƿ�ȷ��Ҫɾ�������걾?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                        End If
                            strSQL = "ZL_����걾��¼_�걾ɾ��(" & mlngKey & ")"
                            zlDatabase.ExecuteProcedure strSQL, Me.Caption
                        intLoop = Me.rptList.FocusedRow.Index
                        DelItem lngSampleID
                        On Error Resume Next
                        If Me.rptList.Rows.Count > 0 Then
                            If Me.rptList.Rows.Count < intLoop Then
                                Me.rptList.FocusedRow = Me.rptList.Rows(Me.rptList.Rows.Count)
                            Else
                                If intLoop = 0 Then intLoop = 1
                                Me.rptList.FocusedRow = Me.rptList.Rows(intLoop - 1)
                            End If
                        End If
'                        Me.rptList.FocusedRow.Record.DeleteAll
'                        Me.rptList.Populate
                    End If
                    'ȡ������
'                    Call frmLisStationCheckCancel.ShowEdit(Me, mlngKey, objLISComm)
                End If
            End If
            gintSelectFocus = 1
            RptListFilter
'            RefreshData
        Case mActS.����                                                                 '����
            
            If mintEditState <> 0 Then
                lngRetuId = mfrmRequest.ZlRefuse()
                If lngRetuId = 0 Then Exit Sub
                If mintContinue = 0 Then
                    '����������
                    mintEditState = 0
                End If
            Else
                If rptList1.FocusedRow Is Nothing Then Exit Sub
                frmLabRefuse.ShowEdit Me, rptList1.FocusedRow.Record(mRCol.ҽ��id).Value, Me.WinsockC
                RefreshData1
            End If
            gintSelectFocus = 1
        Case mActS.����                                                                 '����
            '�ڱ�ǰ����б����������˼���Ƿ�Ϊͬһ�������
            If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                If AuditionCheck = False Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٺ���.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            '���ڻس�ʱ������ܵ�"���鱸ע"����ʹ�����淽�����ԭ����(����)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
''                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
''                mfrmWrite2.txtComment.SetFocus
'            End If
            Call ShowRequest(True)
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            mintEditState = 1
            If mfrmRequest.ZlEditStart(mActS.����, mlngDeptID, mlngMachineID, 0, 0, 0, _
                                        IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan), _
                                        Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex), _
                                        lngPatientID) = False Then
                mfrmWrite.ZlClearForm
                mfrmWrite2.ZlClearForm
                mintEditState = 0
            End If
            
            
        
        Case mActS.�Ǽ�                                                                 '�Ǽ�
            '�ڱ�ǰ����б����������˼���Ƿ�Ϊͬһ�������
            If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                If AuditionCheck = False Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            '���ڻس�ʱ������ܵ�"���鱸ע"����ʹ�����淽�����ԭ����(����)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
'                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
'                mfrmWrite2.txtComment.SetFocus
'            End If
            Call ShowRequest(True)
            mintEditState = 2
            Me.dkpMain.FindPane(Dkp_ID_Request).Select
            
            If mfrmRequest.ZlEditStart(mActS.�Ǽ�, mlngDeptID, mlngMachineID, 0, 0, 0, _
                                        IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan), _
                                        Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex)) = False Then
                mfrmWrite.ZlClearForm
                mfrmWrite2.ZlClearForm
                mintEditState = 0
            End If
        Case mActS.��������                                                             '��������
            
            If frmAddSample.ShowEdit(Me, "", mlngDeptID, mlngMachineID) = True Then
                
                Call RefreshData
                gintSelectFocus = 1
            End If
        
        Case mActS.�����                                                             '�����
            If intExeState = 7 Or intExeState = 8 Then Exit Sub
            If str������ <> "" And mSendReport = 1 Then Exit Sub
            If strPatienName <> "" Then
                '�Ƿ�ֻ�ܻع����ѵı걾
                If InStr(1, mstrPrivs, "�޸����˽��") <= 0 And UserInfo.���� <> strVerifyMan Then
                    MsgBox "�㲻�ܻع����˵ı��浥��", vbInformation, Me.Caption
                    Call SetControlFocus
                    Exit Sub
                End If
            End If
            
            '�ڱ�ǰ����б����������˼���Ƿ�Ϊͬһ�������
            If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                If AuditionCheck = False Then
                    '����˵�½
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٲ����.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            '���ڻس�ʱ������ܵ�"���鱸ע"����ʹ�����淽�����ԭ����(����)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
'                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
'                mfrmWrite2.txtComment.SetFocus
'            End If
            Call ShowRequest(True)
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            mintEditState = 4
            If mfrmRequest.ZlEditStart(mActS.�����, mlngDeptID, mlngMachineID, mlngKey, 0, 0, _
                                    IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan), _
                                    Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex)) = False Then
                mintEditState = 0
            End If
            
        Case mActS.���º���                                                             '���º���
        
            '���ڻس�ʱ������ܵ�"���鱸ע"����ʹ�����淽�����ԭ����(����)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
'                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
'                mfrmWrite2.txtComment.SetFocus
'            End If
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            Me.dkpMain.FindPane(Dkp_ID_Request).Select
            mintEditState = 3
            If mfrmRequest.ZlEditStart(mActS.���º���, mlngDeptID, mlngMachineID, mlngKey, 0, 0, _
                                    IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan)) = False Then
                mintEditState = 0
            End If
        Case mActS.��Ϊ����
'            strsql = "Select Distinct ҽ��ID From (Select ҽ��ID From ������Ŀ�ֲ� Where �걾id = [1] " & _
'                    "Union All Select ҽ��ID From ����걾��¼ Where ID = [1])"
'            Set rsTmp =zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngkey)
'            gcnOracle.BeginTrans
'            Do While Not rsTmp.EOF
'                If Not IsNull(rsTmp(0)) Then
'                    zlDatabase.ExecuteProcedure "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")", gstrSysName
'                End If
'                rsTmp.MoveNext
'            Loop
'            gcnOracle.CommitTrans
            'ȡ������
            '�Ƿ�ֻ�ܻع����ѵı걾
            If InStr(1, mstrPrivs, "�޸����˽��") <= 0 And UserInfo.���� <> strVerifyMan Then
                MsgBox "�㲻�ܻع����˵ı��浥��", vbInformation, Me.Caption
                Call SetControlFocus
                Exit Sub
            End If
            
            frmLisStationCheckCancel.ShowEdit Me, mlngKey, Me.WinsockC, False, True
'            InsertOneRecored mlngKey, False
            Call RefreshData
            gintSelectFocus = 1
        Case mActS.�ϲ��걾
            If Not rptList.FocusedRow Is Nothing Then                                   'û�н�����ʱ�˳�
                
                With Me.rptList.FocusedRow
                    Call mfrmLabMainSampleUnion.zlRefresh(mlngKey, Nvl(.Record(mCol.����).Value), Nvl(.Record(mCol.�Ա�).Value), Nvl(.Record(mCol.����).Value), _
                                                    Nvl(.Record(mCol.������Ŀ).Value), Nvl(.Record(mCol.�걾��).Value), Nvl(.Record(mCol.������).Value), _
                                                    Nvl(.Record(mCol.����ʱ��).Value), Nvl(.Record(mCol.������).Value))
                End With
            End If
        Case mActS.�ϲ��걾����
            mfrmLabMainSampleUnion.ZlSave
            Call RefreshData
        Case mActS.�޸Ĳ�����Ϣ
'            Call ModifyPatientBaseInfo(mlngKey)
            
    End Select
    
    
    '���˽����б�
'    RptListFilter
'    gintSelectFocus = 1
    If Me.rptList.Rows.Count = 0 And mintEditState = 0 Then
        mfrmRequest.ZlCancel
        mfrmWrite.ZlCancel
        mfrmWrite2.zlRefresh -1
    End If
    Exit Sub
errH:
    AutoRefresh = True                                                      '�������(���Խ���ˢ��)
    If blnRollBak = True Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ModifyPatientBaseInfo(lngKey As Long)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                               �޸Ĳ��˻�����Ϣ
    '����    lngKey                     �걾ID
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim strsql As String
'    Dim rsTmp As New ADODB.Recordset
'    strsql = "Select B.����id, B.��ҳid, Decode(b.������Դ, 2, 1, 0) ���� From ����걾��¼ A, ����ҽ����¼ B Where A.ҽ��id = B.Id and a.id = [1]"
'    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lngKey)
'    If rsTmp.EOF = True Then
'        MsgBox "û���ҵ���Ӧ��ҽ������Ҫ�޸ģ�", vbInformation, Me.Caption
'        Exit Sub
'    End If
'    Call zldatabase.zlModiPatiBaseInfo(Val(rsTmp("����id") & ""), Val(rsTmp("��ҳID") & ""), "���鱨��", rsTmp("����"))
'    Call InsertOneRecored(lngKey)
End Sub

Private Sub ReportDisposal(Disposal As Integer)
    '����                   �Ա���ĸ��ֲ���
    '                       ������������
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strPatienName As String                             '��������
    Dim strSaveAs As String                                 '�Ƿ���
    Dim intMicrobe As Integer                               '�Ƿ���΢����
    Dim strVerifyMan As String                              '������
    Dim strAuditingMan As String                            '�����
    Dim blIf As Boolean                                     '��ʱ��¼�ж�
    Dim strAuditingMain                                     '�����
    Dim strVerifydate As String                             '����ʱ��
    Dim strAuditingDate As String                           '���ʱ��
    Dim lngSampleID As Long                                 '�걾ID
    Dim lngPatientType As Integer                           '������Դ
    Dim lngPatientID  As Long                               '����ID
    Dim intBay As Integer                                   'Ӥ��
    Dim lngApplyDept As Long                                '��������ID
    Dim lngAdviceID As Long                                 'ҽ��ID
    Dim intRepotrCount As Integer                           '����������
    Dim lngPatientPage As Integer                           '��ҳID
    Dim strErrInfo As String                                '������ʾ
    Dim intPrivacy As Integer                               '���ͱ��浥��ҽ��վʱ�Ƿ���ʾ��˽��Ŀ
    Dim lngAdvice As Long                                   'ҽ��ID
    Dim intUnion As Integer                                 '�Ƿ���������������ʾ
    Dim blnClueTo As Boolean                                '�Ƿ���ʾ��˶Ի���
    Dim intLook As Integer                                  'ҽ��վ�Ƿ�鿴����
    Dim strSource As String                                 'ȡ�õ��Ӳ���ǩ���ִ�
    Dim lng֤��ID As Long                                   '֤��ID
    Dim strSign As String                                   'ǩ�������ɵ��ִ�
    Dim strTimeStamp As String                              'ʱ���
    Dim blnRollBack As Boolean                              '�Ƿ�ع�
    Dim str������ As String                                 '������
    Dim intLoop As Integer
    Dim strTmp As String
    Dim lngRow As Long
    Dim astrSQL() As String
    On Error GoTo errH
    ReDim astrSQL(0)
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            strPatienName = .Record(mCol.����).Value
            strSaveAs = .Record(mCol.ת��).Value
            intMicrobe = Val(.Record(mCol.΢����걾).Value)
            strVerifyMan = .Record(mCol.������).Value
            If .Record(mCol.�������).Value = "����" Then
                lngPatientType = 1
            ElseIf .Record(mCol.�������).Value = "סԺ" Then
                lngPatientType = 2
            ElseIf .Record(mCol.�������).Value = "Ժ��" Then
                lngPatientType = 3
            ElseIf .Record(mCol.�������).Value = "���" Then
                lngPatientType = 4
            End If
            lngPatientID = Val(.Record(mCol.����ID).Value)
            intBay = Val(.Record(mCol.Ӥ��).Value)
            lngApplyDept = Val(.Record(mCol.��������ID).Value)
            lngAdviceID = Val(.Record(mCol.ҽ��id).Value)
            intRepotrCount = Val(.Record(mCol.������).Value)
            lngPatientPage = Val(.Record(mCol.��ҳID).Value)
            lngSampleID = Val(.Record(mCol.ID).Value)
            strAuditingDate = .Record(mCol.���ʱ��).Value
            lngAdvice = Val(.Record(mCol.ҽ��id).Value)
            intLook = IIf(.Record(mCol.����״̬).Value = "�Ѳ���", 1, 0)
            str������ = .Record(mCol.������).Value
        End With

    End If
    Select Case Disposal
    
        Case mActR.������������ ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Call frmLisStationAdjust.ShowEdit(Me, mlngDeptID, mstrPrivs)
            RefreshData
            gintSelectFocus = 1
        Case mActR.��˱��� ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If strPatienName = "" Then Exit Sub
            If Me.rptList.FocusedRow.Record(mCol.������).Value <> "" Then
                If Me.rptList.FocusedRow.Record(mCol.����ʱ��).Value <> "" Then
                    If CDate(Me.rptList.FocusedRow.Record(mCol.����ʱ��).Value) > zlDatabase.Currentdate Then
                        MsgBox "����ʱ�䣬���ڵ�ǰʱ�䣬���ܽ�����ˣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
            If InStr(1, mstrPrivs, "��˱걾") <= 0 Then
                'û��Ȩ�޺������û���½ʱ�˳�
                MsgBox "��û��Ȩ�޽������,�����µ�½���������Ա�������!", vbInformation, gstrSysName
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            blIf = False

            If (strVerifyMan = mstrAuditingMan Or (mstrAuditingMan = "" And strVerifyMan = UserInfo.����)) And InStr(1, mstrPrivs, "�������") > 0 Then
                'û�е�½�����
                If mintAuditing = 0 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    'ͬһ���˱�Ȩ�޿��Ʋ��ܽ������
                    
'                    MsgBox "�����˺������Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
                End If

                
                '�жϵ�½ʱ���������Ƿ�Ϊͬһ��.
                If strVerifyMan = mstrAuditingMan Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    If strVerifyMan = mstrAuditingMan Then
                        MsgBox "�����˺������Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    '��½���������˺͵�ǰ�û�Ϊͬһ����
'                    MsgBox "��½���������˺͵�ǰ�û�Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
                End If
            End If
            '���ʱ���Ƿ����
            If mintAuditing < 0 Then
                If DateDiff("n", mDataAuditing, Now) > Abs(mintAuditing) * 60 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
'                        MsgBox "�����Чʱ���ѹ�,�����µ�½�����!", vbInformation, gstrSysName
                    '����Чʱ����ڿ��Խ������
                End If
            End If
            
            blnClueTo = zlDatabase.GetPara("���ʱ����Ҫ��ʾ", 100, 1208, 0)
            
'            If blnClueTo = False Then
'                If mintHandleState = 0 Then
'                    If MsgBox("���Ҫ��ˡ�" & strPatienName & "���걾�ı�����", _
'                            vbQuestion + vbYesNo, gstrSysName) = vbNo Then
'                            Call SetControlFocus
'                            Exit Sub
'                    End If
'                End If
'            End If
            
            '11210 Ȩ�ޡ�δ�շ���ˡ�������˵�������ʱ��δ��Ч��
            If InStr(mstrPrivs, "δ�շ����") <= 0 Then
                If CheckChargeState(mlngKey, False) = False Then
                    MsgBox "����δ�շѣ����ܽ�����ˣ�", vbInformation, gstrSysName
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            '21137 �ѹ鵵���治�����
            gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                    "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                    "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                    " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rsTmp.EOF = False Then
                MsgBox "���˱���סԺ�Ĳ������ύ��飬���ܽ�����ˣ�", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '���סԺ�����Ƿ��Ժ���л��۵�
            If CheckExesState(mlngKey) = False Then
                MsgBox "��ǰסԺ���˻��л��۵�δ��ˣ����ѳ�Ժ��Ԥ��Ժ��", vbInformation, Me.Caption
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            
            '���鲡����Ϣ��һ��ʱʹ�ò�����Ϣ�滻
            Call CheckPatientInfo(mlngKey)
            
            
            '������˹����ж�
            strErrInfo = ""
            If VerifyAuditingRule(mlngKey, strErrInfo) = 1 Then
                If Mid(strErrInfo, 1, 2) = "1|" And InStr(mstrPrivs, "ǿ����˹���") <= 0 Then
                    strErrInfo = Mid(strErrInfo, 3)
                    MsgBox "<" & strPatienName & ">�ļ��鵥���δͨ��!" & vbNewLine & strErrInfo
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
                strErrInfo = Mid(strErrInfo, 3)
                If MsgBox("<" & strPatienName & ">�ļ��鵥���δͨ��!�Ƿ�����?" & vbNewLine & strErrInfo, _
                    vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            
            intPrivacy = zlDatabase.GetPara("���浥�Ƿ���ʾ��˽��Ŀ", 100, 1208, 0)
            If mintUnion = 1 Then
                If mSendReport = 1 And str������ = "" Then
                    '����
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_����걾��¼_���󱨸�(" & mlngKey & ",1,'" & UserInfo.���� & "')"
                Else
                    gstrSql = " select id from ����걾��¼ where ҽ��id = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAdvice)
                    Do Until rsTmp.EOF
                        'ǩ�����ɹ�ʱ�˳�
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Signature;" & rsTmp("ID") & ";" & mstrAuditingManID
                        
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & rsTmp("ID") & ",'" & IIf(mstrAuditingMan = "" _
                            , UserInfo.����, mstrAuditingMan) & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                            

                        
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & rsTmp("ID") & "," & intPrivacy & ",'" & gstrUnitName & "')"         '��˺��������浥

                        
                        rsTmp.MoveNext
                    Loop
                End If
            Else
                If mSendReport = 1 And str������ = "" Then
                    '����
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_����걾��¼_���󱨸�(" & mlngKey & ",1,'" & UserInfo.���� & "')"
                Else
                    'ǩ�����ɹ�ʱ�˳�
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Signature;" & mlngKey & ";" & mstrAuditingManID
                    
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & mlngKey & ",'" & IIf(mstrAuditingMan = "" _
                        , UserInfo.����, mstrAuditingMan) & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"

                    
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & mlngKey & "," & intPrivacy & ",'" & gstrUnitName & "')"        '��˺��������浥

                End If
            End If
            
            '����ִ��SQL
            gcnOracle.BeginTrans
            blnRollBack = True
            For intLoop = 1 To UBound(astrSQL)
                If UCase(Mid(astrSQL(intLoop), 1, 3)) = "ZL_" Then
                    zlDatabase.ExecuteProcedure astrSQL(intLoop), "��˱걾"
                Else
                    'ǩ�����ɹ�ʱ�˳�
                    If Signature(Val(Split(astrSQL(intLoop), ";")(1)), mstrAuditingManID) = False Then
                        gcnOracle.RollbackTrans
                        blnRollBack = False
                        Exit Sub
                    End If
                End If
            Next
            gcnOracle.CommitTrans
            
            Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Value = "�Ѽ���"
            Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon = 7
            
            If blnAutoPrint Then ReportPrint True                                           '�Ƿ���ɺ�ֱ�Ӵ�ӡ����
            If mblnAout = False Then
                MoveStation 1, 2
            Else
                MoveStation 1, 0
            End If
            InsertOneRecored mlngKey, False
            gintSelectFocus = 1
        Case mActR.���ͱ���                                                                         '���ͱ���
            '������˹����ж�
            strErrInfo = ""
            If VerifyAuditingRule(mlngKey, strErrInfo) = 1 Then
                If Mid(strErrInfo, 1, 2) = "1|" And InStr(mstrPrivs, "ǿ����˹���") <= 0 Then
                    strErrInfo = Mid(strErrInfo, 3)
                    MsgBox "<" & strPatienName & ">�ļ��鵥���δͨ��!" & vbNewLine & strErrInfo
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
                strErrInfo = Mid(strErrInfo, 3)
                If MsgBox("<" & strPatienName & ">�ļ��鵥���δͨ��!�Ƿ�����?" & vbNewLine & strErrInfo, _
                    vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            

            gstrSql = "Zl_����걾��¼_���󱨸�(" & mlngKey & ",1,'" & UserInfo.���� & "')"
            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            InsertOneRecored mlngKey, False
            MoveStation 1, 2
'            Me.rptList.FocusedRow.Record(mCol.����״̬).Value = "�Ѳ���"
'            Me.rptList.FocusedRow.Record(mCol.����״̬).Icon = 12
'            Me.rptList.Populate
        Case mActR.������˱��� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If InStr(1, mstrPrivs, "��˱걾") <= 0 Then
                'û��Ȩ�޺������û���½ʱ�˳�
                MsgBox "��û��Ȩ�޽������,�����µ�½���������Ա�������!", vbInformation, gstrSysName
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            blIf = False

            If (strVerifyMan = mstrAuditingMan Or (mstrAuditingMan = "" And strVerifyMan = UserInfo.����)) And InStr(1, mstrPrivs, "�������") > 0 Then
                'û�е�½�����
                If mintAuditing = 0 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    'ͬһ���˱�Ȩ�޿��Ʋ��ܽ������

'                    MsgBox "�����˺������Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
                End If

                '�жϵ�½ʱ���������Ƿ�Ϊͬһ��.
                If strVerifyMan = mstrAuditingMan Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    If strVerifyMan = mstrAuditingMan Then
                        MsgBox "�����˺������Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    '��½���������˺͵�ǰ�û�Ϊͬһ����
'                    MsgBox "��½���������˺͵�ǰ�û�Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
                End If
            End If

'            Call frmLisStationAuditing.ShowEdit(Me, mlngDeptID, mstrPrivs, IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan), _
                                                mintAuditing, _
                                                mDataAuditing)
            '���ʱ���Ƿ����
            If mintAuditing < 0 Then
                If DateDiff("n", mDataAuditing, Now) > Abs(mintAuditing) * 60 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "û�������,���ܽ������!��ȡ�����������ٵǼ�.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
'                        MsgBox "�����Чʱ���ѹ�,�����µ�½�����!", vbInformation, gstrSysName
                    '����Чʱ����ڿ��Խ������
                End If
            End If
            
            Call frmBatchAction.ShowMe(Me, 2, mlngMachineID, mstrPrivs, IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan), _
                                                mintAuditing, _
                                                mDataAuditing, mlngDeptID, mstrAuditingManID)
            gintSelectFocus = 1
        Case mActR.���ȡ�� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If strSaveAs = "��" Then
                MsgBox "��ǰ����������ת�뱸�ݣ�����ȡ����ˣ�", vbInformation, gstrSysName
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            
            If InStr(";" & mstrPrivs & ";", ";���ȡ��;") <= 0 Then
                If DateDiff("h", strAuditingDate, zlDatabase.Currentdate) > 24 Then
                    MsgBox "��ֻ��ȡ��24Сʱ�ڵ���˱��浥������ϵ�ϼ���ʦȡ�����!", vbInformation, Me.Caption
                    Call SetControlFocus
                    Exit Sub
                End If
            End If
            '21434
            If InStr(";" & mstrPrivs & ";", ";�����Ѵ�ӡ�ɻع�;") <= 0 Then
                If Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon = 8 Then
                    MsgBox "��ֻ��ȡ��δ��ӡ����˱��浥������ϵ�ϼ���ʦȡ�����!", vbInformation, Me.Caption
                    Call SetControlFocus
                    Exit Sub
                End If
            End If
            '21137 �ѹ鵵���治��ȡ��
            gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                    "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                    "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                    " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rsTmp.EOF = False Then
                MsgBox "���˱���סԺ�Ĳ������ύ��飬����ȡ����ˣ�", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If intLook = 0 Then
                strTmp = "���Ҫȡ����" & strPatienName & "���걾�ı��������"
            Else
                strTmp = "ҽ���Ѳ��ġ�" & strPatienName & "���ı��棬�Ƿ�ȷ��Ҫȡ����ˣ�"
            End If
            
            If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call SetControlFocus
                gintSelectFocus = 1: Exit Sub
            End If
            
            gstrSql = "select ����걾id from ����ǩ����¼ where ����걾id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rsTmp.EOF = False Then
                If gobjESign Is Nothing Then
                    MsgBox "����ȡ��ǩ��������ϵͳ����������ʹ�õ���ǩ����", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
            
            
            If mintUnion = 1 Then
                gstrSql = " select id from ����걾��¼ where ҽ��id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAdvice)
                Do Until rsTmp.EOF
                strSQL = "ZL_����걾��¼_���ȡ��(" & rsTmp("ID") & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    rsTmp.MoveNext
                Loop
            Else
                strSQL = "ZL_����걾��¼_���ȡ��(" & mlngKey & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            InsertOneRecored mlngKey, False
'            Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Value = ""
'            Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon = -1
'            Me.rptList.FocusedRow.Record(mCol.����״̬).Value = ""
'            Me.rptList.FocusedRow.Record(mCol.����״̬).Icon = -1
'            Me.rptList.Populate
'            RptListFilter
            gintSelectFocus = 1
'            RefreshData

        Case mActR.���������
            frmPatinetAuditing.ShowMe Me, mstrPrivs, mstrAuditingManID
            Call RefreshData
            
        Case mActR.������� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
            If intMicrobe = 1 Then Exit Sub     '�����΢�����˳�
            
            If MsgBox("���Ҫ������" & strPatienName & "���걾�ļ�����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
            End If
            strSQL = "ZL_����걾��¼_�걾����(" & mlngKey & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            RefreshData
            gintSelectFocus = 1
        Case mActR.ȡ������ ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If MsgBox("���Ҫȡ����" & strPatienName & "���ļ�������", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
            End If
            strSQL = "ZL_����걾��¼_ȡ������(" & mlngKey & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            RefreshData
            gintSelectFocus = 1
        Case mActR.��д���� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '�г�����ʱ�������޸Ľ����
            If str������ <> "" And mSendReport = 1 Then Exit Sub
            
            strSQL = "select ������,����ʱ�� from ����걾��¼ where id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngKey)
            If rsTmp.EOF = False Then
                strVerifyMan = Nvl(rsTmp("������"))
                strVerifydate = Nvl(rsTmp("����ʱ��"))
            End If
            
            '�����Ƿ����޸����˱���
            If UserInfo.���� <> strVerifyMan And strVerifyMan <> "" Then
                If InStr(1, mstrPrivs, "�޸����˽��") <= 0 Then
                    MsgBox "��û���޸����˱����Ȩ�ޣ��������Ա��ϵ��", vbInformation, gstrSysName
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            '�����ܹ���д���޸ķǱ��ռ���ı�����
            If strVerifydate <> "" Then
                If DateDiff("d", CDate(strVerifydate), Now) > 1 Then
                    If InStr(1, mstrPrivs, "�޸����ս��") <= 0 Then
                        MsgBox "��û��Ȩ����д���޸ķǱ��ռ���ı�����", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            If Val(Me.rptList.FocusedRow.Record(mCol.΢����걾).Value) <> 1 Then
                Me.TabCtlWindow.Item(0).Selected = True
                If mfrmWrite.ZlEditStart(mlngKey) = True Then
                    mintEditState = 5
                End If
            Else
                Me.TabCtlWindow.Item(1).Selected = True
                If mfrmWrite2.ZlEditStart(mlngKey) = True Then
                    mintEditState = 5
                End If
            End If
        Case mActR.��д��������
            
            
            '�����Ƿ����޸����˱���
            If UserInfo.���� <> strVerifyMan And strVerifyMan <> "" Then
                If InStr(1, mstrPrivs, "�޸����˽��") <= 0 Then
                    MsgBox "��û���޸����˱����Ȩ�ޣ��������Ա��ϵ��", vbInformation, gstrSysName
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            '�����ܹ���д���޸ķǱ��ռ���ı�����
            If strVerifydate <> "" Then
                If DateDiff("d", CDate(strVerifydate), Now) > 1 Then
                    If InStr(1, mstrPrivs, "�޸����ս��") <= 0 Then
                        MsgBox "��û��Ȩ����д���޸ķǱ��ռ���ı�����", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            mintEditState = 6
            mfrmLabMicrobe3Report.ZlEditStart
            
            
        Case mActR.д�벡��
            
            If intMicrobe = 1 Then
                strSQL = "Zl_���鱨�浥_Update(" & mlngKey & ",0,'" & gstrUnitName & "')"       '΢�����������浥����
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Case mActR.��֤ǩ��
            Call VerifySignature(mlngKey)
    End Select
    
    '���˽����б�
'    RptListFilter
    If Me.rptList.Rows.Count = 0 And mintEditState = 0 Then
        mfrmRequest.ZlCancel
        mfrmWrite.ZlCancel
        mfrmWrite2.zlRefresh -1
    End If
    Exit Sub
errH:
    If blnRollBack = True Then
        blnRollBack = False
        gcnOracle.RollbackTrans
    End If
    AutoRefresh = True                                                                      '���¿�ʼ�Զ�ˢ��
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LoadRegistSetup()
    mintEditState = 0                                           '��ǰ�༭״̬��0-�Ǳ༭��1-�������գ�2-�����Ǽǣ�4-����ˣ�3-���º��գ�5-����༭
    
    On Error GoTo errH
    
    'Ŀǰ�Ƿ����������յǼ�״̬
    mintContinue = IIf(zlDatabase.GetPara("��������", 100, 1208, False), 1, 0)

    Set mfrmRequest = New frmLabRequest                             '���յǼǴ���
    Set mfrmWrite = New frmLisStationWrite                          '������д����
    Set mfrmWrite2 = New frmLisStationWrite2                        '��д΢����
    Set mfrmTrack = New frmLabTrack                                 '���ζԱ�
    Set mfrmLabMicrobe3Report = New frmLabMicrobe3Report            '��������
    Set mfrmLabMainSampleUnion = frmLabMainSampleUnion              '�걾�ϲ�
'    Set mclsExpenses = New zlCISKernel.clsDockExpense           '���ò���
'    Set mclsOutAdvices = New zlCISKernel.clsDockOutAdvices      '����ҽ��
'    Set mclsInAdvices = New zlCISKernel.clsDockInAdvices        'סԺҽ��
'    Set mclsOutEPRs = New zlRichEPR.cDockOutEPRs                '����ҽ��
'    Set mclsInEPRs = New zlRichEPR.cDockInEPRs                  'סԺ����
'    Set mcolSubForm = New Collection
    
'    mcolSubForm.Add mclsExpenses.zlGetForm, "_����"             '�õ��Ӵ���
'    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_����ҽ��"
'    mcolSubForm.Add mclsInAdvices.zlGetForm, "_סԺҽ��"
'    mcolSubForm.Add mclsOutEPRs.zlGetForm, "_���ﲡ��"
'    mcolSubForm.Add mclsInEPRs.zlGetForm, "_סԺ����"
    
'    Set mfrmLabMainImage = frmLabMainImage
'    Call mclsExpenses.zlDefCommandBars(Me, Me.cbrthis)
    
    
    '���Һ�����ID
    mlngDeptID = zlDatabase.GetPara("ȱʡ����ID", 100, 1208, mlngDeptID)
    mlngMachineID = zlDatabase.GetPara("��������", 100, 1208, mlngMachineID)
    mstrMachineGroup = zlDatabase.GetPara("����С��", 100, 1208, mstrMachineGroup)
    mblnAout = zlDatabase.GetPara("��˺�������һ������걾", 100, 1208, mblnAout)
    
    '����ˢ�¹���
    Call GetVerifying
'    chkSoure(0).Value = IIf(mblnVerifying(0), 1, 0)
'    chkSoure(1) = IIf(mblnVerifying(1), 1, 0)
'    chkSoure(2) = IIf(mblnVerifying(2), 1, 0)
'    chkSoure(3) = IIf(mblnVerifying(3), 1, 0)
'    chkSoure(4) = IIf(mblnVerifying(4), 1, 0)
'    chkSoure(5) = IIf(mblnVerifying(5), 1, 0)
    '���ˢ��Dpk��ʽ���������ݿ�̫����
'    dkpMain.LoadStateFromString zlDatabase.GetPara("DKP����", 100, 1208, "")
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    
    blnComm = Val(zlDatabase.GetPara("��������˫��", 100, 1208, 0))
    blnAutoPrint = zlDatabase.GetPara("��˴�ӡ", 100, 1208, 0)
    blnAutoRefresh = Val(zlDatabase.GetPara("�Զ�ˢ��", 100, 1208, 1))
    mintUnion = zlDatabase.GetPara("������������ʾ������Ŀ", 100, 1208, 0)
    mMakeNoRule = zlDatabase.GetPara("�걾������ɹ���", 100, 1208, "��  ��")
    mSendReport = zlDatabase.GetPara("ʹ�ö����������", 100, 1208, 0)
    mstrPrintDepts = zlDatabase.GetPara("ֻ��ָ�����ұ��浥", 100, 1208, "")
    
    int��촦��ʽ = Val(zlDatabase.GetPara("��첡����Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
    intԺ�⴦��ʽ = Val(zlDatabase.GetPara("Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
    intסԺ����ʽ = Val(zlDatabase.GetPara("סԺ������Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
    int���ﴦ��ʽ = Val(zlDatabase.GetPara("���ﲡ����Ϣ��һ�µĴ���ʽ", 100, 1208, True, 1))
    
    mTodayQCPrivs = GetPrivFunc(100, 1210)
    mHistoryPrivs = GetPrivFunc(100, 1211)
    
    '�����кʹ�����ʱ��ѡ��
    With cboʱ��
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "������"
        .AddItem "��  ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "�Զ���"
    End With
    
    dtpDate.Value = Now
    dtpDateEnd.Value = Now
    
    '����ǩ����֤����
    gintCA = Val(zlDatabase.GetPara("����ǩ����֤����", glngSys))
    '����ǩ�����Ƴ���
    gstrESign = zlDatabase.GetPara("����ǩ��ʹ�ó���", glngSys)
    
    If Mid(gstrESign, 6, 1) = "1" Then
        If gintCA <> 0 Then
            'If InStr(GetInsidePrivs(p����ҽ���´�), "ҽ������ǩ��") > 0 And gobjESign Is Nothing Then
            If gobjESign Is Nothing Then
                On Error Resume Next
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                Err.Clear: On Error GoTo 0
                If Not gobjESign Is Nothing Then
                    Call gobjESign.Initialize(gcnOracle, glngSys)
                End If
            End If
        Else
            Set gobjESign = Nothing
        End If
    End If
    Exit Sub
errH:
    MsgBox Err.Description
End Sub
Private Sub ShowOrHideItem(Control As CommandBarControl, DkpID As Integer)
    '����               '��ʾ������
    Dim Pane As Pane
    Set Pane = Me.dkpMain.FindPane(DkpID)
    If Control.Checked = True Then
        Pane.Close
    Else
        Pane.Select
    End If
    If mlngKey <> 0 Then
        ReadImageData mlngKey, False
    End If
    Me.dkpMain.RecalcLayout
    Me.cbrthis.RecalcLayout
End Sub
Private Sub BackOrNextPatient(Move As Integer)
    '����                 �ƶ�����һ�����˻���һ������
    '����                 Move = 1 ��һ���� =2 ��һ����
    Dim Rerow As ReportRow
    Dim i As Long
    With Me.rptList
        If .Rows.Count <= 0 Then Exit Sub
        i = .SelectedRows(0).Index
        If Move = 1 Then            '�����ƶ�
            If i - 1 >= 0 Then
                i = i - 1
                .FocusedRow = .Rows(i)
            End If
        Else
            If i < .Rows.Count - 1 Then
                i = i + 1
                .FocusedRow = .Rows(i)
            End If
        End If
    End With
End Sub

Private Function FindPatient(strFind As String) As Boolean
    '����:              ���Ҳ���
    '����               ��ѯ�ֶ� ��ʶ�� �걾�� ������ƴ����д
    '����               "����Ϊ�걾�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
    Dim Rerow As ReportRow
    Dim strPatientID As String                                          '��ʶ��
    Dim strSampleID As String                                           '�걾��
    Dim strPatientName As String                                        '��������
    Dim strPatientPY As String                                          '����ƴ������
    Dim lngPatientID As Long                                            '����ID
    Dim strSource As String                                             '�������
    Dim strRegisterNo As String                                         '�Һŵ�
    Dim strChargeNo As String                                           '�շѵ�
    Dim strBarCode As String                                            '��������
    Dim strSQL As String                                                '���ݲ�ѯ���
    Dim rsTmp As New ADODB.Recordset                                    '���ݼ�
    Dim strWhere As String                                              '�����������
    On Error GoTo errH
    
    '��λǰ��ˢ��һ��
'    Call RefreshData
    
    
    
    If Me.TabList(0).Selected = True Then
        If Me.rptList.Rows.Count = 0 Then Exit Function                     'û�м�¼ʱ�˳�
        strFind = UCase(strFind)
        For Each Rerow In Me.rptList.Rows
            '��ȡ������Ҫ���ֶ���Ϣ
            Select Case Mid(strFind, 1, 1)
                Case "-"                                                    '����ID
                    lngPatientID = Val(Rerow.Record(mCol.����ID).Value)
                    strWhere = Mid(strFind, 2)
                    If strWhere = lngPatientID Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "+", "*"                                               '�����/סԺ��
                    strPatientID = Rerow.Record(mCol.��ʶ��).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strPatientID Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "."                                                    '�Һŵ�
                    strRegisterNo = Rerow.Record(mCol.�Һŵ�).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strRegisterNo Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "/"                                                    '�շѵ�
                    strChargeNo = Rerow.Record(mCol.��ʶ��).Value
                    strWhere = zlCommFun.GetFullNO(Mid(strFind, 2))
                    If strWhere = strChargeNo Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case Else                                                   '�걾�š�������������ҡ�����
                    strSampleID = Nvl(Rerow.Record(mCol.�걾��).Value)
                    strPatientName = Rerow.Record(mCol.����).Value
                    strPatientPY = zlCommFun.SpellCode(Rerow.Record(mCol.����).Value)
                    strBarCode = Rerow.Record(mCol.��������).Value
                    If strSampleID = strFind Or (strPatientName Like UCase(strFind) & "*") Or (strPatientPY Like UCase(strFind) & "*") _
                            Or strBarCode = UCase(strFind) Then
                        If Val(Rerow.Record(mCol.��λ).Value) <= 0 Then
                            Rerow.Record(mCol.��λ).Value = 1
                            Me.rptList.FocusedRow = Rerow
                            Me.rptList.Populate
                            mlngKey = Rerow.Record(mCol.ID).Value
                            FindPatient = True
                            Exit Function
                        End If
                    End If
                    
            End Select
        Next
        For Each Rerow In Me.rptList.Rows
            Rerow.Record(mRCol.��λ).Value = 0
        Next
        '������������
        If BlnIsNumber(strFind) Then
            strSQL = "select distinct b.����id  from ����ҽ������ a , ����ҽ����¼ b " & _
                     " Where a.ҽ��id = b.ID And a.�������� = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind)
            If rsTmp.EOF = False Then
                Me.rptList.Tag = ";;,;;;;;;;,True;,;0;;0;;;;1;;" & rsTmp(0)
                RefreshData True
                rptList.Tag = ""
                Exit Function
            End If
        End If
        Me.rptList.Populate
    End If
    
    If Me.TabList(1).Selected = True Then
        If Me.rptList1.Rows.Count = 0 Then Exit Function                     'û�м�¼ʱ�˳�
        strFind = UCase(strFind)
        For Each Rerow In Me.rptList1.Rows
            '��ȡ������Ҫ���ֶ���Ϣ
            Select Case Mid(strFind, 1, 1)
                Case "-"                                                    '����ID
                    lngPatientID = Val(Rerow.Record(mRCol.����ID).Value)
                    strWhere = Mid(strFind, 2)
                    If strWhere = lngPatientID Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mRcol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "+", "*"                                               '�����/סԺ��
                    strPatientID = Rerow.Record(mRCol.��ʶ��).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strPatientID Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "."                                                    '�Һŵ�
                    strRegisterNo = Rerow.Record(mRCol.�Һŵ�).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strRegisterNo Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "/"                                                    '�շѵ�
                    strChargeNo = Rerow.Record(mRCol.��ʶ��).Value
                    strWhere = zlCommFun.GetFullNO(Mid(strFind, 2))
                    If strWhere = strChargeNo Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case Else                                                   '�걾�š�������������ҡ�����
'                    strSampleID = Nvl(Rerow.Record(mCol.�걾��).Value)
                    strPatientName = Rerow.Record(mRCol.����).Value
                    strPatientPY = zlCommFun.SpellCode(Rerow.Record(mRCol.����).Value)
'                    strBarCode = Rerow.Record(mCol.��������).Value
                    If strSampleID = strFind Or (strPatientName Like UCase(strFind) & "*") Or (strPatientPY Like UCase(strFind) & "*") _
                            Or strBarCode = UCase(strFind) Then
                        If Val(Rerow.Record(mRCol.��λ).Value) <= 0 Then
                            Rerow.Record(mRCol.��λ).Value = 1
                            Me.rptList1.FocusedRow = Rerow
                            Me.rptList1.Populate
    '                        mlngkey = Rerow.Record(mCol.ID).Value
                            FindPatient = True
                            Exit Function
                        End If
                    End If
            End Select
        Next
        For Each Rerow In Me.rptList1.Rows
            Rerow.Record(mRCol.��λ).Value = 0
        Next
        '������������
'        If IsNumeric(strFind) = True And Len(strFind) >= 12 Then
'            strsql = "select distinct b.����id  from ����ҽ������ a , ����ҽ����¼ b " & _
'                     " Where a.ҽ��id = b.ID And a.�������� = [1] "
'            Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, strFind)
'            If rsTmp.EOF = False Then
'                rptList.Tag = ";;,;;;;;,True;,;0;;0;0;;" & rsTmp(0)
'                RefreshData True
'                rptList.Tag = ""
'                Exit Function
'            End If
'        End If
        Me.rptList.Populate
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = "���˲ɼ��嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub InsertOneRecored(lngKey As Long, Optional blnNew As Boolean = True, Optional blnGoto As Boolean = True)
    '����                                               'ͨ������걾ID�ҵ�һ��¼��׷�ӵ��б�
    '����   blnNew                                      �Ƿ������걾
    '       blnGoto                                     �Ƿ�λ
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset, rsZk As ADODB.Recordset
    Dim Record As ReportRecord
    Dim i As Integer
    Dim blnNewRecord As Boolean               '�Ƿ�������¼�Զ��ж�
    Dim lngRowIndex As Long                                         '������
    Dim lngRowID As Long                                            '��ID
    Dim lngloop As Long
    Dim intLoop As Integer
    Dim blnPathPatient As Boolean                                   '�ٴ�·������
    Dim blnAdviceKey As Long                                        'ҽ��ID
    
    mblnCompelRefresh = True    'ˢ��ʱ����ǿ��ˢ��
    blnPathPatient = False
    On Error GoTo errH
    
    strSQL = "select ҽ��id from ����걾��¼ where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rsTmp.EOF = True Then Exit Sub
    blnAdviceKey = Nvl(rsTmp("ҽ��ID"), 0)
    
    strSQL = "Select /*+ rule */     Decode(a.�Ƿ���, 1, '', '����ʧ��') As ����," & vbNewLine & _
            "       decode(a.�걾���,1,'����',decode(a.����,1,'����', '')) As ����,Decode(a.����״̬, 1, '������', 2, '�Ѽ���') As ִ��״̬," & vbNewLine & _
            "       Decode(A.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���','����') As �������," & vbNewLine & _
            "       Decode(Sign(Nvl(a.�Ƿ��ʿ�Ʒ, 0)), 0, '��ͨ', 1, '�ʿ�', -1, '�ȶ�') As �걾����," & vbNewLine & _
            "       Decode(a.����id, Null," & vbNewLine & _
            "                 To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
            "                 a.�걾���) As �걾����ʾ,a.�걾���, A.�Һŵ� ," & vbNewLine & _
            "       Decode(A.������Դ, 1, to_char(nvl(a.�����,a.��ʶ��)), 2, to_char(nvl(a.סԺ��,a.��ʶ��)), 3, to_char(nvl(a.NO,a.��ʶ��)), 4, to_char(nvl(a.�����,a.��ʶ��)),to_char(a.��ʶ��)) As ��ʶ��,a.����,a.�Ա�,a.����," & vbNewLine & _
            "       Decode(a.������Դ,2,S.��������,b.��������) as ��������," & vbNewLine & _
            "       a.������ As �������,a.ҽ��ID,a.����ID,'' As ת��,a.Id,a.����ʱ�� ,a.��ӡ����,a.����id," & vbNewLine & _
            "       a.����ʱ��,a.΢����걾,a.������,a.�����,To_Char(A.Ӥ��) As Ӥ��,a.��������,a.�������ID As ��������id," & vbNewLine & _
            "       a.��ҳID,a.������,a.��������,a.���䵥λ,a.�����,a.סԺ��,a.��������,a.�Һŵ�,a.������Ŀ,e.���� as �������,f.���� as ��������, " & vbNewLine & _
            "       a.�������ID as ���˿���ID,a.����,a.������,a.�걾��̬,a.������,a.����ʱ��,a.�걾���� as ����걾,a.NO,a.������,a.����ʱ��, " & vbNewLine & _
            "       abs(nvl(a.�Ƿ��ʿ�Ʒ,0)) as �ȶԴ���,a.���ʱ��,nvl(a.�걾���,0) as �걾���, " & vbNewLine & _
            "       nvl(a.����,0) as ҽ������, nvl(a.�걾���,0) as �걾����,decode(a.���˿���,null,C.����,a.���˿���) as ���˿���, " & vbNewLine & _
            "       a.��������,nvl(R.����״̬,0) as ����״̬,nvl(R.����ID,0) as ���淢��,a.������,a.����ʱ��,b.������λ,p.��Ŀ,p.����,b.������,  " & vbNewLine & _
            "       a.���δͨ��,a.������Դ,a.���Ϊ��,nvl(s.·��״̬,0) as �ٴ�·������,decode(d.�����Ƿ����,1,'�������','����δ���')  as ������� " & vbNewLine & _
            " From ����걾��¼ a ,���ű� E , �������� f , ������Ϣ B , ���ű� C,����ҽ������ R,����ҽ������ P,������ҳ S,������ˮ�߱걾 d " & vbNewLine & _
            " Where a.�������ID = E.id(+) and a.����id=f.id(+) and a.����ID = B.����ID(+) and B.��ǰ����ID = C.id(+) " & vbNewLine & _
            " " & IIf(blnAdviceKey = 0, " and  a.ID = [1] ", " and a.ҽ��id=[2] ") & vbNewLine & _
            " And a.ҽ��ID = R.ҽ��ID(+) and A.ҽ��Id = P.ҽ��ID(+) and a.id=d.�걾id(+)" & vbNewLine & vbNewLine & _
            " and a.����ID = S.����ID(+) and a.��ҳID = s.��ҳID(+)  "
            
    If mlngMachineID > 0 Then
        strSQL = strSQL & " And a.����id = [3] "
    End If
      
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngKey, blnAdviceKey, mlngMachineID)
                                                             
    If rsTmp.EOF = True Then Exit Sub       'û��ʱ�˳�

    'ˢ��ǰ��¼һ��λ��
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For i = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(i).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(i).Record.Index
                    lngRowID = Me.rptList.Rows(i).Record(mCol.ID).Value
                End If
            Next
        End If
    End If

'    If lngRowIndex = 0 And Me.rptList.Rows.Count > 0 Then
'
'    End If
    
    
    blnNewRecord = True
    
    For i = 0 To Me.rptList.Records.Count - 1
        If Me.rptList.Records(i).Item(mCol.ID).Value = lngKey Then
            Set Record = Me.rptList.Records(i)
            blnNewRecord = False
            Exit For
        End If
    Next
    
    If blnNewRecord = True Then
        Set Record = Me.rptList.Records.Add
        For i = 0 To Me.rptList.Columns.Count + 1
            Record.AddItem ""
        Next
    End If
    
    'ǰ�漸����Ҫ����ͼ��
    Record.Item(mCol.����).Value = IIf(Nvl(rsTmp("�걾����")) = 1, "����", "")
    If Record.Item(mCol.����).Value = "����" Then
        Record.Item(mCol.����).Icon = 1
    Else
        Record.Item(mCol.����).Icon = -1
    End If
    
    Record.Item(mCol.����ҽ��).Value = IIf(Nvl(rsTmp("ҽ������")) = 1, "����", "")
    If Record.Item(mCol.����ҽ��).Value = "����" Then
        Record.Item(mCol.����ҽ��).Icon = 14
    Else
        Record.Item(mCol.����ҽ��).Icon = -1
    End If
    
'    If Nvl(rsTmp("������")) <> "" Then
'        Record.Item(mCol.����״̬).Value = "�ѳ���"
'        Record.Item(mCol.����״̬).Icon = 13
'    Else
'        Record.Item(mCol.����״̬).Value = ""
'        Record.Item(mCol.����״̬).Icon = -1
'    End If
    
    If Nvl(rsTmp("����״̬")) = 1 Then
        Record.Item(mCol.����״̬).Value = "�Ѳ���"
        Record.Item(mCol.����״̬).Icon = 11
    End If
    If rsTmp("�������") & "" = "�������" Then
        Record.Item(mCol.�������).Value = "��"
    Else
        Record.Item(mCol.�������).Value = "��"
    End If
    If Val(Nvl(rsTmp("�ٴ�·������"))) = 1 Then
        blnPathPatient = True
        Record.Item(mCol.�ٴ�·������).Icon = 15
    Else
        Record.Item(mCol.�ٴ�·������).Icon = -1
    End If
    
    If CInt(Nvl(rsTmp("��ӡ����"), "0")) > 0 Then
        Record.Item(mCol.ִ��״̬).Value = "�Ѵ�ӡ"
        Record.Item(mCol.ִ��״̬).Icon = 8
    ElseIf Nvl(rsTmp("ִ��״̬")) = "�Ѽ���" Then
        Record.Item(mCol.ִ��״̬).Value = "�Ѽ���"
        Record.Item(mCol.ִ��״̬).Icon = 7
    ElseIf Nvl(rsTmp("������")) <> "" Then
        Record.Item(mCol.ִ��״̬).Value = "����"
        Record.Item(mCol.ִ��״̬).Icon = 13
    ElseIf Nvl(rsTmp("����")) = "" Then
        Record.Item(mCol.ִ��״̬).Value = "�Ѵ���"
        Record.Item(mCol.ִ��״̬).Icon = 6
    Else
        Record.Item(mCol.ִ��״̬).Value = ""
        Record.Item(mCol.ִ��״̬).Icon = -1
    End If

    
    Record.Item(mCol.����).Value = Nvl(rsTmp("����")) '& IIf(Nvl(rsTmp("Ӥ��"), 0) > 0, "(Ӥ��)", "")
    If Nvl(rsTmp("�걾����")) = "�ʿ�" Then
        Record.Item(mCol.�걾����).Value = "�ʿ�"
        Record.Item(mCol.�걾����).Icon = 3
        strSQL = "Select A.�걾id, B.����, B.����, B.ˮƽ From �����ʿؼ�¼ A, �����ʿ�Ʒ B Where A.�ʿ�Ʒid = B.ID And A.�걾id=[1]"
        Set rsZk = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp("ID"))))
        Do Until rsZk.EOF
            Record.Item(mCol.����).Value = "" & rsZk!���� & "," & rsZk!���� & ",ˮƽ" & rsZk!ˮƽ
            rsZk.MoveNext
        Loop
    ElseIf Nvl(rsTmp("�걾����")) = "�ȶ�" Then
        Record.Item(mCol.�걾����).Value = "�ȶ�"
        Record.Item(mCol.�걾����).Icon = 4
        Record.Item(mCol.����).Value = Record.Item(mCol.����).Value & "(" & Nvl(rsTmp("�ȶԴ���")) & ")"
    Else
        Record.Item(mCol.�걾����).Value = ""
        Record.Item(mCol.�걾����).Icon = -1
    End If
    
    Record.Item(mCol.�걾��).Value = Val(Nvl(rsTmp("�걾���")))
    Record.Item(mCol.�걾��).Caption = Trim(rsTmp("�걾����ʾ"))

    If Nvl(rsTmp("��������")) = "" Then
        
        If Nvl(rsTmp("Ӥ��"), 0) = 0 Then
            If IsNumeric(Nvl(rsTmp("����"))) = True Then
                Record.Item(mCol.����).Caption = Nvl(rsTmp("����")) & "��"
            Else
                If Nvl(rsTmp("����")) <> "��" And Nvl(rsTmp("����")) <> "0��" Then
                    Record.Item(mCol.����).Caption = Nvl(rsTmp("����"))
                End If
            End If
            If Record.Item(mCol.����).Caption <> "" Then
                Record.Item(mCol.����).Value = Val(rsTmp("����"))
            End If
        End If
    '            Record.Item(mCol.����).Caption = IIf(Nvl(rstmp("Ӥ��"), 0) > 0, "", _
                                   IIf(Nvl(rstmp("����")) = "��", "", _
                                   IIf(Nvl(rstmp("����")) = "0��", "", IIf(IsNumeric(Nvl(rstmp("����"))) = True, rstmp("����") & "��", rstmp("����")))))
    Else
        Record.Item(mCol.����).Value = Nvl(rsTmp("��������"))
        Record.Item(mCol.����).Caption = Nvl(rsTmp("����")) 'Nvl(rsTmp("��������")) & Nvl(rsTmp("���䵥λ"))
    End If
    
    If Nvl(rsTmp("��������")) <> "" Then
        Record.Item(mCol.����).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp("��������")), False)
    End If
    Record.Item(mCol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
    Record.Item(mCol.�������).Value = Nvl(rsTmp("�������"))
    Record.Item(mCol.������Ŀ).Value = Trim(Nvl(rsTmp("������Ŀ")))
    Record.Item(mCol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
    
    Record.Item(mCol.�������).Value = Nvl(rsTmp("�������"))
    Record.Item(mCol.ҽ��id).Value = Nvl(rsTmp("ҽ��ID"))
    Record.Item(mCol.����id).Value = Nvl(rsTmp("����ID"))
    Record.Item(mCol.ת��).Value = Nvl(rsTmp("ת��"))
    Record.Item(mCol.����ID).Value = Nvl(rsTmp("����id"))
    Record.Item(mCol.ID).Value = Nvl(rsTmp("ID"))
    Record.Item(mCol.�걾ʱ��).Caption = Format(Nvl(rsTmp("����ʱ��")), "MM-dd HH:mm:ss")
    Record.Item(mCol.�걾ʱ��).Value = Format(Nvl(rsTmp("����ʱ��")), "YYYY-MM-dd HH:mm:ss")
    Record.Item(mCol.����ʱ��).Caption = Format(Nvl(rsTmp("����ʱ��")), "MM-dd HH:mm")
    Record.Item(mCol.����ʱ��).Value = Format(Nvl(rsTmp("����ʱ��")), "YYYY-MM-dd HH:mm")
    Record.Item(mCol.΢����걾).Value = Val(Nvl(rsTmp("΢����걾")))
    '        Record.Item(mCol.�շѵ�).Value = Nvl(rstmp("�շѵ�"))
    Record.Item(mCol.�Һŵ�).Value = Nvl(rsTmp("�Һŵ�"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.�����).Value = Nvl(rsTmp("�����"))
    Record.Item(mCol.���˿���).Value = Nvl(rsTmp("���˿���"))
    Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"))
    'Record.Item(mCol.���ͺ�).Value = Nvl(rstmp("���ͺ�"))
    Record.Item(mCol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("��������"))
    Record.Item(mCol.��ҳID).Value = Nvl(rsTmp("��ҳID"))
    Record.Item(mCol.��������ID).Value = Nvl(rsTmp("��������Id"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"))
    Record.Item(mCol.���䵥λ).Value = Nvl(rsTmp("���䵥λ"))
    Record.Item(mCol.����).Value = Nvl(rsTmp("����"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.�걾��̬).Value = Nvl(rsTmp("�걾��̬"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
    Record.Item(mCol.����걾).Value = Nvl(rsTmp("����걾"))
    Record.Item(mCol.NO).Value = Nvl(rsTmp("NO"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
    Record.Item(mCol.���ʱ��).Value = Nvl(rsTmp("���ʱ��"))
    Record.Item(mCol.�걾���).Value = Nvl(rsTmp("�걾���"))
    Record.Item(mCol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
    Record.Item(mCol.�걾����).Value = Nvl(rsTmp("�걾����"))
    Record.Item(mCol.���˿���).Value = Nvl(rsTmp("���˿���"))
    Record.Item(mCol.�������).Value = Nvl(rsTmp("�������"))
    Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"), 0)
    Record.Item(mCol.���淢��).Value = Nvl(rsTmp("���淢��"), 0)
    Record.Item(mCol.���˿���ID).Value = Nvl(rsTmp("���˿���ID"), 0)
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
    Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
    Record.Item(mCol.���δͨ��).Value = Nvl(rsTmp("���δͨ��"))
    Record.Item(mCol.������Դ).Value = Nvl(rsTmp("������Դ"))
    Record.Item(mCol.�����).Value = Nvl(rsTmp("�����"))
    Record.Item(mCol.סԺ��).Value = Nvl(rsTmp("סԺ��"))
    If Nvl(rsTmp("��Ŀ")) = "��������" Then
        Record.Item(mCol.��λ).Value = Nvl(rsTmp("����"))
    End If
    Record.Item(mCol.���Ϊ��).Value = Val(Nvl(rsTmp("���Ϊ��")))
'    Record.Item(mCol.����״̬).Value = Nvl(rsTmp("����״̬"))

    '------��ú����
    For i = 0 To rptList.Columns.Count + 1
        If Val("" & rsTmp!΢����걾) = 0 Then
            If Record.Item(mCol.���Ϊ��).Value > 0 Then
                Record.Item(i).BackColor = vbWhite
            Else
                Record.Item(i).BackColor = &HFDD6C6
            End If
        Else
            Record.Item(i).BackColor = vbWhite
        End If
    Next
    

    Me.rptList.Populate
    
    '���˽����б�
    RptListFilter
    
    'û���ٴ�·������ʱ����ʾ��
    Me.rptList.Columns(6).Visible = blnPathPatient
    
    If blnGoto = True Then
        mfrmRequest.ZlCancel
        If Val(Record.Item(mCol.΢����걾).Value) = 1 Then
            mfrmWrite2.ZlCancel
        Else
            mfrmWrite.ZlCancel
        End If
    End If
    '���¶�λ����ǰ��λ��
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If

    
    '����λʱ�˳�
    If blnGoto = False Then
        If Not Me.rptList.FocusedRow Is Nothing Then
            If mlngKey = lngKey Then
                Call mfrmWrite.zlRefresh(mlngKey)
            End If
        End If
        Exit Sub
    End If
    '��λ����
    With Me.rptList
        For i = 0 To .Rows.Count - 1
            If .Rows(i).Record(mCol.ҽ��id).Value = blnAdviceKey Or .Rows(i).Record(mCol.ID).Value = lngKey Then
                Set .FocusedRow = .Rows(i)
                Exit For
            End If
        Next
    End With
    
    If Me.TabList.Selected.Index = 0 Then
        Call SetControlFocus
    Else
        Call SetControlFocus
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub AuditingRegister()
    Dim i As Integer
    Dim strVerifyMan As String              '������
    Dim blnCancel As Boolean
    Dim strLogID As String
    '����:          �����ע��,������û�������Ȩ��ʱ��½����˺�������
    
    '�����ǰ��(Ϊ�˱���)
    zlDatabase.SetPara "�Ƿ��о������Ȩ��", 0, 100, 1208
    zlDatabase.SetPara "�����", 0, 100, 1208
                    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            strVerifyMan = .Record(mCol.������).Value
        End With
    End If
                        
    frmLabAuditingLand.ShowMe Me, strVerifyMan, blnCancel, strLogID
    
    If blnCancel = True Then Exit Sub   'ȡ��ʱ������
    
    '�õ��Ƿ���Ȩ
    i = zlDatabase.GetPara("�Ƿ��о������Ȩ��", 100, 1208, 0)
    mstrAuditingMan = zlDatabase.GetPara("�����", 100, 1208, 0)
    If mstrAuditingMan = "0" Or mstrAuditingMan = "" Then mstrAuditingMan = ""
    If mstrAuditingMan <> "" Then
        mstrAuditingManID = strLogID
        Me.Caption = "���鼼ʦ����վ" & "-�����(" & mstrAuditingMan & ")"
    Else
        mstrAuditingManID = ""
        Me.Caption = "���鼼ʦ����վ"
    End If
    '��=0ʱ��Ȩ�ޱ仯,���¶���ʱ��
'    If i <> 0 Then
        mintAuditing = i
        mDataAuditing = Now
'    End If
    '���
    zlDatabase.SetPara "�Ƿ��о������Ȩ��", 0, 100, 1208
    zlDatabase.SetPara "�����", 0, 100, 1208
End Sub
Private Function SaveDisposal(intDisposal As Integer) As Boolean
    '����                   '�Ա��� ȡ�����в���
    Dim lngRetuId As Long                                        '��������������Ӧ�����ķ���ֵ
    Dim Pane1 As Pane                                           '��������
    Dim rptRow As ReportRow                                     '�б��¼��
    Dim intLoop As Integer                                      '��ǰ��λ��
    Dim lngLodKey As Long                                       '�ɵ�ID
    
'    Me.SetFocus
    Select Case intDisposal
        Case mFileS.����
            Select Case mintEditState
            Case 1, 2                                                           '�Ǽ�,���ձ���
                '�ڱ�ǰ����б����������˼���Ƿ�Ϊͬһ�������
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = False Then Exit Function
                End If
                
                lngRetuId = mfrmRequest.ZlSave()
               
                If lngRetuId = 0 Then Exit Function
                
                If mintContinue = 0 Or Me.TabList.Selected.Index = 1 Then
                    '����������
                    Me.rptList.Tag = ""   '�����������ı��
                    mfrmRequest.ZlCancel: mlngKey = lngRetuId
                    

                    mintEditState = 0
                    If mlngMachineID > 0 Or mlngMachineID = -1 Then
                        InsertOneRecored lngRetuId, True
                    Else
                        Call RefreshData
                    End If
                    If Me.TabList.Item(1).Selected = True Then
                        Call RefreshData1
                    End If
                    '����д��΢�������������
                    Call ReportDisposal(mActR.д�벡��)
                    '���պ��Ƿ�����������
                    Call SampleDisposal(mActS.��������)
                    
'                    If Me.rptList.Visible Then
                        '11268 �������˲�������֮��û�����
                        '���Ѻ��մ����в�֧�ִ˹���,�����մ�����ȱ�ٺܶ���Ϣ,�������������
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            Call ReportDisposal(mActR.��˱���)
                        End If
'                    End If
                    
                Else
                    If Me.rptList.Tag = "" Then
                        '��һ������ʱ������б�
                        Me.rptList.Records.DeleteAll
                        Me.rptList.Tag = "Continue"
                    End If
                    '��Ӹ������ļ�¼���б���
                    InsertOneRecored lngRetuId, True
                    '���պ��Ƿ�����������
                    Call SampleDisposal(mActS.��������)
                    
                    If Me.rptList.Visible Then
                        '11268 �������˲�������֮��û�����
                        '���Ѻ��մ����в�֧�ִ˹���,�����մ�����ȱ�ٺܶ���Ϣ,�������������
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            Call ReportDisposal(mActR.��˱���)
                        End If
                    End If
                    
                    If Me.TabList.Selected.Index = 0 Then
                        '���ջ�Ǽ�
                        SampleDisposal IIf(mintEditState = 1, mActS.����, mActS.�Ǽ�)
                    Else
                        mintEditState = 0
                    End If
                End If
            Case 3                                                              '���º���
                '�ڱ�ǰ����б����������˼���Ƿ�Ϊͬһ�������
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = False Then Exit Function
                End If
                
                lngRetuId = mfrmRequest.ZlSave()
                mfrmRequest.ZlCancel
                If lngRetuId = 0 Then Exit Function
                mintEditState = 0
                Call RefreshData
            Case 4                                                              '�����
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = False Then Exit Function
                End If
                lngRetuId = mfrmRequest.ZlSave(mintEditState)
                If lngRetuId = 0 Then Exit Function
                mfrmRequest.ZlCancel
                If Me.TabList.Item(0).Selected = True Then
                    InsertOneRecored lngRetuId, False
                Else
                    Call RefreshData1
                End If
                mintEditState = 0

            Case 5                                                              '����༭
                If Val(Me.rptList.FocusedRow.Record(mCol.΢����걾).Value) <> 1 Then
                    If mfrmWrite.ZlSave() = True Then
                        mfrmWrite.ZlCancel
                        mintEditState = 0
                        Call mfrmWrite.zlRefresh(mlngKey)

                        '------��ú����
                        Dim strSQL As String, rsTmp As ADODB.Recordset, i As Integer
                        
                        If Not rptList.FocusedRow Is Nothing Then
                            strSQL = "Select Count(A.ID) - Sum(Decode(A.������, Null, 0, 1)) As �޽����¼,Count(A.ID) as ����� " & vbNewLine & _
                                    "From ������ͨ��� A" & vbNewLine & _
                                    "Where A.����걾id = [1]"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
                            If rsTmp.EOF Then
                                For i = 0 To rptList.Columns.Count - 1
                                   rptList.FocusedRow.Record.Item(i).BackColor = vbWhite
                                Next
                            Else
                                If Val("" & rsTmp.Fields("�޽����¼")) = 0 And Val("" & rsTmp.Fields("�����")) > 0 Then
                                    For i = 0 To rptList.Columns.Count - 1
                                        rptList.FocusedRow.Record.Item(i).BackColor = &HFDD6C6     '&HC0FFFF
                                    Next
                                Else
                                    For i = 0 To rptList.Columns.Count - 1
                                        rptList.FocusedRow.Record.Item(i).BackColor = vbWhite
                                    Next
                                End If
                            End If
                        End If
                        '---------------

                    End If
                Else
                    If mfrmWrite2.ZlSave() = True Then
                        mfrmWrite2.ZlCancel
                        mintEditState = 0
                        Call mfrmWrite2.zlRefresh(mlngKey)
                        '����д��΢�������������
                        Call ReportDisposal(mActR.д�벡��)
                    End If
                End If
            Case 6                      '��д��������
                If mfrmLabMicrobe3Report.ZlSave(mlngKey) = True Then
                    mintEditState = 0
                End If
            End Select
        
        Case mFileS.����
            Select Case mintEditState
                Case 1, 2, 3, 4
                    If mfrmRequest.ZlCancel() = False Then Exit Function
                    Me.rptList.Tag = ""         '�����������ı��
                    mintEditState = 0
                    If Me.TabList.Selected.Index = 0 Then
                        Call RefreshData
                    Else
                        Call RefreshData1
                    End If
                Case 5
                    If Val(Me.rptList.FocusedRow.Record(mCol.΢����걾).Value) <> 1 Then
                        If mfrmWrite.ZlCancel = False Then Exit Function
                        mintEditState = 0
                        Call mfrmWrite.zlRefresh(mlngKey)
                    Else
                        If mfrmWrite2.ZlCancel = False Then Exit Function
                        mintEditState = 0
                        Call mfrmWrite2.zlRefresh(mlngKey)
                    End If
                Case 7
                    If mfrmLabMicrobe3Report.ZlCancel = True Then
                        mintEditState = 0
                    End If
                Case Else
            End Select
            Me.rptList.Tag = ""    '�����������ı��
            mintEditState = 0
            
            If Not Me.rptList.FocusedRow Is Nothing Then
                InsertOneRecored mlngKey, False
            End If
            
            
                            
'            If Me.TabList.Item(0).Selected = True Then
'                Call RefreshData
'            Else
'                Call RefreshData1
'            End If
    End Select
    
    SaveDisposal = True
    
    On Error Resume Next
    Me.MousePointer = 0
    If Me.rptList.Tag = "" Then
        gintSelectFocus = 1
'        Me.cboʱ��.SetFocus
        If Me.TabList.Selected.Index = 0 Then
'            Me.rptList.SetFocus
        Else
'            Me.rptList1.SetFocus
        End If
    End If
    If mintEditState = 0 Then
        Call ShowRequest(False)
    End If
End Function
Private Function SampleRefuse(lngKey As Long) As Boolean
    '�걾ȡ������(�м���Ҫ����˫��ͨѶ)
    '����                   ����ҽ��ID
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset, strQrySQL As String
    Dim strDevices As String, aDevice() As String, strAdviceIDs As String, i As Integer
    Dim intType As Integer                      '�걾���:0=��ͨ��1=����
    Dim lngAdviceID As Long                     'ҽ��ID
    Dim intEmerge As Integer                    '�Ƿ�ʹ�ü����־

    If mlngKey = 0 Then Exit Function
    
    intEmerge = Val(zlDatabase.GetPara("����걾", 100, 1208, 0))
    
    On Error GoTo ErrHand

    Me.MousePointer = vbHourglass
    strAdviceIDs = "": strDevices = ""
    
    strSQL = "select distinct nvl(b.�걾���,0) as �걾���,a.id as ҽ��Id " & _
             " from ����ҽ����¼ a,����걾��¼ b " & _
             " where a.id = b.ҽ��ID and b.id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)

    If rs.EOF = True Then Exit Function
    
    If rs.BOF = False Then
        intType = rs("�걾���")
        lngAdviceID = rs("ҽ��ID")
            
    End If
    
    
    '����˫��ͨ��
    If blnComm Then
        strAdviceIDs = strAdviceIDs & "," & lngAdviceID
        
        strQrySQL = "Select Distinct ����ID From ����걾��¼ A,������Ŀ�ֲ� B" & _
            " Where B.ҽ��ID=[1] And B.�걾ID+0=A.ID"
        Set rs = zlDatabase.OpenSQLRecord(strQrySQL, Me.Caption, lngAdviceID)
        Do While Not rs.EOF
            If InStr(strDevices, "," & zlCommFun.Nvl(rs(0), 0)) = 0 Then
                strDevices = strDevices & "," & zlCommFun.Nvl(rs(0), 0)
            End If
            'CSBmk <Type the bookmark name here>
            rs.MoveNext
        Loop
    End If
    
    '����˫��ͨ��
    If blnComm Then
        If Len(strDevices) > 0 Then strDevices = Mid(strDevices, 2)
        If Len(strAdviceIDs) > 0 Then strAdviceIDs = Mid(strAdviceIDs, 2)
        
        
        aDevice = Split(strDevices, ",")
        For i = 0 To UBound(aDevice)
            SendSample WinsockC, WinsockC.LocalIP, CLng(Val(aDevice(i))), "", 0, strAdviceIDs, True, IIf(intEmerge = 1, 0, intType)
        Next
    End If
    Me.MousePointer = vbDefault
    
    strSQL = "ZL_����걾��¼_ȡ������(" & lngAdviceID & ")"
    zlDatabase.ExecuteProcedure strSQL, gstrSysName
        
    SampleRefuse = True
   
    Exit Function
    
ErrHand:
    
    Me.MousePointer = vbDefault
    If SampleRefuse = False Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
        
End Function

Private Sub TabList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intLoop As Integer
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    If Me.TabList.Item(1).Selected = True Then
        Me.TabCtlWindow.Item(0).Selected = True
        Me.TabCtlWindow.Item(2).Visible = False
        Me.TabCtlWindow.Item(3).Visible = False
        Me.TabCtlWindow.Item(4).Visible = False
        Me.TabCtlWindow.Item(5).Visible = False
        Me.TabCtlWindow.Item(6).Visible = False
        For intLoop = 2 To Me.chkSoure.UBound
            Me.chkSoure(intLoop).Visible = False
        Next
        '��쵥������
        Me.chkSoure(5).Visible = True
        Me.chkSoure(5).Left = Me.chkSoure(2).Left
        Me.picFilter.Left = Me.chkSoure(2).Left + Me.chkSoure(2).Width + 30
        Call mfrmWrite.zlRefresh(-1)
        Call mfrmRequest.ZlCancel
        Call RefreshData1
        
        If mblnTabList1 = False Then
            cboʱ��.Text = "��  ��"
            mblnTabList1 = True
        Else
            cboʱ��.Text = Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(0)
            Me.dtpDate.Value = Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
            Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
        End If
        
        
        Call SetControlFocus
    Else
        Me.rptList.Tag = ""
        Me.TabCtlWindow.Item(0).Visible = True
        Me.TabCtlWindow.Item(0).Selected = True
        
        Me.TabCtlWindow.Item(2).Visible = True
        Me.TabCtlWindow.Item(3).Visible = IIf(Me.TabCtlWindow.Item(3).Tag = "���ò�ѯ", True, False)
        Me.TabCtlWindow.Item(4).Visible = True
        Me.TabCtlWindow.Item(5).Visible = True
        If Me.rptList.FocusedRow Is Nothing Then
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = False
        Else
            With Me.rptList
                If .Records(mCol.�������).Visible = "סԺ" Then
                    Me.TabCtlWindow.Item(6).Visible = False
                    Me.TabCtlWindow.Item(7).Visible = True
'                    Me.TabCtlWindow.Item(6).Selected = True
                Else
                    Me.TabCtlWindow.Item(6).Visible = True
                    Me.TabCtlWindow.Item(7).Visible = False
'                    Me.TabCtlWindow.Item(5).Selected = True
                End If
            End With
        End If
        Me.chkSoure(5).Left = 3780
        For intLoop = 0 To Me.chkSoure.UBound
            Me.chkSoure(intLoop).Visible = True
        Next
        Me.picFilter.Left = Me.chkSoure(5).Left + Me.chkSoure(5).Width + 30
'        Call RefreshData
        Call mfrmWrite.zlRefresh(mlngKey)
        Call mfrmRequest.zlRefresh(Me.rptList.FocusedRow)
        
        If mblnTabList1 = False Then
            cboʱ��.Text = "��  ��"
            mblnTabList1 = True
        Else
            cboʱ��.Text = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0)
            Me.dtpDate.Value = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
            Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
        End If
        
        Call SetControlFocus
    End If
    
    If mintContinue = 1 Then
        Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "�����Ǽ�"
        Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "��������"
    Else
        Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "�Ǽ�"
        Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "����"
    End If
    Me.cbrthis.RecalcLayout
    Call picList_Click
End Sub

Private Sub txtGoto_GotFocus()
    Me.txtGoto.SelStart = 0
    Me.txtGoto.SelLength = Len(Me.txtGoto.Text)
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If FindPatient(txtGoto.Text) = True Then
'            txtGoto.Text = ""
'        Else
            Call FindPatient(txtGoto.Text)
            txtGoto.SelStart = 0
            txtGoto.SelLength = Len(txtGoto.Text)
'        End If
    End If
End Sub

Private Sub RptListFilter()
    '����                   �б����(���ﲡ��;סԺ����;�����걾;����걾;δ��걾;��첡��;����ҽ��;�����걾 ���п��ٹ���)
    Dim lngloop As Long
    Dim strSource As String             '��Դ
    Dim strExeState As String           '����״̬
    Dim strPatientName As String        '����
    Dim lngItemID As Long               '������ĿID
    Dim lngCboItemID As Long            'ѡ��������Ŀ
    Dim intҽ������ As Integer          'ҽ������
    Dim int�걾���� As Integer          '�걾����
    Dim str���δͨ�� As String         '���δͨ��
    Dim rsTmp As New ADODB.Recordset    '��¼��
    Dim lngRowIndex As Long                                         '������
    Dim lngRowID As Long                                            '��ID
    Dim intLoop As Integer
    Dim i As Integer
    Dim str�걾���� As String
    Dim int���Ϊ�� As Integer
    Dim strYiqiShenHe As String
    On Error Resume Next
    
    'ˢ��ǰ��¼һ��λ��
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For i = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(i).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(i).Record.Index
                    lngRowID = Me.rptList.Rows(i).Record(mCol.ID).Value
                End If
            Next
        End If
    End If
        
'    If Me.rptList.Records.Count <= 0 And Me.rptList1.Records.Count <= 0 Then                           'û�м�¼ʱ�˳�
'        Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList.Rows.Count & "�����ˣ�"
'        Exit Sub
'    End If
    If Me.TabList.Selected.Index = 0 Then
        If Me.rptList.Records.Count <= 0 Then
            Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList.Rows.Count & "�����ˣ�"
            Exit Sub
        End If
    Else
        If Me.rptList1.Records.Count <= 0 Then
            Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList1.Rows.Count & "�����ˣ�"
            Exit Sub
        End If
    End If
    With Me.rptList
        For lngloop = 0 To .Records.Count - 1
            .Records(lngloop).Visible = True
            
            strSource = .Records(lngloop).Item(mCol.�������).Value
            strExeState = .Records(lngloop).Item(mCol.ִ��״̬).Value
            strPatientName = .Records(lngloop).Item(mCol.����).Value
            intҽ������ = Val(.Records(lngloop).Item(mCol.ҽ������).Value)
            int�걾���� = Val(.Records(lngloop).Item(mCol.�걾����).Value)
            str�걾���� = Nvl(.Records(lngloop).Item(mCol.�걾����).Value)
            str���δͨ�� = Nvl(.Records(lngloop).Item(mCol.���δͨ��).Value)
            int���Ϊ�� = Val(.Records(lngloop).Item(mCol.���Ϊ��).Value)
            strYiqiShenHe = .Records(lngloop).Item(mCol.�������).Value
            If str���δͨ�� = "" Then
                .Records(lngloop).Visible = mblnVerifying(9) And .Records(lngloop).Visible
            End If

            If str���δͨ�� <> "" Then
                .Records(lngloop).Visible = mblnVerifying(10) And .Records(lngloop).Visible
            End If

            
            '====����
            If strSource = "����" Or strSource = "Ժ��" Then
                .Records(lngloop).Visible = mblnVerifying(0) And .Records(lngloop).Visible
            End If
            
            '=====���
            If strSource = "���" Then
                .Records(lngloop).Visible = mblnVerifying(5) And .Records(lngloop).Visible
            End If
            
            '====סԺ
            If strSource = "סԺ" Then
                .Records(lngloop).Visible = mblnVerifying(1) And .Records(lngloop).Visible
            End If
            
            '====����
            If strSource = "����" Then
                .Records(lngloop).Visible = mblnVerifying(2) And .Records(lngloop).Visible
            End If
            
            '==ҽ������
            If intҽ������ = 1 Then
                .Records(lngloop).Visible = mblnVerifying(6) And .Records(lngloop).Visible
            End If
            
            If int�걾���� = 1 Then
                .Records(lngloop).Visible = mblnVerifying(7) And .Records(lngloop).Visible
            End If
            
            If str�걾���� = "�ʿ�" Then
                .Records(lngloop).Visible = mblnVerifying(8) And .Records(lngloop).Visible
            End If
            
            If str���δͨ�� = "" Then
                .Records(lngloop).Visible = mblnVerifying(9) And .Records(lngloop).Visible
            End If

            If str���δͨ�� <> "" Then
                .Records(lngloop).Visible = mblnVerifying(10) And .Records(lngloop).Visible
            End If
            
            
            '�����
            If strExeState = "�Ѽ���" Or strExeState = "�Ѵ�ӡ" Then
                .Records(lngloop).Visible = (mblnVerifying(3) = True And .Records(lngloop).Visible = True)
                
            End If
            
            If strExeState = "�Ѵ���" Or strExeState = "" Then
                .Records(lngloop).Visible = (mblnVerifying(4) = True And .Records(lngloop).Visible = True)
            End If
            
            'δ��ɵı걾�Ƿ���ʾ by cd 2014-01-08
            If int���Ϊ�� > 0 Then
                .Records(lngloop).Visible = (mblnVerifying(11) = True And .Records(lngloop).Visible = True)
            Else
                .Records(lngloop).Visible = (mblnVerifying(12) = True And .Records(lngloop).Visible = True)
            End If
            If strYiqiShenHe = "��" Then
                .Records(lngloop).Visible = mblnVerifying(13) And .Records(lngloop).Visible
            End If
            
            If strYiqiShenHe = "��" Then
                .Records(lngloop).Visible = mblnVerifying(14) And .Records(lngloop).Visible
            End If
        Next
        .Populate
        If Me.rptList.Rows.Count = 0 Then
            mfrmRequest.ZlCancel
            mfrmWrite2.ZlCancel
            mfrmWrite.ZlCancel
        End If
        Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList.Rows.Count & "�����ˣ�"
    End With
    
    '���¶�λ����ǰ��λ��
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If
    
    
    With Me.rptList1
        If Me.TabList.Item(1).Selected = False Then Exit Sub
        If Me.rptList1.Records.Count <= 0 Then
            Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList1.Rows.Count & "�����ˣ�"
            Exit Sub
        End If
        For lngloop = 0 To .Records.Count - 1
            strSource = .Records(lngloop).Item(mRCol.��Դ).Value
            .Records(lngloop).Visible = True
            '====����
            If strSource = "����" Or strSource = "Ժ��" Then
                .Records(lngloop).Visible = mblnWaitVerify(0) And .Records(lngloop).Visible
            End If
            '====סԺ
            If strSource = "סԺ" Then
                .Records(lngloop).Visible = mblnWaitVerify(1) And .Records(lngloop).Visible
            End If
            '====���
            If strSource = "���" Then
                .Records(lngloop).Visible = mblnWaitVerify(2) And .Records(lngloop).Visible
            End If
        Next
        .Populate

        If mlngMachineID = 0 Then Exit Sub
        If mlngMachineID = -1 Then
            gstrSql = "Select Distinct b.������Ŀid" & vbNewLine & _
                      " From ����������Ŀ a, ���鱨����Ŀ b, ������ĿĿ¼ c" & vbNewLine & _
                      " Where a.��Ŀid = b.������Ŀid And b.������Ŀid = c.Id"
        Else
            gstrSql = "Select Distinct b.������Ŀid" & vbNewLine & _
                      " From ����������Ŀ a, ���鱨����Ŀ b, ������ĿĿ¼ c" & vbNewLine & _
                      " Where a.����id = [1] And a.��Ŀid = b.������Ŀid And b.������Ŀid = c.Id"
        End If
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngMachineID)

        For lngloop = 0 To .Records.Count - 1
            If .Records(lngloop).Visible = True Then
                .Records(lngloop).Visible = True
                lngItemID = Val(.Records(lngloop).Item(mRCol.������ĿID).Value)
                rsTmp.filter = ""
                rsTmp.filter = "������ĿID = " & lngItemID
                If mlngMachineID = -1 Then
                    If rsTmp.RecordCount > 0 Then .Records(lngloop).Visible = False
                Else
                    If rsTmp.RecordCount <= 0 Then .Records(lngloop).Visible = False
                    lngCboItemID = cboUnionItem.ItemData(cboUnionItem.ListIndex)
                    If lngCboItemID = 0 Then
                    
                    ElseIf lngCboItemID = -1 And .Records(lngloop).Item(mRCol.������ĿID).Value = "" Then
                    
                    ElseIf Val(.Records(lngloop).Item(mRCol.������ĿID).Value) = lngCboItemID Then
                        
                    Else
                        .Records(lngloop).Visible = False
                    End If
                End If
            End If
        Next


        .Populate
        Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptList1.Rows.Count & "�����ˣ�"
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    

End Sub

Private Sub mclsExpenses_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.stbThis.Panels(2).Text = Text
End Sub
Private Sub mfrmWrite_StartEdit(Cancel As Boolean)
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    On Error GoTo errH:
    If InStr(",7,8,13,", CStr(Me.rptList.FocusedRow.Record(mCol.ִ��״̬).Icon)) > 0 Then
        '�Ѽ���
        Cancel = True
        mintHandleState = 0
    Else
        '���ڽ��еǼǺ��ղ���ʱ�Զ�����
        If mintEditState >= 1 And mintEditState <= 4 Then
            If Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Visible = True Then
                Call SaveDisposal(mFileS.����)
            End If
        End If
        Select Case Me.rptList.FocusedRow.Record(mCol.�걾����).Icon
            Case 3
                If InStr(mstrPrivs, "�޸��ʿؽ��") = 0 Then
                    Cancel = True
                    mintHandleState = 0
                Else
                    If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
                        Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
                        Cancel = False
                        ReportDisposal mActR.��д����
                    End If
                End If
            Case 4
                If InStr(mstrPrivs, "�޸ıȶԽ��") = 0 Then
                    Cancel = True
                    mintHandleState = 0
                Else
                    Cancel = False
                    If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
                        Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
                        ReportDisposal mActR.��д����
                    End If
                End If
            Case Else
        '        mintHandleState = 2
                If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
                        Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
                    ReportDisposal mActR.��д����
                    Cancel = False
                Else
                    Cancel = True
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AuditionCheck() As Boolean
    Dim strVerifyMan As String
    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            strVerifyMan = .Record(mCol.������).Value
        End With
    End If

    If InStr(1, mstrPrivs, "��˱걾") <= 0 Then
        'û��Ȩ�޺������û���½ʱ�˳�
        MsgBox "��û��Ȩ�޽������,�����µ�½���������Ա�������!", vbInformation, gstrSysName
        Call SetControlFocus
        gintSelectFocus = 1
        Exit Function
    End If

    '��Ȩ�޿���ʱ
    If InStr(1, mstrPrivs, "�������") > 0 And strVerifyMan = UserInfo.���� Then
        'û�е�½�����
        If mintAuditing = 0 Then
            'ͬһ���˱�Ȩ�޿��Ʋ��ܽ������
'            MsgBox "�����˺������Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
            Exit Function
        End If
        '���ʱ���Ƿ����
        If mintAuditing < 0 Then
            If DateDiff("h", mDataAuditing, Now) > Abs(mintAuditing) Then
'                MsgBox "�����Чʱ���ѹ�,�����µ�½�����!", vbInformation, gstrSysName
                '����Чʱ����ڿ��Խ������
                Exit Function
            End If
        End If
        
        '�жϵ�½ʱ���������Ƿ�Ϊͬһ��.
        If strVerifyMan = mstrAuditingMan Then
            '��½���������˺͵�ǰ�û�Ϊͬһ����
'            MsgBox "��½���������˺͵�ǰ�û�Ϊͬһ����,��ʹ�������û���½����!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    AuditionCheck = True
    
End Function


Private Sub ShortWork(SWork As Integer)
    Dim lngRetuId As Long                       '���ؽ��

    '����           ��ݲ���
    Select Case SWork
        Case mSWork.Key_Home, mSWork.Key_End
            Select Case mintEditState
                Case 0
                    If SWork = mSWork.Key_Home Then
                        BackOrNextPatient 1
                    Else
                        BackOrNextPatient 2
                    End If
                Case 1, 4, 5
                    If SaveDisposal(mFileS.����) = True Then
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            If AuditionCheck = False Then
                                Exit Sub
                            End If
                            Call ReportDisposal(mActR.��˱���)
                        Else
                            If SWork = mSWork.Key_End Then
                                If MoveStation(1, 2) = False Then                       '�����ƶ�
                                    'û���ҵ���¼ʱ�˳�����
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.����)
                                    Exit Sub
                                End If
                            Else
                                If MoveStation(0, 2) = False Then                      '�����ƶ�
                                    'û���ҵ���¼ʱ�˳�����
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.����)
                                    Exit Sub
                                End If
                            End If
                        End If
                        If mintHandleState = 1 Then
                            Call SampleDisposal(mActS.�����)
                        Else
                            Call ReportDisposal(mActR.��д����)
                        End If
                    End If
            End Select
        Case mSWork.Key_PageDown, mSWork.Key_PageUP
            Select Case mintEditState
                Case 0
                    If SWork = mSWork.Key_PageUP Then
                        BackOrNextPatient 1
                    Else
                        BackOrNextPatient 2
                    End If
                Case 1, 2                    '�Ǽ�ʱҲ����
                    SaveDisposal (mFileS.����)
                Case 4                     '�����
                    If SaveDisposal(mFileS.����) = True Then
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = False Then
                            Call ReportDisposal(mActR.��д����)
                        Else
                            If SWork = mSWork.Key_PageDown Then
                                If MoveStation(1, 2) = False Then                       '�����ƶ�
                                    'û���ҵ���¼ʱ�˳�����
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.����)
                                    Exit Sub
                                End If
                            Else
                                If MoveStation(0, 2) = False Then                       '�����ƶ�
                                    'û���ҵ���¼ʱ�˳�����
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.����)
                                    Exit Sub
                                End If
                            End If
                            Call SampleDisposal(mActS.�����)
                        End If
                    End If
                    
                Case 5                      '��д����
                    If SaveDisposal(mFileS.����) = True Then
                        '��������
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            If AuditionCheck = False Then
                                Exit Sub
                            End If
                            Call ReportDisposal(mActR.��˱���)
                        Else
                            If SWork = mSWork.Key_PageDown Then
                                If MoveStation(1, 2) = False Then                       '�����ƶ�
                                    'û���ҵ���¼ʱ�˳�����
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.����)
                                    Exit Sub
                                End If
                            Else
                                If MoveStation(0, 2) = False Then                       '�����ƶ�
                                    'û���ҵ���¼ʱ�˳�����
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.����)
                                    Exit Sub
                                End If
                            End If
                        End If
                        If mintHandleState = 1 Then
                            Call SampleDisposal(mActS.�����)
                        Else
                            Call ReportDisposal(mActR.��д����)
                        End If
                        
                    End If
            End Select
    End Select
    
End Sub
Private Function MoveStation(BackOrNext As Integer, Optional intState As Integer) As Boolean
    '����               �ƶ�����һ������һ����¼
    '����               BackOrNext =0 ���� = 1 ����
    '                   intState ����һ����״̬ 0 = ��һ��δ��˼�¼ = 1 ��һ������ =2 ��һ��
    
    Dim NowRow As Long
    Dim lngloop As Long

    If Me.rptList.Rows.Count = 0 Then Exit Function
    If Me.rptList.FocusedRow Is Nothing Then Exit Function

    NowRow = Me.rptList.FocusedRow.Index
    

    With Me.rptList

        If BackOrNext = 1 Then
            If NowRow + 1 = .Rows.Count Then Exit Function
            For lngloop = NowRow + 1 To .Rows.Count - 1
                If intState = 0 Then
                    If Val(.Rows(lngloop).Record(mCol.ҽ��id).Value) > 0 And .Rows(lngloop).Record(mCol.�����).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 1 Then
                    If .Rows(lngloop).Record(mCol.����).Value = "" And .Rows(lngloop).Record(mCol.�걾����).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 2 Then
                    Set .FocusedRow = .Rows(lngloop)
                    .Populate
                    MoveStation = True
                    Exit Function
                End If
            Next
        Else
            If NowRow - 1 = -1 Then Exit Function
            For lngloop = NowRow - 1 To 0 Step -1
                If intState = 0 Then
                    If Val(.Rows(lngloop).Record(mCol.ҽ��id).Value) > 0 And .Rows(lngloop).Record(mCol.�����).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 1 Then
                    If .Rows(lngloop).Record(mCol.����).Value = "" And .Rows(lngloop).Record(mCol.�걾����).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 2 Then
                    Set .FocusedRow = .Rows(lngloop)
                    .Populate
                    MoveStation = True
                    Exit Function
                End If
            Next
        End If
    End With

    
End Function
Public Sub zlRefreshData()
    'ˢ������
    Call RefreshData
End Sub
Private Sub RefreshData1()
    '''''''''''''''''''''''''''''''''''''''''
    '����           ˢ�������б�
    '''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim strSQL As String
    Dim lngAdviceID As Long
    Dim lngCorrelation As Long
    Dim intLoop As Integer
    Dim strStart As String
    Dim strEnd As String
    Dim strDeptID As String
    On Error GoTo errH

    
    strStart = GetDateTime(Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(0), 2)
    
    If strStart = "�Զ���" Then
        strStart = Format(Me.dtpDate.Value, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd.Value, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    End If
    Me.rptList1.Records.DeleteAll
    
    gstrSql = "Select /*+ Rule */ a.id,a.���ID,a.����id,a.������־,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') As ������Դ," & vbNewLine & _
            "       d.����,d.�Ա�,d.����,e.���� As ���˿���," & vbNewLine & _
            "       decode(a.������Դ,1,d.�����,2,d.סԺ��,4,d.�����) As ��ʶ��," & vbNewLine & _
            "       Decode(a.������Դ,2,S.��������,d.��������) as ��������," & vbNewLine & _
            "       d.��ǰ���� As ����,a.ҽ������, a.����ҽ�� , a.����ʱ��,a.������ĿID,b.ִ��״̬,a.�Һŵ�,b.����ʱ�� " & vbNewLine & _
            "From ����ҽ����¼ a , ����ҽ������ b ,������ĿĿ¼ c , ������Ϣ d ,���ű� e,������ҳ s" & vbNewLine & _
            "Where a.Id = b.ҽ��id And b.ִ��״̬ in (0,2) And a.���ID Is Not Null" & vbNewLine & _
            "     And c.Id = a.������ĿID And c.��� = 'C' And a.����ʱ�� Between [2] And [3]" & vbNewLine & _
            "     And a.����id = d.����id And a.���˿���id = e.Id [����]  " & vbNewLine & _
            " and a.����ID = S.����ID(+) and a.��ҳID = s.��ҳID(+)  " & vbNewLine & _
            "Order By a.Id , a.���ID ,����ʱ�� "
    
    
    If mlngDeptID > 0 And rptList.Tag = "" Then
        strDeptID = mlngDeptID
    Else
        If InStr(mstrPrivs, "���п���") = 0 Or InStr(mstrPrivs, "�鿴�������ұ���") > 0 Then
            For intLoop = 1 To Me.cboDept.ListCount - 1
                strDeptID = strDeptID & "," & Me.cboDept.ItemData(intLoop)
            Next
        End If
    End If
    gstrSql = Replace(gstrSql, "[����]", " And A.ִ�п���id In (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) ")

    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strDeptID, CDate(Format(strStart, "yyyy-MM-dd HH:mm:ss")), _
                                        CDate(Format(strEnd, "yyyy-MM-dd HH:mm:ss")))
    
    Do Until rsTmp.EOF
        With Me.rptList1
            If rsTmp("���ID") <> lngCorrelation Then
                Set Record = .Records.Add
                For intLoop = 0 To .Columns.Count
                    Record.AddItem ""
                Next
                Record(mRCol.����ID).Value = Nvl(rsTmp("����ID"))
                Record(mRCol.����).Icon = IIf(Val(Nvl(rsTmp("������־"))) = 1, 0, -1)    '1=����
                Record(mRCol.��Դ).Value = Nvl(rsTmp("������Դ"))
                Record(mRCol.����).Value = Nvl(rsTmp("����")) & IIf(Nvl(rsTmp("ִ��״̬")) = 2, "(����)", "")
                Record(mRCol.����).Value = Nvl(rsTmp("����"))
                Record(mRCol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
                Record(mRCol.���˿���).Value = Nvl(rsTmp("���˿���"))
                Record(mRCol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
                Record(mRCol.����).Value = Nvl(rsTmp("����"))
                Record(mRCol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
                Record(mRCol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
                Record(mRCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mRCol.������ĿID).Value = Nvl(rsTmp("������ĿID"))
                Record(mRCol.ҽ��id).Value = Nvl(rsTmp("���ID"))
                Record(mRCol.ִ��״̬).Value = Nvl(rsTmp("ִ��״̬"))
                Record(mRCol.�Һŵ�).Value = Nvl(rsTmp("�Һŵ�"))
                Record(mRCol.ǩ��ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                If Nvl(rsTmp("��������")) <> "" Then
                    Record(mRCol.����).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp("��������")), False)
                End If
            Else
                Record(mRCol.ҽ������).Value = Record(mRCol.ҽ������).Value & " " & Nvl(rsTmp("ҽ������"))
            End If
            If Nvl(rsTmp("ִ��״̬")) = 2 Then
                For intLoop = 0 To .Columns.Count
                    Record(intLoop).ForeColor = vbRed
                Next
            End If
            lngCorrelation = Val(Nvl(rsTmp("���ID")))
        End With
        rsTmp.MoveNext
    Loop
    If Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
    Else
'        Me.rptList1.SetFocus
    End If
    Me.rptList1.Populate
    Call RptListFilter
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub QuickFindPatient()
    '���ٲ��ҵ�ǰ���˵����μ���
    Dim strPatient As String                        '��������
    Dim lngPatientID As Long                        '����ID
    Dim strStart As String                          '��ʼʱ��
    Dim strEnd As String                            '����ʱ��
    Dim i As Integer                                '�����ж��Ƿ�ʹ�ò���ID
    
    On Error Resume Next
    If Me.TabList.Item(0).Selected = False Or Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    With Me.rptList.FocusedRow
        If .Record(mCol.����).Value = "" Then Exit Sub
        strPatient = .Record(mCol.����).Value
        lngPatientID = .Record(mCol.����ID).Value
        i = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0)
        strStart = GetDateTime(zlDatabase.GetPara("���μ��鷶Χ", 100, 1208, "��  ��"), 1)
        strEnd = GetDateTime(zlDatabase.GetPara("���μ��鷶Χ", 100, 1208, "��  ��"), 2)
        Me.rptList.Tag = strPatient & ";;,;;;;;;;,True;" & strStart & "," & strEnd & ";0;;0;;;;1;;" & IIf(i = 0, lngPatientID, "0")
        Call RefreshData
    End With
    
End Sub






Private Sub DelItem(lngKey As Long)
    '����           'ɾ��ָ���ļ�¼
    Dim intLoop As Integer
    Dim lngloop As Integer
    Dim lngRowIndex As Long                                         '������
    Dim lngRowID As Long                                            '��ID
    
    
    'ˢ��ǰ��¼һ��λ��
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For intLoop = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(intLoop).Record.Index
                    lngRowID = Me.rptList.Rows(intLoop).Record(mCol.ID).Value
                End If
            Next
        End If
    End If
    
    With Me.rptList
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mCol.ID).Value = lngKey Then
                .Records.RemoveAt (intLoop)
                .Populate
                Exit For
            End If
        Next
    End With
    
    '���¶�λ����ǰ��λ��
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If
End Sub
Public Function ReadImageData(lngKeyID As Long, blnSave As Boolean) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer
    Dim strTime As Date
    Dim strErr As String
    Static objImg As Object
        
    On Error GoTo errH
    strTime = Now
    gstrSql = "select id ,�걾ID,ͼ������ from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKeyID)
    'ͼ���Ű�
    ImageTypeSet rsTmp.RecordCount - 1, True
    '����ʾʱ������
    If Me.cbrthis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = True Then Exit Function
    
    If objImg Is Nothing Then Set objImg = CreateObject("zlLisDev.clsDrawGraph")
    objImg.GetSampleImgInit glngSys, gcnOracle, strErr
    Call objImg.GetSampleImages(lngKeyID, App.path, False, strErr)
    Do Until rsTmp.EOF
        If Dir(App.path & "\" & lngKeyID & "_" & rsTmp("ͼ������") & ".cht") = "" Then
            If Dir(App.path & "\" & lngKeyID & "_" & rsTmp("ͼ������") & ".cht") <> "" Then
                Me.ChartThis(DrawIndex).Load App.path & "\" & lngKeyID & "_" & rsTmp("ͼ������") & ".cht"
                If blnSave Then
                    Kill App.path & "\" & lngKeyID & "_" & rsTmp("ͼ������") & ".cht"
                End If
            End If
        Else
            Me.ChartThis(DrawIndex).Load App.path & "\" & lngKeyID & "_" & rsTmp("ͼ������") & ".cht"
            If blnSave Then
                Kill App.path & "\" & lngKeyID & "_" & rsTmp("ͼ������") & ".cht"
            End If
        End If
        DrawIndex = DrawIndex + 1
        rsTmp.MoveNext
    Loop
    ReadImageData = True
'    Debug.Print "ID=" & lngKeyID & ",��ʱ:" & DateDiff("s", strTime, Now)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub VScroll1_Change()
'    Me.PicImage.Top = -200
End Sub

Private Sub VScroll_Change()
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
    For intLoop = 0 To Me.VScroll.Max
        If intLoop < Me.VScroll.Value Then
            Me.ChartThis(intLoop).Visible = False
        Else
            Me.ChartThis(intLoop).Visible = True
            If intLoop = Me.VScroll.Value Then
                Me.ChartThis(intLoop).Top = 0
            Else
                Me.ChartThis(intLoop).Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
            End If
        End If
    Next
End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnReset As Boolean = False)
    '����           �Լ���ͼ������Ű�
    '����           intCount = ͼ����
    '               blnReset = �Ƿ���Ҫ���¶���
    Dim intLoop As Integer
    Dim Pane5 As Pane
    
'    If blnReset = True Then
'        For intLoop = Me.ChartThis.UBound To 1 Step -1
''            Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
''            Me.ChartThis(Me.ChartThis.UBound).Header.Text = ""
'            If intLoop <> 0 Then
'                Unload Me.ChartThis(Me.ChartThis.UBound)
'            End If
'        Next
'    End If
    
    On Error Resume Next
    
    For intLoop = 0 To intCount
        If intLoop = 0 Then
            With Me.ChartThis(intLoop)
                .Interior.Image.LayOut = oc2dImageStretched
                .Visible = True
                .Top = 0
                .Left = 0
                .Width = IIf(Me.PicImage.ScaleWidth - Me.VScroll.Width - 20 <= 300, 300, Me.PicImage.ScaleWidth - Me.VScroll.Width - 20)
                .Height = .Width
            End With
        Else
            If blnReset = True And Me.ChartThis.UBound < intLoop Then
                Load Me.ChartThis(intLoop)
            End If
            With Me.ChartThis(intLoop)
'                .ChartGroups(1).Data.NumSeries = intLoop
'                .ChartGroups(1).Data.NumPoints(intLoop) = intLoop
                .Interior.Image.LayOut = oc2dImageStretched
                .Visible = True
                .Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
                .Left = 0
                .Width = Me.ChartThis(intLoop - 1).Width
                .Height = .Width
                .IsBatched = False
            End With
        End If
    Next
    
    '���ض����Chart�ؼ�
    For intLoop = intCount + 1 To Me.ChartThis.UBound
        Me.ChartThis(intLoop).Visible = False
    Next
    
    Set Pane5 = Me.dkpMain.FindPane(Dkp_ID_Image)
    If Not Pane5 Is Nothing Then
        If intCount < 0 Then
            Pane5.Close
        Else
            If Me.cbrthis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = False Then
                Pane5.Select
            Else
                Pane5.Close
            End If
        End If
    End If
    With Me.VScroll
        .Top = 0
        .Left = Me.PicImage.ScaleWidth - .Width - 10
        .Height = Me.PicImage.ScaleHeight
        .Max = intCount
        .SmallChange = 1
        .LargeChange = 1
    End With
End Sub


Private Function CheckPatientInfo(lngSampleID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim int��ʾ���� As Integer '1-��ʾ������2-����ʾ������3-������

    On Error GoTo errH
    
    gstrSql = "Select A.������Դ,A.����id, A.�Ա� As �Ա�1, B.�Ա� As �Ա�2, A.���� As ����1, B.���� As ����2, A.���� As ����1, B.���� As ����2,nvl(a.Ӥ��,0) as Ӥ�� " & vbNewLine & _
                        "From ����걾��¼ A, ������Ϣ B" & vbNewLine & _
                        "Where A.����id = B.����id And A.ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngSampleID)
    
    '��Ӥ��ʱ�����жԱ�
    If rsTmp("Ӥ��") > 0 Then
        Exit Function
    End If
    
    
    If Nvl(rsTmp("����1")) <> Nvl(rsTmp("����2")) Or Nvl(rsTmp("�Ա�1")) <> Nvl(rsTmp("�Ա�2")) Or _
        Nvl(rsTmp("����1")) <> Nvl(rsTmp("����2")) Then
        
        int��ʾ���� = 1
        
        If rsTmp("������Դ") = 4 Then
            int��ʾ���� = int��촦��ʽ
        ElseIf rsTmp("������Դ") = 3 Then
            int��ʾ���� = intԺ�⴦��ʽ
        ElseIf rsTmp("������Դ") = 2 Then
            int��ʾ���� = intסԺ����ʽ
        ElseIf rsTmp("������Դ") = 1 Then
            int��ʾ���� = int���ﴦ��ʽ
        End If
        
        If int��ʾ���� = 1 Then
            If MsgBox("���ּ�����Ϣ�еĲ�����Ϣ�Ͳ�����Ϣ�в�����Ϣ��һ��!" & vbCrLf & "�Ƿ���Ҫ����?", _
                vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                gstrSql = "zl_����걾��¼_Update(" & lngSampleID & ",'" & Nvl(rsTmp("����2")) & "','" & Nvl(rsTmp("�Ա�2")) & _
                                             "','" & Nvl(rsTmp("����2")) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            End If
        ElseIf int��ʾ���� = 2 Then
            gstrSql = "zl_����걾��¼_Update(" & lngSampleID & ",'" & Nvl(rsTmp("����2")) & "','" & Nvl(rsTmp("�Ա�2")) & _
                                         "','" & Nvl(rsTmp("����2")) & "')"
            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
        End If
        CheckPatientInfo = True
        Exit Function
    End If
    CheckPatientInfo = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub DelButton(Index As Integer)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    '����       ��ʾ�����ذ�ť
    Dim lngCount As Long
    On Error Resume Next
    '���ò�ѯ
    Me.cbrthis.FindControl(, conMenu_EditPopup).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Append).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_NewItem).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Modify).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Delete).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ChargeDelApply).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ChargeDelAudit).Delete
    Me.cbrthis.FindControl(, conMenu_ToolPopup).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_Option).Delete
    Me.cbrthis.FindControl(, conMenu_ToolPopup).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ExtraFeeMove).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ExtraFeeExe).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ExtraFeeUnExe).Delete
    'ҽ����¼
    Me.cbrthis.FindControl(, conMenu_Edit_NewItem).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Modify).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Delete).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Blankoff).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Stop).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Send).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Untread).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Compend).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_MarkMap).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_MarkKeyMap).Delete
    Me.cbrthis.FindControl(, conMenu_Manage_ReportLisView).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_Sign).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_SignNew).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_SignVerify).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_SignEarse).Delete
    Me.cbrthis.FindControl(, conMenu_View_Append, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_Hide, , True).Delete
    Me.cbrthis.FindControl(, conMenu_Report_ClinicBill, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_FontSize, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_FontSize_S, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_FontSize_L, , True).Delete
    
    '���û�� �󱨸� ��ť,������һ�� �󱨸� ��ť
    Set objBar = Me.cbrthis(2)
    With objBar.Controls
        Set objControl = .Find(, conMenu_Edit_Audit)
        If objControl Is Nothing Then
            Set objControl = .Find(, conMenu_Manage_Report) '����水ť֮��ʼ����
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�󱨸�", objControl.Index + 1): objControl.BeginGroup = True   '���Ӱ�ť
            objControl.ID = conMenu_Edit_Audit: objPopup.IconId = conMenu_Edit_Audit    '��ֵID��ͼ��
            objControl.Style = xtpButtonIconAndCaption  'ͬʱ��ʾ�����ͼƬ
        End If
    End With

'    With Me.cbrthis.KeyBindings
'        .Add FCONTROL, Asc("P"), conMenu_File_Print
'        .Add 0, VK_F2, conMenu_Edit_Save
'        .Add 0, VK_ESCAPE, conMenu_LIS_Cancel
'        .Add 0, VK_F12, conMenu_File_Parameter
'        .Add 0, VK_F4, conMenu_Manage_Plan
'        .Add 0, VK_F8, conMenu_Manage_Regist
'        .Add FCONTROL, Asc("T"), conMenu_Tool_Apply
'        .Add FCONTROL, Asc("Z"), conMenu_Edit_SendBack
'        .Add FCONTROL, VK_DELETE, conMenu_Manage_ClearUp
'        .Add 0, VK_F7, conMenu_Manage_Report
'        .Add 0, VK_F6, conMenu_Edit_Audit
'        .Add FCONTROL, VK_LEFT, conMenu_View_Backward
'        .Add FCONTROL, VK_RIGHT, conMenu_View_Forward
'        .Add 0, VK_F1, conMenu_Help_Help
'        .Add 0, VK_F5, conMenu_View_Refresh
'        .Add FCONTROL, Asc("F"), conMenu_Manage_Transfer_Force
'        .Add 0, VK_F3, conMenu_View_Filter
'        .Add 0, VK_HOME, conMenu_Tool_MeetFinish
'        .Add 0, VK_END, conMenu_Tool_MeetCancel
'        .Add 0, VK_PAGEUP, conMenu_Tool_Reference_1
'        .Add 0, VK_PAGEDOWN, conMenu_Tool_Reference_2
'        .Add FCONTROL, Asc("H"), conMenu_View_FindNext
'        .Add 0, VK_F9, conMenu_Edit_QCRes
'        .Add 0, VK_F11, conMenu_Manage_Logout
'    End With
'
'    If Index = 3 Then
'        'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
'        End If
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&M)", objMenu.Index + 1, False)
'        objMenu.ID = conMenu_EditPopup
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����������(&N)")
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "����Ѽӷ���(&A)"): objPopup.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ĸ��ӷ���(&M)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�����ӷ���(&D)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "��������(&L)"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "�������(&U)")
'        End With
'
'        '���߲˵�:���������û��,���ڰ����˵�ǰ��
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
'            objMenu.ID = conMenu_ToolPopup
'        End If
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ������ѡ��(&O)"): objControl.BeginGroup = True
'            objControl.IconId = conMenu_File_Parameter
'        End With
'
'        '����������:���ļ�������˵������ť֮��ʼ����
'        '-----------------------------------------------------
'        Set objBar = cbrthis(2)
'        For Each objControl In objBar.Controls '�����ǰ������һ��Control
'            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
'                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
'            End If
'        Next
'        With objBar.Controls
'            'Set objControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����", objControl.Index + 1): objControl.BeginGroup = True
'            Set objPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "����", objControl.Index + 1): objPopup.BeginGroup = True
'                objPopup.ID = conMenu_Edit_NewItem: objPopup.IconId = conMenu_Edit_NewItem
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�ķ�", objPopup.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", objControl.Index + 1)
'        End With
'
'        '����Ŀ����
'        '-----------------------------------------------------
'        With cbrthis.KeyBindings
'            .Add FCONTROL, vbKeyE, conMenu_Edit_Append '����������
'            .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸ĸ��ӷ���
'            .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ�����ӷ���
'        End With
'
'        '���ò���������
'        '-----------------------------------------------------
'        With cbrthis.Options
'        End With
'    End If
'
'    If Index = 5 Then
'        'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
'        End If
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
'        objMenu.ID = conMenu_EditPopup
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&G)"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)")
'
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
'            objPopup.BeginGroup = True
'            objPopup.IconId = conMenu_Manage_Report
'
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
'        End With
'
'        '�鿴�˵�
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        With objMenu.CommandBar.Controls
'            Set objControl = .Find(, conMenu_View_StatusBar) '״̬�����
'            Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
'        End With
'
'        '���߲˵�:���������û��,���ڰ����˵�ǰ��
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
'            objMenu.ID = conMenu_ToolPopup
'        End If
'        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
'                objControl.IconId = conMenu_Tool_Sign
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
'            End With
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "����ҽ��ѡ��(&O)"): objControl.BeginGroup = True
'            objControl.IconId = conMenu_File_Parameter
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&O)"): objControl.BeginGroup = True
'        End With
'
'        '����������:���ļ�������˵������ť֮��ʼ����
'        '-----------------------------------------------------
'        Set objBar = cbrthis(2)
'        For Each objControl In objBar.Controls '�����ǰ������һ��Control
'            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
'                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
'            End If
'        Next
'        With objBar.Controls
'            'Set objControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�¿�", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ǩ��", objControl.Index + 1): objControl.BeginGroup = True
'            objControl.IconId = conMenu_Tool_Sign
'        End With
'
'        '����Ŀ����
'        '-----------------------------------------------------
'        With cbrthis.KeyBindings
'            .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
'            .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
'            .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
'            .Add FCONTROL, vbKeyG, conMenu_Edit_Send 'ҽ������
'
'            .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '���ı���
'            .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
'
'            .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
'
'            .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
'        End With
'
'        '���ò���������
'        '-----------------------------------------------------
'        With cbrthis.Options
'        End With
'    End If
'
'    If Index = 6 Then
'        'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
'        End If
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
'        objMenu.ID = conMenu_EditPopup
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ҽ��ֹͣ(&S)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "��������(&G)"): objControl.BeginGroup = True
'            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Untread, "ҽ������(&L)")
'
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
'            objPopup.BeginGroup = True
'            objPopup.IconId = conMenu_Manage_Report
'
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
'        End With
'
'        '�鿴�˵�
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        With objMenu.CommandBar.Controls
'            Set objControl = .Find(, conMenu_View_StatusBar) '״̬�����
'            Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
'        End With
'
'        '���߲˵�:���������û��,���ڰ����˵�ǰ��
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
'            objMenu.ID = conMenu_ToolPopup
'        End If
'        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
'                objControl.IconId = conMenu_Tool_Sign
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
'            End With
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "סԺҽ��ѡ��(&O)"): objControl.BeginGroup = True
'            objControl.IconId = conMenu_File_Parameter
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&O)"): objControl.BeginGroup = True
'        End With
'
'        '����������:���ļ�������˵������ť֮��ʼ����
'        '-----------------------------------------------------
'        Set objBar = cbrthis(2)
'        For Each objControl In objBar.Controls '�����ǰ������һ��Control
'            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
'                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
'            End If
'        Next
'        With objBar.Controls
'            'Set objControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ǩ��", objControl.Index + 1): objControl.BeginGroup = True
'            objControl.IconId = conMenu_Tool_Sign
'        End With
'
'        '����Ŀ����
'        '-----------------------------------------------------
'        With cbrthis.KeyBindings
'            .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
'            .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
'            .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
'            .Add FCONTROL, vbKeyS, conMenu_Edit_Stop 'ֹͣҽ��
'            .Add FCONTROL, vbKeyG, conMenu_Edit_Send 'ҽ������
'            .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread 'ҽ������
'
'            .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '���ı���
'            .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
'
'            .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
'
'            .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
'        End With
'
'        '���ò���������
'        '-----------------------------------------------------
'        With cbrthis.Options
'        End With
'    End If
'    cbrthis.ActiveMenuBar.FindControl(, conMenu_LIS_RightMenu).Visible = False
'    Me.cbrthis.RecalcLayout
    'ɾ�����ڵĹ������������˵���
'    For lngCount = cbrthis.ActiveMenuBar.Controls.Count To 1 Step -1
'        cbrthis.ActiveMenuBar.Controls(lngCount).Delete
'    Next
'    For lngCount = cbrthis.Count To 2 Step -1
'        cbrthis(lngCount).Delete
'    Next
'    Call CreateCbs
End Sub

Private Sub WinsockC_DataArrival(ByVal bytesTotal As Long)
    '********************���ظ���ʦվ����Ϣ*****************************
    'Private Const strSend_Refresh = "Refresh"      '�ѱ������ݿ���ˢ��
    'Private Const strSend_True = "True"            '�Ѳ����ɹ�
    'Private Const strSend_False = "False"          '����ʧ��
    '*******************************************************************
    Dim strData As String
    Dim astrData() As String
    
    On Error Resume Next
    
    With Me.WinsockC
        .GetData strData
        astrData = Split(strData, ";")

        Select Case astrData(1)
            Case "Refresh"
                If Me.Tag <> "Refresh" And blnAutoRefresh And mintEditState = 0 Then
                    Me.Tag = "Refresh"
                    Call InsertOneRecored(Val(astrData(2)), False, False)
                    Me.Tag = ""
                End If
            Case "True"
                mblnSendComplete = True
            Case "False"
                mblnSendComplete = False
            Case Else
                If strData Like "AutoQCCompute|*" Then
                    If Split(strData, "|")(1) <> "" Then frmQCShowInfo.ShowMe "�Զ�����", Split(strData, "|")(1), Me
                End If
        End Select
    End With
End Sub

Private Sub ShowRequest(blnShow As Boolean)
    '����       �Ƿ���ʾ�ǼǴ���
    '��ע       ����ѡ��ʱ�Ż���Ч
    Dim Pane1 As Pane
    Dim blnExec As Boolean
    blnExec = Val(zlDatabase.GetPara("ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���", 100, 1208, 0))
    If blnExec = False And blnShow = False Then Exit Sub    'û��ѡ�����ʱ������
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Request)
    If blnShow = True Then
        Pane1.Select
    Else
        Pane1.Close
    End If
    Me.dkpMain.RecalcLayout
End Sub
Private Sub GetVerifying()
    '����           �õ�������ɸѡ�ִ�
    Dim intLoop As Integer
    Dim astrFilter() As String
    
    astrFilter = Split(con_������ɸѡ_������, ";")
    For intLoop = 0 To UBound(astrFilter)
        mblnVerifying(intLoop) = zlDatabase.GetPara("������_" & astrFilter(intLoop), 100, 1208, True)
        If intLoop <= Me.chkSoure.UBound Then
            Me.chkSoure(intLoop).Value = IIf(mblnVerifying(intLoop), 1, 0)
        End If
    Next
End Sub
Private Sub GetWaitVerify()
    '����           �õ��ȴ�����ɸѡ�ִ�
    Dim intLoop As Integer
    Dim astrFilter() As String
    astrFilter = Split(con_������ɸѡ_������, ";")
    For intLoop = 0 To UBound(astrFilter)
        mblnWaitVerify(intLoop) = zlDatabase.GetPara("������_" & astrFilter(intLoop), 100, 1208, True)
        If intLoop < 2 Then
            Me.chkSoure(intLoop).Value = IIf(mblnWaitVerify(intLoop), 1, 0)
        Else
            Me.chkSoure(5).Value = IIf(mblnWaitVerify(intLoop), 1, 0)
        End If
    Next
End Sub
Private Sub CreateChildCbs()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom

    '�Ӵ���˵�����
    Me.cbrChild.VisualTheme = xtpThemeOffice2003
    Set Me.cbrChild.Icons = zlCommFun.GetPubIcons
    With Me.cbrChild.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
'        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .LargeIcons = False
    End With
    Me.cbrChild.EnableCustomization False

    Me.cbrChild.ActiveMenuBar.Title = "�˵�"
    Me.cbrChild.ActiveMenuBar.Position = xtpBarTop
    Me.cbrChild.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbrChild.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "ǰһ��")
        cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "��һ��")
        cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlLabel, 0, "    ��λ")
        Set cbrCustom = .Add(xtpControlCustom, conMenu_File_RoomSet, "")
        cbrCustom.Handle = Me.txtGoto.hWnd
        Me.txtGoto.ToolTipText = "����Ϊ�걾�ź����롢��������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
        
        Set cbrControl = .Add(xtpControlLabel, conMenu_Edit_UnArchive, "    �շ���Ŀ")
        Set cbrCustom = .Add(xtpControlCustom, conMenu_Manage_Transfer_Send, "")
        cbrCustom.Handle = Me.cboExesItem.hWnd
        
        Set cbrControl = .Add(xtpControlLabel, conMenu_View_FindType, "")
        Set cbrPopControl = .Add(xtpControlButtonPopup, 0, "ѡ��     ")
        cbrPopControl.Flags = xtpFlagRightAlign: cbrPopControl.Style = xtpButtonIconAndCaption
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_RequestView, "ʹ������ɨ��", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("ʹ������ɨ��", 100, 1208, False)
        
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_RequestPrint, "��������", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("��������", 100, 1208, False)
        
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_RequestBatPrint, "�����ֱ�����", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("�����ֱ�����", 100, 1208, True, 0)
        
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, XTP_ID_WINDOW_LIST, "��ʾ��ע", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("��ʾ���鱸ע", 100, 1208, False)
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, ""): cbrControl.Visible = False
        
        
        
    End With
    cbrChild.RecalcLayout
'    Call mclsExpenses.zlDefCommandBars(Me, Me.cbrthis)
'    Call mclsInAdvices.zlDefCommandBars(Me, Me.cbrthis, 2)
'    Call mclsOutAdvices.zlDefCommandBars(Me, Me.cbrthis, 2)
'    Call zldatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs)
End Sub

Private Sub SetControlFocus()
    On Error Resume Next
    If Me.Visible = False Or Me.TabList.Enabled = False Then Exit Sub
    If Me.TabList.Selected.Index = 0 Then
        Me.rptList.SetFocus
    Else
        Me.rptList1.SetFocus
    End If
End Sub

Private Sub PrintBarcord()
    Dim intBarCode As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str�ɼ���ʽ As String, strִ�п��� As String, str����ʱ�� As String
    Dim str���� As String, str��Ѫ�� As String, str������ As String
    Dim lng������Դ As Long
    '�������뵽PIC
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    With Me.rptList.FocusedRow
        If Trim(.Record(mCol.��������).Value) = "" Then Exit Sub
        '�ɼ���ʽ,ִ�п���,����ʱ��,����,��Ѫ��,�Թ�����,
        
        strSQL = "Select A.������Դ,D.���� As �ɼ���ʽ, F.���� As ִ�п���, To_Char(C.����ʱ��, 'yyyy-MM-dd HH24:mi:ss') As ����ʱ��, E.���� As ����, E.��Ѫ��," & vbNewLine & _
                "       E.���� As ������, A.������Դ " & vbNewLine & _
                "From ���ű� F, ��Ѫ������ E, ������ĿĿ¼ D, ����ҽ����¼ C, ����ҽ������ B, ����걾��¼ A" & vbNewLine & _
                "Where C.ִ�п���id = F.ID And D.�Թܱ��� = E.����(+) And C.������Ŀid = D.ID And C.ID = B.ҽ��id And A.�������� = B.�������� And A.ID = [1]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(mCol.ID).Value))
        Do Until rsTmp.EOF
            If Trim("" & rsTmp!����) = "" Then
                str�ɼ���ʽ = Trim("" & rsTmp!�ɼ���ʽ)
            Else
                strִ�п��� = Trim("" & rsTmp!ִ�п���)
                str����ʱ�� = Trim("" & rsTmp!����ʱ��)
                str���� = Trim("" & rsTmp!����)
                str��Ѫ�� = Trim("" & rsTmp!��Ѫ��)
                str������ = Trim("" & rsTmp!������)
                lng������Դ = Val(Trim("" & rsTmp!������Դ))
            End If
            rsTmp.MoveNext
        Loop
        intBarCode = zlDatabase.GetPara("ʹ������", "100", "1211", False, 2)
        If intBarCode = 1 Then
            Bar39 Me.picBarCodePrint, 3, CStr(Trim(.Record(mCol.��������).Value)), False, True
        Else
            Bar128 Me.picBarCodePrint, 3, CStr(Trim(.Record(mCol.��������).Value)), True
        End If
        SavePicture Me.picBarCodePrint.Image, App.path & "\BarCode.Bmp"
        '��ʼ��ӡ
        
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me, "��������=" & Trim(.Record(mCol.��������).Value), "��Ŀ=" & Trim(.Record(mCol.������Ŀ).Value), _
        "�������� = " & IIf(Trim(.Record(mCol.����).Value) <> "", Trim(.Record(mCol.����).Value) & IIf(Val(Trim(.Record(mCol.Ӥ��).Value)) = 0, "", "(Ӥ��" & Trim(.Record(mCol.����).Value) & ")"), "��"), _
        "�Ա� = " & IIf(Trim(.Record(mCol.�Ա�).Value) <> "", Trim(.Record(mCol.�Ա�).Value), "��"), _
        "���� = " & IIf(Trim(.Record(mCol.����).Value) & Trim(.Record(mCol.���䵥λ).Value) <> "", Trim(.Record(mCol.����).Value) & Trim(.Record(mCol.���䵥λ).Value), "��"), _
        "���� = " & IIf(Trim(.Record(mCol.����).Value) <> "", Trim(.Record(mCol.����).Value), "��"), _
        "��ʶ�� = " & IIf(Trim(.Record(mCol.��ʶ��).Value) <> "", Trim(.Record(mCol.��ʶ��).Value), "��"), _
        "���ڿ��� = " & IIf(Trim(.Record(mCol.���˿���).Value) <> "", Trim(.Record(mCol.���˿���).Value), "��"), _
        "�ɼ���ʽ = " & IIf(str�ɼ���ʽ <> "", str�ɼ���ʽ, "��"), _
        "�걾 = " & IIf(Trim(.Record(mCol.����걾).Value) <> "", Trim(.Record(mCol.����걾).Value), "��"), _
        "ִ�п��� = " & IIf(strִ�п��� <> "", strִ�п���, "��"), _
        "����ҽ�� = " & IIf(Trim(.Record(mCol.������).Value) <> "", Trim(.Record(mCol.������).Value), "��"), _
        "����ʱ�� = " & IIf(str����ʱ�� <> "", str����ʱ��, "��"), _
        "������ = " & IIf(Trim(.Record(mCol.������).Value) <> "", Trim(.Record(mCol.������).Value), "��"), _
        "����ʱ�� = " & IIf(Trim(.Record(mCol.����ʱ��).Value) <> "", Format(Trim(.Record(mCol.����ʱ��).Value), "yyyy-MM-dd HH:mm:ss"), "��"), _
        "���� = " & IIf(str���� <> "", str����, "��"), _
        "��Ѫ�� = " & IIf(str��Ѫ�� <> "", str��Ѫ��, "��"), _
        "�Թ����� = " & IIf(str������ <> "", str������, "��"), _
        "���� = " & IIf(Trim(.Record(mCol.����).Value) <> "", Trim(.Record(mCol.����).Value), "��"), _
        "������Դ = " & IIf(lng������Դ <> 0, lng������Դ, "��"), _
        "����ͼ��1=" & App.path & "\BarCode.Bmp", 2)
        'ɾ������ͼ��
        Kill App.path & "\BarCode.Bmp"
    End With
End Sub

''''''''''''''''''''
''' ʵ�ֲ����HOST����
''''''''''''''''''''
Private Property Get clsLisQueryHost_OwnerFormHandle() As Long
    clsLisQueryHost_OwnerFormHandle = Me.hWnd
End Property

Private Function clsLisQueryHost_GetRecordSet(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    'ִ�в�ѯ
    On Error GoTo errH
    Dim lngCount As Long
    Dim var(30) As Variant

    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        MsgBox "��֧�ֳ���30��������SQL��", vbInformation, Me.Caption
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    Set clsLisQueryHost_GetRecordSet = zlDatabase.OpenSQLRecord(strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub clsLisQueryHost_RaiseFinished(objQuery As zl9LisQuery_Def.clsLisQuery)
    'ִ�����
    On Error GoTo errH
    If objQuery.Result <> "" Then
        'Ԥ��
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function clsLisQueryHost_ClientTrigger(ByVal Index As Long, ByVal strAction As String, strData As String) As String
    '�ͻ��˴������¼�
    On Error GoTo errH
    If Not mobjPlugin(Index) Is Nothing Then
        Select Case strAction
        Case "Cmd_Start"
        Case "Cmd_End"
        Case "Cmd_OK"
        Case "Cmd_Cancle"
        End Select
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowHideListHead(Cols As ReportColumns, strFiled As String)
    '��ʾ������ͷ
    Dim intLoop As Integer
    
    For intLoop = 0 To Cols.Count - 1
        Cols(intLoop).Visible = (InStr(strFiled & ";", ";" & Cols(intLoop).Caption & ";") > 0)
    Next
End Sub

Private Sub ShowLJAverage()
    frmQCLJAverage.ShowMe Me, mstrPrivs, mlngDeptID, mlngMachineID
End Sub

