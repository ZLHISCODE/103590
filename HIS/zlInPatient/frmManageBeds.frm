VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBeds 
   Caption         =   "������λ����"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManageBeds.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   2760
      ScaleHeight     =   3135
      ScaleWidth      =   2505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   2505
      Begin XtremeReportControl.ReportControl rptUnit 
         Height          =   2130
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   3757
         _StockProps     =   0
         MultipleSelection=   0   'False
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   2505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2505
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2130
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   3757
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
         MultipleSelection=   0   'False
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7860
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15266
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
            AutoSize        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1260
      Left            =   6720
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   2222
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
      BackColorFixed  =   15790320
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
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin MSComctlLib.ImageList imgColor 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":0CCA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":0EE4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":10FE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1318
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1532
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":174C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1966
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1B80
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1D9A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1FB4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":21CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   1440
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":2AA8
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":2DC2
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":30DC
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":33F6
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":3710
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":3A2A
            Key             =   "MASK_�Ӵ�"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":3D44
            Key             =   "MASK_�Ǳ�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":405E
            Key             =   "MASK_����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4378
            Key             =   "MASK_����_�Ӵ�"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4692
            Key             =   "MASK_����_�Ǳ�"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2025
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":49AC
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4CC6
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4FE0
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":52FA
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5614
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":592E
            Key             =   "MASK_�Ӵ�"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5A88
            Key             =   "MASK_�Ǳ�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5BE2
            Key             =   "MASK_����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5D3C
            Key             =   "MASK_����_�Ӵ�"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5E96
            Key             =   "MASK_����_�Ǳ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5FF0
            Key             =   "�Ӵ�_Empty"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":6CCA
            Key             =   "�Ǳ�_Empty"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":79A4
            Key             =   "����_Empty"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":867E
            Key             =   "����_�Ӵ�_Empty"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":9358
            Key             =   "����_�Ǳ�_Empty"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":A032
            Key             =   "�Ӵ�_M_Empty"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":AD0C
            Key             =   "�Ǳ�_M_Empty"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":B9E6
            Key             =   "����_M_Empty"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":C6C0
            Key             =   "����_�Ӵ�_M_Empty"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":D39A
            Key             =   "����_�Ǳ�_M_Empty"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":E074
            Key             =   "�Ӵ�_F_Empty"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":ED4E
            Key             =   "�Ǳ�_F_Empty"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":FA28
            Key             =   "����_F_Empty"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":10702
            Key             =   "����_�Ӵ�_F_Empty"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":113DC
            Key             =   "����_�Ǳ�_F_Empty"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":120B6
            Key             =   "�Ӵ�_Holding"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":12D90
            Key             =   "�Ǳ�_Holding"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":13A6A
            Key             =   "����_Holding"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":14744
            Key             =   "����_�Ӵ�_Holding"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1541E
            Key             =   "����_�Ǳ�_Holding"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":160F8
            Key             =   "�Ӵ�_Remedy"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":169D2
            Key             =   "�Ǳ�_Remedy"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":176AC
            Key             =   "����_Remedy"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":18386
            Key             =   "����_�Ӵ�_Remedy"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":19060
            Key             =   "����_�Ǳ�_Remedy"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":19D3A
            Key             =   ""
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
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmManageBeds.frx":1A614
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmManageBeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnUnload As Boolean
Private mintEmpty As Integer, intHolding, intRemedy As Integer
Private Const STR_HEAD = "����,600,0,1;����,1200,0,2;�����,800,0,2;״̬,600,0,2;�Ա����,1000,0,2;�ȼ�,1000,0,2;��λ����,1000,0,2;����,1000,0,0;�Ա�,600,0,0;����,600,0,0"
Private mstrPrivs As String

Const conPane_Type = 201
Const conPane_List = 202
Const conPane_Edit = 203

Private mintEditState As Integer    '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�༭״̬
Private mfrmEditBed As frmBedEdit

Private mlngUnit As Long
Private mLngEditWidth As Long       '�༭������
Private mstrGroupBy As String       '��¼�������
'����,����,�����,״̬,�Ա����,�ȼ�,��λ����,����,�Ա�,����,����ID,�ȼ�id,����ID,����,����ID
Public Enum mCol
    ͼ�� = 0: ����: ����: �����: ˳���: ״̬: �Ա����: �ȼ�: ��λ����: ����: ����: �Ա�: ����: סԺ״̬: ����ID: ����ID: ����ID: �ȼ�ID: ����
End Enum

Private Enum mIcon
    iEmpty = 0: iM_Empty: iF_Empty: i_To_Empty: i_To_Repair
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Dim objControl As CommandBarControl
    'Dim objCombo As CommandBarComboBox
    Dim objRow As ReportRow, i As Long
    
    Dim strBedNO As String
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    '------------------------------------
    
    'Set objCombo = cbsMain(cbsMain.Count).FindControl(, conMenu_Edit_SelUnit, True)
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_Save:                                                                 '����
        strBedNO = mfrmEditBed.zlEditSave
        If strBedNO <> "" Then
            Call zlRefList(strBedNO)
            If mfrmEditBed.chkContAdd.Value Then
                If mfrmEditBed.zlEditStart(True, mlngUnit) = False Then
                    MsgBox "���ݳ�ʼ������", vbExclamation, gstrSysName
                    Exit Sub
                    
                End If
            Else
                ShowEdit False
                mintEditState = 0: Me.picUnit.Enabled = True: Me.picList.Enabled = True: Me.rptList.SetFocus
            End If
        End If

    Case conMenu_Edit_Untread:                                                              'ȡ��
        Call ShowEdit(False)
        Call mfrmEditBed.zlEditCancel
        mintEditState = 0: Me.picUnit.Enabled = True: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_NewItem                                                               '����

        If mlngUnit <= 0 Then
            MsgBox "��ѡ������", vbExclamation, gstrSysName
            rptUnit.SetFocus
            Exit Sub
        End If
        If mfrmEditBed Is Nothing Then Set mfrmEditBed = New frmBedEdit
        
        Call ShowEdit(True)
        If mfrmEditBed.zlEditStart(True, mlngUnit) = False Then Call ShowEdit(False): Exit Sub
        mintEditState = 1: Me.picUnit.Enabled = False: Me.picList.Enabled = False
    Case conMenu_Edit_Modify                                                            '����
        If mlngUnit <= 0 Then
            MsgBox "��ѡ������", vbExclamation, gstrSysName: Exit Sub
        End If
        
        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "��ѡ��Ҫ�����Ĳ�����", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.״̬).Value = "ռ��" Then
                MsgBox "�ò����ѱ�����ռ��,���ڲ��ܽ��е�����", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.״̬).Value = "ռ��" = "����" Then
                MsgBox "�ò�����������,���ڲ��ܽ��е�����", vbExclamation, gstrSysName: Exit Sub
            End If
        End With
        
'        On Error Resume Next
'        Err.Clear
        Call ShowEdit(True)
        If mfrmEditBed.zlEditStart(False, mlngUnit, rptList.FocusedRow.Record) = False Then Call ShowEdit(False):  Exit Sub
        
        mintEditState = 1: Me.picUnit.Enabled = False: Me.picList.Enabled = False
    Case conMenu_Edit_Delete                                                            'ɾ��

        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "��ѡ��Ҫ�����Ĳ�����", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.״̬).Value = "ռ��" Then
                MsgBox "�ò����ѱ�����ռ��,���ڲ��ܳ�����", vbExclamation, gstrSysName: Exit Sub
            End If
            If MsgBox("ȷʵҪ��������" & .FocusedRow.Record(mCol.����).Value & " ��", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            intIndex = .FocusedRow.Index

            gstrSQL = "zl_��λ״����¼_Delete('" & Trim(.FocusedRow.Record(mCol.����).Value) & "'," & mlngUnit & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            If .Rows.Count > intIndex + 1 Then
                If .Rows(intIndex + 1).GroupRow = False Then strBedNO = .Rows(intIndex + 1).Record(mCol.����).Value
            ElseIf intIndex > 0 Then
                If .Rows(intIndex - 1).GroupRow = False Then strBedNO = .Rows(intIndex - 1).Record(mCol.����).Value
            End If
            Call Me.zlRefList(strBedNO)
        End With
    Case conMenu_Edit_Bed_ToRepair                                                          'ת����
        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "��ѡ��Ҫ���ɵĲ�����", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.״̬).Value <> "�մ�" Then
                MsgBox "�ò������ǿմ�,����ִ�иò�����", vbExclamation, gstrSysName: Exit Sub
            End If
            gstrSQL = "zl_��λ״����¼_STOP('" & Trim(.FocusedRow.Record(mCol.����).Value) & "'," & .FocusedRow.Record(mCol.����ID).Value & ")"
            
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            strBedNO = Trim(.FocusedRow.Record(mCol.����).Value)
            Call zlRefList(strBedNO)

        End With
    Case conMenu_Edit_Bed_ToEmpty                                                           'ת�մ�
        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "��ѡ���Ѿ����ɺõĲ�����", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.״̬).Value <> "����" Then
                MsgBox "�ò���û�н�������,����ִ�иò�����", vbExclamation, gstrSysName: Exit Sub
            End If
            
            gstrSQL = "zl_��λ״����¼_REUSE('" & Trim(.FocusedRow.Record(mCol.����).Value) & "'," & .FocusedRow.Record(mCol.����ID).Value & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            strBedNO = Trim(.FocusedRow.Record(mCol.����).Value)
            Call zlRefList(strBedNO)
        End With
        
    Case conMenu_View_ToolBar_Button                                                        '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text                                                          '��ʾ�ı�
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                If Not (objControl.Type = xtpControlLabel Or objControl.Type = xtpControlComboBox) Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size                                                          '��ͼ��/Сͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar                                                             '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Column                                                                'ѡ����
        
    Case conMenu_View_Refresh                                                               'ˢ��
        '59753:������,2013-4-18
        If Me.rptList.FocusedRow Is Nothing Then
            zlRefList
        Else
            If Me.rptList.FocusedRow.GroupRow Then
                zlRefList
            Else
                zlRefList Trim(Me.rptList.FocusedRow.Record(mCol.����).Value)
            End If
        End If
    Case conMenu_Help_Help:     Call ShowHelp("", Me.hWnd, Me.Name, Int((glngSys) / 100))   '����
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        '--ִ���Զ��屨��
        If Control.ID > 401 And Control.ID < 499 Then
            If Me.rptList.FocusedRow Is Nothing Then
                strBedNO = ""
            Else
                If Me.rptList.FocusedRow.GroupRow Then
                    strBedNO = ""
                Else
                    strBedNO = Trim(Me.rptList.FocusedRow.Record(mCol.����).Value)
                End If
            End If
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                "����=" & mlngUnit, "����=" & strBedNO)
        End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If

    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Visible = InStr(1, mstrPrivs, "��λ�༭") > 0
        Control.Enabled = (mintEditState <> 0)

    Case conMenu_Edit_NewItem
        Control.Visible = InStr(1, mstrPrivs, "��λ�༭") > 0
        Control.Enabled = (InStr(1, mstrPrivs, "��λ�༭") > 0 And mintEditState = 0)
        
    Case conMenu_Edit_Modify
        Control.Visible = InStr(1, mstrPrivs, "��λ�༭") > 0
        Control.Enabled = (InStr(1, mstrPrivs, "��λ�༭") > 0 And mintEditState = 0 And Me.rptList.Rows.Count)
        'If Control.Enabled Then Control.Enabled = mstr���� <> ""
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_Edit_Delete
        Control.Visible = InStr(1, mstrPrivs, "��λ�༭") > 0
        Control.Enabled = (InStr(1, mstrPrivs, "��λ�༭") > 0 And mintEditState = 0 And Me.rptList.Rows.Count And rptList.FocusedRow.Record(mCol.״̬).Value <> "����")
        'If Control.Enabled Then Control.Enabled = mstr���� <> ""
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_Edit_Bed_ToRepair
        Control.Visible = InStr(1, mstrPrivs, "��λ�༭") > 0
        If Me.rptList.Records.Count > 0 Then
            Control.Enabled = (mintEditState = 0 And rptList.FocusedRow.Record(mCol.״̬).Value = "�մ�")
        Else
            Control.Enabled = False
        End If
        
    Case conMenu_Edit_Bed_ToEmpty
        Control.Visible = InStr(1, mstrPrivs, "��λ�༭") > 0
        If Me.rptList.Records.Count > 0 Then
            Control.Enabled = (mintEditState = 0 And rptList.FocusedRow.Record(mCol.״̬).Value = "����")
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsMain(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = 1
    End Select
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If mfrmEditBed Is Nothing Then Set mfrmEditBed = New frmBedEdit

    Select Case Item.ID
    Case conPane_Type
        Item.Handle = Me.picUnit.hWnd
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        Item.Handle = mfrmEditBed.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me
End Sub

Private Sub Form_Load()
    
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    
    mLngEditWidth = frmBedEdit.ScaleWidth
    
    
    Call InitCommandBar
    Call InitDockPannel
    Call InitReportColumn
    Call RestoreWinState(Me, App.ProductName)
    Call ZLCommFun.SetWindowsInTaskBar(Me.hWnd, False)

    Call MakeBedIcon

    '��ȡ����
    If Not InitUnits Then mblnUnload = True: Exit Sub
    If rptUnit.Records.Count = 0 Then
        MsgBox "�㲻�������в�����Ȩ��,���Ҳ���ȷ������������,����ʹ�ô�λ����", vbExclamation, gstrSysName
        mblnUnload = True: Exit Sub
    End If

'    If Not ReadBeds(mlngUnit) Then
'        mblnUnload = True: Exit Sub
'    End If
    
    mstrPrivs = gstrPrivs
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCombo As CommandBarComboBox

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
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        'Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "����(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToRepair, "ת����(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToEmpty, "ת�մ�(&T)")

    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            '.Add xtpControlButton, conMenu_View_Append, "����ѡ��(&U)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        'Set objControl = .Add(xtpControlButton, conMenu_View_Column, "ѡ����(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToRepair, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToEmpty, "�մ�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, ""): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
'    With objBar.Controls
'        Set objControl = .Add(xtpControlLabel, 0, "���� "): objControl.BeginGroup = True
'        Set objCombo = .Add(xtpControlComboBox, conMenu_Edit_SelUnit, "")  '�޷���ʾͼ��
'            objCombo.DropDownListStyle = True
'            objCombo.Width = 150
'            objCombo.DefaultItem = True
'            objCombo.flags = xtpFlagControlStretched
'    End With
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save '����
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread '����
        .Add FCONTROL, vbKeyR, conMenu_Edit_Bed_ToRepair '����
        .Add FCONTROL, vbKeyE, conMenu_Edit_Bed_ToEmpty '�մ�
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����һЩ�����Ĳ���������
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
'        .AddHiddenCommand conMenu_File_Excel '�����Excel
'    End With
    '����Զ��屨��
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Sub

Private Sub InitDockPannel()
    Dim objPaneType As Pane, objPaneList As Pane, objPaneEdit As Pane
    
    If mfrmEditBed Is Nothing Then Set mfrmEditBed = New frmBedEdit
    
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    'Me.dkpMain.Options.AlphaDockingContext = True
    Me.dkpMain.Options.HideClient = True
    
    
    Set objPaneType = Me.dkpMain.CreatePane(conPane_Type, 200, 600, DockLeftOf)
    objPaneType.Title = "�����б�"
    objPaneType.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
    
    Set objPaneList = Me.dkpMain.CreatePane(conPane_List, 750, 600, DockRightOf, objPaneType)
    objPaneList.Title = "������λ�б�"
    objPaneList.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
    
    Set objPaneEdit = Me.dkpMain.CreatePane(conPane_Edit, 250, 600, DockRightOf)
    objPaneEdit.Title = "������λ�༭"
    objPaneEdit.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPaneEdit.MaxTrackSize.SetSize 380, 600
    
    objPaneEdit.Close

End Sub

Private Sub InitReportColumn()
'���ܣ���ʼ�������б���
    Dim objCol As ReportColumn
        
    With rptUnit
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        .AutoColumnSizing = False  '������������֮ǰ���ã�������Ч
        Set objCol = .Columns.Add(mCol.ͼ��, "", 18, False): objCol.Editable = False: objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(1, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(2, "����", 1200, True): objCol.Editable = False
        .ShowHeader = True
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "û�п���ʾ�Ĳ�����Ϣ..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With

    With rptList
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        .AutoColumnSizing = False   '������������֮ǰ���ã�������Ч
        Set objCol = .Columns.Add(mCol.ͼ��, "", 18, False):  objCol.Groupable = False: objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(mCol.����, "����", 50, True):  objCol.Groupable = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����, "����", 100, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.�����, "�����", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.˳���, "˳���", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.״̬, "״̬", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.�Ա����, "�Ա����", 60, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.�ȼ�, "�ȼ�", 140, True):  objCol.Groupable = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.��λ����, "��λ����", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����, "����", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����, "����", 50, True):  objCol.Groupable = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.�Ա�, "�Ա�", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����, "����", 30, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.סԺ״̬, "סԺ״̬", 100, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.�ȼ�ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.����, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
    
        .ShowHeader = True
        .ShowGroupBox = True
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
    
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĵ�λ��Ϣ..."
        End With
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    Dim objItem As TaskPanelGroupItem
    
    If Me.rptList.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    '��ͷ
    objPrint.Title.Text = "������λ�嵥"
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ�ˣ�" & UserInfo.����)
    Call objAppRow.Add("��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    Set objAppRow = New zlTabAppRow

            Call objAppRow.Add("����:" & ZLCommFun.GetNeedName(rptUnit.FocusedRow.Record(2).Value))  'cboUnit.Text

    Call objPrint.UnderAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
   
End Sub

'################################################################################################################
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand
    For Each rptRow In rptList.Rows
        If rptRow.Childs.Count > 0 Then rptRow.Expanded = True
    Next
    If rptList.Rows.Count < 1 Then zlReportToVSFlexGrid = False: Exit Function
        
    With vfgList
        .Clear
        .Rows = 1: .FixedRows = 1: .RowHeight(.Rows - 1) = 280
        .Cols = 0
        .MergeCells = flexMergeFree
        
        '�����и���
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = rptCol.Caption
                .ColData(.Cols - 1) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .colAlignment(.Cols - 1) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .colAlignment(.Cols - 1) = flexAlignCenterCenter
                Case xtpAlignmentRight: .colAlignment(.Cols - 1) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, .Cols - 1, .FixedRows - 1) = flexAlignCenterCenter
                If rptCol.width < 20 * IIf(rptList.GroupsOrder.Count = 0, 1, rptList.GroupsOrder.Count) Then
                    .ColWidth(.Cols - 1) = 0
                Else
                    .ColWidth(.Cols - 1) = rptCol.width * Screen.TwipsPerPixelX
                End If
            End If
        Next
        
        '�����и���
        Dim intTiers As Integer, rptParent As ReportRow, rptChild As ReportRow
        For Each rptRow In rptList.Rows
            .Rows = .Rows + 1: .RowHeight(.Rows - 1) = 280
            If rptRow.GroupRow Then
                intTiers = 0
                Set rptParent = rptRow
                Do While Not (rptParent.ParentRow Is Nothing)
                    intTiers = intTiers + 1
                    Set rptParent = rptParent.ParentRow
                Loop
                Set rptChild = rptRow.Childs(0)
                Do While rptChild.GroupRow
                    Set rptChild = rptChild.Childs(0)
                Loop
                .MergeRow(.Rows - 1) = True
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "��") & rptList.GroupsOrder(intTiers).Caption & ": "
                    .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & rptChild.Record(rptList.GroupsOrder(intTiers).ItemIndex).Value
                Next
            Else
                For lngCol = 0 To .Cols - 1
                    If rptList.Columns(.ColData(lngCol)).TreeColumn Then
                        intTiers = 0
                        Set rptParent = rptRow
                        Do While Not (rptParent.ParentRow Is Nothing)
                            intTiers = intTiers + 1
                            Set rptParent = rptParent.ParentRow
                        Loop
                        .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "��") & rptRow.Record(.ColData(lngCol)).Value
                    Else
                        .TextMatrix(.Rows - 1, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                    End If
                    .Cell(flexcpAlignment, .Rows - 1, lngCol, .Rows - 1) = .colAlignment(lngCol)
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Private Sub Form_Resize()
    Dim panType As Pane
    Dim panEdit As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panType = Me.dkpMain.FindPane(conPane_Type)
    Set panEdit = Me.dkpMain.FindPane(conPane_Edit)
    panType.MinTrackSize.SetSize 150, 600
    panType.MaxTrackSize.SetSize 200, 375
    panEdit.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265
    panEdit.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 375
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
'
'    panEdit.MinTrackSize.SetSize 0, 0
'    panEdit.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 375
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmEditBed Is Nothing Then
        Unload mfrmEditBed
        If mfrmEditBed.mintCancle = 1 Then Cancel = 1: Exit Sub
        Set mfrmEditBed = Nothing
    End If
    mblnUnload = False
    mintEditState = 0
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
'    mfrmEditBed.fraEdit.Width = mfrmEditBed.Width
'    mfrmEditBed.fraEdit.Height = Me.ScaleHeight - Me.picList.ScaleHeight
End Sub

Private Sub MakeBedIcon()
    Dim i As Integer, k As Integer
    
    k = img32.ListImages.Count
    For i = 1 To img32.ListImages.Count
        If Not img32.ListImages(i).Key Like "MASK_*" Then
            img32.ListImages.Add , "�Ӵ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_�Ӵ�", i)
            img32.ListImages.Add , "�Ǳ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_�Ǳ�", i)
            img32.ListImages.Add , "����_" & img32.ListImages(i).Key, img32.Overlay("MASK_����", i)
            img32.ListImages.Add , "����_�Ӵ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_����_�Ӵ�", i)
            img32.ListImages.Add , "����_�Ǳ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_����_�Ǳ�", i)
        End If
    Next
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ����
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim i As Integer, lngUnitID As Long, blnLimitUnit As Boolean
    Dim strUnitIDs As String
    Dim intCurrIndes As Integer
    
    On Error GoTo errH
    
    'Set objCombo = cbsMain(cbsMain.Count).FindControl(, conMenu_Edit_SelUnit, True)
    '��������۲���
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    '����30922 by lesfeng 2010-06-18 b
    If blnLimitUnit Then strUnitIDs = UserInfo.ID
    'by lesfeng 2010-1-8 �����Ż�
    gstrSQL = _
        " Select A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & IIf(blnLimitUnit, ",������Ա C ", "") & _
        " Where B.����ID = A.ID" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.������� IN(1,2,3) And B.��������='����'" & _
        IIf(blnLimitUnit, " And A.ID = C.����ID And C.��ԱID In ([1])", "") & _
        " And (A.վ��=[2] Or A.վ�� is Null)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strUnitIDs), gstrNodeNo)
    
    '����30922 by lesfeng 2010-06-18 e
    If Not rsTmp.EOF Then
        Me.rptUnit.Records.DeleteAll

        Do While Not rsTmp.EOF
            '��������
            Set objRecord = rptUnit.Records.Add()
            Set objItem = objRecord.AddItem(""): objItem.Icon = 35
            objRecord.AddItem Val("" & rsTmp!ID)
            objRecord.AddItem Nvl(rsTmp!����)    'Nvl(rsTmp!����) & "-" &
            rsTmp.MoveNext
        Loop
        With Me.rptUnit
            .Populate
        End With
        If Me.rptUnit.Records.Count - 1 > 0 Then Me.rptUnit.FocusedRow = rptUnit.Rows(IIf(rptUnit.Rows(1).GroupRow, 1, 0))
    ElseIf InStr(";" & mstrPrivs, "���в���") > 0 Then
        MsgBox "û�����ò���,�����ȵ����Ź��������ù�������Ϊ����Ĳ��ţ�", vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBox "��û�� [���в���] ��Ȩ��,���������ڲ��Ų��ǲ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBeds(lngUnitID As Long) As Boolean
    '���ܣ���ȡָ�������Ĵ�λ�б�
    Dim i As Integer, j As Integer
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim intBedLen As Integer
    Dim mrsBeds As ADODB.Recordset
    Dim str�۸�ȼ� As String
    
    intHolding = 0: intRemedy = 0: mintEmpty = 0
    
    '���ͳ������
    On Error GoTo errH
    intBedLen = GetMaxBedLen(lngUnitID)
    
    gstrSQL = "Select a.�۸�ȼ�" & vbNewLine & _
            "  From �շѼ۸�ȼ�Ӧ�� A,�շѼ۸�ȼ� B" & vbNewLine & _
            "  where A.�۸�ȼ�=b.����  And a.����=0 And b.�Ƿ�������ͨ��Ŀ=1 and a.վ��=[1]" & vbNewLine & _
            "        and nvl(b.����ʱ��,sysdate+1)>sysdate"
    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrNodeNo)
    If mrsBeds.RecordCount > 0 Then
        str�۸�ȼ� = mrsBeds!�۸�ȼ� & ""
    End If
    gstrSQL = "Select a.*" & vbNewLine & _
            "From (Select LPad(a.����, [1], ' ') ����, a.����id, a.�����,a.˳���,a.�Ա����, a.��λ����, a.�ȼ�id, a.״̬, a.����id, a.����, e.�ּ� As �۸�," & vbNewLine & _
            "              Nvl(b.����, Decode(a.����, 1, '<���ò���>', Null)) As ����, a.����id, c.���� As �ȼ�, d.����, d.�Ա�, d.����," & vbNewLine & _
            "              Decode(d.״̬, 0, '����סԺ', 2, '׼��ת��', 3, '׼����Ժ' || '(' || To_Char(f.��ʼʱ��, 'YYYY-MM-DD HH24:MI:SS') || ')') As סԺ״̬," & vbNewLine & _
            "              Row_Number() Over(Partition By a.����, e.�շ�ϸĿid Order By Decode(e.�۸�ȼ�, [3], 1, Null, 2, 3)) As Top" & vbNewLine & _
            "       From ��λ״����¼ A, ���ű� B, �շ���ĿĿ¼ C," & vbNewLine & _
            "            (Select m.����id, Nvl(n.����, m.����) ����, Nvl(n.�Ա�, m.�Ա�) �Ա�, Nvl(n.����, m.����) ����, n.״̬" & vbNewLine & _
            "              From ������Ϣ M, ������ҳ N" & vbNewLine & _
            "              Where m.����id = n.����id And m.��ҳid = n.��ҳid And n.��ǰ����id = [2]) D, �շѼ�Ŀ E," & vbNewLine & _
            "            (Select q.����id, ��ʼʱ��" & vbNewLine & _
            "              From ������ҳ P, ���˱䶯��¼ Q" & vbNewLine & _
            "              Where p.����id = q.����id And p.��ҳid = q.��ҳid And ��ʼԭ�� = 10 And p.��ǰ����id = [2] And p.״̬ = 3) F" & vbNewLine & _
            "       Where a.����id = b.Id(+) And a.�ȼ�id = c.Id(+) And a.����id = d.����id(+) And e.�շ�ϸĿid = c.Id And a.����id = f.����id(+) And" & vbNewLine & _
            "             a.����id = [2] And Sysdate Between e.ִ������ And Nvl(e.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             Nvl(e.�۸�ȼ�, [3]) = [3]" & vbNewLine & _
            "       Order By a.˳���,LPad(a.����, [1], ' ')) A" & vbNewLine & _
            "Where Top = 1"

    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intBedLen, lngUnitID, IIf(str�۸�ȼ� = "", "��", str�۸�ȼ�))
    
    Me.rptList.Records.DeleteAll
    Me.rptList.SortOrder.DeleteAll  'ÿ���������Ĭ�ϰ�������������
    Do While Not mrsBeds.EOF
        '��������
        Set objRecord = rptList.Records.Add()
        Select Case mrsBeds!״̬
            Case "�մ�"
                If mrsBeds!�Ա���� = "�д�" Then
                    Set objItem = objRecord.AddItem(""): objItem.Icon = 1
                ElseIf mrsBeds!�Ա���� = "Ů��" Then
                    Set objItem = objRecord.AddItem(""): objItem.Icon = 2
                Else
                    Set objItem = objRecord.AddItem(""): objItem.Icon = 0
                End If
                mintEmpty = mintEmpty + 1
            Case "ռ��"
                Set objItem = objRecord.AddItem(""): objItem.Icon = 3
                intHolding = intHolding + 1
            Case "����"
                Set objItem = objRecord.AddItem(""): objItem.Icon = 4
                intRemedy = intRemedy + 1
            Case Else   '��������
                Set objItem = objRecord.AddItem(""): objItem.Icon = 4
                intRemedy = intRemedy + 1
        End Select
        
        objRecord.AddItem (Trim(mrsBeds!����))
        objRecord.AddItem (Nvl(mrsBeds!����))
        objRecord.AddItem (Nvl(mrsBeds!�����))
        objRecord.AddItem (Nvl(mrsBeds!˳���))
        objRecord.AddItem (Nvl(mrsBeds!״̬))
        objRecord.AddItem (Nvl(mrsBeds!�Ա����))
        objRecord.AddItem (Nvl(mrsBeds!�ȼ�))
        objRecord.AddItem (Nvl(mrsBeds!��λ����))
        objRecord.AddItem Format((Nvl(mrsBeds!�۸�)), "0.00")
        objRecord.AddItem (Nvl(mrsBeds!����))
        objRecord.AddItem (Nvl(mrsBeds!�Ա�))
        objRecord.AddItem (Nvl(mrsBeds!����))
        objRecord.AddItem (Nvl(mrsBeds!סԺ״̬))
        objRecord.AddItem (Nvl(mrsBeds!����ID))
        objRecord.AddItem (Nvl(mrsBeds!����ID))
        objRecord.AddItem (Nvl(mrsBeds!����ID))
        objRecord.AddItem (Nvl(mrsBeds!�ȼ�ID))
        objRecord.AddItem (Nvl(mrsBeds!����))

        If objRecord.Item(mCol.��λ����).Value = "�Ӵ�" Then
            For i = 1 To img16.ListImages.Count
                If img16.ListImages(i).Key = "�Ӵ�_" & img16.ListImages(objRecord.Item(mCol.ͼ��).Icon + 1).Key Then
                    Exit For
                End If
            Next
            objRecord.Item(mCol.ͼ��).Icon = i - 1
        ElseIf objRecord.Item(mCol.��λ����).Value = "�Ǳ�" Then
            For i = 1 To img16.ListImages.Count
                If img16.ListImages(i).Key = "�Ǳ�_" & img16.ListImages(objRecord.Item(mCol.ͼ��).Icon + 1).Key Then
                    Exit For
                End If
            Next
            objRecord.Item(mCol.ͼ��).Icon = i - 1
        End If
        If Val(objRecord(mCol.����).Value) <> 0 Then
            For i = 1 To img16.ListImages.Count
                If img16.ListImages(i).Key = "����_" & img16.ListImages(objRecord.Item(mCol.ͼ��).Icon + 1).Key Then
                    Exit For
                End If
            Next
            objRecord.Item(mCol.ͼ��).Icon = i - 1
        End If
        mrsBeds.MoveNext
    Loop
    With Me.rptList
        .Populate
    End With
    If Me.rptList.Records.Count - 1 > 0 Then Me.rptList.FocusedRow = rptList.Rows(IIf(rptList.Rows(1).GroupRow, 1, 0))
    Call SetBedNOLen(lngUnitID)
    ReadBeds = True
    stbThis.Panels(2) = "��ǰ������ " & rptList.Records.Count & " �Ų���,���в���ռ�� " & intHolding & " ��,�մ� " & mintEmpty & " ��,�������� " & intRemedy & " �ţ�"
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowEdit(blnShow As Boolean)
    '����       �Ƿ���ʾ�༭����
    Dim objPane As Pane
    Set objPane = dkpMain.FindPane(conPane_Edit)
    If blnShow = True Then
        objPane.Select
    Else
        objPane.Close
    End If
    dkpMain.RecalcLayout
End Sub

Public Function zlRefList(Optional strBedNO As String) As Long
    '���ܣ�ˢ��װ�봲λ�嵥������λ��ָ���Ĵ�λ��
    Dim objCombo As CommandBarComboBox
    Dim rptRow As ReportRow
    Dim rptParent As ReportRow
    

    If ReadBeds(mlngUnit) Then
        If strBedNO <> "" Then
            For Each rptRow In Me.rptList.Rows
                If rptRow.GroupRow = False Then
                    If Trim(rptRow.Record(mCol.����).Value) = strBedNO Then
                        Set rptParent = rptRow.ParentRow
                        Set Me.rptList.FocusedRow = rptRow
                        Exit For
                    End If
                End If
            Next
            For Each rptRow In Me.rptList.Rows
                If rptRow.GroupRow Then
                    If Not (rptRow Is rptParent) Then rptRow.Expanded = False
                End If
            Next
            Set Me.rptList.FocusedRow = Me.rptList.FocusedRow
        Else
            For Each rptRow In Me.rptList.Rows
                If rptRow.GroupRow Then rptRow.Expanded = True
            Next
        End If
        If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Public Sub SetBedNOLen(ByVal lngUnitID As Long)
    Dim bytLen As Byte, i As Integer
    If rptList.Records.Count = 0 Then Exit Sub
    
    bytLen = GetMaxBedLen(lngUnitID)
    For i = 0 To rptList.Records.Count - 1
        rptList.Records(i).Item(mCol.����).Value = Space(bytLen - Len(CStr(rptList.Records(i).Item(mCol.����).Value))) & Trim(rptList.Records(i).Item(mCol.����).Value)
    Next
End Sub

Private Sub picUnit_Resize()
    Err = 0: On Error Resume Next
    With Me.rptUnit
        .Left = Me.picUnit.ScaleLeft: .width = Me.picUnit.ScaleWidth - .Left
        .Top = Me.picUnit.ScaleTop: .Height = Me.picUnit.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    If Me.cbsMain.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls(2)
    Set objBar = Me.cbsMain.Add("�����˵�", xtpBarPopup)
    For Each objControl In objPopup.CommandBar.Controls
        Set objControl = objBar.Controls.Add(xtpControlButton, objControl.ID, objControl.Caption)
        objControl.BeginGroup = objControl.BeginGroup
    Next
    objBar.ShowPopup
End Sub

Private Sub rptUnit_SelectionChanged()
    If rptUnit.FocusedRow Is Nothing Then Exit Sub
    mlngUnit = rptUnit.FocusedRow.Record(1).Value
    
    If Not ReadBeds(mlngUnit) Then
        mblnUnload = True: Exit Sub
    End If
    Me.Refresh
End Sub

