VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm�������� 
   AutoRedraw      =   -1  'True
   Caption         =   "���Ӳ�������"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   Icon            =   "frm��������.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk��� 
      Caption         =   "δ����"
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   24
      Top             =   675
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chk��� 
      Caption         =   "δ���"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1215
      TabIndex        =   23
      Top             =   675
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chk��� 
      Caption         =   "�����"
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   2
      Left            =   2160
      TabIndex        =   22
      Top             =   675
      Value           =   1  'Checked
      Width           =   915
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   135
      Top             =   9240
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid fg�б� 
      Height          =   2865
      Left            =   345
      TabIndex        =   16
      Top             =   3630
      Visible         =   0   'False
      Width           =   1635
      _cx             =   2884
      _cy             =   5054
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������.frx":08CA
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.TreeView tvw���� 
      Height          =   1170
      Left            =   2145
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3585
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   2064
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImgСͼ��"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.PictureBox picRight 
      BackColor       =   &H00FAFAFA&
      ClipControls    =   0   'False
      Height          =   4665
      Left            =   7095
      Picture         =   "frm��������.frx":0917
      ScaleHeight     =   4605
      ScaleWidth      =   6015
      TabIndex        =   5
      Top             =   1530
      Width           =   6075
      Begin VSFlex8Ctl.VSFlexGrid fg���_S 
         Height          =   1425
         Left            =   210
         TabIndex        =   1
         Top             =   1920
         Width           =   4920
         _cx             =   8678
         _cy             =   2514
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm��������.frx":0E14
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
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   2550
         TabIndex        =   25
         Top             =   870
         Width           =   810
      End
      Begin VB.Label lbl��ע 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע:"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   1590
         Width           =   450
      End
      Begin VB.Label lbl�����޸� 
         BackStyle       =   0  'Transparent
         Caption         =   "�̷����޸�"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2565
         TabIndex        =   17
         Top             =   645
         Width           =   2580
      End
      Begin VB.Label lbl����ʱ�� 
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��:"
         Height          =   195
         Left            =   2565
         TabIndex        =   15
         Top             =   1113
         Width           =   2580
      End
      Begin VB.Label lbl���ʱ�� 
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��:"
         Height          =   195
         Left            =   2565
         TabIndex        =   14
         Top             =   1350
         Width           =   2580
      End
      Begin VB.Label lbl������Ϣ 
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   405
         Width           =   2580
      End
      Begin VB.Label lbl�ȼ� 
         BackStyle       =   0  'Transparent
         Caption         =   "�ȼ�:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   877
         Width           =   2580
      End
      Begin VB.Label lbl������ 
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1113
         Width           =   2580
      End
      Begin VB.Label lbl����� 
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1350
         Width           =   2580
      End
      Begin VB.Label lbl�ܷ� 
         BackStyle       =   0  'Transparent
         Caption         =   "�ܷ�:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   641
         Width           =   2580
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���ֽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   4950
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2115
      ScaleWidth      =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2460
      Width           =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8625
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm��������.frx":0F6D
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20770
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
   Begin MSComctlLib.ImageList ImgСͼ�� 
      Left            =   5145
      Top             =   6825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��������.frx":1801
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��������.frx":1978
            Key             =   "Dot"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft_S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   300
      ScaleHeight     =   1800
      ScaleWidth      =   4650
      TabIndex        =   12
      Top             =   1440
      Width           =   4650
      Begin VSFlex8Ctl.VSFlexGrid fg����_S 
         Height          =   1020
         Left            =   135
         TabIndex        =   0
         Top             =   345
         Width           =   4365
         _cx             =   7699
         _cy             =   1799
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   26
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm��������.frx":1A50
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
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Image imgSelCols_S 
         Height          =   195
         Left            =   4275
         MouseIcon       =   "frm��������.frx":1DC8
         MousePointer    =   99  'Custom
         Picture         =   "frm��������.frx":1F1A
         ToolTipText     =   "ѡ����Ҫ��ʾ����"
         Top             =   90
         Width           =   195
      End
      Begin VB.Image imgRefresh 
         Height          =   195
         Left            =   4005
         MouseIcon       =   "frm��������.frx":1F6D
         MousePointer    =   99  'Custom
         Picture         =   "frm��������.frx":20BF
         ToolTipText     =   "ˢ������"
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2970
      TabIndex        =   19
      Top             =   105
      Width           =   2760
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3390
      TabIndex        =   21
      Top             =   645
      Width           =   2250
   End
   Begin VB.Image imgLujinPic 
      Height          =   240
      Left            =   11070
      MouseIcon       =   "frm��������.frx":22FF
      MousePointer    =   99  'Custom
      Picture         =   "frm��������.frx":2451
      ToolTipText     =   "ˢ������"
      Top             =   8040
      Width           =   240
   End
   Begin VB.Image imgLujin 
      Height          =   240
      Left            =   11085
      MouseIcon       =   "frm��������.frx":8CA3
      MousePointer    =   99  'Custom
      Picture         =   "frm��������.frx":8DF5
      ToolTipText     =   "ˢ������"
      Top             =   8265
      Width           =   240
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
      Bindings        =   "frm��������.frx":F647
      Left            =   465
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
   Begin VB.Label Label4 
      Caption         =   "����(&T)"
      Height          =   195
      Left            =   2295
      TabIndex        =   20
      Top             =   165
      Width           =   870
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   165
      Picture         =   "frm��������.frx":F65B
      Top             =   9750
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   0
      Left            =   3090
      Picture         =   "frm��������.frx":F81B
      Top             =   9735
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   1
      Left            =   6345
      Picture         =   "frm��������.frx":1003F
      Top             =   9810
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private WithEvents mfrm���ֽ���༭ As frm���ֽ���༭
Attribute mfrm���ֽ���༭.VB_VarHelpID = -1
Private mstrPrivs               As String               'Ȩ�޴�
Private mobjFindKey             As CommandBarPopup      '��ѯ
Private mstrFindKey             As String               '��ѯ��
Private mlngModule              As Long                 'ģ���
Private m_lngOldRow             As Long
Private mcbrPopupBar            As CommandBar           '��������
Private mRecordRating           As Boolean              '���ַ�
Private mRecordAudit            As Boolean              '��˷�
Private mRecordMyAudit          As Boolean              '�Ƿ��˼�¼
Private mRecordReturn           As Boolean              '�������ַ�
Private mvarPara                As Variant              '��ѯ����
Private rsM                     As ADODB.Recordset      '�������ݼ�
Private mstrWhere               As String               '��ǰ�Ĳ�ѯ����
Private mblnSetDept             As Boolean              '�Ƿ��޶�����

Dim m_lng����ID                 As Long                 '��ǰ����ID
Dim m_lng��ҳID                 As Long                 '��ǰ��ҳID
Dim m_lng���ID                 As Long                 '��ǰ���ID
Dim m_lng����ID                 As Long                 '��ǰ���ַ���ID
Dim m_str�б���                 As String               '
Dim mfrm����                    As frm�������ֲ�ѯ      '���ڲ��Ҳ����Ĵ��壬��Ϊһ���ֲ�����ʹ�á�����Form_QueryUnload�йرգ�����ֻ������֮����
Dim cbrPopupItem                As CommandBarControl    '������
'��ѯ���ڱ���
Private mlngSickID              As Long             '����ID
Private mlngHospitalID          As Long             'סԺ��
Private mlngHospitalTimes       As Long             'סԺ����
Private mstrSickName            As String           '��������
Private mstrMainDoctor          As String           '����ҽʦ
Private mstrOutpatientDoctor    As String           '����ҽʦ
Private mstrNurses              As String           '���λ�ʿ
Private mstrRatingMan           As String           '������
Private mstrAuditMan            As String           '�����
Private mstrOutDept             As String           '��Ժ����
Private mstrInDept              As String           '��Ժ����
Private mdatStarOutDate         As Date             '��Ժ��ʼ����
Private mdatEndOutDate          As Date             '��Ժ��ʼ����
Private mdatStarInDate          As Date             '��Ժ��ʼ����
Private mdatEndInDate           As Date             '��Ժ��ʼ����
Private mstrSickType            As String           '��������
Private mfrmArchiveView         As frmArchiveView   '��������

'==============================================================================
'=���ܣ� �ؼ���ʼ��
'==============================================================================
Private Sub InitControl()
    On Error GoTo ErrH
    
    '�˵�����
    Call InitCommandBar
    '��������
    Call InitDockPannel
    '��ʼ������
    Call InitVsf
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ���򻮷�
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane             As Pane

    On Error GoTo ErrH
    
    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "������Ϣ"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 200, 100, DockRightOf, Nothing)
    objPane.Title = "���ֽ��"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ�˵�������
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo ErrH
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint, "ȫ����ӡ(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "��������(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "�޸Ľ��(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Insert, "��������(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "ɾ�����(&D)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_ReportView, "���Ĳ���(&V)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "ͨ�����(&P)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Leave_UndoPost, "ȡ�����(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "ȫ��ѡ��(&L)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeSelect, "ȡ��ѡ��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnAudit, "����ѡ��(&B)")
    
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Find, "����(&F)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & ParamInfo.��Ʒ����)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.��Ʒ���� & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.��Ʒ���� & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)
    
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    If Len(mstrFindKey) <= 2 Then mstrFindKey = "סԺ��"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.סԺ��", , , "סԺ��")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.��������", , , "��������")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.����ID", , , "����ID")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.���￨��", , , "���￨��")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&5.��λ�������", , , "��λ�������")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txt����.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_ReportView, "���Ĳ���(&V)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "ͨ�����(&P)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Leave_UndoPost, "ȡ�����(&C)")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewParent, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_ModifyParent, "�޸�")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Insert, "����")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_DeleteParent, "ɾ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_ReportView, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Audit, "���", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Leave_UndoPost, "����")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Find, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    If ScaleWidth - 4000 > 0 Then txt����.Width = 4000
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10000, "")
    cbrCustom.Handle = chk���(0).hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10001, "")
    cbrCustom.Handle = chk���(1).hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10002, "")
    cbrCustom.Handle = chk���(2).hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    chk���(2).ForeColor = vbRed
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_Help_Help, "����", True)
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10003, "")
    cbrCustom.Handle = txt����.hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    objBar.Controls.Add xtpControlButton, 10004, ""
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    With cbsMain.KeyBindings
        .Add 0, vbKeyF11, conMenu_Manage_ReportView         '����
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyI, conMenu_Edit_CopyNewItem     '����
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '�޸�
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       'ɾ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save     '����
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add 0, vbKeyF4, conMenu_View_Option                'ѡ��λ����
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With
    '------------------------------------------------------------------------------------------------------------------
    '�����˵�����
    Set mcbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "��������(&A)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸Ľ��(&R)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "��������(&M)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ�����(&D)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_ReportView, "���Ĳ���(&V)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Audit, "ͨ�����(&P)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Leave_UndoPost, "ȡ�����(&C)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Select, "ȫ��ѡ��(&L)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_DeSelect, "ȡ��ѡ��(&S)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_UnAudit, "����ѡ��(&B)")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ָ�
'==============================================================================
Private Sub InitVsf()
    Dim i           As Long
    On Error GoTo ErrH
    With fg�б�
        .Rows = 24
        .Cell(flexcpText, 1, 1) = "����ID"
        .Cell(flexcpText, 2, 1) = "·��ͼ��"
        .Cell(flexcpText, 3, 1) = "סԺ��"
        .Cell(flexcpText, 4, 1) = "סԺ����"
        .Cell(flexcpText, 5, 1) = "����"
        .Cell(flexcpText, 6, 1) = "�Ա�"
        .Cell(flexcpText, 7, 1) = "סԺҽʦ"
        .Cell(flexcpText, 8, 1) = "����ҽʦ"
        .Cell(flexcpText, 9, 1) = "���λ�ʿ"
        .Cell(flexcpText, 10, 1) = "��Ժ����"
        .Cell(flexcpText, 11, 1) = "��Ժ����"
        .Cell(flexcpText, 12, 1) = "��Ժ����"
        .Cell(flexcpText, 13, 1) = "��Ժ����"
        .Cell(flexcpText, 14, 1) = "��Ŀ����"
        .Cell(flexcpText, 15, 1) = "������"
        .Cell(flexcpText, 16, 1) = "����ʱ��"
        .Cell(flexcpText, 17, 1) = "�����"
        .Cell(flexcpText, 18, 1) = "���ʱ��"
        .Cell(flexcpText, 19, 1) = "�ܷ�"
        .Cell(flexcpText, 20, 1) = "�ȼ�"
        .Cell(flexcpText, 21, 1) = "�����޸�"
        .Cell(flexcpText, 22, 1) = "��ע"
        .Cell(flexcpText, 23, 1) = "��������"
        .Cell(flexcpChecked, 1, 0, .Rows - 1, 0) = flexUnchecked
        .Editable = flexEDKbdMouse
    End With
    For i = 1 To fg�б�.Rows - 1
        If fg����_S.ColWidth(fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1))) < 100 Then
            fg�б�.Cell(flexcpChecked, i, 0) = flexUnchecked
        Else
            fg�б�.Cell(flexcpChecked, i, 0) = flexChecked
        End If
    Next
    fg�б�.ZOrder 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ָ�
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo ErrH
    
    Select Case Item.ID
        Case 1
            Item.Handle = picLeft_S.hWnd
        Case 2
            Item.Handle = picRight.hWnd
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ؿ�������
'==============================================================================
Private Sub InitTvw()
    Dim rsTemp      As ADODB.Recordset
    Dim strTmp      As String
    Dim varTmp      As Variant
    Dim varAry      As Variant
    Dim lngCount    As Long
    Dim strDept     As String
    Dim intCol      As Integer
    
    On Error GoTo ErrH
    
    '�г����ű�Ͷ�Ӧ��Ա
    strTmp = GetPara("���ֿ��ҷ�Χ", mlngModule)
    varTmp = Split(strTmp, ";")
    strDept = ""
    For lngCount = 0 To UBound(varTmp)
        varAry = Split(varTmp(lngCount), ",")
        If UserInfo.ID = varAry(0) Then
            strDept = varTmp(lngCount)
            strDept = Mid(strDept, InStr(1, strDept, ",")) & ","
            Exit For
        End If
    Next
    
    tvw����.Nodes.Clear
    gstrSQL = "select id,����,����,�ϼ�id From ���ű� where ( TO_CHAR (����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or ����ʱ�� is null)" & _
             " start with �ϼ�id is null connect by prior id = �ϼ�id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If IsPrivs(mstrPrivs, "���п���") Then tvw����.Nodes.Add , , "C0", "��+�����п���", "Search", "Search"
    If strDept = "" Then
        mblnSetDept = False
        Do Until rsTemp.EOF
            If IsNull(rsTemp("�ϼ�id")) Then
                tvw����.Nodes.Add , , "C" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "Search", "Search"
            Else
                tvw����.Nodes.Add "C" & rsTemp("�ϼ�id"), tvwChild, "C" & rsTemp("id").Value, "��" & rsTemp("����") & "��" & rsTemp("����"), "Dot", "Dot"
            End If
            rsTemp.MoveNext
        Loop
    ElseIf IsPrivs(mstrPrivs, "���п���") Then
        mblnSetDept = True
        Do Until rsTemp.EOF
            If InStr(1, strDept, "," & rsTemp("id").Value & ",") Then
                tvw����.Nodes.Add "C0", tvwChild, "C" & rsTemp("id").Value, "��" & rsTemp("����") & "��" & rsTemp("����"), "Dot", "Dot"
            End If
            rsTemp.MoveNext
        Loop
    End If
    mblnSetDept = False
    rsTemp.Close
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��������
'==============================================================================
Private Sub mfrm���ֽ���༭_AferSaveData()
    Call ����ID
    Call Fill���ֽ��
End Sub

'==============================================================================
'=���ܣ� ��������
'==============================================================================
Private Sub RecordRating()
    On Error GoTo ErrH
    Call ����ID
    mfrm���ֽ���༭.ShowForm "����", m_lng���ID, m_lng����ID, m_lng��ҳID, m_lng����ID, Val(txt����.Tag)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �޸�����
'==============================================================================
Private Sub RecordEdit()
    
    On Error GoTo ErrH

    Call ����ID
    mfrm���ֽ���༭.ShowForm "�޸�", m_lng���ID, m_lng����ID, m_lng��ҳID, m_lng����ID, Val(txt����.Tag)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��������
'==============================================================================
Private Sub RecordReturn()
 
    On Error GoTo ErrH

    Call ����ID
    mfrm���ֽ���༭.ShowForm "����", m_lng���ID, m_lng����ID, m_lng��ҳID, m_lng����ID, Val(txt����.Tag)
 
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ɾ������
'==============================================================================
Private Sub RecordDel()
    Dim msgReturn       As VbMsgBoxResult '����Ի��򷵻�ֵ
    
    On Error GoTo ErrH
    
    Call ����ID
    msgReturn = MsgBox("��ȷ��Ҫɾ��" & fg����_S.Cell(flexcpText, fg����_S.Row, 5) & "�Ų��˵ĵ�(" & fg����_S.Cell(flexcpText, fg����_S.Row, 6) & ")��סԺ�������ֽ����¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If msgReturn = vbNo Then Exit Sub
    gstrSQL = "ZL_�������ֽ��_Delete (" & m_lng���ID & ")"
    'ע�⣺�˴�ʹ����������
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '�������б�ѡ����Ϊ�ջ���ҳTabѡ����Ϊ�գ����˳�
    If fg����_S.Row < 1 Then Exit Sub
    '�ύ����
    gcnOracle.CommitTrans
    'ˢ����ҳTAB
    Call ����ID
    Call Fill���ֽ��
    Exit Sub
ErrH:
    '�ع�����
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �鿴��ҳ
'==============================================================================
Private Sub RecordLook()
    
    On Error GoTo ErrH
                                                    
    Call ����ID
    '��ʼ���Ĳ�����ҳ
    If fg����_S.Row < 1 Then Exit Sub
    If mfrmArchiveView Is Nothing Then Set mfrmArchiveView = New frmArchiveView
    Call mfrmArchiveView.ShowArchive(Me, m_lng����ID, m_lng��ҳID, False)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������
'==============================================================================
Private Sub RecordAudit()
    Dim rs              As ADODB.Recordset
    Dim msgReturn       As VbMsgBoxResult '����Ի��򷵻�ֵ
    
    On Error GoTo ErrH
    
    Call ����ID
    msgReturn = MsgBox("��ȷ��ͨ��������ˣ�" & fg����_S.Cell(flexcpText, fg����_S.Row, 5) & "�Ų��˵ĵ�(" & fg����_S.Cell(flexcpText, fg����_S.Row, 6) & ")��סԺ�������ֽ����¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If msgReturn = vbNo Then Exit Sub
    gstrSQL = "select A.����ID,A.��ҳID,A.��Ϣ��,A.��Ϣֵ,B.�ȼ� from ������ҳ�ӱ� A,�������ֽ�� B Where A.����ID=" & m_lng����ID & " and A.��ҳID=" & m_lng��ҳID & " and A.��Ϣ��='��������' " & _
        " and B.����ID=A.����ID and B.��ҳID=A.��ҳID "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rs.EOF Then
        If rs("��Ϣֵ") <> rs("�ȼ�") Then
            If MsgBox("�ѱ�Ŀ�Ĳ����ȼ������ֵȼ�����ͬ��ȷ�ϱ�����" + vbCrLf + _
                "----------------------------------------------------------- " + vbCrLf + _
                "ǰ��Ϊ: [" + rs("��Ϣֵ") + "]������Ϊ: [" + IIf(rs("�ȼ�") = "��", "���ϸ�", rs("�ȼ�")) + "]", vbOKCancel + vbInformation + vbDefaultButton1, gstrSysName) = vbOK Then
                gstrSQL = "ZL_�������ֽ��_���" & _
                    "(" & m_lng���ID & ",'" & gstrUserName & "'," & glngSys & ")"
            Else
                Exit Sub
            End If
        Else
            gstrSQL = "ZL_�������ֽ��_���(" & m_lng���ID & ",'" & gstrUserName & "'," & glngSys & ")"
        End If
    Else
        '�ȼ���ͬ����ֱ�����ͨ����
        gstrSQL = "ZL_�������ֽ��_���" & _
            "(" & m_lng���ID & ",'" & gstrUserName & "'," & glngSys & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    'ˢ����ҳTAB
    Call ����ID
    Call Fill���ֽ��
    Call SetMenu
    
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȡ�����
'==============================================================================
Private Sub RecordUnAudit()
    Dim msgReturn As VbMsgBoxResult '����Ի��򷵻�ֵ
On Error GoTo ErrH
    Call ����ID
    msgReturn = MsgBox("��ȷ��ȡ��������ˣ�" & fg����_S.Cell(flexcpText, fg����_S.Row, 5) & "�Ų��˵ĵ�(" & fg����_S.Cell(flexcpText, fg����_S.Row, 6) & ")��סԺ�������ֽ����¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If msgReturn = vbNo Then Exit Sub
    gstrSQL = "ZL_�������ֽ��_ȡ�����" & "(" & m_lng���ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    'ˢ����ҳTAB
    Call ����ID
    Call Fill���ֽ��
    Call SetMenu
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȫ��ѡ��
'==============================================================================
Private Sub RecordSelect()
    On Error GoTo ErrH
    
    fg����_S.Cell(flexcpChecked, 0, 0, fg����_S.Rows - 1, 0) = flexChecked
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȫ�����
'==============================================================================
Private Sub RecordUnSelect()
    
    On Error GoTo ErrH
    
    fg����_S.Cell(flexcpChecked, 0, 0, fg����_S.Rows - 1, 0) = flexUnchecked
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ѡ��
'==============================================================================
Private Sub RecordSelectOther()
    Dim i As Long

    On Error GoTo ErrH
    
    fg����_S.Cell(flexcpChecked, 0, 0) = flexUnchecked
    For i = 1 To fg����_S.Rows - 1
        If fg����_S.Cell(flexcpChecked, i, 0) = flexUnchecked Then
            fg����_S.Cell(flexcpChecked, i, 0) = flexChecked
        Else
            fg����_S.Cell(flexcpChecked, i, 0) = flexUnchecked
        End If
    Next
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���˲�ѯ
'==============================================================================
Private Sub RecordFind()
    Dim strTemp         As String
    
    On Error GoTo ErrH
    With mfrm����
    
        strTemp = .GetFilter(mstrPrivs, txt����.Text)
        If .mblnCancel Then Exit Sub
        mlngSickID = .lngSickID                         '����ID
        mlngHospitalID = .lngHospitalID                 'סԺ��
        mlngHospitalTimes = .lngHospitalTimes           'סԺ����
        mstrSickName = .strSickName                     '��������
        mstrMainDoctor = .strMainDoctor                 '����ҽʦ
        mstrOutpatientDoctor = .strOutpatientDoctor     '����ҽʦ
        mstrNurses = .strNurses                         '���λ�ʿ
        mstrRatingMan = .strRatingMan                   '������
        mstrAuditMan = .strAuditMan                     '�����
        mstrOutDept = .strOutDept                       '��Ժ����
        mstrInDept = .strInDept                         '��Ժ����
        mdatStarOutDate = .datStarOutDate               '��Ժ��ʼ����
        mdatEndOutDate = .datEndOutDate                 '��Ժ��ʼ����
        mdatStarInDate = .datStarInDate                 '��Ժ��ʼ����
        mdatEndInDate = .datEndInDate                   '��Ժ��ʼ����
        mstrSickType = .strSickType                     '��������
    End With
    mstrWhere = "Where 1=1 " & strTemp
    Call mDataLoad
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������ӡ
'==============================================================================
Private Sub RecordAllPrint()
    Dim i               As Long
    Dim rs              As ADODB.Recordset
    Dim lngID           As Long
    Dim lngNum          As Long
    On Error GoTo ErrH

    If MsgBox("�Ƿ��ӡ��ǰѡ�е����в������ֽ������", vbOKCancel + vbInformation + vbDefaultButton2, gstrSysName) = vbCancel Then Exit Sub
    lngNum = 0
    For i = 1 To fg����_S.Rows - 1
        If fg����_S.Cell(flexcpChecked, i, 0) = flexChecked Then
            lngID = Val(fg����_S.Cell(flexcpText, i, 1))
            If lngID <> 0 Then
                lngNum = lngNum + 1
                stbThis.Panels(2) = "��ӡ����:" & CStr(lngNum)
                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "���ID=" & lngID, 2
            End If
        End If
    Next i
    stbThis.Panels(2) = "��ӡ��������ϣ�������" & CStr(lngNum) & "�ݡ�"
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ������仯ʱ���¼�������
'==============================================================================
Private Sub chk���_Click(Index As Integer)
    On Error GoTo ErrH
    Call mDataLoad
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�chk��� �س��൱��Tab��
'==============================================================================
Private Sub chk���_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�cmb��Χ �س��൱��Tab��
'==============================================================================
Private Sub cmb��Χ_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����ؿ���ѡ������
'==============================================================================
Private Sub DeptSelect()
    Dim i As Integer
    On Error GoTo ErrH
    If tvw����.Nodes.count = 0 Then
        Exit Sub
    End If
    tvw����.Visible = True
    tvw����.ZOrder (0)
    If tvw����.Visible Then
        '��ʾ��ǰ��Ա
        If txt����.Tag = "" Then
            tvw����.Nodes(1).Expanded = True
            tvw����.Nodes(1).Selected = True
        Else
            tvw����.Nodes("C" & txt����.Tag).Selected = True
            tvw����.SelectedItem.EnsureVisible
        End If
        tvw����.SetFocus
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��Ҽ��˵������ֿ�ݲ˵���
'==============================================================================
Private Sub fg����_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If Button = 2 Then
        If fg����_S.MouseRow = -1 And fg����_S.Rows >= 1 Then
            fg����_S.Row = fg����_S.Rows - 1
        ElseIf fg����_S.MouseRow = 0 And fg����_S.Rows > 1 Then
            fg����_S.Row = 1
        Else
            fg����_S.Row = fg����_S.MouseRow
        End If
        fg����_S.Col = fg����_S.MouseCol
        mcbrPopupBar.ShowPopup
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����в����϶�λ��
'==============================================================================
Private Sub fg����_S_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    On Error GoTo ErrH
    If Col = 0 Then
        Position = -1
    Else
        If Position <= 0 Then Position = Col
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ĳ�в����϶���С fg����_S[ͼ��]
'==============================================================================
Private Sub fg����_S_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    If Col = 0 Then Cancel = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg���� ���б䶯�����״̬��
'==============================================================================
Private Sub fg����_S_RowColChange()
    On Error GoTo ErrH
    Call fg����_S_SelChange
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg���� ѡ���б䶯�����״̬��
'==============================================================================
Private Sub fg����_S_SelChange()
    Dim strTag As String
    On Error GoTo ErrH
    If fg����_S.Rows < 1 Then
        fg���_S.Rows = 1
        lbl�ܷ�.Caption = "�ܷ�:"
        lbl�ȼ�.Caption = "�ȼ�:"
        lbl������.Caption = "������:"
        lbl�����.Caption = "�����:"
        lbl�����޸�.Caption = ""
        lbl��������.Caption = "��������:"
        lbl��ע.Caption = "��ע:"
        lbl����ʱ��.Caption = "����ʱ��:"
        lbl���ʱ��.Caption = "���ʱ��:"
        Exit Sub
    End If
    strTag = fg����_S.Cell(flexcpText, fg����_S.Row, 4) & "_" & fg����_S.Cell(flexcpText, fg����_S.Row, 6)
    If fg����_S.Tag <> strTag Then
        fg����_S.Tag = strTag
        Call Fill���ֽ��
    End If
    m_lngOldRow = fg����_S.Row
    mRecordRating = (fg����_S.TextMatrix(fg����_S.Row, fg����_S.ColIndex("����ʱ��")) <> "")
    mRecordAudit = (fg����_S.TextMatrix(fg����_S.Row, fg����_S.ColIndex("���ʱ��")) <> "")
    mRecordMyAudit = (fg����_S.TextMatrix(fg����_S.Row, fg����_S.ColIndex("������")) = UserInfo.����)
    gstrSQL = "select Count(1) from �������ַ��� where ѡ��=1 and ID = [1]"
    mRecordReturn = (zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(fg����_S.TextMatrix(fg����_S.Row, fg����_S.ColIndex("����ID")))).Fields(0) > 0)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȫѡ��ȫ��
'==============================================================================
Private Sub fg����_S_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    
    If Col = 0 And Row = 0 Then
        If fg����_S.Cell(flexcpChecked, 0, 0, 0, 0) = flexUnchecked Then
            fg����_S.Cell(flexcpChecked, 0, 0, fg����_S.Rows - 1, 0) = flexChecked
        Else
            fg����_S.Cell(flexcpChecked, 0, 0, fg����_S.Rows - 1, 0) = flexUnchecked
        End If
        Cancel = True
    ElseIf Col <> 0 Then
        Cancel = True
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��Ҽ��˵������ֿ�ݲ˵���
'==============================================================================
Private Sub fg���_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If Button = 2 Then
        If fg���_S.MouseRow = -1 And fg���_S.Rows >= 1 Then
            fg���_S.Row = fg���_S.Rows - 1
        ElseIf fg���_S.MouseRow = 0 And fg���_S.Rows > 1 Then
            fg���_S.Row = 1
        Else
            fg���_S.Row = fg���_S.MouseRow
        End If
        fg���_S.Col = fg���_S.MouseCol
        mcbrPopupBar.ShowPopup
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg�б� ����ֵ���ж�
'==============================================================================
Private Sub fg�б�_Click()
    Dim i           As Long
    On Error GoTo ErrH
    For i = 1 To fg�б�.Rows - 1
        If fg�б�.Cell(flexcpChecked, i, 0) = flexChecked Then
            If fg����_S.ColWidth(fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1))) < 100 Then
                
                If fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1)) = 3 Then
                    fg����_S.ColWidth(fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1))) = 300
                Else
                    fg����_S.ColWidth(fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1))) = 1000
                End If
            End If
        Else
            If fg����_S.ColWidth(fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1))) > 100 Then
                fg����_S.ColWidth(fg����_S.ColIndex(fg�б�.Cell(flexcpText, i, 1))) = 0
            End If
        End If
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg�б� ��ESCʱ�����б�
'==============================================================================
Private Sub fg�б�_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 27 Then fg�б�.Visible = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg�б� ��ESCʱ�����б�
'==============================================================================
Private Sub fg�б�_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    If Col <> 0 Then
        Cancel = True
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ��ؼ���ʼ��
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo ErrH
    Call InitCommonControls
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ���ʼ��
'==============================================================================
Private Sub Form_Load()
    Dim strKey              As String
    On Error GoTo ErrH
    m_lngOldRow = -1
    mstrPrivs = UserInfo.ģ��Ȩ��
    mlngModule = ParamInfo.ģ���
    mblnSetDept = False
    Set mfrm���ֽ���༭ = New frm���ֽ���༭
    '�ؼ���ʼ��
    mstrFindKey = "סԺ��"
    If GetPersonSet Then
        mstrFindKey = Trim(GetPara("��λ��Χ", mlngModule, "סԺ��", True))
        chk���(0).Value = zlDatabase.GetPara("δ����", glngSys, mlngModule, vbChecked)
        chk���(1).Value = zlDatabase.GetPara("δ���", glngSys, mlngModule, vbChecked)
        chk���(2).Value = zlDatabase.GetPara("�����", glngSys, mlngModule, vbChecked)
    End If
                                       
    Call InitControl
    
    '��ʼ�����Ҵ���
    Set mfrm���� = New frm�������ֲ�ѯ
    Load mfrm����
    mdatStarOutDate = DateAdd("M", -1, Date)
    If mfrm����.mbln��Ŀ������ Then
        mstrWhere = " Where ��Ŀ���� is not null and ��Ժ���� >= [12]"
    Else
        mstrWhere = " Where ��Ժ���� >= [12]"
    End If
    If IsPrivs(mstrPrivs, "���п���") Then
        txt����.Text = "���п���"
        Call InitTvw
    Else
        
        txt����.Text = Get��������(UserInfo.ID, 0)
        txt����.Locked = True
        txt����.BackColor = &H80000000
        mstrOutDept = UserInfo.��������
        
        If txt����.Text <> mstrOutDept Then
            mstrWhere = mstrWhere & " And ��Ժ���� In (" & Get��������(UserInfo.ID, 1) & ")"
        Else
            mstrWhere = mstrWhere & " And ��Ժ���� = [11]"
        End If
    End If
    Call mDataLoad
    stbThis.Panels(2) = "��ǰ��ʾ��" & fg����_S.Rows - 1 & "�ݲ�����"
    strKey = Me.fg����_S.Tag
    Me.fg����_S.Tag = ""
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL1_INSIDE_1562_1")
    fg����_S.Tag = strKey
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����ڹر�ʱ�ر��Ӵ���
'==============================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrH
    If Not (mfrmArchiveView Is Nothing) Then Unload mfrmArchiveView
    Set mfrmArchiveView = Nothing
    mfrm����.mblnForce = True   'ǿ�ƹرգ���ʽ�رգ�����ֻ������֮��
    Unload mfrm����
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����ڴ�С�䶯ʱ�ؼ��仯
'==============================================================================
Private Sub Form_Resize()
    On Error GoTo ErrH

    Call SetPaneRange(dkpMain, 1, 100, 100, ScaleHeight - 200, ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 400, 100, ScaleHeight - 200, ScaleHeight)
    With tvw����
        .Move txt����.Left, stbThis.Height * 2 + txt����.Top + txt����.Height + 70, txt����.Width, 4000
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ��ر�ʱ��������
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrH
    If Not (mfrm���ֽ���༭ Is Nothing) Then Unload mfrm���ֽ���༭
    If Not (mfrmArchiveView Is Nothing) Then Unload mfrmArchiveView
    Set mfrmArchiveView = Nothing
    Call SetPara("δ����", chk���(0).Value, mlngModule)
    Call SetPara("δ���", chk���(1).Value, mlngModule)
    Call SetPara("�����", chk���(2).Value, mlngModule)
    Call SetPara("��λ��Χ", mstrFindKey, mlngModule)
    Me.fg����_S.Tag = ""
    SaveWinState Me, App.ProductName
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg�б� ��ʾ�ֶ��б�
'==============================================================================
Private Sub imgSelCols_S_Click()
    On Error GoTo ErrH
    Call InitTvw
    fg�б�.Visible = Not fg�б�.Visible
    If fg�б�.Visible Then fg�б�.SetFocus
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�fg�б� ʧȥ����ʱ����
'==============================================================================
Private Sub fg�б�_LostFocus()
    On Error GoTo ErrH
    fg�б�.Visible = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�������ӡ
'==============================================================================
Private Sub mnuFilePrintALL_Click()
    Dim i               As Long
    Dim rs              As ADODB.Recordset
    Dim lngID           As Long
    Dim lngNum          As Long
    On Error GoTo ErrH
    If MsgBox("�Ƿ��ӡ��ǰѡ�е����в������ֽ������", vbOKCancel + vbInformation + vbDefaultButton2, gstrSysName) = vbCancel Then Exit Sub
    lngNum = 0
    For i = 1 To fg����_S.Rows - 1
        If fg����_S.Cell(flexcpChecked, i, 0) = flexChecked Then
            lngID = Val(fg����_S.Cell(flexcpText, i, 1))
            If lngID <> 0 Then
                lngNum = lngNum + 1
                stbThis.Panels(2) = "��ӡ����:" & CStr(lngNum)
                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "���ID=" & lngID, 2
            End If
        End If
    Next i
    stbThis.Panels(2) = "��ӡ��������ϣ�������" & CStr(lngNum) & "�ݡ�"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����λ��λ�ñ仯�������
'==============================================================================
Private Sub picLeft_S_Resize()
On Error Resume Next
    imgSelCols_S.Move IIf(imgSelCols_S.Left < 1350, 1350, picLeft_S.ScaleWidth - imgSelCols_S.Width - 100)
    fg�б�.Move Abs(picLeft_S.Width - fg�б�.Width), picLeft_S.Top + imgSelCols_S.Top + imgSelCols_S.Height + 45
    imgRefresh.Move imgSelCols_S.Left - imgRefresh.Width - 175
    picLeft_S.Cls
    picLeft_S.PaintPicture imgBGBlue.Picture, Screen.TwipsPerPixelX, 0, picLeft_S.Width, 360, 0, 0, imgBGBlue.Width, 360
    picLeft_S.PaintPicture imgBGBlue.Picture, Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picLeft_S.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    picLeft_S.PaintPicture imgBGBlue.Picture, picLeft_S.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picLeft_S.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    picLeft_S.PaintPicture imgBGBlue.Picture, Screen.TwipsPerPixelX, picLeft_S.ScaleHeight - Screen.TwipsPerPixelY, picLeft_S.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    fg����_S.Move picLeft_S.Left + 40, fg����_S.Top, picLeft_S.Width - 60, picLeft_S.Height - fg����_S.Top - 420
    Refresh
End Sub

'==============================================================================
'=���ܣ��һ�ʱ�����˵�
'==============================================================================
Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If Button = 2 Then mcbrPopupBar.ShowPopup
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��Ҳ�ؼ�λ�õ���
'==============================================================================
Private Sub picRight_Resize()
    On Error GoTo ErrH
    fg���_S.Move fg���_S.Left, fg���_S.Top, picRight.Width - 400, picRight.Height - fg���_S.Top - 450
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��������ݼ���
'==============================================================================
Private Sub mDataLoad()
    Dim strWhereDept        As String
    Dim strWhere            As String
    Dim blnA                As Boolean
    Dim blnB                As Boolean
    Dim blnC                As Boolean
    Dim i                   As Integer
    
    On Error GoTo ErrH
    
    strWhereDept = "A.��Ժ���� is not null"
    blnA = (chk���(0).Value = vbChecked)
    blnB = (chk���(1).Value = vbChecked)
    blnC = (chk���(2).Value = vbChecked)
    
    strWhere = "("
    If blnA Then
        strWhere = strWhere & "A.����ʱ�� is null"
    End If
    If blnB Then
        If strWhere = "(" Then
            strWhere = strWhere & "A.���ʱ�� is null"
        Else
            strWhere = strWhere & " or A.���ʱ�� is null"
        End If
    End If
    If blnC Then
        If strWhere = "(" Then
            strWhere = strWhere & "A.���ʱ�� is not null"
        Else
            strWhere = strWhere & " or A.���ʱ�� is not null"
        End If
    End If
 
    If strWhere <> "(" Then
        strWhere = strWhereDept & " And " & strWhere & ")"
    Else
        strWhere = strWhereDept
    End If
    If Trim(mstrWhere) = "" Then mstrWhere = "1=1"
    strWhere = IIf(InStr(LCase(mstrWhere), "where") > 0, mstrWhere, " where " & mstrWhere) & " And " & strWhere
    
    gstrSQL = "" & _
        "   Select A.סԺ��, A.����, A.�Ա�,  Decode((select Count(*) from �����ٴ�·�� where ����ID = A.����ID and ��ҳid = A.��ҳID),0,'','lujin') as ·��,A.����id, A.��ҳid, A.��Ժ����, A.��Ժ����, A.��Ժ����, A.��Ժ����, A.����ҽʦ, A.���λ�ʿ, A.סԺҽʦ," & _
        "           A.��Ŀ����, A.���id, A.����id, A.�ܷ�, A.�ȼ�, A.������, A.����ʱ��, A.�����, A.���ʱ��, A.�����޸�, A.��ע,A.�������� " & _
        "   from ��������������ͼ A " & strWhere

    Set rsM = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngSickID, mlngHospitalID, mlngHospitalTimes, mstrSickName, mstrMainDoctor, mstrOutpatientDoctor, mstrNurses, mstrRatingMan, mstrInDept, mstrAuditMan, mstrOutDept, CDate(mdatStarOutDate), CDate(mdatEndOutDate), CDate(mdatStarInDate), CDate(mdatEndInDate), mstrSickType)
    rsM.Sort = "���ʱ�� desc,����ʱ�� desc,��Ժ���� desc,סԺҽʦ,����"
    
    If (txt����.Text = "���п���" Or txt����.Text = "") And Not mblnSetDept Then
        rsM.Filter = ""
    Else
        If txt����.Text = UserInfo.�������� Then
            rsM.Filter = "��Ժ����='" & txt����.Text & "'"
        End If
    End If
    
    Call Fill����
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��������������Ϣ
'==============================================================================
Private Sub Fill����()
    Dim lngIndex            As Long
    Dim lng����״̬         As Long
    Dim sngForeColor        As ColorConstants
    Dim bln����             As Boolean
    Dim str�ȼ�             As String
    Dim i                   As Long
    Dim j                   As Long
    
    On Error GoTo ErrH
    
    With fg����_S
        .Editable = flexEDKbdMouse
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 26
        .Clear
        
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpText, 0, 1) = "���ID"
        .Cell(flexcpText, 0, 2) = "����ID"
        .Cell(flexcpPicture, 0, 3) = imgLujinPic.Picture
        .Cell(flexcpText, 0, 4) = "����ID"
        .Cell(flexcpText, 0, 5) = "סԺ��"
        .Cell(flexcpText, 0, 6) = "סԺ����"
        .Cell(flexcpText, 0, 7) = "����"
        .Cell(flexcpText, 0, 8) = "�Ա�"
        .Cell(flexcpText, 0, 9) = "סԺҽʦ"
        .Cell(flexcpText, 0, 10) = "����ҽʦ"
        .Cell(flexcpText, 0, 11) = "���λ�ʿ"
        .Cell(flexcpText, 0, 12) = "��Ժ����"
        .Cell(flexcpText, 0, 13) = "��Ժ����"
        .Cell(flexcpText, 0, 14) = "��Ժ����"
        .Cell(flexcpText, 0, 15) = "��Ժ����"
        .Cell(flexcpText, 0, 16) = "��Ŀ����"
        .Cell(flexcpText, 0, 17) = "������"
        .Cell(flexcpText, 0, 18) = "����ʱ��"
        .Cell(flexcpText, 0, 19) = "�����"
        .Cell(flexcpText, 0, 20) = "���ʱ��"
        .Cell(flexcpText, 0, 21) = "�ܷ�"
        .Cell(flexcpText, 0, 22) = "�ȼ�"
        .Cell(flexcpText, 0, 23) = "�����޸�"
        .Cell(flexcpText, 0, 24) = "��ע"
        .Cell(flexcpText, 0, 25) = "��������"
        DoEvents
        .FocusRect = flexFocusSolid
        '��������
        .Rows = IIf(rsM.RecordCount < 1000, rsM.RecordCount + 1, 1001)
        i = 1
        Do Until rsM.EOF
        
            If i >= 1001 And bln���� = False Then
                If MsgBox("�Ѿ�װ��1000�ݲ���������" & rsM.RecordCount - .Rows + 1 & "�ݴ�װ��" & vbCrLf & _
                    "�Ƿ������", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                    
                    Exit Do
                End If
                .Rows = rsM.RecordCount + 1
                bln���� = True
            End If
            
            If Trim(rsM("���ʱ��")) <> "" Then
                lng����״̬ = 2
                sngForeColor = RGB(180, 180, 180)
            ElseIf Trim(rsM("����ʱ��")) <> "" Then
                lng����״̬ = 1
                sngForeColor = RGB(0, 0, 255)
            Else
                lng����״̬ = 0
                sngForeColor = vbBlack
            End If
            
            .Cell(flexcpText, i, 1) = NVL(rsM("���ID"), 0)
            .Cell(flexcpText, i, 2) = NVL(rsM("����ID"), 0)
            .Cell(flexcpPicture, i, 3) = IIf(NVL(rsM("·��")) = "", "", imgLujin.Picture)
            .Cell(flexcpText, i, 4) = NVL(rsM("����ID"), 0)
            
            .Cell(flexcpText, i, 5) = NVL(rsM("סԺ��"), 0)
            .Cell(flexcpText, i, 6) = NVL(rsM("��ҳID"))
            .Cell(flexcpText, i, 7) = NVL(rsM("����"))
            .Cell(flexcpText, i, 8) = NVL(rsM("�Ա�"))
            .Cell(flexcpText, i, 9) = NVL(rsM("סԺҽʦ"))
            .Cell(flexcpText, i, 10) = NVL(rsM("����ҽʦ"))
            .Cell(flexcpText, i, 11) = NVL(rsM("���λ�ʿ"))
            .Cell(flexcpText, i, 12) = NVL(rsM("��Ժ����"))
            .Cell(flexcpText, i, 13) = IIf(IsNull(rsM("��Ժ����")), "", Format(rsM("��Ժ����"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 14) = NVL(rsM("��Ժ����"))
            .Cell(flexcpText, i, 15) = IIf(IsNull(rsM("��Ժ����")), "", Format(rsM("��Ժ����"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 16) = IIf(IsNull(rsM("��Ŀ����")), "", Format(rsM("��Ŀ����"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 17) = NVL(rsM("������"))
            .Cell(flexcpText, i, 18) = IIf(IsNull(rsM("����ʱ��")), "", Format(rsM("����ʱ��"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 19) = NVL(rsM("�����"))
            .Cell(flexcpText, i, 20) = IIf(IsNull(rsM("���ʱ��")), "", Format(rsM("���ʱ��"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 21) = IIf(NVL(rsM("�ȼ�")) = "��", "", NVL(rsM("�ܷ�")))
            .Cell(flexcpText, i, 24) = NVL(rsM!��ע)
            .Cell(flexcpText, i, 25) = NVL(rsM!��������)
            Select Case NVL(rsM("�ȼ�"))
                Case "��"
                    str�ȼ� = "�׼�"
                Case "��"
                    str�ȼ� = "�Ҽ�"
                Case "��"
                    str�ȼ� = "����"
                Case "��"
                    str�ȼ� = "���ϸ�"
                Case Else
                    str�ȼ� = ""
            End Select
            .Cell(flexcpText, i, 22) = str�ȼ�
            If NVL(rsM("�����޸�"), 0) = 0 Then
                .Cell(flexcpText, i, 23) = ""
            Else
                .Cell(flexcpText, i, 23) = "��"
            End If
             
            For j = 1 To 25
                .Cell(flexcpForeColor, i, j) = sngForeColor
            Next
            
            rsM.MoveNext
            i = i + 1
        Loop
        .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
            
        If Me.Tag = "" Then
            .ColWidth(.ColIndex("ICON")) = 300
            .ColWidth(.ColIndex("���ID")) = 0
            .ColWidth(.ColIndex("����ID")) = 0
            .ColWidth(.ColIndex("·��ͼ��")) = 300
            
            .ColWidth(.ColIndex("����ID")) = 0
            .ColWidth(.ColIndex("סԺ����")) = 400
            .ColWidth(.ColIndex("סԺ��")) = 650
            .ColWidth(.ColIndex("����")) = 900
            .ColWidth(.ColIndex("�Ա�")) = 600
            .ColWidth(.ColIndex("סԺҽʦ")) = 900
            .ColWidth(.ColIndex("����ҽʦ")) = 0
            .ColWidth(.ColIndex("���λ�ʿ")) = 0
            .ColWidth(.ColIndex("��Ժ����")) = 900
            .ColWidth(.ColIndex("��Ժ����")) = 1600
            .ColWidth(.ColIndex("��Ժ����")) = 0
            .ColWidth(.ColIndex("��Ժ����")) = 0
            .ColWidth(.ColIndex("��Ŀ����")) = 0
            .ColWidth(.ColIndex("������")) = 0
            .ColWidth(.ColIndex("����ʱ��")) = 0
            .ColWidth(.ColIndex("�����")) = 0
            .ColWidth(.ColIndex("���ʱ��")) = 0
            .ColWidth(.ColIndex("�ܷ�")) = 0
            .ColWidth(.ColIndex("�ȼ�")) = 0
            .ColWidth(.ColIndex("�����޸�")) = 0
            .ColWidth(.ColIndex("��ע")) = 0
            .ColWidth(.ColIndex("��������")) = 0
            .ColAlignment(.ColIndex("�����޸�")) = flexAlignCenterCenter
            Me.Tag = "�Ѿ������п�"
        End If
        
        .ColAlignment(.ColIndex("�Ա�")) = flexAlignCenterCenter
        '�и�����
        .RowHeightMin = 300
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        
        'ѡ����ǰ����
        If m_lngOldRow > 0 And m_lngOldRow < i Then
            .Row = m_lngOldRow
            .Col = 2
            .ShowCell m_lngOldRow, 2
            On Error Resume Next
            If .Visible = True Then .SetFocus
            Call fg����_S_SelChange
        ElseIf .Tag = "" And i > 1 And .Rows > 1 Then
            m_lngOldRow = 1
            .Tag = "ѡ�е�һ��"
            .Row = 1
            .Col = 2
            .ShowCell m_lngOldRow, 2
            If .Visible = True Then .SetFocus
            Call fg����_S_SelChange
        Else
            If .Rows > 1 Then .Row = 1
            Call fg����_S_SelChange
        End If
     End With

    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȡ�õ�ǰ���ֽ��ID�ͷ���ID
'==============================================================================
Private Sub ����ID()
    Dim i                   As Long
    Dim sngForeColor        As ColorConstants
    Dim lng����״̬         As Long
    On Error GoTo ErrH
    With fg����_S
        m_lng����ID = Val(.Cell(flexcpText, .Row, 4))
        m_lng��ҳID = Val(.Cell(flexcpText, .Row, 6))
        
        Dim rs As New ADODB.Recordset, str�ȼ� As String
        gstrSQL = " " & _
            "   Select סԺ��, ����, �Ա�, ����id, ��ҳid, ��Ժ����, ��Ժ����, ��Ժ����, ��Ժ����, ����ҽʦ, ���λ�ʿ, סԺҽʦ," & _
            "           ��Ŀ����, ���id, ����id, �ܷ�, �ȼ�, ������, ����ʱ��, �����, ���ʱ��, �����޸�, ��ע,�������� " & _
            "   From ��������������ͼ " & _
            "   where ����ID=[1] and ��ҳID=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID, m_lng��ҳID)
        
        If Not rs.EOF Then
            If Trim(rs("���ʱ��")) <> "" Then
                lng����״̬ = 2
                sngForeColor = RGB(180, 180, 180)
            ElseIf Trim(rs("����ʱ��")) <> "" Then
                lng����״̬ = 1
                sngForeColor = RGB(0, 0, 255)
            Else
                lng����״̬ = 0
                sngForeColor = vbBlack
            End If
            
            For i = 1 To 21
                .Cell(flexcpForeColor, .Row, i) = sngForeColor
            Next
            
            .Cell(flexcpText, .Row, 1) = NVL(rs("���ID"))
            .Cell(flexcpText, .Row, 2) = NVL(rs("����ID"))
            .Cell(flexcpText, .Row, 17) = NVL(rs("������"))
            .Cell(flexcpText, .Row, 18) = NVL(rs("����ʱ��"))
            .Cell(flexcpText, .Row, 19) = NVL(rs("�����"))
            .Cell(flexcpText, .Row, 20) = NVL(rs("���ʱ��"))
            .Cell(flexcpText, .Row, 21) = IIf(NVL(rs("�ȼ�")) = "��", "", NVL(rs("�ܷ�")))
            .Cell(flexcpText, .Row, 23) = IIf(NVL(rs("�����޸�"), 0) = 0, "", "��")
            .Cell(flexcpText, .Row, 24) = NVL(rs("��ע"))
            .Cell(flexcpText, .Row, 25) = NVL(rs("��������"))
            
            Select Case NVL(rs("�ȼ�"))
                Case "��"
                    str�ȼ� = "�׼�"
                Case "��"
                    str�ȼ� = "�Ҽ�"
                Case "��"
                    str�ȼ� = "����"
                Case "��"
                    str�ȼ� = "���ϸ�"
                Case Else
                    str�ȼ� = ""
            End Select
            .Cell(flexcpText, .Row, 22) = str�ȼ�
            '��ɫ����
        End If
        rs.Close
        m_lng���ID = Val(.Cell(flexcpText, .Row, 1))
        m_lng����ID = Val(.Cell(flexcpText, .Row, 2))
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=����:�����ݱ���д�ӡ,Ԥ���������EXCEL
'=����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'==============================================================================
Private Sub subPrint(ByVal bytMode As Byte)
    On Error GoTo ErrH
    Select Case bytMode
        Case 1  'Print
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "���ID=" & m_lng���ID, 2
        Case 2  'Preview
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "���ID=" & m_lng���ID, 1
        Case 3  'Excel
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "���ID=" & m_lng���ID, 3
    End Select
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����ѡ��Ķ������õ�ǰ�˵�����β���Ҫ�Ĳ˵���
'==============================================================================
Public Sub SetMenu()
    On Error GoTo ErrH
    Call ����ID
    If Trim(fg����_S.Cell(flexcpText, fg����_S.Row, 18)) <> "" Then
        mRecordRating = True
    End If
    If Trim(fg����_S.Cell(flexcpText, fg����_S.Row, 20)) <> "" Then
        mRecordAudit = True
    End If
    stbThis.Panels(2) = "��ǰ��ʾ��" & fg����_S.Rows - 1 & "�ݲ�����"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��������ַ���ID���������ֱ�׼����
'==============================================================================
Private Sub Fill���ֱ�׼()
    Dim rsTemp          As ADODB.Recordset
    Dim lngIndex        As Long
    Dim lng�ɷ��޸�     As Long
    Dim i               As Long
    On Error GoTo ErrH
    Call ����ID
    With fg���_S
        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cell(flexcpText, 0, 0) = "��Ŀ"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "��׼��ֵ"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "ȱ������"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3) = "���ֱ�׼"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4) = "����"
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 5) = "�ɷ��޸�"
        .Cell(flexcpText, 0, 6) = "ID"
        .Cell(flexcpText, 0, 7) = "�ϼ�ID"
        .Cell(flexcpText, 0, 8) = "����ID"
        .Cell(flexcpText, 0, 9) = "��ע"

        
        'ȷ����������
        If m_lng����ID < 1 Then .Redraw = flexRDDirect: Exit Sub
        gstrSQL = "" & _
            "   Select �ϼ����, ���, Id, �ϼ�id, ����id, ��Ŀ, ��׼��ֵ, ����Ҫ��, ȱ������, �۷ֱ�׼, ���� " & _
            "   From �������ֱ�׼��ͼ " & _
            "   Where ����='��' and ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
        
        .FocusRect = flexFocusSolid
        '��������
        .Cols = 10
        .Rows = rsTemp.RecordCount + 1
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, 0) = IIf(IsNull(rsTemp.Fields("��Ŀ")), "", rsTemp.Fields("��Ŀ"))
            .Cell(flexcpAlignment, i, 0) = flexAlignCenterCenter
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("��׼��ֵ")), " ", Format(rsTemp.Fields("��׼��ֵ"), "####��"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
            .Cell(flexcpText, i, 2) = IIf(IsNull(rsTemp.Fields("ȱ������")), "", rsTemp.Fields("ȱ������"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftTop
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("�۷ֱ�׼")), "", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�׼�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�Ҽ�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "����", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "������", rsTemp.Fields("�۷ֱ�׼"))))))
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            .Cell(flexcpText, i, 4) = ""
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            .Cell(flexcpText, i, 5) = ""
            .Cell(flexcpText, i, 6) = IIf(IsNull(rsTemp.Fields("ID")), "", rsTemp.Fields("ID"))
            .Cell(flexcpText, i, 7) = IIf(IsNull(rsTemp.Fields("�ϼ�ID")), "", rsTemp.Fields("�ϼ�ID"))
            .Cell(flexcpText, i, 8) = IIf(IsNull(rsTemp.Fields("����ID")), "", rsTemp.Fields("����ID"))
            .Cell(flexcpText, i, 9) = ""
            
            rsTemp.MoveNext
            i = i + 1
        Loop
        '�Զ�����
        .WordWrap = True
        '�ϲ���Ԫ��
        .MergeCells = 2
        .MergeCol(.ColIndex("��Ŀ")) = True
        .MergeCol(.ColIndex("��׼��ֵ")) = True
        '��������
        .ColAlignment(.ColIndex("��Ŀ")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("��׼��ֵ")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("���ֱ�׼")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("�ɷ��޸�")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("��ע")) = flexAlignLeftCenter
        
        '���ص�Ԫ��
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("�ϼ�ID")) = 0
        .ColWidth(.ColIndex("����ID")) = 0
        '�������
'        .ColWidth(.ColIndex("��Ŀ")) = 1500
'        .ColWidth(.ColIndex("��׼��ֵ")) = 850
'        .ColWidth(.ColIndex("ȱ������")) = 3000
'        .ColWidth(.ColIndex("���ֱ�׼")) = 1100
'        .ColWidth(.ColIndex("����")) = 800
'        .ColWidth(.ColIndex("�ɷ��޸�")) = 800
        '�и�����
        .RowHeightMin = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("ȱ������")
        .AllowBigSelection = False
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�װ���Ӧ��ҳ�����ֽ��
'==============================================================================
Private Function Fill���ֽ��() As Boolean
    Dim rs              As ADODB.Recordset
    Dim i               As Long
    Dim bln�۷���       As Boolean
    Dim intSign         As Long

    On Error GoTo ErrH

    Call Fill���ֱ�׼
    With fg����_S
        lbl������Ϣ = "����:" & .Cell(flexcpText, .Row, 7) & ",��" & .Cell(flexcpText, .Row, 6) & "��סԺ"
        lbl�ܷ� = "�ܷ�:" & .Cell(flexcpText, .Row, 21)
        lbl�ȼ� = "�ȼ�:" & .Cell(flexcpText, .Row, 22)
        lbl������ = "������:" & .Cell(flexcpText, .Row, 17)
        lbl����ʱ�� = "����ʱ��:" & .Cell(flexcpText, .Row, 18)
        lbl����� = "�����:" & .Cell(flexcpText, .Row, 19)
        lbl���ʱ�� = "���ʱ��:" & .Cell(flexcpText, .Row, 0)
        lbl�����޸� = IIf(.Cell(flexcpText, .Row, 23) = "", "", "�������޸ġ�")
        lbl��ע.Caption = "��ע:" & .Cell(flexcpText, .Row, .ColIndex("��ע"))
        lbl��������.Caption = "��������:" & .Cell(flexcpText, .Row, .ColIndex("��������"))
    End With
    fg���_S.Redraw = flexRDNone
    For i = 1 To fg���_S.Rows - 1
        fg���_S.Cell(flexcpText, i, 4) = ""
        fg���_S.Cell(flexcpText, i, 5) = ""
        fg���_S.Cell(flexcpText, i, 9) = ""
    Next
    'ȷ������
    gstrSQL = "select ���� from �������ַ��� where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
    bln�۷��� = True
    If Not rs.EOF Then
        bln�۷��� = IIf(NVL(rs("����"), "�ӷ���") = "�ӷ���", False, True)
    End If
    rs.Close
    If bln�۷��� Then
        intSign = -1
    Else
        intSign = 1
    End If
    gstrSQL = "" & _
        "   Select A.ID,A.��Ŀ,A.��׼��ֵ,A.����Ҫ��,A.ȱ������,A.�۷ֱ�׼," & _
        "           (select decode(ȱ�ݵȼ�,null,to_CHAR(�������),ȱ�ݵȼ�) from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[1]) as ����," & _
        "           (select �ɷ��޸� from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[1]) as �ɷ��޸�," & _
        "           (select ��ע from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[1]) as ��ע " & _
        "   from �������ֱ�׼��ͼ A " & _
        "   where A.����='��' and A.����ID=(select B.����ID from �������ֽ�� B where B.ID=[1]) " & _
        "   order by A.�ϼ�ID,A.ID "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng���ID)
    If Not rs.EOF Then
        For i = 1 To fg���_S.Rows - 1
            rs.MoveFirst
            rs.Find "ID=" & Val(fg���_S.Cell(flexcpText, i, 6))
            If Not rs.EOF Then
                If Not IsNull(rs("����")) Then
                    Select Case rs("����")
                    Case "��", "��", "��"
                        fg���_S.Cell(flexcpText, i, 4) = rs("����").Value + "��"
                    Case "��"
                        fg���_S.Cell(flexcpText, i, 4) = "������"
                    Case Else
                        fg���_S.Cell(flexcpText, i, 4) = IIf(Abs(NVL(rs("����").Value, 0)) < 1, Format(Abs(NVL(rs("����").Value, 0)), "0.0"), Abs(NVL(rs("����").Value, 0)))
                    End Select
                    If intSign = -1 Then
                        fg���_S.Cell(flexcpForeColor, i, 4) = RGB(255, 0, 0)
                    Else
                        fg���_S.Cell(flexcpForeColor, i, 4) = RGB(0, 0, 255)
                    End If
                End If
                If Not IsNull(rs("�ɷ��޸�")) Then
                    If rs("�ɷ��޸�") = 1 Then
                        fg���_S.Cell(flexcpText, i, 5) = "��"
                    End If
                End If
                fg���_S.Cell(flexcpText, i, 9) = NVL(rs!��ע)
            End If
        Next
    End If
    fg���_S.Redraw = flexRDBuffered
    If fg���_S.Rows <= 1 Then
        '������
        fg���_S.WallPaper = imgBG_fg(0).Picture
    ElseIf Trim(fg����_S.Cell(flexcpText, fg����_S.Row, 20)) <> "" Then
        '�����
        fg���_S.WallPaper = imgBG_fg(1).Picture
    Else
        'δ���
        fg���_S.WallPaper = LoadPicture("")
    End If
    Fill���ֽ�� = True
    Call SetMenu
    Call fg����_S_SelChange
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Fill���ֽ�� = False
End Function

'==============================================================================
'=���ܣ����ٲ��Ҳ���
'==============================================================================
Private Sub ���Ҳ���(��Χ As String, strID As String)
    Dim lngBRID             As Long
    Dim lngZYID             As Long
    Dim strSQL              As String
    Dim i                   As Long
    Dim rs                  As ADODB.Recordset
    Dim blnFinded           As Boolean
    Dim lngCurRowTMP        As Long
    On Error GoTo ErrH
    If ��Χ = "1-���￨��" Then
        strSQL = _
            "Select A.����ID,B.��ҳID " & _
            " From ������Ϣ A,������ҳ B " & _
            " Where A.����ID=B.����ID " & _
            " And Nvl(B.��ҳID,0)<>0 " & _
            " And A.���￨��=[1]"
    ElseIf ��Χ = "2-����ID" Then '����ID
        strSQL = _
            "Select A.����ID,B.��ҳID " & _
            " From ������Ϣ A,������ҳ B " & _
            " Where A.����ID=B.����ID " & _
            " And Nvl(B.��ҳID,0)<>0 " & _
            " And A.����ID=[1]"
    ElseIf ��Χ = "3-סԺ��" Then 'סԺ��(������Ժ)
        strSQL = _
            " Select A.����ID,B.��ҳID " & _
            " From ������Ϣ A,������ҳ B " & _
            " Where A.����ID=B.����ID  And Nvl(B.��ҳID,0)<>0 And B.סԺ��=[1]"
    ElseIf ��Χ = "4-�����" Then '�����(ҽ������)
        strSQL = _
            " Select A.����ID,B.��ҳID " & _
            " From ������Ϣ A,������ҳ B " & _
            " Where A.����ID=B.����ID   And Nvl(B.��ҳID,0)<>0 And A.�����=[1]"
    Else '��������
        strSQL = _
            " Select A.����ID,B.��ҳID " & _
            " From ������Ϣ A,������ҳ B " & _
            " Where A.����ID=B.����ID  And Nvl(B.��ҳID,0)<>0 And Upper(A.����)=[1]"
    End If
    gstrSQL = strSQL
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(UCase(strID)))
    If Not rs.EOF Then
        lngBRID = rs("����ID")
        If lngBRID <= 0 Then Exit Sub
        With fg����_S
            lngCurRowTMP = .Row
            For i = lngCurRowTMP + 1 To .Rows - 1
                If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                    .Row = i
                    .ShowCell i, 2
                    blnFinded = True
                    Exit For
                End If
            Next
            If blnFinded = False Then '�����ǰ������û��ƥ�����ӵ�һ�п�ʼ���²�ѯ��
                For i = 1 To lngCurRowTMP
                    If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                        .Row = i
                        .ShowCell i, 2
                        blnFinded = True
                        Exit For
                    End If
                Next
            End If
        End With
    End If
    rs.Close
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���˫��ʱ��������
'==============================================================================
Private Sub tvw����_DblClick()
    Dim i       As Integer
    On Error GoTo ErrH
    If tvw����.SelectedItem Is Nothing Then Exit Sub
    With tvw����.SelectedItem
        txt����.Tag = Mid(.Key, 2)
        txt����.Text = Mid(.Text, InStr(.Text, "��") + 1)
        tvw����.Visible = False
        '��Ժ���ҿ���ѡ��
        If txt����.Text = "���п���" Then
            rsM.Filter = ""
        Else
            If mblnSetDept Then
                rsM.Filter = ""
            Else
                rsM.Filter = "��Ժ����='" & Mid(tvw����.SelectedItem.Text, InStr(tvw����.SelectedItem.Text, "��") + 1) & "'"
            End If
        End If
        Call Fill����
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����������
'==============================================================================
Private Sub tvw����_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        '��Ժ���ҿ���ѡ��
        txt����.Tag = Mid(tvw����.SelectedItem.Key, 2)
        txt����.Text = Mid(tvw����.SelectedItem.Text, InStr(tvw����.SelectedItem.Text, "��") + 1)
        If txt����.Text = "���п���" Then
            rsM.Filter = ""
        Else
            rsM.Filter = "��Ժ����='" & Mid(tvw����.SelectedItem.Text, InStr(tvw����.SelectedItem.Text, "��") + 1) & "'"
        End If
        Call Fill����
    ElseIf KeyAscii = 27 Or KeyAscii = vbKeySpace Then
        tvw����.Visible = False
        txt����.SetFocus
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���ʧȥ����ʱ����
'==============================================================================
Private Sub tvw����_LostFocus()
    On Error GoTo ErrH
    tvw����.Visible = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt����_Change()
    Dim lngStart        As Long
    Dim lngLength       As Long
On Error GoTo ErrH

    '�ؼ�¼������ַ��ȵ�
    lngLength = Len(txt����.Text)
    lngStart = txt����.SelStart
    txt����.Text = ConvertString(txt����.Text)
    If lngStart - (lngLength - Len(txt����.Text)) >= 0 Then txt����.SelStart = lngStart - (lngLength - Len(txt����.Text))
 
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��ı��򰴼�����
'==============================================================================
Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim zlInputNot      As String
    On Error GoTo ErrH
    '�����������ַ�
    zlInputNot = "'|"
    If Len(zlInputNot) > 0 Then
        If InStr(1, zlInputNot, Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
    If txt����.Locked Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        Call DeptSelect
    ElseIf KeyAscii = 27 Then
        tvw����.Visible = False
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��ı��õ�����ѡ������
'==============================================================================
Private Sub txt����_GotFocus()
    On Error GoTo ErrH
    Call zlControl.TxtSelAll(txt����)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���λ���ݻس�ˢ��
'==============================================================================
Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo ErrH
    
    lngRow = 0
    If txt����.Locked Then Exit Sub
    If mstrFindKey = "��������" Then mstrFindKey = "����"
    If fg����_S.ColIndex(mstrFindKey) = -1 Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        For lngLoop = fg����_S.Row + 1 To fg����_S.Rows - 1
            If InStr(UCase(fg����_S.TextMatrix(lngLoop, fg����_S.ColIndex(mstrFindKey))), UCase(txt����.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '��ȡС�ڵ�ǰ�еļ�¼����
        If lngRow = 0 Then
            For lngLoop = 0 To fg����_S.Row
                If InStr(UCase(fg����_S.TextMatrix(lngLoop, fg����_S.ColIndex(mstrFindKey))), UCase(txt����.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fg����_S.Rows > 1 And lngRow >= 1 Then fg����_S.Row = lngRow
        Call LocationObj(txt����)
    End If
    If mstrFindKey = "����" Then mstrFindKey = "��������"
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȡ�ø��Ի�����
'==============================================================================
Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ��˵���ť����
'==============================================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    
    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_View_Refresh           'ˢ������
            Call mDataLoad
        Case conMenu_File_Preview           'Ԥ��
            subPrint 2
        Case conMenu_File_Print             '��ӡ
            subPrint 1
        Case conMenu_File_Excel             '�����&Excel
            subPrint 3
        Case conMenu_File_BatPrint          'ȫ����ӡ
            Call RecordAllPrint
        Case conMenu_Edit_NewParent         '��������
            Call RecordRating
        Case conMenu_Edit_ModifyParent      '�޸Ľ��
            Call RecordEdit
        Case conMenu_Edit_Insert            '��������
            Call RecordReturn
        Case conMenu_Edit_DeleteParent      'ɾ�����
            Call RecordDel
        Case conMenu_Manage_ReportView      '������ҳ
            Call RecordLook
        Case conMenu_Manage_Audit           'ͨ�����
            Call RecordAudit
        Case conMenu_Edit_Leave_UndoPost    'ȡ�����
            Call RecordUnAudit
        Case conMenu_Edit_Select            'ȫ��ѡ��
            Call RecordSelect
        Case conMenu_Edit_DeSelect          'ȡ��ѡ��
            Call RecordUnSelect
        Case conMenu_Manage_UnAudit         '����ѡ��
            Call RecordSelectOther
        Case 10004                          '����ѡ��
            Call DeptSelect
        Case conMenu_View_Find              '���˲�ѯ
            Call RecordFind
        Case conMenu_View_Forward           '��һ��
            With fg����_S
                If .Row > 1 Then
                    .Row = .Row - 1
                    .ShowCell .Row, .Col
                End If
            End With
        Case conMenu_View_Backward          '��һ��
            With fg����_S
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .ShowCell .Row, .Col
                End If
            End With
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txt����
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '��ҵ���޹صĹ��ܣ������Ĺ���
                Call CommandBarExecutePublic(Control, Me, fg����_S, "��������")
            End If
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��˵�Ȩ�޿���
'==============================================================================
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_BatPrint 'Ԥ��,��ӡ,�����Excel,ȫ����ӡ
            Control.Enabled = ((fg����_S.Rows > 1) And IsPrivs(mstrPrivs, "����"))
        Case conMenu_Edit_Select, conMenu_Edit_DeSelect, conMenu_Manage_UnAudit
            Control.Enabled = ((fg����_S.Rows > 1))
        Case conMenu_Manage_ReportView
            Control.Enabled = ((fg����_S.Rows > 1))
        Case conMenu_View_Refresh
            Control.Enabled = IsPrivs(mstrPrivs, "����")
        Case conMenu_Edit_NewParent
            Control.Visible = (InStr(mstrPrivs, "����") > 0)
            Control.Enabled = (Not mRecordRating) And (fg����_S.Rows > 1)
        Case conMenu_Edit_ModifyParent      '�޸�[������Ȩ���ң����޸��������ֻ��Լ��ļ�¼����δ���]
            Control.Visible = (InStr(mstrPrivs, "����") > 0)
            Control.Enabled = (InStr(mstrPrivs, "�޸���������") > 0 Or mRecordMyAudit) And (fg����_S.Rows > 1) And mRecordRating And Not mRecordAudit
        Case conMenu_Edit_DeleteParent
            Control.Visible = (InStr(mstrPrivs, "����") > 0)
            Control.Enabled = (InStr(mstrPrivs, "�޸���������") > 0 Or mRecordMyAudit) And (fg����_S.Rows > 1) And mRecordRating And Not mRecordAudit
        Case conMenu_Edit_Insert
            Control.Visible = (InStr(mstrPrivs, "����") > 0)
            Control.Enabled = (InStr(mstrPrivs, "�޸���������") > 0 Or mRecordMyAudit) And (fg����_S.Rows > 1) And mRecordRating And Not mRecordAudit And mRecordReturn
        Case conMenu_Manage_Audit   '���
            Control.Visible = (InStr(mstrPrivs, "���") > 0)
            Control.Enabled = (InStr(mstrPrivs, "�޸���������") > 0 Or mRecordMyAudit) And (fg����_S.Rows > 1) And mRecordRating And Not mRecordAudit
        Case conMenu_Edit_Leave_UndoPost
            Control.Visible = (InStr(mstrPrivs, "���") > 0)
            Control.Enabled = (InStr(mstrPrivs, "�޸���������") > 0 Or mRecordMyAudit) And (fg����_S.Rows > 1) And mRecordRating And mRecordAudit
        Case 10004              '�����п�������ѡ��
            Control.Visible = IsPrivs(mstrPrivs, "���п���")
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_BatPrint  'Ԥ��,��ӡ,�����Excel,ȫ����ӡ
            Control.Enabled = ((fg����_S.Rows > 1) And IsPrivs(mstrPrivs, "��ӡ���ֽ����"))
        Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
            If InStr(Control.Caption, mstrFindKey) > 0 Then
                Control.Checked = True
            Else
                Control.Checked = False
            End If
        Case conMenu_View_Forward
            Control.Enabled = (Control.Visible And fg����_S.Row > 1)
        Case conMenu_View_Backward
                Control.Enabled = (Control.Visible And fg����_S.Row + 1 < fg����_S.Rows)
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Function Get��������(ByVal lng��ԱId As Long, ByVal lngMode As Long) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrH
    ' lngMode =0 ��ʾ��ʽ lngMode =1 ���ڲ�ѯ��ʽ
    
    strSQL = "SELECT  distinct C.���� AS ����" & vbNewLine & _
                "      FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D" & vbNewLine & _
                "      WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                "      AND A.id =[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԱId)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            If lngMode = 0 Then
                If Len(strTmp) = 0 Then
                    strTmp = NVL(rsTemp!����)
                Else
                    strTmp = strTmp & "," & NVL(rsTemp!����)
                End If
            Else
                If Len(strTmp) = 0 Then
                    strTmp = "'" & NVL(rsTemp!����) & "'"
                Else
                    strTmp = strTmp & ",'" & NVL(rsTemp!����) & "'"
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        
        Get�������� = strTmp
    Else
        Get�������� = UserInfo.��������
    End If
    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
End Function
