VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathTable 
   BorderStyle     =   0  'None
   Caption         =   "�ٴ�·����"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15015
      Begin VB.Frame fraPath 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   380
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   10815
         Begin VB.ComboBox cboPath 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   30
            Width           =   2415
         End
         Begin VB.Label lblInPep 
            BackColor       =   &H8000000E&
            Caption         =   "�����ˣ����Ʊ�"
            Height          =   255
            Left            =   3360
            TabIndex        =   18
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblInDate 
            BackColor       =   &H8000000E&
            Caption         =   "����ʱ�䣺2011-01-01 24:24"
            Height          =   255
            Left            =   5160
            TabIndex        =   17
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblOutDate 
            BackColor       =   &H8000000E&
            Caption         =   "����ʱ�䣺2011-01-01 01:01"
            Height          =   255
            Left            =   7620
            TabIndex        =   16
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblInDiag 
            BackColor       =   &H8000000E&
            Caption         =   "������ϣ�"
            Height          =   255
            Left            =   10080
            TabIndex        =   15
            Top             =   120
            Width           =   4995
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "·������"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraSendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   380
         Left            =   11880
         TabIndex        =   7
         Top             =   0
         Width           =   3015
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ȫ��"
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   10
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ҽ��"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   9
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʿ"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   8
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblSendNote 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����ߣ�"
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin zlCISPath.UCAdviceList UCAdvice 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2566
   End
   Begin VB.Frame fraline 
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   5520
      Width           =   8175
   End
   Begin MSComctlLib.ImageList imgCharacter 
      Left            =   8280
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":0000
            Key             =   "�Ѿ�ִ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":059A
            Key             =   "��δִ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":0B34
            Key             =   "ȡ��ִ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":10CE
            Key             =   "����ִ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":1668
            Key             =   "��ǰִ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":1C02
            Key             =   "�Ӻ�ִ��"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8400
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   225
      Begin VB.Image imgMore 
         Height          =   225
         Left            =   0
         Picture         =   "frmPathTable.frx":219C
         Top             =   0
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList imgFlow 
      Left            =   8280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":259D
            Key             =   "node"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":26E4
            Key             =   "currnode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":2833
            Key             =   "multnode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":29B5
            Key             =   "currmultnode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":2B7B
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":2FFE
            Key             =   "arrowlate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":3479
            Key             =   "arrow_Branch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":3899
            Key             =   "arrowlate_Branch"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "˫���鿴·����Ŀ����"
      Top             =   2280
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTable.frx":3CBD
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsFlow 
      Height          =   1920
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "˫���鿴·���׶ζ���"
      Top             =   360
      Width           =   8175
      _cx             =   14420
      _cy             =   3387
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483634
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   1800
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTable.frx":3DFA
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   101
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ѵ�ӡ���߰�·����"
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1725
         WordWrap        =   -1  'True
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPathPrint 
      Height          =   3105
      Index           =   0
      Left            =   0
      TabIndex        =   19
      ToolTipText     =   "˫���鿴·����Ŀ����"
      Top             =   -99999
      Visible         =   0   'False
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTable.frx":3E67
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   8880
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnMoved As Long
Private mint���� As Integer                 '0-ҽ��վ����,1-��ʿվ����
Private mbln����ִ�л��� As Boolean         '�Ƿ�����·��ִ�л���
Private mstrִ�г��� As String                 '"10"-ҽ��,"01"-��ʿ��"11"��ҽ����ʿ
Private mbln���ò����� As Boolean         '����ǰһ�첻���������ɽ����·����Ŀ F-�������,T-�����������
Private mbln������ǰ���� As Boolean         '������ǰ���������·����Ŀ
Private mbytPrintWay As Byte                '0-����ӡ;1-�����ӡ
Private mblnInsideTools As Boolean          '�ڲ�������ģʽ
Private mfrmParent As Object, mcbsMain As Object
Attribute mcbsMain.VB_VarUserMemId = 1073938436
Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣƽ̨����
Private mobjPublicPACS As Object
Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)    'Ҫ��鿴����
Public Event Activate()    '���Ѽ���ʱ
Public Event RequestRefresh(ByVal lngPathState As Long)  'Ҫ��������ˢ��
Attribute RequestRefresh.VB_UserMemId = 3
Public Event StatusTextUpdate(ByVal Text As String)    'Ҫ�����������״̬������
Attribute StatusTextUpdate.VB_UserMemId = 4
Private Const C_Exe = "��"  '��
Private Const C_UnExe = "��"
Private Enum EFixedRow
    R0�׶��� = 0
    R1���� = 1
    R2���� = 2
End Enum
'�������±�ֵ
Private Enum CONST_IX_SENDOR
    IX_ALL = 0
    IX_ҽ�� = 1
    IX_��ʿ = 2
End Enum

Private mblnUnChange As Boolean    '�����õ�Ԫ��仯�¼���ˢ�µ�Ԫ������
Attribute mblnUnChange.VB_VarUserMemId = 1073938439

Private mPP As TYPE_PATH_Pati
Attribute mPP.VB_VarUserMemId = 1073938440
Private mPati As TYPE_Pati
Attribute mPati.VB_VarUserMemId = 1073938441
Private mcolReason As Collection
Attribute mcolReason.VB_VarUserMemId = 1073938442
Private mblnInOverScope As Boolean    '���˵�ǰִ�������Ƿ��ڱ�׼סԺ�շ�Χ���������·����
Attribute mblnInOverScope.VB_VarUserMemId = 1073938443
Private mlng����״̬ As Long    '�������״̬
Attribute mlng����״̬.VB_VarUserMemId = 1073938444
Private mlngҽ������ID As Long
Private mlngӤ������ID As Long
Private mlngӤ������ID As Long
Private mlngState As Long      '���˱䶯״̬  =5Ϊת������
Attribute mlngState.VB_VarUserMemId = 1073938445
Private mrsPlugInBar As ADODB.Recordset '�˵���ʽ
Private mlngPlugInID As Long '�Զ�ִ�еĲ������ID

'ˢ������ʱ����Ĳ���״̬
Public Enum TYPE_PATI_State
    ps��Ժ = 0
    psԤ�� = 1
    ps��Ժ = 2
    ps���� = 3          'ҽ��վ:�����ﲡ��(��Ժ)
    ps���� = 4          'ҽ��վ:�ѻ��ﲡ��
    ps���ת�� = 5      'ҽ��վ:���ת�ƻ�ת�����Ĳ���(��Ժ)
    ps��ת�� = 6        'ҽ��վ:��ƴ���ס��ת��������������
End Enum

Private mlngFontSize As Long
Attribute mlngFontSize.VB_VarUserMemId = 1073938446
Private mlngPathCount As Long   '����סԺ��·����
Attribute mlngPathCount.VB_VarUserMemId = 1073938447

Private Const CON_SmallFontSize As Long = 9     'С����
Private Const CON_BigFontSize As Long = 12     '������
Private Const CON_PathOutItemColor As Long = &HC0FFFF        '·������Ŀ��ǳ��ɫ
Private Const CON_PathOutItemColorBlue As Long = &HFAEADA    '�ݴ�·������Ŀ,ǳ��ɫ��ʶ

Private Sub SetUnImport()
'���ܣ�����δ����ʱ��״̬����Ϣ
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .ForeColorSel = vbBlack
        .TextMatrix(0, 0) = "  �ò���δ�����ٴ�·����"
    End With
    Call ClearPathItem
End Sub

Private Sub SetImportFalse()
'���ܣ����õ����˵����ٴ�·��ʧ��ʱ��״̬����Ϣ
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .TextMatrix(0, 0) = "  �ò��˲�����·������������" & vbCrLf & "  ԭ��" & mPP.δ����ԭ��
        .AutoSize 0
        .ForeColorSel = &HC0&
        If .Visible And .Enabled Then .SetFocus
    End With
    Call ClearPathItem
End Sub

Private Sub ClearPathItem(Optional blnImported As Boolean)
'���ܣ�������û�п��õ��ٴ�·��ʱ���·������Ŀ
    With vsPath
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 0
        .Cols = 0
        
        If blnImported Then
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = vbCrLf & "  �ò��˻�û������·����Ŀ��"
            .Select 0, 0
            .CellAlignment = flexAlignLeftTop
        End If
        fraSendor.Tag = "����"
    End With
End Sub


Private Sub cboPath_Click()
    If cboPath.ListIndex >= 0 Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬, , , , Val(cboPath.ItemData(cboPath.ListIndex)))
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Call zlPopupCommandBars(CommandBar)
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If fraSendor.Visible Or fraPath.Visible Then
        fraTop.Top = lngTop
        fraTop.Left = lngLeft
        fraTop.Width = lngRight - lngLeft
        fraSendor.Left = fraTop.Width - fraSendor.Width
        fraPath.Width = fraSendor.Left
        lngTop = fraTop.Top + fraTop.Height
    End If
    vsFlow.Left = lngLeft
    vsFlow.Top = lngTop
    vsFlow.Width = lngRight - lngLeft
    vsFlow.Height = lngBottom - lngTop
        
    If vsPath.FixedRows = 0 And vsPath.Rows = 0 Then  'û�е���·��ʱ
        vsFlow.Height = vsFlow.Height + vsPath.Height
        vsPath.Visible = False
        UCAdvice.Visible = False
        fraline.Visible = False
    Else
        If vsPath.Visible = False Then vsPath.Visible = True
        If UCAdvice.Visible = False Then UCAdvice.Visible = True
        If fraline.Visible = False Then fraline.Visible = True
        
        If Grid.HScrollVisible(vsFlow) = False Then
            vsFlow.Height = 1140 + IIf(mPP.�ϲ�·������ > 2, (mPP.�ϲ�·������ - 2) * 180, 0)
        Else
            vsFlow.Height = 1440 + IIf(mPP.�ϲ�·������ > 2, (mPP.�ϲ�·������ - 2) * 180, 0)
        End If
        With vsPath
            .Top = lngTop + vsFlow.Height
            .Width = lngRight - lngLeft
            If lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30 > 0 Then
                .Height = lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30
            Else
                .Height = lngBottom - lngTop - vsFlow.Height
            End If
        
            If .FixedRows = 0 And .Rows = 1 Then             'û��������Ŀ
                .ColWidth(0) = .Width - 30
                .RowHeight(0) = .Height
            End If
            fraline.Top = .Top + .Height
            fraline.Width = .Width
            
            UCAdvice.Top = fraline.Top + fraline.Height
            UCAdvice.Width = .Width
        End With
    End If
    
    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub fraline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next
        If vsPath.Height + Y < 1000 Or vsPath.Height - Y < 500 Then Exit Sub
        If UCAdvice.Height + Y < 250 Or UCAdvice.Height - Y < 500 Then Exit Sub
                
        If fraMore.Visible Then fraMore.Visible = False
        
        fraline.Top = fraline.Top + Y
        vsPath.Height = vsPath.Height + Y
        UCAdvice.Top = UCAdvice.Top + Y
        UCAdvice.Height = UCAdvice.Height - Y
    End If
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub Form_Load()
    Set mrsPlugInBar = Nothing
    Call RestoreWinState(Me, App.ProductName)
    Call InitCbsSubBar
    '��ʼ��LIS����
    Call InitObjLis(P�ٴ�·��Ӧ��)
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
    lblPrinted.Top = vsFlow.RowPos(0)
    lblPrinted.Left = vsFlow.ColPos(0)
End Sub

Private Sub LoadPathFlow()
'���ܣ����ݲ��˵����·�������·��������Ϣ������
    Dim strSql As String, i As Long, j As Long, lngCurCol As Long
    Dim rsTmp As ADODB.Recordset, lngDayMin As Long, lngDayMax As Long
    Dim lng�������� As Long
    Dim lng��� As Long
    Dim str��׼סԺ�� As String
    Dim rsBranch As Recordset
    Dim rsMerge As Recordset

    With vsFlow
        .Clear
        .Rows = 1: .Cols = 1
        .ForeColorSel = vbBlack
        mblnInOverScope = False
        On Error GoTo errH
        '�Ѵ�ӡ���߰�·����
        strSql = "Select 1 From ���Ӳ�����ӡ Where �ļ�id = [1] And ���� = 12"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        If rsTmp.RecordCount > 0 Then
            lblPrinted.Caption = "�Ѵ�ӡ���߰�·����"
            lblPrinted.ForeColor = &HC0&
            lblPrinted.Visible = True
        Else
            lblPrinted.Caption = ""
            lblPrinted.Visible = False
        End If

        If mPP.��ǰ�׶η�֧ID <> 0 Then
            strSql = "Select NVL(c.���,b.���) as ��� From �ٴ�·����֧ A,�ٴ�·���׶� B,�ٴ�·���׶� C Where a.ǰһ�׶�ID=b.ID And b.��ID=c.id(+) And a.ID=[1]"
            '�������ǰ��֧��ǰһ�׶ε����
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.��ǰ�׶η�֧ID)
            lng��� = Val(rsTmp!��� & "")
        End If

        If mPP.��ǰ�׶η�֧ID = 0 Then
            strSql = "Select a.ID,a.���� �׶���, Decode(a.��������, Null, 0, 1) ����,b.����,b.���� ·����,b.���°汾,c.��׼סԺ�� ,a.��֧id" & _
                     " From �ٴ�·���׶� a,�ٴ�·��Ŀ¼ b,�ٴ�·���汾 c " & _
                     " Where a.·��id = [1] And a.�汾�� = [2] And a.·��id=b.id And a.��ID is null And b.id = c.·��id And a.�汾�� = c.�汾�� " & _
                     " And a.��֧ID is Null" & _
                     " Order by a.���"
        Else
            strSql = "Select a.ID,a.���� �׶���, Decode(a.��������, Null, 0, 1) ����,b.����,b.���� ·����,b.���°汾,c.��׼סԺ�� ,a.��֧id" & _
                     " From �ٴ�·���׶� a,�ٴ�·��Ŀ¼ b,�ٴ�·���汾 c,�ٴ�·����֧ D,�ٴ�·���׶� E,�ٴ�·���׶� F,�ٴ�·���׶� G " & _
                     " Where a.·��id = [1] And a.�汾�� = [2] And a.·��id=b.id And a.��ID is null And b.id = c.·��id And a.�汾�� = c.�汾�� " & _
                     " And a.��֧ID=d.ID(+) And a.��ID=e.id(+) And d.ǰһ�׶�ID=f.id(+) And f.��ID=g.id(+) And (a.��֧ID=[3] Or NVL(e.���,a.���)<=[4] and a.��֧ID is null )" & _
                     " Order by Decode(a.��֧ID,Null,NVL(e.���,a.���),NVL(e.���,a.���)+NVL(g.���,f.���))"
        End If

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID, mPP.�汾��, mPP.��ǰ�׶η�֧ID, lng���)
        If mPP.��ǰ�׶η�֧ID = 0 Then
            str��׼סԺ�� = rsTmp!��׼סԺ�� & ""
        Else
            strSql = "Select ��׼סԺ�� From �ٴ�·����֧ Where ID=[1]"
            Set rsBranch = zlDatabase.OpenSQLRecord(strSql, "��ȡ��֧��׼סԺ��", mPP.��ǰ�׶η�֧ID)
            str��׼סԺ�� = rsBranch!��׼סԺ�� & ""
        End If
        If rsTmp.RecordCount > 0 Then
            .Rows = 1
            .Cols = rsTmp.RecordCount * 2    '��һ��Ϊ·��������ͷΪ�׶���-1
            .Select 0, 0
            .RowHeight(0) = 1100 + IIf(mPP.�ϲ�·������ > 2, (mPP.�ϲ�·������ - 2) * 180, 0)

            '��һ����ʾ·������
            .ColWidth(0) = 2800

            If mPP.����·��״̬ > 0 Then
                strSql = "Select b.����,a.����ʱ�� From ���˺ϲ�·�� A,�ٴ�·��Ŀ¼ B Where a.·��ID=b.ID And a.��Ҫ·����¼ID=[1]"
                Set rsMerge = zlDatabase.OpenSQLRecord(strSql, "�ϲ�·��", mPP.����·��ID)
                .TextMatrix(0, 0) = rsTmp!·���� & ""
                Do While Not rsMerge.EOF
                    .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "(�ϲ�)" & rsMerge!���� & IIf(IsNull(rsMerge!����ʱ��), "", "(���)")
                    rsMerge.MoveNext
                Loop


                If mPP.����·��״̬ = 3 Then
                    .Cell(flexcpForeColor, 0, 0) = vbRed
                End If
            Else
                .TextMatrix(0, 0) = rsTmp!·����
            End If
            If mPP.��ǰ���� > 0 And mPP.����·��״̬ = 1 Then
                If InStr(str��׼סԺ��, "-") > 0 Then
                    j = Split(str��׼סԺ��, "-")(1)
                    lngDayMin = Val(Split(str��׼סԺ��, "-")(0))
                    lngDayMax = j
                Else
                    j = Val(str��׼סԺ��)   'С�ڵ���n������
                    lngDayMin = 1
                    lngDayMax = j
                End If
                lng�������� = GetMustDay(mPP.����·��ID, mPP.��ǰ����)

                i = Format(lng�������� / j * 100, "0")
                If i = 100 And lng�������� <> j Then i = 99
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "���ȣ�" & i & "%"


                If lng�������� > lngDayMax Then
                    mblnInOverScope = True
                Else
                    mblnInOverScope = Between(lng��������, lngDayMin, lngDayMax)
                End If
            End If
            If mPP.����·��״̬ > 0 Then
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "״̬��" & IIf(mPP.����·��״̬ = 1, "ִ����", IIf(mPP.����·��״̬ = 2, "���", "�����˳�"))
            End If
            .Cell(flexcpTextStyle, 0, 0) = 3

            For i = 1 To .Cols Step 2
                .TextMatrix(0, i) = " " & rsTmp!�׶��� & " "    '���ñ߾�
                .ColAlignment(i) = flexAlignCenterCenter

                .ColWidth(i) = 1750
                .Col = i
                .PicturesOver = True
                .CellPictureAlignment = flexPicAlignLeftCenter
                If mPP.��ǰ�׶�ID = rsTmp!ID Or mPP.�׶θ�ID = rsTmp!ID Or (mPP.��ǰ�׶�ID = 0 And i = 1 And mPP.����·��״̬ = 1) Then
                    lngCurCol = i
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!���� = 1, "currmultnode", "currnode")).Picture
                    Call .ShowCell(0, i)
                Else
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!���� = 1, "multnode", "node")).Picture
                End If
                .ColData(i) = Val(rsTmp!ID)

                rsTmp.MoveNext

                '��ͷ
                If i < .Cols - 1 Then
                    .ColWidth(i + 1) = 550
                    .Col = i + 1
                    .CellPictureAlignment = flexPicAlignCenterCenter
                    .CellPicture = imgFlow.ListImages(IIf(i + 1 > lngCurCol And lngCurCol <> 0 Or mPP.����·��״̬ > 1, "arrowlate", "arrow") & IIf(rsTmp!��֧ID & "" <> "", "_Branch", "")).Picture
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Getִ�н������ͼ��(ByVal lngִ�н������ As Long) As Long
'���ܣ�����ִ�н�����ʷ��ض�Ӧ��ͼ�����
'1-�Ѿ�ִ�У�2-��δִ�У�3-ȡ��ִ�У�4-����ִ�У�5-��ǰִ�У�6-�Ӻ�ִ��
    Dim lngIdx As Long
    
    Select Case lngִ�н������
        Case 1
            lngIdx = imgCharacter.ListImages("�Ѿ�ִ��").Index
        Case 2
            lngIdx = imgCharacter.ListImages("��δִ��").Index
        Case 3
            lngIdx = imgCharacter.ListImages("ȡ��ִ��").Index
        Case 4
            lngIdx = imgCharacter.ListImages("����ִ��").Index
        Case 5
            lngIdx = imgCharacter.ListImages("��ǰִ��").Index
        Case 6
            lngIdx = imgCharacter.ListImages("�Ӻ�ִ��").Index
    End Select
    Getִ�н������ͼ�� = lngIdx
End Function

Private Sub LoadPathItem()
'���ܣ����ز��������ɵ�·����Ŀ
    Dim strSql As String, strOldType As String, str������� As String
    Dim lngRow As Long, lngCol As Long, i As Long, j As Long, arrtmp As Variant, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim CPos As New Collection  'ÿ���������ʼ��
    Dim lngPreRow As Long, lngPreCol As Long, lngPrePathID As Long, lngDayRow As Long
    Dim lngBranchID As Long  '��֧ID
    Dim rsSort As Recordset
    Dim str����ԭ�� As String
    
    With vsPath
        lngPreRow = -1
        lngPreCol = -1
        If .Row >= .FixedRows Then lngPreRow = .Row
        If .Col >= .FixedCols Then lngPreCol = .Col

        '1)���ಿ��
        .Redraw = flexRDNone
        mblnUnChange = True
        .Clear
        .Rows = 3: .FixedRows = 3
        .Cols = 1: .FixedCols = 1
        mblnUnChange = False
        .MergeCol(0) = True
        .MergeRow(0) = True

        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 0, .FixedRows - 1, 0) = "ʱ��׶�"
        On Error GoTo errH

        '�������·����ת����ͬ�׶ε���Ŀ�����ǲ�ͬ��·�����
        strSql = _
        "Select ����, Max(����) As ����,100 as ���" & vbNewLine & _
                 "From (Select Count(a.Id) As ����, a.����, a.�׶�id, a.����" & vbNewLine & _
                 "       From ����·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1]" & vbNewLine & _
                 "       Group By a.����, a.����, a.�׶�id)" & vbNewLine & _
                 "Group By ����"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp)

        '��ȡ���(������·�����ȡ�Ȼ���Ƿ�֧·��������Ǻϲ�·��������н�������Ҳ���������)
        strSql = _
        "Select ����, ���" & vbNewLine & _
                 "From (Select ����, ���," & vbNewLine & _
                 "              Row_Number() Over(Partition By ���� Order By Decode(�ϲ�·����¼id, Null, Decode(��֧id, Null, 1, 2), Decode(��֧id, Null, 3, 4))) As Top" & vbNewLine & _
                 "       From (Select a.���, a.���� As ����, c.��֧id, b.�ϲ�·����¼id" & vbNewLine & _
                 "              From �ٴ�·������ A, ����·��ִ�� B, �ٴ�·����Ŀ C" & vbNewLine & _
                 "              Where a.���� = c.���� And b.·����¼id = [1] And b.��Ŀid = c.Id And c.·��id = a.·��id And c.�汾�� = a.�汾�� And" & vbNewLine & _
                 "                    Nvl(c.��֧id, 0) = Nvl(a.��֧id, 0)" & vbNewLine & _
                 "              Union" & vbNewLine & _
                 "              Select a.���, a.���� As ����, c.��֧id, b.�ϲ�·����¼id" & vbNewLine & _
                 "              From �ٴ�·������ A, ����·��ִ�� B, �ٴ�·���׶� C" & vbNewLine & _
                 "              Where a.���� = b.���� And b.�׶�id+0 = c.Id And b.·����¼id = [1] And b.��Ŀid Is Null And a.·��id = c.·��id And" & vbNewLine & _
                 "                    a.�汾�� = c.�汾�� And Nvl(c.��֧id, 0) = Nvl(a.��֧id, 0)))" & vbNewLine & _
                 "Where Top = 1" & vbNewLine & _
                 "Order By ���"

        Set rsSort = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        '����
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                rsSort.Filter = "����='" & rsTmp!���� & "'"
                If rsSort.RecordCount > 0 Then
                    rsTmp!��� = Val(rsSort!��� & "")
                    rsTmp.Update
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Sort = "���"
            rsTmp.MoveFirst
        End If
        For i = 1 To rsTmp.RecordCount
            CPos.Add .Rows, "T" & rsTmp!����
            .Rows = .Rows + rsTmp!����
            For j = 1 To rsTmp!����
                .TextMatrix(.Rows - j, .FixedCols - 1) = rsTmp!����
            Next
            rsTmp.MoveNext
        Next


        '2)ʱ��׶β���
        '�׶�����ʱ�� NVL(c.���,b.���) ��Ϊ�˴����÷�֧������������⣬ȡֵb.��� ����Ϊ��������Ҫ��ʾ�ǵڼ�����֧����ȡ��֧·�������ʱ��ȡ����һ�׶ε���ż��Ϸ�֧·������ţ�
        If mPP.��ǰ�׶η�֧ID = 0 Then
            strSql = _
            "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, To_Char(a.����, 'day') ����, b.���� As �׶���, b.���, b.˵��, b.��id,Decode(g.·��id,b.·��id,1,0) as ����" & vbNewLine & _
                     "From (Select a.�׶�id, a.����, a.����,a.·����¼id" & vbNewLine & _
                     "       From ����·��ִ�� A" & vbNewLine & _
                     "       Where a.·����¼id = [1]" & vbNewLine & _
                     "       Group By a.�׶�id, a.����, a.����,a.·����¼id) A, �ٴ�·���׶� B,�ٴ�·���׶� C,�����ٴ�·�� G" & vbNewLine & _
                     "Where a.�׶�id = b.Id And b.��id=c.id(+) And g.id=A.·����¼ID " & vbNewLine & _
                     "Order By ����,����, NVL(c.���,b.���)"
        Else
            strSql = _
            "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, To_Char(a.����, 'day') ����, b.���� As �׶���, b.���, b.˵��, b.��id,Decode(g.·��id,b.·��id,1,0) as ����" & vbNewLine & _
                     "From (Select a.�׶�id, a.����, a.����,a.·����¼id" & vbNewLine & _
                     "       From ����·��ִ�� A" & vbNewLine & _
                     "       Where a.·����¼id = [1]" & vbNewLine & _
                     "       Group By a.�׶�id, a.����, a.����,a.·����¼id) A, �ٴ�·���׶� B,�ٴ�·���׶� C,�ٴ�·����֧ D,�ٴ�·���׶� E,�ٴ�·���׶� F,�����ٴ�·�� G" & vbNewLine & _
                     "Where a.�׶�id = b.Id And b.��id=c.id(+) And b.��֧id=d.id(+) and d.ǰһ�׶�id=e.id(+) And e.��id=f.id(+)  And g.id=A.·����¼ID " & vbNewLine & _
                     "Order By ����,����, Decode(b.��֧ID,Null,NVL(c.���,b.���),NVL(c.���,b.���)+NVL(f.���,e.���))"
        End If

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        .AutoSizeMode = flexAutoSizeRowHeight
        .Cols = .Cols + rsTmp.RecordCount
        For i = 1 To rsTmp.RecordCount
            .ColWidth(i) = 2800
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColData(i) = Val("" & rsTmp!�׶�ID)
            If IsNull(rsTmp!��ID) Then
                .TextMatrix(EFixedRow.R0�׶���, i) = Replace(rsTmp!�׶���, vbLf, vbCrLf)    'Ϊ�˴�ӡʱ��������(vbLfʱ������ʾ������)
            Else
                .TextMatrix(EFixedRow.R0�׶���, i) = Replace(rsTmp!�׶���, vbLf, vbCrLf) & ",��֧:" & Nvl(rsTmp!˵��, rsTmp!���)
            End If
            .TextMatrix(EFixedRow.R1����, i) = "��" & rsTmp!���� & "��"
            .Cell(flexcpData, EFixedRow.R1����, i) = rsTmp!����
            .TextMatrix(EFixedRow.R2����, i) = rsTmp!���� & "(" & rsTmp!���� & ")"
            .Cell(flexcpData, EFixedRow.R2����, i) = rsTmp!���� & ""
            If rsTmp!���� = mPP.��ǰ���� Then mPP.��ǰ���� = rsTmp!����
            rsTmp.MoveNext
        Next
        
        For i = 1 To mcolReason.count
            mcolReason.Remove 1 'ɾ���ֲ���������(�����ƶ���,���¼���ʱ��Ҫ��ձ���ԭ��)
        Next i
        strSql = _
        "Select a.Id, Nvl(b.ͼ��id, a.ͼ��id) ͼ��id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, a.����, a.�׶�id, Nvl(a.��Ŀ���, b.��Ŀ���) As ��Ŀ���," & vbNewLine & _
                 "Nvl(b.��Ŀ����, a.��Ŀ����) ��Ŀ����, a.��Ŀid, Decode(a.ִ����, Null, 0, 1) ִ��״̬, Nvl(b.ִ�з�ʽ, 1) ִ�з�ʽ, a.���ԭ��,NVl(a.����ʱ������,0) as ����ʱ������, c.���� As ����ԭ��," & vbNewLine & _
                 "Nvl(b.��Ŀ���, a.��Ŀ���) As ��Ŀ���, a.ִ�н��, d.·��id, d.��֧id,NVL(NVL(A.������,B.������),1) as ������" & vbNewLine & _
                 "From ����·��ִ�� A, �ٴ�·����Ŀ B, ���쳣��ԭ�� C, �ٴ�·���׶� D" & vbNewLine & _
                 "Where a.·����¼id = [1] And a.��Ŀid = b.Id(+) And a.����ԭ�� = c.����(+) And a.�׶�id + 0 = d.Id" & vbNewLine & _
                 Decode(Val(optSelect(IX_ALL).Tag), 0, " ", 1, " And Decode(a.��Ŀid, Null, Nvl(a.������, 1), Nvl(b.������, 1)) = 1 ", 2, " And Decode(a.��Ŀid, Null, Nvl(a.������, 1), Nvl(b.������, 1)) = 2") & vbNewLine & _
                 "Order By a.����, ����, ��Ŀ���"
        'Nvl(a.������,1)��Ϊ�˼�����ǰ�汾,������������ʱĬ��Ϊҽ����
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        For lngCol = .FixedCols To .Cols - 1
            rsTmp.Filter = "�׶�ID='" & .ColData(lngCol) & "' And ����=" & Val(Replace(.TextMatrix(EFixedRow.R1����, lngCol), "��", ""))
            strOldType = ""
            If rsTmp.RecordCount > 0 Then
                If lngPrePathID <> Val(rsTmp!·��ID & "") Or lngBranchID <> Val(rsTmp!��֧ID & "") Then
                    '·����ת�������ָ���
                    If lngPrePathID <> 0 Or Val(rsTmp!��֧ID & "") <> 0 Then Call .CellBorderRange(0, lngCol, .Rows - 1, lngCol, vbBlack, 1, 0, 0, 0, 0, 1)
                    lngPrePathID = rsTmp!·��ID
                    lngBranchID = Val(rsTmp!��֧ID & "")
                End If
            End If

            Do While Not rsTmp.EOF
                If strOldType <> rsTmp!���� Then
                    lngRow = CPos("T" & rsTmp!����)
                    strOldType = rsTmp!����
                End If

                If mbln����ִ�л��� Then
                    .TextMatrix(lngRow, lngCol) = IIf(rsTmp!ִ�з�ʽ = 0, "", IIf(rsTmp!ִ��״̬ = 0, C_UnExe, C_Exe)) & rsTmp!��Ŀ����
                Else
                    .TextMatrix(lngRow, lngCol) = "" & rsTmp!��Ŀ����   'ҽ��������Ӻ󣬻�δ����·������Ŀǰˢ�£���Ŀ����Ϊ��
                End If
                '����������֯��ʽ ID|��ĿID|��Ŀ���|������|����ʱ������
                '·������Ŀ��ĿidΪ��
                '������ 1-ҽ��,2-��ʿ
                .Cell(flexcpData, lngRow, lngCol) = Val(rsTmp!ID) & "|" & Val("" & rsTmp!��ĿID) & "|" & Val("" & rsTmp!��Ŀ���) & "|" & rsTmp!������ & "|" & rsTmp!����ʱ������
                
                If IsNull(rsTmp!��ĿID) Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = CON_PathOutItemColor         '·������Ŀ��ǳ��ɫ
                    If Val(rsTmp!����ʱ������ & "") = 2 Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = CON_PathOutItemColorBlue     '�ݴ�����Ŀ,ǳ��ɫ
                    End If
                    mcolReason.Add "����˵����" & rsTmp!���ԭ�� & vbCrLf & "����ԭ��" & rsTmp!����ԭ��, "C" & rsTmp!ID
                    If rsTmp!����ԭ�� & "" <> "" Or rsTmp!���ԭ�� & "" <> "" Then
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "����ԭ��" & rsTmp!����ԭ�� & vbCrLf & "����˵����" & rsTmp!���ԭ��
                    End If
                ElseIf InStr("124", CStr(rsTmp!ִ�з�ʽ)) > 0 Then '�������ɵģ�δ����
                    If Not IsNull(rsTmp!����ԭ��) Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED    'ǳ��ɫ
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "����ԭ��" & rsTmp!����ԭ��
                    End If
                ElseIf rsTmp!ִ�з�ʽ = 3 Then                          '��ѡ�����ɫ
                    .Cell(flexcpForeColor, lngRow, lngCol) = &HC00000
                    If Not IsNull(rsTmp!����ԭ��) Then '93648 ��ҩ·����Ŀ�ı���ԭ��
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED    'ǳ��ɫ
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "����ԭ��" & rsTmp!����ԭ��
                    End If
                End If

                If InStr(rsTmp!��Ŀ���, "|") > 0 And Not IsNull(rsTmp!ִ�н��) Then
                    i = Val(Mid(rsTmp!��Ŀ���, InStr(rsTmp!��Ŀ���, rsTmp!ִ�н��) + Len(rsTmp!ִ�н��) + 1, 1))
                    If i > 0 Then i = Getִ�н������ͼ��(i)
                Else
                    i = 0
                End If

                If Not IsNull(rsTmp!ͼ��ID) Or i > 0 Then
                    .Cell(flexcpPictureAlignment, lngRow, lngCol) = flexPicAlignRightCenter    ' flexPicAlignLeftCenter
                    If i > 0 Then
                        .Cell(flexcpPicture, lngRow, lngCol) = imgCharacter.ListImages(i).Picture
                    Else
                        .Cell(flexcpPicture, lngRow, lngCol) = GetPathIcon(rsTmp!ͼ��ID)
                    End If
                End If

                lngRow = lngRow + 1
                rsTmp.MoveNext
            Loop
        Next

        '4)��ʾ������Ϣ

        If .Rows = .FixedRows And .Cols = .FixedCols Then
            Call ClearPathItem(True)
            .BackColorSel = vbWhite
            .ForeColorSel = vbBlack
        Else
            If Val(optSelect(IX_ALL).Tag) <> IX_��ʿ Then 'Ŀǰδ���ǻ�ʿ��������,��ѡ�л�ʿ����ʾ������Ϣ
                .BackColorSel = &H8000000D
                .ForeColorSel = &H8000000E
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                lngDayRow = .FixedRows - 2
                .TextMatrix(lngRow, .FixedCols - 1) = "�������"
                .Cell(flexcpBackColor, lngRow, 0) = .BackColorFixed  '&HEFF0E0      '&HD0EFFF
                Call .CellBorderRange(.Rows - 1, 0, .Rows - 1, .Cols - 1, vbBlack, 0, 1, 0, 0, 0, 0)
    
                strSql = "Select a.�׶�id, a.����, a.�������, a.����˵��, a.������,a.����ʱ��, c.���� As ����ԭ��, a.���������, Nvl(a.ʱ�����, 0) ʱ�����, a.��ת�����, a.ԭ·��id" & vbNewLine & _
                        "From ����·������ A, ����·������ B, ���쳣��ԭ�� C" & vbNewLine & _
                        "Where a.·����¼id = b.·����¼id(+) And a.�׶�ID=B.�׶�ID(+) And a.����=b.����(+) And a.·����¼id = [1] And b.����ԭ�� = c.����(+)"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
                For lngCol = .FixedCols To .Cols - 1
                    .Cell(flexcpBackColor, lngRow, lngCol) = &HEDF8FF   '&HD0EFFF
    
                    rsTmp.Filter = "�׶�ID='" & .ColData(lngCol) & "' And ����=" & Val(Replace(.TextMatrix(EFixedRow.R1����, lngCol), "��", ""))
                    str����ԭ�� = ""
                    For j = 1 To rsTmp.RecordCount
                        '��ȡ�������ԭ��
                        str����ԭ�� = str����ԭ�� & rsTmp!����ԭ�� & "��"
                        If j = rsTmp.RecordCount Then
                            str����ԭ�� = Mid(str����ԭ��, 1, Len(str����ԭ��) - 1)
                            If InStr(rsTmp!����˵��, vbCrLf) = 0 Or IsNull(rsTmp!����˵��) Then
                                strTmp = "" & rsTmp!����˵��
                            Else
                                arrtmp = Split(rsTmp!����˵��, vbCrLf)
                                strTmp = ""
                                For i = 0 To UBound(arrtmp)
                                    strTmp = strTmp & vbCrLf & Space(4) & (i + 1) & "." & arrtmp(i)
                                Next
                            End If
                            strTmp = strTmp & vbCrLf & "�� �� �ˣ�" & rsTmp!������
                            If rsTmp!������� = 1 Then
                                str������� = "����"
                            ElseIf mPP.����·��״̬ = 3 And lngCol = .Cols - 1 Then
                                str������� = "������˳�" & vbCrLf & "����ԭ��" & str����ԭ�� & vbCrLf & "�� �� �ˣ�" & rsTmp!���������
        
                            ElseIf mPP.����·��״̬ = 2 And lngCol = .Cols - 1 Then
                                str������� = "��������" & vbCrLf & "����ԭ��" & str����ԭ��
                            Else
                                str������� = "��������" & vbCrLf & "����ԭ��" & str����ԭ��
                                If Not IsNull(rsTmp!���������) Then str������� = str������� & vbCrLf & "�� �� �ˣ�" & rsTmp!���������
                            End If
        
                            .TextMatrix(lngRow, lngCol) = "���������" & str������� & vbCrLf & "����˵����" & strTmp
                            If rsTmp!������� = -1 Then
                                .Cell(flexcpForeColor, lngRow, lngCol) = vbRed     '�����ú�ɫ��ʾ
                            End If
        
                            If rsTmp!ʱ����� = 1 Or rsTmp!ʱ����� = 2 Then
                                '��ǰ
                                .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "��"
                                .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                            ElseIf rsTmp!ʱ����� = -1 Then    '�Ӻ�
                                .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "��"
                                .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                            End If
                            'δ��˵���ת�׶�
                            If rsTmp!ԭ·��ID & "" <> "" And rsTmp!��ת����� & "" = "" Then
                                .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "(δ���)"
                                .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Next
                    If rsTmp.RecordCount = 0 Then
                        .TextMatrix(lngRow, lngCol) = ""
                    End If
                Next
            End If
        End If
        .Redraw = True
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч

        If lngPreRow <> -1 And lngPreCol <> -1 And lngPreRow <= .Rows - 1 And lngPreCol <= .Cols - 1 Then
            .Select lngPreRow, lngPreCol
        Else
            .Select .FixedRows, .FixedCols
        End If
    End With

    Exit Sub
errH:
    vsPath.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckPathIsTurnAduit() As Boolean
'���ܣ�����Ƿ����δ��˵���ת�׶Ρ�trueΪ����
     Dim strSql As String, rsTmp As Recordset
     
     strSql = "Select 1 From ����·������ Where ԭ·��id is not null And ��ת����� is null And ·����¼ID=[1]"
     
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ת���", mPP.����·��ID)
     
     CheckPathIsTurnAduit = rsTmp.RecordCount > 0
     Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Get����·����Ϣ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, Optional ByVal lng·����¼ID As Long, Optional ByVal blnReadOnly As Boolean)
'���ܣ���ȡ���˵��ٴ�·����Ϣ
'������lng·����¼ID=��һ�������ж���·��ʱ��ˢ��ָ��·����¼ID��·����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    'һ��סԺֻ֧��һ��·�������ܿ���
    ' And (����ID = [3] Or Exists(Select 1 From ��������˵�� B Where (a.id = b.����id or b.����id = [3]) and b.��������='ICU'))
    '��ǰ�׶�Ϊ0��ʾ��δ���ɹ�·��
    strSql = "Select a.ID,a.·��ID,c.·��ID as ԭ·��ID,a.�汾��,a.״̬,a.��ǰ�׶�ID,a.��ǰ����,b.���� as δ����ԭ��,c.��ID,c.��֧ID,d.��֧ID as ǰһ�׶η�֧ID,e.����·������,a.�ϲ�·������,e.���� as ·������,a.������,a.����ʱ��,a.����ʱ��" & _
            " From �����ٴ�·�� A,���쳣��ԭ�� B,�ٴ�·���׶� C,�ٴ�·���׶� D,�ٴ�·��Ŀ¼ E" & _
            " Where a.����ID = [1] And a.��ҳID = [2] And a.·��ID=e.id And a.δ����ԭ�� = b.����(+) And a.��ǰ�׶�ID = c.ID(+) And a.ǰһ�׶�ID=d.id(+)" & _
            IIf(lng·����¼ID <> 0, " And a.ID=[4] ", "") & _
            " Order By a.����ʱ�� Desc"  'ȡ���һ�ε����·����֧��һ��סԺ���·��2012-10-25��
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get����·����Ϣ", lng����ID, lng��ҳID, lng����ID, lng·����¼ID)
    If rsTmp.RecordCount > 0 Then
        mPP.ԭ·��ID = Val("" & rsTmp!ԭ·��ID)
        mPP.·��ID = rsTmp!·��ID
        mPP.�汾�� = rsTmp!�汾��
        mPP.����·��ID = rsTmp!ID
        mPP.����·��״̬ = rsTmp!״̬
        mPP.��ǰ�׶�ID = Val("" & rsTmp!��ǰ�׶�ID)
        mPP.�׶θ�ID = Val("" & rsTmp!��ID)
        mPP.��ǰ���� = Val("" & rsTmp!��ǰ����)
        mPP.����ʱ�� = CDate(Nvl(rsTmp!����ʱ��, 0))
        mPP.��ǰ���� = "0" '��LoadPathItem�и�ֵ
        mPP.δ����ԭ�� = "" & rsTmp!δ����ԭ��
        mPP.��ǰ�׶η�֧ID = Val("" & rsTmp!��֧ID)
        '����·�����������յ�ǰ�׶�ID������ȡǰһ�׶εķ�֧ID
        If mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3 Then
            mPP.��ǰ�׶η�֧ID = Val(rsTmp!ǰһ�׶η�֧ID & "")
        End If
        mPP.����·������ = Val(rsTmp!����·������ & "")
        mPP.�ϲ�·������ = Val(rsTmp!�ϲ�·������ & "")
        If lng·����¼ID = 0 Then mlngPathCount = rsTmp.RecordCount
    Else
        mPP.ԭ·��ID = 0
        mPP.·��ID = 0
        mPP.�汾�� = 0
        mPP.����·��ID = 0
        mPP.����·��״̬ = -1
        mPP.��ǰ�׶�ID = 0
        mPP.�׶θ�ID = 0
        mPP.��ǰ���� = 0
        mPP.��ǰ���� = "0"
        mPP.δ����ԭ�� = ""
        mPP.��ǰ�׶η�֧ID = 0
        mPP.����·������ = 0
        mPP.�ϲ�·������ = 0
        mPP.����ʱ�� = CDate(0)
        mlngPathCount = 0
    End If
    
        If blnReadOnly Then Exit Sub
    If mlngPathCount > 1 Then
        fraPath.Visible = True
        lblInDiag.Caption = "�������:" & Get�������(mPP.����·��ID)
        lblInPep.Caption = "������:" & rsTmp!������
        lblInDate.Caption = "����ʱ��:" & Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:mm")
        lblOutDate.Caption = "����ʱ��:" & Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:mm")
        If lng·����¼ID = 0 Then
            cboPath.Clear
            Do While Not rsTmp.EOF
                cboPath.AddItem rsTmp!·������ & ""
                cboPath.ItemData(cboPath.NewIndex) = rsTmp!ID & ""
                rsTmp.MoveNext
            Loop
            zlControl.CboSetIndex cboPath.Hwnd, 0
        End If
    Else
        fraPath.Visible = False
        cboPath.Clear
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get�������(ByVal lng·����¼ID As Long) As String
'���ܣ�ȡ������ϵ�����
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select B.������� From �����ٴ�·�� A,������ϼ�¼ B Where " & _
            " a.����id = b.����id And a.��ҳid = b.��ҳid  and a.������� = b.������� And a.�����Դ = b.��¼��Դ And NVL(a.����id,0) = NVL(b.����id,0) And NVL(a.���id,0) = NVL(b.���id,0) And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get�������", lng·����¼ID)
    If rsTmp.RecordCount > 0 Then Get������� = rsTmp!������� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlPrintOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'���ܣ��ٴ�·������ӡ
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'     blnIsSetup-��ʾ������ӡ�������д�ӡǰ����
'     ��bytStyle=4ʱ����Ҫ����strPDFFile=PDF���Ĭ��·��,�����ļ�������׺
    Call FuncPathTableOutput(bytStyle, blnIsSetup, strPDFFile, strDeviceName)
End Sub

Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
        ByVal int����״̬ As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal blnForceRefresh As Boolean = True, Optional ByVal lngState As Long, _
        Optional ByVal lng·����¼ID As Long, Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object, Optional ByVal blnReadOnly As Boolean) As Long
'������lng·����¼ID=��һ�������ж���·��ʱ��ˢ��ָ��·����¼ID��·����
'      blnForceRefresh=True δ�л�����ˢ��ʱҲ����ˢ�£�����ˢ��
'      objMip ��Ϣ����
'      blnReadOnly=ֻ��ȡ������Ϣ������������
    
    Dim objControl As CommandBarControl
    Dim strPrePati As String
    
    strPrePati = mPati.����ID & "_" & mPati.��ҳID
    If strPrePati = lng����ID & "_" & lng��ҳID And lng����ID <> 0 And Not blnForceRefresh Then Exit Function       '����֮ǰ��Ԫ��λ�ò���
    
    
    If mPati.����ID & "_" & mPati.��ҳID = lng����ID & "_" & lng��ҳID And lng����ID <> 0 And Not blnForceRefresh Then Exit Function       '����֮ǰ��Ԫ��λ�ò���
    
    mPati.����ID = lng����ID
    mPati.��ҳID = lng��ҳID
    mPati.����ID = lng����ID
    mPati.����ID = lng����ID
    mPati.����״̬ = int����״̬
    mlngState = lngState
    mlngҽ������ID = lngҽ������ID
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mlng����״̬ = Get���˲���״̬(lng����ID, lng��ҳID, mlngӤ������ID, mlngӤ������ID)
    mblnMoved = blnMoved

    Set mcolReason = New Collection

    Call Get����·����Ϣ(lng����ID, lng��ҳID, lng����ID, lng·����¼ID, blnReadOnly)
    
    '�°没������ʱ���������������ֻ������������ᵼ������������
    If blnReadOnly = True Then Exit Function
    lblPrinted.Visible = False 'Ĭ�ϲ���ʾ��LoadPathFlow�������Ƿ���ʾ
    fraSendor.Tag = ""
    
    If mPP.����·��ID = 0 Then
        Call SetUnImport
    Else
        If mPP.����·��״̬ = 0 Then
            Call SetImportFalse
        Else
            Call LoadPathFlow
            Call LoadPathItem
        End If
    End If
    fraTop.Visible = True  '��fraTop������������������丳ֵ��Ч��ʼ����false
    fraSendor.Visible = (mPP.����·��״̬ > 0 And fraSendor.Tag <> "����")
    fraTop.Visible = fraPath.Visible Or fraSendor.Visible
    Call Form_Resize    '����·�����̱��Ƿ��й������������߶�
    If strPrePati <> lng����ID & "_" & lng��ҳID And lng����ID <> 0 And mlngPlugInID <> 0 Then
        If mblnInsideTools Then
            Set objControl = cbsSub.FindControl(, mlngPlugInID, , True)
        Else
            Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        End If
        If Not objControl Is Nothing Then
            objControl.Execute
        End If
    End If
End Function


Private Sub InitCbsSubBar()
    Dim objBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    
    
    Set objBar = cbsSub.Add("�ڲ�������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 24, 24
    objBar.Visible = False  'ֻ���ڲ�����ʱ����ʾ(zlDefCommandBars)
    
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByVal int���� As Integer, Optional ByVal blnInsideTools As Boolean = False)
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim lngStart As Long, i As Long
 
    mint���� = int����
    mbln����ִ�л��� = Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, P�ٴ�·��Ӧ��, 1))
    If mbln����ִ�л��� Then
        mstrִ�г��� = zlDatabase.GetPara("·��ִ�л������ó���", glngSys, P�ٴ�·��Ӧ��, 11)
    End If
    mbln���ò����� = Val(zlDatabase.GetPara("����ǰһ�첻���������ɽ����·����Ŀ", glngSys, P�ٴ�·��Ӧ��, 1))
    mbln������ǰ���� = Val(zlDatabase.GetPara("������ǰ���������·����Ŀ", glngSys, P�ٴ�·��Ӧ��, 1))
    mbytPrintWay = Val(zlDatabase.GetPara("·������ӡ��ʽ", glngSys, P�ٴ�·��Ӧ��, "0"))
    mblnInsideTools = blnInsideTools

    Set mfrmParent = frmParent

    If cbsMain Is Nothing Then Exit Sub
    If mrsPlugInBar Is Nothing Then
        Call GetPlugInBar(P�ٴ�·��Ӧ��, mint����, mrsPlugInBar)
    End If
    
    Set mcbsMain = cbsMain
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '�ļ��˵�
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Find(, conMenu_File_Excel)
        objControl.Caption = "�����&Excel(ҽʦ��)��"
        Set objControl = .Find(, conMenu_File_Print)
        objControl.Caption = "��ӡ·����(ҽʦ��)(&P)"
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_PatiPath, "��ӡ·����(���߰�)(&Q)", objControl.Index + 1)
        objControl.IconId = conMenu_File_Print
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objPopup.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "·��(&E)", objPopup.Index + 1, False)
    objPopup.ID = conMenu_EditPopup
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "����·��(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ������")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_ImportMerge, "����ϲ�·��")
        objControl.IconId = conMenu_Edit_Import
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnImportMerge, "ȡ���ϲ�·��")
        objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewMergeImport, "�鿴�ϲ�·����������")
        objControl.IconId = conMenu_Edit_Select

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����·����Ŀ(&C)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����������Ŀ(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "��������ҽ��")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "���·������Ŀ")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�·������Ŀ")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ȡ����������(&X)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ȡ����ǰ��Ŀ(&V)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive, "��Ŀִ�еǼ�(&E)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "ȡ��ִ�еǼ�(&Z)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "����ִ�еǼ�(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "����ȡ��ִ��(&F)")


        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����(&D)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�޸�����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "ȡ������")


        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "���·��(&O)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ȡ�����")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogModi, "�޸ĳ����ǼǱ�")


        Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "����")
        objControl.IconId = conMenu_Manage_Up
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "����")
        objControl.IconId = conMenu_Manage_Down
        '��Ҳ˵�
        Call DefCommandPlugInPopup(objPopup.CommandBar.Controls, mrsPlugInBar)
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "��׼·���ο�")
        objControl.BeginGroup = True
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Select, "�鿴��������")
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "�鿴�����ǼǱ�")
        'Set objControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)", objControl.Index + 1)
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        '
        '        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
        '        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "�鿴������(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_View, "�鿴��Ŀ����(&A)")
    End With

    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objPopup.Index, False)
        objPopup.ID = conMenu_ToolPopup
    End If

    '����������
    '-----------------------------------------------------
    lngStart = 0
    If blnInsideTools Then
        Set cbrToolBar = cbsSub(2)
        Set objControl = cbrToolBar.FindControl(, conMenu_Edit_Import)
        If objControl Is Nothing Then lngStart = 1: cbrToolBar.Visible = True
        For i = cbrToolBar.Controls.count To 1 Step -1
            If cbrToolBar.Controls(i).ID > conMenu_Tool_PlugIn_Item And cbrToolBar.Controls(i).ID < conMenu_Tool_PlugIn_Item + 100 Or cbrToolBar.Controls(i).ID = conMenu_Tool_PlugIn Then
                cbrToolBar.Controls(i).Delete
            End If
        Next i
    Else
        Set cbrToolBar = cbsMain(2)
        For Each objControl In cbrToolBar.Controls    '�����ǰ������һ��Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = cbrToolBar.Controls(objControl.Index - 1): Exit For
            End If
        Next
        lngStart = objControl.Index + 1
    End If

    If lngStart <> 0 Then
        With cbrToolBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "����", lngStart)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", objControl.Index + 1)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����", objControl.Index + 1)
            objControl.ToolTipText = "�������ɿ�ѡ���ɵ�·����Ŀ"

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "ִ��", objControl.Index + 1)
            objControl.BeginGroup = True
            objControl.ToolTipText = "����ִ��·����Ŀ"
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����", objControl.Index + 1)
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "���", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "����", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Up
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "����", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Down
       
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "����", objControl.Index + 1)
            objPopup.BeginGroup = True
            objPopup.IconId = conMenu_Manage_Report
            objPopup.ToolTipText = "���ı���"
        End With

        If blnInsideTools Then
            For Each objControl In cbrToolBar.Controls
                If objControl.Type <> xtpControlLabel Then
                    objControl.Style = xtpButtonIconAndCaption
                End If
            Next
        End If
    End If


    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        '.Add FCONTROL, Asc("O"), conMenu_File_Open
'        .Add 0, vbKeyF11, conMenu_Tool_Option    '·��ѡ��
    End With
    
    '��ҳ����������
    Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
End Sub

Private Sub DefCommandPlugIn(ByVal cbsMain As Object, ByRef rsBar As ADODB.Recordset)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim strFuncName As String, lngFuncID As Long
    Dim strFunc As String, i As Long
    
    Dim blnGroup As Boolean
    Dim lngTmp  As Long
    
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
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������)
                        objControl.IconId = rsBar!ͼ��ID
                        objControl.Parameter = rsBar!������
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                        blnGroup = True
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
            With objMenu.CommandBar.Controls
                If rsBar.RecordCount = 1 Then
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                        objControl.IconId = rsBar!ͼ��ID
                        objControl.Parameter = rsBar!������
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                        blnGroup = True
                    End If
                Else
                    Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����", , False)
                        objPopup.BeginGroup = True
                    With objPopup.CommandBar.Controls
                        For i = 1 To rsBar.RecordCount
                            Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                            objControl.IconId = rsBar!ͼ��ID
                            objControl.Parameter = rsBar!������
                            objControl.Style = xtpButtonIconAndCaption
                            If Val(rsBar!IsGroup) = 1 Then
                                objControl.BeginGroup = True
                                blnGroup = True
                            End If
                            rsBar.MoveNext
                        Next
                    End With
                End If
            End With
        End If
    End If
    
    '��������ť
    If mblnInsideTools Then
        Set objBar = cbsSub(2)
    Else
        Set objBar = cbsMain(2)
    End If
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

Private Sub FuncPatiPathPrint()
'���ܣ�������߰��ٴ�·��
    Dim WordApp As Object   'Word.Application
    Dim WordDoc As Object     'Word.Document
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim strFileName As String, strFilePath As String
    Dim lngRetu As Long, strInfo As String

    If vsPath.FixedRows < 3 Then
        MsgBox "�ò��˻�δ�����ٴ�·����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error GoTo errH
    '���·��
    Screen.MousePointer = 11
    strSql = "Select �ļ��� from �ٴ�·���ļ� where ·��ID=[1] And ���=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID)
    If rsTmp.RecordCount > 0 Then
        strFileName = rsTmp!�ļ��� & ""
        strFilePath = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & strFileName
        If gobjFile.FileExists(strFilePath) Then gobjFile.DeleteFile strFilePath, True
        '�����ݿ���BLOB���ݶ���������ʱ�ļ�Ŀ¼��
        strFilePath = Sys.ReadLob(glngSys, 10, mPP.·��ID & "," & strFileName, strFilePath)
        If Not gobjFile.FileExists(strFilePath) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Sub
        End If
    Else
        Screen.MousePointer = 0
        MsgBox "��·����û�����ö�Ӧ���ٴ�·����(���߰�),�뵽�ٴ�·�����������á�", vbInformation, Me.Caption
        Exit Sub
    End If

    Set WordApp = CreateObject("Word.Application")
    If WordApp Is Nothing Then
        MsgBox "�밲װMicrosoft Office Word��", vbInformation, gstrSysName
        Exit Sub
    End If

    Set WordDoc = WordApp.Documents.Open(strFilePath)      '��RTF�ĵ�
    WordDoc.PrintPreview
    WordApp.Visible = True
    WordApp.ScreenUpdating = True
    WordApp.Activate
    Screen.MousePointer = 0
    
    '��¼��ӡ��Ϣ
    Call zlDatabase.ExecuteProcedure("Zl_���Ӳ�����ӡ_Insert(" & mPP.����·��ID & ",12," & mPati.����ID & "," & mPati.��ҳID & ",'" & UserInfo.���� & "')", "��ӡ���߰�·����")
    '��ӡ��ǿ�����¼�����ʾ��Ϣ��������ʾ��Ϣ
    Call LoadPathFlow
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathTableOutput(bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'���ܣ�����ٴ�·����
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'     blnIsSetup-������ӡ�����д�ӡǰ����
'     strPDFFile=PDF���Ĭ��·��
'     strDeviceName=ָ����ӡ������
    Dim rsTmp As ADODB.Recordset
    Dim vsBody As VSFlexGrid
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngColor As Long, bytR As Byte
    Dim strSql As String
    Dim rsSQLTmp As ADODB.Recordset
    Dim strDisease As String        '�������
    Dim strStandardDate As String   '��׼סԺ��
    Dim i As Long, j As Long
    Dim strTitle As String
    Dim strTmp As String
    Dim lngDefDay As Long
    If mbytPrintWay = 1 Then
    
        If bytStyle = 1 Then
            bytStyle = 2
        ElseIf bytStyle = 2 Then
            bytStyle = 1
        End If
        Call FuncPathTableReport(bytStyle)
    Else
        strSql = "Select a.����id, a.��ҳid, b.����id, b.���id, b.�������, c.��׼סԺ��" & vbNewLine & _
                 "From �����ٴ�·�� A, ������ϼ�¼ B, �ٴ�·���汾 C" & vbNewLine & _
                 "Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.������� = b.�������" & vbNewLine & _
                 "      And a.�����Դ = b.��¼��Դ And c.·��id = a.·��id And c.�汾�� = a.�汾�� And" & vbNewLine & _
                 "      b.��ϴ��� = 1 And a.����id = [1] And a.��ҳid = [2] And a.ID=[3]"
    
        mblnUnChange = True
        If vsPath.FixedRows < 3 Then
            '���PDF���������·�����ˣ���ֱ���˳�����ʾ
            If bytStyle = 4 Then Exit Sub
            '������ӡ����ʾ
            If blnIsSetup Then Exit Sub
            MsgBox "�ò��˻�δ�����ٴ�·����Ŀ��", vbInformation, gstrSysName
            Exit Sub
        End If
        On Error GoTo errH
        Set rsTmp = GetPatiInfo(mPati.����ID, mPati.��ҳID)
        Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.����ID, mPati.��ҳID, mPP.����·��ID)
    
        If rsSQLTmp.RecordCount > 0 Then
            strDisease = rsSQLTmp!������� & ""
            strStandardDate = rsSQLTmp!��׼סԺ�� & ""
        Else
            strDisease = ""
            strStandardDate = ""
        End If
        '��ͷ
        If InStr(vsFlow.TextMatrix(0, 0), vbCrLf) > 0 Then
            strTitle = Mid(vsFlow.TextMatrix(0, 0), 1, InStr(vsFlow.TextMatrix(0, 0), vbCrLf) - 1)
        Else
            strTitle = vsFlow.TextMatrix(0, 0)
        End If
        objOut.Title.Text = strTitle & vbCrLf & "�ٴ�·����"
        objOut.Title.Font.Name = "����_GB2312"
        objOut.Title.Font.Size = 20
        objOut.Title.Font.Bold = True
    
        '����
        strSql = "Select a.�������" & vbNewLine & _
                 "From ������ϼ�¼ A" & vbNewLine & _
                 "Where a.����id = [1] And a.��ҳid = [2] And a.��¼��Դ = 3 And a.������� In (2, 12) Order By a.��ϴ���"
        Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.����ID, mPati.��ҳID)
        If rsSQLTmp.RecordCount > 0 Then
            strTmp = rsSQLTmp!������� & ""
            strTmp = Mid(strTmp, InStr(strTmp, ")") + 1) & Mid(strTmp, 1, InStr(strTmp, ")"))
        Else
            strTmp = ""
        End If
        strSql = "Select a.�������� || Decode(Nvl(a.������Ŀid, 0), 0, '(ICD9CM-3:' || b.���� || ')', '(������Ŀ:' || c.���� || ')') As ��������" & vbNewLine & _
                 "From ���������¼ A, ��������Ŀ¼ B, ������ĿĿ¼ C" & vbNewLine & _
                 "Where a.����id = [1] And a.��ҳid = [2] And a.��¼��Դ = 3 And a.��������id = b.Id(+) And a.������Ŀid = c.Id(+)" & vbNewLine & _
                 "Order By a.������ʼʱ��"
        Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.����ID, mPati.��ҳID)
    
        If rsSQLTmp.RecordCount > 0 Then
            strTmp = strTmp & " �� " & rsSQLTmp!��������
        End If
        Set objRow = New zlTabAppRow
        objRow.Add "���ö��󣺵�һ���Ϊ " & strTmp
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "����������" & rsTmp!���� & " �Ա�" & rsTmp!�Ա� & " ���䣺" & rsTmp!���� & " סԺ�ţ�" & rsTmp!סԺ�� & " ����ţ�" & rsTmp!����� & ""
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "סԺ����:" & Format(rsTmp!��Ժ����, "yyyy��MM��dd��")
        objRow.Add "��Ժ����:" & Format(rsTmp!��Ժ����, "yyyy��MM��dd��")
        objRow.Add "��׼סԺ�գ�" & IIf(InStr(strStandardDate, "-") > 0, "", "��") & strStandardDate & "��"
        objOut.UnderAppRows.Add objRow
        objOut.AppFont.Size = 12
        '����
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & UserInfo.����
        objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
        objOut.BelowAppRows.Add objRow
    
        'ҳ��
        objOut.Footer = ";��[ҳ��]ҳ����[ҳ��]ҳ;"
        objOut.PageFooter = 5
    
        '����
        strTmp = zlDatabase.GetPara("·������ӡ����", glngSys, P�ٴ�·��Ӧ��, "0")
        If strTmp = "1" Then
            Set vsBody = FuncConvertPathTable
        Else
            Set vsBody = vsPath
        End If
        
        '���
        With vsBody
            .Redraw = flexRDNone
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "��ʿǩ��"
            .RowHeight(.Rows - 1) = 440
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "ҽ��ǩ��"
            .RowHeight(.Rows - 1) = 440
            
            'Ĭ�ϴ�ӡ����
            lngDefDay = Val(zlDatabase.GetPara("·����ÿҳ��ӡ������", glngSys, P�ٴ�·��Ӧ��, "2"))
            objOut.PageCols = lngDefDay + .FixedCols
            '�����������ʱ���������
            If (.Cols - 1) Mod lngDefDay <> 0 Then
               .Cols = .Cols + (lngDefDay - ((.Cols - 1) Mod lngDefDay))
            End If
            '��ӡ���ת��
            Call FuncPathTableChange(vsBody, lngDefDay)
           
            '�ƻ��ϲ�������,��ӡ�����жԺϲ����е�������
            For i = .FixedCols To .Cols - 1
                If i Mod 2 = 0 Then
                    .TextMatrix(R0�׶���, i) = .TextMatrix(R0�׶���, i) & vbTab
                End If
            Next
            .Redraw = flexRDDirect
            '�п�����Ӧ
            If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч
    
            objOut.FixCol = vsBody.FixedCols
            objOut.FixRow = vsBody.FixedRows
            Set objOut.Body = vsBody
    
            'ָ����ӡ��
            If strDeviceName <> "" Then SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "DeviceName", strDeviceName
            If bytStyle = 1 Or bytStyle = 4 Then
                If bytStyle = 4 Then
                    bytR = 4
                    objOut.Privileged = True '���Ӳ������� ���ڹ�������Zl9PrintMode�ڲ�������ӡȨ�޼��
                Else
                    If Not blnIsSetup Then
                        bytR = zlPrintAsk(objOut)
                    Else
                        bytR = 1
                    End If
                End If
                Me.Refresh
                
                If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR, strPDFFile
                '��ӡ���˲�����ӡ��¼
                strSql = "zl_���Ӳ�����ӡ_insert(" & mPP.����·��ID & ",11," & mPati.����ID & "," & mPati.��ҳID & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Else
                zlPrintOrView1Grd objOut, bytStyle
            End If
            mblnUnChange = False
            '�ָ�����ʼ״̬
            Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)    'vsPath�䶯�����¼���
            
            If vsPathPrint.UBound = 1 Then Unload vsPathPrint(1)
        End With
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathTableReport(ByVal bytType As Byte)
'����:�ñ����ӡ�ٴ�·����
'bytType:0=ȱʡֵ,�ɲ���,��ʾ����(������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF
    Dim arrSQL As Variant
    Dim i As Long, j As Long
    Dim strTmp As String
    
    On Error GoTo errH
    arrSQL = Array()
    With vsPath
        For i = .FixedCols To .Cols - 1
            For j = 0 To .Rows - 1
                strTmp = ""
                If TypeName(.Cell(flexcpData, j, i)) = "String" Then
                    If .Cell(flexcpData, j, i) & "" <> "" Then strTmp = Split(.Cell(flexcpData, j, i), "|")(0) & ""
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_·����ӡ��¼_Insert(" & ZVal(Val(.ColData(i) & "")) & ",'" & .TextMatrix(j, 0) & "'," & i & "," & j & ",'" & .TextMatrix(j, i) & "'," & ZVal(Val(strTmp)) & ")"
            Next
        Next
    End With
    gcnOracle.BeginTrans
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1256", Me, "����ID=" & mPati.����ID, "��ҳID=" & mPati.��ҳID, "����·��ID=" & mPP.����·��ID, bytType)
    gcnOracle.CommitTrans
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
    Case conMenu_Edit_Import, conMenu_Edit_Untread, conMenu_Edit_ImportMerge, conMenu_Edit_UnImportMerge
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
    Case conMenu_Edit_Send, conMenu_Edit_Append, conMenu_Edit_Delete, conMenu_Edit_Blankoff, conMenu_Edit_SendBack
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
        If Control.ID = conMenu_Edit_SendBack And blnVisible Then
            blnVisible = Not InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0
        End If
    Case conMenu_Edit_Surplus, conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";·������Ŀ;") = 0 Then blnVisible = False

    Case conMenu_Edit_Archive, conMenu_Edit_UnArchive, conMenu_Edit_Merge, conMenu_Edit_DeleteParent
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";ִ��·��;") = 0 Or mbln����ִ�л��� = False Then blnVisible = False
        '����·��ִ�л���ʱ�����ó��Ϻ͵�ǰ���ϲ�һ��ʱ,���ز˵���ť
        If blnVisible Then
            If Mid(mstrִ�г���, mint���� + 1, 1) = "0" Then
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_Audit, conMenu_Edit_Reuse, conMenu_Edit_Clear
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";�׶�����;") = 0 Then blnVisible = False

    Case conMenu_Edit_Stop, conMenu_Edit_ClearUp
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";����·��;") = 0 Then blnVisible = False

    Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView
        If Control.ID = conMenu_Edit_OutLogModi Then
            If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
        End If
        If blnVisible Then blnVisible = CheckPathOutLog
    Case conMenu_Edit_Compend
        '���浯��(����ӡ),���ı���
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";�������;") = 0 Then blnVisible = False
    End Select

    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveItem As Boolean
    Dim lng��ĿID As Long

    If vsPath.Redraw = flexRDNone Then Exit Sub

    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub

    With vsPath
        blnHaveItem = .Row > .FixedRows - 1 And .FixedRows <> 0 And .Col > .FixedCols - 1   '.FixedRows=0ʱ��ֻ��һ����ʾ��Ϣ
    End With
    Select Case Control.ID
        '0.���
    Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel, conMenu_File_Print_PatiPath
        Control.Enabled = mPP.����·��ID <> 0

        '1.����
        '-----------------------------------------
    Case conMenu_Edit_Import    '����·��
        Control.Enabled = mlngState <> ps���ת�� And mlngState <> ps��Ժ And mPati.����״̬ = 0 And (mPP.����·��ID = 0 Or mPP.����·��״̬ <> 1) And mPati.����ID <> 0 And cboPath.ListIndex <= 0

    Case conMenu_Edit_Untread   'ȡ������(���ڵ�һ������ʱ��ȡ������)
        Control.Enabled = mPati.����״̬ = 0 And mPP.����·��ID <> 0 And (mPP.����·��״̬ = 0 Or mPP.����·��״̬ = 1) And vsPath.Cols <= vsPath.FixedCols + 1
    Case conMenu_Edit_Select      '�鿴��������
        Control.Enabled = mPati.����״̬ = 0 And mPP.����·��ID <> 0
    Case conMenu_Edit_ImportMerge  '����ϲ�·��
        Control.Enabled = mPati.����״̬ = 0 And mPP.����·��ID <> 0 And cboPath.ListIndex <= 0
    Case conMenu_Edit_UnImportMerge    'ȡ������ϲ�·��
        Control.Enabled = mPati.����״̬ = 0 And mPP.����·��ID <> 0 And cboPath.ListIndex <= 0 And mPP.�ϲ�·������ > 0
    Case conMenu_Edit_ViewMergeImport    '�鿴�ϲ�·����������
        Control.Enabled = mPati.����״̬ = 0 And mPP.����·��ID <> 0 And mPP.�ϲ�·������ > 0
        '2.����
        '-----------------------------------------
    Case conMenu_Edit_Send      '����·��
        Control.Enabled = mPati.����״̬ = 0 And mPP.����·��ID <> 0 And mPP.����·��״̬ = 1

    Case conMenu_Edit_Append    '��������
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Blankoff  'ȡ����������
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Delete, conMenu_Edit_SendBack   'ȡ��·����Ŀ,��������ҽ��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Col > 0 Then
                    Control.Enabled = ((mint���� = 0 And .ColData(.Col) = mPP.��ǰ�׶�ID And .Col = .Cols - 1) _
                        Or (mint���� = 1 And conMenu_Edit_Delete = Control.ID) _
                        Or (conMenu_Edit_Delete = Control.ID And Split(.Cell(flexcpData, .Row, .Col), "|")(4) = 1))
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Surplus   '���·������Ŀ
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1

    Case conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down     '�޸�·������Ŀ
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" And vsPath.Row <> vsPath.Rows - 1 Then
                lng��ĿID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '·������ĿΪ0
                Control.Enabled = lng��ĿID = 0
            Else
                Control.Enabled = False
            End If
        End If

    Case conMenu_Edit_View      '�鿴��Ŀ����
        Control.Enabled = blnHaveItem
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" Then
                lng��ĿID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '·������ĿΪ0
                Control.Enabled = lng��ĿID <> 0
            End If
        End If


        '3.ִ��
        '-----------------------------------------
    Case conMenu_Edit_Archive, conMenu_Edit_UnArchive   '������Ŀִ��(�����һ�ε��в���) 'ȡ��ִ��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Row >= .FixedRows Then
                    Control.Enabled = (.ColData(.Col) = mPP.��ǰ�׶�ID And mint���� = 0 And .Col = .Cols - 1) Or mint���� = 1
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Merge, conMenu_Edit_DeleteParent    '����ִ��,����ȡ��ִ��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1


        '4.����
        '-----------------------------------------
    Case conMenu_Edit_Audit     '�׶�����
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1 And mint���� = 0
    Case conMenu_Edit_Reuse     '�޸�����
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1 And mint���� = 0
    Case conMenu_Edit_Clear     'ȡ������
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1 And mint���� = 0


        '5.���
        '-----------------------------------------
    Case conMenu_Edit_Stop      '���·��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then    '��ǰ���������׼סԺ�շ�Χ�����������������
            Control.Enabled = mblnInOverScope And vsPath.TextMatrix(vsPath.Rows - 1, vsPath.Cols - 1) <> ""
        End If

    Case conMenu_Edit_ClearUp   'ȡ�����
        If mPP.����·��״̬ = 3 Then
            Control.Caption = "ȡ���˳�"
        Else
            Control.Caption = "ȡ�����"
        End If
        Control.Enabled = (mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3) And cboPath.ListIndex <= 0    '2-������ɣ�3-�������

    Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView   '�����ǼǱ�
        Control.Enabled = (mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3)     '2-������ɣ�3-�������
        If Control.ID = conMenu_Edit_OutLogModi And Control.Enabled Then
            Control.Enabled = mlng����״̬ = 0  '�ύ��˺�Ͳ������޸�
        End If

        '6.����
        '-----------------------------------------
    Case conMenu_Edit_Compend    '�鿴����
        With vsPath
            Control.Enabled = blnHaveItem
            If Control.Enabled Then Control.Enabled = .Cell(flexcpData, .Row, .Col) <> ""
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As ADODB.Recordset, str�������� As String
    Dim blnDo As Boolean
    Dim strTmp As String
    
    If mlngӤ������ID <> 0 Then
        If mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID Then
            MsgBox "�ò����Ѿ�ת���������ˣ�ֻ��Ӥ�����ڱ����ң����������·����", vbInformation, Me.Caption
            Exit Sub
        End If
    End If

    Select Case Control.ID
        '0.���
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3)
    Case conMenu_File_Print_PatiPath
        '��ӡ���߰�·����
        Call FuncPatiPathPrint
        '1.����
        '-----------------------------------------
    Case conMenu_Edit_Import    '����·��
        Call FuncImport
    Case conMenu_Edit_Untread   'ȡ������
        Call FuncUnImport
    Case conMenu_Edit_ImportMerge  '����ϲ�·��
        Call FuncImportMerge
    Case conMenu_Edit_UnImportMerge  'ȡ������ϲ�·��
        Call FuncUnImportMerge
    Case conMenu_Edit_Select      '�鿴��������
        Call frmEvaluate.ShowMe(mfrmParent, 0, 0, mPati, mPP)
    Case conMenu_Edit_ViewMergeImport      '�鿴�ϲ�·����������
        Call ViewMergeImport

        '2.����
        '-----------------------------------------
    Case conMenu_Edit_Send      '����·��
        Call FuncSendItem
    Case conMenu_Edit_Append    '��������
        Call FuncSendItemApend
    Case conMenu_Edit_Delete    'ȡ�������ɵ���Ŀ
        Call FuncDelItem

    Case conMenu_Edit_Blankoff  'ȡ����������
        Call FuncDelAllItem
    Case conMenu_Edit_SendBack  '��������ҽ��
        Call FuncReSendItem
    Case conMenu_Edit_Surplus   '���·������Ŀ
        Call FuncAppendItem(0)
    Case conMenu_Edit_Modify    '�޸�·������Ŀ
        Call FuncAppendItemModify

        '3.ִ��
        '-----------------------------------------
    Case conMenu_Edit_Archive   'ִ��·��
        Call FuncExecuteItem
    Case conMenu_Edit_Merge     '����ִ��
        Call FuncExecuteAll
    Case conMenu_Edit_UnArchive     'ȡ��ִ��
        Call FuncExecuteItemCancel
    Case conMenu_Edit_DeleteParent  '����ȡ��ִ��
        Call FuncExecuteAllCancel

        '4.����
        '-----------------------------------------
    Case conMenu_Edit_Audit     '����
        Call FuncEvaluate
    Case conMenu_Edit_Reuse     '�޸�����
        Call FuncReEvaluate
    Case conMenu_Edit_Clear     'ȡ������
        Call FuncEvaluateCancel


        '5.���
        '-----------------------------------------
    Case conMenu_Edit_Stop      '���·��
        Call FuncOver
    Case conMenu_Edit_ClearUp   'ȡ�����
        Call FuncOverCancel
    Case conMenu_Edit_OutLogModi    '�޸ĳ����ǼǱ�
        Call OutLogModi
    Case conMenu_Edit_OutLogView   '�鿴�����ǼǱ�
        Call frmPathOutLog.ShowMe(mfrmParent, mPati.����ID, mPati.��ҳID, 1, Nothing, mPP.·��ID, mPP.����·��ID)
        '6.���ƣ�����
        '-----------------------------------------
    Case conMenu_Edit_Up    '1-����
        Call MovePathItem(1)
    Case conMenu_Edit_Down    '-1-����
        Call MovePathItem(-1)

        '7.����
        '-----------------------------------------
    Case conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 10  '�������10������
        If InStr(Control.Parameter, ":") > 0 Then
            Call FuncViewReport(Split(Control.Parameter, ":")(0), Split(Control.Parameter, ":")(1))
        End If
    Case conMenu_Edit_MarkMap

        'Case conMenu_Manage_ReportView  '�鿴������

    Case conMenu_Edit_View    '��ʾ·����Ŀ�������Ϣ
        Call vsPath_DblClick
    Case conMenu_View_StPath    '�鿴��׼·���ο�
        Set rsTmp = GetPatiDiagnose(mPati.����ID, mPati.��ҳID, 2)  '��ȡ��Ҫ���
        If rsTmp.RecordCount <> 0 Then
            str�������� = rsTmp!����
        End If
        Call frmStPathList.ShowMe(mfrmParent, str��������)
'    Case conMenu_Tool_Option    '·��ѡ��
'        Dim objControl As CommandBarControl
'
'        If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��������;") = 0 Then
'            MsgBox "��û�в������õ�Ȩ�ޡ�", vbInformation, gstrSysName
'        Else
'            frmPathSetup.mbytFun = 0
'            frmPathSetup.Show 1, mfrmParent
'
'            strTmp = zlDatabase.GetPara("·��ִ�л������ó���", glngSys, p�ٴ�·��Ӧ��, "11")
'            If strTmp <> mstrִ�г��� Then
'                mstrִ�г��� = strTmp
'                blnDo = True
'            End If
'
'            If mbln����ִ�л��� <> CBool(Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, p�ٴ�·��Ӧ��, 0))) Or blnDo Then
'                mbln����ִ�л��� = Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, p�ٴ�·��Ӧ��, 1))
'                If mblnInsideTools Then
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_Archive, , True): objControl.Category = ""
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_Merge, , True): objControl.Category = ""
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_UnArchive, , True): objControl.Category = ""
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_DeleteParent, , True): objControl.Category = ""
'                Else
'                    Set objControl = mcbsMain.ActiveMenuBar.FindControl(, conMenu_Edit_Merge, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_Archive, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_Merge, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_UnArchive, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_DeleteParent, , True): objControl.Category = ""
'                End If
'            End If
'            End If
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
            If CreatePlugInOK(P�ٴ�·��Ӧ��) Then
                Call gobjPlugIn.ExecuteFunc(glngSys, P�ٴ�·��Ӧ��, Control.Parameter, mPati.����ID, mPati.��ҳID, mPP.·��ID, , mint����)
            End If
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
'���ܣ�����˵���ĵ����˵�
    Dim objControl As CommandBarControl
    Dim rsTmp As ADODB.Recordset, i As Long, j As Long
    Dim rsTmpPacs As Recordset
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_Edit_Compend
             With CommandBar.Controls
                .DeleteAll
                With vsPath
                    Set rsTmp = GetReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                    Set rsTmpPacs = GetPACSReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                End With
                
                If rsTmp.RecordCount = 0 And rsTmpPacs.RecordCount = 0 Then
                     .Add xtpControlButton, conMenu_Edit_Compend * 10 + 1, "�ޱ����δ��д"
                Else
                    For i = 1 To rsTmp.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i, rsTmp!�������� & "(&" & i & ")")
                        objControl.Parameter = rsTmp!ID & ":" & rsTmp!ҽ��id
                        rsTmp.MoveNext
                    Next
                    i = rsTmp.RecordCount
                    For j = 1 To rsTmpPacs.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i + j, rsTmpPacs!�ĵ����� & "(&" & i + j & ")")
                        objControl.Parameter = rsTmpPacs!����ID & ":" & rsTmpPacs!ҽ��id
                        rsTmpPacs.MoveNext
                    Next
                End If
                
                
            End With
    End Select
End Sub

Private Function GetReportOfPath(ByVal lng·��ִ��ID As Long) As ADODB.Recordset
'���ܣ���ȡ·����Ӧ�ı�������
    Dim strSql As String
 
    strSql = "Select d.id, d.��������,c.ҽ��Id" & vbNewLine & _
            "From ����·��ִ�� A, ����·��ҽ�� B, ����ҽ������ C, ���Ӳ�����¼ D" & vbNewLine & _
            "Where a.Id = [1] And a.Id = b.·��ִ��id And b.����ҽ��id = c.ҽ��Id And c.����id = d.Id"
    On Error GoTo errH
    Set GetReportOfPath = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ִ��ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPACSReportOfPath(ByVal lng·��ִ��ID As Long) As ADODB.Recordset
'���ܣ���ȡ·����Ӧ�ı�������
    Dim strSql As String
    Dim strIDs As String
    Dim rsTmp As Recordset
 
    strSql = "Select b.����ҽ��id" & vbNewLine & _
            "From ����·��ִ�� A, ����·��ҽ�� B" & vbNewLine & _
            "Where a.Id = [1] And a.Id = b.·��ִ��id "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ִ��ID)
    Do While Not rsTmp.EOF
        strIDs = strIDs & "," & rsTmp!����ҽ��id
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If strIDs <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Set GetPACSReportOfPath = mobjPublicPACS.zlDocGetListWithAdvice(strIDs)
    Else
        Set rsTmp = New Recordset
        rsTmp.Fields.Append "ID", adInteger, 1
        rsTmp.CursorLocation = adUseClient
        rsTmp.LockType = adLockOptimistic
        rsTmp.CursorType = adOpenStatic
        rsTmp.Open
        Set GetPACSReportOfPath = rsTmp
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.����)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS��������δ�����ɹ���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Private Sub FuncOver()
'���ܣ����·��
    Dim strSql As String, blnOK As Boolean, lngValue As Long
    Dim colSQL As New Collection, blnTrans As Boolean, i As Long
    Dim str����� As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long
    
    On Error GoTo errH
    '����黤ʿ���ɵ���Ŀ�Ƿ����δִ�еǼǵ���Ŀ(��Ϊ��ʿû����������)
    If mbln����ִ�л��� Then
        If Mid(mstrִ�г���, 2, 1) = "1" Then  '��ʿ��������ִ�л���
            strSql = "Select a.ִ��ʱ��" & vbNewLine & _
                    "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
                    "Where a.��Ŀid = b.Id(+) And a.·����¼id = [1] and NVl(NVl(a.ִ����,b.ִ����),1)=2 and a.ִ��ʱ�� is null and rownum <2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
            If rsTmp.RecordCount > 0 Then
                MsgBox "���ڻ�ʿδ���ִ�еǼǵ�·����Ŀ���������ִ�еǼǺ������·����"
                Exit Sub
            End If
        End If
    End If
    
    '���жϸ�·���Ƿ�������ϲ�ͬ���·���������������Ժ����Ƿ�͵��������ͬ
    If mPP.����·������ = 0 Then
        If Not CheckPathOutDiag(mPP.����·��ID, mPati.����ID, mPati.��ҳID) Then
            MsgBox "��Ժ��ϲ������ò��ַ�Χ�ڣ��������������·����ֻ�ܱ����˳�·����", vbInformation, gstrSysName
            Call FuncReEvaluate
            Exit Sub
        End If
    End If
    '�ж��Ƿ���δ��˵Ľ׶�
    If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";��ת���;") = 0 Then
        If CheckPathIsTurnAduit Then
            str����� = zlDatabase.UserIdentify(Me, "ǰ��׶δ���δ��˵�·����ת��������˺��������ɡ�", glngSys, P�ٴ�·��Ӧ��, "��ת���")
            If str����� = "" Then Exit Sub
        End If
    Else
        str����� = UserInfo.����
    End If
            
    If MsgBox("��ȷ��Ҫ��ɵ�ǰ���˵��ٴ�·����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If
    If CheckPathOutLog Then
        blnOK = frmPathOutLog.ShowMe(mfrmParent, mPati.����ID, mPati.��ҳID, 0, colSQL, mPP.·��ID, mPP.����·��ID)
        If blnOK = False Then
            lngValue = Val(zlDatabase.GetPara("������д�����ǼǱ�", glngSys, P�ٴ�·��Ӧ��, "0"))
            If lngValue = 1 Then
                MsgBox "�������·��ǰ������д�����ǼǱ���ȡ������д��·����ɲ���δִ�С�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    lngPPStatus = mPP.����·��״̬
    
    strSql = "Zl_����·������_Update(" & mPP.����·��ID & ",'" & str����� & "')"
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            'ִ�г����ǼǱ��SQL
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "�����ǼǱ�")
        Next
        Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·�����")
    gcnOracle.CommitTrans: blnTrans = False
    
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    
    '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
    If lngPPStatus <> mPP.����·��״̬ Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncOverCancel()
'���ܣ�ȡ��·�������
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long
    
    On Error GoTo errH
    '��ɺ�û���������õ�ҽ��������ȡ��
    strSql = "Select Null" & vbNewLine & _
            "From �����ٴ�·�� A, ����ҽ����¼ B" & vbNewLine & _
            "Where a.Id = [1] And a.����id = b.����id And a.��ҳid = b.��ҳid And b.����ʱ�� > Trunc(a.����ʱ��, 'MI') And b.ҽ��״̬ Not In (-1, 4) And Nvl(b.Ӥ��,0)=0 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "·����ɺ��Ѳ������µ�ҽ������ɾ�������Ϻ��ٽ���ȡ��������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mPP.����·��״̬ = 3 Then
        If MsgBox("��ǰ·���Ǳ�����Զ���ɵģ�ȡ����������������ͬʱɾ��������ȡ����������ȷ��Ҫ������?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    lngPPStatus = mPP.����·��״̬
    
    strSql = "Zl_����·������_Delete(" & mPP.����·��ID & "," & mPP.����·��״̬ & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·�����")
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
    If lngPPStatus <> mPP.����·��״̬ Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItem()
'���ܣ�ִ��·����Ŀ
    Dim lngִ��ID As Long, lng��ĿID As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng��ĿID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)    '·������ĿΪ0
    End With
        
    
    strSql = "Select 1 From ����·��ִ�� Where ID = [1] And ִ��ʱ�� is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����Ŀ��ִ�С�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ݼ��
    If lng��ĿID <> 0 Then
        strSql = "Select ִ���� From �ٴ�·����Ŀ Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��ĿID)
    Else
        strSql = "Select ִ���� From ����·��ִ�� Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    End If
    '���ݳ����ж�,��������Ա����
    If (mint���� = 0 And rsTmp!ִ���� = 2) Or (mint���� = 1 And rsTmp!ִ���� = 1) Then
        MsgBox "����Ŀֻ����" & IIf(rsTmp!ִ���� = 1, "ҽ��", "��ʿ") & "ִ�С�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frmPathExecute.ShowMe(mfrmParent, 1, mPati, mPP, lngִ��ID, mint����) Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteAll()
'���ܣ�����ִ��
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If mbln���ò����� Then
'        strSQL = "Select * From (" & _
'            "Select Distinct a.�׶�id, a.����, a.����, a.�Ǽ�ʱ��" & vbNewLine & _
'            "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
'            "Where a.��Ŀid = b.Id(+) And a.·����¼id = [1] And Nvl(a.����ʱ������, 0) = 0 And Nvl(Nvl(a.ִ����, b.ִ����), 0) = " & IIf(mint���� = 0, 1, 2) & " And" & vbNewLine & _
'            "      a.ִ��ʱ�� Is Null " & vbNewLine & _
'            "Order By a.�Ǽ�ʱ��) where Rownum <2 "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPP.����·��ID)
'        If rsTmp.RecordCount > 0 Then
'            mPP.��ǰ�׶�ID = rsTmp!�׶�ID
'            mPP.��ǰ���� = rsTmp!����
'            mPP.��ǰ���� = rsTmp!����
'        End If
        GetPathCurrPhase 1, mPP.��ǰ�׶�ID, mPP.��ǰ����, mPP.��ǰ����
    End If
    
'    If mint���� = 1 Then
'        Call GetPhaseInNurse(0, mPP.��ǰ�׶�ID, mPP.��ǰ����, mPP.��ǰ����)
'    End If
    
    If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, mint����) Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItemCancel()
'���ܣ�ȡ��·����Ŀ��ִ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lngִ��ID As Long
    Dim blnTip As Boolean
    
    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With
        
    strSql = "Select 1 From ����·��ִ�� Where ID = [1] And ִ��ʱ�� is Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����Ŀ��δִ�С�", vbInformation, gstrSysName
        Exit Sub
    End If
    'ҽ��ִ�е���Ŀֻ��ҽ��ȡ��,��ʿִ�е�ֻ�ܻ�ʿȡ��
    strSql = "Select 1 From ����·��ִ�� A,�ٴ�·����Ŀ B Where A.ID=[1] And A.��ĿID=B.ID(+) " & _
            "And NVL(NVL(A.ִ����,B.ִ����),1)=" & IIf(mint���� = 0, 1, 2) & " And A.ִ��ʱ�� is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount = 0 Then
        MsgBox "����Ŀ��" & IIf(mint���� = 0, "��ʿ", "ҽ��") & "ִ�еǼ�,������ȡ����", vbInformation, gstrSysName
        Exit Sub
    End If

    strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, Val(vsPath.ColData(vsPath.Col)), CDate(vsPath.Cell(flexcpData, EFixedRow.R2����, vsPath.Col)))
    If rsTmp.RecordCount > 0 Then
        'ǿ��ȡ�������������Ȩ��
        If mint���� = 0 Then
            If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ�����������ȡ��ִ�С�" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        ElseIf mint���� = 1 Then
            '���·��������
            If CheckPathSendByNurse(2, lngִ��ID) Then
                blnTip = True
            Else
                MsgBox "����Ŀ��ҽ�����ɵ���Ŀ��" & mPP.��ǰ���� & "�ѽ�����������" & vbCrLf & vbCrLf & "������ȡ��ִ�С�", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
    Else
        blnTip = True
    End If
    
    If blnTip Then
        If MsgBox("��ȷ��Ҫȡ��[" & vsPath.TextMatrix(vsPath.Row, vsPath.Col) & "]��ִ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    strSql = "Zl_����·��ִ��_Delete(" & lngִ��ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncExecuteAllCancel(Optional blnRefresh As Boolean = True) As Boolean
'���ܣ�����ȡ��·����Ŀ��ִ��
'˵������ʿ���ɵ���Ŀ���Բ�����������ڡ���ҽ��վ������ʱ�����ҽ�������ߵ�ִ�еǼ������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnDo As Boolean
    Dim blnNurse As Boolean    'Ture -���ڻ�ʿ���ɵ���Ŀ ;False-�������ڻ�ʿ����

    On Error GoTo errH
    
    
    '�Ƿ�����Ѿ�ִ�еǼǵ���Ŀ,ȡ����������ʱ,���ִ�еǼǳ���ֻ���û�ʿʱ,ҽ��ǿ��ȡ����������ʱ,����Ϊ��ǰ������ҽ��ִ�еǼǵ���Ŀ����ֹ�˳�
    If blnRefresh = True Then
        If mbln���ò����� Then
            GetPathCurrPhase 2, mPP.��ǰ�׶�ID, mPP.��ǰ����, mPP.��ǰ����
        End If
        strSql = "Select 1 From ����·��ִ�� A,�ٴ�·����Ŀ B Where A.·����¼ID = [1] And A.�׶�ID = [2] And A.���� = [3] And A.��ĿID=B.ID(+) " & _
                "And NVL(NVL(A.ִ����,B.ִ����),1)=" & IIf(mint���� = 0, 1, 2) & " And A.ִ��ʱ�� is Not Null And Rownum<2"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ǰ��������" & IIf(mint���� = 0, "ҽ��", "��ʿ") & "ִ�еǼǵ��κ���Ŀ��", vbInformation, gstrSysName
            FuncExecuteAllCancel = True
            Exit Function
        End If
    End If
    
    '�������ڼ��
    strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        If mint���� = 1 Then
            strSql = "Select 1 " & vbNewLine & _
                    "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
                    "Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And a.��Ŀid = b.Id(+) And Nvl(Nvl(a.������, b.������), 1) = 2 and rownum<2 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
            If CheckPathSendByNurse(1, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����)) Then
                '���ڻ�ʿ���ɵ�·����Ŀ
                blnNurse = True
            Else
                MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ���������,������ȡ����ҽ�����ɵ���Ŀ��", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        ElseIf mint���� = 0 Then
            'ǿ��ȡ�������������Ȩ��
            If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ�����������ȡ��ִ�С�" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, True)
            Else
                Exit Function
            End If
        End If
    End If
 
    blnDo = frmPathExecute.ShowMe(mfrmParent, 2, mPati, mPP, 0, mint����, blnNurse)
    If blnDo And blnRefresh Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    
    FuncExecuteAllCancel = blnDo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function CheckSameDayOfPhaseTurn() As Boolean
'���ܣ���鵱ǰ·���Ƿ�ո���ת�������鵱���Ƿ��п��õĽ׶�
    Dim strSql As String, rsTmp As Recordset
    
    strSql = "select ԭ·��ID,ԭ·���汾 from ����·������  where ·����¼id=[1] and �׶�ID=[2] and ����=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    If rsTmp.RecordCount > 0 Then
        If rsTmp!ԭ·��ID & "" <> "" Then
            strSql = "Select 1 From �ٴ�·���׶� Where ·��ID=[1] And �汾��=[2] and [3] Between ��ʼ���� And Nvl(��������, ��ʼ����)  And rownum<2 and ��֧ID is null And ��ID is null"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID, mPP.�汾��, mPP.��ǰ����)
            CheckSameDayOfPhaseTurn = rsTmp.RecordCount > 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FuncSendItem(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'���ܣ�ִ������·��(ͨ��clsDockPath�еĽӿڿ��Ÿ�ҽ����������)
'������blnIsCancel��û��·��������ʱ���û��Ƿ�ȡ����������true=ȡ��
'     lngType:1-ҽ���༭������ã��������󲻼������ɣ���Ϊҽ���༭���治���ٵ���ҽ���༭��
    Dim rsTmp As ADODB.Recordset
    '-------
    Dim lng���� As Long, lngʱ����� As Long, lng�������� As Long
    Dim lng�׶�ID As Long
    Dim lngPPStatus As Long
    Dim i As Long
    '-------
    Dim strTmp As String
    Dim strSql As String
    Dim strDate As String
    Dim strPhase As String
    Dim strMsg As String
    '-------
    Dim blnDo As Boolean
    Dim blnIsNext As Boolean
    Dim blnEvaluate As Boolean
    Dim blnRefresh As Boolean
    Dim blnTrans As Boolean
    
    Dim DatCurr As Date
    Dim colSQL As Collection
    
    On Error GoTo errH
    
LineBegin:
    If mint���� = 1 Then
        '��ʿ��������ǰǿ��ˢ�£�����ҽ�����ɻ�ȡ�����ɲ���ʱ����ʿվδ��ͬ�����µ����
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If

    If mPP.��ǰ���� = 0 Then '��һ��
        '��ʿ����
        If mint���� = 1 Then
            MsgBox "ҽ����û�������κ�·����Ŀ,��ʿ������ǰ���ɡ�", vbInformation, gstrSysName
            Exit Function
        End If
            
        strSql = "Select To_number(Trunc(Sysdate)-Trunc(a.��ʼʱ��)+1) as ��Ժ����,Nvl(b.ȷ������,0) as ȷ������,a.��ʼʱ�� as ��Ժʱ��" & _
                " From �����ٴ�·�� a,�ٴ�·��Ŀ¼ b Where a.ID = [1] And a.·��id = b.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        DatCurr = zlDatabase.Currentdate
        If rsTmp!ȷ������ > 0 And DatCurr > Format(DateAdd("d", Val(rsTmp!ȷ������), rsTmp!��Ժʱ��), "yyyy-MM-DD HH:mm:ss") Then
            MsgBox "�ò�������Ժ" & rsTmp!��Ժ���� & "�죬�����˹涨��ȷ������(" & rsTmp!ȷ������ & "��)������������·����", vbInformation, gstrSysName
            Exit Function
        End If
        If mPP.����ʱ�� <> CDate(0) Then
            '������·�����״��������������ڴ��ڵ�������
            DatCurr = zlDatabase.Currentdate
            If Int(DatCurr) - Int(mPP.����ʱ��) >= 1 Then
                Set colSQL = New Collection
                Call CreatePathItem(DatCurr, mPP.����ʱ��, mPati, mPP, mPP.����·��ID, colSQL)
                If colSQL.count > 0 Then
                    gcnOracle.BeginTrans: blnTrans = True
                    For i = 1 To colSQL.count
                        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "·������")
                    Next
                    gcnOracle.CommitTrans: blnTrans = False
                    'ǿ��ˢ��
                    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
                    GoTo LineBegin
                End If
            End If
        End If
        lng���� = rsTmp!��Ժ����
        lngʱ����� = 0
    Else
        If mint���� = 1 Then
            '��ʿ��������·����Ŀ���
            '1)��ǰ·��ҽ��δ�����κ�·����Ŀ,��ʿ����������
            '2)ҽ����û��������һ�׶�,��ʿû��·����Ŀ��������
            '��ȡ��ʿ�������׶μ�����
            strSql = "Select 1 from ����·��ִ�� A where A.·����¼ID=[1] And A.�׶�ID=[2] And A.����=[3] And NVL(a.������,1) =2 And RowNum<2 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����·����Ŀ", mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
            
            If rsTmp.RecordCount > 0 Then
                MsgBox "�ò����ڵ����·�������ɡ�", vbInformation, gstrSysName
                Exit Function
            Else
                Call GetPhaseInNurse(1, mPP.��ǰ�׶�ID, mPP.��ǰ����, mPP.��ǰ����, , lng����, strPhase)
            End If
            lngʱ����� = 0
            
            If Not CheckPathIsExecuted(blnRefresh) Then
                'ǿ��ˢ��
                If blnRefresh Then
                    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
                End If
                Exit Function
            End If
        End If
        
        If mint���� = 0 Then
            '2.��ǰδ�����������������µ�;��ʿ����û����������
            strSql = "Select ʱ����� From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
            If Not mbln���ò����� Then
                If rsTmp.RecordCount = 0 Then
                    If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";�׶�����;") = 0 Then
                        MsgBox "�ò�����" & mPP.��ǰ���� & "��û�н������������ܽ��к���������", vbInformation, gstrSysName
                        Exit Function
                    Else
                        If MsgBox("�ò�����" & mPP.��ǰ���� & "��û�н���������������������" & vbCrLf & "������Ҫ��������������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                            '����ǰ���ȼ��ִ�еǼ����
                            If Not CheckPathIsExecuted() Then
                                Exit Function
                            End If
                            '
                            If frmEvaluate.ShowMe(mfrmParent, 1, 1, mPati, mPP) = False Then
                                Exit Function
                            Else
                                lngPPStatus = mPP.����·��״̬
                                Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
                                '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
                                If lngPPStatus <> mPP.����·��״̬ Then
                                    If Not gobjLIS Is Nothing Then
                                       Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
                                    End If
                                End If
                                '�����󣬿��ܽ������˳�·�������Ը��������е�״̬�����ж��Ƿ�Ҫ��������,�˳�������򲻼�������
                                If mPP.����·��״̬ <> 1 Or lngType = 1 Then
                                    Exit Function
                                End If

                                Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
                                strSql = "Select ʱ����� From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
                                If rsTmp.RecordCount <> 0 Then
                                    lngʱ����� = Val("" & rsTmp!ʱ�����): blnEvaluate = True
                                Else
                                    Exit Function
                                End If
                                
                                blnIsNext = True
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    lngʱ����� = Val("" & rsTmp!ʱ�����): blnEvaluate = True
                End If
            Else
                '��δ���ñ�������ʱ,����û��Ѿ�������,�Ͱ�������ʱ����ȣ���ǰ/�Ӻ�/������������һ�׶�;��δ��������ȱʡʱ����� 0-����
                If rsTmp.RecordCount <> 0 Then
                    lngʱ����� = Val("" & rsTmp!ʱ�����): blnEvaluate = True
                End If
            End If
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            If lngʱ����� = 0 Then
                If mPP.��ǰ���� = strDate Then
                    lng�������� = GetMustDay(mPP.����·��ID, mPP.��ǰ����)
                    'a.������컹�������׶Σ��������������׶Σ����������ǵ���
                    If CheckSameDayOfPhase(mPP.��ǰ�׶�ID, lng��������) Then
                        lng���� = mPP.��ǰ����
                    Else
                        '��鵱ǰ·���Ƿ�ո���ת�������鵱���Ƿ��п��õĽ׶�
                        If CheckSameDayOfPhaseTurn Then
                            lng���� = mPP.��ǰ����
                        Else
                            blnDo = False
                            If mbln������ǰ���� Then
                                'c.��ǰ���ɺ����׶�
                                If MsgBox("�ò����ڵ����·����Ŀ�����ɣ�������Ҫ��ǰ������һ���·����Ŀ��", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then
                                    If CheckSendOfBefore() Then
                                        lng���� = mPP.��ǰ���� + 1: blnDo = True
                                    Else
                                        Exit Function
                                    End If
                                Else
                                    Exit Function
                                End If
                            Else
                                MsgBox "�ò��˵���û���������õĽ׶ο������ɡ�", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                                Exit Function
                            End If
                            
                            If Not blnDo And blnEvaluate Then
                                '���û�к����׶��ˣ��û����Ǹո���������ֱ���˳�
                                If blnIsNext Then Exit Function
                                If MsgBox("�ò����ڽ����ѽ�����������Ҫ����ҽ������ȡ��������������Ҫȡ��������", vbYesNo + vbDefaultButton1 + vbQuestion, "�Ƿ�ȡ��������") = vbYes Then
                                    Call FuncEvaluateCancel(False, True)
                                    blnIsCancel = True
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf mPP.��ǰ���� < strDate Then
                    'b.֮ǰ������û�����ɣ��򲹳�����
                    lng���� = mPP.��ǰ���� + 1
                Else 'c.��ǰ���ɺ����׶�
                    If mbln������ǰ���� Then
                        If CheckSendOfBefore() Then
                            lng���� = mPP.��ǰ���� + 1: blnDo = True
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            ElseIf lngʱ����� = 1 Then '��һ�׶���ǰ������(ʱ�䲻�䣬ͬһ�����ɶ���׶ε�����)
                lng���� = mPP.��ǰ����
            ElseIf lngʱ����� = 2 Then '��һ�׶���ǰ������
                If mPP.��ǰ���� = strDate Then
                    MsgBox "��һ�׶�����Ϊ����һ�׶���ǰ�����족,��������������һ�׶ε�·����Ŀ��", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                lng���� = mPP.��ǰ���� + 1
            Else    '��һ�׶��Ӻ�(������ǰ�׶�)
                If mPP.��ǰ���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
                    MsgBox "�ò����ڽ����·�������ɡ�", vbInformation, gstrSysName
                    Exit Function
                End If
                lng���� = mPP.��ǰ���� + 1
            End If
        End If
    End If

    If frmPathSend.ShowMe(mfrmParent, 0, mint����, mPati, mPP, mPP.��ǰ�׶�ID, lng����, 0, 0, lngʱ�����, mclsMipModule, blnDo, strPhase) Then
        FuncSendItem = True
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncSendItemApend()
'���ܣ���������·��
'      �����·��������ʱ��������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strTmp As String
    Dim strDate As String
    
    On Error GoTo errH
    If mint���� = 1 Then
        '��ʿ���ϸ���ѡ��׶β�������
        Call GetPhaseInNurse(0, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    End If
    
    strSql = "Select Max(ID) as ID From ����·��ִ�� Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    If IsNull(rsTmp!ID) Then
        MsgBox "�ò����ڽ����·����û�����ɡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint���� = 0 Then 'ҽ���Ŷ������������жϣ���ʿ������������
        strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";�׶�����;") = 0 Then
                MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ����������������ٲ���������Ŀ��", vbInformation, gstrSysName
                Exit Sub
            Else
                'ȡ������
                If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ����������ܲ���������Ŀ��" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    Call FuncEvaluateCancel(False, False)
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
    If frmPathSend.ShowMe(mfrmParent, 1, mint����, mPati, mPP, mPP.��ǰ�׶�ID, mPP.��ǰ����, , , , mclsMipModule) Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncReSendItem()
'���ܣ���������·����Ŀ��ҽ��
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngִ��ID As Long, lng��ĿID As Long, blnMust As Boolean, lng���� As Long
            
    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng��ĿID = Val(Split(.Cell(flexcpData, .Row, .Col), "|")(1))
    End With
    If lng��ĿID = 0 Then
        MsgBox "Ҫ��������·������Ŀ����ȡ������Ŀ�����ɺ�������ӡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '1.�Ѿ�ִ�еĲ�������������
    strSql = "Select a.ִ��ʱ��,c.����Ҫ��  From ����·��ִ�� a,�ٴ�·��ҽ�� b,�ٴ�·����Ŀ C Where a.ID = [1] And a.��Ŀid = b.·����Ŀid And a.��ĿID = c.ID And rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!ִ��ʱ��) And mbln����ִ�л��� Then
            If rsTmp.RecordCount > 0 Then
                MsgBox "����Ŀ��ִ�У������������ɡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If Val("" & rsTmp!����Ҫ��) = 0 Then
            strSql = "Select 1" & vbNewLine & _
                    "From ����·��ҽ�� A, ����·��ҽ�� B" & vbNewLine & _
                    "Where a.·��ִ��id = [1] And a.����ҽ��id = b.����ҽ��id And b.·��ִ��id <> a.·��ִ��id  And rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
            If rsTmp.RecordCount > 0 Then
                MsgBox "����Ŀ��Ӧ��ҽ���Ǹ����ϴεĳ������ɵģ������ǿ���ѡ���ɵģ�����ִ���������ɡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        MsgBox "����Ŀ����ҽ������Ŀ�������������ɡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '2.���ҽ��
    If mint���� = 1 Then
        '�����Ѿ�����˵�ҽ�����������޸�ɾ����
        strSql = "Select 1 From ����·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id = [1] And b.����ҽ��id = c.Id And c.����ҽ�� Like '%/%' And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "����Ŀ��Ӧ��ҽ���Ѿ���ҽ����ˣ�����ִ�д˲�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If frmPathSend.ShowMe(mfrmParent, 3, mint����, mPati, mPP, mPP.��ǰ�׶�ID, mPP.��ǰ����, lng��ĿID, lngִ��ID, , mclsMipModule) Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDelPhaseItem()
'���ܣ�ǿ��ɾ�����һ�����е�ִ����Ŀ(���ڲ���ʱ�������)
    Dim strSql As String
    Dim lngִ��ID As Long
    Dim i As Long
        
    On Error GoTo errH
    With vsPath
        For i = .FixedRows To .Rows - 2     '���һ��������
            If .TextMatrix(i, .Cols - 1) <> "" Then
                lngִ��ID = Split(.Cell(flexcpData, i, .Cols - 1), "|")(0)
                strSql = "Zl_����·������_Delete(" & lngִ��ID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
            End If
        Next
    End With
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    Exit Sub
 Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncDelAllItem(Optional ByVal blnRefresh As Boolean = True, Optional ByVal blnPrompt As Boolean = True) As Boolean
'���ܣ�����ȡ���������ɵ�����·����Ŀ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, strIDs As String, strIDSQL As String, blnTrans As Boolean
    Dim strNewIDs As String
    Dim blnExecuted As Boolean
    Dim dat����ʱ�� As Date
    Dim lng���� As Long
    
    If blnPrompt Then
        If mint���� = 0 Then
            If MsgBox("ȡ�����ɽ�ɾ��·����Ŀ��Ӧ��ҽ���Ͳ����ļ���" & vbCrLf & "��ȷʵҪȡ���������ɵ�����·����Ŀ��?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If MsgBox("ȡ�����ɽ�ɾ��·�����л�������Ŀ��" & vbCrLf & "��ȷʵҪȡ���������ɵ����л�����Ŀ��?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    If mint���� = 1 Then
        Call GetPhaseInNurse(0, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    End If
    
    On Error GoTo errH
    
    strSql = "Select A.ID,A.ִ��ʱ��,NVL(NVL(A.ִ����,B.ִ����),1) as ִ����,NVL(NVl(A.������,B.������),1) as ������ From ����·��ִ�� A,�ٴ�·����Ŀ B Where A.·����¼ID = [1] And A.�׶�ID = [2] And A.���� = [3] and A.��ĿID=B.ID(+) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    If mint���� = 0 Then
'   ҽ��վȡ����������ʱ,�����ǻ�ʿ�Ƿ��Ѿ�������Ŀ��ҽ��������Ŀ�ɻ�ʿִ�еǼǵ��������ԭ�򣺽���ҽ��֮��Ĺ�����
'        rsTmp.Filter = "������ =2"
'        If rsTmp.RecordCount > 0 Then
'            MsgBox "��ʿ�����ɻ�������Ŀ������֪ͨ��ʿȡ�����ɡ�", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
'            Exit Function
'        End If
'        ҽ�����ɵ���Ŀ���ܻ�ʿ�Ƿ�Ǽ� ������ҽ��ǿ��ȡ��
'        rsTmp.Filter = "ִ����=2 and ������ =1"
'        If rsTmp.RecordCount > 0 Then
'            If Not IsNull(rsTmp!ִ��ʱ��) Then
'                MsgBox "����ҽ�����ɵ���Ŀ����ʿִ�еǼǣ�����֪ͨ��ʿȡ��ִ�еǼǡ�", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
'                Exit Function
'            End If
'        End If
    Else
        rsTmp.Filter = "������ = 2"
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ǰ�׶�û�����ɹ���������Ŀ��", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    
    rsTmp.Filter = IIf(mint���� = 0, "", "������ = 2") 'ҽ��վʱ
    
    Do While Not rsTmp.EOF
        If blnExecuted = False Then
            If Not IsNull(rsTmp!ִ��ʱ��) And ((Mid(mstrִ�г���, 1, 1) = "1" And mint���� = 0 And Val(rsTmp!ִ����) = 1) Or mint���� = 1) Then
                blnExecuted = True
            End If
        End If
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If blnExecuted Then
        '���ж�Ȩ�ޣ�����ʾ��ǿ��ȡ��
        If FuncExecuteAllCancel(False) = False Then
            Exit Function
        End If
    End If
    
    If mint���� = 0 Then
        strSql = "Select ����ʱ�� from �����ٴ�·�� Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        dat����ʱ�� = Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
        '����Ƿ�������
        If mbln����ִ�л��� = False Or Not blnExecuted Then
            strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
            If rsTmp.RecordCount > 0 Then
                'ǿ��ȡ�������������Ȩ��
                MsgBox "�������ɵ���Ŀ��������ȡ������֮ǰ���Զ�ȡ��������", vbInformation, gstrSysName
                Call FuncEvaluateCancel(False, False)
            End If
        End If
        strIDSQL = "(Select Column_value From Table(f_Str2List([1])))"
        '2.���ҽ��
        '���ǵ������ɵĳ���������ȡ��·����Ŀ�������Ƿ��ͣ�
        '�ǵ������ɵĳ�������У�Ե�δ���ϣ�������ȡ����δУ�Եģ�ȡ��ʱ�Զ�ɾ����Ӧ��ҽ����

        strSql = "Select /*+ Rule*/ distinct A.·��ִ��id" & vbNewLine & _
                 "From ����·��ҽ�� A, ����·��ҽ�� B" & vbNewLine & _
                 "Where a.·��ִ��id In " & strIDSQL & " And a.����ҽ��id = b.����ҽ��id And b.·��ִ��id <> a.·��ִ��id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs)
        If rsTmp.RecordCount = 0 Then
            strNewIDs = strIDs
            'û�зǵ��յĳ���
        Else
            '����ǰ�����˳������ǲ���ȥ����ֻ��鵱���
            strNewIDs = "," & strIDs & ","
            For i = 1 To rsTmp.RecordCount
                If InStr(strNewIDs, "," & rsTmp!·��ִ��id & ",") > 0 Then
                    strNewIDs = Replace(strNewIDs, "," & rsTmp!·��ִ��id & ",", ",")
                End If
                rsTmp.MoveNext
            Next
            If strNewIDs = "," Then
                strNewIDs = ""
            Else
                strNewIDs = Mid(strNewIDs, 2, Len(strNewIDs) - 2)
            End If
        End If
        
        If strNewIDs <> "" Then
            '��ʹ��ֹͣ��ҽ��Ҳ������ɾ������59������Ϊ����ʱ��δ��ȷ����
            strSql = "Select /*+ Rule*/ C.ҽ������ From ����·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id In " & strIDSQL & _
                     " And b.����ҽ��id = c.Id And c.ҽ��״̬ > 1 And c.ҽ��״̬ <> 4 And rownum<2 And to_date(to_char(c.����ʱ�� +59/24/60/60,'yyyy-mm-dd hh24:mi:ss'),'yyyy-mm-dd hh24:mi:ss') >[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewIDs, dat����ʱ��)
            If rsTmp.RecordCount > 0 Then
                strIDs = ""
                For i = 1 To rsTmp.RecordCount
                    If i > 10 Then strIDs = strIDs & "......": Exit For
                    strIDs = strIDs & vbNewLine & rsTmp!ҽ������
                    rsTmp.MoveNext
                Next
                MsgBox "��ǰ���ɵ���Ŀ������У�Ե�δ���ϵ�ҽ����" & strIDs & vbNewLine & "��������ҽ������ִ��ȡ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
'    If mint���� = 1 Then
'        '�����Ѿ�����˵�ҽ�����������޸�ɾ����
'        strSql = "Select /*+ Rule*/ 1 From ����·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id In " & strIDSQL & _
'                 " And b.����ҽ��id = c.Id And c.����ҽ�� Like '%/%' And rownum<2  And to_date(to_char(c.����ʱ�� +59/24/60/60,'yyyy-mm-dd hh24:mi:ss'),'yyyy-mm-dd hh24:mi:ss') >[2]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, dat����ʱ��)
'        If rsTmp.RecordCount > 0 Then
'            MsgBox "��ǰ���ɵ���Ŀ��Ӧ��ҽ���Ѿ���ҽ����ˣ���������ȡ����", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If

        '3.��鲡��
        strSql = "Select /*+ Rule*/ 1 From ���Ӳ�����¼ Where ·��ִ��id In " & strIDSQL & _
                 " And (���ʱ�� is not null or ��ӡ�� is not null) And rownum<2  And ����ʱ�� >[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, dat����ʱ��)
        If rsTmp.RecordCount > 0 Then
            MsgBox "��ǰ���ɵ���Ŀ��Ӧ�Ĳ�����ǩ�����Ѵ�ӡ����������ȡ����", vbInformation, gstrSysName
            Exit Function
        End If
        
        '����°���Ӳ���
        If Not CheckDelNewEMR(strIDs, 1, rsTmp) Then  '������Ҫɾ���ĵ��Ӳ�������ID
            Exit Function
        Else
            'ɾ��
            If Not gobjEmr Is Nothing Then
                If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
                If Not gobjEmr Is Nothing Then
                    For i = 1 To rsTmp.RecordCount
                        strSql = "<parameter><taskid>" & rsTmp!����ID & "</taskid></parameter>"
                        On Error Resume Next
                        Call gobjEmr.DeleteTask(strSql)
                        On Error GoTo 0
                        rsTmp.MoveNext
                    Next
                End If
            End If
        End If
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(Split(strIDs, ","))
        strSql = "Zl_����·������_Delete(" & Split(strIDs, ",")(i) & ",0," & mint���� & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    FuncDelAllItem = True

    If blnRefresh Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncDelItem()
'���ܣ�ȡ�����ɵ�ǰѡ���δִ�е�·����Ŀ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngִ��ID As Long, lng��ĿID As Long, blnMust As Boolean, lng���� As Long
    Dim blnCancel As Boolean, strReason As String, blnTrans As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long
    
    With vsPath

        If mint���� = 0 And Split(.Cell(flexcpData, .Row, .Col), "|")(3) = 2 Then '������ 1-ҽ��,2-��ʿ
            MsgBox "��ǰ��Ŀ�ǻ�ʿ���ɵ�,�㲻��ɾ����", vbInformation, Me.Caption
            Exit Sub
        ElseIf mint���� = 1 And Split(.Cell(flexcpData, .Row, .Col), "|")(3) = 1 Then
            MsgBox "��ǰ��Ŀ��ҽ�����ɵģ��㲻��ɾ����", vbInformation, Me.Caption
            Exit Sub
        End If
         
        If .Cell(flexcpBackColor, .Row, .Col) = &HE0EFED Then
            MsgBox "����ĿΪ�������ɵ�û�����ɵ���Ŀ������ȡ�����ɡ�", vbInformation, Me.Caption
            Exit Sub
        End If
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng��ĿID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
    End With
    
    If mbln����ִ�л��� Then
        '�Ѿ�ִ�еĲ�����ȡ��
        strSql = "Select 1 " & vbNewLine & _
                "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
                "Where a.��Ŀid = b.Id(+) And a.Id = [1] And Nvl(a.����ʱ������,0)<>1 And a.ִ��ʱ�� Is Not Null"

        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "����Ŀ��ִ�У�����ȡ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    
    '1.���·����Ŀ
    strSql = "Select b.ִ�з�ʽ,a.���� From ����·��ִ�� a, �ٴ�·����Ŀ b Where a.��ĿID = b.ID And a.ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then '��ʱ��Ŀ������ȡ��
        lng���� = Val("" & rsTmp!����)
        If rsTmp!ִ�з�ʽ = 1 Then
            blnMust = True
        ElseIf rsTmp!ִ�з�ʽ = 2 Or rsTmp!ִ�з�ʽ = 4 Then  '����һ�λ����һ��
            strSql = "Select ��ʼ����,�������� From �ٴ�·���׶� Where ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.��ǰ�׶�ID)
            If Not IsNull(rsTmp!��ʼ����) Then
                If Not IsNull(rsTmp!��������) Then
                    blnMust = (lng���� = Val("" & rsTmp!��������))    '�Ƿ����һ��
                    If blnMust Then '�жϸ���Ŀ֮ǰ��û��ִ�й�(·������Ŀ����)
                    
                        strSql = "Select 1 From ����·��ִ�� Where ·����¼ID = [1] And �׶�ID = [2] And ��ĿID = [3] And ����<[4] And rownum<2"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, lng��ĿID, lng����)
                        If rsTmp.RecordCount > 0 Then blnMust = False
                    End If
                Else
                    blnMust = True  '����
                End If
            End If
        End If
        
    End If
    
    '2.���ҽ��
    If CheckDelPathItem(lngִ��ID, mint����) = False Then Exit Sub
    '3.�������ɵ���Ŀ��д����ԭ��
    If blnMust Then
        'ȡ���������ɵ���Ŀʱѡ�����ԭ��
        strSql = "Select b.���� as ����,a.���� as ID,a.����,a.����,a.���� From ���쳣��ԭ�� a,���쳣��ԭ�� b" & _
                " Where a.����=1 And a.ĩ��=1 And a.�ϼ�=b.���� And b.ĩ��=0 " & _
                " Order by ����,a.����"
        vPoint = zlControl.GetCoordPos(vsPath.Hwnd, vsPath.CellLeft, vsPath.CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "���쳣��ԭ��", True, , , True, True, True, _
                 vPoint.X, vPoint.Y, vsPath.RowHeight(vsPath.Row), blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "ϵͳû�г�ʼ���쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            strReason = rsTmp!ID
        End If
    End If
    '4.��鲡��
    strSql = "Select 1 From ���Ӳ�����¼ Where ·��ִ��id = [1] And (���ʱ�� is not null or ��ӡ�� is not null) And rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����Ŀ��Ӧ�Ĳ�����ǩ�����Ѵ�ӡ������ȡ����", vbInformation, gstrSysName
        Exit Sub
    End If
    '����°���Ӳ���
    If Not CheckDelNewEMR(lngִ��ID & "", 0, rsTmp) Then
        Exit Sub
    Else
        'ɾ��
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            strSql = "<parameter><taskid>" & rsTmp!����ID & "</taskid></parameter>"
            Call gobjEmr.DeleteTask(strSql)
            rsTmp.MoveNext
        Next
        Err.Clear: On Error GoTo 0
    End If
    
    With vsPath
        If MsgBox("ȷʵҪȡ��·����Ŀ""" & .TextMatrix(.Row, .Col) & """��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    End With
    If Not mbln����ִ�л��� Then
        '�ж��Ƿ��Ѿ�����
        strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            'ǿ��ȡ�������������Ȩ��
            MsgBox "�������ɵ���Ŀ��������ȡ������֮ǰ���Զ�ȡ��������", vbInformation, gstrSysName
            Call FuncEvaluateCancel(False, False)
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If strReason <> "" Then
        strSql = "Zl_����·������_Update(" & lngִ��ID & ",'" & vsPath.TextMatrix(vsPath.Row, 0) & "',Null,NULL,NULL,NULL,NULL,'" & strReason & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "�޸�·����Ŀ")
    End If
    strSql = "Zl_����·������_Delete(" & lngִ��ID & "," & IIf(strReason <> "", "2", "0") & "," & mint���� & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAppendItemModify()
'���ܣ��޸�·������Ŀ
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngִ��ID As Long
            
    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With
    
    strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";�׶�����;") = 0 Then
            MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ������������������޸�·������Ŀ��", vbInformation, gstrSysName
            Exit Sub
        Else
            'ȡ������
            If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ������������޸�·������Ŀ��" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        End If
    End If
    
    If frmPathAppend.ShowMe(mfrmParent, mint����, mPati, mPP, "", 2, "", lngִ��ID, mclsMipModule) Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncAppendItem(ByVal bytUseType As Byte, Optional ByVal strItemType As String, Optional ByVal strAdviceIDs As String, _
                                Optional ByVal lngִ��ID As Long, Optional ByVal datDate As Date) As Boolean
'���ܣ����·������Ŀ(ͨ��clsDockPath�еĽӿڿ��Ÿ�ҽ����������)
'������bytUseType=0-ֱ�����,1-ҽ���¿�ʱ���
'       strItemType=ҽ���ӿڵ���ʱ���루�������һ����Ŀ�ķ��ࣩ
'       strAdviceIDs=ҽ���ӿڵ���ʱ����,ҽ�����
'       datDate =ҽ���Ŀ�ʼִ�����ڣ�ͬһ��·����ҽ����
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim DatCur As Date
    Dim blnRefresh As Boolean
    
    If mint���� = 0 Then 'ҽ�������������ڣ���ʿû����������
        strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";�׶�����;") = 0 Then
                MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ��������������������·������Ŀ��", vbInformation, gstrSysName
                Exit Function
            Else
                'ȡ������
                If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ��������������·������Ŀ��" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    Call FuncEvaluateCancel(False, False)
                Else
                    Exit Function
                End If
            End If
        Else
            'δ���������Ƿ��ǵ���δ���������ǵĻ�����ʾ�Ƿ�Ҫ��ӵ����һ���׶�
            If bytUseType = 0 Then
                DatCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                If DatCur <> Format(mPP.��ǰ����, "yyyy-MM-dd") Then
                    If MsgBox("��Ҫ���·������Ŀ��""" & mPP.��ǰ���� & """?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call FuncSendItem
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf mint���� = 1 Then
        If bytUseType = 0 Then
            Call GetPhaseInNurse(0, mPP.��ǰ�׶�ID, mPP.��ǰ����, mPP.��ǰ����)
            blnRefresh = True
        End If
    End If
    
    If bytUseType = 0 Then
        With vsPath
            If .Row > 0 And .Row < .Rows - 2 Then strItemType = .TextMatrix(.Row, .FixedCols - 1) '���һ����"·������"
        End With
    End If
    If frmPathAppend.ShowMe(mfrmParent, mint����, mPati, mPP, strItemType, bytUseType, strAdviceIDs, lngִ��ID, mclsMipModule, datDate) Or blnRefresh Then
        FuncAppendItem = True
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncImport(Optional ByVal lngHwnd As Long) As Boolean
'���ܣ�����·��
'������lngHwnd=�°没�����븸��������Ĭ��Ϊ0,�°没������ʾ���벻�ɹ���ԭ��
    Dim rsTmp As ADODB.Recordset
    '----
    Dim strSql As String
    '----
    Dim lngPPStatus As Long
    Dim t_pp As TYPE_PATH_Pati
    Dim str���� As String, lngDiagnosisType As Long, lngDiagnosisSorce As Long
    Dim lng����ID As Long, lng���ID As Long
    
    '1.���ò��˵�ǰ�Ƿ��������ִ�е�·���������Ǳ���סԺ�����Ƶ�,��ǰ����������
    strSql = "Select b.���� From �����ٴ�·�� a,���ű� b Where a.����id = b.id And a.����ID = [1] And a.״̬ = 1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "FuncImport", mPati.����ID)
    If rsTmp.RecordCount > 0 Then
        If lngHwnd = 0 Then MsgBox "�ò�����[" & rsTmp!���� & "]��������ִ�е��ٴ�·�������������µ�·����", vbInformation, gstrSysName
        Exit Function
    End If
    
    lngPPStatus = mPP.����·��״̬
    
    FuncImport = frmPathImport.ShowMe(mfrmParent, mPati, 0, t_pp, , , , , , , , lngHwnd, str����, lngDiagnosisType, lngDiagnosisSorce, lng����ID, lng���ID)
    If lngHwnd <> 0 And FuncImport = True And t_pp.·��ID <> 0 Then
        '�°没�������ڴ�������ȥ������������,����ᱻ��С��
        FuncImport = frmEvaluate.ShowMe(mfrmParent, 0, 1, mPati, t_pp, str����, lngDiagnosisType, lngDiagnosisSorce, lng����ID, lng���ID, 0)
    End If
    If lngHwnd <> 0 Then Exit Function
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    
    If lngPPStatus <> mPP.����·��״̬ Then
        '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    If mPP.����·��״̬ = 1 Then
        '�������ɹ��������Ƿ���Ҫ��������ϲ�·��
        Call frmPathImport.ShowMe(mfrmParent, mPati, 2, t_pp, , , , True, mPP.����·��ID, , , lngHwnd)
    End If
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncImportMerge() As Boolean
'���ܣ�����ϲ�·��
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim t_pp As TYPE_PATH_Pati
    
    '1.�жϵ�ǰ�Ƿ�����
    If Val(mPP.��ǰ�׶�ID & "") <> 0 Then
        strSql = "Select ʱ����� From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount = 0 Then
            If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";�׶�����;") = 0 Then
                MsgBox "�ò�����" & mPP.��ǰ���� & "��û�н������������ܽ��к���������", vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("�ò�����" & mPP.��ǰ���� & "��û�н���������������������" & vbCrLf & "������Ҫ��������������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If frmEvaluate.ShowMe(mfrmParent, 1, 1, mPati, mPP) = False Then
                        Exit Function
                    Else
                        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
                        '�����󣬿��ܽ������˳�·�������Ը��������е�״̬�����ж��Ƿ�Ҫ��������,�˳�������򲻼�������
                        If mPP.����·��״̬ <> 1 Then
                            Exit Function
                        End If
                    End If
                Else
                    Exit Function
                End If
            End If
        End If
    
    End If
     '2.���ϲ�·������������5��
    strSql = "Select �ϲ�·������ From �����ٴ�·�� Where ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If Val(rsTmp!�ϲ�·������ & "") >= 5 Then
        MsgBox "�ò������Ѿ�������5���ϲ�·�����������ٵ����µĺϲ�·���ˡ�", vbInformation, gstrSysName
        Exit Function
    End If
        
    FuncImportMerge = frmPathImport.ShowMe(mfrmParent, mPati, 2, t_pp, , , , False, mPP.����·��ID)
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncUnImport(Optional ByVal blnPrompt As Boolean = True)
'���ܣ�ȡ������,δ����·��ʱ��ȡ������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim str����� As String
    Dim lngPPStatus As Long
    
    '�ȼ���Ƿ���ȡ��·����Ȩ��
    If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";ȡ������;") = 0 Then
        str����� = zlDatabase.UserIdentify(Me, "û��ȡ������Ȩ����Ҫ��ˡ�", glngSys, P�ٴ�·��Ӧ��, "ȡ������")
        If str����� = "" Then Exit Sub
    Else
        str����� = UserInfo.����
    End If
    strSql = "Select 1 From ����·��ִ�� Where ·����¼ID = [1] And rownum<2"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount > 0 Then
        
        If MsgBox("��ǰ�׶ε�·����Ŀ�����ɣ����Ƚ�����ȡ�����ɲ�����" & vbCrLf & "��ȷʵҪȡ���ò����ѵ�����ٴ�·����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        'Ҫˢ�£��Ա���ȡ·����Ϣ(��ǰ�׶ε�)
        If FuncDelAllItem(True, False) Then
            Call FuncUnImport(False)    '���µ��ã��ٴμ��
        End If
        Exit Sub
    ElseIf blnPrompt Then
        If MsgBox("��ȷʵҪȡ���ò����ѵ�����ٴ�·����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    lngPPStatus = mPP.����·��״̬
    
    gcnOracle.BeginTrans: blnTrans = True
    strSql = "Zl_����·������_Delete(" & mPP.����·��ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ������")
    '����ȡ�������¼
    strSql = "Zl_����·��ȡ��_Insert(" & mPati.����ID & "," & mPati.��ҳID & ",'" & UserInfo.���� & "','" & str����� & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ������")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    
    '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
    If lngPPStatus <> mPP.����·��״̬ Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncUnImportMerge(Optional ByVal blnPrompt As Boolean = True)
'���ܣ�ȡ������ϲ�·��
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim str����� As String
    Dim colSQL As New Collection
    Dim strIDs As String, i As Long
    Dim t_pp As TYPE_PATH_Pati
    
    '�����Ѿ�����ĺϲ�·��
    strSql = "Select a.Id, a.·��id, b.����, b.����, b.˵��, NVL(Sign(Max(c.Id)),0) As �Ƿ�ִ��" & vbNewLine & _
            "From ���˺ϲ�·�� A, �ٴ�·��Ŀ¼ B, ����·��ִ�� C" & vbNewLine & _
            "Where a.·��id = b.Id And c.�ϲ�·����¼id(+) = a.Id And a.��Ҫ·����¼id = [1]" & vbNewLine & _
            "Group By a.Id, a.·��id, b.����, b.����, b.˵��"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    '���ֻ��һ���ϲ�·������ֱ����ʾ���򵯳�ѡ��
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò���δ�����κκϲ�·��������ȡ�����롣", vbInformation, gstrSysName
        Exit Sub
    ElseIf rsTmp.RecordCount = 1 Then
        If Val(rsTmp!�Ƿ�ִ�� & "") = 1 Then
            MsgBox "�ò��˵ĺϲ�·��:" & rsTmp!���� & "�Ѿ���������Ŀ����ȡ���ϲ�·������Ŀ����ȡ�����롣", vbInformation, gstrSysName
            Exit Sub
        Else
            If MsgBox("��ȷ��Ҫȡ������ϲ�·����" & rsTmp!���� & "?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        strIDs = rsTmp!ID & ""
    Else
        If Not frmPathImport.ShowMe(mfrmParent, mPati, 3, t_pp, , , , , , rsTmp, True) Then Exit Sub
        Unload frmPathImport
        If rsTmp.RecordCount = 0 Then Exit Sub
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            strIDs = strIDs & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        strIDs = Mid(strIDs, 2)
    End If
    If strIDs <> "" Then
        For i = 0 To UBound(Split(strIDs, ","))
            strSql = "Zl_����·������_Delete(" & mPP.����·��ID & "," & Val(Split(strIDs, ",")(i)) & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        Next
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "ȡ���ϲ�·��")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ViewMergeImport()
'���ܣ��鿴�ϲ�·����������
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim t_pp As TYPE_PATH_Pati
    
    '�����Ѿ�����ĺϲ�·��
    strSql = "Select a.Id, a.·��id, b.����, b.����, b.˵��" & vbNewLine & _
            "From ���˺ϲ�·�� A, �ٴ�·��Ŀ¼ B" & vbNewLine & _
            "Where a.·��id = b.Id  And a.��Ҫ·����¼id = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount = 0 Then
        MsgBox "��ǰδ�����κκϲ�·����", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not frmPathImport.ShowMe(mfrmParent, mPati, 3, t_pp, , , , , , rsTmp) Then Exit Sub
    Unload frmPathImport
    If rsTmp.RecordCount = 0 Then Exit Sub
    Call frmEvaluate.ShowMe(mfrmParent, 0, 0, mPati, mPP, , , , , , 1, , Val(rsTmp!ID & ""))
    
    Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'���ܣ�ȡ������,δ����ʱ����ȡ����������Զ�������ֻ��ȡ��������
'������blnPrompt=�Ƿ񵯳�ѯ����ʾ
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim lngPPStatus As Long
    
    On Error GoTo errH
    
    If Not mbln���ò����� Then
        strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount = 0 Then
            MsgBox "�ò�����" & mPP.��ǰ���� & "��û�н���������", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        strSql = "Select A.�׶�ID,A.����,A.���� From (Select t.�׶�id, t.����, t.���� From ����·������ T Where t.·����¼id = [1] Order By t.�Ǽ�ʱ�� Desc, t.���� Desc) A where rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        If rsTmp.RecordCount > 0 Then
            If Format(rsTmp!����, "YYYY-MM-DD") < mPP.��ǰ���� Or (Format(rsTmp!����, "YYYY-MM-DD") = mPP.��ǰ���� And Val(rsTmp!�׶�ID & "") <> mPP.��ǰ�׶�ID) Then
                mPP.��ǰ�׶�ID = rsTmp!�׶�ID
                mPP.��ǰ���� = Format(rsTmp!����, "YYYY-MM-DD")
                mPP.��ǰ���� = rsTmp!����
             End If
         Else
            MsgBox "�ò��˲������κ�������¼��", vbInformation, gstrSysName
            Exit Function
         End If
    End If
        
    If blnPrompt Then
        strSql = "Select 1 From ���˺ϲ�·�� Where ��Ҫ·����¼ID = [1] And ��Ҫ·���׶�ID = [2] And ��Ҫ·������=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
        If rsTmp.RecordCount > 0 Then
            If MsgBox("��ǰ�׶��Ѿ������˺ϲ�·����ȡ��������ͬʱȡ������ϲ�·�����Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("��ȷ��Ҫȡ����" & mPP.��ǰ���� & "���������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    lngPPStatus = mPP.����·��״̬
    
    strSql = "Zl_����·������_Delete(" & mPP.����·��ID & ", " & mPP.��ǰ�׶�ID & ",To_Date('" & mPP.��ǰ���� & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ������")
    FuncEvaluateCancel = True
    If blnRefresh Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
                 
    '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
    If lngPPStatus <> mPP.����·��״̬ Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncEvaluate()
'���ܣ��׶�����,ֻ�ܶԵ�ǰ�׶ε����һ����ִ���˵Ľ���
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim strTmp As String
    Dim bln��¼ As Boolean
    Dim blnRefresh As Boolean
    Dim strDate  As String
    Dim lngPPStatus As Long
    
    '1.�������Ĳ��������� '�ѽ����Ĳ���������(�����˲˵���)
    '2.ֻ�ܶ����һ��ִ�еļ�¼��������(������������ָ��ģ�����������������ɴ���·����û�ж������������ִ������)����Ϊ����Ϊ��������ܽ���·��
    '3.����ý׶ε�������Ŀ��ҽ�����ɵ���Ŀ����ִ�к��������

    On Error GoTo errH
  
    If Not mbln���ò����� Then
        strSql = "Select 1 From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ�����������", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        '��ѯ���ȱʡ�ǰ���·��ִ��ID�Ž���,�ܹ�ȡ����һ��δ�����Ľ׶�
        strSql = "Select * " & vbNewLine & _
                "From (Select a.�׶�id, a.����, a.����" & vbNewLine & _
                "       From ����·��ִ�� A, ����·������ B" & vbNewLine & _
                "       Where a.·����¼id = [1] And a.·����¼id = b.·����¼id(+) And a.�׶�id = b.�׶�id(+) And a.���� = b.����(+) And b.�׶�id Is Null" & vbNewLine & _
                "       Order By a.Id) A" & vbNewLine & _
                "Where Rownum < 2"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        '��ȡ�������ڱ���
        strDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
        If rsTmp.RecordCount > 0 Then
            If Format(rsTmp!����, "YYYY-MM-DD") <= strDate Then  '����������
                If mPP.��ǰ���� & "_" & mPP.��ǰ�׶�ID = Format(rsTmp!����, "YYYY-MM-DD") & "_" & Val(rsTmp!�׶�ID & "") Then
                    '�ǲ�¼����,û����ǰ���ɵ�·���׶�
                    bln��¼ = False
                ElseIf Format(rsTmp!����, "YYYY-MM-DD") <= mPP.��ǰ���� Then
                    mPP.��ǰ�׶�ID = rsTmp!�׶�ID
                    mPP.��ǰ���� = rsTmp!���� & ""
                    mPP.��ǰ���� = rsTmp!����
                    bln��¼ = True
                End If
            ElseIf Format(rsTmp!����, "YYYY-MM-DD") > strDate Then
                MsgBox "������Ҫ�����Ľ׶����ڣ���" & Format(rsTmp!����, "YYYY-MM-DD") & "��" & vbCrLf & "�����˵�ǰ���ڣ���" & strDate & "��" & vbCrLf & "���ܽ�������������", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
        Else
            MsgBox "�ò������н׶ζ��Ѿ����������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    'ִ�еǼǼ��
    If Not CheckPathIsExecuted(blnRefresh) Then
        'ǿ��ˢ��
        If blnRefresh Then
            Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
        End If
        Exit Sub
    End If
    
    lngPPStatus = mPP.����·��״̬
    
    If frmEvaluate.ShowMe(mfrmParent, 1, 1, mPati, mPP, , , , , , , , , bln��¼) Or bln��¼ Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    
    '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
    If lngPPStatus <> mPP.����·��״̬ Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    If mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3 Then RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncReEvaluate()
'���ܣ��޸���������������˺����׶ε���Ŀ�������޸��������Ϊ�����������������ڱ���Ĵ洢�������жϡ�
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim bln��¼ As Boolean
    Dim lng�׶�ID As Long
    Dim strSysDate As String
    Dim lng���� As Long
    Dim lngPPStatus As Long

    On Error GoTo errH
    
    If Not mbln���ò����� Then
        strSql = "Select ԭ·��ID From ����·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount = 0 Then
            MsgBox "�ò����ڵ�ǰ�׶λ�û�н���������", vbInformation, gstrSysName
            Exit Sub
        ElseIf Val("" & rsTmp!ԭ·��ID) <> 0 Then
            MsgBox "�ò�������ת������·�������Ҫ�޸���������ȡ������������������", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        strSql = "Select A.�׶�ID,A.����,A.���� From (Select t.�׶�id, t.����, t.���� From ����·������ T Where t.·����¼id = [1] Order By t.�Ǽ�ʱ�� Desc, t.���� Desc) A where rownum<2"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        If rsTmp.RecordCount > 0 Then
            '���ڵ���������½׶�ID����鲻�����ݱ���
            If Format(rsTmp!����, "YYYY-MM-DD") <= mPP.��ǰ���� Then
                mPP.��ǰ�׶�ID = rsTmp!�׶�ID
                mPP.��ǰ���� = Format(rsTmp!����, "YYYY-MM-DD")
                mPP.��ǰ���� = rsTmp!����
                bln��¼ = True
             End If
         Else
            MsgBox "�ò����ڵ�ǰ�׶λ�û�н���������", vbInformation, gstrSysName
            Exit Sub
         End If
    End If
    
    lngPPStatus = mPP.����·��״̬
    
    If frmEvaluate.ShowMe(mfrmParent, 1, 2, mPati, mPP, , , , , , , , , bln��¼) Or bln��¼ Then
        Call zlRefresh(mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID, mPati.����״̬)
    End If
    
    '��ǰ����·��״̬�����仯ʱ����Lis����·��״̬
    If lngPPStatus <> mPP.����·��״̬ Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.����ID, mPati.��ҳID, mPP.����·��״̬)
        End If
    End If
    
    If mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3 Then RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewReport(ByVal str����ID As String, ByVal lngҽ��ID As Long)
'���ܣ����ı���
        
    '���ж��Ƿ���Լ�������
    If IsNumeric(str����ID) Then
        If CheckEPRReport(Val(str����ID), lngҽ��ID) = 2 Then
            If InStr(GetInsidePrivs(pסԺҽ���´�), "����δ��ɱ���") > 0 Then
                MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
            Else
                MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        RaiseEvent ViewEPRReport(Val(str����ID), False)
    Else
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(0, str����ID, Val(zlDatabase.GetPara("�Զ���Ǳ������״̬", glngSys, pסԺҽ���´�, "1")) = 1, mfrmParent)
    End If
    
End Sub


Public Function CheckEPRReport(ByVal lng����ID As Long, ByVal lngҽ��ID As Long) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lng·��ִ��ID=����·��ִ�м�¼�е�ID
'      lng����ID=���ر��没��ID
'���أ�
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
    '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
    strSql = _
        " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
        " Union ALL" & _
        " Select ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ��" & _
        " From (" & _
            " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
            " Order by A.���" & _
        " ) Where Rownum=1" & _
        " Order by ����,����ʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID)
    If Nvl(rsTmp!ִ�й���, 0) >= 5 Or Nvl(rsTmp!ִ��״̬, 0) = 1 Then
        CheckEPRReport = 1
    Else
        CheckEPRReport = 2
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mcolReason = Nothing
    Set mclsMipModule = Nothing
    SaveWinState Me, App.ProductName
    Set mobjPublicPACS = Nothing
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String, lngId As Long, i As Long
    Dim strSql As String, rsTmp As ADODB.Recordset
        
    lngId = fraMore.Tag
    If lngId = 0 Then
        Call zlCommFun.ShowTipInfo(0, strInfo)
    Else
        strSql = "Select  decode(NVL(Nvl(a.������, b.������),1),1,'ҽ��',2,'��ʿ') As ������," & IIf(mbln����ִ�л���, "A.ִ�н��,A.ִ��˵��,A.ִ����,to_char(A.ִ��ʱ��,'yyyy-mm-dd hh24:mi') as ִ��ʱ��,", "") & _
                " A.�Ǽ���,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi') as �Ǽ�ʱ�� From ����·��ִ�� A,�ٴ�·����Ŀ B Where A.��ĿID=B.ID(+) And A.ID = [1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        If rsTmp.RecordCount > 0 Then
            With rsTmp
                For i = 0 To .Fields.count - 1
                    strInfo = strInfo & .Fields(i).Name & "��" & .Fields(i).Value & vbCrLf
                Next
            End With
            Call zlCommFun.ShowTipInfo(fraMore.Hwnd, strInfo, True)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optSelect_Click(Index As Integer)
    If Me.Visible Then
        optSelect(IX_ALL).Tag = Index  '��ǵ�ǰѡ����
        Call LoadPathItem 'ˢ��
    End If
End Sub

Private Sub vsFlow_DblClick()
    Dim lngPhaseID As Long
    If mPP.����·��״̬ = 0 And mPP.����·��ID <> 0 Then   '����ʧ��
        Call frmEvaluate.ShowMe(mfrmParent, 0, 0, mPati, mPP)
    Else
        lngPhaseID = Val(vsFlow.ColData(vsFlow.Col))
        If lngPhaseID <> 0 Then
            Call frmPathSend.ShowMe(mfrmParent, 2, mint����, mPati, mPP, lngPhaseID, 0, , , , mclsMipModule)
        ElseIf vsFlow.Col = 0 And mPP.·��ID <> 0 Then
            Call frmPathDefinition.ShowMe(mfrmParent, mPP.·��ID)
        Else
            If vsFlow.Col = vsFlow.Cols - 2 And gstrDBUser = "ZLHIS" Then
                vsFlow.Editable = flexEDKbdMouse
            End If
        End If
    End If
End Sub

Private Sub vsFlow_LostFocus()
    If Not (mPP.����·��״̬ = 0 And mPP.����·��ID <> 0) Then
        vsFlow.ForeColorSel = vsFlow.CellForeColor
    End If
End Sub

Private Sub vsFlow_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 '���ܣ����ڲ���ʱǿ��ɾ�����һ�����Ŀ(ѡ�����һ����ͷ������DELA)
    Dim strPass As String, i As Long
        
    If vsFlow.Col = vsFlow.Cols - 2 Then
        strPass = UCase(vsFlow.EditText)
        vsFlow.EditText = ""
        If strPass = "DELA" Then
            If MsgBox("��ȷ��Ҫɾ�����һ���������Ŀ��", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            Call FuncDelPhaseItem
        End If
        vsFlow.Editable = flexEDNone
    End If
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Or OldCol <> NewCol Then
        If fraMore.Visible Then fraMore.Visible = False
        
        If NewRow <> -1 And NewCol <> -1 And mblnUnChange = False Then
            '��ʾ·����Ŀ���ɵ�ҽ���嵥
            Dim strTmp As String
            
            strTmp = vsPath.Cell(flexcpData, NewRow, NewCol)
            If InStr(strTmp, "|") > 0 Then
                Call UCAdvice.ShowAdvice(1, "", Val("" & Split(strTmp, "|")(0)))
            Else
                Call UCAdvice.ShowAdvice(1, "", 0)
            End If
        Else
            Call UCAdvice.ShowAdvice(1, "", 0)
        End If
    End If
End Sub

Private Sub vsPath_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45
End Sub

Private Sub vsPath_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    
    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub vsPath_DblClick()
    Dim lng��ĿID As Long

    With vsPath
        If Trim(.TextMatrix(.Row, .Col)) <> "" And .Cell(flexcpData, .Row, .Col) <> "" And .Row <> .Rows - 1 Then
            lng��ĿID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
            If lng��ĿID <> 0 Then
                Call frmPathItemEdit.ShowView(mfrmParent, lng��ĿID)
            End If
        End If
    End With
End Sub

Private Sub vsPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsPath
        If .MouseCol >= .FixedCols And .MouseRow >= .FixedRows Then
            Dim lngId As Long, lngRow As Long, lngCol As Long, lngItemID As Long
            
            lngRow = .MouseRow: lngCol = .MouseCol
            If .Cell(flexcpData, lngRow, lngCol) <> "" And lngRow <> .Rows - 1 Then
                lngId = Split(.Cell(flexcpData, lngRow, lngCol), "|")(0)
                lngItemID = Split(.Cell(flexcpData, lngRow, lngCol), "|")(1)
                If lngItemID = 0 Then
                    .ToolTipText = ""
                    Call zlCommFun.ShowTipInfo(.Hwnd, mcolReason("C" & lngId), True)      '·������Ŀ�����ԭ��
                Else
                    If .ToolTipText = "" Then .ToolTipText = "˫���鿴·����Ŀ����"
                    Call zlCommFun.ShowTipInfo(.Hwnd, "")
                End If
            Else
                .ToolTipText = ""
            End If
            
            If lngId = 0 Then
                If imgMore.Visible Then fraMore.Visible = False
                fraMore.Tag = ""
            Else
                If lngRow = .Row And lngCol = .Col Then
                    fraMore.BackColor = .BackColorSel
                Else
                    fraMore.BackColor = .BackColor
                End If
            
                fraMore.Tag = lngId
                If fraMore.Visible = False Then fraMore.Visible = True
                fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - imgMore.Height - 30
                fraMore.Left = .Left + .ColPos(lngCol) + .ColWidth(lngCol) - imgMore.Width - 30
            End If
        Else
            If fraMore.Visible Then fraMore.Visible = False
        End If
    End With
End Sub

Private Sub vsPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim lng��ĿID As Long

    '��ʾ�༭�˵����������
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub OutLogModi()
    Dim colSQL As New Collection, i As Long, blnTrans As Boolean

    Call frmPathOutLog.ShowMe(mfrmParent, mPati.����ID, mPati.��ҳID, 2, colSQL, mPP.·��ID, mPP.����·��ID)

    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    'ִ�г����ǼǱ��SQL
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "�޸ĳ����ǼǱ�")
    Next
    gcnOracle.CommitTrans: blnTrans = False

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����·��������ͼ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    mlngFontSize = IIf(bytSize = 0, CON_SmallFontSize, CON_BigFontSize)
    
    vsFlow.Font.Size = mlngFontSize
    vsFlow.Redraw = flexRDDirect
    
    Call Grid.SetFontSize(vsPath, mlngFontSize)
    If vsPath.FixedRows > 1 Then vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45 '��ҪDraw֮�����Ч
    
    Call UCAdvice.SetVsAdviceFontSize(mlngFontSize)
End Sub

Private Sub MovePathItem(ByVal lngWay As Long)
'����:��ǰ��Ԫ��ѡ��·������Ŀʱ������·������Ŀ�������ƶ�
'����:lngWay=1����һ��,-1����һ��(�൱����һ������һ��)
    Dim lngId       As Long
    Dim lngItemNum  As Long
    Dim arrSQL()    As Variant
    Dim i           As Integer
    Dim blnTran     As Boolean
    Dim blnDo As Boolean, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long

    blnDo = True: blnFind = False

    With vsPath
        Do While blnDo

            If .TextMatrix(.Row, .FixedCols - 1) <> .TextMatrix(.Row - lngWay, .FixedCols - 1) Or .Cell(flexcpData, .Row - lngWay, .Col) = "" Then
                MsgBox "��Ŀ����:" & .TextMatrix(.Row, .Col) & vbCrLf & _
                       "�Ѵ��ڡ�" & .TextMatrix(.Row, .FixedCols - 1) & "�������" & IIf(lngWay > 0, "��һ��", "���һ��"), vbInformation, gstrSysName
                blnDo = False: blnFind = False: Exit Do
            Else
                lngRow = .Row - lngWay: lngCol = .Col
                blnFind = True: Exit Do
            End If
        Loop
        '������Ŀ���
        If blnFind Then
            arrSQL = Array()

            lngId = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_����·�����_update(" & lngId & "," & lngItemNum & ")"

            lngId = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_����·�����_update(" & lngId & "," & lngItemNum & ")"

            On Error GoTo errH
            gcnOracle.BeginTrans: blnTran = True
            For i = LBound(arrSQL) To UBound(arrSQL)
                zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
            Next i
            gcnOracle.CommitTrans: blnTran = False

            Call ClearPathItem(True)
            Call LoadPathItem

            '�����ƶ�
            .Row = lngRow: .Col = lngCol
            .ShowCell lngRow, lngCol

        End If
    End With
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckSendOfBefore() As Boolean
'����:��ǰ����ǰ���
'����: T-������ǰ����,F-��������ǰ����
    Dim strTmp As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim blnReturn As Boolean
    
    On Error GoTo errH
    strTmp = "," & UserInfo.���� & ","
    If InStr(strTmp, ",ҽ��,") > 0 Then
        '������ǰ����
        blnReturn = True
    Else
        '��ʿ��ǰ����ʱ������ҽ���������ɹ�û��
        strSql = "Select 1" & vbNewLine & _
            "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
            "Where a.·����¼id = [1] And a.�׶�ID=[2] And a.����=[3] And a.��Ŀid = b.Id(+) And Nvl(Nvl(a.������, b.������),1) = 1 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
        If rsTmp.RecordCount > 0 Then
            blnReturn = True
        Else
            MsgBox "ҽ����û����ǰ������һ���·����Ŀ����ʿ������ǰ���ɡ�", vbInformation + vbOKOnly
            blnReturn = False
        End If
    End If
    CheckSendOfBefore = blnReturn
    Exit Function
errH:
   If ErrCenter() = 1 Then
        Resume
   End If
   Call SaveErrLog

End Function

Private Function CheckDelNewEMR(ByVal str·��ִ��IDs As String, ByVal bytMode As Byte, ByRef rsTmp As ADODB.Recordset) As Boolean
'����:ɾ���°���Ӳ������
'����:
'   str·��ִ��IDs
'   bytMode 0-����·����Ŀ,1-���·����Ŀ
'   str����IDs -����ID�� ID,ID,ID....
    Dim strSql As String
    Dim rsTask As ADODB.Recordset
    Dim i As Long
    
    If bytMode = 0 Then
        strSql = "Select ����ID from ����·������ where ·��ִ��Id=[1] "
    Else
        strSql = "Select ����ID from ����·������ where ·��ִ��Id in (Select Column_value From Table(f_Num2List([2]))) "
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(str·��ִ��IDs), str·��ִ��IDs)
    If rsTmp.RecordCount > 0 Then
        If Not gobjEmr Is Nothing Then
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                For i = 1 To rsTmp.RecordCount
                    strSql = "<parameter><taskid>" & rsTmp!����ID & "</taskid></parameter>"
                    On Error Resume Next
                    Set rsTask = gobjEmr.GetTaskStatus(strSql)
                    Err.Clear: On Error GoTo 0
                    '��¼������0�У������쳣����1�����ݣ�
                    '��¼�������ֶΣ�ID�������ˣ�����ʱ�䣬����ˣ����ʱ�䣬����ˣ����ʱ�䣬�����ӡ�ˣ������ӡʱ�䣻����IDΪ����ID����ID�������ˣ�����ʱ���⣬�����ֶζ�����Ϊ�գ�
                    '�������Ϊ�ǿ�ʱ����ʾ�ò����ļ�����ɡ�
                    If rsTask.State <> adStateClosed Then
                        If rsTask.RecordCount = 1 Then
                            If rsTask!����� <> "" Then
                                If bytMode = 0 Then
                                    MsgBox "����Ŀ��Ӧ�ĵ��Ӳ����ļ������,����ȡ����", vbInformation, gstrSysName
                                Else
                                    MsgBox "���������д��ڶ�Ӧ�ĵ��Ӳ����ļ������,����ȡ���������ɡ�", vbInformation, gstrSysName
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                    rsTmp.MoveNext
                Next
            End If
        End If
    End If
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    CheckDelNewEMR = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPhaseInNurse(ByVal bytType As Byte, ByRef lng�׶�ID As Long, ByRef lng���� As Long, Optional ByRef strDate As String, _
            Optional ByRef lng��һ�׶� As Long, Optional ByRef lng��һ�� As Long, _
             Optional ByRef strPhase As String = "-1")
'����:��ȡ��ʿ���Ͻ׶κ�����
'
'����:
'    bytType:0-��ʿ�������ɡ����\ȡ��·������Ŀ��ȡ������ʱȱʡ�Ľ׶κ�����;��������ɵĽ׶κ�������
'            1-��ʿ����ʱȱʡ�Ľ׶κ�����
'����:
'    ���Σ�
'    lng�׶�ID:
'    lng����: ����������ͬһ�׶�,���ɶ�������
'    strPhase:�׶���ϢSQL =-1������SQL
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim lngRowNUM As Long
    
    On Error GoTo errH
    'ȡ��ʿ����������ɵĽ׶μ�����,�Ǽ�ʱ��ȡ��С����Ϊ�˱�֤ȡ��������ʱ��ִ�м�¼,�������������ɡ��ݴ�·����Ŀ
    If mPP.��ǰ�׶η�֧ID = 0 Then
        strSql = "Select �׶�ID,����,����,������,RowNum as ���� " & vbNewLine & _
        "From (Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') as ����,������ " & vbNewLine & _
                 "From (Select a.�׶�id, a.����, a.����,a.·����¼id,NVl(������,1) as ������" & vbNewLine & _
                 "       From ����·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1] " & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����,a.·����¼id,NVl(������,1)) A, �ٴ�·���׶� B,�ٴ�·���׶� C,�����ٴ�·�� G" & vbNewLine & _
                 "Where a.�׶�id = b.Id And b.��id=c.id(+) And g.id=A.·����¼ID " & vbNewLine & _
                 "Order By ����,Decode(g.·��id,b.·��id,1,0), NVL(c.���,b.���))"
    Else
        strSql = "Select �׶�ID,����,����,������,RowNum as ���� " & vbNewLine & _
            "From (Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') as ����,������ " & vbNewLine & _
                 "From (Select a.�׶�id, a.����, a.����,a.·����¼id,NVl(������,1) as ������" & vbNewLine & _
                 "       From ����·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1] " & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����,a.·����¼id,NVl(������,1)) A, �ٴ�·���׶� B,�ٴ�·���׶� C,�ٴ�·����֧ D,�ٴ�·���׶� E,�ٴ�·���׶� F,�����ٴ�·�� G" & vbNewLine & _
                 "Where a.�׶�id = b.Id And b.��id=c.id(+) And b.��֧id=d.id(+) and d.ǰһ�׶�id=e.id(+) And e.��id=f.id(+)  And g.id=A.·����¼ID " & vbNewLine & _
                 "Order By ����,Decode(g.·��id,b.·��id,1,0), Decode(b.��֧ID,Null,NVL(c.���,b.���),NVL(c.���,b.���)+NVL(f.���,e.���)))"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����·����Ŀ", mPP.����·��ID)
    rsTmp.Filter = "������=2"  '=2��ʿ���� =1ҽ������
    If rsTmp.RecordCount = 0 Then
        '��ʿ���ϻ�û�����ɹ��κ�·����Ŀ,ҽ���������ɵĵ�һ���׶κ�����
        rsTmp.Filter = "������=1"
        If rsTmp.RecordCount > 0 Then
            '��ʿ��һ�׶μ�����
            If bytType = 1 Then
                lng��һ�׶� = Val(rsTmp!�׶�ID & "")
                lng��һ�� = Val(rsTmp!���� & "")
                strDate = rsTmp!���� & ""
                If strPhase <> "-1" Then
                    strPhase = Rec.ToSQL(rsTmp)
                End If
                Exit Sub
            End If
        End If
    Else
        '��ʿ���׶μ�����
        rsTmp.Sort = "���� DESC"
        lng�׶�ID = Val(rsTmp!�׶�ID & "")
        lng���� = Val(rsTmp!���� & "")
        strDate = rsTmp!���� & ""
        If bytType = 1 Then
            rsTmp.Sort = ""
            rsTmp.Filter = "������=1"
            Do
                If rsTmp!�׶�ID & "_" & rsTmp!���� & "_" & rsTmp!���� = lng�׶�ID & "_" & lng���� & "_" & strDate Then
                    '�ҵ���ʿ��һ�׶μ�����
                    lngRowNUM = Val(rsTmp!���� & "")
                    rsTmp.Filter = "������=1 and ���� >" & lngRowNUM
                    lng��һ�׶� = Val(rsTmp!�׶�ID & "")
                    lng��һ�� = Val(rsTmp!���� & "")
                    If strPhase <> "-1" Then
                        strPhase = Rec.ToSQL(rsTmp)
                    End If
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop While Not rsTmp.EOF
       End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncConvertPathTable() As VSFlexGrid
'����:ת���ٴ�·����,����Ӧ������Ĵ�ӡ���� 78233
'����
'   -ת�����·����
    Dim lngFirstSameCol As Long
    Dim lngLastRow As Long
    Dim i As Long, j As Long, k As Long
    
    Grid.CopyTo vsPath, vsPathPrint(0)
    vsPath.Redraw = flexRDNone
    vsPathPrint(0).MergeCol(0) = True
    vsPathPrint(0).MergeRow(0) = True
    
    With vsPathPrint(0)
        'һ���׶δ�ӡһ�У�����ӡ����������
        For i = 2 To .Cols - 1
            If .TextMatrix(R0�׶���, i) = .TextMatrix(R0�׶���, i - 1) Then
            '��������Ϊͬһ�׶�Ҫ�ϲ�
                For j = 1 To i - 1
                    If .TextMatrix(R0�׶���, i) = .TextMatrix(R0�׶���, j) Then
                        lngFirstSameCol = j '�ҵ��뵱ǰ����ͬһ�׶ε�����
                        Exit For
                    End If
                Next
                'j-��ǰ��,i-��ǰ��
                For j = R2���� + 1 To .Rows - 1
                    If .TextMatrix(j, i) <> "" And .TextMatrix(j, 0) <> "�������" Then
                        k = 0
                        For k = R2���� + 1 To .Rows - 1
                            If .TextMatrix(j, i) = .TextMatrix(k, lngFirstSameCol) Then
                                Exit For
                            End If
                        Next
                        If k = .Rows Then
                            '����
                            k = 0
                            lngLastRow = 0
                            For k = R2���� + 1 To .Rows - 1
                                If .TextMatrix(j, 0) = .TextMatrix(k, 0) Then
                                    If .TextMatrix(k, lngFirstSameCol) = "" Then
                                        'ͬ�����¿�������
                                        .TextMatrix(k, lngFirstSameCol) = .TextMatrix(j, i)
                                        Exit For
                                    End If
                                    lngLastRow = k
                                End If
                            Next
                            If k = .Rows Then
                                'ͬ�������һ������һ��
                                .AddItem "", lngLastRow + 1
                                .TextMatrix(lngLastRow + 1, lngFirstSameCol) = .TextMatrix(j, i)
                                .TextMatrix(lngLastRow + 1, 0) = .TextMatrix(j, 0)
                                .RowHeight(lngLastRow + 1) = .RowHeight(j)
                            End If
                        Else
                            'ǰһ�д��ڣ����Բ�����
                        End If
                        
                    End If
                Next
                '���ɾ����
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        'ɾ����
        For i = .Cols - 1 To 0 Step -1
            If .ColHidden(i) = True And .ColWidth(i) = 0 Then
                '���һ��ֱ��ɾ��
                If i = .Cols - 1 Then
                    .Cols = .Cols - 1
                Else
                    '������ǰ��
                    For k = i + 1 To .Cols - 1
                        For j = 0 To .Rows - 1
                            .TextMatrix(j, k - 1) = .TextMatrix(j, k)
                        Next
                    Next
                    .Cols = .Cols - 1
                End If
            End If
        Next
        '�������ں�����
        .RowHidden(R1����) = True: .RowHidden(R2����) = True
        .RowHeight(R1����) = 0: .RowHeight(R2����) = 0
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч
    End With
    Set FuncConvertPathTable = vsPathPrint(0)
End Function

Private Function CheckPathIsExecuted(Optional ByRef blnRefresh As Boolean) As Boolean
'-------------------------------------------------------------------------------------------
'����:��鵱ǰ�׶��Ƿ����δִ�е�·����Ŀ
'������=1 ���ɻ��ڵ���,=2 ����ʱ�����
'���أ�F-����δ���ִ�еǼǵ�·����Ŀ,���������ɻ�����
'     T-������δ���ִ�еǼǵ�·����Ŀ\�����ִ�еǼ�������������ɻ�����
'˵����1.��ʿ����ʱ,��ǰδִ��ʱ�����������µĽ׶�,ҽ������ʱ,��ǰδִ�в����������µĽ׶�
'      2.����ҽ��Ҫ��ǰ���ɺ����׶� mbln���ò�����=trueʱ,ҽ��վ����ʱ������Ƿ����ִ�еǼǣ������������ڼ��
'      3.��ʿվû����������,��Ҫÿ�����ɶ�Ҫ���ǰһ�ε�ִ�еǼ����
'-------------------------------------------------------------------------------------------
    Dim blnHave As Boolean       '������ִ�л��ڵļ��
    Dim blnReturn As Boolean
    Dim blnExePath As Boolean
    Dim blnUnExe As Boolean      '���ڱ��û��ִ��·��Ȩ���Ҵ��ڲ���Աִ�е�·����Ŀʱ,��Ҫ�����û���ʾ
    Dim strSubSQL As String
    Dim strSql As String
    Dim strTmp As String
    
    Dim strMsg As String
    Dim rsTmp As ADODB.Recordset
    
    Dim i As Long
    
    On Error GoTo errH
    
    blnHave = True 'Ĭ�ϼ��ִ�еǼ����
    blnExePath = InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";ִ��·��;") > 0
    blnReturn = True
    
    If mbln����ִ�л��� And mstrִ�г��� <> "00" Then
        If mint���� = 0 Then
            'ҽ��վ����
            If mstrִ�г��� = "11" Then
                strSubSQL = "And NVL(NVL(a.������,b.������),1)=1"
            ElseIf mstrִ�г��� = "10" Then
                strSubSQL = "And NVL(NVL(a.������,b.������),1)=1 And Nvl(Nvl(a.ִ����,b.ִ����),1)=1 "
            ElseIf mstrִ�г��� = "01" Then
                strSubSQL = "And NVL(NVL(a.������,b.������),1)=1 And Nvl(Nvl(a.ִ����,b.ִ����),1)=2 "
            End If
        ElseIf mint���� = 1 Then
            '��ʿվ��û���������ڣ�
            If mstrִ�г��� = "11" Or mstrִ�г��� = "01" Then
                strSubSQL = "And Nvl(Nvl(a.ִ����,b.ִ����),1)=2 "
            ElseIf mstrִ�г��� = "10" Then
                blnHave = False
            End If
        End If
    Else
        blnHave = False
    End If
    
    If blnHave Then
      
        strSql = "Select Nvl(b.��Ŀ����,a.��Ŀ����) ��Ŀ����,NVl(Nvl(a.ִ����,b.ִ����),1) as ִ���� From ����·��ִ�� a,�ٴ�·����Ŀ b " & vbNewLine & _
                        "Where a.��Ŀid=b.id(+) And a.·����¼ID = [1] And a.�׶�ID = [2] And a.���� = [3] And Nvl(a.����ʱ������,0)<>2 And a.ִ��ʱ�� Is null " & strSubSQL

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            If mint���� = 0 Then
                'ҽ�������������ڼ��,����ʱ�����
                If mstrִ�г��� = "11" Then
                    rsTmp.Filter = " ִ���� = 1"
                    If rsTmp.RecordCount > 0 Then
                        Call FuncGetRSTipInfo(rsTmp, "��Ŀ����", strTmp)
                        If blnExePath Then
                            If MsgBox("�ò��˻���δִ�е���Ŀ:" & vbCrLf & strTmp & vbCrLf & "������ִ�С�������Ҫ����ִ�в�����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, mint����) Then
                                    blnRefresh = True
                                    If Not CheckPathIsExecuted() Then
                                        blnReturn = False
                                        '�ٴμ�飬�п��ܴ��ڻ�ʿδִ�е���Ŀ
                                    End If
                                Else
                                    blnReturn = False
                                End If
                            Else
                                blnReturn = False
                            End If
                        Else
                            blnUnExe = True: blnReturn = False
                        End If
                    Else
                        rsTmp.Filter = " ִ���� = 2"
                        Call FuncGetRSTipInfo(rsTmp, "��Ŀ����", strTmp)
                        strMsg = "�ò��˻��л�ʿδִ�е���Ŀ:" & vbCrLf & strTmp & vbCrLf & "����ִ�к���ܼ�����"
                        blnReturn = False
                    End If
                ElseIf mstrִ�г��� = "01" Then
                    'ֻ�����������ҽ����ִ�����ǻ�ʿ��·����Ŀ
                    Call FuncGetRSTipInfo(rsTmp, "��Ŀ����", strTmp)
                    strMsg = "�ò��˻��л�ʿδִ�е���Ŀ:" & vbCrLf & strTmp & vbCrLf & "����ִ�к���ܼ�����"
                    blnReturn = False
                End If
            End If
            
            If mint���� = 1 Or (mint���� = 0 And mstrִ�г��� = "10") Then
                '��ʿ����
                Call FuncGetRSTipInfo(rsTmp, "��Ŀ����", strTmp)
                If (mint���� = 1 And (mstrִ�г��� = "11" Or mstrִ�г��� = "01")) Or (mint���� = 0 And mstrִ�г��� = "10") Then
                     If blnExePath Then
                        If MsgBox("�ò��˻���δִ�е���Ŀ:" & vbCrLf & strTmp & vbCrLf & "������ִ�С�������Ҫ����ִ�в�����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                            Call frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, mint����) '����ִ�еǼ�
                            blnRefresh = True
                        Else
                            blnReturn = False
                        End If
                     Else
                        blnUnExe = True: blnReturn = False
                     End If
                End If
            End If
            '
            If blnUnExe Then
                'û��ִ��·��Ȩ���Ҵ��ڲ���Աִ�е�·����Ŀʱ , ��Ҫ�����û���ʾ
                strMsg = "�ò��˻���δִ�е���Ŀ��" & vbCrLf & strTmp & vbCrLf & "����ִ�к���ܼ�����"
            End If
            
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End If
    
    CheckPathIsExecuted = blnReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncGetRSTipInfo(ByVal rsTmp As ADODB.Recordset, ByVal strFieldName As String, ByRef strTipInfo As String)
'------------------------------------------------------------------------------------------
'����:ѭ����ȡ��¼���м�Ҫ��Ϣ
'-----------------------------------------------------------------------------------------
    Dim i As Long
    
    strTipInfo = ""
    For i = 1 To rsTmp.RecordCount
        strTipInfo = IIf(i = 1, "", strTipInfo & vbCrLf) & rsTmp.Fields(strFieldName)
        If Len(strTipInfo) > 200 Then strTipInfo = strTipInfo & "��": Exit For
        rsTmp.MoveNext
    Next
End Sub

Private Function CheckPathSendByNurse(ByVal bytFunc As Byte, ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long, Optional ByVal dat���� As Date) As Boolean
'--------------------------------------------
'���ܣ���鵱ǰ�׶ε���Ŀ���Ƿ���ڻ�ʿ���ɵ���Ŀ
'����: bytFunc=1 ����ȡ��ִ�еǼ�   =2����ȡ��ִ�еǼ�
'      bytFunc=1ʱ lng·����¼ID ·����¼ID
'      bytFunc=2ʱ lng·����¼ID ·��ִ��ID
'����:T-����,F-������
'--------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim blnRet As Boolean
    
    On Error GoTo errH
    If bytFunc = 1 Then
        strSql = "Select 1 " & vbNewLine & _
                "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
                "Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And a.��Ŀid = b.Id(+) And Nvl(Nvl(a.������, b.������), 1) = 2 and rownum<2 "
    Else
        strSql = "Select 1 " & vbNewLine & _
               "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
               "Where a.ID=[1] And a.��Ŀid = b.Id(+) And Nvl(Nvl(a.������, b.������), 1) = 2 and rownum<2 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·����¼ID, lng�׶�ID, dat����)
    
    If rsTmp.RecordCount > 0 Then
        blnRet = True
    End If
    
    CheckPathSendByNurse = blnRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetPathCurrPhase(ByVal bytType As Byte, ByRef lng�׶�ID As Long, ByRef lng���� As Long, Optional ByRef strDate As String)
'--------------------------------------------------
'����:��ȡ����ִ�еǼǻ�����ȡ��ִ�еǼǵĵ�ǰ�׶�
'����:bytType =1 ����ִ��,=2����ȡ��ִ��
'-------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If bytType = 1 Then
        strSql = "Select *" & vbNewLine & _
            "From (Select Distinct a.�׶�id, a.����, a.����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��" & vbNewLine & _
            "       From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
            "       Where a.��Ŀid = b.Id(+) And a.·����¼id = [1] And Nvl(a.����ʱ������, 0) = 0 And Nvl(Nvl(a.ִ����, b.ִ����), 0) = " & IIf(mint���� = 0, 1, 2) & _
            "             And a.ִ��ʱ�� Is Null" & vbNewLine & _
            "       Group By a.�׶�id, a.����, a.����" & vbNewLine & _
            "       Order By Min(a.�Ǽ�ʱ��))" & vbNewLine & _
            "Where Rownum < 2"
    Else
        strSql = "Select *" & vbNewLine & _
            "From (Select Distinct a.�׶�id, a.����, a.����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��" & vbNewLine & _
            "       From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
            "       Where a.��Ŀid = b.Id(+) And a.·����¼id = [1] And Nvl(a.����ʱ������, 0) = 0 And Nvl(Nvl(a.ִ����, b.ִ����), 0) = " & IIf(mint���� = 0, 1, 2) & _
            "             And a.ִ��ʱ�� Is Not Null" & vbNewLine & _
            "       Group By a.�׶�id, a.����, a.����" & vbNewLine & _
            "       Order By Min(a.�Ǽ�ʱ��) Desc )  " & vbNewLine & _
            "Where Rownum < 2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount > 0 Then
        lng�׶�ID = rsTmp!�׶�ID
        strDate = rsTmp!���� & ""
        lng���� = rsTmp!����
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DefCommandPlugInPopup(ByVal objBar As Object, ByRef rsBar As ADODB.Recordset)
'���ܣ���ҽ�����Ҽ������˵�
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim objCtl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    '������ť
    rsBar.Filter = "IsInTool=1 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        For i = 1 To rsBar.RecordCount
            Set objControl = objBar.Add(xtpControlButton, rsBar!����ID, rsBar!������)
            objControl.IconId = rsBar!ͼ��ID
            objControl.Parameter = rsBar!������
            objControl.Style = xtpButtonIconAndCaption
            If Val(rsBar!IsGroup) = 1 Then
                objControl.BeginGroup = True
            End If
            rsBar.MoveNext
        Next
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        Set objPopup = objBar.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����")
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
End Sub

Private Function GetPlugInBar(ByVal lngģ�� As Long, ByVal int���� As Integer, rsBar As ADODB.Recordset) As String
'���ܣ���֯��Ҳ����Ĳ˵�����ť
    Dim strFunc As String
    Dim strXML As String
    Call CreatePlugInOK(lngģ��, int����)
    If gobjPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, lngģ��, int����, strXML)
    Call zlPlugInErrH(Err, "GetFuncNames")
    Err.Clear: On Error GoTo 0
    Call MakePlugInBar(strFunc, strXML, rsBar)
    GetPlugInBar = strFunc
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'���ܣ���֯�˵������ؼ�¼���У�ע����ϰ汾�ļ��ݴ���
'������strFunc �ϰ汾�����д���strXML��������Ϣ�Ĺ��ܴ�
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    Dim rsBarFuncID As ADODB.Recordset
    
    If strXML = "" And strFunc = "" Then Exit Sub
    If strXML = "" And strFunc <> "" Then
        '������ǰ�ϰ汾�ķ�ʽ
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '�ݶ�Ϊ200����չ���ܲ������ֹ��ѭ��
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'���ܣ������ܴ�ת��Ϊ��¼����ʽ
'������strFunc ���ܴ���intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '��һ��������ť��ʾ�ָ���
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !������ = strFuncName
            !�˵��� = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'���ܣ����书��ID���Ӳ˵����
'������lngV �汾��1-�ϰ棬2-�°�
'���أ��ַ�������ǰ�Ͱ汾��ʽ�Ĺ��ܴ�
    Dim i As Long
    '���书��ID��ͼ��ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !��� = i
            !����ID = conMenu_Tool_PlugIn_Item + i
            !ͼ��ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'���ܣ��趨���
'������lngV �汾��1-�ϰ棬2-�°� intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '���ֻ��һ����Ҳ��Ϊ������ť
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !�˵��� = !�˵��� & "(&" & i & ")"
                    Else
                        !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !�˵��� = !�˵��� & "(&" & i & ")"
                Else
                    !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "���", adBigInt '��������
    rsBar.Fields.Append "����ID", adBigInt '�˵���ť Control.ID
    rsBar.Fields.Append "ͼ��ID", adBigInt
    rsBar.Fields.Append "������", adVarChar, 1000 'ȥ���ؼ���֮��� ���� ���������ϵİ�ť����
    rsBar.Fields.Append "�˵���", adVarChar, 1000 '�˵���/�Ҽ��˵� ����
    rsBar.Fields.Append "IsAuto", adInteger '�Ƿ��Զ�ִ�й���
    rsBar.Fields.Append "IsGroup", adInteger '�Ƿ�ָ���
    rsBar.Fields.Append "IsInTool", adInteger '�Ƿ������ʾ
    rsBar.Fields.Append "BarType", adInteger '1-�˵�����2����������3��������
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub

Private Sub FuncPathTableChange(ByRef vsBody As VSFlexGrid, ByVal lngPageCOL As Long, Optional vsHead As VSFlexGrid)
'����:����ӡ��ת���ɹ̶���,���ڴ�ӡ�����
'��Ҫ�������:89612-���׶��и߳�����ӡ��Ч��ΧʱҪ����һҳ��������ǰ�׶�ʣ����
'            80442-ÿһ�׶ε������Զ������м��,�޳��հ��С�
'����:
'����:vsBody��ӡ����
'���:lngPageCOL ��ӡ����(�����̶���)
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error Resume Next
    Load vsPathPrint(1)
    Err.Clear: On Error GoTo 0
    
    With vsPathPrint(1)
        '���
        .Rows = 0
        .Cols = 0
        
        If lngPageCOL = 0 Then Exit Sub
        If (vsBody.Cols - vsBody.FixedCols) Mod lngPageCOL <> 0 Then Exit Sub
        '
        .Rows = ((vsBody.Cols - vsBody.FixedCols) / lngPageCOL) * vsBody.Rows
        .Cols = vsBody.FixedCols + lngPageCOL
        .FixedCols = vsBody.FixedCols
        .FixedRows = vsBody.FixedRows
        
        '��vsBody�����ݸ��Ƶ�vsPathPrint(1)
        '�̶���
        For lngCol = 0 To .FixedCols
            lngRow = 0
            Do
                '��ԭ���ǹ̶���ת���ɷǹ̶���ʱ��Ҫ�����Ǳ��ڴ�ӡ����ʶ��
                If lngRow Mod vsBody.Rows < vsBody.FixedRows And lngRow >= vsBody.FixedRows And lngCol = 0 Then
                    .RowData(lngRow) = UCase("FIXEDROW")
                End If
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next
        '�ǹ̶���
        For lngCol = .FixedCols To (.FixedCols + lngPageCOL) - 1
            lngRow = 0
            Do
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, (lngPageCOL * (lngRow \ vsBody.Rows)) + lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next
        
        '��ն��ж��ǿհ׵���
        For lngRow = 0 To .Rows - 1
            For lngCol = 1 To lngPageCOL
                If .RowData(lngRow) = UCase("FIXEDROW") Then
                    .Cell(flexcpAlignment, lngRow, lngCol, lngRow, .Cols - 1) = flexAlignCenterCenter
                    Exit For
                ElseIf .TextMatrix(lngRow, 0) = "��ʿǩ��" Or .TextMatrix(lngRow, 0) = "ҽ��ǩ��" Then
                    Exit For
                ElseIf .TextMatrix(lngRow, lngCol) <> "" Then
                    Exit For
                ElseIf lngCol = lngPageCOL Then
                    '��¼��Ҫɾ���Ŀհ���
                   .RemoveItem lngRow
                   lngRow = lngRow - 1  'ɾ��һ��,��һ���������
                End If
            Next
            If lngRow = .Rows - 1 Then Exit For
        Next
        '��ʾ�����
        .MergeCol(0) = True
        '�趨���壬���
        .FontSize = IIf(mlngFontSize = 0, CON_SmallFontSize, mlngFontSize) '·������������ӡmlngFontSize=0
        '��ʾ���
        .Cell(flexcpAlignment, 0, .FixedCols, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
    End With
    
    
    Set vsBody = vsPathPrint(1)
End Sub

Private Sub FuncPathCellCopy(ByRef vsSource As VSFlexGrid, ByRef vsCopy As VSFlexGrid, _
        ByVal lngSourRow As Long, ByVal lngSourCol As Long, ByVal lngCopyRow As Long, ByVal lngCopyCol As Long)
'����:���Ƶ�Ԫ��
'������vsSource-��Copy�ı�
'    vsCopy-copy��ı�
'    lngSourRow ,lngSourCol ��Copy�ı���Ӧ�к���
'    lngCopyRow��lngCopyCol  Copy�����Ӧ�к���
    With vsCopy
        .Cell(flexcpText, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpText, lngSourRow, lngSourCol)
        .Cell(flexcpAlignment, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpAlignment, lngSourRow, lngSourCol)
        .Cell(flexcpBackColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpBackColor, lngSourRow, lngSourCol)
        .Cell(flexcpForeColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpForeColor, lngSourRow, lngSourCol)
        .Cell(flexcpPicture, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpPicture, lngSourRow, lngSourCol)
    End With
End Sub

Public Function GetFormOperation() As String
'���ܣ���ȡ�������ѡ�񣬸ýӿڻ��ڴ���ж��ǰ���ã��°滤ʿվ �������񴰿�
'���أ���¼��ǰ�����пؼ�ѡ��״̬
 
    Dim strXML As String
    Dim lngIdx As Long
     
    If optSelect(IX_ALL).Value Then
        lngIdx = IX_ALL
    ElseIf optSelect(IX_ҽ��).Value Then
        lngIdx = IX_ҽ��
    ElseIf optSelect(IX_��ʿ).Value Then
        lngIdx = IX_��ʿ
    End If
    strXML = "<root><scz>" & lngIdx & "</scz></root>"  '������
    GetFormOperation = strXML
End Function

Public Function RestoreFormOperation(ByVal strValue As String)
'���ܣ��ָ��������ѡ��
'������strValue ǰ�����пؼ�ѡ��״̬

    Dim objXML As New zl9ComLib.clsXML
    Dim strTmp As String
     Dim lngIdx As Long
    
    On Error Resume Next
    
    Call objXML.OpenXMLDocument(strValue)
    
    Call objXML.GetSingleNodeValue("scz", strTmp) 'Ӥ��
    lngIdx = Val(strTmp)

    Set mcolReason = New Collection
    optSelect(lngIdx).Value = True
End Function