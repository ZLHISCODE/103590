VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeeQuery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picNum 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8520
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   33
      Top             =   30
      Width           =   1095
      Begin VB.ComboBox cboNum 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   -30
         Width           =   1185
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid mshInsure 
      Height          =   1095
      Left            =   4980
      TabIndex        =   27
      Top             =   945
      Width           =   3660
      _cx             =   6456
      _cy             =   1931
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      ExplorerBar     =   5
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
   Begin VSFlex8Ctl.VSFlexGrid mshDepost 
      Height          =   1260
      Left            =   30
      TabIndex        =   26
      Top             =   915
      Width           =   4545
      _cx             =   8017
      _cy             =   2222
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      ExplorerBar     =   5
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
   Begin VB.PictureBox picLR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   5040
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1215
      ScaleWidth      =   45
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox pic������Ϣ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   105
      ScaleHeight     =   315
      ScaleWidth      =   7035
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   315
      Width           =   7035
      Begin VB.Label lbl������Ϣ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   135
         TabIndex        =   21
         Top             =   75
         Width           =   900
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   19
      Top             =   2130
      Width           =   7275
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   60
      ScaleHeight     =   780
      ScaleWidth      =   10830
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2235
      Width           =   10830
      Begin VB.Frame fraDeptMode 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   0
         Width           =   2295
         Begin VB.OptionButton optDeptMode 
            Caption         =   "ִ�п���"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optDeptMode 
            Caption         =   "��������"
            Height          =   255
            Index           =   0
            Left            =   100
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame fraTypeMode 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   0
         Width           =   2295
         Begin VB.OptionButton optTypeMode 
            Caption         =   "�վݷ�Ŀ"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   4
            Top             =   0
            Width           =   1485
         End
         Begin VB.OptionButton optTypeMode 
            Caption         =   "������Ŀ"
            Height          =   255
            Index           =   0
            Left            =   100
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��"
         Height          =   315
         Left            =   6420
         TabIndex        =   16
         Top             =   345
         Width           =   630
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   367
         Width           =   1230
      End
      Begin VB.CheckBox chkAdivce 
         Caption         =   "��ҽ������"
         Height          =   300
         Left            =   6075
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   2415
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboFeeType 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.TabStrip tabTime 
         Height          =   315
         Left            =   7560
         TabIndex        =   11
         Top             =   -15
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         Style           =   2
         TabFixedHeight  =   526
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   882
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "���з���"
               Key             =   "All"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   2130
         TabIndex        =   14
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   271253507
         CurrentDate     =   36257
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   271253507
         CurrentDate     =   36257.9999884259
      End
      Begin VB.CheckBox chkNotCheckFee 
         Caption         =   "������������"
         Height          =   210
         Left            =   9675
         TabIndex        =   32
         Top             =   412
         Width           =   1575
      End
      Begin VB.CheckBox chk����ʾ���ʵ��� 
         Caption         =   "����ʾ���ʵ���"
         Height          =   210
         Left            =   8850
         TabIndex        =   17
         Top             =   412
         Width           =   1665
      End
      Begin VB.CheckBox chk��������С�� 
         Caption         =   "��������С��"
         Height          =   210
         Left            =   7035
         TabIndex        =   31
         Top             =   412
         Width           =   2115
      End
      Begin VB.Label lbl���ڷ�Χ 
         AutoSize        =   -1  'True
         Caption         =   "2009-01-01 00:00:00��2009-02-02 23:59:59"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2070
         TabIndex        =   30
         Top             =   420
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.Label lbl�� 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3885
         TabIndex        =   29
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   75
         TabIndex        =   12
         Top             =   420
         Width           =   720
      End
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   10410
         Picture         =   "frmFeeQuery.frx":0000
         ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ���ò�ѯ��ʽ��"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   -15
         TabIndex        =   18
         ToolTipText     =   "F2:ѡ������嵥����"
         Top             =   60
         Width           =   1350
      End
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   315
      Left            =   90
      TabIndex        =   22
      Top             =   0
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   556
      Style           =   2
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "δ�����"
            Key             =   "Main"
            Object.ToolTipText     =   "δ������嵥"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFee 
      Height          =   1260
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   4545
      _cx             =   8017
      _cy             =   2222
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsTools 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblDepost 
      BackColor       =   &H00808080&
      Caption         =   " Ԥ�����嵥"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   690
      Width           =   4050
   End
   Begin VB.Label lblInsure 
      BackColor       =   &H00808080&
      Caption         =   " ����Ԥ�����"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   2010
   End
End
Attribute VB_Name = "frmFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '���Ѽ���ʱ
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������

Private mint���� As Integer '0-���ò�ѯ��1-��ʿվ����
Private mlng����ID As Long, mbln���� As Boolean '33744

Private mcbsMain As CommandBars
Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1

Private mstrUnitIDs As String   '����Ա�����Ĳ�����

Private msngScale As Single
Private Const mlngModul = 1139
Private mrsList As ADODB.Recordset '��¼��ǰ�����嵥
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstrסԺ�� As String
Private mlng����ID As Long
Private mintInsure As Integer
Private mblnDateMoved As Boolean
Private mbln��Ժ As Boolean
Private mbln���� As Boolean
Private mbln�������۲��� As Boolean
Private mblnHavePara As Boolean

Private mintPreCard As Integer
Private mintPreTime As Integer
Private mintPreTimeIndex As Integer

Private mbytList As Byte                'δ������嵥�ͽ����嵥��ѯ�������
Private mblnClinicOrNurse As Boolean     '��ǰ����Ա��ȱʡ�����Ƿ����ٴ�������

Private mbytDateType As Byte '1-����ʱ��,2-�Ǽ�ʱ��
Private mblnPreBalance As Boolean   '�Ƿ�����Ԥ��ҽ��
Private mblnUnBilling As Boolean '�Ƿ������
Private mstrPrivs As String
Private mstr��ֹ���� As String
Private mblnNotClick As Boolean '��ִ����ص�ѡ������

Private Enum ListType
    C0�����嵥 = 0
    C1�ֿ�����ϸ = 1
    C2����Ŀ��ϸ = 2
    C3�������ϸ = 3
    C4���������ϸ = 4
    
    C5����Ŀ���� = 5
    C6�������� = 6
    C7���·������ = 7
    C8���յ��ݻ��� = 8
    C9���շ�Ŀ���� = 9
End Enum

Private Const conTabδ�� = 1

Private Type t_ViewState
    ReBalance As Boolean
    ZeroFee As Boolean
    CheckFee As Boolean
End Type
Private mobjInPati As Object
Private mvs As t_ViewState
Private mbytFontSize As Byte
Private mblnFisrtSetFontSize As Boolean '��һ�����������С
Private mcllBalaceNums As Collection
Private mstrRestoreFeeCons As String
Private mblnContainOutFee As Boolean '�Ƿ�����������

Public Sub SetFontSize(ByVal bytSize As Byte)
      '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2012-06-18 16:50:35
    '����:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub
Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������С
    '����:���˺�
    '����:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("��") + 20
        Case UCase("VsFlexGrid")
            Call zlControl.VSFSetFontSize(objCtrl, mbytFontSize)
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
            
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    Call Form_Resize
    Call picDetail_Resize
    '����:55392
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False, , mblnHavePara
 End Sub

Private Sub InitBaseData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:24913
    '����:2009-08-17 09:47:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType  As String, strBeginDate As String, strEndDate As String, varData As Variant, intTYPE As Integer, i As Long
    With cbo����
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .AddItem "����"
        .ItemData(.NewIndex) = 2
        .AddItem "����"
        .ItemData(.NewIndex) = 3
        .AddItem "����"
        .ItemData(.NewIndex) = 4
        .AddItem "�Զ��巶Χ"
        .ItemData(.NewIndex) = 5
    End With
    strType = zlDatabase.GetPara("���ò�ѯ��Χ", glngSys, mlngModul, "����", Array(cbo����), InStr(1, mstrPrivs, ";��������;") > 0)
    mblnHavePara = InStr(1, mstrPrivs, ";��������;") > 0
       
    varData = Split(strType & "|", "|"): strType = varData(0)
    intTYPE = Switch(strType = "����", 0, strType = "����", 1, strType = "����", 2, strType = "����", 3, strType = "����", 4, True, 5)
    If intTYPE = 5 Then
       varData = Split(varData(1) & ",", ",")
       If varData(0) <> "" And IsDate(varData(0)) Then dtpBegin.Value = Format(CDate(varData(0)), "yyyy-mm-dd 00:00:00")
       If varData(1) <> "" And IsDate(varData(1)) Then dtpEnd.Value = CDate(Format(CDate(varData(1)), "yyyy-mm-dd") & " 23:59:59")
    End If
    For i = 0 To cbo����.ListCount - 1
        If cbo����.ItemData(i) = intTYPE Then
            cbo����.ListIndex = i: Exit For
        End If
    Next
    If cbo����.ListIndex < 0 Then cbo����.ListIndex = 0
    '46646
    chkNotCheckFee.Value = IIf(Val(zlDatabase.GetPara("����������", glngSys, mlngModul, "0", Array(chkNotCheckFee), InStr(1, mstrPrivs, ";��������;") > 0)) = 1, 1, 0)
    
    
    strType = zlDatabase.GetPara("��ϸ�������ʵ���", glngSys, mlngModul, "0", Array(chk����ʾ���ʵ���), InStr(1, mstrPrivs, ";��������;") > 0)
    chk����ʾ���ʵ���.Value = IIf(Val(strType) = 0, 0, 1)
   
    chk����ʾ���ʵ���.Visible = False
    strType = zlDatabase.GetPara("��������ͳ��", glngSys, mlngModul, "0", Array(chk��������С��), InStr(1, mstrPrivs, ";��������;") > 0)
    '����:41673
    chk��������С��.Value = IIf(Val(strType) = 0, 0, 1)
    chk��������С��.Visible = mbytList = ListType.C2����Ŀ��ϸ
    
    Select Case mbytList
        Case ListType.C0�����嵥, ListType.C1�ֿ�����ϸ, ListType.C2����Ŀ��ϸ, ListType.C3�������ϸ, ListType.C4���������ϸ '��ϸ�嵥,�ֿ���ϸ,��Ŀ��ϸ,������ϸ,(��������Ŀ(���վݷ�Ŀ),�շ���Ŀ,��ϸ�ּ���ѯ)
            chk����ʾ���ʵ���.Visible = True
            
        Case ListType.C5����Ŀ����  '��Ŀ����
        Case ListType.C6��������  '�������
        Case ListType.C7���·������  '���»���
        Case ListType.C8���յ��ݻ���  '���շ���
        Case ListType.C9���շ�Ŀ����  '���շ�Ŀ
    End Select
End Sub
Private Sub SetDateVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ڿؼ���visible����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:24913
    '����:2009-08-17 10:10:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String, intTYPE As Integer, blnDateVisible As Boolean, strBeginDate As String, strEndDate As String
    
    strType = cbo����.Text
    intTYPE = Switch(strType = "����", 0, strType = "����", 1, strType = "����", 2, strType = "����", 3, strType = "����", 4, True, 5)
    blnDateVisible = intTYPE = 5
    dtpBegin.Visible = blnDateVisible: dtpEnd.Visible = blnDateVisible: lbl��.Visible = blnDateVisible
    cmdRefresh.Visible = blnDateVisible
    
    lbl���ڷ�Χ.Visible = (Not blnDateVisible) And intTYPE <> 0
    
    If lbl���ڷ�Χ.Visible Then
        zlGetDateRange , strBeginDate, strEndDate
        lbl���ڷ�Χ.Caption = strBeginDate & "��" & strEndDate
    End If
End Sub

Private Sub zlGetDateRange(Optional ByVal blnOnlyDate As Boolean = False, Optional ByRef strBeginDate As String, Optional ByRef strEndDate As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ڷ�Χ
    '���:blnOnlyDate-true:��Ϊ����(2009-01-01),����Ϊ�����������(20009-01-01 23:59:59)
    '����:strBeginDate-��ʼ����,strEndDate-��������
    '����:
    '����:���˺�
    '����:2009-08-17 10:23:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    Dim blnTime As Boolean
    strType = cbo����.Text
    blnTime = False
    Select Case strType
    Case "����"
        strBeginDate = "": strEndDate = ""
    Case "����"
         strBeginDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
         strEndDate = strBeginDate
    Case "����"
         strBeginDate = Format(DateAdd("d", -1, zlDatabase.Currentdate), "yyyy-mm-dd"): strEndDate = strBeginDate
    Case "����"
        strEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd"): strBeginDate = Format(DateAdd("d", -1 * (Weekday(CDate(strEndDate), vbSunday) - 1), CDate(strEndDate)), "yyyy-mm-dd")
    Case "����"
        strEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd"): strBeginDate = Format(CDate(strEndDate), "yyyy-mm") & "-01"
    Case Else
        strBeginDate = Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"): strEndDate = Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
        blnTime = True
    End Select
    If blnOnlyDate = False And strBeginDate <> "" And Not blnTime Then
          strBeginDate = strBeginDate & " 00:00:00": strEndDate = strEndDate & " 23:59:59"
    End If
End Sub

Private Sub RefreshAllData()
'���ܣ�ˢ������
    Dim blnMCPatient As Boolean, i As Integer
    
    '�����Ƿ�ҽ��������ʾ����ģ��������
    If tabClass.SelectedItem.Index = conTabδ�� Then blnMCPatient = mintInsure <> 0
    
            
    '��ȡ����ģ���������嵥
    '���˺�:����        mblnPreBalance = True:ԭ����control.enabled���ܲ�Ϊtrue
    '25657:
    If blnMCPatient Then
        Call ReadInsureMoney(mlng����ID, mlng��ҳID)
        mblnPreBalance = True
    Else
        mshInsure.Clear
        mshInsure.Rows = 2
        mblnPreBalance = True
    End If
    lblInsure.Visible = blnMCPatient
    mshInsure.Visible = blnMCPatient
    picLR.Visible = blnMCPatient
    Call Form_Resize
    
    Call LoadPatientBaby(cboBaby, mlng����ID, mlng��ҳID)
    cboBaby.Visible = cboBaby.ListCount > 1
    If cboBaby.ListCount > 1 Then
        cboBaby.AddItem "���˺�Ӥ��"
        cboBaby.ItemData(cboBaby.NewIndex) = 999
        Call zlControl.CboSetIndex(cboBaby.hWnd, cboBaby.NewIndex)
    End If
    zlControl.CboSetWidth cboBaby.hWnd, cboBaby.Width * 2
    
    
    If LoadPatiClass Then '��ʼ��ѡ�
        Call ChangeList(False)  '����ʱ
        If mstrRestoreFeeCons <> "" Then
            For i = 1 To tabClass.Tabs.Count
                If tabClass.Tabs(i).Key = Nvl(Split(mstrRestoreFeeCons, "|")(1)) Then tabClass.Tabs(i).Selected = True
            Next
        End If
        If i = 0 Then tabClass.Tabs(1).Selected = True '����tabClass_Click��ʾ����
    End If
          
    Call SetCondition '��Ҫ����ΪӤ���ѿ��ܱ仯
    If mstrRestoreFeeCons <> "" Then
        If zlRestorePosition(mlng����ID) = False Then Exit Sub
    End If
End Sub


Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strסԺ�� As String, ByVal lng����ID As Long, _
    ByVal intInsure As Integer, ByVal blnDateMoved As Boolean, ByVal bln��Ժ As Boolean, ByVal bln���� As Boolean, _
    blnOnlyRefreshVar As Boolean, _
    Optional bln���� As Boolean = False, Optional lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:bln��Ժ-�Ƿ��Ժ����
    '       bln����-��ת�ƣ�ת�����Ĳ��˽��в���
    '       lng����ID-������Ϊtrueʱ,���뱾����Ҫ���ѵĿ���ID
    '       lng����ID-������Ϊtrueʱ,���뱾����Ҫ���ѵĲ���ID
    '       blnOnlyRefreshVar-��ˢ���ڲ�����
    '����:
    '����:
    '����:���˺�
    '����:2010-12-10 14:43:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrסԺ�� = strסԺ��
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mbln���� = bln����
    mintPreTimeIndex = 0
    mintPreTime = 0
    mintInsure = intInsure
    mblnDateMoved = blnDateMoved
    mbln��Ժ = bln��Ժ
    mbln���� = bln����
    mbln�������۲��� = ZlIsOutpatientObserve(lng����ID, lng��ҳID)
    If blnOnlyRefreshVar Then Exit Sub
    tabClass.Tag = ""
    Call RefreshAllData
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strBillingPrivs As String
    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub
    blnVisible = True
    
    strBillingPrivs = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    Select Case Control.ID
        Case conMenu_File_PrintPageSet
            blnVisible = InStr(";" & GetInsidePrivs(Enum_Inside_Program.p���ò�ѯ), ";������ҳ") > 0
        Case conMenu_File_PrintMultiBill, conMenu_File_PrintSingleBill
            blnVisible = InStr(";" & GetInsidePrivs(Enum_Inside_Program.p���ò�ѯ), ";�߿��ӡ") > 0
        Case conMenu_Edit_PreBalanceAll
            blnVisible = InStr(";" & GetInsidePrivs(Enum_Inside_Program.p���ò�ѯ), ";Ԥ�����в���") > 0
        Case conMenu_Edit_Billing, conMenu_Edit_Copy, conMenu_Edit_Billing_Mulit
            '54274
            blnVisible = InStr(strBillingPrivs, "סԺ����") > 0
        Case conMenu_Edit_CardBackMoney
            blnVisible = InStr(";" & GetInsidePrivs(9000), ";��Ժ��������˿�;") > 0 Or InStr(";" & GetInsidePrivs(9000), ";��Ժ��������˿�;") > 0
        Case conMenu_Edit_ReBilling
            '55380
            blnVisible = InStr(strBillingPrivs, ";ҩƷ����;") > 0 _
                Or InStr(strBillingPrivs, ";��������;") > 0 _
                Or InStr(strBillingPrivs, ";��������;") > 0
        Case conMenu_Edit_ReBillingApply
            '55380
            blnVisible = (InStr(strBillingPrivs, ";ҩƷ��������;") > 0 _
                Or InStr(strBillingPrivs, ";������������;") > 0 _
                Or InStr(strBillingPrivs, ";������������;") > 0) _
                And InStr(strBillingPrivs, "��������") > 0
                
        Case conMenu_Edit_ReBillingAudit
            blnVisible = InStr(strBillingPrivs, "�������") > 0
        Case conMenu_Edit_ReBillingButton
            '55380
            blnVisible = InStr(strBillingPrivs, "�������") > 0 _
                Or ((InStr(strBillingPrivs, ";ҩƷ��������;") > 0 _
                        Or InStr(strBillingPrivs, ";������������;") > 0 _
                        Or InStr(strBillingPrivs, ";������������;") > 0) And InStr(strBillingPrivs, "��������") > 0)
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub

Private Function GetPatiInsure() As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select A.�Ǽ�ʱ��, B.����, E.����, Nvl(E.ҽ����, D.��Ϣֵ) ҽ����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ҽ�����˵��� E, ҽ�����˹����� F" & vbNewLine & _
            "Where B.����id = [1] And B.��ҳid = [2] And A.����id = B.����id And B.����id = D.����id(+) And B.��ҳid = D.��ҳid(+) And D.��Ϣ��(+) = 'ҽ����' And" & vbNewLine & _
            "      A.����id = F.����id(+) And F.��־(+) = 1 And F.ҽ���� = E.ҽ����(+) And F.���� = E.����(+) And F.���� = E.����(+)"
    On Error GoTo errH
    Set GetPatiInsure = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ExecPreBalance()
    Dim int���� As Integer
    Dim strҽ���� As String, str���� As String
    Dim rsTmp As ADODB.Recordset, str������� As String
    Dim blnDateMoved As Boolean, dat�Ǽ�ʱ�� As Date
    
    Set rsTmp = GetPatiInsure
    If rsTmp.RecordCount > 0 Then
    With rsTmp
        int���� = Val(!����)
        strҽ���� = "" & !ҽ����
        str���� = "" & !����
        dat�Ǽ�ʱ�� = !�Ǽ�ʱ��
    End With
    End If
    If int���� = 0 Then
        MsgBox "��ȡ����ҽ�������Ϣʧ��!", vbExclamation, gstrSysName
        Exit Sub
    End If
    If gclsInsure.GetCapability(support����_�������ú���ýӿ�, mlng����ID, mintInsure) Then
        MsgBox "��ҽ���ӿڲ�֧�ֽ�������ǰԤ����!", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    blnDateMoved = zlDatabase.DateMoved(dat�Ǽ�ʱ��, , , Caption)
    
    Screen.MousePointer = 11
    Set rsTmp = GetVBalance(1, "סԺ���ý���", int����, mlng����ID, , , , , blnDateMoved)
    Screen.MousePointer = 0
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò���û��δ���ʵı�����Ŀ����!", vbInformation, gstrSysName
    Else
        str������� = gclsInsure.WipeoffMoney(rsTmp, mlng����ID, strҽ����, "0", int����, "|0") '������;����
        MsgBox "Ԥ����ɹ�!" & str�������, vbInformation, gstrSysName '�ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
        Call RefreshAllData
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim i As Long, objControl As CommandBarControl
    Select Case Control.ID
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Preview
            PrintList 2
        Case conMenu_File_Print
            PrintList 1
        Case conMenu_File_Excel
            PrintList 3
        Case conMenu_File_PrintBedCard
            Call zlPrintBedCard(Me, mlng����ID, mlng��ҳID)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_PrintSingleBill
            Call zlExecPrintSingleBill(Me, mlng����ID, mstr��ֹ����)
        Case conMenu_File_PrintDayDetail
            Call zlPrintDayDetail(Me, mint����, mlng����ID, mlng����ID, mvs.ReBalance, mvs.ZeroFee, mbytDateType = 1, mlng��ҳID)
        Case conMenu_File_PrintPageSet  '��ӡ��ҳ����
            Call zlPrintAccountPage(Me)
        Case conMenu_Edit_PreBalance     'Ԥ����
            If zlPreBalance(Me, mlng����ID, mlng��ҳID) = True Then RefreshAllData
        Case conMenu_Edit_PreBalanceAll      'Ԥ��������
            Call zlPreBalanceAll(Me, mlng����ID)
        Case conMenu_Edit_PatiMemo '�޸Ĳ��˱�ע��Ϣ
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, mlng����ID, mlng��ҳID, mobjInPati, False)
        Case conMenu_Edit_Billing   '����
            '����:33744
            If zlExecBilling(IIf(mint���� = 0, 6, mint����), gfrmMain, mlng����ID, mlng����ID, mbln��Ժ, mbln����, _
                mstrUnitIDs, mlng��ҳID, mbln����, mlng����ID, , mbln�������۲���) Then Call RefreshAllData
        Case conMenu_Edit_Billing_Mulit '��������
            If zlExecBilling_Mulit(IIf(mint���� = 0, 6, mint����), gfrmMain, mlng����ID, mlng����ID, mbln��Ժ, mbln����, _
                mstrUnitIDs, mlng��ҳID, mbln����, mlng����ID) Then Call RefreshAllData
        Case conMenu_Edit_Copy '���Ƽ��˵�
            '54274
            If zlCopyBill(IIf(mint���� = 0, 6, mint����), gfrmMain, mlng����ID, mlng����ID, mbln��Ժ, mbln����, _
                mstrUnitIDs, mlng��ҳID, mlng����ID, mbln�������۲���) Then Call RefreshAllData
        Case conMenu_Edit_ReBilling '����
            Call ExecUnBilling
        Case conMenu_Edit_CardBackMoney '����˿�
            Call NurseDeposit(mfrmParent, mlng����ID, mlng��ҳID, True, IIf(mbln�������۲���, 1, 2))
        Case conMenu_Edit_ReBillingApply
            If vsfFee.ColIndex("���ݺ�") = -1 Then
                If zlWrite_Off_ApplyAndVerfy(mfrmParent, mlng����ID, mlng����ID, Control.ID = conMenu_Edit_ReBillingApply) = True Then
                    RefreshAllData
                End If
            Else
                If zlWrite_Off_ApplyAndVerfy(mfrmParent, mlng����ID, mlng����ID, Control.ID = conMenu_Edit_ReBillingApply, vsfFee.TextMatrix(vsfFee.Row, vsfFee.ColIndex("���ݺ�"))) = True Then
                    RefreshAllData
                End If
            End If
        Case conMenu_Edit_ReBillingAudit
            If zlWrite_Off_ApplyAndVerfy(mfrmParent, mlng����ID, mlng����ID, Control.ID = conMenu_Edit_ReBillingApply) = True Then RefreshAllData
        Case conMenu_View_DateType * 10 + 1, conMenu_View_DateType * 10 + 2 'ʱ��ģʽ
            mbytDateType = Control.ID - conMenu_View_DateType * 10
            lbl����ʱ��.Caption = IIf(mbytDateType = 1, "����ʱ��", "�Ǽ�ʱ��")
            Call LoadCardData(False, False, True)
        
        Case conMenu_View_DetailType * 10 To conMenu_View_DetailType * 10 + 9 '��ѯ��ʽ'
            
            '�����ϴ�ѡ��Ľ��
            zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False
            
            mbytList = Control.ID - conMenu_View_DetailType * 10
            chk����ʾ���ʵ���.Visible = False
            chk��������С��.Visible = mbytList = ListType.C2����Ŀ��ϸ
            Select Case mbytList
                Case ListType.C0�����嵥, ListType.C1�ֿ�����ϸ, ListType.C2����Ŀ��ϸ, ListType.C3�������ϸ, ListType.C4���������ϸ '��ϸ�嵥,�ֿ���ϸ,��Ŀ��ϸ,������ϸ,(��������Ŀ(���վݷ�Ŀ),�շ���Ŀ,��ϸ�ּ���ѯ)
                    chk����ʾ���ʵ���.Visible = True
                Case ListType.C5����Ŀ����  '��Ŀ����
                Case ListType.C6��������  '�������
                Case ListType.C7���·������  '���»���
                Case ListType.C8���յ��ݻ���  '���շ���
                Case ListType.C9���շ�Ŀ����  '���շ�Ŀ
            End Select
            Call ChangeList(True)
            Call picDetail_Resize
            
        Case conMenu_View_ReBalance '��ʾ��������
            Control.Checked = Not Control.Checked: mvs.ReBalance = Control.Checked
            Call RefreshAllData
        Case conMenu_View_ZeroFee   '��ʾ�����
            Control.Checked = Not Control.Checked: mvs.ZeroFee = Control.Checked
            Call LoadCardData(False, False, True)
        Case conMenu_View_CheckFee  '��ʾ������
            Control.Checked = Not Control.Checked: mvs.CheckFee = Control.Checked
            Call LoadCardData(True, True, True)
        Case conMenu_View_TurnToWardFeeQuery 'ת�������ñ䶯��ѯ
            If CreatePublicExpenseBillOperation() Then
                Call gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(Me, 3, mlng����ID, mlng��ҳID)
            End If
        Case conMenu_View_ToolBar_Button '������
            For i = 1 To cbsTools.Count
                cbsTools(i).Visible = Not cbsTools(i).Visible
            Next
            cbsTools.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            Control.Checked = Not Control.Checked
            For i = 1 To cbsTools.Count
                For Each objControl In cbsTools(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            cbsTools.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            cbsTools.Options.LargeIcons = Not cbsTools.Options.LargeIcons
            cbsTools.RecalcLayout
        Case conMenu_View_PatInfor  '�鿴���˿�Ƭ
            Call ShowPatiCard
        Case conMenu_View_Billing   '�鿴���ʵ�
            Call vsfFee_DblClick
        Case conMenu_View_Refresh
            mintPreCard = 0: mintPreTime = 0
            Call RefreshAllData
        Case conMenu_Tool_Option    '����ѡ��
            frmSetExpence.mlngModul = 1133
            frmSetExpence.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
            frmSetExpence.mbytInFun = 0
            frmSetExpence.mbytUseType = 1   '����סԺ���ʹ���ģ�����
            frmSetExpence.Show 1, Me
        Case conMenu_View_ContainOutFee
            Control.Checked = Not Control.Checked
            mblnContainOutFee = IIf(Control.Checked, 1, 0)
            Call LoadCardData(True, True, True)
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                'ִ�з�������ǰģ��ı���
                If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then '�߿��(��ʹû����ʾ����Ҳ����ʹ��)
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                ElseIf Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then 'סԺ�����ձ�(��ʿվ���ò���)
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                             "����=" & mlng����ID, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID)
                Else
                    If mlng����ID = 0 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "����=" & mlng����ID)
                    Else
                        Dim lng����ID As Long
                        lng����ID = Val(IIf(tabClass.SelectedItem.Index = 1, 0, tabClass.SelectedItem.Tag))
                        If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then  '������ҳ
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "סԺ��=" & mstrסԺ��, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, _
                             "����=" & mlng����ID, "����ID=" & lng����ID)
                        Else
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "סԺ��=" & mstrסԺ��, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, _
                             "����=" & mlng����ID, "����ID=" & lng����ID)
                        End If
                    End If
                End If
          End If
    End Select
End Sub

Private Sub ShowPatiCard()
    frmDegreeCard.mlng����ID = mlng����ID
    frmDegreeCard.mlng��ҳID = mlng��ҳID
    frmDegreeCard.Show 1, Me
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByRef cbsMain As CommandBars, _
    ByVal int���� As Integer, Optional ByVal blnChildToolBar As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
    '���:int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    '       CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
    '       blnChildToolBar = True��ʾ������������Լ��Ĵ����ڲ�
    '����:
    '����:
    '˵��:
    '   �����Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)�����bln�ڲ�������Ϊ�٣������������ϴ������������˵���ȻҪ��������
    '   ����Ҫ���Լ��Ľ����ϴ�������������˶����Լ��������Ѿ����ڹ������ĳ���Ӧ����ؼ����ظ���
    'ע��:
    '         ��ӹ�����ʱע��������ܰ�ť��������Ҫ�ظ�
    '         ����������ģ���޲˵���conMenu_ManagePopup������ӳ����ڴ���ʱ��Ҫ��飬�޴˶���ʱ��ӵ����ѵĲ˵���
    '         ���������ڲ�����������ɾ����������������
    '         δʹ�ù�������ģ����Ҫ��ӳ�ʼ������������
    '         �������Ĺ���״̬�ı仯��ͨ�����������zlUpdateCommandBars��ͳһ����
    '����:���˺�
    '����:2010-10-29 15:14:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Set mfrmParent = frmParent
    Set mcbsMain = cbsMain
    mint���� = int����
        
    Err = 0: On Error GoTo ErrHand:
        
    '�ļ��˵�
    Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_FilePopup, True, False)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(xtpControlButton, conMenu_File_Excel, True, False) '�����Excel֮��
        If mint���� = 1 Then
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintBedCard, "��ӡ��ͷ��(&K)��", objControl.Index + 1) '��ӡ��ͷ��
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintPageSet, "��ӡ��ҳ����(&A)��", objControl.Index + 1)
    End With
    
    '�б༭�˵�ʱ�����ڱ༭�˵���(���ò�ѯģ��)��������ڹ���˵�(���������û��)���ļ��˵�����
    Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_EditPopup, True, False)
    If objMenu Is Nothing Then
        Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ManagePopup, True, False)
        If objMenu Is Nothing Then
            Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_FilePopup, True, False)
        End If
        ''0-���ò�ѯ��1-��ʿվ����
        '���:C;E:63630
        Set objMenu = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, IIf(mint���� = 1, "����(&C)", "�༭(&E)"), objMenu.Index + 1, False)
        objMenu.ID = conMenu_EditPopup
    End If
    With objMenu.CommandBar.Controls
        If mint���� = 1 Then
            '����:40900
            Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalanceAll, "Ԥ�����в���(&I)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "Ԥ�ᵱǰ����(&W)")
            objControl.BeginGroup = True
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing, "����(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing_Mulit, "��������(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBilling, "����(&D)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingApply, "��������(&L)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingAudit, "�������(&U)", objControl.Index + 1)
        If mint���� = 1 Then
            Set objControl = .Add(xtpControlButton, conMenu_View_Billing, "�鿴���ʵ�(&D)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_CardBackMoney, "����˿�(&F)"): objControl.BeginGroup = True
        End If
        '54274
       Set objControl = .Add(xtpControlButton, conMenu_Edit_Copy, "���Ƽ��˵�(&F)"): objControl.BeginGroup = True
       If mint���� = 0 Then .Add(xtpControlButton, conMenu_Edit_PatiMemo, "���˱�ע��Ϣ(&M)").BeginGroup = True
    End With
             
    
    '�鿴�˵�
    '-----------------------------------------------------
    Set objMenu = mcbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ViewPopup, True, False)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(xtpControlButton, conMenu_View_StatusBar, True, False) '״̬�����
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_DetailType, "�嵥����(&M)", objControl.Index + 1): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 0, "�����嵥(&0)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 1, "�ֿ�����ϸ(&1)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 2, "����Ŀ��ϸ(&2)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 3, "�������ϸ(&3)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 4, "���������ϸ(&4)", -1, False
                        
            .Add(xtpControlButton, conMenu_View_DetailType * 10 + 5, "����Ŀ����(&5)", -1, False).BeginGroup = True
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 6, "��������(&6)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 7, "���·������(&7)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 8, "���յ��ݻ���(&8)", -1, False
            .Add xtpControlButton, conMenu_View_DetailType * 10 + 9, "���շ�Ŀ����(&9)", -1, False
        End With
                
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_DateType, "��ѯʱ��(&E)", objPopup.Index + 1)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_DateType * 10 + 1, "����ʱ��(&H)", -1, False
            .Add xtpControlButton, conMenu_View_DateType * 10 + 2, "�Ǽ�ʱ��(&A)", -1, False
        End With
        
        If mint���� = 1 Then
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ViewPopup, True, False)
            With objMenu.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "�鿴������ϸ��Ϣ(&K)"): objControl.BeginGroup = True
            End With
         End If
         
        Set objControl = .Add(xtpControlButton, conMenu_View_ReBalance, "��ʾ��������(&Q)", objPopup.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_ZeroFee, "��ʾ�����(&Z)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_CheckFee, "��ʾ������(&C)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_TurnToWardFeeQuery, "ת�������ñ䶯��ѯ(&T)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_ContainOutFee, "�����������(&B)", objControl.Index + 1)
    End With
    
    
    '����˵�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ReportPopup, True, False)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ViewPopup, True, False)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_ReportPopup, True, False)
    With objMenu.CommandBar.Controls
        If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "��ӡһ���嵥(&D)��", 1)
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "��ӡ�߿(&C)��", objControl.Index + 1)
    End With
    
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "���ʲ���ѡ��(&O)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_File_Parameter
    End With
    
    
    If mint���� <> 1 Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_EditPopup, True, False)
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "�鿴������ϸ��Ϣ(&K)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Billing, "�鿴���ʵ�(&D)", objControl.Index + 1)
        End With
    End If
        
    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    If blnChildToolBar Then
        Set objBar = CreateChildTools
    Else
        cbsTools.DeleteAll
        Set objBar = mcbsMain(2)
    End If
    
    If blnChildToolBar = False Then
        For Each objControl In objBar.Controls '�����ǰ������һ��Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
            End If
        Next
    End If
    Dim intIndex As Integer
    With objBar.Controls
        Set objControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
        If objControl Is Nothing Then
            intIndex = 0
        Else
            intIndex = objControl.Index
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "�߿�", intIndex + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "һ��", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "Ԥ��", objControl.Index + 1)
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing, "����", objControl.Index + 1)
        objControl.BeginGroup = True
        intIndex = objControl.Index
        For Each objControl In objBar.Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
        
        Set objPopup = .Add(xtpControlPopup, conMenu_Edit_ReBillingButton, "����", intIndex + 1)
        objPopup.ID = conMenu_Edit_ReBillingButton
        objPopup.IconId = conMenu_Edit_ReBillingButton
        objPopup.Style = xtpButtonIconAndCaption
                
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add 0, VK_F4, conMenu_Edit_Billing
        .Add 0, VK_F6, conMenu_Edit_ReBilling
        .Add 0, vbKeyF11, conMenu_Tool_Option '����ѡ��
    End With

    '���ò���������
    '-----------------------------------------------------
    With mcbsMain.Options   '��������ˣ��ؼ��ڲ˵���һ����ʾʱû�е���update�¼�
'        .AddHiddenCommand conMenu_View_ReBalance
'        .AddHiddenCommand conMenu_View_ZeroFee
'        .AddHiddenCommand conMenu_View_CheckFee
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function CreateChildTools() As CommandBar
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ӹ�����
    '����:���˺�
    '����:2010-10-29 15:59:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsTools.DeleteAll
    Set CreateChildTools = cbsTools.Add("���ò���", xtpBarTop)
    CreateChildTools.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    CreateChildTools.ModifyStyle XTP_CBRS_GRIPPER, 0
    
End Function


Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_ReBillingButton '��������
        With CommandBar.Controls
            .DeleteAll
            .Add(xtpControlButton, conMenu_Edit_ReBillingApply, "��������(&L)").BeginGroup = True
            .Add xtpControlButton, conMenu_Edit_ReBillingAudit, "�������(&U)"
        End With
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnSelect As Boolean, lngColTmp As Long, blnEnabled As Boolean
    
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    blnSelect = mlng����ID <> 0
    Select Case Control.ID
        '�ļ�
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel
            Control.Enabled = vsfFee.Rows > vsfFee.FixedRows
        Case conMenu_File_PrintSingleBill
            Control.Enabled = blnSelect
            
        '�༭
        Case conMenu_Edit_PreBalance
            Control.Enabled = blnSelect
            If blnSelect Then
                Dim blnMCPatient As Boolean
                If tabClass.SelectedItem.Index = conTabδ�� Then
                    blnMCPatient = mintInsure <> 0
                End If
                Control.Enabled = blnMCPatient
            End If
            mblnPreBalance = Control.Enabled
        Case conMenu_Edit_PatiMemo   '�޸ı�ע��Ϣ
           ' Control.Visible = InStr(1, mstrPrivs, ";���˱�ע�༭;")
            Control.Enabled = mlng����ID > 0 And Control.Visible
        Case conMenu_Edit_Billing
            Control.Enabled = blnSelect
        Case conMenu_Edit_Billing_Mulit '��������
        
        Case conMenu_Edit_Copy
            '54274
            Control.Enabled = blnSelect
            With vsfFee
                If Control.Enabled Then
                    If mbytList = ListType.C0�����嵥 Or mbytList = ListType.C1�ֿ�����ϸ Or mbytList = ListType.C2����Ŀ��ϸ Or mbytList = ListType.C3�������ϸ Or mbytList = ListType.C4���������ϸ Then
                        '.row>=1:61895
                        Control.Enabled = .ColIndex("���ݺ�") >= 0 And .ColIndex("��¼״̬") >= 0 And .ColIndex("��¼����") >= 0 And .Row >= 1
                        If Control.Enabled Then
                            If Trim(.TextMatrix(.Row, .ColIndex("���ݺ�"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 2 Or Val(.TextMatrix(.Row, .ColIndex("��¼����"))) = 3 Then
                                Control.Enabled = False
                            End If
                        End If
                     Else
                        Control.Enabled = False
                    End If
                End If
            End With
        Case conMenu_Edit_ReBilling '����
            Control.Enabled = blnSelect
            With vsfFee
                If Control.Enabled Then
                    If mbytList = ListType.C0�����嵥 Or mbytList = ListType.C1�ֿ�����ϸ Or mbytList = ListType.C2����Ŀ��ϸ Or mbytList = ListType.C3�������ϸ Or mbytList = ListType.C4���������ϸ Then
                        lngColTmp = VsfGetColNum(vsfFee, "��¼״̬")
                        If lngColTmp = -1 Or .Row < 1 Then
                            Control.Enabled = False
                        Else
                            lngColTmp = Val(.TextMatrix(.Row, lngColTmp))
                            Control.Enabled = (lngColTmp = 1 Or lngColTmp = 3)
                        End If
                    Else
                        Control.Enabled = False
                    End If
                End If
            End With
            mblnUnBilling = Control.Enabled
            
       '�鿴
        Case conMenu_View_PatInfor
            Control.Enabled = blnSelect
        Case conMenu_View_Billing
            Control.Enabled = vsfFee.Rows > vsfFee.FixedRows
            
        Case conMenu_View_DateType * 10 + 1, conMenu_View_DateType * 10 + 2
            Control.Checked = (Control.ID - conMenu_View_DateType * 10) = mbytDateType
            Control.Enabled = blnSelect
        Case conMenu_View_DetailType * 10 To conMenu_View_DetailType * 10 + 9
            Control.Checked = (Control.ID - conMenu_View_DetailType * 10) = mbytList
            Control.Enabled = blnSelect
            
        Case conMenu_View_ZeroFee
            Control.Enabled = tabClass.SelectedItem.Index = conTabδ��
            Control.Checked = mvs.ZeroFee
        Case conMenu_View_CheckFee
            Control.Checked = mvs.CheckFee
        Case conMenu_View_ReBalance
            Control.Checked = mvs.ReBalance
        Case conMenu_View_ContainOutFee
            Control.Checked = mblnContainOutFee
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then Control.Enabled = blnSelect  '������ҳ
            End If
        
    End Select
End Sub

Private Sub cboNum_Click()
    If mblnNotClick Then Exit Sub
    Call LoadPages(cboNum.ListIndex + 1)
End Sub

Private Sub cbo����_Click()
    If mblnNotClick Then Exit Sub
    Call SetDateVisible
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbsTools_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
        Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsTools_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
   If CommandBar.Parent Is Nothing Then Exit Sub
    Call zlPopupCommandBars(CommandBar)
End Sub

Private Sub cbsTools_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
       Call Form_Resize
End Sub

Private Sub cbsTools_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
       Call zlUpdateCommandBars(Control)
End Sub

Private Sub chkAdivce_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub


Private Sub cboBaby_Click()
    If mblnNotClick Then Exit Sub
    If Visible Then Call LoadCardData(False, False, True)
End Sub

Private Sub cboFeeType_Click()
    If mblnNotClick Then Exit Sub
    Call FilterDetail
End Sub

Private Sub cboDept_Click()
    '����64817:������,2013-10-28,�л��������б�ʱ��ѯ�Ļ������ݲ�û�а��տ����������й���
    If cboDept.Tag = "��ˢ��" Or mblnNotClick Then Exit Sub
    Call FilterDetail
End Sub

Private Sub FilterDetail()
'����:����ѡ��ķ�Ŀ����ҹ�����ϸ����
'����:
    Dim arrTotal(2) As Currency
    Dim strFilter As String
        
    Select Case mbytList
        Case ListType.C0�����嵥, ListType.C1�ֿ�����ϸ, ListType.C2����Ŀ��ϸ, ListType.C3�������ϸ, ListType.C4���������ϸ
            If mrsList Is Nothing Then Exit Sub
            If mrsList.State = adStateClosed Then Exit Sub
            
            If cboFeeType.ListIndex > 0 Then strFilter = "��Ŀ='" & cboFeeType.Text & "'"
            If cboDept.ListIndex > 0 Then strFilter = IIf(strFilter = "", "", strFilter & " And") & " ��������='" & cboDept.Text & "'"
            
            mrsList.Filter = strFilter
            Set vsfFee.DataSource = mrsList
            
            Call SetVsffeeFormat
            vsfFee.AutoSize 0, vsfFee.Cols - 1
            
            '������Ի�����
            zl_vsGrid_Para_Restore mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False
            
        Case ListType.C5����Ŀ����, ListType.C6��������, ListType.C7���·������, ListType.C8���յ��ݻ���, ListType.C9���շ�Ŀ����
            Call LoadCardData(False, False, True, False)
    End Select
End Sub

Private Function LoadPatiClass() As Boolean
'���ܣ����ò��˵ķ���ѡ�
    Dim strSQL As String, i As Long, intPage As Integer, intCount As Integer
    Dim rsTmp As ADODB.Recordset
    Dim cllPage As Collection
    Dim str(3) As String
    
    mintPreCard = 0
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
        
    '�����ǰ���˵���Ժʱ����ת��ʱ��֮ǰ,����Ҫ��������ݱ��ѯ
    If mblnDateMoved Then
        strSQL = zlGetFullFieldsTable("���˽��ʼ�¼")
    Else
        strSQL = "���˽��ʼ�¼ A"
    End If
    
    strSQL = "Select A.ID,A.NO,A.�շ�ʱ�� as ����,A.��¼״̬" & _
        " From " & strSQL & " " & _
        " Where A.����ID = [1]" & _
        " And A.��¼״̬ IN (1" & IIf(mvs.ReBalance, ",3", "") & ")" & _
        " Order by A.ID Desc"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID)
        
    cboNum.Visible = False
    picNum.Visible = False
    Set mcllBalaceNums = New Collection
    If rsTmp.RecordCount >= 20 Then
        cboNum.Visible = True
        picNum.Visible = True
    End If
    
    Set cllPage = New Collection
    Do While Not rsTmp.EOF
        If rsTmp.RecordCount < 20 Then
            tabClass.Tabs.Add , "_" & rsTmp!NO, Format(rsTmp!����, "yyyy-MM-dd") & IIf(rsTmp!��¼״̬ = 1, " ����", " �˷�")
            tabClass.Tabs(tabClass.Tabs.Count).Tag = rsTmp!ID '��¼����ID,�ӿ��ٶ�
            tabClass.Tabs(tabClass.Tabs.Count).ToolTipText = "����ʱ��:" & Format(rsTmp!����, "yyyy-MM-dd hh:mm:ss")
        Else
            str(0) = Val(rsTmp!ID)
            str(1) = Nvl(rsTmp!NO)
            str(2) = Format(rsTmp!����, "yyyy-MM-dd") & IIf(rsTmp!��¼״̬ = 1, " ����", " �˷�")
            str(3) = "����ʱ��:" & Format(rsTmp!����, "yyyy-MM-dd hh:mm:ss")
            cllPage.Add str
            'cllPage.Add Array(rsTmp!ID, rsTmp!NO, Format(rsTmp!����, "yyyy-MM-dd") & IIf(rsTmp!��¼״̬ = 1, " ����", " �˷�"), "����ʱ��:" & Format(rsTmp!����, "yyyy-MM-dd hh:mm:ss"))
            intPage = intPage + 1
            intCount = intCount + 1
            If intPage >= 5 Or rsTmp.RecordCount = intCount Then
                mcllBalaceNums.Add cllPage
                Set cllPage = New Collection
                intPage = 0
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    '���ط�ҳ����
    If cboNum.Enabled And cboNum.Visible Then
        cboNum.Clear
        For i = 1 To mcllBalaceNums.Count
            cboNum.AddItem "��" & i & "ҳ"
            cboNum.ItemData(cboNum.NewIndex) = i
            If i = 1 Then cboNum.ListIndex = cboNum.NewIndex
        Next
        If mstrRestoreFeeCons <> "" Then
            For i = 0 To cboNum.ListCount - 1
                If Nvl(Split(mstrRestoreFeeCons, "|")(4)) = cboNum.List(i) Then cboNum.ListIndex = i: Exit For
            Next
        End If
    End If
    If picNum.Enabled And picNum.Visible Then
        picNum.Left = IIf(995 + 1680 * (tabClass.Tabs.Count - 1) + 120 < Me.ScaleWidth - picNum.Width, 995 + 1680 * (tabClass.Tabs.Count - 1) + 120, Me.ScaleWidth - picNum.Width - 30)
        tabClass.Width = Me.ScaleWidth - picNum.Width - 30
    Else
        tabClass.Width = Me.ScaleWidth
    End If
    
    LoadPatiClass = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPatiTime() As Boolean
'���ܣ����ò���סԺ����ѡ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
      
    With tabTime
        mintPreTime = 0
        For i = .Tabs.Count To 2 Step -1    '������һ��
            .Tabs.Remove i
        Next
        .Visible = tabClass.SelectedItem.Index = conTabδ��
        If Not (tabClass.SelectedItem.Index = conTabδ��) Then LoadPatiTime = True: Exit Function
        
        On Error GoTo errH
        strSQL = "Select ��ҳID,��Ժ����,��Ժ���� From ������ҳ Where Nvl(��ҳID,0)<>0 And ����ID=[1] Order by ��ҳID Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID)
      
        Do While Not rsTmp.EOF
            .Tabs.Add , "_" & rsTmp!��ҳID, "��" & rsTmp!��ҳID & "��"
            .Tabs((.Tabs.Count)).Tag = rsTmp!��ҳID
            .Tabs((.Tabs.Count)).ToolTipText = "��Ժ:" & Format(rsTmp!��Ժ����, "yyyy-MM-dd") & _
                                                IIf(Not IsNull(rsTmp!��Ժ����), ",��Ժ:" & Format(rsTmp!��Ժ����, "yyyy-MM-dd"), "")
            '�����:53136 �޸���:���˺�,�޸�ʱ��:2012-12-10 13:26:07
            If Val(Nvl(rsTmp!��ҳID)) = mlng��ҳID Then
                .Tag = "1"
                .Tabs((.Tabs.Count)).Selected = True
                .Tag = ""
            End If
            rsTmp.MoveNext
        Loop
    End With
    LoadPatiTime = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintList(bytStyle As Byte)
    Dim objOut As zlPrint1Grd
    Dim objRow As zlTabAppRow, strTmp As String, bytR As Byte, lngTmp As Long
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng��ҳID As Long
            
    On Error GoTo errH
    Set objOut = New zlPrint1Grd
    '����
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objOut.UnderAppRows.Add objRow
    
    If tabClass.SelectedItem.Index = conTabδ�� Then
        If tabTime.SelectedItem.Index = 1 Then
            '��ǰ�����嵥�е���Ϣ
            lng��ҳID = mlng��ҳID
        Else
            'ָ��סԺ��������Ϣ
            lng��ҳID = Val(tabTime.SelectedItem.Tag)
        End If
        strSQL = "" & _
        "   Select Nvl(b.����, a.����) As ����,A.סԺ��,B.��Ժ���� as ����," & _
        "           Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(b.����,a.����) As ����,B.��Ժ����,B.��Ժ����,C.���� as ����" & _
        "   From ������Ϣ A,������ҳ B,���ű� C" & _
        "   Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
        "           And A.����ID=[1] And B.��ҳID=[2] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID, lng��ҳID)
        If rsTmp.EOF Then Exit Sub
        
        Set objRow = New zlTabAppRow
        objRow.Add "����:" & rsTmp!���� & "    סԺ��:" & rsTmp!סԺ�� & "    ����:" & rsTmp!���� & "    �Ա�:" & rsTmp!�Ա� & "    ����:" & rsTmp!����
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "�� " & lng��ҳID & " ��סԺ    ����:" & rsTmp!���� & _
            "    ��Ժ����:" & Format(rsTmp!��Ժ����, "yyyy-MM-dd") & "    ��Ժ����:" & Format(Nvl(rsTmp!��Ժ����), "yyyy-MM-dd")
        objOut.UnderAppRows.Add objRow
    Else
        strSQL = "Select Max(��ҳID) as ��ҳID From סԺ���ü�¼ Where ����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, Val(tabClass.SelectedItem.Tag))
        If Not rsTmp.EOF Then lng��ҳID = Nvl(rsTmp!��ҳID, 0)
        If lng��ҳID <> 0 Then
            '�Խ������סԺ����Ϊ׼
            strSQL = "" & _
            "   Select Nvl(b.����, a.����) As ����,A.סԺ��,B.��Ժ���� as ����," & _
            "           Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(b.����, a.����) as ����,B.��Ժ����,B.��Ժ����,C.���� as ����" & _
            "   From ������Ϣ A,������ҳ B,���ű� C" & _
            "   Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
            "           And A.����ID=[1] And B.��ҳID=[2] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID, lng��ҳID)
            If rsTmp.EOF Then Exit Sub
        
            Set objRow = New zlTabAppRow
            objRow.Add "����:" & rsTmp!���� & "    סԺ��:" & rsTmp!סԺ�� & "    ����:" & rsTmp!���� & "    �Ա�:" & rsTmp!�Ա� & "    ����:" & rsTmp!����
            objOut.UnderAppRows.Add objRow
        
            Set objRow = New zlTabAppRow
            objRow.Add "�� " & lng��ҳID & " ��סԺ    ����:" & rsTmp!���� & _
                "    ��Ժ����:" & Format(rsTmp!��Ժ����, "yyyy-MM-dd") & "    ��Ժ����:" & Format(Nvl(rsTmp!��Ժ����), "yyyy-MM-dd")
            objOut.UnderAppRows.Add objRow
        Else
            '��Ľ����������
            strSQL = "Select A.����,A.סԺ��,A.�Ա�,A.���� From ������Ϣ A Where A.����ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID)
            If rsTmp.EOF Then Exit Sub
        
            Set objRow = New zlTabAppRow
            objRow.Add "����:" & rsTmp!���� & "    סԺ��:" & rsTmp!סԺ�� & "    �Ա�:" & rsTmp!�Ա� & "    ����:" & rsTmp!����
            objOut.UnderAppRows.Add objRow
        End If
    End If
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objOut.UnderAppRows.Add objRow
    
    '���øſ����
    Set objRow = New zlTabAppRow
    objRow.Add lbl������Ϣ.Caption
    objOut.UnderAppRows.Add objRow
    
    objOut.Title.Font.Size = 16
    If tabClass.SelectedItem.Index = conTabδ�� Then
        '�����嵥����
        Dim objControl As CommandBarControl
        Set objControl = mcbsMain.ActiveMenuBar.FindControl(xtpControlButton, conMenu_View_DetailType * 10 + mbytList, True, True)
        If objControl Is Nothing Then
            strTmp = ""
        Else
            strTmp = objControl.Caption
        End If
        objOut.Title.Text = GetUnitName & "����δ�����" & Left(strTmp, Len(strTmp) - 4)
    Else
        objOut.Title.Text = GetUnitName & "���˽�����ϸ�嵥"
        '����
        Set objRow = New zlTabAppRow
        objRow.Add ""
        objOut.UnderAppRows.Add objRow
        '�������
        Set objRow = New zlTabAppRow
        objRow.Add "��ʽ:" & Right(tabClass.SelectedItem.Caption, 3)
        objRow.Add "��������:" & Left(tabClass.SelectedItem.Caption, Len(tabClass.SelectedItem.Caption) - 3)
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ע:"
    objOut.BelowAppRows.Add objRow
    
    If vsfFee.FixedCols = 1 Then
        vsfFee.Redraw = flexRDNone
        vsfFee.OutlineBar = flexOutlineBarNone
        lngTmp = vsfFee.ColWidth(0)
        vsfFee.ColWidth(0) = 0
    End If
    Set objOut.Body = vsfFee
    
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    If vsfFee.FixedCols = 1 Then
        vsfFee.OutlineBar = flexOutlineBarComplete
        vsfFee.ColWidth(0) = lngTmp
        vsfFee.Redraw = flexRDDirect
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadFeeOutline(ByVal lng����ID As Long, ByVal blnDateMoved As Boolean, ByVal lng����ID As Long)
    '����:��ʾ���˷��øſ�:Ԥ�����,�������,Ԥ�����,��ĳ�ν��ʷ��øſ�
    '����:lng����ID:0-��ʾ����ʾ��ʼ��Ϣ.
    Dim rsTmp As ADODB.Recordset, strWhere As String
    Dim strSQL As String, strInfo As String, strTmp As String, i As Long
    Dim lngColor As Long
    Dim dblYbMoney As Double, dblTotal As Double
     
    lngColor = ForeColor

    On Error GoTo errH
    'a.δ����øſ�
    If lng����ID = 0 Then
        If lng����ID > 0 Then
            If mblnContainOutFee = False Then strWhere = " And ����=2"
            strSQL = _
                " Select Nvl(Ԥ�����,0) As Ԥ�����,Nvl(�������,0) As �������,0 as Ԥ�����" & _
                " From �������" & _
                " Where ����=1 And ����ID=[1]" & strWhere
            
            If mblnPreBalance Then
                strSQL = strSQL & " Union ALL " & _
                    " Select 0 as Ԥ�����,0 as �������,Sum(B.���) as Ԥ�����" & _
                    " From ������Ϣ A,����ģ����� B" & _
                    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
            End If
            
            strSQL = _
                " Select Nvl(Sum(Ԥ�����),0) as Ԥ�����,Nvl(Sum(�������),0) as �������,Nvl(Sum(Ԥ�����),0) as Ԥ�����" & _
                " From (" & strSQL & ")"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID)
            If rsTmp.RecordCount > 0 Then
                '�Ը�����:43575
                With rsTmp
                    strInfo = "������Ϣ��Ԥ����:" & Format(!Ԥ�����, "0.00") & Space(4) & "δ�����:" & Format(!�������, gstrDec) & _
                            IIf(mblnPreBalance, Space(4) & "Ԥ�����:" & Format(!Ԥ�����, gstrDec) & Space(4) & "�Ը�����:" & Format(Val(Nvl(!�������)) - Val(Nvl(!Ԥ�����)), gstrDec), "") & _
                            Space(4) & "ʣ���:" & Format((!Ԥ����� - !������� + !Ԥ�����), "0.00")
                            If (Val(Nvl(!Ԥ�����)) - Val(Nvl(!�������)) + Val(Nvl(!Ԥ�����))) < 0 Then
                                    lngColor = vbRed
                            End If
                End With
            Else
                strInfo = "������Ϣ��Ԥ����:0.00" & Space(4) & "δ�����:" & gstrDec & Space(4) & "ʣ���:0.00"
            End If
            
            strTmp = GetPatientDue(lng����ID)
            If Val(strTmp) <> 0 Then strInfo = strInfo & Space(4) & "Ӧ�տ�:" & Format(strTmp, "0.00")
            
            strSQL = "Select ������,������ From ������Ϣ Where ����ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID)
            If rsTmp.RecordCount > 0 Then
                If Not IsNull(rsTmp!������) Or Not IsNull(rsTmp!������) Then
                    strInfo = strInfo & "������:" & rsTmp!������
                    strInfo = strInfo & "������:" & Format(Nvl(rsTmp!������, 0), "0.00")
                End If
            End If
        Else
            strInfo = "������Ϣ��Ԥ����:0.00" & Space(4) & "δ�����:" & gstrDec & Space(4) & "ʣ���:0.00"
        End If
        
    'b.���ʸſ�
    Else
        
        strInfo = "���ݺ�:" & Mid(tabClass.SelectedItem.Key, 2)
        
        strSQL = _
            " Select nvl(���ʽ��,0) as ���ʽ�� From ������ü�¼ where ����ID=[1]" & _
            " Union ALL " & _
            " Select nvl(���ʽ��,0) as ���ʽ�� From סԺ���ü�¼ where ����ID=[1]  "
        If blnDateMoved Then
            strSQL = strSQL & " UNION ALL " & Replace(Replace(strSQL, "������ü�¼", "H������ü�¼"), "סԺ���ü�¼", "HסԺ���ü�¼")
        End If
        
        strSQL = "Select Nvl(Sum(���ʽ��),0) as ��� From  (" & strSQL & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        
        strInfo = strInfo & Space(4) & "���ʽ��:" & Format(rsTmp!���, gstrDec)
        dblTotal = Val(Nvl(rsTmp!���))
        strSQL = _
            " Select Decode(Substr(A.��¼����,Length(A.��¼����),1),1,'��Ԥ��',A.���㷽ʽ) as ���㷽ʽ,Sum(Nvl(��Ԥ��,0)) as ���," & _
            "               Decode(Substr(A.��¼����,Length(A.��¼����),1),1,0,B.����,3,1,b.����,4,1,0) as ҽ�� " & _
            " From " & IIf(blnDateMoved, zlGetFullFieldsTable("����Ԥ����¼"), "����Ԥ����¼ A") & " ,���㷽ʽ B " & _
            " Where A.���㷽ʽ=B.����(+) And A.����ID=[1]" & _
            " Group by Decode(Substr(A.��¼����,Length(A.��¼����),1),1,'��Ԥ��',A.���㷽ʽ)," & _
            "       Decode(Substr(a.��¼����, Length(a.��¼����), 1), 1, 0, b.����, 3, 1, b.����, 4, 1, 0)" & _
            " Order by ���㷽ʽ"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
       
        dblYbMoney = 0
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���㷽ʽ = "��Ԥ��" Then
                strInfo = strInfo & Space(4) & "��Ԥ��:" & Format(rsTmp!���, "0.00")
            Else
                strInfo = strInfo & Space(4) & IIf(rsTmp!��� < 0, "��", "��") & rsTmp!���㷽ʽ & ":" & Format(Abs(rsTmp!���), "0.00")
            End If
            If Val(Nvl(rsTmp!ҽ��)) = 1 Then dblYbMoney = dblYbMoney + Val(Nvl(rsTmp!���))
            rsTmp.MoveNext
        Next
        If dblYbMoney <> 0 Then '43575
            strInfo = strInfo & Space(4) & "�Ը����:" & Format(dblTotal - dblYbMoney, "0.00")
        End If
        strInfo = "������Ϣ��" & strInfo
    End If
    
    lbl������Ϣ.Caption = strInfo
    lbl������Ϣ.ForeColor = lngColor
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDeposit(ByVal lng����ID As Long, ByVal blnDateMoved As Boolean, ByVal lng����ID As Long)
'����:��ʾδ�����Ԥ�����嵥����ʳ�Ԥ����ϸ
'����:lng����ID:0-��ʾ����ʾ��ʼ��Ϣ.
    Dim rsTmp As ADODB.Recordset, strWhere As String
    Dim strSQL As String, strDepost As String, i As Long
    
    On Error GoTo errH
    
    'a.δ�����Ԥ�����嵥
    If lng����ID = 0 Then
        lblDepost.Caption = " Ԥ�����嵥"
        If mblnContainOutFee = False Then
            strWhere = " And Nvl(A.Ԥ�����, 2) = 2"
        End If
        
        '����,���ݺ�,����,���㷽ʽ,�������,Ԥ�����,���ʽ��,ժҪ
        strSQL = "Select To_Char(Max(Decode(a.��¼����, 1, a.�տ�ʱ��, Null)), 'YYYY-MM-DD') As ����, " & vbNewLine & _
                "       a.NO As ���ݺ�, " & vbNewLine & _
                "       Max(Decode(a.��¼����, 1, b.����, Null)) As ����, " & vbNewLine & _
                "       Max(Decode(a.��¼����, 1, a.���㷽ʽ, Null)) As ���㷽ʽ, " & vbNewLine & _
                "       Max(Decode(a.��¼����, 1, a.�������, Null)) as �������, " & vbNewLine & _
                "       To_Char(Sum(Nvl(a.���, 0)), 'FM9999999990.00') As Ԥ�����," & vbNewLine & _
                "       To_Char(Sum(Nvl(a.��Ԥ��, 0)), 'FM9999999990.00') As ���ʽ��," & vbNewLine & _
                "       To_Char(Sum(Nvl(a.���, 0)) - Sum(Nvl(a.��Ԥ��, 0)), 'FM9999999990.00') As ʣ����, " & vbNewLine & _
                "       Max(Decode(a.��¼����, 1, a.ժҪ, Null)) as ժҪ, " & vbNewLine & _
                "       Max(Decode(a.��¼����, 1, a.ʵ��Ʊ��, Null)) as ʵ��Ʊ��, " & vbNewLine & _
                "       Max(Decode(a.��¼����, 1, a.����Ա����, Null)) as ����Ա����," & vbNewLine & _
                "       Decode(a.Ԥ�����,1,'����Ԥ��','סԺԤ��') As Ԥ�����" & vbNewLine & _
                " From ����Ԥ����¼ A, ���ű� B" & vbNewLine & _
                " Where A.����id = B.ID(+) And A.��¼���� In (1, 11) And A.����id = [1]" & strWhere & vbNewLine & _
                " Group By A.NO,a.Ԥ�����" & vbNewLine & _
                " Having Sum(Nvl(A.���, 0)) - Sum(Nvl(A.��Ԥ��, 0)) <> 0" & vbNewLine & _
                " Order By Ԥ����� Desc,����, ���ݺ�"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID)
        strDepost = "4,4,1,4,1,7,7,7,1,1,1"
        
     'b.���ʳ�Ԥ����ϸ
     Else
        lblDepost.Caption = " Ԥ��ʹ���嵥"
        '���ڡ����ݺš����ҡ����㷽ʽ��������롢���ʽ�ժҪ
        strSQL = "Select To_Char(c.�տ�ʱ��, 'YYYY-MM-DD') As ����, c.No As ���ݺ�, b.���� As ����, c.���㷽ʽ, " & vbNewLine & _
                "       c.�������, LTrim(To_Char(Nvl(a.��Ԥ��, 0), '9999999990.00')) As ��Ԥ�����, " & vbNewLine & _
                "       c.ժҪ, c.ʵ��Ʊ��, a.����Ա���� As ���ʲ���Ա, c.����Ա���� As Ԥ���տ����Ա, " & vbNewLine & _
                "       Decode(c.Ԥ�����,1,'����Ԥ��','סԺԤ��') As Ԥ�����" & vbNewLine & _
                " From " & IIf(blnDateMoved, zlGetFullFieldsTable("����Ԥ����¼"), "����Ԥ����¼ A") & ", ���ű� B, " & vbNewLine & _
                        IIf(blnDateMoved, zlGetFullFieldsTable("����Ԥ����¼", , , , "C"), "����Ԥ����¼ C") & vbNewLine & _
                " Where a.No = c.No And c.��¼���� = 1 And c.��¼״̬ In (1, 3) And c.����id = b.Id(+)" & vbNewLine & _
                "       And a.��¼���� In (1, 11) And Nvl(a.��Ԥ��, 0) <> 0 And a.����id = [1] " & vbNewLine & _
                " Order By Ԥ����� Desc,����, ���ݺ�"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID)
        strDepost = "4,4,4,1,1,7,1,1,1,1"
    End If
    
    With mshDepost
        .Redraw = flexRDNone
        
        Set .DataSource = rsTmp       'ʹ�ô˷�ʽ,�´�������ʱ������ж�λ��λ,������ʾ�ϼ���
        If rsTmp.RecordCount = 0 Then .Rows = 2
        '��ʽ����
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4
            .ColKey(i) = Trim(.TextMatrix(0, i))
        Next
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        For i = 0 To UBound(Split(strDepost, ","))
            .ColAlignment(i) = Split(strDepost, ",")(i)
        Next
        zl_vsGrid_Para_Restore mlngModul, mshDepost, Me.Name, "mshDepost", False
        .Redraw = flexRDBuffered
        If rsTmp.RecordCount > 0 Then .Row = 1: .Col = 0
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function LoadCardData(Optional blnHead As Boolean = True, _
    Optional blnDeposit As Boolean = True, Optional blnMoney As Boolean = True, Optional blnLoadDept As Boolean = True) As Boolean
    '���ܣ����ݵ�ǰѡ��Ĳ��˷�����Ŀ��Ƭ����ȡ�����÷����嵥
    '������blnHead=ֻ����ſ�����
    '      blnDeposit=ֻ����Ԥ�����
    '      blnMoney=ֻ������ò���
    '      blnLoadDept=���ܲ�ѯʱ���Ƿ����¶�ȡ�����б�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str���˷��� As String, strIF As String, strWhere��ҳ As String
    Dim lng��ҳID As Long, lng����ID As Long
    Dim strStartDate As String, strEndDate As String, strWhere As String, strDept As String
    Dim strBaby As String, strDateMode As String
    Dim blnDateMoved As Boolean '��¼��ǰѡ��������Ƿ����ں����ݱ���
    Dim str�������� As String, str���෽ʽ As String, str��� As String, str���ϼ� As String, str�������� As String, str�����ջ� As String
    Dim lng��������ID As Long
    Dim lngPre����ID As Long    '�ϴο���ID
    Dim strPre����Text As String, blnNotCheckFee As Boolean
    Dim strWhereCheckFee As String  '�����ù�������
    
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڼ��ط�����Ϣ,���Ժ� ...", Me)
    If mlng����ID <> 0 Then lng��ҳID = Val(tabTime.SelectedItem.Tag)
    
    '����������:46646
    strWhereCheckFee = IIf(chkNotCheckFee.Value = 1, " And nvl(A.�����־,0)<>4 ", "")
    
    strDateMode = IIf(mbytDateType = 1, "����", "�Ǽ�")
    '���һ����:���˺�Ӥ��
    If cboBaby.Visible And cboBaby.ListIndex < cboBaby.ListCount - 1 Then strBaby = " And Nvl(A.Ӥ����,0)=" & cboBaby.ItemData(cboBaby.ListIndex)
        
    If blnMoney Then
        If mbytList = ListType.C1�ֿ�����ϸ Then
            If optDeptMode(0).Value = True Then
                str�������� = optDeptMode(0).Caption
            Else
                str�������� = optDeptMode(1).Caption
            End If
        ElseIf mbytList = ListType.C4���������ϸ Then
            If optTypeMode(0).Value = True Then
                str���෽ʽ = optTypeMode(0).Caption
            Else
                str���෽ʽ = optTypeMode(1).Caption
            End If
        End If
    End If
    
    
    If mlng����ID <> 0 Then
        If tabClass.SelectedItem.Index = conTabδ�� Then 'Ĭ����ʾδ�����,���������ĳ�ν��ʵķ���
            '�����ǰ���˵���Ժʱ����ת��ʱ��֮ǰ,����Ҫ��������ݱ��ѯ
            '��Ժʱ��,�����ǰѡ����ǽ�������,�ڱ���������else��,ȡ��ǰ��ѡ�Ľ��������ж�
            blnDateMoved = mblnDateMoved
        Else
            blnDateMoved = zlDatabase.DateMoved(Format(Mid(tabClass.SelectedItem.ToolTipText, InStr(1, tabClass.SelectedItem.ToolTipText, ":") + 1), "yyyy-MM-dd hh:mm:ss"), , , Caption)
            
            lng����ID = Val(tabClass.SelectedItem.Tag)
        End If
    End If
    
    '���øſ�
    If blnHead Then Call LoadFeeOutline(mlng����ID, blnDateMoved, lng����ID)
    'Ԥ�����嵥
    If blnDeposit Then Call LoadDeposit(mlng����ID, blnDateMoved, lng����ID)
    
    On Error GoTo errH
    '���˺�:24913,  mbytDateType:1-����ʱ��,2-�Ǽ�ʱ��
    zlGetDateRange , strStartDate, strEndDate
    If strStartDate <> "" Then
        strWhere = IIf(mbytDateType = 1, " And  (A.����ʱ�� between [4] and [5] ) ", " And  (A.�Ǽ�ʱ�� between [4] and [5] )")
    Else
        strStartDate = "1901-01-01": strEndDate = "3000-01-01"
    End If
   
    Select Case mbytList
        Case ListType.C0�����嵥, ListType.C1�ֿ�����ϸ, ListType.C2����Ŀ��ϸ, ListType.C3�������ϸ, ListType.C4���������ϸ '��ϸ�嵥,�ֿ���ϸ,��Ŀ��ϸ,������ϸ,(��������Ŀ(���վݷ�Ŀ),�շ���Ŀ,��ϸ�ּ���ѯ)
            strWhere = strWhere & IIf(chk����ʾ���ʵ���.Value = 0, "", " And Exists (Select 1 From סԺ���ü�¼ Where NO = a.No And Mod(��¼����, 10) = Mod(a.��¼����, 10) And ��� = a.��� And ��¼״̬ = 3)")
        Case ListType.C5����Ŀ����, ListType.C6��������, ListType.C7���·������, ListType.C8���յ��ݻ���, ListType.C9���շ�Ŀ����
            If cboDept.ListIndex > 0 And Not blnLoadDept Then   '0�����в���
                strDept = " And A.��������id = [6]"
                strWhere = strWhere & strDept
                lng��������ID = cboDept.ItemData(cboDept.ListIndex)
            End If
    End Select
    
    lngPre����ID = 0
    strPre����Text = cboDept.Text   ''43494
    If cboDept.ListIndex >= 0 Then
        lngPre����ID = cboDept.ItemData(cboDept.ListIndex)
    End If
    If blnLoadDept Then
        cboDept.Clear
        cboDept.AddItem "���п���"
    End If
    
    
    If blnMoney Then
        If lng����ID = 0 Then 'Ĭ����ʾδ�����,���������ĳ�ν��ʵķ���
            
            strIF = " And A.��¼״̬<>0 And A.���ʷ���=1" & strBaby & _
                    IIf(mvs.CheckFee, "", " And A.�����־<>4") & _
                    IIf(chkAdivce.Value = 0, "", " And A.ҽ����� is Null") & strWhere
            
            strWhere��ҳ = IIf(tabTime.SelectedItem.Index = 1, "", " And A.��ҳID=[2]")
            
            If mvs.ZeroFee Or (chk����ʾ���ʵ���.Value = 1 And chk����ʾ���ʵ���.Visible) Then
                '61527        Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and Sum(Nvl(A.���ʽ��,0)) =0 And (Mod(Count(*),2)=0 or  sum(decode(a.����ID,null,0,0,1)) = 0)) " & _
                '       :������count(*)=1������,��Ϊ���ڽ���һ��ʱ,ҲҪ��ʾ,��Ӧ��ʾ�Ŷ�,�ֵ���Ϊ:sum(decode(a.����ID,null,0,0,1)=0
            
                str���˷��� = _
                    "  Select Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.�շ����,A.������,A.��������ID,A.ִ�в���ID,A.���㵥λ,Max(A.ժҪ) as ժҪ,Max(A.���ձ���) as ���ձ���," & _
                    "       A.����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������,Decode(Nvl(A.ҽ�����,0),0,0,(Decode(Sign(A.����),-1,1,0))) �����ջ�,Nvl(A.�۸񸸺�,A.���) as ���,A.ִ��״̬ as ִ��״̬" & _
                    "  From סԺ���ü�¼ A" & _
                    "  Where A.����ID=[1]" & strIF & strWhere��ҳ & _
                    "           And (Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��,0) Or Nvl(A.���ʽ��, 0)=0)" & _
                    "  Having Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0)<>0 " & _
                    "           Or (Sum(Nvl(A.���ʽ��, 0)) = 0 And (Mod(Count(*),2)=0 Or sum(decode(a.����ID,null,0,0,1))=0))" & _
                    "  Group by A.NO,Mod(A.��¼����,10),Nvl(A.�۸񸸺�,A.���),A.����ʱ��,A.�Ǽ�ʱ��,A.��¼״̬,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.�շ����,A.ִ��״̬," & _
                    "          A.������,A.��������ID,A.ִ�в���ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.��������,Decode(Nvl(A.ҽ�����,0),0,0,(Decode(Sign(A.����),-1,1,0))),a.ҽ����� "
                
                If mblnContainOutFee Then
                    str���˷��� = str���˷��� & "  Union ALL " & Replace(str���˷���, "סԺ���ü�¼", "������ü�¼")
                End If
              
            Else
                strSQL = _
                    " Select Distinct NO,Mod(��¼����,10) as ��¼����" & _
                    " From סԺ���ü�¼ A" & _
                    " Where ����ID=[1]" & strIF & strWhere��ҳ & _
                    " Group by NO,Mod(��¼����,10),���" & _
                    " Having Nvl(Sum(ʵ�ս��),0)-Nvl(Sum(���ʽ��),0)<>0"
                    
                str���˷��� = _
                    " Select /*+ optimizer_features_enable('10.2.0.4') */ Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.�շ����,A.������,A.��������ID,A.ִ�в���ID,A.���㵥λ,Max(A.ժҪ) as ժҪ,Max(A.���ձ���) as ���ձ���," & _
                    "        A.����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������,Decode(Nvl(A.ҽ�����,0),0,0,(Decode(Sign(A.����),-1,1,0))) �����ջ�,Nvl(A.�۸񸸺�,A.���) as ���,A.ִ��״̬ as ִ��״̬" & _
                    " From סԺ���ü�¼ A," & _
                    "      (" & strSQL & ") B" & _
                    " Where A.NO=B.NO And Mod(A.��¼����,10)=B.��¼���� " & _
                    "       And A.����ID+0=[1]" & strIF & strWhere��ҳ & _
                    "       And Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��,0)" & _
                    " Having Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0)<>0" & _
                    " Group by A.NO,Mod(A.��¼����,10),Nvl(A.�۸񸸺�,A.���),A.����ʱ��,A.�Ǽ�ʱ��,A.��¼״̬,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.�շ����,A.ִ��״̬ ," & _
                    "          A.������,A.��������ID,A.ִ�в���ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.��������,Decode(Nvl(A.ҽ�����,0),0,0,(Decode(Sign(A.����),-1,1,0))) "
                 
                 If mblnContainOutFee Then
                    str���˷��� = str���˷��� & " Union ALL " & Replace(str���˷���, "סԺ���ü�¼", "������ü�¼")
                 End If
            End If
            
            str�������� = ""
            str��� = " Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & gstrDec & "')) as δ����,"
            str���ϼ� = " Ltrim(To_Char(Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0),'999999999" & gstrDec & "')) as δ����,"
            str�����ջ� = "A.�����ջ�,"
        Else
            str���˷��� = "" & _
            " Select ����ID,����ʱ��,NO,���,ҽ�����,�۸񸸺�,��¼״̬,ִ��״̬,��������ID,ִ�в���id,�շ�ϸĿID,������Ŀid,������,����,����,���㵥λ,��׼����,���ʽ��,��¼����,��������,�վݷ�Ŀ,�շ����,����Ա����,�Ǽ�ʱ��,���ձ���,ժҪ  " & _
            " From סԺ���ü�¼ A" & _
            " Where A.����ID=[1]  " & strBaby & IIf(chkAdivce.Value = 0, "", " And A.ҽ����� is Null") & strWhere
            
            str���˷��� = str���˷��� & vbCrLf & " UNION ALL " & vbCrLf & Replace(str���˷���, "סԺ���ü�¼", "������ü�¼")
            If blnDateMoved Then
                str���˷��� = str���˷��� & vbCrLf & " UNION ALL " & vbCrLf & _
                    Replace(Replace(str���˷���, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
            End If
            
            str�������� = " And A.����ID=[1]" & strBaby & IIf(chkAdivce.Value = 0, "", " And A.ҽ����� is Null") & strWhere
            str��� = " Ltrim(To_Char(A.���ʽ��,'999999999" & gstrDec & "')) as ���ʽ��,"
            str���ϼ� = " Ltrim(To_Char(Nvl(Sum(A.���ʽ��),0),'999999999" & gstrDec & "')) as ���ʽ��,"
            str�����ջ� = "Decode(Nvl(A.ҽ�����,0),0,0,Decode(Sign(A.����),-1,1,0)) �����ջ�,"
        End If
        
        '28078:case  when trunc(A.����)=0 then  case when A.����>=0 then '0' else '-0' end when nvl(A.����,0)<0 then '-' else '' end||abs(A.����)
        '��Ҫ�Ǹ�ʽ.5�����.��ʾ��ʽ��0.5��-.5��ʾΪ-0.5
        
        Select Case mbytList
            Case ListType.C0�����嵥  '��ϸ�嵥
                strSQL = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.������," & _
                "       B.���� as ��������,E.���� as ִ�п���,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||case  when trunc(A.����)=0 then  case when A.����>=0 then '0' else '-0' end when nvl(A.����,0)<0 then '-' else '' end||abs(A.����)||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as ��׼���," & str��� & _
                "       Nvl(A.��������,C.��������) as ����,N.���� ҽ������,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & str�����ջ� & _
                "        Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.���,C.���� as ��Ŀ����,A.���ձ���,C.˵�� as ��Ŀ˵��,A.ժҪ,Decode(A.��¼״̬,2,'',Decode(A.ִ��״̬,0,'δִ��','��ִ��')) as ִ��״̬" & _
                " FROM (" & str���˷��� & ") A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D,���ű� E,����֧����Ŀ M,����֧������ N" & _
                " Where A.��������ID=B.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And C.ID=M.�շ�ϸĿID(+) And M.����(+)=" & IIf(lng����ID = 0, "[3]", "[2]") & " And M.����ID=N.ID(+)" & vbNewLine & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Order by ��������,���ݺ�,��Ŀ"
                
            Case ListType.C1�ֿ�����ϸ  '�ֿ���ϸ
                strSQL = _
                "SELECT " & IIf(str�������� = "��������", " B.���� as ��������,", " E.���� as ִ�п���,") & "To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.������," & _
                        IIf(str�������� = "��������", " E.���� as ִ�п���,", " B.���� as ��������,") & "Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||case  when trunc(A.����)=0 then  case when A.����>=0 then '0' else '-0' end when nvl(A.����,0)<0 then '-' else '' end||abs(A.����)||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as ��׼���," & str��� & _
                "       Nvl(A.��������,C.��������) as ����,N.���� ҽ������,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & str�����ջ� & _
                "       Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.���,C.���� as ��Ŀ����,A.���ձ���,C.˵�� as ��Ŀ˵��,A.ժҪ,Decode(A.��¼״̬,2,'',Decode(A.ִ��״̬,0,'δִ��','��ִ��')) as ִ��״̬" & _
                " FROM (" & str���˷��� & ") A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D,���ű� E,����֧����Ŀ M,����֧������ N" & _
                " Where A.��������ID=B.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=C.ID " & _
                " And C.ID=M.�շ�ϸĿID(+) And M.����(+)=" & IIf(lng����ID = 0, "[3]", "[2]") & " And M.����ID=N.ID(+)" & vbNewLine & _
                " And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Order by " & str�������� & ",��������,���ݺ�"
                
            Case ListType.C2����Ŀ��ϸ  '��Ŀ��ϸ
                strSQL = _
                " SELECT Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ') ���,To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.������," & _
                "        B.���� as ��������,E.���� as ִ�п���,A.�վݷ�Ŀ as ��Ŀ," & _
                "        Nvl(A.����,1)*A.���� as ����,A.���㵥λ," & _
                "        Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "        Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as ��׼���," & str��� & _
                "        Nvl(A.��������,C.��������) as ����,N.���� ҽ������,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & str�����ջ� & _
                "       Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.���,C.���� as ��Ŀ����,A.���ձ���,C.˵�� as ��Ŀ˵��,A.ժҪ,Decode(A.��¼״̬,2,'',Decode(A.ִ��״̬,0,'δִ��','��ִ��')) as ִ��״̬" & _
                " FROM (" & str���˷��� & ") A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D,���ű� E,����֧����Ŀ M,����֧������ N" & _
                " Where A.��������ID=B.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And C.ID=M.�շ�ϸĿID(+) And M.����(+)=" & IIf(lng����ID = 0, "[3]", "[2]") & " And M.����ID=N.ID(+)" & vbNewLine & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Order by ��Ŀ,���,��������,���ݺ�"
                
            Case ListType.C3�������ϸ  '������ϸ
                strSQL = _
                " SELECT A.�վݷ�Ŀ as ��Ŀ,To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.������," & _
                "       B.���� as ��������,E.���� as ִ�п���,Nvl(D.����,C.����) as ��Ŀ,C.���," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||case  when trunc(A.����)=0 then  case when A.����>=0 then '0' else '-0' end when nvl(A.����,0)<0 then '-' else '' end||abs(A.����)||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as ��׼���," & str��� & _
                "       Nvl(A.��������,C.��������) as ����,N.���� ҽ������,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & str�����ջ� & _
                "       Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.���,C.���� as ��Ŀ����,A.���ձ���,C.˵�� as ��Ŀ˵��,A.ժҪ,Decode(A.��¼״̬,2,'',Decode(A.ִ��״̬,0,'δִ��','��ִ��')) as ִ��״̬" & _
                " FROM (" & str���˷��� & ") A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D,���ű� E,����֧����Ŀ M,����֧������ N" & _
                " Where A.��������ID=B.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And C.ID=M.�շ�ϸĿID(+) And M.����(+)=" & IIf(lng����ID = 0, "[3]", "[2]") & " And M.����ID=N.ID(+)" & vbNewLine & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Order by ��Ŀ,��������,���ݺ�"
            Case ListType.C4���������ϸ '��������Ŀ(���վݷ�Ŀ),�շ���Ŀ,��ϸ�ּ���ѯ
                If str���෽ʽ = "������Ŀ" Then
                    If lng����ID = 0 Then
                        str���˷��� = " With PatiFee as (" & str���˷��� & ")"
                        
                        str���˷��� = str���˷��� & "" & _
                        " Select  distinct A.����ʱ��,A.NO,A.������,A.�վݷ�Ŀ," & _
                        "        A.����*decode(A.��¼����,12,0,13,0,1) as ����,A.����*decode(A.��¼����,12,0,13,0,1) as ����, " & _
                        "       A.���㵥λ,A.��׼����,A.ʵ�ս��,A.���ʽ��,A.��������,A.����Ա����,A.�Ǽ�ʱ��," & _
                        "       A.ҽ�����,A.�۸񸸺�,A.���ձ��� ,A.���,A.��¼״̬,A.ִ��״̬,A.��������ID,A.ִ�в���id,A.�շ�ϸĿID,A.������Ŀid,A.��¼����,A.ժҪ  " & _
                        " From סԺ���ü�¼ A ,PatiFee G" & _
                        " Where A.NO = G.NO And Mod(A.��¼����,10)=G.��¼���� And Nvl(A.�۸񸸺�,A.���)=G.��� And A.��¼״̬<>0 "
                        If mblnContainOutFee Then
                            str���˷��� = str���˷��� & " Union ALL " & _
                            " Select distinct  A.����ʱ��,A.NO,A.������,A.�վݷ�Ŀ, " & _
                            "        A.����*decode(A.��¼����,12,0,13,0,1) as ����,A.����*decode(A.��¼����,12,0,13,0,1) as ����, " & _
                            "       A.���㵥λ,A.��׼����,A.ʵ�ս��,A.���ʽ��,A.��������,A.����Ա����,A.�Ǽ�ʱ��," & _
                            "       A.ҽ�����,A.�۸񸸺�,A.���ձ��� ,A.���,A.��¼״̬,A.ִ��״̬,A.��������ID,A.ִ�в���id,A.�շ�ϸĿID,A.������Ŀid,A.��¼����,A.ժҪ  " & _
                            " From ������ü�¼ A ,PatiFee G" & _
                            " Where A.NO = G.NO And Mod(A.��¼����,10)=G.��¼���� And Nvl(A.�۸񸸺�,A.���)=G.��� And A.��¼״̬<>0 " & strWhereCheckFee & _
                            ""
                        End If
                        
                        '�������Ϻ������ʣ���¼����Ϊ12�ļ�¼״̬��1,��¼����Ϊ2�ļ�¼״̬��3,Ҫ����������ҪDecode
                        strSQL = "Select F.���� ������Ŀ, Nvl(D.����, C.����) As �շ���Ŀ, To_Char(A.����ʱ��, 'YYYY-MM-DD') As ��������, A.NO As ���ݺ�,A.������," & vbNewLine & _
                            "       B.���� As ��������, E.���� As ִ�п���, C.���, A.�վݷ�Ŀ As ��Ŀ," & vbNewLine & _
                            "       sum(Nvl(A.����,1)*A.����)  as ����,A.���㵥λ," & vbNewLine & _
                            "       LTrim(To_Char(Nvl(A.��׼����, 0), '999999999" & gstrFeePrecisionFmt & "')) As ��׼����," & vbNewLine & _
                            "       LTrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) As ��׼���," & str���ϼ� & vbNewLine & _
                            "       Nvl(A.��������, C.��������) As ����,N.���� ҽ������, A.����Ա���� As ����Ա," & vbNewLine & _
                            "       To_Char(A.�Ǽ�ʱ��, 'YYYY-MM-DD HH24:MI:SS') As �Ǽ�ʱ��,Decode(Nvl(A.ҽ�����,0),0,0,Decode(Sign(A.����),-1,1,0)) �����ջ�," & _
                            "       Mod(A.��¼����,10) as ��¼����,Decode(A.��¼״̬,3,1,A.��¼״̬) as ��¼״̬,Nvl(A.�۸񸸺�,A.���) ���,max(C.����) as ��Ŀ����, max(A.���ձ���) ���ձ���," & _
                            "       max(C.˵��) as ��Ŀ˵��,max(A.ժҪ) as ժҪ,Decode(Decode(A.��¼״̬,3,1,A.��¼״̬),2,'',Decode(Max(A.ִ��״̬),0,'δִ��','��ִ��')) as ִ��״̬" & vbNewLine & _
                            "From  (" & str���˷��� & ") A, ���ű� B, �շ���ĿĿ¼ C, �շ���Ŀ���� D, ���ű� E, ������Ŀ F,����֧����Ŀ M,����֧������ N" & vbNewLine & _
                            "Where   A.��������id = B.ID(+) And A.ִ�в���id = E.ID(+) And A.�շ�ϸĿid = C.ID And A.������Ŀid = F.ID " & vbNewLine & _
                            "      And C.ID=M.�շ�ϸĿID(+) And M.����(+)=[3] And M.����ID=N.ID(+)" & vbNewLine & _
                            "      And A.�շ�ϸĿid = D.�շ�ϸĿid(+) And D.����(+) = 1 And D.����(+) = " & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & vbNewLine & _
                            " Group By F.����, Nvl(D.����, C.����), To_Char(A.����ʱ��, 'YYYY-MM-DD'), A.NO,A.������,B.����, E.����, C.���, " & vbNewLine & _
                            "          A.�վݷ�Ŀ,A.���㵥λ,Nvl(A.��׼����, 0),Nvl(A.��������, C.��������),N.����, A.����Ա����," & vbNewLine & _
                            "          A.�Ǽ�ʱ��,Decode(Nvl(A.ҽ�����,0),0,0,Decode(Sign(A.����),-1,1,0)),Mod(A.��¼����,10),Decode(A.��¼״̬,3,1,A.��¼״̬),Nvl(A.�۸񸸺�,A.���)" & vbNewLine & _
                            " Order By ������Ŀ,�շ���Ŀ,��������"
                    Else
                        strSQL = "Select F.���� ������Ŀ, Nvl(D.����, C.����) As �շ���Ŀ, To_Char(A.����ʱ��, 'YYYY-MM-DD') As ��������, A.NO As ���ݺ�,A.������," & vbNewLine & _
                            "       B.���� As ��������, E.���� As ִ�п���, C.���, A.�վݷ�Ŀ As ��Ŀ," & vbNewLine & _
                            "       Nvl(A.����,1)*A.���� as ����,A.���㵥λ," & vbNewLine & _
                            "       LTrim(To_Char(Nvl(A.��׼����, 0), '999999999" & gstrFeePrecisionFmt & "')) As ��׼����," & vbNewLine & _
                            "       LTrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) As ��׼���," & str��� & vbNewLine & _
                            "       Nvl(A.��������, C.��������) As ����,N.���� ҽ������, A.����Ա���� As ����Ա," & vbNewLine & _
                            "       To_Char(A.�Ǽ�ʱ��, 'YYYY-MM-DD HH24:MI:SS') As �Ǽ�ʱ��," & str�����ջ� & _
                            "       Mod(A.��¼����,10) as ��¼����,A.��¼״̬,Nvl(A.�۸񸸺�,A.���) ���,C.���� as ��Ŀ����,A.���ձ���,C.˵�� as ��Ŀ˵��,A.ժҪ,Decode(A.��¼״̬,2,'',Decode(A.ִ��״̬,0,'δִ��','��ִ��')) as ִ��״̬" & vbNewLine & _
                            "From (" & str���˷��� & ") A, ���ű� B, �շ���ĿĿ¼ C, �շ���Ŀ���� D, ���ű� E, ������Ŀ F,����֧����Ŀ M,����֧������ N" & vbNewLine & _
                            "Where A.��������id = B.ID(+) And A.ִ�в���id = E.ID(+) And A.�շ�ϸĿid = C.ID And A.������Ŀid = F.ID " & vbNewLine & _
                            "      And C.ID=M.�շ�ϸĿID(+) And M.����(+)=[2] And M.����ID=N.ID(+)" & vbNewLine & _
                            "      And A.�շ�ϸĿid = D.�շ�ϸĿid(+) And D.����(+) = 1 And D.����(+) = " & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & vbNewLine & _
                            "Order By ������Ŀ,�շ���Ŀ,��������"
                    End If
                Else
                    strSQL = "Select A.�վݷ�Ŀ As ��Ŀ, Nvl(D.����, C.����) As �շ���Ŀ, To_Char(A.����ʱ��, 'YYYY-MM-DD') As ��������, A.NO As ���ݺ�,A.������," & vbNewLine & _
                            "       B.���� As ��������, E.���� As ִ�п���, C.���, " & vbNewLine & _
                            "       Nvl(A.����,1)*A.���� as ����,A.���㵥λ," & vbNewLine & _
                            "       LTrim(To_Char(Nvl(A.��׼����, 0), '999999999" & gstrFeePrecisionFmt & "')) As ��׼����," & vbNewLine & _
                            "       LTrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) As ��׼���," & str��� & vbNewLine & _
                            "       Nvl(A.��������, C.��������) As ����,N.���� ҽ������, A.����Ա���� As ����Ա," & vbNewLine & _
                            "       To_Char(A.�Ǽ�ʱ��, 'YYYY-MM-DD HH24:MI:SS') As �Ǽ�ʱ��," & str�����ջ� & _
                            "       Mod(A.��¼����,10) as ��¼����,A.��¼״̬,���,C.���� as ��Ŀ����,A.���ձ���,C.˵�� as ��Ŀ˵��,A.ժҪ,Decode(A.��¼״̬,2,'',Decode(A.ִ��״̬,0,'δִ��','��ִ��')) as ִ��״̬" & vbNewLine & _
                            "From (" & str���˷��� & ") A, ���ű� B, �շ���ĿĿ¼ C, �շ���Ŀ���� D, ���ű� E,����֧����Ŀ M,����֧������ N" & vbNewLine & _
                            "Where A.��������id = B.ID(+) And A.ִ�в���id = E.ID(+) And A.�շ�ϸĿid = C.ID " & vbNewLine & _
                            "      And C.ID=M.�շ�ϸĿID(+) And M.����(+)=" & IIf(lng����ID = 0, "[3]", "[2]") & " And M.����ID=N.ID(+)" & vbNewLine & _
                            "      And A.�շ�ϸĿid = D.�շ�ϸĿid(+) And D.����(+) = 1 And D.����(+) = " & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & str�������� & vbNewLine & _
                            "Order By ��Ŀ,�շ���Ŀ,��������"
                End If
                
            Case ListType.C5����Ŀ����  '��Ŀ����
                strSQL = "case  when trunc(Sum(Nvl(A.����,1)*Nvl(A.����,1)))=0 then  case when Sum(Nvl(A.����,1)*Nvl(A.����,1))>0 then  '0' when Sum(Nvl(A.����,1)*Nvl(A.����,1))=0 then '' else '-0' end when Sum(Nvl(A.����,1)*Nvl(A.����,1))<0 then '-' else '' end||abs(Sum(Nvl(A.����,1)*Nvl(A.����,1)))"
                
                strSQL = _
                " SELECT nvl(Q.���,A.�շ����) as �շ����, Nvl(D.����,C.����) as ��Ŀ,C.���," & strSQL & "||Max(A.���㵥λ) ����," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as ��׼���," & Mid(str���ϼ�, 1, Len(str���ϼ�) - 1) & _
                " FROM (" & str���˷��� & ") A,�շ���ĿĿ¼ C,�շ���Ŀ���� D, �շ���� Q" & _
                " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And A.�շ����=Q.����(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Group by nvl(Q.���,A.�շ����), Nvl(D.����,C.����),���" & _
                " Order by �շ����,��Ŀ,���"
                    
            Case ListType.C6��������  '�������
                'If str�������� <> "" Then str�������� = " Where " & Mid(str��������, InStr(1, str��������, "And") + 3)
                strSQL = _
                " SELECT A.�վݷ�Ŀ as ��Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as ��׼���," & Mid(str���ϼ�, 1, Len(str���ϼ�) - 1) & _
                " FROM (" & str���˷��� & ") A " & _
                " Group by A.�վݷ�Ŀ " & _
                " Order by ��Ŀ"
            Case ListType.C7���·������  '���»���
                strSQL = _
                " SELECT B.�ڼ�,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as ��׼���," & Mid(str���ϼ�, 1, Len(str���ϼ�) - 1) & _
                " FROM (" & str���˷��� & ") A,�ڼ�� B" & _
                " Where A." & strDateMode & "ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                " Group by B.�ڼ�,A.�վݷ�Ŀ" & _
                " Order by �ڼ�,��Ŀ"
                    
            Case ListType.C8���յ��ݻ���  '���շ���
                'If str�������� <> "" Then str�������� = " Where " & Mid(str��������, InStr(1, str��������, "And") + 3)
                strSQL = _
                " SELECT TO_Char(A." & strDateMode & "ʱ��,'YYYY-MM-DD') as " & strDateMode & "����,A.NO as ���ݺ�,A.�վݷ�Ŀ as ������Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as ��׼���," & str���ϼ� & _
                "       A.����Ա���� as ����Ա,A.��¼����" & _
                " FROM (" & str���˷��� & ") A" & _
                " Group by TO_Char(A." & strDateMode & "ʱ��,'YYYY-MM-DD'),A.NO,A.��¼����,A.�վݷ�Ŀ,A.����Ա����" & _
                " Order by " & strDateMode & "����,��¼���� desc,���ݺ�,������Ŀ"
            
            Case ListType.C9���շ�Ŀ����  '���շ�Ŀ
                'If str�������� <> "" Then str�������� = " Where " & Mid(str��������, InStr(1, str��������, "And") + 3)
                strSQL = _
                " SELECT TO_Char(A." & strDateMode & "ʱ��,'YYYY-MM-DD') as " & strDateMode & "����,A.�վݷ�Ŀ as ������Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as ��׼���," & Mid(str���ϼ�, 1, Len(str���ϼ�) - 1) & _
                " FROM (" & str���˷��� & ") A" & _
                " Group by TO_Char(A." & strDateMode & "ʱ��,'YYYY-MM-DD'),A.�վݷ�Ŀ" & _
                " Order by " & strDateMode & "����,������Ŀ"
        End Select
                    
        If lng����ID = 0 Then
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID, lng��ҳID, mintInsure, CDate(strStartDate), CDate(strEndDate), lng��������ID)
        Else
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID, mintInsure, 0, CDate(strStartDate), CDate(strEndDate), lng��������ID)
        End If
        If mbytList = ListType.C0�����嵥 Or mbytList = ListType.C2����Ŀ��ϸ Then Call LoadCbo��Ŀ
        
        If mbytList <> ListType.C1�ֿ�����ϸ Then
            If mbytList = ListType.C0�����嵥 Or mbytList = ListType.C2����Ŀ��ϸ Or mbytList = ListType.C3�������ϸ Or mbytList = ListType.C4���������ϸ Then
                mblnNotClick = True
                Call LoadCbo��������(mrsList, False)
                If strPre����Text <> "" Then
                    Call zlControl.CboLocate(cboDept, strPre����Text)
                End If
                mblnNotClick = False
            ElseIf blnLoadDept Then    '������ܣ��ض����������б�(�ڵ�ǰ��ѯ���ѡ��һ������ʱ�������ض�)
                strSQL = "Select Distinct B.���� as ��������,a.��������ID From (" & Replace(str���˷���, strDept, "") & ") A,���ű� B Where a.��������ID = b.ID"
                If lng����ID = 0 Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, mlng����ID, lng��ҳID, mintInsure, CDate(strStartDate), CDate(strEndDate), lng��������ID)
                Else
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID, mintInsure, 0, CDate(strStartDate), CDate(strEndDate), lng��������ID)
                End If
                mblnNotClick = True
                Call LoadCbo��������(rsTmp, True)
                If strPre����Text <> "" Then
                    Call zlControl.CboLocate(cboDept, strPre����Text)
                End If
                mblnNotClick = False
            End If
        End If
        
        With vsfFee
            .Redraw = flexRDNone
            .Clear
                        
            '�����ڰ�����ǰ����,�����ʹ��ʱҪ����Ϊ1
            If mbytList = ListType.C0�����嵥 Or mbytList = ListType.C5����Ŀ���� Or mbytList = ListType.C6�������� Then
                .FixedCols = 0
            Else
                .FixedCols = 1
                .OutlineCol = 0
                .OutlineBar = flexOutlineBarComplete
            End If
            Set .DataSource = mrsList
            Call SetVsffeeFormat
            
            '�ָ����Ի�����
            zl_vsGrid_Para_Restore mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False
            
            If mbytList = ListType.C0�����嵥 Or mbytList = ListType.C1�ֿ�����ϸ Or mbytList = ListType.C2����Ŀ��ϸ Or mbytList = ListType.C3�������ϸ Or mbytList = ListType.C4���������ϸ Then
            
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                .ColData(.ColIndex("��¼����")) = "-1||2"
                .ColData(.ColIndex("��¼״̬")) = "-1||2"
                .ColData(.ColIndex("���")) = "-1||2"
                .ColData(.ColIndex("�����ջ�")) = "-1||2"
                
                .ColWidth(.ColIndex("��¼����")) = 0    '�����ColHidden��ʽ,��ӡԤ���Կɼ�
                .ColWidth(.ColIndex("��¼״̬")) = 0
                .ColWidth(.ColIndex("���")) = 0
                .ColWidth(.ColIndex("�����ջ�")) = 0
                
                .ColHidden(.ColIndex("��¼����")) = True
                .ColHidden(.ColIndex("��¼״̬")) = True
                .ColHidden(.ColIndex("���")) = True
                .ColHidden(.ColIndex("�����ջ�")) = True
                If .ColIndex("������") >= 0 Then
                    '����:35710
                    If InStr(1, mstrPrivs, ";ҽ����ѯ;") = 0 Then
                        .ColHidden(.ColIndex("������")) = True
                        .ColWidth(.ColIndex("������")) = 0
                        .ColData(.ColIndex("������")) = "-1||2"
                    End If
                End If
                
                If mintInsure = 0 Then
                    .ColWidth(.ColIndex("ҽ������")) = 0
                    .ColHidden(.ColIndex("ҽ������")) = True
                    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                    .ColData(.ColIndex("ҽ������")) = "-1||2"
                End If
            ElseIf mbytList = ListType.C8���յ��ݻ��� Then
                .ColWidth(.ColIndex("��¼����")) = 0
                .ColHidden(.ColIndex("��¼����")) = True
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                .ColData(.ColIndex("��¼����")) = "-1||2"
            End If
            'If .Rows = 1 Then .Rows = .FixedRows + 1
            
            .Redraw = flexRDDirect
        End With
    End If
    If mstrRestoreFeeCons <> "" Then
        If zlRestoreFeeControls(mlng����ID) = False Then
            Call zlCommFun.StopFlash
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    '55107
    Select Case mbytList
    Case ListType.C0�����嵥, ListType.C1�ֿ�����ϸ, ListType.C2����Ŀ��ϸ, ListType.C3�������ϸ, ListType.C4���������ϸ
        Call FilterDetail
    End Select
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    
    LoadCardData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    vsfFee.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetVsffeeFormat()
    Dim i As Long, j As Long, lng�����ջ� As Long, lng��׼��� As Long, lngδ���� As Long, lng���� As Long, lng���� As Long
    Dim arrTotal(2) As Currency, strTmp As String, lng��¼״̬ As Long
    Dim blnSetColor As Boolean
    Dim bln������ As Boolean
    
    With vsfFee
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignGeneral
            .ColKey(i) = Trim(.TextMatrix(0, i))
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            .ColData(i) = "0||0"
        Next
        If .Rows <= 1 Then .Rows = 2
        lng���� = .ColIndex("����")
        If lng���� >= 0 Then .ColAlignment(lng����) = flexAlignRightCenter: .ColFormat(lng����) = "#######0.#####"
    
        If mrsList.RecordCount = 0 Then Exit Sub
        
        lng��׼��� = .ColIndex("��׼���")
        If lng��׼��� < 0 Then lng��׼��� = VsfGetColNum(vsfFee, "��׼���")  '��SQL��ʹ����Group by��,�󶨵������е�Keyû���Զ�����,��colindex��ʽȡ������
        
        If Val(tabClass.SelectedItem.Tag) = 0 Then
            lngδ���� = .ColIndex("δ����")
            If lngδ���� < 0 Then lngδ���� = VsfGetColNum(vsfFee, "δ����")
        Else
            lngδ���� = .ColIndex("���ʽ��")
            If lngδ���� < 0 Then lngδ���� = VsfGetColNum(vsfFee, "���ʽ��")
        End If
               
        Select Case mbytList
            Case ListType.C0�����嵥, ListType.C5����Ŀ����, ListType.C6��������
                .Subtotal flexSTSum, -1, lng��׼���, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "�ܼ�"
                .Subtotal flexSTSum, -1, lngδ����, "#######" & gstrDec
                .MergeRow(.Rows - 1) = True
            Case ListType.C2����Ŀ��ϸ
                bln������ = chk��������С��.Value = 1
                .Subtotal flexSTSum, 0, lng��׼���, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "�ܼ�"
                .Subtotal flexSTSum, 1, lng��׼���, "#######" & gstrDec, &HF5F5F5, vbBlack, True, IIf(bln������, "�ϼ�", "С��")
               If bln������ Then
                 .Subtotal flexSTSum, 2, lng��׼���, "#######" & gstrDec, &HF5F5F5, vbBlack, True, "С��"
                End If
                .Subtotal flexSTSum, 0, lngδ����, "#######" & gstrDec
                .Subtotal flexSTSum, 1, lngδ����, "#######" & gstrDec
                
                If bln������ Then
                    .Subtotal flexSTSum, 2, lngδ����, "#######" & gstrDec
                    .Subtotal flexSTSum, 1, lng����, "#######0.#####"
                    .Subtotal flexSTSum, 2, lng����, "#######0.#####"
                Else
                    .Subtotal flexSTSum, 1, lng����, "#######0.#####"
                End If
                
                
                .MergeCol(1) = True
                .MergeRow(.Rows - 1) = True
            Case ListType.C1�ֿ�����ϸ, ListType.C3�������ϸ, ListType.C7���·������, ListType.C9���շ�Ŀ����
                                
                If mbytList = ListType.C1�ֿ�����ϸ Or mbytList = ListType.C3�������ϸ Then strTmp = "%s "
                
                .Subtotal flexSTSum, 0, lng��׼���, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "�ܼ�"
                .Subtotal flexSTSum, 1, lng��׼���, "#######" & gstrDec, &HF5F5F5, vbBlack, True, strTmp & "С��"
                            
                .Subtotal flexSTSum, 0, lngδ����, "#######" & gstrDec
                .Subtotal flexSTSum, 1, lngδ����, "#######" & gstrDec
                If mbytList = ListType.C2����Ŀ��ϸ Then .Subtotal flexSTSum, 1, lng����, "#######0.#####"
                .MergeCol(1) = True
                .MergeRow(.Rows - 1) = True
                
            Case ListType.C4���������ϸ, ListType.C8���յ��ݻ���
                               
                If mbytList = ListType.C4���������ϸ Then strTmp = "%s "
                
                .Subtotal flexSTSum, 0, lng��׼���, "#######" & gstrDec, &HFFC0C0, vbBlack, True, "�ܼ�"
                .Subtotal flexSTSum, 1, lng��׼���, "#######" & gstrDec, &HF1E8FC, vbBlack, True, strTmp & "�ϼ�"
                .Subtotal flexSTSum, 2, lng��׼���, "#######" & gstrDec, &HF5F5F5, vbBlack, True, "С��"
                            
                .Subtotal flexSTSum, 0, lngδ����, "#######" & gstrDec
                .Subtotal flexSTSum, 1, lngδ����, "#######" & gstrDec
                .Subtotal flexSTSum, 2, lngδ����, "#######" & gstrDec
                          
                If mbytList = ListType.C4���������ϸ Then .Subtotal flexSTSum, 2, lng����, "#######0.#####"
                
                .MergeCol(1) = True
                .MergeCol(2) = True
                
        End Select
             
        lng�����ջ� = .ColIndex("�����ջ�")
        lng��¼״̬ = .ColIndex("��¼״̬") ' '30289
        lng���� = .ColIndex("����")
        If lng�����ջ� >= 0 Or lng��¼״̬ >= 0 Then
            For i = 1 To .Rows - 1
                blnSetColor = False
                
                If lng�����ջ� >= 0 Then
                    If Val(.TextMatrix(i, lng�����ջ�)) = 1 Then blnSetColor = True
                End If
                If blnSetColor = False And lng��¼״̬ >= 0 Then
                    If Val(.TextMatrix(i, lng��¼״̬)) = 2 Then blnSetColor = True
                    '����ҲҪ�ú�ɫ��ʾ
                    If blnSetColor = False And InStr(1, "31", Val(.TextMatrix(i, lng��¼״̬))) > 0 And lng���� >= 0 Then
                        If Left(Trim(.TextMatrix(i, lng����)), 1) = "-" Then blnSetColor = True
                    End If
                End If
                If blnSetColor Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC0  '�����ջ���,�����ɫ��ʾ,��.FillStyle = flexFillRepeat
                End If
            
            Next
        End If
        
        .AutoSize 0, .Cols - 1
    End With
End Sub

Private Sub LoadCbo��Ŀ()
    Dim str��Ŀ As String
    Dim i As Integer
    Dim strOld As String
    Dim strPreText As String
    strPreText = cboFeeType.Text    '43494
    cboFeeType.Clear '��ʱ��Ϊ��Ŀ,ͳ�Ʒ�ʽΪ2ʱ��Ϊ��������
    
    str��Ŀ = ";���з�Ŀ"
    
    If Not mrsList Is Nothing Then
        If mrsList.RecordCount > 0 Then mrsList.MoveFirst
        Do While Not mrsList.EOF
            If strOld <> mrsList!��Ŀ Then
                If InStr(1, ";" & str��Ŀ & ";", ";" & mrsList!��Ŀ & ";") = 0 Then
                    str��Ŀ = str��Ŀ & ";" & mrsList!��Ŀ
                End If
                strOld = mrsList!��Ŀ
            End If
            mrsList.MoveNext
        Loop
    End If
    
    str��Ŀ = Mid(str��Ŀ, 2)
    mblnNotClick = True
    For i = 0 To UBound(Split(str��Ŀ, ";"))
        cboFeeType.AddItem Split(str��Ŀ, ";")(i)
    Next
    zlControl.CboSetIndex cboFeeType.hWnd, 0
    If strPreText <> "" Then zlControl.CboLocate cboFeeType, strPreText
    mblnNotClick = False

End Sub

Private Sub LoadCbo��������(ByRef rsTmp As ADODB.Recordset, ByVal blnAddID As Boolean)
    Dim str�������� As String
    Dim i As Integer
    Dim strOld As String
            
    If blnAddID Then
        For i = 0 To rsTmp.RecordCount - 1
            cboDept.AddItem rsTmp!��������
            cboDept.ItemData(cboDept.NewIndex) = rsTmp!��������ID
            
            If mblnClinicOrNurse And cboDept.ListIndex = -1 Then
                If rsTmp!��������ID = UserInfo.����ID Then
                    cboDept.ListIndex = cboDept.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Next
    Else
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If strOld <> rsTmp!�������� Then
                If InStr(1, ";" & str�������� & ";", ";" & rsTmp!�������� & ";") = 0 Then
                    str�������� = str�������� & ";" & rsTmp!��������
                End If
                strOld = rsTmp!��������
            End If
            rsTmp.MoveNext
        Loop
        str�������� = Mid(str��������, 2)
        
        For i = 0 To UBound(Split(str��������, ";"))
            cboDept.AddItem Split(str��������, ";")(i)
            If mblnClinicOrNurse And cboDept.ListIndex = -1 Then
                If Split(str��������, ";")(i) = UserInfo.�������� Then
                    cboDept.ListIndex = cboDept.NewIndex
                End If
            End If
        Next
    End If
    cboDept.Tag = "��ˢ��"
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then zlControl.CboSetIndex cboDept.hWnd, 0
    cboDept.Tag = ""
End Sub

Private Sub chkNotCheckFee_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub chk��������С��_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub chk����ʾ���ʵ���_Click()
    If mblnNotClick Then Exit Sub
    If Visible = False Then Exit Sub
    Call LoadCardData(False, False, True)
End Sub

Private Sub chk����ʾ���ʵ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdRefresh_Click()
    Call LoadCardData(False, False, True)
    If vsfFee.Enabled And vsfFee.Visible Then vsfFee.SetFocus
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Call SetCondition
    Call Form_Resize
    Call picDetail_Resize
    Call vsfFee_LostFocus
    Call mshDepost_LostFocus
    Call mshInsure_LostFocus
    RaiseEvent Activate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If tabClass.SelectedItem Is Nothing Then Exit Sub
            Call lblMoney_MouseDown(1, 0, 0, 0)
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, strTmp As String
    mstrPrivs = gstrPrivs
    mblnFisrtSetFontSize = True
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsTools.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsTools.VisualTheme = xtpThemeOffice2003
    cbsTools.EnableCustomization False
    Set cbsTools.Icons = zlCommFun.GetPubIcons
    
    fraDeptMode.BackColor = picDetail.BackColor
    fraTypeMode.BackColor = picDetail.BackColor
 
    mstr��ֹ���� = ""
    mblnContainOutFee = zlDatabase.GetPara("�����������", glngSys, mlngModul, "1") = "1"
    msngScale = CSng(zlDatabase.GetPara("�嵥����", glngSys, mlngModul, 0.75))
    mbytDateType = IIf(zlDatabase.GetPara("����ʱ������", glngSys, mlngModul, "1") = "2", 2, 1)
    lbl����ʱ��.Caption = IIf(mbytDateType = 1, "����ʱ��", "�Ǽ�ʱ��")
    dtpEnd.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    dtpEnd.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    
    dtpBegin.Value = Format(DateAdd("m", -1, dtpEnd.Value), "yyyy-mm-dd 00:00:00")
    dtpBegin.MaxDate = dtpEnd.Value
    
    
    mvs.ReBalance = zlDatabase.GetPara("��ʾ��������", glngSys, mlngModul, "1") = "1"
    mvs.ZeroFee = zlDatabase.GetPara("��ʾ�����", glngSys, mlngModul, "0") = "1"
    mvs.CheckFee = zlDatabase.GetPara("��ʾ������", glngSys, mlngModul, "0") = "1"
    
    i = IIf(zlDatabase.GetPara("�ֿ�ģʽ", glngSys, mlngModul) = "1", 1, 0)
    optDeptMode(i).Value = True
    i = IIf(zlDatabase.GetPara("����ģʽ", glngSys, mlngModul) = "1", 1, 0)
    optTypeMode(i).Value = True
    
    mblnClinicOrNurse = isCliniOrNurse(UserInfo.����ID)
    
    mstrUnitIDs = GetUserUnits
    Call InitBaseData
    
    With vsfFee
        .ExplorerBar = flexExSortShowAndMove
        .FillStyle = flexFillRepeat
        .FixedRows = 1
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat  'ȱʡ��flexGridInset
                    
        .Subtotal flexSTClear
        .SubtotalPosition = flexSTBelow
    End With
    
    
    
End Sub

Private Sub Form_Resize()
    Dim tmpW As Long, tmpH As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    tmpW = IIf(mshInsure.Visible, mshInsure.Width + picLR.Width, 0)
    Call cbsTools.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    tabClass.Top = lngTop
    tabClass.Left = lngLeft
    picNum.Top = lngTop + 30
    tabClass.Width = Me.ScaleWidth
    If picNum.Enabled And picNum.Visible Then
        picNum.Left = IIf(995 + 1680 * (tabClass.Tabs.Count - 1) + 120 < Me.ScaleWidth - picNum.Width, 995 + 1680 * (tabClass.Tabs.Count - 1) + 120, Me.ScaleWidth - picNum.Width - 30)
        tabClass.Width = Me.ScaleWidth - picNum.Width - 30
    End If
    
    pic������Ϣ.Left = tabClass.Left
    pic������Ϣ.Top = tabClass.Top + tabClass.Height
    pic������Ϣ.Width = Me.ScaleWidth
    
    lblDepost.Top = pic������Ϣ.Top + pic������Ϣ.Height + 30
    lblDepost.Left = pic������Ϣ.Left + 15
    lblDepost.Width = pic������Ϣ.Width - 30 - tmpW
    mshDepost.Redraw = flexRDNone
    mshDepost.Top = lblDepost.Top + lblDepost.Height
    mshDepost.Left = pic������Ϣ.Left
    tmpH = (Me.ScaleHeight - tabClass.Height - pic������Ϣ.Height - lblDepost.Height - picDetail.Height - fraUD.Height - 30) * (1 - msngScale)
    If tmpH > 0 Then mshDepost.Height = tmpH
    mshDepost.Width = pic������Ϣ.Width - tmpW
    mshDepost.Redraw = flexRDBuffered
    
    picLR.Left = mshDepost.Left + mshDepost.Width
    picLR.Top = mshDepost.Top
    picLR.Height = mshDepost.Height
    
    lblInsure.Top = lblDepost.Top
    lblInsure.Left = picLR.Left + picLR.Width
    lblInsure.Width = mshInsure.Width - 30
    
    mshInsure.Top = mshDepost.Top
    mshInsure.Left = picLR.Left + picLR.Width
    mshInsure.Height = mshDepost.Height
    
    fraUD.Top = mshDepost.Top + mshDepost.Height
    fraUD.Left = pic������Ϣ.Left
    fraUD.Width = pic������Ϣ.Width
    
    picDetail.AutoRedraw = False
    picDetail.Top = fraUD.Top + fraUD.Height
    picDetail.Left = fraUD.Left
    picDetail.Width = fraUD.Width
    picDetail.AutoRedraw = True
    
    vsfFee.Redraw = flexRDNone
    vsfFee.Top = picDetail.Top + picDetail.Height
    vsfFee.Left = pic������Ϣ.Left
    vsfFee.Width = pic������Ϣ.Width
    vsfFee.Height = Me.ScaleHeight - lngTop - tabClass.Height - pic������Ϣ.Height - lblDepost.Height - picDetail.Height - fraUD.Height - mshDepost.Height - 30
    vsfFee.Redraw = flexRDDirect

    zlControl.PicShowFlat pic������Ϣ, -1, , taCenterAlign
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    zlDatabase.SetPara "�嵥����", msngScale, glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "����ʱ������", mbytDateType, glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "��ʾ��������", IIf(mvs.ReBalance, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "��ʾ�����", IIf(mvs.ZeroFee, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "��ʾ������", IIf(mvs.CheckFee, 1, 0), glngSys, mlngModul, mblnHavePara
    
    zlDatabase.SetPara "�ֿ�ģʽ", IIf(optDeptMode(0).Value, 0, 1), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "����ģʽ", IIf(optTypeMode(0).Value, 0, 1), glngSys, mlngModul, mblnHavePara
    
    Call zlDatabase.SetPara("���ò�ѯ��Χ", cbo����.Text & IIf(cbo����.Text = "�Զ��巶Χ", "|" & Format(dtpBegin.Value, "yyyy-mm-dd") & "," & Format(dtpEnd.Value, "yyyy-mm-dd"), ""), glngSys, mlngModul, mblnHavePara)
    Call zlDatabase.SetPara("��ϸ�������ʵ���", IIf(chk����ʾ���ʵ���.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara)
    Call zlDatabase.SetPara("��������ͳ��", IIf(chk��������С��.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara)
    '46646
    Call zlDatabase.SetPara("����������", IIf(chkNotCheckFee.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara)
    Call zlDatabase.SetPara("�����������", IIf(mblnContainOutFee, 1, 0), glngSys, mlngModul, mblnHavePara)
 
    '������Ի�����
    If mbytFontSize <> 9 Then
        zlControl.VSFSetFontSize vsfFee, 9
        zlControl.VSFSetFontSize mshDepost, 9
        zlControl.VSFSetFontSize mshInsure, 9
    End If
    
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False, , mblnHavePara
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False, , mblnHavePara
    Set mrsList = Nothing
    mbytList = ListType.C0�����嵥
    mintPreCard = 0: mintPreTime = 0
    Unload frmDailyListAsk
End Sub
 

Private Sub imgColSel_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picDetail.hWnd)
    lngLeft = vRect.Left + imgColSel.Left
    lngTop = vRect.Top + imgColSel.Height + imgColSel.Top
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfFee, lngLeft, lngTop, imgColSel.Height)
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False
End Sub
 

Private Sub mshDepost_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False
End Sub

Private Sub mshDepost_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, mshDepost, Me.Name, "mshDepost", False
End Sub


Private Sub mshInsure_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False
End Sub

Private Sub mshInsure_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, mshInsure, Me.Name, "mshInsure", False
End Sub

Private Sub pic������Ϣ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pic������Ϣ.ToolTipText = lbl������Ϣ.Caption
End Sub


Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDepost.Height + Y < 268 Or vsfFee.Height - Y < 1000 Then Exit Sub
        
        fraUD.Top = fraUD.Top + Y
        
        mshDepost.Height = mshDepost.Height + Y
        picLR.Height = picLR.Height + Y
        mshInsure.Height = mshInsure.Height + Y
                
        picDetail.Top = picDetail.Top + Y
        vsfFee.Top = vsfFee.Top + Y
        vsfFee.Height = vsfFee.Height - Y
        
        Refresh
        msngScale = vsfFee.Height / (Me.ScaleHeight _
            - tabClass.Height - pic������Ϣ.Height - lblDepost.Height - picDetail.Height - fraUD.Height - 45)
    End If
End Sub

Private Sub lblMoney_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim objPopup As CommandBarPopup
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(xtpControlButtonPopup, conMenu_View_DetailType, True, True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
'
'
'Private Sub SetActiveList(objAct As Object)
'    If objAct Is mshDepost Then
'        mshDepost.BackColorSel = &H800000
'        mshInsure.BackColorSel = &H808080
'        vsfFee.BackColorSel = &H808080
'    ElseIf objAct Is mshInsure Then
'        mshDepost.BackColorSel = &H808080
'        mshInsure.BackColorSel = &H800000
'        vsfFee.BackColorSel = &H808080
'    ElseIf objAct Is vsfFee Then
'        mshDepost.BackColorSel = &H808080
'        mshInsure.BackColorSel = &H808080
'        vsfFee.BackColorSel = &H800000
'    Else
'        mshDepost.BackColorSel = &H808080
'        mshInsure.BackColorSel = &H808080
'        vsfFee.BackColorSel = &H808080
'    End If
'    'Call mshDepost_EnterCell
'    Call mshInsure_EnterCell
'End Sub

Private Sub SetVsGrindSelColor(ByVal objGrid As Object, Optional blnLostFocus As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��������ɫ
    '���:objGrid-�������ؼ�
    '     blnLostFocus-����Ƴ�
    '����:
    '����:
    '����:���˺�
    '����:2009-07-28 18:08:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    objGrid.BackColorSel = IIf(blnLostFocus, &H808080, &H800000)       ' &H8000000F
End Sub


Private Function ReadInsureMoney(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    mshInsure.Redraw = flexRDNone
    mshInsure.Clear
    On Error GoTo errH
        
    strSQL = "Select ���㷽ʽ,To_Char(���,'9999999990.00') as ������" & _
        " From ����ģ����� Where ����ID=[1] And ��ҳID=[2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption, lng����ID, lng��ҳID)
    
    mshInsure.Clear
    Set mshInsure.DataSource = rsTmp
    If rsTmp.RecordCount = 0 Then mshInsure.Rows = 2
    
'    Call Grid.BandRec(mshInsure, rsTmp)
'    Call SetGridWidth(mshInsure, Me)        '���ȡ��,����û�����ó�ʼ�п�,��ӡ���쳣
   With mshInsure
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = IIf(i > 0, 7, 1)
        Next
        .Row = 1: .Col = 0
        .Redraw = flexRDBuffered
        '�ָ����Ի�����
        zl_vsGrid_Para_Restore mlngModul, mshInsure, Me.Name, "mshInsure", False
   End With
    ReadInsureMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    mshInsure.Redraw = flexRDBuffered
    Call SaveErrLog
End Function


Function CurrencyToStr(ByVal Number As Currency) As String
    Dim str1Ary As Variant, str2Ary As Variant
    Dim a As Long, B As Long  'ѭ������
    Dim tmp1 As String        '��ʱת��
    Dim tmp2 As String        '��ʱת�����
    Dim Point As Long         'С����λ��
    
    Number = Val(Trim(Number))
    If Number = 0 Then CurrencyToStr = "": Exit Function
    If Number <= -922337203685477# Or Number >= 922337203685477# Then
       Exit Function
    End If
    
    str1Ary = Split("�� Ҽ �� �� �� �� ½ �� �� ��")
    str2Ary = Split("�� �� Ԫ ʰ �� Ǫ �� ʰ �� Ǫ �� ʰ �� Ǫ �� ʰ ��")
    tmp1 = FormatEx(Number, 2)
    tmp1 = Replace(tmp1, "-", "")  '��ȥ����-����
    Point = InStr(tmp1, ".")       'ȡ��С����λ��
    If Point = 0 Then      '�����С���㣬��������
       B = Len(tmp1) + 2   '��2λС��
    Else
       B = Len(Left(tmp1, Point + 1))  '�������2λС��
    End If
    ''�Ƚ����������滻Ϊ����
    For a = 9 To 0 Step -1
        tmp1 = Replace(Replace(tmp1, a, str1Ary(a)), ".", "")
    Next
    For a = 1 To B
        B = B - 1
        If Mid(tmp1, a, 1) <> "" Then
           If B > UBound(str2Ary) Then Exit For
           tmp2 = tmp2 & Mid(tmp1, a, 1) & str2Ary(B)
        End If
    Next
    If tmp2 = "" Then CurrencyToStr = "": Exit Function
    
'    ''������Ϊ����ʽ�����㷨������ȥ����
'    For a = 1 To Len(tmp2)
'        tmp2 = Replace(tmp2, "����", "����")
'        tmp2 = Replace(tmp2, "����", "����")
'        tmp2 = Replace(tmp2, "��Ǫ", "��")
'        tmp2 = Replace(tmp2, "���", "��")
'        tmp2 = Replace(tmp2, "��ʰ", "��")
'        tmp2 = Replace(tmp2, "��Ԫ", "Ԫ")
'        tmp2 = Replace(tmp2, "����", "��")
'        tmp2 = Replace(tmp2, "����", "��")
'    Next
'    ''������Ϊ����ʽ�����㷨������ȥ����
    
    If Point = 1 Then tmp2 = "��Ԫ" + tmp2
    If Number < 0 Then tmp2 = "��" + tmp2
    If Point = 0 Then tmp2 = tmp2 + "��"
    CurrencyToStr = tmp2
End Function

Private Sub vsfFee_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False
End Sub

Private Sub vsfFee_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsfFee, Me.Name, "��ͷ��Ϣ-" & mbytList, False
End Sub

Private Sub vsfFee_GotFocus()
    Call SetVsGrindSelColor(vsfFee)
End Sub
Private Sub vsfFee_LostFocus()
    Call SetVsGrindSelColor(vsfFee, True)
End Sub

Private Sub vsfFee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If mblnUnBilling Then Call ExecUnBilling
    End If
End Sub



Private Sub ChangeList(blnRefreshData As Boolean)
    Dim strTmp As String
    Dim objControl As CommandBarControl
    '31159
    Set objControl = mcbsMain.ActiveMenuBar.FindControl(xtpControlButton, conMenu_View_DetailType * 10 + mbytList, True, True)
    If objControl Is Nothing Then Exit Sub
        
    strTmp = objControl.Caption
    If tabClass.SelectedItem.Index = conTabδ�� Then
        lblMoney.Caption = " δ��" & Left(strTmp, Len(strTmp) - 4) & "��"
    Else
        lblMoney.Caption = " ����" & Left(strTmp, Len(strTmp) - 4) & "��"
    End If
    
    If mbytList = ListType.C1�ֿ�����ϸ Or mbytList = ListType.C4���������ϸ Then
        fraDeptMode.Visible = mbytList = ListType.C1�ֿ�����ϸ
        fraTypeMode.Visible = mbytList = ListType.C4���������ϸ
        
        If mbytList = ListType.C1�ֿ�����ϸ Then
            optDeptMode(0).Caption = "��������"
            optDeptMode(1).Caption = "ִ�п���"
        Else
            optTypeMode(0).Caption = "������Ŀ"
            optTypeMode(1).Caption = "�վݷ�Ŀ"
        End If
    Else
        fraTypeMode.Visible = False
        fraDeptMode.Visible = False
    End If
    
    cboFeeType.Visible = (mbytList = ListType.C0�����嵥 Or mbytList = ListType.C2����Ŀ��ϸ)
    cboDept.Visible = (mbytList <> ListType.C1�ֿ�����ϸ)
    
    If cboFeeType.ListCount = 0 Then cboFeeType.AddItem "���з�Ŀ"
    
    If Visible Then
        Call SetCondition
        If blnRefreshData Then Call LoadCardData(False, False, True)
    End If
    
End Sub

Private Sub SetCondition()
    Dim lngLeft As Long
    Dim sngTop As Single
    
    picDetail.AutoRedraw = True
    sngTop = lblMoney.Top + (lblMoney.Height - chkAdivce.Height) \ 2
    chkAdivce.Left = lblMoney.Left + lblMoney.Width + 150
    chkAdivce.Top = sngTop
    sngTop = lblMoney.Top + (lblMoney.Height - cboFeeType.Height) \ 2
    lngLeft = chkAdivce.Left + chkAdivce.Width + 50
    If cboFeeType.Visible Then
        cboFeeType.Top = sngTop
        cboFeeType.Left = lngLeft
        lngLeft = cboFeeType.Left + cboFeeType.Width + 50
    ElseIf fraTypeMode.Visible Then
        fraTypeMode.Top = cboFeeType.Top + (cboFeeType.Height - fraTypeMode.Height) \ 2
        fraTypeMode.Left = lngLeft
        lngLeft = fraTypeMode.Left + fraTypeMode.Width + 50
        
        fraTypeMode.Width = optTypeMode(1).Width + optTypeMode(0).Width + 100
        optTypeMode(1).Left = optTypeMode(0).Left + optTypeMode(0).Width + 50
        
        
    ElseIf fraDeptMode.Visible Then
        fraDeptMode.Top = cboFeeType.Top + (cboFeeType.Height - fraDeptMode.Height) \ 2
        fraDeptMode.Left = lngLeft
        fraDeptMode.Width = optDeptMode(1).Width + optDeptMode(0).Width + 100
        optDeptMode(1).Left = optDeptMode(0).Left + optDeptMode(0).Width + 50
        
        lngLeft = fraDeptMode.Left + fraDeptMode.Width + 50
    Else
        lngLeft = chkAdivce.Left + chkAdivce.Width + 50
    End If
    
    If cboDept.Visible Then
        cboDept.Left = lngLeft
        lngLeft = cboDept.Left + cboDept.Width + 50
    End If
    
    tabTime.Top = cboFeeType.Top + (cboFeeType.Height - tabTime.Height) \ 2
    If cboBaby.Visible Then
        cboBaby.Left = lngLeft + 50
        tabTime.Left = cboBaby.Left + cboBaby.Width + 50
    Else
        tabTime.Left = lngLeft + 50
    End If
    sngTop = cboFeeType.Top + cboFeeType.Height + 50
    dtpBegin.Height = cbo����.Height: dtpEnd.Height = cbo����.Height
    dtpBegin.Top = sngTop: dtpEnd.Top = sngTop: cbo����.Top = sngTop
    cmdRefresh.Top = sngTop
    lbl����ʱ��.Top = sngTop + (cboFeeType.Height - lblMoney.Height) \ 2
    lbl��.Top = lbl����ʱ��.Top: lbl���ڷ�Χ.Top = lbl����ʱ��.Top
    chk����ʾ���ʵ���.Top = sngTop + (cboFeeType.Height - chk����ʾ���ʵ���.Height) \ 2
    lngLeft = lbl����ʱ��.Left + lbl����ʱ��.Width + 50
    cbo����.Left = lngLeft
    lngLeft = cbo����.Left + cbo����.Width + 50
    lbl���ڷ�Χ.Left = lngLeft: dtpBegin.Left = lngLeft
    lbl��.Left = dtpBegin.Width + dtpBegin.Left + 50
    dtpEnd.Left = lbl��.Left + lbl��.Width + 50
    cmdRefresh.Left = dtpEnd.Left + dtpEnd.Width + 50
    picDetail.AutoRedraw = False
End Sub

Private Sub mshDepost_GotFocus()
    Call SetVsGrindSelColor(mshDepost)
End Sub
Private Sub mshDepost_LostFocus()
    Call SetVsGrindSelColor(mshDepost, True)
End Sub


Private Sub mshInsure_GotFocus()
    Call SetVsGrindSelColor(mshInsure)
End Sub
Private Sub mshInsure_LostFocus()
    Call SetVsGrindSelColor(mshInsure, True)
End Sub

Private Sub picDetail_Resize()
    Dim sngLeft As Single
    
    On Error Resume Next
    Call SetCondition
    tabTime.Width = picDetail.ScaleWidth - tabTime.Left - imgColSel.Width - 450
    With imgColSel
        .Left = picDetail.ScaleWidth - .Width - 200
    End With
    With chk��������С��
        .Left = imgColSel.Left - .Width - 200
        sngLeft = imgColSel.Left - chk����ʾ���ʵ���.Width - 200
        If .Visible Then sngLeft = .Left
    End With
    With chk����ʾ���ʵ���
        .Left = sngLeft - .Width - 200
        If .Visible Then sngLeft = .Left
    End With
    '46646
    With chkNotCheckFee
        .Top = chk����ʾ���ʵ���.Top
        .Left = sngLeft - .Width - 200
    End With
End Sub

Private Sub picLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshInsure.Width - X < 2000 Or mshDepost.Width + X < 2000 Then Exit Sub
        picLR.Left = picLR.Left + X
        lblDepost.Width = lblDepost.Width + X
        mshDepost.Width = mshDepost.Width + X
        lblInsure.Left = lblInsure.Left + X
        lblInsure.Width = lblInsure.Width - X
        mshInsure.Left = mshInsure.Left + X
        mshInsure.Width = mshInsure.Width - X
        
        Refresh
    End If
End Sub



Private Sub vsfFee_DblClick()
    Dim lngColTmp As Long, strNO As String, byt��¼���� As Byte, byt��¼״̬ As Byte
    If vsfFee.MouseRow = 0 Then Exit Sub
    
    With vsfFee
        If .Row > 0 Then
            lngColTmp = .ColIndex("���ݺ�")
            If lngColTmp <> -1 Then
                strNO = .TextMatrix(.Row, lngColTmp)
                
                lngColTmp = .ColIndex("��¼����")
                If lngColTmp <> -1 Then
                    byt��¼���� = Val(.TextMatrix(.Row, lngColTmp))
                End If
                
                If strNO <> "" And byt��¼���� <> 0 Then
                    lngColTmp = .ColIndex("��¼״̬")
                    If lngColTmp <> -1 Then
                        byt��¼״̬ = Val(.TextMatrix(.Row, lngColTmp))
                    End If
                    Call ShowBilling(strNO, byt��¼����, byt��¼״̬)
                End If
            End If
        End If
    End With
End Sub

Private Sub ShowBilling(ByVal strNO As String, ByVal byt��¼���� As Byte, ByVal byt��¼״̬ As Byte)
    Dim blnNOMoved As Boolean
    
    If Get������Դ(strNO) = 1 Then
        Call ZLShowChargeWindow(Me, 2, 1, 0, 0, 0, 0, False, 0, "", strNO)
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    gbytBilling = 0 '���ʲ���
    blnNOMoved = zlDatabase.NOMoved("סԺ���ü�¼", strNO, , 2, Caption)
    
    If BillisBatch(strNO) Then '��������
        frmBillings.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        frmBillings.mbytInState = 1
        frmBillings.mstrInNO = strNO
        frmBillings.mblnDelete = byt��¼״̬ = 2
        frmBillings.mblnNOMoved = blnNOMoved
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnDelete = byt��¼״̬ = 2
        frmSimpleBilling.mblnNOMoved = blnNOMoved
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        frmCharge.mbytNOType = byt��¼����
        frmCharge.mblnDelete = byt��¼״̬ = 2
        frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        frmCharge.mbytInState = 1
        frmCharge.mstrInNO = strNO
        frmCharge.mlngModule = mlngModul
        frmCharge.mblnNOMoved = blnNOMoved
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
End Sub

Private Function Get������Դ(ByVal strNO As String) As Byte
    '���ݵ��ݺ��жϷ��ü�¼��Դ
    '���أ�0-סԺ,1-����
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = "Select 1 From ������ü�¼ Where ��¼����=2 And NO=[1] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Function
    Get������Դ = 1
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
    
Private Sub ExecUnBilling()
    Dim strNO As String, strTime As String, blnNOMoved As Boolean, lng����ID As Long, strUnitIDs As String
    Dim blnBat As Boolean, intTmp As Integer, i As Long, intInsure As Integer, lngDelRow As Long
    Dim str����IDs As String, strInfo As String, strUser As String
    Dim strInsure As String, arrInsure As Variant, bytType As Byte, blnFlagPrint As Boolean
    Dim byt������Դ As Byte '0-סԺ,1-����
    
    If Not (mbytList = ListType.C0�����嵥 _
            Or mbytList = ListType.C1�ֿ�����ϸ _
            Or mbytList = ListType.C2����Ŀ��ϸ _
            Or mbytList = ListType.C3�������ϸ _
            Or mbytList = ListType.C4���������ϸ) Then Exit Sub
        
    If InStr(GetInsidePrivs(Enum_Inside_Program.pסԺ����), "���в���") = 0 Then
        If InStr("," & mstrUnitIDs & ",", "," & mlng����ID & ",") = 0 Then
            MsgBox "��û�����в�����Ȩ�ޣ����ܶ����������Ĳ������ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    With vsfFee
        i = .ColIndex("���ݺ�")
        If i = -1 Then Exit Sub
        strNO = .TextMatrix(.Row, i)
        If strNO = "" Then Exit Sub
        
        strTime = .TextMatrix(.Row, .ColIndex("�Ǽ�ʱ��"))
        bytType = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
        strUser = .TextMatrix(.Row, .ColIndex("����Ա"))
        lngDelRow = Val(.TextMatrix(.Row, .ColIndex("���")))
    End With
    byt������Դ = Get������Դ(strNO)
    
    'Ȩ���ж�
    If Not BillOperCheck(IIf(byt������Դ = 1, 4, 5), strUser, CDate(strTime), "����", strNO, , bytType) Then Exit Sub
        
    '�Ƿ���ת������ݱ���
    If zlDatabase.NOMoved(IIf(byt������Դ = 1, "������ü�¼", "סԺ���ü�¼"), strNO, , CStr(bytType), Caption) Then
        If Not ReturnMovedExes(strNO, bytType, Caption) Then Exit Sub
    End If
            
    '��Ŀ����Ȩ��
    If Not CheckDelPriv(strNO, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), strTime, bytType, , byt������Դ) Then Exit Sub
        
    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), strTime, bytType, byt������Դ)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '�Ƿ���ִ��
    If byt������Դ = 0 Then blnBat = BillisBatch(strNO)
    i = BillCanDelete(strNO, bytType, blnBat, strTime, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), blnFlagPrint, byt������Դ)
    If i <> 0 Then
        Select Case i
            Case 1 '�õ��ݲ�����
                MsgBox "ָ�������е����ݲ�����,������û������շ���Ŀ������Ȩ�ޣ�", vbInformation, gstrSysName
            Case 2 '�Ѿ�ȫ����ȫִ��
                MsgBox "ָ�������е������Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
            Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                MsgBox "ָ�������е�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '��Ժ���˲���Ȩ���ж�
    If Not BillCanBeOperate(strNO, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "����", _
        strTime, str����IDs, bytType, byt������Դ) Then Exit Sub
    
    '�Ƿ��Ѿ�����
    intTmp = HaveBilling(IIf(byt������Դ = 1, 1, 2), strNO, False, strTime, bytType)
    If intTmp <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True, bytType, byt������Դ)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , arrInsure(i)) Then
                        'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ����
                        If intTmp = 1 Then
                            MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbExclamation, gstrSysName
                        End If
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        Case 2
                            If intTmp = 1 Then
                                MsgBox "�ü��ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                                Exit Sub
                            Else
                                MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbExclamation, gstrSysName
                            End If
                    End Select
                End If
            Next
        End If
    End If
    
    intInsure = BillExistInsure(strNO, , , bytType, byt������Դ) '�ж��Ƿ���ҽ�����˼ǵ���,���ʱ�������ֻҪ��ҽ������
    'ҽ�����ʲ�����Ը�����¼��������
    If intInsure <> 0 Then
        If CheckNONegative(strNO, bytType, byt������Դ) Then
            MsgBox "�õ��ݴ��ڸ������ʼ�¼,���������ҽ�����ʲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
        
    '�Ƿ������������¼
    If CheckRecalcRecord(strNO, byt������Դ) Then
        MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
            "����ǰ�밴�ѱ�������ã������˽����������ʵ��ݵĴ����Żݽ�", vbInformation, Caption
    End If
     
    If byt������Դ = 1 Then
        Call ZLShowChargeWindow(Me, 2, 3, 0, 0, 0, 0, False, 0, "", strNO)
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
        
    gbytBilling = 0
    If blnBat Then '��������
        frmBillings.mbytUseType = 1
        frmBillings.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        frmBillings.mbytInState = 3
        frmBillings.mstrInNO = strNO
        frmBillings.mlngDelRow = lngDelRow
        frmBillings.mstrTime = strTime
        frmBillings.mstr����IDs = str����IDs
        frmBillings.mlngUnitID = 0
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO, bytType) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        frmSimpleBilling.mbytInState = 3
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mlngUnitID = 0
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        
        frmCharge.mbytUseType = 1
        frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        frmCharge.mbytInState = 3
        frmCharge.mstrInNO = strNO
        frmCharge.mlngDelRow = lngDelRow
        frmCharge.mbytNOType = bytType
        frmCharge.mstrTime = strTime
        frmCharge.mlngUnitID = 0
        frmCharge.mlngModule = mlngModul
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If

    If gblnOK Then Call RefreshAllData
End Sub
 
Private Sub ExecPrintDailyDetail()
    frmDailyListAsk.mlngModul = 1141    '��Ȼ��һ���嵥ģ��Ĳ���Ϊ׼
    frmDailyListAsk.mbytInFun = 1
    frmDailyListAsk.mlng����ID = mlng����ID
    frmDailyListAsk.Show vbModal, Me
    If frmDailyListAsk.mblnAskOk Then
        ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me, "����ID=" & mlng����ID, _
            "��ʼʱ��=" & Format(frmDailyListAsk.mdatBegin, "YYYY-MM-DD HH:MM:SS"), _
            "����ʱ��=" & Format(frmDailyListAsk.mdatEnd, "YYYY-MM-DD HH:MM:SS"), _
            "��ʾ�˷�=" & IIf(mvs.ReBalance, "1", "0"), _
            "��ʾ�����=" & IIf(mvs.ZeroFee, "1", "0"), _
            "���˲���=" & mlng����ID, _
            "��ҳID=" & frmDailyListAsk.mlngPageID, _
            "����ʱ��=" & IIf(mbytDateType, "����ʱ��", "�Ǽ�ʱ��"), 1
    End If
End Sub

Private Sub optDeptMode_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    If Visible Then Call LoadCardData(False, False, True)
End Sub

Private Sub optTypeMode_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    If Visible Then Call LoadCardData(False, False, True)
End Sub


Private Sub tabClass_Click()
    Dim strTmp As String
    Dim objControl As CommandBarControl
    Dim intPreSel As Integer
    
    If tabClass.SelectedItem.Index = mintPreCard Then Exit Sub
    
    mintPreCard = tabClass.SelectedItem.Index
    Set objControl = mcbsMain.ActiveMenuBar.FindControl(xtpControlButton, conMenu_View_DetailType * 10 + mbytList, True, True)
    If objControl Is Nothing Then Exit Sub
    strTmp = objControl.Caption
    If tabClass.SelectedItem.Index = conTabδ�� Then
        mblnPreBalance = mintInsure <> 0   '59073
        lblMoney.Caption = " δ��" & Left(strTmp, Len(strTmp) - 4) & "��"
    Else
        lblMoney.Caption = " ����" & Left(strTmp, Len(strTmp) - 4) & "��"
    End If
    intPreSel = mintPreTimeIndex
    If Not tabTime.SelectedItem Is Nothing And mintPreTimeIndex = 0 Then
        intPreSel = tabTime.SelectedItem.Index
    End If
    
    If LoadPatiTime Then
        '43494
        '�����:53136 �޸���:���˺�,�޸�ʱ��:2012-12-10 13:26:07
        If tabTime.SelectedItem Is Nothing Or tabClass.Tag = "Loaded" Then
            If tabTime.Tabs.Count >= intPreSel And intPreSel <> 0 Then
                tabTime.Tabs(intPreSel).Selected = True    '����tabTime_Click
            ElseIf tabTime.Tabs.Count <> 0 Then
                tabTime.Tabs((1)).Selected = True   '����tabTime_Click
            End If
        Else
            Call tabTime_Click
        End If
    End If
    tabClass.Tag = "Loaded"
End Sub

Private Sub tabTime_Click()
    '�����:53136 �޸���:���˺�,�޸�ʱ��:2012-12-10 11:55:54,��mintPreTimeIndex
    If tabTime.Tag = "1" Then Exit Sub
    If tabTime.SelectedItem.Index = mintPreTime Then Exit Sub
    mintPreTime = tabTime.SelectedItem.Index
    '��¼�ϴ����ѡ��
    mintPreTimeIndex = IIf(tabTime.Tabs.Count > 1, mintPreTime, mintPreTimeIndex)
    
    '��ʾ��ǰ��Ƭ����
    Call LoadCardData
End Sub

Private Sub vsfFee_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim objPopup As CommandBarPopup
        If Not Me.ActiveControl Is vsfFee Then vsfFee.SetFocus
        If vsfFee.MouseRow <> vsfFee.Row Then vsfFee.Row = vsfFee.MouseRow
    
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_EditPopup, True, False)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
Private Function zlCopyBill(ByVal int���� As Integer, ByVal frmMain As Object, ByVal lng����ID As Long, _
    ByVal lng����ID As Long, bln��Ժ As Boolean, ByVal bln���� As Boolean, _
    Optional strUnitIDs As String = "", Optional lng��ҳID As Long = 0, _
    Optional lng����ID As Long = 0, Optional ByVal bln�������۲��� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:int����- 0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS);9-���ò�ѯ����
    '����:���˺�
    '����:2013-02-17 10:45:31
    '����:54274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String, intIndex As Integer, strNO As String
    Dim lng����ID As Long, str���ת��ʱ�� As String
    '�ȿ��м���Ȩ��û��
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    If InStr(1, strPrivs, ";סԺ����;") = 0 Then Exit Function
    
    If InStr(GetInsidePrivs(Enum_Inside_Program.pסԺ����), "���в���") = 0 Then
        If strUnitIDs = "" Then
            '���»�ȡ����Ա�����ڲ���
            strUnitIDs = GetUserUnits
        End If
        If InStr("," & strUnitIDs & ",", "," & lng����ID & ",") = 0 Then
            MsgBox "��û�����в�����Ȩ�ޣ����ܶ����������Ĳ��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    lng����ID = 0
    If mbln���� Then
        If InStr(1, "012", int����) > 0 Then '    int���� = 0 - ҽ��վ����, 1 - ��ʿվ����, 2 - ҽ��վ����(PACS / LIS))
            lng����ID = lng����ID
        End If
        '���Ѽ���Ƿ񳬹�ʱ��
        If zlCheckPatiFeeRenewValied(lng����ID, lng��ҳID, lng����ID, lng����ID, str���ת��ʱ��) = False Then Exit Function
    Else
        lng����ID = lng����ID
    End If
    
    With vsfFee
        intIndex = .ColIndex("���ݺ�")
        If intIndex = -1 Then Exit Function
        strNO = .TextMatrix(.Row, intIndex)
        If strNO = "" Then Exit Function
    End With
    '�ಡ�˵���,���ܿ���
    If BillisBatch(strNO) Then
        MsgBox "���ݡ�" & strNO & "���Ǽ��ʱ�,�������Ƹü��ʵ���!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    '��Ժ���˼���Ȩ��
    If bln��Ժ Then
        If bln���� And InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "��Ժ����ǿ�Ƽ���") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf Not bln���� And InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "��Ժδ��ǿ�Ƽ���") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�������۲������������ģʽ
    If bln�������۲��� Then
        If Not (gbln�������� And InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), ";�������ۼ���;") > 0) Then
            MsgBox "��û��Ȩ�޶��������۲��˽��м��ʲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        zlCopyBill = ZLShowChargeWindow(Me, 2, 11, lng����ID, lng��ҳID, _
            lng����ID, lng����ID, False, 0, "", strNO)
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    gblnOK = False
    gbytBilling = 0
    frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
    frmCharge.mbytUseType = 1
    frmCharge.mbytInState = 0
    frmCharge.mblnCopyBill = True
    frmCharge.mstrInNO = strNO
    frmCharge.mlngDeptID = lng����ID
    frmCharge.mlngUnitID = lng����ID
    frmCharge.mlngModule = mlngModul
    frmCharge.mlng����ID = lng����ID
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If Not gblnOK Then Exit Function
    zlCopyBill = True
End Function

Private Function LoadPages(ByVal intPage As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���صڼ�ҳ
    '���:intPage-ҳ��
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-03 10:26:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, cllTemp As Collection
    Dim strKey As String
    On Error GoTo errHandle
    
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
    
    Set cllTemp = mcllBalaceNums(intPage)
    For i = 1 To cllTemp.Count
        If i = 1 Then strKey = "_" & cllTemp(i)(1)
        tabClass.Tabs.Add , "_" & cllTemp(i)(1), cllTemp(i)(2)
        tabClass.Tabs(tabClass.Tabs.Count).Tag = Val(cllTemp(i)(0)) '��¼����ID,�ӿ��ٶ�
        tabClass.Tabs(tabClass.Tabs.Count).ToolTipText = cllTemp(i)(3)
    Next
    If cllTemp.Count > 0 Then tabClass.Tabs(strKey).Selected = True
    If picNum.Enabled And picNum.Visible Then
        picNum.Left = IIf(995 + 1680 * (tabClass.Tabs.Count - 1) + 120 < Me.ScaleWidth - picNum.Width, 995 + 1680 * (tabClass.Tabs.Count - 1) + 120, Me.ScaleWidth - picNum.Width - 30)
        tabClass.Width = Me.ScaleWidth - picNum.Width - 30
    End If
    LoadPages = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFormOperation() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ѡ�񣬴���ж��ǰ����
    '����:
    '     �ϴδ�����������ַ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strKey As String

    On Error GoTo errHandle
    
    If mlng����ID = 0 Then Exit Function

    'zlGetFormOperation�ַ�����ʽ:
    '����ID|��������(tabClass.SelectedItem.Key)|�嵥��ѯ�������|סԺ����(tabTime.SelectedItem.Key)|cboNum.Text|chkAdivce,chkAdivce.Value| _
    chk������������,chk������������.Value|cbo����,cbo����.cbo����.Text,dtpBegin.Value,dtpEnd.Value| _
    chk����ʾ���ʵ���,chk����ʾ���ʵ���.Value|chk��������С��,chk��������С��.Value| _
    cboFeeType,cboFeeType.Text|cboDept,cboDept.Text|cboBaby,cboBaby.Text|optDeptMode,optDeptMode(0).Value| _
    optTypeMode,optTypeMode(0).Value|mshDepost,���ݺ�,���㷽ʽ|mshInsure,���㷽ʽ|vsfFee,��������,���ݺ�,��Ŀ����,��¼״̬...|cboNum,cboNum.Text

    strKey = mlng����ID & "|" & tabClass.SelectedItem.Key & "|" & mbytList & "|" & tabTime.SelectedItem.Key & "|" & cboNum.Text
    strKey = strKey & "|chkAdivce," & IIf(chkAdivce.Value = 1, 1, 0) & "|chkNotCheckFee," & IIf(chkNotCheckFee.Value = 1, 1, 0)
    strKey = strKey & "|cbo����," & cbo����.Text & "," & Format(dtpBegin.Value, "yyyy-mm-dd hh:MM:ss") & "," & Format(dtpEnd.Value, "yyyy-mm-dd hh:MM:ss")
    strKey = strKey & "|chk����ʾ���ʵ���," & IIf(chk����ʾ���ʵ���.Value = 1, 1, 0)
    strKey = strKey & "|chk��������С��," & IIf(chk��������С��.Value = 1, 1, 0)
    strKey = strKey & "|cboFeeType," & cboFeeType.Text
    strKey = strKey & "|cboDept," & cboDept.Text
    strKey = strKey & "|cboBaby," & cboBaby.Text
    strKey = strKey & "|optDeptMode," & IIf(optDeptMode(0).Value, 0, 1)
    strKey = strKey & "|optTypeMode," & IIf(optTypeMode(0).Value, 0, 1)

    With mshDepost
        If .Rows > 1 And .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then
            strKey = strKey & "|mshDepost," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                     .TextMatrix(.RowSel, .ColIndex("���㷽ʽ"))
        End If
    End With
    
    With mshInsure
        If mintInsure <> 0 Then
            If .Rows > 1 And .TextMatrix(1, .ColIndex("���㷽ʽ")) <> "" Then
                strKey = strKey & "|mshInsure," & .TextMatrix(.RowSel, .ColIndex("���㷽ʽ"))
            End If
        End If
    End With
    
    With vsfFee
        Select Case mbytList
            Case ListType.C0�����嵥
                If .Rows > 1 And .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��Ŀ����")) & "," & .TextMatrix(.RowSel, .ColIndex("��¼״̬"))
                End If
            Case ListType.C1�ֿ�����ϸ
                If .Rows > 1 And .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��������")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��Ŀ����")) & "," & .TextMatrix(.RowSel, .ColIndex("��¼״̬"))
                End If
                
            Case ListType.C2����Ŀ��ϸ
                If .Rows > 1 And .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("��Ŀ")) <> "" Then   '������
                        If .TextMatrix(.RowSel, .ColIndex("��Ŀ")) = "�ܼ�" Then
                            strKey = strKey & "|vsfFee,,,,,�ܼ�"
                        Else
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("��������")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("���ݺ�")) & "," & _
                            .TextMatrix(.RowSel - 1, .ColIndex("��Ŀ����")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("��¼״̬")) & ",����"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                        .TextMatrix(.RowSel, .ColIndex("��Ŀ����")) & "," & .TextMatrix(.RowSel, .ColIndex("��¼״̬"))
                    End If
                End If
                
            Case ListType.C3�������ϸ
                If .Rows > 1 And .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��Ŀ")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��Ŀ����")) & "," & .TextMatrix(.RowSel, .ColIndex("��¼״̬"))
                End If
                
            Case ListType.C4���������ϸ
                If .Rows > 1 And .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("������Ŀ")) & "," & .TextMatrix(.RowSel, .ColIndex("�շ���Ŀ")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("��Ŀ����")) & "," & .TextMatrix(.RowSel, .ColIndex("��¼״̬"))
                End If
                
            Case ListType.C5����Ŀ����
                If .Rows > 1 And .TextMatrix(1, .ColIndex("�շ����")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("�շ����")) & "," & .TextMatrix(.RowSel, .ColIndex("��Ŀ")) & "," & _
                    .TextMatrix(.RowSel, .ColIndex("���"))
                End If
            Case ListType.C6��������
                If .Rows > 1 And .TextMatrix(1, .ColIndex("��Ŀ")) <> "" Then
                    strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��Ŀ"))
                End If
            Case ListType.C7���·������
                If .Rows > 1 And .TextMatrix(1, .ColIndex("�ڼ�")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("��Ŀ")) = "" Then   '������
                        If .TextMatrix(.RowSel, .ColIndex("�ڼ�")) <> "�ܼ�" Then
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("�ڼ�")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("��Ŀ")) & ",����"
                        Else
                            strKey = strKey & "|vsfFee,,,�ܼ�"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("�ڼ�")) & "," & .TextMatrix(.RowSel, .ColIndex("��Ŀ")) & ","
                    End If
                End If
            Case ListType.C8���յ��ݻ���
                If .Rows > 1 And .TextMatrix(1, .ColIndex("��������")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("������Ŀ")) = "" Then   '������
                        If .TextMatrix(.RowSel, .ColIndex("��������")) <> "�ܼ�" Then
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("��������")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("���ݺ�")) & "," & _
                            .TextMatrix(.RowSel - 1, .ColIndex("������Ŀ")) & ",����"
                        Else
                            strKey = strKey & "|vsfFee,,,,�ܼ�"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("���ݺ�")) & "," & _
                        .TextMatrix(.RowSel, .ColIndex("������Ŀ")) & ","
                    End If
                End If
            Case ListType.C9���շ�Ŀ����
                If .Rows > 1 And .TextMatrix(1, .ColIndex("��������")) <> "" Then
                    If .TextMatrix(.RowSel, .ColIndex("������Ŀ")) = "" Then   '������
                        If .TextMatrix(.RowSel, .ColIndex("��������")) <> "�ܼ�" Then
                            strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel - 1, .ColIndex("��������")) & "," & .TextMatrix(.RowSel - 1, .ColIndex("������Ŀ")) & ",����"
                        Else
                            strKey = strKey & "|vsfFee,,,�ܼ�"
                        End If
                    Else
                        strKey = strKey & "|vsfFee," & .TextMatrix(.RowSel, .ColIndex("��������")) & "," & .TextMatrix(.RowSel, .ColIndex("������Ŀ")) & ","
                    End If
                End If
        End Select
    End With
    strKey = strKey & "|cboNum," & cboNum.Text
    
    zlGetFormOperation = strKey
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlRestoreFormOperation(ByVal strValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ��������ѡ�񣬴���ˢ��ǰ����
    '���:
    '     strValue-�ϴδ�����������ַ���
    '����:
    '     True-����ָ��ɹ�;False-����ָ�ʧ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Integer
    On Error GoTo errHandle
    
    If strValue = "" Then Exit Function
    mstrRestoreFeeCons = strValue
    varData = Split(strValue, "|")
    mbytList = varData(2)
    zlRestoreFormOperation = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlRestoreFeeControls(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ����ò�ѯҳǩ�����ؼ��Ĳ���ѡ��
    '���:
    '     lng����ID-����ID
    '����:
    '     True-�ָ��ɹ�;False-�ָ�ʧ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Relng����ID As Long
    Dim varData As Variant, i As Integer, j As Integer
    On Error GoTo errHandle
    If mstrRestoreFeeCons = "" Then Exit Function
    varData = Split(mstrRestoreFeeCons, "|")
    Relng����ID = Val(varData(0))
    If Relng����ID <> lng����ID Then Exit Function
    
    For i = 1 To tabTime.Tabs.Count
        If tabTime.Tabs(i).Key = varData(3) Then tabTime.Tabs(i).Selected = True
    Next
    mblnNotClick = True
    For i = 5 To UBound(varData)
        Select Case Split(varData(i), ",")(0)
            Case "cbo����"
                cbo����.ListIndex = -1
                    For j = 0 To cbo����.ListCount - 1
                        If cbo����.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                            cbo����.ListIndex = j
                            Exit For
                        End If
                        cbo����.ListIndex = 0
                    Next
                If cbo����.ListIndex = 5 Then
                    dtpBegin.Value = Format(Split(varData(i), ",")(2), "yyyy-mm-dd hh:MM:ss")
                    dtpEnd.Value = Format(Split(varData(i), ",")(3), "yyyy-mm-dd hh:MM:ss")
                End If
                Call SetDateVisible
            Case "chkAdivce"
                chkAdivce.Value = Val(Split(varData(i), ",")(1))
                
            Case "chkNotCheckFee"
                chkNotCheckFee.Value = Val(Split(varData(i), ",")(1))
                
            Case "chk����ʾ���ʵ���"
                If chk����ʾ���ʵ���.Visible And chk����ʾ���ʵ���.Enabled Then
                    chk����ʾ���ʵ���.Value = Val(Split(varData(i), ",")(1))
                End If
                
            Case "chk��������С��"
                If chk��������С��.Visible And chk��������С��.Enabled Then
                    chk��������С��.Value = Val(Split(varData(i), ",")(1))
                End If
                
            Case "cboFeeType"
                cboFeeType.ListIndex = -1
                For j = 0 To cboFeeType.ListCount - 1
                    If cboFeeType.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                        cboFeeType.ListIndex = j
                        Exit For
                    End If
                    cboFeeType.ListIndex = 0
                Next
                
            Case "cboDept"
                cboDept.ListIndex = -1
                For j = 0 To cboDept.ListCount - 1
                    If cboDept.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                        cboDept.ListIndex = j
                        Exit For
                    End If
                    cboDept.ListIndex = 0
                Next
                
            Case "cboBaby"
                cboBaby.ListIndex = -1
                For j = 0 To cboBaby.ListCount - 1
                    If cboBaby.List(j) = Nvl(Split(varData(i), ",")(1)) Then
                        cboBaby.ListIndex = j
                        Exit For
                    End If
                    cboBaby.ListIndex = 0
                Next
                
            Case "optDeptMode"
                If optDeptMode(0).Value And optDeptMode(0).Enabled Then
                    optDeptMode(Val(Split(varData(i), ",")(1))).Value = True
                End If
                
            Case "optTypeMode"
                If optTypeMode(0).Value And optTypeMode(0).Enabled Then
                    optTypeMode(Val(Split(varData(i), ",")(1))).Value = True
                End If

        End Select
    Next
    mblnNotClick = False
    zlRestoreFeeControls = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlRestorePosition(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ����ò�ѯҳǩ�б�ؼ���ѡ����
    '���:
    '     lng����ID-����ID
    '����:
    '     True-�ָ��ɹ�;False-�ָ�ʧ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Relng����ID As Long
    Dim varData As Variant, i As Integer, j As Integer
    On Error GoTo errHandle
    If mstrRestoreFeeCons = "" Then Exit Function
    varData = Split(mstrRestoreFeeCons, "|")
    Relng����ID = Val(varData(0))
    If Relng����ID <> lng����ID Then Exit Function
    
    For i = 5 To UBound(varData)
        Select Case Split(varData(i), ",")(0)
            Case "mshDepost"
                With mshDepost
                    For j = 1 To .Rows - 1
                        If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                           Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("���㷽ʽ")) Then
                            .Row = j
                            Exit For
                        End If
                    Next
                End With
            
            Case "mshInsure"
                With mshInsure
                    For j = 1 To .Rows - 1
                        If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("���㷽ʽ")) Then
                            .Row = j
                            Exit For
                        End If
                    Next
                End With
                        
            Case "vsfFee"
                With vsfFee
                    Select Case mbytList
                        Case ListType.C0�����嵥
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��������")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("��Ŀ����")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("��¼״̬")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                            
                        Case ListType.C1�ֿ�����ϸ
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��������")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("��������")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("��Ŀ����")) And _
                                   Nvl(Split(varData(i), ",")(5)) = .TextMatrix(j, .ColIndex("��¼״̬")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C2����Ŀ��ϸ
                            If Nvl(Split(varData(i), ",")(5)) = "�ܼ�" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��������")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                                       Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("��Ŀ����")) And _
                                       Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("��¼״̬")) Then
                                        If Nvl(Split(varData(i), ",")(5)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Case ListType.C3�������ϸ
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��Ŀ")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("��������")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("��Ŀ����")) And _
                                   Nvl(Split(varData(i), ",")(5)) = .TextMatrix(j, .ColIndex("��¼״̬")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C4���������ϸ
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("������Ŀ")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("�շ���Ŀ")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("��������")) And _
                                   Nvl(Split(varData(i), ",")(4)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                                   Nvl(Split(varData(i), ",")(5)) = .TextMatrix(j, .ColIndex("��Ŀ����")) And _
                                   Nvl(Split(varData(i), ",")(6)) = .TextMatrix(j, .ColIndex("��¼״̬")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C5����Ŀ����
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("�շ����")) And _
                                   Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("��Ŀ")) And _
                                   Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("���")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C6��������
                            For j = 1 To .Rows - 1
                                If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��Ŀ")) Then
                                    .Row = j: .ShowCell .Row, 0
                                    Exit For
                                End If
                            Next
                        Case ListType.C7���·������
                            If Nvl(Split(varData(i), ",")(3)) = "�ܼ�" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("�ڼ�")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("��Ŀ")) Then
                                        If Nvl(Split(varData(i), ",")(3)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Case ListType.C8���յ��ݻ���
                            If Nvl(Split(varData(i), ",")(4)) = "�ܼ�" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��������")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("���ݺ�")) And _
                                       Nvl(Split(varData(i), ",")(3)) = .TextMatrix(j, .ColIndex("������Ŀ")) Then
                                        If Nvl(Split(varData(i), ",")(4)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        Case ListType.C9���շ�Ŀ����
                            If Nvl(Split(varData(i), ",")(3)) = "�ܼ�" Then
                                .Row = .Rows - 1: .ShowCell .Row, 0
                            Else
                                For j = 1 To .Rows - 1
                                    If Nvl(Split(varData(i), ",")(1)) = .TextMatrix(j, .ColIndex("��������")) And _
                                       Nvl(Split(varData(i), ",")(2)) = .TextMatrix(j, .ColIndex("������Ŀ")) Then
                                        If Nvl(Split(varData(i), ",")(3)) = "" Then
                                            .Row = j: .ShowCell .Row, 0
                                        Else
                                            .Row = j + 1: .ShowCell .Row, 0
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                    End Select
                End With
        End Select
    Next
    mstrRestoreFeeCons = ""
    zlRestorePosition = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


