VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmClinicSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "������Ŀѡ����"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   Icon            =   "frmClinicSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9120
   Begin VB.CheckBox chkShowCause 
      Caption         =   "��ʾδƥ��ɹ�����Ŀ"
      Height          =   195
      Left            =   4080
      TabIndex        =   23
      Top             =   5658
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame fraInfo 
      Height          =   435
      Left            =   30
      TabIndex        =   14
      Top             =   -75
      Width           =   9075
      Begin VB.CheckBox chkSub 
         Caption         =   "�����¼���Ŀ(&T)"
         Height          =   195
         Left            =   7380
         TabIndex        =   10
         Top             =   165
         Width           =   1650
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   180
         Left            =   225
         TabIndex        =   15
         Top             =   165
         Width           =   270
      End
   End
   Begin VB.Frame fraStat 
      Height          =   5160
      Left            =   15
      TabIndex        =   17
      Top             =   285
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdSelClear 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   1425
         TabIndex        =   7
         ToolTipText     =   "Ctrl+R"
         Top             =   3705
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   1425
         TabIndex        =   6
         ToolTipText     =   "Ctrl+A"
         Top             =   3345
         Width           =   1100
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "ȫ����ʾ"
         Height          =   195
         Left            =   1050
         TabIndex        =   4
         Top             =   1965
         Width           =   1020
      End
      Begin VB.CommandButton cmdStat 
         Caption         =   "ͳ��(&S)"
         Height          =   350
         Left            =   1425
         TabIndex        =   5
         Top             =   2835
         Width           =   1100
      End
      Begin VB.TextBox txtCount 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "100"
         Top             =   1620
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   106233859
         CurrentDate     =   38434
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmClinicSelect.frx":058A
         Left            =   1050
         List            =   "frmClinicSelect.frx":05A6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   825
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "���ͳ�Ƶ�ʱ�䷶Χ�ϳ����ٶȿ��ܻ�����������ĵȴ���"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   165
         TabIndex        =   22
         Top             =   2340
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   2100
         TabIndex        =   21
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ʾ��ǰ"
         Height          =   180
         Left            =   285
         TabIndex        =   20
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��ʱ��"
         Height          =   180
         Left            =   285
         TabIndex        =   19
         Top             =   885
         Width           =   720
      End
      Begin VB.Label lblStatTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ�ͳ��""XXXXXX""������õ�������Ŀ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   18
         Top             =   270
         Width           =   2400
      End
   End
   Begin MSComctlLib.ImageList imgOften 
      Left            =   1110
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":060A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":0D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":1AF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOften 
      Height          =   450
      Left            =   495
      TabIndex        =   16
      Top             =   5505
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   794
      ButtonWidth     =   1561
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgOften"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "Often"
            Description     =   "����"
            Object.ToolTipText     =   "��ʾ������Ŀ(F2)"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "ͳ��"
            Key             =   "Stat"
            Description     =   "ͳ��"
            Object.ToolTipText     =   "ͳ�Ƴ�����Ŀ"
            Object.Tag             =   "ͳ��"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "New"
            Description     =   "����"
            Object.ToolTipText     =   "���볣����Ŀ(F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "�Ƴ�"
            Key             =   "Del"
            Description     =   "�Ƴ�"
            Object.ToolTipText     =   "�Ƴ�������Ŀ(Del)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4665
      Left            =   2790
      TabIndex        =   8
      Top             =   375
      Width           =   6300
      _cx             =   11112
      _cy             =   8229
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicSelect.frx":21F2
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   930
         Top             =   900
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":227F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":2759
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   13
      Top             =   555
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   12
      Top             =   5580
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      TabIndex        =   11
      Top             =   5580
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   600
      Left            =   2835
      TabIndex        =   9
      Top             =   4815
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   1058
      TabWidthStyle   =   2
      TabFixedWidth   =   1623
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      Placement       =   1
      ImageList       =   "img16"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫ��(0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�г�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�в�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2235
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
            Picture         =   "frmClinicSelect.frx":2C33
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":31CD
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":3767
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":3D01
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":429B
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":4835
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   4995
      Left            =   15
      TabIndex        =   0
      Top             =   390
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   8811
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "����ҩ�����"
      Height          =   180
      Left            =   2760
      TabIndex        =   24
      Top             =   5595
      Width           =   1080
   End
   Begin VB.Shape Shp 
      Height          =   405
      Left            =   4800
      Top             =   5550
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   10000
      Y1              =   5445
      Y2              =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frmClinicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mint��Χ As Integer
Private mlng�������� As Long
Private mlng���˿���id As Long
Private mblnOK As Boolean
Private mrsItem As ADODB.Recordset
Private mint��Ч As Integer
Private mstr�Ա� As String
Private mstr���� As String
Private mobjTXT As Object
Private mlng����ID As Long
Private mlng��λ����ID As Long
Private mint���� As Integer
Private mintType As Integer  '��mlngҩ��ID>0ʱ �ñ�����ֵ��Ч: 0-��ȡָ��Ʒ��ҩƷ�����й��;1-��ȡָ��Ʒ�����ĵ����й��

Private mstr���Ʒ��� As String
Private mstr�������� As String
Private mstrִ�з��� As String

Private mstrSaveTag As String
Private mstrPreNode As String
Private mblnClick As Boolean

Private mbln�۸� As Boolean
Private mbln���� As Boolean
Private mint���� As Integer
Private mstrLike As String

Private mstr������ҩ�� As String
Private mstr���ó�ҩ�� As String
Private mstr������ҩ�� As String
Private mstr���ϲ��� As String

Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng���ϲ��� As Long
Private mlngҩ��ID As Long '��ȡָ��Ʒ��ҩƷ�����й��;��ȡָ��Ʒ�����ĵ����й��

Private mstr����ID As String '�������õĿ���ID
Private mstrPrivs As String
Private mbytƥ�� As Byte '��ƥ���ԭ��1��ʾ��Ч��ƥ��
Private mbytSize As Byte
Private mbln��ʾ��� As Boolean
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�
Private mlngҽ������ID As Long

Public Function ShowSelect(frmParent As Object, ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng����id As Long, _
    ByVal int��Ч As Integer, ByVal str�Ա� As String, Optional ByVal str���� As String, _
    Optional objTxt As Object, Optional ByVal int��Χ As Integer = 2, _
    Optional ByVal lng����ID As Long, Optional ByVal int���� As Integer, Optional ByVal lng�������� As Integer, Optional ByVal lngҩ��ID As Long, _
    Optional ByVal strʹ�ÿ��� As String, Optional ByRef bytƥ�� As Byte, Optional str���Ʒ��� As String, _
    Optional ByVal str�������� As String, Optional ByVal strִ�з��� As String, Optional ByVal lng��λ����ID As Long, _
    Optional ByVal strҩƷ�۸�ȼ� As String, Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ��Ŀ�۸�ȼ� As String, _
    Optional ByVal intType As Integer, Optional ByVal lngҽ������ID As Long) As ADODB.Recordset
'���ܣ���ʾ������Ŀѡ����
'������int����=(-1)-���׷����༭,0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      lng����ID/lng����ID=���˵Ĳ���/����ID
'      int��Ч=ҽ����Ч
'      str�Ա�=�����Ա�
'      str����=����ƥ�������,���û����Ϊѡ������ʽ,����Ϊ�б�ʽ
'      objTXT=�����б�λ�������
'      blnCancel(O):�Ƿ�ȡ��
'      int��Χ=1-����,2-סԺ,3-�����סԺ
'      lng����ID=ѡ����ʱ(str����="")����������࿪ʼ��ʾ
'      lngҩ��ID=��ȡָ��Ʒ��ҩƷ\���ĵ����й�񣬴���ҩƷ\���ĵ�������ĿID
'      bytƥ�� ���Σ���ƥ��ԭ�� =1��Ч =2����
'      str���Ʒ��� =��-1�����׷����༭��·����Ŀ��������ʱ����
'      str��������=��-1�����׷����༭��·����Ŀ��������ʱ����
'      strִ�з���=��-1�����׷����༭��·����Ŀ��������ʱ����
'      lng��λ����ID=��λ���÷�������
'      intType-�ò����������֡�lngҩ��ID����ҩƷ��������ĿID�������ĵ�������ĿID��=0����lngҩ��ID��ΪҩƷ��������ĿID;=1 ��lngҩ��ID��Ϊ���ĵ�������ĿID(�ٴ�·������ʱ)
'      lngҽ������ID ҽ������վ����ʱ��ǰ�������ID
'���أ����û������,��ȡ��,�򷵻�Nothing������Ϊһ������������Ŀ���ݵļ�¼
    mint���� = int����
    mint��Χ = int��Χ
    mint��Ч = int��Ч
    mstr�Ա� = str�Ա�
    mstr���� = str����
    mlngҩ��ID = lngҩ��ID
    mlng���˿���id = lng����id
    mlngҽ������ID = lngҽ������ID
    If mlngҩ��ID <> 0 Then mstr���� = ""
    
    Set mobjTXT = objTxt
    mlng����ID = lng����ID
    mlng��λ����ID = lng��λ����ID
    mint���� = int����
    mlng�������� = lng��������
    mstrҩƷ�۸�ȼ� = strҩƷ�۸�ȼ�
    mstr���ļ۸�ȼ� = str���ļ۸�ȼ�
    mstr��ͨ��Ŀ�۸�ȼ� = str��ͨ��Ŀ�۸�ȼ�
    
    mstrSaveTag = mint��Χ & IIF(mstr���� <> "", 1, 0) & IIF(gblnҩƷ�������ҽ�� Or mint��Ч = 1, 1, 0)
    
    '�������ÿ���
    If mint���� = -1 Then
        '���׷����༭���ӿڲ����룬ȡ����Ա�������п���
        If strʹ�ÿ��� = "" Then
            mstr����ID = GetUser����IDs
            If mstr����ID <> "" Then mstr����ID = "," & mstr����ID & ","
        Else
            mstr����ID = "," & strʹ�ÿ��� & ","
        End If
    Else
        'ҽ������վ�����ݲ��˿��ң�����������Ŀ��ʹ�ÿ���
        'ҽ������վ��
        '    סԺ�����ݲ��˿��ң�����������Ŀ��ʹ�ÿ���
        '    ������ݲ��˿��ң�����������Ŀ��ʹ�ÿ���
        mstr����ID = "," & lng����id & ","
    End If
    mbytƥ�� = 0
    mstr���Ʒ��� = str���Ʒ���
    mstr�������� = str��������
    mstrִ�з��� = strִ�з���
    mintType = intType
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        bytƥ�� = mbytƥ��
        Set ShowSelect = mrsItem
    Else
        bytƥ�� = 0
        Set ShowSelect = Nothing
    End If
End Function

Private Sub cboDate_Click()
    Dim curDate As Date
    
    If cboDate.ListIndex = cboDate.ListCount - 1 Then
        dtpDate.Enabled = True
        dtpDate.SetFocus
    Else
        dtpDate.Enabled = False
        curDate = zlDatabase.Currentdate
        Select Case cboDate.ListIndex
            Case 0 'һ��
                dtpDate.value = DateAdd("ww", -1, curDate)
            Case 1 '����(15��)
                dtpDate.value = DateAdd("d", -15, curDate)
            Case 2 'һ��
                dtpDate.value = DateAdd("m", -1, curDate)
            Case 3 '����
                dtpDate.value = DateAdd("m", -2, curDate)
            Case 4 '����
                dtpDate.value = DateAdd("m", -3, curDate)
            Case 5 '����
                dtpDate.value = DateAdd("m", -6, curDate)
            Case 6 'һ��
                dtpDate.value = DateAdd("yyyy", -1, curDate)
        End Select
    End If
End Sub

Private Sub chkAll_Click()
    txtCount.Enabled = chkAll.value = 0
End Sub

Private Sub chkShowCause_Click()
    If chkShowCause.value = 1 Then tbrOften.Buttons(1).value = tbrUnpressed
    Call FillList
    Call SetFormSize
    If chkShowCause.value = 1 Then
        cmdOK.Enabled = False
        If mrsItem.RecordCount = 1 Then
            If InStr(mrsItem!δƥ��ԭ��, "��Ŀ���벻ƥ��") > 0 Or InStr(mrsItem!δƥ��ԭ��, "������Ŀ��ִ��Ƶ��Ϊ") > 0 And mint��Χ <> 1 Then
                cmdOK.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkSub_Click()
    If Not Visible Then Exit Sub
    vsItem.SetFocus
    Call FillList(True)
End Sub

Private Sub cmdCancel_Click()
    Set mrsItem = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If chkShowCause.value = 1 Then
        If mrsItem.RecordCount = 1 Then
            If InStr(mrsItem!δƥ��ԭ��, "��Ŀ���벻ƥ��") > 0 Then
                mbytƥ�� = 2
            ElseIf InStr(mrsItem!δƥ��ԭ��, "������Ŀ��ִ��Ƶ��Ϊ") > 0 Then
                mbytƥ�� = 1
            End If
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdStat_Click()
    If chkAll.value = 0 Then
        If Val(txtCount.Text) <= 0 Then
            MsgBox "��������ȷ����ʾ������", vbInformation, gstrSysName
            txtCount.SetFocus: Exit Sub
        End If
    End If
    
    Call FillStat(True)
    vsItem.SetFocus
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        If Val(vsItem.TextMatrix(i, 1)) <> 0 Then
            vsItem.TextMatrix(i, 2) = 1
        End If
    Next
End Sub

Private Sub cmdSelClear_Click()
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        vsItem.TextMatrix(i, 2) = 0
    Next
End Sub

Private Sub Form_Activate()
    If Not tvw_s.Visible And vsItem.Visible Then vsItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf Shift = vbAltMask Then
        If Between(KeyCode, vbKey0, vbKey9) Then
            lngIdx = KeyCode - vbKey0 + 1
        End If
        If tabClass.SelectedItem.Index <> lngIdx And Between(lngIdx, 1, tabClass.Tabs.Count) Then
            tabClass.Tabs(lngIdx).Selected = True
        End If
    ElseIf Shift = vbCtrlMask Then
        If KeyCode = vbKeyA Then
            If fraStat.Visible Then cmdSelALL_Click
        ElseIf KeyCode = vbKeyR Then
            If fraStat.Visible Then cmdSelClear_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If tbrOften.Buttons("Often").Visible And tbrOften.Buttons("Often").Enabled Then
            If tbrOften.Buttons("Often").value = tbrPressed Then
                tbrOften.Buttons("Often").value = tbrUnpressed
            Else
                tbrOften.Buttons("Often").value = tbrPressed
            End If
            Call tbrOften_ButtonClick(tbrOften.Buttons("Often"))
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If tbrOften.Buttons("New").Visible And tbrOften.Buttons("New").Enabled Then
            Call tbrOften_ButtonClick(tbrOften.Buttons("New"))
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If tbrOften.Buttons("Del").Visible And tbrOften.Buttons("Del").Enabled Then
            Call tbrOften_ButtonClick(tbrOften.Buttons("Del"))
        End If
    End If
End Sub

Private Function ExistOftenItem(Optional ByVal lng����ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng����ID <> 0 Then
        strSQL = "Select ID From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Start With ID=[2] Connect by Prior ID=�ϼ�ID"
        strSQL = "Select 1 From ���Ƹ�����Ŀ A,������ĿĿ¼ B" & _
            " Where A.������ĿID=B.ID And B.����ID IN(" & strSQL & ") And A.��ԱID=[1]" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And rownum<2"
    Else
        strSQL = "Select 1 From ���Ƹ�����Ŀ Where ��ԱID=[1] And rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, UserInfo.ID, lng����ID)
    ExistOftenItem = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim blnDo As Boolean
    Dim str������ĿIDs As String, str�շ�ϸĿIDs As String
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    Call SetFontSize(mbytSize)
    mblnOK = False
    mblnClick = True
    mstrPreNode = ""
    Set mrsItem = Nothing
    mstrPrivs = GetInsidePrivs(IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�))
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    mlng��ҩ�� = Val(zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ��ҩ��", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id))
    mlng��ҩ�� = Val(zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ��ҩ��", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id))
    mlng��ҩ�� = Val(zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ��ҩ��", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id))
    mlng���ϲ��� = Val(zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ���ϲ���", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id))
    
    mstr������ҩ�� = zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "������ҩ��", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id)
    mstr���ó�ҩ�� = zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "���ó�ҩ��", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id)
    mstr������ҩ�� = zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "������ҩ��", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id)
    mstr���ϲ��� = zlDatabase.GetPara(Decode(mint��Χ, 1, "����", 2, "סԺ", "") & "���÷��ϲ���", glngSys, Decode(mint��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0), , , , , mlng���˿���id)
    If mint���� = 0 Then
        '��������
        mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")
    ElseIf mint���� = 2 Then
        mbytSize = zlDatabase.GetPara("����", glngSys, pҽ������վ, "0")
    End If
    'ѡ�����е�����
    mbln���� = True '�Ƿ���ʾ���룺��δ�Ӳ������̶���ʾ
    If mint���� <> -1 Then
        mbln�۸� = True '�Ƿ���ʾ�۸���δ�Ӳ������̶���ʾ
        mbln��ʾ��� = Val(zlDatabase.GetPara("��ʾҩƷ���", glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�))) = 1 '�Ƿ���ʾҩƷ���
    Else
        mbln�۸� = False
    End If
    
    lblStatTitle.Caption = Replace(lblStatTitle.Caption, "XXXXXX", UserInfo.����)
    cboDate.ListIndex = 0
    If mlngҩ��ID = 0 Then
        Call SetOftenToolBar(mstr���� = "")
    Else
        tbrOften.Visible = False
    End If
    
    If mstr���� = "" Then
        tvw_s.Visible = mlngҩ��ID = 0
        chkSub.Visible = mlngҩ��ID = 0
        
        '��ȡ���ʧ��,����ʾ,��ȡ���˳�
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '�����,��ʾ,��ȡ���˳�
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "û����������������,���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
        '����и�����Ŀ,ȱʡת��������Ŀ
        If ExistOftenItem(mlng����ID) Then
            tbrOften.Buttons("Often").value = tbrPressed
            Call tbrOften_ButtonClick(tbrOften.Buttons("Often"))
        End If
    Else
        fraInfo.Visible = False
        tvw_s.Visible = False
        fraLR.Visible = False
        chkSub.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        Line1(0).Visible = False
        Line1(1).Visible = False
        Shp.Visible = True
        chkShowCause.Visible = True

        '�����ƥ��ĸ�����Ŀ,������ʾ������Ŀ
        If ExistOftenItem Then
            tbrOften.Buttons("Often").value = tbrPressed
            Call SwitchToOften(False, False)
            Call FillList(True, str������ĿIDs, str�շ�ϸĿIDs)
            
            '���û�����л�����
            If Not cmdOK.Enabled Then
                tbrOften.Buttons("Often").value = tbrUnpressed
                Call SwitchToOften(False, False)
                Call FillList(True, str������ĿIDs, str�շ�ϸĿIDs)
            End If
        Else
            '���ƥ������
            Call FillList(True, str������ĿIDs, str�շ�ϸĿIDs)
        End If
        
        If cmdOK.Enabled And vsItem.Rows = vsItem.FixedRows + 1 Then
            'ֻ��һ����Ŀʱ,ֱ�ӷ���
            If tbrOften.Buttons("Often").value = tbrUnpressed Then
                mblnOK = True: Unload Me: Exit Sub
            Else
                blnDo = True '������Ŀƥ��ʱʼ����ʾ
            End If
        End If
        If (cmdOK.Enabled And vsItem.Rows > vsItem.FixedRows + 1) Or blnDo Then
            '������ͬһ����Ŀʱ,ֱ�ӷ���:�������շ�ϸĿID
            If mstr���� <> "" Then
                If UBound(Split(str������ĿIDs, ",")) = 1 _
                    And UBound(Split(str�շ�ϸĿIDs, ",")) <= 1 Then
                    '������Ŀƥ��ʱʼ����ʾ
                    If tbrOften.Buttons("Often").value = tbrUnpressed Then
                        mblnOK = True: Unload Me: Exit Sub
                    End If
                End If
            End If
        
            vsItem.Appearance = ccFlat
            vsItem.BorderStyle = ccFixedSingle
            
            Call SetFormSize
            Call Form_Resize
        Else
            '������,��ʾ,��ȡ���˳�,��ʾ�Ƿ�鿴δƥ�����Ŀ��
            If MsgBox("δ�ҵ�����ʹ�õ�������Ŀ��ҩƷ�����ģ������ǿ�治������⣬�Ƿ�鿴��ϸ��ԭ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set mrsItem = Nothing
                mblnOK = True: Unload Me: Exit Sub
            Else
                chkShowCause.value = 1
            End If
        End If
    End If
End Sub

Private Sub SetFormSize()
    Dim vRect As RECT, i As Long
    Dim lngUpH As Long, lngDnH As Long
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long

    Call zlControl.FormSetCaption(Me, False, False)
    Call GetWindowRect(mobjTXT.hwnd, vRect) '�����λ��
    
    '���ô���ߴ��λ��
    '������
    Me.Left = vRect.Left * Screen.TwipsPerPixelX
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 60 '+3D�߿�
    For i = 0 To vsItem.Cols - 1
        lngColW = lngColW + IIF(vsItem.ColHidden(i), 0, vsItem.ColWidth(i))
    Next
    If Me.Left + lngColW + lngScrW > Screen.Width - lngScrW Then
        Me.Width = Screen.Width - lngScrW - Me.Left
    Else
        Me.Width = lngColW + lngScrW
    End If
    
    '����߶�
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
    lngUpH = vRect.Top * Screen.TwipsPerPixelY '������ø߶�
    lngDnH = lngScrH - vRect.Bottom * Screen.TwipsPerPixelY '������ø߶�
    Me.Height = vsItem.Rows * vsItem.RowHeight(0) + tbrOften.Height + 45 '395 '+���Ƭ�߶�
    If Me.Height < 2000 Then Me.Height = IIF(mbytSize = 0, 2000, 2500) '������С�߶�
    If Me.Height > lngUpH And Me.Height > lngDnH Then
        Me.Height = IIF(lngUpH < lngDnH, lngDnH, lngUpH)
    End If
    If Me.Height > lngScrH / 2 Then Me.Height = lngScrH / 2 '�������߶�
    If Me.Height <= lngDnH Then
        Me.Top = vRect.Bottom * Screen.TwipsPerPixelY
    ElseIf Me.Height <= lngUpH Then
        Me.Top = vRect.Top * Screen.TwipsPerPixelY - Me.Height
    End If
End Sub
    
Private Sub SetOftenToolBar(ByVal blnCaption As Boolean)
'���ܣ����ù������Ƿ���ʾ�ı�
    Dim lngW As Long, i As Long
    
    For i = 1 To tbrOften.Buttons.Count
        tbrOften.Buttons(i).Caption = IIF(blnCaption, tbrOften.Buttons(i).Description, "")
    Next
    If blnCaption Then
        tbrOften.TextAlignment = tbrTextAlignRight
    Else
        tbrOften.TextAlignment = tbrTextAlignBottom
    End If
    
    For i = 1 To tbrOften.Buttons.Count
        If tbrOften.Buttons(i).Visible Then
            lngW = lngW + tbrOften.Buttons(i).Width
        End If
    Next
    tbrOften.Width = lngW
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    
    On Error Resume Next
    
    If mstr���� = "" Then
        fraInfo.Left = 0
        fraInfo.Width = Me.ScaleWidth
        chkSub.Left = fraInfo.Width - IIF(chkSub.Visible, chkSub.Width, 0) - 45
        lblInfo.Width = IIF(chkSub.Visible, chkSub.Left, fraInfo.Width) - lblInfo.Left - 45
        
        lngLeft = IIF(tvw_s.Visible, tvw_s.Width, 0) + IIF(fraStat.Visible, fraStat.Width, 0) + IIF(fraLR.Visible, fraLR.Width, 0)
        
        If tvw_s.Visible Then
            tvw_s.Left = 0
            tvw_s.Top = fraInfo.Top + fraInfo.Height + 15
            tvw_s.Height = Me.ScaleHeight - tvw_s.Top - 615
        End If
        If fraStat.Visible Then
            fraStat.Left = 0
            fraStat.Top = fraInfo.Top + fraInfo.Height - 90
            fraStat.Height = Me.ScaleHeight - fraStat.Top - 615
        End If
        If fraLR.Visible Then
            fraLR.Top = tvw_s.Top
            fraLR.Left = tvw_s.Left + tvw_s.Width
            fraLR.Height = tvw_s.Height
        End If
        
        vsItem.Top = fraInfo.Top + fraInfo.Height + 15
        vsItem.Left = lngLeft
        vsItem.Width = Me.ScaleWidth - lngLeft
        vsItem.Height = Me.ScaleHeight - vsItem.Top - IIF(mbytSize = 1, 750, 615) - IIF(tabClass.Visible, 350, 0)
        
        If tabClass.Visible Then
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
            tabClass.Left = vsItem.Left + 30
            tabClass.Width = vsItem.Width - 60
        End If
        
        Line1(0).X1 = 0: Line1(0).X2 = Me.ScaleWidth
        Line1(0).Y1 = tvw_s.Top + vsItem.Height + IIF(tabClass.Visible, 350, 0) + 75: Line1(0).Y2 = Line1(0).Y1
        
        Line1(1).X1 = Line1(0).X1: Line1(1).X2 = Line1(0).X2
        Line1(1).Y1 = Line1(0).Y1 - 15: Line1(1).Y2 = Line1(1).Y1
        
        cmdOK.Top = Line1(1).Y1 + 120
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.8 < 4000 Then
            cmdCancel.Left = 4000
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.2
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 15
        
        tbrOften.Top = cmdOK.Top + (cmdOK.Height - tbrOften.Height) / 2
        
        lblStore.Top = cmdOK.Top + 70
        lblStore.Left = tbrOften.Left + tbrOften.Width + 80
        
    Else
        Shp.Left = 0
        Shp.Top = 0
        Shp.Width = Me.ScaleWidth
        Shp.Height = Me.ScaleHeight
        
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        'vsItem.Height = Me.ScaleHeight - IIf(tabClass.Tabs.Count > 1, 380, 0)
        vsItem.Height = Me.ScaleHeight - tbrOften.Height + 15 - chkShowCause.Height
        
        tbrOften.Left = Me.ScaleWidth - tbrOften.Width - 15
        tbrOften.Top = vsItem.Top + vsItem.Height + chkShowCause.Height - 30
        
        If chkShowCause.Visible Then
            chkShowCause.Left = tbrOften.Left - chkShowCause.Width - 60
            chkShowCause.Top = tbrOften.Top + tbrOften.Height - chkShowCause.Height - 60
        End If
        
        If tabClass.Tabs.Count > 1 Then
            tabClass.Left = vsItem.Left + 60
            tabClass.Width = vsItem.Width - tbrOften.Width - 120
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
        End If
        
        lblStore.Top = tbrOften.Top + tbrOften.Height - chkShowCause.Height - 20
        lblStore.Left = 80
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngҩ��ID = 0 Then
        If tbrOften.Buttons("Often").value = tbrPressed Then
            Call SaveColPosition("Often")
            Call SaveColWidth("Often")
        Else
            Call SaveColPosition
            Call SaveColWidth
        End If
        Call SaveWinState(Me, App.ProductName, mstrSaveTag)
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsItem.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        vsItem.Left = vsItem.Left + X
        vsItem.Width = vsItem.Width - X
        tabClass.Left = tabClass.Left + X
        tabClass.Width = tabClass.Width - X
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node
    
    On Error GoTo errH
    
    If mlng����ID <> 0 Then
        strSQL = _
            " Select 1 as ��,����,ID,�ϼ�ID,����,���� From ���Ʒ���Ŀ¼ Where ID=[1]" & _
            " Union ALL " & _
            " Select Level+1 as ��,����,ID,�ϼ�ID,����,���� From ���Ʒ���Ŀ¼" & _
            " Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With �ϼ�ID=[1] Connect by Prior ID=�ϼ�ID" & _
            " Order by ��,����"
    Else
        strSQL = _
            " Select 0 as ��,����,-���� as ID,-Null as �ϼ�ID,����||'' as ����," & _
            " ����||'.'||Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',4,'��ҩ�䷽',5,'������Ŀ',6,'��������','7','��������') as ����" & _
            " From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Group by ����"
        strSQL = strSQL & " Union ALL " & _
            " Select Level as ��,����,ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,����,���� From ���Ʒ���Ŀ¼" & _
            " Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
            " Order by ��,����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID)
        
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!�ϼ�ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, "Close")
        End If
        objNode.Tag = rsTmp!���� '��ŷ�������
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        If mlng��λ����ID = 0 Then
            tvw_s.Nodes(1).Expanded = True
            If tvw_s.Nodes(1).Children > 0 Then
                tvw_s.Nodes(1).Child.Selected = True
            Else
                tvw_s.Nodes(1).Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Else
            tvw_s.Nodes("_" & mlng��λ����ID).Selected = True
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If
    
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tabClass_Click()
    If Not mblnClick Then Exit Sub
    
    If fraStat.Visible Then
        Call FillStat
    Else
        Call FillList
    End If
    vsItem.SetFocus
End Sub

Private Sub SwitchToOften(Optional ByVal blnFill As Boolean = True, Optional ByVal blnSaveColPos As Boolean = True)
'���ܣ��ڳ�����Ŀ��ѡ����Ŀ����֮���л�
'������blnFill=�л�֮���Ƿ�����ˢ���嵥
    Dim blnNoStat As Boolean
    
    '�䷽�ͳ����޷�ͳ�Ƴ���
    If mlng����ID <> 0 And Not tvw_s.SelectedItem Is Nothing Then
        If InStr(",4,6,", Val(tvw_s.SelectedItem.Tag)) > 0 Then
            blnNoStat = True
        End If
    End If
    tbrOften.Buttons("Stat").Visible = tbrOften.Buttons("Often").value = tbrPressed And mstr���� = "" And Not blnNoStat
    If mint��Χ = 1 Or mint��Χ = 2 Then
        '�����סԺ���û�г���ҽ��ͳ�Ƶ�Ȩ�ޣ�������ͳ�ư�ť
        If InStr(mstrPrivs, ";����ҽ��ͳ��;") = 0 Then
            tbrOften.Buttons("Stat").Visible = False
        End If
    End If
    tbrOften.Buttons("New").Visible = tbrOften.Buttons("Often").value = tbrUnpressed
    tbrOften.Buttons("Del").Visible = tbrOften.Buttons("Often").value = tbrPressed
    If mstr���� = "" Then
        chkSub.Visible = tbrOften.Buttons("Often").value = tbrUnpressed
        tvw_s.Visible = tbrOften.Buttons("Often").value = tbrUnpressed
        fraLR.Visible = tbrOften.Buttons("Often").value = tbrUnpressed
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").value = tbrPressed Then
                Call SaveColPosition(tvw_s.SelectedItem.Tag)
                Call SaveColWidth(tvw_s.SelectedItem.Tag)
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    Else
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").value = tbrPressed Then
                Call SaveColPosition
                Call SaveColWidth
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    End If
    Call SetOftenToolBar(mstr���� = "")
    Call Form_Resize
    
    If blnFill Then Call FillList(True)
End Sub

Private Sub SwitchToState(Optional ByVal blnFill As Boolean = True)
'���ܣ��ڳ�����Ŀ���棬����ͳ�ƺͳ��ý���֮���л�
'������blnFill=�л�֮���Ƿ�����ˢ���嵥
    If tbrOften.Buttons("Stat").value = tbrUnpressed Then
        fraStat.Visible = False
        tbrOften.Buttons("Del").Visible = True
        tbrOften.Buttons("New").Visible = False
        Call Form_Resize
        If blnFill Then Call FillList(True)
        If Visible Then vsItem.SetFocus
    Else
        fraStat.Visible = True
        lblInfo.Caption = "��ǰѡ��""" & UserInfo.���� & """�ĸ��˳�����Ŀ"
        tbrOften.Buttons("Del").Visible = False
        tbrOften.Buttons("New").Visible = True
        Call Form_Resize
        vsItem.FixedRows = 0: vsItem.Rows = 0
        vsItem.FixedCols = 0: vsItem.Cols = 0
        If Visible Then cboDate.SetFocus
    End If
End Sub

Private Sub tbrOften_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Often" Then
        If Visible And mstr���� <> "" Then
            zlControl.FormLock Me.hwnd
            Call SwitchToOften
            Call SetFormSize
            Call Form_Resize
            zlControl.FormLock 0
        Else
            '�л���ѡ�����ʱ�ȹر�ͳ�ƽ���
            If tbrOften.Buttons("Stat").value = tbrPressed Then
                tbrOften.Buttons("Stat").value = tbrUnpressed
                Call SwitchToState(False)
                Call SwitchToOften(, False)
            Else
                Call SwitchToOften
            End If
        End If
    ElseIf Button.Key = "Stat" Then
        Call SwitchToState
    ElseIf Button.Key = "New" Then
        Call NewOftenNew
    ElseIf Button.Key = "Del" Then
        Call OftenItemDel
    End If
End Sub

Private Sub NewOftenNew()
'���ܣ�����ǰ������Ŀ������˳�����Ŀ
    Dim arrSQL As Variant, i As Long, blnTran As Boolean
    Dim lngCol��Ŀ As Long, lngCol���� As Long
    Dim lngCol�շ�ϸĿID As Long, lngCol��� As Long
    
    arrSQL = Array()
    If Not fraStat.Visible Then
        If mrsItem.EOF Then Exit Sub
        
        ReDim arrSQL(0)
        arrSQL(0) = "ZL_���Ƹ�����Ŀ_Insert(" & UserInfo.ID & "," & mrsItem!������ĿID & ",Null,'" & _
                 mrsItem!���ID & "'," & ZVal(Val("" & mrsItem!�շ�ϸĿID)) & ")"
    Else
        lngCol��Ŀ = GetCol("������ĿID")
        lngCol�շ�ϸĿID = GetCol("�շ�ϸĿID")
        lngCol��� = GetCol("���ID")
        lngCol���� = GetCol("����")
        With vsItem
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 2)) <> 0 And Val(.TextMatrix(i, lngCol��Ŀ)) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_���Ƹ�����Ŀ_Insert(" & UserInfo.ID & "," & Val(.TextMatrix(i, lngCol��Ŀ)) & "," & _
                        Val(.TextMatrix(i, lngCol����)) & ",'" & .TextMatrix(i, lngCol���) & "'," & ZVal(Val(.TextMatrix(i, lngCol�շ�ϸĿID))) & ")"
                End If
            Next
        End With
        If UBound(arrSQL) < 0 Then
            MsgBox "������ѡ��һ��Ҫ����ĳ�����Ŀ��", vbInformation, gstrSysName
            vsItem.SetFocus: Exit Sub
        Else
            If MsgBox("�㵱ǰѡ���� " & UBound(arrSQL) + 1 & " ����Ŀ��Ҫ����Щ��Ŀ����Ϊ��ĸ��˳�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
        
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    Screen.MousePointer = 0
    
    If Not fraStat.Visible Then
        MsgBox "��Ŀ""" & mrsItem!���� & """�Ѿ�������ĸ��˳�����Ŀ��", vbInformation, gstrSysName
    Else
        MsgBox "��ѡ�����Ŀ�Ѿ�������ĸ��˳�����Ŀ��", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCol(ByVal strName As String) As Long
    Dim i As Long
    For i = 1 To vsItem.Cols - 1
        If UCase(vsItem.TextMatrix(0, i)) = UCase(strName) Then
            GetCol = i: Exit Function
        End If
    Next
End Function

Private Sub OftenItemDel()
'���ܣ�����ǰ����������Ŀ�Ƴ�
    Dim strSQL As String, lngRow As Long
    
    If mrsItem.EOF Then Exit Sub
    If MsgBox("ȷʵҪ��""" & mrsItem!���� & """����ĸ�����Ŀ���Ƴ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    lngRow = vsItem.Row
    
    strSQL = "ZL_���Ƹ�����Ŀ_Delete(" & UserInfo.ID & "," & mrsItem!������ĿID & "," & ZVal(Val("" & mrsItem!�շ�ϸĿID)) & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    With vsItem
        If lngRow = .FixedRows And .Rows = .FixedRows + 1 Then
            vsItem.Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
        Else
            .RemoveItem lngRow
            If lngRow <= .Rows - 1 Then
                .Row = lngRow
            Else
                .Row = .Rows - 1
            End If
        End If
        Call .ShowCell(.Row, .Col)
        Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstrPreNode Then Exit Sub
    '���ı�ʱ,���浱ǰ˳��(������)
    If Visible Then
        Call SaveColPosition(tvw_s.Nodes(mstrPreNode).Tag)
        Call SaveColWidth(tvw_s.Nodes(mstrPreNode).Tag)
    End If
    mstrPreNode = Node.Key
    
    Call FillList(True)
End Sub

Private Function GetTreePath(ByVal objNode As Node) As String
'���ܣ���ȡ����·����
    Dim tmpNode As Node, strTmp As String
    Set tmpNode = objNode
    Do While Not tmpNode Is Nothing
        strTmp = IIF(InStr(tmpNode.Text, "[") > 0, zlCommFun.GetNeedName(tmpNode.Text), Mid(tmpNode.Text, 3)) & "\" & strTmp
        Set tmpNode = tmpNode.Parent
    Loop
    GetTreePath = strTmp
End Function

Private Sub txtCount_GotFocus()
    Call zlControl.TxtSelAll(txtCount)
End Sub

Private Sub txtCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then
        If chkShowCause.value = 1 Then
            cmdOK.Enabled = False
            If NewRow >= vsItem.FixedRows Then
                mrsItem.Filter = "KeyID=" & Val(vsItem.TextMatrix(NewRow, GetCol("KeyID")))
                If mrsItem.RecordCount = 1 Then
                    If InStr(mrsItem!δƥ��ԭ��, "��Ŀ���벻ƥ��") > 0 Or InStr(mrsItem!δƥ��ԭ��, "������Ŀ��ִ��Ƶ��Ϊ") > 0 And mint��Χ <> 1 Then
                        cmdOK.Enabled = True
                    End If
                End If
            End If
        Else
            If NewRow >= vsItem.FixedRows Then
                mrsItem.Filter = "KeyID=" & Val(vsItem.TextMatrix(NewRow, GetCol("KeyID")))
                'ͳ�Ƴ�����Ŀ����ֱ��ѡ��,��Ϊû�й�Ȩ��,�ٳ�����
                cmdOK.Enabled = mrsItem.RecordCount = 1 And Not fraStat.Visible
            Else
                cmdOK.Enabled = False
            End If
        End If
        cmdOK.Visible = Not fraStat.Visible And mstr���� = ""
        Call ShowDrugStore(NewRow)
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    If Order = 0 Then Exit Sub
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '��Ϊ������˳��ı�,���Ա���ԭʼ�к�
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").value = tbrPressed Then strType = "Often" '�̶�
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    If vsItem.ColDataType(Col) = flexDTBoolean Then
        Order = 0
    Else
        'ǿ�Ʊ����а��ַ�������
        If vsItem.TextMatrix(0, Col) = "����" Then
            If Order = 1 Then Order = 7
            If Order = 2 Then Order = 8
        End If
    End If
End Sub

Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) = flexDTBoolean Then Cancel = True
End Sub

Private Sub ShowDrugStore(ByVal lngRow As Long)
'���ܣ���ʾ����ҩ���еĿ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strҩ�� As String
    Dim lngҩƷID As Long
    Dim str��Χ As String
    Dim i As Long
    Dim strTmp As String
    Dim blnDo As Boolean
    Dim lngCol��� As Long
    Dim lng��� As Long
    
    On Error GoTo errH
    
    lblStore.Caption = ""
    lblStore.ToolTipText = ""
    
    With vsItem
        If .Row >= .FixedRows Then
            lng��� = Val(.TextMatrix(.Row, GetCol("���ID")))
            If InStr(",5,6,7,", lng���) > 0 Then
                lngCol��� = GetCol("���")
                If .ColHidden(lngCol���) = False Then
                    If vsItem.TextMatrix(0, lngCol���) = "���" Then
                        lngҩƷID = Val(.TextMatrix(.Row, GetCol("�շ�ϸĿID")))
                        If lngҩƷID <> 0 And (mint��Χ = 1 Or mint��Χ = 2) Then
                            blnDo = True
                        End If
                    End If
                End If
            End If
            If .Cell(flexcpData, .Row, lngCol���) <> "" Then
                lblStore.Caption = .Cell(flexcpData, .Row, lngCol���)
                lblStore.ToolTipText = .Cell(flexcpData, .Row, lngCol���)
                blnDo = False
            End If
        End If
        If blnDo Then
            Select Case lng���
            Case 5
                strҩ�� = mstr������ҩ��
            Case 6
                strҩ�� = mstr���ó�ҩ��
            Case 7
                strҩ�� = mstr������ҩ��
            End Select
            
            str��Χ = Decode(mint��Χ, 1, "C.����", 2, "C.סԺ")
            
            strSQL = "Select a.����,Decode(x.���,Null,Null,Round(x.���/" & str��Χ & "��װ,5)||" & str��Χ & "��λ) As ���" & _
                " From (Select a.ҩƷid, a.�ⷿid, Nvl(Sum(a.��������), 0) As ���  From ҩƷ��� A" & _
                " Where a.���� = 1 And a.�ⷿid In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
                " And (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And a.ҩƷid=[2]" & _
                " Group By a.ҩƷid,a.�ⷿid" & _
                " Having Nvl(Sum(a.��������),0) <> 0) X, ҩƷ��� C, ���ű� A" & _
                " Where x.ҩƷid = c.ҩƷid And a.Id = x.�ⷿid"
                
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strҩ��, lngҩƷID)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    strTmp = strTmp & "," & rsTmp!���� & ":" & rsTmp!���
                    rsTmp.MoveNext
                Next
            End If
            If strTmp = "" Then
                strTmp = "����ҩ�����޿��."
            Else
                strTmp = Mid(strTmp, 2) & "."
            End If
            .Cell(flexcpData, .Row, lngCol���) = strTmp
            lblStore.Caption = strTmp
            lblStore.ToolTipText = strTmp
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) <> flexDTBoolean Then
        Cancel = True
    ElseIf Val(vsItem.TextMatrix(Row, 1)) = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        If cmdOK.Enabled Then
            Call vsItem_KeyPress(13)
        ElseIf fraStat.Visible Then
            Call vsItem_KeyPress(32)
        End If
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then
            cmdOK_Click
        ElseIf fraStat.Visible Then
            If vsItem.Row + 1 <= vsItem.Rows - 1 Then
                vsItem.Row = vsItem.Row + 1
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If fraStat.Visible Then
            If Val(vsItem.TextMatrix(vsItem.Row, 1)) <> 0 Then
                If Val(vsItem.TextMatrix(vsItem.Row, 2)) = 0 Then
                    vsItem.TextMatrix(vsItem.Row, 2) = 1
                Else
                    vsItem.TextMatrix(vsItem.Row, 2) = 0
                End If
            End If
        End If
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub

Private Sub SaveColPosition(Optional ByVal strType As String)
'���ܣ�������˳��:�к�,˳��|...
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    If tbrOften.Buttons("Stat").value = tbrPressed Or fraStat.Visible Then Exit Sub
    
    With vsItem
        For i = 0 To .Cols - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr���� = "" And strType = "" And tvw_s.Visible And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub RestoreColPosition()
'���ܣ��ָ���˳��
'˵����Ӧ����������֮ǰ
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    
    With vsItem
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").value = tbrPressed Then strType = "Often" '�̶�
        strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,�ı������к�Ҳ�ı�
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .Cols - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .Cols - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub SaveColWidth(Optional ByVal strType As String)
'���ܣ������п��
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    If mstr���� = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'���ܣ��ָ��п��
'˵����Ӧ���ڻָ�����֮��
    Dim strType As String
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    
    If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColSort()
'���ܣ�������
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0 Then
            If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
            If tbrOften.Buttons("Often").value = tbrPressed Then strType = "Often" '�̶�
            strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '��Ϊ���ܵ�����˳��,���Բ�����ʵ��������
                For i = 0 To .Cols - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .Cols - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(1).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(2).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function FillList(Optional ByVal blnClass As Boolean, _
    Optional str������ĿIDs As String, Optional str�շ�ϸĿIDs As String) As Boolean
'���ܣ����ݵ�ǰ��������װ��������ĿĿ¼
'������blnClass=�Ƿ��ؽ����࿨(Ӧ��������Ŀ�ı�ʱ���ؽ�)
    Dim objNode As Node, objItem As ListItem
    Dim strSQL As String, strInside As String
    Dim arrClass As Variant, strClass As String
    Dim strSub As String, str�������� As String
    Dim str�Ա� As String, strStock As String
    Dim strInput As String, lngҩ��ID As Long
    Dim blnLoad As Boolean, objTab As MSComctlLib.Tab
    Dim str��Χ As String, strҩƷ As String
    Dim blnOften As Boolean, blnStock As Boolean
    Dim str������� As String, strPriv As String
    Dim i As Long, j As Long
    Dim strCommIF As String, strScope As String
    Dim blnIsHaveKSS As Boolean
    Dim str������� As String
    Dim blnҩƷ As Boolean, bln���� As Boolean, bln���� As Boolean
    Dim lng����ID As Long, int���� As Integer, str��� As String
    Dim bln��ʾ��� As Boolean
    Dim str��ȡ�ֶ� As String
    Dim blnBarcode As Boolean          '����ʱ�����Ƿ����ҩƷ������Ʒ����/�ڲ�������ƥ��
    Dim strҽ������ As String
    Dim strTsPriv As String

    str������ĿIDs = "": str�շ�ϸĿIDs = ""
    Set objNode = tvw_s.SelectedItem '����ΪNothing
    blnOften = tbrOften.Buttons("Often").value = tbrPressed And mlngҩ��ID = 0 '�Ƿ���ʾ������Ŀ
    
    '
    blnҩƷ = True: bln���� = True: bln���� = True
    If mstr���Ʒ��� <> "" And mstr���� <> "" Then
       blnҩƷ = InStr(",5,6,7,", mstr���Ʒ���) > 0
       bln���� = mstr���Ʒ��� = "4"
       bln���� = Not (InStr(",4,5,6,7,", mstr���Ʒ���) > 0)
    ElseIf mlngҩ��ID <> 0 Then
        If mintType = 0 Then
            bln���� = False: blnҩƷ = True: bln���� = False    '��ʾָ��ҩƷ���й��
        ElseIf mintType = 1 Then
            bln���� = True: blnҩƷ = False: bln���� = False    '��ʾָ���������й��
        End If
    End If
    
    '�Ƿ���ʾ���ѡ��
    If mint���� <> -1 Then
        blnStock = mstr���� <> "" And tabClass.SelectedItem.Index = 1 _
            And ((gblnҩƷ�������ҽ�� Or mint��Ч = 1) And Not (mlng��ҩ�� = 0 And mlng��ҩ�� = 0 And mlng��ҩ�� = 0) _
            Or mlng���ϲ��� <> 0)
    Else
        blnStock = False
    End If
    
    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '�����������ֶ�����(mstr���Ʒ��಻Ϊ��ʱ,������ʾ���ܵ���Ӧ�õ���Ŀ)
    '------------------------------------------------------------------------
    If mint���� = 2 Then
        strҽ������ = " and a.��� <> '9' Or a.��� = '9' And Exists (Select 1 From �������ÿ��� Where ��Ŀid = a.Id And Instr([23], ',' || ����id || ',') > 0) "
    End If
    strCommIF = " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.������� IN([8],3) Or [8]=3 And Nvl(A.�������,0)<>0)"
    strScope = " And ((A.���<>'9' Or A.���='9' And (A.��ԱID=[11] Or A.��ԱID is Null))" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And Instr([17],','||����ID||',')>0)" & strҽ������ & _
            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID)))" & _
             IIF(mstr���Ʒ��� <> "", "", "And Nvl(A.����Ӧ��,0)=1") & " And Instr([10],','||Nvl(a.�����Ա�,0)||',')>0 And Nvl(A.ִ��Ƶ��,0) IN(0,[9])"
            
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "0,1"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "0,2"
    ElseIf mstr�Ա� = "" Then
        str�Ա� = "0,1,2"
    Else
        str�Ա� = "0"
    End If
    
    If chkShowCause.value = 1 Then
        strCommIF = "": strScope = " And (A.���<>'9' Or A.���='9' And (A.��ԱID=[11] Or A.��ԱID is Null))"
    End If
    
    '������Ŀ�Ĳ�������
    str�������� = "Decode(A.���," & _
        "'H',Decode(A.��������,'1','����ȼ�','������')," & _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨','4','��ҩ�÷�','5','��������','6','�ɼ�����','7','��Ѫ����','8','��Ѫ;��',Null)," & _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','7','����','8','����','9','����','10','��Σ','11','����','12','��¼�����','14','��ǰ',NULL)," & _
        "A.��������)"
    
    If mstr���� = "" Then
        If mlngҩ��ID <> 0 Then
            strSub = " And A.ID = [19]"
        Else
            int���� = Val(objNode.Tag): lng����ID = Val(Mid(objNode.Key, 2))
            If Not blnOften Then
                '�����еķ���ID
                If chkSub.value = 1 Then
                    '��ʾ�¼�����Ŀ
                    If Val(Mid(objNode.Key, 2)) < 0 Then
                        strSub = " And A.����ID IN(" & _
                            " Select ID From ���Ʒ���Ŀ¼ Where ����=[1] And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                            " )"
                    Else
                        strSub = " And A.����ID IN(" & _
                            " Select ID From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
                            " Start With ID=[3] Connect by Prior ID=�ϼ�ID)"
                    End If
                Else
                    strSub = " And A.����ID=[3]"
                End If
            ElseIf mlng����ID <> 0 Then 'ͨ��������ȷ���ķ�����������г�����Ŀ
                strSub = " And A.����ID IN(" & _
                    " Select ID From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
                    " Start With ID=[4] Connect by Prior ID=�ϼ�ID" & _
                    " )"
            Else
                '��ʾ���з���,����еĸ��˳�����Ŀ
            End If
            
            If Not blnOften Or mlng����ID <> 0 Then
                '�����е�����ȷ�����
                If Val(objNode.Tag) = 5 Then
                    strSub = strSub & " And A.��� Not IN('4','5','6','7','8','9')"
                Else
                    If Val(objNode.Tag) < 8 And Val(objNode.Tag) > 0 Then str��� = Choose(Val(objNode.Tag), "5", "6", "7", "8", "", "9", "4")
                    If str��� <> "" Then strSub = strSub & " And A.���=[2]"
                End If
            End If
        End If
    Else
        '����ƥ��:�޷�ȷ�����༰���,��������Ŀ��ƥ��
        If Len(mstr����) < 2 Then mstrLike = "" '�Ż�
        strInput = " And (A.���� Like [5] And B.����=[7]" & _
            " Or B.���� Like [6] And B.����=[7] Or B.���� Like [6] And B.���� IN([7],3))"
    End If
    
    '���Ƭȷ�����
    If tabClass.SelectedItem.Key <> "" Then
        str��� = Mid(tabClass.SelectedItem.Key, 2)
        strSub = strSub & " And A.���=[2]"
    End If
        
    'ģ��Ȩ��
    If mint��Χ = 1 Then
        strPriv = GetInsidePrivs(p����ҽ���´�)
    ElseIf mint��Χ = 2 Then
        strPriv = GetInsidePrivs(pסԺҽ���´�)
    End If
    
    If mint��Χ = 1 Then
        strTsPriv = GetTsPrivs(p����ҽ���´�)
    ElseIf mint��Χ = 2 Then
        strTsPriv = GetTsPrivs(pסԺҽ���´�)
    End If
    
    '����ҩƷȨ��
    strҩƷ = ""
    If strTsPriv <> "" And chkShowCause.value <> 1 Then
        If InStr(strTsPriv, "�´�����ҩ��") = 0 Then strҩƷ = strҩƷ & " And D.�������<>'����ҩ'"
        If InStr(strTsPriv, "�´ﶾ��ҩ��") = 0 Then strҩƷ = strҩƷ & " And D.�������<>'����ҩ'"
        If InStr(strTsPriv, "�´ﾫ��ҩ��") = 0 Then strҩƷ = strҩƷ & " And D.������� Not IN('����I��')"
        If InStr(strTsPriv, "�´����ҩ��") = 0 Then strҩƷ = strҩƷ & " And D.��ֵ���� Not IN('����','����')"
    End If
    
    '·����������ʱ,ֻ��ʾָ�����Ʒ����������Ŀ
    If mstr���Ʒ��� <> "" Then
        str������� = "A.��� ='" & mstr���Ʒ��� & "'"
        If InStr(",C,D,F,G,E,H,Z,", mstr���Ʒ���) > 0 Then
            If mstr�������� <> "" Then str������� = str������� & " And A.��������='" & mstr�������� & "'"
        End If
        If mstr���Ʒ��� = "E" Or (mstr���Ʒ��� = "D" And mstr�������� = "18") Then
            If Val(mstrִ�з���) <> 0 Then str������� = str������� & " And A.ִ�з���=" & mstrִ�з���
        End If
        
    End If
    
    '��ȡ����
    
    '1.ҩƷ�б�
    If mstr���� <> "" And blnҩƷ Then
        strInput = " And (A.���� Like [5] And B.����=[7]" & _
            " Or B.���� Like [6] And B.����=[7] Or B.���� Like [6] And B.���� IN([7],3))"
        If IsNumeric(mstr����) Then
            '1X.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.���� Like [5] And B.����=[7] Or B.���� Like [6] And B.����=3)"
        ElseIf zlCommFun.IsCharAlpha(mstr����) Then
            'X1.����ȫ����ĸʱֻƥ�����
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.���� Like [6] And B.����=[7]"
        ElseIf zlCommFun.IsCharChinese(mstr����) Then
            '��������,��ֻƥ������
            strInput = " And B.���� Like [6] And B.����=[7]"
        End If
    End If
    '��Ʒ���´�ĳ���
    If Not (gblnҩƷ�������ҽ�� Or mint��Ч = 1) And blnҩƷ Then
        
        'ҩƷ������Ŀ����:��������ҩƷ����ʱ�Ŷ�ȡ
        '--------------------------------------------------------------------------------------
        blnLoad = False
        If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
        End If
        If blnLoad Then
            If mstr���� <> "" Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & IIF(Not mbln����, "Distinct", "") & _
                        " A.��� As ���ID,A.ID as ������ĿID,-Null as �շ�ϸĿID," & _
                        " F.���� As ���,Null as ����,A.����,B.����,Null as ��Ʒ��," & IIF(mbln����, "B.����,", "") & _
                        " A.���㵥λ,Null as ���,Null as ����,D.ҩƷ����," & str�������� & " As ��Ŀ����," & _
                        " Null as ��������,Null as ҽ������,Null as ˵��,D.����ְ�� as ����ְ��ID,Null as �۸�,Null as ���,Decode(d.������,0,'',1,'������ʹ��',2,'����ʹ��',3,'����ʹ��') as �����ȼ�,D.�ٴ��Թ�ҩ as �ٴ��Թ�ҩID,NULL as ����" & _
                        IIF(chkShowCause.value = 1, ",Null as �շѳ���ʱ��,a.����ʱ��,A.վ��,Nvl(D.������,0) as ������,null as ���÷������,D.������� as �������,D.��ֵ���� as ��ֵ����,a.�������,Nvl(a.ִ��Ƶ��, 0) as ִ��Ƶ��,Null as ��ҩ��������," & vbNewLine & _
                        "  Null as ��ҩ��������,Null as ��ҩ��������,Null as ʹ�ÿ���ID,Null as ����Ӧ��,Null as �������,Nvl(A.�����Ա�,0) as �����Ա�,b.����,NULL AS δƥ��ԭ��", "") & _
                    " From ҩƷ���� D,������Ŀ��� F,������Ŀ���� B,������ĿĿ¼ A" & _
                    " Where A.ID=B.������ĿID And A.ID=D.ҩ��ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, str�������, "A.��� = Null"), "A.��� IN ('5','6','7')") & strCommIF & _
                        IIF(chkShowCause.value = 1, "", " And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0 And Nvl(A.ִ��Ƶ��,0) IN(0,[9])") & strInput & strSub & strҩƷ
            Else
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select A.��� As ���ID,A.ID as ������ĿID,-Null as �շ�ϸĿID," & _
                        " F.���� As ���,Null as ����,A.����,A.����,Null as ��Ʒ��," & IIF(mbln����, "Null as ����,", "") & _
                        "A.���㵥λ,Null as ���,Null as ����, D.ҩƷ����," & str�������� & " As ��Ŀ����," & _
                        "Null as ��������,Null as ҽ������,Null as ˵��,D.����ְ�� as ����ְ��ID, Null as �۸�,Null as ��� ,Decode(d.������,0,'',1,'������ʹ��',2,'����ʹ��',3,'����ʹ��') as �����ȼ�,D.�ٴ��Թ�ҩ as �ٴ��Թ�ҩID,NULL as ����" & _
                    " From ҩƷ���� D,������Ŀ��� F,������ĿĿ¼ A" & _
                    " Where A.ID=D.ҩ��ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, str�������, "A.��� = Null"), "A.��� IN ('5','6','7')") & strCommIF & _
                        " And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0 And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strSub & strҩƷ
            End If
        End If
    Else
        
        'ҩƷ��񲿷�:��������ҩƷ����ʱ�Ŷ�ȡ
        '--------------------------------------------------------------------------------------
        blnLoad = False
        If blnҩƷ Then
            If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
                blnLoad = True
            Else
                blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
            End If
        End If
        
        
        If blnLoad Then
            'ҩƷ���,ĳһ��ҩ��δָ��ʱ,����������¼
            strStock = ""
            If mint���� <> -1 Then
                If mstr���� = "" Then '���ݷ������ȷ��ҩƷ���
                    If Val(objNode.Tag) = 1 Then
                        lngҩ��ID = mlng��ҩ��
                    ElseIf Val(objNode.Tag) = 2 Then
                        lngҩ��ID = mlng��ҩ��
                    ElseIf Val(objNode.Tag) = 3 Then
                        lngҩ��ID = mlng��ҩ��
                    End If
                Else
                    'û�з���,�޷�ȷ��ҩƷ���
                    If Mid(tabClass.SelectedItem.Key, 2) = "5" Then
                        lngҩ��ID = mlng��ҩ��
                    ElseIf Mid(tabClass.SelectedItem.Key, 2) = "6" Then
                        lngҩ��ID = mlng��ҩ��
                    ElseIf Mid(tabClass.SelectedItem.Key, 2) = "7" Then
                        lngҩ��ID = mlng��ҩ��
                    End If
                End If
                If chkShowCause.value <> 1 And mbln��ʾ��� Then
                    If lngҩ��ID <> 0 Then
                        strStock = _
                            "Select A.ҩƷID,Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
                            " Where A.���� = 1 And A.�ⷿID=[12]" & _
                            " And (Nvl(A.����, 0) = 0 Or A.Ч�� Is Null Or A.Ч�� > Trunc(Sysdate))" & _
                            " Group by A.ҩƷID Having Nvl(Sum(A.��������),0)<>0"
                        bln��ʾ��� = True
                    ElseIf blnStock And Not (mlng��ҩ�� = 0 And mlng��ҩ�� = 0 And mlng��ҩ�� = 0) Then
                        strStock = _
                            "Select C.ҩƷID,Nvl(Sum(C.��������),0) as ���" & _
                            " From ҩƷ��� C,�շ���ĿĿ¼ A" & IIF(strInput <> "", ",�շ���Ŀ���� B", "") & _
                            " Where C.���� = 1 And (Nvl(C.����,0)=0 Or C.Ч�� Is Null Or C.Ч��>Trunc(Sysdate))" & _
                                " And C.�ⷿID=Decode(A.���,'5',[13],'6',[14],'7',[15],Null)" & _
                                " And C.ҩƷID=A.ID And A.��� IN('5','6','7')" & _
                                 IIF(strInput <> "", " And A.ID=B.�շ�ϸĿid " & strInput, "") & _
                            " Group by C.ҩƷID Having Nvl(Sum(C.��������),0)<>0"
                        bln��ʾ��� = True
                        'strStock = "" '�Ż�
                    End If
                End If
            End If
            
            str��Χ = Decode(mint��Χ, 1, "C.����", 2, "C.סԺ", 3, "A.����")
                
            '�Ƿ�����п��:ָ��ҩ��ʱ����ϵͳ�����Ƿ�Ҫ���ƿ��
            If strStock = "" Then
                str������� = ""
            Else
                str������� = " And A.ID=X.ҩƷID(+)"
            End If
            If Not (mstr������ҩ�� = "" And mstr���ó�ҩ�� = "" And mstr������ҩ�� = "") And chkShowCause.value <> 1 Then
                '��ʹ�ð󶨱�������Ϊ��������������Ծ�̬��
                If gblnStock Then
                    str������� = str������� & " And (D.�ٴ��Թ�ҩ=1 Or (" & _
                        " A.���='5'" & IIF(mstr������ҩ�� = "", "", " And Exists(Select 1 From ҩƷ���" & _
                        " Where ҩƷID = c.ҩƷID And ���� = 1 And (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ��������>0 And �ⷿID In(" & mstr������ҩ�� & "))") & _
                        " Or A.���='6'" & IIF(mstr���ó�ҩ�� = "", "", " And Exists(Select 1 From ҩƷ���" & _
                        " Where ҩƷID = c.ҩƷID And ���� = 1 And (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ��������>0 And �ⷿID In(" & mstr���ó�ҩ�� & "))") & _
                        " Or A.���='7'" & IIF(mstr������ҩ�� = "", "", " And Exists(Select 1 From ҩƷ��� " & _
                        " Where ҩƷID = c.ҩƷID And ���� = 1 And (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ��������>0 And �ⷿID In(" & mstr������ҩ�� & "))") & _
                        "))"
                Else
                'ֻ��ʾָ��ҩ����ҩƷ�����ù����õĲ��˿��ң�
                    str������� = str������� & " And (D.�ٴ��Թ�ҩ=1 Or Exists(Select 1 From �շ�ִ�п��� X Where x.�շ�ϸĿid = c.ҩƷID And " & _
                           "(A.���='5'" & IIF(mstr������ҩ�� = "", "", " And x.ִ�п���id In (" & mstr������ҩ�� & ")") & _
                        " Or A.���='6'" & IIF(mstr���ó�ҩ�� = "", "", " And x.ִ�п���id In (" & mstr���ó�ҩ�� & ")") & _
                        " Or A.���='7'" & IIF(mstr������ҩ�� = "", "", " And x.ִ�п���id In (" & mstr������ҩ�� & ")") & _
                        ")" & IIF(mint��Χ <> 3, " And (x.������Դ is NULL Or x.������Դ=[8])", "") & "))"
                End If
            End If
            If mstr���� <> "" Then
                '���Ƹ��������ƥ����ʾ
                strInside = "Select " & IIF(Not mbln����, "Distinct", "") & _
                    " A.ID,A.���,A.����," & IIF(gbyt����ҩƷ��ʾ = 1, "C2.���� ,C1.���� as ��Ʒ��,", "B.����,Null as ��Ʒ��,") & IIF(mbln����, "B.����,", "") & _
                    " A.���㵥λ as ���۵�λ,1 as ���۰�װ,A.���,A.����,A.��������," & _
                        IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������,", "Null as ҽ������,") & "A.˵��,A.�Ƿ���" & _
                        IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,b.����", "") & _
                    " From �շ���Ŀ���� B,�շ���ĿĿ¼ A" & IIF(gbyt����ҩƷ��ʾ = 1, ",�շ���Ŀ���� C2,�շ���Ŀ���� C1", "") & _
                      IIF(mint���� <> 0, ",����֧����Ŀ M,����֧������ N", "") & _
                    " Where A.ID=B.�շ�ϸĿID And A.��� IN ('5','6','7')" & _
                    IIF(gbyt����ҩƷ��ʾ = 1, " And A.ID=C1.�շ�ϸĿID(+) And C1.����(+)=1 And C1.����(+)=3 And A.ID=C2.�շ�ϸĿID(+) And C2.����(+)=1 And C2.����(+)=1", "") & _
                    strCommIF & strInput & _
                    IIF(mint���� <> 0, " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[18]", "")
                If mbln�۸� Then
                    strInside = "Select A.ID,A.���,A.����,A.����,A.��Ʒ��," & IIF(mbln����, "A.����,", "") & _
                        " A.���۵�λ,A.���۰�װ,A.���,A.����,A.��������,A.ҽ������,A.˵��,Sum(Decode(A.�Ƿ���,1,NULL,B.�ּ�)) as �۸�,Sum(b.�ּ�) as �ּ�" & _
                        IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,A.����", "") & _
                        " From �շѼ�Ŀ B,(" & strInside & ") A" & _
                        " Where A.ID=B.�շ�ϸĿID And Sysdate Between B.ִ������+0 And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "20", "21", "22") & _
                        " Group by A.ID,A.���,A.����,A.����,A.��Ʒ��," & IIF(mbln����, "A.����,", "") & _
                        " A.���۵�λ,A.���۰�װ,A.���,A.����,A.��������,A.ҽ������,A.˵��" & IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,A.����", "")
                ElseIf mstrLike = "" And strStock <> "" Then
                    '���������ü�������ʱ(����ƥ��),�����(+)����(ҩƷ���),����ҪGroup Byһ��(���)
                    '��Group by ��Distinct ͬʱ����ʱ(Not mbln����)��Oracle��ֻѡ�����Group by
                    strInside = Replace(strInside, "A.�Ƿ���", "Null as �۸�")
                    strInside = strInside & " Group By A.ID,A.���,A.����," & IIF(gbyt����ҩƷ��ʾ = 1, "C2.���� ,C1.����,", "B.����,") & IIF(mbln����, "B.����,", "") & _
                        " A.���㵥λ,A.���,A.����,A.��������," & IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')'),", "") & "A.˵��,A.�Ƿ���" & IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,b.����", "")
                End If
                
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & " Select " & _
                        " A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                        " F.���� AS ���,C.����ҩ�� as ����,A.����,A.����,A.��Ʒ��," & IIF(mbln����, "A.����,", "") & _
                        " E.���㵥λ,A.���,A.����,D.ҩƷ����,Null as ��Ŀ����,A.��������,A.ҽ������,A.˵��,D.����ְ�� as ����ְ��ID" & _
                        IIF(mbln�۸�, IIF(chkShowCause.value = 1, ",Sum(Decode(A.�۸�, Null, decode(C.�ϴ��ۼ�,null,A.�ּ�,C.�ϴ��ۼ�), A.�ּ�)) * ", ",Decode(A.�۸�, Null, decode(C.�ϴ��ۼ�,null,A.�ּ�,C.�ϴ��ۼ�), A.�ּ�) * ") & str��Χ & "��װ || '/' || " & str��Χ & "��λ As �۸�", ",Null as �۸�") & _
                        IIF(strStock <> "", _
                            IIF(InStr(strPriv, "��ʾҩƷ���") = 0, _
                                ",Decode(Sign(Nvl(X.���,0)),1,'��','') as ���", _
                                ",Decode(X.���,NULL,NULL,Round(X.���/" & str��Χ & "��װ,5)||" & str��Χ & "��λ) as ���"), _
                            ",Null as ���") & ",Decode(d.������,0,'',1,'������ʹ��',2,'����ʹ��',3,'����ʹ��') as �����ȼ�,D.�ٴ��Թ�ҩ as �ٴ��Թ�ҩID,NULL as ����" & _
                            IIF(chkShowCause.value = 1, ",a.����ʱ�� as �շѳ���ʱ��,e.����ʱ��,A.վ��,Nvl(D.������,0)  as ������,a.������� as ���÷������,D.�������,D.��ֵ����," & vbNewLine & _
                    "              e.�������,Nvl(e.ִ��Ƶ��, 0) as ִ��Ƶ��,decode(D.�ٴ��Թ�ҩ,1,1,max(decode(a.���,'5', decode(instr('," & mstr������ҩ�� & ",',',' || n." & IIF(gblnStock, "�ⷿid", "ִ�п���id") & " || ','),0,0,NVL(N." & IIF(gblnStock, "��������", "ִ�п���id") & ",0)),0))) as ��ҩ��������," & vbNewLine & _
                    "              decode(D.�ٴ��Թ�ҩ,1,1,max(decode(a.���,'6', decode(instr('," & mstr���ó�ҩ�� & ",',',' || n." & IIF(gblnStock, "�ⷿid", "ִ�п���id") & " || ','),0,0,NVL(N." & IIF(gblnStock, "��������", "ִ�п���id") & ",0)),0))) as ��ҩ��������," & vbNewLine & _
                    "              decode(D.�ٴ��Թ�ҩ,1,1,max(decode(a.���,'7',decode(instr('," & mstr������ҩ�� & ",',',' || n." & IIF(gblnStock, "�ⷿid", "ִ�п���id") & " || ','),0,0,NVL(N." & IIF(gblnStock, "��������", "ִ�п���id") & ",0)),0))) as ��ҩ��������,null as ʹ�ÿ���ID, Null as ����Ӧ��,null as �������, nvl(e.�����Ա�,0) as �����Ա�,A.����,NULL AS δƥ��ԭ��", "") & _
                    " From ҩƷ��� C,ҩƷ���� D,������ĿĿ¼ E,�շ���Ŀ��� F,(" & strInside & ") A" & IIF(chkShowCause.value = 1, IIF(gblnStock, ",ҩƷ��� N", ",�շ�ִ�п��� N"), "") & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                    " Where A.ID=C.ҩƷID And C.ҩ��ID=D.ҩ��ID And D.ҩ��ID=E.ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, Replace(str�������, "A.", "E."), "E.��� = Null"), " E.��� IN ('5','6','7')") & _
                        IIF(chkShowCause.value = 1, "", " And Instr([10],','||Nvl(e.�����Ա�,0)||',')>0") & _
                        IIF(chkShowCause.value = 1, IIF(gblnStock, " And N.ҩƷID(+) = c.ҩƷID AND n.����(+) = 1 And (Nvl(n.����, 0) = 0 Or n.Ч�� Is Null Or n.Ч�� > Trunc(Sysdate))", " And N.�շ�ϸĿid(+) = c.ҩƷID " & IIF(mint��Χ <> 3, " And (N.������Դ is NULL Or N.������Դ=[8])", "")), "") & _
                        IIF(chkShowCause.value <> 1, " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                        " And (E.������� IN([8],3) Or [8]=3 And Nvl(E.�������,0)<>0) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])", "") & _
                        str������� & strҩƷ & Replace(strSub, "A.", "E.") & _
                        IIF(chkShowCause.value = 1, " Group by a.���, e.Id, a.Id, f.����,c.����ҩ��,a.����, a.����," & vbNewLine & _
                        "              a.��Ʒ��" & IIF(mbln����, ",A.����", "") & ", e.���㵥λ, a.���, a.����, d.ҩƷ����, a.��������, a.ҽ������, a.˵��, d.����ְ��," & IIF(mbln�۸�, "a.�۸�," & str��Χ & "��װ, " & str��Χ & "��λ,", "") & vbNewLine & _
                        "              d.������,a.����ʱ��,e.����ʱ��,A.վ��,a.�������,e.�������,e.ִ��Ƶ��,D.�ٴ��Թ�ҩ,D.�������,D.��ֵ����,A.����,e.�����Ա� ", "")

            Else
                '��ҩ���Ƹ��ݲ���������ʾ
                If mbln�۸� Then
                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                        "Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                            " F.���� AS ���,C.����ҩ�� as ����,A.����,Nvl(G1.����,A.����) as ����," & IIF(gbytҩƷ������ʾ = 2, "G2.����", "Null") & " as ��Ʒ��," & IIF(mbln����, "Null as ����,", "") & "E.���㵥λ,A.���,A.����," & _
                            " D.ҩƷ����,Null as ��Ŀ����,A.��������," & IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������,", "Null as ҽ������,") & "A.˵��,D.����ְ�� as ����ְ��ID," & _
                            " Decode(A.�Ƿ���,1,decode(Sum(c.�ϴ��ۼ�),null,Sum(B.�ּ�),sum(c.�ϴ��ۼ�))*" & str��Χ & "��װ || '/' || " & str��Χ & "��λ,Sum(B.�ּ�)* " & str��Χ & "��װ || '/' || " & str��Χ & "��λ) As �۸�" & _
                            IIF(strStock <> "", _
                                IIF(InStr(strPriv, "��ʾҩƷ���") = 0, _
                                    ",Decode(Sign(Nvl(X.���,0)),1,'��','') as ���", _
                                    ",Decode(X.���,NULL,NULL,Round(X.���/" & str��Χ & "��װ,5)||" & str��Χ & "��λ) as ���"), _
                                ",Null as ���") & ",Decode(d.������,0,'',1,'������ʹ��',2,'����ʹ��',3,'����ʹ��') as �����ȼ�,D.�ٴ��Թ�ҩ as �ٴ��Թ�ҩID,NULL as ����" & _
                        " From �շѼ�Ŀ B,�շ���ĿĿ¼ A,ҩƷ��� C,ҩƷ���� D,������ĿĿ¼ E,�շ���Ŀ��� F,�շ���Ŀ���� G1" & IIF(gbytҩƷ������ʾ = 2, ",�շ���Ŀ���� G2", "") & _
                          IIF(strStock <> "", ",(" & strStock & ") X", "") & IIF(mint���� <> 0, ",����֧����Ŀ M,����֧������ N", "") & _
                        " Where A.ID=C.ҩƷID And C.ҩ��ID=D.ҩ��ID And D.ҩ��ID=E.ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, Replace(str�������, "A.", "E."), "E.��� = Null"), " E.��� IN ('5','6','7')") & _
                            " And Instr([10],','||Nvl(e.�����Ա�,0)||',')>0 And A.ID=G1.�շ�ϸĿID(+) And G1.����(+)=1 And G1.����(+)=" & IIF(gbytҩƷ������ʾ = 1, 3, 1) & _
                            IIF(gbytҩƷ������ʾ = 2, " And A.ID=G2.�շ�ϸĿID(+) And G2.����(+)=1 And G2.����(+)=3", "") & _
                            " And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, Replace(str�������, "A.", "E."), "E.��� = Null"), " E.��� IN ('5','6','7')") & strCommIF & _
                            " And (E.������� IN([8],3) Or [8]=3 And Nvl(E.�������,0)<>0) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                            " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                            str������� & strҩƷ & Replace(strSub, "A.", "E.") & _
                            IIF(mint���� <> 0, " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[18]", "") & _
                            " And A.ID=B.�շ�ϸĿID And Sysdate Between B.ִ������+0 And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "20", "21", "22") & _
                        " Group by A.���,E.ID,A.ID,F.����,C.����ҩ��,A.����,Nvl(G1.����,A.����)," & IIF(gbytҩƷ������ʾ = 2, "G2.����,", "") & "E.���㵥λ,A.���,A.����,D.ҩƷ����,A.��������," & _
                        IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')'),", "") & "A.˵��,D.����ְ��,A.�Ƿ���," & str��Χ & "��װ," & str��Χ & "��λ" & IIF(strStock <> "", ",X.���", "") & ",d.������,D.�ٴ��Թ�ҩ"
                Else
                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                        " Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                            " F.���� AS ���,C.����ҩ�� as ����,A.����,Nvl(G1.����,A.����) as ����," & IIF(gbytҩƷ������ʾ = 2, "G2.����", "Null") & " as ��Ʒ��," & IIF(mbln����, "Null as ����,", "") & _
                            " E.���㵥λ,A.���,A.����,D.ҩƷ����,Null as ��Ŀ����,A.��������," & _
                            IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������,", "Null as ҽ������,") & "A.˵��,D.����ְ�� as ����ְ��ID,Null as �۸�" & _
                            IIF(strStock <> "", _
                                IIF(InStr(strPriv, "��ʾҩƷ���") = 0, _
                                    ",Decode(Sign(Nvl(X.���,0)),1,'��','') as ���", _
                                    ",Decode(X.���,NULL,NULL,Round(X.���/" & str��Χ & "��װ,5)||" & str��Χ & "��λ) as ���"), _
                                ",Null as ���") & ",Decode(d.������,0,'',1,'������ʹ��',2,'����ʹ��',3,'����ʹ��') as �����ȼ�,D.�ٴ��Թ�ҩ as �ٴ��Թ�ҩID,NULL as ����" & _
                        " From �շ���ĿĿ¼ A,ҩƷ��� C,ҩƷ���� D,������ĿĿ¼ E,�շ���Ŀ��� F,�շ���Ŀ���� G1" & IIF(gbytҩƷ������ʾ = 2, ",�շ���Ŀ���� G2", "") & _
                            IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                            IIF(mint���� <> 0, ",����֧����Ŀ M,����֧������ N", "") & _
                        " Where A.ID=C.ҩƷID And C.ҩ��ID=D.ҩ��ID And D.ҩ��ID=E.ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, Replace(str�������, "A.", "E."), "E.��� = Null"), " E.��� IN ('5','6','7')") & _
                            " And Instr([10],','||Nvl(e.�����Ա�,0)||',')>0 And A.ID=G1.�շ�ϸĿID(+) And G1.����(+)=1 And G1.����(+)=" & IIF(gbytҩƷ������ʾ = 1, 3, 1) & _
                            IIF(gbytҩƷ������ʾ = 2, " And A.ID=G2.�շ�ϸĿID(+) And G2.����(+)=1 And G2.����(+)=3", "") & _
                            " And A.��� " & IIF(mstr���Ʒ��� <> "" And InStr(",5,6,7,", "," & mstr���Ʒ��� & ",") > 0, "='" & mstr���Ʒ��� & "'", " IN ('5','6','7')") & strCommIF & _
                            " And (E.������� IN([8],3) Or [8]=3 And Nvl(E.�������,0)<>0) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                            " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                            str������� & strҩƷ & Replace(strSub, "A.", "E.") & _
                            IIF(mint���� <> 0, " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[18]", "")
                End If
            End If
        End If
    End If
        
        
    '2.��ҩƷ���ĵ�������Ŀ����:���಻��ҩƷ����ʱ���ض�ȡ
    '--------------------------------------------------------------------------------------
    blnLoad = False
    If bln���� Then
        If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,7,", Val(objNode.Tag)) = 0
        End If
    End If
    
    If blnLoad Then
        If mstr���� <> "" Then
            strInput = " And (A.���� Like [5] Or B.���� Like [6] Or B.���� Like [6]) And B.����=[7]"
            If IsNumeric(mstr����) Then
                '1X.����ȫ������ʱֻƥ�����
                If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���� Like [5] And B.����=[7]"
            ElseIf zlCommFun.IsCharAlpha(mstr����) Then
                'X1.����ȫ����ĸʱֻƥ�����
                If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.���� Like [6] And B.����=[7]"
            ElseIf zlCommFun.IsCharChinese(mstr����) Then
                '��������,��ֻƥ������
                strInput = " And B.���� Like [6] And B.����=[7]"
            End If
            
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & IIF(Not mbln����, "Distinct", "") & _
                    " A.��� As ���ID,A.ID as ������ĿID,-Null as �շ�ϸĿID," & _
                    " D.���� As ���,Null as ����,A.����,B.����,Null as ��Ʒ��," & IIF(mbln����, "B.����,", "") & _
                    " A.���㵥λ,A.�걾��λ as ���,Null as ����,Null as ҩƷ����," & str�������� & " As ��Ŀ����," & _
                    " Null as ��������,Null as ҽ������,Null as ˵��,Null as ����ְ��ID,Null as �۸�,Null as ���" & ",Null As �����ȼ�,NULL AS �ٴ��Թ�ҩID,NULL as ����" & _
                    IIF(chkShowCause.value = 1, ",Null as �շѳ���ʱ��,a.����ʱ��,A.վ��,null as ������,null as ���÷������,Null as �������,Null as ��ֵ����,a.�������,Nvl(a.ִ��Ƶ��, 0) as ִ��Ƶ��,Null as ��ҩ��������," & vbNewLine & _
                "              Null as ��ҩ��������,Null as ��ҩ��������,Max(Decode(NVL(e.����id,0),0,0,Decode(instr([17],',' || e.����id || ','),0,-1,e.����id))) as ʹ�ÿ���ID,a.����Ӧ��,Null as �������,Nvl(A.�����Ա�,0) as �����Ա�,b.����,NULL AS δƥ��ԭ��", "") & _
                " From ������Ŀ��� D,������Ŀ���� B,������ĿĿ¼ A" & IIF(chkShowCause.value = 1, ",�������ÿ��� E", "") & _
                " Where A.ID=B.������ĿID And A.���=D.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",4,5,6,7,", "," & mstr���Ʒ��� & ",") = 0, str�������, "A.��� = Null"), " A.��� Not IN ('4','5','6','7')") & strScope & strSub & strInput & _
                IIF(mlng�������� = 1 And strCommIF <> "", Mid(strCommIF, 1, IIF(Len(strCommIF) = 0, 1, Len(strCommIF)) - 1) & " Or A.��� = 'Z')", strCommIF) & _
                IIF(chkShowCause.value = 1, " And E.��ĿID(+)=A.ID Group by a.��� , a.Id ,  d.����, a.����, b.����,  " & IIF(mbln����, "B.����,", "") & _
                " a.���㵥λ,a.�걾��λ ,a.��������,a.����ʱ��,A.վ��,a.�������,a.ִ��Ƶ��,a.����Ӧ��,A.�����Ա�,b.����", "")
        Else
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & _
                    " A.��� As ���ID,A.ID as ������ĿID,-Null as �շ�ϸĿID,D.���� As ���,Null as ����," & _
                    " A.����,A.����,Null as ��Ʒ��," & IIF(mbln����, "Null as ����,", "") & "A.���㵥λ,A.�걾��λ as ���,Null as ����," & _
                    " Null as ҩƷ����," & str�������� & " As ��Ŀ����,Null as ��������,Null as ҽ������,Null as ˵��,Null as ����ְ��ID," & _
                    " Null as �۸�,Null as ���" & ",Null As �����ȼ�,NULL AS �ٴ��Թ�ҩID,NULL as ����" & _
                " From ������Ŀ��� D,������ĿĿ¼ A" & _
                " Where A.���=D.���� And " & IIF(mstr���Ʒ��� <> "", IIF(InStr(",4,5,6,7,", "," & mstr���Ʒ��� & ",") = 0, str�������, "A.��� = Null"), " A.��� Not IN ('4','5','6','7')") & strScope & strSub & _
                IIF(mlng�������� = 1 And strCommIF <> "", Mid(strCommIF, 1, IIF(Len(strCommIF) = 0, 1, Len(strCommIF)) - 1) & " Or A.��� = 'Z')", strCommIF)
        End If
    End If
    
    '3.���Ĳ���:����������������ʱ�Ŷ�ȡ�����������Ϣ��ȡ
    '--------------------------------------------------------------------------------------
    strStock = "" '���Ŀ��,���ϲ���δָ��ʱ,����������¼
    If mint���� <> -1 Then
        blnLoad = False
        If mstr���� = "" Then
            If Val(objNode.Tag) = 7 Then blnLoad = True
        Else
            If Mid(tabClass.SelectedItem.Key, 2) = "4" Then blnLoad = True
        End If
        If (blnLoad Or blnStock And mbln��ʾ���) And mlng���ϲ��� <> 0 And chkShowCause.value <> 1 Then
            strStock = _
                "Select A.ҩƷID,Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
                " Where A.���� = 1 And A.�ⷿID=[16]" & _
                " And (Nvl(A.����, 0) = 0 Or A.Ч�� Is Null Or A.Ч�� > Trunc(Sysdate))" & _
                " Group by A.ҩƷID Having Nvl(Sum(A.��������),0)<>0"
        End If
    End If
    
    '�Ƿ�����п��
    If strStock = "" Then
        str������� = ""
    Else
        str������� = " And A.ID=X.ҩƷID(+)"
    End If
    
    If mstr���ϲ��� <> "" And chkShowCause.value <> 1 Then
        '��ʹ�ð󶨱�������������Ծ�̬��
        If gblnStock Then
            str������� = str������� & " And A.���='4' And Exists(Select 1 From ҩƷ��� Where ҩƷID = c.����ID And ���� = 1 And (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ��������>0 And �ⷿID In (" & mstr���ϲ��� & "))"
        Else
            'ֻ��ʾָ��ҩ����ҩƷ�����ù����õĲ��˿��ң�
            str������� = str������� & " And Exists(Select 1 From �շ�ִ�п��� X Where x.�շ�ϸĿid = c.����ID And A.���='4' And x.ִ�п���id In (" & mstr���ϲ��� & ")" & IIF(mint��Χ <> 3, " And (x.������Դ is NULL Or x.������Դ=[8])", "") & ")"
        End If
    End If
    
    blnLoad = False
    If bln���� Then
        If mlngҩ��ID <> 0 Then
            blnLoad = True
        Else
            If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
                blnLoad = True
            Else
                blnLoad = Val(objNode.Tag) = 7
            End If
        End If
    End If
    
    If blnLoad Then
        If mstr���� <> "" Then
            'ʹ������ƥ��Ĺ���1��ȫ���ֻ�������+��ĸ������10λ�����ϣ�
            If (Not zlCommFun.IsCharChinese(mstr����)) And Len(mstr����) >= 10 Then
                strInput = " And (A.���� Like [5] Or B.���� Like [6] Or B.���� Like [6] Or c.��Ʒ���� Like [5] Or c.�ڲ����� Like [5] ) And B.����=[7] "
                blnBarcode = True
            Else
                strInput = " And (A.���� Like [5] Or B.���� Like [6] Or B.���� Like [6] ) And B.����=[7] "
                blnBarcode = False
            End If
            If IsNumeric(mstr����) Then
                '1X.����ȫ������ʱֻƥ�����
                If Mid(gstrMatchMode, 1, 1) = "1" Then
                    If Len(mstr����) >= 10 Then
                        strInput = " And (A.���� Like [5] Or c.��Ʒ���� Like [5] Or c.�ڲ����� Like [5] ) And B.����=[7] "
                        blnBarcode = True
                    Else
                        strInput = " And (A.���� Like [5] ) And B.����=[7] "
                        blnBarcode = False
                    End If
                End If
            ElseIf zlCommFun.IsCharAlpha(mstr����) Then
                'X1.����ȫ����ĸʱֻƥ�����
                If Mid(gstrMatchMode, 2, 1) = "1" Then
                    If Len(mstr����) >= 10 Then
                        strInput = " And (B.���� Like [6] Or c.��Ʒ���� Like [5] Or c.�ڲ����� Like [5] ) And B.����=[7] "
                        blnBarcode = True
                    Else
                        strInput = " And (B.���� Like [6] ) And B.����=[7] "
                        blnBarcode = False
                    End If
                End If
            ElseIf zlCommFun.IsCharChinese(mstr����) Then
                '��������,��ֻƥ������
                strInput = " And B.���� Like [6] And B.����=[7]"
                blnBarcode = False
            End If
            '���Ƹ��������ƥ����ʾ
            strInside = "Select " & IIF(Not mbln���� Or blnBarcode, "Distinct", "") & _
                " A.ID,A.���,A.����,B.����," & IIF(mbln����, "B.����,", "") & "A.���㵥λ,A.���,A.����,A.��������," & _
                IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������,", "Null as ҽ������,") & "A.˵��,A.�Ƿ���," & IIF(blnBarcode, "C.���� ", "NULL as ����") & _
                IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,b.����", "") & _
                " From �շ���Ŀ���� B,�շ���ĿĿ¼ A " & IIF(blnBarcode, " , ҩƷ��� C ", "") & _
                    IIF(mint���� <> 0, ",����֧����Ŀ M,����֧������ N", "") & _
                " Where A.ID=B.�շ�ϸĿID " & IIF(blnBarcode, " And a.Id = c.ҩƷid ", "") & "  And A.���='4'" & strCommIF & strInput & _
                IIF(mint���� <> 0, " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[18]", "")
            If mbln�۸� Then
                strInside = "Select A.ID,A.���,A.����,A.����," & IIF(mbln����, "A.����,", "") & _
                    " A.���㵥λ,A.���,A.����,A.��������,A.ҽ������,A.˵��,Sum(Decode(A.�Ƿ���,1,NULL,B.�ּ�)) as �۸�,Sum(b.�ּ�) as �ּ�," & IIF(blnBarcode, "a.����", "null as ����") & _
                    IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,A.����", "") & _
                    " From �շѼ�Ŀ B,(" & strInside & ") A" & _
                    " Where A.ID=B.�շ�ϸĿID And Sysdate Between B.ִ������+0 And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "20", "21", "22") & _
                    " Group by A.ID,A.���,A.����,A.����," & IIF(mbln����, "A.����,", "") & _
                    " A.���㵥λ,A.���,A.����,A.��������,A.ҽ������,A.˵��" & IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,A.����", "") & IIF(blnBarcode, ",a.���� ", "")
            ElseIf mstrLike = "" And strStock <> "" Then
                '���������ü�������ʱ(����ƥ��),�����(+)����(ҩƷ���),����ҪGroup Byһ��(���)
                '��Group by ��Distinct ͬʱ����ʱ(Not mbln����)��Oracle��ֻѡ�����Group by
                strInside = strInside & " Group By A.ID,A.���,A.����,B.����," & IIF(mbln����, "B.����,", "") & "A.���㵥λ,A.���,A.����," & _
                " A.��������," & IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')'),", "") & "A.˵��,A.�Ƿ���" & IIF(chkShowCause.value = 1, ",a.����ʱ��,A.վ��,a.�������,b.����", "") & IIF(blnBarcode, ",C.���� ", "")
            End If
            
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & _
                    " A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                    " F.���� AS ���,Null as ����,A.����,A.����,Null as ��Ʒ��," & IIF(mbln����, "A.����,", "") & _
                    " A.���㵥λ,A.���,A.����,Null as ҩƷ����,Null as ��Ŀ����,A.��������,A.ҽ������,A.˵��,Null as ����ְ��ID" & _
                    IIF(mbln�۸�, ",Decode(A.�۸�, Null, decode(C.�ϴ��ۼ�,null,A.�ּ�,C.�ϴ��ۼ�), A.�ּ�)||'/'||A.���㵥λ as �۸�", ",Null as �۸�") & _
                    IIF(strStock <> "", _
                        IIF(InStr(strPriv, "��ʾҩƷ���") = 0, _
                            ",Decode(Sign(Nvl(X.���,0)),1,'��','') as ���", _
                            ",Decode(X.���,NULL,NULL,X.���||A.���㵥λ) as ���"), _
                        ",Null as ���") & ",Null As �����ȼ�,NULL AS �ٴ��Թ�ҩID,A.����" & _
                IIF(chkShowCause.value = 1, ",a.����ʱ�� as �շѳ���ʱ��,e.����ʱ��,A.վ��,NULL AS  ������,a.������� as ���÷������,NULL AS  �������,NULL AS ��ֵ����," & vbNewLine & _
                "              e.�������,Nvl(e.ִ��Ƶ��, 0) as ִ��Ƶ��,Null as ��ҩ��������,Null as ��ҩ��������," & vbNewLine & _
                "              Null as ��ҩ��������,null as ʹ�ÿ���ID, Null as ����Ӧ��,c.������� ,Nvl(e.�����Ա�,0) as �����Ա�,A.����,NULL AS δƥ��ԭ��", "") & _
                " From �������� C,������ĿĿ¼ E,�շ���Ŀ��� F,(" & strInside & ") A" & IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                " Where A.ID=C.����ID And C.����ID=E.ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(mstr���Ʒ��� = "4", Replace(str�������, "A.", "E."), " E.��� = Null"), " E.��� ='4'") & _
                    IIF(chkShowCause.value <> 1, " And Instr([10],','||Nvl(e.�����Ա�,0)||',')>0 And C.�������=0 And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                    " And (E.������� IN([8],3) Or [8]=3 And Nvl(E.�������,0)<>0) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])", "") & _
                    str������� & Replace(strSub, "A.", "E.")
        Else
            If mbln�۸� Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    "Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                        " F.���� AS ���,Null as ����,A.����,A.����,Null as ��Ʒ��," & IIF(mbln����, "Null as ����,", "") & "A.���㵥λ,A.���,A.����,Null as ҩƷ����," & _
                        " Null as ��Ŀ����,A.��������," & IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������,", "Null as ҽ������,") & "A.˵��,Null as ����ְ��ID," & _
                        " Decode(A.�Ƿ���,1,decode(Sum(c.�ϴ��ۼ�),null,Sum(B.�ּ�),sum(c.�ϴ��ۼ�)) || '/' || A.���㵥λ,Sum(B.�ּ�)|| '/' || A.���㵥λ) As �۸�" & _
                        IIF(strStock <> "", _
                            IIF(InStr(strPriv, "��ʾҩƷ���") = 0, _
                                ",Decode(Sign(Nvl(X.���,0)),1,'��','') as ���", _
                                ",Decode(X.���,NULL,NULL,X.���||A.���㵥λ) as ���"), _
                            ",Null as ���") & ",Null As �����ȼ�,NULL AS �ٴ��Թ�ҩID,NULL as ����" & _
                    " From �շѼ�Ŀ B,�շ���ĿĿ¼ A,�������� C,������ĿĿ¼ E,�շ���Ŀ��� F" & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        IIF(mint���� <> 0, ",����֧����Ŀ M,����֧������ N", "") & _
                    " Where A.ID=C.����ID And C.����ID=E.ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(mstr���Ʒ��� = "4", Replace(str�������, "A.", "E."), " E.��� = Null"), " E.��� ='4'") & " And C.�������=0" & _
                        " And A.���='4' And Instr([10],','||Nvl(e.�����Ա�,0)||',')>0" & strCommIF & _
                        " And (E.������� IN([8],3) Or [8]=3 And Nvl(E.�������,0)<>0) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                        " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                        IIF(strStock <> "", IIF(gblnStock, " And A.ID=X.ҩƷID", " And A.ID=X.ҩƷID(+)"), "") & Replace(strSub, "A.", "E.") & _
                        IIF(mint���� <> 0, " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[18]", "") & _
                        " And A.ID=B.�շ�ϸĿID And Sysdate Between B.ִ������+0 And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "20", "21", "22") & _
                    " Group by A.���,E.ID,A.ID,F.����,A.����,A.����,A.���㵥λ,A.���,A.����,A.��������," & IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')'),", "") & _
                    " A.˵��,A.�Ƿ���" & IIF(strStock <> "", ",X.���", "")
            Else
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                        " F.���� AS ���,Null as ����,A.����,A.���� as ����,Null as ��Ʒ��," & IIF(mbln����, "Null as ����,", "") & "A.���㵥λ,A.���,A.����," & _
                        " Null as ҩƷ����,Null as ��Ŀ����,A.��������," & IIF(mint���� <> 0, "n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������,", "Null as ҽ������,") & _
                        " A.˵��,Null as ����ְ��ID,Null as �۸�" & _
                        IIF(strStock <> "", _
                            IIF(InStr(strPriv, "��ʾҩƷ���") = 0, _
                                ",Decode(Sign(Nvl(X.���,0)),1,'��','') as ���", _
                                ",Decode(X.���,NULL,NULL,X.���||A.���㵥λ) as ���"), _
                            ",Null as ���") & ",Null As �����ȼ�,NULL AS �ٴ��Թ�ҩID,NULL as ����" & _
                    " From �շ���ĿĿ¼ A,�������� C,������ĿĿ¼ E,�շ���Ŀ��� F" & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        IIF(mint���� <> 0, ",����֧����Ŀ M,����֧������ N", "") & _
                    " Where A.ID=C.����ID And C.����ID=E.ID And A.���=F.���� And " & IIF(mstr���Ʒ��� <> "", IIF(mstr���Ʒ��� = "4", Replace(str�������, "A.", "E."), " E.��� = Null"), " E.��� ='4'") & " And C.�������=0" & _
                        " And A.���='4' And Instr([10],','||Nvl(e.�����Ա�,0)||',')>0" & strCommIF & _
                        " And (E.������� IN([8],3) Or [8]=3 And Nvl(E.�������,0)<>0) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                        " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                        str������� & Replace(strSub, "A.", "E.") & _
                        IIF(mint���� <> 0, " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[18]", "")
            End If
        End If
    End If
    
    'ͳһ��SQL��ȡ����ֶΣ��ɼ�����[]��Ŀ����������������
    '���ID,������ĿID,�շ�ϸĿID,���,����,����,����,[��Ʒ��],[����],���㵥λ,[���],[����],[ҩƷ����],[��Ŀ����],[��������],[ҽ������],[˵��],����ְ��ID,[�۸�],[���],[�����ȼ�],�ٴ��Թ�ҩid
    '-------------------------------------------------------------------------------------------------------------------------------------------------
    str��ȡ�ֶ� = "a.���ID,a.������ĿID,a.�շ�ϸĿID,a.���,a.����,a.����,a.����,a.��Ʒ��,a.����,a.���㵥λ,a.���,a.���,a.����,a.ҩƷ����," & _
        "a.��Ŀ����,a.��������,a.ҽ������,a.˵��,a.����ְ��ID,a.�۸�,a.�����ȼ�,a.�ٴ��Թ�ҩid,A.����"
        
    If chkShowCause.value = 1 Then
        str��ȡ�ֶ� = str��ȡ�ֶ� & ",a.�շѳ���ʱ��,a.����ʱ��,a.վ��,a.������,a.���÷������,a.�������,a.��ֵ����,a.�������,a.ִ��Ƶ��," & _
            "a.��ҩ��������,a.��ҩ��������,a.��ҩ��������,a.ʹ�ÿ���ID,a.����Ӧ��,a.�������,a.�����Ա�,a.����,a.δƥ��ԭ��"
    End If
    
    If blnOften Then
        '������Ʒ���´�ĳ���ҩƷ(��ҩƷ�࣬avg(R.Ƶ��)=R.Ƶ��,���Ķ���ʱ�򻯺�ҩƷһ��)
        If Not (gblnҩƷ�������ҽ�� Or mint��Ч = 1) Then
            strSQL = "Select /*+ rule*/Rownum as KeyID," & str��ȡ�ֶ� & ",r.Ƶ��ID " & vbNewLine & _
                    "From (" & strSQL & ") A," & vbNewLine & _
                    "(Select ������Ŀid, Avg(Ƶ��) Ƶ��ID From ���Ƹ�����Ŀ Where ��Աid = [11] Group By ������Ŀid) R Where r.������Ŀid = a.������Ŀid" & vbNewLine & _
                    "Order by Ƶ��ID Desc,Decode(���ID,'4','Z',���ID),���,����"
        Else
            strSQL = "Select " & str��ȡ�ֶ� & ",R.Ƶ�� as Ƶ��ID From (" & strSQL & ") A,���Ƹ�����Ŀ R" & _
                    " Where R.������ĿID=A.������ĿID And (A.�շ�ϸĿID is Null Or A.�շ�ϸĿID = R.�շ�ϸĿID) And R.��ԱID=[11]"
                    
            strSQL = "Select /*+ rule*/Rownum as KeyID,A.* From (" & strSQL & ") A Order by Ƶ��ID Desc,Decode(���ID,'4','Z',���ID),���,����"
        End If
    ElseIf mint��Χ = 1 And (mint���� = 0 Or mint���� = 2) And mstr���� <> "" Then
        strSQL = "Select " & str��ȡ�ֶ� & ",R.Ƶ��ID From (" & strSQL & ") A," & _
                " (select ������Ŀid,Avg(ʹ�ô���) as Ƶ��ID from ҽ������ҽ�� Where ��Աid =[11] Group By ������Ŀid) R" & _
                " Where A.������ĿID=R.������ĿID(+)"
        strSQL = "Select /*+ rule*/Rownum as KeyID,A.* From (" & strSQL & ") A Order by nvl(Ƶ��ID,0) Desc,Decode(���ID,'4','Z',���ID),���,����"
    Else
        strSQL = "Select /*+ rule*/Rownum as KeyID," & str��ȡ�ֶ� & " From (" & strSQL & ") A Order by Decode(���ID,'4','Z',���ID),���,����"
    End If
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    If chkShowCause.value = 1 Then
        '�滻����
        strSQL = Replace(strSQL, "B.����=[7]", "B.���� In(1,2)")
        strSQL = Replace(strSQL, "B.���� IN([7],3)", "B.���� IN(1,2,3)")
    End If
    
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int����, str���, lng����ID, mlng����ID, _
        UCase(mstr����) & "%", mstrLike & UCase(mstr����) & "%", mint���� + 1, mint��Χ, IIF(mint��Ч = 0, 2, 1), _
        "," & str�Ա� & ",", UserInfo.ID, lngҩ��ID, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��, mlng���ϲ���, mstr����ID, mint����, mlngҩ��ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "," & mlngҽ������ID & ",")
    
    'δƥ��ԭ�����
    If chkShowCause.value = 1 Then
        If mrsItem.RecordCount > 0 Then
            Set mrsItem = zlDatabase.CopyNewRec(mrsItem)
            mrsItem.MoveFirst
            Do While Not mrsItem.EOF
                If mrsItem!����ʱ�� & "" <> "" And Format(mrsItem!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    mrsItem!δƥ��ԭ�� = Decode(mrsItem!���ID & "", 5, "ҩƷ", 6, "ҩƷ", 7, "ҩƷ", 4, "����", "������Ŀ") & "�Ѿ�ͣ�á�"
                ElseIf mrsItem!�շѳ���ʱ�� & "" <> "" And Format(mrsItem!�շѳ���ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    mrsItem!δƥ��ԭ�� = Decode(mrsItem!���ID & "", 5, "ҩƷ���", 6, "ҩƷ���", 7, "ҩƷ���", 4, "���Ĺ��", "�շ���Ŀ") & "�Ѿ�ͣ�á�"
                ElseIf mrsItem!վ�� & "" <> "" And mrsItem!վ�� & "" <> gstrNodeNo Then
                    mrsItem!δƥ��ԭ�� = "���Ǳ�վ���µ���Ŀ��"
                ElseIf (Val(mrsItem!���ID & "") = 4 Or InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And (gblnҩƷ�������ҽ�� Or mint��Ч = 1)) And Val(mrsItem!���÷������ & "") <> 3 And Val(mrsItem!���÷������ & "") <> mint��Χ And mint��Χ <> 3 Then
                    mrsItem!δƥ��ԭ�� = "���������ƥ�䡣"
                ElseIf Val(mrsItem!������� & "") <> 3 And Val(mrsItem!������� & "") <> mint��Χ And mint��Χ <> 3 Or (mint��Χ = 3 And Val(mrsItem!������� & "") = 0) Then
                    mrsItem!δƥ��ԭ�� = "������Ŀ�������ƥ�䡣"
                ElseIf InStr(",4,5,6,7,", "," & mrsItem!���ID & ",") = 0 And Val(mrsItem!ʹ�ÿ���ID & "") = -1 Then
                    mrsItem!δƥ��ԭ�� = "������Ŀ���ÿ��Ҳ�ƥ�䣬��ǰ���Ҳ����á�"
                ElseIf InStr(",4,5,6,7,", "," & mrsItem!���ID & ",") = 0 And Val(mrsItem!����Ӧ�� & "") <> 1 Then
                    mrsItem!δƥ��ԭ�� = "������Ŀ���ɵ���ʹ�á�"
                ElseIf InStr("," & str�Ա� & ",", "," & mrsItem!�����Ա� & ",") = 0 Then
                    mrsItem!δƥ��ԭ�� = "������Ŀ�Ա�ƥ�䵱ǰ���ˡ�"
                ElseIf mrsItem!ִ��Ƶ�� & "" <> "0" And mrsItem!ִ��Ƶ�� & "" <> IIF(mint��Ч = 0, 2, 1) & "" Then
                    mrsItem!δƥ��ԭ�� = "������Ŀ��ִ��Ƶ��Ϊ" & IIF(mrsItem!ִ��Ƶ�� & "" = "1", "һ����", "������") & ",������Ϊ" & IIF(mrsItem!ִ��Ƶ�� & "" = "1", "������", "������")
                ElseIf InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And InStr(strTsPriv, "�´�����ҩ��") = 0 And mrsItem!������� & "" = "����ҩ" Then
                    mrsItem!δƥ��ԭ�� = "ҩƷΪ������ҩƷ����ǰ�û�û���´�������ҩƷ��Ȩ�ޡ�"
                ElseIf InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And InStr(strTsPriv, "�´ﶾ��ҩ��") = 0 And mrsItem!������� & "" = "����ҩ" Then
                    mrsItem!δƥ��ԭ�� = "ҩƷΪ������ҩƷ����ǰ�û�û���´ﶾ����ҩƷ��Ȩ�ޡ�"
                ElseIf InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And InStr(strTsPriv, "�´ﾫ��ҩ��") = 0 And (mrsItem!������� & "" = "����I��") Then
                    mrsItem!δƥ��ԭ�� = "ҩƷΪ������ҩƷ����ǰ�û�û���´ﾫ����ҩƷ��Ȩ�ޡ�"
                ElseIf InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And InStr(strTsPriv, "�´����ҩ��") = 0 And (mrsItem!��ֵ���� & "" = "����" Or mrsItem!��ֵ���� & "" = "����") Then
                    mrsItem!δƥ��ԭ�� = "ҩƷΪ������ҩƷ����ǰ�û�û���´������ҩƷ��Ȩ�ޡ�"
                ElseIf gblnKSSStrict And mint��Χ = 1 And InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And mrsItem!������ & "" = "3" Then
                    mrsItem!δƥ��ԭ�� = "ҩƷΪ�����࿹��ҩ����ﲻ�����´"
                ElseIf InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And Not (mstr������ҩ�� = "" And mstr���ó�ҩ�� = "" And mstr������ҩ�� = "") And gblnStock Then
                    If mrsItem!���ID = "5" And Val(mrsItem!��ҩ�������� & "") = 0 Or mrsItem!���ID = "6" And Val(mrsItem!��ҩ�������� & "") = 0 Or mrsItem!���ID = "7" And Val(mrsItem!��ҩ�������� & "") = 0 Then
                        mrsItem!δƥ��ԭ�� = "ҩƷ��治�㣬ϵͳ�����˲������´��治���ҩƷ��"
                    End If
                End If
                If mrsItem!δƥ��ԭ�� & "" = "" And InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And Not (mstr������ҩ�� = "" And mstr���ó�ҩ�� = "" And mstr������ҩ�� = "") And Not gblnStock Then
                    If mrsItem!���ID = "5" And Val(mrsItem!��ҩ�������� & "") = 0 Or mrsItem!���ID = "6" And Val(mrsItem!��ҩ�������� & "") = 0 Or mrsItem!���ID = "7" And Val(mrsItem!��ҩ�������� & "") = 0 Then
                        mrsItem!δƥ��ԭ�� = "ҩƷֻ�������õ�ҩ����ִ�С�"
                    End If
                End If
                If mrsItem!δƥ��ԭ�� & "" = "" And mrsItem!���ID = "4" And mrsItem!������� <> 0 Then
                    mrsItem!δƥ��ԭ�� = "�����Ǻ�����ϣ����������´"
                ElseIf mrsItem!δƥ��ԭ�� & "" = "" And mrsItem!���� <> mint���� + 1 And zlCommFun.IsCharAlpha(mstr����) And zlCommFun.IsCharChinese(mrsItem!���� & "") Then
                    mrsItem!δƥ��ԭ�� = "��Ŀ���벻ƥ�䣬��ǰ��ʹ�õ�" & IIF(mint���� + 1 = 1, "ƴ��", "���") & "��"
                End If
                mrsItem.MoveNext
            Loop
            mrsItem.Filter = "δƥ��ԭ�� <> ''"
            If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst
        End If
    End If
    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '����ͳ�Ƴ�����Ŀʱ����Ϊ��0��0��
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        If InStr(",���,�۸�,", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.TextMatrix(0, i) = "ҽ������" Or vsItem.TextMatrix(0, i) = "ҽ������(���õȼ�)" Then
            vsItem.ColWidth(i) = 2000
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    
    '�ָ���˳��:Ӧ����������֮ǰ
    Call RestoreColPosition
    Call RestoreColWidth
    '������:������,�Ա���洦���к�
    Call RestoreColSort
    
    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '�ռ����Ƭ��Ϣ
        If InStr(strClass & ",", "," & mrsItem!���ID & mrsItem!��� & ",") = 0 Then
            strClass = strClass & "," & mrsItem!���ID & mrsItem!���
        End If
        If InStr(",5,6,7,", "," & mrsItem!���ID & ",") > 0 And bln��ʾ��� = True And chkShowCause.value <> 1 Then
            If Val(mrsItem!�ٴ��Թ�ҩID & "") = 0 And mrsItem!��� & "" = "" And (mlng��ҩ�� <> 0 And mrsItem!���ID = "5" Or mlng��ҩ�� <> 0 And mrsItem!���ID = "6" Or mlng��ҩ�� <> 0 And mrsItem!���ID = "7") Then
                '��ʾ�˿�浫��治���ҩƷ���û�ɫ������ʾ�����ų��Թ�ҩ
                vsItem.Cell(flexcpBackColor, i, vsItem.FixedCols, i, vsItem.Cols - 1) = &H8000000F
            End If
        End If
        
        '�ռ���ĿID:ֻ�ռ����2��
        If mstr���� <> "" Then
            If UBound(Split(str������ĿIDs, ",")) < 2 Then
                If InStr(str������ĿIDs & ",", "," & mrsItem!������ĿID & ",") = 0 Then
                    str������ĿIDs = str������ĿIDs & "," & mrsItem!������ĿID
                End If
            End If
            If UBound(Split(str�շ�ϸĿIDs, ",")) < 2 Then
                If Not IsNull(mrsItem!�շ�ϸĿID) Then
                    If InStr(str�շ�ϸĿIDs & ",", "," & mrsItem!�շ�ϸĿID & ",") = 0 Then
                        str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & mrsItem!�շ�ϸĿID
                    End If
                End If
            End If
        End If
        If NVL(mrsItem!�����ȼ�) <> "" Then blnIsHaveKSS = True
        mrsItem.MoveNext
    Next
    
    '�������࿨Ƭ:�ж���ʱ����Ŀ���϶�ʱ
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '��Alt��ݼ������޷�����
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '���ݽ�����������������һЩ����Ҫ����
    For i = 1 To vsItem.Cols - 1
        If vsItem.TextMatrix(0, i) = "��Ʒ��" Then
            If (mstr���� <> "" And gbyt����ҩƷ��ʾ = 0) Or (mstr���� = "" And gbytҩƷ������ʾ <> 2) Then
                vsItem.ColHidden(i) = True '����ʱ����ʾ��ѡ����ֱ�Ӹ��ݲ�����
            ElseIf Not ((InStr(strClass, ",5") > 0 Or InStr(strClass, ",6") > 0) And (gblnҩƷ�������ҽ�� Or mint��Ч = 1)) Then
                vsItem.ColHidden(i) = True '��ҩ������´����Ҫ
            End If
        ElseIf vsItem.TextMatrix(0, i) = "����" Then
            'ѡ������ʽʱΪ���ֶ�
            If mstr���� = "" Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "���" Then
            '��Ʒ���´��������ҩû��(������ĿΪ�걾��λ����)
            If (InStr(strClass, ",5") > 0 Or InStr(strClass, ",6") > 0) _
                And Not (gblnҩƷ�������ҽ�� Or mint��Ч = 1) Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "����" Then
            '��Ʒ���´��������ҩ����ʾ��û��ҩƷ��ʱ����ʾ
            If Not (gblnҩƷ�������ҽ�� Or mint��Ч = 1) Then
                vsItem.ColHidden(i) = True
            ElseIf InStr(strClass, ",5") = 0 And InStr(strClass, ",6") = 0 And InStr(strClass, ",7") = 0 Then
                vsItem.ColHidden(i) = True
            End If
        
        ElseIf InStr(",����,��������,ҽ������,ҽ������(���õȼ�),˵��,�۸�,���,", vsItem.TextMatrix(0, i)) > 0 Then
            '�̶����շ�ϸĿ�����ġ���ҩ�����߰�����´��������ҩ
            If Not (InStr(strClass, ",4") > 0 Or InStr(strClass, ",7") > 0 _
                Or ((InStr(strClass, ",5") > 0 Or InStr(strClass, ",6") > 0) _
                    And (gblnҩƷ�������ҽ�� Or mint��Ч = 1))) Then
                vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "����" And Not gblnShowOrigin Then
                vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "�۸�" Then
                If Not mbln�۸� Then vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "ҽ������" Then
                vsItem.TextMatrix(0, i) = "ҽ������(���õȼ�)"
                If mint���� = 0 Then vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "ҽ������(���õȼ�)" Then
                If mint���� = 0 Then vsItem.ColHidden(i) = True
            End If
            If vsItem.TextMatrix(0, i) = "�۸�" Or vsItem.TextMatrix(0, i) = "���" Then
                For j = vsItem.FixedRows To vsItem.Rows - 1
                    If Mid(vsItem.TextMatrix(j, i), 1, 1) = "." Then vsItem.TextMatrix(j, i) = "0" & vsItem.TextMatrix(j, i)
                Next
            End If
            If chkShowCause.value = 1 Then
                If vsItem.TextMatrix(0, i) = "���" Then
                    vsItem.ColHidden(i) = True
                End If
            End If
        ElseIf vsItem.TextMatrix(0, i) = "ҩƷ����" Then
            'ֻ��ҩƷ����
            If InStr(strClass, ",5") = 0 And InStr(strClass, ",6") = 0 _
                And InStr(strClass, ",7") = 0 And strClass <> "" Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "��Ŀ����" Then
            'ҩƷ�����Ĳ���Ҫ
            strSub = Replace(strClass, "4����", "")
            strSub = Replace(strSub, "5����ҩ", "")
            strSub = Replace(strSub, "6�г�ҩ", "")
            strSub = Replace(strSub, "7�в�ҩ", "")
            strSub = Replace(strSub, ",", "")
            If strSub = "" Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "�����ȼ�" Then
            If Not blnIsHaveKSS Then vsItem.ColHidden(i) = True
        ElseIf InStr(",�շѳ���ʱ��,����ʱ��,վ��,������,���÷������,�������,��ֵ����,�������,ִ��Ƶ��,��ҩ��������,��ҩ��������,��ҩ��������,����Ӧ��,�������,�����Ա�,����,", vsItem.TextMatrix(0, i)) > 0 Then
            '����δƥ��ԭ��ļ�����
            vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "δƥ��ԭ��" Then
            vsItem.ColWidth(i) = 4500
                ElseIf vsItem.TextMatrix(0, i) = "����" Then
            vsItem.ColHidden(i) = True
        End If
    Next
    
    '�к��п��
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    If blnOften Then
        lblInfo.Caption = "��ǰѡ��""" & UserInfo.���� & """�ĸ��˳�����Ŀ������ " & mrsItem.RecordCount & " ����Ŀ"
    Else
        lblInfo.Caption = "��ǰѡ��" & GetTreePath(tvw_s.SelectedItem) & tabClass.SelectedItem.Tag & "������ " & mrsItem.RecordCount & " ����Ŀ"
    End If
    
    vsItem.FrozenCols = 0
    vsItem.Editable = flexEDNone
    vsItem.SheetBorder = vsItem.BackColor
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    zlControl.FormLock 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function

Private Function FillStat(Optional ByVal blnClass As Boolean) As Boolean
'������blnClass=�Ƿ��ؽ����࿨(Ӧ��������Ŀ�ı�ʱ���ؽ�)
    Dim objNode As Node, objItem As ListItem
    Dim strSQL As String, i As Long, j As Long
    Dim arrClass As Variant, strClass As String
    Dim str���� As String, str��� As String
    Dim str�������� As String, strѡ����� As String
    Dim objTab As MSComctlLib.Tab

    Set objNode = tvw_s.SelectedItem '����ΪNothing
    
    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vsItem.Editable = flexEDKbdMouse
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '�����������ֶ�����
    '------------------------------------------------------------------------
    '������Ŀ�Ĳ�������
    str�������� = "Decode(A.���," & _
        "'H',Decode(A.��������,'1','����ȼ�','������')," & _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨','4','��ҩ�÷�','5','��������','6','�ɼ�����','7','��Ѫ����',Null)," & _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','7','����','8','����','9','����','10','��Σ','11','����','12','��¼�����','14','��ǰ',NULL)," & _
        "A.��������)"
    
    'ֻͳ�Ƹ÷�������ĳ�����Ŀ
    If mlng����ID <> 0 Then
        str���� = " And A.����ID IN(" & _
            " Select ID From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With ID=[1] Connect by Prior ID=�ϼ�ID" & _
            " )"
        
        '�����е�����ȷ�����
        If Val(objNode.Tag) = 5 Then
            str��� = str��� & " And A.������� Not IN('4','5','6','7','8','9')"
        Else
            If Val(objNode.Tag) = 1 Then
                strѡ����� = "5"
            ElseIf Val(objNode.Tag) = 2 Then
                strѡ����� = "6"
            ElseIf Val(objNode.Tag) = 3 Then
                strѡ����� = "7"
            ElseIf Val(objNode.Tag) = 4 Then
                strѡ����� = "8"
            ElseIf Val(objNode.Tag) = 6 Then
                strѡ����� = "9"
            ElseIf Val(objNode.Tag) = 7 Then
                strѡ����� = "4"
            End If
            If strѡ����� <> "" Then
                str��� = str��� & " And A.�������=[2]"
            End If
        End If
    End If

    '���Ƭȷ�����
    If tabClass.SelectedItem.Key <> "" Then
        strѡ����� = Mid(tabClass.SelectedItem.Key, 2)
        str��� = str��� & " And A.�������=[2]"
    End If
    
    '��ȡ����:û�����Ƴ���/����Ӧ��,�������,�Ա�Χ,ҩƷȨ��;��ҩȱʡ���ܵ���Ӧ��
    '------------------------------------------------------------------------------
    If InStr(",5,6,7,", strѡ�����) > 0 And strѡ����� <> "" Then
        strSQL = " Select A.�շ�ϸĿID,Count(A.�շ�ϸĿID) As ����" & _
            " From ����ҽ����¼ A" & _
            " Where A.����ʱ��>=[3] And A.����ҽ��=[4] " & str��� & _
            " Group By A.�շ�ϸĿID"
    Else
        strSQL = " Select A.������ĿID,Count(A.������ĿID) As ����" & _
            " From ����ҽ����¼ A" & _
            " Where A.����ʱ��>=[3] And A.����ҽ��=[4] " & str��� & _
            " Group By A.������ĿID"
    End If
    If zlDatabase.DateMoved(dtpDate.value) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    
    If InStr(",5,6,7,", strѡ�����) > 0 And strѡ����� <> "" Then
        'ҩƷ������Ŀ����
        strSQL = _
            " Select A.��� as ���ID,A.ID as ������ĿID,B.�շ�ϸĿID," & _
            " D.���� as ���,A.����,A.����,F.���,A.���㵥λ,C.ҩƷ����,B.����" & _
            " From ������Ŀ��� D,ҩƷ���� C,������ĿĿ¼ A,ҩƷ��� E,(" & strSQL & ") B,�շ���ĿĿ¼ F" & _
            " Where A.���=D.���� And A.ID=C.ҩ��ID And A.ID=E.ҩ��ID And E.ҩƷID=B.�շ�ϸĿID And E.ҩƷID=F.ID" & _
            "   And Not (A.���='E' And Nvl(A.��������,'0')<>'0')" & str���� & _
            "   And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            "   And Nvl(A.�������,0)<>0 And (Nvl(A.����Ӧ��,0)=1 Or A.��� IN('4','7'))"
        If chkAll.value = 0 Then
            strSQL = _
                "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���,���㵥λ,ҩƷ����,-1*���� as ���� From (" & strSQL & ")" & _
                " Group by -1*����,���ID,������ĿID,�շ�ϸĿID,���,����,����,���,���㵥λ,ҩƷ����"
            strSQL = "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���,���㵥λ,ҩƷ����,Abs(����) as ���� From (" & strSQL & ") Where Rownum<=[5]"
        End If
    Else
        '��ҩƷ���ݻ����в���
        strSQL = _
            " Select A.��� as ���ID,A.ID as ������ĿID,Nvl(h.id,f.id) as �շ�ϸĿID," & _
            " D.���� as ���,Nvl(Nvl(h.����,f.����),A.����) as ����,A.����,Nvl(h.���,f.���) as ���,A.���㵥λ,A.�걾��λ," & str�������� & " As ��Ŀ����,B.����" & _
            " From ������Ŀ��� D,������ĿĿ¼ A,(" & strSQL & ") B,�������� G,�շ���ĿĿ¼ H,ҩƷ��� E,�շ���ĿĿ¼ F" & _
            " Where A.���=D.���� And A.ID=B.������ĿID" & str���� & _
            "   And Not (A.���='E' And Nvl(A.��������,'0')<>'0') And a.id=g.����id(+) And g.����id=h.id(+) and a.id=e.ҩ��id(+) and e.ҩƷid=f.id(+)" & _
            "   And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            "   And Nvl(A.�������,0)<>0 And (Nvl(A.����Ӧ��,0)=1 Or A.��� IN('4','7'))"
        If chkAll.value = 0 Then
            strSQL = _
                "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���,���㵥λ,�걾��λ,��Ŀ����,-1*���� as ���� From (" & strSQL & ")" & _
                " Group by -1*����,���ID,������ĿID,�շ�ϸĿID,���,����,����,���,���㵥λ,�걾��λ,��Ŀ����"
            strSQL = "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���,���㵥λ,�걾��λ,��Ŀ����,Abs(����) as ���� From (" & strSQL & ") Where Rownum<=[5]"
        End If
    End If
    strSQL = "Select Rownum as KeyID,Null as ѡ��,A.* From (" & strSQL & ") A Order by ���� Desc,����"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, strѡ�����, CDate(Format(dtpDate.value, "yyyy-MM-dd")), UserInfo.����, Val(txtCount.Text))
    
    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '����ͳ�Ƴ�����Ŀʱ����Ϊ��0��0��
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        vsItem.ColAlignment(i) = 1
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    
    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '�ռ����Ƭ��Ϣ
        If InStr(strClass & ",", "," & mrsItem!���ID & mrsItem!��� & ",") = 0 Then
            strClass = strClass & "," & mrsItem!���ID & mrsItem!���
        End If
        mrsItem.MoveNext
    Next
    
    '�������࿨Ƭ:�ж���ʱ����Ŀ���϶�ʱ
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '��Alt��ݼ������޷�����
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '�к��п��
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    lblInfo.Caption = "��ǰѡ��""" & UserInfo.���� & """�ĸ��˳�����Ŀ������ " & mrsItem.RecordCount & " ����Ŀ"
    
    vsItem.TextMatrix(0, 2) = ""
    vsItem.ColWidth(2) = 300
    vsItem.ColDataType(2) = flexDTBoolean
    vsItem.FrozenCols = 2
    vsItem.Editable = flexEDKbdMouse
    vsItem.SheetBorder = vbBlack
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillStat = True
    Exit Function
errH:
    zlControl.FormLock 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function

Private Sub SetFontSize(ByVal bytSize As Byte)
'���ܣ����н��������ͳһ����
'������bytSize  0-9�����壬1-12������
    Call zlControl.SetPubFontSize(Me, bytSize)
    If bytSize = 1 Then vsItem.RowHeightMin = 300: tvw_s.Font.Size = IIF(bytSize = 0, 9, 12)
End Sub
