VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmOutAdviceSend 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҽ������"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmOutAdviceSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9615
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   540
      Width           =   9435
      Begin VB.ComboBox cboDrugType 
         Height          =   300
         Left            =   6630
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   15
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   900
         MaxLength       =   1000
         TabIndex        =   2
         Top             =   360
         Width           =   8415
      End
      Begin VB.Label lblDrugType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ������"
         Height          =   180
         Left            =   5865
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ˣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ժҪ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6615
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   5
      Top             =   6210
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6150
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOutAdviceSend.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   714
      BandCount       =   2
      FixedOrder      =   -1  'True
      BandBorders     =   0   'False
      _CBWidth        =   9615
      _CBHeight       =   405
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinWidth1       =   3300
      MinHeight1      =   345
      Width1          =   3300
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "tbrSys"
      MinHeight2      =   345
      Width2          =   9195
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   345
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   609
         ButtonWidth     =   2619
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����Ϊ�շѵ�"
               Key             =   "����Ϊ�շѵ�"
               Description     =   "����Ϊ�շѵ�"
               Object.ToolTipText     =   "����Ϊ�շѵ�(Ctrl+1)"
               Object.Tag             =   "����Ϊ�շѵ�"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����Ϊ���ʵ�"
               Key             =   "����Ϊ���ʵ�"
               Description     =   "����Ϊ���ʵ�"
               Object.ToolTipText     =   "����Ϊ���ʵ�(Ctrl+2)"
               Object.Tag             =   "����Ϊ���ʵ�"
               ImageKey        =   "����"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrSys 
         Height          =   345
         Left            =   3525
         TabIndex        =   8
         Top             =   30
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Ctrl+A)"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "ȫѡ"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Ctrl+R)"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ȫ��"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�(ALT+X)"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
         Begin VB.CheckBox chk�Ӱ�Ӽ� 
            Caption         =   "ִ�мӰ�Ӽ�(&V)"
            Height          =   195
            Left            =   4350
            TabIndex        =   3
            Top             =   150
            Width           =   1650
         End
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   9
      Top             =   4605
      Width           =   9495
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3285
      Left            =   0
      TabIndex        =   0
      Top             =   1245
      Width           =   9540
      _cx             =   1981497788
      _cy             =   1981486754
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
      BackColorSel    =   16771802
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOutAdviceSend.frx":0E1E
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
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1470
      Left            =   0
      TabIndex        =   1
      Top             =   4665
      Width           =   9525
      _cx             =   1981497761
      _cy             =   1981483553
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
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
Attribute VB_Name = "frmOutAdviceSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean) '�༭�������

Private mMainPrivs As String 'IN
Private mlng����ID As Long 'IN
Private mstr�Һŵ� As String 'IN
Private mstrǰ��IDs As String 'IN
Private mblnAuto As Boolean 'IN
Private mblnSend As Boolean 'OUT:�Ƿ�ɹ����͹���
Private mint���� As Integer '���ó���: 2-ҽ��վ

Private mlng�Һ�ID As Long
Private mlng�������ID As Long
Private mint���� As Integer
Private mint����ģʽ As Integer
Private mstr��λ As String '�����Ƿ��Լ��λ���˼�����ĵ�λ����
Private mlngҽ������ID As Long 'ҽ������ID

Private mint�������� As Integer '0-����Ϊ�շѵ�,1-����Ϊ���ʵ�,2-�ֹ�ѡ��
Private mblnһ����ҩ����Ϊһ�� As Boolean 'һ����ҩ��ҩƷ��Ӧ�Ĵ����㲻ͬʱ���Ƿ��Է���Ϊһ�ŵ���
Private mbln��λ���� As Boolean '�Ƿ����Լ��λ���˷���Ϊ���ʵ�
Private mstr���������� As String 'ִ�п�����ͬʱ������Ϊͬһ���ݵ�ҽ�����
Private mblnNOCtrl As Boolean '��ͬ��ϵ�ҽ���ֱ��������
Private mblnStartTimeDef As Boolean '��ʼʱ�䲻��ͬһ��ķֱ��������

Private mintSendNo As Integer
Private mstrLike As String
Private mint���� As Integer
Private mblnAutoExe As Boolean
Private mbytSize As Byte

Private mlngNOSequence As Long
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀⷿ�ĳ����鷽ʽ
Private mrsPati As ADODB.Recordset '����������Ϣ
Private mrsPrice As ADODB.Recordset '�����Ƽ۹�ϵ
Private mrsBill As ADODB.Recordset
Private mrsRXKey As ADODB.Recordset
Private mstr���� As String
Private mstr����� As String
Private mint���� As Integer
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln���鵥���������� As Boolean  '����ҽ��������������
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng���ϲ��� As Long
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�

Private mblnFirst As Boolean
Private mblnUnload As Boolean
Private mbln���֧�� As Boolean '�жϵ�ǰ�����Ƿ������֧�����˲������������֧������
Private mstr֧����ʽ As String  '�������֧���ӿ�ʱ�Ĵ������ "1"--���֧����""--�����֧��
Private mlngCardType As Long    '���֧���Ŀ����ID
Private mblnʹ��Ԥ�� As Boolean '���������֧������ʹ��Ԥ����
Private mbln������ҩ As Boolean  'Ƥ��������ҩ �����������ô˲��������ж�Ƥ�Խ��������Ҫ��дƤ��������ҩ˵��
Private mstrAdDrugIDs As String '���һ���������˵����ҩƷ��ҽ��ID����
Private mblnԤԼ���� As Boolean '�жϱ��η���ҽ��ʱ�Ƿ���Ҫ����ԤԼ���ķ���
Private mlngԤ��Ժҽ��ID As Long  '������ԤԼ�к��¼�ķ���ԤԼ��Ժҽ��

'--------------------------------------------------
Private Const COL_ѡ�� = 0
Private Const COL_Ӥ�� = 1
Private Const col_ҽ������ = 2
Private Const COL_���� = 3
Private Const COL_������λ = 4
Private Const COL_���� = 5
Private Const COL_������λ = 6
Private Const COL_��� = 7
Private Const COL_Ƶ�� = 8
Private Const COL_�÷� = 9
Private Const COL_ҽ������ = 10 'Data���ڴ��ժҪ(ҽ��)
Private Const COL_ִ��ʱ�� = 11
Private Const COL_ִ�п��� = 12
Private Const COL_ִ������ = 13
Private Const COL_ID = 14 '������
Private Const COL_���ID = 15
Private Const COL_���˿���ID = 16
Private Const COL_��������ID = 17
Private Const COL_����ҽ�� = 18
Private Const COL_������� = 19
Private Const COL_������ĿID = 20
Private Const COL_�걾��λ = 21
Private Const COL_��鷽�� = 22
Private Const COL_ִ�б�� = 23
Private Const COL_�Ƽ����� = 24
Private Const COL_ִ������ID = 25
Private Const COL_ִ�п���ID = 26
Private Const COL_�շ�ϸĿID = 27
Private Const COL_Ƶ�ʴ��� = 28
Private Const COL_Ƶ�ʼ�� = 29
Private Const COL_�����λ = 30
Private Const COL_����ϵ�� = 31
Private Const COL_�����װ = 32
Private Const COL_���ﵥλ = 33
Private Const COL_�ɷ���� = 34
    Private Const COL_�������� = 34
Private Const COL_��� = 35
Private Const COL_���� = 36
Private Const COL_�ֽ�ʱ�� = 37
Private Const COL_�״�ʱ�� = 38
Private Const COL_ĩ��ʱ�� = 39
Private Const COL_ǰ��ID = 40
Private Const COL_ǩ��ID = 41
Private Const COL_�Թܱ��� = 42
Private Const COL_�������� = 43
Private Const COL_������־ = 44
Private Const COL_��Ѽ��� = 45
Private Const COL_���㷽ʽ = 46
Private Const COL_��ʼʱ�� = 47
Private Const COL_ִ�а��� = 48
Private Const COL_ִ�з��� = 49
Private Const COL_������� = 50
Private Const COL_��ҩ���� = 51

'-------------------------------------------------
Private Const COLP_�к� = 0
Private Const COLP_�շ�ϸĿID = 1
Private Const COLP_�̶� = 2
Private Const COLP_��� = 3
Private Const COLP_�Ƽ�ҽ�� = 4 '�ɼ���
Private Const COLP_��� = 5
Private Const COLP_�շ���Ŀ = 6
Private Const COLP_�Ƽ����� = 7
Private Const COLP_���� = 8
Private Const COLP_��λ = 9
Private Const COLP_���� = 10
Private Const COLP_Ӧ�ս�� = 11
Private Const COLP_ʵ�ս�� = 12
Private Const COLP_ִ�п��� = 13
Private Const COLP_�������� = 14
Private Const COLP_���� = 15
Private Const COLP_�շѷ�ʽ = 16
Private Const COLP_�շ���� = 17 '������
Private Const COLP_ִ�п���ID = 18
Private Const COLP_�������� = 19
Private Const COLP_�������� = 20

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, ByVal lng����ID As Long, ByVal str�Һŵ� As String, _
    ByVal strǰ��IDs As String, Optional ByVal blnAuto As Boolean, Optional ByVal lngҽ������ID As Long, _
    Optional ByVal int���� As Integer, Optional ByRef objMip As Object) As Boolean
'���ܣ�����ҽ��
'������
'       blnAuto=����ҽ��վ����ҽ��ʱ�Զ���ɷ��Ͳ��������ز���ȷ���˷��͵��ݵ����ͣ�
'       lngҽ������ID=ҽ������վ��������ҽ��ʱ������ѡ���ҽ������
'       int����=2 ҽ������վ����ҽ��
'       strǰ��IDs ҽ��վ�´�ҽ����ǰ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mMainPrivs = MainPrivs
    mlng����ID = lng����ID
    mstr�Һŵ� = str�Һŵ�
    mstrǰ��IDs = strǰ��IDs
    mlngҽ������ID = lngҽ������ID
    mint���� = int����
    mblnAuto = blnAuto
    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng����ID, 0, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
        
    'ȡ�Һ�ID(�������)
    mlng�Һ�ID = 0: mlng�������ID = 0
    strSQL = "Select ID,Nvl(Nvl(�������ID,ת�����ID),ִ�в���ID) as ����ID,�����,���� as ��������,���� From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowMe", mstr�Һŵ�)
    If Not rsTmp.EOF Then
        mlng�Һ�ID = Val(rsTmp!ID & "")
        mlng�������ID = Val(rsTmp!����ID & "")
        mstr���� = rsTmp!�������� & ""
        mstr����� = rsTmp!����� & ""
        mint���� = Val(rsTmp!���� & "")
    End If
    
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    ShowMe = mblnSend
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboDrugType_Click()
    If Val(cboDrugType.Tag) <> cboDrugType.ListIndex Then
        '���¶�ȡ�����嵥
        Me.Refresh
        Call LoadAdviceSend
        vsAdvice.SetFocus
        cboDrugType.Tag = cboDrugType.ListIndex
    End If
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub chk�Ӱ�Ӽ�_Click()
    gbln�Ӱ�Ӽ� = chk�Ӱ�Ӽ�.value = 1
    '���¶�ȡ�����嵥
    Me.Refresh
    Call LoadAdviceSend
    vsAdvice.SetFocus
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        
        '��ȡ�����嵥
        Me.Refresh
        If Not LoadAdviceSend Then Unload Me: Exit Sub
        
        '�Զ���ʼ����:�̶���������ʱ
        If mblnAuto Then
            If tbrMain.Buttons("����Ϊ�շѵ�").Enabled And tbrMain.Buttons("����Ϊ���ʵ�").Enabled _
                Or Not tbrMain.Buttons("����Ϊ�շѵ�").Enabled And Not tbrMain.Buttons("����Ϊ���ʵ�").Enabled Then
                mblnAuto = False
            End If
        End If
        Call tbrSys_ButtonClick(tbrSys.Buttons("ȫѡ"))
        If mblnAuto Then
            mblnUnload = True
            If tbrMain.Buttons("����Ϊ�շѵ�").Enabled Then
                Call tbrMain_ButtonClick(tbrMain.Buttons("����Ϊ�շѵ�"))
            ElseIf tbrMain.Buttons("����Ϊ���ʵ�").Enabled Then
                Call tbrMain_ButtonClick(tbrMain.Buttons("����Ϊ���ʵ�"))
            End If
            If mblnUnload Then Unload Me: Exit Sub '�����ظ�Unload
        End If
    End If
End Sub

Private Function GetPatiInfo() As Boolean
'���ܣ���ȡ������Ϣ
    Dim strSQL As String
    
    On Error GoTo errH
 
    'ִ�в���(�ű����)�����˿���
    strSQL = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ���� = 1 And ����ID=[1]"
    strSQL = "Select Decode(A.��ͬ��λID,NULL,NULL,Nvl(A.������λ,D.����)) as ��λ,Nvl(c.����,A.����) ����,Nvl(c.�Ա�,A.�Ա�) �Ա� ,Nvl(c.����,A.����) ���� ,A.�����," & _
        " A.�ѱ�,A.����,A.����ģʽ,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������,Nvl(B.Ԥ�����,0)-Nvl(B.�������,0) as ʣ���,a.��ͥ�绰 as PhoneNO,a.סԺ�� as InPatNo," & _
        "To_Char(A.��������,'YYYY-MM-DD HH24:MI:SS') as Birthdate,a.��ͥ��ַ as Address" & _
        " From ������Ϣ A,(" & strSQL & ") B,���˹Һż�¼ C,��Լ��λ D" & _
        " Where A.����ID=B.����ID(+) And A.��ͬ��λID=D.ID(+)" & _
        " And A.����id = C.����id(+) And A.����� = C.�����(+) " & _
        " And A.����ID=[1] And c.id(+)=[2]"
    'Set mrsPati = New ADODB.Recordset
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    mstr��λ = NVL(mrsPati!��λ)
    mint���� = NVL(mrsPati!����, 0)
    mint����ģʽ = NVL(mrsPati!����ģʽ, 0)
    lblPati.Caption = _
        "������" & mrsPati!���� & "���Ա�" & NVL(mrsPati!�Ա�) & "�����䣺" & NVL(mrsPati!����) & "���ѱ�" & NVL(mrsPati!�ѱ�)
    
    'ҽ��ָ����������ʱ���뷢��ժҪ���ͻ
    If mint���� <> 0 Then
        If gclsInsure.GetCapability(supportҽ��ȷ����������, mlng����ID, mint����) Then
            txtNote.Text = "": txtNote.Enabled = False
            fraInfo.Height = txtNote.Top
        End If
    End If
    
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("����"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("�˳�"))
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("ȫѡ"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("ȫ��"))
    ElseIf KeyCode = vbKey1 And Shift = vbCtrlMask Then
        If tbrMain.Buttons("����Ϊ�շѵ�").Enabled And tbrMain.Buttons("����Ϊ�շѵ�").Visible Then
            Call tbrMain_ButtonClick(tbrMain.Buttons("����Ϊ�շѵ�"))
        End If
    ElseIf KeyCode = vbKey2 And Shift = vbCtrlMask Then
        If tbrMain.Buttons("����Ϊ���ʵ�").Enabled And tbrMain.Buttons("����Ϊ���ʵ�").Visible Then
            Call tbrMain_ButtonClick(tbrMain.Buttons("����Ϊ���ʵ�"))
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strPrivs As String
    Dim strMsg As String
    Dim bln�޿�֧�� As Boolean
    
    If Not OutPatiFeeUsable(mlng����ID) Then Unload Me: Exit Sub
    
    '���ù�����ťͼ��
    Set tbrMain.HotImageList = frmIcons.imgColor
    Set tbrMain.ImageList = frmIcons.imgGray
    Set tbrSys.HotImageList = frmIcons.imgColor
    Set tbrSys.ImageList = frmIcons.imgGray
    tbrSys.Buttons("ȫѡ").Image = "ȫѡ"
    tbrSys.Buttons("ȫ��").Image = "ȫ��"
    tbrMain.Buttons("����Ϊ�շѵ�").Image = "ִ��"
    tbrMain.Buttons("����Ϊ���ʵ�").Image = "ִ��"
    tbrSys.Buttons("����").Image = "����"
    tbrSys.Buttons("�˳�").Image = "�˳�"
    tbrSys.ButtonHeight = 500
    tbrMain.ButtonHeight = 500
    
    Call InitAdviceTable
    Call InitPriceTable
    strPrivs = GetInsidePrivs(p����ҽ���´�)
    mbln���֧�� = False
    mstr֧����ʽ = ""
    '�ж�Ȩ��
    bln�޿�֧�� = InStr(strPrivs, ";����޿�֧��;") > 0
    
    If bln�޿�֧�� Then
        '�ж��Ƿ���Ҫ���֧��
        mbln���֧�� = Val(zlDatabase.GetPara("����ҽ�����ͺ��������֧��", glngSys, p����ҽ���´�)) = 1
    End If
        
    mlngCardType = 0
    
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    mbln���鵥���������� = Val(zlDatabase.GetPara("����ҽ��������������", glngSys, p����ҽ���´�, "0")) = 1
    mblnʹ��Ԥ�� = Val(zlDatabase.GetPara("���֧������ʹ��Ԥ����", glngSys, p����ҽ���´�, "1"))
    '��������
    mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")
    mblnSend = False
    mblnFirst = True
    
    Call SetFontSize(mbytSize)
    Call RestoreWinState(Me, App.ProductName)
    
    '���͵��ݺ�
    mintSendNo = 0
    If mstrǰ��IDs = "" Then
        mintSendNo = Val(zlDatabase.GetPara("���͵��ݺŹ���", glngSys, p����ҽ���´�)) '0-���,1-����,2-����
    End If
    mstr���������� = zlDatabase.GetPara("����Ϊͬһ���ݵ�ҽ�����", glngSys, p����ҽ���´�)

    '����Լ��λ���˷���Ϊ���ʵ�
    mbln��λ���� = Val(zlDatabase.GetPara("��λ����", glngSys, p����ҽ���´�)) <> 0
    '��ʾ������Ϣ(��ȡ��ҩ��λ��Ϣ)
    If Not GetPatiInfo Then Unload Me: Exit Sub
    
    '��ͬ��ϵ�ҽ���ֱ��������
    mblnNOCtrl = Val(zlDatabase.GetPara("��ͬ��ϵ�ҽ���ֱ��������", glngSys, p����ҽ���´�, 0)) = 1
    
    '��ʼʱ�䲻��ͬһ��ķֱ��������
    mblnStartTimeDef = Val(zlDatabase.GetPara("��ʼʱ�䲻��ͬһ��ķֱ��������", glngSys, p����ҽ���´�, 0)) = 1
    
    '����ѡ��:0-����Ϊ�շѵ�,1-����Ϊ���ʵ�,2-�ֹ�ѡ��
    mint�������� = Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�))
    mblnһ����ҩ����Ϊһ�� = Val(zlDatabase.GetPara("һ����ҩ����Ϊһ��", glngSys, p����ҽ���´�, 1)) = 1
    
    If mint����ģʽ = 1 Then
        tbrMain.Buttons("����Ϊ�շѵ�").Visible = False
        tbrMain.Buttons("����Ϊ�շѵ�").Enabled = False
        cbr.Bands(1).MinWidth = cbr.Bands(1).MinWidth / 2
        cbr.Bands(1).Width = cbr.Bands(1).MinWidth
        
        If InStr(strPrivs, ";����Ϊ���ʵ�;") = 0 Then
            strMsg = "�ò��˲��õ��������ƺ����ģʽ��ֻ�ܷ���Ϊ���ʵ���������û�з���Ϊ���ʵ���Ȩ�ޡ�"
        End If
    Else
        If mint�������� = 0 Or InStr(strPrivs, ";����Ϊ���ʵ�;") = 0 Or (mbln��λ���� And mstr��λ = "") Then
            '������Ϊ��Լ����ʱҪ��ʾ"���ʵ�"��ť
            If Not (mint�������� = 0 And InStr(strPrivs, ";����Ϊ���ʵ�;") > 0 And mbln��λ���� And mstr��λ <> "") Then
                tbrMain.Buttons("����Ϊ���ʵ�").Visible = False
                tbrMain.Buttons("����Ϊ���ʵ�").Enabled = False
                cbr.Bands(1).MinWidth = cbr.Bands(1).MinWidth / 2
                cbr.Bands(1).Width = cbr.Bands(1).MinWidth
            End If
        End If
        If mint�������� = 1 Or InStr(strPrivs, ";����Ϊ�շѵ�;") = 0 Then
            tbrMain.Buttons("����Ϊ�շѵ�").Visible = False
            tbrMain.Buttons("����Ϊ�շѵ�").Enabled = False
            cbr.Bands(1).MinWidth = cbr.Bands(1).MinWidth / 2
            cbr.Bands(1).Width = cbr.Bands(1).MinWidth
        End If
    
        If mint�������� = 0 And InStr(strPrivs, ";����Ϊ�շѵ�;") = 0 Then
            strMsg = "��û�з���Ϊ�շѵ���Ȩ�ޡ�"
        ElseIf mint�������� = 1 Then
            If InStr(strPrivs, ";����Ϊ���ʵ�;") = 0 Then
                strMsg = "��û�з���Ϊ���ʵ���Ȩ�ޡ�"
                If mbln��λ���� And mstr��λ <> "" Then strMsg = "��ǰ�����Ǻ�Լ��λ���ˣ����뷢��ΪΪ���ʵ���������û�з���Ϊ���ʵ���Ȩ�ޡ�"
            Else
                If mbln��λ���� And mstr��λ = "" Then strMsg = "��ǰ���˲��Ǻ�Լ��λ���ˣ����ܷ���Ϊ���ʵ���"
            End If
        ElseIf mint�������� = 2 Then
            If InStr(strPrivs, ";����Ϊ�շѵ�;") = 0 And InStr(strPrivs, ";����Ϊ���ʵ�;") = 0 Then
                strMsg = "��û�з���Ϊ�շѵ��ͷ���Ϊ���ʵ���Ȩ�ޡ�"
            ElseIf InStr(strPrivs, ";����Ϊ�շѵ�;") = 0 Then
                If mbln��λ���� And mstr��λ = "" Then strMsg = "��ǰ���˲��Ǻ�Լ��λ���ˣ����뷢��Ϊ�շѵ���������û�з���Ϊ�շѵ���Ȩ�ޡ�"
            End If
        End If
    End If
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, Me.Caption
        On Error Resume Next
        Unload Me: Exit Sub
        err.Clear: On Error GoTo 0
    End If
    
    mlng��ҩ�� = Val(zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�, , , , , mlng�������ID))
    mlng��ҩ�� = Val(zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�, , , , , mlng�������ID))
    mlng��ҩ�� = Val(zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, p����ҽ���´�, , , , , mlng�������ID))
    mlng���ϲ��� = Val(zlDatabase.GetPara("����ȱʡ���ϲ���", glngSys, p����ҽ���´�, , , , , mlng�������ID))
    
    cboDrugType.AddItem "0-ȫ��"
    cboDrugType.AddItem "1-��Ʒ��"
    cboDrugType.AddItem "2-����;���I��"
    cboDrugType.AddItem "3-����(��1��2��)"
    cboDrugType.Visible = gbln����ҩƷ�ֿ�����
    lblDrugType.Visible = gbln����ҩƷ�ֿ�����
    Call Cbo.SetIndex(cboDrugType.hwnd, 0)
    cboDrugType.Tag = "0"
    
    '����ִ���Զ����
    mblnAutoExe = Val(zlDatabase.GetPara("���ﱾ���Զ�ִ��", glngSys, p����ҽ���´�)) <> 0
    
    mbln������ҩ = Val(zlDatabase.GetPara("Ƥ��������ҩ", glngSys, p����ҽ���´�)) <> 0
    
    If gobjSquareCard Is Nothing Then
        On Error Resume Next
        Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If gobjSquareCard.zlInitComponents(Me, p����ҽ���´�, glngSys, gstrDBUser, gcnOracle, False) = False Then Set gobjSquareCard = Nothing
        err.Clear: On Error GoTo 0
    End If
    
    If gobjSquareCard Is Nothing Then
        mbln���֧�� = False
    End If
    
    '�����ⷿҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '��ȡ��̬�ѱ�,���͹رպ����
    gstr��̬�ѱ� = Load��̬�ѱ�(mlng�������ID)
        
    Call ShowPatiMoney
End Sub

Private Function TheStockCheck(ByVal lng�ⷿID As Long, ByVal str��� As String) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    On Error Resume Next
    If InStr(",5,6,7,", str���) > 0 Then
        intStyle = mcolStock1("_" & lng�ⷿID)
    ElseIf str��� = "4" Then
        intStyle = mcolStock2("_" & lng�ⷿID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Sub ShowPatiMoney()
    Dim rsTmp As ADODB.Recordset
    '��ʾ����Ԥ�����
    Set rsTmp = GetMoneyInfo(mlng����ID, 0)
    If Not rsTmp Is Nothing Then
        If NVL(rsTmp!Ԥ�����, 0) - NVL(rsTmp!�������, 0) <> 0 Then
            stbThis.Panels(4).Text = "Ԥ��:" & Format(NVL(rsTmp!Ԥ�����, 0) - NVL(rsTmp!�������, 0), "0.00")
            stbThis.Panels(4).Visible = True
        Else
            stbThis.Panels(4).Visible = False
        End If
    End If
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    txtNote.Left = lblNote.Left + lblNote.Width + 30
    txtNote.Width = fraInfo.Width - txtNote.Left - 150
    fraInfo.Height = txtNote.Height + txtNote.Top + 60
    
    cboDrugType.Left = fraInfo.Width - cboDrugType.Width - 150
    lblDrugType.Left = cboDrugType.Left - lblDrugType.Width - 30
    lblDrugType.Top = lblPati.Top
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - vsPrice.Height - fraUD.Height - cbr.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = 0
    fraUD.Width = Me.ScaleWidth
    
    vsPrice.Left = 0
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = Me.ScaleWidth
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    chk�Ӱ�Ӽ�.Top = tbrSys.Top + (tbrSys.Height - chk�Ӱ�Ӽ�.Height) / 2 - 15
    If Me.ScaleWidth - tbrSys.Left - chk�Ӱ�Ӽ�.Width - 100 < 4300 Then
        chk�Ӱ�Ӽ�.Left = 4300
    Else
        chk�Ӱ�Ӽ�.Left = Me.ScaleWidth - tbrSys.Left - chk�Ӱ�Ӽ�.Width - 100
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '�ͷ�˽�м�IN����
    mblnAuto = False
    mMainPrivs = ""
    mstr�Һŵ� = ""
    mlng����ID = 0
    mlngCardType = 0
    Set mrsPati = Nothing
    Set mrsPrice = Nothing
    Set mrsBill = Nothing
    Set mrsRXKey = Nothing
    Set mcolStock1 = Nothing
    Set mcolStock2 = Nothing
    
    gbln�Ӱ�Ӽ� = False
    gstr��̬�ѱ� = ""
    Set mclsMipModule = Nothing
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Top = vsPrice.Top + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Function ExpendSendClear(ByVal strNO As String, Optional ByVal blnShowCell As Boolean) As String
'���ܣ�������˹Һŵ��ѳ����Һ���Ч���������Զ���ѡ���ٴ��´��ҽ��
'������strNO=�Һ�NO
'      blnShowCell=�Ƿ�λ��ʾ�������
'���أ��������ѡ��Ҫ���͵��ٴ�ҽ�����򷵻���ʾ��Ϣ
    Dim strMsg As String, i As Long
    
    If BillExpend(strNO) Then
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If Val(.TextMatrix(i, COL_ID)) <> 0 And Val(.TextMatrix(i, COL_ǰ��ID)) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                        Call RowSelectSame(i, COL_ѡ��)
                        If strMsg = "" Then
                            strMsg = "�ò��˹Һ��ѳ�����Ч�������ٴ��´��ҽ�������ٷ���Ϊ�շѵ���"
                            If blnShowCell Then Call .ShowCell(i, COL_ѡ��)
                        End If
                    End If
                End If
            Next
        End With
        ExpendSendClear = strMsg
    End If
End Function

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim rsDiag As ADODB.Recordset
    Dim lng���ͺ� As Long, strMsg As String
    Dim bln���� As Boolean, i As Long
    Dim lngCount As Long, str��� As String
    Dim blnDiagnose As Boolean, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim bln�ϰల�� As Boolean, bytDay As Long, blnFirst As Boolean '�Ƿ��һ���ж��ϰల��
    Dim str�������� As String
    Dim lngָ������ӡ As Long
    
    On Error GoTo errH
    
    blnFirst = True
    With vsAdvice
        '���ҽ��������ϵ���д
        str��� = zlDatabase.GetPara("Ҫ�������������", glngSys, p����ҽ���´�)
        If str��� <> "" Then
            strSQL = "Select B.ҽ��ID From ������ϼ�¼ A,�������ҽ�� B" & _
                " Where A.����ID=[1] And A.��ҳID=[2] And A.��¼��Դ=3 And A.������� In(1,11) And A.ID=B.���ID"
            Set rsDiag = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
        End If
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                lngCount = lngCount + 1

                blnDiagnose = False
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    If InStr(str���, "5") > 0 Then
                        blnDiagnose = True
                    End If
                Else
                    If InStr(str���, .TextMatrix(i, COL_�������)) > 0 Then
                        blnDiagnose = True
                    End If
                End If
                If blnDiagnose And mstrǰ��IDs = "" Then
                    rsDiag.Filter = "ҽ��ID=" & IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)))
                    If rsDiag.EOF Then
                        MsgBox """" & .TextMatrix(i, col_ҽ������) & """û�ж�Ӧ�����Ϣ�����������Ӧ����ϡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                '��������ҽ��ʱ���
                If .TextMatrix(i, COL_�������) = "Z" And .TextMatrix(i, COL_��������) = "1" Then
                    strSQL = "Select b.���� From ������ҳ a,���ű� b Where a.��Ժ����id=b.id And a.����id=[1] And a.�������� in (1,2) And a.��Ժ���� Is Null"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
                    If Not rsTmp.EOF Then
                        MsgBox "�ò��˻��ڡ�" & rsTmp!���� & "�����ۣ����ܷ��� """ & .TextMatrix(i, col_ҽ������) & """ ���Ȱ������۳�Ժ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                '����סԺҽ��ʱ���
                If .TextMatrix(i, COL_�������) = "Z" And (.TextMatrix(i, COL_��������) = "2" Or .TextMatrix(i, COL_��������) = "1") Then
                    If Sys.NewSystemSvr("ԤԼ����", "סԺ����", "", "") Then
                        '����סԺԤԼ����ж�ͨ��ҽ��
                        strSQL = "Select 1 From ����ҽ����¼ a,������ĿĿ¼ b Where a.������Ŀid=b.id and a.�������='Z' and b.�������� in ('1','2') and a.ҽ��״̬=8 and a.�Һŵ�=[1] and rownum<2"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
                        If Not rsTmp.EOF Then
                            MsgBox "�ò����Ѿ����͹�һ��סԺ����ҽ���������ٷ��ͣ������������ҽ����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
            '��鵱ǰҽ���ķ�ҩҩ���Ƿ��ϰ࣬���ϰ��������ʾ
            If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing And _
               InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                If blnFirst Then bln�ϰల�� = Check�ϰల��(True): blnFirst = False
                If bln�ϰల�� Then
                   bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
                   strSQL = "Select A.����id" & vbNewLine & _
                           "From   ��������˵�� A, ���Ű��� B" & vbNewLine & _
                           "Where  A.����id = B.����id And A.�������� In ('��ҩ��', '��ҩ��', '��ҩ��') And b.����id = [1] And B.���� = [2] " & _
                           "And To_Char(Sysdate, 'HH24:MI:SS') Between To_Char(B.��ʼʱ��, 'HH24:MI:SS') And To_Char(B.��ֹʱ��, 'HH24:MI:SS')"
                   Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_ִ�п���ID)), bytDay)
                   If rsTmp.RecordCount = 0 Then
                       str�������� = Sys.RowValue("���ű�", Val(.TextMatrix(i, COL_ִ�п���ID)), "����")
                       If MsgBox(str�������� & "�Ѿ��°�,�Ƿ�������͵�" & str�������� & " ��" & vbNewLine & "��Ҫ��������ҩ����ҩƷ����ѡ��", vbYesNo + vbInformation, gstrSysName) = vbNo Then
                           Exit Sub
                       End If
                   End If
                End If
            End If
        Next
        If lngCount = 0 Then
            MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If zlCommFun.ActualLen(txtNote.Text) > txtNote.MaxLength Then
        MsgBox "����ժҪ�����ݹ������������ " & txtNote.Text \ 2 & " �����ֻ� " & txtNote.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtNote.SetFocus: Exit Sub
    End If
    
    If Button.Key = "����Ϊ�շѵ�" Then
        '���Һ���Ч������������������Ϊ�շѵ�
        '����鷢��Ϊ���ʵ�
        'δ���ҽ��ҽ������
        '����鼱��Һ�
        strMsg = ExpendSendClear(mstr�Һŵ�, True)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����Ϊ��Ѽ��ʵ��в���ѡ
        With vsAdvice
            .Redraw = flexRDNone
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If .TextMatrix(i, COL_��Ѽ���) = 1 Then
                        .Row = i
                        .Col = COL_ѡ��
                        Call vsAdvice_KeyPress(32)
                    End If
                End If
            Next
            .Redraw = flexRDDirect
        End With
        
        '��λ���˷���Ϊ�շѵ�ʱ��������
        If mbln��λ���� And mstr��λ <> "" Then
            If MsgBox("��ǰ�����Ǻ�Լ��λ���ˣ�����""" & mstr��λ & """���Ƿ�Ҫ���ͳ�Ϊ�շѵ��ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnUnload = False: Exit Sub
            End If
        End If
        
        If tbrMain.Buttons("����Ϊ�շѵ�").Enabled And tbrMain.Buttons("����Ϊ���ʵ�").Enabled Then
            If MsgBox("����ҽ�����͵ķ��ý�����Ϊ�շѵ��ݣ�ȷʵҪ������ѡ���ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                mblnUnload = False: Exit Sub '�Զ�����ʱ,�û��жϲ��رմ���
            End If
        End If
        bln���� = False
    ElseIf Button.Key = "����Ϊ���ʵ�" Then
               
        If tbrMain.Buttons("����Ϊ�շѵ�").Enabled And tbrMain.Buttons("����Ϊ���ʵ�").Enabled Then
            If MsgBox("����ҽ�����͵ķ��ý�����Ϊ���ʵ��ݣ�ȷʵҪ������ѡ���ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                mblnUnload = False: Exit Sub
            End If
        End If
        bln���� = True
    End If
    
    lng���ͺ� = SendAdvice(bln����)
    If lng���ͺ� <> 0 Then
        mblnSend = True
        
        '��ӡ�������ָ����
        lngָ������ӡ = Val(zlDatabase.GetPara("ָ������ӡ��ʽ", glngSys, p����ҽ���´�))
        If lngָ������ӡ = 1 Then
            If MsgBox("�Ƿ�Ҫ��ӡ����ָ������", vbQuestion + vbYesNo + vbDefaultButton1, "����ָ������ӡ") = vbYes Then
                lngָ������ӡ = 2
            End If
        End If
        If lngָ������ӡ = 2 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "���ͺ�=" & lng���ͺ�, "����ID=" & mlng����ID, "�Һŵ�=" & mstr�Һŵ�, "PrintEmpty=0", 2)
        End If
        
        '��ӡ���Ƶ���
        SwitchPrintSet glngSys & "\" & p����ҽ���´�
        Call frmSendBillPrint.ShowMe(lng���ͺ�, 1, Me, mstrǰ��IDs)
        SwitchPrintSet glngSys & "\" & p����ҽ���´�, True
        '���ȫ���������,���˳�
        If vsAdvice.Rows = 2 Then
            If Val(vsAdvice.TextMatrix(1, COL_ID)) = 0 Then
                Unload Me: Exit Sub
            End If
        End If
        Call GetPatiInfo
        Call ShowPatiMoney
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbrSys_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    Select Case Button.Key
        Case "ȫѡ"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "ȫ��"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "����"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "�˳�"
            Unload Me
    End Select
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, Optional strҽ��IDs As String)
'���ܣ����ݿɼ��е�ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_ѡ�� Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            'ȡ��ѡ��ʱ
            If Not (.Cell(flexcpData, lngRow, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, lngRow, COL_ѡ��) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
                '1.�����Ӧ�ķ��ü����ͼ�¼��д
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "ҽ��ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '��ΪҪʹ��BookMark����˻ָ�
                End If
                '2.�����Ӧ�ķ��ͼƼ������ۼ�
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "ҽ��ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '4.��������͵�ǩ��ҽ����ID
                If strҽ��IDs <> "" Then
                    strҽ��IDs = strҽ��IDs & ","
                    strҽ��IDs = Replace(strҽ��IDs, "," & i & ",", ",")
                    If strҽ��IDs <> "" Then
                        strҽ��IDs = Left(strҽ��IDs, Len(strҽ��IDs) - 1)
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'���ܣ�����ָ��ҽ���У����ظ�ҽ���пɼ�����
    Dim lng��ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        'һ����ҩ�Ķ�λ����һҩƷ��
        If blnFirst Then
            If .TextMatrix(lngRow, COL_�������) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 _
                And Val(.TextMatrix(lngRow, COL_���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_���ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Sub txtNote_GotFocus()
    Call zlControl.TxtSelAll(txtNote)
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowAdvicePrice(NewRow)
                
                'ȱʡѡ��Ƽ�ҽ��(�������)
                Call ShowDefaultRow
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_ѡ�� + 1 - .FixedCols Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_ҽ������ Then
            .AutoSize col_ҽ������
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ƶ��: lngRight = COL_�÷�
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
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_ѡ�� Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_ѡ��) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_ѡ��) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        End If
    End With
End Sub

Private Sub ShowDefaultRow()
'���ܣ����ڿ��ԼƼ۵�ҽ��,ȱʡ����һ�в�����ȱʡ�Ƽ�ҽ��
'˵����ComboList="#ҽ��ID1;�Ƽ�ҽ��1|#ҽ��ID2;�Ƽ�ҽ��2|..."
'      ���ڵ�һ����ʾ�Ƽ۱�ͻس�������ʱ����
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lngҽ��ID As Long, lng�к� As Long, str�Ƽ�ҽ�� As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_�Ƽ�ҽ��) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_�Ƽ�ҽ��), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_�к�)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_�շ�ϸĿID)) <> 0 Then
                '��һ����ʾʱȱʡ����һ��
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '���ǵ�һ����ʾʱȱʡ�Ƽ�ҽ������һ����ͬ
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_�̶�)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_�к�)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lngҽ��ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str�Ƽ�ҽ�� = Replace(arrCombo(i), "#" & lngҽ��ID & ";", "")
                lng�к� = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                If blnHave Then
                    If lng�к� = Val(.TextMatrix(lngRow - 1, COLP_�к�)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            'ģ��ѡ������Ƽ�ҽ��
            .TextMatrix(lngRow, COLP_�к�) = lng�к�
            .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
            .Cell(flexcpData, lngRow, COLP_�Ƽ�ҽ��) = .TextMatrix(lngRow, COLP_�Ƽ�ҽ��)
            
            'ֻ��һ���Ƽ�ҽ��ʱ����ͣ��
            If UBound(arrCombo) = 0 Then
                .Col = COLP_�շ���Ŀ
            Else
                .Col = COLP_�Ƽ�ҽ��
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngԭ��ID As Long, lngҽ��ID As Long
    Dim int�������� As Integer, intԭ�������� As Integer
    Dim lng�շ�ϸĿID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_�Ƽ�ҽ�� Then
            '�������ComboData,TextMatrixȡֵ��ΪComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lngҽ��ID = .ComboData
                If lngҽ��ID < 0 Then
                    int�������� = Val(Left(Abs(lngҽ��ID), 1))
                    lngҽ��ID = Val(Mid(Abs(lngҽ��ID), 2))
                End If
                lngԭ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
                intԭ�������� = Val(.TextMatrix(Row, COLP_��������))
                lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                                
                '���üƼ�ҽ���Ƿ�������ͬ�շ�ϸĿ
                If lng�շ�ϸĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�Ѿ��������շ���Ŀ""" & .TextMatrix(Row, COLP_�շ���Ŀ) & """��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                'ԭ����ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                If lngԭ��ID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '���������˵ļƼ�ҽ������
                i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                .TextMatrix(Row, COLP_�к�) = i
                .TextMatrix(Row, COLP_��������) = int��������
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng�շ�ϸĿID <> 0 Then
                    '��ѡ���ҽ���Ƿ��д�������޸ĺ����Ŀ�Ƿ����
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " ��������=" & int�������� & " And ����=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_����) = IIF(blnHaveSub, "��", "")
                
                    '���»����Ӽ�¼������
                    If lngԭ��ID = 0 Then
                        mrsPrice.AddNew '����
                    Else '����
                        mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    End If
                    mrsPrice!ҽ��ID = lngҽ��ID
                    If Val(vsAdvice.TextMatrix(i, COL_���ID)) <> 0 Then
                        mrsPrice!���ID = vsAdvice.TextMatrix(i, COL_���ID)
                    Else
                        mrsPrice!���ID = Null
                    End If
                    mrsPrice!�������� = int��������
                    mrsPrice!�շѷ�ʽ = 0
                    If lngԭ��ID = 0 Then
                        mrsPrice!�շ�ϸĿID = lng�շ�ϸĿID
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_�Ƽ�����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_��������))
                        mrsPrice!��� = Val(.TextMatrix(Row, COLP_���))
                        mrsPrice!�̶� = 0
                    End If
                    mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
            End If
        ElseIf Col = COLP_�շ���Ŀ Or Col = COLP_ִ�п��� Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        ElseIf Col = COLP_�Ƽ����� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '���ݿɷ�༭����
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_�Ƽ�ҽ�� Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_�շ���Ŀ Or NewCol = COLP_ִ�п��� Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    If NewRow <> OldRow Then
        '��ʾҩƷ���
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_�к�))
            If lngRow <> 0 And .TextMatrix(NewRow, COLP_�շ����) <> "" Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_�շ����)) > 0 _
                    Or .TextMatrix(NewRow, COLP_�շ����) = "4" And Val(.TextMatrix(NewRow, COLP_��������)) = 1 Then
                    '��ʾҩƷ���������ĵĿ��
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                        If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, col_ҽ������) & "," & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & IIF(Val(vsAdvice.TextMatrix(lngRow, COL_���)) > 0, "�п��", "�޿��")
                        Else
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, col_ҽ������) & "," & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & "���ÿ��:" & FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_���)), 5) & vsAdvice.TextMatrix(lngRow, COL_���ﵥλ)
                        End If
                    Else
                        'ͬһ������ȡ:ҩƷ�����ﵥλ,���İ��ۼ۵�λ
                        If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                            If GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))) > 0 Then
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "�п��"
                            Else
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "�޿��"
                            End If
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ��:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))), 5) & .TextMatrix(NewRow, COLP_��λ)
                        End If
                    End If
                End If
            End If
        End With
        
        '��ʾҽ������
        stbThis.Panels(3).Text = Getҽ������(NewRow)
    End If
End Sub

Private Function Getҽ������(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еķ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str���� As String
    
    With vsPrice
        If Val(.TextMatrix(lngRow, COLP_�շ�ϸĿID)) <> 0 Then
            strSQL = "Select N.���� From ����֧����Ŀ M,����֧������ N Where M.�շ�ϸĿID=[1] And M.����ID=N.ID And M.����=[2]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COLP_�շ�ϸĿID)), mint����)
            If Not rsTmp.EOF Then str���� = NVL(rsTmp!����)
        End If
    End With
    Getҽ������ = IIF(str���� <> "", "ҽ������:" & str����, "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵�������ص��кŷ�Χ��������ҩ;�����к�
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

Private Sub InitAdviceTable()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;" & _
        "Ӥ��,550,1;ҽ������,3000,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;���,850,7;" & _
        "Ƶ��,1000,1;�÷�,1000,1;ҽ������,1500,1;ִ��ʱ��,1000,1;ִ�п���,850,1;ִ������,850,1;" & _
        "ID;���ID;���˿���ID;��������ID;����ҽ��;�������;������ĿID;�걾��λ;��鷽��;ִ�б��;�Ƽ�����;ִ������ID;" & _
        "ִ�п���ID;�շ�ϸĿID;Ƶ�ʴ���;Ƶ�ʼ��;�����λ;����ϵ��;�����װ;���ﵥλ;�ɷ����;���;" & _
        "����;�ֽ�ʱ��;�״�ʱ��;ĩ��ʱ��;ǰ��ID;ǩ��ID;�Թܱ���;��������;������־;��Ѽ���;���㷽ʽ;��ʼʱ��;ִ�а���;ִ�з���;�������;��ҩ����"
    arrHead = Split(strHead, ";")
    With vsAdvice
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
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�к�;�շ�ϸĿID;�̶�;���;�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2000,1;�Ƽ�����,900,7;" & _
        "����,800,7;��λ,500,1;����,1000,7;Ӧ�ս��,1050,7;ʵ�ս��,1050,7;ִ�п���,1000,1;��������,850,1;" & _
        "����,450,4;�շѷ�ʽ,1500,1;�շ����;ִ�п���ID;��������;��������"
    arrHead = Split(strHead, ";")
    With vsPrice
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
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub DeleteCurRow(ByVal lngRow As Long)
'���ܣ��ڴ���������嵥�Ĺ�����ɾ������������(��ҩ�ƻ��ҩ)
    Dim lngҽ��ID As Long, lng���ID As Long, i As Long
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
                
        'ɾ����ǰ��
        .RemoveItem lngRow
        
        'ɾ�������
        If lng���ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID _
                    Or Val(.TextMatrix(i, COL_ID)) = lng���ID Then
                    .RemoveItem i
                End If
            Next
        Else
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Sub InitSeekSet(rsSeek As ADODB.Recordset)
'���ܣ���ʼ�����ڻ��ܼ����ۿ۵���ʱ��¼��
    Set rsSeek = New ADODB.Recordset
    rsSeek.Fields.Append "��������", adInteger
    rsSeek.Fields.Append "�����ǩ", adVariant
    rsSeek.Fields.Append "������ID", adBigInt
    rsSeek.Fields.Append "�ϼ�", adCurrency, , adFldIsNullable
    rsSeek.CursorLocation = adUseClient
    rsSeek.LockType = adLockOptimistic
    rsSeek.CursorType = adOpenStatic
    rsSeek.Open
End Sub

Private Sub InitPriceRecordset()
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "��������", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շѷ�ʽ", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շ����", adVarChar, 1
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
    mrsPrice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable '��ۼ۸�
    mrsPrice.Fields.Append "����", adInteger '�����Ƿ��������
    mrsPrice.Fields.Append "���", adInteger
    mrsPrice.Fields.Append "����", adInteger
    mrsPrice.Fields.Append "�̶�", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset)
'��ʼ����¼��
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-�Ƽ�,2-����,3-ǩ��,4-����,5-����
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '����NO�滻����ʱ����
    rsSQL.Fields.Append "�������", adVarChar, 8
    rsSQL.Fields.Append "��ǰ��ҽ��ID", adInteger
    rsSQL.Fields.Append "����", adVarChar, 38  'ҽ�������к� �շ���ĿID ִ�в���ID �� ��_���ָ�
    rsSQL.Fields.Append "NewNO", adVarChar, 30, adFldIsNullable '��¼�滻���NOֵ,�����󷽴���
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '�Ƽ������ۼƼ�¼��
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsTotal.Fields.Append "��ĿID", adBigInt
    rsTotal.Fields.Append "�ⷿID", adBigInt
    rsTotal.Fields.Append "����", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    '��¼�Թܱ���
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "����", adVarChar, 18
    rsNumber.Fields.Append "���ID", adBigInt
    rsNumber.Fields.Append "��������", adVarChar, 18
    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
    rsNumber.Fields.Append "������ĿID", adVarChar, 18
    rsNumber.Fields.Append "Ӥ��", adBigInt
    rsNumber.Fields.Append "������־", adBigInt
    rsNumber.Fields.Append "�걾", adVarChar, 18
    rsNumber.Fields.Append "�ɼ�����ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '��ǰ���˱���Ҫ���͵ķ���
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsMoneyNow.Fields.Append "������ĿID", adBigInt
    rsMoneyNow.Fields.Append "�շ���ĿID", adBigInt
    rsMoneyNow.Fields.Append "�Թܱ���", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "��������", adVarChar, 50, adFldIsNullable
    rsMoneyNow.Fields.Append "�շѷ�ʽ", adInteger
    rsMoneyNow.Fields.Append "�շ�ʱ��", adVarChar, 10
    rsMoneyNow.Fields.Append "ִ�в���ID", adBigInt
    
    rsMoneyNow.Fields.Append "��ҽ��ID", adBigInt '���ID��Ϊ�յ�ҽ���е�ҽ��ID
    rsMoneyNow.Fields.Append "��鲿λ", adVarChar, 100
    rsMoneyNow.Fields.Append "��鷽��", adVarChar, 100
    rsMoneyNow.Fields.Append "����", adDouble '�շ�����
    
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '��ǰ���˱��η��͵ķ�����Ŀ����
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "����ID", adBigInt
    rsItems.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "ҽ��ID", adBigInt
    rsItems.Fields.Append "�շ����", adVarChar, 1
    rsItems.Fields.Append "�շ�ϸĿID", adBigInt
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "ʵ�ս��", adDouble
    rsItems.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "���ID", adBigInt, , adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
End Sub

Private Function LoadAdvicePrice(ByVal lngRow As Long, rsSend As ADODB.Recordset, cur�ϼ� As Currency) As Boolean
'���ܣ���ȡָ��ҽ��(����ǰ��)�ļƼ۹�ϵ����ʱ��¼��,������ȱʡ���ͽ��(���ѱ����)
'���أ�cur�ϼ�=�������ҽ�����ͽ��(��ҩ���δ��,��Ҫ����۸�����)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, strPrice As String
    Dim str�������� As String, arr�������� As Variant
    Dim blnDo As Boolean, i As Long, k As Long
    Dim dbl���� As Double, dbl���� As Double, dblӦ�� As Double
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim bln�������� As Boolean, lng��ĿID As Long
    Dim lng������ID As Long, blnHaveSub As Boolean
    Dim lngִ�п���ID As Long, cur��� As Currency
    Dim lng����ID As Long, bln��Ѽ��� As Boolean
    
    On Error GoTo errH
    
    cur��� = 0
    With vsAdvice
        bln��Ѽ��� = .TextMatrix(lngRow, COL_��Ѽ���) = 1
        
        If InStr(",4,5,6,7,", rsSend!�������) > 0 Then
            '��ΪԺ��ִ��(�Ա�ҩ),ҩƷ������Ϊ����,�ҹ̶������Ƽ�
            If NVL(rsSend!ִ������, 0) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!ҽ��ID = rsSend!ID
                mrsPrice!���ID = rsSend!���ID
                mrsPrice!�������� = 0
                mrsPrice!�շѷ�ʽ = 0
                mrsPrice!�շ���� = rsSend!�������
                mrsPrice!�շ�ϸĿID = rsSend!�շ�ϸĿID
                mrsPrice!ִ�п���ID = rsSend!ִ�п���ID
                mrsPrice!���� = 1
                mrsPrice!���� = NVL(rsSend!��������, 0)
                mrsPrice!��� = NVL(rsSend!�Ƿ���, 0)
                mrsPrice!�̶� = 1
                mrsPrice!���� = 0
                                
                '���͵���������
                If rsSend!������� = "7" Then
                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                    If NVL(rsSend!�ɷ����, 0) = 0 Then
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����)) / NVL(rsSend!����ϵ��, 1)
                    Else
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_����)) / NVL(rsSend!����ϵ��, 1) / NVL(rsSend!�����װ, 1)) * NVL(rsSend!�����װ, 1)
                    End If
                Else
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * NVL(rsSend!�����װ, 1)
                End If
                dbl���� = Format(dbl����, "0.00000")
                                
                '��¼�ۼ۵���
                If NVL(rsSend!�Ƿ���, 0) = 0 Or rsSend!������� = "4" And NVL(rsSend!��������, 0) = 0 Then
                    mrsPrice!���� = Format(CalcPrice(rsSend!�շ�ϸĿID, , , True, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                    mrsPrice!���� = Format(CalcDrugPrice(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), dbl����, , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                End If
                mrsPrice.Update
                 
                If Not bln��Ѽ��� Then
                    '����ҽ�����ͽ��(���ѱ���۵�ʵ�ս��)
                    If Not IsNull(mrsPati!�ѱ�) Then
                        If NVL(rsSend!�Ƿ���, 0) = 0 Or rsSend!������� = "4" And NVL(rsSend!��������, 0) = 0 Then
                            cur��� = Format(CalcPrice(rsSend!�շ�ϸĿID, mrsPati!�ѱ�, dbl����, , NVL(rsSend!ִ�п���ID, 0), , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDec)
                        Else
                            cur��� = Format(CalcDrugPrice(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), dbl����, mrsPati!�ѱ�, , 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), "0.00000")
                        End If
                    Else
                        If gbln�Ӱ�Ӽ� Then
                            '����Ӱ�Ӽ�
                            If NVL(rsSend!�Ƿ���, 0) = 0 Or rsSend!������� = "4" And NVL(rsSend!��������, 0) = 0 Then
                                dbl���� = Format(CalcPrice(rsSend!�շ�ϸĿID, , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                                dbl���� = Format(CalcDrugPrice(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), dbl����, , , 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            End If
                            cur��� = Format(mrsPrice!���� * dbl���� * dbl����, gstrDec)
                        Else
                            cur��� = Format(mrsPrice!���� * dbl���� * mrsPrice!����, gstrDec)
                        End If
                    End If
                End If
            End If
            
            cur�ϼ� = cur���
        Else
            'ȡ�����շ� ��ϵ�еĶ���(����ʱ�Ŷ��Ƽ�):�����Ƽ�,��Ϊ������Ժ��ִ��
            If NVL(rsSend!�Ƽ�����, 0) = 0 And InStr(",0,5,", NVL(rsSend!ִ������, 0)) = 0 Then
                lng����ID = 0 '�����Թܷ���,ֻ��ȡ�Թܶ�Ӧ�����ķ���
                If .TextMatrix(lngRow, COL_�Թܱ���) <> "" Then
                    lng����ID = GetTubeMaterial(.TextMatrix(lngRow, COL_�Թܱ���))
                End If
            
                dbl���� = Format(Val(.TextMatrix(lngRow, COL_����)), "0.00000")
                
                '���ֶ�Ӧ�ļƼ����
                If Not IsNull(rsSend!�걾��λ) And Not IsNull(rsSend!��鷽��) Then
                    strPrice = " And c.��鲿λ=[3] And c.��鷽��=[4] And Nvl(c.��������,0)=0"
                ElseIf NVL(rsSend!ִ�б��, 0) = 0 Then
                    strPrice = " And c.��鲿λ Is Null And c.��鷽�� is Null And Nvl(c.��������,0)=0"
                Else 'Ŀǰ�������Ի����м��յ����
                    strPrice = " And c.��鲿λ Is Null And c.��鷽�� is Null And Nvl(c.��������,0) IN(0,1)"
                End If
                
                bln�������� = (rsSend!������� = "F" And Not IsNull(rsSend!���ID))
                
                strPrice = "Select * From (" & _
                        "Select C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,c.���ÿ���id" & _
                        " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                        " From �����շѹ�ϵ C Where C.������ĿID=[1]" & strPrice & _
                        "      And (C.���ÿ���ID is Null And C.������Դ = 0 or C.���ÿ���ID = [2] And C.������Դ = 1)" & _
                        " ) Where Nvl(���ÿ���id, 0) = Top"
                strSQL = _
                    " Select C.���,A.�շ���ĿID,A.�շ�����,A.���ж���,B.������ĿID," & _
                    " C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,B.ȱʡ�۸�,B.�ּ�)" & IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "") & " as ����,C.�Ƿ���," & _
                    " Nvl(A.������Ŀ,0) as ����,D.��������,[2] as ִ�п���ID,C.���ηѱ�,Nvl(A.��������,0) as ��������," & _
                    " Nvl(A.�շѷ�ʽ,0) as �շѷ�ʽ" & _
                    " From (" & strPrice & ") A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D" & _
                    " Where A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And A.�շ���ĿID=D.����ID(+)" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "B", "6", "7", "8") & _
                    " And C.������� IN(1,3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And (Nvl(A.�շѷ�ʽ,0)=1 And C.���='4' And A.�շ���ĿID=[5] Or Not(Nvl(A.�շѷ�ʽ,0)=1 And C.���='4' And [5]<>0))" & _
                    " Order by ��������,����,A.�շ���ĿID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!������ĿID), _
                    Val(NVL(rsSend!ִ�п���ID, 0)), CStr(NVL(rsSend!�걾��λ)), CStr(NVL(rsSend!��鷽��)), lng����ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                
                'ȷ���Ƽ�֮���Ƿ���������Լ���������ID
                arr�������� = Array()
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        If InStr(str�������� & ",", "," & rsTmp!�������� & ",") = 0 Then
                            str�������� = str�������� & "," & rsTmp!��������
                        End If
                        rsTmp.MoveNext
                    Loop
                    arr�������� = Split(Mid(str��������, 2), ",")
                End If
                
                For k = 0 To UBound(arr��������)
                    rsTmp.Filter = "��������=" & arr��������(k)
                    lng��ĿID = 0: cur��� = 0
                    lng������ID = 0: blnHaveSub = False
                    If Not rsTmp.EOF And gbln��������ۿ� Then
                        Do While Not rsTmp.EOF
                            If NVL(rsTmp!����, 0) = 0 Then
                                'SQL����������ǰ��,ֻȡ����Ŀ�ĵ�һ������
                                If lng������ID = 0 Then lng������ID = rsTmp!������ĿID
                            ElseIf NVL(rsTmp!����, 0) = 1 Then
                                blnHaveSub = True: Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsTmp.MoveFirst
                    End If
                    
                    Do While True
                        blnDo = False
                        If rsTmp.EOF Then
                            If lng��ĿID <> 0 Then blnDo = True
                        Else
                            If rsTmp!�շ���ĿID <> lng��ĿID And lng��ĿID <> 0 Then blnDo = True
                        End If
                        If blnDo Then
                            If Not IsNull(mrsPrice!����) Then
                                mrsPrice!���� = Format(mrsPrice!����, gstrDecPrice)
                            End If
                            mrsPrice.Update
                            
                            'ҽ�����ͽ��
                            cur��� = cur��� + Format(curʵ��, gstrDec)
                        End If
                        If rsTmp.EOF Then Exit Do
                        
                        '------------------------------------
                        If rsTmp!�շ���ĿID <> lng��ĿID Then
                            curʵ�� = 0
                            mrsPrice.AddNew
                            mrsPrice!ҽ��ID = rsSend!ID
                            mrsPrice!���ID = rsSend!���ID
                            mrsPrice!�������� = NVL(rsTmp!��������, 0)
                            mrsPrice!�շѷ�ʽ = NVL(rsTmp!�շѷ�ʽ, 0)
                            mrsPrice!�շ���� = rsTmp!���
                            mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
                            mrsPrice!���� = NVL(rsTmp!�շ�����, 0)
                            mrsPrice!���� = NVL(rsTmp!��������, 0)
                            mrsPrice!��� = NVL(rsTmp!�Ƿ���, 0)
                            mrsPrice!�̶� = NVL(rsTmp!���ж���, 0)
                            mrsPrice!���� = NVL(rsTmp!����, 0)
                            
                            If .TextMatrix(lngRow, COL_�������) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" And InStr(",5,6,", rsTmp!���) > 0 Then
                                'ԭҺƤ�����⡣�󶨵�ҩƷ�������û��ָ��������ԭ���߼�
                                If Val(.TextMatrix(lngRow, COL_��ҩ����)) <> 0 Then
                                    lngִ�п���ID = Val(.TextMatrix(lngRow, COL_��ҩ����))
                                Else
                                    lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                                End If
                                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, NVL(rsSend!���˿���id, 0), 0, 1, lngִ�п���ID)
                            Else
                                'ִ�п���:��ҩ��ҩƷ���������ĵ�ר��ȡ
                                lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                                If rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Or InStr(",5,6,7,", rsTmp!���) > 0 Then
                                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, NVL(rsSend!���˿���id, 0), 0, 1, lngִ�п���ID)
                                End If
                            End If
                                                        
                            If lngִ�п���ID <> 0 Then
                                mrsPrice!ִ�п���ID = lngִ�п���ID
                            Else
                                mrsPrice!ִ�п���ID = Null
                            End If
                        End If
                        lng��ĿID = rsTmp!�շ���ĿID
                        
                        '���㵥�ۺ�ʵ��
                        If NVL(rsTmp!�Ƿ���, 0) = 1 And InStr(",5,6,7,", rsTmp!���) > 0 Then
                            '��ҩ��ҩƷ�Ƽ۰�ʱ�ۼ���(��һ������),���������Ҫ��ҽ������
                            mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, NVL(mrsPrice!ִ�п���ID, 0), dbl���� * NVL(rsTmp!�շ�����, 0), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
    
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If Not bln��Ѽ��� Then
                                If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                    curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                        mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                                Else
                                    curʵ�� = curʵ�� + curӦ��
                                End If
                            End If
                        ElseIf NVL(rsTmp!�Ƿ���, 0) = 1 And rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Then
                            '�������õ�ʱ�����ĺ�ҩƷһ������
                            mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, NVL(mrsPrice!ִ�п���ID, 0), dbl���� * NVL(rsTmp!�շ�����, 0), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
    
                            curӦ�� = Format(dblӦ��, gstrDec)
                                                        
                            If Not bln��Ѽ��� Then
                                If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                    curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                        mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                                Else
                                    curʵ�� = curʵ�� + curӦ��
                                End If
                            End If
                        Else '�̶��۸����ͨ���(ֻ��һ��������Ŀ)
                            mrsPrice!���� = NVL(mrsPrice!����, 0) + NVL(rsTmp!����, 0)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(NVL(rsTmp!����, 0), gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
                            
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If Not bln��Ѽ��� Then
                                If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                    curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                        mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                                Else
                                    curʵ�� = curʵ�� + curӦ��
                                End If
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                    
                    '������Ŀ���ܼ����ۿ�
                    If gbln��������ۿ� And blnHaveSub And lng������ID <> 0 And Not bln��Ѽ��� Then
                        cur��� = Format(ActualMoney(NVL(mrsPati!�ѱ�) & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), lng������ID, cur���), gstrDec)
                    End If
                    
                    cur�ϼ� = cur�ϼ� + cur���
                Next
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'���ܣ����ݵ�ǰҽ���л�ȡ��ѡ��ļƼ�ҽ������
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
'˵����ע�������Ǹ��ݾ���ҽ����ȡ
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If .Cell(flexcpData, lngRow, COL_ID) = 3 Then
            '��ҩ�÷�����ҩ�÷�,��ҩ�巨
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", CLng(.Cell(flexcpData, i, COL_ID))) > 0 Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�̶�, 0) = 0 Then
                                    If .Cell(flexcpData, i, COL_ID) = 2 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, col_ҽ������)
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            If .Cell(flexcpData, i, COL_ID) = 2 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, col_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf .Cell(flexcpData, lngRow, COL_ID) = 4 Then
            '�ɼ�������
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!�̶�, 0) = 0 Then
                                If .TextMatrix(i, COL_�������) = "C" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, col_ҽ������)
                                ElseIf .TextMatrix(i, COL_�������) = "E" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, col_ҽ������)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        If .TextMatrix(i, COL_�������) = "C" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, col_ҽ������)
                        ElseIf .TextMatrix(i, COL_�������) = "E" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, col_ҽ������)
                        End If
                        If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                            strCombo = strCombo & "|#" & strTmp
                        End If
                    End If
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '���г�ҩ����ҩ;��
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!�̶�, 0) = 0 Then
                                strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, col_ҽ������)
                                Exit For
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, col_ҽ������)
                    End If
                End If
            End If
        Else
            'һ���������飬����Ѫҽ���������ҽ��
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�̶�, 0) = 0 Then
                                    If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .TextMatrix(i, COL_�걾��λ) & "(" & .TextMatrix(i, COL_��鷽��) & ")"
                                    ElseIf .TextMatrix(i, COL_�������) = "E" And .TextMatrix(lngRow, COL_�������) = "K" Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��Ѫ;��-" & .Cell(flexcpData, i, col_ҽ������)
                                    Else
                                        If mrsPrice!�������� <> 0 Then
                                            '���շ��ã�Ŀǰ�������Ĵ��Ժ����м���
                                            lngTmp = -1 * Val(mrsPrice!�������� & Val(.TextMatrix(i, COL_ID)))
                                            strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������) & _
                                                "(" & Decode(Val(.TextMatrix(i, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                                        Else
                                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������)
                                        End If
                                    End If
                                    
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            'δ���üƼ۵ģ�����ѡ����ӼƼ���Ŀ
                            If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .TextMatrix(i, COL_�걾��λ) & "(" & .TextMatrix(i, COL_��鷽��) & ")"
                            ElseIf .TextMatrix(i, COL_�������) = "E" And .TextMatrix(lngRow, COL_�������) = "K" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��Ѫ;��-" & .Cell(flexcpData, i, col_ҽ������)
                            Else
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                            
                            '���շ��ã�Ŀǰ�������Ĵ��Ի����м���
                            If .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 _
                                And (Val(.TextMatrix(i, COL_ִ�б��)) = 1 Or Val(.TextMatrix(i, COL_ִ�б��)) = 2) Then
                                lngTmp = -1 * Val(1 & Val(.TextMatrix(i, COL_ID)))
                                strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������) & _
                                    "(" & Decode(Val(.TextMatrix(i, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'���ܣ�����ҽ���Ƽ۹�ϵ�����㲢��ʾָ��ҽ���ķ���(����ҽ�������ܶ���)
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
    Dim rsTmp As New ADODB.Recordset
    Dim rsExeDays As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str�Ƽ�ҽ�� As String
    Dim str��λ As String, dbl���� As Double
    Dim bln�������� As Boolean, strCombo As String, str�к� As String, str�ֽ�ʱ�� As String
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim dbl��ǰ���� As Double, dbl��ǰӦ�� As Double, cur��ǰӦ�� As Currency, cur��ǰʵ�� As Currency
    Dim lng�к� As Long, cur�ϼ� As Currency, bln��Ѽ��� As Boolean
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    Dim strPriceType As String
        
    On Error GoTo errH
    
    '���ڻ��ܼ����ۿ۵���ʱ��¼��
    rsMain.Fields.Append "ҽ���к�", adBigInt
    rsMain.Fields.Append "��������", adInteger
    rsMain.Fields.Append "�����к�", adBigInt
    rsMain.Fields.Append "������ID", adBigInt
    rsMain.Fields.Append "ҽ���ϼ�", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        bln��Ѽ��� = .TextMatrix(lngRow, COL_��Ѽ���) = 1
    
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnFirst = False 'һ����ҩ���Ƿ��һҩƷ��
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or ҽ��ID=" & Val(.TextMatrix(lngRow, COL_���ID))
            Else
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or ���ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '�Ƽ�ҽ��
            bln�������� = False
            lng�к� = .FindRow(CStr(mrsPrice!ҽ��ID), , COL_ID)
            If .TextMatrix(lng�к�, COL_�������) = "4" Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf InStr(",5,6,7", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                str�Ƽ�ҽ�� = "ҩƷҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 1 Then
                str�Ƽ�ҽ�� = "��ҩ;��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 2 Then
                str�Ƽ�ҽ�� = "��ҩ�巨-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 3 Then
                str�Ƽ�ҽ�� = "��ҩ�÷�-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 4 Then
                str�Ƽ�ҽ�� = "�ɼ�����-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 5 Then
                str�Ƽ�ҽ�� = "��Ѫ;��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "C" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "������Ŀ-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "F" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                bln�������� = True
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "G" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "D" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��鲿λ-" & .TextMatrix(lng�к�, COL_�걾��λ) & "(" & .TextMatrix(lng�к�, COL_��鷽��) & ")"
            Else
                If NVL(mrsPrice!��������, 0) = 1 Then
                    '���Ի����м��շ���
                    str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������) & _
                        "(" & Decode(Val(.TextMatrix(lng�к�, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                Else
                    str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
                End If
            End If
            str�Ƽ�ҽ�� = Replace(str�Ƽ�ҽ��, "'", "''")
            
            '����:ҩƷ�����ﵥλ������,��������������
            If InStr(",5,6,", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                dbl���� = Val(.TextMatrix(lng�к�, COL_����))
            ElseIf .TextMatrix(lng�к�, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lng�к�, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) * Val(.TextMatrix(lng�к�, COL_����)) _
                        / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_�����װ))
                Else
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) _
                        * IntEx(Val(.TextMatrix(lng�к�, COL_����)) / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_�����װ)))
                End If
            Else
                If InStr(",3,4,5,6,", Val("" & mrsPrice!�շѷ�ʽ)) > 0 Then 'һ��ֻ��һ�ε�
                     '�ֽ�ʱ��
                    If .TextMatrix(lng�к�, COL_�ֽ�ʱ��) <> "" Then
                        str�ֽ�ʱ�� = .TextMatrix(lng�к�, COL_�ֽ�ʱ��)
                    Else
                        str�ֽ�ʱ�� = .Cell(flexcpData, lng�к�, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                    End If
                    
                    Set rsExeDays = GetExecDays(str�ֽ�ʱ��)
                    dbl���� = rsExeDays.RecordCount
                ElseIf InStr(",1,2,", Val("" & mrsPrice!�շѷ�ʽ)) > 0 Then 'һ�η���ֻ��һ��
                    dbl���� = 1
                Else
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����))
                End If
            End If
            dbl���� = Format(dbl���� * NVL(mrsPrice!����, 0), "0.00000")
                        
            '���SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as ���," & mrsPrice!ҽ��ID & " as ҽ��ID,ID," & _
                NVL(mrsPrice!�̶�, 0) & " as �̶�,'" & str�Ƽ�ҽ�� & "' as �Ƽ�ҽ��,���,����,����,���," & _
                "���㵥λ as ��λ," & NVL(mrsPrice!����, 0) & " as �Ƽ�����," & dbl���� & " as ����," & _
                Format(NVL(mrsPrice!����, 0), gstrDecPrice) & " as ����,��������," & lng�к� & " as �к�," & _
                " �Ƿ���,�Ӱ�Ӽ�," & IIF(bln��������, 1, 0) & " as ��������," & mrsPrice!���� & " as ����," & _
                NVL(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID,���ηѱ�," & mrsPrice!�������� & " as ��������," & _
                mrsPrice!�շѷ�ʽ & " as �շѷ�ʽ From �շ���ĿĿ¼ Where ID=" & mrsPrice!�շ�ϸĿID
                
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '��Ҫ�Ƽ۵�ҽ��ѡ��
        '���ݴ�����ҽ��ȡ�ɼƼ�ҽ��(���ܴ�mrsPriceȡ,��Ϊ�������շѹ�ϵ����ɾ��,����Ҳ�����ڼƼ���ȫ��ɾ��)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_�Ƽ�ҽ��) = strCombo
            .Editable = flexEDKbdMouse '����ѡ������Ա༭
        Else
            .ColData(COLP_�Ƽ�ҽ��) = ""
        End If
        
        '��ʾ���еļƼ���Ŀ
        If strSQL <> "" Then
            strSQL = "Select A.�к�,A.ID AS �շ�ϸĿID,A.�̶�,A.����,A.�Ƽ�ҽ��,A.���,C.���� as �������,A.ִ�п���ID,G.���� as ִ�п���," & _
                " Nvl(E.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ����," & _
                " A.��λ,A.�Ƽ�����,A.����,D.�����װ,D.���ﵥλ,Decode(A.�Ƿ���,1,A.����,B.�ּ�) as ����,F.��������," & _
                " A.��������,A.�շѷ�ʽ,A.��������,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.ԭ��,B.�ּ�,A.��������,B.�����շ���,B.������ĿID" & _
                " From (" & strSQL & ") A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,�շ���Ŀ���� E,�������� F,���ű� G" & _
                " Where A.ID=B.�շ�ϸĿID And A.���=C.���� And A.ID=D.ҩƷID(+)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3") & _
                " And A.ID=F.����ID(+) And A.ִ�п���ID=G.ID(+)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIF(gbytҩƷ������ʾ = 0, 1, 3) & _
                " Order by A.���"
                '��Ϊ������ǵ��ñ�����ˢ��,Ҫ���ֶ�̬��¼���м�¼˳��
                'Ҫ��֤��������ǰ��,LoadAdvicePriceʱ������������ǰ�棬���ұ༭��ֻ���ܼ��˴���
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�) 'û��
            
            If Not rsTmp.EOF And gbln��������ۿ� Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str�к� <> rsTmp!�к� & "_" & rsTmp!�������� & "_" & rsTmp!�շ�ϸĿID Then
                    If str�к� <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                            .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, gstrDecPrice)
                            .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                            .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                        End If
                        cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
                    End If
                    str�к� = rsTmp!�к� & "_" & rsTmp!�������� & "_" & rsTmp!�շ�ϸĿID
                    dbl���� = 0: curӦ�� = 0: curʵ�� = 0
                    .Rows = .Rows + 1
                    
                    '��ʶ�̶�����Ϊ��ɫ
                    If rsTmp!�̶� <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_�к�) = rsTmp!�к�
                    .TextMatrix(.Rows - 1, COLP_�շ�ϸĿID) = rsTmp!�շ�ϸĿID
                    .TextMatrix(.Rows - 1, COLP_�̶�) = rsTmp!�̶�
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��) = rsTmp!�Ƽ�ҽ��
                    .TextMatrix(.Rows - 1, COLP_��������) = rsTmp!��������
                    .TextMatrix(.Rows - 1, COLP_�շѷ�ʽ) = getChargeMode(Val(NVL(rsTmp!�շѷ�ʽ, 0)))
                        .Cell(flexcpData, .Rows - 1, COLP_�շѷ�ʽ) = Val(NVL(rsTmp!�շѷ�ʽ, 0))
                    .TextMatrix(.Rows - 1, COLP_���) = rsTmp!�������
                    .TextMatrix(.Rows - 1, COLP_�շ����) = rsTmp!���
                    .TextMatrix(.Rows - 1, COLP_�շ���Ŀ) = rsTmp!����
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�����) = NVL(rsTmp!�Ƽ�����, 0) '�������
                    
                    dbl���� = NVL(rsTmp!����, 0)
                    If InStr(",5,6,7,", rsTmp!���) > 0 Then '�����װ
                        .TextMatrix(.Rows - 1, COLP_��λ) = NVL(rsTmp!���ﵥλ)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0), 5)
                            dbl���� = dbl���� * NVL(rsTmp!�����װ, 1)
                        Else
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,���ת��Ϊҩ����λ��ʾʱ���������㴦��
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0) / NVL(rsTmp!�����װ, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_��λ) = NVL(rsTmp!��λ)
                        .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_ִ�п���) = NVL(rsTmp!ִ�п���)
                    .TextMatrix(.Rows - 1, COLP_ִ�п���ID) = NVL(rsTmp!ִ�п���ID, 0)
                    
                    '��ʾҽ����������
                    If Val(rsTmp!�շ�ϸĿID) > 0 Then
                        strPriceType = GetPriceType(Val(mlng����ID), Val(rsTmp!�շ�ϸĿID & ""), Val(mint����), True)
                    End If
                    '��������
                    If strPriceType = "" Then
                        .TextMatrix(.Rows - 1, COLP_��������) = NVL(rsTmp!��������)
                    Else
                        .TextMatrix(.Rows - 1, COLP_��������) = strPriceType
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_����) = IIF(NVL(rsTmp!����, 0) = 0, "", "��")
                    .TextMatrix(.Rows - 1, COLP_��������) = NVL(rsTmp!��������, 0)
                    
                    '��¼��������ָ�
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�ҽ��) = .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��)
                    .Cell(flexcpData, .Rows - 1, COLP_�շ���Ŀ) = .TextMatrix(.Rows - 1, COLP_�շ���Ŀ)
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�����) = .TextMatrix(.Rows - 1, COLP_�Ƽ�����)
                    .Cell(flexcpData, .Rows - 1, COLP_ִ�п���) = .TextMatrix(.Rows - 1, COLP_ִ�п���)
                    
                    '��¼�����������Ϣ���Ա����
                    If gbln��������ۿ� And rsTmp!���� = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") = 0 Then
                            rsClone.Filter = "�к�=" & rsTmp!�к� & " And ��������=" & rsTmp!�������� & " And ����=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!ҽ���к� = rsTmp!�к�
                                rsMain!�������� = rsTmp!��������
                                rsMain!�����к� = .Rows - 1
                                rsMain!������ID = rsTmp!������ĿID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!�к� & "_" & rsTmp!��������
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!�к� & "_" & rsTmp!��������
                            End If
                        End If
                    End If
                    
                    '��ҩƷ������ҽ����ҩƷ�͸������ļƼۣ���ʹ�̶�Ҳ�����޸�ִ�п���
                    If InStr(",5,6,7,", rsTmp!���) > 0 _
                        Or rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '���ۼ��㴦��
                If InStr(",5,6,7,", rsTmp!���) > 0 Then
                    If NVL(rsTmp!�Ƿ���, 0) = 0 Then
                        dbl��ǰ���� = NVL(rsTmp!����, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0) * NVL(rsTmp!�����װ, 1), "0.00000"), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        Else
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0), "0.00000"), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                        dbl��ǰ���� = dbl��ǰ���� * NVL(rsTmp!�����װ, 1)
                        dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    Else
                        dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                        dbl��ǰ���� = dbl��ǰ���� * NVL(rsTmp!�����װ, 1)
                    End If
                ElseIf rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 And NVL(rsTmp!�Ƿ���, 0) = 1 Then
                    '�������õ�ʱ�����ĺ�ҩƷһ������
                    dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0), "0.00000"), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                Else
                    dbl��ǰ���� = NVL(rsTmp!����, 0) '�������Ϊ��������û������
                    dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    If NVL(rsTmp!�Ƿ���, 0) = 1 Then '��¼��ҩ��۷�Χ
                        .TextMatrix(.Rows - 1, COLP_���) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_Ӧ�ս��) = CCur(NVL(rsTmp!ԭ��, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_ʵ�ս��) = CCur(NVL(rsTmp!�ּ�, 0))
                        .Editable = flexEDKbdMouse '��ҩƷ���,��ʹ�̶�Ҳ���Զ���
                    End If
                End If
                'Ӧ��
                If rsTmp!�������� = 1 Then
                    dbl��ǰӦ�� = dbl��ǰӦ�� * NVL(rsTmp!�����շ���, 100) / 100
                End If
                '����Ӱ�Ӽ�
                If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                    dbl��ǰӦ�� = dbl��ǰӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                End If
                cur��ǰӦ�� = Format(dbl��ǰӦ��, gstrDec)
                
                'ʵ��
                If gbln��������ۿ� And (rsTmp!���� = 1 Or InStr(strHaveSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") > 0) Then
                    If bln��Ѽ��� Then
                        cur��ǰʵ�� = 0
                    Else
                        cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                    End If
                    '�ۼ�ҽ���ϼ��������ۿ�
                    rsMain.Filter = "ҽ���к�=" & rsTmp!�к� & " And ��������=" & rsTmp!��������
                    rsMain!ҽ���ϼ� = NVL(rsMain!ҽ���ϼ�, 0) + cur��ǰʵ��
                    rsMain.Update
                ElseIf NVL(rsTmp!���ηѱ�, 0) = 0 And Not IsNull(mrsPati!�ѱ�) Then
                    If bln��Ѽ��� Then
                        cur��ǰʵ�� = 0
                    Else
                        cur��ǰʵ�� = Format(ActualMoney(mrsPati!�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, cur��ǰӦ��, rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), _
                            dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                    End If
                Else
                    If bln��Ѽ��� Then
                        cur��ǰʵ�� = 0
                    Else
                        cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                    End If
                End If
                
                dbl���� = dbl���� + dbl��ǰ����
                curӦ�� = curӦ�� + cur��ǰӦ��
                curʵ�� = curʵ�� + cur��ǰʵ��
                
                rsTmp.MoveNext
            Next
            If str�к� <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                    .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, gstrDecPrice)
                    .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                    .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                End If
                cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
            End If
        End If
        
        '���ܼ����ۿ�
        If gbln��������ۿ� And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                If bln��Ѽ��� Then
                    cur��ǰʵ�� = 0
                Else
                    cur��ǰʵ�� = Format(ActualMoney(NVL(mrsPati!�ѱ�) & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsMain!������ID, rsMain!ҽ���ϼ�), gstrDec)
                End If
                cur�ϼ� = cur�ϼ� - Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                .TextMatrix(rsMain!�����к�, COLP_ʵ�ս��) = Format(Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��)) + (cur��ǰʵ�� - rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� + Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                rsMain.MoveNext
            Loop
        End If
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '��λȱʡ��Ԫ
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_�Ƽ�ҽ�� And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_�Ƽ�ҽ��
        End If
        '��λ�������λ��
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_�Ƽ�ҽ�� And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '���»�����ʾ�ɼ��еķ���ҽ�����
    vsAdvice.TextMatrix(lngRow, COL_���) = Format(cur�ϼ�, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln�Ǳ��� As Boolean) As Boolean
'���ܣ��жϼ۱��е�Ԫ���Ƿ���Ա༭
    Dim lng�к� As Long
    
    With vsPrice
        bln�Ǳ��� = False
        CellEditable = .Editable
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lngCol = COLP_ִ�п��� Then
            '�������õ�����,��ҩ��ҩƷ�Ƽ۵�ִ�п��ҿ����޸�
            If Not ((.TextMatrix(lngRow, COLP_�շ����) = "4" And Val(.TextMatrix(lngRow, COLP_��������)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�շ����)) > 0) And InStr(",4,5,6,7,", vsAdvice.TextMatrix(lng�к�, COL_�������)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Or .TextMatrix(lngRow, COLP_�к�) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_�̶�)) <> 0 Then
            '�̶������н������޸ı��
            If Not (Val(.TextMatrix(lngRow, COLP_���)) = 1 And lngCol = COLP_����) Then
                CellEditable = False
            ElseIf Val(.TextMatrix(lngRow, COLP_���)) = 1 And lngCol = COLP_���� Then
                '�Ǳ���ִ�еı����Ŀ�������۸�
                If lng�к� <> 0 Then
                    If (Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID)))) And InStr(GetInsidePrivs(p����ҽ���´�), "�޸����Ʒ���") = 0 Then
                        bln�Ǳ��� = True: CellEditable = False
                    End If
                End If
            End If
        Else
            If lngCol = COLP_���� Then
                If Val(.TextMatrix(lngRow, COLP_���)) <> 1 Then
                    CellEditable = False
                Else
                    '�Ǳ���ִ�еı����Ŀ�������۸�
                    If lng�к� <> 0 Then
                        If (Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID)))) And InStr(GetInsidePrivs(p����ҽ���´�), "�޸����Ʒ���") = 0 Then
                            bln�Ǳ��� = True: CellEditable = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_�Ƽ�ҽ�� And lngCol <> COLP_�Ƽ����� And lngCol <> COLP_�շ���Ŀ Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Function LoadAdviceSend() As Boolean
'���ܣ�����������ȡ����ʾҪ���͵�ҩƷҽ���嵥
'˵����ע��CellData�д�ŵ��и�������
'   RowData��0-δ���͵�,-1-�ѳɹ����͵�
'   COL_ѡ��0-������ѡ���,1-��ֹ�ı�ѡ��״̬��
'   COL_ID��1-��ҩ;����2-��ҩ�巨��3-��ҩ�÷���4-�ɼ�������5-��Ѫ;��
'   COL_Ӥ�������Ӥ�����
'   COL_������𣺴������������ƣ�������ʾ�Ƽ�ҽ��
'   COL_ҽ�����ݣ����������Ŀ���ƻ�걾��λ��������ʾ�Ƽ�ҽ��
'   COL_�ֽ�ʱ�䣺��ŷ��õķ���ʱ��(�޷ֽ�ʱ��ʱ)
'   COL_Ƶ�ʣ�1-"һ����"����
'   COL_��ԭʼ�Ľ��������ۼ���ʾ��
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDelҽ��ID As Long, lngDel���ID As Long
    Dim bln����ʱ�� As Boolean, lng���� As Long, lng��С���� As Long
    Dim str�ֽ�ʱ�� As String, dbl���� As Double, cur��� As Currency
    
    Dim vMsg As VbMsgBoxResult, strNoneIDs As String
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    Dim str�÷� As String, i As Long, j As Long
    Dim strͣ�� As String
    Dim blnTmp As Boolean
    Dim blnҩƷ������ʾ As Boolean
      
    Screen.MousePointer = 11
    
    stbThis.Panels(3).Text = "": stbThis.Panels(5).Text = "": Call Form_Resize
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '��ɾ���й���
    
    vsAdvice.ColHidden(COL_Ӥ��) = True
    Me.Refresh
    
    Call InitPriceRecordset '�Ƽ۹�ϵ��
    mstrAdDrugIDs = ""
    If mstrǰ��IDs = "" Then
        strNoneIDs = GetNoneSendID(mlng����ID, mstr�Һŵ�, 1, False, mlng�Һ�ID, mstrAdDrugIDs)
    End If
    
    '��ȡ�����嵥:ÿ��ҽ����¼(ҩƷ�ͷ�ҩƷ)
    '----------------------------------------------------------------------------------------------------------
    '����(����,���,���鲻����Ϊ����)������ȼ�������Ƥ�Բ�����,�������ȶ�ȡ����(��ҩ;��,�÷�,�巨,�ɼ�����,��Ѫ;��)
    strSQL = _
        " Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
        " A.�������,F.���� as �������,A.������ĿID,B.���� as ������Ŀ,A.�շ�ϸĿID,C.���,A.Ӥ��," & _
        " A.ҽ������,A.�걾��λ,A.��鷽��,A.ִ�б��,A.����,A.�ܸ�����,D.���ﵥλ,A.��������," & _
        " Decode(A.�������,'4',C.���㵥λ,B.���㵥λ) as ���㵥λ,D.����ϵ��,D.�����װ," & _
        " A.��ʼִ��ʱ��,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ��ʱ�䷽��," & _
        " A.���˿���ID,A.��������ID,A.����ҽ��,A.�Ƽ�����,A.ִ������,A.ִ�п���ID,Nvl(E.����,Decode(Nvl(A.ִ������,0),5,'-')) as ִ�п���," & _
        " D.����ɷ���� As �ɷ����,Decode(A.�������,'4',G.���÷���,D.ҩ������) as ����,C.�Ƿ���,G.��������," & _
        " C.����ʱ��,C.�������,A.ǰ��ID,A.�¿�ǩ��ID as ǩ��ID,B.�Թܱ���,B.��������,b.ִ�з���,A.ժҪ,a.������־,A.��Ѽ���,c.����ʱ��,B.���㷽ʽ,a.��ʼִ��ʱ��,b.ִ�а���,h.�������,a.��ҩ����" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,�շ���ĿĿ¼ C,ҩƷ��� D,���ű� E,������Ŀ��� F,�������� G,ҩƷ���� H,����ҽ����¼ X" & _
        " Where A.����ID+0=[1] And A.�Һŵ�=[2] And Nvl(A.ǰ��ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([4]) As zlTools.t_Numlist)) X) " & _
        " And A.ҽ��״̬=1 And A.ҽ����Ч=1 And A.���ID=X.ID(+) And B.���=F.����" & _
        " And A.������ĿID=B.ID And A.�շ�ϸĿID=C.ID(+) And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=G.����ID(+)" & _
        " And A.ִ�п���ID=E.ID(+) And Not (A.�������='H' And B.��������='1') And NVL(A.ִ�б��,0)<>-1 and b.id=h.ҩ��id(+) " & _
        IIF(gblnKSSStrict Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ, " And Nvl(A.���״̬,0) Not in " & IIF(gblnѪ��ϵͳ = True, " (1,3,7)", " (1,3,4,5,7)"), "") & _
        IIF(strNoneIDs <> "" And Not mbln������ҩ, " And Instr([3],','||A.ID||',')=0", "") & IIF(mint�������� = 0, " And A.��Ѽ��� Is Null", "") & _
        " And Nvl(A.Ƥ�Խ��,'��')<>'����' And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & IIF(mint���� = 2, " And A.����ҽ��=[5]", "") & _
        " Order by A.Ӥ��,���,��ID,A.���"
    
    On Error GoTo errH
    
    Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, "," & strNoneIDs & ",", IIF(mstrǰ��IDs = "", "0", mstrǰ��IDs), UserInfo.����)
    
    '���㲢��ʾ�����嵥
    '----------------------------------------------------------------------------------------------------------
    If Not rsSend.EOF Then
        With vsAdvice
            blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
            bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
            blnҩƷ������ʾ = True
            .Redraw = flexRDNone
            For i = 1 To rsSend.RecordCount
                'һ����ҩ���䷽���������е�һ�������Ѿ����ܷ���,�����鲻�ܷ���
                If lngDel���ID <> 0 Then
                    If (rsSend!ID = lngDel���ID Or NVL(rsSend!���ID, 0) = lngDel���ID) Then
                        GoTo NextLoop
                    Else
                        lngDel���ID = 0
                    End If
                End If
                '�����ϻ���������е�һ�������Ѿ����ܷ���,�����鲻�ܷ���
                If lngDelҽ��ID <> 0 Then
                    If NVL(rsSend!���ID, 0) = lngDelҽ��ID Then
                        GoTo NextLoop
                    Else
                        lngDelҽ��ID = 0
                    End If
                End If
                                                
                '���뵱ǰ��
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .Cell(flexcpPictureAlignment, lngRow, COL_ѡ��) = 4
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                
                '�����ͣ�õģ�����ʾ���ܷ���
                If Format(NVL(rsSend!����ʱ��, "3000-1-1"), "YYYY-MM-DD") <> Format("3000-1-1", "YYYY-MM-DD") Then
                    .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                    Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    If InStr(strͣ�� & ",", "," & rsSend!ҽ������ & ",") = 0 Then strͣ�� = strͣ�� & "," & rsSend!ҽ������
                End If
                
                '���������
                If rsSend!������� = "7" Then
                    .RowHidden(lngRow) = True '�в�ҩ
                ElseIf rsSend!������� = "E" Then
                    If Not IsNull(rsSend!���ID) Then
                        .RowHidden(lngRow) = True
                        If .TextMatrix(lngRow - 1, COL_�������) = "K" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 5 '��Ѫ;��
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 2 '��ҩ�巨
                        End If
                    ElseIf Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID Then
                        If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_ID) = 1 '��ҩ;��
                        ElseIf .TextMatrix(lngRow - 1, COL_�������) = "C" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 4 '�ɼ�����
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 3 '��ҩ�÷�
                        End If
                    End If
                ElseIf InStr(",5,6,", rsSend!�������) = 0 And Not IsNull(rsSend!���ID) Then
                    '��������,��������,��鲿λ,һ���ɼ��ļ�����Ŀ
                    .RowHidden(lngRow) = True
                End If
                
                '�ſ�һ��Ķ���(������ҩ;��,��ҩ�巨,�÷�,�ɼ�����,��Ѫ;��)
                If NVL(rsSend!ִ������, 0) = 0 Then
                    If InStr(",1,2,3,4,5,", CLng(.Cell(flexcpData, lngRow, COL_ID))) = 0 _
                        And InStr(",5,6,7,", rsSend!�������) = 0 Then
                        Call .RemoveItem(lngRow): GoTo NextLoop
                    End If
                End If
                
                'һ���и�ֵ
                '---------------------------------------------------------------
                .Cell(flexcpData, lngRow, COL_Ӥ��) = CLng(NVL(rsSend!Ӥ��, 0))
                If NVL(rsSend!Ӥ��, 0) = 0 Then
                    .TextMatrix(lngRow, COL_Ӥ��) = "����"
                Else
                    .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsSend!Ӥ��
                    .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
                End If
                
                .TextMatrix(lngRow, COL_ID) = rsSend!ID
                .TextMatrix(lngRow, COL_���ID) = NVL(rsSend!���ID)
                .TextMatrix(lngRow, COL_�������) = rsSend!�������
                .TextMatrix(lngRow, COL_������ĿID) = rsSend!������ĿID
                .TextMatrix(lngRow, col_ҽ������) = NVL(rsSend!ҽ������)
                .TextMatrix(lngRow, COL_ǰ��ID) = NVL(rsSend!ǰ��ID)
                
                .TextMatrix(lngRow, COL_�걾��λ) = NVL(rsSend!�걾��λ)
                .TextMatrix(lngRow, COL_��鷽��) = NVL(rsSend!��鷽��)
                .TextMatrix(lngRow, COL_ִ�б��) = NVL(rsSend!ִ�б��, 0)
                .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������)
                .TextMatrix(lngRow, COL_������־) = NVL(rsSend!������־, 0)
                .TextMatrix(lngRow, COL_��Ѽ���) = NVL(rsSend!��Ѽ���, 0)
                If InStr(",4,5,6,7,", "," & rsSend!������� & ",") = 0 Then .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsSend!���㷽ʽ, 0)
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(NVL(rsSend!��ʼִ��ʱ��), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(lngRow, COL_ִ�а���) = NVL(rsSend!ִ�а���, "")
                .TextMatrix(lngRow, COL_ִ�з���) = NVL(rsSend!ִ�з���, "")
                .TextMatrix(lngRow, COL_��ҩ����) = NVL(rsSend!��ҩ����)
                '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    j = .FindRow(CStr(rsSend!ID), .FixedRows, COL_���ID)
                    If j <> -1 Then
                        .TextMatrix(lngRow, COL_�Թܱ���) = .TextMatrix(j, COL_�Թܱ���)
                    End If
                Else
                    .TextMatrix(lngRow, COL_�Թܱ���) = NVL(rsSend!�Թܱ���)
                End If
                
                '����ǩ����ʶ
                .TextMatrix(lngRow, COL_ǩ��ID) = NVL(rsSend!ǩ��ID)
                If Val(.TextMatrix(lngRow, COL_ǩ��ID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgSign.ListImages("ǩ��").Picture
                End If
                
                '������ʾ�Ƽ�ҽ����
                .Cell(flexcpData, lngRow, COL_�������) = CStr(NVL(rsSend!�������))
                .Cell(flexcpData, lngRow, col_ҽ������) = CStr(NVL(rsSend!������Ŀ))
                .Cell(flexcpData, lngRow, COL_�շ�ϸĿID) = CStr(NVL(rsSend!���))
                
                .TextMatrix(lngRow, COL_ҽ������) = NVL(rsSend!ҽ������)
                .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(NVL(rsSend!ժҪ))
                
                .TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsSend!ִ��ʱ�䷽��)
                .TextMatrix(lngRow, COL_Ƶ��) = NVL(rsSend!ִ��Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = NVL(rsSend!Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = NVL(rsSend!Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = NVL(rsSend!�����λ)
                
                .TextMatrix(lngRow, COL_���˿���ID) = NVL(rsSend!���˿���id)
                .TextMatrix(lngRow, COL_��������ID) = NVL(rsSend!��������id)
                .TextMatrix(lngRow, COL_����ҽ��) = NVL(rsSend!����ҽ��)
                                
                '�ɼ���������ʾ������Ŀ��ִ�п���
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngRow - 1, COL_ִ�п���)
                Else
                    .TextMatrix(lngRow, COL_ִ�п���) = NVL(rsSend!ִ�п���)
                End If
                .TextMatrix(lngRow, COL_ִ�п���ID) = NVL(rsSend!ִ�п���ID)
                
                .TextMatrix(lngRow, COL_�Ƽ�����) = NVL(rsSend!�Ƽ�����, 0)
                .TextMatrix(lngRow, COL_ִ������ID) = NVL(rsSend!ִ������, 0)
                                
                'ҩƷ�����Ϣ
                If InStr(",5,6,7", rsSend!�������) > 0 Then
                    'ҩƷ��Ӧ�Ĺ���ѳ�����������(������Ŀ����Ҳ������ͬ����,Ŀǰ��δ����)
                    If Format(NVL(rsSend!����ʱ��, "3000-01-01"), "yyyy-MM-dd") <> "3000-01-01" Or InStr(",1,3,", NVL(rsSend!�������, 0)) = 0 Then
                        If rsSend!������� = "7" Then
                            strTmp = "���в�ҩ��Ӧ����ҩ�䷽�޷����ͣ�" & vbCrLf & vbCrLf & "����" & NVL(rsSend!ҽ������)
                        Else
                            strTmp = "��ҩƷ(��һ����ҩ������ҩƷ)�޷����ͣ�" & vbCrLf & vbCrLf & "����" & NVL(rsSend!ҽ������)
                        End If
                        strTmp = strTmp & vbCrLf & vbCrLf & "û�з�����Ч��ҩƷ�����Ϣ����ҩƷ�����Ѿ���ͣ�û����������ﲡ�ˡ�"
                        strTmp = strTmp & vbCrLf & "���ȵ�ҩƷĿ¼�����д�����[ȷ��]������������ҽ����"
                        
                        .Redraw = flexRDDirect
                        Call .ShowCell(lngRow, COL_ѡ��)
                        Screen.MousePointer = 0
                        MsgBox strTmp, vbInformation, gstrSysName
                        
                        'ɾ����ǰ��(�������),��������һҽ��
                        Screen.MousePointer = 11
                        lngDelҽ��ID = rsSend!ID
                        lngDel���ID = NVL(rsSend!���ID, 0)
                        Call DeleteCurRow(lngRow)
                        .Refresh: .Redraw = flexRDNone
                        lng��С���� = 0: GoTo NextLoop
                    End If
                    
                    '��������ж�
                    If gbln����ҩƷ�ֿ����� Then
                        strTmp = ""
                        Select Case cboDrugType.ListIndex
                        Case 1
                            If rsSend!������� & "" <> "����ҩ" Then strTmp = "1"
                        Case 2
                            If InStr(",����ҩ,����I��,", "," & rsSend!������� & ",") = 0 Then strTmp = "1"
                        Case 3
                            If InStr(",����ҩ,����ҩ,����I��,", "," & rsSend!������� & ",") > 0 Then strTmp = "1"
                        End Select
                        
                        If strTmp <> "" Then
                            lngDelҽ��ID = rsSend!ID
                            lngDel���ID = rsSend!���ID
                            Call DeleteCurRow(lngRow)
                            lng��С���� = 0: GoTo NextLoop
                        End If
                        .TextMatrix(lngRow, COL_�������) = NVL(rsSend!�������, "��")
                    End If
                
                    .TextMatrix(lngRow, COL_�շ�ϸĿID) = rsSend!�շ�ϸĿID
                    .TextMatrix(lngRow, COL_����ϵ��) = NVL(rsSend!����ϵ��, 1)
                    .TextMatrix(lngRow, COL_�����װ) = NVL(rsSend!�����װ, 1)
                    .TextMatrix(lngRow, COL_���ﵥλ) = NVL(rsSend!���ﵥλ)
                    .TextMatrix(lngRow, COL_�ɷ����) = NVL(rsSend!�ɷ����, 0)
                    .TextMatrix(lngRow, COL_���) = GetStock(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), 1) '�������װ
                ElseIf rsSend!������� = "4" Then
                    .TextMatrix(lngRow, COL_�շ�ϸĿID) = rsSend!�շ�ϸĿID
                    .TextMatrix(lngRow, COL_����ϵ��) = 1
                    .TextMatrix(lngRow, COL_�����װ) = 1
                    .TextMatrix(lngRow, COL_���ﵥλ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_���) = GetStock(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), 1)
                End If
                                                                        
                '���㷢�ʹ�����ִ�еķֽ�ʱ���
                '---------------------------------------------------------------
                If rsSend!������� = "7" Then
                    .TextMatrix(lngRow, COL_����) = rsSend!�ܸ�����
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Calc�����ֽ�ʱ��(rsSend!�ܸ�����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(rsSend!�ܸ����� - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!��������) '����
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = rsSend!�ܸ����� '����
                    .TextMatrix(lngRow, COL_������λ) = "��"
                ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                    '����������ҩ����
                    If NVL(rsSend!����, 0) <> 0 And Not IsNull(rsSend!ִ��Ƶ��) Then
                        'һ��Ƶ�����ڵĴ���
                        If rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / 7))
                        ElseIf rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��))
                        ElseIf rsSend!�����λ = "Сʱ" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * 24)
                        ElseIf rsSend!�����λ = "����" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * (24 * 60))
                        End If
                    Else
                        '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,���ɷ�����һ����ʹ��ҩƷʱ���������ԣ����������ϵ����ֵȡ�����ı��������ҩ;���Ĵ�����
                        '����һ��Ƶ�����ڵĴ�������
                        If NVL(rsSend!�ɷ����, 0) = 0 And NVL(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� * rsSend!����ϵ�� / rsSend!��������)
                        ElseIf (NVL(rsSend!�ɷ����, 0) = 1 Or NVL(rsSend!�ɷ����, 0) = 2) And NVL(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� / IntEx(rsSend!�������� / rsSend!����ϵ��))
                        Else
                            lng���� = NVL(rsSend!Ƶ�ʴ���, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!Ƶ�ʴ���) And (Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����") Then
                        str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        If str�ֽ�ʱ�� <> "" Then
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = FormatEx(rsSend!�ܸ����� / rsSend!�����װ, 5) '�����ﵥλ��ʾ
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���ﵥλ)
                    
                    If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                ElseIf rsSend!������� = "E" And CLng(.Cell(flexcpData, lngRow, COL_ID)) <> 0 Then
                    '��ҩ;��,��ҩ�巨,��ҩ�÷�,�ɼ�����,��Ѫ;��
                    'һ����ҩ�İ���С��������(Ӱ���ҩ;���Ʒ�)
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '��ҩ;��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                    .TextMatrix(j, COL_����) = lng��С����
                                    If .TextMatrix(j, COL_�ֽ�ʱ��) <> "" Then
                                        .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                        .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                                        .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "yyyy-MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng��С���� = 0
                    End If
                    
                    blnTmp = False
                    If Val(rsSend!�ܸ����� & "") <> 0 Then
                        If Val(rsSend!�ܸ����� & "") < Val(.TextMatrix(lngRow - 1, COL_����)) Then
                            If .TextMatrix(lngRow, COL_��������) = "2" And (.TextMatrix(lngRow, COL_ִ�з���) = "1" Or .TextMatrix(lngRow, COL_ִ�з���) = "2") Then
                                blnTmp = True
                            End If
                        End If
                    End If
                    
                    If blnTmp Then
                        .TextMatrix(lngRow, COL_����) = Val(rsSend!�ܸ����� & "")
                        .TextMatrix(lngRow, COL_����) = Val(rsSend!�ܸ����� & "")
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(Val(rsSend!�ܸ����� & ""), .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��))
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(Val(rsSend!�ܸ����� & "") - 1), "yyyy-MM-dd HH:mm")
                    Else
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����) '���������
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                    End If
                    
                    .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                    If .Cell(flexcpData, lngRow, COL_ID) = 3 Then '��ҩ�÷�
                        .TextMatrix(lngRow, COL_������λ) = "��"
                    Else
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    End If
                Else
                    '������ҩ����:�ɼ�����������ķ�֧����������
                    If IsNull(rsSend!���ID) Or (Not IsNull(rsSend!���ID) And rsSend!������� = "C") Then '��Ҫҽ��,�����������
                        If rsSend!������� = "K" Then
                            '��Ѫ;����ִ�д���
                            dbl���� = NVL(rsSend!�ܸ�����, 0)
                            If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                                lng���� = 1
                            Else
                                lng���� = NVL(rsSend!Ƶ�ʴ���, 1)
                            End If
                        Else
                            dbl���� = NVL(rsSend!�ܸ�����, 1)
                            lng���� = IntEx(dbl���� / NVL(rsSend!��������, 1))
                        End If
                        
                        If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                            'ִ��Ƶ��Ϊ"һ����"����Ŀ
                            str�ֽ�ʱ�� = "" '����Ҫ
                            .Cell(flexcpData, lngRow, COL_Ƶ��) = 1
                        Else
                            'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ:��ҽ��ʱӦ����������
                            If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                                str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                            Else
                                str�ֽ�ʱ�� = "" '����Ҳ��δ����ִ��ʱ��,�޷��ֽ�
                            End If
                        End If
                        .TextMatrix(lngRow, COL_����) = lng����
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                        If str�ֽ�ʱ�� <> "" Then
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                        Else
                            '��¼���÷���ʱ��(���޷ֽ�ʱ��ʱ),��ҽ���Ŀ�ʼִ��ʱ��
                            .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = CStr(Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
                        End If
                        
                        .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                        If Not IsNull(rsSend!��������) Then
                            .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                        End If
                        .TextMatrix(lngRow, COL_����) = IIF(dbl���� = 0, "", FormatEx(dbl����, 5))
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    Else
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                    End If
                End If
                
                '������Ŀ���ͽ��
                cur��� = 0
                If Not LoadAdvicePrice(lngRow, rsSend, cur���) Then
                    lngDelҽ��ID = rsSend!ID
                    lngDel���ID = rsSend!���ID
                    Call DeleteCurRow(lngRow)
                    lng��С���� = 0: GoTo NextLoop
                End If
                .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                .Cell(flexcpData, lngRow, COL_���) = CCur(.TextMatrix(lngRow, COL_���))
                
                '�����ʱ��һЩ�����ۼ���ʾ���,��ҩ;��,�÷�,ִ�п���,ִ������
                '---------------------------------------------------------------
                If rsSend!������� = "E" And InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_ID))) > 0 Then '��ҩ;������ҩ�÷�
                    cur��� = 0
                    lngTmp = .FindRow(CStr(rsSend!ID), , COL_���ID)
                    
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '��ҩ;��
                        'һ����ҩʱ,��ҩ;���Ľ���ۼ���ʾ�ڵ�һ����ҩ��
                        .TextMatrix(lngTmp, COL_���) = Format(Val(.TextMatrix(lngTmp, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                        '��ʾ��ҩ;��,ִ������
                        For j = lngTmp To lngRow - 1
                            strTmp = ""
                            If Val(.TextMatrix(j, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                                strTmp = "�Ա�ҩ"
                            ElseIf Val(.TextMatrix(j, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                                strTmp = "��Ժ��ҩ"
                            End If
                            .TextMatrix(j, COL_ִ������) = strTmp
                            .TextMatrix(j, COL_�÷�) = rsSend!������Ŀ
                        Next
                    Else
                        'ҩƷ��ִ������
                        strTmp = ""
                        If Val(.TextMatrix(lngTmp, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                            strTmp = "�Ա�ҩ"
                        ElseIf Val(.TextMatrix(lngTmp, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                            strTmp = "��Ժ��ҩ"
                        End If
                    
                        '��ҩ�÷�,�巨
                        str�÷� = rsSend!������Ŀ
                        If Val(.Cell(flexcpData, lngRow - 1, COL_ID)) = 2 Then
                            str�÷� = str�÷� & "|" & Sys.RowValue("������ĿĿ¼", Val(.TextMatrix(lngRow - 1, COL_������ĿID)), "����")
                        End If
                        For j = lngTmp To lngRow
                            .TextMatrix(j, COL_�÷�) = str�÷� '������д�շ���¼
                            cur��� = cur��� + Val(.TextMatrix(j, COL_���))
                        Next
                        .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                        '��ʾִ������
                        .TextMatrix(lngRow, COL_ִ������) = strTmp
                        '��ʾ�䷽ִ�п���
                        .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngTmp, COL_ִ�п���)
                    End If
                    
                    'ʹ���ҽ��ѡ��״̬��ͬ(��Ϊ����ԭ�򣻷�ҩҽ������)
                    For j = lngTmp To lngRow
                        If .Cell(flexcpData, j, COL_ѡ��) <> 0 Then
                            Call RowSelectSame(j, COL_ѡ��)
                            Exit For 'һ����ֹ,ȫ����ֹ
                        End If
                    Next
                    If j > lngRow Then
                        For j = lngRow To lngTmp Step -1
                            If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 Then
                                If .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                    Call RowSelectSame(j, COL_ѡ��)
                                    Exit For '���ѡ,ȫ����ѡ
                                End If
                            End If
                        Next
                    End If
                ElseIf InStr(",5,6,7,", rsSend!�������) = 0 Then
                    If Not IsNull(rsSend!���ID) And rsSend!������� <> "C" Then
                        '������ҩҽ��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_ID)) = rsSend!���ID Then
                                .TextMatrix(j, COL_���) = Format(Val(.TextMatrix(j, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                                Exit For
                            End If
                        Next
                        
                        '��Ѫ;��
                        If rsSend!������� = "E" And Val(.Cell(flexcpData, lngRow, COL_ID)) = 5 Then
                            .TextMatrix(lngRow - 1, COL_�÷�) = rsSend!������Ŀ
                        End If
                    ElseIf Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                        '����걾�ɼ�����Ϊ��ʾ��
                        .TextMatrix(lngRow, COL_�÷�) = rsSend!������Ŀ
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                .TextMatrix(lngRow, COL_���) = Format(Val(.TextMatrix(lngRow, COL_���)) + Val(.TextMatrix(j, COL_���)), gstrDec)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If

                'ҩƷ�����Ŀ����(0-�����;1-���,��������;2-��飬�����ֹ),�Ա�ҩ�����
                '---------------------------------------------------------------
                If InStr(",5,6,7,", rsSend!�������) > 0 And NVL(rsSend!ִ������, 0) <> 5 Then
                    Call CheckStock(lngRow, rsSend, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                    Call CheckDrug����(lngRow, blnҩƷ������ʾ)
                ElseIf rsSend!������� = "4" And NVL(rsSend!��������, 0) = 1 Then
                    Call CheckStock(lngRow, rsSend, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                End If

NextLoop:       '---------------------------------------------------------------
                Progress = i / rsSend.RecordCount * 100
                rsSend.MoveNext
            Next
        End With
        
        '���Һ���Ч������������������Ϊ�շѵ�
        Call ExpendSendClear(mstr�Һŵ�)
    End If
    
    With vsAdvice
        .AutoSize col_ҽ������
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        '����ǩ��ͼ�����
        .Cell(flexcpPictureAlignment, .FixedRows, col_ҽ������, .Rows - 1, col_ҽ������) = 0
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        '�����ͣ�õ���Ŀ������ʾ
        If strͣ�� <> "" Then
            Call MsgBox("������Ŀ��" & Mid(strͣ��, 2) & " �Ѿ�ͣ�ã����ܷ��͡�", vbInformation, Me.Caption)
        End If
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Sub CheckDrug����(ByVal lngRow As Long, ByRef bln��ʾ As Boolean)
'���ܣ����͹����ж�����ҩƷ���м���ֹ
    Dim strTmp As String
    Dim blnTmp As Boolean
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        If 0 <> Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) And 0 <> Val(.TextMatrix(lngRow, COL_ִ�п���ID)) And .Cell(flexcpData, lngRow, COL_ѡ��) <> 1 Then
            If InitObjPublicDrug Then
                blnTmp = gobjPublicDrug.zlCheckPriceAdjustBySell(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), False)
                If Not blnTmp Then
                    strTmp = "��(" & .TextMatrix(lngRow, COL_ִ�п���) & ")��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """" & vbCrLf & vbCrLf & _
                        "���������۹����Ҫ�󣺳ɱ��ۺ��ۼ۲�һ�£��������۳��⡣" & vbCrLf & vbCrLf & _
                        "����ϵҩ����ҩ���ƽ��е��۴���"
                    
                    If bln��ʾ Then
                        .Redraw = flexRDDirect:
                        Call .ShowCell(lngRow, COL_ѡ��)
                        Screen.MousePointer = 0
                        vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
                        If vMsg = vbIgnore Then bln��ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        Screen.MousePointer = 11
                        .Refresh: .Redraw = flexRDNone
                    Else
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub CheckStock(ByVal lngRow As Long, rsSend As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean)
'���ܣ����ݿ���������鷢��ҩƷ���������ĵĿ��
'������lngRow=ҽ���к�,rsSend=��ǰ����ҽ����Ϣ
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = TheStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID)), .TextMatrix(lngRow, COL_�������))
        bln���� = NVL(rsSend!����, 0) = 1
        blnʱ�� = NVL(rsSend!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = .TextMatrix(lngRow, COL_���ﵥλ) '������ɢװ��λ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����:�����װ
            If .TextMatrix(lngRow, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����))
                    dbl���� = dbl���� / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_�����װ))
                Else
                    dbl���� = IntEx(Val(.TextMatrix(lngRow, COL_����)) / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_�����װ)))
                    dbl���� = dbl���� * Val(.TextMatrix(lngRow, COL_����))
                End If
            Else
                dbl���� = Val(.TextMatrix(lngRow, COL_����))
            End If
            
            '��ǰ���ÿ��:�����װ,��ȥǰ����ͬҩƷҪ���͵Ŀ��
            For i = lngRow - 1 To .FixedRows Step -1
                If rsSend!������� = "4" Then
                    blnDo = .TextMatrix(i, COL_�������) = "4"
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                End If
                If blnDo Then
                    blnDo = Val(.TextMatrix(i, COL_�շ�ϸĿID)) = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) _
                        And Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                End If
                If blnDo Then
                    blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                End If
                If blnDo Then
                    If .TextMatrix(i, COL_�������) = "7" Then
                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            dbl�ѷ���� = dbl�ѷ���� + _
                                Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ))
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ)))
                        End If
                    Else
                        dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                    End If
                End If
            Next
            dbl���ÿ�� = Val(.TextMatrix(lngRow, COL_���))
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "������ʱ��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��" & vbCrLf & vbCrLf & _
                                "��" & .TextMatrix(lngRow, COL_ִ�п���) & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "������ʱ��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��治�㣺" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    Else
                        If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��" & vbCrLf & vbCrLf & _
                                "��" & .TextMatrix(lngRow, COL_ִ�п���) & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��治�㣺" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҩƷ��"
                    End If
                    If rsSend!������� = "4" Then
                        strTmp = Replace(strTmp, "ҩƷ", "����")
                    End If
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                        End If
                    End If
                    
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    rsTotal As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean) As Boolean
'���ܣ����͹�����ʱ���Է�ҩ��ҩƷ���������õ����ļƼ۽��п����(�ۼƼ��)
'������lngRow=ҽ���к�
'      dbl����=�Ѽ���õļƼ�����(�ۼ۵�λ)
'      rsTotal=��ǰ����ǰ�����ۼƷ��͵ļƼ�ҩƷ����������(�ۼ۵�λ)
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = TheStockCheck(lng�ⷿID, rsPrice!���)
        bln���� = NVL(rsPrice!����, 0) = 1
        blnʱ�� = NVL(rsPrice!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = NVL(rsPrice!���ﵥλ, NVL(rsPrice!���㵥λ)) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����������:�����װ
            dbl���� = Format(dbl���� / NVL(rsPrice!�����װ, 1), "0.00000")
            
            '��ǰ���ÿ��:�����װ,��ȥǰ����ͬҩƷҽ��Ҫ���͵Ŀ��
            If InStr(",5,6,7,", rsPrice!���) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_�շ�ϸĿID)) = rsPrice!ID And Val(.TextMatrix(i, COL_ִ�п���ID)) = lng�ⷿID
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_�������) = "7" Then
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                dbl�ѷ���� = dbl�ѷ���� + _
                                    Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                    / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ))
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                    * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ)))
                            End If
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                Next
            End If
            '�Ƽ۲���Ҫ���͵��ۼ�����
            rsTotal.Filter = "��ĿID=" & rsPrice!ID & " And �ⷿID=" & lng�ⷿID
            Do While Not rsTotal.EOF
                dbl�ѷ���� = dbl�ѷ���� + Format(rsTotal!���� / NVL(rsPrice!�����װ, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl���ÿ�� = Format(GetStock(rsPrice!ID, lng�ⷿID, 2), "0.00000")
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ��" & vbCrLf & vbCrLf & _
                                """" & rsPrice!���� & """��" & Sys.RowValue("���ű�", lng�ⷿID, "����") & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                                vbCrLf & vbCrLf & Sys.RowValue("���ű�", lng�ⷿID, "����") & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    Else
                        If InStr(GetInsidePrivs(p����ҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ļƼ���Ŀ��" & vbCrLf & vbCrLf & _
                                """" & rsPrice!���� & """��" & Sys.RowValue("���ű�", lng�ⷿID, "����") & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ļƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                                vbCrLf & vbCrLf & Sys.RowValue("���ű�", lng�ⷿID, "����") & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҽ����"
                    End If
                    
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '���δ��ʾ��Ҫ����,�����ۼƷ�������
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_���ID))
            Else
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!��ĿID = rsPrice!ID
            rsTotal!�ⷿID = lng�ⷿID
            rsTotal!���� = dbl����
            rsTotal.Update
        End If
    End With
End Function

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, blnCancel As Boolean
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim int�������� As Integer, vPoint As POINTAPI
    Dim strSQL2 As String
    
    With vsPrice
        lng�к� = Val(.TextMatrix(Row, COLP_�к�))
        If Col = COLP_�շ���Ŀ Then
            '����ѡ�����е���Ŀ
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_�к�)) = lng�к� And lng�к� <> 0 And i <> Row Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                End If
            Next
            str��ĿIDs = Mid(str��ĿIDs, 2)
            
            strSQL = _
                " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
                " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',7,'��������') as ����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,-ID as ID,Nvl(-�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL2 = _
                " Select ĩ��,ID,�ϼ�ID,����,����,��λ,���,����,���,��������,ҽ������,˵��," & _
                " Decode(Nvl(�Ƿ���,0),1,Decode(Instr('567',���ID),0,Sum(Nvl(ԭ��,0))||'-'||Sum(Nvl(�ּ�,0)),'ʱ��'),Sum(�ּ�)) as �۸�," & _
                " Sum(ԭ��) as ԭ��ID,Sum(�ּ�) as �ּ�ID,Sum(ȱʡ�۸�) as ȱʡ�۸�ID,�Ƿ��� as �Ƿ���ID,���ID,��������ID" & _
                " From (" & _
                " Select Distinct 1 as ĩ��,A.ID,Decode(Instr('567',A.���),0,A.����ID,-E.����ID) as �ϼ�ID,A.����,A.����," & _
                " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���," & _
                " A.��� as ���ID,-Null as ��������ID" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                " Where A.ID=B.�շ�ϸĿID  [ѡ���滻�Ĺ�����1] And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.������� IN(1,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.��� Not IN('4','J','1') And A.���=C.���� And A.ID=D.ҩƷID(+) And D.ҩ��ID=E.ID(+)" & _
                " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=1 or (w.������Դ is Null And Nvl(w.��������id,[3]) = [3]))))" & _
                " And (a.��� Not in ('5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[3])=[3]))"
            If DeptExist("���ϲ���", 1) Then
                strSQL2 = strSQL2 & " Union ALL " & _
                    " Select Distinct 1 as ĩ��,A.ID,-E.����ID as �ϼ�ID,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��," & _
                    " B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���,A.��� as ���ID,D.�������� as ��������ID" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,�������� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                    " Where A.ID=B.�շ�ϸĿID  [ѡ���滻�Ĺ�����2] And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(1,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.���='4' And A.���=C.���� And A.ID=D.����ID And D.����ID=E.ID" & _
                    " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                    " And Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[3])=[3])"
            End If
            strSQL2 = strSQL2 & " ) Group by ĩ��,ID,�ϼ�ID,���,����,����,��λ,���,����,��������,ҽ������,˵��,�Ƿ���,���ID,��������ID"
            '[ѡ���滻�Ĺ�����1],[ѡ���滻�Ĺ�����2],����������ѡ���д����
            'Ҫȷ�� "ռλ����" �����һλ���ò�����ѡ������ƴ�ӣ�Ҫ���4000���ȵ�����
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "�շ���Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str��ĿIDs & ",", mint����, mlng�������ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "ռλ����")
            If Not rsTmp Is Nothing Then
                '�Ǳ���ִ�е�ҽ����������������Ŀ
                If lng�к� <> 0 Then
                    If NVL(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And NVL(rsTmp!��������ID, 0) = 1) Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                'ҽ��������
                If CheckItemInsure(rsTmp) Then
                    .SetFocus: Exit Sub
                End If
                
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                If lng�к� <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ��շ���Ŀ�����ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_ִ�п��� Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_�շ����) = "4" Then
                '�������õ�����
                strSQL = _
                    " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                    " And B.������� IN(1,3) And B.����ID=C.ID" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (A.������Դ is NULL Or A.������Դ=1)" & _
                    " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And A.�շ�ϸĿID=[1]" & _
                    " Order by B.�������,C.����"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                'ҩƷ
                'ҩƷ��ϵͳָ���Ĵ���ҩ������
                If Not Check�ϰల��(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(1,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=1)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(1,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=1)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                    Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!ִ�п���ID = rsTmp!ID
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset) As Boolean
'���ܣ��������(ѡ��)�Ƽ���Ŀ�Ƿ�ҽ������
'���أ����δ���룬������ʾѡ�񲻼������򷵻��档
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int���� As Integer
    
    If gintҽ������ = 0 Then Exit Function
    
    On Error GoTo errH

    strSQL = "Select ���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", mlng����ID)
    If Not rsTmp.EOF Then int���� = NVL(rsTmp!����, 0)
    If int���� <> 0 Then
        If Not ItemExistInsure(mlng����ID, rsInput!ID, int����) Then
            If gintҽ������ = 1 Then
                If MsgBox("��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gintҽ������ = 2 Then
                MsgBox "��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_�Ƽ�ҽ�� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_�̶�)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_�к�)) <> 0 And Val(.TextMatrix(.Row, COLP_�շ�ϸĿID)) <> 0 Then
                    'ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & _
                        " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_�Ƽ�ҽ��) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("ȷʵҪɾ����ǰ�Ƽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & _
                        " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And �շ�ϸĿID=" & Val(.TextMatrix(.Row, COLP_�շ�ϸĿID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_�շ���Ŀ Or .Col = COLP_ִ�п���) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ���۱�����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    
    With vsPrice
        '��ǰ��Ԫ�����δ��������,���˳�
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_���� And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_�Ƽ�ҽ��) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '��ǰ�����û���ҵ���һ���ɱ༭��Ԫ,�������Ƽ�ҽ��,������һ����
            If CStr(.ColData(COLP_�Ƽ�ҽ��)) <> "" Then
                '��ǰ��δ��������,��λ����������Ԫ
                If .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = "" Then
                    .Col = COLP_�Ƽ�ҽ��
                ElseIf .TextMatrix(lngRow, COLP_�Ƽ�����) = "" Then
                    .Col = COLP_�Ƽ�����
                ElseIf .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Then
                    .Col = COLP_�շ���Ŀ
                ElseIf Val(.TextMatrix(lngRow, COLP_���)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_����)) = 0 _
                    And CellEditable(lngRow, COLP_����) Then
                    .Col = COLP_����
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_�Ƽ�ҽ��
                    
                    'ȱʡѡ��Ƽ�ҽ��(�������)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '���ɱ༭ʱ���ⶨһ��
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, int�������� As Integer
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As POINTAPI
    Dim strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng�к� = Val(.TextMatrix(Row, COLP_�к�))
            If Col = COLP_�Ƽ�ҽ�� Then
                '����ʱ�س�
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '��ȻEnterNextCell����Ҫ�˳�
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_�Ƽ����� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�Ƽ�����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շѵ���������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��������뷶Χ
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_�շ���Ŀ And .EditText <> "" Then
                '����ѡ�����е���Ŀ
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_�к�)), COL_ID)) = Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) <> 0 And i <> Row Then
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                    End If
                Next
                str��ĿIDs = Mid(str��ĿIDs, 2)
                
                If mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng���ϲ��� <> 0 Then
                    strStock = _
                        "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ���" & _
                        " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                        " Where A.���� = 1 And (Nvl(A.����,0)=0 Or A.Ч�� Is Null Or A.Ч��>Trunc(Sysdate))" & _
                        " And A.�ⷿID=Decode(B.���,'5',[7],'6',[8],'7',[9],'4',[10],Null)" & _
                        " And A.ҩƷID=B.ID And B.��� IN('4','5','6','7')" & _
                        " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
                Else
                    strStock = "Select Null as ҩƷID,Null as ��� From Dual"
                End If
                
                '��ͬ������ƥ�䷽ʽ
                strInput = UCase(.EditText)
                strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
                ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strMatch = " And C.���� Like [2] And C.����=[3]"
                End If
                
                strSQL = ""
                If Not DeptExist("���ϲ���", 1) Then strSQL = " And A.���<>'4'"
                
                strSQL = _
                    " Select A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����," & _
                    " Decode(Instr('4567',A.���ID),0,NULL,1," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���,'999990.0000'))||A.��λ)," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���/Nvl(C.�����װ,1),'999990.0000'))||C.���ﵥλ)) as ���," & _
                    "   A.��������,N.���� as ҽ������,A.˵��," & _
                    " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(Nvl(A.ԭ��,0))||'-'||Sum(Nvl(A.�ּ�,0)),'ʱ��'),Sum(A.�ּ�)) as �۸�," & _
                    " Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,Sum(A.ȱʡ�۸�) as ȱʡ�۸�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,B.�������� as ��������ID" & _
                    " From (" & _
                    " Select Distinct 1 as ĩ��,A.ID,a.ִ�п���,A.��� as ���ID,D.���� as ���,A.����,A.����,A.���㵥λ as ��λ," & _
                    " A.���,A.����,A.��������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ���� C,�շ���Ŀ��� D" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "11", "12", "13") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(1,3)" & IIF(str��ĿIDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.�շ�ϸĿID And A.���=D.���� And A.��� Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,�������� B,ҩƷ��� C,����֧����Ŀ M,����֧������ N,(" & strStock & ") S" & _
                    " Where A.ID=B.����ID(+) And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[5]  And A.ID=C.ҩƷID(+) And A.ID=S.ҩƷID(+)" & _
                    " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=1 or (w.������Դ is Null And Nvl(w.��������id,[6]) = [6]))))" & _
                    " And (a.���id not in ('4','5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[6])=[6]))" & _
                    " Group by A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,C.���ﵥλ,C.�����װ,S.���,N.����,A.˵��,A.�Ƿ���,A.���ID,B.��������" & _
                    " Order by A.���,A.����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ���Ŀ", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint���� + 1, "," & str��ĿIDs & ",", mint����, mlng�������ID, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��, mlng���ϲ���, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                If Not rsTmp Is Nothing Then
                    '�Ǳ���ִ�е�ҽ����������������Ŀ
                    If lng�к� <> 0 Then
                        If NVL(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And NVL(rsTmp!��������ID, 0) = 1) Then
                            If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                                MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    'ҽ��������
                    If CheckItemInsure(rsTmp) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                        .SetFocus: Exit Sub
                    End If
                
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    If lng�к� <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õ��շ���Ŀ��", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            ElseIf Col = COLP_ִ�п��� And .EditText <> "" Then 'ִ�п���
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_�շ����) = "4" Then
                    '�������õ�����
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                        " And B.������� IN(1,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=1)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1] And (C.���� Like [3] Or C.���� Like [4] Or C.���� Like [4])" & _
                        " Order by B.�������,C.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                    'ҩƷ��ϵͳָ���Ĵ���ҩ������
                    If Not Check�ϰల��(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(1,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And (A.������Դ is NULL Or A.������Դ=1)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(1,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                            " And (A.������Դ is NULL Or A.������Դ=1)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                        Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!����
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    
                    '���¼�¼��
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                        mrsPrice!ִ�п���ID = rsTmp!ID
                        mrsPrice.Update
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lngҽ��ID As Long, ByVal int�������� As Integer, ByVal lngԭ��ĿID As Long)
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng�к� As Long, dbl���� As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '��¼������
        '�������:����ʱ��ʾ�����������Ŀ,Ҳ���Դ���Ϊδ���Ƽ�ҽ��������������Ŀ
        .TextMatrix(lngRow, COLP_���) = rsInput!���
        .TextMatrix(lngRow, COLP_�շ����) = rsInput!���ID
        .TextMatrix(lngRow, COLP_�շ�ϸĿID) = rsInput!ID
        .TextMatrix(lngRow, COLP_�շ���Ŀ) = rsInput!����
        If Not IsNull(rsInput!����) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & "(" & rsInput!���� & ")"
        End If
        If Not IsNull(rsInput!���) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & " " & rsInput!���
        End If
        .TextMatrix(lngRow, COLP_��λ) = NVL(rsInput!��λ) '�������۵�λ(������ҩ��ҩƷ�Ƽ�)
        .TextMatrix(lngRow, COLP_�Ƽ�����) = 1 'ȱʡ��ԼƼ�1,ҩƷΪ��1�����۵�λ
        
        'ִ�п���
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lng�к� <> 0 Then
            lngִ�п���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))
            '��ҩ��ҩƷ�͸������õ�����ר����ִ�п���
            If rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 Or InStr(",5,6,7,", rsInput!���ID) > 0 Then
                lng���˿���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID))
                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsInput!���ID, rsInput!ID, 4, lng���˿���ID, 0, 1, lngִ�п���ID)
            End If
        End If
        .TextMatrix(lngRow, COLP_ִ�п���) = Sys.RowValue("���ű�", lngִ�п���ID, "����")
        .TextMatrix(lngRow, COLP_ִ�п���ID) = lngִ�п���ID
        
        '���ۼ��㴦��:ҩ����ҩƷ�Ƽ۲����������ﴦ��
        If InStr(",5,6,7,", rsInput!���ID) > 0 Then
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                dbl���� = NVL(rsInput!�ּ�ID, 0)
            ElseIf lng�к� <> 0 Then
                '��ÿ��ȱʡһ�����۵�λ,��ǰ�������μ���
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            End If
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
                        
            'ʱ��ҩƷ������۸�
            .TextMatrix(lngRow, COLP_���) = 0
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        ElseIf rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 And NVL(rsInput!�Ƿ���ID, 0) = 1 Then
            '�������õ�ʱ�����ĺ�ҩƷһ������
            dbl���� = 0
            If lng�к� <> 0 Then
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            End If
            .TextMatrix(lngRow, COLP_���) = 0
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        Else
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_���) = 0
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!�ּ�ID, 0), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
            Else
                .TextMatrix(lngRow, COLP_���) = 1
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!ȱʡ�۸�ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = NVL(rsInput!ԭ��ID, 0)
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = NVL(rsInput!�ּ�ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_��������) = NVL(rsInput!��������)
        .TextMatrix(lngRow, COLP_�̶�) = 0
        
        '��������ָ�
        .Cell(flexcpData, lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ)
        .Cell(flexcpData, lngRow, COLP_�Ƽ�����) = .TextMatrix(lngRow, COLP_�Ƽ�����)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_ִ�п���) = .TextMatrix(lngRow, COLP_ִ�п���)
        
        '��¼������
        If lngҽ��ID <> 0 Then
            If lngԭ��ĿID = 0 Then
                '��ǰҽ���Ƿ��д��������������Ŀ�Ƿ����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And ����=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_����) = IIF(blnHaveSub, "��", "")
            
                mrsPrice.AddNew '����
            Else '����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
            End If
            If lngԭ��ĿID = 0 Then
                mrsPrice!ҽ��ID = lngҽ��ID
                lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
                If Val(vsAdvice.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                    mrsPrice!���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���ID))
                Else
                    mrsPrice!���ID = Null
                End If
                mrsPrice!�������� = int��������
                mrsPrice!���� = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!�շѷ�ʽ = 0
            mrsPrice!�շ���� = rsInput!���ID
            mrsPrice!�շ�ϸĿID = rsInput!ID
            If lngִ�п���ID <> 0 Then
                mrsPrice!ִ�п���ID = lngִ�п���ID
            Else
                mrsPrice!ִ�п���ID = Null
            End If
            mrsPrice!���� = NVL(rsInput!��������ID, 0)
            mrsPrice!��� = NVL(rsInput!�Ƿ���ID, 0)
            mrsPrice!���� = Val(.TextMatrix(lngRow, COLP_����))
            mrsPrice!���� = 1
            mrsPrice!�̶� = 0
            mrsPrice.Update
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln�Ǳ��� As Boolean
    
    If Not CellEditable(Row, Col, bln�Ǳ���) Then
        '�Ǳ���ִ�еı����Ŀ�������۸�
        If bln�Ǳ��� Then
            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_�Ƽ����� Or Col = COLP_���� Or Col = COLP_ִ�п��� Then
            '������ȷ���շ���Ŀ
            If vsPrice.TextMatrix(Row, COLP_�շ���Ŀ) = "" Then Cancel = True
        End If
        If Col = COLP_���� Then
            '������ǰ������ȷ���Ƽ�ҽ��,�Ծ����Ƿ��������(����ִ��)
            If vsPrice.TextMatrix(Row, COLP_�Ƽ�ҽ��) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub InitBillSet()
'���ܣ���ʼ��ҽ�����ʵ������ɼ�¼��
    Set mrsBill = New ADODB.Recordset
    mrsBill.Fields.Append "Key", adVarChar, 200
    mrsBill.Fields.Append "NO", adVarChar, 30
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
    
    Set mrsRXKey = New ADODB.Recordset
    mrsRXKey.Fields.Append "Key", adVarChar, 200
    mrsRXKey.Fields.Append "ҽ��ID", adVarChar, 200
    mrsRXKey.Fields.Append "����", adBigInt
    mrsRXKey.Fields.Append "����", adBigInt
    mrsRXKey.CursorLocation = adUseClient
    mrsRXKey.LockType = adLockOptimistic
    mrsRXKey.CursorType = adOpenStatic
    mrsRXKey.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng������� As Long, lng������� As Long, bln���� As Boolean)
'���ܣ���ȡ��ǰ���õ��ݵ�NO�����
'������lng�������=���ü�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'      lng�������=���ͼ�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'˵����strKey=���ݼ��ʵ������ɹ��򶨵�Ψһ�ؼ���
'1.������ҩ��"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
'2.һ���䷽�е����в�ҩ����һ���������ݺ�
'3.����ҽ�����ҩ�ֺŹ�����ͬ��
'***��ҩ��������ģ�����"����Ϊͬһ���ݵ�ҽ�����"������ͬһ����Ƿ�ִ�п��һ��ֵ��ݺš�
'4.������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
'5.��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
'6.һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        
        'ȡ���ݺ�
        'mrsBill!NO = zlDatabase.GetNextNo(IIF(bln����, 14, 13))
        mlngNOSequence = mlngNOSequence + 1
        mrsBill!NO = "TemporaryNO=" & IIF(bln����, 14, 13) & Format(mlngNOSequence, "00000")
        
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill.Update
    Else
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng������� <> -1 Then lng������� = mrsBill!�������
    If lng������� <> -1 Then lng������� = mrsBill!�������
End Sub

Private Sub ReplaceTrueNO(rsSQL As ADODB.Recordset)
'���ܣ�����ʱ������NO�滻�����ձ������ʵNO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = zlDatabase.GetNextNo(Val(Left(strCur, 2)))
            End If
            
            rsSQL!Sql = Replace(rsSQL!Sql, rsSQL!NO, strNO)
            rsSQL!NewNO = strNO
            'rsSQL!NO = strNO '��������£����⵼��Sort��˳������
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Private Sub DeleteSendRow()
'���ܣ���������ҽ���嵥���ѷ��ͳɹ��ĵ���ɾ��
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_ѡ��
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Getʵ�ս��(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Beginʵ��"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "Endʵ��" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Getʵ�ս�� = CCur(strSQL)
End Function

Private Function Setʵ�ս��(ByVal strSQL As String, ByVal cur��� As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Beginʵ��"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "Endʵ��" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Setʵ�ս�� = strLeft & cur��� & strRight
End Function

Private Function Set��̬�ѱ�(ByVal strSQL As String, ByVal str�ѱ� As String) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin�ѱ�"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End�ѱ�" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Set��̬�ѱ� = strLeft & str�ѱ� & strRight
End Function

Private Function CheckSignSend() As Boolean
'���ܣ����һ��ǩ����ҽ���Ƿ�һ���͵�
    Dim colǩ��ID As New Collection, strǩ��ID As String
    Dim lngǩ��id As Long, strTmp As String
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            '�ռ���ǩ��ҽ���ķ���״̬
            lngǩ��id = Val(.TextMatrix(i, COL_ǩ��ID))
            If lngǩ��id <> 0 Then
                If InStr(strǩ��ID & ",", "," & lngǩ��id & ",") > 0 Then
                    strTmp = Split(colǩ��ID("_" & lngǩ��id), "=")(1)
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If InStr(strTmp, "1") = 0 Then
                            colǩ��ID.Remove "_" & lngǩ��id
                            colǩ��ID.Add lngǩ��id & "=" & strTmp & "1", "_" & lngǩ��id
                        End If
                    Else
                        If InStr(strTmp, "0") = 0 Then
                            colǩ��ID.Remove "_" & lngǩ��id
                            colǩ��ID.Add lngǩ��id & "=" & strTmp & "0", "_" & lngǩ��id
                        End If
                    End If
                Else
                    strǩ��ID = strǩ��ID & "," & lngǩ��id
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        colǩ��ID.Add lngǩ��id & "=1", "_" & lngǩ��id
                    Else
                        colǩ��ID.Add lngǩ��id & "=0", "_" & lngǩ��id
                    End If
                End If
            End If
        Next
        
        '���ǩ�����(һ��ǩ����ҽ������һ����)
        strTmp = ""
        For i = 1 To colǩ��ID.Count
            lngǩ��id = Split(colǩ��ID(i), "=")(0)
            strǩ��ID = Split(colǩ��ID(i), "=")(1)
            If Not (strǩ��ID = "1" Or strǩ��ID = "0") Then
                '���ǩ�������ݲ���"��Ҫ���ͻ򶼲�����"�����
                j = .FindRow(CStr(lngǩ��id), , COL_ǩ��ID)
                Do While j <> -1
                    If Not .RowHidden(j) Then
                        If .Cell(flexcpData, j, COL_ѡ��) = 1 Or .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                            strTmp = strTmp & vbCrLf & "��" & .TextMatrix(j, col_ҽ������)
                        End If
                    End If
                    j = .FindRow(CStr(lngǩ��id), j + 1, COL_ǩ��ID)
                Loop
                Exit For '��ֻ��ʾ��һ��
            End If
        Next
    End With
    
    If strTmp <> "" Then
        MsgBox "����ҽ������������Ҫ���͵�ҽ��һ��ǩ��������ǰ����Ϊ�����ͣ�" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "һ��ǩ����ҽ������һ���ͣ���������ҽ���ķ���״̬��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSignSend = True
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng��ĿID As Long, ByVal int�������� As Integer, ByVal lngCol As Long)
'���ܣ���λ������ʾָ��ҽ����ָ���Ƽ���
'������lngRow=ҽ���к�
'      lng��ĿID=�Ƽ���ĿID
'      lngCol=�Ƽ۱����ʾ��
    Dim k As Long
    
    With vsAdvice
        .Col = col_ҽ������ '�������Զ�ShowPrice,mrsPrice�����仯
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_�������)) > 0 And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                '��������,��������,��鲿λ,���������Ŀ
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '��ҩ;��
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '��ҩ�巨
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_�к�)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_��������)) = int�������� _
                And Val(vsPrice.TextMatrix(k, COLP_�շ�ϸĿID)) = lng��ĿID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function GetMergeDrugStore(ByVal lngRow As Long) As Long
'���ܣ���ȡһ����ҩ�Ļ�׼ҩ�����������ɷ���NO��Keyֵ
'˵����һ����ҩ��ҩƷ���͵�һ�𣬰����Ա�ҩ�Ͳ�ͬҩ�������
    Dim lngҩ��ID As Long, lngBegin As Long, i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_���ID)) <> Val(.TextMatrix(lngRow - 1, COL_���ID)) And Val(.TextMatrix(lngRow, COL_ִ�п���ID)) <> 0 Then
            lngҩ��ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        Else
            lngBegin = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
            For i = lngBegin To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                        lngҩ��ID = Val(.TextMatrix(i, COL_ִ�п���ID)): Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetMergeDrugStore = lngҩ��ID
End Function

Private Sub InitExecRecordset(rsExec As Recordset)
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "ҽ��ID", adBigInt
    rsExec.Fields.Append "���ͺ�", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "Ҫ��ʱ��", adDate, , adFldIsNullable
    rsExec.Fields.Append "����", adDouble, , adFldIsNullable
    rsExec.Fields.Append "��������", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
End Sub

Private Function SendAdvice(ByVal bln���� As Boolean) As Long
'���ܣ�����ҽ������(��������м��ʱ���)
'˵����������˷����ύ
'���أ�����ɹ��򷵻ط��ͺ�
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsNumber As ADODB.Recordset '������������Ķ�̬��¼��
    Dim rsItems As ADODB.Recordset '����ҽ���ܿصķ��ü�¼��,��̬��¼��
    Dim rsMoneyNow As ADODB.Recordset '��ǰ���˱���Ҫ���͵ķ���,��̬��¼��
    Dim rsMoneyDay As ADODB.Recordset '��ǰ���˵����ѷ��͵ķ���,��̬��¼��
    Dim rsExec As ADODB.Recordset  'ҽ��ִ�мƼ�
    
    Dim rsTmp As ADODB.Recordset
    Dim rsMoney As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date
    Dim blnTran As Boolean, blnBool As Boolean
    Dim str��� As String, str���� As String
    
    Dim bln���� As Boolean, int���� As Integer, strTmp As String
    Dim lng���ͺ� As Long, int�Ʒ�״̬ As Integer, strNO As String
    Dim str�շ���Ŀ As String, lng������� As Long, lng���ø��� As Long, lng������� As Long
    Dim int���� As Integer, dbl���� As Double, cur�ϼ� As Currency, cur���ʺϼ� As Currency
    Dim dbl���� As Double, dblӦ�� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim str�ֽ�ʱ�� As String, str�״�ʱ�� As String, strĩ��ʱ�� As String
    Dim int�䷽�� As Integer, strNOKey As String, str�Զ����� As String, strPre���Ƶ���ID As String
    Dim str����ʱ�� As String, str�Ǽ�ʱ�� As String
    Dim dbl�������� As Double, blnFirst As Boolean '�䷽�����ֺŹؼ���
    Dim lngҩƷ���ID As Long, lng�������ID As Long
    Dim lngִ�п���ID As Long, intִ��״̬ As Integer
    Dim bln��Ժ��ҩ As Boolean, bln�������� As Boolean, str�ѱ� As String
    
    Dim rsClone As ADODB.Recordset
    Dim rsSeek As ADODB.Recordset
    Dim strNoneSub As String, strHaveSub As String
    Dim int����� As Integer, lng����ĿID As Long, strʵ�� As String
    Dim bln������Ŀ�� As Boolean, lng���մ���ID As Long, str���ձ��� As String, str�������� As String
    
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    
    '����ǩ��
    Dim lng��ID As Long, strҽ��IDs As String, strSource As String
    Dim intRule As Integer, strSign As String, strTimeStamp As String, strTimeStampCode As String
    Dim lng֤��ID As Long, lngǩ��id As Long
    
    Dim strCuvetteNumber As String  '��������
    Dim blnʵʱ��� As Boolean, rsҽ����� As ADODB.Recordset
    Dim strժҪ As String
    Dim lng���ô��� As Long 'һ��ֻ��һ��ʱ�����η���Ӧ��ȡ�ķ��ô���
    Dim str����ҽ��IDs As String, bln���֧��Tmp As Boolean
    Dim lngҽ��ID As Long
    Dim strҽ�����ids As String
    Dim lng��ҽ���� As Long
    Dim str���ҽ��IDs As String
    Dim lng�ɼ�����ID As Long
    Dim str��ҩIDs As String, str������ҩIDs As String, strҽ������ As String
    Dim bln���������� As Boolean
    Dim str��λ���� As String '�����Ŀ�Ĳ�λ�������̶���ʽ����鲿λ<sTab>��鷽�����磺"ͷ��<sTab>ƽɨ"
    Dim dblOther���� As Double '������Ŀ�շѴ���
    Dim str����ҩ��  As String '������ҩƷ��ҽ�� ,"Ƥ��ҽ��ID,ҩƷ��ҽ��ID"
    Dim rsƤ�� As ADODB.Recordset
    Dim strMinDate As String
    Dim lngԤԼ���� As Long
    
    On Error GoTo errH
    
    '��������������ҩ�󷽽���ж�
    Call Check�������
    Call FuncPassPharmReview
    
    '���һ��ǩ����ҽ���Ƿ�һ����
    If Not CheckSignSend Then Exit Function
    
    'RISԤԼ����ж���ʾ
    Call CheckRISScheduling
    
    Call InitExecRecordset(rsExec)   'ҽ��ִ�мƼ�
    
    'Ʒʱ��ȡҩƷ������
    lngҩƷ���ID = ExistIOClass(IIF(bln����, 9, 8))
    If lngҩƷ���ID = 0 Then
        MsgBox "����ȷ��ҩƷ�������ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    lng�������ID = ExistIOClass(IIF(bln����, 41, 40)) '����ȷ���Ƿ�ʹ���������շ�,�������ж�
    
    Screen.MousePointer = 11
    
    blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
    bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
    
    If Not IsNull(mrsPati!����) Then
        blnʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mrsPati!����)
        If blnʵʱ��� Then
            strSQL = "Select A.���ID,A.����ID,B.ҽ��ID From ������ϼ�¼ A,�������ҽ�� B Where A.����ID=[1] And ��ҳID=[2] And A.ID=B.���ID"
            Set rsҽ����� = zlDatabase.OpenSQLRecord(strSQL, "SendAdvice", mlng����ID, mlng�Һ�ID)
        End If
    Else
        blnʵʱ��� = False
    End If
    
    Call InitBillSet
    Call InitRecordSet(rsSQL, rsTotal, rsNumber, rsMoneyNow, rsItems)
    mlngNOSequence = 0 '���ݺ��������³�ʼ
    mblnԤԼ���� = False
    mlngԤ��Ժҽ��ID = 0
    lng���ͺ� = zlDatabase.GetNextNo(10)
    curDate = zlDatabase.Currentdate
    '����(ҽ��ID,����ʱ��)�ظ�
    '����ҽ��Insertʱ��-2s�ˣ�������ȡ����1sҲ�����ظ�
    If mblnAuto Then
        curDate = DateAdd("s", 1, curDate)
    End If
    bln���� = True '��ʼȫ���ǻ���
    int�䷽�� = 1 '��ʾ���͵ĵڼ����䷽,���ڷֵ��ݺ�
    
    With vsAdvice
        If InitObjRecipeAudit(p����ҽ���´�) Then
            '�������ϵͳ������������
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "2" Then
                        str��ҩIDs = str��ҩIDs & "," & .TextMatrix(i, COL_ID)
                    End If
                End If
            Next
            If Mid(str��ҩIDs, 2) <> "" Then
                Call gobjRecipeAudit.BuildData(Mid(str��ҩIDs, 2), mlng�������ID, 0, mlng����ID, mlng�Һ�ID, str������ҩIDs)
            End If
            For i = .FixedRows To .Rows - 1
                If str������ҩIDs <> "" And (InStr("," & str������ҩIDs & ",", "," & .TextMatrix(i, COL_ID) & ",") > 0 Or InStr("," & str������ҩIDs & ",", "," & .TextMatrix(i, COL_���ID) & ",") > 0) Then
                    Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                    .Cell(flexcpData, i, COL_ѡ��) = 1
                    If Val(.TextMatrix(i, COL_���ID)) <> 0 Then strҽ������ = strҽ������ & vbCrLf & .TextMatrix(i, col_ҽ������)
                End If
            Next
            If str������ҩIDs <> "" Then
                Call MsgBox("��ǰ�����ô������ϵͳ�����·��͵�ҽ����Ҫ����ҽ������ȴ������ɺ��ٷ���ҽ����" & strҽ������, vbInformation, Me.Caption)
            End If
        End If
        
        '����Ƿ��д���
        strժҪ = Replace(txtNote.Text, "'", "''")
        If mint���� <> 0 Then
            If gclsInsure.GetCapability(supportҽ��ȷ����������, mlng����ID, mint����) Then
                strժҪ = "2"
                strTmp = zlCommFun.ShowMsgBox("��������", "��ȷ����ǰҽ�����˱���Ҫ���͵�ҩƷ���������͡�", "!ҽ����(&A),ҽ����(&B),?ȡ��(&C)", Me)
                If strTmp = "" Then Exit Function
                If strTmp = "ҽ����" Then strժҪ = "1"
            End If
        End If
        
        '��������ж�
        If gbln����ҩƷ�ֿ����� Then
            If cboDrugType.ListIndex = 0 Then
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If InStr("," & str���� & ",", "," & .TextMatrix(i, COL_�������) & ",") = 0 Then
                            str���� = str���� & "," & .TextMatrix(i, COL_�������)
                        End If
                    End If
                Next
                If str���� <> "" Then
                    If Not (str���� = ",����ҩ" Or str���� = ",����I��" Or str���� = ",����ҩ" Or str���� = ",����ҩ,����I��" Or str���� = ",����I��,����ҩ") Then
                        If Not (InStr(str���� & ",", ",����ҩ,") = 0 And InStr(str���� & ",", ",����ҩ,") = 0 And InStr(str���� & ",", ",����I��,") = 0) Then
                            Screen.MousePointer = 0
                            MsgBox "���η��͵�ҽ���п��ܰ������龫��ҩƷ����ֱ��ͣ����޸Ĺ����������¶�ȡҽ�����ٷ��͡�", vbInformation, gstrSysName
                            mblnUnload = False
                            Exit Function
                        Else
                            str���� = ""
                        End If
                    End If
                End If
            ElseIf cboDrugType.ListIndex = 3 Then
                str���� = ""
            Else
                str���� = ",����ҩ"
            End If
        End If
        
        '��Сʱ�����
        strMinDate = "3000-01-01 00:00"
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�״�ʱ��) < strMinDate Then
                    strMinDate = .TextMatrix(i, COL_�״�ʱ��)
                End If
            End If
        Next
        If strMinDate = "3000-01-01 00:00" Then strMinDate = ""
        
        '������ҩ
        If mbln������ҩ Then
            blnBool = Set������ҩ()
            If Not blnBool Then
                GoTo FuncEnd
            End If
        End If
        
        If Not zlPluginAdviceBeforeSend Then
            GoTo FuncEnd
        End If
        
        'ҽ�����ʹ���
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '�������ݺŷ���ؼ���
                '-----------------------------------------------------------------------------------------
                bln���������� = False
                If mintSendNo = 1 And Not gblnִ��ǰ�Ƚ��� Then
                    strNOKey = "ֻ����һ�����ݺ�"
                ElseIf mintSendNo = 2 Then
                    strNOKey = Val(.TextMatrix(i, COL_ִ�п���ID))
                    bln���������� = (InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 And gintRXCount > 0)
                Else
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        '������ҩ��"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                        'һ����ҩ�ģ����͵�һ�𣺰����Ա�ҩ�Ͳ�ͬҩ�������
                        strNOKey = "������ҩ_" & mlng����ID & "_" & mstr�Һŵ� & "_" & _
                            Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                             GetMergeDrugStore(i)
                                                    
                        If mblnһ����ҩ����Ϊһ�� Then
                            If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                                '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�(һ����ҩ�ģ�ֻȡ��һ��ҩƷ�����Ƶ���ID)
                                strPre���Ƶ���ID = GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 1)
                            End If
                            strNOKey = strNOKey & "_" & strPre���Ƶ���ID
                        Else
                            strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 1)
                        End If
                        bln���������� = (gintRXCount > 0)
                    ElseIf InStr(",4,M,", .TextMatrix(i, COL_�������)) > 0 Then
                        '���ϰ�"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                        strNOKey = "����ҽ��_" & mlng����ID & "_" & mstr�Һŵ� & "_" & _
                            Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                            Val(.TextMatrix(i, COL_ִ�п���ID))
                        '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�
                        strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 1)
                        
                    ElseIf .TextMatrix(i, COL_�������) = "7" Then
                        'һ���䷽�е����в�ҩ����һ���������ݺ�
                        strNOKey = "��ҩ�䷽_" & mlng����ID & "_" & mstr�Һŵ� & "_" & int�䷽��
                    
                    '��ҩ����ͬһ����Ƿ���ִͬ�п�����ϵ���
                    ElseIf InStr(mstr����������, .TextMatrix(i, COL_�������)) > 0 Then
                        strNOKey = "��ҩҽ��_" & .TextMatrix(i, COL_�������) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                        
                    ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "C" Then
                        'һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
                        'ͬһ��������ͣ�ͬһ������ִ�п��ң�ͬһ�ɼ��ܣ�ͬһ���ɼ���ʽ��ͬһ���ɼ�ִ�п��ҵļ��������ͬ�ĵ��ݺ�
                        If mbln���鵥���������� Then
                            strNOKey = "һ���ɼ�_" & Val(.TextMatrix(i, COL_���ID))
                        Else
                            lng��ҽ���� = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                            strNOKey = "һ���ɼ�_" & mlng����ID & "_" & mstr�Һŵ� & "_" & .TextMatrix(i, COL_�걾��λ) & "_" & _
                                .TextMatrix(i, COL_ִ�п���ID) & "_" & .TextMatrix(i, COL_��������) & "_" & .TextMatrix(i, COL_�Թܱ���) & "_" & _
                                .TextMatrix(lng��ҽ����, COL_������ĿID) & "_" & .TextMatrix(lng��ҽ����, COL_ִ�п���ID)
                        End If
                    ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_�������)) > 0 Then
                        '��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
                        strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_���ID))
                        
                    Else
                        '������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷����ɼ���ʽ������ʽ����Ѫҽ��/��Ѫ;��)
                        strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_ID))
                    End If
                End If
                
                '��ͬ��ϵ�ҽ���ֱ��������,Ҫ���¼ӹ�strNOKey
                If mblnNOCtrl Then
                    lngҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)))
                    If lngҽ��ID <> lng��ID Then strҽ�����ids = GetAdviceDiag(lngҽ��ID)
                    lng��ID = lngҽ��ID

                    If strҽ�����ids <> "" Then strNOKey = strNOKey & "_" & strҽ�����ids
                End If
                
                '��ʼʱ�䲻��ͬһ��ķֱ��������
                If mblnStartTimeDef Then
                    strNOKey = strNOKey & "_" & Format(.TextMatrix(i, COL_��ʼʱ��), "YYYY-MM-DD")
                End If
 
                '�����˲�ͬ�ģ�Ĭ��ȫ���ֱ��������
                strNOKey = strNOKey & "_" & .TextMatrix(i, COL_����ҽ��)
                
                '���ò���������ҩƷ�ֿ����� ʱ������ҩƷҽ����ҩƷ�е������ɵ��ݺţ�һ��ҽ������һ����
                If str���� <> "" Then
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        strNOKey = "������ҩ_" & .TextMatrix(i, COL_���ID)
                    End If
                End If
                
                '������������Ӧ�÷ŵ����
                If bln���������� Then
                    strTmp = ""
                    If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                        strTmp = GetMergeIDs(vsAdvice, i, COL_���ID, COL_ID) 'һ����ҩ��ʼ�л����ҩƷ�в�ȡֵ
                    End If
                    strTmp = GetRXKey(mrsRXKey, strNOKey, strTmp)
                    If strTmp <> "1" Then
                        strNOKey = strNOKey & "_" & strTmp
                    End If
                End If
                
                '�Ƿ���Ժ��ҩ
                bln��Ժ��ҩ = False
                If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                    If .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ" Then bln��Ժ��ҩ = True
                ElseIf .TextMatrix(i, COL_�������) = "7" Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                    If j <> -1 Then
                        If .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ" Then bln��Ժ��ҩ = True
                    End If
                End If
                
                '����ҽ�����ʷ���:�����¼۸����
                '-----------------------------------------------------------------------------------------
                strSQL = "": str�շ���Ŀ = ""
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
                    If Val(.TextMatrix(i, COL_ִ������ID)) <> 5 Then
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,RTrim(A.����||' '||A.���) as ����," & _
                            " A.���㵥λ,A.�Ƿ���,A.���ηѱ�,A.����ȷ��,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,100 as �����շ���," & _
                            " Y.���ﵥλ,Y.�����װ,Y.����ϵ��,Y.ҩ������ as ����,0 as ��������,B.������ĿID," & _
                            " C.�վݷ�Ŀ,1 as ����,B.�ּ� as ����,[2] as ִ�п���ID,0 as ����,0 as ��������,0 as �շѷ�ʽ" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,ҩƷ��� Y" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.���=D.����" & _
                            GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "3", "4", "5") & _
                            " And A.ID=Y.ҩƷID(+) And A.ID=[1]" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " Order by A.����"
                    End If
                Else
                    '��ɾ��ԭ��ҩҽ���ļƼ�(Ӧ��û��)
                    rsSQL.AddNew
                    rsSQL!���� = 1: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    rsSQL!Sql = "ZL_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                    rsSQL.Update
                    
                    '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�е�ҽ������ȡ
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�շ�ϸĿID, 0) <> 0 And NVL(mrsPrice!����, 0) <> 0 Then '��������Ϊ0���Զ����˵�
                                    '��ͨ��Ŀ�ı�۵���Ҫ�����룬�����Ǹ������õ�ʱ������ҽ��
                                    If NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 _
                                        And Not (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                        Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, mrsPrice!��������, COLP_����)
                                        Screen.MousePointer = 0
                                        MsgBox "����Ϊ��۵��շ���Ŀȷ��һ���շѼ۸�", vbInformation, gstrSysName
                                        mblnUnload = False: vsPrice.SetFocus: GoTo FuncEnd
                                    End If
                                    
                                    '�Ƽ�ִ�п���:ֻ�����ҩƷ������ҽ���ģ�ҩƷ�����ļƼ۵�ִ�п���
                                    If InStr(",4,5,6,7,", .TextMatrix(i, COL_�������)) = 0 _
                                        And (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                        lngִ�п���ID = NVL(mrsPrice!ִ�п���ID, 0)
                                        
                                        '���ı�������ִ�п���
                                        If lngִ�п���ID = 0 And mrsPrice!�շ���� = "4" Then
                                            Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, mrsPrice!��������, COLP_ִ�п���)
                                            Screen.MousePointer = 0
                                            MsgBox "����""" & vsPrice.TextMatrix(vsPrice.Row, COLP_�շ���Ŀ) & """û��ȷ��ִ�п��ң����ֹ�������ȷ��ִ�п��ҡ�" & vbCrLf & _
                                                "�������ȷ����ȷ��ִ�п��ң��뵽""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                            mblnUnload = False: vsPrice.SetFocus: GoTo FuncEnd
                                        End If
                                    Else
                                        lngִ�п���ID = 0
                                    End If
                                    
                                    'ҩƷ������ҽ���ļƼ۹̶���Ӧ�����棻�Ǹ������õ�ʱ�����ĵı����Ҫ���룬���Ҫ���浽�Ƽ۱���
                                    If InStr(",4,5,6,7,", .TextMatrix(i, COL_�������)) = 0 _
                                        Or .TextMatrix(i, COL_�������) = "4" And NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 Then
                                        rsSQL.AddNew
                                        rsSQL!���� = 1: rsSQL!��ĿID = mrsPrice!�շ�ϸĿID: rsSQL!��� = i
                                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                        rsSQL!Sql = "ZL_����ҽ���Ƽ�_INSERT(" & _
                                            mrsPrice!ҽ��ID & "," & mrsPrice!�շ�ϸĿID & "," & _
                                            NVL(mrsPrice!����, 0) & "," & NVL(mrsPrice!����, 0) & "," & _
                                            NVL(mrsPrice!����, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                            NVL(mrsPrice!��������, 0) & "," & NVL(mrsPrice!�շѷ�ʽ, 0) & ")"
                                        rsSQL.Update
                                    End If
                                    
                                    '��ʱ����ҽ���Ƽ۱�
                                    If Val(.TextMatrix(i, COL_����)) <> 0 Then
                                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                            "Select " & mrsPrice!�շ�ϸĿID & " as �շ�ϸĿID," & _
                                            NVL(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID," & _
                                            NVL(mrsPrice!����, 0) & " as ����," & Format(NVL(mrsPrice!����, 0), gstrDecPrice) & " as ����," & _
                                            NVL(mrsPrice!����, 0) & " as ����," & NVL(mrsPrice!��������, 0) & " as ��������," & _
                                            NVL(mrsPrice!�շѷ�ʽ, 0) & " as �շѷ�ʽ From Dual"
                                    End If
                                End If
                                
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    
                    If strSQL <> "" Then
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,A.����,A.���㵥λ,A.�Ƿ���," & _
                            " A.���ηѱ�,A.����ȷ��,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,Y.���ﵥλ,Y.�����װ,Y.����ϵ��," & _
                            " Decode(A.���,'4',E.���÷���,Y.ҩ������) as ����,E.��������,B.������ĿID," & _
                            " C.�վݷ�Ŀ,X.����,Decode(A.�Ƿ���,1,X.����,B.�ּ�) as ����,X.ִ�п���ID,X.����,X.��������,X.�շѷ�ʽ" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,�������� E,(" & strSQL & ") X,ҩƷ��� Y" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.ID=E.����ID(+)" & _
                            GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "3", "4", "5") & _
                            " And A.���=D.���� And X.�շ�ϸĿID=A.ID And A.ID=Y.ҩƷID(+)" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " Order by X.��������,X.����,X.�շѷ�ʽ Desc,A.ID"
                            'һ��Ҫ����������ǰ��,�Ա��ڼ�����ڷ��ü�¼�б������ӹ�ϵ
                    End If
                End If
                                
                '�����ۿ۱�����ʼ
                int����� = 0: lng����ĿID = 0
                strHaveSub = "": strNoneSub = ""
                Call InitSeekSet(rsSeek)
                
                '��ǰ������������(����"ҽ����������������"û������ʱҲ����һ����������룬�����ж��Ƿ��ղ�Ѫ�ܷ���)
                strCuvetteNumber = ""
                If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                    If j > 0 Then lng�ɼ�����ID = Val(.TextMatrix(j, COL_ִ�п���ID))
                    strCuvetteNumber = GetCuvetteNumber(rsNumber, .TextMatrix(i, COL_�Թܱ���), _
                        Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)), .TextMatrix(i, COL_�������), Val(.TextMatrix(i, COL_��������)), _
                        Val(.TextMatrix(i, COL_ִ�п���ID)), Val(.TextMatrix(i, COL_Ӥ��)), Val(.TextMatrix(i, COL_������ĿID)), _
                        Val(.TextMatrix(i, COL_������־)), .TextMatrix(i, COL_�걾��λ), lng�ɼ�����ID)
                End If
                If gobjSquareCard Is Nothing Then
                    bln���֧��Tmp = False
                Else
                    bln���֧��Tmp = gobjSquareCard.zlIsAllowCliniqueRoomPay(p����ҽ��վ, mlng����ID, Val(.TextMatrix(i, COL_ID)), mlngCardType)
                End If
                '�ж��Ƿ��������֧�������ҵ�ǰ���������֧����Լ����,ֻ�з����շѵ�ʱ�����֧��
                '74233�������˲��������������շѻ������˺�����ҽ��������
                If mbln���֧�� And Not bln���� And bln���֧��Tmp Then
                    mstr֧����ʽ = "1" 'Ȩ�ޣ��������ӿڷ�����  ��������������Ϊ���֧��
                Else
                    bln���֧��Tmp = False
                End If
                If bln���֧��Tmp Or gbln�����������շѻ������� Then
                    str����ҽ��IDs = str����ҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
                '����ִ�е��Զ�ִ�У�����ҽ�����ô���
                intִ��״̬ = 0
                If mblnAutoExe Then
                    If (mstrǰ��IDs <> "" And mlngҽ������ID = Val(.TextMatrix(i, COL_ִ�п���ID)) Or _
                        mstrǰ��IDs = "" And Val(.TextMatrix(i, COL_���˿���ID)) = Val(.TextMatrix(i, COL_ִ�п���ID))) _
                        And Not (.TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) <> 0) Then
                        str���ҽ��IDs = str���ҽ��IDs & "," & .TextMatrix(i, COL_ID)
                        'ִ��ǰ�Ƚ���ʱ�������ڡ�ִ�к��Զ���˼��ʻ��۵���
                        If Not (bln���֧��Tmp Or gbln�����������շѻ�������) Then
                            If gblnִ��ǰ�Ƚ��� And Not gobjSquareCard Is Nothing Then
                                str����ҽ��IDs = str����ҽ��IDs & "," & .TextMatrix(i, COL_ID)
                            Else
                                intִ��״̬ = 1
                                'Ѫ��������⴦��
                                If gblnѪ��ϵͳ Then
                                    strTmp = .TextMatrix(i, COL_�������) & .TextMatrix(i, COL_��������)
                                    If strTmp = "E8" Or strTmp = "E9" Then
                                        strTmp = "Select 1 From ������ĿĿ¼ a where a.id=[1] and nvl(a.ִ�з���,0) in (0,1)"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, Val(.TextMatrix(i, COL_������ĿID)))
                                        If Not rsTmp.EOF Then
                                            intִ��״̬ = 0
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "D" Then
                    str��λ���� = .TextMatrix(i, COL_�걾��λ) & "<sTab>" & .TextMatrix(i, COL_��鷽��)
                Else
                    str��λ���� = ""
                End If
                
                int�Ʒ�״̬ = IIF(Val(.TextMatrix(i, COL_�Ƽ�����)) = 1, -1, 0) '����Ʒѻ�δ�Ʒ�
                If strSQL <> "" Then
                    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_ִ�п���ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    If Not rsMoney.EOF Then
                        int�Ʒ�״̬ = 1 '�ѼƷ�
                        Set rsClone = rsMoney.Clone
                    End If
                    
                    '����������Ŀ���ķ�����ϸ
                    bln�������� = .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0
                    Do While Not rsMoney.EOF
MoneyItemBegin:
                        'ִ�п���ID
                        lngִ�п���ID = NVL(rsMoney!ִ�п���ID, 0)
                        '��ԭֵ������ȡ��Ч�ķ�ҩ��ҩƷ���������ĵ�ִ�п���
                        If InStr(",4,5,6,7", .TextMatrix(i, COL_�������)) = 0 _
                            And (rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Or InStr(",5,6,7", rsMoney!���) > 0) Then
                            lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsMoney!���, rsMoney!ID, 4, Val(.TextMatrix(i, COL_���˿���ID)), 0, 1, lngִ�п���ID)
                        End If
                    
                        '�ֽ�ʱ��
                        If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                            str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                        Else
                            str�ֽ�ʱ�� = .Cell(flexcpData, i, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                        End If
                    
                        '----------------------------------------
                        '�����շѷ�ʽ��ȷ����ǰ�շ���Ŀ�Ƿ�Ӧ�շ�
                        If rsMoney!�������� & "_" & rsMoney!ID <> str�շ���Ŀ Then
                            If Not AdviceMoneyMake(mlng����ID, 0, rsMoneyNow, rsMoneyDay, _
                                IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))), _
                                Val(.TextMatrix(i, COL_������ĿID)), rsMoney!ID, lngִ�п���ID, .TextMatrix(i, COL_�Թܱ���), _
                                rsMoney!���, NVL(rsMoney!�շѷ�ʽ, 0), str�ֽ�ʱ��, 1, lng���ô���, Val(.TextMatrix(i, COL_����)), _
                                 Val(.TextMatrix(i, COL_ID)), lng���ͺ�, Val(rsMoney!���� & ""), rsExec, Val(.TextMatrix(i, COL_���㷽ʽ)), _
                                .TextMatrix(i, COL_Ƶ��), Val(.TextMatrix(i, COL_����)), , , .TextMatrix(i, COL_�������), strCuvetteNumber, str��λ����, dblOther����, strMinDate) Then
                                '������ǰ�շ���Ŀ(���������Ŀ)
                                str�շ���Ŀ = rsMoney!�������� & "_" & rsMoney!ID
                                Do While rsMoney!�������� & "_" & rsMoney!ID = str�շ���Ŀ
                                    rsMoney.MoveNext
                                    If rsMoney.EOF Then Exit Do
                                Loop
                                If rsMoney.EOF Then Exit Do
                                GoTo MoneyItemBegin
                            End If
                        End If
                        '----------------------------------------

                        If InStr(",5,6,7", rsMoney!���) > 0 Then
                            If InStr(",5,6,7", .TextMatrix(i, COL_�������)) > 0 Then
                                If .TextMatrix(i, COL_�������) = "7" Then
                                    int���� = Val(.TextMatrix(i, COL_����))
                                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                    If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                        dbl���� = Val(.TextMatrix(i, COL_����)) / NVL(rsMoney!����ϵ��, 1)
                                    Else
                                        dbl���� = IntEx(Val(.TextMatrix(i, COL_����)) / NVL(rsMoney!����ϵ��, 1) / NVL(rsMoney!�����װ, 1)) * NVL(rsMoney!�����װ, 1)
                                    End If
                                Else
                                    int���� = 1
                                    dbl���� = Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!�����װ, 1)
                                    If rsƤ�� Is Nothing Then
                                        Set rsƤ�� = GetԭҺƤ��(0, 0, mstr�Һŵ�)
                                    End If
                                    rsƤ��.Filter = "ҩƷID=" & Val(rsMoney!ID & "")
                                    If Not rsƤ��.EOF Then
                                        If Val(rsƤ��!��� & "") = 0 Then
                                            '���м���������
                                            dbl���� = (Val(.TextMatrix(i, COL_����)) - 1) * NVL(rsMoney!�����װ, 1)
                                            rsƤ��!��� = Val(.TextMatrix(i, COL_ID))
                                            
                                            str����ҩ�� = "'" & rsƤ��!Ƥ��ҽ��ID & "," & rsƤ��!��� & "'"
                                            rsƤ��.Update
                                            If dbl���� <= 0 Then
                                                rsMoney.MoveNext
                                                If rsMoney.EOF Then Exit Do
                                                GoTo MoneyItemBegin
                                            End If
                                        End If
                                    End If
                                    
                                End If
                            Else
                                int���� = 1
                                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,��˲��������㴦��
                                '�����շѶ����е�ҩƷ����Ϊ����ֻ��ȡһ�Σ�����Ϊ���ô���*��������
                                If InStr(",2,3,4,5,6,7,9,", Val("" & rsMoney!�շѷ�ʽ)) > 0 Then
                                    If dblOther���� > 0 Then
                                        dbl���� = Format(dblOther����, "0.00000")
                                    Else
                                        dbl���� = Format(lng���ô��� * NVL(rsMoney!����, 0), "0.00000")
                                    End If
                                Else
                                    dbl���� = Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!����, 0)
                                End If
                            End If
                            dbl���� = Format(dbl����, "0.00000")
                            
                            If NVL(rsMoney!�Ƿ���, 0) = 1 Then
                                dbl���� = Format(CalcDrugPrice(rsMoney!ID, lngִ�п���ID, int���� * dbl����, , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            Else
                                dbl���� = Format(NVL(rsMoney!����, 0), gstrDecPrice)
                            End If
                        ElseIf rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Then
                            '�����������������
                            If lng�������ID = 0 Then
                                Screen.MousePointer = 0
                                MsgBox "����ȷ���������ϵ��ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                            
                            int���� = 1
                            If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!�շѷ�ʽ)) > 0 Then
                                If dblOther���� > 0 Then
                                    dbl���� = Format(dblOther����, "0.00000")
                                Else
                                    dbl���� = Format(lng���ô��� * NVL(rsMoney!����, 0), "0.00000")
                                End If
                            Else
                                dbl���� = Format(Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!����, 0), "0.00000")
                            End If
                            
                            'ȷ��ʱ�����ļ۸�
                            If NVL(rsMoney!�Ƿ���, 0) = 1 Then
                                dbl���� = Format(CalcDrugPrice(rsMoney!ID, lngִ�п���ID, dbl����, , True, 1, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            Else
                                dbl���� = Format(NVL(rsMoney!����, 0), gstrDecPrice)
                            End If
                        Else
                            '�������ڵ������������Ρ�һ��ֻ��һ��ʱ���ж�����Ҫִ�У����ն��ٴΣ����ܵ������������磺ÿ�����Σ�,��Ҫ���շѶ��յĴ���
                            int���� = 1
                            If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!�շѷ�ʽ)) > 0 Then
                                If dblOther���� > 0 Then
                                    dbl���� = Format(dblOther����, "0.00000")
                                Else
                                    dbl���� = Format(lng���ô��� * NVL(rsMoney!����, 0), "0.00000")
                                End If
                            Else
                                dbl���� = Format(Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!����, 0), "0.00000")
                            End If
                            dbl���� = Format(NVL(rsMoney!����, 0), gstrDecPrice)
                        End If
                        
                        '��ҩ��ҩƷ���������ĵĿ����
                        If InStr(",4,5,6,7", .TextMatrix(i, COL_�������)) = 0 _
                            And (rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Or InStr(",5,6,7", rsMoney!���) > 0) Then
                            If TheStockCheck(lngִ�п���ID, rsMoney!���) <> 0 Or NVL(rsMoney!�Ƿ���, 0) = 1 Or NVL(rsMoney!����, 0) = 1 Then
                                If rsMoney!��� = "4" Then
                                    blnBool = CheckPriceStock(i, rsMoney, lngִ�п���ID, int���� * dbl����, rsTotal, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                                Else
                                    blnBool = CheckPriceStock(i, rsMoney, lngִ�п���ID, int���� * dbl����, rsTotal, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                                End If
                                If blnBool Then
                                    Call RowSelectSame(i, COL_ѡ��, rsSQL, rsTotal, strҽ��IDs)
                                    '�����ǩ��ҽ��������Ƿ�һͬǩ����ҽ������һ����
                                    If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 Then
                                        If Not CheckSignSend Then
                                            GoTo FuncEnd
                                        Else
                                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                            GoTo NextAdvice
                                        End If
                                    Else
                                        Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                        GoTo NextAdvice
                                    End If
                                End If
                            End If
                        End If
                            
                        '���ͽ��
                        dblӦ�� = int���� * dbl���� * dbl����
                        If bln�������� Then
                            dblӦ�� = dblӦ�� * NVL(rsMoney!�����շ���, 100) / 100
                        End If
                        
                        '����Ӱ�Ӽ�
                        If gbln�Ӱ�Ӽ� And NVL(rsMoney!�Ӱ�Ӽ�, 0) = 1 Then
                            dblӦ�� = dblӦ�� * (1 + NVL(rsMoney!�Ӱ�Ӽ���, 0) / 100)
                        End If
                        
                        curӦ�� = Format(dblӦ��, gstrDec)
                        
                        'NO,���
                        Call GetCurBillSet(strNOKey, strNO, lng�������, -1, bln����)
                        rsSQL.AddNew: blnBool = False
                        If rsMoney!�������� & "_" & rsMoney!ID <> str�շ���Ŀ Then
                            lng���ø��� = lng�������
                            If rsMoney!���� = 0 Then
                                '��¼������Ϣ������϶��ڴ���ǰ
                                '��ʹ�������ۿۣ�ҲҪ��¼�������ϵ
                                If InStr(strHaveSub & ",", "," & rsMoney!�������� & ",") = 0 _
                                    And InStr(strNoneSub & ",", "," & rsMoney!�������� & ",") = 0 Then
                                    rsClone.Filter = "��������=" & rsMoney!�������� & " And ����=1"
                                    If Not rsClone.EOF Then
                                        int����� = lng�������
                                        lng����ĿID = rsMoney!ID
                                        
                                        rsSeek.AddNew
                                        rsSeek!�������� = rsMoney!��������
                                        rsSeek!�����ǩ = rsSQL.Bookmark 'Variant(Double)
                                        rsSeek!������ID = rsMoney!������ĿID
                                        rsSeek.Update
                                        strHaveSub = strHaveSub & "," & rsMoney!��������
                                        
                                        blnBool = True
                                    Else
                                        strNoneSub = strNoneSub & "," & rsMoney!��������
                                    End If
                                End If
                            End If
                        End If
                        
                        '��������ۿۺϼ�
                        str�ѱ� = NVL(mrsPati!�ѱ�)
                        If gbln��������ۿ� And (rsMoney!���� = 1 Or InStr(strHaveSub & ",", "," & rsMoney!�������� & ",") > 0) Then
                            If .TextMatrix(i, COL_��Ѽ���) = 1 Then
                                curʵ�� = 0
                            Else
                                curʵ�� = curӦ��
                            End If
                            
                            '�ۼ�ҽ���ϼ��������ۿ�
                            rsSeek.Filter = "��������=" & rsMoney!��������
                            rsSeek!�ϼ� = NVL(rsSeek!�ϼ�, 0) + curʵ��
                            rsSeek.Update
                        ElseIf NVL(rsMoney!���ηѱ�, 0) = 0 Then
                            str�ѱ� = NVL(mrsPati!�ѱ�) & IIF(gstr��̬�ѱ� <> "" And Not bln����, "," & gstr��̬�ѱ�, "")
                            
                            If .TextMatrix(i, COL_��Ѽ���) = 1 Then
                                curʵ�� = 0
                            Else
                                curʵ�� = Format(ActualMoney(str�ѱ�, rsMoney!������ĿID, curӦ��, rsMoney!ID, lngִ�п���ID, _
                                    int���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsMoney!�Ӱ�Ӽ�, 0) = 1, NVL(rsMoney!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            End If
                            If InStr(str�ѱ�, ",") > 0 Then str�ѱ� = NVL(mrsPati!�ѱ�)
                        Else
                            If .TextMatrix(i, COL_��Ѽ���) = 1 Then
                                curʵ�� = 0
                            Else
                                curʵ�� = curӦ��
                            End If
                        End If
                        '�����ۿ�ʱ���������ʵ�ս�������⴦��
                        If gbln��������ۿ� And blnBool Then
                            str�ѱ� = Chr(0) & Chr(1) & "Begin�ѱ�" & str�ѱ� & "End�ѱ�" & Chr(0) & Chr(1)
                            strʵ�� = Chr(0) & Chr(1) & "Beginʵ��" & curʵ�� & "Endʵ��" & Chr(0) & Chr(1)
                        Else
                            strʵ�� = curʵ��
                        End If
                        
                        'ҽ������ֶ�
                        bln������Ŀ�� = False: lng���մ���ID = 0: str���ձ��� = "": str�������� = ""
                        If Not IsNull(mrsPati!����) Then
                            strTmp = gclsInsure.GetItemInsure(mlng����ID, rsMoney!ID, curʵ��, True, mrsPati!����, .Cell(flexcpData, i, COL_ҽ������) & "||" & int���� * dbl����)
                            If strTmp <> "" Then
                                bln������Ŀ�� = Val(Split(strTmp, ";")(0)) <> 0
                                lng���մ���ID = Val(Split(strTmp, ";")(1))
                                str���ձ��� = CStr(Split(strTmp, ";")(3))
                                If UBound(Split(strTmp, ";")) >= 5 Then
                                    If Split(strTmp, ";")(5) <> "" Then
                                        str�������� = Split(strTmp, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                        
                        '�ռ����ʱ������
                        If InStr(str���, rsMoney!���) = 0 Then
                            str��� = str��� & rsMoney!���
                        End If
                        
                        '����ʱ��
                        If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                            str����ʱ�� = "To_Date('" & Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str����ʱ�� = "To_Date('" & .Cell(flexcpData, i, COL_�ֽ�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        
                        '��Ϊ���ڲ��Ƽ۵�ҽ������������,���Դ���ļƼ����Զ�Ϊ(0-�����Ƽ�)
                        rsSQL!���� = 2: rsSQL!��ĿID = rsMoney!ID: rsSQL!��� = i
                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                        rsSQL!NO = strNO
                        rsSQL!������� = IIF(InStr(",5,6,7,", "," & .TextMatrix(i, COL_�������) & ",") > 0, "ҩƷ", "0")
                        rsSQL!��ǰ��ҽ��ID = Val(.TextMatrix(i, COL_ID))
                        rsSQL!���� = i & "_" & rsMoney!ID & "_" & lngִ�п���ID
                        curӦ�� = Format(curӦ��, gstrDec)
                        strʵ�� = Format(strʵ��, gstrDec)
                        cur�ϼ� = cur�ϼ� + curʵ��
                        If Not bln���� Then
                            '��δȡ��ҩ����
                            rsSQL!Sql = "ZL_���ﻮ�ۼ�¼_INSERT(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & ",NULL," & _
                                IIF(IsNull(mrsPati!�����), "NULL", "'" & mrsPati!����� & "'") & ",NULL,'" & mrsPati!���� & "'," & _
                                "'" & NVL(mrsPati!�Ա�) & "','" & NVL(mrsPati!����) & "'," & _
                                "'" & str�ѱ� & "',NULL," & _
                                ZVal(.TextMatrix(i, COL_��������ID)) & "," & ZVal(.TextMatrix(i, COL_��������ID)) & "," & _
                                "'" & .TextMatrix(i, COL_����ҽ��) & "'," & IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!��� & "','" & NVL(rsMoney!���㵥λ) & "',NULL," & _
                                int���� & "," & dbl���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & NVL(rsMoney!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                str����ʱ�� & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "'ҽ������','" & UserInfo.���� & "'," & _
                                "'" & IIF(strժҪ = "", .TextMatrix(i, col_ҽ������), strժҪ) & "'," & _
                                Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                ZVal(.TextMatrix(i, COL_����)) & ",'" & .TextMatrix(i, COL_�÷�) & "',1," & _
                                IIF(bln��Ժ��ҩ, 3, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",1," & _
                                "'" & str���ձ��� & "','" & str�������� & "'," & IIF(bln������Ŀ��, 1, 0) & "," & ZVal(lng���մ���ID) & ",NULL,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & ")"
                        Else
                            '�Ƿ񻮼۷���
                            If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                int���� = IIF(InStr(gstr���﷢�ͻ��۵�, "5") > 0, 1, 0)
                            Else
                                int���� = IIF(InStr(gstr���﷢�ͻ��۵�, .TextMatrix(i, COL_�������)) > 0, 1, 0)
                            End If
                            If int���� = 0 Then int���� = IIF(NVL(rsMoney!����ȷ��, 0) = 1, 1, 0)
                            
                            If int���� = 0 Or intִ��״̬ = 1 Then
                                bln���� = False
                                If gdblԤ��������鿨 <> 0 Then cur���ʺϼ� = cur���ʺϼ� + curʵ��
                            End If
                            
                            '�Ǽ�ʱ��
                            If int���� = 1 Then '��ǻ��۵�ʱ�������ֿ�
                                str�Ǽ�ʱ�� = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str�Ǽ�ʱ�� = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                            rsSQL!Sql = "ZL_������ʼ�¼_INSERT(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & "," & _
                                IIF(IsNull(mrsPati!�����), "NULL", "'" & mrsPati!����� & "'") & ",'" & mrsPati!���� & "'," & _
                                "'" & NVL(mrsPati!�Ա�) & "','" & NVL(mrsPati!����) & "'," & _
                                "'" & str�ѱ� & "',NULL," & Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                ZVal(.TextMatrix(i, COL_��������ID)) & "," & _
                                ZVal(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & rsMoney!ID & "," & _
                                "'" & rsMoney!��� & "','" & NVL(rsMoney!���㵥λ) & "'," & _
                                int���� & "," & dbl���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & NVL(rsMoney!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                str����ʱ�� & "," & str�Ǽ�ʱ�� & ",'ҽ������'," & int���� & ",'" & UserInfo.��� & "'," & _
                                "'" & UserInfo.���� & "',NULL," & _
                                "'" & IIF(strժҪ = "", .TextMatrix(i, col_ҽ������), strժҪ) & "'," & _
                                Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_����)) & "," & _
                                "'" & .TextMatrix(i, COL_�÷�) & "',1," & IIF(bln��Ժ��ҩ, 3, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",1,NULL,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & ")"
                        End If
                        rsSQL.Update
                        
                        '��¼�Զ����ϵ�SQL
                        If gbln�����Զ����� And bln���� And int���� = 0 And lngִ�п���ID <> 0 And rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Then
                            If InStr(str�Զ����� & ";", ";" & strNO & "," & lngִ�п���ID & ";") = 0 Then
                                rsSQL.AddNew
                                rsSQL!���� = 5
                                rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                rsSQL!��ĿID = 0
                                rsSQL!��� = i
                                rsSQL!NO = strNO
                                rsSQL!Sql = "zl_�����շ���¼_��������(" & lngִ�п���ID & ",25,'" & strNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                                rsSQL.Update
                                str�Զ����� = str�Զ����� & ";" & strNO & "," & lngִ�п���ID
                            End If
                        End If
                        
                        'ҽ���ܿ�ʵʱ��⣺���ɷ�����Ŀ��¼��,���շ�ϸĿ����
                        If Not IsNull(mrsPati!����) And blnʵʱ��� Then
                            rsItems.Filter = "�շ�ϸĿID=" & rsMoney!ID
                            If rsItems.EOF Then
                                '�����շ���Ŀ��Ӧ��ԭʼ��Ϣ
                                rsItems.AddNew
                                rsItems!����ID = mlng����ID
                                rsItems!��ҳID = Null
                                rsItems!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                                rsItems!�շ���� = rsMoney!���
                                rsItems!�շ�ϸĿID = rsMoney!ID
                                rsItems!������ = .TextMatrix(i, COL_����ҽ��)
                                rsItems!�������� = CStr(Sys.RowValue("���ű�", Val(.TextMatrix(i, COL_��������ID)), "����"))
                                
                                rsItems!���� = int���� * dbl����
                                rsItems!���� = dbl����
                                
                                rsҽ�����.Filter = "ҽ��ID=" & IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                If Not rsҽ�����.EOF Then
                                    rsItems!���id = rsҽ�����!���id
                                    rsItems!����id = rsҽ�����!����id
                                End If
                            Else
                                '����һ��ҽ��(������Ŀ)���շѶ��ղ������ظ����շ�ϸĿ
                                '������ͬһ�շ���Ŀ�Ĳ�ͬ������Ŀ��¼��ͬ
                                If rsMoney!�������� & "_" & rsMoney!ID <> str�շ���Ŀ Then
                                    rsItems!���� = NVL(rsItems!����, 0) + int���� * dbl����
                                End If
                                '���ۣ�ͬһ�շ���Ŀ�Ĳ�ͬ������Ŀ�ۼ�
                                If Val(.TextMatrix(i, COL_ID)) = rsItems!ҽ��ID Then
                                    rsItems!���� = NVL(rsItems!����, 0) + dbl����
                                End If
                            End If
                            rsItems!ʵ�ս�� = NVL(rsItems!ʵ�ս��, 0) + curʵ��
                            rsItems.Update
                        End If
                        
                        str�շ���Ŀ = rsMoney!�������� & "_" & rsMoney!ID
                        rsMoney.MoveNext
                    Loop
                End If
                
                '��ҽ�������л����ۿ۴���
                If gbln��������ۿ� And strHaveSub <> "" Then
                    rsSeek.Filter = 0
                    Do While Not rsSeek.EOF
                        rsSQL.Bookmark = rsSeek!�����ǩ
                        
                        str�ѱ� = NVL(mrsPati!�ѱ�) & IIF(gstr��̬�ѱ� <> "" And Not bln����, "," & gstr��̬�ѱ�, "")
                        If .TextMatrix(i, COL_��Ѽ���) = 1 Then
                            curʵ�� = 0
                        Else
                            curʵ�� = Format(ActualMoney(str�ѱ�, rsSeek!������ID, rsSeek!�ϼ�), gstrDec)
                        End If
                        
                        If InStr(str�ѱ�, ",") > 0 Then str�ѱ� = NVL(mrsPati!�ѱ�)
                        rsSQL!Sql = Set��̬�ѱ�(rsSQL!Sql, str�ѱ�)
                        
                        curʵ�� = curʵ�� - rsSeek!�ϼ� '���۲��
                        
                        'ҽ���ܿ�ʵʱ��⣺������Ŀ����滻
                        If Not IsNull(mrsPati!����) And blnʵʱ��� Then
                            rsItems.Filter = "�շ�ϸĿID=" & lng����ĿID
                            If Not rsItems.EOF Then
                                rsItems!ʵ�ս�� = NVL(rsItems!ʵ�ս��, 0) + curʵ��
                                rsItems.Update
                            End If
                        End If
                        
                        '����SQL�����滻
                        curʵ�� = Getʵ�ս��(rsSQL!Sql) + curʵ��
                        rsSQL!Sql = Setʵ�ս��(rsSQL!Sql, curʵ��)
                        rsSQL.Update
                    
                        rsSeek.MoveNext
                    Loop
                End If
                
                '����ҽ�����ͼ�¼
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then '����������(��ҩ;�����䷽�巨���÷�,�ɼ���������Ѫ;������Ϊ)
                    'ҽ���ķ��ͺ��Լ����ʾ
                    strSQL = "Select zl_AdviceSendCheck([1],[2]) as ��� From Dual"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceSendCheck", Val(.TextMatrix(i, COL_ID)), Replace(strժҪ, "''", "'"))
                    If Not rsTmp.EOF Then
                        strTmp = NVL(rsTmp!���)
                        If strTmp <> "" Then
                            Select Case Val(Split(strTmp, "|")(0))
                            Case 1 '��ʾ
                                If MsgBox(Split(strTmp, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    GoTo FuncEnd
                                End If
                            Case 2 '��ֹ
                                MsgBox Split(strTmp, "|")(1), vbInformation, gstrSysName
                                GoTo FuncEnd
                            End Select
                        End If
                    End If
                    
                    'һ��Ҫ��������NO
                    Call GetCurBillSet(strNOKey, strNO, -1, lng�������, bln����)
                                                            
                    '�Ƿ�һ��ҽ���ĵ�һҽ����:ҩ�Ƶĵ�һҩƷ��Ϊ��һҽ����
                    blnFirst = False
                    If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = True
                        End If
                    ElseIf .TextMatrix(i, COL_�������) = "C" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = True '��������еĵ�һ������
                        End If
                    ElseIf InStr(",1,2,3,4,5,", CLng(.Cell(flexcpData, i, COL_ID))) = 0 Then '�ſ���ҩ;������ҩ�巨����ҩ�÷����ɼ���������Ѫ;��
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            blnFirst = True
                        End If
                    End If
                                        
                    '��������:ҩƷΪ������λ������,����Ϊ����
                    If .TextMatrix(i, COL_�������) = "7" Then
                        dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))
                    ElseIf InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_�����װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                    Else
                        dbl�������� = Val(.TextMatrix(i, COL_����))
                    End If
                    dbl�������� = Format(dbl��������, "0.00000")
                                                            
                    '��ĩʱ��
                    str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                    If str�ֽ�ʱ�� <> "" Then
                        str�״�ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        strĩ��ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(UBound(Split(str�ֽ�ʱ��, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        '�޷��ֽ��Ϊ"һ����"��������Ϊ��ʼִ��ʱ�䣨74366��
                        str�״�ʱ�� = "To_Date('" & .TextMatrix(i, COL_��ʼʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                        strĩ��ʱ�� = "To_Date('" & .TextMatrix(i, COL_��ʼʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    
                    If Not gbln�������������� Then strCuvetteNumber = ""
                    'ԤԼ��Ժҽ��
                    If .TextMatrix(i, COL_�������) = "Z" And .TextMatrix(i, COL_��������) = "2" Then
                        lngԤԼ���� = SvrԤԼ��Ժ����(0)
                        mblnԤԼ���� = lngԤԼ���� = 1
                        If mblnԤԼ���� Then mlngԤ��Ժҽ��ID = Val(.TextMatrix(i, COL_ID))
                    Else
                        lngԤԼ���� = 0
                    End If
                    rsSQL.AddNew
                    rsSQL!���� = 3: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    rsSQL!NO = strNO
                    rsSQL!Sql = "ZL_����ҽ������_Insert(" & _
                        Val(.TextMatrix(i, COL_ID)) & "," & lng���ͺ� & "," & IIF(bln����, 2, 1) & ",'" & strNO & "'," & _
                        lng������� & "," & ZVal(dbl��������) & "," & str�״�ʱ�� & "," & strĩ��ʱ�� & "," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        intִ��״̬ & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & int�Ʒ�״̬ & "," & _
                        IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIF(InStr(str����ҩ��, "," & Val(.TextMatrix(i, COL_ID)) & "'") > 0, str����ҩ��, "Null") & "," & lngԤԼ���� & ")"
                    rsSQL.Update
                    str����ҩ�� = "''"
                    If gblnѪ��ϵͳ And .TextMatrix(i, COL_�������) = "K" Then
                        rsSQL.AddNew
                        rsSQL!���� = 9
                        rsSQL!��ĿID = 0
                        rsSQL!��� = 0
                        rsSQL!Sql = "Zl_ѪҺ��Ѫ����_Insert(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        rsSQL.Update
                    End If
                    
                    'ҽ��ִ�мƼ�
                    If rsExec.RecordCount > 0 Then
                        rsExec.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID)) & " And ���ͺ�=" & lng���ͺ�
                        If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                        Do While Not rsExec.EOF
                            rsSQL.AddNew
                            rsSQL!���� = 8
                            rsSQL!��ĿID = 0
                            rsSQL!��� = 0
                            rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                            rsSQL!Sql = "Zl_ҽ��ִ�мƼ�_Insert(" & rsExec!ҽ��ID & "," & rsExec!���ͺ� & ",To_date('" & _
                            rsExec!Ҫ��ʱ�� & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!�շ�ϸĿID & "")) & "," & rsExec!���� & ")"
                            rsSQL.Update
                            rsExec.MoveNext
                        Loop
                        rsExec.Filter = 0
                    End If
                    
                    'Ҫ���͵���δǩ����ҽ��ID(��ID,һ���еĶ���Ҳ�ᱻǩ��)
                    If Val(.TextMatrix(i, COL_ǩ��ID)) = 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            lng��ID = Val(.TextMatrix(i, COL_���ID))
                        Else
                            lng��ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(strҽ��IDs & ",", "," & lng��ID & ",") = 0 Then
                            strҽ��IDs = strҽ��IDs & "," & lng��ID
                        End If
                    End If
                End If
                
                '������ҩ�䷽��
                If .Cell(flexcpData, i, COL_ID) = 3 Then '��ҩ�÷�
                    int�䷽�� = int�䷽�� + 1
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
        Next
        
        '�Զ����е���ǩ��(δǩ������)
        '-----------------------------------------------------------------------------------------
        If Not gobjESign Is Nothing And CheckSign(IIF(mlngҽ������ID <> 0, 3, 0), 0, mlngҽ������ID, mlng�������ID, 1, , gobjESign) And strҽ��IDs <> "" Then
            strҽ��IDs = Mid(strҽ��IDs, 2) '��������ID,����Ϊ��ϸ��ID
            intRule = ReadAdviceSignSource(1, mlng����ID, mstr�Һŵ�, strҽ��IDs, 0, False, strSource, mstrǰ��IDs)
            If intRule = 0 Then GoTo FuncEnd
            If strSource = "" Then
                Screen.MousePointer = 0
                MsgBox "���ܶ�ȡҪǩ����ҽ��Դ�ġ�", vbInformation, gstrSysName
                GoTo FuncEnd
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign = "" Then GoTo FuncEnd
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
            rsSQL.AddNew
            rsSQL!���� = 4: rsSQL!ҽ��ID = 0: rsSQL!��ĿID = 0: rsSQL!��� = 0
            rsSQL!Sql = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��IDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            rsSQL.Update
        End If
        
        
        'ҽ���ܿ�ʵʱ���
        If Not IsNull(mrsPati!����) And blnʵʱ��� Then
            rsItems.Filter = 0
            If Not rsItems.EOF Then
                If Not gclsInsure.CheckItem(mrsPati!����, 0, 2, rsItems, Replace(strժҪ, "''", "'")) Then GoTo FuncEnd
            End If
        End If
        str����ҽ��IDs = Mid(str����ҽ��IDs, 2)
        str���ҽ��IDs = Mid(str���ҽ��IDs, 2)
        '�ύ��������
        '-----------------------------------------------------------------------------------------
        If Not CompletePatiSend(bln����, rsSQL, cur�ϼ�, str���, bln����, blnTran, cur���ʺϼ�, lng���ͺ�, str����ҽ��IDs, str���ҽ��IDs, CStr(curDate)) Then GoTo errH
    End With
    SendAdvice = lng���ͺ�
    '������ҽӿ�
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        On Error Resume Next
        Call gobjPlugIn.AdviceSendEnd(glngSys, p����ҽ���´�, lng���ͺ� & "")
        Call zlPlugInErrH(err, "AdviceSendEnd")
        On Error GoTo 0
    End If
FuncEnd:
    'ɾ�������ѳɹ����͵���
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    If blnTran Then
        gcnOracle.RollbackTrans
    End If
    If err.Number <> 0 Then
        Screen.MousePointer = 0
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        Screen.MousePointer = 0
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function CompletePatiSend(ByVal bln���� As Boolean, rsSQL As ADODB.Recordset, _
    ByVal cur�ϼ� As Currency, ByVal str��� As String, _
    ByVal bln���� As Boolean, blnTran As Boolean, _
    ByVal cur���ʺϼ� As Currency, ByVal lng���ͺ� As Long, ByVal str����ҽ��IDs As String, _
    ByVal str���ҽ��IDs As String, Optional ByVal strCurDate As String) As Boolean
'���ܣ��ύһ�����˵�ҽ����������,����֮ǰ������ʱ���
'������
'      bln����=�Ƿ�ȫ�����ö��ǻ���ģʽ�����ڱ��������⴦��
'      cur�ϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼ�,�������ʻ��۵��Ľ��
'      cur���ʺϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼƣ���������ִ�к��Զ���˵Ļ��۷��ã������������۷���
'      str���=���˱��η��ͼ��ʷ��õ��շ����,���ڼ��ʱ���
'      lng���ͺ�=���η��͵����ؼ���
'      str����ҽ��IDs=һ��ͨ�����ҽ��ID��
'      str���ҽ��IDs=��Ҫ�Զ�ִ����ɵ�ҽ��ID��
'˵�����������,���ڵ��ú����д���,blnTran�����Ƿ�����������
    Dim rsWarn As New ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, intR As Integer, lng��ID As Long, strҽ��IDs As String
    Dim cur���� As Currency, cur��� As Currency, i As Long
    Dim arrNOs() As String, strDiag As String, strAdviceInfo As String
    Dim arrSQL As Variant, arrAdviceID As Variant
    Dim strErr As String
    Dim bln����OK As Boolean
    Dim blnClearPatiCache As Boolean
    Dim blnPlugIn As Boolean
    Dim rsAdviceRis As ADODB.Recordset
    Dim strAdvices��Ѫ As String
    Dim var��Ѫ As Variant
    
'    ������ҽӿڷ���ǰ���ҽ������
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        blnPlugIn = True
        On Error Resume Next
        blnPlugIn = gobjPlugIn.AdviceCheckSendFee(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, cur�ϼ�, mint����)
        If Not blnPlugIn And err.Number <> 0 Then blnPlugIn = True
        Call zlPlugInErrH(err, "AdviceCheckSendFee")
        err.Clear: On Error GoTo 0
        If Not blnPlugIn Then
            Exit Function
        End If
    End If
    
    '���˷��ñ���
    blnClearPatiCache = True
    If bln���� And cur�ϼ� > 0 Then
        If InitObjPublicExpense Then
            For i = 1 To Len(str���)
                Call gobjPublicExpense.zlBillingWarn.zlBillingWarnCheck(Me, 0, IIF(bln����, 1, 0), mlng����ID, 0, 0, Mid(str���, i, 1), IIF(gbln�����������۷���, cur�ϼ�, cur���ʺϼ�), False, False, blnClearPatiCache, intR, , , , True)
                blnClearPatiCache = False
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        If bln���� And gdblԤ��������鿨 <> 0 And cur���ʺϼ� > 0 Then
            If Not zlDatabase.PatiIdentify(Me, glngSys, mlng����ID, cur���ʺϼ�, , , , IIF(-1 * gdblԤ��������鿨 >= Val(cur���ʺϼ�), False, True), , , (gdblԤ��������鿨 <> 0), (2 = gdblԤ��������鿨)) Then Exit Function
        End If
        Call InitObjLis(p����ҽ��վ)
        
        '�ȵ���LIS����ӿ�
        If Not gobjLIS Is Nothing Then
            strAdviceInfo = Get����ҽ����Ϣ
            If strAdviceInfo <> "" Then
                Set rsTmp = Get������ϼ�¼(mlng����ID, mlng�Һ�ID, "1")
                If rsTmp.RecordCount > 0 Then strDiag = rsTmp!�������
            End If
        End If
        
        If gblnѪ��ϵͳ Then
            If InitObjBlood(True) Then
                strAdvices��Ѫ = Get��Ѫҽ����Ϣ
                If strAdvices��Ѫ <> "" Then
                    var��Ѫ = Split(strAdvices��Ѫ, ",")
                End If
            End If
        End If
                
        Call ReplaceTrueNO(rsSQL)
        'ִ��˳��:�Ƽ�,����,����,ǩ��,����
        '1.�Է��ü�¼���շ�ϸĿID�������
        rsSQL.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        rsSQL.Sort = "����,��ĿID,���"

        gcnOracle.BeginTrans: blnTran = True
        'ִ��HIS�����ύ
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!Sql, Me.Caption)
            rsSQL.MoveNext
        Loop
                            
        '����LIS����ӿ�
        If strAdviceInfo <> "" Then
            If gobjLIS.SendLisApplicationForm(strAdviceInfo, strDiag) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                Call Del��������
                MsgBox "����ӿڵ���ʧ�ܣ����ܷ��ͼ���ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        'ҽ�������ϴ��ӿ�(������������)
        If mint���� <> 0 Then
            If gclsInsure.GetCapability(support�ϴ����ﵵ��, mlng����ID, mint����) Then
                If Not gclsInsure.TranElecDossier(1, mlng����ID, mlng�Һ�ID, mint����) Then Exit Function
            End If
        End If
        If strAdvices��Ѫ <> "" Then
            For i = 0 To UBound(var��Ѫ)
                If gobjPublicBlood.AdviceOperation(p����ҽ���´�, Val(var��Ѫ(i)), 5, False, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName

                    Exit Function
                End If
            Next
        End If
        gcnOracle.CommitTrans: blnTran = False
        Screen.MousePointer = 0
        
        'һ��ͨ����(������ɺ���ý��㣬����ɹ����ٵ���ִ�У�ȡ����������ʧ�ܣ�����ִ��)
        If str����ҽ��IDs <> "" Then
            If gobjSquareCard.zlSquareAffirm(Me, p����ҽ���´�, GetInsidePrivs(p����ҽ���´�), mlng����ID, mlngCardType, False, IIF(bln����, 2, 1), , str����ҽ��IDs, mstr֧����ʽ, , mblnʹ��Ԥ��) Then
                
                bln����OK = True
                
                arrSQL = Array()
                arrAdviceID = Split(str���ҽ��IDs, ",")
                
                For i = 0 To UBound(arrAdviceID)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����ҽ��ִ��_Finish(" & arrAdviceID(i) & "," & lng���ͺ� & ",Null,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIF(mlngҽ������ID <> 0, mlngҽ������ID, mlng�������ID) & ")"
                Next
                                
                gcnOracle.BeginTrans: blnTran = True
                For i = 0 To UBound(arrSQL)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
                Next
                gcnOracle.CommitTrans: blnTran = False
            End If
        End If
        If Not (mclsMipModule Is Nothing) Then
            If mclsMipModule.IsConnect Then
                rsSQL.Filter = "�������='ҩƷ'"
                If Not rsSQL.EOF Then
                    Call SendMsgҩƷҽ������(rsSQL, "," & str����ҽ��IDs & ",", bln����OK, IIF(bln����, 1, 2), lng���ͺ�, strCurDate)
                End If
                Call SendMsg����(lng���ͺ�, IIF(bln����, 1, 2))
            End If
        End If
        Call BulidBarCode(lng���ͺ�)
        'RIS�ӿ�
        If HaveRIS Then
            If GetAdviceRis(rsAdviceRis) Then
                On Error Resume Next
                If gobjRis.HISSendAdvice(rsAdviceRis, 1, mlng����ID, 0, mstr�Һŵ�, lng���ͺ�) <> 1 Then
                    MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISSendAdvice)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo 0
            End If
        ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������RIS�ӿڴ���ʧ��δ����(HISSendAdvice)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
        'ԤԼ���ķ������
        If mblnԤԼ���� And mlngԤ��Ժҽ��ID <> 0 Then
            Call SvrԤԼ��Ժ����(1)
        End If
        '�ύ�ɹ�,������ҽ���б��Ϊ��ɾ��
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    .RowData(i) = -1
                End If
            Next
            '������ҽӿ�
            If CreatePlugInOK(p����ҽ���´�, mint����) Then
                On Error Resume Next
                Call gobjPlugIn.AdviceSend(glngSys, p����ҽ���´�, mlng����ID, mlng�Һ�ID, lng���ͺ�)
                Call zlPlugInErrH(err, "AdviceSend")
                On Error GoTo 0
            End If
            If gobjExchange Is Nothing Then
                On Error Resume Next
                Set gobjExchange = CreateObject("zlExchange.clsExchange")
                If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
                err.Clear: On Error GoTo 0
            End If
            '�������ݽ���ƽ̨����LIS,PACS�������뵥
            If Not gobjExchange Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        'c-����,d-���
                        If .TextMatrix(i, COL_�������) = "C" Or .TextMatrix(i, COL_�������) = "D" Then
                            If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                lng��ID = Val(.TextMatrix(i, COL_���ID))
                            Else
                                lng��ID = Val(.TextMatrix(i, COL_ID))
                            End If
                            If InStr(strҽ��IDs & ",", "," & lng��ID & ",") = 0 Then
                                strҽ��IDs = strҽ��IDs & "," & lng��ID
                                Call gobjExchange.SendMsg(IIF(.TextMatrix(i, COL_�������) = "C", 1, 2), "����ID::" & mlng����ID & "||��ҳID::0||ҽ��ID::" & lng��ID & "||��������::1")
                            End If
                        End If
                    End If
                Next
            End If
        End With
        
        CompletePatiSend = True
    End If
End Function

Private Sub SendMsgҩƷҽ������(ByVal rsIn As ADODB.Recordset, ByVal str����ҽ��IDs As String, ByVal bln����OK As Boolean, ByVal int�������� As Integer, ByVal lng���ͺ� As String, ByVal str����ʱ�� As String)
'����ҩƷҽ�����ͺ����ҩ����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim rsMsg As ADODB.Recordset
    Dim intB As Integer
    Dim intE As Integer
    Dim i As Long
    Dim j As Long
    Dim lngRow As Long
    Dim strNO As String
    Dim blnKey As Boolean
    Dim byt�շ� As Byte
    Dim str���ݺ� As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    Set rsMsg = New ADODB.Recordset
    rsMsg.Fields.Append "ҽ��ID", adBigInt
    rsMsg.Fields.Append "���ID", adBigInt
    rsMsg.Fields.Append "ҽ������", adVarChar, 120
    rsMsg.Fields.Append "ִ��Ƶ��", adVarChar, 60
    rsMsg.Fields.Append "��ҩ;��id", adBigInt
    rsMsg.Fields.Append "��ҩ;��", adVarChar, 60
    rsMsg.Fields.Append "��ʼʱ��", adVarChar, 60
    rsMsg.Fields.Append "����", adDouble
    rsMsg.Fields.Append "����", adDouble
    rsMsg.Fields.Append "ҽ������", adVarChar, 120
    rsMsg.Fields.Append "Ʒ��ID", adBigInt
    rsMsg.Fields.Append "ҩƷ���", adVarChar, 6
    rsMsg.Fields.Append "ҩƷID", adBigInt
    rsMsg.Fields.Append "ִ�в���id", adBigInt
    rsMsg.CursorLocation = adUseClient
    rsMsg.LockType = adLockOptimistic
    rsMsg.CursorType = adOpenStatic
    rsMsg.Open
    
    Set rsTmp = zlDatabase.CopyNewRec(rsIn)
    
    With vsAdvice
        For i = 1 To rsIn.RecordCount
            If strNO <> rsIn!NO Then
                strNO = rsIn!NO: rsTmp.Filter = "NO='" & strNO & "'"
                If Not rsTmp.EOF Then
                    str���ݺ� = "": strTmp = rsIn!Sql
                    intB = InStr(strTmp, "'") + 1
                    intE = InStr(intB, strTmp, "'")
                    str���ݺ� = Mid(strTmp, intB, intE - intB)
                    For j = 1 To rsTmp.RecordCount
                        lngRow = Val(Split(rsTmp!���� & "", "_")(0))
                        rsMsg.AddNew
                        rsMsg!ҽ��ID = .TextMatrix(lngRow, COL_ID)
                        rsMsg!���ID = Val(.TextMatrix(lngRow, COL_���ID))
                        rsMsg!ҽ������ = .TextMatrix(lngRow, col_ҽ������)
                        rsMsg!ִ��Ƶ�� = .TextMatrix(lngRow, COL_Ƶ��)
                        rsMsg!��ҩ;��ID = Val(.TextMatrix(lngRow, COL_���ID))
                        rsMsg!��ҩ;�� = .TextMatrix(lngRow, COL_�÷�)
                        rsMsg!��ʼʱ�� = .TextMatrix(lngRow, COL_��ʼʱ��)
                        rsMsg!���� = .TextMatrix(lngRow, COL_����)
                        rsMsg!���� = .TextMatrix(lngRow, COL_����)
                        rsMsg!ҽ������ = .TextMatrix(lngRow, COL_ҽ������)
                        rsMsg!Ʒ��ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                        rsMsg!ҩƷ��� = rsTmp!�������
                        rsMsg!ҩƷID = Val(Split(rsTmp!���� & "", "_")(1))
                        rsMsg!ִ�в���ID = Val(Split(rsTmp!���� & "", "_")(2))
                        rsMsg.Update
                        If bln����OK And InStr(str����ҽ��IDs, "," & Val(.TextMatrix(lngRow, COL_ID)) & ",") = 0 And blnKey = False Then blnKey = True
                        rsTmp.MoveNext
                    Next
                End If
                
                byt�շ� = 1
                If bln����OK And Not blnKey Then byt�շ� = 2
                
                '������Ϣ
                If rsMsg.RecordCount > 0 Then
                    rsMsg.MoveFirst
                    Call ZLHIS_CIS_006(mclsMipModule, mlng����ID, mstr����, , mstr�����, 1, mlng�Һ�ID, mlng�������ID, "", , , lng���ͺ�, str����ʱ��, _
                        UserInfo.����, str���ݺ�, int��������, byt�շ�, rsMsg)
                    rsMsg.MoveFirst
                    For j = 1 To rsMsg.RecordCount
                        rsMsg.Delete
                        rsMsg.MoveNext
                    Next
                End If
            End If
            rsIn.MoveNext
        Next
    End With
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SendMsg����(ByVal lng���ͺ� As Long, ByVal int�������� As Integer)
    Dim strIDs As String
    Dim lngTmp As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    Dim i As Long
    Dim j As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '���밲��
                If Val(.TextMatrix(i, COL_ִ�а���)) = 1 Then
                    Call ZLHIS_CIS_004(mclsMipModule, mlng����ID, mstr����, , mstr�����, 1, _
                        mlng�Һ�ID, .TextMatrix(i, COL_���˿���ID), "", , , Val(.TextMatrix(i, COL_ID)), 1, .TextMatrix(i, COL_�������), .TextMatrix(i, COL_��������), _
                        lng���ͺ�, .TextMatrix(i, COL_ִ�п���ID))
                End If
                '����ҽ��
                If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 6 Then
                    strIDs = "": lngTmp = 0
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "C" Then
                                strIDs = strIDs & "," & Val(.TextMatrix(j, COL_ID))
                                lngTmp = Val(.TextMatrix(j, COL_ִ�п���ID))
                            End If
                        End If
                    Next
                    strIDs = Mid(strIDs, 2)
                    If strIDs <> "" Then
                        Call ZLHIS_CIS_016(mclsMipModule, mlng����ID, mstr����, , mstr�����, 1, mlng�Һ�ID, mlng�������ID, , Val(.TextMatrix(i, COL_ID)), _
                            .TextMatrix(i, COL_�걾��λ), .TextMatrix(i, COL_������ĿID), , .TextMatrix(i, COL_ִ�п���ID), , strIDs, , lngTmp, , lng���ͺ�, "", _
                            int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_��ʼʱ��), .TextMatrix(i, COL_��������ID), , "")
                    End If
                '�������
                ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    strTmp1 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "D" Then
                                strTmp1 = strTmp1 & "," & .TextMatrix(j, COL_�걾��λ)
                            End If
                        End If
                    Next
                    strTmp1 = Mid(strTmp1, 2)
                    Call ZLHIS_CIS_017(mclsMipModule, mlng����ID, mstr����, , mstr�����, 1, mlng�Һ�ID, Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), _
                        .TextMatrix(i, COL_������ĿID), .TextMatrix(i, col_ҽ������), strTmp1, .TextMatrix(i, COL_ִ�п���ID), , lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_��ʼʱ��), .TextMatrix(i, COL_��������ID), , "")
                '��������
                ElseIf .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    strTmp1 = Getҽ����������(Val(.TextMatrix(i, COL_ID)), "����ҽ��")
                    strTmp2 = Getҽ����������(Val(.TextMatrix(i, COL_ID)), "����ҽ��")
                    strIDs = "": lngTmp = 0
                    strIDs = strIDs & "," & Val(.TextMatrix(i, COL_ID))
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "F" Then
                                strIDs = strIDs & "," & .TextMatrix(j, COL_ID)
                            ElseIf .TextMatrix(j, COL_�������) = "G" Then
                                lngTmp = Val(.TextMatrix(j, COL_ID))
                            End If
                        End If
                    Next
                    strIDs = Mid(strIDs, 2)
                    Call ZLHIS_CIS_018(mclsMipModule, mlng����ID, mstr����, , mstr�����, 1, _
                        mlng�Һ�ID, mlng�������ID, "", Val(.TextMatrix(i, COL_ID)), strIDs, , lngTmp, , strTmp1, strTmp2, .TextMatrix(i, COL_ִ�п���ID), , lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_��ʼʱ��), .TextMatrix(i, COL_��������ID), , "")
                End If
            End If
        Next
    End With
End Sub

Private Sub ShowSendTotal()
'���ܣ����ݵ�ǰѡ��Ҫ���͵�ҽ�������㲢��ʾ���͵�ҽ���ϼ�
    Dim cur��� As Currency, curҩƷ��� As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '�ɼ��еĽ��:��һ��Ļ��ܽ��
                If Not .RowHidden(i) Then
                    cur��� = cur��� + Val(.TextMatrix(i, COL_���))
                End If
                'ҩƷ�Ľ��,ȡԭʼ���
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    curҩƷ��� = curҩƷ��� + Val(.Cell(flexcpData, i, COL_���))
                End If
            End If
        Next
    End With
    stbThis.Panels(5).Text = "���:" & FormatEx(cur���, gbytDec) & "(ҩ" & FormatEx(curҩƷ���, gbytDec) & ")"
    Call Form_Resize
End Sub

Private Sub Del��������()
'���ܣ�ҽ������ʧ�ܣ�������˺󣬵��ü�������ɾ���ӿ�
    Dim i As Long, strҽ��IDs As String, strErr As String
        
    '�ռ��ɼ�����
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strҽ��IDs = strҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    
    If strҽ��IDs <> "" Then
        strҽ��IDs = Mid(strҽ��IDs, 2)
        Call InitObjLis(p����ҽ��վ)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(strҽ��IDs, strErr) = False Then
                MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function Get����ҽ����Ϣ() As String
'���ܣ���ȡ����ҽ����Ϣ�����ݸ�����ӿڳ���
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                '����ҽ��ID1,�ɼ�ҽ��ID1,ִ�п���ID1,�걾1;.....
                'LIS�ӿڲ����ļ��飬һ���ɼ���ʽֻ��һ������ҽ����û��һ���ɼ��������
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strInfo = strInfo & ";" & .TextMatrix(i - 1, COL_ID) & "," & .TextMatrix(i, COL_ID) & "," & .TextMatrix(i - 1, COL_ִ�п���ID) & "," & .TextMatrix(i - 1, COL_�걾��λ)
                End If
            End If
        Next
    End With
    Get����ҽ����Ϣ = Mid(strInfo, 2)
End Function

Private Function Get��Ѫҽ����Ϣ() As String
'���ܣ���ȡ��Ѫҽ����Ϣ�����ݸ��ӿڳ��򣬽�ȡ��ҽ��ID
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_�������) = "K" Then
                '����ҽ��ID1,�ɼ�ҽ��ID1,ִ�п���ID1,�걾1;.....
                'LIS�ӿڲ����ļ��飬һ���ɼ���ʽֻ��һ������ҽ����û��һ���ɼ��������
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strInfo = strInfo & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Get��Ѫҽ����Ϣ = Mid(strInfo, 2)
End Function

Private Sub SetFontSize(ByVal bytSize As Byte)
'���ܣ����н��������ͳһ����
'������bytSize  0-9�����壬1-12������
    Call zlControl.SetPubFontSize(Me, bytSize)
    Me.Width = IIF(bytSize = 0, 10000, 11000)
    Me.Height = IIF(bytSize = 0, 7000, 8800)
End Sub

Private Function zlPluginAdviceBeforeSend() As Boolean
'���ܣ�ҽ������ǰ������Һ�
    Dim i As Long, j As Long
    Dim strAdviceIDs As String, strMsg  As String
    Dim rsDataPlugIn As ADODB.Recordset
    Dim lng���� As Long
    Dim str�ֽ�ʱ�� As String, strTmp As String
    
    zlPluginAdviceBeforeSend = True
    
    '������ҽӿڣ�ҽ������ǰ�ļ��
    If CreatePlugInOK(p����ҽ���´�, mint����) Then
        Call InitPlugInRs(rsDataPlugIn)
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                        str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                    Else
                        str�ֽ�ʱ�� = .Cell(flexcpData, i, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                    End If
                    rsDataPlugIn.AddNew
                    rsDataPlugIn!����ID = mlng����ID
                    rsDataPlugIn!����ID = mlng�Һ�ID
                    rsDataPlugIn!�Һŵ� = mstr�Һŵ�
                    rsDataPlugIn!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                    rsDataPlugIn!���ID = Val(.TextMatrix(i, COL_���ID))
                    rsDataPlugIn!�շ�ϸĿID = Val(.TextMatrix(i, COL_�շ�ϸĿID))
                    rsDataPlugIn!�ֽ�ʱ�� = str�ֽ�ʱ��
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!������λ = .TextMatrix(i, COL_������λ)
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!������λ = .TextMatrix(i, COL_������λ)
                    rsDataPlugIn!���� = mint����
                    rsDataPlugIn.Update
                End If
            Next
            If rsDataPlugIn.RecordCount > 0 Then rsDataPlugIn.MoveFirst
            strAdviceIDs = "": strMsg = ""
            On Error Resume Next
            Call gobjPlugIn.AdviceBeforeSend("", rsDataPlugIn, strAdviceIDs, strMsg)
            Call zlPlugInErrH(err, "AdviceBeforeSend")
            err.Clear
            On Error GoTo 0
             
            If strAdviceIDs <> "" Then
                strTmp = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If InStr("," & strAdviceIDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                            If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                                j = Val(.TextMatrix(i, COL_ID))
                            Else
                                j = Val(.TextMatrix(i, COL_���ID))
                            End If
                            
                            If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                                strTmp = strTmp & "," & j
                            End If
                        End If
                    End If
                Next
                strAdviceIDs = Mid(strTmp, 2)
                lng���� = 0
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            j = Val(.TextMatrix(i, COL_ID))
                        Else
                            j = Val(.TextMatrix(i, COL_���ID))
                        End If
                        lng���� = lng���� + 1
                        If InStr("," & strAdviceIDs & ",", "," & j & ",") > 0 Then
                            .Cell(flexcpData, i, COL_ѡ��) = 1
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                            lng���� = lng���� - 1
                        End If
                    End If
                Next
                
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                If lng���� = 0 Then
                    MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
                    zlPluginAdviceBeforeSend = False
                End If
            End If
        End With
    End If
End Function

Private Sub BulidBarCode(ByVal lng���ͺ� As Long)
'���ܣ�ҽ�����͵��ӿ����ɶ�ά���������Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strNOs As String
    Dim int��¼���� As Integer
    Dim strExpand As String

    If gobjSquareCard Is Nothing Then
        On Error Resume Next
        Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        If Not gobjSquareCard Is Nothing Then
            If gobjSquareCard.zlInitComponents(Me, p����ҽ���´�, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set gobjSquareCard = Nothing
            End If
        End If
    End If

    On Error GoTo errH
    If Not gobjSquareCard Is Nothing Then
        strSQL = "Select ��¼����, NO From ����ҽ������ Where ���ͺ� =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���ͺ�)
        If Not rsTmp.EOF Then
            int��¼���� = Val("" & rsTmp!��¼����)
            For i = 1 To rsTmp.RecordCount
                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNOs = strNOs & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Next
            strNOs = Mid(strNOs, 2)
        End If
        Call gobjSquareCard.zlAdviceSendBulidBarCode(Me, p����ҽ���´�, 0, int��¼����, strNOs, strExpand)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetAdviceRis(ByRef rsData As ADODB.Recordset) As Boolean
'���ܣ���ȡ���͵�RIS��ҽ����Ϣ
    Dim i As Long
    
    On Error GoTo errH
    
    Set rsData = New ADODB.Recordset
    
    rsData.Fields.Append "ҽ��ID", adBigInt
    rsData.Fields.Append "��������ID", adBigInt
    rsData.Fields.Append "ִ�п���ID", adBigInt
    rsData.Fields.Append "������ĿID", adBigInt
    rsData.Fields.Append "������Դ", adInteger '1-����;2-סԺ;
    rsData.Fields.Append "���", adVarChar, 10
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If InStr(",D,F,", .TextMatrix(i, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_�������) = "E" Then
                    If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                        rsData.AddNew
                        rsData!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                        rsData!��������id = Val(.TextMatrix(i, COL_��������ID))
                        rsData!ִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                        rsData!������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                        rsData!������Դ = 1
                        rsData!��� = .TextMatrix(i, COL_�������)
                        rsData.Update
                    End If
                End If
            End If
        Next
    End With
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        GetAdviceRis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRISScheduling() As Boolean
'���ܣ������Ŀ�Ƿ��Ǳ���ԤԼ
    Dim i As Long
    Dim blnDo As Boolean
    Dim lngҽ��ID As Long
    Dim lng������ĿID As Long
    Dim lngRst As Long
    Dim strMsg As String
    
    CheckRISScheduling = True
    
    If HaveRIS Then
        If gbln����Ӱ����ϢϵͳԤԼ Then
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Function
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If Val(.TextMatrix(i, COL_������־)) <> 1 And mint���� <> 1 Then
                    If InStr(",D,F,", .TextMatrix(i, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_�������) = "E" Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            lngҽ��ID = Val(.TextMatrix(i, COL_ID))
                            lng������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                            lngRst = -1
                            lngRst = gobjRis.HISScheduling(1, lngҽ��ID, lng������ĿID, False)
                            If lngRst <> 0 Then
                            '�ӿڷ���ʧ�ܸ�����ʾ
                                .Cell(flexcpData, i, COL_ѡ��) = 1 '��ǰ��ֹѡ��
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                                Call RowSelectSame(i, COL_ѡ��)
                                strMsg = IIF("" = strMsg, "", strMsg & "��") & .TextMatrix(i, col_ҽ������)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
    If strMsg <> "" Then
        MsgBox "����������RISϵͳԤԼ���̣�" & vbCrLf & "��" & strMsg & "��" & _
                vbCrLf & "ҽ��û��ԤԼ��ԤԼ�ɹ�����ܷ��͡�", vbInformation, gstrSysName
        CheckRISScheduling = False
    End If
End Function

Private Sub FuncPassPharmReview()
'����:ҩʦ��ϵͳ
    Dim strGroupID As String
    Dim strMsg As String
    Dim i As Long, j As Long
    Dim dblAdviceID As Double
    
    If gobjPass Is Nothing Then Exit Sub
    
    With vsAdvice
        'ҩ����ID
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�������) = "E" And InStr(",2,4,", "," & .TextMatrix(i, COL_��������) & ",") > 0 Then
                    strGroupID = strGroupID & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    
        'û��ҩ������ʱ���������
        If strGroupID = "" Then Exit Sub
        strGroupID = Mid(strGroupID, 2)
        If Not gobjPass.zlPassPharmReview(mlng����ID, mlng�Һ�ID, mstr�Һŵ�, strGroupID) Then Exit Sub
        
        If strGroupID <> "" Then
            'ȡ��ѡ��
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    dblAdviceID = IIF(0 = Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    If InStr("," & strGroupID & ",", "," & dblAdviceID & ",") > 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                        .Cell(flexcpData, i, COL_ѡ��) = 1
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 And InStr(",5,6,", "," & .TextMatrix(i, COL_�������) & ",") > 0 Or (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "4") Then
                            If j <= 10 Then
                                strMsg = strMsg & vbCrLf & .TextMatrix(i, col_ҽ������)
                                j = j + 1
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If strMsg <> "" Then
            Call MsgBox("����ҽ��δͨ��������飬���ܷ��ͣ�" & strMsg, vbInformation, Me.Caption)
        End If
    End With
End Sub

Private Function Set������ҩ() As Boolean
'���ܣ�����ҩƷҽ���е�������ҩ˵��
    Dim i As Long
    Dim strMsg As String
    Dim str������ҩ As String
    Dim strSQL As String
    Dim strҽ��IDs As String
    
    On Error GoTo errH
    If mstrAdDrugIDs = "" Then
        Set������ҩ = True
        Exit Function
    End If
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If InStr("," & mstrAdDrugIDs & ",", "," & .TextMatrix(i, COL_ID) & ",") > 0 Then
                    strMsg = strMsg & "," & .TextMatrix(i, col_ҽ������)
                    strҽ��IDs = strҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    If strMsg = "" Then
        Set������ҩ = True
        Exit Function
    End If
    Call frmMsgDruExcess.ShowMe(Me, 1, Mid(strMsg, 2), str������ҩ)
    If str������ҩ = "*NULL*" Then
        Exit Function
    End If
    strSQL = "Zl_����ҽ����¼_������ҩ('" & Mid(strҽ��IDs, 2) & "','" & str������ҩ & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set������ҩ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function OutPatiFeeUsable(ByVal lng����ID As Long) As Boolean
'���ܣ����˵ĵ�ǰ�����Ƿ���Ч������true������ǰ�ѱ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnʧЧ As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select  Sysdate as ��ǰ,Nvl(b.��Ч��ʼ, To_Date('1900-01-01', 'yyyy-mm-dd')) as ��ʼ,Nvl(b.��Ч����, To_Date('3000-01-01', 'yyyy-mm-dd')) as ����  From ������Ϣ A, �ѱ� B Where a.�ѱ�=b.���� And a.����id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    OutPatiFeeUsable = True
    
    If rsTmp.EOF Then
        blnʧЧ = True
    Else
        If Not Between(Format(rsTmp!��ǰ, "YYYY-MM-DD"), Format(rsTmp!��ʼ, "YYYY-MM-DD"), Format(rsTmp!����, "YYYY-MM-DD")) Then
            blnʧЧ = True
        End If
    End If
    
    If blnʧЧ Then
        MsgBox "�ò��˵ĵ�ǰ�ѱ��Ѿ�ʧЧ�����ܷ���ҽ�������ڲ�����Ϣ�е������˷ѱ�", vbInformation, gstrSysName
        OutPatiFeeUsable = False
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SvrԤԼ��Ժ����(ByVal intType As Integer) As Long
'���ܣ�ԤԼ�з������,
'������
'      intType 0-��֤�Ե����ж��Ƿ����÷���1-����������
'���أ�0-ʧ�ܣ�1-�ɹ�

    Dim blnTmp As Boolean
    Dim strErr As String
    Dim strJsIn As String
    Dim strJsOut As String
    Dim lng�к� As Long
    Dim lng��������ҽ��ID As Long
    Dim str��������ҽ�� As String
    Dim lng�����������ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim rsAppend As ADODB.Recordset
    Dim i As Long
    Dim str���븽�� As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        blnTmp = Sys.NewSystemSvr("ԤԼ����", "סԺ����", strJsIn, strJsOut, strErr)
        SvrԤԼ��Ժ���� = IIF(blnTmp, 1, 0)
    Else
        strSQL = "select b.id as ����ҽ��id, a.����ҽ��,a.��������id,c.���� as ��������,To_Char(a.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & vbNewLine & _
            "a.ִ�п���id, d.���� as ִ�п���,To_Char(a.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI:SS') as ��ʼִ��ʱ��,e.��ͥ�绰,e.��ϵ�˵绰" & vbNewLine & _
            "from ����ҽ����¼ a,��Ա�� b,���ű� c,���ű� d,������Ϣ e" & vbNewLine & _
            "where a.����ҽ��=b.���� and a.��������id=c.id and a.ִ�п���id=d.id and a.����ID = e.����ID and a.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngԤ��Ժҽ��ID)
        If Not rsTmp.EOF Then
            strSQL = "select ����,��Ŀ,���� from ����ҽ������ where ҽ��id=[1] order by ����"
            Set rsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngԤ��Ժҽ��ID)
            If Not rsAppend.EOF Then
                For i = 1 To rsAppend.RecordCount
                    str���븽�� = IIF("" = str���븽��, "", str���븽�� & ",") & "��" & rsAppend!��Ŀ & "��" & rsAppend!����
                    rsAppend.MoveNext
                Next
            End If
            
            strJsIn = "{  ""input_in"": {" & vbNewLine & _
                    "  ""iba_reg_rec_id"": """"," & vbNewLine & _
                    "  ""rgst_id"": """ & mlng�Һ�ID & """," & vbNewLine & _
                    "  ""rgst_no"": """ & mstr�Һŵ� & """," & vbNewLine & _
                    "  ""pid"": """ & mlng����ID & """," & vbNewLine & _
                    "  ""pat_name"": """ & mstr���� & """," & vbNewLine & _
                    "  ""pat_sex"": """ & mrsPati!�Ա� & """," & vbNewLine & _
                    "  ""pat_age"": """ & mrsPati!���� & """," & vbNewLine & _
                    "  ""pat_brsdate"": """ & mrsPati!Birthdate & """," & vbNewLine & _
                    "  ""insure_sign"": """ & IIF(mint���� = 0, 0, 1) & """," & vbNewLine & _
                    "  ""outp_apply_dr_id"": """ & rsTmp!����ҽ��id & """," & vbNewLine & _
                    "  ""outp_apply_dr"": """ & rsTmp!����ҽ�� & """," & vbNewLine & _
                    "  ""outp_apply_dept_id"": """ & rsTmp!��������id & """," & vbNewLine & _
                    "  ""outp_apply_dept"": """ & rsTmp!�������� & """," & vbNewLine & _
                    "  ""outp_apply_time"": """ & rsTmp!����ʱ�� & ""","
            strJsIn = strJsIn & vbNewLine & _
                    "  ""iba_dept_id"": """ & rsTmp!ִ�п���ID & """," & vbNewLine & _
                    "  ""iba_dept"": """ & rsTmp!ִ�п��� & """," & vbNewLine & _
                    "  ""iba_time"": """ & rsTmp!��ʼִ��ʱ�� & """," & vbNewLine & _
                    "  ""harea_code"": """"," & vbNewLine & _
                    "  ""harea_name"": """"," & vbNewLine & _
                    "  ""outp_dept_id"": """"," & vbNewLine & _
                    "  ""outp_dept_name"": """"," & vbNewLine & _
                    "  ""iba_reg_sign"": ""0""," & vbNewLine & _
                    "  ""apply_item"": """ & str���븽�� & """," & vbNewLine & _
                    "  ""order_id"": """ & mlngԤ��Ժҽ��ID & """," & vbNewLine & _
                    "  ""home_phno"": """ & NVL(rsTmp!��ͥ�绰) & """," & vbNewLine & _
                    "  ""contacts_phno"": """ & NVL(rsTmp!��ϵ�˵绰) & """ " & vbNewLine & _
                "}}"
            Call Sys.NewSystemSvr("ԤԼ����", "סԺ����", strJsIn, strJsOut, strErr)
            If strErr <> "" Then
                MsgBox "ԤԼ��Ժ����:" & strErr, vbInformation, gstrSysName
            End If
        End If

    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Check�������()
'���ܣ������󷽽ӿ��жϵ�ǰҽ���ǲ���������
    Dim i As Long
    Dim str��ҩIDs As String '���뵽�ӿ��еĲ���
    Dim strOutҽ��IDs As String '���ܹ����͵�ҽ��ID
    Dim strErr As String
    Dim lngҽ��ID As Long
    Dim strҽ������ As String
    Dim blnTmp As Boolean
    Dim strҩ��ҽ��IDs As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If Not gbln��ϵͳ Then Exit Sub
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "2" Then
                    str��ҩIDs = str��ҩIDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
        If str��ҩIDs <> "" Then
            str��ҩIDs = Mid(str��ҩIDs, 2)
            '���ú����ò������ṩ�Ľӿ�
            blnTmp = gobjPass.ZLPharmReviewResultOut(Me, mlng����ID, mlng�Һ�ID, mstr�Һŵ�, str��ҩIDs, rsTmp, strErr)
            If blnTmp Then
                If strErr = "" Then
                    If Not rsTmp Is Nothing Then
                        If Not rsTmp.EOF Then
                            For i = 1 To rsTmp.RecordCount
                                If InStr("," & strOutҽ��IDs & ",", "," & rsTmp!���ID & ",") = 0 Then
                                    strOutҽ��IDs = strOutҽ��IDs & "," & rsTmp!���ID
                                End If
                                strҩ��ҽ��IDs = strҩ��ҽ��IDs & "," & rsTmp!ҽ��ID
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                    
                End If
            End If
        End If
        
        
        If strOutҽ��IDs <> "" Then
            'ȡ��ѡ��
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    lngҽ��ID = IIF(0 = Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    If InStr("," & strOutҽ��IDs & ",", "," & lngҽ��ID & ",") > 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                        .Cell(flexcpData, i, COL_ѡ��) = 1
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            If InStr("," & strҩ��ҽ��IDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                                strҽ������ = strҽ������ & vbCrLf & .TextMatrix(i, col_ҽ������)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        
        If strҽ������ <> "" Then
            Call MsgBox("����ҽ��δͨ��������飬���ܷ��ͣ�" & strҽ������, vbInformation, Me.Caption)
        End If
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
