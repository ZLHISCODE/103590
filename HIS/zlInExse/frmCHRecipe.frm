VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCHRecipe 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ҩ�䷽"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCHRecipe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   10635
      TabIndex        =   18
      ToolTipText     =   "�ȼ���F2"
      Top             =   240
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   10635
      TabIndex        =   19
      Top             =   750
      Width           =   1170
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   400
      Left            =   10635
      TabIndex        =   20
      Top             =   7725
      Width           =   1170
   End
   Begin VB.Frame fraInfo 
      Height          =   600
      Left            =   15
      TabIndex        =   22
      Top             =   15
      Width           =   10455
      Begin VB.Frame fraAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   8940
         TabIndex        =   24
         Top             =   465
         Width           =   330
      End
      Begin VB.TextBox txtAuto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   8970
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "5"
         Top             =   225
         Width           =   285
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "�Զ�ʶ��   λ������"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7680
         TabIndex        =   0
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ����ҩ��̬����������ÿζ�в�ҩ��������ѡ���в�ҩ�밴 * ����"
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   255
         Width           =   7560
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vs��ҩ��� 
      Height          =   1365
      Left            =   30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4785
      Width           =   10410
      _cx             =   18362
      _cy             =   2408
      Appearance      =   1
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCHRecipe.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Begin VB.CommandButton cmd��̬ 
         Caption         =   "��Ϊɢװ"
         Height          =   330
         Left            =   6675
         TabIndex        =   25
         Top             =   630
         Width           =   1245
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   21
      Top             =   8175
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmCHRecipe.frx":061C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Object.Tag             =   "��ҩζ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   15
      TabIndex        =   26
      Top             =   495
      Width           =   10455
      Begin VB.PictureBox pic��̬ 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   45
         ScaleHeight     =   390
         ScaleWidth      =   10260
         TabIndex        =   27
         Top             =   165
         Width           =   10260
         Begin VB.OptionButton opt��̬ 
            Caption         =   "ɢװ(&0)"
            Height          =   420
            Index           =   0
            Left            =   750
            TabIndex        =   3
            Top             =   -15
            Width           =   1245
         End
         Begin VB.OptionButton opt��̬ 
            Caption         =   "��Ƭ(&1)"
            Height          =   420
            Index           =   1
            Left            =   1980
            TabIndex        =   4
            Top             =   -15
            Width           =   1245
         End
         Begin VB.OptionButton opt��̬ 
            Caption         =   "����(&2)"
            Height          =   420
            Index           =   2
            Left            =   3225
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.ComboBox cboҩ�� 
            Height          =   360
            Left            =   7635
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   30
            Width           =   2625
         End
         Begin VB.TextBox txt���� 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   6315
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "1"
            Top             =   45
            Width           =   495
         End
         Begin VB.Label lbl��̬ 
            AutoSize        =   -1  'True
            Caption         =   "��̬"
            Height          =   240
            Left            =   90
            TabIndex        =   2
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lblҩ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ��"
            Height          =   240
            Left            =   7035
            TabIndex        =   8
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   240
            Left            =   5760
            TabIndex        =   6
            Top             =   90
            Width           =   480
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   3525
      Left            =   30
      TabIndex        =   10
      Top             =   1215
      Width           =   10425
      _cx             =   18389
      _cy             =   6218
      Appearance      =   1
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
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
      Rows            =   11
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCHRecipe.frx":0EB0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Frame fraSplit 
      Height          =   705
      Left            =   45
      TabIndex        =   28
      Top             =   6120
      Width           =   10410
      Begin VB.TextBox txtӦ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   6915
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1395
      End
      Begin VB.TextBox txtʵ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   8955
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1395
      End
      Begin VB.ComboBox cbo�巨 
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lblӦ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ��"
         Height          =   240
         Left            =   6360
         TabIndex        =   14
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblʵ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʵ��"
         Height          =   240
         Left            =   8400
         TabIndex        =   16
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lbl�巨 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�巨"
         Height          =   240
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.Frame fra��� 
      Height          =   1380
      Left            =   45
      TabIndex        =   29
      Top             =   6765
      Width           =   10410
      Begin VSFlex8Ctl.VSFlexGrid vsSpecShow 
         Height          =   1095
         Left            =   60
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   195
         Width           =   10275
         _cx             =   18124
         _cy             =   1931
         Appearance      =   0
         BorderStyle     =   0
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
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483644
         BackColorAlternate=   -2147483644
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483644
         FloodColor      =   192
         SheetBorder     =   -2147483644
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCHRecipe.frx":0F0B
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
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   8430
      Left            =   10515
      TabIndex        =   31
      Top             =   -150
      Width           =   45
   End
End
Attribute VB_Name = "frmCHRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ�����
Private mbytFun As Byte  '0-����,1-����
Private mstrPrivs As String
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mlng����ID As Long
Private mint������Դ As Integer '�ӵ��÷�����
Private mlng���˿���ID As Long '���˿���
Private mlng��������ID As Long '��������ID
Private mlng��ҩ�� As Long '��ҩ��ID
Private mobjDetails As BillDetails
Private mstr�ѱ� As String
Private mint���� As Integer '�����ҽ�����ˣ���Ϊ��������
Private mbln�Ӱ� As Boolean
Private mcolStock As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mrsPati As ADODB.Recordset
Private mrsWarn As ADODB.Recordset
Private mstrWarn As String
Private mblnFirst As Boolean

Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean

Private mcurModiMoney As Currency
Private mcur����ҩ��� As Currency      '�����䷽֮ǰ,���ݵĽ��,��������ʱ���㵱ǰ���ݽ��

Public mstr�巨 As String   'out
Private mcll���  As Collection  '��Ʒ��IDΪ����������:���1,����;���2,����|δ��������
Private mcllInput����ժҪ As Collection  '��ҩƷIDΪ������¼����¼�����ժҪ
Private mint��ҩ��̬ As Integer
Private Const mlngModul = 1150
Private Const MIPTS = 4 '�䷽������
Private Const MCOLS = 3 'ÿһ������
Private Const MROWS = 12 '����ɼ�����
Private Const STR_HEAD = "�в�ҩ,1280,1;����,700,7;,400,1"
Private Enum COL_BILL
    col��ҩ = 0
    col���� = 1
    col��λ = 2
End Enum
'--�Զ������񼰷���ҩ��̬(ɢװ,��Ƭ,����)��������:31867
Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal bytFun As Byte, ByVal curModiMoney As Currency, _
    ByVal lng����ID As Long, ByVal int������Դ, ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, ByVal lng��ҩ�� As Long, _
    ByVal objDetails As BillDetails, ByVal str�ѱ� As String, _
    ByVal int���� As Integer, ByVal bln�Ӱ� As Boolean, ByVal str�巨 As String, rsWarn As ADODB.Recordset, colStock As Collection, _
    Optional int��ҩ��̬ As Integer = -1) As BillDetails
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ҩ�䷽�༭����(�������)
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-02-02 14:37:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    mbytFun = bytFun  '0-����,1-����  ��ʱδ�õ��˱���,Ϊ����Ԥ��
    mstrPrivs = strPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    
    mcurModiMoney = curModiMoney
    mlng����ID = lng����ID
    mint������Դ = int������Դ
    mlng���˿���ID = lng���˿���ID
    mlng��������ID = lng��������ID
    mlng��ҩ�� = lng��ҩ��
    mstr�ѱ� = str�ѱ�
    mint���� = int����
    mbln�Ӱ� = bln�Ӱ�
    mstr�巨 = str�巨
    Set mrsWarn = rsWarn
    Set mcolStock = colStock
    mint��ҩ��̬ = int��ҩ��̬

    mcur����ҩ��� = 0
    
    '���봫��ĵ�����ϸ���ݵ��в�ҩ��
    Set mobjDetails = New BillDetails
    For i = 1 To objDetails.Count
        With objDetails(i)
            If .�շ���� = "7" Then
                 Call mobjDetails.Add(.Detail, .�շ�ϸĿID, .���, .��������, .����ID, .��ҳID, .����ID, .����ID, _
                 .����, .�Ա�, .����, .סԺ��, .����, .�ѱ�, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, _
                 .���ӱ�־, .ִ�в���ID, .InComes, .���￨��, "", .������, .ҽ�Ƹ���, .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ)
            Else
                For j = 1 To .InComes.Count
                    mcur����ҩ��� = mcur����ҩ��� + .InComes(j).ʵ�ս��
                Next
            End If
        End With
    Next
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then
        Set ShowMe = mobjDetails
    End If
End Function

Private Sub cbo�巨_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboҩ��_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    If Visible = False Then Exit Sub
    If cboҩ��.ListIndex < 0 Then Exit Sub
    
    
    If Val(cboҩ��.Tag) <> cboҩ��.ItemData(cboҩ��.ListIndex) Then
        Call ����ˢ��������ҩ���
         cboҩ��.Tag = cboҩ��.ItemData(cboҩ��.ListIndex)
        Call ReCalcӦ�պϼ�
        mblnChange = True
        Call ShowSpecs(Val(vsBill.Cell(flexcpData, vsBill.Row, (vsBill.Col \ MCOLS) * MCOLS + 2)))
    End If
End Sub

Private Sub cboҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chkAuto_Click()
    txtAuto.Enabled = chkAuto.Value = 1
    If txtAuto.Enabled And Visible Then txtAuto.SetFocus
End Sub

Private Sub chkAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean, i As Long, j As Long, strStock As String
    
    If mobjDetails.Count = 0 Then
        MsgBox "�����䷽����������һζ�в�ҩ��", vbInformation, gstrSysName
        vsBill.Row = vsBill.FixedRows
        vsBill.Col = vsBill.FixedCols
        vsBill.SetFocus: Exit Sub
    End If
    If cboҩ��.Visible And cboҩ��.ListIndex = -1 Then
        MsgBox "��ȷ����ҩ�䷽�ķ�ҩҩ����", vbInformation, gstrSysName
        cboҩ��.SetFocus: Exit Sub
    End If
    
     '��¼��ѡ��ҩ�巨
    mstr�巨 = Mid(cbo�巨.Text, InStr(1, cbo�巨.Text, "-") + 1)
    
    'ǿ��ʹ���븶����Ч
    If Me.ActiveControl Is txt���� Then
        Call txt����_Validate(blnCancel)
        If blnCancel Then Exit Sub
    End If
    
    '�����:������cboҩ����Click�м��
    Dim lngҩ��ID As Long
    For i = 1 To mobjDetails.Count
        With mobjDetails(i)
            If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
                strStock = FormatEx(.Detail.���, 5) & IIf(gblnסԺ��λ, .Detail.סԺ��λ, .���㵥λ)
            End If
            lngҩ��ID = mobjDetails(i).Detail.ҩ��ID
            If .���� * .���� > .Detail.��� Then
                If Not gbln���뷢ҩ Then
                    If .Detail.���� Or .Detail.��� Then
                        MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ����ǰ���" & strStock & ",��������������", vbInformation, gstrSysName
                        Exit For
                    ElseIf mcolStock("_" & .ִ�в���ID) <> 0 Then
                        If mcolStock("_" & .ִ�в���ID) = 1 Then
                            If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���" & strStock & ",��������������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit For
                            End If
                        ElseIf mcolStock("_" & .ִ�в���ID) = 2 Then
                            MsgBox """" & .Detail.���� & """�ĵ�ǰ���" & strStock & ",��������������", vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                ElseIf gblnStock And gstr��ҩ�� <> "" Then
                    MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���" & strStock & ",��������������", vbInformation, gstrSysName
                    Exit For
                End If
            End If
        End With
    Next
    
    '������,��Ҫ�ƶ���������
    If i <= mobjDetails.Count Then
        With vsBill
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step MCOLS
                    If lngҩ��ID = Val(.Cell(flexcpData, i, j + 2)) Then
                        .Row = i: .Col = j
                        If vsBill.Editable And vsBill.Visible Then vsBill.SetFocus
                        vsBill.ShowCell .Row, .Col
                        Exit Sub
                    End If
                Next
            Next
        End With
        Exit Sub
    End If
    
    '���¼�������,��������Ĺ��̾�������ɾ��,���,��ϸ���ݿ��ܲ�һ��,���,��Ҫ��������
    Dim ObjBillDetails As BillDetails
    Set ObjBillDetails = New BillDetails
    Dim q As Integer, intRow As Integer
    
    With vsBill
        intRow = 1
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1 Step MCOLS
                lngҩ��ID = Val(.Cell(flexcpData, i, j + 2))
                If opt��̬(0).Value = False And lngҩ��ID <> 0 Then
                
                    '��ɢװ,��Ҫ����Ƿ��̯���
                    If InStr(1, mcll���("_" & lngҩ��ID), "|") > 0 Or mcll���("_" & lngҩ��ID) = "" Then
                            ShowMsgbox "ҩ��Ϊ" & .TextMatrix(i, j) & "�Ĳ�ҩδ�������,���ܼ���!"
                            .Row = i: .Col = j
                            If vsBill.Enabled Then vsBill.SetFocus
                            Exit Sub
                    End If
                End If
                If lngҩ��ID <> 0 Then
                For q = 1 To mobjDetails.Count
                    If lngҩ��ID = mobjDetails(q).Detail.ҩ��ID Then
                        '���¸�ֵ
                        With mobjDetails(q)
                            ObjBillDetails.Add .Detail, .�շ�ϸĿID, intRow, .��������, .����ID, .��ҳID, .����ID, .����ID, .����, .�Ա�, .����, _
                                .סԺ��, .����, .�ѱ�, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, _
                                .���￨��, .Key, .������, .ҽ�Ƹ���, .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .ԭʼ����, .ԭʼִ�в���ID, .Ӥ����
                        End With
                        intRow = intRow + 1
                    End If
                Next
            End If
            Next
        Next
    End With
    Set mobjDetails = ObjBillDetails
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd��̬_Click()
    '���ܣ���ζҩ����ɢװ���䲻��ʱ������ɢװ���
    Dim lngҩ��ID As Long, dbl���� As Double, lngҩƷID As Long
    Dim objBillDetail As BillDetail, objDetail As Detail
    With vsBill
        lngҩ��ID = Val(.Cell(flexcpData, .Row, (.Col \ MCOLS) * MCOLS + 2))
        dbl���� = Val(.TextMatrix(.Row, (.Col \ MCOLS) * MCOLS + 1))
        lngҩƷID = Val(cmd��̬.Tag)    'ȱʡ���
        If zlGetDetail(lngҩƷID, dbl����, objDetail) = False Then
            Exit Sub
        End If
        dbl���� = FormatEx(dbl���� / IIf(objDetail.����ϵ�� = 0, 1, objDetail.����ϵ��), 5)
        If CheckStock(lngҩƷID, dbl����, objDetail) = False Then
            Exit Sub
        End If
        
        'ɾ��ҩ��Ϊ
        Call DeleteDetails(lngҩ��ID)
        Call mcll���.Remove("_" & lngҩ��ID)
        mcll���.Add lngҩƷID & "," & dbl����, "_" & lngҩ��ID
        
         '������ϸ
         If SetBillDetail(lngҩƷID, dbl����, 1, Nothing, objBillDetail) = False Then
            '�ֽ�ʧ��
            Call DeleteDetails(lngҩ��ID)
             Call ReCalcӦ�պϼ�
         Else
            '�����շ���Ŀ����
            Call zlCalcMoney(objBillDetail, True)
         End If
        Call ReCalcӦ�պϼ�
        Call Show��ҩ���(lngҩ��ID, dbl����, 0)
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mobjDetails Is Nothing Then Exit Sub
    If mobjDetails.Count <> 0 Then
        vsBill.SetFocus
    Else
        If opt��̬(0).Enabled Then opt��̬(0).SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    mblnFirst = True
    mstrWarn = ""
    
    mblnOK = False
    mblnChange = False
                        
                       
    chkAuto.Value = IIf(zlDatabase.GetPara("��ҩ�Զ�����", glngSys, mlngModul) = "1", 1, 0)
    txtAuto.Text = Val(zlDatabase.GetPara("��ҩ�Զ����볤��", glngSys, mlngModul, 5))
            
    
    '��ʼ������
    If Not InitData Then Unload Me: Exit Sub
                        
    '��ʾ��������
    Call ShowDetails
    Call vsBill_GotFocus
End Sub

Private Function InitData() As Boolean
'���ܣ���ʼ����Ӧ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, cur������� As Currency
    
    On Error GoTo errH
    '��ȡ������Ϣ,������ö����������mbytFun = 2 And
    If mlng����ID <> 0 Then
        Set rsTmp = GetMoneyInfo(mlng����ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
    
        If Not rsTmp Is Nothing And Not mrsWarn Is Nothing Then
            cur������� = Val("" & rsTmp!Ԥ�����) - Val("" & rsTmp!�������)
            If gbln�����������۷��� Then cur������� = cur������� - GetPriceMoneyTotal(1, mlng����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)   '�޸Ļ��۵��ұ���Ҫ�㻮�۵�ʱ,�ӵ�ǰ���ݽ��
            
            strSQL = "," & Val("" & rsTmp!Ԥ�����) & " as Ԥ�����," & (Val("" & rsTmp!Ԥ�����) - cur�������) & " as �������," & cur������� & " as �������"
        Else
            strSQL = ",0 as Ԥ�����,0 as �������,0 as �������"
        End If
        '76451,Ƚ����,2014-8-19
        strSQL = "Select A.����,A.סԺ��,A.��ǰ���� As ����,A.����ID,A.��ҳID ��ҳId,Nvl(A.��ǰ����ID,0) as ����ID,Zl_Patiwarnscheme(A.����id, A.��ҳID) As ���ò���," & _
            " Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,A.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�" & _
            strSQL & _
            " From ������Ϣ A Where A.����ID=[1]"
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    End If
    
    '��ȡ��ҩ��
    If gbln���뷢ҩ Then
        lblҩ��.Visible = False
        cboҩ��.Visible = False
    Else
        Set rsTmp = GetDepartments("��ҩ��", mint������Դ & ",3")
        For i = 1 To rsTmp.RecordCount
            cboҩ��.AddItem IIf(zlIsShowDeptCode, rsTmp!���� & "-", "") & rsTmp!����
            cboҩ��.ItemData(cboҩ��.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng��ҩ�� Then cboҩ��.ListIndex = cboҩ��.NewIndex
            rsTmp.MoveNext
        Next
    End If
    
     '��ȡ��ҩ�巨
    strSQL = "select ID,rownum||'-'||���� as ���� from ������ĿĿ¼ where ���='E' and ��������='3' order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo�巨.Clear
    cbo�巨.AddItem ""
    
    Do While Not rsTmp.EOF
        cbo�巨.AddItem rsTmp!����
        rsTmp.MoveNext
    Loop
    
    If mstr�巨 <> "" Then  '����δ����֮ǰ���½���
        For i = 0 To cbo�巨.ListCount
            If Mid(cbo�巨.List(i), InStr(1, cbo�巨.List(i), "-") + 1) = mstr�巨 Then
                cbo�巨.ListIndex = i
                Exit For
            End If
        Next
        If i > cbo�巨.ListCount Then
            cbo�巨.AddItem mstr�巨
            cbo�巨.ListIndex = cbo�巨.NewIndex
        End If
    Else
        If cbo�巨.ListCount = 0 Then cbo�巨.Enabled = False
        'Ĭ��Ϊ��ѡ�巨
    End If
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowDetails()
    '���ܣ�ȫ��ˢ����ʾ��ǰ�䷽����
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim lngRow As Long, lngCol As Long
    Dim i As Long, j As Long, k As Long, intIndex As Integer
    Dim strҩ��ID As String
    Dim varData As Variant, dbl���� As Double
    Dim str������� As String
    Dim lngҩ��ID As Long, dblTemp As Double
    
    Set mcll��� = New Collection
    Set mcllInput����ժҪ = New Collection
    Call InitFace
    
    strҩ��ID = ""
    For i = 1 To mobjDetails.Count
        If InStr(1, strҩ��ID & ",", "," & mobjDetails(i).Detail.ҩ��ID & ",") = 0 Then
            strҩ��ID = strҩ��ID & "," & mobjDetails(i).Detail.ҩ��ID
        End If
        If i = 1 Then
            txt����.Text = mobjDetails(i).����
            intIndex = Decode(mobjDetails(i).Detail.��ҩ��̬, 0, 0, 1, 1, 2)
            opt��̬(intIndex).Value = True
        End If
        If mobjDetails(i).ժҪ <> "" Then
            mcllInput����ժҪ.Add mobjDetails(i).ժҪ, "K" & mobjDetails(i).Detail.ID
        End If
        
        '�ۼƷ���
        For j = 1 To mobjDetails(i).InComes.Count
            curӦ�� = curӦ�� + mobjDetails(i).InComes(j).Ӧ�ս��
            curʵ�� = curʵ�� + mobjDetails(i).InComes(j).ʵ�ս��
        Next
    Next
    
    varData = Split(strҩ��ID, ",")
    Dim str�������� As String, str���㵥λ As String, dbl���� As Double
    With vsBill
        .Redraw = flexRDNone
        For i = 1 To UBound(varData)
            lngRow = ((i - 1) \ MIPTS) + 1
            lngCol = ((i - 1) Mod MIPTS) * MCOLS
            If i = 1 Then lngҩ��ID = Val(varData(i))
            If lngRow > .Rows - 1 Then
                .AddItem ""
                Call SetSplitLine
            End If
            'ҩƷid,����;...|ʣ������
            dbl���� = 0: str������� = ""
            For j = 1 To mobjDetails.Count
                If Val(varData(i)) = mobjDetails(j).Detail.ҩ��ID Then
                    dblTemp = mobjDetails(j).���� * mobjDetails(j).Detail.����ϵ��
                    If gblnסԺ��λ Then    '52722
                        dblTemp = dblTemp * mobjDetails(j).Detail.סԺ��װ
                    End If
                    dbl���� = dbl���� + dblTemp
                    str�������� = mobjDetails(j).Detail.��������
                    str���㵥λ = mobjDetails(j).Detail.������λ
                    str������� = str������� & ";" & mobjDetails(j).Detail.ID & "," & dblTemp
                End If
            Next
            
            If str������� <> "" Then str������� = Mid(str�������, 2)
            mcll���.Add str�������, "_" & Val(varData(i))
            .TextMatrix(lngRow, lngCol) = str��������
            .TextMatrix(lngRow, lngCol + 1) = FormatEx(dbl����, 5)
            .TextMatrix(lngRow, lngCol + 2) = str���㵥λ
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            .Cell(flexcpData, lngRow, lngCol + 1) = .TextMatrix(lngRow, lngCol + 1)
            .Cell(flexcpData, lngRow, lngCol + 2) = Val(varData(i))
            If i = UBound(varData) Then
                '��λ�����һ�еļ�������
                .Row = lngRow
                If dbl���� <> 0 Then
                    If lngCol + MCOLS > .Cols - 1 Then
                        If .Rows - 1 > .Row Then
                            .Row = .Row + 1
                        Else
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                        End If
                        .Col = .FixedCols
                    Else
                        .Col = lngCol + MCOLS
                    End If
                Else
                         .Col = lngCol + 1
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    txtӦ��.Text = Format(curӦ��, gstrDec)
    txtʵ��.Text = Format(curʵ��, gstrDec)
    If mobjDetails.Count > 0 Then
        If Not gbln���뷢ҩ Then
            cboҩ��.ListIndex = cbo.FindIndex(cboҩ��, mobjDetails(1).ִ�в���ID)
        End If
        If cboҩ��.ListIndex < 0 And cboҩ��.ListCount > 0 Then cboҩ��.ListIndex = 0
        Show��ҩ��� lngҩ��ID, Val(vsBill.TextMatrix(vsBill.Row, GetBillCol(1, vsBill.Col)))
    End If

End Sub

Private Sub InitFace()
'���ܣ���ʼ����ҩ�䷽����ʽ������
'������mstrExtData=����ÿζ��ҩ��Ϣ���巨��Ϣ�Ĵ�,Ϊ��ʱ��ʾ��������ҩ�䷽
    Dim arrCols As Variant
    Dim blnPre As Boolean, i As Integer
    
    arrCols = Split(STR_HEAD, ";")
    
    With vsBill
        blnPre = .Redraw
        .Redraw = flexRDNone
        .Rows = 0: .Cols = 0
        .Rows = MROWS: .Cols = (UBound(arrCols) + 1) * MIPTS
        .FixedCols = 0: .FixedRows = 1
        .RowHidden(0) = True
        
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(arrCols(i Mod 3), ",")(0)
            .ColWidth(i) = Split(arrCols(i Mod 3), ",")(1)
            .ColAlignment(i) = Split(arrCols(i Mod 3), ",")(2)
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .GridColor = .BackColor
        .GridColorFixed = .BackColorFixed
        
        .Editable = flexEDKbdMouse
        .Row = .FixedRows: .Col = .FixedCols
                
        Call SetSplitLine
        
        .Redraw = blnPre
    End With
    
    txtӦ��.Text = gstrDec
    txtʵ��.Text = gstrDec
    txt����.Text = 1: txt����.Tag = 1

End Sub

Private Sub SetSplitLine()
'���ܣ�������ҩ�䷽��������зָ���
    Dim lngRow As Long, lngCol As Long
    Dim blnPre As Boolean, i As Long
    
    With vsBill
        blnPre = .Redraw
        lngRow = .Row: lngCol = .Col
        
        .Redraw = flexRDNone
        For i = 0 To .Cols - 1 Step MCOLS
            .Select .FixedRows, i + MCOLS - 1, .Rows - 1, i + MCOLS - 1
            .CellBorder &H808080, 0, 0, 1, 0, 0, 0
        Next
        
        .Row = lngRow: .Col = lngCol
        .Redraw = blnPre
    End With
End Sub

Private Function GetRow(ByVal lngRow As Long, ByVal lngCol As Long) As Long
'���ܣ���ȡ��ǰ��Ԫ��Ӧ�ķ����к�
    GetRow = (lngRow - 1) * MIPTS + lngCol \ MCOLS + 1
End Function

Private Sub opt��̬_Click(Index As Integer)
    Dim lngҩ��ID As Long, lngҩ��ID As Long, str������� As String, lngҩƷID As Long
    Dim rsTemp As ADODB.Recordset, lngTemp As Long
    
    If Not Me.Visible Then Exit Sub
    With vsBill
        lngҩ��ID = Val(.Cell(flexcpData, .FixedRows, .FixedCols + 2))
        If gblnStock = False Or lngҩ��ID = 0 Then
            Call ����ˢ��������ҩ���
            Exit Sub  '���޶����ʱ,�˳�
        End If
        
        '�޶����ʱ,��һζԼ��ȱʡ�����ܱ仯,�仯��,����ҩ���ͱ���
         str������� = mcll���("_" & lngҩ��ID)
         lngҩ��ID = mlng��ҩ��
         If cboҩ��.ListIndex >= 0 Then lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
         If str������� <> "" Then
                'ȷ��ҩƷID
                Set rsTemp = Get��ҩ���(lngҩ��ID, Index)
                If rsTemp.RecordCount > 0 Then
                    lngҩƷID = Val(Split(str�������, ",")(0))
                    If lngҩƷID <> Val(rsTemp!ҩƷID) Then
                        If mlng����ID <> 0 Then
                            lngTemp = Get�շ�ִ�п���ID("7", Val(NVL(rsTemp!ҩƷID)), NVL(rsTemp!ִ�п���, 0), mlng���˿���ID, mlng��������ID, mint������Դ, mlng��ҩ��, mrsPati!����ID)
                        Else
                            lngTemp = Get�շ�ִ�п���ID("7", Val(NVL(rsTemp!ҩƷID)), NVL(rsTemp!ִ�п���, 0), mlng���˿���ID, mlng��������ID, mint������Դ, mlng��ҩ��)
                        End If
                       '������ҩ��
                        If Not gbln���뷢ҩ Then
                            If lngTemp <> lngҩ��ID And lngTemp <> 0 Then
                                cboҩ��.ListIndex = cbo.FindIndex(cboҩ��, lngTemp)
                            End If
                        End If
                    End If
                End If
         End If
    End With
    '��̬���ˣ�Ҫ���·����������
    Call ����ˢ��������ҩ���
End Sub

Private Sub opt��̬_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAuto_GotFocus()
    Call zlControl.TxtSelAll(txtAuto)
End Sub

Private Sub txtAuto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAuto_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAuto.Text) Then txtAuto.Text = 5
    If Val(txtAuto.Text) > 20 Then txtAuto.Text = 20
    If Val(txtAuto.Text) < 2 Then txtAuto.Text = 2
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txt����_Validate(Cancel As Boolean)
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim i As Integer, strStock As String
    '�������
    If Not IsNumeric(txt����.Text) Then
        MsgBox "������һ����Ч����ֵ��", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    If Val(txt����.Text) <> Int(txt����.Text) Then
        MsgBox "��ҩ����Ӧ����������ֵ��", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    
    If Val(txt����.Text) = 0 Then
        MsgBox "������һ������ĸ�����", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt����)
        Cancel = True: Exit Sub
    End If
    If Val(txt����.Tag) = Val(txt����.Text) Then Exit Sub
    
    If Get��ҩ��̬ = 0 Then
        'ɢװ��̬��,��Ҫ�����
        For i = 1 To mobjDetails.Count
            If CheckStock(mobjDetails(i).�շ�ϸĿID, mobjDetails(i).����, mobjDetails(i).Detail) = False Then
                Cancel = True: Exit Sub
            End If
        Next
        'Ϊ������ʱ�����ø���
        For i = 1 To mobjDetails.Count
            mobjDetails(i).���� = Val(txt����.Text)
            Call zlCalcMoney(mobjDetails(i), True)
        Next
        '����Ӧ�պϼ�
        Call ReCalcӦ�պϼ�
        txt����.Tag = Val(txt����.Text)
        Exit Sub
    End If
    '��ɢװ��̬��,��Ҫ����ˢ�¹��
    Call ����ˢ��������ҩ���
    txt����.Tag = Val(txt����.Text)
End Sub

Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cur���� As Currency, cur��� As Currency
    Dim i As Long, strStock As String
    
    If NewRow <= 0 Or NewCol = -1 Then Exit Sub
    
    Call vsBill.ShowCell(NewRow, vsBill.LeftCol)
     
     If OldRow <> NewRow Or (OldCol \ MCOLS) <> (NewCol \ MCOLS) Then   '���л򻻵���һҩƷ��
        If vsBill.Cell(flexcpData, NewRow, (NewCol \ MCOLS) * MCOLS + 2) <> 0 Then
            Call Show��ҩ���(Val(vsBill.Cell(flexcpData, NewRow, (NewCol \ MCOLS) * MCOLS + 2)), Val(vsBill.TextMatrix(NewRow, (NewCol \ MCOLS) * MCOLS + 1)))
        Else
            vs��ҩ���.Rows = vs��ҩ���.FixedRows
            cmd��̬.Visible = False
        End If
        Call ShowSpecs(Val(vsBill.Cell(flexcpData, NewRow, (NewCol \ MCOLS) * MCOLS + 2)))
    End If
End Sub

Private Sub vsBill_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    '��λ����겻�ɽ���
    If Button = 1 And (vsBill.MouseCol Mod MCOLS) = col��λ Then Cancel = True
End Sub

Private Sub vsBill_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    '��λ�а������ɽ���
    If Not Visible Or vsBill.Redraw = flexRDNone Then Exit Sub
    If (NewCol Mod MCOLS) = col��λ Then
        Cancel = True
        If OldCol > NewCol Then '�����ƶ�ʱ����
            vsBill.Col = NewCol - 1
        Else
            If NewCol + 1 <= vsBill.Cols - 1 Then
                vsBill.Col = NewCol + 1
            Else
                vsBill.Col = NewCol - 1
            End If
        End If
        vsBill.Row = NewRow
    End If
End Sub

Private Sub vsBill_GotFocus()
    With vsBill
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .BackColorSel = vbBlue
    End With
End Sub

Private Sub vsBill_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim i As Long, j As Long, k As Long
    Dim lngҩ��ID As Long
    
    If KeyCode = vbKeyDelete Then
        With vsBill
            If .TextMatrix(.Row, (.Col \ MCOLS) * MCOLS) <> "" Then
                If MsgBox("Ҫɾ��""" & .TextMatrix(.Row, (.Col \ MCOLS) * MCOLS) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                lngҩ��ID = Val(.Cell(flexcpData, .Row, GetBillCol(2, .Col)))
                'ɾ����ϸ����
                Call DeleteDetails(lngҩ��ID)
                '�Ƴ����
                mcll���.Remove "_" & lngҩ��ID
                '����۸�
                Call ReCalcӦ�պϼ�
                '�����ǰζҩ��Ϣ
                For i = 0 To MCOLS - 1
                    .TextMatrix(.Row, (.Col \ MCOLS) * MCOLS + i) = ""
                    .Cell(flexcpData, .Row, (.Col \ MCOLS) * MCOLS + i) = Empty
                Next
                
                '�����������ǰ��
                For i = .Row To .Rows - 1
                    For j = 0 To .Cols - 1 Step MCOLS
                        If Not (i = .Row And j <= (.Col \ MCOLS) * MCOLS) Then
                            For k = 0 To MCOLS - 1
                                If j = 0 Then
                                    .TextMatrix(i - 1, .Cols - (MCOLS - k)) = .TextMatrix(i, j + k)
                                    .Cell(flexcpData, i - 1, .Cols - (MCOLS - k)) = .Cell(flexcpData, i, j + k)
                                Else
                                    .TextMatrix(i, j + k - MCOLS) = .TextMatrix(i, j + k)
                                    .Cell(flexcpData, i, j + k - MCOLS) = .Cell(flexcpData, i, j + k)
                                End If
                                .TextMatrix(i, j + k) = ""
                                .Cell(flexcpData, i, j + k) = Empty
                            Next
                        End If
                    Next
                Next
                'ɾ������Ŀ���
                If .Rows > MROWS Then
                    For i = .Rows - 1 To MROWS Step -1
                        If .TextMatrix(i, 0) = "" Then
                            .RemoveItem i
                        End If
                    Next
                End If
                Call .ShowCell(.Row, .Col)
                sta.Panels(3).Text = "��" & mcll���.Count & "ζҩ"
            End If
        End With
    End If
End Sub

Private Sub vsBill_KeyPress(KeyAscii As Integer)
'���ܣ��Ǳ༭״̬ʱ���Զ��ƶ���Ԫ��
    If KeyAscii = 13 Then
        KeyAscii = 0
        '��λ����һӦ���뵥Ԫ��
        If vsBill.TextMatrix(vsBill.Row, (vsBill.Col \ MCOLS) * MCOLS) = "" Then
            If GetRow(vsBill.Row, vsBill.Col) > 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        Else
            Call EnterNextCell(vsBill.Row, vsBill.Col)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If CellCanEdit(vsBill.Row, vsBill.Col) Then
            If vsBill.Col <> (vsBill.Col \ MCOLS) * MCOLS Then
                Exit Sub
            End If
            If SelectChineDrug("") = False Then Exit Sub
        End If
    End If
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    If Not mblnReturn Then '�ǻس�ȷ��ʧЧ
        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
    End If
End Sub

Private Sub vsBill_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col Mod MCOLS = col��ҩ And chkAuto.Value = 1 Then
        '�Զ���ɿ����ҩ����
        If Len(vsBill.EditText) >= Val(txtAuto.Text) Then
            Call vsBill_KeyPressEdit(Row, Col, 13)
        End If
    ElseIf Col Mod MCOLS = col���� Then
        '�Զ���ɿ����������
        If InStr(gstrABC, UCase(Chr(KeyCode))) > 0 And Between(KeyCode, vbKeyA, vbKeyZ) Then
            vsBill.EditCell
            vsBill.EditText = UCase(Chr(KeyCode))
            Call vsBill_KeyPressEdit(Row, Col, 13)
            vsBill.FinishEditing False  '�ؼ�bugδ��Ч
        End If
    End If
End Sub

Private Sub vsBill_LostFocus()
  With vsBill
        .FocusRect = flexFocusLight
        .HighLight = flexHighlightAlways
        .BackColorSel = &HE7CFBA
    End With
End Sub

Private Sub vsBill_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsBill.EditSelStart = 0
    vsBill.EditSelLength = zlCommFun.ActualLen(vsBill.EditText)
End Sub

Private Sub vsBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�����ĳЩ�в�����༭(���¼�����BeforeEdit,��EditText��ֵ֮ǰ)
    mblnReturn = False

    '������������
    If Not CellCanEdit(Row, Col) Then Cancel = True

    If Col Mod MCOLS = col���� Then
        vsBill.EditMaxLength = 8
    Else
        vsBill.EditMaxLength = 0
    End If
End Sub

Private Function CellCanEdit(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ�������ҩ�䷽ʱ,�ж�ָ���ĵ�Ԫ��ǰ�Ƿ���������
'˵�������䷽��������,���ǰһ��δ����,��ǰ����������
    '��λ����һ����ҩ���뵥Ԫ
    lngCol = (lngCol \ MCOLS) * MCOLS
    If lngCol - MCOLS >= vsBill.FixedCols Then
        lngCol = lngCol - MCOLS
    Else
        If lngRow - 1 >= vsBill.FixedRows Then
            lngRow = lngRow - 1
            lngCol = vsBill.Cols - MCOLS
        Else
            CellCanEdit = True
            Exit Function
        End If
    End If
    CellCanEdit = vsBill.TextMatrix(lngRow, lngCol) <> ""
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ�������һ����ҩ�䷽�����뵥Ԫ��
    '��ǰλ��δ������ҩ
    If vsBill.TextMatrix(lngRow, (lngCol \ MCOLS) * MCOLS) = "" Then Exit Sub

    '����δ����
    If lngCol Mod MCOLS = 1 And vsBill.TextMatrix(lngRow, lngCol) = "" Then Exit Sub

    If lngCol + 1 <= vsBill.Cols - 1 Then
        If (lngCol + 1) Mod MCOLS = col��λ And lngCol \ MCOLS + 1 = MIPTS Then
            If lngRow + 1 > vsBill.Rows - 1 Then
                vsBill.AddItem "", vsBill.Rows
                Call SetSplitLine
            End If
            lngCol = 0
            lngRow = lngRow + 1
        Else
            lngCol = lngCol + 1
        End If
    Else
        If lngRow + 1 > vsBill.Rows - 1 Then
            vsBill.AddItem "", vsBill.Rows
            Call SetSplitLine
        End If
        lngRow = lngRow + 1
        lngCol = vsBill.FixedCols
    End If

    vsBill.Row = lngRow: vsBill.Col = lngCol
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange And Not mblnOK Then
        If MsgBox("�䷽�����ѱ��ı䣬ȷʵҪ������Щ�ı��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
        
    Set mcolStock = Nothing
    Set mrsWarn = Nothing
    Set mrsPati = Nothing
    
    zlDatabase.SetPara "��ҩ�Զ�����", chkAuto.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ҩ�Զ����볤��", Val(txtAuto.Text), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function SelectChineDrug(ByVal strInput As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ����ҩ
    '��Σ�strInput-Ҫ���ҵ�ֵ
    '���Σ�
    '���أ��ɹ�,����true, ���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-27 13:56:04
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    'gblnStock:��ʾ�շѻ����ʱ���ָ���˷���ҩ�����Ƿ�ֻ�������ҩ���п���ҩƷ
    Dim vPoint As POINTAPI, strTmp As String
    Dim rsTemp As ADODB.Recordset, str���� As String
    Dim lngҩ��ID As Long, strStock As String, strSQLAdd As String, str��׼��Ŀ As String, strSQL As String
    Dim strSQLInput As String, str����ʱ�� As String, strWhere As String
    Dim int��ҩ��̬ As Integer, blnCancel As Boolean
    Dim str��� As String, str���� As String, lngҩƷID As Long, lngTmp As Long
    Dim lngҩ��ID As Long, lng�ϴ�ҩ��ID As Long
    
    int��ҩ��̬ = Get��ҩ��̬
    '������ö��������mbytFun = 0 And,�ſ��˻���
    If cboҩ��.ListIndex < 0 Then cboҩ��.ListIndex = cbo.FindIndex(cboҩ��, lngҩ��ID)
    If cboҩ��.ListIndex < 0 And cboҩ��.ListCount > 0 Then cboҩ��.ListIndex = 0
    If cboҩ��.ListIndex < 0 Then
        MsgBox "ҩ��δѡ��,��ѡ��ҩ��?", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
         If cboҩ��.Enabled And cboҩ��.Visible Then cboҩ��.SetFocus
        Exit Function
    End If
    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
  '����ҩƷȨ��
    str���� = ""
    If InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then str���� = str���� & " And E.�������<>'����ҩ'"
    If InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then str���� = str���� & " And E.�������<>'����ҩ'"
    If InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then str���� = str���� & " And E.��ֵ���� Not IN('����','����')"
    '��δ�д�����
   ' If InStr(mstrPrivsOpt, "����ҩƷ����") = 0 Then str���� = str���� & " And E.��ֵ���� Not IN('����I��','����II��')"
    If int��ҩ��̬ = 0 And lngҩ��ID <> 0 Then
        'ֻ��ɢװ���п��
        strStock = _
            " Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
            "  Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
            "           And ���� = 1 And �ⷿID=[4]" & _
            "  Group by ҩƷID  " & _
            "  Having Sum(Nvl(��������,0))<>0"
     Else
        strStock = "Select NULL as ҩƷID,NULL as ��� From Dual"
     End If
    
    If int��ҩ��̬ = 0 Then
        str��� = _
        "   And Nvl(C.��ҩ��̬,0) = [6] And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([7],3)" & _
        "   And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & _
         IIf(gblnStock And lngҩ��ID <> 0, " And nvl(X.���,0)<>0", "")
    Else
         str��� = " And Exists(Select 1 From ҩƷ��� C Where C.ҩ��ID=E.ҩ��ID And Nvl(C.��ҩ��̬,0) = [6])"
    End If
    
    
    str����ʱ�� = "" & _
        "   And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " & _
        "   And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        "   And A.������� IN([7],3)"
    
    str��׼��Ŀ = ""
    If int��ҩ��̬ = 0 Then
        If mint���� <> 0 And mlng����ID <> 0 Then
            '���˺�:24862
            If zl_Check��׼��Ŀ(gclsInsure, mint����, mlng����ID, False) Then str��׼��Ŀ = Get������׼��Ŀ(mlng����ID, "D.ID")
        End If
    End If
        
    If strInput <> "" Then
            strWhere = " And (A.���� Like [1] And B.����=[3] Or B.���� Like [2] And B.����=[3] Or B.���� Like upper([2]) And B.���� IN([3],3))"
            If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                If Mid(gstrMatchMode, 1, 1) = "1" Then strWhere = " And (A.���� Like [1] And B.����=[3] Or B.���� Like Upper([2]) And B.����=3)"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                If Mid(gstrMatchMode, 2, 1) = "1" Then strWhere = " And B.���� Like Upper([2]) And B.����=[3]"
            ElseIf zlCommFun.IsCharChinese(strInput) Then
                strWhere = " And B.���� Like [2] And B.����=[3]"
            End If
             '��ɢװʱ��Ʒ����ʾ���Ҳ���ʾ���
            strSQL = "" & _
            "   Select  distinct A.ID,A.����,A.����,A.���㵥λ" & _
            "   From ������ĿĿ¼ A,������Ŀ���� B" & _
            "   Where A.ID=B.������ĿID  And A.���='7' " & str����ʱ�� & strWhere
            
            If int��ҩ��̬ = 0 Then
                'ɢװ����ʾ�����,����ԭ������
                strSQL = _
                " Select distinct  A.ID as ҩ��ID,C.ҩƷID as ID,C.ҩƷID,D.����,A.����,D.���,A.���㵥λ as ������λ," & _
                        IIf(gblnסԺ��λ, "C.סԺ��λ", "D.���㵥λ") & " as ��λ,D.����,D.��������,d.ִ�п��� AS ִ�п���_ID," & IIf(mint���� <> 0, "N.���� ҽ������,", "") & _
                "       Decode(D.�Ƿ���,1,'ʱ��',LTrim(To_Char(Sum(F.�ּ�)" & _
                        IIf(gblnסԺ��λ, "*Nvl(C.סԺ��װ,1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as ����," & _
                        IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, " LTrim(To_Char(X.���" & IIf(gblnסԺ��λ, "/Nvl(C.סԺ��װ,1)", "") & ",'9999990.00000'))", "Decode(Sign(X.���),1,'��','��')") & " as ���" & _
                " From ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,�շѼ�Ŀ F, " & vbNewLine & _
                            IIf(mint���� <> 0, "����֧����Ŀ M,����֧������ N,", "") & vbNewLine & _
                "          (" & strSQL & ") A, " & vbNewLine & _
                "          (" & strStock & ") X" & vbNewLine & _
                " Where   A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And C.ҩƷID=X.ҩƷID(+) " & vbNewLine & _
                "        And D.ID=F.�շ�ϸĿID " & vbNewLine & _
                         IIf(mint���� <> 0, " And C.ҩƷID=M.�շ�ϸĿID(+) And M.����(+)=[5] And M.����ID=N.ID(+)" & vbNewLine, "") & _
                "        And exists(Select 1 From �շ�ִ�п��� A1 Where A1.�շ�ϸĿID=C.ҩƷID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine & _
                "        And Sysdate Between F.ִ������ and Nvl(F.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                         str��� & str���� & str��׼��Ŀ & _
                " Group by A.ID,C.ҩƷID,A.���㵥λ,D.����,A.����,D.���,D.����,D.��������,d.ִ�п���,D.�Ƿ���," & IIf(mint���� <> 0, "N.����,", "") & "X.���," & _
                        IIf(gblnסԺ��λ, "C.סԺ��λ,C.סԺ��װ", "D.���㵥λ") & _
                " Order by D.����"
            Else
                 '��ɢװʱ��Ʒ����ʾ���Ҳ���ʾ���
                strSQL = strSQL & _
                "        And exists(Select 1 From ����ִ�п��� A1 Where A1.������ĿID=A.ID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine
                strSQL = _
                    " Select Distinct A.ID,A.ID as ҩ��ID,A.����,A.����,A.���㵥λ as ��λ" & _
                    " From ҩƷ���� E,(" & strSQL & ") A" & _
                    " Where A.ID=E.ҩ��ID  " & _
                    "         And Exists(Select 1 From ҩƷ��� C Where C.ҩ��ID=E.ҩ��ID And Nvl(C.��ҩ��̬,0) = [6])" & _
                    "         And Rownum<=100" & _
                    " Order by A.����"
            End If
        
            vPoint = zlControl.GetCoordPos(vsBill.hWnd, vsBill.CellLeft, vsBill.CellTop)
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�в�ҩ", False, "", "", False, False, True, vPoint.X, vPoint.Y, _
                        vsBill.CellHeight, blnCancel, True, True, strInput & "%", gstrLike & strInput & "%", gbytCode + 1, lngҩ��ID, mint����, int��ҩ��̬, mint������Դ, mlng��������ID)
    Else
            If int��ҩ��̬ = 0 Then
                'ɢװ����ʾ�����,����ԭ������
            strSQL = "" & _
                " Select 0 as ĩ��,ID,ID as ҩ��ID ,�ϼ�ID,����,����,Null as ���,NULL as ������λ,NULL as ��λ," & _
                "       NULL as ����,NULL as �������� , NULL as ִ�п���_ID" & IIf(mint���� = 0, "", ",Null as ҽ������") & ",NULL as ����,NULL as ���,NULL as ҩƷID" & _
                " From ���Ʒ���Ŀ¼ " & _
                "  Where ����=3 And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                " Union ALL "
                strSQL = strSQL & _
                "  Select ĩ��,-1*Rownum As Id,ҩ��ID,�ϼ�ID,����,����,���,������λ,��λ,����,��������,ִ�п���_ID,����,���,ҩƷID " & _
                "  From ( " & _
                " Select 1 as ĩ��,A.ID,A.ID as ҩ��ID,A.����ID as �ϼ�ID,D.����,D.����,D.���,A.���㵥λ as ������λ," & _
                            IIf(gblnסԺ��λ, " C.סԺ��λ", "D.���㵥λ") & " as ��λ,D.����,D.��������,d.ִ�п��� as ִ�п���_ID" & IIf(mint���� = 0, "", ",N.���� ҽ������") & "," & _
                "           Decode(D.�Ƿ���,1,'ʱ��',LTrim(To_Char(Sum(F.�ּ�)" & _
                            IIf(gblnסԺ��λ, "*Nvl(C.סԺ��װ,1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as ����," & _
                            IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, " LTrim(To_Char(X.���" & IIf(gblnסԺ��λ, "/Nvl(C.סԺ��װ,1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))", "Decode(Sign(X.���),1,'��','��')") & " as ���,C.ҩƷID" & _
                " From  ������ĿĿ¼ A,ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,�շѼ�Ŀ F," & _
                            IIf(mint���� = 0, "", "           ����֧����Ŀ M,����֧������ N,") & _
                "           (" & strStock & ") X" & _
                " Where A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And C.ҩƷID =F.�շ�ϸĿID And A.���='7'  " & _
                        IIf(mint���� = 0, "", "       And C.ҩƷID=M.�շ�ϸĿID(+) And   M.����(+)=" & mint���� & " And M.����ID=N.ID(+)") & _
                "       And C.ҩƷID=X.ҩƷID(+) " & _
                "        And exists(Select 1 From �շ�ִ�п��� A1 Where A1.�շ�ϸĿID=C.ҩƷID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine & _
                "       And Sysdate Between F.ִ������ and Nvl(F.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                "       And D.������� IN(" & mint������Դ & ",3)" & str��׼��Ŀ & str��� & str����ʱ�� & _
                " Group by A.ID,A.���㵥λ ,A.����ID,D.����,D.����,D.���,D.����,D.��������,d.ִ�п���" & IIf(mint���� = 0, "", ",N.����") & ",D.�Ƿ���,X.���,C.ҩƷID," & _
                     IIf(gblnסԺ��λ, "C.סԺ��λ,C.סԺ��װ", "D.���㵥λ") & _
                ")"
            Else

                strSQL = "" & _
                " Select 0 as ĩ��,ID,ID as ҩ��ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ����ְ��ID" & _
                " From ���Ʒ���Ŀ¼ Where ����=3 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
                strSQL = strSQL & " UNION ALL " & _
                "Select Distinct 1 as ĩ��,A.ID,ID as ҩ��ID,A.����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,E.����ְ�� as ����ְ��ID" & _
                " From ������ĿĿ¼ A,ҩƷ���� E" & _
                " Where A.ID=E.ҩ��ID" & str���� & str����ʱ�� & str��� & _
                "        And exists(Select 1 From ����ִ�п��� A1 Where A1.������ĿID=A.ID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine
            End If
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "�в�ҩ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, True, "", "", "", lngҩ��ID, "", int��ҩ��̬, mint������Դ, mlng��������ID)
    End If
    
    With vsBill
        If rsTemp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õ���ҩ��Ŀ�����ȵ�������Ŀ���������á���", vbInformation, gstrSysName
            End If
            With vsBill
              If strInput <> "" Then .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col))
              Exit Function
            End With
        End If
        
        lngҩ��ID = Val(NVL(rsTemp!ҩ��ID)): lngҩƷID = 0
        If int��ҩ��̬ = 0 Then lngҩƷID = Val(NVL(rsTemp!ҩƷID))
        
        If ItemExist(lngҩ��ID, .Row, .Col) Then
           MsgBox "��ζ��ҩ���䷽���Ѿ�¼�롣", vbInformation, gstrSysName
            If strInput <> "" Then .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col))
            Exit Function
        End If
        
        lng�ϴ�ҩ��ID = Val(.Cell(flexcpData, .Row, GetBillCol(2, .Col))) '�ϴε�Ʒ��
        
        lngҩƷID = -1
        If lng�ϴ�ҩ��ID <> 0 Then
            If int��ҩ��̬ = 0 Then   '����ǵ�һζɢװҩ�������ˣ�����ҩ�����Ÿı�
                If .Row = .FixedRows And .Col = .FixedCols Then
                    If mcll���("_" & lng�ϴ�ҩ��ID) <> "" Then
                        lngҩƷID = Val(Split(mcll���("_" & lng�ϴ�ҩ��ID), ",")(0))
                    Else
                        lngҩƷID = 0
                    End If
                End If
            End If
            mcll���.Remove "_" & lng�ϴ�ҩ��ID
        End If
        
        '��ȡ����ֵ
        If strInput <> "" Then .EditText = rsTemp!����     'ֱ������ƥ��ʱ��Ҫ
         .TextMatrix(.Row, .Col) = rsTemp!����
         If int��ҩ��̬ = 0 Then
            .TextMatrix(.Row, .Col + 2) = NVL(rsTemp!������λ)
         Else
            .TextMatrix(.Row, .Col + 2) = rsTemp!��λ
         End If
         .Cell(flexcpData, .Row, .Col) = .TextMatrix(.Row, .Col)
         .Cell(flexcpData, .Row, .Col + 2) = lngҩ��ID    '��¼��ҩID
         If lng�ϴ�ҩ��ID <> lngҩ��ID And lng�ϴ�ҩ��ID <> 0 Then
            'ɾ���ϴ�ҩ��ID
            Call DeleteDetails(lng�ϴ�ҩ��ID)
            Err = 0: On Error Resume Next
            mcll���.Remove "_" & lng�ϴ�ҩ��ID
            Err = 0: On Error GoTo 0
         End If
         
        If mcll��� Is Nothing Then Set mcll��� = New Collection
        If int��ҩ��̬ = 0 Then
            Err = 0: On Error Resume Next
            mcll���.Remove "_" & lngҩ��ID
            Err = 0: On Error GoTo 0
            mcll���.Add NVL(rsTemp!ҩƷID) & ",0", "_" & lngҩ��ID
            
            '���ɢװҩƷ�Ĺ����ˣ��������ҩ��
            If lngҩƷID <> Val(NVL(rsTemp!ҩƷID)) Then
                If cboҩ��.ListIndex < 0 Then
                    If mlng����ID <> 0 Then
                        lngҩ��ID = Get�շ�ִ�п���ID("7", Val(NVL(rsTemp!ҩƷID)), NVL(rsTemp!ִ�п���_ID, 0), mlng���˿���ID, mlng��������ID, mint������Դ, mlng��ҩ��, mrsPati!����ID)
                    Else
                        lngҩ��ID = Get�շ�ִ�п���ID("7", Val(NVL(rsTemp!ҩƷID)), NVL(rsTemp!ִ�п���_ID, 0), mlng���˿���ID, mlng��������ID, mint������Դ, mlng��ҩ��)
                    End If
                Else
                    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
                End If
                
                '������ҩ��
                If Not gbln���뷢ҩ Then
                    If cboҩ��.ListIndex <> -1 Then
                        lngTmp = cboҩ��.ItemData(cboҩ��.ListIndex)
                    End If
                    If lngTmp <> lngҩ��ID And lngҩ��ID <> 0 Then
                        cboҩ��.ListIndex = cbo.FindIndex(cboҩ��, lngҩ��ID)
                        '�ı��˿ⷿ,��Ҫ����ˢ�¹��
                         Call ����ˢ��������ҩ���
                    End If
                End If
            End If
        Else
            Err = 0: On Error Resume Next
            mcll���.Remove "_" & lngҩ��ID
            Err = 0: On Error GoTo 0
            mcll���.Add "", "_" & lngҩ��ID
            If cboҩ��.ListIndex < 0 Then
                    If mlng��ҩ�� <> 0 Then
                        cboҩ��.ListIndex = cbo.FindIndex(cboҩ��, mlng��ҩ��)
                    Else
                       If cboҩ��.ListCount <> 0 Then cboҩ��.ListIndex = 0
                    End If
            End If
        End If
         If cboҩ��.ListCount <> 0 And cboҩ��.ListIndex < 0 Then cboҩ��.ListIndex = 0
        '����������ʱ���޸�ҩ��
        Call �ֽ���ҩ���(lngҩ��ID, Val(.TextMatrix(.Row, .Col + 1)))
        If Val(.TextMatrix(.Row, .Col + 1)) <> 0 Then
            Call Show��ҩ���(lngҩ��ID, Val(.TextMatrix(.Row, .Col + 1)))
        End If
    End With
    Call ShowSpecs(lngҩ��ID)
    '����:39319
    If Not mcll��� Is Nothing Then
        sta.Panels(3).Text = "��" & mcll���.Count & "ζҩ"
    End If
    SelectChineDrug = True
End Function



Private Function ItemExist(ByVal lng��ҩID As Long, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    '���ܣ��ж���ҩ�䷽��������,ָ������ҩ�Ƿ��Ѿ�����
    Dim i As Long, j As Long, lngTemp As Long
    Dim lngCurCol As Long
    
    lngTemp = GetBillCol(2, lngCol)
    For i = 1 To vsBill.Rows - 1
        For j = 0 To vsBill.Cols - 1 Step MCOLS
            lngCurCol = GetBillCol(2, j)
            If lngRow = i And lngTemp <> lngCurCol Or lngRow <> i Then
                If Val(vsBill.Cell(flexcpData, i, lngCurCol)) = lng��ҩID Then
                    ItemExist = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Function is����ҩƷ(ByVal lngƷ��ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��Ƿ����ҩƷ
    '���أ��Ƿ���true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-03 17:05:01
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngҩƷID As Long, rsTemp As ADODB.Recordset
    Dim i As Long, strSQL As String
    
    If lngƷ��ID = 0 Then Exit Function
    Err = 0: On Error Resume Next
    lngҩƷID = Val(Split(mcll���("_" & lngƷ��ID) & ",", ",")(0))
    If Err <> 0 Then
        ShowMsgbox "δ�ҵ�����,����!"
        is����ҩƷ = False
        Exit Function
    End If
    If lngҩƷID = 0 Then Exit Function
    For i = 1 To mobjDetails.Count
        If mobjDetails(i).Detail.ID = lngҩƷID Then
           is����ҩƷ = mobjDetails(i).Detail.����
           Exit Function
        End If
    Next
    
    On Error GoTo errHandle
    
    'δ������,ֱ�Ӵӿ��ж�ȡ
    strSQL = "Select ҩ������ From ҩƷ��� where ҩƷID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩƷID)
    If Not rsTemp.EOF Then
        is����ҩƷ = NVL(rsTemp!ҩ������, 0) <> 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetBillCol(ByVal int���� As Integer, lngCol As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ʵ���
    '���:  int����-0:��ҩ������,1-������;2-��λ��
    '���أ�ָ�����ʵ���
    '���ƣ����˺�
    '���ڣ�2010-08-03 17:29:28
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    GetBillCol = (lngCol \ MCOLS) * MCOLS + int����
End Function
Private Function GetBillDetailObject(ByVal lngҩƷID As Long) As BillDetail
    '��ȡ��ϸ���ݶ���
    Dim i As Long
    For i = 1 To mobjDetails.Count
         If mobjDetails(i).Detail.ID = lngҩƷID Then
            Set GetBillDetailObject = mobjDetails(i)
            Exit Function
         End If
    Next
End Function
Private Function CheckStock(ByVal lngҩƷID As Long, dbl���� As Double, Optional objDetail As Detail) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҩƷ���
    '���أ����ڿ��,����True,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-04 11:23:47
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngҩ��ID As Long
    If lngҩƷID = 0 Then CheckStock = True: Exit Function
    
    If objDetail Is Nothing Then
        Set objDetail = GetBillDetailObject(lngҩƷID).Detail
    End If
    
    If objDetail Is Nothing Then CheckStock = True: Exit Function
    If cboҩ��.ListIndex < 0 Then
        If cboҩ��.ListCount = 0 Then '33188
            CheckStock = True: Exit Function
        End If
        cboҩ��.ListIndex = 0
    End If
    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    
    'ҩƷ�����
    With objDetail
        If Not gbln���뷢ҩ Then
            If .���� Or .��� Then
                If Val(txt����.Text) * dbl���� > .��� Then
                    MsgBox """" & .���� & """Ϊ������ʱ��ҩƷ����ǰ���ÿ�治������������", vbInformation, gstrSysName
                    Exit Function
                End If
            ElseIf mcolStock("_" & lngҩ��ID) <> 0 Then
                If Val(txt����.Text) * dbl���� > .��� Then
                    If mcolStock("_" & lngҩ��ID) = 1 Then
                        If MsgBox("""" & .���� & """�ĵ�ǰ���ÿ�治������������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                             Exit Function
                        End If
                    ElseIf mcolStock("_" & lngҩ��ID) = 2 Then
                        MsgBox """" & .���� & """�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                         Exit Function
                    End If
                End If
            End If
        ElseIf gstr��ҩ�� <> "" And Val(txt����.Text) * dbl���� > .��� Then
            If gblnStock Then
                MsgBox "[" & .���� & "]�ĵ�ǰ���ÿ�治����������!", vbInformation, gstrSysName
                 Exit Function
            Else
                If MsgBox("[" & .���� & "]�ĵ�ǰ���ÿ�治������������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End With
    CheckStock = True
End Function
Private Function IsCheckStockEnough(ByVal lngҩƷID As Long, dbl���� As Double, Optional objDetail As Detail = Nothing) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��жϿ���Ƿ����
    '���أ����㷵��true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-04 17:11:52
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If objDetail Is Nothing Then
        Set objDetail = GetBillDetailObject(lngҩƷID).Detail
    End If
    If objDetail Is Nothing Then IsCheckStockEnough = True: Exit Function
    
    With objDetail
        If Val(txt����.Text) * dbl���� > .��� Then Exit Function
    End With
    IsCheckStockEnough = True
End Function
Private Sub vsBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'���ܣ���������ȷ��
    Dim rsTmp As ADODB.Recordset
    Dim lngҩ��ID As Long, lngҩƷID As Long
    Dim strSQL As String, blnCancel As Boolean
    Dim lngҩ��ID As Long, strStock As String, i As Long
    Dim vPoint As POINTAPI, strTmp As String
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim str��׼��Ŀ As String, blnOverFlow As Boolean
    Dim strInput As String, strSQLInput As String, strSQLAdd As String, strSQLItem As String
    Dim int��ҩ��̬ As Integer
    Dim dblϵ�� As Double
    
    If KeyAscii = 13 Then
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        KeyAscii = 0
        
        '��ȡ�س���,�����MsgboxʹEdit���㶪ʧ,�����ɱ༭,�����ἤ��AfterEdit�¼�
        If Col Mod MCOLS = col��ҩ Then
            '��ҩ����
            If vsBill.EditText = "" Then   'zyk
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            strInput = vsBill.EditText
            If SelectChineDrug(strInput) = False Then
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
            End If
        ElseIf Col Mod MCOLS = col���� Then
            '�������ת��
            vsBill.EditText = ConvertABCtoNUM(vsBill.EditText)
            '��������Ϸ��Լ��
            If Not IsNumeric(vsBill.EditText) Or Val(vsBill.EditText) > LONG_MAX Then
                MsgBox "ҩƷ����������󣬲�����ֵ���ͻ�������ֵ����", vbInformation, gstrSysName
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
            End If
            
            
            '��ҩδ����,��Ч
            lngҩ��ID = Val(vsBill.Cell(flexcpData, Row, GetBillCol(2, Col)))
            If lngҩ��ID = 0 Then
                MsgBox "���������в�ҩ��", vbInformation, gstrSysName
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
            End If
            
            If Val(vsBill.EditText) < 0 Then
                If InStr(mstrPrivsOpt, ";��ҩ��������;") = 0 Then
                    MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '����ɢװ����,������ֻ����������
                If Get��ҩ��̬ <> 0 Then
                    MsgBox "ҩƷ��ֻ̬��ɢװ�Ĳ��ܸ������ʡ�", vbInformation, gstrSysName
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                Else
                    'ɢװ��,Ҫ����Ƿ������,������Ҳ������������
                    If is����ҩƷ(Val(vsBill.Cell(flexcpData, Row, Col + 1))) Then
                        MsgBox "����ҩƷ���������븺����", vbInformation, gstrSysName
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                If mint���� > 0 Then
                    If Not gclsInsure.GetCapability(support��������, mlng����ID, mint����) Then
                        MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            End If
            If InStr(mstrPrivsOpt, ";ҩƷ����С��;") = 0 Then
                If Val(vsBill.EditText) <> Int(vsBill.EditText) Then
                    MsgBox "��û��Ȩ������С����", vbInformation, gstrSysName
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                End If
            End If
            
            strTmp = vsBill.EditText
            If Val(vsBill.EditText) = 0 Then
                If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                Else
                    vsBill.EditCell: vsBill.EditText = strTmp '���㶪ʧ��EditTextҲ��ʧ
                End If
            End If
            
            int��ҩ��̬ = Get��ҩ��̬
            If int��ҩ��̬ = 0 Then
                '��Ҫ�����
                lngҩƷID = Val(Split(mcll���("_" & lngҩ��ID) & ",", ",")(0))
                dblϵ�� = Get����ϵ��(lngҩƷID)
                dblϵ�� = IIf(dblϵ�� = 0, 1, dblϵ��)
                If CheckStock(lngҩƷID, FormatEx(Val(strTmp) / dblϵ��, 5)) = False Then
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                Else
                    vsBill.EditCell: vsBill.EditText = strTmp
                End If
            End If
            
            '��鴦������:����ʱ���
            vsBill.EditText = FormatEx(Val(vsBill.EditText), 5)
            strTmp = vsBill.EditText  '����Msgbox��vsBill.EditText�ᱻ���,������Ҫ���ȼ�¼
            
            If �ֽ���ҩ���(lngҩ��ID, Val(strTmp)) = False Then
                
            End If
            Call Show��ҩ���(lngҩ��ID, Val(strTmp))
            '�������
            If gcurMaxMoney > 0 Then
                For i = 1 To mobjDetails.Count
                        If mobjDetails(i).Detail.ҩ��ID = lngҩ��ID Then
                                If mobjDetails(i).InComes(1).Ӧ�ս�� > gcurMaxMoney Then
                                    If MsgBox("ҩƷΪ:" & mobjDetails(i).Detail.���� & " ���Ϊ" & mobjDetails(i).Detail.��� & vbCrLf & _
                                                      "�ĵ�ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                        mobjDetails(i).���� = Val(vsBill.Cell(flexcpData, Row, Col))
                                        Call �ֽ���ҩ���(lngҩ��ID, Val(vsBill.TextMatrix(Row, Col)))
                                        Call Show��ҩ���(lngҩ��ID, Val(vsBill.TextMatrix(Row, Col)))
                                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                                        Exit Sub
                                    End If
                                End If
                        End If
                Next
            End If
            
            
            Call GetBillTotalIncomes(curӦ��, curʵ��, blnOverFlow)
            If blnOverFlow Then
                '���,�ָ�����
                MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                Call �ֽ���ҩ���(lngҩ��ID, Val(vsBill.TextMatrix(Row, Col)))
                Call Show��ҩ���(lngҩ��ID, Val(vsBill.TextMatrix(Row, Col)))
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
                     
            '��Ҫ��ʵ�����֮��
            '���ʷ��ñ���������ǰ����,���䲡������������䲡�������������󱣴�ʱ�ж�
            mrsWarn.Filter = ""
            If mrsWarn.RecordCount > 0 And Not mrsPati Is Nothing Then
                Call GetBillTotalIncomes(, curʵ��)
                If curʵ�� > 0 Then
                    gbytWarn = BillingWarn(mstrPrivsOpt, mrsPati!���� & IIf(NVL(mrsPati!סԺ��) = "", "", "(סԺ��:" & mrsPati!סԺ�� & " ����:" & mrsPati!���� & ")"), Val("" & mrsPati!����ID), mrsPati!���ò���, mrsWarn, mrsPati!�������, _
                                Val("" & mrsPati!���ն�) - mcurModiMoney, curʵ�� + mcur����ҩ���, Val("" & mrsPati!������), 7, "�в�ҩ", mstrWarn, , gblnPrice)
                                        
                    If gbytWarn = 2 Or gbytWarn = 3 Then
                        Call �ֽ���ҩ���(lngҩ��ID, Val(vsBill.TextMatrix(Row, Col)))
                        Call Show��ҩ���(lngҩ��ID, Val(vsBill.TextMatrix(Row, Col)))
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                        sta.Panels(2).Text = "Ԥ��:" & Format(mrsPati!Ԥ�����, "0.00") & "/����:" & Format(mrsPati!Ԥ����� - mrsPati!�������, "0.00") & "/ʣ��:" & Format(mrsPati!�������, "0.00")
                        Exit Sub
                    End If
                End If
            End If
            'ȷ������
            vsBill.TextMatrix(Row, Col) = strTmp
            vsBill.Cell(flexcpData, Row, Col) = vsBill.TextMatrix(Row, Col)
            'ˢ�½����ʾ
            txtӦ��.Text = Format(curӦ��, gstrDec)
            txtʵ��.Text = Format(curʵ��, gstrDec)
            mblnChange = True
        End If
        Call EnterNextCell(Row, Col)
    ElseIf Col Mod MCOLS = col���� Then
        lngҩ��ID = Val(vsBill.Cell(flexcpData, Row, GetBillCol(2, Col)))
        'ҩ��δ����,�����������
        If lngҩ��ID = 0 Then
            KeyAscii = 0: Exit Sub
        End If
        strTmp = "0123456789" & gstrABC
        If InStr(mstrPrivsOpt, ";��ҩ��������;") > 0 Then
            If mint���� > 0 Then
                If gclsInsure.GetCapability(support��������, mlng����ID, mint����) Then strTmp = strTmp & "-"
            Else
                strTmp = strTmp & "-"
            End If
        End If
        '����ɢװ����,������ֻ����������
        If InStr(1, strTmp, "-") > 0 Then
            If Get��ҩ��̬ <> 0 Then
                 strTmp = Replace(strTmp, "-", "")
            Else
                'ɢװ��,Ҫ����Ƿ������,������Ҳ������������
                If is����ҩƷ(lngҩ��ID) Then
                    strTmp = Replace(strTmp, "-", "")
                End If
            End If
        End If
        If InStr(mstrPrivsOpt, ";ҩƷ����С��;") > 0 Then
            strTmp = strTmp & "."
        End If
        If InStr(strTmp & Chr(8) & Chr(27), UCase(Chr(KeyAscii))) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub
Public Function CheckDrugDataValied(ByVal lngҩƷID As Long, Optional strName As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҩƷ���ݵĺϷ���
    '���أ����ݺϷ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-04 12:01:51
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    '�������ҩƷȨ�޼��
    Set rsTemp = ReadҩƷ��Ϣ(lngҩƷID)
    If Not rsTemp Is Nothing Then
        If IIf(IsNull(rsTemp!�������), "", rsTemp!�������) = "����ҩ" _
            And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
            MsgBox """" & strName & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf IIf(IsNull(rsTemp!�������), "", rsTemp!�������) = "����ҩ" _
            And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
            MsgBox """" & strName & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf (IIf(IsNull(rsTemp!��ֵ����), "", rsTemp!��ֵ����) = "����" _
            Or IIf(IsNull(rsTemp!��ֵ����), "", rsTemp!��ֵ����) = "����") _
            And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
            MsgBox """" & strName & """Ϊ���ػ򰺹�ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDrugDataValied = True
    
End Function
Private Function zlGetDetail(ByVal lngҩƷID As Long, Optional dbl���� As Double, Optional ByRef objDetail As Detail) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ϸ����
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-08-04 16:28:00
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngTmp As Long, lngҩ��ID As Long
    
    Set objDetail = New Detail
    If cboҩ��.ListIndex < 0 Then
        MsgBox "��ѡ����ҩ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errH:
    

    If mint���� > 0 Then
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,A.����," & _
        "       A.���,A.���㵥λ,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�," & _
        "       A.ִ�п���,A.��������,A.����ժҪ,M.Ҫ������,C.ҩ������,C.ҩ��ID," & _
        "       C.סԺ��λ,C.סԺ��װ,J1.���� as ��������,J1.���㵥λ as ������λ,C.����ϵ��,A.�������" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,����֧����Ŀ M,����֧������ N,������ĿĿ¼ J1" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID And A.ID=[1] and C.ҩ��ID=J1.ID" & _
        " And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[2] And M.����ID=N.ID(+)"
    Else
        strSQL = _
            "   Select A.ID,A.���,B.���� as �������,A.����,A.����," & _
            "           A.���,A.���㵥λ,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�," & _
            "           A.ִ�п���,A.��������,A.����ժҪ,0 as Ҫ������,C.ҩ������,C.ҩ��ID," & _
            "           C.סԺ��λ,C.סԺ��װ,J1.���� as ��������,J1.���㵥λ as ���Ƶ�λ,J1.���㵥λ as ������λ,C.����ϵ��,A.�������" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,������ĿĿ¼ J1" & _
            " Where A.���=B.���� And A.ID=C.ҩƷID and C.ҩ��ID=J1.ID And A.ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩƷID, mint����)
    '����շ��뷢ҩ����ʱ����������ʱ�ۼ�����ҩƷ
    If gbln���뷢ҩ Then
        If NVL(rsTmp!�Ƿ���, 0) = 1 Or NVL(rsTmp!ҩ������, 0) = 1 Then
            MsgBox "��ҩ���봦��ʱ��������ʱ�ۻ����ҩƷ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '����Ӧ����֧����Ŀ,������ö����������:mbytFun=0
    If mint���� <> 0 Then
        If Not CheckMediCareItem(lngҩƷID, mint����, "" & rsTmp!����, Val(NVL(rsTmp!�Ƿ���)) <> 1) Then
            Exit Function
        End If
    End If
    '�������ҩƷȨ�޼��
    If CheckDrugDataValied(lngҩƷID, NVL(rsTmp!����)) = False Then Exit Function
    '��ؿ����
    '---------------------------------------------------------------------------------------
    objDetail.ID = rsTmp!ID
    objDetail.ҩ��ID = rsTmp!ҩ��ID
    objDetail.���� = rsTmp!����
    objDetail.���� = rsTmp!����
    objDetail.���㵥λ = NVL(rsTmp!���㵥λ)
    objDetail.��� = NVL(rsTmp!���)
    objDetail.��� = rsTmp!���
    objDetail.������� = rsTmp!�������
    objDetail.��� = NVL(rsTmp!�Ƿ���, 0) <> 0
    objDetail.���� = NVL(rsTmp!ҩ������, 0) <> 0
    objDetail.����ժҪ = NVL(rsTmp!����ժҪ, 0) <> 0
    objDetail.����ְ�� = Get����ְ��(rsTmp!ID)
    objDetail.�������� = Get��������(rsTmp!ID)
    objDetail.�Ӱ�Ӽ� = NVL(rsTmp!�Ӱ�Ӽ�, 0) <> 0
    objDetail.���ηѱ� = NVL(rsTmp!���ηѱ�, 0) <> 0
    objDetail.סԺ��װ = NVL(rsTmp!סԺ��װ, 1)
    objDetail.סԺ��λ = NVL(rsTmp!סԺ��λ)
    objDetail.ִ�п��� = NVL(rsTmp!ִ�п���, 0)
    objDetail.���� = NVL(rsTmp!��������)
    objDetail.Ҫ������ = NVL(rsTmp!Ҫ������, 0) = 1
    objDetail.��ҩ��̬ = Get��ҩ��̬
    objDetail.�������� = NVL(rsTmp!��������)
    objDetail.������λ = NVL(rsTmp!������λ)
    objDetail.����ϵ�� = NVL(rsTmp!����ϵ��)
    objDetail.������� = Val(NVL(rsTmp!�������))
    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    If Not gbln���뷢ҩ Then
        objDetail.��� = GetStock(rsTmp!ID, lngҩ��ID)
        If gblnסԺ��λ Then
            objDetail.��� = objDetail.��� / objDetail.סԺ��װ
        End If
    ElseIf gstr��ҩ�� <> "" Then
        objDetail.��� = GetMultiStock(rsTmp!ID, gstr��ҩ��)
        If objDetail.��� = 0 And gblnStock Then
            MsgBox "[" & objDetail.���� & "]�Ŀ��ÿ��Ϊ��!", vbInformation, gstrSysName
            Exit Function
        End If
        If gblnסԺ��λ Then
            objDetail.��� = objDetail.��� / objDetail.סԺ��װ
        End If
    End If
    zlGetDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SetBillDetail(ByVal lngҩƷID As Long, Optional dbl���� As Double, Optional lng��� As Long = 0, _
    Optional objDetail As Detail = Nothing, Optional objBillDetail As BillDetail) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������һζ�в�ҩʱ������۸����ö����Լ���ʾ����
    '         lng���-��ǰ��������
    '���� : ������ϸ������
    '���ƣ����˺�
    '���ڣ�2010-08-02 17:55:57
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim objInComes As New BillInComes
    Dim rsTmp As ADODB.Recordset, strSQL As String, lngRow  As Integer
    Dim strժҪ As String, lngҩ��ID As Long
    Dim rsҩƷ��Ϣ As ADODB.Recordset
    
    lngRow = GetRow(vsBill.Row, vsBill.Col)
    If objDetail Is Nothing Then
        If zlGetDetail(lngҩƷID, dbl����, objDetail) = False Then Exit Function
    End If
    
    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    strժҪ = GetժҪ(lngҩƷID)
    
    If objDetail.����ժҪ Then
        If frmInputBox.InputBox(Me, "ժҪ", "������""" & objDetail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
            strժҪ = strժҪ
        End If
    Else
'        If mint���� <> 0 Then '90304
            strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, lngҩƷID, strժҪ, 2)
'        End If
    End If
    Call SetժҪ(lngҩƷID, strժҪ)
    Dim dblTemp As Double
    dblTemp = FormatEx(dbl���� / objDetail.����ϵ��, 5)
    If gblnסԺ��λ Then     '52722
        dblTemp = dblTemp / IIf(objDetail.סԺ��װ = 0, 1, objDetail.סԺ��װ)
        Set objBillDetail = mobjDetails.Add(objDetail, lngҩƷID, lngRow, 0, mlng����ID, 0, 0, 0, "", "", "", 0, 0, _
                mstr�ѱ�, 0, objDetail.���, objDetail.סԺ��λ, "", Val(txt����.Text), dblTemp, 0, lngҩ��ID, objInComes, "", lngRow & "_" & lngҩƷID, , , , , , strժҪ)
    Else
        Set objBillDetail = mobjDetails.Add(objDetail, lngҩƷID, lngRow, 0, mlng����ID, 0, 0, 0, "", "", "", 0, 0, _
                mstr�ѱ�, 0, objDetail.���, objDetail.���㵥λ, "", Val(txt����.Text), dblTemp, 0, lngҩ��ID, objInComes, "", lngRow & "_" & lngҩƷID, , , , , , strժҪ)
    End If


    
    mblnChange = True
    SetBillDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetժҪ(ByVal lngҩƷID As Long) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ����¼�����ժҪ
    '���ƣ����˺�
    '���ڣ�2010-08-03 12:02:37
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strժҪ As String
    Err = 0: On Error Resume Next
    strժҪ = mcllInput����ժҪ("K" & lngҩƷID)
    GetժҪ = strժҪ
    Err = 0: On Error GoTo 0
End Function
Private Sub SetժҪ(ByVal lngҩƷID As Long, strժҪ As String)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ����¼�����ժҪ
    '���ƣ����˺�
    '���ڣ�2010-08-03 12:02:37
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    mcllInput����ժҪ.Remove "K" & lngҩƷID
    mcllInput����ժҪ.Add strժҪ, "K" & lngҩƷID
    Err = 0: On Error GoTo 0
End Sub
Private Sub ����ˢ��������ҩ���()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����������ҩ�Ĺ��(����������),��������ʾ��ǰ��ҩ����������������б�
    '���ƣ����˺�
    '���ڣ�2010-08-03 14:44:47
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngҩ��ID, dbl���� As Double, lngҩƷID As Long
    Dim int��ҩ��̬ As Long, rsTmp As ADODB.Recordset
    
    int��ҩ��̬ = Get��ҩ��̬
    With vsBill
        For i = .FixedRows To .Rows - 1
            For j = 0 To .Cols - 1 Step MCOLS
                lngҩ��ID = Val(.Cell(flexcpData, i, j + 2))
                dbl���� = Val(.TextMatrix(i, j + 1))
                If lngҩ��ID <> 0 Then
                    If int��ҩ��̬ = 0 Then
                        '��¼�ϴ�ѡ���ҩƷID,�Ա�ָ�ѡ��
                        '����:45410
                        lngҩƷID = Val(Split(mcll���("_" & lngҩ��ID) & ",", ",")(0))
                        mcll���.Remove ("_" & lngҩ��ID)  '��ѡ���ID
                        Set rsTmp = Getɢװ���(lngҩ��ID)   'ȡȱʡ���
                        If rsTmp.RecordCount > 0 Then
                            If lngҩƷID <> 0 Then rsTmp.Find "ҩƷID=" & lngҩƷID, , adSearchForward, 1
                            If rsTmp.EOF Then rsTmp.MoveFirst
                            mcll���.Add rsTmp!ҩƷID & ",0", "_" & lngҩ��ID
                        Else
                            mcll���.Add IIf(lngҩƷID = 0, "", lngҩƷID & ",0"), "_" & lngҩ��ID
                        End If
                    End If
                    
                    Call �ֽ���ҩ���(lngҩ��ID, dbl����, , False)
                    If mcll���("_" & lngҩ��ID) = "" Or InStr(mcll���("_" & lngҩ��ID), "|") > 0 Then
                        .Cell(flexcpForeColor, i, j + 1) = vbRed
                    Else
                        .Cell(flexcpForeColor, i, j + 1) = .ForeColor
                    End If
                End If
            Next
        Next
        lngҩ��ID = Val(.Cell(flexcpData, .Row, (.Col \ MCOLS) * MCOLS + 2))
        dbl���� = Val(.TextMatrix(.Row, (.Col \ MCOLS) * MCOLS + 1))
    End With
    If lngҩ��ID <> 0 Then Call Show��ҩ���(lngҩ��ID, dbl����)
    Call ReCalcӦ�պϼ�
End Sub

Private Function Getɢװ���(ByVal lngҩ��ID As Long) As ADODB.Recordset
'���ܣ���ȡ��ǰҩƷ�����п��õ�ɢװ���
    Dim lngҩ��ID As Long, strSQL As String

    On Error GoTo errHandle
    
    If gblnStock Then
        If cboҩ��.ListIndex <> -1 Then lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
        If lngҩ��ID <> 0 Then
            strSQL = " And Exists(Select 1 From ҩƷ��� B" & _
                " Where (Nvl(b.����, 0) = 0 Or b.Ч�� Is Null Or b.Ч��>Trunc(Sysdate))" & _
                " And b.����=1 And b.�ⷿID=[4] And a.ҩƷID=b.ҩƷID Group by b.ҩƷID" & _
                " Having Sum(b.��������)>0)"
        End If
    End If
    strSQL = "Select a.ҩ��id, a.ҩƷid, d.���, d.����, a.����ϵ��,A.סԺ��λ,A.סԺ��װ,D.���㵥λ," & _
            "       d.����, d.����,A.��ҩ��̬,D.ִ�п���" & vbNewLine & _
            "From ҩƷ��� A, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where a.ҩ��id = [1] And a.��ҩ��̬ = 0 And a.ҩƷID = d.ID" & strSQL & vbNewLine & _
            " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([3],3)" & _
            " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) Order By D.����"
    Set Getɢװ��� = zlDatabase.OpenSQLRecord(strSQL, "����б�", lngҩ��ID, lngҩ��ID, mint������Դ, lngҩ��ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function �ֽ���ҩ���(ByVal lngҩ��ID As Long, ByVal dbl���� As Double, _
    Optional objDetail As Detail, Optional ReCalӦ�� As Boolean = True, Optional int��ҩ��̬ As Integer = -1) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ֽ���ҩ�������
    '���: dbl����-������λ����
    '���� :��������ɹ�,����true, ���򷵻�ʧ��!
    '���ƣ����˺�
    '���ڣ�2010-08-02 17:50:45
    '˵����31867
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, i As Long
    Dim str������� As String, lngҩ��ID As Long
    Dim varData As Variant, varTemp As Variant
    Dim blnObj As Boolean
    
    If int��ҩ��̬ = -1 Then int��ҩ��̬ = Get��ҩ��̬
    blnObj = Not objDetail Is Nothing
    If int��ҩ��̬ = 0 Then
        'ɢװ��������ʱ��ȷ�����
        str������� = mcll���("_" & lngҩ��ID)
        If str������� <> "" Then str������� = Split(str�������, ",")(0) & "," & dbl����
    Else
        '2.������,ҩƷid,����;ҩƷid,����;...|ʣ������
        On Error GoTo errH
        If cboҩ��.ListIndex <> -1 Then lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
        strSQL = "Select Zl_Dispensechspecs([1],[2],[3],[4],[5],[6]) as txt From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "������", lngҩ��ID, int��ҩ��̬, dbl����, Val(txt����.Text), lngҩ��ID, IIf(gbln���뷢ҩ, 1, 0))
        str������� = "" & rsTemp!txt
    End If
    
    Call mcll���.Remove("_" & lngҩ��ID)
    mcll���.Add str�������, "_" & lngҩ��ID
    
    If str������� = "" Then
        'ɾ������ҩ��ID��Ϣ
        Call DeleteDetails(lngҩ��ID)
        Exit Function
    End If
    
    '--����:ҩƷid,����;ҩƷid,����;...(ɢװֻѡ��һ�����)
    '--                             ������ȫ����ʱ����:����Ϊ6��10�������,17�˵ķ���=23755,6;23756,10|1
    '--                             ���ܷ���ʱ���ؿ�,����:����Ϊ6��10�������,3�˵ķ���
    'ɾ������ҩ��ID��Ϣ
    Call DeleteDetails(lngҩ��ID)
    Dim objBillDetail As BillDetail
    '�����ϸ����
    varData = Split(Split(str�������, "|")(0), ";")
    For i = 0 To UBound(varData)
         varTemp = Split(varData(i), ",")
         '������ϸ
         If Not blnObj Then Set objDetail = Nothing
         
         If SetBillDetail(Val(varTemp(0)), Val(varTemp(1)), i, objDetail, objBillDetail) = False Then
            '�ֽ�ʧ��
            Call DeleteDetails(lngҩ��ID)
            If ReCalӦ�� Then Call ReCalcӦ�պϼ�
            Exit Function
         Else
            '�����շ���Ŀ����
            Call zlCalcMoney(objBillDetail, True)
         End If
    Next
    If ReCalӦ�� Then Call ReCalcӦ�պϼ�
    �ֽ���ҩ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ReCalcӦ�պϼ�()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����Ӧ��ʵ�պϼƺϼ�
    '���ƣ����˺�
    '���ڣ�2010-08-04 11:40:25
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    '����ϼ���
    Dim curӦ�� As Currency, curʵ�� As Currency
    Call GetBillTotalIncomes(curӦ��, curʵ��)
    txtӦ��.Text = Format(curӦ��, gstrDec)
    txtʵ��.Text = Format(curʵ��, gstrDec)
End Sub
Private Sub DeleteDetails(ByVal lngҩ��ID As Long)
    'ɾ��ҩ��ID�����й��
    Dim blnNotFond As Boolean, i As Long
     Do While True
        'ɾ��ҩ��ID
        blnNotFond = True
        For i = 1 To mobjDetails.Count
            If mobjDetails(i).Detail.ҩ��ID = lngҩ��ID Then
                 blnNotFond = False
                  mobjDetails.Remove i: Exit For
            End If
        Next
        If blnNotFond = True Then Exit Do
    Loop
End Sub
 
Private Sub GetBillTotalIncomes(Optional curӦ�� As Currency, Optional curʵ�� As Currency, Optional blnOvweFlow As Boolean)
'������blnOvweFlow=�����Ƿ����
    Dim i As Long, j As Long
    
    curӦ�� = 0: curʵ�� = 0: blnOvweFlow = False
    For i = 1 To mobjDetails.Count
        For j = 1 To mobjDetails(i).InComes.Count
            'Ҫ��VALתΪDouble��������
            If Abs(Val(curӦ��) + Val(mobjDetails(i).InComes(j).Ӧ�ս��)) > 922337203685477# Then
                blnOvweFlow = True: Exit Sub
            End If
            If Abs(Val(curʵ��) + Val(mobjDetails(i).InComes(j).ʵ�ս��)) > 922337203685477# Then
                blnOvweFlow = True: Exit Sub
            End If
            curӦ�� = curӦ�� + mobjDetails(i).InComes(j).Ӧ�ս��
            curʵ�� = curʵ�� + mobjDetails(i).InComes(j).ʵ�ս��
        Next
    Next
End Sub

 
Private Function Get��ҩ��̬() As Integer
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ��ǰ����ҩ��̬
    '���أ�0-ɢװ;1-��Ƭ;2-����
    '���ƣ����˺�
    '���ڣ�2010-07-27 14:58:51
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    For i = 0 To opt��̬.UBound
        If opt��̬(i).Value = True Then Exit For
    Next
    Get��ҩ��̬ = i
End Function
Private Sub zlCalcMoney(objBillDetail As BillDetail, Optional bln������ As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����¼���ָ��ҩƷ�еļ۸�ͽ��
    '��Σ�objbillDetail-ָ������ϸ����
    '          bln������-�����е��۴���
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-08-03 14:18:34
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, objInCome As New BillInCome
    Dim dblAllTime As Double, dblMoney As Double, dblPrice As Double, dbl�Ӱ�Ӽ��� As Double
    Dim str�ѱ� As String, cur��� As Currency
    Dim strInfo As String, strSQL As String, i As Long, dblPriceSingle As Double

    On Error GoTo errH
     If Not bln������ Then Call AdjustCpt(objBillDetail.�շ�ϸĿID)

    strSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ��� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objBillDetail.�շ�ϸĿID)
    
    If Not rsTmp.EOF Then
        '�Ȼ�ȡ����Ա��ǰ����ı�۽��
        If objBillDetail.Detail.��� Then
            If InStr(",5,6,7,", objBillDetail.�շ����) > 0 Then
                '����ҩƷʱ��(�����򲻷���)
                '��Ȼ�м�¼(�������Ŀʱ���ж�)
                dblAllTime = objBillDetail.���� * objBillDetail.����
                If gblnסԺ��λ Then
                    '���ʱ�۰��ۼ��������м���
                    dblAllTime = dblAllTime * objBillDetail.Detail.סԺ��װ
                End If
                If dblAllTime <> 0 Then
                    dblPrice = Getʱ��ҩƷӦ�ս��(objBillDetail.ִ�в���ID, objBillDetail.�շ�ϸĿID, dblAllTime, gstrDec, dblPriceSingle)
                    If dblAllTime <> 0 Then
                        '����δ�ֽ����
                        MsgBox "�� " & Split(objBillDetail.Key, "_")(0) & " ��ʱ��ҩƷ""" & objBillDetail.Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                        dblMoney = 0
                    Else
                        'ע�⣺���������ֻ�ܱ���4λС��,�Ұ�Round����,������Ҫ�ֹ�����;�����������ڼ��㾫������������
                        dblAllTime = objBillDetail.���� * objBillDetail.����
                        If gblnסԺ��λ Then
                            '���ۼ���������ʵ��
                            dblAllTime = dblAllTime * objBillDetail.Detail.סԺ��װ
                        End If
                        dblMoney = IIf(dblPriceSingle = 0, Format(dblPrice / dblAllTime, gstrFeePrecisionFmt), dblPriceSingle) '�������ǰ��ۼ۵�λ
                    End If
                Else
                    dblMoney = 0
                End If
            Else
                If objBillDetail.InComes.Count = 0 Then
                    '�����һ�μ�����,���Ĭ��ȡԭ��
                    dblMoney = 0    'dblMoney = Nvl(rsTmp!ԭ��, 0)
                Else
                    dblMoney = objBillDetail.InComes(1).��׼����
                    '����û�����ı�۲������۷�Χ����ȡĬ��ֵ
                    If Abs(dblMoney) > Abs(NVL(rsTmp!�ּ�, 0)) Then
                        dblMoney = NVL(rsTmp!ԭ��, 0)
                    End If
                End If
            End If
        End If

        '�����ԭ�м�¼
        Set objBillDetail.InComes = New BillInComes

        '��д���з��ü�¼
        For i = 1 To rsTmp.RecordCount
            Set objInCome = New BillInCome
            With objInCome
                .������ĿID = rsTmp!������ĿID
                .������Ŀ = rsTmp!����
                .�վݷ�Ŀ = NVL(rsTmp!�վݷ�Ŀ)
                .ԭ�� = NVL(rsTmp!ԭ��, 0)
                .�ּ� = NVL(rsTmp!�ּ�, 0)
                If objBillDetail.Detail.��� Then
                    If InStr(",5,6,7,", objBillDetail.�շ����) > 0 And gblnסԺ��λ Then
                        .��׼���� = Format(dblMoney * objBillDetail.Detail.סԺ��װ, gstrFeePrecisionFmt)
                    Else
                        .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                    End If
                Else
                    If InStr(",5,6,7,", objBillDetail.�շ����) > 0 And gblnסԺ��λ Then
                        .��׼���� = Format(NVL(rsTmp!�ּ�, 0) * objBillDetail.Detail.סԺ��װ, gstrFeePrecisionFmt)
                    Else
                        .��׼���� = Format(NVL(rsTmp!�ּ�, 0), gstrFeePrecisionFmt)
                    End If
                End If

                'Ӧ�ս��=���� * ���� * ����
                If InStr(",5,6,7,", objBillDetail.�շ����) > 0 _
                    And objBillDetail.Detail.��� Then
                    .Ӧ�ս�� = dblPrice '��֤Ӧ�ս�������۽��û�����
                Else
                    .Ӧ�ս�� = .��׼���� * objBillDetail.���� * objBillDetail.����
                End If

                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If mbln�Ӱ� And objBillDetail.Detail.�Ӱ�Ӽ� Then
                    dbl�Ӱ�Ӽ��� = NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100
                    .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
                End If

                .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))

                dblAllTime = objBillDetail.���� * objBillDetail.����
                If gblnסԺ��λ Then dblAllTime = dblAllTime * objBillDetail.Detail.סԺ��װ
                
                If objBillDetail.Detail.���ηѱ� Then
                    .ʵ�ս�� = .Ӧ�ս��
                Else
                    'ҩƷ���ɱ��ۼ���,��������
                    
                    .ʵ�ս�� = CCur(Format(ActualMoney(mstr�ѱ�, .������ĿID, .Ӧ�ս��, _
                        objBillDetail.�շ�ϸĿID, objBillDetail.ִ�в���ID, dblAllTime, dbl�Ӱ�Ӽ���), gstrDec))
                End If
                objBillDetail.�ѱ� = mstr�ѱ�

                '��ȡ��Ŀ������Ϣ,����ֻ��ҽ�����˲���,������ö����������:And mbytFun = 0
                If mint���� <> 0 Then
                    strInfo = gclsInsure.GetItemInsure(mlng����ID, objBillDetail.�շ�ϸĿID, .ʵ�ս��, True, mint����, _
                        objBillDetail.ժҪ & "||" & dblAllTime)
                    If strInfo <> "" Then
                        objBillDetail.������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        objBillDetail.���մ���ID = Val(Split(strInfo, ";")(1))
                        .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        objBillDetail.���ձ��� = CStr(Split(strInfo, ";")(3))
                                                
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then objBillDetail.ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then objBillDetail.Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If

                objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
            End With
            rsTmp.MoveNext
        Next
    Else
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set objBillDetail.InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
Private Sub Show��ҩ���(ByVal lngҩ��ID As Long, dbl���� As Double, Optional int��ҩ��̬ As Long = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݵ�ǰ�к��У���ʾ��������ҩ����б�
    '���ƣ����˺�
    '���ڣ�2010-08-03 14:55:39
    '˵���������ɢװ��̬������ؿ�ѡ��Ĺ�������б�
    '------------------------------------------------------------------------------------------------------------------------
    Dim str������� As String, varData As Variant, arrValue As Variant
    Dim i As Long, strҩƷIDs As String, lngColBegin As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String, lngҩƷID As Long
    Dim bln��ɢװ��ɢװ As Boolean
    
    lngColBegin = (vsBill.Col \ MCOLS) * MCOLS
    vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vsBill.ForeColor
    
    cmd��̬.Visible = False
        
    With vs��ҩ���
        .Rows = .FixedRows
        .ColComboList(.ColIndex("���")) = ""
        If dbl���� = 0 Then Exit Sub
        If int��ҩ��̬ = -1 Then int��ҩ��̬ = Get��ҩ��̬
        
        str������� = Trim(mcll���("_" & lngҩ��ID))
        .Redraw = flexRDNone
        
        If str������� = "" Then
            .Rows = .FixedRows + 1
            '���ܷ���ʱ���ؿ�,����:����Ϊ6��10�������,3�˵ķ���
            .MergeCells = flexMergeRestrictRows
            If int��ҩ��̬ = 0 Then
                strMsg = "��ҩƷû�п��õ�ɢװ��̬����ѡ������ҩƷ����̬��"
            Else
                strMsg = "�޷����������������ù����䣬�����������"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("���")) = strMsg
            .TextMatrix(.Rows - 1, .ColIndex("����")) = strMsg
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, .ColIndex("����")) = dbl����
            .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = .TextMatrix(.Rows - 1, .ColIndex("����"))
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("����")) = vbRed
            vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
        Else
            varData = Split(Split(str�������, "|")(0), ";")
            If InStr(str�������, "|") > 0 Then
                 .Rows = .FixedRows + UBound(varData) + 2
            Else
                 .Rows = .FixedRows + UBound(varData) + 1
            End If
            
            For i = 0 To UBound(varData)
                arrValue = Split(varData(i), ",")
                strҩƷIDs = strҩƷIDs & "," & Val(arrValue(0))
                .Cell(flexcpData, .FixedRows + i, .ColIndex("���")) = Val(arrValue(0)) '���ID
                .TextMatrix(.FixedRows + i, .ColIndex("����")) = arrValue(1)    '����
                .Cell(flexcpData, .FixedRows + i, .ColIndex("����")) = .TextMatrix(.FixedRows + i, .ColIndex("����"))
            Next
            strҩƷIDs = Mid(strҩƷIDs, 2)
            
            On Error GoTo errH:
            If int��ҩ��̬ = 0 Then
                '�������п���(�п��)��ɢװ����Ա����ѡ�������Ĺ��
                Set rsTmp = Getɢװ���(lngҩ��ID)
            Else
                strSQL = "" & _
                    "   Select /*+ Rule*/A.ҩƷID,D.���,D.����,A.����ϵ��,A.��ҩ��̬ ," & _
                    "           A.סԺ��λ,A.סԺ��װ,D.���㵥λ From ҩƷ��� A,�շ���ĿĿ¼ D,Table(f_Num2List([1])) B" & vbNewLine & _
                    "   Where A.ҩƷID = B.Column_value And A.ҩƷID = D.ID"
                '������̬��������Ϊ�ù������ǻ��ɵ�ɢװ������ǰ��ѡ�����Ƭ
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����б�", strҩƷIDs)
                '���ֻ��һ����¼,����ɢװ��̬��,����Ϊ�Ƿ�ɢװ��,����ɢװ��
                If rsTmp.RecordCount <> 0 Then
                    bln��ɢװ��ɢװ = rsTmp.RecordCount = 1 And Val(NVL(rsTmp!��ҩ��̬)) = 0
                End If
            End If
            For i = .FixedRows To .Rows - 1
                If InStr(str�������, "|") > 0 And i = .Rows - 1 Then
                '���һ����ʾδ��������
                    .MergeCells = flexMergeRestrictRows
                    strMsg = "�޷����������������ù����䣬�����������"
                    .TextMatrix(i, .ColIndex("���")) = strMsg
                    .TextMatrix(i, .ColIndex("����")) = strMsg
                    .MergeRow(i) = True
                    .Cell(flexcpForeColor, i, .ColIndex("���")) = vbRed
                    .TextMatrix(i, .ColIndex("����")) = Split(str�������, "|")(1)
                    .Cell(flexcpData, i, .ColIndex("����")) = .TextMatrix(i, .ColIndex("����"))
                    vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
                Else
                    lngҩƷID = Val(CStr(.Cell(flexcpData, i, .ColIndex("���"))))
                    rsTmp.Filter = "ҩƷID = " & lngҩƷID
                    If rsTmp.RecordCount = 0 Then 'ɢװ����治��ʱ�������棩
                        strMsg = "��ǰҩ����治�㣬����û��ɢװ���"
                        .TextMatrix(.Rows - 1, .ColIndex("���")) = strMsg
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = strMsg
                        .MergeRow(.Rows - 1) = True
                        .Cell(flexcpForeColor, i, .ColIndex("����")) = vbRed
                         vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
                    Else
                        .TextMatrix(i, .ColIndex("���")) = "" & rsTmp!���
                        .Cell(flexcpData, i, .ColIndex("���")) = "" & rsTmp!��� '����ɢװ���ȡ������ѡ��ʱ�ָ�
                        .TextMatrix(i, .ColIndex("����")) = "" & rsTmp!����
                        .Cell(flexcpData, i, .ColIndex("����ϵ��")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("����"))) / IIf(Val(NVL(rsTmp!����ϵ��)) = 0, 1, Val(NVL(rsTmp!����ϵ��))), 5) '�ۼ۵�λ����
                        .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("����"))) / IIf(Val(NVL(rsTmp!����ϵ��)) = 0, 1, Val(NVL(rsTmp!����ϵ��))), 5)
                        .Cell(flexcpData, i, .ColIndex("����")) = Val(.Cell(flexcpData, i, .ColIndex("����"))) & ":" & IIf(Val(NVL(rsTmp!����ϵ��)) = 0, 1, Val(NVL(rsTmp!����ϵ��))) & ":" & IIf(Val(NVL(rsTmp!סԺ��װ)) = 0, 1, Val(NVL(rsTmp!סԺ��װ)))
                        '�ð�װ��λ��ʾ
                        If gblnסԺ��λ Then
                             .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(.TextMatrix(i, .ColIndex("����"))) / IIf(Val(NVL(rsTmp!סԺ��װ)) = 0, 1, Val(NVL(rsTmp!סԺ��װ))), 5) & NVL(rsTmp!סԺ��λ)
                        Else
                             .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) & NVL(rsTmp!���㵥λ)
                        End If
                        .TextMatrix(i, .ColIndex("����ϵ��")) = "" & rsTmp!����ϵ��
                        If int��ҩ��̬ = 0 Then
                            '��ҩ��̬��,����������ʾ��ɫ����
                            If IsCheckStockEnough(lngҩƷID, Val(.Cell(flexcpData, i, .ColIndex("����ϵ��")))) Then
                                .Cell(flexcpForeColor, i, .ColIndex("����")) = .ForeColor
                            Else
                                .Cell(flexcpForeColor, i, .ColIndex("����")) = vbRed
                                vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
                            End If
                        End If
                    End If
                End If
            Next
            
            'ɢװ��̬������ѡ����
            If int��ҩ��̬ = 0 Or bln��ɢװ��ɢװ Then
                If bln��ɢװ��ɢװ Then
                    '��Ҫ���Ĺ��
                    Set rsTmp = Getɢװ���(lngҩ��ID)
                End If
                
                rsTmp.Filter = ""
                If rsTmp.RecordCount > 1 Then
                    strҩƷIDs = ""
                    For i = 1 To rsTmp.RecordCount
                        strҩƷIDs = strҩƷIDs & "|#" & rsTmp!ҩƷID & ";" & rsTmp!���� & "-" & rsTmp!���� & IIf(Not IsNull(rsTmp!���), "(" & rsTmp!��� & ")", "")
                        rsTmp.MoveNext
                    Next
                    .ColComboList(.ColIndex("���")) = Mid(strҩƷIDs, 2)
                    rsTmp.MoveFirst
                    .RowData(.FixedRows) = rsTmp   'ֻ��һ��
                    .Cell(flexcpBackColor, .FixedRows, .ColIndex("���")) = &HF0F4E4
                End If
            End If
        End If
        
        If int��ҩ��̬ <> 0 Then
            '��ɢװ��̬��δ������ʱ������Ϊɢװ
            If str������� = "" Or InStr(str�������, "|") > 0 Then
                Set rsTmp = Getɢװ���(lngҩ��ID)
                If rsTmp.RecordCount > 0 Then
                    strMsg = "�޷����������������ù����䣬��������������ɢװ��"
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = strMsg
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = strMsg
                    .MergeRow(.Rows - 1) = True
                    
                    .Select .Rows - 1, .ColIndex("����ϵ��")
                    cmd��̬.Visible = True
                    cmd��̬.Tag = rsTmp!ҩƷID  'ȱʡ���
                    cmd��̬.Caption = "ɢװ(&D)"
                    cmd��̬.Top = vs��ҩ���.CellTop
                    cmd��̬.Left = vs��ҩ���.CellLeft
                    cmd��̬.Width = vs��ҩ���.CellWidth
                    cmd��̬.Height = vs��ҩ���.CellHeight
                End If
            End If
        End If
        .Redraw = True
        .Visible = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vs��ҩ���_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim str������� As String, lngҩƷID As Long
    Dim objDetail As Detail, dbl���� As Double
    
    With vs��ҩ���
        Select Case .Col
        Case .ColIndex("���")
            If .ComboData = "" Then
                'û��ѡ��ʱ�ƿ�����
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
            Else
                'ɢװ��ҩ��ѡ����֮��
                lngҩƷID = CLng(.ComboData)
                If zlGetDetail(lngҩƷID, Val(.Cell(flexcpData, Row, .ColIndex("����ϵ��"))), objDetail) = False Then
                     .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '���㵥λ����:���ڼ���ϵ����
                If CheckStock(lngҩƷID, Val(.Cell(flexcpData, Row, .ColIndex("����ϵ��"))), objDetail) = False Then
                     .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                Set rsTmp = .RowData(.FixedRows)
                rsTmp.Filter = "ҩƷID = " & lngҩƷID
                str������� = mcll���("_" & rsTmp!ҩ��ID)
                dbl���� = Val(.Cell(flexcpData, Row, .ColIndex("����")))
                mcll���.Remove "_" & rsTmp!ҩ��ID
                mcll���.Add rsTmp!ҩƷID & "," & dbl����, "_" & rsTmp!ҩ��ID
                
                If �ֽ���ҩ���(Val(NVL(rsTmp!ҩ��ID)), dbl����, objDetail, True, 0) = False Then
                        mcll���.Remove "_" & rsTmp!ҩ��ID
                        mcll���.Add str�������, "_" & rsTmp!ҩ��ID
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                
                If IsCheckStockEnough(lngҩƷID, FormatEx(dbl���� / IIf(Val(NVL(rsTmp!����ϵ��)) = 0, 1, Val(NVL(rsTmp!����ϵ��))), 5), objDetail) = False Then
                    vsBill.Cell(flexcpForeColor, vsBill.Row, GetBillCol(1, vsBill.Col)) = vbRed
                    .Cell(flexcpForeColor, Row, .ColIndex("����")) = vbRed
                Else
                    vsBill.Cell(flexcpForeColor, vsBill.Row, GetBillCol(1, vsBill.Col)) = vsBill.ForeColor
                    .Cell(flexcpForeColor, Row, .ColIndex("����")) = .ForeColor
                End If
                .TextMatrix(Row, .ColIndex("���")) = Trim(NVL(rsTmp!���))
                .Cell(flexcpData, Row, Col) = Trim(NVL(rsTmp!���))   '���ڻָ�
                .TextMatrix(Row, .ColIndex("����")) = Trim(NVL(rsTmp!����))
                .TextMatrix(Row, .ColIndex("����ϵ��")) = Trim(NVL(rsTmp!����ϵ��))
                .Cell(flexcpData, Row, .ColIndex("����ϵ��")) = FormatEx(dbl���� / IIf(Val(NVL(rsTmp!����ϵ��)) = 0, 1, Val(NVL(rsTmp!����ϵ��))), 5)
                
                If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs��ҩ���_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vs��ҩ���
        Select Case NewCol
        Case .ColIndex("���")
            If opt��̬(0).Value And .ColComboList(NewCol) <> "" Then
                 .FocusRect = flexFocusSolid
            Else
                 .FocusRect = flexFocusLight
            End If
        End Select
    End With
End Sub

Private Sub vs��ҩ���_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs��ҩ���
        Select Case Col
        Case .ColIndex("���")
             If Not (opt��̬(0).Value Or .ColComboList(Col) <> "") Then
                    Cancel = True
             End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vs��ҩ���_ChangeEdit()
    Call vs��ҩ���_AfterEdit(vs��ҩ���.Row, vs��ҩ���.Col)
End Sub

Private Sub vs��ҩ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vs��ҩ���_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vs��ҩ���.ComboIndex <> -1 Then
            Call vs��ҩ���_KeyPress(13)
        End If
    End If
End Sub
Private Function Get��ҩ���(ByVal lngҩ��ID As Long, Optional ByVal lng��̬ As Long = -1) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ҩ����ID��ȡ��ҩ���
    '���أ����ع��ļ�¼��
    '���ƣ����˺�
    '���ڣ�2010-08-05 11:19:04
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngҩ��ID As Long
    
    On Error GoTo errH
    If lng��̬ = 0 Then
        Set Get��ҩ��� = Getɢװ���(lngҩ��ID)
    Else
        If gblnStock Then
            lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
            If lngҩ��ID <> 0 Then
                strSQL = " And Exists(Select 1 From ҩƷ��� B" & _
                    " Where (Nvl(b.����, 0) = 0 Or b.Ч�� Is Null Or b.Ч��>Trunc(Sysdate))" & _
                    " And b.����=1 And b.�ⷿID=[4] And a.ҩƷID=b.ҩƷID Group by b.ҩƷID" & _
                    " Having Sum(b.��������)>0)"
            End If
        End If
    
        strSQL = "" & _
        "   Select A.ҩƷID,A.��ҩ��̬,D.����,D.ִ�п��� " & _
        "   From ҩƷ��� A,�շ���ĿĿ¼ D Where A.ҩ��ID = [1] And A.ҩƷID = D.ID" & _
                    IIf(lng��̬ = -1, "", " And A.��ҩ��̬ = [3]") & strSQL & _
         "          And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([2],3)" & _
         "          And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) " & _
         "  Order by D.����"
         
        Set Get��ҩ��� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҩ���", lngҩ��ID, mint������Դ, lng��̬, lngҩ��ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowSpecs(ByVal lngҩ��ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ҩ����ع����
    '����:���˺�
    '����:2011-01-04 14:13:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngҩ��ID As Long, rsTemp As ADODB.Recordset
    Dim intCols As Integer, lngRow As Long, intCol As Integer
    Dim lngWidth As Long, i As Long
    
    On Error GoTo errH
    lngҩ��ID = 0
    If lngҩ��ID = 0 Then
        vsSpecShow.Clear
        vsSpecShow.Rows = 0
    End If
    If cboҩ��.ListIndex >= 0 Then lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
     
    strSQL = "" & _
    "   Select  D.����,D.���,D.����,E.���� as ҩ��," & _
    "      " & IIf(gblnסԺ��λ, "A.סԺ��λ", "D.���㵥λ") & " as סԺ��λ ," & _
                IIf(gblnסԺ��λ, "nvl(A.סԺ��װ,1)", "1") & "   as סԺ��װ,D.���㵥λ," & _
    "      Sum(nvl(M.��������,0))/" & IIf(gblnסԺ��λ, "nvl(A.סԺ��װ,1)", "1") & " as ��������" & _
    "   From ҩƷ��� M,ҩƷ��� A,�շ���ĿĿ¼ D,���ű� E" & _
    "   Where M.ҩƷID = D.ID and M.ҩƷID=A.ҩƷID and M.�ⷿID=E.ID " & _
    "            And  (Nvl(M.����, 0) = 0 Or M.Ч�� Is Null Or M.Ч��>Trunc(Sysdate)) " & _
    "            And A.ҩ��ID = [1]   " & IIf(lngҩ��ID = 0, "", " And M.�ⷿID=[3] ") & _
     "           And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([2],3)" & _
     "           And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) " & _
     "  Group by E.����,D.����,D.���,D.���� ,D.���㵥λ" & IIf(gblnסԺ��λ, ",A.סԺ��λ", "") & "" & IIf(gblnסԺ��λ, ",nvl(A.סԺ��װ,1)", "") & _
     "  Having Sum(nvl(M.��������,0))>0 " & _
     "  Order by D.����"
     
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҩ���", lngҩ��ID, mint������Դ, lngҩ��ID)
    intCols = 3
    With vsSpecShow
        .Clear
        .Rows = 0: .Cols = intCols
        intCol = intCols: lngRow = -1
        lngWidth = (.Width / intCol) - 30
        For i = 0 To .Cols - 1
            .ColWidth(i) = lngWidth
        Next
        Do While Not rsTemp.EOF
            If intCol >= intCols Then
                lngRow = lngRow + 1
                .Rows = .Rows + 1
                intCol = 0
            End If
           .TextMatrix(lngRow, intCol) = IIf(lngҩ��ID = 0, "(" & NVL(rsTemp!ҩ��) & ")", "") & NVL(rsTemp!���, "�޹��") & ":" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") = 0, "�п��", Val(NVL(rsTemp!��������)) & NVL(rsTemp!סԺ��λ))
           intCol = intCol + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount = 0 Then
            .Rows = 1
            .TextMatrix(0, 0) = "�޿��!"
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function Get����ϵ��(ByVal lngҩƷID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ϵ��
    '����:
    '����:���˺�
    '����:2011-02-18 11:35:29
    '����:35786
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl����ϵ�� As Double, strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
     strSQL = "Select max(����ϵ��) as ����ϵ�� From ҩƷ��� where ҩƷID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩƷID)
    dbl����ϵ�� = NVL(rsTemp!����ϵ��, 0)
    Get����ϵ�� = dbl����ϵ��
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


