VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frm�������� 
   Caption         =   "��������"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   Icon            =   "frm��������.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11790
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraҩƷ������Ϣ 
      Height          =   615
      Left            =   3720
      TabIndex        =   33
      Top             =   600
      Width           =   7215
      Begin VB.Label lblҩƷ������Ϣ 
         Caption         =   "ҩƷ��Ϣ"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7920
      TabIndex        =   29
      Top             =   3960
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frm��������.frx":6852
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.PictureBox pic������Ϣ 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   3375
      TabIndex        =   15
      Top             =   3840
      Width           =   3375
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cbo����ʱ�� 
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1720
         Width           =   2055
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   920
         Width           =   2055
      End
      Begin VB.TextBox txt��ʼNO 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txt����NO 
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   520
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTP�������� 
         Height          =   300
         Left            =   960
         TabIndex        =   23
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin MSComCtl2.DTPicker DTP��ʼ���� 
         Height          =   300
         Left            =   960
         TabIndex        =   24
         Top             =   2115
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin VB.Label lbl������ 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1350
         Width           =   615
      End
      Begin VB.Label lbl�������� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label lbl��ʼ���� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2145
         Width           =   735
      End
      Begin VB.Label lbl����ʱ�� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label lbl����� 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   945
         Width           =   615
      End
      Begin VB.Label lbl��ʼNO 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼNO"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lbl����NO 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "����NO"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   555
         Width           =   615
      End
   End
   Begin VB.PictureBox pic������Ϣ 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   11
      Top             =   1080
      Width           =   3375
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   97
         Width           =   2055
      End
      Begin VB.ComboBox cboʱ�䷶Χ 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   933
         Width           =   2055
      End
      Begin VB.CommandButton cmdҩƷ 
         Caption         =   "��"
         Height          =   300
         Left            =   2720
         TabIndex        =   3
         Top             =   517
         Width           =   255
      End
      Begin VB.TextBox txtҩƷ 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   517
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTP����ʱ�� 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   1770
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin MSComCtl2.DTPicker DTP��ʼʱ�� 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   31
         Top             =   1357
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin VB.Label lbl����ʱ�� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lbl��ʼʱ�� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label lblʱ�䷶Χ 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl�ⷿid 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ����"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblҩƷ 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   735
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2295
      Left            =   3840
      TabIndex        =   32
      Top             =   1680
      Width           =   7215
      _cx             =   12726
      _cy             =   4048
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483644
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483641
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fraEW 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   3840
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   240
      Width           =   45
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _Version        =   589884
      _ExtentX        =   5953
      _ExtentY        =   1508
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   7080
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm��������.frx":6DA0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14446
            Text            =   "ҩƷ�Ŀ��ÿ��"
            TextSave        =   "ҩƷ�Ŀ��ÿ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frm��������.frx":7634
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frm��������.frx":7B36
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1335
      Left            =   4440
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   2355
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
      FormatString    =   $"frm��������.frx":8038
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��������.frx":8086
            Key             =   "��ǰ"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgicon 
      Left            =   1680
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frm��������.frx":E8E8
   End
   Begin XtremeCommandBars.CommandBars combars 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane docpane 
      Bindings        =   "frm��������.frx":1435E
      Left            =   480
      Top             =   720
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const INTPANETYPE As Integer = 0
Private Const INTPANEDETAIL As Integer = 1
Private mfrmMain As Form
Private MStrCaption As String
Private mintUnit As Integer
Private mintCostDigit As Integer  '�ɱ��۵�С����λ��
Private mintPriceDigit As Integer
Private mintNumberDigit As Integer
Private mintMoneyDigit As Integer
Private mint����� As Integer
Private mint��������ⷿ As Integer
Private mbln�¿������� As Boolean
Private mlngģ��� As Long
Private mstr�����п� As String  '�����е�λ�ú��п�
Private mrsReturn As Recordset  '����ҩƷѡ����ѡ���ҩƷ
Private mInt���ݺ� As Integer   '����ҵ������

Private Const MCST_INVALIDCHAR As String = "'" '��ֹ������ַ�

'���嵥λ����
Private Const MCONINTPRICEUNIT As Integer = 1   '�ۼ۵�λ
Private Const MCONINTOUTUNIT As Integer = 2     '���ﵥλ
Private Const MCONINTINUNIT As Integer = 3      'סԺ��λ
Private Const MCONINTSTOREUNIT As Integer = 4   'ҩ�ⵥλ

'������ɫ����
Private Const CSTCOLOR_FIXED = &H808080        '��ɫ����ѡ���в��ܱ༭��������ɫ
Private Const CSTCOLOR_MODIFY = &HE0E0E0       '���ɫ��������޸�֮��ı���ɫ
Private Const CSTCOLOR_FONT = vbRed            '��ɫ��������Ԫ��������Ϊ0��������ɫ
Private Const CSTCOLOR_NOFONT = vbBlack        '��ɫ��������Ԫ������Ϊ0��������ɫ
Private Const CSTCOLOR_NOMODIFY = vbWhite      '��ɫ��������޸�֮ǰ�ı���ɫ
Private Const CSTCOLOR_ENTERCELL = &HFF0000    '��ɫ���������������Ԫ��ı߿���ɫ
Private Const CSTCOLOR_LOSTFORCE = &H80000005  '���ʧȥ����֮�󣬱��ѡ�е���ɫ

'�����г���
Private mintcolѡ�� As Integer
Private mintcolҩƷid As Integer
Private mintcol�к� As Integer
Private mIntColNO As Integer
Private mintcolҩƷ��������� As Integer
Private mintcol��Ʒ�� As Integer
Private mintcolҩƷ��Դ As Integer
Private mintcol����ҩ�� As Integer
Private mintcolҩ�ۼ��� As Integer
Private mintcol��� As Integer
Private mintcol��λ As Integer
Private mintcol���� As Integer
Private mintcol�������� As Integer
Private mintcol���� As Integer
Private mintcol���� As Integer
Private mintcol�������� As Integer
Private mintcol��Ч���� As Integer
Private mintcol�������� As Integer
Private mintcol��׼�ĺ� As Integer
Private mintcol������� As Integer   '���������ֶ�

Private mintcol�ɹ��޼� As Integer
Private mintcol�ɹ��� As Integer
Private mintcol���� As Integer
Private mintcol�ɱ��� As Integer    '�⹺�����Ϊ�����
Private mintcol�ɱ���� As Integer  '�⹺�����Ϊ������
Private mintcol�ӳ��� As Integer
Private mintcol�ۼ� As Integer
Private mintcol�ۼ۽�� As Integer
Private mintcol��� As Integer
Private mintcol������ As Integer
Private mintcol�������� As Integer
Private mintcol����� As Integer
Private mintcol������� As Integer

'�⹺�����Ҫ���ֶ�
Private mintcol���ۼ� As Integer
Private mintcol���۵�λ As Integer
Private mintcol���۽�� As Integer
Private mintcol���۲�� As Integer
Private mintcol�⹺��׼�ĺ� As Integer
Private mintcol������� As Integer
Private mintcol��Ʊ�� As Integer
Private mintcol��Ʊ���� As Integer
Private mintcol��Ʊ��Ϣ As Integer
Private mintcol��Ʊ��� As Integer

'��Ҫ���ص���
Private mintcol��ʵ���� As Integer
Private mintcol��� As Integer
Private mintcol����ϵ�� As Integer
Private mintcolҩ�� As Integer
Private mintcol���� As Integer
Private mintcol��¼״̬ As Integer
Private mintcol�������� As Integer
Private mintcol�������� As Integer
Private mintcol���Ч�� As Integer
Private mintcolʵ�ʲ�� As Integer
Private mintcolʵ�ʽ�� As Integer
Private mintcol�ϴι�Ӧ��ID As Integer
Private mintcolժҪ As Integer
Private mintcol�Է����� As Integer
Private mintcol�Ƿ��� As Integer
Private Const MINTCOL������ As Integer = 59

'��������ť�Ķ���
Private Const MINTBTNFILTER As Integer = 1          '���˰�ť
Private Const MINTBTNALLWRITEOFF As Integer = 2     'ȫ�尴ť
Private Const MINTBTNALLELIMINATE As Integer = 3    'ȫ�尴ť
Private Const MINTBTNDEL As Integer = 4             'ɾ����ť
Private Const MINTBTNWRITEOFF As Integer = 5        '������ť
Private Const MINTBTNHELP  As Integer = 6           '������ť
Private Const MINTBTNEXIT  As Integer = 7           '�˳���ť
Private Const MINTBTNSIMPLE  As Integer = 8         '��ఴť
Private Const MINTBTNCONPLETE   As Integer = 9      '������ť
Private mrecSort As Recordset

Private Sub cbo�ⷿ_Click()
    mint����� = MediWork_GetCheckStockRule(Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    mint��������ⷿ = MediWork_GetCheckStockRule(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
End Sub

Private Sub cbo�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub cboʱ�䷶Χ_Click()
    If Me.cboʱ�䷶Χ.ListIndex = 0 Then
        Me.DTP��ʼʱ��.Value = Date
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    ElseIf Me.cboʱ�䷶Χ.ListIndex = 1 Then
        Me.DTP��ʼʱ��.Value = Date - 1
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    ElseIf Me.cboʱ�䷶Χ.ListIndex = 2 Then
        Me.DTP��ʼʱ��.Value = Date - 2
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = False
        Me.DTP����ʱ��.Enabled = False
    Else
        Me.DTP��ʼʱ��.Value = Date - 30
        Me.DTP����ʱ��.Value = Date
        Me.DTP��ʼʱ��.Enabled = True
        Me.DTP����ʱ��.Enabled = True
    End If
End Sub
Private Sub cboʱ�䷶Χ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub cbo����ʱ��_Click()
    If Me.cbo����ʱ��.ListIndex = 0 Then
        Me.DTP��ʼ����.Value = Date
        Me.DTP��������.Value = Date
        Me.DTP��ʼ����.Enabled = False
        Me.DTP��������.Enabled = False
    ElseIf Me.cbo����ʱ��.ListIndex = 1 Then
        Me.DTP��ʼ����.Value = Date - 1
        Me.DTP��������.Value = Date
        Me.DTP��ʼ����.Enabled = False
        Me.DTP��������.Enabled = False
    ElseIf Me.cbo����ʱ��.ListIndex = 2 Then
        Me.DTP��ʼ����.Value = Date - 2
        Me.DTP��������.Value = Date
        Me.DTP��ʼ����.Enabled = False
        Me.DTP��������.Enabled = False
    Else
        Me.DTP��ʼ����.Value = Date - 30
        Me.DTP��������.Value = Date
        Me.DTP��ʼ����.Enabled = True
        Me.DTP��������.Enabled = True
    End If
End Sub
Private Sub cbo����ʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
'ҩƷʹ��ҩƷ��ͨ��ѡ����
Private Sub cmdҩƷ_Click()
    Dim vRect As RECT
    Dim strsql As String
    
    On Error GoTo errRow
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(6, "ҩƷ�⹺������", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    End If
    
'    Set mrsReturn = FrmҩƷѡ����.ShowME(Me, 6, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), , True, True, False, False, True, 0)
    Set mrsReturn = frmSelector.showMe(Me, 0, 6, , , , cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), , 0, True, True, True, , False)

    If Not mrsReturn.EOF Then
        Me.txtҩƷ.Text = mrsReturn!ͨ����
        Me.txtҩƷ.Tag = mrsReturn!ҩƷID
    End If
    Exit Sub
errRow:
    If ErrCenter <> 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cmdҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub combars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
Dim strKey As String
Dim count As Integer
Dim dblSum As Integer
Dim dblOldSum As Integer

Select Case Control.Id
    Case MINTBTNFILTER            'ִ�й��˲���
        Call Filter
    Case MINTBTNALLWRITEOFF       'ִ��ȫ�����
        Call AllWriteOff
    Case MINTBTNALLELIMINATE      'ִ��ȫ�����
        Call AllEliminate
    Case MINTBTNWRITEOFF          'ִ�г�������
        Call WriteOff
    Case MINTBTNDEL               'ִ��ɾ������
        Call DelRow
    Case MINTBTNHELP              'ִ�а�������
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
    Case MINTBTNEXIT              'ִ���˳�����
        Unload Me
    Case MINTBTNSIMPLE            'ִ������Ϊ���
        Call SetSimple(Control)
    Case MINTBTNCONPLETE          'ִ������Ϊ����
        Call SetConplete(Control)
End Select
End Sub

Private Sub InitData()
'----------------------------------------------
'ִ�����ݳ�ʼ����������Ҫ�Ǹ��������ó�sql��䣬Ȼ�󽫲�ѯ������浽vsflexfrid�����
'----------------------------------------------
    Dim strsql As String
    Dim rsData As Recordset
    Dim strOrder As String
    Dim strCompare As String
    Dim strSqlOrder As String
    Dim strUnitQuantity As String
    Dim i As Long
    Dim j As Long
    Dim int�ⷿid  As Long
    Dim int��װϵ�� As String
    Dim strҩƷ��Ϣ As String
    
    On Error GoTo errRow
    strOrder = zldatabase.GetPara("����", glngSys, mlngģ���)
    strCompare = Mid(strOrder, 1, 1)
    strSqlOrder = "���"
    
    '���������
    If strCompare = "0" Then
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        strSqlOrder = "ҩƷ����"
    ElseIf strCompare = "2" Then
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strSqlOrder = "ͨ����"
        Else
            strSqlOrder = "Nvl(��Ʒ��, ͨ����)"
        End If
    End If
    
    '����ķ�ʽ
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    int�ⷿid = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    
    'ȡ���ּ۸�ľ���
    Call GetDrugDigit(int�ⷿid, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Select Case mlngģ���
        Case ģ���.ҩƷ�ƿ�
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "C.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,(A.ʵ������ / B.�����װ) AS ʵ������,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,B.�����װ as ����ϵ��,"
                Case MCONINTINUNIT
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,B.סԺ��װ as ����ϵ��,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,(A.ʵ������ / B.ҩ���װ) AS ʵ������,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,B.ҩ���װ as ����ϵ��,"
            End Select
            
            strsql = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                        " FROM " & _
                        "     (SELECT DISTINCT a.NO,d.��¼״̬,A.ҩƷID,A.���,'[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, E.���� As ��Ʒ��," & _
                        "     B.ҩƷ��Դ,B.����ҩ��,C.���,C.���� AS ԭ����,A.����, A.����,A.����,B.�ӳ���,B.ҩ����� AS ��������," & _
                        "     B.���Ч��,A.Ч��," & strUnitQuantity & _
                        "     A.�ɱ����,0 ���۽��, 0 ���,D.ժҪ,A.�ⷿID,A.�Է�����ID,C.�Ƿ���,B.ҩ������ AS ҩ����������,A.�ϴι�Ӧ��ID,A.��׼�ĺ�,A.��д���� ��ʵ���� " & _
                        " ,D.������,D.��������,D.�����,D.�������,Y.���� �Է��ⷿ" & _
                        "     FROM " & _
                        "         (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,0 ʵ������,SUM(�ɱ����) AS �ɱ����,NO,ҩƷID,���,����," & _
                        "   ����,Ч��,NVL(����,0) ����,����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0) �ϴι�Ӧ��ID,��׼�ĺ�" & _
                        "          FROM ҩƷ�շ���¼ X " & _
                        "          WHERE ҩƷid=[1] AND ����=6 AND ���ϵ��=1 " & _
                        "          GROUP BY NO,ҩƷID,���,����,����,Ч��,NVL(����,0),����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0),��׼�ĺ�" & _
                        "          HAVING SUM(ʵ������)<>0) A," & _
                        "     ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� E,���ű� Y, " & _
                        " (Select NO,���,ժҪ,��¼״̬,������,��������,�����,������� From ҩƷ�շ���¼ " & _
                        "  Where ���� = 6 And ҩƷID = [1] and �ⷿID=[2] And ���ϵ�� = 1 And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) D " & _
                        "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=E.�շ�ϸĿID(+) AND A.�Է�����id=Y.id AND E.����(+)=3 AND B.ҩƷID=C.ID And A.��� = D.��� and d.no=a.no) W," & _
                        "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                        "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                        " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0)" & _
                        " And w.No Not in (Select NO From ҩƷ�շ���¼ Where ���� = 6 And ���ϵ�� = 1 And" & _
                        " ҩƷid = [1] And �ⷿid = [2] Having Sum(ʵ������) = 0 Group By NO, ���, ҩƷid)"
        Case ģ���.ҩƷ����
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "F.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,(A.ʵ������ / B.�����װ) AS ʵ������,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,B.�����װ as ����ϵ��,"
                Case MCONINTINUNIT
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,B.סԺ��װ as ����ϵ��,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,(A.ʵ������ / B.ҩ���װ) AS ʵ������,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,B.ҩ���װ as ����ϵ��,"
            End Select
        
            strsql = "Select w.*, z.�������� / w.����ϵ�� ��������, z.ʵ�ʽ��, z.ʵ�ʲ��" & _
                " From" & _
                " (Select Distinct a.No, x.��¼״̬, ������, ��������, �����, �������, a.ҩƷid," & _
                "a.���, '[' || f.���� || ']' As ҩƷ����, f.���� As ͨ����,e.���� As ��Ʒ��," & _
                "Nvl(e.����, f.����) ����, b.ҩƷ��Դ, b.����ҩ��, f.���, f.���� As ԭ����, a.����," & _
                "a.����, Nvl(a.����, 0) ����,b.�ӳ���, a.Ч��," & strUnitQuantity & _
                "a.�ɱ����, 0 ���۽��, 0 ���, a.ժҪ, a.�ⷿid,a.�Է�����id, c.���� As �Է��ⷿ," & _
                "f.�Ƿ���, b.ҩ������ As ҩ����������, a.������, a.��׼�ĺ�, a.��ҩ��ʽ, a.��д���� ԭʼ����" & _
                " From " & _
                " (Select Min(ID) As ID, Sum(ʵ������) As ��д����, 0 ʵ������, Sum(�ɱ����) As �ɱ����," & _
                "NO, ҩƷid, ���, ����, ����, Ч��, Nvl(����, 0) ����,����, �ɱ���, ���ۼ�, ժҪ, �ⷿid," & _
                "�Է�����id, ������id, Nvl(x.������, '') As ������, x.��׼�ĺ�, x.��ҩ��ʽ From ҩƷ�շ���¼ X" & _
                " Where ���� = 7 And ҩƷid = [1] Group By NO, ҩƷid, ���, ����, ����, Ч��, Nvl(����, 0), ����," & _
                "�ɱ���, ���ۼ�, ժҪ, �ⷿid, �Է�����id, ������id, ������, ��׼�ĺ�, ��ҩ��ʽ" & _
                " Having Sum(ʵ������) <> 0) A, ҩƷ��� B, �շ���Ŀ���� E, �շ���ĿĿ¼ F, ���ű� C," & _
                "(Select NO, ���, ժҪ, ��¼״̬, ������, ��������, �����, ������� From ҩƷ�շ���¼" & _
                " Where ���� = 7 And ҩƷid = [1] And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) X" & _
                " Where a.No = x.No And a.��� = x.��� And a.ҩƷid = b.ҩƷid And b.ҩƷid = f.Id" & _
                " And a.�Է�����id = c.Id And b.ҩƷid = e.�շ�ϸĿid(+) And e.����(+) = 3) W, ҩƷ��� Z" & _
                " Where w.ҩƷid = z.ҩƷid(+) And Nvl(w.����, 0) = Nvl(z.����(+), 0) And z.�ⷿid(+) = [2] And z.����(+) = 1"

        Case ģ���.�������
             Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,F.���㵥λ AS ��λ, A.��д���� AS ��д����,b.ָ�������� as ָ��������, a.�ɱ���,A.���ۼ�,1 as ����ϵ��,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,b.ָ��������*B.�����װ as ָ�������� , a.�ɱ���*B.�����װ as �ɱ���,A.���ۼ�*B.�����װ as ���ۼ� ,B.�����װ as ����ϵ��,"
                Case MCONINTINUNIT
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,b.ָ��������*B.סԺ��װ as ָ�������� , a.�ɱ���*B.סԺ��װ as �ɱ���,A.���ۼ�*B.סԺ��װ as ���ۼ� ,  B.סԺ��װ as ����ϵ��,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,b.ָ��������*B.ҩ���װ as ָ�������� , a.�ɱ���*B.ҩ���װ as �ɱ���,A.���ۼ�*B.ҩ���װ as ���ۼ� ,B.ҩ���װ as ����ϵ��,"
            End Select
            
            strsql = " Select w.*,z.�������� / w.����ϵ�� ��������" & _
                " From (Select Distinct a.no,a.ҩƷid, a.���, x.��¼״̬,  x.������, x. ��������,  x.�����,  x.�������,'[' || f.���� || ']' As ҩƷ����," & _
                "f.���� As ͨ����, e.���� As ��Ʒ��, b.ҩƷ��Դ, b.����ҩ��,b.ҩ�ۼ���,f.���,f.���� As ԭ����, a.����, a.����, b.���Ч��, a.Ч��," & _
                strUnitQuantity & "a.�ɱ����, 0 ���۽��, 0 ���, b.�ӳ��� ," & _
                "f.�Ƿ���, b.ҩ������ As ҩ����������, a.ժҪ, a.�ⷿid, g.���� As ����, a.������id," & _
                "a.��������, a.��׼�ĺ�, a.���,a.��д���� ��ʵ����, a.����,a.����" & _
                " From (Select Min(ID) As ID, Sum(ʵ������) As ��д����, Sum(�ɱ����) As �ɱ����, Sum(To_Number(Nvl(�÷�, 0))) As ����, no,ҩƷid," & _
                "��� , ����, ����, Ч��, ����, �ɱ���, ���ۼ�, ժҪ, �ⷿid, ������id, x.��������, x.��׼�ĺ�, x.���,nvl(����,0) ����" & _
                " From ҩƷ�շ���¼ X Where ���� = 4 And ҩƷid = [1]" & _
                " Group By no,ҩƷid, nvl(����,0),���, ����, ����, Ч��, ����, �ɱ���, ���ۼ�, ժҪ, �ⷿid, ������id, x.��������, x.��׼�ĺ�, x.���" & _
                " Having Sum(ʵ������) <> 0) A, ҩƷ��� B, �շ���Ŀ���� E, �շ���ĿĿ¼ F, ���ű� G," & _
                "(Select NO, ���, ժҪ, ��¼״̬, ������, ��������, �����, ������� From ҩƷ�շ���¼" & _
                " Where ���� = 4 And ҩƷid = [1] And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) X" & _
                " Where a.no=x.no and a.���=x.��� and a.ҩƷid = b.ҩƷid And b.ҩƷid = f.Id And " & _
                "a.�ⷿid = g.Id And b.ҩƷid = e.�շ�ϸĿid(+) And e.����(+) = 3 And e.����(+) = 1 and a.�ⷿid=[2]) w, ҩƷ��� Z" & _
                " Where w.ҩƷid = z.ҩƷid(+) And Nvl(w.����, 0) = Nvl(z.����(+), 0) And z.�ⷿid(+) = [2] And z.����(+) = 1"
        Case ģ���.��������
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "F.���㵥λ AS ��λ, A.��д���� as ��д����,a.�ɱ���,a.���ۼ�,nvl(a.����,0) As �����,'1' as ����ϵ��,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,nvl(a.����,0)*B.�����װ As �����,B.�����װ as ����ϵ��,"
                Case MCONINTINUNIT
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,nvl(a.����,0)*B.סԺ��װ As �����,B.סԺ��װ as ����ϵ��,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,nvl(a.����,0)*B.ҩ���װ As �����,B.ҩ���װ as ����ϵ��,"
            End Select
        
            strsql = "Select w.*, z.��������, z.ʵ�ʽ��, z.ʵ�ʲ��" & _
                    " From (Select Distinct a.no,a.ҩƷid,x.��¼״̬,x.������, x.��������,x. �����,x. �������, a.���, '[' || f.���� || ']' As ҩƷ����, f.���� As ͨ����," & _
                    "e.���� As ��Ʒ��, b.ҩƷ��Դ, b.����ҩ��, f.���,f.���� As ԭ����, a.����, a.����, a.����," & _
                    "b.�ӳ���, a.Ч��, g.���� As �����λ, h.���� As ������λ, a.��ֵ˰��," & strUnitQuantity & _
                    "a.�ɱ����, 0 ���۽��,0 ���, a.ժҪ, a.�ⷿid,a.������id, f.�Ƿ���, b.ҩ������ As ҩ����������," & _
                    "a.��׼�ĺ� From (Select Min(ID) As ID, Sum(ʵ������) As ��д����, Sum(�ɱ����) As �ɱ����, no,ҩƷid," & _
                    "���, ����, ����, Ч��, Nvl(����, 0) ����, ����,�ɱ���, ���ۼ�, ժҪ, �ⷿid, ������id, ����, ��ҩ����," & _
                    "��׼�ĺ�,To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000'))) As ��ֵ˰�� From ҩƷ�շ���¼ X" & _
                    " Where ���� = 11 And ҩƷid = [1] Group By no,ҩƷid, ���, ����, ����, Ч��, Nvl(����, 0), ����, �ɱ���," & _
                    "���ۼ�, ժҪ, �ⷿid, ������id, ����, ��ҩ����, ��׼�ĺ�," & _
                    "To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000')))" & _
                    " Having Sum(ʵ������) <> 0) A, ҩƷ��� B, �շ���Ŀ���� E, �շ���ĿĿ¼ F, ҩƷ�����λ G, ҩƷ������λ H," & _
                    "(Select NO, ���, ժҪ, ��¼״̬, ������, ��������, �����, ������� From ҩƷ�շ���¼" & _
                    " Where ���� = 11 And ҩƷid = [1] And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) X" & _
                    " Where a.no=x.no and a.���=x.��� and a.ҩƷid = b.ҩƷid And b.ҩƷid = f.Id And a.��ҩ���� = g.����(+)" & _
                    " And a.��ҩ���� = h.����(+) And b.ҩƷid = e.�շ�ϸĿid(+) And e.����(+) = 3 And e.����(+) = 1) W," & _
                    " (Select ҩƷid, Nvl(����, 0) ����, ��������, ʵ�ʽ��, ʵ�ʲ�� From ҩƷ���" & _
                    " Where �ⷿid = [2] And ���� = 1) Z" & _
                    " Where w.ҩƷid = z.ҩƷid(+) And Nvl(w.����, 0) = Nvl(z.����(+), 0)"
        Case ģ���.�⹺���
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,D.���㵥λ AS ��λ, A.��д���� AS ��д����,'1' as ����ϵ��, "
                    int��װϵ�� = "1"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,B.�����װ as ����ϵ��,"
                    int��װϵ�� = "B.�����װ"
                Case MCONINTINUNIT
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,B.סԺ��װ as ����ϵ��,"
                    int��װϵ�� = "B.סԺ��װ"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,B.ҩ���װ as ����ϵ��,"
                    int��װϵ�� = "B.ҩ���װ"
            End Select
        
            strsql = "Select w.*,z.�������� / w.����ϵ�� ��������" & _
                    " From (Select Distinct a.no,a.ҩƷid, a.���, x.��¼״̬, x.������, x.��������, x.�����, x.�������, '[' || d.���� || ']' As ҩƷ����," & _
                    "d.���� As ͨ����, e.���� As ��Ʒ��, b.ҩƷ��Դ, b.����ҩ��, d.���,d.���� As ԭ����, a.����, a.����, Nvl(b.�б�ҩƷ, 0) �б�ҩƷ," & _
                    "Nvl(b.���������, 0) ���������, b.���Ч��, a.Ч��," & strUnitQuantity & _
                    " nvl(A.����,b.ָ��������)*" & int��װϵ�� & " AS ָ�������� ,A.�ɱ���*" & int��װϵ�� & " AS �ɱ���," & _
                    " A.�ɱ���� AS �ɹ����,D.�Ƿ���,B.ҩ������ ҩ����������,  " & _
                    " DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�*" & int��װϵ�� & " AS ���ۼ� ,0 AS ���۽��,0 AS ���,A.����, " & _
                    "a.��׼�ĺ�, a.�������, a.��Ʊ��,a.��Ʊ����, a.��Ʊ����, a.��Ʊ���,a.��ҩ��λid, f.���� As ��Ӧ��, a.�ⷿid," & _
                    "g.���� As ����, Nvl(a.�������, 0) As �������, a.�˻�, a.��������, a.����,a.��ҩ�� As �˲���," & _
                    "a.��ҩ���� As �˲�����, b.ҩ�ۼ���, a.�ӳ��� From (Select Min(x.Id) As ID, Sum(ʵ������) As ��д����," & _
                    "Sum(�ɱ����) As �ɱ����, �������, ��Ʊ��,��Ʊ����, ��Ʊ����, Sum(��Ʊ���) As ��Ʊ���,x.no,x.ҩƷid, x.���," & _
                    "x.����, x.����, x.Ч��, x.����, x.�ɱ���, x.���ۼ�, x.����, x.��ҩ��λid, �ⷿid," & _
                    "Nvl(y.�������, 0) As �������, Nvl(x.��ҩ��ʽ, 0) As �˻�, x.��������, x.��׼�ĺ�, Nvl(x.����, 0) ����," & _
                    "x.��ҩ��, x.��ҩ����,Sum(To_Number(Nvl(�÷�, 0))) As ����, Ƶ�� As �ӳ��� From ҩƷ�շ���¼ X," & _
                    "(Select �շ�id, �������, �������, ��Ʊ��,��Ʊ����, ��Ʊ����, ��Ʊ��� From Ӧ����¼" & _
                    " Where ϵͳ��ʶ = 1 And ��¼���� =0) Y" & _
                    " Where x.Id = y.�շ�id(+) And  ���� = 1 and ҩƷid=[1]" & _
                    " Group By x.no,x.ҩƷid, x.���, x.����, x.����, x.Ч��, x.����, x.�ɱ���, x.���ۼ�, x.����, x.��ҩ��λid, x.�ⷿid, �������, ��Ʊ��,��Ʊ����, ��Ʊ����," & _
                    "Nvl(y.�������, 0), Nvl(x.��ҩ��ʽ, 0), x.��������, x.��׼�ĺ�, Nvl(x.����, 0), x.��ҩ��, x.��ҩ����, x.Ƶ��" & _
                    " Having Sum(ʵ������) <> 0) A, ҩƷ��� B, �շ���Ŀ���� E, �շ���ĿĿ¼ D, ��Ӧ�� F, ���ű� G," & _
                    "(Select NO, ���, ժҪ, ��¼״̬, ������, ��������, �����, ������� From ҩƷ�շ���¼" & _
                    " Where ���� = 1 And ҩƷid = [1] And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) X" & _
                    " Where a.No = x.No And a.��� = x.��� And a.ҩƷid = b.ҩƷid And b.ҩƷid = d.Id And a.�ⷿid = g.Id And" & _
                    " b.ҩƷid = e.�շ�ϸĿid(+) and a.�ⷿid=[2] And e.����(+) = 3 And a.��ҩ��λid = f.Id And Substr(f.����, 1, 1) = 1) w, ҩƷ��� Z" & _
                    " Where w.ҩƷid = z.ҩƷid(+) And Nvl(w.����, 0) = Nvl(z.����(+), 0) And z.�ⷿid(+) = [2] And z.����(+) = 1"
    End Select

    If Me.DTP��ʼʱ��.Value <> "" And Me.DTP����ʱ��.Value <> "" Then
        strsql = strsql + " and w.�������>=to_date('" & Me.DTP��ʼʱ��.Value & "','yyyy-mm-dd') and w.�������<=to_date('" & Me.DTP����ʱ��.Value & " 23:59:59','yyyy-mm-dd HH24:MI:SS')"
    End If
    
    If Me.txt�����.Text <> "" Then
        strsql = strsql + " and w.����� like '%" & Me.txt�����.Text & "%'"
    End If
    
    If Me.txt��ʼNO.Text <> "" Then
        If Me.txt����NO.Text <> "" Then
            strsql = strsql + " and w.NO>='" & Me.txt��ʼNO.Text & "'  and w.NO<='" & Me.txt����NO.Text & "'"
        Else
            strsql = strsql + " and w.NO='" & Me.txt��ʼNO.Text & "'"
        End If
    Else
        If Me.txt����NO.Text <> "" Then
            strsql = strsql + " and w.NO='" & Me.txt����NO.Text & "'"
        End If
    End If
    
    

    If Me.txt������.Text <> "" Then
        strsql = strsql + " and w.������ like '%" & Me.txt������.Text & "%'"
    End If
    
    If Me.DTP��ʼ����.Value <= Date And Me.DTP��������.Value <= Date Then
        strsql = strsql + " and w.��������>=to_date('" & Me.DTP��ʼ����.Value & "','yyyy-mm-dd') and w.��������<=to_date('" & Me.DTP��������.Value & " 23:59:59','yyyy-mm-dd HH24:MI:SS')"
    End If
    
    strsql = strsql + " ORDER BY NO," & strSqlOrder
    Set rsData = zldatabase.OpenSQLRecord(strsql, Me.Caption, Me.txtҩƷ.Tag, Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.ListIndex))
    
    If rsData.RecordCount > 0 Then
       '����ҩƷ�Ļ�����Ϣ
        strҩƷ��Ϣ = rsData!ҩƷ����
        strҩƷ��Ϣ = strҩƷ��Ϣ & IIf(gintҩƷ������ʾ <> 1, rsData!ͨ����, "")
        strҩƷ��Ϣ = strҩƷ��Ϣ & IIf(gintҩƷ������ʾ <> 0 And zlStr.Nvl(rsData!��Ʒ��) <> "", "(" & zlStr.Nvl(rsData!��Ʒ��) & ")", "")
        strҩƷ��Ϣ = strҩƷ��Ϣ & "   " & zlStr.Nvl(rsData!���)
        strҩƷ��Ϣ = strҩƷ��Ϣ & "   (" & zlStr.Nvl(rsData!��λ) & ")"
        
        Me.lblҩƷ������Ϣ.Caption = strҩƷ��Ϣ
        
        Me.vsfList.rows = rsData.RecordCount + 1
        For i = 1 To rsData.RecordCount
            With Me.vsfList
                .TextMatrix(i, mintcolҩƷid) = rsData!ҩƷID
                .TextMatrix(i, mIntColNO) = rsData!NO
                
               'ҩƷ������ʾ��ʽ
                If gintҩƷ������ʾ = 1 Then
                    .TextMatrix(i, mintcolҩƷ���������) = rsData!ҩƷ����
                Else
                    .TextMatrix(i, mintcolҩƷ���������) = rsData!ҩƷ���� & rsData!ͨ����
                End If
                
                .TextMatrix(i, mintcolҩ��) = zlStr.Nvl(rsData!ͨ����)
                .TextMatrix(i, mintcol��Ʒ��) = zlStr.Nvl(rsData!��Ʒ��)
                .TextMatrix(i, mintcolҩƷ��Դ) = zlStr.Nvl(rsData!ҩƷ��Դ)
                .TextMatrix(i, mintcol����ҩ��) = zlStr.Nvl(rsData!����ҩ��)
                .TextMatrix(i, mintcol���) = zlStr.Nvl(rsData!���)
                .TextMatrix(i, mintcol����) = zlStr.Nvl(rsData!����)
                .TextMatrix(i, mintcol��λ) = zlStr.Nvl(rsData!��λ)
                .TextMatrix(i, mintcol����) = zlStr.Nvl(rsData!����)
                .TextMatrix(i, mintcol��Ч����) = Format(zlStr.Nvl(rsData!Ч��), "yyyy-mm-dd")
                .TextMatrix(i, mintcol������) = zlStr.Nvl(rsData!������)
                .TextMatrix(i, mintcol��������) = Format(zlStr.Nvl(rsData!��������), "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, mintcol�����) = zlStr.Nvl(rsData!�����)
                .TextMatrix(i, mintcol�������) = Format(zlStr.Nvl(rsData!�������), "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, mintcol����) = zlStr.FormatEx(Val(rsData!��д����), mintNumberDigit, , True)
                .TextMatrix(i, mintcol��������) = 0
                .TextMatrix(i, mintcol�ɱ���) = zlStr.FormatEx(Val(rsData!�ɱ���), mintCostDigit, , True)
                .TextMatrix(i, mintcol�ɱ����) = 0
                .TextMatrix(i, mintcol�ۼ�) = zlStr.FormatEx(Val(rsData!���ۼ�), mintPriceDigit, , True)
                .TextMatrix(i, mintcol�ۼ۽��) = 0
                .TextMatrix(i, mintcol���) = 0
                .TextMatrix(i, mintcol���) = Val(rsData!���)
                .TextMatrix(i, mintcol����ϵ��) = Val(rsData!����ϵ��)
                .TextMatrix(i, mintcol��¼״̬) = Val(rsData!��¼״̬)
                .TextMatrix(i, mintcol��������) = zlStr.FormatEx(zlStr.Nvl(rsData!��������, 0), mintNumberDigit, , True)
                .TextMatrix(i, mintcol�Ƿ���) = zlStr.Nvl(rsData!�Ƿ���, 0)
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 0 And .TextMatrix(i, mintcol��Ч����) <> "" Then
                    '����ΪʧЧ��
                    .TextMatrix(i, mintcol��Ч����) = Format(DateAdd("D", 1, .TextMatrix(i, mintcol��Ч����)), "yyyy-mm-dd")
                End If
                
                If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Then .TextMatrix(i, mintcolҩ�ۼ���) = zlStr.Nvl(rsData!ҩ�ۼ���)
                If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Then .TextMatrix(i, mintcol��������) = Format(zlStr.Nvl(rsData!��������), "yyyy-mm-dd")

                If mlngģ��� = ģ���.ҩƷ���� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then .TextMatrix(i, mintcol��������) = zlStr.Nvl(rsData!�Է��ⷿ)
                If mlngģ��� <> ģ���.�⹺��� Then .TextMatrix(i, mintcol��׼�ĺ�) = zlStr.Nvl(rsData!��׼�ĺ�)
                If mlngģ��� = ģ���.������� Then .TextMatrix(i, mintcol�������) = zlStr.Nvl(rsData!���)

                If mlngģ��� = ģ���.�⹺��� Then
                    .TextMatrix(i, mintcol�ɹ���) = zlStr.FormatEx(Val(rsData!�ɱ���) / (Val(rsData!���� / 100)), mintCostDigit, , True)
                    .TextMatrix(i, mintcol�ɹ��޼�) = zlStr.FormatEx(Val(rsData!ָ��������), mintCostDigit, , True)
                    .TextMatrix(i, mintcol����) = Val(rsData!����)
                    .TextMatrix(i, mintcol�ӳ���) = Val(rsData!�ӳ���) * 100 & "%"
                End If
                
                If mlngģ��� <> ģ���.������� Then .TextMatrix(i, mintcol����) = Val(rsData!����)
                If mlngģ��� = ģ���.ҩƷ�ƿ� Or mlngģ��� = ģ���.������� Then .TextMatrix(i, mintcol��ʵ����) = zlStr.FormatEx(Val(rsData!��ʵ����), mintNumberDigit, , True)
                
                If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Then
                    .TextMatrix(i, mintcolʵ�ʲ��) = zlStr.FormatEx(.TextMatrix(i, mintcol�ۼ۽��) - .TextMatrix(i, mintcol�ɱ����), mintMoneyDigit, , True)
                Else
                    .TextMatrix(i, mintcolʵ�ʲ��) = zlStr.FormatEx(zlStr.Nvl(rsData!ʵ�ʲ��, 0), mintMoneyDigit, , True)
                End If
                If mlngģ��� <> ģ���.�⹺��� And mlngģ��� <> ģ���.������� Then .TextMatrix(i, mintcolʵ�ʽ��) = zlStr.FormatEx(zlStr.Nvl(rsData!ʵ�ʽ��, 0), mintMoneyDigit, , True)
                
                If mlngģ��� = ģ���.ҩƷ�ƿ� Then .TextMatrix(i, mintcol�Է�����) = Val(rsData!�Է�����id)
                
                If mlngģ��� = ģ���.�⹺��� Then
                    If Val(.TextMatrix(i, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(i, mintcol����), 0) <> 0 Then
                        .TextMatrix(i, mintcol���ۼ�) = zlStr.FormatEx(rsData!���ۼ� / Val(rsData!����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�, , True)
                        .TextMatrix(i, mintcol���۵�λ) = rsData!�ۼ۵�λ
                        .TextMatrix(i, mintcol���۽��) = 0
                        .TextMatrix(i, mintcol���۲��) = 0
                     End If
                    .TextMatrix(i, mintcol�⹺��׼�ĺ�) = zlStr.Nvl(rsData!��׼�ĺ�)
                    .TextMatrix(i, mintcol�������) = zlStr.Nvl(rsData!�������)
                    .TextMatrix(i, mintcol��Ʊ��) = zlStr.Nvl(rsData!��Ʊ��)
                    .TextMatrix(i, mintcol��Ʊ����) = zlStr.Nvl(rsData!��Ʊ����)
                    .TextMatrix(i, mintcol��Ʊ��Ϣ) = zlStr.Nvl(rsData!��Ʊ����)
                    .TextMatrix(i, mintcol��Ʊ���) = zlStr.FormatEx(zlStr.Nvl(rsData!��Ʊ���, 0), mintMoneyDigit, , True)
                End If
                .TextMatrix(i, mintcolժҪ) = ""
            End With
            
            If Not rsData.EOF Then rsData.MoveNext
        Next
    End If
    
    If vsfList.rows > 1 Then
        vsfList.Row = 1
        Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = True
        Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = True
        Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = True
        Me.vsfList.Cell(flexcpFontBold, 1, mintcol��������, Me.vsfList.rows - 1, mintcol��������) = True
    Else
        Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
        Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
        Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
    End If
    
    Exit Sub
errRow:
    If ErrCenter <> 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub combars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Id = MINTBTNCONPLETE Then
        If Control.Checked Then
            Control.IconId = 90003
        Else
            Control.IconId = 90004
        End If
    ElseIf Control.Id = MINTBTNSIMPLE Then
        If Control.Checked Then
            Control.IconId = 90003
        Else
            Control.IconId = 90004
        End If
    End If
End Sub

Private Sub Form_Load()
    Call SetTitle
    '�ָ�����
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        mstr�����п� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Caption, "�����п�", "")
    End If
    
    '����ҩƷ�����¼��ҩƷ�ƿ�ļ�¼��ͬ�����Ե����������ʱ���ڲ������ƿ�ķ�ʽ����
    If mlngģ��� = 1341 Then
        mlngģ��� = ģ���.ҩƷ�ƿ�
    End If
    
    Call InitComman
    Call InitTool
    Call InitTask
    Call InitCbo
    Call initGrid
    Call InitVSFColSel
    
    '��Ϊû�����ݣ����Բ������ݵİ�ť������
    Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
    Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
    Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False

    Me.DTP��ʼʱ��.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate))
    Me.DTP����ʱ��.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate))
    Me.DTP��ʼʱ��.Enabled = False
    Me.DTP����ʱ��.Enabled = False
    
    Me.DTP��ʼ����.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate)) + 1
    Me.DTP��������.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate)) + 1
    Me.DTP��ʼ����.Enabled = False
    Me.DTP��������.Enabled = False
    
    Call combars_Execute(Me.combars.Item(1).Controls.Item(MINTBTNSIMPLE))
    mbln�¿������� = (gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = 1)
End Sub

Private Sub SetTitle()
'------------------------------
'���ݲ�ͬ��ҵ�����ô���ı���
'-------------------------------
    Select Case mlngģ���
        Case ģ���.ҩƷ�ƿ�
            Me.Caption = "ҩƷ�ƿ���������"
            mInt���ݺ� = ���ݺ�.ҩƷ�ƿ�
            MStrCaption = "ҩƷ�ƿ����"
        Case ģ���.��������
            Me.Caption = "ҩƷ����������������"
            mInt���ݺ� = ���ݺ�.��������
            MStrCaption = "ҩƷ�����������"
        Case ģ���.�������
            Me.Caption = "ҩƷ���������������"
            mInt���ݺ� = ���ݺ�.�������
            MStrCaption = "ҩƷ����������"
        Case ģ���.�⹺���
            Me.Caption = "ҩƷ�⹺�����������"
            mInt���ݺ� = ���ݺ�.�⹺���
            MStrCaption = "ҩƷ�⹺������"
        Case ģ���.ҩƷ����
            Me.Caption = "ҩƷ������������"
            mInt���ݺ� = ���ݺ�.ҩƷ����
            MStrCaption = "ҩƷ���ù���"
        Case 1341
            Me.Caption = "ҩƷ������������"
            mInt���ݺ� = ���ݺ�.ҩƷ�ƿ�
            MStrCaption = "ҩƷ�������"
    End Select
    
End Sub


Private Sub InitVSFColSel()
'-----------------------------------------
'������ѡ����У��Լ����ܽ���ѡ�����
'-----------------------------------------
    Dim rows As Integer
    Dim i As Integer
    Dim sum As Integer
    
    For i = 1 To Me.vsfList.Cols - 1
        If Me.vsfList.ColHidden(i) = False Then
            With vsfColSel
                .rows = .rows + 1
                .TextMatrix(.rows - 2, 1) = Me.vsfList.TextMatrix(0, i)
                .RowData(.rows - 2) = i
            End With
        End If
    Next
    
    sum = 6
    If Me.vsfList.ColHidden(mintcol��������) = False Then sum = 7
    
    Me.vsfColSel.rows = vsfColSel.rows - 1
    For i = 1 To sum
        Me.vsfColSel.Cell(flexcpForeColor, i, 0, i, 1) = CSTCOLOR_FIXED
    Next
End Sub

Private Sub InitComman()
'--------------------------------------
'��ʼ��CommandBars1�ؼ�

'--------------------------------------
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����������ϵͳ�Զ�ʶ��
    End With

    With combars.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With combars
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        .Item(1).Delete
        .Icons = Me.imgicon.Icons
    End With
End Sub


Private Sub InitTool()
'-----------------------------------------------------
'���ù�����
'----------------------------------------------------
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    Set objBar = combars.Add("������1", xtpBarTop)
    objBar.ContextMenuPresent = False '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, MINTBTNFILTER, "����")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        Set objControl = .Add(xtpControlButton, MINTBTNALLWRITEOFF, "ȫ��")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        Set objControl = .Add(xtpControlButton, MINTBTNALLELIMINATE, "ȫ��")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        Set objControl = .Add(xtpControlButton, MINTBTNDEL, "ɾ��")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, MINTBTNWRITEOFF, "����")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        Set objControl = .Add(xtpControlButton, MINTBTNHELP, "����")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, MINTBTNEXIT, "�˳�")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        
        Set objControl = .Add(xtpControlButton, MINTBTNSIMPLE, "���")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
            objControl.Checked = True
        Set objControl = .Add(xtpControlButton, MINTBTNCONPLETE, "����")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
    End With
End Sub
Private Sub InitTask()
'---------------------------------------
'��ʼ���������
'----------------------------------------
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
   
    Call tkpMain.SetMargins(0, 0, 0, 0, 0)
    Call tkpMain.SetItemInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetItemOuterMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupOuterMargins(3, 3, 3, 0)
        
    Set objGroup = tkpMain.Groups.Add(1, "��������")
    objGroup.Expandable = False '��������
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic������Ϣ
    pic������Ϣ.BackColor = objItem.BackColor
   
    Set objGroup = tkpMain.Groups.Add(2, "��������")
    objGroup.Expandable = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic������Ϣ
    pic������Ϣ.BackColor = objItem.BackColor
    objGroup.Expanded = False  'û�д�

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    
    Me.tkpMain.Move 0, 530, Me.tkpMain.Width, Me.ScaleHeight - 530 - Me.staThis.Height
    Me.fraEW.Move Me.tkpMain.Left + Me.tkpMain.Width, Me.tkpMain.Top, 45, Me.tkpMain.Height
    Me.vsfList.Move Me.fraEW.Left + Me.fraEW.Width, Me.fraEW.Top + Me.fraҩƷ������Ϣ.Height + 50, Me.ScaleWidth - (Me.fraEW.Left + Me.fraEW.Width), Me.tkpMain.Height - Me.fraҩƷ������Ϣ.Height
    Me.fraҩƷ������Ϣ.Move vsfList.Left, Me.fraEW.Top - 30, vsfList.Width, Me.fraҩƷ������Ϣ.Height

    fraColSel.Left = Me.tkpMain.Width + Me.tkpMain.Left - fraColSel.Width + 265
    fraColSel.Top = (vsfList.RowHeight(0) - fraColSel.Height) / 2 + 540 + Me.fraҩƷ������Ϣ.Height + 50
    fraColSel.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim str�����п� As String
    Dim i As Integer
    
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        For i = 0 To MINTCOL������ - 1
            str�����п� = str�����п� & vsfList.ColKey(i) & "," & vsfList.ColWidth(i) & "|"
        Next
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Caption, "�����п�", str�����п�)
    End If
    mstr�����п� = ""
    Call ReleaseSelectorRS
End Sub

Private Sub fraEW_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------
'�������͵�����������
'------------------------------------------
    On Error Resume Next
    If Me.tkpMain.Width + x < 1700 Or Me.vsfList.Width - x < 500 Then
        Exit Sub
    End If
    
    If Button = 1 Then
        Me.fraEW.Move Me.fraEW.Left + x, Me.fraEW.Top, Me.fraEW.Width, Me.fraEW.Height
        Me.pic������Ϣ.Move Me.pic������Ϣ.Left, Me.pic������Ϣ.Top, Me.pic������Ϣ.Width + x, Me.pic������Ϣ.Height
        Me.tkpMain.Move Me.tkpMain.Left, Me.tkpMain.Top, Me.tkpMain.Width + x, Me.tkpMain.Height
        Me.vsfList.Move Me.vsfList.Left + x, Me.vsfList.Top, Me.vsfList.Width - x, Me.vsfList.Height
        Me.fraҩƷ������Ϣ.Move Me.fraҩƷ������Ϣ.Left + x, Me.fraҩƷ������Ϣ.Top, Me.fraҩƷ������Ϣ.Width - x, Me.fraҩƷ������Ϣ.Height
        fraColSel.Left = fraColSel.Left + x
        Me.cbo�ⷿ.Width = Me.cbo�ⷿ.Width + x
        Me.txt����NO.Width = Me.txt����NO.Width + x
        Me.txt��ʼNO.Width = Me.txt��ʼNO.Width + x
        Me.txt�����.Width = Me.txt�����.Width + x
        Me.txt������.Width = Me.txt������.Width + x
        Me.txtҩƷ.Width = Me.txtҩƷ.Width + x
        Me.cmdҩƷ.Left = Me.cmdҩƷ.Left + x
        Me.DTP����ʱ��.Width = Me.DTP����ʱ��.Width + x
        Me.DTP��������.Width = Me.DTP��������.Width + x
        Me.DTP��ʼʱ��.Width = Me.DTP��ʼʱ��.Width + x
        Me.DTP��ʼ����.Width = Me.DTP��ʼ����.Width + x
        Me.cboʱ�䷶Χ.Width = Me.cboʱ�䷶Χ.Width + x
        Me.cbo����ʱ��.Width = Me.cbo����ʱ��.Width + x
    End If
End Sub
Private Sub InitCbo()
    '-----------------
    '��ʼ��������
    '-----------------
    With Me.cboʱ�䷶Χ
        .Clear
        .AddItem "һ����"
        .AddItem "������"
        .AddItem "������"
        .AddItem "ָ��ʱ�䷶Χ"
    End With
    
    With Me.cbo����ʱ��
        .Clear
        .AddItem "һ����"
        .AddItem "������"
        .AddItem "������"
        .AddItem "ָ��ʱ�䷶Χ"
    End With
    
    Me.cboʱ�䷶Χ.ListIndex = 0
End Sub
Private Sub initGrid()
'----------------------------------
'��ʼ�������
'----------------------------------
    Dim i As Integer
    Dim arr������
    
    mintcolѡ�� = 0
    mintcolҩƷid = 1
    mintcol�к� = 2
    mIntColNO = 3
    mintcol�������� = 4
    mintcolҩƷ��������� = 5
    mintcol��Ʒ�� = 6
    mintcolҩƷ��Դ = 7
    mintcol����ҩ�� = 8
    mintcolҩ�ۼ��� = 9
    mintcol��� = 10
    mintcol��λ = 11
    mintcol���� = 12
    mintcol�������� = 13
    mintcol���� = 14
    mintcol���� = 15
    mintcolժҪ = 16
    mintcol�������� = 17
    mintcol��Ч���� = 18
    mintcol��׼�ĺ� = 19
    mintcol������� = 20 '���������ֶ�
    mintcol������ = 21
    mintcol�������� = 22
    mintcol����� = 23
    mintcol������� = 24
    mintcol�ɹ��޼� = 25
    mintcol�ɹ��� = 26
    mintcol���� = 27
    mintcol�ɱ��� = 28 '�⹺�����Ϊ�����
    mintcol�ɱ���� = 29 '�⹺�����Ϊ������
    mintcol�ӳ��� = 30
    mintcol�ۼ� = 31
    mintcol�ۼ۽�� = 32
    mintcol��� = 33
    '�⹺�����Ҫ���ֶ�
    mintcol���ۼ� = 34
    mintcol���۵�λ = 35
    mintcol���۽�� = 36
    mintcol���۲�� = 37
    mintcol�⹺��׼�ĺ� = 38
    mintcol������� = 39
    mintcol��Ʊ�� = 40
    mintcol��Ʊ���� = 41
    mintcol��Ʊ��Ϣ = 42
    mintcol��Ʊ��� = 43
    '��Ҫ���ص���
    mintcol��ʵ���� = 44
    mintcol��� = 45
    mintcol����ϵ�� = 46
    mintcol���� = 47
    mintcol��¼״̬ = 48
    mintcol�������� = 49
    mintcol�������� = 50
    mintcol���Ч�� = 51
    mintcolʵ�ʲ�� = 52
    mintcolʵ�ʽ�� = 53
    mintcol�ϴι�Ӧ��ID = 54
    mintcol�Է����� = 55
    mintcolҩ�� = 56
    mintcol�Ƿ��� = 57
    
    With Me.vsfList
        .rows = 1
        .Cols = MINTCOL������
        If mstr�����п� <> "" Then
            arr������ = Split(mstr�����п�, "|")
            If UBound(arr������) <> MINTCOL������ Then
                mstr�����п� = ""
            Else
                For i = 0 To UBound(arr������) - 1
                    SetColValue Split(arr������(i), ",")(0), i, Split(arr������(i), ",")(1)
                Next
            End If
        End If
    
        .TextMatrix(0, mintcolѡ��) = ""
        .TextMatrix(0, mintcolҩƷid) = "ҩƷid"
        .TextMatrix(0, mintcol�к�) = "�к�"
        .TextMatrix(0, mIntColNO) = "NO"
        .TextMatrix(0, mintcolҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mintcol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mintcolҩƷ��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mintcol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mintcolҩ�ۼ���) = "ҩ�ۼ���"
        .TextMatrix(0, mintcol���) = "���"
        .TextMatrix(0, mintcol����) = "����"
        .TextMatrix(0, mintcol��λ) = "��λ"
        .TextMatrix(0, mintcol����) = "����"
        .TextMatrix(0, mintcol��������) = "��������"
        .TextMatrix(0, mintcol��Ч����) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mintcol��������) = "�Է��ⷿ"
        .TextMatrix(0, mintcol������) = "������"
        .TextMatrix(0, mintcol��������) = "��������"
        .TextMatrix(0, mintcol�����) = "�����"
        .TextMatrix(0, mintcol�������) = "�������"
        .TextMatrix(0, mintcol�������) = "���"
        .TextMatrix(0, mintcol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mintcol����) = "����"
        .TextMatrix(0, mintcol��������) = "��������"
        .TextMatrix(0, mintcol�ɹ��޼�) = "�ɹ��޼�"
        .TextMatrix(0, mintcol�ɹ���) = "�ɹ���"
        .TextMatrix(0, mintcol����) = "����"
        .TextMatrix(0, mintcol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mintcol�ɱ����) = "�ɱ����"
        .TextMatrix(0, mintcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mintcol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mintcol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mintcol���) = "���"
        .TextMatrix(0, mintcol���ۼ�) = "���ۼ�"
        .TextMatrix(0, mintcol���۵�λ) = "���۵�λ"
        .TextMatrix(0, mintcol���۽��) = "���۽��"
        .TextMatrix(0, mintcol���۲��) = "���۲��"
        .TextMatrix(0, mintcol�⹺��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mintcol�������) = "�������"
        .TextMatrix(0, mintcol��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, mintcol��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, mintcol��Ʊ��Ϣ) = "��Ʊ��Ϣ"
        .TextMatrix(0, mintcol��Ʊ���) = "��Ʊ���"

        .TextMatrix(0, mintcol��ʵ����) = "��ʵ����"
        .TextMatrix(0, mintcol���) = "���"
        .TextMatrix(0, mintcol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mintcolҩ��) = "ҩ��"
        .TextMatrix(0, mintcol����) = "����"
        .TextMatrix(0, mintcol��¼״̬) = "��¼״̬"
        .TextMatrix(0, mintcol��������) = "��������"
        .TextMatrix(0, mintcol��������) = "��������"
        .TextMatrix(0, mintcol���Ч��) = "���Ч��"
        .TextMatrix(0, mintcolʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mintcolʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mintcol�ϴι�Ӧ��ID) = "�ϴι�Ӧ��ID"
        .TextMatrix(0, mintcolժҪ) = "ժҪ"
        .TextMatrix(0, mintcol�Է�����) = "�Է�����id"
        .TextMatrix(0, mintcol�Ƿ���) = "�Ƿ���"

        .ColKey(mintcolѡ��) = "ѡ��"
        .ColKey(mintcolҩƷid) = "ҩƷid"
        .ColKey(mintcol�к�) = "�к�"
        .ColKey(mIntColNO) = "NO"
        .ColKey(mintcolҩƷ���������) = "ҩƷ���������"
        .ColKey(mintcol��Ʒ��) = "��Ʒ��"
        .ColKey(mintcolҩƷ��Դ) = "ҩƷ��Դ"
        .ColKey(mintcol����ҩ��) = "����ҩ��"
        .ColKey(mintcolҩ�ۼ���) = "ҩ�ۼ���"
        .ColKey(mintcol����) = "����"
        .ColKey(mintcol���) = "���"
        .ColKey(mintcol��λ) = "��λ"
        .ColKey(mintcol����) = "����"
        .ColKey(mintcol��������) = "��������"
        .ColKey(mintcol��Ч����) = "��Ч����"
        .ColKey(mintcol��������) = "��������"
        .ColKey(mintcol�������) = "�������"
        .ColKey(mintcol������) = "������"
        .ColKey(mintcol��������) = "��������"
        .ColKey(mintcol�����) = "�����"
        .ColKey(mintcol�������) = "�������"
        .ColKey(mintcol��׼�ĺ�) = "��׼�ĺ�"
        .ColKey(mintcol����) = "����"
        .ColKey(mintcol��������) = "��������"
        .ColKey(mintcol�ɹ��޼�) = "�ɹ��޼�"
        .ColKey(mintcol�ɹ���) = "�ɹ���"
        .ColKey(mintcol����) = "����"
        .ColKey(mintcol�ɱ���) = "�ɱ���"
        .ColKey(mintcol�ɱ����) = "�ɱ����"
        .ColKey(mintcol�ӳ���) = "�ӳ���"
        .ColKey(mintcol�ۼ�) = "�ۼ�"
        .ColKey(mintcol�ۼ۽��) = "�ۼ۽��"
        .ColKey(mintcol���) = "���"
        .ColKey(mintcol���ۼ�) = "���ۼ�"
        .ColKey(mintcol���۵�λ) = "���۵�λ"
        .ColKey(mintcol���۽��) = "���۽��"
        .ColKey(mintcol���۲��) = "���۲��"
        .ColKey(mintcol�⹺��׼�ĺ�) = "�⹺��׼�ĺ�"
        .ColKey(mintcol�������) = "�������"
        .ColKey(mintcol��Ʊ��) = "��Ʊ��"
        .ColKey(mintcol��Ʊ����) = "��Ʊ����"
        .ColKey(mintcol��Ʊ��Ϣ) = "��Ʊ��Ϣ"
        .ColKey(mintcol��Ʊ���) = "��Ʊ���"
        .ColKey(mintcol��ʵ����) = "��ʵ����"
        .ColKey(mintcol���) = "���"
        .ColKey(mintcol����ϵ��) = "����ϵ��"
        .ColKey(mintcolҩ��) = "ҩ��"
        .ColKey(mintcol����) = "����"
        .ColKey(mintcol��¼״̬) = "��¼״̬"
        .ColKey(mintcol��������) = "��������"
        .ColKey(mintcol��������) = "��������"
        .ColKey(mintcol���Ч��) = "���Ч��"
        .ColKey(mintcolʵ�ʲ��) = "ʵ�ʲ��"
        .ColKey(mintcolʵ�ʽ��) = "ʵ�ʽ��"
        .ColKey(mintcol�ϴι�Ӧ��ID) = "�ϴι�Ӧ��ID"
        .ColKey(mintcolժҪ) = "ժҪ"
        .ColKey(mintcol�Է�����) = "�Է�����"
        .ColKey(mintcol�Ƿ���) = "�Ƿ���"
        .ColKey(mintcol��������) = "��������"

        If mstr�����п� = "" Then
            .ColWidth(mintcolѡ��) = 270
            .ColWidth(mintcol�к�) = 800
            .ColWidth(mIntColNO) = 900
            .ColWidth(mintcolҩƷ���������) = 2500
            .ColWidth(mintcol��Ʒ��) = 1100
            .ColWidth(mintcol���) = 1100
            .ColWidth(mintcol����) = 1100
            .ColWidth(mintcol��λ) = 600
            .ColWidth(mintcol��������) = 1100
            .ColWidth(mintcolժҪ) = 2300
            .ColWidth(mintcol��Ч����) = 1100
            .ColWidth(mintcol��������) = 1100
            .ColWidth(mintcol��׼�ĺ�) = 1100
            .ColWidth(mintcol�������) = 1100
            .ColWidth(mintcol��������) = 1100
            .ColWidth(mintcol�������) = 1100
            .ColWidth(mintcol����) = 1100
            .ColWidth(mintcol��������) = 1200
            .ColWidth(mintcol�ɹ��޼�) = 1100
            .ColWidth(mintcol�ɹ���) = 1100
            .ColWidth(mintcol����) = 1100
            .ColWidth(mintcol�ӳ���) = 1100
            .ColWidth(mintcol�ɱ���) = 1100
            .ColWidth(mintcol�ɱ����) = 1100
            .ColWidth(mintcol�ۼ�) = 1100
            .ColWidth(mintcol�ۼ۽��) = 1100
            .ColWidth(mintcol���) = 1100
            .ColWidth(mintcol����) = 1100
            .ColWidth(mintcol���ۼ�) = 1100
            .ColWidth(mintcol���۵�λ) = 1100
            .ColWidth(mintcol���۽��) = 1100
            .ColWidth(mintcol���۲��) = 1100
            .ColWidth(mintcol�⹺��׼�ĺ�) = 1100
            .ColWidth(mintcol�������) = 1100
            .ColWidth(mintcol��Ʊ��) = 1100
            .ColWidth(mintcol��Ʊ����) = 1100
            .ColWidth(mintcol��Ʊ��Ϣ) = 1100
            .ColWidth(mintcol��Ʊ���) = 1100
        End If
        
        '�Ƿ���ʾ��Ʒ��
        If gintҩƷ������ʾ <> 0 Then
            .ColHidden(mintcol��Ʒ��) = False
        Else
            .ColHidden(mintcol��Ʒ��) = True
        End If
        
        '���������ʾ���кͲ���ʾ����
        If mlngģ��� = ģ���.������� Then
            .ColHidden(mintcol�������) = False
        Else
            .ColHidden(mintcol�������) = True
        End If
        
        If mlngģ��� = ģ���.ҩƷ���� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
            .ColHidden(mintcol��������) = False
        Else
            .ColHidden(mintcol��������) = True
        End If
        'ֻ���⹺����ʾ����
        If mlngģ��� = ģ���.�⹺��� Then
            .ColHidden(mintcol�ɹ��޼�) = False
            .ColHidden(mintcol�ɹ���) = False
            .ColHidden(mintcol����) = False
            .ColHidden(mintcol�ӳ���) = False
            .ColHidden(mintcol���ۼ�) = False
            .ColHidden(mintcol���۵�λ) = False
            .ColHidden(mintcol���۽��) = False
            .ColHidden(mintcol���۲��) = False
            .ColHidden(mintcol�⹺��׼�ĺ�) = False
            .ColHidden(mintcol�������) = False
            .ColHidden(mintcol��Ʊ��) = False
            .ColHidden(mintcol��Ʊ����) = False
            .ColHidden(mintcol��Ʊ��Ϣ) = False
            .ColHidden(mintcol��Ʊ���) = False
            .ColHidden(mintcolҩ�ۼ���) = False
            .ColHidden(mintcol��������) = False
            .ColHidden(mintcol��׼�ĺ�) = True
        Else
            .ColHidden(mintcol�ɹ��޼�) = True
            .ColHidden(mintcol�ɹ���) = True
            .ColHidden(mintcol����) = True
            .ColHidden(mintcol�ӳ���) = True
            .ColHidden(mintcol���ۼ�) = True
            .ColHidden(mintcol���۵�λ) = True
            .ColHidden(mintcol���۽��) = True
            .ColHidden(mintcol���۲��) = True
            .ColHidden(mintcol�⹺��׼�ĺ�) = True
            .ColHidden(mintcol�������) = True
            .ColHidden(mintcol��Ʊ��) = True
            .ColHidden(mintcol��Ʊ����) = True
            .ColHidden(mintcol��Ʊ��Ϣ) = True
            .ColHidden(mintcol��Ʊ���) = True
            .ColHidden(mintcolҩ�ۼ���) = True
            .ColHidden(mintcol��������) = True
            .ColHidden(mintcol��׼�ĺ�) = False
        End If
        
        If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
            .ColHidden(mintcolժҪ) = False
        Else
            .ColHidden(mintcolժҪ) = True
        End If
        
        'ȫ��ҵ�����ص���
        .ColHidden(mintcolҩƷ���������) = True
        .ColHidden(mintcol��Ʒ��) = True
        .ColHidden(mintcol���) = True
        .ColHidden(mintcol��λ) = True
        .ColHidden(mintcol�к�) = True
        .ColHidden(mintcolҩƷid) = True
        .ColHidden(mintcol��ʵ����) = True
        .ColHidden(mintcol���) = True
        .ColHidden(mintcol����ϵ��) = True
        .ColHidden(mintcolҩ��) = True
        .ColHidden(mintcol����) = True
        .ColHidden(mintcol��¼״̬) = True
        .ColHidden(mintcol��������) = True
        .ColHidden(mintcol��������) = True
        .ColHidden(mintcol���Ч��) = True
        .ColHidden(mintcolʵ�ʲ��) = True
        .ColHidden(mintcolʵ�ʽ��) = True
        .ColHidden(mintcol�ϴι�Ӧ��ID) = True
        .ColHidden(mintcol�Է�����) = True
        .ColHidden(mintcol�Ƿ���) = True
        .ColHidden(mintcolҩƷ��Դ) = True
        .ColHidden(mintcolҩ�ۼ���) = True
        .ColHidden(mintcol����ҩ��) = True
        
        '�����ݶ��뷽ʽ
        .ColAlignment(mintcolҩƷid) = flexAlignRightCenter
        .ColAlignment(mIntColNO) = flexAlignLeftCenter
        .ColAlignment(mintcolҩƷ���������) = flexAlignLeftCenter
        .ColAlignment(mintcol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mintcolҩƷ��Դ) = flexAlignLeftCenter
        .ColAlignment(mintcol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mintcolҩ�ۼ���) = flexAlignLeftCenter
        .ColAlignment(mintcol���) = flexAlignLeftCenter
        .ColAlignment(mintcol����) = flexAlignLeftCenter
        .ColAlignment(mintcol��λ) = flexAlignLeftCenter
        .ColAlignment(mintcol����) = flexAlignLeftCenter
        .ColAlignment(mintcol��������) = flexAlignLeftCenter
        .ColAlignment(mintcol��Ч����) = flexAlignLeftCenter
        .ColAlignment(mintcol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mintcol�������) = flexAlignLeftCenter
        .ColAlignment(mintcol����) = flexAlignRightCenter
        .ColAlignment(mintcol��������) = flexAlignRightCenter
        .ColAlignment(mintcol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mintcol�ɱ����) = flexAlignRightCenter
        .ColAlignment(mintcol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mintcol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mintcol���) = flexAlignRightCenter
        .ColAlignment(mintcol���ۼ�) = flexAlignRightCenter
        .ColAlignment(mintcol���۵�λ) = flexAlignLeftCenter
        .ColAlignment(mintcol���۽��) = flexAlignRightCenter
        .ColAlignment(mintcol���۲��) = flexAlignRightCenter
        .ColAlignment(mintcol�⹺��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mintcol�������) = flexAlignLeftCenter
        .ColAlignment(mintcol��Ʊ��) = flexAlignLeftCenter
        .ColAlignment(mintcol��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(mintcol��Ʊ��Ϣ) = flexAlignLeftCenter
        .ColAlignment(mintcol��Ʊ���) = flexAlignLeftCenter
        .ColAlignment(mintcolժҪ) = flexAlignLeftCenter
        
        '�б�����뷽ʽ
        .FixedAlignment(mIntColNO) = flexAlignCenterCenter
        .FixedAlignment(mintcolҩƷ���������) = flexAlignCenterCenter
        .FixedAlignment(mintcol��Ʒ��) = flexAlignCenterCenter
        .FixedAlignment(mintcolҩƷ��Դ) = flexAlignCenterCenter
        .FixedAlignment(mintcol����ҩ��) = flexAlignCenterCenter
        .FixedAlignment(mintcolҩ�ۼ���) = flexAlignCenterCenter
        .FixedAlignment(mintcol���) = flexAlignCenterCenter
        .FixedAlignment(mintcol����) = flexAlignCenterCenter
        .FixedAlignment(mintcol��λ) = flexAlignCenterCenter
        .FixedAlignment(mintcol����) = flexAlignCenterCenter
        .FixedAlignment(mintcol��������) = flexAlignCenterCenter
        .FixedAlignment(mintcol��Ч����) = flexAlignCenterCenter
        .FixedAlignment(mintcol��׼�ĺ�) = flexAlignCenterCenter
        .FixedAlignment(mintcol�������) = flexAlignCenterCenter
        .FixedAlignment(mintcol��������) = flexAlignCenterCenter
        .FixedAlignment(mintcol������) = flexAlignCenterCenter
        .FixedAlignment(mintcol��������) = flexAlignCenterCenter
        .FixedAlignment(mintcol�����) = flexAlignCenterCenter
        .FixedAlignment(mintcol�������) = flexAlignCenterCenter
        .FixedAlignment(mintcol����) = flexAlignCenterCenter
        .FixedAlignment(mintcol��������) = flexAlignCenterCenter
        .FixedAlignment(mintcol�ɱ���) = flexAlignCenterCenter
        .FixedAlignment(mintcol�ɱ����) = flexAlignCenterCenter
        .FixedAlignment(mintcol�ۼ�) = flexAlignCenterCenter
        .FixedAlignment(mintcol�ۼ۽��) = flexAlignCenterCenter
        .FixedAlignment(mintcol���) = flexAlignCenterCenter
        .FixedAlignment(mintcol���ۼ�) = flexAlignCenterCenter
        .FixedAlignment(mintcol���۵�λ) = flexAlignCenterCenter
        .FixedAlignment(mintcol���۽��) = flexAlignCenterCenter
        .FixedAlignment(mintcol���۲��) = flexAlignCenterCenter
        .FixedAlignment(mintcol�⹺��׼�ĺ�) = flexAlignCenterCenter
        .FixedAlignment(mintcol�������) = flexAlignCenterCenter
        .FixedAlignment(mintcol��Ʊ��) = flexAlignCenterCenter
        .FixedAlignment(mintcol��Ʊ����) = flexAlignCenterCenter
        .FixedAlignment(mintcol��Ʊ��Ϣ) = flexAlignCenterCenter
        .FixedAlignment(mintcol��Ʊ���) = flexAlignCenterCenter
        .FixedAlignment(mintcolժҪ) = flexAlignCenterCenter
        
        .RowHeight(0) = 300
        .AllowUserResizing = flexResizeBoth
        .ExplorerBar = flexExSortShowAndMove
        
        .Cell(flexcpForeColor, 0, mintcol��������) = &HFF0000
        .Cell(flexcpForeColor, 0, mintcolժҪ) = &HFF0000
        .Cell(flexcpFontBold, 0, mintcol��������) = True
        .Cell(flexcpFontBold, 0, mintcolժҪ) = True
        
    End With
End Sub
Public Sub showMe(ByVal intģ��� As Integer, Optional ByVal FrmMain As Form, Optional strtock As String, Optional int�ⷿIndex As Integer)
    '�ù���������������򿪸ô���
    '������intģ��ţ�ҵ��ģ���
    'FrmMain:������
    'strtock:�ⷿ�ַ���(��ʽ:���ⷿ����1,�ⷿid|�ⷿ����2,�ⷿid|......��
    'int�ⷿIndex:�������пⷿ�����б��listindex
    Dim i As Integer
    Dim strsql As String
    Dim rsDepend As Recordset
    Dim arr�ⷿ
    
    On Error Resume Next
    Set mfrmMain = FrmMain
    mlngģ��� = intģ���
     
    If strtock <> "" Then
        arr�ⷿ = Split(strtock, "|")
        
        For i = 0 To UBound(arr�ⷿ) - 1
            Me.cbo�ⷿ.AddItem Split(arr�ⷿ(i), ",")(0)
            Me.cbo�ⷿ.ItemData(i) = Split(arr�ⷿ(i), ",")(1)
        Next
    End If
    
    Me.cbo�ⷿ.ListIndex = int�ⷿIndex
    
    Me.Show 1
End Sub
Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.ColHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                If .Top + .Height > Me.ScaleHeight - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub
Private Sub Txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    Select Case mlngģ���
        Case ģ���.�⹺���
            intNO = 21
        Case ģ���.�������
            intNO = 24
        Case ģ���.ҩƷ�ƿ�
            intNO = 26
        Case ģ���.��������
            intNO = 28
        Case ģ���.ҩƷ����
            intNO = 27
    End Select
    
    lng�ⷿid = Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿid)
        End If
        OS.PressKey (vbKeyTab)
    End If
    
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub Txt��ʼNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿid As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    Select Case mlngģ���
        Case ģ���.�⹺���
            intNO = 21
        Case ģ���.�������
            intNO = 24
        Case ģ���.ҩƷ�ƿ�
            intNO = 26
        Case ģ���.��������
            intNO = 28
        Case ģ���.ҩƷ����
            intNO = 27
    End Select
    
    lng�ⷿid = Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNO) < 8 And Len(txt��ʼNO) > 0 Then
            txt��ʼNO.Text = zlCommFun.GetFullNO(txt��ʼNO.Text, intNO, lng�ⷿid)
        End If
        Me.txt����NO.SetFocus
    End If
    
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Function SaveStrike() As Boolean
'-------------------------------------------
'��������Ĺ��̣�����Boolean���͵�ֵ
'-------------------------------------------
    Dim �д�_IN As Integer
    Dim ԭ��¼״̬_IN As Integer
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim ҩƷID_IN As Long
    Dim ��������_IN As Double
    Dim ������_IN As String
    Dim ��������_IN  As String
    Dim ��Ʊ��_IN As String
    Dim ��Ʊ����_In As String
    Dim ��Ʊ����_IN As Date
    Dim ��Ʊ���_IN As Double
    Dim intRow As Integer
    Dim rstemp As New ADODB.Recordset
    Dim i As Integer
    Dim ժҪ_IN As String
    Dim strҩƷid As String
    Dim lastNO As String
    Dim intȫ������ As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    SaveStrike = False
    With Me.vsfList
        If Val(.TextMatrix(1, mintcol����)) = 0 Then
            For intRow = 1 To .rows - 1
                ��������_IN = ��������_IN + zlStr.FormatEx(.TextMatrix(intRow, mintcol��������) * .TextMatrix(intRow, mintcol����ϵ��), gtype_UserSaleDigits.Digit_����, , True)
            Next
        End If
        For intRow = 1 To .rows - 1
            '����������������С����
            If Val(.TextMatrix(intRow, mintcol��������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mintcol����)), Val(.TextMatrix(intRow, mintcol��������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            If mlngģ��� <> ģ���.�������� And mlngģ��� <> ģ���.ҩƷ���� Then
                '�����������Ƿ��㹻����������Ϊ�������ʱ������
                If mint��������ⷿ <> 0 And .TextMatrix(intRow, 1) <> "" Then
                    If Val(.TextMatrix(intRow, mintcol��������)) = Val(.TextMatrix(intRow, mintcol����)) Then
                        intȫ������ = 1
                        If Val(.TextMatrix(1, mintcol����)) <> 0 Then
                            ��������_IN = zlStr.FormatEx(ͬ���γ�������(Val(.TextMatrix(intRow, mintcol����))), gtype_UserSaleDigits.Digit_����, , True) 'Val(.TextMatrix(intRow, mintcol����)) * Val(.TextMatrix(intRow, mintcol����ϵ��))
                        End If
                    Else
                        intȫ������ = 0
                        If Val(.TextMatrix(1, mintcol����)) <> 0 Then
                            ��������_IN = zlStr.FormatEx(ͬ���γ�������(Val(.TextMatrix(intRow, mintcol����))), gtype_UserSaleDigits.Digit_����, , True) 'zlStr.FormatEx(.TextMatrix(intRow, mintcol��������) * .TextMatrix(intRow, mintcol����ϵ��), gtype_UserSaleDigits.Digit_����, , True)
                        End If
                    End If

                    If CheckStrickUsable(mInt���ݺ�, Me.cbo�ⷿ.ItemData(Me.cbo�ⷿ.ListIndex), Val(.TextMatrix(intRow, 1)), .TextMatrix(intRow, mintcolҩ��), _
                        IIf(mlngģ��� = ģ���.�������, 0, (.TextMatrix(intRow, mintcol����))), Val(��������_IN), mint��������ⷿ, Trim(.TextMatrix(intRow, mIntColNO)), Val(.TextMatrix(intRow, mintcol���))) = False Then
                        Exit Function
                    End If
                End If
            End If
        Next
        
        ������_IN = UserInfo.�û�����
        ��������_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")

        On Error GoTo errHandle
        �д�_IN = 0
        '��ҩƷID˳���������
        Call SetSortRecord
        mrecSort.Sort = "ҩƷid,����,���"
        mrecSort.MoveFirst
        
        For i = 1 To mrecSort.RecordCount
            intRow = mrecSort!�к�
            If .TextMatrix(intRow, 1) <> "" And Val(.TextMatrix(intRow, mintcol��������)) <> 0 Then
                NO_IN = .TextMatrix(intRow, mIntColNO)
                If lastNO <> NO_IN Then
                    lastNO = NO_IN
                    �д�_IN = 0
                End If
                
                �д�_IN = �д�_IN + 1
                ҩƷID_IN = .TextMatrix(intRow, 1)
                strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & ҩƷID_IN
                If Val(.TextMatrix(intRow, mintcol��������)) = Val(.TextMatrix(intRow, mintcol����)) Then
                    ��������_IN = Val(.TextMatrix(intRow, mintcol����)) * Val(.TextMatrix(intRow, mintcol����ϵ��))
                Else
                    ��������_IN = zlStr.FormatEx(.TextMatrix(intRow, mintcol��������) * .TextMatrix(intRow, mintcol����ϵ��), gtype_UserSaleDigits.Digit_����, , True)
                End If
                
                ��������_IN = ��������_IN
                ԭ��¼״̬_IN = .TextMatrix(intRow, mintcol��¼״̬)
                ժҪ_IN = .TextMatrix(intRow, mintcolժҪ)
                ���_IN = IIf(mlngģ��� <> ģ���.ҩƷ�ƿ�, .TextMatrix(intRow, mintcol���), Val(.TextMatrix(intRow, mintcol���)) - 1)
                
                If mlngģ��� = ģ���.�⹺��� Then
                    ��Ʊ��_IN = Trim(.TextMatrix(intRow, mintcol��Ʊ��))
                    ��Ʊ����_In = Trim(.TextMatrix(intRow, mintcol��Ʊ����))
                    ��Ʊ���_IN = Val(IIf(.TextMatrix(intRow, mintcol��Ʊ���) = "", "", .TextMatrix(intRow, mintcol��Ʊ���)))
                End If
                
                Select Case mlngģ���
                    Case ģ���.��������
                        gstrSQL = "ZL_ҩƷ��������_STRIKE("
                    Case ģ���.�������
                        gstrSQL = "ZL_ҩƷ�������_STRIKE("
                    Case ģ���.�⹺���
                        gstrSQL = "ZL_ҩƷ�⹺_STRIKE("
                    Case ģ���.ҩƷ����
                        gstrSQL = "ZL_ҩƷ����_STRIKE("
                    Case ģ���.ҩƷ�ƿ�
                        gstrSQL = "ZL_ҩƷ�ƿ�_STRIKE("
                End Select
                
                '�д�
                gstrSQL = gstrSQL & �д�_IN
                'ԭ��¼״̬
                gstrSQL = gstrSQL & "," & ԭ��¼״̬_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '���
                gstrSQL = gstrSQL & "," & ���_IN
                'ҩƷID
                gstrSQL = gstrSQL & "," & ҩƷID_IN
                If mlngģ��� = ģ���.������� Then
                    gstrSQL = gstrSQL & "," & IIf(ժҪ_IN = "", "Null", "'" & ժҪ_IN & "'")
                End If
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '������
                gstrSQL = gstrSQL & ",'" & ������_IN & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & Format(��������_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                
                If mlngģ��� = ģ���.�⹺��� Then
                    '��Ʊ��
                    gstrSQL = gstrSQL & "," & IIf(��Ʊ��_IN = "", "Null", "'" & ��Ʊ��_IN & "'")
                    '��Ʊ���
                    gstrSQL = gstrSQL & "," & ��Ʊ���_IN
                    '�Ƿ�ȫ������
                    gstrSQL = gstrSQL & "," & intȫ������
                    '�Ƿ�������
                    gstrSQL = gstrSQL & "," & 0
                End If
                
                If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
                    'ժҪ
                    gstrSQL = gstrSQL & "," & IIf(ժҪ_IN = "", "Null", "'" & ժҪ_IN & "'")
                End If
                
                If mlngģ��� = ģ���.�⹺��� Then
                    '��Ʊ����
                    gstrSQL = gstrSQL & "," & IIf(��Ʊ����_In = "", "Null", "'" & ��Ʊ����_In & "'")
                End If
                
                If mlngģ��� = ģ���.ҩƷ�ƿ� Then
                    '������ʽ
                    gstrSQL = gstrSQL & ",0"
                End If
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            mrecSort.MoveNext
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
            MsgBox "û��ѡ��һ��ҩƷ����������¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    SaveStrike = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function
Private Sub SetSortRecord()
'------------------------------------------------
'Ҫ������������
'------------------------------------------------
    Dim n As Integer
    
    If Me.vsfList.rows < 2 Then Exit Sub
    If vsfList.TextMatrix(1, 1) = "" Then Exit Sub
    
    Set mrecSort = New ADODB.Recordset
    With mrecSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfList.rows - 1
            If vsfList.TextMatrix(n, 1) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(vsfList.TextMatrix(n, mintcol���)) = 0, n, Val(vsfList.TextMatrix(n, mintcol���)))
                !ҩƷID = Val(vsfList.TextMatrix(n, 0))
                !���� = Val(vsfList.TextMatrix(n, mintcol����))
                
                .Update
            End If
        Next
    End With
End Sub
Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------
'����ҩƷͨ��ѡ������ѡ��ҩƷ
'--------------------------------------------------------------------
    Dim vRect As RECT
    Dim strsql As String
    Dim sngLeft As Single
    Dim sngTop As Single
    
    If KeyCode = 13 Then
        sngLeft = Me.Left + Me.txtҩƷ.Left + Screen.TwipsPerPixelX + 100
        sngTop = Me.Top + Me.Height - Me.ScaleHeight + Me.txtҩƷ.Top + Me.pic������Ϣ.Top + Me.txtҩƷ.Height + 400
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(6, "ҩƷ�⹺������", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
        End If
        
'        Set mrsReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 6, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), , Me.txtҩƷ.Text, sngLeft, sngTop, True, True, False, False, True, 0)
        Set mrsReturn = frmSelector.showMe(Me, 1, 6, UCase(Me.txtҩƷ.Text), sngLeft, sngTop, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), , 0, True, True, True, , False)
        If Not (mrsReturn Is Nothing) Then
            If Not mrsReturn.EOF Then
                Me.txtҩƷ.Text = mrsReturn!ͨ����
                Me.txtҩƷ.Tag = mrsReturn!ҩƷID
            Else
                Me.txtҩƷ.SetFocus
                Me.txtҩƷ.SelStart = 0
                Me.txtҩƷ.SelLength = Len(Me.txtҩƷ.Text)
            End If
        End If
    End If
End Sub
Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�����е����غ���ʾ
    If Me.vsfColSel.TextMatrix(Row, 0) <> 0 Then
        Me.vsfList.ColHidden(Me.vsfColSel.RowData(Row)) = False
    Else
        Me.vsfList.ColHidden(Me.vsfColSel.RowData(Row)) = True
    End If
End Sub
Private Sub vsfColSel_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim sum As Integer
    
    sum = 6
    If Me.vsfList.ColHidden(mintcol��������) = False Then sum = 7
    
    If Row > sum Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfColSel_LostFocus()
    Me.vsfColSel.Visible = False
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '-------------------------------------------------------------
    '�༭��������������۸�
    '-------------------------------------------------------------
    Dim strKey As String
    Dim i As Integer
    Dim dblNum As Double
    Dim count As Integer
    
    If Col <> mintcol�������� Then Exit Sub
    
    If vsfList.TextMatrix(Row, Col) = "" And strKey = "" Then
        vsfList.TextMatrix(Row, mintcol��������) = 0
        vsfList.Cell(flexcpForeColor, Row, mintcol��������) = CSTCOLOR_NOFONT
        vsfList.Cell(flexcpBackColor, Row, 1, Row, MINTCOL������ - 1) = CSTCOLOR_NOMODIFY
        Exit Sub
    End If
    
    If Not IsNumeric(strKey) And strKey <> "" Then
        MsgBox "�Բ��𣬳�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
        vsfList.TextMatrix(Row, mintcol��������) = 0
        vsfList.Cell(flexcpForeColor, Row, mintcol��������) = CSTCOLOR_NOFONT
        vsfList.Cell(flexcpBackColor, Row, 1, Row, MINTCOL������ - 1) = CSTCOLOR_NOMODIFY
        Exit Sub
    End If
    
    If CDbl(Me.vsfList.TextMatrix(Row, mintcol��������)) > CDbl(Me.vsfList.TextMatrix(Row, mintcol����)) Then
        Me.vsfList.TextMatrix(Row, mintcol��������) = Me.vsfList.TextMatrix(Row, mintcol����)
    End If
    
    For i = 1 To Me.vsfList.rows - 1
        If zlStr.Nvl(Me.vsfList.TextMatrix(i, mintcol��������)) <> 0 Then
            count = 1
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
            Exit For
        End If
    Next
    
    If count <> 1 Then
        Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    End If
    
    strKey = vsfList.TextMatrix(Row, mintcol��������)
    If Val(strKey) <> 0 Then
        If InStr(1, strKey, ".") <> 0 Then vsfList.TextMatrix(Row, mintcol��������) = zlStr.FormatEx(vsfList.TextMatrix(Row, mintcol��������), mintNumberDigit, , True)
    Else
        vsfList.TextMatrix(Row, mintcol��������) = 0
    End If
    If Me.vsfList.TextMatrix(Row, Col) <> 0 Then
        With Me.vsfList
            If .TextMatrix(Row, mintcol�ۼ�) <> "" Then
                .TextMatrix(Row, mintcol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(Row, mintcol�ۼ�) * strKey, mintMoneyDigit, , True)
            End If

            .TextMatrix(Row, mintcol�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(Row, mintcol�ɱ���)) * strKey, mintMoneyDigit, , True)
            .TextMatrix(Row, mintcol���) = zlStr.FormatEx(Val(.TextMatrix(Row, mintcol�ۼ۽��)) - Val(.TextMatrix(Row, mintcol�ɱ����)), mintMoneyDigit, , True)
            
            If mlngģ��� = ģ���.�⹺��� And Val(.TextMatrix(.Row, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol����), 0) <> 0 Then
                 .TextMatrix(.Row, mintcol���۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol���ۼ�)) * Val(strKey), mintMoneyDigit, , True)
                 .TextMatrix(.Row, mintcol���۲��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol���۽��)) - Val(.TextMatrix(.Row, mintcol�ɱ����)), mintMoneyDigit, , True)
            End If
            
            .Cell(flexcpForeColor, Row, mintcol��������) = CSTCOLOR_FONT
            .Cell(flexcpBackColor, Row, 1, Row, MINTCOL������ - 1) = CSTCOLOR_MODIFY
        End With
    Else
        With Me.vsfList
            .TextMatrix(Row, mintcol�ۼ۽��) = 0
            .TextMatrix(Row, mintcol���) = 0
            .TextMatrix(Row, mintcol�ɱ����) = 0
            
            If mlngģ��� = ģ���.�⹺��� And Val(.TextMatrix(Row, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(Row, mintcol����), 0) <> 0 Then
                 .TextMatrix(Row, mintcol���۽��) = 0
                 .TextMatrix(Row, mintcol���۲��) = 0
            End If
            
            .Cell(flexcpForeColor, Row, mintcol��������) = CSTCOLOR_NOFONT
            .Cell(flexcpBackColor, Row, 1, Row, MINTCOL������ - 1) = CSTCOLOR_NOMODIFY
        End With
    End If
    If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
        If IIf(vsfList.TextMatrix(Row, mintcol����) = "", 0, vsfList.TextMatrix(Row, mintcol����)) = 0 Then
            dblNum = Val(vsfList.TextMatrix(Row, mintcol��������)) - Val(vsfList.TextMatrix(Row, mintcol��������))
            For i = 1 To Me.vsfList.rows - 1
                vsfList.TextMatrix(i, mintcol��������) = zlStr.FormatEx(dblNum, mintNumberDigit, , True)
            Next
        End If
    End If
    
End Sub
Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '----------------------------------------------------------------------
    '���ƿ��Ա༭����
    'ֻ�г��������п��Ա༭
    '----------------------------------------------------------------------
    If Col = mintcol�������� Or Col = mintcolժҪ Or Row = 0 Then
        Cancel = False
    Else
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub vsfList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = mintcolѡ�� Or Col = mIntColNO Or Col = mintcolҩƷ��������� Or Position = mintcolҩƷ��������� Or Position = mIntColNO Or Position = mintcolѡ�� Then
        Position = Col
    End If
End Sub
Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mintcolѡ�� Or Col = mIntColNO Or Col = mintcolҩƷ��������� Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub
Private Sub vsfList_DblClick()
'--------------------------------------------
'��������ֵ������������
'--------------------------------------------
    Dim strKey As String
    Dim dblNum As Double
    Dim i As Integer
    Dim count As Integer

    If vsfList.Row = 0 Or vsfList.Col <> mintcol���� Then Exit Sub
    
    If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
        If Val(Me.vsfList.TextMatrix(vsfList.Row, mintcol����)) = 0 Then
            Me.vsfList.TextMatrix(vsfList.Row, mintcol��������) = zlStr.FormatEx(Val(Me.vsfList.TextMatrix(vsfList.Row, mintcol��������)) + Val(Me.vsfList.TextMatrix(vsfList.Row, mintcol��������)), mintNumberDigit, , True)
        End If
    End If
    
    If zlStr.Nvl(Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol��������), 0) = 0 Then
        Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol��������) = zlStr.FormatEx(Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol����), mintNumberDigit, , True)
    Else
        Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol��������) = 0
    End If
    
    For i = 1 To Me.vsfList.rows - 1
        If Me.vsfList.TextMatrix(i, mintcol��������) <> 0 Then
            count = 1
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
            Exit For
        End If
    Next
    
    If count <> 1 Then
        Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    End If
    
    If Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol��������) <> 0 Then
        With Me.vsfList
            strKey = .TextMatrix(.Row, mintcol��������)
            If .TextMatrix(.Row, mintcol�ۼ�) <> "" Then
                .TextMatrix(.Row, mintcol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mintcol�ۼ�) * strKey, mintMoneyDigit, , True)
            End If
            
'            .TextMatrix(.Row, mintcol�ɱ���) =Str.FormatEx((Val(.TextMatrix(.Row, mintcol�ۼ۽��)) - Val(.TextMatrix(.Row, mintcol���))) / strkey, mintCostDigit)
            .TextMatrix(.Row, mintcol�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol�ɱ���)) * strKey, mintMoneyDigit, , True)
            .TextMatrix(.Row, mintcol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol�ۼ۽��)) - Val(.TextMatrix(.Row, mintcol�ɱ����)), mintMoneyDigit, , True)
            
            If mlngģ��� = ģ���.�⹺��� And Val(.TextMatrix(.Row, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol����), 0) <> 0 Then
                 .TextMatrix(.Row, mintcol���۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol���ۼ�)) * Val(strKey), mintMoneyDigit, , True)
                 .TextMatrix(.Row, mintcol���۲��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol���۽��)) - Val(.TextMatrix(.Row, mintcol�ɱ����)), mintMoneyDigit, , True)
            End If
            
            .Cell(flexcpForeColor, .Row, mintcol��������) = CSTCOLOR_FONT
            .Cell(flexcpBackColor, .Row, 1, .Row, MINTCOL������ - 1) = CSTCOLOR_MODIFY
        End With
    Else
        With Me.vsfList
            .TextMatrix(.Row, mintcol�ۼ۽��) = 0
            .TextMatrix(.Row, mintcol���) = 0
            .TextMatrix(.Row, mintcol�ɱ����) = 0
            
            If mlngģ��� = ģ���.�⹺��� And Val(.TextMatrix(.Row, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol����), 0) <> 0 Then
                 .TextMatrix(.Row, mintcol���۽��) = 0
                 .TextMatrix(.Row, mintcol���۲��) = 0
            End If
            
            .Cell(flexcpForeColor, .Row, mintcol��������) = CSTCOLOR_NOFONT
            .Cell(flexcpBackColor, .Row, 1, .Row, MINTCOL������ - 1) = CSTCOLOR_NOMODIFY
        End With
    End If
    
    If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
        If IIf(vsfList.TextMatrix(vsfList.Row, mintcol����) = "", 0, vsfList.TextMatrix(vsfList.Row, mintcol����)) = 0 Then
            dblNum = Val(vsfList.TextMatrix(vsfList.Row, mintcol��������)) - Val(vsfList.TextMatrix(vsfList.Row, mintcol��������))
            For i = 1 To Me.vsfList.rows - 1
                vsfList.TextMatrix(i, mintcol��������) = zlStr.FormatEx(dblNum, mintNumberDigit, , True)
            Next
        End If
    End If
End Sub
Private Sub vsfList_EnterCell()
    Dim i As Integer
    
    If Me.vsfList.Row = 0 Then Exit Sub
    Me.staThis.Panels(2).Text = "��ǰ���εĿ��ÿ��Ϊ" & Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol��������) & Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol��λ)
    
    With Me.vsfList
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.imgList.ListImages(1).Picture
            
        '�Ӵֱ༭�еı߿�
        If .MouseCol = mintcol�������� Or .MouseCol = mintcolժҪ Then
            .BackColorSel = CSTCOLOR_ENTERCELL
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
        End If
        
        If .MouseCol = mintcolժҪ Then
            i = .Row
            Do While i >= 1
                If .TextMatrix(i, mintcolժҪ) <> "" Then
                    .TextMatrix(.Row, mintcolժҪ) = .TextMatrix(i, mintcolժҪ)
                    Exit Sub
                End If
                i = i - 1
            Loop
        End If
    End With
End Sub
Private Sub vsfList_GotFocus()
    Me.vsfList.BackColorSel = CSTCOLOR_ENTERCELL
    If Me.vsfList.MouseCol = mintcol�������� Then Me.vsfList.FocusRect = flexFocusSolid
End Sub
Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    '-----------------------------------------------------------
    'ͨ���س���������һ�г��������ı༭,ɾ��ѡ����
    '-----------------------------------------------------------
    Dim strText As String
    Dim count As Integer
    
    With Me.vsfList
        If .Row = 0 Then Exit Sub
        If KeyCode = 46 Then
            .RemoveItem (.Row)
            
            If .rows = 1 Then
                Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
                Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
                Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
                Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
            End If
        End If
        
        If KeyCode = 13 And .Col = mintcolժҪ Then
            If .Row <> .rows - 1 Then
                .Row = .Row + 1
                .Col = mintcol��������
            End If
        ElseIf KeyCode = 13 And .Col = mintcol�������� Then
            If .ColHidden(mintcolժҪ) Then
                If .Row <> .rows - 1 Then
                    .Row = .Row + 1
                    .Col = mintcol��������
                End If
            Else
                .Col = mintcolժҪ
            End If
        End If
        
    End With
End Sub
Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strText As String
    Dim count As Integer
    
    If Col <> mintcol�������� Or Row = 0 Then Exit Sub
    
    If InStr(MCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13) + IIf(Val(vsfList.TextMatrix(Row, mintcol����)) > 0, "", Chr(45)), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc(".") Then
        If InStr(vsfList.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc("-") Then
        If InStr(vsfList.EditText, "-") <> 0 Then     'ֻ�ܴ���һ��-
            KeyAscii = 0
        End If
    End If
    strText = ""
End Sub

Private Sub vsfList_LostFocus()
    Me.vsfList.BackColorSel = CSTCOLOR_LOSTFORCE
    Me.vsfList.FocusRect = flexFocusLight
End Sub

Private Sub vsfList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '-----------------------------------------------------------------
    '��������������У���ʾ˫�����пɸı����������ֵ
    '------------------------------------------------------------
    'Ϊ���������������
    If Me.vsfList.MouseRow <= 0 Then Exit Sub
    
    If Me.vsfList.MouseCol = mintcol���� Then
        If Me.vsfList.TextMatrix(Me.vsfList.MouseRow, mintcol��������) = 0 Then
            Me.vsfList.ToolTipText = "˫�����У����г�����������" & Me.vsfList.TextMatrix(Me.vsfList.MouseRow, mintcol����)
        Else
            Me.vsfList.ToolTipText = "˫�����У����г�����������0"
        End If
    Else
        Me.vsfList.ToolTipText = ""
    End If
End Sub
Private Sub SetSimple(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '---------------------------------------------------------------------
    '���ü���е���ʽ
    '---------------------------------------------------------------------
    Dim i As Integer
    
    For i = mintcol�������� To Me.vsfList.Cols - 1
        If vsfList.ColHidden(i) = False Then
            vsfList.ColData(i) = i
            vsfList.ColHidden(i) = True
        End If
    Next
    
    If Control.Checked = False Then
        Control.Checked = True
        Me.combars.Item(1).Controls.Item(MINTBTNCONPLETE).Checked = False
    End If
End Sub
Private Sub SetConplete(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '---------------------------------------------------------------------
    '���������е���ʽ
    '---------------------------------------------------------------------
    Dim i As Integer
    If Control.Checked = False Then
        For i = mintcol�������� To Me.vsfList.Cols - 1
            If vsfList.ColData(i) Then vsfList.ColHidden(i) = False
        Next
        
        Control.Checked = True
        Me.combars.Item(1).Controls.Item(MINTBTNSIMPLE).Checked = False
    End If
End Sub
Private Sub SetColValue(ByVal str���� As String, ByVal intValue As Integer, ByVal intW As Integer)
'-----------------------------------------------------------------
'�����ø��Ի����õ�ǰ���£����е�˳����еĿ�Ȼָ�Ϊ֮ǰ��״̬
'-----------------------------------------------------------------
    Select Case str����
        Case "ѡ��"
            mintcolѡ�� = intValue
        Case "ҩƷid"
            mintcolҩƷid = intValue
        Case "�к�"
            mintcol�к� = intValue
        Case "NO"
            mIntColNO = intValue
        Case "ҩƷ���������"
            mintcolҩƷ��������� = intValue
        Case "��Ʒ��"
            mintcol��Ʒ�� = intValue
        Case "ҩƷ��Դ"
            mintcolҩƷ��Դ = intValue
        Case "����ҩ��"
            mintcol����ҩ�� = intValue
        Case "ҩ�ۼ���"
            mintcolҩ�ۼ��� = intValue
        Case "���"
            mintcol��� = intValue
        Case "��λ"
            mintcol��λ = intValue
        Case "����"
            mintcol���� = intValue
        Case "��������"
            mintcol�������� = intValue
        Case "����"
            mintcol���� = intValue
        Case "����"
            mintcol���� = intValue
        Case "��������"
            mintcol�������� = intValue
        Case "��Ч����"
            mintcol��Ч���� = intValue
        Case "�������"
            mintcol������� = intValue
        Case "�ɹ��޼�"
            mintcol�ɹ��޼� = intValue
        Case "�ɹ���"
            mintcol�ɹ��� = intValue
        Case "����"
            mintcol���� = intValue
        Case "�ɱ���"
            mintcol�ɱ��� = intValue
        Case "�ɱ����"
            mintcol�ɱ���� = intValue
        Case "�ӳ���"
            mintcol�ӳ��� = intValue
        Case "�ۼ�"
            mintcol�ۼ� = intValue
        Case "�ۼ۽��"
            mintcol�ۼ۽�� = intValue
        Case "���"
            mintcol��� = intValue
        Case "���ۼ�"
            mintcol���ۼ� = intValue
        Case "���۵�λ"
            mintcol���۵�λ = intValue
        Case "���۽��"
            mintcol���۽�� = intValue
        Case "���۲��"
            mintcol���۲�� = intValue
        Case "�⹺��׼�ĺ�"
            mintcol�⹺��׼�ĺ� = intValue
        Case "�������"
            mintcol������� = intValue
        Case "��Ʊ��"
            mintcol��Ʊ�� = intValue
        Case "��Ʊ����"
            mintcol��Ʊ���� = intValue
        Case "��Ʊ��Ϣ"
            mintcol��Ʊ��Ϣ = intValue
        Case "��Ʊ���"
            mintcol��Ʊ��� = intValue
        Case "��ʵ����"
            mintcol��ʵ���� = intValue
        Case "���"
            mintcol��� = intValue
        Case "����ϵ��"
            mintcol����ϵ�� = intValue
        Case "ҩ��"
            mintcolҩ�� = intValue
        Case "����"
            mintcol���� = intValue
        Case "��¼״̬"
            mintcol��¼״̬ = intValue
        Case "��������"
            mintcol�������� = intValue
        Case "��������"
            mintcol�������� = intValue
        Case "���Ч��"
            mintcol���Ч�� = intValue
        Case "ʵ�ʲ��"
            mintcolʵ�ʲ�� = intValue
        Case "ʵ�ʽ��"
            mintcolʵ�ʽ�� = intValue
        Case "�ϴι�Ӧ��ID"
            mintcol�ϴι�Ӧ��ID = intValue
        Case "ժҪ"
            mintcolժҪ = intValue
        Case "�Է�����"
            mintcol�Է����� = intValue
        Case "�Ƿ���"
            mintcol�Ƿ��� = intValue
        Case "������"
            mintcol������ = intValue
        Case "��������"
            mintcol�������� = intValue
        Case "�����"
            mintcol����� = intValue
        Case "�������"
            mintcol������� = intValue
        Case "��������"
            mintcol�������� = intValue
    End Select
    
    vsfList.ColWidth(intValue) = intW
End Sub
Private Sub vsfList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
        If Val(Me.vsfList.TextMatrix(Row, mintcol����)) = 0 Then
            Me.vsfList.TextMatrix(Row, mintcol��������) = Val(Me.vsfList.TextMatrix(Row, mintcol��������)) + Val(Me.vsfList.TextMatrix(Row, mintcol��������))
        End If
    End If
End Sub
Private Sub AllWriteOff()
    Dim i As Integer
    Dim dblOldSum As Double
    Dim dblSum As Double
    Dim strKey As String
    
    With Me.vsfList
        For i = 1 To .rows - 1
            dblOldSum = .TextMatrix(i, mintcol��������) + dblOldSum
            .TextMatrix(i, mintcol��������) = zlStr.FormatEx(.TextMatrix(i, mintcol����), mintNumberDigit, , True)
            
            If i = .Row Then
                .EditText = .TextMatrix(i, mintcol����)
            End If
            
            If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
                If Val(.TextMatrix(i, mintcol����)) = 0 Then
                    dblSum = dblSum + Val(.TextMatrix(i, mintcol��������))
                    .Cell(flexcpText, 1, mintcol��������, .rows - 1, mintcol��������) = zlStr.FormatEx(.TextMatrix(1, mintcol��������) + dblOldSum - dblSum, mintNumberDigit, , True)
                End If
            End If
            
            If .TextMatrix(i, mintcol��������) <> 0 Then
                strKey = .TextMatrix(i, mintcol��������)
                If .TextMatrix(i, mintcol�ۼ�) <> "" Then
                    .TextMatrix(i, mintcol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(i, mintcol�ۼ�) * strKey, mintMoneyDigit, , True)
                End If
                
                .TextMatrix(i, mintcol�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol�ɱ���)) * strKey, mintMoneyDigit, , True)
                .TextMatrix(i, mintcol���) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol�ۼ۽��)) - Val(.TextMatrix(i, mintcol�ɱ����)), mintMoneyDigit, , True)
                
                If mlngģ��� = ģ���.�⹺��� And Val(.TextMatrix(i, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(i, mintcol����), 0) <> 0 Then
                     .TextMatrix(i, mintcol���۽��) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol���ۼ�)) * Val(.TextMatrix(i, mintcol��������)), mintMoneyDigit, , True)
                     .TextMatrix(i, mintcol���۲��) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol���۽��)) - Val(.TextMatrix(i, mintcol�ɱ����)), mintMoneyDigit, , True)
                End If
                
                .Cell(flexcpForeColor, i, mintcol��������) = CSTCOLOR_FONT
                .Cell(flexcpBackColor, i, 1, i, MINTCOL������ - 1) = CSTCOLOR_MODIFY
            End If
        Next
    End With
    
    For i = 1 To Me.vsfList.rows - 1
        If Me.vsfList.TextMatrix(i, mintcol��������) <> 0 Then
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
        End If
    Next
End Sub
Private Sub AllEliminate()
    Dim i As Integer
    Dim dblOldSum As Double
    Dim dblSum As Double
    
    Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
        
    With Me.vsfList
        If .rows <= 1 Then Exit Sub
        For i = 1 To .rows - 1
            If mlngģ��� = ģ���.�⹺��� Or mlngģ��� = ģ���.������� Or mlngģ��� = ģ���.ҩƷ�ƿ� Then
                If Val(.TextMatrix(i, mintcol����)) = 0 Then
                    dblSum = dblSum + Val(.TextMatrix(i, mintcol��������))
                End If
            End If
            
            .TextMatrix(i, mintcol��������) = 0
            .TextMatrix(i, mintcol�ۼ۽��) = 0
            .TextMatrix(i, mintcol���) = 0
            .TextMatrix(i, mintcol�ɱ����) = 0
            If i = .Row Then
                .EditText = 0
            End If
            
            If mlngģ��� = ģ���.�⹺��� And Val(.TextMatrix(.Row, mintcol�Ƿ���)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol����), 0) <> 0 Then
                 .TextMatrix(i, mintcol���۽��) = 0
                 .TextMatrix(i, mintcol���۲��) = 0
            End If
            
            .Cell(flexcpForeColor, i, mintcol��������) = CSTCOLOR_NOFONT
            .Cell(flexcpBackColor, i, 1, i, MINTCOL������ - 1) = CSTCOLOR_NOMODIFY
        Next
        
        If .rows > 1 And Val(.TextMatrix(1, mintcol����)) = 0 Then .Cell(flexcpText, 1, mintcol��������, .rows - 1, mintcol��������) = zlStr.FormatEx(.TextMatrix(1, mintcol��������) + dblSum, mintNumberDigit, , True)
    End With
End Sub
Private Sub Filter()
'-----------------------
'���˲���
'-----------------------
    Dim i As Integer
    
    '����������
    For i = 1 To Me.vsfList.rows - 1
        Me.vsfList.RemoveItem (1)
    Next
    
    Me.lblҩƷ������Ϣ.Caption = "ҩƷ��Ϣ"
    
    Call InitData

    '�����и�
    For i = 1 To Me.vsfList.rows - 1
        Me.vsfList.RowHeight(i) = 300
    Next
End Sub
Private Sub WriteOff()
'-----------------------
'��������
'-----------------------
    Dim i As Integer
    
    With Me.vsfList
        If .rows = 1 Then
            Exit Sub
        End If
    End With

    Call SetSortRecord
    
    If mlngģ��� = ģ���.�⹺��� Then
        If CheckPay = True Then Exit Sub
    End If
    
    If SaveStrike = True Then
        Call combars_Execute(Me.combars.Item(1).Controls.Item(MINTBTNFILTER))
        If Me.vsfList.rows = 1 Then
            '��Ϊû�����ݣ����Բ������ݵİ�ť������
            Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
        End If
        Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    End If
End Sub

Private Function CheckPay() As Boolean
    '����Ƿ�����Ѿ����߲��ָ���ĵ���
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    For n = 1 To vsfList.rows - 1
        If vsfList.TextMatrix(n, mintcol���) <> "" Then
            gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
                " where �շ�id=(Select Id From ҩƷ�շ���¼ Where ����=1 And No=[1] And (Mod(��¼״̬,3)=0 Or ��¼״̬=1) " & _
                " And ���=[2]) "
            Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�������]", vsfList.TextMatrix(n, mIntColNO), Val(vsfList.TextMatrix(n, mintcol���)))

            If rs.EOF Then CheckPay = False: Exit Function

            If rs!������� = 0 Then
                CheckPay = False
            Else
                CheckPay = True
                MsgBox "��" & n & "��ҩƷ�Ѿ�������߲��ָ�����ܳ�����", vbInformation, gstrSysName
                vsfList.Row = n
                vsfList.Col = 2
                Exit Function
            End If
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub DelRow()
'-----------------------
'ɾ���в���
'-----------------------
    Dim count As Integer
    Dim i As Integer
    
    With Me.vsfList
        If .Row = 0 Then Exit Sub
        .RemoveItem (.Row)
        
        For i = 1 To Me.vsfList.rows - 1
            If zlStr.Nvl(Me.vsfList.TextMatrix(i, mintcol��������)) <> 0 Then
                count = 1
                Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
                Exit For
            End If
        Next
    
        If count <> 1 Then
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
        End If
        
        If .rows = 1 Then
            Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
        End If
    End With
End Sub

Private Function ͬ���γ�������(ByVal lng���� As Long) As Double
    '��������ָ���˿ⷿ��ҩƷ
    '��ȡ�б�����ͬ���εĳ���������
    Dim dbl�������� As Double
    Dim intRow As Integer
    
    For intRow = 1 To vsfList.rows - 1
        If lng���� = Val(vsfList.TextMatrix(intRow, mintcol����)) Then
            dbl�������� = dbl�������� + (Val(vsfList.TextMatrix(intRow, mintcol��������)) * Val(vsfList.TextMatrix(intRow, mintcol����ϵ��)))
        End If
    Next
    
    ͬ���γ������� = dbl��������
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

