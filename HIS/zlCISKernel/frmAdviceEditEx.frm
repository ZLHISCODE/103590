VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdviceEditEx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmAdviceEditEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraMethod 
      BackColor       =   &H8000000E&
      Height          =   2175
      Left            =   1680
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdMethodOK 
         Caption         =   "ȷ��"
         Height          =   300
         Left            =   1065
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMethod 
         Height          =   1815
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   2055
         _cx             =   1993543209
         _cy             =   1993542785
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
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAdviceEditEx.frx":000C
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
   Begin VB.PictureBox picSentence 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2880
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   1185
      Begin VB.TextBox txtSentence 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   15
         TabIndex        =   2
         Top             =   30
         Width           =   930
      End
      Begin VB.Image imgSentence 
         Height          =   210
         Left            =   960
         Picture         =   "frmAdviceEditEx.frx":0048
         ToolTipText     =   "�밴 * �ż�ѡ��"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   4
      Left            =   1470
      MousePointer    =   7  'Size N S
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   1155
      MousePointer    =   9  'Size W E
      TabIndex        =   17
      Top             =   2265
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   16
      Top             =   2250
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   495
      TabIndex        =   15
      Top             =   2535
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   495
      MousePointer    =   7  'Size N S
      TabIndex        =   14
      Top             =   2265
      Width           =   615
   End
   Begin VB.OptionButton optMode 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   3090
      TabIndex        =   11
      Top             =   2745
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.OptionButton optMode 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   180
      Index           =   1
      Left            =   2415
      TabIndex        =   10
      Top             =   2745
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.OptionButton optMode 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   1740
      TabIndex        =   9
      Top             =   2745
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "��"
      Height          =   240
      Left            =   2225
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����Ŀ(*)"
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   525
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   3555
      Picture         =   "frmAdviceEditEx.frx":0572
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "ȡ��(Esc)"
      Top             =   1920
      Width           =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4140
      _cx             =   1993546886
      _cy             =   1993542838
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
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceEditEx.frx":0AFC
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
      Begin MSComctlLib.ImageList img16 
         Left            =   1650
         Top             =   975
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":0BF7
               Key             =   "c0"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":1191
               Key             =   "c1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":172B
               Key             =   "o0"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceEditEx.frx":1CC5
               Key             =   "o1"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   240
         Left            =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ComboBox cbo�걾 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   3015
      Picture         =   "frmAdviceEditEx.frx":225F
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "ȷ��(F2)"
      Top             =   1920
      Width           =   450
   End
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   870
      Left            =   135
      TabIndex        =   4
      Top             =   3015
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1535
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAdviceEditEx.frx":27E9
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
   Begin VB.Line lin 
      Index           =   7
      X1              =   2475
      X2              =   3150
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   2475
      X2              =   3150
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   2475
      X2              =   3150
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   2475
      X2              =   3150
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   2475
      X2              =   3150
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   2475
      X2              =   3150
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   2475
      X2              =   3150
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   2475
      X2              =   3150
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label lblAppend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݸ��(��~��������ʾ�ʾ��)"
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   2745
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   105
      TabIndex        =   5
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmAdviceEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Const EM_POSFROMCHAR = &HD6
Private Const EM_EXGETSEL = (&H400 + 52)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
'=============================================================================================================
'��ڲ�����
Private mlngHwnd As Long '���ڶ�λ�Ŀؼ����
Private mint��Ч As Integer
Private mstr�Ա� As String
Private mint�������� As Integer  '1-����,2-סԺ
Private mint������� As Integer '1-����,2-סԺ
Private mbytUseType As Byte      '0=ҽ���´�,1-·����Ŀ��ҽ������,2-���·������Ŀ

'0-������,1-��������,4-������ϣ�5-��Ѫ�����������������Ҫ��д���븽���
Private mintType As Integer

'��:��������ĿID
Private mlng��ĿID As Long


'��/��:���Ӷ�������,����ʱһ��Ϊ��
'      ���="��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
'      ����="����ID1,����ID2,...;����ID",���п���û�и�������������
'      �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾"
Private mstrExtData As String

'��/��:���븽������,����ʱΪ��
'     ��ʽ="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
Private mstrAppend As String

'�룺��δ�����ҽ������¼��ĸ�������,�����µ�Ϊ׼,��������ʱ��ȡʹ��
'     ��ʽ=��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>...
Private mstrAdvItem As String

'�룺��δ�����ҽ���Ѷ�Ӧ¼������(�������ǵ�ǰҽ������Ӧ��)
Private mstrDiagnosis As String


'��:�����������Ҫ,��������ȡ��������
Private mlng����ID As Long
Private mvar����ID As Variant '��ҳID��Һŵ���
Private mintӤ�� As Integer
Private mlng���˿���id As Long

Private mint���� As Integer  '0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Private mblnNew As Boolean  '�ж��Ƿ����¿�������Ŀʱ���룬����Ϊ���¼�ͷ����

'��������ҽ������������λҪ�ض�Ӧ��ֵ���Ա�ƴ��ҽ��������
Private mstr������λ As String


'�룺�жϼ�������Ƿ�ʹ���°�LIS�ļ������ģʽ
Private mblnNewLIS As Boolean

'���ڲ�����
Private mblnOK As Boolean '��

'�������
Private mstrMatchMode As Boolean
Private mint���� As Integer
Private mstrLike As String
Private mbln���� As Boolean
Private mblnFirst As Boolean
Private mblnReturn As Boolean '�Ƿ��˻س�ȷ��
Private mblnNotAddNew As Boolean '�Ƿ���������
Private mbytSize As Byte '�����С 0-С���壨9�ţ���1-�����壨12�ţ�
Private mstr�����ȼ� As String   '����ҽ���������ȼ�
Private mbln�����ּ����� As Boolean   '�Ƿ����������ּ�����
Private mbln������Ȩ���� As Boolean
Private mbln�����ȼ����� As Boolean  '�Ƿ����ò���������ҽʦ�ﵽ�����ȼ��������
Private mrsAppend As ADODB.Recordset
Private mbln��鲿λ As Boolean    '�ü����Ŀ�Ƿ���Ҫ���ò�λ
Private mstr��ѡ��Ŀ As String
Private mbln��ʦվ As Boolean '�Ƿ��Ǽ�ʦվ����
'ģ��Ŷ���
Public Enum Enum_Program_Modual
    pm����ҽ���´� = 1252
    pmסԺҽ���´� = 1253
    pm����ҽ��վ = 1260
    pmסԺҽ��վ = 1261
    pmסԺ��ʿվ = 1262
    pmҽ������վ = 1263
End Enum

Private mblnChangeSel As Boolean
Private mstrPrivs As String             'Ȩ��
Private mfrmParent As Object
Private mobjEmrInterface As Object           '�°没�����븽���ȡ����

Public Function ShowMe(ByVal frmParent As Object, ByVal lngHwnd As Long, ByRef t_Pati As TYPE_PatiInfoEx, ByVal int���� As Integer, _
            ByVal intType As Integer, ByVal bytUseType As Byte, ByVal int��Ч As Integer, ByVal int������� As Integer, Optional ByVal int�������� As Integer, _
            Optional ByVal blnNewLIS As Boolean, Optional ByVal blnNew As Boolean, Optional ByVal lng��Ŀid As Long, Optional ByRef strExtData As String, _
            Optional ByRef strAppend As String, Optional ByVal strAdvItem As String, Optional ByVal strDiagnosis As String, Optional ByRef str������λ As String, Optional ByVal bln��ʦվ As Boolean) As Boolean
'����:
'     frmParent         ������
'     lngHwnd           ���ڶ�λ�Ŀؼ����,�����øô���Ŀؼ�
'     t_Pati            ������Ϣ
'     int����           0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'     intType           0-������,1-��������,4-������ϣ�5-��Ѫ����������Ҫ��д���븽���
'     bytUseType        0=ҽ���´�,1-·����Ŀ��ҽ������,2-���·������Ŀ
'     int��Ч           ��Ҫ�����ҽ����Ч 0-������1-����
'     int�������       ��ҽ��Ҫ����Ĳ������� 1-����������ﲡ�ˣ���첡�ˣ��������˵�) 2-סԺ��ֻ��סԺ���ˣ�
'     int��������       ���øô���Ĺ���վ���� 1-����ҽ������վ 2-סԺҽ������վ
'     blnNewLIS         �жϼ�������Ƿ�ʹ���°�LIS�ļ������ģʽ
'     blnNew            �ж��Ƿ����¿�������Ŀʱ���룬����Ϊ���¼�ͷ���롣 true-�¿�������Ŀʱ���룬 false-���¼�ͷ���루����ֻ��Լ��飬ֻ���°�LIS��ʹ�ã�blnNewLIS=true)��
'     lng��Ŀid         ��������ĿID
'     strAdvItem        ��δ�����ҽ������¼��ĸ�������,�����µ�Ϊ׼,��������ʱ��ȡʹ��
'                       ��ʽ=��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>...
'    strDiagnosis       ��δ�����ҽ���Ѷ�Ӧ¼������(�������ǵ�ǰҽ������Ӧ��)
'���أ�
'     strExtData        ���Ӷ������� , ����ʱһ��Ϊ��
'                       ��� = "��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
'                       ����="����ID1,����ID2,...;����ID",���п���û�и�������������
'                       �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾"
'     strAppend         ����ʱΪ�գ���ʽ="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
'     str������λ       ����ҽ�������ء�������λ��Ҫ�ض�Ӧ��ֵ
'     bln��ʦվ         ��ǰ���÷��Ǽ�ʦ����վ
    Set mfrmParent = frmParent
    mlngHwnd = lngHwnd
    With t_Pati
        mintӤ�� = .intӤ��
        mlng����ID = .lng����ID
        mlng���˿���id = .lng���˿���ID
        mvar����ID = IIF(.str�Һŵ� = "", .lng��ҳID, .str�Һŵ�)
        mstr�Ա� = .str�Ա�
    End With
    mint���� = int����
    mintType = intType
    mbytUseType = bytUseType
    mint��Ч = int��Ч
    mint������� = int�������
    mint�������� = int��������
    mblnNewLIS = blnNewLIS
    mblnNew = blnNew
    mlng��ĿID = lng��Ŀid
    mstrExtData = strExtData
    mstrAppend = strAppend
    mstrAdvItem = strAdvItem
    mstrDiagnosis = strDiagnosis
    mbln��ʦվ = bln��ʦվ
    mblnOK = False
    
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    
    strExtData = mstrExtData
    strAppend = mstrAppend
    str������λ = mstr������λ
    
    
    ShowMe = mblnOK
End Function

Private Sub cbo�걾_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo�걾.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cbo�걾.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo�걾.ListCount > 0 Then lngIdx = 0
        cbo�걾.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdMethodOK_Click()
    Call vsMethod_KeyPress(vbKeyReturn)
End Sub

Private Sub rtfAppend_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtSentence.Tag = "����ҽ��" Then
        If picSentence.Visible = False And KeyCode > 127 Then KeyCode = 0: Call rtfAppend_SelChange
        If txtSentence.Tag = "����ҽ��" And KeyCode = vbKeyBack Then KeyCode = 0: Call rtfAppend_SelChange
    End If
End Sub

Private Sub cmd_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim int�Ա� As Integer, strSQLItem As String, i As Long
    Dim strStock As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strSamples As String, strPrivs As String
    
    If mstr�Ա� Like "*��*" Then
        int�Ա� = 1
    ElseIf mstr�Ա� Like "*Ů*" Then
        int�Ա� = 2
    End If
    
    On Error GoTo errH
    
    If mintType = 1 Then
        '���븽������:���ﲻ�ǵ���Ӧ��,��˲�����
        '"-1*������ID"�ǲ��ſ�������ID������Ϊ�����������շ���
        strSQLItem = _
            " From ������ĿĿ¼ A Where A.���='F' And A.ID<>-1*" & mlng��ĿID & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[4])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                " And A.������� IN([1],3) And Nvl(A.ִ��Ƶ��,0) IN(0,[2]) And Nvl(A.�����Ա�,0) IN(0,[3])"
        
        strSQL = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��ģ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
            " Group by ID,�ϼ�ID,����,����"
        strSQL = strSQL & " Union ALL" & _
            " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ" & _
            strSQLItem & " Order By ĩ��,��ID Desc,����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�, mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õ�������Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ø��������Ѿ���������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set��������(vsExt.Row, rsTmp)
    ElseIf mintType = 4 Then
        '������Ŀ
        With Me.cbo�걾
            For i = 0 To .ListCount - 1
                strSamples = strSamples & ",'" & .List(i) & "'"
            Next
        End With
        If Len(strSamples) > 0 Then
            strSamples = Mid(strSamples, 2)
        Else
            strSamples = "''"
        End If
        
        strSQL = "Select 0 as ĩ��,���� as ID,-Null as �ϼ�ID,����,����,' ' as ��������,' ' As �걾��λ,NULL as �Թܱ��� From ���Ƽ�������" & _
            " Union ALL" & _
            " Select Distinct 1 as ĩ��,''||A.ID as ID,A.�������� as �ϼ�ID,A.����,A.����,A.�������� as ��������,A.�걾��λ,A.�Թܱ��� " & _
            " From ������ĿĿ¼ A,������Ŀ�ο� C,���鱨����Ŀ D " & _
            " Where A.ID=D.������Ŀid(+) And D.������ĿID=C.��Ŀid(+)" & _
            " And A.���='C' " & _
            IIF(mint���� = 2, "", " And Nvl(A.����Ӧ��,0)=1 ") & _
            " And Nvl(A.�����Ա�,0) In (0,[2])" & _
            " And A.������� IN([1],3" & IIF(mint���� = 2, ",4", "") & ") " & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[3])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " And (C.�걾���� In (" & strSamples & ") Or C.�걾���� Is Null)" & _
            " Order By ĩ��,���� "
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint�������, int�Ա�, mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õļ�����Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
'        If rsTmp!�������� = "΢����" And vsExt.Rows > 2 Then
'            If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '��������ֻ�ܿ�һ��΢������Ŀ
'                MsgBox "΢������Ŀֻ�ܵ������룡", vbInformation, gstrSysName
'                Exit Sub
'            End If
'        End If
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ü�����Ŀ�Ѿ�¼�룡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '���������ͣ��Թܱ����Ƿ���ͬ
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 And i <> vsExt.Row Then
                If Not (vsExt.TextMatrix(i, 1) = NVL(rsTmp!��������) _
                    Or vsExt.TextMatrix(i, 1) = "" Or NVL(rsTmp!��������) = "") Then
                    MsgBox "��������ͬ�������͵���Ŀ����������Ŀ�ļ�������Ϊ""" & vsExt.TextMatrix(i, 1) & """��", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Not (vsExt.Cell(flexcpData, i, 1) = CStr(NVL(rsTmp!�Թܱ���)) _
                    Or vsExt.Cell(flexcpData, i, 1) = "" Or NVL(rsTmp!�Թܱ���) = "") Then
                    MsgBox "��������ͬ�Թܱ������Ŀ����������Ŀ�Ĺܱ���Ϊ""" & vsExt.Cell(flexcpData, i, 1) & """��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '���³�ʼ�걾
        If Not InitCombox(rsTmp!ID, NVL(rsTmp!�걾��λ)) Then Exit Sub
        
        Call Set������Ŀ(vsExt.Row, rsTmp)
        If rsTmp("��������") = "΢����" Then
            mblnNotAddNew = False
'            vsExt.Rows = 2
        Else
            mblnNotAddNew = False
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdData_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str�Ա� As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "0,1"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "0,2"
    Else
        str�Ա� = "0"
    End If
    
    If mintType = 1 Then
        '����������Ŀ:���ﲻ�ǵ���Ӧ��,��˲�����
        strSQLItem = " From ������ĿĿ¼ A Where A.���='G'" & _
                " And A.������� IN([2],3) And A.ID<>[1]" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[3])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"

        strSQL = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��������" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
            " Group by ID,�ϼ�ID,����,����"
        strSQL = strSQL & " Union ALL" & _
            " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��������" & _
            strSQLItem & " Order By ĩ��,��ID Desc,����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng��ĿID, mint�������, mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
            End If
            txtData.SetFocus: Exit Sub
        End If
        txtData.Tag = rsTmp!ID
        txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
        cmdData.Tag = txtData.Text
        
        txtData.SetFocus
    ElseIf mintType = 4 Then
        '����걾
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Get��鲿λ����(str��鲿λ As String, str��鷽�� As String)
'���ܣ��ռ�����Ӧ�Ĳ�λ������,��","�ż��
    Dim i As Long
    
    str��鲿λ = "": str��鷽�� = ""
    
    With vsExt
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, 1) = 1 Then
                str��鲿λ = str��鲿λ & "," & .TextMatrix(i, 1)
                If .TextMatrix(i, 2) <> "" Then
                    str��鷽�� = str��鷽�� & "," & .TextMatrix(i, 2)
                End If
            End If
        Next
        str��鲿λ = Mid(str��鲿λ, 2)
        str��鷽�� = Mid(str��鷽��, 2)
    End With
End Sub

Private Function GetMax�����ȼ�(ByVal str������Ŀ As String) As String
'���ܣ�ȡ�õ�ǰҽ�������������
'������str������Ŀ��������ĿID�ã��ָ���lng�����ȼ�������ߵ������ȼ�
    Dim strSQL As String, rsTmp As Recordset
    Dim str�����ȼ� As String
    
    On Error GoTo errH
    strSQL = "Select a.�������� From ��������Ŀ¼ A,������϶��� B Where a.ID=b.����ID And a.���='S' And instr([1], b.����id)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, str������Ŀ)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If decode(rsTmp!�������� & "", "��", 1, "��", 2, "��", 3, "��", 4, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) > _
                decode(str�����ȼ�, "��", 1, "��", 2, "��", 3, "��", 4, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) Then
                str�����ȼ� = rsTmp!�������� & ""
            End If
            rsTmp.MoveNext
        Loop
    End If
    GetMax�����ȼ� = str�����ȼ�
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim str�������� As String
    Dim str��Ա�ȼ� As String
    Dim blnMsg As Boolean
    
    Dim lngBegin As Long, lngEnd As Long
    Dim strAppend As String, strData As String
    
    If mintType = 0 Then '��鲿λ���
        '�ռ���λ�����������
        With vsExt
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, 1) = 1 Then
                    If .TextMatrix(i, 2) = "" Then
                        .Row = i: .ShowCell .Row, .Col
                        MsgBox "û��Ϊ��鲿λ""" & .TextMatrix(i, 1) & """ȷ����鷽����", vbInformation, gstrSysName
                        vsExt.SetFocus: Exit Sub
                    End If
                    
                    strTmp = strTmp & "|" & .TextMatrix(i, 1) & ";" & .TextMatrix(i, 2)
                End If
            Next
            If strTmp = "" And vsExt.Editable <> flexEDNone Then
                MsgBox "������ѡ��һ����鲿λ��", vbInformation, gstrSysName
                vsExt.SetFocus: Exit Sub
            End If
            strTmp = Mid(strTmp, 2) & vbTab & IIF(optMode(0).Value, 0, IIF(optMode(1).Value, 1, 2))
        End With
    ElseIf mintType = 1 Or mintType = 4 Then '����������������Ŀ��������Ŀ���걾
        'ȷ�����������Ŀ
        If mintType = 1 Or mintType = 4 And mblnNewLIS = False Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 Then
                    If vsExt.RowData(i) = mlng��ĿID And mintType = 1 Then
                        MsgBox "������������г�������Ҫ������ͬ��������", vbInformation, gstrSysName
                        vsExt.SetFocus: Exit Sub
                    End If
                    strTmp = strTmp & "," & vsExt.RowData(i)
                End If
            Next
        ElseIf mintType = 4 And mblnNewLIS Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And (Val(vsExt.Cell(flexcpChecked, i, 0)) = 1 Or Val(vsExt.TextMatrix(i, 3)) = 0) Then
                    strTmp = strTmp & IIF(Val(vsExt.TextMatrix(i, 3)) = 1, "|", ",") & vsExt.RowData(i)
                End If
            Next
        End If
        strTmp = Mid(strTmp, 2)
        If mintType = 1 And mbln�����ּ����� Then
            '��������ȼ�
            str��Ա�ȼ� = IIF(mstr�����ȼ� <> "", mstr�����ȼ�, UserInfo.�����ȼ�)
            str�������� = GetMax�����ȼ�(mlng��ĿID & "," & strTmp)
            If decode(str��������, "��", 1, "��", 2, "��", 3, "��", 4, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) > _
                decode(str��Ա�ȼ�, "��", 1, "��", 2, "��", 3, "��", 4, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) Then
                blnMsg = True
            End If
        End If
        If strTmp = "" And mintType = 4 Then
            MsgBox "����Ҫѡ��һ��������Ŀ��", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
        strTmp = strTmp & ";" & IIF(mintType = 4, Me.cbo�걾.Text, IIF(Val(txtData.Tag) = 0, "", Val(txtData.Tag)))
    End If
    
    '��鲢�ռ���������������еĵط� rtfAppend.Find�������ܲ�֧�֣�����Ҫ�� Instr ���ж���
    If rtfAppend.Visible Then
        mrsAppend.MoveFirst
        For i = 1 To mrsAppend.RecordCount
            strData = "": lngBegin = -1: lngEnd = -1
            lngBegin = rtfAppend.Find(mrsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngBegin = -1 Then
                lngBegin = InStr(rtfAppend.Text, mrsAppend!��Ŀ & "��")
                lngBegin = lngBegin - 1
            End If
            If lngBegin <> -1 Then
                lngBegin = lngBegin + Len(mrsAppend!��Ŀ & "��")
                If i = mrsAppend.RecordCount Then
                    lngEnd = Len(rtfAppend.Text)
                Else
                    mrsAppend.MoveNext
                    lngEnd = rtfAppend.Find(vbCrLf & mrsAppend!��Ŀ & "��", lngBegin, , rtfNoHighlight Or rtfMatchCase)
                    If lngEnd = -1 Then
                        lngEnd = InStr(rtfAppend.Text, vbCrLf & mrsAppend!��Ŀ & "��")
                        lngEnd = lngEnd - 1
                    End If
                    If lngEnd = -1 Then
                        lngEnd = InStr(rtfAppend.Text, mrsAppend!��Ŀ & "��")
                        lngEnd = lngEnd - 1
                    End If
                    mrsAppend.MovePrevious
                End If
            End If
            If lngBegin <> -1 And lngEnd <> -1 Then
                'MID��������1Ϊ������rtf����0Ϊ����
                lngBegin = lngBegin + 1
                lngEnd = lngEnd + 1
                strData = Mid(rtfAppend.Text, lngBegin, lngEnd - lngBegin)
                'ȥ��Ϊ��������ı����һ��λ�ò���ֱ��¼�뺺������ӵĿո�
                If Left(strData, 1) = " " Then strData = Mid(strData, 2)
                If Right(strData, 1) = " " Then strData = Left(strData, Len(strData) - 1)
                
                If Trim(strData) = "" And NVL(mrsAppend!����, 0) = 1 Then
                    MsgBox "���ݸ���""" & mrsAppend!��Ŀ & """������û����д��", vbInformation, gstrSysName
                    If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
                        rtfAppend.SelStart = lngBegin
                    Else
                        rtfAppend.SelStart = lngBegin - 1
                    End If
                    rtfAppend.SetFocus: Exit Sub
                ElseIf zlCommFun.ActualLen(strData) > 4000 Then
                    MsgBox "���ݸ���""" & mrsAppend!��Ŀ & """�����ݹ������������2000�����ֻ�4000���ַ���", vbInformation, gstrSysName
                    If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
                        rtfAppend.SelStart = lngBegin
                    Else
                        rtfAppend.SelStart = lngBegin - 1
                    End If
                    If rtfAppend.SelText = " " Then rtfAppend.SelStart = lngBegin
                    rtfAppend.SetFocus: Exit Sub
                End If
            End If
            
            'û���������ݵĸ���Ҳ�����˱���
            strAppend = strAppend & "<Split1>" & mrsAppend!��Ŀ & "<Split2>" & NVL(mrsAppend!����, 0) & "<Split2>" & NVL(mrsAppend!Ҫ��ID) & "<Split2>" & strData
            If mintType = 1 And mrsAppend!������ & "" = "������λ" Then mstr������λ = strData
            mrsAppend.MoveNext
        Next
        strAppend = Mid(strAppend, Len("<Split1>") + 1)
    End If
    
    
    If blnMsg Then
         MsgBox "��ǰ�����ȼ�Ϊ" & str�������� & "������ҽʦ" & IIF(mstr�����ȼ� = "", "���ܿ�չ������", "ֻ�ܿ�չ" & mstr�����ȼ� & "��������"), vbInformation, gstrSysName
    End If
    
    '���������������Ȩ������������ҽʦִ��Ȩ
    If mintType = 1 And mbln������Ȩ���� And mint�������� = 2 Then
        If CheckDocEmpowerEx(mlng��ĿID, strAppend) = False Then
            If Not mbln�����ȼ����� Then
                MsgBox "����ҽ�����߱���������ִ��Ȩ���������´", vbInformation, "������Ȩ����"
                Exit Sub
            Else
                MsgBox "����ҽ�����߱���������ִ��Ȩ��", vbInformation, "������Ȩ����"
            End If
        End If
    End If
    
    
    mstrExtData = strTmp
    mstrAppend = strAppend
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst And vsExt.TabStop And vsExt.Enabled And vsExt.Visible And Not Me.ActiveControl Is vsExt Then
        mblnFirst = False: vsExt.SetFocus '�������Ϊʲô�Զ���λ��rtfAppend����ȥ�ˡ�
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyEscape Then
        If fraMethod.Visible Then
            fraMethod.Visible = False
            vsExt.SetFocus
        ElseIf picSentence.Visible Then
            Call HideWordInput(True) '���شʾ�����
        Else
            Call cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '����������ָ�����������
    End If
End Sub

Private Sub Form_Resize()
    Dim lngAppend As Long
    Dim lngMinRows As Long
    Dim lngRows As Long, i As Long
    Dim lngHeight As Long, lngTotalHeight As Long
    Call HideWordInput(True) '���شʾ�����
    
    On Error Resume Next
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    
    vsExt.Left = fraBorder(3).Width
    vsExt.Top = fraBorder(0).Top + fraBorder(0).Height
    vsExt.Width = Me.ScaleWidth - fraBorder(3).Width * 2
    
    If mbln���� Then
        lngTotalHeight = Me.ScaleHeight - fraBorder(4).Height * 3 - lblAppend.Height * 2 - (cbo�걾.Height + 200) - IIF(vsExt.Visible, vsExt.Top, 0)
        vsExt.Height = lngTotalHeight * 0.618

        fraBorder(4).Left = fraBorder(3).Width
        fraBorder(4).Top = IIF(vsExt.Visible, vsExt.Top, 0) + IIF(vsExt.Visible, vsExt.Height, 0)
        fraBorder(4).Width = Me.ScaleWidth - fraBorder(3).Width * 2
        
        lblAppend.Left = fraBorder(3).Width * 2
        lblAppend.Top = fraBorder(4).Top + fraBorder(4).Height * 2
        
        rtfAppend.Left = fraBorder(3).Width
        rtfAppend.Top = lblAppend.Top + lblAppend.Height + fraBorder(4).Height
        rtfAppend.Width = Me.ScaleWidth - fraBorder(3).Width * 2
        rtfAppend.Height = Me.ScaleHeight - rtfAppend.Top - fraBorder(2).Height - (cbo�걾.Height + 200)
        
        lngAppend = rtfAppend.Top + rtfAppend.Height - fraBorder(4).Top
    Else
        vsExt.Height = Me.ScaleHeight - fraBorder(2).Height * 2 - (cbo�걾.Height + 200)
    End If
    
    cbo�걾.Top = Me.ScaleHeight - fraBorder(2).Height - ((Me.ScaleHeight - fraBorder(0).Height * 2 - IIF(vsExt.Visible, vsExt.Height, 0) - lngAppend) - cbo�걾.Height) / 2 - cbo�걾.Height
    txtData.Top = cbo�걾.Top
    lblData.Top = cbo�걾.Top + (cbo�걾.Height - lblData.Height) / 2
    cmdOK.Top = cbo�걾.Top + (cbo�걾.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
        
    optMode(0).Top = cbo�걾.Top + (cbo�걾.Height - optMode(0).Height) / 2
    optMode(1).Top = optMode(0).Top: optMode(2).Top = optMode(0).Top
    optMode(0).Left = 500
    optMode(1).Left = optMode(0).Left + optMode(0).Width + 100
    optMode(2).Left = optMode(1).Left + optMode(1).Width + 100

    lblData.Left = 200
    cbo�걾.Left = lblData.Left + lblData.Width + fraBorder(3).Width
    txtData.Left = cbo�걾.Left
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdCancel.Height
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - fraBorder(1).Width * 3
        
    cbo�걾.Width = cmdOK.Left - cbo�걾.Left - 200

    txtData.Width = cbo�걾.Width
    cmdData.Top = txtData.Top + 30
    cmdData.Left = txtData.Left + txtData.Width - cmdData.Width - 45

    Me.Refresh
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    Dim str���� As String, i As Long, lngBaseHeight As Long
    
    Me.Height = 2325
    
    '�߿�����
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
    
    '����ƥ��
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    If mint������� = 0 Then mint������� = 2 'ȱʡΪסԺ
    mblnOK = False
    mblnNotAddNew = False
    mblnFirst = True
    mstr������λ = ""
    mbln�����ּ����� = False
    If mintType = 1 Then
        '�Ƿ����������ּ�����
        mbln�����ּ����� = Val(zlDatabase.GetPara(209, glngSys)) <> 0
        '�Ƿ�����������ҽʦ��Ȩ����
        mbln������Ȩ���� = Val(zlDatabase.GetPara(217, glngSys)) <> 0
        '�Ƿ����ò���������ҽʦ�ﵽ�����ȼ��������
        mbln�����ȼ����� = Val(zlDatabase.GetPara(254, glngSys)) <> 0
        '�����������Ȩ�����������ּ�����ļ�鲻ʹ��,����Ҳ��ʹ��
        If mbln������Ȩ���� Or mint�������� = 1 Then mbln�����ּ����� = False
    End If
    If mint���� = 0 Then
        If mint�������� = 1 Then
            mbytSize = zlDatabase.GetPara("����", glngSys, pm����ҽ��վ, "0")
        Else
            mbytSize = zlDatabase.GetPara("����", glngSys, pmסԺҽ��վ, "0")
        End If
    ElseIf mint���� = 1 Then
        mbytSize = zlDatabase.GetPara("����", glngSys, pmסԺ��ʿվ, "0")
    Else
        mbytSize = zlDatabase.GetPara("����", glngSys, pmҽ������վ, "0")
    End If
    '�������ã���Ϊ RichTextBox�������ԣ��������÷ŵ�ǰ����е���
    Call SetControlFontSize(Me, mbytSize)
    '��ȡ����
    Call Init���븽��
    mbln���� = Not mrsAppend Is Nothing
    If mbln���� Then mbln���� = mrsAppend.State = 1
    If mbln���� Then mbln���� = mrsAppend.RecordCount > 0
    
    '��ʼ�������ʽ
    If mintType = 0 Then
        If Not Init������ Then Unload Me: Exit Sub
    ElseIf mintType = 1 Then
        lblData.Visible = True
        txtData.Visible = True
        cmdData.Visible = True
        lblData.Caption = "����"
        If Not Init������Ŀ Then Unload Me: Exit Sub
    ElseIf mintType = 4 Then
        lblData.Visible = True
        lblData.Caption = "�걾"
        With cbo�걾
            .Left = txtData.Left: .Top = txtData.Top: .Width = txtData.Width
            .Visible = True
        End With
        If Not Init������� Then Unload Me: Exit Sub
        If Not InitCombox(DefaultValue:=Me.txtData) Then Unload Me: Exit Sub
    ElseIf mintType = 5 Then
        vsExt.Visible = False
        fraBorder(3).Visible = False
        fraBorder(4).Visible = False
        Me.Height = Me.Height - 800
    End If
    If mbln���� Then
        Me.Height = Me.Height + (lblAppend.Height + rtfAppend.Height + fraBorder(4).Height * 3)
    End If
    If mbytUseType = 1 Then
        vsExt.Editable = flexEDNone
        If mintType = 4 Then cmd.Enabled = False  '121475
         '�����޸ļ�鲿λ������ mbln��鲿λ-�����Ŀ��Ч�Լ��ʱ�����õ�vsExt.Editable����,���ڲ�Ҫ�����ò�λ�ļ����Ŀ,Ӧ�ý�ֹ��༭
        If mintType = 0 And mbln��鲿λ Then vsExt.Editable = flexEDKbdMouse
        txtData.Enabled = False
        cmdData.Visible = False
        cbo�걾.Enabled = False
    End If
    
    '��������
    If mintType = 0 Then
        If vsExt.Rows = vsExt.FixedRows + 1 Then
            If vsExt.Editable = flexEDNone Then
                'û�����ò�λʱ���������Ҫȷ���������У�Ҳ����Ҫ���븽����Զ�ȷ��
                If Not mbln���� And Not optMode(0).Enabled Then Call cmdOK_Click: Exit Sub
            ElseIf vsExt.TextMatrix(vsExt.FixedRows, 1) <> "" Then
                'ֻ��һ����λ���Ҳ�λֻ��һ��������ѡʱ���Զ�ȷ��
                
                'ֻ��һ����λ���Զ�ѡ�иò�λ
                vsExt.Cell(flexcpData, vsExt.FixedRows, 1) = 1
                Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1) = img16.ListImages("c1").Picture
                '���û��Ĭ�Ϸ�����ֻ��һ������Ҳѡ��
                str���� = GetOnlyOneMethod(vsExt.Cell(flexcpData, vsExt.FixedRows, 2))
                If vsExt.TextMatrix(vsExt.FixedRows, 2) = "" And str���� <> "" Then
                    vsExt.TextMatrix(vsExt.FixedRows, 2) = str����
                End If
                If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" Then vsExt.TabStop = False
                
                'ֻ��һ��������ѡʱ���������Ҫ�������븽������Ҳ������
                If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" And str���� <> "" Then
                    If Not mbln���� Then Call cmdOK_Click: Exit Sub
                End If
            End If
        End If
    ElseIf mintType = 4 Then
        '������������⴦��
        If Not mbln��ʦվ Then
            blnMulti = Val(zlDatabase.GetPara(84, glngSys)) = 1 '�Ƿ�����һ��ҽ��������������Ŀ
            If Len(Trim(mstrExtData)) > 0 Then
                If Len(Trim(Split(mstrExtData, ";")(0))) > 0 And Not blnMulti Then
                    vsExt.Enabled = False
                    '���ֻ��һ���걾����ʾ������
                    If cbo�걾.ListCount < 2 And Not mbln���� Then cmdOK_Click: Exit Sub
                End If
            End If
        End If
    End If
    
    Call Grid.SetFontSize(vsExt, IIF(mbytSize = 0, 9, 12))
    
    '�ָ����Ի�
    lngBaseHeight = Me.Height
    Call RestoreWinState(Me, App.ProductName, mintType)
    
     '10.26.80���ӹ����ʾ����ǰ���Ի�����ĸ߶ȿ��ܲ���
    If Me.Height < lngBaseHeight Then
        Me.Height = lngBaseHeight
    End If
    
    '���嶨λ
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    Call Form_Resize
    
End Sub

Private Function Init������Ŀ() As Boolean
'���ܣ���ʼ����������ʽ������
'������mstrExtData=��������������������Ŀ����Ϣ,���п���û�и���������Ϊ��ʱ��ʾ������������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng����ID As Long
    Dim arr����IDs As Variant, str����IDs As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = ";"
    str����IDs = CStr(Split(strSQL, ";")(0))
    lng����ID = Val(Split(strSQL, ";")(1))
    
    '��������
    If str����IDs <> "" Then
        strSQL = "Select /*+ Rule*/ A.ID,A.����,A.����,A.��������" & _
            " From ������ĿĿ¼ A" & _
            " Where A.���='F' And A.ID IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[2])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " Order by A.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs, mlng���˿���id)
        i = rsTmp.RecordCount
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "��������"
    vsExt.TextMatrix(0, 1) = "��ģ"
    vsExt.ColWidth(0) = 3200: vsExt.ColWidth(1) = 800
    vsExt.FixedAlignment(0) = 4: vsExt.FixedAlignment(1) = 4
    vsExt.ColAlignment(0) = 1: vsExt.ColAlignment(1) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If str����IDs <> "" And i <> 0 Then
        arr����IDs = Split(str����IDs, ",") '����ԭ������˳��
        For i = 0 To UBound(arr����IDs)
            rsTmp.Filter = "ID=" & CStr(arr����IDs(i))
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!���� & "]" & rsTmp!����
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '���ڻָ���ʾ
                vsExt.TextMatrix(j, 1) = NVL(rsTmp!��������, 0)
            End If
        Next
    End If
    
    '������Ŀ
    If lng����ID <> 0 Then
        strSQL = "Select A.ID,A.����,A.����,�������� From ������ĿĿ¼ A Where A.���='G' And A.ID=[1]" & _
                " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[2])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, mlng���˿���id)
        If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
        If Not rsTmp.EOF Then
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
            cmdData.Tag = txtData.Text '���ڻָ���ʾ
        End If
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init������Ŀ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init������() As Boolean
'���ܣ���ʼ����鲿λ����ʽ������
'������mstrExtData=������鲿λ����Ϣ,Ϊ��ʱ��ʾ�������������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long, i As Integer
    Dim str���� As String, str���� As String
    Dim arrData As Variant, strNoneRegion As String
    Dim blnNone As Boolean
    Dim Y As Long, str���� As String
    
    On Error GoTo errH
    
    '��ȡ�����Ŀ������Ϣ
    strSQL = "Select ����,��������,ִ�б�� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID)
    If mint������� = 2 And NVL(rsTmp!ִ�б��, 0) = 1 Then
        '��������ִ�б��
        optMode(0).Visible = True: optMode(1).Visible = True: optMode(2).Visible = True
        optMode(0).Enabled = True: optMode(1).Enabled = True: optMode(2).Enabled = True
        If UBound(Split(mstrExtData, vbTab)) >= 1 Then
            optMode(Val(Split(mstrExtData, vbTab)(1))).Value = True
        End If
    End If
    str���� = rsTmp!��������
    str���� = rsTmp!����
        
    '��ȡ��鲿λ��Ϣ
    strSQL = "Select B.����,A.��λ,A.����,A.Ĭ��,B.��ע,B.���� as ��鷽�� From ������Ŀ��λ A,���Ƽ�鲿λ B" & _
        " Where A.����=B.���� And A.��λ=B.���� And A.��ĿID=[1] And A.����=[2] Order by B.����,B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID, str����)
    blnNone = rsTmp.EOF
    mbln��鲿λ = Not blnNone
'    If rsTmp.EOF Then
'        '����ü����Ŀ��û�����ü�鲿λ,�������еĹ�ѡ��
'        strSQL = "Select ����,���� as ��λ,Null as ����,Null as Ĭ��,��ע,���� as ��鷽�� From ���Ƽ�鲿λ Where ����=[1] Order by ����,����"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
'        If rsTmp.EOF Then
'            MsgBox "����Ŀ�ļ������""" & str���� & """����û�������κμ�鲿λ�����ȵ���鲿λ�����н������á�", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    With vsExt
        '��ʾ��׼�Ĳ�λ��Ĭ�Ϸ���
        If blnNone Then
            .HighLight = flexHighlightNever
            .Editable = flexEDNone
            .TabStop = False
        Else
            .HighLight = flexHighlightAlways
            .Editable = flexEDKbdMouse
        End If
        .WordWrap = True
        .FocusRect = flexFocusNone
        .BackColorSel = &HFFCC99
        .ForeColorSel = &H0&
        .FixedRows = 1: .FixedCols = 0
        .Rows = .FixedRows + 1: .Cols = 4
        .MergeCellsFixed = flexMergeFree: .MergeRow(0) = True
        .MergeCells = flexMergeFree: .MergeCol(0) = True
        
        If str���� = "����" Then
            .TextMatrix(0, 0) = "�걾����"
            .TextMatrix(0, 1) = "�걾����"
            .TextMatrix(0, 2) = "�������"
        Else
            .TextMatrix(0, 0) = "��鲿λ"
            .TextMatrix(0, 1) = "��鲿λ"
            .TextMatrix(0, 2) = "��鷽��"
        End If
        
        .TextMatrix(0, 3) = "��ע"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        Do While Not rsTmp.EOF
            If .TextMatrix(.Rows - 1, 1) <> rsTmp!��λ Then
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(.Rows - 1, 0) = zlCommFun.GetNeedName("" & rsTmp!����)
                .TextMatrix(.Rows - 1, 1) = rsTmp!��λ
                Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                .Cell(flexcpData, .Rows - 1, 2) = CStr(NVL(rsTmp!��鷽��)) '������ѡ����ʹ��
                .TextMatrix(.Rows - 1, 3) = NVL(rsTmp!��ע)
            End If
            If NVL(rsTmp!Ĭ��, 0) = 1 Then '��"������1,������2,..."�ķ�ʽ��ʾ��λ��鷽��
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & NVL(rsTmp!����)
                If Left(.TextMatrix(.Rows - 1, 2), 1) = "," Then
                    .TextMatrix(.Rows - 1, 2) = Mid(.TextMatrix(.Rows - 1, 2), 2)
                End If
            End If
            rsTmp.MoveNext
        Loop
        
        '�޸�ʱ�������е�����
        '  ���Ϊ�գ�Ҳ��������ǰ�ĵ���λ�����Ŀ����ʱҪ�������ķ�ʽ����ѡ��λ
        '  ���߶�����ǰ�ĵ���λ��Ŀ��ǿ�д�����ǰ�Ĳ�λ(û�з���)���ֻ�������ͬ����λ
        If mstrExtData <> "" Then
            arrData = Split(Split(mstrExtData, vbTab)(0), "|")
            For i = 0 To UBound(arrData)
                lngIdx = .FindRow(CStr(Split(arrData(i), ";")(0)), 1, 1, , True)
                str���� = ""
                If lngIdx <> -1 Then
                    '��鷽����û�в����ڵ�
                    For Y = 0 To UBound(Split(Split(arrData(i), ";")(1), ","))
                        If InStr(.Cell(flexcpData, lngIdx, 2), CStr(Split(Split(arrData(i), ";")(1), ",")(Y))) = 0 Then
                            strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0) & "(" & Split(Split(arrData(i), ";")(1), ",")(Y) & ")"
                        Else
                            str���� = str���� & "," & Split(Split(arrData(i), ";")(1), ",")(Y)
                        End If
                    Next
                    '�ò�λ�ķ���:������ǰ������ֻ�в�λû�з���
                    If UBound(Split(arrData(i), ";")) >= 1 Then
                        .TextMatrix(lngIdx, 2) = Mid(str����, 2)
                    Else
                        .TextMatrix(lngIdx, 2) = ""
                    End If
                    .Cell(flexcpData, lngIdx, 1) = 1 '�����ò�λ��ѡ��
                    Set .Cell(flexcpPicture, lngIdx, 1) = img16.ListImages("c1").Picture
                Else
                    '�ò�λ�����Ѳ�����
                    strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0)
                End If
            Next
        End If
        
        .Row = 1: .Col = 1
        .ShowCell .Row, .Col
        
        'ȷ�����ߴ�
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) < 500 Then .ColWidth(0) = 500
        If .ColWidth(0) > 850 Then .ColWidth(0) = 850
        If .ColWidth(1) < 800 Then .ColWidth(1) = 800
        If .ColWidth(1) > 1600 Then .ColWidth(1) = 1600
        If .ColWidth(2) < 2500 Then .ColWidth(2) = 2500
        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        If .ColWidth(3) < 800 Then .ColWidth(3) = 800
        If .ColWidth(3) > 2000 Then .ColWidth(3) = 2000
        
        lngIdx = 0
        For i = 0 To .Cols - 1
            lngIdx = lngIdx + .ColWidth(i) + 15
        Next
        Me.Width = lngIdx + 90
        
        .Height = (.Rows - 1) * (.RowHeightMin + 15) + .RowHeight(0) + 60
        If Not blnNone Then
            If .Height < 1590 Then .Height = 1590 '����5�в�λ
            If .Height > 2865 + 50 Then .Height = 2865 + 50 '���10�в�λ
        End If
    End With
    
    Me.Height = (vsExt.Height + 90) + cmdOK.Height + (cmdOK.Height * 0.65)
    
    '�Ѳ����ڵĲ�λ��ʾ
    If strNoneRegion <> "" Then
        If str���� = "����" Then
            MsgBox "���²���걾����Ŀ�������Ѳ����ڣ�" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        Else
            MsgBox "���¼�鲿λ�򷽷�����Ŀ�������Ѳ����ڻ����ã�" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        End If
    End If
    
    Init������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Init���븽��() As Boolean
'���ܣ���ȡ��Ŀ�ĵ������븽��
'���أ���Ӧ�ĵ��ݶ��������븽��ʱ����True
    Dim strSQL As String, lngIdx As Long
    Dim arrData As Variant, strData As String
    Dim strNoneAppend As String, strHaveAppend As String
    Dim arrSub As Variant, i As Long
    Dim str����ҽ�� As String
    Dim str����ҽ������ As String
    Dim blnHave����ҽ�� As Boolean
    Dim rsTmp As ADODB.Recordset
    
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    strSQL = "Select C.��Ŀ,C.����,C.Ҫ��ID,C.����,d.������,decode(D.��ʾ��,4,D.��ֵ��,NULL) as ��ֵ��,c.ֻ��" & _
        " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C,����������Ŀ D" & _
        " Where A.������ĿID=[1] And A.Ӧ�ó���=[2]" & _
        " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID And c.Ҫ��id=d.id(+)" & _
        " Order by C.����"
    
    On Error GoTo errH
    Set mrsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID, mint�������)
    If Not mrsAppend.EOF Then
    
        mrsAppend.Filter = "������='����ҽ��'"
        If mrsAppend.RecordCount > 0 Then
            blnHave����ҽ�� = True
            str����ҽ�� = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mrsAppend!������ & "")
            str����ҽ�� = Trim(str����ҽ��)
        End If
        
        mrsAppend.Filter = "������='����ҽ������'"
        If mrsAppend.RecordCount > 0 Then
            If str����ҽ�� <> "" And blnHave����ҽ�� Then
                strSQL = "Select a.����, c.ȱʡ From ���ű� A, ��Ա�� B, ������Ա C, ��������˵�� D" & _
                    " Where a.Id = c.����id And b.Id = c.��Աid And a.Id = d.����id And d.�������� = '�ٴ�' And" & _
                    " (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And" & _
                    " (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null) And b.���� = [1]"
                    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ҽ��)
                
                If Not rsTmp.EOF Then
                    str����ҽ������ = rsTmp!���� & ""
                    rsTmp.Filter = "ȱʡ = 1"
                    If rsTmp.RecordCount > 0 Then str����ҽ������ = rsTmp!���� & ""
                End If
            End If
        End If
        
        mrsAppend.Filter = 0
        arrData = Split(mstrAppend, "<Split1>")
        With rtfAppend
            Do While Not mrsAppend.EOF
                'ȷ����������
                strData = ""
                If mrsAppend!��ֵ�� & "" <> "" Then mstr��ѡ��Ŀ = mstr��ѡ��Ŀ & "," & mrsAppend!������
                If mstrAppend <> "" Then
                    '�޸�ʱ������ԭ������
                    For i = 0 To UBound(arrData)
                        arrSub = Split(arrData(i), "<Split2>")
                        If arrSub(0) = mrsAppend!��Ŀ Then
                            strData = arrSub(3)
                            If strData = "" And UBound(arrSub) >= 4 Then
                                '���Ը��ƻ���ײ�����ҽ�������޸�ʱ�����븽��ҲҪȡȱʡֵ
                                If Val(arrSub(4)) = 1 Then
                                    If Not IsNull(mrsAppend!����) Then
                                        strData = mrsAppend!����
                                    ElseIf mlng����ID <> 0 Then
                                        strData = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mrsAppend!������ & "")
                                    End If
                                End If
                            End If
                            
                            '���ڵĸ���
                            strHaveAppend = strHaveAppend & "," & arrSub(0)
                            strNoneAppend = Replace(strNoneAppend & ",", "," & arrSub(0) & ",", ",")
                            If Right(strNoneAppend, 1) = "," Then strNoneAppend = Left(strNoneAppend, Len(strNoneAppend) - 1)
                        ElseIf InStr(strNoneAppend & ",", "," & arrSub(0) & ",") = 0 _
                             And InStr(strHaveAppend & ",", "," & arrSub(0) & ",") = 0 Then
                            strNoneAppend = strNoneAppend & "," & arrSub(0) '�ȼǵ�û�еĸ�����
                        End If
                    Next
                Else
                    '����ʱ��ʹ��Ԥ�������ݻ�Ӳ�����������ȡ
                    If Not IsNull(mrsAppend!����) Then
                        strData = mrsAppend!����
                    ElseIf mlng����ID <> 0 Then
                        If mrsAppend!������ & "" = "����ҽ��" Then
                            strData = str����ҽ��
                        ElseIf mrsAppend!������ & "" = "����ҽ������" And blnHave����ҽ�� And str����ҽ�� <> "" Then
                            strData = str����ҽ������
                        Else
                            strData = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mrsAppend!������ & "")
                        End If
                    End If
                End If
                
                '��������ʾ��RTF��:�����ı����һ��λ�ò���ֱ��¼�뺺��,����ȶ��һ���������Ŀո�
                .SelText = IIF(.Text = "", "", vbCrLf) & mrsAppend!��Ŀ & "�� " & strData
                lngIdx = .Find(mrsAppend!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
                '���������ҽ�������ȡ�������Ǽ�
                If mrsAppend!������ & "" = "����ҽ��" And mbln�����ּ����� Then
                    mstr�����ȼ� = GetDoctorLevel(strData)
                End If
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(mrsAppend!��Ŀ & "��")
                    .SelBold = True
                    .SelIndent = 100
                    .SelProtected = True
                End If
                If Val(mrsAppend!ֻ�� & "") = 1 And strData <> "" Then
                    lngIdx = lngIdx + Len(mrsAppend!��Ŀ & "�� ")
                    .SelStart = lngIdx
                    .SelLength = Len(strData)
                    .SelProtected = True
                End If
                .SelStart = Len(.Text)
                
                mrsAppend.MoveNext
            Loop
            mstr��ѡ��Ŀ = Mid(mstr��ѡ��Ŀ, 2)
            
            '��궨λ�ڵ�һ�����븽��
            mrsAppend.MoveFirst
            lngIdx = .Find(mrsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(mrsAppend!��Ŀ & "��") + 1
            
            'ȷ��RTF�ؼ��ߴ�
            .Height = (mrsAppend.RecordCount + 2) * 250 + 30
            If .Height < 3 * 265 + 30 Then .Height = 3 * 250 + 30 '����3��
            If .Height > 8 * 265 + 30 Then .Height = 8 * 250 + 30 '���8��
        End With
        
        lblAppend.Visible = True: rtfAppend.Visible = True: fraBorder(4).Visible = True
        Init���븽�� = True
    End If
    
    '�Ѳ����ڵ�������Ŀ��ʾ
    If strNoneAppend <> "" Then
        MsgBox "���¸�������Ŀ��Ӧ�ĵ����������Ѳ����ڣ�" & vbCrLf & Mid(strNoneAppend, 2), vbInformation, gstrSysName
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetOrderInspectInfo(ByVal lng����ID As Long, ByVal strCondition As String, ByVal intType As Integer, ByVal lng����ID As Long) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵ�
    Dim strText As String
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, lng����ID, lng����ID, strCondition)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(lng����ID, strCondition)
        End If
    End If
    GetOrderInspectInfo = strText
End Function

Private Function GetAppendItemValue(ByVal str��Ŀ As String, ByVal lngҪ��ID As Long, ByVal str������ As String) As String
'���ܣ���ȡָ�������븽��ֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim arrItem As Variant, i As Long
    Dim lng����ID As Long
    Dim intType As Integer '1-���2��סԺ
    
    On Error GoTo errH
    
    If TypeName(mvar����ID) = "String" Then
        intType = 1
    Else
        intType = 2
    End If
    
    If intType = 1 And strText = "" Then
        '�����ϣ������δ�������¼���������ȡ
        If str��Ŀ Like "*���" And strText = "" And mstrDiagnosis <> "" Then
            strText = mstrDiagnosis
        End If
    End If
    
    '����Ҫ�ص���ȡ����סԺ����Դ�
    '������ȡ����ҽ���еĸ���Ŀ����ȡ�����еģ�2-סԺ��ȡ��������ȡҽ��
    If intType = 1 And strText = "" Then
        '4.δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
        strSQL = " Select ���� From (" & _
            " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
            IIF(TypeName(mvar����ID) = "String", " And A.�Һŵ�=[2]", " And A.��ҳID=[3]") & _
            " And B.��Ŀ=[5] And B.���� is Not Null and nvl(a.ҽ��״̬,0)<>4" & _
            " Order by A.����ʱ�� Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CStr(mvar����ID), Val(mvar����ID), mintӤ��, str��Ŀ)
        If Not rsTmp.EOF Then strText = nvl(rsTmp!����)
    End If
    
    
    '����ж�ӦҪ�أ���Ҫ����ȡ������ȡ
    If lngҪ��ID <> 0 And strText = "" Then
        '���ϰ棬���°�
        If TypeName(mvar����ID) = "String" Then '����
            strSQL = "Select Zl_Replace_Element_Value(B.������,[1],A.ID,1) as ����" & _
                " From ���˹Һż�¼ A,����������Ŀ B Where A.NO=[2] And B.ID=[3] And a.��¼����=1 And a.��¼״̬=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CStr(mvar����ID), lngҪ��ID)
        Else
            strSQL = "Select Zl_Replace_Element_Value(������,[1],[2],2) as ���� From ����������Ŀ Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mvar����ID), lngҪ��ID)
        End If
        If Not rsTmp.EOF Then strText = nvl(rsTmp!����)
        If strText = "" Then
            
            If TypeName(mvar����ID) = "String" Then
                strSQL = "select a.id From ���˹Һż�¼ A Where A.NO=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(mvar����ID))
                lng����ID = Val(rsTmp!ID & "")
                intType = 1
            Else
                lng����ID = Val(mvar����ID)
                intType = 2
            End If
            strText = GetOrderInspectInfo(mlng����ID, str������, intType, lng����ID)
        End If
    End If
    
    'δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰδ�����ҽ������ȡ,�������д��Ϊ׼
    If strText = "" And mstrAdvItem <> "" Then
        arrItem = Split(mstrAdvItem, "<Split1>")
        For i = 0 To UBound(arrItem)
            If Split(arrItem(i), "<Split2>")(0) = str��Ŀ Then
                strText = Split(arrItem(i), "<Split2>")(1): Exit For
            End If
        Next
    End If
    
    If strText = "" And intType = 2 Then
        'δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
        strSQL = " Select ���� From (" & _
            " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
            IIF(TypeName(mvar����ID) = "String", " And A.�Һŵ�=[2]", " And A.��ҳID=[3]") & _
            " And B.��Ŀ=[5] And B.���� is Not Null and nvl(a.ҽ��״̬,0)<>4" & _
            " Order by A.����ʱ�� Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CStr(mvar����ID), Val(mvar����ID), mintӤ��, str��Ŀ)
        If Not rsTmp.EOF Then strText = nvl(rsTmp!����)
    End If
    
    GetAppendItemValue = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init�������() As Boolean
'���ܣ���ʼ��������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnLis As Boolean
    Dim arrItems As Variant, strItems As String
    Dim i As Long, j As Long
    Dim strLIS As String
    Dim strTmp As String
    Dim colTmp As New Collection
    Dim strItemTmp As String
    Dim lng��ID As Long
    Dim Y As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = IIF(mlng��ĿID <> 0, mlng��ĿID, "") & ";"
    strItems = CStr(Split(strSQL, ";")(0))
    Me.txtData.Text = Split(strSQL, ";")(1)
    cmdData.Tag = txtData.Text
    
    If strItems <> "" Then
        '�ж��Ƿ����°�LISģʽ�������Ŀ
        If Not gobjLIS Is Nothing Then
            blnLis = gobjLIS.CheckLisSate
        End If
        If mblnNewLIS And blnLis Then
            strLIS = " Union All" & vbNewLine & _
                    "       Select e.Id, e.����, e.����, e.��������, e.�Թܱ���, ���������Ŀ.���� As ���,���������Ŀ.id as ��ID " & vbNewLine & _
                    "       From ���������Ŀ, ���鱨����Ŀ C, ���鱨����Ŀ D, ������ĿĿ¼ E" & vbNewLine & _
                    "       Where ���������Ŀ.Id = c.������Ŀid And c.������Ŀid = d.������Ŀid And d.������Ŀid = e.Id And e.�����Ŀ <> 1 And ���������Ŀ.Id <> e.Id"
            '�ֽ�����
            For i = 0 To UBound(Split(strItems, ","))
                strTmp = Split(strItems, ",")(i)
                If InStr(strTmp, "|") > 0 Then
                    colTmp.Add Mid(strTmp, InStr(strTmp, "|") + 1), "_" & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                    strItemTmp = strItemTmp & "," & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                Else
                    strItemTmp = strItemTmp & "," & strTmp
                End If
            Next
            strItems = Mid(strItemTmp, 2)
            Me.Height = Me.Height + 1200
            vsExt.Height = vsExt.Height + 1200
        End If
        strSQL = "Select * From (With ���������Ŀ As (Select /*+ Rule*/ A.ID,A.����,A.����,A.��������,A.�Թܱ���, a.���� As ���,null as ��ID  From ������ĿĿ¼ A " & _
            " Where A.���='C' " & _
            IIF(mint���� = 2, "", " And Nvl(A.����Ӧ��,0)=1") & _
            " And A.������� IN(" & mint������� & ",3" & IIF(mint���� = 2, ",4", "") & ") " & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And A.ID In(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[2])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID)))" & _
            " Select * from ���������Ŀ" & _
            strLIS & _
            ") Order by ���,����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strItems, mlng���˿���id)
    End If
        
    vsExt.Clear
    If strItems <> "" Then
        vsExt.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
    Else
        vsExt.Rows = 2
    End If
    vsExt.Cols = 4
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 2) = "������Ŀ"
    If mblnNewLIS Then
        vsExt.ColWidth(2) = 3700
        vsExt.ColWidth(0) = 300
    Else
        vsExt.ColWidth(2) = 4000
        vsExt.ColHidden(0) = True
    End If
    vsExt.ColHidden(1) = True
    vsExt.ColHidden(3) = True
    vsExt.FixedAlignment(2) = 4
    vsExt.ColAlignment(2) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If strItems <> "" Then
        If Not rsTmp.EOF Then
            arrItems = Split(strItems, ",") '����ԭ������˳��
            For i = 0 To UBound(arrItems)
                rsTmp.Filter = "ID=" & arrItems(i)
                If Not rsTmp.EOF Then
                    Y = vsExt.FindRow(CLng(rsTmp!ID))
                    '�ظ���ָ�겻����
                    If Y = -1 Then
                        j = j + 1
                        vsExt.RowData(j) = CLng(rsTmp!ID)
                        '����Ĭ�Ϲ�ѡ���Ҳ���ȡ��
                        vsExt.TextMatrix(j, 0) = " "
                        vsExt.Cell(flexcpBackColor, j, 0) = &H8000000F
                        vsExt.TextMatrix(j, 2) = "[" & rsTmp!���� & "]" & rsTmp!����
                        vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '���ڻָ���ʾ
                        vsExt.TextMatrix(j, 1) = NVL(rsTmp!��������)
                        vsExt.Cell(flexcpData, j, 1) = CStr(NVL(rsTmp!�Թܱ���)) '����ͬ����������
                        vsExt.TextMatrix(j, 3) = 0   '����
    '                    If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                    End If
                    If mblnNewLIS Then
                        lng��ID = CLng(rsTmp!ID)
                        rsTmp.Filter = "��ID=" & CLng(rsTmp!ID)
                        Do While Not rsTmp.EOF
                            Y = vsExt.FindRow(CLng(rsTmp!ID))
                            '�ظ���ָ�겻����
                            If Y = -1 Then
                                j = j + 1
                                vsExt.RowData(j) = CLng(rsTmp!ID)
                                On Error Resume Next
                                strItemTmp = ""
                                strItemTmp = colTmp("_" & lng��ID)
                                On Error GoTo errH
                                If InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                    vsExt.Cell(flexcpChecked, j, 0) = 1
                                ElseIf strItemTmp = "" And mblnNew Then  '��һ�ν���Ĭ�Ϲ�ѡ
                                    vsExt.Cell(flexcpChecked, j, 0) = 1
                                Else
                                    vsExt.Cell(flexcpChecked, j, 0) = 2
                                End If
                                '��������
                                vsExt.TextMatrix(j, 2) = "    [" & rsTmp!���� & "]" & rsTmp!����
                                vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '���ڻָ���ʾ
                                vsExt.TextMatrix(j, 1) = NVL(rsTmp!��������)
                                vsExt.Cell(flexcpData, j, 1) = CStr(NVL(rsTmp!�Թܱ���)) '����ͬ����������
        '                       If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                                vsExt.TextMatrix(j, 3) = 1    '����
                            Else
                                '����ظ���ָ�깴ѡ��ǰ���ָ��δ��ѡ����ɾ��ǰ���ָ����غ����ָ��
                                On Error Resume Next
                                strItemTmp = ""
                                strItemTmp = colTmp("_" & lng��ID)
                                On Error GoTo errH
                                If vsExt.Cell(flexcpChecked, Y, 0) = 1 And InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                    vsExt.RemoveItem Y
                                    vsExt.AddItem ""
                                    vsExt.RowData(j) = CLng(rsTmp!ID)
                                    vsExt.Cell(flexcpChecked, j, 0) = 1
                                    '��������
                                    vsExt.TextMatrix(j, 2) = "    [" & rsTmp!���� & "]" & rsTmp!����
                                    vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '���ڻָ���ʾ
                                    vsExt.TextMatrix(j, 1) = NVL(rsTmp!��������)
                                    vsExt.Cell(flexcpData, j, 1) = CStr(NVL(rsTmp!�Թܱ���)) '����ͬ����������
            '                       If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                                    vsExt.TextMatrix(j, 3) = 1    '����
                                End If
                            End If
                            rsTmp.MoveNext
                        Loop
                    End If
                End If
            Next
        End If
        If j > 0 Then vsExt.Rows = j + 1
    End If
    
    vsExt.Row = 1: vsExt.Col = 2
    Init������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitCombox(Optional ByVal strNewItemID As String = "", Optional ByVal DefaultValue As String = "") As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String, lngItemCount As Long
    
    InitCombox = False
    
    On Error GoTo DBError
    strTmp = "": lngItemCount = 0
    For i = 1 To vsExt.Rows - 1
        If vsExt.RowData(i) <> 0 And (i <> vsExt.Row Or Len(strNewItemID) = 0) Then
            lngItemCount = lngItemCount + 1
            strTmp = strTmp & "," & vsExt.RowData(i)
        End If
    Next
    If Len(strNewItemID) > 0 Then
        lngItemCount = lngItemCount + 1
        strTmp = strTmp & "," & strNewItemID
    End If
    If Len(strTmp) > 0 Then strTmp = Mid(strTmp, 2)

    If lngItemCount = 0 Then
        strSQL = "Select ���� From ���Ƽ���걾" & _
            "     Where (Instr(Nvl(�����Ա�,'����'),'��')=0 And Instr(Nvl(�����Ա�,'����'),'Ů')=0" & _
            "         Or Instr(Nvl([1],'����'),'��')=0 And Instr(Nvl([1],'����'),'Ů')=0" & _
            "         Or Instr([1],'��')>0 And Instr(�����Ա�,'��')>0 Or Instr([1],'Ů')>0 And Instr(�����Ա�,'Ů')>0)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr�Ա�)
    Else
        strSQL = "Select /*+ Rule*/ �걾����,Sum(1) From (" & _
            "   Select Distinct A.ID,B.���� As �걾����" & _
            "   From ������ĿĿ¼ A,���Ƽ���걾 B,������Ŀ�ο� C,���鱨����Ŀ D" & _
            "   Where A.ID=D.������ĿID(+) And D.������ĿID=C.��ĿID(+)" & _
            "        And (NVL(C.�걾����,'') Is Null Or NVL( C.�걾����,'')=B.����) " & _
            "       And A.ID In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            "       And (Instr(Nvl(B.�����Ա�,'����'),'��')=0 And Instr(Nvl(B.�����Ա�,'����'),'Ů')=0" & _
            "         Or Instr(Nvl([3],'����'),'��')=0 And Instr(Nvl([3],'����'),'Ů')=0" & _
            "         Or Instr([3],'��')>0 And Instr(B.�����Ա�,'��')>0 Or Instr([3],'Ů')>0 And Instr(B.�����Ա�,'Ů')>0)" & _
            "           And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[4])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " ) Group By �걾���� Having Sum(1)=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strTmp, lngItemCount, mstr�Ա�, mlng���˿���id)
    End If
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "δ���ü���걾���뵽�ֵ�����������á�", _
            lngItemCount = 1, "ѡȡ�ļ�����Ŀδ�������걾�����ȵ�������Ŀ����������", _
            lngItemCount > 1, "ѡȡ�ļ�����Ŀ�ļ���걾��������Ŀ�Ĳ�һ�£����ȵ�������Ŀ����������"), vbInformation, gstrSysName
        Exit Function
    End If
    
    With cbo�걾
        strTmp = .Text
        
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp(0)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
        On Error Resume Next
        If Len(DefaultValue) > 0 Then
            .Text = DefaultValue
        Else
            .Text = strTmp
        End If
    End With
    InitCombox = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mintType)

    mlngHwnd = 0
    mintӤ�� = 0
    mlng����ID = 0
    mlng���˿���id = 0
    mvar����ID = Empty
    mstr�Ա� = ""
    mint���� = 0
    mintType = 0
    mbytUseType = 0
    mint��Ч = 0
    mint������� = 0
    mint�������� = 0
    mblnNewLIS = False
    mblnNew = False
    mlng��ĿID = 0
    mstrAdvItem = ""
    mstrDiagnosis = ""
    mstr�����ȼ� = ""
    Set mrsAppend = Nothing

    Set mfrmParent = Nothing
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 0 Then
            If Me.Height - Y < 2355 Or Me.Height - Y > 7200 Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 1 Then
            If Me.Width + x < 4140 Or Me.Width + x > 9600 Then Exit Sub
            Me.Width = Me.Width + x
        ElseIf Index = 4 Then
            If vsExt.Height + Y < 1000 Or vsExt.Height + Y > Me.Height * 0.7 Then Exit Sub
            vsExt.Height = vsExt.Height + Y
            Call Form_Resize
        End If
    End If
End Sub

Private Function CursorInItem(Optional ByRef str��Ŀ As String, Optional ByRef bln���� As Boolean) As Boolean
'���ܣ��жϵ�ǰ����Ƿ���ĳ����Ŀ�������
    Dim lngLoc As Long, i As Long
    
    With rtfAppend
        mrsAppend.MoveFirst
        For i = 1 To mrsAppend.RecordCount
            lngLoc = .Find(mrsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngLoc = -1 Then
                lngLoc = InStr(rtfAppend.Text, mrsAppend!��Ŀ & "��")
                lngLoc = lngLoc - 1
            End If
            If lngLoc <> -1 Then
                lngLoc = lngLoc + Len(mrsAppend!��Ŀ & "��")
                If .SelStart >= lngLoc And InStr(Mid(.Text, lngLoc, IIF(.SelStart - lngLoc < 0, 0, .SelStart - lngLoc)), vbCrLf) = 0 Then
                    bln���� = True
                    str��Ŀ = NVL(mrsAppend!������)
                End If
                If .SelStart = lngLoc Then CursorInItem = True: Exit Function
            End If
            mrsAppend.MoveNext
        Next
    End With
End Function

Private Sub imgSentence_Click()
    Dim strSentence As String
    Dim str��鲿λ As String, str��鷽�� As String
    
    If txtSentence.Tag = "����ҽ��" Then
        Call FindDoctor("")
    ElseIf txtSentence.Tag = "�������" Then
        Call FindDept("")
    ElseIf txtSentence.Tag = "ѡ��ѡ��" Then
        Call FindItem
    Else
        Call Get��鲿λ����(str��鲿λ, str��鷽��)
        
        strSentence = frmSentenceSel.ShowMe(Me, mint�������, mlng����ID, mvar����ID, mlng��ĿID, str��鲿λ, str��鷽��, , , , mobjEmrInterface)
        If strSentence <> "" Then
            rtfAppend.SelText = strSentence
            Call HideWordInput(True)
        End If
    End If
End Sub

Private Sub rtfAppend_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub rtfAppend_KeyPress(KeyAscii As Integer)
    Dim str��Ŀ���� As String
    Dim bln���� As Boolean
    
    If KeyAscii = 13 Then
        If txtSentence.Tag <> "����ҽ��" Then
            '�������λس������ת
            With rtfAppend
                If .SelStart - 1 > 0 Then
                    If Mid(.Text, .SelStart - 1, 2) = vbCrLf Then
                        KeyAscii = 0
                        Call zlCommFun.PressKey(vbKeyBack)
                        If InStr(Mid(.Text, .SelStart + 1), "�� ") > 0 Then
                            Call zlCommFun.PressKey(vbKeyDown)
                            Call zlCommFun.PressKey(vbKeyEnd)
                        Else
                            Call zlCommFun.PressKey(vbKeyTab)
                        End If
                    End If
                End If
            End With
        Else
            KeyAscii = 0
            bln���� = CursorInItem(str��Ŀ����, False)
            If str��Ŀ���� = "����ҽ��" Or str��Ŀ���� = "����ҽ��" Then
                With rtfAppend
                    .SelStart = IIF(InStr(Mid(.Text, .SelStart + 1), "�� ") > 0, .SelStart + 1 + InStr(Mid(.Text, .SelStart + 1), "�� ") + 1, .SelStart - IIF(bln����, 0, 1))
                    .SelLength = IIF(InStr(Mid(.Text, .SelStart), vbCrLf) - 1 > 0, InStr(Mid(.Text, .SelStart), vbCrLf) - 1, Len(.Text))
                End With
            End If
        End If
    ElseIf KeyAscii = 8 And txtSentence.Tag <> "����ҽ��" Then
        '������ɾ��������" "
        With rtfAppend
            If .SelLength = 0 And .SelStart > 0 Then
                If Mid(.Text, .SelStart, 1) = "��" Then
                    If CursorInItem Then
                         If Mid(.Text, .SelStart + 1, 1) <> " " Then .SelText = " "
                    End If
                End If
            End If
        End With
    Else
        If txtSentence.Tag = "����ҽ��" Then KeyAscii = 0: Call rtfAppend_SelChange
        If txtSentence.Tag = "ѡ��ѡ��" Then KeyAscii = 0: Call rtfAppend_SelChange
    End If
End Sub

Private Sub rtfAppend_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub rtfAppend_SelChange()
    Dim str��Ŀ As String
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    Dim bytType  As Byte 'bytType 0-��Ϊ�ʾ䣬1-������Ա��2-�������
    Dim str��ֵ�� As String
    
    With rtfAppend
        If .Visible And .SelLength = 0 And .SelStart > 0 Then
            bln���� = CursorInItem(str��Ŀ, bln����)
            If str��Ŀ = "����ҽ��" Or str��Ŀ = "����ҽ��" Then
                bytType = 1
            ElseIf str��Ŀ = "����ҽ������" Then
                bytType = 2
            ElseIf InStr("," & mstr��ѡ��Ŀ & ",", "," & str��Ŀ & ",") > 0 And mstr��ѡ��Ŀ <> "" Then
                bytType = 3
                mrsAppend.Filter = "������='" & str��Ŀ & "'"
                If mrsAppend.RecordCount > 0 Then mrsAppend.MoveFirst: str��ֵ�� = mrsAppend!��ֵ�� & ""
                 mrsAppend.Filter = 0
            End If
            
            If bln���� And bytType <> 0 And Not picSentence.Visible Then
                .SelStart = IIF(InStrRev(Mid(.Text, 1, .SelStart + 1), "�� ") > 0, InStrRev(Mid(.Text, 1, .SelStart + 1), "�� ") + 1, .SelStart - IIF(bln����, 0, 1))
                .SelLength = IIF(InStr(Mid(.Text, .SelStart), vbCrLf) - 1 > 0, InStr(Mid(.Text, .SelStart), vbCrLf) - 1, Len(.Text))
                Call ShowWordInput(bytType, str��ֵ��)
            Else
                If Mid(.Text, .SelStart, 2) = "�� " Then
                    '��겻����λ��������" "��
                    If CursorInItem() Then .SelStart = .SelStart + 1
                ElseIf Mid(.Text, .SelStart, 1) = "`" Then
                    '�ʾ����������⴦��
                    '��vbBack�ﲻ��Ч��
                    .SelStart = .SelStart - 1
                    .SelLength = 1: .SelText = ""
                    Call ShowWordInput
                Else
                    If Not (str��Ŀ = "����ҽ��" Or str��Ŀ = "����ҽ��") Then txtSentence.Tag = ""
                End If
            End If
        End If
    End With
End Sub

Private Sub txtData_GotFocus()
    zlControl.TxtSelAll txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset, vRect As RECT
    Dim strSQL As String, str�Ա� As String
    Dim strLike As String, blnCancel As Boolean
    
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "0,1"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "0,2"
    Else
        str�Ա� = "0"
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtData.Text = "" Then
            If mintType = 1 Then '�������Բ�����������Ŀ
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        ElseIf txtData.Text = cmdData.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        '�Ż�
        strLike = mstrLike
        If Len(txtData.Text) < 2 Then strLike = ""
        
        If mintType = 1 Then
            '����������Ŀ
            strSQL = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��������" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='G' And A.������� IN([3],3)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
                    " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[5])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                " Order by A.����"
            vRect = zlControl.GetControlRect(txtData.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, vRect.Left, vRect.Top, txtData.Height, blnCancel, False, True, _
                UCase(txtData.Text) & "%", strLike & UCase(txtData.Text) & "%", mint�������, mint���� + 1, mlng���˿���id)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                txtData.Text = cmdData.Tag
                zlControl.TxtSelAll txtData
                Exit Sub
            End If
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
            cmdData.Tag = txtData.Text
            
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf mintType = 4 Then
            '����걾
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmdData_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
'���ܣ��ָ���ʾԭ����
    If txtData.Text <> cmdData.Tag Then
        txtData.Text = cmdData.Tag
    End If
End Sub

Private Sub txtSentence_GotFocus()
    Call zlControl.TxtSelAll(txtSentence)
End Sub

Private Function GetDoctorLevel(ByVal str���� As String) As String
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "Select �����ȼ� From ��Ա�� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, str����)
    If rsTmp.RecordCount > 0 Then
        GetDoctorLevel = rsTmp!�����ȼ� & ""
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindDoctor(ByVal strTmp As String) As Recordset
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vPoint As PointAPI
    Dim blnCancel As Boolean, str��Ŀ���� As String, blnDo As Boolean
    Dim lngStart As Long
    Dim lng��Աid As Long
    Dim str���� As String
    
    On Error GoTo errH
    strInput = Trim(UCase(strTmp))   '�����ֵ����ǰ׺�ո�
    strSQL = "Select A.ID,A.���,A.����,A.����,A.�����ȼ�" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.���"
    vPoint = zlControl.GetCoordPos(txtSentence.Hwnd, txtSentence.Left + 15, txtSentence.Top + 3300 + txtSentence.Height)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ��", False, "", "", False, False, True, _
        vPoint.x, vPoint.Y, 3000, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call CursorInItem(str��Ŀ����, False)
            If str��Ŀ���� = "����ҽ��" Then
                If MsgBox("û���ҵ�ƥ���ҽ������ȷ��Ҫ����û�н�����Ա������ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnDo = True
                    strTmp = strInput
                Else
                    blnDo = False
                End If
            Else
                Call MsgBox("û���ҵ�ƥ���ҽ��!", vbInformation, gstrSysName)
                blnDo = False
            End If
        End If
    Else
        blnDo = True
        strTmp = rsTmp!���� & ""
        lng��Աid = rsTmp!ID
        Call CursorInItem(str��Ŀ����, False)
        If str��Ŀ���� = "����ҽ��" And mbln�����ּ����� Then
            mstr�����ȼ� = rsTmp!�����ȼ� & ""
        End If
    End If
    
    If blnDo Then
        rtfAppend.SelText = strTmp
        
        strSQL = "Select b.����, a.ȱʡ From ������Ա A, ���ű� B, ��������˵�� C" & _
            " Where a.����id = b.Id And b.Id = c.����id And c.�������� = '�ٴ�' And a.��Աid = [1]" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Աid)
        If Not rsTmp.EOF Then
            str���� = rsTmp!���� & ""
            rsTmp.Filter = "ȱʡ = 1"
            If rsTmp.RecordCount > 0 Then str���� = rsTmp!���� & ""
        End If
         
        lngStart = rtfAppend.SelStart
        Call Do������������("����ҽ������", True, str����)
        rtfAppend.SelStart = lngStart
 
        Call HideWordInput(True)
    Else
        txtSentence.SetFocus
        Call zlControl.TxtSelAll(txtSentence)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindItem() As Recordset
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vPoint As PointAPI
    Dim blnCancel As Boolean, strTmp As String, blnDo As Boolean
    
    strInput = Replace(imgSentence.Tag, ";", ",")
    If strInput = "" Then Exit Function
    
    On Error GoTo errH
    strSQL = "Select rownum as ID, Column_Value as ѡ���� From Table(Cast(f_Str2List([1]) As zlTools.t_StrList))"
    vPoint = zlControl.GetCoordPos(txtSentence.Hwnd, txtSentence.Left + 15, txtSentence.Top + 3300 + txtSentence.Height)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ��", False, "", "", False, False, True, _
        vPoint.x, vPoint.Y, 3000, blnCancel, False, True, strInput)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            blnDo = False
        End If
    Else
        blnDo = True
        strTmp = rsTmp!ѡ���� & ""
    End If
    
    If blnDo Then
        rtfAppend.SelText = strTmp
        Call HideWordInput(True)
    Else
        txtSentence.SetFocus
        Call zlControl.TxtSelAll(txtSentence)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindDept(ByVal strTmp As String) As Recordset
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vPoint As PointAPI
    Dim blnCancel As Boolean, str��Ŀ���� As String, blnDo As Boolean
    Dim strDoctor As String
    
    On Error GoTo errH
    strInput = Trim(UCase(strTmp))   '�����ֵ����ǰ׺�ո�
    strDoctor = Do������������("����ҽ��")
    strDoctor = Trim(Replace(strDoctor, vbCrLf, "")) 'ȥ���س��Ϳհ�
    strSQL = "Select Distinct A.ID,A.����,A.���� as ����,A.���� From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And a.Id = b.����id And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And B.��������='�ٴ�'" & _
        IIF(strDoctor <> "", " And a.id in (select x.����id from ������Ա X, ��Ա�� Y where x.��Աid=y.id and y.����=[3])", "") & _
        " Order by A.����"
    
    vPoint = zlControl.GetCoordPos(txtSentence.Hwnd, txtSentence.Left + 15, txtSentence.Top + 3300 + txtSentence.Height)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ��", False, "", "", False, False, True, _
        vPoint.x, vPoint.Y, 3000, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", strDoctor)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ��Ŀ���!", vbInformation, gstrSysName)
            blnDo = False
        End If
    Else
        blnDo = True
        strTmp = rsTmp!���� & ""
    End If
    
    If blnDo Then
        rtfAppend.SelText = strTmp
        Call HideWordInput(True)
    Else
        txtSentence.SetFocus
        Call zlControl.TxtSelAll(txtSentence)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Do������������(ByVal str��Ŀ���� As String, Optional ByVal blnSet As Boolean, Optional ByVal strValue As String) As String
'�ܹ������û����ǻ�ȡ��Ӧ��str��Ŀ���ƣ����������븽������
'������str��Ŀ���� ������Ŀ��������
'      blnSet true ����Ŀ��ֵ��false ȡֵȻ�󷵻�
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strData As String
    
    If rtfAppend.Visible Then
        mrsAppend.MoveFirst
        For i = 1 To mrsAppend.RecordCount
            If mrsAppend!������ = str��Ŀ���� Then
                strData = "": lngBegin = -1: lngEnd = -1
                lngBegin = rtfAppend.Find(mrsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
                If lngBegin = -1 Then
                    lngBegin = InStr(rtfAppend.Text, mrsAppend!��Ŀ & "��")
                    lngBegin = lngBegin - 1
                End If
                If lngBegin <> -1 Then
                    lngBegin = lngBegin + Len(mrsAppend!��Ŀ & "��")
                    If i = mrsAppend.RecordCount Then
                        lngEnd = Len(rtfAppend.Text)
                    Else
                        mrsAppend.MoveNext
                        lngEnd = rtfAppend.Find(vbCrLf & mrsAppend!��Ŀ & "��", lngBegin, , rtfNoHighlight Or rtfMatchCase)
                        If lngEnd = -1 Then
                            lngEnd = InStr(rtfAppend.Text, vbCrLf & mrsAppend!��Ŀ & "��")
                            lngEnd = lngEnd - 1
                        End If
                        mrsAppend.MovePrevious
                    End If
                End If
                If lngBegin <> -1 And lngEnd <> -1 Then
                    'MID��������1Ϊ������rtf����0Ϊ����
                    lngBegin = lngBegin + 1
                    lngEnd = lngEnd + 1
                    strData = Mid(rtfAppend.Text, lngBegin, lngEnd - lngBegin)
                    'ȥ��Ϊ��������ı����һ��λ�ò���ֱ��¼�뺺������ӵĿո�
                    If Left(strData, 1) = " " Then strData = Mid(strData, 2)
                    If Right(strData, 1) = " " Then strData = Left(strData, Len(strData) - 1)
                End If
                If blnSet Then
                    rtfAppend.SelStart = lngBegin
                    rtfAppend.SelLength = lngEnd - lngBegin
                    rtfAppend.SelText = strValue
                    Exit Function
                End If
            End If
            mrsAppend.MoveNext
        Next
    End If
    Do������������ = strData
End Function

Private Sub txtSentence_KeyPress(KeyAscii As Integer)
    Dim strSentence As String, blnCancel As Boolean
    Dim str��鲿λ As String, str��鷽�� As String
    Dim str��Ŀ���� As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtSentence.Tag = "����ҽ��" Then
            If txtSentence.Text = "" Then
                rtfAppend.SelText = ""
                Call HideWordInput(True)
                Call CursorInItem(str��Ŀ����, False)
                If str��Ŀ���� = "����ҽ��" And mbln�����ּ����� Then
                    mstr�����ȼ� = ""
                End If
            Else
                Call FindDoctor(txtSentence.Text)
            End If
        ElseIf txtSentence.Tag = "�������" Then
            If txtSentence.Text = "" Then
                rtfAppend.SelText = ""
                Call HideWordInput(True)
            Else
                Call FindDept(txtSentence.Text)
            End If
        Else
            Call Get��鲿λ����(str��鲿λ, str��鷽��)
            
            strSentence = frmSentenceSel.ShowMe(Me, mint�������, mlng����ID, mvar����ID, mlng��ĿID, str��鲿λ, str��鷽��, txtSentence.Text, picSentence.Hwnd, blnCancel, mobjEmrInterface)
            If strSentence <> "" Then
                rtfAppend.SelText = strSentence
                Call HideWordInput(True)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ĵʾ䡣", vbInformation, gstrSysName
                End If
                Call zlControl.TxtSelAll(txtSentence)
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call imgSentence_Click
    End If
End Sub

Private Sub txtSentence_LostFocus()
    If Not frmSentenceSel.mblnShow Then
        Call HideWordInput(False) '���شʾ�����
    End If
End Sub

Private Sub ShowWordInput(Optional ByVal bytType As Byte, Optional ByVal str��ֵ�� As String)
'���ܣ���ʾ�ʾ�����
'������bytType 0-��Ϊ�ʾ䣬1-������Ա��2-�������,3-ѡ��ѡ��
    Dim vPos As PointAPI
    Dim blnLocked As Boolean
    
    imgSentence.Tag = ""
    If bytType = 1 Then
        txtSentence.Tag = "����ҽ��"
    ElseIf bytType = 2 Then
        txtSentence.Tag = "�������"
    ElseIf bytType = 3 Then
        txtSentence.Tag = "ѡ��ѡ��"
        imgSentence.Tag = str��ֵ��
        blnLocked = True
    Else
        txtSentence.Tag = ""
    End If
    txtSentence.Locked = blnLocked
    
    If rtfAppend.Visible And rtfAppend.Enabled Then
        vPos = GetCaretPos(rtfAppend.Hwnd)
        If vPos.x <> -1 And vPos.Y <> -1 Then
            If rtfAppend.Left + vPos.x + Screen.TwipsPerPixelX * 2 < rtfAppend.Left + rtfAppend.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX Then
                picSentence.Left = rtfAppend.Left + vPos.x + Screen.TwipsPerPixelX * 2
            Else
                picSentence.Left = rtfAppend.Left + rtfAppend.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX
            End If
            picSentence.Top = rtfAppend.Top + vPos.Y + Screen.TwipsPerPixelY
            If bytType <> 0 Then
                txtSentence.Text = rtfAppend.SelText
            Else
                txtSentence.Text = ""
            End If
            picSentence.Visible = True
            txtSentence.SetFocus
        End If
    End If
End Sub

Private Sub HideWordInput(ByVal blnFocus As Boolean)
'���ܣ����شʾ�����
    picSentence.Visible = False
    txtSentence.Text = ""
    If blnFocus And rtfAppend.Visible And rtfAppend.Enabled Then
        rtfAppend.SetFocus
    End If
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����:��ʾѡ��ť,����֤��ǰ��Ԫ��ɼ�
    Dim strKey As String, lngҩ��ID As Long
    
    If mblnChangeSel = True Then Exit Sub
    '��֤��ǰ��Ԫ��ɼ�
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If
    
    If mintType = 1 Or mintType = 4 Then
        '��ʾ/��������ѡ��ť
        If NewCol = 0 And mintType = 1 Or NewCol = 2 And mintType = 4 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(NewRow, 3) = "1" Then
                    cmd.Visible = False
                Else
                    cmd.Visible = True
                End If
            Else
                cmd.Visible = True
            End If
        Else
            cmd.Visible = False
        End If
        If cmd.Visible Then
            vsExt.FocusRect = flexFocusSolid
        Else
            vsExt.FocusRect = flexFocusLight
        End If
    End If
    
End Sub

Private Sub vsExt_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'����:����ĳЩ�п�ķ�Χ
    If Row = -1 Then
        If mintType = 1 Or mintType = 4 Then
            Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ���������ťλ��
        End If
    End If
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mintType = 0 Then
        If NewCol = 0 Or NewCol = 3 Then
            Cancel = True
            If NewRow <> OldRow Then vsExt.Row = NewRow
        End If
    End If
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
    If fraMethod.Visible Then fraMethod.Visible = False
End Sub

Private Function GetOnlyOneMethod(ByVal strMethod As String) As String
'���ܣ����ݲ�λ�ķ������壬���ֻ��һ��������ѡ���򷵻ظ÷���
'ע�⣺��3������Ϊ�������ţ�< vbTab  ;  , >
    Dim strTmp As String
    
    If strMethod = "" Then Exit Function
    strTmp = strMethod
    
    strTmp = Replace(strTmp, vbTab, ";")
    strTmp = Replace(strTmp, ",", ";")
    strTmp = Replace(strTmp, ";;", ";")
    strTmp = "<spdel>" & strTmp & "<spdel>"
    strTmp = Replace(strTmp, "<spdel>;", "")
    strTmp = Replace(strTmp, ";<spdel>", "")
    strTmp = Replace(strTmp, "<spdel>", "")
    
    If InStr(strTmp, ";") = 0 Then GetOnlyOneMethod = Mid(strTmp, 2)        'ȥ��ǰ��λ��Ӱ���
End Function

Private Sub vsExt_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strMethod As String, i As Long, j As Long
    Dim arrMethod As Variant, arrSub As Variant
    Dim lngTmp As Long
    Dim k As Long
    Dim blnDo As Boolean

    strMethod = vsExt.Cell(flexcpData, Row, Col)
    If strMethod = "" Then
        MsgBox "�ü�鲿λû�����ÿɹ�ѡ��ļ�鷽����", vbInformation, gstrSysName
        Exit Sub
    End If
    With vsMethod
        .Rows = 0
        
        arrMethod = Split(Replace(strMethod, vbTab, ";" & vbTab), ";")
        
        For i = 0 To UBound(arrMethod)
            arrSub = Split(arrMethod(i), ",")
            
            For j = 0 To UBound(arrSub)
                .Rows = .Rows + 1
                If j = 0 Then
                    If InStr(1, arrMethod(i), vbTab) > 0 Then
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 2 '�����ǹ�ѡ��
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 3) '��һλ����Ӱ����־
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 3) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    Else
                        '�ų���
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 1 '�������ų���
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 2) '��һλ����Ӱ����־
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1 '1Ϊѡ��
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    End If
                Else
                    '��ѡ����
                    .RowData(.Rows - 1) = 3 '�����ǹ�ѡ����
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Mid(arrSub(j), 2)
                    If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                        blnDo = True
                        '����û��ѡ��ʱ,�����ѡ��
                        For k = .Rows - 2 To 0 Step -1
                            If .RowData(k) <> 3 Then
                                If .Cell(flexcpData, k, 0) = 0 Then blnDo = False
                                Exit For
                            End If
                        Next
                    Else
                        blnDo = False
                    End If
                    
                    If blnDo Then
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c1").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                    End If
                End If
            Next
        Next

        .Row = 0: .Col = 1
        
        .Height = .Rows * (.RowHeightMin + 15) + 30
        If .Height > Me.ScaleHeight - 100 Then .Height = Me.ScaleHeight - 100
        If .Height < 3 * (.RowHeightMin + 15) + 30 Then .Height = 3 * (.RowHeightMin + 15) + 30
        
        If (vsExt.Width - 30) - (vsExt.CellLeft + 15) <= 0 Then
            For i = 0 To vsExt.Cols - 1
                lngTmp = vsExt.ColWidth(i) + lngTmp
            Next
            Me.Width = lngTmp
        End If
        
        .Width = (vsExt.Width - 30) - (vsExt.CellLeft + 15)
        
        .Left = vsExt.Left + vsExt.CellLeft + 15
        
        .Top = vsExt.Top + vsExt.CellTop + vsExt.CellHeight + 15
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If
        fraMethod.Top = .Top: .Top = 0
        fraMethod.Left = .Left: .Left = 0
        fraMethod.Width = .Width
        fraMethod.Height = .Height + cmdMethodOK.Height + 20
        cmdMethodOK.Top = .Height
        cmdMethodOK.Left = .Width - cmdMethodOK.Width - 20
        
        fraMethod.ZOrder
        If .Tag = "AutoPopup" Then
            fraMethod.Visible = .Rows > 1
        Else
            fraMethod.Visible = True
        End If
        If fraMethod.Visible Then .SetFocus
    End With
End Sub

Private Sub vsExt_DblClick()
    If mintType = 0 Then
        If vsExt.Editable <> flexEDNone And vsExt.MouseCol = 1 And vsExt.MouseRow >= vsExt.FixedRows Then
            Call vsExt_KeyPress(vbKeySpace)
        End If
    End If
End Sub

Private Sub vsExt_GotFocus()
    If fraMethod.Visible Then fraMethod.Visible = False
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ�
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim i As Long, j As Long, k As Long, g As Long
    Dim intRow As Integer        '��Ч��
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strKey As String, lngҩƷID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim rsTmp As Recordset
    Dim lngҩ��ID As Long
    
    If KeyCode = vbKeyDelete Then
        If (mintType = 1 Or mintType = 4) And vsExt.RowData(vsExt.Row) <> 0 Then
            '������°�LIS�����Ŀģʽ��������ɾ������
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(vsExt.Row, 3) = "1" Then Exit Sub
            End If
            If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            '��������Ŀģʽ����ͬʱɾ������
            If mintType = 4 And mblnNewLIS Then
                lngBegin = vsExt.Row + 1
                For j = vsExt.Row + 1 To vsExt.Rows - 1
                    If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
                    lngEnd = j
                Next
                For j = lngEnd To lngBegin Step -1
                    vsExt.RowData(j) = 0
                    For i = 0 To vsExt.Cols - 1
                        vsExt.TextMatrix(j, i) = ""
                        vsExt.Cell(flexcpData, j, i) = ""
                    Next
                    If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                        vsExt.RemoveItem j
                    End If
                Next
            End If
            
            vsExt.RowData(vsExt.Row) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(vsExt.Row, i) = ""
                vsExt.Cell(flexcpData, vsExt.Row, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And vsExt.Row = vsExt.FixedRows) Then
                vsExt.RemoveItem vsExt.Row
            End If
            
            '���³�ʼ�걾
            If mintType = 4 Then InitCombox
        End If
    End If
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'���ܣ��Ǳ༭״̬ʱ���Զ��ƶ���Ԫ��
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '��λ����һӦ���뵥Ԫ��
        If mintType = 0 Then
            If vsExt.Col <= 1 Then
                vsExt.Col = vsExt.Col + 1
            ElseIf vsExt.Col = 2 And vsExt.Row <= vsExt.Rows - 2 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 1
            ElseIf vsExt.Col = 2 And vsExt.Row = vsExt.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If vsExt.Row = vsExt.Rows - 1 Then
                If vsExt.RowData(vsExt.Row) = 0 Or mblnNotAddNew Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    vsExt.AddItem ""
                End If
            End If
            If vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                If mintType = 1 Then
                    vsExt.Col = 0
                Else
                    vsExt.Col = 2
                End If
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If mintType = 0 Then
            If vsExt.Col = 2 Then
                Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            KeyAscii = 0
            If cmd.Visible Then cmd_Click
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If mintType = 0 Then
            If vsExt.Editable <> flexEDNone Then
                If vsExt.Col = 1 Then
                    If vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1 Then
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 0
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c0").Picture
                    Else
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c1").Picture
                        
                        '�Զ���������ѡ����
                        vsExt.Col = 2
                        vsMethod.Tag = "AutoPopup"
                        Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                        vsMethod.Tag = ""
                    End If
                ElseIf vsExt.Col = 2 Then
                    Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                End If
            End If
        End If
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    Dim strPrivs As String, i As Long
    Dim strKey As String, lngҩ��ID As Long
    
    If Not mblnReturn Then
        If mintType = 1 Or mintType = 4 Then
            If Col = 0 And mintType = 1 Or Col = 2 And mintType = 4 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                            
                '���³�ʼ�걾
                If mintType = 4 Then InitCombox
                
            End If
        End If
    End If
End Sub

Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'���ܣ���������ȷ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, int�Ա� As Integer, strҩƷ As String
    Dim strStock As String, blnCancel As Boolean, i As Long
    Dim vPoint As PointAPI, strLike As String
    Dim strSamples As String, strPrivs As String
    Dim strKey As String, lngҩ��ID As Long
    
    If KeyAscii = 13 Then
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        KeyAscii = 0
        
        If mstr�Ա� Like "*��*" Then
            int�Ա� = 1
        ElseIf mstr�Ա� Like "*Ů*" Then
            int�Ա� = 2
        End If
        '�Ż�
        strLike = mstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        
        If mintType = 1 Then
            '���븽������:���ﲻ�ǵ���Ӧ��,��˲�����
            '"-1*������ID"�ǲ��ſ�������ID������Ϊ�����������շ���
            strSQL = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='F' And A.ID<>-1*[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
                    " And A.������� IN([5],3) And Nvl(A.ִ��Ƶ��,0) IN(0,[6]) And Nvl(A.�����Ա�,0) IN(0,[7])" & _
                    " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[8])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                " Order by A.����"
            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng��ĿID, mint���� + 1, mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�, mlng���˿���id)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ø��������Ѿ���������¼�롣", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set��������(Row, rsTmp)
        ElseIf mintType = 4 Then
            '������Ŀ
            With Me.cbo�걾
                For i = 0 To .ListCount - 1
                    strSamples = strSamples & ",'" & .List(i) & "'"
                Next
            End With
            If Len(strSamples) > 0 Then
                strSamples = Mid(strSamples, 2)
            Else
                strSamples = "''"
            End If
            strSQL = "Select A.ID,A.����,A.����,A.��������,A.�걾��λ,A.�Թܱ���" & _
                " From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID" & _
                " And (A.���� Like [1] Or C.���� Like [2] Or C.���� Like [2]) And C.����=[3]" & _
                " And A.���='C' " & _
                IIF(mint���� = 2, "", " And Nvl(A.����Ӧ��,0)=1 ") & _
                " And Nvl(A.�����Ա�,0) In (0,[5])" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And A.������� IN([4],3" & IIF(mint���� = 2, ",4", "") & ") " & _
                " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[6])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)"
            If strLike = "" Then
                '���������ü�������ʱ(����ƥ��),�����(+)����,����ҪGroup Byһ��(���)
                strSQL = strSQL & " Group by A.ID,A.����,A.����,A.��������,A.�걾��λ,A.�Թܱ���"
            End If
            
            strSQL = "Select Distinct A.ID,A.����,A.����,A.�������� as ��������,A.�걾��λ,A.�Թܱ���" & _
                " From ������Ŀ�ο� D,���鱨����Ŀ E,(" & strSQL & ") A" & _
                " Where A.ID=E.������Ŀid(+) And E.������ĿID=D.��Ŀid(+)" & _
                " And (D.�걾���� In (" & strSamples & ") Or D.�걾���� Is Null)" & _
                " Order by A.����"

            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mint���� + 1, mint�������, int�Ա�, mlng���˿���id)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
'            If rsTmp!�������� = "΢����" And vsExt.Rows > 2 Then
'                If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '��������ֻ�ܿ�һ��΢������Ŀ
'                    MsgBox "΢������Ŀֻ�ܵ������룡", vbInformation, gstrSysName
'                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
'                    Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
'                    Exit Sub
'                End If
'            End If
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ü�����Ŀ�Ѿ�¼�룡", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '���������͡��Թܱ����Ƿ���ͬ
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And i <> Row Then
                    If Not (vsExt.TextMatrix(i, 1) = NVL(rsTmp!��������) _
                        Or vsExt.TextMatrix(i, 1) = "" Or NVL(rsTmp!��������) = "") Then
                        MsgBox "��������ͬ�������͵���Ŀ����������Ŀ�ļ�������Ϊ""" & vsExt.TextMatrix(i, 1) & """��", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                        Exit Sub
                    End If
                    If Not (vsExt.Cell(flexcpData, i, 1) = CStr(NVL(rsTmp!�Թܱ���)) _
                        Or vsExt.Cell(flexcpData, i, 1) = "" Or NVL(rsTmp!�Թܱ���) = "") Then
                        MsgBox "��������ͬ�Թܱ������Ŀ����������Ŀ���Թܱ���Ϊ""" & vsExt.Cell(flexcpData, i, 1) & """��", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                        Exit Sub
                    End If
                End If
            Next
            
            '���³�ʼ�걾
            If Not InitCombox(rsTmp!ID, NVL(rsTmp!�걾��λ)) Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set������Ŀ(Row, rsTmp)
            If rsTmp!�������� = "΢����" Then
                mblnNotAddNew = False
'                vsExt.Rows = 2
            Else
                mblnNotAddNew = False
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set��������(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '��������
    vsExt.EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!���� & "]" & rsInput!����
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = NVL(rsInput!��ģ)

    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub Set������Ŀ(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    
    '������Ŀ
    '�����LIS�����Ŀģʽ����ɾ��������·��
    '��������Ŀģʽ����ͬʱɾ������
    If mblnNewLIS Then
        lngBegin = lngRow + 1
        For j = lngRow + 1 To vsExt.Rows - 1
            If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
            lngEnd = j
        Next
        For j = lngEnd To lngBegin Step -1
            vsExt.RowData(j) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(j, i) = ""
                vsExt.Cell(flexcpData, j, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                vsExt.RemoveItem j
            End If
        Next
    End If
    
    vsExt.EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 2) = "[" & rsInput!���� & "]" & rsInput!����
    vsExt.Cell(flexcpData, lngRow, 2) = vsExt.TextMatrix(lngRow, 2)
    vsExt.TextMatrix(lngRow, 1) = NVL(rsInput!��������)
    vsExt.Cell(flexcpData, lngRow, 1) = CStr(NVL(rsInput!�Թܱ���))
    vsExt.TextMatrix(lngRow, 0) = " "
    vsExt.Cell(flexcpBackColor, lngRow, 0) = &H8000000F
    vsExt.TextMatrix(lngRow, 3) = 0 '����
    
    If mblnNewLIS Then
        strSQL = "" & vbNewLine & _
            "       Select e.Id, e.����, e.����, e.��������, e.�Թܱ���, a.���� As ���, a.Id As ��id" & vbNewLine & _
            "       From ������ĿĿ¼ a, ���鱨����Ŀ C, ���鱨����Ŀ D, ������ĿĿ¼ E" & vbNewLine & _
            "       Where a.Id = c.������Ŀid And c.������Ŀid = d.������Ŀid And d.������Ŀid = e.Id And e.�����Ŀ <> 1 And a.Id <> e.Id and a.id=[1]" & vbNewLine & _
            "       Order By ���, ����"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsInput!ID))
        Do While Not rsTmp.EOF
            i = vsExt.FindRow(CLng(rsTmp!ID))
            '�ظ���ָ�겻����
            If i = -1 Then
                If vsExt.RowData(vsExt.Rows - 1) & "" <> "" Then vsExt.AddItem ""
                vsExt.RowData(vsExt.Rows - 1) = CLng(rsTmp!ID)
                vsExt.Cell(flexcpChecked, vsExt.Rows - 1, 0) = 1
                '��������
                vsExt.TextMatrix(vsExt.Rows - 1, 2) = "    [" & rsTmp!���� & "]" & rsTmp!����
                vsExt.Cell(flexcpData, vsExt.Rows - 1, 2) = vsExt.TextMatrix(vsExt.Rows - 1, 2) '���ڻָ���ʾ
                vsExt.TextMatrix(vsExt.Rows - 1, 1) = NVL(rsTmp!��������)
                vsExt.Cell(flexcpData, vsExt.Rows - 1, 1) = CStr(NVL(rsTmp!�Թܱ���)) '����ͬ����������
    '                       If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                vsExt.TextMatrix(vsExt.Rows - 1, 3) = 1  '����
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsExt_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim strTip As String
    
    If mintType = 0 Then
        lngRow = vsExt.MouseRow: lngCol = vsExt.MouseCol
        If Between(lngRow, 0, vsExt.Rows - 1) And Between(lngCol, 0, vsExt.Cols - 1) Then
            If vsExt.Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            Else
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 - 240 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            End If
        End If
        vsExt.ToolTipText = strTip
    End If
End Sub

Private Sub vsExt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mintType = 0 Then
        If vsExt.Col = 1 And vsExt.MouseCol = 1 Then
            If x <= vsExt.CellLeft + 250 Then
                Call vsExt_KeyPress(vbKeySpace)
            End If
        End If
    End If
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�����ĳЩ�в�����༭(���¼�����BeforeEdit,��EditText��ֵ֮ǰ)
    mblnReturn = False
        
    If mintType = 0 Then
        'ֻ����ѡ���鷽��
        If Col <> 2 Then Cancel = True
    ElseIf mintType = 1 Or mintType = 4 Then
        'ֻ����༭��������
        If cmd.Visible Then cmd.Visible = False '��ʼ�༭�������ذ�ť
        If Col <> 0 And mintType = 1 Or Col <> 2 And Col <> 0 And mintType = 4 Then Cancel = True
        '����������°�LIS�������Ŀģʽ�������������
        If mblnNewLIS And mintType = 4 And Col = 2 Then
            If vsExt.TextMatrix(Row, 3) = "1" Then Cancel = True
        ElseIf mblnNewLIS And mintType = 4 And Col = 0 Then
            If Val(vsExt.TextMatrix(Row, 3)) = 0 Then Cancel = True
        End If
    End If
End Sub

Private Sub vsMethod_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 And NewRow <> -1 Then
        If vsMethod.TextMatrix(NewRow, 0) = "" Then
            Cancel = True
            vsMethod.Col = 1
        End If
    End If
End Sub

Private Sub vsMethod_Click()
    If fraMethod.Visible And vsMethod.Row >= 0 And vsMethod.Col >= 0 Then Call vsMethod_KeyPress(vbKeySpace)
End Sub

Private Sub ConfirmMethod()
'���ܣ���鷽����ȷ��
    Dim strMethod As String, i As Long
        
    With vsMethod
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strMethod = strMethod & "," & .TextMatrix(i, 1)
            End If
        Next
        vsExt.TextMatrix(vsExt.Row, 2) = Mid(strMethod, 2)
        
        '�������ú��Զ�ѡ�иò�λ
        If vsExt.TextMatrix(vsExt.Row, 2) <> "" Then
            vsExt.Cell(flexcpData, vsExt.Row, 1) = 1
            Set vsExt.Cell(flexcpPicture, vsExt.Row, 1) = img16.ListImages("c1").Picture
        End If
    End With
End Sub
    
Private Sub vsMethod_KeyPress(KeyAscii As Integer)
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMethod
        If KeyAscii = 13 Then
            Call ConfirmMethod
            fraMethod.Visible = False
            vsExt.SetFocus
        ElseIf KeyAscii = vbKeySpace Then
            '��鷽����ѡ����ȡ��
            If .Cell(flexcpData, .Row, 0) = 1 Then
                '��ѡ��ĿǰҲ����ȡ��ѡ��
                .Cell(flexcpData, .Row, 0) = 0
                Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o0", "c0")).Picture
                'ͬʱȡ���õ�ѡ�������
                If .RowData(.Row) = 1 Then
                    For i = .Row + 1 To .Rows - 1
                        If .RowData(i) = 3 Then
                            If .Cell(flexcpData, i, 0) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 1) = img16.ListImages("c0").Picture
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            Else
                blnDo = True
                If .RowData(.Row) = 3 Then
                    '����û��ѡ��ʱ,�����ѡ��
                    For i = .Row - 1 To 0 Step -1
                        If .RowData(i) <> 3 Then
                            If .Cell(flexcpData, i, 0) = 0 Then blnDo = False
                            Exit For
                        End If
                    Next
                End If
                If blnDo Then
                    .Cell(flexcpData, .Row, 0) = 1
                    Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o1", "c1")).Picture
                    If .RowData(.Row) = 1 Then '��ѡ��ѡ��ʱ��ȡ��������ѡ��
                        For i = 0 To .Rows - 1
                            If i <> .Row And .RowData(i) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 0, i, 1) = img16.ListImages("o0").Picture
                                For j = i + 1 To .Rows - 1 'ͬʱȡ���õ�ѡ�������
                                    If .RowData(j) = 3 Then
                                        If .Cell(flexcpData, j, 0) = 1 Then
                                            .Cell(flexcpData, j, 0) = 0
                                            Set .Cell(flexcpPicture, j, 1) = img16.ListImages("c0").Picture
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            
            Call ConfirmMethod
        End If
    End With
End Sub

Private Function GetCaretPos(ByVal lngHwnd As Long) As PointAPI
'���ܣ����ر༭�ؼ��е�ǰ��������
'������lngHwnd=Edit�ؼ��ľ��
'���أ�����ֵ������Edit�ؼ�,��TwipΪ��λ
'      ��������ڿؼ���Χ֮�⣬�򷵻�(-1,-1)����
    Dim lngPos As Long
    Dim vSel As CHARRANGE
    Dim vPos As PointAPI
    Dim vRect As RECT
    
    SendMessage lngHwnd, EM_EXGETSEL, 0, vSel
    lngPos = SendMessage(lngHwnd, EM_POSFROMCHAR, vSel.cpMin, 0)
    
    vPos.x = lngPos Mod 2 ^ 16
    vPos.Y = lngPos \ 2 ^ 16
    
    '����Χ�ж�
    GetWindowRect lngHwnd, vRect
    If vPos.x >= 0 And vPos.x <= vRect.Right - vRect.Left + 1 _
        And vPos.Y >= 0 And vPos.Y <= vRect.Bottom - vRect.Top + 1 Then
        vPos.x = vPos.x * Screen.TwipsPerPixelX
        vPos.Y = vPos.Y * Screen.TwipsPerPixelY
    Else
        vPos.x = -1: vPos.Y = -1
    End If
    
    GetCaretPos = vPos
End Function

Private Sub SetControlFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'���ܣ����ô��弰���пؼ��������С
'������frmMe=��Ҫ��������Ĵ������
'      bytSize:����Ϊ9������,0:����Ϊ9������,1,����Ϊ12������
'      strOther:�������������õĿؼ��������ļ���,��ʽΪ����������1,��������2,��������3,....
'˵����1.����漰��VsFlexGrid�ȱ��ؼ�����Ҫ�������ڵĻ������µ����п���и�
'      2.�������δ�г��������ؼ����Զ���ؼ�,��Ҫ���ض�����ָ�������С����ش���ģ������ⵥ������

    Dim objCtrol As Control, objrptCol As ReportColumn
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    
    lngFontSize = IIF(bytSize = 0, 9, IIF(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '����CommandBars�û��Զ���ؼ���ȡobjCtrol.Container�����
            On Error Resume Next
            If InStr(1, strOther, "," & objCtrol.Container.Name & ",") > 0 Then
                 blnDo = False
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = frmMe.TextHeight("��") + 20
                        'Label�����Ҫ���е���
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("����" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("��") + IIF(bytSize = 0, 100, 120)
                Case "TextBox", "RichTextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("��")
                        If objCtrol.Name = "txtSentence" Then
                            imgSentence.Width = imgSentence.Width * dblRate
                            imgSentence.Height = imgSentence.Height * dblRate
                            imgSentence.Left = objCtrol.Width + objCtrol.Left
                            picSentence.Width = picSentence.Width * dblRate
                            picSentence.Height = picSentence.Height * dblRate
                        End If
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                        objCtrol.Height = frmMe.TextHeight("��")
                Case "ReportControl"
                        lngOldSize = objCtrol.PaintManager.TextFont.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        Set CtlFont = objCtrol.PaintManager.TextFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.TextFont = CtlFont
                        For Each objrptCol In objCtrol.Columns
                            objrptCol.Width = objrptCol.Width * dblRate
                        Next
                        objCtrol.Redraw
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                        
            End Select
        End If
    Next
End Sub
