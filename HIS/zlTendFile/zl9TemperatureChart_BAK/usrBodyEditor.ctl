VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl usrBodyEditor 
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8100
   ScaleWidth      =   10455
   Begin VB.PictureBox picSerach 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1515
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1515
      Begin VB.Label lbl�鿴 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ԭʼ��С"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   32
         Top             =   60
         Width           =   960
      End
      Begin VB.Image imgPic 
         Height          =   360
         Left            =   60
         Picture         =   "usrBodyEditor.ctx":0000
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7095
      ScaleWidth      =   10215
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   10215
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   9600
         TabIndex        =   29
         Top             =   4920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   7200
         TabIndex        =   28
         Top             =   6120
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   240
         ScaleHeight     =   360
         ScaleWidth      =   2730
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   6120
         Width           =   2730
         Begin VB.ComboBox cboBaby 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   0
            Width           =   1920
         End
         Begin VB.Label lblSerach 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�鿴"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   9
            Left            =   345
            TabIndex        =   24
            Top             =   105
            Width           =   360
         End
      End
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   9375
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   9375
         Begin VB.PictureBox picDraw 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   7335
            TabIndex        =   33
            Top             =   2160
            Width           =   7335
         End
         Begin zl9TemperatureChart.VsfGrid vsf 
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3480
            Width           =   7215
            _ExtentX        =   12515
            _ExtentY        =   450
         End
         Begin VB.PictureBox picCard 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   810
            Index           =   0
            Left            =   120
            ScaleHeight     =   810
            ScaleWidth      =   8640
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   120
            Width           =   8640
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   7
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   10
               TabStop         =   0   'False
               Text            =   "���"
               Top             =   375
               Width           =   2370
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   6
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "����"
               Top             =   60
               Width           =   645
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   5
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   8
               TabStop         =   0   'False
               Text            =   "�Ա�"
               Top             =   60
               Width           =   420
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   4
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   7
               TabStop         =   0   'False
               Text            =   "12"
               Top             =   375
               Width           =   615
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   3
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   6
               TabStop         =   0   'False
               Text            =   "��Ժ����"
               Top             =   60
               Width           =   1140
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   2
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   5
               TabStop         =   0   'False
               Text            =   "����"
               Top             =   375
               Width           =   2400
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   1
               Left            =   6645
               Locked          =   -1  'True
               TabIndex        =   4
               TabStop         =   0   'False
               Text            =   "1234567"
               Top             =   60
               Width           =   3825
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   0
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "������"
               Top             =   60
               Width           =   1425
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��    ��:"
               Height          =   180
               Index           =   7
               Left            =   4065
               TabIndex        =   18
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   6
               Left            =   2910
               TabIndex        =   17
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ա�:"
               Height          =   180
               Index           =   4
               Left            =   1980
               TabIndex        =   16
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����:"
               Height          =   180
               Index           =   5
               Left            =   4050
               TabIndex        =   15
               Top             =   60
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   3
               Left            =   2910
               TabIndex        =   14
               Top             =   390
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   13
               Top             =   375
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ��:"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   12
               Top             =   60
               Width           =   630
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   11
               Top             =   60
               Width           =   450
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid mshDownTab 
            Height          =   975
            Left            =   90
            TabIndex        =   20
            Top             =   3840
            Width           =   7215
            _cx             =   12726
            _cy             =   1720
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   4
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"usrBodyEditor.ctx":076A
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
            OwnerDraw       =   1
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
         Begin VSFlex8Ctl.VSFlexGrid mshUpTab 
            Height          =   1095
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   7275
            _cx             =   12832
            _cy             =   1931
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
            OwnerDraw       =   1
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
            Begin VB.PictureBox picDisplay 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   240
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   60
               Width           =   165
               Begin VB.Image imgDisPlay 
                  Appearance      =   0  'Flat
                  Height          =   240
                  Left            =   -30
                  Picture         =   "usrBodyEditor.ctx":08D7
                  Stretch         =   -1  'True
                  Top             =   -30
                  Width           =   240
               End
            End
            Begin VB.Label lblCur 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   240
               TabIndex        =   27
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Label lblCommText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "˵��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   240
            TabIndex        =   21
            Top             =   4920
            Width           =   360
         End
      End
      Begin VB.PictureBox picBuffer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   7200
         ScaleHeight     =   117
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   139
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "��ʱ��ͼ��,ǧ���ɾ"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrBodyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrTitle As String = "���»�ͼ"
'--�˵�
Private mcbrToolBarҳ�� As CommandBarControl
Private mcbrTools   As CommandBar
Private mcbrToolBar As CommandBar
Private mcbrItem         As CommandBarControl

'--����
Public mblnResize As Boolean '��¼�����С�Ƿ����仯
Public mblnMoved As Boolean
Private mlngWidth As Long
Private mlngHeight As Long
Private mintPage      As Integer '��¼��ǰҳ��
Private mintAllPage As Integer '���µ�����ҳ��
Private mstrParam  As String, mstrParam1 As String, mstrParam2 As String
Private mfrmParent As Object
Private mIntDataEditor As Integer '0 ��ʾ�������µ����ݱ༭ 1��ʾ�������µ�������ʾ����
Private mstrSQL As String
Private mintColMin, mintColMax As Integer
Private mint����Ӧ�� As Double
Private msinVStep As Single      '�������Ĳ���
Private msinHStep As Single      '�������Ĳ���
Private mblnAutoAdjust As Boolean  '�������µ���ʽ �����µ��Ƿ���洰���С�Զ�����
Private mblnAutoRedraw As Boolean  '�����Ƿ��Զ��ػ�:�Ƿ��Զ�����ػ�,���ݰ���:��ʼ����,��ȡ���ݲ�����,�滭,
Private mblnRefresh    As Boolean
Private mintOpDays As Integer '������־����
Private mblnStopFlag As Boolean '����ֹͣ��־
Private mintOpFormat As Integer '��������ȱʡ��ʽ 0-����ʾ;1-��ʾ0;2-��ʾ��������
Private mintRepairRows As Integer '���±��̶��������
Private mbln��ʾƤ�� As Boolean
Private mblnKeyDown As Boolean
Private mstrOpdays(1 To 7) As String
Private mstrOpValue(1 To 7) As String
Private mstrNewString() As String '����Ƥ�Խ����Ϣ
Private mlng�߶� As Single '���µ��̶��������ʾ�ĸ߶ȷ�Χ
Private mbln��Ժ As Boolean '�����Ƿ��Ժ
Private mbytSize As Byte '�����С 0-9������ 1-12������
'--���µ�ʱ��
Private mstr��ʼʱ�� As String  'һ�ܿ�ʼʱ��
Private mstr����ʱ�� As String  'һ�ܽ���ʱ��
Private mstrEnterDate As String '���µ���ʼʱ��
Private mstrEndDate As String   '���µ�����ʱ��
Private mstrComeInDate As String '������Ժʱ��

'����
Private mbln�೦�����ӷ�ĸ��ʾ As Boolean

Private WithEvents mfrmCaseTendBodyPrint As frmCaseTendBodyPrint
Attribute mfrmCaseTendBodyPrint.VB_VarHelpID = -1

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2 'ǳ����
Private Const BDR_RAISEDINNER = &H4 'ǳ͹��
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)


'***************************************************************
'���µ��滭��ر���
'***************************************************************
Private mobjDraw          As Object
Private mobjBuffer        As Object                     '����,����ʹ��
Private mblnRedraw        As Boolean                    '�Ƿ���Ҫ�ػ�
Private mlngHwnd          As Long
Private mlngDC            As Long
Private mlngMemDC         As Long
Private mlngBitmap        As Long
Private mlngOldBitmap     As Long
Private mlngMemBitmap     As Long
Private mlngPen           As Long
Private mlngBrush         As Long
Private mlngOldPen        As Long
Private mlngOldBrush      As Long
Private mlngFont           As Long
Private mlngOldFont       As Long


'***************************************************************
'���µ��滭����ڴ�ӳ���¼��
'***************************************************************
Private mrsItems As New ADODB.Recordset              '������Ŀ������
Private mrsGraph As New ADODB.Recordset              '�����������ͼ�����(ȫ����ȡ��picBuffer��,�˴��������Ŀ�Ĳ�λ�����Ӧ��ͼ�����)
Private mrsDrawItems As New ADODB.Recordset              '����������Ŀ����Ч��������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
Private mrsPoint As New ADODB.Recordset              '���е�ı��ּ���
Private mrsNote  As New ADODB.Recordset              '�ı��������,��ָ����ɫ

Private Type Type_NO
    Ѫѹ As Integer
    ����ѹ As Integer
    ���� As Integer
End Type

Private mItemNO As Type_NO

Private Type Type_row
    Ѫѹ As Integer
    ������ As Integer
    �ų��� As Integer
End Type
Private mItemRow As Type_row

Private Type T_LPoint
    X As Long
    Y As Long
    W As Single
End Type

'***************************************************************
'���˻�����Ϣ
'***************************************************************
Private Type type_Patient
    lng����ID As Long
    lng��ҳID As Long
    lng����ID As Long
    lng����ID As Long
    lng��Ժ As Long
    lngӤ�� As Long
    lng�༭ As Long
    lng����ȼ� As Long
    lng�ļ�ID As Long
    lngԭʼ��С As Long
    lngPage As Long
End Type
Private T_Patient As type_Patient

'--�¼�����
Public Event CmdClick(ByVal strParam As String)
Public Event zlAfterPrint()
Public Event DbClickCur(ByVal intDataEditor As Integer)

Public Property Get ParentForm() As Object
    Set ParentForm = mfrmParent
End Property

Public Property Set ParentForm(objParent As Object)
    Set mfrmParent = objParent
End Property

Public Property Get ScrollBarY() As FlatScrollBar
    Set ScrollBarY = vsb
End Property

Public Property Get ScrollBarX() As FlatScrollBar
    Set ScrollBarX = hsb
End Property

Public Property Get DateEditor() As Integer
     DateEditor = mIntDataEditor
End Property

Public Property Let DateEditor(intDataEditor As Integer)
     mIntDataEditor = intDataEditor
End Property

Public Property Let lng����ID(lng����ID As Long)
     T_Patient.lng����ID = lng����ID
End Property

Public Property Let lng��ҳID(lng��ҳID As Long)
     T_Patient.lng��ҳID = lng��ҳID
End Property

Public Property Let lng�ļ�ID(lng�ļ�ID As Long)
     T_Patient.lng�ļ�ID = lng�ļ�ID
End Property

Public Property Let lng����ID(lng����ID As Long)
     T_Patient.lng����ID = lng����ID
End Property

Public Property Let lngӤ��(lngӤ�� As Long)
     T_Patient.lngӤ�� = lngӤ��
End Property

Public Property Let intPage(intPage1 As Long)
     mintPage = intPage1
End Property

Public Property Get intPage() As Long
    intPage = mintPage + 1
End Property

Public Property Get AllPage() As Integer
    AllPage = mintAllPage
End Property

Public Property Get FontSize() As Byte
    FontSize = mbytSize
End Property

Public Property Let FontSize(bytSize As Byte)
     mbytSize = bytSize
End Property

Private Function InitCommandBar() As Boolean

    '******************************************************************************************************************
    '���ܣ���ʼ���˵���ť
    '������
    '���أ�
    '******************************************************************************************************************

    Dim objCustom  As CommandBarControlCustom
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo Errhand
    '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
'    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003

    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CallPrevious, "��һҳ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CallNext, "��һҳ")
    End With

    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    
    Set mcbrToolBar = cbsMain.Add("Ӥ��", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    Set objCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Option, "")
    objCustom.flags = xtpFlagAlignLeft
    picTmp.Visible = True
    objCustom.Handle = picTmp.hWnd
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyLeft, conMenu_Manage_CallPrevious      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_Manage_CallNext    '��һ��
    End With

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function InitBody(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As Boolean

    '******************************************************************************************************************
    '���ܣ���ȡ���˻���ʱ�䷶Χ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSql        As String, strNewSql As String
    Dim strParam1 As String '���ָ����ҳ�� ������¼��Ӧҳ�ŵĲ���ֵ
    Dim RS            As New ADODB.Recordset
    Dim rsTmp         As New ADODB.Recordset
    Dim ArrControlId() As Variant
    Dim cbrPre  As CommandBarButton
    Dim cbrWeek As CommandBarButton
    Dim objCostom As CommandBarControlCustom
    Dim intCount      As Integer
    Dim strDateFrom   As String 'ÿһҳ ��ʼʱ��
    Dim strDateTo     As String 'ÿһҳ ����ʱ��
    Dim strEnterDate  As String  '��Ժʱ��
    Dim strOutDate  As String   '��ֹʱ��
    Dim strMarkDate As String '���µ�����ʱ��
    Dim intCOl        As Integer
    Dim strCaption    As String
    Dim strParameter  As String
    Dim strSvrCaption As String, strSvrCaption1 As String
    Dim strNow        As String
    Dim strCut        As String
    Dim lngLoop       As Long
    Dim strTmp        As String
    Dim lnglast����id As Long
    
    On Error GoTo Errhand

    If lng����ID = 0 And lng�ļ�ID = 0 And lng��ҳID = 0 Then Exit Function
    mbln��Ժ = False
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    'ɾ������ҳ��˵���

    If Not mcbrToolBarҳ�� Is Nothing Then mcbrToolBarҳ��.Delete
    Set mcbrToolBarҳ�� = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewItem, "ҳ��"):  mcbrToolBarҳ��.BeginGroup = True
    mcbrToolBarҳ��.IconId = conMenu_Edit_Modify
    mcbrToolBarҳ��.Style = xtpButtonIconAndCaption
    
    
    ArrControlId = Array(conMenu_View_OneWeek, conMenu_View_TwotWeek, conMenu_View_ThreeWeek, conMenu_View_FourWeek, _
        conMenu_View_Forward, conMenu_View_Backward)

    For lngLoop = 0 To UBound(ArrControlId)
        If Not mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))) Is Nothing Then mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))).Delete
    Next lngLoop
    
    
    With mcbrToolBar.Controls
        
        '��������ҳ
        Set cbrPre = .Add(xtpControlButton, conMenu_View_Forward, "��һҳ", -1, False)
        Set cbrPre = .Add(xtpControlButton, conMenu_View_Backward, "��һҳ", -1, False)
        
        '���� 4������ �˴�Ĭ��Ϊ��Ժʱ�俪ʼ4����
        Set cbrPre = .Add(xtpControlButton, conMenu_View_OneWeek, " " & 1 & " ", -1, False)
        cbrPre.ToolTipText = "��һ��"
        Set cbrPre = .Add(xtpControlButton, conMenu_View_TwotWeek, " " & 2 & " ", -1, False)
        cbrPre.ToolTipText = "�ڶ���"
        Set cbrPre = .Add(xtpControlButton, conMenu_View_ThreeWeek, " " & 3 & " ", -1, False)
        cbrPre.ToolTipText = "������"
        Set cbrPre = .Add(xtpControlButton, conMenu_View_FourWeek, " " & 4 & " ", -1, False)
        cbrPre.ToolTipText = "������"
    End With
    
    If Not mcbrToolBar.FindControl(, conMenu_ViewPopup) Is Nothing Then
        mcbrToolBar.FindControl(, conMenu_ViewPopup).Delete
    End If
    
    If mblnAutoAdjust = True Then
        Set objCostom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_ViewPopup, "�鿴ԭͼ")
        objCostom.Handle = picSerach.hWnd
    Else
        picSerach.Visible = False
    End If
    
    '��ȡ�û����õ����µ���ʼʱ��(Ӥ��������Ӥ������ʱ��Ϊ׼)
    strSql = "select ��ʼʱ�� from ���˻����ļ� where ID=[1] and ����ID=[2] and ��ҳid=[3] and nvl(Ӥ��,0)=[4]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ���µ���ʼʱ��", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��)
    If rsTmp.RecordCount <> 0 Then
        strEnterDate = Format(rsTmp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
    End If
    
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ)����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "   (SELECT /*+ RULE */  ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0 AND C.��� = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.�������� = COLUMN_VALUE) And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"

    strMarkDate = "to_date('" & strEnterDate & "','yyyy-MM-dd hh24:mi:ss')"
    '------------------------------------------------------------------------------------------------------------------
    '��ȡ���µ�ҳ����Ӥ����Ժʱ�������Ƿ����ҽ��,������ҽ��ʱ��Ϊ׼��������ĸ�׵�Ϊ׼��
    strSql = "SELECT DECODE(C.����ʱ��,NULL," & IIf(strEnterDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��) AS ��Ժʱ��," & vbNewLine & _
                " DECODE(C.����ʱ��,NULL,B.��Ժʱ��,C.����ʱ��) AS ʵ����Ժʱ��," & vbNewLine & _
                " DECODE(E.��¼,0,DECODE(SIGN(nvl(E.Ӥ��ʱ��,B.��Ժʱ��) - D.����ʱ��), 1,nvl(E.Ӥ��ʱ��,B.��Ժʱ��) ,D.����ʱ��),nvl(E.Ӥ��ʱ��,B.��Ժʱ��))  ��Ժʱ��," & vbNewLine & _
                " 1 + TRUNC((TO_DATE(TO_CHAR(DECODE(E.��¼,0,DECODE(SIGN(nvl(E.Ӥ��ʱ��,B.��Ժʱ��) - D.����ʱ��), 1,nvl(E.Ӥ��ʱ��,B.��Ժʱ��) ,D.����ʱ��),nvl(E.Ӥ��ʱ��,B.��Ժʱ��)),'yyyy-MM-dd'),'yyyy-MM-dd') - " & vbNewLine & _
                " TO_DATE(TO_CHAR(DECODE(C.����ʱ��,NULL," & IIf(strEnterDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��),'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS ҳ��,D.����ʱ��,E.��¼" & vbNewLine & _
                "    FROM (SELECT ����ID,��ҳID,MIN(��ʼʱ��) AS ��Ժʱ��,MAX(NVL(��ֹʱ��, SYSDATE)) AS ��Ժʱ��" & vbNewLine & _
                "    FROM ���˱䶯��¼" & vbNewLine & _
                "    WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID =[3] GROUP BY ����ID,��ҳID) B," & vbNewLine & _
                "    (SELECT ����ID,��ҳID,����ʱ�� FROM ������������¼ WHERE ����ID = [2] AND ��ҳID = [3] AND ���=[4]) C," & vbNewLine & _
                "    (SELECT NVL(����ʱ��,SYSDATE) ����ʱ�� FROM (SELECT MAX(����ʱ��) ����ʱ�� FROM ���˻����ļ� A,���˻������� B" & vbNewLine & _
                "           WHERE A.ID=B.�ļ�ID AND A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4])) D," & vbNewLine & _
                strNewSql & vbNewLine & _
                "    WHERE B.����ID=E.����ID And B.��ҳID=E.��ҳID And B.����ID=C.����ID(+) AND B.��ҳID=C.��ҳID(+)"

    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "usrBodyEditor", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��)
    If rsTmp.BOF Then
        MsgBox "�޲��˱���סԺ��¼��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mintAllPage = rsTmp("ҳ��").Value
    If T_Patient.lngPage > mintAllPage Then T_Patient.lngPage = 0
    
    If strEnterDate = "" Then strEnterDate = Format(rsTmp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
    mstrEnterDate = strEnterDate
    mstrComeInDate = Format(rsTmp!ʵ����Ժʱ��, "yyyy-MM-dd HH:mm:ss")
    
    strOutDate = Format(rsTmp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
    mbln��Ժ = Not (Val(zlCommFun.Nvl(rsTmp!��¼)) = 0)
    
    '------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT 1 + TRUNC((TO_DATE(TO_CHAR(A.��ʼʱ��,'yyyy-MM-dd'),'yyyy-MM-dd') - TO_DATE(TO_CHAR(B.��Ժʱ��,'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS ��ʼҳ��," & vbNewLine & _
            "1 + TRUNC((TO_DATE(TO_CHAR(DECODE(A.���,F.LAST,DECODE(E.��¼,0,DECODE(SIGN(nvl(E.Ӥ��ʱ��,A.��ֹʱ��) - D.����ʱ��), 1,nvl(E.Ӥ��ʱ��,A.��ֹʱ��) ,D.����ʱ��),nvl(E.Ӥ��ʱ��,A.��ֹʱ��)),nvl(E.Ӥ��ʱ��,A.��ֹʱ��)),'yyyy-MM-dd'),'yyyy-MM-dd') - TO_DATE(TO_CHAR(B.��Ժʱ��,'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS ����ҳ��," & vbNewLine & _
            "      B.��Ժʱ��,D.����ʱ��,����ID,C.����,A.��ʼʱ��,DECODE(A.���,F.LAST,DECODE(E.��¼,0,DECODE(SIGN(nvl(E.Ӥ��ʱ��,A.��ֹʱ��) - D.����ʱ��), 1,nvl(E.Ӥ��ʱ��,A.��ֹʱ��) ,D.����ʱ��),nvl(E.Ӥ��ʱ��,A.��ֹʱ��)),nvl(E.Ӥ��ʱ��,A.��ֹʱ��))  ��ֹʱ��" & vbNewLine & _
            "FROM (SELECT ROWNUM ���, ����ID,��ʼʱ��,��ֹʱ��" & vbNewLine & _
            "      FROM(SELECT  ����ID,MIN(��ʼʱ��) AS ��ʼʱ��,MAX(NVL(��ֹʱ��, SYSDATE)) AS ��ֹʱ��" & vbNewLine & _
            "           FROM ���˱䶯��¼" & vbNewLine & _
            "                WHERE ��ʼʱ�� IS NOT NULL AND ����ID =[2] AND ��ҳID =[3] GROUP BY ����ID  ORDER BY ��ʼʱ��)) A," & vbNewLine & _
            "      (SELECT DECODE(Y.����ʱ��,NULL,X.��Ժʱ��,Y.����ʱ��) AS ��Ժʱ��,X.����ID,X.��ҳID FROM (SELECT ����ID,��ҳID,MIN(��ʼʱ��) AS ��Ժʱ��" & vbNewLine & _
            "      FROM ���˱䶯��¼" & vbNewLine & _
            "      WHERE ��ʼʱ�� IS NOT NULL AND ����ID =[2] AND ��ҳID =[3] GROUP BY ����ID,��ҳID) X," & vbNewLine & _
            "      (SELECT ����ID,��ҳID,����ʱ�� FROM ������������¼ WHERE ����ID =[2] AND ��ҳID =[3] AND ���=[4]) Y" & vbNewLine & _
            "      WHERE X.����ID=Y.����ID(+) AND X.��ҳID=Y.��ҳID(+) ) B,���ű� C ," & vbNewLine & _
            "      (SELECT NVL(����ʱ��,SYSDATE) ����ʱ�� FROM (SELECT MAX(����ʱ��) ����ʱ�� FROM ���˻����ļ� A,���˻������� B" & vbNewLine & _
            "      WHERE A.ID=B.�ļ�ID AND A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4])) D," & vbNewLine & _
            strNewSql & "," & vbNewLine & _
            "      (SELECT  COUNT(*) LAST FROM" & vbNewLine & _
            "      (SELECT ����ID FROM ���˱䶯��¼" & vbNewLine & _
            "                WHERE ��ʼʱ�� IS NOT NULL AND ����ID =[2] AND ��ҳID = [3] GROUP BY ����ID )) F" & vbNewLine & _
            "WHERE B.����ID=E.����ID And B.��ҳID=E.��ҳID And C.ID(+)=A.����ID" & vbNewLine & _
            "ORDER BY A.��ʼʱ��"
            
    Set RS = zldatabase.OpenSQLRecord(strSql, "usrBodyEditor", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��)
    
    For lngLoop = 0 To rsTmp("ҳ��").Value - 1

        strDateFrom = Format(rsTmp("��Ժʱ��").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"

        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then

            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")

            RS.Filter = ""
            RS.Filter = "��ʼҳ��<=" & lngLoop + 1 & " And ����ҳ��>=" & lngLoop + 1

            If RS.RecordCount > 0 Then RS.MoveFirst

            For intCOl = 1 To RS.RecordCount

                If strDateFrom < Format(RS("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strTmp = Format(RS("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strTmp = strDateFrom
                End If

                If strDateTo > Format(RS("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strCaption = Format(RS("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strCaption = strDateTo
                End If

                strCaption = Format(strTmp, "yyyy-MM-dd") & "��" & Format(strCaption, "yyyy-MM-dd")
                strCaption = "��" & lngLoop + 1 & "ҳ��" & strCaption & "(" & RS("����").Value & ")"

                '��Ժʱ��;����id;��ʼʱ��;����ʱ��;
                Set mcbrItem = mcbrToolBarҳ��.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
                mcbrItem.Parameter = strEnterDate & ";" & RS!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                
                If lngLoop + 1 <= 4 Then
                    Set cbrWeek = mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop)))
                    cbrWeek.Parameter = strEnterDate & ";" & RS!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                End If
                 
                lnglast����id = Val(Nvl(RS("����ID").Value))

                RS.MoveNext

                strParameter = mcbrItem.Parameter
                
                'ָ��ҳ�Ų�Ϊ0 ���Һ͸�ҳ����Ⱦͼ�¼����ֵ
                If T_Patient.lngPage <> 0 And Val(T_Patient.lngPage - 1) = lngLoop Then
                    strParam1 = strParameter
                    strSvrCaption1 = strCaption
                End If
                
                strSvrCaption = strCaption
            Next
        Else
            mintAllPage = lngLoop
            Exit For
        End If
    Next
    
    '������Ժ��̶�ǰ���ܵ�״̬
    For lngLoop = 0 To 3
        If mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))).Parameter = "" Then mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))).Enabled = False
    Next lngLoop
    
    'ҳ�Ų�Ϊ�վͰ�ָ��ҳ����ʾ
    If strParam1 <> "" Then strParameter = strParam1: strSvrCaption = strSvrCaption1
    
    '������һҳ��һҳ״̬
    Call InitWeekDays(strParameter)

    If strParameter <> "" Then
        mstrParam = strParameter
        mcbrToolBarҳ��.Caption = strSvrCaption
        Call zlMenuClick("װ������", mstrParam)
    End If
    
    cbsMain.RecalcLayout
    
    InitBody = True

    Exit Function

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub InitWeekDays(ByVal strParameter As String)

    '-----------------------------------------------------------------------------------
    '����:�����ж��ٸ���һҳ����һҳ
    '����:strParameter '��Ժʱ��;����id;��ʼʱ��;����ʱ��;ҳ��;��ֹʱ��
    '------------------------------------------------------------------------------------
    Dim ArrCode      As Variant

    Dim strBeginTime As String, strEndTime As String, strDateFrom As String

    Dim cbrMunu      As CommandBarButton
    
    Dim lngLoop As Long
    
    Dim lngPage As Long

    ArrCode = Split(strParameter, ";")
    
    
    On Error GoTo Errhand
    
    If CDate(Format(CStr(ArrCode(0)), "yyyy-MM-dd HH:mm:ss")) > CDate(Format(CStr(ArrCode(5)), "yyyy-MM-dd HH:mm:ss")) Then ArrCode(0) = Format(ArrCode(5), "yyyy-MM-dd HH:mm:ss")
    
    
    For lngLoop = 0 To Round((DateDiff("D", CDate(ArrCode(0)), CDate(ArrCode(5))) + 1) / 7)

        strDateFrom = Format(CDate(ArrCode(0)) + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"

        If strDateFrom < Format(ArrCode(0), "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(ArrCode(0), "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(ArrCode(5), "yyyy-MM-dd HH:mm:ss") Then
            lngPage = lngLoop
        End If
    Next lngLoop

    With mcbrToolBar.Controls
        
        '���������²���.
        Set cbrMunu = .Find(, conMenu_View_Forward) '��һҳ
        cbrMunu.Parameter = ArrCode(4)  '��Ż��м�����һҳ
        
        Set cbrMunu = .Find(, conMenu_View_Backward) '��һ��
        cbrMunu.Parameter = Val(lngPage - Val(ArrCode(4)))  '���м�����һ��

    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    Dim RS As New ADODB.Recordset
    Dim varParam   As Variant
    Dim blnRefresh As Boolean
    Dim strStartDate As String
    Dim strEndDate   As String
    Dim intCOl  As Long
    Dim strTime As String, strInput As String
    
    On Error GoTo Errhand
    If Trim(strParam) <> "" Then varParam = Split(strParam, ";")
    
    Select Case strMenuItem
        Case "��ʼ��"
            mstrParam1 = "": mstrParam2 = ""
            mstrParam1 = strParam
            mbln��Ժ = False
            picMain.Tag = ""
            'strParam ����ID;��ҳID;����ID;�ļ�ID;��Ժ;�༭;Ӥ��;����ȼ�;�Ƿ���ݾߴ�����Զ�У�����µ���ʽ(1 �� 0 ��)
            T_Patient.lng����ID = varParam(0)
            T_Patient.lng��ҳID = varParam(1)
            T_Patient.lng����ID = varParam(2)
            T_Patient.lng����ID = varParam(2)
            T_Patient.lng�ļ�ID = varParam(3)
            T_Patient.lngPage = 0
            
            If UBound(varParam) > 3 Then
                T_Patient.lng��Ժ = varParam(4)
            Else
                T_Patient.lng��Ժ = 0
            End If
            
            If UBound(varParam) > 4 Then
                T_Patient.lng�༭ = varParam(5)
            Else
                T_Patient.lng�༭ = 0
            End If
            
            If UBound(varParam) > 5 Then
                T_Patient.lngӤ�� = varParam(6)
            Else
                T_Patient.lngӤ�� = 0
            End If
            
            If UBound(varParam) > 6 Then
                T_Patient.lng����ȼ� = varParam(7)
            Else
                T_Patient.lng����ȼ� = 3
            End If
            
            If UBound(varParam) > 7 Then
                T_Patient.lngԭʼ��С = Val(varParam(8))
            Else
                T_Patient.lngԭʼ��С = 0
            End If
            
            If UBound(varParam) > 8 Then
                T_Patient.lngPage = Val(varParam(9))
            End If
            
            mblnAutoAdjust = IIf(T_Patient.lngԭʼ��С = 1, False, True)
            mblnRedraw = False
            mblnRefresh = True
            
            mstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
            Set RS = zldatabase.OpenSQLRecord(mstrSQL, "��ȡ����ID", T_Patient.lng����ID, T_Patient.lng��ҳID)
            If RS.BOF = False Then
                T_Patient.lng����ID = Val(zlCommFun.Nvl(RS("��Ժ����ID").Value))
            End If

            mstrSQL = "SELECT A.���,A.���� FROM(" & vbNewLine & _
                        "SELECT A.���,A.����,A.����ID,A.��ҳID FROM (SELECT 0 ���, B.����,A.����ID,A.��ҳID" & vbNewLine & _
                        "            FROM ������ҳ A, ������Ϣ B" & vbNewLine & _
                        "            WHERE A.����ID = B.����ID AND A.����ID =[1] AND A.��ҳID =[2]" & vbNewLine & _
                        "            UNION ALL" & vbNewLine & _
                        "            SELECT A.���, DECODE(A.Ӥ������, NULL, B.���� || '֮��' || TRIM(TO_CHAR(A.���, '9')), A.Ӥ������) AS ����,A.����ID,A.��ҳID" & vbNewLine & _
                        "            FROM ������������¼ A, ������Ϣ B" & vbNewLine & _
                        "            WHERE A.����ID =[1] AND A.��ҳID =[2] AND A.����ID = B.����ID) A," & vbNewLine & _
                        "            (SELECT A.����ID,A.��ҳID , NVL(A.Ӥ��,0) Ӥ�� FROM ���˻����ļ� A,�����ļ��б� B" & vbNewLine & _
                        "            WHERE A.��ʽID=B.ID AND B.����=3 AND B.����=-1) B" & vbNewLine & _
                        "            WHERE A.����ID=B.����ID AND A.��ҳID=B.��ҳID AND A.���=B.Ӥ��) A" & vbNewLine & _
                        "ORDER BY A.���"
            Set RS = zldatabase.OpenSQLRecord(mstrSQL, mstrTitle, T_Patient.lng����ID, T_Patient.lng��ҳID)
            
            cboBaby.Clear
            If RS.BOF = False Then
                Do While Not RS.EOF
                    cboBaby.AddItem RS("����").Value
                    cboBaby.ItemData(cboBaby.NewIndex) = RS("���").Value
                    RS.MoveNext
                    If cboBaby.ListIndex = -1 And T_Patient.lngӤ�� = Val(cboBaby.ItemData(cboBaby.NewIndex)) Then
                        Call zlControl.CboSetIndex(cboBaby.hWnd, cboBaby.NewIndex)
                        T_Patient.lngӤ�� = cboBaby.ItemData(cboBaby.ListIndex)
                    End If
                Loop
            End If
            
            If cboBaby.ListCount > 0 And cboBaby.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboBaby.hWnd, 0)
                T_Patient.lngӤ�� = cboBaby.ItemData(cboBaby.ListIndex)
            End If
            
            '��ʼ����
            Call Paint_Init(picDraw, picBuffer)
            
            If Not InitData(T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng��Ժ, T_Patient.lng�༭, T_Patient.lngӤ��) Then Exit Function
            If Not InitBody(T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��) Then Exit Function
            Call ReSetFontSize
        Case "װ������"
            'strParam��ʽ����ʼʱ��;����ID;��ʼʱ��;����ʱ��;ҳ��
            
            'Debug.Print Now & ":װ������"
            mstrParam2 = strParam
            mblnRedraw = True
            mbln�������� = True
            mstrEnterDate = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
            strStartDate = Format(varParam(2), "YYYY-MM-DD HH:mm:ss")
            strEndDate = Format(varParam(3), "YYYY-MM-DD HH:mm:ss")
            mintPage = Val(varParam(4))
            glngCurPage = mintPage + 1
            mstrEndDate = Format(varParam(5), "YYYY-MM-DD HH:mm:ss")
            If mbln��Ժ = True Then
                '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
                mstrEndDate = Format(RetrunEndTime(CDate(mstrEnterDate), CDate(mstrEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
                strEndDate = Format(RetrunEndTime(CDate(mstrEnterDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
            End If
            If strStartDate & ";" & strEndDate = picMain.Tag Then
                mblnRefresh = False
            Else
                mblnRefresh = True
            End If
            
            picMain.Tag = strStartDate & ";" & strEndDate
                        
            mstr��ʼʱ�� = strStartDate
            mstr����ʱ�� = strEndDate
            
            If mstr��ʼʱ�� = "" Or mstr����ʱ�� = "" Then
                Call FaceInitTable(False)
                Call picDraw_Paint '���ڴ���Copy������PIC
            Else
                If mblnRefresh = True Then
                    Call ReadBodyInfo '���ز��˻�����Ϣ
                    'Debug.Print Now & ":��ʼ�����±��"
                    Call FaceInitTable '��ʼ�����±��
                    'Debug.Print Now & ":���ز�����������"
                    Call ReadBoyData(mblnAutoAdjust) '���ز�����������
                    'Debug.Print Now & ":��ʼ���ͼ��"
                    Call Paint_Construct   '������ߺ�ͼ��
                    Call Paint_Assistant '����ϱ�,�±�,δ��˵��,��Ժ��ר�ƣ�����Ϣ
                    Call picDraw_Paint '���ڴ���Copy������PIC
                    'Debug.Print Now & ":���ر������"
                    Call ShowDowntab '�����±������
                End If
            End If
            
            '���ٴ�����������Ϣ
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            mlngOldFont = 0: mlngFont = 0
            
            mlngWidth = UserControl.Width
            mlngHeight = UserControl.Height
            
            'Debug.Print Now & ":װ������Over"
        Case "��ʾ������Ϣ"
            If T_Patient.lngӤ�� = 0 Then
                txtCard(0).Text = txtCard(0).Tag
                txtCard(7).Text = txtCard(7).Tag
            Else
                txtCard(5).Text = ""
                txtCard(6).Text = ""
                txtCard(7).Text = ""
                
                mstrSQL = "Select Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,a.Ӥ���Ա�,a.����ʱ�� From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id And a.���=[3]"
                Set RS = zldatabase.OpenSQLRecord(mstrSQL, "��ȡӤ����Ϣ", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
                If RS.BOF = False Then
                    txtCard(0).Text = RS("Ӥ������").Value
                    txtCard(5).Text = RS("Ӥ���Ա�").Value
                    txtCard(6).Text = "������"
                End If
            End If
            
        Case "����������ʾ����"
            If T_Patient.lng�༭ = 0 Then Exit Function
            If mstr��ʼʱ�� <> "" Then
                '����ѡ�����
                intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                intCOl = intCOl - 5
                If intCOl < mintColMin Then intCOl = mintColMin
                
                '����õ��з��ص�ʱ�䷶Χ
                If Trim(strParam) <> "" Then '�����±༭���������ʾ�Ǵ���ʱ��(��Ϊ�����������µ�ˢ�º�,�ᶨλ����һ��)
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
                Else
                    strTime = Split(GetCurveDate(intCOl, mstr��ʼʱ��, gintHourBegin), ";")(0)
                End If
                
                If Format(strTime, "YYYY-MM-DD HH:mm:ss") < Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss") Then
                    strTime = Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
                End If
                
                strInput = T_Patient.lng����ID & ";" & T_Patient.lng��ҳID & ";" & T_Patient.lng�ļ�ID & ";" & T_Patient.lngӤ�� & ";" & T_Patient.lng����ID & ";" & T_Patient.lng����ȼ�
                If frmCaseTendBodySetShowData.ShowEdit(UserControl.Extender.ParentForm, strInput, CDate(strTime), CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")), mint����Ӧ��, mblnMoved, FontSize) = True Then
                    '����ɹ���ˢ�����µ���ʾ
                    strParam = mstrParam2
                    picMain.Tag = ""
                    Call zlMenuClick("װ������", strParam)
                End If
            End If
            
        Case "�������ݱ༭"
            If T_Patient.lng�༭ = 0 Then Exit Function
            Dim strCurDate As String, strDay As String
            If mstr��ʼʱ�� <> "" Then
                If picMain.Tag = "" Then picMain.Tag = mstr��ʼʱ�� & ";" & mstr����ʱ��
                
               strCurDate = zldatabase.Currentdate
               
    
                
'               intCOl = (lblCur.Left - mshUpTab.ColWidth(0) - mshUpTab.Left - ((mshUpTab.ColWidth(1) - lblCur.Width) / 2)) / mshUpTab.ColWidth(1) + 1
                'intCOl = mshUpTab.Col
                '����õ��з��ص�ʱ�䷶Χ
                If Trim(strParam) <> "" Then
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss") & ";" & Format(varParam(1), "YYYY-MM-DD HH:mm:ss")
                Else
                    '����ѡ�����
                    intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                    intCOl = intCOl - 5
                    If intCOl < mintColMin Then intCOl = mintColMin
                    strTime = GetCurveDate(intCOl, mstr��ʼʱ��, gintHourBegin)
                End If
                
                If Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss") < Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss") Then
                    strTime = Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss") & ";" & Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss")
                ElseIf Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss") > Format(mstr����ʱ��, "YYYY-MM-DD HH:mm:ss") Then
                    strTime = Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss") & ";" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm:ss")
                End If
                
                
                
                strInput = T_Patient.lng����ID & ";" & T_Patient.lng��ҳID & ";" & T_Patient.lng�ļ�ID & ";" & T_Patient.lngӤ�� & ";" & T_Patient.lng����ID & ";" & T_Patient.lng����ȼ�
                If frmCaseTendBodySetData.ShowEditor(UserControl.Extender.ParentForm, strInput, strTime, mstr��ʼʱ��, mint����Ӧ��, mblnMoved, FontSize) = True Then
                    '����ɹ���ˢ�����µ���ʾ
                    mstrParam1 = mstrParam1 & String(9 - UBound(Split(mstrParam1, ";")), ";")
                    varParam = Split(mstrParam1, ";")
                    varParam(3) = T_Patient.lng�ļ�ID
                    varParam(6) = T_Patient.lngӤ��
                    varParam(9) = mintPage + 1
                    mstrParam1 = Join(varParam, ";")
                    strParam = mstrParam1
                    
                    Call zlMenuClick("��ʼ��", strParam)
                End If
            End If
    End Select
    
    zlMenuClick = True
    Exit Function

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng��Ժ As Long, ByVal lng�༭ As Long, ByVal intӤ�� As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    '��ȡ��������
    T_Patient.lng����ID = lng����ID
    T_Patient.lng��ҳID = lng��ҳID
    T_Patient.lng��Ժ = lng��Ժ
    T_Patient.lng�༭ = lng�༭

    '���س�ʼ������,��������ʱ���
    Call InitPara

    '���б�Ҫ�ļ��
    '��ȡ���˵�ǰ����ȼ�
    T_Patient.lng����ȼ� = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵�ǰ����ȼ�", T_Patient.lng����ID, T_Patient.lng��ҳID)
    If rsTemp.BOF = False Then T_Patient.lng����ȼ� = zlCommFun.Nvl(rsTemp("����ȼ�"), 3)

    '����Ƿ�������������Ŀ
    gstrSQL = " Select 1 From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
              " Where C.��Ŀ���=A.��Ŀ��� " & _
                        "AND C.��ĿID=B.ID(+) " & _
                        "AND C.����ȼ�>=[1] " & _
                        "And A.��¼��=1 And RowNum<2 And C.��Ŀ���<>" & gint����
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ����������Ŀ", T_Patient.lng����ȼ�)
    If rsTemp.EOF Then
        MsgBox "����Ҫ��һ��������Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�жϸò����Ƿ��Ѿ�ת��
    If T_Patient.lng����ID > 0 And T_Patient.lng��Ժ = 1 Then
        gstrSQL = "select nvl(����ת��,0) ת�� from ������ҳ where ����ID=[1] and ��ҳID=[2]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��鲡���Ƿ�ת��", T_Patient.lng����ID, T_Patient.lng��ҳID)
        mblnMoved = (Val(rsTemp("ת��")) <> 0)
    End If
    
    vsf.Body.Appearance = flexFlat
    vsf.Body.RowHidden(0) = True
    vsf.Body.ColHidden(0) = True
    vsf.Body.ScrollBars = flexScrollBarNone
    vsf.Body.BorderStyle = flexBorderNone
    vsf.Body.OwnerDraw = flexODOver
    vsf.FixedCols = 1
    vsf.FixedRows = 1
    vsf.Rows = 2
    vsf.Body.RowHeight(vsf.FixedRows) = 400
    vsf.Height = vsf.Body.RowHeight(vsf.FixedRows)
     
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReadBodyInfo()
'����:��ȡ������Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, strTime As String, strStart As String, strTo As String
    Dim intCOl As Integer
    Dim bln�����ʾ��Ժ As Boolean
    On Error GoTo hErr
    
    strStart = mstr��ʼʱ��
    strTo = mstr����ʱ��
    
    If zldatabase.GetPara("���µ���ʾ���", glngSys, 1255, 1) = 0 Then
        lblCard(7).Visible = False
        txtCard(7).Visible = False
    Else
        lblCard(7).Visible = True
        txtCard(7).Visible = True
    End If
    
    If CStr(mstrEndDate) < Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln��Ժ = False Then
        mstrEndDate = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If mintAllPage = mintPage + 1 Then
        If CStr(mstr����ʱ��) < Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln��Ժ = False Then
            mstr����ʱ�� = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    txtCard(3).Text = ""
    
    '������������������¼���ʱ�䣬��Ӥ�����µ��Ŀ�ʼʱ��
    If T_Patient.lngӤ�� > 0 Then
        mstrSQL = " Select  b.����ʱ�� From ������������¼ B Where ����id=[1] And ��ҳid=[2] And ���=[3] "
        Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "��ȡ��������Ϣ", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID), T_Patient.lngӤ��)
        If rsTmp.BOF = False Then
            mstrEnterDate = Format(zlCommFun.Nvl(rsTmp("����ʱ��").Value), "yyyy-MM-dd HH:mm:ss")
            txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("����ʱ��").Value), "yyyy-MM-dd")
            strStart = mstrEnterDate
        End If
    End If
    
    '�˴�����ʱ��ת��
    intCOl = GetCurveColumn(CDate(strStart), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strStart = Split(GetCurveDate(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(0)
    
    If CDate(strStart) < CDate(mstr��ʼʱ��) Then
        strStart = Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    intCOl = GetCurveColumn(CDate(strTo), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strTo = Split(GetCurveDate(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(1)
    If CDate(Format(strTo, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")) Then
        strTo = Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")
    End If
    
    mstr��ʼʱ�� = Format(strStart, "yyyy-MM-dd HH:mm:ss")
    mstr����ʱ�� = Format(strTo, "yyyy-MM-dd HH:mm:ss")
    
    picMain.Tag = mstr��ʼʱ�� & ";" & mstr����ʱ��
    
    bln�����ʾ��Ժ = False
    If CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) Then
        bln�����ʾ��Ժ = True
    ElseIf CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) = CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) And T_BodyFlag.��Ժ = 0 Then
        bln�����ʾ��Ժ = True
    End If
    
    '��Ժʱ��(�����ʱ��Ϊ׼)
    mstrSQL = "select ��ʼʱ�� from ���˱䶯��¼ where ����id=[1] And ��ҳid=[2] and ��ʼԭ��=2 order by ��ʼʱ��"
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "�䶯��¼", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID))
    If rsTmp.BOF = False Then
        If txtCard(3).Text = "" And bln�����ʾ��Ժ = True Then txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("��ʼʱ��").Value), "yyyy-MM-dd")
    End If
    
    '��ȡ���˻�����Ϣ
    mstrSQL = " Select  b.����,A.סԺ��,A.��Ժ���� ��Ժʱ��,b.�Ա�,A.���� From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "��ȡ������Ϣ", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID))
    If rsTmp.BOF = False Then
        txtCard(0).Text = zlCommFun.Nvl(rsTmp("����").Value)
        txtCard(0).Tag = zlCommFun.Nvl(rsTmp("����").Value)
        txtCard(1).Text = zlCommFun.Nvl(rsTmp("סԺ��").Value)
        txtCard(5).Text = zlCommFun.Nvl(rsTmp("�Ա�").Value)
        txtCard(6).Text = zlCommFun.Nvl(rsTmp("����").Value)
        If txtCard(3).Text = "" Then txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("��Ժʱ��").Value), "yyyy-MM-dd")
    End If
    
    
    '��ȡ���˿��ҡ����ŵ���Ϣ
    
    txtCard(2).Text = ""
    txtCard(4).Text = ""
    
    mstrSQL = " Select  c.���� As ����,b.���� As ����,a.����,a.��ʼԭ�� " & _
                "From ���˱䶯��¼ a,���ű� b,���ű� c " & _
                "Where a.����id=[1] And a.��ҳid=[2] And a.����id Is Not Null And a.����id=b.id and a.����id=c.id And a.��ʼʱ��-4/24<=[3] And Nvl(a.��ֹʱ��,Sysdate)>=[4] Order By a.��ʼʱ��"
    
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "��ȡ���˿��ҡ����ŵ���Ϣ", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID), CDate(mstr����ʱ��), CDate(mstr��ʼʱ��))
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            
            If zlCommFun.Nvl(rsTmp("����").Value) <> strTmp And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
            
                strTmp = zlCommFun.Nvl(rsTmp("����").Value)
                
                If txtCard(2).Text = "" Then
                    txtCard(2).Text = strTmp
                Else
                    txtCard(2).Text = txtCard(2).Text & "->" & strTmp
                End If
                
            End If

            If zlCommFun.Nvl(rsTmp("����").Value) <> strTime And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
                strTime = zlCommFun.Nvl(rsTmp("����").Value)
                
                If txtCard(4).Text = "" Then
                    txtCard(4).Text = strTime
                Else
                    txtCard(4).Text = txtCard(4).Text & "->" & strTime
                End If
                
            End If
                        
            rsTmp.MoveNext
        Loop
        
        If Left(txtCard(2).Text, 2) = "->" Then txtCard(2).Text = Mid(txtCard(2).Text, 3)
        If Left(txtCard(4).Text, 2) = "->" Then txtCard(4).Text = Mid(txtCard(4).Text, 3)
    End If
    
    '��ȡ���������Ϣ
    mstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As ������ From Dual"
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "������", "������", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID), CDate(strStart))
    If rsTmp.BOF = False Then
        If T_Patient.lngӤ�� = 0 Then
            txtCard(7).Text = zlCommFun.Nvl(rsTmp("������").Value)
        Else
            txtCard(7).Text = ""
        End If
    Else
        txtCard(7).Text = ""
    End If
    txtCard(7).Tag = txtCard(7).Text
    
    Call zlMenuClick("��ʾ������Ϣ")

    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FaceInitTable(Optional ByVal blnInitUpdate As Boolean = True)
'---------------------------------------------------------------------
'���ܣ�������ʾ���ϱ�������
'---------------------------------------------------------------------
    '�����������±��
    Dim rsTemp As New ADODB.Recordset
    Dim intCOl As Integer, intRow As Integer
    Dim lngCount As Long
    Dim lngWith As Long
    Dim strPace As String
    
    On Error GoTo Errhand
    
    If T_DrawClient.�е�λ = 0 Then T_DrawClient.�е�λ = glngColStep
    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    '�õ���������
    lngCount = CurveCount
    
    '�������µ��̶���������ұ߾�
    If lngCount <= 3 Then
        T_DrawClient.�̶�����.Right = T_DrawClient.�̶�����.Left + glngLableWith
    Else
        T_DrawClient.�̶�����.Right = T_DrawClient.�̶�����.Left + lngCount * glngLableStep
    End If
    
    lngWith = T_DrawClient.�е�λ * Screen.TwipsPerPixelX
    
    With mshUpTab
        .Cols = 43
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        .Cell(flexcpText, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
    
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(2) = True
        .ColWidth(0) = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) * Screen.TwipsPerPixelX
        .TextMatrix(0, 0) = "��       ��"
        .TextMatrix(1, 0) = IIf(T_Patient.lngӤ�� = 0, "ס Ժ �� ��", "�� �� �� ��")
        .TextMatrix(2, 0) = "����������"
        .TextMatrix(3, 0) = "ʱ       ��"
        
        '.Cell(flexcpWidth, 0, 1, .Rows - 1, .Cols - 1) = lngWith
        For intCOl = 1 To .Cols - 1
            .ColWidth(intCOl) = lngWith
        Next
        .ColWidthMin = lngWith
        .Redraw = flexRDBuffered
    End With
    
    '�ϲ���Ԫ�����
    For intRow = 0 To 2
        Call UniteCellCol(mshUpTab, 6, intRow, mshUpTab.FixedCols)
    Next intRow
    
    If blnInitUpdate = True Then Call ShowUptab
    
    With vsf
        .Cols = 0
        .NewColumn "", 0, 1
        .NewColumn "��Ŀ", mshUpTab.ColWidth(0) + 10, 1
    
        For intCOl = 1 To 42
            .NewColumn intCOl, lngWith + 7, 1, , 1
        Next
        
        .Left = T_DrawClient.ƫ����X * Screen.TwipsPerPixelX
        .FixedCols = 2
        .Rows = 2
        .Body.Appearance = flexFlat
        .Body.RowHidden(0) = True
        .Body.ColHidden(0) = True
        .Body.ScrollBars = flexScrollBarNone
        .Body.BorderStyle = flexBorderNone
        .Body.OwnerDraw = flexODOver
        .Cell(flexcpAlignment, 1, 1) = flexAlignCenterCenter
        .Cell(flexcpFontName, 1, 2, 1, .Cols - 1) = "Times New Roman"
        .Cell(flexcpFontSize, 1, 2, 1, .Cols - 1) = 7.5
        .Cell(flexcpForeColor, 1, 2, 1, .Cols - 1) = RGB_RED
        .Body.Select 1, 1
        .Body.CellBorder 0, 1, 0, 0, 0, 0, 0
        .Body.Select 1, vsf.Cols - 1
        .Body.CellBorder 0, 0, 0, 1, 0, 0, 0
        .Body.BackColorFixed = .Body.BackColor
        .Visible = False
        For intCOl = 3 To .Cols - 1 Step 2
            .Cell(flexcpBackColor, 1, intCOl, 1, intCOl) = &HF7ECE6
        Next
        For intCOl = 1 To .Cols - 1
            .EditMode(intCOl) = 0
        Next
        .Height = .Body.RowHeight(.FixedRows)
    End With
    
    '�����±��(������Ŀ)
    With mshDownTab
        .Cols = 46
        .Rows = 1
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(0) = True
        .Tag = 0
        
        For intCOl = .FixedCols To .Cols - 1
            .ColWidth(intCOl) = mshUpTab.ColWidth(1)
            If (intCOl - .FixedCols + 1) Mod 2 = 0 Then
                .Cell(flexcpBackColor, 0, intCOl, .Rows - 1, intCOl) = &H80000013
            End If
        Next intCOl

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
    mItemNO.���� = 0
    mintRepairRows = zldatabase.GetPara("���±������", glngSys, 1255, 8)
    mbln��ʾƤ�� = (Val(zldatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0")) = 1)
    
    '�������Ƿ��Ǳ����Ŀ
    gstrSQL = "select ��¼�� From ���¼�¼��Ŀ where ��Ŀ���=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "���¼�¼��Ŀ", gint����)
    If rsTemp.RecordCount > 0 Then
         mintRepairRows = mintRepairRows - IIf(Val(Nvl(rsTemp!��¼��)) = 2, 1, 0)
    End If
    If mintRepairRows < 0 Then mintRepairRows = 0

    '�������б����Ŀ�������̶���Ŀ�������ݵĻ��Ŀ
    Set rsTemp = GetAppendGridItem(T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng����ȼ�, T_Patient.lngӤ��, Int(CDate(mstr��ʼʱ��)), CDate(mstr����ʱ��), IIf(T_Patient.lngӤ�� = 0, 1, 2), T_Patient.lng����ID, mblnMoved)
    With rsTemp
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            mshDownTab.Rows = 0
            Call AppenGridItem(rsTemp)
        Else
            mshDownTab.Rows = 0
        End If
    End With
    
    mshDownTab.Rows = mintRepairRows
    
    '������ʣ�µĿ���
    If mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
        For intRow = Val(mshDownTab.Tag) To mshDownTab.Rows - 1

            mshDownTab.MergeRow(intRow) = True
            For intCOl = 0 To mshDownTab.FixedCols
                strPace = " " & String(intCOl, " ") & String(intRow, " ")
                mshDownTab.TextMatrix(intRow, intCOl) = strPace & "" & strPace
            Next intCOl
            
            Call UniteCellCol(mshDownTab, 6, intRow, mshDownTab.FixedCols)
        Next intRow
    End If
    
    If mbln��ʾƤ�� And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
        intRow = Val(mshDownTab.Tag)
        strPace = " " & String(1, " ") & String(intRow, " ")
        mshDownTab.TextMatrix(intRow, 0) = strPace & "Ƥ�Խ��" & strPace
    End If
    
    '����������λ��
    If mItemNO.���� <> 0 Then
        mbln�������� = False
    End If
    
    '���ñ����ɫ
    For intCOl = mshDownTab.FixedCols To mshDownTab.Cols - 1
        If (intCOl - mshDownTab.FixedCols + 1) Mod 2 = 0 Then
            mshDownTab.Cell(flexcpBackColor, 0, intCOl, mshDownTab.Rows - 1, intCOl) = &HF7ECE6
        End If
    Next intCOl
    mshDownTab.Cell(flexcpAlignment, 0, 0, mshDownTab.Rows - 1, mshDownTab.Cols - 1) = 4
    
    Call picBack_Resize
    
    Call Paint_Canvas(mblnAutoAdjust) '��ʼ����������
    
    Call picBack_Resize
    
    Call SetVisible
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function SetVisible() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------------------
    If T_Patient.lng�༭ = 0 Then
        mshUpTab.Enabled = False
        mshDownTab.Enabled = False
    Else
        mshUpTab.Enabled = True
        mshDownTab.Enabled = True
    End If
End Function


Private Function ShowUptab() As Boolean
'----------------------------------------------------------------
'����:�������������Ϣ ������Ժ���ڣ�סԺ������������ע
'----------------------------------------------------------------
    Dim lngValue  As Long, intCOl As Long
    Dim lngDays   As Long
    Dim i As Long, j As Long
    Dim lngColor  As Long
    Dim intMinCol As Long, intMaxCol As Long
    Dim strTmp As String
    Dim arrOperDay, strTmp1 As String
    Dim rsTmp  As New ADODB.Recordset
    Dim strʱ�� As String
    Dim intDays As Integer
    Dim lng���� As Long
    Dim lngWith As Long

    On Error GoTo Errhand

    With mshUpTab
        
        lngValue = 0
        gstrSQL = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As ��ʼ���� From Dual"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡסԺ����", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, Int(CDate(mstr��ʼʱ��)))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("��ʼ����").Value
        End If
        
        '�ϱ��ʽ�е�Ԫ��ϲ��ģ��˴���Ҫ���д���
        For intCOl = 1 To 7

            .ColData(intCOl) = 0
            .Row = 0
            .Col = intCOl
            .ColAlignment(intCOl) = 4

            strTmp = Format(CDate(mstr��ʼʱ��) + intCOl - 1, "yyyy-MM-dd")

            lngDays = lngValue + (intCOl - 1)
            
            For i = 1 To 6
                .Row = 0
                .Col = (intCOl - 1) * 6 + i
                
                If Right(strTmp, 5) = "01-01" Then
                    'һ��ĵ�һ��
                    .Text = strTmp
                ElseIf strTmp = Format(mstrEnterDate, "yyyy-MM-dd") Then
                    '��Ժ��һ�죬д�����
                    .Text = strTmp
                ElseIf intCOl = 1 Then
                    .Text = Right(strTmp, 5)
                ElseIf Right(strTmp, 2) = "01" Then
                    .Text = Right(strTmp, 5)
                Else
                    .Text = Right(strTmp, 2)
                End If

                .Row = 1
                .Text = lngDays
            Next i
        Next
        
        '����ϱ�ʱ�����Ϣ
        If picMain.Tag <> "" Then
            Call CalcMinMaxCol(picMain.Tag, intMinCol, intMaxCol)
            mintColMin = intMinCol
            mintColMax = intMaxCol
            
            With picDisplay
                .Left = ((((intMaxCol - 1) \ 6) + 1) * 6 - 1) * mshUpTab.ColWidth(intMinCol) + mshUpTab.ColWidth(0)
                mshUpTab.Row = mshUpTab.FixedRows
                .Top = (mshUpTab.RowHeight(mshUpTab.FixedRows) - .Height) / 2
                .Enabled = IIf(T_Patient.lng�༭ = 1, True, False)
            End With
            
            lblCur.Left = (intMinCol - 1) * .ColWidth(intMinCol) + .ColWidth(0)
            '������ʾ
            lblCur.Left = lblCur.Left + (.ColWidth(intMinCol) - lblCur.Width) / 2
            lblCur.Top = .Height - lblCur.Height
            lblCur.Enabled = IIf(T_Patient.lng�༭ = 1, True, False)
        End If
        '��ΪDrawCell��� �����п�̫Сʱ��������ʾ������
'        For i = 1 To 7
'            '�����������ʱ��
'            .Row = 3
'            For j = 1 To 6
'
'                Select Case j
'
'                    Case 1
'                        strTmp = gintHourBegin + 4 * 0
'                        lngColor = &H8080FF
'
'                    Case 2
'                        strTmp = gintHourBegin + 4 * 1
'                        lngColor = &H8080FF
'
'                    Case 3
'                        strTmp = gintHourBegin + 4 * 2
'                        lngColor = &H80000012
'
'                    Case 4
'                        lngColor = &H80000012
'                        strTmp = gintHourBegin + 4 * 3
'
'                    Case 5
'                        lngColor = &H80000012
'                        strTmp = gintHourBegin + 4 * 4
'
'                    Case 6
'                        lngColor = &H8080FF
'                        strTmp = gintHourBegin + 4 * 5
'                End Select
'
'                .Col = j + (i - 1) * 6
'                .ColAlignment(.Col) = 4
'
'                If .Col >= intMinCol And .Col <= intMaxCol Then
'                    lngColor = lngColor
'                Else
'                    lngColor = RGB_FleetGRAY
'                End If
'
'                .CellForeColor = lngColor
'
'                If picMain.Tag <> "" Then
'                    .Text = strTmp
'                End If
'
'            Next j
'        Next i
        
        For i = 1 To 7
            mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
            mstrOpdays(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
        Next i
        
        '��ȡ�����־������ֹͣ������־
        mintOpDays = Val(zldatabase.GetPara("�������ע����", glngSys, 1255, "10"))
        mblnStopFlag = (Val(zldatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0")) = 1)
        '51338,������,2012-07-06
        strTmp = zldatabase.GetPara("��������ȱʡ��ʽ", glngSys, 1255, "2")
        If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
            mintOpFormat = Val(strTmp)
        Else
            mintOpFormat = 0
        End If
        
        strTmp = ""
        '��ʾ��ǰ�ε��������
        gstrSQL = "select B.����ʱ�� ʱ��" & _
            "   From ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & _
            "   where A.ID=B.�ļ�ID And  B.ID=C.��¼ID And A.ID=[1] And nvl(A.Ӥ��,0)=[4]" & _
            "   and A.����ID=[2] and A.��ҳID=[3] and C.��¼����=4 and C.��ֹ�汾 is null" & _
            "   and B.����ʱ�� between [5] and [6] order by B.����ʱ��"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", Val(T_Patient.lng�ļ�ID), T_Patient.lng����ID, T_Patient.lng��ҳID, Val(T_Patient.lngӤ��), Int(CDate(mstr��ʼʱ��) - 14), CDate(mstr����ʱ��))
        
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
            gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        Do While Not rsTmp.EOF
            strʱ�� = Format(rsTmp("ʱ��"), "YYYY-MM-DD")
            For i = 1 To 7
                If DateDiff("d", mstr��ʼʱ��, mstr����ʱ��) + 1 >= i Then
                    intDays = DateDiff("d", strʱ��, mstr��ʼʱ��) + (i - 1)

                    Select Case intDays

                        Case 0 '��ǰ�����ڵ�������ʼʱ��
                            'Modify 2012-03-05 �޸�һ������ж������
                            If Trim(mstrOpdays(i)) <> "" Then
                                mstrOpdays(i) = strʱ�� & "/" & mstrOpdays(i)
                            Else
                                mstrOpdays(i) = strʱ��
                            End If
                            
                        Case 1 To mintOpDays '������ʼ����

                            If mblnStopFlag Then '������ע�������ڴ�����ʱֹͣǰһ�α�ע
                                mstrOpValue(i) = intDays
                            Else
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = intDays & "/" & mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = intDays
                                End If
                            End If
                    End Select
                End If
            Next i
            rsTmp.MoveNext
        Loop
        
        
        '��ȡ��ǰ��ʼ����-14��ǰ��������¼��Ϣ
        gstrSQL = "Select Nvl(Count(B.����ʱ��),0) ����" & _
            "   From ���˻����ļ� A, ���˻������� B,���˻�����ϸ C" & _
            "   Where A.ID=B.�ļ�ID And B.ID=C.��¼ID and A.ID=[1] and nvl(A.Ӥ��,0)=[4]" & _
            "   And A.����ID=[2] And A.��ҳID=[3] And C.��¼����=4 and C.��ֹ�汾 is null" & _
            "   And B.����ʱ�� <[5] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", Val(T_Patient.lng�ļ�ID), T_Patient.lng����ID, T_Patient.lng��ҳID, Val(T_Patient.lngӤ��), Int(CDate(mstr��ʼʱ��)))
        
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
            gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        lng���� = 0
        If rsTmp.BOF = False Then lng���� = Val(rsTmp("����"))
        For i = 1 To 7
            If DateDiff("d", mstr��ʼʱ��, mstr����ʱ��) + 1 >= i Then
                '�޸�һ����ܴ��ڶ������
                If Trim(mstrOpdays(i)) <> "" Then
                    arrOperDay = Split(mstrOpdays(i), "/")
                Else
                    arrOperDay = Split("1", "/")
                End If
                lngValue = lng����
                If Trim(mstrOpdays(i)) <> "" And lngValue + UBound(arrOperDay) < 12 Then
                    strTmp = "": strTmp1 = ""
                    For j = UBound(arrOperDay) + 1 To 1 Step -1
                        lng���� = lngValue + j
                        strTmp1 = Switch(lng���� = 1, "��", lng���� = 2, "��", lng���� = 3, "��", lng���� = 4, "��", lng���� = 5, "��", lng���� = 6, "��", lng���� = 7, "��", lng���� = 8, "��", lng���� = 9, "��", lng���� = 10, "��", lng���� = 11, "��", lng���� = 12, "��")
                        If strTmp = "" Then
                            strTmp = strTmp1
                        Else
                            strTmp = strTmp & "/" & strTmp1
                        End If
                        If mblnStopFlag Then Exit For
                    Next j
                    lng���� = lngValue + UBound(arrOperDay) + 1
                    If mblnStopFlag Then '������ע�������ڴ�����ʱֹͣǰһ�α�ע
                        Select Case mintOpFormat
                            Case 1 '--��ʾ0
                                mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1)) & "0" & .TextMatrix(2, ((i - 1) * 6 + 1))
                            Case 2 '--��ʾ����
                                If strTmp = "��" Then
                                    mstrOpValue(i) = 0
                                Else
                                    mstrOpValue(i) = strTmp & "-0"
                                End If
                            Case Else '--����ʾ
                                 mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
                        End Select
                    Else
                        Select Case mintOpFormat
                            Case 1 '--��ʾ0
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = 0 & "/" & mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = 0
                                End If
                            Case 2 '--��ʾ����
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = strTmp & "/" & mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = strTmp
                                End If
                            Case Else  '--����ʾ
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
                                End If
                        End Select
                    End If
                    .Row = 2
                    For j = 1 To 6
                        .Col = j + (i - 1) * 6
                        .Text = mstrOpValue(i)
                    Next j
                Else
                    .Row = 2
                    For j = 1 To 6
                        .Col = j + (i - 1) * 6
                        .Text = mstrOpValue(i)
                    Next j
                End If
            End If
        Next i
        '�趨���ڣ�סԺ�����ı���ɫ
        mshUpTab.Cell(flexcpForeColor, 0, mshUpTab.FixedCols, 1, mshUpTab.Cols - 1) = 16711680
        '�趨���� �����ı���ɫ
        '51283,������,2012-07-11
        lngColor = Val(zldatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "255"))
        mshUpTab.Cell(flexcpForeColor, 2, mshUpTab.FixedCols, 2, mshUpTab.Cols - 1) = lngColor

        lngWith = T_DrawClient.�е�λ * Screen.TwipsPerPixelX
        'mshUpTab.Cell(flexcpWidth, 0, 1, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = lngWith
        For intCOl = 1 To mshUpTab.Cols - 1
            mshUpTab.ColWidth(intCOl) = lngWith
        Next intCOl
        mshUpTab.ColWidthMin = lngWith
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With

    ShowUptab = True
    Exit Function
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowDowntab() As Boolean
    '����������ݣ�������Ŀ��Ϣ��
    Dim rsTemp   As New ADODB.Recordset
    Dim rsDownTab As New ADODB.Recordset
    Dim intRow As Integer, intRow1 As Integer
    Dim intCOl As Integer, intCol1 As Integer
    Dim intColCount As Integer, intRowCount As Integer
    Dim intDay As Integer
    Dim strItems As String, strItemName As String, strSql As String
    Dim lngItemCode As Long
    Dim strPace As String
    Dim str��Ŀ���� As String, str��Ŀ����1 As String
    Dim int��¼Ƶ�� As Integer, int��Ŀ���� As Integer, int��Ŀ���� As Integer, int��Ŀ��ʾ As Integer, int��Ժ�ײ� As Integer
    Dim strBegin As String, str��� As String, strPart As String
    Dim int����ѹ As Integer, int����ѹ As Integer, Int�к� As Integer
    Dim blnColor As Boolean
    Dim lngColor As Long
    Dim arrTmpString0(1 To 42) As String, arrTmpString1(1 To 42) As String, arrTmpString2(1 To 42) As String
    Dim blnAdd As Boolean, blnValue As Boolean
    Dim SinX As Single
    Dim i As Integer
    Dim int����λ�� As Integer, intValue As Integer, int������������ʽ As Integer
    Dim bln���ܵ��� As Boolean, bln¼��Сʱ As Boolean
    Dim arrTmp() As String
    Dim dtBegin As Date, dtEnd As Date
    
    On Error GoTo Errhand
    
    Call InitPublicData '��ȡ��������
    
    ReDim mstrNewString(mintRepairRows, 6)
    'mstrNewString = Split(String(6, ";"), ";")
    int������������ʽ = zldatabase.GetPara("����������", glngSys, 1255, 0)
    bln���ܵ��� = (Val(zldatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    mbln�೦�����ӷ�ĸ��ʾ = (Val(zldatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
    bln¼��Сʱ = (Val(zldatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, 0)) = 1)
    
    gbln��Ժ = mbln��Ժ
    dtBegin = Int(CDate(mstr��ʼʱ��) - 1)
    dtEnd = CDate(CDate(mstr����ʱ��) + 1)
    
    If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) Then _
        dtBegin = CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss"))
    If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")) Then _
        dtEnd = CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss"))
    
    '��ȡ��Ŀ����(ƴ���ַ���)
    strItems = ""
    For intRow = mshDownTab.FixedRows To Val(mshDownTab.Tag) - 1
        If Val(mshDownTab.RowData(intRow)) <> mItemNO.Ѫѹ Then
            i = InStr(1, mshDownTab.TextMatrix(intRow, 0), "(")
            If i > 0 Then
                strItemName = Trim(Left(mshDownTab.TextMatrix(intRow, 0), i - 1))
            Else
                strItemName = Trim(mshDownTab.TextMatrix(intRow, 0))
            End If
            If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                strItems = strItems & ",'" & strItemName & "'"
            End If
        End If
    Next
    
    If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
    If Not mbln�������� Then strItems = strItems & ",'����'"
    strItems = strItems & ",'����ѹ','����ѹ'"
    If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
    'Debug.Print "��ȡ���ݿ�ʼ---" & Now
    '��ȡ�������±���¼
    gstrSQL = " SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & _
        "   DECODE(E.��Ŀ����,2,C.���²�λ || D.��¼�� ,D.��¼��) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,E.��Ŀ���� " & _
        "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
        "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID   AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
        "   AND B.����id=[2]  AND B.��ҳid=[3] AND INSTR([6],decode(E.��Ŀ����,2,C.���²�λ || D.��¼�� ,D.��¼��))>0 " & _
        "   AND D.��Ŀ���=C.��Ŀ���  AND MOD(c.��¼����,10)=1  AND E.��Ŀ���=D.��Ŀ��� " & _
        "   AND nvl(E.����ȼ�,0)>=[8]  AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null AND D.��¼��=2 "
    
    '��ȡ�����±��Ļ�����Ŀ
    strSql = "  SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & _
        "   D.��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,D.��Ŀ����" & _
        "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,(SELECT A.��Ŀ���,A.��Ŀ����, 1 ��Ŀ����,B.����� FROM �����¼��Ŀ A,���������Ŀ B" & vbNewLine & _
        "       WHERE A.��Ŀ���=B.��� AND NOT EXISTS (SELECT C.��Ŀ��� FROM ���¼�¼��Ŀ C,���������Ŀ E WHERE C.��Ŀ���=E.��� AND C.��Ŀ���=A.��Ŀ���)" & vbNewLine & _
        "       AND NVL(A.Ӧ�÷�ʽ,0)=1 AND NVL(A.����ȼ�,0)>=[8] AND NVL(A.���ò���,0) IN (0,[9])" & vbNewLine & _
        "       AND (A.���ÿ���=1 OR (A.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D WHERE D.��Ŀ���=A.��Ŀ��� AND D.����ID=[10])))) D" & _
        "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID   AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
        "   AND B.����id=[2]  AND B.��ҳid=[3]  AND D.��Ŀ���=C.��Ŀ���  AND C.��¼����=1" & _
        "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null"

    gstrSQL = "Select ID,ʱ��,��¼����,��ʾ,���,���²�λ,δ��˵��,������Դ,��Ŀ����,��Ŀ���,��ԴID,����,��Ŀ���� From (" & _
        "   " & gstrSQL & " UNION ALL " & strSql & ")" & _
        "   Order By  Decode(��Ŀ����,'����ѹ',0,1)," & strItems & ",ʱ��"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���±������", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, _
                        CDate(dtBegin), CDate(dtEnd), strItems, T_Patient.lngӤ��, T_Patient.lng����ȼ�, IIf(T_Patient.lngӤ�� = 0, 1, 2), T_Patient.lng����ID)
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
    
    'Debug.Print "��ȡ���ݽ���---" & Now
    '1---��������������
    vsf.Cell(flexcpText, 1, 2, 1, vsf.Cols - 1) = ""
    vsf.Cell(flexcpData, 1, 2, 1, vsf.Cols - 1) = ""
    vsf.Cell(flexcpForeColor, 1, 2, 1, vsf.Cols - 1) = 200
    
    rsTemp.Filter = "��Ŀ���=" & gint����
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    With rsTemp
        Do While Not .EOF
            If CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")) Then
                blnAdd = False
                intCOl = GetCurveColumn(rsTemp!ʱ��, mstr��ʼʱ��, gintHourBegin) + vsf.FixedCols - 1
                str��� = zlCommFun.Nvl(rsTemp!���) & ";" & Nvl(rsTemp!���²�λ)
                If intCOl < vsf.Cols Then
                    If arrTmpString1(intCOl - vsf.FixedCols + 1) <> "" Then
                        If (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) <> 1 And Val(zlCommFun.Nvl(!��ʾ, 0)) <> 1) Or _
                            (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 1 And Val(zlCommFun.Nvl(!��ʾ, 0)) = 1) Then
                            
                            '����Ǹ����ص�ʱ�����
                            SinX = GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
                            blnAdd = GetCanvasCenter(CDate(Format(arrTmpString1(intCOl - vsf.FixedCols + 1), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 1 Then
                            blnAdd = False
                        Else
                            blnAdd = True
                        End If
                        
                        If blnAdd = True Then
                            If Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 2 Then
                                arrTmpString0(intCOl - vsf.FixedCols + 1) = str���
                                arrTmpString1(intCOl - vsf.FixedCols + 1) = Format(rsTemp!ʱ��, "YYYY-MM-DD HH:mm:ss")
                                arrTmpString2(intCOl - vsf.FixedCols + 1) = 2
                                GoTo ErrNext
                            End If
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                arrTmpString2(intCOl - vsf.FixedCols + 1) = 2
                                GoTo ErrNext
                            End If
                        End If
                    Else
                        blnAdd = True
                    End If
                    
                    If blnAdd = True Then
                        arrTmpString0(intCOl - vsf.FixedCols + 1) = str���
                        arrTmpString1(intCOl - vsf.FixedCols + 1) = Format(rsTemp!ʱ��, "YYYY-MM-DD HH:mm:ss")
                        arrTmpString2(intCOl - vsf.FixedCols + 1) = Val(zlCommFun.Nvl(!��ʾ, 0))
                    End If
                End If
            End If
ErrNext:
        .MoveNext
        Loop
    End With
    
    'һ��ѭ����������,�����鵽��ʾ=2���������
    For i = 1 To 42
        If Val(arrTmpString2(i)) = 2 Then arrTmpString0(i) = ""
    Next i
    
    '2----��ʼ����������� ������Ϊͼ�����
    int����λ�� = 0
    blnValue = False
     'ѭ���������ֵ
    vsf.Cell(flexcpForeColor, 1, vsf.FixedCols, 1, vsf.Cols - 1) = Val(vsf.Tag)
    For i = 1 To 42
        intCOl = i + vsf.FixedCols - 1
        If InStr(1, arrTmpString0(i), ";") > 0 Then
            str��� = Split(arrTmpString0(i), ";")(0)
            strPart = Split(arrTmpString0(i), ";")(1)
        Else
            str��� = arrTmpString0(i)
            strPart = ""
        End If
        
        '��ӡ����ֵ���������ӡ�� ��һ��ʼ��������
        If IsNumeric(str���) Then
            vsf.TextMatrix(1, intCOl) = str���
            If blnValue = False Then
                intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                blnValue = True
                int����λ�� = 2
            End If
            
            If int������������ʽ = 0 Then '˳��������ʾ
                If intCOl Mod 2 = intValue Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterTop
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 1
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterBottom
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 2
                    End If
                End If
                
            Else        '������ʱ����֮��������ʾ
                If int����λ�� = 2 Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterTop
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 1
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterBottom
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 2
                    End If
                End If
                
                int����λ�� = int����λ�� + 1
                If int����λ�� > 2 Then int����λ�� = 1
            End If
        End If
    Next i
       
    'Debug.Print "���ݿ�ʼ---" & Now
    '��ȡ�����Ŀ������Ϣ
    With mshDownTab
        lngItemCode = 0
        str��Ŀ���� = ""
        For intRow = .FixedRows To .Tag - 1
            i = InStr(1, .TextMatrix(intRow, 0), "(")

            If i > 0 Then
                str��Ŀ����1 = Trim(Mid(.TextMatrix(intRow, 0), 1, i - 1))
            Else
                str��Ŀ����1 = Trim(.TextMatrix(intRow, 0))
            End If
            
            blnColor = False
            If str��Ŀ����1 & ";" & .RowData(intRow) <> str��Ŀ���� & ";" & lngItemCode Then
                
                lngItemCode = .RowData(intRow)
                str��Ŀ���� = str��Ŀ����1
                int��Ŀ���� = Val(Split(.TextMatrix(intRow, 1), ",")(0))
                int��¼Ƶ�� = Val(Split(.TextMatrix(intRow, 1), ",")(2))
                int��Ŀ��ʾ = Val(Split(.TextMatrix(intRow, 1), ",")(3))
                int��Ŀ���� = Val(Split(.TextMatrix(intRow, 1), ",")(4))
                int��Ժ�ײ� = Val(Split(.TextMatrix(intRow, 1), ",")(6))
                blnColor = (int��Ŀ���� = 2 And int��Ŀ���� = 1 And int��Ŀ��ʾ = 0)
                
                For intDay = 0 To 6
                    strBegin = DateAdd("D", intDay, CDate(mstr��ʼʱ��))
                    If CDate(strBegin) > CDate(mstr����ʱ��) Then strBegin = mstr����ʱ��
                    int����ѹ = 0
                    int����ѹ = 0
                    Int�к� = 0
                    'ѭ���õ�ĳ����Ŀĳ���������Ϣ
                    Set rsDownTab = ReturnItemRecord(rsTemp, Int(CDate(strBegin)), CDate(mstrEnterDate), lngItemCode & ";" & str��Ŀ���� & ";" & _
                                int��¼Ƶ�� & ";" & int��Ŀ��ʾ & ";" & int��Ŀ���� & ";" & int��Ժ�ײ�, bln���ܵ���, bln¼��Сʱ)
                    If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                    rsDownTab.Sort = "ʱ��,��Ŀ���,���"
                    Do While Not rsDownTab.EOF
                        str��� = zlCommFun.Nvl(rsDownTab!��¼����, "")
                        lngColor = 0
                        If blnColor Then lngColor = Val(zlCommFun.Nvl(rsDownTab!δ��˵��, 0))
                        intCOl = Val(rsDownTab!���)
                        intColCount = 0
                        intRow1 = 0
                        strPace = ""
                        
                        Select Case int��¼Ƶ��
                            Case 1
                                intRow1 = intRow
                                intCOl = intDay * 6 + .FixedCols
                                intColCount = 6
                                strPace = " "
                            Case 2
                                intRow1 = intRow
                                intCOl = (intCOl - 1) * 3 + intDay * 6 + .FixedCols
                                intColCount = 3
                                strPace = String(intCOl, " ")
                            Case 3
                                intRow1 = intRow + (intCOl - 1)
                                intCOl = intDay * 6 + .FixedCols
                                intColCount = 6
                                strPace = " "
                            Case 4
                                intRow1 = intRow + Fix((intCOl - 1) / 2)
                                Select Case intCOl
                                    Case 1, 3
                                        intCOl = 1
                                    Case 2, 4
                                        intCOl = 2
                                End Select
                                intCOl = (intCOl - 1) * 3 + intDay * 6 + .FixedCols
                                intColCount = 3
                                strPace = String(intCOl, " ")
                            Case 6
                                intRow1 = intRow
                                intCOl = (intCOl - 1) + intDay * 6 + .FixedCols
                                intColCount = 1
                                strPace = String(intCOl, " ")
                        End Select
                        
                        '��鱾����������Ƿ����������֮��
                        If mintRepairRows > 0 And mintRepairRows - 1 >= intRow1 Then
                            strPace = strPace & String(intDay + 1, " ") & String(intRow1, " ")
                            '������չʾ�ڱ����
                            Select Case rsDownTab!��Ŀ���
                                Case mItemNO.����ѹ
                                    If int����ѹ <> Val(rsDownTab!���) Then
                                        For i = 1 To intColCount
                                            intCol1 = intCOl + (i - 1)
                                            If intCol1 < mshDownTab.Cols Then
                                                If Trim(mshDownTab.TextMatrix(intRow1, intCol1)) <> "" Or str��� <> "" Then
                                                    If InStr(1, mshDownTab.TextMatrix(intRow1, intCol1), "/") > 0 Then
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & Trim(Split(mshDownTab.TextMatrix(intRow1, intCol1), "/")(0)) & "/" & str��� & strPace
                                                    Else
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & "/" & str��� & strPace
                                                    End If
                                                    '--����ţ�53505���޸��ˣ����Σ�Ѫѹ��ʾ���֡�
                                                    If str��� = "���" Or str��� = "�ܲ�" Or str��� = "���" Or str��� = "δ��" Then
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str��� & strPace
                                                    End If
                                                End If
                                            End If
                                        Next i
                                        int����ѹ = Val(rsDownTab!���)
                                    End If
                                Case mItemNO.Ѫѹ '����ѹ
                                    If int����ѹ <> Val(rsDownTab!���) Then
                                        For i = 1 To intColCount
                                            intCol1 = intCOl + (i - 1)
                                            If intCol1 < mshDownTab.Cols Then
                                                If Trim(mshDownTab.TextMatrix(intRow1, intCol1)) <> "" Or str��� <> "" Then
                                                    If InStr(1, mshDownTab.TextMatrix(intRow1, intCol1), "/") > 0 Then
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str��� & "/" & Trim(Split(mshDownTab.TextMatrix(intRow1, intCol1), "/")(1)) & strPace
                                                    Else
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str��� & "/" & strPace
                                                    End If
                                                End If
                                            End If
                                        Next i
                                        int����ѹ = Val(rsDownTab!���)
                                    End If
                                Case Else
                                    If Int�к� <> Val(rsDownTab!���) Then
                                        For i = 1 To intColCount
                                            intCol1 = intCOl + (i - 1)
                                            If intCol1 < mshDownTab.Cols Then
                                                mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str��� & strPace
                                                If int��Ŀ���� = 2 And int��Ŀ���� = 1 And int��Ŀ��ʾ = 0 Then
                                                    mshDownTab.Cell(flexcpForeColor, intRow1, intCol1, intRow1, intCol1) = lngColor
                                                End If
                                            End If
                                        Next i
                                        Int�к� = Val(rsDownTab!���)
                                    End If
                            End Select
                        End If
                    rsDownTab.MoveNext
                    Loop
                    If Format(strBegin, "YYYY-MM-DD") = Format(mstr����ʱ��, "YYYY-MM-DD") Then
                        Exit For
                    End If
                Next intDay
            End If
        Next intRow
        
        '��ʼ���Ƥ�Խ��
        If mbln��ʾƤ�� = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
            strSql = _
               "SELECT ʱ��,F_LIST2STR(CAST(COLLECT(ҩ����) AS T_STRLIST)) ҩ���� FROM (" & vbNewLine & _
                "   SELECT TO_CHAR(��ʼִ��ʱ��,'YYYY-MM-DD') ʱ��,DECODE(Ƥ�Խ��,'(+)',255,0) || '-#' || REPLACE(REPLACE(ҽ������,',',''),'-#','') || Ƥ�Խ��  ҩ����" & vbNewLine & _
                "   FROM ����ҽ����¼" & vbNewLine & _
                "   WHERE  ����ID=[1] AND ��ҳID=[2] AND Ӥ��=[3] AND Ƥ�Խ�� IS NOT NULL" & vbNewLine & _
                "   AND ��ʼִ��ʱ��  BETWEEN [4] AND [5]" & vbNewLine & _
                "   ORDER BY TO_DATE(TO_CHAR(��ʼִ��ʱ��,'YYYY-MM-DD'),'YYYY-MM-DD'),Ƥ�Խ��" & vbNewLine & _
                ") GROUP BY ʱ��"

            If mblnMoved Then
                strSql = Replace(strSql, "���˹�����¼", "H���˹�����¼")
            End If

            Set rsDownTab = zldatabase.OpenSQLRecord(strSql, "��ȡ���˹�����¼��Ϣ", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��))

            Do While Not rsDownTab.EOF
                intCOl = DateDiff("D", CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD")), CDate(Format(rsDownTab!ʱ��, "YYYY-MM-DD")))
                str��� = Nvl(rsDownTab!ҩ����)
                Call ShowTestis(str���, intCOl)
                rsDownTab.MoveNext
            Loop
        End If
    End With
    'Debug.Print "���ݽ���---" & Now
    
    
    ShowDowntab = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowTestis(ByVal strValue As String, ByVal intCOl As Integer)
'----------------------------------------------------------------------
'����:����Ƥ�Խ��Ҫ���������
'----------------------------------------------------------------------
    Dim intNum As Integer, i As Integer
    Dim lngColor As Long
    Dim strTmp As String, strPart As String, strPic As String
    Dim arrTmp() As String
    Dim LPoint As T_LPoint
    Dim lngDC As Long
    Dim objDraw As Object
    Dim lngH As Long, lngW As Long, lngX1 As Long, lngLen As Long
    Dim intRowCount As Integer
    Dim sngLen As Single
    Dim intRow As Integer
    
    Set objDraw = picBack
    intRowCount = Val(mshDownTab.Tag)
    intNum = 1
    strTmp = strValue
    If strTmp = "" Then Exit Sub
    LPoint.X = 0
    LPoint.W = mshDownTab.ColWidth(mshDownTab.FixedCols) / Screen.TwipsPerPixelX * 6
    lngW = LPoint.W
    lngX1 = 0
    
    '��ʼ�����Ƿ���Ҫ����
    strPart = ""
    arrTmp = Split(strTmp, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        lngColor = Val(Split(arrTmp(i), "-#")(0))
        strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") 'Ƥ�Խ��
        If Trim(strTmp) <> "" Then
            Do While True
                T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                strPic = strTmp
                If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                    sngLen = Round((LPoint.W - (LPoint.X - lngX1)) / T_Size.W, 2)
                    lngLen = Len(StrConv(strTmp, vbFromUnicode)) * sngLen
                    '�����תΪȫ��
                    strTmp = StrConv(strTmp, vbWide)
                    strPart = StrConv(Mid(StrConv(strTmp, vbFromUnicode), lngLen + 1), vbUnicode)
                    strTmp = StrConv(Mid(StrConv(strTmp, vbFromUnicode), 1, lngLen), vbUnicode)
                    '��ȡԭʼ�ַ���
                    strPart = Mid(strPic, Len(strTmp) + 1)
                    strTmp = Mid(strPic, 1, Len(strTmp))
                    
                    mstrNewString(intRow, intCOl) = mstrNewString(intRow, intCOl) & "," & lngColor & "-#" & strTmp
                    If Left(mstrNewString(intRow, intCOl), 1) = "," Then mstrNewString(intRow, intCOl) = Mid(mstrNewString(intRow, intCOl), 2)
                    
                    T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                    LPoint.X = LPoint.X + T_Size.W
                    strTmp = strPart
                    T_Size.W = objDraw.TextWidth("��") / T_TwipsPerPixel.X
                    If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                        LPoint.X = lngX1
                        intRow = intRow + 1
                        intNum = intNum + 1

                        If intRowCount + intNum > mintRepairRows Then Exit Sub
                    End If
                    If strTmp = "" Then Exit Do
                Else
                    mstrNewString(intRow, intCOl) = mstrNewString(intRow, intCOl) & "," & lngColor & "-#" & strTmp
                    If Left(mstrNewString(intRow, intCOl), 1) = "," Then mstrNewString(intRow, intCOl) = Mid(mstrNewString(intRow, intCOl), 2)
                    If T_Size.W + objDraw.TextWidth("��") / T_TwipsPerPixel.X - LPoint.W > 0 Then
                        LPoint.X = lngX1
                    Else
                        LPoint.X = LPoint.X + T_Size.W
                    End If

                    Exit Do
                End If
            Loop
        End If
    Next i
End Sub

Public Sub AppenGridItem(ByVal rsTemp As ADODB.Recordset)
    '��д������
    Dim intRow  As Integer, intRowStart As Integer
    Dim intƵ�� As Integer
    Dim intRowNum As Integer, intColNum As Integer
    Dim intRowCount As Integer, intNum As Integer
    Dim i As Integer, j As Integer
    Dim strText As String, strֵ�� As String

    On Error GoTo Errhand
    
    With rsTemp
        j = 0
        Do While Not .EOF
            intRowCount = mshDownTab.Rows
            Select Case !��¼��
                Case "����"
                    mItemNO.���� = !��Ŀ���
                    vsf.TextMatrix(1, 1) = Nvl(!��¼��, "����") & IIf(Not IsNull(!��λ), "(" & !��λ & ")", "")
                    vsf.Tag = Val(Nvl(!��¼ɫ, RGB_RED))
                Case "����ѹ"
                    mItemNO.����ѹ = !��Ŀ���
                Case Else
                    If mintRepairRows > 0 And mintRepairRows > intRowCount Then
                        j = j + 1
                        intƵ�� = zlCommFun.Nvl(!��¼Ƶ��, 2)
                        
                        '������Ŀ�򲨶���ĿƵ�����Ϊ2
                        If Val(zlCommFun.Nvl(!��Ŀ��ʾ)) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!��Ŀ���))) Then
                            If intƵ�� > 2 Then intƵ�� = 2
                        End If
                        
                        Select Case intƵ��
                            'intColNum Ҫ�ϲ�������
                            'intRowNum Ҫ�ϲ�����
                            Case 1
                                intRowNum = 1
                                intColNum = 6
                            Case 2
                                intRowNum = 1
                                intColNum = 3
                            Case 3
                                intRowNum = 3
                                intColNum = 6
                            Case 4
                                intRowNum = 2
                                intColNum = 3
                            Case 6
                                intRowNum = 1
                                intColNum = 1
                        End Select
                        
                        '����Ҫ��ӵ�����
                        If mshDownTab.Rows + intRowNum > mintRepairRows Then
                            intNum = mintRepairRows - mshDownTab.Rows
                        Else
                            intNum = intRowNum
                        End If
                        
                        intRowNum = intNum
                        mshDownTab.Rows = mshDownTab.Rows + intRowNum
                        mshDownTab.Tag = mshDownTab.Rows '��¼ʵ������ı������
                        intRowStart = mshDownTab.Rows - intRowNum
                        
                        '�ϲ��в���ֵ
                        For i = 1 To intRowNum
                            intRow = intRowStart + i - 1
                            
                            mshDownTab.MergeCol(0) = True
                            mshDownTab.MergeRow(intRow) = True
                            
                            If !��¼�� = "����ѹ" Then
                                mshDownTab.TextMatrix(intRow, 0) = String(j, "��") & "Ѫѹ" & IIf(Not IsNull(!��λ), "(" & !��λ & ")", "") & String(j, "��")
                                mItemNO.Ѫѹ = !��Ŀ���
                                mItemRow.Ѫѹ = intRowStart
                            Else
                                mshDownTab.TextMatrix(intRow, 0) = String(j, "��") & Replace(Nvl(!��¼��), ";", ":") & IIf(Not IsNull(!��λ), "(" & !��λ & ")", "") & String(j, "��")
                            End If
                            
                            strText = !��Ŀ���
                            mshDownTab.RowData(intRow) = strText
                            mshDownTab.RowHeight(intRow) = 255
                            
                            mshDownTab.TextMatrix(intRow, 1) = zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��ĿС��) & "," & _
                                intƵ�� & "," & zlCommFun.Nvl(!��Ŀ��ʾ) & "," & zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��Ժ�ײ�, 0)
                            mshDownTab.TextMatrix(intRow, 2) = zlCommFun.Nvl(!���ֵ, "")
                            mshDownTab.TextMatrix(intRow, 3) = zlCommFun.Nvl(!��Сֵ, "")
                            
                            Call UniteCellCol(mshDownTab, intColNum, intRow, mshDownTab.FixedCols)
                        Next i
                    End If
            End Select
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    picBack.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    picBuffer.Move lngLeft, lngTop
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboBaby.hWnd, KeyAscii)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_View_Jump '�˵�
            mcbrToolBarҳ��.Caption = Control.Caption
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("װ������", mstrParam)
            cbsMain.RecalcLayout
        Case conMenu_View_OneWeek To conMenu_View_FourWeek '4�����ڰ�ť
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("װ������", mstrParam)
            mcbrToolBarҳ��.Caption = mcbrItem.Controls.Item(mintPage + 1).Caption
        Case conMenu_View_Forward, conMenu_Manage_CallPrevious '��һҳ
            Call picDraw_KeyDown(vbKeyLeft, vbCtrlMask)
        Case conMenu_View_Backward, conMenu_Manage_CallNext '��һҳ
            Call picDraw_KeyDown(vbKeyRight, vbCtrlMask)
    End Select

End Sub


Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.Id

        Case conMenu_View_Jump '�˵�

            If Control.Parameter = "" Then
                Control.Checked = True
            Else
                Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = mintPage)
            End If

        Case conMenu_View_OneWeek To conMenu_View_FourWeek '4�����ڰ�ť

            If Control.Parameter <> "" Then
                Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = mintPage)
            End If

        Case conMenu_View_Forward, conMenu_View_Backward '����ҳ
            Control.Enabled = IIf(Val(Control.Parameter) > 0, True, False)
    End Select

End Sub

Private Sub cmdPrimitive_Click()
'�鿴���µ�ԭʼ����
    Dim strParams As String
    
    strParams = ""
    strParams = T_Patient.lng����ID & ";"
    strParams = strParams & T_Patient.lng��ҳID & ";"
    strParams = strParams & T_Patient.lng����ID & ";"
    strParams = strParams & T_Patient.lng�ļ�ID & ";"
    strParams = strParams & T_Patient.lng��Ժ & ";"
    strParams = strParams & T_Patient.lng�༭ & ";"
    strParams = strParams & T_Patient.lngӤ�� & ";1;" & mintPage + 1

    RaiseEvent CmdClick(strParams)
End Sub

Private Sub hsb_Change()
    picMain.Left = -1 * hsb.Value * msinHStep
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lbl�鿴_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl�鿴_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub mshDownTab_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim i As Integer
    Dim lngColor As Long
    Dim strTmp As String
    Dim arrTmp() As String, arrText() As String
    Dim LPoint As T_LPoint
    Dim T_ClientRect As RECT
    Dim lngBrush As Long, lngOldBrush As Long, lngBackColor As Long
    Dim lngDC As Long, lngFont As Long, lngOldFont As Long
    Dim objDraw As Object, stdset As Object
    Dim lngX1 As Long
    Dim intCOl As Integer, intRow As Integer
    
    On Error GoTo Errhand
    Err = 0
    intRow = UBound(mstrNewString)
Errhand:
    If Err <> 0 Then Exit Sub
    
    lngDC = hDC
    Set objDraw = picBack
    If mbln��ʾƤ�� = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 And Col >= mshDownTab.FixedCols And Row >= Val(mshDownTab.Tag) Then
        If (Col - mshDownTab.FixedCols) Mod 6 = 0 And UBound(mstrNewString) >= (Row - Val(mshDownTab.Tag)) Then
            intCOl = (Col - mshDownTab.FixedCols) / 6
            intRow = Row - Val(mshDownTab.Tag)
            strTmp = CStr(mstrNewString(intRow, intCOl))
            If strTmp = "" Then Exit Sub
            
            '�趨�ͻ������С
            With T_ClientRect
                .Left = Left + 1
                .Top = Top + 1
                .Right = Right - 1
                .Bottom = Bottom - 1
            End With

            LPoint.X = Left
            Call GetTextExtentPoint32(hDC, "��", Len("��"), T_Size)
            LPoint.Y = Top + (Bottom - Top) / 2 '+ T_Size.H / 2
            lngX1 = 0
            
            '1���������
            '�����뱳��ɫ��ͬ��ˢ��
            lngBackColor = GetRBGFromOLEColor(mshDownTab.BackColor)
            lngBrush = CreateSolidBrush(lngBackColor)
            'ʹ�ø�ˢ����䱳��ɫ
            lngOldBrush = SelectObject(lngDC, lngBrush)
            Call FillRect(hDC, T_ClientRect, lngBrush)
            '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
            Call SelectObject(lngDC, lngOldBrush)
            Call DeleteObject(lngBrush)
        
'            '��������
            Set stdset = New StdFont
            stdset.Name = "����"
            stdset.Size = 9
            stdset.Bold = False
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)

            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                lngColor = Val(Split(arrTmp(i), "-#")(0))
                '����������ɫ
                Call SetTextColor(lngDC, lngColor)
                strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") 'Ƥ�Խ��
                If i < UBound(arrTmp) Then strTmp = strTmp & ","
                If Trim(strTmp) <> "" Then
                    T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                    Call GetTextRect(objDraw, LPoint.X + lngX1, LPoint.Y, CStr(strTmp), , True)
                    Call DrawText(lngDC, CStr(strTmp), -1, T_LableRect, DT_CENTER)
                    lngX1 = lngX1 + T_Size.W
                End If
            Next i
           Call SelectObject(lngDC, lngOldFont)
           Call DeleteObject(lngFont)
        End If
    End If
    
    '���������
    If Col >= mshDownTab.FixedCols And Row >= mshDownTab.FixedRows Then
        strTmp = mshDownTab.TextMatrix(Row, Col)
        If AnsyGrade(Val(mshDownTab.RowData(Row)), strTmp, arrText) = True Then
            'lngColor = mshDownTab.Cell(flexcpForeColor, Row, Col, Row, Col)
            Call DrawDownTabAnsyGrade(lngDC, picMain, arrText, Row, Col, Left, Top, Right, Bottom, Done, mbln�೦�����ӷ�ĸ��ʾ)
        End If
    End If
End Sub

Private Sub mshUpTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strTime As String
    If NewRow = 0 And T_Patient.lng�༭ = 1 Then
        strTime = GetCurveDate(NewCol, mstr��ʼʱ��, gintHourBegin)
        If Format(Split(strTime, ";")(0), "YYYY-MM-DD") > Format(mstr����ʱ��, "YYYY-MM-DD") Then
            mshUpTab.FocusRect = flexFocusLight
        Else
            mshUpTab.FocusRect = flexFocusSolid
            If mblnKeyDown = True Then
                picDisplay.Left = ((((NewCol - 1) \ 6) + 1) * 6 - 1) * mshUpTab.ColWidth(NewCol) + mshUpTab.ColWidth(0)
                picDisplay.Top = (mshUpTab.RowHeight(NewRow) - picDisplay.Height) / 2
                picDisplay.Enabled = IIf(T_Patient.lng�༭ = 1, True, False)
            End If
        End If
    Else
        mshUpTab.FocusRect = flexFocusNone
    End If
    mblnKeyDown = False
End Sub

Private Sub mshUpTab_DblClick()
    If T_Patient.lng�༭ = 0 Then Exit Sub
    With mshUpTab
        If .Row = 0 And .FocusRect = flexFocusSolid Then
            RaiseEvent DbClickCur(mIntDataEditor)
        End If
    End With
End Sub

Private Sub mshUpTab_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim intMinCol As Long, intMaxCol As Long
    Dim i As Integer, j As Integer
    Dim strTmp As String
    Dim lngColor As Long, lngDC As Long
    Dim objDraw As Object, stdset As Object
    
    lngDC = hDC
    
    If picMain.Tag = "" Then Exit Sub
    If Row = mshUpTab.Rows - 1 And Col >= mshUpTab.FixedCols Then
        Set objDraw = picBack
        Call CalcMinMaxCol(picMain.Tag, intMinCol, intMaxCol)
        j = (Col - mshUpTab.FixedCols) Mod 6
        Select Case j
            Case 0
                strTmp = gintHourBegin + 4 * 0
                lngColor = &H8080FF
            Case 1
                strTmp = gintHourBegin + 4 * 1
                lngColor = &H8080FF
            Case 2
                strTmp = gintHourBegin + 4 * 2
                lngColor = &H80000012
            Case 3
                lngColor = &H80000012
                strTmp = gintHourBegin + 4 * 3
            Case 4
                lngColor = &H80000012
                strTmp = gintHourBegin + 4 * 4
            Case 5
                lngColor = &H8080FF
                strTmp = gintHourBegin + 4 * 5
        End Select
        '���ݲ�������ҹ��ʱ�䷶Χ����ʱ����ɫ
        lngColor = GetTimeColor(Val(strTmp))
        If Col >= intMinCol And Col <= intMaxCol Then
            lngColor = lngColor
        Else
            lngColor = RGB_FleetGRAY
        End If
        
        Call SetTextColor(lngDC, lngColor)
        Call GetTextRect(objDraw, Left, Top + (Bottom - Top) / 2, CStr(strTmp), Right - Left - 3, True)
        Call DrawText(lngDC, CStr(strTmp), -1, T_LableRect, DT_CENTER)
    End If
End Sub

Private Sub mshUpTab_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnKeyDown = False
    With mshUpTab
        If .Row = 0 And .FocusRect = flexFocusSolid Then
            Select Case KeyCode
                Case vbKeyReturn
                    Call mshUpTab_DblClick
                Case vbKeyLeft
                    mblnKeyDown = True
                Case vbKeyRight
                    mblnKeyDown = True
            End Select
        End If
    End With
End Sub

Private Sub picBack_Resize()
    Dim lngLeft As Long
    
    On Error Resume Next
    '�趨�����ڸ����ռ�ĳ�ʼλ��
    T_DrawClient.ƫ����Y = 0
    picMain.Move 0, 0
    picMain.BackColor = &H80000005
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    
    lngLeft = T_DrawClient.ƫ����X * T_TwipsPerPixel.X
    
    With vsb
        .Left = picBack.Width - .Width
        .Top = 0
        .Height = picBack.Height - hsb.Height
    End With
    
    With hsb
        .Left = 0
        .Top = picBack.Height - .Height
        .Width = picBack.Width - vsb.Width
    End With
    
    picCard(0).Move lngLeft, 10
    
    mshUpTab.Redraw = False
    mshDownTab.Redraw = False
    
    With mshUpTab
        .ColWidth(0) = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) * 15
        .Left = lngLeft
        .Top = picCard(0).Top + picCard(0).Height
        .RowHeight(3) = 400
        .Height = (3 * mshUpTab.RowHeight(0) + 520)
        .Width = ((T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) + T_DrawClient.�е�λ * 6 * 7 + 1) * T_TwipsPerPixel.X
        .ColWidthMin = T_DrawClient.�е�λ * Screen.TwipsPerPixelX
         picCard(0).Width = .Width
         .Refresh
    End With
    
    picDraw.Move 0, mshUpTab.Top + mshUpTab.Height, (T_DrawClient.��������.Right + 1) * T_TwipsPerPixel.X, _
        (T_DrawClient.�̶�����.Bottom - T_DrawClient.�̶�����.Top) * Screen.TwipsPerPixelY

    picDisplay.Height = 165
     
    With vsf
        .Top = mshUpTab.Top + mshUpTab.Height + (T_DrawClient.�̶�����.Bottom - T_DrawClient.�̶�����.Top) * Screen.TwipsPerPixelY
        .Left = lngLeft
        .Width = mshUpTab.Width
        .Height = .Body.RowHeight(vsf.FixedRows)
        .Visible = Not mbln��������
    End With
        
    With mshDownTab
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .Left = lngLeft
        .Top = mshUpTab.Top + mshUpTab.Height + (IIf(mbln�������� = False, vsf.Height, 0)) + (T_DrawClient.�̶�����.Bottom - T_DrawClient.�̶�����.Top) * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        .Width = mshUpTab.Width
        .Height = .Rows * .RowHeight(0)
        .Refresh
    End With
    
    lblCommText.Left = lngLeft
    lblCommText.Top = mshDownTab.Top + mshDownTab.Height
    lblCommText.Visible = True
    
    mshUpTab.Redraw = True
    mshDownTab.Redraw = True
    
    picMain.Width = mshUpTab.Width + mshUpTab.Left
    picMain.Height = lblCommText.Top + lblCommText.Height
    
    '���������
    Call CalcScrollBarSize
    
    '�������µ��Ŀɻ������С
    mlng�߶� = (picBack.Height - mshUpTab.Top - mshUpTab.Height - mshDownTab.Height - lblCommText.Height - _
        IIf(mbln�������� = False, vsf.Height, 0) - IIf(hsb.Visible = True, hsb.Height, 0)) / Screen.TwipsPerPixelY
    
    hsb.Value = 0
    vsb.Value = 0
End Sub

Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ� ���óɹ�����TRUE������FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    'ֻ����û��ʾ�������ǲ��������㲽��
    msinHStep = (picMain.Width - picBack.Width) / 100
    msinVStep = (picMain.Height - picBack.Height) / 100

    
    hsb.Max = 0 - Int(0 - ((picMain.Width - picBack.Width) / 300)) - 1
    vsb.Max = 0 - Int(0 - ((picMain.Height - picBack.Height) / 300)) - 1
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    
    With vsb
        .Height = picBack.Height - IIf(hsb.Visible = True, hsb.Height, 0)
    End With
    
    With hsb
        .Width = picBack.Width - IIf(vsf.Visible = True, vsb.Width, 0)
    End With
    
    '�㶨Ϊ100,ֻ�ǲ��������仯
    If hsb.Enabled Then
        hsb.Max = 100
        hsb.LargeChange = 100 / Int((Round((picMain.Width - picBack.Width) / picBack.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 100
        vsb.LargeChange = 100 / Int((Round((picMain.Height - picBack.Height) / picBack.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
    
End Function

 Private Sub lblCur_DblClick()
 
    If T_Patient.lng�༭ = 0 Then Exit Sub
    'RaiseEvent DbClickCur
End Sub

Private Sub mshUpTab_BeforeMouseDown(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single, _
                                     Cancel As Boolean)

    
    Dim strTemp   As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim intCOl As Long
    If Button <> vbLeftButton Then Exit Sub
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    '����ָ��������ſɽ��в���
    If T_Patient.lng�༭ = 1 Then
        intCOl = ((mintColMax - 1) \ 6 + 1) * 6
        
        If X > mshUpTab.ColWidth(0) And X < mshUpTab.ColWidth(0) + (intCOl * mshUpTab.ColWidth(intCOl)) Then
            '�������꣬������������
            strTemp = GetXCoordinate(X / T_TwipsPerPixel.X + mshUpTab.Left / T_TwipsPerPixel.X - 1, mstr��ʼʱ��, False)
            strTemp = mstr��ʼʱ�� & ";" & Split(strTemp, ",")(1)
            '����ʱ�������
            Call CalcMinMaxCol(strTemp, intMinCol, intMaxCol)
            picDisplay.Visible = True
            If Y < mshUpTab.RowHeight(0) + 40 Then
                picDisplay.Left = ((((intMaxCol - 1) \ 6) + 1) * 6 - 1) * mshUpTab.ColWidth(intMaxCol) + mshUpTab.ColWidth(0)
                picDisplay.Top = (mshUpTab.RowHeight(mshUpTab.FixedRows) - picDisplay.Height) / 2
                picDisplay.Enabled = IIf(T_Patient.lng�༭ = 1, True, False)
                mshUpTab.Col = intMaxCol
                mshUpTab.Row = mshUpTab.FixedRows
            End If
            
            If X > mshUpTab.ColWidth(0) + ((mintColMin - 1) * mshUpTab.ColWidth(mintColMin)) And X < mshUpTab.ColWidth(0) + ((mintColMax) * mshUpTab.ColWidth(mintColMax)) Then
                If Y > 3 * mshUpTab.RowHeight(0) Then
                    lblCur.Left = (intMaxCol - 1) * mshUpTab.ColWidth(intMaxCol) + mshUpTab.ColWidth(0)
                    '������ʾ
                    lblCur.Left = lblCur.Left + (mshUpTab.ColWidth(intMaxCol) - lblCur.Width) / 2
                    lblCur.Top = mshUpTab.Height - lblCur.Height
                End If
            End If
            
        End If
    End If
    
End Sub

Private Sub cboBaby_Click()
    Dim RS As New ADODB.Recordset
    
    If T_Patient.lngӤ�� = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    T_Patient.lngӤ�� = cboBaby.ItemData(cboBaby.ListIndex)
    
    On Error GoTo Errhand
    '������ȡ�ļ�ID
    mstrSQL = "select A.ID from ���˻����ļ� A,�����ļ��б� B" & _
       "    where A.����ID=[1] and A.��ҳId=[2] and nvl(A.Ӥ��,0)=[3] and A.��ʽID=B.ID and B.����=3 and B.����=-1"
    If mblnMoved = True Then
        mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    Set RS = zldatabase.OpenSQLRecord(mstrSQL, "��ȡ�ļ�ID", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
    
    If RS.BOF = False Then
        T_Patient.lng�ļ�ID = Val(zlCommFun.Nvl(RS("ID")))
        cboBaby.Enabled = True
    Else
        cboBaby.Enabled = False
        T_Patient.lngӤ�� = 0
        cboBaby.ListIndex = 0
    End If
   
    If Not InitBody(T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��) Then Exit Sub
    Call zlMenuClick("��ʾ��������")
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picDraw_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight And Shift = vbCtrlMask Then  '��һ��
        If mintPage < mcbrItem.Controls.Count - 1 Then
            mintPage = mintPage + 2
            mstrParam = mcbrItem.Controls.Item(mintPage).Parameter '�õ���ǰҳ��ʱ��
            Call InitWeekDays(mstrParam)
            mcbrToolBarҳ��.Caption = mcbrItem.Controls.Item(mintPage).Caption
            cbsMain.RecalcLayout
            Call zlMenuClick("װ������", mstrParam)
        End If

    ElseIf KeyCode = vbKeyLeft And Shift = vbCtrlMask Then

        If mintPage > 0 Then '��һ��
            mstrParam = mcbrItem.Controls.Item(mintPage).Parameter '�õ���ǰҳ��ʱ��
            Call InitWeekDays(mstrParam)
            mcbrToolBarҳ��.Caption = mcbrItem.Controls.Item(mintPage).Caption
            cbsMain.RecalcLayout
            Call zlMenuClick("װ������", mstrParam)
        End If

    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        'mblnAutoRedraw = mblnAutoRedraw Xor True
    End If

End Sub

Private Sub picDraw_Paint()
    '----------------------------------------------------------------------------
    '����:���ڴ���Copyͼ��PIC��
    '----------------------------------------------------------------------------
    picDraw.Cls
    Call BitBlt(mlngDC, 0, 0, T_ClientRect.Right, T_ClientRect.Bottom, mlngMemDC, 0, 0, SRCCOPY)
End Sub


Private Sub picCard_Paint(Index As Integer)
    Dim intLoop As Integer
    Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
    On Error Resume Next
    
    picCard(Index).Cls
    For intLoop = 0 To txtCard.UBound
        txtCard(intLoop).Height = 180
        If txtCard(intLoop).Visible Then
            X1 = txtCard(intLoop).Left
            Y1 = txtCard(intLoop).Top + txtCard(intLoop).Height + 15
            X2 = txtCard(intLoop).Left + txtCard(intLoop).Width
            Y2 = txtCard(intLoop).Top + txtCard(intLoop).Height + 15
            picCard(Index).ForeColor = &H8000000C
            picCard(Index).DrawStyle = 0
            picCard(Index).DrawWidth = 1
            picCard(Index).Line (X2, Y2)-(X1, Y1)
        End If
    Next
End Sub

Private Sub picCard_Resize(Index As Integer)
    On Error Resume Next
    txtCard(1).Move txtCard(1).Left, txtCard(1).Top, picCard(Index).Width - txtCard(1).Left - 45
    txtCard(7).Move txtCard(7).Left, txtCard(7).Top, picCard(Index).Width - txtCard(7).Left - 45
End Sub


Private Sub picSerach_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Call RaisEffect(picSerach, -2)
End Sub

Private Sub picSerach_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Call RaisEffect(picSerach, 2)
    Call cmdPrimitive_Click
End Sub

Private Sub txtCard_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtCard(Index))
End Sub

Private Sub UserControl_Initialize()
    picDraw.AutoRedraw = False
    Call InitCommandBar
    picBack.BackColor = &H80000005
    T_DrawClient.�е�λ = glngColStep
    T_DrawClient.ƫ����X = 5
    T_DrawClient.ƫ����Y = 0
    
    Call RaisEffect(picSerach, 2)
End Sub

Public Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = IIf(FontSize = 0, 9, IIf(FontSize = 1, 12, FontSize))
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "����"

    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont

    lblSerach(9).FontSize = bytFontSize
    cboBaby.FontSize = bytFontSize
    cboBaby.Top = (picTmp.Height - cboBaby.Height) \ 2
    lblSerach(9).Top = cboBaby.Top + (cboBaby.Height - lblSerach(9).Height) \ 2
End Sub


Private Sub UserControl_Resize()
    If UserControl.Parent.Visible = False Then Exit Sub
    
    If mblnAutoAdjust = True And Not mblnResize Then
        '���ʵ�ʴ�С�Ƿ����仯
        If Abs(mlngHeight - UserControl.Height) > 20 Then
            'Debug.Print "--��С�ı����--"
            Call zlMenuClick("װ������", mstrParam)
        End If
    End If
    
    Call RaisEffect(picSerach, 2)
    Call CalcScrollBarSize
End Sub

Private Sub UserControl_Terminate()
    Call ReleaseObj
End Sub

Private Sub vsb_Change()
    picMain.Top = -1 * vsb.Value * msinVStep
End Sub

Private Sub mfrmCaseTendBodyPrint_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

'------��ͼ��غ���

Private Sub Paint_Init(ByVal objDraw As Object, ByVal objBuffer As Object)

    On Error GoTo Errhand

    '��ͼǰ�ĳ�ʼ������
    '��Σ�������ľ��
    RGB_BLACK = RGB(0, 0, 0)
    RGB_RED = RGB(255, 0, 0)
    RGB_WRITE = RGB(255, 255, 255)
    RGB_BLUE = RGB(0, 0, 255)
    RGB_GRAY = &H808080
    RGB_FleetGRAY = &HC0C0C0
    mblnRedraw = True
    
    mlngHwnd = objDraw.hWnd
    Set mobjDraw = objDraw
    Set mobjBuffer = objBuffer
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    
    '�Ƚ��Զ����ͷ�
    Call Paint_Destory
    
    '�õ��ͻ�����
    Call GetClientRect(GetDesktopWindow, T_ClientRect)      'ȡ����Ļ����Ч����
    '�õ���ǰDC���
    mlngDC = GetDC(mlngHwnd)
    '��������DC
    mlngMemDC = CreateCompatibleDC(mlngDC)
    '��������λͼ����ֱ���ڴ�λͼ������
    mlngMemBitmap = CreateCompatibleBitmap(mlngDC, T_ClientRect.Right, T_ClientRect.Bottom) '������ԴDC���ܱ�֤�ǲ�ɫ��λͼ
    '�ڼ���DC��ʹ�ô����ļ���λͼ
    mlngOldBitmap = SelectObject(mlngMemDC, mlngMemBitmap)
    
    Call SetBkMode(mlngMemDC, TRANSPARENT)
    
    '������ʱˢ�����ñ���ɫ
    Dim lngBrush As Long, lngOldBrush As Long

    '������ɫˢ��
    lngBrush = GetStockObject(WHITE_BRUSH)
    'ʹ�ø�ˢ����䱳��ɫ��ȫ�ף�
    lngOldBrush = SelectObject(mlngMemDC, lngBrush)
    Call FillRect(mlngMemDC, T_ClientRect, lngBrush)
    '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
    Call SelectObject(mlngMemDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    '����������������Ŀ��ͼ��װ�����ڴ�
    Call PrepareGraph

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Reset()

    On Error GoTo Errhand

    '�ָ���������ʼ״̬
    
    '������ʱˢ�����ñ���ɫ
    Dim lngBrush As Long, lngOldBrush As Long

    '������ɫˢ��
    lngBrush = GetStockObject(WHITE_BRUSH)
    'ʹ�ø�ˢ����䱳��ɫ��ȫ�ף�
    lngOldBrush = SelectObject(mlngMemDC, lngBrush)
    Call FillRect(mlngMemDC, T_ClientRect, lngBrush)
    '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
    Call SelectObject(mlngMemDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    
    If Not mrsDrawItems Is Nothing Then If mrsDrawItems.State = 1 Then mrsDrawItems.Close
    '����������Ŀ����ͼ����(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ,��ɫ,������)
    gstrFields = "��Ŀ���," & adDouble & ",18|���ֵ," & adDouble & ",18|��Сֵ," & adDouble & ",18|" & _
        "��λֵ," & adDouble & ",18|���ֵ����," & adLongVarChar & ",20|��Сֵ����," & adLongVarChar & ",20|" & _
        "��λ�̶�," & adLongVarChar & ",20|��ʾģʽ," & adDouble & ",5|��ɫ," & adDouble & ",18"
    Call Record_Init(mrsDrawItems, gstrFields)
    
    mblnRedraw = True           '��mrsDrawItems���,����ǿ��ˢ��
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Destory()

    On Error GoTo Errhand

    '�������ж���
    If mlngOldBitmap <> 0 Then Call SelectObject(mlngMemDC, mlngOldBitmap)
    If mlngMemBitmap <> 0 Then Call DeleteObject(mlngMemBitmap)
    If mlngMemDC <> 0 Then Call DeleteDC(mlngMemDC)
    If mlngDC <> 0 Then Call ReleaseDC(mlngHwnd, mlngDC)
    mlngOldBitmap = 0
    mlngMemBitmap = 0
    mlngMemDC = 0
    mlngDC = 0
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Canvas(Optional ByVal blnAdjust As Boolean = False)
    '׼����������ɿ̶ȼ�������š����ⵥ�������Լ������趨���л�׼�ߵ���棩
    '��Сģʽ��,����ʾ���ϱ��,�ı�������
    'blnAdjust=False��ʾ�̶���С�����������������е���
    
    Static SlngMaxY As Long                 '��¼��һ�ε����߶ȣ��Ծ��������Ƿ���Ҫ�ػ�
    Dim lngCurX     As Long, lngCurY As Single  '��ǰλ��
    Dim lngMaxX     As Long, lngMaxY As Single  '�߽�
    Dim lngCurAlerY As Single
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln˫�� As Boolean                  '�˲������û�ָ��,bln˫��=TRUE��ʾֻ��ʾ����;������ʾʮ��
    Dim bln���� As Boolean                  '�˲������û�ָ��,���зֽ��Ǵ��߻���ϸ��
    Dim rsTemp        As New ADODB.Recordset
    
    '���¶��Ǳ�׼�߶�
    Dim intLineMode   As Integer
    Dim blnDoubleRow  As Boolean             '������Ϊһ�д�ӡ���
    Dim intTens_digit As Integer            '3����10�ı��������2����5�ı��������1���Ǹ�λ�����������
    Dim sinAlertness  As Single              '������,��������
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sigRowStepNew As Single, sinRowStep As Single, lngInitRowStep As Long
    Dim lng����� As Long, lngMaxRows As Long
    Dim lng������Сֵ As Long
    Dim arrTemp()     As String
    Dim sinY��λ As Single '���ߵ�λ�����Bottom
    Dim lngCurveRow As Long

    '�������ͼ�������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    Dim sin�̶� As Single, bln��ʾ�̶� As Boolean
    Dim sin�̶ȼ�� As Single, sinBegin�̶� As Single, dbl��λֵ As Double

    Dim str���ֵ���� As String, str��Сֵ���� As String

    On Error GoTo Errhand
    
    'ʵ�����ŵ�ԭ��˵����
    '1����ͨģʽ���������ݾ���ʾ
    '2����Сģʽ=2��ʱ��̶Ȳ���ʾ��ÿ��10С�и�Ϊ5С��
    '3����Сģʽ<=4��תΪ������ʾ
    
    '��ǰ�ǹ̶���������2����������ݣ����Դ˴���ȥ2��
    '�������ʱΪ�˶���ÿ����ٴμ�2�������
    lngCurveRow = Val(zldatabase.GetPara("�������߹̶��������", glngSys, 1255, "0"))
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    T_DrawClient.������ = glngMaxRows
    
    gstrSQL = " Select A.��Ŀ���,A.�������,A.��¼��,A.��¼��,A.��¼ɫ,nvl(A.���ֵ,0) ���ֵ,nvl(A.��Сֵ,0) ��Сֵ," & _
        "nvl(A.��λֵ,0) ��λֵ,A.�̶ȼ��,A.��ʾ��,C.��Ŀ��λ ��λ,nvl(A.�����,2)-2 AS �����,B.��λ " & _
        " From ���¼�¼��Ŀ A,���²�λ B,�����¼��Ŀ C" & _
        " Where A.��Ŀ���=B.��Ŀ���(+) And B.ȱʡ��(+)=1" & _
        " And A.��¼��=1 And A.��Ŀ���=C.��Ŀ��� and nvl(C.Ӧ�÷�ʽ,0)=1 and C.����ȼ�>=[1]" & _
        " and nvl(C.���ò���,0) in (0,[2]) and (C.���ÿ���=1 or (C.���ÿ���=2 and Exists (select 1 from �������ÿ��� D where C.��Ŀ���=D.��Ŀ��� and D.����ID=[3])))" & _
        " Order by �������"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ʼ��", T_Patient.lng����ȼ�, IIf(T_Patient.lngӤ�� = 0, 1, 2), T_Patient.lng����ID)
    
    '------------------------------------------------------------------------------------------------------------------
    rsTemp.Filter = "��Ŀ���=" & gint����
    '�����ӡ���������
    With rsTemp
        Do While Not .EOF
            lng����� = Val(zlCommFun.Nvl(!�����))
            If lng����� < 0 Then lng����� = 0
            
             '�޸�����51442
            If Val(zlCommFun.Nvl(!��Сֵ, 0)) > 34 Then
                lngMaxRows = lng����� + (Val(zlCommFun.Nvl(!���ֵ, 0)) - 35) / 0.1 + 10
            Else
                lngMaxRows = lng����� + (Val(zlCommFun.Nvl(!���ֵ, 0)) - Val(zlCommFun.Nvl(!��Сֵ, 0))) / 0.1
            End If

            lngMaxRows = lngMaxRows + lngCurveRow
            
            If lngMaxRows > T_DrawClient.������ Then
                T_DrawClient.������ = lngMaxRows
            End If
        .MoveNext
        Loop
    End With
    
    rsTemp.Filter = 0
    rsTemp.Sort = "�������"
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
    '����ֵ
    intLineMode = PS_SOLID
    lngLableStep = glngLableStep
    lngColStep = glngColStep
    lngInitRowStep = glngInitRowStep
    sigRowStepNew = lngInitRowStep
    intTens_digit = 3
    
    '���µ��Ե�����ʾ(������ѡ����˫����ʾ��û�����̶���ʾһ��) 1��������ʾ 0��˫����ʾ
    If zldatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, 0) = 1 Then
        bln˫�� = False
    Else
        bln˫�� = True
    End If
    'True��ʾ����ֻ���һ��,Ч����һ���̶�ֻ��ʾ������;����һ���̶���ʾʮ��,���û�������������,��blnDoubleRow�޹�
    bln���� = True
    
    If Not bln���� Then intLineMode = PS_DASHDOTDOT
    
    '�����
    intLables = rsTemp.RecordCount
    lngCurX = T_DrawClient.ƫ����X
    lngCurY = T_DrawClient.ƫ����Y
    lngMaxX = (intLables * lngLableStep) + (7 * 6 * lngColStep) + T_DrawClient.ƫ����X  '�̶�+7*��� +T_DrawClient.ƫ����X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.������ * sigRowStepNew + T_DrawClient.ƫ����Y '��Ϊ����С�����������ʼY����,����̶�����6��Ϊ���ʱ����Ϣ��
    
    '����������ݵ�У��
    If blnAdjust Then

        '���С�ڿɼ������С���������
        If lngMaxY > mlng�߶� Then
            lngMaxY = mlng�߶� - 2 * mintNullRow * lngInitRowStep
            sigRowStepNew = Round((lngMaxY) / T_DrawClient.������, 1)
        End If

        '����и�̫С���򽫷�����Ϊһ����ʾ
        If sigRowStepNew <= 2 Then
            sinRowStep = 1.5
            blnDoubleRow = True
        End If

        '����̶ȵ��������
        lngMaxY = T_DrawClient.������ * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + T_DrawClient.ƫ����Y + lngInitRowStep * 2 * mintNullRow

        If Not mblnRedraw Then mblnRedraw = (lngMaxY <> SlngMaxY)
        If sigRowStepNew < 4 Then intLineMode = PS_DOT
    End If
    
    '���п̶ȵ�У��(��������ĿС��3ʱ)
    If intLables <= 3 Then
        lngLableStep = glngLableWith / intLables
        lngMaxX = (intLables * lngLableStep) + (7 * 6 * lngColStep) + T_DrawClient.ƫ����X     '�̶�+7*��� +ƫ����X
    End If
    
    lblCommText.Caption = ""
    
    Call Paint_Reset                                                    '�������
    
    SlngMaxY = lngMaxY
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�е�λ = IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.˫�� = blnDoubleRow
    
    
    '���̶�����
'    For lngRow = 1 To intLables
'        Call DrawRect(mlngMemDC, lngCurX - IIf(lngRow = 1, 0, 1), lngCurY, lngCurX + lngLableStep + 1, lngMaxY, PS_SOLID, 1, RGB_BLACK)
'        lngCurX = lngCurX + lngLableStep
'    Next
    
    For lngRow = 1 To intLables
         Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
         lngCurX = lngCurX + lngLableStep
    Next
    Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
    
    '���̶ȿ�
    Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)

    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    T_DrawClient.�̶�����.Top = lngCurY
    T_DrawClient.�̶�����.Right = lngCurX
    T_DrawClient.�̶�����.Bottom = lngMaxY
    
    'Ĭ�����һ��������ʾ��Ŀ����
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    
    '�����µ�������
    For lngRow = 0 To T_DrawClient.������
        If lngRow <> 0 Then
            lngCurY = lngCurY + IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
        End If
        '�����µ���������
        If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
            Call DrawLine(mlngMemDC, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sigRowStepNew >= 4 And bln����, 2, 1), RGB_FleetGRAY)
        End If
    Next
    
    lngCurY = T_DrawClient.�̶�����.Top
    
    '�����µ�������
    For lngRow = 1 To 6 * 7
        lngCurX = lngCurX + lngColStep
        
        Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod 6 = 0, 2, 1), IIf(lngRow Mod 6 = 0, RGB_RED, RGB_GRAY))
    Next
    
    lngCurX = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.��������.Right = lngMaxX
    T_DrawClient.��������.Bottom = lngMaxY
    
    '�������������
    Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngMaxY - 1, lngMaxX, lngMaxY - 1, PS_SOLID, 1, RGB_BLACK)

    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
    With rsTemp
        Do While Not .EOF
            '��ʾ�̶ȿ���Ŀ�����Ƽ�����,�����¡�
            lngCurX = T_DrawClient.�̶�����.Left + ((.AbsolutePosition - 1) * T_DrawClient.�̶ȵ�λ)
            lngCurY = T_DrawClient.�̶�����.Top
            
            '���������С
            gstdSet.Name = "����"
            gstdSet.Size = 9
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)
            
            '���������Ŀ������
            Call SetTextColor(mlngMemDC, zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + mobjDraw.TextHeight(zlCommFun.Nvl(!��¼��)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!��¼��)), T_DrawClient.�̶ȵ�λ)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!��¼��)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            
            '���������С
            gstdSet.Name = "����"
            gstdSet.Size = 8
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)

            '�����Ŀ��λ
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + lngInitRowStep * 2 + mobjDraw.TextHeight(zlCommFun.Nvl(!��λ)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!��λ)), T_DrawClient.�̶ȵ�λ)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!��λ)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            mobjDraw.Font.Size = 9
            sinY��λ = T_LableRect.Bottom
            '��������
            gstdSet.Name = "����"
            gstdSet.Size = 9
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)
            '���������Ŀ�ı�ʶ��
            'Call DrawMarker(False, !��Ŀ���, zlCommFun.NVL(!��λ, "��"), lngCurX + T_TwipsPerPixel.x / 2, lngCurY + 15, "��", True)
    
            'ǿ���趨����������Ŀ����ʾģʽ
            Select Case !��Ŀ���

                Case gint����  '��������ʱ����̶�
                    intTens_digit = 1
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 1)
                    dbl��λֵ = 0.1
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 37)
                    arrTemp = Split(zlCommFun.Nvl(!��¼��, "��,��,��"), ",")
                    lblCommText.Caption = lblCommText.Caption & "��" & zlCommFun.Nvl(!��¼��) & "(����" & arrTemp(0) & ",Ҹ��" & arrTemp(1) & ",����" & arrTemp(2) & ")"

                Case gint����, gint����  '����/������10�ı�������̶�
                    intTens_digit = 3
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 10)
                    dbl��λֵ = 2
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    
                    If !��Ŀ��� = gint���� Then
                        lblCommText.Caption = lblCommText.Caption & "��" & zlCommFun.Nvl(!��¼��) & "(ȱʡ��¼��" & zlCommFun.Nvl(!��¼��, "+") & ",����H)"
                    Else
                        lblCommText.Caption = lblCommText.Caption & "��" & zlCommFun.Nvl(!��¼��) & "(" & zlCommFun.Nvl(!��¼��, "��") & ")"
                    End If

                Case gint����  '������5�ı�������̶�
                    mbln�������� = True
                    intTens_digit = 2
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 5)
                    dbl��λֵ = 1
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    lblCommText.Caption = lblCommText.Caption & "��" & zlCommFun.Nvl(!��¼��) & "(��������" & zlCommFun.Nvl(!��¼��, "*") & ",������R)"
                Case Else
                    intTens_digit = 1
                    dbl��λֵ = Val(zlCommFun.Nvl(!��λֵ, 1))
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, Val(zlCommFun.Nvl(!��λֵ, 0)) * 10)
                    If sin�̶ȼ�� > Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ)) Then
                        sin�̶ȼ�� = Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ))
                    End If
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    lblCommText.Caption = lblCommText.Caption & "��" & zlCommFun.Nvl(!��¼��) & "(" & zlCommFun.Nvl(!��¼��, "*") & ")"
            End Select

            '����ֵ
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow)   '�̶�ǰ2 * mintNullRow�еĸ߶Ȳ�����̶�

            '�������Сģʽ,�ӵ�30�п�ʼ�����ʶ
            'If blnDoubleRow Then lngCurY = lngCurY + lngInitRowStep * 2 * mintNullRow
            
            '��������ж�λ����Чλ��
            lngCurY = lngCurY + (T_DrawClient.�е�λ * zlCommFun.Nvl(!�����, 2))
            
            Do While True
                bln��ʾ�̶� = False
                If sin�̶� = 0 Then     '�ս���ѭ������ʱȡ�����ֵ
                    sin�̶� = zlCommFun.Nvl(!���ֵ, 0)
                    sinBegin�̶� = sin�̶�
                    str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                Else                    '����õ�ÿ���̶ȵ�ֵ
                    sin�̶� = sin�̶� - dbl��λֵ    '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                End If
                
                If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                
                If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - IIf(T_DrawClient.˫��, sin�̶ȼ�� * 2, sin�̶ȼ��)
                
                If sinBegin�̶� < 0 Then sinBegin�̶� = 0
                
                If bln��ʾ�̶� Then
                    '�������ֵ�������ߵ�λ�ظ�
                    If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY < sinY��λ Then
                        Call GetTextRect(mobjDraw, lngCurX, sinY��λ, Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ)
                    ElseIf Format(lngCurY, "#0") = T_DrawClient.�̶�����.Bottom Then
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY - (mobjDraw.TextHeight("1") / (T_TwipsPerPixel.Y * 2)), Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ)
                    Else
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY, Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ)
                    End If
                    Call DrawText(mlngMemDC, Format(sin�̶�, "#0"), -1, T_LableRect, DT_CENTER)
                End If
                
                '���������Ч��Χ�ڣ����߳����������˳�
                If Val(Format(sin�̶�, "#0.00")) <= Val(Format(!��Сֵ, "#0.00")) Or Format(lngCurY, "#0") > T_DrawClient.�̶�����.Bottom Then
                    str��Сֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                    '��Ӹ���Ŀ(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
                    gstrFields = "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ"
                    gstrValues = zlCommFun.Nvl(!��Ŀ���) & "|" & zlCommFun.Nvl(!���ֵ) & "|" & zlCommFun.Nvl(!��Сֵ) & "|" & dbl��λֵ & "|" & _
                        str���ֵ���� & "|" & str��Сֵ���� & "|" & T_DrawClient.�е�λ & "," & T_DrawClient.�е�λ & "|" & intTens_digit & "|" & !��¼ɫ
                    Call Record_Add(mrsDrawItems, gstrFields, gstrValues)
                    
                    '�����߻�ʾ��
                    If blnDoubleRow = False And (sinAlertness < Val(Nvl(!���ֵ)) And sinAlertness > Val(Nvl(!��Сֵ))) Then
                        lngCurAlerY = Val(GetYCoordinate(mobjDraw, mrsDrawItems, Val(Nvl(!��Ŀ���)), sinAlertness))
                        Call DrawLine(mlngMemDC, T_DrawClient.��������.Left, lngCurAlerY, lngMaxX, lngCurAlerY, PS_SOLID, 1, RGB_RED)
                    End If
                    Exit Do
                End If
                
                lngCurY = lngCurY + T_DrawClient.�е�λ
            Loop
            
            '��ԭ������Ϣ
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
                
            sinBegin�̶� = 0
            sin�̶� = 0                 '���ƴӵ�һ�п�ʼ���
            .MoveNext
        Loop
    End With
        
    '��������
    gstdSet.Name = "����"
    gstdSet.Size = 9
    Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngMemDC, mlngFont)
    
    lblCommText.Caption = "˵��:" & Mid(lblCommText.Caption, 2)
    mblnRedraw = False                      '����һ�κ�Ͳ��ٻ���
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub Paint_Date()
'-------------------------------------
'����:������ʾ�����µ��ϵ�ʱ��
'˵��:�˺���Ŀǰδʹ��
'------------------------------------
    Dim i As Long, j As Long
    Dim lngColor  As Long
    Dim intMinCol As Long, intMaxCol As Long
    Dim strTmp As String
    
    On Error GoTo Errhand
    
    If picMain.Tag <> "" Then
        Call CalcMinMaxCol(picMain.Tag, intMinCol, intMaxCol)
    End If
    
    For i = 1 To 7
        '������� ������Ϣ
        Call SetTextColor(mlngMemDC, RGB_BLACK)
        Call GetTextRect(mobjDraw, T_DrawClient.��������.Left + (i - 1) * 6 * T_DrawClient.�е�λ, T_DrawClient.��������.Top + T_DrawClient.ʱ���е�λ * 2, "����", T_DrawClient.�е�λ * 3)
        Call DrawText(mlngMemDC, "����", -1, T_LableRect, DT_CENTER)

        Call SetTextColor(mlngMemDC, RGB_BLACK)
        Call GetTextRect(mobjDraw, T_DrawClient.��������.Left + 3 * T_DrawClient.�е�λ + (i - 1) * 6 * T_DrawClient.�е�λ, T_DrawClient.��������.Top + T_DrawClient.ʱ���е�λ * 2, "����", T_DrawClient.�е�λ * 3)
        Call DrawText(mlngMemDC, "����", -1, T_LableRect, DT_CENTER)
        
        '���ʱ����Ϣ
        For j = 1 To 6

            Select Case j

                Case 1
                    strTmp = gintHourBegin + 4 * 0
                    lngColor = &H8080FF

                Case 2
                    strTmp = gintHourBegin + 4 * 1
                    lngColor = &H8080FF

                Case 3
                    strTmp = gintHourBegin + 4 * 2
                    lngColor = &H80000012

                Case 4
                    lngColor = &H80000012
                    strTmp = gintHourBegin + 4 * 0

                Case 5
                    lngColor = &H80000012
                    strTmp = gintHourBegin + 4 * 1

                Case 6
                    lngColor = &H8080FF
                    strTmp = gintHourBegin + 4 * 2
            End Select
            
            If j + (i - 1) * 6 >= intMinCol And j + (i - 1) * 6 <= intMaxCol Then
                lngColor = lngColor
            Else
                lngColor = RGB_FleetGRAY
            End If
            
            If picMain.Tag <> "" Then
                Call SetTextColor(mlngMemDC, lngColor)
                Call GetTextRect(mobjDraw, T_DrawClient.��������.Left + ((i - 1) * T_DrawClient.�е�λ * 6) + ((j - 1) * T_DrawClient.�е�λ), T_DrawClient.��������.Top + T_DrawClient.ʱ���е�λ * 6, Trim(strTmp), T_DrawClient.�е�λ)
                Call DrawText(mlngMemDC, Trim(strTmp), -1, T_LableRect, DT_CENTER)
            End If

        Next j
    Next i
    
    Exit Sub

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Construct()

    Dim lngRGB  As Long

    Dim blnLine As Boolean              '��������������ʱ,���ʲ�����

    Dim str���� As String               '��¼�����������ڵ���(X����)

    Dim strԭֵ As String, sinX����ԭ As Single, sinY����ԭ As Single
    
    Dim dbl��ֵ As Double, dblMinValue As Double, dblMaxValue As Double
    Dim bln�������� As Boolean
    Dim lng���²�����ʾ��ʽ As Long
    
    
    On Error GoTo Errhand

    '��ʼ��ͼ�����ͼ�����ݵ�������ص��Ĵ������¸��ˡ�ͼ�α�������������׾��
    '�Ȼ���(��������)
    '�ٴ���������׾
    '�����ͼ��
    
    lng���²�����ʾ��ʽ = Val(zldatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0"))
    
    With mrsPoint
        .Filter = ""
        '�Ȼ���
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "��Ŀ���,ʱ��"
        Do While Not .EOF
            If Val(zlCommFun.Nvl(!״̬)) <> 3 Then
                '�����µĺ��洦��,������
                If Not (!��Ŀ��� = gint���� And !��� = 1) Then
                    If strԭֵ <> !��Ŀ��� Then
                        dblMinValue = GetMinValue(!��Ŀ���)
                        dblMaxValue = GetMaxValue(!��Ŀ���)
                        blnLine = True
                        mrsDrawItems.Filter = "��Ŀ���=" & !��Ŀ���
                        If mrsDrawItems.RecordCount = 0 Then
                            '����������������
                            blnLine = False
                            mrsDrawItems.Filter = "��Ŀ���=" & gint����
                        End If

                        lngRGB = mrsDrawItems!��ɫ
                        mrsDrawItems.Filter = 0
                        
                        sinX����ԭ = 0
                        sinY����ԭ = 0
                        strԭֵ = !��Ŀ���
                    End If
                    
                    '����ϸ�
                    If !��Ŀ��� = gint���� And Val(zlCommFun.Nvl(!����)) = 1 Then
                        Call SetTextColor(mlngMemDC, lngRGB)
                        Call GetTextRect(mobjDraw, !X����, !Y���� - Screen.TwipsPerPixelY, "v", T_DrawClient.�е�λ, False)
                        Call DrawText(mlngMemDC, "v", -1, T_LableRect, DT_CENTER)
                    End If
                    
                    If sinX����ԭ <> 0 And blnLine Then
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y����, sinX����ԭ + T_DrawClient.�е�λ / 2, sinY����ԭ, PS_SOLID, 1, lngRGB)
                    End If
                    
                    If !�Ͽ� = 0 Then
                        sinX����ԭ = !X����
                        sinY����ԭ = !Y����
                    Else
                        sinX����ԭ = 0
                    End If
                    
                End If

                '�˴�������Ŀ�߳���Ŀ�����ֵ ��С����Ŀ��Сֵ
                If !��Ŀ��� = gint���� And Trim(Nvl(!��ֵ)) = "����" Then
                    dbl��ֵ = dblMinValue
                Else
                    dbl��ֵ = Val(zlCommFun.Nvl(!��ֵ))
                End If
                '�ص�ʱ����ſ�ǰ��Ϊ׼
                If !�ص� = 0 Then
                    If dbl��ֵ < dblMinValue Then
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + IIf(T_DrawClient.�е�λ < glngInitRowStep, glngInitRowStep, T_DrawClient.�е�λ) * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, 1, lngRGB, True)
                    End If
                    
                    If dbl��ֵ > dblMaxValue Then
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - IIf(T_DrawClient.�е�λ < glngInitRowStep, glngInitRowStep, T_DrawClient.�е�λ) * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, 1, lngRGB, True)
                    End If
                End If
            End If

            .MoveNext
        Loop
        
        '��������׾����(������������,����һ�����Ӧ��������������,����������������;����жϻ�ֻ�е���Ӧ,��ֻ��һ��)
        If .RecordCount <> 0 Then .MoveFirst
        .Filter = "��Ŀ���=" & gint����

        Do While Not .EOF
            str���� = str���� & "," & !X���� & ";" & !Y����
            .MoveNext
        Loop

        If str���� <> "" Then str���� = Mid(str����, 2)
        .Filter = 0

        '�γɷ���������
        If str���� <> "" Then Call CreatePoly(mrsPoint, mobjDraw, mlngMemDC, mstr��ʼʱ��, str����)

        '������ͼ��
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "��Ŀ���,ʱ��"

        Do While Not .EOF
            If Val(zlCommFun.Nvl(!״̬)) <> 3 Then
                If !��Ŀ��� = gint���� And !��� = 1 Then
                    '���µ������������ɫ�Ŀ���Բ
                    '�ַ����
                    Call SetTextColor(mlngMemDC, RGB_RED)
                    Call GetTextRect(mobjDraw, !X����, !Y����, "��", T_DrawClient.�е�λ)
                    Call DrawText(mlngMemDC, "��", -1, T_LableRect, DT_CENTER)
                    T_Size.H = mobjDraw.TextHeight("��") / Screen.TwipsPerPixelY
                    strԭֵ = Split(!��ע, ",")(0)
                    sinX����ԭ = Val(Split(Split(!��ע, ",")(1), ";")(0))
                    sinY����ԭ = Val(Split(Split(!��ע, ",")(1), ";")(1))
                    
                    If Val(!��ֵ) > Val(strԭֵ) Then
                        '������ʧ�ܣ�������ͷ�ĺ�ɫʵ�ߣ��ַ��̶��á�
                        'Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y����, sinX����ԭ + T_DrawClient.�е�λ / 2, sinY����ԭ, PS_SOLID, 1, RGB_RED, True)
                        '����ʧ��ҲΪ����(ҽԺҪ��)
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + (T_Size.H / 4), sinX����ԭ + T_DrawClient.�е�λ / 2, sinY����ԭ, PS_DOT, 1, RGB_RED, True)
                    ElseIf Val(!��ֵ) < Val(strԭֵ) Then
                        '�����³ɹ�������ɫ���ߣ��ַ��̶��á� ������ͷֱ������������
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - (T_Size.H / 2), sinX����ԭ + T_DrawClient.�е�λ / 2, sinY����ԭ, PS_DOT, 1, RGB_RED, False)
                    End If
                Else
                    If !��Ŀ��� = gint���� And Trim(Nvl(!��ֵ)) = "����" And (lng���²�����ʾ��ʽ = 0 Or lng���²�����ʾ��ʽ = 1) Then
                        bln�������� = False
                    Else
                        bln�������� = True
                    End If
                    
                    If !�ص� = 0 And bln�������� Then
                        Call DrawMarker(True, !��Ŀ���, !��λ, !X����, !Y����, !�ص���Ŀ, False, !����)
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Assistant()

    '��42�ȿ�ʼ��ӡ,���ж���������������ں�����д�ӡ,���һ�в��ؿ����ܷ��ȫ
    Dim rsNote As New ADODB.Recordset
    Dim bytδ��˵����ʾλ�� As Byte
    Dim Y As Long, X As Long, Y1 As Long
    Dim bln�ı� As Boolean
    Dim lngX          As Long, lngY As Long
    Dim strComment    As String, strTemp As String, strText As String
    Dim intNum        As Integer
    Dim intAscCharNum As Integer
    Dim varNote()     As String
    Dim i  As Integer, j As Integer

    On Error GoTo Errhand

    '������ͼ��������ֵ���������µĲ�������������С��10�����±���δ��˵���������������ת��������ҩ�ȣ�
    
    'mrsNote "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����"
    '����˵�� 2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��
    
    '����δ��˵����Ϣ
    mrsNote.Filter = "����=99"
    mrsNote.Sort = "X����"
    With mrsNote
        Do While Not .EOF
            If X = !X���� Then
                If InStr(1, "," & strTemp & ",", "," & zlCommFun.Nvl(!����) & ",") <> 0 Then
                    mrsNote.Delete
                Else
                    strTemp = strTemp & "," & zlCommFun.Nvl(!����)
                End If
            Else
                X = !X����
                strTemp = zlCommFun.Nvl(!����)
            End If
        .MoveNext
        Loop
    End With
    
    bytδ��˵����ʾλ�� = Val(zldatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0"))
    
    Y1 = GetYCoordinate(mobjDraw, mrsDrawItems, gint����, 42, mlngMemDC)
    
    mrsNote.Filter = 0
    mrsNote.Sort = "X����,��Ŀ���"
    With mrsNote
        Do While Not .EOF
            If !���� = 99 Then '���ݲ������ü��δ��˵����ʾ��ʽ
                varNote = Split(!����, ";")
                strComment = ""
                strTemp = ""

                For i = 0 To UBound(varNote)
                    'δ��˵����ʾ���Ϸ� ���²�����Ϊ�±꣬���������ϱ�
                    If Not (varNote(i) = "����" And bytδ��˵����ʾλ�� = 0) And varNote(i) <> "" Then
                        If InStr(1, strTemp, varNote(i)) = 0 Then
                            strTemp = IIf(strTemp = "", varNote(i), strTemp & ";" & varNote(i))
                        End If
                    End If
                Next i
                
                If strTemp <> "" Then
                    strComment = ""
                    varNote = Split(strTemp, ";")

                    For i = 0 To UBound(varNote)

                        If strComment = "" Then
                            strComment = varNote(i)
                        Else
                            strComment = strComment & " " & varNote(i)
                        End If

                    Next i

                End If
                
                '���ݲ����ж��Ƿ�ֱ����δ��˵��
                If bytδ��˵����ʾλ�� = 1 Then
                    Y = GetYCoordinate(mobjDraw, mrsDrawItems, gint����, 35, mlngMemDC)
                    If lngY <> 0 And X = Val(!X����) Then Y = lngY: strComment = " " & strComment
                    X = Val(!X����)

                    If strComment <> "" Then
                        strComment = Replace(strComment, ";", " ")
                        '�����Ϣδ��˵��
                        For i = 1 To Len(strComment)

                            If Y < T_DrawClient.�̶�����.Bottom Then
                                strText = Mid(strComment, i, 1)
                                Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)
                                '���������Ϣ
                                If T_DrawClient.�̶�����.Bottom - Y > T_Size.H Then
                                    Call DrawRotateText(mobjDraw, mlngMemDC, X, Y, strText, !��ɫ)
                                End If
                                If Asc(strText) < 0 Then
                                    Y = Y + T_Size.H
                                Else
                                    Y = Y + T_Size.H / 2
                                End If
                            End If

                        Next i

                        strComment = " "
                        lngY = Y
                    End If

                    mrsNote!���� = 1
                Else
                    mrsNote!���� = strComment
                    strComment = ""
                    mrsNote!Y���� = Y1
                    lngY = 0
                End If

            ElseIf !���� = 6 Then '����±�˵��
                Y = GetYCoordinate(mobjDraw, mrsDrawItems, gint����, 35, mlngMemDC)
                strComment = ""
                If lngY <> 0 And X = Val(!X����) Then Y = lngY: strComment = " "
                X = Val(!X����)
                
                '���δ��˵��������·����˴����������±��ϴ���δ��˵��,�Ա㱣֤��ʽ
                If strComment <> "" Then
                    If Asc(strComment) < 0 Then
                        intNum = 0
                    Else
                        intNum = 1
                    End If
                End If
                
                '�����Ϣδ��˵��
                strComment = strComment & !����
                intAscCharNum = 0
                If strComment <> "" Then
                    strComment = Replace(strComment, ";", " ")
                End If
                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.�̶�����.Bottom Then
                        strText = Mid(strComment, i, 1)
                        Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)

                        If Asc(strText) < 0 Then
                            If (intAscCharNum - intNum) Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If
                         
                        '���������Ϣ
                        If T_DrawClient.�̶�����.Bottom - Y > T_Size.H Then
                            Call DrawRotateText(mobjDraw, mlngMemDC, X, Y, strText, !��ɫ)
                        End If
                        If Asc(strText) < 0 Then
                            Y = Y + T_Size.H
                            intAscCharNum = 0
                        Else
                            Y = Y + T_Size.H / 2
                            intAscCharNum = intAscCharNum + 1
                        End If
                    End If
                Next i
                mrsNote!���� = 1
                lngY = 0
                strComment = ""
            Else
                mrsNote!Y���� = Y1
            End If

            .MoveNext
        Loop

    End With
    
    If mrsNote.RecordCount > 0 Then
        mrsNote.MoveFirst
        mrsNote.Update
    End If
    '������±����Ϣ(���ת���������� �ϱ����Ϣ)
    Call OutPutText(mobjDraw, mrsDrawItems, mlngMemDC, mrsNote, mstr��ʼʱ��)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ReleaseObj()
'----------------------------------------------------------
'�رյ����ж���
'----------------------------------------------------------
    If Not (mrsNote Is Nothing) Then Set mrsNote = Nothing
    If Not (mrsItems Is Nothing) Then Set mrsItems = Nothing
    If Not (mrsPoint Is Nothing) Then Set mrsPoint = Nothing
    If Not (mrsGraph Is Nothing) Then Set mrsGraph = Nothing
    If Not (mrsDrawItems Is Nothing) Then Set mrsDrawItems = Nothing
    If Not (mrsTabTime Is Nothing) Then Set mrsDrawItems = Nothing
    If Not (mrsCollect Is Nothing) Then Set mrsCollect = Nothing
    Set mobjDraw = Nothing
    Set mobjBuffer = Nothing
    Set gstdSet = Nothing
    
    mintPage = 0
    mintAllPage = 0
    mblnKeyDown = False
    mstrParam = ""
    mstrParam1 = ""
    mstrParam2 = ""
    If Not (mfrmParent Is Nothing) Then Set mfrmParent = Nothing
    
    Call Paint_Destory
End Sub

'-------------------���µ�������ȡ ����
Private Function GetMinValue(ByVal lng��Ŀ��� As Long) As Double
    '-------------------------------------------------------------------
    '����:������Ŀ��Ż�ȡ��Сֵ
    '˵��:Ŀǰ������Ŀֵ��ȷ����¼��ķ�Χ�����ֵ��Сֵȷ���������������С���꣬�����Լ�ͷ��ʾ
    '-------------------------------------------------------------------
    Dim dblvalue As Double
    
    mrsItems.Filter = "��Ŀ���=" & lng��Ŀ���
    If mrsItems.EOF Then Exit Function
    
'    If InStr(1, Nvl(mrsItems!��Ŀֵ��), ";") = 0 Then
'        dblvalue = Val(Nvl(mrsItems!��Сֵ, 0))
'    Else
'        dblvalue = Val(Split(mrsItems!��Ŀֵ��, ";")(0))
'    End If
    dblvalue = Val(Nvl(mrsItems!��Сֵ, 0))
    GetMinValue = dblvalue
End Function

Private Function GetMaxValue(ByVal lng��Ŀ��� As Long) As Double
    '-------------------------------------------------------------------
    '����:������Ŀ��Ż�ȡ���ֵ
    '˵��:Ŀǰ������Ŀֵ��ȷ����¼��ķ�Χ�����ֵ��Сֵȷ���������������С���꣬�����Լ�ͷ��ʾ
    '-------------------------------------------------------------------
    Dim dblvalue As Double
    Dim strValue As String
    
    mrsItems.Filter = "��Ŀ���=" & lng��Ŀ���
    If mrsItems.EOF Then Exit Function
    
'    If InStr(1, Nvl(mrsItems!��Ŀֵ��), ";") = 0 Then
'        dblvalue = Val(Nvl(mrsItems!���ֵ, 0))
'    Else
'        dblvalue = Val(Split(mrsItems!��Ŀֵ��, ";")(1))
'        If dblvalue = 0 Then dblvalue = Val(Nvl(mrsItems!���ֵ))
'    End If
    dblvalue = Val(Nvl(mrsItems!���ֵ, 0))
    strValue = Nvl(mrsItems!�ٽ�ֵ)
    If strValue <> "" And Val(strValue) <= Val(Nvl(mrsItems!���ֵ)) And Val(strValue) >= Val(Nvl(mrsItems!��Сֵ)) Then dblvalue = strValue
    GetMaxValue = dblvalue
End Function

Private Sub ReadBoyData(ByVal blnAutoAdjust As Boolean)
    
    On Error GoTo Errhand
    
    mint����Ӧ�� = 0
    If Not (mrsItems Is Nothing) Then If mrsItems.State = 1 Then mrsItems.Close
    '���ִ������øò��˵Ļ����¼��Ŀ
    gstrSQL = " Select C.��Ŀ���,C.��Ŀ����,C.��Ŀ����,C.��Ŀ����,C.��Ŀ����,C.��ĿС��,C.��Ŀ��ʾ,C.��Ŀ��λ,C.��Ŀֵ��,A.���ֵ,A.��Сֵ,A.�ٽ�ֵ,C.����ȼ�,C.Ӧ�÷�ʽ,C.���ò���" & _
              " From ���¼�¼��Ŀ A,�����¼��Ŀ C" & _
              " where A.��Ŀ���(+)=C.��Ŀ���" & _
              " And nvl(C.Ӧ�÷�ʽ,0)<>0" & _
              " and nvl(C.���ò���,0) in (0,[1])" & _
              " and (C.���ÿ���=1 or (C.���ÿ���=2 and Exists (select 1 from �������ÿ��� D where C.��Ŀ���=D.��Ŀ��� and D.����ID=[2])))" & _
              " Order by C.��Ŀ���"
              
    Set mrsItems = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ʼ��", IIf(T_Patient.lngӤ�� = 0, 1, 2), T_Patient.lng����ID)
    mrsItems.Filter = "��Ŀ���=-1"
    If mrsItems.RecordCount > 0 Then mint����Ӧ�� = zlCommFun.Nvl(mrsItems("Ӧ�÷�ʽ").Value, 2): mrsItems.Filter = ""
    
    '���е�ı��ּ���
    '   �ص��Ƿ��ص����.
    '   �ص���Ŀ��¼�ص���Ŀ
    '   �Ͽ�������:����һ��������,����δ��˵��
    '   ״̬:0-δ�༭;1-����;2-�޸�;3-ɾ��
    '   ��ע:������ʱ��¼ԭֵ
    '   ����:������ע���²���������ֵС�ڵ�����Ŀ��Сֵ���ڵ�����Ŀ���ֵ�ǵ��������.����Ĭ��Ϊ��
    If Not (mrsPoint Is Nothing) Then If mrsPoint.State = 1 Then mrsPoint.Close
    
    gstrFields = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",200|" & _
                 "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|" & _
                 "״̬," & adDouble & ",1|����," & adDouble & ",1|�Ͽ�," & adDouble & ",1|�ص���Ŀ," & adLongVarChar & ",50|" & _
                 "�ص�," & adDouble & ",5|X����," & adDouble & ",5|Y����," & adDouble & ",5|��ע," & adLongVarChar & ",50|" & _
                 "����," & adLongVarChar & ",10|��ʾ," & adDouble & ",1"
    Call Record_Init(mrsPoint, gstrFields)
    
    '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,13-����,99-δ��˵��)
    '���ñ�ʾ��Ϣ�Ƿ����
    
    If Not mrsNote Is Nothing Then If mrsNote.State = 1 Then mrsNote.Close
    
    gstrFields = "ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|����," & adDouble & ",2|" & _
        "����," & adLongVarChar & ",200|��ɫ," & adLongVarChar & ",20|X����," & adDouble & ",20|" & _
        "Y����," & adDouble & ",20|�߶�," & adDouble & ",20|��ӡX����," & adDouble & ",20|" & _
        "����," & adInteger & ",1|��ʾ," & adDouble & ",1"
        
    Call Record_Init(mrsNote, gstrFields)
    
    '������������
    Call SaveMemory(blnAutoAdjust)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveMemory(ByVal blnAutoAdjust As Boolean)
    Dim bytShow As Byte
    Dim strPart As String
    Dim strValue As String  '�����µ�ԭֵ
    Dim blnAdd As Boolean, bln���� As Boolean
    Dim SinX As Single, sinY As Single
    Dim strʱ�� As String, str���� As String
    Dim rsData As New ADODB.Recordset
    Dim rsPart As New ADODB.Recordset
    Dim strSql As String
    Dim lngColor As Long, lng�к� As Long, lng��Ŀ���  As Long
    Dim str���� As String
    Dim dbl��ֵ As Double, dblMinValue As Double, dblMaxValue As Double
    Dim strTmpString0 As String, strTmpString1 As String, strTmpString2 As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim arrValues() As String
    Dim arrTmpValue() As Variant, arrTmpNote As Variant
    Dim i As Integer
    Dim int��ʾ As Integer
    Dim rs���� As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim blnӤ�����µ���ʾ��Ժ As Boolean, bln�����ʾ��Ժ As Boolean
    Dim lng���²�����ʾ��ʽ As Long
    Dim int��� As Integer
    On Error GoTo Errhand
    
    '��¼������Ϣ
    strFileds = "��Ŀ���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|X����," & adDouble & ",5|ʱ��," & adLongVarChar & ",20"
    Call Record_Init(rs����, strFileds)
    
    '��ȡ���в�λ��Ϣ
    strSql = "Select ��Ŀ���, ��λ,ȱʡ�� From ���²�λ"
    Call zldatabase.OpenRecordset(rsPart, strSql, "��ȡ���²�λ")
    
    
    '����������Ŀ��Ҫ�����ֶ����������������Ƿ������ⵥ����ʾ,ĿǰȱʡΪ��ʾ
    '-----------------------------------------------------------------------
    gstrSQL = "SELECT C.ID ���,a.����ʱ�� As ʱ��,C.��ʾ,C.��¼���� As ��ֵ,C.���²�λ,C.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵�� " & _
                "FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                "Where B.ID=A.�ļ�ID " & _
                    "And A.ID = C.��¼ID " & _
                    "AND B.ID=[1] " & _
                    "AND Nvl(B.Ӥ��,0)=[6] " & _
                    "AND B.����id=[2] " & _
                    "AND B.��ҳid=[3] " & _
                    "AND D.��Ŀ���=c.��Ŀ��� " & _
                    "AND c.��¼����=1 " & _
                    "AND E.��Ŀ���=D.��Ŀ��� " & _
                    "AND E.����ȼ�>=[7]  " & _
                    "AND a.����ʱ�� BETWEEN [4] And [5] And c.��ֹ�汾 Is Null " & _
                    "AND D.��¼��=1 " & _
                "Order By A.����ʱ��,DECODE(C.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���)"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ����", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), T_Patient.lngӤ��, T_Patient.lng����ȼ�)
        
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
    
    strTmpString0 = ""
    strTmpString1 = ""
    strTmpString2 = ""
    With rsData
        Do While Not .EOF
            str���� = ""
            blnAllow = False
            strPart = zlCommFun.Nvl(!���²�λ)
            lng��Ŀ��� = Val(zlCommFun.Nvl(!��Ŀ���))
            Select Case lng��Ŀ���
                Case gint����
                    int��� = 1
                Case Else
                    int��� = Val(Nvl(!��¼���))
            End Select
            If strPart = "" Then
                rsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
                If rsPart.BOF = False Then
                    strPart = zlCommFun.Nvl(rsPart!��λ)
                Else
                    Select Case lng��Ŀ���
                        Case gint����
                            strPart = "Ҹ��"
                        Case gint����
                            strPart = "��������"
                        Case Else
                            strPart = ""
                    End Select
                End If
            End If
            
            mrsItems.Filter = "��Ŀ���=" & lng��Ŀ���
            If mrsItems.RecordCount > 0 Then
                SinX = GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinate(SinX, Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinate(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
                
                '��¼����������Ϣ
                If lng��Ŀ��� = gint���� Then
                    strFileds = "��Ŀ���|��ֵ|X����|ʱ��"
                    strValues = lng��Ŀ��� & "|" & zlCommFun.Nvl(!��ֵ) & "|" & SinX & "|" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                    Call Record_Add(rs����, strFileds, strValues)
                End If
                
                If (Not IsNull(!δ��˵��)) And zlCommFun.Nvl(!��ֵ) <> "����" Then
                    mrsNote.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���)) & " AND X����=" & SinX
                    blnAdd = (mrsNote.RecordCount = 0)
                    '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��)
                    gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����|��ʾ"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
                    gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & !��Ŀ��� & "|99|" & _
                        !δ��˵�� & "|" & RGB_BLUE & "|" & SinX & "|0|0|0|0|" & Val(zlCommFun.Nvl(!��ʾ))
                   
                    If blnAdd Then
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                         Call Record_Add(mrsNote, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(mrsNote!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(mrsNote!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                             blnAllow = GetCanvasCenter(CDate(Format(mrsNote!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                        If blnAllow = True Then
                            If Val(mrsNote!��ʾ) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(mrsNote, gstrFields, gstrValues, "ʱ��|" & Format(mrsNote!ʱ��, "yyyy-MM-dd HH:mm:ss"))
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                gstrFields = "��ʾ"
                                gstrValues = "2"
                                Call Record_Update(mrsNote, gstrFields, gstrValues, "ʱ��|" & Format(mrsNote!ʱ��, "yyyy-MM-dd HH:mm:ss"))
                            End If
                        End If
                        
                    End If
                Else
                    blnAdd = False
                    
                    mrsPoint.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���, 0)) & " AND X����=" & SinX & " And ���=" & int���
                    blnAdd = (mrsPoint.RecordCount = 0)
                    
                    dbl��ֵ = Val(zlCommFun.Nvl(!��ֵ))
                    
                    dblMinValue = GetMinValue(!��Ŀ���)
                    dblMaxValue = GetMaxValue(!��Ŀ���)
                    
                    '��ָ�����ţ���Ŀ��ֵ�������ֵ����Сֵ����Ŀ���������ʾ
                    If dbl��ֵ <= dblMinValue Then
                        dbl��ֵ = dblMinValue
                        'str���� = "��"
                    End If
                    
                    If dbl��ֵ >= dblMaxValue Then
                        dbl��ֵ = dblMaxValue
                        'str���� = "��"
                    End If
                    
                     '���²���������ʾ��35�̶�
                    If Trim(Nvl(!��ֵ)) = "����" And lng��Ŀ��� = gint���� Then dbl��ֵ = 35
                    sinY = Val(GetYCoordinate(mobjDraw, mrsDrawItems, !��Ŀ���, dbl��ֵ, mlngMemDC, True))
                     
                    gstrFields = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|״̬|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����|��ʾ"
                    gstrValues = Val(zlCommFun.Nvl(!���)) & "|" & !��ֵ & "|" & strPart & "|" & int��� & "|" & _
                                 Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng��Ŀ��� & "|0|" & Val(zlCommFun.Nvl(!���Ժϸ�, 0)) & "|" & IIf(zlCommFun.Nvl(!��ֵ, 0) = "����", 1, 0) & "|��|0|" & _
                                 SinX & "|" & sinY & "||" & str���� & "|" & Val(zlCommFun.Nvl(!��ʾ, 0))
                    If blnAdd Then '���
                        Call Record_Add(mrsPoint, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(mrsPoint!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(mrsPoint!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                            blnAllow = GetCanvasCenter(CDate(Format(mrsPoint!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                        If blnAllow = True Then
                            If Val(mrsPoint!��ʾ) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(mrsPoint, gstrFields, gstrValues, "���|" & mrsPoint!���)
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                gstrFields = "��ʾ"
                                gstrValues = "2"
                                Call Record_Update(mrsPoint, gstrFields, gstrValues, "���|" & mrsPoint!���)
                            End If
                        End If
                    End If
                End If
            End If
            mrsItems.Filter = 0
        .MoveNext
        Loop
    End With
    
     '�����Ѿ��õ���������Ŀ��������Ϣ���������������º���������������
    arrTmpValue = Array()
    If mint����Ӧ�� = 2 Then
        mrsPoint.Filter = "��Ŀ���=" & gint����
        With mrsPoint
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !X���� & ";" & Format(!ʱ��, "yyyy-MM-DD HH:mm:ss")
            .MoveNext
            Loop
        End With
    End If
    mrsPoint.Filter = ""
    
    '������Ϊ��������ʱ����������Ƿ�����Ϊ����
    mrsItems.Filter = "��Ŀ���=" & gint����
    If mrsItems.RecordCount > 0 Then
        For i = 0 To UBound(arrTmpValue)
            '��������Ƿ����������Ӧ
            rs����.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
            mrsPoint.Filter = "��Ŀ���=" & gint���� & " and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
            If mrsPoint.RecordCount = 0 Then
                If rs����.RecordCount = 0 Then
                    mrsPoint.Filter = ""
                    gstrFields = "��Ŀ���": gstrValues = gint����
                    Call Record_Update(mrsPoint, gstrFields, gstrValues, "���|" & Val(Split(CStr(arrTmpValue(i)), ";")(0)))
                Else
                    mrsPoint.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                    mrsPoint.Delete
                End If
            End If
        Next i
    End If
    
    If mint����Ӧ�� = 2 Then
        Set rs���� = New ADODB.Recordset
        strFileds = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",200|" & _
                    "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|" & _
                    "״̬," & adDouble & ",1|����," & adDouble & ",1|�Ͽ�," & adDouble & ",1|�ص���Ŀ," & adLongVarChar & ",50|" & _
                    "�ص�," & adDouble & ",5|X����," & adDouble & ",5|Y����," & adDouble & ",5|��ע," & adLongVarChar & ",50|" & _
                    "����," & adLongVarChar & ",10|��ʾ," & adDouble & ",1"
        Call Record_Init(rs����, strFileds)
        
        mrsPoint.Filter = "��Ŀ���=" & gint����
        With mrsPoint
            Do While Not .EOF
                rs����.AddNew
                For i = 0 To .Fields.Count - 1
                    rs����.Fields(.Fields(i).Name).Value = .Fields(i).Value
                Next i
                rs����.Update
            .MoveNext
            Loop
        End With
        
        mrsPoint.Filter = "��Ŀ���=" & gint����
        Do While Not mrsPoint.EOF
            mrsPoint.Delete
            mrsPoint.MoveNext
        Loop
        
        rs����.Filter = ""
        rs����.Sort = "ʱ��"
        With rs����
            Do While Not .EOF
                blnAdd = False
                blnAllow = False
                
                SinX = Val(zlCommFun.Nvl(!X����))
                sinY = Val(zlCommFun.Nvl(!Y����))
                mrsPoint.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���, 0)) & " AND X����=" & SinX
                blnAdd = IIf(mrsPoint.RecordCount = 0, True, False)
                
                strFileds = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|״̬|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����|��ʾ"
                strValues = Val(zlCommFun.Nvl(!���)) & "|" & !��ֵ & "|" & zlCommFun.Nvl(!��λ) & "|" & Val(zlCommFun.Nvl(!���, 0)) & "|" & _
                             Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!��Ŀ���)) & "|0|0|" & Val(zlCommFun.Nvl(!�Ͽ�)) & "|��|0|" & _
                             SinX & "|" & sinY & "||" & zlCommFun.Nvl(!����) & "|" & Val(zlCommFun.Nvl(!��ʾ, 0))
                
                If blnAdd Then '���
                    Call Record_Add(mrsPoint, strFileds, strValues)
                Else
                    If (zlCommFun.Nvl(mrsPoint!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(mrsPoint!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                        blnAllow = GetCanvasCenter(CDate(Format(mrsPoint!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
                    ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                        blnAllow = True
                    End If
                    
                    '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                    If blnAllow = True Then
                        If Val(mrsPoint!��ʾ) = 2 Then
                            arrValues = Split(strValues, "|")
                            arrValues(UBound(arrValues)) = 2
                            strValues = Join(arrValues, "|")
                        End If
                        Call Record_Update(mrsPoint, strFileds, strValues, "���|" & mrsPoint!���)
                    Else
                        If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                            strFileds = "��ʾ"
                            strValues = "2"
                            Call Record_Update(mrsPoint, strFileds, strValues, "���|" & mrsPoint!���)
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
    End If
    
    '��������������
    arrTmpValue = Array()
    mrsPoint.Filter = "��Ŀ���=1 and ���=0"
    With mrsPoint
        Do While Not .EOF
            ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
            arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
        .MoveNext
        Loop
    End With
    
    mrsPoint.Filter = "��Ŀ���=1"
    If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
    For i = 0 To UBound(arrTmpValue)
        mrsPoint.Filter = "��Ŀ���=1 and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
        If mrsPoint.RecordCount <> 0 Then
            gstrFields = "��ע": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
            Call Record_Update(mrsPoint, gstrFields, gstrValues, "���|" & zlCommFun.Nvl(mrsPoint!���))
        End If
    Next i
    
    arrTmpValue = Array()
    mrsPoint.Filter = "��Ŀ���=1 and ���=1"
    With mrsPoint
        Do While Not .EOF
            ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
            arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
        .MoveNext
        Loop
    End With
    
    mrsPoint.Filter = "��Ŀ���=1"
    If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
    For i = 0 To UBound(arrTmpValue)
        mrsPoint.Filter = "��Ŀ���=1 and ���=0 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
        If mrsPoint.RecordCount = 0 Then
            mrsPoint.Filter = "��Ŀ���=1 and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            mrsPoint.Delete
        End If
    Next i
    
    
    'ɾ����ʾΪ2������
    mrsPoint.Filter = "��ʾ=2"
    Do While Not mrsPoint.EOF
        mrsPoint.Delete
    mrsPoint.MoveNext
    Loop
    
    mrsNote.Filter = ""
    mrsNote.Filter = "��ʾ=2"
    Do While Not mrsNote.EOF
        mrsNote.Delete
    mrsNote.MoveNext
    Loop

    '����δ��˵�����������ݸ���ʾ��һ��
    mrsNote.Filter = ""
    mrsPoint.Filter = ""
    
    arrTmpValue = Array()
    arrTmpNote = Array()
    mrsNote.Sort = "��Ŀ���,X����"
    With mrsNote
        Do While Not .EOF
            blnAllow = False
            SinX = Val(!X����)
            mrsPoint.Filter = "��Ŀ���=" & Val(!��Ŀ���) & " And X����=" & SinX
            If mrsPoint.RecordCount > 0 Then
                If (zlCommFun.Nvl(mrsPoint!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(mrsPoint!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                    blnAllow = GetCanvasCenter(CDate(Format(mrsPoint!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
                ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                    blnAllow = True
                End If
                If blnAllow = True Then
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !��Ŀ��� & ";" & SinX
                Else
                    ReDim Preserve arrTmpNote(UBound(arrTmpNote) + 1)
                    arrTmpNote(UBound(arrTmpNote)) = !��Ŀ��� & ";" & SinX
                End If
            End If
        .MoveNext
        Loop
    End With
    
    For i = 0 To UBound(arrTmpValue)
        mrsPoint.Filter = "��Ŀ���=" & Val(Split(CStr(arrTmpValue(i)), ";")(0)) & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(1))
        Do While Not mrsPoint.EOF
            mrsPoint.Delete
        mrsPoint.MoveNext
        Loop
    Next i
    
    For i = 0 To UBound(arrTmpNote)
        mrsNote.Filter = "��Ŀ���=" & Val(Split(CStr(arrTmpNote(i)), ";")(0)) & " And X����=" & Val(Split(CStr(arrTmpNote(i)), ";")(1))
        Do While Not mrsNote.EOF
            mrsNote.Delete
        mrsNote.MoveNext
        Loop
    Next i
    
    '�������²��� ����Ϊ������Ҫ��35��������������²�������
    mrsPoint.Filter = "��Ŀ���=" & gint���� & " and ��ֵ='����' and ���<>1"
    mrsPoint.Sort = "ʱ��"
    With mrsPoint
        Do While Not .EOF
            strTmpString0 = strTmpString0 & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!��Ŀ���)) & "|99|" & _
                  "����|" & RGB_BLUE & "|" & !X���� & "|0|0|0|0"
            strTmpString2 = strTmpString2 & ";" & !X����
        .MoveNext
        Loop
    End With
    '��ȡ���������±���Ϣ
    '-----------------------------------------------------------------------
    gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
    gstrSQL = "" & _
             " Select A.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.��¼����,C.��Ŀ����,C.δ��˵��" & _
             " FROM ���˻����ļ� B, ���˻������� A, ���˻�����ϸ C" & _
             " Where B.ID=A.�ļ�ID And A.ID = C.��¼ID AND B.ID=[1] AND Nvl(B.Ӥ��, 0)=[6] AND B.����id=[2] AND B.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
             " AND MOD(C.��¼����,10) <> 1  AND A.����ʱ�� BETWEEN [4]  And [5]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���������±����Ϣ", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, Int(CDate(mstr��ʼʱ��)), CDate(mstr����ʱ��), T_Patient.lngӤ��, T_Patient.lng����ȼ�)
    
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
        
    With rsData
        Do While Not .EOF
            bytShow = 1
            str���� = Trim(zlCommFun.Nvl(!��¼����))
            
            lng�к� = IIf(!��¼���� = 2, 10, IIf(!��¼���� = 6, 11, 4))
            
            '����������ʾ��Ҫ���⴦��
            If !��¼���� = 4 Then
                str���� = Trim(zlCommFun.Nvl(!��Ŀ����))
                
                If str���� = "����" Then
                    bytShow = T_BodyFlag.����
                Else
                    bytShow = T_BodyFlag.����
                End If
                
                If bytShow = 2 And Not blnAutoAdjust Then
                    str���� = str���� & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                Else
                    str���� = !��Ŀ����
                End If
                lngColor = RGB_RED
            Else
                lngColor = IIf(Not IsNumeric(Nvl(!δ��˵��)), RGB_BLUE, Val(Nvl(!δ��˵��)))
            End If
            
            If bytShow > 0 Then
                SinX = Val(GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), mstr��ʼʱ��))
                
                mrsNote.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=" & !��¼����
                If mrsNote.BOF Then
                    gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|" & !��¼���� & "|" & _
                        str���� & "|" & lngColor & "|" & SinX & "|0|0|0|0"
                    Call Record_Add(mrsNote, gstrFields, gstrValues)
                Else
                    mrsNote!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                    mrsNote!���� = str����
                    mrsNote.Update
                End If
            End If
            mrsNote.Filter = 0
            .MoveNext
        Loop
    End With
    
    blnӤ�����µ���ʾ��Ժ = (zldatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, 1) = 1)
    
    bln�����ʾ��Ժ = False
    If CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) Then
        bln�����ʾ��Ժ = True
    ElseIf CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) = CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) And T_BodyFlag.��Ժ = 0 Then
        bln�����ʾ��Ժ = True
    End If
    
    '��ȡ���ת����Ϣ
    '-----------------------------------------------------------------------
    '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��)
    '1-��Ժ��2-��ƣ�3-ת�ƣ�4-����
    gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
    Set rsData = GetDataFromHis(T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), 2) ' Int(cdate(mstr��ʼʱ��))
    With rsData
        Do While Not .EOF
            If Trim(zlCommFun.Nvl(!����)) <> "" Then
                bytShow = 0
                lng�к� = Val(!�к�)
                str���� = zlCommFun.Nvl(!����)
                Select Case lng�к�
                Case 5
                    bytShow = T_BodyFlag.��Ժ
                Case 6, 3 '6ת�룬3ת��
                    bytShow = T_BodyFlag.ת��
                Case 7
                    bytShow = T_BodyFlag.����
                Case 8
                    bytShow = T_BodyFlag.��Ժ
                    If T_Patient.lngӤ�� > 0 Then
                        bytShow = IIf(blnӤ�����µ���ʾ��Ժ, bytShow, 0)
                    End If
                Case 9
                    bytShow = T_BodyFlag.���
                End Select
                 
                If bytShow > 0 Then
                    'Ŀǰ3��4 �����ת�� 3-��ʾ˵���Ϳ��� 4 ��ʾ˵�������ң�ʱ��
                    If lng�к� = 9 And bln�����ʾ��Ժ = True Then
                        str���� = "��Ժ"
                    End If
                
                    If bytShow = 2 Then
                        str���� = str���� & IIf(blnAutoAdjust = False, gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm")), "")
                    ElseIf bytShow = 3 Then
                        str���� = str���� & IIf(blnAutoAdjust = False, gstrCaveSplit & zlCommFun.Nvl(!����), "")
                    ElseIf bytShow = 4 Then
                        str���� = str���� & IIf(blnAutoAdjust = False, gstrCaveSplit & zlCommFun.Nvl(!����) & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm")), "")
                    ElseIf bytShow = 1 Then
                        str���� = str����
                    End If
                    
                    '���µ�����ģʽ�� ��������ʾ����
                    If bytShow = T_BodyFlag.���� And blnAutoAdjust = True Then
                        If InStr(1, str����, "(") <> 0 Then
                            str���� = Split(str����, "(")(0)
                        End If
                    End If
                    
                    SinX = Val(GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), mstr��ʼʱ��))
                    mrsNote.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=3"
                    
                    If mrsNote.BOF Then
                        gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|3|" & _
                            str���� & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                        Call Record_Add(mrsNote, gstrFields, gstrValues)
                    Else
                        mrsNote!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                        mrsNote!���� = str����
                        mrsNote.Update
                    End If
                End If
                mrsNote.Filter = 0
            End If
            .MoveNext
        Loop
    End With
    
    '��ȡӤ��������Ϣ
    If T_Patient.lngӤ�� > 0 Then
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
        Set rsData = GetDataFromHis(T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), 3)
        With rsData
            Do While Not .EOF
                bytShow = 0
                If Trim(zlCommFun.Nvl(!����)) <> "" Then
                    lng�к� = 12
                    bytShow = T_BodyFlag.����
                    If bytShow > 0 Then
                        Select Case bytShow
                            Case 1
                                str���� = zlCommFun.Nvl(!����)
                            Case 2
                                If Not blnAutoAdjust Then
                                    str���� = zlCommFun.Nvl(!����) & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                                Else
                                    str���� = zlCommFun.Nvl(!����)
                                End If
                        End Select
                        
                        SinX = Val(GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), mstr��ʼʱ��))
                        mrsNote.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=13"
                        
                        If mrsNote.BOF Then
                            gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|13|" & _
                                str���� & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                            Call Record_Add(mrsNote, gstrFields, gstrValues)
                        Else
                            mrsNote!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                            mrsNote!���� = str����
                            mrsNote.Update
                        End If
                    End If
                End If
                mrsNote.Filter = 0
            .MoveNext
            Loop
        End With
    End If
    
    str���� = ""
    Dim bytTag As Byte
    '51512,������,2012-07-11,δ��˵����ʾλ�� 0-��ʾ������,1-��ʾ������,2-����ʾ
    '��ҽ��ԺҪ��δ��˵������ʾ������ע��δ�ǵ����ߵ��������߲�����
    bytTag = Abs(Val(zldatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0")))
    lng���²�����ʾ��ʽ = Val(zldatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0"))
    '�������²��� ���²���ʼ����ʾ�� 35 �����棬ֻ��δ��˵����ʾ�������������Ž���������δ��˵���У���������������±���
    If Left(strTmpString0, 1) = ";" Then
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"
        If lng���²�����ʾ��ʽ = 0 Or lng���²�����ʾ��ʽ = 2 Then
            arrValues = Split(strTmpString0, "|")
            arrValues(3) = "�� "
            strTmpString0 = Join(arrValues, "|")
        End If
        strTmpString0 = Mid(strTmpString0, "2")
        strTmpString2 = Mid(strTmpString2, 2)
        For i = 0 To UBound(Split(strTmpString0, ";"))
            str���� = Split(strTmpString0, ";")(i)
            mrsNote.Filter = "����=" & IIf(bytTag = 1, 99, 6) & " and X����=" & Val(Split(strTmpString2, ";")(i))
            mrsNote.Sort = "��Ŀ���"
            If mrsNote.RecordCount > 0 Then
                mrsNote!���� = IIf(lng���²�����ʾ��ʽ = 0 Or lng���²�����ʾ��ʽ = 2, "�� ", "����") & IIf(bytTag = 1, ";", " ") & zlCommFun.Nvl(mrsNote!����)
                mrsNote.Update
            Else
                If lng���²�����ʾ��ʽ = 0 Or lng���²�����ʾ��ʽ = 2 Then
                    str���� = Replace(str����, "����", "�� ")
                End If
                Call Record_Add(mrsNote, gstrFields, str����)
                mrsNote!���� = IIf(bytTag = 1, 99, 6)
                mrsNote.Update
            End If
        Next i
    End If
    
    '���¶Ͽ���־(ʱ�䳬��һ������δ��˵����������)
    Call ProcessPoint(mstr��ʼʱ��)
    '������֯�ظ��ĵ�
    Call GetConverPoint(mrsPoint)
    
    '���δ��˵������ʾ����ȡ����¼��mrsNote������Ϊ99�ļ�¼
    If bytTag = 2 Then
        mrsNote.Filter = "����=99"
        Do While Not mrsNote.EOF
            mrsNote.Delete
            mrsNote.MoveNext
        Loop
        mrsNote.Filter = ""
    End If
    '�������������������Ϣ
    'Call OutputRsData(mrsPoint, True)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PrepareGraph()
    Dim strPic As String, strOverlap As String, strPart As String, lngID As Long    'ͼƬ,�ص����,��λ,����Ŀ���
    Dim lngCurX As Long, sinCurY As Single, lngCount As Long, lngMax As Long        'һ���ܱ�����ٸ�ͼƬ?
    Dim rsTemp As New ADODB.Recordset
    Dim rsOverlap As New ADODB.Recordset
    Dim ArrCode() As String, arrChar() As String, arrItem() As String
    Dim strChar As String
    Dim i As Integer
    On Error GoTo Errhand
    
   If Not mrsGraph Is Nothing Then If mrsGraph.State = 1 Then mrsGraph.Close
    
    lngMax = mobjBuffer.ScaleWidth \ gintBmpW      'һ���ܱ�����ٸ�ͼƬ?
    '�����������ͼ�����(���������ص����),ȫ����ȡ��picBuffer��,�˴��������Ŀ�Ĳ�λ�����Ӧ��ͼ�����
    gstrFields = "��Ŀ���," & adDouble & ",18|��λ," & adLongVarChar & ",50|��¼��," & adLongVarChar & ",50|" & _
                 "��¼ɫ," & adDouble & ",18|�ص���Ŀ," & adLongVarChar & ",20|��," & adDouble & ",5|��," & adDouble & ",5"    '�ص���ĿӦ����Ŀ��Ŵ�С����,��:1,4,5
    Call Record_Init(mrsGraph, gstrFields)
    
    '�ȸ������²�λװ��
    gstrSQL = " Select ��Ŀ���,'' AS ��λ, ��¼�� ��Ƿ���,��¼ɫ �����ɫ,1 չ�ַ�ʽ,'��' AS �ص���Ŀ From ���¼�¼��Ŀ Order by ��Ŀ���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��������Ŀ��չ�ַ�ʽ")
    With rsTemp
        Do While Not .EOF
            If !չ�ַ�ʽ = 1 Then
                
                If !��Ŀ��� = 1 Then '����
                    ArrCode = Split("����,Ҹ��,����", ",")
                    strChar = zlCommFun.Nvl(!��Ƿ���, "��,��,��")
                    strChar = strChar & String(2 - UBound(Split(strChar, ",")), ",")
                    arrChar = Split(strChar, ",")
                    For i = 0 To UBound(ArrCode)
                        gstrFields = "��Ŀ���|��λ|�ص���Ŀ|��¼��|��¼ɫ"
                        gstrValues = !��Ŀ��� & "|" & ArrCode(i) & "|" & zlCommFun.Nvl(!�ص���Ŀ) & "|" & arrChar(i) & "|" & zlCommFun.Nvl(!�����ɫ, 0)
                        Call Record_Add(mrsGraph, gstrFields, gstrValues)
                    Next i
                Else
                    strPart = ""
                    strChar = zlCommFun.Nvl(!��Ƿ���)
                    '������Ӧ���ڴ��¼����
                    gstrFields = "��Ŀ���|��λ|�ص���Ŀ|��¼��|��¼ɫ"
                    gstrValues = !��Ŀ��� & "|" & strPart & "|" & zlCommFun.Nvl(!�ص���Ŀ) & "|" & strChar & "|" & zlCommFun.Nvl(!�����ɫ, 0)
                    Call Record_Add(mrsGraph, gstrFields, gstrValues)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '��������ͺ�����ͼ��
    arrItem = Split("2,3", ",")
    ArrCode = Split("����,������", ",")
    arrChar = Split("PACEMAKER,BREATH", ",")
    For i = 0 To UBound(ArrCode)
        strPic = arrChar(i)  '��Դ�ļ�
        If strPic <> "" Then
            If DrawPicture(mobjBuffer, strPic, lngCurX, sinCurY, lngCurX + gintBmpW, sinCurY + gintBmpH, True) Then
                '������Ӧ���ڴ��¼����
                gstrFields = "��Ŀ���|��λ|�ص���Ŀ|��¼��|��|��"
                gstrValues = Val(arrItem(i)) & "|" & ArrCode(i) & "|" & "��" & "||" & sinCurY \ gintBmpH & "|" & lngCount
                Call Record_Add(mrsGraph, gstrFields, gstrValues)
                
                'λ�Ƽ���
                lngCurX = lngCurX + gintBmpW
                lngCount = lngCount + 1
                If lngCount >= lngMax Then
                    lngCount = 0
                    lngCurX = 0
                    sinCurY = sinCurY + gintBmpH
                End If
            End If
            'If !չ�ַ�ʽ = 2 Then Call FileSystem.Kill(strPic)
        End If
    Next i

    '�ٸ��������ص����װ��
    gstrSQL = " Select ���,��Ƿ���,�����ɫ From �����ص���� Where nvl(�ص���Ŀ,0)>0 Order by ���"
    Set rsOverlap = zldatabase.OpenSQLRecord(gstrSQL, "�ٸ��������ص����װ��")
    gstrSQL = " Select ���,�ϼ����,��Ŀ���,���²�λ From �����ص���� Where ��Ŀ��� is not null Order by ���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�ص�������Ŀ")
    Do While Not rsOverlap.EOF
        strPart = ""
        strOverlap = ""
        With rsTemp
            .Filter = "�ϼ����=" & rsOverlap!���
            If rsTemp.RecordCount > 0 Then
                '��һ����¼��Ϊ����¼
                lngID = zlCommFun.Nvl(!��Ŀ���, 0)
                strPart = rsOverlap!��� ' zlCommFun.Nvl(!���²�λ)
                 
                .Sort = "��Ŀ���"
                Do While Not .EOF
                    If !��Ŀ��� <> lngID Then
                        strOverlap = strOverlap & "," & !��Ŀ���
                    End If
                    .MoveNext
                Loop
                .Sort = "���"              '�˴���Ҫ����Ż�ԭ,����ȡ��һ���ص���Ŀ������Ŀʱ��ȡ��
                
                strOverlap = Mid(strOverlap, 2)
                If Not IsNull(rsOverlap!��Ƿ���) Then
                    '����ַ�
                    '������Ӧ���ڴ��¼����
                    gstrFields = "��Ŀ���|��λ|�ص���Ŀ|��¼��|��¼ɫ"
                    gstrValues = lngID & "|" & strPart & "|" & strOverlap & "|" & zlCommFun.Nvl(rsOverlap!��Ƿ���) & "|" & rsOverlap!�����ɫ
                    Call Record_Add(mrsGraph, gstrFields, gstrValues)
                Else
                    '���ͼ���ļ�
                    strPic = zlBlobRead(9, rsOverlap!���)
                    If strPic <> "" Then
                        If DrawPicture(mobjBuffer, strPic, lngCurX, sinCurY, lngCurX + gintBmpW, sinCurY + gintBmpH, False) Then
                            '������Ӧ���ڴ��¼����
                            gstrFields = "��Ŀ���|��λ|�ص���Ŀ|��¼��|��|��"
                            gstrValues = lngID & "|" & strPart & "|" & strOverlap & "||" & sinCurY \ gintBmpH & "|" & lngCount
                            Call Record_Add(mrsGraph, gstrFields, gstrValues)
                            
                            'λ�Ƽ���
                            lngCurX = lngCurX + gintBmpW
                            lngCount = lngCount + 1
                            If lngCount >= lngMax Then
                                lngCount = 0
                                lngCurX = 0
                                sinCurY = sinCurY + gintBmpH
                            End If
                        End If
                        Call FileSystem.Kill(strPic)
                    End If
                End If
            End If
        End With
        rsOverlap.MoveNext
    Loop
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ProcessPoint(ByVal strBeginDate As String)
    Dim arrData
    Dim lngOrder As Long
    Dim lngCurX As Long                         '��¼δ��˵����������
    Dim strPrimary As String
    Dim strDate As String
    On Error GoTo Errhand
    '�������е�ĸ��ֱ�־λ����Ͽ�
    
    '�ȴ���δ��˵������δ��˵��ǰһ����ĶϿ���־����Ϊ1
    '---------------------------------------------------
    strPrimary = "���|"        '��ʽ:�ֶ���,ֵ
    gstrFields = "�Ͽ�"         '��ʽ:�ֶ���|�ֶ���
    gstrValues = 1              '��ʽ:ֵ|ֵ
    
    mrsPoint.Filter = ""

    With mrsNote
        .Filter = "����=99"
        Do While Not .EOF
            lngCurX = GetXCoordinate(!ʱ��, strBeginDate)
            If mint����Ӧ�� = 2 And !��Ŀ��� = -1 Then
                mrsPoint.Filter = "��Ŀ���=" & gint���� & " And  X����<=" & !X����
            Else
                If Val(!��Ŀ���) = 1 Then
                    mrsPoint.Filter = "��Ŀ���=" & !��Ŀ��� & " And  ���<>1 And X����<" & !X����
                Else
                    mrsPoint.Filter = "��Ŀ���=" & !��Ŀ��� & " And X����<" & !X����
                End If
            End If
            
            mrsPoint.Sort = "ʱ��"
            If mrsPoint.RecordCount <> 0 Then
                mrsPoint.MoveLast
                lngOrder = mrsPoint!���
                
                Call Record_Update(mrsPoint, gstrFields, gstrValues, strPrimary & lngOrder)
            End If
            mrsPoint.Filter = 0
            
            .MoveNext
        Loop
        
        .Filter = 0
    End With
        
    '��һ��δ�����ݵ�Ҳ���öϿ���־
    '---------------------------------------------------
    lngCurX = 0
    lngOrder = 0
    strPrimary = ""
    mrsPoint.Filter = ""
    mrsPoint.Sort = "��Ŀ���,ʱ��,���"
    
    With mrsPoint
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Not (Val(!��Ŀ���) = gint���� And Val(zlCommFun.Nvl(!���)) = 1) Then
                If lngCurX <> 0 Then
                    If lngCurX <> !��Ŀ��� Then strDate = ""
                End If
                lngCurX = !��Ŀ���
                
                If strDate <> "" Then
                    If DateDiff("d", CDate(strDate), CDate(Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss"))) > 1 Then
                        strPrimary = strPrimary & "," & lngOrder
                    End If
                End If
                '��¼��ǰ��������Ϣ,����һ����ʱ���
                strDate = Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss")
                lngOrder = Val(zlCommFun.Nvl(!���))
            End If
            .MoveNext
        Loop
    End With
    
    arrData = Split(strPrimary, ",")
    lngOrder = UBound(arrData)
    strPrimary = "���|"        '��ʽ:�ֶ���,ֵ
    
    For lngCurX = 1 To lngOrder
        Call Record_Update(mrsPoint, gstrFields, gstrValues, strPrimary & arrData(lngCurX))
    Next
    
    '�������²�����.��ǰһ����ĶϿ���־����Ϊ1
    mrsPoint.Filter = ""
    mrsPoint.Filter = "��Ŀ���=" & gint���� & " and ���<>1"
    mrsPoint.Sort = "ʱ��,���"
    With mrsPoint
        Do While Not .EOF
            If !��ֵ = "����" And .AbsolutePosition <> 1 Then
                .MovePrevious '������һ�жϿ����
                If Val(zlCommFun.Nvl(!�Ͽ�)) <> 1 Then
                    lngOrder = !���
                    Call Record_Update(mrsPoint, gstrFields, gstrValues, strPrimary & lngOrder)
                End If
                .MoveNext
            End If
        .MoveNext
        Loop
    End With
    mrsPoint.Filter = 0
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawMarker(ByVal bln��ͼ���� As Boolean, ByVal lng��Ŀ��� As Long, ByVal str��λ As String, ByVal lngCurX As Long, ByVal sinCurY As Single, Optional ByVal str�ص���Ŀ As String = "��", Optional ByVal bln��Ŀ As Boolean = False, Optional ByVal str���� As String = "")

    'bln��ͼ����=True,���µ���ͼ����,����������ʾ;����,�մ����������ʾ
    Dim blnGraph As Boolean
    Dim bln�ص� As Boolean
    Dim str��¼�� As String
    Dim lngRGB As Long

    On Error GoTo Errhand

    '����ַ���ͼ��
  
    mrsGraph.Filter = "��Ŀ���=" & lng��Ŀ��� & " And ��λ='" & str��λ & "' And �ص���Ŀ='" & str�ص���Ŀ & "'"

    If mrsGraph.RecordCount = 0 Then    'δ�����ص���Ŀ�������ʽ,����Ŀ���+��λ���
        mrsGraph.Filter = "��Ŀ���=" & lng��Ŀ��� & " And ��λ='" & str��λ & "'"
    Else
        bln�ص� = True
    End If
    
    If mrsGraph.RecordCount = 0 Then    'δ���ø���Ŀ����λ�������ʽ,����Ŀ���������
        mrsGraph.Filter = "��Ŀ���=" & lng��Ŀ���
    End If
    
    If mrsGraph.RecordCount = 0 Then Exit Sub
    blnGraph = (zlCommFun.Nvl(mrsGraph!��¼��) = "")
    
    If Not blnGraph Then
        If bln�ص� = True And str�ص���Ŀ <> "��" Then
            str��¼�� = zlCommFun.Nvl(mrsGraph!��¼��)
        Else

            If str���� <> "" Then
                str��¼�� = str����
            Else
                str��¼�� = zlCommFun.Nvl(mrsGraph!��¼��)
            End If
        End If
        
        lngRGB = Val(mrsGraph!��¼ɫ)
        
        If lng��Ŀ��� = -1 And mint����Ӧ�� = 2 Then lngRGB = RGB_RED
        
        '�ַ����
        Call SetTextColor(mlngMemDC, lngRGB)
        Call GetTextRect(mobjDraw, lngCurX - IIf(bln��Ŀ = True, Screen.TwipsPerPixelY / 2, 0), sinCurY + IIf(bln��Ŀ = True, Screen.TwipsPerPixelY / 2, 0), Trim(Split(str��¼�� & ",", ",")(0)), IIf(bln��ͼ����, T_DrawClient.�е�λ, T_DrawClient.�̶ȵ�λ))
        Call DrawText(mlngMemDC, Trim(Split(str��¼�� & ",", ",")(0)), -1, T_LableRect, DT_CENTER)
    Else

        '���������Ŀ��ͼ��
        If bln��ͼ���� Then
            '������ͼ������д�ӡ
            Call BitBlt(mlngMemDC, lngCurX + 2, sinCurY - gintBmpH / 2, gintBmpW, gintBmpH, mobjBuffer.hDC, mrsGraph!�� * gintBmpW, mrsGraph!�� * gintBmpH, SRCCOPY)
        Else
            '�̶�����ָ���������
            Call BitBlt(mlngMemDC, lngCurX, sinCurY, gintBmpW, gintBmpH, mobjBuffer.hDC, mrsGraph!�� * gintBmpW, mrsGraph!�� * gintBmpH, SRCCOPY)
        End If
    End If
    
    mrsGraph.Filter = ""

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CurveCount() As Long
'--------------------------------------------------
'����:�õ�����������Ŀ����
'--------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    On Error GoTo Errhand
    
    strSql = " Select Count(*) ��¼" & _
             " From ���¼�¼��Ŀ A, �����¼��Ŀ C" & _
             " Where A.��Ŀ���=C.��Ŀ��� And A.��¼��=1" & _
             " And nvl(C.Ӧ�÷�ʽ,0)=1" & _
             " And nvl(C.���ò���,0) in (0,[1]) And  nvl(C.����ȼ�,3)>=[3] " & _
             " and (C.���ÿ���=1 OR (C.���ÿ���=2 and Exists (select 1 from �������ÿ��� D where C.��Ŀ���=D.��Ŀ��� and D.����ID=[2])))" & _
             " Order by C.��Ŀ���"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "ȡ��ʼ��", IIf(T_Patient.lngӤ�� = 0, 1, 2), T_Patient.lng����ID, T_Patient.lng����ȼ�)
    
    lngCount = Val(zlCommFun.Nvl(rsTemp!��¼))
    
    CurveCount = lngCount
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintState(ByVal intPrintRange As Integer, ByVal blnPrint As Boolean, Optional lngBeginY As Long, _
    Optional ByVal intPageNo As Integer = -1, Optional ByVal strPrintDevice As String, Optional strPage As String, Optional strParam As String = "") As Boolean
    '******************************************************************************************************************
    '����:����ǰ���±��ǰ��ʼ���������±��������ӡ���ϻ�Ԥ������
    '����:blnCurState = �Ƿ�Ϊֻ��ӡ��ǰ���±�,�����ӡ�ӵ�ǰ��ʼ���������±�
    '     blnPrint    = �Ƿ��������ӡ���Ϸ��������Ԥ��������
    '******************************************************************************************************************
    
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strNewSql As String
    Dim strPaper As String
    Dim strPrintName As String
    Dim blnYesPrinter As Boolean
    Dim intCOl As Integer
    Dim intBeginPage As Integer
    Dim intEndPage As Integer
    Dim byeReturn As Byte
    Dim strArrFromTo() As String
    Dim intBaby As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim lngIndex As Long, lngIndexEnd As Long
    Dim intCount As Integer
    Dim objPrint As Object
    Dim strMarkDate As String
    Dim arrParam() As String
    
    On Error GoTo ErrHandle
    
    If strParam <> "" Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) < 2 Then
            MsgBox "strParam������Ϊ��ʱ,���봫���ļ�ID;����ID;��ҳID��", vbInformation, gstrSysName
            Exit Function
        End If
        T_Patient.lng�ļ�ID = Val(arrParam(0))
        T_Patient.lng����ID = Val(arrParam(1))
        T_Patient.lng��ҳID = Val(arrParam(2))
        If UBound(arrParam) > 2 Then T_Patient.lng����ID = Val(arrParam(3))
        If UBound(arrParam) > 3 Then T_Patient.lngӤ�� = Val(arrParam(4))
    End If
    
    intBaby = T_Patient.lngӤ��
    
    '------------------------------------------------------------------------------------------------------------------
    '��ӡ���ָ�������
    If Not ExistsPrinter Then
        MsgBox "ϵͳû�а�װ�κδ�ӡ�����ܼ�����ӡ�������˳���", vbInformation, gstrSysName
        Exit Function
    End If
    
    gPrinter.lngLeft = OFFSET_LEFT
    gPrinter.lngRight = OFFSET_RIGHT
    gPrinter.lngTop = OFFSET_TOP
    gPrinter.lngBottom = OFFSET_BOTTOM
    '��ȡ��ӡ����
    strSql = "Select ��ʽ From ����ҳ���ʽ Where ���� = 3 And ��� In (Select A.ҳ�� From �����ļ��б� A,���˻����ļ� B Where A.Id = B.��ʽID and B.ID=[1])"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ�ļ���ӡ����", T_Patient.lng�ļ�ID)
    If Not rsTmp.EOF Then strPaper = "" & rsTmp!��ʽ
    
    If UBound(Split(strPaper, ";")) >= 0 Then
        gPrinter.intPage = Val(Split(strPaper, ";")(0))
        If UBound(Split(strPaper, ";")) >= 1 Then gPrinter.intOrient = Val(Split(strPaper, ";")(1))
        If UBound(Split(strPaper, ";")) >= 2 Then gPrinter.lngHeight = Val(Split(strPaper, ";")(2))
        If UBound(Split(strPaper, ";")) >= 3 Then gPrinter.lngWidth = Val(Split(strPaper, ";")(3))
        If UBound(Split(strPaper, ";")) >= 4 Then gPrinter.lngLeft = CLng(Val(Split(strPaper, ";")(4)) / conRatemmToTwip)
        If UBound(Split(strPaper, ";")) >= 5 Then gPrinter.lngRight = CLng(Val(Split(strPaper, ";")(5)) / conRatemmToTwip)
        If UBound(Split(strPaper, ";")) >= 6 Then gPrinter.lngTop = CLng(Val(Split(strPaper, ";")(6)) / conRatemmToTwip)
        If UBound(Split(strPaper, ";")) >= 7 Then gPrinter.lngBottom = CLng(Val(Split(strPaper, ";")(7)) / conRatemmToTwip)
    End If
    
    If strPrintDevice = "" Then
        If Trim(zldatabase.GetPara("���µ���ӡ��", glngSys, 1255, "")) = "" Then
            MsgBox "û�����ô�ӡ��,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
            strPrintName = Printer.DeviceName
        Else
            strPrintName = Trim(zldatabase.GetPara("���µ���ӡ��", glngSys, 1255, Printer.DeviceName))
        End If
    Else
        strPrintName = strPrintDevice
    End If
    
    '��ӡ��
    blnYesPrinter = False
    If Printer.DeviceName <> strPrintName Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrintName Then Set Printer = Printers(i): blnYesPrinter = True: Exit For
        Next
        If blnYesPrinter = False Then
            MsgBox "���õĴ�ӡ���Ѳ�����,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
        End If
    End If
    'ȱʡʹ�ô�ӡ��Ĭ�Ͻ�ֹ���˴���������(ֻҪ�����˽�ֹ��ʽ��ӡ�Ͳ�������
    gPrinter.intBin = Val(zldatabase.GetPara("���µ���ֽ", glngSys, 1255, Printer.PaperBin))
    
    On Error Resume Next
    'ֽ��
    If gPrinter.intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = gPrinter.lngWidth
        Printer.Height = gPrinter.lngHeight
    Else
        Printer.PaperSize = gPrinter.intPage
    End If
    
    Printer.Orientation = gPrinter.intOrient
    If IsWindowsNT And gPrinter.intPage = 256 Then
        Call SetNTPrinterPaper(frmFlash.hWnd, gPrinter.lngWidth / conRatemmToTwip, gPrinter.lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
        Unload frmFlash
    End If
    
    On Error GoTo ErrHandle
    
    '------------------------------------------------------------------------------------------------------------------
    lngBeginY = IIf(gPrinter.lngTop > lngBeginY, gPrinter.lngTop, lngBeginY)
    lngIndex = mintPage
    
    '���ֻ��ӡ��ǰ��ֻ����ʼ�ͽ���дͬһҳ��
    Set mfrmCaseTendBodyPrint = New frmCaseTendBodyPrint
    Load frmTendFileRead
    Call frmTendFileRead.InitRechBox(T_Patient.lng�ļ�ID)
    strMarkDate = ""
    '��ȡ�û����õ����µ���ʼʱ��(Ӥ��������Ӥ������ʱ��Ϊ׼)
    strSql = "select ��ʼʱ�� from ���˻����ļ� where ID=[1] and ����ID=[2] and ��ҳid=[3] and nvl(Ӥ��,0)=[4]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ���µ���ʼʱ��", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
    If rsTmp.RecordCount <> 0 Then
        strMarkDate = Format(rsTmp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If strMarkDate <> "" Then strMarkDate = "to_date('" & strMarkDate & "','yyyy-MM-dd hh24:mi:ss')"
    
     '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ)����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "   (SELECT /*+ RULE */  ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0 AND C.��� = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.�������� = COLUMN_VALUE) And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
                
    '��ȡ�˲��˵����µ���ҳ��
    '------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��) AS ��Ժʱ��," & vbNewLine & _
                "    DECODE(E.��¼,0,DECODE(SIGN(NVL(E.Ӥ��ʱ��,B.��Ժʱ��) - D.����ʱ��), 1,NVL(E.Ӥ��ʱ��,B.��Ժʱ��) ,D.����ʱ��),NVL(E.Ӥ��ʱ��,B.��Ժʱ��))  ��Ժʱ��," & vbNewLine & _
                "    1 + TRUNC((TO_DATE(TO_CHAR(DECODE(E.��¼,0,DECODE(SIGN(NVL(E.Ӥ��ʱ��,B.��Ժʱ��) - D.����ʱ��), 1,NVL(E.Ӥ��ʱ��,B.��Ժʱ��) ,D.����ʱ��),NVL(E.Ӥ��ʱ��,B.��Ժʱ��)),'yyyy-MM-dd'),'yyyy-MM-dd') - " & vbNewLine & _
                "    TO_DATE(TO_CHAR(DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��),'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS ҳ��,D.����ʱ��" & vbNewLine & _
                "    FROM (SELECT ����ID,��ҳID,MIN(��ʼʱ��) AS ��Ժʱ��," & vbNewLine & _
                "    MAX(NVL(��ֹʱ��, SYSDATE)) AS ��Ժʱ��" & vbNewLine & _
                "    FROM ���˱䶯��¼" & vbNewLine & _
                "    WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID = [3] GROUP BY ����ID,��ҳID) B," & vbNewLine & _
                "    (SELECT ����ID,��ҳID,����ʱ�� FROM ������������¼ WHERE ����ID =[2] AND ��ҳID =[3] AND ���=[4]) C ," & vbNewLine & _
                "    (SELECT NVL(����ʱ��,SYSDATE) ����ʱ�� FROM ( SELECT MAX(����ʱ��) ����ʱ�� FROM ���˻����ļ� A,���˻������� B" & vbNewLine & _
                "    WHERE A.ID=B.�ļ�ID AND A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4])) D," & vbNewLine & _
                strNewSql & vbNewLine & _
                "WHERE B.����ID=E.����ID And B.��ҳID=E.��ҳID And B.����ID=C.����ID(+) AND B.��ҳID=C.��ҳID(+)"

    Set rsTmp = zldatabase.OpenSQLRecord(strSql, mstrTitle, T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
    intCount = 0
    For intCOl = 0 To rsTmp("ҳ��").Value - 1
    
        strDateFrom = Format(rsTmp("��Ժʱ��").Value + 7 * intCOl, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + 7 * (intCOl + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
        End If
        
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
        
            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            
            ReDim Preserve strArrFromTo(intCount)
            strArrFromTo(intCount) = "0;" & intCOl + 1 & ";" & intCOl + 1
            intCount = intCount + 1
        End If
    Next
    
    If blnPrint = True Then
        Set objPrint = Printer
    Else
        Set objPrint = mfrmCaseTendBodyPrint
    End If
    
    Select Case intPrintRange
    Case 0                  '��ӡ��ǰҳ
        If InStr(1, strPage, ";") <> 0 Then
            lngIndex = Val(Split(strPage, ";")(1))
        End If
        strPage = lngIndex & ";" & lngIndex
        
        If blnPrint = True Then Printer.Print ""
        If PrintOrPreviewBodyState(objPrint, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng�ļ�ID, intBaby, _
                T_Patient.lng����ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, False, _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, , mblnMoved) = True Then
                
                If blnPrint = False Then
                    mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, T_Patient.lng����ID, T_Patient.lng��ҳID, _
                        T_Patient.lng�ļ�ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                        CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, T_Patient.lng����ID, T_Patient.lngӤ��
                Else
                    'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                    Printer.EndDoc
                End If
        Else
            MsgBox "δ֪����������µ�ʧ�ܣ�", vbExclamation, gstrSysName
        End If
        
    Case 1              '�ӵ�ǰҳ������ӡ
        If InStr(1, strPage, ";") <> 0 Then
            lngIndex = Val(Split(strPage, ";")(0))
            lngIndexEnd = Val(Split(strPage, ";")(1))
            If lngIndexEnd > UBound(strArrFromTo) Then lngIndexEnd = UBound(strArrFromTo)
            
        Else
            lngIndexEnd = UBound(strArrFromTo)
        End If
        
        strPage = lngIndex & ";" & lngIndexEnd
        
        For intCOl = lngIndex To lngIndexEnd
            If blnPrint = True Then Printer.Print ""
            If PrintOrPreviewBodyState(objPrint, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng�ļ�ID, intBaby, _
                T_Patient.lng����ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> lngIndex, _
                CInt(Split(strArrFromTo(intCOl), ";")(1)), CInt(Split(strArrFromTo(intCOl), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "δ֪���󣬴�ӡʧ�ܣ�", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCOl = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, T_Patient.lng����ID, T_Patient.lng��ҳID, _
            T_Patient.lng�ļ�ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, T_Patient.lng����ID, T_Patient.lngӤ��
        Else '������ӡ�Ǽ�¼��ӡ�Ŀ�ʼҳ�źͽ���ҳ��
            strSql = "zl_���µ�����_Printer(" & T_Patient.lng�ļ�ID & "," & lngIndex + 1 & "," & lngIndexEnd + 1 & ")"
            Call zldatabase.ExecuteProcedure(strSql, "zl_���µ�����_Printer")
        End If
        
    Case 2          '�ӵ�һҳ������ӡ,��ȫ����ӡ
        strPage = 0
        For intCOl = 0 To UBound(strArrFromTo)
            If blnPrint = True Then Printer.Print ""
            If PrintOrPreviewBodyState(objPrint, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng�ļ�ID, intBaby, _
                T_Patient.lng����ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> 0, _
                CInt(Split(strArrFromTo(intCOl), ";")(1)), CInt(Split(strArrFromTo(intCOl), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "δ֪���󣬴�ӡʧ�ܣ�", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCOl = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, T_Patient.lng����ID, T_Patient.lng��ҳID, _
            T_Patient.lng�ļ�ID, CInt(Split(strArrFromTo(0), ";")(1)), _
                CInt(Split(strArrFromTo(0), ";")(1)), intPageNo, strArrFromTo, strPage, T_Patient.lng����ID, T_Patient.lngӤ��
        End If
    End Select
    
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And gPrinter.intPage = 256 Then DelCustomPaper
    
    Unload frmTendFileRead
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'��PictureBoxģ���3Dƽ�水ť
'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
Private Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hDC, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hDC, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Private Sub vsf_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call VsfDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done)
End Sub

Private Sub VsfDrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'----------------------------------
'����:��ɵ�Ԫ���ͼ
'-----------------------------------
    Dim T_ClientRect As RECT
    Dim strText As String, strTmp As String, strPart
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngBackColor As Long, lngForeColor As Long
    Dim i As Integer
    Dim int�� As Integer, int�� As Integer
    
     
    If mbln�������� Or mItemNO.���� = 0 Then Exit Sub
    If vsf.Body.RowHidden(Row) Or vsf.Body.ColHidden(Col) Then Exit Sub
    If Col < vsf.FixedCols Then Exit Sub
             
    On Error GoTo Errhand
    '�趨�ͻ������С
    With T_ClientRect
        .Left = Left + 1
        .Top = Top + 1
        .Right = Right - 1
        .Bottom = Bottom - 1
    End With
    
    'ֻ��ͼ��
    If Val(vsf.ColData(Col)) <> 0 Then
        '��ȡͼ��λ��
        mrsGraph.Filter = "��Ŀ���=" & gint���� & " And ��λ='������'"
        If mrsGraph.RecordCount = 0 Then
            int�� = -1: int�� = -1
        Else
            int�� = Val(mrsGraph!��)
            int�� = Val(mrsGraph!��)
        End If
        
        '1���������
        '�����뱳��ɫ��ͬ��ˢ��
        lngBackColor = vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col)
        If lngBackColor = 0 Then lngBackColor = vsf.Body.BackColor
        lngBackColor = GetRBGFromOLEColor(lngBackColor)
        lngForeColor = 200
        lngBrush = CreateSolidBrush(lngBackColor)
        'ʹ�ø�ˢ����䱳��ɫ
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, T_ClientRect, lngBrush)
        '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        T_ClientRect.Left = Left + (T_ClientRect.Right - Left - gintBmpW) / 2
        If Val(vsf.ColData(Col)) = 2 Then
            T_ClientRect.Top = Top + (T_ClientRect.Bottom - gintBmpH)
        End If
        '��ʼ����ͼ��
        Call BitBlt(hDC, T_ClientRect.Left, T_ClientRect.Top, gintBmpW, gintBmpH, mobjBuffer.hDC, int�� * gintBmpW, int�� * gintBmpH, SRCCOPY)
        mrsGraph.Filter = 0
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawDownTabAnsyGrade(ByVal lngDC As Long, ByVal objDraw As Object, arrText() As String, ByVal Row As Long, ByVal Col As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean, Optional ByVal blnFormat As Boolean = False)
'---------------------------------------------------
'���� ���������
'˵�� AnsyGrade=True���ܵ��ô˺���
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer, intOldSize As Integer
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngBackColor As Long, lngForeColor As Long
    Dim stdset As StdFont, stdOldset As StdFont
    Dim LPoint As T_LPoint, T_ClientRect As RECT
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    
    On Error GoTo Errhand
    
    If UBound(arrText) < 2 Then Exit Sub
    
     '�趨�ͻ������С
    With T_ClientRect
        .Left = Left + 1
        .Top = Top + 1
        .Right = Right - 1
        .Bottom = Bottom - 1
        LPoint.W = .Right - .Left
        LPoint.X = .Left
        LPoint.Y = .Top + (.Bottom - .Top) / 2
    End With
    
    '1���������
    '�����뱳��ɫ��ͬ��ˢ��
    lngBackColor = mshDownTab.Cell(flexcpBackColor, Row, Col, Row, Col)
    If lngBackColor = 0 Then lngBackColor = objDraw.BackColor
    lngBackColor = GetRBGFromOLEColor(lngBackColor)
    lngForeColor = GetRBGFromOLEColor(mshDownTab.Cell(flexcpForeColor, Row, Col, Row, Col))
    lngBrush = CreateSolidBrush(lngBackColor)
    'ʹ�ø�ˢ����䱳��ɫ
    lngOldBrush = SelectObject(lngDC, lngBrush)
    Call FillRect(lngDC, T_ClientRect, lngBrush)
    '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
    Call SelectObject(lngDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        If Len(str2) > Len(str3) Then
            strTmp = str1 & str2
        Else
            strTmp = str1 & str3
        End If
    Else
        strTmp = str1 & str2 & "/" & str3
    End If
    intSize = objDraw.Font.Size
    intOldSize = intSize
    objDraw.Font.Size = intSize
    Set stdset = New StdFont
    stdset.Name = "����"
    stdset.Size = intSize
    stdset.Bold = False
    Set stdOldset = stdset 'ԭʼ����
    
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , 1)
    '������
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngForeColor)
        Call DrawText(lngDC, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + 1
    Else
        lngX = T_LableRect.Left
    End If

    If blnFormat = True Then '���ӷ�ĸ��ʾ
        intSize = 7
        objDraw.Font.Size = intSize
        Set stdset = New StdFont
        stdset.Name = "����"
        stdset.Size = intSize
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngForeColor)
        T_LableRect.Left = lngX
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        'If T_LableRect.Top < Top Then T_LableRect.Top = Top - 1
        T_LableRect.Bottom = T_ClientRect.Bottom
        Call DrawText(lngDC, str2, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        '������
        objDraw.Font.Size = intOldSize
        Call DrawLine(lngDC, lngX, lngY, lngX + (objDraw.TextWidth("A") / T_TwipsPerPixel.X), lngY)
        '�����ĸ
        lngY = lngY
        T_LableRect.Left = lngX
        T_LableRect.Top = lngY
        intSize = 7.5
        Set stdset = New StdFont
        stdset.Name = "����"
        stdset.Size = intSize
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngForeColor)
        Call DrawText(lngDC, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
    Else
        If str1 <> "" Then
            '����ϱ�
            intSize = 7
            objDraw.Font.Size = intSize
            Set stdset = New StdFont
            stdset.Name = "����"
            stdset.Size = intSize
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngForeColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < T_ClientRect.Top Then T_LableRect.Top = T_ClientRect.Top - 1
            Call DrawText(lngDC, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            '�����벿��
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngForeColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDC, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
        Else
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngForeColor)
            Call DrawText(lngDC, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
        End If
    End If
    
    objDraw.Font.Size = intOldSize
    Set stdset = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

