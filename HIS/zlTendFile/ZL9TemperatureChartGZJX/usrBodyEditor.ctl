VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl usrBodyEditor 
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8205
   ScaleWidth      =   10800
   Begin VB.PictureBox picSerach 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1515
      TabIndex        =   26
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
         TabIndex        =   33
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   10215
      Begin VB.TextBox txtLength 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5505
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   6165
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   9600
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   9375
         Begin VB.PictureBox picCommText 
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
            Height          =   300
            Left            =   75
            ScaleHeight     =   300
            ScaleWidth      =   7260
            TabIndex        =   36
            Top             =   4860
            Width           =   7260
         End
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
            TabIndex        =   34
            Top             =   2160
            Width           =   7335
         End
         Begin VSFlex8Ctl.VSFlexGrid mshUpTab 
            Height          =   1095
            Left            =   120
            TabIndex        =   24
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
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   60
               Width           =   165
               Begin VB.Image imgDisPlay 
                  Appearance      =   0  'Flat
                  Height          =   240
                  Left            =   -30
                  Picture         =   "usrBodyEditor.ctx":076A
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
               TabIndex        =   28
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid mshDownTab 
            Height          =   975
            Left            =   90
            TabIndex        =   25
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
            FormatString    =   $"usrBodyEditor.ctx":6FBC
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
         Begin zl9TemperatureChartGZJX.VsfGrid vsf 
            Height          =   255
            Left            =   120
            TabIndex        =   27
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
            TabIndex        =   7
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
               TabIndex        =   15
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
               Index           =   4
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   12
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
               Index           =   6
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   14
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
               TabIndex        =   13
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
               Index           =   0
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   8
               TabStop         =   0   'False
               Text            =   "������"
               Top             =   60
               Width           =   1425
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   2
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   10
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
               Index           =   3
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   11
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
               Index           =   1
               Left            =   6645
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "1234567"
               Top             =   60
               Width           =   3825
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��    ��:"
               Height          =   180
               Index           =   7
               Left            =   4065
               TabIndex        =   23
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   3
               Left            =   2910
               TabIndex        =   19
               Top             =   390
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   6
               Left            =   2910
               TabIndex        =   22
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
               TabIndex        =   21
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   16
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   18
               Top             =   375
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����:"
               Height          =   180
               Index           =   5
               Left            =   4050
               TabIndex        =   20
               Top             =   60
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ��:"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   17
               Top             =   60
               Width           =   630
            End
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
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "��ʱ��ͼ��,ǧ���ɾ"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2115
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
         ScaleWidth      =   5220
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   6120
         Width           =   5220
         Begin VB.ComboBox cboFile 
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
            Left            =   3165
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   2085
         End
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
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   0
            Width           =   1920
         End
         Begin VB.Label lblSerach 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ļ�"
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
            Index           =   0
            Left            =   2730
            TabIndex        =   3
            Top             =   60
            Width           =   360
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
            Left            =   225
            TabIndex        =   1
            Top             =   60
            Width           =   360
         End
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
Private mcbrItem    As CommandBarControl

'--����
Public mblnResize As Boolean '��¼�����С�Ƿ����仯
Public mblnMoved As Boolean
Private mlngWidth As Long
Private mlngHeight As Long
Private mintPage      As Integer '��¼��ǰҳ��
Private mintAllPage As Integer '���µ�����ҳ��
Private mintIndex As Integer '��¼ҳ����ת������
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
Private mstrOpdays() As String
Private mstrOpValue() As String
Private mstrNewString() As String '����Ƥ�Խ����Ϣ
Private mlngNewHeight() As Long '����Ƥ�Խ���и�

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
    lng��ʽID As Long
End Type
Private T_Patient As type_Patient

'--�¼�����
Public Event CmdClick(ByVal strParam As String)
Public Event zlAfterPrint()
Public Event DbClickCur(ByVal intDataEditor As Integer)
Public Event zlFileChange(ByVal blnRefresh As Boolean, ByVal lngFileID As Long, ByVal lngBaby As Long)
Public Event zlDataChange(ByVal blnChange As Boolean)
Public Event ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)

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
    Dim strSQL        As String, strNewSql As String
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
    Dim strMarkDate As String, strFileBeginTime As String, strFileEndTime As String '���µ�����ʱ��
    Dim intCOl        As Integer
    Dim strCaption    As String, strCategory As String, strUnitName As String, blnAddMenu As Boolean
    Dim strParameter  As String
    Dim strSvrCaption As String, strSvrCaption1 As String
    Dim strNow        As String
    Dim strCut        As String
    Dim lngLoop       As Long
    Dim strTmp        As String
    Dim lnglast����id As Long
    Dim lng���� As Long
    
    On Error GoTo Errhand

    If lng����ID = 0 And lng�ļ�ID = 0 And lng��ҳID = 0 Then Exit Function
    mbln��Ժ = False
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
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
    strSQL = "Select ��ʼʱ��,����ʱ�� From ���˻����ļ� where ID=[1] and ����ID=[2] and ��ҳid=[3] and nvl(Ӥ��,0)=[4]"
    If mblnMoved = True Then strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���µ���ʼʱ��", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��)
    If rsTmp.RecordCount <> 0 Then
        strEnterDate = Format(rsTmp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
        strFileEndTime = Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:mm:ss")
    End If
    strFileBeginTime = strEnterDate
    strMarkDate = "To_date('" & strEnterDate & "','yyyy-MM-dd hh24:mi:ss')"
    '------------------------------------------------------------------------------------------------------------------
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ)����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "   (SELECT ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0 AND B.������� = 'Z'" & vbNewLine & _
                "                And Instr(',3,5,11,', ',' || c.�������� || ',') > 0 And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    '˵��:Ŀǰ����ר�����µ������˿���ͬʱ���ڶ�����µ������µ���ʼʱ�����ֹʱ��Ĺ�������:
    '����ļ��Ŀ�ʼʱ�䲻Ϊ�ղ��Ҵ��ڵ��ڲ�����Ժʱ���Ӥ������ʱ��,���µ��Ŀ�ʼʱ�����ļ���ʼʱ��Ϊ׼,�����Բ�����Ժʱ���Ӥ������ʱ��Ϊ׼
    '����ļ�����ֹʱ�䲻Ϊ�ղ���С�ڵ��ڲ��˻�Ӥ����Ժʱ�䣨δ��Ժ���ܲ��ܴ��ڵ�ǰʱ�䣩,���µ�����ʱ�����ļ���ʼʱ��Ϊ׼���������µ�����ʱ���Բ��˻�Ӥ����Ժʱ��Ϊ׼(δ��ԺΪ��ǰʱ��)
    '����ļ�����ֹʱ��Ϊ��,����ԭ�з�ʽ,��������Ѿ���Ժ�����ѳ�Ժʱ��Ϊ׼,δ��Ժ���ѵ�ǰʱ������ݽ���ʱ��Ϊ׼.
    '��ȡ�˲��˵����µ���ҳ��
    '------------------------------------------------------------------------------------------------------------------
    strSQL = " SELECT  ��Ժʱ��,ʵ����Ժʱ��,��Ժʱ��,1 + TRUNC((TO_DATE(TO_CHAR(��Ժʱ��,'yyyy-MM-dd'),'yyyy-MM-dd') -TO_DATE(TO_CHAR(��Ժʱ��,'yyyy-MM-dd'),'yyyy-MM-dd')) / " & T_BodyStyle.lng���� & ") AS ҳ��,����ʱ��,��¼ " & _
             "  From (" & _
             "      SELECT DECODE(D.��ʼʱ��,NULL,DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��)," & vbNewLine & _
             "                 DECODE(SIGN(D.��ʼʱ�� - DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��))," & vbNewLine & _
             "                        1," & vbNewLine & _
             "                        D.��ʼʱ��," & vbNewLine & _
             "                        DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��))) AS ��Ժʱ��," & vbNewLine & _
             "      DECODE(C.����ʱ��,NULL,B.��Ժʱ��,C.����ʱ��) AS ʵ����Ժʱ��," & vbNewLine & _
             "      DECODE(D.����ʱ��,NULL," & vbNewLine & _
             "                 DECODE(E.��¼,0," & vbNewLine & _
             "                        DECODE(SIGN(NVL(E.Ӥ��ʱ��, B.��Ժʱ��) - D.����ʱ��), 1, NVL(E.Ӥ��ʱ��, B.��Ժʱ��), D.����ʱ��)," & vbNewLine & _
             "                        NVL(E.Ӥ��ʱ��, B.��Ժʱ��))," & vbNewLine & _
             "                 DECODE(SIGN(NVL(E.Ӥ��ʱ��, B.��Ժʱ��) - D.����ʱ��), 1, D.����ʱ��, NVL(E.Ӥ��ʱ��, B.��Ժʱ��))) ��Ժʱ��," & vbNewLine & _
             "      D.����ʱ��,DECODE(D.����ʱ��, NULL, E.��¼, 1) ��¼" & vbNewLine & _
             "      FROM (SELECT ����ID,��ҳID,MIN(��ʼʱ��) AS ��Ժʱ��," & vbNewLine & _
             "      MAX(NVL(��ֹʱ��, SYSDATE)) AS ��Ժʱ��" & vbNewLine & _
             "      FROM ���˱䶯��¼" & vbNewLine & _
             "      WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID = [3] GROUP BY ����ID,��ҳID) B," & vbNewLine & _
             "      (SELECT ����ID,��ҳID,����ʱ�� FROM ������������¼ WHERE ����ID =[2] AND ��ҳID =[3] AND ���=[4]) C ," & vbNewLine & _
             "      (SELECT NVL(����ʱ��, SYSDATE) ����ʱ��, ��ʼʱ��, ����ʱ��" & vbNewLine & _
             "         FROM (SELECT MAX(B.����ʱ��) ����ʱ��, MAX(A.��ʼʱ��) ��ʼʱ��, MAX(A.����ʱ��) ����ʱ��" & vbNewLine & _
             "                FROM ���˻����ļ� A, ���˻������� B" & vbNewLine & _
             "                WHERE A.ID = B.�ļ�ID(+) AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND A.Ӥ�� = [4])) D," & vbNewLine & _
             "  " & strNewSql & vbNewLine & _
             "   WHERE B.����ID=E.����ID And B.��ҳID=E.��ҳID And B.����ID=C.����ID(+) AND B.��ҳID=C.��ҳID(+))"
                
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��)
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
    strSQL = " SELECT /*+ RULE */ 1 + TRUNC((TO_DATE(TO_CHAR(DECODE(D.��ʼʱ��,NULL,A.��ʼʱ��, DECODE(SIGN(D.��ʼʱ�� - A.��ʼʱ��), 1, D.��ʼʱ��, A.��ʼʱ��)),'YYYY-MM-DD'),'YYYY-MM-DD') -" & vbNewLine & _
            "                  TO_DATE(TO_CHAR(DECODE(D.��ʼʱ��, NULL, B.��Ժʱ��, D.��ʼʱ��), 'YYYY-MM-DD'), 'YYYY-MM-DD')) / " & T_BodyStyle.lng���� & ") AS ��ʼҳ��," & vbNewLine & _
            "       1 + TRUNC((TO_DATE(TO_CHAR(DECODE(A.���,F.LAST,DECODE(D.����ʱ��,NULL," & vbNewLine & _
            "                                                 DECODE(E.��¼,0,DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹʱ��) - D.����ʱ��),1," & vbNewLine & _
            "                                                        NVL(E.Ӥ��ʱ��, A.��ֹʱ��),D.����ʱ��),NVL(E.Ӥ��ʱ��, A.��ֹʱ��))," & vbNewLine & _
            "                                                 DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹʱ��) - D.����ʱ��),1,D.����ʱ��,NVL(E.Ӥ��ʱ��, A.��ֹʱ��)))," & vbNewLine & _
            "                                          DECODE(D.����ʱ��, NULL,NVL(E.Ӥ��ʱ��, A.��ֹʱ��)," & vbNewLine & _
            "                                                 DECODE(SIGN(D.����ʱ�� - NVL(E.Ӥ��ʱ��, A.��ֹʱ��)),1,NVL(E.Ӥ��ʱ��, A.��ֹʱ��),D.����ʱ��)))," & vbNewLine & _
            "                                   'YYYY-MM-DD')," & vbNewLine & _
            "                           'YYYY-MM-DD') - TO_DATE(TO_CHAR(DECODE(D.��ʼʱ��, NULL, B.��Ժʱ��, D.��ʼʱ��), 'YYYY-MM-DD'), 'YYYY-MM-DD')) / " & T_BodyStyle.lng���� & ") AS ����ҳ��," & vbNewLine & _
            "                          D.����ʱ��, ����ID, C.����, DECODE(D.��ʼʱ��,NULL,A.��ʼʱ��,DECODE(SIGN(D.��ʼʱ�� - A.��ʼʱ��), 1, D.��ʼʱ��, A.��ʼʱ��)) ��ʼʱ��," & vbNewLine & _
            "      DECODE(A.���,F.LAST,DECODE(D.����ʱ��,NULL," & vbNewLine & _
            "                           DECODE(E.��¼,0,DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹʱ��) - D.����ʱ��),1," & vbNewLine & _
            "                                  NVL(E.Ӥ��ʱ��, A.��ֹʱ��),D.����ʱ��),NVL(E.Ӥ��ʱ��, A.��ֹʱ��))," & vbNewLine & _
            "                           DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹʱ��) - D.����ʱ��),1,D.����ʱ��,NVL(E.Ӥ��ʱ��, A.��ֹʱ��)))," & vbNewLine & _
            "                    DECODE(D.����ʱ��, NULL,NVL(E.Ӥ��ʱ��, A.��ֹʱ��)," & vbNewLine & _
            "                           DECODE(SIGN(D.����ʱ�� - NVL(E.Ӥ��ʱ��, A.��ֹʱ��)),1,NVL(E.Ӥ��ʱ��, A.��ֹʱ��),D.����ʱ��))) ��ֹʱ��"
    strSQL = strSQL & _
            " FROM (SELECT ROWNUM ���, ����ID, ��ʼʱ��, ��ֹʱ��" & vbNewLine & _
            "       FROM (SELECT ����ID, MIN(��ʼʱ��) AS ��ʼʱ��, MAX(NVL(��ֹʱ��, SYSDATE)) AS ��ֹʱ��" & vbNewLine & _
            "              FROM ���˱䶯��¼" & vbNewLine & _
            "              WHERE ((��ʼʱ��>=[5]" & IIf(IsDate(strFileEndTime), " And ��ʼʱ��<[6]", " ") & ") OR (��ʼʱ��<=[5] And (��ֹʱ�� IS NULL OR ��ֹʱ��>[5]))) AND ����ID = [2] AND ��ҳID = [3]" & vbNewLine & _
            "              GROUP BY ����ID" & vbNewLine & _
            "              ORDER BY ��ʼʱ��,��ֹʱ��)) A," & vbNewLine & _
            "     (SELECT DECODE(Y.����ʱ��, NULL, X.��Ժʱ��, Y.����ʱ��) AS ��Ժʱ��, X.����ID, X.��ҳID" & vbNewLine & _
            "       FROM (SELECT ����ID, ��ҳID, MIN(��ʼʱ��) AS ��Ժʱ��" & vbNewLine & _
            "              FROM ���˱䶯��¼" & vbNewLine & _
            "              WHERE ((��ʼʱ��>=[5]" & IIf(IsDate(strFileEndTime), " And ��ʼʱ��<[6]", " ") & ") OR (��ʼʱ��<=[5] And (��ֹʱ�� IS NULL OR ��ֹʱ��>[5]))) AND ����ID = [2] AND ��ҳID = [3]" & vbNewLine & _
            "              GROUP BY ����ID,��ҳID) X," & vbNewLine & _
            "            (SELECT ����ID, ��ҳID, ����ʱ�� FROM ������������¼ WHERE ����ID = [2] AND ��ҳID = [3] AND ��� = [4]) Y" & vbNewLine & _
            "       WHERE X.����ID = Y.����ID(+) AND X.��ҳID = Y.��ҳID(+)) B, ���ű� C," & vbNewLine & _
            "     (SELECT NVL(����ʱ��, SYSDATE) ����ʱ��, ��ʼʱ��, ����ʱ��" & vbNewLine & _
            "       FROM (SELECT MAX(����ʱ��) ����ʱ��, MAX(A.��ʼʱ��) ��ʼʱ��, MAX(A.����ʱ��) ����ʱ��" & vbNewLine & _
            "              FROM ���˻����ļ� A, ���˻������� B" & vbNewLine & _
            "              WHERE A.ID = B.�ļ�ID(+) AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3]  AND A.Ӥ�� = [4])) D," & vbNewLine & strNewSql & "," & vbNewLine & _
            "     (SELECT COUNT(*) LAST" & vbNewLine & _
            "       FROM (SELECT ����ID" & vbNewLine & _
            "              FROM ���˱䶯��¼" & vbNewLine & _
            "              WHERE ((��ʼʱ��>=[5]" & IIf(IsDate(strFileEndTime), " And ��ʼʱ��<[6]", " ") & ") OR (��ʼʱ��<=[5] And (��ֹʱ�� IS NULL OR ��ֹʱ��>[5]))) AND ����ID = [2] AND ��ҳID = [3]" & vbNewLine & _
            "              GROUP BY ����ID)) F" & vbNewLine & _
            " WHERE B.����ID = E.����ID AND B.��ҳID = E.��ҳID AND C.ID(+) = A.����ID" & vbNewLine & _
            " ORDER BY A.��ʼʱ��"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    If (IsDate(strFileEndTime)) Then
        Set RS = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��, CDate(strFileBeginTime), CDate(strFileEndTime))
    Else
        Set RS = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��, CDate(strFileBeginTime))
    End If
    
    For lngLoop = 0 To rsTmp("ҳ��").Value - 1

        strDateFrom = Format(rsTmp("��Ժʱ��").Value + T_BodyStyle.lng���� * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + T_BodyStyle.lng���� * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"

        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then

            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            
            RS.Filter = ""
            RS.Filter = "��ʼҳ��<=" & lngLoop + 1 & " And ����ҳ��>=" & lngLoop + 1
            RS.Sort = "��ʼʱ��,��ֹʱ��"
            blnAddMenu = (RS.RecordCount = 1)
            strCategory = ""
ToStart:
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
                
                If blnAddMenu = False Then
                    If intCOl = 1 Then
                        strCategory = Format(strTmp, "yyyy-MM-dd")
                        strUnitName = Nvl(RS("����").Value)
                    ElseIf intCOl = RS.RecordCount Then
                        blnAddMenu = True
                        strCategory = strCategory & "��" & Format(strCaption, "yyyy-MM-dd")
                        If strUnitName <> Nvl(RS("����").Value) Then
                            strUnitName = strUnitName & "->" & Nvl(RS("����").Value)
                        End If
                        strCategory = "��" & lngLoop + 1 & "ҳ��" & strCategory & "(" & strUnitName & ")"
                    End If
                    RS.MoveNext
                    If blnAddMenu = True Then GoTo ToStart
                Else
                    strCaption = Format(strTmp, "yyyy-MM-dd") & "��" & Format(strCaption, "yyyy-MM-dd")
                    strCaption = "��" & lngLoop + 1 & "ҳ��" & strCaption & "(" & RS("����").Value & ")"
                    If strCategory = "" Then strCategory = strCaption
                    '��Ժʱ��;����id;��ʼʱ��;����ʱ��;
                    Set mcbrItem = mcbrToolBarҳ��.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
                    mcbrItem.Parameter = strEnterDate & ";" & RS!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                    mcbrItem.Category = strCategory
                    
                    If lngLoop + 1 <= 4 Then
                        Set cbrWeek = mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop)))
                        cbrWeek.Parameter = strEnterDate & ";" & RS!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                        cbrWeek.Category = strCategory
                    End If
                     
                    lnglast����id = Val(Nvl(RS("����ID").Value))
                    RS.MoveNext
                    strParameter = mcbrItem.Parameter
                    
                    'ָ��ҳ�Ų�Ϊ0 ���Һ͸�ҳ����Ⱦͼ�¼����ֵ
                    If T_Patient.lngPage <> 0 And Val(T_Patient.lngPage - 1) = lngLoop Then
                        strParam1 = strParameter
                        strSvrCaption1 = strCategory
                    End If
                    
                    strSvrCaption = strCategory
                End If
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
    
    
    For lngLoop = 0 To Round((DateDiff("D", CDate(ArrCode(0)), CDate(ArrCode(5))) + 1) / T_BodyStyle.lng����)

        strDateFrom = Format(CDate(ArrCode(0)) + T_BodyStyle.lng���� * lngLoop, "yyyy-MM-dd") & " 00:00:00"

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
            mblnMoved = False
            
            mstrSQL = "Select ��Ժ����ID,nvl(����ת��,0) ת�� from ������ҳ Where ����id=[1] And ��ҳid=[2] "
            Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ����ID", T_Patient.lng����ID, T_Patient.lng��ҳID)
            If RS.BOF = False Then
                T_Patient.lng����ID = Val(zlCommFun.Nvl(RS("��Ժ����ID").Value))
                If T_Patient.lng��Ժ = 1 Then mblnMoved = (Val(RS("ת��")) <> 0)
            End If
            
            '��ȡ��ʼ���¸�ʽ��������
            If Not GetStyleBody(T_Patient.lng�ļ�ID, T_Patient.lng����ȼ�, T_Patient.lngӤ��, T_Patient.lng����ID) Then Exit Function
    
            mstrSQL = "SELECT A.���,A.���� FROM(" & vbNewLine & _
                        "SELECT A.���,A.����,A.����ID,A.��ҳID FROM (SELECT 0 ���,'���˱���' AS ����,A.����ID,A.��ҳID" & vbNewLine & _
                        "            FROM ������ҳ A, ������Ϣ B" & vbNewLine & _
                        "            WHERE A.����ID = B.����ID AND A.����ID =[1] AND A.��ҳID =[2]" & vbNewLine & _
                        "            UNION ALL" & vbNewLine & _
                        "            SELECT A.���, DECODE(A.Ӥ������, NULL, NVL(C.����,B.����) || '֮��' || TRIM(TO_CHAR(A.���, '9')), A.Ӥ������) AS ����,A.����ID,A.��ҳID" & vbNewLine & _
                        "            FROM ������Ϣ B,������ҳ C,������������¼ A" & vbNewLine & _
                        "            WHERE B.����ID=C.����ID And C.����ID=A.����ID And C.��ҳID=A.��ҳID And C.����ID =[1] AND C.��ҳID =[2]) A," & vbNewLine & _
                        "            (SELECT A.����ID,A.��ҳID , NVL(A.Ӥ��,0) Ӥ�� FROM ���˻����ļ� A,�����ļ��б� B" & vbNewLine & _
                        "            WHERE A.��ʽID=B.ID AND B.����=3 AND B.����=-1 And A.����ID =[1] AND A.��ҳID =[2] GROUP BY A.����ID,A.��ҳID,NVL(A.Ӥ��,0)) B" & vbNewLine & _
                        "            WHERE A.����ID=B.����ID AND A.��ҳID=B.��ҳID AND A.���=B.Ӥ��) A" & vbNewLine & _
                        "ORDER BY A.���"
            If mblnMoved = True Then
                mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            End If
            Set RS = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng��ʽID)
            
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
            '��ȡ�����ļ��б�
            mstrSQL = "select A.ID,A.�ļ����� From ���˻����ļ� A,�����ļ��б� B" & _
               "    where A.����ID=[1] and A.��ҳId=[2] and nvl(A.Ӥ��,0)=[3] and A.��ʽID=B.ID and B.����=3 and B.����=-1 Order by A.��ʼʱ��"
            If mblnMoved = True Then
                mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            End If
            Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�ļ�", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
            cboFile.Clear
            With RS
                Do While Not .EOF
                    cboFile.AddItem Nvl(!�ļ�����)
                    cboFile.ItemData(cboFile.NewIndex) = !Id
                    .MoveNext
                    If cboFile.ListIndex = -1 And T_Patient.lng�ļ�ID = Val(cboFile.ItemData(cboFile.NewIndex)) Then
                        Call zlControl.CboSetIndex(cboFile.hWnd, cboFile.NewIndex)
                        T_Patient.lng�ļ�ID = cboFile.ItemData(cboFile.ListIndex)
                    End If
                Loop
            End With
            If cboFile.ListCount > 0 And cboFile.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboFile.hWnd, 0)
                T_Patient.lng�ļ�ID = cboFile.ItemData(cboFile.ListIndex)
            End If
            
            RaiseEvent zlFileChange(False, T_Patient.lng�ļ�ID, T_Patient.lngӤ��)
            
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
            mstrEnterDate = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
            strStartDate = Format(varParam(2), "YYYY-MM-DD HH:mm:ss")
            strEndDate = Format(varParam(3), "YYYY-MM-DD HH:mm:ss")
            mintPage = Val(varParam(4))
            glngCurPage = mintPage + 1
            mstrEndDate = Format(varParam(5), "YYYY-MM-DD HH:mm:ss")
            If mbln��Ժ = True Then
                '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
                mstrEndDate = Format(RetrunEndTimeNew(CDate(mstrEnterDate), CDate(mstrEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
                strEndDate = Format(RetrunEndTimeNew(CDate(mstrEnterDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
            End If
            If strStartDate & ";" & strEndDate = picMain.Tag And mblnResize = True Then
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
                    'Debug.Print Now & ":���ر������"
                    Call ShowDowntab '�����±������
                    Call picDraw_Paint '���ڴ���Copy������PIC
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
                
                mstrSQL = "Select Decode(a.Ӥ������,Null,NVL(C.����,B.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,a.Ӥ���Ա�,a.����ʱ�� " & _
                    " From ������Ϣ B,������ҳ C,������������¼ A " & _
                    " Where B.����ID=C.����ID And C.����ID=A.����ID And C.��ҳID=A.��ҳID And C.����id=[1] And C.��ҳid=[2] And a.���=[3]"
                Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡӤ����Ϣ", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
                If RS.BOF = False Then
                    txtCard(0).Text = RS("Ӥ������").Value
                    txtCard(5).Text = RS("Ӥ���Ա�").Value
                End If
            End If
            txtCard(6).Text = GetElementValue("����", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��, mstr��ʼʱ��)
        Case "����������ʾ����"
            If T_Patient.lng�༭ = 0 Then Exit Function
            If mstr��ʼʱ�� <> "" Then
                '����ѡ�����
                intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                intCOl = intCOl - T_BodyStyle.lng������ + 1
                If intCOl < mintColMin Then intCOl = mintColMin
                
                '����õ��з��ص�ʱ�䷶Χ
                If Trim(strParam) <> "" Then '�����±༭���������ʾ�Ǵ���ʱ��(��Ϊ�����������µ�ˢ�º�,�ᶨλ����һ��)
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
                Else
                    strTime = Split(GetCurveDateNew(intCOl, mstr��ʼʱ��, gintHourBegin), ";")(0)
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
                    RaiseEvent zlDataChange(True)
                End If
            End If
            
        Case "�������ݱ༭"
            If T_Patient.lng�༭ = 0 Then Exit Function
            Dim strCurDate As String, strDay As String
            If mstr��ʼʱ�� <> "" Then
            If picMain.Tag = "" Then picMain.Tag = mstr��ʼʱ�� & ";" & mstr����ʱ��
                strCurDate = zlDatabase.Currentdate
                '����õ��з��ص�ʱ�䷶Χ
                If Trim(strParam) <> "" Then
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss") & ";" & Format(varParam(1), "YYYY-MM-DD HH:mm:ss")
                Else
                    '����ѡ�����
                    intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                    intCOl = intCOl - T_BodyStyle.lng������ + 1
                    If intCOl < mintColMin Then intCOl = mintColMin
                    strTime = GetCurveDateNew(intCOl, mstr��ʼʱ��, gintHourBegin)
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
                    RaiseEvent zlDataChange(True)
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

Private Function InitData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng��Ժ As Long, ByVal lng�༭ As Long, ByVal intӤ�� As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
            
    '��ȡ��������
    T_Patient.lng����ID = lng����ID
    T_Patient.lng��ҳID = lng��ҳID
    T_Patient.lng��Ժ = lng��Ժ
    T_Patient.lng�༭ = lng�༭
    
    '���س�ʼ������,��������ʱ���
    Call InitPara(T_BodyStyle.blnר��)
    
    '���б�Ҫ�ļ��
    '��ȡ���˵�ǰ����ȼ�
    T_Patient.lng����ȼ� = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵�ǰ����ȼ�", T_Patient.lng����ID, T_Patient.lng��ҳID)
    If rsTemp.BOF = False Then T_Patient.lng����ȼ� = zlCommFun.Nvl(rsTemp("����ȼ�"), 3)

    '����Ƿ�������������Ŀ
    gstrSQL = " Select 1 From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
              " Where C.��Ŀ���=A.��Ŀ��� " & _
                        "AND C.��ĿID=B.ID(+) " & _
                        "AND C.����ȼ�>=[1] " & _
                        "And A.��¼��=1 And RowNum<2 And C.��Ŀ���<>" & gint����
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����������Ŀ", T_Patient.lng����ȼ�)
    If rsTemp.EOF Then
        MsgBox "����Ҫ��һ��������Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�жϸò����Ƿ��Ѿ�ת��
    If T_Patient.lng����ID > 0 And T_Patient.lng��Ժ = 1 Then
        gstrSQL = "select nvl(����ת��,0) ת�� from ������ҳ where ����ID=[1] and ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鲡���Ƿ�ת��", T_Patient.lng����ID, T_Patient.lng��ҳID)
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
    vsf.Body.RowHeight(vsf.FixedRows) = 300
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
    Dim bln�����ʾ��Ժ As Boolean, bln��ʾ��� As Boolean
    On Error GoTo hErr
    
    strStart = mstr��ʼʱ��
    strTo = mstr����ʱ��
    bln��ʾ��� = (Val(zlDatabase.GetPara("���µ���ʾ���", glngSys, 1255, 1)) = 1)
    If Not bln��ʾ��� Then
        lblCard(7).Visible = False
        txtCard(7).Visible = False
    Else
        lblCard(7).Visible = True
        txtCard(7).Visible = True
    End If
    
    If CStr(mstrEndDate) < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln��Ժ = False Then
        mstrEndDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If mintAllPage = mintPage + 1 Then
        If CStr(mstr����ʱ��) < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln��Ժ = False Then
            mstr����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    txtCard(3).Text = ""
    
    '������������������¼���ʱ�䣬��Ӥ�����µ��Ŀ�ʼʱ��
    If T_Patient.lngӤ�� > 0 Then
        mstrSQL = " Select  b.����ʱ�� From ������������¼ B Where ����id=[1] And ��ҳid=[2] And ���=[3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ��������Ϣ", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID), T_Patient.lngӤ��)
        If rsTmp.BOF = False Then
            txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("����ʱ��").Value), "yyyy-MM-dd")
        End If
    End If
    
    '�˴�����ʱ��ת��
    intCOl = GetCurveColumnNew(CDate(strStart), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strStart = Split(GetCurveDateNew(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(0)
    
    If CDate(strStart) < CDate(mstr��ʼʱ��) Then
        strStart = Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    intCOl = GetCurveColumnNew(CDate(strTo), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strTo = Split(GetCurveDateNew(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(1)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "�䶯��¼", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID))
    If rsTmp.BOF = False Then
        If txtCard(3).Text = "" And bln�����ʾ��Ժ = True Then txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("��ʼʱ��").Value), "yyyy-MM-dd")
    End If
    
    '��ȡ���˻�����Ϣ
    mstrSQL = " Select  NVL(A.����,b.����) ����,A.סԺ��,A.��Ժ���� ��Ժʱ��,NVL(A.�Ա�,b.�Ա�) �Ա�,NVL(A.����,b.����) ����" & _
        " From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ������Ϣ", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID))
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
            " From ���˱䶯��¼ a,���ű� b,���ű� c " & _
            " Where a.����id=[1] And a.��ҳid=[2] And a.����id Is Not Null And a.����id=b.id and a.����id=c.id  And NVL(A.���Ӵ�λ,0)=0 " & _
            " And a.��ʼʱ��-4/24<=[3] And Nvl(a.��ֹʱ��,Sysdate)>=[4] Order By a.��ʼʱ��"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���˿��ҡ����ŵ���Ϣ", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID), CDate(mstr����ʱ��), CDate(mstr��ʼʱ��))
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
    If bln��ʾ��� = True Then
        '��ȡ��ϵ���Сʱ��
        strStart = GetDiagnoseMinTime(T_Patient.lng����ID, T_Patient.lng��ҳID, CDate(strStart), mblnMoved)
        '��ȡ���������Ϣ
        mstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As ������ From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "������", "������", Val(T_Patient.lng����ID), Val(T_Patient.lng��ҳID), CDate(strStart))
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
    End If
    
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
    Dim lng���� As Long
    Dim PicRect  As RECT
    On Error GoTo Errhand
    
    '��ȡ��������
    Call InitPublicData
    
    mbln�������� = True
    lng���� = T_BodyStyle.lng������ * T_BodyStyle.lng����
    T_DrawClient.�е�λ = T_BodyStyle.lng�����п� \ Screen.TwipsPerPixelX
    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    '�õ���������
    lngCount = CurveCount
    
    '�������µ��̶���������ұ߾�
    If lngCount <= 3 Then
        T_DrawClient.�̶�����.Right = T_DrawClient.�̶�����.Left + T_BodyStyle.lng�̶ȿ�� \ Screen.TwipsPerPixelX
    Else
        T_DrawClient.�̶�����.Right = T_DrawClient.�̶�����.Left + T_BodyStyle.lng�̶ȿ�� \ Screen.TwipsPerPixelX
    End If
    
    lngWith = T_DrawClient.�е�λ * Screen.TwipsPerPixelX
    
    With mshUpTab
        .Cols = lng���� + 1
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        .Cell(flexcpText, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
        .ColWidthMin = lngWith
        .RowHeightMin = T_BodyStyle.lng���߶�
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(2) = True
        .ColWidth(0) = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) * Screen.TwipsPerPixelX
        .RowHeight(-1) = .RowHeightMin
        .TextMatrix(0, 0) = Split(T_BodyStyle.str��ͷ����, "@")(0)
        If UBound(Split(T_BodyStyle.str��ͷ����, "@")) > 0 Then
            .TextMatrix(1, 0) = Split(T_BodyStyle.str��ͷ����, "@")(1)
        Else
            .TextMatrix(1, 0) = IIf(T_Patient.lngӤ�� = 0, "ס Ժ �� ��", "�� �� �� ��")
        End If
        If UBound(Split(T_BodyStyle.str��ͷ����, "@")) > 1 Then
            .TextMatrix(2, 0) = Split(T_BodyStyle.str��ͷ����, "@")(2)
        Else
            .TextMatrix(2, 0) = "����������"
        End If
        If UBound(Split(T_BodyStyle.str��ͷ����, "@")) > 2 Then
            .TextMatrix(3, 0) = Split(T_BodyStyle.str��ͷ����, "@")(3)
        Else
            .TextMatrix(3, 0) = "ʱ       ��"
        End If
        
        For intCOl = 1 To .Cols - 1
            .ColWidth(intCOl) = lngWith
        Next
        .Redraw = flexRDBuffered
    End With
    
    '�ϲ���Ԫ�����
    For intRow = 0 To 2
        Call UniteCellCol(mshUpTab, T_BodyStyle.lng������, intRow, mshUpTab.FixedCols)
    Next intRow
    
    If blnInitUpdate = True Then Call ShowUptab
    
    With vsf
        .Cols = 0
        .NewColumn "", 0, 1
        .NewColumn "��Ŀ", mshUpTab.ColWidth(0) + 10, 1
    
        For intCOl = 1 To lng����
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
        .Cols = lng���� + 4
        .Rows = 1
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(0) = True
        .Tag = 0
        .RowHeightMin = T_BodyStyle.lng�±��߶�
        .RowHeight(-1) = T_BodyStyle.lng�±��߶�
        
        For intCOl = .FixedCols To .Cols - 1
            .ColWidth(intCOl) = mshUpTab.ColWidth(1)
            If (intCOl - .FixedCols + 1) Mod 2 = 0 Then
                .Cell(flexcpBackColor, 0, intCOl, .Rows - 1, intCOl) = &H80000013
            End If
        Next intCOl

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
    mItemNO.���� = 0
    Dim bln���� As Boolean
    Dim intRows  As Integer
    intRows = GetRows(bln����, T_BodyItem.str�����Ŀ)
    mintRepairRows = T_BodyStyle.lng������ + intRows
    mbln��ʾƤ�� = (Val(zlDatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0")) = 1)
    
    '�������Ƿ��Ǳ����Ŀ
    gstrSQL = "select ��¼�� From ���¼�¼��Ŀ where ��Ŀ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���¼�¼��Ŀ", gint����)
    If rsTemp.RecordCount > 0 Then
         mintRepairRows = mintRepairRows - IIf(Val(Nvl(rsTemp!��¼��)) = 2 And bln���� = True, 1, 0)
    End If
    If mintRepairRows < 0 Then mintRepairRows = 0

    '�������б����Ŀ�������̶���Ŀ�������ݵĻ��Ŀ
    Set rsTemp = GetAppendGridItemNew(T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lng����ȼ�, T_Patient.lngӤ��, Int(CDate(mstr��ʼʱ��)), CDate(mstr����ʱ��), IIf(T_Patient.lngӤ�� = 0, 1, 2), T_Patient.lng����ID, T_BodyItem.str�����Ŀ, mblnMoved)
    With rsTemp
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            mshDownTab.Rows = 0
            Call AppenGridItemNew(rsTemp)
        Else
            mshDownTab.Rows = 0
        End If
    End With
    
    mshDownTab.Rows = mintRepairRows
    mshDownTab.RowHeightMin = T_BodyStyle.lng�±��߶�
    mshDownTab.RowHeight(-1) = mshDownTab.RowHeightMin
    
    '������ʣ�µĿ���
    If mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
        For intRow = Val(mshDownTab.Tag) To mshDownTab.Rows - 1

            mshDownTab.MergeRow(intRow) = True
            For intCOl = 0 To mshDownTab.FixedCols
                strPace = " " & String(intCOl, " ") & String(intRow, " ")
                mshDownTab.TextMatrix(intRow, intCOl) = strPace & "" & strPace
            Next intCOl
            
            Call UniteCellCol(mshDownTab, T_BodyStyle.lng������, intRow, mshDownTab.FixedCols)
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
    If mshDownTab.Rows > mshDownTab.FixedRows Then
        For intCOl = mshDownTab.FixedCols To mshDownTab.Cols - 1
            If (intCOl - mshDownTab.FixedCols + 1) Mod 2 = 0 Then
                mshDownTab.Cell(flexcpBackColor, 0, intCOl, mshDownTab.Rows - 1, intCOl) = &HF7ECE6
            End If
        Next intCOl
        mshDownTab.Cell(flexcpAlignment, 0, 0, mshDownTab.Rows - 1, mshDownTab.Cols - 1) = 4
    End If
    Call picBack_Resize
    Call Paint_CanvasNew(mblnAutoAdjust) '��ʼ����������
    Call picBack_Resize
    
    PicRect.Top = 0
    PicRect.Left = 0
    PicRect.Right = picCommText.Width \ Screen.TwipsPerPixelX
    PicRect.Bottom = picCommText.Height \ Screen.TwipsPerPixelY
    picCommText.Cls
    Call PrintCurveInfo(picCommText, PicRect)
    
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
    Dim lngValue  As Long, intCOl As Integer
    Dim lngDays   As Long
    Dim i As Long, j As Long
    Dim lngColor  As Long
    Dim intMinCol As Integer, intMaxCol As Integer
    Dim strTmp As String
    Dim arrOperDay, strTmp1 As String
    Dim rsTmp  As New ADODB.Recordset
    Dim strʱ�� As String
    Dim intDays As Integer
    Dim lng���� As Long
    Dim lngWith As Long
    Dim lng���� As Long, lngƵ�� As Long, lngʱ���� As Long
    Dim bln������ʾ As String
    Dim str����ʱ�� As String
    

    On Error GoTo Errhand
    
    lng���� = T_BodyStyle.lng����
    lngƵ�� = T_BodyStyle.lng������
    lngʱ���� = T_BodyStyle.lngʱ����

    With mshUpTab
        
        lngValue = 0
        gstrSQL = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As ��ʼ���� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡסԺ����", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, Int(CDate(mstr��ʼʱ��)))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("��ʼ����").Value
        End If
        
        '�ϱ��ʽ�е�Ԫ��ϲ��ģ��˴���Ҫ���д���
        For intCOl = 1 To lng����
            .ColData(intCOl) = 0
            .Row = 0
            .Col = intCOl
            .ColAlignment(intCOl) = 4

            strTmp = Format(CDate(mstr��ʼʱ��) + intCOl - 1, "yyyy-MM-dd")

            lngDays = lngValue + (intCOl - 1)
            
            For i = 1 To lngƵ��
                .Row = 0
                .Col = (intCOl - 1) * lngƵ�� + i
                
                If Right(strTmp, 5) = "01-01" Then
                    'һ��ĵ�һ��
                    .Text = strTmp
                ElseIf strTmp = Format(mstrEnterDate, "yyyy-MM-dd") Then
                    '��Ժ��һ�죬д�����
                    .Text = strTmp
                ElseIf intCOl = 1 Then
                    '70299:������,2014-4-4,ÿҳ����������ʾΪ������(1-��-��-��,0:Ĭ�ϸ�ʽ:��������ʾ)
                    If Val(zlDatabase.GetPara("�������ڸ�ʽ", glngSys, 1255, "0")) = 1 Then
                        .Text = strTmp
                    Else
                        .Text = Right(strTmp, 5)
                    End If
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
            Call CalcMinMaxColNew(picMain.Tag, intMinCol, intMaxCol)
            mintColMin = intMinCol
            mintColMax = intMaxCol
            
            With picDisplay
                .Left = ((((intMaxCol - 1) \ lngƵ��) + 1) * lngƵ�� - 1) * mshUpTab.ColWidth(intMinCol) + mshUpTab.ColWidth(0)
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

        ReDim mstrOpValue(T_BodyStyle.lng����) As String
        ReDim mstrOpdays(T_BodyStyle.lng����) As String
        
        For i = 1 To T_BodyStyle.lng����
            mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng������ + 1))
            mstrOpdays(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng������ + 1))
        Next i
        
        '��ȡ�����־������ֹͣ������־
        mintOpDays = Val(zlDatabase.GetPara("�������ע����", glngSys, 1255, "10"))
        mblnStopFlag = (Val(zlDatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0")) = 1)
        bln������ʾ = (Val(zlDatabase.GetPara("����������14���Ժ�����ʾ", glngSys, 1255, "0")) = 1)
        
        '51338,������,2012-07-06
        strTmp = zlDatabase.GetPara("��������ȱʡ��ʽ", glngSys, 1255, "2")
        If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
            mintOpFormat = Val(strTmp)
        Else
            mintOpFormat = 0
        End If
        
        strTmp = ""
        '��ʾ��ǰ�ε��������
        gstrSQL = "select B.����ʱ�� ʱ��" & _
            "   From ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & _
            "   where A.ID=B.�ļ�ID And  B.ID=C.��¼ID And A.ID=[1] And nvl(A.Ӥ��,0)=[4]" & _
            "   and A.����ID=[2] and A.��ҳID=[3] and C.��¼����=4 And NVL(C.���Ժϸ�,0)<>1 and C.��ֹ�汾 is null" & _
            "   and B.����ʱ�� between [5] and [6] order by B.����ʱ��"
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
            gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", Val(T_Patient.lng�ļ�ID), T_Patient.lng����ID, T_Patient.lng��ҳID, Val(T_Patient.lngӤ��), Int(CDate(mstr��ʼʱ��) - 14), CDate(mstr����ʱ��))

        str����ʱ�� = mstr����ʱ��
        
        Do While Not rsTmp.EOF
            strʱ�� = Format(rsTmp("ʱ��"), "YYYY-MM-DD")
            
             '�����:56005,����,2013-04-27
            If Not rsTmp.EOF Then
                If bln������ʾ And DateDiff("d", CDate(Format(strʱ��, "YYYY-MM-DD")), str����ʱ��) < mintOpDays Then
                    str����ʱ�� = Format(DateAdd("D", mintOpDays, CDate(Format(strʱ��, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
                End If
            End If
            
            For i = 1 To lng����
                If DateDiff("d", mstr��ʼʱ��, str����ʱ��) + 1 >= i Then
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
                                    If mintOpFormat = 3 Then
                                        mstrOpValue(i) = mstrOpValue(i) & "/" & intDays
                                    Else
                                        mstrOpValue(i) = intDays & "/" & mstrOpValue(i)
                                    End If
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
            "   And A.����ID=[2] And A.��ҳID=[3] And C.��¼����=4 And NVL(C.���Ժϸ�,0)<>1 and C.��ֹ�汾 is null" & _
            "   And B.����ʱ�� <[5] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", Val(T_Patient.lng�ļ�ID), T_Patient.lng����ID, T_Patient.lng��ҳID, Val(T_Patient.lngӤ��), Int(CDate(mstr��ʼʱ��)))
        
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
            gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        lng���� = 0
        If rsTmp.BOF = False Then lng���� = Val(rsTmp("����"))
        For i = 1 To lng����
            If DateDiff("d", mstr��ʼʱ��, str����ʱ��) + 1 >= i Then
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
                        '�����:57771,���Σ�2013-05-02
                        If mintOpFormat = 3 Then
                            strTmp1 = Switch(lng���� = 1, "����", lng���� = 2, "��2", lng���� = 3, "��3", lng���� = 4, "��4", lng���� = 5, "��5", lng���� = 6, "��6", lng���� = 7, "��7", lng���� = 8, "��8", lng���� = 9, "��9", lng���� = 10, "��10", lng���� = 11, "��11", lng���� = 12, "��12")
                        Else
                            strTmp1 = Switch(lng���� = 1, "��", lng���� = 2, "��", lng���� = 3, "��", lng���� = 4, "��", lng���� = 5, "��", lng���� = 6, "��", lng���� = 7, "��", lng���� = 8, "��", lng���� = 9, "��", lng���� = 10, "��", lng���� = 11, "��", lng���� = 12, "��")
                        End If
                       
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
                                mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng������ + 1)) & "0" & .TextMatrix(2, ((i - 1) * T_BodyStyle.lng������ + 1))
                            Case 2 '--��ʾ����
                                If strTmp = "��" Then
                                    mstrOpValue(i) = 0
                                Else
                                    mstrOpValue(i) = strTmp & "-0"
                                End If
                            Case 3
                                  If strTmp = "���� 1" Then
                                    mstrOpValue(i) = "����"
                                Else
                                    mstrOpValue(i) = strTmp
                                End If
                            Case Else '--����ʾ
                                 mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng������ + 1))
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
                            Case 3
                                  If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = mstrOpValue(i) & "/" & strTmp
                                Else
                                    mstrOpValue(i) = strTmp
                                End If
                            Case Else  '--����ʾ
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng������ + 1))
                                End If
                        End Select
                    End If
                    .Row = 2
                    For j = 1 To T_BodyStyle.lng������
                        .Col = j + (i - 1) * T_BodyStyle.lng������
                        .Text = mstrOpValue(i)
                    Next j
                Else
                    .Row = 2
                    For j = 1 To T_BodyStyle.lng������
                        .Col = j + (i - 1) * T_BodyStyle.lng������
                        .Text = mstrOpValue(i)
                    Next j
                End If
            End If
        Next i
        '�趨���ڣ�סԺ�����ı���ɫ
        mshUpTab.Cell(flexcpForeColor, 0, mshUpTab.FixedCols, 1, mshUpTab.Cols - 1) = 16711680
        '�趨���� �����ı���ɫ
        '51283,������,2012-07-11
        lngColor = Val(zlDatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "255"))
        mshUpTab.Cell(flexcpForeColor, 2, mshUpTab.FixedCols, 2, mshUpTab.Cols - 1) = lngColor

        lngWith = T_DrawClient.�е�λ * Screen.TwipsPerPixelX
        mshUpTab.ColWidthMin = lngWith
        'mshUpTab.Cell(flexcpWidth, 0, 1, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = lngWith
        For intCOl = 1 To mshUpTab.Cols - 1
            mshUpTab.ColWidth(intCOl) = lngWith
        Next intCOl
        
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
    Dim strItems As String, strItemName As String, strSQL As String
    Dim lngItemCode As Long
    Dim strPace As String
    Dim str��Ŀ���� As String, str��Ŀ����1 As String
    Dim int��¼Ƶ�� As Integer, int��Ŀ���� As Integer, int��Ŀ���� As Integer, int��Ŀ��ʾ As Integer, strTabItemTemp As String
    Dim strBegin As String, str��� As String, strPart As String
    Dim int����ѹ As Integer, int����ѹ As Integer, Int�к� As Integer
    Dim blnColor As Boolean
    Dim lngColor As Long
    Dim arrTmpString0() As String, arrTmpString1() As String, arrTmpString2() As String
    Dim blnAdd As Boolean, blnValue As Boolean
    Dim SinX As Single
    Dim i As Integer, j As Integer
    Dim int����λ�� As Integer, intValue As Integer, int������������ʽ As Integer
    Dim bln���ܵ��� As Boolean, bln¼��Сʱ As Boolean
    Dim arrTmp() As String
    Dim dtBegin As Date, dtEnd As Date
    Dim int������ As Integer
    '73316�����޸���ر�������
    Dim arrBreathe, blnBreathe As Boolean, intBegin As Integer, intEnd As Integer
    Dim lngX As Long, lngY As Long, lngBottomY As Long
    Dim blnBreatheShowType As Boolean  '����:����Ϊ���ʱ�����������ʽ
    
    On Error GoTo Errhand
    
    ReDim mstrNewString(mintRepairRows, T_BodyStyle.lng���� - 1)
    ReDim mlngNewHeight(mintRepairRows)
    For i = 0 To UBound(mlngNewHeight)
        mlngNewHeight(i) = mshDownTab.RowHeightMin
    Next i
    ReDim arrTmpString0(1 To T_BodyStyle.lng������ * T_BodyStyle.lng����) As String
    ReDim arrTmpString1(1 To T_BodyStyle.lng������ * T_BodyStyle.lng����) As String
    ReDim arrTmpString2(1 To T_BodyStyle.lng������ * T_BodyStyle.lng����) As String
    
    'mstrNewString = Split(String(T_BodyStyle.lng����-1, ";"), ";")
    int������������ʽ = zlDatabase.GetPara("����������", glngSys, 1255, 0)
    If int������������ʽ < 0 Or int������������ʽ > 3 Then int������������ʽ = 0
    bln���ܵ��� = (Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    mbln�೦�����ӷ�ĸ��ʾ = (Val(zlDatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
    bln¼��Сʱ = (Val(zlDatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, 0)) = 1)
    '73316:������,2014-06-26,���첿��ҽԺҪ��:
    '��1����������ɫ���ں�������Ӧʱ������д���������κ������½�����д�����Ϻ���
    '��2������������ʶ������ʼ��Ӧʱ������ɫ�ֱ������µ������������Ϸ�����
    '��д�������������á�������ʶ��ʼ����ֹ�ԡ�������ʶ���������趨Ƶ�������ֱ�ʾ������
    'ɫ���ں�������Ӧʱ������д���������κ������½�����д�����Ϻ�
    '2----��ʼ����������� ������Ϊͼ�����
    blnBreatheShowType = (Val(zlDatabase.GetPara("�����������������ʽ", glngSys, 1255, 0)) = 1)
    
    int������ = T_BodyStyle.lng������
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
            strItemName = mshDownTab.TextMatrix(intRow, 3)
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
        "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID   AND B.ID=[1] AND Nvl(B.Ӥ��,0)=[7] " & _
        "   AND B.����id=[2]  AND B.��ҳid=[3] AND INSTR([6],decode(E.��Ŀ����,2,C.���²�λ || D.��¼�� ,D.��¼��))>0 " & _
        "   AND D.��Ŀ���=C.��Ŀ���  AND MOD(c.��¼����,10)=1  AND E.��Ŀ���=D.��Ŀ��� " & _
        "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null AND ��¼��=2"
    
    '��ȡ�����±��Ļ�����Ŀ
    strSQL = "  SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & _
        "   Decode(d.��Ŀ����, 2, c.���²�λ || d.��Ŀ����, d.��Ŀ����) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,D.��Ŀ����" & _
        "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,(SELECT A.��Ŀ���,A.��Ŀ����, A.��Ŀ����,B.����� FROM �����¼��Ŀ A,���������Ŀ B" & vbNewLine & _
        "       WHERE A.��Ŀ���=B.��� AND  B.����� is not NULL " & vbNewLine & _
        "       AND NVL(A.Ӧ�÷�ʽ,0)=1 AND NVL(A.����ȼ�,0)>=[8] AND NVL(A.���ò���,0) IN (0,[9])" & vbNewLine & _
        "       AND (A.���ÿ���=1 OR (A.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D WHERE D.��Ŀ���=A.��Ŀ��� AND D.����ID=[10])))) D" & _
        "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID AND Instr([6], Decode(d.��Ŀ����, 2, c.���²�λ || d.��Ŀ����, d.��Ŀ����)) = 0  AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
        "   AND B.����id=[2]  AND B.��ҳid=[3]  AND D.��Ŀ���=C.��Ŀ���  AND C.��¼����=1" & _
        "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null"
        
    gstrSQL = "Select /*+ Rule*/ ID,ʱ��,��¼����,��ʾ,���,���²�λ,δ��˵��,������Դ,��Ŀ����,��Ŀ���,��ԴID,����,��Ŀ���� From (" & _
        "   " & gstrSQL & " UNION ALL " & strSQL & ")" & _
        "   Order By  Decode(��Ŀ����,'����ѹ',0,1)," & strItems & ",ʱ��"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���±������", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, _
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
                intCOl = GetCurveColumnNew(rsTemp!ʱ��, mstr��ʼʱ��, gintHourBegin) + vsf.FixedCols - 1
                str��� = zlCommFun.Nvl(rsTemp!���) & ";" & Nvl(rsTemp!���²�λ)
                If intCOl < vsf.Cols Then
                    If arrTmpString1(intCOl - vsf.FixedCols + 1) <> "" Then
                        If (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) <> 1 And Val(zlCommFun.Nvl(!��ʾ, 0)) <> 1) Or _
                            (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 1 And Val(zlCommFun.Nvl(!��ʾ, 0)) = 1) Then
                            
                            '����Ǹ����ص�ʱ�����
                            SinX = GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
                            blnAdd = GetCanvasCenterNew(CDate(Format(arrTmpString1(intCOl - vsf.FixedCols + 1), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
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
    For i = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
        If Val(arrTmpString2(i)) = 2 Then arrTmpString0(i) = ""
    Next i
    
    '2----��ʼ����������� ������Ϊͼ�����
    int����λ�� = 0
    blnValue = False
    arrBreathe = Array(): blnBreathe = False
     'ѭ���������ֵ
    vsf.Cell(flexcpForeColor, 1, vsf.FixedCols, 1, vsf.Cols - 1) = Val(vsf.Tag)
    If blnBreatheShowType = True Then
        vsf.Body.OwnerDraw = flexODNone
    Else
        vsf.Body.OwnerDraw = flexODOver
    End If
    
    For i = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
        intCOl = i + vsf.FixedCols - 1
        If InStr(1, arrTmpString0(i), ";") > 0 Then
            str��� = Split(arrTmpString0(i), ";")(0)
            strPart = Split(arrTmpString0(i), ";")(1)
        Else
            str��� = arrTmpString0(i)
            strPart = ""
        End If
        '����ÿ�κ�������һ����Χ
        If mbln�������� = False Then
            If strPart = "������" And IsNumeric(str���) Then
                If blnBreathe = False Then
                    ReDim Preserve arrBreathe(UBound(arrBreathe) + 1)
                    arrBreathe(UBound(arrBreathe)) = i & ";" & i
                    blnBreathe = True
                Else
                    arrBreathe(UBound(arrBreathe)) = Split(arrBreathe(UBound(arrBreathe)), ";")(0) & ";" & i
                End If
            Else
                blnBreathe = False
            End If
        End If
        '��ӡ����ֵ���������ӡ�� ��һ��ʼ��������
        If IsNumeric(str���) Then
            vsf.TextMatrix(1, intCOl) = str���
            If blnValue = False Then
                intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                blnValue = True
                int����λ�� = 2
            End If
            
            If int������������ʽ = 0 Or int������������ʽ = 2 Then '˳��������ʾ
                If intCOl Mod 2 = intValue Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int������������ʽ = 0, flexAlignCenterTop, flexAlignCenterBottom)
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int������������ʽ = 0, 1, 2)
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int������������ʽ = 0, flexAlignCenterBottom, flexAlignCenterTop)
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int������������ʽ = 0, 2, 1)
                    End If
                End If
                
            Else        '������ʱ����֮��������ʾ
                If int����λ�� = 2 Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int������������ʽ = 1, flexAlignCenterTop, flexAlignCenterBottom)
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int������������ʽ = 1, 1, 2)
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int������������ʽ = 1, flexAlignCenterBottom, flexAlignCenterTop)
                    If strPart <> "������" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int������������ʽ = 1, 2, 1)
                    End If
                End If
                
                int����λ�� = int����λ�� + 1
                If int����λ�� > 2 Then int����λ�� = 1
            End If
        End If
    Next i
    
    '��ʼ�ں������Ϸ��������
    If blnBreatheShowType = True Then
        lngBottomY = T_DrawClient.����������.Bottom
        For i = 0 To UBound(arrBreathe)
            intBegin = Split(arrBreathe(i), ";")(0)
            intEnd = Split(arrBreathe(i), ";")(1)
            '�������������
            strPart = "������"
            Call SetTextColor(mlngMemDC, Val(vsf.Tag))
            T_Size.H = mobjDraw.TextHeight("��") / T_TwipsPerPixel.Y
            T_Size.W = mobjDraw.TextWidth("��") / T_TwipsPerPixel.X
            '����GetTextRect����Ĭ�ϸ�X+1���Դ˴�-1
            If intBegin = intEnd Then
                If T_DrawClient.�е�λ >= T_Size.W + 6 Then
                    lngX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ + ((T_DrawClient.�е�λ - T_Size.W - 6) \ 2) - 1
                Else
                    lngX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ - ((T_Size.W + 6 - T_DrawClient.�е�λ)) - 1
                End If
            Else
                If T_DrawClient.�е�λ >= T_Size.W + 3 Then
                    lngX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ + ((T_DrawClient.�е�λ - T_Size.W - 3) \ 2) - 1
                Else
                    lngX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ - ((T_Size.W + 3 - T_DrawClient.�е�λ)) - 1
                End If
            End If
            lngY = lngBottomY - T_Size.H * Len(strPart)
            For j = 1 To Len(strPart)
                Call GetTextRect(mobjDraw, lngX, lngY, Mid(strPart, j, 1), 0, False)
                Call DrawText(mlngMemDC, Mid(strPart, j, 1), -1, T_LableRect, DT_CENTER)
                lngY = lngY + T_Size.H
            Next j
            '��ʼ�����ϼ�ͷ�����������¼�ͷ
            If intBegin = intEnd Then
                lngY = T_Size.H * Len(strPart) - T_Size.H
                lngX = lngX + T_Size.W + 3
                Call DrawLine(mlngMemDC, lngX, lngBottomY - lngY - (T_Size.H \ 2), lngX, lngBottomY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
                lngX = lngX + 3
                Call DrawLine(mlngMemDC, lngX, lngBottomY - (T_Size.H \ 2), lngX, lngBottomY - lngY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
            Else
                lngY = T_Size.H * Len(strPart) - T_Size.H
                lngX = lngX + T_Size.W + 3
                Call DrawLine(mlngMemDC, lngX, lngBottomY - lngY - (T_Size.H \ 2), lngX, lngBottomY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
                lngX = T_DrawClient.��������.Left + (intEnd - 1) * T_DrawClient.�е�λ + T_DrawClient.�е�λ \ 2
                Call DrawLine(mlngMemDC, lngX, lngBottomY - (T_Size.H \ 2), lngX, lngBottomY - lngY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
            End If
        Next i
    End If
    
    'Debug.Print "���ݿ�ʼ---" & Now
    '��ȡ�����Ŀ������Ϣ
    With mshDownTab
        lngItemCode = 0
        str��Ŀ���� = ""
        For intRow = .FixedRows To .Tag - 1
            str��Ŀ����1 = .TextMatrix(intRow, 3)
            blnColor = False
            If str��Ŀ����1 & ";" & .RowData(intRow) <> str��Ŀ���� & ";" & lngItemCode Then
                
                lngItemCode = .RowData(intRow)
                str��Ŀ���� = str��Ŀ����1
                int��Ŀ���� = Val(Split(.TextMatrix(intRow, 1), ",")(0))
                int��¼Ƶ�� = Val(Split(.TextMatrix(intRow, 1), ",")(2))
                int��Ŀ��ʾ = Val(Split(.TextMatrix(intRow, 1), ",")(3))
                int��Ŀ���� = Val(Split(.TextMatrix(intRow, 1), ",")(4))
                strTabItemTemp = Val(Split(.TextMatrix(intRow, 1), ",")(6)) & ";" & Split(.TextMatrix(intRow, 1), ",")(7)
                blnColor = (int��Ŀ���� = 2 And int��Ŀ���� = 1 And int��Ŀ��ʾ = 0)
                
                For intDay = 0 To T_BodyStyle.lng���� - 1
                    strBegin = DateAdd("D", intDay, CDate(mstr��ʼʱ��))
                    If CDate(strBegin) > CDate(mstr����ʱ��) Then strBegin = mstr����ʱ��
                    int����ѹ = 0
                    int����ѹ = 0
                    Int�к� = 0
                    'ѭ���õ�ĳ����Ŀĳ���������Ϣ
                    Set rsDownTab = ReturnItemRecord(rsTemp, Int(CDate(strBegin)), CDate(mstrEnterDate), lngItemCode & ";" & str��Ŀ���� & ";" & _
                                int��¼Ƶ�� & ";" & int��Ŀ��ʾ & ";" & int��Ŀ���� & ";" & strTabItemTemp, bln���ܵ���, bln¼��Сʱ)
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
                                intCOl = intDay * int������ + .FixedCols
                                intColCount = int������
                                strPace = " "
                            Case 2
                                intRow1 = intRow
                                intCOl = (intCOl - 1) * (int������ / 2) + intDay * int������ + .FixedCols
                                intColCount = (int������ / 2)
                                strPace = String(intCOl, " ")
                            Case 3
                                intRow1 = intRow + (intCOl - 1)
                                intCOl = intDay * int������ + .FixedCols
                                intColCount = int������
                                strPace = " "
                            Case 4
                                intRow1 = intRow + Fix((intCOl - 1) / 2)
                                Select Case intCOl
                                    Case 1, 3
                                        intCOl = 1
                                    Case 2, 4
                                        intCOl = 2
                                End Select
                                intCOl = (intCOl - 1) * (int������ / 2) + intDay * int������ + .FixedCols
                                intColCount = int������ / 2
                                strPace = String(intCOl, " ")
                            Case 6
                                intRow1 = intRow
                                intCOl = (intCOl - 1) * (int������ / 6) + intDay * int������ + .FixedCols
                                intColCount = int������ / 6
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
                                                    mrsCurInfo.Filter = "����='" & str��� & "'"
                                                    If Not mrsCurInfo.EOF Then
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
        lngColor = RGB(0, 0, 255)
        If mbln��ʾƤ�� = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
            '83477:LPF,Ƥ�Խ����ȡSQL����
            strSQL = _
                " Select ʱ��, f_List2str(Cast(Collect(ҩ����) As t_Strlist)) ҩ����" & vbNewLine & _
                " From (Select To_Char(a.��ʼִ��ʱ��, 'YYYY-MM-DD') ʱ��," & vbNewLine & _
                "              Decode(Ƥ�Խ��, '(+)', 255, '(����)', 255, " & lngColor & ") || '-#' ||" & vbNewLine & _
                "               Replace(Replace(Replace(Decode(b.�Թܱ���, Null, a.ҽ������, b.�Թܱ���), ',', ''), '-#', ''), 'Ƥ��', '') || a.Ƥ�Խ�� ҩ����" & vbNewLine & _
                "       From ����ҽ����¼ a, ������ĿĿ¼ b" & vbNewLine & _
                "       Where a.������Ŀid = b.Id And a.������� = 'E' And b.�������� = '1' And a.ҽ��״̬ = 8 And a.Ƥ�Խ�� Is Not Null And a.Ƥ�Խ�� <> '����' And" & vbNewLine & _
                "             a.����id = [1] And a.��ҳid = [2] And a.Ӥ�� = [3] And a.��ʼִ��ʱ�� Between [4] And [5]" & vbNewLine & _
                "       Order By a.��ʼִ��ʱ��, a.Ƥ�Խ��)" & vbNewLine & _
                " Group By ʱ��"

            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If

            Set rsDownTab = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˹�����¼��Ϣ", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��))

            Do While Not rsDownTab.EOF
                intCOl = DateDiff("D", CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD")), CDate(Format(rsDownTab!ʱ��, "YYYY-MM-DD")))
                str��� = Nvl(rsDownTab!ҩ����)
                Call ShowTestis(str���, intCOl)
                rsDownTab.MoveNext
            Loop
            
            '��������Ƥ�Խ���и�
            For intRow = Val(mshDownTab.Tag) To mintRepairRows - 1
                If mshDownTab.RowHeight(intRow) < mlngNewHeight(intRow - Val(mshDownTab.Tag)) Then
                    mshDownTab.RowHeight(intRow) = mlngNewHeight(intRow - Val(mshDownTab.Tag))
                End If
            Next intRow
            Call picBack_Resize
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
    Dim intNum As Integer, i As Integer, j As Integer
    Dim lngColor As Long
    Dim strTmp As String, strPart As String, strSpace As String
    Dim arrTmp() As String, arrData
    Dim LPoint As T_LPoint
    Dim lngDc As Long
    Dim objDraw As Object
    Dim lngH As Long, lngW As Long, lngX1 As Long, lngLen As Long
    Dim intRowCount As Integer
    Dim sngLen As Single
    Dim intRow As Integer
    Dim sgnSize As Single, strFontName As String
    Dim lngRowHeight As Long
    
    Set objDraw = picBack
    intRowCount = Val(mshDownTab.Tag)
    intNum = 1
    strTmp = strValue
    If strTmp = "" Then Exit Sub
    LPoint.X = 0
    LPoint.W = (mshDownTab.ColWidth(mshDownTab.FixedCols) / Screen.TwipsPerPixelX) * T_BodyStyle.lng������
    lngW = LPoint.W
    lngX1 = 0
    intRow = 0
    
    '��ʼ�����Ƿ���Ҫ����
    strPart = ""
    arrTmp = Split(strTmp, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        lngColor = Val(Split(arrTmp(i), "-#")(0))
        strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") 'Ƥ�Խ��
        If Trim(strTmp) <> "" Then
            strFontName = "����"
            sgnSize = GetFontSize(objDraw, strTmp & "L", LPoint.W)
            '��С��������Ҫ�����ʵ������
            With txtLength
                .Width = LPoint.W * Screen.TwipsPerPixelX
                .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = "����"
                .FontSize = sgnSize
                .FontBold = False
                .FontItalic = False
            End With
            arrData = GetData(frmTendFileRead.txtLength.Text, txtLength)
            
            '����ÿһ��Ƥ�Խ��������и�
            If Val(objDraw.TextHeight("��") * (UBound(arrData) + 1)) > mshDownTab.RowHeightMin Then
                lngRowHeight = objDraw.TextHeight("��") * (UBound(arrData) + 1)
            Else
                lngRowHeight = mshDownTab.RowHeightMin
            End If
            If mlngNewHeight(intRow) < lngRowHeight Then mlngNewHeight(intRow) = lngRowHeight
            'Ƥ�Խ�����ڶ�����ͷ�ϲ���ʾ
            If mshDownTab.Rows > intRow + Val(mshDownTab.Tag) Then
                mshDownTab.MergeRow(intRow) = True
                strSpace = " " & String(1, " ") & String(Val(mshDownTab.Tag), " ")
                mshDownTab.TextMatrix(intRow + Val(mshDownTab.Tag), 0) = strSpace & "Ƥ�Խ��" & strSpace
                
                For j = 0 To T_BodyStyle.lng������ - 1
                    strSpace = " " & String(intCOl + 1, " ") & String(intRow + Val(mshDownTab.Tag), " ")
                    If intCOl * T_BodyStyle.lng������ + mshDownTab.FixedCols + j < mshDownTab.Cols Then
                        mshDownTab.TextMatrix(intRow + Val(mshDownTab.Tag), intCOl * T_BodyStyle.lng������ + mshDownTab.FixedCols + j) = strSpace & strTmp & strSpace
                    End If
                Next j
            End If
            '��ʼ�������
            mstrNewString(intRow, intCOl) = sgnSize & "'" & strFontName & "'" & lngColor & "-#" & strTmp
            
            intRow = intRow + 1
            intNum = intNum + 1
            If intRowCount + intNum > mintRepairRows Then Exit Sub
        End If
    Next i
End Sub

Public Sub AppenGridItemNew(ByVal rsTemp As ADODB.Recordset)
     '��д������
    Dim intRow  As Integer, intRowStart As Integer
    Dim intƵ�� As Integer
    Dim intRowNum As Integer, intColNum As Integer
    Dim intRowCount As Integer, intNum As Integer
    Dim i As Integer, j As Integer
    Dim strText As String, strֵ�� As String
    Dim int������ As Long
    Dim strArray() As String


    On Error GoTo Errhand
    int������ = T_BodyStyle.lng������
    
    
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
                        
                        'intColNum Ҫ�ϲ�������
                        'intRowNum Ҫ�ϲ�����
                        Select Case intƵ��
                            'intColNum Ҫ�ϲ�������
                            'intRowNum Ҫ�ϲ�����
                            Case 1
                                intRowNum = 1
                                intColNum = int������
                            Case 2
                                intRowNum = 1
                                intColNum = int������ / 2
                            Case 3
                                intRowNum = 3
                                intColNum = int������
                            Case 4
                                intRowNum = 2
                                intColNum = int������ / 2
                            Case 6
                                intRowNum = 1
                                intColNum = int������ / 6
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
                            
                            mshDownTab.TextMatrix(intRow, 1) = zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��ĿС��) & "," & _
                                intƵ�� & "," & zlCommFun.Nvl(!��Ŀ��ʾ) & "," & zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��Ժ�ײ�, 0) & "," & Nvl(!���²�λ)
                            mshDownTab.TextMatrix(intRow, 2) = zlCommFun.Nvl(!��Сֵ, "") & ";" & zlCommFun.Nvl(!���ֵ, "")
                            mshDownTab.TextMatrix(intRow, 3) = Nvl(!��¼��)
                            
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
    Dim lngIndex As Long
    Select Case Control.Id
        Case conMenu_View_Jump '�˵�
            mcbrToolBarҳ��.Caption = Control.Category
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("װ������", mstrParam)
            cbsMain.RecalcLayout
        Case conMenu_View_OneWeek To conMenu_View_FourWeek '4�����ڰ�ť
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("װ������", mstrParam)
            lngIndex = GetMenuPageIndex(0)
            mcbrToolBarҳ��.Caption = mcbrItem.Controls.Item(lngIndex).Category
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
    
    If T_Patient.lng�ļ�ID = 0 Then Exit Sub
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
    Dim lngDc As Long, lngFont As Long, lngOldFont As Long
    Dim objDraw As Object, stdSet As Object
    Dim intCOl As Integer, intRow As Integer
    Dim sgnSize As Single, strFontName As String
    Dim arrData
    
    On Error Resume Next
    Err = 0
    intRow = UBound(mstrNewString)
    If Err <> 0 Then Exit Sub
    
    On Error GoTo Errhand
    
    lngDc = hDC
    Set objDraw = picBack
    If mbln��ʾƤ�� = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 And Col >= mshDownTab.FixedCols And Row >= Val(mshDownTab.Tag) Then
        If (Col - mshDownTab.FixedCols) Mod T_BodyStyle.lng������ = 0 And UBound(mstrNewString) >= (Row - Val(mshDownTab.Tag)) Then
            intCOl = (Col - mshDownTab.FixedCols) / T_BodyStyle.lng������
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
            LPoint.Y = Top
            
            '1���������
            '�����뱳��ɫ��ͬ��ˢ��
            lngBackColor = GetRBGFromOLEColor(mshDownTab.BackColor)
            lngBrush = CreateSolidBrush(lngBackColor)
            'ʹ�ø�ˢ����䱳��ɫ
            lngOldBrush = SelectObject(lngDc, lngBrush)
            Call FillRect(hDC, T_ClientRect, lngBrush)
            '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
            Call SelectObject(lngDc, lngOldBrush)
            Call DeleteObject(lngBrush)
        
            sgnSize = 9: strFontName = "����"
            If UBound(Split(strTmp, "'")) > 0 Then
                sgnSize = Split(strTmp, "'")(0)
                strFontName = Split(strTmp, "'")(1)
                strTmp = Split(strTmp, "'")(2)
            End If
            
            arrTmp = Split(strTmp, "-#")
            lngColor = Val(arrTmp(0))
            strTmp = arrTmp(1)
            
            With txtLength
                .Width = mshDownTab.ColWidth(mshDownTab.FixedCols) * T_BodyStyle.lng������
                .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = "����"
                .FontSize = sgnSize
                .FontBold = False
                .FontItalic = False
            End With
            arrData = GetData(frmTendFileRead.txtLength.Text, txtLength)
            
            '��������
            Set stdSet = New StdFont
            stdSet.Name = strFontName
            stdSet.Size = sgnSize
            stdSet.Bold = False
            Call SetFontIndirect(stdSet, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            
            If (UBound(arrData) + 1) * objDraw.TextHeight("��") < mshDownTab.RowHeight(Row) Then
                LPoint.Y = Top + (Val(mshDownTab.RowHeight(Row)) - ((UBound(arrData) + 1) * objDraw.TextHeight("��"))) / T_TwipsPerPixel.Y / 2
            Else
                LPoint.Y = Top
            End If
            
            '��ʼ�������
            Call SetTextColor(lngDc, lngColor)
            For i = 0 To UBound(arrData)
                Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(arrData(i)), , False)
                Call DrawText(lngDc, CStr(arrData(i)), -1, T_LableRect, DT_CENTER)
                LPoint.Y = LPoint.Y + Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0")
            Next i
                    
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        End If
    End If
    
    '���������
    If Col >= mshDownTab.FixedCols And Row >= mshDownTab.FixedRows Then
        strTmp = mshDownTab.TextMatrix(Row, Col)
        If AnsyGrade(Val(mshDownTab.RowData(Row)), strTmp, arrText) = True Then
            'lngColor = mshDownTab.Cell(flexcpForeColor, Row, Col, Row, Col)
            Call DrawDownTabAnsyGrade(lngDc, picMain, arrText, Row, Col, Left, Top, Right, Bottom, Done, mbln�೦�����ӷ�ĸ��ʾ)
        End If
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mshDownTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RaiseShowTipInfo(mshDownTab, 3, X, Y)
End Sub

Private Sub mshUpTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strTime As String
    If NewRow = 0 And T_Patient.lng�༭ = 1 Then
        strTime = GetCurveDateNew(NewCol, mstr��ʼʱ��, gintHourBegin)
        If Format(Split(strTime, ";")(0), "YYYY-MM-DD") > Format(mstr����ʱ��, "YYYY-MM-DD") Then
            mshUpTab.FocusRect = flexFocusLight
        Else
            mshUpTab.FocusRect = flexFocusSolid
            If mblnKeyDown = True Then
                picDisplay.Left = ((((NewCol - 1) \ T_BodyStyle.lng������) + 1) * T_BodyStyle.lng������ - 1) * mshUpTab.ColWidth(NewCol) + mshUpTab.ColWidth(0)
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
    Dim intMinCol As Integer, intMaxCol As Integer
    Dim i As Integer, j As Integer
    Dim strTmp As String
    Dim lngColor As Long, lngDc As Long
    Dim objDraw As Object, stdSet As Object
    Dim lngƵ�� As Long
    Dim lngʱ���� As Long
    
    lngDc = hDC
    
    lngƵ�� = T_BodyStyle.lng������
    lngʱ���� = T_BodyStyle.lngʱ����
    
    If picMain.Tag = "" Then Exit Sub
    If Row = mshUpTab.Rows - 1 And Col >= mshUpTab.FixedCols Then
        Set objDraw = picBack
        Call CalcMinMaxColNew(picMain.Tag, intMinCol, intMaxCol)
        j = (Col - mshUpTab.FixedCols) Mod lngƵ��
        
        strTmp = gintHourBegin + lngʱ���� * j

        '���ݲ�������ҹ��ʱ�䷶Χ����ʱ����ɫ
        lngColor = GetTimeColor(Val(strTmp))
        If Col >= intMinCol And Col <= intMaxCol Then
            lngColor = lngColor
        Else
            lngColor = RGB_FleetGRAY
        End If
        
        Call SetTextColor(lngDc, lngColor)
        Call GetTextRect(objDraw, Left, Top + (Bottom - Top) / 2, CStr(strTmp), Right - Left - 3, True)
        Call DrawText(lngDc, CStr(strTmp), -1, T_LableRect, DT_CENTER)
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

Private Sub mshUpTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RaiseShowTipInfo(mshUpTab, 1, X, Y)
End Sub

Private Sub RaiseShowTipInfo(ByVal vfgObj As Object, ByVal intType As Byte, ByVal X As Single, ByVal Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim lngHeadWidth As Long
    Dim lngWidth As Long, lngHeight As Long, lngHeight1 As Long
    Dim i As Long
    
    If Not vfgObj.Visible Then Exit Sub
    Select Case intType
    Case 1 '�ϱ��
        lngHeadWidth = vfgObj.ColWidth(0)
        lngWidth = vfgObj.ColWidth(vfgObj.FixedCols)
        lngHeight = vfgObj.RowHeight(vfgObj.FixedRows)
    Case 3 '�±��
        lngHeadWidth = vfgObj.ColWidth(0)
        lngWidth = vfgObj.ColWidth(vfgObj.FixedCols)
        lngHeight = vfgObj.RowHeight(vfgObj.FixedRows)
    Case 2 '�������
        lngHeadWidth = vfgObj.ColWidth(1)
        lngWidth = vfgObj.ColWidth(vfgObj.FixedCols)
        lngHeight = vfgObj.RowHeight(vfgObj.FixedRows)
    Case Else
        Exit Sub
    End Select
    
    lngHeight = 0
    lngHeight1 = 0
    For i = 0 To vfgObj.Rows - 1
        If vfgObj.RowHidden(i) = False Then
            lngHeight = lngHeight + vfgObj.RowHeight(i)
            If Y > lngHeight1 And Y < lngHeight Then Exit For
            lngHeight1 = lngHeight
        End If
    Next i
    
    If i < vfgObj.Rows Then
        lngRow = i
    Else
        Exit Sub
    End If
    
    If X <= lngHeadWidth Then
        lngCol = IIf(intType = 2, 1, 0)
    Else
        lngCol = (X - lngHeadWidth) \ lngWidth + vfgObj.FixedCols
    End If
    If lngRow >= 0 And lngCol >= 0 And lngRow < vfgObj.Rows - IIf(intType = 1, 1, 0) And lngCol < vfgObj.Cols Then
        RaiseEvent ShowTipInfo(vfgObj, vfgObj.TextMatrix(lngRow, lngCol), True)
    Else
        RaiseEvent ShowTipInfo(vfgObj, "", True)
    End If
End Sub

Private Sub picBack_Resize()
    Dim lngLeft As Long
    Dim lngHeight As Long, lngRow As Long
    
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
        .ColWidth(0) = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) * Screen.TwipsPerPixelX
        .Left = lngLeft
        .Top = picCard(0).Top + picCard(0).Height
        .Height = .Rows * mshUpTab.RowHeight(0)
        .Width = ((T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) + T_DrawClient.�е�λ * T_BodyStyle.lng������ * T_BodyStyle.lng���� + 1) * T_TwipsPerPixel.X
        .ColWidthMin = T_DrawClient.�е�λ * Screen.TwipsPerPixelX
         picCard(0).Width = .Width
         .Refresh
    End With
    
    picDraw.Move 0, mshUpTab.Top + mshUpTab.Height, (T_DrawClient.��������.Right + 1) * T_TwipsPerPixel.X, _
        (T_DrawClient.����������.Bottom - T_DrawClient.����������.Top + 1) * Screen.TwipsPerPixelY

    picDisplay.Height = 165
     
    With vsf
        .Top = mshUpTab.Top + mshUpTab.Height + (T_DrawClient.����������.Bottom - T_DrawClient.����������.Top + 1) * Screen.TwipsPerPixelY
        .Left = lngLeft
        .Width = mshUpTab.Width
        .Height = .Body.RowHeight(vsf.FixedRows)
        .Visible = Not mbln��������
    End With
        
    With mshDownTab
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .Left = lngLeft
        .Top = mshUpTab.Top + mshUpTab.Height + (IIf(mbln�������� = False, vsf.Height, 0)) + (T_DrawClient.����������.Bottom - T_DrawClient.����������.Top + 1) * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        .Width = mshUpTab.Width
        lngHeight = 0
        For lngRow = 0 To .Rows - 1
            lngHeight = lngHeight + .RowHeight(lngRow)
        Next lngRow
        .Height = lngHeight
        .Refresh
    End With
    
    picCommText.Left = lngLeft
    picCommText.Top = mshDownTab.Top + mshDownTab.Height
    picCommText.Width = mshDownTab.Width
    picCommText.Visible = True
    
    mshUpTab.Redraw = True
    mshDownTab.Redraw = True
    
    picMain.Width = mshUpTab.Width + mshUpTab.Left
    picMain.Height = picCommText.Top + picCommText.Height
    
    '���������
    Call CalcScrollBarSize
    
    '�������µ��Ŀɻ������С
    mlng�߶� = (picBack.Height - mshUpTab.Top - mshUpTab.Height - mshDownTab.Height - picCommText.Height - _
        IIf(mbln�������� = False, vsf.Height, 0) - IIf(hsb.Max > 0, hsb.Height, 0)) / Screen.TwipsPerPixelY
    
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

    
    hsb.Max = 0 - Int(0 - ((picMain.Width - picBack.Width) / 100)) - 1
    vsb.Max = 0 - Int(0 - ((picMain.Height - picBack.Height) / 100)) - 1
    If vsb.Max > 0 Then
        hsb.Max = 0 - Int(0 - ((picMain.Width - picBack.Width + vsb.Width) / 100)) - 1
    End If
    If hsb.Max > 0 Then
        vsb.Max = 0 - Int(0 - ((picMain.Height - picBack.Height + hsb.Height) / 100)) - 1
    End If
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    If hsb.Visible = True Then hsb.ZOrder 0
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    If vsb.Visible = True Then vsb.ZOrder 0
    
    With vsb
        .Height = picBack.Height
    End With
    
    With hsb
        .Width = picBack.Width - IIf(vsb.Max > 0, vsb.Width, 0)
    End With
    
    'ֻ����û��ʾ�������ǲ��������㲽��
    msinHStep = (picMain.Width - picBack.Width + IIf(vsb.Max > 0, vsb.Width, 0)) / 10
    msinVStep = (picMain.Height - picBack.Height + IIf(hsb.Max > 0, hsb.Height, 0)) / 10
    
    '�㶨Ϊ100,ֻ�ǲ��������仯
    If hsb.Enabled Then
        hsb.Max = 10
        hsb.LargeChange = 10 / Int((Round((picMain.Width - picBack.Width + IIf(vsb.Max > 0, vsb.Width, 0)) / picBack.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange
    End If
    
    If vsb.Enabled Then
        vsb.Max = 10
        vsb.LargeChange = 10 / Int((Round((picMain.Height - picBack.Height + IIf(hsb.Max > 0, hsb.Height, 0)) / picBack.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange
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
    Dim intMinCol As Integer
    Dim intMaxCol As Integer
    Dim intCOl As Integer
    If Button <> vbLeftButton Then Exit Sub
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    '����ָ��������ſɽ��в���
    If T_Patient.lng�༭ = 1 Then
        intCOl = ((mintColMax - 1) \ T_BodyStyle.lng������ + 1) * T_BodyStyle.lng������
        
        If X > mshUpTab.ColWidth(0) And X < mshUpTab.ColWidth(0) + (intCOl * mshUpTab.ColWidth(intCOl)) Then
            '�������꣬������������
            strTemp = GetXCoordinateNew(X / T_TwipsPerPixel.X + mshUpTab.Left / T_TwipsPerPixel.X - 1, mstr��ʼʱ��, False)
            strTemp = mstr��ʼʱ�� & ";" & Split(strTemp, ",")(1)
            '����ʱ�������
            Call CalcMinMaxColNew(strTemp, intMinCol, intMaxCol)
            picDisplay.Visible = True
            If Y < mshUpTab.RowHeight(0) + 40 Then
                picDisplay.Left = ((((intMaxCol - 1) \ T_BodyStyle.lng������) + 1) * T_BodyStyle.lng������ - 1) * mshUpTab.ColWidth(intMaxCol) + mshUpTab.ColWidth(0)
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
    '��ȡ�����ļ��б�
    mstrSQL = "Select A.ID,A.�ļ����� From ���˻����ļ� A,�����ļ��б� B" & _
       "    where A.����ID=[1] and A.��ҳId=[2] and nvl(A.Ӥ��,0)=[3] and A.��ʽID=B.ID and B.����=3 and B.����=-1 Order by A.��ʼʱ��"
    If mblnMoved = True Then
        mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�ļ�ID", T_Patient.lng����ID, T_Patient.lng��ҳID, T_Patient.lngӤ��)
    cboFile.Clear
    With RS
        Do While Not .EOF
            cboFile.AddItem Nvl(!�ļ�����)
            cboFile.ItemData(cboFile.NewIndex) = !Id
        .MoveNext
        Loop
    End With
    
    If cboFile.ListCount > 1 Then
        cboFile.Enabled = True
    Else
        cboFile.Enabled = False
    End If
    
    If cboFile.ListCount > 0 And cboFile.ListIndex = -1 Then cboFile.ListIndex = 0
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboFile_Click()
    Dim strParam As String
    
    If T_Patient.lng�ļ�ID = cboFile.ItemData(cboFile.ListIndex) Then Exit Sub
    T_Patient.lng�ļ�ID = cboFile.ItemData(cboFile.ListIndex)
    If mblnAutoAdjust = False Then '����ģʽ
        '��ȡ��ʼ���¸�ʽ��������
        strParam = T_Patient.lng����ID & ";" & T_Patient.lng��ҳID & ";" & T_Patient.lng����ID & ";" & T_Patient.lng�ļ�ID & ";" & _
        T_Patient.lng��Ժ & ";" & T_Patient.lng�༭ & ";" & T_Patient.lngӤ�� & ";" & T_Patient.lng����ȼ� & ";1"
        Call zlMenuClick("��ʼ��", strParam)
    Else
        RaiseEvent zlFileChange(True, T_Patient.lng�ļ�ID, T_Patient.lngӤ��)
    End If
End Sub

Private Sub picDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIndex As Long
    If KeyCode = vbKeyRight And Shift = vbCtrlMask Then  '��һ��
        If mintPage < mintAllPage - 1 Then
            lngIndex = GetMenuPageIndex(1)
            mstrParam = mcbrItem.Controls.Item(lngIndex).Parameter '�õ���ǰҳ��ʱ��
            Call InitWeekDays(mstrParam)
            mcbrToolBarҳ��.Caption = mcbrItem.Controls.Item(lngIndex).Category
            cbsMain.RecalcLayout
            Call zlMenuClick("װ������", mstrParam)
        End If

    ElseIf KeyCode = vbKeyLeft And Shift = vbCtrlMask Then
        If mintPage > 0 Then '��һ��
            lngIndex = GetMenuPageIndex(-1)
            mstrParam = mcbrItem.Controls.Item(lngIndex).Parameter '�õ���ǰҳ��ʱ��
            Call InitWeekDays(mstrParam)
            mcbrToolBarҳ��.Caption = mcbrItem.Controls.Item(lngIndex).Category
            cbsMain.RecalcLayout
            Call zlMenuClick("װ������", mstrParam)
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        'mblnAutoRedraw = mblnAutoRedraw Xor True
    End If
End Sub

Private Function GetMenuPageIndex(ByVal intType As Integer) As Long
    '����:��ȡ���µ�ҳ���Ӧ�Ĳ˵�����
    'intType:��Ե�ǰҳҪ��ת��ҳ��
    '72090:������,2014-07-23
    Dim i As Long, lngIndex As Long, lngPage As Long
    
    lngPage = mintPage + intType
    If lngPage < 0 Then
        lngPage = 0
    ElseIf lngPage > mintAllPage - 1 Then
        lngPage = mintAllPage - 1
    End If
    
    For i = 1 To mcbrItem.Controls.Count
        If Val(Split(mcbrItem.Controls.Item(i).Parameter, ";")(4)) = lngPage Then
            lngIndex = i
            Exit For
        End If
    Next i
    
    GetMenuPageIndex = lngIndex
End Function

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '82732:LPF,����ƶ�����Ӧ�����ݵ㣬��ʾ������Ϣ
    Dim sgnX As Single, sgnY As Single, sgnTmp As Single
    Dim strInfo As String, strTmp As String
    Dim colPonit As Collection
    Dim arrPoint(0 To 2) As String, i As Integer
    
    If mrsPoint Is Nothing Then Exit Sub
    If mrsPoint.State = adStateClosed Then Exit Sub
    
    sgnX = picDraw.ScaleX(X, vbTwips, vbPixels)
    sgnY = picDraw.ScaleX(Y, vbTwips, vbPixels)
    If sgnX >= T_DrawClient.��������.Left And sgnX <= T_DrawClient.����������.Right And sgnY >= T_DrawClient.����������.Top And sgnY <= T_DrawClient.����������.Bottom Then
        '������ռ�¼��mrsPoint�е���������λ�����������ƶ���׼ȷ�ĵ�(�����Բ���),��˲�ȡ����Χ��λ
        '1���������λ�����¼����Ӧ���ʵ��X����
        strTmp = GetXCoordinateNew(sgnX, Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"), False)
        sgnTmp = GetXCoordinateNew(Format(Split(strTmp, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
        mrsPoint.Filter = "X����=" & sgnTmp
        mrsPoint.Sort = "��Ŀ���"
        '2��ѭ��ʵ��X�����Ӧ�����ݣ������������λ�ö�Ӧ������(�������Ҹ��ƶ�4���㣬������긡����ʾ)
        Set colPonit = New Collection
        Do While Not mrsPoint.EOF
            sgnTmp = Val(mrsPoint!X����) + T_DrawClient.�е�λ \ 2
            If Val(mrsPoint!Y����) > sgnY - 4 And Val(mrsPoint!Y����) < sgnY + 4 And sgnTmp > sgnX - 4 And sgnTmp < sgnX + 4 Then
                arrPoint(0) = Val(mrsPoint!��Ŀ���)
                arrPoint(1) = Nvl(mrsPoint!��λ)
                arrPoint(2) = Nvl(mrsPoint!��ֵ)
                colPonit.Add arrPoint
            End If
            mrsPoint.MoveNext
        Loop
        mrsPoint.Filter = ""
        '3.��������������ʽ:��Ŀ����[(��λ)]����ֵ[(��Ŀ��λ)]
        For i = 1 To colPonit.Count
            mrsItems.Filter = "��Ŀ��� =" & Val(colPonit.Item(i)(0))
            If mrsItems.RecordCount > 0 Then
                strInfo = IIf(strInfo = "", "", strInfo & vbCrLf) & mrsItems!��Ŀ���� & IIf(colPonit.Item(i)(1) = "", "", "(" & colPonit.Item(i)(1) & ")") & "��" & colPonit.Item(i)(2) & "" & IIf(IsNumeric(colPonit.Item(i)(2)) = True, Nvl(mrsItems!��Ŀ��λ), "")
            End If
        Next i
        
        RaiseEvent ShowTipInfo(picDraw, strInfo, True)
    Else
        RaiseEvent ShowTipInfo(picDraw, "", False)
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
    T_DrawClient.�е�λ = T_BodyStyle.lng�����п� \ Screen.TwipsPerPixelX
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
    
    lblSerach(0).FontSize = bytFontSize
    cboFile.FontSize = bytFontSize
    cboFile.Top = (picTmp.Height - cboFile.Height) \ 2
    lblSerach(0).Top = cboFile.Top + (cboFile.Height - lblSerach(0).Height) \ 2
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
   
    If UserControl.Parent.Visible = False Then Exit Sub
    If mblnAutoAdjust = True And Not mblnResize Then
        '���ʵ�ʴ�С�Ƿ����仯
        If Abs(mlngHeight - UserControl.Height) > 20 Then
            'Debug.Print "--��С�ı����--"
            Call LockWindowUpdate(UserControl.hWnd)
            Call zlMenuClick("װ������", mstrParam)
            Call LockWindowUpdate(0)
            mblnResize = True
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
    T_ClientRect.Right = T_ClientRect.Right * 2
    T_ClientRect.Bottom = T_ClientRect.Bottom * 2
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

Public Sub Paint_CanvasNew(Optional ByVal blnAdjust As Boolean = False)
    '׼����������ɿ̶ȼ�������š����ⵥ�������Լ������趨���л�׼�ߵ���棩
    '��Сģʽ��,����ʾ���ϱ��,�ı�������
    'blnAdjust=False��ʾ�̶���С�����������������е���
    
    Static SlngMaxY As Long                 '��¼��һ�ε����߶ȣ��Ծ��������Ƿ���Ҫ�ػ�
    Dim lngCurX     As Long, lngCurY As Long   '��ǰλ��
    Dim lngMaxX     As Long, lngMaxY As Long, lngAllMaxY As Long  '�߽�
    Dim lngCurAlerY As Long
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
    Dim lngCurveRows As Long '��������������
    Dim lngY As Long, lngX As Long
    Dim str˵�� As String
    Dim sinCurAlerY As Single
    '�������ͼ�������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    Dim sin�̶� As Single, bln��ʾ�̶� As Boolean, blnFirst As Boolean
    Dim sin�̶ȼ�� As Single, sinBegin�̶� As Single, dbl��λֵ As Double

    Dim str���ֵ���� As String, str��Сֵ���� As String
    Dim lng�̶ȿ�� As Long
    

    On Error GoTo Errhand
    
    'ʵ�����ŵ�ԭ��˵����
    '1����ͨģʽ���������ݾ���ʾ
    '2����Сģʽ=2��ʱ��̶Ȳ���ʾ��ÿ��10С�и�Ϊ5С��
    '3����Сģʽ<=4��תΪ������ʾ
    
    '��ǰ�ǹ̶���������2����������ݣ����Դ˴���ȥ2��
    '�������ʱΪ�˶���ÿ����ٴμ�2�������
    lngCurveRow = T_BodyStyle.lng���߿���
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    T_DrawClient.������ = glngMaxRows
    
    gstrSQL = " Select /*+ Rule*/ A.��Ŀ���,A.�������,A.��¼��,A.��¼��,A.��¼ɫ,nvl(A.���ֵ,0) ���ֵ,nvl(A.��Сֵ,0) ��Сֵ," & _
        "nvl(A.��λֵ,0) ��λֵ,A.�̶ȼ��,A.��ʾ��,C.��Ŀ��λ ��λ,Decode(��¼��,3,A.�����,nvl(A.�����,2)-2) AS �����,B.��λ,A.��¼��" & _
        " From ���¼�¼��Ŀ A,���²�λ B,�����¼��Ŀ C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) D" & _
        " Where A.��Ŀ���=B.��Ŀ���(+) And B.ȱʡ��(+)=1" & _
        " And  A.��Ŀ���=C.��Ŀ��� AND A.��¼��<>2 AND NOT (NVL(C.Ӧ�÷�ʽ,0)=2 And C.��Ŀ���=-1) and C.��Ŀ���=D.COLUMN_VALUE" & _
        " Order by �������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ʼ��", T_BodyItem.str������Ŀ)
    
    '------------------------------------------------------------------------------------------------------------------
    rsTemp.Filter = "��Ŀ���=" & gint����
    '�����ӡ���������
    With rsTemp
        Do While Not .EOF
            lng����� = Val(zlCommFun.Nvl(!�����))
            If lng����� < 0 Then lng����� = 0
            
             '�޸�����51442
            If Val(zlCommFun.Nvl(!��Сֵ, 0)) > 34 Then
                lngMaxRows = lng����� + (Val(zlCommFun.Nvl(!���ֵ, 0)) - 35) / 0.1
            Else
                lngMaxRows = lng����� + (Val(zlCommFun.Nvl(!���ֵ, 0)) - Val(zlCommFun.Nvl(!��Сֵ, 0))) / 0.1
            End If

            lngMaxRows = lngMaxRows + lngCurveRow
            T_DrawClient.������ = lngMaxRows
        .MoveNext
        Loop
    End With
    
    T_DrawClient.�������������� = 0
    rsTemp.Filter = "��¼��=3 And ��Ŀ���<>1"
    rsTemp.Sort = "�������"
    Do While Not rsTemp.EOF
        lngRow = ((Val(Nvl(rsTemp!���ֵ, 0)) - Val(Nvl(rsTemp!��Сֵ, 0))) / Val(Nvl(rsTemp!��λֵ, 1)))
        If Val(Nvl(rsTemp!�����, 0)) > 0 Then lngRow = lngRow + Val(Nvl(rsTemp!�����, 0))
        If lngRow Mod 2 = 1 Then lngRow = lngRow + 1
        T_DrawClient.�������������� = T_DrawClient.�������������� + lngRow
    rsTemp.MoveNext
    Loop
    
    rsTemp.Filter = "��¼��=1"
    rsTemp.Sort = "�������"
    intLables = rsTemp.RecordCount
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
    '����ֵ
    intLineMode = PS_SOLID
    
    lngColStep = T_BodyStyle.lng�����п� \ Screen.TwipsPerPixelX
    lngInitRowStep = T_BodyStyle.lng�����и� \ Screen.TwipsPerPixelY
    sigRowStepNew = lngInitRowStep
    lng�̶ȿ�� = (T_BodyStyle.lng�̶ȿ�� \ Screen.TwipsPerPixelX)
    lngLableStep = Fix(lng�̶ȿ�� / intLables)
    intTens_digit = 3
    '���µ��Ե�����ʾ(������ѡ����˫����ʾ��û�����̶���ʾһ��) 1��������ʾ 0��˫����ʾ
    If zlDatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, 0) = 1 Then
        bln˫�� = False
    Else
        bln˫�� = True
    End If
    'True��ʾ����ֻ���һ��,Ч����һ���̶�ֻ��ʾ������;����һ���̶���ʾʮ��,���û�������������,��blnDoubleRow�޹�
    bln���� = True
    
    If Not bln���� Then intLineMode = PS_DASHDOTDOT
    
    '�����
    lngCurX = T_DrawClient.ƫ����X
    lngCurY = T_DrawClient.ƫ����Y
    lngMaxX = lng�̶ȿ�� + (T_BodyStyle.lng���� * T_BodyStyle.lng������ * lngColStep) + T_DrawClient.ƫ����X   '�̶�+������*��� +T_DrawClient.ƫ����X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.������ * sigRowStepNew + T_DrawClient.ƫ����Y '�Ƕ������߲������Y����
    lngAllMaxY = 2 * mintNullRow * lngInitRowStep + (T_DrawClient.������ + T_DrawClient.��������������) * sigRowStepNew + T_DrawClient.ƫ����Y '�������߲������Y����
    '����������ݵ�У��
    If blnAdjust Then
        '���С�ڿɼ������С���������
        If lngAllMaxY > mlng�߶� Then
            lngAllMaxY = mlng�߶� - 2 * mintNullRow * lngInitRowStep
            sigRowStepNew = Round((lngAllMaxY) / (T_DrawClient.������ + T_DrawClient.��������������), 1)
            sigRowStepNew = Fix(sigRowStepNew + 0.5)
        End If

        '����и�̫С���򽫷�����Ϊһ����ʾ
        If sigRowStepNew <= 2 Then
            sinRowStep = 2
            blnDoubleRow = True
        End If

        If Not mblnRedraw Then mblnRedraw = (lngAllMaxY <> SlngMaxY)
        If sigRowStepNew < 4 Then intLineMode = PS_DOT
    End If
    '����̶ȵ��������
    lngMaxY = (lngInitRowStep * 2 * mintNullRow) + T_DrawClient.������ * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + T_DrawClient.ƫ����Y
    lngAllMaxY = (lngInitRowStep * 2 * mintNullRow) + (T_DrawClient.������ + T_DrawClient.��������������) * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + T_DrawClient.ƫ����Y
    
    Call Paint_Reset                                                    '�������
    
    SlngMaxY = lngMaxY
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�е�λ = IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.˫�� = blnDoubleRow
    
    For lngRow = 1 To intLables
        Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
        '���ڿ��ܴ��ڿ̶��ܿ��/��Ŀ����������(�磺90/4),����ʽΪǰ3��ΪFix(90/4),���һ�еĿ��Ϊ�̶ȿ��-ǰ3�еĿ��
        '��֤���ϱ����ͷ�����±����ͷ�Ϳ̶ȿ����ͬ
        If lngRow = intLables Then
            lngCurX = lng�̶ȿ�� + T_DrawClient.ƫ����X
        Else
            lngCurX = lngCurX + lngLableStep
        End If
    Next
    Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
    
    '���̶ȿ�
    Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)

    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    T_DrawClient.�̶�����.Top = lngCurY
    T_DrawClient.�̶�����.Right = lng�̶ȿ�� + T_DrawClient.ƫ����X
    T_DrawClient.�̶�����.Bottom = lngMaxY
    
    'Ĭ�����һ��������ʾ��Ŀ����
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    
    '�����µ�������
    For lngRow = 0 To T_DrawClient.������ - 1
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
    For lngRow = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
        lngCurX = lngCurX + lngColStep
        
        Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng������ = 0, 2, 1), IIf(lngRow Mod T_BodyStyle.lng������ = 0, RGB_RED, RGB_GRAY))
    Next
    
    lngCurX = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.��������.Right = lngMaxX
    T_DrawClient.��������.Bottom = lngMaxY
    
    T_DrawClient.����������.Left = T_DrawClient.�̶�����.Left
    T_DrawClient.����������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.����������.Right = lngMaxX
    T_DrawClient.����������.Bottom = lngAllMaxY
    
    '�������������
    Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
    
    Set mobjPart = New Collection
    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
    rsTemp.Filter = "��¼��=1"
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            '��ʾ�̶ȿ���Ŀ�����Ƽ�����,�����¡�
            lngCurX = T_DrawClient.�̶�����.Left + ((.AbsolutePosition - 1) * T_DrawClient.�̶ȵ�λ)
            If .AbsolutePosition = .RecordCount Then
                lngLableStep = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) - ((.AbsolutePosition - 1) * T_DrawClient.�̶ȵ�λ)
            Else
                lngLableStep = T_DrawClient.�̶ȵ�λ
            End If
            lngCurY = T_DrawClient.�̶�����.Top
            
            '���������С
            Set gstdSet = New StdFont
            gstdSet.Name = "����"
            gstdSet.Size = 9
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)
            
            '���������Ŀ������
            Call SetTextColor(mlngMemDC, zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + mobjDraw.TextHeight(zlCommFun.Nvl(!��¼��)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!��¼��)), lngLableStep)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!��¼��)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            Call ReleaseFontIndirect(mobjDraw)
            '���������С
            Set gstdSet = New StdFont
            gstdSet.Name = "����"
            gstdSet.Size = 8
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)

            '�����Ŀ��λ
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + lngInitRowStep * 2 + mobjDraw.TextHeight(zlCommFun.Nvl(!��λ)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!��λ)), lngLableStep)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!��λ)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            Call ReleaseFontIndirect(mobjDraw)
            sinY��λ = T_LableRect.Bottom
            '��������
            Set gstdSet = New StdFont
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

                Case gint����, gint����  '����/������10�ı�������̶�
                    intTens_digit = 3
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 10)
                    dbl��λֵ = 2
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)

                Case gint����  '������5�ı�������̶�
                    mbln�������� = True
                    intTens_digit = 2
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 5)
                    dbl��λֵ = 1
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                Case Else
                    intTens_digit = 1
                    dbl��λֵ = Val(zlCommFun.Nvl(!��λֵ, 1))
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, Val(zlCommFun.Nvl(!��λֵ, 0)) * 10)
                    If sin�̶ȼ�� > Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ)) Then
                        sin�̶ȼ�� = Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ))
                    End If
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
            End Select
            
            If !��Ŀ��� = gint���� Then
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "����", arrTemp(0), Nvl(!��¼ɫ, RGB_BLACK), "B"), "B" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "Ҹ��", arrTemp(1), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "����", arrTemp(2), Nvl(!��¼ɫ, RGB_BLACK), "C"), "C" & !��Ŀ���
            ElseIf !��Ŀ��� = gint���� Then
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "ȱʡ��¼��", Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "����", "H", RGB_RED, "B"), "B" & !��Ŀ���
                If mint����Ӧ�� = 2 Then
                    mrsItems.Filter = "��Ŀ���=" & gint����
                    If mrsItems.RecordCount > 0 Then
                        mobjPart.Add Array("" & gint����, Nvl(mrsItems!��Ŀ����), "", Nvl(mrsItems!��¼��), RGB_RED, "A"), "A" & gint����
                    End If
                    mrsItems.Filter = ""
                End If
            ElseIf !��Ŀ��� = gint���� Then
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "��������", Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "������", "R", RGB_BLACK, "B"), "B" & !��Ŀ���
            Else
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "", Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
            End If
            
            '����ֵ
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow)   '�̶�ǰ2 * mintNullRow�еĸ߶Ȳ�����̶�

            '�������Сģʽ,�ӵ�30�п�ʼ�����ʶ
            'If blnDoubleRow Then lngCurY = lngCurY + lngInitRowStep * 2 * mintNullRow
            
            '��������ж�λ����Чλ��
            lngCurY = lngCurY + (T_DrawClient.�е�λ * zlCommFun.Nvl(!�����, 2))
            blnFirst = False
            Do While True
                bln��ʾ�̶� = False
                If blnFirst = False Then    '�ս���ѭ������ʱȡ�����ֵ
                    sin�̶� = zlCommFun.Nvl(!���ֵ, 0)
                    sinBegin�̶� = sin�̶�
                    str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                    blnFirst = True
                Else                    '����õ�ÿ���̶ȵ�ֵ
                    sin�̶� = sin�̶� - dbl��λֵ    '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                End If
                
                If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                
                If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - IIf(T_DrawClient.˫��, sin�̶ȼ�� * 2, sin�̶ȼ��)
                
                If sinBegin�̶� < Val(Format(!��Сֵ, "#0.00")) Then sinBegin�̶� = Val(Format(!��Сֵ, "#0.00"))
                
                If bln��ʾ�̶� Then
                    '�������ֵ�������ߵ�λ�ظ�
                    If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY < sinY��λ Then
                        Call GetTextRect(mobjDraw, lngCurX, sinY��λ, Format(sin�̶�, "#0"), lngLableStep)
                    ElseIf lngCurY = T_DrawClient.�̶�����.Bottom Then
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY - (mobjDraw.TextHeight("1") / (T_TwipsPerPixel.Y * 2)), Format(sin�̶�, "#0"), lngLableStep)
                    Else
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY, Format(sin�̶�, "#0"), lngLableStep)
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
                    If (sinAlertness < Val(Nvl(!���ֵ)) And sinAlertness > Val(Nvl(!��Сֵ))) Then
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
            Call ReleaseFontIndirect(mobjDraw)
            sinBegin�̶� = 0
            sin�̶� = 0                 '���ƴӵ�һ�п�ʼ���
            .MoveNext
        Loop
    End With

    '��ɶ������߲��ֵ����
    lngMaxY = T_DrawClient.�̶�����.Bottom
    rsTemp.Filter = "��¼��=3"
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            lngY = lngMaxY
            lngCurY = lngY
            lngCurX = T_DrawClient.ƫ����X
            lngCurveRows = ((Val(Nvl(!���ֵ, 0)) - Val(Nvl(!��Сֵ, 0))) / Val(Nvl(!��λֵ)))
            If Val(Nvl(!�����, 0)) > 0 Then lngCurveRows = lngCurveRows + Val(Nvl(!�����, 0))
            If lngCurveRows Mod 2 = 1 Then lngCurveRows = lngCurveRows + 1
            If lngCurveRows > 0 Then
                lngMaxY = lngCurveRows * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + lngCurY
                '��ɿ̶�����Ļ���
                Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                Call DrawLine(mlngMemDC, lngCurX + lng�̶ȿ��, lngCurY, lngCurX + lng�̶ȿ��, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                Call DrawLine(mlngMemDC, lngCurX, lngMaxY, lngCurX + lng�̶ȿ��, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                '��������еĻ���
                lngCurX = lngCurX + lng�̶ȿ��
                For lngRow = 1 To lngCurveRows
                    '�����µ���������
                    If lngRow <> 0 Then
                        lngCurY = lngCurY + IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
                    End If

                    If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
                        Call DrawLine(mlngMemDC, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sigRowStepNew >= 4 And bln����, 2, 1), RGB_FleetGRAY)
                    End If
                Next
                lngCurY = lngY

                 '�����µ�������
                For lngRow = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
                    lngCurX = lngCurX + lngColStep
                    Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng������ = 0, 2, 1), IIf(lngRow Mod T_BodyStyle.lng������ = 0, RGB_RED, RGB_GRAY))
                Next
                
                '�������������
                Call DrawLine(mlngMemDC, T_DrawClient.ƫ����X, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                '�����Ŀ���ƺͿ̶ȵ����
                lngCurX = lngX: lngCurY = lngY
                '���������Ŀ������
                '��������
                Set gstdSet = New StdFont
                gstdSet.Name = "����"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
                mlngFont = CreateFontIndirect(T_Font)
                mlngOldFont = SelectObject(mlngMemDC, mlngFont)
                Call SetTextColor(mlngMemDC, Nvl(!��¼ɫ, RGB_BLACK))
                T_Size.H = mobjDraw.ScaleY(mobjDraw.TextHeight("��"), vbTwips, vbPixels)
                If T_Size.H * Len(Nvl(!��¼��)) >= lngCurveRows * T_DrawClient.�е�λ Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * T_DrawClient.�е�λ) - (T_Size.H * Len(Nvl(!��¼��)))) \ 2
                End If
                For lngRow = 1 To Len(Nvl(!��¼��))
                    Call GetTextRect(mobjDraw, lngCurX, lngCurY, Mid(Nvl(!��¼��), lngRow, 1), lng�̶ȿ�� \ 2, False)
                    Call DrawText(mlngMemDC, Mid(Nvl(!��¼��), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call SelectObject(mlngMemDC, mlngOldFont)
                Call DeleteObject(mlngFont)
                Call ReleaseFontIndirect(mobjDraw)
                '�����Ŀ��λ
                lngCurY = lngY: If Nvl(!��¼��) <> "" Then lngCurX = T_LableRect.Right
                If Trim(Nvl(!��λ)) <> "" And Nvl(!��¼��) <> "" Then
                    '���������С
                    Set gstdSet = New StdFont
                    gstdSet.Name = "����"
                    gstdSet.Size = 8
                    Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
                    mlngFont = CreateFontIndirect(T_Font)
                    mlngOldFont = SelectObject(mlngMemDC, mlngFont)
                    T_Size.H = mobjDraw.ScaleY(mobjDraw.TextHeight("��"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(Nvl(!��λ))) >= lngCurveRows * T_DrawClient.�е�λ Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * T_DrawClient.�е�λ) - (T_Size.H * Len(Nvl(!��λ)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(Nvl(!��λ)))
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY, Mid(Trim(Nvl(!��λ)), lngRow, 1), 0, False)
                        Call DrawText(mlngMemDC, Mid(Trim(Nvl(!��λ)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call SelectObject(mlngMemDC, mlngOldFont)
                    Call DeleteObject(mlngFont)
                    Call ReleaseFontIndirect(mobjDraw)
                End If
                '���������С
                Set gstdSet = New StdFont
                gstdSet.Name = "����"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
                mlngFont = CreateFontIndirect(T_Font)
                mlngOldFont = SelectObject(mlngMemDC, mlngFont)
                dbl��λֵ = Val(Nvl(!��λֵ, 0))
                sin�̶ȼ�� = Nvl(!�̶ȼ��, Val(Nvl(!��λֵ, 0)) * 10)
                If sin�̶ȼ�� > Val(Nvl(!���ֵ)) - Val(Nvl(!��Сֵ)) Then
                    sin�̶ȼ�� = Val(Nvl(!���ֵ)) - Val(Nvl(!��Сֵ))
                End If
                sinAlertness = Nvl(!��ʾ��, 0)
                str˵�� = str˵�� & "��" & Nvl(!��¼��) & "(" & Nvl(!��¼��, "*") & ")"
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), "", Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
                
                intTens_digit = 1
                lngCurY = lngY + (T_DrawClient.�е�λ * Val(Nvl(!�����, 0)))
                blnFirst = False
                Do While True
                    bln��ʾ�̶� = False
                    If blnFirst = False Then     '�ս���ѭ������ʱȡ�����ֵ
                        sin�̶� = Nvl(!���ֵ, 0)
                        sinBegin�̶� = sin�̶�
                        str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                        blnFirst = True
                    Else                    '����õ�ÿ���̶ȵ�ֵ
                        sin�̶� = sin�̶� - dbl��λֵ     '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                    End If
    
                    '�������õĿ̶ȼ����ʾ�̶�ֵ
                    If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                    If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - IIf(T_DrawClient.˫��, sin�̶ȼ�� * 2, sin�̶ȼ��)
                    If sinBegin�̶� < Val(Format(Nvl(!��Сֵ), "#0.00")) Then sinBegin�̶� = Val(Format(Nvl(!��Сֵ), "#0.00"))
    
                    If bln��ʾ�̶� Then
                        '�������ֵ�������ߵ�λ�ظ�
                        lngCurX = T_DrawClient.��������.Left - mobjDraw.ScaleX(mobjDraw.TextWidth(Val(Format(sin�̶�, "#0.0"))), vbTwips, vbPixels)
                        lngCurX = lngCurX - (mobjDraw.ScaleY(mobjDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY = lngY Then
                            Call GetTextRect(mobjDraw, lngCurX, lngCurY + (mobjDraw.ScaleY(mobjDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        ElseIf lngCurY = lngMaxY Then
                            Call GetTextRect(mobjDraw, lngCurX, lngCurY - (mobjDraw.ScaleY(mobjDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        Else
                            Call GetTextRect(mobjDraw, lngCurX, lngCurY, Val(Format(sin�̶�, "#0.0")))
                        End If
                        Call DrawText(mlngMemDC, Val(Format(sin�̶�, "#0.0")), -1, T_LableRect, DT_CENTER)
                    End If
                    If Val(Format(sin�̶�, "#0.00")) <= Val(Format(Nvl(!��Сֵ), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                        str��Сֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                        '��Ӹ���Ŀ(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
                        gstrFields = "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ"
                        gstrValues = Nvl(!��Ŀ���) & "|" & Nvl(!���ֵ) & "|" & Nvl(!��Сֵ) & "|" & dbl��λֵ & "|" & _
                            str���ֵ���� & "|" & str��Сֵ���� & "|" & T_DrawClient.�е�λ & "," & T_DrawClient.�е�λ & "|" & intTens_digit & "|" & !��¼ɫ
                        Call Record_Add(mrsDrawItems, gstrFields, gstrValues)
                        '���������
                        If blnDoubleRow = False And sinAlertness > Val(Nvl(!��Сֵ)) And sinAlertness < Val(Nvl(!���ֵ)) Then
                            '�������ֵ�뵱ǰֵ֮��Ĳ��,�Լ���Сֵ,����õ������ٸ��̶�,�ٸ��ݵ�λ�̶ȵõ�ʵ������
                            sinCurAlerY = Val(GetYCoordinate(mobjDraw, mrsDrawItems, Val(Nvl(!��Ŀ���)), sinAlertness))
                            Call DrawLine(mlngMemDC, T_DrawClient.��������.Left, CLng(sinCurAlerY), lngMaxX, CLng(sinCurAlerY), PS_SOLID, 1, RGB_RED)
                        End If
                        Exit Do
                    End If
                    lngCurY = lngCurY + T_DrawClient.�е�λ
                Loop
                '��ԭ������Ϣ
                Call SelectObject(mlngMemDC, mlngOldFont)
                Call DeleteObject(mlngFont)
                Call ReleaseFontIndirect(mobjDraw)
                sinBegin�̶� = 0
                sin�̶� = 0
            End If
        .MoveNext
        Loop
    End With
        
    '��������
    Set gstdSet = New StdFont
    gstdSet.Name = "����"
    gstdSet.Size = 9
    Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngMemDC, mlngFont)
    
    mblnRedraw = False                      '����һ�κ�Ͳ��ٻ���
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub Paint_Construct()

    Dim lngRGB  As Long
    Dim blnLine As Boolean              '��������������ʱ,���ʲ�����
    Dim str���� As String               '��¼�����������ڵ���(X����)
    Dim strԭֵ As String, sinX����ԭ As Single, sinY����ԭ As Single
    Dim dbl��ֵ As Double, dblMinValue As Double, dblMaxValue As Double
    Dim bln�������� As Boolean
    Dim lng���²�����ʾ��ʽ As Long
    Dim lng�и� As Long
    Dim lngWith As Long
    Dim bln���� As Boolean
    Dim strWaveReview As String, lngWaveReviewColor As Long '����:���¸���ķ��ż���ɫ
    On Error GoTo Errhand

    '��ʼ��ͼ�����ͼ�����ݵ�������ص��Ĵ������¸��ˡ�ͼ�α�������������׾��
    '�Ȼ���(��������)
    '�ٴ���������׾
    '�����ͼ��
    lng�и� = T_BodyStyle.lng�����и� \ Screen.TwipsPerPixelY
    
    lng���²�����ʾ��ʽ = Val(zlDatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0"))
    strWaveReview = zlDatabase.GetPara("���¸��Ժϸ����", glngSys, 1255, "v")
    '75319
    lngWaveReviewColor = Val(zlDatabase.GetPara("���¸��Ժϸ���ɫ", glngSys, 1255, "10485760"))
    
    With mrsPoint
        .Filter = ""
        '�Ȼ���
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "��Ŀ���,ʱ��"
        Do While Not .EOF
            If Val(zlCommFun.Nvl(!״̬)) <> 3 Then
                '�����µĺ��洦��,������
                If Not ((!��Ŀ��� = gint���� Or !��Ŀ��� = gint��ʹǿ��) And !��� = 1) Then
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
                        Call SetTextColor(mlngMemDC, lngWaveReviewColor)
                        Call GetTextRect(mobjDraw, !X����, !Y���� - Screen.TwipsPerPixelY, strWaveReview, T_DrawClient.�е�λ, False)
                        Call DrawText(mlngMemDC, strWaveReview, -1, T_LableRect, DT_CENTER)
                    End If
                    
                    '�����:56886,����,2013-05-06
                    bln���� = GetSymbol(!��Ŀ���, !��λ, !�ص���Ŀ, !����)
                    lngWith = 0
                    If bln���� Then
                        lngWith = mobjDraw.TextWidth("��") / 4 / T_TwipsPerPixel.X
                    End If
                    
                    If sinX����ԭ <> 0 And blnLine Then
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2 - lngWith, !Y����, sinX����ԭ + T_DrawClient.�е�λ / 2, sinY����ԭ, PS_SOLID, 1, lngRGB)
                    End If
                    

                    If !�Ͽ� = 0 Then
                        sinX����ԭ = !X���� + lngWith
                        sinY����ԭ = !Y����
                    Else
                        sinX����ԭ = 0
                    End If

                    '�˴�������Ŀ�߳���Ŀ�����ֵ ��С����Ŀ��Сֵ
                    If Not (!��Ŀ��� = gint���� And Trim(Nvl(!��ֵ)) = "����") Then
                        dbl��ֵ = Val(zlCommFun.Nvl(!��ֵ))
                        '�ص�ʱ����ſ�ǰ��Ϊ׼
                        If !�ص� = 0 Then
                            If dbl��ֵ < dblMinValue Then
                                Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + IIf(T_DrawClient.�е�λ < lng�и�, lng�и�, T_DrawClient.�е�λ) * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, 1, lngRGB, True)
                            End If
                            
                            If dbl��ֵ > dblMaxValue Then
                                Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - IIf(T_DrawClient.�е�λ < lng�и�, lng�и�, T_DrawClient.�е�λ) * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, 1, lngRGB, True)
                            End If
                        End If
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
        If str���� <> "" Then Call CreatePolyNew(mrsPoint, mobjDraw, mlngMemDC, mstr��ʼʱ��, str����, mint����Ӧ�� = 2)

        '������ͼ��
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "��Ŀ���,ʱ��"
        
    
        Do While Not .EOF
            If Val(zlCommFun.Nvl(!״̬)) <> 3 Then
                If (!��Ŀ��� = gint���� Or !��Ŀ��� = gint��ʹǿ��) And !��� = 1 Then
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
                        Call DrawLine(mlngMemDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + (T_Size.H / 4), sinX����ԭ + T_DrawClient.�е�λ / 2, sinY����ԭ, PS_DOT, 1, RGB_RED, False)
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
    Dim blnCurBeginTop As Boolean  '���±�־�Ƿ�Ӷ������
    Dim Y As Long, X As Long, Y1 As Long
    Dim bln�ı� As Boolean
    Dim lngX          As Long, lngY As Long
    Dim strComment    As String, strTemp As String, strText As String
    Dim intNum        As Integer
    Dim intAscCharNum As Integer
    Dim varNote()     As String
    Dim i  As Integer, j As Integer

    On Error GoTo Errhand
    
    '����
    bytδ��˵����ʾλ�� = Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0"))
    blnCurBeginTop = (Val(zlDatabase.GetPara("���±�־���λ��", glngSys, 1255, "0")) = 1)
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
    
    If blnCurBeginTop = False Then
        Y1 = GetYCoordinate(mobjDraw, mrsDrawItems, gint����, 42, mlngMemDC)
    Else
        Y1 = T_DrawClient.��������.Top
    End If
    
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

                            If Y < T_DrawClient.��������.Bottom Then
                                strText = Mid(strComment, i, 1)
                                Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)
                                '���������Ϣ
                                If T_DrawClient.��������.Bottom - Y >= T_Size.H - 1 Then
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
                    If Y < T_DrawClient.��������.Bottom Then
                        strText = Mid(strComment, i, 1)
                        Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)

                        If Asc(strText) < 0 Then
                            If (intAscCharNum - intNum) Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If
                         
                        '���������Ϣ
                        If T_DrawClient.��������.Bottom - Y >= T_Size.H - 1 Then
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
    Call OutPutTextNew(mobjDraw, mrsDrawItems, mlngMemDC, mrsNote, mstr��ʼʱ��, blnCurBeginTop)
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
    Dim strValue As String
    
    mrsItems.Filter = "��Ŀ���=" & lng��Ŀ���
    If mrsItems.EOF Then Exit Function
    
'    If InStr(1, Nvl(mrsItems!��Ŀֵ��), ";") = 0 Then
'        dblvalue = Val(Nvl(mrsItems!��Сֵ, 0))
'    Else
'        dblvalue = Val(Split(mrsItems!��Ŀֵ��, ";")(0))
'    End If
    dblvalue = Val(Nvl(mrsItems!��Сֵ, 0))
    strValue = Nvl(mrsItems!�ٽ�ֵ)
    If InStr(1, strValue, ";") <> 0 Then
        strValue = Split(strValue, ";")(0)
    Else
        strValue = ""
    End If
    
    If IsNumeric(strValue) = True And Val(strValue) <= Val(Nvl(mrsItems!���ֵ)) And Val(strValue) >= Val(Nvl(mrsItems!��Сֵ)) Then
        dblvalue = Val(strValue)
    Else
        '���������Сֵ��Ч�����������СֵΪ35
        If lng��Ŀ��� = gint���� And dblvalue < 35 Then dblvalue = 35
    End If
    
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
    If InStr(1, strValue, ";") <> 0 Then strValue = Split(strValue, ";")(1)
    If IsNumeric(strValue) = True And Val(strValue) <= Val(Nvl(mrsItems!���ֵ)) And Val(strValue) >= Val(Nvl(mrsItems!��Сֵ)) Then dblvalue = Val(strValue)
    GetMaxValue = dblvalue
End Function

Private Sub ReadBoyData(ByVal blnAutoAdjust As Boolean)
    
    On Error GoTo Errhand
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
    Dim strSQL As String
    Dim lngColor As Long, lng�к� As Long, lng��Ŀ���  As Long
    Dim str���� As String
    Dim dbl��ֵ As Double, dblMinValue As Double, dblMaxValue As Double
    Dim strTmpString0 As String, strTmpString1 As String, strTmpString2 As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim arrValues() As String
    Dim arrTmpValue() As Variant, arrTmpNote As Variant
    Dim i As Integer, j As Integer
    Dim int��ʾ As Integer
    Dim rs���� As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim blnӤ�����µ���ʾ��Ժ As Boolean, bln�����ʾ��Ժ As Boolean
    Dim lng���²�����ʾ��ʽ As Long
    Dim int��� As Integer
    Dim lngSignColor As Long '����:�����Զ���ʶ����ɫ
    Dim lngNoRecordColor As Long '����:δ��˵����ʾ��ɫ
    Dim bln��Ʋ�ת��Ժ As Boolean
    
    On Error GoTo Errhand
    
    '71950:������,2014-06-11,���µ�δ��˵����ʾ��ɫ
    lngNoRecordColor = Val(zlDatabase.GetPara("δ��˵����ʾ��ɫ", glngSys, 1255, "16711680"))
    '��¼������Ϣ
    strFileds = "��Ŀ���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|X����," & adDouble & ",5|ʱ��," & adLongVarChar & ",20"
    Call Record_Init(rs����, strFileds)
    
    '��ȡ���в�λ��Ϣ
    strSQL = "Select ��Ŀ���, ��λ,ȱʡ�� From ���²�λ"
    Call zlDatabase.OpenRecordset(rsPart, strSQL, "��ȡ���²�λ")
    
    
    '����������Ŀ��Ҫ�����ֶ����������������Ƿ������ⵥ����ʾ,ĿǰȱʡΪ��ʾ
    '-----------------------------------------------------------------------
    gstrSQL = "SELECT /*+ Rule*/  C.ID ���,a.����ʱ�� As ʱ��,C.��ʾ,C.��¼���� As ��ֵ,C.���²�λ,C.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵�� " & _
                "FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E,Table(Cast(f_num2list([6]) As zlTools.t_Numlist)) F " & _
                "Where B.ID=A.�ļ�ID " & _
                    "And A.ID = C.��¼ID " & _
                    "AND B.ID=[1] " & _
                    "AND B.����id=[2] " & _
                    "AND B.��ҳid=[3] " & _
                    "AND D.��Ŀ���=c.��Ŀ��� " & _
                    "AND c.��¼����=1 " & _
                    "AND E.��Ŀ���=D.��Ŀ��� " & _
                    "AND F.COLUMN_VALUE=D.��Ŀ��� " & _
                    "AND a.����ʱ�� BETWEEN [4] And [5] And c.��ֹ�汾 Is Null And D.��¼��<>2" & _
                "Order By A.����ʱ��,DECODE(C.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���)"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ����", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), T_BodyItem.str������Ŀ)
        
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
                SinX = GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinateNew(SinX, Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinateNew(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss"))
                
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
                        !δ��˵�� & "|" & lngNoRecordColor & "|" & SinX & "|0|0|0|0|" & Val(zlCommFun.Nvl(!��ʾ))
                   
                    If blnAdd Then
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                         Call Record_Add(mrsNote, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(mrsNote!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(mrsNote!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                             blnAllow = GetCanvasCenterNew(CDate(Format(mrsNote!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
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
                            blnAllow = GetCanvasCenterNew(CDate(Format(mrsPoint!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
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
                        blnAllow = GetCanvasCenterNew(CDate(Format(mrsPoint!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
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
    
    '��������������,��ʹ��ʹ
    For j = 0 To 1
        lng��Ŀ��� = IIf(j = 0, gint����, gint��ʹǿ��)
        arrTmpValue = Array()
        mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=0"
        With mrsPoint
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ���
        If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If mrsPoint.RecordCount <> 0 Then
                gstrFields = "��ע": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                Call Record_Update(mrsPoint, gstrFields, gstrValues, "���|" & zlCommFun.Nvl(mrsPoint!���))
            End If
        Next i
        
        arrTmpValue = Array()
        mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=1"
        With mrsPoint
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ���
        If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=0 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If mrsPoint.RecordCount = 0 Then
                mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                mrsPoint.Delete
            End If
        Next i
    Next j
    
    
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
                    blnAllow = GetCanvasCenterNew(CDate(Format(mrsPoint!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")), SinX)
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
    
    '62989:������,2013-07-24,���µ������ʾ��ɫ
    lngSignColor = Val(zlDatabase.GetPara("���µ������ʾ��ɫ", glngSys, 1255, "255"))
    
    '��ȡ���������±���Ϣ
    '-----------------------------------------------------------------------
    gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
    gstrSQL = "" & _
             " Select A.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.��¼����,C.��Ŀ����,C.δ��˵��" & _
             " FROM ���˻����ļ� B, ���˻������� A, ���˻�����ϸ C" & _
             " Where B.ID=A.�ļ�ID And A.ID = C.��¼ID AND B.ID=[1] AND Nvl(B.Ӥ��, 0)=[6] AND B.����id=[2] AND B.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
             " AND MOD(C.��¼����,10) <> 1  AND A.����ʱ�� BETWEEN [4]  And [5]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������±����Ϣ", T_Patient.lng�ļ�ID, T_Patient.lng����ID, T_Patient.lng��ҳID, Int(CDate(mstr��ʼʱ��)), CDate(mstr����ʱ��), T_Patient.lngӤ��, T_Patient.lng����ȼ�)
    
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
        
    With rsData
        Do While Not .EOF
            bytShow = 1
            str���� = Trim(zlCommFun.Nvl(!��¼����))
            
            lng�к� = IIf(!��¼���� = 2, 10, IIf(!��¼���� = 6, 11, 14))
            
            '����������ʾ��Ҫ���⴦��
            If !��¼���� = 4 Then
                str���� = Trim(zlCommFun.Nvl(!��Ŀ����))
                
                If str���� = "����" Then
                    bytShow = T_BodyFlag.����
                ElseIf str���� = "����" Then
                    bytShow = T_BodyFlag.����
                Else
                    bytShow = T_BodyFlag.����
                End If
                
                If bytShow = 2 And Not blnAutoAdjust Then
                    str���� = str���� & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                Else
                    str���� = !��Ŀ����
                End If
                lngColor = lngSignColor
            Else
                lngColor = IIf(Not IsNumeric(Nvl(!δ��˵��)), RGB_BLUE, Val(Nvl(!δ��˵��)))
            End If
            
            If bytShow > 0 Then
                SinX = Val(GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), mstr��ʼʱ��))
                
                mrsNote.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=" & !��¼���� & " And ʱ��='" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "'"
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
    
    blnӤ�����µ���ʾ��Ժ = (zlDatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, 1) = 1)
    '�����:63525,�޸���:����,��Ժ��ʶ����ʾ����Ʊ�ʶ��ʾʱ�����Զ�תΪ��Ժ��
    bln��Ʋ�ת��Ժ = (zlDatabase.GetPara("��Ʊ�ʶ���Զ�ת��Ϊ��Ժ", glngSys, 1255, 1) = 0)
    
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
                Case 15
                    bytShow = T_BodyFlag.ת����
                End Select
                 
                If bytShow > 0 Then
                    'Ŀǰ3��4 �����ת�� 3-��ʾ˵���Ϳ��� 4 ��ʾ˵�������ң�ʱ��
                    If lng�к� = 9 And bln�����ʾ��Ժ = True And bln��Ʋ�ת��Ժ = True Then
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
                    
                    SinX = Val(GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), mstr��ʼʱ��))
                    mrsNote.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=3 And ʱ��='" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "'"
                    
                    If mrsNote.BOF Then
                        gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|3|" & _
                            str���� & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
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
                        
                        SinX = Val(GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), mstr��ʼʱ��))
                        mrsNote.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=13 And ʱ��='" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "'"
                        
                        If mrsNote.BOF Then
                            gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|13|" & _
                                str���� & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
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
    bytTag = Abs(Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0")))
    lng���²�����ʾ��ʽ = Val(zlDatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0"))
    '�������²��� ���²���ʼ����ʾ�� 35 �����棬ֻ��δ��˵����ʾ�������������Ž���������δ��˵���У���������������±���
    If Left(strTmpString0, 1) = ";" Then
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"
        If lng���²�����ʾ��ʽ = 0 Or lng���²�����ʾ��ʽ = 2 Then
            arrValues = Split(strTmpString0, "|")
            arrValues(3) = "�� "
            strTmpString0 = Join(arrValues, "|")
        End If
        strTmpString0 = Mid(strTmpString0, 2)
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
    
    mint����Ӧ�� = 0
    If Not (mrsItems Is Nothing) Then If mrsItems.State = 1 Then mrsItems.Close
    '���ִ������øò��˵Ļ����¼��Ŀ
    gstrSQL = " Select C.��Ŀ���,C.��Ŀ����,C.��Ŀ����,C.��Ŀ����,C.��Ŀ����,C.��ĿС��,C.��Ŀ��ʾ,C.��Ŀ��λ,C.��Ŀֵ��,A.���ֵ,A.��Сֵ,A.�ٽ�ֵ,A.��¼��,A.��¼ɫ,C.����ȼ�,C.Ӧ�÷�ʽ,C.���ò���" & _
              " From ���¼�¼��Ŀ A,�����¼��Ŀ C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist))  D" & _
              " where A.��Ŀ���(+)=C.��Ŀ���" & _
              " And C.��Ŀ���=D.COLUMN_VALUE " & _
              " Order by C.��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ʼ��", T_BodyItem.str������Ŀ & "," & T_BodyItem.str�������)
    mrsItems.Filter = "��Ŀ���=-1"
    If mrsItems.RecordCount > 0 Then mint����Ӧ�� = zlCommFun.Nvl(mrsItems("Ӧ�÷�ʽ").Value, 2): mrsItems.Filter = ""
    
    If Not mrsGraph Is Nothing Then If mrsGraph.State = 1 Then mrsGraph.Close
    
    lngMax = mobjBuffer.ScaleWidth \ gintBmpW      'һ���ܱ�����ٸ�ͼƬ?
    '�����������ͼ�����(���������ص����),ȫ����ȡ��picBuffer��,�˴��������Ŀ�Ĳ�λ�����Ӧ��ͼ�����
    gstrFields = "��Ŀ���," & adDouble & ",18|��λ," & adLongVarChar & ",50|��¼��," & adLongVarChar & ",50|" & _
                 "��¼ɫ," & adDouble & ",18|�ص���Ŀ," & adLongVarChar & ",20|��," & adDouble & ",5|��," & adDouble & ",5"    '�ص���ĿӦ����Ŀ��Ŵ�С����,��:1,4,5
    Call Record_Init(mrsGraph, gstrFields)
    
    '�ȸ������²�λװ��
    gstrSQL = " Select ��Ŀ���,'' AS ��λ, ��¼�� ��Ƿ���,��¼ɫ �����ɫ,1 չ�ַ�ʽ,'��' AS �ص���Ŀ From ���¼�¼��Ŀ Order by ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������Ŀ��չ�ַ�ʽ")
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
    Set rsOverlap = zlDatabase.OpenSQLRecord(gstrSQL, "�ٸ��������ص����װ��")
    gstrSQL = " Select ���,�ϼ����,��Ŀ���,���²�λ From �����ص���� Where ��Ŀ��� is not null Order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ص�������Ŀ")
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
            lngCurX = GetXCoordinateNew(!ʱ��, strBeginDate)
            If mint����Ӧ�� = 2 And !��Ŀ��� = -1 Then
                mrsPoint.Filter = "��Ŀ���=" & gint���� & " And  X����<=" & !X����
            Else
                If Val(!��Ŀ���) = gint���� Or Val(!��Ŀ���) = gint��ʹǿ�� Then
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
            If Not ((Val(!��Ŀ���) = gint���� Or Val(!��Ŀ���) = gint��ʹǿ��) And Val(zlCommFun.Nvl(!���)) = 1) Then
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
        T_LableRect.Left = T_LableRect.Left - 1
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
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    On Error GoTo Errhand
    
    strSQL = " Select /*+ Rule*/ Count(*) ��¼" & _
             " From ���¼�¼��Ŀ A, �����¼��Ŀ B,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) C" & _
             " Where A.��Ŀ���=B.��Ŀ��� And B.��Ŀ���=C.COLUMN_VALUE" & _
             " Order by B.��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ��ʼ��", T_BodyItem.str������Ŀ)
    
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
    Dim strSQL As String, strNewSql As String
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
    Dim strMarkDate As String, strFileBeginTime As String
    Dim arrParam() As String
    '----------------------------------------------------
    '������Ϣ����
    '----------------------------------------------------
    Dim lng�ļ�ID As Long, lng����ID As Long, lng��ҳID  As Long
    Dim lng����ID As Long, lngӤ��  As Long, lng����ȼ� As Long
    '----------------------------------------------------
    '���µ���ʽ����
    '----------------------------------------------------
    Dim MT_BodyStyle As type_BodyStyle
    Dim MT_BodyItem As type_BodyItem
    
    On Error GoTo ErrHandle
    
    If strParam <> "" Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) < 2 Then
            MsgBox "strParam������Ϊ��ʱ,���봫���ļ�ID;����ID;��ҳID��", vbInformation, gstrSysName
            Exit Function
        End If
        lng�ļ�ID = Val(arrParam(0))
        lng����ID = Val(arrParam(1))
        lng��ҳID = Val(arrParam(2))
        If UBound(arrParam) > 2 Then lng����ID = Val(arrParam(3))
        If UBound(arrParam) > 3 Then lngӤ�� = Val(arrParam(4))
        lng����ȼ� = 3
        gstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ȼ�", lng����ID, lng��ҳID)
        If rsTmp.BOF = False Then lng����ȼ� = Nvl(rsTmp("����ȼ�"), 3)
    Else
        lng�ļ�ID = T_Patient.lng�ļ�ID
        lng����ID = T_Patient.lng����ID
        lng��ҳID = T_Patient.lng��ҳID
        lng����ID = T_Patient.lng����ID
        lngӤ�� = T_Patient.lngӤ��
        lng����ȼ� = T_Patient.lng����ȼ�
    End If
    
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
    strSQL = "Select ��ʽ From ����ҳ���ʽ Where ���� = 3 And ��� In (Select A.ҳ�� From �����ļ��б� A,���˻����ļ� B Where A.Id = B.��ʽID and B.ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ļ���ӡ����", lng�ļ�ID)
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
        If Trim(zlDatabase.GetPara("���µ���ӡ��", glngSys, 1255, "")) = "" Then
            MsgBox "û�����ô�ӡ��,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
            strPrintName = Printer.DeviceName
        Else
            strPrintName = Trim(zlDatabase.GetPara("���µ���ӡ��", glngSys, 1255, Printer.DeviceName))
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
    gPrinter.intBin = Val(zlDatabase.GetPara("���µ���ֽ", glngSys, 1255, Printer.PaperBin))
    
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
    '�ڶ�ȡ�ļ�֮ǰ���Ƚ�֮ǰ�ļ�����ʽ��������
    With MT_BodyStyle
        .lng��ʼʱ�� = T_BodyStyle.lng��ʼʱ��
        .lngʱ���� = T_BodyStyle.lngʱ����
        .lng������ = T_BodyStyle.lng������
        .lng���� = T_BodyStyle.lng����
        .lng�̶ȿ�� = T_BodyStyle.lng�̶ȿ��
        .lng�����п� = T_BodyStyle.lng�����п�
        .lng�����и� = T_BodyStyle.lng�����и�
        .lng���߶� = T_BodyStyle.lng���߶�
        .str��ͷ���� = T_BodyStyle.str��ͷ����
        .str�����ı� = T_BodyStyle.str�����ı�
        .str�������� = T_BodyStyle.str��������
        .lng���߿��� = T_BodyStyle.lng���߿���
        .lng������ = T_BodyStyle.lng������
        .lng�±��߶� = T_BodyStyle.lng�±��߶�
        .blnר�� = T_BodyStyle.blnר��
    End With
    With MT_BodyItem
        .str������� = T_BodyItem.str�������
        .str�����Ŀ = T_BodyItem.str�����Ŀ
        .str������Ŀ = T_BodyItem.str������Ŀ
    End With
    '��ȡ���ļ�����ʽ(��Ҫɾ����������ӡ��Ҫ������ȡ)
    If Not GetStyleBody(lng�ļ�ID, lng����ȼ�, lngӤ��, lng����ID, blnPrint) Then Exit Function
    intBaby = lngӤ��
    '------------------------------------------------------------------------------------------------------------------
    lngBeginY = gPrinter.lngTop
    lngIndex = mintPage
    
    '���ֻ��ӡ��ǰ��ֻ����ʼ�ͽ���дͬһҳ��
    Set mfrmCaseTendBodyPrint = New frmCaseTendBodyPrint
    Load frmTendFileRead
    Call frmTendFileRead.InitRechBox(lng�ļ�ID)
    strMarkDate = ""
    '��ȡ�û����õ����µ���ʼʱ��(Ӥ��������Ӥ������ʱ��Ϊ׼)
    strSQL = "select ��ʼʱ�� from ���˻����ļ� where ID=[1] and ����ID=[2] and ��ҳid=[3] and nvl(Ӥ��,0)=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���µ���ʼʱ��", lng�ļ�ID, lng����ID, lng��ҳID, lngӤ��)
    If rsTmp.RecordCount <> 0 Then
        strMarkDate = Format(rsTmp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If strMarkDate <> "" Then strMarkDate = "to_date('" & strMarkDate & "','yyyy-MM-dd hh24:mi:ss')"
    
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ)����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "   (SELECT ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0 AND B.������� = 'Z'" & vbNewLine & _
                "                AND Instr(',3,5,11,', ',' || c.�������� || ',') > 0 And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    '˵��:Ŀǰ����ר�����µ������˿���ͬʱ���ڶ�����µ������µ���ʼʱ�����ֹʱ��Ĺ�������:
    '����ļ��Ŀ�ʼʱ�䲻Ϊ�ղ��Ҵ��ڵ��ڲ�����Ժʱ���Ӥ������ʱ��,���µ��Ŀ�ʼʱ�����ļ���ʼʱ��Ϊ׼,�����Բ�����Ժʱ���Ӥ������ʱ��Ϊ׼
    '����ļ�����ֹʱ�䲻Ϊ�ղ���С�ڵ��ڲ��˻�Ӥ����Ժʱ�䣨δ��Ժ���ܲ��ܴ��ڵ�ǰʱ�䣩,���µ�����ʱ�����ļ���ʼʱ��Ϊ׼���������µ�����ʱ���Բ��˻�Ӥ����Ժʱ��Ϊ׼(δ��ԺΪ��ǰʱ��)
    '����ļ�����ֹʱ��Ϊ��,����ԭ�з�ʽ,��������Ѿ���Ժ�����ѳ�Ժʱ��Ϊ׼,δ��Ժ���ѵ�ǰʱ������ݽ���ʱ��Ϊ׼.
    '��ȡ�˲��˵����µ���ҳ��
    '------------------------------------------------------------------------------------------------------------------
    strSQL = " SELECT  ��Ժʱ��,��Ժʱ��,1 + TRUNC((TO_DATE(TO_CHAR(��Ժʱ��,'yyyy-MM-dd'),'yyyy-MM-dd') -TO_DATE(TO_CHAR(��Ժʱ��,'yyyy-MM-dd'),'yyyy-MM-dd')) / " & T_BodyStyle.lng���� & ") AS ҳ��,����ʱ�� " & _
            "  From (" & _
                " SELECT DECODE(D.��ʼʱ��,NULL,DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��)," & vbNewLine & _
                "               DECODE(SIGN(D.��ʼʱ�� - DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��))," & vbNewLine & _
                "                      1," & vbNewLine & _
                "                      D.��ʼʱ��," & vbNewLine & _
                "                      DECODE(C.����ʱ��,NULL," & IIf(strMarkDate = "", "B.��Ժʱ��", strMarkDate) & ",C.����ʱ��))) AS ��Ժʱ��," & vbNewLine & _
                "    DECODE(D.����ʱ��,NULL," & vbNewLine & _
                "               DECODE(E.��¼,0," & vbNewLine & _
                "                      DECODE(SIGN(NVL(E.Ӥ��ʱ��, B.��Ժʱ��) - D.����ʱ��), 1, NVL(E.Ӥ��ʱ��, B.��Ժʱ��), D.����ʱ��)," & vbNewLine & _
                "                      NVL(E.Ӥ��ʱ��, B.��Ժʱ��))," & vbNewLine & _
                "               DECODE(SIGN(NVL(E.Ӥ��ʱ��, B.��Ժʱ��) - D.����ʱ��), 1, D.����ʱ��, NVL(E.Ӥ��ʱ��, B.��Ժʱ��))) ��Ժʱ��," & vbNewLine & _
                "    D.����ʱ��" & vbNewLine & _
                "    FROM (SELECT ����ID,��ҳID,MIN(��ʼʱ��) AS ��Ժʱ��," & vbNewLine & _
                "    MAX(NVL(��ֹʱ��, SYSDATE)) AS ��Ժʱ��" & vbNewLine & _
                "    FROM ���˱䶯��¼" & vbNewLine & _
                "    WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID = [3] GROUP BY ����ID,��ҳID) B," & vbNewLine & _
                "    (SELECT ����ID,��ҳID,����ʱ�� FROM ������������¼ WHERE ����ID =[2] AND ��ҳID =[3] AND ���=[4]) C ," & vbNewLine & _
                "    (SELECT NVL(����ʱ��, SYSDATE) ����ʱ��, ��ʼʱ��, ����ʱ��" & vbNewLine & _
                "       FROM (SELECT MAX(B.����ʱ��) ����ʱ��, MAX(A.��ʼʱ��) ��ʼʱ��, MAX(A.����ʱ��) ����ʱ��" & vbNewLine & _
                "              FROM ���˻����ļ� A, ���˻������� B" & vbNewLine & _
                "              WHERE A.ID = B.�ļ�ID(+) AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND A.Ӥ�� = [4])) D," & vbNewLine & _
                strNewSql & vbNewLine & _
                "WHERE B.����ID=E.����ID And B.��ҳID=E.��ҳID And B.����ID=C.����ID(+) AND B.��ҳID=C.��ҳID(+))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, mstrTitle, lng�ļ�ID, lng����ID, lng��ҳID, lngӤ��)
    intCount = 0
    For intCOl = 0 To rsTmp("ҳ��").Value - 1
    
        strDateFrom = Format(rsTmp("��Ժʱ��").Value + T_BodyStyle.lng���� * intCOl, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + T_BodyStyle.lng���� * (intCOl + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
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
        If PrintOrPreviewBodyStateNew(objPrint, lng����ID, lng��ҳID, lng�ļ�ID, intBaby, _
                lng����ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, False, _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, , mblnMoved) = True Then
                
                If blnPrint = False Then
                    mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, lng����ID, lng��ҳID, _
                        lng�ļ�ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                        CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, lng����ID, lngӤ��
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
            If PrintOrPreviewBodyStateNew(objPrint, lng����ID, lng��ҳID, lng�ļ�ID, intBaby, _
                lng����ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> lngIndex, _
                CInt(Split(strArrFromTo(intCOl), ";")(1)), CInt(Split(strArrFromTo(intCOl), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "δ֪���󣬴�ӡʧ�ܣ�", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                '68407:������,2013-12-05,�޸�intCOl = UBound(strArrFromTo)ΪintCOl=lngIndexEnd,��Ȼ�ᵼ�´�ӡ������
                If intCOl = lngIndexEnd Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, lng����ID, lng��ҳID, _
            lng�ļ�ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, lng����ID, lngӤ��
        Else '������ӡ�Ǽ�¼��ӡ�Ŀ�ʼҳ�źͽ���ҳ��
            strSQL = "zl_���µ�����_Printer(" & lng�ļ�ID & "," & lngIndex + 1 & "," & lngIndexEnd + 1 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "zl_���µ�����_Printer")
        End If
        
    Case 2          '�ӵ�һҳ������ӡ,��ȫ����ӡ
        strPage = 0
        For intCOl = 0 To UBound(strArrFromTo)
            If blnPrint = True Then Printer.Print ""
            If PrintOrPreviewBodyStateNew(objPrint, lng����ID, lng��ҳID, lng�ļ�ID, intBaby, _
                lng����ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> 0, _
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
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, lng����ID, lng��ҳID, _
            lng�ļ�ID, CInt(Split(strArrFromTo(0), ";")(1)), _
                CInt(Split(strArrFromTo(0), ";")(1)), intPageNo, strArrFromTo, strPage, lng����ID, lngӤ��
        End If
    End Select
    
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And gPrinter.intPage = 256 Then DelCustomPaper
    
    Unload frmTendFileRead
    
    '------------------------------------------------------------------------------------------------------------------
ReStoreCuve:
    '��Ԥ������ӡ��ɺ�ָ�֮ǰѡ���ļ�����ʽ(������ӡ���ܵ���֮ǰ���ļ���ʽ�����仯)
    With T_BodyStyle
        .lng��ʼʱ�� = MT_BodyStyle.lng��ʼʱ��
        .lngʱ���� = MT_BodyStyle.lngʱ����
        .lng������ = MT_BodyStyle.lng������
        .lng���� = MT_BodyStyle.lng����
        .lng�̶ȿ�� = MT_BodyStyle.lng�̶ȿ��
        .lng�����п� = MT_BodyStyle.lng�����п�
        .lng�����и� = MT_BodyStyle.lng�����и�
        .lng���߶� = MT_BodyStyle.lng���߶�
        .str��ͷ���� = MT_BodyStyle.str��ͷ����
        .str�����ı� = MT_BodyStyle.str�����ı�
        .str�������� = MT_BodyStyle.str��������
        .lng���߿��� = MT_BodyStyle.lng���߿���
        .lng������ = MT_BodyStyle.lng������
        .lng�±��߶� = MT_BodyStyle.lng�±��߶�
        .blnר�� = MT_BodyStyle.blnר��
    End With
    With T_BodyItem
        .str������� = MT_BodyItem.str�������
        .str�����Ŀ = MT_BodyItem.str�����Ŀ
        .str������Ŀ = MT_BodyItem.str������Ŀ
    End With
    Call InitPara(T_BodyStyle.blnר��)
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GoTo ReStoreCuve
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

Private Sub vsf_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call RaiseShowTipInfo(vsf.Body, 2, X, Y)
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

Private Sub DrawDownTabAnsyGrade(ByVal lngDc As Long, ByVal objDraw As Object, arrText() As String, ByVal Row As Long, ByVal Col As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean, Optional ByVal blnFormat As Boolean = False)
'---------------------------------------------------
'���� ���������
'˵�� AnsyGrade=True���ܵ��ô˺���
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer, intOldSize As Integer
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngBackColor As Long, lngForeColor As Long
    Dim stdSet As StdFont, stdOldset As StdFont
    Dim LPoint As T_LPoint, T_ClientRect As RECT
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    Dim lngMaxWidth As Long
    
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
    lngOldBrush = SelectObject(lngDc, lngBrush)
    Call FillRect(lngDc, T_ClientRect, lngBrush)
    '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
    Call SelectObject(lngDc, lngOldBrush)
    Call DeleteObject(lngBrush)
    
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        '60529:������,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
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
    Set stdSet = New StdFont
    stdSet.Name = "����"
    stdSet.Size = intSize
    stdSet.Bold = False
    Set stdOldset = stdSet 'ԭʼ����
    
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , 1)
    '������
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngForeColor)
        Call DrawText(lngDc, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + 1
        Call ReleaseFontIndirect(objDraw)
    Else
        lngX = T_LableRect.Left
    End If

    If blnFormat = True Then '���ӷ�ĸ��ʾ
        intSize = 7
        objDraw.Font.Size = intSize
        '60529:������,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
            lngMaxWidth = objDraw.TextWidth(str2) / T_TwipsPerPixel.X
        Else
            lngMaxWidth = objDraw.TextWidth(str3) / T_TwipsPerPixel.X
        End If
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = intSize
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngForeColor)
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str2) / T_TwipsPerPixel.X) \ 2
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        'If T_LableRect.Top < Top Then T_LableRect.Top = Top - 1
        T_LableRect.Bottom = T_ClientRect.Bottom
        Call DrawText(lngDc, str2, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call ReleaseFontIndirect(objDraw)
        '������
        objDraw.Font.Size = intOldSize
        Call DrawLine(lngDc, lngX, lngY, lngX + lngMaxWidth, lngY)
        '�����ĸ
        intSize = 7
        objDraw.Font.Size = intSize
        lngY = lngY
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str3) / T_TwipsPerPixel.X) \ 2
        T_LableRect.Top = lngY
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = intSize
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngForeColor)
        Call DrawText(lngDc, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(mobjDraw)
    Else
        If str1 <> "" Then
            '����ϱ�
            intSize = 7
            objDraw.Font.Size = intSize
            Set stdSet = New StdFont
            stdSet.Name = "����"
            stdSet.Size = intSize
            Call SetFontIndirect(stdSet, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngForeColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < T_ClientRect.Top Then T_LableRect.Top = T_ClientRect.Top - 1
            Call DrawText(lngDc, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            Call ReleaseFontIndirect(mobjDraw)
            '�����벿��
            Call SetFontIndirect(stdOldset, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngForeColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDc, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(mobjDraw)
        Else
            Call SetFontIndirect(stdOldset, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngForeColor)
            Call DrawText(lngDc, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(mobjDraw)
        End If
    End If
    
    objDraw.Font.Size = intOldSize
    Set stdSet = Nothing
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


Private Function GetStyleBody(ByVal lng�ļ�ID As Long, ByVal lng����ȼ� As Long, lngӤ�� As Long, lng����ID As Long, Optional ByVal blnPrint As Boolean = False) As Boolean
'-------------------------------------------------------------------------------------------
'����:��ȡ�ļ����µ��ļ���ʽ
'-------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lng��ʽID As Long
    Dim str�����Ŀ As String
    Dim str������Ŀ As String
    Dim str������� As String
    Dim lngTabRows As Long
    Dim i As Integer
    Dim sinTwipsPerPixelX As Single, sinTwipsPerPixelY As Single
    
    On Error GoTo Errhand
    
    If blnPrint = True Then
        sinTwipsPerPixelX = Printer.TwipsPerPixelX
        sinTwipsPerPixelY = Printer.TwipsPerPixelY
    Else
        sinTwipsPerPixelX = Screen.TwipsPerPixelX
        sinTwipsPerPixelY = Screen.TwipsPerPixelY
    End If
    
    gstrSQL = "Select A.��ʽID,B.���� From ���˻����ļ� A, �����ļ��б� B Where a.��ʽid = b.Id And b.���� = 3 And b.���� = -1 And A.Id = [1]"
    If mblnMoved = True Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˻����ļ�", lng�ļ�ID)
    lng��ʽID = CLng(rsTemp!��ʽID)
    T_Patient.lng��ʽID = lng��ʽID
    T_BodyStyle.blnר�� = (Nvl(rsTemp!����, "0") = "1")
    If T_BodyStyle.blnר�� = True Then
        '�����ʽ��������
        gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, d.Ҫ�ر�ʾ " & _
                " From �����ļ��ṹ D, �����ļ��ṹ P" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ʽ����'" & _
                " Order By d.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���µ���ʽ����", lng��ʽID)
        With rsTemp
            Do While Not .EOF
                Select Case "" & !Ҫ������
                    Case "��ʼʱ��"
                        T_BodyStyle.lng��ʼʱ�� = Val(Nvl(!�����ı�))
                    Case "ʱ����"
                        T_BodyStyle.lngʱ���� = Val(Nvl(!�����ı�))
                    Case "������"
                        T_BodyStyle.lng������ = Val(Nvl(!�����ı�))
                    Case "����"
                        T_BodyStyle.lng���� = Val(Nvl(!�����ı�))
                    Case "�̶ȿ��"
                        T_BodyStyle.lng�̶ȿ�� = Fix(Val(Nvl(!�����ı�)) / sinTwipsPerPixelX) * sinTwipsPerPixelX
                    Case "�����п�"
                        T_BodyStyle.lng�����п� = Fix(Val(Nvl(!�����ı�)) / sinTwipsPerPixelX) * sinTwipsPerPixelX
                    Case "�����и�"
                        T_BodyStyle.lng�����и� = Fix(Val(Nvl(!�����ı�)) / sinTwipsPerPixelY) * sinTwipsPerPixelY
                    Case "���߶�"
                        T_BodyStyle.lng���߶� = Fix(Val(Nvl(!�����ı�)) / sinTwipsPerPixelY) * sinTwipsPerPixelY
                    Case "��ͷ����"
                        T_BodyStyle.str��ͷ���� = Nvl(!�����ı�)
                    Case "�����ı�"
                        T_BodyStyle.str�����ı� = Nvl(!�����ı�)
                    Case "��������"
                        T_BodyStyle.str�������� = Nvl(!�����ı�)
                    Case "���߿���"
                        T_BodyStyle.lng���߿��� = Val(Nvl(!�����ı�))
                    Case "���߶�1"
                        T_BodyStyle.lng�±��߶� = Fix(Val(Nvl(!�����ı�)) / sinTwipsPerPixelY) * sinTwipsPerPixelY
                    Case "������"
                        T_BodyStyle.lng������ = Val(Nvl(!�����ı�))
                End Select
                .MoveNext
            Loop
        End With
        
        '������Ŀ��������
        gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, d.Ҫ�ر�ʾ " & _
            " From �����ļ��ṹ D, �����ļ��ṹ P " & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '������Ŀ����'" & _
            " Order By d.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀ����", lng��ʽID)
        i = 0: str������Ŀ = ""
        With rsTemp
            Do While Not .EOF
                If i = 0 Then
                    str������Ŀ = !�����ı�
                Else
                    str������Ŀ = str������Ŀ & "," & !�����ı�
                End If
                i = i + 1
                .MoveNext
            Loop
            T_BodyItem.str������Ŀ = str������Ŀ
        End With
        
        '�����Ŀ��������
        gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, d.Ҫ�ر�ʾ " & _
            " From �����ļ��ṹ D, �����ļ��ṹ P " & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����Ŀ����'" & _
            " Order By d.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀ����", lng��ʽID)
        i = 0: str�����Ŀ = "": str������� = ""
        With rsTemp
            Do While Not .EOF
                 If i = 0 Then
                    str�����Ŀ = Nvl(!�����ı�) & ":" & Nvl(!Ҫ�ر�ʾ)
                    str������� = Nvl(!�����ı�)
                 Else
                    str�����Ŀ = str�����Ŀ & "@" & Nvl(!�����ı�) & ":" & Nvl(!Ҫ�ر�ʾ)
                    str������� = str������� & "," & Nvl(!�����ı�)
                 End If
                 i = i + 1
                .MoveNext
            Loop
            
            T_BodyItem.str������� = str�������
            T_BodyItem.str�����Ŀ = GetString(str�����Ŀ)
        End With
    Else '��׼���µ�
        '�����ʽ��������
        T_BodyStyle.lng��ʼʱ�� = Val(zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4))
        T_BodyStyle.lngʱ���� = 4
        T_BodyStyle.lng������ = 6
        T_BodyStyle.lng���� = 7
        T_BodyStyle.lng�̶ȿ�� = Fix(1350 / sinTwipsPerPixelX) * sinTwipsPerPixelX
        T_BodyStyle.lng�����п� = Fix(225 / sinTwipsPerPixelX) * sinTwipsPerPixelX
        T_BodyStyle.lng�����и� = Fix(90 / sinTwipsPerPixelY) * sinTwipsPerPixelY
        T_BodyStyle.lng���߶� = Fix(255 / sinTwipsPerPixelY) * sinTwipsPerPixelY
        T_BodyStyle.str��ͷ���� = "��       ��@" & IIf(T_Patient.lngӤ�� = 0, "ס Ժ �� ��", "�� �� �� ��") & "@����������@ʱ       ��"
        T_BodyStyle.str�����ı� = "���µ�"
        T_BodyStyle.str�������� = "����,20"
        T_BodyStyle.lng���߿��� = Val(zlDatabase.GetPara("�������߹̶��������", glngSys, 1255, "0"))
        T_BodyStyle.lng�±��߶� = Fix(255 / sinTwipsPerPixelY) * sinTwipsPerPixelY
        T_BodyStyle.lng������ = 0
        '��ȡ��Ŀ��Ϣ
        gstrSQL = _
            " SELECT Decode(b.��Ŀ���, 3, Decode(b.��¼��, 2, 1, b.�������), b.�������) �������, b.��Ŀ���, Decode(b.��Ŀ���, 4, 'Ѫѹ', b.��¼��) ��Ŀ����, b.��λ," & vbNewLine & _
            "       b.��¼��," & vbNewLine & _
            "       Decode(b.��¼��," & vbNewLine & _
            "               2," & vbNewLine & _
            "               Decode(b.��Ŀ���," & vbNewLine & _
            "                      3," & vbNewLine & _
            "                      6," & vbNewLine & _
            "                      Decode(Decode(c.��Ŀ���, NULL, a.��Ŀ��ʾ, 4)," & vbNewLine & _
            "                             4," & vbNewLine & _
            "                             Decode(Sign(Nvl(b.��¼Ƶ��, 2) - 2), 1, 2, Nvl(b.��¼Ƶ��, 2))," & vbNewLine & _
            "                             Nvl(b.��¼Ƶ��, 2)))," & vbNewLine & _
            "               NULL) ��¼Ƶ��" & vbNewLine & _
            " FROM �����¼��Ŀ a, ���¼�¼��Ŀ b, ��������Ŀ c" & vbNewLine & _
            " WHERE a.��Ŀ��� = b.��Ŀ��� AND a.��Ŀ��� = c.��Ŀ���(+) AND NVL(a.Ӧ�÷�ʽ,0) <> 0 AND a.��Ŀ���� = 1  and A.����ȼ�>=[1]" & vbNewLine & _
            " And nvl(A.���ò���,0) in (0,[2]) and (A.���ÿ���=1 or (A.���ÿ���=2 and Exists (select 1 from �������ÿ��� D where A.��Ŀ���=D.��Ŀ��� and D.����ID=[3])))" & vbNewLine & _
            " ORDER BY Decode(b.��¼��, 2, 2, 1), Decode(b.��Ŀ���, 3, Decode(b.��¼��, 2, 1, b.�������), b.�������)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ", lng����ȼ�, IIf(lngӤ�� = 0, 1, 2), lng����ID)
        '����������Ŀ
        rsTemp.Filter = "��¼��=1 OR ��¼��=3"
        rsTemp.Sort = "��¼��,�������"
        i = 0: str������Ŀ = ""
        With rsTemp
            Do While Not .EOF
                If i = 0 Then
                    str������Ŀ = !��Ŀ���
                Else
                    str������Ŀ = str������Ŀ & "," & !��Ŀ���
                End If
                i = i + 1
                .MoveNext
            Loop
            T_BodyItem.str������Ŀ = str������Ŀ
        End With
        
        '���±����Ŀ
        rsTemp.Filter = "��¼��=2 And ��Ŀ���<>5"
        rsTemp.Sort = "�������"
        i = 0: str�����Ŀ = "": str������� = "": lngTabRows = 0
        With rsTemp
            Do While Not .EOF
                If i = 0 Then
                   If Val(!��Ŀ���) = 4 Then
                       str�����Ŀ = "4,5:" & Nvl(!��¼Ƶ��)
                       str������� = "4,5"
                   Else
                       str�����Ŀ = Nvl(!��Ŀ���) & ":" & Nvl(!��¼Ƶ��)
                       str������� = Nvl(!��Ŀ���)
                   End If
                Else
                   If Val(!��Ŀ���) = 4 Then
                       str�����Ŀ = str�����Ŀ & "@" & "4,5" & ":" & Nvl(!��¼Ƶ��)
                       str������� = str������� & "," & "4,5"
                   Else
                       str�����Ŀ = str�����Ŀ & "@" & Nvl(!��Ŀ���) & ":" & Nvl(!��¼Ƶ��)
                       str������� = str������� & "," & Nvl(!��Ŀ���)
                   End If
                End If
                 '��������ռ�õ�������
                If Val(!��Ŀ���) = 3 Then '˵������Ϊ�����Ŀ
                    lngTabRows = lngTabRows + 1
                Else
                    Select Case Val(Nvl(!��¼Ƶ��, 2))
                    Case 3
                        lngTabRows = lngTabRows + 3
                    Case 4
                        lngTabRows = lngTabRows + 2
                    Case Else
                        lngTabRows = lngTabRows + 1
                    End Select
                End If
                 i = i + 1
                .MoveNext
            Loop
            T_BodyItem.str������� = str�������
            T_BodyItem.str�����Ŀ = GetString(str�����Ŀ)
        End With
        T_BodyStyle.lng������ = Val(zlDatabase.GetPara("���±������", glngSys, 1255, 8)) - lngTabRows
    End If
    Call GetPainDegreeNO
    GetStyleBody = True
    Exit Function
Errhand:
    If ErrCenter() Then
        Resume
    End If
End Function

Private Function GetString(ByVal strValue As String) As String
    Dim strOld() As String
    Dim strNew As String
    Dim strѪѹ As String
    Dim i As Integer
    
    strOld = Split(strValue, "@")
    For i = 0 To UBound(strOld)
        If InStr(strOld(i), ",") > 0 Then
            strѪѹ = Split(strOld(i), ",")(0) & ":" & Split(strOld(i), ":")(1)
            strѪѹ = strѪѹ & "," & Split(strOld(i), ",")(1)
        Else
            If i = 0 Then
                strNew = strOld(i)
            Else
                strNew = strNew & "," & strOld(i)
            End If
        End If
    Next
    If strѪѹ = "" Then
        GetString = strNew
    Else
        GetString = strNew & "," & strѪѹ
    End If
End Function

Private Function GetSymbol(ByVal lng��Ŀ��� As Long, ByVal str��λ As String, Optional ByVal str�ص���Ŀ As String = "��", Optional ByVal str���� As String = "") As Boolean

    'bln��ͼ����=True,���µ���ͼ����,����������ʾ;����,�մ����������ʾ
    Dim blnGraph As Boolean
    Dim bln�ص� As Boolean
    Dim str��¼�� As String

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
    
    If mrsGraph.RecordCount = 0 Then Exit Function
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
    End If

    
    mrsGraph.Filter = ""
    
    If str��¼�� <> "��" Then
        GetSymbol = False
    Else
        GetSymbol = True
    End If

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

