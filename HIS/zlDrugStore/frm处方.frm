VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���� 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picRecInfo_CM 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   7455
      TabIndex        =   45
      Top             =   7560
      Width           =   7455
      Begin VB.Label lblԭʼ���� 
         AutoSize        =   -1  'True
         Caption         =   "ԭʼ������"
         Height          =   180
         Left            =   0
         TabIndex        =   47
         Tag             =   "ԭʼ����:"
         Top             =   60
         Width           =   900
      End
      Begin VB.Label lbl��ҩ�巨 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ�巨��"
         Height          =   180
         Left            =   1830
         TabIndex        =   46
         Tag             =   "��ҩ�巨:"
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox picRecipe 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Index           =   0
      Left            =   240
      ScaleHeight     =   7395
      ScaleWidth      =   8475
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   4
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   8175
         TabIndex        =   30
         Top             =   6480
         Width           =   8175
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "���ڣ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   2
            Left            =   3720
            TabIndex        =   44
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "2009��07��07��"
            Height          =   180
            Index           =   3
            Left            =   4320
            TabIndex        =   39
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ҽʦ��"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   8
            Left            =   3720
            TabIndex        =   38
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��С��"
            Height          =   180
            Index           =   9
            Left            =   4320
            TabIndex        =   37
            Top             =   600
            Width           =   540
         End
         Begin VB.Line lineRP��� 
            Index           =   0
            X1              =   240
            X2              =   7680
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ӧ��/ʵ�պϼƣ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   36
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "1013.45/1010.00Ԫ"
            Height          =   180
            Index           =   1
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�շ�Ա��"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   4
            Left            =   6000
            TabIndex        =   34
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "������"
            Height          =   180
            Index           =   5
            Left            =   6840
            TabIndex        =   33
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ҩ�ˣ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   6
            Left            =   900
            TabIndex        =   32
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblRP��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�Ž���"
            Height          =   180
            Index           =   7
            Left            =   1680
            TabIndex        =   31
            Top             =   600
            Width           =   540
         End
      End
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Index           =   3
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   8175
         TabIndex        =   27
         Top             =   2280
         Width           =   8175
         Begin VSFlex8Ctl.VSFlexGrid vsfRecipe 
            Height          =   1815
            Left            =   480
            TabIndex        =   28
            Top             =   600
            Width           =   2175
            _cx             =   3836
            _cy             =   3201
            Appearance      =   0
            BorderStyle     =   0
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
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
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
            ComboSearch     =   0
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
         Begin VB.Line lineRP���� 
            Index           =   0
            X1              =   240
            X2              =   7680
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblRP���� 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "RP"
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   120
            Width           =   390
         End
      End
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   2
         Left            =   0
         ScaleHeight     =   1575
         ScaleWidth      =   8175
         TabIndex        =   11
         Top             =   720
         Width           =   8175
         Begin VB.TextBox txt������� 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   48
            Text            =   "frm����.frx":0000
            Top             =   1100
            Width           =   6135
         End
         Begin VB.Label lblRPǰ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���أ�"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   21
            Left            =   3240
            TabIndex        =   54
            Tag             =   "���ţ�"
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "55kg"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   3840
            TabIndex        =   53
            Tag             =   "���أ�"
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "1234567890"
            Height          =   180
            Index           =   20
            Left            =   6600
            TabIndex        =   52
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "���￨�ţ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   19
            Left            =   5640
            TabIndex        =   51
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "5"
            Height          =   180
            Index           =   17
            Left            =   5040
            TabIndex        =   43
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����ţ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   4
            Left            =   5820
            TabIndex        =   42
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "���ࣺ"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   41
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�Է�"
            Height          =   180
            Index           =   1
            Left            =   1680
            TabIndex        =   40
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�ѱ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   0
            Left            =   1125
            TabIndex        =   26
            Top             =   0
            Width           =   540
         End
         Begin VB.Line lineRPǰ�� 
            Index           =   0
            X1              =   240
            X2              =   7680
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ҽ��"
            Height          =   180
            Index           =   3
            Left            =   3840
            TabIndex        =   25
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "H00040015"
            Height          =   180
            Index           =   5
            Left            =   6600
            TabIndex        =   24
            Top             =   0
            Width           =   825
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��������"
            Height          =   180
            Index           =   7
            Left            =   1680
            TabIndex        =   23
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   180
            Index           =   9
            Left            =   3840
            TabIndex        =   22
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "40"
            Height          =   180
            Index           =   11
            Left            =   5040
            TabIndex        =   21
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����ڿ�"
            Height          =   180
            Index           =   15
            Left            =   6240
            TabIndex        =   20
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "������"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   6
            Left            =   1125
            TabIndex        =   19
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�Ա�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   8
            Left            =   3240
            TabIndex        =   18
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "���䣺"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   10
            Left            =   4440
            TabIndex        =   17
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "���ң�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   14
            Left            =   5640
            TabIndex        =   16
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "���ţ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   16
            Left            =   4440
            TabIndex        =   15
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʶ�ţ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   12
            Left            =   945
            TabIndex        =   14
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "3434242123213"
            Height          =   180
            Index           =   13
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblRPǰ�� 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�ٴ���ϣ�"
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   18
            Left            =   765
            TabIndex        =   12
            Top             =   1080
            Width           =   900
         End
      End
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   8175
         TabIndex        =   7
         Top             =   0
         Width           =   8175
         Begin VB.PictureBox picRecipe 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   5
            Left            =   6840
            ScaleHeight     =   435
            ScaleWidth      =   945
            TabIndex        =   8
            Top             =   68
            Width           =   975
            Begin VB.Label lblRP��ʶ 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ͨ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   75
               Width           =   720
            End
         End
         Begin VB.Label lblRP���� 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "�����е�������ҽԺ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1080
            TabIndex        =   10
            Top             =   120
            Width           =   5055
         End
      End
   End
   Begin VB.PictureBox picProcess 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   7920
      Width           =   9855
      Begin VB.ComboBox cbo�˲��� 
         Height          =   300
         Left            =   6600
         TabIndex        =   49
         Text            =   "cbo�˲���"
         Top             =   0
         Width           =   1935
      End
      Begin VB.ComboBox cbo��ҩ�� 
         Height          =   300
         Left            =   3720
         TabIndex        =   3
         Text            =   "cbo��ҩ��"
         Top             =   25
         Width           =   1935
      End
      Begin VB.ComboBox cbo����ҽ�� 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   25
         Width           =   1935
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "��ҩ(&S)"
         Height          =   350
         Left            =   8640
         TabIndex        =   1
         ToolTipText     =   "�ȼ���F2"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lbl�˲��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˲���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6000
         TabIndex        =   50
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lbl��ҩ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3120
         TabIndex        =   5
         Top             =   85
         Width           =   540
      End
      Begin VB.Label lbl����ҽ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   85
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�û�����Ĵ�����ɫ����ע���ȡ���ַ�������;�ָ�
Private mstrUserRecipeColor As String

Private mstrDosUser As String
Private mstrPrivs As String
Private mbln��ҩ���� As Boolean
Private mbln��־ As Boolean
Private mstr�˲��� As String

Private Type Type_Condition
    intListType As Integer
    lngҩ��ID As Long
    bln�Զ���ҩ As Boolean
    bln�Ƿ���Ҫ��ҩ���� As Boolean
    blnУ�鴦�� As Boolean
    str��ҩ�� As String
    str�˲��� As String
    int�����ʾ As Integer      '�����ʾ��ʽ��0-��ʾӦ�ս��,1-��ʾʵ�ս��,2-��ʾӦ�պ�ʵ�ս��
    bln������� As Boolean
    intRowNum As Integer
    
End Type
Private mcondition As Type_Condition

'�б�����
Private Enum mListType
    ��ҩȷ�� = 0
    ����ҩ = 1
    ����ҩ = 2
    ����ҩ = 3
    ��ʱδ�� = 4
    ��ҩ = 5
End Enum

'�������ͣ���ͨ�����ơ������������һ������
Private Enum ��������
    ��ͨ = 0
    ���� = 1
    ���� = 2
    ���� = 3
    ��һ = 4
    ���� = 5
End Enum

Private Enum ��ҩ����
    ���� = 0
    ҩƷ = 1
    Ƥ�Խ�� = 2
    ���� = 3
    ���� = 4
    �÷� = 5
    ���� = 6
    ���id = 7
    
    ���� = 8
End Enum

Private Enum ��ҩ����
    ����1 = 0
    ����1��ע = 1
    ���1 = 2
    ����2 = 3
    ����2��ע = 4
    ���2 = 5
    ����3 = 6
    ����3��ע = 7
    ���3 = 8
    ����4 = 9
    ����4��ע = 10
    
    ���� = 8
End Enum

Private Enum ����ǩ����
    ���� = 0
    ���� = 1
    ǰ�� = 2
    ���� = 3
    ��� = 4
    ��ʶ = 5
End Enum

Private Enum RP����
    ҽԺ���� = 0
End Enum

Private Enum RP��ʶ
    ��ʶ = 0
End Enum

Private Enum RPǰ��
    �ѱ��ǩ = 0
    �ѱ� = 1
    
    �����ǩ = 2
    ���� = 3
    
    �����ű�ǩ = 4
    ������ = 5
    
    ������ǩ = 6
    ���� = 7
    
    �Ա��ǩ = 8
    �Ա� = 9
    
    �����ǩ = 10
    ���� = 11
    
    ��ʶ�ű�ǩ = 12
    ��ʶ�� = 13
    
    ���ұ�ǩ = 14
    ���� = 15
    
    ���ű�ǩ = 16
    ���� = 17
    
    �ٴ���ϱ�ǩ = 18
    
    ���￨�ű�ǩ = 19
    ���￨�� = 20
    
    ���ر�ǩ = 21
    ���� = 22
End Enum

Private Enum RP����
    ��ʶ = 0
End Enum

Private Enum RP���
    �ϼƽ���ǩ = 0
    �ϼƽ�� = 1
    
    ���ڱ�ǩ = 2
    ���� = 3
    
    �շ�Ա��ǩ = 4
    �շ�Ա = 5
    
    ��ҩ�˱�ǩ = 6
    ��ҩ�� = 7
    
    ����ҽ����ǩ = 8
    ����ҽ�� = 9
End Enum

Public Sub CmdProcess()
    If CmdSend.Enabled Then CmdSend_Click
End Sub
Public Sub FormClear()
    '����
    lblRP����(RP����.ҽԺ����).Caption = GetUnitName
    
    '��ʶ
    lblRP��ʶ(RP��ʶ.��ʶ).Caption = "��ͨ"
    
    'ǰ��
    lblRPǰ��(RPǰ��.�ѱ�).Caption = ""
    lblRPǰ��(RPǰ��.����).Caption = ""
    lblRPǰ��(RPǰ��.������).Caption = ""
    lblRPǰ��(RPǰ��.����).Caption = ""
    lblRPǰ��(RPǰ��.�Ա�).Caption = ""
    lblRPǰ��(RPǰ��.����).Caption = ""
    lblRPǰ��(RPǰ��.���￨��).Caption = ""
    lblRPǰ��(RPǰ��.��ʶ��).Caption = ""
    lblRPǰ��(RPǰ��.����).Caption = ""
    lblRPǰ��(RPǰ��.����).Caption = ""
    lblRPǰ��(RPǰ��.����).Caption = ""
    txt�������.Text = ""
    txt�������.Tag = ""
    
    '����
    vsfRecipe.rows = 1
    
    '���
    lblRP���(RP���.�ϼƽ��).Caption = ""
    lblRP���(RP���.����).Caption = ""
    lblRP���(RP���.�շ�Ա).Caption = ""
    lblRP���(RP���.��ҩ��).Caption = ""
    lblRP���(RP���.����ҽ��).Caption = ""
    
    '������ɫ
    SetRecipeColor 0
    
    CmdSend.Enabled = False
End Sub


Private Sub Loadҽ��()
    Dim rsData As ADODB.Recordset
    
    Set rsData = RecipeSendWork_Getҽ��
    
    Me.cbo����ҽ��.Clear
    cbo����ҽ��.AddItem ""
    Do While Not rsData.EOF
        cbo����ҽ��.AddItem rsData!ҽ��
        rsData.MoveNext
    Loop
    cbo����ҽ��.ListIndex = 0
End Sub
Public Sub SetParams()
    Dim bln�Ƿ���ҩȷ�� As Boolean

    mstrUserRecipeColor = zldatabase.GetPara("������ɫ", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
    
    With mcondition
        If .lngҩ��ID <> Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1341)) Then
            .lngҩ��ID = Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1341))
            .bln�Ƿ���Ҫ��ҩ���� = RecipeSendWork_DispensingMedi(.lngҩ��ID, bln�Ƿ���ҩȷ��)
            Call Load��ҩ��(.lngҩ��ID)
        End If
        
        .str��ҩ�� = zldatabase.GetPara("��ҩ��", glngSys, 1341)
        .str�˲��� = zldatabase.GetPara("�˲���", glngSys, 1341)
        
        If .str��ҩ�� = "|��ǰ����Ա|" Then
            mstrDosUser = gstrUserName
        Else
            mstrDosUser = .str��ҩ��
        End If
        
        If .str�˲��� = "|��ǰ����Ա|" Then
            mstr�˲��� = gstrUserName
        Else
            mstr�˲��� = .str�˲���
        End If
    
        .bln�Զ���ҩ = (Val(zldatabase.GetPara("�Զ���ҩ", glngSys, 1341)) = 1)
        .int�����ʾ = Val(zldatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0))
        .bln������� = ((gtype_UserSysParms.P240_ҩ��������� = 1 Or gtype_UserSysParms.P240_ҩ��������� = 3) And gtype_UserSysParms.P241_�������ʱ�� = 2)
        .intRowNum = gtype_UserSysParms.P213_��ҩ�䷽ÿ����ҩζ��
        
        
        If zlStr.IsHavePrivs(mstrPrivs, "��ҩ") = True Then
            If .bln�Զ���ҩ = False Then
                Cbo��ҩ��.Enabled = True
            Else
                Cbo��ҩ��.Enabled = False
            End If
        Else
            Cbo��ҩ��.Enabled = False
        End If
        
        cbo�˲���.Enabled = True
        .blnУ�鴦�� = IsInString(gstrprivs, "У�鴦��", ";")
        
        Call Load��ҩ��(.lngҩ��ID)
        
        Call Load�˲���(.lngҩ��ID)
    End With
End Sub

Private Sub SetRecipeMedi(ByVal rsData As ADODB.Recordset)
    Dim intRow As Integer
    Dim n As Integer
    Dim lng�������ID As Long
    Dim lng�������ID As Long
    Dim lng�������ID As Long
    Dim blnƤ�� As Boolean
    Dim i As Integer
    Dim lngҩ��id As Long
    Dim dblAmount As Double
    Dim strDiag As String
    Dim int���� As Integer
    Dim intCol As Integer
    Dim dateCurrent As Date
    
    dateCurrent = Sys.Currentdate
    
    rsData.Filter = ""
    rsData.Sort = "���ID,���"
    
    With vsfRecipe
        .Redraw = flexRDNone
        
        Do While Not rsData.EOF
            If rsData!��¼���� = 1 Or (rsData!��¼���� = 2 And (rsData!�����־ = 1 Or rsData!�����־ = 4)) Then
                int���� = 1
            Else
                int���� = 2
            End If
            strDiag = RecipeSendWork_GetDiagnosis(int����, IIf(int���� = 1, Val(rsData!���id), Val(rsData!����ID)), Val(rsData!��ҳid), IIf(mbln��ҩ����, 1, 2))
            If int���� = 1 And rsData!��Ժ And strDiag = "" Then
                int���� = 2
                strDiag = RecipeSendWork_GetDiagnosis(int����, IIf(int���� = 1, Val(rsData!���id), Val(rsData!����ID)), Val(rsData!��ҳid), IIf(mbln��ҩ����, 1, 2))
            End If
            
            If strDiag <> "" Then
                strDiag = strDiag & "|"
                For i = 0 To UBound(Split(strDiag, "|"))
                    If Split(strDiag, "|")(i) <> "" Then
                        If InStr(1, txt�������.Text & " ��", "��" & Split(strDiag, "|")(i) & " ��") < 1 Then
                            txt�������.Text = IIf(txt�������.Text = "", " ��", txt�������.Text & " ��") & Split(strDiag, "|")(i)
                            txt�������.Tag = IIf(txt�������.Tag = "", "�� ", txt�������.Tag & vbCrLf & "�� ") & Split(strDiag, "|")(i)
                        End If
                    End If
                Next
            End If
        
            If mbln��ҩ���� Then
                .MergeCells = flexMergeRestrictColumns
                .MergeCol(��ҩ����.����1) = True
                .MergeCol(��ҩ����.����2) = True
                .MergeCol(��ҩ����.����3) = True
                If mcondition.intRowNum = 4 Then .MergeCol(��ҩ����.����4) = True

                If rsData!ҩ��ID <> lngҩ��id Then
                    If intCol = mcondition.intRowNum Then
                        intCol = 0
                        intRow = intRow + 2
                    ElseIf intCol = 0 Then
                        intRow = intRow + 2
                    End If
                    .rows = intRow + 1
                    
                    lngҩ��id = rsData!ҩ��ID
                    
                    If NVL(rsData!����, 0) = 0 Then
                        dblAmount = rsData!���� * rsData!���� * rsData!��װ * rsData!����ϵ�� / NVL(rsData!ԭʼ����, 1)
                    Else
                        dblAmount = rsData!����
                    End If
                Else
                    intCol = intCol - 1
                    
                    If NVL(rsData!����, 0) = 0 Then
                        dblAmount = dblAmount + rsData!���� * rsData!���� * rsData!��װ * rsData!����ϵ�� / NVL(rsData!ԭʼ����, 1)
                    Else
                        dblAmount = rsData!����
                    End If
                End If
                
                If intCol = 0 Then
                    .TextMatrix(intRow, ��ҩ����.����1) = rsData!ҩ��
                    .TextMatrix(intRow, ��ҩ����.����1��ע) = FormatEx(Abs(dblAmount), 1) & rsData!���㵥λ
                    .TextMatrix(intRow - 1, ��ҩ����.����1) = rsData!ҩ��
                    .TextMatrix(intRow - 1, ��ҩ����.����1��ע) = IIf(IsNull(rsData!ҽ������), "", "(" & rsData!ҽ������ & ")")
                ElseIf intCol = 1 Then
                    .TextMatrix(intRow, ��ҩ����.����2) = rsData!ҩ��
                    .TextMatrix(intRow, ��ҩ����.����2��ע) = FormatEx(Abs(dblAmount), 1) & rsData!���㵥λ
                    .TextMatrix(intRow - 1, ��ҩ����.����2) = rsData!ҩ��
                    .TextMatrix(intRow - 1, ��ҩ����.����2��ע) = IIf(IsNull(rsData!ҽ������), "", "(" & rsData!ҽ������ & ")")
                ElseIf intCol = 2 Then
                    .TextMatrix(intRow, ��ҩ����.����3) = rsData!ҩ��
                    .TextMatrix(intRow, ��ҩ����.����3��ע) = FormatEx(Abs(dblAmount), 1) & rsData!���㵥λ
                    .TextMatrix(intRow - 1, ��ҩ����.����3) = rsData!ҩ��
                    .TextMatrix(intRow - 1, ��ҩ����.����3��ע) = IIf(IsNull(rsData!ҽ������), "", "(" & rsData!ҽ������ & ")")
                ElseIf intCol = 3 Then
                    .TextMatrix(intRow, ��ҩ����.����4) = rsData!ҩ��
                    .TextMatrix(intRow, ��ҩ����.����4��ע) = FormatEx(Abs(dblAmount), 1) & rsData!���㵥λ
                    .TextMatrix(intRow - 1, ��ҩ����.����4) = rsData!ҩ��
                    .TextMatrix(intRow - 1, ��ҩ����.����4��ע) = IIf(IsNull(rsData!ҽ������), "", "(" & rsData!ҽ������ & ")")

                End If
        
                intCol = intCol + 1
                
                .RowHeight(intRow) = 250
                .RowHeight(intRow - 1) = 250

            Else
                intRow = intRow + 1
                .rows = intRow + 1
                
                dblAmount = rsData!����
                
                .TextMatrix(intRow, ��ҩ����.ҩƷ) = rsData!ҩƷ���� & vbCrLf & rsData!ҩƷ���
                
                If rsData!�Ƿ�Ƥ�� = 1 Then
                    .TextMatrix(intRow, ��ҩ����.Ƥ�Խ��) = GetƤ�Խ��(rsData!����ID, rsData!ҩ��ID, dateCurrent, rsData!����ʱ��)
                    If .TextMatrix(intRow, ��ҩ����.Ƥ�Խ��) <> "" Then
                        blnƤ�� = True
                    End If
                End If
                
                .TextMatrix(intRow, ��ҩ����.����) = zlStr.FormatEx(dblAmount, 5) & rsData!��λ
                
                .ColWidth(��ҩ����.����) = 750
                .TextMatrix(intRow, ��ҩ����.����) = IIf(IsNull(rsData!����), "", zlStr.FormatEx(rsData!����, 5) & "(" & zlStr.NVL(rsData!���㵥λ) & ")")
                
                .TextMatrix(intRow, ��ҩ����.�÷�) = IIf(IsNull(rsData!�÷�), "", rsData!�÷�) & " " & IIf(IsNull(rsData!Ƶ��), "", rsData!Ƶ��)
                .TextMatrix(intRow, ��ҩ����.����) = IIf(IsNull(rsData!ҽ������), "", rsData!ҽ������)
                .TextMatrix(intRow, ��ҩ����.���id) = Val(rsData!���id)
                
                'Ĭ������+��������д�����д�ڵڶ��У�������Ƴ����п������Ӷ����������Ĭ��ÿ���ָ�250��ÿ�ֿ�200
                .RowHeight(intRow) = 250 * ((-1 * Int(-1 * Len(rsData!ҩƷ����) / Int((.ColWidth(��ҩ����.ҩƷ) / 200)))) + 1)
            End If
            
            rsData.MoveNext
            
            '�����һ��ҩƷ����һ��ҩƷ������һ����
            If Not rsData.EOF And Not mbln��ҩ���� Then
                If Val(rsData!���id) = 0 Then
                    intRow = intRow + 1
                    .rows = intRow + 1
                    .RowHeight(intRow) = 200
                End If
            End If
        Loop
        
        If Not mbln��ҩ���� Then
            '���÷���
            For n = 1 To .rows - 1
                If Val(.TextMatrix(n, ��ҩ����.���id)) <> 0 Then
                    lng�������ID = .TextMatrix(n, ��ҩ����.���id)
                    If n + 1 <= .rows - 1 Then
                        If Val(.TextMatrix(n + 1, ��ҩ����.���id)) <> 0 Then    '�������Ϊ��¼��ʱ
                            lng�������ID = IIf(.TextMatrix(n + 1, ��ҩ����.���id) = 0, -1, .TextMatrix(n + 1, ��ҩ����.���id))
                        ElseIf n + 2 <= .rows - 1 Then  '�������Ϊ��������ʱ
                            If Val(.TextMatrix(n + 2, ��ҩ����.���id)) <> 0 Then    '���������Ϊ��¼��ʱ
                                lng�������ID = IIf(Val(.TextMatrix(n + 2, ��ҩ����.���id)) = 0, -1, Val(.TextMatrix(n + 2, ��ҩ����.���id)))
                            Else
                                lng�������ID = -1
                            End If
                        Else
                            lng�������ID = -1
                        End If
                    Else
                        lng�������ID = -1
                    End If
                    
                    If lng�������ID = lng�������ID Then
                        If lng�������ID = lng�������ID Then
                            .TextMatrix(n, ��ҩ����.����) = "��"
                        Else
                            .TextMatrix(n, ��ҩ����.����) = "��"
                        End If
                    ElseIf lng�������ID = lng�������ID Then
                        .TextMatrix(n, ��ҩ����.����) = "��"
                    End If
                
                    lng�������ID = IIf(lng�������ID = 0, -1, lng�������ID)
                End If
            Next
            
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(��ҩ����.�÷�) = False
            
            '����Ƥ��
            If blnƤ�� = True Then
                .ColWidth(��ҩ����.Ƥ�Խ��) = 800
                For i = 1 To .rows - 1
                    If .TextMatrix(i, ��ҩ����.Ƥ�Խ��) = "(+)" Then
                        .Cell(flexcpForeColor, i, ��ҩ����.Ƥ�Խ��, i, ��ҩ����.Ƥ�Խ��) = vbRed
                    ElseIf .TextMatrix(i, ��ҩ����.Ƥ�Խ��) = "(-)" Then
                        .Cell(flexcpForeColor, i, ��ҩ����.Ƥ�Խ��, i, ��ҩ����.Ƥ�Խ��) = vbBlue
                    Else
                        .Cell(flexcpForeColor, i, ��ҩ����.Ƥ�Խ��, i, ��ҩ����.Ƥ�Խ��) = &H80000008
                    End If
                Next
            Else
                .ColWidth(��ҩ����.Ƥ�Խ��) = 0
            End If
        End If
        
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub IniRecipe()
    
    With vsfRecipe
        .rows = 1
        
        'ҩƷ������񣩣��������������÷������У����ID
        
        If Not mbln��ҩ���� Then
            .Cols = ��ҩ����.����
            .ColWidth(0) = 500
            .ColWidth(��ҩ����.ҩƷ) = 2000
            .ColWidth(��ҩ����.Ƥ�Խ��) = 400
            .ColWidth(��ҩ����.����) = 750
            
            
            .ColWidth(��ҩ����.����) = 750
            .ColWidth(��ҩ����.�÷�) = 2000
            .ColWidth(��ҩ����.����) = 1500
            .ColWidth(��ҩ����.���id) = 0
            
            .FixedAlignment(��ҩ����.ҩƷ) = flexAlignCenterCenter
            .FixedAlignment(��ҩ����.Ƥ�Խ��) = flexAlignCenterCenter
            .FixedAlignment(��ҩ����.����) = flexAlignCenterCenter
            .FixedAlignment(��ҩ����.����) = flexAlignCenterCenter
            .FixedAlignment(��ҩ����.�÷�) = flexAlignCenterCenter
            .FixedAlignment(��ҩ����.����) = flexAlignCenterCenter
            
            .TextMatrix(0, ��ҩ����.ҩƷ) = "ҩƷ"
            .TextMatrix(0, ��ҩ����.Ƥ�Խ��) = ""
            .TextMatrix(0, ��ҩ����.����) = "����"
            .TextMatrix(0, ��ҩ����.����) = "����"
            .TextMatrix(0, ��ҩ����.�÷�) = "�÷�"
            .TextMatrix(0, ��ҩ����.����) = "����"
            
            .ColAlignment(��ҩ����.����) = flexAlignRightCenter
            .ColAlignment(��ҩ����.ҩƷ) = flexAlignLeftCenter
            .ColAlignment(��ҩ����.Ƥ�Խ��) = flexAlignLeftCenter
            .ColAlignment(��ҩ����.����) = flexAlignCenterCenter
            .ColAlignment(��ҩ����.����) = flexAlignCenterCenter
            .ColAlignment(��ҩ����.�÷�) = flexAlignLeftCenter
            .ColAlignment(��ҩ����.����) = flexAlignLeftCenter
            
            .RowHeight(0) = 255
        Else
            .Cols = ��ҩ����.���� + IIf(mcondition.intRowNum = 4, 3, 0)
            
            If mcondition.intRowNum = 4 Then
                .ColWidth(��ҩ����.����1) = 1100
                .ColWidth(��ҩ����.����2) = 1100
                .ColWidth(��ҩ����.����3) = 1100
                .ColWidth(��ҩ����.����4) = 1100
                .ColWidth(��ҩ����.���1) = 50
                .ColWidth(��ҩ����.���2) = 50
                .ColWidth(��ҩ����.���3) = 50
                .ColWidth(��ҩ����.����1��ע) = 750
                .ColWidth(��ҩ����.����2��ע) = 750
                .ColWidth(��ҩ����.����3��ע) = 750
                .ColWidth(��ҩ����.����4��ע) = 750
                
                .ColAlignment(��ҩ����.����4) = flexAlignRightCenter
                .TextMatrix(0, ��ҩ����.����4) = ""
            Else
                .ColWidth(��ҩ����.����1) = 1700
                .ColWidth(��ҩ����.����2) = 1700
                .ColWidth(��ҩ����.����3) = 1700
                .ColWidth(��ҩ����.���1) = 50
                .ColWidth(��ҩ����.���2) = 50
                .ColWidth(��ҩ����.����1��ע) = 750
                .ColWidth(��ҩ����.����2��ע) = 750
                .ColWidth(��ҩ����.����3��ע) = 750
            End If
            .ColAlignment(��ҩ����.����1) = flexAlignRightCenter
            .ColAlignment(��ҩ����.����2) = flexAlignRightCenter
            .ColAlignment(��ҩ����.����3) = flexAlignRightCenter
            .ColAlignment(��ҩ����.����1��ע) = flexAlignLeftCenter
            .ColAlignment(��ҩ����.����2��ע) = flexAlignLeftCenter
            .ColAlignment(��ҩ����.����3��ע) = flexAlignLeftCenter
            
            .TextMatrix(0, ��ҩ����.����1) = ""
            .TextMatrix(0, ��ҩ����.����2) = ""
            .TextMatrix(0, ��ҩ����.����3) = ""
            .TextMatrix(0, ��ҩ����.����1��ע) = ""
            .TextMatrix(0, ��ҩ����.����2��ע) = ""
            .TextMatrix(0, ��ҩ����.����3��ע) = ""
            
            .RowHidden(0) = 0
        End If
        
    End With
End Sub
Public Sub ShowRecipe(ByVal intType As Integer)
    Dim i As Integer
    
    With mcondition
        .intListType = intType
        
        If .intListType = mListType.����ҩ Or .intListType = mListType.��ʱδ�� Then
            cbo����ҽ��.Enabled = True
        Else
            cbo����ҽ��.Enabled = False
        End If
        
        If .intListType <> mListType.��ҩ Then
            For i = 0 To Cbo��ҩ��.ListCount - 1
                If mstrDosUser = Cbo��ҩ��.List(i) Then
                    Cbo��ҩ��.ListIndex = i
                    Exit For
                End If
            Next
            
            For i = 0 To cbo�˲���.ListCount - 1
                If mstr�˲��� = cbo�˲���.List(i) Then
                    cbo�˲���.ListIndex = i
                    Exit For
                End If
            Next
            Lbl��ҩ��.Caption = "��ҩ��"
        Else
            Lbl��ҩ��.Caption = "��ҩ��"
        End If
        
'        cbo��ҩ��.Enabled = (.intListType <> mListType.��ҩ)
        If zlStr.IsHavePrivs(mstrPrivs, "��ҩ") = True Then
            If .bln�Զ���ҩ = False Then
                Cbo��ҩ��.Enabled = True
            Else
                Cbo��ҩ��.Enabled = False
            End If
        Else
            Cbo��ҩ��.Enabled = False
        End If

        cbo�˲���.Enabled = True
        If .intListType = mListType.��ҩ Then
            Cbo��ҩ��.Enabled = False
            cbo�˲���.Enabled = False
        End If
        
        Select Case .intListType
            Case mListType.��ҩȷ��
                Me.cbo����ҽ��.Enabled = False
                Me.Cbo��ҩ��.Enabled = False
                Me.cbo�˲���.Enabled = False
                CmdSend.Caption = "��ҩȷ��(&O)"
            Case mListType.����ҩ
                CmdSend.Caption = "��ҩ(&V)"
            Case mListType.����ҩ
                CmdSend.Caption = "ȡ����ҩ(&C)"
            Case mListType.����ҩ, mListType.��ʱδ��
                CmdSend.Caption = "��ҩ(&S)"
        End Select
        
        CmdSend.Visible = (.intListType <> mListType.��ҩ)
    End With
    
    SetCmdSendPrivs intType
End Sub

Private Sub SetCmdSendPrivs(ByVal int����� As Integer)
    'Ȩ�޿���
    Select Case mcondition.intListType
    Case mListType.��ҩȷ��
       '��ҩȷ��
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��ҩȷ��")
    Case mListType.����ҩ
        '��ҩ
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "��ҩ") And mcondition.bln�Զ���ҩ = False And (mcondition.bln������� = False Or (mcondition.bln������� = True And int����� = 1)))
    Case mListType.����ҩ
        'ȡ��
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��ҩ")
    Case mListType.����ҩ, mListType.��ʱδ��
        '��ҩ
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "��ҩ") And (mcondition.bln������� = False Or (mcondition.bln������� = True And int����� = 1)))
    End Select
End Sub


Private Sub cbo�˲���_Click()
    Dim i As Integer
    
    If mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ�� Then
        mstr�˲��� = Me.cbo�˲���.Text
    End If
End Sub

Private Sub cbo��ҩ��_Click()
    Dim i As Integer
    
    If mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ�� Then
        mstrDosUser = Me.Cbo��ҩ��.Text
    End If
End Sub

Private Sub CmdSend_Click()
    If frmҩƷ������ҩNew.RecipeWork(mcondition.intListType, frm������ҩ��ϸ.mblnInput, frm������ҩ��ϸ.vsfList) = False Then
        FormClear
    End If
End Sub

Public Function Get��ҩ��() As String
    If Cbo��ҩ��.ListIndex = -1 Then
        Get��ҩ�� = ""
    ElseIf InStr(Cbo��ҩ��.Text, "-") > 0 Then
        Get��ҩ�� = Mid(Cbo��ҩ��.Text, InStr(Cbo��ҩ��.Text, "-") + 1)
    Else
        Get��ҩ�� = Cbo��ҩ��.Text
    End If
End Function
Public Function Get�˲���() As String
    If cbo�˲���.ListIndex = -1 Then
        Get�˲��� = ""
    ElseIf InStr(cbo�˲���.Text, "-") > 0 Then
        Get�˲��� = Mid(cbo�˲���.Text, InStr(cbo�˲���.Text, "-") + 1)
    Else
        Get�˲��� = cbo�˲���.Text
    End If
End Function
Public Function Get����ҽ��() As String
    If cbo����ҽ��.ListIndex = -1 Then
        Get����ҽ�� = ""
    ElseIf InStr(cbo����ҽ��.Text, "-") > 0 Then
        Get����ҽ�� = Mid(cbo����ҽ��.Text, InStr(cbo����ҽ��.Text, "-") + 1)
    Else
        Get����ҽ�� = cbo����ҽ��.Text
    End If
End Function
Private Sub Form_Load()
    mstrPrivs = gstrprivs
    
    Call SetParams
    Call Loadҽ��
    
    Call IniRecipe
    Call FormClear
    
    If InStr(1, mstrPrivs, "ҽ����ѯ") = 0 Then
        lblRP���(9).Visible = False
    End If
End Sub

Private Sub Load��ҩ��(ByVal lngҩ��ID As Long)
    '��ҩ��
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = " Select ����||'-'||���� As ����,���� As ���� From ��Ա��  Where ID in " & _
             " (Select Distinct ��ԱID From ��Ա����˵�� Where ��Ա����='ҩ����ҩ��' " & _
             " And ��ԱID IN (Select ��ԱID From ������Ա Where ����ID=[1]))" & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩ��", lngҩ��ID)
    
    With rsData
        Me.Cbo��ҩ��.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            Cbo��ҩ��.AddItem !����
            
            If mstrDosUser = !���� Then
                intIndex = .AbsolutePosition - 1
            End If
            
            .MoveNext
        Loop
        
        Cbo��ҩ��.Enabled = Not Cbo��ҩ��.ListCount = 0
        
        If intIndex <> -1 Then Cbo��ҩ��.ListIndex = intIndex
        
        mstrDosUser = Me.Cbo��ҩ��.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Load�˲���(ByVal lngҩ��ID As Long)
    '�˲���
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select ����||'-'||���� As ����,���� As ���� From ��Ա�� Where Id In (Select ��Աid from ������Ա Where ����id=[1]) " & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��˴�����", lngҩ��ID)
    
    With rsData
        Me.cbo�˲���.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            cbo�˲���.AddItem !����
            
            If mstr�˲��� = !���� Then
                intIndex = .AbsolutePosition - 1
            End If
            
            .MoveNext
        Loop
        
        cbo�˲���.Enabled = Not cbo�˲���.ListCount = 0
        
        If intIndex <> -1 Then cbo�˲���.ListIndex = intIndex
        
        mstr�˲��� = Me.cbo�˲���.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    With picRecipe(����ǩ����.����)
        .Left = (Me.Width - .Width) / 2
        .Height = Me.Height - IIf(picRecInfo_CM.Visible, picRecInfo_CM.Height, 0) - picProcess.Height - 200
    End With
    
    With picRecInfo_CM
'        If .Visible Then
            .Top = picRecipe(����ǩ����.����).Top + picRecipe(����ǩ����.����).Height + 100
            .Left = picRecipe(����ǩ����.����).Left
            .Width = picRecipe(����ǩ����.����).Width
'        End If
    End With
    
    With picProcess
        .Left = picRecipe(����ǩ����.����).Left
        .Top = Me.Height - .Height - 50
'        .Width = picRecipe(����ǩ����.����).Width
    End With
    
'    If Me.lbl����ҽ��.Visible = False And mbln��־ = False Then
'        Me.lbl��ҩ��.Left = Me.lbl��ҩ��.Left - 2200
'        Me.cbo��ҩ��.Left = Me.cbo��ҩ��.Left - 2200
'        Me.lbl�˲���.Left = Me.lbl�˲���.Left - 2200
'        Me.cbo�˲���.Left = Me.cbo�˲���.Left - 2200
'    End If
    
    mbln��־ = True
End Sub


Private Sub SetRecipeColor(index As Integer)
    Dim lngBackColor As Long
    Dim objTmp As Object
    Dim strTypeName As String
    
    Select Case index
        Case 0
            lblRP��ʶ(RP��ʶ.��ʶ).Caption = "��ͨ"
        Case 1
            lblRP��ʶ(RP��ʶ.��ʶ).Caption = "����"
        Case 2
            lblRP��ʶ(RP��ʶ.��ʶ).Caption = "����"
        Case 3
            lblRP��ʶ(RP��ʶ.��ʶ).Caption = "����"
        Case 4
            lblRP��ʶ(RP��ʶ.��ʶ).Caption = "��һ"
        Case 5
            lblRP��ʶ(RP��ʶ.��ʶ).Caption = "��"
    End Select
    
    lngBackColor = Val(Split(mstrUserRecipeColor, ";")(index))
    
    For Each objTmp In lblRP����
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRPǰ��
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP����
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP���
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP��ʶ
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In picRecipe
        objTmp.BackColor = lngBackColor
    Next
    
    With vsfRecipe
        .BackColorFixed = lngBackColor
        .BackColor = lngBackColor
        .BackColorBkg = lngBackColor
    End With
    
    txt�������.BackColor = lngBackColor
End Sub


Private Sub picProcess_Resize()
    With CmdSend
        .Left = picProcess.Width - .Width - 100
    End With
End Sub


Private Sub picRecipe_Resize(index As Integer)
    On Error Resume Next
    
    If index = 0 Then
        With lblRP����(RP����.ҽԺ����)
            .Left = 0
            .Width = picRecipe(����ǩ����.����).Width
        End With
        
        With picRecipe(����ǩ����.���)
            .Top = picRecipe(����ǩ����.����).Height - .Height
        End With
        
        With picRecipe(����ǩ����.����)
            .Height = picRecipe(����ǩ����.���).Top - .Top
        End With
        
        With vsfRecipe
            .Left = lblRP����(RP����.��ʶ).Left
            .Top = lblRP����(RP����.��ʶ).Top + lblRP����(RP����.��ʶ).Height + 150
            .Width = picRecipe(����ǩ����.����).Width - .Left - 100
            .Height = picRecipe(����ǩ����.����).Height - .Top - 100
        End With
    ElseIf index = ����ǩ����.ǰ�� Then
        picRecipe(����ǩ����.ǰ��).Height = txt�������.Top + txt�������.Height + 50
        
        With picRecipe(����ǩ����.����)
            .Top = picRecipe(����ǩ����.ǰ��).Top + picRecipe(����ǩ����.ǰ��).Height + 50
            .Height = picRecipe(����ǩ����.���).Top - .Top
        End With
    End If
End Sub
Public Sub RefreshRecipe(ByVal rsData As ADODB.Recordset, ByVal strWeight As String, Optional ByVal int�ɲ��� As Integer = 0, Optional int�Ŷ�״̬ As Integer, Optional int����� As Integer)
    Dim dblӦ�ս��, dblʵ�ս�� As Double
    Dim str����Ա As String
    Dim IntLocate As Integer
    Dim strDiag As String
    Dim int���� As Integer
    Dim i As Integer
 
    FormClear
    
    CmdSend.Enabled = False
    
    With rsData
        .Filter = ""
        
        If .EOF Then Exit Sub
        
        
        mbln��ҩ���� = False
        If �ж��Ƿ���ҩ����(!ҩ��ID, !����, !NO) Then
            mbln��ҩ���� = True
        End If
            
        Call IniRecipe
        
        If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
            int���� = 1
        Else
            int���� = 2
        End If
        
        '����
        lblRP����(RP����.ҽԺ����).Caption = GetUnitName
        
        '��ʶ
        lblRP��ʶ(RP��ʶ.��ʶ).Caption = Split(gconstrRecipeType, ";")(Val(!��������))
        
        'ǰ��
        lblRPǰ��(RPǰ��.�ѱ�).Caption = IIf(IsNull(!�ѱ�), "", !�ѱ�)
'        lblRPǰ��(RPǰ��.����).Caption = IIf(IsNull(!����), "", !����)
        lblRPǰ��(RPǰ��.������).Caption = !NO
        lblRPǰ��(RPǰ��.����).Caption = IIf(IsNull(!����), "", !����)
        lblRPǰ��(RPǰ��.����).ForeColor = zldatabase.GetPatiColor(IIf(IsNull(!��������), "", !��������))
        
        lblRPǰ��(RPǰ��.�Ա�).Caption = IIf(IsNull(!�Ա�), "", !�Ա�)
        lblRPǰ��(RPǰ��.����).Caption = IIf(IsNull(!����), "", !����)
        
        lblRPǰ��(RPǰ��.���￨��).Caption = IIf(IsNull(!���￨��), "", !���￨��)
        
        If !�����־ = 1 Or !�����־ = 4 Then
            lblRPǰ��(RPǰ��.��ʶ�ű�ǩ).Caption = "����ţ�"
        Else
            lblRPǰ��(RPǰ��.��ʶ�ű�ǩ).Caption = "סԺ�ţ�"
        End If
        
        lblRPǰ��(RPǰ��.��ʶ��).Caption = IIf(IsNull(!סԺ��), "", !סԺ��)
        
        lblRPǰ��(RPǰ��.����).Caption = IIf(IsNull(!����), "", !����)
        lblRPǰ��(RPǰ��.����).Caption = IIf(IsNull(!����), "", !����)
        lblRPǰ��(RPǰ��.����).Caption = IIf(IsNumeric(strWeight), strWeight & "kg", strWeight)
        
        '�����Ϣ
        txt�������.Text = ""
        txt�������.Tag = ""
        txt�������.Height = 180
        
        Call picRecipe_Resize(����ǩ����.ǰ��)

        '����
        SetRecipeMedi rsData
        
        '���
        .Filter = ""
        Do While Not .EOF
            dblӦ�ս�� = dblӦ�ս�� + Val(!���۽��)
            dblʵ�ս�� = dblʵ�ս�� + Val(!ʵ�ս��)
            .MoveNext
        Loop
        .MoveFirst
        
        If mcondition.int�����ʾ = 1 Then
            lblRP���(RP���.�ϼƽ���ǩ).Caption = "ʵ�պϼƣ�"
            lblRP���(RP���.�ϼƽ��).Caption = zlStr.FormatEx(dblʵ�ս��, 2, , True) & "Ԫ"
        ElseIf mcondition.int�����ʾ = 2 Then
            lblRP���(RP���.�ϼƽ���ǩ).Caption = "Ӧ��/ʵ�պϼƣ�"
            lblRP���(RP���.�ϼƽ��).Caption = zlStr.FormatEx(dblӦ�ս��, 2, , True) & "Ԫ/" & zlStr.FormatEx(dblʵ�ս��, 2, , True) & "Ԫ"
        Else
            lblRP���(RP���.�ϼƽ���ǩ).Caption = "Ӧ�պϼƣ�"
            lblRP���(RP���.�ϼƽ��).Caption = zlStr.FormatEx(dblӦ�ս��, 2) & "Ԫ"
        End If
        lblRP���(RP���.�ϼƽ���ǩ).Left = lblRP���(RP���.��ҩ�˱�ǩ).Left - (lblRP���(RP���.�ϼƽ���ǩ).Width - lblRP���(RP���.��ҩ�˱�ǩ).Width)
        
        lblRP���(RP���.����).Caption = IIf(IsNull(!��������), "", Format(!��������, "yyyy-mm-dd"))
        lblRP���(RP���.�շ�Ա).Caption = IIf(IsNull(!����Ա����), "", !����Ա����)
        lblRP���(RP���.��ҩ��).Caption = IIf(IsNull(!��ҩ��), "", !��ҩ��)
        lblRP���(RP���.����ҽ��).Caption = IIf(IsNull(!������), "", !������)
                
        '���ô�����ɫ
        SetRecipeColor Val(!��������)
        
        '���ÿ���ҽ��
        Me.cbo����ҽ��.ListIndex = 0
        If (mcondition.blnУ�鴦�� = False) And zlStr.IsHavePrivs(gstrprivs, "ҽ����ѯ") Then
            str����Ա = IIf(IsNull(!������), "", !������)
        Else
            If mcondition.intListType = mListType.��ҩ And mcondition.blnУ�鴦�� = True Then
                str����Ա = IIf(IsNull(!������), "", !������)
            Else
                str����Ա = ""
            End If
        End If
        If str����Ա <> "" Then
            '��λҽ��
            For IntLocate = 1 To cbo����ҽ��.ListCount
                If Mid(cbo����ҽ��.List(IntLocate), InStr(1, cbo����ҽ��.List(IntLocate), "-") + 1) = str����Ա Then
                    cbo����ҽ��.ListIndex = IntLocate
                    Exit For
                End If
            Next
        End If
        
        cbo����ҽ��.Enabled = ((mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ��) And mcondition.blnУ�鴦�� = True And cbo����ҽ��.ListIndex = 0)
        
        Lbl��ҩ��.Caption = IIf(mcondition.intListType = mListType.��ҩ, IIf(int�ɲ��� <> 3, "��ҩ��", "��ҩ��"), "��ҩ��")
        '������ҩ��
        If mcondition.intListType = mListType.��ҩ Then
            If IIf(IsNull(!��ҩ��), "", !��ҩ��) <> "" Then
                Me.Cbo��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            End If
        Else
        
            If IIf(IsNull(!��ҩ��), "", !��ҩ��) <> "" Then
                Me.Cbo��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            End If
        End If
        
        If mcondition.intListType = mListType.��ҩ Then
            If IIf(IsNull(!�˲���), "", !�˲���) <> "" Then
                Me.cbo�˲���.Text = IIf(IsNull(!�˲���), "", !�˲���)
            End If
        End If
        
        '��ҩ����
        picRecInfo_CM.Visible = False
        If mbln��ҩ���� Then
            picRecInfo_CM.Visible = True
            Call ��ҩ�����ر���(!ҩ��ID, !����, !NO, !��¼����, !�����־)
        End If
        Call Form_Resize
    End With
    
    SetCmdSendPrivs int�����
    
    If Me.CmdSend.Caption = "��ҩȷ��(&O)" And int�Ŷ�״̬ = 1 Then
        Me.CmdSend.Caption = "ȡ��ȷ��(&C)"
    ElseIf Me.CmdSend.Caption = "ȡ��ȷ��(&C)" And int�Ŷ�״̬ = 0 Then
        Me.CmdSend.Caption = "��ҩȷ��(&O)"
    End If
End Sub

Private Function �ж��Ƿ���ҩ����(ByVal lngNOҩ��id As Long, ByVal BillType As Integer, ByVal BillNo As String) As Boolean
    'ͨ��ҩƷid�ж��Ƿ�����ҩ
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim lngҩ��ID As Long
    Dim blnMoved As Boolean
    
    On Error GoTo errHandle
    
    lngҩ��ID = lngNOҩ��id
    If lngNOҩ��id = 0 Then lngҩ��ID = mcondition.lngҩ��ID
    
    strsql = "Select a.��� as ��� From �շ���ĿĿ¼ a ,ҩƷ�շ���¼ b Where b.ҩƷid=a.Id And b.����=[2] and b.No=[1] And (b.��¼״̬=1 Or Mod(b.��¼״̬,3)=0) and (b.�ⷿID+0=[3] OR b.�ⷿID IS NULL) " _
   
    '�������ת������ֱ�ӴӺ󱸱�����ȡ����
    blnMoved = Sys.IsMovedByNO("ҩƷ�շ���¼", BillNo, " ���� = ", BillType)
    If blnMoved Then
        gstrSQL = Replace(gstrSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(strsql, Me.Caption & "[�ж��Ƿ���ҩ����]", BillNo, BillType, lngҩ��ID)
    
    �ж��Ƿ���ҩ���� = IIf(rs!��� = 7, True, False)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ��ҩ�����ر���(ByVal lngNOҩ��id As Long, ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer)
    '��ҩ������ʾԭʼ��������ҩ�巨
    Dim rs As New ADODB.Recordset
    Dim lngҩ��ID As Long
    
    On Error GoTo errHandle
    lngҩ��ID = lngNOҩ��id
    If lngNOҩ��id = 0 Then lngҩ��ID = mcondition.lngҩ��ID

    gstrSQL = "Select a.���,b.���� From ҩƷ�շ���¼ a ,������ü�¼ b Where a.����id=b.Id " _
        & " And a.����=[2] And a.No=[1] " _
        & " And (a.��¼״̬=1 Or Mod(a.��¼״̬,3)=0) and (a.�ⷿID+0=[3] OR a.�ⷿID IS NULL) "
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ҩ�����ر���]", BillNo, BillStyle, lngҩ��ID)
    
    lblԭʼ����.Caption = lblԭʼ����.Tag & CStr(IIf(IsNull(rs!����), 1, rs!����))
    lbl��ҩ�巨.Caption = lbl��ҩ�巨.Tag & IIf(IsNull(rs!���), "", rs!���)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt�������_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetTip(txt�������, txt�������.Tag)
End Sub
