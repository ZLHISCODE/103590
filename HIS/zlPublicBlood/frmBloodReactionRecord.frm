VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmBloodReactionRecord 
   Caption         =   "��Ѫ��Ӧ��¼"
   ClientHeight    =   11895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14370
   Icon            =   "frmBloodReactionRecord.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11895
   ScaleWidth      =   14370
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer TimNotify 
      Interval        =   500
      Left            =   1560
      Top             =   0
   End
   Begin VB.PictureBox picTips 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   5400
      ScaleHeight     =   1455
      ScaleWidth      =   1815
      TabIndex        =   42
      Top             =   8520
      Width           =   1815
      Begin XtremeReportControl.ReportControl rptTips 
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   735
         _Version        =   589884
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox PicPane 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9465
      ScaleHeight     =   330
      ScaleWidth      =   4185
      TabIndex        =   22
      Top             =   11595
      Width           =   4185
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��Ѫ�ƴ��ύ"
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
         Index           =   2
         Left            =   2910
         TabIndex        =   28
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ҽ�����ύ"
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
         Index           =   1
         Left            =   1455
         TabIndex        =   27
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�����"
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
         Index           =   0
         Left            =   390
         TabIndex        =   26
         Top             =   60
         Width           =   585
      End
      Begin VB.Label lbl��ǩ״̬ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   2535
         TabIndex        =   25
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lbl��ǩ״̬ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lbl��ǩ״̬ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   23
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   4440
      ScaleHeight     =   6495
      ScaleWidth      =   8895
      TabIndex        =   17
      Top             =   540
      Width           =   8895
      Begin zlPublicBlood.usrCardEdit UCE 
         Height          =   8715
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   8805
         _extentx        =   15531
         _extenty        =   15372
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10050
      Left            =   120
      ScaleHeight     =   10050
      ScaleWidth      =   4575
      TabIndex        =   14
      Top             =   465
      Width           =   4575
      Begin VB.ComboBox cbo1 
         Height          =   300
         Left            =   570
         TabIndex        =   0
         Text            =   "���п���"
         Top             =   300
         Width           =   2700
      End
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   3480
         ScaleHeight     =   165
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   8160
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox picUCP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2460
         Left            =   60
         ScaleHeight     =   2460
         ScaleWidth      =   3855
         TabIndex        =   21
         Top             =   5595
         Width           =   3855
         Begin zlPublicBlood.usrCardPeople UCP 
            Height          =   2055
            Left            =   30
            TabIndex        =   12
            Top             =   0
            Width           =   3135
            _extentx        =   5530
            _extenty        =   3625
         End
      End
      Begin XtremeSuiteControls.TabControl tbcthis 
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   4530
         Width           =   3855
         _Version        =   589884
         _ExtentX        =   6800
         _ExtentY        =   1296
         _StockProps     =   64
      End
      Begin VB.Frame Fra1 
         Caption         =   "��������"
         Height          =   3795
         Left            =   75
         TabIndex        =   15
         Top             =   660
         Width           =   3855
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   75
            TabIndex        =   36
            Top             =   3050
            Width           =   3615
            Begin VB.OptionButton opt 
               Caption         =   "δ��"
               Height          =   255
               Index           =   3
               Left            =   2760
               TabIndex        =   44
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton opt 
               Caption         =   "����"
               Height          =   255
               Index           =   2
               Left            =   840
               TabIndex        =   39
               Top             =   0
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton opt 
               Caption         =   "��"
               Height          =   255
               Index           =   0
               Left            =   2160
               TabIndex        =   38
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton opt 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   37
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "��Ѫ��Ӧ"
               Height          =   180
               Left            =   20
               TabIndex        =   41
               Top             =   10
               Width           =   720
            End
         End
         Begin VB.TextBox TXTDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   930
            MaxLength       =   4
            TabIndex        =   6
            Text            =   "7"
            Top             =   1275
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Frame frmLine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   930
            TabIndex        =   32
            Top             =   1470
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cbotime 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   210
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "ˢ��"
            Height          =   300
            Left            =   2955
            TabIndex        =   11
            Top             =   3360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox ChkRection 
            Caption         =   "������Ѫ��Ӧ��д�������"
            Height          =   225
            Left            =   120
            TabIndex        =   7
            Top             =   1620
            Width           =   2475
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   300
            Left            =   945
            TabIndex        =   2
            Top             =   2310
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42593
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   300
            Left            =   945
            TabIndex        =   1
            Top             =   1965
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42593
         End
         Begin VB.CheckBox chk1 
            Caption         =   "δ�ύ"
            Height          =   225
            Index           =   2
            Left            =   2835
            TabIndex        =   10
            Top             =   2730
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chk1 
            Caption         =   "���ύ"
            Height          =   225
            Index           =   1
            Left            =   1965
            TabIndex        =   9
            Top             =   2730
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chk1 
            Caption         =   "������д"
            Height          =   225
            Index           =   0
            Left            =   915
            TabIndex        =   8
            Top             =   2730
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTP4 
            Height          =   300
            Left            =   930
            TabIndex        =   5
            Top             =   915
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42711
         End
         Begin MSComCtl2.DTPicker DTP3 
            Height          =   300
            Left            =   930
            TabIndex        =   4
            Top             =   585
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42711
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label lbl6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "��ʾ���       ��ת���Ĳ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   33
            Top             =   1260
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.Label lbl5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "��Ժ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "��ʼʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   630
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   990
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl4 
            Caption         =   "~"
            Height          =   135
            Left            =   720
            TabIndex        =   16
            Top             =   2460
            Width           =   135
         End
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   11535
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2461
            MinWidth        =   882
            Picture         =   "frmBloodReactionRecord.frx":07AA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11774
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7585
            MinWidth        =   7585
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
   Begin VSFlex8Ctl.VSFlexGrid VSFBRlist 
      Height          =   855
      Left            =   3000
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   2580
      _cx             =   4551
      _cy             =   1508
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpPeoPle 
      Bindings        =   "frmBloodReactionRecord.frx":1290
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBloodReactionRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSys As Long   '����ϵͳ��
Private mlngModule As Long           '����ģ���
Private mlng�׶� As Long             '0-���ﴦ��׶Ρ�1-סԺ����׶�  2-��Ѫ�ƴ���׶Σ���Ӧclspublicbloodģ���еĳ���
Private mstr���� As String
Private mstr��ʼʱ�� As String
Private mstr����ʱ�� As String
Private mstr��д�� As String
Private mlng�ύ״̬ As Long         '0-ȫ�����ݣ�1-δ�ύ���ݣ�2-���ύ����
Private mArr��������                 '��ſ��ҡ�ʱ�䡢�Ƿ�����д���ύ״̬
Private mstrPrivs As String          'Ȩ�޴�
Private mblnButtonChecked As Boolean '��׼��ť
Private mblnTextChecked As Boolean   '�ı���ǩ
Private mblnSizeChecked As Boolean   '��������С
Private mblnStatuChecked As Boolean  '״̬����ʾ
Private mfrmMain As Object           '������
Private mRsBR As ADODB.Recordset     '������Ϣ��¼��
Private mblnHaveBR As Boolean        '�ж��Ƿ��ѯ������
Private mlngtbcIndex As Long         '��¼tbcthisѡ�п�Ƭ��index
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mblnStart As Boolean         '�жϳ����Ƿ�ʼ
Private mArrPosition(0 To 2)         '��Ų����ڲ�ͬ״̬�µĶ�λ���ݣ�������Ժ����Ժ��ת����
Private mArrCheckData(0 To 9)        '��ſؼ��ĵ�ǰ����0~9��Ӧ��cbo1��dtp1~4��chk1(0~2)��cbotime��TXTDay
Private mblnIsSelect As Boolean      '�ж��Ƿ��в��˱�ѡ�� true��ʾ�У�false��ʾ��
Private mblnIsSubmit As Boolean      '�ж��Ƿ��ύ���߻��ˣ������ύ���˺����α����ȡ��ʹ��
Public mblnBloodReactionRecordIsOpen As Boolean '��ģ̬״̬�£��жϴ����Ƿ���
Private mstrFindKey As String             '��Ѫ������ʱ��ͨ����ѯҳ���ѯ���Ĳ���,��ʽ������ID-����ID-����(0-סԺ/2-����)
Private mblnADDPeoPle As Boolean     '��Ѫ������ʱ���ж����������˻��Ƕ�λ�����ˣ��������˾ͻ����¶�ȡ���˲��ҵ�sql����λ������ֱ��ѡ�в�ѯ�Ĳ��ˡ�
Private mintDeptIndex As Integer  '��������
Private mintNotify As Integer '�����Զ�ˢ�¼��(����)
Private mblnFirst As Boolean         '��һ����������ˢ��,���л����š���Ϣ�ı�ʱǿ��ˢ��

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��CommandBar
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ������
    
    Call CommandBarInit(cbsMain)
    '�˵�����:������������
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_File_MedRec, "��Ӧ��¼��ӡ")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_File_MedRecSetup, "��ӡ����"): objControl.IconId = conMenu_File_PrintSet
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_File_MedRecPreview, "Ԥ����Ӧ"): objControl.IconId = conMenu_File_Preview
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_File_MedRecPrint, "��ӡ��Ӧ"): objControl.IconId = conMenu_File_Print
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "��������", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    '�༭
    
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "����")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸�")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Save, "����", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "�ύ")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "����")
    
    '�鿴
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_FindNext, "������һ��(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    
    mblnButtonChecked = True
    mblnTextChecked = True
    mblnSizeChecked = True
    mblnStatuChecked = True
    
    '����
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_MedRecPreview, "Ԥ��"): objControl.ToolTipText = "Ԥ����Ѫ��Ӧ��": objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_File_MedRecPrint, "��ӡ"): objControl.ToolTipText = "��ӡ��Ѫ��Ӧ��": objControl.IconId = conMenu_File_Print
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True '
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, IIf(mlng�׶� = 2, "����", "�޸�"))
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, IIf(mlng�׶� = 2, "���", "�ύ")): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����")
        Set objControl = .Add(xtpControlButton, conMenu_View_Detail, "��Ѫִ��"): objControl.ToolTipText = "��Ѫִ���������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")

    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then objControl.Style = xtpButtonIconAndCaption
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyDelete, conMenu_Edit_Delete            'ɾ��
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify          '�޸�
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh                 'ˢ��
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '����
        .Add 0, vbKeyF3, conMenu_View_FindNext              '��������
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save            '����
        .Add FCONTROL, vbKeyC, conMenu_Edit_Transf_Cancle   'ȡ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add 0, vbKeyF12, conMenu_File_Parameter             '��������
        .Add FCONTROL, vbKeyX, conMenu_File_Exit            '�˳�
    End With
    
    Call gobjDatabase.ShowReportMenu(Me, 2200, p��Ѫ��Ӧ����, mstrPrivs)
    InitCommandBar = True
    
    Exit Function
ErrHand:
    
End Function

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    With rptTips
        Set objCol = .Columns.Add(0, "��������", 30, True)
        Set objCol = .Columns.Add(1, "��Ϣ����", 60, True)
        Set objCol = .Columns.Add(2, "����ID", 40, True): objCol.Visible = False
        Set objCol = .Columns.Add(3, "����ID", 40, True): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û����������..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
    End With
End Sub
Private Sub cbo1_Click()
    If mblnStart = False Then Exit Sub
    If mintDeptIndex = cbo1.ListIndex Then Exit Sub
    mintDeptIndex = cbo1.ListIndex
    RefreshBR
    mblnFirst = True
End Sub

Private Sub cbo1_KeyPress(KeyAscii As Integer)
    '���ܣ���Сдת��Ϊ��д����ѯƥ��Ĳ�����ʾ�������ת����һ���ؼ���
    Dim olddata As String
    Dim intIndex As Integer
    olddata = cbo1.Text
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = vbKeyReturn Then
        intIndex = findDepart(cbo1.Text)
        If intIndex = -1 Then '�Ҳ�����ԭ
            cbo1.ListIndex = mintDeptIndex
            gobjCommFun.PressKey vbKeyTab
        Else
            cbo1.ListIndex = intIndex
        End If
    End If
End Sub

Private Function findDepart(key As String) As Long
    '���ܣ����Ҳ����б��з��������Ĳ���
    Dim lngi As Long
    Dim blnfind As Boolean
    Dim ArrCbo
    
    For lngi = 0 To cbo1.ListCount - 1
        If cbo1.List(lngi) Like key & "*" Then
            blnfind = True
            findDepart = lngi
            Exit For
        End If
        ArrCbo = Split(cbo1.List(lngi), "-")
        If ArrCbo(0) Like key & "*" Then
            blnfind = True
            findDepart = lngi
            Exit For
        ElseIf UBound(ArrCbo) > 0 Then
            If ArrCbo(1) Like key & "*" Then
                blnfind = True
                findDepart = lngi
                Exit For
            End If
        End If
    Next
    If blnfind = False Then
        findDepart = -1
    End If
End Function

Private Sub Form_Activate()
    Set gobjFScrollBar = UCP.FScrollBar
    glngBooldPepWinProc = GetWindowLong(UCP.objPicBack.hWnd, GWL_WNDPROC)
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, glngBooldPepWinProc
End Sub

Private Sub rptTips_SelectionChanged()
    Dim lng�շ�id As Long, lng����ID As Long, lng����id As Long, int������Դ As Integer
    Dim strSQL As String, rs As Recordset, bytMode As Byte
    Dim strKey As String, lng��������4 As Long
    If Not Visible Then Exit Sub
    lng�շ�id = Val(rptTips.SelectedRows(0).Record(2).Value)
    lng����ID = Val(rptTips.SelectedRows(0).Record(3).Value)
    lng����id = Val(rptTips.SelectedRows(0).Record(4).Value)
    int������Դ = IIf(Val(rptTips.SelectedRows(0).Record(5).Value) = 2, 0, 1)
    lng��������4 = mArr��������(4)
    If mArr��������(4) = 0 Then mArr��������(4) = 2
    strKey = lng����ID & "-" & lng����id
    mblnADDPeoPle = Not UCP.findIdPeoPle(strKey, False)
    If mblnADDPeoPle Then
        mstrFindKey = lng����ID & "-" & lng����id & "-" & int������Դ
        Call ExecuteCommand("ˢ������")
        If UCP.findIdPeoPle(strKey, False) Then
            If Not UCE.BloodLocation(lng�շ�id) Then Call MsgBox("δ�ҵ���ӦѪҺ��", vbInformation, Me.Caption)
        Else
            Call MsgBox("δ�ҵ���Ӧ���ˡ�", vbInformation, Me.Caption)
        End If
    Else
        If Not UCE.BloodLocation(lng�շ�id) Then Call MsgBox("δ�ҵ���ӦѪҺ��", vbInformation, Me.Caption)
    End If
    mblnADDPeoPle = False
    mArr��������(4) = lng��������4
End Sub
Private Sub TimNotify_Timer()
    Static strPreTime1 As String
    Dim curTime As Date
    curTime = Now
    'ˢ������
    If mintNotify > 0 Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Or mblnFirst Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call ExecuteCommand("ˢ����ʾ")
            mblnFirst = False
        End If
     Else
        If mblnFirst = True Then
            Call ExecuteCommand("ˢ����ʾ")
            mblnFirst = False
        End If
    End If
End Sub
Private Sub cbo1_LostFocus()
    If cbo1.Text <> cbo1.List(cbo1.ListIndex) Then cbo1.Text = cbo1.List(cbo1.ListIndex)
End Sub

Private Sub cboTime_Click()
    Dim blnEnable As Boolean, strCurDate As String
    
    blnEnable = Val(cbotime.ItemData(cbotime.ListIndex)) = -1
    strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    DTP3.Value = Format(CDate(strCurDate) - Val(cbotime.ItemData(cbotime.ListIndex)), "YYYY-MM-DD")
    DTP4.Value = Format(strCurDate, "YYYY-MM-DD") & " 23:59:59"
    DTP3.Enabled = blnEnable
    DTP4.Enabled = blnEnable
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey  As String
    Dim arrKey
    Select Case Control.id
        Case conMenu_File_Preview 'Ԥ�������б�
            Call zlRptPrint(2, VSFBRlist, "�����б�")
        Case conMenu_File_Print '��ӡ�����б�
            Call zlRptPrint(1, VSFBRlist, "�����б�")
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
        Case conMenu_File_MedRecSetup '��Ѫ��Ӧ�����ӡ����
            Call UCE.showPrintSet
        Case conMenu_File_MedRecPreview '��Ѫ��Ӧ��ӡԤ��
            If UCE.lngFYCount = 1 Then
                Call UCE.ShowPrint(2)
            Else
                Call UCE.ShowPrintList(2)
            End If
        Case conMenu_File_MedRecPrint '��Ѫ��Ӧ��ӡ
            If UCE.lngFYCount = 1 Then
                Call UCE.ShowPrint(1)
            Else
                Call UCE.ShowPrintList(1)
            End If
        Case conMenu_Edit_NewItem: '����
            If mlng�׶� = 2 Then
                strKey = frmBloodPeoPleSerch.SerchPeople(mfrmMain, mlngModule)
                mstrFindKey = strKey
                If strKey <> "" Then
                    arrKey = Split(strKey, "-")
                    strKey = arrKey(0) & "-" & arrKey(1)
                    mblnADDPeoPle = Not UCP.findIdPeoPle(strKey, False)
                    If mblnADDPeoPle = True Then '������еĲ����б�����û��Ҫ�ҵ��Ĳ����������ò���
                        Call ExecuteCommand("ˢ������")
                        If UCP.findIdPeoPle(strKey, False) = True Then
'                            UCE.DataChanged = True
                            UCE.AddPage
                            mblnIsSubmit = False
                        Else
                            mstrFindKey = ""
                        End If
                        mblnADDPeoPle = False
                    Else
                        '�����ǰ�����б��в�ѯ�����ݣ��������Ĳ��ˣ�������ҳ��
                        UCE.AddPage
                        mblnIsSubmit = False
                    End If
                    UCP.locked = IIf(mstrFindKey = "", False, True) '������״̬������ucp�ؼ�
                End If
            Else
                UCE.AddPage
                mblnIsSubmit = False
                UCP.locked = True
            End If
        Case conMenu_Edit_Modify: '�޸�
            UCE.ShowModify
            mblnIsSubmit = False
        Case conMenu_Edit_Delete: 'ɾ��
            If IsPrivs(mstrPrivs, "ɾ������") = False Then
                If UCE.Doctor <> "" And UCE.Doctor <> UserInfo.���� Then
                    MsgBox "��û��Ȩ��ɾ�����˼�¼�����ݣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            UCE.ShowDelete
            mblnIsSubmit = True
        Case conMenu_Edit_Save: '����
            If UCE.ShowSave = False Then Exit Sub
            If mlng�׶� = 2 Then mblnFirst = True
            UCP.locked = False '�����ȡ������
        Case conMenu_Edit_Transf_Cancle: 'ȡ��
            UCE.ShowCancel
            UCP.locked = False 'ȡ����ȡ������
            Call ExecuteCommand("ˢ������")
        Case conMenu_Edit_Audit: '�ύ
            UCE.SubmitData
            mblnIsSubmit = True
        Case conMenu_Edit_Untread: '����
            UCE.ShowUntread
            mblnIsSubmit = True
        Case conMenu_View_Detail 'ִ������鿴
            Call frmBloodExecEdit.ViewExecution(Me, UCE.BloodID)
        Case conMenu_View_Refresh: 'ˢ��
            Call ExecuteCommand("ˢ������")
        Case conMenu_View_ToolBar_Button
            mblnButtonChecked = Not mblnButtonChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Text
            mblnTextChecked = Not mblnTextChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Size
            mblnSizeChecked = Not mblnSizeChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_StatusBar
            mblnStatuChecked = Not mblnStatuChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_File_Parameter
            Call ExecuteCommand("���ز�������")
        Case conMenu_Help_Help              '��������
            Call gobjComlib.ShowHelp(App.ProductName, Me.hWnd, Me.name, Int((2200) / 100))
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call gobjComlib.zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum         'Web�ϵ���̳
            Call gobjComlib.zlWebForum(Me.hWnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call gobjComlib.zlMailTo(Me.hWnd)
        Case conMenu_Help_About '����
            Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Exit '�˳�
            Unload Me
        Case conMenu_View_Find, conMenu_View_FindNext '���ң���������
            Call UCP.FindPatiByVbKey(Control.id = conMenu_View_FindNext)
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = Me.stbThis.Height
        PicPane.Visible = True
        PicPane.Top = Me.stbThis.Top + 60
        PicPane.Left = stbThis.Panels(6).Left + 120
    Else
        PicPane.Visible = False
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_File_MedRec, conMenu_File_MedRecSetup, conMenu_File_MedRecPreview, conMenu_File_MedRecPrint
            Control.Visible = IsPrivs(mstrPrivs, "���ݴ�ӡ")
            Control.Enabled = Control.Visible
        Case conMenu_Edit_Modify: '�޸�
            '�޼�¼��Ӧ��Ȩ�����޸İ�ť���ɼ���
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            
            Control.Caption = IIf(UCE.��Ѫ������ = False And mlng�׶� = 2, "����", "�޸�")
            'ҽ���׶����ύ״̬ ���� ��Ѫ�����ύ״̬ ���� ҽ���׶���Ѫ������������ ���� ����״̬ ���� �޸�״̬ ���� �޲������� ���� δѡ�в��˵�������޸İ�ť��ʹ�ܣ��������ʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 0) Or (mlng�׶� = 2 And UCE.lng״̬ = 2) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or UCE.strST = ���� Or UCE.strST = �޸� Or mblnHaveBR = False Or mblnIsSelect = False)
            '����ʱ�����������Ѫ���ô���
            If mlng�׶� = 2 And Control.Enabled And Control.Caption = "����" Then
                Control.Enabled = UCE.������Ѫ��Ӧ
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Delete: 'ɾ��
            '��ɾ����¼��Ȩ�޻�������Ѫ�ƽ׶Σ���ɾ����ť���ɼ���
            Control.Visible = IsPrivs(mstrPrivs, "ɾ����¼")   'and not (mlng�׶�=2 and not IsPrivs(mstrPrivs, "��Ѫ������"))������������Ѫ���������Ȩ�޵����������ɾ��
            'ҽ���׶����ύ״̬ ���� ��Ѫ�ƽ׶����ύ״̬ ���� ҽ���׶���Ѫ������������ ���� ����״̬ ���� �޸�״̬ ���� �޲������� ���� δѡ�в��˵������ɾ����ť��ʹ�ܣ��������ʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 0) Or (mlng�׶� = 2 And UCE.lng״̬ <> 0) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or UCE.strST = ���� Or UCE.strST = �޸� Or mblnHaveBR = False Or mblnIsSelect = False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_NewItem: '����
            '��Ѫ��û������Ȩ��ʱ��������ť���ɼ����������������ť�ɼ���
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ") And Not (mlng�׶� = 2 And Not IsPrivs(mstrPrivs, "��Ѫ������")) '����������Ѫ���������Ȩ�޵��������������
            '���blnAddPage=true����û�в��˻���û��ѡ�в���ʱ��������ʹ��
            Control.Enabled = Not (UCE.blnAddPage = True Or mblnHaveBR = False Or mblnIsSelect = False)
            'blnAddPage=true ���ߵ�ǰ�����������޸�״̬���ߵ�ǰ�޲��˻���δѡ�в���ʱ��������ť��ʹ�ܣ�����������Խ���������
'            Control.Enabled = Not (UCE.strST = ���� Or UCE.strST = �޸� Or mblnHaveBR = False Or mblnIsSelect = False) 'UCE.blnAddPage = True Or
            '���⣺����Ѫ�ƽ׶Σ��û�����Ѫ������Ȩ��ʱ����ʹ�����б��޲��˻���δѡ�в��ˣ�Ҳ���ǿ��Խ�������������
            If mlng�׶� = 2 And (mblnHaveBR = False Or mblnIsSelect = False) Then
                Control.Enabled = True
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Save: '����
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            'δ�ύ�����ݱ仯ʱ������ʹ��
            Control.Enabled = UCE.DataChanged And mblnIsSubmit = False
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Transf_Cancle: 'ȡ��
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            'δ�ύ�����ݱ仯ʱ��ȡ��ʹ��
            Control.Enabled = UCE.DataChanged And mblnIsSubmit = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Audit: '�ύ
            Control.Visible = IsPrivs(mstrPrivs, "�ύ����")
            
            Control.Caption = IIf(mlng�׶� = 2, "���", "�ύ")
            'ҽ���׶����ύ���� ���� ��Ѫ�ƽ׶����ύ���� ���� ҽ���׶���Ѫ������״̬ ���� ���������޸�״̬ ���� �޲��˻���δѡ�в���ʱ�ύ��ʹ�ܣ�����״̬�ύʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 0) Or (mlng�׶� = 2 And UCE.lng״̬ = 2) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or UCE.strST = ���� Or UCE.strST = �޸� Or mblnHaveBR = False Or mblnIsSelect = False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Untread: '����
            Control.Visible = IsPrivs(mstrPrivs, "�ύ����")
            'ҽ���׶η�ҽ�����ύ״̬ ���� ��Ѫ�ƽ׶�δ�ύ״̬ ���� ҽ���׶���Ѫ���������� ���� ��Ѫ�ƽ׶���Ѫ����������δ�ύ״̬ ���� ���������޸�״̬ ���� �޲��˻���δѡ�в��� ʱ���˲�ʹ�ܣ�����״̬����ʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 1) Or (mlng�׶� = 2 And UCE.lng״̬ <> 2) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or (mlng�׶� = 2 And UCE.lng״̬ = 0 And UCE.��Ѫ������) Or UCE.strST = ���� Or UCE.strST = �޸� Or mblnHaveBR = False Or mblnIsSelect = False)
        Case conMenu_View_Detail
            Control.Enabled = UCE.BloodID > 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_Parameter     'ҽ����ʱ�����õ���������
            Control.Visible = mlng�׶� = 2
            Control.Enabled = mlng�׶� = 2
        Case conMenu_View_ToolBar_Button
            Control.Checked = mblnButtonChecked
        Case conMenu_View_ToolBar_Text
            Control.Checked = mblnTextChecked
            
        Case conMenu_View_ToolBar_Size
            Control.Checked = mblnSizeChecked
            
        Case conMenu_View_StatusBar
            stbThis.Visible = mblnStatuChecked
            Control.Checked = mblnStatuChecked
    End Select
End Sub

Private Sub chk1_Click(Index As Integer)
    If chk1(1).Value = Unchecked And chk1(2).Value = Unchecked Then
        If Index = 1 Then
            chk1(2).Value = Checked
        ElseIf Index = 2 Then
            chk1(1).Value = Checked
        End If
    End If
End Sub


Private Sub RefreshBR()
    Dim strCurDate As String
    If cbo1.ListIndex = -1 Then Exit Sub
    
    If mblnStart = True Then
        If Not Me.ActiveControl Is Nothing Then
            Select Case Me.ActiveControl.name
                Case "DTP1", "DTP2", "DTP3", "DTP4"
                    Call gobjControl.ControlSetFocus(picUCP)
            End Select
        End If
    End If
    
    mArr��������(0) = Val(cbo1.ItemData(cbo1.ListIndex)) 'ѡ�еĿ���id��-1��ʾ���п���
    mArr��������(2) = IIf(chk1(0).Value = Checked, UserInfo.����, "") '�Ƿ�����д,�������ȡ��д�ˣ���Ϊ��,���˸ģ���ǰchecked��true�����ǵ��¹�ѡ������д��Ч����Ҫԭ��
    If opt(0).Value Then mArr��������(4) = 0
    If opt(1).Value Then mArr��������(4) = 1
    If opt(2).Value Then mArr��������(4) = 2
    If opt(3).Value Then mArr��������(4) = 3
    
    If mlng�׶� = 2 Then
        mstr��ʼʱ�� = DTP1.Value
        mstr����ʱ�� = DTP2.Value
        mArr��������(1) = mstr��ʼʱ�� & "'" & mstr����ʱ�� '����ʱ��
    Else
        If (mlngtbcIndex = 2 And mlng�׶� = 1) Or (mlng�׶� = 0 And mlngtbcIndex = 1) Then
            strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
            If cbotime.ItemData(cbotime.ListIndex) = -1 Then
                mstr��ʼʱ�� = DTP3.Value
                mstr����ʱ�� = DTP4.Value
            Else
                mstr��ʼʱ�� = Format(CDate(strCurDate) - Val(cbotime.ItemData(cbotime.ListIndex)), "YYYY-MM-DD")
                mstr����ʱ�� = Format(strCurDate, "YYYY-MM-DD") & " 23:59:59"
            End If
        End If
        mArr��������(1) = mstr��ʼʱ�� & "'" & mstr����ʱ��
    End If
    Call ExecuteCommand("ˢ������")
End Sub

Public Sub BloodReactionRecord(frmMain As Variant, lng�׶� As Long, ByVal lngSys As Long, ByVal lngMoudle As Long, Optional strPrivs As String, Optional lngisModul As Long = 0)
    '���ܣ���Ѫ��Ӧ��¼�ĵ��ú���
    '������frmMain-������
    '      lng�׶�-0:����ҽ������׶�1:סԺҽ������׶Ρ�2����Ѫ�ƴ���׶�
    '      lngMoudle-ģ���
    '      strPrivs-Ȩ�޴�
    '      lngisModul-0-��ģ̬��1-ģ̬
    Dim strSQL As String
    Dim rs���� As ADODB.Recordset
    Dim lngi As Long
    Dim rs�ϼ����� As ADODB.Recordset
    Dim objPane As Object
    Dim lngIndex As Long
    Dim strCurDate As String
    
    mblnFirst = True
    ReDim mArr��������(0 To 4)
    lngIndex = 0
    If mblnBloodReactionRecordIsOpen = True Then GoTo TOSHOW
    '��ʼ��ȫ�ֱ���
    Set mfrmMain = frmMain
    mstr���� = cbo1.Text
    mlngtbcIndex = 0
    mlngSys = lngSys
    mlngModule = lngMoudle
    mstrPrivs = strPrivs
    mlng�׶� = lng�׶�
    mblnADDPeoPle = False
    mstrFindKey = ""
    mlng�ύ״̬ = 0 '0-ȫ�����ݣ�1-δ�ύ���ݣ�2-���ύ����
    mblnStart = False
    strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    DTP1.Value = Format(CDate(strCurDate) - 29, "YYYY-MM-DD 00:00:00")
    DTP2.Value = Format(strCurDate, "YYYY-MM-DD 23:59:59")
    TimNotify.Enabled = mlng�׶� = 2
    
    If mlng�׶� = 2 Then
        mstr��ʼʱ�� = DTP1.Value
        mstr����ʱ�� = DTP2.Value
        mintNotify = gobjDatabase.GetPara("��Ϣ���Ѽ��", 2200, 1938, 0)
    Else
        mstr��ʼʱ�� = Format(strCurDate, "YYYY-MM-DD")
        mstr����ʱ�� = Format(strCurDate, "YYYY-MM-DD") & " 23:59:59"
    End If
    
    '��ʼ��commandbar
    InitCommandBar
    '��ʼ��cboTime�ؼ�����ؿؼ�
    initComboTime
    
    '��ʼ��dockingpane
    Me.dkpPeoPle.SetCommandBars Me.cbsMain
    Me.dkpPeoPle.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpPeoPle.Options.ThemedFloatingFrames = True
    Me.dkpPeoPle.Options.AlphaDockingContext = True
    Me.dkpPeoPle.Options.HideClient = True
    
    Set objPane = dkpPeoPle.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "�����б�": objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set objPane = dkpPeoPle.CreatePane(2, 100, 30, DockBottomOf, objPane): objPane.Title = "��Ϣ����": objPane.Options = PaneNoFloatable Or PaneNoHideable Or PaneNoCloseable
    If mlng�׶� <> 2 Then objPane.Options = PaneActionClosed: dkpPeoPle(2).Close
    Set objPane = dkpPeoPle.CreatePane(3, 700, 100, DockRightOf, Nothing): objPane.Title = "��¼": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    '��ʼ����Ϣ��ʾ�ؼ�
    InitReportColumn
    '��ʼ�����ſ�����Ϣ
    If mlng�׶� = 2 Then
        Set rs���� = GetDeptList("Ѫ��", 3, IsPrivs(mstrPrivs, "���п���"))
    Else
        Set rs���� = GetDeptList("�ٴ�", mlng�׶� + 1, IsPrivs(mstrPrivs, "���п���"))
    End If
    
    If rs����.RecordCount <= 0 Then
        MsgBox "�㲻����" & IIf(mlng�׶� = 2, "Ѫ��", "�ٴ�") & "���ţ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If IsPrivs(mstrPrivs, "���п���") Then
        cbo1.AddItem "���п���"
        cbo1.ItemData(cbo1.NewIndex) = -1 '���п��ҵ�idĬ��Ϊ-1
    End If
    
    For lngi = 0 To rs����.RecordCount - 1
        cbo1.AddItem rs����.Fields("����") & "-" & rs����.Fields("����").Value
        cbo1.ItemData(cbo1.NewIndex) = rs����.Fields("id").Value
        If rs����.Fields("id").Value = UserInfo.����ID Then lngIndex = IIf(IsPrivs(mstrPrivs, "���п���") = True, lngi + 1, lngi)
        rs����.MoveNext
    Next
    cbo1.ListIndex = lngIndex
    mintDeptIndex = lngIndex
    
    mArr��������(0) = -1
    mArr��������(1) = mstr��ʼʱ�� & "'" & mstr����ʱ��
    mArr��������(2) = IIf(chk1(0).Value = Checked, UserInfo.����, "")
    mArr��������(3) = mlng�ύ״̬
    If opt(0).Value Then mArr��������(4) = 0
    If opt(1).Value Then mArr��������(4) = 1
    If opt(2).Value Then mArr��������(4) = 2
    If opt(3).Value Then mArr��������(4) = 3
    
    '��ʼ��tabControl
    Call initTabControl(mlng�׶�)
    
    '��ʼ��usrCardEdit�ؼ�
    UCE.InitEdit
    '��ʼ��usrCardPeople
    UCP.UserInit Me, "��ɫ|ID|1||||255;סԺ���|��ҳID;����;����;������;�Ա������;��Ժ����;��д��", , p��Ѫ��Ӧ����
    
    '��ʼ�����Ͳ�ѯ������Ϣ
    Call ExecuteCommand("��ʼ���")
    Call RefreshBR
        
    If Not mRsBR Is Nothing Then
        If mRsBR.RecordCount > 0 Then
            mblnHaveBR = True
        End If
    Else
        mblnHaveBR = False
    End If
    mblnStart = True
TOSHOW:
    If mblnBloodReactionRecordIsOpen = True Then
        mArr��������(0) = -1
        mArr��������(1) = mstr��ʼʱ�� & "'" & mstr����ʱ��
        mArr��������(2) = UserInfo.����
        mArr��������(3) = mlng�ύ״̬
        If opt(0).Value Then mArr��������(4) = 0
        If opt(1).Value Then mArr��������(4) = 1
        If opt(2).Value Then mArr��������(4) = 2
        If opt(3).Value Then mArr��������(4) = 3
    End If
    mblnBloodReactionRecordIsOpen = True
    If IsObject(mfrmMain) Then
        If frmMain Is Nothing Then
            Me.Show lngisModul
        Else
            Me.Show lngisModul, mfrmMain
        End If
    Else
        gobjComlib.os.ShowChildWindow Me.hWnd, Val(mfrmMain)
    End If
    
End Sub

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��DockPannel
    '������
    '���أ�
    '******************************************************************************************************************
    
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long
    Dim lngj As Long
    Dim rsSAD As New ADODB.Recordset
    Dim Arr����
    Dim lng����ID As Long, lng��ҳid As Long
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
            Case "��ʼ���"
                Set mclsVsf = New clsVsf
                With mclsVsf
                    Call .Initialize(Me.Controls, VSFBRlist, True, True)
                    Call .ClearColumn
                    Call .AppendColumn("����id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("��ҳid", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("����", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("�Ա������", 1000, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("����", 700, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("������", 1400, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("סԺ���", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��Ժ����", 1100, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��������", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��д��", 1000, flexAlignLeftCenter, flexDTString, "", , True)

                    .AppendRows = False
                    .SysHidden(.ColIndex("����id")) = True
                    .SysHidden(.ColIndex("��ҳid")) = True
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    
                End With
                
            Case "�������˲�ѯ":
                '���ܣ����ݲ�ѯ����ѯ���ݿ��з������������ݲ�����ָ���ļ����У�
                Dim strSQL As String
                Dim strSql1 As String
                Dim strSqlRection As String
                Dim lng�ύ״̬ As String
                Dim str�����ύ As String
                Dim lngColor As Long
                Dim rsBR As ADODB.Recordset
                Dim lng����ID As Long
                
                lng����ID = Val(cbo1.ItemData(cbo1.ListIndex))
                
                If chk1(1).Value = Checked And chk1(2).Value = Unchecked Then  '���ύ
                    mlng�ύ״̬ = 2
                ElseIf chk1(1).Value = Unchecked And chk1(2).Value = Checked Then  'δ�ύ
                    mlng�ύ״̬ = 1
                Else 'ȫ������
                    mlng�ύ״̬ = 0
                End If

                Select Case mlng�׶�
                    Case 0: '���ﲡ��   ȥ�����"and f.ִ����=[4]" ,����˶���Ѫ��Ӧ��¼�ļ�¼�˵��жϣ�
                        strSQL = " Select f.����ID || '-' || f.id As id, f.����id, f.Id As ��ҳid, f.����� As ������, '��' As סԺ���, f.����, f.�Ա� || '/' || f.���� As �Ա������, f.ִ�в���id As ����id, " & vbNewLine & _
                                 "        b.���� As ��������, f.ִ��ʱ�� As ��Ժ����, '' As ����, f.����, '' As ����, 255 As ��ɫ, f.ִ���� As ��д�� " & vbNewLine & _
                                 " From ���˹Һż�¼ f, ���ű� b " & vbNewLine & _
                                 " Where f.ִ�в���id = b.Id And f.ִ��״̬ = [5] " & vbNewLine & _
                                 " " & IIf(lng����ID <> -1, " and b.id=[3] ", "") & IIf(mlngtbcIndex = 1, " and f.ִ��ʱ�� Between [1] And [2] ", "")
                        
                        If ChkRection.Value = 1 Then
                            strSqlRection = " And Exists (Select 1 " & vbNewLine & _
                                 "        From ѪҺ��Ѫ��¼ e, ѪҺ�շ���¼ g, ����ҽ����¼ h,��Ѫ��Ӧ��¼ J " & vbNewLine & _
                                 "        Where h.Id = e.����id And e.Id = g.�䷢id  and g.id=j.�շ�id " & IIf(chk1(0).Value = Checked, " and j.��¼��=[4] ", "") & " " & IIf(mlng�ύ״̬ = 1, " and J.״̬=0 ", IIf(mlng�ύ״̬ = 2, " and J.״̬<>0 ", "")) & " And " & vbNewLine & _
                                 "              Mod(g.��¼״̬, 3) = 1 And g.����� Is Not Null And H.�������='K' And h.�Һŵ� = f.No And  h.����id = f.����id)"
                        Else
                            strSqlRection = " And Exists (Select 1 " & vbNewLine & _
                                 "        From ѪҺ��Ѫ��¼ e, ѪҺ�շ���¼ g, ����ҽ����¼ h " & vbNewLine & _
                                 "        Where h.Id = e.����id And e.Id = g.�䷢id And h.����id = f.����id And " & vbNewLine & _
                                 "              h.�Һŵ� = f.No And Mod(g.��¼״̬, 3) = 1 And g.����� Is Not Null)"
                        End If
                        
                        strSQL = strSQL & strSqlRection
                        
                        If mlngtbcIndex = 0 Then '���ھ���
                            Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), lng����ID, UserInfo.����, 2)
                        ElseIf mlngtbcIndex = 1 Then '��ɾ���
                            Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), lng����ID, UserInfo.����, 1)
                        End If
                        
                    Case 1: 'סԺ����
                        If mlngtbcIndex = 0 Then '1��ʾ��Ժ
                            strSQL = " Select a.����ID || '-' || a.��ҳid As id,a.����id, a.��ҳid, a.סԺ�� As ������, 'ס' As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա������," & vbNewLine & _
                                     " f.����id As ����id, b.���� As ��������, a.��Ժ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ, a.�Ǽ��� As ��д�� " & vbNewLine & _
                                     " From ������ҳ a,������Ϣ C, ���ű� b, ��Ժ���� f " & vbNewLine & _
                                     " Where b.id = f.����id And a.����id = c.����id And a.��ҳID = c.��ҳID And c.����id = f.����id " & IIf(lng����ID <> -1, " and f.����id=[3] ", "") & ""
                        ElseIf mlngtbcIndex = 2 Then '2��ʾ��Ժ
                            strSQL = " Select a.����ID || '-' || a.��ҳid As id,a.����id, a.��ҳid, a.סԺ�� As ������, 'ס' As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա������," & vbNewLine & _
                                     "        a.��Ժ����id As ����id, b.���� As ��������, a.��Ժ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ, a.�Ǽ��� As ��д�� " & vbNewLine & _
                                     " From ������ҳ a, ������Ϣ c, ���ű� b " & vbNewLine & _
                                     " Where a.����id = c.����id And a.��ҳid = c.��ҳid And a.��Ժ����id = b.Id " & IIf(lng����ID <> -1, " and a.��Ժ����id=[3] ", "") & " And a.��Ժ���� Between [1] And [2] "
                        Else '1��ʾת��
                            strSQL = " Select a.����ID || '-' || a.��ҳid As id,a.����id, a.��ҳid, a.סԺ��  As ������, 'ס' As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա������, " & vbNewLine & _
                                     "       f.����id As ����id, b.���� As ��������, a.��Ժ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ, a.�Ǽ��� As ��д�� " & vbNewLine & _
                                     " From ������ҳ a, ������Ϣ c, ���ű� b, ���˱䶯��¼ f  " & vbNewLine & _
                                     " Where a.����id = c.����id And a.��ҳid = c.��ҳid And f.����id = a.����id And f.��ҳid = a.��ҳid And b.Id = f.����id And f.��ʼԭ�� = 3 And " & vbNewLine & _
                                     "       Nvl(f.���Ӵ�λ, 0) = 0 " & IIf(lng����ID <> -1, " and f.����id=[3] ", "") & " And f.��ʼʱ�� Between Sysdate - [5] And Sysdate  "
                        End If
                        
                        If ChkRection.Value = 1 Then
                            strSqlRection = " And Exists " & vbNewLine & _
                                     " (Select 1 From ѪҺ��Ѫ��¼ e, ѪҺ�շ���¼ g,��Ѫ��Ӧ��¼ h " & vbNewLine & _
                                     "        Where e.Id = g.�䷢id And e.����id = a.����id And e.��ҳid = a.��ҳid and g.id=h.�շ�id  " & IIf(chk1(0).Value = Checked, " and h.��¼��=[4] ", "") & " " & IIf(mlng�ύ״̬ = 1, " and h.״̬=0 ", IIf(mlng�ύ״̬ = 2, " and h.״̬<>0 ", "")) & " And  Mod(g.��¼״̬, 3) = 1 And g.����� Is Not Null)"
                        Else
                            strSqlRection = " And Exists " & vbNewLine & _
                                     " (Select 1 From ѪҺ��Ѫ��¼ e, ѪҺ�շ���¼ g " & vbNewLine & _
                                     "        Where e.Id = g.�䷢id And e.����id = a.����id And e.��ҳid = a.��ҳid And  Mod(g.��¼״̬, 3) = 1 And g.����� Is Not Null)"
                        End If
                        
                        strSQL = strSQL & strSqlRection
                        
                        Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), lng����ID, UserInfo.����, Val(TXTDay.Text))
                    Case 2: '��Ѫ��
                        'ȷ���ύ״̬�Ĺ�������
                        If mlng�ύ״̬ = 0 Then 'ȫ������
                            strSql1 = " and (e.״̬<>0 or e.�Ƿ���Ѫ������ =1) "
                        ElseIf mlng�ύ״̬ = 1 Then 'δ�ύ����
                            strSql1 = " and (e.״̬<>2 and e.�Ƿ���Ѫ������=1 OR e.״̬=1)"
                        Else '���ύ����
                            strSql1 = " and e.״̬=2 "
                        End If
                        If mArr��������(4) = 0 Then
                            strSql1 = strSql1 & " and e.������Ѫ��Ӧ = 2 "
                        ElseIf mArr��������(4) = 1 Then
                            strSql1 = strSql1 & " and e.������Ѫ��Ӧ = 1 "
                        ElseIf mArr��������(4) = 2 Then
                        ElseIf mArr��������(4) = 3 Then
                            strSql1 = strSql1 & " and e.������Ѫ��Ӧ = 0 "
                        End If
                        'ȥ����䡰and a.�Ǽ���=[5]�� ��  and f.ִ����=[5]�������˶���Ѫ��ȷ���˵��ж�
                        strSQL = " Select a.����ID || '-' || a.��ҳid As id,a.����id, a.��ҳid, a.סԺ�� As ������, 'ס' As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա������, " & vbNewLine & _
                                 "        a.��Ժ����id As ����id, b.���� As ��������, a.��Ժ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ, a.�Ǽ��� As ��д�� " & vbNewLine & _
                                 " From ������ҳ a, ���ű� b, " & vbNewLine & _
                                 "      (Select Distinct c.����id, c.��ҳid " & vbNewLine & _
                                 "       From ѪҺ��Ѫ��¼ c, ѪҺ�շ���¼ d, ��Ѫ��Ӧ��¼ e " & vbNewLine & _
                                 "       Where c.Id = d.�䷢id And d.Id = e.�շ�id " & IIf(chk1(0).Value = Checked, " And (e.ȷ����=[5] or e.�Ƿ���Ѫ������ =1 And e.��¼��=[5]) ", "") & " And e.��Ӧʱ�� Between [1] And [2] " & IIf(lng����ID = -1, "", " And c.ִ�в���id = [3] ") & vbNewLine & _
                                 " And c.��¼���� = 1 " & strSql1 & "  And Mod(d.��¼״̬, 3) = 1 And d.����� Is Not Null) K " & vbNewLine & _
                                 "Where K.����id = a.����id And K.��ҳid = a.��ҳid And a.��Ժ����id = b.Id(+) "

                        strSQL = strSQL & " Union ALL" & vbNewLine & _
                                " Select f.����ID || '-' || f.id As id,f.����id, f.Id As ��ҳid, f.����� As ������, '��' As סԺ���, f.����, f.�Ա� || '/' || f.���� As �Ա������, f.ִ�в���id As ����id, " & vbNewLine & _
                                "        b.���� As ��������, f.ִ��ʱ�� As ��Ժ����, '' As ����, f.����, '' As ����, 255 As ��ɫ, f.ִ���� As ��д�� " & vbNewLine & _
                                " From ���˹Һż�¼ f, ���ű� b, " & vbNewLine & _
                                "      (Select Distinct g.����id, g.�Һŵ� " & vbNewLine & _
                                "       From ѪҺ��Ѫ��¼ c, ѪҺ�շ���¼ d, ��Ѫ��Ӧ��¼ e, ����ҽ����¼ g " & vbNewLine & _
                                "       Where c.Id = d.�䷢id And d.Id = e.�շ�id " & IIf(chk1(0).Value = Checked, " And (e.ȷ����=[5] or e.�Ƿ���Ѫ������ =1 And e.��¼��=[5]) ", "") & " And g.Id = c.����id And c.��¼���� = 1 " & strSql1 & " And Mod(d.��¼״̬, 3) = 1 And " & vbNewLine & _
                                "             d.����� Is Not Null And e.��Ӧʱ�� Between [1] And [2] " & IIf(lng����ID = -1, "", " And c.ִ�в���id = [3] ") & " And g.������� = 'K') h " & vbNewLine & _
                                " Where h.����id = f.����id And h.�Һŵ� = f.No And f.ִ�в���id = b.Id "
                        
                        '���ϲ�ѯ���Ĳ���
                        If mblnADDPeoPle And mstrFindKey <> "" Then
                            Arr���� = Split(mstrFindKey, "-")  '����ID-����ID-�����סԺ
                            lng����ID = Val(Arr����(0))
                            lng��ҳid = Val(Arr����(1))
                            If Val(Arr����(2)) = 0 Then 'סԺ
                                strSQL = strSQL & " Union ALL" & vbNewLine & _
                                    " Select a.����id || '-' || a.��ҳid As Id, a.����id, a.��ҳid, a.סԺ��  As ������, 'ס' As סԺ���, a.����," & vbNewLine & _
                                    "       a.�Ա� || '/' || a.���� As �Ա������, a.��Ժ����id As ����id, d.���� As ��������, a.��Ժ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ," & vbNewLine & _
                                    "       a.�Ǽ��� As ��д��" & vbNewLine & _
                                    " From ���ű� d, ������ҳ a" & vbNewLine & _
                                    " Where a.��Ժ����id = d.Id(+) And a.����id = [6] And a.��ҳid = [7] And Exists" & vbNewLine & _
                                    "  (Select 1" & vbNewLine & _
                                    "       From ѪҺ�շ���¼ c, ѪҺ��Ѫ��¼ b" & vbNewLine & _
                                    "       Where Mod(c.��¼״̬, 3) = 1 And c.����� Is Not Null And b.Id = c.�䷢id And b.����id = a.����id And b.��ҳid = a.��ҳid)"

                            Else '����
                                strSQL = strSQL & " Union ALL" & vbNewLine & _
                                    " Select a.����id || '-' || a.Id As Id, a.����id, a.Id As ��ҳid, a.����� As ������, '��' As סԺ���, a.����," & vbNewLine & _
                                    "       a.�Ա� || '/' || a.���� As �Ա������, a.ִ�в���id As ����id, d.���� As ��������, a.ִ��ʱ�� As ��Ժ����, '' As ����, a.����, '' As ����, 255 As ��ɫ," & vbNewLine & _
                                    "       a.ִ���� As ��д��" & vbNewLine & _
                                    " From ���ű� d, ���˹Һż�¼ a" & vbNewLine & _
                                    " Where a.ִ�в���id = d.Id And a.����id = [6] And a.Id = [7] And Exists" & vbNewLine & _
                                    "  (Select 1" & vbNewLine & _
                                    "       From ѪҺ�շ���¼ c, ѪҺ��Ѫ��¼ b, ����ҽ����¼ e" & vbNewLine & _
                                    "       Where Mod(c.��¼״̬, 3) = 1 And c.����� Is Not Null And b.Id = c.�䷢id And b.����id = e.Id And e.������� = 'K' And" & vbNewLine & _
                                    "             e.����id = a.����id And e.�Һŵ� = a.No)"
                            End If
                        End If
                        Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), lng����ID, mlng�ύ״̬, UserInfo.����, lng����ID, lng��ҳid)
                End Select
                
                Call RsTitelCopy(rsBR, mRsBR)
                
                With mRsBR
                    If rsBR.RecordCount > 0 Then '��ǰû�ж�rsbr���������жϻᱨ��
                        For lngi = 0 To rsBR.RecordCount - 1
                            .AddNew
                            For lngj = 0 To rsBR.Fields.Count - 1
                                .Fields(lngj).Value = rsBR.Fields(lngj).Value
                                
                                If .Fields(lngj).name = "��Ժ����" Then '�����ڽ���ģʽ������Ҫ��Ȼ������ʾ��������
                                    .Fields(lngj).Value = Format(rsBR.Fields("��Ժ����").Value, "YYYY-MM-DD HH:mm:ss")
                                End If
                                
                                If .Fields(lngj).name = "��ɫ" Then '���¸������ͺ����������ɫ
                                    If Not IsNull(rsBR!����) And Len(rsBR!����) > 0 Then
                                        '������ɫ
                                        lngColor = gobjDatabase.GetPatiColor(Nvl(rsBR!����))
                                        .Fields("��ɫ").Value = lngColor
                                    End If
                                End If
                            Next
                            .Update
                            rsBR.MoveNext
                        Next
                        rsBR.MoveFirst
                        If .RecordCount > 0 Then
                            .MoveFirst
                        End If
                    End If
                End With
                
            Case "ˢ������"
                Dim rsTemp As ADODB.Recordset
                Dim StrPosition As String
                mblnFirst = True
                mlng�ύ״̬ = 0
                mblnIsSubmit = True
                
                If InStr(1, cbo1.Text, "-") > 0 Then
                    Arr���� = Split(cbo1.Text, "-")
                    mstr���� = Arr����(1)
                Else
                    mstr���� = cbo1.Text
                End If

                Call ExecuteCommand("�������˲�ѯ")

                mArr��������(3) = mlng�ύ״̬ '�ύ״̬
'                Set rsTemp = mRsBR
                Call CopyRecord(mRsBR, rsTemp)
                
                If rsTemp.RecordCount > 0 Then
                    mblnHaveBR = True
'                    rsTemp.MoveFirst
                Else
                    mblnHaveBR = False
                End If
                UCP.ShowPeople rsTemp, True
                Call mclsVsf.LoadGrid(rsTemp) '�����ݷ������صĴ�ӡ�б��С�
                
                Set rsTemp = Nothing

                StrPosition = mArrPosition(mlngtbcIndex)
                UCP.SetCardFocus "����id'��ҳid", StrPosition
            Case "ˢ����ʾ"
                If mlng�׶� <> 2 Then Exit Function
                With rptTips
                    Set rsSAD = GetReactionTips(Val(cbo1.ItemData(cbo1.ListIndex)))
                    .Records.DeleteAll
                    .Populate
                    If rsSAD.RecordCount <> 0 Then
                        rsSAD.MoveFirst
                        Call LoadRptTips(rsSAD)
                    End If
                End With
            Case "���ز�������"
                ExecuteCommand = frmBloodReactionRecordSetup.ShowPara(Me)
                mintNotify = gobjDatabase.GetPara("��Ϣ���Ѽ��", 2200, 1938, 0)
        End Select
    Next
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    
    ExecuteCommand = False
End Function

Private Sub LoadRptTips(rsData As Recordset)
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String, rsTmp As New Recordset
    Dim strTmp As String
    rptTips.Records.DeleteAll
    strTmp = ""
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount = 0 Then Exit Sub
    rsData.MoveFirst
    Do While Not rsData.EOF
        If InStr(1, strTmp & ",", "," & Val(Nvl(rsData!����id, 0))) & "," = 0 Then strTmp = strTmp & "," & Val(Nvl(rsData!����id, 0))
        rsData.MoveNext
    Loop
    If strTmp = "" Then Exit Sub
    strSQL = "select /*+ CARDINALITY(b,10) */ a.����,a.����id from ������Ϣ a,table(f_str2list([1],',')) b where a.����id = b.column_value"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", strTmp)
    strTmp = ""
    rsData.MoveFirst
    Do While Not rsData.EOF
        If InStr(1, strTmp & "|", "|" & Val(Nvl(rsData!����id, 0)) & "|") = 0 Then
            Set objRecord = Me.rptTips.Records.Add()
            rsTmp.Filter = "����id = " & Val(Nvl(rsData!����id, 0))
            Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!����)))
            Set objItem = objRecord.AddItem(CStr(rsData!��Ϣ����))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!�շ�ID, 0)))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!����id, 0)))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!����id, 0)))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!������Դ, 0)))
            strTmp = strTmp & "|" & Val(Nvl(rsData!����id, 0))
        End If
        rsData.MoveNext
    Loop
    rptTips.Populate
End Sub
Private Sub RsTitelCopy(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '���ܣ��½�ToRs��¼������RsProm�Ľṹ���Ƶ�ToRs��
    '������RsProm-ԭ��¼����ToRs-�½��ļ�¼��
    Dim lngi As Long
    Set ToRs = New ADODB.Recordset
    With ToRs '��ʼ��rsReturn
        For lngi = 0 To RsProm.Fields.Count - 1
            .Fields.Append RsProm.Fields(lngi).name, adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub CopyRecord(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '���ܣ�����¼��RsProm�Ľṹ�������ݶ����Ƹ�ToRs
    '������RsProm-Ҫ��ֵ�ļ�¼����ToRs-Ŀ���¼��
    Dim lngi As Long
    Dim lngj As Long
    Call RsTitelCopy(RsProm, ToRs)
    With ToRs
        If RsProm.RecordCount > 0 Then '��ǰû�ж�rsbr���������жϻᱨ��
            For lngi = 0 To RsProm.RecordCount - 1
                .AddNew
                For lngj = 0 To RsProm.Fields.Count - 1
                    .Fields(lngj).Value = RsProm.Fields(lngj).Value
                Next
                .Update
                RsProm.MoveNext
            Next
            RsProm.MoveFirst
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
    End With
End Sub

Private Sub chk1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub ChkRection_Click()
    Call pic1_Resize
End Sub

Private Sub ChkRection_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd2_Click()
    RefreshBR
    pic1_Resize
End Sub

Private Sub dkpPeoPle_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
        Case 1
            Item.Handle = pic1.hWnd
        Case 2
            Item.Handle = picTips.hWnd
        Case 3
            Item.Handle = pic2.hWnd
    End Select
End Sub

Private Sub DTP1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTP2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTP3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTP4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpPeoPle, 1, 260, 100, 320, Me.ScaleHeight)
    Call SetPaneRange(dkpPeoPle, 2, 260, 40, 320, 100)
    Call SetPaneRange(dkpPeoPle, 3, 100, 100, Me.ScaleWidth, Me.ScaleHeight)
    dkpPeoPle.RecalcLayout
End Sub

Public Function SetPaneRange(dkpM As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�����dockingpane�Ĵ�С��Χ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpM.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If UCE.strST = ���� Or UCE.strST = �޸� Then
        Cancel = (MsgBox("����δ���棬�Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    mblnBloodReactionRecordIsOpen = False
    mblnStart = False
End Sub

Private Sub initTabControl(Index As Long)
    '���ܣ���ʼ��tbcthis
    With tbcthis
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .COLOR = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = False
        End With
        
        Select Case Index
            Case 0:
                .InsertItem(0, "���ھ���", picTmp.hWnd, 0).Tag = "���ھ���"
                .InsertItem(1, "��ɾ���", picTmp.hWnd, 0).Tag = "��ɾ���"
                .Item(0).Selected = True
            Case 1:
                .InsertItem(0, "��Ժ", picTmp.hWnd, 0).Tag = "��Ժ"
                .InsertItem(1, "ת��", picTmp.hWnd, 0).Tag = "ת��"
                .InsertItem(2, "��Ժ", picTmp.hWnd, 0).Tag = "��Ժ"
                .Item(0).Selected = True
        End Select
    End With
End Sub

Private Sub opt_Click(Index As Integer)
    opt(2).Tag = Index
End Sub
Private Sub picTips_Resize()
    rptTips.Move 0, 0, picTips.Width, picTips.Height
End Sub
Private Sub picUCP_Resize()
    '���ܣ�����ҳ��Ĳ���
    On Error Resume Next
    UCP.Move 0, 0, picUCP.ScaleWidth, picUCP.ScaleHeight
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '���ܣ�tbcthis�ؼ��л�ѡ���ˢ������
    Dim ArrReturn
    ArrReturn = Split(UCP.strReturn, "'")
    If UBound(ArrReturn) >= 0 Then
        mArrPosition(mlngtbcIndex) = Val(ArrReturn(1)) & "'" & Val(ArrReturn(3)) '��¼��λ��Ϣ����"����id'��ҳid"����ʽ��ArrReturn(1)��ArrReturn(3)�ֱ������id����ҳid
    End If
    mlngtbcIndex = Item.Index
    pic1_Resize
    If mblnStart = True Then
        Call RefreshBR
    End If
End Sub

Private Sub initComboTime()
    '���ܣ���cobtime�����dtp�ؼ����г�ʼ��
    cbotime.Clear '��Ժ
    With cbotime
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 6
        .AddItem "30����"
        .ItemData(.NewIndex) = 29
        .AddItem "60����"
        .ItemData(.NewIndex) = 59
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cbotime.ListCount > 0 Then cbotime.ListIndex = 0
End Sub

Private Sub TXTDay_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = True Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
    Else
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub UCP_CardChanged()
    '���ܣ�usrCardPeople�ؼ��л����˺�ͬʱˢ�¸ò��˵ķ�Ӧ��¼
    Dim ArrReturn
    Dim lngi As Long
    Dim lng�׶� As Long
    Dim rsData As ADODB.Recordset
    Dim strFilter As String
    Dim blnHaveRection As Boolean
    
    On Error GoTo Errorhand:

    Set rsData = mRsBR
    '�������������ϳ��ַ��������㴫��
    strFilter = mArr��������(0) & "|" & mArr��������(1) & "|" & mArr��������(2) & "|" & mArr��������(3) & "|" & mArr��������(4)
    
    lng�׶� = mlng�׶�
    If lng�׶� = 0 Then '������׶κ�סԺ�׶�ͳһΪҽ���׶Σ���������Ϊ1
        lng�׶� = 1
    End If
    
    ArrReturn = Split(UCP.strReturn, "'")
    If UBound(ArrReturn) >= 0 Then
        mArrPosition(mlngtbcIndex) = Val(ArrReturn(1)) & "'" & Val(ArrReturn(3)) '��¼��λ��Ϣ����"����id'��ҳid"����ʽ��ArrReturn(1)��ArrReturn(3)�ֱ������id����ҳid
    End If
    '������Ѫ��Ӧ��¼
    If UBound(ArrReturn) = -1 Then
        UCE.ShowClear
        mblnIsSelect = False
    ElseIf rsData.RecordCount > 0 Then
        With rsData
            .MoveFirst
            For lngi = 0 To .RecordCount - 1
                If .Fields("����ID").Value = Val(ArrReturn(1)) And .Fields("��ҳid").Value = Val(ArrReturn(3)) Then 'סԺ������ҳid������ҳid�����ﲡ�˵Ļ�����ҳid��������
                    If .Fields("סԺ���").Value = "ס" Then
                        UCE.showInfor Val(.Fields("����ID").Value), 2, Val(IIf(IsNull(.Fields("��ҳid").Value) = True, 0, .Fields("��ҳid").Value)), lng�׶�, gcnOracle, Me, p��Ѫ��Ӧ����, strFilter, IsPrivs(mstrPrivs, "��Ѫ������")
                    Else  '���ﲡ��ʹ������������ҳid�����ֶ�ιҺŵĲ���
                        UCE.showInfor Val(.Fields("����ID").Value), 1, Val(IIf(IsNull(.Fields("��ҳid").Value) = True, 0, .Fields("��ҳid").Value)), lng�׶�, gcnOracle, Me, p��Ѫ��Ӧ����, strFilter, IsPrivs(mstrPrivs, "��Ѫ������")
                    End If
                End If
                .MoveNext
            Next
            .MoveFirst
        End With
        mblnIsSelect = True
    End If
    Set rsData = Nothing
Errorhand:
End Sub

Private Sub pic1_Resize()
    Dim intType As Integer
    On Error Resume Next
    

    lbl1.Left = 120
    cbo1.Left = lbl1.Left + lbl1.Width + 90
    cbo1.Top = 60
    lbl1.Top = cbo1.Top + (cbo1.Height - lbl1.Height) \ 2
    
    Fra1.Left = 60
    Fra1.Top = cbo1.Top + cbo1.Height + 60
    Fra1.Width = pic1.ScaleWidth - 60
    
    cmd2.Visible = True
    If mlng�׶� = 2 Then '��Ѫ��
        '��ʾ��ʾ
        picTips.Visible = True
        '����ʾChkRection�ؼ�
        ChkRection.Visible = False
        '��ʾ��Ѫ��Ӧ���������ؼ���������λ��
        chk1(0).Visible = True
        chk1(1).Visible = True
        chk1(2).Visible = True
        tbcthis.Visible = False
        
        '��Ѫ��Ӧʱ��
        lbl2.Left = 120
        DTP1.Left = lbl2.Left + lbl2.Width + 90
        DTP1.Top = 210
        lbl2.Top = DTP1.Top + (DTP1.Height - lbl2.Height) \ 2
        DTP2.Left = DTP1.Left
        DTP2.Top = DTP1.Top + DTP1.Height + 60
        lbl4.Left = lbl2.Left + lbl2.Width - lbl4.Width
        lbl4.Top = DTP2.Top + (DTP2.Height - lbl4.Height) \ 2
        
        chk1(0).Left = DTP1.Left
        chk1(0).Top = DTP2.Top + DTP2.Height + 60
        chk1(1).Top = chk1(0).Top
        chk1(2).Top = chk1(0).Top
        
        chk1(1).Value = Checked
        chk1(2).Value = Checked
        
        fra.Top = chk1(0).Top + chk1(0).Height + 60
        cmd2.Top = fra.Top + fra.Height + 60
        
        Fra1.Height = cmd2.Top + cmd2.Height + 120
        
        picUCP.Left = Fra1.Left
        picUCP.Top = Fra1.Top + Fra1.Height + 60
        picUCP.Width = Fra1.Width
        If pic1.ScaleHeight - Fra1.Top - Fra1.Height > 0 Then
            picUCP.Height = pic1.ScaleHeight - Fra1.Top - Fra1.Height
        End If
        '�����������õ�����
        lbl5.Visible = False
        cbotime.Visible = False
        lbl8.Visible = False
        DTP3.Visible = False
        lbl9.Visible = False
        DTP4.Visible = False
        lbl6.Visible = False
        TXTDay.Visible = False
        frmLine.Visible = False
    Else 'סԺ��������
        '����ʾ��ʾ
        picTips.Visible = False
        ChkRection.Visible = True
        If (mlng�׶� = 1 Or mlng�׶� = 0) And mlngtbcIndex = 0 Then '��Ժ�������ھ���
            '�ؼ���ʾ����
            lbl5.Visible = False
            cbotime.Visible = False
            lbl8.Visible = False
            DTP3.Visible = False
            lbl9.Visible = False
            DTP4.Visible = False
            lbl6.Visible = False
            TXTDay.Visible = False
            frmLine.Visible = False
            ChkRection.Left = 120
            ChkRection.Top = 240
            intType = 0
        ElseIf (mlng�׶� = 1 And mlngtbcIndex = 2) Or (mlng�׶� = 0 And mlngtbcIndex = 1) Then  '��Ժ������ɾ���
            '�ؼ���ʾ����
            lbl5.Visible = True
            cbotime.Visible = True
            lbl8.Visible = True
            DTP3.Visible = True
            lbl9.Visible = True
            DTP4.Visible = True
            lbl6.Visible = False
            TXTDay.Visible = False
            frmLine.Visible = False
            
            lbl5.Left = 120
            cbotime.Left = lbl5.Left + lbl5.Width + 90
            cbotime.Top = 210
            
            lbl5.Top = cbotime.Top + (cbotime.Height - lbl5.Height) \ 2
            DTP3.Left = cbotime.Left
            DTP3.Top = cbotime.Top + cbotime.Height + 60
            lbl8.Left = lbl5.Left
            lbl8.Top = DTP3.Top + (DTP3.Height - lbl8.Height) \ 2
            DTP4.Left = DTP3.Left
            DTP4.Top = DTP3.Top + DTP3.Height + 60
            lbl9.Left = lbl5.Left
            lbl9.Top = DTP4.Top + (DTP4.Height - lbl9.Height) \ 2
            ChkRection.Left = 120
            ChkRection.Top = DTP4.Top + DTP4.Height + 60
            
            If mlng�׶� = 1 Then
                lbl5.Caption = "��Ժ����"
            ElseIf mlng�׶� = 0 Then
                lbl5.Caption = "��������"
            End If
            intType = 1
        ElseIf mlng�׶� = 1 And mlngtbcIndex = 1 Then 'ת��
            '�ؼ���ʾ����
            lbl5.Visible = False
            cbotime.Visible = False
            lbl8.Visible = False
            DTP3.Visible = False
            lbl9.Visible = False
            DTP4.Visible = False
            lbl6.Visible = True
            TXTDay.Visible = True
            frmLine.Visible = True
            
            lbl6.Left = 120
            lbl6.Top = 240
            TXTDay.Left = lbl6.Left + 810
            TXTDay.Top = lbl6.Top
            frmLine.Left = TXTDay.Left
            frmLine.Top = TXTDay.Top + TXTDay.Height + 15
            ChkRection.Left = 120
            ChkRection.Top = lbl6.Top + lbl6.Height + 120
            intType = 2
        End If
        If ChkRection.Value = 0 Then
            chk1(0).Visible = False
            chk1(1).Visible = False
            chk1(2).Visible = False
            Select Case intType
                Case 0
                    cmd2.Visible = False
                    Fra1.Height = ChkRection.Top + ChkRection.Height + 120
                Case Else
                    cmd2.Top = ChkRection.Top + ChkRection.Height + 60
                    Fra1.Height = cmd2.Top + cmd2.Height + 120
            End Select
        Else
            chk1(0).Visible = True
            chk1(1).Visible = True
            chk1(2).Visible = True
            
            chk1(0).Left = ChkRection.Left + 180
            chk1(0).Top = ChkRection.Top + ChkRection.Height + 60
            chk1(1).Left = chk1(0).Left + chk1(0).Width
            chk1(1).Top = chk1(0).Top
            chk1(2).Left = chk1(1).Left + chk1(1).Width
            chk1(2).Top = chk1(0).Top
            
            chk1(1).Value = Checked
            chk1(2).Value = Checked
            
            cmd2.Top = chk1(0).Top + chk1(0).Height + 60
            Fra1.Height = cmd2.Top + cmd2.Height + 120
        End If
        
        lbl2.Visible = False
        lbl4.Visible = False
        DTP1.Visible = False
        DTP2.Visible = False
        
        tbcthis.Left = Fra1.Left
        tbcthis.Top = Fra1.Top + Fra1.Height + 60
        tbcthis.Width = Fra1.Width
        tbcthis.Height = 350
        
        picUCP.Left = Fra1.Left
        picUCP.Top = tbcthis.Top + tbcthis.Height
        picUCP.Width = Fra1.Width
        If pic1.ScaleHeight - tbcthis.Top - tbcthis.Height > 0 Then
            picUCP.Height = pic1.ScaleHeight - tbcthis.Top - tbcthis.Height
        End If
    End If
End Sub

Private Sub pic2_Resize()
    On Error Resume Next
    UCE.Move 0, 0, pic2.Width, pic2.ScaleHeight
End Sub
