VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.0#0"; "zlIDKind.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPathologyStation 
   Caption         =   "Ӱ������վ"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   11400
   Icon            =   "frmPathologyStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtLocate 
      Height          =   300
      Left            =   5040
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox PicWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   1200
      ScaleHeight     =   3495
      ScaleWidth      =   10035
      TabIndex        =   1
      Top             =   3600
      Width           =   10035
      Begin VB.PictureBox picVideoContainer 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2055
         Left            =   4800
         ScaleHeight     =   1995
         ScaleWidth      =   3555
         TabIndex        =   18
         Top             =   840
         Width           =   3615
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   625
         Left            =   0
         ScaleHeight     =   630
         ScaleWidth      =   9990
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   9990
         Begin VB.Frame fraInfo 
            ForeColor       =   &H00000000&
            Height          =   700
            Left            =   2040
            TabIndex        =   7
            Top             =   0
            Width           =   7860
            Begin VB.Label lblCash 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   540
               Left            =   6945
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lbl������Ϣ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������Ϣ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   250
               Left            =   90
               TabIndex        =   9
               Top             =   150
               Width           =   900
            End
            Begin VB.Label lbl�����Ϣ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����Ϣ"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   90
               TabIndex        =   8
               Top             =   450
               Width           =   720
            End
         End
         Begin VB.Frame fraRegist 
            Height          =   700
            Left            =   15
            TabIndex        =   4
            Top             =   -75
            Width           =   1980
            Begin VB.ComboBox cboTimes 
               Height          =   300
               Left            =   60
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   340
               Width           =   1875
            End
            Begin VB.Label lblRegist 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����¼(&G)"
               Height          =   180
               Left            =   95
               TabIndex        =   6
               Top             =   140
               Width           =   990
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   45
      ScaleHeight     =   4275
      ScaleWidth      =   4500
      TabIndex        =   12
      Top             =   525
      Width           =   4495
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   661
         _StockProps     =   64
      End
      Begin VB.TextBox txtAppend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1605
         Width           =   2010
      End
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   250
         Left            =   870
         TabIndex        =   13
         ToolTipText     =   "*����ţ�+סԺ�ţ�����ѡ���ҷ�ʽ������+��*��Ϊģ����ѯ��������ɺ�ֱ�ӻس���ʼ����"
         Top             =   45
         Width           =   1485
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   3360
         _cx             =   5927
         _cy             =   4736
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
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
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Left            =   2730
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   270
            Visible         =   0   'False
            Width           =   270
         End
      End
      Begin XtremeCommandBars.CommandBars cbrdock 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   720
   End
   Begin zlIDKind.IDKind IDKind 
      Bindings        =   "frmPathologyStation.frx":1CFA
      Height          =   360
      Left            =   5010
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathologyStation.frx":1D0E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7938
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   6690
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":25A2
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":2B3C
            Key             =   "סԺ"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":3416
            Key             =   "����"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":3570
            Key             =   "Ӱ��"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":3CEA
            Key             =   "�շ�"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":4084
            Key             =   "��ɫͨ��"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":41DE
            Key             =   "·��"
            Object.Tag             =   "7"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5955
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":4778
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":4D12
            Key             =   "��ѡ����"
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":5064
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90003"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathologyStation.frx":53E6
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
      EndProperty
   End
   Begin DicomObjects.DicomViewer dcmRelateViewer 
      Height          =   1095
      Left            =   6240
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
      _Version        =   262147
      _ExtentX        =   4471
      _ExtentY        =   1931
      _StockProps     =   35
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPathologyStation.frx":5980
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPathologyStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mintCurҵ������ As Integer = 1 '��ǰϵͳ������ҵ������

Private Const ConstrCol = "·��;400|����;300|��Դ;400|�շ�;300|����;300|����;300|����;1200|�����;800|����ִ�й���;1400|������;800|�Ա�;450|����;450" & _
                        "|��ʶ��;1400|ҽ������;2400|��λ����;1400|����ʱ��;1800|����ʱ��;1800|����ҽ��;800" & _
                        "|���;450|����;450|Ӥ��;450|�Ǽ���;800|������;800|�����;800|�������;800" & _
                        "|��ɫͨ��;0|�����ӡ;800|������;800|������;800|��ͼʱ��;1800|�������;2400|����;1400|�������;1200" & _
                        "|������;0|����ID;0|��ҳID;0|�Һŵ�;0|���˿���ID;0|ҽ��ID;1200|���ͺ�;0|���UID;0" & _
                        "|���״̬;0|NO;0|��¼����;0|ת��;0|����;0|��ǰ����ID;0|���淢��;800|��Ϸ���;800|����ID;0" & _
                        "|���˿���;800|���￨��;800|���ݺ�;800|���֤��;800"
Private mstrCol As String   '�б�˳�������ʱ��ȡע�������ֵ��ConstrColΪĬ��ֵ

'ID_���ҷ�ʽ+100֮����7������Ϊ���ҷ�ʽѡ���
'ID_Ӱ�����֮����40��������ΪӰ����𣬴�4021-4060
Private Enum FilterID
    ID_���� = 4001: ID_סԺ = 4002: ID_��� = 4003: ID_���� = 4004
    ID_���� = 4005: ID_�ѽ� = 4006: ID_δ�� = 4007: ID_�Ǽ� = 4008
    ID_���� = 4009: ID_��� = 4010: ID_���� = 4011: ID_��� = 4012
    ID_��� = 4013
    ID_���ҷ�ʽ = 4014: ID_����ֵ = 4015: ID_��ʼ���� = 4016: ID_����סԺ = 4017
    
    
    ID_������� = 4100
    ID_�������_���� = 4101: ID_�������_���� = 4102: ID_�������_ϸ�� = 4103: ID_�������_ʬ�� = 4104: ID_�������_���� = 4105
    
    ID_�걾���� = 4110: ID_�걾����_���� = 4111: ID_�걾����_С�걾 = 4112: ID_�걾����_���� = 4113: ID_�걾����_���� = 4114: ID_�걾����_Һ�� = 4115
End Enum

Private mblncmd���� As Boolean, mblncmdסԺ As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean, mblncmd�ѽ� As Boolean, mblncmdδ�� As Boolean
Private mblncmd�Ǽ� As Boolean, mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd��� As Boolean
Private mblncmd���� As Boolean


Private mblncmd���� As Boolean
Private mblncmdС�걾 As Boolean
Private mblncmd���� As Boolean
Private mblncmd���� As Boolean
Private mblncmdҺ�� As Boolean


Private mblncmd���� As Boolean
Private mblncmdϸ�� As Boolean
Private mblncmd���� As Boolean
Private mblncmdʬ�� As Boolean
Private mblncmd���� As Boolean


Private mstrFirstTab As String '�״���ʾ��ҳ��

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0��������￨ = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
End Enum

'�Ӵ������
Private WithEvents mfrmPacsReport As frmReport                          'PACS����༭����Ƕ��������Ĵ���
Attribute mfrmPacsReport.VB_VarHelpID = -1
Private WithEvents mfrmPacsReportDock As frmReport                      'PACS����༭��,��������
Attribute mfrmPacsReportDock.VB_VarHelpID = -1
Private WithEvents mobjReport As zlRichEPR.cDockReport                  '�������
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mobjInAdvice As zlCISKernel.clsDockInAdvices         'סԺҽ������
Attribute mobjInAdvice.VB_VarHelpID = -1
Private WithEvents mobjOutAdvice As zlCISKernel.clsDockOutAdvices       '����ҽ������
Attribute mobjOutAdvice.VB_VarHelpID = -1
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer                '��Ƭվ����
Attribute mobjPacsCore.VB_VarHelpID = -1

Private WithEvents mfrmPatholSpecimen As frmPatholSpecimen              '�걾����
Attribute mfrmPatholSpecimen.VB_VarHelpID = -1
Private WithEvents mfrmPatholMaterial As frmPatholMaterials             'ȡ��
Attribute mfrmPatholMaterial.VB_VarHelpID = -1
Private WithEvents mfrmPatholSlices As frmPatholSlices                  '��Ƭ
Attribute mfrmPatholSlices.VB_VarHelpID = -1
Private WithEvents mfrmPatholSpeExam As frmPatholSpecialExamined        '�ؼ�
Attribute mfrmPatholSpeExam.VB_VarHelpID = -1
Private mfrmPatholProRep As frmPatholProcedureRep                       '���̱���
Private mfrmPatholDecalinTask As New frmPatholDecalcification           '�Ѹ�����


Private mobjExpense As zlCISKernel.clsDockExpense       '���ö���
Private mobjInEPRs As zlRichEPR.cDockInEPRs             'סԺ��������
Private mobjOutEPRs As zlRichEPR.cDockOutEPRs           '���ﲡ������
Private mobjQueue As zlQueueManage.clsQueueManage          '�Ŷӽк�

Private mobjPacsReportArry() As frmReport                   'PACS����༭������


'���ڱ���
Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrCur���� As String                               '��ǰ���� ����-����
Private mstrCanUse���� As String                            '��ǰ���ÿ���  ID_����-����
Private mstrCurFindtype As String                           '��������
Private mlngFilterTab As Long                               '����tabҳ
Private mstrLocateType As String                            '��λ����
Private mblnInitOk As Boolean, mblnvsRefresh As Boolean     '��ʼ�����,װ�ر��
Private mstrPrivs As String, mlngModul As Long              'ģ��ţ���ģ��Ȩ��
Private mlngSortCol As Long                                 '�����б��У���ǰ�����������
Private mintSortOrder As Integer                            '�����б��У���ǰ��������ķ�ʽ

'���̿��Ʊ���
Private mblnFinishCommit As Boolean                         '�ޱ��������,�Ƿ������ٴ�ȷ��
Private mblnCompleteCommit As Boolean                       '��˺������ٴ�ȷ��
Private mblnIgnoreResult As Boolean                         '���������� '=true ����
Private mintResultInput As Integer                          '��ʾ���������Ժ�Ӱ������
Private mblnReportWithImage As Boolean                      '��ͼ�����д���棬��ͼ�񲻿�д����
Private mblnReportWithResult As Boolean                     '��Ӱ�����Ϊ����
Private mblnLocalizerBackward As Boolean                    '��λƬ����
Private mblnPacsReport As Boolean                           '�Ƿ�ʹ��PACS����༭����Fasleʱʹ�õ��Ӳ����༭��
Private mblnPrintCommit As Boolean                          '��ӡ��ֱ�����
Private mblnCanPrint As Boolean                             'ƽ����Ҫ��˲��ܴ�ӡ =true
Private mBeforeDays As Integer                              'Ĭ�ϲ�ѯ������
Private mlngRefreshInterval As Long                         '�����б��Զ�ˢ�¼��
Private mAstr��������() As String                           '�������ƣ�ִ�м������
Private mblnRelatingPatient As Boolean                      '�Ƿ����ù�������
'��������
Private mstrRoom As String                                  'ֻ����ִ�м��ڵĲ���
Private mblnPatTrack As Boolean                             '�Ƿ�Խ����˽��и���
Private mblnֱ�Ӽ�� As Boolean                             '�ǼǺ�ֱ�ӽ�����
Private mblnNoShowCancel As Boolean                         '����ʾȡ���ļ��
Private mblnMoved As Boolean                                '��ǰʱ������Ƿ�ת�ƹ�
Private mblnDockVideo As Boolean                            '�Ƿ�ʹ�ø������ڲɼ�ͼ��True-��������mfrmDockVideo��False��Ƕ�봰��mfrmCapture
Private mblnOpenReport As Boolean                           '��ʼ����Զ��򿪱���
Private mblnWriteCapDoctor As Boolean                       '�Ƿ��ڲɼ�ͼ����Զ��ѵ�ǰ�û���дΪ��鼼ʦ
Private mblnTechReptSame As Boolean                         'ֻ����д�Լ����ı���
Private mblnPacsReportShowVideoCapture As Boolean           '��PACS����༭���У��Ƿ���ʾ��Ƶ�ɼ�����


'������������
Private Type Type_SQLCondition
    ��ʼʱ�� As Date
    ����ʱ�� As Date
    ʱ������ As Integer                                 'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
    ���ݺ� As String
    ����� As Double
    סԺ�� As Double
    ���￨ As String
    ���� As String
    �Ա� As String
    ��ʼ���� As Long
    �������� As Long
    �������� As String
    ���� As Double
    ����� As String
    ���֤  As String
    IC�� As String
    ���˿��� As Long
    �걾��λ As String
    ���ҽ�� As String
    ���ҽ�� As String
    ������� As String
    �������� As String
    ������� As Integer
    Ӱ������ As String
    ��鼼ʦ As String
    ������ As String
    ������� As String
    ������ As String
    ���� As String
    ��� As String
End Type

Private SQLCondition As Type_SQLCondition
Private WithEvents mobjSysHook As clsHookKey '���õ�ǰ�����HOOK
Attribute mobjSysHook.VB_VarHelpID = -1

'��ʷ��¼����ʾ
Private mblnIsHistory As Boolean
Private mlngHOrderID As Long
Private mlngHSendNo As Long
Private mstrHStudyUID As String
Private mblnHMoved As Boolean

'�Ŷӽк�


Private Sub Menu_File_Excel_click()
Dim bytMode As Byte
   
    On Error GoTo errHandle
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vsList
    objPrint.Title.Text = "��鲡���嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub Menu_File_BatPrint()
    Dim cbrControl As CommandBarControl, strReturn As String, i As Integer
    Dim objReportPrint As New zlRichEPR.cDockReport
    Dim objPacsReport As New frmReport
    Dim strReportString As String

    Set cbrControl = Me.cbrMain(2).FindControl(, conMenu_File_Print)
    If Not cbrControl Is Nothing Then
        cbrControl.ID = conMenu_File_BatPrint
    Else
        Exit Sub
    End If

    'ѡ����
    strReturn = frmDocPrintPatiList.Showfrm(vsList, Me, mblnCanPrint, mblnPacsReport, mlngCur����ID)
    
    'ѭ�����ñ����ӡ
    '����ʹ��PACS����༭����ӡ�ģ�����PACS����༭����������ӡ
    '����ֵ��"ҽ��ID-�Ƿ�PACS����༭��-ִ�п���ID|ҽ��ID-�Ƿ�PACS����༭��-ִ�п���ID|..."���
    For i = 0 To UBound(Split(strReturn, "|"))
        strReportString = Split(strReturn, "|")(i)
        '�ж��Ƿ�ʹ��PACS����༭��
        If Split(strReportString, "-")(1) = 1 Then  'ʹ��PACS����༭��
            Call objPacsReport.InitReportWindow(CLng(Split(strReportString, "-")(2)), mlngModul, mstrPrivs, True) '���һ���������Ϊtrue���ɲ���ʾ��Ƶ�ɼ�
            objPacsReport.zlRefresh CLng(Split(strReportString, "-")(0)), Me, False, ""
            Call objPacsReport.zlExecuteCommandBars(cbrControl)
            '��ҪAfterPrint��
        Else    'ʹ�ò����༭��
            If objReportPrint.zlRefresh(CLng(Split(strReportString, "-")(0)), CLng(Split(strReportString, "-")(2)), , , True) > 0 Then
                Call objReportPrint.zlExecuteCommandBars(cbrControl)
                Call AfterPrinted(CLng(Split(strReportString, "-")(0)))
            End If
        End If
    Next
    
    cbrControl.ID = conMenu_File_Print
    Unload objReportPrint.zlGetForm
End Sub


Private Sub Menu_RichEPR(ByVal cbrID As Long)
    Dim cbrControl As CommandBarControl, i As Integer, blnCanPrint As Boolean
    
    '����ҳ�治�ɼ�ʱ��ִ���κβ���
    If TabWindow.Selected.Tag <> "������д" Then
        For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
            If TabWindow(i).Tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.Tag <> "������д" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    With vsList
        blnCanPrint = IIf(mblnCanPrint, IIf(.Cell(flexcpData, .Row, GetCN("����")) = 1, .TextMatrix(.Row, GetCN("������")) <> "", .TextMatrix(.Row, GetCN("������")) <> ""), True)
        'ˢ��Ƕ��ҳ������
        If mblnPacsReport = True Then
            Call mfrmPacsReport.zlRefresh(Val(.TextMatrix(.Row, GetCN("ҽ��ID"))), Me, .TextMatrix(.Row, GetCN("ת��")) = 1, .TextMatrix(.Row, GetCN("����")))
        Else
            Call mobjReport.zlRefresh(Val(.TextMatrix(.Row, GetCN("ҽ��ID"))), mlngCur����ID, True, .TextMatrix(.Row, GetCN("ת��")) = 1, blnCanPrint)
        End If
    End With
    
    '�жϰ���������
    Set cbrControl = Me.cbrMain.FindControl(, IIf(mblnPacsReport, conMenu_PacsReport_Open, cbrID))
    If cbrControl Is Nothing Then Exit Sub
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    Call cbrMain_Execute(cbrControl)
End Sub

Private Sub Menu_File_Parmeter_click()
    With frmTechnicSetup
        .mlngModul = mlngModul
        .mlng����ID = mlngCur����ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        If .mblnOK Then
            InitLocalPars
            Call RefreshList
        End If
    End With
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '���ܣ����ð�������
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub

Private Sub Menu_Manage_ȡ������()
'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
Dim strFilter As String, rsTmp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    On Error GoTo errHandle
    '��ʾ����ѡ�񴰿�
    With vsList
        lngAdviceID = Nvl(.TextMatrix(.Row, GetCN("ҽ��ID")), 0)
        lngSendNO = Nvl(.TextMatrix(.Row, GetCN("���ͺ�")), 0)
    End With
    
    gstrSQL = "select 0 as ѡ��,B.����UID as ID ,B.���к�,B.��������,SUM(1) AS ͼ���� from Ӱ�����¼ A ," & _
            "Ӱ�������� B, Ӱ����ͼ�� C Where a.���UID = B.���UID And B.����UID = C.����UID" & _
            " And a.ҽ��ID = [1] and A.���ͺ�= [2] group by B.����UID,B.���к�,B.��������"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceID, lngSendNO)
    
    frmSelectMuli.ShowSelect rsTmp, "ID,3000,0,1;���к�,800,0,1;��������,2000,0,1;ͼ����,800,0,1", 0, 0, 14000, 10000, "ȡ������"
    
    If frmSelectMuli.mblnOK = True Then
        strFilter = frmSelectMuli.strFilter
        rsTmp.Filter = strFilter
        '�����ѡ�����У�����ÿһ�����е�ȡ��
        While Not rsTmp.EOF
            subCancelSeriesRelate Me, lngAdviceID, lngSendNO, rsTmp!ID, True
            rsTmp.MoveNext
        Wend
        
        '����Ӱ����״̬�������ǰҽ���Ѿ�û��ͼ�񣬶��Ҽ�����Ϊ3�����޸�Ϊ2
         If vsList.TextMatrix(vsList.Row, GetCN("���״̬")) = 3 Then
            gstrSQL = "Select ���uid From Ӱ�����¼ Where  ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceID, lngSendNO)
            If IsNull(rsTmp!���uid) Then
                gstrSQL = "Zl_Ӱ����_State(" & lngAdviceID & "," & lngSendNO & ",2)"
                zlDatabase.ExecuteProcedure gstrSQL, "ȡ������"
            End If
        End If
        
        Call RefreshList '����ȡ��������ȷ����ˢ��
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub Menu_Manage_�ޱ������()
'ֻ�н����еı�����Բ����ò˵�,��Ϊ��ʱ��û��ǩ��
        On Error GoTo errHandle
        With vsList
            If .TextMatrix(.Row, GetCN("������")) <> "" Or .TextMatrix(.Row, GetCN("�������")) <> "" Then
                If MsgBoxD(Me, "�Ƿ��ޱ���ֱ�����,ֱ����ɽ�ɾ������д�ı���!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            If mblnFinishCommit And InStr(mstrPrivs, "������") > 0 Then '�ޱ�����ɺ������ٴ�ȷ�����,����Ҫ�м����ɵ�Ȩ��
                '�˹���,��״̬=6,���ұ���ID��Ϊ�ս�ɾ�����Ӳ�����¼
                If bln����δ���(.TextMatrix(.Row, GetCN("����ID")), Val(.TextMatrix(.Row, GetCN("��ҳID"))), _
                    .TextMatrix(.Row, GetCN("ҽ��ID")), CLng(Decode(.TextMatrix(.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4))) Then
                    
                    'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
                    MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵�������ɣ�", vbExclamation, gstrSysName
                Else
                    gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & "," & .TextMatrix(.Row, GetCN("���ͺ�")) & ",6,1)"
                End If
            Else
                gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & "," & .TextMatrix(.Row, GetCN("���ͺ�")) & ",5,1)"
            End If
        End With

        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ı������")
        
            
        If mblnPatTrack Then
            If mblnFinishCommit Then
                Call StateCheck(6)
            Else
                Call StateCheck(5)
            End If
        Else
            Call RefreshList
        End If
        Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Edit_�ޱ������()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If MsgBoxD(Me, "ȷ��Ҫ���˸�������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    With vsList
            '�����ͼ������˵����Ѽ�顱��������˵����ѱ�����
            gstrSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ͼ��", CLng(.TextMatrix(.Row, GetCN("ҽ��ID"))))
            
            gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & "," & .TextMatrix(.Row, GetCN("���ͺ�")) & "," & IIf(Nvl(rsTemp!���uid) = "", 2, 3) & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    
    If mblnPatTrack Then
        Call StateCheck(2)
    Else
        Call RefreshList
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_����������(Optional lngҽ��ID As Long = 0, Optional blnRefresh As Boolean = True)
'�������������̵��ã���ʱ������ҽ��ID������ҪȨ���ж�
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lngҽ��ID = 0 Then
        lngҽ��ID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
    End If
    If InStr(mstrPrivs, "������") <= 0 Then Exit Sub
    
    gstrSQL = "Select a.���ͺ�,b.����ID,b.��ҳID From ����ҽ������ a,����ҽ����¼ b Where a.ҽ��id = [1] And a.ҽ��ID=b.Id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����������", lngҽ��ID)
    
    If rsTemp.EOF = True Then Exit Sub
    
    If bln����δ���(rsTemp!����ID, Nvl(rsTemp!��ҳID, 0), Nvl(lngҽ��ID), _
        CLng(Decode(vsList.TextMatrix(vsList.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4))) Then
       
        'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
        MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵���������ɣ�", vbExclamation, gstrSysName
    Else
    
        Call gcnOracle.BeginTrans
        On Error GoTo errTrans
        
        gstrSQL = "ZL_Ӱ����_STATE(" & lngҽ��ID & "," & rsTemp!���ͺ� & ",6)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ı������")
        
        gstrSQL = "Zl_������_���(" & lngҽ��ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
        
        GoTo errCommit
errTrans:
        Call gcnOracle.RollbackTrans
        GoTo errHandle
errCommit:
        Call gcnOracle.CommitTrans
        
        If blnRefresh Then Call StateCheck(6)
    End If

    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_ȡ��������()
    Dim intState As Integer
    
    On Error GoTo errHandle
    With vsList
            If .TextMatrix(.Row, GetCN("ת��")) = 1 Then
                MsgBoxD Me, "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Call gcnOracle.BeginTrans
            On Error GoTo errTrans
            
            intState = getStudyState(.TextMatrix(.Row, GetCN("ҽ��ID")))
            gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & "," & .TextMatrix(.Row, GetCN("���ͺ�")) & "," & intState & ")"
            zlDatabase.ExecuteProcedure gstrSQL, "ȡ��������"
            
            gstrSQL = "Zl_������_ȡ�����(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "������ȡ�����")
            
            GoTo errCommit
            
errTrans:
    Call gcnOracle.RollbackTrans
    GoTo errHandle
errCommit:
    Call gcnOracle.CommitTrans
            
    End With

    Call StateCheck(intState)
    Exit Sub

errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�������(ByVal lngID As Long)
    Dim iresult As Integer

    On Error GoTo errHandle
    Select Case lngID
        Case conMenu_Manage_Negative
            iresult = 1
        Case conMenu_Manage_Positive
            iresult = 0
    End Select
    With vsList
        gstrSQL = "ZL_Ӱ����_���(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & "," & iresult & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
        
        If iresult = 1 Then
            Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("����").Picture
        Else
            Set .Cell(flexcpPicture, .Row, GetCN("����")) = Nothing
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_��ɫͨ��(ByVal lngID As Long)
    Dim intResult As Integer

    On Error GoTo errHandle
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    With vsList
        gstrSQL = "Zl_��ɫͨ��_Update(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & ",'" & intResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɫͨ��")
        .TextMatrix(.Row, GetCN("��ɫͨ��")) = intResult
        If intResult = 1 Then
            Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("��ɫͨ��").Picture
        Else
            Set .Cell(flexcpPicture, .Row, GetCN("����")) = Nothing
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_Ӱ������(ByVal lngID As Long)
    Dim strResult As String

    On Error GoTo errHandle
    Select Case lngID
        Case conMenu_Manage_First
            strResult = "��"
        Case conMenu_Manage_Second
            strResult = "��"
    End Select
    With vsList
        gstrSQL = "Zl_Ӱ������_Update(" & .TextMatrix(.Row, GetCN("ҽ��ID")) & ",'" & strResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "Ӱ������")
        .TextMatrix(.Row, GetCN("����")) = strResult
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�޸�()
    With frmPatholRIS
        .mlngModul = mlngModul
        .mlngSendNo = vsList.TextMatrix(vsList.Row, GetCN("���ͺ�"))
        .mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
        .mintEditMode = IIf(vsList.TextMatrix(vsList.Row, GetCN("���״̬")) > 1, 3, 1) '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .InitMvar
        If .RefreshPatiInfor(False) = True Then  'ˢ�²���
            .mblnOK = False
            .zlShowMe Me
        End If
        If .mblnOK Then RefreshList '�ɹ�����
    End With
End Sub
Private Sub Menu_Manage_���ƵǼ�()
    With frmPatholRIS
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .mblnOK = False
        .InitMvar
        If .CopyCheck(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")), vsList.TextMatrix(vsList.Row, GetCN("���ͺ�"))) = True Then   'ˢ�²���
            .zlShowMe Me
        End If
        If .mblnOK Then '�ɹ�����
            If mblnֱ�Ӽ�� Then
                Call StateCheck(2, .mlngAdviceID)
            Else
                Call RefreshList
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_�Ǽ�()
    With frmPatholRIS
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .mblnOK = False
        .InitMvar
        .zlShowMe Me
        If .mblnOK Then '�ɹ�����
            If mblnֱ�Ӽ�� Then
                Call StateCheck(2, .mlngAdviceID)
            Else
                Call RefreshList
            End If
            
            If vsList.Rows = 2 Then
              Call vsList.Select(1, 1)
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_ȡ���Ǽ�()
    On Error GoTo errHandle
    
    If MsgBoxD(Me, "ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) & "," & vsList.TextMatrix(vsList.Row, GetCN("���ͺ�")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����Ǽ�")
    Call RefreshList
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�ٻ�ȡ��()
'���ܣ��ٻر�ȡ���ĵǼ�
    On Error GoTo errH
    
    If MsgBoxD(Me, "ȷʵҪ�ٻر�ȡ���Ǽǵ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) & "," & vsList.TextMatrix(vsList.Row, GetCN("���ͺ�")) & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Menu_Manage_����()
Dim blnFocusFind As Boolean
Dim rsTemp As ADODB.Recordset
    If Me.ActiveControl Is Nothing Then
        blnFocusFind = False
    Else
        blnFocusFind = (Me.ActiveControl.Name = "txtFilter")
    End If
    With frmPatholRIS
        .mstrPrivs = mstrPrivs
        .mlngModul = mlngModul
        .mlngSendNo = vsList.TextMatrix(vsList.Row, GetCN("���ͺ�"))
        .mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
        .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .InitMvar
        If .RefreshPatiInfor(True) = True Then  'ˢ�²���
            .mblnOK = False
            .zlShowMe Me
        End If
        If .mblnOK Then  '�ɹ�����
            Call StateCheck(2)
            If mblnOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
        End If
        If blnFocusFind Then txtFilter.SetFocus '�Զ���λ����λ��
    End With
End Sub
Private Sub Menu_Manage_ȡ������()
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long
    
    On Error GoTo errHandle
    With vsList
        If .TextMatrix(.Row, GetCN("���״̬")) <= 1 Then Call Menu_Manage_ȡ���Ǽ�: Exit Sub '����������
        '------------------------------------��ǩ������Ҫ�Ȼ���ǩ�����ٳ���
        lngAdviceID = .TextMatrix(.Row, GetCN("ҽ��ID"))
        gstrSQL = "Select Distinct B.���ʱ�� From ����ҽ������ A, ���Ӳ�����¼ B Where A.����ID=B.Id And A.ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ƿ�ǩ��", lngAdviceID)
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!���ʱ��, "") <> "" Then 'ǩ������
                MsgBoxD Me, "��ǰ���˵ļ�鱨���Ѿ�ǩ��,����ȡ�����,���Ȼ���ǩ��!", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If

        If MsgBoxD(Me, "ȡ�����μ�齫ɾ����Ӧ�ļ��ͼ��ͼ�鱨�棬�Ƿ������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        If .TextMatrix(.Row, GetCN("���UID")) <> "" And InStr(mstrPrivs, "���ͼ��") <= 0 Then
            MsgBoxD Me, "��û�����ͼ��Ȩ��,�������ͼ��,���в���ȡ��������!", vbInformation, gstrSysName
            Exit Sub
        End If
                
        
        gstrSQL = "ZL_Ӱ����_CANCEL(" & lngAdviceID & "," & .TextMatrix(.Row, GetCN("���ͺ�")) & ",0)"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        'ɾ��Ӱ���ļ���Ŀ¼
        RemoveCheckImages lngAdviceID, .TextMatrix(.Row, GetCN("���ͺ�"))
    End With
    
    Call StateCheck(1)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_����Ӱ��()
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long

    On Error GoTo errHandle
    With vsList
        lngAdviceID = .TextMatrix(.Row, GetCN("ҽ��ID"))
        lngSendNO = .TextMatrix(.Row, GetCN("���ͺ�"))
        
        Call funRelateSeries(Me, lngAdviceID, lngSendNO, True, mblnMoved, dcmRelateViewer)
        '����Ӱ����״̬�����ԭ����״̬���ѱ��������޸ĳ��Ѽ�飬
        If .TextMatrix(.Row, GetCN("���״̬")) < 3 Then
            '��������Ѿ���ͼ�����޸ĳ��Ѽ��
            gstrSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ͼ��", lngAdviceID)
            
            If Not IsNull(rsTemp!���uid) Then
                gstrSQL = "Zl_Ӱ����_State(" & lngAdviceID & "," & lngSendNO & ",3)"
                zlDatabase.ExecuteProcedure gstrSQL, "����Ӱ��"
            End If
        End If
    End With
    Call RefreshList
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_View_Locate_Type_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    mstrLocateType = Split(control.Caption, "(")(0)
    cbrMain.RecalcLayout
    If mstrLocateType = "�ɣÿ�" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Else
            txtLocate.Text = mobjICCard.Read_Card(Me)
        End If
    End If
    txtLocate.SetFocus
End Sub

Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    If mlngCur����ID <> control.DescriptionText Then
        mlngCur����ID = control.DescriptionText
        mstrCur���� = Split(control.Caption, "(")(0)
        
        Call ReadStudyListColor(mlngCur����ID)
        Call cbrMain.RecalcLayout
        Call InitMvar(False)
        
        If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Call frmVideoCapture.InitDeptPara(mlngCur����ID)
        
        Call mfrmPacsReport.InitReportWindow(mlngCur����ID, mlngModul, mstrPrivs, False)
        
'        If Not frmPACSFilter Is Nothing Then
'            frmPACSFilter.mBeforeDays = mBeforeDays
'            frmPACSFilter.dtpBegin.value = SQLCondition.��ʼʱ��
'        End If
        
        mblnInitOk = False '��ֹ���Ӵ�����ع����ж��Ӵ������ˢ��
        Call InitSubForm
        mblnInitOk = True

        
        
        Call RefreshList
    End If
End Sub

Private Sub Menu_Help_Web_Home_click()
    zlHomePage hWnd
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub
Private Sub cboTimes_Click()
    If cboTimes.ListCount <= 1 Then Exit Sub
    If cboTimes.Tag = "" Then Exit Sub '��ʱcbotime��Ŀδ������ɣ���listindex��ֵ����
    
    On Error GoTo errHandle
    Dim lngAdviceID As Long
    lngAdviceID = cboTimes.ItemData(cboTimes.ListIndex)
    If lngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) Then Call vsList_RowColChange: Exit Sub '�����뵱ǰѡ��ҽ��ID��ͬʱ���ɱ���������

    mblnIsHistory = True: mlngHOrderID = lngAdviceID '�����������̵������Ⱥ�˳�������
    Call FillTxtInfor(mlngHOrderID)  '������Ϸ����˻�����Ϣ
    Call FillTxtAppend(mlngHOrderID) '������½�ҽ������
    Call ShowTab(mlngHOrderID)  '���ݲ����ṩ��ͬѡ�
    Call RefreshTabWindow(mlngHOrderID) 'ˢ���Ӵ���

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.hWnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim strTemp As String
    
    Select Case control.ID
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_סԺ
            mblncmdסԺ = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_�ѽ�
            mblncmd�ѽ� = Not control.Checked
            If mblncmd�ѽ� Then mblncmdδ�� = False
        Case ID_δ��
            mblncmdδ�� = Not control.Checked
            If mblncmdδ�� Then mblncmd�ѽ� = False
'        Case ID_Ӱ����� + 1 To ID_Ӱ����� + 40
'            control.Checked = Not control.Checked
'            mblncmdӰ�����(control.ID - ID_Ӱ����� - 1) = control.Checked
'            If control.Checked = True Then
'                mintcmdӰ����� = mintcmdӰ����� + 1
'            Else
'                mintcmdӰ����� = mintcmdӰ����� - 1
'            End If
'            Set objControl = cbrdock.FindControl(, ID_Ӱ�����)
'            If mintcmdӰ����� = 0 Then
'                strTemp = "Ӱ�����"
'            Else
'                strTemp = ""
'                For i = 1 To objControl.CommandBar.Controls.Count
'                    If objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Checked = True Then
'                        strTemp = IIf(strTemp = "", objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, strTemp & "," & objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption)
'                    End If
'                Next i
'            End If
'            objControl.Caption = strTemp
        Case ID_�Ǽ�
            mblncmd�Ǽ� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_����סԺ
            control.Checked = Not control.Checked
            mblncmd���� = Not mblncmd����
        Case ID_�������_����
            mblncmd���� = Not control.Checked
        Case ID_�������_����
            mblncmd���� = Not control.Checked
        Case ID_�������_ϸ��
            mblncmdϸ�� = Not control.Checked
        Case ID_�������_ʬ��
            mblncmdʬ�� = Not control.Checked
        Case ID_�������_����
            mblncmd���� = Not control.Checked
        Case ID_�걾����_����
            mblncmd���� = Not control.Checked
        Case ID_�걾����_С�걾
            mblncmdС�걾 = Not control.Checked
        Case ID_�걾����_����
            mblncmd���� = Not control.Checked
        Case ID_�걾����_����
            mblncmd���� = Not control.Checked
        Case ID_�걾����_Һ��
            mblncmdҺ�� = Not control.Checked
        Case ID_���ҷ�ʽ * 100# To ID_���ҷ�ʽ * 100# + 8
            mstrCurFindtype = Split(control.Caption, "(")(0)
            If InStr(mstrCurFindtype, "�ɣÿ�") > 0 Then
                If mobjICCard Is Nothing Then
                    Set mobjICCard = CreateObject("zlICCard.clsICCard")
                End If
                txtFilter.Text = mobjICCard.Read_Card(Me)
            End If
            
            If txtFilter.PasswordChar = "*" Then '֮ǰ�Ǿ��￨�ţ���Ҫ������������
                txtFilter.Text = "": txtFilter.PasswordChar = ""
            End If
            
            txtFilter_GotFocus
            cbrdock.RecalcLayout
            Exit Sub
        Case ID_��ʼ����
            Call subRefreshFilterCondition(txtFilter.Text)
    End Select
cbrdock.RecalcLayout
Call RefreshList
End Sub



Private Function GetPatholNum(ByVal strSureNum As String) As String
'�ֽ�ȷ�Ϻ���
    Dim lngFindSplitChar As Long
    
    lngFindSplitChar = InStr(1, strSureNum, "-")
    
    If lngFindSplitChar > 0 Then
        GetPatholNum = Mid(strSureNum, 1, lngFindSplitChar - 1)
    Else
        GetPatholNum = strSureNum
    End If
    
End Function



Private Sub subRefreshFilterCondition(strFilter As String)
'------------------------------------------------
'���ܣ���txtFilter�ؼ������ݸ��¹�������
'������ strFilter --- ��������
'���أ���
'------------------------------------------------

    On Error GoTo err
    
    With SQLCondition
        .���� = ""
        .���￨ = ""
        .����� = 0
        .סԺ�� = 0
        .���ݺ� = ""
        .���� = 0
        .���֤ = ""
        .IC�� = ""
        .����� = ""
        Select Case mstrCurFindtype
            Case "��  ��"
                .���� = Trim(strFilter)
            Case "���￨"
                .���￨ = Trim(strFilter)
            Case "�����"   '��ݷ�ʽ�ǡ�*+���֡�,VAL��ȡǰ����*��Ҫ���⴦��
                If Left(strFilter, 1) = "*" Then
                    strFilter = Mid(strFilter, 2)
                End If
                .����� = Val(strFilter)
            Case "סԺ��"   '��ݷ�ʽ�ǡ�++���֡�
                .סԺ�� = Val(strFilter)
            Case "���ݺ�"
                .���ݺ� = Trim(strFilter)
            Case "����"
                .���� = Val(strFilter)
            Case "���֤"
                .���֤ = Trim(strFilter)
            Case "�ɣÿ�"
                .IC�� = Trim(strFilter)
            Case "�����"
                .����� = GetPatholNum(Trim(strFilter))
        End Select
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrdock_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    If CommandBar.Parent Is Nothing Then Exit Sub
    If CommandBar.Parent.ID = ID_���ҷ�ʽ Then
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 0, "�����(&1)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 1, "סԺ��(&2)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 2, "���￨(&3)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 3, "��  ��(&4)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 4, "���ݺ�(&5)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 5, "����(&6)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 6, "���֤(&7)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 7, "�ɣÿ�(&8)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 8, "�����(&9)")
            End If
        End With
    End If
End Sub
Private Sub cbrdock_Resize()
Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    tabFilter.Top = lngTop
    tabFilter.Left = lngLeft
    tabFilter.Width = picList.Width
    
    vsList.Top = lngTop + IIf(tabFilter.Visible, tabFilter.Height, 0) + 7
    vsList.Left = lngLeft
    vsList.Width = picList.Width
    vsList.Height = picList.Height - lngTop - txtAppend.Height - 100

    txtAppend.Top = vsList.Top + vsList.Height + 100
    txtAppend.Left = lngLeft + 100
    txtAppend.Width = picList.Width - 200
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_סԺ
            control.Checked = mblncmdסԺ
            control.IconId = IIf(mblncmdסԺ, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd�ѽ� Xor mblncmdδ��
            control.Caption = IIf(mblncmd�ѽ� Xor mblncmdδ��, IIf(mblncmd�ѽ�, " �ѽɷ�", " δ�ɷ�"), " ��  ��")
        Case ID_�ѽ�
            control.Checked = mblncmd�ѽ�
            control.IconId = IIf(mblncmd�ѽ�, 90001, 90000)
        Case ID_δ��
            control.Checked = mblncmdδ��
            control.IconId = IIf(mblncmdδ��, 90001, 90000)
'        Case ID_Ӱ�����
'            control.IconId = IIf(mintcmdӰ����� = 0, 90000, 90001)
'        Case ID_Ӱ����� + 1 To ID_Ӱ����� + 40
'            control.Checked = mblncmdӰ�����(control.ID - ID_Ӱ����� - 1)
'            control.IconId = IIf(control.Checked, 90001, 90000)
        Case ID_�Ǽ�
            control.Checked = mblncmd�Ǽ�
            control.IconId = IIf(mblncmd�Ǽ�, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_�������_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�������_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�������_ϸ��
            control.Checked = mblncmdϸ��
            control.IconId = IIf(mblncmdϸ��, 90001, 90000)
        Case ID_�������_ʬ��
            control.Checked = mblncmdʬ��
            control.IconId = IIf(mblncmdʬ��, 90001, 90000)
        Case ID_�������_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�걾����
            control.Checked = mblncmd���� Or mblncmdС�걾 Or mblncmd���� Or mblncmd���� Or mblncmdҺ��
            control.IconId = IIf(control.Checked, 90001, 90000)
            control.Caption = "�걾����(" & IIf(mblncmd����, "����,", "") & IIf(mblncmdС�걾, "С�걾,", "") & IIf(mblncmd����, "����,", "") & IIf(mblncmd����, "����,", "") & IIf(mblncmdҺ��, "Һ��,", "") & ")"
            control.Caption = Replace(control.Caption, "()", "")
            control.Caption = Replace(control.Caption, ",)", ")")
        Case ID_�걾����_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�걾����_С�걾
            control.Checked = mblncmdС�걾
            control.IconId = IIf(mblncmdС�걾, 90001, 90000)
        Case ID_�걾����_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�걾����_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�걾����_Һ��
            control.Checked = mblncmdҺ��
            control.IconId = IIf(mblncmdҺ��, 90001, 90000)
        Case ID_����סԺ
            control.IconId = IIf(control.Checked, 90001, 90000)
        Case ID_���ҷ�ʽ
            control.Caption = mstrCurFindtype
        Case ID_���ҷ�ʽ * 100# To ID_���ҷ�ʽ * 100# + 7
            control.Checked = (InStr(control.Caption, mstrCurFindtype) > 0)
    End Select
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub


'����ִ��
Private Sub ExecuteStudyMoney()
    On Error GoTo errHandle
      
    Dim lngAdviceID As Long, lngSendNO As Long
    
    With vsList
        lngAdviceID = Nvl(.TextMatrix(.Row, GetCN("ҽ��ID")), 0)
        lngSendNO = Nvl(.TextMatrix(.Row, GetCN("���ͺ�")), 0)
    End With
    
    gstrSQL = "Zl_Ӱ�����ִ��(" & lngAdviceID & "," & lngSendNO & ",2)"
    zlDatabase.ExecuteProcedure gstrSQL, "����ִ��"
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    
    If control.ID <> 0 Then
        If cbrMain.FindControl(, control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    cbrMain.RecalcLayout
    Select Case control.ID
    
'--------------------------�ļ�------------------
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
            
        Case conMenu_File_Excel '�嵥��ӡ
            Call Menu_File_Excel_click
            
        Case conMenu_File_BatPrint '������ӡ
            Call Menu_File_BatPrint
            
        Case conMenu_File_Parameter '��������
            Call Menu_File_Parmeter_click
            
        Case conMenu_File_SendImg '����ͼ��
            frmPacsSendImage.ShowMe Me
            
        Case conMenu_Manage_Change_In   '�����б�
            If dkpMain.Panes(1).Hidden = False Then
                dkpMain.Panes(1).Hide
            Else
                dkpMain.ShowPane (1)
            End If
            
        Case conMenu_File_Exit '�˳�
            Unload Me
            
'---------------------------���-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '��ӡ���Ƶ���
            Call FuncBillPrint(control)
            
        Case conMenu_Manage_Regist                          '�Ǽ�
            Call Menu_Manage_�Ǽ�
            
        Case conMenu_Manage_CopyCheck                       '���ƵǼ�
            Call Menu_Manage_���ƵǼ�
            
        Case conMenu_Manage_Receive                         '����
            Call Menu_Manage_����
            
        Case conMenu_Manage_Redo                            'ȡ���Ǽ�
            Call Menu_Manage_ȡ���Ǽ�
            
        Case conMenu_Manage_ReGet                           '�ٻ�ȡ��
            Call Menu_Manage_�ٻ�ȡ��
        
        Case conMenu_Manage_ThingModi                       '�޸ĵǼ�
            Call Menu_Manage_�޸�
            
        Case conMenu_Manage_Logout                          'ȡ������
            Call Menu_Manage_ȡ������
            
        Case conMenu_Manage_Transfer                        '����Ӱ��
            Call Menu_Manage_����Ӱ��
            
        Case conMenu_Manage_Cancel                          'ȡ������
            Call Menu_Manage_ȡ������
            
        Case conMenu_Manage_Review                          '��ע
            Call Menu_Manage_���
            
        Case conMenu_Manage_ReportRelease                   '���淢��
            Call Menu_Manage_���淢��
            
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '���������
            Call Menu_Manage_�������(control.ID)
        
        Case conMenu_Manage_First, conMenu_Manage_Second
            Call Menu_Manage_Ӱ������(control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_��ɫͨ��(control.ID)
            
        Case conMenu_Manage_ClearUp                           '�ޱ������
            Call Menu_Edit_�ޱ������
                    
        Case conMenu_Manage_Finish                          '�ޱ���ֱ�����
            Call Menu_Manage_�ޱ������
            
        Case conMenu_Manage_Complete                        '������
            Call Menu_Manage_����������
        
        Case conMenu_Manage_Undone                          'ȡ��������
            Call Menu_Manage_ȡ��������
            
        Case conMenu_Manage_RelatingPatiet                  '��������
            Call Menu_Manage_��������
        Case conMenu_File_Preview, conMenu_File_Print       '����Ԥ���ʹ�ӡ
            Dim i As Integer
            'û���治�ܴ�ӡ��Ԥ��
            If vsList.TextMatrix(vsList.Row, GetCN("������")) = "" Then
                MsgBoxD Me, "��ǰ����û�м�鱨�棬���ܲ��������飡", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '����ҳ�治�ɼ�ʱ��ִ���κβ���
            If TabWindow.Selected.Tag <> "������д" Then
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
                Next
            End If
            If TabWindow.Selected.Tag = "������д" Then
                If mblnPacsReport = True Then
                    mfrmPacsReport.zlExecuteCommandBars control
                Else
                    mobjReport.zlExecuteCommandBars control
                End If
            End If
'-------------------------�������---------------------
        Case conMenu_Antibody_Manage    '�������
            Call Menu_Manage_�������
            
        Case conMenu_Meal_Manage        '�ײ�ά��
            Call Menu_Manage_�ײ�ά��
            
        Case conMenu_Pathol_Request     '��������
            Call Menu_Manage_��������
            
        Case conMenu_Report_Delay       '�ӳٵǼ�
            Call Menu_Manage_�ӳٵǼ�
        
        Case conMenu_Con_Request, conMenu_Con_Feedback       '�������뷴��
            Call Menu_Manage_�������뷴��(control.ID)
            
        Case conMenu_Decalin_Task       '�Ѹ�����
            Call Menu_Manage_�Ѹ��������

'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(control)
        Case conMenu_Manage_LocateType * 10# To conMenu_Manage_LocateType * 10# + 6 '��λ
            Call Menu_View_Locate_Type_click(control)
        Case conMenu_View_Filter '����
            Call Menu_View_Filter_click
        Case conMenu_View_Refresh 'ˢ��
            Call RefreshList
'--------------------------�����ɼ�-----------------
        Case comMenu_Cap_Process    '�����ɼ�
            control.Checked = Not control.Checked
            Call Menu_Manage_�����ɼ�(True)
            
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|"))
            Call Menu_Dept_Select(control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If control.parameter <> "" Then 'ִ�з�������ǰģ��ı���
                With vsList
                    If .TextMatrix(.Row, GetCN("ҽ��ID")) <> "" Then
                        Call ReportOpen(gcnOracle, Split(control.parameter, ",")(0), Split(control.parameter, ",")(1), Me, _
                            "NO=" & .TextMatrix(.Row, GetCN("NO")), "����=" & .TextMatrix(.Row, GetCN("��¼����")), "ҽ��id=" & .TextMatrix(.Row, GetCN("ҽ��ID")), 1)

                    Else
                        Call ReportOpen(gcnOracle, Split(control.parameter, ",")(0), Split(control.parameter, ",")(1), Me, "", 1)
                    End If
                End With
            End If
        Case Else
            If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then Exit Sub
            Select Case TabWindow.Selected.Tag
                Case "������д"
                    '���汻ĳ�˴򿪺��ٱ��������˱༭���޶�
                    If control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Modify Or control.ID = conMenu_PacsReport_Open Or control.ID = conMenu_Edit_Delete Then
                        If CheckConcurrentReport(Me, vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = False Then Exit Sub
                    End If
                    
                    '���� ֻ����д�Լ����ı���,'��������д���޶���ɾ��
                    If mblnTechReptSame = True _
                        And (control.ID = conMenu_Edit_Modify Or control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Delete) _
                        And Nvl(vsList.TextMatrix(vsList.Row, GetCN("��鼼ʦ"))) <> "" _
                And Nvl(vsList.TextMatrix(vsList.Row, GetCN("��鼼ʦ"))) <> UserInfo.���� Then
                        MsgBoxD Me, "�㲻��������ߵļ�鼼ʦ���޷�������ݱ��档", vbInformation, gstrSysName
                    Else
                        If mblnPacsReport = True Then
                            If control.ID = conMenu_PacsReport_Open Then   '�򿪱��洰��
                                Call Menu_Manage_PACS����
                            Else
                                mfrmPacsReport.zlExecuteCommandBars control
                            End If
                        Else
                            mobjReport.zlExecuteCommandBars control
                        End If
                    End If
                Case "�������"
                    mobjExpense.zlExecuteCommandBars control
                    
                    '----------------------����ʱ��ִ�з���------------------
                    If control.ID = conMenu_Edit_Append _
                    Or control.ID = conMenu_Edit_Modify _
                    Or control.ID = conMenu_Edit_NewItem * 10# + 1 _
                    Or control.ID = conMenu_Edit_NewItem * 10# + 2 _
                    Or control.ID = conMenu_Edit_NewItem * 10# + 3 Then
            
                        If vsList.TextMatrix(vsList.Row, GetCN("���״̬")) >= 2 Then
                            Call ExecuteStudyMoney
                        End If
                    End If
                    
                Case "סԺҽ��"
                    mobjInAdvice.zlExecuteCommandBars control
                Case "����ҽ��"
                    mobjOutAdvice.zlExecuteCommandBars control
                Case "סԺ����"
                    mobjInEPRs.zlExecuteCommandBars control
                Case "���ﲡ��"
                    mobjOutEPRs.zlExecuteCommandBars control
                Case "�Ŷӽк�"
                    If Not mobjQueue Is Nothing Then
                        mobjQueue.zlExecuteCommandBars control
                    End If
            End Select
    End Select
End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo errHandle
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays - 1
        .mDept = mlngCur����ID '��ǰ����
        .Show 1, Me
        If Not .mblnOK Then Exit Sub 'û�з�������
        
        '��ʹ��ʱ������ʱ����չ̶�����
        txtFilter.Text = ""
        SQLCondition.���� = ""
        SQLCondition.���￨ = ""
        SQLCondition.����� = 0
        SQLCondition.סԺ�� = 0
        SQLCondition.���ݺ� = ""
        SQLCondition.���� = 0
        SQLCondition.���֤ = ""
        SQLCondition.IC�� = ""
        
        SQLCondition.��ʼʱ�� = Format(.dtpBegin.value, "yyyy-MM-dd HH:mm:00")
        If Format(.dtpEnd.value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            SQLCondition.����ʱ�� = CDate(0) '��ʾȡ��ǰʱ��
        Else
            SQLCondition.����ʱ�� = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
        End If
        
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
        
        If .optFindType(1).value = True Then 'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
            SQLCondition.ʱ������ = 1
        ElseIf .optFindType(2).value = True Then
            SQLCondition.ʱ������ = 2
        Else
            SQLCondition.ʱ������ = 3
        End If
        
        If .cboPart.ListIndex <> 0 Then '���걾��λ
            SQLCondition.�걾��λ = NeedName(.cboPart.Text)
        Else
            SQLCondition.�걾��λ = ""
        End If
        
        '�����Ա�
        If NeedName(.cboSex.Text) = "ȫ��" Then
            SQLCondition.�Ա� = ""
        Else
            SQLCondition.�Ա� = NeedName(.cboSex.Text)
        End If
        
        '��������
        Select Case NeedName(.cboAgeType.Text)
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 365
                SQLCondition.�������� = Val(.txtEndAge.Text) * 365
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 30
                SQLCondition.�������� = Val(.txtEndAge.Text) * 30
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 7
                SQLCondition.�������� = Val(.txtEndAge.Text) * 7
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 1
                SQLCondition.�������� = Val(.txtEndAge.Text) * 1
        End Select
        
        If Trim(.txtBeginAge.Text) = "" Then SQLCondition.��ʼ���� = -1
        If Trim(.txtEndAge.Text) = "" Then SQLCondition.�������� = -1
        
        SQLCondition.�������� = Trim(.cboAgeWhere.Text)
        
        If .cboDept.ListIndex <> 0 Then '���˿���
            SQLCondition.���˿��� = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            SQLCondition.���˿��� = 0
        End If

        If .cbodiagdoc.ListIndex <> 0 Then '���ҽ��
            SQLCondition.���ҽ�� = NeedName(.cbodiagdoc.Text)
        Else
            SQLCondition.���ҽ�� = ""
        End If
        
        If .cboAuditing.ListIndex <> 0 Then '���ҽ��
            SQLCondition.���ҽ�� = NeedName(.cboAuditing.Text)
        Else
            SQLCondition.���ҽ�� = ""
        End If
        
'        If .cboCheckStep.ListIndex <> 0 Then '������
'            SQLCondition.������ = .cboCheckStep.Text
'        Else
'            SQLCondition.������ = ""
'        End If
        
'        If .cboModality.ListIndex <> 0 Then 'Ӱ�����
'            SQLCondition.Ӱ����� = Split(.cboModality.Text, "--")(1)
'        Else
'            SQLCondition.Ӱ����� = ""
'        End If
        
        If Trim(.TxtӰ�����) <> "" Then 'Ӱ�����
            SQLCondition.������� = Trim(.TxtӰ�����)
        Else
            SQLCondition.������� = ""
        End If
        
        If Trim(.txt��������) <> "" Then '��������
            SQLCondition.�������� = Trim(.txt��������)
        Else
            SQLCondition.�������� = ""
        End If
        
        If NeedName(.cboYinYangXing.Text) = "����" Then
            SQLCondition.������� = 1
        ElseIf NeedName(.cboYinYangXing.Text) = "����" Then
            SQLCondition.������� = 0
        Else
            SQLCondition.������� = -1
        End If
        
        If .cbo����.ListIndex = 0 Then
            SQLCondition.Ӱ������ = ""
        Else
            SQLCondition.Ӱ������ = NeedName(.cbo����.Text)
        End If
        
        If .cbo��鼼ʦ.ListIndex = 0 Then
            SQLCondition.��鼼ʦ = ""
        Else
            SQLCondition.��鼼ʦ = NeedName(.cbo��鼼ʦ.Text)
        End If
        
        
        If Trim(.txtPacsRpt(0)) <> "" Then 'PACS�������
            SQLCondition.������� = Trim(.txtPacsRpt(0))
        Else
            SQLCondition.������� = ""
        End If
        
        If Trim(.txtPacsRpt(1)) <> "" Then
            SQLCondition.������ = Trim(.txtPacsRpt(1))
        Else
            SQLCondition.������ = ""
        End If
        
        If Trim(.txtPacsRpt(2)) <> "" Then
            SQLCondition.���� = Trim(.txtPacsRpt(2))
        Else
            SQLCondition.���� = ""
        End If
        
        If Trim(.txt���.Text) <> "" Then
            SQLCondition.��� = Trim(.txt���.Text)
        Else
            SQLCondition.��� = ""
        End If
        
        Call RefreshList '����ˢ��
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_Manage_LocateType
            With CommandBar.Controls
                If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10#, "��ʶ��(&1)"): objControl.Category = "Main": objControl.Checked = True
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 1, "���￨(&2)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 2, "����(&3)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 3, "���ݺ�(&4)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 4, "����(&5)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 5, "���֤(&6)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 6, "�ɣÿ�(&7)"): objControl.Category = "Main"
                End If
            End With
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    For i = 0 To UBound(Split(mstrCanUse����, "|")) 'mstrCanUse����=id_����-����|id_����-����
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i, Split(Split(mstrCanUse����, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse����, "|")(i), "_")(0)
                        If mlngCur����ID = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.Tag
                Case "סԺҽ��"
                    mobjInAdvice.zlPopupCommandBars CommandBar
                Case "����ҽ��" '����
                    mobjOutAdvice.zlPopupCommandBars CommandBar
                Case "�������"
                    mobjExpense.zlPopupCommandBars CommandBar
            End Select
    End Select
End Sub
Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim blnNoRecord As Boolean, intState As Integer, blnCancel As Boolean
    If Not mblnInitOk Then Exit Sub
    
    blnNoRecord = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0
    control.Style = xtpButtonIconAndCaption
    
    If Not blnNoRecord Then
        intState = Val(vsList.TextMatrix(vsList.Row, GetCN("���״̬")))
        blnCancel = vsList.TextMatrix(vsList.Row, GetCN("������")) = "�Ѿܾ�"
    End If
    
    Select Case control.ID
        Case conMenu_Manage_LocateType
            control.Caption = "��" & mstrLocateType & "��λ(&G)"
            control.Enabled = Not blnNoRecord
        Case conMenu_Manage_LocateType * 10# To conMenu_Manage_LocateType * 10# + 6
            control.Checked = (InStr(control.Caption, mstrLocateType) > 0)
        Case conMenu_Manage_LocateValue
            control.Enabled = Not blnNoRecord
        Case comMenu_Cap_Process
            control.Enabled = Not blnNoRecord
            
            If Not CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then
                control.Visible = False
            End If
            
        Case conMenu_View_Filter * 10#
            control.Caption = "��ǰ����:" & mstrCur����
            
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|"))
            control.Checked = (control.DescriptionText = mlngCur����ID)
            
        Case conMenu_View_ToolBar_Button '������
            If cbrMain.Count >= 2 Then
                control.Checked = Me.cbrMain(2).Visible
            End If
            
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbrMain.Count >= 2 Then
                control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
            
        Case conMenu_View_ToolBar_Size '��ͼ��
            control.Checked = Me.cbrMain.Options.LargeIcons
            
        Case conMenu_View_StatusBar '״̬��
            control.Checked = Me.stbThis.Visible
            
        Case conMenu_View_Filter   '����
        
        Case conMenu_View_Refresh  'ˢ��
        
        Case conMenu_Manage_RequestPrint
            control.Enabled = control.CommandBar.Controls.Count > 0 And Not blnNoRecord
                
        Case conMenu_Manage_Regist   '���Ǽ�(&I)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            End If
        Case conMenu_Manage_CopyCheck '�ٴεǼ�
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Redo   'ȡ���Ǽ�(&R)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '�ٻ�ȡ��
            If Not blnNoRecord Then
                control.Enabled = blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ThingModi   '�޸���Ϣ(&M)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 3 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '��鱨��(&L)
            If InStr(mstrPrivs, "��鱨��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   'ȡ������(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent.Type = xtpControlPopup Then
                If InStr(mstrPrivs, "ȡ������") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "ȡ������"
                    control.Caption = "ȡ������(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' �������е���ȡ��������ȡ���Ǽ�,ͬһ�������ȡ���ǼǺ�ȡ����鹦��
                control.Visible = IIf(intState <= 1, InStr(mstrPrivs, "���Ǽ�") > 0, InStr(mstrPrivs, "ȡ������") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And Not blnCancel) '���ܾ��Ĳ��ܱ��ٴξܾ�
                control.ToolTipText = IIf(intState <= 1, "ȡ���Ǽ�", "ȡ������")
                control.Caption = "ȡ��"
            End If
        Case conMenu_Manage_Transfer   '����Ӱ��(&C)
            If InStr(mstrPrivs, "���ͼ��") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_Cancel   'ȡ������(&B)
            If InStr(mstrPrivs, "���ͼ��") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("���UID")) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_First, conMenu_Manage_Second, conMenu_Manage_Quality
            If InStr(mstrPrivs, "Ӱ���ʿ�") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("���UID")) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Review  '��ע
            If InStr(mstrPrivs, "���") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord And intState > 1 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReportRelease       '���淢��,��������ɺ󶼿���ִ��
            If intState >= 2 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
            
            '�޸ı��淢�Ű�ť�ı���
            If Not blnNoRecord Then
                If vsList.TextMatrix(vsList.Row, GetCN("���淢��")) = "�ѷ���" Then
                    control.Caption = "�ջ�"
                    control.ToolTipText = "�ջ��Ѿ����ŵı���"
                Else
                    control.Caption = "����"
                    control.ToolTipText = "���ű���"
                End If
            End If
        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive   '���������(&X)
            If (InStr(GetInsidePrivs(p���Ʊ������), "������д") <= 0 And InStr(GetInsidePrivs(p���Ʊ������), "�����޶�") <= 0) Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '��ɫͨ�����/ȡ��
            If InStr(mstrPrivs, "��ɫͨ��") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_Finish   '�ޱ������(&F)
            If InStr(mstrPrivs, "�ޱ������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 2 Or intState = 3
            End If
        Case conMenu_Manage_ClearUp   '�ޱ������(&U)
            If InStr(mstrPrivs, "�ޱ������") <= 0 Then
                control.Visible = False
            ElseIf intState = 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("������")) = ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '������(&E)
            If InStr(mstrPrivs, "������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState = 4 Or intState = 5)
            End If
        Case conMenu_Manage_Undone   'ȡ�����(&U)
            If InStr(mstrPrivs, "ȡ��������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 6
            End If
        Case conMenu_Manage_RelatingPatiet  '��������
            If InStr(mstrPrivs, "��������") <= 0 Or mblnRelatingPatient = False Then
                control.Visible = False
            ElseIf blnNoRecord Or intState < 2 Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
            
        '---------------------------------���������-------------------------------------
        Case conMenu_Antibody_Manage
            If Not (CheckPopedom(mstrPrivs, "�������") <= 0 Or CheckPopedom(mstrPrivs, "���巴��")) Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Meal_Manage
            If Not CheckPopedom(mstrPrivs, "�ײ�ά��") Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_Request
            If Not (CheckPopedom(mstrPrivs, "�ؼ�����") Or CheckPopedom(mstrPrivs, "��Ƭ����") Or CheckPopedom(mstrPrivs, "��ȡ����")) Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Report_Delay
            If Not CheckPopedom(mstrPrivs, "�����ӳ�") Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Con_Request
            If Not CheckPopedom(mstrPrivs, "��������") Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Con_Feedback
            If Not CheckPopedom(mstrPrivs, "���ﷴ��") Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Decalin_Task
            If Not CheckPopedom(mstrPrivs, "����ȡ��") Then
                control.Enabled = False
            Else
            
            End If
        
        Case conMenu_File_SendImg
            If InStr(mstrPrivs, "�ļ�����") <= 0 Then control.Visible = False
        Case conMenu_File_PrintSet     '��ӡ����(&S)
        Case conMenu_File_Preview, conMenu_File_Print '����Ԥ��(&V) �����ӡ(&P)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Excel         '�嵥��ӡ(&L)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_BatPrint    ' ������ӡ(&B)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter     '��������(&O)
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup, conMenu_PatholManage
        Case conMenu_Help_Help, conMenu_Help_About  '����
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '����WEB
        Case conMenu_File_Exit
        Case conMenu_View_ToolBar
        Case conMenu_Manage_Change_In   '�����б�
        Case Else
            If blnNoRecord Then control.Enabled = False: Exit Sub
            Select Case TabWindow.Selected.Tag
                Case "������д"
                    If mblnPacsReport = True Then
                        mfrmPacsReport.zlUpdateCommandBars control
                    Else
                        mobjReport.zlUpdateCommandBars control
                    End If
                Case "�������"
                    mobjExpense.zlUpdateCommandBars control
                Case "סԺҽ��"
                    mobjInAdvice.zlUpdateCommandBars control
                Case "����ҽ��"
                    mobjOutAdvice.zlUpdateCommandBars control
                Case "סԺ����"
                    mobjInEPRs.zlUpdateCommandBars control
                Case "���ﲡ��"
                    mobjOutEPRs.zlUpdateCommandBars control
            End Select

            If Not blnNoRecord Then
                'ɾ��ֻ�����ѱ���ͽ����п���
                If control.ID = conMenu_Edit_Delete And Val(vsList.TextMatrix(vsList.Row, GetCN("���״̬"))) >= 4 Then
                    control.Enabled = False
                End If
                '��ǰ�鿴�������μ�¼��˵���������
                If cboTimes.ListIndex <> -1 Then
                    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) <> cboTimes.ItemData(cboTimes.ListIndex) Then
                        If control.ID = conMenu_Edit_Copy Or control.ID = conMenu_File_ExportToXML Or control.ID = conMenu_Tool_Search Then
                            '�⼸���˵�������
                        Else
                            control.Enabled = False
                        End If
                    End If
                End If
                '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
                If Val(vsList.TextMatrix(vsList.Row, GetCN("���״̬"))) = 6 Then
                    Select Case control.ID
                        Case conMenu_Edit_MarkMap, conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                            control.Enabled = True
                        Case conMenu_Edit_Copy, conMenu_File_ExportToXML, conMenu_Tool_Search, conMenu_File_Open, conMenu_EditPopup
                            '�⼸���˵�������
                        Case Else
                            control.Enabled = False
                    End Select
                End If
            End If
    End Select
End Sub

Private Sub chkSource_Click(Index As Integer)
    If Not mblnInitOk Then Exit Sub
    Call RefreshList
End Sub

Private Sub InitMvar(Optional blnIsUpdateSearchTime As Boolean = True)
'����:��ʼ��ģ�鼶����,���������ʱ����һ��

    On Error GoTo err
    
    mblnIgnoreResult = GetDeptPara(mlngCur����ID, "���Խ��������", 0) = "1" '        '���Խ��������
    mblnFinishCommit = GetDeptPara(mlngCur����ID, "�ޱ�����ɺ�ֱ�����", 0) = "1" '  '�ޱ�����ɺ�ֱ�����
    mblnReportWithImage = GetDeptPara(mlngCur����ID, "��ͼ�����д����", 0) = "1" '   '��ͼ�����д����
    mblnReportWithResult = GetDeptPara(mlngCur����ID, "��Ӱ�����Ϊ����", 0) = "1" '  '��Ӱ�����Ϊ����
    mblnLocalizerBackward = GetDeptPara(mlngCur����ID, "��λƬ����", 0) = "1" '       '��λƬ����
    mblnCompleteCommit = GetDeptPara(mlngCur����ID, "��˺�ֱ�����", 0) = "1" '      '��˺�ֱ�����
    mBeforeDays = Val(GetDeptPara(mlngCur����ID, "Ĭ�Ϲ�������", 2)) '                   'Ĭ�Ϲ�������
    If mBeforeDays > 15 Or mBeforeDays <= 0 Then
        mBeforeDays = 2
    End If
    mblnTechReptSame = GetDeptPara(mlngCur����ID, "ֻ����д�Լ����ı���", 0) = "1"  'ֻ����д�Լ����ı���
    mblnWriteCapDoctor = GetDeptPara(mlngCur����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", 0) = "1"  '�ɼ�ͼ����Ϊ��鼼ʦ
    mblnPacsReport = GetDeptPara(mlngCur����ID, "����༭��", 0) = "1" '              '����༭��
    mintResultInput = Val(GetDeptPara(mlngCur����ID, "��ʾ������", 1))    '              '��ʾ������
    mblnPrintCommit = GetDeptPara(mlngCur����ID, "��ӡ��ֱ�����", 0) = "1" '         '��ӡ��ֱ�����
    mblnCanPrint = GetDeptPara(mlngCur����ID, "ƽ������˲��ܴ򱨸�") = "1"           'ƽ����Ҫ��˲��ܴ�ӡ =true
    mblnPacsReportShowVideoCapture = GetDeptPara(mlngCur����ID, "��ʾ��Ƶ�ɼ�", 0) = "1" '��ʾ��Ƶ�ɼ�
    mblnRelatingPatient = GetDeptPara(mlngCur����ID, "������������", 0) = "1"       '�Ƿ�ʹ�ù�������
    mlngRefreshInterval = Val(GetDeptPara(mlngCur����ID, "�Զ�ˢ�¼��", 0)) '      '�Զ�ˢ�¼��,Ĭ�ϲ��Զ�ˢ��
    If mlngRefreshInterval > 0 Then
        If mlngRefreshInterval > 65 Then mlngRefreshInterval = 65
        TimerRefresh.Interval = mlngRefreshInterval * 1000
        TimerRefresh.Enabled = True
    Else
        TimerRefresh.Enabled = False
    End If
    
    If blnIsUpdateSearchTime Then
        SQLCondition.��ʼʱ�� = CDate(Format(zlDatabase.Currentdate - (mBeforeDays - 1), "yyyy-mm-dd 00:00"))
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
    End If
        
    
    '��ʼ�����������б�
    Dim iCount As Integer, rsTemp As ADODB.Recordset
    Dim strSql As String
    
    iCount = 1
    gstrSQL = "Select ִ�м�,����豸 From ҽ��ִ�з��� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡִ�м�����", mlngCur����ID)
    If rsTemp.EOF <> True Then
        ReDim mAstr��������(rsTemp.RecordCount) As String
        While rsTemp.EOF = False
            'mAstr��������(iCount) = Split(mstrCur����, "-")(1) & Nvl(rsTemp!ִ�м�)
            mAstr��������(iCount) = mlngCur����ID & ":" & Nvl(rsTemp!ִ�м�)
            iCount = iCount + 1
            rsTemp.MoveNext
        Wend
    Else
        ReDim mAstr��������(0) As String
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_�����ɼ�(Optional blnUnload As Boolean = True)
    Dim lngAdviceID As Long
    Dim lngSendNO As Long
    Dim blnReadOnly As Boolean
    Dim intState As Integer
    Dim strInfor As String
    Dim blnMoved As Boolean
    
    On Error GoTo errHandle
    
    If Not GetIsValidOfStorageDevice(mlngCur����ID) Then
      MsgBoxD Me, "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
      Exit Sub
    End If
    
    'Call frmVideoCapture.SetRestoreContainer(picVideoContainer)
    Call frmVideoDockWindow.Show
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_PACS����()
    Call OpenPacsReport
End Sub

Private Sub OpenPacsReport()
    Dim i As Integer
    
    If Not mfrmPacsReportDock Is Nothing Then
        '���жϵ�ǰ�����Ƿ�����Ҫ�򿪵Ĵ��壬������ǣ�����Ҵ�������
        If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = mfrmPacsReportDock.mlngAdviceID Then
            
            mfrmPacsReportDock.WindowState = 0  'normal
            mfrmPacsReportDock.ZOrder
            Exit Sub
        End If
    End If
    
    '���Ҵ�������,�ҵ���Ҫ�򿪵Ĵ��壬��ͨ��Zorder�Ѵ�����ʾ����ǰ��
    If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
        For i = 1 To UBound(mobjPacsReportArry)
            If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = mobjPacsReportArry(i).mlngAdviceID Then
                Set mfrmPacsReportDock = mobjPacsReportArry(i)
                
                mfrmPacsReportDock.WindowState = 0  'normal
                mfrmPacsReportDock.ZOrder
                Exit Sub
            End If
        Next i
    End If
    
    'û���ҵ���Ҫ�򿪵Ĵ��壬�Ҵ��´���,����¼��ǰ����
    Set mfrmPacsReportDock = New frmReport
    Set mfrmPacsReportDock.pobjPacsCore = mobjPacsCore
    
    Call mfrmPacsReportDock.InitReportWindow(mlngCur����ID, mlngModul, mstrPrivs, False)
    
    mfrmPacsReportDock.zlEditReport vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")), vsList.TextMatrix(vsList.Row, GetCN("���ͺ�")), Me, vsList.TextMatrix(vsList.Row, GetCN("ת��")) = 1, vsList.TextMatrix(vsList.Row, GetCN("����"))
    
    If SafeArrayGetDim(mobjPacsReportArry) = 0 Then
        ReDim mobjPacsReportArry(1) As frmReport
    Else
        ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) + 1) As frmReport
    End If
    
    Set mobjPacsReportArry(UBound(mobjPacsReportArry)) = mfrmPacsReportDock
End Sub
Private Sub cmdInfo_Click()
    On Error GoTo errHandle
    frmDegreeCard.ShowMe Val(vsList.TextMatrix(vsList.Row, GetCN("����ID"))), Val(vsList.TextMatrix(vsList.Row, GetCN("��ҳID")))
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = PicWindow.hWnd
    End If
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs 'Ȩ��
    mlngModul = glngModul 'ģ���
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    mstrCurFindtype = "���￨"
    mblnInitOk = False  '��ʼ����,��ʼ�����֮ǰ���������ݵ���ȡ
    mblnvsRefresh = False
    mlngSortCol = 0
    mintSortOrder = 0
    mlngFilterTab = 0
    
    Call InitLocalPars '����ע������
    If Not InitDepts Then Unload Me: Exit Sub '��ʼ��ҽ������
    
    ReDim gConnectedShardDir(0) As String   '��ʼ������Ŀ¼���Ӵ�
    
    Call InitMvar '��ʼ��ģ�鼶����
    
    
    '��ʼ�Ӵ���
    
    
    frmVideoCapture.mlngModul = mlngModul
    frmVideoCapture.mlngCurDeptId = mlngCur����ID
    frmVideoCapture.mstrPrivs = mstrPrivs
    frmVideoCapture.mIsShowing = False
    Set frmVideoCapture.MainFormObj = Me
    'Call mfrmCapture.InitVideoCaptureWindow(mlngCur����ID, mlngModul, mstrPrivs)
        
    Set mfrmPatholSpecimen = New frmPatholSpecimen
    Set mfrmPatholMaterial = New frmPatholMaterials
    Set mfrmPatholSlices = New frmPatholSlices
    Set mfrmPatholSpeExam = New frmPatholSpecialExamined
    Set mfrmPatholProRep = New frmPatholProcedureRep
        
    Set mfrmPacsReport = New frmReport  'PACS����
    Set mobjReport = New zlRichEPR.cDockReport
    Set mobjPacsCore = New zl9PacsCore.clsViewer
        mobjReport.PacsCore = mobjPacsCore
    Set mobjExpense = New zlCISKernel.clsDockExpense
    Set mobjInAdvice = New zlCISKernel.clsDockInAdvices
    Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices
    Set mobjInEPRs = New zlRichEPR.cDockInEPRs
    Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs
    
    If CheckPopedom(mstrPrivs, "����ȡ��") Then
        Call mfrmPatholDecalinTask.Hide
    End If
    
    Set mfrmPacsReport.pobjPacsCore = mobjPacsCore
    Call mfrmPacsReport.InitReportWindow(mlngCur����ID, mlngModul, mstrPrivs, False)
    
    Call ReadStudyListColor(mlngCur����ID)
    Call InitFilterCmd
    Call InitCommandBars
    Call InitFilterPage
    Call InitSubForm
    Call InitFaceScheme
    Call InitList(vsList)

    
    Set frmVideoCapture.pobjPacsCore = mobjPacsCore
    
    'ȥ��PACS���洰��Ŀ��ƿ�
    FormSetCaption mfrmPacsReport, False, False
    FormSetCaption mfrmPatholSpecimen, False, False
    FormSetCaption mfrmPatholMaterial, False, False
    FormSetCaption mfrmPatholSlices, False, False
    FormSetCaption mfrmPatholSpeExam, False, False
    FormSetCaption mfrmPatholProRep, False, False
    
    mblnInitOk = True '��ʼ�����
    Call RestoreWinState(Me, App.ProductName)
    
    Call RefreshList
    
    ClearCacheFolder App.Path & "\TmpImage\"    '����ʱĿ¼���ˣ�����ո�Ŀ¼
      
  
    '�ж���ʱĿ¼�Ƿ����
    If Dir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage", vbDirectory) = "" Then
        Call MkDir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage")
    End If
    
    
    Me.stbThis.Panels(3).Text = "����ҽ����" & UserInfo.����
    ReDim mobjPacsReportArry(0) As frmReport
    
    
    '��ʼ��hook����
    Set mobjSysHook = New clsHookKey
    
    mobjSysHook.ActiveHwnd = Me.hWnd
    mobjSysHook.IsOnlyActive = True
    
    Call mobjSysHook.EnableHook
End Sub


Private Sub InitFilterPage()
    Dim lngHideCount As Long
    
    lngHideCount = 0
    
    With tabFilter
        .RemoveAll
'        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        
        'ȡ��
        .InsertItem 0, "��ȡ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "��ȡ��"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����ȡ��")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
                
        .InsertItem 1, "��ȡ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "��ȡ��"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����ȡ��")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        '��Ƭ
        .InsertItem 2, "����Ƭ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "����Ƭ"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "������Ƭ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 3, "����Ƭ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "����Ƭ"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "������Ƭ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 4, "��Ƭ����", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "��Ƭ����"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "������Ƭ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        '����
        .InsertItem 5, "������", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "������"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "�����黯")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 6, "������", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "������"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "�����黯")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 7, "���߽���", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "���߽���"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "�����黯")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        '����
        .InsertItem 8, "�����", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�����"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���Ӳ���")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 9, "�ѷ���", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�ѷ���"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���Ӳ���")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 10, "���ӽ���", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "���ӽ���"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���Ӳ���")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        '��Ⱦ
        .InsertItem 11, "����Ⱦ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "����Ⱦ"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����Ⱦɫ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 12, "����Ⱦ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "����Ⱦ"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����Ⱦɫ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 13, "��Ⱦ����", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "��Ⱦ����"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����Ⱦɫ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        
        '����
        .InsertItem 14, "���ڻ���", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "���ڻ���"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���ﷴ��")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 15, "�ѻ���", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�ѻ���"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���ﷴ��")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 16, "�� ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�� ��"
        
    End With


    tabFilter.Visible = (lngHideCount < tabFilter.ItemCount - 1)
    tabFilter.Tag = (lngHideCount < tabFilter.ItemCount - 1)
    
    
    If tabFilter.Tag Then
        If Not tabFilter.Item(mlngFilterTab).Visible Then
            tabFilter.Item(tabFilter.ItemCount - 1).Selected = True
        Else
            tabFilter.Item(mlngFilterTab).Selected = True
        End If
    End If
    
    
    On Error Resume Next
    
    tabFilter.Height = tabFilter.Height - Fix((lngHideCount + 3) / 4) * 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim i As Integer
    
    On Error Resume Next
    
    Call mobjSysHook.FreeHook
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", IIf(mblncmdסԺ, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", IIf(mblncmd�ѽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", IIf(mblncmdδ��, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", IIf(mblncmd�Ǽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鲡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���˷�ʽ", mstrCurFindtype
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", mstrLocateType
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mlngSortCol
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mintSortOrder
    
    Call zlDatabase.SetPara("�������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("��������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("ϸ������", IIf(mblncmdϸ��, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("�������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("ʬ�����", IIf(mblncmdʬ��, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("���ι���", IIf(mblncmd����, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("С�걾����", IIf(mblncmdС�걾, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("���̹���", IIf(mblncmd����, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("�������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("Һ������", IIf(mblncmdҺ��, 1, 0), glngSys, glngModul)
    Call zlDatabase.SetPara("����ҳ��", tabFilter.Selected.Index, glngSys, glngModul)
    
    
'    If UBound(mblncmdӰ�����) >= 0 Then
'        strTemp = mblncmdӰ�����(0)
'    End If
'    For i = 1 To UBound(mblncmdӰ�����)
'        strTemp = strTemp & "," & mblncmdӰ�����(i)
'    Next i
'    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ӱ��������", strTemp
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    Call SaveWinState(Me, App.ProductName)
    '�ж�Ƕ��ʽ����༭���еı����Ƿ�û�б���
    If mblnPacsReport = True Then    'ʹ��PACS����༭��
        Call mfrmPacsReport.PromptModify
    End If
    
    
    '�ͷŴ������

    
    Unload frmVideoDockWindow
    Unload frmVideoCapture
    Unload mfrmPacsReport
    Unload mfrmPacsReportDock
    
    Unload mfrmPatholSpecimen
    Unload mfrmPatholMaterial
    Unload mfrmPatholSlices
    Unload mfrmPatholSpeExam
    Unload mfrmPatholProRep
    Unload mfrmPatholDecalinTask
    
    Unload mobjReport.zlGetForm
    Unload mobjExpense.zlGetForm
    Unload mobjInAdvice.zlGetForm
    Unload mobjOutAdvice.zlGetForm
    Unload mobjInEPRs.zlGetForm
    Unload mobjOutEPRs.zlGetForm
    Unload mobjQueue.zlGetForm


    For i = LBound(mobjPacsReportArry) To UBound(mobjPacsReportArry)
        Unload mobjPacsReportArry(i)
        Set mobjPacsReportArry(i) = Nothing
    Next i
    
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    
    Set mobjIDCard = Nothing
    Set mfrmPacsReport = Nothing
    Set mfrmPacsReportDock = Nothing
    
    Set mfrmPatholSpecimen = Nothing
    Set mfrmPatholMaterial = Nothing
    Set mfrmPatholSlices = Nothing
    Set mfrmPatholSpeExam = Nothing
    Set mfrmPatholProRep = Nothing
    Set mfrmPatholDecalinTask = Nothing
    
    Set mobjReport = Nothing
    Set mobjExpense = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjPacsCore = Nothing
    Set mobjQueue = Nothing
    
End Sub
Private Function GetCN(ByVal Col As String) As Integer
Dim arrCol As Variant, i As Integer
    If mstrCol = "" Then mstrCol = ConstrCol
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        If Split(arrCol(i), ";")(0) = Col Then GetCN = i: Exit Function
    Next
    GetCN = 0
End Function
Private Function GetCW(ByVal Col As String) As Long
    Dim arrCol As Variant, i As Integer
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        If Split(arrCol(i), ";")(0) = Col Then GetCW = Split(arrCol(i), ";")(1): Exit Function
    Next
    GetCW = 0
End Function
Private Sub InitLocalPars()
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
    
'��ʼ����ʱ���ز������Ը������ã�ע������Ϊ��,������أ����ˣ��������õȵ���
    On Error GoTo err
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmdסԺ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", 1))
    mblncmd�ѽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", 0))
    mblncmdδ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", 0))
    mblncmd�Ǽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鲡��", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", 1))
    
    
    mblncmd���� = Val(zlDatabase.GetPara("�������", glngSys, glngModul))
    mblncmdϸ�� = Val(zlDatabase.GetPara("ϸ������", glngSys, glngModul))
    mblncmd���� = Val(zlDatabase.GetPara("�������", glngSys, glngModul))
    mblncmdʬ�� = Val(zlDatabase.GetPara("ʬ�����", glngSys, glngModul))
    mblncmd���� = Val(zlDatabase.GetPara("��������", glngSys, glngModul))
    mblncmd���� = Val(zlDatabase.GetPara("���ι���", glngSys, glngModul))
    mblncmdС�걾 = Val(zlDatabase.GetPara("С�걾����", glngSys, glngModul))
    mblncmd���� = Val(zlDatabase.GetPara("���̹���", glngSys, glngModul))
    mblncmd���� = Val(zlDatabase.GetPara("�������", glngSys, glngModul))
    mblncmdҺ�� = Val(zlDatabase.GetPara("Һ������", glngSys, glngModul))
    mlngFilterTab = Val(zlDatabase.GetPara("����ҳ��", glngSys, glngModul))
    
    
    mstrCurFindtype = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���˷�ʽ", "����")
    mstrLocateType = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", "����")
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", "0"))
    mlngSortCol = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", 0))
    
'    strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ӱ��������", "")
'    ReDim strTempArry(0)
'    ReDim mblncmdӰ�����(0)
'    On Error Resume Next
'    strTempArry = Split(strTemp, ",")
'    If UBound(strTempArry) >= 0 Then ReDim mblncmdӰ�����(UBound(strTempArry))
'    For i = 0 To UBound(strTempArry)
'        mblncmdӰ�����(i) = IIf(UCase(strTempArry(i)) = "TRUE", True, False)
'    Next i
    
    On Error GoTo err
    mstrFirstTab = zlDatabase.GetPara("������ҳ", glngSys, mlngModul, "") 'Ϊ�ձ�ʾ��ʹ�ö��ƹ�����ҳ����
    mblnֱ�Ӽ�� = (Val(GetDeptPara(mlngCur����ID, "�ǼǺ�ֱ�Ӽ��", 0)) = 1)
    mblnOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModul, 0)) = 1)
    mblnNoShowCancel = (Val(zlDatabase.GetPara("����ʾ��ȡ���ĵǼ�", glngSys, mlngModul, 0)) = 1)
    mblnPatTrack = (Val(zlDatabase.GetPara("���˸���", glngSys, mlngModul, 0)) = 1)
    mstrRoom = zlDatabase.GetPara("ִ�м䷶Χ", glngSys, mlngModul, "")
    If mstrRoom <> "" Then mstrRoom = "'," & Replace(mstrRoom, "|", ",") & ",'"
    
    With SQLCondition '------------------------ '����������ʼ
        'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
        .ʱ������ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", 1))
        .���ݺ� = ""
        .����� = 0
        .סԺ�� = 0
        .���￨ = ""
        .���� = ""
        .�Ա� = ""
        .��ʼ���� = -1
        .�������� = -1
        .�������� = "="
        .���� = 0
        .���֤ = ""
        .IC�� = ""
        .����� = ""
        .���˿��� = 0
        .�걾��λ = ""
        .���ҽ�� = ""
        .���ҽ�� = ""
        .������� = ""
        .�������� = ""
        .������� = -1
        .Ӱ������ = ""
        .��鼼ʦ = ""
        .������ = ""
'        .Ӱ����� = ""
        .������� = ""
        .������ = ""
        .���� = ""
        .��� = ""
    End With
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    On Error GoTo errH
    
 
    str��Դ = "1,2,3"
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   

    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr("," & str��Դ & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
            Exit Function
        End If
        InitDepts = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 240, 250, DockLeftOf, Nothing)
    Pane1.Title = "����б�"
    Pane1.Handle = picList.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set Pane2 = dkpMain.CreatePane(2, 700, 250, DockRightOf, Nothing)
    Pane2.Title = "�Ӵ���"
    Pane2.Handle = PicWindow.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
End Sub
Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    Dim i As Integer

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrdock.VisualTheme = xtpThemeOfficeXP
    With Me.cbrdock.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    cbrdock.AddImageList img16 '��VB.ImageList��Tag��ID���й���
    cbrdock.EnableCustomization False
    cbrdock.ActiveMenuBar.Visible = False
    
    Set objBar = cbrdock.Add("��Դ", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ���ﲡ��"
        Set objControl = .Add(xtpControlButton, ID_סԺ, "סԺ")
            objControl.ToolTipText = "��ʾסԺ����"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ���ﲡ��"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ��첡��"
        Set objControl = .Add(xtpControlButtonPopup, ID_����, " ��  ��")
            objControl.ToolTipText = "��ʾ�����ѽ�/δ�ɲ���"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_δ��, "δ��")
            cbrPopControl.ToolTipText = "��ʾ����δ�ɲ���"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�ѽ�, "�ѽ�")
            cbrPopControl.ToolTipText = "��ʾ�����ѽɲ���"
        
        
'        '�������Ӱ�����
'        Set objControl = .Add(xtpControlButtonPopup, ID_Ӱ�����, "Ӱ�����")
'        objControl.ToolTipText = "��ʾӰ�����"
'        strSQL = "select ����,���� from Ӱ�������"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Ӱ�������")
'        i = 1
'        mintcmdӰ����� = 0
'        strTemp = ""
'        ReDim Preserve mblncmdӰ�����(rsTemp.RecordCount - 1)
'        While rsTemp.EOF = False
'            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_Ӱ����� + i, rsTemp("����"))
'            cbrPopControl.DescriptionText = rsTemp("����")
'            cbrPopControl.Style = xtpButtonIconAndCaption
'            cbrPopControl.Checked = mblncmdӰ�����(i - 1)
'            cbrPopControl.CloseSubMenuOnClick = False
'            If mblncmdӰ�����(i - 1) = True Then
'                mintcmdӰ����� = mintcmdӰ����� + 1
'                strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
'            End If
'            rsTemp.MoveNext
'            i = i + 1
'        Wend
'        If strTemp <> "" Then objControl.Caption = strTemp
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("״̬", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_�Ǽ�, "�Ǽ�")
            objControl.ToolTipText = "��ʾ�ѵǼǲ���"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ�ѱ�������"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ�Ѽ�鲡��"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ�ѱ��没��"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ����˲���"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ����ɲ���"
    End With
    
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    
    
    
    
    '----------------������ز˵�---------------------------------
    Set objBar = cbrdock.Add("����", xtpBarTop)
        objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        objBar.ContextMenuPresent = False
        
    With objBar.Controls

    
        Set objControl = .Add(xtpControlButton, ID_�������_����, "����")
            objControl.ToolTipText = "��ʾ���������ͼ��"
            
        Set objControl = .Add(xtpControlButton, ID_�������_����, "����")
            objControl.ToolTipText = "��ʾ����������ͼ��"
            
        Set objControl = .Add(xtpControlButton, ID_�������_ϸ��, "ϸ��")
            objControl.ToolTipText = "��ʾ����ϸ�����ͼ��"
            
        Set objControl = .Add(xtpControlButton, ID_�������_ʬ��, "ʬ��")
            objControl.ToolTipText = "��ʾ����ʬ�����ͼ��"
        
        Set objControl = .Add(xtpControlButton, ID_�������_����, "����")
            objControl.ToolTipText = "��ʾ����������ͼ��"
                 

                
        Set objControl = .Add(xtpControlButtonPopup, ID_�걾����, "�걾����")
            objControl.ToolTipText = "��ʾ����걾����"
        
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�걾����_����, "����")
                cbrPopControl.DescriptionText = "���α걾"
                cbrPopControl.Style = xtpButtonIconAndCaption
                cbrPopControl.Checked = True
                cbrPopControl.CloseSubMenuOnClick = False
                
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�걾����_С�걾, "С�걾")
                cbrPopControl.DescriptionText = "С�걾"
                cbrPopControl.Style = xtpButtonIconAndCaption
                cbrPopControl.Checked = True
                cbrPopControl.CloseSubMenuOnClick = False
                
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�걾����_����, "����")
                cbrPopControl.DescriptionText = "����ϸ��"
                cbrPopControl.Style = xtpButtonIconAndCaption
                cbrPopControl.Checked = True
                cbrPopControl.CloseSubMenuOnClick = False
                
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�걾����_����, "����")
                cbrPopControl.DescriptionText = "����ϸ��"
                cbrPopControl.Style = xtpButtonIconAndCaption
                cbrPopControl.Checked = True
                cbrPopControl.CloseSubMenuOnClick = False
                
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�걾����_Һ��, "Һ��")
                cbrPopControl.DescriptionText = "Һ��ϸ��"
                cbrPopControl.Style = xtpButtonIconAndCaption
                cbrPopControl.Checked = True
                cbrPopControl.CloseSubMenuOnClick = False
            
    End With
            
    
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next

    
    
    
    
    
    
    
    
    
    Set objBar = cbrdock.Add("����", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    Set objPopbar = objBar.Controls.Add(xtpControlPopup, ID_���ҷ�ʽ, "���ҷ�ʽ")
        objPopbar.ID = ID_���ҷ�ʽ
        objPopbar.flags = xtpFlagRightAlign
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_����ֵ, "����ֵ")
        objCusControl.Handle = txtFilter.hWnd
        objCusControl.flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_��ʼ����, "��ʼ����")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_����סԺ, "����")
    objControl.ToolTipText = "ֻ��ʾ����סԺ����¼"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter

    
    With cbrdock.KeyBindings
        .Add FCONTROL, vbKey0, ID_����
        .Add FCONTROL, vbKey1, ID_סԺ
        .Add FCONTROL, vbKey2, ID_����
        .Add FCONTROL, vbKey3, ID_���
        
        .Add FCONTROL, vbKey4, ID_�Ǽ�
        .Add FCONTROL, vbKey5, ID_����
        .Add FCONTROL, vbKey6, ID_���
        .Add FCONTROL, vbKey7, ID_����
        .Add FCONTROL, vbKey8, ID_���
        .Add FCONTROL, vbKey9, ID_���
        .Add FCONTROL, Asc("G"), ID_��ʼ����
    End With
    cbrdock.RecalcLayout
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    

'�˵�����
'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.Title = "�˵�"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)"): cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.IconId = 102
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "������ӡ(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�嵥��ӡ(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&O)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_SendImg, "����ͼ��(&T)"): cbrControl.IconId = 3061
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "�����б�")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "���(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "���Ǽ�(&I)"): cbrControl.IconId = 211: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CopyCheck, "���ƵǼ�(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "ȡ���Ǽ�(&R)"): cbrControl.IconId = 742
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "�ٻ�ȡ��(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "�޸���Ϣ(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "��鱨��(&L)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ������(&D)"): cbrControl.IconId = 743
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "����Ӱ��(&C)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 505: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "ȡ������(&B)"): cbrControl.IconId = 506
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Review, "��ע(&R)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 232
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportRelease, "����"): cbrControl.ToolTipText = "���淢��": cbrControl.IconId = 3013
        
        '���Խ�������ԣ�����ʾ����˵�
        If mblnIgnoreResult = False Then
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "�����(&X)"): cbrControl.ID = conMenu_Manage_Result
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "�������(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "�������(&X)"): cbrPopControl.IconId = 3507
        End If
        
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "Ӱ������(&Y)"): cbrControl.ID = conMenu_Manage_Quality
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "�׵�(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "�ҵ�(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_GChannel, "��ɫͨ��(&G)"): cbrControl.ID = conMenu_Manage_GChannel
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelOk, "���(&J)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelCancel, "ȡ��(&Y)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "�ޱ������(&F)"): cbrControl.IconId = 216: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "�ޱ������(&U)"):  cbrControl.IconId = 3012
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "������(&E)"): cbrControl.IconId = 225
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ�����(&U)"): cbrControl.IconId = 219
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet, "��������"): cbrControl.IconId = 803
    End With
    
    
'Begin----------------------�������˵�---------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholManage, "�������(&P)", -1, False)
    cbrMenuBar.ID = conMenu_PatholManage
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Antibody_Manage, "�������(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Meal_Manage, "�ײ�ά��(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Decalin_Task, "�Ѹ��������(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Pathol_Request, "��������(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Report_Delay, "�ӳٵǼ�(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Con_Request, "��������(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Con_Feedback, "���ﷴ��(&F)")
    End With
    
    
'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '�����˵�
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_LocateType, "��λ��ʽ(&G)"): cbrControl.ID = conMenu_Manage_LocateType
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "������"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "���ٹ���(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)")
    End With

'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������", -1, False)
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    

'��ȡ��������ģ��ı���(��������ģ���)
'-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModul, mstrPrivs)
    
'----------------------�����------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print '��ӡ------------------Ctrl+P
        .Add 0, VK_F12, conMenu_File_Parameter      '��������--------------F12
        
        .Add 0, VK_F2, conMenu_Manage_Regist       '�Ǽ�-----------------F2
        .Add 0, VK_F7, conMenu_Manage_CopyCheck    '���ƵǼ�-------------F7
        .Add 0, VK_F4, conMenu_Manage_Receive       '����-----------------F4
        .Add 0, VK_F9, conMenu_Manage_ClearUp       '���ر���------------F9
        .Add 0, VK_F6, conMenu_Manage_Complete         '��˱���----------F6
        
        
        .Add 0, VK_F1, conMenu_Help_Help              '����-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           'ˢ��-------------F5
        .Add FCONTROL, Asc("G"), conMenu_Manage_LocateType    '��λ��ʽ---------Ctrl+F
        .Add 0, VK_F3, conMenu_View_Filter            '����-------------F3
    End With

    
'---------------------�������Ͻǵ�ǰ����----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "������")
            cbrControl.ID = conMenu_View_Filter * 10#: cbrControl.flags = xtpFlagRightAlign: cbrControl.Category = "Main"
        
        Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Manage_LocateType, "��ʶ��(&D)")
            cbrMenuBar.ID = conMenu_Manage_LocateType
            cbrMenuBar.flags = xtpFlagRightAlign
            cbrMenuBar.Category = "Main"
            
        Set cbrCustom = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Manage_LocateValue, "��λ����")
            cbrCustom.Handle = txtLocate.hWnd
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Style = xtpButtonIconAndCaption
            cbrCustom.Category = "Main"
            
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, comMenu_Cap_Process, "�����ɼ�")
            cbrControl.ToolTipText = "�����ɼ�"
            cbrControl.flags = xtpFlagRightAlign
            cbrControl.Category = "Main"
    

'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
'    cbrToolBar.EnableDocking xtpFlagStretched '+ xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.IconId = 102: cbrControl.ToolTipText = "����Ԥ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.IconId = 103: cbrControl.ToolTipText = "�����ӡ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Ǽ�"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����"): cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ��"): cbrControl.IconId = 743: cbrControl.ToolTipText = "ȡ������"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Review, "��ע"):  cbrControl.BeginGroup = True: cbrControl.IconId = 232
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportRelease, "����"): cbrControl.ToolTipText = "���淢��": cbrControl.IconId = 3013
        
        '���Խ�������ԣ�����ʾ���������
        If mblnIgnoreResult = False Then
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "���"): cbrControl.ID = conMenu_Manage_Result: cbrControl.IconId = 3506: cbrControl.ToolTipText = "�����������"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "����(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "����(&Y)"): cbrPopControl.IconId = 3507
        End If
        
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "����"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061: cbrControl.ToolTipText = "Ӱ������"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "�׼�(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "�Ҽ�(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "���"): cbrControl.IconId = 225: cbrControl.ToolTipText = "����������"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
    End With
End Sub
Private Sub InitSubForm()
Dim i As Integer
Dim strFirstTitle As String

    With TabWindow
        .RemoveAll
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        
        
        
        .InsertItem 0, "Ӱ��ɼ�", picVideoContainer.hWnd, conMenu_Cap_Dynamic
        .Item(TabWindow.ItemCount - 1).Tag = "Ӱ��ɼ�"
        .Item(TabWindow.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "��Ƶ�ɼ�")
        If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "Ӱ��ɼ�", strFirstTitle)
        
        
        .InsertItem 1, "�걾����", mfrmPatholSpecimen.hWnd, 10015
        .Item(TabWindow.ItemCount - 1).Tag = "�걾����"
        .Item(TabWindow.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "�걾����")
        If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "�걾����", strFirstTitle)
        
        
        .InsertItem 2, "����ȡ��", mfrmPatholMaterial.hWnd, 10016
        .Item(TabWindow.ItemCount - 1).Tag = "����ȡ��"
        .Item(TabWindow.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����ȡ��")
        If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "����ȡ��", strFirstTitle)
        
        
        .InsertItem 3, "������Ƭ", mfrmPatholSlices.hWnd, 10017
        .Item(TabWindow.ItemCount - 1).Tag = "������Ƭ"
        .Item(TabWindow.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "������Ƭ")
        If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "������Ƭ", strFirstTitle)
        
        
        .InsertItem 4, "������", mfrmPatholSpeExam.hWnd, 10018
        .Item(TabWindow.ItemCount - 1).Tag = "������"
        .Item(TabWindow.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "����Ⱦɫ") Or CheckPopedom(mstrPrivs, "���Ӳ���")
        If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "������", strFirstTitle)
        
        
        .InsertItem 5, "����/�ؼ챨��", mfrmPatholProRep.hWnd, 10019
        .Item(TabWindow.ItemCount - 1).Tag = "����/�ؼ챨��"
        .Item(TabWindow.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "��������") _
            Or CheckPopedom(mstrPrivs, "��Ⱦ����") Or CheckPopedom(mstrPrivs, "���ӱ���") Or CheckPopedom(mstrPrivs, "���߱���") Or CheckPopedom(mstrPrivs, "�����ؼ챨�����")
        If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "����/�ؼ챨��", strFirstTitle)
        
       
       
       
       
       
        If GetInsidePrivs(p���Ʊ������, True) <> "" Then
            If mblnPacsReport = True Then
                .InsertItem 6, "������", mfrmPacsReport.hWnd, conMenu_Edit_Compend '10008 '
            Else
                .InsertItem 6, "������", mobjReport.zlGetForm.hWnd, conMenu_Edit_Compend '10008 '
            End If
            .Item(TabWindow.ItemCount - 1).Tag = "������д"
            
            If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "������д", strFirstTitle)
        End If
        
        If GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
            .InsertItem 7, "���ü�¼", mobjExpense.zlGetForm.hWnd, conMenu_Manage_Request '10007  '
            .Item(TabWindow.ItemCount - 1).Tag = "�������"
            
            If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "�������", strFirstTitle)
        End If
        
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Then
            .InsertItem 8, "ҽ����¼", mobjInAdvice.zlGetForm.hWnd, conMenu_Edit_NewItem ' 10010 '
            .Item(TabWindow.ItemCount - 1).Tag = "סԺҽ��"
            
            If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "סԺҽ��", strFirstTitle)
        End If
        
        If GetInsidePrivs(p����ҽ���´�, True) <> "" Then
            .InsertItem 9, "ҽ����¼", mobjOutAdvice.zlGetForm.hWnd, conMenu_Edit_NewItem ' 10010 '
            .Item(TabWindow.ItemCount - 1).Tag = "����ҽ��": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "����ҽ��", strFirstTitle)
        End If
        
        If GetInsidePrivs(pסԺ��������, True) <> "" Then
            .InsertItem 10, "������¼", mobjInEPRs.zlGetForm.hWnd, conMenu_Edit_Archive ' 10009 '
            .Item(TabWindow.ItemCount - 1).Tag = "סԺ����"
            
            If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "סԺ����", strFirstTitle)
        End If
        
        If GetInsidePrivs(p���ﲡ������, True) <> "" Then
            .InsertItem 11, "������¼", mobjOutEPRs.zlGetForm.hWnd, conMenu_Edit_Archive ' 10009 '
            .Item(TabWindow.ItemCount - 1).Tag = "���ﲡ��": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If .Item(TabWindow.ItemCount - 1).Visible Then strFirstTitle = IIf(Trim(strFirstTitle) = "", "���ﲡ��", strFirstTitle)
        End If


        If Trim(mstrFirstTab) <> "" Then strFirstTitle = mstrFirstTab
        
        i = .ItemCount
        
        If strFirstTitle <> "" Then
            If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Set frmVideoCapture.ParentContainerObj = picVideoContainer
            
            For i = 0 To .ItemCount - 1
                If InStr(.Item(i).Tag, strFirstTitle) > 0 And .Item(i).Visible Then
                    .Item(i).Selected = True

                    
                    If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then
                        If InStr("������д", strFirstTitle) > 0 Then
                            If mblnPacsReport = True Then Call mfrmPacsReport.ShowVideoWindow
                        ElseIf InStr("Ӱ��ɼ�", strFirstTitle) > 0 Then
                            Call frmVideoCapture.ShowVideoWindow(picVideoContainer)
                        Else
                            Call frmVideoCapture.ShowVideoWindow(picVideoContainer)
                        End If
                    End If
                    
                    Exit Sub
                End If
            Next
        End If
        
        '���δ�ҵ���Ч��tabҳ����ʹ�õ�һ���ɼ���tab
        If i = .ItemCount Then
            For i = 0 To .ItemCount - 1
                If .Item(i).Visible Then
                    .Item(i).Selected = True
                    Exit For
                End If
            Next i
        End If
        
'        Call frmVideoCapture.SetRestoreContainer(picVideoContainer) 'RefreshTabWindow�л�Ը÷������е���
        If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Call frmVideoCapture.ShowVideoWindow(picVideoContainer)
    End With


End Sub



Private Sub InitList(lst As VSFlexGrid)
'��ʼ�����
Dim C·�� As Long, C���� As Long, C��Դ As Long, C���� As Long, C���� As Long, C���� As Long, C���� As Long, C������ As Long, C�Ա� As Long, C���� As Long
Dim C��ʶ�� As Long, Cҽ������ As Long, C��λ���� As Long, C����ʱ�� As Long, C����ʱ�� As Long, C����ҽ�� As Long, C����ִ�й��� As Long
Dim C��� As Long, C���� As Long, CӤ�� As Long, C�Ǽ��� As Long, C������ As Long, C����� As Long, C������� As Long
Dim C��ɫͨ�� As Long, C�����ӡ As Long, C������ As Long, C������ As Long, C��ͼʱ�� As Long, C������� As Long
Dim C������ As Long, C����ID As Long, C��ҳID As Long, C�Һŵ� As Long, C���˿���ID As Long, Cҽ��ID As Long, C���ͺ� As Long, C���UID As Long
Dim C���״̬ As Long, CNO As Long, C��¼���� As Long, Cת�� As Long, C���� As Long, C��ǰ����ID As Long, C���淢�� As Long, C����� As Long, C������� As Long
Dim C��Ϸ��� As Long, C����ID As Long, C���˿��� As Long, C���￨�� As Long, C���ݺ� As Long, C���֤�� As Long
Dim C�շ� As Long

    If mstrCol = "" Then
        mstrCol = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, ConstrCol)
        '�ж��Ƿ��޸Ĺ���ʾ������������޸Ĺ������ȡĬ��ֵ�������Ƕ�ȡע���
        If UBound(Split(mstrCol, "|")) <> UBound(Split(ConstrCol, "|")) Then
            mstrCol = ConstrCol
        End If
    End If
    With lst
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 54
        '��ȡ����
        C·�� = GetCN("·��")
        C���� = GetCN("����"):           C��Դ = GetCN("��Դ"):          C���� = GetCN("����")
        C���� = GetCN("����"):          C���� = GetCN("����"):          C���� = GetCN("����")
        C������ = GetCN("������"):  C�Ա� = GetCN("�Ա�"):          C���� = GetCN("����")
        C��ʶ�� = GetCN("��ʶ��"):      Cҽ������ = GetCN("ҽ������"):  C��λ���� = GetCN("��λ����")
        C����ʱ�� = GetCN("����ʱ��"):  C����ʱ�� = GetCN("����ʱ��")
        C����ҽ�� = GetCN("����ҽ��"):   C��� = GetCN("���"):          C���� = GetCN("����")
        CӤ�� = GetCN("Ӥ��"):          C�Ǽ��� = GetCN("�Ǽ���"):      C������ = GetCN("������")
        C����� = GetCN("�����"):      C������� = GetCN("�������")
        C��ɫͨ�� = GetCN("��ɫͨ��"):  C�����ӡ = GetCN("�����ӡ"):  C������ = GetCN("������")
        C������ = GetCN("������"):      C��ͼʱ�� = GetCN("��ͼʱ��")
        C������� = GetCN("�������"):  C������ = GetCN("������"):  C����ID = GetCN("����ID")
        C��ҳID = GetCN("��ҳID"):      C�Һŵ� = GetCN("�Һŵ�"):      Cҽ��ID = GetCN("ҽ��ID")
        C���ͺ� = GetCN("���ͺ�"):      C���˿���ID = GetCN("���˿���ID"): C���UID = GetCN("���UID")
        C���״̬ = GetCN("���״̬"):  CNO = GetCN("NO"):              C��¼���� = GetCN("��¼����")
        Cת�� = GetCN("ת��"):          C���� = GetCN("����"):          C��ǰ����ID = GetCN("��ǰ����ID")
        C���淢�� = GetCN("���淢��"):  C����� = GetCN("�����"):      C������� = GetCN("�������")
        C��Ϸ��� = GetCN("��Ϸ���"):  C����ID = GetCN("����ID"):      C���˿��� = GetCN("���˿���")
        C���￨�� = GetCN("���￨��"):  C���ݺ� = GetCN("���ݺ�"):      C���֤�� = GetCN("���֤��")
        C�շ� = GetCN("�շ�"):          C����ִ�й��� = GetCN("����ִ�й���")
        

        '��ȡ��ָ���п�
        .ColWidth(C·��) = GetCW("·��")
        .ColWidth(C����) = GetCW("����"):           .ColWidth(C��Դ) = GetCW("��Դ"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(C����) = GetCW("����"):           .ColWidth(C����) = GetCW("����"):           .ColWidth(C�����) = GetCW("�����"): .ColWidth(C����ִ�й���) = GetCW("����ִ�й���")
        .ColWidth(C������) = GetCW("������"):   .ColWidth(C�Ա�) = GetCW("�Ա�"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(C��ʶ��) = GetCW("��ʶ��"):       .ColWidth(Cҽ������) = GetCW("ҽ������"):   .ColWidth(C��λ����) = GetCW("��λ����")
        .ColWidth(C����ʱ��) = GetCW("����ʱ��"):   .ColWidth(C����ʱ��) = GetCW("����ʱ��")
        .ColWidth(C����ҽ��) = GetCW("����ҽ��"):   .ColWidth(C���) = GetCW("���"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(CӤ��) = GetCW("Ӥ��"):           .ColWidth(C�Ǽ���) = GetCW("�Ǽ���"):       .ColWidth(C������) = GetCW("������")
        .ColWidth(C�����) = GetCW("�����"):       .ColWidth(C�������) = GetCW("�������")
        .ColWidth(C��ɫͨ��) = GetCW("��ɫͨ��"):   .ColWidth(C�����ӡ) = GetCW("�����ӡ"):   .ColWidth(C������) = GetCW("������")
        .ColWidth(C������) = GetCW("������"):       .ColWidth(C��ͼʱ��) = GetCW("��ͼʱ��")
        .ColWidth(C�������) = GetCW("�������"):   .ColWidth(C������) = GetCW("C������"):   .ColWidth(C����ID) = GetCW("����ID")
        .ColWidth(C��ҳID) = GetCW("��ҳID"):       .ColWidth(C�Һŵ�) = GetCW("�Һŵ�"):       .ColWidth(Cҽ��ID) = GetCW("ҽ��ID")
        .ColWidth(C���ͺ�) = GetCW("���ͺ�"):       .ColWidth(C���˿���ID) = GetCW("���˿���ID"): .ColWidth(C���UID) = GetCW("���UID")
        .ColWidth(C���״̬) = GetCW("���״̬"):   .ColWidth(CNO) = GetCW("NO"):               .ColWidth(C��¼����) = GetCW("��¼����")
        .ColWidth(Cת��) = GetCW("ת��"):           .ColWidth(C����) = GetCW("����"):           .ColWidth(C��ǰ����ID) = GetCW("��ǰ����ID")
        .ColWidth(C���淢��) = GetCW("���淢��"):   .ColWidth(C����) = GetCW("����"):       .ColWidth(C�������) = GetCW("�������")
        .ColWidth(C��Ϸ���) = GetCW("��Ϸ���"):
        .ColWidth(C����ID) = GetCW("����ID"):
        .ColWidth(C���˿���) = GetCW("���˿���")
        .ColWidth(C���￨��) = GetCW("���￨��"):
        .ColWidth(C���ݺ�) = GetCW("���ݺ�"):
        .ColWidth(C���֤��) = GetCW("���֤��")
        .ColWidth(C�շ�) = GetCW("�շ�"):
        
        
        '������
        .Cell(flexcpData, 0, C·��) = "·��"
        .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C��Դ) = "��Դ":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C�����) = "�����": .Cell(flexcpData, 0, C����ִ�й���) = "����ִ�й���"
        .Cell(flexcpData, 0, C������) = "������":       .Cell(flexcpData, 0, C�Ա�) = "�Ա�":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, C��ʶ��) = "��ʶ��":           .Cell(flexcpData, 0, Cҽ������) = "ҽ������":       .Cell(flexcpData, 0, C��λ����) = "��λ����"
        .Cell(flexcpData, 0, C����ʱ��) = "����ʱ��":       .Cell(flexcpData, 0, C����ʱ��) = "����ʱ��"
        .Cell(flexcpData, 0, C����ҽ��) = "����ҽ��":       .Cell(flexcpData, 0, C���) = "���":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, CӤ��) = "Ӥ��":               .Cell(flexcpData, 0, C�Ǽ���) = "�Ǽ���":           .Cell(flexcpData, 0, C������) = "������"
        .Cell(flexcpData, 0, C�����) = "�����":           .Cell(flexcpData, 0, C�������) = "�������"
        .Cell(flexcpData, 0, C��ɫͨ��) = "��ɫͨ��":       .Cell(flexcpData, 0, C�����ӡ) = "�����ӡ":       .Cell(flexcpData, 0, C������) = "������"
        .Cell(flexcpData, 0, C������) = "������":           .Cell(flexcpData, 0, C��ͼʱ��) = "��ͼʱ��"
        .Cell(flexcpData, 0, C�������) = "�������":       .Cell(flexcpData, 0, C������) = "������":       .Cell(flexcpData, 0, C����ID) = "����ID"
        .Cell(flexcpData, 0, C��ҳID) = "��ҳID":           .Cell(flexcpData, 0, C�Һŵ�) = "�Һŵ�":           .Cell(flexcpData, 0, C���˿���ID) = "���˿���ID"
        .Cell(flexcpData, 0, Cҽ��ID) = "ҽ��ID":           .Cell(flexcpData, 0, C���ͺ�) = "���ͺ�":           .Cell(flexcpData, 0, C���UID) = "���UID"
        .Cell(flexcpData, 0, C���״̬) = "���״̬":       .Cell(flexcpData, 0, CNO) = "NO":                   .Cell(flexcpData, 0, C��¼����) = "��¼����"
        .Cell(flexcpData, 0, Cת��) = "ת��":               .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C��ǰ����ID) = "��ǰ����ID"
        .Cell(flexcpData, 0, C���淢��) = "���淢��":       .Cell(flexcpData, 0, C����) = "����":           .Cell(flexcpData, 0, C�������) = "�������"
        .Cell(flexcpData, 0, C��Ϸ���) = "��Ϸ���":       .Cell(flexcpData, 0, C����ID) = "����ID":           .Cell(flexcpData, 0, C���˿���) = "���˿���"
        .Cell(flexcpData, 0, C���￨��) = "���￨��":       .Cell(flexcpData, 0, C���ݺ�) = "���ݺ�":           .Cell(flexcpData, 0, C���֤��) = "���֤��"
        .Cell(flexcpData, 0, C�շ�) = "�շ�":
        
        '��ʾ������
        .TextMatrix(0, C·��) = "·��"
        Set .Cell(flexcpPicture, 0, C����) = Imglist.ListImages("����").Picture
        Set .Cell(flexcpPicture, 0, C��Դ) = Imglist.ListImages("סԺ").Picture
        Set .Cell(flexcpPicture, 0, C����) = Imglist.ListImages("����").Picture
        Set .Cell(flexcpPicture, 0, C�շ�) = Imglist.ListImages("�շ�").Picture
        .TextMatrix(0, C����) = "��":               .TextMatrix(0, C����) = "����":             .TextMatrix(0, C�����) = "�����": .TextMatrix(0, C����ִ�й���) = "����ִ�й���"
        .TextMatrix(0, C������) = "������":     .TextMatrix(0, C�Ա�) = "�Ա�":             .TextMatrix(0, C����) = "����"
        .TextMatrix(0, C��ʶ��) = "��ʶ��":         .TextMatrix(0, Cҽ������) = "ҽ������":     .TextMatrix(0, C��λ����) = "��λ����"
        .TextMatrix(0, C����ʱ��) = "����ʱ��":     .TextMatrix(0, C����ʱ��) = "����ʱ��"
        .TextMatrix(0, C����ҽ��) = "����ҽ��":     .TextMatrix(0, C���) = "���":             .TextMatrix(0, C����) = "����"
        .TextMatrix(0, CӤ��) = "Ӥ��":             .TextMatrix(0, C�Ǽ���) = "�Ǽ���":         .TextMatrix(0, C������) = "������"
        .TextMatrix(0, C�����) = "�����":         .TextMatrix(0, C�������) = "�������"
        .TextMatrix(0, C��ɫͨ��) = "��ɫͨ��":     .TextMatrix(0, C�����ӡ) = "�����ӡ":     .TextMatrix(0, C������) = "������"
        .TextMatrix(0, C������) = "������":         .TextMatrix(0, C��ͼʱ��) = "��ͼʱ��"
        .TextMatrix(0, C�������) = "�������":     .TextMatrix(0, C������) = "������":     .TextMatrix(0, C����ID) = "����ID"
        .TextMatrix(0, C��ҳID) = "��ҳID":         .TextMatrix(0, C�Һŵ�) = "�Һŵ�":         .TextMatrix(0, C���˿���ID) = "���˿���ID"
        .TextMatrix(0, Cҽ��ID) = "ҽ��ID":         .TextMatrix(0, C���ͺ�) = "���ͺ�":         .TextMatrix(0, C���UID) = "���UID"
        .TextMatrix(0, C���״̬) = "���״̬":     .TextMatrix(0, CNO) = "NO":                 .TextMatrix(0, C��¼����) = "��¼����"
        .TextMatrix(0, Cת��) = "ת��":             .TextMatrix(0, C����) = "����":             .TextMatrix(0, C��ǰ����ID) = "��ǰ����ID"
        .TextMatrix(0, C���淢��) = "���淢��":      .TextMatrix(0, C����) = "����":        .TextMatrix(0, C�������) = "�������"
        .TextMatrix(0, C��Ϸ���) = "��Ϸ���":     .TextMatrix(0, C����ID) = "����ID":         .TextMatrix(0, C���˿���) = "���˿���"
        .TextMatrix(0, C���￨��) = "���￨��":     .TextMatrix(0, C���ݺ�) = "���ݺ�":         .TextMatrix(0, C���֤��) = "���֤��"
        
        
        Dim i As Integer
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignLeftCenter
        Next

        '��ȡ�����ò����б������
        .FontName = zlDatabase.GetPara("�����б���������", glngSys, mlngModul, "����")
        .FontSize = Val(zlDatabase.GetPara("�����б������ֺ�", glngSys, mlngModul, 9))
        .FontBold = zlDatabase.GetPara("�����б����ݴ���", glngSys, mlngModul, 0) = 1
        .FontItalic = zlDatabase.GetPara("�����б�����б��", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("�����б��ͷ����", glngSys, mlngModul, "����")
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(zlDatabase.GetPara("�����б��ͷ�ֺ�", glngSys, mlngModul, 9))
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("�����б��ͷ����", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("�����б��ͷб��", glngSys, mlngModul, 0) = 1
        .Editable = flexEDNone
    End With
End Sub

Private Sub mfrmCapture_StudyChangeEvent(lngAdviceID As Long, strPatientName As String, blnIsLock As Boolean)
    '�޸ı�ǩҳ����ʾ��ʽ�ͱ���
    Dim i As Integer
    
    For i = 0 To TabWindow.ItemCount - 1
        If TabWindow(i).Caption Like "*Ӱ��ɼ�*" Then
            If blnIsLock Then
                TabWindow(i).Image = 10013
                TabWindow(i).Caption = "��" & strPatientName & "�� Ӱ��ɼ�"
            Else
                TabWindow(i).Image = conMenu_Cap_Dynamic
                TabWindow(i).Caption = "Ӱ��ɼ�"
            End If
            
            'TabWindow(i).Image
            
            Exit For
        End If
    Next i
End Sub



Private Sub mfrmPacsReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
    
    'Ƕ��ʽ��д����ʱ������֮�����¿����Զ�ˢ�¹���
    Call subTriggleRefreshTimer(True)
End Sub

Private Sub mfrmPacsReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub mfrmPacsReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Form)
    Call AfterReportSaved(lngOrderID, frmOwnerForm)
End Sub

Private Sub mfrmPacsReport_BeforeEdit()
Dim lngOrderID As Long

    On Error GoTo errHandle
    lngOrderID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
    If CheckConcurrentReport(Me, lngOrderID) Then '����Ƿ��������ڲ�������
        Call UpdateReporter(lngOrderID, UserInfo.����)
    Else
        Call mfrmPacsReport.PromptModify(True)
    End If
    
    'Ƕ��ʽ��д����ʱ���༭����֮ǰ���ȹر��Զ�ˢ�¹���
    Call subTriggleRefreshTimer(False)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsReportDock_AfterOpen()
    Call AfterReportOpen
End Sub

Private Sub mfrmPacsReportDock_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmSample_StateChanged(lngState As Long, str����� As String, str��������� As String)
    vsList.TextMatrix(vsList.Row, GetCN("�������")) = IIf(lngState = 1, "��", IIf(lngState = 2, "��", ""))
    If lngState = 1 Then
        vsList.TextMatrix(vsList.Row, GetCN("�����")) = str�����
        vsList.TextMatrix(vsList.Row, GetCN("Ӱ�����")) = str���������
    End If
End Sub



Private Sub mfrmPatholMaterial_OnMaterialSure(ByVal lngAdviceID As Long)
'�걾ȡ��ִ���¼�
On Error Resume Next
    Call RefreshList(lngAdviceID)
End Sub

Private Sub mfrmPatholSlices_OnSlicesSure(ByVal lngAdviceID As Long)
'������Ƭִ���¼�
On Error Resume Next
    Call RefreshList(lngAdviceID)
End Sub

Private Sub mfrmPatholSpecimen_OnAccept(ByVal lngAdviceID As Long)
'�걾����ִ���¼�
On Error Resume Next
    Call RefreshList(lngAdviceID)
End Sub

Private Sub mfrmPatholSpeExam_OnSpeExamSure(ByVal lngAdviceID As Long)
'�����ؼ�ִ���¼�
On Error Resume Next
    Call RefreshList(lngAdviceID)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtFilter.Text = "" And Me.ActiveControl Is txtFilter Then
        IDKind.IDKind = IDKinds.C2���֤��
        mstrCurFindtype = "���֤"
        txtFilter = strID
        Call txtFilter_KeyDown(vbKeyReturn, 0)
    ElseIf txtLocate.Text = "" And Me.ActiveControl Is txtLocate Then
        IDKind.IDKind = IDKinds.C2���֤��
        mstrLocateType = "���֤"
        txtLocate = strID
        Call txtLocate_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub mobjInAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lngҽ��ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select ҽ��ID FROM ����ҽ������ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ��ID", CLng(����ID))
    If rsTemp.EOF Then Exit Sub
    
    lngҽ��ID = Nvl(rsTemp!ҽ��ID, 0)
    mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False '�Բ���Edit��ʽˢ�¶���
    
    Set cbrControl = cbrMain.FindControl(, conMenu_Help_Help, , True)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '���ò��ı���
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
    Call OpenViewer(mobjPacsCore, ҽ��ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lngҽ��ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select ҽ��ID FROM ����ҽ������ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ��ID", CLng(����ID))
    If rsTemp.EOF Then Exit Sub
    
    lngҽ��ID = Nvl(rsTemp!ҽ��ID, 0)
    mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False '�Բ���Edit��ʽˢ�¶���
    
    Set cbrControl = cbrMain.FindControl(, conMenu_Help_Help, , True)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '���ò��ı���
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
    Call OpenViewer(mobjPacsCore, ҽ��ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    If mblnPacsReport = True Then
        mfrmPacsReport.RefPacsPic 'ˢ��ͼƬ
        If Not mfrmPacsReportDock Is Nothing Then
            mfrmPacsReportDock.RefPacsPic 'ˢ��ͼƬ
        End If
    Else
        mobjReport.RefPacsPic 'ˢ��ͼƬ
    End If
End Sub
Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub
Public Sub EditorClosed(ByVal lngOrderID As Long)
    Dim i As Integer
    Dim j As Integer
    
    Call UpdateReporter(lngOrderID, "")
    
    '����PACS����༭���Ĵ�������
    On Error Resume Next
    If mblnPacsReport = True Then
        '���Ҵ������飬�ҵ���Ӧ�Ĵ��ڲ�ɾ��
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If mobjPacsReportArry(i).mlngAdviceID = lngOrderID Then
                    '��������ɾ��
                    For j = i To UBound(mobjPacsReportArry)
                        Set mobjPacsReportArry(j) = mobjPacsReportArry(j + 1)
                    Next j
                    ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) - 1) As frmReport
                    Exit For
                End If
            Next i
        End If
        
        If Not mfrmPacsReportDock Is Nothing Then
            If lngOrderID = mfrmPacsReportDock.mlngAdviceID Then
                '�رյ�ǰ���洰�ڣ�����ǰ�������óɿ�
                Set mfrmPacsReportDock = Nothing
            End If
        End If
    End If
End Sub
Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
    On Error GoTo errHandle
    gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
    Call RefreshList
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterOpen(ByVal intEditType As zlRichEPR.EditTypeEnum)
    Call AfterReportOpen
End Sub

Private Sub AfterReportOpen()
Dim lngOrderID As Long
    lngOrderID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
    Call UpdateReporter(lngOrderID, UserInfo.����)
End Sub
Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub
Public Sub AfterPrinted(lngOrderID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "���´�ӡ���"
    If Not mblnIgnoreResult And mintResultInput = 2 Then
        strSql = "Select �������  From  ����ҽ������ Where ҽ��id= [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�������", lngOrderID)
        
        If IsNull(rsTemp!�������) Then  '�ڱ���ʱ��ʾ���������
            Call PromptResult(lngOrderID, mlngModul, Me)
        End If
    End If
    
    If mblnPrintCommit = True Then
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    
    Call RefreshList
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID, Me)
End Sub

Public Sub AfterReportSaved(lngOrderID As Long, frmOwnerForm As Form)
'���汨��֮��Ĵ���
'ִ�й��̣�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����

    Dim intState As Integer, lngSendId As Long
    Dim strǩ�� As String
    Dim str������ As String
    Dim str������ As String
    Dim bln���������� As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()
    
    On Error GoTo errHandle
    
    If mblnPacsReport = True Then
'        mfrmPacsReport.zlRefresh 0, 0, 0
    Else
        mobjReport.zlRefresh 0, mlngCur����ID, False
    End If

    '��ȡ���μ���ִ�й���
    intState = getStudyState(lngOrderID, lngSendId, str������, strǩ��, str������, bln����������)
    
    'intState =1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--����ɣ������̲������������ֵ��
    If intState = 2 Or intState = 3 Then
        gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & intState & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & str������ & "','')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    Else
        If intState = 4 Then
            '���ǩ�������һ��ǩ��Ϊҽʦ,ִ�й���Ϊ�ѱ���
            '�п��ܵ���� 1-ҽʦ��N��ǩ�� 2-���μ������һ����ǩ 3-�޶�ģʽ�±���(ǩ������=0)
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & intState & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            'Ӧ����д�����˲�׼ȷ�����˵�ʱ�򣬻��˵����Ǳ����ˣ����ǲ��Ǳ��洴����
            'ҽ�����ǩ��,�����ǵ�N�Σ���ʱ����������Ҫ���棬��������Ҫ���;
            gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & str������ & "','')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        ElseIf intState = 5 Then
            '���ǩ�������μ����ϼ���ǩ����ǩ������>=2,ִ�й���Ϊ�����
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & intState & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & str������ & "','" & IIf(strǩ�� <> "", strǩ��, str������) & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    gcnOracle.BeginTrans        '----------������״̬��������
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������״̬��������")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If intState = 4 Or intState = 5 Then
        If Not mblnIgnoreResult And Not bln���������� Then  '�ڱ���ʱ��ʾ���������
            If mblnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
                gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
                zlDatabase.ExecuteProcedure gstrSQL, "���������"
            ElseIf mintResultInput = 1 Then
                Call PromptResult(lngOrderID, mlngModul, frmOwnerForm)  ' Me)
            End If
        End If
    End If
    
    If intState = 5 And mblnCompleteCommit Then   '�������˺�ֱ����ɡ�
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    
    '����״̬����
    Call StateCheck(intState)
    Exit Sub
errHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub UpdateStudyListState(lngAdviceID As Long, strStudyUID As String, blnAddImage As Boolean, blnStateChanged As Boolean)
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then Exit Sub
    
    Dim intRowIndex As Integer
    
    For intRowIndex = 0 To vsList.Rows - 1
        If vsList.TextMatrix(intRowIndex, GetCN("ҽ��ID")) = CStr(lngAdviceID) Then
            Exit For
        End If
    Next intRowIndex
    
    If blnStateChanged Then
        If blnAddImage Then '��ͼ
            vsList.TextMatrix(intRowIndex, GetCN("���UID")) = Nvl(strStudyUID, "A123456789")
            Set vsList.Cell(flexcpPicture, intRowIndex, GetCN("����")) = Imglist.ListImages("Ӱ��").Picture '�ı�ͼ��
        Else '���һ�β�ͼ
            vsList.TextMatrix(intRowIndex, GetCN("���UID")) = ""
            Set vsList.Cell(flexcpPicture, intRowIndex, GetCN("����")) = Nothing '�ı�ͼ��
        End If
    End If
    
    '�������ø���Ӱ���鼼ʦ
    If mblnWriteCapDoctor = True And blnStateChanged = True Then
        gstrSQL = "Zl_Ӱ����_��鼼ʦ( " & vsList.TextMatrix(intRowIndex, GetCN("ҽ��ID")) & "," & vsList.TextMatrix(intRowIndex, GetCN("���ͺ�")) & ",'" & IIf(blnAddImage = True, UserInfo.����, "") & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
End Sub


Private Sub StateCheck(ByVal intState As Integer, Optional ByVal lngAdviceID As Long)
    
    If mblnPatTrack Then
        Select Case intState '���ݲ�����״̬ȷ����״̬�����Ƿ�ѡ��
            Case 0, 1
                If Not mblncmd�Ǽ� Then mblncmd�Ǽ� = True
            Case 2
                If Not mblncmd���� Then mblncmd���� = True
            Case 3
                If Not mblncmd��� Then mblncmd��� = True
            Case 4
                If Not mblncmd���� Then mblncmd���� = True
            Case 5
                If Not mblncmd��� Then mblncmd��� = True
            Case 6
                If Not mblncmd��� Then mblncmd��� = True
        End Select
        Call RefreshList(lngAdviceID)
    Else '������ֻˢ���б�
        Call RefreshList
    End If
End Sub
Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
        
    On Error GoTo errH
    
    objPopup.CommandBar.Controls.DeleteAll
    With vsList
        gstrSQL = "Select Distinct C.���,C.����,C.˵��" & _
            " From ����ҽ����¼ A,��������Ӧ�� B,�����ļ��б� C" & _
            " Where A.ID=[1] And A.���ID IS NULL" & _
            " And A.������ĿID=B.������ĿID" & _
            " And B.Ӧ�ó���=[2] And B.�����ļ�ID=C.ID And C.����=7" & _
            " Order by C.���"
        If .TextMatrix(.Row, GetCN("ת��")) = 1 Then
            gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.TextMatrix(.Row, GetCN("ҽ��ID"))), CLng(Decode(.TextMatrix(.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4)))
    End With
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!���� & "(&0)")
            objControl.parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
    On Error GoTo errH
    If objControl.parameter = "" Then '��֣�ֱ�Ӱ�F10ʱ����һ���յ�Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.parameter = "" Then Exit Sub
    
    If ReportPrintSet(gcnOracle, glngSys, objControl.parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.parameter, Me, "NO=" & vsList.TextMatrix(vsList.Row, GetCN("NO")), _
                        "����=" & vsList.TextMatrix(vsList.Row, GetCN("��¼����")), "ҽ��ID=" & vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")), 1)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshList(Optional ByVal lngAdviceID As Long = 0)
Dim i As Integer, lngcurҽ��ID As Long, lngRow As Long, lngTopRow As Long
    With vsList
        If lngAdviceID <> 0 Then
            lngcurҽ��ID = lngAdviceID
        Else
            lngcurҽ��ID = Val(.TextMatrix(.Row, GetCN("ҽ��ID"))) '��ǰ��ҽ��ID
            lngRow = .Row: lngTopRow = .TopRow               '��ǰ�кͶ���֮��Ĳ��
        End If
        
        Call LoadPatiList
        
        If lngcurҽ��ID = 0 Then
            Call .Select(1, GetCN("����"))
            Exit Sub
        End If
        
        '�м�¼ʱҪ���¶�λ��֮ǰ��¼
        On Error Resume Next
        lngcurҽ��ID = .FindRow(CStr(lngcurҽ��ID), , GetCN("ҽ��ID"))
        If lngcurҽ��ID <> -1 Then
            lngRow = Abs(lngRow - lngTopRow)
            If .Row = lngcurҽ��ID Then '��ͬʱ���ᴥ��CHANGE�¼�
                Call vsList_RowColChange 'ǿ��ˢ���ұ��Ӵ���
            Else
                .Row = lngcurҽ��ID
            End If
            .TopRow = .Row - lngRow
        Else
            If .Row <> 1 Then
                .Row = 1
            Else
                Call vsList_RowColChange 'ǿ��ˢ���ұ��Ӵ���
            End If
        End If
        err.Clear
    End With
End Sub

Private Sub mobjSysHook_OnHookProcess(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case wParam
        Case 119
            '�жϼ��̰����Ƿ��ɿ���Ϊ0��ʾ���¼���
            If (lParam And &H80000000) = 0 Then
                Exit Sub
            End If
                        
            If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then
                'ִ�п�ݲɼ�
                Call frmVideoCapture.CaptureImage
            End If
    End Select
End Sub



Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraInfo.Top = -75
    fraInfo.Left = fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left
    
    lblCash.Top = (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 60
    
    
    lbl������Ϣ.Width = lblCash.Left
    lbl�����Ϣ.Width = lblCash.Left
    
    lbl������Ϣ.Left = 60
    lbl�����Ϣ.Left = 60
End Sub

Private Function GetFilterData() As ADODB.Recordset
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
Dim strSQLBak As String
Dim str��Դ As String
Dim strFilter As String
Dim i As Integer
Dim strModalitys As String
Dim blnUseTime As Boolean       '�Ƿ�ʹ��ʱ������
Dim strTemp As String
Dim strLinkTab As String

    
    On Error GoTo errHandle
    
    Set GetFilterData = Nothing
    
    With SQLCondition
        blnUseTime = False  'Ĭ�ϲ�ʹ��ʱ������
        '�������������ʹ��ʱ������
        If .����� <> 0 Then
            strFilter = " And C.�����=[1]"
        ElseIf .סԺ�� <> 0 Then
            strFilter = " And C.סԺ��=[2]"
        ElseIf .���￨ <> "" Then
            strFilter = " And C.���￨��=[3]"
        ElseIf .���� <> "" And InStr(.����, "*") = 0 Then   '�������⴦����*�ű�ʾģ����ѯ
            strFilter = " And C.����=[4]"
        ElseIf .���֤ <> "" Then
            strFilter = " And C.���֤��=[5]"
        ElseIf .IC�� <> "" Then
            strFilter = " And C.IC��=[6]"
        ElseIf .���ݺ� <> "" Then
            strFilter = " And A.NO=[7] "
        ElseIf .���� <> 0 Then
            strFilter = " And H.����=[8] "
        ElseIf .����� <> "" Then
            strFilter = " And o.�����=[9] "
        Else
        '����������ѯ��ʹ��ʱ������
            blnUseTime = True
            '��д����ʱ������
            'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
            If .ʱ������ = 1 Then       '������ʱ��
                strFilter = " And A.����ʱ�� Between [10] and "
            ElseIf .ʱ������ = 2 Then   '������ʱ��
                strFilter = " And A.�״�ʱ�� Between [10] and "
            Else                        '��ͼʱ��
                strFilter = " And H.�������� Between [10] and "
            End If
            If .����ʱ�� <> CDate(0) Then
                strFilter = strFilter & " [11] "
            Else
                strFilter = strFilter & " Sysdate+1/(24*3600) "
            End If
            
            '�ȴ��������д�*�ŵģ����д�ʱ��������ģ����ѯ
            If .���� <> "" And InStr(.����, "*") <> 0 Then
                .���� = Replace(.����, "*", "%")
                strFilter = strFilter & " And C.���� like [4]"
            End If
            
            If .�Ա� <> "" Then
                strFilter = strFilter & " And Nvl(H.�Ա�,C.�Ա�)=[30]"
            End If
        
        
            '��������-��ʼ����(ֻ�е�����ʹ�á����������ڶ�������֮��ʱ����ʹ�ÿ�ʼ����)
            If .��ʼ���� <> -1 Then
                If .�������� = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.����)>=[31]"
                End If
            End If
            
            '��������-��������
            If .�������� <> -1 Then
                If .�������� = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.����)<=[32]"
                Else
                    strFilter = strFilter & " And ZL_AgeToDays(C.����)" & .�������� & "[32]"
                End If
            End If
            
            If .���˿��� <> 0 Then
                strFilter = strFilter & " And B.���˿���ID+0=[12] "
            End If
        
            If .�걾��λ <> "" Then
                strFilter = strFilter & " And instr(B.ҽ������,[13])>0"
            End If
            
            If .������� <> -1 Then
                strFilter = strFilter & " And Nvl(A.�������, 0)=[29]"
            End If
            
            If .���ҽ�� <> "" Then
                strFilter = strFilter & " And H.������=[14] "
            End If
            
            If .���ҽ�� <> "" Then
                strFilter = strFilter & " And H.������=[15] "
            End If
            
            If .Ӱ������ <> "" Then
                strFilter = strFilter & " And H.Ӱ������=[16]"
            End If
            
            If .��鼼ʦ <> "" Then
                strFilter = strFilter & " And H.��鼼ʦ=[17]"
            End If
            
            'Ӱ������������ط�������������ѡ�񣬹��˴��ں����������棬���������е�Ϊ��
'            If mintcmdӰ����� > 0 Then
'                Dim objControl As CommandBarControl
'
'                Set objControl = cbrdock.FindControl(, ID_Ӱ�����)
'                For i = 1 To objControl.CommandBar.Controls.Count
'                    If objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Checked = True Then
'                        strModalitys = strModalitys & "," & objControl.CommandBar.FindControl(, ID_Ӱ����� + i).DescriptionText
'                    End If
'                Next i
'                If strModalitys <> "" Then
'                    strFilter = strFilter & " And instr([27],H.Ӱ�����)>0 "
'                End If
'            Else
'                If .Ӱ����� <> "" Then
'                    strFilter = strFilter & " And H.Ӱ�����=[18] "
'                End If
'            End If
            
            If .��� <> "" Then
                strFilter = strFilter & " And  Instr(H.�������, [19]) > 0 "
            End If
            
            If .������� <> "" Then
                strFilter = strFilter & " And B.ID IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id IN " & _
                                                                        " (Select Distinct A.ID  " & _
                                                                        "From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                        "Where A.����ʱ��>[1] AND A.Id=B.�ļ�ID  " & _
                                                                            "And B.��������=7 And instr(B.��������,'52;')>0 And instr(B.�����ı�,[20])>0))"
            End If
            
            Dim strSubFilter As String '����PACS�����������
            If .������� <> "" Then
                strSubFilter = " (b.�����ı� ='�������' And Instr(c.�����ı�, [21]) > 0)"
            End If
            
            If .������ <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.�����ı� ='������' And Instr(c.�����ı�, [22]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.�����ı� ='������' And Instr(c.�����ı�, [22]) > 0)"
                End If
            End If
            
            If .���� <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.�����ı� ='����' And Instr(c.�����ı�, [23]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.�����ı� ='����' And Instr(c.�����ı�, [23]) > 0)"
                End If
            End If
            
            If strSubFilter <> "" Then
                strSubFilter = " (" & strSubFilter & ")"
                strFilter = strFilter & " And B.ID IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id IN " _
                    & " (Select Distinct a.ID From ���Ӳ�����¼ a, ���Ӳ������� b,���Ӳ������� c " _
                    & " Where a.����ʱ�� > [10] And a.Id = b.�ļ�id And b.Id = C.��ID And b.�������� = 3 And c.�������� = 2 And c.��ֹ�� = 0 and " _
                    & strSubFilter & "))"
            End If
        End If
        
        '�����˴��ڡ��͡�������ҡ������������������������ʹ��ʱ����������������Ϊ��������
        '������Դ (1-����,2-סԺ,3-����,4-���)
        '���������Դ��ѡ���ˣ���ʾ�������в��ˣ�����Ӳ�����Դ�Ĳ�ѯ����
        If mblncmd���� And mblncmdסԺ And mblncmd��� And mblncmd���� Then
        
        Else
            If mblncmd���� Then str��Դ = "1,"
            If mblncmdסԺ Then str��Դ = str��Դ & "2,"
            If mblncmd���� Then str��Դ = str��Դ & "3,"
            If mblncmd��� Then str��Դ = str��Դ & "4,"
            If str��Դ <> "" Then       'str��ԴΪ�գ���ʾû��ѡ���κ���Դ������Ӳ�����Դ�Ĳ�ѯ����
                str��Դ = Mid(str��Դ, 1, Len(str��Դ) - 1)
                strFilter = strFilter & " And Instr([24],B.������Դ)> 0"
            End If
        End If
        
    
        If mstrRoom <> "" Then  'ֻ��ʾִ�м䷶Χ�ڵ�
            If Not mblncmd�Ǽ� Then
                strFilter = strFilter & " And Instr([25],','|| A.ִ�м� || ',' )>0"
            Else
                strFilter = strFilter & " And (Instr([25],','|| A.ִ�м� || ',' )>0 And Nvl(A.ִ�й���,0)>1 OR Nvl(A.ִ�й���,0)<2)"
            End If
        End If
    
        If mblnNoShowCancel Then '����ʾȡ���Ǽǵļ��
            strFilter = strFilter & " And A.ִ��״̬<>2 "
        End If
        
        If mblncmd���� Then        'ֻ��ʾ����סԺ��¼
            strFilter = strFilter & vbNewLine & " And (B.������Դ=2 And B.��ҳID=C.סԺ���� Or Nvl(B.������Դ,0)<>2)"
        End If
        
        
        
        
        '����ָ���Ĳ��������ͽ��й���
        If mblncmd���� Or mblncmd���� Or mblncmdϸ�� Or mblncmd���� Or mblncmdʬ�� Then
            strTemp = ""
            
            If mblncmd���� Then
                strTemp = strTemp & vbNewLine & " o.�������=0"
            End If
            
            If mblncmd���� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " o.�������=1"
            End If
            
            If mblncmdϸ�� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " o.�������=2"
            End If
            
            If mblncmd���� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " o.�������=3"
            End If
            
            If mblncmdʬ�� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " o.�������=4"
            End If
            
            If Trim(strTemp) <> "" Then strFilter = strFilter & vbNewLine & " and (" & strTemp & " ) "
        End If
        
      
        
        
        
        '���ݱ걾���ͽ��й���
        If mblncmd���� Or mblncmdС�걾 Or mblncmd���� Or mblncmd���� Or mblncmdҺ�� Then
            strTemp = ""
            
            strLinkTab = strLinkTab & " ����걾��Ϣ p"
            
            If mblncmd���� Then
                strTemp = strTemp & " p.�걾����=0"
            End If
            
            If mblncmdС�걾 Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " p.�걾����=1"
            End If
            
            If mblncmd���� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " p.�걾����=2"
            End If
            
            If mblncmd���� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " p.�걾����=3"
            End If
            
            If mblncmdҺ�� Then
                If strTemp <> "" Then strTemp = strTemp & " or "
                strTemp = strTemp & " p.�걾����=4"
            End If
            
            If Trim(strTemp) <> "" Then strFilter = strFilter & vbNewLine & " and a.ҽ��ID=p.ҽ��ID and ( " & strTemp & " ) "
        End If
        
        
        '���˵�ǰҳ������
        If tabFilter.Tag Then
            Select Case tabFilter.Selected.Tag
                Case "��ȡ��"
                    strFilter = strFilter & " and (o.��ǰ���� = 1 or o.��ǰ���� = 8)"
                Case "��ȡ��"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " ����ȡ����Ϣ q"
                    
                    strFilter = strFilter & " and o.����� = q.�����"
                    
                Case "����Ƭ"
                    strFilter = strFilter & " and (o.��ǰ���� = 2 or o.��ǰ���� = 9)"
                Case "����Ƭ"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " ������Ƭ��Ϣ r"
                    
                    strFilter = strFilter & " and (o.�����=r.����� and r.��ǰ״̬=2)"
                    
                Case "��Ƭ����"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " ������Ƭ��Ϣ r"
                    
                    strFilter = strFilter & " and (o.�����=r.����� and r.��ǰ״̬=1)"
                    
                Case "������"
                    strFilter = strFilter & " and (o.��ǰ���� = 4)"
                Case "������"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " �����ؼ���Ϣ s"
                    
                    strFilter = strFilter & " and (o.�����=s.����� and s.�ؼ�����=0 and s.��ǰ״̬=2)"
                    
                Case "���߽���"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " �����ؼ���Ϣ s"
                    
                    strFilter = strFilter & " and (o.�����=s.����� and s.�ؼ�����=0 and s.��ǰ״̬=1)"
                    
                Case "����Ⱦ"
                    strFilter = strFilter & " and (o.��ǰ���� = 5)"
                Case "����Ⱦ"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " �����ؼ���Ϣ s"
                    
                    strFilter = strFilter & " and (o.�����=s.����� and s.�ؼ�����=1 and s.��ǰ״̬=2)"
                    
                Case "��Ⱦ����"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " �����ؼ���Ϣ s"
                    
                    strFilter = strFilter & " and (o.�����=s.����� and s.�ؼ�����=1 and s.��ǰ״̬=1)"
                    
                Case "�����"
                    strFilter = strFilter & " and (o.��ǰ���� = 6)"
                Case "�ѷ���"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " �����ؼ���Ϣ s"
                    
                    strFilter = strFilter & " and (o.�����=s.����� and s.�ؼ�����=2 and s.��ǰ״̬=2)"
                    
                Case "���ӽ���"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " �����ؼ���Ϣ s"
                    
                    strFilter = strFilter & " and (o.�����=s.����� and s.�ؼ�����=2 and s.��ǰ״̬=1)"
                    
                Case "���ڻ���"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " ���������Ϣ t"
                    
                    strFilter = strFilter & " and (o.�����=t.����� and t.��ǰ״̬=0 and t.����ҽʦ='" & UserInfo.���� & "')"
                    
                Case "�ѻ���"
                    If Trim(strLinkTab) <> "" Then strLinkTab = strLinkTab & ","
                    strLinkTab = strLinkTab & " ���������Ϣ t"
                    
                    strFilter = strFilter & " and (o.�����=t.����� and t.��ǰ״̬<>0 and t.����ҽʦ='" & UserInfo.���� & "')"
                    
                Case "�� ��"
            End Select
        End If
        
        
        
        '������������
        If .�������� <> "" Then
            strFilter = strFilter & " And B.id IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id In " & _
                                                                    " (Select Distinct A.ID " & _
                                                                    " From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                    " Where A.����ʱ��>[10] AND A.Id=B.�ļ�ID " & _
                                                                    " And B.��������=2 And instr(B.�����ı�,[28])>0 And B.��ֹ�� = 0)) "
        End If
        
        gstrSQL = "Select /*+ RULE */ Distinct" & vbNewLine & _
                    "       A.ҽ��ID,B.���ID,A.���ͺ�,A.�״�ʱ�� ����ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.������� ����," & vbNewLine & _
                    "       decode(o.��ǰ����,1,'ȡ��',2,'��Ƭ',3,'���',4,'�����黯',5,'����Ⱦɫ',6,'���Ӳ���',8,'��ȡ��',9,'����Ƭ',10,'���',null) as ����ִ�й���, " & vbNewLine & _
                    "       decode(o.�������,0,'����',1,'����',2,'ϸ��',3,'����',4,'ʬ��',null) as  ������, " & vbNewLine & _
                    "       decode(o.�����,null,'δ����','�Ѻ���') as �������, " & vbNewLine & _
                    "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,Decode(B.������Դ, 1, '��', 2, 'ס', 3, '��', 4, '��') ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
                    "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��," & vbNewLine & _
                    "       Nvl(H.����,C.����) ����,H.����,Nvl(H.�Ա�,C.�Ա�) �Ա�,Nvl(H.����,C.����) ����,H.���,H.����,H.Ӱ������," & vbNewLine & _
                    "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������,o.�����,H.���淢��,H.����ID,A.��¼����, " & vbNewLine & _
                    "       H.�����,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.������,H.��鼼ʦ,H.�������� ��ͼʱ��, " & vbNewLine & _
                    "       H.�������,H.��Ϸ���,H.���UID,0 as ת��,F.���� AS ���˿���, " & vbNewLine & _
                    "       C.���￨��,A.NO as ���ݺ�,C.���֤��,D.״̬ as ·��״̬,A.�Ʒ�״̬,Decode(A.��¼����,2,1,Decode(a.�Ʒ�״̬,3,1,0)) as �շ� " & vbNewLine & _
                    " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,�����ٴ�·�� D,Ӱ�����¼ H,Ӱ������Ŀ G,���ű� F, " & vbNewLine & _
                    "       ��������Ϣ o " & IIf(Trim(strLinkTab) <> "", ",", "") & strLinkTab & vbNewLine & _
                    " Where A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) " & vbNewLine & _
                    "       And B.������ĿID=G.������ĿID And B.����ID=C.����ID And B.���˿���id=F.ID " & vbNewLine & _
                    "       and A.ҽ��ID=o.ҽ��ID(+) " & vbNewLine & _
                    "       And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) And D.����ʱ��(+) is Null "
        gstrSQL = gstrSQL & vbNewLine & strFilter & " And A.ִ�в���ID+0=[26]"
        
        'ͨ��"����ҽ������.�Ʒ�״̬"ֱ���ж�,ԭ��ֵ��-1-����Ʒ�;0-δ�Ʒ�;1-�ѼƷѣ����ڼ��ʵ�������������ʵ���������ԭ��ֵ���䡣
        '�����շѵ��ķ��ͼ�¼����������״̬��2-�����շѣ�3-ȫ���շ�
'        If mblncmd�ѽ� = True Then
'            gstrSQL = gstrSQL & " and (A.��¼���� <> 1 Or (A.��¼���� = 1 And a.�Ʒ�״̬ = 3)) "
'        ElseIf mblncmdδ�� = True Then
'            gstrSQL = gstrSQL & " and (A.��¼���� = 1 And A.�Ʒ�״̬ <>3) "
'        End If
        
        '��ʹ�ü��Ż���Ų���ʱһ���Ǳ������ģ�Ӱ�����¼���м�¼����ʱȡ�������ӱ���ȫ��ɨ��
        'ʹ�òɼ�ʱ����ˣ�Ӱ�����¼���м�¼
        If .���� <> 0 Or .����� <> "" Or (blnUseTime = True And SQLCondition.ʱ������ = 3) Then
            gstrSQL = Replace(Replace(gstrSQL, "H.ҽ��ID(+)", "H.ҽ��ID"), "H.���ͺ�(+)", "H.���ͺ�")
            If .����� <> "" Then
                gstrSQL = Replace(gstrSQL, "I.ҽ��ID(+)", "I.ҽ��ID")
            End If
        End If

        '���������ת����Ҫ�����󱸱�
        If mblnMoved Then
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
            strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
            strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")

            strSQLBak = Replace(strSQLBak, "���Ӳ�����¼", "H���Ӳ�����¼")
            strSQLBak = Replace(strSQLBak, "���Ӳ�������", "H���Ӳ�������")
            strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
            gstrSQL = gstrSQL & " Union ALL " & strSQLBak
        End If
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order by ������,����ʱ��,����ʱ��"
    
        Set GetFilterData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����б�", .�����, .סԺ��, .���￨, .����, .���֤, _
                                            .IC��, .���ݺ�, .����, .�����, .��ʼʱ��, .����ʱ��, .���˿���, _
                                            .�걾��λ, .���ҽ��, .���ҽ��, .Ӱ������, .��鼼ʦ, "", .���, _
                                            .�������, .�������, .������, .����, str��Դ, mstrRoom, mlngCur����ID, _
                                            strModalitys, .��������, .�������, .�Ա�, .��ʼ����, .��������)
    End With
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub LoadPatiList()
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
Dim rsList As ADODB.Recordset
Dim strFilter As String

    If Not mblnInitOk Then Exit Sub      '��ʼ��δ���
    mblnvsRefresh = True
    On Error GoTo errHandle
    
    Set rsList = GetFilterData()
   
    strFilter = ""
    If mblncmd�Ǽ� Then strFilter = "������=0 or ������=1 or "
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=2 or ", "������=2 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=3 or ", "������=3 or ")
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=4 or ", "������=4 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=5 or ", "������=5 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=6 or ", "������=6 or ")
    If mblncmd�Ǽ� And mblncmd���� And mblncmd��� And mblncmd���� And mblncmd��� And mblncmd��� Then
        strFilter = ""
    End If
    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
        rsList.Filter = strFilter
    End If
    
    Call FillList(vsList, rsList)
    
    stbThis.Panels(2).Text = "�� " & vsList.Rows - 1 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
    
    mblnvsRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Function OpenPatiListWind(ByRef lngAdviceID As Long, ByRef strPatientName As String) As Boolean
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
'���أ�����ѡ���ҽ��ID
Dim rsList As ADODB.Recordset
Dim strFilter As String

    On Error GoTo errHandle
    
    lngAdviceID = -1
    strPatientName = ""
    OpenPatiListWind = False
    
    Set rsList = GetFilterData()

    strFilter = ""
'    If mblncmd�Ǽ� Then strFilter = "������=0 or ������=1 or "
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=2 or ", "������=2 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=3 or ", "������=3 or ")
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=4 or ", "������=4 or ")
    
    If strFilter = "" Then strFilter = "������=2 or ������=3 or ������=4 or "
    
'    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=5 or ", "������=5 or ")
'    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=6 or ", "������=6 or ")
'    If mblncmd�Ǽ� And mblncmd���� And mblncmd��� And mblncmd���� And mblncmd��� And mblncmd��� Then
'        strFilter = ""
'    End If

    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
        rsList.Filter = strFilter
    End If
    
    Call FillList(frmOpenStudyList.vsStudyList, rsList)


    frmOpenStudyList.Show 1
    
    If frmOpenStudyList.blnOK Then
        lngAdviceID = Val(Nvl(frmOpenStudyList.vsStudyList.TextMatrix(frmOpenStudyList.vsStudyList.Row, GetCN("ҽ��ID")), 0))
        strPatientName = Nvl(frmOpenStudyList.vsStudyList.TextMatrix(frmOpenStudyList.vsStudyList.Row, GetCN("����")), "")
    Else
        lngAdviceID = -1
    End If
    
    frmOpenStudyList.blnOK = False
    
    OpenPatiListWind = IIf(lngAdviceID <= 0, False, True)
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub FillList(lst As VSFlexGrid, ByVal rsTemp As ADODB.Recordset)
    Dim rsBaby As ADODB.Recordset
    Dim blnShowPath As Boolean      '�Ƿ���ʾ·����
    Dim intPathColNum As Integer
    Dim rsClone As New ADODB.Recordset
    Dim rsList As New ADODB.Recordset
    Dim intRow As Integer
    Dim blnCharged As Boolean
    Dim i As Integer
    
    On Error GoTo errHandle
    Call InitList(lst)
    
    If rsTemp.EOF Then stbThis.Panels(2).Text = "û���ҵ��κ�ƥ��ļ�¼": Exit Sub
    
    Set rsList = rsTemp.Clone
    Set rsClone = rsTemp.Clone
    
    '����rsList����������
    rsList.Filter = rsTemp.Filter
    
    intRow = 1
    
    With lst
        Do Until rsList.EOF
            
            blnCharged = True
            
            '�ж��Ƿ��Ѿ��շ�
            '"����ҽ������.��¼����"--- 1���շѵģ�2�Ǽ��ʵġ�

            'ͨ��"����ҽ������.�Ʒ�״̬"ֱ���ж�,ԭ��ֵ��-1-����Ʒ�;0-δ�Ʒ�;1-�ѼƷѣ����ڼ��ʵ�������������ʵ���������ԭ��ֵ���䡣
            '�����շѵ��ķ��ͼ�¼����������״̬��2-�����շѣ�3-ȫ���շ�
            
            'û�ж�Ӧ���õ�ҽ�������������һ����"-1-����Ʒ�"����û�������շѶ��գ�һ����"0-δ�Ʒ�"������Ȼ�������շѶ��գ�������Ϊ���ͺ��ֹ��Ʒѣ�����ҽ������ȥ���ɡ�
            '"1-�ѼƷ�"���Ƿ���ʱ�����˷��õġ��������˷��õ��ݲ���ʾ�շ��ˣ����ɿ����Ǽ��ʻ��۵������շѻ��۵��������շѻ��۵��Ͷ�����״̬��
            '"2-�����շ�"��ʾ�����շѺͲ����˷ѵ����������û�յ��ꡣ
            
            '���շ��ж�������
            '1����ҽ���Ǽ��˵����շ�-------����¼����=2��
            '2����ҽ�����շѵ��ģ����������������շ�
            '   (1)��ҽ���Ͳ�λҽ���� �Ʒ�״̬in(-1,0,3)���շ�-----����¼����=1 and �Ʒ�״̬in(-1,0,3)��
            
            If Nvl(rsList!���ID) = "" Then
                If Nvl(rsList!��¼����, 2) = 2 Then
                    blnCharged = True
                Else
                    If Nvl(rsList!�Ʒ�״̬, -1) = 1 Or Nvl(rsList!�Ʒ�״̬, -1) = 2 Then
                        blnCharged = False
                    Else
                        '��ѯ��ҽ��δ�Ʒѻ����Ѿ��շ��ˣ���Ҫ�鲿λҽ�����շ����������ҽ�����Ѿ��շѣ��������շ�
                        rsClone.Filter = "���ID = " & Nvl(rsList!ҽ��ID)
                        Do While rsClone.EOF = False
                            If Nvl(rsClone!�Ʒ�״̬, -1) = 1 Or Nvl(rsClone!�Ʒ�״̬, -1) = 2 Then
                                blnCharged = False
                                Exit Do
                            End If
                            rsClone.MoveNext
                        Loop
                    End If
                End If
            End If
            
            If Nvl(rsList!���ID) = "" And ((mblncmd�ѽ� = True And blnCharged = True) Or (mblncmdδ�� = True And blnCharged = False) _
                Or (mblncmd�ѽ� = False And mblncmdδ�� = False)) Then
                '�����շ�������շѹ���������ȷ���Ƿ���ӵ��б���
                
                .Rows = intRow + 1
                .Row = intRow
                intRow = intRow + 1
            
                If Nvl(rsList!·��״̬, 0) = 1 Then
                   Set .Cell(flexcpPicture, .Row, GetCN("·��")) = Imglist.ListImages("·��").Picture
                   .TextMatrix(.Row, GetCN("·��")) = " "
                   blnShowPath = True
                End If
                
                .Cell(flexcpData, .Row, GetCN("����")) = Val(rsList!������־)
                If rsList!������־ <> 0 Then
                    Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("����").Picture
                End If
                If rsList!��Դ = "ס" Then
                    Set .Cell(flexcpPicture, .Row, GetCN("��Դ")) = Imglist.ListImages("סԺ").Picture
                End If
                .TextMatrix(.Row, GetCN("��Դ")) = rsList!��Դ
                .Cell(flexcpData, .Row, GetCN("��Դ")) = Decode(rsList!��Դ, "��", 1, "ס", 2, "��", 3, 4)
                
                If blnCharged = True Then
                    Set .Cell(flexcpPicture, .Row, GetCN("�շ�")) = Imglist.ListImages("�շ�").Picture
                    .TextMatrix(.Row, GetCN("�շ�")) = " "  ' ��������
                End If
                
                If Nvl(rsList!����, 0) <> 0 Then
                    Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("����").Picture
                    .TextMatrix(.Row, GetCN("����")) = " "  ' ��������
                End If
                
                If Nvl(rsList!��ɫͨ��, 0) <> 0 Then
                    Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("��ɫͨ��").Picture
                End If
                
                If Nvl(rsList!���uid) <> "" Then
                    Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("Ӱ��").Picture
                End If
                .TextMatrix(.Row, GetCN("����")) = Nvl(rsList!Ӱ������)
                .TextMatrix(.Row, GetCN("����")) = Nvl(rsList!����)
                .TextMatrix(.Row, GetCN("�����")) = Nvl(rsList!�����)
                .TextMatrix(.Row, GetCN("������")) = IIf(rsList!ִ��״̬ = 2, "�Ѿܾ�", Decode(Nvl(rsList!������, 0), 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                                                            2, IIf(Nvl(rsList!�������) <> "", "������", _
                                                                                                    IIf(Nvl(rsList!������) = "", "�ѱ���", "������")), _
                                                                                            3, IIf(Nvl(rsList!�������) <> "", "������", _
                                                                                                    IIf(Nvl(rsList!������) = "", "�Ѽ��", "������")), _
                                                                                            4, IIf(Nvl(rsList!�������) <> "", "������", _
                                                                                                    IIf(Nvl(rsList!������) <> "", "�����", "�ѱ���")), _
                                                                                            5, "�����", "�����"))
                .TextMatrix(.Row, GetCN("�Ա�")) = Nvl(rsList!�Ա�)
                .TextMatrix(.Row, GetCN("����")) = Nvl(rsList!����)
                If InStr(Nvl(rsList!ҽ������), ":") > 0 Then '�µ�ģʽ������ҽ����������Ϣ�� ����,ִ�б��:��λ(����,����),��λ---
                    .TextMatrix(.Row, GetCN("ҽ������")) = Split(rsList!ҽ������, ":")(0)
                    .TextMatrix(.Row, GetCN("��λ����")) = Split(rsList!ҽ������, ":")(1)
                Else
                    .TextMatrix(.Row, GetCN("ҽ������")) = Nvl(rsList!ҽ������)
                End If
                .TextMatrix(.Row, GetCN("����ʱ��")) = Nvl(rsList!����ʱ��)
                .TextMatrix(.Row, GetCN("����ִ�й���")) = Nvl(rsList!����ִ�й���)
                .TextMatrix(.Row, GetCN("����ʱ��")) = Nvl(rsList!����ʱ��)
                .TextMatrix(.Row, GetCN("����ҽ��")) = Nvl(rsList!����ҽ��)
                .TextMatrix(.Row, GetCN("���")) = Nvl(rsList!���)
                .TextMatrix(.Row, GetCN("����")) = Nvl(rsList!����)
                .TextMatrix(.Row, GetCN("Ӥ��")) = Nvl(rsList!Ӥ��)
                .TextMatrix(.Row, GetCN("�Ǽ���")) = Nvl(rsList!�Ǽ���)
                .TextMatrix(.Row, GetCN("������")) = Nvl(rsList!������)
                .TextMatrix(.Row, GetCN("�����")) = Nvl(rsList!�����)
                .TextMatrix(.Row, GetCN("�������")) = Nvl(rsList!�������)
                .TextMatrix(.Row, GetCN("��ɫͨ��")) = Nvl(rsList!��ɫͨ��)
                .TextMatrix(.Row, GetCN("�����ӡ")) = IIf(Nvl(rsList!�����ӡ) = 1, "�Ѵ�ӡ", "δ��ӡ")
                .TextMatrix(.Row, GetCN("������")) = Nvl(rsList!������)
                .TextMatrix(.Row, GetCN("������")) = Nvl(rsList!������)
                .TextMatrix(.Row, GetCN("��ͼʱ��")) = Nvl(rsList!��ͼʱ��)
                .TextMatrix(.Row, GetCN("����")) = Nvl(rsList!����)
                .TextMatrix(.Row, GetCN("������")) = Nvl(rsList!������)
                .TextMatrix(.Row, GetCN("�������")) = Nvl(rsList!�������) ' Decode(Nvl(rsList!�������, "δ����"), "�Ѻ���", "��", "")
                .TextMatrix(.Row, GetCN("����ID")) = Nvl(rsList!����ID, 0)
                .TextMatrix(.Row, GetCN("��ҳID")) = Nvl(rsList!��ҳID, 0)
                .TextMatrix(.Row, GetCN("�Һŵ�")) = Nvl(rsList!�Һŵ�)
                .TextMatrix(.Row, GetCN("���˿���ID")) = Nvl(rsList!���˿���ID, 0)
                .TextMatrix(.Row, GetCN("ҽ��ID")) = Nvl(rsList!ҽ��ID)
                .TextMatrix(.Row, GetCN("���ͺ�")) = Nvl(rsList!���ͺ�)
                .TextMatrix(.Row, GetCN("���UID")) = Nvl(rsList!���uid)
                .TextMatrix(.Row, GetCN("���״̬")) = Nvl(rsList!������)
                .TextMatrix(.Row, GetCN("�������")) = Nvl(rsList!�������)
                .TextMatrix(.Row, GetCN("NO")) = Nvl(rsList!NO)
                .TextMatrix(.Row, GetCN("��¼����")) = Nvl(rsList!��¼����)
                .TextMatrix(.Row, GetCN("ת��")) = Nvl(rsList!ת��)
                .TextMatrix(.Row, GetCN("����")) = Nvl(rsList!��ǰ����)
                .TextMatrix(.Row, GetCN("��ǰ����ID")) = Nvl(rsList!��ǰ����ID, 0)
                .TextMatrix(.Row, GetCN("��ʶ��")) = Nvl(rsList!��ʶ��)
                .TextMatrix(.Row, GetCN("���淢��")) = IIf(Nvl(rsList!���淢��, 0) = 0, "δ����", "�ѷ���")
                .TextMatrix(.Row, GetCN("��Ϸ���")) = Nvl(rsList!��Ϸ���)
                .TextMatrix(.Row, GetCN("����ID")) = Nvl(rsList!����ID, 0)
                .TextMatrix(.Row, GetCN("���˿���")) = Nvl(rsList!���˿���)
                .TextMatrix(.Row, GetCN("���￨��")) = Nvl(rsList!���￨��)
                .TextMatrix(.Row, GetCN("���ݺ�")) = Nvl(rsList!���ݺ�)
                .TextMatrix(.Row, GetCN("���֤��")) = Nvl(rsList!���֤��)
                
                If Nvl(rsList!Ӥ��) <> 0 Then
                    gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                                "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                                "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", CLng(rsList!����ID), CLng(Nvl(rsList!��ҳID, 0)), CLng(rsList!Ӥ��))
                    If Not rsBaby.EOF Then
                        .TextMatrix(.Row, GetCN("����")) = rsBaby!Ӥ������
                        .TextMatrix(.Row, GetCN("�Ա�")) = Nvl(rsBaby!Ӥ���Ա�)
                        .TextMatrix(.Row, GetCN("����")) = Nvl(rsBaby!����ʱ��)
                    End If
                End If
    
                If .TextMatrix(.Row, GetCN("������")) = "�Ѿܾ�" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�Ѿܾ�
                If .TextMatrix(.Row, GetCN("������")) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�����
                If .TextMatrix(.Row, GetCN("������")) = "�ѱ���" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�ѱ���
                If .TextMatrix(.Row, GetCN("������")) = "�ѵǼ�" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�ѵǼ�
                If .TextMatrix(.Row, GetCN("������")) = "�Ѽ��" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�Ѽ��
                If .TextMatrix(.Row, GetCN("������")) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�����
                If .TextMatrix(.Row, GetCN("������")) = "������" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor������
                If .TextMatrix(.Row, GetCN("������")) = "������" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor������
                If .TextMatrix(.Row, GetCN("������")) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�����
                If .TextMatrix(.Row, GetCN("������")) = "�ѱ���" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�ѱ���
            End If
            rsList.MoveNext
        Loop
    End With
    
    '���û��·���в��ˣ�����ʾ·����
    intPathColNum = GetCN("·��")
    If blnShowPath = False Then
        vsList.ColWidth(intPathColNum) = 0
    Else
        vsList.ColWidth(intPathColNum) = GetCW("·��")
    End If
    
    '�ָ�����
    If mlngSortCol <> 0 And mintSortOrder <> 0 Then
        If mlngSortCol < lst.Cols Then
            lst.Col = mlngSortCol
            lst.Sort = mintSortOrder
        End If
    End If
    
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Private Sub picVideoContainer_Paint()
    On Error Resume Next
    
    Dim i As Integer
    Dim Count As Integer
    Dim wordRect As RECT
    
    If Not CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Exit Sub
    
    Count = 2
    wordRect.Bottom = 45
    wordRect.Right = 200
    
    If frmVideoCapture.picBackImg.Height * 3 >= picVideoContainer.Height Then Count = 1
    
    Call picVideoContainer.Cls
    For i = 0 To Count
        Call picVideoContainer.PaintPicture(frmVideoCapture.picBackImg.Picture, _
            Round(picVideoContainer.Width / (i + 1)) - frmVideoCapture.picBackImg.Width + 200, _
            Round((picVideoContainer.Height / 3) * (i + 1) - frmVideoCapture.picBackImg.Height), _
            frmVideoCapture.picBackImg.Width, frmVideoCapture.picBackImg.Height)
            
        wordRect.Left = ScaleX(Round(picVideoContainer.Width / (i + 1)) - frmVideoCapture.picBackImg.Width, vbTwips, vbPixels)
        wordRect.Top = ScaleY(Round((picVideoContainer.Height / 3) * (i + 1) - frmVideoCapture.picBackImg.Height), vbTwips, vbPixels) - 25
        
        wordRect.Right = wordRect.Left + 200
        wordRect.Bottom = wordRect.Top + 45
        
        Call DrawText(picVideoContainer.hdc, "��Ƶ�ѱ��������ڴ򿪣�", 24, wordRect, 0)
    Next i
End Sub

Private Sub picVideoContainer_Resize()
    On Error Resume Next
    
    If Not CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Exit Sub
    
    If frmVideoCapture.ParentContainerObj.hWnd = picVideoContainer.hWnd Then
        Call frmVideoCapture.UpdateSize
    End If
End Sub

Private Sub PicWindow_Resize()
    On Error Resume Next
    With picInfo
        .Top = 0
        .Left = 0
        .Width = PicWindow.ScaleWidth
    End With
        
    With TabWindow
        .Top = picInfo.ScaleHeight
        .Left = 0
        .Width = PicWindow.ScaleWidth
        .Height = PicWindow.ScaleHeight - picInfo.ScaleHeight
    End With
End Sub


Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    If tabFilter.ItemCount < 17 Then Exit Sub
    If Not vsList.Visible Then Exit Sub
    
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mblnInitOk Then Exit Sub

    On Error GoTo errHandle
    If mblnIsHistory Then
        RefreshTabWindow mlngHOrderID
    ElseIf Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0 Then
        RefreshTabWindow 0, True
    Else
        RefreshTabWindow 0, False, True
    End If
    
    'ɾ�����ڵĹ������������˵���
    Call LockWindowUpdate(Me.hWnd)
    Dim lngCount As Long
    For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbrMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbrMain.Count To 2 Step -1
        cbrMain(lngCount).Delete
    Next
    Call InitCommandBars
    
    Select Case Item.Tag
        Case "������д"
            If mblnPacsReport = True Then    'ʹ��PACS����༭��
                mfrmPacsReport.zlDefCommandBars Me.cbrMain
            Else
                mobjReport.zlDefCommandBars Me.cbrMain
            End If
        Case "�������"
            mobjExpense.zlDefCommandBars Me, Me.cbrMain
        Case "סԺҽ��"
            mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
        Case "����ҽ��"
            mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
        Case "סԺ����"
            mobjInEPRs.zlDefCommandBars cbrMain
        Case "���ﲡ��"
            mobjOutEPRs.zlDefCommandBars cbrMain
        Case "�Ŷӽк�"
            If Not mobjQueue Is Nothing Then
                mobjQueue.zlDefCommandBars cbrMain
            End If
    End Select
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) <> 0 Then
        '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If
    
    Call LockWindowUpdate(0)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub TimerRefresh_Timer()
    'ˢ�²����б�
    Call RefreshList
End Sub

Private Sub txtFilter_Change()
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (txtFilter.Text = "" And Me.ActiveControl Is txtFilter)
    End If
    If txtFilter.Text = "" Then txtFilter.Tag = ""
    Call subRefreshFilterCondition(txtFilter.Text)
End Sub

Private Sub txtFilter_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '���֤ʶ�����
    
    If txtFilter.Text <> "" Then Call zlControl.TxtSelAll(txtFilter)
    If InStr(mstrCurFindtype, "��  ��") > 0 Then
        Call zlCommFun.OpenIme(True)
    End If

    If Not mobjIDCard Is Nothing And txtFilter.Text = "" Then '�������֤�����豸
        mobjIDCard.SetEnabled (True)
    End If
End Sub
Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtFilter_Validate(False)
        Call zlControl.TxtSelAll(txtFilter)
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrCurFindtype
            Case "�����", "סԺ��"
                If InStr("*+0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "�����", "סԺ��", "����"
                If Len(txtFilter.Text) > 18 Then KeyAscii = 0 '����
            Case "���￨"
                Dim blnCard As Boolean
    
                'ȥ���ſ��������������ַ�
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = zlCommFun.InputIsCard(Me.txtFilter, KeyAscii, glngSys)
                
                'ˢ����ɻ�ȷ������
                If blnCard And Len(Me.txtFilter.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtFilter.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txtFilter.Text = Me.txtFilter.Text & Chr(KeyAscii)
                        Me.txtFilter.SelStart = Len(Me.txtFilter.Text)
                    End If
                    KeyAscii = 0
                    Me.txtFilter.Text = UCase(Me.txtFilter)
                    Me.txtFilter.SetFocus
                End If
            Case "���ݺ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtFilter.Text = "" Or txtFilter.SelLength = Len(txtFilter.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "����"
            
        End Select
    Else
        If Trim(txtFilter.Text) <> "" Then
            If Mid(txtFilter.Text, 1, 1) = "*" And IsNumeric(Mid(txtFilter.Text, 2)) = True Then mstrCurFindtype = "�����"
            If Mid(txtFilter.Text, 1, 1) = "+" Then mstrCurFindtype = "סԺ��"
        End If
        Dim cbrControl As CommandBarControl
        Set cbrControl = cbrdock.FindControl(, ID_��ʼ����)
        If Not cbrControl Is Nothing Then
            cbrdock_Execute cbrControl
        End If
    End If
End Sub
Private Sub txtFilter_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
    End If
End Sub
Private Sub txtFilter_Validate(Cancel As Boolean)
    If InStr(mstrCurFindtype, "���ݺ�") > 0 Then
        If IsNumeric(txtFilter.Text) Then
            txtFilter.Text = GetFullNO(txtFilter.Text, 0)
        End If
    End If
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean)
'------------------------------------------------
'���ܣ��ڲ����б��ж�λָ���ļ�¼
'������ blnFirst -- �Ƿ��һ�β���
'���أ��ޣ�ֱ���ڲ����б��ж�λ
'------------------------------------------------
    Dim blnOK As Boolean, lngCount As Long, intB As Integer
    Dim lngRow As Long

    '���û�м�¼�����˳�
    If Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0 Then Exit Sub

    intB = 0
    
    On Error GoTo err
    
    If Not blnFirst Then
        intB = vsList.Row + 1
        If intB >= vsList.Rows Then intB = 1
    End If

    blnOK = False
    For lngCount = intB To vsList.Rows - 1 '�ڵ�ǰ״̬�в���
        Select Case mstrLocateType
            Case "��ʶ��"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("��ʶ��")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "���￨", "�ɣÿ�"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("���￨��")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "���ݺ�"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("NO")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "����"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("����")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "����"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("����")), "") Like txtLocate.Text & "*" Then blnOK = True
                If zlCommFun.SpellCode(Nvl(vsList.TextMatrix(lngCount, GetCN("����")), "")) Like UCase(txtLocate.Text) & "*" Then blnOK = True
            Case "���֤"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("���֤��")), 0) Like txtLocate.Text & "*" Then blnOK = True
        End Select
        
        If blnOK Then
            txtLocate.Tag = txtLocate.Text
            On Error Resume Next
            '���㵱ǰ�кͶ���֮��Ĳ��
            lngRow = Abs(vsList.Row - vsList.TopRow)
            
            vsList.Row = lngCount
            vsList.TopRow = vsList.Row - lngRow
            
            Exit Sub
        End If
    Next
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_���()
    Dim strReview As String
    Dim strDeptName As String
    
    On Error GoTo errHandle
    
    strDeptName = Split(mstrCur����, "-")(1)
    If frmReview.ShowMe(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")), vsList.TextMatrix(vsList.Row, GetCN("���ͺ�")), _
        Me, strDeptName, strReview) = True Then
        vsList.TextMatrix(vsList.Row, GetCN("�������")) = strReview
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_���淢��()
    '���淢��
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "Zl_Ӱ�񱨸淢��(" & vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "���淢��")
    vsList.TextMatrix(vsList.Row, GetCN("���淢��")) = IIf(vsList.TextMatrix(vsList.Row, GetCN("���淢��")) = "δ����", "�ѷ���", "δ����")
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub txtLocate_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtLocate.Text = "" And Me.ActiveControl Is txtLocate)
    If txtLocate.Text = "" Then txtLocate.Tag = ""
End Sub

Private Sub txtLocate_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '���֤ʶ�����
    
    If txtLocate.Text <> "" Then Call zlControl.TxtSelAll(txtLocate)
    If mstrLocateType = "����" Then
        Call zlCommFun.OpenIme(True)
    End If
    If Not mobjIDCard Is Nothing And txtLocate.Text = "" Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtLocate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtLocate_Validate(False)
        Call zlControl.TxtSelAll(txtLocate)
        Call SeekNextPati(txtLocate.Tag <> txtLocate.Text)
    End If
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrLocateType
            Case "��ʶ��"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "���￨"
                Dim blnCard As Boolean
    
                'ȥ���ſ��������������ַ�
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = zlCommFun.InputIsCard(Me.txtLocate, KeyAscii, glngSys)
                
                'ˢ����ɻ�ȷ������
                If blnCard And Len(Me.txtLocate.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtLocate.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txtLocate.Text = Me.txtLocate.Text & Chr(KeyAscii)
                        Me.txtLocate.SelStart = Len(Me.txtLocate.Text)
                    End If
                    KeyAscii = 0
                    Me.txtLocate.Text = UCase(Me.txtLocate)
                    Me.txtLocate.SetFocus
                End If
            Case "���ݺ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtLocate.Text = "" Or txtLocate.SelLength = Len(txtLocate.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "����"
            
        End Select
    End If
End Sub

Private Sub txtLocate_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtLocate_Validate(Cancel As Boolean)
    If InStr(mstrLocateType, "���ݺ�") > 0 Then
        If IsNumeric(txtLocate.Text) Then
            txtLocate.Text = GetFullNO(txtLocate.Text, 0)
        End If
    End If
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����: ��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(NewRow, GetCN("ҽ��ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("����") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, NewRow, GetCN("����")) + vsList.Cell(flexcpWidth, NewRow, GetCN("����")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("����")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub
Private Sub vsList_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If NewLeftCol > GetCN("����") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("����")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("����")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("����")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub

Private Sub vsList_AfterSort(ByVal Col As Long, Order As Integer)
    mlngSortCol = Col
    mintSortOrder = Order
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("����") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("����")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("����")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("����")) + 15
            cmdInfo.Visible = True
        End If
    End If
    
    Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1 '�ݴ������п�����ر�ʱ����ע���
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < GetCN("����") Then Cancel = True
End Sub

Private Sub vsList_DblClick()
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) <> "" Then
        Select Case vsList.TextMatrix(vsList.Row, GetCN("���״̬"))
            Case 1, 0
                Call Menu_Manage_����
            Case 2, 3               '˫������д����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Modify)
            Case 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '����
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Set Popup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        For Each Menucontrol In cbrMain.ActiveMenuBar.Controls
'            If Menucontrol.Parent.BarID = conMenu_ManagePopup Then
            If (Menucontrol.ID <> conMenu_FilePopup And Menucontrol.ID <> conMenu_ToolPopup _
                And Menucontrol.ID <> conMenu_ViewPopup And Menucontrol.ID <> conMenu_HelpPopup) And Menucontrol.Type = xtpControlPopup Then
                For Each control In Menucontrol.CommandBar.Controls
                    If control.ID <> conMenu_Antibody_Manage And control.ID <> conMenu_Meal_Manage And control.ID <> conMenu_Decalin_Task Then control.Copy Popup
                Next
            End If
        Next
        Popup.ShowPopup
    End If
End Sub

Private Sub vsList_RowColChange()
    On Error GoTo errHandle
    mblnIsHistory = False
    If mblnvsRefresh Then Exit Sub
    '�ж�Ƕ��ʽ����༭���еı����Ƿ�û�б���
    If mblnPacsReport = True Then    'ʹ��PACS����༭��
        Call mfrmPacsReport.PromptModify
    End If
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0 Then '�޼�¼ʱ����
        Call RefreshTabWindow(0, True)
        cboTimes.Clear
        txtAppend = ""
        lbl������Ϣ.Caption = "��  ��:" & Space(12) & "��  ��:" & Space(10) & "��  ��:" & Space(10)
        lbl�����Ϣ.Caption = "����:" & Space(17) & "���˿���:" & Space(15) & "��ʶ��:" & Space(12) & "��  ��:" & Space(10)
        lblCash.Visible = False
    Else
        Call FillHistory '������μ���¼
        Call FillTxtInfor '������Ϸ����˻�����Ϣ
        Call FillTxtAppend '������½�ҽ������
        Call ShowTab '���ݲ����ṩ��ͬѡ�
        
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        
        If mstrFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ,��TabWindow����ˢ��
            Dim i As Integer
            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).Tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    If TabWindow.Item(i).Selected Then
                        Call RefreshTabWindow
                    Else
                        TabWindow.Item(i).Selected = True
                    End If
                    
                    Exit Sub
                End If
            Next
            
            If i = TabWindow.ItemCount Then
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow(i).Visible Then
                        TabWindow(i).Selected = True 'ûѭ�����˴�����1������tab
                        Exit For
                    End If
                Next i
            End If
        Else
            Call RefreshTabWindow
        End If
        
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtInfor(Optional lngAdviceID As Long = 0)
'������Ϸ����˻�����Ϣ
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    With vsList
        lbl������Ϣ.Caption = "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")), 12, " ") & "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("�Ա�")), 10, " ") & _
                          "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")), 10, " ")
                          
        If lngAdviceID = 0 Then '---------------------------�����μ��ֱ�����б��м�¼���
            gstrSQL = "Select ���� From ���ű� Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˿���", CLng(.TextMatrix(.Row, GetCN("���˿���ID"))))
            lbl�����Ϣ.Caption = "�����:" & Rpad(.TextMatrix(.Row, GetCN("�����")), 17, " ") & "���˿���:" & Rpad(rsTemp!����, 15, " ") & _
                                    "��ʶ��:" & Rpad(.TextMatrix(.Row, GetCN("��ʶ��")), 12, " ") & _
                                    "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")) & "", 10, " ")
                                  
            lblCash.Caption = "��": lblCash.Visible = False
            lblCash.Visible = (.TextMatrix(.Row, GetCN("�շ�")) = " ")
        Else
            Dim strSQLBak As String
            gstrSQL = "Select A.ID, A.���˿���id, A.����ҽ��,A.������Դ, A.ҽ������, Nvl(A.Ӥ��, 0) Ӥ��, A.����id, A.��ҳid, A.����, A.�Һŵ�, B.����, B.���uid, C.����, D.���ͺ�,D.ִ��״̬,D.ִ�й���,0 as ת��" & vbNewLine & _
                        "From ����ҽ����¼ A, Ӱ�����¼ B, ���ű� C, ����ҽ������ D" & vbNewLine & _
                        "Where A.ID = [1] And A.ID = B.ҽ��id And A.���˿���id = C.ID And A.ID = D.ҽ��id"
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
            strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
            strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
            strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
            gstrSQL = gstrSQL & vbNewLine & " Union ALL " & strSQLBak
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����μ�¼��Ϣ", lngAdviceID)
            If Not rsTemp.EOF Then
                mlngHOrderID = lngAdviceID
                mlngHSendNo = Nvl(rsTemp!���ͺ�, 0)
                mstrHStudyUID = Nvl(rsTemp!���uid)
                mblnHMoved = IIf(rsTemp!ת�� = 1, True, False)
                fraInfo.Tag = rsTemp!����ID & "|" & rsTemp!��ҳID & "|" & rsTemp!ID & "|" & rsTemp!���ͺ� & "|" & rsTemp!���˿���ID & "|" & rsTemp!�Һŵ� & "|" & Nvl(rsTemp!������Դ, 3) & "|" & rsTemp!���uid & "|" & rsTemp!ת�� & "|" & rsTemp!ִ��״̬ & "|" & rsTemp!ִ�й��� & "|" & rsTemp!����
                lbl�����Ϣ.Caption = "�����:" & Rpad(Nvl(rsTemp!�����), 17, " ") & "���˿���:" & Rpad(rsTemp!����, 15, " ") & _
                                      "��ʶ��:" & Rpad(.TextMatrix(.Row, GetCN("��ʶ��")), 12, " ") & _
                                      "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")) & "", 10, " ")
                If rsTemp!Ӥ�� <> 0 Then
                    Dim lngBaby As Integer, lngPatID As Long, lngPageID As Long
                    lngBaby = rsTemp!Ӥ��: lngPatID = rsTemp!����ID: lngPageID = Nvl(rsTemp!��ҳID, 0)
                    gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                            "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", lngPatID, lngPageID, lngBaby)
                    If Not rsTemp.EOF Then
                        lbl������Ϣ.Caption = "��  ��:" & Rpad(rsTemp!Ӥ������, 12, " ") & "��  ��:" & Rpad(rsTemp!Ӥ���Ա�, 10, " ") & _
                                            "��  ��:" & Rpad(rsTemp!����ʱ��, 10, " ") & "ִ�й���:" & Nvl(rsTemp!����ִ�й���)
                    End If
                End If
            Else
                lbl�����Ϣ.Caption = "�����:" & Space(17) & "���˿���:" & Space(15) & "��ʶ��:" & Space(12) & "��  ��:" & Space(10)
            End If
            lblCash.Caption = "��": lblCash.Visible = True
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillTxtAppend(Optional lngAdviceIDtmp As Long = 0)
'������½�ҽ������
Dim lngAdviceID As Long, strAppend As String, rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo errHandle
    With vsList
        If lngAdviceIDtmp = 0 Then
            lngAdviceID = Val(.TextMatrix(.Row, GetCN("ҽ��ID")))
        Else
            lngAdviceID = lngAdviceIDtmp
        End If
        
        If lngAdviceIDtmp = 0 Then '-------------------------------------------�б�ѡ�����
            txtAppend = "�����Ŀ:" & .TextMatrix(.Row, GetCN("ҽ������")) & vbCrLf
            txtAppend = txtAppend & "����ҽ��:" & Rpad(.TextMatrix(.Row, GetCN("����ҽ��")), 8, " ") & vbCrLf
            
            If .TextMatrix(.Row, GetCN("��λ����")) <> "" Then
                For i = 0 To UBound(Split(.TextMatrix(.Row, GetCN("��λ����")), "),"))
                    If i = 0 Then
                        txtAppend = txtAppend & "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(.TextMatrix(.Row, GetCN("��λ����")), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.TextMatrix(.Row, GetCN("��λ����")), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) 'ȡ����������
            Else
                txtAppend = txtAppend & "��鲿λ:" & .TextMatrix(.Row, GetCN("ҽ������"))
            End If
            gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
            If .TextMatrix(.Row, GetCN("ת��")) = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        Else                    '-------------------------------------------���μ�¼ѡ�����
            Dim strTemp As String
            txtAppend = ""
            
            gstrSQL = "Select ����ҽ��,ҽ������ From ����ҽ����¼ Where  id =[1]"
            If Split(fraInfo.Tag, "|")(8) = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ������", lngAdviceID)
            
            If rsTemp.EOF = False Then
                strTemp = Nvl(rsTemp!ҽ������)
                If InStr(strTemp, ":") > 0 Then
                    txtAppend = "�����Ŀ:" & Split(strTemp, ":")(0) & vbCrLf
                Else
                    txtAppend = "�����Ŀ:" & strTemp & vbCrLf
                End If
                
                txtAppend = txtAppend & "����ҽ��:" & rsTemp!����ҽ�� & vbCrLf
            End If
            
            If strTemp <> "" Then
                If InStr(strTemp, ":") > 0 Then
                    strTemp = Split(strTemp, ":")(1)
                    For i = 0 To UBound(Split(strTemp, "),"))
                        If i = 0 Then
                            txtAppend = txtAppend & "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(strTemp, "),")(i) & ")"
                        Else
                            txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(strTemp, "),")(i) & ")"
                        End If
                    Next
                    If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) 'ȡ����������
                Else
                    txtAppend = txtAppend & strTemp
                End If
            End If
            gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����" '�������μ�¼�Ƿ�ת���жϲ���ʷ��
            If Split(fraInfo.Tag, "|")(8) = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˸���", lngAdviceID)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!��Ŀ & ":" & Nvl(rsTemp!����) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        txtAppend = txtAppend & vbCrLf & vbCrLf & strAppend
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillHistory()
'������μ���¼
Dim rsTemp As ADODB.Recordset, strTemp As String
    On Error GoTo errHandle
    With vsList
        cboTimes.Tag = "" 'cbotime����ʱ�õ�������������"������Ŀ"ʱ��������"���cbotimes"����
        gstrSQL = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                   " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C" & _
                   " Where A.����id = [1] And A.���id Is Null And A.ִ�п���id+0 =[2] And B.ҽ��ID=A.ID " & _
                   "" & IIf(.TextMatrix(.Row, GetCN("������")) = "�Ѿܾ�", "", " And B.ִ��״̬<>2 ") & _
                   " AND A.ID=C.ҽ��ID"
        
        '���ù������ˣ��Ų�ѯ����ID
        If mblnRelatingPatient = True And .TextMatrix(.Row, GetCN("����ID")) <> 0 Then
            gstrSQL = gstrSQL & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                    " From ����ҽ����¼ A " & _
                    " Where A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[3]) "
        End If
        
        strTemp = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
        strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
        strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
        gstrSQL = gstrSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order By ����ʱ�� Asc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(.TextMatrix(.Row, GetCN("����ID"))), mlngCur����ID, _
                    CLng(.TextMatrix(.Row, GetCN("����ID"))))
        
        cboTimes.Clear
        Do Until rsTemp.EOF
           cboTimes.AddItem "��" & rsTemp.AbsolutePosition & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & Trim(rsTemp!ҽ������)
           cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!ҽ��ID
           If rsTemp!ҽ��ID = .TextMatrix(.Row, GetCN("ҽ��ID")) Then cboTimes.ListIndex = cboTimes.NewIndex
           rsTemp.MoveNext
        Loop
        cboTimes.Tag = "���"
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ShowTab(Optional lngAdviceID As Long = 0)
'���ݲ�����Դ���Ʋ�����ҽ��ѡ�
Dim int��Դ As Integer, i As Integer
Dim strFirstTab As String
Dim intDefaultIndex As Integer

    On Error GoTo errHandle
    
    If lngAdviceID = 0 Then '-------------------------------------------�б�ѡ�����
        int��Դ = Val(vsList.Cell(flexcpData, vsList.Row, GetCN("��Դ")))
        Dim blnShowReport As Boolean
        '�ж� ��ͼ����д����
        blnShowReport = True
        If mblnReportWithImage = True Then
            If vsList.TextMatrix(vsList.Row, GetCN("���UID")) = "" Then blnShowReport = False
        End If
    Else                    '-------------------------------------------���μ�¼ѡ�����
        '���μ�¼ʱfraInfo.Tag = 0����ID|1��ҳID|2ҽ��ID|3���ͺ�|4���˿���ID|5�Һŵ�|6������Դ|7���UID|8ת��
        int��Դ = Split(fraInfo.Tag, "|")(6)
    End If
    
    If int��Դ <> 2 Then '���ݲ�����Դ���Ʋ�����ҽ��ѡ�
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = True
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = False
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, GetCN("���״̬")) > 1 And blnShowReport, True)
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = False
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = True
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, GetCN("���״̬")) > 1 And blnShowReport, True)
            End Select
        Next
    End If
    
    
    
    '�����ǰ��ѡ���ҳ�治�ɼ�������ʾ�û�����Ҫ����ҳ��
    If TabWindow.Selected.Visible = False Then
        strFirstTab = mstrFirstTab
'        If strFirstTab = "" Then strFirstTab = "Ӱ��"
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).Tag, strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            ElseIf InStr(TabWindow(i).Tag, "Ӱ��") > 0 Then
                intDefaultIndex = i
            End If
        Next i
        
        If i = TabWindow.ItemCount Then
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Visible Then
                    TabWindow(i).Selected = True
                    Exit For
                End If
            Next i
        End If
    End If
    
    If TabWindow.Selected.Visible = False Then
        TabWindow(intDefaultIndex).Visible = True
    End If
    
'    '@�޸�����30490
'    For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
'        If TabWindow(i).Tag = "����ȡ��" Then
'            TabWindow(i).Visible = IIf(vsList.TextMatrix(vsList.Row, GetCN("������")) = "�ѵǼ�", False, True)
'        End If
'    Next
'    '@�޸�����30490
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0, Optional blnClear As Boolean = False, Optional blnRefresh As Boolean = False)
'lngAdviceIDtmp���μ�¼ʱ���� , ������0, blnclear��յ�ǰ�б�, blnRefreshǿ��ˢ��
'ˢ�µ�ǰҳ��,���ã��б�ѡ�����μ�¼ѡ���Ӵ���ѡ��
'���μ�¼ʱfraInfo.Tag = 0����ID|1��ҳID|2ҽ��ID|3���ͺ�|4���˿���ID|5�Һŵ�|6������Դ|7���UID|8ת��|9ִ��״̬
Dim lngAdviceID As Long, lngSendNO As Long, lngPatID As Long, lngPageID As Long, blnCanPrint As Boolean, blnIsInsidePatient As Boolean
Dim lngUnit As Long, lngPatDept As Long, strRegNo As String, intMoved As Boolean, intState As Integer, intStep As Integer, i As Integer, intPatientForm As Integer
Dim strInfo As String

    On Error GoTo errHandle
    If lngAdviceIDtmp = 0 Then '-----------------------�б�ѡ�����
        If blnClear Then       '�޼�¼ʱ��������Ӵ���
            lngAdviceID = 0: lngSendNO = 0: lngPatID = 0: lngPageID = 0
            lngPatDept = 0: strRegNo = "": intMoved = 0: intState = 0: lngUnit = 0: blnCanPrint = False
        Else
            With vsList
                strInfo = .TextMatrix(.Row, GetCN("����"))
                lngAdviceID = .TextMatrix(.Row, GetCN("ҽ��ID")): lngSendNO = .TextMatrix(.Row, GetCN("���ͺ�"))
                lngPatID = .TextMatrix(.Row, GetCN("����ID")): lngPageID = Val(.TextMatrix(.Row, GetCN("��ҳID")))
                lngPatDept = .TextMatrix(.Row, GetCN("���˿���ID")): strRegNo = .TextMatrix(.Row, GetCN("�Һŵ�"))
                intMoved = .TextMatrix(.Row, GetCN("ת��"))
                intState = IIf(.TextMatrix(.Row, GetCN("������")) = "�Ѿܾ�", 2, IIf(.TextMatrix(.Row, GetCN("������")) = "�����", 1, 3))
                intStep = .TextMatrix(.Row, GetCN("���״̬")) '��ȡִ�й���
                lngUnit = Val(.TextMatrix(.Row, GetCN("��ǰ����ID")))
                blnCanPrint = IIf(mblnCanPrint, IIf(.Cell(flexcpData, .Row, GetCN("����")) = 1, .TextMatrix(.Row, GetCN("������")) <> "", .TextMatrix(.Row, GetCN("������")) <> ""), True)
                intPatientForm = Decode(.TextMatrix(.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4)
            End With
        End If
    Else                       '----------------------���μ�¼ѡ�����
        lngAdviceID = lngAdviceIDtmp: lngSendNO = Split(fraInfo.Tag, "|")(3)
        lngPatID = Split(fraInfo.Tag, "|")(0): lngPageID = Val(Split(fraInfo.Tag, "|")(1))
        lngPatDept = Split(fraInfo.Tag, "|")(4): strRegNo = Split(fraInfo.Tag, "|")(5)
        intMoved = Split(fraInfo.Tag, "|")(8): intState = Split(fraInfo.Tag, "|")(9)
        intStep = Split(fraInfo.Tag, "|")(10)
        strInfo = Split(fraInfo.Tag, "|")(11)
        lngUnit = lngPatDept
        blnCanPrint = True
        intPatientForm = Split(fraInfo.Tag, "|")(6)
    End If
    
    blnIsInsidePatient = (intPatientForm = 1) Or (intPatientForm = 2)
    
    Select Case TabWindow(TabWindow.Selected.Index).Tag
        Case "�������"
            mobjExpense.zlRefresh mlngCur����ID, lngAdviceID, lngSendNO, intMoved = 1
        Case "������д"
            
            If mblnPacsReport = True Then
                mfrmPacsReport.zlRefresh lngAdviceID, Me, intMoved = 1, strInfo
                
                If GetActiveWindow = Me.hWnd Then Call mfrmPacsReport.ShowVideoWindow
            Else
                '���Ӳ����༭��
                mobjReport.zlRefresh lngAdviceID, mlngCur����ID, Not mblnIsHistory, intMoved = 1, blnCanPrint
            End If
            
        Case "�Ŷӽк�"
            If Not mblnIsHistory And Not mobjQueue Is Nothing Then
                mobjQueue.zlRefresh mAstr��������, Split(mstrCur����, "-")(1) & vsList.TextMatrix(vsList.Row, GetCN("ִ�м�")), lngAdviceID
            End If
        Case "סԺҽ��"
            If TabWindow.Selected.Visible Then '������סԺ��¼ת�����������¼,��ʱ����û����Ȩ����ҽ��Ȩ��
                mobjInAdvice.zlRefresh lngPatID, lngPageID, lngUnit, lngPatDept, 0, intMoved = 1, lngAdviceID, intState, lngPatDept
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "����ҽ��" Then
                        If strRegNo = "" Then   '���еǼǵĲ���û�йҺŵ���
                            mobjOutAdvice.zlRefresh lngPatID, "", False
                        Else
                            mobjOutAdvice.zlRefresh lngPatID, strRegNo, Not mblnIsHistory And blnIsInsidePatient, intMoved = 1, lngAdviceID
                        End If
                    End If
                Next
            End If
        Case "����ҽ��"
            If TabWindow.Selected.Visible Then '�����������¼ת������סԺ��¼,��ʱ����û����ȨסԺҽ��Ȩ��
                If strRegNo = "" Then   '���еǼǵĲ���û�йҺŵ���
                    mobjOutAdvice.zlRefresh lngPatID, "", False
                Else
                    mobjOutAdvice.zlRefresh lngPatID, strRegNo, Not mblnIsHistory And blnIsInsidePatient, intMoved = 1, lngAdviceID
                End If
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "סԺҽ��" Then
                      mobjInAdvice.zlRefresh lngPatID, lngPageID, lngUnit, lngPatDept, 0, intMoved = 1, lngAdviceID, intState, lngPatDept
                    End If
                Next
            End If
        Case "סԺ����"
            If TabWindow.Selected.Visible Then '������סԺ��¼ת�����������¼,��ʱ����û����Ȩ���ﲡ��Ȩ��
                mobjInEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "���ﲡ��" Then
                       mobjOutEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
                    End If
                Next
            End If
        Case "���ﲡ��"
            If TabWindow.Selected.Visible Then '�����������¼ת������סԺ��¼,��ʱ����û����ȨסԺ����Ȩ��
                mobjOutEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "סԺ����" Then
                        mobjInEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
                    End If
                Next
            End If
            
        Case "�걾����"
'            If mfrmPatholSpecimen.Visible Then
                If intState = 6 Or intState = 0 Or intState = 1 Then '�鿴ģʽ
                    mfrmPatholSpecimen.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                Else
                    mfrmPatholSpecimen.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                End If
'            End If
        Case "����ȡ��"
'            If mfrmPatholMaterial.Visible Then
                If intState = 6 Or intState = 0 Or intState = 1 Then '�鿴ģʽ
                    mfrmPatholMaterial.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                Else
                    mfrmPatholMaterial.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                End If
'            End If
        Case "������Ƭ"
'            If mfrmPatholSlices.Visible Then
                If intState = 6 Or intState = 0 Or intState = 1 Then '�鿴ģʽ
                    mfrmPatholSlices.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                Else
                    mfrmPatholSlices.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                End If
'            End If
            
        Case "������"
'            If mfrmPatholSpeExam.Visible Then
                If intState = 6 Or intState = 0 Or intState = 1 Then '�鿴ģʽ
                    mfrmPatholSpeExam.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                Else
                    mfrmPatholSpeExam.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                End If
'            End If
        Case "����/�ؼ챨��"
            If mfrmPatholProRep.Visible Then
                If intState = 6 Or intState = 0 Or intState = 1 Then '�鿴ģʽ
                    mfrmPatholProRep.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                Else
                    mfrmPatholProRep.zlRefresh lngAdviceID, mstrPrivs, intMoved = 1, mlngCur����ID
                End If
            End If
        Case "Ӱ��ɼ�"
            If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then
                Call frmVideoCapture.SetRestoreContainer(picVideoContainer)
                
                If intStep = 6 Or intStep = 0 Or intStep = 1 Then  '�鿴ģʽ
                    Call frmVideoCapture.zlBeginCapture(lngAdviceID, True, False, intMoved = 1, strInfo)
                Else
                    Call frmVideoCapture.zlBeginCapture(lngAdviceID, InStr(mstrPrivs, "��Ƶ�ɼ�") <= 0, False, intMoved = 1, strInfo)
                End If
                
                '���û�п����������ڣ�����Ƕ��ҳ������ʾ��Ƶ
                If Not (TypeOf frmVideoCapture.ParentContainerObj Is frmVideoDockWindow) Then
                    If GetActiveWindow = Me.hWnd Then Call frmVideoCapture.ShowVideoWindow(picVideoContainer)
                End If
            End If
    End Select
    
    If CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then
        '���Ϊ�����ɼ�״̬������ı�֮���޸Ĳɼ�ģ��������Ϣ
        If TypeOf frmVideoCapture.ParentContainerObj Is frmVideoDockWindow Then
            If GetActiveWindow = Me.hWnd Then
                Call frmVideoCapture.zlBeginCapture(lngAdviceID, InStr(mstrPrivs, "��Ƶ�ɼ�") <= 0, False, intMoved = 1, strInfo)
            End If
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subTriggleRefreshTimer(blnEnable As Boolean)
    '�������߹ر��Զ�ˢ�µ�Timer
    If blnEnable = False Then
        TimerRefresh.Enabled = False
    Else
        TimerRefresh.Enabled = mlngRefreshInterval > 0
    End If
End Sub

Private Sub Menu_Manage_��������()
'��������
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0 Then Exit Sub
    
    On Error GoTo err
    Call frmReferencePatient.zlShowMe(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")), _
        vsList.TextMatrix(vsList.Row, GetCN("����")), Me, True)
    
    'ˢ�²����б�
     Call RefreshList
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub Menu_Manage_�������()
'�������
If Not (CheckPopedom(mstrPrivs, "�������") Or CheckPopedom(mstrPrivs, "���巴��")) Then
    Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
    Exit Sub
End If


Dim frmAntibody As New frmPatholAntibody
On Error GoTo errFree
    Call frmAntibody.ShowAntibodyManageWind(mstrPrivs, Me)
errFree:
    Call Unload(frmAntibody)
    Set frmAntibody = Nothing
End Sub



Private Sub Menu_Manage_�ײ�ά��()
'�ײ�ά��

If Not CheckPopedom(mstrPrivs, "�ײ�ά��") Then
    Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
    Exit Sub
End If

Dim frmMeal As New frmPatholMeal
On Error GoTo errFree
    Call frmMeal.ShowMealWindow(mstrPrivs, Me)
errFree:
    Call Unload(frmMeal)
    Set frmMeal = Nothing
End Sub


Private Sub Menu_Manage_��������()
'��������
If Not (CheckPopedom(mstrPrivs, "�ؼ�����") Or CheckPopedom(mstrPrivs, "��Ƭ����") Or CheckPopedom(mstrPrivs, "��ȡ����")) Then
    Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
    Exit Sub
End If

Dim lngAdviceID As Long
Dim frmRequest As New frmPatholRequisition
On Error GoTo errFree
    lngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")))
    Call frmRequest.zlRefresh(lngAdviceID, mstrPrivs, mblnMoved, mlngCur����ID, Me)
errFree:
    Call Unload(frmRequest)
    Set frmRequest = Nothing
End Sub


Private Sub Menu_Manage_�ӳٵǼ�()
'�ӳٵǼ�
If Not CheckPopedom(mstrPrivs, "�����ӳ�") Then
    Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
    Exit Sub
End If

Dim lngAdviceID As Long
Dim frmDelay As New frmPatholReportDelay
On Error GoTo errFree
    lngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")))
    Call frmDelay.zlRefresh(lngAdviceID, mstrPrivs, mblnMoved, mlngCur����ID, Me)
errFree:
    Call Unload(frmDelay)
    Set frmDelay = Nothing
End Sub



Private Sub Menu_Manage_�������뷴��(ByVal lngMenuId As Long)
'�������뷴��

If Not (CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "���ﷴ��")) Then
    Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
    Exit Sub
End If

Dim lngAdviceID As Long
Dim frmConRequest As New frmPatholConsultation
On Error GoTo errFree
    lngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")))
    
    If lngMenuId = conMenu_Con_Feedback Then
        Call frmConRequest.zlRefresh(lngAdviceID, mstrPrivs, mblnMoved, mlngCur����ID, True, Me)
    Else
        Call frmConRequest.zlRefresh(lngAdviceID, mstrPrivs, mblnMoved, mlngCur����ID, False, Me)
    End If
errFree:
'    Call Unload(frmConRequest)
'    Set frmConRequest = Nothing
End Sub


Private Sub Menu_Manage_�Ѹ��������()
'�Ѹ��������

If Not CheckPopedom(mstrPrivs, "����ȡ��") Then
    Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
    Exit Sub
End If

Call mfrmPatholDecalinTask.ShowDecalinTaskWind(Me)
End Sub



Public Sub VideoCallBack(EventType As TVideoEventType, lngAdviceID As Long, Optional strStudyUID As String, Optional strPatientName As String, Optional blnIsLock As Boolean)

    Select Case EventType
        Case vetLockStudy
            '�޸ı�ǩҳ����ʾ��ʽ�ͱ���
            Dim i As Integer
    
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*Ӱ��ɼ�*" Then
                    If blnIsLock Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "��" & strPatientName & "�� Ӱ��ɼ�"
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "Ӱ��ɼ�"
                    End If
            
                    Exit For
                End If
            Next i
        Case vetAddFirstImg, vetDelLastImg
            '�����������б���ʾ
            If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then Exit Sub

            If EventType = vetAddFirstImg Then
                '���¼���б�
                Call UpdateStudyListState(lngAdviceID, strStudyUID, True, True)
            Else
                '���¼���б�
                Call UpdateStudyListState(lngAdviceID, strStudyUID, False, True)
            End If


            If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) <> CStr(lngAdviceID) Then Exit Sub
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mfrmPacsReport Is Nothing Then
                If mfrmPacsReport.mblnShowImage Then
                    mfrmPacsReport.RefPacsPic
                End If
            End If

            'ˢ�µ������洰���е�ͼ��
            If Not mfrmPacsReportDock Is Nothing Then
                If mfrmPacsReportDock.mblnShowImage Then
                 mfrmPacsReportDock.RefPacsPic
                End If
            End If

            'ˢ�µ��Ӳ�����ͼ��
            If Not mobjReport Is Nothing Then
                mobjReport.RefPacsPic
            End If
        Case vetRecVideo
    End Select
        
    '���±�����Ƕ�ײɼ�״̬
    Call mfrmPacsReport.VideoCallBack(EventType, lngAdviceID, strStudyUID, strPatientName, blnIsLock)
    
    On Error Resume Next
    
    Dim j As Integer
    For j = LBound(mobjPacsReportArry) To UBound(mobjPacsReportArry)
        Call mobjPacsReportArry(j).VideoCallBack(EventType, lngAdviceID, strStudyUID, strPatientName, blnIsLock)
    Next j
End Sub

