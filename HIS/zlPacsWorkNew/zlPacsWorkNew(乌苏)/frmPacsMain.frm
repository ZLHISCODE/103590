VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.commandbars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#71.0#0"; "zl9PacsControl.ocx"
Object = "*\A..\..\0_zlhisPro\common vbp\zlIDKind.vbp"
Begin VB.Form frmPacsMain 
   Caption         =   "Ӱ����վ"
   ClientHeight    =   7605
   ClientLeft      =   8595
   ClientTop       =   975
   ClientWidth     =   11400
   Icon            =   "frmPacsMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerVideoEvent 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9015
      Top             =   165
   End
   Begin VB.Timer timerCapture 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8505
      Top             =   135
   End
   Begin VB.PictureBox picTemp 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   4815
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Timer timerOperHint 
      Interval        =   500
      Left            =   7920
      Tag             =   "0"
      Top             =   120
   End
   Begin VB.PictureBox picWindow 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   5400
      ScaleHeight     =   4575
      ScaleWidth      =   5535
      TabIndex        =   11
      Top             =   2160
      Width           =   5535
      Begin zl9PacsControl.TranControl tcDisable 
         Height          =   975
         Left            =   4560
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1720
         AlphaValue      =   25
      End
      Begin VB.PictureBox picLoadState 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1320
         ScaleHeight     =   1095
         ScaleWidth      =   3855
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
         Begin VB.PictureBox picSmile 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   240
            Picture         =   "frmPacsMain.frx":1CFA
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   25
            Top             =   240
            Width           =   360
         End
         Begin VB.Label labLoadState 
            Caption         =   " ���ڼ��ع���ģ�飬�����ĵȴ�..."
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.PictureBox picReportContainer 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Left            =   3720
         ScaleHeight     =   2055
         ScaleWidth      =   1815
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin DicomObjects.DicomViewer dcmRelateViewer 
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
      _Version        =   262147
      _ExtentX        =   4471
      _ExtentY        =   1931
      _StockProps     =   35
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7320
      Top             =   120
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
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPacsMain.frx":2771
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   6675
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":3005
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":359F
            Key             =   "סԺ"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":3E79
            Key             =   "����"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":3FD3
            Key             =   "Ӱ��"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":474D
            Key             =   "��ɫͨ��"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":48A7
            Key             =   "·��"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":4E41
            Key             =   "�޷�"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":51DB
            Key             =   "Ƿ��"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":5575
            Key             =   "�շ�"
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":590F
            Key             =   "����"
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":6021
            Key             =   "����"
            Object.Tag             =   "11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":63BB
            Key             =   "Σ��"
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":6755
            Key             =   "��鼼ʦ"
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":6E4F
            Key             =   "�Ƿ�"
            Object.Tag             =   "14"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6060
      Top             =   90
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
            Picture         =   "frmPacsMain.frx":71E9
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":7783
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":7D1D
            Key             =   "��λ"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":80AF
            Key             =   "����"
            Object.Tag             =   "4"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6540
      Left            =   45
      ScaleHeight     =   6540
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   555
      Width           =   4495
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picExeState 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton optNeed 
            Caption         =   "��ִ��"
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   50
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAccept 
            Caption         =   "�ѽ���"
            Height          =   180
            Left            =   1080
            TabIndex        =   18
            Top             =   50
            Width           =   975
         End
         Begin VB.OptionButton optFinal 
            Caption         =   "��ִ��"
            Height          =   180
            Left            =   2040
            TabIndex        =   17
            Top             =   50
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "����"
            Height          =   180
            Left            =   3000
            TabIndex        =   16
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.PictureBox picAppend 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   240
         ScaleHeight     =   2775
         ScaleWidth      =   3945
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3945
         Begin VB.ComboBox cboTimes 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Width           =   2235
         End
         Begin VB.TextBox txtAppend 
            BackColor       =   &H00FDD6C6&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   10
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1260
            Width           =   3920
         End
         Begin VB.Label labStudyNum 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "���ţ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label labHistory 
            Caption         =   "��ʷ��飺"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
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
            Height          =   465
            Left            =   3360
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl������Ϣ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������      �Ա�    ���䣺  "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label lbl�����Ϣ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   270
         End
      End
      Begin VB.PictureBox PicLine 
         BorderStyle     =   0  'None
         Height          =   90
         Left            =   240
         MousePointer    =   7  'Size N S
         ScaleHeight     =   90
         ScaleWidth      =   3975
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3975
      End
      Begin zl9PACSWork.ucFlexGrid ufgStudyList 
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
         DefaultCols     =   ""
         HeadCheckValue  =   1
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         ReadOnly        =   -1  'True
         IsShowPopupMenu =   0   'False
         HeadFontCharset =   134
         HeadFontWeight  =   400
         HeadColor       =   0
         DataFontCharset =   134
         DataFontWeight  =   400
         DataColor       =   -2147483640
         GridLineColor   =   14737632
         ExtendLastCol   =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   4005
         _Version        =   589884
         _ExtentX        =   7064
         _ExtentY        =   661
         _StockProps     =   64
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   330
         Left            =   360
         TabIndex        =   28
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPacsMain.frx":8441
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         CaptionAlignment=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "���￨"
         IDkindBorderStyle=   1
         IDKindWidth     =   1330
         FindPatiShowName=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Bindings        =   "frmPacsMain.frx":84C8
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPacsMain.frx":84DC
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPacsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#Const DebugImmediately = False

Private Const M_BLN_ALL_FUNCTIONS_OPEN As Boolean = True
Private Const M_STR_MODULE_MENU_TAG As String = "Main"

Private Const mintCurҵ������ As Integer = 1 '��ǰϵͳ������ҵ�����ͣ������Ŷӽк�ʹ��



'������
'����ϵͳ��ͬ����[------]������������ݽ����в�������滻
Private Const M_STR_PUBLIC_COLS = "|·��>·��״̬,w400|����>������־,headimg1,w300|��Դ,headimg2,w400" & _
                        "|�շ�,headimg9,w300|Σ��,headimg12,w800|����,headimg3,w300|����,btn,txtleft,w1200,uncfg" & _
                        "|���뵥>���뵥ҽ��,w1100|������>[placecol],w800|ִ��״̬,hide,uncfg|�Ա�,w450|����,w450|��ʶ��,w1400|[------]|��������,w800|ҽ������,w2400" & _
                        "|��λ����>[placecol],w1400|����ʱ��,w1800,shortdatetime|����ʱ��,w1800,shortdatetime|����ҽ��,w800|���,hide,w450" & _
                        "|����,hide,w450|Ӥ��,w450|�Ǽ���,w800|������,w800|�����,w800|�������,w800|��ɫͨ��,hide,uncfg" & _
                        "|�����ӡ,w800|������,w800|������,w800|��ͼʱ��,w1800,shortdatetime|�������,w2400|����ID,hide,uncfg" & _
                        "|��ҳID,hide,uncfg|�Һŵ�,hide|���˿���ID,hide,uncfg|ҽ��ID,key,hide,w1200|���ͺ�,hide,uncfg" & _
                        "|���UID,hide,uncfg|���״̬>������,hide,uncfg|NO,hide,uncfg|��¼����,hide,uncfg|ת��,hide,uncfg" & _
                        "|����>��ǰ����,hide|��ǰ����ID,hide,uncfg|���淢��,w800|��Ϸ���,w800|���˿���,w800|����ID,hide,uncfg" & _
                        "|���￨��,w800|���ݺ�,w800|���֤��,w800|����ʱ��,hide,uncfg,shortdatetime|ͼ��λ��,hide,uncfg|�Ƿ�ʦȷ��,hide,uncfg|"

Private Const M_STR_PUBLIC_COLS_NEW = "|·��>·��״̬,w400|����>������־,headimg1,w300|��Դ,headimg2,w400" & _
                        "|�շ�,headimg9,w300|Σ��,headimg12,w800|����,btn,txtleft,w1200,uncfg" & _
                        "|���뵥>���뵥ҽ��,w1100|������>[placecol],w800|ִ��״̬,hide,uncfg|�Ա�,w450|����,w450|��ʶ��,w1400|[------]|ҽ������,w2400" & _
                        "|��λ����>[placecol],w1400|����ʱ��,w1800,shortdatetime|����ʱ��,w1800,shortdatetime|����ҽ��,w800|���,hide,w450" & _
                        "|����,hide,w450|Ӥ��,w450|�Ǽ���,w800|������,w800|�����,w800|��ɫͨ��,hide,uncfg" & _
                        "|��ͼʱ��,w1800,shortdatetime|�������,w2400|����ID,hide,uncfg" & _
                        "|��ҳID,hide,uncfg|�Һŵ�,hide|���˿���ID,hide,uncfg|ҽ��ID,key,hide,w1200|���ͺ�,hide,uncfg" & _
                        "|���UID,hide,uncfg|���״̬>������,hide,uncfg|NO,hide,uncfg|��¼����,hide,uncfg|ת��,hide,uncfg" & _
                        "|����>��ǰ����,hide|��ǰ����ID,hide,uncfg|��Ϸ���,w800|���˿���,w800|����ID,hide,uncfg" & _
                        "|���￨��,w800|���ݺ�,w800|���֤��,w800|����ʱ��,hide,uncfg,shortdatetime|ͼ��λ��,hide,uncfg|�Ƿ�ʦȷ��,hide,uncfg|"

'����
Private Const M_STR_PATHOL_COLS = "�����,w1400|����>�ۺ�����,w280|����ִ��״̬,w1400|������|�������,w1200|ȡ�Ĺ���,hide,uncfg|��Ƭ����,hide,uncfg|���߹���,hide,uncfg|���ӹ���,hide,uncfg|��Ⱦ����,hide,uncfg|"
'ҽ��
Private Const M_STR_IMAGES_COLS = "����,w1400|Ӱ�����|Ӱ������,w280|�������,w280|ִ�м�,w600|���ӽ�Ƭ>�Ƿ���ӽ�Ƭ,W600|��Ƭ��ӡ>�Ƿ��ӡ,w800|��鼼ʦ,w800|��鼼ʦ��,w1000|��Ƭ����>���Ž�Ƭ,w800|ִ�п���ID,hide,uncfg"
'�ɼ�
Private Const M_STR_CAPTOR_COLS = "����,w1400|Ӱ�����|Ӱ������,w280|�������,w280|ִ�м�,w600|��鼼ʦ,w800|��鼼ʦ��,w1000"


'��û������ʱ��ʹ�ô���ʾ��Ϣ
Private Const M_STR_HINT_NoSelectData As String = "��ѡ����Ҫִ�еļ�����ݡ�"

'���ݲ�ͬϵͳ����[------]�������滻Ϊ�����š����ߡ�����š�
Private Const CONST_STR_LOCAL_CARD_TYPE As String = "����;���￨;��ʶ��;���ݺ�;[------];���֤��;������;IC����;"
Private Const CONST_STR_FIND_CARD_TYPE As String = "����;���￨;�����;סԺ��;���ݺ�;[------];���֤��;������;IC����;"

Private Enum TLocateFindType
    lftLocate = 0
    lftFind = 1
End Enum


'��ǰҽ����Ϣ
Private Type TAdviceInf
    lngPatID As Long                '1 ����ID
    lngPageID As Long               '2 ��ҳID
    lngAdviceId As Long             '3 ҽ��ID
    lngSendNo As Long               '4 ���ͺ�
    strPatientName As String        '5 ��������
    
    lngPatDept As Long              '6 ������������
    strRegNo As String              '7 �Һŵ�
    lngRegId As Long                '8 �Һ�id
    intMoved As Integer             '9 �Ƿ�ת��
    intState As Integer             '10 ���״̬
    intStep As Integer              '11 ������
    lngUnit As Long                 '12 ����ID
    strStudyUID As String           '13 ���UID
    blnCanPrint As Boolean          '14 �Ƿ��ܹ���ӡ
    blnIsInsidePatient As Boolean   '15 �Ƿ������סԺ����
    lngExeDepartmentId As Long      '16 ִ�в���ID
    strDoDoctor As String           '17 ��鼼ʦ
    strExeRoom As String            '18 ִ�м�
    lngPatientFrom As Long          '19 ������Դ
    
    strStudyNum As String           '20 ����
    strBedNum As String             '21 ����
    lngMarkNum As Double            '22 ��־��
    lngBaby As Long                 '23 Ӥ��
    strPatientDepartment As String  '24 ���˿�������
    
    strReportDoctor As String       '25 ������
    strReportOperation As String    '26 �������
    lngLinkId As Long               '27 ����ID
    strImgType As String            '28 Ӱ�����
    intImageLocation As Integer     '29 PACSӰ�����ڵ�λ�ã�0������PACS��1������PACS
    strPatientType As String        '30 ��������
End Type


'������������
Private Type Type_SQLCondition
    ��ʼʱ�� As Date
    ����ʱ�� As Date
    ʱ������ As Integer                                 'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
    ���ݺ� As String
    ����� As Double
    ������ As String
    סԺ�� As Double
    ���￨ As String
    ���� As String
    �Ա� As String
    ��ʼ���� As Long
    �������� As Long
    �������� As String
    ���� As Variant
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
    Ӱ����� As String
    ������� As String
    ������ As String
    ���� As String
    ��� As String
    ����ID As Long
End Type

'ϵͳ�������Ͷ���
Private Type TSystemPar

    '���ز���
    strFirstTab As String                               '�״���ʾ��ҳ��
    blnֱ�Ӽ�� As Boolean                              '�ǼǺ�ֱ�ӽ�����
    blnWriteCapDoctor As Boolean                        '�Ƿ��ڲɼ�ͼ����Զ��ѵ�ǰ�û���дΪ��鼼ʦ
    blnAutoOpenReport As Boolean                        '��ʼ����Զ��򿪱���
    blnNoShowCancel As Boolean                          '����ʾȡ���ļ��
    blnPatTrack As Boolean                              '�Ƿ�Խ����˽��и���
    strLocalRoom As String                              '����ִ�м�����
    
    '���̲���
    blnFinishCommit As Boolean                          '�ޱ��������,�Ƿ������ٴ�ȷ��
    blnCompleteCommit As Boolean                        '��˺������ٴ�ȷ��
    blnIgnoreResult As Boolean                          '���������� '=true ����
    
    blnReportWithImage As Boolean                       '��ͼ�����д���棬��ͼ�񲻿�д����
    blnReportWithResult As Boolean                      '�������Խ������д����
    blnLocalizerBackward As Boolean                     '��λƬ����
    
    blnPrintCommit As Boolean                           '��ӡ��ֱ�����
    blnCanPrint As Boolean                              'ƽ����Ҫ��˲��ܴ�ӡ =true
    lngBeforeDays As Long                               'Ĭ�ϲ�ѯ������
    lngRefreshInterval As Long                          '�����б��Զ�ˢ�¼��
    blnUseQueue As Boolean                              '�Ƿ������Ŷӽк�
    blnSynStudylist As Boolean                          '�Ŷӽк�ʱ������Ŷ��б������б����ݺ��Ƿ�ͬ����λ������б�
    blnAutoInQueue As Boolean                           '�����Ŷӽкź��Ƿ��Զ����
    
    blnRelatingPatient As Boolean                       '�Ƿ����ù�������
    lngQueueWay As Long                                 '�жϷ�ʽ��0��ִ�м������Ŷӣ�1�����������Ŷ�
'    lngSameTime As Long                                 '���ŷ�ʽ��0����ͽ�Ƭ�ֱ𷢷� 1 ����ͽ�Ƭͬʱ����
    
    lngCriticalValues As Long                           'Σ��ֵ
    lngConformDetermine As Long                         '�������
    strImageLevel As String                             'Ӱ�������ȼ���
    strReportLevel As String                            '���������ȼ���
    lngImageLevel As Long                               'Ӱ�������ж�
    lngReportLevel As Long                              '���������ж�
    lngHintType As Long                                 '��Ͻ����ʾ����
    
    blnIsPetitionScan As Boolean                        '�Ƿ��������뵥ɨ��
    blnChangeUser As Boolean                            '�Ƿ������û�����
    blnSwitchUser As Boolean                            '�Ƿ������û��л�
    
    lngMoneyExeModle As Long                            '����ִ��ģʽ
    
    lngListColorMark As Long                            'Ϊ0ʱ����б�ǰ��ɫ��Ϊ1ʱ����б���ɫ
    blnNameColColorCfg As Boolean                       '�Ƿ���ݲ������������б���������ɫ
    blnOrdinaryNameColColorCfg As Boolean               'ȱʡ���͵Ĳ����Ƿ���ݲ�����������������ɫ
    
    blnAutoSendWorkList As Boolean                      '�Ƿ񱨵�ʱ�Զ�����WorkList
    
    '״̬����
    lngEnregAfterTimeLen As Long                        '�ǼǺ�����
    lngCheckInAfterTimeLen As Long                      '����������
    lngStudyAfterTimeLen As Long                        '��������
    lngReportAfterTimeLen As Long                       '���������
    lngAuditAfterTimeLen As Long                        '��˺�����
    
    blnAutoPrint As Boolean    '�������Զ���ӡ���뵥
    
    blnShowImgAfterReport As Boolean                    '����ʱ��Ƭ
End Type


'��Ƶ�ɼ��¼���Ϣ
Private Type TVideoEventInf
    vetEventType As TVideoEventType
    lngAdviceId As Long
    lngSendNo As Long
    strOtherInf As String
End Type

'��Ƶ�ɼ���Ϣ����
Private Type TCaptureMsgInf
    lngMsg As Long
    lngVirtualKey As Long
    lngScanKey As Long
    lngFlags As Long
End Type


'ID_���ҷ�ʽ+100֮����7������Ϊ���ҷ�ʽѡ���
'ID_Ӱ�����֮����40��������ΪӰ����𣬴�4021-4060
Private Enum FilterID
    ID_��Դ = 4000: ID_���� = 4001: ID_סԺ = 4002: ID_��� = 4003: ID_���� = 4004
    ID_���� = 4005: ID_�ѽ� = 4006: ID_δ�� = 4007: ID_���� = 4008: ID_�޷� = 4009: ID_���� = 4018
    ID_״̬ = 4010: ID_�Ǽ� = 4011: ID_���� = 4012: ID_��� = 4013: ID_���� = 4014: ID_��� = 4015: ID_���� = 4016: ID_��� = 4017
    ID_����ֵ = 4020: ID_��ʼ���� = 4021: ID_����סԺ = 4022: ID_���ҷ�ʽ = 4023
    
    ID_Ӱ����� = 4030
    
    ID_������� = 4100
    ID_�������_���� = 4101: ID_�������_���� = 4102: ID_�������_ϸ�� = 4103: ID_�������_ʬ�� = 4104: ID_�������_���� = 4105: ID_�������_����ʯ�� = 4106
        
    ID_Ӱ��ִ�м� = 4110
    
    ID_��鲿λ = 4210 '4210--4500
End Enum

Private mblncmd���� As Boolean, mblncmdסԺ As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean
Private mblncmd�ѽ� As Boolean, mblncmdδ�� As Boolean, mblncmd���� As Boolean, mblncmd�޷� As Boolean, mblncmd���� As Boolean
Private mblncmd�Ǽ� As Boolean, mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean
Private mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd��� As Boolean

Private mblncmd���� As Boolean


Private mblncmd���� As Boolean
Private mblncmd����ʯ�� As Boolean
Private mblncmdϸ�� As Boolean
Private mblncmd���� As Boolean
Private mblncmdʬ�� As Boolean
Private mblncmd���� As Boolean


Private mintcmdӰ����� As Integer      '0��ʾû��ѡ��Ӱ������������ֱ�ʾѡ���Ӱ����������
Private mblncmdӰ�����() As Boolean    '���浱ǰѡ���Ӱ������Ƿ�ѡ��

Private mintcmdӰ��ִ�м� As Integer    '��ѡ�����Ҫ���˵�Ӱ��ִ�м�������ֻ��Ϊ0ʱ���Ų���Ҫ����ִ�м����
Private mblncmdӰ��ִ�м�() As Boolean

Private mstrcmd��λ���� As String
Private mstrcmd��λ As String

Private mstrFirstTab As String '�״���ʾ��ҳ��

Private mintToolBarWriteReg As Integer        '������ע���״ֵ̬


Private mstrPrivs As String, mlngModule As Long              'ģ��ţ���ģ��Ȩ��


'�Ӵ������
Private WithEvents mobjEvent As clsEvent            '�¼��������
Attribute mobjEvent.VB_VarHelpID = -1
Private WithEvents mfrmRISRequest As frmRISRequest
Attribute mfrmRISRequest.VB_VarHelpID = -1
'��Ϣ��������
Private WithEvents mobjMsgCenter As clsPacsMsgProcess
Attribute mobjMsgCenter.VB_VarHelpID = -1

'����ģ�������ˢ��ģʽ�����������
'1.����ģ��ֻҪ���ڣ�ǿ�ƶ����е����ݽ���ˢ��
'2.����ģ������ʾʱ���Ŷ����е����ݽ���ˢ��
'3.����ģ����������ݱ仯ʱ����ʾ��ģ���ǵ�ǰģ�飬�Ŷ����е����ݽ���ˢ��

Private mfrmWork_PacsImg As frmWork_Image           'Ӱ���Ӵ���
Attribute mfrmWork_PacsImg.VB_VarHelpID = -1
Private mobjWork_Pathol As clsWorkModule_Pathol     '�������ģ��
Private mobjWork_His As clsWorkModule_His           'HIS���ģ��

Private mobjWork_ActiveVideo As Object  ' zl9PacsCapture.clsPacsCapture  '��Ƶ�ɼ�ģ��
Attribute mobjWork_ActiveVideo.VB_VarHelpID = -1
Private WithEvents mobjWork_Report As clsWorkModule_Report     '����ģ��
Attribute mobjWork_Report.VB_VarHelpID = -1
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer            '��Ƭվ����
Attribute mobjPacsCore.VB_VarHelpID = -1
Private WithEvents mobjQueue As frmWork_Queue  'zlQueueManage.clsQueueManage          '�Ŷӽк�
Attribute mobjQueue.VB_VarHelpID = -1

Private mfrmPatholSpecimen As frmPatholSpecimen              '�걾����

Private mfrmPACSFilter As frmPACSFilter

Private mclsCISKernel As clsCISKernel 'ֻ���˲鿴��鵥����

'���ڱ���
Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrCur���� As String                               '��ǰ���� ����-����
Private mstrCanUse���� As String                            '��ǰ���ÿ���  ID_����-����
Private mlngFilterTab As Long                               '����tabҳ
Private mblnInitOk As Boolean, mblnvsRefresh As Boolean     '��ʼ�����,װ�ر��
Private mblnLoadSubFrom As Boolean                          '�Ƿ����ڼ����Ӵ���
Private mblnAllDepts As Boolean                             '�Ƿ�ѡ��ȫ������
Private mstrCanUse����IDs As String                         '��ǰ���õĿ���ID�����á������ָ�������ֱ����ΪSQL��ѯ����
Private mlngSortCol As Long                                 '�����б��У���ǰ�����������
Private mintSortOrder As Integer                            '�����б��У���ǰ��������ķ�ʽ
Private mblnMenuDownState As Boolean                        '����˫����������������
Private mblnIsLoadPatholModule As Boolean                   '�Ƿ������˲���ģ��
Private mblnFormLoadState As Boolean

Private mblnIsPrintMode As Boolean                          '�Ƿ����嵥��ӡ

Private mstrDefaultPatientType As String                    'ȱʡ��������
Private mlngOldAdviceId As Long                             'ǰһ��ѡ��ļ���ҽ��ID

Private mstrRPTExecutor As String                           '����ѡ��ı�����
Private mrtReportType As ReportType

'���̿��Ʊ���
Private mSysPar As TSystemPar                               'ϵͳ����

'Private mlngOldSameTime As Long                             '�л�����ǰ��ǰ���ҷ��ŷ�ʽ��0����ͽ�Ƭ�ֱ𷢷� 1 ����ͽ�Ƭͬʱ����
Private mblnObserve As Boolean                              '�Ƿ��й�Ƭ����Ȩ��   true��  false��
Private mblnSetXWParam As Boolean                           '�Ƿ��С�Ӱ���豸Ŀ¼��Ȩ�ޣ�����У��������������PACS�Ĳ���
Private mintImgCount As Integer                             '��ɨ�����뵥����

Private mAstr��������() As String       '�������ƣ�ִ�м������

Private WithEvents mobjCaptureHot As zl9PacsControl.clsHookKey
Attribute mobjCaptureHot.VB_VarHelpID = -1
Private mVideoEventInf As TVideoEventInf
Private mblnUseActivexCapture As Boolean                            '�Ƿ�ʹ��ActivexExe����Ƶ�ɼ���ʽ
Private mstrCaptureHot As String                                    '�ɼ��ȼ�����
Private mstrCaptureAfterHot As String                               '��̨�ɼ��ȼ�����
Private mstrCaptureAfterTagHot As String                            '��Ǹ����ȼ�����
Private mCaptureMsg As TCaptureMsgInf
Private mobjSquareCard As Object

'�������ز���
Private mstrSelQueueRooms As String                         'ֻ����ִ�м��ڵĲ���
Private mstrAllQueueRooms As String

Private mblnMoved As Boolean                                '��ǰʱ������Ƿ�ת�ƹ�
Private mstrWorkModule As String

Private mblnPopChangGuiWindow As Boolean
Private mblnPopBingDongWindow As Boolean
Private mblnPopXiBaoWindow As Boolean
Private mblnPopHuiZhenWindow As Boolean
Private mblnPopShiJianWindow As Boolean
Private mblnPopKuaiShuWindow As Boolean

Private SQLCondition As Type_SQLCondition

Private mblnAssignment As Boolean
Private mstrFindWay As String
Private mstrLocateWay As String
Private mlngLocateFindType As Long
Private mstrAllExamineRoomCfg As String    '���п���ִ�м�ѡ�����
Private mstrCurExamineRoomCfg As String    '��ǰ����ִ�м�ѡ�����

Private mbytFontSize As Byte  '�����С
Private mbyrFontState As Byte '����״̬�������ж��Ƿ�����ؼ�λ��


Private mcurAdviceInf As TAdviceInf          '����Ӽ���б������ʷ�б���ѡ��ĵ�ǰ�����Ϣ
Private mListAdviceInf As TAdviceInf         'ֻ����Ӽ���б���ѡ��ļ����Ϣ


'��ʷ��¼����ʾ
Private mblnIsHistory As Boolean


'˫�û���¼
Private mcnOracleHIS As New ADODB.Connection    '��¼HIS����̨��½ʱʹ�õ����ݿ����Ӵ�
Private mstrUserNameHIS As String               '��¼HIS����̨��½ʱʹ�õ��û���
Private mstrUserIDHIS As String                 '��¼HIS����̨��¼ʱʹ�õ��û�ID
Private mstrUserNameNew As String               '��¼˫�û���½�ĵڶ����û���
Private mstrUserIDNew As String                 '��¼˫�û���¼�ĵڶ����û�ID
Private mblnCnOracleIsHIS As Boolean            '��ǰ���ݿ������Ƿ�HIS����̨������
Private mintChangeUserState As Integer          '��¼�û������������1- ͳһ��2-����

'�ղع���
Private mlngShareFatherID As Long
Private mlngCollectionFatherID As Long

Private mblnIsIntegratedQuery As Boolean        '�Ƿ��ڽ����ۺϲ�ѯ
Private mlngDefQuerySchemeId As Long            'Ĭ�ϲ�ѯ����id
Private mlngSysQuerySchemeId As Long            'ϵͳ��ѯ����id
Private blnIsLoading As Boolean

Dim mlngTempCharged As Long

Private mblnIsCallModuleRefresh As Boolean          '�Ƿ����ģ��ˢ�²���
Private mblnAutoRefreshList As Boolean          '�Ƿ��Զ�ˢ�¼���б�


Public Sub ShowStation(ByVal lngModule As Long, owner As Object)
    
    mblnInitOk = False
    mblnLoadSubFrom = False
    mlngModule = lngModule
    mblnUseActivexCapture = False
    mblnAutoRefreshList = False
    mblnIsIntegratedQuery = False
    
    '���������㲿��
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    
    '��ʼ�������㲿��
    mobjSquareCard.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    
    PatiIdentify.zlInit Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjSquareCard, InitCardType(Replace(IIf(mlngLocateFindType = TLocateFindType.lftLocate, CONST_STR_LOCAL_CARD_TYPE, CONST_STR_FIND_CARD_TYPE), "[------]", GetStudyNumberDisplayName))

    If lngModule = G_LNG_VIDEOSTATION_MODULE Or lngModule = G_LNG_PATHOLSYS_NUM Then
        mblnUseActivexCapture = GetSetting("ZLSOFT", "����ģ��", "UseActiveVideo", "1")
        Call SaveSetting("ZLSOFT", "����ģ��", "UseActiveVideo", mblnUseActivexCapture)
    End If
    
    Call WriteLog("ShowStation -> Step 1������Ӱ�������ڳ�ʼ�����̡�")
    
    If Not mblnFormLoadState Then Call InitForm
    
    Call WriteLog("ShowStation -> Step 2")
    
    '����ʾ����ǰϵͳ����
    Me.Show , owner
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    DoEvents
    
    Call WriteLog("ShowStation -> Step 3����ʼ��������ģ�顣")
    '��������Ĺ���ģ��
    Call Me.InitSubForm
    

    DoEvents
    
    Call WriteLog("ShowStation -> Step 4��������ʾ��ģ�顣")
    
    If Not TabWindow.Selected Is Nothing Then
        Call ConfigSubForm(TabWindow.Selected)
    End If
    
    mblnInitOk = True
    
    Call WriteLog("ShowStation -> Step 5��ˢ�������б�")
    
    'ˢ�¼������
    Call Me.RefreshList
    
    DoEvents
    
    Call WriteLog("ShowStation -> Step 6������ģ��˵���")
    '����ģ��˵�
    Call CreateWorkModuleMenu
    
    Call WriteLog("ShowStation -> Step 7������Ӱ�������ڳ�ʼ�����̡�")
End Sub


Private Sub Menu_File_Excel_click()
'����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
'����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
'       lngSelectedRow,��¼���ô�ӡ����ǰ��ѡ���У����嵥�رպ�ָ�
On Error GoTo ErrHandle
    Dim bytMode As Byte
    Dim lngSelectedRow As Long
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = ufgStudyList.DataGrid
    objPrint.Title.Text = "��鲡���嵥"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & zlDatabase.Currentdate())
    Call objPrint.BelowAppRows.Add(objAppRow)

    '�� �Ƿ��Ǵ�ӡ�嵥������ֵ
    mblnIsPrintMode = True
    '�õ���ӡ�嵥ǰ�ĵ�ǰѡ����
    lngSelectedRow = ufgStudyList.SelectionRow
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    
    '��ӡ��Ԥ�������� �ָ�ѡ����
    ufgStudyList.DataGrid.Row = lngSelectedRow
    mblnIsPrintMode = False
    
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_RichEPR(ByVal cbrID As Long)
'�Զ��򿪱���༭����ͬʱ������PACS����༭���͵��Ӳ����༭��
On Error GoTo ErrHandle
    Dim cbrControl As CommandBarControl, i As Long
    
    '���û��ѡ�������ݣ���ֱ���˳�ִ��
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '����ҳ�治�ɼ�ʱ��ִ���κβ���
    If TabWindow.Selected.tag <> "������д" Then
        For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
            If TabWindow(i).tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.tag <> "������д" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    '�ҵ�����ҳ�棬�ٴ��������ҳ��
    With ufgStudyList
        'ˢ��Ƕ��ҳ������
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.zlUpdateAdviceInf(mListAdviceInf.lngAdviceId, mListAdviceInf.lngPatID, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, mListAdviceInf.intMoved = 1)
            Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, ufgStudyList, mblnIsHistory, mListAdviceInf.blnCanPrint, mListAdviceInf.strDoDoctor, mListAdviceInf.strStudyUID)
            
            Call mobjWork_Report.zlRefreshFace
        End If
    End With
    
    '�жϰ���������
    Set cbrControl = Me.cbrMain.FindControl(, conMenu_PacsReport_Open + 1000000)
    
    If cbrControl Is Nothing Then
        Set cbrControl = Me.cbrMain.FindControl(, cbrID + 1000000)
        If cbrControl Is Nothing Then Exit Sub
    End If
    
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    '����˫����ť����ı���������Ҫ�������ó�False����Ϊ��������ʱ�򿪱��洰�塱ʱ��ʵ���ϴ˱���ΪTrue
    mblnMenuDownState = False
    Call cbrMain_Execute(cbrControl)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_File_Parmeter_click()
On Error GoTo ErrHandle
    With frmTechnicSetup
        .mlngModul = mlngModule
        .mlng����ID = mlngCur����ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        
        If .mblnOK Then
            InitLocalPars
            
            If Not mobjWork_Report Is Nothing Then
                '���¼��غͱ�����ص����ò���
                Call mobjWork_Report.InitReportParameter
            End If
            
            Call RefreshList
        End If
    End With
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'��ʾ��ݷ�ʽ����
Private Sub Menu_File_ShortcutSet_click()
    Dim frmShortcut As New frmShortcutConfig
    
On Error GoTo ErrHandle
    Dim lngCount As Long
    
    Call frmShortcut.ShowShortcutConfig(App.ProductName, mlngModule, Me)
        
    If frmShortcut.blnIsOk Then
        'ɾ�����ڵĹ������������˵���
        Call LockWindowUpdate(Me.hWnd)
        
        For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
            cbrMain.ActiveMenuBar.Controls(lngCount).Delete
        Next
        
        For lngCount = cbrMain.Count To 2 Step -1
            cbrMain(lngCount).Delete
        Next
    
        Call InitCommandBars
        Call CreateWorkModuleMenu
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        
        Call LockWindowUpdate(0)
    End If
    
    
    Call Unload(frmShortcut)
    Set frmShortcut = Nothing
Exit Sub
ErrHandle:
    Call Unload(frmShortcut)
    Set frmShortcut = Nothing
End Sub


Private Sub Menu_Help_About_click()
On Error GoTo ErrHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo ErrHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo ErrHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo ErrHandle
    zlMailTo hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ȡ������()
'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    lngResult = -1
    
    '�����ģ���Ϊ1298��RIS����վ����������������ݿ��ѯ��ƥ���ͼ���¼
    If mlngModule = G_LNG_PACSSTATION_MODULE And mListAdviceInf.intImageLocation = 1 Then
        lngResult = XWShowMatched(Me, mListAdviceInf.lngAdviceId)
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mListAdviceInf.lngAdviceId, mstrPrivs, mblnMoved, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur����ID, 1
        
        If frmSelectMuli.mblnOK = True Then lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call AfterReleationImage(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, 1, True)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_�ޱ������()
'ֻ�н����еı�����Բ����ò˵�,��Ϊ��ʱ��û��ǩ��
On Error GoTo ErrHandle
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mListAdviceInf.strReportDoctor <> "" Or mListAdviceInf.strReportOperation <> "" Then
        If MsgBoxD(Me, "�Ƿ��ޱ���ֱ�����,ֱ����ɽ�ɾ������д�ı���!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    If mSysPar.blnFinishCommit And InStr(mstrPrivs, "������") > 0 Then '�ޱ�����ɺ������ٴ�ȷ�����,����Ҫ�м����ɵ�Ȩ��
        '�˹���,��״̬=6,���ұ���ID��Ϊ�ս�ɾ�����Ӳ�����¼
        
        If bln����δ���(mListAdviceInf.lngPageID, mListAdviceInf.lngPageID, mListAdviceInf.lngAdviceId, mListAdviceInf.lngPatientFrom) Then
            'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
            MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵�������ɣ�", vbExclamation, gstrSysName
        Else
            If mrtReportType = �����ĵ��༭�� Then
                gstrSQL = "ZL_Ӱ����_״̬����(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",'',6,1,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            Else
                gstrSQL = "ZL_Ӱ����_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",6,1,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            End If
        End If
    Else
        If mrtReportType = �����ĵ��༭�� Then
            gstrSQL = "ZL_Ӱ����_״̬����(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",'',5,1,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
        Else
            gstrSQL = "ZL_Ӱ����_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",5,1,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
        End If
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ı������")
        
    'ȡ���Ŷ���Ϣ
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(mListAdviceInf.lngAdviceId)
    End If
    
        
    If mSysPar.blnFinishCommit Then
        Call StateCheck(6)
    Else
        Call StateCheck(5)
    End If
    
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Edit_�ޱ������()
On Error GoTo ErrHandle
    Dim rsTemp As ADODB.Recordset

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫ���˸�������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub

    '�����ͼ������˵����Ѽ�顱��������˵����ѱ�����
    gstrSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ͼ��", mListAdviceInf.lngAdviceId)
    
    gstrSQL = "ZL_Ӱ����_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & "," & IIf(Nvl(rsTemp!���uid) = "", 2, 3) & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        
    Call StateCheck(2)

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetAdviceDetailInf(Optional ByVal lngAdviceId As Long = 0) As TAdviceInf
'����ҽ��id��ȡ��ϸ��ҽ����Ϣ
'lngAdviceId:���Ϊ0�����ȡ��ǰ�б�ѡ�еļ��ҽ����Ϣ

    Dim strSql As String
    Dim strSQLBak As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long
    Dim i As Long
    
    lngIndex = -1
    
    '����Ĭ�ϵ�ҽ����Ϣ
    GetAdviceDetailInf = GetNullAdviceInf
    
    
    '����б��м��������ݣ�����б��ж�ȡҽ����Ϣ
    If ufgStudyList.GridRows > 1 And ufgStudyList.GridCols > 1 Then
        If lngAdviceId <= 0 Then
            lngIndex = ufgStudyList.SelectionRow
        Else
            For i = 1 To ufgStudyList.GridRows - 1
                If Val(ufgStudyList.KeyValue(i)) = lngAdviceId Then
                    lngIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
    
    
    If lngIndex <= 0 And lngAdviceId > 0 Then
    
        '�����ݿ��в�ѯָ��ҽ��id����ϸ��Ϣ
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            strSql = "Select A.ID,A.����, A.���˿���id, A.����ҽ��,A.������Դ, A.ҽ������, Nvl(A.Ӥ��, 0) Ӥ��,A.����id,e.��ǰ����,e.סԺ��,e.�����,decode(A.������Դ,2,F.��������,E.��������) AS ��������, " & vbNewLine & _
                    " A.��ҳid, A.�Һŵ�, B.����,B.Ӱ�����,B.��鼼ʦ, B.���uid,B.ͼ��λ��,B.������,B.�������,B.����ID, C.����, D.���ͺ�,D.ִ��״̬,D.ִ�й���,D.ִ�м�, 0 as ת��,A.ִ�п���ID " & vbNewLine & _
                    " From ����ҽ����¼ A, Ӱ�����¼ B, ���ű� C, ����ҽ������ D,������Ϣ E,������ҳ F " & vbNewLine & _
                    " Where A.ID = B.ҽ��id And A.���˿���id = C.ID And A.ID = D.ҽ��id and A.����ID=E.����ID and A.����ID = F.����ID(+) And A.��ҳID+0 = F.��ҳID(+) and A.ID = [1]"
        Else
            strSql = "Select A.ID,A.����, A.���˿���id, A.����ҽ��,A.������Դ, A.ҽ������, Nvl(A.Ӥ��, 0) Ӥ��, A.����id,F.��ǰ����,F.סԺ��,F.�����,decode(A.������Դ,2,G.��������,F.��������) AS ��������, " & vbNewLine & _
                    " A.��ҳid, A.�Һŵ�, E.�����,B.Ӱ�����,B.��鼼ʦ, B.���uid,B.ͼ��λ��,B.������,B.�������,B.����ID, C.����, D.���ͺ�,D.ִ��״̬,D.ִ�й���,D.ִ�м�,0 as ת��,A.ִ�п���ID " & vbNewLine & _
                    " From ����ҽ����¼ A, Ӱ�����¼ B, ���ű� C, ����ҽ������ D, ��������Ϣ E, ������Ϣ F,������ҳ G " & vbNewLine & _
                    " Where A.ID = B.ҽ��id And A.���˿���id = C.ID And A.ID = D.ҽ��id and A.ID=E.ҽ��ID and A.����ID=F.����ID and A.����ID = G.����ID(+) And A.��ҳID+0 = G.��ҳID(+) and A.ID = [1]"
        End If
                    
        strSQLBak = strSql
        strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
        strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
        strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
'        strSQLBak = Replace(strSQLBak, "��������Ϣ", "H��������Ϣ")    '��������Ϣ��10.32.0֮�󲻲���ת��
'        strSQLBak = Replace(strSQLBak, "������Ϣ", "H������Ϣ")            '������Ϣ��δ����ת��
        
        strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
        
        strSql = strSql & vbNewLine & " Union ALL " & strSQLBak
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�����μ�¼��Ϣ", lngAdviceId)
        
        If Not rsTemp.EOF Then
            With GetAdviceDetailInf
                .lngPatID = Val(Nvl(rsTemp!����ID))
                .lngAdviceId = lngAdviceId
                .lngSendNo = Val(Nvl(rsTemp!���ͺ�))
                .lngPageID = Val(Nvl(rsTemp!��ҳID))
                .lngPatDept = Val(Nvl(rsTemp!���˿���ID))
                .strPatientName = Nvl(rsTemp!����)
                .lngUnit = .lngPatDept
                .blnCanPrint = True
                
                .lngPatientFrom = Val(Nvl(rsTemp!������Դ, 3))
                .strPatientType = Nvl(rsTemp!��������)
                
                .blnIsInsidePatient = (.lngPatientFrom = 1) Or (.lngPatientFrom = 2)
                .intMoved = Val(Nvl(rsTemp!ת��))
                .intState = Val(rsTemp!ִ��״̬)
                .intStep = Val(Nvl(rsTemp!ִ�й���))
                .strRegNo = Val(Nvl(rsTemp!�Һŵ�))
                .lngRegId = getRegID(.strRegNo)
                .strStudyUID = Val(Nvl(rsTemp!���uid))
                .lngExeDepartmentId = Val(Nvl(rsTemp!ִ�п���ID))
                .strDoDoctor = Nvl(rsTemp!��鼼ʦ)
                .strExeRoom = Nvl(rsTemp!ִ�м�)
                .strStudyNum = Nvl(rsTemp(GetStudyNumberDisplayName))
                .strBedNum = Nvl(rsTemp!��ǰ����)
                .lngBaby = Val(Nvl(rsTemp!Ӥ��))
                .strPatientDepartment = Nvl(rsTemp!����)
                .lngMarkNum = IIf(.lngPatientFrom = 1, Val(Nvl(rsTemp!�����)), IIf(.lngPatientFrom = 2, Val(Nvl(rsTemp!סԺ��)), 0))
                
                .strReportDoctor = Nvl(rsTemp!������)
                .strReportOperation = Nvl(rsTemp!�������)
                
                .lngLinkId = Val(Nvl(rsTemp!����ID))
                
                .strImgType = Nvl(rsTemp!Ӱ�����)
                .intImageLocation = Nvl(rsTemp!ͼ��λ��)
            End With
        End If
        
        Exit Function
    End If
    
    '�����ǰ�б���û�м�飬��ҽ��idΪ0�����˳��ú���
    If lngIndex <= 0 And lngAdviceId <= 0 Then Exit Function
    
    
    '�ӽ����ж�ȡҽ��id��ص���ϸ��Ϣ
    With GetAdviceDetailInf
        .lngPatID = Val(ufgStudyList.Text(lngIndex, "����ID"))
        .lngPageID = Val(ufgStudyList.Text(lngIndex, "��ҳID"))
        .lngAdviceId = Val(ufgStudyList.KeyValue(lngIndex))
        .lngSendNo = Val(ufgStudyList.Text(lngIndex, "���ͺ�"))
        .lngPatDept = Val(ufgStudyList.Text(lngIndex, "���˿���ID"))
        .strPatientName = ufgStudyList.Text(lngIndex, "����")
        .strRegNo = ufgStudyList.Text(lngIndex, "�Һŵ�")
        .lngRegId = getRegID(.strRegNo)
        .intMoved = Val(ufgStudyList.Text(lngIndex, "ת��"))
        .intState = IIf(ufgStudyList.Text(lngIndex, "������") = "�Ѿܾ�", 2, IIf(ufgStudyList.Text(lngIndex, "������") = "�����", 1, 3))
        .intStep = Val(ufgStudyList.Text(lngIndex, "���״̬")) '��ȡִ�й���
        .lngUnit = Val(ufgStudyList.Text(lngIndex, "��ǰ����ID"))
        
        If mrtReportType = �����ĵ��༭�� Then
            If ufgStudyList.Text(lngIndex, "����") = " " Or Val(ufgStudyList.Text(lngIndex, "��ɫͨ��")) = 1 Then
                .blnCanPrint = True
            Else
                .blnCanPrint = False
            End If
        Else
            .blnCanPrint = IIf(mSysPar.blnCanPrint, IIf(Val(ufgStudyList.Text(lngIndex, "����")) = 1, ufgStudyList.Text(lngIndex, "������") <> "", ufgStudyList.Text(lngIndex, "������") <> ""), True)
        End If
        
        .strStudyUID = ufgStudyList.Text(lngIndex, "���UID")
        .lngExeDepartmentId = Val(ufgStudyList.Text(lngIndex, "ִ�п���ID"))
        .strDoDoctor = ufgStudyList.Text(lngIndex, "��鼼ʦ")
        
        '��ִ��ˢ�²����󣬵�Ԫ���flexcpdata���ݲ��������ͱ�ˢ�£�ֻ��ͨ����Ӧ����ʾ�ı���ֵ����ת����flexcpdataֵ�ĸ������첽�¼�����
        .lngPatientFrom = Decode(ufgStudyList.Text(lngIndex, "��Դ"), "��", 1, "ס", 2, "��", 3, 4)
        
        .strPatientType = ufgStudyList.Text(lngIndex, "��������")
        
        .blnIsInsidePatient = (.lngPatientFrom = 1) Or (.lngPatientFrom = 2)
        .strExeRoom = ufgStudyList.Text(lngIndex, "ִ�м�")
        .strStudyNum = ufgStudyList.Text(lngIndex, GetStudyNumberDisplayName)
        .strBedNum = ufgStudyList.Text(lngIndex, "����")
        .lngMarkNum = Val(ufgStudyList.Text(lngIndex, "��ʶ��"))
        .lngBaby = 0
        
        .strReportDoctor = ufgStudyList.Text(lngIndex, "������")
        .strReportOperation = ufgStudyList.Text(lngIndex, "�������")
        
        .lngLinkId = Val(ufgStudyList.Text(lngIndex, "����ID"))
        .strImgType = ufgStudyList.Text(lngIndex, "Ӱ�����")
        .intImageLocation = Val(ufgStudyList.Text(lngIndex, "ͼ��λ��"))
        
        strSql = "Select ���� From ���ű� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˿���", .lngPatDept)
        
        .strPatientDepartment = ""
        If rsTemp.RecordCount > 0 Then .strPatientDepartment = Nvl(rsTemp!����)

    End With
        
End Function

Private Function getRegID(ByVal strRegNo As String) As Long
'����:��ȡ�Һ�id
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    getRegID = 0
    
    strSql = "select id from ���˹Һż�¼ where no=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, strRegNo)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    getRegID = Nvl(rsTemp!ID, 0)
    
    Exit Function

ErrHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function IsAlreadyInputQuality(ByVal lngAdviceId As Long) As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    IsAlreadyInputQuality = False
    
    strSql = "select �ۺ����� from ��������Ϣ where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If Nvl(rsData!�ۺ�����) <> "" Then IsAlreadyInputQuality = True
    
End Function

Private Sub Menu_Manage_����������(Optional lngAdviceId As Long = 0, Optional blnRefresh As Boolean = True, Optional strReportId As String = "")
'�������������̵��ã���ʱ������ҽ��ID������ҪȨ���ж�
On Error GoTo ErrHandle
    Dim strSql As String
    Dim curAdviceInf As TAdviceInf
    Dim intState As Integer
    
    If InStr(mstrPrivs, "������") <= 0 Then Exit Sub
    
    curAdviceInf = GetAdviceDetailInf(lngAdviceId)
    
    If curAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If lngAdviceId = 0 Then
    '����ǻ��б���δ���ʱ��ɼ��
        If mrtReportType = �����ĵ��༭�� Then
            intState = getStudyStateRich(curAdviceInf.lngAdviceId, strReportId, False)
        Else
            intState = getStudyState(curAdviceInf.lngAdviceId, False)
        End If
        
        If intState = 4 Then
            If MsgBoxD(Me, "���б���û��д�꣬�����ʱ��ɼ�飬��Ҫ�С���¼���桱Ȩ�޵��˲��ܼ�����д����!" & vbCrLf & "ȷ��Ҫ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    '����ǲ���ϵͳ��������ʱ������Ҫ�����������ƴ���
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If (mblnPopChangGuiWindow And ufgStudyList.CurText("������") = "����") _
            Or (mblnPopKuaiShuWindow And ufgStudyList.CurText("������") = "����ʯ��") _
            Or (mblnPopBingDongWindow And ufgStudyList.CurText("������") = "����") _
            Or (mblnPopXiBaoWindow And ufgStudyList.CurText("������") = "ϸ��") _
            Or (mblnPopHuiZhenWindow And ufgStudyList.CurText("������") = "����") _
            Or (mblnPopShiJianWindow And ufgStudyList.CurText("������") = "ʬ��") Then
            
            If Not IsAlreadyInputQuality(curAdviceInf.lngAdviceId) Then
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlMenu.zlExecuteMenu(conMenu_Pathol_Quality_Manage)
            End If
            
            If Not IsAlreadyInputQuality(curAdviceInf.lngAdviceId) Then
                Call MsgBoxD(Me, "δ¼��������������ִ����ɲ�����", vbInformation, GetWindowCaption)
                Exit Sub
            End If
            
        End If
    End If
    
    If bln����δ���(curAdviceInf.lngPatID, curAdviceInf.lngPageID, Nvl(curAdviceInf.lngAdviceId), curAdviceInf.lngPatientFrom) Then
        'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
        MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵���������ɣ�", vbExclamation, gstrSysName
        Exit Sub
    Else
        If mrtReportType = �����ĵ��༭�� Then
            strSql = "Zl_Ӱ����_״̬����(" & curAdviceInf.lngAdviceId & "," & curAdviceInf.lngSendNo & ",'',6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "�ı������")
        Else
            strSql = "ZL_Ӱ����_STATE(" & curAdviceInf.lngAdviceId & "," & curAdviceInf.lngSendNo & ",6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "�ı������")
        End If
        
        If mlngModule = G_LNG_PATHOLSYS_NUM Then
            gstrSQL = "Zl_������_���(" & curAdviceInf.lngAdviceId & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
        End If
    End If

    
    'ȡ���Ŷ���Ϣ
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(mListAdviceInf.lngAdviceId)
    End If

    If blnRefresh Then Call StateCheck(6)
        
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow
    
    '���ͼ�������Ϣ
    Call mobjMsgCenter.Send_Msg_StudyComplete(mListAdviceInf.lngAdviceId, strReportId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_ȡ��������()
On Error GoTo ErrHandle
    Dim strSql As String
    Dim intState As Integer

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    If mListAdviceInf.intMoved = 1 Then
        MsgBoxD Me, "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If CheckIsArchived(mListAdviceInf.lngAdviceId) Then
            MsgBoxD Me, "�ò��˵ĵ����Ѿ��鵵�������������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mrtReportType = �����ĵ��༭�� Then
        intState = getStudyStateRich(mListAdviceInf.lngAdviceId, "", True)
        strSql = "ZL_Ӱ����_״̬����(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",''," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & To_Date(zlDatabase.Currentdate) & ")"
    Else
        intState = getStudyState(mListAdviceInf.lngAdviceId, True)
        strSql = "ZL_Ӱ����_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & To_Date(zlDatabase.Currentdate) & ")"
    End If
    
    zlDatabase.ExecuteProcedure strSql, "ȡ��������"
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSql = "Zl_������_ȡ�����(" & mListAdviceInf.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "������ȡ�����")
    End If
    
    Call StateCheck(intState)
    
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow
    
    '���ͼ�鳷�������Ϣ
    Call mobjMsgCenter.Send_Msg_CancelComplete(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function CheckIsArchived(lngAdviceId As Long) As Boolean
 '���ò��˵����Ƿ��Ѿ��鵵���ѹ鵵�ļ�飬��Ҫ��������ȡ�����  0--δ�鵵  1--�ѹ鵵
 On Error GoTo ErrHandle
 
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select distinct c.����״̬ as ״̬ from ��������Ϣ a,����鵵��Ϣ b,��������Ϣ c where a.����ҽ��ID = b.����ҽ��ID and b.����id = c.id and a.ҽ��ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ��ѹ鵵", lngAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        CheckIsArchived = False
        Exit Function
    End If
    
    CheckIsArchived = IIf(Nvl(rsTemp!״̬, 0) = 1, True, False)
Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Menu_Manage_CriticalMark(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim intCritical As Integer
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_CriticalValues, conMenu_Manage_Critical
            intCritical = 1
        Case conMenu_Manage_Normal
            intCritical = 0
    End Select
    
    With ufgStudyList
        If intCritical = 1 Then
            If lngID = conMenu_Manage_CriticalValues Then
                Call frmCriticalValues.ShowMe(mListAdviceInf.lngAdviceId, Me)
                If Not frmCriticalValues.mblnSave Then Exit Sub
            End If
            
            strSql = "zl_Ӱ����_Σ������(" & mListAdviceInf.lngAdviceId & ",1)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("Σ��")) = imgList.ListImages("Σ��").Picture
                .CurText("Σ��") = " "
                
            Menu_Manage_������� conMenu_Manage_Negative
            
            '����Σ��ֵ��Ϣ
            Call mobjMsgCenter.Send_Msg_Critical(mListAdviceInf.lngAdviceId)
        ElseIf intCritical = 0 Then
            If .CurText("Σ��") = "" Then Exit Sub
            If MsgBoxD(Me, "ȷ��Ҫȡ�����ˡ�" & .CurText("����") & "����Σ��״̬��", vbOKCancel, "Σ����������") = 2 Then Exit Sub
            
            strSql = "Zl_Ӱ��Σ��ֵ��¼_ȡ��(" & mListAdviceInf.lngAdviceId & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("Σ��")) = Nothing
                .CurText("Σ��") = ""
        End If
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "Σ��", intCritical)
    End With

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_�������(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim iResult As Integer
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_Negative
            iResult = 1
        Case conMenu_Manage_Positive
            iResult = 0
    End Select
    
    If mrtReportType = �����ĵ��༭�� Then
        Call mobjWork_Report.Menu_Manage_�������(mListAdviceInf.lngAdviceId, iResult)
        Exit Sub
    End If
    
    strSql = "ZL_Ӱ����_���(" & mListAdviceInf.lngAdviceId & "," & iResult & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "���������")
    
    With ufgStudyList
        If iResult = 1 Then
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("����")) = imgList.ListImages("����").Picture
            .CurText("����") = " "
        Else
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("����")) = Nothing
            .CurText("����") = ""
        End If
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "����", iResult)
    End With

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_��ɫͨ��(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim intResult As Integer
    Dim blnCanPrint As Boolean
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    
    strSql = "Zl_��ɫͨ��_Update(" & mListAdviceInf.lngAdviceId & ",'" & intResult & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "��ɫͨ��")
    
    With ufgStudyList
        .CurText("��ɫͨ��") = intResult
        
        If intResult = 1 Then
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("����")) = imgList.ListImages("��ɫͨ��").Picture
        Else
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("����")) = Nothing
        End If
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "��ɫͨ��", intResult)
    End With
    
    If mrtReportType = �����ĵ��༭�� Then
        blnCanPrint = ufgStudyList.CurText("����") = " " Or Val(ufgStudyList.CurText("��ɫͨ��")) = 1
        
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, ufgStudyList, mblnIsHistory, blnCanPrint, mcurAdviceInf.strDoDoctor, mcurAdviceInf.strStudyUID)
            Call mobjWork_Report.zlRefreshFace(True, False, False)
        End If
    End If

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_�������(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strResult As String
    Dim strSql As String
    Dim lngColIndex As Long

    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    Select Case lngID
        Case conMenu_Manage_FuHe
            strResult = "����"
        Case conMenu_Manage_JiBenFuHe
            strResult = "��������"
        Case conMenu_Manage_BuFuHe
            strResult = "������"
    End Select

    strSql = "Zl_�������_Update(" & mListAdviceInf.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "�������")
        
    With ufgStudyList
        .CurText("�������") = strResult
        
        lngColIndex = ufgStudyList.GetColIndex("�������")
         
        If strResult = "����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, .DataGrid.Row, lngColIndex) = vbGreen
        If strResult = "��������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, .DataGrid.Row, lngColIndex) = vbYellow
        If strResult = "������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, .DataGrid.Row, lngColIndex) = vbRed
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "�������", strResult)
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_�޸�()
On Error GoTo ErrHandle
    Dim strOldName As String
    Dim strOldRoom As String
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mListAdviceInf.lngSendNo
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mstrPatientName = mListAdviceInf.strPatientName
            .mintEditMode = IIf(mListAdviceInf.intStep > 1, 3, 1)  '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mstrCur���� = NeedName(mstrCur����)
'            .mlngQueueWay = mSysPar.lngQueueWay
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then
                strOldName = mListAdviceInf.strPatientName
                strOldRoom = mListAdviceInf.strExeRoom
                
                Call RefreshList(.mlngAdviceId, True) '�ɹ�����
                
                If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                    '����Ǳ������޸ģ��Ҹı���ִ�м䣬����Ҫ���½����Ŷ�
                    If .mintEditMode = 3 And .mlngResultState = 3 Then
                        If .mstrTechnicRoom <> strOldRoom Then
                            If .mstrTechnicRoom = "" Then
                                '���Ϊ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                                Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                            Else
                                '�����Ϊ�գ���д���Ӧ��ִ�м�����
                                strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                                strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                            End If
                            
                            Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                        Else
                            '������ʽ���޸ģ���ֻ���Ŷӽк��е������Ϣ���и���
                            If .mstrPatientName <> strOldName Then
                                Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = Val(ufgStudyList.CurText("���ͺ�"))
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mintEditMode = IIf(Val(ufgStudyList.CurText("���״̬")) > 1, 3, 1)  '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mintImgCount = mintImgCount
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";�걾����ģ��;") > 0, True, False)
            .InitMvar
            
            If .RefreshPatiInfor(False) = True Then  'ˢ�²���
                .mblnOK = False
                .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOK Then RefreshList '�ɹ�����
        End With
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub Menu_Manage_ModifBaseInfo()
''������Ϣ����
'On Error GoTo errHandle
'    Dim int���� As Integer
'    Dim str����ID As String
'
'    With mcurAdviceInf
'        int���� = Decode(.lngPatientFrom, 1, 1, 2, 2, 3, 3, 4, 4)
'
'        str����ID = Decode(.lngPatientFrom, 1, .lngRegId, 2, .lngPageID, 3, .lngAdviceID, 4, .strRegNo)
'
'        If zlDatabase.zlModiPatiBaseInfo(.lngPatID, str����ID, mlngModule, int����) Then RefreshList    '�ɹ�����
'    End With
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub Menu_Manage_���ƵǼ�()
On Error GoTo ErrHandle
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mstrCur���� = NeedName(mstrCur����)
            .mlngResultState = 0
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1), mblnAllDepts, mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo
            
            If .mlngResultState <> 0 Then '�ɹ�����
                Call StateCheck(2, .mlngAdviceId)
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mblnOK = False
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";�걾����ģ��;") > 0, True, False)
            .InitMvar
            
            If .CopyCheck(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo) = True Then  'ˢ�²���
                .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOK Then '�ɹ�����
                Call StateCheck(2, .mlngAdviceId)
            End If
        End With
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub AutoPrintApplication(lngAdviceId As Long, lngSendNo As Long, lngClinicId As Long, intSourceType As Long)
'����:�����������Զ���ӡ���뵥
Dim rsTemp As ADODB.Recordset, strBillNo As String, strExseNo As String, intExseKind As Integer

On Error GoTo errHand
    Dim strSql As String
    
    If Not mSysPar.blnAutoPrint Then Exit Sub
    
    strSql = "select NO,��¼���� from ����ҽ������ where ҽ��ID=[1] and ���ͺ�=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡNO", lngAdviceId, lngSendNo)
    If rsTemp.EOF Then Exit Sub
    
    strExseNo = rsTemp!NO: intExseKind = rsTemp!��¼����
    
    strSql = "Select B.ID, B.���" & vbNewLine & _
                "From ��������Ӧ�� A, �����ļ��б� B" & vbNewLine & _
                "Where A.������Ŀid =[1] And A.Ӧ�ó��� =[2] And A.�����ļ�id = B.ID And B.���� = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ݱ��", lngClinicId, CLng(Decode(intSourceType, 1, 1, 2, 2, 1)))
    
    If rsTemp.EOF Then Exit Sub
    strBillNo = "ZLCISBILL" & Format(rsTemp!���, "00000") & "-1"
    ReportOpen gcnOracle, glngSys, strBillNo, Me, "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & lngAdviceId, 2
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub Menu_Manage_�Ǽ�()
On Error GoTo ErrHandle
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set mfrmRISRequest = New frmRISRequest
        With mfrmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mstrPatientName = ""
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mstrCur���� = NeedName(mstrCur����)
'            .mlngQueueWay = mSysPar.lngQueueWay
            .mlngResultState = 0
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1), mblnAllDepts
            
            If .mlngResultState <> 0 Then '�ɹ�����
                Call StateCheck(2, .mlngAdviceId)
                
                If ufgStudyList.DataGrid.Rows = 2 Then
                    Call ufgStudyList.LocateRow(1)
                End If
                
                If .mlngResultState = 2 Then
                
                    '��������Ŷӽкţ��򱨵�����Ҫ�����ŶӽкŶ���......
                    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                        '������Ҫ����Ķ�������
                        If .mstrTechnicRoom = "" Then
                            '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                            Call mobjQueue.zlGetInQueueInf(.mlngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '�����Ϊ�գ���д���Ӧ��ִ�м�����
                            strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                
                    '��ӡ���뵥
                    Call AutoPrintApplication(.mlngAdviceId, .mlngSendNo, .mlngClinicID, .mintSourceType)
                End If
                
                '������������Ϣ
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mintImgCount = 0
            .mblnOK = False
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";�걾����ģ��;") > 0, True, False)
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            
            If .mblnOK Then '�ɹ�����
    
                Call StateCheck(2, .mlngAdviceId)
    
                
                If ufgStudyList.DataGrid.Rows = 2 Then
                    Call ufgStudyList.LocateRow(1)
                End If
                
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_Edit_Modify)
                
                '������������Ϣ
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ȡ���Ǽ�()
On Error GoTo ErrHandle
    Dim strSql As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "�����Ǽ�")
    
    Call RefreshList
    
    '����ҽ��������Ϣ
    Call mobjMsgCenter.Send_Msg_CancelAdvice(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_�ٻ�ȡ��()
'���ܣ��ٻر�ȡ���ĵǼ�
On Error GoTo ErrHandle
    Dim strSql As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷʵҪ�ٻر�ȡ���Ǽǵ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Call RefreshList
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_����()
On Error GoTo ErrHandle
    Dim blnFocusFind As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mcurAdviceInf.lngPatientFrom = 4 Then 'ֻ����첡�˲�ִ�����¹���
        Call zlDatabase.ExecuteProcedure("zl_PeisLockAdviceState(" & mListAdviceInf.lngAdviceId & ")", Me.Caption)
    End If
    
    If Me.ActiveControl Is Nothing Then
        blnFocusFind = False
    Else
        blnFocusFind = (Me.ActiveControl.Name = "txtFilter")
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mListAdviceInf.lngSendNo
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mstrPatientName = mListAdviceInf.strPatientName
            .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mstrCur���� = NeedName(mstrCur����)
'            .mlngQueueWay = mSysPar.lngQueueWay
            .mlngResultState = 0
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then  '�ɹ�����
                Call StateCheck(2)
                
                If .mblnIsRelationImage = True Then
                    '�������ǰ����ͼ��������Զ��������������ｫ��Ӱ��ͼ��ģ�����ˢ��
                    If Not mfrmWork_PacsImg Is Nothing Then
                        Call mfrmWork_PacsImg.zlUpdateAdviceInf(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, mListAdviceInf.intMoved = 1)
                        Call mfrmWork_PacsImg.zlRefreshFace(True)
                    End If
                End If
                
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
                
                If .mlngResultState = 2 Then
                    '��������Ŷӽкţ����ұ������Զ��Ŷӣ��򱨵�����Ҫ�����ŶӽкŶ���......
                    If mSysPar.blnUseQueue And mSysPar.blnAutoInQueue And Not mobjQueue Is Nothing Then
                        '������Ҫ����Ķ�������
                        If .mstrTechnicRoom = "" Then
                            '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                            Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '�����Ϊ�գ���д���Ӧ��ִ�м�����
                            strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                    
                                    '��ӡ���뵥
                    Call AutoPrintApplication(.mlngAdviceId, .mlngSendNo, .mlngClinicID, .mintSourceType)
                End If
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
                If mcurAdviceInf.lngPatientFrom <> 3 Then
                    Call mobjMsgCenter.Send_Msg_Arrange(.mlngAdviceId)
                End If
            End If

        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mListAdviceInf.lngSendNo
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = mlngCur����ID
            .mintImgCount = mintImgCount
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";�걾����ģ��;") > 0, True, False)
            .InitMvar
            If .RefreshPatiInfor(True) = True Then  'ˢ�²���
                .mblnOK = False
                .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            End If
            If .mblnOK Then  '�ɹ�����
                Call StateCheck(2)
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
            End If
            
        End With
    End If
    
    If blnFocusFind Then PatiIdentify.SetFocus '�Զ���λ����λ��
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'�Ŷӽк����
Private Sub zlInPacsQueue()
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mobjQueue Is Nothing Then Exit Sub
    
    '������Ҫ����Ķ�������
    If Trim(ufgStudyList.CurText("ִ�м�")) = "" Then
        '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
        Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, mlngCur����ID, strQueueName, strCodeNo)
    Else
        '�����Ϊ�գ���д���Ӧ��ִ�м�����
        strQueueName = NeedName(mstrCur����) & "-" & ufgStudyList.CurText("ִ�м�")
        strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(ufgStudyList.CurText("ִ�м�"), mlngCur����ID)
    End If
    
    Call mobjQueue.zlInQueue(mListAdviceInf.lngAdviceId, ufgStudyList.CurText("����"), mlngCur����ID, strQueueName, ufgStudyList.CurText("ִ�м�"), strCodeNo)
End Sub




Private Sub Menu_Manage_ȡ������()
On Error GoTo ErrHandle
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim lngResult As Long

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
  
    If mListAdviceInf.intStep <= 1 Then Call Menu_Manage_ȡ���Ǽ�: Exit Sub  '����������
    '------------------------------------��ǩ������Ҫ�Ȼ���ǩ�����ٳ���
    strSql = "Select Distinct B.���ʱ�� From ����ҽ������ A, ���Ӳ�����¼ B Where A.����ID=B.Id And A.ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ƿ�ǩ��", mListAdviceInf.lngAdviceId)
    
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!���ʱ��, "") <> "" Then 'ǩ������
            MsgBoxD Me, "��ǰ���˵ļ�鱨���Ѿ�ǩ��,����ȡ�����,���Ȼ���ǩ��!", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    '��������ȡ�Ļ�����Ƭ�����ܽ���ȡ��
    strSql = "select count(1) as ���� from ��������Ϣ a, ����ȡ����Ϣ b where a.����ҽ��ID=b.����ҽ��ID and a.ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, mListAdviceInf.lngAdviceId)
    If rsTemp.RecordCount > 0 Then
        If Val(Nvl(rsTemp!����)) > 0 Then
            Call MsgBoxD(Me, "�ü����ִ��ȡ�Ĳ��������ܽ���ȡ����", vbInformation, GetWindowCaption)
            Exit Sub
        End If
    End If

    If MsgBoxD(Me, "ȡ�����μ�齫ɾ����Ӧ�ļ��ͼ��ͼ�鱨�棬�Ƿ������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    If mListAdviceInf.strStudyUID <> "" And InStr(mstrPrivs, "���ͼ��") <= 0 Then
        MsgBoxD Me, "��û��������ͼ��Ȩ��,�������ͼ��,���Բ���ȡ��������!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ȡ���Ŷ���Ϣ
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCancelPacsQueue(mListAdviceInf.lngAdviceId)
    End If
    
    '�����RIS����վ������ͼ��������PACS�У�����Ҫ��ȡ��������Ȼ���ٵ���ZL_Ӱ����_CANCEL����ȡ������
    If mlngModule = G_LNG_PACSSTATION_MODULE And mListAdviceInf.intImageLocation = 1 Then
        'ȡ��ͼ�����
        Call XWUnmatchImage(mListAdviceInf.lngAdviceId, 0)
    End If
    
    'ȡ�����棬�޸����ݿ�״̬��ɾ����Ӱ�����¼��
    strSql = "ZL_Ӱ����_CANCEL(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",0," & mlngCur����ID & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSql = "ZL_������_����(" & mListAdviceInf.lngAdviceId & ")"
        zlDatabase.ExecuteProcedure strSql, GetWindowCaption
    End If
    
    '���ͼ��������PACS����ɾ��Ӱ���ļ���Ŀ¼
    If mListAdviceInf.intImageLocation = 0 Then
        RemoveCheckImages mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo
    End If
    
    Call StateCheck(1)
    
    '����״̬������Ϣ
    Call mobjMsgCenter.Send_Msg_StateCancel(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_����Ӱ��()
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    lngResult = -1
    '�����ģ���ΪRIS����վ����������������ݿ��ѯδƥ���ͼ���¼
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        lngResult = XWShowUnMatched(Me, mListAdviceInf.lngAdviceId, mListAdviceInf.strImgType)
        
        If lngResult = 0 Then
            'ͼ������ɹ���,ʹ��ֵΪ1
            mListAdviceInf.intImageLocation = 1
            
            ufgStudyList.CurText("ͼ��λ��") = "1"
            
            Call ufgStudyList.UpdateSourceData(mListAdviceInf.lngAdviceId, "ͼ��λ��", 1)
        End If
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mListAdviceInf.lngAdviceId, mstrPrivs, mListAdviceInf.intMoved = 1, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur����ID, 2, mListAdviceInf.strImgType
        
        If Not frmSelectMuli.mblnOK Then Exit Sub
        lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call AfterReleationImage(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, 2, True)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer
    Dim objDepartmentMenu As CommandBarControl
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    
    If Not mblnInitOk Then Exit Sub
    
    If mlngCur����ID <> control.DescriptionText Or (control.DescriptionText <> 0 And mblnAllDepts = True) Then
        mstrRPTExecutor = UserInfo.����
        
        If Not mobjWork_Report Is Nothing And mrtReportType = �����ĵ��༭�� Then
            Call mobjWork_Report.SetDocCreator(mstrRPTExecutor)
        End If
        
        stbThis.Panels(4).Text = "����ҽ����" & mstrRPTExecutor & "   ���ҽ����" & Split(stbThis.Panels(4).Text, "���ҽ����")(1)
        
        mcurAdviceInf = GetNullAdviceInf
        mListAdviceInf = mcurAdviceInf
        
        '�����л�������û�����´����˵��͹���ģ�飬Ҳû�е���cbrMain.RecalcLayout�������Ҫʹ�øö������ÿ����л���Ŀ�����Ϣ
        Set objDepartmentMenu = cbrMain.FindControl(, conMenu_View_Filter * 10#)
        
        If control.DescriptionText = 0 Then
            'ѡ�����п���
            mblnAllDepts = True
        
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "��ǰ����:ȫ������"
            
            If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                Set objControl = cbrdock.FindControl(, ID_Ӱ��ִ�м�)
                For i = 1 To objControl.CommandBar.Controls.Count
                    objControl.CommandBar.Controls(1).Delete
                Next
                
                Call InitExamineRoom(objControl, cbrPopControl, 0)
            End If
        Else
            'ѡ�񵥸�����
            mblnAllDepts = False
            
            mlngCur����ID = control.DescriptionText
            mstrCur���� = Split(control.Caption, "(")(0)
            mrtReportType = GetDeptPara(mlngCur����ID, "����༭��", 0)                 '����༭��
            
            Call InitStudyList
            
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "��ǰ����:" & mstrCur����

            If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                Set objControl = cbrdock.FindControl(, ID_Ӱ��ִ�м�)
                For i = 1 To objControl.CommandBar.Controls.Count
                    objControl.CommandBar.Controls(1).Delete
                Next
                
                Call InitExamineRoom(objControl, cbrPopControl, mlngCur����ID)
            End If
            
            Call InitModuleParameter(False)
            
            Call ReadStudyListColor(mlngCur����ID)
            
            Call RefreshCustomQueryMenu(cbrMain.FindControl(, conMenu_Manage_Query), mlngCur����ID)

        
            If mblnUseActivexCapture Then
                'ʹ��ActivexExe��ʽ����Ƶ�ɼ���ʽ����
                If Not mobjWork_ActiveVideo Is Nothing Then
                    Call mobjWork_ActiveVideo.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngCur����ID, Me.hWnd, Me, True)
                End If
            End If
            
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
            If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
            If Not mobjWork_His Is Nothing Then
                If mblnAllDepts Then
                    Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, UserInfo.����ID, Me)
                Else
                    Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
                End If
            End If
            
            '�����л�������������Ŷӽкţ�������Ŷӽк�ҳ��
            If mSysPar.blnUseQueue = True Then
                If mobjQueue Is Nothing Then
                    mstrWorkModule = mstrWorkModule & ";�Ŷӽк�ģ��;"
                    
                    Set mobjQueue = New frmWork_Queue
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, NeedName(mstrCur����), mstrPrivs)
                    
                    TabWindow.InsertItem 13, "�Ŷӽк�", mobjQueue.hWnd, 10011
                    TabWindow.Item(TabWindow.ItemCount - 1).tag = "�Ŷӽк�"
                    
                    Call picWindow_Resize
                Else
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, NeedName(mstrCur����), mstrPrivs)
                End If
            Else
                If mSysPar.blnUseQueue = False And Not mobjQueue Is Nothing Then
                    mstrWorkModule = Replace(mstrWorkModule, ";�Ŷӽк�ģ��;", "")
                    
                    For i = 0 To TabWindow.ItemCount - 1
                        If TabWindow.Item(i).tag = "�Ŷӽк�" Then
                            Call TabWindow.RemoveItem(i)
                            Exit For
                        End If
                    Next i
                    
                    Set mobjQueue = Nothing
                    
                    Call picWindow_Resize
                End If
            End If
            
            If mlngModule = G_LNG_PACSSTATION_MODULE Then
                If Not mfrmWork_PacsImg Is Nothing And InStr(mstrWorkModule, ";Ӱ��ͼ��ģ��;") > 0 Then
                    '����Ӱ���������Ӳ˵��͹�����
                    Call mfrmWork_PacsImg.zlMenu.zlCreateMenu(Me.cbrMain)
                    Call mfrmWork_PacsImg.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
                End If
            End If
            
            'Ϊ���ֱ���˵��ܹ�һֱ��ʾ��������Ҫ�Ա���˵����д���
            If Not mobjWork_Report Is Nothing And (InStr(mstrWorkModule, ";Ӱ�񱨸�ģ��;") > 0 Or InStr(mstrWorkModule, ";�������ģ��;") > 0) Then
                Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
                
                '���������Ӧ�˵��͹�����������༭��ʹ�ò�ͬ��ʽ��ʱ�򣬴����Ĳ˵���ͬ��
                Call mobjWork_Report.zlMenu.zlCreateMenu(Me.cbrMain)
                Call mobjWork_Report.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
            End If
            
            '�л���Ϣ�Ľ��տ���
            Call mobjMsgCenter.ChangeMsgReceiveDept(mlngCur����ID)
        End If
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        Call cbrMain.RecalcLayout
        
        '�����л�������ˢ�¿��Ҷ�Ӧ�ļ������
        Call RefreshList
        
        'ˢ���Ŷӽк�ģ�����ݣ�����Ѿ�����
        Call RefreshPacsQueueData
    
        Call FillCurAdviceTxtInfor
        Call FillCurAdviceAppend
        
        '�����л��󣬻ָ��������ѵĶ�ʱ��
        timerOperHint.Enabled = True
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur����ID
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub RefreshCustomQueryMenu(objQueryMenu As Object, ByVal lngDeptID As Long)
'���ݿ���Id,ˢ���Զ����ѯ�˵�
    Dim objCurQueryMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Dim i As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If objQueryMenu Is Nothing Then Exit Sub
    
    Set objCurQueryMenu = objQueryMenu
    
    For i = 1 To objCurQueryMenu.CommandBar.Controls.Count
        objCurQueryMenu.CommandBar.Controls(1).Delete
    Next
    
    
    Set rsTemp = zlDatabase.OpenSQLRecord("select Id, ��������, �Ƿ�Ĭ�� from Ӱ���ѯ���� where ʹ��״̬=1 and (��������=0 or �������� is null or ��������=[1]) order by �������� desc, �������", "������ѯ�˵�", lngDeptID)
    
    With objCurQueryMenu.CommandBar
        If rsTemp.RecordCount > 0 Then
            '�����Զ���Ĳ�ѯ����
            i = 65
            While Not rsTemp.EOF
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CustomQuery * 1000# + Val(Nvl(rsTemp!ID)), Nvl(rsTemp!��������) & "(&" & Chr(i) & ")", "", 0, False)
                
                i = i + 1
                If Chr(i) = "F" Or Chr(i) = "C" Then i = i + 1
                
                If Val(Nvl(rsTemp!�Ƿ�Ĭ��)) = 1 Then
                    cbrControl.IconId = 3558
                End If
                
                Call rsTemp.MoveNext
            Wend
        End If
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CustomQuery, "�ۺϲ�ѯ", "", 721, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ConfigQuery, "��ѯ����", "", 3965, False)
    End With
    
End Sub

Private Sub Menu_View_Refresh_click()
On Error GoTo ErrHandle
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo ErrHandle
    zlHomePage hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cboTimes_Click()
On Error GoTo ErrHandle
    Dim lngAdviceId As Long
    
    If cboTimes.ListCount <= 1 Then Exit Sub
    If cboTimes.tag = "" Then Exit Sub '��ʱcbotime��Ŀδ������ɣ���listindex��ֵ����
    
    lngAdviceId = cboTimes.ItemData(cboTimes.ListIndex)
    
    If lngAdviceId = mListAdviceInf.lngAdviceId Then
        Call ufgStudyList_OnSelChange
        Exit Sub  '�����뵱ǰѡ��ҽ��ID��ͬʱ���ɱ���������
    End If

    mblnIsHistory = True
    
    '�����������̵������Ⱥ�˳�������
    mcurAdviceInf = GetAdviceDetailInf(lngAdviceId)
    
    Call FillCurAdviceTxtInfor    '������Ϸ����˻�����Ϣ
    Call FillCurAdviceAppend   '������½�ҽ������
    
    'ѡ����ȫ�����Һ������л��˿���
    If mlngCur����ID <> mcurAdviceInf.lngExeDepartmentId And mblnAllDepts = True Then
        mlngCur����ID = mcurAdviceInf.lngExeDepartmentId
        mstrCur���� = GetDeptName(mlngCur����ID, mstrCanUse����)
    End If
    
    Call ShowTab    '���ݲ����ṩ��ͬѡ�
    
    Call RefreshModuleAdviceInf
    Call RefreshTabWindow   'ˢ���Ӵ���

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function GetDeptName(lngDeptID As Long, strDeptStrings As String) As String
'ͨ�����õĿ��Ҵ�����ȡָ������ID�Ŀ�������
On Error GoTo ErrHandle
    Dim strDepts() As String
    Dim i As Integer
    
    strDepts = Split(strDeptStrings, "|")
    For i = 0 To UBound(strDepts)
        If Split(strDepts(i), "_")(0) = lngDeptID Then
            GetDeptName = Split(strDepts(i), "_")(1)
            Exit For
        End If
    Next i
Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
End Function


Private Sub cboTimes_DropDown()
On Error GoTo ErrHandle
    Call SendMessage(cboTimes.hWnd, &H160, 500, 0)
ErrHandle:
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim objTmpControl As CommandBarControl
    
    Dim i As Integer, j As Integer
    Dim strTemp As String
    Dim strCardName As String
    Dim strCardText As String
    Dim lngPatientID As Long
    Dim strAllRooms As String
    Dim strRoomName As String
    Dim strStudyTypes As String
    
    Select Case control.ID
        Case ID_���ҷ�ʽ
            If control.IconId = 3 Then
                control.IconId = 4
                
                mstrLocateWay = PatiIdentify.GetCurCard.����
                '�ڿ��ٹ����������λ�Ͳ���ʱ������ˢ���ؼ�IDKindStrʱ�������ItemClick�¼���
                '�����޷��ֱ��¼��λ�Ͳ����ֶ���Ϣ�������ô˱�����ǣ�Ϊtrueʱ������ItemClick�¼�
                mblnAssignment = True
                PatiIdentify.IDKindStr = InitCardType(Replace(CONST_STR_FIND_CARD_TYPE, "[------]", GetStudyNumberDisplayName))
                PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(mstrFindWay)
                mblnAssignment = False
                
                cbrdock.FindControl(, ID_��ʼ����).Caption = "��ʼ����"
                
                Call zlDatabase.SetPara("��λ���ҷ�ʽ", 1, glngSys, mlngModule)
            Else
                control.IconId = 3
                
                mstrFindWay = PatiIdentify.GetCurCard.����
                
                Call subRefreshFilterCondition("", "")
                Call RefreshList
                
                PatiIdentify.tag = ""
                mblnAssignment = True
                PatiIdentify.IDKindStr = InitCardType(Replace(CONST_STR_LOCAL_CARD_TYPE, "[------]", GetStudyNumberDisplayName))
                PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(mstrLocateWay)
                mblnAssignment = False
                
                cbrdock.FindControl(, ID_��ʼ����).Caption = "��ʼ��λ"
                
                Call zlDatabase.SetPara("��λ���ҷ�ʽ", 0, glngSys, mlngModule)
            End If
            
            Exit Sub
            
            
            
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
            
        Case ID_δ��
            mblncmdδ�� = Not control.Checked
            
        Case ID_����
            mblncmd���� = Not control.Checked
            
        Case ID_����
            mblncmd���� = Not control.Checked
            
        Case ID_�޷�
            mblncmd�޷� = Not control.Checked
            
        Case ID_Ӱ����� + 1 To ID_Ӱ����� + 40
            control.Checked = Not control.Checked
            mblncmdӰ�����(control.ID - ID_Ӱ����� - 1) = control.Checked
            
            If control.Checked = True Then
                mintcmdӰ����� = mintcmdӰ����� + 1
            Else
                mintcmdӰ����� = mintcmdӰ����� - 1
            End If
            
            Set objControl = cbrdock.FindControl(, ID_Ӱ�����)
            
            If mintcmdӰ����� = 0 Then
                strTemp = "���"
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Checked = True Then
                        strTemp = IIf(strTemp = "", Mid(objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, 1, InStr(objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, "(") - 1), strTemp & "," & Mid(objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, 1, InStr(objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, "(") - 1))
                        strStudyTypes = strStudyTypes & "," & Mid(objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, 1, InStr(objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, "(") - 1)
                    End If
                Next i
            End If
            
            If strStudyTypes <> "" Then strStudyTypes = Mid(strStudyTypes, 2)
            
            Call InitStudyPlace(cbrdock.FindControl(, ID_��鲿λ), cbrPopControl, strStudyTypes)
            
            If strTemp = "���" Or strTemp = "" Then
                objControl.ToolTipText = "����Ӱ�������й���"
            Else
                objControl.ToolTipText = "��ʾӰ�����Ϊ[" & strTemp & "]�ļ��"
            End If
            
            objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            
        Case ID_Ӱ��ִ�м� + 1 To ID_Ӱ��ִ�м� + 40
            control.Checked = Not control.Checked
            mblncmdӰ��ִ�м�(control.ID - ID_Ӱ��ִ�м� - 1) = control.Checked
            
            If control.Checked = True Then
                mintcmdӰ��ִ�м� = mintcmdӰ��ִ�м� + 1
            Else
                mintcmdӰ��ִ�м� = mintcmdӰ��ִ�м� - 1
            End If
            
                        
            Set objControl = cbrdock.FindControl(, ID_Ӱ��ִ�м�)
            
            mstrSelQueueRooms = ""
            
            If mintcmdӰ��ִ�м� <= 0 Then
                strTemp = "ִ�м�"
                mintcmdӰ��ִ�м� = 0
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                
                    Set cbrPopControl = objControl.CommandBar.FindControl(, ID_Ӱ��ִ�м� + i)
                    If Not cbrPopControl Is Nothing Then
                        strRoomName = Mid(cbrPopControl.Caption, 1, InStr(cbrPopControl.Caption, "(") - 1)
                        
                        If cbrPopControl.Checked = True Then
                            strTemp = IIf(strTemp = "", strRoomName, strTemp & "," & strRoomName)
                            
                            If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
                            mstrSelQueueRooms = mstrSelQueueRooms & cbrPopControl.ToolTipText
                        End If
                    End If
                    
                Next i
            End If
            
            '���û��ѡ���κ�ִ�м䣬��Ĭ����ʾ����ִ�м������
            If Trim(mstrSelQueueRooms) = "" Then mstrSelQueueRooms = mstrAllQueueRooms
            
            If strTemp = "ִ�м�" Or strTemp = "" Then
                objControl.ToolTipText = "����Ӱ��ִ�м���й���"
            Else
                objControl.ToolTipText = "��ʾӰ��ִ�м�Ϊ[" & strTemp & "]�ļ��"
            End If
            
            '���˵���������6���ַ�ʱ��������ַ�ʹ��ʡ�Ժ���ʾ
            objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
 
            '����������Ŷӽкţ���ˢ���Ŷӽк�����
            Call RefreshPacsQueueData
        
        Case ID_��鲿λ To ID_��鲿λ + 4500
            control.Checked = Not control.Checked

            Set objControl = cbrdock.FindControl(, ID_��鲿λ)

            strTemp = ""
            mstrcmd��λ = ""
            
            For i = 1 To objControl.CommandBar.Controls.Count
                Set objTmpControl = objControl.CommandBar.Controls(i)

                If Not objTmpControl Is Nothing Then
                    If objTmpControl.Checked = True Then
                        strTemp = IIf(strTemp = "", objTmpControl.Caption, strTemp & "," & objTmpControl.Caption)
                        mstrcmd��λ = mstrcmd��λ & "," & objTmpControl.Category
                    End If
                End If
            Next i
            
            If control.Checked Then
                mstrcmd��λ���� = mstrcmd��λ���� & "," & control.Caption
            Else
                mstrcmd��λ���� = Replace(mstrcmd��λ����, "," & control.Caption, "")
            End If

            If strTemp = "" Then
                strTemp = "��λ"
                objControl.ToolTipText = "���ݼ�鲿λ���й���"
            Else
                objControl.ToolTipText = "��ʾ��鲿λΪ[" & strTemp & "]�ļ��"
            End If

            objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            
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
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
            
            
            
        Case ID_�������_����
            mblncmd���� = Not control.Checked
        Case ID_�������_����ʯ��
            mblncmd����ʯ�� = Not control.Checked
        Case ID_�������_����
            mblncmd���� = Not control.Checked
        Case ID_�������_ϸ��
            mblncmdϸ�� = Not control.Checked
        Case ID_�������_ʬ��
            mblncmdʬ�� = Not control.Checked
        Case ID_�������_����
            mblncmd���� = Not control.Checked
            
            
            
        Case ID_����סԺ
            control.Checked = Not control.Checked
            mblncmd���� = Not mblncmd����
        Case ID_��ʼ����
            Call StartReadCard
            Call SaveFilterCmd
            
            Exit Sub
    End Select
    
    '������ٹ�������������
    Call SaveFilterCmd
    
    cbrdock.RecalcLayout
    
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subRefreshFilterCondition(ByVal strCardName As String, ByVal strFilter As String)
'------------------------------------------------
'���ܣ���txtFilter�ؼ������ݸ��¹�������
'������ strFilter --- ��������
'���أ���
'------------------------------------------------

On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strTemp As String
    
    With SQLCondition
        .���� = ""
        .���￨ = ""
        .����� = 0
        .סԺ�� = 0
        .������ = ""
        .���ݺ� = ""
        .���� = 0
        .���֤ = ""
        .IC�� = ""
        .������� = -1
        .����ID = 0
        
        Select Case strCardName
            Case "����", "��  ��", "��   ��"  '��������ǰ�ķ�ʽ����
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
                
            Case "������"
                .������ = Trim(strFilter)
                
            Case "���ݺ�"
                If Len(Trim(strFilter)) = 0 Then
                     .���ݺ� = ""
                Else
                    If Len(Trim(strFilter)) < 8 And Not IsNumeric(Trim(strFilter)) Then
                        strTemp = GetFullNO(0, 0)
                        strTemp = Mid(strTemp, 1, Len(strTemp) - Len(strFilter)) & strFilter
                    Else
                        strTemp = GetFullNO(Nvl(strFilter, 0), 0)
                    End If
                    
                    PatiIdentify.Text = strTemp
                    .���ݺ� = strTemp
                End If
                
            Case GetStudyNumberDisplayName
                If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                    .���� = Val(strFilter)
                Else
                    If Trim(strFilter) = "" Then
                        Exit Sub
                    End If
                    
                    If UCase(Mid(strFilter, Len(strFilter), 1)) = UCase("Z") Then       '���ͨ��ɨ��ǹ��ɨ�����Z����ͷ�ĺ��룬˵������Ƭ��
                        strSql = "select ����� from ��������Ϣ a, ������Ƭ��Ϣ b where a.����ҽ��ID=b.����ҽ��Id and b.ID=[1]"
                        Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, Mid(strFilter, 1, Len(strFilter) - 1))
                        
                        If rsData.RecordCount > 0 Then
                            .���� = Nvl(rsData!�����)
                            
                            PatiIdentify.Text = .����
                        End If
                    ElseIf UCase(Mid(strFilter, Len(strFilter), 1)) = UCase("T") Then   '���ͨ��ɨ��ǹ��ɨ�����T����ͷ�ĺ��룬˵�����ؼ���Ƭ��
                        strSql = "select ����� from ��������Ϣ a, �����ؼ���Ϣ b where a.����ҽ��ID=b.����ҽ��Id and b.ID=[1]"
                        Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, Mid(strFilter, 1, Len(strFilter) - 1))
                        
                        If rsData.RecordCount > 0 Then
                            .���� = Nvl(rsData!�����)
                            
                            PatiIdentify.Text = .����
                        End If
                    Else
                        .���� = GetPatholNum(Trim(strFilter))
                    End If
                End If
                
            Case "���֤��", "���֤"
                .���֤ = Trim(strFilter)
                
            Case Else
                .����ID = Val(strFilter)
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetPatholNum(ByVal strSureNum As String) As String
'�ֽ�ȷ�Ϻ���
    Dim lngFindSplitChar As Long
    
    lngFindSplitChar = InStr(1, strSureNum, "-")
    
    If lngFindSplitChar > 0 Then
        GetPatholNum = UCase(Mid(strSureNum, 1, lngFindSplitChar - 1))
    Else
        GetPatholNum = UCase(strSureNum)
    End If
    
End Function

Private Sub cbrdock_Resize()
On Error GoTo ErrHandle
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    If tabFilter.Visible Then
        'ֻ�в�����վ����ʾtab����ҳ��
        tabFilter.Top = lngTop
        tabFilter.Left = lngLeft
        tabFilter.Width = PicList.Width
        
        picExeState.Left = lngLeft
        picExeState.Top = lngTop + IIf(tabFilter.Visible, tabFilter.Height, 0)
        picExeState.Width = PicList.Width
    End If
    
    ufgStudyList.Top = IIf(tabFilter.Visible, picExeState.Top + picExeState.Height, lngTop)
    ufgStudyList.Left = lngLeft
    ufgStudyList.Width = PicList.Width
    ufgStudyList.Height = Abs(PicList.Height - lngTop - picAppend.Height - IIf(tabFilter.Visible, tabFilter.Height + picExeState.Height, 0))

    PicLine.Top = lngTop + ufgStudyList.Height + IIf(tabFilter.Visible, tabFilter.Height + picExeState.Height, 0)
    PicLine.Left = PicList
    PicLine.Width = PicList.Width
    PicLine.Height = 90

    picAppend.Top = PicLine.Top + PicLine.Height
    picAppend.Left = lngLeft
    picAppend.Width = PicList.Width
    picAppend.Height = PicList.Height - lngTop - ufgStudyList.Height - IIf(tabFilter.Visible, tabFilter.Height + picExeState.Height, 0)

ErrHandle:
End Sub


Private Sub Form_Activate()
On Error GoTo ErrHandle
    '�жϵ�ǰ����ģ���Ƿ�Ӱ��ɼ�ģ�飬����ǣ����жϲɼ�ģ���Ƿ��ʼ��������Ѿ���ʼ�������˳��ù��̣�����Ͷ�����г�ʼ��������ʾ
    '��Ϊ��ͬһ����̨�У����ͬʱ�򿪲�����Ƶ�ɼ�ģ�齫���л�������һϵͳ�˳�ʱ���ɼ�ģ��Ҳ�����ͷţ�����л��ص�ǰϵͳ����Ҫ�ж��Ƿ���³�ʼ���ɼ�ģ��
    If Not mblnInitOk Then Exit Sub
    If TabWindow.Selected Is Nothing Then Exit Sub
    If TabWindow.Selected.tag <> "Ӱ��ɼ�" Then Exit Sub
    
    If mblnUseActivexCapture Then
        'ʹ��ActivexExe��ʽ����Ƶ�ɼ�����
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved)
            Call mobjWork_ActiveVideo.zlRefreshVideoWindow
            Call mobjWork_ActiveVideo.zlRefreshData(True)
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '���ع���ģ��ʱ���������˳�����
    If Not mblnInitOk Then Cancel = True
    
    If mblnUseActivexCapture Then
        'TODO:�ȴ�������ɺ󣬲������˳���Ƶ
    End If
End Sub


Private Sub labStudyNum_Change()
On Error GoTo ErrHandle
    Call picAppend_Resize
ErrHandle:
End Sub

Private Sub lbl������Ϣ_Change()
On Error GoTo ErrHandle
    Call picAppend_Resize
ErrHandle:
End Sub

Private Sub mfrmRISRequest_HaveRegist()
    Dim strQueueName As String
    Dim strCodeNo As String
    With mfrmRISRequest
        If .mlngResultState <> 0 Then '�ɹ�����
            '��������Ŷӽкţ��򱨵�����Ҫ�����ŶӽкŶ���......
            If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And .mlngResultState = 2 Then
                '������Ҫ����Ķ�������
                If .mstrTechnicRoom = "" Then
                    '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                    Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                Else
                    '�����Ϊ�գ���д���Ӧ��ִ�м�����
                    strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                    strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                End If
                
                Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
            End If
            
            '������������Ϣ
            Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
        End If
    End With
End Sub

Private Sub mobjCaptureHot_OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo ErrHandle
    Dim lngWindowPID As Long
    Dim lngVideoPID As Long
    Dim lngCurrentPID As Long

    If lngMsg <> WM_KEYDOWN Then Exit Sub
    If Trim(mstrCaptureHot) = "" And Trim(mstrCaptureAfterHot) = "" And Trim(mstrCaptureAfterTagHot) = "" Then Exit Sub
    
    mCaptureMsg.lngMsg = lngMsg
    mCaptureMsg.lngVirtualKey = lngVkCode
    mCaptureMsg.lngScanKey = lngScanCode
    mCaptureMsg.lngFlags = lngFlags
    
    '����ֱ����Hook�ص�������ʹ��ActiveExe�������ط�������������δ֪�������
    timerCapture.Enabled = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjEvent_OnWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceId As Long, ByVal other As Variant)
'��Ӧ����ģ��ִ�в����󴥷����¼�
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strRoom As String
    Dim i As Integer
    Dim j As Integer
    Dim strStudyUID As String
    
    Dim lngcurRow As Long
    Dim lngColIndex As Long
    
    Select Case lngWorkType
        Case TWorkEventType.wetGetImg           '��ȡͼ��QR��***************************************
            Call RefreshList
            
        Case TWorkEventType.wetTechDo           '��ʦִ��***************************************
            If mListAdviceInf.lngAdviceId = lngAdviceId Then
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "��鼼ʦ", UserInfo.����)
                
                If ufgStudyList.CurText("�Ƿ�ʦȷ��") = "1" Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, ufgStudyList.DataGrid.RowSel, ufgStudyList.GetColIndex("��鼼ʦ")) = imgList.ListImages("��鼼ʦ").Picture
                    ufgStudyList.CurText("��鼼ʦ") = UserInfo.����
                Else
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, ufgStudyList.DataGrid.RowSel, ufgStudyList.GetColIndex("��鼼ʦ")) = Nothing
                    ufgStudyList.CurText("��鼼ʦ") = IIf(ufgStudyList.CurText("��鼼ʦ") = UserInfo.����, "", ufgStudyList.CurText("��鼼ʦ"))
                End If
            End If
            
        Case TWorkEventType.wetChangeImgType    '�ı�Ӱ������***************************************
            Call RefreshList(lngAdviceId)
        
        Case TWorkEventType.wetLockStudy, TWorkEventType.wetUnLockStudy        '�������,�������
            '�޸ı�ǩҳ����ʾ��ʽ�ͱ���
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*Ӱ��ɼ�*" Then
                    If lngWorkType = wetLockStudy Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "��" & other & "�� Ӱ��ɼ�"
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "Ӱ��ɼ�"
                    End If
                    Exit For
                End If
            Next i
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngWorkType, lngAdviceId, other)
            
        Case TWorkEventType.wetCaptureFirstImg, TWorkEventType.wetDelAllImg, TWorkEventType.wetUpdateImg  '�ɼ���һ��ͼ��***************************************
            '���¼���б�
            
            strStudyUID = other
            
            If lngWorkType = wetCaptureFirstImg Then
                '��д���ִ�м�
                If mSysPar.lngQueueWay = 1 And mSysPar.blnUseQueue Then
                    strRoom = mSysPar.strLocalRoom
                    If InStr(strRoom, "-") > 0 Then strRoom = Mid(mSysPar.strLocalRoom, 1, InStr(mSysPar.strLocalRoom, "-") - 1)
        
                    strSql = "zl_Ӱ����_����ִ�м�(" & lngAdviceId & ",'" & strRoom & "','" & NeedName(mSysPar.strLocalRoom) & "')"
                    Call zlDatabase.ExecuteProcedure(strSql, GetWindowCaption)
                End If
                
                '���¼���б�
                Call UpdateStudyListState(lngAdviceId, strStudyUID, True, True)
                
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(lngAdviceId)
            ElseIf lngWorkType = wetDelAllImg Then
                '���¼���б�
                Call UpdateStudyListState(lngAdviceId, strStudyUID, False, True)
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateCancel(lngAdviceId)
            End If


            If Val(ufgStudyList.CurKeyValue) <> CStr(lngAdviceId) Then Exit Sub
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngWorkType, lngAdviceId, other)
            
            'ˢ��Ƕ���ؼ챨��������½�����ͼͼ��
            If lngWorkType = wetUpdateImg Then If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
        Case wetChangeUser
            '�����û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
        
            Call ChangeUser
            
            '�����û�����Ҫˢ�±���༭������Ϊ�û�������ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case wetPatholRequest       '��������
            Call RefreshList(lngAdviceId)
            
        Case wetPatholQuality       '��������
            lngcurRow = ufgStudyList.FindRowIndex(CStr(lngAdviceId), "ҽ��ID", True)
            
            If lngcurRow > 0 Then
                ufgStudyList.Text(lngcurRow, "����") = other
                
                lngColIndex = ufgStudyList.GetColIndex("����")
                
                If other = "����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, lngColIndex) = vbGreen
                If other = "��������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, lngColIndex) = vbYellow
                If other = "������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, lngColIndex) = vbRed
                
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "�ۺ�����", other)
            End If
        
        Case wetPatholBatSlices     '��Ƭ��������
            Call RefreshList(lngAdviceId)
            
        Case wetPatholBatSpeExm     '�ؼ���������
            Call RefreshList(lngAdviceId)
            
        Case wetSpecimenAccept      '�걾����
            Call RefreshPatholExecuteState(lngAdviceId)
            
            With ufgStudyList
                lngcurRow = .DataGrid.FindRow(CStr(lngAdviceId), , .GetColIndex("ҽ��ID"))
                
                If lngcurRow > 0 Then
                    If Trim(.Text(lngcurRow, "�����")) = "" Then
                        .Text(lngcurRow, "�����") = other
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "�����", other)
                        
                        .Text(lngcurRow, "���״̬") = 2
                        
                        .Text(lngcurRow, "������") = "�ѱ���"
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "������", 2)
                        
                        .Text(lngcurRow, "����ʱ��") = zlDatabase.Currentdate
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "����ʱ��", zlDatabase.Currentdate)
                        
                        .Text(lngcurRow, "������") = UserInfo.����
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "������", UserInfo.����)
                        
                        .Text(lngcurRow, "�������") = "�Ѻ���"
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "�������", "�Ѻ���")
                        
                        labStudyNum.Caption = "[�����:" & IIf(other <> "", other, "---") & "]"
                        
                        'ˢ����������ģ������
                        If Not mobjWork_Pathol Is Nothing Then
                            Call mobjWork_Pathol.zlUpdateAdviceInf(lngAdviceId, 0, 2, False)
                            Call mobjWork_Pathol.NotificationRefresh(mtAll)
                        End If
                    End If
                End If
            End With
        
        Case wetSpecimenReject      '�걾����
        
        Case wetSpecimenSave        '�걾����
            '�걾�����ˢ��ȡ��ģ������
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtMaterial)
            
        Case wetMaterialSure        'ȡ��ȷ��
            Call RefreshPatholExecuteState(lngAdviceId)
            
            'ˢ����Ƭģ������
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtSlices)
            
        Case wetMaterialSave        '�Ŀ鱣��
            'ˢ����Ƭģ������
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtSlices)
            
        Case wetSlicesSure          '��Ƭȷ��
            Call RefreshPatholExecuteState(lngAdviceId)
    
        Case wetSpeExamSure         '�ؼ�ȷ��
            Call RefreshPatholExecuteState(lngAdviceId)
            
        Case wetViewEprReport       'Ԥ�����Ӳ�������
            Dim strRepInf() As String
            
            strRepInf = Split(other & ",,", ",")
            
            If Val(strRepInf(0)) <= 0 Then Exit Sub
            
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.ViewEPRReport(Val(strRepInf(0)), IIf(Val(strRepInf(1)) = 1, True, False))
        
        Case wetViewPacsImage       'Ԥ��Pacsͼ��
            '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
            Call OpenViewer(2, mobjPacsCore, lngAdviceId, False, Me, , , mSysPar.blnLocalizerBackward)
            
        Case wetRejectReport        '���汻����
            lngcurRow = ufgStudyList.DataGrid.FindRow(CStr(lngAdviceId), , ufgStudyList.GetColIndex("ҽ��ID"))
            
            If lngcurRow <= 0 Then Exit Sub
                        
            ufgStudyList.Text(lngcurRow, "������") = "�Ѳ���"
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, 1, lngcurRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�Ѳ���
            
            Call ufgStudyList.UpdateSourceData(lngAdviceId, "������", -1)
            
            '����״̬ͬ����Ϣ
            Call mobjMsgCenter.Send_Msg_StateSync(lngAdviceId)
        Case wetPrintFilm
            '����Ƭ��ӡ��Ϣ
            lngcurRow = ufgStudyList.DataGrid.FindRow(CStr(lngAdviceId), , ufgStudyList.GetColIndex("ҽ��ID"))
            
            If lngcurRow <= 0 Then Exit Sub
            
            ufgStudyList.Text(lngcurRow, "��Ƭ��ӡ") = "1"
            Call ufgStudyList.UpdateSourceData(lngAdviceId, "��Ƭ��ӡ", "1")
                        
        End Select
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub RefreshPatholExecuteState(ByVal lngAdviceId As Long)
'���²���ִ��״̬
    Dim lngcurRow As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select �������,ȡ�Ĺ���,��Ƭ����,���߹���,���ӹ���,��Ⱦ���� from ��������Ϣ where ҽ��Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngcurRow = ufgStudyList.DataGrid.FindRow(CStr(lngAdviceId), , ufgStudyList.GetColIndex("ҽ��ID"))
        
    If lngcurRow > 0 Then
        ufgStudyList.Text(lngcurRow, "����ִ��״̬") = GetPatholExecuteState(rsData)
        ufgStudyList.Text(lngcurRow, "������") = Decode(Nvl(rsData!�������), 1, "����", 2, "ϸ��", 3, "����", 4, "ʬ��", 5, "����ʯ��", "����")
        
    End If
End Sub

Private Function GetPatholExecuteState(rsData As ADODB.Recordset) As String
    Dim strState As String

    strState = ""
    
    If Nvl(rsData!ȡ�Ĺ���) = 1 Then strState = "��ȡ��"

    If Nvl(rsData!��Ƭ����) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "����Ƭ"
    End If
    
    If Nvl(rsData!���߹���) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "������"
    End If
    
    If Nvl(rsData!���ӹ���) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "�����"
    End If
    
    If Nvl(rsData!��Ⱦ����) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "����Ⱦ"
    End If
    
    
    If Nvl(rsData!��Ƭ����) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "��Ƭ����"
    End If
    
    If Nvl(rsData!���߹���) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "���߽���"
    End If
    
    If Nvl(rsData!���ӹ���) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "���ӽ���"
    End If
    
    If Nvl(rsData!��Ⱦ����) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "��Ⱦ����"
    End If
    
    If Trim(strState) = "" Then strState = ""
    
    GetPatholExecuteState = strState
End Function

Private Sub mobjMsgCenter_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, objMsgPro As clsMipModule, objXML As clsXML)
'��Ϣ���մ���
    Dim lngRowIndex As Long
    Dim lngAdviceId As Long
    Dim lngStudyState As Long
    Dim strHint As String
    Dim strSql As String
    Dim rsReport As ADODB.Recordset
    Dim rsDataMulite As ADODB.Recordset
    Dim rsDataMuliteClone As ADODB.Recordset
    Dim strCurNo As String
    Dim strNodeId As String
    Dim lngChargeState As Long
    
    
    lngAdviceId = 0
    
    '��ȡ��Ϣ�ж�Ӧ��ҽ��ID����
    If strMsgItemIdentity = G_STR_MSG_ZLHIS_PACS_003 Then
        rsData.Filter = "node_name='study_order_id'"
    Else
        rsData.Filter = "node_name='order_id'"
    End If
    
    If rsData.RecordCount > 0 Then
        lngAdviceId = Val(Nvl(rsData!node_value))
    End If
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_ZLHIS_CIS_017    '�������
            '������Ϣ��ʾ@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & Nvl(rsData!node_value) & " ��Ҫ���м�飬�뼰ʱ����"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
            
            '�����ݿ���ˢ������
            Call RefreshList(0, True)
            
        Case G_STR_MSG_ZLHIS_CIS_024    'ҽ������
            '����������ʾ@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & Nvl(rsData!node_value) & " �ļ��ҽ���ѱ������� "
        
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
        
            '�ж�ҽ��ID�Ƿ����б��д��ڣ����������ɾ����Ӧ����
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID")
            If lngRowIndex > 0 Then
                Call ufgStudyList.SyncText(lngRowIndex, "������", "�Ѿܾ�", False)
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "ִ��״̬", 2)
            End If
            
            'ִ��UpdateSourceData��AdoData.Filter���
            ufgStudyList.AdoData.Filter = ""
            
            '���ݵ�ǰѡ���ҽ���ж��Ƿ���Ҫˢ���б�
            Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
            
        Case G_STR_MSG_ZLHIS_CIS_025    'Σ��ֵ�Ķ�
            '����Ϣƽ̨���õ�����ʾ
            
        Case G_STR_MSG_ZLHIS_CHARGE_003 '���ﻼ�߻��۵���
            'ˢ���շ�״̬��ʾ
            '���ݵ��ݺŲ��Ҷ�Ӧ��ҽ��ID
            rsData.Filter = "node_name='bill_no'"
            rsData.Sort = "node_name"
            If rsData.RecordCount <= 0 Then
                Exit Sub
            End If
            
            '��ȡ���е��ݵ���Ϣ
            If objXML.GetMultiNodeRecord("charge_bill", rsDataMulite) = False Then Exit Sub
            
            Set rsDataMuliteClone = rsDataMulite.Clone
            
            rsDataMulite.Filter = "node_name='charge_bill'"
            
            If rsDataMulite.RecordCount <= 0 Then Exit Sub
            
            Do While Not rsDataMulite.EOF
                '��ȡ����charge_bill��ӦID���ڴ��ڶ�����ݵ�����£������charge_bill��D���Ҷ�Ӧ�ĵ�����Ϣ
                strNodeId = Val(Nvl(rsDataMulite!ID))
                
                '��ȡcharge_bill�ڵ��µĵ��ݺţ�����charge_bill��id����
                rsDataMuliteClone.Filter = "parent_id=" & strNodeId & " and node_name='bill_no'"
                If rsDataMuliteClone.RecordCount > 0 Then strCurNo = Nvl(rsDataMuliteClone!node_value)
                
                '��ȡcharge_bill�ڵ��µĵ��ݷ���״̬,����charge_bill��id����
                rsDataMuliteClone.Filter = "parent_id=" & strNodeId & " and node_name='charge_state'"
                If rsDataMuliteClone.RecordCount > 0 Then lngChargeState = Val(Nvl(rsDataMuliteClone!node_value))
                
                lngRowIndex = ufgStudyList.FindRowIndex(strCurNo, "���ݺ�")
                
                If lngChargeState = 2 Then  '=2��ʾ�Ѿ��շ�
                    ufgStudyList.Text(lngRowIndex, "�շ�") = " "
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRowIndex, ufgStudyList.GetColIndex("�շ�")) = imgList.ListImages("�շ�").Picture
                    
                    lngAdviceId = Val(ufgStudyList.KeyValue(lngRowIndex))
                    
                    If mListAdviceInf.lngAdviceId = lngAdviceId Then
                        'ˢ���б��·��ķ���״̬��ʾ
                        lblCash.Caption = "��"
                        lblCash.ForeColor = &H8000&
                    End If
                End If
                
                rsDataMulite.MoveNext
            Loop
        
        Case G_STR_MSG_ZLHIS_PACS_001   '��鱨����ɣ������ɲ����鱨���������
            '�����б��е���ʾ״̬
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID")
            If lngRowIndex > 0 Then
                Call ufgStudyList.SyncText(lngRowIndex, "������", "�����", False)
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "������", 6)
                
                'ִ��UpdateSourceData��AdoData.Filter���
                ufgStudyList.AdoData.Filter = ""
                
                '���ݸ��º�����ݣ�ˢ���б���ʾ
                Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
            End If
            
        Case G_STR_MSG_ZLHIS_PACS_002, G_STR_MSG_ZLHIS_PACS_003  '���״̬ͬ������״̬���˴���
            '������汻���أ���Ҫ��������@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='study_cur_state'"
            If Nvl(rsData!node_value) = -1 Then
                
                
                '��Ҫ�жϵ�ǰ�û��Ƿ�Ϊ������
                strSql = "select ������ from Ӱ�����¼ where ҽ��ID=[1]"
                Set rsReport = zlDatabase.OpenSQLRecord(strSql, "��ѯ������", lngAdviceId)
                If rsReport.RecordCount > 0 Then
                    If Nvl(rsReport!������) = UserInfo.���� Then
                        '������Ϣ
                        rsData.Filter = "node_name='patient_name'"
                        strHint = "����" & Nvl(rsData!node_value) & "�ı����ѱ����أ���ע�⴦��"
                        
                        Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
                    End If
                End If
            End If
            
            
        
            'ˢ���б��Ӧ��ʾ
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID")
            If lngRowIndex > 0 Then
            
                rsData.Filter = "node_name='study_cur_state'"
                If rsData.RecordCount > 0 Then
                    lngStudyState = Val(Nvl(rsData!node_value))
                    Call ufgStudyList.UpdateSourceData(lngAdviceId, "������", lngStudyState)
                    
                    '���ݸ��º�����ݣ�ˢ���б���ʾ
                    Call RefreshList '(IIf(lngAdviceID = mcurAdviceInf.lngAdviceID, lngAdviceID, 0), False)
                End If
            End If
        
        Case G_STR_MSG_ZLHIS_PACS_004   '��鱨�泷��
            '�����б��е���ʾ״̬
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID")
            If lngRowIndex > 0 Then
                rsData.Filter = "node_name='cur_state'"
                If rsData.RecordCount > 0 Then
                    lngStudyState = Val(Nvl(rsData!node_value))
                    Call ufgStudyList.UpdateSourceData(lngAdviceId, "������", lngStudyState)
                    
                    ufgStudyList.AdoData.Filter = ""
                    '���ݸ��º�����ݣ�ˢ���б���ʾ
                    Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
                End If
            End If
            
        
        Case G_STR_MSG_ZLHIS_PACS_005   '���Σ��ֵ֪ͨ
            '�ڿ����ڵ���Σ������@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & Nvl(rsData!node_value) & "��"
            
            rsData.Filter = "node_name='check_item_title'"
            strHint = strHint & "�����Ŀ " & Nvl(rsData!node_value) & " ����Σ�������"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
        
            '�����б��е���ʾ״̬
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID")
            If lngRowIndex > 0 Then
                Call ufgStudyList.SyncText(lngRowIndex, "Σ��", " ", False)
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "Σ��", 1)
                
                ufgStudyList.AdoData.Filter = ""
                Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
            End If
            
    End Select
    
    
End Sub

Private Sub mobjPacsCore_AfterSaveOuterImage(strStudyUID As String)
    '�������ⲿͼ��ˢ��ͼ��������б�
On Error GoTo ErrHandle
    
    'û�м�¼���˳�
    If mListAdviceInf.lngAdviceId = 0 Then Exit Sub
    
    '�ǵ�ǰ�ļ�飬��ˢ�¼��������б�
    If mListAdviceInf.strStudyUID = strStudyUID Then
        Call mfrmWork_PacsImg.zlRefreshFace(True)
    End If
    
    Exit Sub
ErrHandle:
    '������
End Sub


Public Sub OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'��Ƶ�ɼ������ص��¼�
    mVideoEventInf.vetEventType = lngEventType
    mVideoEventInf.lngAdviceId = lngAdviceId
    mVideoEventInf.lngSendNo = lngSendNo
    mVideoEventInf.strOtherInf = strOther

    timerVideoEvent.Enabled = True
End Sub

Public Sub OnImgLoadState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean)
    If blnLoadFinish Then
        Me.stbThis.Panels(5).Text = ""
    Else
        If blnUpLoad Then
            Me.stbThis.Panels(5).Text = "�����ϴ��ļ�������رչ���վ��"
        Else
            Me.stbThis.Panels(5).Text = "���������ļ�������رչ���վ��"
        End If
    End If
End Sub

Public Sub OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadId As Long)
    Call mdlPACSWork.OnState(blnLoadFinish, blnUpLoad, lngThreadId)
End Sub

Public Sub OnDockClose()
'�������ڹرջص��¼�
End Sub

Private Sub DoOnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'��Ӧ����ģ��ִ�в����󴥷����¼�
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strRoom As String
    Dim strStudyUID As String
    Dim i As Long
    
    Select Case lngEventType
        Case TVideoEventType.vetLockStudy, TVideoEventType.vetUnLockStudy         '�������,�������
            '�޸ı�ǩҳ����ʾ��ʽ�ͱ���
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*Ӱ��ɼ�*" Then
                    If lngEventType = vetLockStudy Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "��" & strOther & "�� Ӱ��ɼ�"
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "Ӱ��ɼ�"
                    End If
                    Exit For
                End If
            Next i
            
     
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)

            
        Case TVideoEventType.vetCaptureFirstImg, TVideoEventType.vetDelAllImg, TVideoEventType.vetUpdateImg  '�ɼ���һ��ͼ��***************************************
            '���¼���б�
            
            strStudyUID = strOther
            
            If lngEventType = TVideoEventType.vetCaptureFirstImg Then
                '����ʱִ�з��û�ΪӰ��ɼ�ϵͳʱִ�з���
                If mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngMoneyExeModle = 1 Then
                    strSql = "Zl_Ӱ�����ִ��(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, "ִ�м�����")
                End If
        
                '��д���ִ�м�
                If mSysPar.lngQueueWay = 1 And mSysPar.blnUseQueue Then
                    strRoom = mSysPar.strLocalRoom
                    If InStr(strRoom, "-") > 0 Then strRoom = Mid(mSysPar.strLocalRoom, 1, InStr(mSysPar.strLocalRoom, "-") - 1)
        
                    strSql = "zl_Ӱ����_����ִ�м�(" & lngAdviceId & ",'" & strRoom & "','" & NeedName(mSysPar.strLocalRoom) & "')"
                    Call zlDatabase.ExecuteProcedure(strSql, GetWindowCaption)
                End If
                
                '���¼���б�
                Call UpdateStudyListState(lngAdviceId, strStudyUID, True, True)
                
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
            ElseIf lngEventType = TVideoEventType.vetDelAllImg Then
                '���¼���б�
                Call UpdateStudyListState(lngAdviceId, strStudyUID, False, True)
            End If


            If Val(ufgStudyList.CurKeyValue) <> CStr(lngAdviceId) Then Exit Sub
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
            'ˢ��Ƕ���ؼ챨��������½�����ͼͼ��
            If lngEventType = TVideoEventType.vetUpdateImg Then If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
        Case TVideoEventType.vetAfterUpdateImg
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub AfterReleationImage(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long, ByVal blnUseMenuReleation As Boolean)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If lngReleationType = 1 Then
        If InStr("345", intStep) > 0 Then
            gstrSQL = "Select ���uid From Ӱ�����¼ Where  ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceId, lngSendNo)
            
            If IsNull(rsTemp!���uid) Then
                '����Ӱ����״̬�������ǰҽ���Ѿ�û��ͼ�񣬶��Ҽ�����Ϊ3�����޸�Ϊ2
                If intStep = 3 Then
                    gstrSQL = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                    zlDatabase.ExecuteProcedure gstrSQL, "ȡ������"
                End If
            End If
        End If
    Else
        '����Ӱ����״̬�����ԭ����״̬���ѱ��������޸ĳ��Ѽ�飬
        If intStep = 2 Then
            '��������Ѿ���ͼ�����޸ĳ��Ѽ��
            strSql = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ���ͼ��", lngAdviceId)
            
            If Not IsNull(rsTemp!���uid) Then
                strSql = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                zlDatabase.ExecuteProcedure strSql, "����Ӱ��"
            End If
        End If
    End If
    
    Call RefreshList
    
    If mblnUseActivexCapture Then
        'ʹ��ActivexExe��Ƶ�ɼ�ͼ��ˢ�´���
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlRefreshData(True)
            Call mobjWork_ActiveVideo.zlUpdateAfterCaptureInfo
        End If
    End If
    
    If Not mfrmWork_PacsImg Is Nothing Then
        Call mfrmWork_PacsImg.zlRefreshFace(True)
    End If
    
    If Not mobjWork_Report Is Nothing And blnUseMenuReleation Then
        Call mobjWork_Report.UpdateVideoCaptureState(TVideoEventType.vetAfterUpdateImg, lngAdviceId, "")
    End If
End Sub

Private Sub mobjWork_Report_AfterReleationImage(ByVal lngOrderID As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long)
On Error GoTo ErrHandle
    Call AfterReleationImage(lngOrderID, lngSendNo, intStep, lngReleationType, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjWork_Report_DocPluginAction(ByVal actionType As Long, ByVal data As String, ByVal tag As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    If actionType = 5 And Trim(data) <> "" And (Trim(tag) = "����ͼ��" Or Trim(tag) = "ȡ������") Then
        '����ҽ��ID��ȡ���ͺźͼ�����
        strSql = "select b.���ͺ�,b.ִ�й��� from  Ӱ�����¼ a,����ҽ������ b where a.ҽ��ID =[1] and a.ҽ��id = b.ҽ��id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", Val(data))
        
        If rsTemp.RecordCount > 0 Then
            Call AfterReleationImage(data, Val(Nvl(rsTemp!���ͺ�)), Val(Nvl(rsTemp!ִ�й���)), IIf(Trim(tag) = "����ͼ��", 2, 1), False)
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optAccept_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optAll_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optFinal_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optNeed_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
'¼���¼�
On Error GoTo ErrHandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long
    
    If Trim(PatiIdentify.GetCurCard.����) = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    If KeyAscii = 13 Then
        Call StartReadCard
        
        Exit Sub
    End If
    
    If PatiIdentify.GetCurCard.�Ƿ�ˢ�� Then
        blnCard = PatiIdentify.zlIsBrushCard(PatiIdentify.objTxtInput, KeyAscii)
            
        If blnCard And Len(PatiIdentify.Text) = PatiIdentify.GetCardNoLen - 1 And KeyAscii <> 8 Then  'ˢ����ϴ���
            PatiIdentify.Text = PatiIdentify.Text & Chr(KeyAscii)
    
            KeyAscii = 0
            
            If PatiIdentify.GetCurCard.�ӿ���� > 0 Then
                Call mobjSquareCard.zlGetPatiID(PatiIdentify.GetCurCard.�ӿ����, PatiIdentify.Text, , lngPatientID)
                
                Call OnFilterRead(PatiIdentify.GetCurCard.����, PatiIdentify.Text, IIf(lngPatientID > 0, lngPatientID, ""))
            Else
                Call OnFilterRead(PatiIdentify.GetCurCard.����, PatiIdentify.Text, "")
            End If
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picAppend_Resize()
On Error GoTo ErrHandle
    labHistory.Left = 120
    labHistory.Top = 120
    
    cboTimes.Left = labHistory.Left + labHistory.Width
    cboTimes.Top = 60
    cboTimes.Width = picAppend.Width - labHistory.Width - lblCash.Width - 360
    
    lblCash.Left = cboTimes.Left + cboTimes.Width + 120
    lblCash.Top = 0
    
    labStudyNum.Left = 120
    labStudyNum.Top = cboTimes.Top + cboTimes.Height + 90
    labStudyNum.Width = picAppend.Width - 240
    
    lbl������Ϣ.Left = 120
    lbl������Ϣ.Top = labStudyNum.Top + labStudyNum.Height + 30
    
    If picAppend.Width > lbl�����Ϣ.Width + lbl������Ϣ.Width + 360 Then
        lbl�����Ϣ.Left = lbl������Ϣ.Left + lbl������Ϣ.Width + 240
        lbl�����Ϣ.Top = lbl������Ϣ.Top
    Else
        lbl�����Ϣ.Left = 120
        lbl�����Ϣ.Top = lbl������Ϣ.Top + lbl������Ϣ.Height + 60
    End If
    
    txtAppend.Top = lbl�����Ϣ.Top + lbl�����Ϣ.Height + 120
    txtAppend.Left = 60
    txtAppend.Width = picAppend.Width - 70
    txtAppend.Height = picAppend.Height - cboTimes.Height - lbl������Ϣ.Height - lbl�����Ϣ.Height - 430
    
ErrHandle:
End Sub



Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long, lngTop As Long, lngRight  As Long, lngBottom  As Long
 On Error GoTo ErrHandle
    
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If Button = 1 Then
        
        '��ֵ�ﵽһ����Χ���˳�����
        If Me.PicLine.Top + Y < lngTop + 700 Or PicLine.Top + Y > PicList.Height - 450 Then
            Exit Sub
        End If

        '�ƶ��ؼ�λ��
        ufgStudyList.Height = ufgStudyList.Height + Y
        PicLine.Top = PicLine.Top + Y
        picAppend.Top = picAppend.Top + Y
        picAppend.Height = picAppend.Height - Y
        txtAppend.Height = txtAppend.Height - Y
    End If
    
ErrHandle:
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim strTemp As String
    
    Select Case control.ID
        Case ID_��Դ
            control.IconId = IIf(Not (mblncmd���� Or mblncmdסԺ Or mblncmd���� Or mblncmd���), 90000, 90001)
            
            strTemp = IIf(mblncmd����, "����", "")
            strTemp = strTemp & IIf(mblncmdסԺ, IIf(strTemp <> "", ",", "") & "סԺ", "")
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd���, IIf(strTemp <> "", ",", "") & "���", "")
            
            If strTemp = "" Then
                strTemp = "��Դ"
                control.ToolTipText = "���ݲ�����Դ���й���"
            Else
                control.ToolTipText = "��ʾ������ԴΪ[" & strTemp & "]�ļ��"
            End If
        
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
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
            control.IconId = IIf(Not (mblncmd�ѽ� Or mblncmdδ�� Or mblncmd���� Or mblncmd�޷� Or mblncmd����), 90000, 90001)
            
            strTemp = strTemp & IIf(mblncmdδ��, IIf(strTemp <> "", ",", "") & "δ��", "")
            strTemp = strTemp & IIf(mblncmd�ѽ�, IIf(strTemp <> "", ",", "") & "�ѽ�", "")
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd�޷�, IIf(strTemp <> "", ",", "") & "�޷�", "")
            
            If strTemp = "" Then
                strTemp = "����"
                control.ToolTipText = "���ݷ���״̬���й���"
            Else
                control.ToolTipText = "��ʾ����״̬Ϊ[" & strTemp & "]�ļ��"
            End If
            
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
        Case ID_�ѽ�
            control.Checked = mblncmd�ѽ�
            control.IconId = IIf(mblncmd�ѽ�, 90001, 90000)
        Case ID_δ��
            control.Checked = mblncmdδ��
            control.IconId = IIf(mblncmdδ��, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�޷�
            control.Checked = mblncmd�޷�
            control.IconId = IIf(mblncmd�޷�, 90001, 90000)
            
            
        Case ID_Ӱ�����
            control.IconId = IIf(mintcmdӰ����� = 0, 90000, 90001)
        Case ID_Ӱ����� + 1 To ID_Ӱ����� + 40
            control.Checked = mblncmdӰ�����(control.ID - ID_Ӱ����� - 1)
            control.IconId = IIf(control.Checked, 90001, 90000)
       
        If control.ID = ID_Ӱ��ִ�м� Then Stop
        Case ID_Ӱ��ִ�м�
            control.IconId = IIf(mintcmdӰ��ִ�м� = 0, 90000, 90001)
        Case ID_Ӱ��ִ�м� + 1 To ID_Ӱ��ִ�м� + 40
            control.Checked = mblncmdӰ��ִ�м�(control.ID - ID_Ӱ��ִ�м� - 1)
            control.IconId = IIf(control.Checked, 90001, 90000)

        Case ID_״̬
            control.IconId = IIf(Not (mblncmd�Ǽ� Or mblncmd���� Or mblncmd��� Or mblncmd���� Or mblncmd��� Or mblncmd���� Or mblncmd���), 90000, 90001)
            
            strTemp = IIf(mblncmd�Ǽ�, "�Ǽ�", "")
            
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd���, IIf(strTemp <> "", ",", "") & "���", "")
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd���, IIf(strTemp <> "", ",", "") & "���", "")
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd���, IIf(strTemp <> "", ",", "") & "���", "")
            
            If strTemp = "" Then
                strTemp = "״̬"
                control.ToolTipText = "���ݼ��״̬���й���"
            Else
                control.ToolTipText = "��ʾ���״̬Ϊ[" & strTemp & "]�ļ��"
            End If
            
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
        
        Case ID_��鲿λ
            control.IconId = IIf(mstrcmd��λ���� = "", 90000, 90001)
        Case ID_��鲿λ + 1 To 4500
            control.Checked = InStr(mstrcmd��λ����, control.Caption) > 0
            control.IconId = IIf(control.Checked, 90001, 90000)
            
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
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
            
        Case ID_�������
            control.IconId = IIf(Not (mblncmd���� Or mblncmd����ʯ�� Or mblncmd���� Or mblncmdϸ�� Or mblncmdʬ�� Or mblncmd����), 90000, 90001)
            
            strTemp = IIf(mblncmd����, "����", "")
            
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmdϸ��, IIf(strTemp <> "", ",", "") & "ϸ��", "")
            strTemp = strTemp & IIf(mblncmdʬ��, IIf(strTemp <> "", ",", "") & "ʬ��", "")
            strTemp = strTemp & IIf(mblncmd����, IIf(strTemp <> "", ",", "") & "����", "")
            strTemp = strTemp & IIf(mblncmd����ʯ��, IIf(strTemp <> "", ",", "") & "����ʯ��", "")
            
            If strTemp = "" Then
                strTemp = "���"
                control.ToolTipText = "���ݲ��������й���"
            Else
                control.ToolTipText = "��ʾ�������Ϊ[" & strTemp & "]�ļ��"
            End If
            
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            
        Case ID_�������_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_�������_����ʯ��
            control.Checked = mblncmd����ʯ��
            control.IconId = IIf(mblncmd����ʯ��, 90001, 90000)
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
            
        Case ID_����סԺ
            control.IconId = IIf(control.Checked, 90001, 90000)
    End Select
    
ErrHandle:
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub

'����ִ��
Private Sub ExecuteStudyMoney()
On Error GoTo ErrHandle
    Dim strSql  As String

    If mListAdviceInf.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    strSql = "Zl_Ӱ�����ִ��(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",2,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
    zlDatabase.ExecuteProcedure strSql, "����ִ��"
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub conMenu_WorkModule_Click()
On Error GoTo ErrHandle
    Dim frmWorkModule As New frmWorkModuleCfg
    
    frmWorkModule.blnIsUseQueue = mSysPar.blnUseQueue
    Call frmWorkModule.ShowWorkModuleCfg(mlngModule, Me)
    
    '�������ù���ģ��ҳ��
    If frmWorkModule.blnIsOk Then
        
        mblnInitOk = False '��ֹ���Ӵ�����ع����ж��Ӵ������ˢ��
        
        Call InitSubForm
        
        mblnInitOk = True
        
        Call ShowTab
        
        Call picWindow_Resize
        
        '���û�м�����ݣ��������������ģ�飬ֻ��ʾģ�鱳��
        If tcDisable.Visible Then Call tcDisable.Translucence
        
        If Not TabWindow.Selected Is Nothing Then Call TabWindow_SelectedChanged(TabWindow.Selected)
        
    End If
    
    Call Unload(frmWorkModule)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal objControl As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim control As XtremeCommandBars.ICommandBarControl
    Dim i As Long
    
    If mblnMenuDownState Then Exit Sub
    
    '������Ҫ����id���Ҷ�Ӧ�Ĳ˵���Ŀ����Ϊͨ���󶨿�ݼ�ִ��ʱ����������һ��ֻ��id��û�������κ���Ϣ��control�˵���
    Set control = cbrMain.FindControl(, objControl.ID, True, True)
    If control Is Nothing Then
        '����ò˵�Ϊ���Ӳ����༭�����Ҽ��˵�������Ҫ�޸��Ҽ��˵���id����Ϣ
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(objControl)
            
            Set control = cbrMain.FindControl(, objControl.ID, True, True)
        End If
        
        If control Is Nothing Then Exit Sub
    End If
    
    If control.ID = 0 Then Exit Sub
    
    mblnMenuDownState = True
        
    cbrMain.RecalcLayout
    
    'ִ��Ӱ��ͼ���Ӧ�Ĺ���
    If Not mfrmWork_PacsImg Is Nothing Then
        If mfrmWork_PacsImg.zlMenu.zlIsModuleMenu(control) Then
            Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(control.ID)
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    If mblnUseActivexCapture Then
        'ʹ��ActivexExc��Ƶ��ʽ��ͼ��ɼ�����
        If Not mobjWork_ActiveVideo Is Nothing Then
'            If mobjWork_ActiveVideo.zlMenu.zlIsModuleMenu(control) Then
'                'ִ��ActivexExe��Ƶ�ɼ���Ӧ�˵�����
'                Call mobjWork_ActiveVideo.zlMenu.zlExecuteMenu(control.ID)
'
'                mblnMenuDownState = False
'                Exit Sub
'            End If
        End If
    End If

    
    'ִ�в������Ӧ����
    If Not mobjWork_Pathol Is Nothing Then
        If mobjWork_Pathol.zlMenu.zlIsModuleMenu(control) Then
            Call mobjWork_Pathol.zlMenu.zlExecuteMenu(control.ID)
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    'ִ��HISģ���Ӧ����
    If Not mobjWork_His Is Nothing Then
        If mobjWork_His.zlMenu.zlIsModuleMenu(control) Then
            If mintChangeUserState = 2 Then  '�������û������������
                MsgBoxD Me, "��ͳһ�û����ٲ�����"
            Else
                Call mobjWork_His.zlMenu.zlExecuteMenu(control.ID)
                
'                '----------------------����ʱ��ִ�з���------------------
'                If control.ID = conMenu_Edit_Append _
'                Or control.ID = conMenu_Edit_Modify _
'                Or control.ID = conMenu_Edit_NewItem * 10# + 1 _
'                Or control.ID = conMenu_Edit_NewItem * 10# + 2 _
'                Or control.ID = conMenu_Edit_NewItem * 10# + 3 Then
'                    If Val(ufgStudyList.CurText("���״̬")) >= 2 Then
'                        Call ExecuteStudyMoney
'                    End If
'                End If
            End If

            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    If Not mobjWork_Report Is Nothing Then
        If mobjWork_Report.zlMenu.zlIsModuleMenu(control) Then
            'ִ�б�����ع���ʱ���������л�������ģ�飬��������ִ��

            If TabWindow.Selected.tag <> "������д" Then
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
                Next
            End If
            
            If control.Caption <> "������ӡ" Then
                If TabWindow.Selected.tag <> "������д" Then
                    mblnMenuDownState = False
                    Exit Sub
                End If
            End If
            
            Call mobjWork_Report.zlMenu.zlExecuteMenu(control.ID)
            
            '�����ѡ�򿪱���ͬʱ��Ƭ��������򿪹�Ƭվ
            If (control.ID = conMenu_PacsReport_Open + 1000000 Or control.ID = conMenu_Edit_Modify + 1000000 _
                Or control.ID = conMenu_Edit_Audit + 1000000 Or control.ID = conMenu_File_Open + 1000000) And mSysPar.blnShowImgAfterReport = True Then
                If Not mfrmWork_PacsImg Is Nothing Then
                    Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(conMenu_Img_Look)
                End If
            End If
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    
    Select Case control.ID

'--------------------------�ļ�------------------
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
            
        Case conMenu_File_Excel '�嵥��ӡ
            Call Menu_File_Excel_click
            
        Case conMenu_File_Parameter '��������
            Call Menu_File_Parmeter_click
            
        Case ConMenu_File_ShortcutSet '��ݼ�����
            Call Menu_File_ShortcutSet_click
            
        Case conMenu_Pathol_WorkModule  'վ��ģʽ����
            Call conMenu_WorkModule_Click
            
        Case conMenu_Manage_SetXWParam  '��������PACS�Ĳ���
            Call Menu_Manage_SetXWParam_click
            
        Case conMenu_File_SendImg '����ͼ��
            Call conMenu_File_SendImg_click
            
        Case conMenu_Cap_DevSet         '��Ƶ����
            If mblnUseActivexCapture Then
                If Not mobjWork_ActiveVideo Is Nothing Then
                    Call mobjWork_ActiveVideo.zlShowVideoConfig
                    mstrCaptureHot = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
                    mstrCaptureAfterHot = GetSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", "F7")
                    mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", "F6")
                End If
            Else
                Exit Sub
            End If
            
        Case conMenu_Manage_ChangeUser
            '�����û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
        
            Call ChangeUser
            
            '�����û�����Ҫˢ�±���༭������Ϊ�û�������ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case conMenu_Manage_SwitchUser
            '�л��û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
            Call SwitchUser
            
            '�л��û�����Ҫˢ�±���༭������Ϊ�û��л���ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
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
            
        Case comMenu_Petition_Capture                       'ɨ�����뵥
            Call Menu_Petition_ɨ�����뵥(1)
            
        Case comMenu_Petition_View
            Call Menu_Petition_ɨ�����뵥(0)                '�鿴���뵥
            
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
        
'        Case conMenu_Manage_ModifBaseInfo               '������Ϣ����
'            Call Menu_Manage_ModifBaseInfo
        
        Case conMenu_Manage_Logout                          'ȡ������
            Call Menu_Manage_ȡ������
            
        Case conMenu_Manage_InQueue                         '�Ŷӽк����
            Call zlInPacsQueue
            
        Case conMenu_Manage_Transfer                        '����Ӱ��
            Call Menu_Manage_����Ӱ��
            
        Case conMenu_Manage_Cancel                          'ȡ������
            Call Menu_Manage_ȡ������
            
        Case conMenu_Manage_Review                          '���
            Call Menu_Manage_���
            
        Case conMenu_Tool_Analyse
            Call OpenViewer(1, mobjPacsCore, mcurAdviceInf.lngAdviceId, True, Me, "", mblnMoved, mSysPar.blnLocalizerBackward)
        
        Case conMenu_Manage_ReportRelease                   '���淢��
            Call Menu_Manage_���淢��
            
        Case conMenu_Manage_FilmRelease                     '��Ƭ����
            Call Menu_Manage_��Ƭ����
            
        Case conMenu_Manage_ReportFilmRelease               '���潺Ƭͬʱ����
            Call Menu_Manage_���潺Ƭͬʱ����
            
        Case conMenu_Manage_SendArrange                     '���Ͱ���
            Call frmSendArrange.ShowMe(Me, mlngCur����ID, mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo)
            
        Case conMenu_Manage_ReportExecutor                  '����ִ�У�����Ǳ�����
            Call Menu_Manage_ReportExecutor
            
        Case conMenu_Manage_CriticalValues, conMenu_Manage_Normal, conMenu_Manage_Critical        'Σ��ֵ�Ǽ�
            Call Menu_Manage_CriticalMark(control.ID)
            
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '���������
            Call Menu_Manage_�������(control.ID)
           
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe   '�������
            Call Menu_Manage_�������(control.ID)
            
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
            
        Case conMenu_Manage_Burn                            'ͼ���¼
            Call Menu_Manage_ͼ���¼
            
        Case conMenu_Manage_CheckList                       '�鿴�������뵥
            Call Menu_Manage_CheckList
            
'----------------------------------------�ղ�---------------------------------------
        Case conMenu_Collection_Manage  '�ղع���
           Call Menu_Manage_�ղع���
        Case conMenu_Collection_To      '�ղص�
           Call Menu_Manage_�ղص�
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '��̬�ղ����˵�
           Call Menu_Manage_�ղ�������ʾ(control, 0)
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999   '�鿴����
           Call Menu_Manage_�ղ�������ʾ(control, 1)
           
           
'----------------------------------------�Զ����ѯ---------------------------------------
        Case conMenu_Manage_ConfigQuery '���ò�ѯ
            Call ShowCustomQueryConfig
            
        Case conMenu_Manage_CustomQuery * 1000# To conMenu_Manage_CustomQuery * 1000# + 9999
            Call ExecuteCustomQuery(control.ID - conMenu_Manage_CustomQuery * 1000#)   'ִ���Զ����ѯ
            
        Case conMenu_Manage_CustomQuery 'ִ���ۺϲ�ѯ
            mblnIsIntegratedQuery = True
            Call Menu_View_Filter_click
            
        Case conMenu_View_Filter '����
                If mlngDefQuerySchemeId >= 0 Then
                    Call ExecuteCustomQuery(mlngDefQuerySchemeId)
                Else
                    mblnIsIntegratedQuery = True
                    Call Menu_View_Filter_click
                End If

'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
            
        Case conMenu_View_FontSize_S    'С����
            Call SetFontSize(0)
        Case conMenu_View_FontSize_M    '������
            Call SetFontSize(1)
        Case conMenu_View_FontSize_L    '������
            Call SetFontSize(2)
            
            
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_Refresh 'ˢ��
            mblnIsCallModuleRefresh = True    '�ֶ�ˢ��ʱ����Ҫ֪ͨ����ģ�������и���
            Call RefreshList
            
            '�����Ŷӽкŵ�ˢ�²������������
            Call RefreshPacsQueueData
        Case comMenu_Cap_Process
            Call Menu_Manage_�����ɼ�
            
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
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|")) + 1
            Call Menu_Dept_Select(control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If control.Parameter <> "" Then 'ִ�з�������ǰģ��ı���
                With ufgStudyList
                    If Val(.CurKeyValue) <> 0 Then
                        Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, _
                            "NO=" & .CurText("NO"), "����=" & .CurText("��¼����"), "ҽ��id=" & Val(.CurKeyValue), 1)
                    Else
                        Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, "", 1)
                    End If
                End With
            End If
        Case Else
            If Val(ufgStudyList.CurKeyValue) = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            Select Case TabWindow.Selected.tag
                    
                    
                Case "�Ŷӽк�"
                    If Not mobjQueue Is Nothing Then
                        If mintChangeUserState = 2 Then  '�������û������������
                            MsgBoxD Me, "��ͳһ�û����ٲ�����"
                        Else
                            mobjQueue.zlExecuteCommandbar control
                        End If
                    End If
                Case "�������", "סԺҽ��", "����ҽ��", "סԺ����", "���ﲡ��", "������Ӳ���", "סԺ���Ӳ���"
                    If Not mobjWork_His Is Nothing Then
                        Call mobjWork_His.zlMenu.zlExecuteMenu(control.ID)
                    End If
                Case "������д"
                    If Not mobjWork_Report Is Nothing Then
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(control.ID)
                    End If
            End Select
            
    End Select
    
    mblnMenuDownState = False
Exit Sub
ErrHandle:
    mblnMenuDownState = False
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub RefreshPacsQueueData()
'ˢ���Ŷ�ģ������
    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlRefreshQueueData(mstrSelQueueRooms)
    End If
End Sub

Private Sub ShowCustomQueryConfig()
'��ʾ�Զ����ѯ����
    Dim frmCusQuery As New frmCustomQueryCfg
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrHandle
    frmCusQuery.Show 1, Me
    
    If frmCusQuery.mblnIsChange Then
        Call RefreshCustomQueryMenu(cbrMain.FindControl(, conMenu_Manage_Query), mlngCur����ID)
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        
        mlngDefQuerySchemeId = -1
        mlngSysQuerySchemeId = -1
        
        Set rsTemp = zlDatabase.OpenSQLRecord("select id,�Ƿ�Ĭ��,�Ƿ�ϵͳ��ѯ from Ӱ���ѯ���� where (�Ƿ�Ĭ��=1 or �Ƿ�ϵͳ��ѯ=1) and( ��������=0 or �������� is null or ��������=[1]) order by �������� desc,�������", "��ȡĬ�Ϲ��˷���", mlngCur����ID)
        rsTemp.Filter = "�Ƿ�Ĭ��=1"
        If rsTemp.RecordCount > 0 Then mlngDefQuerySchemeId = Val(Nvl(rsTemp!ID))
        rsTemp.Filter = "�Ƿ�ϵͳ��ѯ=1"
        If rsTemp.RecordCount > 0 Then mlngSysQuerySchemeId = Val(Nvl(rsTemp!ID))
    End If
    
ErrHandle:
    Unload frmCusQuery
End Sub

Private Sub ExecuteCustomQuery(ByVal lngSchemeId As Long)
    Dim strReturn As String
    Dim strPars As Variant
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strWithCustomQueryTab As String   '�Զ����Ӳ�ѯ
    Dim strWithOrderTab As String   'ҽ���Ӳ�ѯ
    Dim blnEnabledRules As Boolean  '�Ƿ������˹���
    
    '�����Զ����ѯ����ʱ�����жϴ˷����Ƿ������¼����
    strSql = "select id from Ӱ���ѯ���� where ����id=[1] and rownum<=1"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", lngSchemeId)
    
    If rsData.RecordCount <= 0 Then
        '�˷�����������¼����
        Call frmCustomQueryCall.GetQuerySqlAndPars(lngSchemeId, strReturn, strPars)
    Else
        strReturn = frmCustomQueryCall.ShowCustomQuery(lngSchemeId, IIf(mblnAllDepts, 0, mlngCur����ID), mlngModule, strPars, Me)
    End If
    
    If Trim(strReturn) = "" Then Exit Sub
    
    blnEnabledRules = frmCustomQueryCall.mintEnabledRules = 1
    
    strWithCustomQueryTab = " with tmpCustom  as (" & strReturn & ") "
    
    strWithOrderTab = "tmpOrder as ( select Id as ID from tmpCustom Union All select a.ID from ����ҽ����¼ a, tmpCustom b where a.���ID=b.ID and a.ҽ��״̬ <> 4 ) "
    
    'ִ���Զ����ѯ
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        '����ɾ���ò�ѯ�еġ�Ӱ������Ŀ�����ݱ���Ϊɾ���󣬻�������ݵĲ�ѯЧ�ʽϵͣ�ɾ��������Ҫʹ�ò���ҽ�����͵�ִ�в���ID��Ϊ�������˼�飬Ȼ�����ֶ�û��������
        strSql = "Select  Distinct" & vbNewLine & _
                    "       A.ҽ��ID,B.���ID,A.���ͺ�,A.�״�ʱ�� ����ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.ִ�м�,A.������� ����,h.Σ��״̬ Σ��," & vbNewLine & _
                    "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,B.������Դ ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
                    "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��," & vbNewLine & _
                    "       Nvl(B.����,H.����) ����,G.Ӱ�����,H.����,Nvl(B.�Ա�,H.�Ա�) �Ա�,Nvl(B.����,H.����) ����,H.���,H.����,H.Ӱ������,H.��������,H.�������," & vbNewLine & _
                    "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������, H.���淢��,H.���Ž�Ƭ,H.����ID,A.��¼����, " & vbNewLine & _
                    "       H.�����,H.�Ƿ���ӽ�Ƭ,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.������,H.�Ƿ�ʦȷ��,H.��鼼ʦ,H.��鼼ʦ��,H.�������� ��ͼʱ��," & vbNewLine & _
                    "       H.�������,H.��Ϸ���,H.���UID,H.ͼ��λ��,A.ִ�в���ID as ִ�п���ID,0 as ת��,F.���� AS ���˿���, a.����ʱ��, " & vbNewLine & _
                    "       C.���￨��,A.NO as ���ݺ�,C.���֤��,C.����ģʽ,decode(B.������Դ,2,D.��������,C.��������) as ��������,D.·��״̬,A.�Ʒ�״̬,Decode(A.��¼����,2,1,Decode(a.�Ʒ�״̬,3,1,0)) as �շ� ,m.ҽ��ID as ���뵥ҽ�� " & vbNewLine & _
                    " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,������ҳ D,Ӱ�����¼ H,Ӱ������Ŀ G,���ű� F,Ӱ�����뵥ͼ�� m , tmpOrder " & vbNewLine & _
                    " Where A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) " & vbNewLine & _
                    "       And B.������ĿID=G.������ĿID And B.����ID=C.����ID And B.���˿���id=F.ID" & vbNewLine & _
                    "       And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) and a.ҽ��id = m.ҽ��id(+) and a.ҽ��Id=tmpOrder.ID and b.ҽ��״̬ <> 4 "

    Else
        '���ﵥ���Բ���Ĳ�ѯ���д�����Ϊ������Ҫ������һЩ��ѯ�����ݱ�
        strSql = "Select Distinct" & vbNewLine & _
                    "       A.ҽ��ID,B.���ID,A.���ͺ�,A.�״�ʱ�� ����ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.������� ����,h.Σ��״̬ Σ��," & vbNewLine & _
                    "       '' as ����ִ��״̬, o.ȡ�Ĺ���,o.��Ƭ����,o.���߹���,o.���ӹ���,o.��Ⱦ����, " & vbNewLine & _
                    "       decode(o.�������,0,'����',1,'����',2,'ϸ��',3,'����',4,'ʬ��',5,'����ʯ��',null) as  ������, " & vbNewLine & _
                    "       decode(o.�����,null,'δ����','�Ѻ���') as �������, A.ִ�в���ID as ִ�п���ID, " & vbNewLine & _
                    "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID, B.������Դ ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
                    "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��," & vbNewLine & _
                    "       Nvl(B.����,H.����) ����,Nvl(B.�Ա�,H.�Ա�) �Ա�,Nvl(B.����,H.����) ����,H.���,H.����,o.�ۺ�����," & vbNewLine & _
                    "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������,o.�����,H.���淢��,H.���Ž�Ƭ,H.����ID,A.��¼����, " & vbNewLine & _
                    "       H.�����,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.��������,H.������,H.������,H.�Ƿ�ʦȷ��,H.��鼼ʦ,H.�������� ��ͼʱ��, " & vbNewLine & _
                    "       H.�������,H.��Ϸ���,H.���UID,H.ͼ��λ��,0 as ת��,F.���� AS ���˿���, a.����ʱ��, t.��ǰ״̬ as ����״̬, t.����ҽʦ,t.ID as ����ID, " & vbNewLine & _
                    "       C.���￨��,A.NO as ���ݺ�,C.���֤��,C.����ģʽ,decode(B.������Դ,2,D.��������,C.��������) as ��������,D.·��״̬,A.�Ʒ�״̬,Decode(A.��¼����,2,1,Decode(a.�Ʒ�״̬,3,1,0)) as �շ� ,m.ҽ��ID as ���뵥ҽ��, " & vbNewLine & _
                    "      (select count(1) from ��������Ϣ V , ����������Ϣ W where V.����ҽ��ID=w.����ҽ��id and v.ҽ��id=A.ҽ��ID and w.����״̬=1) as ���� " & vbNewLine & _
                    " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,������ҳ D,Ӱ�����¼ H,���ű� F, " & vbNewLine & _
                    "       ��������Ϣ o ,Ӱ�����뵥ͼ�� m,  ���������Ϣ t, tmpOrder " & _
                    " Where A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) " & vbNewLine & _
                    "       And B.����ID=C.����ID And B.���˿���id=F.ID " & vbNewLine & _
                    "       and A.ҽ��ID=o.ҽ��ID(+) " & vbNewLine & _
                    "       And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) and a.ҽ��id = m.ҽ��id(+) and o.����ҽ��ID=t.����ҽ��ID(+) and a.ҽ��id= tmpOrder.ID and b.ҽ��״̬ <> 4 "
    End If
    
    strSql = "select " & IIf(blnEnabledRules, "/*+ RULE*/", "") & " * from (" & strWithCustomQueryTab & "," & strWithOrderTab & strSql & ")"
    
    Set ufgStudyList.AdoData = GetDataToLocal(strSql, "�Զ����ѯ", strPars(1), strPars(2), strPars(3), strPars(4), strPars(5), strPars(6), strPars(7), strPars(8), strPars(9), strPars(10), _
                                            strPars(11), strPars(12), strPars(13), strPars(14), strPars(15), strPars(16), strPars(17), strPars(18), strPars(19), strPars(20))
    
    ufgStudyList.AdoFilter = GetFilterWhere
    
    '��binddata�ķ�����ʹ��refreshdata�ķ�����
    Call ufgStudyList.BindData
    
    '�ָ�����
    Call ufgStudyList.ResetSort(mlngSortCol, mintSortOrder)
    
    Call RefreshStatusBarInf
 
    If ufgStudyList.GridRows > 1 Then
        Call ufgStudyList.LocateRow(1)
        Call ufgStudyList_OnSelChange
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    '���������С
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, IIf(bytSize = 2, 15, bytSize)))
    
    Call ReMoveCtrl(mbytFontSize)
    Call ReSetFormFontSize
    Call ReSetModuleFontSize(mbytFontSize, IIf(bytSize = 2, 1, bytSize))
    Call SetSelectRowColor
End Sub


Private Sub ReSetModuleFontSize(ByVal bytFontSize As Byte, ByVal bytSize As Byte)
'����:�������ø���ҵ��ģ�鴰��������С
    On Error Resume Next
    
    '�ж� ��ǰѡ�е�
    Select Case mlngModule
        Case 1290
            If Not mfrmWork_PacsImg Is Nothing Then
                If TabWindow.Selected.tag = "Ӱ��ͼ��" Then
                    Call mfrmWork_PacsImg.ReSetFormFontSize(mbytFontSize)
                End If
            End If
            
            If Not mobjWork_His Is Nothing Then
                If Not mobjWork_His.GetExpenseObj Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(bytSize)
                If Not mobjWork_His.GetAdviceObj Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
                If Not mobjWork_His.GetEPRsObj Is Nothing Then Call mobjWork_His.GetEPRsObj.SetFontSize(bytSize)
            End If
            
        Case 1291
            If Not mobjWork_His Is Nothing Then
               If Not mobjWork_His.GetExpenseObj Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(bytSize)
               If Not mobjWork_His.GetAdviceObj Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
               If Not mobjWork_His.GetEPRsObj Is Nothing Then Call mobjWork_His.GetEPRsObj.SetFontSize(bytSize)
            End If
            
        Case 1294
        
            If Not mobjWork_Pathol Is Nothing Then
                Select Case TabWindow.Selected.tag
                    Case "�걾����"
                        Call mobjWork_Pathol.GetModule(mtSpecimen).ReSetFormFontSize(mbytFontSize)
                        
                    Case "����ȡ��"
                        Call mobjWork_Pathol.GetModule(mtMaterial).ReSetFormFontSize(mbytFontSize)
                        
                    Case "������Ƭ"
                        Call mobjWork_Pathol.GetModule(mtSlices).ReSetFormFontSize(mbytFontSize)
                        
                        
                    Case "�����ؼ�"
                        Call mobjWork_Pathol.GetModule(mtSpeExam).ReSetFormFontSize(mbytFontSize)
                        
                    Case "���̱���"
                        Call mobjWork_Pathol.GetModule(mtProRep).ReSetFormFontSize(mbytFontSize)
                        
                    Case "�������"
                        If Not mobjWork_His Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(mbytFontSize, bytSize)
                        
                    Case "����ҽ��", "סԺҽ��"
                        If Not mobjWork_His Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
                    
                End Select
            End If
    End Select
End Sub

Private Sub ReSetFormFontSize()
'����:�������ù���վ����������С
    On Error Resume Next
    
    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType As String
    
    Me.FontSize = mbytFontSize
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    CtlFont.Name = strFontType
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            If objCtrl.Name <> "lblCash" Then
                objCtrl.Font.Name = strFontType
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("��") + 60
            End If
        Case UCase("vsFlexGrid")
        
            CtlFont.Name = strFontType
            CtlFont.Size = mbytFontSize
            objCtrl.DataGrid.Font = CtlFont
            
        Case UCase("ucFlexGrid")
            ufgStudyList.HeadFont.Size = mbytFontSize
            objCtrl.DataGrid.Cell(flexcpFontSize, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = mbytFontSize
            objCtrl.DataGrid.FontName = strFontType
            objCtrl.DataGrid.FontSize = mbytFontSize
            objCtrl.DataGrid.RowHeight(0) = TextHeight("��") + 150
        Case UCase("ComboBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontName = strFontType
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    
    Call picAppend_Resize
    
End Sub
Private Sub ReMoveCtrl(ByVal bytFontSize As Byte)
'����:�ƶ��ؼ�λ��
    On Error GoTo ErrHandle
    Dim lngMove As Long '�ؼ��ƶ�����
    
    mbytFontSize = bytFontSize
    lngMove = IIf(bytFontSize = 9, 1200, IIf(bytFontSize = 12, 1500, 2000))
    
    
    If glngModul = 1294 Then
        optAccept.Left = optNeed.Left + lngMove
        optFinal.Left = optAccept.Left + lngMove
        optAll.Left = optFinal.Left + lngMove
        tabFilter.Width = IIf(bytFontSize = 15, 4800, 4000)
        tabFilter.Height = IIf(bytFontSize = 9, 375, IIf(bytFontSize = 12, 400, 425))
    End If
    
    '���ò�����ϸ��Ϣ �������÷���
    Call picAppend_Resize
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo ErrHandle

    If mfrmPACSFilter Is Nothing Then Set mfrmPACSFilter = New frmPACSFilter
    
    With mfrmPACSFilter
        .mlngModul = mlngModule
        .mBeforeDays = mSysPar.lngBeforeDays - 1
        .mDept = mlngCur����ID '��ǰ����
        .Show 1, Me
        If Not .mblnOK Then Exit Sub 'û�з�������
        
        '��ʹ��ʱ������ʱ����չ̶�����
        PatiIdentify.Text = ""
        SQLCondition.���� = ""
        SQLCondition.���￨ = ""
        SQLCondition.����� = 0
        SQLCondition.סԺ�� = 0
        SQLCondition.������ = ""
        SQLCondition.���ݺ� = ""
        SQLCondition.���� = 0
        SQLCondition.���֤ = ""
        SQLCondition.IC�� = ""
        SQLCondition.������� = -1
        
        
        SQLCondition.��ʼʱ�� = Format(.dtpBegin.value, "yyyy-MM-dd HH:mm:00")
        SQLCondition.����ʱ�� = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
        
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
        
        If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
            gblnXWMoved = mblnMoved
        End If
        
        If .optFindType(1).value = True Then 'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ���4=����ʱ�䣨����ҽ����¼.��ʼִ��ʱ�䣩
            SQLCondition.ʱ������ = 1
        ElseIf .optFindType(2).value = True Then
            SQLCondition.ʱ������ = 2
        ElseIf .optFindType(3).value = True Then
            SQLCondition.ʱ������ = 3
        Else
            SQLCondition.ʱ������ = 4
        End If
        
        If NeedName(.cboPart.Text) <> "���в�λ" Then '���걾��λ
            SQLCondition.�걾��λ = NeedName(.cboPart.Text)
        Else
            SQLCondition.�걾��λ = ""
        End If
        
        '�����Ա�
        If NeedName(.CboSex.Text) = "ȫ��" Then
            SQLCondition.�Ա� = ""
        Else
            SQLCondition.�Ա� = NeedName(.CboSex.Text)
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
        
        If NeedName(.cboDept.Text) <> "���п���" Then '���˿���
            SQLCondition.���˿��� = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            SQLCondition.���˿��� = 0
        End If

        If NeedName(.cboDiagDOC.Text) <> "����ҽ��" Then '���ҽ��
            SQLCondition.���ҽ�� = NeedName(.cboDiagDOC.Text)
        Else
            SQLCondition.���ҽ�� = ""
        End If
        
        If NeedName(.cboAuditing.Text) <> "����ҽ��" Then '���ҽ��
            SQLCondition.���ҽ�� = NeedName(.cboAuditing.Text)
        Else
            SQLCondition.���ҽ�� = ""
        End If
       
      
        If .cboModality.Text <> "�������" Then 'Ӱ�����
            SQLCondition.Ӱ����� = Split(.cboModality.Text, "-")(1)
        Else
            SQLCondition.Ӱ����� = ""
        End If
        
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
        
        If NeedName(.cbo��鼼ʦ.Text) = "����ҽ��" Then
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
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
On Error GoTo ErrHandle
    Dim objControl As CommandBarControl, i As Integer
    Dim aryKindInfo() As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    If mlngModule = G_LNG_PACSSTATION_MODULE Then
                        'ֻ��ҽ����Ҫ��ӡ�ȫ�����ҡ��Ŀ���ѡ��˵�
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "ȫ������")
                    
                        objControl.Category = "Main"
                        objControl.DescriptionText = 0
                        If mblnAllDepts = True Then objControl.Checked = True
                    End If
                    
                    '�����ÿһ���������
                    For i = 0 To UBound(Split(mstrCanUse����, "|"))  'mstrCanUse����=id_����-����|id_����-����
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstrCanUse����, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse����, "|")(i), "_")(0)
                        
                        If mblnAllDepts = False And mlngCur����ID = objControl.DescriptionText Then
                            objControl.Checked = True
                        End If
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.tag
                Case "סԺҽ��", "����ҽ��", "�������"
                    Call mobjWork_His.zlMenu.zlRefreshSubMenu(CommandBar)
            End Select
    End Select
ErrHandle:
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim blnNoRecord As Boolean
    Dim intState As Integer
    Dim blnCancel As Boolean
    Dim tt As CommandBarControl
    Dim objControl As XtremeCommandBars.ICommandBarControl
    
    If Not mblnInitOk Then Exit Sub
    
    '����ò˵�Ϊ���Ӳ����༭�����Ҽ��˵�������Ҫ�޸Ĳ˵�id����Ϣ
    Set objControl = cbrMain.FindControl(, control.ID, True, True)
    If objControl Is Nothing Then
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(control)
        End If
    End If
    
    If ufgStudyList.GridCols <= 1 Or ufgStudyList.GridRows <= 1 Or Not ufgStudyList.IsSelectionRow Then
        blnNoRecord = True
    Else
        blnNoRecord = Val(ufgStudyList.CurKeyValue) = 0
    End If
    
    If Not blnNoRecord Then
        intState = Val(ufgStudyList.CurText("���״̬"))
        blnCancel = ufgStudyList.CurText("������") = "�Ѿܾ�"
    End If
    
    If TabWindow.ItemCount > 0 Then
        If TabWindow.Selected Is Nothing Then Exit Sub
        
        '����Ӱ��ͼ��˵�
        If Not mfrmWork_PacsImg Is Nothing Then
            If mfrmWork_PacsImg.zlMenu.zlIsModuleMenu(control) Then
                Call mfrmWork_PacsImg.zlMenu.zlUpdateMenu(control)
                Exit Sub
            End If
        End If
        
        '���²�����˵�
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu(control) Then

                Select Case control.ID
                    Case conMenu_PatholSpecimen
                        control.Visible = IIf(TabWindow.Selected.tag = "�걾����", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholMaterial
                        control.Visible = IIf(TabWindow.Selected.tag = "����ȡ��", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholSlices
                        control.Visible = IIf(TabWindow.Selected.tag = "������Ƭ", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholSpeExam
                        control.Visible = IIf(TabWindow.Selected.tag = "�����ؼ�", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholProRep
                        control.Visible = IIf(TabWindow.Selected.tag = "���̱���", True, False)
                        
                        Exit Sub
                End Select
                
                Call mobjWork_Pathol.zlMenu.zlUpdateMenu(control)
                
                Exit Sub
            End If
        End If
        
        '����HISģ��˵�
        If Not mobjWork_His Is Nothing Then
            
            If InStr("�������, סԺҽ��, ����ҽ��, סԺ����, ���ﲡ��, ������Ӳ���, סԺ���Ӳ���", TabWindow.Selected.tag) > 0 Then
                If mobjWork_His.zlMenu.zlIsModuleMenu(control) Then
                    Call mobjWork_His.zlMenu.zlUpdateMenu(control)
                    
                    '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
                    If Val(ufgStudyList.CurText("���״̬")) = 6 Then
                        Select Case control.ID
                            Case conMenu_Edit_MarkMap, conMenu_Tool_PlugIn, conMenu_Edit_Compend, conMenu_Manage_ReportLisView, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                                control.Enabled = True
                            Case conMenu_Edit_Copy, conMenu_File_ExportToXML, conMenu_Tool_Search, conMenu_File_Open, conMenu_EditPopup, conMenu_Edit_ChargeDelAudit
                                '�⼸���˵�������
                            Case Else
                                control.Enabled = False
                        End Select
                    End If
                    
                    Exit Sub
                End If
            End If
        End If
        
        If mblnUseActivexCapture Then
            '����ʹ��ActivexExe��ʽ����Ƶ�ɼ��˵�
            If Not mobjWork_ActiveVideo Is Nothing Then
'                If mobjWork_ActiveVideo.zlMenu.zlIsModuleMenu(control) Then
'                    '������Ƶ�ɼ��˵�...
'                    Call mobjWork_ActiveVideo.zlMenu.zlUpdateMenu(control)
'                    Exit Sub
'                End If
            End If
        End If

        
        '���±���ģ��˵�
        If Not mobjWork_Report Is Nothing Then
            If mobjWork_Report.zlMenu.zlIsModuleMenu(control) Then
                Call mobjWork_Report.zlMenu.zlUpdateMenu(control)
                
                '��ǰ�鿴�������μ�¼��˵���������
                If cboTimes.ListIndex <> -1 Then
                    If mListAdviceInf.lngAdviceId <> cboTimes.ItemData(cboTimes.ListIndex) Then
                        If control.ID = conMenu_Edit_Copy + 1000000 Or control.ID = conMenu_File_ExportToXML + 1000000 Or control.ID = conMenu_EditPopup + 1000000 _
                            Or control.ID = conMenu_Tool_Search + 1000000 Or control.ID = conMenu_File_Preview + 1000000 Or control.ID = conMenu_File_Print + 1000000 Or control.ID = conMenu_File_NoAskPrint + 1000000 Then
                            '�⼸���˵�������
                        Else
                            control.Enabled = False
                        End If
                    End If
                End If
            
                Exit Sub
            End If
        End If
    End If
    
    
    Select Case control.ID
        Case conMenu_Manage_LocateValue
            control.Enabled = Not blnNoRecord
        Case comMenu_Cap_Process
            control.Enabled = True 'Not blnNoRecord
        Case conMenu_View_Filter * 10#
            control.Caption = "��ǰ����:" & IIf(mblnAllDepts = True, "ȫ������", mstrCur����)
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|")) + 1
            If mblnAllDepts = True Then
                control.Checked = (control.DescriptionText = 0)
            Else
                control.Checked = (control.DescriptionText = mlngCur����ID)
            End If
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
        Case conMenu_Manage_CopyCheck '���ƵǼ�
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
                control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
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
                control.Enabled = intState < 6 And Not blnCancel
            Else
                control.Enabled = False
            End If
'        Case conMenu_Manage_ModifBaseInfo '������Ϣ����
'            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
'                control.Visible = False
'            ElseIf Not blnNoRecord Then
'                control.Enabled = intState < 6 And Not blnCancel
'            Else
'                control.Enabled = False
'            End If
        Case conMenu_Manage_Receive   '��鱨��(&L)
            If InStr(mstrPrivs, "��鱨��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   'ȡ������(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent Is Nothing Then '��ʹ���ȼ�ʱ��������ж�parent����������쳣
                Exit Sub
            ElseIf control.Parent.type = xtpControlPopup Then
                If InStr(mstrPrivs, "ȡ������") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "ȡ������"
                    control.Caption = "ȡ������(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' �������е���ȡ��������ȡ���Ǽ�,ͬһ�������ȡ���ǼǺ�ȡ����鹦��
                control.Visible = IIf(intState <= 1 And intState <> -1, InStr(mstrPrivs, "���Ǽ�") > 0, InStr(mstrPrivs, "ȡ������") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And intState <> -1 And Not blnCancel) '���ܾ��Ĳ��ܱ��ٴξܾ�
                control.ToolTipText = IIf(intState <= 1 And intState <> -1, "ȡ���Ǽ�", "ȡ������")
                control.Caption = "ȡ��"
            End If
        Case conMenu_Manage_InQueue    '�Ŷӽк����
            control.Visible = mSysPar.blnUseQueue And Not mSysPar.blnAutoInQueue
            control.Enabled = (intState >= 2 And intState <= 5)
            
        Case conMenu_Manage_Transfer   '����Ӱ��(&C)
            If InStr(mstrPrivs, "ͼ�����") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
            End If
        Case conMenu_Manage_Cancel   'ȡ������(&B)
            If InStr(mstrPrivs, "ͼ�����") <= 0 Then
                control.Visible = False
            ElseIf (intState >= 2 And intState <= 5) Or intState = -1 Then
                control.Enabled = ufgStudyList.CurText("���UID") <> ""
            Else
                control.Enabled = False
            End If
            
        Case conMenu_Manage_Review  '���
            If InStr(mstrPrivs, "���") <= 0 Then
                control.Visible = False
            ElseIf (Not blnNoRecord And intState > 1 And intState <= 6) Or intState = -1 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Tool_Analyse   '�߼�ͼ����
            If InStr(GetPrivFunc(glngSys, 1289), "����") <= 0 Then
                control.Visible = False
            ElseIf (Not blnNoRecord And intState > 1 And intState < 6) Or intState = -1 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Release, conMenu_Manage_ReportFilmRelease     '���淢��,��������ɺ󶼿���ִ��

            control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
            If mrtReportType = �����ĵ��༭�� Then
                If control.ID = conMenu_Manage_ReportFilmRelease Then
                    If mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) = 3 And Nvl(ufgStudyList.CurText("��Ƭ����"), "0") = "1" Then
                        control.Caption = "�ջ�����"
                        control.ToolTipText = "�ջ��Ѿ����ŵı����Ƭ"
                    Else
                        control.Caption = "��������"
                        control.ToolTipText = "����ͽ�Ƭͬʱ����"



                    End If
                End If
            Else
                If Not blnNoRecord Then
                  '�޸ı��淢�Ű�ť�ı���
                     If Not blnNoRecord Then
                         If Nvl(ufgStudyList.CurText("���淢��"), "0") = "1" And Nvl(ufgStudyList.CurText("��Ƭ����"), "0") = "1" Then
                             control.Caption = "�ջ�"
                             control.ToolTipText = "�ջ��Ѿ����ŵı����Ƭ"
                         Else
                             control.Caption = "����"
                             control.ToolTipText = IIf(control.ID = conMenu_Manage_Release, "�����Ƭ����", "����ͽ�Ƭͬʱ����")
                         End If
                     End If
                End If
            End If
            
            control.Enabled = Not control.Enabled
            control.Enabled = Not control.Enabled
        Case conMenu_Manage_FilmRelease
            control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
            If Not blnNoRecord Then
                 If Nvl(ufgStudyList.CurText("��Ƭ����"), "0") = "1" Then
                     control.Caption = "��Ƭ�ջ�"
                     control.ToolTipText = "�ջ��Ѿ����ŵĽ�Ƭ"
                     
                     If InStr(mstrPrivs, "ȡ������") > 0 Then
                         control.Enabled = True
                     Else
                         control.Enabled = False
                     End If
                 Else
                     control.Caption = "��Ƭ����"
                     control.ToolTipText = "��Ƭ����"
                 End If
             End If


        Case conMenu_Manage_ReportRelease
            control.Enabled = IIf(intState >= 2 And intState <= 6, True, False)
            
            If mlngModule = G_LNG_PATHOLSYS_NUM Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                If Not blnNoRecord Then
                    If mrtReportType = �����ĵ��༭�� Then
                        If mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) > 1 Then
                            control.Caption = "�����ջ�"
                            control.ToolTipText = "�ջ��Ѿ����ŵı���"
                        Else
                            control.Caption = "���淢��"
                            control.ToolTipText = "���淢��"
                        End If
                    Else
                        If ufgStudyList.CurText("���淢��") = "1" Then
                            control.Caption = "�����ջ�"
                            control.ToolTipText = "�ջ��Ѿ����ŵı���"
                        Else
                            control.Caption = "���淢��"
                            control.ToolTipText = "���淢��"
                        End If
                    End If
                End If
            End If
            control.Enabled = Not control.Enabled
            control.Enabled = Not control.Enabled
        
        Case conMenu_Manage_SendArrange                     '���Ͱ���
            control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
        Case conMenu_Manage_ReportExecutor      '����ִ��
            control.Enabled = IIf(intState >= 2 And intState <= 6, True, False)
        
        Case conMenu_Manage_CriticalValues, conMenu_Manage_CriticalSituation, conMenu_Manage_Normal, conMenu_Manage_Critical 'Σ��ֵ
            If mSysPar.lngCriticalValues = 0 Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
            End If

        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive   '���������(&X)
            If mSysPar.blnIgnoreResult = True Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
                If ufgStudyList.CurText("Σ��") = " " And control.ID = conMenu_Manage_Result Then control.Enabled = False
            End If
            
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe, conMenu_Manage_FuHeLevel '�������
            If mSysPar.lngConformDetermine = 0 Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
            End If
        
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '��ɫͨ�����/ȡ��
            If InStr(mstrPrivs, "��ɫͨ��") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
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
                control.Enabled = IIf(mrtReportType = �����ĵ��༭��, mobjWork_Report.GetCurrDocId(mcurAdviceInf.lngAdviceId) = "", ufgStudyList.CurText("������") = "")
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '������(&E)
            If InStr(mstrPrivs, "������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState < 6
            End If
        Case conMenu_Manage_Undone   'ȡ�����(&U)
            If InStr(mstrPrivs, "ȡ��������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 6
            End If
        Case conMenu_File_SendImg  '����ͼ��
            If InStr(mstrPrivs, "�ļ�����") <= 0 Then control.Visible = False
        Case conMenu_Img_Contrast, conMenu_Img_Look     'Ӱ��Ա�,Ӱ���Ƭ
            If mblnObserve Then
                If blnNoRecord Then control.Enabled = False: Exit Sub

                control.Enabled = mcurAdviceInf.strStudyUID <> ""
            Else
                control.Visible = False
            End If
        Case conMenu_Manage_RelatingPatiet  '��������
            If InStr(mstrPrivs, "��������") <= 0 Or mSysPar.blnRelatingPatient = False Then
                control.Visible = False
            ElseIf blnNoRecord Or (intState < 2 And intState <> -1) Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_Manage_Burn
            control.Visible = IIf(InStr(mstrPrivs, "ͼ���¼") <= 0, False, True)
        
        Case conMenu_Manage_CheckList   '�鿴���뵥
            control.Visible = True
            If mListAdviceInf.lngAdviceId > 0 And mListAdviceInf.lngPatientFrom <> 3 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_File_SendImg
            If InStr(mstrPrivs, "�ļ�����") <= 0 Then control.Visible = False
        Case conMenu_File_PrintSet     '��ӡ����(&S)
        Case conMenu_File_Excel         '�嵥��ӡ(&L)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter, conMenu_Cap_DevSet
        
        Case conMenu_Manage_ChangeUser  '�û�����
            If mSysPar.blnChangeUser Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            
        Case conMenu_Manage_SwitchUser  '�л��û�
            If mSysPar.blnSwitchUser Then
                control.Visible = True
            Else
                control.Visible = False
            End If
        
        Case conMenu_Manage_SetXWParam      '����PACS�������ã�����д˲˵�������ʾ
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_Help_Help, conMenu_Help_About  '����
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '����WEB
        Case conMenu_File_Exit
        Case ConMenu_File_ShortcutSet
        Case conMenu_Pathol_WorkModule
        Case conMenu_View_ToolBar
        Case conMenu_Manage_Query
        Case conMenu_Manage_CustomQuery * 1000# To conMenu_Manage_CustomQuery * 1000# + 999
        Case conMenu_Manage_CustomQuery
        Case conMenu_Manage_ConfigQuery '��ѯ����
            control.Visible = CheckPopedom(mstrPrivs, "��ѯ����")
        Case conMenu_Cap_DevSet     'Ӱ���豸����
        Case conMenu_Manage_Change_In   '�����б�
        Case conMenu_Img_3D_MMPR, conMenu_Img_3D_MPR, conMenu_Img_3D_PF, conMenu_Img_3D_SA, conMenu_Img_3D_VA, conMenu_Img_3D_VE '��ά�ؽ��ļ����Ӳ˵�����Ҫ����
        Case conMenu_View_FontSize_S    'С����
             control.Checked = mbytFontSize = 9
        Case conMenu_View_FontSize_M    '������
             control.Checked = mbytFontSize = 12
        Case conMenu_View_FontSize_L    '������
             control.Checked = mbytFontSize = 15
        
   '-------------------------------------------------�ղع�����----------------------------------------------------------
 
        Case conMenu_Collection    '�ղ�(&C)
            control.Enabled = True
        Case conMenu_Collection_Manage  '�ղع���˵�
            control.Enabled = True
        Case conMenu_Collection_ViewShare      '�鿴����
            control.Enabled = True
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '��̬�ղز˵�
            control.Enabled = True
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999  '��̬����˵�
            control.Enabled = True
         Case conMenu_Collection_To
            
            
    '-------------------------------------------ɨ�����뵥����-----------------------------------------------

        'ɨ�����뵥
        Case comMenu_Petition_Capture
            If Val(ufgStudyList.CurKeyValue) = 0 Or blnCancel Then
                control.Enabled = False
            Else
                control.Enabled = IIf((intState >= 2 And intState <= 5) Or intState = -1, True, False)
            End If
            
        '�鿴���뵥
        Case comMenu_Petition_View
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Enabled = False
            End If
            
        Case Else
            If blnNoRecord Then
                control.Enabled = False
                Exit Sub
            End If
                    
            
            '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
            If Val(ufgStudyList.CurText("���״̬")) = 6 Then
                control.Enabled = False
            End If
            
    End Select
ErrHandle:
End Sub

Private Sub InitModuleParameter(Optional blnIsUpdateSearchTime As Boolean = True)
'����:��ʼ��ģ�鼶����,���������ʱ����һ��
    Dim rsTemp As ADODB.Recordset
    
    '��ȡĬ�ϵĲ�ѯ����id��ϵͳ��ѯ����id
    mlngDefQuerySchemeId = -1
    mlngSysQuerySchemeId = -1
    
    Set rsTemp = zlDatabase.OpenSQLRecord("select id,�Ƿ�Ĭ��,�Ƿ�ϵͳ��ѯ from Ӱ���ѯ���� where (�Ƿ�Ĭ��=1 or �Ƿ�ϵͳ��ѯ=1) and( ��������=0 or �������� is null or ��������=[1]) order by �������� desc,�������", "��ȡĬ�Ϲ��˷���", mlngCur����ID)
    rsTemp.Filter = "�Ƿ�Ĭ��=1"
    If rsTemp.RecordCount > 0 Then mlngDefQuerySchemeId = Val(Nvl(rsTemp!ID))
    rsTemp.Filter = "�Ƿ�ϵͳ��ѯ=1"
    If rsTemp.RecordCount > 0 Then mlngSysQuerySchemeId = Val(Nvl(rsTemp!ID))
    
    mSysPar.lngListColorMark = Nvl(GetDeptPara(mlngCur����ID, "��ɫ��ʾ����", 0))
    mSysPar.blnNameColColorCfg = GetDeptPara(mlngCur����ID, "������ɫ����", 0) = "1"         '������ɫ����
    mSysPar.blnOrdinaryNameColColorCfg = GetDeptPara(mlngCur����ID, "ȱʡ���Ͳ���������ɫ����", 0) = "1"       'ȱʡ���Ͳ���������ɫ����
    
    If mSysPar.blnNameColColorCfg Then
        gstrSQL = "select ���� from �������� where ȱʡ��־=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡȱʡ��������")
        
        If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = Nvl(rsTemp!����)
    End If
    
    mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("�������Զ���ӡ���뵥", glngSys, mlngModule, 0)) '�������Զ���ӡ���뵥
    
    mSysPar.blnChangeUser = GetDeptPara(mlngCur����ID, "�������û�", 0) = "1"              '�������û�
    mSysPar.blnSwitchUser = GetDeptPara(mlngCur����ID, "�����л��û�", 0) = "1"              '�����л��û�
    
    mSysPar.blnIsPetitionScan = IIf(Val(GetDeptPara(mlngCur����ID, "�������뵥ɨ��", 1)) = 1, True, False)   '��ȡ�������뵥ɨ�����
    mSysPar.strImageLevel = Nvl(GetDeptPara(mlngCur����ID, "Ӱ�������ȼ�", "��,��"))
    mSysPar.strReportLevel = Nvl(GetDeptPara(mlngCur����ID, "���������ȼ�", "��,��"))
    mSysPar.blnֱ�Ӽ�� = (Val(GetDeptPara(mlngCur����ID, "�ǼǺ�ֱ�Ӽ��", 0)) = 1)         '�ǼǺ�ֱ�Ӽ��

    mSysPar.lngCriticalValues = Val(GetDeptPara(mlngCur����ID, "Σ������ж�", 0))           'Σ������ж�
    mSysPar.blnIgnoreResult = GetDeptPara(mlngCur����ID, "���Խ��������", 0) = "1" '        '���Խ��������
    mSysPar.lngConformDetermine = Val(GetDeptPara(mlngCur����ID, "��������ж�", 0))         '��������ж�
    mSysPar.lngImageLevel = Val(GetDeptPara(mlngCur����ID, "Ӱ�������ж�", 0))               'Ӱ�������ж�
    mSysPar.lngReportLevel = Val(GetDeptPara(mlngCur����ID, "���������ж�", 0))
    
    mSysPar.lngHintType = Val(GetDeptPara(mlngCur����ID, "��Ͻ����ʾ����", 0))
    
    mSysPar.blnFinishCommit = GetDeptPara(mlngCur����ID, "�ޱ�����ɺ�ֱ�����", 0) = "1" '  '�ޱ�����ɺ�ֱ�����
    mSysPar.blnReportWithImage = GetDeptPara(mlngCur����ID, "��ͼ�����д����", 0) = "1" '   '��ͼ�����д����
    mSysPar.blnReportWithResult = GetDeptPara(mlngCur����ID, "��Ӱ�����Ϊ����", 0) = "1" '  '��Ӱ�����Ϊ����
    mSysPar.blnLocalizerBackward = GetDeptPara(mlngCur����ID, "��λƬ����", 0) = "1" '       '��λƬ����
    mSysPar.blnCompleteCommit = GetDeptPara(mlngCur����ID, "��˺�ֱ�����", 0) = "1" '      '��˺�ֱ�����
    
    mSysPar.lngBeforeDays = Val(GetDeptPara(mlngCur����ID, "Ĭ�Ϲ�������", 2)) '                   'Ĭ�Ϲ�������
    If mSysPar.lngBeforeDays > 15 Or mSysPar.lngBeforeDays <= 0 Then
        mSysPar.lngBeforeDays = 2
    End If
    
    mSysPar.blnWriteCapDoctor = GetDeptPara(mlngCur����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", 0) = "1"  '�ɼ�ͼ����Ϊ��鼼ʦ
    
    mSysPar.blnPrintCommit = GetDeptPara(mlngCur����ID, "��ӡ��ֱ�����", 0) = "1" '           '��ӡ��ֱ�����
    mSysPar.blnCanPrint = GetDeptPara(mlngCur����ID, "ƽ������˲��ܴ򱨸�") = "1"             'ƽ����Ҫ��˲��ܴ�ӡ =true
    mSysPar.blnAutoSendWorkList = GetDeptPara(mlngCur����ID, "����ʱ�Զ�����WorkList") = "1"   '����ʱ�Զ�����WorkList

                
    '״̬����
    mSysPar.lngEnregAfterTimeLen = Val(GetDeptPara(mlngCur����ID, "�ǼǺ�����", 0))
    mSysPar.lngCheckInAfterTimeLen = Val(GetDeptPara(mlngCur����ID, "����������", 0))
    mSysPar.lngStudyAfterTimeLen = Val(GetDeptPara(mlngCur����ID, "��������", 0))
    mSysPar.lngReportAfterTimeLen = Val(GetDeptPara(mlngCur����ID, "���������", 0))
    mSysPar.lngAuditAfterTimeLen = Val(GetDeptPara(mlngCur����ID, "��˺�����", 0))
    
    '����ʱ��Ƭ
    mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModule, 0)) = 1)
    
    If InStr(mstrPrivs, "�Ŷӽк�") > 0 And mlngModule <> G_LNG_PATHOLSYS_NUM Then    '��Ȩ��ʹ�òŸ��ݲ�������
        mSysPar.blnUseQueue = GetDeptPara(mlngCur����ID, "�����Ŷӽк�", 0) = "1" '          'Ĭ�ϲ������Ŷӽк�
        
        If mSysPar.blnUseQueue Then
            mSysPar.lngQueueWay = GetDeptPara(mlngCur����ID, "�Ŷӽкŷ�ʽ", 0)             '�Ŷӽкŵ��Ŷӷ�ʽ
            mSysPar.blnSynStudylist = GetDeptPara(mlngCur����ID, "ͬ����λ����б�", 0)
            mSysPar.blnAutoInQueue = GetDeptPara(mlngCur����ID, "�������Զ��Ŷ�", 1)
        Else
            mSysPar.lngQueueWay = 0
        End If
    End If
    
    mSysPar.blnRelatingPatient = GetDeptPara(mlngCur����ID, "������������", 0) = "1"       '�Ƿ�ʹ�ù�
    mSysPar.lngRefreshInterval = Val(GetDeptPara(mlngCur����ID, "�Զ�ˢ�¼��", 0))  '     '�Զ�ˢ�¼��,Ĭ�ϲ��Զ�ˢ��
    
    gblnXWLog = (Val(zlDatabase.GetPara("XW��¼�ӿ���־", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '�Ƿ��¼�ӿ���־
    
    If mSysPar.lngRefreshInterval > 0 Then
        If mSysPar.lngRefreshInterval > 65 Then mSysPar.lngRefreshInterval = 65
        timerRefresh.Interval = mSysPar.lngRefreshInterval * 1000
        timerRefresh.Enabled = True
    Else
        timerRefresh.Enabled = False
    End If

    If blnIsUpdateSearchTime Then
        SQLCondition.��ʼʱ�� = CDate(Format(zlDatabase.Currentdate - (mSysPar.lngBeforeDays - 1), "yyyy-mm-dd 00:00"))
        
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
        
        If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
            gblnXWMoved = mblnMoved
        End If
    End If
        

    '��ʼ�����������б�
    If mSysPar.lngQueueWay = 0 Then
        '��ʼ�����������б�
        Dim iCount As Integer
        Dim strSql As String
        
        iCount = 1
        gstrSQL = "Select ִ�м�,����豸 From ҽ��ִ�з��� where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡִ�м�����", mlngCur����ID)
        If rsTemp.EOF <> True Then
            ReDim mAstr��������(rsTemp.RecordCount) As String
            While rsTemp.EOF = False
                mAstr��������(iCount) = mlngCur����ID & ":" & Nvl(rsTemp!ִ�м�)
                iCount = iCount + 1
                rsTemp.MoveNext
            Wend
    
    '       �����������ڲ���
    '        ReDim mAstr��������(8) As String
    '        mAstr��������(1) = "42:CT98"
    '        mAstr��������(2) = "42:CT99"
    '        mAstr��������(3) = "61:CT2"
    '        mAstr��������(4) = "61:CT1"
    '        mAstr��������(5) = "81:jy1"
    '        mAstr��������(6) = "81:jy2"
    '        mAstr��������(7) = "82:�����"
    '        mAstr��������(8) = "83:" & Nvl(rsTemp!ִ�м�)
            
        Else
            ReDim mAstr��������(0) As String
        End If
    Else
        ReDim mAstr��������(1) As String

        mAstr��������(1) = mstrCur����

    End If
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = PicList.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picWindow.hWnd
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo ErrHandle
    '��ֹ����б� �϶�
    Cancel = IIf(((Action = 4 Or Action = 6 Or Action = 5) And Not Pane.Hidden), True, False)
ErrHandle:
End Sub


Private Sub InitStudyList()
    Dim strCols As String
    Dim strDefaultCols As String
    Dim i As Integer
    Dim arrCol() As String
    Dim strTemp As String
    
    
    strCols = zlDatabase.GetPara("����б�", glngSys, mlngModule, "")
    
    Set ufgStudyList.ImageList = imgList
    
    Select Case mlngModule
         Case G_LNG_PACSSTATION_MODULE   'ҽ��
            strDefaultCols = Replace(IIf(mrtReportType = �����ĵ��༭��, M_STR_PUBLIC_COLS_NEW, M_STR_PUBLIC_COLS), "[------]", M_STR_IMAGES_COLS)
                
        Case G_LNG_PATHOLSYS_NUM        '����
            strDefaultCols = Replace(IIf(mrtReportType = �����ĵ��༭��, M_STR_PUBLIC_COLS_NEW, M_STR_PUBLIC_COLS), "[------]", M_STR_PATHOL_COLS)
            
        Case G_LNG_VIDEOSTATION_MODULE  '�ɼ�
            strDefaultCols = Replace(IIf(mrtReportType = �����ĵ��༭��, M_STR_PUBLIC_COLS_NEW, M_STR_PUBLIC_COLS), "[------]", M_STR_CAPTOR_COLS)
    End Select
    
    
    arrCol() = Split(strCols, "|")
    
    For i = 0 To UBound(arrCol())
        If arrCol(i) <> "" Then
            If InStr(arrCol(i), "���뵥") > 0 Then
                strTemp = arrCol(i)
                
                If mSysPar.blnIsPetitionScan Then
                    '���������뵥ɨ��ʱ�������뵥�������������
                    strCols = Replace(strCols, strTemp, Replace(strTemp, ",uncfg", ""))
                Else
                    '��δ�������뵥ʱ������������뵥�н�������
                    strCols = Replace(strCols, strTemp, Replace(Replace(strTemp, ",hide", ""), ",uncfg", "") & ",hide,uncfg")
                    
                    strDefaultCols = Replace(strDefaultCols, "���뵥>���뵥ҽ��,w1100", "���뵥>���뵥ҽ��,w1100,hide,uncfg")
                End If

                Exit For
            End If
        End If
    Next i
    
    
    ufgStudyList.DefaultColNames = strDefaultCols
    ufgStudyList.ColNames = IIf(strCols = "", strDefaultCols, strCols)
    
    ufgStudyList.IsKeepRows = False
    ufgStudyList.IsCopyMode = False
    ufgStudyList.IsAutoRowHeight = False
End Sub


Private Sub InitForm()
    Dim strKinds As String
    Dim blnDo As Boolean
    Dim lngKey As Long
    Dim bytFontSize As Byte
    
    Call WriteLog("InitForm -> Step 1����ʼִ��...")
    
    '�õ����Ի�������
    blnDo = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0
    
    mstrPrivs = gstrPrivs 'Ȩ��
    mlngModule = glngModul 'ģ���
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    mblnAllDepts = False
    mlngSortCol = 0
    mintSortOrder = 0
    mSysPar.lngQueueWay = 0
    
    '��ȡ�����С
    bytFontSize = Val(zlDatabase.GetPara("��ʾ�����С", glngSys, glngModul))
    mbytFontSize = IIf(bytFontSize = 0, 9, IIf(bytFontSize = 1, 12, 15))
    '��ʼ����״̬
    mbyrFontState = 2
    
    
    
    mblnInitOk = False  '��ʼ����,��ʼ�����֮ǰ���������ݵ���ȡ
    mblnvsRefresh = False
    mblnMenuDownState = False
    mlngFilterTab = 0
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then labHistory.Caption = "������ʷ��"
    
    
    '�жϵ�ǰ�û��Ƿ���� ��Ƭվ�Ļ���Ȩ��
    mblnObserve = IIf(InStr(GetPrivFunc(glngSys, 1289), "����") > 0, True, False)
    
    If mlngModule = G_LNG_PATHSTATION_MODULE Then
        mlngFilterTab = Val(zlDatabase.GetPara("����ҳ��", glngSys, glngModul))
        
        tabFilter.Visible = True
        picExeState.Visible = True
        
        Call InitFilterPage
    End If
    
    Call WriteLog("InitForm -> Step 2�����뱾��ע������...")
    
    '�жϵ�ǰ�û��Ƿ���С�Ӱ���豸Ŀ¼����Ȩ�ޣ��д�Ȩ�޲ſ�������������PACS����
    mblnSetXWParam = IIf(InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "PACS��������") > 0, True, False)
    
    Call InitLocalPars '����ע������
    
    Call WriteLog("InitForm -> Step 3�����벿�������Ϣ...")
    If Not InitDepts Then Unload Me: Exit Sub '��ʼ��ҽ������
    
    mrtReportType = GetDeptPara(mlngCur����ID, "����༭��", 0)                 '����༭��
    
    ReDim gConnectedShardDir(0) As String   '��ʼ������Ŀ¼���Ӵ�
    
    Call WriteLog("InitForm -> Step 4����ʼ�����ż�����...")
    Call InitModuleParameter '��ʼ��ģ�鼶����
    
    
    '��ʼ�Ӵ���
    Set mobjEvent = New clsEvent
    Set gobjEvent = mobjEvent
    
    
    '���ݲ����ж��Ƿ�������Ϣ����
    Set mobjMsgCenter = New clsPacsMsgProcess
    Call mobjMsgCenter.OpenMsgCenter(mlngModule, mlngCur����ID, mstrPrivs)
    
    Set mobjPacsCore = New zl9PacsCore.clsViewer
    
    
    If mSysPar.blnUseQueue And InStr(GetPrivFunc(glngSys, 1160), "����") > 0 Then
        Set mobjQueue = New frmWork_Queue
        Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, NeedName(mstrCur����), mstrPrivs)
    Else
        Set mobjQueue = Nothing
    End If

    Call WriteLog("InitForm -> Step 5����ȡ�б���ɫ����...")
    Call ReadStudyListColor(mlngCur����ID)
    
    Call WriteLog("InitForm -> Step 6����ȡ���ٹ�������...")
    Call InitFilterCmd
    
    Call WriteLog("InitForm -> Step 7����ʼ�����ڲ˵�...")
    Call InitCommandBars
'    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call WriteLog("InitForm -> Step 8����ʼ�����沼��...")
    Call InitFaceScheme
    
    Call WriteLog("InitForm -> Step 9����ʼ����������б�...")
    Call InitStudyList
    
     '���ע����й��������ֵΪ�� ���� �ѹ�ѡ���Ի����ã���ô��ע���д�빤������ʾģʽֵ
    If mintToolBarWriteReg = 9 Or (mintToolBarWriteReg = 0 And blnDo) Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 3
    End If
    
    '�ָ������״̬   ע���ָ�����״̬ ������� ��ע���д�빤������ʾģʽֵ �������棬�������ɹ�������ʾģʽ����
    Call RestoreWinState(Me, App.ProductName)
    
    picAppend.Height = Nvl(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "StudyInfHeight", picAppend.Height))
    
     '������--- �ı���ǩ ������ʹ��RestoreWinState �ָ����ˣ�����Ҫ����������δ��ѡ���Ի����ã��򹤾���Ĭ����ʾͼ����ı�
    If blnDo Then
        If Me.cbrMain(2).Controls(1).Style = xtpButtonIconAndCaption Then
            Me.cbrMain(2).ShowTextBelowIcons = True
        Else
            Me.cbrMain(2).ShowTextBelowIcons = False
        End If
    Else
        Me.cbrMain(2).ShowTextBelowIcons = True
    End If
    
    ClearCacheFolder App.Path & "\TmpImage\"    '����ʱĿ¼���ˣ�����ո�Ŀ¼
    
    
    '�ж���ʱĿ¼�Ƿ����
    If Dir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage", vbDirectory) = "" Then
        Call MkDir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage")
    End If
    
    
    '��ʼ��˫�û���½�Ĳ���
    mblnCnOracleIsHIS = True
    mintChangeUserState = 1
    mstrUserNameHIS = UserInfo.����
    mstrUserNameNew = UserInfo.����
    mstrUserIDHIS = UserInfo.�û���
    mstrUserIDNew = UserInfo.�û���
    
    Set mcnOracleHIS = gcnOracle
    
    Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameHIS & "   ���ҽ����" & mstrUserNameNew
    
    ReDim mobjPacsReportArry(0) As frmReport
    
    gblnUseXinWangView = False
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        gblnUseXinWangView = IsUseXwViewer
    
        '�����RIS����վ���������������ݿ⣬��ȡ����
        If gblnUseXinWangView Then
            '���Ͻػ���Ϣ��hook
            plngXWPreWndProc = XWHook(Me.hWnd)
        End If
    End If
    
    mblnFormLoadState = True
    
    Call WriteLog("InitForm -> Step 10������ִ��...")
End Sub


'Private Sub Form_Load()
'On Error GoTo errHandle
'    '��ʼ����ط�����showstation�е���InitForm���д���......
'    '���ﲻ�ܽ�����صĳ�ʼ����������Ϊ��clsPacsWork��BHCodeMain�����У�������ʾ��ʽ��ʱ�򣬻ᴥ��Load�¼���
'    '��Load�¼��е�ĳЩ������Ҫ��ز���������ȷִ�У������Ҫ��Load�еĴ�����������ȡ����������ShowStation������ִ��...
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub


Private Sub RefreshStatusBarInf()
    Dim i As Long
    
    Dim lngDengJi As Long
    Dim lngBaoDao As Long
    Dim lngJianCha As Long
    Dim lngBaoGao As Long
    Dim lngShenHe As Long
    Dim lngBoHui As Long
    Dim lngWanCheng As Long
    Dim lngYiBaoGao As Long
    Dim strTemp As String
    
    
    lngDengJi = 0
    lngBaoDao = 0
    lngJianCha = 0
    lngBaoGao = 0
    lngShenHe = 0
    lngBoHui = 0
    lngWanCheng = 0
    lngYiBaoGao = 0
    
    
    For i = 1 To ufgStudyList.GridRows - 1
        Select Case ufgStudyList.Text(i, "������")
            Case "�ѵǼ�"
                lngDengJi = lngDengJi + 1
            Case "�ѱ���"
                lngBaoDao = lngBaoDao + 1
            Case "�Ѽ��"
                lngJianCha = lngJianCha + 1
            Case "�ѱ���"
                lngYiBaoGao = lngYiBaoGao + 1
            Case "������"
                lngBaoGao = lngBaoGao + 1
            Case "�����"
                lngShenHe = lngShenHe + 1
            Case "�Ѳ���"
                lngBoHui = lngBoHui + 1
            Case "�����"
                lngWanCheng = lngWanCheng + 1
        End Select
    Next i
    
    strTemp = ""
    If lngDengJi > 0 Then strTemp = "�ѵǼǣ�" & lngDengJi & "    "
    If lngBaoDao > 0 Then strTemp = strTemp & "�ѱ�����" & lngBaoDao & "    "
    If lngJianCha > 0 Then strTemp = strTemp & "�Ѽ�飺" & lngJianCha & "    "
    If lngBaoGao > 0 Then strTemp = strTemp & "�����У�" & lngBaoGao & "    "
    If lngYiBaoGao > 0 Then strTemp = strTemp & "�ѱ��棺" & lngYiBaoGao & "    "
    If lngShenHe > 0 Then strTemp = strTemp & "����ˣ�" & lngShenHe & "    "
    If lngBoHui > 0 Then strTemp = strTemp & "�Ѳ��أ�" & lngBoHui & "    "
    If lngWanCheng > 0 Then strTemp = strTemp & "����ɣ�" & lngWanCheng & "    "
    
    stbThis.Panels(2).Text = "�� " & ufgStudyList.GridRows - 1 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
    stbThis.Panels(3).Text = strTemp
End Sub


Private Sub InitFilterPage()
    Dim lngHideCount As Long
    
    lngHideCount = 0
    With tabFilter
        .RemoveAll
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
        



        .InsertItem 0, "ȡ  ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "ȡ��"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����ȡ��")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 1, "��  Ƭ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "��Ƭ"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "������Ƭ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 2, "��  ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "����"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "�����黯")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 3, "��  ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "����"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���Ӳ���")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1


        .InsertItem 4, "��  Ⱦ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "��Ⱦ"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "����Ⱦɫ")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 5, "��  ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "����"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "���ﷴ��")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 6, "��  ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "����"
        
    End With

    '�����й��ܱ�ǩ������ʱ����ֱ������tabFilter�ؼ�
    tabFilter.Visible = (lngHideCount < tabFilter.ItemCount - 1)
    tabFilter.tag = (lngHideCount < tabFilter.ItemCount - 1)
    
    On Error GoTo errContinue1
    If tabFilter.tag Then
        If Not tabFilter.Item(mlngFilterTab).Visible Then
            tabFilter.Item(tabFilter.ItemCount - 1).Selected = True
        Else
            tabFilter.Item(mlngFilterTab).Selected = True
        End If
    End If
    
    optAccept.Enabled = IIf(tabFilter.Selected.tag = "ȡ��" Or tabFilter.Selected.tag = "����" Or tabFilter.Selected.tag = "����", False, True)
    
    optNeed.Enabled = IIf(tabFilter.Selected.tag = "����", False, True)
    optFinal.Enabled = IIf(tabFilter.Selected.tag = "����", False, True)
    optAll.Enabled = IIf(tabFilter.Selected.tag = "����", False, True)
errContinue1:
End Sub


Private Function GetWindowCaption() As String
    GetWindowCaption = Mid(Me.Caption & " ", 1, InStr(Me.Caption & " ", " "))
End Function


Private Sub DisposeObj()
    If Not mfrmWork_PacsImg Is Nothing Then
        Unload mfrmWork_PacsImg
        Set mfrmWork_PacsImg = Nothing
    End If
    
    If Not mobjQueue Is Nothing Then
        Unload mobjQueue
        Set mobjQueue = Nothing
    End If
    
    If Not mobjPacsCore Is Nothing Then
        mobjPacsCore.Closefrom
        Set mobjPacsCore = Nothing
    End If
    
    If Not mfrmPACSFilter Is Nothing Then
        Unload mfrmPACSFilter
        Set mfrmPACSFilter = Nothing
    End If
    
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.Free
        Set mobjWork_Pathol = Nothing
    End If
    
    If Not mobjWork_His Is Nothing Then
        Call mobjWork_His.Free
        Set mobjWork_His = Nothing
    End If
    
    If Not mobjWork_Report Is Nothing Then
        Call mobjWork_Report.Free
        Set mobjWork_Report = Nothing
    End If
    
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        If Not mobjCaptureHot Is Nothing Then
            Call mobjCaptureHot.FreeHook
            Set mobjCaptureHot = Nothing
        End If
    End If
    
    If mblnUseActivexCapture Then
        'ʹ��Activex����Ƶ�ɼ���ʽ�˳�
        Set mobjWork_ActiveVideo = Nothing
    End If
    
    If Not gobjMsgCenter Is Nothing Then
        Set gobjMsgCenter = Nothing
    End If

        
    Set mobjEvent = Nothing
    Set mobjSquareCard = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    If mblnUseActivexCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlNotifyQuit
        End If
    End If
    
    '�ر���Ϣ����
    If Not gobjMsgCenter Is Nothing Then
        Call gobjMsgCenter.CloseMsgCenter
    End If
    
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mlngSortCol)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mintSortOrder)
    
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "StudyInfHeight", picAppend.Height)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "StudyListWidth", PicList.Width / Me.ScaleWidth)
        
    '�����������
    zlDatabase.SetPara "��ʾ�����С", IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, IIf(mbytFontSize = 15, 2, mbytFontSize))), glngSys, glngModul
    
    '�ָ���������
    Me.Caption = GetWindowCaption
    
    Call SaveWinState(Me, App.ProductName)
    
    Call DisposeObj
    
    '�ָ�����̨�����ݿ�����
    If mblnCnOracleIsHIS = False Then
        Set gcnOracle = mcnOracleHIS
        InitCommon gcnOracle
        RegCheck
        SetDbUser mstrUserIDHIS
        Call GetUserInfo
        Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
    End If
    
    frmTwoUser.intDBState = 1
    
    '�����RIS����վ����Ͽ����������ݿ������
    If gblnUseXinWangView Then
        '    ж��hook
        XWUnhook Me.hWnd, plngXWPreWndProc
    End If
    
    mblnFormLoadState = False
    
    Exit Sub
ErrHandle:
    Debug.Print err.Description
End Sub

Private Function InitCardType(ByVal strCardNames As String) As String
'��ָ����ʽ��ʼ��������
    Dim i As Integer
    Dim aryKindInfo() As String
    Dim strKinds As String
    
    aryKindInfo = Split(strCardNames, ";")
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo) - 1
        If strKinds <> "" Then strKinds = strKinds & ";"
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    InitCardType = strKinds & ";"
End Function

Private Sub InitLocalPars()
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
'��ʼ����ʱ���ز������Ը�������Ϊ��,������أ����ˣ��������õȵ���

    mstrCaptureHot = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
    mstrCaptureAfterHot = GetSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", "F7")
    mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", "F6")
    
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmdסԺ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", 1))
    mblncmd�ѽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", 0))
    mblncmdδ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", 0))
    mblncmd�޷� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����޷�", 0))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ü���", 0))
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ò���", 0))
        
        mblnPopChangGuiWindow = (Val(zlDatabase.GetPara("������������", glngSys, mlngModule, 0)) = 1)
        mblnPopKuaiShuWindow = (Val(zlDatabase.GetPara("����ʯ����������", glngSys, mlngModule, 0)) = 1)
        mblnPopBingDongWindow = (Val(zlDatabase.GetPara("������������", glngSys, mlngModule, 0)) = 1)
        mblnPopXiBaoWindow = (Val(zlDatabase.GetPara("ϸ����������", glngSys, mlngModule, 0)) = 1)
        mblnPopHuiZhenWindow = (Val(zlDatabase.GetPara("������������", glngSys, mlngModule, 0)) = 1)
        mblnPopShiJianWindow = (Val(zlDatabase.GetPara("ʬ����������", glngSys, mlngModule, 0)) = 1)
    End If
    
    mblncmd�Ǽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鲡��", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ز���", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", 1))
    
    mlngLocateFindType = zlDatabase.GetPara("��λ���ҷ�ʽ", glngSys, mlngModule, 0)
        
    mstrFindWay = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���˷�ʽ", GetStudyNumberDisplayName)
    mstrLocateWay = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", GetStudyNumberDisplayName)
    
    PatiIdentify.IDKindStr = InitCardType(Replace(IIf(mlngLocateFindType = TLocateFindType.lftLocate, CONST_STR_LOCAL_CARD_TYPE, CONST_STR_FIND_CARD_TYPE), "[------]", GetStudyNumberDisplayName))
    PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(IIf(mlngLocateFindType = 0, mstrLocateWay, mstrFindWay))
    
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", "0"))
    mlngSortCol = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", 0))
    
    strTemp = zlDatabase.GetPara("Ӱ��������", glngSys, mlngModule, "")
    
    ReDim strTempArry(0)
    ReDim mblncmdӰ�����(0)
    
    On Error GoTo errContinue1
    strTempArry = Split(strTemp, ",")
    If UBound(strTempArry) >= 0 Then ReDim mblncmdӰ�����(UBound(strTempArry))
    
    For i = 0 To UBound(strTempArry)
        mblncmdӰ�����(i) = IIf(UCase(strTempArry(i)) = "TRUE", True, False)
    Next i
        
    ReDim mblncmdӰ��ִ�м�(0)
errContinue1:
    mSysPar.strFirstTab = zlDatabase.GetPara("������ҳ", glngSys, mlngModule, "") 'Ϊ�ձ�ʾ��ʹ�ö��ƹ�����ҳ����
    mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnNoShowCancel = (Val(zlDatabase.GetPara("����ʾ��ȡ���ĵǼ�", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnPatTrack = (Val(zlDatabase.GetPara("���˸���", glngSys, mlngModule, 0)) = 1)
    mSysPar.strLocalRoom = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModule, "")
    
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        '����ǲɼ�ģ�飬����Ҫִ�иò���
        mSysPar.lngMoneyExeModle = Val(zlDatabase.GetPara("�ɼ�����ִ��ģʽ", glngSys, mlngModule, 0))
    End If
    
    '�õ�ע����й��ڹ�������ʾ״̬��ֵ�����Ϊ�������9
    mintToolBarWriteReg = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 9))
    
    
    With SQLCondition '------------------------ '����������ʼ
        'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
       .ʱ������ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", 1))
       .���ݺ� = ""
       .����� = 0
       .סԺ�� = 0
       .������ = ""
       .���￨ = ""
       .���� = ""
       .�Ա� = ""
       .��ʼ���� = -1
       .�������� = -1
       .�������� = "="
       .���� = 0
       .���֤ = ""
       .IC�� = ""
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
       .Ӱ����� = ""
       .������� = ""
       .������ = ""
       .���� = ""
       .��� = ""
    End With
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
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
   

    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, CStr("," & str��Դ & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            mstrCanUse����IDs = mstrCanUse����IDs & "," & rsTmp!ID
            
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        mstrCanUse����IDs = Mid(mstrCanUse����IDs, 2)
        
        If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ�ô˹���վ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        InitDepts = True
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur����ID
    End If
End Function

Private Sub InitFaceScheme()
    Dim lngListWidth As Double
    
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    lngListWidth = Nvl(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "StudyListWidth", 0.35))
    If lngListWidth >= 1 Then lngListWidth = 0.35
    
    'ע����б���Ľ��沼��Pnae�������ԣ������Ĭ�ϵ�Pane����
    If dkpMain.PanesCount <> 3 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, lngListWidth * 100, 250, DockLeftOf, Nothing)
        Pane1.Title = "����б�"
        Pane1.Handle = PicList.hWnd
        Pane1.Options = PaneNoCloseable Or PaneNoFloatable
        
        Set Pane2 = dkpMain.CreatePane(2, (1 - lngListWidth) * 100, 300, DockRightOf, Nothing)
        Pane2.Title = "�Ӵ���"
        Pane2.Handle = picWindow.hWnd
        Pane2.Options = PaneNoCaption Or PaneNoCloseable
    End If
End Sub

'�����ٹ����������仯ʱ����
Private Sub SaveFilterCmd()
    Dim strTemp As String
    Dim i As Integer
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", IIf(mblncmdסԺ, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", IIf(mblncmd�ѽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", IIf(mblncmdδ��, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ü���", IIf(mblncmd����, 1, 0)
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ò���", IIf(mblncmd����, 1, 0)
    End If
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����޷�", IIf(mblncmd�޷�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", IIf(mblncmd�Ǽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鲡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ز���", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���˷�ʽ", mstrFindWay
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", mstrLocateWay
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", IIf(mblncmd����, 1, 0)
    
    If mlngModule = G_LNG_PATHSTATION_MODULE Then
        '����ģ�鵥������Ĳ���
        Call zlDatabase.SetPara("�������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("����ʯ������", IIf(mblncmd����ʯ��, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("��������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("ϸ������", IIf(mblncmdϸ��, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("�������", IIf(mblncmd����, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("ʬ�����", IIf(mblncmdʬ��, 1, 0), glngSys, glngModul)
        
        Call zlDatabase.SetPara("����ҳ��", tabFilter.Selected.Index, glngSys, glngModul)
    End If
    
    If UBound(mblncmdӰ�����) >= 0 Then
        strTemp = mblncmdӰ�����(0)
    End If
    For i = 1 To UBound(mblncmdӰ�����)
        strTemp = strTemp & "," & mblncmdӰ�����(i)
    Next i
    Call zlDatabase.SetPara("Ӱ��������", strTemp, glngSys, mlngModule)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        If UBound(mblncmdӰ��ִ�м�) >= 0 Then
            strTemp = mlngCur����ID & ":" & IIf(mblncmdӰ��ִ�м�(0), "T", "F")
        End If
        
        For i = 1 To UBound(mblncmdӰ��ִ�м�)
            strTemp = strTemp & "," & IIf(mblncmdӰ��ִ�м�(i), "T", "F")
        Next i
        
        '�滻��Ӧ���ҵ�ִ�м�����
        If mstrAllExamineRoomCfg = "" Or InStr(mstrAllExamineRoomCfg, ":") <= 0 Then
            mstrAllExamineRoomCfg = strTemp
        Else
            If InStr(mstrAllExamineRoomCfg, mlngCur����ID & ":" & mstrCurExamineRoomCfg) > 0 Then
                mstrAllExamineRoomCfg = Replace(mstrAllExamineRoomCfg, mlngCur����ID & ":" & mstrCurExamineRoomCfg, strTemp)
            Else
                mstrAllExamineRoomCfg = mstrAllExamineRoomCfg & "|" & strTemp
            End If
        End If
        
        '���óɵ�ǰ���õ�ֵ
        If strTemp <> "" Then mstrCurExamineRoomCfg = Split(strTemp, ":")(1)
        
        Call zlDatabase.SetPara("Ӱ��ִ�м����", mstrAllExamineRoomCfg, glngSys, mlngModule)
    End If
    
    '�����鲿λ
    Call zlDatabase.SetPara("��鲿λ����", mstrcmd��λ����, glngSys, mlngModule)
End Sub

Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    Dim i As Integer
    Dim strStudyTypes As String

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
        '��Դ.........................................................
        Set objControl = .Add(xtpControlButtonPopup, ID_��Դ, "��Դ")
        objControl.ToolTipText = "���ݲ�����Դ���й���"
        
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&1)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_סԺ, "סԺ(&2)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&3)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_���, "���(&4)")
        
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
            
            
        '״̬.........................................................
        Set objControl = .Add(xtpControlButtonPopup, ID_״̬, "״̬")
        objControl.ToolTipText = "���ݼ��״̬���й���"
        
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�Ǽ�, "�Ǽ�(&1)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&2)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_���, "���(&3)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&4)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_���, "���(&5)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&6)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_���, "���(&7)")
    
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
        
            
        If mlngModule = G_LNG_PATHSTATION_MODULE Then
            'ֻ�в���ϵͳ���в������
            Set objControl = .Add(xtpControlButtonPopup, ID_�������, "���")
            objControl.ToolTipText = "���ݲ��������й���"
            
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�������_����, "����(&1)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�������_����, "����(&2)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�������_ϸ��, "ϸ��(&3)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�������_ʬ��, "ʬ��(&4)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�������_����, "����(&5)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�������_����ʯ��, "����ʯ��(&6)")
        Else
            '�������Ӱ�����
            Set objControl = .Add(xtpControlButtonPopup, ID_Ӱ�����, "���   ")
            objControl.ToolTipText = "����Ӱ�������й���"
            
            strSql = "select ����,���� from Ӱ�������"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ�������")
            
            i = 1
            mintcmdӰ����� = 0
            strTemp = ""
            If rsTemp.RecordCount > 0 Then
                ReDim Preserve mblncmdӰ�����(rsTemp.RecordCount - 1)
                
                While rsTemp.EOF = False
                    Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_Ӱ����� + i, rsTemp("����") & "(&" & Chr(64 + i) & ")")
                    
                    cbrPopControl.DescriptionText = rsTemp("����")
                    cbrPopControl.Style = xtpButtonIconAndCaption
                    cbrPopControl.Checked = mblncmdӰ�����(i - 1)
                    cbrPopControl.CloseSubMenuOnClick = False
                    
                    If mblncmdӰ�����(i - 1) = True Then
                        strStudyTypes = strStudyTypes & "," & rsTemp("����")
                        
                        mintcmdӰ����� = mintcmdӰ����� + 1
                        strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
                    End If
                    
                    rsTemp.MoveNext
                    i = i + 1
                Wend
                
                If strStudyTypes <> "" Then strStudyTypes = Mid(strStudyTypes, 2)
                
                If strTemp <> "" Then
                    objControl.ToolTipText = "��ʾӰ�����Ϊ[" & strTemp & "]�ļ��"
                    objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
                End If
            End If
        End If
        
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
        
        '����.........................................................
        Set objControl = .Add(xtpControlButtonPopup, ID_����, " ����")
            objControl.ToolTipText = "���ݷ���״̬���й���"
            
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_δ��, "δ��(&1)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�ѽ�, "�ѽ�(&2)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&3)")
        
        If mlngModule = G_LNG_PATHOLSYS_NUM Then
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����(&4)")
        End If
        
        '���û�в��ɲ˵�����ʹ������3�İ�����Ϊ��ݼ�
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�޷�, "�޷�(&" & IIf(mlngModule = G_LNG_PATHOLSYS_NUM, 5, 4) & ")")
        
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
        
        '�������Ӱ��ִ�м�
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set objControl = .Add(xtpControlButtonPopup, ID_Ӱ��ִ�м�, "ִ�м�   ")
            objControl.ToolTipText = "����Ӱ��ִ�м���й���"
            
            Call InitExamineRoom(objControl, cbrPopControl, mlngCur����ID)
        End If
        
        '������м�鲿λ
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set objControl = .Add(xtpControlButtonPopup, ID_��鲿λ, "��λ   ")
            objControl.ToolTipText = "���ݼ�鲿λ���й���"
            
            Call InitStudyPlace(objControl, cbrPopControl, strStudyTypes)
        End If
    End With
    
    For Each objControl In objBar.Controls
        If objControl.type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("����", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_���ҷ�ʽ, "")
        objControl.Style = xtpButtonIcon
        objControl.IconId = IIf(mlngLocateFindType = TLocateFindType.lftLocate, 3, 4)
        
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_����ֵ, "����ֵ")
        objCusControl.Handle = PatiIdentify.hWnd
        objCusControl.flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_��ʼ����, IIf(mlngLocateFindType = TLocateFindType.lftLocate, "��ʼ��λ", "��ʼ����"))
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_����סԺ, "����")
    objControl.ToolTipText = "ֻ��ʾ����סԺ����¼"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, Asc("G"), ID_��ʼ����
    End With
    cbrdock.RecalcLayout
End Sub

Private Sub InitExamineRoom(objControl As CommandBarControl, cbrPopControl As CommandBarControl, ByVal lngCur����ID As Long)
'��ʼ��ִ�м��������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    Dim strTemp As String
    Dim strTempArry() As String
    
    Dim i As Integer
    Dim strID As String
    
    '��ȡִ�м�����,��ʽ:����1ID:ִ�м�1ѡ�����,ִ�м�2ѡ�����,...|����2ID:ִ�м�1ѡ�����,ִ�м�2ѡ�����,...|...
    'ʾ����64:T,F,T,F|65:T,F,T,F|...
    mstrAllExamineRoomCfg = zlDatabase.GetPara("Ӱ��ִ�м����", glngSys, mlngModule, "")
    
    For i = 0 To UBound(Split(mstrAllExamineRoomCfg, "|"))
        If Val(Split(Split(mstrAllExamineRoomCfg, "|")(i), ":")(0)) = mlngCur����ID Then
            mstrCurExamineRoomCfg = Split(Split(mstrAllExamineRoomCfg, "|")(i), ":")(1)
            strTemp = mstrCurExamineRoomCfg
            Exit For
        End If
    Next
    
    If mblnAllDepts Then
        If InStr(mstrPrivs, "���п���") > 0 Then
            strSql = "select ����,ִ�м� from ҽ��ִ�з��� a, ���ű� b where a.����Id=b.Id "
        Else
            '��ѯ��Ӧ��Ա���ڿ�������������ִ�м�
            strSql = "select ����,ִ�м� from ҽ��ִ�з��� a, ������Ա b,���ű� c where a.����id=b.����id and a.����Id=c.Id and b.��Աid = [1]"
            
            strID = UserInfo.ID
        End If
                
    Else
        strSql = "Select ����,ִ�м� From ҽ��ִ�з��� a, ���ű� b Where a.����Id=b.Id and  ����ID=[1]"
        
        strID = lngCur����ID
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strID)
        
    mintcmdӰ��ִ�м� = 0
    mstrSelQueueRooms = ""
    mstrAllQueueRooms = ""
    
    If rsData.RecordCount <= 0 Then
        objControl.Caption = "ִ�м�    "
        objControl.Enabled = False
        
        Exit Sub
    End If
    
    If rsData.RecordCount - 1 > UBound(Split(strTemp, ",")) Then strTemp = strTemp & String(rsData.RecordCount - 1 - UBound(Split(strTemp, ",")), ",")
    strTempArry = Split(strTemp, ",")
    
    i = 1
    strTemp = ""
    
    objControl.Enabled = True
    ReDim Preserve mblncmdӰ��ִ�м�(rsData.RecordCount - 1)

    While rsData.EOF = False
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_Ӱ��ִ�м� + i, Nvl(rsData("ִ�м�")) & "(&" & Chr(64 + i) & ")")
    
        cbrPopControl.ToolTipText = Nvl(rsData!����) & "-" & Nvl(rsData!ִ�м�)
        cbrPopControl.DescriptionText = Nvl(rsData!ִ�м�) 'Nvl(rsData!����) & "-" & Nvl(rsData!ִ�м�)
        
        cbrPopControl.Style = xtpButtonIconAndCaption
        cbrPopControl.Checked = False
        cbrPopControl.CloseSubMenuOnClick = False
    
        '�ŶӽкŶ��е�����Ϊ���������ơ� + ��-�� + ��ִ�м����ơ�
        If mstrAllQueueRooms <> "" Then mstrAllQueueRooms = mstrAllQueueRooms & ","
        mstrAllQueueRooms = mstrAllQueueRooms & Nvl(rsData!����) & "-" & Nvl(rsData!ִ�м�)
                
        If UCase(strTempArry(i - 1)) = UCase("T") Then
            mintcmdӰ��ִ�м� = mintcmdӰ��ִ�м� + 1
            mblncmdӰ��ִ�м�(i - 1) = True
            cbrPopControl.Checked = True
            
            strTemp = IIf(strTemp = "", Mid(cbrPopControl.Caption, 1, InStr(cbrPopControl.Caption, "(") - 1), strTemp & "," & Mid(cbrPopControl.Caption, 1, InStr(cbrPopControl.Caption, "(") - 1))
            
            If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
            mstrSelQueueRooms = mstrSelQueueRooms & Nvl(rsData!����) & "-" & Nvl(rsData!ִ�м�)
        Else
            mblncmdӰ��ִ�м�(i - 1) = False
        End If
        
        rsData.MoveNext
        i = i + 1
    Wend
    
    '���û��ѡ���κ�ִ�м䣬��Ĭ��Ϊѡ��������ִ�м�
    If Trim(mstrSelQueueRooms) = "" Then mstrSelQueueRooms = mstrAllQueueRooms
        
    If strTemp <> "" Then
        objControl.ToolTipText = "��ʾӰ��ִ�м�Ϊ[" & strTemp & "]�ļ��"
        objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
    Else
        objControl.Caption = "ִ�м�    "
    End If
End Sub

Private Sub InitStudyPlace(objControl As CommandBarControl, cbrPopControl As CommandBarControl, ByVal strStudyTypes As String)
'��ʼ����鲿λ����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsGroup As ADODB.Recordset
    
    Dim strTemp As String
    Dim objTmpControl As CommandBarControl
    
    Dim i As Integer, j As Integer
    
    objControl.CommandBar.Controls.DeleteAll
    
    mstrcmd��λ���� = zlDatabase.GetPara("��鲿λ����", glngSys, mlngModule, "")
    
    If strStudyTypes = "" Then
        strSql = "Select Distinct ����, substr(����,instr(����,'-')+1) as ����,���� From ���Ƽ�鲿λ Order By ����,����"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        strSql = "Select Distinct ����, substr(����,instr(����,'-')+1) as ����,���� " & _
                 "From ���Ƽ�鲿λ A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
                 "Where A.���� =B.Column_Value Order By ����,����"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strStudyTypes)
    End If
    
    If rsData.RecordCount <= 0 Then
        objControl.Caption = "��λ    "
        objControl.Enabled = False
        mstrcmd��λ���� = ""
        mstrcmd��λ = ""
        
        Exit Sub
    End If
    
    i = 1
    objControl.Enabled = True
    
    Dim str��λ���� As String
    Dim str��λ As String
    Dim blnIsExist As Boolean
    
    While rsData.EOF = False
        blnIsExist = False
        
        For j = 1 To objControl.CommandBar.Controls.Count
            Set objTmpControl = objControl.CommandBar.Controls(j)
            
            If Not objTmpControl Is Nothing Then
                If objTmpControl.Caption = Nvl(rsData!����) Then
                    If InStr(objTmpControl.Category, Nvl(rsData!����)) <= 0 Then
                        objTmpControl.Category = objTmpControl.Category & "|" & Nvl(rsData!����)
                    End If
                    
                    blnIsExist = True
                    Exit For
                End If
            End If
        Next
        
        If blnIsExist = False Then
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_��鲿λ + i, Nvl(rsData!����))
        
            cbrPopControl.ToolTipText = Nvl(rsData!����)
            
            If InStr(cbrPopControl.Category, Nvl(rsData!����)) <= 0 Then
                cbrPopControl.Category = cbrPopControl.Category & "|" & Nvl(rsData!����)
            End If
            
            cbrPopControl.DescriptionText = Nvl(rsData!����)
            
            cbrPopControl.Style = xtpButtonIconAndCaption
            cbrPopControl.Checked = False
            cbrPopControl.CloseSubMenuOnClick = False
                    
            If InStr(mstrcmd��λ����, Nvl(rsData!����)) > 0 Then
                str��λ���� = str��λ���� & "," & Nvl(rsData!����)
                
                cbrPopControl.Checked = True
                
                strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
            End If
            
            i = i + 1
        End If
        
        rsData.MoveNext
    Wend
    
    For i = 1 To objControl.CommandBar.Controls.Count
        Set objTmpControl = objControl.CommandBar.Controls(i)

        If Not objTmpControl Is Nothing Then
            If objTmpControl.Checked = True Then
                 str��λ = str��λ & "," & objTmpControl.Category
            End If
        End If
    Next i
    
    mstrcmd��λ���� = str��λ����
    mstrcmd��λ = str��λ
    
    If strTemp <> "" Then
        objControl.ToolTipText = "��ʾ��鲿λΪ[" & strTemp & "]�ļ��"
        objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
    Else
        objControl.Caption = "��λ    "
    End If
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim blnShowCaption As Boolean
    
    Dim rsCollection As ADODB.Recordset
    Dim rsViewShare As ADODB.Recordset
    Dim rsShareCount As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer
    Dim i3DFunc As Integer
    Dim intTxtLen As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
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
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_FilePopup, "�ļ�", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_PrintSet, "��ӡ����", "", 181, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Excel, "�嵥��ӡ", "", 103, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Parameter, "��������", "", 181, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, ConMenu_File_ShortcutSet, "��ݼ�����", "", 181, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_WorkModule, "վ��ģʽ����", "", 9004, False)
        
        If mblnSetXWParam = True And mlngModule = G_LNG_PACSSTATION_MODULE Then    '�С�Ӱ���豸Ŀ¼����Ȩ�ޣ���������������PACS�Ĳ���
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SetXWParam, "PACS��������", "", 9004, False)
        End If
        
        '������Ƶ�ɼ����ò˵�
        If mlngModule <> G_LNG_PACSSTATION_MODULE And mblnUseActivexCapture = True Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "��Ƶ����", "��Ƶ����", 815, False)
        End If
        
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            '�����û������˵�
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "�û�����", "�������ҽ���ͱ���ҽ��", 3012, True)
        End If
        
        If mlngModule = G_LNG_PATHSTATION_MODULE Then
            '�����û��л��˵�
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "�л��û�", "�л��û�", 3012, True)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_SendImg, "����ͼ��", "", 3061, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Change_In, "�����б�", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Exit, "�˳�", "", 191, True)
    End With


'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ManagePopup, "���", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��", "", 0, False)
        
        '����������뵥ɨ����� ��ѡ������ء�ɨ�����뵥���˵���δ��ѡ�� ������
        If mSysPar.blnIsPetitionScan Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, comMenu_Petition_Capture, "ɨ�����뵥", "", 3935, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, comMenu_Petition_View, "�鿴���뵥", "�鿴��ɨ������뵥ͼ��", 3935, True)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Regist, "���Ǽ�", "", 2110, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CopyCheck, "���ƵǼ�", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Redo, "ȡ���Ǽ�", "", 742, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReGet, "�ٻ�ȡ��", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ThingModi, "�޸���Ϣ", "", 0, False)
'        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ModifBaseInfo, "������Ϣ����", "", 4113, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Receive, "��鱨��", "", 744, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Logout, "ȡ������", "", 743, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_InQueue, "���", "��ʼ�Ŷ�", 3534, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Transfer, "����Ӱ��", "", 505, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Cancel, "ȡ������", "", 506, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Review, "������Ϣ", "", 232, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CheckList, "�鿴��������", "�鿴�������뵥", 3564, True)
        If mlngModule = G_LNG_PACSSTATION_MODULE Then
            '���ݲ����жϱ���ͽ�Ƭ�ķ��ŷ�ʽ
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_Release, "���Ŵ���", "�����Ƭ���Ŵ���", 3013, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportFilmRelease, "��������", "", 3013, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "", 8215, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "��Ƭ����", "", 8216, False)
        Else
            Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "", 8215, False)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "����ִ��", "������", 3967, True)
        
        If mlngModule = G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SendArrange, "���Ͱ���", "", 232, False)
        End If
        
        'Σ��ֵ
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_CriticalSituation, "Σ��ֵ", "", 8338, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Normal, "��ͨ", "", 8344, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Critical, "Σ��", "", 8345, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CriticalValues, "��������", "", 8338, True)
    
        '�����
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_Result, "�����", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "�������", "", 3506, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "�������", "", 3507, False)

        '�������
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_FuHeLevel, "�������", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "����", "", 3587, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "��������", "", 3010, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "������", "", 3010, False)
        End If
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_GChannel, "��ɫͨ��", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelOk, "���", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelCancel, "ȡ��", "", 0, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Finish, "�ޱ������", "", 216, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ClearUp, "�ޱ������", "", 3012, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Complete, "������", "", 225, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Undone, "ȡ�����", "", 219, False)

        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_RelatingPatiet, "��������", "", 803, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Burn, "ͼ���¼", "", 0, True)
        
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Analyse, "�߼�����"): cbrControl.ToolTipText = "�߼�ͼ����"
        End If
        
    End With
    
    
    
'Begin-------------------------------------------------------�ղز˵�(Ĭ�Ͽɼ�)----------------------------------------------------------

    gstrSQL = "select ID ,�ϼ�id,������,�ղ���� from Ӱ���ղ���� where ������='" & UserInfo.���� & "' Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id"
    Set rsCollection = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)

    gstrSQL = "select ID ,�ϼ�id,������,�ղ����,�Ƿ��� from Ӱ���ղ���� where ������<>'" & UserInfo.���� & "' Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id"
    Set rsViewShare = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)
        
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Collection, "�ղ�", "", 0, False) ' Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Collection, "�ղ�", -1, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Collection_Manage, "�ղع���", "", 0, False) '.Add(xtpControlButton, conMenu_Collection_Manage, "�ղع���", -1, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Collection_To, "�ղص�...", "", 0, False) '.Add(xtpControlButton, conMenu_Collection_To, "�ղص�...")
        
        
        '��¡���� ɸѡ����������ݽ����ж�
        Set rsShareCount = zlDatabase.CopyNewRec(rsViewShare)
        rsShareCount.Filter = "�Ƿ���=1"
        
        If rsShareCount.RecordCount <> 0 Then
           '�ݹ鴴������˵�
           mlngShareFatherID = 0
           Set rsTemp = zlDatabase.CopyNewRec(rsViewShare)
           rsViewShare.Filter = "�ϼ�ID=" & Nvl(rsViewShare!�ϼ�ID, 1) & " and ������<> '" & UserInfo.���� & "'"
           
           Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Collection_ViewShare, "����鿴", "", 0, False)
           Call RecursionCreateShareMenu(rsViewShare, rsTemp, cbrControl)
        End If

        If rsCollection.RecordCount > 0 Then
            '�ݹ鴴���ղ����˵�
                 mlngCollectionFatherID = 0
                 Set rsTemp = zlDatabase.CopyNewRec(rsCollection)
                 rsCollection.Filter = "�ϼ�ID=" & Nvl(rsCollection!�ϼ�ID, 1)
                 Call RecursionCreateCollectionMenu(rsCollection, rsTemp, cbrMenuBar)
        End If
        
    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
'-----------------------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)")
    cbrMenuBar.ID = conMenu_ReportPopup
    
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModule, mstrPrivs, _
                                    "ZL1_INSIDE_1294_01", _
                                    "ZL1_INSIDE_1294_02", _
                                    "ZL1_INSIDE_1294_03", _
                                    "ZL1_INSIDE_1294_04", _
                                    "ZL1_INSIDE_1294_05", _
                                    "ZL1_INSIDE_1294_06", _
                                    "ZL1_INSIDE_1294_07", _
                                    "ZL1_INSIDE_1294_08", _
                                    "ZL1_INSIDE_1294_09", _
                                    "ZL1_INSIDE_1294_10", _
                                    "ZL1_INSIDE_1294_11", _
                                    "ZL1_INSIDE_1294_12", _
                                    "ZL1_INSIDE_1294_13")
    
    If cbrMenuBar.CommandBar.Controls.Count > 0 Then
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        For i = 1 To cbrMenuBar.CommandBar.Controls.Count
            cbrMenuBar.CommandBar.Controls(i).Category = M_STR_MODULE_MENU_TAG
        Next i
    Else
        cbrMenuBar.Delete
    End If
    
'Begin----------------------�Զ����ѯ�˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Manage_Query, "��ѯ", "", 0, False)
    Call RefreshCustomQueryMenu(cbrMenuBar, mlngCur����ID)
    
    
'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ViewPopup, "�鿴", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_ToolBar, "������", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��", "", 0, False): cbrPopControl.Checked = True
            End With
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_FontSize, "�����С", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_S, "С����", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_M, "������", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_L, "������", "", 0, False)
            End With
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_StatusBar, "״̬��", "", 0, True): cbrControl.Checked = True
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_Filter * 10#, "������", "", 0, False)
'        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Filter, "���ٹ���", "", 0, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Refresh, "ˢ��", "", 0, False)
    End With


'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_HelpPopup, "����", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Help, "��������", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����", "", 0, False)
            With cbrControl.CommandBar
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Forum, "������̳", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Home, "������ҳ", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���", "", 0, False)
            End With
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_About, "���ڡ�", "", 0, True)
    End With
    

'---------------------�������Ͻǵ�ǰ����----------------------------------
    Set cbrControl = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_View_Filter * 10#, "������", "", 0, False): cbrControl.flags = xtpFlagRightAlign
            
    '���ұ���ʾ�����ɼ���ť
    If mblnUseActivexCapture Then
        Set cbrControl = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "�����ɼ�", "���������ɼ�����", 0, False): cbrControl.flags = xtpFlagRightAlign
    End If
        
'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Regist, "�Ǽ�", "���Ǽ�", 211, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Receive, "����", "��鱨��", 744, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Logout, "ȡ��", "ȡ������", 743, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_InQueue, "���", "��ʼ�Ŷ�", 3534, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, comMenu_Petition_View, "�鿴���뵥", "�鿴��ɨ������뵥ͼ��", 3935, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_CheckList, "�鿴��������", "�鿴�������뵥", 3564, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Review, "��ע", "������Ϣ", 232, True)
    
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Tool_Analyse, "�߼�"): cbrControl.ToolTipText = "�߼�ͼ����"
    End If
    
    'ֻ��Ӱ����ϵͳ�ž����û��л�����
    If mlngModule = G_LNG_PATHSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "�л�", "�л��û�", 3012, False, conMenu_Tool_Analyse)
    End If
        
    '���ݲ����жϱ���ͽ�Ƭ�ķ��ŷ�ʽ
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Release, "���Ŵ���", "�����Ƭ���Ŵ���", 3013, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportFilmRelease, "��������", "��������", 3013, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "���淢��", 8215, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "��Ƭ����", "��Ƭ����", 8216, False)
    Else
        Set cbrPopControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "���淢��", 8215, False)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "����ִ��", "������", 3967, False)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SendArrange, "���Ͱ���", "���Ͱ���", 232, False)
    End If
    
    'Σ�����
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_CriticalSituation, "Σ��ֵ", "Σ�����", 8338, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Normal, "��ͨ", "��ͨ", 8344, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Critical, "Σ��", "Σ��", 8345, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CriticalValues, "��������", "", 8338, True)
    
    '�����������
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Result, "���", "�����������", 3506, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "����", "����", 3506, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "����", "����", 3507, False)
    
    '����ǲ���ϵͳ����û�з��������ť
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_FuHeLevel, "�������", "�������", 8044, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "����", "����", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "��������", "��������", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "������", "������", 0, False)
    End If
        
    'ֻ��Ӱ��ɼ�ϵͳ�ž����û���������
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "����", "�������ҽ���ͱ���ҽ��", 3012, False)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Complete, "���", "����������", 225, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Filter, "����", "����", 0, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Refresh, "ˢ��", "ˢ��", 0, False)
  
  
     '��ʼ���������� �ӵ�����Ϊ�˷�ֹ��һЩ���������ʱ�򣬻ᵼ������ָ��ɳ�ʼ��
    Call SetFontSize(IIf(mbytFontSize = 12, 1, IIf(mbytFontSize = 15, 2, 0)))
    
'    '��������ģ������Ĳ˵�
'    Call CreateWorkModuleMenu
End Sub


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If

    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub CreateWorkModuleMenu()
'��������ģ��˵�
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
On Error GoTo err
    If Not mobjWork_Pathol Is Nothing And mblnIsLoadPatholModule Then
        Call mobjWork_Pathol.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mobjWork_Pathol.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    '����Ӱ��ͼ��ģ����ز˵���������
    If Not mfrmWork_PacsImg Is Nothing And InStr(mstrWorkModule, ";Ӱ��ͼ��ģ��;") > 0 Then
        Call mfrmWork_PacsImg.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mfrmWork_PacsImg.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    If mblnUseActivexCapture Then
        'ʹ��ActivexExe��ͼ��ɼ��˵�����
        If Not mobjWork_ActiveVideo Is Nothing And InStr(mstrWorkModule, ";Ӱ��ɼ�ģ��;") > 0 Then
            'TODO:������Ƶ�ɼ�ģ��˵�
'            Call mobjWork_ActiveVideo.zlMenu.zlCreateMenu(Me.cbrMain)
'            Call mobjWork_ActiveVideo.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
        End If
    End If

    
    '���뽫����˵��Ĵ�������mobjWork_His�����˵�֮ǰ���������л�������ģ��ʱ����Ӧ��ģ��˵����ܹ�������
    If Not mobjWork_Report Is Nothing And _
        (InStr(mstrWorkModule, ";Ӱ�񱨸�ģ��;") > 0 Or InStr(mstrWorkModule, ";�������ģ��;") > 0) Then
        Call mobjWork_Report.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mobjWork_Report.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    If Not mobjWork_His Is Nothing Then
        '��Ϊ��PACSϵͳ�� ����ӡ�� �˵����ڱ༭�˵����£������������ļ��˵��£������ڵ��ò����Ĳ˵���������ʱ��
        '���ļ��˵����Ҳ�����ӡ�˵����������PACS�У��嵥��ӡ���ļ��˵��£����Ե��ò����Ĳ˵���������ʱ��
        '�嵥��ӡ��id�ĳɴ�ӡ��id��������󣬻ָ��嵥��ӡԭ����id
        If TabWindow.Selected.tag = "������Ӳ���" Or TabWindow.Selected.tag = "סԺ���Ӳ���" Then
            Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
            Set cbrControl = cbrMenuBar.CommandBar.Controls.Find(, conMenu_File_Excel)
            cbrControl.ID = conMenu_File_Print
        End If
        
        Call mobjWork_His.zlMenu.zlCreateMenu(Me.cbrMain)
        
        If TabWindow.Selected.tag = "������Ӳ���" Or TabWindow.Selected.tag = "סԺ���Ӳ���" Then
            cbrControl.ID = conMenu_File_Excel
        End If
    End If

    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call cbrMain.RecalcLayout
    Exit Sub
err:
    cbrControl.ID = conMenu_File_Excel
End Sub

Private Sub RecursionCreateShareMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrParentControl As CommandBarControl, Optional blnIsShare As Boolean = False)
'�ݹ�ѭ����������˵�
    Dim rsFilterTemp As ADODB.Recordset
    Dim i As Long
    Dim cbrControl As CommandBarControl
    Static j As Long
    
    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst
    
    With cbrParentControl.CommandBar.Controls
        If mlngShareFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + mlngShareFatherID, "�鿴��ǰ�ղ�", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If
        
        For i = 1 To rsFilterADO.RecordCount
            rsFullADO.Filter = " �ϼ�ID=" & Nvl(rsFilterADO!ID)

            If rsFullADO.RecordCount > 0 Then
                Set cbrControl = Nothing
  
                If Nvl(rsFilterADO!�Ƿ���) = 1 Or blnIsShare = True Then
                    mlngShareFatherID = Nvl(rsFilterADO!ID)
                    '���������˵� ����ϼ�ID=1 ����ʾ����������
                    Set cbrControl = .Add(xtpControlButtonPopup, CLng(conMenu_Collection_ViewShare) * 10000# + j, Nvl(rsFilterADO!�ղ����) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & Nvl(rsFilterADO!������) & ")", ""), -1, False)
                    cbrControl.DescriptionText = Nvl(rsFilterADO!������)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '�����Լ�
                Call RecursionCreateShareMenu(rsFilterTemp, rsFullADO, IIf(cbrControl Is Nothing, cbrParentControl, cbrControl), IIf(Nvl(rsFilterADO!�Ƿ���) = 0, False, True))
            Else
            '�����Ӽ��˵�
                If Nvl(rsFilterADO!�Ƿ���) = 1 Or blnIsShare = True Then
                    Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + j, Nvl(rsFilterADO!�ղ����) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & Nvl(rsFilterADO!������) & ")", ""), -1, False)
                    cbrControl.DescriptionText = Nvl(rsFilterADO!������)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                mlngShareFatherID = 0
            End If

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With
End Sub



Private Sub RecursionCreateCollectionMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrMenuBar As CommandBarPopup)
'�ݹ�ѭ�������ղ����˵�
    Dim rsFilterTemp As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim i As Long
    Static j As Long

    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst

    With cbrMenuBar.CommandBar.Controls
        If mlngCollectionFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + mlngCollectionFatherID, "�鿴��ǰ�ղ�", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If

        For i = 1 To rsFilterADO.RecordCount

            rsFullADO.Filter = " �ϼ�ID=" & Nvl(rsFilterADO!ID)
            mlngCollectionFatherID = Nvl(rsFilterADO!ID)
            If rsFullADO.RecordCount > 0 Then
            '���������˵�
                Set cbrControl = .Add(xtpControlButtonPopup, CLng(comMenu_Collection_Type) * 10000# + j, Nvl(rsFilterADO!�ղ����), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '�����Լ�
                Call RecursionCreateCollectionMenu(rsFilterTemp, rsFullADO, cbrControl)
                
            Else
            '�����Ӽ��˵�
                Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + j, Nvl(rsFilterADO!�ղ����), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
            End If
            If i = 1 Then cbrControl.BeginGroup = True

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With

End Sub


Private Sub ReadWorkModuleCfg()
    '���õ�ǰ��Ҫ�����Ĺ���ҳ��
    mstrWorkModule = zlDatabase.GetPara("վ��ģ��", glngSys, mlngModule, "")
    mstrWorkModule = IIf(mstrWorkModule <> "", ";" & mstrWorkModule & ";", "")
    
    '���ģ��Ϊ�գ������������Ŷӽкţ���ֻ��ʾ�ŶӽкŹ���ģ��
    If mstrWorkModule = "" Then 'And Not mblnUseQueue
        Select Case mlngModule
            Case G_LNG_PACSSTATION_MODULE
                mstrWorkModule = ";Ӱ��ͼ��ģ��;Ӱ�񱨸�ģ��;������¼ģ��;���ü�¼ģ��;ҽ����¼ģ��;"
            
            Case G_LNG_VIDEOSTATION_MODULE
                mstrWorkModule = ";Ӱ��ɼ�ģ��;Ӱ�񱨸�ģ��;������¼ģ��;���ü�¼ģ��;ҽ����¼ģ��;"
            
            Case G_LNG_PATHOLSYS_NUM
                mstrWorkModule = ";�걾����ģ��;Ӱ��ɼ�ģ��;����ȡ��ģ��;������Ƭģ��;�����ؼ�ģ��;���̱���ģ��;�������ģ��;������¼ģ��;���ü�¼ģ��;ҽ����¼ģ��;"
            Case Else
                Exit Sub
        End Select
    End If
    
'    '���Դ���
'    mstrWorkModule = ";Ӱ��ͼ��ģ��;Ӱ��ɼ�ģ��;�걾����ģ��;����ȡ��ģ��;������Ƭģ��;�����ؼ�ģ��;���̱���ģ��;Ӱ�񱨸�ģ��;���ü�¼ģ��;ҽ����¼ģ��;������¼ģ��;"
End Sub


Private Sub InitPatholModuleObj()
    '��ʼ���������ģ�����
    If mobjWork_Pathol Is Nothing Then
        Set mobjWork_Pathol = New clsWorkModule_Pathol
        Call mobjWork_Pathol.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
    End If
End Sub

Private Sub InitHisModuleObj()
    '��ʼ��HIS���ģ�����
    If mobjWork_His Is Nothing Then
        Set mobjWork_His = New clsWorkModule_His
        
        If mblnAllDepts Then
            Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, UserInfo.����ID, Me)
        Else
            Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
        End If
    End If
End Sub

Private Sub InitActiveVideoModuleObj()
'��ʼ��ActivexExe��Ƶ�ɼ�ģ�����
    If mlngModule = G_LNG_PACSSTATION_MODULE Then Exit Sub
    If Not CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Exit Sub
    If InStr(mstrWorkModule, ";Ӱ��ɼ�ģ��;") < 0 Then Exit Sub
    
    If mobjWork_ActiveVideo Is Nothing Then
        Set mobjWork_ActiveVideo = CreateObject("zl9PacsCapture.clsPacsCapture") ' New zl9PacsCapture.clsPacsCapture
        
        mobjWork_ActiveVideo.ParentWindowKey = Me.Name
        mobjWork_ActiveVideo.AllowEventNotify = True
        
        Call mobjWork_ActiveVideo.RegEventObj(Me)
        
        Call mobjWork_ActiveVideo.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngCur����ID, Me.hWnd, Me, True, gblnUseDebugLog)
    End If
End Sub

Private Sub ShowModuleLoadState(Optional ByVal strState As String = "")
'��ʾ����״̬
On Error GoTo ErrHandle
    picLoadState.Left = 0
    picLoadState.Top = 350
    picLoadState.Width = picWindow.Width - 0
    picLoadState.Height = picWindow.Height - 350
    
    
    If strState <> "" Then
        labLoadState.Caption = strState
        Call picLoadState_Resize
    End If
    
    picLoadState.Visible = True
    
ErrHandle:
End Sub

Private Sub HideModuleLoadState()
'��������״̬
    picLoadState.Visible = False
End Sub

Public Sub InitSubForm()
    Dim i As Integer
    Dim blnDoEvents As Boolean

    mblnIsLoadPatholModule = False   '���ñ��������ȻΪfalseʱ�����������ɾ������˵�
    blnDoEvents = True  '��ֵΪtrueʱ�������ι���ģ����ع����е��¼�����
    
    Call ShowModuleLoadState
    DoEvents
    
    With TabWindow
        .RemoveAll
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ButtonMargin.Top = 4
        .PaintManager.ButtonMargin.Bottom = 4
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        '��ȡ����ģ������
        Call ReadWorkModuleCfg
    
        If InStr(mstrWorkModule, ";Ӱ��ͼ��ģ��;") > 0 Then
            '����Ӱ���¼ģ��
            If mfrmWork_PacsImg Is Nothing Then
                Set mfrmWork_PacsImg = New frmWork_Image
                
                Set mfrmWork_PacsImg.PacsCore = mobjPacsCore
                Call mfrmWork_PacsImg.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
            End If
    
            .InsertItem 0, "Ӱ���¼", picTemp.hWnd, conMenu_Img_Look
            .Item(TabWindow.ItemCount - 1).tag = "Ӱ��ͼ��"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
            
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mfrmWork_PacsImg Is Nothing Then
                Call mfrmWork_PacsImg.zlMenu.zlClearMenu
                Call mfrmWork_PacsImg.zlMenu.zlClearToolBar
            End If
        End If
                        
        If mlngModule <> G_LNG_PACSSTATION_MODULE And CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") _
            And InStr(mstrWorkModule, ";Ӱ��ɼ�ģ��;") > 0 Then
            
            If mobjCaptureHot Is Nothing Then
                Set mobjCaptureHot = New zl9PacsControl.clsHookKey
                Call mobjCaptureHot.EnableHook(WM_KEYDOWN, True)
            End If

            If mblnUseActivexCapture Then
                Call InitActiveVideoModuleObj
                
                .InsertItem 1, "Ӱ��ɼ�", mobjWork_ActiveVideo.ContainerHwnd, conMenu_Cap_Dynamic
                .Item(TabWindow.ItemCount - 1).tag = "Ӱ��ɼ�"
            End If


            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            If mblnUseActivexCapture Then
                'TODO:ʹ��activex��Ƶ�ɼ���ʽ��Ĳ˵�����...
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "�걾����") And InStr(mstrWorkModule, ";�걾����ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 2, "�걾����", picTemp.hWnd, G_INT_ICONID_SPECIMEN
            .Item(TabWindow.ItemCount - 1).tag = "�걾����"
            
            mblnIsLoadPatholModule = True

            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "����ȡ��") And InStr(mstrWorkModule, ";����ȡ��ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 3, "����ȡ��", picTemp.hWnd, G_INT_ICONID_MATERIAL
            .Item(TabWindow.ItemCount - 1).tag = "����ȡ��"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "������Ƭ") And InStr(mstrWorkModule, ";������Ƭģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 4, "������Ƭ", picTemp.hWnd, G_INT_ICONID_SLICES
            .Item(TabWindow.ItemCount - 1).tag = "������Ƭ"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If (CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "����Ⱦɫ") Or CheckPopedom(mstrPrivs, "���Ӳ���")) _
            And InStr(mstrWorkModule, ";�����ؼ�ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 5, "�����ؼ�", picTemp.hWnd, G_INT_ICONID_SPEEXAM
            .Item(TabWindow.ItemCount - 1).tag = "�����ؼ�"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If (CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "��Ⱦ����") _
            Or CheckPopedom(mstrPrivs, "���ӱ���") Or CheckPopedom(mstrPrivs, "���߱���") _
            Or CheckPopedom(mstrPrivs, "�����ؼ챨�����")) And InStr(mstrWorkModule, ";���̱���ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 6, "����/�ؼ챨��", picTemp.hWnd, G_INT_ICONID_PROREPORT
            .Item(TabWindow.ItemCount - 1).tag = "���̱���"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If GetInsidePrivs(p���Ʊ������, True) <> "" And _
            (InStr(mstrWorkModule, ";Ӱ�񱨸�ģ��;") > 0 Or InStr(mstrWorkModule, ";�������ģ��;") > 0) Then
            
            If mobjWork_Report Is Nothing Then
                Set mobjWork_Report = New clsWorkModule_Report
                
                Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
                
                Set mobjWork_Report.PacsCore = mobjPacsCore
            End If

            .InsertItem 7, "Ӱ�񱨸�", picReportContainer.hWnd, 10008 'conMenu_Edit_Compend
            .Item(TabWindow.ItemCount - 1).tag = "������д"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlMenu.zlClearMenu
                Call mobjWork_Report.zlMenu.zlClearToolBar
            End If
        End If
        
        
        If Not mblnIsLoadPatholModule And Not mobjWork_Pathol Is Nothing Then
            'û�м��ز���ģ�飬��mobjWork_Pathol��Ϊ��ʱ��ɾ������˵�
            Call mobjWork_Pathol.zlMenu.zlClearMenu
            Call mobjWork_Pathol.zlMenu.zlClearToolBar
        End If
        
        
        If GetInsidePrivs(pҽ�����ѹ���, True) <> "" And InStr(mstrWorkModule, ";���ü�¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 8, "���ü�¼", picTemp.hWnd, 10007
            .Item(TabWindow.ItemCount - 1).tag = "�������"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" And InStr(mstrWorkModule, ";ҽ����¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 9, "ҽ����¼", picTemp.hWnd, 10010
            .Item(TabWindow.ItemCount - 1).tag = "סԺҽ��"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(p����ҽ���´�, True) <> "" And InStr(mstrWorkModule, ";ҽ����¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 10, "ҽ����¼", picTemp.hWnd, 10010  ' conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).tag = "����ҽ��": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(pסԺ��������, True) <> "" And InStr(mstrWorkModule, ";������¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 11, "������¼", picTemp.hWnd, 10009 ' conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).tag = "סԺ����"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(p���ﲡ������, True) <> "" And InStr(mstrWorkModule, ";������¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 12, "������¼", picTemp.hWnd, 10009 ' conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).tag = "���ﲡ��": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(p������Ӳ���, True) <> "" And InStr(mstrWorkModule, ";���Ӳ���ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 13, "���Ӳ���", picTemp.hWnd, 10009
            .Item(TabWindow.ItemCount - 1).tag = "������Ӳ���": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(pסԺ���Ӳ���, True) <> "" And InStr(mstrWorkModule, ";���Ӳ���ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 14, "���Ӳ���", picTemp.hWnd, 10009
            .Item(TabWindow.ItemCount - 1).tag = "סԺ���Ӳ���": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        '����Ŷӽк�ҳ��
        If mSysPar.blnUseQueue = True Then
            mstrWorkModule = mstrWorkModule & ";�Ŷӽк�ģ��;"
            
            If mobjQueue Is Nothing Then
                Set mobjQueue = New frmWork_Queue
                Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, NeedName(mstrCur����), mstrPrivs)
            End If
            
            .InsertItem 15, "�Ŷӽк�", picTemp.hWnd, 10011
            .Item(TabWindow.ItemCount - 1).tag = "�Ŷӽк�"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
    
'        If Not GetVideoForm Is Nothing Then Call GetVideoForm.ShowVideoWindow(picVideoContainer)
    End With
    
    DoEvents
    
    If GetWorkModuleCount = 1 Then
        TabWindow.PaintManager.ClientMargin.Top = -30
    Else
        TabWindow.PaintManager.ClientMargin.Top = 0
    End If
    
    Call HideModuleLoadState
End Sub

Private Function GetWorkModuleCount() As Long
'��ȡ�ɼ�tabwindow������
    Dim i As Long
    Dim lngCount As Long
    Dim aryWorkModule() As String
    
    
    aryWorkModule = Split(mstrWorkModule, ";")
    
    For i = LBound(aryWorkModule) To UBound(aryWorkModule)
        If aryWorkModule(i) <> "" Then lngCount = lngCount + 1
    Next i
    
    GetWorkModuleCount = lngCount
End Function


Private Function GetTabWindowIndex() As Long
'��ȡ��һ���ɼ�tabwindow������
    Dim i As Long
    
    GetTabWindowIndex = -1
    For i = 0 To TabWindow.ItemCount - 1
        If TabWindow.Item(i).Visible Then
            GetTabWindowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub mobjWork_Report_AfterDeleted(ByVal lngOrderID As Long)
    Call AfterDeleted(lngOrderID)
End Sub

Private Sub mobjWork_Report_AfterDeletedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Call AfterDeletedRich(lngOrderID, strDocId)
End Sub

Private Sub mobjWork_Report_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjWork_Report_AfterPrintedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Call AfterPrintedRich(lngOrderID, strDocId)
End Sub

Private Sub mobjWork_Report_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Object, ByVal lngSaveType As Long)
    Call AfterReportSaved(lngOrderID, frmOwnerForm, lngSaveType)
End Sub

Private Sub mobjWork_Report_AfterSavedRich(ByVal lngOrderID As Long, ByVal strDocId As String, frmOwnerForm As Object, ByVal lngSaveType As Long)
    Call AfterReportSavedRich(lngOrderID, strDocId, frmOwnerForm, lngSaveType)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    On Error GoTo ErrHandle
    
    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.RefreshReportImage
    
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjQueue_OnDiagnose(ByVal lngAdviceId As Long, ByVal strExeRoom As String, ByVal strTurnPage As String)
'�Ŷӽ����¼�
On Error GoTo ErrHandle
    Dim lngIndex As String
    Dim i As Long
    
    lngIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID", True)
    
    If lngIndex > 0 Then
        Call ufgStudyList.LocateRow(lngIndex)
        
        If Trim(strTurnPage) <> "" Then
            '��ת��ָ���Ĺ���ģ��

            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow(i).tag, strTurnPage) > 0 And TabWindow(i).Visible Then
                    TabWindow(i).Selected = True
                    Exit For
                End If
            Next i
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjQueue_OnCompleted(ByVal lngAdviceId As Long, ByVal strExeRoom As String)
'�Ŷ�����¼�
On Error GoTo ErrHandle
    Dim lngIndex As String

    
    lngIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID", True)
    
    If lngIndex > 0 Then
        Call ufgStudyList.LocateRow(lngIndex)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjQueue_OnSelChange(ByVal lngAdviceId As Long)
'��ѡ��ı��¼�
On Error GoTo ErrHandle
    Dim lngIndex As Long
    
    If mSysPar.blnSynStudylist Then
        lngIndex = ufgStudyList.FindRowIndex(lngAdviceId, "ҽ��ID", True)
        
        If lngIndex > 0 Then
            Call ufgStudyList.LocateRow(lngIndex)
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub AfterDeletedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Dim intState As Integer
    Dim lngSendNo As Long
    
On Error GoTo ErrHandle
    intState = getStudyStateRich(lngOrderID, strDocId, False, , lngSendNo)
    If intState = 6 Then Exit Sub
    
    gstrSQL = "Zl_Ӱ����_״̬����(" & lngOrderID & "," & lngSendNo & ",''," & intState & ",0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������״̬��������")
    
    If intState < 4 Then
        gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
    End If
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
On Error GoTo ErrHandle
    gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub AfterPrintedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    Dim bln���������� As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim blnAllReportFinished As Boolean
    Dim intState As Integer, lngSendId As Long
    
    strResultInput = ""
    
    intState = getStudyStateRich(lngOrderID, strDocId, False, blnAllReportFinished, lngSendId, bln����������, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    BUGEX "AfterPrintedRich lngOrderID=" & lngOrderID & "   strDocId=" & strDocId
    
    strSql = "Select B.Σ��״̬, A.�������, B.Ӱ������, A.��������, B.�������,B.ҽ��ID " & _
                 "From Ӱ�񱨸��¼ A, Ӱ�����¼ B " & _
                 "Where A.ID=[1] and A.ҽ��id = B.ҽ��id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�������", strDocId)
    
    If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "Σ��״̬|"
    If (Not bln���������� And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "�������|"
    If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "Ӱ���ʿ�") > 0 Then strResultInput = strResultInput & "Ӱ������|"
    If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "�����ʿ�") > 0 Then strResultInput = strResultInput & "��������|"
    If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "�������|"
    
    If strResultInput <> "" Then Call PromptResultRich(lngOrderID, strDocId, mlngModule, Me, mlngCur����ID, strResultInput)
    
    If mSysPar.blnPrintCommit = True Then
        If blnAllReportFinished Then    '���б�����ɺ�������
            Call Menu_Manage_����������(lngOrderID, False, strDocId)
        End If
    End If
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub AfterPrinted(lngOrderID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    
    strResultInput = ""
    gstrSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "���´�ӡ���"
    
    strSql = "Select B.Σ��״̬, A.�������, B.Ӱ������, B.��������, B.������� " & _
             "From ����ҽ������ A, Ӱ�����¼ B " & _
             "Where A.ҽ��id = B.ҽ��id and B.ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�������", lngOrderID)
    
    If IsNull(rsTemp!Σ��״̬) And mSysPar.lngCriticalValues <> 0 Then strResultInput = "Σ��״̬|"
    If IsNull(rsTemp!�������) And Not mSysPar.blnIgnoreResult Then strResultInput = strResultInput & "�������|"
    If IsNull(rsTemp!Ӱ������) And mSysPar.strImageLevel <> "" And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "Ӱ���ʿ�") > 0 Then strResultInput = strResultInput & "Ӱ������|"
    If IsNull(rsTemp!��������) And mSysPar.strReportLevel <> "" And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "�����ʿ�") > 0 Then strResultInput = strResultInput & "��������|"
    If IsNull(rsTemp!�������) And mSysPar.lngConformDetermine <> 0 Then strResultInput = strResultInput & "�������|"

    If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, Me, mlngCur����ID, strResultInput)
    
    If mSysPar.blnPrintCommit = True Then
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub AfterReportSavedRich(ByVal lngOrderID As Long, ByVal strDocId As String, frmOwnerForm As Form, ByVal lngSaveType As Long)
'���汨��֮��Ĵ���
'ִ�й��̣�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����
On Error GoTo ErrHandle
    Dim intState As Integer, lngSendId As Long
    Dim strǩ�� As String
    Dim str������ As String
    Dim str������ As String
    Dim bln���������� As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim blnAllReportFinished As Boolean
    
    arrSQL = Array()

    Call mobjWork_Report.zlRefreshFace(True)
    
    'intState =1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--�����
    BUGEX "AfterReportSavedRich lngOrderID=" & lngOrderID & "   strDocId=" & strDocId
    
    '��ȡ���μ���ִ�й���
    intState = getStudyStateRich(lngOrderID, strDocId, False, blnAllReportFinished, lngSendId, bln����������, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    If intState = 2 Or intState = 3 Then
        '���汣��ʱִ�з���
        If mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngMoneyExeModle = 2 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            gstrSQL = "Zl_Ӱ�����ִ��(" & lngOrderID & "," & lngSendId & ",4,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    gstrSQL = "Zl_Ӱ����_״̬����(" & lngOrderID & "," & lngSendId & ",'" & strDocId & "'," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gcnOracle.BeginTrans        '----------������״̬��������
    
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������״̬��������")
    Next i
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If (intState = 4 Or intState = 5) And IIf(mSysPar.lngHintType = 0, lngSaveType = 1, IIf(mSysPar.lngHintType = 1, lngSaveType = 2, False)) Then
        Dim strResultInput As String
        
        strResultInput = ""
        If mSysPar.blnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
            Call mobjWork_Report.Menu_Manage_�������(mListAdviceInf.lngAdviceId, "0")
        End If
            
        If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "Σ��״̬|"
        If (Not bln���������� And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "�������|"
        If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "Ӱ���ʿ�") > 0 Then strResultInput = strResultInput & "Ӱ������|"
        If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "�����ʿ�") > 0 Then strResultInput = strResultInput & "��������|"
        If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "�������|"
 
        If strResultInput <> "" Then Call PromptResultRich(lngOrderID, strDocId, mlngModule, frmOwnerForm, mlngCur����ID, strResultInput)
    End If
    
    If intState = 5 And mSysPar.blnCompleteCommit Then   '�������˺�ֱ����ɡ�
        If blnAllReportFinished Then '���б�����ɺ�������
            Call Menu_Manage_����������(lngOrderID, False, strDocId)
        End If
    End If
    '����״̬����
    Call StateCheck(intState)
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(lngOrderID)
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub AfterReportSaved(lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long)
'���汨��֮��Ĵ���
'ִ�й��̣�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����
On Error GoTo ErrHandle
    Dim intState As Integer, lngSendId As Long
    Dim strǩ�� As String
    Dim str������ As String
    Dim str������ As String
    Dim bln���������� As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()

    Call mobjWork_Report.zlRefreshFace(True)

    '��ȡ���μ���ִ�й���
    intState = getStudyState(lngOrderID, False, lngSendId, str������, strǩ��, str������, bln����������, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    'intState =1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--����ɣ������̲������������ֵ��
    If intState = 2 Or intState = 3 Then
        gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str������) & "','')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        '���汣��ʱִ�з���
        If mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngMoneyExeModle = 2 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            gstrSQL = "Zl_Ӱ�����ִ��(" & lngOrderID & "," & lngSendId & ",4,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    Else
        If intState = 4 Then
            '���ǩ�������һ��ǩ��Ϊҽʦ,ִ�й���Ϊ�ѱ���
            '�п��ܵ���� 1-ҽʦ��N��ǩ�� 2-���μ������һ����ǩ 3-�޶�ģʽ�±���(ǩ������=0)
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            'Ӧ����д�����˲�׼ȷ�����˵�ʱ�򣬻��˵����Ǳ����ˣ����ǲ��Ǳ��洴����
            'ҽ�����ǩ��,�����ǵ�N�Σ���ʱ����������Ҫ���棬��������Ҫ���;
            gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str������) & "','')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        ElseIf intState = 5 Then
            '���ǩ�������μ����ϼ���ǩ����ǩ������>=2,ִ�й���Ϊ�����
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str������) & "','" & IIf(strǩ�� <> "", strǩ��, str������) & "')"
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
    
    If (intState = 4 Or intState = 5) And IIf(mSysPar.lngHintType = 0, lngSaveType = 1, IIf(mSysPar.lngHintType = 1, lngSaveType = 2, False)) Then
        Dim strResultInput As String
        
        strResultInput = ""
        If mSysPar.blnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
            gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
            zlDatabase.ExecuteProcedure gstrSQL, "���������"
        End If
            
        If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "Σ��״̬|"
        If (Not bln���������� And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "�������|"
        If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "Ӱ���ʿ�") > 0 Then strResultInput = strResultInput & "Ӱ������|"
        If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "�����ʿ�") > 0 Then strResultInput = strResultInput & "��������|"
        If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "�������|"
 
        If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, frmOwnerForm, mlngCur����ID, strResultInput)
    End If
    
    If intState = 5 And mSysPar.blnCompleteCommit Then   '�������˺�ֱ����ɡ�
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    '����״̬����
    Call StateCheck(intState)
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(lngOrderID)
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub UpdateStudyListState(lngAdviceId As Long, strStudyUID As String, blnAddImage As Boolean, blnStateChanged As Boolean)
    Dim strSql As String
    Dim intRowIndex As Integer
    
    With ufgStudyList
        intRowIndex = .FindRowIndex(CStr(lngAdviceId), "ҽ��ID", True)
        
        If blnStateChanged And intRowIndex > 0 Then
            If blnAddImage Then '��ͼ
                .Text(intRowIndex, "���UID") = Nvl(strStudyUID, "A123456789")
                Call .UpdateSourceData(lngAdviceId, "���UID", Nvl(strStudyUID, "A123456789"))
                
                Set .DataGrid.Cell(flexcpPicture, intRowIndex, .GetColIndex(GetStudyNumberDisplayName)) = imgList.ListImages(IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "����", "Ӱ��")).Picture '�ı�ͼ��
                
                If .Text(intRowIndex, "������") = "�ѱ���" Then
                    .Text(intRowIndex, "������") = "�Ѽ��"
                    Call .UpdateSourceData(lngAdviceId, "������", 3)
                    
                    .Text(intRowIndex, "���״̬") = 3
                End If
            Else '���һ�β�ͼ
                .Text(intRowIndex, "���UID") = ""
                Call .UpdateSourceData(lngAdviceId, "���UID", "")
                
                Set .DataGrid.Cell(flexcpPicture, intRowIndex, .GetColIndex(GetStudyNumberDisplayName)) = Nothing '�ı�ͼ��
                
                If .Text(intRowIndex, "������") = "�Ѽ��" Then
                    .Text(intRowIndex, "������") = "�ѱ���"
                    Call .UpdateSourceData(lngAdviceId, "������", 2)
                    
                    .Text(intRowIndex, "���״̬") = 2
                End If
            End If
        End If
        
        '�������ø���Ӱ���鼼ʦ
        If mSysPar.blnWriteCapDoctor = True And blnStateChanged = True Then
            If mblnCnOracleIsHIS Then
                strSql = "Zl_Ӱ����_��鼼ʦ( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrUserNameNew, "") & "')"
                .Text(intRowIndex, "��鼼ʦ") = IIf(blnAddImage = True, mstrUserNameNew, "")
            Else
                strSql = "Zl_Ӱ����_��鼼ʦ( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrUserNameHIS, "") & "')"
                .Text(intRowIndex, "��鼼ʦ") = IIf(blnAddImage = True, mstrUserNameHIS, "")
            End If
            
            zlDatabase.ExecuteProcedure strSql, GetWindowCaption
        End If
    End With
End Sub

Private Sub StateCheck(ByVal intState As Integer, Optional ByVal lngAdviceId As Long)
'----------------------------------------------------------
'���ܣ��ڲ����б��ж�λָ���ļ�¼
'������ intState --���˼��״̬   lngAdviceID --����ҽ��ID
'���أ��ޣ�ֱ���ڲ����б��ж�λ
'----------------------------------------------------------
    If mSysPar.blnPatTrack Then
        If Not mblncmd�Ǽ� And Not mblncmd���� And Not mblncmd��� And Not mblncmd���� And Not mblncmd��� And Not mblncmd���� And Not mblncmd��� Then
            Call RefreshList(lngAdviceId)
            Exit Sub
        End If
        
        Select Case intState '���ݲ�����״̬ȷ����״̬�����Ƿ�ѡ��
            Case -1
                If Not mblncmd���� Then mblncmd���� = True
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
        
        Call RefreshList(lngAdviceId)
    Else '������ֻˢ���б�
        Call RefreshList
    End If
End Sub

Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSql As String
        
    On Error GoTo errH
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Function
    End If
    
    objPopup.CommandBar.Controls.DeleteAll
    
    strSql = "Select Distinct C.���,C.����,C.˵��" & _
        " From ����ҽ����¼ A,��������Ӧ�� B,�����ļ��б� C" & _
        " Where A.ID=[1] And A.���ID IS NULL" & _
        " And A.������ĿID=B.������ĿID" & _
        " And B.Ӧ�ó���=[2] And B.�����ļ�ID=C.ID And C.����=7" & _
        " Order by C.���"
        
    If mListAdviceInf.intMoved = 1 Then
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mListAdviceInf.lngAdviceId, mListAdviceInf.lngPatientFrom)
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!���� & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
            objControl.Category = M_STR_MODULE_MENU_TAG
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
On Error GoTo ErrHandle
    If objControl.Parameter = "" Then '��֣�ֱ�Ӱ�F10ʱ����һ���յ�Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    
    If objControl.Parameter = "" Then Exit Sub
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & ufgStudyList.CurText("NO"), _
                       "����=" & Val(ufgStudyList.CurText("��¼����")), "ҽ��ID=" & mListAdviceInf.lngAdviceId, 1)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub NotificationAllModuleRefresh()
'֪ͨ����ģ��ˢ��
    If Not mobjWork_His Is Nothing Then Call mobjWork_His.NotificationRefresh(hmAll)
    If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtAll)
    If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.NotificationRefresh
    
    If mblnUseActivexCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then Call mobjWork_ActiveVideo.zlNotifyRefresh
    End If

    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.NotificationRefresh
End Sub


Private Sub DisableWorkModule()
'���ù���ģ��
    tcDisable.Visible = True
    tcDisable.Translucence
End Sub


Private Sub EnableWorkModule()
'�򿪹���ģ��
    tcDisable.Visible = False
End Sub


Public Sub RefreshList(Optional ByVal lngAdviceId As Long = 0, Optional ByVal blnFromDB As Boolean = True)
'ˢ�������б�
    Dim i As Integer
    Dim lngcurҽ��ID As Long
    Dim lngRow As Long
    Dim lngTopRow As Long
    
    If blnIsLoading = True Or ufgStudyList.IsLoading = True Then
        MsgBoxD Me, "���ݼ����У����Ժ�����...", vbInformation, Me.Caption
        Exit Sub
    End If
    
    blnIsLoading = True

On Error GoTo ErrHandle
    
    mblnAutoRefreshList = True
    
    With ufgStudyList
        If lngAdviceId <> 0 Then
            lngcurҽ��ID = lngAdviceId
        Else
            lngcurҽ��ID = Val(ufgStudyList.CurKeyValue) '��ǰ��ҽ��ID
            lngRow = .DataGrid.Row: lngTopRow = .DataGrid.TopRow               '��ǰ�кͶ���֮��Ĳ��
        End If
    
        
        Call LoadPatiList(blnFromDB)
        
        If ufgStudyList.GridRows <= 1 Then
            '��û������ʱ��֪ͨˢ�¹���ģ������ص�����
            mcurAdviceInf = GetNullAdviceInf
            mListAdviceInf = GetNullAdviceInf
            
            Call RefreshModuleAdviceInf
            Call NotificationAllModuleRefresh
            
            If TabWindow.Selected Is Nothing Then
                'ѡ���һ������ģ��
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow.Item(i).Visible Then
                        TabWindow(i).Selected = True
                        
                        mblnAutoRefreshList = False
                        Exit For
                    End If
                Next i
            End If
            
            Call RefreshTabWindow
            
            mblnAutoRefreshList = False
            blnIsLoading = False
            
            Exit Sub
        End If
        
        
        If lngcurҽ��ID = 0 Then
            'Call .LocateRow(1)
            Call ufgStudyList_OnSelChange
            
            mblnAutoRefreshList = False
            blnIsLoading = False
            
            Exit Sub
        End If
        
        '�м�¼ʱҪ���¶�λ��֮ǰ��¼\
        lngcurҽ��ID = .FindRowIndex(CStr(lngcurҽ��ID), "ҽ��ID", True)
        
        If lngcurҽ��ID <> -1 Then
            lngRow = Abs(lngRow - lngTopRow)
            If .DataGrid.Row = lngcurҽ��ID Then '����δ�����ı�ʱ�����ᴥ��OnSelChange�¼�����˵�����ͬʱ���ֶ�����CHANGE�¼�
                Call ufgStudyList_OnSelChange  'ǿ��ˢ���ұ��Ӵ���
            Else
                .DataGrid.Row = lngcurҽ��ID
            End If
            
            .DataGrid.TopRow = IIf((.DataGrid.Row - lngRow) < 1, 1, (.DataGrid.Row - lngRow))
        Else
            If .DataGrid.Row <> 1 Then
                .DataGrid.Row = 1
            Else
                Call ufgStudyList_OnSelChange 'ǿ��ˢ���ұ��Ӵ���
            End If
        End If
        
    End With
    
    mblnAutoRefreshList = False
    blnIsLoading = False

    Exit Sub
ErrHandle:
    blnIsLoading = False
End Sub


Private Function GetExecuteState() As Long
'��ȡ�������ִ��״̬
    GetExecuteState = -1
    
    Select Case True
        Case optNeed.value And optNeed.Enabled
            GetExecuteState = 0
        Case optAccept.value And optAccept.Enabled
            GetExecuteState = 1
        Case optFinal.value And optFinal.Enabled
            GetExecuteState = 2
        Case optAll.value And optAll.Enabled
            GetExecuteState = 3
    End Select
End Function


Private Function GetFilterData() As ADODB.Recordset
'���ܣ�ȡ�õ�ǰ���˵�SQL
    Dim strSQLBak As String
    Dim str��Դ As String
    
    Dim lngType As Long          '1-ʹ��ҽ�������Ϊ����������2-ʹ�ü�������Ϊ��������
    Dim strSql As String
    Dim strSubSql As String
    
    Dim strWithOrderTab As String   'ҽ���Ӳ�ѯ
    Dim strWithStudyTab As String   '����Ӳ�ѯ
    Dim strWithOrderCols As String  'ҽ����ز�ѯ��
    Dim strWithStudyCols As String  '�����ز�ѯ��
    
    Dim strFilterOrder As String            'ҽ����Ϣ�������
    Dim strFilterStudy As String            '�����Ϣ�������
    Dim strFilterDate As String             '��ѯ���ڷ�Χ����
    Dim strFilterReportContext As String    '���ݱ������ݹ���
    Dim strFilterReportAdvice As String     '���ݱ��潨�����
    Dim strFilterIllnessDiagnose As String  '���ݼ�����Ϲ���
    Dim strFilterTemp As String
    
    Dim strPatholCol As String
    Dim strPatholTab As String
    Dim strPatholFilter As String
    
    Set GetFilterData = Nothing
    
    strPatholCol = ""
    strPatholTab = ""
    strPatholFilter = ""
    
    '�ж��Ƿ����Ӳ����ѯ����ر�
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strPatholTab = " ��������Ϣ o, ���������Ϣ p "
        strPatholCol = " o.ȡ�Ĺ���,o.��Ƭ����,o.���߹���,o.���ӹ���,o.��Ⱦ����,o.�������,o.�����,o.�ۺ����� "
                        
        strPatholFilter = " h.ҽ��ID=o.ҽ��ID and o.����ҽ��ID=p.����ҽ��ID(+) "
    End If
    
    lngType = 1 'Ĭ��ʹ��ҽ����ؼ���
    
    With SQLCondition
        '�������������ʹ��ʱ������
        If .����� <> 0 Then
            strFilterOrder = " And C.�����=[1]"
        ElseIf .סԺ�� <> 0 Then
            strFilterOrder = " And C.סԺ��=[2]"
        ElseIf .������ <> "" Then
            strFilterOrder = " And C.������=[8]"
        ElseIf .���￨ <> "" Then
            strFilterOrder = " And C.���￨��=[3]"
        ElseIf .���� <> "" And InStr(.����, "*") = 0 Then   '�������⴦����*�ű�ʾģ����ѯ
            .���� = .���� & "%"
            strFilterOrder = " And C.���� like [4]"
        ElseIf .���֤ <> "" Then
            strFilterOrder = " And C.���֤��=[5]"
        ElseIf .IC�� <> "" Then
            strFilterOrder = " And C.IC��=[6]"
        ElseIf .���ݺ� <> "" Then
            strFilterOrder = " And A.NO=[7] "
        ElseIf .���� <> 0 Then
            lngType = 2 'ʹ�ü�������Ϣ��Ϊ��������
            
            If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
            
            If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                strFilterStudy = strFilterStudy & " H.����=[9] "
            Else
                '����ǲ���ϵͳ����������Ҫ���ݲ���Ž��в�ѯ
                strFilterStudy = strFilterStudy & " o.�����=upper([9]) "
            End If
        ElseIf .����ID <> 0 Then
            strFilterOrder = " And C.����ID=[31]"
        Else
        '����������ѯ��ʹ��ʱ������
            
            '��д����ʱ������
            'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
            If .ʱ������ = 1 Then       '������ʱ��
                strFilterDate = " A.����ʱ�� Between [10] and "
            ElseIf .ʱ������ = 2 Then   '������ʱ��
                strFilterDate = " A.����ʱ�� Between [10] and "
            ElseIf .ʱ������ = 3 Then                        '��ͼʱ����߲����ڲ���������ʱ��
                lngType = 2
                
                If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                    strFilterDate = strFilterDate & " H.�������� Between [10] and  "
                Else
                    strPatholTab = strPatholTab & " , ����������Ϣ q"

                    strFilterDate = strFilterDate & " o.����ҽ��ID = q.����ҽ��ID and q.����ʱ�� between [10] and "
                End If
            Else                        '������ʱ��
                strFilterDate = strFilterDate & " B.��ʼִ��ʱ�� Between [10] and  "
            End If
            
            If .����ʱ�� <> CDate(0) Then
                strFilterDate = strFilterDate & " [11] "
            Else
                strFilterDate = strFilterDate & " Sysdate+1/(24*3600) "
            End If
            
            '�ȴ��������д�*�ŵģ����д�ʱ��������ģ����ѯ
            If .���� <> "" And InStr(.����, "*") <> 0 Then
                .���� = Replace(.����, "*", "%")
                strFilterOrder = strFilterOrder & " And C.���� || '' like [4]"
            End If
            
            If .�Ա� <> "" Then
                strFilterOrder = strFilterOrder & " And C.�Ա�=[27]"
            End If
        
        
            '��������-��ʼ����(ֻ�е�����ʹ�á����������ڶ�������֮��ʱ����ʹ�ÿ�ʼ����)
            If .��ʼ���� <> -1 Then
                If .�������� = "~" Then
                    strFilterOrder = strFilterOrder & " And ZL_AgeToDays(C.����)>=[28]"
                End If
            End If
            
            '��������-��������
            If .�������� <> -1 Then
                If .�������� = "~" Then
                    strFilterOrder = strFilterOrder & " And ZL_AgeToDays(C.����)<=[29]"
                Else
                    strFilterOrder = strFilterOrder & " And ZL_AgeToDays(C.����)" & .�������� & "[29]"
                End If
            End If
            
            If .���˿��� <> 0 Then
                strFilterOrder = strFilterOrder & " And B.���˿���ID+0=[12] "
            End If

            If .�걾��λ <> "" Then
                strFilterOrder = strFilterOrder & " And instr(B.ҽ������,[13])>0"
            End If
            
            If .������� <> -1 Then
                strFilterOrder = strFilterOrder & " And Nvl(A.�������, 0)=[30]"
            End If
            
            If .���ҽ�� <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.������=[14] "
            End If
            
            If .���ҽ�� <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.������=[15] "
            End If
            
            If .Ӱ������ <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.Ӱ������=[16] "
            End If
            
            If .��鼼ʦ <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.��鼼ʦ=[17] "
            End If
            
            'Ӱ������������ط�������������ѡ�񣬹��˴��ں����������棬���������е�Ϊ��
            If mintcmdӰ����� <= 0 Then
                If .Ӱ����� <> "" Then
                    If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                    strFilterStudy = strFilterStudy & " H.Ӱ�����=[18] "
                End If
            End If
            
            If .��� <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " Instr(H.�������,[19])>0 "
            End If
            
            If .������� <> "" Then
                strFilterIllnessDiagnose = "( Select t.ҽ��id From ����ҽ������ t Where t.����id IN " & _
                                                                        " (Select Distinct A.ID  " & _
                                                                        "From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                        "Where A.����ʱ��>[10] AND A.Id=B.�ļ�ID  " & _
                                                                            "And B.��������=7 And instr(B.��������,'52;')>0 And instr(B.�����ı�,[20])>0))"
            End If
            
            
            If .������� <> "" Then
                strFilterTemp = " (b.�����ı� ='�������' And Instr(c.�����ı�, [21]) > 0)"
            End If
            
            If .������ <> "" Then
                If strFilterTemp = "" Then
                    strFilterTemp = " (b.�����ı� ='������' And Instr(c.�����ı�, [22]) > 0)"
                Else
                    strFilterTemp = strFilterTemp & " or (b.�����ı� ='������' And Instr(c.�����ı�, [22]) > 0)"
                End If
            End If
            
            If .���� <> "" Then
                If strFilterTemp = "" Then
                    strFilterTemp = " (b.�����ı� ='����' And Instr(c.�����ı�, [23]) > 0)"
                Else
                    strFilterTemp = strFilterTemp & " or (b.�����ı� ='����' And Instr(c.�����ı�, [23]) > 0)"
                End If
            End If
            
            If strFilterTemp <> "" Then
                strFilterTemp = " (" & strFilterTemp & ")"
                strFilterReportAdvice = "( Select t.ҽ��id From ����ҽ������ t Where t.����id IN " _
                    & " (Select Distinct a.ID From ���Ӳ�����¼ a, ���Ӳ������� b,���Ӳ������� c " _
                    & " Where a.����ʱ�� > [10] And a.Id = b.�ļ�id And b.Id = C.��ID And b.�������� = 3 And c.�������� = 2 And c.��ֹ�� = 0 and " _
                    & strFilterTemp & "))"
            End If
        End If
    
        If mSysPar.blnNoShowCancel Then '����ʾȡ���Ǽǵļ��
            strFilterOrder = strFilterOrder & " And A.ִ��״̬<>2 "
        End If
        
        If mblncmd���� Then        'ֻ��ʾ����סԺ��¼
            strFilterOrder = strFilterOrder & vbNewLine & " And (B.������Դ=2 And B.��ҳID=C.סԺ���� Or Nvl(B.������Դ,0)<>2)"
        End If
        
        '�Ƿ�ѡ����ȫ������
        If mblnAllDepts = True Then
            strFilterOrder = strFilterOrder & " And Instr( [25],A.ִ�в���ID ) >0"
            
            If lngType = 2 Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & "  Instr([25],H.ִ�п���ID) >0 "
            End If
        Else
            strFilterOrder = strFilterOrder & " And A.ִ�в���ID+0=[24]"
            
            If lngType = 2 Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & "  H.ִ�п���ID+0=[24] "
            End If
        End If
        
        '������������
        If .�������� <> "" Then
            strFilterReportContext = "( Select t.ҽ��id From ����ҽ������ t Where t.����id In " & _
                                                                    " (Select Distinct A.ID " & _
                                                                    " From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                    " Where A.����ʱ��>[10] AND A.Id=B.�ļ�ID " & _
                                                                    " And B.��������=2 And instr(B.�����ı�,[26])>0 And B.��ֹ�� = 0)) "
        End If
        
        
        '�����ѯ���
        
        'ҽ������Ӳ�ѯ��
        strWithOrderCols = "A.ҽ��ID,A.���ͺ�,A.�״�ʱ��,A.����ʱ��,A.ִ��״̬,A.ִ�й���,nvl(A.ִ�м�,' ') as ִ�м�, A.������� , " & _
                            " A.NO, A.������, A.ִ�в���ID, A.��¼����, A.�Ʒ�״̬, A.����ʱ��, " & _
                            " B.ID, B.���ID,B.��ҳID,B.����ID, B.���˿���ID,B.�Һŵ�,B.������Դ,B.ҽ������,B.�걾��λ, " & _
                            " B.������־,B.Ӥ��,B.����ҽ��,B.����,B.�Ա�,B.����,B.������ĿID,F.Ӱ�����, " & _
                            " C.���￨��, C.���֤��, C.����ģʽ, decode(B.������Դ,2,D.��������,C.��������) as ��������, C.סԺ��, C.�����, C.��ǰ����, C.��ǰ����ID, D.·��״̬, E.���� "
    
        '�������Ӳ�ѯ��
        strWithStudyCols = "H.ҽ��ID, H.����,H.����,H.�Ա�,H.����,H.���,H.����,H.Ӱ������,H.��������,H.�������,H.�Ƿ�ʦȷ��," & _
                            " H.�����,H.�Ƿ���ӽ�Ƭ,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.������,H.��鼼ʦ,H.��鼼ʦ��,H.��������,H.ͼ��λ��, " & _
                            " H.�������,H.��Ϸ���,H.���UID,H.���ͺ�,H.����ID,H.������, H.���淢��,H.���Ž�Ƭ,H.Σ��״̬,J.ҽ��ID as ���뵥ҽ�� " & _
                            IIf(strPatholCol = "", "", "," & strPatholCol & ",p.id as ����ID,p.��ǰ״̬ as ����״̬,p.����ҽʦ" & _
                            ", (select count(1) from ��������Ϣ V , ����������Ϣ W where V.����ҽ��ID=w.����ҽ��id and v.ҽ��id=H.ҽ��ID and w.����״̬=1) as ���� ")
    
    
        strSql = ""
        strSubSql = ""
        
        strWithOrderTab = ""
        strWithStudyTab = ""
        
        If lngType = 1 Then
            '��ҽ����ѯΪ��
            
            '����ɾ���ò�ѯ�еġ�Ӱ������Ŀ�����ݱ���Ϊɾ���󣬻�������ݵĲ�ѯЧ�ʽϵͣ�ɾ��������Ҫʹ�ò���ҽ�����͵�ִ�в���ID��Ϊ�������˼�飬Ȼ�����ֶ�û��������
            strWithOrderTab = "tmpOrder as(select " & strWithOrderCols & vbNewLine & _
                              " from ����ҽ������  A, ����ҽ����¼ B,������Ϣ C, ������ҳ D,���ű� E,Ӱ������Ŀ F " & vbNewLine & _
                              " Where a.ҽ��ID = b.ID And b.����ID = c.����ID " & vbNewLine & _
                                      "     And B.���˿���ID=E.ID " & vbNewLine & _
                                      "     And B.����ID = D.����ID(+) And B.��ҳID+0 = D.��ҳID(+) And B.������ĿID+0 =F.������ĿID " & vbNewLine & _
                                      "     " & IIf(strFilterOrder = "", " ", strFilterOrder) & vbNewLine & _
                                                IIf(strFilterDate = "", "", "     And " & strFilterDate) & "  and B.ҽ��״̬ <> 4" & ")"
                                                      
'            strWithStudyTab = "tmpStudy as(select " & strWithStudyCols & vbNewLine & _
'                                " from Ӱ�����¼ H " & IIf(strPatholTab = "", "", " ," & strPatholTab) & " , tmpOrder I,Ӱ�����뵥ͼ�� J " & vbNewLine & _
'                                " where h.ҽ��ID=I.ҽ��ID  and H.ҽ��ID=J.ҽ��ID(+) " & vbNewLine & _
'                                IIf(strPatholFilter = "", "", "     And " & strPatholFilter) & _
'                                IIf(strFilterStudy = "", "", "     And " & strFilterStudy) & ")"
                             
                             
            '��ѯ������ҽ���Ͳ�λҽ����ҽ��id����
            If strFilterStudy <> "" Then
                strSubSql = "(select id from tmpOrder I, tmpStudy J where I.���Id=J.ҽ��Id " & _
                                " Union All " & _
                                " select I.ҽ��Id as id from tmpStudy I) K "
            End If
            
            strSql = " with " & strWithOrderTab '& "," & vbNewLine & strWithStudyTab
            
            '�������with���ѯ��oracle10g�´���һ��bug��������һ��with��ѯ��ʹ�������with��ѯ���ݱ�ʱ��������α���������Լ�ռ���ڴ����
            
            strSql = strSql & vbNewLine & _
                        " select distinct L.ҽ��ID,L.���ID,L.���ͺ�,L.�״�ʱ�� ����ʱ��,L.����ʱ�� ����ʱ��,L.ִ��״̬,nvl(L.ִ�й���,0) ������,L.ִ�м�,L.������� ����,M.Σ��״̬ Σ��," & vbNewLine & _
                        "     L.����ID,L.��ҳID,L.�Һŵ�,L.���˿���ID,L.������Դ ��Դ,L.ҽ������,L.�걾��λ," & vbNewLine & _
                        "     Nvl(L.������־, 0) ������־, Nvl(L.Ӥ��, 0) Ӥ��,L.����ҽ��,L.NO,L.��ǰ����,L.��ǰ����ID,Decode(L.������Դ,2,L.סԺ��,L.�����) ��ʶ��," & vbNewLine & _
                        "     Nvl(L.����,M.����) ����,L.Ӱ�����,M.����,Nvl(L.�Ա�,M.�Ա�) �Ա�,Nvl(L.����,M.����) ����,M.���,M.����,M.Ӱ������,M.��������,M.�������," & vbNewLine & _
                        "     Decode(L.������Դ,3,L.����ҽ��,L.������) �Ǽ���,M.������, M.���淢��,M.���Ž�Ƭ,M.����ID,L.��¼����, " & vbNewLine & _
                        "     M.�����,M.�Ƿ���ӽ�Ƭ,M.�Ƿ��ӡ,M.�������,M.��ɫͨ��,M.�����ӡ,M.������,M.������,M.�Ƿ�ʦȷ��,M.��鼼ʦ,M.��鼼ʦ��,M.�������� ��ͼʱ��," & vbNewLine & _
                        "     M.�������,M.��Ϸ���,M.���UID,M.ͼ��λ��,L.ִ�в���ID as ִ�п���ID,0 as ת��,L.���� AS ���˿���, L.����ʱ��, " & vbNewLine & _
                        "     L.���￨��,L.NO as ���ݺ�,L.���֤��,L.����ģʽ,L.��������,L.·��״̬,L.�Ʒ�״̬,Decode(L.��¼����,2,1,Decode(L.�Ʒ�״̬,3,1,0)) as �շ� ,f.ҽ��ID As ���뵥ҽ�� " & vbNewLine & _
                                IIf(strPatholCol = "", "", _
                                        "    ,M.����ID,M.����״̬,M.����ҽʦ,nvl(M.����,0) as ����, '' as ����ִ��״̬,decode(M.�����,null,'δ����','�Ѻ���') as �������," & _
                                        "    decode(M.�������,0,'����',1,'����',2,'ϸ��',3,'����',4,'ʬ��',5,'����ʯ��',null) as  ������, " & _
                                        Replace(strPatholCol, "o.", "M.") & vbNewLine) & _
                        " From tmpOrder L,Ӱ�����¼ M, Ӱ�����뵥ͼ�� f " & vbNewLine & _
                                IIf(strSubSql = "", "", "," & strSubSql) & vbNewLine & _
                        " Where L.Id=M.ҽ��ID(+) And m.ҽ��ID=f.ҽ��ID(+) " & IIf(strSubSql = "", "", " And L.Id=K.Id" & vbNewLine) & _
                                IIf(strPatholFilter = "", "", "     And " & strPatholFilter) & _
                                IIf(strFilterStudy = "", "", "     And " & strFilterStudy) & _
                                IIf(strFilterIllnessDiagnose = "", "", "     And L.Id In" & strFilterIllnessDiagnose & vbNewLine) & _
                                IIf(strFilterReportAdvice = "", "", "     And L.Id In" & strFilterReportAdvice & vbNewLine) & _
                                IIf(strFilterReportContext = "", "", "     And L.Id In" & strFilterReportContext & vbNewLine)
            
            
            
        Else
            '�Լ���ѯΪ��
            strWithStudyTab = "tmpStudy as(select " & strWithStudyCols & vbNewLine & _
                                " from Ӱ�����¼ H " & IIf(strPatholTab = "", "", " ," & strPatholTab) & ",Ӱ�����뵥ͼ�� J " & vbNewLine & _
                                " where H.ҽ��ID=J.ҽ��ID(+) " & vbNewLine & _
                                IIf(strPatholFilter = "", "", "     And " & strPatholFilter) & vbNewLine & _
                                IIf(strFilterStudy = "", "", "     And " & strFilterStudy) & vbNewLine & _
                                IIf(strFilterDate = "", "", "     And " & strFilterDate) & ")"
                                
            strWithOrderTab = "tmpOrder as(select " & strWithOrderCols & vbNewLine & _
                              " from ����ҽ������  A, ����ҽ����¼ B,������Ϣ C, tmpStudy G, ������ҳ D,���ű� E,Ӱ������Ŀ F " & vbNewLine & _
                              " Where a.ҽ��ID = b.ID And b.����ID = c.����ID " & vbNewLine & _
                                      "     And B.���˿���ID=E.ID " & vbNewLine & _
                                      "     And (B.ID=G.ҽ��ID) " & vbNewLine & _
                                      "     And B.����ID = D.����ID(+) And B.��ҳID+0 = D.��ҳID(+) And B.������ĿID+0 =F.������ĿID " & vbNewLine & _
                                      "     " & strFilterOrder & " and B.ҽ��״̬ <> 4" & vbNewLine & _
                                      " Union All " & vbNewLine & _
                                " select " & strWithOrderCols & vbNewLine & _
                              " from ����ҽ������  A, ����ҽ����¼ B,������Ϣ C, tmpStudy G, ������ҳ D,���ű� E,Ӱ������Ŀ F " & vbNewLine & _
                              " Where a.ҽ��ID = b.ID And b.����ID = c.����ID " & vbNewLine & _
                                      "     And B.���˿���ID=E.ID " & vbNewLine & _
                                      "     And (B.���ID=G.ҽ��ID) " & vbNewLine & _
                                      "     And B.����ID = D.����ID(+) And B.��ҳID+0 = D.��ҳID(+) And B.������ĿID+0 =F.������ĿID " & vbNewLine & _
                                      "     " & strFilterOrder & " and B.ҽ��״̬ <> 4" & vbNewLine & _
                                      ")"
                                                     
            strSql = " with " & strWithStudyTab & "," & vbNewLine & strWithOrderTab
                                             
        
        
            strSql = strSql & vbNewLine & _
                        " select distinct L.ҽ��ID,L.���ID,L.���ͺ�,L.�״�ʱ�� ����ʱ��,L.����ʱ�� ����ʱ��,L.ִ��״̬,nvl(L.ִ�й���,0) ������,L.ִ�м�,L.������� ����,M.Σ��״̬ Σ��," & vbNewLine & _
                        "     L.����ID,L.��ҳID,L.�Һŵ�,L.���˿���ID,L.������Դ ��Դ,L.ҽ������,L.�걾��λ," & vbNewLine & _
                        "     Nvl(L.������־, 0) ������־, Nvl(L.Ӥ��, 0) Ӥ��,L.����ҽ��,L.NO,L.��ǰ����,L.��ǰ����ID,Decode(L.������Դ,2,L.סԺ��,L.�����) ��ʶ��," & vbNewLine & _
                        "     Nvl(L.����,M.����) ����,L.Ӱ�����,M.����,Nvl(L.�Ա�,M.�Ա�) �Ա�,Nvl(L.����,M.����) ����,M.���,M.����,M.Ӱ������,M.��������,M.�������," & vbNewLine & _
                        "     Decode(L.������Դ,3,L.����ҽ��,L.������) �Ǽ���,M.������, M.���淢��,M.���Ž�Ƭ,M.����ID,L.��¼����, " & vbNewLine & _
                        "     M.�����,M.�Ƿ���ӽ�Ƭ,M.�Ƿ��ӡ,M.�������,M.��ɫͨ��,M.�����ӡ,M.������,M.������,M.�Ƿ�ʦȷ��,M.��鼼ʦ,M.��鼼ʦ��,M.�������� ��ͼʱ��," & vbNewLine & _
                        "     M.�������,M.��Ϸ���,M.���UID,M.ͼ��λ��,L.ִ�в���ID as ִ�п���ID,0 as ת��,L.���� AS ���˿���, L.����ʱ��, " & vbNewLine & _
                        "     L.���￨��,L.NO as ���ݺ�,L.���֤��,L.����ģʽ,L.��������,L.·��״̬,L.�Ʒ�״̬,Decode(L.��¼����,2,1,Decode(L.�Ʒ�״̬,3,1,0)) as �շ� ,M.���뵥ҽ�� " & vbNewLine & _
                                IIf(strPatholCol = "", "", _
                                        "    ,M.����ID,M.����״̬,M.����ҽʦ,nvl(M.����,0) as ����, '' as ����ִ��״̬,decode(M.�����,null,'δ����','�Ѻ���') as �������," & _
                                        "    decode(M.�������,0,'����',1,'����',2,'ϸ��',3,'����',4,'ʬ��',5,'����ʯ��',null) as  ������, " & _
                                        Replace(strPatholCol, "o.", "M.") & vbNewLine) & _
                        " From tmpOrder L,tmpStudy M " & vbNewLine & _
                                IIf(strSubSql = "", "", "," & strSubSql) & vbNewLine & _
                        " Where L.Id=M.ҽ��ID(+) " & IIf(strSubSql = "", "", " And L.Id=K.Id" & vbNewLine) & _
                                IIf(strFilterIllnessDiagnose = "", "", "     And L.Id In" & strFilterIllnessDiagnose & vbNewLine) & _
                                IIf(strFilterReportAdvice = "", "", "     And L.Id In" & strFilterReportAdvice & vbNewLine) & _
                                IIf(strFilterReportContext = "", "", "     And L.Id In" & strFilterReportContext & vbNewLine)

        End If
        
        '���������ת����Ҫ�����󱸱�
        If mblnMoved Then
            strSql = "Select /*+ RULE*/ * From (" & vbNewLine & strSql & vbNewLine & ")"
            strSQLBak = strSql
            strSQLBak = GetHistoryQuerySql(strSQLBak)
            
            strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
            strSql = strSql & " Union ALL " & strSQLBak
            
            strSql = "Select * From (" & vbNewLine & strSql & vbNewLine & ") Order by ������,����ʱ��,����ʱ��"
        Else
            strSql = "Select /*+ RULE*/ * From (" & vbNewLine & strSql & vbNewLine & ") Order by ������,����ʱ��,����ʱ��"
        End If
        
        '1: �����    2: סԺ��    3: ���￨��    4: ����    5: ���֤��    6: IC��    7: ���ݺ�    8: ������
        '9: ����    10: ��ʼʱ��    11: ����ʱ��    12: ���˿���ID    13: ҽ������    14: ������    15: ����    16: Ӱ������
        '17: ��鼼ʦ    18: Ӱ�����    19: �������    20: �����ı�-�������    21: �����ı�-�������    22: �����ı�-������    23: �����ı� -����
        '24: ִ�в���Id    25: ��ǰ��������Ids    26: ��������    27: �Ա�    28: ��ʼ����    29: ��������    30: �������    31: ����ID
        Set GetFilterData = GetDataToLocal(strSql, "��ȡ�����б�", .�����, .סԺ��, .���￨, .����, .���֤, _
                                            .IC��, .���ݺ�, .������, .����, .��ʼʱ��, .����ʱ��, .���˿���, .�걾��λ, _
                                            .���ҽ��, .���ҽ��, .Ӱ������, .��鼼ʦ, .Ӱ�����, .���, _
                                            .�������, .�������, .������, .����, mlngCur����ID, _
                                            mstrCanUse����IDs, .��������, .�Ա�, .��ʼ����, .��������, .�������, .����ID)
    End With
End Function


Private Function GetFilterWhere() As String
    Dim objControl As CommandBarControl
    Dim strFilter As String
    Dim strModalitys As String
    Dim lngCurExecuteState As Long
    Dim i As Long
    
    strFilter = ""
        
    '���˼�����
    If mlngModule <> G_LNG_PATHOLSYS_NUM And mintcmdӰ����� <> 0 Then
        'Ӱ������������ط�������������ѡ�񣬹��˴��ں����������棬���������е�Ϊ��
        Set objControl = cbrdock.FindControl(, ID_Ӱ�����)
        For i = 1 To objControl.CommandBar.Controls.Count
            If objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Checked = False Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & " Ӱ�����<>'" & objControl.CommandBar.FindControl(, ID_Ӱ����� + i).DescriptionText & "'"
            End If
        Next i
    End If

    '���˼��ִ�м�
    If mlngModule <> G_LNG_PATHOLSYS_NUM And mintcmdӰ��ִ�м� <> 0 Then
        Set objControl = cbrdock.FindControl(, ID_Ӱ��ִ�м�)
        For i = 1 To objControl.CommandBar.Controls.Count
            If objControl.CommandBar.FindControl(, ID_Ӱ��ִ�м� + i).Checked = False Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & " ִ�м�<>'" & objControl.CommandBar.FindControl(, ID_Ӱ��ִ�м� + i).DescriptionText & "'"
            End If
        Next i
    End If

    '���˲�����Դ
    If (Abs(mblncmd����) + Abs(mblncmdסԺ) + Abs(mblncmd���) + Abs(mblncmd����)) Mod 4 <> 0 Then
        If Not mblncmd���� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " ��Դ<>1"
        End If
        
        If Not mblncmdסԺ Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " ��Դ<>2"
        End If
        
        If Not mblncmd��� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " ��Դ<>4"
        End If
        
        If Not mblncmd���� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " ��Դ<>3"
        End If
    End If


    '�����̹���
    If (Abs(mblncmd�Ǽ�) + Abs(mblncmd����) + Abs(mblncmd���) + Abs(mblncmd����) + Abs(mblncmd���) + Abs(mblncmd����) + Abs(mblncmd���)) Mod 7 <> 0 Then
        If Not mblncmd�Ǽ� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " ������<>0 and ������<>1"
        End If
        
        If Not mblncmd���� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "������<>2"
        End If
        
        If Not mblncmd��� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "������<>3"
        End If
        
        If Not mblncmd���� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "������<>4"
        End If
        
        If Not mblncmd��� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "������<>5 "
        End If
        
        If Not mblncmd���� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "������<>-1 "
        End If
        
        If Not mblncmd��� Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "������<>6"
        End If
    End If


    '�Բ����֣���Ҫ��������һЩ���˴���
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        '���ȫѡ�����൱��ȫ��ѡ
        If (Abs(mblncmd����) + Abs(mblncmd����) + Abs(mblncmdϸ��) + Abs(mblncmd����) + Abs(mblncmdʬ��) + Abs(mblncmd����ʯ��)) Mod 6 <> 0 Then

            If Not mblncmd���� Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "������<>'����'"
            End If

            If Not mblncmd���� Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "������<>'����'"
            End If

            If Not mblncmdϸ�� Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "������<>'ϸ��'"
            End If

            If Not mblncmd���� Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "������<>'����'"
            End If

            If Not mblncmdʬ�� Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "������<>'ʬ��'"
            End If

            If Not mblncmd����ʯ�� Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "������<>'����ʯ��'"
            End If
            
        End If

        '���˵�ǰҳ������
        If tabFilter.tag Then

            lngCurExecuteState = GetExecuteState
            Select Case tabFilter.Selected.tag
                Case "ȡ��"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '��ȡ��
                        strFilter = strFilter & "ȡ�Ĺ��� = 1"
                    ElseIf lngCurExecuteState = 2 Then                      '��ȡ��
                        strFilter = strFilter & "ȡ�Ĺ��� = 2"
                    ElseIf lngCurExecuteState = 3 Then                      '����
                        strFilter = strFilter & "ȡ�Ĺ��� > 0"
                    End If

                Case "��Ƭ"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '����Ƭ
                        strFilter = strFilter & "��Ƭ���� = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '��Ƭ����
                        strFilter = strFilter & "��Ƭ���� = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '����Ƭ
                        strFilter = strFilter & "��Ƭ���� = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '����
                        strFilter = strFilter & "��Ƭ���� > 0"
                    End If

                Case "����"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '������
                        strFilter = strFilter & "���߹��� = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '���߽���
                        strFilter = strFilter & "���߹��� = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '������
                        strFilter = strFilter & "���߹��� = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '����
                        strFilter = strFilter & "���߹��� > 0"
                    End If

                Case "��Ⱦ"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '����Ⱦ
                        strFilter = strFilter & "��Ⱦ���� = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '��Ⱦ����
                        strFilter = strFilter & "��Ⱦ���� = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '����Ⱦ
                        strFilter = strFilter & "��Ⱦ���� = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '����
                        strFilter = strFilter & "��Ⱦ���� > 0"
                    End If


                Case "����"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '�����
                        strFilter = strFilter & "���ӹ��� = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '���ӽ���
                        strFilter = strFilter & "���ӹ��� = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '�ѷ���
                        strFilter = strFilter & "���ӹ��� = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '����
                        strFilter = strFilter & "���ӹ��� > 0"
                    End If

                Case "����"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '�����
                        strFilter = strFilter & "����״̬=0 and ����ҽʦ='" & UserInfo.���� & "'"
                    ElseIf lngCurExecuteState = 2 Then                      '�ѻ���
                        strFilter = strFilter & "����״̬<>0 and ����ҽʦ='" & UserInfo.���� & "'"
                    ElseIf lngCurExecuteState = 3 Then                      '����
                        strFilter = strFilter & " ����ID > 0 and ����ҽʦ='" & UserInfo.���� & "'"
                    End If

                Case "����"
            End Select
        End If
    End If
        
    GetFilterWhere = strFilter
End Function


Private Sub LoadPatiList(Optional ByVal blnFromDB As Boolean = True)
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
    Dim rsList As ADODB.Recordset

    If Not mblnInitOk Then Exit Sub      '��ʼ��δ���
    
    mblnvsRefresh = True
    
    If blnFromDB Then
        If Not mblnIsIntegratedQuery Then
            If mlngSysQuerySchemeId > 0 Then    'ʹ���Զ���ϵͳ��ѯ����
                Call ExecuteCustomQuery(mlngSysQuerySchemeId)
                mblnvsRefresh = False
                Exit Sub
            End If
        Else
            mblnIsIntegratedQuery = False
        End If
        
        Set rsList = GetFilterData()
        Set ufgStudyList.AdoData = rsList
    End If
    
    ufgStudyList.AdoFilter = GetFilterWhere
    
    '��binddata�ķ�����ʹ��refreshdata�ķ�����
    Call ufgStudyList.BindData
    
    '�ָ�����
    Call ufgStudyList.ResetSort(mlngSortCol, mintSortOrder)
    
    Call RefreshStatusBarInf
    
    mblnvsRefresh = False
End Sub


Private Sub picLoadState_Resize()
On Error GoTo ErrHandle
    labLoadState.Left = Fix((picLoadState.Width - labLoadState.Width) / 2)
    labLoadState.Top = Fix((picLoadState.Height - labLoadState.Height) / 2)
    
    picSmile.Left = labLoadState.Left - picSmile.Width
    picSmile.Top = labLoadState.Top - 80
    
ErrHandle:
End Sub

Private Sub picReportContainer_Resize()
On Error GoTo ErrHandle
    
    If mobjWork_Report Is Nothing Then Exit Sub
    
    Call mobjWork_Report.UpdateSize
    
ErrHandle:
End Sub



Private Sub picWindow_Resize()
On Error GoTo ErrHandle
    With TabWindow
        If GetWorkModuleCount = 1 Then
            TabWindow.PaintManager.ClientMargin.Top = -30
        Else
            TabWindow.PaintManager.ClientMargin.Top = 0
        End If
        
        .Left = 0
        .Width = picWindow.ScaleWidth
        .Height = picWindow.ScaleHeight + IIf(GetWorkModuleCount = 1, ScaleY(30, vbTwips, vbPixels), 0)
    End With
    
    tcDisable.Left = 0
    tcDisable.Top = IIf(TabWindow.PaintManager.ClientMargin.Top < 0, 0, IIf(mbytFontSize = 9, 440, 470))
    tcDisable.Width = picWindow.ScaleWidth
    tcDisable.Height = picWindow.ScaleHeight - IIf(TabWindow.PaintManager.ClientMargin.Top < 0, 0, IIf(mbytFontSize = 9, 440, 470))
ErrHandle:
End Sub

Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo ErrHandle
    If Not mblnInitOk Then Exit Sub
    
    If tabFilter.ItemCount < 7 Then Exit Sub
    If Not ufgStudyList.Visible Then Exit Sub
    
    optAccept.Enabled = IIf(Item.tag = "ȡ��" Or Item.tag = "����" Or Item.tag = "����", False, True)
    
    optNeed.Enabled = IIf(Item.tag = "����", False, True)
    optFinal.Enabled = IIf(Item.tag = "����", False, True)
    optAll.Enabled = IIf(Item.tag = "����", False, True)
    
    If (Item.tag = "ȡ��" Or Item.tag = "����") And optAccept.value Then
        '��checkֵ���ı�ʱ���ᴥ���ؼ���click�¼���ִ��RefreshList����
        optNeed.value = True
    Else
        Call RefreshList(, False)
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ConfigSubForm(ByVal Item As XtremeSuiteControls.ITabControlItem)
'�����Ӵ��ڽ���
On Error GoTo ErrHandle
    Dim lngIndex As Integer
    Dim objItem As XtremeSuiteControls.TabControlItem
    
    If mblnLoadSubFrom Then Exit Sub
    If Item.Handle <> picTemp.hWnd Then Exit Sub
    
    mblnLoadSubFrom = True
    lngIndex = Item.Index
    
    Set objItem = Nothing
    
    Select Case Item.tag
        Case "Ӱ��ͼ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "Ӱ���¼", mfrmWork_PacsImg.hWnd, Item.Image)
                
        Case "�걾����"
            Set objItem = TabWindow.InsertItem(lngIndex, "�걾����", mobjWork_Pathol.GetModule(mtSpecimen).hWnd, Item.Image)

        Case "����ȡ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "����ȡ��", mobjWork_Pathol.GetModule(mtMaterial).hWnd, Item.Image)
            
        Case "������Ƭ"
            Set objItem = TabWindow.InsertItem(lngIndex, "������Ƭ", mobjWork_Pathol.GetModule(mtSlices).hWnd, Item.Image)
            
        Case "�����ؼ�"
            Set objItem = TabWindow.InsertItem(lngIndex, "�����ؼ�", mobjWork_Pathol.GetModule(mtSpeExam).hWnd, Item.Image)
        
        Case "���̱���"
            Set objItem = TabWindow.InsertItem(lngIndex, "����/�ؼ챨��", mobjWork_Pathol.GetModule(mtProRep).hWnd, Item.Image)
            
        Case "�������"
            Set objItem = TabWindow.InsertItem(lngIndex, "���ü�¼", mobjWork_His.GetModule(hmExpense).hWnd, Item.Image)
            
        Case "סԺҽ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "ҽ����¼", mobjWork_His.GetModule(hmInAdvice).hWnd, Item.Image)
            
        Case "����ҽ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "ҽ����¼", mobjWork_His.GetModule(hmOutAdvices).hWnd, Item.Image)
            
        Case "סԺ����"
            Set objItem = TabWindow.InsertItem(lngIndex, "������¼", mobjWork_His.GetModule(hmInEPRs).hWnd, Item.Image)
            
        Case "���ﲡ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "������¼", mobjWork_His.GetModule(hmOutEPRs).hWnd, Item.Image)
           
        Case "������Ӳ���", "סԺ���Ӳ���"
            Set objItem = TabWindow.InsertItem(lngIndex, "���Ӳ���", mobjWork_His.GetModule(hmEMR).hWnd, Item.Image)
              
        Case "�Ŷӽк�"
            Set objItem = TabWindow.InsertItem(lngIndex, "�Ŷӽк�", mobjQueue.hWnd, Item.Image)
            
        Case "Ӱ��ɼ�", "������д"
            '���ﲻ���д���
    End Select
    
    Call RefreshModuleAdviceInf
    
    If Not objItem Is Nothing Then
        objItem.tag = Item.tag
        objItem.Selected = True
        
        Call TabWindow.RemoveItem(lngIndex + 1)
    End If
    
    mblnLoadSubFrom = False
Exit Sub
ErrHandle:
    If Not objItem Is Nothing Then
        If objItem.tag = "" Then
            Call TabWindow.RemoveItem(objItem.Index)
        End If
    End If
    
    mblnLoadSubFrom = False
End Sub

Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo ErrHandle
    Dim intStyle As Integer
    Dim blnVisible As Boolean
    Dim blnLargeIcon As Boolean
    Dim cbrControl As CommandBarControl

    
    Call ConfigSubForm(Item)

    If Not mblnInitOk Then Exit Sub
    
    Call ReSetModuleFontSize(mbytFontSize, IIf(mbytFontSize = 9, 0, 1))
    
    Call RefreshTabWindow
    
    'ˢ���Ŷӽк�ģ�����ݣ�����Ѿ����ò�����ѡ����Ŷӽк�ҳ��
    If Trim(Item.tag) = "�Ŷӽк�" Then
        Call RefreshPacsQueueData
    End If

    Call LockWindowUpdate(Me.hWnd)

    '�еĲ˵���ֻ�ڹ���ģ����ʾ��ʱ�� ����ʾ
    Call CreateWorkModuleMenu
    
    If Val(ufgStudyList.CurKeyValue) <> 0 Then
        '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If
    
    Call LockWindowUpdate(0)
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub GetRGB(ByVal lngColor As Long, lngR As Long, lngG As Long, lngB As Long)
    Dim lngMinVal As Long
    Dim lngMaxVal As Long
    
    lngMinVal = 80
    lngMaxVal = 225
    
    lngR = lngColor Mod 256
    
    If lngR <= lngMinVal Then
        lngR = lngMinVal
    ElseIf lngR > lngMaxVal Then
        lngR = lngMaxVal
    End If
    
    lngG = (Fix(lngColor \ 256)) Mod 256
 
    If lngG <= lngMinVal Then
        lngG = lngMinVal
    ElseIf lngG > lngMaxVal Then
        lngG = lngMaxVal
    End If
    
    lngB = Fix(lngColor \ 256 \ 256)
 
    If lngB <= lngMinVal Then
        lngB = lngMinVal
    ElseIf lngB > lngMaxVal Then
        lngB = lngMaxVal
    End If
End Sub


Private Sub timerCapture_Timer()
On Error GoTo ErrHandle

    timerCapture.Enabled = False
    
    'ʹ���ȼ����вɼ�
    If GetKeyAliasEx(mCaptureMsg.lngVirtualKey) = mstrCaptureHot Then
        If mblnUseActivexCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
                Call mobjWork_ActiveVideo.zlCaptureImg
            End If
        End If
    'ʹ���ȼ����к�̨�ɼ�
    ElseIf GetKeyAliasEx(mCaptureMsg.lngVirtualKey) = mstrCaptureAfterHot Then
        If mblnUseActivexCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
                Call mobjWork_ActiveVideo.zlCaptureAfterImg
            End If
        End If
    
    'ʹ���ȼ����б�Ǹ���
    ElseIf GetKeyAliasEx(mCaptureMsg.lngVirtualKey) = mstrCaptureAfterTagHot Then
        If mblnUseActivexCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
                Call mobjWork_ActiveVideo.zlUpdateAfterCaptureInfo
            End If
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume

End Sub

Private Sub timerOperHint_Timer()
On Error GoTo ErrHandle
    Dim i As Long
    Dim strText As String
    Dim dtOper As Date
    Dim lngColor1 As Long
    Dim lngR As Long, lngG As Long, lngB As Long
    
    If Not (mSysPar.lngEnregAfterTimeLen > 0 Or mSysPar.lngCheckInAfterTimeLen > 0 _
        Or mSysPar.lngStudyAfterTimeLen > 0 Or mSysPar.lngReportAfterTimeLen > 0 Or mSysPar.lngAuditAfterTimeLen > 0) Then
        timerOperHint.Enabled = False
        Exit Sub
    End If
    
    If ufgStudyList.DataGrid.Rows <= 1 Then Exit Sub
    
    '1��ʾ��ɫ��˸ʱ��ʾΪ������ɫ����һ�����ɫ��-1��ʾ��ʾΪ������ɫ��ǳһ�����ɫ��0��ʾ��ʾ���õ���ɫ
    If timerOperHint.tag = "1" Then
        timerOperHint.tag = "-1"
    ElseIf timerOperHint.tag = "-1" Then
        timerOperHint.tag = "0"
    ElseIf timerOperHint.tag = "0" Then
        timerOperHint.tag = "1"
    End If
    
    For i = ufgStudyList.DataGrid.TopRow To ufgStudyList.DataGrid.BottomRow
    
        dtOper = IIf(Nvl(ufgStudyList.Text(i, "����ʱ��")) = "", Now, ufgStudyList.Text(i, "����ʱ��"))
        strText = ufgStudyList.Text(i, "������")
        
        Select Case strText
            Case "�ѵǼ�"
                If mSysPar.lngEnregAfterTimeLen > 0 Then
                    dtOper = Nvl(ufgStudyList.Text(i, "����ʱ��"))
                    
                    Call SetFlickerColor(i, gdblColor�ѵǼ�, dtOper, mSysPar.lngEnregAfterTimeLen)
                End If
            Case "�ѱ���"
                If mSysPar.lngCheckInAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor�ѱ���, dtOper, mSysPar.lngCheckInAfterTimeLen)
                End If
            Case "�Ѽ��"
                If mSysPar.lngStudyAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor�Ѽ��, dtOper, mSysPar.lngStudyAfterTimeLen)
                End If
            Case "�ѱ���"
                If mSysPar.lngReportAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor�ѱ���, dtOper, mSysPar.lngReportAfterTimeLen)
                End If
            Case "�����"
                If mSysPar.lngAuditAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor�����, dtOper, mSysPar.lngAuditAfterTimeLen)
                End If
        End Select
    Next i
ErrHandle:
End Sub

Private Sub SetFlickerColor(ByVal lngRow As Long, ByVal lngStateColor As Long, ByVal dtOper As Date, ByVal lngAfterTimeLen As Long)
'���ܣ������ѳ�ʱ�е���˸��ɫ
'������lngRow---��ǰ��
'      lngStateColor---�����õ���ɫ
    Dim lngR As Long, lngG As Long, lngB As Long
    Dim lngPreStateColor As Long
    Dim lngNextStateColor As Long
    
    Call GetRGB(lngStateColor, lngR, lngG, lngB)
    lngNextStateColor = RGB(lngR - 30, lngG - 30, lngB - 30)
    lngPreStateColor = RGB(lngR + 30, lngG + 30, lngB + 30)
    
    If DateDiff("N", dtOper, Now) >= lngAfterTimeLen Then
        If timerOperHint.tag = "1" Then
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = lngPreStateColor
        ElseIf timerOperHint.tag = "-1" Then
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = lngStateColor
        ElseIf timerOperHint.tag = "0" Then
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = lngNextStateColor
        End If
    End If
End Sub

Private Sub timerRefresh_Timer()
On Error GoTo ErrHandle
    'ˢ�²����б�
    If Not mblnInitOk Then Exit Sub
    If Not Me.Visible Then Exit Sub

    timerRefresh.Enabled = False
    
    Call RefreshList
    
    timerRefresh.Enabled = True
    
ErrHandle:
End Sub


Private Sub ChangeUser()
    Dim strPrivs As String
    Dim strUserID As String
    
    frmTwoUser.intDBState = mintChangeUserState
    frmTwoUser.strUserNameHIS = mstrUserNameHIS
    frmTwoUser.strUserIDHIS = mstrUserIDHIS
    frmTwoUser.Show 1, Me
    
    If frmTwoUser.blnOk = True Then
        If frmTwoUser.intDBState = 1 Then   'ͳһ����ָ���HISԭ�������ݿ����Ӻ��û���
            mstrUserNameNew = mstrUserNameHIS
            mstrUserIDNew = mstrUserIDHIS
            mblnCnOracleIsHIS = True
            mintChangeUserState = 1
            Set gcnOracle = mcnOracleHIS
            InitCommon gcnOracle
            SetDbUser mstrUserIDHIS
            RegCheck
            Call GetUserInfo
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
        ElseIf frmTwoUser.intDBState = 2 Then   '�������򽻻����ݿ�����
            '�����ʹ�������ݿ����ӣ��ȼ��Ȩ��
            mstrUserNameNew = frmTwoUser.strUserNameNew
            mstrUserIDNew = frmTwoUser.strUserIDNew
            mintChangeUserState = 2
            If frmTwoUser.blnCnOracleIsNew = True Then
                Set gcnOracle = frmTwoUser.cnOracle
                mblnCnOracleIsHIS = False
                
                '��ʼ��zlComLib������ȷ��GetPrivFunc��ȡ������ȷ����Ϣ
                InitCommon gcnOracle
                RegCheck
                SetDbUser mstrUserIDNew
                
                '�����û�Ȩ��
                strPrivs = GetPrivFunc(100, 1291)       'Ӱ��ɼ�����վ
                If strPrivs = "" Then
                    MsgBoxD Me, "�㲻�߱�ʹ�á�Ӱ��ɼ�����վ��ģ���Ȩ�ޣ�"
                    
                    '�л���ԭ�����û�
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
                    RegCheck
                    SetDbUser mstrUserIDHIS
                
                    mstrUserNameNew = mstrUserNameHIS
                    mstrUserIDNew = mstrUserIDHIS
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                End If
                
                strPrivs = GetPrivFunc(100, 1258)       '���Ʊ������
                If strPrivs = "" Then
                    MsgBoxD Me, "�㲻�߱�ʹ�á����Ʊ��桱ģ���Ȩ�ޣ�"
                    
                    '�л���ԭ�����û�
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
                    RegCheck
                    SetDbUser mstrUserIDHIS
                    
                    mstrUserNameNew = mstrUserNameHIS
                    mstrUserIDNew = mstrUserIDHIS
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                End If
            Else
                Set gcnOracle = mcnOracleHIS
                
                InitCommon gcnOracle
                RegCheck
                SetDbUser mstrUserIDHIS
                
                mblnCnOracleIsHIS = True
            End If
            
            Call GetUserInfo
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
        End If
    End If
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameHIS & "   ���ҽ����" & mstrUserNameNew
    Else
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameNew & "   ���ҽ����" & mstrUserNameHIS
    End If
End Sub

Private Sub SwitchUser()
    Dim strPrivs As String
    
    frmSwitchUser.Show 1, Me
    
    If frmSwitchUser.blnOk Then
        '�����ʹ�������ݿ����ӣ��ȼ��Ȩ��
        mstrUserNameNew = frmSwitchUser.strUserNameNew
        mstrUserIDNew = frmSwitchUser.strUserIDNew

        If frmSwitchUser.blnCnOracleIsNew = True Then
            Set gcnOracle = frmSwitchUser.cnOracle
            mblnCnOracleIsHIS = False
            
            '��ʼ��zlComLib������ȷ��GetPrivFunc��ȡ������ȷ����Ϣ
            zl9comlib.CloseWindows
            InitCommon gcnOracle
            RegCheck
            SetDbUser mstrUserIDNew
            
            '�����û�Ȩ��
            strPrivs = GetPrivFunc(100, 1294)       'Ӱ������վ
            If strPrivs = "" Then
                MsgBoxD Me, "�㲻�߱�ʹ�á�Ӱ������վ��ģ���Ȩ�ޣ�"
                
                '�л���ԭ�����û�
                Set gcnOracle = mcnOracleHIS
                
                InitCommon gcnOracle
                RegCheck
                SetDbUser mstrUserIDHIS
            
                mstrUserNameNew = mstrUserNameHIS
                mstrUserIDNew = mstrUserIDHIS
                mblnCnOracleIsHIS = True
            End If
        Else
            Set gcnOracle = mcnOracleHIS
            
            zl9comlib.CloseWindows
            InitCommon gcnOracle
            RegCheck
            SetDbUser mstrUserIDHIS
            
            mblnCnOracleIsHIS = True
        End If
    End If
    
    Call GetUserInfo
    Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameHIS & "   ���ҽ����" & mstrUserNameHIS
    Else
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameNew & "   ���ҽ����" & mstrUserNameNew
    End If
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strCardName As String, _
    ByVal strFilter As String, Optional blnIsReSeek As Boolean = False)
'------------------------------------------------
'���ܣ��ڲ����б��ж�λָ���ļ�¼
'������ blnFirst -- �Ƿ��һ�β���
'���أ��ޣ�ֱ���ڲ����б��ж�λ
'------------------------------------------------
    Dim i As Long
    Dim intB As Integer
    Dim lngEndRow As Long
    Dim lngSelRow As Long
    Dim strTemp As String
    Dim lngRowIndex As Long

    
    '���û�м�¼�����˳�
    If ufgStudyList.ShowingRowCount <= 0 Then Exit Sub

    intB = 0
    
    If Not blnFirst Then
        intB = ufgStudyList.DataGrid.Row + 1
        If intB >= ufgStudyList.DataGrid.Rows Then intB = 1
    End If
    
    lngSelRow = ufgStudyList.DataGrid.Row
    lngEndRow = ufgStudyList.DataGrid.Rows - 1

continue1:

    Select Case strCardName
        Case "��ʶ��", "סԺ��", "�����"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("��ʶ��"), False, False)
            
        Case "���ݺ�"
            strTemp = ""
            
            '��ȫ���ݺ�
            If Len(Trim(strFilter)) > 0 Then
                If Len(Trim(strFilter)) < 8 And Not IsNumeric(Trim(strFilter)) Then
                    strTemp = GetFullNO(0, 0)
                    strTemp = Mid(strTemp, 1, Len(strTemp) - Len(strFilter)) & strFilter
                Else
                    strTemp = GetFullNO(Nvl(strFilter, 0), 0)
                End If
            End If
            
            PatiIdentify.Text = strTemp
            
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strTemp, intB, ufgStudyList.GetColIndex("NO"), False, False)
            
        Case GetStudyNumberDisplayName
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex(GetStudyNumberDisplayName), False, False)
            
        Case "����", "�� ��", "��  ��", "��   ��"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("����"), False, False)
            
            '���û���ҵ������ж������Ƿ�Ϊȫ��ĸ������ǣ���ʹ��ƴ������
            If lngRowIndex <= 0 And LenB(StrConv(strFilter, vbFromUnicode)) = Len(strFilter) Then
                For i = intB To lngEndRow
                    If zlCommFun.SpellCode(Nvl(ufgStudyList.Text(i, "����"), "")) Like UCase(strFilter) & "*" Then
                        lngRowIndex = i
                        Exit For
                    End If
                Next i
            End If
            
        Case "���￨", "���￨��"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("���￨��"), False, False)
            
        Case "���֤��", "���֤"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("���֤��"), False, False)
        
        Case "ҽ��ID"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("ҽ��ID"), False, False)
            
        Case Else
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("����ID"), False, True)
            
    End Select


    If lngRowIndex > 0 Then
        PatiIdentify.tag = PatiIdentify.Text
        
        On Error GoTo errContinue1

            ufgStudyList.DataGrid.Row = lngRowIndex

            If ufgStudyList.DataGrid.TopRow > ufgStudyList.DataGrid.Row Then ufgStudyList.DataGrid.TopRow = ufgStudyList.DataGrid.Row
            If ufgStudyList.DataGrid.BottomRow - 1 < ufgStudyList.DataGrid.Row Then
                ufgStudyList.DataGrid.TopRow = ufgStudyList.DataGrid.TopRow + (ufgStudyList.DataGrid.Row - ufgStudyList.DataGrid.BottomRow) + 1
            End If

            If lngSelRow = ufgStudyList.DataGrid.Row Then
                '����ü��Ϊ�ѵǼ�״̬����ִ�б�������
                If ufgStudyList.CurText("������") = "�ѵǼ�" Then
                    Call Menu_Manage_����
                End If
            End If
        
errContinue1:
        
        Exit Sub
    End If
    
    '���û���ҵ�������ִ��ˢ���б�Ȼ���ٶ�λ����������ÿ�ζ�λ��Ҫˢ���б�
    If lngRowIndex <= 0 Then
        If blnIsReSeek Then
        
            Call RefreshList
            blnIsReSeek = False
            
            GoTo continue1
        
        End If
    End If
    
    If intB > 1 Then
        lngEndRow = intB
        intB = 1
        
        GoTo continue1
    End If
    
    ufgStudyList.DataGrid.Row = -1
End Sub

Private Sub Menu_Manage_���()
On Error GoTo ErrHandle
    Dim strReview As String
    Dim strDeptName As String

    If mListAdviceInf.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    strDeptName = Split(mstrCur����, "-")(1)
    If frmReview.ShowMe(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, Me, strDeptName, strReview) = True Then
        ufgStudyList.CurText("�������") = strReview
        Call ufgStudyList.UpdateSourceData(mListAdviceInf.lngAdviceId, "�������", strReview)
    End If

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_���淢��()
'���淢��
On Error GoTo ErrHandle
    Dim strSql As String

    If mListAdviceInf.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    With ufgStudyList
        If mrtReportType = �����ĵ��༭�� Then
            Call mobjWork_Report.Menu_Manage_���淢��(mListAdviceInf.lngAdviceId, IIf(mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) > 1, 0, 1))
        Else
            strSql = "Zl_Ӱ�񱨸淢��(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSql, "���淢��")
            
            .CurText("���淢��") = IIf(Nvl(.CurText("���淢��"), "0") = "0", "1", "0")
            Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���淢��", IIf(.CurText("���淢��") = "1", "1", "0"))
        End If
    End With
    
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_��Ƭ����()
'��Ƭ����
On Error GoTo ErrHandle
    Dim strSql As String

    With ufgStudyList

        If mListAdviceInf.lngAdviceId <= 0 Then
            MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
            Exit Sub
        End If
        
        strSql = "Zl_Ӱ��Ƭ����(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "��Ƭ����")
        
        .CurText("��Ƭ����") = IIf(Nvl(.CurText("��Ƭ����"), "0") = "0", "1", "0")
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���Ž�Ƭ", IIf(Nvl(.CurText("��Ƭ����"), "0") = "1", "1", "0"))
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_���潺Ƭͬʱ����()
'���潺Ƭͬʱ����
On Error GoTo ErrHandle
    Dim strSql As String
    
    With ufgStudyList
        
        If mListAdviceInf.lngAdviceId <= 0 Then
            MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
            Exit Sub
        End If
        
        If mrtReportType = �����ĵ��༭�� Then
            If mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) = 3 And Nvl(.CurText("��Ƭ����"), "0") = "1" Then
                Call mobjWork_Report.Menu_Manage_���淢��(mListAdviceInf.lngAdviceId, 0)
                
                strSql = "Zl_Ӱ��Ƭ����(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "��Ƭ����")
                
                .CurText("��Ƭ����") = "0"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���Ž�Ƭ", IIf(Nvl(.CurText("��Ƭ����"), "0") = "1", "1", "0"))
            Else
                Call mobjWork_Report.Menu_Manage_���淢��(mListAdviceInf.lngAdviceId, 1)
                
                strSql = "Zl_Ӱ��Ƭ����(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "��Ƭ����")
                
                .CurText("��Ƭ����") = "1"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���Ž�Ƭ", IIf(Nvl(.CurText("��Ƭ����"), "0") = "1", "1", "0"))
            End If
        Else
            If Nvl(.CurText("���淢��"), "0") = "1" And Nvl(.CurText("��Ƭ����"), "0") = "1" Then
                strSql = "Zl_Ӱ�񱨸淢��(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "���淢��")
                
                .CurText("���淢��") = "0"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���淢��", IIf(Nvl(.CurText("���淢��"), "0") = "1", "1", "0"))
            
                strSql = "Zl_Ӱ��Ƭ����(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "��Ƭ����")
                
                .CurText("��Ƭ����") = "0"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���Ž�Ƭ", IIf(Nvl(.CurText("��Ƭ����"), "0") = "1", "1", "0"))
            Else
                strSql = "Zl_Ӱ�񱨸淢��(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "���淢��")
                
                .CurText("���淢��") = "1"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���淢��", IIf(Nvl(.CurText("���淢��"), "0") = "1", "1", "0"))
            
                strSql = "Zl_Ӱ��Ƭ����(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "��Ƭ����")
                
                .CurText("��Ƭ����") = "1"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "���Ž�Ƭ", IIf(Nvl(.CurText("��Ƭ����"), "0") = "1", "1", "0"))
            End If
        End If
        
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_ReportExecutor()
    Dim strSql As String
    
    Dim strRPTExecutor As String
On Error GoTo ErrHandle
    strRPTExecutor = frmSelectRPTExecutor.GetRPTExecutor(mlngCur����ID, Me, mstrRPTExecutor)
    
    If strRPTExecutor <> "" Then
        '���±�����
        strSql = "ZL_Ӱ�񱨸汣��_���±�����(" & mcurAdviceInf.lngAdviceId & ",'" & strRPTExecutor & "')"
        Call zlDatabase.ExecuteProcedure(CStr(strSql), "���±�����")
        
        'ˢ�¶�Ӧ���ı�����
        mstrRPTExecutor = strRPTExecutor
        ufgStudyList.CurText("������") = strRPTExecutor
        
        If Not mobjWork_Report Is Nothing And mrtReportType = �����ĵ��༭�� Then Call mobjWork_Report.SetDocCreator(mstrRPTExecutor)
        
        stbThis.Panels(4).Text = "����ҽ����" & strRPTExecutor & "   ���ҽ����" & Split(stbThis.Panels(4).Text, "���ҽ����")(1)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub timerVideoEvent_Timer()
On Error GoTo ErrHandle
    timerVideoEvent.Enabled = False
    
    Call DoOnStateChange(mVideoEventInf.vetEventType, mVideoEventInf.lngAdviceId, mVideoEventInf.lngSendNo, mVideoEventInf.strOtherInf)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume

End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
On Error GoTo ErrHandle
    PatiIdentify.Text = ""  '�л�Itemʱ��Ҫ����������
    If cbrdock.FindControl(, ID_���ҷ�ʽ) Is Nothing Then Exit Sub
    '�ڿ��ٹ����������λ�Ͳ���ʱ������ˢ���ؼ�IDKindStrʱ�������ItemClick�¼���
    '�����޷��ֱ��¼��λ�Ͳ����ֶ���Ϣ�������ô˱�����ǣ�Ϊtrueʱ������mstrLocateWay��mstrFindWay
    If mblnAssignment Then Exit Sub
    
    If cbrdock.FindControl(, ID_���ҷ�ʽ).IconId = 3 Then
        mstrLocateWay = objCard.����
    Else
        mstrFindWay = objCard.����
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub StartReadCard()
'��ʼ����
    Dim lngPatientID As Long
    Dim strCurCardName As String
    
    If cbrdock.FindControl(, ID_���ҷ�ʽ).IconId = 3 Then
        strCurCardName = mstrLocateWay
    Else
        strCurCardName = mstrFindWay
    End If
    
    If PatiIdentify.GetCurCard.�Ƿ�ˢ�� Then
        Call OnFilterRead(strCurCardName, PatiIdentify.Text, IIf(lngPatientID > 0, lngPatientID, ""))
    Else
        If PatiIdentify.GetCurCard.�ӿ���� > 0 Then
            Call mobjSquareCard.zlGetPatiID(PatiIdentify.GetCurCard.�ӿ����, PatiIdentify.Text, , lngPatientID)
            
            Call OnFilterRead(strCurCardName, PatiIdentify.Text, IIf(lngPatientID > 0, lngPatientID, ""))
        Else
            Call OnFilterRead(strCurCardName, PatiIdentify.Text, "")
        End If
    End If
End Sub

Private Sub OnFilterRead(ByVal strCardName As String, ByVal strFilter As String, ByVal strPatientId As String)
'��ʼ��������
On Error GoTo ErrHandle
    If cbrdock.FindControl(, ID_���ҷ�ʽ).IconId = 3 Then
        '��λ�������
        If strPatientId <> "" Then
            Call SeekNextPati(PatiIdentify.tag <> PatiIdentify.Text, "����ID", strPatientId, True)
        Else
            Call SeekNextPati(PatiIdentify.tag <> PatiIdentify.Text, strCardName, strFilter, True)
        End If
    Else
        '���Ҽ������
        If strPatientId <> "" Then
            Call subRefreshFilterCondition("����ID", strPatientId)
        Else
            Call subRefreshFilterCondition(strCardName, strFilter)
        End If
        
        Call RefreshList
    End If
    
    Call PatiIdentify.SetFocus
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetStudyNumberDisplayName() As String
'��ȡ��������ʾ����
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "�����", "����")
End Function




Private Sub ufgStudyList_OnBindFilter(strBindFilter As String, strCloneFilter As String)
    strBindFilter = " ���ID=NULL"
    strCloneFilter = " ���ID<>NULL"
End Sub

Private Sub ufgStudyList_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    frmDegreeCard.ShowMe Val(ufgStudyList.Text(Row, "����ID")), Val(ufgStudyList.Text(Row, "��ҳID")), Me
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnColFormartChange()
On Error GoTo ErrHandle
    Call zlDatabase.SetPara("����б�", ufgStudyList.GetColsString(ufgStudyList), glngSys, mlngModule)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnColsNameReSet()
On Error GoTo ErrHandle
    '��ͷ�ָ�Ĭ�Ϻ����¼��ز����б�
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnDblClick()
On Error GoTo ErrHandle
    If Val(ufgStudyList.CurKeyValue) <> 0 Then
        '˫�����˼���б�ʱ��������˼��״̬Ϊ �Ѿܾ���Ŀǰ�����κδ���
        If Nvl(ufgStudyList.CurText("������")) = "�Ѿܾ�" Then Exit Sub
        
        Select Case Val(ufgStudyList.CurText("���״̬"))
            Case 1, 0
                Call Menu_Manage_����
            Case 2, 3               '˫������д����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Modify)
            Case -1, 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '����
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnFilterRowData(rsData As ADODB.Recordset, rsClone As ADODB.Recordset, blnFilterOut As Boolean)
On Error GoTo ErrHandle
    '�ж��Ƿ��Ѿ��շ�
    '"����ҽ������.��¼����"--- 1���շѵģ�2�Ǽ��ʵġ�
    
    'ͨ��"����ҽ������.�Ʒ�״̬"ֱ���ж�,ԭ��ֵ��-1-����Ʒ�;0-δ�Ʒ�;1-�ѼƷѣ����ڼ��ʵ�������������ʵ���������ԭ��ֵ���䡣
    '�����շѵ��ķ��ͼ�¼����������״̬��2-�����շѣ�3-ȫ���շ�
    
    'û�ж�Ӧ���õ�ҽ�������������һ����"-1-����Ʒ�"����û�������շѶ��գ�һ����"0-δ�Ʒ�"������Ȼ�������շѶ��գ�������Ϊ���ͺ��ֹ��Ʒѣ�����ҽ������ȥ���ɡ�
    '"1-�ѼƷ�"���Ƿ���ʱ�����˷��õġ��������˷��õ��ݲ���ʾ�շ��ˣ����ɿ����Ǽ��ʻ��۵������շѻ��۵��������շѻ��۵��Ͷ�����״̬��
    '"2-�����շ�"��ʾ�����շѺͲ����˷ѵ����������û�յ��ꡣ
    
    '���շ���ʾ״̬�����շѣ��޷��ã�δ�շѣ�
    'δ�շ�----
    '1����ҽ�����շѵ��ģ���������������δ�շ�
    '   (1)��һ����ҽ���Ͳ�λҽ���� �Ʒ�״̬ in (1,2)��δ�շ� ------����¼����=1 and �Ʒ�״̬ in (1,2)��
    '���շѣ�
    '1����ҽ���Ǽ��˵����շ�-------����¼����=2��
    '2����ҽ�����շѵ��ģ����������������շ�
    '   (1)�ų�δ�շѺ���һ����ҽ���Ͳ�λҽ���� �Ʒ�״̬ =3 ���շ�-----����¼����=1 and �Ʒ�״̬ = 3��
    '�޷���
    '1����ҽ�����շѵ��ģ����������������޷���
    '   (1)������ҽ���Ͳ�λҽ���� �Ʒ�״̬ in (-1,0)���޷��� ------����¼����=1 and �Ʒ�״̬ in (-1,0)��
    
    Dim intNum As Integer
    
    intNum = 0
    If Nvl(rsData!���ID) = "" Then '˵������ҽ��
        If mstrcmd��λ <> "" Then
            '���˳���ҽ��
            rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID)
            
            Do While rsClone.EOF = False
                '���˲�λû�й�ѡʱ����ҽ������Ҫ��ʾ���б���
                If Nvl(rsClone!�걾��λ) <> "" Then
                    If InStr(mstrcmd��λ, Nvl(rsClone!�걾��λ)) > 0 Then '��ҽ���в�λ�ڷ����У���ҽ����Ҫ��ʾ
                        intNum = 1
                        Exit Do
                    End If
                End If
                
                rsClone.MoveNext
            Loop
            
            If rsClone.RecordCount > 0 And intNum <= 0 Then '˵��������ҽ���Ĳ�λ��û��ѡ����ҽ������Ҫ��ʾ���б���
                blnFilterOut = True
                Exit Sub
            End If
        End If
    End If
    
    ' intCharged  '0--δ�շѣ�1--���շѣ�2--�޷���
    
    If Nvl(rsData!���ID) <> "" Then
        '���id��Ϊ��ʱ��˵���鲿λҽ��������Ҫ��ʾ���б���
        blnFilterOut = True
        Exit Sub
    End If

    mlngTempCharged = 2 '�޷���
    
    If Nvl(rsData!��¼����, 2) = 2 Then
        '�жϲ��˽���ģʽ��������˽���ģʽΪ1�����ʾ�����ƺ���㼴���˲��ˣ���ʱ����״̬��Ҫ��ʾΪ���ǡ�
        If Val(Nvl(rsData!����ģʽ)) = 1 Then
            mlngTempCharged = 5         '������˲�����ʾΪ���ǡ�
        Else
            'סԺ�ǼǵĲ��ˣ����û�мƷѣ����Ϊ�޷���
            If Nvl(rsData!�Ʒ�״̬, -1) = 0 Then
    
                    rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID)
                    Do While rsClone.EOF = False
                        If Nvl(rsClone!�Ʒ�״̬, -1) = 1 Then
                            '����Ǽ���ҽ���������ѼƷѺ�ȫ���շѵģ���ʾΪ���շ�
                            mlngTempCharged = 5      '�Ѽ���
                            
                        ElseIf Nvl(rsClone!�Ʒ�״̬, -1) = 3 Then
                            mlngTempCharged = 1      '���շ�
       
                        ElseIf Nvl(rsClone!�Ʒ�״̬, -1) = 2 Then
                            mlngTempCharged = 0  'δ�շ�
                            
                            Exit Do
                        End If
                        rsClone.MoveNext
                    Loop
                    
            Else
'                mlngTempCharged = 1  '���շ�
                mlngTempCharged = 5         'סԺ���˲�����ʾΪ���ǡ�
            End If
            
        End If
    Else
        If Nvl(rsData!�Ʒ�״̬, -1) = 1 Or Nvl(rsData!�Ʒ�״̬, -1) = 2 Then
            mlngTempCharged = 0      'δ�շ�
        Else        '��ҽ���ļƷ�״̬�� -1,0,3  ��3--���շѣ�-1��0--�޷��ã�
            '��ѯ��ҽ��δ�Ʒѻ����Ѿ��շ��ˣ���Ҫ�鲿λҽ�����շ����������ҽ�����Ѿ��շѣ��������շ�
            
            '��������������շѵģ��ȼ�¼�����շ�
            If Nvl(rsData!�Ʒ�״̬, -1) = 3 Then
                mlngTempCharged = 1      '���շ�
            End If
            
            rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID)
            Do While rsClone.EOF = False
                If Nvl(rsClone!�Ʒ�״̬, -1) = 1 Or Nvl(rsClone!�Ʒ�״̬, -1) = 2 Then
                    mlngTempCharged = 0      'δ�շ�

                    Exit Do
                ElseIf Nvl(rsClone!�Ʒ�״̬, -1) = 3 Then
                    mlngTempCharged = 1      '���շ�
                End If

                rsClone.MoveNext
            Loop
            
'            '�Ʒ�״̬��-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷѣ����շѵ��ݶ�����״̬:2-�����շѣ�3-ȫ���շ�
'            rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID) & " and �Ʒ�״̬=1 and �Ʒ�״̬=2"
'            If rsClone.RecordCount > 0 Then
'                mlngTempCharged = 0 'δ�շ�
'            Else
'                rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID) & " and �Ʒ�״̬=3"
'                If rsClone.RecordCount > 0 Then mlngTempCharged = 1 '���շ�
'            End If
            
        End If
    End If

    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If Nvl(rsData!����) > 0 Then mlngTempCharged = 4 '��Ҫ���ѣ��貹�ѵļ��Ҳ��δ�շѵļ��
    End If
    
    If Nvl(rsData!���ID) = "" And ((mblncmd�ѽ� = True And mlngTempCharged = 1) Or (mblncmdδ�� = True And (mlngTempCharged = 0 Or mlngTempCharged = 4)) _
        Or (mblncmd�޷� = True And mlngTempCharged = 2) Or (mblncmd���� = True And mlngTempCharged = 4) Or (mblncmd���� And mlngTempCharged = 5) _
        Or (mblncmd�ѽ� = False And mblncmdδ�� = False And mblncmd���� = False And mblncmd�޷� = False And mblncmd���� = False)) Then
        blnFilterOut = False
        
        Call RowDataConvert(rsData)
    Else
        blnFilterOut = True
    End If
ErrHandle:
End Sub



Private Sub RowDataConvert(rsData As ADODB.Recordset)
On Error Resume Next
    Dim rsBaby As ADODB.Recordset
    Dim intTxtLen As Long
    
    '���������Ҫ��ʾ������Ҫת�������еĲ���ֵ
    rsData!���뵥 = IIf(Nvl(rsData!���뵥) = "", "", "��ɨ��")
    If mrtReportType = �����ĵ��༭�� Then
        rsData!������ = IIf(Val(Nvl(rsData!ִ��״̬)) = 2, "�Ѿܾ�", Decode(Val(Nvl(rsData!���״̬, 0)), -1, "�Ѳ���", 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                                                2, "�ѱ���", 3, "�Ѽ��", 4, "�ѱ���", 5, "�����", "�����"))
    Else
        rsData!������ = IIf(Val(Nvl(rsData!ִ��״̬)) = 2, "�Ѿܾ�", Decode(Val(Nvl(rsData!���״̬, 0)), -1, "�Ѳ���", 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                                                2, IIf(Nvl(rsData!�������) <> "", "������", _
                                                                                        IIf(Nvl(rsData!������) = "", "�ѱ���", "������")), _
                                                                                3, IIf(Nvl(rsData!�������) <> "", "������", _
                                                                                        IIf(Nvl(rsData!������) = "", "�Ѽ��", "������")), _
                                                                                4, IIf(Nvl(rsData!�������) <> "", "������", _
                                                                                        IIf(Nvl(rsData!������) <> "", "�����", "�ѱ���")), _
                                                                                5, "�����", "�����"))
    End If
                                                                                
    If Nvl(rsData!Ӥ��) <> 0 Then
        gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                    "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                    "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
        
        Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", CLng(rsData!����ID), CLng(Nvl(rsData!��ҳID, 0)), CLng(rsData!Ӥ��))
        
        If Not rsBaby.EOF Then
            rsData!���� = rsBaby!Ӥ������
            rsData!�Ա� = Nvl(rsBaby!Ӥ���Ա�)
            rsData!���� = Nvl(rsBaby!����ʱ��)
        End If
        
        rsData!Ӥ�� = "��"
    Else
        rsData!Ӥ�� = ""
    End If
    
    
    If InStr(Nvl(rsData!ҽ������), ":") > 0 Then '�µ�ģʽ������ҽ����������Ϣ�� ����,ִ�б��:��λ(����,����),��λ---
        rsData!��λ���� = Split(Nvl(rsData!ҽ������), ":")(1)
        rsData!ҽ������ = Split(Nvl(rsData!ҽ������), ":")(0)
    End If
    

    rsData!�����ӡ = IIf(Val(Nvl(rsData!�����ӡ)) = 1, "��", "")
    rsData!���淢�� = IIf(Nvl(rsData!���淢��) = "0", "0", "1")
    
    If mlngModule = G_LNG_PATHSTATION_MODULE Then   'ֻ��ҽ���ž߱���Ƭ��ӡ�ͽ�Ƭ�������
        rsData!��Ƭ��ӡ = IIf(Val(Nvl(rsData!��Ƭ��ӡ)) = 1, "��", "")
        rsData!��Ƭ���� = IIf(Val(Nvl(rsData!��Ƭ����)) = 1, "1", "0")
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then    '��ȡ������ִ��״̬
        rsData!����ִ��״̬ = GetPatholExecuteState(rsData)
    End If
    
    
    If Val(Nvl(rsData!·��)) = 1 Then
        rsData!·�� = " "
    Else
        rsData!·�� = ""
    End If
    
    
    If Val(Nvl(rsData!����)) <> 0 Then
        rsData!���� = " "
    Else
        rsData!���� = ""
    End If
    
    '������Դ
    If Val(Nvl(rsData!��Դ)) = 1 Then
        rsData!��Դ = "��"
    ElseIf Val(Nvl(rsData!��Դ)) = 2 Then
        rsData!��Դ = "ס"
    ElseIf Val(Nvl(rsData!��Դ)) = 3 Then
        rsData!��Դ = "��"
    ElseIf Val(Nvl(rsData!��Դ)) = 4 Then
        rsData!��Դ = "���"
    Else
        rsData!��Դ = ""
    End If
    
    If mlngTempCharged = 0 Then         'δ�շ�
        rsData!�շ� = ""
    ElseIf mlngTempCharged = 1 Then     '���շ�
        rsData!�շ� = " "
    ElseIf mlngTempCharged = 2 Then     '�޷���
        rsData!�շ� = "  "
    ElseIf mlngTempCharged = 5 Then     '���ʷ���
        rsData!�շ� = "    "
    Else                                '�貹��
        rsData!�շ� = "   "
    End If
    
    If Val(Nvl(rsData!����)) <> 0 Then
        rsData!���� = " "  ' ��������
    Else
        rsData!���� = ""
    End If
    
    If Val(Nvl(rsData!Σ��)) <> 0 Then
        rsData!Σ�� = " "
    Else
        rsData!Σ�� = ""
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        intTxtLen = Len(mSysPar.strImageLevel) - Len(Replace(mSysPar.strImageLevel, ",", "")) + 1

        If Trim(Val(Nvl(rsData!Ӱ������))) <> 0 Then
            If Val(rsData!Ӱ������) <= intTxtLen Then
                If Trim(Split(mSysPar.strImageLevel, ",")(Val(rsData!Ӱ������) - 1)) <> "" Then
                    rsData!Ӱ������ = Trim(Split(mSysPar.strImageLevel, ",")(Val(rsData!Ӱ������) - 1))
                Else
                    rsData!Ӱ������ = "δ����"
                End If

            Else
                rsData!Ӱ������ = "��Ч�ȼ�"
            End If
        End If
    End If


    intTxtLen = Len(mSysPar.strReportLevel) - Len(Replace(mSysPar.strReportLevel, ",", "")) + 1

    If Trim(Val(Nvl(rsData!��������))) <> 0 Then
        If Val(rsData!��������) <= intTxtLen Then
            If Trim(Split(mSysPar.strReportLevel, ",")(Val(rsData!��������) - 1)) <> "" Then
                rsData!�������� = Trim(Split(mSysPar.strReportLevel, ",")(Val(rsData!��������) - 1))
            Else
                rsData!�������� = "δ����"
            End If

        Else
            rsData!�������� = "��Ч�ȼ�"
        End If
    End If
    
    If Val(Nvl(rsData!���ӽ�Ƭ)) <> 0 Then
        rsData!���ӽ�Ƭ = "������"
    Else
        rsData!���ӽ�Ƭ = "δ����"
    End If
err.Clear
End Sub


Private Sub ufgStudyList_OnOrderChange(ByVal lngCol As Long, ByVal lngOrder As Integer, blnCustom As Boolean)
'���浱ǰ��������Ϣ
On Error GoTo ErrHandle
    mlngSortCol = lngCol
    mintSortOrder = lngOrder
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnRefreshRowData(rsBind As ADODB.Recordset, ByVal lngRow As Long)
On Error GoTo ErrHandle
    Dim strTag As String
    Dim strTemp As String
    Dim i As Long
    
    Dim strPatientType As String
    
    ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = &H8000000F '&HE0E0E0
    
    For i = 0 To ufgStudyList.DataGrid.Cols - 1
        Select Case ufgStudyList.DataGrid.TextMatrix(0, i)
                
            Case "·��"
                If ufgStudyList.Text(lngRow, "·��") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("·��").Picture
                End If
                
            Case "����"
                If ufgStudyList.Text(lngRow, "����") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("����").Picture
                End If
        
            Case "��Դ"
                strTag = Decode(ufgStudyList.Text(lngRow, "��Դ"), "��", 1, "ס", 2, "��", 3, 4)
                ufgStudyList.DataGrid.Cell(flexcpData, lngRow, i) = strTag
                
                If ufgStudyList.Text(lngRow, "��Դ") = "ס" Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("סԺ").Picture
                End If
                
            Case "�շ�" 'TODO:������Ҫ���ǲ��ɷ��õ����
                If ufgStudyList.Text(lngRow, "�շ�") = "" Then          'δ�շ�
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("Ƿ��").Picture
                ElseIf ufgStudyList.Text(lngRow, "�շ�") = " " Then     '���շ�
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("�շ�").Picture
                ElseIf ufgStudyList.Text(lngRow, "�շ�") = "   " Then   '����
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("����").Picture
                ElseIf ufgStudyList.Text(lngRow, "�շ�") = "    " Then  '�Ƿ�
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("�Ƿ�").Picture
                Else '�޷���("  ")
                    '�޷��ò���ʾͼ��
                End If
                
            Case "Σ��"
                If ufgStudyList.Text(lngRow, "Σ��") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("Σ��").Picture
                End If
                
            Case "����"
                If ufgStudyList.Text(lngRow, "����") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("����").Picture
                End If
                
            Case "����" '���Ϊ��ɫͨ��������Ҫ��������ǰ���ͼ��
                If Val(ufgStudyList.Text(lngRow, "��ɫͨ��")) <> 0 Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("��ɫͨ��").Picture
                End If
                
            Case GetStudyNumberDisplayName  '���Ż��߲����
                If ufgStudyList.Text(lngRow, "���UID") <> "" Then
                    '����ϵͳ�У�����б��еļ�����ʾΪ�����
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages(IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "����", "Ӱ��")).Picture
                End If
            
            Case "��鼼ʦ"
                If Val(ufgStudyList.Text(lngRow, "�Ƿ�ʦȷ��")) = 1 Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("��鼼ʦ").Picture
                End If
                
            Case "������"
                '���ݼ����̣����ò�ͬ����ɫ
                If mSysPar.lngListColorMark = 0 Then
                    ufgStudyList.DataGrid.Cell(flexcpForeColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = getRowColor(lngRow)
                Else
                    ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = getRowColor(lngRow)
                End If
        End Select
        
    Next i
    
    If mSysPar.blnNameColColorCfg Then
        rsBind.Filter = "ҽ��id=" & Val(ufgStudyList.Text(lngRow, "ҽ��ID"))
        
        If rsBind.RecordCount > 0 Then strPatientType = Nvl(rsBind!��������)
        
        If mstrDefaultPatientType = strPatientType Then
            If mSysPar.blnOrdinaryNameColColorCfg Then
                Call NameColorCfg(lngRow, strPatientType)
            End If
        Else
            Call NameColorCfg(lngRow, strPatientType)
        End If
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        i = ufgStudyList.GetColIndex("����")
        
        If ufgStudyList.Text(lngRow, "����") = "����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbGreen
        If ufgStudyList.Text(lngRow, "����") = "��������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbYellow
        If ufgStudyList.Text(lngRow, "����") = "������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbRed
    Else
        i = ufgStudyList.GetColIndex("�������")
        
        If ufgStudyList.Text(lngRow, "�������") = "����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbGreen
        If ufgStudyList.Text(lngRow, "�������") = "��������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbYellow
        If ufgStudyList.Text(lngRow, "�������") = "������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbRed
    End If
    
ErrHandle:
    Exit Sub
End Sub

Private Sub NameColorCfg(ByVal lngRow As Long, ByVal strPatientType As String)
    Dim lngR1 As Long, lngG1 As Long, lngB1 As Long
    Dim lngR2 As Long, lngG2 As Long, lngB2 As Long
    Dim lngPatiColor As Long, lngForeColor As Long
        
    lngPatiColor = zlDatabase.GetPatiColor(strPatientType)
    lngForeColor = getRowColor(lngRow)
    
    Call GetRGB(lngPatiColor, lngR1, lngG1, lngB1)
    Call GetRGB(lngForeColor, lngR2, lngG2, lngB2)
    
    '������������ɫ���б��е�ǰ��ɫ���ʱ��Ϊ�˱������忴���壬���ǰ��ɫ�����ʵ�����
    If Abs(lngR1 - lngR2) < 30 Then
        lngR2 = lngR1 - 30
    End If
    
    If Abs(lngG1 - lngG2) < 30 Then
        lngG2 = lngG1 - 30
    End If
    
    If Abs(lngB1 - lngB2) < 30 Then
        lngB2 = lngG1 - 30
    End If
    
    lngForeColor = RGB(lngR2, lngG2, lngB2)
    
    ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, ufgStudyList.GetColIndex("����")) = lngPatiColor
    ufgStudyList.DataGrid.Cell(flexcpForeColor, lngRow, ufgStudyList.GetColIndex("����")) = lngForeColor
End Sub

Private Sub ufgStudyList_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
On Error GoTo ErrHandle
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Dim i As Long
        
        Set Popup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        
        For i = 1 To cbrMain.ActiveMenuBar.Controls.Count
            Set Menucontrol = cbrMain.ActiveMenuBar.Controls(i)
            
'            If Menucontrol.Parent.BarID = conMenu_ManagePopup Then
            If (Menucontrol.ID = conMenu_ManagePopup Or Menucontrol.ID = conMenu_Collection) And Menucontrol.type = xtpControlPopup Then
                For Each control In Menucontrol.CommandBar.Controls
                    '�����Ҽ� "�ղص�" �˵�
                    If control.ID <> conMenu_Collection_ViewShare And control.ID <> conMenu_Collection_Manage _
                    And Mid(control.ID, 1, Decode(InStr(control.ID, "0") - 1, -1, 0, InStr(control.ID, "0") - 1)) <> comMenu_Collection_Type _
                    And Mid(control.ID, 1, Decode(InStr(control.ID, "0") - 1, -1, 0, InStr(control.ID, "0") - 1)) <> conMenu_Collection_ViewShare Then
                        '���ޱ������֮ǰ������ģ�鴴�����Ҽ��˵�
                        If control.ID = conMenu_Manage_Finish Then
                            If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlMenu.zlPopupMenu(Popup)
                            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(Popup)
                        End If
                        
                        control.Copy Popup
                    End If
                Next
            End If
        Next i
        
'        If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlMenu.zlPopupMenu(Popup)
'        If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(Popup)
        
        Popup.ShowPopup
    End If
ErrHandle:
End Sub

Private Function GetNullAdviceInf() As TAdviceInf
    With GetNullAdviceInf
        .lngPatID = 0
        .strPatientName = ""
        .lngPatDept = 0
        .strPatientDepartment = ""
        .lngAdviceId = 0
        .lngUnit = 0
        .lngSendNo = 0
        .strStudyUID = ""
        .blnCanPrint = False
        .blnIsInsidePatient = False
        .intMoved = -1
        .intState = -1
        .intStep = -1
        .strRegNo = ""
        .lngRegId = 0
        .lngExeDepartmentId = 0
        .strExeRoom = ""
        .lngPatientFrom = 0
        .strDoDoctor = ""
        .strStudyNum = ""
        .strBedNum = ""
        .lngMarkNum = 0
        .lngBaby = -1
    End With
End Function

Private Sub FillCurAdviceTxtInfor()
'������Ϸ����˻�����Ϣ
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If mcurAdviceInf.lngAdviceId <= 0 Then
        labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":---]"
        lbl������Ϣ.Caption = "����:  �Ա�:  ����:"
        lbl�����Ϣ.Caption = "���˿���:" & "  ��ʶ��:" & "  ��  ��:"
        Exit Sub
    End If
    
    With ufgStudyList
        lbl������Ϣ.Caption = "����:" & .CurText("����") & "  �Ա�:" & .CurText("�Ա�") & "  ����:" & .CurText("����")
        
        If mSysPar.blnNameColColorCfg Then
            If mcurAdviceInf.strPatientType = "" Or (mstrDefaultPatientType = mcurAdviceInf.strPatientType And Not mSysPar.blnOrdinaryNameColColorCfg) Then
                lbl������Ϣ.ForeColor = &HC00000
                lbl�����Ϣ.ForeColor = &HC00000
            Else
                lbl������Ϣ.ForeColor = zlDatabase.GetPatiColor(mcurAdviceInf.strPatientType)
                lbl�����Ϣ.ForeColor = zlDatabase.GetPatiColor(mcurAdviceInf.strPatientType)
            End If
        End If
        
        If Not mblnIsHistory Then  '---------------------------�����μ��ֱ�����б��м�¼���
            
            labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":" & IIf(mcurAdviceInf.strStudyNum <> "", mcurAdviceInf.strStudyNum, "---") & "]"
            lbl�����Ϣ.Caption = "���˿���:" & mcurAdviceInf.strPatientDepartment & _
                                "  ��ʶ��:" & mcurAdviceInf.lngMarkNum & _
                                "  ����:" & mcurAdviceInf.strBedNum
                                  
            Select Case .CurText("�շ�")
                Case ""
                    lblCash.Caption = "Ƿ"
                    lblCash.ForeColor = &H80FF&
                Case " "
                    lblCash.Caption = "��"
                    lblCash.ForeColor = &H8000&
                Case "  "
                    lblCash.Caption = "��"
                    lblCash.ForeColor = &HC00000
                Case "   "
                    lblCash.Caption = "��"
                    lblCash.ForeColor = &HFF&
                Case "    "
                    lblCash.Caption = "��"
                    lblCash.ForeColor = &HFF00FF
            End Select
            
            lblCash.Visible = True

        Else
            If mcurAdviceInf.lngAdviceId > 0 Then
                labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":" & IIf(mcurAdviceInf.strStudyNum <> "", mcurAdviceInf.strStudyNum, "---") & "]"
                lbl�����Ϣ.Caption = "���˿���:" & mcurAdviceInf.strPatientDepartment & _
                                      "  �� ʶ ��:" & mcurAdviceInf.lngMarkNum & _
                                      "  ��ǰ����:" & mcurAdviceInf.strBedNum
                
                If mcurAdviceInf.lngBaby <> 0 Then
                    
                    strSql = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                            "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
                            
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡӤ����Ϣ", mcurAdviceInf.lngPatID, mcurAdviceInf.lngPageID, mcurAdviceInf.lngBaby)
                    
                    If Not rsTemp.EOF Then
                        lbl������Ϣ.Caption = "����:" & Nvl(rsTemp!Ӥ������) & "  �Ա�:" & Nvl(rsTemp!Ӥ���Ա�) & _
                                            "  ����:" & Nvl(rsTemp!����ʱ��)
                    End If
                End If
            Else
                labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":---]"
                lbl�����Ϣ.Caption = "���˿���:" & "  ��ʶ��:" & "  ��  ��:"
            End If
            
            lblCash.Caption = "��"
            lblCash.ForeColor = &HC00000
            lblCash.Visible = True
        End If
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetScanRequestCount(ByVal lngAdviceId As Long) As Long
'��ȡɨ�����뵥������
On Error GoTo ErrHandle
    Dim lngCount As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    GetScanRequestCount = 0
    
    '����������뵥ɨ����� ��ѡ������ִ�в�ѯ�õ����뵥ͼ��������δ��ѡ�� ��ִ��
    If mSysPar.blnIsPetitionScan Then
        '����ҽ��ID��ѯ Ӱ�����뵥ͼ����õ���ɨ������ ����ҽ����������� VSList
        strSql = "select count(*) as ͼ���� from Ӱ�����뵥ͼ�� where ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�õ�ͼ������", lngAdviceId)
        
        lngCount = Val(rsTemp!ͼ����)
    Else
        lngCount = 0
    End If
    
    GetScanRequestCount = lngCount
Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function



Private Sub FillCurAdviceAppend(Optional ByVal intImgCount As Integer = 0)
'������½�ҽ������
On Error GoTo ErrHandle
    Dim strAppend As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim strTemp As String
    Dim lngCount As Long
    
    With ufgStudyList
    
        If Not mblnIsHistory Then '-------------------------------------------�б�ѡ�����
            If .GridRows <= 1 Then
                txtAppend.Text = ""
                Exit Sub
            End If
    
            txtAppend = "�����Ŀ:" & .CurText("ҽ������") & vbCrLf
            
            '����������뵥ɨ����� ��ѡ������ҽ��������ʾ�����뵥״̬����δ��ѡ�� ����ʾ
            If mSysPar.blnIsPetitionScan Then txtAppend = txtAppend & "���뵥״̬:" & IIf(intImgCount = 0, "δɨ��", "��ɨ�裨" & intImgCount & "�ţ�") & vbCrLf
            
            txtAppend = txtAppend & "����ҽ��:" & Rpad(.CurText("����ҽ��"), 8, " ") & vbCrLf
            
            If .CurText("��λ����") <> "" Then
                For i = 0 To UBound(Split(.CurText("��λ����"), "),"))
                    If i = 0 Then
                        txtAppend = txtAppend & "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(.CurText("��λ����"), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.CurText("��λ����"), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) 'ȡ����������
            Else
                txtAppend = txtAppend & "��鲿λ:" & .CurText("ҽ������")
            End If
        Else                    '-------------------------------------------���μ�¼ѡ�����
            txtAppend = ""
            
            lngCount = GetScanRequestCount(mcurAdviceInf.lngAdviceId)
            
            gstrSQL = "Select ����ҽ��,ҽ������ From ����ҽ����¼ Where  id =[1]"
            If mcurAdviceInf.intMoved = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ������", mcurAdviceInf.lngAdviceId)
            
            If rsTemp.EOF = False Then
                strTemp = Nvl(rsTemp!ҽ������)
                If InStr(strTemp, ":") > 0 Then
                    txtAppend = "�����Ŀ:" & Split(strTemp, ":")(0) & vbCrLf
                Else
                    txtAppend = "�����Ŀ:" & strTemp & vbCrLf
                End If
                
                If mSysPar.blnIsPetitionScan Then txtAppend = txtAppend & "���뵥״̬:" & IIf(lngCount = 0, "δɨ��", "��ɨ�裨" & lngCount & "�ţ�") & vbCrLf
                
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
        End If
        
        gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����" '�������μ�¼�Ƿ�ת���жϲ���ʷ��
        If mcurAdviceInf.intMoved = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˸���", mcurAdviceInf.lngAdviceId)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!��Ŀ & ":" & Nvl(rsTemp!����) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        gstrSQL = "select ��Ϣ��,��Ϣֵ from ������Ϣ�ӱ� where ����ID=[1] and ����id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ������Ϣ", mcurAdviceInf.lngPatID, mcurAdviceInf.lngAdviceId)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!��Ϣ�� & ":" & Nvl(rsTemp!��Ϣֵ) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        txtAppend = txtAppend & vbCrLf & vbCrLf & strAppend
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FillHistoryStudy()
'������μ���¼
On Error GoTo ErrHandle
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    
    If mListAdviceInf.lngAdviceId = 0 Then
        cboTimes.Clear
        Exit Sub
    End If
    
    cboTimes.tag = "" 'cbotime����ʱ�õ�������������"������Ŀ"ʱ��������"���cbotimes"����
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        'mListAdviceInf.intState = 2��ʾ�Ѿܾ�
        strSql = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
               "" & IIf(mListAdviceInf.intState = 2, "", " And B.ִ��״̬<>2 ") & _
               " AND A.ID=C.ҽ��ID"
    Else
        strSql = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
               " From ����ҽ����¼ A,����ҽ������ B,��������Ϣ C" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
               "" & IIf(mListAdviceInf.intState = 2, "", " And B.ִ��״̬<>2 ") & _
               " AND A.ID=C.ҽ��ID"
    End If
               
    '�Ƿ�ѡ����ȫ������
    If mblnAllDepts = True Then
        strSql = strSql & " And Instr( [3],A.ִ�п���id ) >0 "
    Else
        strSql = strSql & " And A.ִ�п���id+0 =[2] "
    End If
    
    '���ù������ˣ��Ų�ѯ����ID
    If mSysPar.blnRelatingPatient = True And mListAdviceInf.lngLinkId <> 0 Then
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            strSql = strSql & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                " From ����ҽ����¼ A " & _
                " Where A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) "
        Else
            strSql = strSql & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                " From ����ҽ����¼ A, ��������Ϣ B " & _
                " Where A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) and a.id=b.ҽ��ID "
        End If
    End If
    
    strTemp = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
    strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
    strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
    strTemp = Replace(strTemp, "���˼����Ϣ", "H���˼����Ϣ")
    strSql = strSql & vbNewLine & " Union ALL " & vbNewLine & strTemp
    strSql = "Select * From (" & vbNewLine & strSql & vbNewLine & ") Order By ����ʱ�� Asc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mListAdviceInf.lngPatID, _
            mlngCur����ID, mstrCanUse����IDs, mListAdviceInf.lngLinkId)
    
    cboTimes.Clear
    Do Until rsTemp.EOF
       cboTimes.AddItem "��" & rsTemp.AbsolutePosition & "��/��" & rsTemp.RecordCount & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & Trim(rsTemp!ҽ������)
       cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!ҽ��ID
       
       If rsTemp!ҽ��ID = mListAdviceInf.lngAdviceId Then cboTimes.ListIndex = cboTimes.NewIndex
       
       rsTemp.MoveNext
    Loop
    
    If cboTimes.ListCount > 1 Then
        cboTimes.ForeColor = &HC0&
    Else
        cboTimes.ForeColor = &H80000008
    End If
    
    cboTimes.tag = "���"

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ShowTab()
'���ݲ�����Դ���Ʋ�����ҽ��ѡ�
On Error GoTo ErrHandle
    Dim i As Integer
    Dim strFirstTab As String
    Dim intDefaultIndex As Integer
    Dim blnShowReport As Boolean
    
    If TabWindow.ItemCount <= 0 Then Exit Sub
    
    blnShowReport = False
     
    If Not mblnIsHistory Then '-------------------------------------------�б�ѡ�����
        '�ж� ��ͼ����д����
        blnShowReport = True
        
        If mSysPar.blnReportWithImage = True Then
            If mcurAdviceInf.strStudyUID = "" Then blnShowReport = False
        End If
    End If
    
    If mcurAdviceInf.lngPatientFrom <> 2 Then '���ݲ�����Դ���Ʋ�����ҽ��ѡ�
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = True
                    
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = False
                    
                Case "������Ӳ���"
                    TabWindow(i).Visible = True
                
                Case "סԺ���Ӳ���"
                    TabWindow(i).Visible = False
                    
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(Not mblnIsHistory, (mcurAdviceInf.intStep > 1 Or mcurAdviceInf.intStep = -1) And blnShowReport Or GetWorkModuleCount = 1, True)
                Case "�Ŷӽк�"
                    TabWindow(i).Visible = mSysPar.blnUseQueue 'True '
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = False

                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = True
                
                Case "������Ӳ���"
                    TabWindow(i).Visible = False
                
                Case "סԺ���Ӳ���"
                    TabWindow(i).Visible = True

                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(Not mblnIsHistory, (mcurAdviceInf.intStep > 1 Or mcurAdviceInf.intStep = -1) And blnShowReport Or GetWorkModuleCount = 1, True)
                Case "�Ŷӽк�"
                    TabWindow(i).Visible = mSysPar.blnUseQueue 'True '
            End Select
        Next
    End If
    
    
    
    intDefaultIndex = GetTabWindowIndex
    
    
    '�����ǰ��ѡ���ҳ�治�ɼ�������ʾ�û�����Ҫ����ҳ��
    If TabWindow.Selected Is Nothing Then
        strFirstTab = mstrFirstTab
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).tag, strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    If TabWindow.Selected Is Nothing Then TabWindow(intDefaultIndex).Selected = True

    If TabWindow.Selected.Visible = False Then
        strFirstTab = mstrFirstTab
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).tag, strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    If TabWindow.Selected.Visible = False Then
        If intDefaultIndex < 0 Then
            TabWindow.Selected.Visible = True
        Else
            TabWindow(intDefaultIndex).Selected = True
            TabWindow(intDefaultIndex).Visible = True
        End If
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshModuleAdviceInf()
'ˢ��ģ��ҽ����Ϣ
On Error GoTo ErrHandle
    Dim intStep As Long

    If mcurAdviceInf.intState = 2 Then intStep = -2
    
    'ˢ��Ӱ��ҽ��ģ���ҽ����Ϣ
    If Not mfrmWork_PacsImg Is Nothing Then
        Call mfrmWork_PacsImg.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        Call mfrmWork_PacsImg.zlUpdateOtherInf(ufgStudyList, cboTimes)
    End If
    
    'ˢ����Ƶ�ɼ�ģ���ҽ����Ϣ
    If mblnUseActivexCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        End If
    End If

    
    'ˢ�²������ģ���ҽ����Ϣ
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
    End If
    
    'ˢ��HIS���ģ���ҽ����Ϣ
    If Not mobjWork_His Is Nothing Then
        Call mobjWork_His.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        Call mobjWork_His.zlUpdateOtherInf(mcurAdviceInf.lngPatID, mcurAdviceInf.lngUnit, mcurAdviceInf.lngPatDept, mcurAdviceInf.lngPageID, _
            mcurAdviceInf.intState, mcurAdviceInf.strRegNo, mblnIsHistory, mcurAdviceInf.blnIsInsidePatient)
    End If
    
    'ˢ�±���ģ������ҽ����Ϣ
    If Not mobjWork_Report Is Nothing Then
        'δ����ǰ������༭���治��ʾ
        If mcurAdviceInf.intStep < 2 And mcurAdviceInf.intStep <> -1 Then
            Call mobjWork_Report.zlUpdateAdviceInf(0, 0, 0, 0, mblnMoved)
            Call mobjWork_Report.zlRefreshFace
        Else
            Call mobjWork_Report.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngPatID, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        End If
        
        Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, ufgStudyList, mblnIsHistory, mcurAdviceInf.blnCanPrint, mcurAdviceInf.strDoDoctor, mcurAdviceInf.strStudyUID)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ�ˢ��TABҳ��
'������ lngAdviceIDtmp���μ�¼ʱ���� , ������0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ErrHandle
    
    If TabWindow.Selected Is Nothing Then Exit Sub
    
    If TabWindow.Selected.tag = "" Then Exit Sub
    
    Select Case TabWindow.Selected.tag
        Case "Ӱ��ͼ��"
            Call mfrmWork_PacsImg.zlRefreshFace
            
        Case "�걾����"
            Call mobjWork_Pathol.GetModule(mtSpecimen).zlRefreshFace
            
        Case "����ȡ��"
            Call mobjWork_Pathol.GetModule(mtMaterial).zlRefreshFace
            
        Case "������Ƭ"
            Call mobjWork_Pathol.GetModule(mtSlices).zlRefreshFace
            
        Case "�����ؼ�"
            Call mobjWork_Pathol.GetModule(mtSpeExam).zlRefreshFace
            
        Case "���̱���"
            Call mobjWork_Pathol.GetModule(mtProRep).zlRefreshFace
            
        Case "������д"
            Call mobjWork_Report.zlRefreshFace
            If GetActiveWindow = Me.hWnd Then Call mobjWork_Report.zlShowReportVideo
            
        Case "�������", "סԺҽ��", "����ҽ��", "סԺ����", "���ﲡ��", "������Ӳ���", "סԺ���Ӳ���"
            Call mobjWork_His.zlRefreshFace(, mcurAdviceInf.lngPatientFrom)
            
        Case "Ӱ��ɼ�"
            If mblnUseActivexCapture Then
                'ʹ��ActivexExe��Ƶ�ɼ��Ĵ���ʽ
                If Not mobjWork_ActiveVideo Is Nothing Then
'                    Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceID, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved)
'                    Call mobjWork_ActiveVideo.zlRefreshVideoWindow
'                    Call mobjWork_ActiveVideo.zlRefreshData
                End If
            End If

    End Select

    If mblnUseActivexCapture Then
        'ʹ��ActivexExe��Ƶ�ɼ��Ĵ���ʽ
        '�����л����ǲɼ�ҳ�棬Ȼ���л����󣬲ɼ�����ͼ�������
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved)
            Call mobjWork_ActiveVideo.zlRefreshVideoWindow
            Call mobjWork_ActiveVideo.zlRefreshData
'            If mobjWork_ActiveVideo.VideoDockState Then
'                '������ڸ�������״̬������Ҫ��Ӧˢ������
'                mobjWork_ActiveVideo.zlRefreshData
'            End If
        End If
    End If

    
    If TabWindow.Selected.tag <> "Ӱ��ɼ�" And TabWindow.Selected.tag <> "�Ŷӽк�" Then
        If mcurAdviceInf.lngAdviceId <= 0 Then
            Call DisableWorkModule
        Else
            Call EnableWorkModule
        End If
    Else
        EnableWorkModule
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_��������()
'��������
On Error GoTo ErrHandle
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Call frmReferencePatient.zlShowMe(mListAdviceInf.lngAdviceId, ufgStudyList.CurText("����"), Me, True, mlngCur����ID)
    
    'ˢ�²����б�
     Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_�����ɼ�()
On Error GoTo ErrHandle

    If Not GetIsValidOfStorageDevice(mlngCur����ID) Then
      MsgBoxD Me, "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
      Exit Sub
    End If
    
    If Not mobjWork_ActiveVideo Is Nothing Then
        Call mobjWork_ActiveVideo.zlShowPopupVideo
        
        If mlngOldAdviceId <> mcurAdviceInf.lngAdviceId And TabWindow.Selected.Caption <> "Ӱ��ɼ�" Then
            Call mobjWork_ActiveVideo.zlRefreshData
            mlngOldAdviceId = mcurAdviceInf.lngAdviceId
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_ͼ���¼()
'ͼ���¼
    Dim lngCurAdviceId As Long
    Dim objBurn As Object
    Dim frmBurn As frmImageBurn
    
    If mListAdviceInf.intImageLocation = 1 Then
        Call subXWShowArchiveManager(3)
    Else
        On Error GoTo errExit
            Set objBurn = CreateObject("IMAPI2.MsftDiscMaster2")
            Set objBurn = Nothing
            GoTo continueBurn
errExit:
            Call MsgBoxD(Me, "���ܴ�����¼�������ڰ�װIMAPI2��¼��������½��롣", vbOKOnly, Me.Caption)
            Exit Sub
            
continueBurn:
            
            Set frmBurn = New frmImageBurn
        On Error GoTo errFree
            
            lngCurAdviceId = Val(ufgStudyList.CurKeyValue)
            
            Set frmBurn = New frmImageBurn
            Call frmBurn.ShowBurn(mlngModule, mlngCur����ID, lngCurAdviceId, mblnMoved, Me)
errFree:
            Call Unload(frmBurn)
            Set frmBurn = Nothing
    End If
End Sub

Private Sub Menu_Manage_�ղع���()
'�ղع���
On Error GoTo errFree
    Dim frmCollectionManage As New frmCollectionManage
    Dim lngCount As Long

    Call frmCollectionManage.ShowCollectionManageWind(Me)
    
    'ɾ�����ڵĹ������������˵���
    Call LockWindowUpdate(Me.hWnd)
    For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbrMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbrMain.Count To 2 Step -1
        cbrMain(lngCount).Delete
    Next
    
    Call InitCommandBars
    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call CreateWorkModuleMenu
    
    Call LockWindowUpdate(0)
    
errFree:
    Call Unload(frmCollectionManage)
    Set frmCollectionManage = Nothing
End Sub

Private Sub Menu_Manage_�ղص�()
'�ղص�
    Dim frmToCollection As New frmToCollection
    Dim rsTemp As ADODB.Recordset
    Dim lngAdviceId As Long
    Dim lngSendNo As Long
On Error GoTo errFree

    lngAdviceId = Val(ufgStudyList.CurText("ҽ��ID"))
    lngSendNo = Val(ufgStudyList.CurText("���ͺ�"))
    
    If lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    gstrSQL = "select �״�ʱ�� from ����ҽ������ where ҽ��ID= " & lngAdviceId & ""
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    '�ж�ѡ�м�¼�Ƿ񱨵������û�б������ܽ����ղز���
    Do While Not rsTemp.EOF
        If Nvl(rsTemp!�״�ʱ��) = "" Then
            Call MsgBoxD(Me, "�ü��δ�����������ղأ�", vbOKOnly, "Ӱ������վ")
            Exit Sub
        End If
        
        rsTemp.MoveNext
    Loop
    
    Call frmToCollection.ShowToCollectionWind(Me, lngAdviceId, lngSendNo)
    
errFree:
    Call Unload(frmToCollection)
    Set frmToCollection = Nothing
End Sub


Private Sub Menu_Manage_�ղ�������ʾ(ByVal control As XtremeCommandBars.ICommandBarControl, ByVal bytStyle As Byte)
'�ղ�������ʾ����
On Error GoTo errHand
    Dim rsList As ADODB.Recordset
    Dim strCollectionType As String
    Dim lngFatherID As Long
    Dim strUser As String
    
    '�����ղ�����ַ���
    If InStr(control.Caption, "(") = 0 Then
        strCollectionType = control.Caption
    Else
        strCollectionType = Mid(control.Caption, 1, InStr(control.Caption, "(") - 1)
    End If
    
    '������������
    strUser = control.DescriptionText ' Category
    
    '������ID�ַ���
    If bytStyle = 0 Then
        lngFatherID = CLng(control.ID) - CLng(comMenu_Collection_Type) * 10000#
    ElseIf bytStyle = 1 Then
        lngFatherID = CLng(control.ID) - CLng(conMenu_Collection_ViewShare) * 10000#
    End If
    
    '���������� ���ݼ��ط���
    Set rsList = GetCollectionData(strCollectionType, lngFatherID, strUser)
   

    
    Set ufgStudyList.AdoData = rsList
    
    ufgStudyList.AdoFilter = ""
    
    Call ufgStudyList.BindData
   
    If ufgStudyList.AdoData.RecordCount > 0 Then Call ufgStudyList_OnSelChange

    Call RefreshStatusBarInf
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_CheckList()
    If mListAdviceInf.lngAdviceId > 0 Then
        Set mclsCISKernel = New clsCISKernel
        Call mclsCISKernel.ShowPacsApplication(Me, mListAdviceInf.lngAdviceId)
        Set mclsCISKernel = Nothing
    Else
        MsgBox "û��ѡ�в��ˡ�", vbInformation + vbOKOnly, gstrSysName
    End If
End Sub
Private Function GetCollectionData(ByVal strCollectionType As String, ByVal lngFatherID As Long, ByVal strUser As String) As ADODB.Recordset
'���ع�������
    Dim strSql As String
    Dim strWithCollectionTab As String
    Dim strWithOrderTab As String   'ҽ���Ӳ�ѯ
    
    Set GetCollectionData = Nothing
    
    '���ݲ����ж�������һ��SQL���
    If Len(Trim(strCollectionType)) <> 0 And strCollectionType <> "�鿴��ǰ�ղ�" Then
        strWithCollectionTab = " with tmpCollection as (select m.ҽ��id as id from Ӱ���ղ���� L,Ӱ���ղ����� M where " & vbNewLine & _
                                        " L.id=M.�ղ�id" & vbNewLine & _
                                        " and l.������='" & Decode(strUser, "", UserInfo.����, strUser) & "' and l.�ղ����='" & strCollectionType & "' )"
    ElseIf lngFatherID <> 0 Then
        strWithCollectionTab = " with tmpCollection as (select m.ҽ��id as id from Ӱ���ղ���� L,Ӱ���ղ����� M where " & vbNewLine & _
                                        "L.id=M.�ղ�id" & vbNewLine & _
                                        " and L.id in (select distinct id from Ӱ���ղ���� start with id =" & lngFatherID & " connect by prior id=�ϼ�id) )"
    End If
    
    strWithOrderTab = "tmpOrder as (select id from tmpCollection Union All select a.ID from ����ҽ����¼ a, tmpCollection b where a.���ID=b.ID and a.ҽ��״̬ <> 4 )"
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSql = "Select * From (" & vbNewLine & _
             strWithCollectionTab & "," & vbNewLine & strWithOrderTab & vbNewLine & _
             "Select  Distinct" & vbNewLine & _
                    "       A.ҽ��ID,B.���ID,A.���ͺ�,A.�״�ʱ�� ����ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.ִ�м�,A.������� ����,h.Σ��״̬ Σ��," & vbNewLine & _
                    "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,Decode(B.������Դ, 1, '��', 2, 'ס', 3, '��', 4, '��') ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
                    "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��,b.����ʱ��,c.�����,c.סԺ��," & vbNewLine & _
                    "       Nvl(B.����,H.����) ����,H.Ӱ�����,H.����,Nvl(B.�Ա�,H.�Ա�) �Ա�,Nvl(B.����,H.����) ����,H.���,H.����,H.Ӱ������,H.�������," & vbNewLine & _
                    "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������,H.���淢��,H.���Ž�Ƭ,H.����ID,A.��¼����, " & vbNewLine & _
                    "       H.�����,H.�Ƿ���ӽ�Ƭ,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.��������,H.������,H.�Ƿ�ʦȷ��,H.��鼼ʦ,H.��鼼ʦ��,H.�������� ��ͼʱ��," & vbNewLine & _
                    "       H.�������,H.��Ϸ���,H.���UID,H.ͼ��λ��,A.ִ�в���ID as ִ�п���ID,0 as ת��,F.���� AS ���˿���, a.����ʱ��, " & vbNewLine & _
                    "       C.���￨��,A.NO as ���ݺ�,C.���֤��,C.����ģʽ,decode(B.������Դ,2,D.��������,C.��������) as ��������,D.·��״̬,A.�Ʒ�״̬,Decode(A.��¼����,2,1,Decode(a.�Ʒ�״̬,3,1,0)) as �շ� ,z.ҽ��ID as ���뵥ҽ��" & vbNewLine & _
                    " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,������ҳ D,Ӱ�����¼ H,���ű� F, tmpOrder ,Ӱ�����뵥ͼ�� Z" & vbNewLine & _
                    " Where A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) " & vbNewLine & _
                    " And B.����ID=C.����ID And B.���˿���id=F.ID " & vbNewLine & _
                    " And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) and a.ҽ��ID = z.ҽ��ID(+) and a.ҽ��ID=tmpOrder.id) "
    Else
        strSql = "Select * From (" & vbNewLine & _
             strWithCollectionTab & "," & vbNewLine & strWithOrderTab & vbNewLine & _
             "Select Distinct" & vbNewLine & _
             "       A.ҽ��ID,B.���ID,A.���ͺ�,A.�״�ʱ�� ����ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.������� ����,h.Σ��״̬ Σ��," & vbNewLine & _
             "       '' as ����ִ��״̬, o.ȡ�Ĺ���,o.��Ƭ����,o.���߹���,o.���ӹ���,o.��Ⱦ����,b.����ʱ��,c.�����,c.סԺ��, " & vbNewLine & _
             "       decode(o.�������,0,'����',1,'����',2,'ϸ��',3,'����',4,'ʬ��',5,'����ʯ��',null) as  ������, " & vbNewLine & _
             "       decode(o.�����,null,'δ����','�Ѻ���') as �������, " & vbNewLine & _
             "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,Decode(B.������Դ, 1, '��', 2, 'ס', 3, '��', 4, '��') ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
             "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��," & vbNewLine & _
             "       Nvl(B.����,H.����) ����,Nvl(B.�Ա�,H.�Ա�) �Ա�,Nvl(B.����,H.����) ����,H.���,H.����,o.�ۺ�����," & vbNewLine & _
             "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������,o.�����,H.���淢��,H.���Ž�Ƭ,H.����ID,A.��¼����, " & vbNewLine & _
             "       H.�����,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.��������,H.������,H.�Ƿ�ʦȷ��,H.��鼼ʦ,H.��鼼ʦ��,H.�������� ��ͼʱ��, " & vbNewLine & _
             "       H.�������,H.��Ϸ���,H.���UID,H.ͼ��λ��,0 as ת��,F.���� AS ���˿���, a.����ʱ��, Y.��ǰ״̬ as ����״̬, Y.����ҽʦ, Y.Id as ����ID, " & vbNewLine & _
             "       C.���￨��,A.NO as ���ݺ�,C.���֤��,C.����ģʽ,decode(B.������Դ,2,D.��������,C.��������) as ��������,D.·��״̬,A.�Ʒ�״̬,Decode(A.��¼����,2,1,Decode(a.�Ʒ�״̬,3,1,0)) as �շ�,z.ҽ��ID as ���뵥ҽ��, " & vbNewLine & _
             "      (select count(1) from ��������Ϣ V , ����������Ϣ W where V.����ҽ��ID=w.����ҽ��id and v.ҽ��id=A.ҽ��ID and w.����״̬=1) as ���� " & vbNewLine & _
             " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,������ҳ D,Ӱ�����¼ H,���ű� F, " & vbNewLine & _
             "       ��������Ϣ o, tmpOrder ,Ӱ�����뵥ͼ�� Z, ���������Ϣ Y" & vbNewLine & _
             " Where A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) " & vbNewLine & _
             "       And B.����ID=C.����ID And B.���˿���id=F.ID " & vbNewLine & _
             "       and A.ҽ��ID=o.ҽ��ID(+) and o.����ҽ��ID=Y.����ҽ��ID(+) " & vbNewLine & _
             "       And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) and a.ҽ��ID = z.ҽ��ID(+) and a.ҽ��ID=tmpOrder.id) "
    End If
      
    strSql = strSql & vbNewLine & "Order by ������,����ʱ��,����ʱ��"
    
    Set GetCollectionData = GetDataToLocal(strSql, GetWindowCaption)
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Sub Menu_Petition_ɨ�����뵥(ByVal intType As Integer)
'intType:0--�鿴���뵥��1--ɨ�����뵥
Dim frmPetitionCap As New frmPetitionCapture

On Error GoTo errFree

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    With ufgStudyList
        Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                                mlngCur����ID, _
                                                Nvl(.CurText("���˿���")), _
                                                Nvl(.CurText("����")), _
                                                Nvl(.CurText("����")), _
                                                Nvl(.CurText("�Ա�")), _
                                                Nvl(.CurText("ҽ������")), _
                                                Nvl(.CurText("��λ����")), _
                                                IIf(InStr(mstrPrivs, "���Ǽ�") <= 0, True, IIf(intType = 0, True, False)), _
                                                False, _
                                                mListAdviceInf.lngAdviceId, _
                                                IIf(.CurText("������") = "�Ѿܾ�", 1, IIf(.CurText("������") = "�����", 2, 0)))
    End With
errFree:
    Call Unload(frmPetitionCap)
    Set frmPetitionCap = Nothing
End Sub

Private Sub ufgStudyList_OnSelChange()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    '����Ǵ�ӡ�嵥�Ĳ��� ��ֹͣ�иı��¼�����Ȼ����ɽ������ˢ��
    If mblnIsPrintMode Then Exit Sub
    
    mblnIsHistory = False
    
    If mblnvsRefresh Then Exit Sub
    
    gblnIsStudyChage = True
    
    mcurAdviceInf = GetAdviceDetailInf()
    mListAdviceInf = mcurAdviceInf
    
    Call FillCurAdviceTxtInfor '������Ϸ����˻�����Ϣ
    Call FillHistoryStudy '������μ���¼
    Call SetSelectRowColor
    
    If mListAdviceInf.lngAdviceId <= 0 Then '�޼�¼ʱ����
        cboTimes.Clear
        txtAppend = ""

        lblCash.Visible = False
        
        If Not TabWindow.Selected Is Nothing Then
            Call ConfigSubForm(TabWindow.Selected)
        End If
    
        Call RefreshModuleAdviceInf
        Call RefreshTabWindow
    Else
        mintImgCount = GetScanRequestCount(mListAdviceInf.lngAdviceId)

        Call RefreshModuleAdviceInf
        
        Call FillCurAdviceAppend(mintImgCount) '������½�ҽ������
        Call ShowTab '���ݲ����ṩ��ͬѡ�
        
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        
        If Not TabWindow.Selected Is Nothing Then
            Call ConfigSubForm(TabWindow.Selected)
        End If

        
        '�ж��Ƿ��ֶ�ˢ�µļ���б�������ֶ�ˢ�£�����Ҫ֪ͨ��������ģ�����ˢ�£�...
        If mblnIsCallModuleRefresh Then
            mblnIsCallModuleRefresh = False
            
            Call NotificationAllModuleRefresh
        End If
        
        If mstrFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ,��TabWindow����ˢ��
            
            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    Exit For
                End If
            Next
            
            If i = TabWindow.ItemCount Then    'ûѭ�����˴�����1������TAB
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow.Item(i).Visible Then
                        Exit For
                    End If
                Next i
            End If
            
            'ˢ��ҳ�棬����ʾ������ҳ
            If TabWindow.Item(i).Selected Then
                Call RefreshTabWindow
            Else
                TabWindow.Item(i).Selected = True
            End If
        Else
            Call RefreshTabWindow
        End If
        
    End If
    
    '�ָ����㣬��������ˢ�¹����У���������б����ʧȥ��ʧȥ����󣬽�����ʹ�������ֹ����б�
    If ufgStudyList.Visible And Not mblnAutoRefreshList Then 'GetActiveWindow = Me.hWnd
        Me.dkpMain.Panes(1).Selected = True
        Call ufgStudyList.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetSelectRowColor()
    Dim lngRowSel As Long
    
    lngRowSel = ufgStudyList.DataGrid.RowSel
    
    If lngRowSel < 0 Then Exit Sub
    
    Call SetStateColor(lngRowSel)
    
    If ufgStudyList.DataGrid.Cols > 1 And ufgStudyList.DataGrid.Rows > 1 Then
        ufgStudyList.DataGrid.Cell(flexcpFontBold, 1, 1, ufgStudyList.DataGrid.Rows - 1, ufgStudyList.DataGrid.Cols - 1) = False

        ufgStudyList.DataGrid.Cell(flexcpFontBold, lngRowSel, 1, lngRowSel, ufgStudyList.DataGrid.Cols - 1) = True
        
        ufgStudyList.DataGrid.Cell(flexcpFontSize, 1, 1, ufgStudyList.DataGrid.Rows - 1, ufgStudyList.DataGrid.Cols - 1) = mbytFontSize
    End If
End Sub

Private Sub SetStateColor(ByVal lngRowSel As Long)
    Dim lngForeColor As Long
    Dim lngR As Long, lngG As Long, lngB As Long
    
    If mSysPar.lngListColorMark <> 0 Then Exit Sub
    
    lngForeColor = getRowColor(lngRowSel)
    
    Call GetRGB(lngForeColor, lngR, lngG, lngB)
    
    ufgStudyList.DataGrid.ForeColorSel = RGB(lngR - 30, lngG - 30, lngB - 30)
    ufgStudyList.DataGrid.BackColorSel = &HFEE0E2      '&HFECFD2
End Sub

Private Function getRowColor(ByVal lngRowSel As Long) As Long
    Dim lngRowColor As Long
    
    If ufgStudyList.Text(lngRowSel, "������") = "�Ѿܾ�" Then lngRowColor = gdblColor�Ѿܾ�
    If ufgStudyList.Text(lngRowSel, "������") = "�����" Then lngRowColor = gdblColor�����
    If ufgStudyList.Text(lngRowSel, "������") = "�ѱ���" Then lngRowColor = gdblColor�ѱ���
    If ufgStudyList.Text(lngRowSel, "������") = "�ѵǼ�" Then lngRowColor = gdblColor�ѵǼ�
    If ufgStudyList.Text(lngRowSel, "������") = "�Ѽ��" Then lngRowColor = gdblColor�Ѽ��
    If ufgStudyList.Text(lngRowSel, "������") = "�����" Then lngRowColor = gdblColor�����
    If ufgStudyList.Text(lngRowSel, "������") = "������" Then lngRowColor = gdblColor������
    If ufgStudyList.Text(lngRowSel, "������") = "������" Then lngRowColor = gdblColor������
    If ufgStudyList.Text(lngRowSel, "������") = "�����" Then lngRowColor = gdblColor�����
    If ufgStudyList.Text(lngRowSel, "������") = "�ѱ���" Then lngRowColor = gdblColor�ѱ���
    If ufgStudyList.Text(lngRowSel, "������") = "�Ѳ���" Then lngRowColor = gdblColor�Ѳ���
    
    getRowColor = lngRowColor
End Function

Private Sub Menu_Manage_SetXWParam_click()
'------------------------------------------------
'���ܣ�������PACS�Ĳ������ô���
'���أ�
'------------------------------------------------
    On Error GoTo err
    
    Call frmXWSetParams.zlShowMe(Me)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub conMenu_File_SendImg_click()
'------------------------------------------------
'���ܣ�����ͼ��
'���أ�
'------------------------------------------------
    On Error GoTo err
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        If mListAdviceInf.lngAdviceId <= 0 Or mListAdviceInf.intImageLocation = 1 Then
            Call subXWShowArchiveManager(2)
        Else
            frmPacsSendImage.ShowMe Me
        End If
    Else
        frmPacsSendImage.ShowMe Me
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


