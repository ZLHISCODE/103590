VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSReqEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ǽ�"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9075
   Icon            =   "frmPACSReqEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9075
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picAdvice 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   -120
      ScaleHeight     =   4425
      ScaleWidth      =   9195
      TabIndex        =   64
      Top             =   2790
      Width           =   9195
      Begin VB.Frame fraCheck 
         Caption         =   "�����Ϣ"
         Height          =   2205
         Left            =   300
         TabIndex        =   88
         Top             =   1770
         Width           =   8700
         Begin VB.OptionButton optMatch 
            Caption         =   "����"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   49
            ToolTipText     =   "�����Ž����˺ͽ��յ�Ӱ�����ƥ��"
            Top             =   1890
            Width           =   1395
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "����/סԺ��"
            Height          =   195
            Index           =   1
            Left            =   2265
            TabIndex        =   50
            ToolTipText     =   "�����˱�ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
            Top             =   1890
            Width           =   1395
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "����ʶ��"
            Height          =   195
            Index           =   2
            Left            =   4320
            TabIndex        =   51
            ToolTipText     =   "������ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
            Top             =   1890
            Width           =   1395
         End
         Begin VB.ComboBox cboRoom 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   525
            Width           =   1995
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   7
            Left            =   3960
            MaxLength       =   30
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   525
            Width           =   1410
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   1
            Left            =   6480
            MaxLength       =   30
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   525
            Width           =   2055
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   3
            Left            =   1050
            MaxLength       =   30
            TabIndex        =   42
            Top             =   885
            Width           =   1995
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   6
            Left            =   6480
            MaxLength       =   3
            TabIndex        =   46
            Top             =   885
            Width           =   1815
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   5
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   44
            Top             =   885
            Width           =   1215
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "������(&C)"
            Height          =   255
            Left            =   180
            TabIndex        =   47
            Top             =   1260
            Width           =   1455
         End
         Begin VB.CheckBox chk��Ƭ 
            Caption         =   "���Ž�Ƭ(&F)"
            Height          =   255
            Left            =   2010
            TabIndex        =   48
            Top             =   1260
            Width           =   1335
         End
         Begin VB.CheckBox chkUnicode 
            Caption         =   "ͬһ���ߵļ����ڱ�����ͳһ���"
            Height          =   210
            Left            =   180
            TabIndex        =   89
            Top             =   240
            Width           =   4080
         End
         Begin MSComCtl2.DTPicker DTBirth 
            Height          =   300
            Left            =   6480
            TabIndex        =   93
            Top             =   525
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   62455811
            CurrentDate     =   38156
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "������ǰ���еļ�飬��������Ŀƥ����ͼ��"
            Height          =   180
            Left            =   165
            TabIndex        =   92
            Top             =   1620
            Width           =   3960
         End
         Begin VB.Label lblRoom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�м�(&R)"
            Height          =   180
            Left            =   165
            TabIndex        =   35
            Top             =   585
            Width           =   810
         End
         Begin VB.Label Label18 
            Caption         =   "����(&U)"
            Height          =   255
            Left            =   3105
            TabIndex        =   37
            Top             =   585
            Width           =   975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "KG"
            Height          =   180
            Left            =   8370
            TabIndex        =   91
            Top             =   990
            Width           =   180
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "CM"
            Height          =   180
            Left            =   5205
            TabIndex        =   90
            Top             =   990
            Width           =   180
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "����豸(&D)"
            Height          =   180
            Index           =   8
            Left            =   5445
            TabIndex        =   39
            Top             =   585
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "Ӣ����(&E)"
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   41
            Top             =   945
            Width           =   810
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "����(&W)"
            Height          =   180
            Index           =   3
            Left            =   5805
            TabIndex        =   45
            Top             =   945
            Width           =   630
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "���(&H)"
            Height          =   180
            Index           =   7
            Left            =   3285
            TabIndex        =   43
            Top             =   945
            Width           =   630
         End
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         ItemData        =   "frmPACSReqEdit.frx":08CA
         Left            =   4260
         List            =   "frmPACSReqEdit.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1440
         Width           =   1995
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7935
         TabIndex        =   86
         Top             =   4065
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6675
         TabIndex        =   85
         Top             =   4065
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   285
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1100
      End
      Begin VB.CommandButton cmd�ɼ� 
         Height          =   285
         Left            =   6645
         Picture         =   "frmPACSReqEdit.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "ѡ�����걾"
         Top             =   360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt�ɼ� 
         Height          =   300
         Left            =   4740
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chk��ʼʱ�� 
         BackColor       =   &H80000004&
         Caption         =   "Ҫ��ʱ��"
         Height          =   225
         Left            =   315
         TabIndex        =   23
         ToolTipText     =   "�Ƿ���ʱ��"
         Top             =   420
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7230
         MaxLength       =   3
         TabIndex        =   31
         Top             =   1080
         Width           =   1380
      End
      Begin VB.TextBox txtƵ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1350
         TabIndex        =   29
         Top             =   1080
         Width           =   2500
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4740
         MaxLength       =   3
         TabIndex        =   30
         Top             =   1080
         Width           =   1500
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000004&
         Caption         =   "����(&J)"
         Height          =   225
         Left            =   7710
         TabIndex        =   27
         Top             =   405
         Width           =   945
      End
      Begin VB.CommandButton cmdExt 
         Height          =   285
         Left            =   8340
         Picture         =   "frmPACSReqEdit.frx":09C4
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "ѡ�����걾"
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   285
         Left            =   5280
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   0
         Width           =   285
      End
      Begin VB.ComboBox cboִ�п��� 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmPACSReqEdit.frx":0ABA
         Left            =   1350
         List            =   "frmPACSReqEdit.frx":0ABC
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1440
         Width           =   1995
      End
      Begin VB.TextBox txtҽ������ 
         Height          =   300
         Left            =   1350
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   0
         Width           =   3945
      End
      Begin VB.ComboBox cboҽ�� 
         Height          =   300
         Left            =   7230
         TabIndex        =   34
         Text            =   "cboҽ��"
         Top             =   1425
         Width           =   1380
      End
      Begin VB.TextBox txtҽ������ 
         Height          =   300
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   28
         Top             =   720
         Width           =   7245
      End
      Begin VB.CommandButton cmdƵ�� 
         Enabled         =   0   'False
         Height          =   240
         Left            =   3575
         Picture         =   "frmPACSReqEdit.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   1110
         Width           =   270
      End
      Begin MSComCtl2.DTPicker txt��ʼʱ�� 
         Height          =   300
         Left            =   1350
         TabIndex        =   24
         Top             =   360
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   62455811
         CurrentDate     =   38022
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   3450
         TabIndex        =   87
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lbl�ɼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɼ���ʽ"
         Height          =   180
         Left            =   3930
         TabIndex        =   77
         Top             =   405
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Line lineTitleSplit 
         BorderColor     =   &H80000000&
         X1              =   400
         X2              =   1440
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����걾"
         Height          =   180
         Left            =   5940
         TabIndex        =   76
         Top             =   45
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ��"
         Height          =   180
         Left            =   6840
         TabIndex        =   75
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   8460
         TabIndex        =   74
         Top             =   1140
         Width           =   15
      End
      Begin VB.Label lblƵ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ƶ��"
         Height          =   180
         Left            =   960
         TabIndex        =   73
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   6150
         TabIndex        =   72
         Top             =   1140
         Width           =   15
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   4455
         TabIndex        =   71
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lblִ�п��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         Height          =   180
         Left            =   600
         TabIndex        =   70
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ(&I)"
         Height          =   180
         Left            =   330
         TabIndex        =   69
         Top             =   45
         Width           =   990
      End
      Begin VB.Label lbl��ʼʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ҫ��ʱ��"
         Height          =   180
         Left            =   600
         TabIndex        =   68
         Top             =   435
         Width           =   720
      End
      Begin VB.Label lbl����ҽ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   6840
         TabIndex        =   67
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   585
         TabIndex        =   66
         Top             =   795
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   ">>"
      Height          =   300
      Left            =   8520
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "���ಡ����Ϣ"
      Top             =   450
      Width           =   315
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���#���շѵ��ݺ�"
      Top             =   60
      Width           =   2160
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   6300
      MaxLength       =   10
      TabIndex        =   3
      Top             =   60
      Width           =   2220
   End
   Begin VB.ComboBox cbo�Ա� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3990
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1635
   End
   Begin VB.ComboBox cbo�ѱ� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   450
      Width           =   2160
   End
   Begin VB.ComboBox cbo���ʽ 
      Height          =   300
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   450
      Width           =   2220
   End
   Begin MSComctlLib.ImageList iLstItem 
      Left            =   8280
      Top             =   3480
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
            Picture         =   "frmPACSReqEdit.frx":0BB4
            Key             =   "Ԫ��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMain 
      Left            =   7680
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":0CC6
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":0EE2
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":10FE
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":131A
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":1536
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":1752
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":1B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":1DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":1FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":21E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":23FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":261A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":2834
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":2A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":31C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":35FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":3816
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":3A30
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":41AA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":4924
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":4B3E
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":4D58
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMainHot 
      Left            =   6360
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":53D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":55F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":5812
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":5A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":5C52
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":5E72
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":6092
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":62AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":64CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":66EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":690C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":6B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":6D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":6F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":717A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":78F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":7B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":7D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":7F42
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":815C
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":88D6
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":9050
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":926A
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":9484
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLstTab 
      Left            =   6960
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":9AFE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSReqEdit.frx":A098
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt����� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   10
      TabIndex        =   82
      Top             =   450
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   -30
      TabIndex        =   52
      Top             =   840
      Width           =   9135
      Begin VB.CommandButton cmd��λ���� 
         Caption         =   "��"
         Height          =   285
         Left            =   8220
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   780
         Width           =   285
      End
      Begin VB.CommandButton cmd��ͥ��ַ 
         Caption         =   "��"
         Height          =   285
         Left            =   8220
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   1170
         Width           =   285
      End
      Begin VB.TextBox txt��ͥ�ʱ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7275
         MaxLength       =   6
         TabIndex        =   18
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5460
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Height          =   300
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1170
         Width           =   6945
      End
      Begin VB.TextBox txt��λ�ʱ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3315
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt��λ�绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt��λ���� 
         Height          =   300
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   11
         Top             =   780
         Width           =   6945
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1260
         MaxLength       =   18
         TabIndex        =   10
         Top             =   390
         Width           =   7245
      End
      Begin VB.ComboBox cboְҵ 
         Height          =   300
         Left            =   7275
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   1260
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   1260
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   3315
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   0
         Width           =   1260
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʱ�"
         Height          =   180
         Left            =   6825
         TabIndex        =   63
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   240
         Left            =   4680
         TabIndex        =   62
         Top             =   1620
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   240
         Left            =   480
         TabIndex        =   61
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʱ�"
         Height          =   180
         Left            =   2865
         TabIndex        =   60
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   240
         Left            =   480
         TabIndex        =   59
         Top             =   1620
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         Height          =   240
         Left            =   480
         TabIndex        =   58
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   240
         Left            =   480
         TabIndex        =   57
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   825
         TabIndex        =   56
         Top             =   60
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   2865
         TabIndex        =   55
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   240
         Left            =   6825
         TabIndex        =   54
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   240
         Left            =   4680
         TabIndex        =   53
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ѱ�"
      Height          =   240
      Left            =   810
      TabIndex        =   81
      Top             =   510
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   240
      Left            =   5850
      TabIndex        =   80
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "���ʽ"
      Height          =   420
      Left            =   5850
      TabIndex        =   79
      Top             =   420
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      Height          =   240
      Left            =   3525
      TabIndex        =   78
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPACSReqEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public strPrivs As String       '�û����б�����ľ���Ȩ��
Private blnOK As Boolean

Private FileID As String
Private PatientID As String '����ID
Private CheckID As String '����ID��Һŵ�ID
Private PatientType As Integer '1=���ﲡ�� 2=סԺ����
Private FileTypeID As String '����ģ���ļ�ID
Private bSample As Boolean '�Ƿ�ʾ��
Private bln��ʿվ As Boolean
Private ParentForm As Object
Private DeptID As Long '��������
Private ItemType As Integer  '������Ŀ��� 1=PACS 2=LIS
Private ItemDeptID As Long '��Ŀִ�п���

Private PatientDate As Date '���˾������Ժʱ��
Private AdviceID As Long, SendNO As Long 'ҽ��ID�����ͺ�
Private sCheckNo As String '���͵��ݺ�
Private iRecordType As Integer '��¼����
Private alngFileID(1) As Long '����ͱ���ID
Private intType As Integer '�������:-1=������0=�����ϡ�1=������2=��ҩ��4=����
Private iTabIndex As Integer
Private mlngǰ��ID As Long, blnҽ��ִ�� As Boolean

'ҽ���༭
Private strAdviceText As String 'ҽ������
Private str��� As String, lngClinicID As Long, strClinicName As String, str�걾��λ As String
Private strSequence As String, lngƵ�ʴ��� As Long, lngƵ�ʼ�� As Long, str�����λ As String 'Ƶ��
Private int�Ƽ����� As Integer, intִ������ As Integer, lng���˿���ID As Long
Private mstr�Ա� As String
Private mstrLike As String
Private gint�����Ǽ���Ч���� As Integer
Private rsRelativeAdvice As ADODB.Recordset '���ҽ��
Private strExtData As String '������Ŀ

Private ifInitItem As Boolean '�Ƿ��ڽ�������ʱֱ����ʾ������Ŀ

Private iInputType As Integer
'����������ǰ����״̬�����һֱ�Ը�״̬���Բ�����ǰ����
'0�����￨
'1������ID
'2��סԺ��
'3�������
'4���Һŵ�
'5���շѵ��ݺ�
'6������

Private iCurrElementIndex As Integer '��ǰԪ��˳���

Private mblnCheck As Boolean '�Ƿ���ʼ���
Private aDevices() As String
Private strImageType As String 'Ӱ�����
Private lng��ʶ�� As Long
Private mblnContiAdd As Boolean '�Ƿ���������

Private Declare Function GetParent Lib "user32" (ByVal Hwnd As Long) As Long

Public Function ShowMe_Request(frmParent As Object, ByVal lngDeptID As Long, Optional ByVal iItemType As Integer = 1, Optional ByVal ModalWindow As Boolean = True, Optional ByVal lngǰ��ID As Long = 0, _
    Optional ByVal blnCheck As Boolean = True) As Boolean
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    
    On Error Resume Next
    '��ʼ��
    Set rsRelativeAdvice = Nothing
    
    alngFileID(0) = 0
    PatientType = 1: AdviceID = 0: PatientID = 0: CheckID = ""
    mlngǰ��ID = lngǰ��ID: ItemType = iItemType: ItemDeptID = lngDeptID
    mblnCheck = blnCheck
    lngClinicID = 0: strDiagName = "": strDrAdvice = ""
    strExtData = ""
    mblnContiAdd = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�����Ǽ�����", 0))
    '��ʼ������
    
    '��ȡ������Ϣ
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
    DeptID = UserInfo.����ID
    
    '��ʼ������
    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    '��ʼҽ���б�
'    Call Get����ҽ��(0, bln��ʿվ, "", 0, Me.cboҽ��, PatientType)
    
    Set ParentForm = frmParent
    
    initForm
    Me.cmdCancel.Caption = IIf(mblnContiAdd, "�ر�(&C)", "ȡ��(&C)")
    ifInitItem = True
    
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
    ShowMe_Request = blnOK
End Function

Private Sub ClearData()
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    
    On Error Resume Next
    '��ʼ��
    alngFileID(0) = 0
    PatientType = 1: AdviceID = 0: PatientID = 0: CheckID = ""
    strDiagName = "": strDrAdvice = ""
    '��ʼ������
    
    '��ȡ������Ϣ
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    DeptID = UserInfo.����ID
    
    '�����Ļ
    Me.txt���� = "": Me.txt���֤�� = "": Me.txt��λ���� = "": Me.txt��ͥ��ַ = ""
    Me.txt��λ�绰 = "": Me.txt��λ�ʱ� = "": Me.txt��ͥ�绰 = "": Me.txt��ͥ�ʱ� = ""

    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
End Sub

Private Sub initForm()
    intType = Switch(ItemType = 1, 0, ItemType = 2, 4)
    Select Case intType
        Case 0
            Me.Caption = "���Ǽ�"
        Case 1
            Me.Caption = "�����Ǽ�"
        Case 4
            Me.Caption = "����Ǽ�"
        Case Else
            Me.Caption = "�Ǽ�"
    End Select

    SetItemFormat
End Sub

Private Sub SetItemFormat()   '����������Ŀ������ʾ��ʽ
    Select Case intType
        Case 0
            Me.lblҽ������.Caption = "�����Ŀ": Me.lbl����.Caption = "��鲿λ": Me.cmdExt.ToolTipText = "ѡ���鲿λ"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case 1
            Me.lblҽ������.Caption = "������Ŀ": Me.lbl����.Caption = "����ʽ": Me.cmdExt.ToolTipText = "ѡ������ʽ"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case 4
            Me.lblҽ������.Caption = "������Ŀ": Me.lbl����.Caption = "����걾": Me.cmdExt.ToolTipText = "ѡ�����걾"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
            Me.lbl�ɼ�.Visible = True: Me.txt�ɼ�.Visible = True: Me.cmd�ɼ�.Visible = True
        Case Else
            Me.lbl����.Visible = False: Me.txt����.Visible = False: Me.cmdExt.Visible = False
    End Select
End Sub

Private Sub cboRoom_Click()
    On Error Resume Next
    txtItem(1) = aDevices(cboRoom.ListIndex)
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��������_Click()
    InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim StrInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboҽ��.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboҽ��.Text = "" Then '������
        Exit Sub
    End If
    
    StrInput = UCase(NeedName(cboҽ��.Text))
    'ȫԺҽ��
    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And B.����ID IN(" & strSQL & ")" & _
        " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    
    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl����ҽ��.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboҽ��.Height, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
    If Not rsTmp Is Nothing Then
        cboҽ��.Text = Nvl(rsTmp!����) & "-" & rsTmp!����
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkUnicode_Click()
    If mblnCheck Then
        txtItem(7).Text = Next����(lngClinicID, PatientID, chkUnicode.Value = 1)
    End If
End Sub

Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk��Ƭ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMore_Click()
    Me.Frame1.Visible = Not Me.Frame1.Visible
    If Me.Frame1.Visible Then
        Me.cbo����.SetFocus
    Else
        Me.txtҽ������.SetFocus
    End If
    Me.Height = Me.Height + IIf(Me.Frame1.Visible, 1, -1) * Me.Frame1.Height
    
    Form_Resize
End Sub

Private Sub cmdOK_Click()
    If Len(sCheckNo) > 0 Then
        If MsgBox("��ǰ������Ŀ�����շѵ��ݣ�" & sCheckNo & " �������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    If Not SaveFile Then Exit Sub
    
    blnOK = True
    If mblnContiAdd Then
        ClearData
        Me.txt����.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub cmd�ɼ�_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID))
    Me.txt�ɼ�.SetFocus
    If Not rsTmp Is Nothing Then
        Me.cmd�ɼ�.Tag = rsTmp("ID")
        Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    If picAdvice.Enabled And ifInitItem Then
        Me.txtҽ������ = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������Ŀ", "")
        Me.txt���� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���븽��", "")
        
        'Ĭ��Ϊ���һ�ε���Ŀ
        If Len(Trim(Me.txtҽ������)) > 0 Then
            On Error Resume Next
            Call txtҽ������_KeyPress(vbKeyReturn)
            
            Me.txt����.SetFocus
        End If
        
        ifInitItem = False
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnShowDetail As Boolean
    Dim strExeRoom As String
    
    On Error GoTo errH
    
    blnShowDetail = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������ϸ��Ϣ", "False")
    Me.Height = Me.Height - IIf(blnShowDetail, 0, Me.Frame1.Height)
    Me.Frame1.Visible = blnShowDetail
    
    blnOK = False
    iInputType = -1
    '�й�ҽ���Ĳ���
    mstrLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    
    strSQL = "Select ������,������,����ֵ from ϵͳ������"
    Call OpenRecord(rsTmp, strSQL, "mdlCISCore")
    'Ƥ�Խ����Чʱ��
    rsTmp.Filter = "������=2"
    If Not rsTmp.EOF Then gint�����Ǽ���Ч���� = Val(Nvl(rsTmp!����ֵ, 0))
    
    '---------Ȩ�޿���-------------
    'strPrivs = gstrPrivs
    '��ʼ������Ϣ
    lng���˿���ID = UserInfo.����ID
    Call InitData
    
    '��ʼ�����Ϣ
    If mblnCheck Then
        'ִ�м�����
        strExeRoom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰִ�м�")
            
        strSQL = "Select * From ҽ��ִ�з��� Where ����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, ItemDeptID)
        cboRoom.Clear
        If rsTmp.EOF Then
            cboRoom.AddItem "": cboRoom.ListIndex = 0
        Else
            Do While Not rsTmp.EOF
                cboRoom.AddItem rsTmp!ִ�м�
                rsTmp.MoveNext
            Loop
        End If
        InitDevice
        
        'Ӱ��ƥ������
        i = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "Ӱ��ƥ�䷽ʽ", 0))
        optMatch(i).Value = True
        chkUnicode.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "��ִ�п��ұ��", 0))
        
        On Error Resume Next
        cboRoom.ListIndex = 0
        cboRoom.Text = strExeRoom
        On Error GoTo errH
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim lngTxtWidth As Single
    Dim lngDistance As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = 0
    lngStatus = 0
    lngDistance = 300
    
    On Error Resume Next
    With picAdvice
        .Width = Me.ScaleWidth
    End With
    With Me.chk����
        .Left = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Width
        If .Left < Me.txt�ɼ�.Left + Me.txt�ɼ�.Width + lngDistance Then .Left = Me.txt�ɼ�.Left + Me.txt�ɼ�.Width + lngDistance
    End With
'    With Me.chk����
'        .Left = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Width
'        If .Left < Me.txt��ʼʱ��.Left + Me.txt��ʼʱ��.Width + lngDistance Then .Left = Me.txt��ʼʱ��.Left + Me.txt��ʼʱ��.Width + lngDistance
'    End With
    
    lngTxtWidth = (picAdvice.ScaleWidth - Me.lbl��ʼʱ��.Left - Me.cmdSel.Width - Me.txtҽ������.Left - lngDistance - _
        Me.lbl����.Width - Me.cmdExt.Width - 60) / 2
    With Me.txtҽ������
        .Width = lngTxtWidth
        Me.cmdSel.Left = .Left + .Width
        Me.lbl����.Left = Me.cmdSel.Left + Me.cmdSel.Width + lngDistance
    End With
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
        Me.cmdExt.Left = .Left + .Width
    End With
    Me.lineTitleSplit.X2 = Me.cmdExt.Left + Me.cmdExt.Width + 200

    With Me.txtҽ������
        .Width = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Left
    End With
    
    lngTxtWidth = (picAdvice.Width - Me.lbl��ʼʱ��.Left - Me.txtƵ��.Left - Me.txtƵ��.Width - _
        (Me.lbl������λ.Width + Me.lbl����.Width + lngDistance + 2 * 30) - _
        (Me.lbl������λ.Width + Me.lbl����.Width + lngDistance + 2 * 30)) / 2
    If lngTxtWidth < 1000 Then lngTxtWidth = 1000
    Me.lbl����.Left = Me.txtƵ��.Left + Me.txtƵ��.Width + lngDistance
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl������λ.Left = Me.txt����.Left + Me.txt����.Width + 30
    Me.lbl����.Left = Me.lbl������λ.Left + Me.lbl������λ.Width + lngDistance
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl������λ.Left = Me.txt����.Left + Me.txt����.Width + 30
    
    With Me.cboҽ��
        .Left = Me.txt����.Left
'        .Width = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Left
    End With
    Me.lbl����ҽ��.Left = Me.cboҽ��.Left - Me.lbl����ҽ��.Width

    Me.picAdvice.Top = Me.Frame1.Top + IIf(Me.Frame1.Visible, Me.Frame1.Height, 0)
    
    With Me.cmdMore
        .Caption = IIf(Me.Frame1.Visible, "<<", ">>")
        .ToolTipText = IIf(Me.Frame1.Visible, "����������Ϣ", "��ϸ������Ϣ")
    End With

    If Not mblnCheck And Me.fraCheck.Visible Then
        Me.fraCheck.Visible = False
        
        Me.cmdHelp.Top = Me.cmdHelp.Top - Me.fraCheck.Height
        Me.cmdCancel.Top = Me.cmdCancel.Top - Me.fraCheck.Height
        Me.cmdOK.Top = Me.cmdOK.Top - Me.fraCheck.Height
        Me.Height = Me.Height - Me.fraCheck.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlCommFun.OpenIme False
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������ϸ��Ϣ", Me.Frame1.Visible
    '���������������Ŀ
    If Len(Trim(Me.txtҽ������)) > 0 Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "������Ŀ", Me.txtҽ������
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���븽��", Me.txt����
    End If
End Sub

Private Sub InitDevice()
    Dim i As Integer, iPos As Integer
    Dim strDevices As String, aTmpArray() As String, aTmpArray1() As String
    On Error Resume Next
    
    strDevices = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����豸")
    If Len(Trim(strDevices)) = 0 Then
        ReDim aTmpArray(1, 0) As String
    Else
        aTmpArray1 = Split(strDevices, "||")
        ReDim aTmpArray(1, UBound(aTmpArray1)) As String
        For i = 0 To UBound(aTmpArray1)
            iPos = InStr(aTmpArray1(i), "|")
            If iPos = 0 Then
                aTmpArray(0, i) = ""
                aTmpArray(1, i) = aTmpArray1(i)
            Else
                aTmpArray(0, i) = Mid(aTmpArray1(i), 1, iPos - 1)
                aTmpArray(1, i) = Mid(aTmpArray1(i), iPos + 1)
            End If
        Next
    End If
    
    ReDim aDevices(cboRoom.ListCount - 1) As String
    For i = 0 To cboRoom.ListCount - 1
        iPos = GetIndex(aTmpArray, cboRoom.List(i))
        aDevices(i) = aTmpArray(1, iPos)
    Next
End Sub

Private Function SaveFile() As Boolean
    Dim sTmpFileID As String
    
    SaveFile = False
        
    '��������
    
    If Not ValidAdvice Then Exit Function
    If Not SaveAdvice Then Exit Function

    SaveFile = True
End Function

'���ҽ�����ݵĺϷ���
Private Function ValidAdvice() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    ValidAdvice = True
    
    On Error Resume Next
'    If txt�����.Text = "" Then
'        ValidAdvice = False
'        MsgBox "�����벡�˵�����ţ�", vbInformation, gstrSysName
'        txt�����.SetFocus: Exit Function
'    End If
    If cbo�ѱ�.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "��ѡ���˵ķѱ�", vbInformation, gstrSysName
        cbo�ѱ�.SetFocus: Exit Function
    End If
    If txt����.Text = "" Then
        ValidAdvice = False
        MsgBox "�����벡�˵�������", vbInformation, gstrSysName
        txt����.SetFocus: Exit Function
    End If
    
    If Len(Trim(strAdviceText)) = 0 Then
        ValidAdvice = False
        MsgBox "��������������Ŀ��", vbInformation, gstrSysName
        Me.txtҽ������.SetFocus: Exit Function
    End If
    If Len(Trim(strSequence)) = 0 Then
        ValidAdvice = False
        MsgBox "����ָ��Ƶ�ʣ�", vbInformation, gstrSysName
        Me.txtƵ��.SetFocus: Exit Function
    End If
    If Not Check��ʼʱ��(CStr(Me.txt��ʼʱ��)) Then
        ValidAdvice = False
        Me.txt��ʼʱ��.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt����)) = 0 Then
        ValidAdvice = False
        MsgBox "������������", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt����)) = 0 And Me.txt����.Enabled Then
        ValidAdvice = False
        MsgBox "�����뵥����", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If Val(Me.txt����) > Val(Me.txt����) Then
        ValidAdvice = False
        MsgBox "�������ܴ���������", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If Me.cbo��������.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "��ָ���������ң�", vbInformation, gstrSysName
        Me.cbo��������.SetFocus: Exit Function
    End If
'    If Me.cboҽ��.ListIndex = -1 Then
'        ValidAdvice = False
'        MsgBox "��ָ������ҽ����", vbInformation, gstrSysName
'        Me.cboҽ��.SetFocus: Exit Function
'    End If
    
    '�����ϢУ��
    If mblnCheck Then
        If Len(Trim(txtItem(3))) = 0 Then
            ValidAdvice = False
            MsgBox "������Ӣ������", vbInformation, gstrSysName
            txtItem(3).SetFocus: Exit Function
        End If
        If LenB(StrConv(Trim(txtItem(3).Text), vbFromUnicode)) > txtItem(3).MaxLength Then
            ValidAdvice = False
            MsgBox "Ӣ�������������" & txtItem(3).MaxLength & "���ַ���" & CInt(txtItem(3).MaxLength / 2) & "�����֣���", vbInformation, gstrSysName
            txtItem(3).SetFocus: Exit Function
        End If
        '�жϼ����Ƿ��ظ�
        strSQL = "Select ����,�Ա�,���� From Ӱ�����¼ Where Ӱ�����=[1] And ����=[2]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strImageType, txtItem(7))
        If Not rsTmp.EOF Then
            If MsgBox("��ǰ���������л����ظ����Ƿ������" & Chr(10) & Chr(13) & "������Ϣ��" & Nvl(rsTmp(0)) & " " & Nvl(rsTmp(1)) & " " & Nvl(rsTmp(2)), vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                ValidAdvice = False
                txtItem(7).SetFocus: Exit Function
            End If
        End If
    End If
End Function
'����ҽ��
Private Function SaveAdvice() As Boolean
    On Error GoTo DBError
    SaveAdvice = True
    
    SaveAdviceData
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveAdvice = False
    SaveErrLog
End Function

Private Sub SaveAdviceData()
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim iMaxSeq As Integer, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng��������ID As Long, lng����id As Long, strDoctor As String, i As Integer
    Dim strִ�п���ID As String, strִ�п���ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr��� As String, tmplngClinicID As Long, tmpint�Ƽ����� As Integer, tmpintִ������ As Integer
    Dim rsDept As ADODB.Recordset

    gcnOracle.BeginTrans
    On Error GoTo DBError
    
    '���没����Ϣ
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If PatientType = 1 Then '���ﲡ��
        If PatientID > 0 Then '���еĲ���
            lng����id = PatientID
            strSQL = _
                "zl_�ҺŲ��˲���_INSERT(3," & lng����id & "," & IIf(Len(Trim(txt�����.Text)) = 0, "Null", txt�����.Text) & "," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & "'," & _
                "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "'," & _
                "'" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
                "'" & NeedName(cboְҵ.Text) & "','" & txt���֤��.Text & "','" & txt��λ����.Text & "'," & _
                Val(txt��λ����.Tag) & ",'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & txt��ͥ��ַ.Text & "'," & _
                "'" & txt��ͥ�绰.Text & "','" & txt��ͥ�ʱ�.Text & "'," & strDate & ",NULL)"
        Else '�²���
            lng����id = zlDatabase.NextNo(1)
            strSQL = _
                "zl_�ҺŲ��˲���_INSERT(1," & lng����id & "," & IIf(Len(Trim(txt�����.Text)) = 0, "Null", txt�����.Text) & "," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & "'," & _
                "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "'," & _
                "'" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
                "'" & NeedName(cboְҵ.Text) & "','" & txt���֤��.Text & "','" & txt��λ����.Text & "'," & _
                Val(txt��λ����.Tag) & ",'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & txt��ͥ��ַ.Text & "'," & _
                "'" & txt��ͥ�绰.Text & "','" & txt��ͥ�ʱ�.Text & "'," & strDate & ",NULL)"
        End If
        ExecuteProc strSQL, Me.Caption
'        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'        gcnOracle.Execute strSQL, , adCmdStoredProc
'        Call SQLTest
    Else
        lng����id = PatientID
    End If
    '����ҽ��������
    lngAdviceID = zlDatabase.GetNextId("����ҽ����¼")
    iMaxSeq = 0
    
    lng��������ID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex) 'Get��������ID(Me.cboҽ��.ItemData(Me.cboҽ��.ListIndex), lng���˿���ID, PatientType)
    lng���˿���ID = lng��������ID
    
    i = InStr(Me.cboҽ��.Text, "-")
    If i > 0 Then
        strDoctor = Trim(Mid(Me.cboҽ��.Text, i + 1))
    Else
        strDoctor = Trim(Me.cboҽ��.Text)
    End If
    If Len(Me.cboִ�п���.Text) = 0 Then
        strִ�п���ID = "NULL"
    Else
        strִ�п���ID = Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex)
    End If
    
    tmpstr��� = str���: tmplngClinicID = lngClinicID: tmpint�Ƽ����� = int�Ƽ�����
    tmpintִ������ = intִ������
    iSendSeq = 1
    If intType = 4 Then
        '������Ŀ���ɼ���ʽ��Ϊ��ҽ��
        strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Me.cmd�ɼ�.Tag)
        tmpstr��� = rsTmp("���"): tmplngClinicID = rsTmp("ID"): tmpint�Ƽ����� = Nvl(rsTmp("�Ƽ�����"), 0)
        tmpintִ������ = Nvl(rsTmp("ִ�п���"), 0)
        'ȡ�ɼ���ʽ��ִ�в���
        Set rsDept = GetExeDepart(rsTmp("ID"), PatientType + 1, DeptID)
        If rsDept Is Nothing Then
            strִ�п���ID1 = "NULL"
        Else
            strִ�п���ID1 = rsDept("ID")
        End If
        lngSendNO = zlDatabase.NextNo(10)
        If Len(sCheckNo) = 0 Then
            strNO = zlDatabase.NextNo(IIf(PatientType = 1, 13, 14))
        Else
            strNO = sCheckNo
        End If
    End If
    
    If intType <> 4 Then
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & PatientType & "," & lng����id & "," & IIf(PatientType = 2, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr��� & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txtҽ������, "'", "''") & "'," & _
            "'" & str�걾��λ & "','" & strSequence & "'," & _
            IIf(lngƵ�ʴ��� = 0, "NULL", lngƵ�ʴ���) & "," & _
            IIf(lngƵ�ʼ�� = 0, "NULL", lngƵ�ʼ��) & "," & _
            "'" & str�����λ & "',NULL," & _
            tmpint�Ƽ����� & "," & _
            strִ�п���ID & "," & _
            tmpintִ������ & "," & Me.chk����.Value & "," & _
            IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "NULL,", "To_Date('" & Format(Me.txt��ʼʱ��.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng���˿���ID & "," & lng��������ID & ",'" & strDoctor & "'," & _
            "Sysdate,'" & IIf(PatientType = 2, "", CheckID) & "'," & _
            IIf(mlngǰ��ID = 0, "Null", mlngǰ��ID) & ")"
        
        ExecuteProc strSQL, Me.Caption
'        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'        gcnOracle.Execute strSQL, , adCmdStoredProc
'        Call SQLTest
        '����ҽ��
        lngSendNO = zlDatabase.NextNo(10)
        If Len(sCheckNo) = 0 Then
            strNO = zlDatabase.NextNo(IIf(PatientType = 1, 13, 14))
        Else
            strNO = sCheckNo
        End If
    End If
    '�������ҽ��
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("����ҽ����¼")
            With rsRelativeAdvice
                strSQL = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    (iMaxSeq + i) & "," & PatientType & "," & lng����id & "," & IIf(PatientType = 2, CheckID, "NULL") & "," & _
                    "0,1," & _
                    "1,'" & .Fields("���") & "'," & _
                    .Fields("ID") & ",NULL,NULL," & _
                    IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
                    IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
                    "'" & Replace(.Fields("����"), "'", "''") & "','" & Replace(Me.txtҽ������, "'", "''") & "'," & _
                    "'" & IIf(intType = 4, str�걾��λ, .Fields("�걾��λ")) & "','" & strSequence & "'," & _
                    IIf(lngƵ�ʴ��� = 0, "NULL", lngƵ�ʴ���) & "," & _
                    IIf(lngƵ�ʼ�� = 0, "NULL", lngƵ�ʼ��) & "," & _
                    "'" & str�����λ & "',NULL," & _
                    .Fields("�Ƽ�����") & "," & _
                    strִ�п���ID & "," & _
                    .Fields("ִ�п���") & "," & Me.chk����.Value & "," & _
                    IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "NULL,", "To_Date('" & Format(Me.txt��ʼʱ��.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
                    "NULL," & _
                    lng���˿���ID & "," & lng��������ID & ",'" & strDoctor & "'," & _
                    "Sysdate,'" & IIf(PatientType = 1, CheckID, "") & "'," & _
                    IIf(mlngǰ��ID = 0, "Null", mlngǰ��ID) & ")"
                    ExecuteProc strSQL, Me.Caption
'                    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'                    gcnOracle.Execute strSQL, , adCmdStoredProc
'                    Call SQLTest
                
                iSendSeq = iSendSeq + 1
                strSQL = "ZL_����ҽ������_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & "," & Me.txt���� & ",NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & strִ�п���ID & "," & IIf(Len(sCheckNo) = 0, 0, 1) & ",0)"
                    
                ExecuteProc strSQL, Me.Caption
'                Call SQLTest(App.ProductName, Me.Caption, strSQL)
'                gcnOracle.Execute strSQL, , adCmdStoredProc
'                Call SQLTest
                
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    If intType = 4 Then
        '��������Ĳɼ���ʽ�ŵ����
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & PatientType & "," & lng����id & "," & IIf(PatientType = 2, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr��� & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txtҽ������, "'", "''") & "'," & _
            "'" & str�걾��λ & "','" & strSequence & "'," & _
            IIf(lngƵ�ʴ��� = 0, "NULL", lngƵ�ʴ���) & "," & _
            IIf(lngƵ�ʼ�� = 0, "NULL", lngƵ�ʼ��) & "," & _
            "'" & str�����λ & "',NULL," & _
            tmpint�Ƽ����� & "," & _
            strִ�п���ID1 & "," & _
            tmpintִ������ & "," & Me.chk����.Value & "," & _
            IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "NULL,", "To_Date('" & Format(Me.txt��ʼʱ��.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng���˿���ID & "," & lng��������ID & ",'" & strDoctor & "'," & _
            "Sysdate,'" & IIf(PatientType = 2, "", CheckID) & "'," & _
            IIf(mlngǰ��ID = 0, "Null", mlngǰ��ID) & ")"
        ExecuteProc strSQL, Me.Caption
'        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'        gcnOracle.Execute strSQL, , adCmdStoredProc
'        Call SQLTest
        
        iSendSeq = iSendSeq + 1
    End If
    
    '������ҽ��
    If intType <> 4 Then iSendSeq = 1 '�Ǽ��������ҽ������ǰ��
    strSQL = "ZL_����ҽ������_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & "," & Me.txt���� & ",NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & strִ�п���ID & "," & IIf(Len(sCheckNo) = 0, 0, 1) & ",1)"
'        "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
'        "0," & strִ�п���ID & ",0,1)"
    ExecuteProc strSQL, Me.Caption
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    '�޸ķ��ü�¼��ҽ�����
    If Len(sCheckNo) > 0 Then
        strSQL = "zl_���˷��ü�¼_ҽ��('" & strNO & "',1," & lngAdviceID & ")"
        ExecuteProc strSQL, Me.Caption
'        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'        gcnOracle.Execute strSQL, , adCmdStoredProc
'        Call SQLTest
    End If
    
    '����ʼ�������
    If mblnCheck Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰִ�м�", Me.cboRoom.Text
        strSQL = ""
        For i = 0 To cboRoom.ListCount - 1
            strSQL = strSQL & "||" & cboRoom.List(i) & "|" & aDevices(i)
        Next
        If Len(strSQL) > 0 Then strSQL = Mid(strSQL, 3)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����豸", strSQL
        For i = 0 To optMatch.Count - 1
            If optMatch(i).Value Then Exit For
        Next
        If i > optMatch.Count - 1 Then i = 0
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "Ӱ��ƥ�䷽ʽ", i
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "��ִ�п��ұ��", chkUnicode.Value
    
        strSQL = "ZL_Ӱ����_BEGIN('" & cboRoom.Text & "'," & txtItem(7).Text & "," & lngAdviceID & "," & lngSendNO & ",'" & strImageType & "','" & _
            Me.txt���� & "','" & Trim(txtItem(3)) & "','" & Trim(Me.cbo�Ա�.Text) & "','" & _
            Me.txt���� & "'," & IIf(IsNull(DTBirth.Value), "Null", "to_Date('" & Format(DTBirth.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')") & ",'" & txtItem(5) & "','" & txtItem(6) & "'," & _
            Me.chk����.Value & "," & Me.chk��Ƭ.Value & ",'" & Trim(txtItem(1)) & "'," & _
            0 & ",'" & Me.txt��ͥ�绰 & "')"
        ExecuteProc strSQL, Me.Caption
            
        '������ǰ���еļ��
        strSQL = "Select A.���UID As ID From Ӱ����ʱ��¼ a " & _
            " Where a.����=[1] And a.Ӱ�����=[2]"
        If optMatch(0).Value Then '����
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtItem(7).Text, strImageType)
        End If
        If optMatch(1).Value Then '����/סԺ��
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng��ʶ��, strImageType)
        End If
        If optMatch(2).Value Then '����ʶ�ţ�ҽ��ID��
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, strImageType)
        End If
        If rsTmp.RecordCount = 1 Then '��ͼ��ͼ���Զ�ƥ��
            strSQL = "ZL_Ӱ����_SET(" & lngAdviceID & "," & lngSendNO & ",'" & _
                rsTmp("ID") & "')"
            ExecuteProc strSQL, Me.Caption
        End If
    End If

    gcnOracle.CommitTrans
    AdviceID = lngAdviceID
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "����ҽ������"
End Sub

Private Function GetOneDept(lng�շ�ϸĿID As Long) As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.ִ�в���ID From �շ�ϸĿ A,�շ�ִ�в��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng�շ�ϸĿID)
    If Not rsTmp.EOF Then
        GetOneDept = rsTmp!ִ�в���ID 'Ĭ��ȡ��һ��(���ж��)
    Else
        GetOneDept = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'========������ҽ���༭==========

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk����_Click()
    On Error Resume Next
    Me.txtҽ������.SetFocus
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk��ʼʱ��_Click()
    On Error Resume Next
    If Me.chk��ʼʱ��.Value = 1 Then
        Me.txt��ʼʱ��.Enabled = True: Me.txt��ʼʱ��.SetFocus
    Else
        Me.txt��ʼʱ��.Enabled = False
    End If
    
    If str��� = "D" Then
        strAdviceText = Get�����������(1, strClinicName)
    ElseIf str��� = "F" Then
        strAdviceText = Get�����������(2, strClinicName)
    End If
End Sub

Private Sub chk��ʼʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboҽ��_GotFocus()
    Call zlControl.TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdExt_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim tmpExtData As String
    
    frmAdviceEditEx.mlngHwnd = Me.cboҽ��.Hwnd 'txt����.Hwnd
    frmAdviceEditEx.mintType = IIf(intType = 4, 3, intType)
    frmAdviceEditEx.mint��Ч = 1
    frmAdviceEditEx.mstr�Ա� = mstr�Ա�
    If intType = 4 Then
        '������Ŀ
        frmAdviceEditEx.mlng��ĿID = 0 'Split(strExtData, ";")(0)
        frmAdviceEditEx.mstrExtData = strExtData ' Split(strExtData, ";")(1)
    Else
        frmAdviceEditEx.mlng��ĿID = lngClinicID
        frmAdviceEditEx.mstrExtData = strExtData
    End If
    frmAdviceEditEx.mint������� = PatientType

    On Error Resume Next
    frmAdviceEditEx.Show 1, Me

    If Not frmAdviceEditEx.mblnOK Then
        zlControl.TxtSelAll Me.txt����
        Me.txt����.SetFocus
        Exit Sub
    Else
        tmpExtData = frmAdviceEditEx.mstrExtData
        If intType = 4 Then
            strExtData = Split(strExtData, ";")(0) + ";" + tmpExtData
        Else
            strExtData = tmpExtData
        End If
    End If
    Select Case intType
        Case 0 '�����ϲ�λ
            Call AdviceSet�������(1, strExtData)
            strAdviceText = Get�����������(1, strClinicName)
            Me.txt���� = Get��λ����
        Case 1 '������Ŀ
            Call AdviceSet�������(2, strExtData)
            txtҽ������.Text = Get�����������(2, strClinicName)
            strAdviceText = Get�����������(2, strClinicName)
            Me.txt���� = Get��������
        Case 4 '������Ŀ
            strAdviceText = strClinicName & "(" & tmpExtData & ")"
            Me.txt���� = tmpExtData: str�걾��λ = tmpExtData
    End Select
    txt����.Tag = txt����.Text
    Me.txt����.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset
    
    If intType = 4 Then
        '������Ŀ
        If LabsInput Then
            txtҽ������.Tag = txtҽ������.Text
            txt����.Tag = txt����.Text
            Me.txtҽ������.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            txt����.Text = txt����.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus
        End If
        Exit Sub
    End If
    
    With txtҽ������
        .Text = ""
        Set rsTmp = SelectDiagItem()
    End With
    
    If rsTmp Is Nothing Then 'ȡ����������
        '�ָ�ԭֵ
        zlControl.TxtSelAll txtҽ������
        txtҽ������.SetFocus: Exit Sub
    End If
    '����Ŀ��¼��
    
    '����ѡ����Ŀ����ȱʡҽ����Ϣ
    If AdviceInput(rsTmp) Then
        '��ʾ��ȱʡ���õ�ֵ
        txtҽ������.Tag = txtҽ������.Text
        txt����.Tag = txt����.Text
        If mblnCheck Then ReadCheckData
        
        Me.txtҽ������.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '�ָ�ԭֵ
        txtҽ������.Text = txtҽ������.Tag
        txt����.Text = txt����.Tag
        zlControl.TxtSelAll txtҽ������
        txtҽ������.SetFocus
    End If
End Sub

Private Sub cmdƵ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
        
    int��Χ = 1
    strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
        " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,nvl(A.�����λ,' ') As �����λ" & _
        " From ����Ƶ����Ŀ A Where A.���÷�Χ=" & int��Χ & _
        " Order by A.����"
    vRect = GetControlRect(txtƵ��.Hwnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "����Ƶ��", , , , , , True, vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
        End If
        txtƵ��.Text = strSequence
        Call zlControl.TxtSelAll(txtƵ��)
        txtƵ��.SetFocus: Exit Sub
    End If
    Me.cmdƵ��.Tag = rsTmp("����"): Me.txtƵ�� = Me.cmdƵ��.Tag: strSequence = Me.cmdƵ��.Tag
    lngƵ�ʴ��� = rsTmp("Ƶ�ʴ���"): lngƵ�ʼ�� = rsTmp("Ƶ�ʼ��"): str�����λ = Trim(rsTmp("�����λ"))

    txtƵ��.SetFocus
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    With Me.txtItem(Index)
        .SelStart = 0: .SelLength = .MaxLength
    End With
    Select Case Index
        Case 1, 2
            Call zlCommFun.OpenIme(True)
        Case Else
            Call zlCommFun.OpenIme(False)
    End Select
End Sub

Private Sub txtItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    
    If LenB(StrConv(Trim(txtItem(Index).Text), vbFromUnicode)) >= txtItem(Index).MaxLength Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 5, 6, 7
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
        Case 1, 2
            Call zlCommFun.OpenIme(False)
            If Index = 1 Then aDevices(cboRoom.ListIndex) = txtItem(1)
    End Select
End Sub

'�ж��Ƿ�Ϊ�༭��
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Sub txt�ɼ�_GotFocus()
    Call zlControl.TxtSelAll(txt�ɼ�)
End Sub

Private Sub txt�ɼ�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt�ɼ�.Text = txt�ɼ�.Tag Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID), Me.txt�ɼ�)
    If Not rsTmp Is Nothing Then
        Me.cmd�ɼ�.Tag = rsTmp("ID")
        Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�ɼ�_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txt�ɼ�.Text <> txt�ɼ�.Tag Then
        txt�ɼ�.Text = txt�ɼ�.Tag
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt����) Then Me.txt���� = 1: Exit Sub
    Me.txt���� = CInt(Me.txt����)
    If CInt(Me.txt����) < 1 Then Me.txt���� = 1
End Sub

Private Sub txt����_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt����)
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text = txt����.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        cmdExt_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txt����.Text <> txt����.Tag Then
        txt����.Text = txt����.Tag
    End If
End Sub

Private Sub txt��ʼʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��ʼʱ��_Validate(Cancel As Boolean)
    On Error Resume Next
    If Not Check��ʼʱ��(CStr(txt��ʼʱ��)) Then
        Cancel = True
        txt��ʼʱ��.SetFocus
    Else
        If str��� = "D" Then
            strAdviceText = Get�����������(1, strClinicName)
        ElseIf str��� = "F" Then
            strAdviceText = Get�����������(2, strClinicName)
        End If
    End If
End Sub

Private Sub txtƵ��_GotFocus()
    Call zlControl.TxtSelAll(txtƵ��)
End Sub

Private Sub txtƵ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdƵ��.Tag <> "" And txtƵ��.Text = strSequence And txtƵ��.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txtƵ��.Text = "" Then
            If cmdƵ��.Enabled And cmdƵ��.Visible Then cmdƵ��_Click
        Else
            int��Χ = 1 '��ѡƵ��
            strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
                " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ" & _
                " From ����Ƶ����Ŀ A Where A.���÷�Χ=" & int��Χ & _
                " And (A.���� Like '" & UCase(txtƵ��.Text) & "%'" & _
                " Or Upper(A.����) Like '" & mstrLike & UCase(txtƵ��.Text) & "%'" & _
                " Or Upper(A.����) Like '" & mstrLike & UCase(txtƵ��.Text) & "%'" & _
                " Or Upper(A.Ӣ������) Like '" & mstrLike & UCase(txtƵ��.Text) & "%')" & _
                " Order by A.����"
            vRect = GetControlRect(txtƵ��.Hwnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "����Ƶ��", , , , , , True, vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ�������Ƶ����Ŀ��", vbInformation, gstrSysName
                End If
                txtƵ��.Text = strSequence
                Call zlControl.TxtSelAll(txtƵ��)
                txtƵ��.SetFocus: Exit Sub
            End If
            Me.cmdƵ��.Tag = rsTmp("����"): Me.txtƵ�� = Me.cmdƵ��.Tag: strSequence = Me.cmdƵ��.Tag
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtƵ��_Validate(Cancel As Boolean)
    If cmdƵ��.Tag <> "" And txtƵ��.Text <> strSequence Then
        txtƵ��.Text = strSequence
    End If
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    On Error Resume Next
    If zlCommFun.ActualLen(txtҽ������.Text) > txtҽ������.MaxLength Then
        MsgBox "�������ݲ������� " & txtҽ������.MaxLength \ 2 & " �����ֻ� " & txtҽ������.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtҽ������.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txtҽ������_DblClick()
    If cmdSel.Visible And cmdSel.Enabled Then cmdSel_Click
End Sub

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txtҽ������)
    End If
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtҽ������.Text = "" Then cmdSel_Click: Exit Sub
        If txtҽ������.Text = txtҽ������.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        With txtҽ������
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
        '����Ŀ��¼��
        
        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        If AdviceInput(rsTmp) Then
            '��ʾ��ȱʡ���õ�ֵ
            txtҽ������.Tag = txtҽ������.Text
            txt����.Tag = txt����.Text
            If mblnCheck Then ReadCheckData
            
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            txt����.Text = txt����.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txtҽ������.Text <> txtҽ������.Tag Then
        txtҽ������.Text = txtҽ������.Tag
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(Me.txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt����) Then Me.txt���� = 1: Exit Sub
    Me.txt���� = CInt(Me.txt����)
    If CInt(Me.txt����) < 1 Then Me.txt���� = 1
End Sub

Private Function Check��ʼʱ��(ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ��������Ŀ�ʼʱ���Ƿ�Ϸ�
'˵����
'1.��ʼʱ�䲻��С�ڲ��˵���Ժʱ��
'2.��ʼʱ�����С����ֹʱ��
'3.����¼��ʱ,��ʼʱ�䲻��С�ڵ�ǰʱ��֮ǰ30����(�Ӷ�������ɿ���ʱ����ڿ�ʼʱ��30����)
'4.��¼��ҽ����ʼʱ�䲻�ܴ��ڵ�ǰʱ��
    Dim strInDate As String
    
    If Not IsDate(strStart) Then
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
        
    strInDate = Format(PatientDate, "yyyy-MM-dd HH:mm")
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵�" & IIf(PatientType = 2, "��Ժ", "����") & "ʱ�� " & strInDate & " ��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
'    If IsDate(strEnd) Then
'        If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(strEnd, "yyyy-MM-dd HH:mm") Then
'            strMsg = "ҽ���Ŀ�ʼִ��ʱ�����С��ִ����ֹʱ�䡣"
'            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    If DateDiff("n", CDate(strStart), zlDatabase.Currentdate) > 30 Then
        strMsg = "��ʼִ��ʱ�䲻��̫���ڵ�ǰʱ�䡣"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check��ʼʱ�� = True
End Function
Private Function SelectDiagItem() As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID "
    Select Case ItemType
        Case 1 'PACS
            strSQL = strSQL + "From ������ĿĿ¼ A,Ӱ������Ŀ B,������Ŀ���� C,����ִ�п��� D Where A.ID=B.������ĿID And A.ID=C.������ĿID And A.ID=D.������ĿID And D.ִ�п���ID=" & ItemDeptID
        Case 2 'LIS
            strSQL = strSQL + "From ������ĿĿ¼ A,������Ŀ���� C,����ִ�п��� D Where A.ID=C.������ĿID And A.ID=D.������ĿID And A.���='C' And D.ִ�п���ID=" & ItemDeptID
    End Select
    strSQL = strSQL + " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN(" & PatientType & ",3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And (A.���� Like '" + txtҽ������ + "%' Or Upper(A.����) Like '" + mstrLike + txtҽ������ + "%' Or Upper(C.����) Like '" + mstrLike + UCase(txtҽ������) + "%')"
            
    With txtҽ������
        Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "ѡ��������Ŀ", True, .Text, "", True, True, True, .Left + Me.picAdvice.Left + Me.Left, .Top + Me.picAdvice.Top + Me.Top, .Height, False, True)
    End With
End Function

Private Function SelectCap(Optional ByVal lngItemId As Long = 0, Optional ByVal QryStr As String = "", Optional blnNotSelect As Boolean = False) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
    If Len(QryStr) > 0 Then
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A,������Ŀ���� C,�����÷����� D Where A.ID=C.������ĿID And A.ID=D.�÷�ID" + _
            " And A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
            " And D.��ĿID=" & lngItemId & _
            " And (A.���� Like '" + QryStr + "%' Or Upper(A.����) Like '" + mstrLike + QryStr + "%' Or Upper(C.����) Like '" + mstrLike + UCase(QryStr) + "%')"
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.����,A.���� " + _
                "From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID" + _
                " And A.���='E' And A.��������='6'" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
                IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
                " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
                " And (A.���� Like '" + QryStr + "%' Or Upper(A.����) Like '" + mstrLike + QryStr + "%' Or Upper(C.����) Like '" + mstrLike + UCase(QryStr) + "%')"
        End If
    Else
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A,�����÷����� D Where A.ID=D.�÷�ID" + _
            " And A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
            " And D.��ĿID=" & lngItemId
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.����,A.���� " + _
                "From ������ĿĿ¼ A Where " + _
                " A.���='E' And A.��������='6'" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
                IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
                " And Nvl(A.ִ��Ƶ��,0) IN(0,1)"
        End If
    End If
    If blnNotSelect Then
        If rsTmp.State = adStateOpen Then rsTmp.Close: Set rsTmp = New ADODB.Recordset
        OpenRecord rsTmp, strSQL, Me.Caption
        If Not rsTmp.EOF Then Set SelectCap = rsTmp
    Else
        tmpRect = GetControlRect(Me.txt�ɼ�.Hwnd)
        Set SelectCap = zlDatabase.ShowSelect(Me, strSQL, 0, "�ɼ���ʽ", True, , , , , True, _
            tmpRect.Left, tmpRect.Top, Me.txt�ɼ�.Height, , , True)
    End If
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'���أ�����¼���Ƿ���Ч
    Dim str���� As String, blnGroup As Boolean, i As Long
    Dim lng�÷�ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String

    On Error GoTo errH

    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    txtҽ������.Text = rsInput!���� '��ʱ��ʾ

    '��Ҫ����������ݵ�һЩ��Ŀ
    '---------------------------------------------------------------------------------------------------------------
    intTmpType = -1
    If rsInput!���ID = "D" And zlCommFun.Nvl(GetItemField(rsInput!������ĿID, "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intTmpType = 0
        strHelpText = "��鲿λ"
    ElseIf rsInput!���ID = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intTmpType = 1
        strHelpText = "��������������ʽ"
    ElseIf InStr(",7,8,", rsInput!���ID) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intTmpType = 2
    ElseIf rsInput!���ID = "C" Then
        '������Ŀѡ�����걾
        intTmpType = 4
        strHelpText = "������Ŀ"
    End If

    If intTmpType <> -1 Then
        frmAdviceEditEx.mlngHwnd = Me.cboִ�п���.Hwnd ' txtҽ������.Hwnd
        frmAdviceEditEx.mintType = intTmpType
        frmAdviceEditEx.mint��Ч = 1
        frmAdviceEditEx.mstr�Ա� = mstr�Ա�
        frmAdviceEditEx.mlng��ĿID = IIf(intTmpType = 4, 0, rsInput!������ĿID)
        frmAdviceEditEx.mstrExtData = IIf(intTmpType = 4, rsInput!������ĿID & ";", "") '��������Ŀ
        frmAdviceEditEx.mint������� = PatientType

        On Error Resume Next
        frmAdviceEditEx.Show 1, Me
        On Error GoTo errH

        If Not frmAdviceEditEx.mblnOK Then Exit Function
        If frmAdviceEditEx.mstrExtData = "" Or Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" Then Exit Function
        
        If rsInput!���ID = "D" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txtҽ������.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str�걾��λ = Trim(rsInput("�걾��λ"))
            
            '������ϲ�λ��
            Call AdviceSet�������(1, strExtData)
            txtҽ������.Text = Get�����������(1, rsInput!����)
            strAdviceText = Get�����������(1, rsInput!����)
            Me.txt���� = Get��λ����
        ElseIf rsInput!���ID = "F" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txtҽ������.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str�걾��λ = Trim(rsInput("�걾��λ"))
            
            '�����ĸ���������������Ŀ��
            Call AdviceSet�������(2, strExtData)
            txtҽ������.Text = Get�����������(2, rsInput!����)
            strAdviceText = Get�����������(2, rsInput!����)
            Me.txt���� = Get��������
        ElseIf rsInput!���ID = "C" And frmAdviceEditEx.mstrExtData <> "" Then
            '��ȡ�ɼ���ʽ
            Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
            If rsTmp Is Nothing Then
                MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
                Exit Function
            End If
            Me.cmd�ɼ�.Tag = rsTmp("ID")
            Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
            
            strAdviceText = txtҽ������.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str�걾��λ = Trim(rsInput("�걾��λ"))
            
            '������Ŀ
            strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
                "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
                "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
                "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID " + _
                "From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID " + _
                "And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                "And A.������� IN([1],3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
                IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
                " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
                " And A.ID=[2]"
            If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
            Set rsInput = OpenSQLRecord(strSQL, Me.Caption, PatientType, Split(Split(strExtData, ";")(0), ",")(0))
            
            Call AdviceSet�������(3, strExtData)
            txtҽ������.Text = Get�����������(2, "")
            strAdviceText = txtҽ������.Text & "(" & Split(strExtData, ";")(1) & ")"
            Me.txt���� = Split(strExtData, ";")(1)
            str�걾��λ = Me.txt����
        End If
    Else
        str�걾��λ = Trim(rsInput("�걾��λ"))
        txtҽ������.Text = txtҽ������.Text & "(" & str�걾��λ & ")"
        strAdviceText = txtҽ������.Text
        
        '������ϲ�λ��
        Call AdviceSet�������(1, "")
    End If
    
    '��ʼʱ��
    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("ִ�а���ID") = 1 Then
        Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ��.Enabled = False
    Else
        Me.lbl��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Visible = False
        Me.txt��ʼʱ��.Enabled = True
    End If
    
    '����Ƶ��
    If rsInput("ִ��Ƶ��ID") = 1 Then
        Me.txtƵ��.Enabled = False: Me.txtƵ�� = "һ����": Me.cmdƵ��.Enabled = False
    Else
        Me.txtƵ��.Enabled = True: Me.txtƵ�� = "": Me.cmdƵ��.Enabled = True
    End If
    strSequence = Me.txtƵ��
    
    '����
    Me.txt���� = "1": Me.lbl������λ.Caption = rsInput("���㵥λ")
    
    '����
    If (rsInput("ִ��Ƶ��ID") = 0 And InStr(",1,2,", rsInput("���㷽ʽID")) > 0) _
                    Or InStr(",5,6,", rsInput("���ID")) > 0 Then
        Me.txt����.Enabled = True: Me.txt���� = "": Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = rsInput("���㵥λ")
    Else
        Me.txt����.Enabled = False: Me.txt���� = "": Me.txt����.BackColor = Me.BackColor: Me.lbl������λ.Caption = "" ' rsInput("���㵥λ")
    End If
    
    'ִ�п���
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType, ItemDeptID)
    If rsTmp Is Nothing Then
        Me.cboִ�п���.Clear: Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cboִ�п���.Clear
        Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(0) = rsTmp("ID"): Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    Else
        Me.cboִ�п���.Clear
        Do While Not rsTmp.EOF
            Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(Me.cboִ�п���.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = True: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    End If
    
    '����ҽ��
    If cboҽ��.ListCount <= 0 Then
        MsgBox "��ǰ����������ҽ��,����!", vbInformation, gstrSysName
        Exit Function
    End If

    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '����������Ŀ������ʾ��ʽ
    
    str��� = rsInput("���ID"): lngClinicID = rsInput("������ĿID")
    int�Ƽ����� = rsInput("�Ƽ�����ID"): intִ������ = rsInput("ִ�п���ID"): strClinicName = IIf(intType = 4, Me.txtҽ������, rsInput("����"))
    
    AdviceInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LabsInput() As Boolean
'���ܣ��༭������Ŀ
'���أ�����¼���Ƿ���Ч
    Dim str���� As String, blnGroup As Boolean, i As Long
    Dim lng�÷�ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String, rsInput As New ADODB.Recordset

    On Error GoTo errH
    
    intTmpType = 4
    strHelpText = "������Ŀ"

    frmAdviceEditEx.mlngHwnd = Me.cboִ�п���.Hwnd ' txtҽ������.Hwnd
    frmAdviceEditEx.mintType = intTmpType
    frmAdviceEditEx.mint��Ч = 1
    frmAdviceEditEx.mstr�Ա� = mstr�Ա�
    frmAdviceEditEx.mlng��ĿID = 0 ' FileTypeID
    frmAdviceEditEx.mstrExtData = strExtData
    frmAdviceEditEx.mint������� = PatientType

    On Error Resume Next
    frmAdviceEditEx.Show 1, Me
    On Error GoTo errH

    If Not frmAdviceEditEx.mblnOK Then Exit Function
    If frmAdviceEditEx.mstrExtData = "" Or Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" Then Exit Function
    '��ȡ�ɼ���ʽ
    Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
    If rsTmp Is Nothing Then
        MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    Me.cmd�ɼ�.Tag = rsTmp("ID")
    Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
    
    strAdviceText = txtҽ������.Text
    strExtData = frmAdviceEditEx.mstrExtData

    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID " + _
        "From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID " + _
        "And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN([1],3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And A.ID=[2]"
    If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
    Set rsInput = OpenSQLRecord(strSQL, Me.Caption, PatientType, Split(Split(strExtData, ";")(0), ",")(0))
    
    Call AdviceSet�������(3, strExtData)
    txtҽ������.Text = Get�����������(2, "")
    strAdviceText = txtҽ������.Text & "(" & Split(strExtData, ";")(1) & ")"
    Me.txt���� = Split(strExtData, ";")(1)
    str�걾��λ = Me.txt����
    
    '��ʼʱ��
    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("ִ�а���ID") = 1 Then
        Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ��.Enabled = False
    Else
        Me.lbl��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Visible = False
        Me.txt��ʼʱ��.Enabled = True
    End If
    
    '����Ƶ��
    If rsInput("ִ��Ƶ��ID") = 1 Then
        Me.txtƵ��.Enabled = False: Me.txtƵ�� = "һ����": Me.cmdƵ��.Enabled = False
    Else
        Me.txtƵ��.Enabled = True: Me.txtƵ�� = "": Me.cmdƵ��.Enabled = True
    End If
    strSequence = Me.txtƵ��
    
    '����
    Me.txt���� = "1": Me.lbl������λ.Caption = rsInput("���㵥λ")
    
    '����
    If (rsInput("ִ��Ƶ��ID") = 0 And InStr(",1,2,", rsInput("���㷽ʽID")) > 0) _
                    Or InStr(",5,6,", rsInput("���ID")) > 0 Then
        Me.txt����.Enabled = True: Me.txt���� = "": Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = rsInput("���㵥λ")
    Else
        Me.txt����.Enabled = False: Me.txt���� = "": Me.txt����.BackColor = Me.BackColor: Me.lbl������λ.Caption = "" ' rsInput("���㵥λ")
    End If
    
    'ִ�п���
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType, ItemDeptID)
    If rsTmp Is Nothing Then
        Me.cboִ�п���.Clear: Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cboִ�п���.Clear
        Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(0) = rsTmp("ID"): Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    Else
        Me.cboִ�п���.Clear
        Do While Not rsTmp.EOF
            Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(Me.cboִ�п���.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = True: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    End If
    
    '����ҽ��
    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '����������Ŀ������ʾ��ʽ
    
    str��� = rsInput("���ID"): lngClinicID = rsInput("������ĿID")
    int�Ƽ����� = rsInput("�Ƽ�����ID"): intִ������ = rsInput("ִ�п���ID"): strClinicName = IIf(intType = 4, Me.txtҽ������, rsInput("����"))
    
    LabsInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '���¼��벿λ�л򸽼������м�������Ŀ��
    If int���� = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    ElseIf int���� = 3 Then
        '���������Ŀ
        strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    End If
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
        "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п��� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        OpenRecord rsRelativeAdvice, strSQL, Me.Caption
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    strDate = IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "", Format(Me.txt��ʼʱ��, "yy��MM��dd��"))
    
    If rsRelativeAdvice Is Nothing Then
        If int���� = 1 Then
            Get����������� = txtMainAdvice & IIf(Len(str�걾��λ) = 0, "", "(" & str�걾��λ & ")"): Exit Function
        Else
            Get����������� = IIf(Len(strDate) = 0, "", strDate & " �� ") & txtMainAdvice & IIf(Len(str�걾��λ) = 0, "", "(" & str�걾��λ & ")"): Exit Function
        End If
    End If
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If int���� = 1 Then
            If Len(Trim(rsRelativeAdvice("�걾��λ"))) > 0 Then
                strTmp = strTmp & "," & rsRelativeAdvice("�걾��λ")
            End If
        ElseIf Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") = "G" Then
                str���� = rsRelativeAdvice("����")
            Else
                strTmp = strTmp & "," & rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If int���� = 1 Then
        If strTmp <> "" Then
            Get����������� = txtMainAdvice & "(" & Mid(strTmp, 2) & ")"
        Else
            Get����������� = txtMainAdvice
        End If
    Else
        If strTmp <> "" Or str���� <> "" Then
            If str���� <> "" Then
                Get����������� = IIf(Len(strDate) = 0, "", strDate & " ") & "�� " & str���� & " ���� " & txtMainAdvice
            Else
                Get����������� = IIf(Len(strDate) = 0, "", strDate & " �� ") & txtMainAdvice
            End If
            If strTmp <> "" Then
                Get����������� = Get����������� & " �� " & Mid(strTmp, 2)
            End If
        Else
            Get����������� = IIf(Len(strDate) = 0, "", strDate & " �� ") & txtMainAdvice
        End If
    End If
End Function

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") <> "G" Then
                strTmp = strTmp & "," & rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get����������� = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " �� ") & Mid(strTmp, 2)
    Else
        Get����������� = txtMainAdvice
    End If
End Function

Private Function Get��������() As String
    If rsRelativeAdvice Is Nothing Then Get�������� = "": Exit Function
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") = "G" Then
                Get�������� = rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
End Function

Private Function Get��λ����() As String
    If rsRelativeAdvice Is Nothing Then Get��λ���� = "": Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("�걾��λ"))) > 0 Then
            Get��λ���� = Get��λ���� & "," & rsRelativeAdvice("�걾��λ")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(Get��λ����) > 0 Then Get��λ���� = Mid(Get��λ����, 2)
End Function

Private Function GetExeDepart(ByVal lngDiagItem As Long, ByVal iPatientType As Integer, Optional ByVal lngDepartID As Long = 0) As ADODB.Recordset
'���ܣ���ȡִ�п���
'   iPatientType���������� 1=���2=סԺ
'   lngDepartID����������
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo DBError
    
    If lngDepartID = 0 Then lngDepartID = UserInfo.����ID
    
    zlDatabase.OpenRecordset rsTmp, "Select B.ID,B.����,B.���� From ���ű� B Where B.ID=" & lngDepartID & " Order by B.����", Me.Caption
    
    If Not rsTmp.EOF Then Set GetExeDepart = rsTmp
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupCount(lng���ID As Long) As Long
'���ܣ���ȡ�����Ŀ�е���Ŀ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(*) as NUM From ������Ŀ��� Where �������ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng���ID)
    If Not rsTmp.EOF Then GetGroupCount = zlCommFun.Nvl(rsTmp!Num, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Getȱʡ�÷�ID(int���� As Integer) As Long
'���ܣ�����ȱʡ�ĸ�ҩ;������ҩ�巨
'������int����=2-��ҩ;��,3-��ҩ�巨,4-��ҩ�÷�
'      str�Ա�=�����Ա�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From ������ĿĿ¼" & _
        " Where ���='E' And ��������=[1]" & _
        " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
        " Order by ����"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int����)
    If Not rsTmp.EOF Then Getȱʡ�÷�ID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemField(ByVal lng��ĿID As Long, ByVal strField As String) As Variant
'���ܣ���ȡָ��������Ŀ��ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get����ҽ��(ByVal lng����id As Long, ByVal bln��ʿվ As Boolean, strȱʡҽ�� As String, lngҽ��ID As Long, _
    Optional objCbo As Object, Optional ByVal int��Χ As Integer = 2) As Boolean
'���ܣ���ȡ���õĿ���ҽ����ָ������������
'������lng���˿���ID=�������ڿ���ID
'      bln��ʿվ=�Ƿ��ɻ�ʿ��ҽ����ҽ��
'      objCbo=Ҫ����ҽ���嵥��������
'      strȱʡҽ��=ȱʡ��λ��ҽ��,�������objCbo,�������ȶ�λ,�ٷ���ȱʡҽ����ҽ��ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
    
    If bln��ʿվ Then
        '�������ڿ��ҵ�ҽ��
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID=" & lng���˿���ID & _
            " Order by A.����"
        '�������ڲ������Ƶ�ҽ��
        strSQL = "Select Distinct ����ID From ��λ״����¼ Where ����ID=" & lng���˿���ID
        strSQL = "Select Distinct ����ID From ��λ״����¼ Where ����ID=(" & strSQL & ")"
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID IN(" & strSQL & ")" & _
            " Order by A.����"
        'ȫԺסԺ���ҵ�ҽ��
        strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(" & int��Χ & ",3)"
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID IN(" & strSQL & ")" & _
            " Order by A.����"
    Else 'ҽ����ҽ��ʱ,����Ϊֻ��Ϊҽ������
        strSQL = "Select ID,���,����,���� From ��Ա�� Where ID=" & UserInfo.ID
    End If

    OpenRecord rsTmp, strSQL, "zlCISCore"
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln��ʿվ Then
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            ElseIf bln��ʿվ Then
                If strȱʡҽ�� <> "" Then
                    'ȱʡҽ��(סԺҽʦ)����
                    rsTmp.Filter = "����='" & strȱʡҽ�� & "'"
                Else
                    '���˿��ҵ�ҽ������
                    rsTmp.Filter = "����ID=" & lng���˿���ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            End If
        End If
    Else
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem zlCommFun.Nvl(rsTmp!����) & "-" & rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!���� = strȱʡҽ�� Then
                Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
            End If
            rsTmp.MoveNext
        Next
    End If
    Get����ҽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get��������ID(ByVal lngҽ��ID As Long, ByVal lng���˿���ID As Long, Optional ByVal int��Χ As Integer = 2) As Long
'���ܣ���ҽ��ȷ����������
'������int��Χ=1-����,2-סԺ(ȱʡ)
'˵������ҽ���������ҷ�Χ��,����˳�����£�
'      1�����˿���
'      2������������/סԺ���˵Ŀ�����ΪĬ�Ͽ���
'      3������������/סԺ���˵Ŀ���
'      4��Ĭ�Ͽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr����ID(1 To 4) As Long
    
    '���ܲ���û������
    strSQL = "Select Distinct C.����,A.����ID,Nvl(A.ȱʡ,0) as ȱʡ,Nvl(B.�������,0) as �������" & _
        " From ������Ա A,��������˵�� B,���ű� C" & _
        " Where A.����ID=C.ID And A.����ID=B.����ID(+) And A.��ԱID=[1]" & _
        " Order by C.����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ID = lng���˿���ID Then
            arr����ID(1) = rsTmp!����ID
        ElseIf InStr("," & int��Χ & ",3,", rsTmp!�������) > 0 And rsTmp!ȱʡ = 1 Then
            arr����ID(2) = rsTmp!����ID
        ElseIf InStr("," & int��Χ & ",3,", rsTmp!�������) > 0 Then
            If arr����ID(3) = 0 Then arr����ID(3) = rsTmp!����ID
        ElseIf rsTmp!ȱʡ = 1 Then
            arr����ID(4) = rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    For i = LBound(arr����ID) To UBound(arr����ID)
        If arr����ID(i) <> 0 Then
            Get��������ID = arr����ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'===����Ϊ������Ϣ
Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo�ѱ�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
    
    If SendMessage(cbo�ѱ�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�ѱ�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
    If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
End Sub

Private Sub cbo���ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���ʽ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo���ʽ.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo���ʽ.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo�Ա�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�Ա�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    If cbo�Ա�.ListIndex = -1 And cbo�Ա�.ListCount > 0 Then cbo�Ա�.ListIndex = 0
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboְҵ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboְҵ.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cboְҵ.ListIndex = lngIdx
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmd��λ����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = zlDatabase.ShowSelect(Me, _
            " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID", _
            2, "��λ", , txt��λ����.Text)
    If Not rsTmp Is Nothing Then
        txt��λ����.Tag = rsTmp!ID
        txt��λ����.Text = rsTmp!����
        txt��λ����.SelStart = Len(txt��λ����.Text)
    End If
    txt��λ����.SetFocus
End Sub

Private Sub cmd��ͥ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = zlDatabase.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , txt��ͥ��ַ.Text)
    If Not rsTmp Is Nothing Then
        txt��ͥ��ַ.Text = rsTmp!����
        txt��ͥ��ַ.SelStart = Len(txt��ͥ��ַ.Text)
    End If
    txt��ͥ��ַ.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
        DoEvents
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Function InitData() As Boolean
'���ܣ���ʼ����Ҫ����
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '�Ա�
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("�Ա�")
    cbo�Ա�.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '�ѱ�
    Init�ѱ� True

    'ҽ�Ƹ��ʽ
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("ҽ�Ƹ��ʽ")
    cbo���ʽ.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo���ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo���ʽ.ItemData(cbo���ʽ.NewIndex) = 1
                cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '����
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("����")
    cbo����.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo����.ItemData(cbo����.NewIndex) = 1
                cbo����.ListIndex = cbo����.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '����
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("����")
    cbo����.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo����.ItemData(cbo����.NewIndex) = 1
                cbo����.ListIndex = cbo����.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '����״��
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("����״��")
    cbo����.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo����.ItemData(cbo����.NewIndex) = 1
                cbo����.ListIndex = cbo����.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    'ְҵ
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("ְҵ")
    cboְҵ.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cboְҵ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cboְҵ.ItemData(cboְҵ.NewIndex) = 1
                cboְҵ.ListIndex = cboְҵ.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '��ʼ��������
    InitDepts
    
    InitData = True
End Function

Private Function Init�ѱ�(bln���� As Boolean, Optional blnKeepIndex As Boolean) As Boolean
'������bln����=�Ƿ�������޳������Ŀ
'      blnKeepIndex=�Ƿ񱣳�ԭ�еķѱ�ѡ��
    Dim strSQL As String, i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strKeep As String
    
    On Error GoTo errH
    
    strKeep = cbo�ѱ�.Text
    
    '�ѱ�:���Ψһ����Ŀ(������ȱʡ�ѱ�),�����ǳ���,������Ч�ڼ估����
    strSQL = "Select ����,����,����," & _
        " Nvl(���޳���,0) as ����,Nvl(ȱʡ��־,0) as ȱʡ" & _
        " From �ѱ� Where ����=1" & IIf(Not bln����, " And Nvl(���޳���,0)=0", "") & _
        " Order by ����"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL) 'SQLTest
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    cbo�ѱ�.Clear
    Do While Not rsTmp.EOF
        cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            If cbo�ѱ�.ListIndex = -1 Then
                cbo�ѱ�.ItemData(cbo�ѱ�.NewIndex) = 1
                cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            End If
        End If
        
        '����ԭ�зѱ�ѡ��
        If blnKeepIndex Then
            If strKeep = rsTmp!���� & "-" & rsTmp!���� Then
                cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            End If
        End If
        
        '��¼������Ŀ:�����Ǳ���ȱʡ��ϵͳȱʡ
        If rsTmp!���� = 1 Then
            cbo�ѱ�.ItemData(cbo�ѱ�.NewIndex) = 2
        End If
        rsTmp.MoveNext
    Loop
    
    Init�ѱ� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt��λ�绰_GotFocus()
    zlControl.TxtSelAll txt��λ�绰
End Sub

Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ�绰, KeyAscii
End Sub

Private Sub txt��λ����_GotFocus()
    zlControl.TxtSelAll txt��λ����
    zlCommFun.OpenIme True
End Sub

Private Sub txt��λ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd��λ����.Enabled And cmd��λ����.Visible Then cmd��λ����_Click
End Sub

Private Sub txt��λ����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ����, KeyAscii
End Sub

Private Sub txt��λ����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��λ�ʱ�
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��λ�ʱ�, KeyAscii
End Sub

Private Sub txt��ͥ��ַ_GotFocus()
    zlControl.TxtSelAll txt��ͥ��ַ
    zlCommFun.OpenIme True
End Sub

Private Sub txt��ͥ��ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd��ͥ��ַ.Enabled And cmd��ͥ��ַ.Visible Then cmd��ͥ��ַ_Click
End Sub

Private Sub txt��ͥ��ַ_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��ͥ��ַ, KeyAscii
End Sub

Private Sub txt��ͥ��ַ_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    zlControl.TxtSelAll txt��ͥ�绰
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��ͥ�绰, KeyAscii
End Sub

Private Sub txt��ͥ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��ͥ�ʱ�
End Sub

Private Sub txt��ͥ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��ͥ�ʱ�, KeyAscii
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt�����, KeyAscii
End Sub

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt����, KeyAscii
End Sub

Private Sub txt���֤��_GotFocus()
    zlControl.TxtSelAll txt���֤��
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt���֤��, KeyAscii
End Sub

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    zlCommFun.OpenIme True
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
        CheckLen txt����, KeyCode
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strField As String
    
    If Len(Trim(txt����)) = 0 Then Exit Sub
    If txt���� = txt����.Tag Then Exit Sub
    
    Set rsTmp = GetPatient(txt����)
    If rsTmp.EOF Then
        Me.txt��λ�绰 = ""
        Me.txt��λ���� = ""
        Me.txt��λ�ʱ� = ""
        Me.txt��ͥ��ַ = ""
        Me.txt��ͥ�绰 = ""
        Me.txt��ͥ�ʱ� = ""
        Me.txt����� = ""
        Me.txt���� = ""
        Me.txt���֤�� = ""
        If InStr("+-*.", Left(Me.txt����.Text, 1)) > 0 Then Cancel = True 'Me.txt����.Text = "": Me.txt����.SetFocus
        
        PatientID = 0: PatientType = 1: CheckID = "": lng��ʶ�� = 0
        
        DTBirth.Value = Empty
    Else
        On Error Resume Next
        Me.txt����.Text = Nvl(rsTmp("����"))
        Me.txt��λ�绰 = Nvl(rsTmp("��λ�绰"))
        Me.txt��λ���� = Nvl(rsTmp("������λ"))
        Me.txt��λ�ʱ� = Nvl(rsTmp("��λ�ʱ�"))
        Me.txt��ͥ��ַ = Nvl(rsTmp("��ͥ��ַ"))
        Me.txt��ͥ�绰 = Nvl(rsTmp("��ͥ�绰"))
        Me.txt��ͥ�ʱ� = Nvl(rsTmp("�����ʱ�"))
        Me.txt����� = Nvl(rsTmp("�����"))
        Me.txt���� = Nvl(rsTmp("����"))
        Me.txt���֤�� = Nvl(rsTmp("���֤��"))
        Me.cbo�ѱ�.ListIndex = CombIndex(cbo�ѱ�, Nvl(rsTmp("�ѱ�")))
        Me.cbo���ʽ.ListIndex = CombIndex(cbo���ʽ, Nvl(rsTmp("ҽ�Ƹ��ʽ")))
        Me.cbo����.ListIndex = CombIndex(cbo����, Nvl(rsTmp("����")))
        Me.cbo����.ListIndex = CombIndex(cbo����, Nvl(rsTmp("����״��")))
        Me.cbo����.ListIndex = CombIndex(cbo����, Nvl(rsTmp("����")))
        Me.cbo�Ա�.ListIndex = CombIndex(cbo�Ա�, Nvl(rsTmp("�Ա�")))
        Me.cboְҵ.ListIndex = CombIndex(cboְҵ, Nvl(rsTmp("ְҵ")))
        
        PatientID = Nvl(rsTmp("����ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1): CheckID = Nvl(rsTmp("��ҳID"))
        lng��ʶ�� = Nvl(rsTmp("��ʶ��"), 0)
        '����Ĭ�Ͽ������ҡ�ҽ��
        For i = 0 To Me.cbo��������.ListCount - 1
            If Me.cbo��������.ItemData(i) = Nvl(rsTmp("���˿���"), 0) Then
                Me.cbo��������.ListIndex = i
                Exit For
            End If
        Next
        DoEvents
        strField = ""
        strField = rsTmp.Fields("ҽ��").Name
        If strField = "ҽ��" Then
            Me.cboҽ��.Text = Nvl(rsTmp("ҽ��"))
            For i = 0 To Me.cboҽ��.ListCount - 1
                If Me.cboҽ��.List(i) Like "*-" & Nvl(rsTmp("ҽ��")) Then
                    Me.cboҽ��.ListIndex = i
                    Exit For
                End If
            Next
        End If
    
        If IsNull(rsTmp!��������) Then
            DTBirth.Value = Empty
        Else
            DTBirth.Value = rsTmp!��������
        End If
    End If
    'Ӣ����
    If mblnCheck Then txtItem(3).Text = UCase(Replace(zlCommFun.mGetFullPY(Trim(txt����)), vbCrLf, ""))
    If mblnCheck Then txtItem(7).Text = Next����(lngClinicID, PatientID, chkUnicode.Value = 1)

    txt����.Tag = txt����.Text
End Sub

Private Function CombIndex(objComboBox As Object, ByVal strText As String) As Integer
    Dim i As Integer
    CombIndex = 0
    For i = 0 To objComboBox.ListCount - 1
        With objComboBox
            If .List(i) Like "*-" & strText Then CombIndex = i: Exit For
        End With
    Next
End Function

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Function GetPatient(strCode As String) As ADODB.Recordset
'���ܣ���ȡ������Ϣ������ʾ�ò��˴��ڵ�ҽ��ʱ��
    Dim strSQL As String, i As Long
    Dim strNO As String, str���� As String, lng����id As Long
    Dim strSeek As String
    
    On Error GoTo errH
    
    sCheckNo = ""
    strSeek = strCode
    '�жϵ�ǰ����ģʽ
    If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) And iInputType = -1 Then 'ˢ��
        iInputType = 0
    ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
        iInputType = 1
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        iInputType = 2
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����
        iInputType = 3
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '�Һŵ�
        iInputType = 4
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "#" Then '�շѵ��ݺ�
        iInputType = 5
        strSeek = Mid(strCode, 2)
    ElseIf iInputType = -1 Then '��������
        iInputType = 6
    End If
    
    If iInputType = 0 Then 'ˢ��
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,Nvl(A.סԺ����,0) As ��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���," & _
            "Decode(A.��ǰ����id,Null,A.�����,A.סԺ��) As ��ʶ��,A.*" & _
            " From ������Ϣ A,���˹Һż�¼ B Where A.���￨��=[1] And A.����ID=B.����ID(+) And A.�����=B.�����(+)" & _
            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 1 Then '����ID
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,Nvl(A.סԺ����,0) As ��ҳID,Nvl(A.��ǰ����id,0) As ���˿���," & _
            "Decode(A.��ǰ����id,Null,A.�����,A.סԺ��) As ��ʶ��,A.*" & _
            " From ������Ϣ A Where A.����ID=[2]"
    ElseIf iInputType = 2 Then 'סԺ��
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,Nvl(A.סԺ����,0) As ��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���," & _
            "Decode(A.��ǰ����id,Null,A.�����,A.סԺ��) As ��ʶ��,A.*" & _
            " From ������Ϣ A,���˹Һż�¼ B Where A.סԺ��=[2] And A.����ID=B.����ID(+) And A.�����=B.�����(+)" & _
            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 3 Then '�����
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,Nvl(A.סԺ����,0) As ��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���," & _
            "Decode(A.��ǰ����id,Null,A.�����,A.סԺ��) As ��ʶ��,A.*" & _
            " From ������Ϣ A,���˹Һż�¼ B Where A.�����=[2] And A.����ID=B.����ID(+) And A.�����=B.�����(+)" & _
            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 4 Then '�Һŵ�
        strNO = GetFullNO(strSeek, 12)
        strSQL = "Select Decode(B.��ҳID,Null,1,2) As PatientType,Nvl(B.��ҳID,0) As ��ҳID,Nvl(B.ִ�в���ID,0) As ���˿���," & _
            "Decode(B.��ҳID,Null,A.�����,A.סԺ��) As ��ʶ��,A.*" & _
            " From ������Ϣ A,���˷��ü�¼ B" & _
            " Where B.��¼����=4 And B.��¼״̬ IN(1,3) And B.NO=[3] And B.����ID=A.����ID"
    ElseIf iInputType = 5 Then '�շѵ��ݺ�
        strNO = GetFullNO(strSeek, 13)
        sCheckNo = strNO
        
        strSQL = "Select Decode(B.��ҳID,Null,1,2) As PatientType,Nvl(B.��ҳID,0) As ��ҳID,B.��������ID As ���˿���,B.������ As ҽ��,B.����,B.�Ա�,B.����," & _
            "A.����ID,A.��λ�绰,A.������λ,A.��λ�ʱ�,A.��ͥ��ַ,A.��ͥ�绰,A.�����ʱ�,A.�����,A.���֤��,A.�ѱ�,A.ҽ�Ƹ��ʽ," & _
            "A.����,A.����״��,A.����,A.ְҵ,Decode(B.��ҳID,Null,A.�����,A.סԺ��) As ��ʶ�� From ������Ϣ A,���˷��ü�¼ B" & _
            " Where B.��¼����=1 And B.��¼״̬ IN(1,3) And B.NO=[3] And B.����ID=A.����ID(+) And B.ҽ����� Is Null"
    Else '��������
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,Nvl(A.סԺ����,0) As ��ҳID,Nvl(A.��ǰ����id,0) As ���˿���," & _
            "Decode(A.��ǰ����id,Null,A.�����,A.סԺ��) As ��ʶ��,A.*" & _
            " From ������Ϣ A Where A.����=[1]"
    End If
    
    Set GetPatient = OpenSQLRecord(strSQL, Me.Caption, strCode, Val(strSeek), strNO)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=" & intNum
    Call OpenRecord(rsTmp, strSQL, "mdlPublic")
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlRegEvent", strSQL) 'SQLTest
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���') Or A.ID=" & ItemDeptID & " Or A.ID=" & UserInfo.����ID & ")" & _
        " Order by A.����"
    Me.cbo��������.Clear
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cbo��������.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    If cbo��������.ListCount > 0 Then cbo��������.ListIndex = 0
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cboҽ��.Clear
    
    '����ҽ����ʿ
    strSQL = _
        "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
        " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա���� IN('ҽ��') And B.����ID=[1]"
    strSQL = strSQL & " Order by ����,��Ա���� Desc"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ��.AddItem IIf(IsNull(rsTmp!����), "", rsTmp!���� & "-") & rsTmp!����
            cboҽ��.ItemData(cboҽ��.ListCount - 1) = rsTmp!����ID
            
            If rsTmp!ID = UserInfo.ID And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = cboҽ��.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboҽ��.ListCount = 1 And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = 0
    End If
End Sub

Private Function GetIndex(aSeekArray() As String, ByVal vSeekValue As Variant) As Integer
    Dim i As Integer
    For i = 0 To UBound(aSeekArray, 2)
        If aSeekArray(0, i) = vSeekValue Then Exit For
    Next
    If i > UBound(aSeekArray, 2) Then
        GetIndex = 0
    Else
        GetIndex = i
    End If
End Function

Public Function Next����(ByVal lngItemId As Long, ByVal lngPatientID As Long, Optional ByVal blnUnicode As Boolean = False) As Double
    Dim rsCtrl As New ADODB.Recordset
    Dim strSQL As String, lngNO As Double
    Dim lngExeDept As Long

ReStart:
    Err = 0
    On Error GoTo errH
    
    If Not blnUnicode Then '�������
        strSQL = "Select ���� From Ӱ�����¼ A,����ҽ����¼ B,Ӱ������Ŀ C" & _
            " Where A.ҽ��ID=B.ID And B.����ID=[1] And A.Ӱ�����+0=C.Ӱ����� And C.������ĿID=[2] Order By B.ͣ��ʱ�� Desc"
        Set rsCtrl = OpenSQLRecord(strSQL, "��ȡ����", lngPatientID, lngItemId)
        If Not rsCtrl.EOF Then
            lngNO = Val(Nvl(rsCtrl("����"), 0))
        Else
            strSQL = "Select A.������ From Ӱ������� A,Ӱ������Ŀ B Where A.����=B.Ӱ����� And B.������ĿID=[1]"
            Set rsCtrl = OpenSQLRecord(strSQL, "��ȡ����", lngItemId)
            If rsCtrl.EOF Then Exit Function
            
            lngNO = Val(Nvl(rsCtrl("������"), 0)) + 1
        End If
    Else '��ִ�п��ұ��
        lngExeDept = ItemDeptID
        strSQL = "Select A.���� From Ӱ�����¼ A,����ҽ����¼ B" & _
            " Where A.ҽ��ID=B.ID+0 And B.ִ�п���ID+0=[1] And B.����ID=[2] Order By B.ͣ��ʱ�� Desc"
        Set rsCtrl = OpenSQLRecord(strSQL, "��ȡ����", lngExeDept, lngPatientID)
        If rsCtrl.EOF Then 'ȡ����������
            strSQL = "SELECT DISTINCT C.����,Nvl(C.������,0) FROM Ӱ������Ŀ A,����ִ�п��� B,Ӱ������� C" & _
                " WHERE A.������ĿID=B.������ĿID+0 AND A.Ӱ�����||''=C.���� AND B.ִ�п���ID=[1] ORDER BY Nvl(C.������,0) DESC"
            Set rsCtrl = OpenSQLRecord(strSQL, "��ȡ����", lngExeDept)
            If rsCtrl.EOF Then
                Next���� = 0
                Exit Function
            End If
            
            lngNO = Val(Nvl(rsCtrl(1), 0)) + 1
        Else
            lngNO = Val(Nvl(rsCtrl("����"), 0))
        End If
    End If
    
    Next���� = lngNO
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReadCheckData()
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    
    strSQL = "Select Nvl(���в���,0) As ���в���,Nvl(�ɷ���Ƭ,0) As �ɷ���Ƭ,Nvl(Ӱ�����,'CT') As Ӱ����� From Ӱ������Ŀ" & _
        " Where ������ĿID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngClinicID)
    If Not rsTmp.EOF Then
        Select Case rsTmp("���в���")
            Case 0, 1
                chk����.Value = rsTmp("���в���"): chk����.Enabled = False
            Case Else
                chk����.Enabled = True
        End Select
        
        Select Case rsTmp("�ɷ���Ƭ")
            Case 0, 1
                chk��Ƭ.Value = rsTmp("�ɷ���Ƭ"): chk��Ƭ.Enabled = False
            Case Else
                chk��Ƭ.Enabled = True
        End Select
        
        strImageType = rsTmp("Ӱ�����")
    Else
        chk��Ƭ.Enabled = True
        chk����.Enabled = True
        
        strImageType = "CT"
    End If
    
    txtItem(7).Text = Next����(lngClinicID, PatientID, chkUnicode.Value = 1)
End Sub
