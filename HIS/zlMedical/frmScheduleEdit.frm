VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmScheduleEdit 
   Caption         =   "���ԤԼ����"
   ClientHeight    =   7500
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11580
   Icon            =   "frmScheduleEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ils32 
      Left            =   10365
      Top             =   4305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   6060
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":6FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":C036
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":C330
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":C8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":CE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":CFBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraGroup 
      Caption         =   "&1.���"
      Height          =   2280
      Left            =   225
      TabIndex        =   36
      Top             =   2865
      Width           =   2445
      Begin MSComctlLib.ListView lvwGroup 
         Height          =   1380
         Left            =   90
         TabIndex        =   37
         Top             =   300
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   2434
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils32"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   3
         Left            =   1920
         Picture         =   "frmScheduleEdit.frx":DE10
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F10"
         Top             =   1770
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   2
         Left            =   1545
         Picture         =   "frmScheduleEdit.frx":14662
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F9"
         Top             =   1770
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   0
         Left            =   1170
         Picture         =   "frmScheduleEdit.frx":1AEB4
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F8"
         Top             =   1770
         Width           =   345
      End
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   435
      ScaleHeight     =   555
      ScaleWidth      =   10650
      TabIndex        =   66
      Top             =   6375
      Width           =   10650
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8100
         TabIndex        =   55
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9315
         TabIndex        =   56
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   90
         TabIndex        =   57
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   645
      Left            =   -15
      TabIndex        =   60
      Top             =   -90
      Width           =   10635
      Begin VB.PictureBox picNo 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   8745
         ScaleHeight     =   315
         ScaleWidth      =   1815
         TabIndex        =   62
         Top             =   240
         Width           =   1815
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
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
            Height          =   210
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   30
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   37
            Left            =   30
            TabIndex        =   64
            Top             =   45
            Width           =   360
         End
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "�������ԤԼ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   61
         Top             =   255
         Width           =   2040
      End
   End
   Begin VB.Frame fraGroupInfo 
      Height          =   660
      Left            =   75
      TabIndex        =   58
      Top             =   1455
      Width           =   11190
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   10
         Left            =   10350
         Picture         =   "frmScheduleEdit.frx":21706
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F12"
         Top             =   225
         Width           =   345
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   8
         Left            =   8400
         TabIndex        =   28
         Text            =   "cfr@zlsoft.cn"
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   11
         Left            =   6000
         TabIndex        =   26
         Text            =   "1399090980"
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   12
         Left            =   4395
         TabIndex        =   24
         Text            =   "���˴�"
         Top             =   240
         Width           =   810
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   13
         Left            =   780
         TabIndex        =   21
         Text            =   "ĳĳ�������������ι�˾"
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   4
         Left            =   3405
         Picture         =   "frmScheduleEdit.frx":22548
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   210
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʼ�"
         Height          =   180
         Index           =   15
         Left            =   7605
         TabIndex        =   27
         Top             =   300
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgNew 
         Height          =   240
         Index           =   0
         Left            =   480
         Picture         =   "frmScheduleEdit.frx":22AD2
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ��"
         Height          =   180
         Index           =   14
         Left            =   3810
         TabIndex        =   23
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�绰"
         Height          =   180
         Index           =   12
         Left            =   5220
         TabIndex        =   25
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   90
         TabIndex        =   20
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.Frame fraSingle 
      Height          =   990
      Left            =   75
      TabIndex        =   0
      Top             =   465
      Width           =   11805
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   14
         Left            =   5010
         TabIndex        =   7
         Text            =   "90"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   14
         Left            =   10890
         Picture         =   "frmScheduleEdit.frx":2305C
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "����Ϣд��IC��"
         Top             =   225
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   15
         Left            =   10500
         Picture         =   "frmScheduleEdit.frx":298AE
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "��IC������Ϣ"
         Top             =   225
         Width           =   345
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   10
         Left            =   5190
         TabIndex        =   11
         Text            =   "1399090980"
         Top             =   615
         Width           =   1515
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   1
         Left            =   2205
         Picture         =   "frmScheduleEdit.frx":30100
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   11
         Left            =   11280
         Picture         =   "frmScheduleEdit.frx":3068A
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F11"
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   3
         Left            =   3225
         TabIndex        =   5
         Text            =   "90"
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   7
         Left            =   7740
         TabIndex        =   19
         Text            =   "cfr@zlsoft.cn"
         Top             =   600
         Width           =   1800
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   4
         Left            =   7545
         TabIndex        =   9
         Text            =   "123456789012345678901"
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   9
         Left            =   2205
         TabIndex        =   15
         Text            =   "90"
         Top             =   615
         Width           =   510
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   5
         Left            =   750
         TabIndex        =   2
         Text            =   "��ĳĳ"
         Top             =   255
         Width           =   1425
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   3255
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   615
         Width           =   1110
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   16
         Left            =   4425
         TabIndex        =   6
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��ϵ�绰"
         Height          =   180
         Index           =   4
         Left            =   4410
         TabIndex        =   10
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʼ�"
         Height          =   180
         Index           =   10
         Left            =   6960
         TabIndex        =   18
         Top             =   690
         Width           =   720
      End
      Begin VB.Image imgNew 
         Height          =   240
         Index           =   1
         Left            =   465
         Picture         =   "frmScheduleEdit.frx":314CC
         Top             =   105
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Index           =   7
         Left            =   6765
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   6
         Left            =   2745
         TabIndex        =   16
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   11
         Left            =   1800
         TabIndex        =   14
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   8
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   12
         Top             =   690
         Width           =   540
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   615
      Left            =   285
      TabIndex        =   59
      Top             =   2115
      Width           =   10635
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   5655
         TabIndex        =   35
         Text            =   "ĳĳ�������������ι�˾"
         Top             =   210
         Width           =   4890
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   3075
         TabIndex        =   33
         Top             =   210
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1170
         TabIndex        =   31
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��ַ(&A)"
         Height          =   180
         Index           =   3
         Left            =   4965
         TabIndex        =   34
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�绰(&T)"
         Height          =   180
         Index           =   2
         Left            =   2400
         TabIndex        =   32
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ԥ Լ ��(&L)"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   30
         Top             =   270
         Width           =   990
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   67
      Top             =   7140
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmScheduleEdit.frx":31A56
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2160
      Top             =   6825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tbs 
      Height          =   2580
      Left            =   3345
      TabIndex        =   41
      Top             =   2730
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   4551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   635
      WordWrap        =   0   'False
      TabCaption(0)   =   "&4.�����Ŀ"
      TabPicture(0)   =   "frmScheduleEdit.frx":322EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(18)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(19)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfPrice"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vsf"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd(18)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd(17)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtSum(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSum(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSum(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&5.�ܼ���Ա"
      TabPicture(1)   =   "frmScheduleEdit.frx":32306
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfPerson"
      Tab(1).Control(1)=   "cmd(8)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmd(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmd(13)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd(16)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   0
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   435
         Width           =   930
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   1
         Left            =   3135
         MaxLength       =   16
         TabIndex        =   78
         Top             =   435
         Width           =   870
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   2
         Left            =   4680
         MaxLength       =   16
         TabIndex        =   77
         Top             =   435
         Width           =   1020
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   17
         Left            =   4620
         Picture         =   "frmScheduleEdit.frx":32322
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "ȫ������"
         Top             =   720
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   18
         Left            =   4230
         Picture         =   "frmScheduleEdit.frx":38B74
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   "ȫ���շ�"
         Top             =   720
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   16
         Left            =   -70035
         Picture         =   "frmScheduleEdit.frx":3F3C6
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "��λ��Աѡ��"
         Top             =   1275
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   12
         Left            =   -70665
         Picture         =   "frmScheduleEdit.frx":45C18
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "����Ϣд��IC��"
         Top             =   795
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   13
         Left            =   -70545
         Picture         =   "frmScheduleEdit.frx":4C46A
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "��IC������Ϣ"
         Top             =   1260
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   9
         Left            =   -69630
         Picture         =   "frmScheduleEdit.frx":52CBC
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "������Ա(F7)"
         Top             =   825
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   8
         Left            =   -70080
         Picture         =   "frmScheduleEdit.frx":57D16
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "��������(F6)"
         Top             =   810
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   5
         Left            =   5025
         Picture         =   "frmScheduleEdit.frx":58B58
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "��ѡ����ݼ���F3"
         Top             =   705
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsfPerson 
         Height          =   1545
         Left            =   -74655
         TabIndex        =   45
         Top             =   450
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   2725
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   1530
         Left            =   165
         TabIndex        =   42
         Top             =   795
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2699
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   6
         Left            =   5475
         Picture         =   "frmScheduleEdit.frx":590E2
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "�������ѡ��F4"
         Top             =   720
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsfPrice 
         Height          =   1635
         Left            =   2985
         TabIndex        =   69
         Top             =   855
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2884
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����۸�(&B)"
         Height          =   180
         Index           =   19
         Left            =   150
         TabIndex        =   82
         Top             =   495
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���۸�(E)"
         Height          =   180
         Index           =   18
         Left            =   2130
         TabIndex        =   81
         Top             =   495
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۿ�(Z)"
         Height          =   180
         Index           =   17
         Left            =   4035
         TabIndex        =   80
         Top             =   495
         Width           =   630
      End
   End
   Begin VB.Frame fraOther 
      Height          =   570
      Left            =   0
      TabIndex        =   65
      Top             =   5475
      Width           =   10635
      Begin VB.TextBox txt 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Index           =   31
         Left            =   4455
         TabIndex        =   52
         Top             =   180
         Width           =   480
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ҫ���(&X)"
         Height          =   195
         Left            =   2475
         TabIndex        =   50
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   6045
         TabIndex        =   54
         Top             =   180
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   870
         TabIndex        =   49
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   91291651
         CurrentDate     =   38545
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��"
         Height          =   180
         Index           =   5
         Left            =   75
         TabIndex        =   48
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&Y)      ��"
         Height          =   180
         Index           =   29
         Left            =   3810
         TabIndex        =   51
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��ע(&L)"
         Height          =   180
         Index           =   13
         Left            =   5325
         TabIndex        =   53
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmScheduleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnGroup As Boolean
Private mlngDept As Long
Private mrsItems As New ADODB.Recordset                 '������ʱ����ѡ��������Ŀ
Private mrsPersons As New ADODB.Recordset                 '������ʱ�����Ա
Private mrsGroup As New ADODB.Recordset                 '������ʱ�������
Private mlngLoop As Long
Private mblnDataChange As Boolean
Private mbytMode As Byte                        '��־,
Private mstrGroup As String
Private mstrSQL As String

Private Enum mCol
    ��Ŀ = 1
    ִ�п���
    ��鲿λ
    �ɼ���ʽ
    �ɼ�����
    ����걾
    �����۸�
    ���۸�
    �ۿ�
    �������
    ���
    ���㷽ʽ
    ִ�п���id
    �ɼ���ʽid
    �ɼ�����id
    ��鲿λid
    �Ʒ���ϸ
    �¼�
    ǰ��ɫ
    ɾ��
    ����
    
    p�Ƽ���Ŀ = 1
    p����
    p���㵥λ
    p����
    p��׼����
    p��쵥��
    p�ۿ�
    p��׼���
    p�����
    pִ�п���
    pִ�п���id
    p�շ���Ŀid
    p�Ƽ�����
    p���
    p���ÿ��
End Enum

Private Enum mPersonCol
    ���� = 1
    �����
    ������
    �Ա�
    ����
    ����״��
    ��������
    ���֤
    ����
    ����
    ѧ��
    ְҵ
    ���
    ��ϵ������
    ��ϵ�˵绰
    �����ʼ�
    ��ϵ�˵�ַ
    ������λ
    �Ǽ�ʱ��
    ����id
    IC����
    ���￨��
    ǰ��ɫ
    
End Enum

Private Enum mColChar
    
    ���� = 66
    �Ա�
    ����
    ��������
    ����״��
    ���֤��
    �����
    ������
    ���￨��
    ������λ
    �����ʼ�
    ����
    ѧ��
    ְҵ
    ����
    �����
    
End Enum

'�������Զ�����̻���************************************************************************************************

Private Property Let DataChange(ByVal vData As Boolean)
        mblnDataChange = vData
End Property

Private Property Get DataChange() As Boolean
        DataChange = mblnDataChange
End Property

Private Function CountGroup() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:�����ͳ����Ŀ�������������С�Ů��
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    
    If mblnGroup Then
        strTmp = """" & lvwGroup.SelectedItem.Text & """�����"
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If vsf.TextMatrix(lngLoop, mCol.���) = "���" Then
                lngCount1 = lngCount1 + 1
            Else
                lngCount2 = lngCount2 + 1
            End If
        End If
    Next
    
    strTmp = strTmp & "������Ŀ" & lngCount1 + lngCount2 & "��(���:" & lngCount1 & "��,����:" & lngCount2 & "��)"
    
    If mblnGroup Then
        lngCount1 = 0
        lngCount2 = 0
        
        For lngLoop = 1 To vsfPerson.Rows - 1
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.����)) <> "" Then
                If InStr(vsfPerson.TextMatrix(lngLoop, mPersonCol.�Ա�), "��") > 0 Then
                    lngCount1 = lngCount1 + 1
                Else
                    lngCount2 = lngCount2 + 1
                End If
            End If
        Next
        
        strTmp = strTmp & ";������Ա" & lngCount1 + lngCount2 & "��(����:" & lngCount1 & "��,Ů��:" & lngCount2 & "��)"
    End If
    
    stbThis.Panels(2).Text = strTmp
    
End Function

Private Function ChangeTotal(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim db�ۿ� As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim dbTotal As Double
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '�仯���
        
        '1.�����ۿ�
        db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")

    Else
        '�仯�ۿ�
        db�ۿ� = dbTmp

    End If
    
    txtSum(1).Text = Format(dbMoney * db�ۿ� / 10, "0.00")
    txtSum(2).Text = Format(db�ۿ�, "0.0000")
    dbTotal = 0
    
    For lngLoop = 1 To vsf.Rows - 1
    
        vsf.TextMatrix(lngLoop, mCol.�ۿ�) = db�ۿ�
        vsf.TextMatrix(lngLoop, mCol.���۸�) = Format(Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)) * (db�ۿ� / 10), "0.00")
        
        dbTotal = dbTotal + Val(vsf.TextMatrix(lngLoop, mCol.���۸�))
                    
        varRow = Split(vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db�ۿ� / 10), "0.00000")
                varCol(10) = db�ۿ�
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ) = Join(varRow, ";")
    Next

    '����
    '------------------------------------------------------------------------------------------------------------------
    If dbTotal <> Val(txtSum(1).Text) Then

        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) <> 0 Then
            
                vsf.TextMatrix(lngLoop, mCol.���۸�) = Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) + (Val(txtSum(1).Text) - dbTotal)
                
                If Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)) <> 0 Then
                    vsf.TextMatrix(lngLoop, mCol.�ۿ�) = Format(10 * Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) / Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)), "0.0000")
                Else
                    vsf.TextMatrix(lngLoop, mCol.�ۿ�) = 0
                End If
                
                varRow = Split(vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ), ";")
                For lngRow = 0 To UBound(varRow)
                    If varRow(lngRow) <> "" Then
                        varCol = Split(varRow(lngRow), ":")
                        If Val(varCol(4)) <> 0 Then
                            varCol(4) = Val(varCol(4)) + (Val(txtSum(1).Text) - dbTotal)
                            If Val(varCol(3)) <> 0 Then
                                varCol(10) = Format(10 * Val(varCol(4)) / Val(varCol(3)), "0.0000")
                            Else
                                varCol(10) = 0
                            End If
                        End If
                    End If
                    varRow(lngRow) = Join(varCol, ":")
                Next
                vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ) = Join(varRow, ";")
                Exit For
            End If
        Next
    End If

    ChangeTotal = True
    
End Function

Private Function ChangeItem(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1, Optional ByVal blnUpdate As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db�ۿ� As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    
    If blnUpdate Then
        If dbMoney = 0 Then Exit Function
        
        Call WritePrice(vsf.Row)
        
        If bytMode = 1 Then
            '�仯���
            
            '1.�����ۿ�
            db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")
        Else
            '�仯�ۿ�
            db�ۿ� = dbTmp
            
        End If
        
        vsf.TextMatrix(vsf.Row, mCol.���۸�) = Format(dbMoney * db�ۿ� / 10, "0.00")
        vsf.TextMatrix(vsf.Row, mCol.�ۿ�) = Format(db�ۿ�, "0.0000")
    End If
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.�����۸�))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.���۸�))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
    
    '���¼۸�
    '------------------------------------------------------------------------------------------------------------------
    If blnUpdate Then
        varRow = Split(vsf.TextMatrix(vsf.Row, mCol.�Ʒ���ϸ), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db�ۿ� / 10), "0.00000")
                varCol(10) = db�ۿ�
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(vsf.Row, mCol.�Ʒ���ϸ) = Join(varRow, ";")
    End If
        
    ChangeItem = True
    
End Function

Private Function ChangePrice(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db�ۿ� As Double
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '�仯���
        
        '1.�����ۿ�
        db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")
    Else
        '�仯�ۿ�
        db�ۿ� = dbTmp
        
    End If
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p��쵥��) = Format(dbMoney * db�ۿ� / 10, "0.00000")
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p�ۿ�) = Format(db�ۿ�, "0.0000")
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p�����) = Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p����)) * Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p��쵥��))
    
    '������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p��׼���))
    Next
    vsf.TextMatrix(vsf.Row, mCol.�����۸�) = dbSum
    
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p�����))
    Next
    vsf.TextMatrix(vsf.Row, mCol.���۸�) = dbSum
    
    If Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)) <> 0 Then
        vsf.TextMatrix(vsf.Row, mCol.�ۿ�) = Format(10 * Val(vsf.TextMatrix(vsf.Row, mCol.���۸�)) / Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)), "0.0000")
    Else
        vsf.TextMatrix(vsf.Row, mCol.�ۿ�) = "0.0000"
    End If
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.�����۸�))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.���۸�))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
        
    ChangePrice = True
    
End Function

Private Function SumPrice(ByVal bytMode As Byte) As Single
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim sglSum As Single
    
    For lngLoop = 1 To vsfPrice.Rows - 1
        If bytMode = 2 Then
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p�����))
        Else
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p��׼���))
        End If
    Next
    SumPrice = sglSum
    
End Function

Private Function GetPatientInfo(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    strSQL = "SELECT A.* FROM ������Ϣ A WHERE A.����id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        
        If mblnGroup Then
                If Val(cmd(4).Tag) <> Val(zlCommFun.NVL(rs("��ͬ��λid"))) And Val(zlCommFun.NVL(rs("��ͬ��λid"))) > 0 And Val(cmd(4).Tag) > 0 Then
                    
                    If MsgBox("���ǵ�ǰ�������Ա���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    
                End If
                
                vsfPerson.EditText = zlCommFun.NVL(rs("����"))
                vsfPerson.Cell(flexcpData, vsfPerson.Row, vsfPerson.Col) = zlCommFun.NVL(rs("����").Value)
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                
                Call SetRowDefault(0, vsfPerson.Row, "ȱʡ��Ϣ")
                
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����"))
                
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��"))
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = Format(zlCommFun.NVL(rs("��������")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = zlCommFun.NVL(rs("����id"))
                
                Call FillPatient(Val(zlCommFun.NVL(rs("����id"))), 2)
                
                vsfPerson.EditMode(mPersonCol.�����) = 0
        Else
                    cmd(1).Tag = zlCommFun.NVL(rs("����id").Value)
                    txt(5).Text = zlCommFun.NVL(rs("����").Value)
                    txt(4).Text = zlCommFun.NVL(rs("���֤��").Value)
                    txt(9).Text = zlCommFun.NVL(rs("����").Value)
                    
                    txt(3).Text = zlCommFun.NVL(rs("�����").Value)
                    
                    zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("�Ա�").Value)
                    zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("����״��").Value)
                    
                    Call FillPatient(Val(cmd(1).Tag))
                    
                    
                    txt(5).Tag = ""
                    imgNew(1).Visible = False
                    
                    txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
        End If
        
        DataChange = True
        
    End If
    
    GetPatientInfo = True
    
End Function


Private Function CreatePriceList(ByVal intRow As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    
    strKeys = CStr(Val(vsf.RowData(intRow))) & "'" & CStr(Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid))) & "'" & vsf.TextMatrix(intRow, mCol.��鲿λid)
    
    Dim str�Ƽ���Ŀ As String
    Dim str�Ƽ����� As String
    
    vsfPrice.Rows = 2
    str�Ƽ���Ŀ = vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ)
    str�Ƽ����� = vsfPrice.TextMatrix(1, mCol.p�Ƽ�����)
    
    vsfPrice.Body.Cell(flexcpText, 1, mCol.p�Ƽ���Ŀ + 1, 1, vsfPrice.Cols - 1) = ""
    vsfPrice.RowData(1) = 0

    vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ) = str�Ƽ���Ŀ
    vsfPrice.TextMatrix(1, mCol.p�Ƽ�����) = str�Ƽ�����
    
    mstrSQL = GetPublicSQL(SQL.�����Ŀ�۱�, strKeys)
    
    If vsf.TextMatrix(intRow, mCol.��鲿λid) = "" Then
        '����򵥲�λ���
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf.RowData(intRow)), Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)))
    Else
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    End If
    
    If rs.BOF = False Then
        With vsfPrice
            Do While Not rs.EOF
                
                If Val(.TextMatrix(.Rows - 1, mCol.p�շ���Ŀid)) > 0 Then
                    .Rows = .Rows + 1
                End If
                
                If zlCommFun.NVL(rs("�Ƽ�����")) = 2 Then
                    .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�ɼ���ʽ-" & vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ)
                ElseIf vsf.TextMatrix(vsf.Row, mCol.���) = "����" Then
                    .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "������Ŀ-" & vsf.TextMatrix(vsf.Row, mCol.��Ŀ)
                Else
                    .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�����Ŀ-" & vsf.TextMatrix(vsf.Row, mCol.��Ŀ)
                End If
                
                .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rs("����"))
                .TextMatrix(.Rows - 1, mCol.p���㵥λ) = zlCommFun.NVL(rs("���㵥λ"))
                .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rs("�շ�����"))
                .TextMatrix(.Rows - 1, mCol.p��׼����) = zlCommFun.NVL(rs("�ּ�"))
                .TextMatrix(.Rows - 1, mCol.p��쵥��) = zlCommFun.NVL(rs("�ּ�"))
                .TextMatrix(.Rows - 1, mCol.p�ۿ�) = 10
                .TextMatrix(.Rows - 1, mCol.p��׼���) = zlCommFun.NVL(rs("�շ�����"), 0) * zlCommFun.NVL(rs("�ּ�"), 0)
                .TextMatrix(.Rows - 1, mCol.p�����) = zlCommFun.NVL(rs("�շ�����"), 0) * zlCommFun.NVL(rs("�ּ�"), 0)
                .TextMatrix(.Rows - 1, mCol.p�շ���Ŀid) = zlCommFun.NVL(rs("ID"))
                
                .TextMatrix(.Rows - 1, mCol.p�Ƽ�����) = zlCommFun.NVL(rs("�Ƽ�����"))
                .TextMatrix(.Rows - 1, mCol.p���) = zlCommFun.NVL(rs("���"))
                
                Call SetRowDefault(zlCommFun.NVL(rs("ID"), 0), .Rows - 1, "�շ�ִ�п���")
                
                If InStr("567", .TextMatrix(.Rows - 1, mCol.p���)) > 0 Then
                    .TextMatrix(.Rows - 1, mCol.p���ÿ��) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.pִ�п���id)))
                    Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p����)), Val(.TextMatrix(.Rows - 1, mCol.p���ÿ��)), .TextMatrix(.Rows - 1, mCol.p����), .TextMatrix(.Rows - 1, mCol.pִ�п���), .TextMatrix(.Rows - 1, mCol.p���㵥λ), 1)
                End If
                
                rs.MoveNext
            Loop
        End With
        
    End If
    
    vsf.TextMatrix(intRow, mCol.�����۸�) = SumPrice(1)
    vsf.TextMatrix(intRow, mCol.���۸�) = SumPrice(2)
    
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngDept As Long, Optional blnGroup As Boolean = False, Optional ByVal bytMode As Byte = 1) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    
    mblnOK = False
    mlngKey = lngKey
    mblnGroup = blnGroup
    mlngDept = lngDept
    mbytMode = bytMode
    
    Set mfrmMain = frmMain
    
    Call ClearData
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
    
        imgNew(0).Visible = False
        imgNew(1).Visible = False
        
        If ReadData(mlngKey) = False Then Exit Function
'        stbThis.Panels(2).Text = "�޸����ԤԼ��"
    Else
        If mblnGroup Then
            Call ReadGroup(0)
        End If
        
        imgNew(0).Visible = True
        imgNew(1).Visible = True
        
'        stbThis.Panels(2).Text = "�¿����ԤԼ��"
    End If
    
    Call CountGroup
    
    DataChange = False
    txt(5).Tag = ""
    txt(13).Tag = ""
            
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next
    
    For lngLoop = 0 To txt.UBound
        txt(lngLoop).Text = ""
        txt(lngLoop).Tag = ""
    Next
    
    On Error GoTo 0
    
    lvwGroup.ListItems.Clear
    Call ResetVsf(vsf)
    Call ResetVsf(vsfPrice)
    Call ResetVsf(vsfPerson)
    
    DataChange = False
    
        
End Function

Private Function InitMaxLength() As Boolean
    
    '����������볤��
    txt(5).MaxLength = GetMaxLength("������Ϣ", "����")
    txt(4).MaxLength = GetMaxLength("������Ϣ", "���֤��")
    txt(13).MaxLength = GetMaxLength("��Լ��λ", "����")
    txt(12).MaxLength = GetMaxLength("��Լ��λ", "��ϵ��")
    txt(11).MaxLength = GetMaxLength("��Լ��λ", "��ϵ�绰")
    txt(0).MaxLength = GetMaxLength("���ǼǼ�¼", "��ϵ��")
    txt(1).MaxLength = GetMaxLength("���ǼǼ�¼", "��ϵ�绰")
    txt(2).MaxLength = GetMaxLength("���ǼǼ�¼", "��ϵ��ַ")
    txt(6).MaxLength = GetMaxLength("���ǼǼ�¼", "����˵��")
    
    txt(10).MaxLength = GetMaxLength("������Ϣ", "��ϵ�˵绰")
    
    InitMaxLength = True
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrGroup = ""
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "����", 2100, 1, "...", 1
        .NewColumn "ִ�п���", 1080, 1, " ", 1
        
        .NewColumn "��鲿λ", 1800, 1, "...", 1
        .NewColumn "�ɼ���ʽ", 1200, 1, " ", 1
        .NewColumn "�ɼ�����", 1080, 1, " ", 1
        
        .NewColumn "����걾", 900, 1, " ", 1
        .NewColumn "�����۸�", 900, 7
        .NewColumn "���۸�", 900, 7, , 1
        .NewColumn "�ۿ�", 900, 7, , 1
        .NewColumn "�������", 0, 1
        .NewColumn "���", 0, 1
        .NewColumn "���㷽ʽ", 900, 1, "����|�շ�", 1
        .NewColumn "ִ�п���id", 0, 1
        .NewColumn "�ɼ���ʽid", 0, 1
        .NewColumn "�ɼ�����id", 0, 1
        .NewColumn "��鲿λid", 0, 1
        .NewColumn "�Ʒ���ϸ", 0, 1
        .NewColumn "�¼�", 0, 1
        .NewColumn "ǰ��ɫ", 0, 1
        .NewColumn "ɾ��", 0, 1
        .NewColumn "����", 0, 1
        .FixedCols = 1
        
        .SelectMode = True
        
        .Body.ColFormat(mCol.�����۸�) = "0.00"
        .Body.ColFormat(mCol.���۸�) = "0.00"
        .Body.ColFormat(mCol.�ۿ�) = "0.0000"
    End With
    
    With vsfPrice
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "�Ƽ���Ŀ", 2100, 1, " ", 1
        .NewColumn "�շ���Ŀ", 2700, 1, "...", 1
        .NewColumn "��λ", 600, 1
        .NewColumn "����", 540, 7, , 1
        .NewColumn "��׼����", 900, 7
        .NewColumn "��쵥��", 900, 7, , 1
        .NewColumn "�ۿ�", 900, 7, , 1
        .NewColumn "��׼�۸�", 900, 7
        .NewColumn "���۸�", 900, 7
        .NewColumn "ִ�п���", 1080, 1, " ", 1
        .NewColumn "ִ�п���id", 0
        .NewColumn "�շ���Ŀid", 0
        .NewColumn "�Ƽ�����", 0
        .NewColumn "���", 0
        .NewColumn "", 0
        .FixedCols = 1
        .Body.ColFormat(mCol.p��׼����) = "0.00000"
        .Body.ColFormat(mCol.p��쵥��) = "0.00000"
        .Body.ColFormat(mCol.p��׼���) = "0.00"
        .Body.ColFormat(mCol.p�����) = "0.00"
        .Body.ColFormat(mCol.p�ۿ�) = "0.0000"
        .SelectMode = True
    End With
    
    With vsfPerson
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "����", 990, 1, "...", 1, GetMaxLength("������Ϣ", "����")
        .NewColumn "�����", 810, 1
        .NewColumn "������", 810, 1, , 1, GetMaxLength("������Ϣ", "������")
        .NewColumn "�Ա�", 750, 1, GetCombList("SELECT ���� FROM �Ա�"), 1, GetMaxLength("������Ϣ", "�Ա�")
        .NewColumn "����", 540, 1, , 1, GetMaxLength("������Ϣ", "����")
        .NewColumn "����״��", 900, 1, GetCombList("SELECT ���� FROM ����״��"), 1, GetMaxLength("������Ϣ", "����״��")
        .NewColumn "��������", 990, 1, , 1
        .NewColumn "���֤", 1800, 1, , 1, GetMaxLength("������Ϣ", "���֤��")
                
        .NewColumn "����", 0, 1, , , GetMaxLength("������Ϣ", "����")
        .NewColumn "����", 0, 1, , , GetMaxLength("������Ϣ", "����")
        .NewColumn "ѧ��", 0, 1, , , GetMaxLength("������Ϣ", "ѧ��")
        .NewColumn "ְҵ", 0, 1, , , GetMaxLength("������Ϣ", "ְҵ")
        .NewColumn "���", 0, 1, , , GetMaxLength("������Ϣ", "���")
        .NewColumn "��ϵ������", 0, 1, , , GetMaxLength("������Ϣ", "��ϵ������")
        .NewColumn "��ϵ�˵绰", 0, 1, , , GetMaxLength("������Ϣ", "��ϵ�˵绰")
        .NewColumn "�����ʼ�", 0, 1, , , GetMaxLength("������Ϣ", "�����ʼ�")
        .NewColumn "��ϵ�˵�ַ", 0, 1, , , GetMaxLength("������Ϣ", "��ϵ�˵�ַ")
        .NewColumn "������λ", 0, 1, , , GetMaxLength("������Ϣ", "������λ")
        .NewColumn "�Ǽ�ʱ��", 0, 1
        .NewColumn "����id", 0, 1
        .NewColumn "IC����", 0, 1
        .NewColumn "���￨��", 0, 1
        
        .FixedCols = 1
        .SelectMode = True
        .Body.ColEditMask(mPersonCol.��������) = "0000-00-00"
    End With
       
       
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM �Ա� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(1), rs)
    
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM ����״�� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs)
    
    '����������볤��
    Call InitMaxLength
    
    '����͸������
    fraGroupInfo.Visible = False
    fraSingle.Visible = False
    
    If mblnGroup Then
    
        lblTitle.Caption = "�������" & IIf(mbytMode = 2, "�Ǽ�", "ԤԼ����")
        fraGroupInfo.Visible = True
        fraGroup.Visible = True
        tbs.TabVisible(1) = True
        
    Else
        lblTitle.Caption = "�������" & IIf(mbytMode = 2, "�Ǽ�", "ԤԼ����")
                
        tbs.TabVisible(1) = False
        fraSingle.Visible = True
        fraGroup.Visible = False
        
    End If
    
    If mbytMode = 2 Then
        Me.Caption = "���Ǽ�"
        fraInfo.Visible = False
        dtp(0).Enabled = False
        dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
    Else
        dtp(0).Value = Format(DateAdd("d", 7, CDate(zlDatabase.Currentdate)), dtp(0).CustomFormat)
    End If
    
    '��ʼ������
    'lvwGroup.TextMatrix(1, 1) = "ȱʡ"
    lvwGroup.ListItems.Add , , "ȱʡ", 1, 1
    
    '1.������¼��,���ڱ���ѡ��������Ŀ
    Call MedicalItemsRecord(mrsItems)
    
    '2.������¼��,���ڱ���ѡ��������Ա
    Call MedicalItemsRecord(mrsPersons, 2)
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
            
    On Error GoTo errHand
    
    '��ȡԤԼ������Ϣ
    gstrSQL = "SELECT * FROM ���ǼǼ�¼ A WHERE A.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        txt����.Text = zlCommFun.NVL(rs("����").Value)
        txt(0).Text = zlCommFun.NVL(rs("��ϵ��").Value)
        txt(1).Text = zlCommFun.NVL(rs("��ϵ�绰").Value)
        txt(2).Text = zlCommFun.NVL(rs("��ϵ��ַ").Value)
        txt(6).Text = zlCommFun.NVL(rs("����˵��").Value)
        cmd(4).Tag = zlCommFun.NVL(rs("��Լ��λid").Value, 0)
        dtp(0).Value = Format(zlCommFun.NVL(rs("���ʱ��").Value), dtp(0).CustomFormat)
        txt(31).Text = zlCommFun.NVL(rs("�������").Value, 0)
        chk.Value = IIf(Val(txt(31).Text) > 0, 1, 0)
    End If
                                        
    If mblnGroup Then Call ReadGroup(Val(cmd(4).Tag))
            
    Set rs = zlDatabase.OpenSQLRecord(GetPublicSQL(SQL.�����Ա����), Me.Caption, lngKey)
    If WriteItems(rs, mrsPersons, , 2) = False Then Exit Function
    
    If mrsPersons.RecordCount > 0 And mblnGroup = False Then Call ReadPersons("ȱʡ", 2)
    
    '��ȡ�����������Ŀ
    
    lvwGroup.ListItems.Clear
    
    gstrSQL = "SELECT A.������� AS ����, rownum AS ID,1 As ͼ�� FROM ������ A WHERE A.�Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillLvw(lvwGroup, rs)
    Else
        lvwGroup.ListItems.Add , , "ȱʡ", 1, 1
    End If
        
    '��ȡ�����Ŀ
    Set rs = zlDatabase.OpenSQLRecord(GetPublicSQL(SQL.���������Ŀ), Me.Caption, mlngKey)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
        
            mrsItems.AddNew
            
            mrsItems("���").Value = zlCommFun.NVL(rs("�������").Value)
            mrsItems("ID").Value = zlCommFun.NVL(rs("ID").Value)
            mrsItems("���").Value = zlCommFun.NVL(rs("���").Value)
            mrsItems("����").Value = zlCommFun.NVL(rs("����").Value)
            mrsItems("�����۸�").Value = Format(zlCommFun.NVL(rs("�����۸�").Value), "0.00")
            mrsItems("���۸�").Value = Format(zlCommFun.NVL(rs("���۸�").Value), "0.00")
            mrsItems("�������").Value = zlCommFun.NVL(rs("�������").Value)
            mrsItems("���㷽ʽ").Value = zlCommFun.NVL(rs("���㷽ʽ").Value)
            mrsItems("ִ�п���").Value = zlCommFun.NVL(rs("ִ�п���").Value)
            mrsItems("�ɼ�����").Value = zlCommFun.NVL(rs("�ɼ�����").Value)
            mrsItems("�ɼ�����id").Value = zlCommFun.NVL(rs("�ɼ�����id").Value)
            mrsItems("ִ�п���id").Value = zlCommFun.NVL(rs("ִ�п���id").Value)
            mrsItems("�ɼ���ʽ").Value = zlCommFun.NVL(rs("�ɼ���ʽ").Value)
            mrsItems("�ɼ���ʽid").Value = zlCommFun.NVL(rs("�ɼ���ʽid").Value)
            mrsItems("����걾").Value = zlCommFun.NVL(rs("����걾").Value)
            mrsItems("��鲿λ").Value = zlCommFun.NVL(rs("��鲿λ").Value)
            mrsItems("��鲿λid").Value = zlCommFun.NVL(rs("��鲿λid").Value)
            mrsItems("�ۿ�").Value = Format(zlCommFun.NVL(rs("�ۿ�").Value), "0.0000")
            mrsItems("�Ʒ���ϸ").Value = GetPriceList(zlCommFun.NVL(rs("�嵥id").Value))
            
            rs.MoveNext
        Loop
    End If
        
    If Not (lvwGroup.SelectedItem Is Nothing) Then Call lvwGroup_ItemClick(lvwGroup.SelectedItem)
    
    ReadData = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ReadGroup(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT A.* FROM ��Լ��λ A WHERE A.ID=" & lngKey
    
    If mrsGroup.State = adStateOpen Then mrsGroup.Close
    mrsGroup.Open gstrSQL, gcnOracle, adOpenStatic, adLockBatchOptimistic
    If mrsGroup.BOF Then mrsGroup.AddNew
    
    Call ShowGroupInfo
    
    ReadGroup = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ShowGroupInfo() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��д������Ϣ���ؼ�
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    If mrsGroup.RecordCount > 0 Then
        mrsGroup.MoveFirst
        
        txt(13).Text = zlCommFun.NVL(mrsGroup("����").Value)
        txt(12).Text = zlCommFun.NVL(mrsGroup("��ϵ��").Value)
        txt(11).Text = zlCommFun.NVL(mrsGroup("�绰").Value)
        txt(8).Text = zlCommFun.NVL(mrsGroup("�����ʼ�").Value)
        cmd(4).Tag = zlCommFun.NVL(mrsGroup("ID").Value)
        
    End If
    
    ShowGroupInfo = True
    
End Function

Private Function SaveGroupInfo() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '����:
    '------------------------------------------------------------------------------------------------------------------
    If mblnGroup Then
        If mrsGroup.RecordCount > 0 Then
            mrsGroup.MoveFirst
            
            mrsGroup("����").Value = txt(13).Text
            mrsGroup("��ϵ��").Value = txt(12).Text
            mrsGroup("�绰").Value = txt(11).Text
            mrsGroup("�����ʼ�").Value = txt(8).Text
            mrsGroup("ID").Value = Val(cmd(4).Tag)
            
        End If
    End If
    
    SaveGroupInfo = True
    
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function CheckHavePerson(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsfPerson.Rows - 1
        If Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)) = lngKey And vsfPerson.Row <> lngLoop And Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)) > 0 Then
            CheckHavePerson = True
            Exit Function
        End If
    Next
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal lngCol As Long = 0) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '����:  ���б�ʽ��ʾ����
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strClass As String
    Dim strPath As String
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    ShowOpenList = 2
    
    Select Case lngCol
        Case mCol.��Ŀ
            strText = UCase(strText)
            
            strLvw = "����,1200,0,1;����,2700,0,0;��λ,900,0,0;�걾��λ,900,0,0;���,900,0,0"
            strPath = Me.Name & "\�����Ŀѡ��"
            
            gstrSQL = GetPublicSQL(SQL.�����Ŀ����ѡ��, strText)
            If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                strTmp = strText & "%"
            Else
                strTmp = "%" & strText & "%"
            End If
            Dim bytParam1 As Byte
            Dim bytParam2 As Byte
            
            bytParam1 = 1
            bytParam2 = 2
                    
            If mblnGroup = False Then
                Select Case zlCommFun.GetNeedName(cbo(1).Text)
                Case "��"
                    bytParam1 = 1
                    bytParam2 = 1
                Case "Ů"
                    bytParam1 = 2
                    bytParam2 = 2
                End Select
            End If
            
            If Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, bytParam1, bytParam2)
            ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "", strText & "%", strTmp, bytParam1, bytParam2)
            Else
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "D", "", strText & "%", strTmp, bytParam1, bytParam2)
            End If

        Case mCol.��鲿λ
            
            strText = "'%" & UCase(strText) & "%'"
            
            strLvw = "����,3300,0,0"
            strPath = Me.Name & "\��鲿λѡ��"
            
            gstrSQL = "select B.�걾��λ AS ����,B.ID,0 AS ѡ�� from ������Ŀ��� A,������ĿĿ¼ B WHERE (B.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or B.����ʱ�� is NULL) AND A.������ĿID=B.ID AND A.�������ID=" & Val(vsf.RowData(vsf.Row)) & ""
            
            rs.CursorLocation = adUseClient
            If rs.State = adStateOpen Then rs.Close
            rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
            
    End Select
    
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    If rs.RecordCount = 1 And strText <> "'%%'" Then GoTo PointOver
    Call CalcPosition(sglX, sglY, vsf)
    
    If lngCol = mCol.��鲿λ Then
        If vsf.TextMatrix(vsf.Row, mCol.��鲿λid) <> "" Then
            Do While Not rs.EOF
                If InStr("," & vsf.TextMatrix(vsf.Row, mCol.��鲿λid) & ",", "," & rs("ID").Value & ",") > 0 Then rs("ѡ��").Value = 1
                rs.MoveNext
            Loop
        End If
        rs.MoveFirst
        
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "�������ѡ������Ŀ,Ȼ��س���˫���˳�", sglX + 60, sglY + 30, 9000, 4500, 300, , strPath, , False, True) Then GoTo PointOver
        
    Else
                
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "�������ѡ��һ����Ŀ", sglX + 60, sglY + 30, 9000, 4500, vsf.Body.RowHeight(1), , strPath, , False) Then GoTo PointOver
        
    End If
        
    Exit Function
    
PointOver:
    Select Case lngCol
        Case mCol.��Ŀ
            If CheckHave(zlCommFun.NVL(rs("ID").Value, 0)) Then
                MsgBox "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���Ѿ���ѡ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            vsf.Cell(flexcpText, vsf.Row, mCol.��Ŀ + 1, vsf.Row, vsf.Cols - 1) = ""
            
            vsf.EditText = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            
        Case mCol.��鲿λ
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = ""
            
            rs.Filter = ""
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    vsf.TextMatrix(vsf.Row, vsf.Col) = vsf.TextMatrix(vsf.Row, vsf.Col) & zlCommFun.NVL(rs("����").Value) & ","
                    vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = vsf.TextMatrix(vsf.Row, mCol.��鲿λid) & zlCommFun.NVL(rs("ID").Value) & ","
                    rs.MoveNext
                Loop
                
                If vsf.TextMatrix(vsf.Row, mCol.��鲿λ) <> "" Then vsf.TextMatrix(vsf.Row, mCol.��鲿λ) = Mid(vsf.TextMatrix(vsf.Row, mCol.��鲿λ), 1, Len(vsf.TextMatrix(vsf.Row, mCol.��鲿λ)) - 1)
                If vsf.TextMatrix(vsf.Row, mCol.��鲿λid) <> "" Then vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = Mid(vsf.TextMatrix(vsf.Row, mCol.��鲿λid), 1, Len(vsf.TextMatrix(vsf.Row, mCol.��鲿λid)) - 1)
                
            End If

    End Select
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SetRowData(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
'------------------------------------------------------------------------------------------------------------------
    '����:���������ݣ����в�ͬ����ͬ��
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error Resume Next
    
    For lngLoop = 0 To UBound(arryMode)
        Select Case arryMode(lngLoop)
        Case "�շ�ִ�п���"
        
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p���)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.ҩƷִ�п���)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p���))
            Else
                gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���, "1")
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
            End If
            If rs.RecordCount > 1 Then
                vsfPrice.EditMode(mCol.pִ�п���) = 1
                vsfPrice.Body.ColComboList(mCol.pִ�п���) = vsfPrice.Body.BuildComboList(rs, "����", "ID")
            Else
                vsfPrice.EditMode(mCol.pִ�п���) = 0
                vsfPrice.Body.ColComboList(mCol.pִ�п���) = ""
            End If
        
        Case "�Ƽ���Ŀ"
            
            If Trim(vsf.TextMatrix(intRow, mCol.���)) = "���" Then
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(intRow, mCol.��Ŀ))
                vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ�����) = "1"
            Else
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(intRow, mCol.��Ŀ))
                If Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(intRow, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                End If
            End If
            
        Case "����ִ�п���"
        
            gstrSQL = GetPublicSQL(SQL.����ִ�п���, "1")
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.ִ�п���) = 1
                    vsf.Body.ColComboList(mCol.ִ�п���) = vsf.Body.BuildComboList(rs, "����", "ID")
                Else
                    vsf.EditMode(mCol.ִ�п���) = 0
                    vsf.Body.ColComboList(mCol.ִ�п���) = ""
                End If
            End If
        
        Case "�ɼ���ʽ"
        
            gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A,�����÷����� B WHERE A.ID=B.�÷�id AND A.���='E' AND A.��������='6' AND B.��ĿID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                vsf.EditMode(mCol.�ɼ���ʽ) = 1
                vsf.Body.ColComboList(mCol.�ɼ���ʽ) = vsf.Body.BuildComboList(rs, "����", "ID")
            Else
                gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A WHERE A.���='E' AND A.��������='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.�ɼ���ʽ) = 1
                    vsf.Body.ColComboList(mCol.�ɼ���ʽ) = vsf.Body.BuildComboList(rs, "����", "ID")
                Else
                    vsf.EditMode(mCol.�ɼ���ʽ) = 0
                    vsf.Body.ColComboList(mCol.�ɼ���ʽ) = ""
                End If
            End If
            
        Case "�ɼ�����"
        
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)), mlngDept, UserInfo.����ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.�ɼ�����) = 1
                    vsf.Body.ColComboList(mCol.�ɼ�����) = vsf.Body.BuildComboList(rs, "*����", "ID")
                Else
                    vsf.EditMode(mCol.�ɼ�����) = 0
                    vsf.Body.ColComboList(mCol.�ɼ�����) = ""
                End If
            End If
        
        Case "����걾"
        
            gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '�������Ŀ
                
                gstrSQL = "SELECT DISTINCT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID<>[1] AND nvl(C.�����Ŀ,0)=0 " & _
                            "AND B.������Ŀid=A.��Ŀid and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.������Ŀid IN (SELECT C.ID " & _
                             "FROM ���鱨����Ŀ A," & _
                                  "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = [1]) B," & _
                                  "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                            "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> [1] AND " & _
                                  "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID=[1] AND nvl(C.�����Ŀ,0)=0 AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                
                vsf.EditMode(mCol.����걾) = 1
                vsf.Body.ColComboList(mCol.����걾) = vsf.Body.BuildComboList(rs, "����", "����")
                
            Else
                
                'û�ж�Ӧʱ����ȡ���б걾����
                gstrSQL = "SELECT ���� FROM ���Ƽ���걾 A"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                
                    vsf.EditMode(mCol.����걾) = 1
                    vsf.Body.ColComboList(mCol.����걾) = vsf.Body.BuildComboList(rs, "����", "����")
                Else
                    vsf.EditMode(mCol.����걾) = 0
                    vsf.Body.ColComboList(mCol.����걾) = ""
                End If
                
            End If
        
        End Select
    Next
    
    SetRowData = True
    
End Function

Private Function SetRowDefault(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error GoTo errHand
    
    For lngLoop = 0 To UBound(arryMode)
        
        Select Case arryMode(lngLoop)
        Case "ȱʡ��Ϣ"
            '�Ȱ����ж�ȡ
            With vsfPerson
                If vsfPerson.Row > 1 Then
                    .TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = .TextMatrix(vsfPerson.Row - 1, mPersonCol.�Ա�)
                    .TextMatrix(vsfPerson.Row, mPersonCol.����״��) = .TextMatrix(vsfPerson.Row - 1, mPersonCol.����״��)
                End If
            End With
            
        Case "���㷽ʽ"
            
            If mblnGroup Then
                vsf.TextMatrix(vsf.Row, mCol.���㷽ʽ) = "����"
            Else
                vsf.TextMatrix(vsf.Row, mCol.���㷽ʽ) = "�շ�"
            End If
            
        Case "ִ�п���"
'            lng��������id = mlngDept
            
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���) = zlCommFun.NVL(rs("����").Value)
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���id) = zlCommFun.NVL(rs("ID").Value)
                Else
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���) = gstrDeptName
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���id) = UserInfo.����ID
                End If
            End If
        
        Case "�ɼ���ʽ"
           
            
            gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A,�����÷����� B WHERE A.ID=B.�÷�id AND A.���='E' AND A.��������='6' AND B.��ĿID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("ID").Value)
            Else
                gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A WHERE A.���='E' AND A.��������='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("����").Value)
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
            
        Case "�ɼ�����"
                    
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)), mlngDept, UserInfo.����ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ�����) = zlCommFun.NVL(rs("����").Value)
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ�����id) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
        
        Case "����걾"
            
            
            gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '�������Ŀ
                
                gstrSQL = "SELECT DISTINCT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID<>[1] AND nvl(C.�����Ŀ,0)=0 " & _
                            "AND B.������Ŀid=A.��Ŀid and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.������Ŀid IN (SELECT C.ID " & _
                             "FROM ���鱨����Ŀ A," & _
                                  "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = [1]) B," & _
                                  "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                            "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> [1] AND " & _
                                  "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID=[1] AND nvl(C.�����Ŀ,0)=0 AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.����걾) = rs("����").Value
            Else
                
                'û�ж�Ӧʱ����ȡ���б걾����
                gstrSQL = "SELECT ���� FROM ���Ƽ���걾 A where rownum<2"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.����걾) = rs("����").Value
                End If
                
            End If
        
        Case "�շ�ִ�п���"
            
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p���)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.ҩƷִ�п���)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p���))
            Else
                gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
            End If
            If rs.BOF = False Then
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���) = zlCommFun.NVL(rs("����").Value)
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���id) = zlCommFun.NVL(rs("ID").Value)
            Else
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���) = vsf.TextMatrix(vsf.Row, mCol.ִ�п���)
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���id) = vsf.TextMatrix(vsf.Row, mCol.ִ�п���id)
            End If
            
        Case "�Ƽ���Ŀ"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "���" Then
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ�����) = "1"
            Else
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                If Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                End If
            End If

        End Select
    Next
    
    SetRowDefault = True
    
    Exit Function
    
errHand:
    
End Function

Private Function SaveItems(ByVal strGroup As String) As Boolean
    
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '������ѡ��ļ�����Ŀ
    mrsItems.Filter = ""
    mrsItems.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
    
    Call DeleteRecord(mrsItems)
    
    For lngLoop = 1 To vsf.Rows - 1
        
        If Val(vsf.RowData(lngLoop)) > 0 Then
            mrsItems.AddNew
            
            mrsItems("���").Value = strGroup
            mrsItems("ID").Value = vsf.RowData(lngLoop)
            mrsItems("���").Value = vsf.TextMatrix(lngLoop, mCol.���)
            mrsItems("����").Value = vsf.TextMatrix(lngLoop, mCol.��Ŀ)
            mrsItems("ִ�п���").Value = vsf.TextMatrix(lngLoop, mCol.ִ�п���)
            mrsItems("��鲿λ").Value = vsf.TextMatrix(lngLoop, mCol.��鲿λ)
            mrsItems("�ɼ���ʽ").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ���ʽ)
            mrsItems("�ɼ�����").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ�����)
            mrsItems("����걾").Value = vsf.TextMatrix(lngLoop, mCol.����걾)
            mrsItems("�������").Value = vsf.TextMatrix(lngLoop, mCol.�������)
            mrsItems("�����۸�").Value = vsf.TextMatrix(lngLoop, mCol.�����۸�)
            mrsItems("���۸�").Value = vsf.TextMatrix(lngLoop, mCol.���۸�)
            mrsItems("���㷽ʽ").Value = vsf.TextMatrix(lngLoop, mCol.���㷽ʽ)
            mrsItems("ִ�п���id").Value = vsf.TextMatrix(lngLoop, mCol.ִ�п���id)
            mrsItems("�ɼ���ʽid").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ���ʽid)
            mrsItems("�ɼ�����id").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ�����id)
            mrsItems("��鲿λid").Value = vsf.TextMatrix(lngLoop, mCol.��鲿λid)
            mrsItems("�Ʒ���ϸ").Value = vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ)
            
            mrsItems("�¼�").Value = vsf.TextMatrix(lngLoop, mCol.�¼�)
            mrsItems("ǰ��ɫ").Value = vsf.TextMatrix(lngLoop, mCol.ǰ��ɫ)
            mrsItems("ɾ��").Value = ""
            mrsItems("����").Value = vsf.TextMatrix(lngLoop, mCol.����)
            
        End If
    Next
    
    SaveItems = True
    
errHand:

End Function

Private Function WritePersons(ByVal strGroup As String, Optional bytMode As Byte = 1) As Boolean
    
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '������ѡ��ļ�����Ŀ
    If bytMode = 1 Then
        mrsPersons.Filter = ""
        mrsPersons.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
    
        Call DeleteRecord(mrsPersons)
    
        For lngLoop = 1 To vsfPerson.Rows - 1
            
            If vsfPerson.TextMatrix(lngLoop, mPersonCol.����) <> "" Then
                mrsPersons.AddNew
                
                mrsPersons("���").Value = strGroup
                
                With vsfPerson
                    mrsPersons("IC����").Value = .TextMatrix(lngLoop, mPersonCol.IC����)
                    mrsPersons("������").Value = .TextMatrix(lngLoop, mPersonCol.������)
                    mrsPersons("����id").Value = .TextMatrix(lngLoop, mPersonCol.����id)
                    mrsPersons("����").Value = .TextMatrix(lngLoop, mPersonCol.����)
                    mrsPersons("�����").Value = .TextMatrix(lngLoop, mPersonCol.�����)
                    mrsPersons("���֤").Value = .TextMatrix(lngLoop, mPersonCol.���֤)
                    mrsPersons("�Ա�").Value = .TextMatrix(lngLoop, mPersonCol.�Ա�)
                    mrsPersons("��������").Value = .TextMatrix(lngLoop, mPersonCol.��������)
                    mrsPersons("����״��").Value = .TextMatrix(lngLoop, mPersonCol.����״��)
                    mrsPersons("����").Value = .TextMatrix(lngLoop, mPersonCol.����)
                    mrsPersons("����").Value = .TextMatrix(lngLoop, mPersonCol.����)
                    mrsPersons("����").Value = .TextMatrix(lngLoop, mPersonCol.����)
                    mrsPersons("ѧ��").Value = .TextMatrix(lngLoop, mPersonCol.ѧ��)
                    mrsPersons("ְҵ").Value = .TextMatrix(lngLoop, mPersonCol.ְҵ)
                    mrsPersons("���").Value = .TextMatrix(lngLoop, mPersonCol.���)
                    mrsPersons("��ϵ������").Value = .TextMatrix(lngLoop, mPersonCol.��ϵ������)
                    mrsPersons("��ϵ�˵绰").Value = .TextMatrix(lngLoop, mPersonCol.��ϵ�˵绰)
                    mrsPersons("�����ʼ�").Value = .TextMatrix(lngLoop, mPersonCol.�����ʼ�)
                    mrsPersons("��ϵ�˵�ַ").Value = .TextMatrix(lngLoop, mPersonCol.��ϵ�˵�ַ)
                    mrsPersons("������λ").Value = .TextMatrix(lngLoop, mPersonCol.������λ)
                    mrsPersons("�Ǽ�ʱ��").Value = .TextMatrix(lngLoop, mPersonCol.�Ǽ�ʱ��)
                    mrsPersons("���￨��").Value = .TextMatrix(lngLoop, mPersonCol.���￨��)
                    mrsPersons("ɾ��").Value = ""
        
                End With
                
            End If
        Next
    End If
    
    If bytMode = 2 Then
        mrsPersons.Filter = ""
        mrsPersons.Filter = "ɾ��<>'1'"
        If mrsPersons.RecordCount = 0 Then mrsPersons.AddNew
                
        mrsPersons("���").Value = "ȱʡ"
        mrsPersons("����id").Value = Val(cmd(1).Tag)
        mrsPersons("����").Value = txt(5).Text
        mrsPersons("���֤").Value = txt(4).Text
        mrsPersons("�Ա�").Value = zlCommFun.GetNeedName(cbo(1).Text)
        mrsPersons("����").Value = txt(9).Text
        mrsPersons("����״��").Value = zlCommFun.GetNeedName(cbo(0).Text)
        mrsPersons("��ϵ�˵绰").Value = txt(10).Text

        mrsPersons("�����").Value = txt(3).Text
        mrsPersons("������").Value = txt(14).Text
    End If
    WritePersons = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    
    If mbytMode = 1 Then
        '���ԤԼ��
        If Trim(txt(0).Text) = "" Then
            ShowSimpleMsg "ԤԼ�˲���Ϊ��ֵ���������룡"
            Call LocationObj(txt(0))
            Exit Function
        End If
        
        If mblnGroup = False Then
            If Format(dtp(0).Value, "yyyy-MM-dd") < Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
                ShowSimpleMsg "ԤԼ���ʱ�䲻��С�ڵ��죡"
                dtp(0).SetFocus
                Exit Function
            End If
        End If
    End If
    
    '�������
    If fraGroupInfo.Visible And mblnGroup Then
        If Trim(txt(13).Text) = "" Then
            ShowSimpleMsg "����Ҫȷ�����壡"
            Call LocationObj(txt(13))
            Exit Function
        End If
    End If
            
    '�����������Ƿ���Ч
    For lngLoop = 1 To lvwGroup.ListItems.Count
        If Trim(lvwGroup.ListItems(lngLoop).Text) = "" Then
            ShowSimpleMsg "��������Ϊ�գ�"
            lvwGroup.SetFocus
            Exit Function
        End If
        
        If StrIsValid(lvwGroup.ListItems(lngLoop).Text, 30) = False Then
            lvwGroup.SetFocus
            Exit Function
        End If
        
    Next
    
    '�������ʼ�
    If mblnGroup Then
        If CheckStrValid(txt(8).Text, CHECKFORMAT.�����ʼ�) = False Then
        
            ShowSimpleMsg "����ĵ����ʼ���ʽ�������ʼ���ʽ���£�" & vbCrLf & "1.�������@�ַ���" & vbCrLf & "2.@�ַ�ֻ�����м䣬�� xxx@163.com��"
            Call LocationObj(txt(8))
            Exit Function
            
        End If
    Else
        
        If CheckStrValid(txt(7).Text, CHECKFORMAT.�����ʼ�) = False Then
        
            ShowSimpleMsg "����ĵ����ʼ���ʽ�������ʼ���ʽ���£�" & vbCrLf & "1.�������@�ַ���" & vbCrLf & "2.@�ַ�ֻ�����м䣬�� xxx@163.com��"
            Call LocationObj(txt(7))
            Exit Function
            
        End If
        
    End If
    
    If mbytMode = 2 Then
        
        mrsItems.Filter = ""
        mrsItems.Filter = "ID>0"
        If mrsItems.RecordCount = 0 Then
            ShowSimpleMsg "��ǰû�������Ŀ"
            tbs.Tab = 0
            Call tbs_Click(1)
            vsf.SetFocus
            Exit Function
        End If
        
        If mblnGroup = False Then
            If Trim(txt(5).Text) = "" Then
                ShowSimpleMsg "��ǰ��컹û�����������Ա"
                
                Call LocationObj(txt(5))
                Exit Function
            End If
        Else
                mrsPersons.Filter = ""
                mrsPersons.Filter = "����<>''"
                If mrsPersons.RecordCount = 0 Then
                    ShowSimpleMsg "��ǰ��컹û�����������Ա"
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    vsfPerson.SetFocus
                    Exit Function
                End If
        End If
        
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) < 0 Then
            tbs.Tab = 0
            Call tbs_Click(1)
            
            ShowSimpleMsg "���۸���Ϊ����"
            vsf.Row = lngLoop
            vsf.Col = mCol.���۸�
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
        
'        If Format(Val(vsf.TextMatrix(lngLoop, mCol.���۸�)), "0.00") > Format(Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)), "0.00") Then
'
'            tbs.Tab = 0
'            Call tbs_Click(1)
'
'            ShowSimpleMsg "���۸��ܴ��ڻ����۸�"
'            vsf.Row = lngLoop
'            vsf.Col = mCol.���۸�
'            vsf.ShowCell vsf.Row, vsf.Col
'            vsf.SetFocus
'
'            Exit Function
'        End If
        
    Next
    
    If mblnGroup Then
        For lngLoop = 1 To vsfPerson.Rows - 1
            
            '���������Ƿ����
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����)) <> "" And Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)) = 0 Then
                gstrSQL = "Select 1 From ������Ϣ Where �����=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����)))
                If rs.BOF = False Then
                    
                    ShowSimpleMsg "��ǰ����ţ�" & Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����)) & "�Ѿ����ڣ��������ظ���"
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.�����
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
                End If
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.����), GetMaxLength("������Ϣ", "����")) = False Then
                
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.����
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.���֤), GetMaxLength("������Ϣ", "���֤��")) = False Then
                
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.���֤
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.����״��), GetMaxLength("������Ϣ", "����״��")) = False Then
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.����״��
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����ʼ�), GetMaxLength("�����Ա����", "�����ʼ�")) = False Then
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.�����ʼ�
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.�Ա�), GetMaxLength("������Ϣ", "�Ա�")) = False Then
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.�Ա�
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
        
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.��������)) <> "" Then
                
                If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.��������), CHECKFORMAT.����) = False Then
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                    ShowSimpleMsg "�Ƿ��ĳ������ڣ�"
                    
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.��������
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
                End If
            End If
            
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����ʼ�), CHECKFORMAT.�����ʼ�) = False Then
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                    ShowSimpleMsg "�����ʼ�������� @ ���ţ��Ҳ��ڵ�1λ�����һλ�ϣ�"
                    
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.�����ʼ�
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
                
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.���֤), CHECKFORMAT.���֤��) = False Then
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                    ShowSimpleMsg "���֤�ŷǷ�������Ϊ15λ��18λ��Ϊ0-9��X�ַ�����"
                    
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.���֤
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
        Next
    End If
    
    ValidEdit = True
    
End Function

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim lngRow As Long
    Dim strSQL() As String
    Dim strNow As String
    Dim rsPati As New ADODB.Recordset
    Dim lng����id As Long
    Dim strGroup As String
    Dim intCount1 As Integer
    Dim str����� As String
    Dim intCount2 As Integer
    Dim bytNew As Byte
    Dim strRegisteDate As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If mblnGroup Then
        If Val(cmd(4).Tag) = 0 Then
            '��������
            cmd(4).Tag = zlDatabase.GetNextId("��Լ��λ")
            
            gstrSQL = "zl_��Լ��λ_Insert(" & Val(cmd(4).Tag) & "," & _
                                            "NULL,'" & _
                                            IIf(zlCommFun.NVL(mrsGroup("����")) = "", GetNextCode("��Լ��λ", "����", ""), mrsGroup("����")) & "','" & _
                                            txt(13).Text & "','" & _
                                            zlCommFun.SpellCode(txt(13).Text) & "'," & _
                                            IIf(IsNull(mrsGroup("��ַ").Value), "NULL", "'" & mrsGroup("��ַ").Value & "'") & ",'" & _
                                            txt(11).Text & "'," & _
                                            IIf(IsNull(mrsGroup("��������").Value), "NULL", "'" & mrsGroup("��������").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("�ʺ�").Value), "NULL", "'" & mrsGroup("�ʺ�").Value & "'") & ",'" & _
                                            txt(12).Text & "'," & _
                                            "1," & _
                                            IIf(IsNull(mrsGroup("�����ʼ�").Value), "NULL", "'" & mrsGroup("�����ʼ�").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("˵��").Value), "NULL", "'" & mrsGroup("˵��").Value & "'") & _
                                            ")"
                                            
            strSQL(ReDimArray(strSQL)) = gstrSQL
        Else
            '�޸�����

            gstrSQL = "zl_��Լ��λ_Update(" & Val(cmd(4).Tag) & "," & _
                                            IIf(IsNull(mrsGroup("�ϼ�ID").Value), "NULL", mrsGroup("�ϼ�ID").Value) & "," & _
                                            IIf(IsNull(mrsGroup("����").Value), "NULL", "'" & mrsGroup("����").Value & "'") & ",'" & _
                                            txt(13).Text & "','" & _
                                            zlCommFun.SpellCode(txt(13).Text) & "'," & _
                                            IIf(IsNull(mrsGroup("��ַ").Value), "NULL", "'" & mrsGroup("��ַ").Value & "'") & ",'" & _
                                            txt(11).Text & "'," & _
                                            IIf(IsNull(mrsGroup("��������").Value), "NULL", "'" & mrsGroup("��������").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("�ʺ�").Value), "NULL", "'" & mrsGroup("�ʺ�").Value & "'") & ",'" & _
                                            txt(12).Text & _
                                            "',0," & _
                                            IIf(IsNull(mrsGroup("�����ʼ�").Value), "NULL", "'" & mrsGroup("�����ʼ�").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("˵��").Value), "NULL", "'" & mrsGroup("˵��").Value & "'") & _
                                            ")"
            strSQL(ReDimArray(strSQL)) = gstrSQL

        End If
    End If
    
    If mlngKey = 0 Then
        
        'ȡ����
        txt����.Text = GetNextNo(78)
        
        '����ԤԼ
        If Val(tbs.Tag) > 0 Then
            lngKey = Val(tbs.Tag)
        Else
            lngKey = zlDatabase.GetNextId("���ǼǼ�¼")
        End If

        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_INSERT(" & lngKey & ",'" & _
                                                            txt����.Text & "'," & _
                                                            "1," & _
                                                            "1,'" & _
                                                            txt(0).Text & "','" & _
                                                            txt(1).Text & "'," & _
                                                            "NULL,'" & _
                                                            txt(2).Text & "'," & _
                                                            IIf(Val(cmd(4).Tag) = 0, "NULL", Val(cmd(4).Tag)) & "," & _
                                                            "1," & _
                                                            "TO_DATE('" & Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00','yyyy-mm-dd hh24:mi:ss')," & _
                                                            mlngDept & ",'" & _
                                                            txt(6).Text & "'," & _
                                                            "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "NULL," & _
                                                            IIf(mblnGroup, 1, 0) & "," & _
                                                            "1," & _
                                                            IIf(chk.Value = 1, Val(txt(31).Text), "NULL") & ")"
        
        '�����������....
        
        
    Else
        '�޸�ԤԼ
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_UPDATE(" & lngKey & ",'" & _
                                                            txt����.Text & "'," & _
                                                            "1," & _
                                                            "1,'" & _
                                                            txt(0).Text & "','" & _
                                                            txt(1).Text & "'," & _
                                                            "NULL,'" & _
                                                            txt(2).Text & "'," & _
                                                            IIf(Val(cmd(4).Tag) = 0, "NULL", Val(cmd(4).Tag)) & "," & _
                                                            "1," & _
                                                            "TO_DATE('" & Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00','yyyy-mm-dd hh24:mi:ss')," & _
                                                            mlngDept & ",'" & _
                                                            txt(6).Text & "'," & _
                                                            "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "NULL," & _
                                                            "1," & _
                                                            IIf(chk.Value = 1, Val(txt(31).Text), "NULL") & ")"
                                                            
    End If
    
    strSQL(ReDimArray(strSQL)) = "zl_�����Ա����_Delete(" & lngKey & ")"
    
    strGroup = ""
    mrsPersons.Filter = ""
    If mrsPersons.RecordCount > 0 Then
        mrsPersons.Filter = ""
        If mblnGroup Then mrsPersons.Sort = "���"
        If mrsPersons.RecordCount > 0 Then mrsPersons.MoveFirst
        
        Dim intCount As Integer

        intCount = -1
        Do While Not mrsPersons.EOF
            
            '����������
            If mrsPersons("��������") <> "" Then
                
                If CheckStrValid(mrsPersons("��������"), CHECKFORMAT.����) = False Then
                    ShowSimpleMsg mrsPersons("����").Value & "�ĳ���������Ч��"
                    Exit Function
                End If
            End If
            
            If mblnGroup Then
                If strGroup <> mrsPersons("���").Value Then strGroup = mrsPersons("���").Value
            Else
                strGroup = "ȱʡ"
            End If
            
            lng����id = zlCommFun.NVL(mrsPersons("����id"), 0)
            bytNew = 0
            If lng����id = 0 Then
                bytNew = 1
                intCount = intCount + 1
                'lng����id = GetNextPatientID + intCount
                lng����id = GetNextNo(1) + intCount
            End If
            
            intCount1 = intCount1 + 1
            
            If zlCommFun.NVL(mrsPersons("�����").Value, 0) < 1 Then
                'lng����� = NextNo(3) + intCount2
                str����� = CStr(GetNextNo(3) + intCount2)
                intCount2 = intCount2 + 1
            Else
                str����� = CStr(zlCommFun.NVL(mrsPersons("�����").Value, 0))
            End If
            
            If zlCommFun.NVL(mrsPersons("�Ǽ�ʱ��").Value, "") <> "" Then
                strRegisteDate = mrsPersons("�Ǽ�ʱ��").Value
                strRegisteDate = "To_Date('" & Format(strRegisteDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
            Else
                strRegisteDate = "Null"
            End If
            
            strSQL(ReDimArray(strSQL)) = "ZL_�����Ա����_INSERT(" & lngKey & "," & _
                                                            lng����id & "," & _
                                                            "'" & strGroup & "','" & _
                                                            mrsPersons("����").Value & "','" & _
                                                            mrsPersons("���֤").Value & "','" & _
                                                            mrsPersons("�Ա�").Value & "'," & _
                                                            IIf(mrsPersons("��������").Value = "", "NULL", "TO_DATE('" & mrsPersons("��������").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                            mrsPersons("����״��").Value & "','" & _
                                                            mrsPersons("����").Value & "','" & _
                                                            mrsPersons("����").Value & "','" & _
                                                            mrsPersons("ѧ��").Value & "','" & _
                                                            mrsPersons("ְҵ").Value & "','" & _
                                                            mrsPersons("��ϵ������").Value & "','" & _
                                                            mrsPersons("��ϵ�˵绰").Value & "','" & _
                                                            mrsPersons("�����ʼ�").Value & "','" & _
                                                            mrsPersons("��ϵ�˵�ַ").Value & "','" & _
                                                            mrsPersons("������λ").Value & "','" & _
                                                            mrsPersons("����").Value & "'," & _
                                                            Val(str�����) & ",'" & _
                                                            mrsPersons("IC����").Value & "','" & _
                                                            mrsPersons("������").Value & "','" & _
                                                            mrsPersons("���￨��").Value & "'," & _
                                                            "1," & _
                                                            IIf(intCount1 = mrsPersons.RecordCount, "1", "0") & ",0," & bytNew & "," & strRegisteDate & _
                                                            ")"
            mrsPersons.MoveNext
        Loop
    End If

    
    '����ѡ��������Ŀ
    strSQL(ReDimArray(strSQL)) = "ZL_�����Ŀ�嵥_DELETE(" & lngKey & ")"
    strSQL(ReDimArray(strSQL)) = "ZL_������_DELETE(" & lngKey & ")"
        
    For lngLoop = 1 To lvwGroup.ListItems.Count
        If Trim(lvwGroup.ListItems(lngLoop).Text) <> "" Then
            strSQL(ReDimArray(strSQL)) = "ZL_������_INSERT(" & lngKey & ",'" & Trim(lvwGroup.ListItems(lngLoop).Text) & "')"
        End If
    Next
    
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    
    mrsItems.Filter = ""
    If mrsItems.RecordCount > 0 Then
        
        mrsItems.Filter = ""
        mrsItems.Sort = "���"
        If mrsItems.RecordCount > 0 Then mrsItems.MoveFirst
        
        strGroup = ""
        
        Do While Not mrsItems.EOF
            
            If strGroup <> mrsItems("���").Value Then strGroup = mrsItems("���").Value
                        
            If mrsItems("ID").Value > 0 Then
                
                strTmp = ""
                varRow = Split(mrsItems("�Ʒ���ϸ").Value, ";")
                For lngLoop = 0 To UBound(varRow)
                    
                    varCol = Split(varRow(lngLoop), ":")
                    
                    If strTmp <> "" Then strTmp = strTmp & ";"
                    strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                    
                Next
                                
                strSQL(ReDimArray(strSQL)) = "ZL_�����Ŀ�嵥_INSERT(" & lngKey & "," & _
                                            "'" & strGroup & "'," & _
                                            mrsItems("ID").Value & ",'" & _
                                            mrsItems("�������").Value & "'," & _
                                            Val(mrsItems("�����۸�").Value) & "," & _
                                            Val(mrsItems("���۸�").Value) & "," & _
                                            mrsItems("ִ�п���id").Value & "," & _
                                            IIf(mrsItems("�ɼ���ʽid") = "", "NULL", mrsItems("�ɼ���ʽid")) & "," & _
                                            IIf(mrsItems("�ɼ�����id") = "", "NULL", mrsItems("�ɼ�����id")) & ",'" & _
                                            mrsItems("����걾").Value & "','" & _
                                            mrsItems("��鲿λ").Value & "','" & _
                                            mrsItems("��鲿λid").Value & "',NULL," & IIf(mrsItems("���㷽ʽ").Value = "����", "1", "2") & ",'" & _
                                            strTmp & "')"
            End If
            
            mrsItems.MoveNext
        Loop
    End If
    
    strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_�������(" & lngKey & ")"
    
    '��������Ǽ�ʱ����Ҫ���� ԤԼȷ��
    If mbytMode = 2 Then
        
        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_STATE(" & lngKey & ",2)"
        
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function ReadItems(ByVal strGroup As String) As Boolean
    
    mrsItems.Filter = ""
    mrsItems.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
    If mrsItems.RecordCount > 0 Then
        mrsItems.MoveFirst
        Call FillGrid(vsf, mrsItems)
    End If
    Call ReadPrice(vsf.Row)
    
    Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
    Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
    
    Call ChangeItem(Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)), Val(vsf.TextMatrix(vsf.Row, mCol.���۸�)), 1, False)
    
    ReadItems = True
    
End Function

Private Function ReadPersons(ByVal strGroup As String, Optional ByVal bytMode As Byte = 1) As Boolean
    Dim lngLoop As Long
    
    If bytMode = 1 Then
        mrsPersons.Filter = ""
        mrsPersons.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
        If mrsPersons.RecordCount > 0 Then
            mrsPersons.MoveFirst
            Call FillGrid(vsfPerson, mrsPersons)
        End If
        
        For lngLoop = 1 To vsfPerson.Rows - 1
            
            If Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)) = 0 Then
                vsfPerson.Cell(flexcpForeColor, lngLoop, 0, lngLoop, vsfPerson.Cols - 1) = COLOR.��ɫ
            End If
        Next
        
    End If
    
    If bytMode = 2 Then
        
        cmd(1).Tag = Val(mrsPersons("����id").Value)
        
        txt(5).Text = mrsPersons("����").Value
        txt(4).Text = mrsPersons("���֤").Value
                
        txt(9).Text = mrsPersons("����").Value
        txt(10).Text = mrsPersons("��ϵ�˵绰").Value
        
        zlControl.CboLocate cbo(1), zlCommFun.NVL(mrsPersons("�Ա�").Value)
        zlControl.CboLocate cbo(0), zlCommFun.NVL(mrsPersons("����״��").Value)
        
        txt(3).Text = zlCommFun.NVL(mrsPersons("�����").Value)
        txt(14).Text = zlCommFun.NVL(mrsPersons("������").Value)
        
        imgNew(0).Visible = (Val(cmd(1).Tag) = 0)
        
        txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
        
    End If
    
    ReadPersons = True
    
End Function

Private Function ReadTemplate(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    
    Dim strKeys As String
    Dim bytParam1 As Byte
    Dim bytParam2 As Byte
    
    bytParam1 = 1
    bytParam2 = 2
            
    If mblnGroup = False Then
        Select Case zlCommFun.GetNeedName(cbo(1).Text)
        Case "��"
            bytParam1 = 1
            bytParam2 = 1
        Case "Ů"
            bytParam1 = 2
            bytParam2 = 2
        End Select
    End If
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT DISTINCT A.ID,DECODE(A.���,'C','����','D','���') AS ���,A.����,A.����,C.���� AS �������,D.���� As �ɼ���ʽ,B.�ɼ���ʽid,B.����걾,B.��鲿λ,B.��鲿λid " & _
                "FROM ������ĿĿ¼ A,�������Ŀ¼ B,������� C,������ĿĿ¼ D " & _
                "WHERE A.ID=B.������ĿID AND C.���=B.��� AND D.ID(+)=B.�ɼ���ʽid AND B.���=[1] And Nvl(a.�����Ա�,0) In (0,[2],[3])"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, bytParam1, bytParam2)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            vsf.Row = vsf.Rows - 1
            If Val(vsf.RowData(vsf.Row)) > 0 Then
                vsf.Rows = vsf.Rows + 1
                vsf.Row = vsf.Rows - 1
            End If
            
            If CheckHave(rs("ID").Value) = False Then
            
                vsf.TextMatrix(vsf.Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf.TextMatrix(vsf.Row, mCol.��Ŀ) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(vsf.Row, mCol.�������) = zlCommFun.NVL(rs("�������").Value)
                
                vsf.TextMatrix(vsf.Row, mCol.����걾) = zlCommFun.NVL(rs("����걾").Value)
                vsf.TextMatrix(vsf.Row, mCol.��鲿λ) = zlCommFun.NVL(rs("��鲿λ").Value)
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("�ɼ���ʽ").Value)
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("�ɼ���ʽid").Value)
                vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = zlCommFun.NVL(rs("��鲿λid").Value)
                
                vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            End If
                        
            If vsf.TextMatrix(vsf.Row, mCol.���) = "����" Then
                
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "ִ�п���")
                
                If Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)) = 0 Then
                    Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "�ɼ���ʽ")
                End If
                
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "�ɼ�����")
                
                If Trim(vsf.TextMatrix(vsf.Row, mCol.����걾)) = "" Then
                    Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "����걾")
                End If
                
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "���㷽ʽ", "�Ƽ���Ŀ")
                                
            ElseIf vsf.TextMatrix(vsf.Row, mCol.���) = "���" Then
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
            End If
            
            gstrSQL = "Select z.����,y.����,y.���㵥λ,x.�ּ�,x.�ּ�*Nvl(z.�ۿ�,1) As ��쵥��,y.id,Nvl(z.�Ƽ�����,1) As �Ƽ�����,y.���,10*Nvl(z.�ۿ�,1) As �ۿ� " & _
                        "From " & _
                            "( Select a.���,a.������Ŀid,a.�շ�ϸĿid,Sum(c.�ּ�) As �ּ� " & _
                              "From �շѼ�Ŀ c, " & _
                                   "������ͼƼ� a " & _
                              "Where a.�շ�ϸĿid = c.�շ�ϸĿid " & _
                                    "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                                    "and A.���=[2] " & _
                                    "and A.������Ŀid=[1] " & _
                              "Group by a.���,a.������Ŀid,a.�շ�ϸĿid " & _
                            ") x, " & _
                            "�շ���ĿĿ¼ y, " & _
                            "������ͼƼ� z " & _
                        "Where x.�շ�ϸĿid = y.ID " & _
                              "and z.���=x.��� " & _
                              "and z.������Ŀid=x.������Ŀid " & _
                              "and z.�շ�ϸĿid=x.�շ�ϸĿid "
                        
            Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), lngKey)
            If rsPrice.BOF = False Then
                With vsfPrice
                    Do While Not rsPrice.EOF
                        
                        If Val(.TextMatrix(.Rows - 1, mCol.p�շ���Ŀid)) > 0 Then
                            .Rows = .Rows + 1
                        End If
                        
                        .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rsPrice("����"))
                        .TextMatrix(.Rows - 1, mCol.p���㵥λ) = zlCommFun.NVL(rsPrice("���㵥λ"))
                        .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rsPrice("����"))
                        .TextMatrix(.Rows - 1, mCol.p��׼����) = zlCommFun.NVL(rsPrice("�ּ�"))
                        .TextMatrix(.Rows - 1, mCol.p��쵥��) = zlCommFun.NVL(rsPrice("��쵥��"))
                        .TextMatrix(.Rows - 1, mCol.p�ۿ�) = zlCommFun.NVL(rsPrice("�ۿ�"))
                        .TextMatrix(.Rows - 1, mCol.p��׼���) = zlCommFun.NVL(rsPrice("����"), 0) * zlCommFun.NVL(rsPrice("�ּ�"), 0)
                        .TextMatrix(.Rows - 1, mCol.p�����) = zlCommFun.NVL(rsPrice("����"), 0) * zlCommFun.NVL(rsPrice("��쵥��"), 0)
                        .TextMatrix(.Rows - 1, mCol.p�շ���Ŀid) = zlCommFun.NVL(rsPrice("ID"))
                        .TextMatrix(.Rows - 1, mCol.p�Ƽ�����) = zlCommFun.NVL(rsPrice("�Ƽ�����"))
                        .RowData(.Rows - 1) = zlCommFun.NVL(rsPrice("ID"), 0)
                        .TextMatrix(.Rows - 1, mCol.p���) = zlCommFun.NVL(rsPrice("���"))
                        
                        If zlCommFun.NVL(rsPrice("�Ƽ�����"), 1) = 2 Then
                            .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                        ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
                            .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                        Else
                            .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                        End If
                        
                        Call SetRowDefault(Val(.RowData(.Rows - 1)), vsfPrice.Rows - 1, "�շ�ִ�п���")
                        
                        If InStr("567", .TextMatrix(.Rows - 1, mCol.p���)) > 0 Then
                            .TextMatrix(.Rows - 1, mCol.p���ÿ��) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.pִ�п���id)))
                            Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p����)), Val(.TextMatrix(.Rows - 1, mCol.p���ÿ��)), .TextMatrix(.Rows - 1, mCol.p����), .TextMatrix(.Rows - 1, mCol.pִ�п���), .TextMatrix(.Rows - 1, mCol.p���㵥λ), 1)
                        End If
                        
                        
                        
                        rsPrice.MoveNext
                    Loop
                    
                    
                End With
                
                Call ChangePrice(Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p��쵥��)), 1)
                
                vsf.TextMatrix(vsf.Row, mCol.�����۸�) = SumPrice(1)
                vsf.TextMatrix(vsf.Row, mCol.���۸�) = SumPrice(2)
                
            End If
            
            Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
            Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
            Call WritePrice(vsf.Row)
                                    
            rs.MoveNext
        Loop
    End If
    
    ReadTemplate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        
    End If
End Sub

Private Sub chk_Click()
    
    txt(31).Enabled = (chk.Value = 1)
    txt(31).BackColor = IIf(chk.Value = 1, &H80000005, &H8000000F)
    If chk.Value <> 1 Then
        txt(31).Text = ""
    End If
    
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strDate As String
    Dim objPoint As POINTAPI
    Dim strTmp As String
    Dim strItem As String
    Dim strValue As String
    Dim strCardNo1 As String
    Dim strCardNo2 As String
    Dim rsPrice As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngLoop As Long
    Dim objItem As ListItem
    Dim intRow As Long
    Dim strKeys As String

    Dim clsCard As Object
    Dim strInfo() As String
    
    On Error GoTo errHand
    
    Call ClientToScreen(cmd(Index).hWnd, objPoint)
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0
        
        strTmp = ""
        If frmInputBox.ShowInputBox(Me, "������������", "�����µ����������ƣ���Ϊ�������������Ŀ���ܼ���Ա��", "���(&G)", strTmp, 1, 20) Then
        
            '�������������,���Ҫ��:1.�������Ƿ��Ѿ�����;2.�޸���Ա����Ŀ��Ӧ���������
        
            For lngLoop = 1 To lvwGroup.ListItems.Count
                If lngLoop <> lvwGroup.SelectedItem.Index Then
                    If Trim(lvwGroup.ListItems(lngLoop).Text) = Trim(strTmp) Then
                        ShowSimpleMsg "��" & strTmp & "������Ѿ����ڣ�"
                        Exit Sub
                    End If
                End If
            Next
            
            Set objItem = lvwGroup.ListItems.Add(, , strTmp, 1, 1)
            objItem.Selected = True
        
            Call lvwGroup_ItemClick(objItem)

        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        If lvwGroup.SelectedItem Is Nothing Then Exit Sub
        
        strTmp = Trim(lvwGroup.SelectedItem.Text)
        If frmInputBox.ShowInputBox(Me, "������", "�޸��Ѿ����ڵ����������ơ�", "������(&G)", strTmp, 1, 20) Then
        
            '�������������,���Ҫ��:1.�������Ƿ��Ѿ�����;2.�޸���Ա����Ŀ��Ӧ���������
            
            If Trim(strTmp) = Trim(lvwGroup.SelectedItem.Text) Then Exit Sub
            
            For lngLoop = 1 To lvwGroup.ListItems.Count
                If lngLoop <> lvwGroup.SelectedItem.Index Then
                    If Trim(lvwGroup.ListItems(lngLoop).Text) = Trim(strTmp) Then
                        ShowSimpleMsg "��" & strTmp & "������Ѿ����ڣ�"
                        Exit Sub
                    End If
                End If
            Next
            
            '2.�޸���Ա����Ŀ��Ӧ���������
            Call WritePrice(vsf.Row)
            Call SaveItems(lvwGroup.SelectedItem.Text)
            Call WritePersons(lvwGroup.SelectedItem.Text)
            
            mrsItems.Filter = ""
            mrsItems.Filter = "���='" & lvwGroup.SelectedItem.Text & "'"
            If mrsItems.RecordCount > 0 Then
                mrsItems.MoveFirst
                Do While Not mrsItems.EOF
                    mrsItems("���").Value = strTmp
                    mrsItems.MoveNext
                Loop
            End If
            mrsItems.Filter = ""
        
            mrsPersons.Filter = ""
            mrsPersons.Filter = "���='" & lvwGroup.SelectedItem.Text & "'"
            If mrsPersons.RecordCount > 0 Then
                mrsPersons.MoveFirst
                Do While Not mrsPersons.EOF
                    mrsPersons("���").Value = strTmp
                    mrsPersons.MoveNext
                Loop
            End If
            mrsPersons.Filter = ""
                    
            lvwGroup.SelectedItem.Text = strTmp
            
            Call ResetVsf(vsf)
            Call ResetVsf(vsfPrice)
            Call ResetVsf(vsfPerson)
            
            mstrGroup = lvwGroup.SelectedItem.Text
            
            Call ReadItems(mstrGroup)
            Call ReadPersons(mstrGroup)
            Call ReadPrice(vsf.Row)
            Call CountGroup
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 3
        
        If lvwGroup.ListItems.Count = 1 Then
            ShowSimpleMsg "�������ʱ������Ҫһ�����"
            Exit Sub
        End If
    
        If MsgBox("ɾ�����ʱ�����Զ�����������Ϣ��" & vbCrLf & "  1.ɾ����Ӧ�������Ŀ" & vbCrLf & "  2.ɾ������Ҫ����������������Ա����" & vbCrLf & "������", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
                
        mrsItems.Filter = ""
        mrsItems.Filter = "���='" & lvwGroup.SelectedItem.Text & "'"
        If mrsItems.RecordCount > 0 Then
            mrsItems.MoveFirst
            Call DeleteRecord(mrsItems)
        End If
        
        Call WritePersons(lvwGroup.SelectedItem.Text)
        
        mrsPersons.Filter = ""
        mrsPersons.Filter = "���='" & lvwGroup.SelectedItem.Text & "'"
        If mrsPersons.RecordCount > 0 Then
            mrsPersons.MoveFirst
            Do While Not mrsPersons.EOF
                If lvwGroup.SelectedItem.Index = 1 Then
                    mrsPersons("���").Value = lvwGroup.ListItems(2).Text
                Else
                    mrsPersons("���").Value = lvwGroup.ListItems(1).Text
                End If
                mrsPersons.MoveNext
            Loop
        End If
        mrsPersons.Filter = ""
    
        Call ResetVsf(vsfPerson)
        Call FillGrid(vsfPerson, mrsPersons)
        
        lngLoop = lvwGroup.SelectedItem.Index
        lvwGroup.ListItems.Remove lngLoop
        Call NextLvwPos(lvwGroup, lngLoop)
        
        If Not (lvwGroup.SelectedItem Is Nothing) Then
            mstrGroup = lvwGroup.SelectedItem.Text
        
            Call ReadItems(mstrGroup)
            Call ReadPersons(mstrGroup)
            Call ReadPrice(vsf.Row)
            
        End If
        Call CountGroup
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '�򿪲��˲��ҶԻ���
        If frmPatientFind.ShowFind(Me, lngKey) Then
            If lngKey > 0 Then
                
                gstrSQL = "SELECT A.* FROM ������Ϣ A WHERE A.����id=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
                If rs.BOF = False Then
                    cmd(1).Tag = zlCommFun.NVL(rs("����id").Value)
                    txt(5).Text = zlCommFun.NVL(rs("����").Value)
                    txt(4).Text = zlCommFun.NVL(rs("���֤��").Value)
                    txt(9).Text = zlCommFun.NVL(rs("����").Value)
                    
                    txt(3).Text = zlCommFun.NVL(rs("�����").Value)
                    txt(14).Text = zlCommFun.NVL(rs("������").Value)
                    
                    zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("�Ա�").Value)
                    zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("����״��").Value)
                    
                    Call FillPatient(Val(cmd(1).Tag))
                    
                    DataChange = True
                    
                    txt(5).Tag = ""
                    imgNew(1).Visible = False
                    
                    txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
                    
                End If
                
            End If
        End If
        
        LocationObj txt(5)
    '------------------------------------------------------------------------------------------------------------------
    Case 4      '������(��ͬ��λ)ѡ����
        lngKey = Val(cmd(Index).Tag)
        gstrSQL = GetPublicSQL(SQL.�������ѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If ShowTxtSelect(Me, txt(13), "����,900,0,1;����,1500,0,1;����,900,0,1;��ַ,3000,0,1", Me.Name & "\�������ѡ��", "�����±���ѡ��һ������/��λ��", rsData, rs, 8790, 5100) Then
              
            Call ReadGroup(zlCommFun.NVL(rs("ID").Value, 0))
                        
            If lngKey <> Val(cmd(Index).Tag) Then DataChange = True

            imgNew(0).Visible = False
            
            txt(0).Text = txt(12).Text
            txt(1).Text = txt(11).Text
            
        End If
        
        LocationObj txt(13)
    '------------------------------------------------------------------------------------------------------------------
    Case 5
    
        Dim bytParam1 As Byte
        Dim bytParam2 As Byte
        
        bytParam1 = 1
        bytParam2 = 2
                
        If mblnGroup = False Then
            Select Case zlCommFun.GetNeedName(cbo(1).Text)
            Case "��"
                bytParam1 = 1
                bytParam2 = 1
            Case "Ů"
                bytParam1 = 2
                bytParam2 = 2
            End Select
        End If
            
        gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, bytParam1, bytParam2)
        
        If ShowTxtSelect(Me, cmd(Index), "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then

                rs.MoveFirst
                Do While Not rs.EOF
                    'ѡȡ��һ����Ŀ
                    vsf.Row = 0

                    If CheckHave(zlCommFun.NVL(rs("ID").Value)) = False Then

                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                            vsf.Rows = vsf.Rows + 1
                        End If
                        intRow = vsf.Rows - 1
                        vsf.Row = vsf.Rows - 1

                        vsf.Cell(flexcpText, intRow, mCol.��Ŀ + 1, intRow, vsf.Cols - 1) = ""

                        vsf.TextMatrix(intRow, mCol.���) = zlCommFun.NVL(rs("���").Value)
                        vsf.TextMatrix(intRow, mCol.��Ŀ) = zlCommFun.NVL(rs("����").Value)
                        vsf.RowData(intRow) = zlCommFun.NVL(rs("ID").Value)

                        If vsf.TextMatrix(intRow, mCol.���) = "����" Then
                            Call SetRowDefault(Val(vsf.RowData(intRow)), intRow, "ִ�п���", "�ɼ���ʽ", "�ɼ�����", "����걾", "���㷽ʽ", "�Ƽ���Ŀ")
                            
                        ElseIf vsf.TextMatrix(intRow, mCol.���) = "���" Then
                            Call SetRowDefault(Val(vsf.RowData(intRow)), intRow, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
                        End If
                            
                        Call CreatePriceList(intRow)
                        Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                        Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
                        
                        Call WritePrice(intRow)

                        DataChange = True
                    End If

                    rs.MoveNext
                Loop
            End If

        End If

        EnterFocus vsf
    '------------------------------------------------------------------------------------------------------------------
    Case 6
    
        gstrSQL = GetPublicSQL(SQL.������ͷ���ѡ��)

        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(mblnGroup, 2, 1))
    
        If ShowTxtSelect(Me, cmd(Index), "����,1080,0,1;����,2400,0,0;����,900,0,0;˵��,1500,0,0", Me.Name & "\�������ѡ��", "����б���ѡ��һ��������͡�", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then

                If Val(vsf.RowData(1)) > 0 Then
                    If MsgBox("�Ƿ�Ҫ�����ѡ��������Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call ResetVsf(vsf)
                        Call ResetVsf(vsfPrice)
                    End If
                End If

                rs.MoveFirst

                Do While Not rs.EOF

                    Call ReadTemplate(rs("ID").Value)
                    rs.MoveNext

                Loop

                DataChange = True
            End If

        End If

        EnterFocus vsf
    '------------------------------------------------------------------------------------------------------------------
    Case 8
    
        Dim strParam As String
        Dim varParam As Variant

        strParam = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) & "'"
        strParam = strParam & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������), "yyyy-MM-dd") & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) & "'"

        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������)
        
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")

            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(1)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = varParam(2)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = varParam(3)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = varParam(4)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = varParam(5)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = Val(varParam(0))

            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(6)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(7)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) = varParam(8)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) = varParam(9)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) = varParam(10)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) = varParam(11)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) = varParam(12)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) = varParam(14)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) = varParam(15)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(16)
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������) = varParam(17)
            
'            imgNew(1).Visible = False

        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 9
    
        On Error GoTo 0

        dlg.CancelError = True

        On Error GoTo ErrHandler

        dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
        dlg.Filter = "�������(*.xls)| *.xls"
        dlg.FilterIndex = 0

        dlg.DialogTitle = "��������ռ�"
        dlg.FileName = App.Path & "\��������ռ�.xls"
        dlg.ShowOpen

        If Dir(dlg.FileName) <> "" Then
            If ReadExcelFile(dlg.FileName) Then
                If Not (lvwGroup.SelectedItem Is Nothing) Then
                    mstrGroup = ""
                    Call lvwGroup_ItemClick(lvwGroup.SelectedItem)
                End If
                'Call lvwGroup_AfterDeleteRow(lvwGroup.Row, lvwGroup.Col)
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 10
                
        Call SaveGroupInfo
        
        If frmGroupEdit.ShowEdit(Me, mrsGroup) Then
            Call ShowGroupInfo
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 11
            
        Call WritePersons("ȱʡ", 2)
        
        strParam = ""
        strParam = strParam & mrsPersons("����id").Value & "'"
        strParam = strParam & mrsPersons("����").Value & "'"
        strParam = strParam & mrsPersons("���֤").Value & "'"
        strParam = strParam & mrsPersons("�Ա�").Value & "'"
        strParam = strParam & mrsPersons("��������").Value & "'"
        strParam = strParam & mrsPersons("����״��").Value & "'"
        strParam = strParam & mrsPersons("����").Value & "'"
        strParam = strParam & mrsPersons("����").Value & "'"
        strParam = strParam & mrsPersons("ѧ��").Value & "'"
        strParam = strParam & mrsPersons("ְҵ").Value & "'"
        strParam = strParam & mrsPersons("���").Value & "'"
        strParam = strParam & mrsPersons("��ϵ������").Value & "'"
        strParam = strParam & mrsPersons("��ϵ�˵绰").Value & "'"
        strParam = strParam & mrsPersons("�����ʼ�").Value & "'"
        strParam = strParam & mrsPersons("��ϵ�˵�ַ").Value & "'"
        strParam = strParam & mrsPersons("������λ").Value & "'"
        strParam = strParam & mrsPersons("����").Value & "'"
        strParam = strParam & mrsPersons("������").Value
        
        
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")
            
            mrsPersons("����").Value = varParam(1)
            mrsPersons("���֤").Value = varParam(2)
            mrsPersons("�Ա�").Value = varParam(3)
            mrsPersons("��������").Value = varParam(4)
            mrsPersons("����״��").Value = varParam(5)
            mrsPersons("����").Value = varParam(6)
            mrsPersons("����").Value = varParam(7)
            mrsPersons("ѧ��").Value = varParam(8)
            mrsPersons("ְҵ").Value = varParam(9)
            mrsPersons("���").Value = varParam(10)
            mrsPersons("��ϵ������").Value = varParam(11)
            mrsPersons("��ϵ�˵绰").Value = varParam(12)
            mrsPersons("�����ʼ�").Value = varParam(13)
            mrsPersons("��ϵ�˵�ַ").Value = varParam(14)
            mrsPersons("������λ").Value = varParam(15)
            mrsPersons("����").Value = varParam(16)
            mrsPersons("������").Value = varParam(17)
                                    
            Call ReadPersons("ȱʡ", 2)
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 12, 14    'д��
        
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            ReDim strInfo(1 To 16)
            
            strCardNo1 = clsCard.GetCardNo
            
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����)
            
            If strCardNo2 <> "" Then
                '�����п������͵�ǰ�Ŀ�����ͬһ�ſ�
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "�˿����ǵ�ǰ���˵Ŀ���"
                    Exit Sub
                End If
            Else
                '����û�п�
                
                If strCardNo1 = "" Then
                
                    '�¿����Զ�����
                    strCardNo1 = "11111111"
                    strCardNo2 = strCardNo1
                    
                    'д����
                    If clsCard.SetCardNo(strCardNo1) = False Then Exit Sub
                    vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����) = strCardNo2
                    
                Else
                
                    '�����¿�
                    ShowSimpleMsg "�˿������¿������ܽ���д�������"
                    Exit Sub
                    
                End If
                                                
            End If
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����) = strCardNo2
            
            If mblnGroup Then
                strInfo(1) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
                strInfo(2) = "���֤��=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤)
                strInfo(3) = "�Ա�=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�)
                strInfo(4) = "��������=" & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������), "yyyy-MM-dd")
                strInfo(5) = "����״��=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��)
                strInfo(6) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
                strInfo(7) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
                strInfo(8) = "ѧ��=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��)
                strInfo(9) = "ְҵ=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ)
                strInfo(10) = "���=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���)
                strInfo(11) = "��ϵ������=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������)
                strInfo(12) = "��ϵ�˵绰=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰)
                strInfo(13) = "��ϵ�˵�ַ=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ)
                strInfo(14) = "�����ʼ�=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�)
                strInfo(15) = "������λ=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ)
                strInfo(16) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
            Else
                strInfo(1) = "����=" & mrsPersons("����").Value
                strInfo(2) = "���֤��=" & mrsPersons("���֤").Value
                strInfo(3) = "�Ա�=" & mrsPersons("�Ա�").Value
                strInfo(4) = "��������=" & mrsPersons("��������").Value
                strInfo(5) = "����״��=" & mrsPersons("����״��").Value
                strInfo(6) = "����=" & mrsPersons("����").Value
                strInfo(7) = "����=" & mrsPersons("����").Value
                strInfo(8) = "ѧ��=" & mrsPersons("ѧ��").Value
                strInfo(9) = "ְҵ=" & mrsPersons("ְҵ").Value
                strInfo(10) = "���=" & mrsPersons("���").Value
                strInfo(11) = "��ϵ������=" & mrsPersons("��ϵ������").Value
                strInfo(12) = "��ϵ�˵绰=" & mrsPersons("��ϵ�˵绰").Value
                strInfo(13) = "��ϵ�˵�ַ=" & mrsPersons("��ϵ�˵�ַ").Value
                strInfo(14) = "�����ʼ�=" & mrsPersons("�����ʼ�").Value
                strInfo(15) = "������λ=" & mrsPersons("������λ").Value
                strInfo(16) = "����=" & mrsPersons("����").Value
            End If
                        
            If clsCard.SetPatient(strInfo) Then
                ShowSimpleMsg "���µ�ǰ������Ϣ�ɹ���"
            End If
        End If
        
        If mblnGroup Then
            If vsfPerson.Visible Then vsfPerson.SetFocus
        Else
            If txt(5).Visible Then txt(5).SetFocus
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 13, 15    '����
        
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            If mblnGroup = False Then Call WritePersons("ȱʡ", 2)
            
            strCardNo1 = clsCard.GetCardNo
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����)
            
            If strCardNo2 <> "" Then
                '��¼�Ĳ����п������͵�ǰ�Ŀ�����ͬһ�ſ�
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "�˿����ǵ�ǰ���˵Ŀ���"
                    Exit Sub
                End If
            Else
            
                '����û�п����򽫵�ǰ�Ŀ��Ÿ�������
                strCardNo2 = strCardNo1
                                
            End If
            
            If mblnGroup = False Then
                mrsPersons("IC����").Value = strCardNo2
            Else
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����) = strCardNo2
            End If
            
            
            If GetPatientID(strCardNo2) > 0 Then
                
                '��ϵͳ���ҵ��˲���
                If mblnGroup Then
                    vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = GetPatientID(strCardNo2)
                    Call GetPatientInfo(Val(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id)))
                Else
                    mrsPersons("IC����").Value = GetPatientID(strCardNo2)
                    Call GetPatientInfo(Val(mrsPersons("IC����").Value))
                End If
                
            ElseIf clsCard.GetPatient(strInfo) Then
                    
                For lngLoop = LBound(strInfo) To UBound(strInfo)
                    If InStr(strInfo(lngLoop), "=") > 0 Then
                        strItem = Mid(strInfo(lngLoop), 1, InStr(strInfo(lngLoop), "=") - 1)
                        strValue = Mid(strInfo(lngLoop), InStr(strInfo(lngLoop), "=") + 1)
                        
                        Select Case strItem
                        Case "����"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            Else
                                mrsPersons("����").Value = strValue
                            End If
                            
                        Case "���֤��"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = strValue
                            Else
                                mrsPersons("���֤").Value = strValue
                            End If
                            
                        Case "�Ա�"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = strValue
                            Else
                                mrsPersons("�Ա�").Value = strValue
                            End If
                            
                        Case "��������"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = strValue
                            Else
                                mrsPersons("��������").Value = strValue
                            End If
                            
                        Case "����״��"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = strValue
                            Else
                                mrsPersons("����״��").Value = strValue
                            End If
                            
                        Case "����"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            Else
                                mrsPersons("����").Value = strValue
                            End If
                            
                        Case "����"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            Else
                                mrsPersons("����").Value = strValue
                            End If
                            
                        Case "ѧ��"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) = strValue
                            Else
                                mrsPersons("ѧ��").Value = strValue
                            End If
                            
                        Case "ְҵ"
                        
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) = strValue
                            Else
                                mrsPersons("ְҵ").Value = strValue
                            End If
                            
                        Case "���"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) = strValue
                            Else
                                mrsPersons("���").Value = strValue
                            End If
                            
                        Case "��ϵ������"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) = strValue
                            Else
                                mrsPersons("��ϵ������").Value = strValue
                            End If
                            
                        Case "��ϵ�˵绰"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) = strValue
                            Else
                                mrsPersons("��ϵ�˵绰").Value = strValue
                            End If
                            
                        Case "��ϵ�˵�ַ"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) = strValue
                            Else
                                mrsPersons("��ϵ�˵�ַ").Value = strValue
                            End If
                            
                        Case "�����ʼ�"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) = strValue
                            Else
                                mrsPersons("�����ʼ�").Value = strValue
                            End If
                            
                        Case "������λ"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) = strValue
                            Else
                                mrsPersons("������λ").Value = strValue
                            End If
                            
                        Case "����"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            Else
                                mrsPersons("����").Value = strValue
                            End If
                            
                        End Select
                    End If
                Next
                
                If mblnGroup = False Then Call ReadPersons("ȱʡ", 2)
            End If
        End If
        
        If mblnGroup Then
            If vsfPerson.Visible Then vsfPerson.SetFocus
        Else
            If txt(5).Visible Then txt(5).SetFocus
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 16
        'ѡ��λ��Ա

        If frmSelectGroupPerson.ShowFilter(Me, Val(cmd(4).Tag), rs) Then
        
            rs.Filter = 0
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then

                If Val(vsfPerson.RowData(1)) > 0 Then
                    If MsgBox("�Ƿ�Ҫ�����ѡ����ܼ���Ա��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call ResetVsf(vsfPerson)
                    End If
                End If

                rs.MoveFirst

                Do While Not rs.EOF

                    If CheckHavePerson(rs("ID").Value) = False Then
                        With vsfPerson

                            .Row = .Rows - 1
                            If Val(.RowData(.Row)) > 0 Then
                                .Rows = .Rows + 1
                                .Row = .Rows - 1
                            End If

                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����").Value)
                            .TextMatrix(.Row, mPersonCol.������) = zlCommFun.NVL(rs("������").Value)
                            .TextMatrix(.Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                            .TextMatrix(.Row, mPersonCol.��������) = zlCommFun.NVL(rs("��������").Value)
                            .TextMatrix(.Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��").Value)
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.ѧ��) = zlCommFun.NVL(rs("ѧ��").Value)
                            .TextMatrix(.Row, mPersonCol.ְҵ) = zlCommFun.NVL(rs("ְҵ").Value)
                            .TextMatrix(.Row, mPersonCol.���) = zlCommFun.NVL(rs("���").Value)
                            .TextMatrix(.Row, mPersonCol.��ϵ������) = zlCommFun.NVL(rs("��ϵ������").Value)
                            .TextMatrix(.Row, mPersonCol.��ϵ�˵绰) = zlCommFun.NVL(rs("��ϵ�˵绰").Value)
'                            .TextMatrix(.Row, mPersonCol.�����ʼ�) = zlCommFun.NVL(rs("�����ʼ�").Value)
                            .TextMatrix(.Row, mPersonCol.��ϵ�˵�ַ) = zlCommFun.NVL(rs("��ϵ�˵�ַ").Value)
                            .TextMatrix(.Row, mPersonCol.������λ) = zlCommFun.NVL(rs("������λ").Value)
                            .TextMatrix(.Row, mPersonCol.����id) = zlCommFun.NVL(rs("ID").Value, 0)
                            .TextMatrix(.Row, mPersonCol.IC����) = zlCommFun.NVL(rs("IC����").Value)
                            .TextMatrix(.Row, mPersonCol.���￨��) = zlCommFun.NVL(rs("���￨��").Value)

                            .RowData(.Row) = zlCommFun.NVL(rs("ID").Value)
                        End With
                    End If

                    rs.MoveNext

                Loop
                Call CountGroup
                DataChange = True
            End If
        End If
        
        Call EnterFocus(vsfPerson)
        
    '------------------------------------------------------------------------------------------------------------------
    Case 17
        
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, mCol.���㷽ʽ) = "����"
            End If
        Next
        
    '------------------------------------------------------------------------------------------------------------------
    Case 18
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, mCol.���㷽ʽ) = "�շ�"
            End If
        Next
        
ErrHandler:


    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    
    Dim lngKey As Long
    
    If Trim(lvwGroup.SelectedItem.Text) <> "" Then
        
        Call WritePrice(vsf.Row)
        Call SaveItems(Trim(lvwGroup.SelectedItem.Text))
        
        If mblnGroup = False Then
            Call WritePersons("ȱʡ", 2)
        Else
            Call WritePersons(Trim(lvwGroup.SelectedItem.Text))
        End If
    End If
    
    Call SaveGroupInfo
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit(lngKey) Then
        
        mblnOK = True
                
        If mlngKey = 0 And mbytMode = 1 Then
        
            Call ClearData
            
            lvwGroup.ListItems.Add , , "ȱʡ"
            
            DataChange = False
            
            ShowSimpleMsg "ԤԼ�Ǽǳɹ���������һ��ԤԼ��"
            
            If mblnGroup Then
                Call LocationObj(txt(13))
            Else
                Call LocationObj(txt(5))
            End If
        
        Else
            DataChange = False
            Unload Me
        End If
    End If
End Sub


Private Sub dtp_Change(Index As Integer)
    DataChange = True
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        
    ElseIf Shift = 0 Then
        Select Case KeyCode
        Case vbKeyF3
        
            If cmd(5).Enabled And tbs.Tab = 0 Then
                Call cmd_Click(5)
            End If
            
        Case vbKeyF4
        
            If cmd(6).Enabled And tbs.Tab = 0 Then
                Call cmd_Click(6)
            End If
                
                
        Case vbKeyF5
        
            If cmd(7).Enabled And tbs.Tab = 0 Then
                Call cmd_Click(7)
            End If
            
        Case vbKeyF6
        
            If cmd(8).Enabled And tbs.Tab = 1 Then
                Call cmd_Click(8)
            End If
            
        Case vbKeyF7
        
            If cmd(9).Enabled And tbs.Tab = 1 Then
                Call cmd_Click(9)
            End If
            
        Case vbKeyF8
        
            If cmd(0).Enabled And mblnGroup Then
                Call cmd_Click(0)
            End If
            
        Case vbKeyF9
    
            If cmd(2).Enabled And mblnGroup Then
                Call cmd_Click(2)
            End If
            
        Case vbKeyF10
        
            If cmd(3).Enabled And mblnGroup Then
                Call cmd_Click(3)
            End If
            
        Case vbKeyF11
        
            If cmd(11).Enabled And mblnGroup = False Then
                Call cmd_Click(11)
            End If
            
        Case vbKeyF12
        
            If cmd(10).Enabled And mblnGroup Then
                Call cmd_Click(10)
            End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    glngFormW = 12000
    glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    Dim lngY As Long
    
    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With fraGroupInfo
        .Left = fraTitle.Left
        .Top = fraTitle.Top + fraTitle.Height - 90
        .Width = fraTitle.Width
    End With
    
    If fraGroupInfo.Visible Then lngY = fraGroupInfo.Top + fraGroupInfo.Height
    
    With fraSingle
        .Left = fraGroupInfo.Left
        .Top = fraGroupInfo.Top
        .Width = fraGroupInfo.Width
    End With
    
    If fraSingle.Visible Then lngY = fraSingle.Top + fraSingle.Height
    
    With fraInfo
        .Left = fraGroupInfo.Left
        .Top = fraGroupInfo.Top + IIf(mblnGroup, fraGroupInfo.Height, fraSingle.Height) - 90
        .Width = fraGroupInfo.Width
    End With
    
    If fraInfo.Visible Then lngY = fraInfo.Top + fraInfo.Height
    
    With fraGroup
        .Left = fraInfo.Left
        .Top = lngY + 60
        .Height = Me.ScaleHeight - .Top - fraOther.Height - picButton.Height + 90 - stbThis.Height
    End With
    
    With tbs
        .Left = IIf(fraGroup.Visible, fraGroup.Left + fraGroup.Width + 45, 0)
        .Top = fraGroup.Top
        .Width = Me.ScaleWidth - .Left - 45
        .Height = fraGroup.Height
    End With
    
    With fraOther
        .Left = fraInfo.Left
        .Top = tbs.Top + tbs.Height - 90
        .Width = fraInfo.Width
    End With
                
    With cmd(11)
        .Left = fraSingle.Width - .Width - 60
    End With
    
    With cmd(14)
        .Left = cmd(11).Left - .Width - 45
    End With
    
    With cmd(15)
        .Left = cmd(14).Left - .Width - 45
    End With
    
    With cmd(10)
        .Left = fraGroupInfo.Width - .Width - 60
    End With
                
    With picButton
        .Left = fraOther.Left
        .Top = fraOther.Top + fraOther.Height
        .Width = fraOther.Width
    End With
    
    With picNo
        .Left = fraTitle.Width - .Width - 45
    End With
    
    With txt(2)
        .Width = fraInfo.Width - .Left - 45
    End With
    
    With lvwGroup
        .Left = 75
        .Top = 225
        .Width = fraGroup.Width - .Left - 75
        .Height = fraGroup.Height - .Top - 60 - cmd(0).Height - 60
    End With
    
    cmd(0).Top = lvwGroup.Top + lvwGroup.Height + 60
    cmd(2).Top = cmd(0).Top
    cmd(3).Top = cmd(0).Top
    
    If mblnGroup Then
        tbs.Tab = 1
        With vsfPerson
            .Left = 90
            .Top = 450
            .Width = tbs.Width - .Left - 90
            .Height = tbs.Height - .Top - 90
        End With
        With cmd(9)
            .Left = tbs.Width - .Width
            .Top = 0
        End With

        With cmd(16)
            .Left = cmd(9).Left - .Width - 45
            .Top = cmd(9).Top
        End With
        
        With cmd(8)
            .Left = cmd(16).Left - .Width - 45
            .Top = cmd(16).Top
        End With
        
        With cmd(12)
            .Left = cmd(8).Left - .Width - 45
            .Top = cmd(8).Top
        End With
        
        
        With cmd(13)
            .Left = cmd(12).Left - .Width - 45
            .Top = cmd(12).Top
        End With
        
    End If
    
    tbs.Tab = 0
    With vsf
        .Left = 90
        .Top = 450 + 300
        .Width = tbs.Width - .Left - 90
        .Height = tbs.Height - .Top - 90 - vsfPrice.Height - 30
    End With
    
    With vsfPrice
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height + 30
        .Width = vsf.Width
    End With
    
    With cmd(6)
        .Left = tbs.Width - .Width
        .Top = 0
    End With

    With cmd(5)
        .Left = cmd(6).Left - .Width - 45
        .Top = cmd(6).Top
    End With
    
    With cmd(17)
        .Left = cmd(5).Left - .Width - 45
        .Top = cmd(6).Top
    End With
    
    With cmd(18)
        .Left = cmd(17).Left - .Width - 45
        .Top = cmd(6).Top
    End With
    
    txt(6).Width = fraOther.Width - txt(6).Left - 45
        
    cmdCancel.Left = picButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If DataChange Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub lvwGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If mstrGroup <> Item.Text Then

        On Error Resume Next
        
        Call WritePrice(vsf.Row)
        Call SaveItems(mstrGroup)
        Call WritePersons(mstrGroup)
        
        Call ResetVsf(vsf)
        Call ResetVsf(vsfPerson)
        Call ResetVsf(vsfPrice)
        
        mstrGroup = Item.Text
        
        Call ReadItems(mstrGroup)
        Call ReadPrice(vsf.Row)
        
        Call ReadPersons(mstrGroup)
        
        Call CountGroup
    End If
    
End Sub

Private Function ReadExcelFile(ByVal strExcelFile As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim objExcel As Object
    Dim ExWorkbook As Object
    Dim ExWorkSheet As Object
    Dim lngLoop As Long
    Dim lngLoop2 As Long
    Dim rsTmp As New ADODB.Recordset
    Dim str���� As String
    Dim str���֤�� As String
    Dim lng����id As Long
    Dim str��� As String
    Dim str����� As String
    
    On Error GoTo errHand
    
    frmWait.OpenWait Me, "���������Ա"
    frmWait.WaitInfo = "���ڴ�""" & strExcelFile & """..."
    
    Set objExcel = CreateObject("Excel.Application")
    Set ExWorkbook = Nothing
    Set ExWorkSheet = Nothing
    
    Set ExWorkbook = objExcel.Workbooks.Open(strExcelFile)
    If ExWorkbook Is Nothing Then Exit Function
    
    Set ExWorkSheet = ExWorkbook.Worksheets("��Ա����")
    If ExWorkSheet Is Nothing Then Exit Function
    
    '��ɾ��
    mrsPersons.Filter = ""
    
    Call CopyRecord(mrsPersons, rsTmp)
    Call DeleteRecord(rsTmp)
        
    For lngLoop = 4 To ExWorkSheet.UsedRange.Cells.Rows.Count
        
        str���� = Trim(ExWorkSheet.Range(Chr(mColChar.����) & lngLoop).Value)
        
        If Trim(str����) = "" Then
            frmWait.WaitInfo = "����������Ա����..."
        Else
            frmWait.WaitInfo = "���ڵ���""" & str���� & """����..."
        
            str���֤�� = Trim(ExWorkSheet.Range(Chr(mColChar.���֤��) & lngLoop).Value)
            str��� = ExWorkSheet.Range(Chr(mColChar.�����) & lngLoop).Value
            
            '�����������֤������Ա����
            str����� = ""
            lng����id = 0
            Call SearchArchive(str���֤��, lng����id, str�����)
            If Val(str�����) = 0 Then
                str����� = Val(ExWorkSheet.Range(Chr(mColChar.�����) & lngLoop).Value)
            End If
            rsTmp.AddNew
                                        
            '�����
            If str��� <> "" Then
                For lngLoop2 = 1 To lvwGroup.ListItems.Count
                    If str��� = lvwGroup.ListItems(lngLoop2).Text Then
                        Exit For
                    End If
                Next
                
                If lngLoop2 = lvwGroup.ListItems.Count + 1 Then
                    
                    '�������
                    'lvwGroup.Rows = lvwGroup.Rows + 1
                    'lvwGroup.TextMatrix(lvwGroup.Rows - 1, 1) = str���
                    
                    lvwGroup.ListItems.Add , , str���, 1, 1
                End If
            Else
                str��� = lvwGroup.ListItems(1).Text
            End If
            
            rsTmp("���").Value = FitlerImport(str���, 30)
            rsTmp("����id").Value = lng����id
            rsTmp("����").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.����) & lngLoop).Value, 20)
            rsTmp("���֤").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.���֤��) & lngLoop).Value, 18)
            rsTmp("�Ա�").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.�Ա�) & lngLoop).Value, 4)
            rsTmp("��������").Value = FitlerImport(Format(ExWorkSheet.Range(Chr(mColChar.��������) & lngLoop).Value, "yyyy-MM-dd"), , "����")
            rsTmp("����״��").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.����״��) & lngLoop).Value, 4)
            rsTmp("����").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.����) & lngLoop).Value, 20)
            rsTmp("����").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.����) & lngLoop).Value, 30)
            rsTmp("ѧ��").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.ѧ��) & lngLoop).Value, 10)
            rsTmp("ְҵ").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.ְҵ) & lngLoop).Value, 20)
            rsTmp("�����ʼ�").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.�����ʼ�) & lngLoop).Value, 50)
            rsTmp("������λ").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.������λ) & lngLoop).Value, 100)
            rsTmp("����").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.����) & lngLoop).Value, 10)
            rsTmp("������").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.������) & lngLoop).Value, 50)
            rsTmp("���￨��").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.���￨��) & lngLoop).Value, 10)
                        
            rsTmp("�����").Value = Val(str�����)
'            rsTmp("���ʱ��").Value = Format(DateAdd("d", 7, CDate(zlDatabase.Currentdate)), "yyyy-MM-dd")
            rsTmp("ɾ��").Value = ""
            
        End If
    Next
    
    Call DeleteRecord(mrsPersons)
    Call CopyRecord(rsTmp, mrsPersons)
    
    objExcel.Quit
    ReadExcelFile = True
    
    frmWait.CloseWait
    
    Exit Function
    
errHand:
    objExcel.Quit
    frmWait.CloseWait
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function FitlerImport(ByVal strText As String, Optional ByVal intLen As Integer = 0, Optional ByVal strMode As String = "�ַ�") As String
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    Select Case strMode
    Case "�ַ�"
        
        If InStr(strText, "'") > 0 Then strText = ReplaceAll(strText, "'", "")
        
        If intLen > 0 Then
        
            If LenB(StrConv(strText, vbFromUnicode)) > intLen Then
            
                'ȡֵ
                strTmp = StrConv(MidB(StrConv(strText, vbFromUnicode), 1, intLen), vbUnicode)
                
                Clipboard.Clear
                Clipboard.SetText strTmp
                strText = Trim(Clipboard.GetText)
                                
            End If
        End If
        
    Case "����"
        If CheckStrValid(strText, CHECKFORMAT.����) = False Then strText = ""
    End Select
     
    FitlerImport = strText
    
End Function

Private Function SearchArchive(ByVal str���֤�� As String, ByRef lng����id As Long, ByRef str����� As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If str���֤�� <> "" Then
        strSQL = "SELECT ����id,����� FROM ������Ϣ WHERE ���֤��=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���֤��)
        If rs.BOF = False Then
            lng����id = rs("����id").Value
            str����� = CStr(zlCommFun.NVL(rs("�����").Value, 0))
        End If
        
    End If
    
    SearchArchive = True
    
End Function

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
    
    If PreviousTab = 0 Then
        vsf.Visible = False
        vsfPerson.Visible = True
        cmd(5).Visible = False
        cmd(6).Visible = False
        cmd(8).Visible = True
        cmd(9).Visible = True
        cmd(12).Visible = True
        cmd(13).Visible = True
        cmd(16).Visible = True
    Else
        vsf.Visible = True
        vsfPerson.Visible = False
        cmd(5).Visible = True
        cmd(6).Visible = True
        cmd(8).Visible = False
        cmd(9).Visible = False
        cmd(12).Visible = False
        cmd(13).Visible = False
        cmd(16).Visible = False
    End If
End Sub

Private Sub txt_Change(Index As Integer)

    DataChange = True
    
    If Index = 13 Or Index = 5 Then
        txt(Index).Tag = "Changed"
                
        cmd(4).Tag = ""
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 0, 2, 5, 6, 12, 13
        zlCommFun.OpenIme True
    End Select
End Sub

Private Function FillPatient(ByVal lngKey As Long, Optional ByVal bytMode As Byte = 1)
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    strSQL = GetPublicSQL(SQL.��Ա����)
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        If bytMode = 1 Then
            If mrsPersons.RecordCount = 0 Then mrsPersons.AddNew
            mrsPersons("���").Value = "ȱʡ"
            mrsPersons("����").Value = zlCommFun.NVL(rs("����").Value)
            mrsPersons("���֤").Value = zlCommFun.NVL(rs("���֤").Value)
            mrsPersons("�Ա�").Value = zlCommFun.NVL(rs("�Ա�").Value)
            mrsPersons("��������").Value = zlCommFun.NVL(rs("��������").Value)
            mrsPersons("����״��").Value = zlCommFun.NVL(rs("����״��").Value)
            mrsPersons("����id").Value = zlCommFun.NVL(rs("����id").Value)
            mrsPersons("����").Value = zlCommFun.NVL(rs("����").Value)
            mrsPersons("����").Value = zlCommFun.NVL(rs("����").Value)
            mrsPersons("ѧ��").Value = zlCommFun.NVL(rs("ѧ��").Value)
            mrsPersons("ְҵ").Value = zlCommFun.NVL(rs("ְҵ").Value)
            mrsPersons("���").Value = zlCommFun.NVL(rs("���").Value)
            mrsPersons("��ϵ������").Value = zlCommFun.NVL(rs("��ϵ������").Value)
            mrsPersons("��ϵ�˵绰").Value = zlCommFun.NVL(rs("��ϵ�˵绰").Value)
            'mrsPersons("�����ʼ�").Value = zlCommFun.NVL(rs("�����ʼ�").Value)
            mrsPersons("��ϵ�˵�ַ").Value = zlCommFun.NVL(rs("��ϵ�˵�ַ").Value)
            mrsPersons("������λ").Value = zlCommFun.NVL(rs("������λ").Value)
        Else
        
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) = zlCommFun.NVL(rs("ѧ��").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) = zlCommFun.NVL(rs("ְҵ").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) = zlCommFun.NVL(rs("���").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) = zlCommFun.NVL(rs("��ϵ������").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) = zlCommFun.NVL(rs("��ϵ�˵绰").Value)
            'vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) = zlCommFun.NVL(rs("��ϵ�˵�ַ").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) = zlCommFun.NVL(rs("������λ").Value)
            
        End If
    End If
    
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strInput As String
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim blnCard As Boolean
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If Index = 5 Then
        '���￨��

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard Then
            If Len(txt(Index).Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txt(Index).Text <> "" Then
                If KeyAscii <> 13 Then
                    txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                    txt(Index).SelStart = Len(txt(Index).Text)
                    KeyAscii = 0
                End If
    
                strInput = strInput & " AND C.���￨��=[1] "
    
            End If
        End If
    End If
        
    If KeyAscii <> vbKeyReturn Then
    
        If Index = 4 Then If FilterKeyAscii(KeyAscii, 99, "0123456789X") = 0 Then KeyAscii = 0
        If Index = 14 Then If FilterKeyAscii(KeyAscii, 2) = 0 Then KeyAscii = 0
        If Index = 31 Then If FilterKeyAscii(KeyAscii, 1) = 0 Then KeyAscii = 0
        
        DataChange = True
        
    ElseIf txt(Index).Tag = "Changed" And Index = 5 And KeyAscii = 13 Then
        If InStr(txt(Index).Text, "'") Then
            ShowSimpleMsg "�ڸ����������зǷ��ַ� ' ��"
            Exit Sub
        End If
        
        imgNew(1).Visible = False
        
        Select Case UCase(Left(txt(Index).Text, 1))
        Case "-", "A"                 '����id
            strInput = strInput & " AND C.����id=[1]"
        
        Case "+", "B"                 'סԺ��
            strInput = " AND C.סԺ��=[1]"

        Case "*", "D"                 '�����
            strInput = strInput & " AND C.�����=[1]"
            
        Case "/", "C"                 '��ǰ����
            strInput = strInput & " AND C.��ǰ����=[1]"
        Case Else
        
            cmd(1).Tag = ""
            imgNew(1).Visible = True
            txt(3).Text = ""
            
            txt(Index).Tag = ""
        
            'ԤԼ��ȱʡΪ���˱���
            txt(0).Text = txt(5).Text
            
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
            
            Exit Sub
            
        End Select
    ElseIf txt(Index).Tag = "Changed" And Index = 13 And KeyAscii = 13 Then
        If InStr(txt(Index).Text, "'") Then
            ShowSimpleMsg "�������������зǷ��ַ� ' ��"
            Exit Sub
        End If
        
        Dim lngKey As Long
        
        lngKey = Val(cmd(4).Tag)
        
        gstrSQL = GetPublicSQL(SQL.�������ѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
        
        If ShowTxtFilter(Me, txt(Index), "����,1800,0,0;����,900,0,0;����,900,0,0;��ϵ��,900,0,0;�绰,1200,0,0", Me.Name & "\�������ѡ��", "�������ѡ��һ�����嵥λ", rsData, rs, , , , False) Then
            
            Call ReadGroup(zlCommFun.NVL(rs("ID").Value, 0))
            
            If lngKey <> Val(cmd(4).Tag) Then DataChange = True
            
'            cmd(Index).Tag = lngKey
            
            imgNew(0).Visible = False
            
            txt(0).Text = txt(12).Text
            txt(1).Text = txt(11).Text
            
        Else
            cmd(4).Tag = ""
            imgNew(0).Visible = True
            
            mrsGroup("����").Value = ""
        End If
        
        txt(Index).Tag = ""
        
        'ԤԼ��ȱʡΪ������ϵ��
        txt(0).Text = txt(12).Text
        txt(1).Text = txt(11).Text
        
        zlCommFun.PressKey vbKeyTab
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    ElseIf KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    
    If strInput <> "" Then
    
        strText = txt(Index).Text
        
        KeyAscii = 0
        
        gstrSQL = GetPublicSQL(SQL.��Ա����ѡ��, strInput)
        
        If blnCard Then
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(txt(Index).Text))
        Else
            Select Case UCase(Left(txt(Index).Text, 1))
            Case "/", "C"                 '��ǰ����
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Mid(txt(Index).Text, 2)))
            Case Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(txt(Index).Text, 2)))
            End Select
        End If

        If ShowTxtFilter(Me, txt(Index), "����,1200,0,0;�Ա�,810,0,0;��������,1200,0,0;����״��,900,0,0;���֤��,1500,0,0", Me.Name & "\��Ա����ѡ��", "�������ѡ��һ����Ա", rsData, rs, , , , False) Then
                                    
            txt(Index).Text = zlCommFun.NVL(rs("����"))
            txt(4).Text = zlCommFun.NVL(rs("���֤��"))
            txt(9).Text = zlCommFun.NVL(rs("����"))
            txt(3).Text = zlCommFun.NVL(rs("�����"))
            txt(14).Text = zlCommFun.NVL(rs("������"))
            
            zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("�Ա�").Value)
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("����״��").Value)
            
            cmd(1).Tag = zlCommFun.NVL(rs("ID"))
            
            Call FillPatient(Val(cmd(1).Tag))
            
            txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
            
        Else
            cmd(1).Tag = ""
            imgNew(1).Visible = True
            txt(3).Text = ""
            txt(4).Text = ""
        End If
        
        txt(Index).Tag = ""
        
        'ԤԼ��ȱʡΪ���˱���
        txt(0).Text = txt(5).Text

        zlCommFun.PressKey vbKeyTab
        zlCommFun.PressKey vbKeyTab
        
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 2, 5, 6, 12, 13
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
End Sub

Private Sub txtSum_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtSum(Index)
        
End Sub

Private Sub txtSum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call WritePrice(vsf.Row)
        
        If Index = 1 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(1).Text), 1)
        If Index = 2 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(2).Text), 2)
        
        Call ReadPrice(vsf.Row)
   
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtSum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        glngTXTProc = GetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtSum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtSum_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txtSum(Index).Text, txtSum(Index).MaxLength)
    
    If Index = 1 Then
        If InStr(txtSum(1).Text, ".") > 0 Then
            If Len(Mid(txtSum(1).Text, InStr(txtSum(1).Text, ".") + 1)) > 2 Then
                MsgBox "ֻ����������λС��λ����", vbExclamation, gstrSysName
                Cancel = True
            End If
        End If
    End If
    
End Sub


Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    If vsf.Rows = 2 And Val(vsf.RowData(1)) = 0 Then
        Call ResetVsf(vsfPrice)
    Else
        Call ReadPrice(vsf.Row)
    End If
    Call CountGroup
    
    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
    
    DataChange = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.ִ�п���
        
        vsf.TextMatrix(Row, mCol.ִ�п���id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.ִ�п���) = vsf.Cell(flexcpTextDisplay, Row, mCol.ִ�п���)
    
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.�ɼ���ʽ
        
        vsf.TextMatrix(Row, mCol.�ɼ���ʽid) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.�ɼ���ʽ) = vsf.Cell(flexcpTextDisplay, Row, mCol.�ɼ���ʽ)
    
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.�ɼ�����
    
        vsf.TextMatrix(Row, mCol.�ɼ�����id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.�ɼ�����) = vsf.Cell(flexcpTextDisplay, Row, mCol.�ɼ�����)
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.���۸�
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
        Call ReadPrice(Row)

    '------------------------------------------------------------------------------------------------------------------
    Case mCol.�ۿ�
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.�ۿ�)), 2)
        Call ReadPrice(Row)
        
    End Select
    DataChange = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = OldRow Then Exit Sub
    
    Call ReadPrice(NewRow)
    
    Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
End Sub

Private Function ReadPrice(ByVal intRow As Integer) As Boolean
    '��ȡ��Ӧ�ļƷ���ϸ
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    Call ResetVsf(vsfPrice)
    
    If intRow = 0 Then Exit Function
    
    If vsf.TextMatrix(intRow, mCol.�Ʒ���ϸ) <> "" Then
        
        varRow = Split(vsf.TextMatrix(intRow, mCol.�Ʒ���ϸ), ";")
        
        vsfPrice.Rows = UBound(varRow) + 2
        
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
'                For lngCol = 0 To UBound(varCol)
                    
                    If Val(varCol(6)) = 2 Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ���Ŀ) = "�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ���Ŀ) = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                    Else
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ���Ŀ) = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                    End If
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p����) = varCol(0)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p���㵥λ) = varCol(1)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p����) = varCol(2)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p��׼����) = varCol(3)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p��쵥��) = varCol(4)
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p��׼���) = Val(varCol(2)) * Val(varCol(3))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�����) = Val(varCol(2)) * Val(varCol(4))
                                        
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�շ���Ŀid) = varCol(5)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ�����) = varCol(6)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.pִ�п���) = varCol(7)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.pִ�п���id) = varCol(8)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p���) = varCol(9)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�ۿ�) = varCol(10)
                    
                    vsfPrice.RowData(lngRow + 1) = Val(varCol(5))
                    
'                Next
            End If
        Next
        
    End If
    
    ReadPrice = True
    
End Function

Private Function WritePrice(ByVal intRow As Integer) As Boolean
    Dim strTmp As String
    Dim lngRow As Long
    Dim varCol As Variant
    
    On Error GoTo errHand
    
    If intRow <= 0 Then Exit Function
    
    For lngRow = 1 To vsfPrice.Rows - 1
        If Val(vsfPrice.TextMatrix(lngRow, mCol.p�շ���Ŀid)) > 0 Then
            
            varCol = Split(String(11, ":"), ":")
            
            varCol(0) = vsfPrice.TextMatrix(lngRow, mCol.p����)
            varCol(1) = vsfPrice.TextMatrix(lngRow, mCol.p���㵥λ)
            varCol(2) = vsfPrice.TextMatrix(lngRow, mCol.p����)
            varCol(3) = vsfPrice.TextMatrix(lngRow, mCol.p��׼����)
            varCol(4) = vsfPrice.TextMatrix(lngRow, mCol.p��쵥��)
            varCol(5) = vsfPrice.TextMatrix(lngRow, mCol.p�շ���Ŀid)
            varCol(6) = vsfPrice.TextMatrix(lngRow, mCol.p�Ƽ�����)
            
            If Val(varCol(6)) <> 2 Then varCol(6) = 1
                        
            varCol(7) = vsfPrice.TextMatrix(lngRow, mCol.pִ�п���)
            varCol(8) = vsfPrice.TextMatrix(lngRow, mCol.pִ�п���id)
            varCol(9) = vsfPrice.TextMatrix(lngRow, mCol.p���)
            varCol(10) = vsfPrice.TextMatrix(lngRow, mCol.p�ۿ�)
            
            If strTmp = "" Then
                strTmp = Join(varCol, ":")
            Else
                strTmp = strTmp & ";" & Join(varCol, ":")
            End If
        End If
    Next
    
    vsf.TextMatrix(intRow, mCol.�Ʒ���ϸ) = strTmp
    
    WritePrice = True
    
errHand:
    
End Function


Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
    Cancel = (Val(vsf.TextMatrix(Row, mCol.ִ�п���id)) <= 0)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error GoTo errHand
    
    If NewRow = OldRow Then Exit Sub
    
    '���ñ༭״̬
    If Val(vsf.TextMatrix(NewRow, mCol.�¼�)) = 1 Then
        vsf.EditMode(mCol.��Ŀ) = 0
        vsf.EditMode(mCol.ִ�п���) = 0
        vsf.EditMode(mCol.��鲿λ) = 0
        vsf.EditMode(mCol.�ɼ���ʽ) = 0
        vsf.EditMode(mCol.�ɼ�����) = 0
        vsf.EditMode(mCol.����걾) = 0
        vsf.EditMode(mCol.���㷽ʽ) = 0
        
        vsf.ComboList(mCol.��Ŀ) = ""
        vsf.ComboList(mCol.ִ�п���) = ""
        vsf.ComboList(mCol.��鲿λ) = ""
        vsf.ComboList(mCol.�ɼ���ʽ) = ""
        vsf.ComboList(mCol.�ɼ�����) = ""
        vsf.ComboList(mCol.����걾) = ""
        vsf.ComboList(mCol.���㷽ʽ) = ""
    Else
        
        vsf.EditMode(mCol.��Ŀ) = 1
        vsf.EditMode(mCol.ִ�п���) = 1
        vsf.EditMode(mCol.���㷽ʽ) = 1
        
        vsf.ComboList(mCol.��Ŀ) = "..."
        vsf.ComboList(mCol.ִ�п���) = " "
        vsf.ComboList(mCol.���㷽ʽ) = "����|�շ�"
        
        If mblnGroup Then
            vsf.EditMode(mCol.���㷽ʽ) = 0
            vsf.ComboList(mCol.���㷽ʽ) = ""
        End If
        
        Select Case vsf.TextMatrix(NewRow, mCol.���)
            Case "���"
                vsf.EditMode(mCol.�ɼ���ʽ) = 0
                vsf.EditMode(mCol.����걾) = 0
                vsf.EditMode(mCol.��鲿λ) = 1
                vsf.EditMode(mCol.�ɼ�����) = 0
                
                vsf.ComboList(mCol.�ɼ�����) = ""
                vsf.ComboList(mCol.�ɼ���ʽ) = ""
                vsf.ComboList(mCol.����걾) = ""
                vsf.ComboList(mCol.��鲿λ) = "..."
            Case "����"
                vsf.EditMode(mCol.�ɼ���ʽ) = 1
                vsf.EditMode(mCol.����걾) = 1
                vsf.EditMode(mCol.��鲿λ) = 0
                vsf.EditMode(mCol.�ɼ�����) = 1
                
                vsf.ComboList(mCol.�ɼ�����) = " "
                vsf.ComboList(mCol.�ɼ���ʽ) = " "
                vsf.ComboList(mCol.����걾) = " "
                vsf.ComboList(mCol.��鲿λ) = ""
        End Select
    End If
    
    Call WritePrice(OldRow)
    
    If vsf.TextMatrix(NewRow, mCol.���) = "����" Then
        Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "�Ƽ���Ŀ", "����ִ�п���", "�ɼ���ʽ", "����걾")
        Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "�ɼ�����")
    ElseIf vsf.TextMatrix(NewRow, mCol.���) = "���" Then
        Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "�Ƽ���Ŀ", "����ִ�п���")
    End If
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim bytResult As Byte
    Dim rsPrice As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    Dim rsData As New ADODB.Recordset
    
    Select Case Col
        Case mCol.��Ŀ
            
            gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)
            Dim bytParam1 As Byte
            Dim bytParam2 As Byte
            
            bytParam1 = 1
            bytParam2 = 2
                    
            If mblnGroup = False Then
                Select Case zlCommFun.GetNeedName(cbo(1).Text)
                Case "��"
                    bytParam1 = 1
                    bytParam2 = 1
                Case "Ů"
                    bytParam1 = 2
                    bytParam2 = 2
                End Select
            End If
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, bytParam1, bytParam2)
            
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;�걾��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 4500) Then
                'ѡȡ��һ����Ŀ
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                
                vsf.Cell(flexcpText, Row, mCol.��Ŀ + 1, Row, vsf.Cols - 1) = ""
                
                vsf.EditText = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf.TextMatrix(Row, mCol.��Ŀ) = zlCommFun.NVL(rs("����").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                  
                If vsf.TextMatrix(Row, mCol.���) = "����" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "�ɼ���ʽ", "�ɼ�����", "����걾", "���㷽ʽ", "�Ƽ���Ŀ")
                    
                ElseIf vsf.TextMatrix(Row, mCol.���) = "���" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
                End If
                
                Call CreatePriceList(Row)
                Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
                
                Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
                
                Call CountGroup
                
                DataChange = True
                
            End If
                
        Case mCol.��鲿λ
            
            bytResult = ShowOpenList("", mCol.��鲿λ)
            If bytResult = 0 Then ShowSimpleMsg "û���ҵ���ƥ�����Ŀ��"
            If bytResult = 1 Then
                Call CreatePriceList(Row)
                DataChange = True
            End If
            
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim bytResult As Byte
    Dim rs As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." And Col = mCol.��Ŀ Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
                        
            bytResult = ShowOpenList(UCase(vsf.EditText), Col)
            
            If bytResult = 0 Then
                'û��ƥ�����Ŀ
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
                MsgBox "û���ҵ���ƥ��������Ŀ��", vbInformation, gstrSysName
            End If
            
            If bytResult = 1 Then
                'ѡȡ��һ����Ŀ
                DataChange = True
                
                If Col = mCol.��Ŀ Then
                    
                    If vsf.TextMatrix(Row, mCol.���) = "����" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "�ɼ���ʽ", "�ɼ�����", "����걾", "���㷽ʽ", "�Ƽ���Ŀ")
                        
                    ElseIf vsf.TextMatrix(Row, mCol.���) = "���" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
                    End If
                    
                    Call CreatePriceList(Row)
                    
                    Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                    Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
                    Call CountGroup
                    
                    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
                    
                End If
            End If
            
            If bytResult = 2 Then
                'ȡ���˱���ѡ��
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
            End If
            
        End If
    Else
        DataChange = True
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    If KeyAscii = vbKeyReturn Then
                
        If Col = 1 Then
            If Trim(vsf.TextMatrix(Row, Col)) = "" Then
                
                KeyAscii = 0
                
                If mblnGroup Then
                                            
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                Else
                    If dtp(0).Enabled Then
                        dtp(0).SetFocus
                    Else
                        chk.SetFocus
                    End If
                End If
                
                Cancel = True
                
            End If
        End If
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    Select Case Col
    Case mPersonCol.�����
        '���������Ƿ����
        If Trim(vsfPerson.EditText) <> "" Then
            gstrSQL = "Select 1 From ������Ϣ Where �����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(vsfPerson.EditText))
            If rs.BOF = False Then
                '����
                Cancel = True
                
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.EditText
                
                ShowSimpleMsg "��ǰ����ţ�" & Val(vsfPerson.EditText) & "�Ѿ����ڣ��������ظ���"
                vsfPerson.EditText = ""
                vsfPerson.TextMatrix(Row, Col) = ""
                
            End If
        End If
    End Select
End Sub

Private Sub vsfPerson_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call CountGroup
End Sub

Private Sub vsfPerson_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Select Case Col
    Case mPersonCol.��������
    
        If Trim(vsfPerson.TextMatrix(Row, Col)) <> "" Then
            vsfPerson.TextMatrix(Row, Col) = Format(zlCommFun.AddDate(vsfPerson.TextMatrix(Row, Col)), "yyyy-MM-dd")
            If CheckStrValid(vsfPerson.TextMatrix(Row, Col), CHECKFORMAT.����) = False Then vsfPerson.TextMatrix(Row, Col) = ""
        End If

    Case mPersonCol.�����ʼ�
    
        If CheckStrValid(vsfPerson.TextMatrix(Row, Col), CHECKFORMAT.�����ʼ�) = False Then vsfPerson.TextMatrix(Row, Col) = ""
        
    Case mPersonCol.���֤
    
        If CheckStrValid(vsfPerson.TextMatrix(Row, Col), CHECKFORMAT.���֤��) = False Then vsfPerson.TextMatrix(Row, Col) = ""
    
    Case mPersonCol.�Ա�
        
        Call CountGroup
        
    End Select
    
End Sub

Private Sub vsfPerson_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfPerson.Rows = 2 Then
        vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
    End If
End Sub

Private Sub vsfPerson_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error GoTo errHand
    
    If Val(vsfPerson.TextMatrix(NewRow, mPersonCol.����id)) = 0 Or Val(vsfPerson.TextMatrix(NewRow, mPersonCol.�����)) = 0 Then
        vsfPerson.EditMode(mPersonCol.�����) = 1
    Else
        vsfPerson.EditMode(mPersonCol.�����) = 0
    End If
errHand:
End Sub

Private Sub vsfPerson_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    
    If frmPatientFind.ShowFind(Me, lngKey) Then
        If lngKey > 0 Then
            
            gstrSQL = "SELECT A.* FROM ������Ϣ A WHERE A.����id=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                
                If Val(cmd(4).Tag) <> Val(zlCommFun.NVL(rs("��ͬ��λid"))) And Val(zlCommFun.NVL(rs("��ͬ��λid"))) > 0 And Val(cmd(4).Tag) > 0 Then
                    
                    If MsgBox("���ǵ�ǰ�������Ա���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    
                End If
                
                vsfPerson.EditText = zlCommFun.NVL(rs("����"))
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("����").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                
                Call SetRowDefault(0, Row, "ȱʡ��Ϣ")
                
                vsfPerson.TextMatrix(Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����"))
                vsfPerson.TextMatrix(Row, mPersonCol.������) = zlCommFun.NVL(rs("������"))
                vsfPerson.TextMatrix(Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��"))
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                vsfPerson.TextMatrix(Row, mPersonCol.��������) = Format(zlCommFun.NVL(rs("��������")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����id) = zlCommFun.NVL(rs("����id"))
                
                Call FillPatient(Val(zlCommFun.NVL(rs("����id"))), 2)
                
                vsfPerson.EditMode(mPersonCol.�����) = 0
                Call CountGroup
                DataChange = True
                                
            End If
            
        End If
    End If

End Sub

Private Sub vsfPerson_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strText As String
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strSvrText As String
    Dim rsData As New ADODB.Recordset
    Dim blnCard As Boolean
    
    If Chr(KeyCode) = "'" Then KeyCode = 0
    
    If Col = mPersonCol.���� Then
        
        strText = vsfPerson.EditText
        If KeyCode <> 8 And KeyCode <> 13 Then
            strText = strText & Chr(KeyCode)
        End If
        
        '���Ƿ��ַ�
        If InStr(strText, "'") > 0 Then
            KeyCode = 0
            ShowSimpleMsg "�ڸ����������зǷ��ַ� ' ��"
            vsfPerson.EditText = ""
            vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
                
        '����Ƿ�Ϊ���￨����
        blnCard = InputIsCard(vsfPerson.EditText, KeyCode)

        If blnCard And Len(vsfPerson.EditText) = ParamInfo.���￨���볤�� - 1 And KeyCode <> 8 And KeyCode <> vbKeyReturn Then
            vsfPerson.Body.EditSelStart = Len(vsfPerson.EditText)
            strInput = strInput & " AND C.���￨��=[1] "
        End If

        If KeyCode = vbKeyReturn Then

            If blnCard Then
                '�Ǿ��￨
                strInput = strInput & " AND C.���￨��=[1] "
            Else
                '�Ǿ��￨
                blnCard = False
                
                strText = vsfPerson.EditText
                
                Select Case UCase(Left(strText, 1))
                Case "-", "A"                 '����id,���￨��
                    strInput = strInput & " AND C.����id=[1]"
                Case "+", "B"                 'סԺ��
                    strInput = " AND C.סԺ��=[1]"
                Case "*", "D"                 '�����
                    strInput = strInput & " AND C.�����=[1]"
                Case "/", "C"                 '��ǰ����
                    strInput = strInput & " AND C.��ǰ����=[1]"
                Case Else                     '����
                    strSvrText = vsfPerson.Cell(flexcpData, Row, Col)
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                End Select
            End If
                
        End If
    
        
        If strInput <> "" Then
        
            gstrSQL = GetPublicSQL(SQL.��Ա����ѡ��, strInput)
                    
            If blnCard Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strText))
            ElseIf UCase(Left(strText, 1)) = "/" Or UCase(Left(strText, 1)) = "C" Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Mid(strText, 2)))
            Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(strText, 2)))
            End If
                    
            If ShowGrdFilter(Me, vsfPerson, "����,1200,0,0;�Ա�,810,0,0;��������,1200,0,0;����״��,900,0,0;���֤��,1500,0,0", Me.Name & "\��Ա����ѡ��Grid", "�������ѡ��һ����Ա", rsData, rs, , , , False) Then
                                                                        
                vsfPerson.EditText = zlCommFun.NVL(rs("����"))
                
                If Val(cmd(4).Tag) <> Val(zlCommFun.NVL(rs("��ͬ��λid"))) And Val(zlCommFun.NVL(rs("��ͬ��λid"))) > 0 And Val(cmd(4).Tag) > 0 Then
                    
                    If MsgBox("���ˡ�" & zlCommFun.NVL(rs("����").Value) & "�����ǵ�ǰ�������Ա���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        KeyCode = 0
                        vsfPerson.EditText = ""
                        vsfPerson.TextMatrix(Row, Col) = strSvrText
                        Cancel = True
                        Exit Sub
                    End If
                
                End If
                
                If CheckHavePerson(Val(zlCommFun.NVL(rs("ID")))) Then
                    ShowSimpleMsg "���ˡ�" & zlCommFun.NVL(rs("����").Value) & "���Ѿ����ڣ�"
                    KeyCode = 0
                    vsfPerson.EditText = ""
                    vsfPerson.TextMatrix(Row, Col) = strSvrText
                    Cancel = True
                    Exit Sub
                End If

                strText = vsfPerson.EditText
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("����").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                Call SetRowDefault(0, Row, "ȱʡ��Ϣ")
                vsfPerson.TextMatrix(Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��"))
                vsfPerson.TextMatrix(Row, mPersonCol.��������) = Format(zlCommFun.NVL(rs("��������")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����id) = zlCommFun.NVL(rs("ID"))
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                vsfPerson.TextMatrix(Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����"))
                vsfPerson.TextMatrix(Row, mPersonCol.������) = zlCommFun.NVL(rs("������"))
                
                Call FillPatient(Val(zlCommFun.NVL(rs("ID"))), 2)
                
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
                
                vsfPerson.EditMode(mPersonCol.�����) = 0
                Call CountGroup
                
                If blnCard Then
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                    vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                    KeyCode = 13
                End If
                
                DataChange = True
            Else
                'ȡ���˱���ѡ����Ϊ�²���
    
                vsfPerson.EditMode(mPersonCol.�����) = 1
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
                
                vsfPerson.Cell(flexcpData, Row, Col) = vsfPerson.EditText
                vsfPerson.EditText = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, mPersonCol.�����) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.���֤) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.����id) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.��������) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.����) = ""
                
                Call SetRowDefault(0, Row, "ȱʡ��Ϣ")
            End If
        ElseIf KeyCode = vbKeyReturn Then
    
            '�²��ˣ��������������
            
            vsfPerson.EditMode(mPersonCol.�����) = 1
            vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
            vsfPerson.TextMatrix(Row, mPersonCol.����id) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.�����) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.���֤) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.��������) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.����) = ""
            Call SetRowDefault(0, Row, "ȱʡ��Ϣ")
        End If
    End If
End Sub

Private Sub vsfPerson_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        
        If Col = 1 Then
            If Trim(vsfPerson.TextMatrix(Row, Col)) = "" Then
                KeyAscii = 0
                
                If dtp(0).Enabled Then
                    dtp(0).SetFocus
                Else
                    chk.SetFocus
                End If
                
                Cancel = True
                
            End If
        End If
    End If
    
End Sub

Private Sub vsfPerson_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    Select Case Col
    Case mPersonCol.�����
        '���������Ƿ����
        If Val(vsfPerson.EditText) > 0 Then
            gstrSQL = "Select 1 From ������Ϣ Where �����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsfPerson.EditText))
            If rs.BOF = False Then
                '����
                Cancel = True
                
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.EditText
                
                ShowSimpleMsg "��ǰ����ţ�" & Trim(vsfPerson.EditText) & "�Ѿ����ڣ��������ظ���"
                vsfPerson.EditText = ""
                vsfPerson.TextMatrix(Row, Col) = ""
                
            End If
        End If
    End Select
End Sub

Private Sub vsfPrice_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)

    Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
    
    DataChange = True
End Sub

Private Sub vsfPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With vsfPrice
        Select Case Col
        Case mCol.p�Ƽ���Ŀ
        
            If Left(.TextMatrix(Row, mCol.p�Ƽ���Ŀ), 4) = "�ɼ���ʽ" Then
                .TextMatrix(Row, mCol.p�Ƽ�����) = "2"
            Else
                .TextMatrix(Row, mCol.p�Ƽ�����) = "1"
            End If
            .TextMatrix(Row, mCol.p�Ƽ���Ŀ) = .Cell(flexcpTextDisplay, Row, mCol.p�Ƽ���Ŀ)
            
        Case mCol.p����
            vsfPrice.TextMatrix(Row, mCol.p��׼���) = Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)) * Val(vsfPrice.TextMatrix(Row, mCol.p����))
            vsfPrice.TextMatrix(Row, mCol.p�����) = Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)) * Val(vsfPrice.TextMatrix(Row, mCol.p����))
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
                    
            If InStr("567", .TextMatrix(Row, mCol.p���)) > 0 Then
                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p����)), Val(.TextMatrix(Row, mCol.p���ÿ��)), .TextMatrix(Row, mCol.p����), .TextMatrix(Row, mCol.pִ�п���), .TextMatrix(Row, mCol.p���㵥λ), 1)
            End If
                
        Case mCol.p��쵥��
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
        
        Case mCol.p�ۿ�
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p�ۿ�)), 2)
        
        Case mCol.pִ�п���
            .TextMatrix(Row, mCol.pִ�п���id) = .Body.ComboData
            .TextMatrix(Row, mCol.pִ�п���) = .Cell(flexcpTextDisplay, Row, mCol.pִ�п���)
            
            If InStr("567", .TextMatrix(Row, mCol.p���)) > 0 Then
                .TextMatrix(Row, mCol.p���ÿ��) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.pִ�п���id)))
                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p����)), Val(.TextMatrix(Row, mCol.p���ÿ��)), .TextMatrix(Row, mCol.p����), .TextMatrix(Row, mCol.pִ�п���), .TextMatrix(Row, mCol.p���㵥λ), 1)
            End If
                    
        End Select
    End With
    
    DataChange = True
    
End Sub

Private Sub vsfPrice_AfterNewRow(ByVal Row As Long, Col As Long)
    
    If Row > 1 Then
        vsfPrice.TextMatrix(Row, mCol.p�Ƽ���Ŀ) = vsfPrice.TextMatrix(Row - 1, mCol.p�Ƽ���Ŀ)
        
        If Left(vsfPrice.TextMatrix(Row, mCol.p�Ƽ���Ŀ), 4) = "�ɼ���ʽ" Then
            vsfPrice.TextMatrix(Row, mCol.p�Ƽ�����) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.p�Ƽ�����) = "1"
        End If
        
    End If
    
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call SetRowData(Val(vsfPrice.RowData(NewRow)), NewRow, "�շ�ִ�п���")
End Sub

Private Sub vsfPrice_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str�Ƽ���Ŀ As String
    Dim str�Ƽ����� As String
    
    If vsfPrice.Rows = 2 Then
        
        str�Ƽ���Ŀ = vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ)
        str�Ƽ����� = vsfPrice.TextMatrix(1, mCol.p�Ƽ�����)
        
        vsfPrice.Body.Cell(flexcpText, 1, mCol.p�Ƽ���Ŀ + 1, 1, vsfPrice.Cols - 1) = ""
        vsfPrice.RowData(1) = 0

        vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ) = str�Ƽ���Ŀ
        vsfPrice.TextMatrix(1, mCol.p�Ƽ�����) = str�Ƽ�����
        Call vsfPrice_AfterDeleteRow(1, Col)
        
        Cancel = True
    End If
End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If Col = mCol.p���� Then
        
        
        gstrSQL = GetPublicSQL(SQL.�շ���Ŀѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowGrdSelect(Me, vsfPrice, "����,1200,0,1;����,2700,0,0;��λ,600,0,0;���,1200,0,0;����,900,0,0;���,900,0,0", Me.Name & "\�շ���Ŀѡ��", "����б���ѡ��һ���շ���Ŀ��", rsData, rs, 8790, 5100) Then

            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                Exit Sub
            End If
            With vsfPrice
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mCol.p����) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mCol.p���㵥λ) = zlCommFun.NVL(rs("��λ").Value)
    
                .TextMatrix(Row, mCol.p��׼����) = zlCommFun.NVL(rs("����").Value, 0)
                .TextMatrix(Row, mCol.p��쵥��) = .TextMatrix(Row, mCol.p��׼����)
    
                .TextMatrix(Row, mCol.p�շ���Ŀid) = zlCommFun.NVL(rs("ID").Value, 0)
                If Val(.TextMatrix(Row, mCol.p����)) < 1 Then .TextMatrix(Row, mCol.p����) = 1
    
                .TextMatrix(Row, mCol.p��׼���) = Val(.TextMatrix(Row, mCol.p��׼����)) * Val(.TextMatrix(Row, mCol.p����))
                .TextMatrix(Row, mCol.p�����) = .TextMatrix(Row, mCol.p��׼���)
                
                .TextMatrix(Row, mCol.p���) = zlCommFun.NVL(rs("���").Value)
                
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
    
                Call SetRowDefault(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                Call SetRowData(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                
                Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
                
                If InStr("567", .TextMatrix(Row, mCol.p���)) > 0 Then
                    .TextMatrix(Row, mCol.p���ÿ��) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.pִ�п���id)))
                    Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p����)), Val(.TextMatrix(Row, mCol.p���ÿ��)), .TextMatrix(Row, mCol.p����), .TextMatrix(Row, mCol.pִ�п���), .TextMatrix(Row, mCol.p���㵥λ), 1)
                End If
'
            End With
            
            DataChange = True

        End If
        
    End If
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsfPrice.EditText, "'") > 0 Then
                KeyCode = 0
                vsfPrice.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.p����
                    
                    strText = UCase(vsfPrice.EditText)
                    gstrSQL = GetPublicSQL(SQL.�շ���Ŀ����, strText)
                    
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
                    
                    If ShowGrdFilter(Me, vsfPrice, "����,1200,0,1;����,2700,0,0;��λ,600,0,0;���,1200,0,0;����,900,0,0;���,900,0,0", Me.Name & "\�շ���Ŀ����", "����б���ѡ��һ���շ���Ŀ��", rsData, rs, 8790, 5100) Then
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                        
                        With vsfPrice
                            .EditText = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, mCol.p����) = zlCommFun.NVL(rs("����").Value)
                            
                            .Cell(flexcpData, Row, mCol.p����, Row, mCol.p����) = zlCommFun.NVL(rs("����").Value)
                            
                            .TextMatrix(Row, mCol.p���㵥λ) = zlCommFun.NVL(rs("��λ").Value)
                            
                            .TextMatrix(Row, mCol.p��׼����) = zlCommFun.NVL(rs("����").Value, 0)
                            .TextMatrix(Row, mCol.p��쵥��) = .TextMatrix(Row, mCol.p��׼����)
                            
                            .TextMatrix(Row, mCol.p�շ���Ŀid) = zlCommFun.NVL(rs("ID").Value, 0)
                            If Val(.TextMatrix(Row, mCol.p����)) < 1 Then .TextMatrix(Row, mCol.p����) = 1
                            
                            .TextMatrix(Row, mCol.p��׼���) = Val(.TextMatrix(Row, mCol.p��׼����)) * Val(.TextMatrix(Row, mCol.p����))
                            .TextMatrix(Row, mCol.p�����) = .TextMatrix(Row, mCol.p��׼���)
                            .TextMatrix(Row, mCol.p���) = zlCommFun.NVL(rs("���").Value)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                            
                            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
                            
                            Call SetRowDefault(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                            Call SetRowData(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                            
                            If InStr("567", .TextMatrix(Row, mCol.p���)) > 0 Then
                                .TextMatrix(Row, mCol.p���ÿ��) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.pִ�п���id)))
                                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p����)), Val(.TextMatrix(Row, mCol.p���ÿ��)), .TextMatrix(Row, mCol.p����), .TextMatrix(Row, mCol.pִ�п���), .TextMatrix(Row, mCol.p���㵥λ), 1)
                            End If
                        End With
                        
                        DataChange = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsfPrice.Cell(flexcpData, Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.EditText = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.TextMatrix(Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        DataChange = True
    End If
End Sub








