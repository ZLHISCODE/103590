VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGuide 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "frmGuide.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Tag             =   "������"
   Begin VB.PictureBox picGuide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4605
      Index           =   0
      Left            =   2490
      ScaleHeight     =   4605
      ScaleWidth      =   5130
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "��������"
      Top             =   0
      Width           =   5130
      Begin VB.OptionButton optType 
         Caption         =   "�򵥽�����ܱ�(&2)"
         Height          =   255
         Index           =   1
         Left            =   1095
         TabIndex        =   21
         Top             =   2295
         Width           =   1830
      End
      Begin VB.OptionButton optType 
         Caption         =   "����������(&1)"
         Height          =   255
         Index           =   0
         Left            =   1095
         TabIndex        =   0
         Top             =   1845
         Value           =   -1  'True
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ������Ҫ����ı������ͣ�"
         Height          =   180
         Left            =   345
         TabIndex        =   68
         Top             =   1140
         Width           =   2520
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "��ӭ��ʹ���Զ��屨���򵼣�ͨ������,���������ɡ���ݵض����������Ҫ�ı���"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   345
         TabIndex        =   67
         Top             =   570
         Width           =   4515
      End
   End
   Begin VB.PictureBox picGuide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4665
      Index           =   2
      Left            =   2520
      ScaleHeight     =   4665
      ScaleWidth      =   5130
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   "������Դ"
      Top             =   -15
      Visible         =   0   'False
      Width           =   5130
      Begin VB.CommandButton cmdMainConn 
         Caption         =   "��"
         Height          =   285
         Left            =   4755
         TabIndex        =   3
         Top             =   330
         Width           =   300
      End
      Begin VB.ComboBox cboMainConn 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   3225
      End
      Begin MSComctlLib.ListView lvwSub 
         Height          =   1500
         Left            =   195
         TabIndex        =   16
         Top             =   3090
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   2646
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��������Դ"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Դ������"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��Դ������"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.ComboBox cboFor 
         Height          =   300
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2220
         Width           =   2955
      End
      Begin VB.ComboBox cboKey 
         Height          =   300
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1890
         Width           =   2955
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "�Ƴ�(&R)"
         Height          =   350
         Left            =   3960
         TabIndex        =   14
         ToolTipText     =   "�Ƴ�"
         Top             =   2610
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   2685
         TabIndex        =   13
         ToolTipText     =   "����"
         Top             =   2610
         Width           =   1100
      End
      Begin VB.ComboBox cboSub 
         Height          =   300
         Left            =   2115
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1545
         Width           =   2955
      End
      Begin VB.ComboBox cboMain 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   675
         Width           =   3540
      End
      Begin VB.Label lblMainConn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������(&T)"
         Height          =   180
         Left            =   255
         TabIndex        =   1
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������Դ(&S)"
         Height          =   180
         Left            =   900
         TabIndex        =   7
         Top             =   1605
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������Դ�嵥(&L)"
         Height          =   180
         Left            =   255
         TabIndex        =   15
         Top             =   2865
         Width           =   1710
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Դ������(&F)"
         Height          =   180
         Left            =   900
         TabIndex        =   11
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Դ������(&K)"
         Height          =   180
         Left            =   900
         TabIndex        =   9
         Top             =   1950
         Width           =   1170
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ҫ��������ѡ��������������Դ������ĸ���������Դ����Ҫ��ȷ������������������Դ֮��Ķ�Ӧ��ϵ��"
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   210
         TabIndex        =   69
         Top             =   1035
         Width           =   4695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������Դ(&M)"
         Height          =   180
         Left            =   255
         TabIndex        =   4
         Top             =   705
         Width           =   1170
      End
   End
   Begin VB.PictureBox picGuide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4665
      Index           =   3
      Left            =   15
      ScaleHeight     =   4665
      ScaleWidth      =   7620
      TabIndex        =   58
      TabStop         =   0   'False
      Tag             =   "�����ʽ"
      Top             =   -15
      Visible         =   0   'False
      Width           =   7620
      Begin VB.ComboBox cboAlign 
         Height          =   300
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4170
         Width           =   1635
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3825
         Width           =   1635
      End
      Begin MSComctlLib.ListView lvwState 
         Height          =   1155
         Left            =   3495
         TabIndex        =   38
         Top             =   3390
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   2037
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "����"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView lvwHsc 
         Height          =   1155
         Left            =   3495
         TabIndex        =   35
         Top             =   1890
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   2037
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView lvwVsc 
         Height          =   1155
         Left            =   3495
         TabIndex        =   32
         Top             =   375
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   2037
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.CommandButton cmdDelState 
         Caption         =   "<"
         Height          =   345
         Left            =   3060
         TabIndex        =   39
         Tag             =   "�Ƴ�ͳ����"
         Top             =   3975
         Width           =   390
      End
      Begin VB.CommandButton cmdAddState 
         Caption         =   ">"
         Height          =   345
         Left            =   3060
         TabIndex        =   37
         Tag             =   "����ͳ����"
         Top             =   3615
         Width           =   390
      End
      Begin VB.CommandButton cmdDelHsc 
         Caption         =   "<"
         Height          =   345
         Left            =   3060
         TabIndex        =   36
         Tag             =   "�Ƴ��������"
         Top             =   2490
         Width           =   390
      End
      Begin VB.CommandButton cmdAddHsc 
         Caption         =   ">"
         Height          =   345
         Left            =   3060
         TabIndex        =   34
         Tag             =   "����������"
         Top             =   2130
         Width           =   390
      End
      Begin VB.CommandButton cmdDelVsc 
         Caption         =   "<"
         Height          =   345
         Left            =   3045
         TabIndex        =   33
         Tag             =   "�Ƴ��������"
         Top             =   930
         Width           =   390
      End
      Begin VB.CommandButton cmdAddVsc 
         Caption         =   ">"
         Height          =   345
         Left            =   3045
         TabIndex        =   31
         Tag             =   "�����������"
         Top             =   570
         Width           =   390
      End
      Begin VB.TextBox txtAS 
         Height          =   300
         Left            =   795
         TabIndex        =   28
         Top             =   3480
         Width           =   1635
      End
      Begin VB.ListBox lstAll 
         Height          =   2940
         Left            =   105
         TabIndex        =   27
         Top             =   360
         Width           =   2880
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   75
         Top             =   4230
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   74
         Top             =   3885
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ������н����Ҫ���ܵ�����"
         Height          =   180
         Left            =   3540
         TabIndex        =   73
         Top             =   3150
         Width           =   2700
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������Ŀ����Ϊ�б��������ܵ���Ŀ"
         Height          =   180
         Left            =   3540
         TabIndex        =   72
         Top             =   1650
         Width           =   3420
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������Ŀ����Ϊ�б��������ܵ���Ŀ"
         Height          =   180
         Left            =   3540
         TabIndex        =   71
         Top             =   135
         Width           =   3420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   70
         Top             =   3540
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������Ŀ(&U)"
         Height          =   180
         Left            =   195
         TabIndex        =   26
         Top             =   120
         Width           =   1350
      End
   End
   Begin VB.PictureBox picGuide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4635
      Index           =   1
      Left            =   2445
      ScaleHeight     =   4635
      ScaleWidth      =   5205
      TabIndex        =   56
      TabStop         =   0   'False
      Tag             =   "������Դ"
      Top             =   0
      Visible         =   0   'False
      Width           =   5205
      Begin VB.CommandButton cmdConn 
         Caption         =   "��"
         Height          =   285
         Left            =   4485
         TabIndex        =   18
         Top             =   345
         Width           =   300
      End
      Begin VB.ComboBox cboConn 
         Height          =   300
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   345
         Width           =   4380
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "��"
         Height          =   435
         Left            =   4530
         TabIndex        =   25
         ToolTipText     =   "����"
         Top             =   2865
         Width           =   495
      End
      Begin VB.CommandButton cmdUP 
         Caption         =   "��"
         Height          =   435
         Left            =   4530
         TabIndex        =   24
         ToolTipText     =   "����"
         Top             =   2265
         Width           =   495
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2970
         Left            =   105
         TabIndex        =   23
         Top             =   1515
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   5239
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.ComboBox cboList 
         Height          =   300
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   930
         Width           =   4380
      End
      Begin VB.Label lblConnect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&T)"
         Height          =   180
         Left            =   165
         TabIndex        =   5
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�븴ѡҪ��Ϊ������Ŀ,��˫������Ŀ��������(&L)"
         Height          =   180
         Left            =   165
         TabIndex        =   22
         Top             =   1305
         Width           =   4140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������Դ(&S)"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   675
         Width           =   1350
      End
   End
   Begin VB.PictureBox picGuide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Index           =   4
      Left            =   2490
      ScaleHeight     =   4605
      ScaleWidth      =   5130
      TabIndex        =   59
      TabStop         =   0   'False
      Tag             =   "��������"
      Top             =   30
      Visible         =   0   'False
      Width           =   5130
      Begin VB.ComboBox cboOper 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   855
         Width           =   855
      End
      Begin VB.CommandButton cmdDelIf 
         Caption         =   "�Ƴ�(&R)"
         Height          =   350
         Left            =   2910
         TabIndex        =   48
         Top             =   1770
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddIf 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   1410
         TabIndex        =   47
         Top             =   1770
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvwIF 
         Height          =   1455
         Left            =   180
         TabIndex        =   49
         Top             =   2280
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   2566
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��������"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ӧ��Ŀ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ȱʡֵ"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.ComboBox cboValue 
         Height          =   300
         Left            =   1215
         TabIndex        =   46
         Top             =   1290
         Width           =   3195
      End
      Begin VB.ComboBox cboIF 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   435
         Width           =   3180
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   43
         Top             =   855
         Width           =   2310
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "��������������ÿ�δ򿪱���ʱ���ò�ͬ������ֵ���Ի�ȡ����Ҫ�����ݡ�"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   300
         TabIndex        =   76
         Top             =   3900
         Width           =   4230
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡֵ(&D)"
         Height          =   180
         Left            =   360
         TabIndex        =   45
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ��Ŀ(&I)"
         Height          =   180
         Left            =   180
         TabIndex        =   40
         Top             =   495
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   42
         Top             =   915
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   225
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "F1"
      Top             =   4875
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5250
      TabIndex        =   52
      ToolTipText     =   "Ctrl+Enter"
      Top             =   4875
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6420
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "ESC"
      Top             =   4875
      Width           =   1100
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Height          =   350
      Left            =   3660
      TabIndex        =   51
      Top             =   4875
      Width           =   1100
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "��һ��(&P)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2265
      TabIndex        =   50
      Top             =   4875
      Width           =   1100
   End
   Begin VB.Frame fraFont 
      Height          =   120
      Left            =   -60
      TabIndex        =   55
      Top             =   4620
      Width           =   8595
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00EBFFFF&
      Height          =   300
      Left            =   3555
      MaxLength       =   30
      TabIndex        =   64
      Top             =   1995
      Width           =   3495
   End
   Begin VB.TextBox txtNO 
      BackColor       =   &H00EBFFFF&
      Height          =   300
      Left            =   3555
      MaxLength       =   20
      TabIndex        =   62
      Top             =   1545
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.TextBox txtNote 
      BackColor       =   &H00EBFFFF&
      Height          =   300
      Left            =   3555
      TabIndex        =   66
      Top             =   2445
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����˵��"
      Height          =   180
      Left            =   2760
      TabIndex        =   65
      Top             =   2505
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   2760
      TabIndex        =   61
      Top             =   1605
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   180
      Left            =   2760
      TabIndex        =   63
      Top             =   2055
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuide.frx":014A
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   2670
      TabIndex        =   60
      Top             =   810
      Width           =   4590
   End
   Begin VB.Image imgFlag 
      BorderStyle     =   1  'Fixed Single
      Height          =   4500
      Left            =   135
      Stretch         =   -1  'True
      Top             =   90
      Width           =   2250
   End
End
Attribute VB_Name = "frmGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmParent As Object '��
Public objGuide As Report '�����򵼲����ı�������
Public objReport As Report '�룺ԭ�б�������(ֻ��)
Public mobjFmt As RPTFmt '�룺ԭ�б���ǰ��ʽ(ֻ��)
Public blnNew As Boolean '�룺�����Ƿ�ȫ�²���һ������

Private bytStep As Byte '��ǰ��������
Private mcolRS As New Collection       '�����������ӵļ���

Private Const MSTR_CN_ITEM  As String = "��ǰ��¼"
Private Const MSTR_OWNER_FILTER As String = _
    "OWNER<>'SYS' and OWNER<>'SYSTEM' and OWNER<>'SCOTT' and OWNER<>'OUTLN' and OWNER<>'DBSNMP' and OWNER<>'MTSSYS'" & _
    " and OWNER<>'MDSYS' and OWNER<>'ORDSYS' and OWNER<>'ORDPLUGINS' and OWNER<>'CTXSYS' and OWNER<>'ZLTOOLS'" & _
    " and OWNER<>'XDB' and OWNER<>'WMSYS' and OWNER<>'TSMSYS' and OWNER<>'SYSMAN' and OWNER<>'SI_INFORMTN_SCHEMA'" & _
    " and OWNER<>'OLAPSYS' and OWNER<>'MGMT_VIEW' and OWNER<>'MDDATA' and OWNER<>'EXFSYS' and OWNER<>'DMSYS'" & _
    " and OWNER<>'DIP' and OWNER<>'ANONYMOUS'"

Private Function GetName(str As String, bytAlign As Byte) As String
'���ܣ��ֽ���"ZLHIS.���ű�.����"�������ַ���
    If InStr(str, ".") = 0 Then GetName = str: Exit Function
    Select Case bytAlign
        Case 0
            GetName = Left(str, InStr(str, ".") - 1)
        Case 1
            GetName = Mid(str, InStr(str, ".") + 1)
            If InStr(GetName, ".") = 0 Then Exit Function
            GetName = Left(GetName, InStr(GetName, ".") - 1)
        Case 2
            GetName = Mid(str, InStr(str, ".") + 1)
            If InStr(GetName, ".") = 0 Then Exit Function
            GetName = Mid(GetName, InStr(GetName, ".") + 1)
    End Select
End Function

Private Function OpenRecord(strObject As String, Optional ByVal intConnect As Integer) As ADODB.Recordset
'���ܣ���һ�������ͼ,�Ի�ȡ���ֶ�����
'������
'  strObject=��������ǰ׺�Ķ�����
'  intConnect=�������ӱ��

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim cn As ADODB.Connection
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    '׼���������Ӷ���
    Set cn = mdlPublic.GetDBConnection(intConnect)
    
    '��Rownum< 1 ��ȫ����ʱ��
    strSQL = "Select * From " & strObject & " Where Rownum<1"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, cn, adOpenKeyset
    Set rsTmp.ActiveConnection = Nothing
    Set OpenRecord = rsTmp
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Sub cboConn_Click()
    If Me.Visible = False Then Exit Sub
    If Val(cboConn.Tag) = cboConn.ListIndex Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Call SetSourceControl(cboConn, cboList)
    
    cboList.ListIndex = -1
    cboList.Tag = CStr(cboList.ListIndex)
    Call cboList_Click
    
    cboConn.Tag = CStr(cboConn.ListIndex)
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
hErr:
    Screen.MousePointer = vbDefault
    Call mdlPublic.ErrCenter
End Sub

Private Sub cboIF_Click()
    If InStr(cboIF.Text, ".") > 0 Then
        txtName.Text = GetName(cboIF.Text, 2)
    Else
        txtName.Text = cboIF.Text
    End If
    
    If cboIF.ItemData(cboIF.ListIndex) = adDBTimeStamp Then
        cboValue.ListIndex = 0
    Else
        cboValue.Text = ""
    End If
End Sub

Private Sub cboMain_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If Val(cboMain.Tag) = cboMain.ListIndex Then Exit Sub
    
    lvwSub.ListItems.Clear
    cboMain.Tag = cboMain.ListIndex
    
    cboKey.Clear
    Set rsTmp = OpenRecord(cboMain.Text, cboMainConn.ItemData(cboMainConn.ListIndex))
    If Not rsTmp Is Nothing Then
        For i = 0 To rsTmp.Fields.count - 1
            If Not IsType(rsTmp.Fields(i).type, adLongVarBinary) Then
                cboKey.AddItem rsTmp.Fields(i).name
                cboKey.ItemData(i) = rsTmp.Fields(i).type '�������
            End If
        Next
    End If
    
    cboMain.Tag = CStr(cboMain.ListIndex)
End Sub

Private Sub cboMainConn_Click()
    Dim i As Integer

    If Me.Visible = False Then Exit Sub
    If Val(cboMainConn.Tag) = cboMainConn.ListIndex Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Call SetSourceControl(cboMainConn, cboMain)
    cboMain.ListIndex = -1
    cboMain.Tag = CStr(cboMain.ListIndex)
    Call cboMain_Click
    
    cboSub.Clear
    For i = 1 To cboMain.ListCount
        cboSub.AddItem cboMain.List(i)
    Next
    cboSub.ListIndex = -1
    cboSub.Tag = CStr(cboSub.ListIndex)
    Call cboSub_Click
    
    cboMainConn.Tag = CStr(cboMainConn.ListIndex)
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
hErr:
    Screen.MousePointer = vbDefault
    Call mdlPublic.ErrCenter
End Sub

Private Sub cboSub_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If Val(cboSub.Tag) = cboSub.ListIndex Then Exit Sub
    
    cboFor.Clear
    Set rsTmp = OpenRecord(cboSub.Text, cboMainConn.ItemData(cboMainConn.ListIndex))
    If Not rsTmp Is Nothing Then
        For i = 0 To rsTmp.Fields.count - 1
            If Not IsType(rsTmp.Fields(i).type, adLongVarBinary) Then
                cboFor.AddItem rsTmp.Fields(i).name
                cboFor.ItemData(i) = rsTmp.Fields(i).type '�������
            End If
        Next
    End If
    
    cboSub.Tag = CStr(cboSub.ListIndex)
End Sub

Private Sub cboValue_GotFocus()
    SelAll cboValue
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
    If InStr("&~`!@#$^""��" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    Dim objItem As Object, i As Integer
    Dim blnSame As Boolean, strFields As String
    
    If cboMain.ListIndex = -1 Then
        MsgBox "��ѡ����������Դ��", vbInformation, App.Title: cboMain.SetFocus: Exit Sub
    End If
    
    If cboSub.ListIndex = -1 Then
        MsgBox "��ѡ������������Դ��", vbInformation, App.Title: cboSub.SetFocus: Exit Sub
    End If
    If cboSub.Text = cboMain.Text Then
        MsgBox "����������Դ��������������Դ��ͬ��", vbInformation, App.Title: cboSub.SetFocus: Exit Sub
    End If
    If cboKey.ListIndex = -1 Then
        MsgBox "��ѡ����������Դ�����", vbInformation, App.Title: cboKey.SetFocus: Exit Sub
    End If
    If cboFor.ListIndex = -1 Then
        MsgBox "��ѡ������������Դ�����", vbInformation, App.Title: cboFor.SetFocus: Exit Sub
    End If
    
    For i = 1 To lvwSub.ListItems.count
        If lvwSub.ListItems(i).Text = cboSub.Text Then
            If lvwSub.ListItems(i).SubItems(1) = cboFor.Text Or lvwSub.ListItems(i).SubItems(2) = cboKey.Text Then
                MsgBox "����������Դ�������������Դ�������Ѿ����룡", vbInformation, App.Title: cboSub.SetFocus: Exit Sub
            End If
        End If
    Next
    
    If cboFor.ItemData(cboFor.ListIndex) = cboKey.ItemData(cboKey.ListIndex) Then
        blnSame = True
    Else
        Select Case cboFor.ItemData(cboFor.ListIndex)
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                If IsType(cboKey.ItemData(cboKey.ListIndex), adVarChar) Then blnSame = True
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                If IsType(cboKey.ItemData(cboKey.ListIndex), adVarNumeric) Then blnSame = True
        End Select
    End If
    If Not blnSame Then MsgBox "������������������Ͳ���ͬ��", vbInformation, App.Title: cboFor.SetFocus: Exit Sub
    
    For i = 0 To cboFor.ListCount - 1
        strFields = strFields & "|" & cboFor.List(i) & "," & cboFor.ItemData(i)
    Next
    strFields = Mid(strFields, 2)
    
    Set objItem = lvwSub.ListItems.Add(, , cboSub.Text)
    objItem.SubItems(1) = cboFor.Text
    objItem.SubItems(2) = cboKey.Text
    objItem.Tag = strFields '�����Դ�ֶ��б�,���ٺ�����ٶ�
    Set lvwSub.SelectedItem = objItem
    lvwSub.SelectedItem.EnsureVisible
    cboFor.ListIndex = -1: cboKey.ListIndex = -1
    cboSub.SetFocus
End Sub

Private Sub cmdAddHsc_Click()
'����������
    Dim objItem As Object, i As Integer
    
    If Not (IsType(lstAll.ItemData(lstAll.ListIndex), adVarChar) Or IsType(lstAll.ItemData(lstAll.ListIndex), adDBTimeStamp)) Then
        MsgBox "ֻ���ַ�����������Ŀ���ʺ���Ϊ������Ŀ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    End If
    For i = 1 To lvwVsc.ListItems.count
        If lvwVsc.ListItems(i).SubItems(1) = txtAS.Text Then
            MsgBox "�ñ����������Ѿ�����������࣡", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
        End If
    Next
    For i = 1 To lvwHsc.ListItems.count
        If lvwHsc.ListItems(i).SubItems(1) = txtAS.Text Then
            MsgBox "�ñ����������Ѿ����������࣡", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
        End If
    Next
    
    Set objItem = lvwHsc.ListItems.Add(, , lstAll.Text)
    objItem.SubItems(1) = txtAS.Text
    If cboState.ListIndex <> 0 Then objItem.SubItems(2) = cboState.Text
    objItem.Tag = lstAll.ItemData(lstAll.ListIndex) '�������
    Set lvwHsc.SelectedItem = objItem
    lvwHsc.SelectedItem.EnsureVisible
    lstAll.SetFocus
End Sub

Private Sub cmdAddIf_Click()
    Dim objItem As Object, i As Integer
    
    If cboIF.ListIndex = -1 Then
        MsgBox "��ѡ��������Ŀ��", vbInformation, App.Title: cboIF.SetFocus: Exit Sub
    End If
    
    If txtName.Text = "" Then
        MsgBox "�������������ƣ�", vbInformation, App.Title: txtName.SetFocus: Exit Sub
    End If
    If TLen(txtName.Text) > 20 Then
        MsgBox "�������Ʋ��ܳ���20���ַ���", vbInformation, App.Title: txtName.SetFocus: Exit Sub
    End If
    
'    If cboValue.Text = "" Then
'        MsgBox "������ȱʡ������ֵ��", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
'    End If
    
    If TLen(cboValue.Text) > 255 Then
        MsgBox "ȱʡ����ֵ���Ȳ��ܳ���255���ַ���", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
    End If
    
    Select Case cboIF.ItemData(cboIF.ListIndex)
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            If cboOper.Text = "IN" And cboValue.Text <> "" Then
                For i = 0 To UBound(Split(cboValue.Text, ","))
                    If Left(Trim(Split(cboValue.Text, ",")(i)), 1) <> "'" Or _
                        Right(Trim(Split(cboValue.Text, ",")(i)), 1) <> "'" Then
                        MsgBox "IN�����ȱʡֵ����,ӦΪ "" '��A','��B','��C'..."" ����ʽ��", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
                    End If
                Next
            End If
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            If cboOper.Text = "IN" And cboValue.Text <> "" Then
                For i = 0 To UBound(Split(cboValue.Text, ","))
                    If Not IsNumeric(Trim(Split(cboValue.Text, ",")(i))) Then
                        MsgBox "IN�����ȱʡֵ����,ӦΪ"" ֵA,ֵB,ֵC..."" ����ʽ��", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
                    End If
                Next
            End If
            If cboOper.Text = "LIKE" Then
                MsgBox "LIKE��������ʺ��������͵���Ŀ��", vbInformation, App.Title: cboOper.SetFocus: Exit Sub
            End If
            If Not IsNumeric(cboValue.Text) And cboValue.Text <> "" Then
                MsgBox "����ȱʡֵ���������������ݣ�", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
            End If
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            If cboOper.Text = "IN" And cboValue.Text <> "" Then
                For i = 0 To UBound(Split(cboValue.Text, ","))
                    If Left(Trim(Split(cboValue.Text, ",")(i)), 1) <> "'" Or _
                        Right(Trim(Split(cboValue.Text, ",")(i)), 1) <> "'" Then
                        MsgBox "IN�����ȱʡֵ����,ӦΪ "" '����A','����B','����C'..."" ����ʽ��", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
                    ElseIf Not IsDate(Mid(Trim(Split(cboValue.Text, ",")(i)), 2, Len(Trim(Split(cboValue.Text, ",")(i))) - 1)) Then
                        MsgBox "IN�����ȱʡֵ����,ӦΪ "" '����A','����B','����C'..."" ����ʽ��", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
                    End If
                Next
            End If
            If cboOper.Text = "LIKE" Then
                MsgBox "LIKE��������ʺ��������͵���Ŀ��", vbInformation, App.Title: cboOper.SetFocus: Exit Sub
            End If
            If Not IsDate(cboValue.Text) And cboValue.ListIndex = -1 And cboValue.Text <> "" Then
                MsgBox "����ȱʡֵ���������������ݣ�", vbInformation, App.Title: cboValue.SetFocus: Exit Sub
            End If
    End Select
    
    For i = 1 To lvwIF.ListItems.count
        If lvwIF.ListItems(i).Text = txtName.Text Then
            MsgBox "�����������з�������ͬ���Ƶ���������,������������ƣ�", vbInformation, App.Title: txtName.SetFocus: Exit Sub
        End If
    Next
    
    Set objItem = lvwIF.ListItems.Add(, , txtName.Text)
    objItem.SubItems(1) = cboIF.Text
    objItem.SubItems(2) = cboOper.Text
    objItem.SubItems(3) = cboValue.Text
    objItem.Tag = cboIF.ItemData(cboIF.ListIndex) '�������
    Set lvwIF.SelectedItem = objItem
    lvwIF.SelectedItem.EnsureVisible
    
    cboIF.SetFocus
End Sub

Private Sub cmdAddState_Click()
'����ͳ��
    Dim objItem As Object, i As Integer
    
    If cboState.ListIndex = 0 Then MsgBox "ͳ����Ŀ����ѡ����ܷ�ʽ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    If Not IsType(lstAll.ItemData(lstAll.ListIndex), adNumeric) Then
        MsgBox "ֻ����������Ŀ���ʺ���Ϊͳ����Ŀ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    End If
    For i = 1 To lvwState.ListItems.count
        If lvwState.ListItems(i).SubItems(1) = txtAS.Text Then
            MsgBox "�ñ�����Ŀ�Ѿ�������ͳ����Ŀ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
        End If
    Next
    
    Set objItem = lvwState.ListItems.Add(, , lstAll.Text)
    objItem.SubItems(1) = txtAS.Text
    objItem.SubItems(2) = cboState.Text
    objItem.SubItems(3) = cboAlign.Text
    objItem.Tag = lstAll.ItemData(lstAll.ListIndex)
    Set lvwState.SelectedItem = objItem
    lvwState.SelectedItem.EnsureVisible
    lstAll.SetFocus
End Sub

Private Sub cmdAddVsc_Click()
'�����������
    Dim objItem As Object, i As Integer
    
    If Not (IsType(lstAll.ItemData(lstAll.ListIndex), adVarChar) Or IsType(lstAll.ItemData(lstAll.ListIndex), adDBTimeStamp)) Then
        MsgBox "ֻ���ַ�����������Ŀ���ʺ���Ϊ������Ŀ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    End If
    For i = 1 To lvwVsc.ListItems.count
        If lvwVsc.ListItems(i).SubItems(1) = txtAS.Text Then
            MsgBox "�ñ����������Ѿ�����������࣡", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
        End If
    Next
    For i = 1 To lvwHsc.ListItems.count
        If lvwHsc.ListItems(i).SubItems(1) = txtAS.Text Then
            MsgBox "�ñ����������Ѿ����������࣡", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
        End If
    Next
    
    Set objItem = lvwVsc.ListItems.Add(, , lstAll.Text)
    objItem.SubItems(1) = txtAS.Text
    If cboState.ListIndex <> 0 Then objItem.SubItems(2) = cboState.Text
    objItem.Tag = lstAll.ItemData(lstAll.ListIndex)
    Set lvwVsc.SelectedItem = objItem
    lvwVsc.SelectedItem.EnsureVisible
    lstAll.SetFocus
End Sub

Private Sub cmdConn_Click()
    Dim blnModified As Boolean
    Dim intIndex As Integer
    
    If Me.Visible = False Then Exit Sub
    
    If gfrmDBConnect Is Nothing Then
        MsgBox "�����������ӹ���ʧ�ܣ�", vbInformation, App.Title
        Exit Sub
    End If
    
    intIndex = cboConn.ListIndex
    If gfrmDBConnect.ShowMe(Me, blnModified) Then
        If blnModified Then
            '�������Ӽ�¼������
            Call mdlPublic.SetControlDBConnect(grsConnect)
            '���µ�ǰ����
            cboConn.Clear
            cboConn.AddItem MSTR_CN_ITEM
            Call mdlPublic.SetControlDBConnect(cboConn)
            If intIndex > cboConn.ListCount Then
                cboConn.ListIndex = 0
            Else
                cboConn.ListIndex = intIndex
            End If
            '��ն��󼯺�
            Call gclsCNs.Clear
        End If
    End If
End Sub

Private Sub cmdDel_Click()
    If lvwSub.SelectedItem Is Nothing Then
        MsgBox "û������������Դ����ɾ����", vbInformation, App.Title: Exit Sub
    End If
    lvwSub.ListItems.Remove lvwSub.SelectedItem.Index
    cboSub.SetFocus
End Sub

Private Sub cmdDelHsc_Click()
    If lvwHsc.SelectedItem Is Nothing Then
        MsgBox "û�з�����Ŀ����ɾ����", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    End If
    lvwHsc.ListItems.Remove lvwHsc.SelectedItem.Index
    lstAll.SetFocus
End Sub

Private Sub cmdDelIf_Click()
    If lvwIF.SelectedItem Is Nothing Then
        MsgBox "û����������ɾ����", vbInformation, App.Title: cboIF.SetFocus:  Exit Sub
    End If
    lvwIF.ListItems.Remove lvwIF.SelectedItem.Index
    cboIF.SetFocus
End Sub

Private Sub cmdDelState_Click()
    If lvwState.SelectedItem Is Nothing Then
        MsgBox "û��ͳ����Ŀ����ɾ����", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    End If
    lvwState.ListItems.Remove lvwState.SelectedItem.Index
    lstAll.SetFocus
End Sub

Private Sub cmdDelVsc_Click()
    If lvwVsc.SelectedItem Is Nothing Then
        MsgBox "û�з�����Ŀ����ɾ����", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
    End If
    lvwVsc.ListItems.Remove lvwVsc.SelectedItem.Index
    lstAll.SetFocus
End Sub

Private Sub cmdDown_Click()
    Dim strText As String, strOrder As String, strTag As String, blnCheck As Boolean
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    If lvwItem.SelectedItem.Index = lvwItem.ListItems.count Then Exit Sub
    
    With lvwItem.ListItems(lvwItem.SelectedItem.Index + 1)
        strText = .Text
        strOrder = .SubItems(1)
        strTag = .Tag
        blnCheck = .Checked
        
        .Text = lvwItem.SelectedItem.Text
        .SubItems(1) = lvwItem.SelectedItem.SubItems(1)
        .Tag = lvwItem.SelectedItem.Tag
        .Checked = lvwItem.SelectedItem.Checked
    End With
    
    With lvwItem.SelectedItem
        .Text = strText
        .SubItems(1) = strOrder
        .Tag = strTag
        .Checked = blnCheck
    End With
    
    Set lvwItem.SelectedItem = lvwItem.ListItems(lvwItem.SelectedItem.Index + 1)
    lvwItem.SelectedItem.EnsureVisible
    lvwItem.SetFocus
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelpRpt(Me.hwnd, "guide", 0)
End Sub

Private Sub cmdMainConn_Click()
    Dim blnModified As Boolean
    Dim intIndex As Integer
    
    If Me.Visible = False Then Exit Sub
    
    If gfrmDBConnect Is Nothing Then
        MsgBox "�����������ӹ���ʧ�ܣ�", vbInformation, App.Title
        Exit Sub
    End If
    
    intIndex = cboConn.ListIndex
    If gfrmDBConnect.ShowMe(Me, blnModified) Then
        If blnModified Then
            '�������Ӽ�¼������
            Call mdlPublic.SetControlDBConnect(grsConnect)
            '���µ�ǰ����
            cboConn.Clear
            cboConn.AddItem MSTR_CN_ITEM
            Call mdlPublic.SetControlDBConnect(cboMainConn)
            If intIndex > cboConn.ListCount Then
                cboConn.ListIndex = 0
            Else
                cboConn.ListIndex = intIndex
            End If
            '��ն��󼯺�
            Call gclsCNs.Clear
        End If
    End If
End Sub

Private Sub cmdPre_Click()
    '��һ��
    Select Case bytStep
        Case 1 '�������
            picGuide(bytStep).Visible = False
            cmdPre.Enabled = False
            bytStep = 0
        Case 2 '��������
            picGuide(bytStep).Visible = False
            cmdPre.Enabled = False
            bytStep = 0
        Case 3 '�����Ų�
            picGuide(bytStep).Visible = False
            bytStep = 2
        Case 4 '��������
            picGuide(bytStep).Visible = False
            If optType(0).Value Then
                bytStep = 1
            Else
                bytStep = 3
            End If
        Case 5
            bytStep = 4
            cmdNext.Enabled = True
            cmdOK.Enabled = False
    End Select
    Caption = Tag & " - " & picGuide(bytStep).Tag
    picGuide(bytStep).ZOrder
    picGuide(bytStep).Visible = True
    Me.Refresh
    picGuide(bytStep).SetFocus
    SendKeys "{Tab}": SendKeys "{Tab}"
End Sub

Private Sub cmdUp_Click()
    Dim strText As String, strOrder As String, strTag As String, blnCheck As Boolean
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    If lvwItem.SelectedItem.Index = 1 Then Exit Sub
    
    With lvwItem.ListItems(lvwItem.SelectedItem.Index - 1)
        strText = .Text
        strOrder = .SubItems(1)
        strTag = .Tag
        blnCheck = .Checked
        
        .Text = lvwItem.SelectedItem.Text
        .SubItems(1) = lvwItem.SelectedItem.SubItems(1)
        .Tag = lvwItem.SelectedItem.Tag
        .Checked = lvwItem.SelectedItem.Checked
    End With
    
    With lvwItem.SelectedItem
        .Text = strText
        .SubItems(1) = strOrder
        .Tag = strTag
        .Checked = blnCheck
    End With
    
    Set lvwItem.SelectedItem = lvwItem.ListItems(lvwItem.SelectedItem.Index - 1)
    lvwItem.SelectedItem.EnsureVisible
    lvwItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 2 And cmdOK.Enabled Then cmdOK_Click
        Case vbKeyF1
            cmdHelp_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If MsgBox("�򵼲�����δ���,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Cancel = 1: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnNew = False
    Set objGuide = Nothing
    Set mcolRS = Nothing
    grsObject.Filter = 0
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("�򵼲�����δ���,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub lstAll_Click()
    txtAS.Text = GetName(lstAll.Text, 2)
    Select Case lstAll.ItemData(lstAll.ListIndex)
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR, _
            adDBTimeStamp, adDBTime, adDBDate, adDate '�ַ���������Ϊ������
            cboState.ListIndex = 0
            cboAlign.ListIndex = -1
            cboAlign.Enabled = False
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
            adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
            adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt '������Ϊͳ����
            cboState.ListIndex = 1
            cboAlign.ListIndex = 2
            cboAlign.Enabled = True
    End Select
End Sub

Private Sub lvwItem_DblClick()
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    If lvwItem.SelectedItem.SubItems(1) = "" Then
        lvwItem.SelectedItem.SubItems(1) = "��"
    ElseIf lvwItem.SelectedItem.SubItems(1) = "��" Then
        lvwItem.SelectedItem.SubItems(1) = "��"
    ElseIf lvwItem.SelectedItem.SubItems(1) = "��" Then
        lvwItem.SelectedItem.SubItems(1) = ""
    End If
End Sub

Private Sub txtAS_GotFocus()
    SelAll txtAS
End Sub

Private Sub txtAS_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&;"",", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&;"",'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&;"",'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtTitle_GotFocus()
    SelAll txtTitle
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    Set imgFlag.Picture = LoadCustomPicture("Report")
        
    grsObject.Filter = MSTR_OWNER_FILTER
    
    gblnOK = False

    If blnNew Then
        lblNO.Visible = True
        txtNO.Visible = True
        lblNote.Visible = True
        txtNote.Visible = True
    End If
    SetComboBoxHeight cboList, 350
    SetComboBoxHeight cboMain, 350
    SetComboBoxHeight cboSub, 350
    
    '��ʼ����
    cboConn.AddItem MSTR_CN_ITEM
    Call mdlPublic.SetControlDBConnect(cboConn)
    cboConn.ListIndex = 0
    cboConn.Tag = CStr(cboConn.ListIndex)
    
    cboMainConn.AddItem MSTR_CN_ITEM
    Call mdlPublic.SetControlDBConnect(cboMainConn)
    cboMainConn.ListIndex = 0
    cboMainConn.Tag = CStr(cboMainConn.ListIndex)
    
    For i = 1 To grsObject.RecordCount
        cboList.AddItem grsObject!Owner & "." & grsObject!OBJECT_NAME
        cboMain.AddItem grsObject!Owner & "." & grsObject!OBJECT_NAME
        cboSub.AddItem grsObject!Owner & "." & grsObject!OBJECT_NAME
        grsObject.MoveNext
    Next
    '��������
    cboValue.AddItem "&��ǰ����" '����ʽ
    cboValue.AddItem "&��ǰ����ʱ��"
    
    cboValue.AddItem "&���쿪ʼʱ��"
    cboValue.AddItem "&�������ʱ��"
    cboValue.AddItem "&ǰһ�쿪ʼʱ��"
    cboValue.AddItem "&ǰһ�����ʱ��"
    cboValue.AddItem "&ǰһ��ͬʱ��"
    cboValue.AddItem "&��һ��ͬʱ��"
    
    cboValue.AddItem "&ǰһ������"
    cboValue.AddItem "&ǰһ������"
    cboValue.AddItem "&ǰһ������"
    cboValue.AddItem "&ǰһ������"
    
    cboValue.AddItem "&��һ������"
    cboValue.AddItem "&��һ������"
    cboValue.AddItem "&��һ������"
    cboValue.AddItem "&��һ������"
    
    cboValue.AddItem "&���³�ʱ��"
    cboValue.AddItem "&����ĩʱ��"
    cboValue.AddItem "&���³�ʱ��"
    cboValue.AddItem "&����ĩʱ��"
    cboValue.AddItem "&�����ʱ��"
    cboValue.AddItem "&����ĩʱ��"
    cboValue.AddItem "&�����ʱ��"
    cboValue.AddItem "&����ĩʱ��"
    
    '���ܷ�ʽ(ע��:�����ڷ������ͳ����)
    '��ͳ����,ȱʡ���(��ѡ)
    '�Է�����,ȱʡ��(��ѡ)
    cboState.AddItem "��"
    cboState.AddItem "���"
    cboState.AddItem "��ƽ��ֵ"
    cboState.AddItem "�����ֵ"
    cboState.AddItem "����Сֵ"
    cboState.AddItem "���¼��"

    '���뷽ʽ,��ͳ������Ч,ȱʡ�Ҷ���
    cboAlign.AddItem "�����"
    cboAlign.AddItem "�м����"
    cboAlign.AddItem "�Ҷ���"
    cboAlign.ListIndex = 2
    
    cboOper.AddItem "="
    cboOper.AddItem ">"
    cboOper.AddItem ">="
    cboOper.AddItem "<"
    cboOper.AddItem "<="
    cboOper.AddItem ">"
    cboOper.AddItem "LIKE"
    cboOper.AddItem "IN" '������
    cboOper.ListIndex = 0
    
    '��ʼ��
    bytStep = 0
    Caption = Tag & " - " & picGuide(bytStep).Tag
    picGuide(bytStep).ZOrder
    picGuide(bytStep).Visible = True
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboList_Click()
'���ܣ���������,����ѡ���������Դ����ֶ��б�
    Dim rsTmp As ADODB.Recordset
    Dim objItem As Object, i As Integer
    
    lvwItem.ListItems.Clear
    
    If Val(cboList.Tag) = cboList.ListIndex Then Exit Sub
    
    If cboList.ListIndex <> -1 Then
        Set rsTmp = OpenRecord(cboList.Text, cboConn.ItemData(cboConn.ListIndex))
        If Not rsTmp Is Nothing Then
            For i = 0 To rsTmp.Fields.count - 1
                If Not IsType(rsTmp.Fields(i).type, adLongVarBinary) Then
                    Set objItem = lvwItem.ListItems.Add(, , rsTmp.Fields(i).name)
                    objItem.SubItems(1) = ""
                    objItem.Tag = rsTmp.Fields(i).type '�����ֶ�����
                    If InStr(objItem.Text, "ID") = 0 Then objItem.Checked = True
                End If
            Next
            lvwItem.SetFocus
        End If
    End If
    
    cboList.Tag = CStr(cboList.ListIndex)
End Sub

Private Sub cmdNext_Click()
    Dim i As Integer, j As Integer, strFlag As String
    
    '��һ��
    Select Case bytStep
        Case 0 'ѡ������
            picGuide(bytStep).Visible = False
            cmdPre.Enabled = True
            If optType(0).Value Then
                bytStep = 1
            Else
                bytStep = 2
            End If
        Case 1 '�������
            '���Ϸ���
            If cboList.ListIndex = -1 Then
                MsgBox "��ѡ�������������Դ��", vbInformation, App.Title
                cboList.SetFocus: Exit Sub
            End If
            For i = 1 To lvwItem.ListItems.count
                If lvwItem.ListItems(i).Checked Then
                    j = j + 1: Exit For
                End If
            Next
            If j = 0 Then
                MsgBox "û��ѡ�������������Ŀ��", vbInformation, App.Title
                lvwItem.SetFocus: Exit Sub
            End If
            '������ʼ(���������Դ����)
            If lvwIF.Tag <> cboList.List(cboList.ListIndex) Then
                lvwIF.Tag = cboList.List(cboList.ListIndex) '��¼��־(���)
                cboIF.Clear: lvwIF.ListItems.Clear
                cboValue.Text = "": txtName.Text = ""
                For i = 1 To lvwItem.ListItems.count
                    cboIF.AddItem lvwItem.ListItems(i).Text
                    cboIF.ItemData(i - 1) = lvwItem.ListItems(i).Tag '�ֶ�����
                Next
            End If
            
            picGuide(bytStep).Visible = False
            bytStep = 4
        Case 2 '��������
            If cboMain.ListIndex = -1 Then
                MsgBox "��ѡ����������Դ��", vbInformation, App.Title: cboMain.SetFocus: Exit Sub
            End If
            If cboKey.ListCount = 0 Then
                MsgBox "�޷�����������Դ����,���ܽ�����һ����", vbInformation, App.Title: cboMain.SetFocus: Exit Sub
            End If
            
            strFlag = "," & cboMain.Text
            For i = 1 To lvwSub.ListItems.count
                strFlag = strFlag & "," & lvwSub.ListItems(i).Text
            Next
            strFlag = Mid(strFlag, 2)
            '���¼���,����Ѹ���������
            If lstAll.Tag <> strFlag Then
                lstAll.Clear: lvwVsc.ListItems.Clear
                lvwHsc.ListItems.Clear: lvwState.ListItems.Clear
                lstAll.Tag = strFlag
                '�����ѡ�ֶ�(������)
                For i = 0 To cboKey.ListCount - 1
                    lstAll.AddItem cboMain.Text & "." & cboKey.List(i)
                    lstAll.ItemData(lstAll.ListCount - 1) = cboKey.ItemData(i)
                Next
                For i = 1 To lvwSub.ListItems.count
                    With lvwSub.ListItems(i)
                        For j = 0 To UBound(Split(.Tag, "|"))
                            lstAll.AddItem .Text & "." & Split(Split(.Tag, "|")(j), ",")(0)
                            lstAll.ItemData(lstAll.ListCount - 1) = Split(Split(.Tag, "|")(j), ",")(1)
                        Next
                    End With
                Next
                lstAll.ListIndex = 0
            End If
            
            picGuide(bytStep).Visible = False
            bytStep = 3
        Case 3 '�����Ų�
            If lvwVsc.ListItems.count = 0 Then
                MsgBox "������ѡ��һ�����������Ŀ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
            End If
            If lvwState.ListItems.count = 0 Then
                MsgBox "������ѡ��һ��ͳ����Ŀ��", vbInformation, App.Title: lstAll.SetFocus: Exit Sub
            End If
            
            '������ʼ(���������Դ����)
            strFlag = "," & cboMain.Text
            For i = 1 To lvwSub.ListItems.count
                strFlag = strFlag & "," & lvwSub.ListItems(i).Text
            Next
            strFlag = Mid(strFlag, 2)
            If lvwIF.Tag <> strFlag Then
                lvwIF.Tag = strFlag '��¼��־(����)
                cboIF.Clear: lvwIF.ListItems.Clear
                cboValue.Text = "": txtName.Text = ""
                
                For i = 0 To lstAll.ListCount - 1
                    cboIF.AddItem lstAll.List(i)
                    cboIF.ItemData(i) = lstAll.ItemData(i)  '�ֶ�����
                Next
            End If
            
            picGuide(bytStep).Visible = False
            bytStep = 4
        Case 4 '��������
            picGuide(bytStep).Visible = False
            cmdNext.Enabled = False
            cmdOK.Enabled = True
            bytStep = 5
            
            If txtNO.Visible Then txtNO.Text = GetNextNO
            
            If optType(0).Value Then
                txtTitle.Text = GetName(cboList.Text, 2) & "���"
            Else
                txtTitle.Text = GetName(cboMain.Text, 1) & "����"
            End If
    End Select
    If bytStep <> 5 Then
        Caption = Tag & " - " & picGuide(bytStep).Tag
        picGuide(bytStep).ZOrder
        picGuide(bytStep).Visible = True
        Me.Refresh
        picGuide(bytStep).SetFocus
        SendKeys "{Tab}": SendKeys "{Tab}"
    Else
        Caption = Tag & " - ���"
    End If
End Sub

Private Sub cmdOK_Click()
    Dim msgR As Integer, i As Long, j As Long, intMax As Integer
    Dim tmpPar As RPTPar, tmpData As RPTData, objPars As RPTPars, tmpItem As RPTItem, tmpID As RelatID
    Dim strSQL As String, strOrder As String, strWhere As String, strGroup As String, strFields As String

    Set objGuide = New Report
    
    If Not blnNew Then
        If TLen(txtTitle.Text) > 255 Then '���Բ�д����
            MsgBox "������ⳤ�Ȳ��ܳ���255���ַ���", vbInformation, App.Title: txtTitle.SetFocus: Exit Sub
        End If
        msgR = MsgBox("Ҫ������������е�������", vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title)
        If msgR = vbCancel Then Exit Sub
    Else
        If txtNO.Text = "" Then MsgBox "�����뱨���ţ�", vbInformation, App.Title: txtNO.SetFocus: Exit Sub
        If txtTitle.Text = "" Then MsgBox "�����뱨����⣡", vbInformation, App.Title: txtTitle.SetFocus: Exit Sub
        
        If Not CheckLen(txtNO, 20, "���") Then Exit Sub
        If Not CheckLen(txtTitle, 40, "����") Then Exit Sub
        If Not CheckLen(txtNote, 255, "˵��") Then Exit Sub
        
        If CheckExist("zlReports", "���", txtNO.Text) Then
            MsgBox "�ñ���Ѿ�����������ʹ��,���������룡", vbInformation, App.Title
            txtNO.SetFocus: Exit Sub
        End If
    End If
    
    intMax = 0
    If msgR = vbYes Then
        objGuide.��� = "" '�Դ���Ϊ������
    Else
        'SQL�������
        If lvwIF.ListItems.count > 0 Then
            For Each tmpData In objReport.Datas
                For Each tmpPar In tmpData.Pars
                    For i = 1 To lvwIF.ListItems.count
                        If lvwIF.ListItems(i).SubItems(2) = "IN" Then
                            msgR = 3
                        Else
                            Select Case lvwIF.ListItems(i).Tag
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    msgR = 0
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                                        , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    msgR = 1
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    msgR = 2
                                Case Else
                                    msgR = 0
                            End Select
                        End If
                        If tmpPar.���� = lvwIF.ListItems(i).Text _
                            And (tmpPar.ȱʡֵ <> lvwIF.ListItems(i).SubItems(3) Or tmpPar.���� <> msgR) Then
                            MsgBox "�ڱ������е�����Դ�з�����������[" & tmpPar.���� & "]ͬ��,��ȱʡֵ�����Ͳ���ͬ.���������ø�������" _
                                , vbInformation, App.Title
                            Exit Sub
                        End If
                    Next
                Next
            Next
        End If
        
        '��������,���������Ҫ��ͷ
        For Each tmpItem In objReport.Items
            If tmpItem.id > intMax Then intMax = tmpItem.id
            'ע�����
            For Each tmpID In tmpItem.CopyIDs
                If tmpID.id > intMax Then intMax = tmpID.id
            Next
        Next
        objGuide.��� = "NOT"
    End If
    
    
    '�����򵼲����ı���
    
    '����
    If txtTitle.Text <> "" Then
        intMax = intMax + 1
        '###��ʽ
        '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
        objGuide.Items.Add intMax, 1, "Ԫ��" & intMax, 0, 2, 0, "", 0, txtTitle.Text, "", _
            (mobjFmt.W - picGuide(0).TextWidth(txtTitle.Text)) / 2, _
            picGuide(0).TextHeight(txtTitle.Text), picGuide(0).TextWidth(txtTitle.Text), _
            picGuide(0).TextHeight(txtTitle.Text), 0, 1, False, picGuide(0).Font.name, _
            picGuide(0).Font.Size, True, False, False, 0, 0, &HFFFFFF, False, 0, 0, 0, 0, _
            False, False, , False, , , , "_" & intMax
    
        '����ȱʡ��ʽ
        objGuide.Fmts.Add 1, txtTitle.Text, mobjFmt.W, mobjFmt.H, mobjFmt.ֽ��, mobjFmt.ֽ��, mobjFmt.��ֽ̬��, mobjFmt.ͼ��, "_" & 1
    Else
        
        '����ȱʡ��ʽ
        objGuide.Fmts.Add 1, "ȱʡ��ʽ", mobjFmt.W, mobjFmt.H, mobjFmt.ֽ��, mobjFmt.ֽ��, mobjFmt.��ֽ̬��, mobjFmt.ͼ��, "_" & 1
    End If
    
    If optType(0).Value Then
        '��������
        '1.����Դ
        For i = 1 To lvwItem.ListItems.count
            If lvwItem.ListItems(i).Checked Then
                strSQL = strSQL & "," & lvwItem.ListItems(i).Text
                strFields = strFields & "|" & lvwItem.ListItems(i).Text & "," & lvwItem.ListItems(i).Tag
            End If
            If lvwItem.ListItems(i).SubItems(1) <> "" Then
                strOrder = strOrder & "," & lvwItem.ListItems(i).Text
                If lvwItem.ListItems(i).SubItems(1) = "��" Then strOrder = strOrder & " Desc"
            End If
        Next
        If strOrder <> "" Then strOrder = " Order by " & Mid(strOrder, 2)
        If strFields <> "" Then strFields = Mid(strFields, 2)
        
        Set objPars = New RPTPars
        For i = 1 To lvwIF.ListItems.count
            '��������
            If lvwIF.ListItems(i).SubItems(2) = "IN" Then
                strWhere = strWhere & " And " & lvwIF.ListItems(i).SubItems(1) & " " & _
                    lvwIF.ListItems(i).SubItems(2) & " ([" & i - 1 & "])"
                msgR = 3
            Else
                If lvwIF.ListItems(i).Tag = adDBTimeStamp Then
                    If InStr(lvwIF.ListItems(i).SubItems(3), "ʱ��") = 0 And InStr(lvwIF.ListItems(i).SubItems(3), ":") = 0 Then
                        strWhere = strWhere & " And Trunc(" & lvwIF.ListItems(i).SubItems(1) & ") " & _
                            lvwIF.ListItems(i).SubItems(2) & " [" & i - 1 & "]"
                    Else
                        strWhere = strWhere & " And " & lvwIF.ListItems(i).SubItems(1) & " " & _
                            lvwIF.ListItems(i).SubItems(2) & " [" & i - 1 & "]"
                    End If
                Else
                    strWhere = strWhere & " And " & lvwIF.ListItems(i).SubItems(1) & " " & _
                        lvwIF.ListItems(i).SubItems(2) & " [" & i - 1 & "]"
                End If
                Select Case lvwIF.ListItems(i).Tag
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                        msgR = 0
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt _
                        , adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                        msgR = 1
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        msgR = 2
                    Case Else
                        msgR = 0
                End Select
            End If
            '###��ʽ����
            objPars.Add "", i - 1, lvwIF.ListItems(i).Text, CByte(msgR), lvwIF.ListItems(i).SubItems(3), 1 _
                , "", "", "", "", "", "", "_" & i - 1
        Next
        If strWhere <> "" Then strWhere = " Where " & Mid(strWhere, 6)
        
        strSQL = "Select " & Mid(strSQL, 2) & " From " & GetName(cboList.Text, 2) & strWhere & strOrder
        
        strWhere = GetName(cboList.Text, 2) 'strWhere��������Դ����
        If objGuide.��� <> "" Then
            For Each tmpData In objReport.Datas
                If tmpData.���� = strWhere Then
                    strWhere = strWhere & "_" '��֤��������ͬ��
                    Exit For
                End If
            Next
        End If
        
        '����������Դ
        objGuide.Datas.Add strWhere, cboConn.ItemData(cboConn.ListIndex), strSQL, strFields, cboList.Text _
                                , 0, "", objPars, "_" & strWhere
        
        '2.���
        intMax = intMax + 1
        '���
        i = IIF((UBound(Split(strFields, "|")) + 1) * 1000 + 300 > mobjFmt.W - 600 _
            , mobjFmt.W - 600 _
            , (UBound(Split(strFields, "|")) + 1) * 1000 + 300)
        '###��ʽ
        '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
        Set tmpItem = objGuide.Items.Add(intMax, 1, "Ԫ��" & intMax, 0, 4, 0, "", 0, strWhere, "", _
            (mobjFmt.W - i) / 2, picGuide(0).TextHeight(txtTitle.Text) * 3, i, _
            mobjFmt.H - picGuide(0).TextHeight(txtTitle.Text) * 4, 255, 0, False, "����", 9, _
            False, False, False, 0, 0, &HFFFFFF, True, 1, "", "", "", False, False, , False, , , , "_" & intMax)
        
        msgR = 0
        For i = 1 To lvwItem.ListItems.count
            If lvwItem.ListItems(i).Checked Then
                Select Case lvwItem.ListItems(i).Tag
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                        , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                        j = 2 '�Ҷ���
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        j = 1 '�ж���
                    Case Else
                        j = 0 '�����
                End Select
                intMax = intMax + 1
                '###��ʽ
                '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
                objGuide.Items.Add intMax, 1, "Ԫ��" & intMax, tmpItem.id, 6, msgR, "", 0, _
                    "[" & strWhere & "." & lvwItem.ListItems(i).Text & "]", "4^255^" & lvwItem.ListItems(i).Text, _
                    0, 0, 1000, 0, 0, CByte(j), False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intMax
                msgR = msgR + 1
                objGuide.Items("_" & tmpItem.id).SubIDs.Add intMax, "_" & intMax
            End If
        Next
    Else
        '���������������ܱ�
        '1:SQL������Դ

        For i = 1 To lvwVsc.ListItems.count
            With lvwVsc.ListItems(i)
                strSQL = strSQL & "," & .Text & IIF(GetName(.Text, 2) = .SubItems(1), "", " AS " & .SubItems(1))
                strGroup = strGroup & "," & .Text
                strFields = strFields & "|" & .SubItems(1) & "," & .Tag
            End With
        Next
        For i = 1 To lvwHsc.ListItems.count
            With lvwHsc.ListItems(i)
                strSQL = strSQL & "," & .Text & IIF(GetName(.Text, 2) = .SubItems(1), "", " AS " & .SubItems(1))
                strGroup = strGroup & "," & .Text
                strFields = strFields & "|" & .SubItems(1) & "," & .Tag
            End With
        Next
        strGroup = " Group by " & Mid(strGroup, 2)
        
        For i = 1 To lvwState.ListItems.count
            With lvwState.ListItems(i)
                Select Case .SubItems(2)
                    Case "���"
                        strSQL = strSQL & ",SUM(" & .Text & ") AS " & .SubItems(1)
                    Case "��ƽ��ֵ"
                        strSQL = strSQL & ",AVG(" & .Text & ") AS " & .SubItems(1)
                    Case "�����ֵ"
                        strSQL = strSQL & ",MAX(" & .Text & ") AS " & .SubItems(1)
                    Case "����Сֵ"
                        strSQL = strSQL & ",MIN(" & .Text & ") AS " & .SubItems(1)
                    Case "���¼��"
                        strSQL = strSQL & ",COUNT(" & .Text & ") AS " & .SubItems(1)
                End Select
                strFields = strFields & "|" & .SubItems(1) & "," & .Tag
            End With
        Next
        
        strFields = Mid(strFields, 2)
        strSQL = "Select " & Mid(strSQL, 2) & " From " & cboMain.Text & ","
        
        For i = 1 To lvwSub.ListItems.count
            With lvwSub.ListItems(i)
                '���ܶ����
                If InStr(strSQL, .Text & ",") = 0 Then
                    strSQL = strSQL & .Text & ","
                End If
                strWhere = strWhere & " And " & .Text & "." & .SubItems(1) & " = " & cboMain.Text & "." & .SubItems(2)
            End With
        Next
        Set objPars = New RPTPars
        For i = 1 To lvwIF.ListItems.count
            With lvwIF.ListItems(i)
                '��������
                If .SubItems(2) = "IN" Then
                    strWhere = strWhere & " And " & .SubItems(1) & " " & .SubItems(2) & " ([" & i - 1 & "])"
                    msgR = 3
                Else
                    If .Tag = adDBTimeStamp Then
                        If InStr(.SubItems(3), "ʱ��") = 0 And InStr(.SubItems(3), ":") = 0 Then
                            strWhere = strWhere & " And Trunc(" & .SubItems(1) & ") " & .SubItems(2) & " [" & i - 1 & "]"
                        Else
                            strWhere = strWhere & " And " & .SubItems(1) & " " & .SubItems(2) & " [" & i - 1 & "]"
                        End If
                    Else
                        strWhere = strWhere & " And " & .SubItems(1) & " " & .SubItems(2) & " [" & i - 1 & "]"
                    End If
                    Select Case .Tag
                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                            msgR = 0
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt _
                            , adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt _
                            , adUnsignedTinyInt
                            msgR = 1
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                            msgR = 2
                        Case Else
                            msgR = 0
                    End Select
                End If
                '###��ʽ����
                objPars.Add "", i - 1, .Text, CByte(msgR), .SubItems(3), 1, "", "", "", "", "", "", "_" & i - 1
            End With
        Next
        If strWhere <> "" Then strWhere = " Where " & Mid(strWhere, 6)
        
        strSQL = Left(strSQL, Len(strSQL) - 1) & strWhere & strGroup
        
        '�򻯴���SQL������ȥ��������
        msgR = 65 '������"A"��ʼ
        strSQL = Replace(strSQL, cboMain.Text & ".", Chr(msgR) & ".") 'ע���Ⱥ�
        strSQL = Replace(strSQL, cboMain.Text, GetName(cboMain.Text, 2) & " " & Chr(msgR))
        For i = 1 To lvwSub.ListItems.count
            msgR = msgR + 1
            With lvwSub.ListItems(i)
                strSQL = Replace(strSQL, .Text & ".", Chr(msgR) & ".") 'ע���Ⱥ�
                strSQL = Replace(strSQL, .Text, GetName(.Text, 2) & " " & Chr(msgR))
            End With
        Next
        
        '����Դ
        strWhere = GetName(cboMain.Text, 2) 'strWhere��������Դ����
        If objGuide.��� <> "" Then
            For Each tmpData In objReport.Datas
                If tmpData.���� = strWhere Then
                    strWhere = strWhere & "_" '��֤��������ͬ��
                    Exit For
                End If
            Next
        End If
        strGroup = "," & cboMain.Text 'strGroup��������
        For i = 1 To lvwSub.ListItems.count
            strGroup = strGroup & "," & lvwSub.ListItems(i).Text
        Next
        strGroup = Mid(strGroup, 2)
        
        objGuide.Datas.Add strWhere, cboMainConn.ItemData(cboMainConn.ListIndex), strSQL, strFields, strGroup, 1, "" _
                , objPars, "_" & strWhere
        
        '2:���
        
        '������
        intMax = intMax + 1
        '###��ʽ
        '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
        Set tmpItem = objGuide.Items.Add(intMax, 1, "Ԫ��" & intMax, 0, 5, 0, "", 0, strWhere, "", 300, _
            picGuide(0).TextHeight(txtTitle.Text) * 3, mobjFmt.W - 600, _
            mobjFmt.H - picGuide(0).TextHeight(txtTitle.Text) * 4, 255, 0, False, "����", 9, False, False, _
            False, 0, 0, &HFFFFFF, True, 1, "", "", "", False, False, , False, , , , "_" & intMax)
        
        '���������
        For i = 1 To lvwVsc.ListItems.count
            With lvwVsc.ListItems(i)
                intMax = intMax + 1
                
                Select Case .SubItems(2) 'strFields�������ܷ�ʽ
                    Case ""
                        strFields = ""
                    Case "���"
                        strFields = "SUM"
                    Case "��ƽ��ֵ"
                        strFields = "AVG"
                    Case "�����ֵ"
                        strFields = "MAX"
                    Case "����Сֵ"
                        strFields = "MIN"
                    Case "���¼��"
                        strFields = "COUNT"
                End Select
                '###��ʽ
                '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
                objGuide.Items.Add intMax, 1, "Ԫ��" & intMax, tmpItem.id, 7, i - 1, "", 0, .SubItems(1), _
                    "", 0, 0, 1000, 0, 0, 0, False, "", 0, False, False, False, 0, 0, 0, False, 0, _
                    .SubItems(1), "", strFields, False, False, , False, , , , "_" & intMax
                
                objGuide.Items("_" & tmpItem.id).SubIDs.Add intMax, "_" & intMax
            End With
        Next
        
        '���������
        For i = 1 To lvwHsc.ListItems.count
            With lvwHsc.ListItems(i)
                intMax = intMax + 1
                
                Select Case .SubItems(2) 'strFields�������ܷ�ʽ
                    Case ""
                        strFields = ""
                    Case "���"
                        strFields = "SUM"
                    Case "��ƽ��ֵ"
                        strFields = "AVG"
                    Case "�����ֵ"
                        strFields = "MAX"
                    Case "����Сֵ"
                        strFields = "MIN"
                    Case "���¼��"
                        strFields = "COUNT"
                End Select
                '###��ʽ
                '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
                objGuide.Items.Add intMax, 1, "Ԫ��" & intMax, tmpItem.id, 8, i - 1, "", 0, .SubItems(1), _
                    "", 0, 0, 1000, 0, 0, 0, False, "", 0, False, False, False, 0, 0, 0, False, 0, _
                    .SubItems(1), "", strFields, False, False, , False, , , , "_" & intMax
                
                objGuide.Items("_" & tmpItem.id).SubIDs.Add intMax, "_" & intMax
            End With
        Next
        
        'ͳ����
        For i = 1 To lvwState.ListItems.count
            With lvwState.ListItems(i)
                intMax = intMax + 1
                
                '����
                Select Case .SubItems(3)
                    Case "�����"
                        msgR = 0
                    Case "�м����"
                        msgR = 1
                    Case "�Ҷ���"
                        msgR = 2
                End Select
                '###��ʽ
                '###���ơ�ͼƬ�����ա����ʡ��Ե����ԵĴ���
                objGuide.Items.Add intMax, 1, "Ԫ��" & intMax, tmpItem.id, 9, i - 1, "", 0, .SubItems(1), _
                    "", 0, 0, 1000, 0, 0, CByte(msgR), False, "����", 9, False, False, False, 0, 0, 0, False, 0, "", _
                    IIF(InStr(lvwState.ListItems(i).SubItems(2), "�ϼ�ƽ��") > 0, "0.00", ""), _
                    "", False, False, , False, , , , "_" & intMax
                
                objGuide.Items("_" & tmpItem.id).SubIDs.Add intMax, "_" & intMax
            End With
        Next
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub txtNO_GotFocus()
    SelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNote_GotFocus()
    SelAll txtNote
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function GetItemRS(ByVal colRS As Collection, ByVal intIndex As Integer) As Recordset
    Dim i As Integer
    Dim rsTmp As Recordset
    
    If colRS Is Nothing Then Exit Function
    
    For i = 1 To colRS.count
        If i = intIndex Then
            On Error Resume Next
            Set rsTmp = colRS("_" & i)
            On Error GoTo 0
            If Not rsTmp Is Nothing Then
                Set GetItemRS = colRS(i)
            End If
            Exit For
        End If
    Next
End Function

Private Sub SetSourceControl(ByVal vConnect As ComboBox, ByRef vSource As ComboBox)
'���ܣ���ָ����������ˢ�����ݱ��ComBox�ؼ�
'������
'  vConnect����������
'  vSource������Դ

    Dim rsTemp As ADODB.Recordset
    Dim intConnect As Integer

    vSource.Clear
    If vConnect.ListIndex <= 0 Then
        '��ǰ��¼
        Set rsTemp = grsObject
    Else
        '������¼
        intConnect = vConnect.ItemData(vConnect.ListIndex)
        If mcolRS.count <= 0 Then
            GoTo makAdd
        ElseIf GetItemRS(mcolRS, intConnect) Is Nothing Then
makAdd:
            Set rsTemp = mdlPublic.UserObject(intConnect, True)
            If Not rsTemp Is Nothing Then
                mcolRS.Add rsTemp, "_" & intConnect
            End If
        Else
            Set rsTemp = mcolRS("_" & intConnect)
        End If
    End If
    
    If Not rsTemp Is Nothing Then
        With rsTemp
            If rsTemp.State = adStateOpen Then
                .Filter = MSTR_OWNER_FILTER
                Do While .EOF = False
                    vSource.AddItem rsTemp!Owner & "." & rsTemp!OBJECT_NAME
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub
