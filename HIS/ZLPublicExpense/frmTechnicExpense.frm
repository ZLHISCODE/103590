VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTechnicExpense 
   AutoRedraw      =   -1  'True
   Caption         =   "���˼ƷѴ���"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTechnicExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelWholeSet 
      Caption         =   "����(&T)"
      Height          =   375
      Left            =   90
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   525
      Width           =   1080
   End
   Begin VB.CommandButton cmdSaveWholeSet 
      Caption         =   "����Ϊ�����շ���Ŀ(&W)"
      Height          =   375
      Left            =   1215
      TabIndex        =   53
      Top             =   525
      Width           =   2715
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   7872
      Width           =   11856
      _ExtentX        =   20902
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmTechnicExpense.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14684
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "�������"
            Object.ToolTipText     =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "ҽ������"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   952
            MinWidth        =   952
            Picture         =   "frmTechnicExpense.frx":0E1E
            Key             =   "BarCode"
            Object.Tag             =   "BarCode"
            Object.ToolTipText     =   "��ʾ���������"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":1548
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":1B82
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   11850
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5004
      Width           =   11850
      Begin VB.CommandButton cmdSaveAndPay 
         BackColor       =   &H00C0C0C0&
         Caption         =   "���֧��(&P)"
         Height          =   420
         Left            =   8730
         TabIndex        =   21
         ToolTipText     =   "�ȼ���Shift+F2"
         Top             =   1740
         Width           =   1680
      End
      Begin VB.ComboBox cboTemp 
         Height          =   360
         Left            =   7320
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   4000
         Width           =   1695
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   11070
         Top             =   1980
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicExpense.frx":21BC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   8730
         TabIndex        =   22
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   2250
         Width           =   1680
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   8730
         TabIndex        =   20
         ToolTipText     =   "�ȼ���F2"
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   37
         ToolTipText     =   "���:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�������"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk�Ӱ� 
            Caption         =   "�Ӱ�(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6555
            TabIndex        =   16
            Top             =   180
            Width           =   2085
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9360
            TabIndex        =   17
            Top             =   180
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            HideSelection   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBaby 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ӥ����(&B)"
            Height          =   240
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   240
            Left            =   5790
            TabIndex        =   39
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   240
            Left            =   8820
            TabIndex        =   38
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraDrawDept 
         Height          =   720
         Left            =   0
         TabIndex        =   48
         Top             =   360
         Width           =   13575
         Begin VB.TextBox txt���˱�ע 
            BackColor       =   &H00E0E0E0&
            Height          =   360
            Left            =   1095
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   2700
         End
         Begin VB.Label lbl���˱�ע 
            Caption         =   "���˱�ע"
            Height          =   225
            Left            =   105
            TabIndex        =   50
            Top             =   315
            Width           =   1005
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         RowHeightMin    =   320
         BackColorBkg    =   15466495
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "^         ��Ŀ|^          ���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraStat 
         Height          =   1770
         Left            =   3510
         TabIndex        =   40
         Top             =   1065
         Width           =   3675
         Begin VB.TextBox txtPreNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1230
            Width           =   1845
         End
         Begin VB.TextBox txtʵ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   750
            Width           =   1845
         End
         Begin VB.TextBox txtӦ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   250
            Width           =   1845
         End
         Begin VB.Label lblPreNO 
            AutoSize        =   -1  'True
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   270
            TabIndex        =   52
            Top             =   1298
            Width           =   690
         End
         Begin VB.Label lblʵ�� 
            AutoSize        =   -1  'True
            Caption         =   "ʵ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   270
            TabIndex        =   42
            Top             =   818
            Width           =   690
         End
         Begin VB.Label lblӦ�� 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   270
            TabIndex        =   41
            Top             =   318
            Width           =   690
         End
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   30
      TabIndex        =   25
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   11865
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   1425
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   18000
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   18000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "���˼Ʒѵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   225
         TabIndex        =   29
         ToolTipText     =   "���:F6"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9540
         TabIndex        =   26
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   9375
      TabIndex        =   24
      Top             =   855
      Width           =   2505
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "cbo��������"
         Top             =   615
         Width           =   2265
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   23
      Top             =   855
      Width           =   9345
      Begin VB.TextBox txtסԺ�� 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1290
      End
      Begin VB.TextBox txt�ѱ� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txt���ʽ 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   615
         Width           =   2085
      End
      Begin VB.TextBox txt�Ա� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   1290
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   765
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   240
         Left            =   7140
         TabIndex        =   47
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   7140
         TabIndex        =   46
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   5145
         TabIndex        =   45
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl���ʽ 
         Caption         =   "���� ��ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   44
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   5385
         TabIndex        =   35
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   33
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   2370
         TabIndex        =   32
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Left            =   3705
         TabIndex        =   31
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2520
      Left            =   15
      TabIndex        =   11
      Top             =   2505
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   4445
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fraBarCode 
      Height          =   630
      Left            =   30
      TabIndex        =   56
      Top             =   1800
      Width           =   11850
      Begin VB.TextBox txtBarCode 
         Height          =   360
         Left            =   705
         TabIndex        =   57
         Top             =   195
         Width           =   11040
      End
      Begin VB.Label lblBarCode 
         Caption         =   "����"
         Height          =   300
         Left            =   150
         TabIndex        =   58
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ�:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmTechnicExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'��ڲ���
'����������������������������������������������������������������������������������������������������������������������������������������
Private mlngҽ��ID As Long '��������ʱ��
Private mlng���ͺ� As Long '��������ʱ��
Private mlng����ID As Long 'ȷ��Ҫ�ƷѵĲ���ID
Private mlng��ҳID As Long 'ȷ��Ҫ�Ʒѵ���ҳID
Private mint������Դ As Integer '1-���ﲡ��,2-סԺ����
Private mint��¼���� As Integer '1-�շ�(����),2-����(��/ס)
Private mstrFeeTab As String
Private mbln���õǼ� As Boolean '���Ǽ�,����ʵ�ս��
Private mlng��������ID As Long 'Ϊ��ǰ������ҽ������
Private mlng���˿���id As Long '��Ҫ������ȷ�����ﲡ�˵Ŀ���ID

Private mlng��������ID As Long
Private mstr����ҽ�� As String

Private mbytInState As Byte '0-ִ��,1-����,2-����(��֧��),3-ɾ��
Private mstrInNO As String '�������ĵ��ݺ�(ִ��ʱΪ�޸�)
Private mstrOriginalNO As String '����������ʱ,ҽ�������еĵ��ݺ�

Private mstrTime As String '�����������ݵĵǼ�ʱ��
Private mblnDelete As Boolean '�Ƿ����˷ѵ���(����)
Private mbytӦ�ó��� As Byte
Private mobjSquareCard As Object
Private mblnSetControl As Boolean
'����������������������������������������������������������������������������������������������������������������������������������������
Private mstrSuccesSaveNos As String
Private mblnWarnCloseed As Boolean  '���˺�:����ñ��������Ĺر�
Private mblnSendMateria  As Boolean
Private mbytSendMateria As Byte '0-���ʺ󲻷�ҩ,1-�Զ���ҩ,2-��ʾ��ҩ
Private mbytȱʡ���� As Byte    '0-ҽ������;1-���˿���
Private mobjBaseItem As Object
Private mstrסԺҽ�� As String
Private mrsAll�������� As ADODB.Recordset
Private mlng���� As Long
Private Enum BillColType       '���ݿؼ���������
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum
Private Enum BillCol
    �� = 0
    ��� = 1
    ��Ŀ = 2
    ��Ʒ�� = 3
    ��� = 4
    ��λ = 5
    ���� = 6
    ���� = 7
    ���� = 8
    Ӧ�ս�� = 9
    ʵ�ս�� = 10
    ִ�п��� = 11
    ��־ = 12
    ���� = 13
End Enum

Private Enum Pan
    C2��ʾ��Ϣ = 2
End Enum

Private mstrRegNo As String '���ﲡ�˹Һŵ���
'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ʵʱ��� As Boolean
    ҽ��ȷ���������� As Boolean 'Ŀǰֻ�б���ҽ��ר��
End Type
Private MCPAR As TYPE_MedicarePAR
Private mrsDept As ADODB.Recordset
Private mstrPrivs As String
'ҽ������վ���ط��ò���
'����������������������������������������������������������������������������������������������������������������������������������������
Private mstrLike As String '����ƥ�䷽ʽ
Private mblnPay As Boolean '��ҩ�Ƿ����븶��
Private mblnTime As Boolean '����Ƿ����븶��
Private mbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ�����
Private mbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ����
Private mstr�շ���� As String '��������շ����
Private mblnҩ����λ As Boolean '�Ƿ������ﵥλ��סԺ��λ��ʾҩƷ
Private mstrҩ����λ As String '���ݲ�����Դ������"���ﵥλ"��"סԺ��λ"
Private mstrҩ����װ As String '���ݲ�����Դ������"�����װ"��"סԺ��װ"
Private mlngPreRow As Long '��¼��ǰ��,�����ı���ʱ
Private mlng��ҩ�� As Long, mlng��ҩ�� As Long, mlng��ҩ�� As Long, mlng���ϲ��� As Long
'����������������������������������������������������������������������������������������������������������������������������������������
'���ݶ���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsWork As New ADODB.Recordset '�����ϰ��ҩ��
Private mblnWork As Boolean '��ǰ�Ƿ��������ϰ��ҩ��
Private mlngҩƷ���ID As Long '��ǰ���ݲ�����ҩƷ������ID
Private mlng�������ID As Long '��ǰ���ݲ���������������ID
'�������
Private mobjBill As ExpenseBill '���õ��ݶ���
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes '������Ŀ���ܼ���

'�������
Private mbytWarn As Byte '���ʱ����ķ���ֵ
Private mintWarn As Integer '���ʱ�����ʾ�ļ���ѡ��
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mrsWarn As New ADODB.Recordset '����������
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀ�ĳ����鷽ʽ

Private WithEvents mobjBrushCheck As clsBrushCardInput
Attribute mobjBrushCheck.VB_VarHelpID = -1
Private mobjCard As New clsBrushCard
Private mbln����ˢ�� As Boolean

Private mcurModiMoney As Currency '�޸ĵ���ʱԭ���ݵĽ��
Private mblnDrop As Boolean '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnNewRow As Boolean
Private mobjSaveData As Object
Private mblnOne As Boolean '�Ƿ�ֻ��һ�������շ����
Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private mdblItemNum As Double '���ݿ��е�ǰ�����Ŀ������
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����
Private marrDr() As String '��¼ҽ����"ID|����ID|���|����|����"
Private mblnEnterCell As Boolean '�����Ƿ�ִ��Entercell�¼�
Private Const STR_HEAD = "��,450,4;���,750,1;��Ŀ,2175,1;��Ʒ��,2000,1;���,1105,1;��λ,520,4;����,520,1;����,570,1;����,1055,7;" & _
                        "Ӧ�ս��,1030,7;ʵ�ս��,1080,7;ִ�п���,1255,1;��־,520,4;����,520,1"
Private mblnOK As Boolean
Private mrsMainOperation As ADODB.Recordset
Private mbln���֧�� As Boolean
Private mobjPublicDrug As Object 'ҩƷ��������,105872
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String

Private mblnDrugMachine As Boolean
Private mobjDrugMachine As Object '�·�ҩ��
Private mblnShowBarCode As Boolean '��ʾ���������

Public Function EditCard(ByVal frmMain As Object, ByVal strPrivs As String, _
    ByVal bytInState As Byte, _
    ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, _
    ByVal lng����ID As Long, ByVal lng��ҳId As Long, int������Դ As Integer, _
    ByVal int��¼���� As Integer, ByVal lng��������id As Long, ByVal lng���˿���ID As Long, _
    ByVal lng��������ID As Long, ByVal str����ҽ�� As String, Optional ByVal strOriginalNO As String, _
    Optional ByVal strInNO As String = "", Optional ByVal strTime As String = "", Optional blnDelete As Boolean, _
    Optional strFeeTable As String = "", Optional bln���õǼ� As Boolean, _
    Optional ByRef strOutNos As String, Optional bytӦ�ó��� As Byte = 0, Optional ByRef objSaveData As Object, _
    Optional objSquareCard As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѳ������
    '���: bytInState:0-ִ��,1-����,2-����(��֧��),3-ɾ��
    '      strInNO:�������ĵ��ݺ�(ִ��ʱΪ�޸�)
    '      strOriginalNO:����������ʱ,ҽ�������еĵ��ݺ�
    '     lngҽ��ID:��������ʱ��
    '     lng���ͺ�:��������ʱ��
    '     lng����ID:ȷ��Ҫ�ƷѵĲ���ID
    '     lng��ҳID:ȷ��Ҫ�Ʒѵ���ҳID
    '     int������Դ:1-���ﲡ��,2-סԺ����
    '     int��¼����:1-�շ�(����),2-����(��/ס)
    '     bln���õǼ�:���Ǽ�,����ʵ�ս��
    '     lng��������ID:Ϊ��ǰ������ҽ������
    '     lng���˿���id:��Ҫ������ȷ�����ﲡ�˵Ŀ���ID
    '     lng��������ID:
    '     str����ҽ��:
    '     strTime:�����������ݵĵǼ�ʱ��
    '     blnDelete:�Ƿ����˷ѵ���(����)
    '     strFeeTab:������ü�¼��סԺ���ü�¼
    '     bytӦ�ó���-0-ҽ������;1-��첹��(��ѡ����)
    '����:strOutNos-�ɹ�����ĵ��ݺ�
    '����:�ɹ�����һ�ŵ�������,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 14:13:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
     mbytInState = bytInState: mstrInNO = strInNO: mstrOriginalNO = strOriginalNO
     mstrTime = strTime: mblnDelete = blnDelete
     mlngҽ��ID = lngҽ��ID: mlng���ͺ� = lng���ͺ�: mlng����ID = lng����ID
     mlng��ҳID = lng��ҳId: mint������Դ = int������Դ: mint��¼���� = int��¼����
     mstrFeeTab = strFeeTable: mbln���õǼ� = bln���õǼ�
     mlng��������ID = lng��������id: mlng���˿���id = lng���˿���ID
     mlng��������ID = lng��������ID: mstr����ҽ�� = str����ҽ��
     mbytӦ�ó��� = bytӦ�ó���
     Set mobjSaveData = objSaveData: Set mobjSquareCard = objSquareCard
     mstrPrivs = strPrivs: mblnOK = False
     
     mstrSuccesSaveNos = ""
     Err = 0: On Error Resume Next
     If frmMain Is Nothing Then
        Me.Show 1
        'Call zlCommFun.ShowChildWindow(Me.hWnd, glngMainHwnd)
     Else
        Me.Show 1, frmMain
     End If

     Err = 0: On Error GoTo 0
     If mstrSuccesSaveNos <> "" Then mstrSuccesSaveNos = Mid(mstrSuccesSaveNos, 2)
     strOutNos = mstrSuccesSaveNos
     EditCard = mblnOK
End Function
Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    '���˺� ����:27378 ����:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "ִ�п���"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "��ҩҩ��"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .�շ����) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, 1150, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, 1150, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then
        mobjBill.��������ID = 0
        cbo��������.Text = ""
    End If
End Sub
Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln��������ۿ� As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Then Cancel = True: Exit Sub
    
    If mobjBill.Details.Count >= Row Then
        '��������Ŀ����ɾ��ȷ��
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("����Ŀ���� " & bytsubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
            If MsgBox("����Ŀ��[" & mobjBill.Details(mobjBill.Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln��������ۿ� = gSysPara.bln��������ۿ�
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If

        If bln��������ۿ� Then lngMainRow = mobjBill.Details(Bill.Row).�������� '����Ǵ���,ɾ��֮ǰ���´���Ĵ�������,���������,����ɾ��,��������
        
        'ɾ������
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        
        '���¼��㲢ˢ��
        If bln��������ۿ� Then
            If CheckItemHaveSub(lngMainRow) Then
                Call Calc��������ʵ��(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) 'ֻ��һ��������,����ȫ����ɾ��ʱ,������ͨ���������
            End If
        End If
        
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ�������ɾ��
        
        mlngPreRow = 0  '��ʾ�иı���
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub ShowStock(strҩƷ As String, dbl��� As Double)
'���ܣ���ʾҩƷ�����ĵĿ��
    If InStr(1, mstrPrivs, "��ʾ���") > 0 Then
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]���ÿ��:" & dbl���
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]" & IIf(dbl��� > 0, "��", "��") & "���."
    End If
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, i As Long
    Dim lngִ�п��� As Long, strִ�п��� As String
    'ҩƷ�����
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                    lngִ�п��� = .ִ�в���ID
                    strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0)  '29680
                        If mblnҩ����λ Then
                            dblStock = dblStock / .Detail.ҩ����װ
                        End If
                        .Detail.��� = dblStock  '��¼��ǰ��ҩƷ���
                        Call ShowStock(.Detail.����, .Detail.���)
                        
                        'ҩ���ı�,ʵ��ҩƷ���¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf .�շ���� = "4" And .Detail.�������� Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0, , , .Detail.����) '29680
                        .Detail.��� = dblStock
                        Call ShowStock(.Detail.����, .Detail.���)
                        
                        '���ϲ��Ÿı�,ʱ���������¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf InStr(",4,5,6,7,", .�շ����) = 0 Then
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                    End If
                    If mint��¼���� = 2 Then
                        If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                            If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), 2, 0, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = strִ�п���
                                .ִ�в���ID = lngִ�п���: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, pҽ�����ѹ���, IIf(mint������Դ = 2, 1, 0), 0, _
                            MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), mint��¼����, IIf(mint��¼���� = 1, 1, 0), Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.cboObj.Text = strִ�п���
                             .ִ�в���ID = lngִ�п���: Exit Sub
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub bill_CellCheck(Row As Long, Col As Long)
'˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Byte
    
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    If mbytInState = 3 Then Exit Sub
    
    '������δ��������Ч
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "F" And mobjBill.Details(i).���ӱ�־ = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    
    If bytTime = 0 Then
        '114157����ǰ�����в������������������ҽ�����Ƿ����������
        If HaveMainOperation(mint������Դ, mlngҽ��ID, mstrInNO) Then bytTime = 1
    End If
    
    If bytTime > 0 Then
        mobjBill.Details(Row).���ӱ�־ = IIf(strCheck = "", 0, 1)
        Call CalcMoneys(Row)
        Call ShowDetails(Row)
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "��ǰҽ�������л�û����Ҫ������������Ӹ���������", vbInformation, gstrSysName
    End If
End Sub

Private Function HaveMainOperation(ByVal int������Դ As Integer, ByVal lngҽ��ID As Long, _
    ByVal strModifyNo As String) As Boolean
    '��ǰҽ��(�����ҽ��)�Ƿ������ҽ������
    '��Σ�
    '   int������Դ 1-���ﲡ��,2-סԺ����
    '   strModifyNo �޸�ʱҪ�ų���ǰ����
    '˵����
    Dim strSql As String
    Dim strTable As String
    Dim strWhere As String
    
    On Error GoTo ErrHandler
    If int������Դ = 2 Then
        strTable = "סԺ���ü�¼"
    Else
        strTable = "������ü�¼"
    End If
    If strModifyNo <> "" Then strWhere = " And c.No <> [2]"
    
    strSql = _
        "Select Nvl(Sum(c.����), 0) As ����" & vbNewLine & _
        "From " & strTable & " C, ����ҽ����¼ B, ����ҽ����¼ A" & vbNewLine & _
        "Where c.ҽ����� = b.Id And (b.���id = a.���id Or b.���id = a.Id Or b.Id = a.���id Or b.Id = a.Id)" & vbNewLine & _
        "      And c.�շ���� = 'F' And Nvl(c.���ӱ�־, 0) = 0 And a.Id = [1]" & strWhere & vbNewLine & _
        "Having Nvl(Sum(c.����), 0) <> 0"
    If mrsMainOperation Is Nothing Then
        Set mrsMainOperation = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID, strModifyNo)
    ElseIf mrsMainOperation.State <> adStateOpen Then
        Set mrsMainOperation = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID, strModifyNo)
    End If
    If mrsMainOperation.RecordCount = 0 Then Exit Function
    HaveMainOperation = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function CheckMainOperation() As Boolean
    '����:����������������(�����������Ҫ����,�����ڸ�������,���ֹ
    '���:
    '����:lngRow-���ظ�����������
    '����:������������û�����븽������,����true,���򷵻�False
    Dim lngCount As Long, lngRow As Long   'ָ����
    Dim i As Long
    
    On Error GoTo ErrHandler
    lngCount = 0
    For i = 1 To mobjBill.Details.Count
        lngCount = 0
        If mobjBill.Details(i).�շ���� = "F" Then
           If mobjBill.Details(i).���ӱ�־ = 0 Then lngCount = 0: Exit For  '������Ҫ����,�򲻼��,ֱ�ӷ���true
           lngCount = lngCount + 1  '��ʾ��������
           If lngRow <= 0 Then lngRow = i
        End If
    Next
    
    If lngCount <> 0 Then
        If HaveMainOperation(mint������Դ, mlngҽ��ID, mstrInNO) Then
            '114157����ǰ�����в������������������ҽ�����Ƿ����������
        Else
            MsgBox "��ǰҽ�������л�û����Ҫ������������Ӹ���������", vbInformation, gstrSysName
            Bill.Row = lngRow '��λ��
            If Bill.Visible Then Bill.SetFocus
            Exit Function
        End If
    End If
    CheckMainOperation = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function SelectIsNurse() As Boolean
'���ܣ��жϵ�ǰ�������Ƿ�ʿ
    Dim str���� As String
    
    If cbo������.ListIndex <> -1 Then
        If cbo������.ItemData(cbo������.ListIndex) = 0 Then Exit Function
        
        If cbo������.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 6 Then
                str���� = Split(marrDr(cbo������.ListIndex), "|")(6)
                SelectIsNurse = str���� = "��ʿ"
            End If
        End If
    End If
End Function

Private Sub bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean
    Dim str��� As String, str��׼��Ŀ As String, int���� As Integer
    
    If gSysPara.bln�շ���� Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str��� = IIf(SelectIsNurse, "'E','M','4'", mstr�շ����)
        End If
    Else
        str��� = IIf(SelectIsNurse, "'E','M','4'", mstr�շ����)
    End If
    If Not IsNull(mrsInfo!����) Then
        int���� = mrsInfo!����
        '���˺�:24862
        'mint������Դ As Integer '1-���ﲡ��,2-סԺ����
        'mint��¼���� As Integer '1-�շ�(����),2-����(��/ס)
        If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), (mint��¼���� = 1 Or mint������Դ = 1)) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
    End If
    mlng���� = -1
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, mint������Դ, int����, True, str���, , , str��׼��Ŀ, _
        zl��ȡ��ҩ��̬(Bill.Row), IIf(mbln���õǼ�, True, False), mbln����ˢ��, mlng����, _
        mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If lng��Ŀid <> 0 Then
        Bill.Text = lng��Ŀid
        mblnSelect = True
        Call bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call gobjCommFun.PressKey(13)
        End If
    End If
    mblnSelect = False
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency
    Dim blnStock As Boolean, lngDoUnit As Long, strժҪ As String
    Dim lng��Ŀid As Long, str��׼��Ŀ As String, str��� As String
    Dim blnInput As Boolean, cur��� As Currency, lng���˿���ID As Long, int���� As Integer, lngOld���� As Long
    Dim colStock As Collection
    Dim strPriceGrade As String
    On Error GoTo errH

    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "ִ�п���" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = 0 Then Exit Sub
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "���"
                If Bill.ListIndex <> -1 Then '�������������򲻻���������
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        'һ���ĸ��շ����,�����(����)ԭ�и���Ŀ����
                        For i = 2 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .�շ�ϸĿID = 0: .�շ���� = ""
                            End With
                            Call CalcMoneys
                            Call ShowMoney
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '��ʱ��RowData��¼��ѡ����շ����
                End If
            Case "��Ŀ"
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                    If mobjBill.Details.Count >= Bill.Row Then
                        'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                        If Bill.TextMatrix(Bill.Row, BillCol.��Ŀ) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                
                    sta.Panels(2).Text = ""
                    sta.Panels(4).Text = ""
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '��������ñ�־
                        Set mobjDetail = GetInputDetail(Val(Bill.Text), mlng����)
                    Else
                        If gSysPara.bln�շ���� Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "û��ȷ���������,�����������"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            str��� = IIf(SelectIsNurse, "'E','M','4'", mstr�շ����)
                        End If
                        If Not IsNull(mrsInfo!����) Then
                            int���� = mrsInfo!����
                           
                            '���˺�:24862
                            'mint������Դ As Integer '1-���ﲡ��,2-סԺ����
                            'mint��¼���� As Integer '1-�շ�(����),2-����(��/ס)
                            If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), (mint��¼���� = 1 Or mint������Դ = 1)) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
                        End If
                        mlng���� = -1
                        lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, mint������Դ, int����, True, str���, Bill.Text, _
                            Bill.TxtHwnd, str��׼��Ŀ, zl��ȡ��ҩ��̬(Bill.Row), IIf(mbln���õǼ�, True, False), _
                            mbln����ˢ��, mlng����, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                        If lng��Ŀid <> 0 Then
                            Set mobjDetail = GetInputDetail(lng��Ŀid, mlng����)
                            If int���� <> 0 Then sta.Panels(4).Text = Getҽ������(lng��Ŀid, int����)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Bill.TxtVisible = False '(���Ӳ���)
                    

                    'ҽ��������Ŀ�Ƿ��������
                    If CheckYBFeeItemVerfiy(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!����)), mobjDetail) = False Then
                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                     
                    
                    '�������ò��˲�������
                    If mint������Դ = 2 And mint��¼���� = 2 Then
                        If InStr(",5,6,7,", mobjDetail.���) = 0 Then
                            If Not CheckFeeItemLimitDept(mobjDetail.ID) Then
                                MsgBox "���շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    
                    '���ҩƷ�����Ƿ��ظ�:������ʱ��ͬһҩ���������ظ�(����ֻ����)
                    If InStr(",5,6,7,", mobjDetail.���) > 0 _
                        Or (mobjDetail.��� = "4" And mobjDetail.��������) Then
                        If PhysicExist(mobjDetail, Bill.Row) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��鴦��ְ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        mobjDetail.����ְ�� = Get����ְ��(mobjDetail.ID)
                        'ҽ���򹫷Ѳ���
                        '����:45605
                        If zlIsCheckMedicinePayMode(txt���ʽ) Then
                            If CheckDuty(mobjDetail, False) > 0 Then
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                        '���в���
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '���˿���ID
                    lng���˿���ID = mobjBill.����ID
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    
                    'ȱʡִ�п���
                    lngDoUnit = Val("" & mrsInfo!����ID)
                    If mobjDetail.��� = "4" And mlng���ϲ��� > 0 Then lngDoUnit = mlng���ϲ���
                    If lngDoUnit = 0 Then lngDoUnit = lng���˿���ID
                    
                    lngDoUnit = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mobjDetail.���, mobjDetail.ID, _
                        mobjDetail.ִ�п���, lng���˿���ID, Get��������ID, mint������Դ, lngDoUnit, 1, 1)
                    
                    
                    '��ȡҩƷ�����������Ϣ
                    Call ReadDrugAndStuffStock(lngDoUnit, mobjDetail)
                   
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        '��������
                        mobjDetail.�������� = Get��������(mobjDetail.ID)
                    End If
                    
                    
                    '����ժҪ(ȡ���е����Ա��޸�)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            strժҪ = mobjBill.Details(Bill.Row).ժҪ
                        End If
                    End If
                    
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051
                    '����ժҪ(������������и���ժҪ)
                    If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        End If
                    '83877,Ƚ����,2015-4-14,���ﲡ�˸���"���˹Һż�¼"��ȡ"����",����GetItemInfo
                    Else
                        If mrsInfo.State = 1 Then    '����:��Ҫ�Ǵ���һԺҪ��,����BH���ܵǼ�,����û��BugNo  '90304
                            strժҪ = gclsInsure.GetItemInfo(Val(Nvl(mrsInfo!����)), mrsInfo!����ID, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        Else
                            strժҪ = gclsInsure.GetItemInfo(0, 0, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        End If
                        mobjBill.Details(Bill.Row).ժҪ = strժҪ
                    End If
                    
                    Call CalcMoneys(Bill.Row)
                    
                    'Calcmoney��ҽ�����ܷ���ժҪ
                    If mobjBill.Details(Bill.Row).ժҪ <> "" Then strժҪ = mobjBill.Details(Bill.Row).ժҪ
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mint��¼���� = 2 And mrsWarn.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        '���˺�:30504: and mbln���õǼ�=False
                        If curTotal > 0 And mbln���õǼ� = False Then
                            cur��� = Val(txtʵ��.Tag)
                            If gSysPara.bln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(IIf(mint������Դ = 1, 0, 1), mrsInfo!����ID)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mint��¼���� = 2 Then
                        If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                            If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), 2, 0, Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, pҽ�����ѹ���, IIf(mint������Դ = 2, 1, 0), 0, _
                            MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), mint��¼����, IIf(mint��¼���� = 1, 1, 0), Bill.Row)) = False Then
                            mobjBill.Details.Remove Bill.Row  'ɾ���ո���Ҫ����ķ�����
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '�������ͼ��
                    Call Check��������(Bill.Row)
                    

                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    mlngPreRow = 0  '�޸�������ʱ,�ָ���ֵ,�Ա���ʾ���
                    With mobjBill.Details(Bill.Row)
                        '��һ�е�����ȷ��
                        If .�շ���� = "7" And mblnPay Then Bill.ColData(BillCol.����) = 4 '����
                        If .�շ���� = "F" Then Bill.ColData(BillCol.��־) = -1 '���ӱ�־
                        
                        '���������������
                        If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                            And Not (.�շ���� = "4" And .Detail.��������) Then
                            Bill.ColData(BillCol.����) = IIf(mblnTime, 4, 5) '����
                            Bill.ColData(BillCol.����) = 4 '����
                        Else
                            Bill.ColData(BillCol.����) = 4 '����
                            Bill.ColData(BillCol.����) = 5 '����
                        End If
                        
                        'ִ�п���
                        mblnEnterCell = False: Bill.Col = BillCol.ִ�п���: mblnEnterCell = True
                        Call FillBillComboBox(Bill.Row, BillCol.ִ�п���, Not blnInput) 'ֱ�ӻس�ʱ����ִ�п���
                        mblnEnterCell = False: Bill.Col = BillCol.��Ŀ: mblnEnterCell = True
                        
                        blnSkip = Bill.ListCount = 1
                        If Not blnSkip And InStr(",4,5,6,7,", .�շ����) > 0 Then
                            'ָ���˹̶�ҩ��ʱ,��������ѡ��
                            Select Case .�շ����
                                Case "4"
                                    blnSkip = mlng���ϲ��� > 0 And .ִ�в���ID = mlng���ϲ���
                                Case "5"
                                    blnSkip = mlng��ҩ�� > 0 And .ִ�в���ID = mlng��ҩ��
                                Case "6"
                                    blnSkip = mlng��ҩ�� > 0 And .ִ�в���ID = mlng��ҩ��
                                Case "7"
                                    blnSkip = mlng��ҩ�� > 0 And .ִ�в���ID = mlng��ҩ��
                            End Select
                        End If
                        If blnSkip Then
                            Bill.ColData(BillCol.ִ�п���) = 5: .Key = 1
                        Else
                            Bill.ColData(BillCol.ִ�п���) = 3: .Key = Bill.ListCount
                        End If
                        
                        If lngDoUnit <> .ִ�в���ID Then
                            '��ȡҩƷ�����������Ϣ
                            Call ReadDrugAndStuffStock(.ִ�в���ID, .Detail)
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                                                
                         '������Ŀ����,�������շ���Ŀ�д�����Ŀ����δȡ��ȡ,ҩƷ�����ж�,ҩƷ��������������
                        If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" And InStr(",5,6,7,", .�շ����) = 0 Then
                            If (gSysPara.bln��������ۿ� And mobjBill.Details(Bill.Row).�������� = 0) Or Not gSysPara.bln��������ۿ� Then  '(����м���,ֻȡһ��)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem
                                   mlngPreRow = 0 'ͨ���б仯��־������ȷ��������
                                End If
                            End If
                        End If
                        
                    End With
                End If
                
                'ֻ����һ�θ���
                If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Details(i).�շ���� = "7" Then
                                '����ִ�иù��̣�����ᶨλ��һ����Ԫ,�ȶ�λ������,����һ����Ԫ������
                                'ѡ����øù��̣����ú���͸��س������ﲻ���ٻس��������������س���Ч��(�ؼ�ԭ��)��
                                Bill.Col = BillCol.����: Exit For
                            End If
                        Next
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "����Ӧ��Ϊ����������", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    '�������
                    If gSysPara.dblMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gSysPara.dblMaxMoney Then
                            If MsgBox("��ǰ������" & gSysPara.dblMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                
                    '����ҩ���Ǵ�����Ŀ�ſɸ��ĸ���(������ı�,����Ҳ��)
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then 'And mobjBill.Details(Bill.Row).�������� = 0 Then
                        '������ʱ��ҩƷ�����ֹ����(û�з�����ʱ��ҩƷ�����޸ĸ���������)
                        If mobjBill.Details(Bill.Row).Detail.���� Or mobjBill.Details(Bill.Row).Detail.��� Then
                            If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� > mobjBill.Details(Bill.Row).Detail.��� Then
                                MsgBox """" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                              
                        '�������ʱ�ۻ������ҩ���ĸ��������Ƿ��㹻
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" _
                                And (mobjBill.Details(i).Detail.��� Or mobjBill.Details(i).Detail.����) Then
                                If Val(Bill.Text) * mobjBill.Details(i).���� > mobjBill.Details(i).Detail.��� Then
                                    MsgBox "�� " & i & " ��ҩƷ""" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                        
                        '���㲢ˢ�¸���
                        lngOld���� = mobjBill.Details(Bill.Row).����
                        mobjBill.Details(Bill.Row).���� = Bill.Text
                        Call CalcMoneys(Bill.Row)
                                                
                        If mint��¼���� = 2 Then
                            If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                                If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), 2, 0, Bill.Row)) = False Then
                                    mobjBill.Details(Bill.Row).���� = lngOld����
                                    Bill.Text = lngOld����
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, pҽ�����ѹ���, IIf(mint������Դ = 2, 1, 0), 0, _
                                MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), mint��¼����, IIf(mint��¼���� = 1, 1, 0), Bill.Row)) = False Then
                                mobjBill.Details(Bill.Row).���� = lngOld����
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                                               
                         '����������ҩ����,����Ƕ�����,���޸������Ǵ����,����Ǵ���,���޸�ͬһ����Ĵ����.��Ϊ�޶�Ϊ�в�ҩ,������������
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" And mobjBill.Details(i).�������� = mobjBill.Details(Bill.Row).�������� Then
                                If mobjBill.Details(i).�������� = 0 Or (mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).Detail.���д��� = 0) Then     '1��2�̶��Ͱ������Ĳ���
                                    mobjBill.Details(i).���� = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                                                
                        Call ShowMoney
                    Else
                        sta.Panels(2) = "������Ŀ�ĸ������ܸ��ģ�"
                        Bill.Text = mobjBill.Details(Bill.Row).����: Beep '�ָ�ԭ�и���ֵ
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                    End If
                    '�������
                    If gSysPara.dblMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gSysPara.dblMaxMoney Then
                            If MsgBox("��ǰ������" & gSysPara.dblMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '�����Ϸ��Լ��
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� < 0 Then
                        'Ȩ��
                        If Not ((InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 And InStr(mstrPrivs, "ҩƷ��������") > 0) _
                             Or (InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) = 0 And InStr(mstrPrivs, "���Ƹ�������") > 0)) Then
                            MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        Else
                            If mobjBill.Details(Bill.Row).Detail.���� Then
                                MsgBox "����ҩƷ���������븺����", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                            If mrsInfo.State = 1 And mint��¼���� = 2 Then
                                If Not IsNull(mrsInfo!����) Then
                                    If Not MCPAR.�������� Then
                                        MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    'ҩƷ�����
                    With mobjBill.Details(Bill.Row)
                        If (.�շ���� = "4" And .Detail.��������) Or InStr(",5,6,7,", .�շ����) > 0 Then
                            If .Detail.���� Or .Detail.��� Then
                                '������ʱ��ҩƷ�����ֹ����
                                If .���� * CSng(Bill.Text) > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .ִ�в���ID) <> 0 And Bill.ColData(BillCol.ִ�п���) = 5 Then
                                    '����ҩƷ�������
                                    If .���� * CSng(Bill.Text) > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .����: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox """" & .Detail.���� & """�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Bill.Text = .����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                        dblPreTime = .����
                        .���� = Bill.Text
                        
                        '�����������
                        If Not CheckLimit(mobjBill, Bill.Row, mblnҩ����λ) Then
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                        '�������¼������
                        If .Detail.¼������ > 0 And .���� * .���� * IIf(mblnҩ����λ, .Detail.ҩ����װ, 1) > .Detail.¼������ Then
                            If MsgBox("��������γ�����¼������" & .Detail.¼������ & ",�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                                .���� = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        '����ʹ������
                        If mint������Դ = 2 And mint��¼���� = 2 And mrsInfo.State = 1 Then
                            If .Detail.Ҫ������ And Not IsNull(mrsInfo!����) And Not mrsMedAudit Is Nothing Then
                                mrsMedAudit.Filter = "��ĿID=" & .�շ�ϸĿID
                                If mrsMedAudit.RecordCount > 0 Then
                                    If Not IsNull(mrsMedAudit!��������) Then
                                        If .���� * .���� * IIf(mblnҩ����λ, .Detail.ҩ����װ, 1) > mrsMedAudit!�������� Then
                                            MsgBox "��������γ�������׼�Ŀ�������" & FormatEx(mrsMedAudit!�������� / IIf(mblnҩ����λ, .Detail.ҩ����װ, 1), 5) & "��", vbInformation, gstrSysName
                                            .���� = dblPreTime: Bill.Text = dblPreTime
                                            Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                        If .�������� <> 0 And .Detail.���д��� <> 0 Then
                            sta.Panels(2) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                    End With
                
                    Call CalcMoneys(Bill.Row)
                    
                    '����������(���Ѿ�������з��õ�δ��ʾǰ)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                        mobjBill.Details(Bill.Row).���� = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mint��¼���� = 2 And mrsWarn.State = 1 Then
                        curTotal = GetBillTotal(mobjBill)
                        
                        '���˺�:2010-07-01 10:23:11:30504:and mbln���õǼ�=False
                        If curTotal > 0 And mbln���õǼ� = False Then
                            cur��� = Val(txtʵ��.Tag)
                            If gSysPara.bln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(IIf(mint������Դ = 0, 0, 1), mrsInfo!����ID)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details(Bill.Row).���� = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                                     
                    If mint��¼���� = 2 Then
                        If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                            If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), 2, 0, Bill.Row)) = False Then
                                mobjBill.Details(Bill.Row).���� = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, pҽ�����ѹ���, IIf(mint������Դ = 2, 1, 0), 0, _
                            MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), mint��¼����, IIf(mint��¼���� = 1, 1, 0), Bill.Row)) = False Then
                            mobjBill.Details(Bill.Row).���� = dblPreTime
                            Call CalcMoneys(Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '��������д���������
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).�������� = Bill.Row Then
                            '28136
                            '���������ĸ���,��Ҫ���¼��еĸ������и��³ɸ���
                            With mobjBill.Details(i)
                                If .Detail.���д��� = 0 Then  '�ǹ��д���
                                    If Abs(.����) <> Abs(.Detail.��������) Then GoTo NotCalc:
                                    .���� = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.��������
                                ElseIf .Detail.���д��� = 1 Then '�̶��Ĺ��д���
                                    .���� = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.�������� = 0, 1, .Detail.��������)
                                ElseIf .Detail.���д��� = 2 Then   '�������Ĺ��д���
                                    .���� = Val(Bill.Text) * .Detail.��������
                                Else
                                     GoTo NotCalc:
                                End If
                            End With
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
NotCalc:
                        End If
                    Next
                    
                    Call ShowMoney

                 ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
               End If
                If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "��Ŀ�۸�Ӧ��Ϊ������Ҫɾ�����ã������븺��������ʵ�֣�", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '�������
                    If gSysPara.dblMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * mobjBill.Details(Bill.Row).���� > gSysPara.dblMaxMoney Then
                            If MsgBox("��ǰ������" & gSysPara.dblMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If

                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '���û�ж�Ӧ��������Ŀ,���޷�����
                    If mobjBill.Details(Bill.Row).Detail.��� And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                        If Not (mobjBill.Details(Bill.Row).InComes(1).�ּ� = 0 And mobjBill.Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                            strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).ԭ��, mobjBill.Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(2) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).��׼����
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        
                        dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).��׼����
                        
                        mobjBill.Details(Bill.Row).InComes(1).��׼���� = Bill.Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                        Call CalcMoneys(Bill.Row)
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        If mint��¼���� = 2 And mrsWarn.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            '30504:and mbln���õǼ�=False
                            If curTotal > 0 And mbln���õǼ� = False Then
                                cur��� = Val(txtʵ��.Tag)
                                If gSysPara.bln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(IIf(mint������Դ = 1, 0, 1), mrsInfo!����ID)
                                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                    Nvl(mrsInfo!������, 0), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, mintWarn)
                                If mbytWarn = 2 Or mbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).��׼���� = dblPreMoney
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        Call ShowMoney
                    Else
                        Bill.Text = "0"
                        sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                        Beep
                    End If
                End If
            Case "ִ�п���"
                If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Details(Bill.Row)
                            If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                                .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                                If CheckItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                            End If
                    
                            'ҩƷ�����:��̬ҩ��,������ʱ��ҩƷҲҪ�����
                            If (.�շ���� = "4" And .Detail.��������) Or InStr(",5,6,7,", .�շ����) > 0 Then
                                If .Detail.���� Or .Detail.��� Then '������ʱ��ҩƷ��治���ֹ����
                                    If .���� * .���� > .Detail.��� Then
                                        If .�շ���� = "4" Then
                                            MsgBox "[" & .Detail.���� & "]Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                        Else
                                            MsgBox "[" & .Detail.���� & "]Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                        End If
                                        Cancel = True
                                    End If
                                Else
                                    Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                    If colStock("_" & .ִ�в���ID) <> 0 Then
                                        If .���� * .���� > .Detail.��� Then
                                            If colStock("_" & .ִ�в���ID) = 1 Then
                                                If MsgBox("[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                    Cancel = True
                                                End If
                                            ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                                MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                                Cancel = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                            '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                            If .�շ���� = "4" And .Detail.�������� Then
                                Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                            End If
                        
                            If CheckItemHaveSub(Bill.Row) Then
                                KeyCode = 0
                                Call LocateMainItemNextRow(Bill.Row)
                            End If
                            If mint��¼���� = 2 Then
                                If Not IsNull(mrsInfo!����) And mobjBill.Details(Bill.Row).���� <> 0 And MCPAR.ʵʱ��� Then
                                    If gclsInsure.CheckItem(mrsInfo!����, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), 2, 0, Bill.Row)) = False Then
                                        Bill.Text = "": Cancel = True
                                        Bill.TxtVisible = False: Exit Sub
                                    End If
                                End If
                            End If
                            
                            If mobjBill.Details(Bill.Row).���� <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, pҽ�����ѹ���, IIf(mint������Դ = 2, 1, 0), 0, _
                                    MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), mint��¼����, IIf(mint��¼���� = 1, 1, 0), Bill.Row)) = False Then
                                    Bill.Text = "": Cancel = True
                                    Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                    End With
                End If

        End Select
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� = lngRow Then
            If mobjBill.Details(i).Detail.���д��� = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.����
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    End If
    '����:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetSubItem()
'����:�����շ���Ŀ��,���ص�ǰ�շ���Ŀ�Ĵ�����Ŀ�����ü�����,����ʾ�ڵ��ݿؼ���
'����:
'������:Bill_KeyDown��������Ŀ��
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long               'ִ�п���ID
Dim bln��������ۿ� As Boolean
Dim strժҪ As String
Dim strPriceGrade As String

lngMainRow = Bill.Row               '�������
If gSysPara.bln��������ۿ� Then            '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
    bln��������ۿ� = Not mobjBill.Details(lngMainRow).Detail.���ηѱ�
End If

With mobjBill.Details(lngMainRow)
    Set mcolDetails = New Details
    Set mcolDetails = GetSubDetails(.�շ�ϸĿID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)
            mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = "" '�б�Ҫ����
        
        'a.������ĿΪ��ҩƷ��Ŀ��ִ�п���
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).���) = 0 Then
             If mcolDetails(i).��� = .�շ���� Then
                '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                lngDoUnit = .ִ�в���ID
             Else
                If mcolDetails(i).ִ�п��� = 0 Then
                    '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                    lngDoUnit = .ִ�в���ID
                Else
                    '������ҩ��Ŀ��ִ�п���
                    lngDoUnit = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mcolDetails(i).���, _
                        mcolDetails(i).ID, mcolDetails(i).ִ�п���, lngDoUnit, Get��������ID, mint������Դ, , 1, 1)
                End If
             End If
        End If
        
        'b.������ĿΪҩƷ,���ĵ�ִ�п���(��������ִ�п���Ϊ��,Ҳ��ִ�е�����)
        If lngDoUnit = 0 Then
            lngDoUnit = mobjBill.����ID
            If lngDoUnit = 0 And cbo��������.ListIndex <> -1 Then
                lngDoUnit = cbo��������.ItemData(cbo��������.ListIndex)
            End If
            lngDoUnit = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mcolDetails(i).���, mcolDetails(i).ID, _
                mcolDetails(i).ִ�п���, lngDoUnit, Get��������ID, mint������Դ, .ִ�в���ID, 1, 1) '���Ĵ���ȱʡ������ִ�п�����ͬ
        End If
            
        '����֧����Ŀ��Ӧ���
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) Then
                If InStr(",5,6,7,", mcolDetails(i).���) > 0 Then
                    strPriceGrade = mstrҩƷ�۸�ȼ�
                ElseIf mcolDetails(i).��� = "4" Then
                    strPriceGrade = mstr���ļ۸�ȼ�
                Else
                    strPriceGrade = mstr��ͨ�۸�ȼ�
                End If
                If zlCheck������۸����(mcolDetails(i).ID, Not mcolDetails(i).���, strPriceGrade) Then
                    '����:27286
                Else
                    If Not ItemExistInsure(mrsInfo!����ID, mcolDetails(i).ID, mrsInfo!����) Then
                        If gSysPara.intҽ������ = 1 Then
                            If MsgBox("��Ŀ""" & mcolDetails(i).���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        ElseIf gSysPara.intҽ������ = 2 Then
                            MsgBox "��Ŀ""" & mcolDetails(i).���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        Call SetDetailtStock(lngDoUnit, mcolDetails(i))
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln��������ۿ�)
        Call ShowDetails(Bill.Rows - 1)
        
        '83877,Ƚ����,2015-4-14,���ﲡ�˸���"���˹Һż�¼"��ȡ"����",����GetItemInfo
        'mint������Դ = 2:41136
        'CalcMoney���ȵ���GetuItemInsure���ܷ���ժҪ
        strժҪ = mobjBill.Details(Bill.Rows - 1).ժҪ  '90304
        If mrsInfo.State = 1 Then
            strժҪ = gclsInsure.GetItemInfo(Val(Nvl(mrsInfo!����)), mrsInfo!����ID, mcolDetails(i).ID, strժҪ, 2)
        Else
            strժҪ = gclsInsure.GetItemInfo(0, 0, mcolDetails(i).ID, strժҪ, 2)
        End If
        mobjBill.Details(Bill.Rows - 1).ժҪ = strժҪ
    Next
    
    If bln��������ۿ� And Not mbln���õǼ� Then
        Call CalcMoney(lngMainRow, bln��������ۿ�) '�����������Ӧ����ʵ��,��Ϊ��û�м������ǰ�����ǰ������������.
        
        Call Calc��������ʵ��(lngMainRow)
    End If
    
    Call ShowMoney
End With

End Sub

Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = UserInfo.����ID
    End If
End Function

Private Sub Calc��������ʵ��(ByVal lngMainRow As Long)
'����:����������ۿ�ʱ,����ָ�����������ID�ĵ�һ��������Ŀ���������ʵ�ս��
'����:  lngMainRow-������ID

Dim i As Long, j As Long
Dim cur����ǰӦ�պϼ� As Currency     '��¼�����������Ӧ�պϼ�
Dim cur���ۺ�ʵ�� As Currency


With mobjBill
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).�������� <> lngMainRow Then Exit For    '��ȻĿǰ�����˲������ڴ����м������������,����һ�ŵ�����������,Ϊ�˽������ܵ�����,����ȫ��ɨ��
        
        If i = lngMainRow Or .Details(i).�������� = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur����ǰӦ�պϼ� = cur����ǰӦ�պϼ� + .Details(i).InComes(j).Ӧ�ս��
            Next
        End If
    Next
   
    cur���ۺ�ʵ�� = CCur(Format(ActualMoney(.�ѱ�, .Details(lngMainRow).InComes(1).������ĿID, cur����ǰӦ�պϼ�), gSysPara.Money_Decimal.strFormt_VB))
    
    cur���ۺ�ʵ�� = cur���ۺ�ʵ�� - cur����ǰӦ�պϼ� + .Details(lngMainRow).InComes(1).Ӧ�ս��
    
    .Details(lngMainRow).InComes(1).ʵ�ս�� = Format(cur���ۺ�ʵ��, gSysPara.Money_Decimal.strFormt_VB)
    .Details(lngMainRow).InComes(1).Key = "_" & Format(cur���ۺ�ʵ��, gSysPara.Money_Decimal.strFormt_VB)
    
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubDept(ByVal lngRow As Long)
Dim i As Long, j As Long
    With mobjBill
        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID) '������ȡ
        
        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).ִ�п��� = 0 Then
                                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                 .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                            Else
                                '3.������ҩ��Ŀ��ִ�п���
                                If cbo��������.ListIndex <> -1 Then
                                    .Details(i).ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                                End If
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, mcolDetails(j).���, _
                                    mcolDetails(j).ID, mcolDetails(j).ִ�п���, .Details(i).ִ�в���ID, Get��������ID, mint������Դ, , 1, 1)
                            End If
                        End If
                    End If
                    
                    '��ʾ����ִ�п���
                    If .Details(i).ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.Details(i).ִ�в���ID)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.Details(i).ִ�в���ID)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = lngRow Then
                CheckItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim strStock As String, i As Long
    
    If Not Bill.Active Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If Not mblnEnterCell Then Exit Sub
    
    If mbytInState = 3 Then
        '����б༭����������ɫ
        Bill.SetColColor BillCol.���, &HE7CFBA '��ȻҪ�ɰ�ɫ
        Exit Sub
    End If
    If mlngPreRow <> Row Then sta.Panels(2).Text = "" '�л���ʱ�����,70160
     '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '��ʾ���
            If InStr(",5,6,7,", .�շ����) > 0 And .�շ�ϸĿID <> 0 Then
                If mbln����ҩ�� Or mbln����ҩ�� Then
                    strStock = GetStockInfo(.�շ�ϸĿID, mbln����ҩ��, mbln����ҩ��, mblnҩ����λ, mstrҩ����װ)
                    If strStock <> "" Then
                        If InStr(1, mstrPrivs, "��ʾ���") > 0 Then
                            sta.Panels(Pan.C2��ʾ��Ϣ) = "[" & .Detail.���� & "]���ÿ��:" & strStock
                        Else
                            sta.Panels(Pan.C2��ʾ��Ϣ) = "[" & .Detail.���� & "]�п��."
                        End If
                    End If
                End If
                If strStock = "" Then
                    '��ʱ���¿����ʾ
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0) '29680
                    If mblnҩ����λ Then
                        .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                    End If
                    Call ShowStock(.Detail.����, .Detail.���)
                End If
            ElseIf .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0, , , .Detail.����) '29680
                Call ShowStock(.Detail.����, .Detail.���)
            ElseIf .Detail.��� And .InComes.Count > 0 And Bill.TextMatrix(0, Bill.Col) = "����" Then
                sta.Panels(2) = "�۸�Χ:" & FormatEx(.InComes(1).ԭ��, 5) & "-" & FormatEx(.InComes(1).�ּ�, 5)
            Else
                sta.Panels(2) = ""
            End If
            
            Bill.ColData(BillCol.���) = IIf(gSysPara.bln�շ����, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
            
             '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
            If CheckItemHaveSub(Row) Or .�������� > 0 Then
                Bill.ColData(BillCol.���) = BillColType.Text_UnModify
                Bill.ColData(BillCol.��Ŀ) = BillColType.Text_UnModify
            End If
            
            '����Ƿǵ���״̬
            If mbytInState <> 2 Then
                If .�շ���� = "7" And mblnPay Then
                    Bill.ColData(BillCol.����) = 4
                Else
                    Bill.ColData(BillCol.����) = 5
                End If
                
                '���������������
                If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                    And Not (.�շ���� = "4" And .Detail.��������) Then
                    Bill.ColData(BillCol.����) = IIf(mblnTime, 4, 5) '����
                    Bill.ColData(BillCol.����) = 4 '���
                Else
                    Bill.ColData(BillCol.����) = 4
                    Bill.ColData(BillCol.����) = 5
                End If
                
                If .Key = "1" Then    'ָ���˹̶�ҩ��ʱ,��������ѡ��ִ�п���
                    Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus
                Else
                    Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
                End If
                
                If .�շ���� = "F" Then
                    Bill.ColData(BillCol.��־) = -1
                Else
                    Bill.ColData(BillCol.��־) = 5
                End If
                
                 'ֻ����һ�����
                If mblnOne Then Bill.ColData(BillCol.���) = 5
            End If
        End With
        
        '��ʾժҪ
        If mobjBill.Details(Bill.Row).ժҪ <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  ժҪ:" & mobjBill.Details(Bill.Row).ժҪ
        End If
    End If
   
    '������δ�������,��ָ��е�����
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.���) = IIf(gSysPara.bln�շ����, BillColType.ComboBox, BillColType.UnFocus) '�����,��������ʱ�ᱻ�ı�
        Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
    End If
    
    
    '-----------------------------------------------------------------
    '2.�иı��������ݴ������ʾ����
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then
        Call FillBillComboBox(Bill.Row, Bill.Col, True) '�������
    End If
    
    If gSysPara.bln�շ���� And Bill.TextMatrix(Row, BillCol.���) = "" And mblnOne Then
        mrsClass.Filter = "����=" & mstr�շ����
        Bill.TextMatrix(Row, BillCol.���) = mrsClass!���
        Bill.RowData(Row) = Asc(mrsClass!����)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "���" '�������շ����ʱ������������
            Call gobjControl.CboSetWidth(Bill.CboHwnd, 1000)
            '������Ϊ��,���Զ�Ĭ��Ϊ��һ�շ�ϸĿ�����
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "����=" & mstr�շ����
                    Bill.TextMatrix(Row, Col) = mrsClass!���
                    Bill.RowData(Row) = Asc(mrsClass!����)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "ִ�п���"
            Call gobjControl.CboSetWidth(Bill.CboHwnd, 2000)
        Case "����"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789" & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                '�ɷ�����С��
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                    If InStr(mstrPrivs, "ҩƷС������") > 0 Then
                        Bill.TextMask = "." & Bill.TextMask
                    End If
                Else
                    Bill.TextMask = "." & Bill.TextMask
                End If
                
                '�ɷ����븺��
                If Not mobjBill.Details(Bill.Row).Detail.���� Then
                    If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                        If InStr(mstrPrivs, "ҩƷ��������") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivs, "���Ƹ�������") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                                    
                    If InStr(Bill.TextMask, "-") > 0 Then
                        If mrsInfo.State = 1 And mint��¼���� = 2 Then
                            If Not IsNull(mrsInfo!����) Then
                                If Not MCPAR.�������� Then
                                    Bill.TextMask = Replace(Bill.TextMask, "-", "")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case "����"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
    
    '����,����������е����ʱ,�������л�û�п�ʼ
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'bill.ToolTipText = bill.TextMatrix(bill.MouseRow, bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, strDoctor As String
    Dim rsReturn As ADODB.Recordset
    Dim intIndex As Integer
    If mbytInState <> 0 Then Exit Sub
    mrsAll��������.Filter = ""
    If cbo��������.ItemData(cbo��������.ListIndex) = 0 And cbo��������.Text Like "��������*" Then
        If gobjDatabase.zlShowListSelect(Me, glngSys, 1150, cbo��������, mrsAll��������, True, "", "ȱʡ,���ȼ�", rsReturn) = False Then
            mobjBill.��������ID = 0
            Exit Sub
        End If
        If rsReturn Is Nothing Then Exit Sub
        If rsReturn.State <> 1 Then Exit Sub
        If rsReturn.RecordCount = 0 Then Exit Sub
        rsReturn.MoveFirst
        If gobjControl.CboLocate(cbo��������, Val(rsReturn!ID), True) = False Then
            cbo��������.RemoveItem cbo��������.ListCount - 1
            cbo��������.AddItem IIf(zlIsShowDeptCode, rsReturn!���� & "-", "") & rsReturn!����
            cbo��������.ItemData(cbo��������.ListCount - 1) = Val(Nvl(rsReturn!ID))
            intIndex = cbo��������.NewIndex
            cbo��������.AddItem "�������ҡ�"
            cbo��������.ItemData(cbo��������.ListCount - 1) = 0
            cbo��������.ListIndex = intIndex
        End If
        Exit Sub
    End If
    '��λҽ��
    cbo������.Clear
    If cbo��������.ListIndex <> -1 Then
        Call Load������(cbo��������.ItemData(cbo��������.ListIndex))
    End If
    
    '���ݶ���
    If mbytInState = 0 Then
        If cbo��������.ListIndex = -1 Then
            mobjBill.��������ID = 0
        Else
            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
    End If
    
    
    '�������������Ŀ��ִ�п���
    If mbytInState = 0 And cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '�������շ���Ŀ
                If InStr(",4,5,6,7,", .Detail.���) = 0 And .Detail.ִ�п��� = 6 Then '6-�����˿���
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    'ˢ����ʾ����ִ�п���
                    If i <= Bill.Rows - 1 And .ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = Get��������(.ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                    '����8113���޸�
'                ElseIf InStr(",4,5,6,7,", .Detail.���) > 0 Then
'                '�������ҩ��Ϊ�洢�ⷿ�����õķ����ڲ��˿���(��������)��ִ�п���
'                    If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
'                        Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
'                    End If
'                    If .Key = "1" Then .Key = "0"        '1��ʾִ�п��Ҳ���ѡ��
'                    mlngPreRow = 0      '������Entercell�¼���������ִ�п��ҵĿ�ѡ��
                End If
            End With
        Next
    End If
End Sub

Private Sub cbo������_Click()
    Dim arrDepts As Variant, i As Long, k As Long
    
    If mbytInState = 0 Then
        mobjBill.������ = IIf(cbo������.ListIndex = -1, "", NeedName(cbo������.Text))
                        
        '��ʿ���
        If Bill.Active Then
            If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
                And Bill.RowData(Bill.Rows - 1) <> 0 Then
                '�����Ч����
                Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = ""
                Bill.RowData(Bill.Rows - 1) = 0
            ElseIf Bill.Col = BillCol.��� Then
                Call Bill_EnterCell(Bill.Row, Bill.Col) 'ˢ��
            End If
        End If
        
        '��ʿ���:�жϷǷ�����
'        If HaveStopClass > 0 Then
'            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
'        End If
    End If
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        Call GetCboIndex(cbo������, NeedName(cbo������.Text))
        If cbo������.ListIndex = -1 Then cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�Ӱ�_Click()
    If mbytInState = 1 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk�Ӱ�.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    blnAdd = OverTime
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count > 0 Or mblnOK Then
        If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Function CheckNegative() As Boolean
'���ܣ���鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim strItems As String, str���� As String
    Dim str��λ As String, dbl���� As Double
    Dim strValues(0 To 10) As String, intR As Long
    Dim strSubTable As String, dbl���κϼ� As Double, dbl�ѽ����� As Double
    
    '����:26951
    If InStr(1, mstrPrivs, ";�������ʲ���鷢����Ŀ;") > 0 Then
        '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
        CheckNegative = True: Exit Function
    End If
    
    CheckNegative = True
    If mobjBill.����ID = 0 Then Exit Function
    
    strItems = ""
    strSubTable = ""
    intR = 0
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                If Len(strItems) > 2000 Then
                    If intR <= 10 Then
                        strValues(intR) = Mid(strItems, 2)
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As �շ�ϸĿID,  " & _
                        "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As ִ�в���ID,0 as ����,0 as �������� " & _
                        " From Table(Cast(f_str2list([" & intR + 3 & "]) As ZLTOOLS.t_strlist))"
                    Else
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As �շ�ϸĿID,  " & _
                        "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As ִ�в���ID,0 as ����,0 as �������� " & _
                        " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
                    End If
                    strItems = "": intR = intR + 1
                End If
                strItems = strItems & "," & .�շ�ϸĿID & "_" & .ִ�в���ID & ""
'                strSQL = strSQL & " Union ALL Select " & .�շ�ϸĿID & "," & .ִ�в���ID & ",0 From Dual"
            End If
        End With
    Next
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
            " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As �շ�ϸĿID,  " & _
            "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As ִ�в���ID,0 as ����,0 as �������� " & _
            " From Table(Cast(f_str2list([" & intR + 3 & "]) As ZLTOOLS.t_strlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
            " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As �շ�ϸĿID,  " & _
            "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As ִ�в���ID,0 as ����,0 as �������� " & _
            " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
        End If
    End If
    
    If strSubTable = "" Then Exit Function
    strSubTable = Mid(strSubTable, 11)
    
    strSql = " " & _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.�շ�ϸĿID,A.ִ�в���ID,  " & _
    "             Nvl(Sum(Decode(A.��¼����, 2, 1, 3, 1, 0) * Nvl(A.����, 1) * A.����), 0) As ����, " & _
     "            Sum(Decode(nvL(Mod(M.��¼״̬ , 3),1),  0, 1, 1, 1, -1) * Decode(A.����id, Null, 0, 1) * Nvl(����, 1) * ����) As �������� " & _
     "     From " & mstrFeeTab & " A, ���˽��ʼ�¼ M " & _
     "     Where  A.����id = M.ID(+)  And A.���ʷ���=1 And A.�۸񸸺� Is Null  And A.��¼״̬<>0 " & _
     "             And A.����ID=[1] " & IIf(mint������Դ = 2, "  And Nvl(A.��ҳID,0)=[2]", "") & _
     "             And (A.�շ�ϸĿID+0,ִ�в���ID,0,0) in (select * From C1) " & _
     "     Group By A.�շ�ϸĿID,A.ִ�в���ID" & _
     "     Union ALL Select * From C1 "
   ' strSQL = _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.�շ�ϸĿID,A.ִ�в���ID,Sum(Nvl(A.����,1)*A.����) as ����," & _
    "           Sum(decode(A.����ID,NULL,0,1)* Nvl(A.����,1)*A.����) as �������� " & _
    " From  " & mstrFeeTab & " A " & _
    " Where A.��¼״̬<>0 And A.���ʷ���=1 And A.�۸񸸺� is NULL" & _
    "           And A.����ID=[1] " & IIF(mint������Դ = 2, "  And Nvl(A.��ҳID,0)=[2]", "") & _
    "           And (A.�շ�ϸĿID+0,ִ�в���ID,0,0) in (select * From C1) " & _
    " Group by A.�շ�ϸĿID,A.ִ�в���ID" & _
    " Union ALL Select * From C1"
    
    strSql = "" & _
    "   Select �շ�ϸĿID,ִ�в���ID,Sum(����) as ����,sum(��������) as �������� " & _
    "   From (" & strSql & ") " & _
    "   Group by �շ�ϸĿID,ִ�в���ID"
    
    On Error GoTo errH
     Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mobjBill.����ID, mobjBill.��ҳID, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID & " And ִ�в���ID=" & .ִ�в���ID
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        str��λ = .Detail.ҩ����λ
                        dbl���� = Nvl(rsTmp!����, 0) / .Detail.ҩ����װ
                        dbl���κϼ� = Abs(.����) * .����
                        dbl�ѽ����� = Val(Nvl(rsTmp!��������)) / .Detail.ҩ����װ
                    Else
                        str��λ = .Detail.���㵥λ
                        dbl���� = Nvl(rsTmp!����, 0)
                        dbl���κϼ� = Abs(.����) * .����
                        dbl�ѽ����� = Val(Nvl(rsTmp!��������))
                        '���ܴ���������ͬ�ļ�¼
                        '����:29412
                        For j = i + 1 To mobjBill.Details.Count
                             If .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID _
                                And mobjBill.Details(j).���� < 0 And mobjBill.Details(j).ִ�в���ID = .ִ�в���ID Then
                                dbl���κϼ� = dbl���κϼ� + Abs(.����) * .����
                             End If
                        Next
                    End If
                    '����:32106
                    If dbl���κϼ� > dbl���� - dbl�ѽ����� Then
                        Select Case gSysPara.bytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
                        Case 0  '����
                            If dbl���κϼ� > dbl���� Then
                                str���� = Get��������(.ִ�в���ID)
                                MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                    " �����ѼƷ����� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                            End If
                        Case 1   '����
                            str���� = Get��������(.ִ�в���ID)
                            If dbl���κϼ� > dbl���� Then
                                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                        " �����ѼƷ����� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                    CheckNegative = False: Exit Function
                            End If
                            
                            If MsgBox("�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                " �а������ѽᲿ��(δ��:" & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "; �ѽ�:" & FormatEx(dbl�ѽ�����, 5) & str��λ & ") ��" & vbCrLf & _
                                " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                CheckNegative = False: Exit Function
                            End If
                        Case 2   '��ֹ
                            str���� = Get��������(.ִ�в���ID)
                            MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                " �����ѼƷ����� " & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "��", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                        End Select
                    End If
                Else
                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]����������Ϊ�㣬�����������", vbInformation, gstrSysName
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call gobjComlib.SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSql As String, strTmp As String
    Dim i As Long, j As Long, lng����ID As Long
    Dim curTotal As Currency, intInsure As Integer
    Dim dblTotal As Double, cur��� As Currency, dbl���� As Double
    Dim cur���ն� As Currency, colStock As Collection
    Dim blnTrans As Boolean, strNos As String
    Dim rsItems As ADODB.Recordset
    
    If CheckValid = False Then Exit Sub
    
    If mbytInState = 3 Then
        If mint��¼���� <> 1 And (False Or mlngҽ��ID <> 0) Then '������ȫ��ɾ��
            For i = 1 To Bill.Rows - 1
                'If Bill.TextMatrix(i, Bill.Cols - 1) = "��" And Bill.RowData(i) > 0 Then
                If Bill.RowData(i) > 0 Then
                    strSql = strSql & "," & Bill.RowData(i)
                End If
            Next
            If strSql = "" Then
                MsgBox "������ѡ��һ��Ҫɾ���ķ��ã�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
            
            '������ѡ����
            strSql = Mid(strSql, 2)
            i = GetBillRows(mstrInNO, mint��¼����, mint������Դ)
            If UBound(Split(strSql, ",")) + 1 = i Then strSql = ""
        Else
            '��ΪҪ����Ϊȫ�ˣ�������ʺ��������ʣ����ݽ��ʺ��Ҫ���
            j = 0
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then j = j + 1
            Next
            i = GetBillRows(mstrInNO, mint��¼����, mint������Դ)
            If j < i Then
                MsgBox "�����еĲ�����Ŀ��ǰ�Ѳ���������(������ִ�л��ѽ��ʵ���Ŀ)��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        'ҽ�����������ϴ�(ע���ж�˳��)
        If mint������Դ = 2 Then
            intInsure = BillExistInsure(mstrInNO) '�ж��Ƿ�ҽ�����˼ǵ���
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) Then
                    'ȥ����ҽ������ƥ����
                    If strSql <> "" Then '���ܲ�������
                        MsgBox "��Ϊҽ��������Ҫ,�õ����е���Ŀ����ȫ��ɾ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If mint������Դ = 2 Then
            strSql = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','" & strSql & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            If mint��¼���� = 2 Then
                strSql = "zl_������ʼ�¼_DELETE('" & mstrInNO & "','" & strSql & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Else
                strSql = "zl_���ﻮ�ۼ�¼_DELETE('" & mstrInNO & "')"
            End If
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        
            Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
                        
            'ҽ�����������ϴ�
            If mint������Դ = 2 And intInsure > 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                End If
            End If
        
            If Not mobjSaveData Is Nothing Then
                If mobjSaveData.SaveData(mstrInNO) = False Then
                    gcnOracle.RollbackTrans
                    blnTrans = False
                    Exit Sub
                End If
            End If
            
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ�����������ϴ�
        If mint������Դ = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """��ɾ��������ҽ������ʧ�ܣ��õ�����ɾ����", vbInformation, gstrSysName
                End If
            End If
        End If
        
        If mint������Դ = 1 And mint��¼���� = 2 Then
            '110319
            If mblnDrugMachine Then
                Dim rsTemp As ADODB.Recordset
                Dim strReturn As String, strData As String '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
                strSql = "Select Id As ����id, -1 * Nvl(����, 1) * ���� As ��ҩ����" & vbNewLine & _
                        " From ������ü�¼" & vbNewLine & _
                        " Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1] And �շ���� In ('5', '6', '7')" & vbNewLine & _
                        "       And �Ǽ�ʱ�� + 0 = (Select Max(�Ǽ�ʱ��)" & vbNewLine & _
                        "                       From ������ü�¼" & vbNewLine & _
                        "                       Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1])"
                Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ѯ�����˷���Ŀ", mstrInNO)
                Do While Not rsTemp.EOF
                    strData = strData & ";" & Nvl(rsTemp!����id) & "," & Nvl(rsTemp!��ҩ����)
                    rsTemp.MoveNext
                Loop
                If strData <> "" Then
                    strData = Mid(strData, 2)
                    Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
                End If
            End If
        End If
        
        On Error GoTo 0
        
        mblnOK = True: Unload Me: Exit Sub
    Else '�������뵥��״̬
        If mobjBill.����ID = 0 Or mrsInfo.State = 0 Then
            MsgBox "û�з��ֲ�����Ϣ�����ݲ��ܱ��档", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mobjBill.Details.Count = 0 Then
            MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Checkִ�п���
        If i <> 0 Then
            MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.��������ID = 0 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus: Exit Sub
        End If
        
        If mobjBill.������ = "" Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        '��ʿ���:�жϷǷ�����
'        If HaveStopClass > 0 Then
'            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
'            Exit Sub
'        End If
                
        '����ʱ����
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '��Ժǿ�Ƽ���Ȩ�޼��
        If mint������Դ = 2 Then
            If Not PatiCanBilling(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0), mstrPrivs, pҽ�����ѹ���) Then Exit Sub
            If zlPatiIS�����ѱ�Ŀ(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0)) Then Exit Sub
            '49501:סԺ
            If zlIsAllowFeeChange(mrsInfo!����ID, Val(Nvl(mrsInfo!��ҳID))) = False Then Exit Sub
        End If
        
        '���˺� ����:?? ����:2010-01-07 10:37:09
        If zlCheck����ҽ��(Val(Nvl(mrsInfo!����))) = False Then Exit Sub
        
        If Check������� > 0 Then Exit Sub
        
        '����ʱ����
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "ǿ�ƶԳ�Ժ���˼���ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "���õķ���ʱ�䲻��С�ڲ��˵���Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '�Ƿ���
        dbl���� = 0
        For i = 1 To mobjBill.Details.Count
            '27467,52828
            If mobjBill.Details(i).���� <> 0 And dbl���� = 0 Then
                dbl���� = mobjBill.Details(i).����
            End If
            If mobjBill.Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
        Next
        '27467,52828
        If mbytInState = 0 And FormatEx(dbl����, 7) = 0 Then
            MsgBox "����������Ҫ��һ����Ϊ�������,���飡", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '����ְ����
        '����:45605
        If zlIsCheckMedicinePayMode(txt���ʽ) Then
            i = CheckDuty(, False)
            If i > 0 Then
                Bill.Row = i: Bill.MsfObj.TopRow = i
                Bill.Col = BillCol.��Ŀ: Bill.SetFocus
                Exit Sub
            End If
        End If

        '���в�����Ŀ
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.��Ŀ: Bill.SetFocus
            Exit Sub
        End If
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
                
        '���ʷ��౨��
        If mint��¼���� = 2 And mrsWarn.State = 1 And mstrWarn <> "-" Then
            '���ݷ���
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                'ˢ�²��˷���״��
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(mint������Դ = 1, 0, mlng��ҳID), mcurModiMoney)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!Ԥ�����
                    cmdCancel.Tag = rsTmp!�������
                    txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
                End If
                sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
                sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + curTotal, gSysPara.Money_Decimal.strFormt_VB)
                sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - curTotal, "0.00")
                
                '���¶�ȡ���ն�
                cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
                
                '�Ƿ�ҽ������
                cur��� = Val(txtʵ��.Tag)
                If gSysPara.bln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(IIf(mint������Դ = 0, 0, 1), mrsInfo!����ID)
                
                If mbln���õǼ� = False Then    '30504
                    For i = 1 To mobjBill.Details.Count
                        mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), cur���, cur���ն� - mcurModiMoney, curTotal, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn, mintWarn)
                        If mbytWarn = 2 Or mbytWarn = 3 Then Exit Sub
                    Next
                End If
            End If
        End If
        
        If mint��¼���� = 2 And Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� Then
            If gclsInsure.CheckItem(mrsInfo!����, 1, 2, MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), 2, 0)) = False Then
                Exit Sub
            End If
        End If
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, pҽ�����ѹ���, IIf(mint������Դ = 2, 1, 0), 1, _
            MakeDetailRecord(mobjBill, NeedName(cbo������.Text), NeedName(cbo��������.Text), mint��¼����, IIf(mint��¼���� = 1, 1, 0))) = False Then
            Exit Sub
        End If
        
        '��鵥�ݺ�ҽ�����Ƿ�ֻ�и�������,���ֻ�и�������,ֱ���˳�
        If CheckMainOperation() = False Then Exit Sub
        
        'ҩƷ���ɼ��
        strInfo = CheckDisable(mobjBill)
        If strInfo <> "" Then
            If strInfo Like "*(�������)*" Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            Else
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
                    
        '�����������
        If Not CheckLimit(mobjBill, , mblnҩ����λ) Then Exit Sub
        
        '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.���� Or .Detail.���) _
                    And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .ִ�в���ID = mobjBill.Details(j).ִ�в���ID Then
                            If .�շ���� = "4" Then
                                If .Detail.���� = mobjBill.Details(j).Detail.���� And .Detail.���� > 0 Then
                                    MsgBox "�� " & j & " �з�����������""" & .Detail.���� & """ ��ͬһ�����ϲ��ű��ظ�������ͬ���Σ���ϲ���", vbInformation, gstrSysName
                                    Exit Sub
                                ElseIf .Detail.���� <= 0 Then
                                    MsgBox "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����룬��ϲ���", vbInformation, gstrSysName
                                    Exit Sub
                                End If
                            Else
                                MsgBox "�� " & j & " �еķ�����ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����룬��ϲ���", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            
                        End If
                    Next
                End If
            End With
        Next
        
        'ҩƷ�����(�������ֹʱ�����ʱ��ҩƷ)
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .�շ����) > 0 Then
                    If .Detail.���� Or .Detail.��� Then
                        '�����:113637,����,2017/09/15,�ڱ�������֮ǰ��Ҫ���ҩƷ���,����GetDrugTotal����
                        '��ȡҩƷ���ʱû�д�������,���¿����ʧЧ
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0) '29680
                        If mblnҩ����λ Then
                            .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        End If
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0) '29680
                        If mblnҩ����λ Then
                            .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        End If
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0, , , .Detail.����) '29680
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, 0, , , .Detail.����) '29680
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ����������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Next
    
        '���ۼ��,105875
        If Not mobjPublicDrug Is Nothing Then
            'Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
            '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
            '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
            'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
            '���۳���ʱֻ�ж�ҩ��
            '���أ�True-�����������۳��⣻false-���ܽ������۳���
            For i = 1 To mobjBill.Details.Count
                With mobjBill.Details(i)
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        If mobjPublicDrug.zlCheckPriceAdjustBySell(.�շ�ϸĿID, .ִ�в���ID) = False Then
                            Exit Sub
                        End If
                    End If
                End With
            Next
        End If
        
        '����������ϵ����Ч��
        '����Զ���ҩ:25490
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .�շ���� = "4" And .Detail.�������� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                    If Not CheckValidity(.�շ�ϸĿID, .ִ�в���ID, dblTotal) Then Exit Sub
                ElseIf InStr(1, ",5,6,7,", .�շ����) > 0 Then
                    '��ӡ��ҩ��,����ͨ����,�һ��۵�����
                    If mbytSendMateria <> 0 And mint��¼���� = 2 And mint������Դ = 2 Then
                        'ȫ��ҩƷ��ȷ����ҩ���Ĳ��Զ���ҩ(���뷢ҩʱ,û��ȷ��ҩ��)
                        mblnSendMateria = .ִ�в���ID <> 0
                    End If
                End If
            End With
        Next
        If InStr(mstrPrivs, ";ҩƷ��ҩ;") = 0 Then mblnSendMateria = False
        
        If mblnSendMateria And mbytSendMateria = 2 Then
            If MsgBox("������ɺ��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        '�����˷Ѽ��
        If mint��¼���� = 2 Then
            If Not CheckNegative Then Exit Sub
        End If
        
        'ˢ�������鿨
        If mint������Դ = 1 And mint��¼���� = 2 And gSysPara.dblԤ��������鿨 <> 0 Then
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                If Not gobjDatabase.PatiIdentify(Me, glngSys, mobjBill.����ID, curTotal, _
                    , , , IIf(-1 * gSysPara.dblԤ��������鿨 >= curTotal, False, True), , , , (gSysPara.dblԤ��������鿨 = 2)) Then Exit Sub
            End If
        End If
        '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
        If mobjSquareCard Is Nothing Then
            If mint������Դ = 1 And (gSysPara.bln�������������� Or mbln���֧��) Then
                If MsgBox("ע�⣺" & vbCrLf & "      ҽ�ƿ�������zl9CardSquare��δ���������������󽫲��ܽ����շѻ������ˣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
        mobjBill.�Ǽ�ʱ�� = gobjDatabase.Currentdate
        If zlGetSaveDataItems_Plugin(mobjBill, rsItems) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(pҽ�����ѹ���, mint��¼����, False, mint��¼���� = 1, "", rsItems) = False Then Exit Sub
        
        If Not SaveBill(strNos) Then Exit Sub
        
        Call zlChargeSaveAfter_Plugin(pҽ�����ѹ���, mlng����ID, mlng��ҳID, mint������Դ = 1, mint��¼����, strNos)
        '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
        If mint������Դ = 1 And (gSysPara.bln�������������� Or mbln���֧��) And strNos <> "" Then
            If Not mobjSquareCard Is Nothing Then
                Call mobjSquareCard.zlSquareAffirm(Me, pҽ�����ѹ���, mstrPrivs, mlng����ID, , , mint��¼����, strNos)
            End If
        End If
        
        '���˺�:��ӡ��ҩ��:25490
        If mblnSendMateria Then
            If InStr(1, mstrPrivs, ";��ҩ�嵥��ӡ;") > 0 Then
                If MsgBox("����""" & mobjBill.NO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
        End If
        
        If mint������Դ = 1 And mint��¼���� = 2 Then
            '110319
            If mblnDrugMachine Then
                If mstrInNO <> "" Then
                    '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
                    strSql = "Select Id As ����id, -1 * Nvl(����, 1) * ���� As ��ҩ����" & vbNewLine & _
                            " From ������ü�¼" & vbNewLine & _
                            " Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1] And �շ���� In ('5', '6', '7')" & vbNewLine & _
                            "       And �Ǽ�ʱ�� + 0 = (Select Max(�Ǽ�ʱ��)" & vbNewLine & _
                            "                       From ������ü�¼" & vbNewLine & _
                            "                       Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1])"
                    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ѯ�����˷���Ŀ", mstrInNO)
                    Do While Not rsTemp.EOF
                        strData = strData & ";" & Nvl(rsTemp!����id) & "," & Nvl(rsTemp!��ҩ����)
                        rsTemp.MoveNext
                    Loop
                    If strData <> "" Then
                        strData = Mid(strData, 2)
                        Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
                    End If
                End If
                
                '�����ʽ��1|����1,������1;����2,������2
                strData = "1|" & "9," & Replace(Replace(strNos, "'", ""), ",", ";9,")
                Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
            End If
        End If
        
        If mstrInNO <> "" Or mstrOriginalNO <> "" Then
            mblnOK = True: Unload Me: Exit Sub
        Else
            txtPreNO.Text = mobjBill.NO
            Set mrsMainOperation = Nothing
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            Call SetMoneyList
            Call NewBill
            
            '���¶�ȡ������Ϣ
            Call GetPatient(mlng����ID, mlng��ҳID)
            Bill.SetFocus
        End If
    End If
    mblnOK = True
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function Check�������() As Integer
'���ܣ���鵱ǰ���˵ļ��ʷ�����Ŀ�ķ�������Ƿ�һ��
'˵������Ϊ�������������۲���,�����д˼��
'���أ���һ�µķ�����,Ϊ0ʱ����
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    For i = 1 To mobjBill.Details.Count

        If mobjBill.Details(i).Detail.������� = 1 And mint������Դ <> 1 Then
            MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """������������,�ò��˲���ʹ��.", vbInformation, gstrSysName
            Check������� = i: Exit Function
        End If

        If mobjBill.Details(i).Detail.������� = 2 And mint������Դ <> 2 Then
            MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """��������סԺ,�ò��˲���ʹ��.", vbInformation, gstrSysName
            Check������� = i: Exit Function
        End If

        If mobjBill.Details(i).Detail.������� = 0 Then
            MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�������ڲ���,�ò��˲���ʹ��.", vbInformation, gstrSysName
            Check������� = i: Exit Function
        End If
    Next
End Function

Private Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str������ As String, ByVal str�������� As String, _
    ByVal intMode As Integer, ByVal intPrice As Integer, Optional ByVal lngRow As Long) As ADODB.Recordset
'���ܣ����ݵ��ݶ������ݴ���һ����ϸ��¼����Ϣ(���ۼ۵�λ)
'�ֶΣ�����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������,ִ�п���ID,
'         �������ʣ�1-�շѵ�,2-���ʵ�),�Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)
'������intPage=ָ���ĵ���,lngRow=ָ�����У���ָ��ʱ�������е��ݵ�������
    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl���� As Double, curʵ�� As Currency
    Dim rsTmp As New ADODB.Recordset
    
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    '79420,���ϴ�,2014/11/10:������¼���ֶδ�С
    rsTmp.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    '131048,����,2018-9-5,��CheckChargeItem���ӿ��е�rsDetail ��¼���������ֶ�:
    '                                  ִ�п���ID���������ʣ�1-�շѵ�,2-���ʵ�)���Ƿ񻮼�(1-����,0-�������շѼ����ʵ�)
    rsTmp.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��������", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�Ƿ񻮼�", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    If lngRow = 0 Then
        intB = 1
        intE = objBill.Details.Count
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl���� = 0: curʵ�� = 0
        With objBill.Details(i)
            If lngRow = 0 Then
                rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID
                blnNew = rsTmp.RecordCount = 0
            Else
                blnNew = True
            End If
                            
            If blnNew Then
                rsTmp.AddNew
                
                rsTmp!����ID = objBill.����ID
                rsTmp!��ҳID = objBill.��ҳID
                
                rsTmp!�շ���� = .�շ����
                rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                rsTmp!ִ�п���ID = .ִ�в���ID
                rsTmp!�������� = intMode
                rsTmp!�Ƿ񻮼� = intPrice
                
                For j = 1 To .InComes.Count
                    dbl���� = dbl���� + .InComes(j).��׼����
                    curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                Next
                If InStr(",5,6,7,", .�շ����) > 0 And mblnҩ����λ Then
                    '��ҩ����λת��Ϊ�ۼ۵�λ
                    rsTmp!���� = IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                    rsTmp!���� = Format(dbl���� / .Detail.ҩ����װ, gSysPara.Price_Decimal.strFormt_VB)
                Else
                    rsTmp!���� = IIf(.���� = 0, 1, .����) * .����
                    rsTmp!���� = Format(dbl����, gSysPara.Price_Decimal.strFormt_VB)
                End If
                rsTmp!ʵ�ս�� = Format(curʵ��, gSysPara.Money_Decimal.strFormt_VB)
                
                rsTmp!������ = str������
                rsTmp!�������� = str��������
            Else
                For j = 1 To .InComes.Count
                    dbl���� = dbl���� + .InComes(j).��׼����
                    curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                Next
                If InStr(",5,6,7,", .�շ����) > 0 And mblnҩ����λ Then
                    '��ҩ����λת��Ϊ�ۼ۵�λ
                    rsTmp!���� = rsTmp!���� + IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                    rsTmp!���� = Format((rsTmp!���� + Format(dbl���� / .Detail.ҩ����װ, gSysPara.Price_Decimal.strFormt_VB)) / 2, gSysPara.Price_Decimal.strFormt_VB)
                Else
                    rsTmp!���� = rsTmp!���� + IIf(.���� = 0, 1, .����) * .����
                    rsTmp!���� = Format((rsTmp!���� + Format(dbl����, gSysPara.Price_Decimal.strFormt_VB)) / 2, gSysPara.Price_Decimal.strFormt_VB)
                End If
                rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + Format(curʵ��, gSysPara.Money_Decimal.strFormt_VB)
            End If
            
            rsTmp.Update
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub cmdSaveAndPay_Click()
    mbln���֧�� = True
    Call cmdOK_Click
    mbln���֧�� = False
End Sub

Private Sub Form_Activate()
    Dim objTemp As Object
    If mbytInState <> 0 Then
        If cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus
        ElseIf cmdCancel.Visible And cmdCancel.Enabled Then
            cmdCancel.SetFocus
        End If
    End If
    '101218
    If mblnSetControl Then
        mblnSetControl = False
        Set objTemp = Me.ActiveControl
        If cboTemp.Visible And cboTemp.Enabled Then cboTemp.SetFocus
        If Not objTemp Is Nothing Then
            If TypeName(objTemp) = "BillEdit" Then
                'BillEdit�ؼ����⴦��
                If objTemp.Visible And objTemp.Active Then objTemp.SetFocus
            Else
                If objTemp.Visible And objTemp.Enabled Then objTemp.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is Bill Then Exit Sub
    If Me.ActiveControl Is txtBarCode Then Exit Sub
    If InStr("',;|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill, i As Long
    gSysPara.strLike = IIf(gobjDatabase.GetPara("����ƥ��") = "0", "%", "")
    gSysPara.bytCode = Val(gobjDatabase.GetPara("���뷽ʽ"))
    mblnSetControl = True
    mblnWarnCloseed = False
    glngFormW = 12000: glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName, mbytInState)
    mblnEnterCell = True
    mintWarn = -1: mstrWarn = ""
    mstrFeeTab = IIf(mint������Դ = 2 And mint��¼���� = 2, "סԺ���ü�¼", "������ü�¼")
    Call InitLocPar
    
    mbln���֧�� = False
    If mbytInState = 0 Then
        If CheckValid = False Then
            Unload Me: Exit Sub
        End If
        If CreatePublicDrug(glngSys, gcnOracle, gstrDBUser) = False Then
            Unload Me: Exit Sub
        End If
        
        If GetPriceGrade(gstrNodeNo, mlng����ID, mlng��ҳID, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�) = False Then
            mstrҩƷ�۸�ȼ� = "": mstr���ļ۸�ȼ� = "": mstr��ͨ�۸�ȼ� = ""
        End If
    End If
    Call CreateDrugPacker
    
    '��ʼ����������
    Set mobjBill = New ExpenseBill
    If mbytInState = 0 Then
        If Not InitData Then
            Unload Me: Exit Sub
        End If
    End If
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then
        If Not ReadBill(mstrInNO, mbytInState = 3) Then
            Unload Me: Exit Sub
        End If
    Else
        Call CreatePlugIn(pҽ�����ѹ���)
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���
            Set mobjBill = ImportBill(mint������Դ, mstrInNO, mint��¼����, mbln���õǼ�, _
                mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            If mobjBill.NO = "" Then
                MsgBox "������ȷ��ȡ�Ʒѵ��ݵ����ݣ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                
                '����б༭����������ɫ
                Bill.SetColColor BillCol.���, &HE7CFBA
                Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
                Bill.SetColColor BillCol.����, &HE7CFBA
                Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
                Bill.SetColColor BillCol.����, &HE0E0E0
                Bill.SetColColor BillCol.����, &HE0E0E0
                Bill.SetColColor BillCol.��־, &HE0E0E0
                
                cboNO.Text = mobjBill.NO
                               
                
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                If mint��¼���� = 2 Then
                    mcurModiMoney = GetBillMoney(mobjBill.NO) '�ڶ�ȡ����ǰȡ
                End If
                
                '�µ�ʱ��ȡ����,������ʱ���ݵ�����ʾ������Ϣ
                Call GetPatient(mlng����ID, mlng��ҳID)
                If mrsInfo.State = 0 Then
                    If Not mblnWarnCloseed Then
                        MsgBox "���ܶ�ȡ������Ϣ���������㲻���жԸò��˼Ʒѵ�Ȩ�ޡ�", vbInformation, gstrSysName
                    End If
                    Unload Me: Exit Sub
                End If
                
'                Call FindCboIndex(cbo��������, mobjBill.��������ID, False)
                Call Find��������(mobjBill.��������ID)
                Call GetCboIndex(cbo������, mobjBill.������)
                Call gobjControl.CboLocate(cboBaby, mobjBill.Ӥ����, True)
                
                If gSysPara.bln��������ۿ� Then CalcMoneys
                Call ShowDetails
                Call ShowMoney
                
                '�������:�޸�ʱ���Ͻ�Ҫ�˻صĿ��
                For i = 1 To mobjBill.Details.Count
                    With mobjBill.Details(i)
                        Bill.RowData(i) = Asc(.�շ����) '���⴦��
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            .Detail.��� = .Detail.��� + .���� * .����
                        ElseIf .�շ���� = "4" And .Detail.�������� Then
                            .Detail.��� = .Detail.��� + .���� * .����
                        End If
                    End With
                Next
                
                Call SetColNum
            End If
        Else
            '�µ�ʱ��ȡ����,������ʱ���ݵ�����ʾ������Ϣ
            Call GetPatient(mlng����ID, mlng��ҳID)
            If mrsInfo.State = 0 Then
                If Not mblnWarnCloseed Then
                    MsgBox "���ܶ�ȡ������Ϣ���������㲻���жԸò��˼Ʒѵ�Ȩ�ޡ�", vbInformation, gstrSysName
                End If
                Unload Me: Exit Sub
            End If
            If Not IsNull(mrsInfo!����) Then
                MCPAR.�������� = gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����)
                MCPAR.�����ϴ� = gclsInsure.GetCapability(support�����ϴ�, mrsInfo!����ID, mrsInfo!����)
                MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, mrsInfo!����ID, mrsInfo!����)
                MCPAR.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mrsInfo!����ID, mrsInfo!����)
                MCPAR.ҽ��ȷ���������� = gclsInsure.GetCapability(supportҽ��ȷ����������, mrsInfo!����ID, mrsInfo!����)
            End If
        End If
        
        If mstrInNO <> "" And mint��¼���� = 2 And mint������Դ = 2 Then
            Call ReCalcInsure '���¼���ͳ����
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraBarCode.Width = Me.ScaleWidth - fraBarCode.Left
    txtBarCode.Width = ScaleWidth - txtBarCode.Left - 100
    
    Bill.Top = IIf(Not mblnShowBarCode, fraInfo.Top + fraInfo.Height, fraBarCode.Top + fraBarCode.Height) + 50
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - Bill.Top - 25
    
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    
    cboNO.Left = fraTitle.Width - cboNO.Width - 90
    lblNO.Left = cboNO.Left - lblNO.Width - 45
        
    fraUnit.Left = Me.ScaleWidth - fraUnit.Width
    fraInfo.Width = Me.ScaleWidth - fraUnit.Width - fraInfo.Left
    
    Bill.Width = Me.ScaleWidth - Bill.Left
    
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    fraDrawDept.Width = fraAppend.Width
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
        
    cbo������.Left = lblDate.Left - cbo������.Width - 200
    lbl������.Left = cbo������.Left - lbl������.Width - 45
    
    If (Me.ScaleWidth + fraStat.Left + fraStat.Width - cmdOK.Width) / 2 < fraStat.Left + fraStat.Width + 500 Then
        cmdOK.Left = fraStat.Left + fraStat.Width + 500
    Else
        cmdOK.Left = (Me.ScaleWidth + fraStat.Left + fraStat.Width - cmdOK.Width) / 2
    End If
    cmdSaveAndPay.Left = cmdOK.Left
    cmdCancel.Left = cmdOK.Left
    If cmdSaveAndPay.Tag = "0" Then
        cmdOK.Top = cmdSaveAndPay.Top - cmdSaveAndPay.Height / 3 * 2
        cmdCancel.Top = cmdSaveAndPay.Top + cmdSaveAndPay.Height / 3 * 2
    End If

    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mbytInState)
    
    mlngҽ��ID = 0
    mlng���ͺ� = 0
    mlng����ID = 0
    mlng��ҳID = 0
    mint������Դ = 0
    mint��¼���� = 0
    mbln���õǼ� = False
    mlng��������ID = 0
    mlng���˿���id = 0
    
    mlng��������ID = 0
    mstr����ҽ�� = ""
    mstrOriginalNO = ""
    
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    
    mbytInState = 0
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mstrPrivs = ""
    mstrRegNo = ""
    
    Set mrsInfo = Nothing
    Set mrsUnit = Nothing
    Set mrsClass = Nothing
    Set mrsWork = Nothing
    Set mrsMedAudit = Nothing
    Set mobjSquareCard = Nothing
    Set mobjCard = Nothing
    Set mobjBrushCheck = Nothing
    Set mobjPublicDrug = Nothing
    Set mrsMainOperation = Nothing
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
End Sub


Private Sub picAppend_Resize()
    Err = 0: On Error Resume Next
    With picAppend
        txt���˱�ע.Width = .ScaleWidth - txt���˱�ע.Left - 100
    End With
End Sub

Private Sub mobjBrushCheck_ReadCardNoed(ByVal strCardNO As String, ByVal blnBrushCard As Boolean)
    If blnBrushCard Then
        mbln����ˢ�� = True
    Else
        mbln����ˢ�� = False
    End If
End Sub

Private Sub txtBarCode_GotFocus()
    gobjCommFun.OpenIme False
    gobjControl.TxtSelAll txtBarCode
End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode <> vbKeyReturn Then Exit Sub
   
   If AddStuffItemFromBarCode(txtBarCode.Text) = False Then
        If txtBarCode.Enabled And txtBarCode.Visible Then txtBarCode.SetFocus
        gobjControl.TxtSelAll txtBarCode: Exit Sub
   End If
   txtBarCode.Text = ""
   If txtBarCode.Enabled And txtBarCode.Visible Then txtBarCode.SetFocus
End Sub
 
Private Sub txtBarCode_LostFocus()
    gobjCommFun.OpenIme False
End Sub

Private Sub ShowAndHideBarCodeInput()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���������������
    '����:���˺�
    '����:2017-11-22 11:42:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    fraBarCode.Visible = mblnShowBarCode
    txtBarCode.Visible = mblnShowBarCode
    lblBarCode.Visible = mblnShowBarCode
    Call Form_Resize
 End Sub
 
Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Key
    Case "BarCode"
        '��ʾ����
        mblnShowBarCode = Not mblnShowBarCode
        Panel.Bevel = IIf(Not mblnShowBarCode, sbrRaised, sbrInset)
        Panel.ToolTipText = IIf(Not mblnShowBarCode, "��ʾ���������", "�������������")
        Call ShowAndHideBarCodeInput
        If txtBarCode.Enabled And txtBarCode.Visible Then txtBarCode.SetFocus
        Call gobjDatabase.SetPara("�ϴ�ѡ���������", IIf(mblnShowBarCode, 1, 0), glngSys, pҽ�����ѹ���)
        Exit Sub
    Case "PY", "WB"
        If Not gSysPara.bln����ƥ�䷽ʽ�л� Then Exit Sub
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call gobjDatabase.SetPara("���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0)))
   End Select

End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call gobjCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long
    If mbytInState = 3 Then
        Bill.Row = 1: Call gobjCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        If mbytInState <> 2 Then
            .ColData(BillCol.���) = IIf(gSysPara.bln�շ����, BillColType.ComboBox, BillColType.UnFocus) '�����,��������ʱ�ᱻ�ı�
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
            .ColData(BillCol.����) = 5 '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = 5 '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.��־) = 5 '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        End If
        '����б༭����������ɫ
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        
        .TextMatrix(Row, BillCol.��) = Row
        
        '����ط��ֶ����ò�ִ��
        If Row > 0 And .ColData(BillCol.���) <> 5 And Me.Visible And Not mblnNewRow Then
            Call gobjCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
     '���˺� ����:27378 ����:2010-01-27 16:20:02
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo��������.ListIndex <> -1 Then
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If cbo��������.Locked Then Exit Sub
    If mrsAll�������� Is Nothing Then Exit Sub
    If InStr(mstrPrivs, ";���п���;") = 0 Then Exit Sub

    If zlSelectDept(Me, 1150, cbo��������, mrsAll��������, cbo��������.Text, True, , , True) = False Then
        mobjBill.��������ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
 
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        If cbo������.Locked Then
            Call gobjCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = cbo������.Text
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then
                Call gobjControl.CboSetIndex(cbo������.hWnd, -1)
            Else
                gobjCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            If IsNumeric(strText) Then
                For i = 0 To cbo������.ListCount - 1
                    If i > UBound(marrDr) Then Exit For
                    If CStr(Split(marrDr(i), "|")(2)) = strText Then
                        If intIdx = -1 Then cbo������.ListIndex = i
                        intIdx = i
                    End If
                Next
                If intIdx = -1 Then
                    For i = 0 To cbo������.ListCount - 1
                        If i > UBound(marrDr) Then Exit For
                        If Val(Split(marrDr(i), "|")(2)) = Val(strText) Then
                            If intIdx = -1 Then cbo������.ListIndex = i
                            intIdx = i
                        End If
                    Next
                End If
            Else
                For i = 0 To cbo������.ListCount - 1
                    If UCase(cbo������.List(i)) Like UCase(strText) & "*" Then
                        If intIdx = -1 Then cbo������.ListIndex = i
                        intIdx = i
                    End If
                Next
            End If
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Call gobjCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.������ = ""
        Else
            mobjBill.������ = NeedName(cbo������.Text)
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Call gobjCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
        Call gobjCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp "zl9InExse", Me.hWnd, "frmCharge"
        Case vbKeyF2
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If Shift = vbShiftMask Then
                If cmdSaveAndPay.Enabled And cmdSaveAndPay.Visible Then Call cmdSaveAndPay_Click
            Else
                If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
            End If
        Case vbKeyF6 '�����ǰ��������,�����µ�״̬
            If mbytInState = 0 Then
                txtʵ��.Text = gSysPara.Money_Decimal.strFormt_VB: txtӦ��.Text = gSysPara.Money_Decimal.strFormt_VB
                Call ClearRows: Call Bill.ClearBill
                Call SetColNum: Call ClearMoney
                Call NewBill
                Bill.SetFocus
            End If
        Case vbKeyF7 '�л����뷨
            If Not gSysPara.bln����ƥ�䷽ʽ�л� Then Exit Sub
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape, vbKeyX
            If KeyCode = vbKeyX And Shift <> 4 Then Exit Sub
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Sub SetMoneyList()
'����:���ݵ�ǰ������Ŀ�����������п�
    Dim lngW As Long
    lngW = mshMoney.Width - 60
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    mshMoney.ColWidth(0) = lngW * 0.5
    mshMoney.ColWidth(1) = lngW * 0.5
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    
    mshMoney.TextMatrix(0, 0) = "��Ŀ"
    mshMoney.TextMatrix(0, 1) = "���"
    mshMoney.Row = 0
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strOperDoc As String
    
    On Error GoTo errH
    
    '��ͬҩ��ҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '��������
    If mbytӦ�ó��� = 0 Or mlngҽ��ID <> 0 Then
        strSql = "Select ��������ID,����ҽ��,������Դ,Decode(Nvl(������Դ,0),4,'',�Һŵ�) As �Һŵ� From ����ҽ����¼ Where ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID)
        If Not rsTmp.EOF Then
            mlng��������ID = Nvl(rsTmp!��������id, 0)
            mstr����ҽ�� = Nvl(rsTmp!����ҽ��)
            mstrRegNo = Nvl(rsTmp!�Һŵ�)   '99842,��첡�˹Һŵ��������쵥��
            If Val(Nvl(rsTmp!������Դ, 0)) = 4 Then
                mbytӦ�ó��� = 1 '�������첡�ˣ����޸�Ӧ�ó��ϣ��Ա���д���۵�������ԴΪ���
            End If
        End If
        If mlng��������ID = 0 Or mstr����ҽ�� = "" Then
            MsgBox "û�з���Դҽ����Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    
    ElseIf mlng��������ID = 0 Then
        If mstrInNO <> "" Then
            strSql = _
            " Select  A.��������ID ,A.����ID" & IIf(mstrFeeTab = "סԺ���ü�¼", ",A.��ҳID ", ",0 as ��ҳID") & _
            " From " & mstrFeeTab & " A " & _
            " Where Rownum<2 And  NO=[1] And A.��¼����=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mstrInNO, mint��¼����)
            If rsTmp.EOF = False Then
                mlng��������ID = Val(Nvl(rsTmp!��������ID))
                mlng����ID = Val(Nvl(rsTmp!����ID))
                mlng��ҳID = Val(Nvl(rsTmp!��ҳID))
            End If
            If mlng��������ID = 0 Then
                MsgBox "δ�ҵ����ݺ�:" & mstrInNO & "�ļ�¼.", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "ָ���Ŀ������Ҳ����ڡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
      
    strSql = "Select A.ID, A.����, A.����, A.����, 0 As ȱʡ, B.��������, D.���ȼ�," & vbNewLine & _
            "        Row_Number() Over(Partition By ID Order By ��������) As R" & vbNewLine & _
            " From ���ű� A, ��������˵�� B," & vbNewLine & _
            "      (Select ����id, Max(Decode(�������, 2, 1, 2)) As ���ȼ� From ��������˵�� Where ������� <> 0 Group By ����id) D" & vbNewLine & _
            " Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And A.ID = B.����id" & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "       And B.����id = D.����id And (B.������� IN(1,2,3) AND B.�������� IN('�ٴ�','����') Or b.��������='����')"
    '87562,ÿһ������ֻ��ѯ��һ����¼������"��������"�ϲ���ʾ,��:������,��������������"�ٴ�"��"����",���ѯ���Ĺ��������ֶ�ֵΪ"�ٴ�,����"
    strSql = "Select ID, ����, ����, ����, ȱʡ, ���ȼ�, ��������" & vbNewLine & _
            " From (Select Row_Number() Over(Partition By ID Order By Lvl Desc) As Rn, ID, ����, ����, ����, ȱʡ, ���ȼ�, ��������" & vbNewLine & _
            "        From (Select ID, ����, ����, ����, ȱʡ, ���ȼ�, Level Lvl, LTrim(Sys_Connect_By_Path(��������, ','), ',') As ��������" & vbNewLine & _
            "               From (" & strSql & ")" & vbNewLine & _
            "               Connect By ID = Prior ID And R - 1 = Prior R))" & vbNewLine & _
            " Where Rn = 1" & vbNewLine & _
            " Order By ���ȼ�,����"
    Set mrsAll�������� = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '69912:������,2014-02-12,�������������б���������ҽ������
    strOperDoc = Getҽ����������(mlngҽ��ID, "����ҽ������")
    
    If mbln���õǼ� Then
        '��Ϊ��ǰѡ���ҽ������
        strSql = "(Select ID,����,����,���� From ���ű� Where ID=[1]"
    Else
        '��Ϊ��ǰѡ���ҽ�����һ�������
        strSql = "(Select ID,����,����,���� From ���ű� Where ID IN([1],[2])"
    End If
    
    If strOperDoc <> "" Then
        strSql = strSql & " Union " & _
                "Select ID,����,����,���� From ���ű� Where ����=[3]"
    End If
    
    strSql = strSql & ") Order By ����"
    Set mrsDept = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlng��������ID, mlng��������ID, strOperDoc)
    
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cbo��������.AddItem IIf(zlIsShowDeptCode, mrsDept!���� & "-", "") & mrsDept!����
            cbo��������.ItemData(cbo��������.ListCount - 1) = mrsDept!ID
            If mbytȱʡ���� = 0 Then    'ȱʡҽ������:36060
                If mrsDept!ID = mlng��������ID Then
                    cbo��������.ListIndex = cbo��������.NewIndex
                End If
            Else
                If mrsDept!ID = mlng���˿���id Then
                    cbo��������.ListIndex = cbo��������.NewIndex
                End If
            End If
            mrsDept.MoveNext
        Next
        If InStr(mstrPrivs, ";���п���;") > 0 Then
            cbo��������.AddItem "�������ҡ�"
            cbo��������.ItemData(cbo��������.ListCount - 1) = 0
        End If
        If cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    Else
        MsgBox "����ȷ���������ң����ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����շ����:"'5','E','Z'"
    If mstr�շ���� = "" Then
        strSql = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
    Else
        strSql = "Select ����,���� as ��� From �շ���Ŀ��� Where Instr([1],����)>0 Order by ���"
    End If
    'Set mrsClass = New ADODB.Recordset
    Set mrsClass = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mstr�շ����)
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ����,�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    '��ֻ��һ�ֿ�ѡ�շ����ʱ,�����û�ѡ��
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(mstr�շ����, "'5'") > 0 Or InStr(mstr�շ����, "'6'") > 0 Or InStr(mstr�շ����, "'7'") > 0 Or mstr�շ���� = "" Then
        mlngҩƷ���ID = ExistIOClass(IIf(mint��¼���� = 1, 8, 9))
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ��ҩƷ���ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(mstr�շ����, "'4'") > 0 Or mstr�շ���� = "" Then
        mlng�������ID = ExistIOClass(IIf(mint��¼���� = 1, 40, 41))
        If mlng�������ID = 0 Then
            MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'ִ�в���
    strSql = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID and B.������� IN([1],3) " & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by B.�������,A.����"
    'Set mrsUnit = New ADODB.Recordset
    Set mrsUnit = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mint������Դ)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    InitData = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str��� As String, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'���ܣ���ȡ����������ͬ�����Ŀ��ִ�п���ID
    Dim i As Long
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).�շ���� = str��� _
            And mobjBill.Details(i).ִ�в���ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                Exit Function
            End If
        End If
    Next
    If str��� = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).ִ�в���ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, Optional blnEnter As Boolean)
'���ܣ����ݵ��������������б������
'������blnEnter=�Ƿ񰴽�����д���,����ִ�п��ұ��ֲ���
    Dim rsTmp As New ADODB.Recordset
    Dim str��Ա���� As String, strTmp As String
    Dim lng����id As Long, strIDs As String
    Dim strSql As String, i As Long
    Dim rsUnit As ADODB.Recordset
    Bill.Clear
    On Error GoTo errH
    Select Case Bill.TextMatrix(0, lngCol)
        Case "���"
            Bill.cboStyle = DropOlnyDown
            If cbo������.ListIndex <> -1 Then
                If cbo������.ListIndex <= UBound(marrDr) Then
                    If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 6 Then
                        str��Ա���� = Split(marrDr(cbo������.ListIndex), "|")(6)
                    End If
                End If
            End If
        
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                For i = 1 To mrsClass.RecordCount
                    '��ʿ���:����
'                    If Not (str��Ա���� = "��ʿ" And InStr(",E,M,4,", mrsClass!����) = 0) Then
                        Bill.AddItem Bill.ListCount + 1 & "-" & mrsClass!���
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!����)  '����������ASCII��
'                    End If
                    mrsClass.MoveNext
                Next
            End If
        Case "ִ�п���"
            Bill.cboStyle = DropDownAndEdit
            
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .�շ����) > 0 Then
                        Call GetWorkUnit(.�շ�ϸĿID, .�շ����)
                        If mrsWork.RecordCount > 0 Then
                            'ȡ��һ��ҩ��ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                lng����id = GetLastDeptID(.�շ����, lngRow, Mid(strIDs, 2))
                            End If
                            If lng����id = 0 Then lng����id = .ִ�в���ID
                            
                            'ȷ����ǰ�е�ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!���� & "-", "") & mrsWork!����
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng����id Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        lng����id = Get��������ID
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        '0-����ȷ,1-���˿���,2-���˲���,3-�����˿���,4-ָ������
                        Select Case .Detail.ִ�п���
                            Case 0 '����ȷ
                                mrsUnit.Filter = 0
                                '101736,�ֹ�����ȱʡִ�п���
                                If mint������Դ = 2 And Not blnEnter Then
                                    '1 ������Ŀѡ���Ҵ���ȱʡ��ִ�п��ҵ� ������Ŀ��ִ�в���ID
                                    '   (���ﲻ�����ǳ�����Ŀ)
                                    '2 �շ���Ŀ.ȱʡ����(�ֹ�����ȱʡִ�п���)
                                    strSql = "Select a.ִ�п���id" & vbNewLine & _
                                            " From �շ�ִ�п��� A, ���ű� C" & vbNewLine & _
                                            " Where a.ִ�п���id + 0 = c.Id And a.�շ�ϸĿid = [1]" & vbNewLine & _
                                            "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                                            "       And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null)" & vbNewLine & _
                                            "       And a.������Դ = [2] And a.��������id Is Null"
                                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, .�շ�ϸĿID, 2)
                                    If Not rsTmp.EOF Then lng����id = Val(Nvl(rsTmp!ִ�п���ID))
                                    '3 ���˿���
                                    If lng����id = 0 Then lng����id = mobjBill.����ID
                                    '4 ��������
                                    If lng����id = 0 Then lng����id = Get��������ID
                                    '5 ����Ա��������ID
                                    If lng����id = 0 Then lng����id = UserInfo.����ID
                                End If
                            Case 1 '���˿���
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!����ID, 0) & " Or ID=" & .ִ�в���ID
                            Case 2 '���˲���
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!����ID, 0) & " Or ID=" & .ִ�в���ID
                            Case 3 '����Ա����
                                mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                            Case 4 'ָ������
                                strSql = "Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                                    " From �շ�ִ�п��� A,���ű� B" & _
                                    " Where A.ִ�п���ID=B.ID And A.�շ�ϸĿID=[1]" & _
                                    " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                                    " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                                    " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
                                    " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, .�շ�ϸĿID, mint������Դ, Val(Nvl(mrsInfo!����ID, 0)))
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .ִ�в���ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                                End If
                            Case 6 '�����˿���
                                mrsUnit.Filter = "ID=" & lng����id & " Or ID=" & .ִ�в���ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        Set rsUnit = rec.CopyNew(mrsUnit)
                        If Not rsUnit.EOF Then
                            For i = 1 To rsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, rsUnit!���� & "-", "") & rsUnit!����
                                '���˺�:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(rsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.NewIndex) = rsUnit!ID
                                    
                                    '����ȱʡִ�п���
                                    If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                        If lngRow = 1 Then
                                            If rsUnit!ID = lng����id Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '����һ�з�ҩƷ��ͬ
                                            If rsUnit!ID = mobjBill.Details(lngRow - 1).ִ�в���ID And mobjBill.Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).�շ����) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf rsUnit!ID = lng����id And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                rsUnit.MoveNext
                            Next
                        End If
                            
                        If Not blnEnter And .Detail.ִ�п��� = 4 Then 'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Long, arrBaby As Variant
    
    '���õ��ݱ��ʽ
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.��Ŀ
        .PrimaryCol = BillCol.��Ŀ
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.��) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 Then
            .ColData(BillCol.��) = 5
            
            .ColData(BillCol.���) = IIf(gSysPara.bln�շ����, 3, 5)
            If mblnOne Then .ColData(BillCol.���) = 5
            
            .ColData(BillCol.��Ŀ) = 1 '��Ŀ����,��Ť��ѡ
            .ColData(BillCol.����) = 4 '��/������
            '���˺�:27990 2010-02-23 12:04:37
            .ColData(BillCol.��Ʒ��) = 5 '�������
            .ColData(BillCol.���) = 5 '�������
            .ColData(BillCol.��λ) = 5 '��λ����
            .ColData(BillCol.����) = 5 '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = 5 '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.Ӧ�ս��) = 5 'Ӧ�ս������
            .ColData(BillCol.ʵ�ս��) = 5 'ʵ�ս������
            .ColData(BillCol.ִ�п���) = 3 'Ĭ��ȡ�������һ���һ����
            .ColData(BillCol.��־) = 5 '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
            .ColData(BillCol.����) = 5 '����ȱʡ����
        End If
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call gobjComlib.RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    
    Me.KeyPreview = True
    Set mobjBrushCheck = New clsBrushCardInput
    mobjBrushCheck.OnlyLegalCardNo = False
'    mobjCard.���ų��� = 18
'    mobjCard.������С���� = 8
'    mobjCard.������Ч�ַ� = Asc("=")
    'mobjCard.���Ž����� = Asc("=")
    'mobjCard.ˢ�������� = 13
    'mobjCard.�������Ĺ��� = "1-3"
'    mobjCard.������Ч�ַ� = "0" '��������(0-�����ַ�,1-����,2-��ĸ;3-���ֻ���ĸ;4-ָ���ַ�)|Ascii��1��Ascii��2....
    Call mobjBrushCheck.InitCompents(Me, Bill, mobjCard)
    
    If gSysPara.bytҩƷ������ʾ <> 2 Then
        '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        Bill.ColWidth(BillCol.��Ʒ��) = 0
    Else
        If Bill.ColWidth(BillCol.��Ʒ��) = 0 Then
             Bill.ColWidth(BillCol.��Ʒ��) = GetOrigColWidth(BillCol.��Ʒ��)
        End If
    End If
    
    Call SetMoneyList

    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gSysPara.bln����ƥ�䷽ʽ�л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gSysPara.bln����ƥ�䷽ʽ�л�
    sta.Panels("BarCode").Visible = mbytInState = 0
    
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���
        i = Val(gobjDatabase.GetPara("���뷽ʽ"))
        If i = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf i = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
        sta.Panels("BarCode").Bevel = IIf(Not mblnShowBarCode, sbrRaised, sbrInset)
        sta.Panels("BarCode").ToolTipText = IIf(Not mblnShowBarCode, "��ʾ���������", "�������������")
        Call ShowAndHideBarCodeInput
    End If

    '����
    If mbln���õǼ� Then
        lblTitle.Caption = gstrUnitName & "��Ѻ��õǼ�"
    Else
        If mint��¼���� = 1 Then
            lblTitle.Caption = gstrUnitName & "�����շѵ�"
        ElseIf mint��¼���� = 2 Then
            lblTitle.Caption = gstrUnitName & "���˼��ʵ�"
        End If
    End If
    txtӦ��.Text = gSysPara.Money_Decimal.strFormt_VB: txtʵ��.Text = gSysPara.Money_Decimal.strFormt_VB
    cmdSaveAndPay.Tag = "0"
    cmdSaveAndPay.Visible = False
    
    Select Case mbytInState
        Case 0 'ִ��
            Call SetShowCol
            cmdSelWholeSet.Visible = True
            cmdSaveWholeSet.Visible = zlCheckPrivs(mstrPrivs, "���ӳ�����Ŀ")
            '���ʵ�����ʱ��ֱ������˵ģ����������
            If gSysPara.bln�������������� = False And mint��¼���� = 1 Then
                If IsCan���֧��(mobjSquareCard, mlng����ID, mlng��ҳID) Then
                    cmdSaveAndPay.Visible = True
                    cmdSaveAndPay.Tag = "1"
                End If
            End If
        Case 1 '����
            cboNO.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
            cmdSelWholeSet.Visible = False
            cmdSaveWholeSet.Visible = zlCheckPrivs(mstrPrivs, "���ӳ�����Ŀ")
            cmdSaveWholeSet.Left = cmdSelWholeSet.Left
        Case 3 '����
            cboNO.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            cmdSelWholeSet.Visible = False
            cmdSaveWholeSet.Visible = zlCheckPrivs(mstrPrivs, "���ӳ�����Ŀ")
            cmdSaveWholeSet.Left = cmdSelWholeSet.Left
            '��ʱ��֧�ֲ���ɾ��
            If mint��¼���� <> 1 And False Then
                Call ShowDeleteCol(True)
                Bill.Active = True
            Else
                Bill.Active = False
            End If
    End Select
    
    If mbytInState <> 0 Then
        lblPreNO.Visible = False: txtPreNO.Visible = False
        lblӦ��.Top = lblӦ��.Top + txtPreNO.Height / 2
        txtӦ��.Top = txtӦ��.Top + txtPreNO.Height / 2
        lblʵ��.Top = lblʵ��.Top + txtPreNO.Height * 0.75
        txtʵ��.Top = txtʵ��.Top + txtPreNO.Height * 0.75
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    cboNO.Locked = Not bln
    txt����.Locked = Not bln
    cbo��������.Locked = Not bln
    cbo������.Locked = Not bln
    
    chk�Ӱ�.Enabled = bln
    cboBaby.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
End Sub

Private Function GetPatient(ByVal lng����ID As Long, ByVal lng��ҳId As Long) As Boolean
'���ܣ���ȡ������Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    mblnWarnCloseed = False
    mintWarn = -1: mstrWarn = ""
    Set mrsWarn = New ADODB.Recordset
    
    txt����.ForeColor = Me.ForeColor
    Set mrsInfo = New ADODB.Recordset
    
    If mint������Դ = 2 Then '��סԺ�����Ƿ����ǿ�Ƽ���Ȩ��
        If InStr(mstrPrivs, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(mstrPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
            strSql = ""
        ElseIf InStr(mstrPrivs, "��Ժδ��ǿ�Ƽ���") > 0 Then
            strSql = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
        ElseIf InStr(mstrPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
            strSql = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
        Else
            strSql = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
        End If
    End If
    
    '83877,Ƚ����,2015-4-14,���ﲡ�˸���"���˹Һż�¼"��ȡ"����",����GetItemInfo
    '�ֶ���ʹ�ò���ʱ���������ȷ����(��Nullֵ),����ΪadVarChar����
    If mint������Դ = 2 Or mstrRegNo = "" Then
        strSql = "Select" & _
            " A.����ID,Nvl(B.��ҳID,0) ��ҳID,To_Number(Nvl(B.��ǰ����ID,[3])) as ����ID," & _
            "       Nvl(B.��Ժ����ID,[3]) as ����ID,B.��Ժ����,B.��Ժ����," & _
            "       A.�����,B.סԺ��,B.��Ժ���� as ����,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ���� ,Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�," & _
            "       A.������," & IIf(mint������Դ = 2 And mint��¼���� = 2, "Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,", "A.������,") & _
            "       Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Y.���� as ������,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���," & _
            "       B.סԺҽʦ,zl_PatiDayCharge(A.����ID) as ���ն�,Nvl(B.����,A.����) as ����,Nvl(B.��������,0) as ��������,B.��˱�־,B.��ע as ���˱�ע" & _
            " From ������Ϣ A,������ҳ B,������� X,ҽ�Ƹ��ʽ Y" & _
            " Where A.����ID=B.����ID(+) And A.����ID=X.����ID(+) And X.����(+) = " & IIf(mint������Դ = 1, 1, 2) & strSql & _
            " And A.����ID=[1] And B.��ҳID(+)=[2] And A.ҽ�Ƹ��ʽ=Y.����(+)"
    Else
        strSql = "Select" & _
            " A.����ID,Nvl(B.��ҳID,0) ��ҳID,To_Number(Nvl(B.��ǰ����ID,[3])) as ����ID," & _
            "       Nvl(B.��Ժ����ID,[3]) as ����ID,B.��Ժ����,B.��Ժ����," & _
            "       A.�����,B.סԺ��,B.��Ժ���� as ����,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ���� ,Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�," & _
            "       A.������," & IIf(mint������Դ = 2 And mint��¼���� = 2, "Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,", "A.������,") & _
            "       Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Y.���� as ������,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���," & _
            "       B.סԺҽʦ,zl_PatiDayCharge(A.����ID) as ���ն�,Nvl(C.����,A.����) as ����,Nvl(B.��������,0) as ��������,B.��˱�־,B.��ע as ���˱�ע" & _
            " From ������Ϣ A,������ҳ B,���˹Һż�¼ C,������� X,ҽ�Ƹ��ʽ Y" & _
            " Where A.����ID=B.����ID(+) And A.����ID=X.����ID(+) And X.����(+) = " & IIf(mint������Դ = 1, 1, 2) & strSql & _
            " And A.����ID=[1] And B.��ҳID(+)=[2] And A.ҽ�Ƹ��ʽ=Y.����(+) And A.����ID=C.����ID And C.NO=[4]"
    End If

    On Error GoTo errH
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳId, mlng���˿���id, mstrRegNo)
    If Not mrsInfo.EOF Then
        mstrסԺҽ�� = Nvl(mrsInfo!סԺҽʦ)
        If Not IsNull(mrsInfo!����) Then
            txt����.ForeColor = vbRed
        End If
        
        '�������ﻮ������Ҫ���������
        If mint��¼���� = 2 Then
            If mint������Դ = 2 Then
                '49501:סԺ
                If zlIsAllowFeeChange(mrsInfo!����ID, Val(Nvl(mrsInfo!��ҳID)), Val(Nvl(mrsInfo!��˱�־))) = False Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt����.Text = "":
                    mlng����ID = 0
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
            End If
            
            'ˢ�²��˷���״��
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(mint������Դ = 1, 0, mlng��ҳID), mcurModiMoney)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!Ԥ�����
                cmdCancel.Tag = rsTmp!�������
                txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
            End If
            sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag), gSysPara.Money_Decimal.strFormt_VB)
            sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag), "0.00")
            
            'ˢ�±�����Ϣ
            strSql = "Select Nvl(��������,1) as ��������," & _
                " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
                " Where ���ò���=[2] And " & IIf(mint������Դ = 1, "Nvl(����ID,0)=0", "����ID=[1]")
            Set mrsWarn = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!����ID, 0)), CStr(Nvl(mrsInfo!���ò���)))
            
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------
            '���˺�:26952
            Dim cur��� As Currency, curItemMoney As Currency, curTotal As Double
            curItemMoney = 0
            curTotal = GetBillTotal(mobjBill)
            cur��� = Val(txtʵ��.Tag)
            If gSysPara.bln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(IIf(mint������Դ = 0, 0, 1), mrsInfo!����ID)

            If mbln���õǼ� = False Then    '30504
            
                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), cur���, Val(Nvl(mrsInfo!���ն�)) - mcurModiMoney, curTotal, _
                     Nvl(mrsInfo!������, 0), "", "", mstrWarn, mintWarn, , True)
                '����:0;û�б���,����
                '     1:������ʾ���û�ѡ�����
                '     2:������ʾ���û�ѡ���ж�
                '     3:������ʾ�����ж�
                '     4:ǿ�Ƽ��ʱ���,����
                '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
                If mbytWarn = 2 Or mbytWarn = 3 Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt����.Text = "":
                    mlng����ID = 0
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                If mrsWarn.EOF Then mrsWarn.Close '���ں���״̬�ж�
            End If
        End If
                            
        'סԺ���ʲŴ��������
        If mint������Դ = 2 Then
            '�������
            If Not IsNull(mrsInfo!����) Then
                chk����.Value = 0: chk����.Visible = True
            Else
                chk����.Value = 0: chk����.Visible = False
            End If
            
            '����ʱ��
            If Not IsNull(mrsInfo!��Ժ����) Then
                txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
            Else
                txtDate.Text = Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        
        Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
        
        '��ʾ������Ϣ
        txt����.Text = Nvl(mrsInfo!����)
        txt�Ա�.Text = Nvl(mrsInfo!�Ա�)
        txt����.Text = Nvl(mrsInfo!����)
        txt�ѱ�.Text = Nvl(mrsInfo!�ѱ�)
        txt���ʽ.Text = Nvl(mrsInfo!ҽ�Ƹ��ʽ)
        txt���ʽ.Tag = Nvl(mrsInfo!������, 0) '��Ҫ��дΪ��
        txt����.Text = Nvl(mrsInfo!����)
        
        '���˺� ����:26953 ����:2009-12-25 15:21:47
        txt���˱�ע = Nvl(mrsInfo!���˱�ע)
        If mint������Դ = 1 Then
            lblסԺ��.Caption = "�����"
            txtסԺ��.Text = Nvl(mrsInfo!�����)
        Else
            lblסԺ��.Caption = "סԺ��"
            txtסԺ��.Text = Nvl(mrsInfo!סԺ��)
        End If
        
        txt������.Text = Nvl(mrsInfo!������)
        txt������.Text = Format(Nvl(mrsInfo!������), "0.00")
        
        With mobjBill
            .����ID = Nvl(mrsInfo!����ID, 0)
            .��ҳID = Nvl(mrsInfo!��ҳID, 0)
            .����ID = Nvl(mrsInfo!����ID, 0)
            .����ID = Nvl(mrsInfo!����ID, 0)
            .���� = Nvl(mrsInfo!����)
            .��ʶ�� = IIf(mint������Դ = 1, Nvl(mrsInfo!�����), Nvl(mrsInfo!סԺ��))
            .���� = Nvl(mrsInfo!����)
            .�Ա� = Nvl(mrsInfo!�Ա�)
            .���� = Nvl(mrsInfo!����)
            .�ѱ� = Nvl(mrsInfo!�ѱ�)
        End With
        
        '�ڵ�һ�ν���ʱ��ȡ��������������Ŀ��Ϣ
        If Not Visible And mint������Դ = 2 And mint��¼���� = 2 And mbytInState = 0 Then Set mrsMedAudit = GetAuditRecord(mrsInfo!����ID, mrsInfo!��ҳID)
        
        GetPatient = True
    Else
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    Dim strMainRows As String
    Dim bln��������ۿ� As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln��������ۿ� = False
        If gSysPara.bln��������ۿ� And Not mbln���õǼ� Then                    '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
            If mobjBill.Details(i).�������� > 0 Then    '����
                bln��������ۿ� = Not mobjBill.Details(mobjBill.Details(i).��������).Detail.���ηѱ�
                If bln��������ۿ� And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).��������      '��������һ�е�ʱ��
            Else
                If CheckItemHaveSub(i) Then                          '����������
                     bln��������ۿ� = Not mobjBill.Details(i).Detail.���ηѱ�
                     If bln��������ۿ� Then strMainRows = strMainRows & "," & i  'һҳ�����ж��������,�ȼ�¼�����к�,���������������ۿ�
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln��������ۿ�)
    Next
    
    '������������,������bln��������ۿ۱���,��Ϊ�������������Ǵ������ʱ�Ѹı�
    If gSysPara.bln��������ۿ� And Not mbln���õǼ� Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc��������ʵ��(Split(strMainRows, ",")(i))
        Next
    End If
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln��������ۿ� As Boolean)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim rsTmp As New ADODB.Recordset
    Dim strInfo As String, i As Long
    Dim dblMoney As Double '�û�����ı�۽��
    Dim dbl�Ӱ�Ӽ��� As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim dblAllTime As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        strPriceGrade = mstrҩƷ�۸�ȼ�
    ElseIf mobjBill.Details(lngRow).�շ���� = "4" Then
        strPriceGrade = mstr���ļ۸�ȼ�
    Else
        strPriceGrade = mstr��ͨ�۸�ȼ�
    End If
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).�շ�ϸĿID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null "
    End If
    
    gstrSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL)) " & _
        "       And A.ID=[1]" & vbNewLine & _
                strWherePriceGrade
    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID, strPriceGrade)
    If Not rsTmp.EOF Then
        With mobjBill.Details(lngRow)
            If InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������) Then
                '����ҩƷʱ��(�����򲻷���),��Ȼ�м�¼(�������Ŀʱ���ж�)
                dblAllTime = .���� * .����
                If mblnҩ����λ And InStr(",5,6,7,", .�շ����) > 0 Then
                    dblAllTime = dblAllTime * .Detail.ҩ����װ '���ʱ�۰��ۼ��������м���
                End If
                If dblAllTime <> 0 Or Not .Detail.��� Then
                    If .Detail.���� <= 0 Then
                        gstrSQL = "Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual"
                    Else
                        gstrSQL = "Select Zl_Fun_Getprice([1],[2],[3],[4],[5]) As Price From Dual"
                    End If
                    Set rsPrice = gobjDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .�շ�ϸĿID, .ִ�в���ID, dblAllTime, 0, .Detail.����)
                    If rsPrice.EOF Then
                        '��ȡ�۸�ʧ��
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            MsgBox "�� " & lngRow & " ��ҩƷ""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                        Else
                            MsgBox "�� " & lngRow & " ����������""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                        End If
                    Else
                        strPrice = Nvl(rsPrice!Price) & "|||"
                        varPrice = Split(strPrice, "|")
                        dblMoney = Val(varPrice(0))
                        dblʣ������ = Val(varPrice(2))
                        
                        If dblʣ������ <> 0 And .Detail.��� Then
                            '����δ�ֽ����
                            If InStr(",5,6,7,", .�շ����) > 0 Then
                                MsgBox "�� " & lngRow & " ��ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            Else
                                MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            End If
                            dblMoney = 0
                        End If
                    End If
                Else
                    dblMoney = 0
                End If
            Else
                If .Detail.��� Then
                    If .InComes.Count = 0 Then  '��һ�μ�����ȡȱʡֵ
                        dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                    Else                        '��ȡ����Ա��ǰ����ı�۽��
                        dblMoney = .InComes(1).��׼����
                        '����û�����ı�۲������۷�Χ����ȡȱʡֵ
                        If Abs(dblMoney) > Abs(Val(Nvl(rsTmp!�ּ�))) Then
                            dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                        End If
                    End If
                End If
            End If
        End With
        
        '�����ԭ�м�¼
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        
        '��д���з��ü�¼
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .������ĿID = rsTmp!������ĿID
                .������Ŀ = rsTmp!����
                .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
                .ԭ�� = Val(Nvl(rsTmp!ԭ��))
                .�ּ� = Val(Nvl(rsTmp!�ּ�))
                
                If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                    If mblnҩ����λ Then
                        .��׼���� = Format(dblMoney * mobjBill.Details(lngRow).Detail.ҩ����װ, gSysPara.Price_Decimal.strFormt_VB)
                    Else
                        .��׼���� = Format(dblMoney, gSysPara.Price_Decimal.strFormt_VB)
                    End If
                Else
                    If mobjBill.Details(lngRow).Detail.��� Then
                        .��׼���� = Format(dblMoney, gSysPara.Price_Decimal.strFormt_VB)
                    Else
                        .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), gSysPara.Price_Decimal.strFormt_VB)
                    End If
                End If
                
                'Ӧ�ս��=���� * ���� * ����
                .Ӧ�ս�� = .��׼���� * mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
                
                '�������������ü���(����������Ŀ)
                If mobjBill.Details(lngRow).���ӱ�־ = 1 And mobjBill.Details(lngRow).�շ���� = "F" Then
                    .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
                End If
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                    dbl�Ӱ�Ӽ��� = Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100
                    .Ӧ�ս�� = .Ӧ�ս�� * (1 + dbl�Ӱ�Ӽ���)
                End If
                
                .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gSysPara.Money_Decimal.strFormt_VB))
                
                dblAllTime = mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
                If mblnҩ����λ And InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                    dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.ҩ����װ
                End If
                
                If mbln���õǼ� Or .Ӧ�ս�� = 0 Then
                    .ʵ�ս�� = 0
                Else
                    If mobjBill.Details(lngRow).Detail.���ηѱ� Or bln��������ۿ� Then
                        .ʵ�ս�� = .Ӧ�ս��
                    Else
                        .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��, _
                            mobjBill.Details(lngRow).�շ�ϸĿID, mobjBill.Details(lngRow).ִ�в���ID, _
                            dblAllTime, dbl�Ӱ�Ӽ���), gSysPara.Money_Decimal.strFormt_VB))
                    End If
                End If
                
                '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
                If Not IsNull(mrsInfo!����) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, mrsInfo!����, _
                        mobjBill.Details(lngRow).ժҪ & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Details(lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                        .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gSysPara.Money_Decimal.strFormt_VB)
                        mobjBill.Details(lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
                
                'ʵ�ս�����Key��,�Դ���ֱ�����(��Key�д��ԭʼʵ�ս��,����)
                mobjBill.Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
            End With
            rsTmp.MoveNext
        Next
    Else
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, curTotal As Currency
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
    
    curTotal = GetBillTotal(mobjBill)
    
    If IsNumeric(cmdOK.Tag) Then
        sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + curTotal, gSysPara.Money_Decimal.strFormt_VB)
        sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - curTotal, "0.00")
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '���������
    For i = 1 To Bill.Cols - 1
        '����ʱ�շ�������
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).�շ���� <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).�շ����)
    End If
    
    'ˢ�µ�����
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "���"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.�������
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "��Ʒ��"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.��Ʒ��
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���
            Case "��λ"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And mblnҩ����λ Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.ҩ����λ
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���㵥λ
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).����, 5)
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, gSysPara.Price_Decimal.strFormt_VB)
            Case "Ӧ�ս��"
                'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gSysPara.Money_Decimal.strFormt_VB)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gSysPara.Money_Decimal.strFormt_VB)
            Case "ִ�п���"
                If mobjBill.Details(lngRow).ִ�в���ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!���� & "-" & mrsUnit!����
                        Else
                            Bill.TextMatrix(lngRow, i) = Get��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        '�������ֻ(��)��ʾ����
                        Bill.TextMatrix(lngRow, i) = Get��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "��־"
                If mobjBill.Details(lngRow).�շ���� = "F" And mobjBill.Details(lngRow).���ӱ�־ = 1 Then
                    Bill.TextMatrix(lngRow, i) = "��"
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
'���ܣ�ˢ����ʾ������Ŀ������
    Dim blnExist As Boolean
    Dim curʵ�� As Currency, curӦ�� As Currency
    Dim i As Long, j As Long, k As Long
    
    mshMoney.Redraw = False
    
    '�������ܷ�Ŀ
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '�����Ƿ��Ѿ��������������Ŀ,������ϼ�,��������
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).������ĿID = mobjBill.Details(i).InComes(j).������ĿID Then
                    blnExist = True: Exit For
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + mobjBill.Details(i).InComes(j).ʵ�ս��
                mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + mobjBill.Details(i).InComes(j).Ӧ�ս��
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��
                End With
            End If
        Next
    Next
    
    'ˢ����ʾ
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).������Ŀ
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).ʵ�ս��, gSysPara.Money_Decimal.strFormt_VB)
        curʵ�� = curʵ�� + mcolMoneys(i).ʵ�ս��
        curӦ�� = curӦ�� + mcolMoneys(i).Ӧ�ս��
    Next
    
    txtӦ��.Text = Format(curӦ��, gSysPara.Money_Decimal.strFormt_VB)
    txtʵ��.Text = Format(curʵ��, gSysPara.Money_Decimal.strFormt_VB)
    
    mshMoney.TopRow = mshMoney.Rows - 1
    mshMoney.Redraw = True
End Sub

Private Function GetCurӦ��() As Currency
'���ܣ���ȡ���˵�ǰ���ݺϼƽ��(�շѲ����ۼӵ���ʱ��)
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCurӦ�� = GetCurӦ�� + mcolMoneys(i).Ӧ�ս��
    Next
End Function

Private Function GetInputDetail(ByVal lng��Ŀid As Long, ByVal lng���� As Long) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    
    If lngMediCareNO > 0 Then
        strSql = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,A.��������,A.����ժҪ,M.Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C." & mstrҩ����װ & ") as ҩ����װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C." & mstrҩ����λ & ") as ҩ����λ,D.��������,A.¼������,C.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,����֧����Ŀ M" & _
        " Where A.ID=C.ҩƷID(+) And A.ID=D.����ID(+) And B.����=A.���" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[2] " & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        "       And A.ID=[1] And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[3]"

    Else
        strSql = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.�������,A.��������,A.����ժҪ,0 as Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "        Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C." & mstrҩ����װ & ") as ҩ����װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C." & mstrҩ����λ & ") as ҩ����λ,D.��������,A.¼������,C.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1" & _
        " Where A.ID=C.ҩƷID(+) And A.ID=D.����ID(+) And B.����=A.���" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=[2] " & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng��Ŀid, IIf(gSysPara.bytҩƷ������ʾ = 1, 3, 1), lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .ҩ����װ = Nvl(rsTmp!ҩ����װ, 1)
        .ҩ����λ = Nvl(rsTmp!ҩ����λ)
        .���� = Nvl(rsTmp!����, 0) = 1
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .������� = Nvl(rsTmp!�������, 0)
        .���� = Nvl(rsTmp!��������)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .�������� = Nvl(rsTmp!��������, 0) = 1
        .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
        .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
        .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
        .���� = lng����
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = GetPay(lngRow)
    If Detail.��� <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '���ø���RowData
                Bill.RowData(lngRow) = Asc(Detail.���)
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIf(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                
                If InStr(",5,6,7,", Detail.���) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Details.Add tmpIncomes, Detail, .ID, CByte(lngRow), CInt(bytParent), .���, .���㵥λ, intPay, dblTime, 0, lngDoUnit, ""
        End With
    Else '��������Ѿ�����,���޸�
        
        If InStr(",5,6,7,", Detail.���) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .���� = intPay
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = dblTime
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSql As String
    
    strSql = "Select count(����ID) as NUM from �շѴ�����Ŀ where ����ID=[1]"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
    
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetSubDetails(ByVal lng��Ŀid As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
            
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    If lngMediCareNO > 0 Then
        strSql = _
        " Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D." & mstrҩ����װ & ") as ҩ����װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D." & mstrҩ����λ & ") as ҩ����λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,G.Ҫ������,D.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,����֧����Ŀ G" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+)" & _
        "       And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        "       And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[2] " & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        "       And C.����ID=[1] And A.ID=G.�շ�ϸĿID(+) And G.����(+)=[3] " & _
        " Order by ����"
    Else
        strSql = _
        " Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D." & mstrҩ����װ & ") as ҩ����װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D." & mstrҩ����λ & ") as ҩ����λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,0 as Ҫ������,D.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+)" & _
        "       And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        "       And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[2] " & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        "       And C.����ID=[1] " & _
        " Order by ����"
    End If

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng��Ŀid, IIf(gSysPara.bytҩƷ������ʾ = 1, 3, 1), lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
                .ID = rsTmp!ID
                .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
                .���� = rsTmp!����
                .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
                .��� = Nvl(rsTmp!���)
                .ҩ����װ = Nvl(rsTmp!ҩ����װ, 1)
                .ҩ����λ = Nvl(rsTmp!ҩ����λ)
                .���㵥λ = Nvl(rsTmp!���㵥λ)
                .���� = Nvl(rsTmp!����, 0) = 1
                .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
                .��� = rsTmp!���
                .������� = rsTmp!�������
                .���� = rsTmp!����
                .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
                .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
                .������� = Nvl(rsTmp!�������, 0)
                .���д��� = Nvl(rsTmp!���д���, 0)
                .�������� = Nvl(rsTmp!��������, 1)
                .���� = Nvl(rsTmp!��������)
                .�������� = Nvl(rsTmp!��������, 0) = 1
                .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
                .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
                .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
                GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                    .ҩ����װ, .ҩ����λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, .����ժҪ, .���д���, .��������, .��������, , , , , , .Ҫ������, , .��ҩ��̬, .��Ʒ��
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'���ܣ�ɾ��ָ���շ���Ŀ��
'˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).�������� > lngRow Then
            mobjBill.Details(i).�������� = mobjBill.Details(i).�������� - 1
        End If
        mobjBill.Details(i).��� = mobjBill.Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
    Set mobjBill = New ExpenseBill
    
    mcurModiMoney = 0
    mlngPreRow = 0
    cboNO.Text = ""
    chk�Ӱ�.Value = IIf(OverTime, 1, 0)
    txtDate.Text = Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                
    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo��������_Click
    With mobjBill
        .�����־ = mint������Դ
        .������ = NeedName(cbo������.Text)
        If cbo��������.ListIndex = -1 Then
            .��������ID = 0
        Else
            .��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        .������ = UserInfo.����
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
    End With
End Sub

Private Sub ClearMoney()
'���ܣ����������ʾ��
    Dim i As Long, j As Long
    mshMoney.Redraw = False
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.Cols - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mshMoney.Rows = 5
    mshMoney.Redraw = True
End Sub

Private Function SaveBill(Optional ByRef strNos As String) As Boolean
'����:���浱ǰ����ļ��ʵ���(����סԺ���ʡ����ۡ�������ߵ��޸�)
'���:mobjBill=���ݶ���
'����:�����Ƿ�ɹ�
    Dim int�к� As Integer, int��� As Integer, int�۸񸸺� As Integer
    Dim dbl���� As Double, dbl���� As Double
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim arrSQL As Variant, i As Long, j As Long
    Dim int���� As Integer, bln�ϴ� As Boolean
    Dim strSql As String, strStuffDept As String '��¼���Ϸ��ϲ���
    Dim strDeptIDs As String, str���ܺ� As String
    Dim cllProExeute As New Collection, varTemp As Variant
    Dim rsTmp As ADODB.Recordset
    Dim lngҽ��С��ID As Long, lng���˿���ID As Long, lng����ID As Long
    Dim blnTrans As Boolean
    
    strNos = ""
    If mstrOriginalNO = "" Then
        If mint��¼���� = 1 And mint������Դ <> 2 Then
            mobjBill.NO = gobjDatabase.GetNextNo(13)
        Else
            mobjBill.NO = gobjDatabase.GetNextNo(14)
        End If
    Else
        mobjBill.NO = mstrOriginalNO
    End If
    
    If mlngҽ��ID <> 0 And mint������Դ = 2 Then
        Call GetPatiInforFromAdvice(mlngҽ��ID, mlng����ID, mlng��ҳID, lng���˿���ID, lng����ID, lngҽ��С��ID)
    End If
    If lng���˿���ID = 0 Then
        lngҽ��С��ID = zlGetҽ��С��ID
        lng����ID = mobjBill.����ID
        lng���˿���ID = mobjBill.����ID
    End If
    
    int��� = 0
    arrSQL = Array()
    Set cllProExeute = New Collection
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.���� <> 0 Then
            For Each mobjBillIncome In mobjBillDetail.InComes
                int��� = int��� + 1 '��ǰ��¼���
                '��������
                With mobjBill
                    If mint������Դ = 2 Then
                        gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ҳID) & "," & _
                            IIf(.��ʶ�� = "", "NULL", "'" & .��ʶ�� & "'") & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "'," & _
                            ZVal(lng����ID) & "," & ZVal(lng���˿���ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                    Else
                        If mint��¼���� = 2 Then
                            gstrSQL = "zl_������ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & _
                                IIf(.��ʶ�� = "", "NULL", "'" & .��ʶ�� & "'") & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "'," & _
                                "'" & .�ѱ� & "'," & .�Ӱ��־ & "," & .Ӥ���� & "," & _
                                ZVal(.����ID) & "," & .��������ID & ",'" & .������ & "',"
                        Else
                            gstrSQL = "zl_���ﻮ�ۼ�¼_Insert('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ҳID) & "," & _
                                IIf(.��ʶ�� = "", "NULL", "'" & .��ʶ�� & "'") & ",'" & IIf(Val(txt���ʽ.Tag) = 0, "", txt���ʽ.Tag) & "','" & .���� & "'," & _
                                "'" & .�Ա� & "','" & .���� & "','" & .�ѱ� & "'," & .�Ӱ��־ & "," & _
                                  ZVal(.����ID) & "," & .��������ID & ",'" & .������ & "',"
                        End If
                    End If
                End With
                
                '�շ�ϸĿ����
                With mobjBillDetail
                    '�����������
                    If .��� <> int�к� Then
                        int�к� = .���
                        int�۸񸸺� = int���
                        
                        '���´����������
                        If mobjBill.Details(.���).�������� = 0 Then
                            For i = .��� + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).�������� = .��� Then
                                    mobjBill.Details(i).�������� = int���
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                    
                    If mint������Դ = 2 Then
                        gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & ZVal(.���մ���ID) & ",'" & .���ձ��� & "',"
                    ElseIf mint��¼���� = 1 Then
                        gstrSQL = gstrSQL & "NULL,"
                    End If
                    
                    dbl���� = .����
                    If InStr(",5,6,7,", .�շ����) > 0 And mblnҩ����λ Then
                        dbl���� = Format(.���� * .Detail.ҩ����װ, "0.00000")
                    End If
                    gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & "," & ZVal(.ִ�в���ID) & ","
                End With
                
                '������Ŀ����
                With mobjBillIncome
                    dbl���� = .��׼����
                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And mblnҩ����λ Then
                        dbl���� = Format(.��׼���� / mobjBillDetail.Detail.ҩ����װ, gSysPara.Price_Decimal.strFormt_VB)
                    End If
                    gstrSQL = gstrSQL & IIf(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & .������ĿID & "," & _
                        "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                    If mint������Դ = 2 Then
                        gstrSQL = gstrSQL & IIf(.ͳ���� = 0, "NULL", .ͳ����) & ","
                    End If
                End With
                                                
                '��������
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                
                '�Ƿ�ֻ���ɻ��۵�
    '                If mbln���õǼ� Then
    '                    int���� = 0 '��ķ��õǼǲ������ɻ��۵�
    '                Else
    '                    '����Ӧ���������,Ӧ��ִ��ҽ����������ж�
    '                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 Then
    '                        int���� = IIF(InStr(gstr���ͻ��۵�, "5") > 0, 1, 0)
    '                    Else
    '                        int���� = IIF(InStr(gstr���ͻ��۵�, mobjBillDetail.�շ����) > 0, 1, 0)
    '                    End If
    '                End If
                If int���� = 0 Then bln�ϴ� = True 'ֻҪ���ڲ��ǻ��۵���Ҫ�ϴ�
                
                '�ռ����Ϸ��ϲ���,�Ա��Զ�����,���ﲡ�˽�����ʱ(����Ϊ����ʱ����),סԺ����ֻ�м���
                'mint������Դ :1-���ﲡ��,2-סԺ����
                'mint��¼���� :1-�շ�(����),2-����(��/ס)
                With mobjBillDetail
                    Select Case .�շ����
                    Case "4"    '����
                        If (mint������Դ = 1 And mint��¼���� = 2 And gSysPara.bln�����Զ����� Or mint������Դ = 2 And gSysPara.intסԺ�Զ����� <> 0) And int���� = 0 Then
                            If .ִ�в���ID <> 0 And .Detail.�������� Then
                                If InStr("," & strStuffDept, "," & .ִ�в���ID & ",") = 0 Then
                                    strStuffDept = strStuffDept & "," & .ִ�в���ID
                                End If
                            End If
                        End If
                    Case "5", "6", "7"  'ҩƷ
                            If gSysPara.bln�շѺ��Զ���ҩ And mint������Դ = 1 And int���� = 0 Then
                                   If .ִ�в���ID <> 0 And Not gSysPara.bln���뷢ҩ Then
                                       If InStr(strDeptIDs & ",", "," & .ִ�в���ID & ",") = 0 Then
                                           strDeptIDs = strDeptIDs & "," & .ִ�в���ID
                                       End If
                                   End If
                               End If
                    End Select
                End With
                
                If mint������Դ = 2 Then
                    '�����:117445,����,2017/12/4,�����ҩƷ�����ĵ�����Ϊ0ʱû�������ݴ���
                    gstrSQL = gstrSQL & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                        "0," & IIf(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                        "NULL,'" & mobjBillDetail.ժҪ & "'," & chk����.Value & "," & ZVal(mlngҽ��ID) & "," & _
                        "Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.���� & "'," & _
                        IIf(mobjBill.��������ID = mlng��������ID, "1", "0") & "," & mlng��������ID & ",NULL" & _
                        IIf(lngҽ��С��ID = 0, ",-1", "," & lngҽ��С��ID) & ",0," & IIf(mobjBillDetail.Detail.���� = -1 Or mobjBillDetail.Detail.���� = 0, "Null", mobjBillDetail.Detail.����) & ")"
                Else
                    If mint��¼���� = 2 Then
                        gstrSQL = gstrSQL & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                            "NULL,'" & mobjBillDetail.ժҪ & "'," & ZVal(mlngҽ��ID) & ","
                        '  Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ����_In       ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  �����־_In   ������ü�¼.�����־%Type := 1,
                        gstrSQL = gstrSQL & "" & IIf(mbytӦ�ó��� = 1, 4, 1) & ","
                        '  ��ҩ��̬_In   ������ü�¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ��������_In   Number := 0,
                        gstrSQL = gstrSQL & "" & "0" & ","
                        '  ����_In       ҩƷ�շ���¼.����%Type := Null
                        gstrSQL = gstrSQL & "" & IIf(mobjBillDetail.Detail.���� = -1 Or mobjBillDetail.Detail.���� = 0, "Null", mobjBillDetail.Detail.����) & ","
                        '  ��ҳid_In     ������ü�¼.��ҳid%Type := Null,
                        gstrSQL = gstrSQL & "" & ZVal(mobjBill.��ҳID) & ","
                        '  ���˲���id_In ������ü�¼.���˲���id%Type := Null
                        gstrSQL = gstrSQL & "" & ZVal(lng����ID) & ")"
                    Else
                        gstrSQL = gstrSQL & "'" & UserInfo.���� & "'," & _
                            "'" & mobjBillDetail.ժҪ & "'," & ZVal(mlngҽ��ID) & ","
                        '  Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ����_In       ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ������Դ_In   Number := 1,
                        gstrSQL = gstrSQL & "" & IIf(mbytӦ�ó��� = 1, 4, 1) & ","
                        '  ���ձ���_In   ������ü�¼.���ձ���%Type := Null,
                        gstrSQL = gstrSQL & "'" & mobjBillDetail.���ձ��� & "',"
                        '  ��������_In   ������ü�¼.��������%Type := Null,
                        gstrSQL = gstrSQL & "'" & mobjBillDetail.Detail.���� & "',"
                        '  ������Ŀ��_In ������ü�¼.������Ŀ��%Type := Null,
                        gstrSQL = gstrSQL & "" & IIf(mobjBillDetail.������Ŀ��, 1, 0) & ","
                        '  ���մ���id_In ������ü�¼.���մ���id%Type := Null,
                        gstrSQL = gstrSQL & "" & ZVal(mobjBillDetail.���մ���ID) & ","
                        '  ��ҩ��̬_In   ������ü�¼.����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ��������_In   Number := 0,
                        gstrSQL = gstrSQL & "" & "0" & ","
                        '  ����_In       ҩƷ�շ���¼.����%Type := Null
                        gstrSQL = gstrSQL & "" & IIf(mobjBillDetail.Detail.���� = -1 Or mobjBillDetail.Detail.���� = 0, "Null", mobjBillDetail.Detail.����) & ","
                        '  ִ����_In     ������ü�¼.ִ����%Type := Null,
                        gstrSQL = gstrSQL & "" & "NULL" & ","
                        '  ���˲���id_In ������ü�¼.���˲���id%Type := Null
                        gstrSQL = gstrSQL & "" & ZVal(lng����ID) & ")"
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
            Next
        End If
    Next
    
    '-----------------------------------------------------------------------------------------------------------------
    If mbytӦ�ó��� = 0 Or mlngҽ��ID <> 0 Then
        'ҽ������ʱ,����ҽ��ID
        If mstrOriginalNO = "" Then
            '����ҽ��Ժ�ӷ���
            gstrSQL = "ZL_����ҽ������_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mint��¼���� & ",'" & mobjBill.NO & "')"
        Else
            '��������
            gstrSQL = "ZL_����ҽ������_�Ʒ�(" & mlngҽ��ID & "," & mlng���ͺ� & ")"
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
    End If
    
    '�޸�ǰ�˳�ԭ����
    If mstrInNO <> "" Then
        '���ж��Ƿ�ҽ�����˼ǵ���,�����Ϸ��Լ��(�����޸�ʱ������һ������ж�)
        If mint������Դ = 2 Then
            'ȥ����ҽ������ƥ����
            intInsure = BillExistInsure(mstrInNO)
        End If
        If mint������Դ = 2 Then
            gstrSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            If mint��¼���� = 2 Then
                gstrSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Else
                gstrSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & mstrInNO & "')"
            End If
        End If
        If gstrSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
        End If
    End If
    
    If UBound(arrSQL) >= 0 Then
        '��SQL���а��շ�ϸĿID����
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        'ִ��SQL���
        strTmp = ""
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        
            For i = 0 To UBound(arrSQL)
                Call gobjDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '-----------------------------------------------------------------------
            'ִ���Զ�����
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                varTemp = Split(strStuffDept, ",")
                For i = 0 To UBound(varTemp)
                    '69902:������,2014-02-09,ֻ��ͬ��������һ�µ�ִ�п�����Ŀ�����Զ�����
                    If Val(varTemp(i)) = Val(cbo��������.ItemData(cbo��������.ListIndex)) Then
                        strSql = "zl_�����շ���¼_��������(" & Val(varTemp(i)) & ",25,'" & mobjBill.NO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                        zlAddArray cllProExeute, strSql
                    End If
                Next
            End If
             
            ''            '-----------------------------------------------------------------------
            ''            '�շѺ��Զ���ҩ,���ʲ��Զ���ҩ,�շ��Ҳ��Ǳ���Ϊ���۵�,�����������
            ''            '--���˺�:�����ݲ�����
            ''            If strDeptIDs <> "" Then
            ''                strDeptIDs = Mid(strDeptIDs, 2)
            ''                varTemp = Split(strDeptIDs, ",")
            ''                For i = 0 To UBound(varTemp)
            ''                    strSQL = "ZL_ҩƷ�շ���¼_������ҩ(" & Val(varTemp(i)) & ",8,'" & strBillNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & mobjBill.Pages(P).������ & "')"
            ''                    zlAddArray cllProExeute, strSQL
            ''                Next
            ''            End If
            ''
            '׼���Զ���ҩ(����ͨ����),�����������в��ܶ�������
            If mblnSendMateria Then
                Set rsTmp = Get����ҩ�嵥(mobjBill.NO, Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"))
                If rsTmp.RecordCount > 0 Then
                    str���ܺ� = gobjDatabase.GetNextNo(20)
                    For i = 0 To rsTmp.RecordCount - 1
                        strSql = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "',to_date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null,Null,Null," & str���ܺ� & ")"
                        zlAddArray cllProExeute, strSql
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            'ִ�з�ҩ�ͷ���
            zlExecuteProcedureArrAy cllProExeute, Me.Caption, False, False
            '-----------------------------------------------------------------------
            
            
            'ҽ���ӿ�
            '1.ҽ�����������ϴ�
            If mint������Դ = 2 And mstrInNO <> "" And intInsure <> 0 Then
                If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.����ʵʱ�ϴ�
            If mint������Դ = 2 And bln�ϴ� And Not IsNull(mrsInfo!����) Then
                'ҽ�����������ϸ
                If gclsInsure.GetCapability(support�����ϴ�, mlng����ID, mrsInfo!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, mrsInfo!����) Then
                    strTmp = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!����) Then
                        gcnOracle.RollbackTrans
                        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        
        If Not mobjSaveData Is Nothing Then
            If mobjSaveData.SaveData(mobjBill.NO) = False Then
                SaveBill = False
                gcnOracle.RollbackTrans
                blnTrans = False
                Exit Function
            End If
        End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ���ӿ�
        '1.ҽ�����������ϴ�
        If mint������Դ = 2 And mstrInNO <> "" And intInsure > 0 Then
            If gclsInsure.GetCapability(support���������ϴ�, mlng����ID, intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """��ҽ������ʧ��,�õ��ݵķ�����ɾ����", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.����ʵʱ�ϴ�
        If mint������Դ = 2 And bln�ϴ� And Not IsNull(mrsInfo!����) Then
            'ҽ�����������ϸ
            If MCPAR.�����ϴ� And MCPAR.������ɺ��ϴ� Then
                strTmp = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!����) Then
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mobjBill.NO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '���뵥����ʷ��¼(�������͵���)
        cboNO.AddItem mobjBill.NO, 0
        mstrSuccesSaveNos = mstrSuccesSaveNos & "," & mobjBill.NO
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i 'ֻ��ʾ10��
        Next
        
        'ҽ���ӿ�
        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
    End If
    '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
    strNos = mobjBill.NO
    
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Integer
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������strNO=���ݺ�
'      blnDelete=�Ƿ��ȡҪ�˷ѵĵ���
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String, intSign As Integer
    Dim curTotal As Currency, curӦ��Total As Currency
    Dim strSql As String, i As Long
    Dim intInsure As Integer, blnDo As Boolean
    Dim blnNOMoved As Boolean
    
    If mbytInState = 1 Then
        If mint��¼���� = 1 Or (mint��¼���� = 2 And mint������Դ = 1) Then
            blnNOMoved = gobjDatabase.NOMoved("������ü�¼", strNO, "��¼����=", mint��¼����)
        Else
            blnNOMoved = gobjDatabase.NOMoved("סԺ���ü�¼", strNO, "��¼����=", mint��¼����)
        End If
    End If
    
    On Error GoTo errH
    
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    
    If mstrFeeTab = "סԺ���ü�¼" Then
        strSql = _
        " Select A.����ID,Nvl(A.��ҳID,0) ��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.����,A.��ʶ��," & _
        " A.���˲���ID,A.��������ID,A.�Ӱ��־,A.Ӥ����,A.������,A.������,A.����Ա����," & _
        " A.��������ID," & IIf(zlIsShowDeptCode, "C.����||'-'||", "") & "C.���� as ��������,A.����ʱ��," & _
        " B.ҽ�Ƹ��ʽ,B.������,B.������,A.�Ƿ���,B1.��ע as ���˱�ע" & _
        " From סԺ���ü�¼ A,������Ϣ B,���ű� C,������ҳ B1" & _
        " Where Rownum=1  And A.����id=B1.����id(+) and A.��ҳid=B1.��ҳID(+) And NO=[1] And A.��¼����=[2]" & _
        " And A.����ID=B.����ID And Instr([3],A.��¼״̬)>0" & _
        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[4]", "") & _
        " And A.��������ID=C.ID"
    Else
        strSql = _
        " Select A.����ID,0 as ��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.���ʽ as ����,A.��ʶ��," & _
        " 0 as ���˲���ID,A.��������ID,A.�Ӱ��־,A.Ӥ����,A.������,A.������,A.����Ա����," & _
        " A.��������ID," & IIf(zlIsShowDeptCode, "C.����||'-'||", "") & "C.���� as ��������,A.����ʱ��," & _
        " B.ҽ�Ƹ��ʽ,B.������,B.������,A.�Ƿ���,Null as ���˱�ע" & _
        " From ������ü�¼ A,������Ϣ B,���ű� C" & _
        " Where Rownum=1  And NO=[1] And A.��¼����=[2]" & _
        " And A.����ID=B.����ID And Instr([3],A.��¼״̬)>0" & _
        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[4]", "") & _
        " And A.��������ID=C.ID"
    End If
    If blnNOMoved Then
        strSql = Replace(strSql, mstrFeeTab, "H" & mstrFeeTab)
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, mint��¼����, _
        IIf(mblnDelete, "2", "0,1,3"), CDate(IIf(mstrTime = "", "1990-01-01", mstrTime)))
    If rsTmp.EOF Then
        MsgBox "û�з��ָõ��ݡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mlng����ID = 0 Then mlng����ID = Nvl(rsTmp!����ID, 0)
    
    cboNO.Text = strNO
    txt����.Text = Nvl(rsTmp!����)
    txt�Ա�.Text = Nvl(rsTmp!�Ա�)
    txt����.Text = Nvl(rsTmp!����)
    If Nvl(rsTmp!��ҳID, 0) <> 0 Then
        txt����.Text = Nvl(rsTmp!����)
    End If
    
    '���˺� ����:26953 ����:2009-12-25 15:23:48
    txt���˱�ע.Text = Nvl(rsTmp!���˱�ע)
    If mint������Դ = 1 Then
        lblסԺ��.Caption = "�����"
    Else
        lblסԺ��.Caption = "סԺ��"
    End If
    txtסԺ��.Text = Nvl(rsTmp!��ʶ��)
    
    txt�ѱ�.Text = Nvl(rsTmp!�ѱ�)
    txt������.Text = Nvl(rsTmp!������)
    txt������.Text = Format(Nvl(rsTmp!������), "0.00")
    txt���ʽ.Text = Nvl(rsTmp!ҽ�Ƹ��ʽ)
    
    cbo��������.AddItem Nvl(rsTmp!��������)
    cbo��������.ItemData(cbo��������.NewIndex) = Nvl(rsTmp!��������ID, 0)
    cbo��������.ListIndex = cbo��������.NewIndex
    
    If Nvl(rsTmp!�Ƿ���, 0) = 1 Then
        chk����.Value = 1: chk����.Visible = True
    End If
    
    chk�Ӱ�.Value = Nvl(rsTmp!�Ӱ��־, 0)
    Call LoadPatientBaby(cboBaby, rsTmp!����ID, rsTmp!��ҳID)
    Call gobjControl.CboLocate(cboBaby, Nvl(rsTmp!Ӥ����, 0), True)
    
    '������
    Call GetCboIndex(cbo������, Nvl(rsTmp!������))
    If cbo������.ListIndex = -1 And Not IsNull(rsTmp!������) Then
        cbo������.AddItem rsTmp!������
        cbo������.ListIndex = cbo������.NewIndex
    End If
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    If mint��¼���� = 2 Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!����ID, IIf(mint������Դ = 1, 0, mlng��ҳID))
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "Ԥ��:" & Format(rsPatiMoney!Ԥ�����, "0.00") & _
                "/����:" & Format(rsPatiMoney!�������, gSysPara.Money_Decimal.strFormt_VB) & _
                "/ʣ��:" & Format(rsPatiMoney!Ԥ����� - rsPatiMoney!�������, "0.00")
        End If
    End If
    
    '------------------------------------------------------------------------------------
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        
        '��ȡ������ԭʼ��¼�ķ���ID
        strSQL1 = _
            " Select A.ID,A.���,A.�շ�ϸĿID," & _
            " Nvl(A.����,1)*A.����" & IIf(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & " as ԭʼ����" & _
            " From " & mstrFeeTab & " A,ҩƷ��� B" & _
            " Where A.NO=[1] And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
            " And A.�շ�ϸĿID=B.ҩƷID(+) And A.��¼����=[2]"
        
        '��ȡҩƷ�շ���¼�е�׼����
        strSQL2 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.����� is NULL" & _
            " And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
        
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
        strSql = "Select Nvl(�۸񸸺�,���) From " & mstrFeeTab & _
            " Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1]" & _
            " And Nvl(ִ��״̬,0)<>1" & IIf(mlngҽ��ID <> 0, " And ҽ�����+0=[8]", "")
        
        '����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
        If mint��¼���� = 2 Then
            If mint������Դ = 2 Then intInsure = BillExistInsure(strNO)
            If intInsure <> 0 Then
                blnDo = Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, rsTmp!����ID, intInsure)
            Else
                blnDo = gSysPara.bytBillOpt = 2
            End If
            If blnDo Then
                strSql = strSql & " And Nvl(�۸񸸺�,���) IN" & _
                    " (" & _
                    " Select Nvl(�۸񸸺�,���) as ���" & _
                    " From " & mstrFeeTab & _
                    " Where NO=[1] And ��¼���� IN(2,12)" & _
                    " Group by Nvl(�۸񸸺�,���)" & _
                    " Having Sum(Nvl(���ʽ��,0))=0" & _
                    " )"
            End If
        End If
        
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSql = _
            " Select A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(mblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & mstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(A.����" & IIf(mblnҩ����λ, "/Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(mblnҩ����λ, "*Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From " & mstrFeeTab & " A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=[2]" & _
            " And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSql & ")" & _
            " Group by A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),C.����,C.����,A.�շ�ϸĿID,B.����," & _
            " B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X." & mstrҩ����λ & ",X." & mstrҩ����װ
            
        '��������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSql = _
            " Select A.���,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������,A.���㵥λ," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
            " Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
            " A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSql & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.���=B.��� And B.ID=C.����ID(+)" & _
            " Group by A.���,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������," & _
            " A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�в���,A.���ӱ�־" & _
            " Having Sum(A.����*A.����)<>0"
            
        strSql = _
            " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���," & _
            "       A.��������,A.���㵥λ,A.׼�˸��� as ����,A.׼������ as ����,A.����," & _
            "       A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
            "       A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
            "       A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSql & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[6]" & _
            "       And  A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " Order by A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        
        strSql = _
            "Select A.�շ�ϸĿID,A.�շ����,A.ִ�в���ID,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " A.���㵥λ,A.����,A.����,A.��׼����,A.Ӧ�ս��,A.ʵ�ս��,A.���ӱ�־,A.��������" & _
            " From " & mstrFeeTab & " A Where A.��¼����=[2]" & _
            " And Instr([4],A.��¼״̬)>0 And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[5]", "")
        If blnNOMoved Then
            strSql = strSql & " Union ALL " & Replace(strSql, mstrFeeTab, "H" & mstrFeeTab)
        End If
        
        strSql = _
            " Select A.���,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(mblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & mstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg([7]*A.����" & IIf(mblnҩ����λ, "/Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(mblnҩ����λ, "*Nvl(X." & mstrҩ����װ & ",1)", "") & ") as ����," & _
            " Sum([7]*A.Ӧ�ս��) as Ӧ�ս��,Sum([7]*A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From (" & strSql & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ����" & _
            " And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            " Group by A.���,C.����,C.����,A.�շ�ϸĿID,B.����,B.���," & _
            " Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X." & mstrҩ����λ
            
        strSql = _
            " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������," & _
            "       A.���㵥λ,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSql & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[6]" & _
            "       And  A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " Order by ���"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, mint��¼����, IIf(mint��¼���� = 2, ",9,25,", ",8,24,"), _
        IIf(mblnDelete, "2", "0,1,3"), CDate(IIf(mstrTime = "", "1990-01-01", mstrTime)), IIf(gSysPara.bytҩƷ������ʾ = 1, 3, 1), intSign, mlngҽ��ID)
    If rsTmp.EOF Then
        If blnDelete Then
            MsgBox "�����е�ǰ�޿��Բ����ļ�¼�����ܵ����е���Ŀ�Ѿ�ȫ��ִ�С�", vbInformation, gstrSysName
        Else
            MsgBox "�����е�ǰ�޿��Բ����ļ�¼��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        Bill.RowData(i) = rsTmp!��� '���ڼ�������
        
        Bill.TextMatrix(i, BillCol.���) = rsTmp!���
        Bill.TextMatrix(i, BillCol.��Ŀ) = rsTmp!����
        Bill.TextMatrix(i, BillCol.��Ʒ��) = Nvl(rsTmp!��Ʒ��)
        Bill.TextMatrix(i, BillCol.���) = Nvl(rsTmp!���)
        Bill.TextMatrix(i, BillCol.��λ) = Nvl(rsTmp!���㵥λ)
        Bill.TextMatrix(i, BillCol.����) = Nvl(rsTmp!����)
        Bill.TextMatrix(i, BillCol.����) = FormatEx(rsTmp!����, 5)
        Bill.TextMatrix(i, BillCol.����) = Format(rsTmp!����, gSysPara.Price_Decimal.strFormt_VB)
        Bill.TextMatrix(i, BillCol.Ӧ�ս��) = Format(rsTmp!Ӧ�ս��, gSysPara.Money_Decimal.strFormt_VB)
        Bill.TextMatrix(i, BillCol.ʵ�ս��) = Format(rsTmp!ʵ�ս��, gSysPara.Money_Decimal.strFormt_VB)
        Bill.TextMatrix(i, BillCol.ִ�п���) = Nvl(rsTmp!ִ�в���)
        Bill.TextMatrix(i, BillCol.��־) = IIf(rsTmp!���ӱ�־ = 1, "��", "")
        Bill.TextMatrix(i, BillCol.����) = Nvl(rsTmp!��������)
        
        '�������ʱ�־
        If Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "��"
        End If
        
        rsTmp.MoveNext
    Next
    '����б༭����������ɫ
    Bill.SetColColor BillCol.���, &HE7CFBA
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.��־, &HE0E0E0
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))

        '��ȡҩƷ�շ���¼�е�׼����
        strSQL1 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(mblnҩ����λ, "/Nvl(B." & mstrҩ����װ & ",1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.����� is NULL" & _
            " And Instr([3],','||A.����||',')>0" & _
            " Group by A.����ID"
        
        '���ŷ��õ���(��ϸ��������Ŀ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSql = "Select Nvl(�۸񸸺�,���) From " & mstrFeeTab & _
            " Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1]" & _
            " And Nvl(ִ��״̬,0)<>1" & IIf(mlngҽ��ID <> 0, " And ҽ�����+0=[7]", "")
        If blnDo Then
            strSql = strSql & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From " & mstrFeeTab & _
                " Where NO=[1] And ��¼���� IN(2,12)" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
        
        strSql = _
            " Select Sum(A.ID) as ID,A.���,A.����,A.�շ����," & _
            " Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
            " Sum(A.ʵ�ս��) as ʣ��ʵ�� From (" & _
            " Select Decode(A.��¼״̬,2,0,A.ID) as ID,A.���,B.����,A.�շ����," & _
            " Nvl(A.����,1)*A.����" & IIf(mblnҩ����λ, "/Nvl(X." & mstrҩ����װ & ",1)", "") & " as ����," & _
            " A.Ӧ�ս��,A.ʵ�ս��" & _
            " From " & mstrFeeTab & " A,������Ŀ B,ҩƷ��� X" & _
            " Where A.��¼����=[2] And A.NO=[1]" & _
            " And A.������ĿID=B.ID And Nvl(A.�۸񸸺�,A.���) IN(" & strSql & ")" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+)) A" & _
            " Group by A.���,A.����,A.�շ����" & _
            " Having Sum(A.����)<>0"
                    
        '��������
        strSql = _
            " Select A.����,Sum(A.ʣ��Ӧ��*(A.׼������/A.ʣ������)) as Ӧ�ս��," & _
            " Sum(ʣ��ʵ��*(A.׼������/A.ʣ������)) as ʵ�ս�� From (" & _
            " Select A.����,A.ʣ������,A.ʣ��Ӧ��,A.ʣ��ʵ��," & _
            " Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,Nvl(B.׼������,A.ʣ������)) as ׼������" & _
            " From (" & strSql & ") A,(" & strSQL1 & ") B" & _
            " Where A.ID=B.����ID(+)" & _
            " ) A Group by A.����"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        
        strSql = "Select A.������ĿID,A.Ӧ�ս��,A.ʵ�ս�� From " & mstrFeeTab & " A" & _
            " Where Instr([4],A.��¼״̬)>0 And A.��¼����=[2] And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[5]", "")
        If blnNOMoved Then
            strSql = strSql & " Union ALL " & Replace(strSql, mstrFeeTab, "H" & mstrFeeTab)
        End If
        
        strSql = _
            " Select B.����,Sum([6]*A.Ӧ�ս��) as Ӧ�ս��,Sum([6]*A.ʵ�ս��) as ʵ�ս�� " & _
            " From (" & strSql & ") A,������Ŀ B Where A.������ĿID=B.ID Group By B.����"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, mint��¼����, IIf(mint��¼���� = 2, ",9,25,", ",8,24,"), _
        IIf(mblnDelete, "2", "0,1,3"), CDate(IIf(mstrTime = "", "1990-01-01", mstrTime)), intSign, mlngҽ��ID)
    If rsTmp.EOF Then Exit Function
    
    'ˢ����ʾ(�շ�Ҫ����)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    Call SetMoneyList
    
    For i = 1 To rsTmp.RecordCount
        mshMoney.TextMatrix(i, 0) = rsTmp!����
        mshMoney.TextMatrix(i, 1) = Format(rsTmp!ʵ�ս��, gSysPara.Money_Decimal.strFormt_VB)
        curTotal = curTotal + rsTmp!ʵ�ս��
        curӦ��Total = curӦ��Total + rsTmp!Ӧ�ս��
        rsTmp.MoveNext
    Next
    
    txtʵ��.Text = Format(curTotal, gSysPara.Money_Decimal.strFormt_VB)
    txtӦ��.Text = Format(curӦ��Total, gSysPara.Money_Decimal.strFormt_VB)
    
    ReadBill = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub SetShowCol()
'���ܣ������еĿ���(���ʱչ��)
    mrsClass.Filter = "����='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.����) = 0
    ElseIf Bill.ColWidth(BillCol.����) = 0 Then
        Bill.ColWidth(BillCol.����) = 520
    End If
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Function GetPay(lngRow As Long) As Integer
    Dim i As Long
    'ȡ������ҩ�ĸ���
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).����
            Exit For
        End If
    Next
End Function

Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
'���ܣ�ȡ���пɹ�ѡ���ҩ��
    Dim strSql As String, bytDay As Byte
    Dim strҩ�� As String, lng��������id As Long
    
    lng��������id = mrsInfo!����ID    '������������
    If lng��������id = 0 And cbo��������.ListIndex <> -1 Then lng��������id = cbo��������.ItemData(cbo��������.ListIndex)
    
    If str��� = "4" Then
        strSql = _
            "Select Distinct c.Id, c.����, c.����, c.����, b.��������, b.�������" & vbNewLine & _
            "From �շ�ִ�п��� A, ��������˵�� B, ���ű� C" & vbNewLine & _
            "Where a.ִ�п���id + 0 = b.����id And b.�������� = '���ϲ���' And b.������� IN([1],3) And b.����id = c.Id And" & vbNewLine & _
            "      (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null) And" & vbNewLine & _
            "      (a.������Դ Is Null Or a.������Դ = [1]) And" & vbNewLine & _
            "      (a.��������id Is Null Or a.��������id = [2] Or Exists (Select 1 From �������Ҷ�Ӧ Where ����id = [2] And a.��������id = ����id)) And a.�շ�ϸĿid = [3]" & vbNewLine & _
            "Order By b.�������, c.����"
            
    Else
        '��ҩƷ����ȷ��ҩ������
        Select Case str���
            Case "5"
                strҩ�� = "��ҩ��"
            Case "6"
                strҩ�� = "��ҩ��"
            Case "7"
                strҩ�� = "��ҩ��"
        End Select
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not Check�ϰల��(True) Then
            strSql = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
                " And B.������� IN([1],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[1])" & _
                " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[3]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(gobjDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSql = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
                " And B.������� IN([1],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[1])" & _
                " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[3]" & _
                " Order by B.�������,C.����"
        End If
    End If
    
    On Error GoTo errH
    'Set mrsWork = New ADODB.Recordset
    Set mrsWork = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mint������Դ, lng��������id, lngҩƷID, strҩ��, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub Load������(ByVal lng����id As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngOldID As Long
    
    cbo������.Clear
    
    '����ҽ����ʿ
    strSql = _
        "   Select Distinct A.ID,B.����ID,A.���,A.����, Upper(A.����) as ����," & _
        "       C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        "   From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        "   Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        "       And C.��Ա���� IN('ҽ��','��ʿ') And B.����ID=[1]  " & _
        "   Order by ����,��Ա���� Desc"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng����id)
    
    i = IIf(rsTmp.RecordCount = 0, 0, rsTmp.RecordCount - 1)
    ReDim marrDr(i)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If lngOldID <> rsTmp!ID Then
                cbo������.AddItem IIf(IsNull(rsTmp!����), "", rsTmp!���� & "-") & rsTmp!����
                cbo������.ItemData(cbo������.ListCount - 1) = rsTmp!����ID
                marrDr(cbo������.ListCount - 1) = rsTmp!ID & "|" & rsTmp!����ID & "|" & Nvl(rsTmp!���) & "|" & rsTmp!���� & "|" & Nvl(rsTmp!����) & "|" & rsTmp!ְ�� & "|" & Nvl(rsTmp!��Ա����)
                
                If rsTmp!���� = mstr����ҽ�� Then cbo������.ListIndex = cbo������.NewIndex
                If lng����id = mlng���˿���id Then
                    'ȱʡΪ���˿���ʱ,����Ƿ�ΪסԺҽ��
                    '����:36862
                    If rsTmp!���� = mstrסԺҽ�� Then cbo������.ListIndex = cbo������.NewIndex
                End If
                If rsTmp!ID = UserInfo.ID And cbo������.ListIndex = -1 Then cbo������.ListIndex = cbo������.NewIndex
                lngOldID = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        
        If cbo������.ListCount > 0 Then ReDim Preserve marrDr(cbo������.ListCount - 1)
        If cbo������.ListCount = 1 And cbo������.ListIndex = -1 Then cbo������.ListIndex = 0
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function CalcGridToTal(Optional blnӦ�� As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If blnӦ�� Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.Ӧ�ս��
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.ʵ�ս��
                End If
            Next
        Next
    Else
        For i = 1 To Bill.Cols - 1
            If blnӦ�� Then
                If Bill.TextMatrix(0, i) = "Ӧ�ս��" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "ʵ�ս��" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'���ܣ���ʾ\�������ʱ�־��
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "ɾ��" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = -1
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���) - 120
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ) - 100
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���) - 200
            
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����) - 50
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��) - 50
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "ɾ��" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���)
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ)
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���)
            
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����)
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��)
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub SetColNum(Optional intRow As Long = 1)
'���ܣ�������ʾ���е��к�
'������intRow=�Ӹ��п�ʼ
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.��) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True) As Integer
'���ܣ����ָ��ҩƷ�е�ְ���Ƿ��뵱ǰҽ����ְ����ƥ��
'������tmpDetail=�������Ŀ,����Ϊ������,blnCommon=�Ƿ��������ж�,����Ϊҽ���򹫷Ѳ��˵��ж�
'���أ���ƥ�����,0Ϊ��ȷ
'˵����ְ��1=����,2=����,3=�м�,4=����/ʦ��,5=Ա/ʿ,9=��Ƹ
    Dim i As Long, intְ��A As Integer, intְ��B As Integer
    Dim strTmp As String
    
    strTmp = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    
    If cbo������.ListIndex = -1 Then Exit Function
    If cbo������.ListIndex <= UBound(marrDr) Then
        If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 5 Then
            intְ��A = Val(Split(marrDr(cbo������.ListIndex), "|")(5))
        End If
    End If
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                If Not blnCommon Then
                    intְ��B = Val(Right(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    intְ��B = Val(Left(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.���) = 0 Then Exit Function
        If Not blnCommon Then
            intְ��B = Val(Right(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        Else
            intְ��B = Val(Left(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer, Optional blnCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ��ҩƷ�ڵ������Ƿ��Ѿ�����
    '���:objDetail=��Ŀ,intRow=Ҫ�жϵ���
    '����:blnCancel-�����Ҫǿ��ȡ��������true
    '����:true��ʾ��������ʾ��false-��ʾ�Ϸ�
    '����:���˺�
    '����:2017-11-22 17:23:06
    '˵��:ʱ�ۻ����ҩƷ��ͬһҩ����ֹ�ظ�����(�������ʾ,����ʱ��ֹ(blnCancel=trueʱ����))
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    blnCancel = False
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.���� Or mobjBill.Details(i).Detail.���) _
                    And (objDetail.���� Or objDetail.���) Then
                    If objDetail.��� = "4" Then
                        If mobjBill.Details(i).Detail.���� <> objDetail.���� And objDetail.���� >= 0 Then
                           'ɨ����ģ�����ˢ������ε�
                        Else
                            If mobjBill.Details(i).Detail.���� = objDetail.���� And objDetail.���� > 0 Then
                                Call MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�������ͬ������,��ֹ����!", vbInformation + vbOKOnly, gstrSysName)
                                blnCancel = True
                                PhysicExist = True
                                Exit Function
                            End If
                            If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                                vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵķ��ϲ��Ų�ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                PhysicExist = True
                            End If
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺��ҩƷΪ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function
Private Function Check��������(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSql As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnҽ�� As Boolean, bln���� As Boolean
    Check�������� = True
    
    '�޷����
    If txt���ʽ.Tag = "" Then Exit Function
    '45605
    'ֻ���ҽ�����˺͹��Ѳ���
    If zlIsCheckMedicinePayMode(txt���ʽ.Text, blnҽ��, bln����) = False Then Exit Function
    'ȷ����������
    bytType = IIf(blnҽ��, 1, 2) ' Val(txt���ʽ.Tag)
    '��ȡ�������
    If bytType = 1 Then
        strSql = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gSysPara.strҽ���������� & ") Order by ����"
    Else
        strSql = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gSysPara.str���ѷ������� & ") Order by ����"
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)

    If rsTmp.EOF Then Exit Function
    
    If intRow > 0 Then
        If mobjBill.Details(intRow).Detail.���� = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            Check�������� = False
        Else
            rsTmp.Filter = "����='" & mobjBill.Details(intRow).Detail.���� & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������Ϊ""" & _
                    mobjBill.Details(intRow).Detail.���� & """,����" & _
                    IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
                Check�������� = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).Detail.���� = "" Then
                If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check�������� = False: Exit For
                End If
            Else
                rsTmp.Filter = "����='" & mobjBill.Details(i).Detail.���� & "'"
                If rsTmp.EOF Then
                    If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                        mobjBill.Details(i).Detail.���� & """,����" & _
                        IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check�������� = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    
    If Not IsNull(mrsInfo!����) Then
        For i = 1 To mobjBill.Details.Count
            For j = 1 To mobjBill.Details(i).InComes.Count
                dblAllTime = mobjBill.Details(i).���� * mobjBill.Details(i).����
                If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                    If mblnҩ����λ Then dblAllTime = dblAllTime * mobjBill.Details(i).Detail.ҩ����װ
                End If
                strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, mrsInfo!����, _
                    mobjBill.Details(i).ժҪ & "||" & dblAllTime)
                    
                If strInfo <> "" Then
                    mobjBill.Details(i).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(i).���մ���ID = Val(Split(strInfo, ";")(1))
                    mobjBill.Details(i).InComes(j).ͳ���� = Val(Split(strInfo, ";")(2))
                    mobjBill.Details(i).���ձ��� = CStr(Split(strInfo, ";")(3))
                    
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).ժҪ = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.���� = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            Next
        Next
    End If
End Sub

Private Function HaveStopClass() As Integer
'���ܣ��жϵ�ǰ�������Ƿ��л�ʿ��ֹ���������
    Dim i As Long, str���� As String
    
    If cbo������.ListIndex <> -1 Then
        If cbo������.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo������.ListIndex), "|")) >= 6 Then
                str���� = Split(marrDr(cbo������.ListIndex), "|")(6)
            End If
        End If
    End If
    
    For i = 1 To mobjBill.Details.Count
        If str���� = "��ʿ" And InStr(",E,M,4,", mobjBill.Details(i).�շ����) = 0 Then
            HaveStopClass = i: Exit Function
        End If
    Next
End Function

Private Function Checkִ�п���() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).ִ�в���ID = 0 Or Bill.TextMatrix(i, BillCol.ִ�п���) = "" Then
            Checkִ�п��� = i: Exit Function
        End If
    Next
End Function

Public Sub InitLocPar()
'���ܣ���ʼ�����ñ�������
    mstrLike = IIf(Val(gobjDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mblnPay = Val(gobjDatabase.GetPara("��ҩ���븶��", glngSys, pҽ�����ѹ���)) <> 0
    mblnTime = Val(gobjDatabase.GetPara("�����������", glngSys, pҽ�����ѹ���)) <> 0
    mbln����ҩ�� = Val(gobjDatabase.GetPara("��ʾ����ҩ�����", glngSys, pҽ�����ѹ���)) = 1
    mbln����ҩ�� = Val(gobjDatabase.GetPara("��ʾ����ҩ����", glngSys, pҽ�����ѹ���)) = 1
    mstr�շ���� = gobjDatabase.GetPara("�շ����", glngSys, pҽ�����ѹ���, "")
    
    'ҩƷ��λ
    mblnҩ����λ = Val(gobjDatabase.GetPara("ҩƷ��λ", glngSys, pҽ�����ѹ���)) <> 0
    If mint������Դ = 1 Then
        mstrҩ����λ = "���ﵥλ": mstrҩ����װ = "�����װ"
    Else
        mstrҩ����λ = "סԺ��λ": mstrҩ����װ = "סԺ��װ"
    End If
    mbytSendMateria = Val(gobjDatabase.GetPara("���ʺ�ҩ", glngSys, pҽ�����ѹ���))
    mbytȱʡ���� = Val(gobjDatabase.GetPara("����ȱʡ����", glngSys, pҽ�����ѹ���))
    'ȱʡҩ��
    mlng��ҩ�� = Val(gobjDatabase.GetPara(IIf(mint������Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, pҽ�����ѹ���))
    mlng��ҩ�� = Val(gobjDatabase.GetPara(IIf(mint������Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, pҽ�����ѹ���))
    mlng��ҩ�� = Val(gobjDatabase.GetPara(IIf(mint������Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, pҽ�����ѹ���))
    mlng���ϲ��� = Val(gobjDatabase.GetPara(IIf(mint������Դ = 2, "סԺ", "����") & "ȱʡ���ϲ���", glngSys, pҽ�����ѹ���))
    mblnShowBarCode = Val(gobjDatabase.GetPara("�ϴ�ѡ���������", glngSys, pҽ�����ѹ���)) = 1
    
End Sub

Public Function zlCheck����ҽ��(ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ա���ҽ����һЩ���
    '���:intInsuer-����
    '����:
    '����:���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-07 10:25:04
    '����:27278
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    If intInsure = 0 Then zlCheck����ҽ�� = True: Exit Function
    
    Err = 0: On Error GoTo Errhand:
    '���˺�:???
    'mint������Դ:1-���ﲡ��,2-סԺ����
    'mint��¼����:1-�շ�(����),2-����(��/ס)
    'mbytInState :0-ִ��,1-����,2-����(��֧��),3-ɾ��
    
    'ֻ�л��۲�֧�ּ��
    If (mint������Դ = 2 Or mint��¼���� = 2) And mbytInState <> 0 Or MCPAR.ҽ��ȷ���������� = False Then
        zlCheck����ҽ�� = True: Exit Function
    End If
    
    'showmsgbox
    '������strCaption=��Ϣ�������
    '      strInfo=������ʾ����,����"^"��ʾ����,">"��ʾ������
    '      strCmds=��ť����,��"����(&R),!����(&A),?ȡ��(&C)"
    '              ����Ҫ��������ť,"!"��ʾȱʡ��ť,"?"��ʾȡ����ť
    '              ÿ����ť�������֧��4������
    '      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
    '���أ���ť����,��"��ť2"(������()��&),������رջ�ȡ���򷵻�""
    strTemp = gobjCommFun.ShowMsgBox("��������", "��ȷ����ǰҽ�����˱���Ҫ���͵�ҩƷ���������͡�", "!ҽ����(&A),ҽ����(&B),?ȡ��(&C)", Me)
    If strTemp = "" Then Exit Function
    '����ǲ������շѻ��۵�������ҽ�����ˣ���ҽ��������supportҽ��ȷ���������͡���Чʱ������ʱ��ʾ�õ����ǡ�ҽ���ڣ�ҽ���⡱�������ҽ���ڷ��ü�¼ժҪ�д��1��ҽ������2��
    strTemp = IIf(strTemp = "ҽ����", 1, 2)
    For Each mobjBillDetail In mobjBill.Details
        mobjBillDetail.ժҪ = strTemp
    Next
    zlCheck����ҽ�� = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function

Public Function zlCheck������۸����(ByVal lng�շ�ϸĿID As Long, bln���� As Boolean, _
    ByVal strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ������(����Ϊ���)
    '���:
    '����:
    '����:���������ĿΪ��,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-12 11:22:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
   '���˺� ����:27286 ���۵ļ۸�Ϊ��Ĳ����м����� ����:2010-01-07 15:13:45
    Dim strSql As String, rs�۸� As ADODB.Recordset, dbl�۸� As Double
    Dim strWherePriceGrade As String
    
    Err = 0: On Error GoTo Errhand:
    zlCheck������۸���� = False
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null "
    End If
    If bln���� Then
        strSql = _
        " Select  B.�ּ� " & _
        " From �շѼ�Ŀ B " & _
        " Where   ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        "       And B.�շ�ϸĿID=[1]" & vbNewLine & _
                strWherePriceGrade
        Set rs�۸� = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��ǰ�۸�", lng�շ�ϸĿID, strPriceGrade)
        If rs�۸�.EOF = False Then
            dbl�۸� = Val(Nvl(rs�۸�!�ּ�))
        Else
            dbl�۸� = 0
        End If
        If dbl�۸� = 0 Then zlCheck������۸���� = True: Exit Function
    End If
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function
Public Function zl��ȡ��ҩ��̬(Optional ByVal lngRow As Long = -1, Optional blnOnly�г�ҩ As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ƿ�¼�����в�ҩ��
    '���:blnOnly�г�ҩ-���ж��Ƿ����г�ҩ(���䷽ʱ�ж���Ч):ԭ�����л�ҩ���䷽���Ѿ�����,�Ͳ���Ҫ���
    '     lngRow-��ǰ��������
    '����:
    '����:¼�����в�ҩ��,�򷵻��������(1-���,0-��Ҫ��),���򷵻�-1 ��ʾ��û��¼�������Ŀ
    '����:���˺�
    '����:2010-02-02 11:44:17
    '����:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl��ȡ��ҩ��̬ = -1
    '���δָ��ҳ,���õ�ǰҳ
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly�г�ҩ, ",6,", ",6,7,")
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).�շ���� & ",") > 0 And .Item(i).�շ�ϸĿID <> 0 And i <> lngRow Then
                zl��ȡ��ҩ��̬ = .Item(i).Detail.��ҩ��̬
                Exit Function
            End If
        Next
    End With
End Function
Public Function Get����ҩ�嵥(strNO As String, strTime As String) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݷ��õ��ݺ�,�Ǽ�ʱ��,��ȡ����ҩƷ�嵥
    '��Σ�strNO-���ݺ�
    '          strTime-�Ǽ�ʱ��
    '���Σ�
    '���أ�����ҩ�嵥
    '���ƣ����˺�
    '���ڣ�2010-03-19 18:59:27
    '˵������ͨ��ҩʱΪ���˿��ң����ҽ����Ϊ�������ҡ�
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = _
        " Select A.ID,A.�ⷿID,A.�Է�����ID" & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B" & _
        " Where A.NO=[1] And A.����=[2] And Mod(A.��¼״̬,3)=1 And A.����� is NULL" & _
        " And A.NO=B.NO And A.����ID=B.ID And B.��¼״̬<>0 And B.�Ǽ�ʱ��+0=[3]" & _
        " Order by A.ҩƷID"
    If strTime <> "" Then
        Set Get����ҩ�嵥 = gobjDatabase.OpenSQLRecord(strSql, "mdlInExse", strNO, 9, CDate(strTime))
    Else
        Set Get����ҩ�嵥 = gobjDatabase.OpenSQLRecord(strSql, "mdlInExse", strNO, 9)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Private Sub SetDetailtStock(ByVal lngִ�п���ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ϸ�Ŀ������
    '���ƣ����˺�
    '���ڣ�2010-07-12 14:27:51
    '˵����
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim strҩ��IDs As String, dblStock As Double
    
    '��ȡ���
    '�������ҩƷ������
    If InStr(1, "5,6,7,4", objDetail.���) = 0 Then Exit Sub
    If objDetail.��� = "4" And objDetail.�������� = False Then Exit Sub
    If objDetail.��� = "4" Then
        '����
        dblStock = GetStock(objDetail.ID, lngִ�п���ID, , , , objDetail.����)
        objDetail.��� = dblStock
        Exit Sub
    End If
    
    dblStock = GetStock(objDetail.ID, lngִ�п���ID)
    If mblnҩ����λ Then
        dblStock = dblStock / objDetail.ҩ����װ
    End If
    objDetail.��� = dblStock  '��¼��ǰ��ҩƷ���
End Sub

Private Sub cmdSelWholeSet_Click()
    'ѡ������Ŀ
    Dim rsSel As ADODB.Recordset, lng����ID As Long, lng��������ID As Long
    Dim tmpBill As New ExpenseBill, bytӤ���� As Byte, curDate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim lng���˿���ID As Long, str�ѱ� As String, intInsure As Integer
    intInsure = 0
    If mobjBill Is Nothing Then
        If mrsInfo Is Nothing Then
            MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        lng����ID = Val(Nvl(mrsInfo!����ID))
        intInsure = Val(Nvl(mrsInfo!����))
        If cbo��������.ListIndex < 0 Then
            lng��������ID = 0
        Else
            lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        If cboBaby.ListIndex < 0 Then
            bytӤ���� = 0
        Else
            bytӤ���� = cboBaby.ItemData(cboBaby.ListIndex)
        End If
        lng���˿���ID = mlng���˿���id: str�ѱ� = Nvl(mrsInfo!�ѱ�)
    Else
        lng����ID = mobjBill.����ID: lng��������ID = mobjBill.��������ID: bytӤ���� = mobjBill.Ӥ����
        lng���˿���ID = mobjBill.����ID: str�ѱ� = mobjBill.�ѱ�
        If Not mrsInfo Is Nothing Then
           If mrsInfo.State = 1 Then intInsure = Val(Nvl(mrsInfo!����))
        End If
    End If
    
    If lng����ID = 0 Then
        MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If SelectWholeItems(Me, pҽ�����ѹ���, mstrPrivs, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    Err = 0: On Error GoTo Errhand:
    Screen.MousePointer = 11
    
    Set tmpBill = ImportWholeSet(Me, intInsure, rsSel, lng����ID, lng��������ID, bytӤ����, IIf(mint������Դ = 2 And mint��¼���� = 2, 2, 0), chk�Ӱ�.Value = 1, _
        0, mint������Դ, UserInfo.����, NeedName(cbo������.Text), , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    '��������
    '�������Ĳ�����Ϣ
    '����:37500
    With tmpBill
        .����ID = mobjBill.����ID
        .��ҳID = mobjBill.��ҳID
        .����ID = mobjBill.����ID
        .����ID = mobjBill.����ID
        .���� = mobjBill.����
        .��ʶ�� = mobjBill.��ʶ��
        .���� = mobjBill.����
        .�Ա� = mobjBill.�Ա�
        .���� = mobjBill.����
        .�ѱ� = mobjBill.�ѱ�
    End With
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    Dim bln��ҩ As Boolean
    bln��ҩ = False
    With mobjBill
        For i = 1 To .Details.Count - 1
            If .Details(i).�շ���� = "7" Then
                bln��ҩ = True
                Exit For
            End If
            Exit For
        Next
    End With
    curDate = gobjDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.�Ǽ�ʱ�� = curDate
    mobjBill.����Ա��� = UserInfo.���
    mobjBill.����Ա���� = UserInfo.����
    mobjBill.�Ӱ��־ = chk�Ӱ�.Value
    If mobjBill.�ѱ� = "" Then mobjBill.�ѱ� = str�ѱ�
    If mobjBill.����ID = 0 Then mobjBill.����ID = lng���˿���ID
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
    txtDate.Text = Format(curDate, "yyyy-MM-dd HH:mm:ss")
    Bill.Redraw = False
    Bill.ClearBill
    Bill.Rows = mobjBill.Details.Count + 1
    
   ' Call InitBillColumnColor
    '���ʷ��౨��
    mstrWarn = ""
        
   ' Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
        
    '������Ķ����˺�ȷ���ѱ��,�ټ���۸�
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    With Bill
        For i = 1 To .Rows - 1
            .TextMatrix(i, BillCol.��) = i
        Next
    End With
    
    Bill.Redraw = True
    'ˢ�²��˷�����Ϣ
    If mrsInfo.State = 1 Then
        'ˢ�²���Ԥ������Ϣ
        curTotal = GetBillTotal(mobjBill)
        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(mint������Դ = 1, 0, mlng��ҳID))
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!Ԥ�����
            cmdCancel.Tag = rsTmp!�������
            txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
        End If
    End If
    '���¼���ͳ����
    Call ReCalcInsure
    '����б༭����������ɫ
    Bill.SetColColor BillCol.���, &HE7CFBA
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.��־, &HE0E0E0
    Screen.MousePointer = 0
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lngִ�п���ID As Long
    Dim rsTemp As ADODB.Recordset, dbl���� As Double, dbl�۸� As Double
    Dim strSql As String, blnNOMoved As Boolean
    '����Ϊ�����շ���Ŀ
    '����:27327
    Err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    If mint��¼���� = 1 Or (mint��¼���� = 2 And mint������Դ = 1) Then
        blnNOMoved = gobjDatabase.NOMoved("������ü�¼", mstrInNO, "��¼����=", mint��¼����)
    Else
        blnNOMoved = gobjDatabase.NOMoved("סԺ���ü�¼", mstrInNO, "��¼����=", mint��¼����)
    End If

    
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:���,����,�շ�ϸĿID,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    Err = 0: On Error GoTo Errhand:
   If mbytInState = 1 Then
        '�鿴
        
         strSql = _
        " Select Nvl(A.�۸񸸺�,A.���) as ���,A.�շ����,A.��������,A.�շ�ϸĿID,A.ִ�в���ID," & _
        "       ��   Avg(Nvl(A.����,1)) as ����, Avg(A.����) ����, Sum(A.��׼����) as ����,B.ִ�п���, B.�Ƿ���,M.��������" & _
        " From " & IIf(blnNOMoved, "H" & mstrFeeTab, mstrFeeTab & " A") & ",�շ���ĿĿ¼ B,�������� M" & _
        " Where  A.��¼״̬  IN(0,1,3)  And A.NO=[1]  And A.��¼����=[2] " & _
        "               And a.�շ�ϸĿID=b.ID And a.�շ�ϸĿID=M.����ID(+) " & _
                        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
        "  Group by Nvl(A.�۸񸸺�,A.���),A.�շ����,A.�շ�ϸĿID,A.��������,A.ִ�в���id,B.ִ�п���,B.�Ƿ���,M.��������" & _
        " Order by ���"
        If mstrTime <> "" Then
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mstrInNO, mint��¼����, CDate(mstrTime))
        Else
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mstrInNO, mint��¼����)
        End If
        With rsTemp
            Do While Not .EOF
                 '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                If InStr(1, ",4,5,6,7,", "," & Nvl(!�շ����)) > 0 Then
                    lngִ�п���ID = 0
                ElseIf InStr(1, ",0,4", Val(Nvl(!ִ�п���))) > 0 Then
                    lngִ�п���ID = Val(Nvl(!ִ�в���ID))
                Else
                    lngִ�п���ID = 0
                End If
                dbl�۸� = 0
                If Val(Nvl(!�Ƿ���)) = 1 Then
                    If InStr(1, "5,6,7", Nvl(!�շ����)) > 0 Or (Nvl(!�շ����) = "4" And Val(Nvl(!��������)) = 1) Then
                        'ҩƷ,����������Ϊ��ȱʡ�۸�,���Բ�����(ͨ��������)
                        dbl�۸� = 0
                    Else
                        dbl�۸� = Val(Nvl(!����))
                    End If
                End If
                strItems = strItems & "|" & Val(Nvl(!���)) & "," & Val(Nvl(!��������)) & "," & Val(Nvl(!�շ�ϸĿID)) & "," & Val(Nvl(!����)) & "," & Val(Nvl(!����)) & "," & dbl�۸� & "," & lngִ�п���ID
                .MoveNext
            Loop
        End With
         If strItems = "" Then
            MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Sub
        End If
        strItems = Mid(strItems, 2)
   Else
        With mobjBill
            strItems = ""
            For i = 1 To .Details.Count
                 '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                If InStr(1, ",4,5,6,7,", "," & .Details(i).Detail.���) > 0 Then
                    lngִ�п���ID = 0
                ElseIf InStr(1, ",0,4", .Details(i).Detail.ִ�п���) > 0 Then
                    lngִ�п���ID = .Details(i).ִ�в���ID
                Else
                    lngִ�п���ID = 0
                End If
                '����:52349
                dbl���� = .Details(i).����
                dbl�۸� = IIf(.Details(i).Detail.���, .Details(i).InComes(1).��׼����, 0)
                If InStr(",5,6,7,", .Details(i).�շ����) > 0 And mblnҩ����λ Then
                    dbl���� = Format(dbl���� * .Details(i).Detail.ҩ����װ, "0.00000")
                    dbl�۸� = Format(dbl�۸� / .Details(i).Detail.ҩ����װ, gSysPara.Price_Decimal.strFormt_VB)
                End If
                strItems = strItems & "|" & .Details(i).��� & "," & .Details(i).�������� & "," & .Details(i).�շ�ϸĿID & "," & .Details(i).���� & "," & dbl���� & "," & dbl�۸� & "," & lngִ�п���ID
             Next
             
             If strItems = "" Then
                MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
            strItems = Mid(strItems, 2)
        End With
    End If
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, 1150, mstrPrivs, strItems)
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Function ImportWholeSet(frmParent As Object, ByVal intInsure As Integer, rsSel As ADODB.Recordset, Optional lng����ID As Long = 0, _
    Optional lng��������ID As Long = 0, Optional bytӤ���� As Byte, _
    Optional int�����־ As Integer, Optional bln�Ӱ�Ӽ� As Boolean = False, _
    Optional ByVal lngUnitID As Long, Optional int��Χ As Integer, _
    Optional str������ As String = "", Optional str������ As String = "", _
    Optional lng��ҳId As Long = 0, _
    Optional ByVal strҩƷ�۸�ȼ� As String, _
    Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ�۸�ȼ� As String) As ExpenseBill
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���õ��ݵ����ݶ�����
    '���:rsSel-ѡ�еĳ�����Ŀ
    '       lngUnitID    ��ǰ��������ID
    '      int��Χ=1.����,2-סԺ
    '      intInsure:����
    '����:
    '����:��ŵ�����Ϣ�ĵ��ݶ���
    '����:���˺�
    '����:2010-09-02 16:17:54
    '˵��:��Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
    '       ��������ͣ���շ�ϸĿ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue(0 To 10) As String, strSubItem As String, str�շ�ϸĿID As String, j As Long
    Dim rsItems As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng���˿���ID As Long, strժҪ As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim lngDoUnit As Long
    Dim i As Long, intCurNo As Integer
    Dim int��� As Integer
    Dim strͣ����Ŀ��� As String
    Dim dblAllTime As Double, dbl�Ӱ�Ӽ��� As Double
    Dim colSerial As New Collection
    Dim bytType As Byte '0-����;1-סԺ;2-�����סԺ
    Dim strTable  As String
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    '�۸�ȼ�
    If strҩƷ�۸�ȼ� <> "" Or str���ļ۸�ȼ� <> "" Or str��ͨ�۸�ȼ� <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [14])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [15])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And d.�۸�ȼ� = [16])" & vbNewLine & _
            "            Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From �շѼ�Ŀ" & vbNewLine & _
            "                                Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [14])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [15])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And �۸�ȼ� = [16])))))"
    Else
        strWherePriceGrade = " And d.�۸�ȼ� Is Null "
    End If
    
    With rsSel
        str�շ�ϸĿID = "": j = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Len(str�շ�ϸĿID) > 1990 And j <= 10 Then
                strValue(j) = Mid(str�շ�ϸĿID, 2)
                strSubItem = strSubItem & " Union ALL " & _
                " Select Column_Value as �շ�ϸĿID From Table(f_Num2List([" & j + 1 & "])) B "
                str�շ�ϸĿID = "": j = j + 1
            End If
            str�շ�ϸĿID = str�շ�ϸĿID & "," & Val(Nvl(!�շ�ϸĿID))
            .MoveNext
        Loop
    End With
    
    If str�շ�ϸĿID <> "" Then
        If j > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From �շ���ĿĿ¼ Where id in (" & Mid(str�շ�ϸĿID, 2) & ")"
        Else
            strValue(j) = Mid(str�շ�ϸĿID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as �շ�ϸĿID From Table(f_Num2List([" & j + 1 & "])) B "
        End If
    End If
    
    gstrSQL = "" & _
       "   Select A.����id, A.����id, A.���д���, A.�������� " & _
       "   From �շѴ�����Ŀ A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.����id = D.�շ�ϸĿid "
    Set rsOthers = gobjDatabase.OpenSQLRecord(gstrSQL, "mdlInExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    strSubItem = Mid(strSubItem, 11)
    strTable = " Select [13] as ����ID,�շ�ϸĿID From (" & strSubItem & ")"
    
    gstrSQL = "" & _
    " Select  X.ҩƷID,W.����ID,W.��������," & _
    "       nvl(G.�ѱ�,F.�ѱ�) as �ѱ�,NVL(NVL(Q.����,G.����),F.����) ����,NVL(NVL(Q.�Ա�,G.�Ա�),F.�Ա�) �Ա�,NVL(NVL(Q.����,G.����),F.����) ����,F.������," & _
    "       G.��Ժ���� as ����,F.סԺ�� as ��ʶ��,F.����ID,G.��ҳID,G.��ǰ����ID as ���˲���ID,G.��Ժ����ID as ���˿���ID," & _
    "       G.��������,B.��� as �շ����,A.�շ�ϸĿID," & _
    "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(H.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
    "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������, B.��������  ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
    "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
    "       Decode(B.���,'4',1,X." & mstrҩ����װ & ") as ҩ����װ,Decode(B.���,'4',B.���㵥λ,X." & mstrҩ����λ & ") as ҩ����λ," & _
    "       Decode(b.���,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,B.¼������, " & _
    "       M1.���� as ���Ʊ���,M1.���� as ��������,X.��ҩ��̬,x.����ϵ��,M1.���㵥λ as ������λ" & _
    "   From  (" & strTable & ") A ,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,������Ϣ F, " & _
    "          ������ҳ G,���˹Һż�¼ Q,�շ���Ŀ���� H,�շ���Ŀ���� E1,�������� W,ҩƷ��� X,������ĿĿ¼ M1" & _
    " Where  A.�շ�ϸĿID=D.�շ�ϸĿID And A.�շ�ϸĿID=B.ID " & _
    "       And b.���=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) and X.ҩ��ID=M1.ID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
    "       And A.�շ�ϸĿID=H.�շ�ϸĿID(+) And H.����(+)=1 And H.����(+)=[12]" & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "       And A.����ID=F.����ID(+) And F.����ID=G.����ID(+) And F.����ID=Q.����ID(+) And F.�����=Q.�����(+) And F.��ҳID=G.��ҳID(+)" & _
    "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
    "       And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
            strWherePriceGrade
    
    If Not gSysPara.bln���뷢ҩ Then
        gstrSQL = "Select * From (" & gstrSQL & ")"
    Else
        '���뷢ҩʱ�ſ�ʱ�ۺͷ���ҩƷ������
        gstrSQL = "Select * From (" & gstrSQL & ") Where Not( Instr(',5,6,7,',�շ����)>0 And (����=1 Or �Ƿ���=1))"
    End If
    Set rsItems = gobjDatabase.OpenSQLRecord(gstrSQL, "mdlExse", strValue(0), strValue(1), strValue(2), strValue(3), _
        strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10), _
        IIf(gSysPara.bytҩƷ������ʾ = 1, 3, 1), lng����ID, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�)
    'û�м�¼���ǿյ���
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    
    With rsSel
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
NextRecord: Do While Not .EOF
            '����շ���Ŀ�Ƿ�ͣ�û���������ﲡ��
            '����ͣ��ʱ,��������
            rsItems.Filter = "�շ�ϸĿID=" & Val(Nvl(!�շ�ϸĿID))
            If rsItems.EOF Then 'δ�ҵ�.������
                 .MoveNext
                GoTo NextRecord:
            End If
            If InStr(",5,6,7,", rsItems!�շ����) = 0 Then
                If InStr(1, strͣ����Ŀ��� & ",", "," & !�������� & ",") > 0 Then
                    .MoveNext
                    GoTo NextRecord
                Else
                    If Not CheckFeeItemAvailable(!�շ�ϸĿID, 2) Then
                        strͣ����Ŀ��� = strͣ����Ŀ��� & "," & !���
                        MsgBox "�����շ���Ŀ�еĵ�" & !��� & "���շ���Ŀ:" & rsItems!���� & "" & vbCrLf & _
                            "��ͣ�û��ٷ����ڲ���,�����ᱻ����." & IIf(IsNull(!��������), "����д�����Ŀ,Ҳ���ᱻ����.", ""), vbInformation, gstrSysName
                        .MoveNext
                        GoTo NextRecord
                    End If
                End If
            End If
            
            If i = 1 Then
                objBill.NO = ""
                objBill.����ID = Val(Nvl(rsItems!����ID))
                objBill.��ҳID = Val(Nvl(rsItems!��ҳID))
                objBill.����ID = Val(Nvl(rsItems!���˲���ID))
                objBill.����ID = Val(Nvl(rsItems!���˿���id))
                objBill.���� = Nvl(rsItems!����)
                objBill.�Ա� = Nvl(rsItems!�Ա�)
                objBill.���� = Nvl(rsItems!����)
                objBill.��ʶ�� = Val(Nvl(rsItems!��ʶ��))
                objBill.���� = "" & rsItems!����
                objBill.�ѱ� = Nvl(rsItems!�ѱ�)
                objBill.�����־ = int�����־
                objBill.�Ӱ��־ = IIf(bln�Ӱ�Ӽ�, 1, 0)
                objBill.Ӥ���� = bytӤ����
                objBill.��������ID = lng��������ID
                objBill.������ = str������
                objBill.������ = str������
                objBill.����Ա��� = UserInfo.���
                objBill.����Ա���� = UserInfo.����
                objBill.����ʱ�� = gobjDatabase.Currentdate   ' !����ʱ��
                objBill.�Ǽ�ʱ�� = gobjDatabase.Currentdate
                objBill.�ಡ�˵� = 0
            End If
            '�����շ�ϸĿ=====================================================
            Set objBillDetail = New BillDetail
            Set objBillDetail.Detail = New Detail
                
            '������źʹ�������
            intCurNo = intCurNo + 1
            objBillDetail.��� = intCurNo
            colSerial.Add Array(Val(Nvl(!�շ�ϸĿID)), intCurNo), "_" & !���  '��¼ԭ������ڵ��к�
            objBillDetail.�������� = Nvl(!��������, 0) '��Ϊ������������,�ȼ�¼ԭ����,�����ٴ���
            objBillDetail.�շ���� = Nvl(rsItems!�շ����)
            objBillDetail.�շ�ϸĿID = Val(Nvl(!�շ�ϸĿID))
            objBillDetail.���㵥λ = Nvl(rsItems!���㵥λ)
            objBillDetail.���� = IIf(Val(Nvl(!����)) = 0, 1, Val(Nvl(!����)))
            
            If InStr(",5,6,7,", rsItems!�շ����) > 0 And mblnҩ����λ Then
                objBillDetail.���� = Nvl(!����, 0) / Nvl(rsItems!ҩ����װ, 1)
            Else
                objBillDetail.���� = Nvl(!����, 0)
            End If
            objBillDetail.��ҩ���� = ""
            
            objBillDetail.���ӱ�־ = 0 ' IIf(IsNull(!���ӱ�־), 0, !���ӱ�־)
            objBillDetail.ժҪ = "" ' IIf(IsNull(!ժҪ), "", !ժҪ)
            '���ĺ�ҩƷ����
            '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
            If objBillDetail.�շ���� = "4" Then
                lngDoUnit = IIf(mlng���ϲ��� > 0, mlng���ϲ���, objBill.����ID)
                If lngDoUnit = 0 Then lngDoUnit = lng��������ID
            ElseIf InStr(1, ",5,6,7,", "," & objBillDetail.�շ���� & ",") > 0 Then
                '����Ƿ���ȱʡҩ��,����ȱʡ��,��ȡȱʡҩ��,����ȡ��һ��ҩ��
                '����:36736
                Select Case objBillDetail.�շ����
                    Case "5"
                        If mlng��ҩ�� > 0 Then lngDoUnit = mlng��ҩ��
                    Case "6"
                        If mlng��ҩ�� > 0 Then lngDoUnit = mlng��ҩ��
                    Case "7"
                        If mlng��ҩ�� > 0 Then lngDoUnit = mlng��ҩ��
                End Select
            Else
                If Val(Nvl(!ִ�п���ID)) <> 0 Then lngDoUnit = Val(Nvl(!ִ�п���ID))
            End If
            
            '���˿���ID
            lng���˿���ID = objBill.����ID
            If lng���˿���ID = 0 Then lng���˿���ID = lng��������ID
            objBillDetail.Detail.ִ�п��� = IIf(IsNull(rsItems!ִ�п���), 0, rsItems!ִ�п���)
            objBillDetail.ִ�в���ID = Val(Nvl(!ִ�п���ID))
            
           lngDoUnit = Get�շ�ִ�п���ID(Val(Nvl(rsItems!����ID)), Val(Nvl(rsItems!��ҳID)), objBillDetail.�շ����, objBillDetail.�շ�ϸĿID, _
                        objBillDetail.Detail.ִ�п���, lng���˿���ID, lng��������ID, int��Χ, lngDoUnit, 1, 1, , objBillDetail.ִ�в���ID)          '0-ҽ���������,1-���ѳ������
            objBillDetail.ִ�в���ID = lngDoUnit

            If InStr(",5,6,7,", rsItems!�շ����) > 0 And gSysPara.bln���뷢ҩ Then
                objBillDetail.ִ�в���ID = 0
            End If
            objBillDetail.Detail.ID = !�շ�ϸĿID
            objBillDetail.Detail.���� = Nvl(rsItems!����)
            objBillDetail.Detail.��� = (Val(Nvl(rsItems!�Ƿ���)) = 1)
            objBillDetail.Detail.�������� = 0
            objBillDetail.Detail.���д��� = 0
            If objBillDetail.�������� <> 0 Then
                'A.����id, A.����id, A.���д���, A.�������� "
                rsOthers.Filter = "����ID=" & colSerial("_" & !��������)(0) & " And ����ID=" & objBillDetail.�շ�ϸĿID
                If Not rsOthers.EOF Then
                    objBillDetail.Detail.�������� = Val(Nvl(rsOthers!��������))
                    objBillDetail.Detail.���д��� = Val(Nvl(rsOthers!���д���))
                End If
            End If
            
            objBillDetail.Detail.��� = Nvl(rsItems!���)
            objBillDetail.Detail.���㵥λ = Nvl(rsItems!���㵥λ)
            
            objBillDetail.Detail.ҩ����λ = Nvl(rsItems!ҩ����λ)
            objBillDetail.Detail.ҩ����װ = Val(Nvl(rsItems!ҩ����װ))
            
            objBillDetail.Detail.�Ӱ�Ӽ� = 0 ' (IIf(IsNull(!�Ӱ�Ӽ�), 0, !�Ӱ�Ӽ�) = 1)
            objBillDetail.Detail.��� = Nvl(rsItems!���)
            objBillDetail.Detail.������� = Nvl(rsItems!�������)
            objBillDetail.Detail.���� = Nvl(rsItems!����)
            objBillDetail.Detail.��Ʒ�� = Nvl(rsItems!��Ʒ��)
            objBillDetail.Detail.���ηѱ� = (Val(Nvl(rsItems!���ηѱ�)) = 1)
            objBillDetail.Detail.˵�� = ""
            objBillDetail.Detail.������� = IIf(IsNull(rsItems!�������), 0, rsItems!�������)
            objBillDetail.Detail.���� = IIf(IsNull(rsItems!��������), "", rsItems!��������)
            
            If InStr(",5,6,7,", rsItems!�շ����) > 0 Then
                objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                objBillDetail.Detail.�������� = Get��������(objBillDetail.Detail.ID)
            End If
            objBillDetail.Detail.¼������ = Val(Nvl(rsItems!¼������))
            objBillDetail.Detail.ҩ��ID = Val(Nvl(rsItems!ҩ��ID))
            objBillDetail.Detail.��� = Val(Nvl(rsItems!�Ƿ���)) = 1
            objBillDetail.Detail.���� = Val(Nvl(rsItems!����)) = 1
            objBillDetail.Detail.�������� = Val(Nvl(rsItems!��������)) = 1
            objBillDetail.Detail.Ҫ������ = 0
            objBillDetail.Detail.��ҩ��̬ = Val(Nvl(rsItems!��ҩ��̬))
            '����:41136
            strժҪ = objBillDetail.ժҪ
            '83877,Ƚ����,2015-4-14,���ﲡ�˸���"���˹Һż�¼"��ȡ"����",����GetItemInfo
'            If lng����ID <> 0 Then'90304
                strժҪ = gclsInsure.GetItemInfo(intInsure, lng����ID, objBillDetail.�շ�ϸĿID, strժҪ, 2, , "|1")
                objBillDetail.ժҪ = strժҪ
'            Else
'                objBillDetail.ժҪ = ""
'            End If
            
            '����۸񲿷�=====================================================
            rsItems.MoveFirst
            Set objBillDetail.InComes = New BillInComes
            Do While Not rsItems.EOF
                '�������еļ۸��������¼���
                If InStr(",5,6,7,", rsItems!�շ����) > 0 Or (rsItems!�շ���� = "4" And Nvl(rsItems!��������, 0) = 1) Then
                    '----------------------------------------------------------------------------------------------
                    'ʱ��ҩƷ����۸�(�����ɲ�����)
                    dblAllTime = Val(Nvl(!����))     '�������ۼ�����
                    If dblAllTime <> 0 Or Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                        Set rsPrice = gobjDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                    "��ȡҩƷ��ǰ�ۼ�", CLng(!�շ�ϸĿID), objBillDetail.ִ�в���ID, dblAllTime)
                        If rsPrice.EOF Then
                            '��ȡ�۸�ʧ��
                            objBillIncome.��׼���� = 0
                        Else
                            strPrice = Nvl(rsPrice!Price) & "|||"
                            varPrice = Split(strPrice, "|")
                            objBillIncome.��׼���� = Val(varPrice(0))
                            dblʣ������ = Val(varPrice(2))
                            
                            If dblʣ������ <> 0 And Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                                '����δ�ֽ����
                                objBillIncome.��׼���� = 0
                            End If
                        End If
                    Else
                        objBillIncome.��׼���� = 0
                    End If
                ElseIf Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                    If Abs(Val(Nvl(!����))) > Abs(Val(Nvl(rsItems!�ּ�))) Or Abs(Val(Nvl(!����))) = 0 Then
                        objBillIncome.��׼���� = Val(Nvl(rsItems!ȱʡ�۸�))
                    Else
                        objBillIncome.��׼���� = Val(Nvl(!����))
                    End If
                Else
                    objBillIncome.��׼���� = Val(Nvl(rsItems!�ּ�))
                End If

                If InStr(",5,6,7,", rsItems!�շ����) > 0 And mblnҩ����λ Then
                    objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(rsItems!ҩ����װ, 1), gSysPara.Price_Decimal.strFormt_VB)
                Else
                    objBillIncome.��׼���� = Format(objBillIncome.��׼����, gSysPara.Price_Decimal.strFormt_VB)
                End If
                
                objBillIncome.�ּ� = Val(Nvl(rsItems!�ּ�))  '�ּ�ԭ�۶�ҩƷ�������
                objBillIncome.ԭ�� = Val(Nvl(rsItems!ԭ��))
                objBillIncome.������ĿID = Val(Nvl(rsItems!������ID))
                objBillIncome.������Ŀ = Nvl(rsItems!������Ŀ)
                objBillIncome.�վݷ�Ŀ = Nvl(rsItems!�ַ�Ŀ)
                
                'Ӧ�ս��=����*����*����
                objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                
'                '�������������ü���(����������Ŀ)
'                If Val(Nvl(rsItems!���ӱ�־)) = 1 And Nvl(rsItems!�շ����) = "F" Then
'                    objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * IIf(IsNull(rsItems!�����շ���), 1, rsItems!�����շ��� / 100)
'                End If
'
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If bln�Ӱ�Ӽ� And Val(Nvl(rsItems!�Ӱ�Ӽ�)) = 1 Then
                    dbl�Ӱ�Ӽ��� = Val(Nvl(rsItems!�Ӱ�Ӽ�)) / 100
                    objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� + objBillIncome.Ӧ�ս�� * dbl�Ӱ�Ӽ���
                End If
                objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gSysPara.Money_Decimal.strFormt_VB)
                
                '����ʵ�ս��
                If Val(Nvl(rsItems!���ηѱ�)) = 1 Then
                    objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                Else
                    objBillIncome.ʵ�ս�� = ActualMoney(objBill.�ѱ�, Val(Nvl(rsItems!������ID)), objBillIncome.Ӧ�ս��, _
                        objBillDetail.�շ�ϸĿID, objBillDetail.ִ�в���ID, objBillDetail.����, dbl�Ӱ�Ӽ���)
                End If
                With objBillIncome
                    objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��
                End With
                '�ж���һ����¼�Ƿ����ڵ�ǰ��
                int��� = !���
                i = i + 1
                rsItems.MoveNext
            Loop
            With objBillDetail
                objBill.Details.Add .InComes, .Detail, .�շ�ϸĿID, .���, .��������, .�շ����, .���㵥λ, .����, .����, .���ӱ�־, .ִ�в���ID, .��ҩ����, .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .Key
                '���뷢ҩʱ,Key����Ϊ1,��ʾ�༭ʱִ�п����в��ɽ���
                If InStr(",5,6,7,", .�շ����) > 0 And gSysPara.bln���뷢ҩ Then
                    objBill.Details(objBill.Details.Count).Key = 1
                End If
            End With
            .MoveNext
        Loop
    End With
     '�����´����������
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).�������� <> 0 Then
            objBill.Details(i).�������� = colSerial("_" & objBill.Details(i).��������)(1)
        End If
    Next
    Set ImportWholeSet = objBill
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetҽ��С��ID() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ʱ��ҽ��С��ID
    '����:�����������Ϊ���˿����Ҳ��ǵ�ǰ���˿���ʱ����ȡ���˱䶯��¼�е����һ�α䶯��ҽ��С��ID
    '        ���򷵻�0,������������д���(�ڴ洢�����д���)
    '����:���˺�
    '����:2011-05-23 10:45:39
    '����:37793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��������ID As Long, rsTemp As ADODB.Recordset, strSql As String
    
    If cbo��������.ListIndex < 0 Then Exit Function
    lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If Not (mlng���˿���id = lng��������ID) Then
        Exit Function
    End If
    'ֻ��סԺ�Ż����
    If Not (mlng����ID <> 0 And mlng��ҳID <> 0) Then Exit Function
    strSql = "" & _
    "   Select ҽ��С��ID From ���˱䶯��¼ A,������Ϣ B " & _
    "   Where  A.����ID=B.����ID  And nvl(A.��ֹԭ��,3)=3 " & _
    "               And A.����ID<>B.��ǰ����ID And A.����ID=[1] and A.��ҳID=[2]  " & _
    "               And A.����ID=[3] "
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID, lng��������ID)
    If rsTemp.EOF = False Then
        zlGetҽ��С��ID = Val(Nvl(rsTemp!ҽ��С��id))
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
 
Private Function CheckValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����������
    '����:�������������-False��������-True
    '����:���ϴ�
    '����:2015/7/14 17:32:53
    '����:86396
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strTable As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    '�ڲ���������ʱ�ż��
    If mstrOriginalNO = "" Then CheckValid = True: Exit Function
    strTable = IIf(mint������Դ = 1, "������ü�¼", "סԺ���ü�¼")
    strSql = "Select 1 From ����ҽ������ A, " & strTable & " B " & _
            " Where A.NO=B.NO And A.��¼����=B.��¼���� And A.NO=[1] And A.���ͺ�=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mstrOriginalNO, mlng���ͺ�)
    
    If rsTmp.RecordCount > 0 Then
        MsgBox "�Ѵ���һ�������ã����ܼ�����ӣ�", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub Find��������(ByVal lng����id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ�������ң����û�������
    '����:
    '����:���ϴ�
    '����:2016/1/18 18:02:57
    '����:91763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strSql As String, rsTmp As ADODB.Recordset
    Dim intIndex As Integer
    On Error GoTo Errhand
    If lng����id <> 0 Then
        For i = 0 To cbo��������.ListCount - 1
            If cbo��������.ItemData(i) = lng����id Then
                cbo��������.ListIndex = i: Exit Sub
            End If
        Next
    End If
    'û�ҵ�,�Զ�����
    strSql = "Select ID,����,����,���� From ���ű� Where ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��������", lng����id)
    If rsTmp.RecordCount = 0 Then Exit Sub
    If Not rsTmp.EOF Then
        '111904:���ϴ�,2017/8/13,����Ȩ���ж��Ƿ��������п���
        If InStr(mstrPrivs, ";���п���;") > 0 Then cbo��������.RemoveItem cbo��������.ListCount - 1
        cbo��������.AddItem IIf(zlIsShowDeptCode, Nvl(rsTmp!����) & "-", "") & Nvl(rsTmp!����)
        cbo��������.ItemData(cbo��������.ListCount - 1) = rsTmp!ID
        intIndex = cbo��������.NewIndex
        If InStr(mstrPrivs, ";���п���;") > 0 Then
            cbo��������.AddItem "�������ҡ�"
            cbo��������.ItemData(cbo��������.ListCount - 1) = 0
        End If
        cbo��������.ListIndex = intIndex
    End If
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function IsCan���֧��(objSquareCard As Object, _
    ByVal lng����ID As Long, ByVal lng��ҳId As Long) As Boolean
    '���ܣ��ж��Ƿ��ܹ���ʾ"���֧��"��ť
    '˵����
    '   �����ǰû�����ô����˻���������֧����ʽ���Ҹò���Ҳû��Ԥ�����ʱ����ʾ"���֧��"��ť
    Dim objCards As Cards, rsTemp As ADODB.Recordset
    Dim curMoney As Currency
    
    On Error GoTo Errhand
    If objSquareCard Is Nothing Then Exit Function
    
    '1.�Ƿ�������õ������˻���ҽ�ƿ�
    'zlGetCards(ByVal bytType As Byte) As Cards
     '����:��ȡ��Ч�Ŀ�����
     '���:bytType-0-����ҽ�ƿ�;
     '             1-���õ�ҽ�ƿ�,
     '             2-���д��������˻���������
     '             3-���õ������˻���ҽ�ƿ�
    Set objCards = objSquareCard.zlGetCards(3)
    If objCards.Count > 0 Then IsCan���֧�� = True: Exit Function
    
    '2.�Ƿ���Ԥ�����
    Set rsTemp = GetMoneyInfo(lng����ID, lng��ҳId)
    If rsTemp Is Nothing Then Exit Function
    Do While Not rsTemp.EOF
        curMoney = curMoney + (Val(rsTemp!Ԥ�����) - Val(rsTemp!�������))
        rsTemp.MoveNext
    Loop
    IsCan���֧�� = curMoney > 0
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function CreatePublicDrug(ByVal lngSys As Long, _
    cnOracle As ADODB.Connection, ByVal strDbUser As String) As Boolean
    '���ܣ���̬����ҩƷ��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    If Not mobjPublicDrug Is Nothing Then CreatePublicDrug = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
    
    Err = 0: On Error GoTo ErrHandler
    If mobjPublicDrug Is Nothing Then
        MsgBox "ҩƷ����������zlPublicDrug������ʧ�ܣ�����ϵͳ��Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    'Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If mobjPublicDrug.zlInitCommon(lngSys, cnOracle, strDbUser) = False Then
        MsgBox "ҩƷ����������zlPublicDrug����ʼ��ʧ�ܣ�����ϵͳ��Ա��ϵ��", vbInformation, gstrSysName
        Set mobjPublicDrug = Nothing: Exit Function
    End If
    CreatePublicDrug = True
    Exit Function
ErrHandler:
    Set mobjPublicDrug = Nothing
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub CreateDrugPacker()
    '����:����������ҩ��(�Զ���ҩ��)
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugMachine = False

    Err = 0: On Error Resume Next
    If Val(gobjDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        'Ȩ�޼��
        strPrivs = gobjComlib.GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))
        If InStr(";" & strPrivs & ";", ";����;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, gobjComlib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    End If
End Sub

Private Function AddStuffItemFromBarCode(ByVal strBarCode As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������������Ŀ
    '���:strBarCode-��������
    '����:
    '����:���ӳɹ�������True,���򷵻�False
    '����:���˺�
    '����:2017-11-22 13:58:00
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��Ŀid As Long, str��� As String, lng���� As Long
    Dim intInsure As Integer, lng����ID As Long, dblStock As Double
    Dim lng���˿���ID As Long, lngDoUnit As Long, strժҪ As String
    Dim lngRow As Long, blnCancel As Boolean, str��׼��Ŀ As String
    Dim blnAdd As Boolean
    On Error GoTo errHandle
    
    If Trim(strBarCode) = "" Then Exit Function
    strBarCode = Trim(strBarCode)
    
    str��� = "'4'"
    If SelectIsNurse = False Then
        If InStr(mstr�շ����, "'4'") = 0 And mstr�շ���� <> "" Then
            MsgBox "��ǰվ�㲻�߱����������Ͻ����շѻ���ʵ�Ȩ�ޣ���ϵͳ����Ա��ϵ,�ڲ��������п����������ϡ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If Not mrsInfo Is Nothing Then
       If mrsInfo.State = 1 Then
            intInsure = Val(Nvl(mrsInfo!����))
            lng����ID = Val(Nvl(mrsInfo!����ID))
            lngDoUnit = Val(Nvl(mrsInfo!����ID))
       End If
    End If
    
    If intInsure <> 0 Then
        If zl_Check��׼��Ŀ(gclsInsure, intInsure, lng����ID, (mint��¼���� = 1 Or mint������Դ = 1)) Then str��׼��Ŀ = Get������׼��Ŀ(lng����ID, "A.ID")
    End If
    
    mlng���� = -1
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, mint������Դ, intInsure, True, str���, strBarCode, txtBarCode.hWnd, str��׼��Ŀ, -1, IIf(mbln���õǼ�, True, False), True, lng����)
    If lng��Ŀid = 0 Then Exit Function
    mlng���� = lng����
    
    blnAdd = False
    lngRow = mobjBill.Details.Count
    If lngRow >= Bill.Rows - 1 Then
        Bill.MsfObj.Rows = Bill.MsfObj.Rows + 1
        Bill.Row = Bill.Rows - 1
        Call bill_AfterAddRow(Bill.Row)
        Bill.Col = BillCol.��Ŀ
        blnAdd = True
    End If
    Bill.Col = BillCol.��Ŀ
    Bill.SetFocus
    Bill.TxtVisible = True: Bill.Text = lng��Ŀid
    mblnSelect = True
    Call bill_KeyDown(13, 0, blnCancel)
    Bill.SetFocus
    If blnCancel Then
        Bill.Text = "": Bill.TxtVisible = False: mblnSelect = False
        If blnAdd And Bill.Rows >= 2 Then
            Bill.Rows = Bill.Rows - 1
            Bill.Row = Bill.Rows - 1
        End If
        AddStuffItemFromBarCode = False: Exit Function
    End If
    mblnSelect = False
    AddStuffItemFromBarCode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function CheckYBFeeItemVerfiy(ByVal lng��ID As Long, ByVal intInsure As Integer, objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ��������Ŀ�Ƿ��������Ƿ��������Ƿ�Ϸ�
    '���:lng��ID-����ID
    '     intInsure-����
    '     blnҪ������-ָ����Ŀ�Ƿ�Ҫ������
    '     lng��ĿID-��ĿID
    '����:
    '����:�Ϸ�������true,���򷵻�False
    '����:���˺�
    '����:2017-11-22 14:22:26
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblNum As Double
    Dim strPriceGrade As String
    On Error GoTo errHandle
    
    If intInsure = 0 Or lng��ID = 0 Then CheckYBFeeItemVerfiy = True: Exit Function
    
    If InStr(",5,6,7,", objDetail.���) > 0 Then
        strPriceGrade = mstrҩƷ�۸�ȼ�
    ElseIf objDetail.��� = "4" Then
        strPriceGrade = mstr���ļ۸�ȼ�
    Else
        strPriceGrade = mstr��ͨ�۸�ȼ�
    End If
            
                            
    '1.����֧����Ŀ��Ӧ���
    If Not zlCheck������۸����(objDetail.ID, Not objDetail.���, strPriceGrade) Then
        If Not ItemExistInsure(lng��ID, objDetail.ID, intInsure) Then
            If gSysPara.intҽ������ = 1 Then
                If MsgBox("��Ŀ""" & objDetail.���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gSysPara.intҽ������ = 2 Then
                MsgBox "��Ŀ""" & objDetail.���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '2.ҽ��������Ŀ�Ƿ��������
    If objDetail.Ҫ������ = False Or mrsMedAudit Is Nothing Then CheckYBFeeItemVerfiy = True: Exit Function
    If Not (mint������Դ = 2 And mint��¼���� = 2) Then CheckYBFeeItemVerfiy = True: Exit Function
    
    mrsMedAudit.Filter = "��ĿID=" & objDetail.ID
    If mrsMedAudit.RecordCount = 0 Then
        MsgBox "��ǰ����δ����׼ʹ�ø���Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    dblNum = FormatEx(Val(Nvl(mrsMedAudit!��������)), 6)
    If dblNum <= 0 Then
        dblNum = dblNum / IIf(mblnҩ����λ, objDetail.ҩ����װ, 1)
        MsgBox "��ǰ����ʹ��[" & objDetail.���� & "]�Ѵﵽ��׼��ʹ������" & FormatEx(dblNum, 5) & "��", vbInformation, gstrSysName
       Exit Function
    End If
     
    CheckYBFeeItemVerfiy = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ReadDrugAndStuffStock(ByVal lng�ⷿID As Long, ByRef objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�������ϵĿ����Ϣ
    '���:lng�ⷿID-�ⷿID
    '����:objDetail-Detail����
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-10 09:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblStock As Double, strҩ��IDs As String
    
    On Error GoTo errHandle
    If objDetail Is Nothing Then Exit Function
    
    If InStr(",5,6,7,4,", objDetail.���) = 0 Then ReadDrugAndStuffStock = True: Exit Function
    If objDetail.��� = "4" And objDetail.�������� = False Then ReadDrugAndStuffStock = True: Exit Function

    If objDetail.��� = "4" And objDetail.�������� Then
        dblStock = GetStock(objDetail.ID, lng�ⷿID, 0, , , objDetail.����)  ''29680
        objDetail.��� = dblStock
        Call ShowStock(objDetail.����, objDetail.���)
        ReadDrugAndStuffStock = True: Exit Function
    End If
    
    '��ȡҩƷ�����Ϣ
    If InStr(",5,6,7,", objDetail.���) > 0 Then
        '��ǰ��ҩƷ���
        dblStock = GetStock(objDetail.ID, lng�ⷿID, 0)  '29680
        If mblnҩ����λ Then
            dblStock = dblStock / objDetail.ҩ����װ
        End If
        objDetail.��� = dblStock
        Call ShowStock(objDetail.����, objDetail.���)
    End If

    ReadDrugAndStuffStock = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
