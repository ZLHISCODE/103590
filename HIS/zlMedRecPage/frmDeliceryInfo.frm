VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeliceryInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������¼��"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9690
   Icon            =   "frmDeliceryInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdGetDeliceryInfo 
      Caption         =   "��ȡ��������Ϣ"
      Height          =   300
      Left            =   120
      TabIndex        =   171
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   3
      Left            =   240
      TabIndex        =   93
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   3
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   111
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   3
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   105
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   3
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   3
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   3
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   3
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   3
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   3
         Left            =   240
         TabIndex        =   109
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   3
         Left            =   1200
         TabIndex        =   108
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   107
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   110
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   3
         Left            =   9000
         TabIndex        =   106
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   3
         Left            =   6720
         TabIndex        =   104
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   3
         Left            =   3480
         TabIndex        =   102
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   100
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   3
         Left            =   6720
         TabIndex        =   98
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   3
         Left            =   3480
         TabIndex        =   96
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   94
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   2
      Left            =   240
      TabIndex        =   74
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   2
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   92
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   2
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   86
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   2
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   2
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   2
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   2
         Left            =   240
         TabIndex        =   90
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   89
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   88
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   91
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   2
         Left            =   9000
         TabIndex        =   87
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   2
         Left            =   6720
         TabIndex        =   85
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   2
         Left            =   3480
         TabIndex        =   83
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   81
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   2
         Left            =   6720
         TabIndex        =   79
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   2
         Left            =   3480
         TabIndex        =   77
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   75
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   6
      Left            =   240
      TabIndex        =   150
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   6
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   168
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   6
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   162
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   6
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   160
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   6
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   158
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   6
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   6
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   154
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   6
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   6
         Left            =   240
         TabIndex        =   166
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   6
         Left            =   1200
         TabIndex        =   165
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   164
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   6
         Left            =   270
         TabIndex        =   167
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   6
         Left            =   9000
         TabIndex        =   163
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   6
         Left            =   6720
         TabIndex        =   161
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   159
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   157
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   6
         Left            =   6720
         TabIndex        =   155
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   153
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   151
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   5
      Left            =   240
      TabIndex        =   131
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   5
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   149
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   5
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   143
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   5
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   141
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   5
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   139
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   5
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   5
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   5
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   5
         Left            =   240
         TabIndex        =   147
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   5
         Left            =   1200
         TabIndex        =   146
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   145
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   148
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   5
         Left            =   9000
         TabIndex        =   144
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   5
         Left            =   6720
         TabIndex        =   142
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   5
         Left            =   3480
         TabIndex        =   140
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   138
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   5
         Left            =   6720
         TabIndex        =   136
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   5
         Left            =   3480
         TabIndex        =   134
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   132
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   4
      Left            =   240
      TabIndex        =   112
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   4
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   130
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   4
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   124
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   4
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   4
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   4
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   4
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   4
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   4
         Left            =   240
         TabIndex        =   128
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   4
         Left            =   1200
         TabIndex        =   127
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   126
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   4
         Left            =   270
         TabIndex        =   129
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   4
         Left            =   9000
         TabIndex        =   125
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   4
         Left            =   6720
         TabIndex        =   123
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   4
         Left            =   3480
         TabIndex        =   121
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   119
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   4
         Left            =   6720
         TabIndex        =   117
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   4
         Left            =   3480
         TabIndex        =   115
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   113
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   1
      Left            =   240
      TabIndex        =   55
      Top             =   3840
      Width           =   9255
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   1
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   1
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   67
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   1
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   73
         Top             =   3435
         Width           =   975
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   1
         Left            =   240
         TabIndex        =   71
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   70
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   56
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   58
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   60
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   62
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   64
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   66
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   1
         Left            =   9000
         TabIndex        =   68
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   72
         Top             =   3480
         Width           =   810
      End
   End
   Begin VB.Frame frḁ����Ϣ 
      Caption         =   "̥����Ϣ"
      Height          =   3855
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar���� 
         Height          =   330
         Index           =   0
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   54
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txtӤ������ 
         Height          =   330
         Index           =   0
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   48
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmbӤ���Ա� 
         Height          =   300
         Index           =   0
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb����ȱ�� 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb������� 
         Height          =   300
         Index           =   0
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb����̥λ 
         Height          =   300
         Index           =   0
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb���䷽ʽ 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   330
         Index           =   0
         Left            =   1200
         TabIndex        =   51
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183828483
         CurrentDate     =   38518
      End
      Begin ZL9BillEdit.BillEdit bill���������� 
         Height          =   1695
         Index           =   0
         Left            =   240
         TabIndex        =   52
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
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
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar���� 
         AutoSize        =   -1  'True
         Caption         =   "Apgar����"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   53
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         Caption         =   "g"
         Height          =   180
         Index           =   0
         Left            =   9000
         TabIndex        =   49
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblӤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   0
         Left            =   6720
         TabIndex        =   47
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblӤ���Ա� 
         AutoSize        =   -1  'True
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   0
         Left            =   3480
         TabIndex        =   45
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl����ȱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ȱ��"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   0
         Left            =   6720
         TabIndex        =   41
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl����̥λ 
         AutoSize        =   -1  'True
         Caption         =   "����̥λ"
         Height          =   180
         Index           =   0
         Left            =   3480
         TabIndex        =   39
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl���䷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   8520
      TabIndex        =   170
      Top             =   7920
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   7320
      TabIndex        =   169
      Top             =   7920
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Tag             =   "##:##:##"
      Top             =   1320
      Width           =   9495
      Begin VB.TextBox txt�ܲ���ʱ�� 
         Height          =   270
         Left            =   1200
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskDate����1 
         Height          =   330
         Left            =   1200
         TabIndex        =   21
         Tag             =   "##:##:##"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483646
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt������� 
         Height          =   330
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   15
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txt̥�� 
         Height          =   330
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "1"
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txt�����Ѫ�� 
         Height          =   330
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmb̥�� 
         Height          =   300
         ItemData        =   "frmDeliceryInfo.frx":000C
         Left            =   8055
         List            =   "frmDeliceryInfo.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   315
         Width           =   1215
      End
      Begin VB.ComboBox cbo����֢ 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1515
         Width           =   4755
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   8055
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1095
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskDate����2 
         Height          =   330
         Left            =   4800
         TabIndex        =   23
         Tag             =   "##:##:##"
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483646
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskDate����3 
         Height          =   300
         Left            =   8040
         TabIndex        =   25
         Tag             =   "##:##:##"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483646
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����������"
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   31
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���Ʋ���֢"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   33
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lbl�ܲ���ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�ܲ���ʱ��"
         Height          =   180
         Left            =   45
         TabIndex        =   26
         Top             =   1155
         Width           =   900
      End
      Begin VB.Label lbl����ʱ��1 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��1"
         Height          =   180
         Left            =   135
         TabIndex        =   20
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl����ʱ��2 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��2"
         Height          =   180
         Left            =   3810
         TabIndex        =   22
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl�����Ѫ�� 
         AutoSize        =   -1  'True
         Caption         =   "�����Ѫ��"
         Height          =   180
         Left            =   3720
         TabIndex        =   28
         Top             =   1155
         Width           =   900
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   225
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl̥�� 
         AutoSize        =   -1  'True
         Caption         =   "̥��"
         Height          =   180
         Left            =   4260
         TabIndex        =   16
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl̥�� 
         AutoSize        =   -1  'True
         Caption         =   "̥��"
         Height          =   180
         Left            =   7425
         TabIndex        =   18
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         Caption         =   "Ml"
         Height          =   180
         Left            =   6060
         TabIndex        =   30
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label lbl����ʱ��3 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��3"
         Height          =   180
         Left            =   6975
         TabIndex        =   24
         Top             =   765
         Width           =   810
      End
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txt��Ժ���� 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txt��Ժ���� 
         Enabled         =   0   'False
         Height          =   330
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txt��Ժ��Ҫ��� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4110
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   623
         Width           =   5190
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   623
         Width           =   1770
      End
      Begin VB.TextBox txtסԺ�� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   263
         Width           =   1215
      End
      Begin VB.TextBox txt������ 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Top             =   263
         Width           =   1770
      End
      Begin VB.Label lbl��Ժ���� 
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   4830
         TabIndex        =   5
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2790
         TabIndex        =   3
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   300
         TabIndex        =   9
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lbl��Ժ���� 
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   7200
         TabIndex        =   7
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl��Ժ��Ҫ��� 
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��Ҫ���"
         Height          =   180
         Left            =   2790
         TabIndex        =   11
         Top             =   690
         Width           =   1080
      End
   End
   Begin MSComctlLib.TabStrip tab̥�� 
      Height          =   4335
      Left            =   120
      TabIndex        =   35
      Top             =   3480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��1��̥��"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDeliceryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mlng����ID As Long                '����ID
Private mlng��ҳID As Long                '��ҳID
Private mblnEditable As Boolean           '�Ƿ���Ա༭
Private mlng����ID As Long                '���˳�Ժ��Ҫ���(��ҽ)�ļ���ID
Private mrs������Ϣ As ADODB.Recordset    '���˷�����Ϣ
Private mrs̥����� As ADODB.Recordset    '������������Ϣ
Private mrs��������� As ADODB.Recordset  '�����������Ϣ
Private mbytModel As Byte                 'mbytModel=1 ����ϵͳ;=2�������Ǽ�
Private mstrSumTime As String

Private Enum BabyDiag
    col���� = 0
    Col���� = 1
    Col��� = 2
End Enum

Private mblnOK As Boolean
Private mintFlag As Integer

Public Function EditDelivery(ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal blnEditable As Boolean, _
                        ByRef rs������Ϣ As ADODB.Recordset, ByRef rs̥����Ϣ As ADODB.Recordset, ByRef rs���������� As ADODB.Recordset, _
                        Optional ByRef blnOK As Boolean, Optional ByVal bytModel As Byte = 1) As Boolean
    Dim ctlTemp As Control

    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID
    mblnEditable = blnEditable
    rs������Ϣ.Filter = "": rs̥����Ϣ.Filter = "": rs����������.Filter = ""
    Set mrs������Ϣ = zlDatabase.CopyNewRec(rs������Ϣ)
    Set mrs̥����� = zlDatabase.CopyNewRec(rs̥����Ϣ)
    Set mrs��������� = zlDatabase.CopyNewRec(rs����������)
    mbytModel = bytModel
    
    mblnOK = False
    Me.Show 1, frmParent
    rs������Ϣ.Filter = "": rs̥����Ϣ.Filter = "": rs����������.Filter = ""
    Set rs������Ϣ = mrs������Ϣ
    Set rs̥����Ϣ = mrs̥�����
    Set rs���������� = mrs���������
    blnOK = mblnOK
    EditDelivery = True
    
End Function

Private Sub bill����������_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    Dim intIndex As Integer
    Dim intCol As Integer
    
    intIndex = tab̥��.SelectedItem.Index
    With bill����������(intIndex - 1)
        For intCol = Col���� To Col���
            .TextMatrix(.Row, intCol) = ""
        Next
        Cancel = True
    End With
End Sub

Private Sub bill����������_CommandClick(Index As Integer)
    Dim intIndex As Integer
    Dim strSql As String, str�Ա� As String
    Dim rsTmp As ADODB.Recordset
    
    intIndex = tab̥��.SelectedItem.Index - 1
    str�Ա� = cmbӤ���Ա�(intIndex).Text
    If str�Ա� Like "*��*" Then
        str�Ա� = "��"
    ElseIf str�Ա� Like "*Ů*" Then
        str�Ա� = "Ů"
    Else
        str�Ա� = ""
    End If
    
    With bill����������(intIndex)
        Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "D", gclsPros.��Ժ����ID, str�Ա�, False)
        If Not rsTmp Is Nothing Then
            SetBillInput intIndex, bill����������(intIndex).Row, rsTmp
        End If
'        .SetFocus
    End With
End Sub

Private Sub bill����������_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim intIndex As Integer
    Dim strFilter As String
    Dim arrFilter As Variant, lngCount As Long
    Dim blnOK As Boolean
    Dim vPoint As POINTAPI
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    intIndex = tab̥��.SelectedItem.Index - 1
    With bill����������(intIndex).MsfObj
        If .Col = Col���� Then
            strFilter = UCase(Replace(Trim(bill����������(intIndex).Text), "'", "''"))
            strFilter = Replace(strFilter, "��", ".")

            If strFilter = "%" Or strFilter = "_" Then
                strSql = "select A.ID,A.ID ��ĿID,A.����,����,A.����,A.����,A.˵��,A.�Ա�����,A.��Ч���� as ������Ч " & _
                          " From ��������Ŀ¼ A  " & _
                          " Where A.���='D' and Rownum<1000  and (a.����ʱ�� is null or a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd'))"
            Else
                If InStr(strFilter, " ") = 0 Then
                    strSql = "select A.ID,A.ID ��ĿID,A.����,����,A.����,A.����,A.˵��,A.�Ա�����,A.��Ч���� as ������Ч " & _
                              " From ��������Ŀ¼ A  " & _
                              " Where A.���='D' and Rownum<1000     and (a.����ʱ�� is null or a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd')) " & _
                              "         and (upper(A.����) like '" & strFilter & "%' or upper(A.����) like '%" & strFilter & "%' or upper(A.����) like '%" & strFilter & "%')  "
                Else
                    arrFilter = Split(strFilter, " ")
                    strSql = ""
                    For lngCount = LBound(arrFilter) To UBound(arrFilter)
                        If Trim(arrFilter(lngCount)) <> "" Then
                            strSql = strSql & " and upper(A.����) like '%" & Trim(arrFilter(lngCount)) & "%'"
                        End If
                    Next
                    strSql = Mid(strSql, 5) 'ȥ����һ��and
                    If Trim(strSql) = "" Then
                        strSql = " upper(A.����) like '" & strSql & "%'"
                    Else
                        strSql = "(" & strSql & ") or upper(A.����) like '" & strFilter & "%'"
                    End If
                    strSql = "select A.ID, A.ID ��ĿID,A.����,����,A.����,A.����,A.˵��,A.�Ա�����,A.��Ч���� as ������Ч " & _
                              " From ��������Ŀ¼ A  " & _
                              " Where A.���='D' and Rownum<1000     and (a.����ʱ�� is null or a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd')) and (" & strSql & ")"
                End If
            End If
            
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)

            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "��������", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
            If blnCancel Or rsTmp Is Nothing Then
                Cancel = True
                Call Beep
                bill����������(intIndex).TxtSetFocus
            Else
                blnOK = SetBillInput(intIndex, bill����������(intIndex).Row, rsTmp)
                If blnOK = False Then
                    Cancel = True
                    bill����������(intIndex).TxtSetFocus
                Else
                    Cancel = False
                    bill����������(intIndex).TxtVisible = False
                End If
            End If
        End If
    End With
End Sub

Private Sub cbo����֢_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyCode = vbKeyDelete Then cbo����֢.ListIndex = -1
End Sub

 
Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyCode = vbKeyDelete Then cbo����.ListIndex = -1
End Sub

Private Sub cmb����ȱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb����̥λ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb���䷽ʽ_Click(Index As Integer)
    Dim intCount As Integer
    Dim i As Integer
    Dim str���� As String
    
    '����22275���� by lesfeng 2009-09-24
    intCount = Val(cmb̥��.Text) - 1
    If intCount < 1 Then Exit Sub
    If mintFlag = 1 Then Exit Sub
        
    If Index = 0 And mintFlag = 0 Then
        str���� = cmb���䷽ʽ(0).Text
        For i = 1 To intCount
            cmb���䷽ʽ(i).Text = str����
        Next
    End If
    
    If Not tab̥��.Tabs(1).Selected Then
        If Index > 0 Then mintFlag = 1
    End If
End Sub

Private Sub cmb���䷽ʽ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb�������_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb̥��_Click()
    Dim intIndex As Integer
    Dim intRow As Integer
    Dim str���� As String
    
    str���� = cmb���䷽ʽ(0).Text
    For intIndex = tab̥��.Tabs.Count To cmb̥��.Text - 1
        initFetus intIndex
        '����22275���� by lesfeng 2009-09-24
        cmb���䷽ʽ(intIndex).Text = str����
    Next
    
    tab̥��.Tabs.Clear
    For intIndex = 1 To cmb̥��.Text
        tab̥��.Tabs.Add intIndex, , "��" & intIndex & "��̥��"
    Next
    
    For intIndex = frḁ����Ϣ.LBound To frḁ����Ϣ.UBound
        frḁ����Ϣ(intIndex).Visible = False
    Next
    
    tab̥��.Tabs(1).Selected = True
    frḁ����Ϣ(0).Visible = True
    
End Sub

Private Sub cmb̥��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbӤ���Ա�_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetDeliceryInfo_Click()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim int̥�� As Integer
    Dim intIndex As Integer
    If mlng����ID <> 0 And mlng��ҳID <> 0 Then
        strSql = "Select ���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,��,����,Ѫ��,����ʱ��,����ʱ��,��ע˵�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] Order by ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��������Ϣ", mlng����ID, mlng��ҳID)
        If rsTmp.EOF Then
            MsgBox "�ò���û����������¼��", vbInformation, gstrSysName
        Else
            If Not mrs������Ϣ.RecordCount < 0 Then
                If MsgBox("�ò��˴�����������¼���Ƿ񸲸ǵ�ǰ������Ϣ��", vbYesNo, Me.Caption) = vbYes Then
                    int̥�� = rsTmp.RecordCount
                    cmb̥��.ListIndex = Cbo.FindIndex(cmb̥��, CStr(int̥��))
                    cmb̥��_Click
                    For intIndex = 1 To tab̥��.Tabs.Count
                        With rsTmp
                            .Filter = "��� = " & intIndex
                            If Not .EOF Then
                                If IsNull(!����ʱ��) Then
                                    dtp����ʱ��(intIndex - 1).Value = zlDatabase.Currentdate
                                Else
                                    dtp����ʱ��(intIndex - 1).Value = CDate(!����ʱ��)
                                End If
                                AddComboItem cmb���䷽ʽ(intIndex - 1), IIf(!���䷽ʽ = "����", "��������", !���䷽ʽ)
                                AddComboItem cmbӤ���Ա�(intIndex - 1), !Ӥ���Ա�
                                txtӤ������(intIndex - 1) = IIf(IsNull(!����), "", Val(Nvl(!����, "")))
                            End If
                        End With
                    Next
                End If
            Else
                int̥�� = rsTmp.RecordCount
                cmb̥��.ListIndex = Cbo.FindIndex(cmb̥��, CStr(int̥��))
                cmb̥��_Click
                For intIndex = 1 To tab̥��.Tabs.Count
                    With rsTmp
                        .Filter = "��� = " & intIndex
                        If Not .EOF Then
                            If IsNull(!����ʱ��) Then
                                dtp����ʱ��(intIndex - 1).Value = zlDatabase.Currentdate
                            Else
                                dtp����ʱ��(intIndex - 1).Value = CDate(!����ʱ��)
                            End If
                            AddComboItem cmb���䷽ʽ(intIndex - 1), IIf(!���䷽ʽ = "����", "��������", !���䷽ʽ)
                            AddComboItem cmbӤ���Ա�(intIndex - 1), !Ӥ���Ա�
                            txtӤ������(intIndex - 1) = IIf(IsNull(!����), "", Val(Nvl(!����, "")))
                        End If
                    End With
                Next
            End If
        End If
    End If
End Sub

Private Sub cmdOk_Click()
    
    If Validate Then
        Save������Ϣ
        mblnOK = True
        Unload Me
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp����ʱ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim intRow As Integer
    Dim intIndex As Integer
    Dim ctlTemp As Control
    Dim arrFetus(0 To 6) As Integer
    
    mintFlag = 0
    
    '��ʼ����Ϣ
    Call InitCombox
    
    If mbytModel = 1 Then
        mrs������Ϣ.Filter = "����=1": mrs������Ϣ.Sort = "��Ϣ��"
        
        Do While Not mrs������Ϣ.EOF
            Select Case mrs������Ϣ!��Ϣ��
                Case "������"
                    txt������.Text = mrs������Ϣ!��Ϣֵ
                Case "סԺ��"
                    txtסԺ��.Text = mrs������Ϣ!��Ϣֵ
                Case "����"
                    txt����.Text = mrs������Ϣ!��Ϣֵ
                Case "��Ժ����"
                    txt��Ժ����.Text = mrs������Ϣ!��Ϣֵ
                Case "��Ժ����"
                    txt��Ժ����.Text = mrs������Ϣ!��Ϣֵ
                Case "��Ҫ���"
                    txt��Ժ��Ҫ���.Text = mrs������Ϣ!��Ϣֵ
            End Select
            mrs������Ϣ.MoveNext
        Loop
        If mrs̥�����.RecordCount <= 0 Then
             dtp����ʱ��(0).Value = CDate(txt��Ժ����.Text)
        End If
    Else
        Me.Height = Me.Height - 1200
        fra������Ϣ.Visible = False
        fra������Ϣ.Move 120, 120
        tab̥��.Move 120, fra������Ϣ.Top + fra������Ϣ.Height + 120
        For intIndex = frḁ����Ϣ.LBound To frḁ����Ϣ.UBound
            frḁ����Ϣ(intIndex).Move tab̥��.Left + 120, tab̥��.Top + 360
        Next
        cmdGetDeliceryInfo.Move cmdGetDeliceryInfo.Left, tab̥��.Top + tab̥��.Height + 120
        cmdOK.Move cmdOK.Left, Me.ScaleHeight - cmdOK.Height - 90
        cmdCancel.Move cmdCancel.Left, Me.ScaleHeight - cmdOK.Height - 90
        If mrs̥�����.RecordCount <= 0 Then
            dtp����ʱ��(0).Value = GetInDate
        End If
    End If
    
    '�����ؼ������С
    For Each ctlTemp In Me.Controls
        If InStr("DTPickerTabStripBillEdit", TypeName(ctlTemp)) = 0 Then
            ctlTemp.FontSize = 10.5
        Else
            ctlTemp.Font.Size = 10.5
        End If
    Next
    
    '��ʼ��̥����
    For intRow = 1 To 7
        arrFetus(intRow - 1) = intRow
    Next
    LoadComboFromArray arrFetus, cmb̥��
    
    '��ʼ�����
    initFetus 0
    For intIndex = frḁ����Ϣ.LBound To frḁ����Ϣ.UBound
        frḁ����Ϣ(intIndex).Visible = False
    Next
    frḁ����Ϣ(0).Visible = True
    
    LoadData
        
    If Not mblnEditable Then
        cmdOK.Visible = False
        cmdCancel.Caption = "�ر�(&C)"
        
        For Each ctlTemp In Me.Controls
            If InStr("Label,TabStrip,CommandButton", TypeName(ctlTemp)) = 0 Then
                ctlTemp.Enabled = False
            End If
        Next
    End If
End Sub

Private Function LoadComboFromArray(ByVal varArray As Variant, cmbTemp As Variant) As Boolean
'�������Ĺ����������ж����б�ֵװ����������
    Dim cmbArray As Variant
    Dim intArray As Long
    Dim intCount As Long
    
    On Error GoTo errHandle
    
    If IsArray(cmbTemp) Then
        cmbArray = cmbTemp
    Else
        'ǿ�����һ������
        cmbArray = Array(cmbTemp)
    End If
    
    For intCount = LBound(cmbArray) To UBound(cmbArray)
        cmbArray(intCount).Clear
        For intArray = LBound(varArray) To UBound(varArray)
            cmbArray(intCount).AddItem varArray(intArray)
        Next
        cmbArray(intCount).ListIndex = 0
    Next
    
    LoadComboFromArray = True
    Exit Function
errHandle:
    LoadComboFromArray = False
End Function
Private Sub InitCombox()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��combox�ؼ���Ϣ,������̥����Ϣ
    '����:���˺�
    '����:2009-02-06 10:08:28
    '-----------------------------------------------------------------------------------------------------------
    '15126
    LoadComboFromArray Array("1.����", "2.������Ѫ", "3.������"), cbo����֢
    LoadComboFromArray Array("1.��", "2.��"), cbo����
    'Ĭ��Ϊ��
    cbo����.ListIndex = -1: cbo����֢.ListIndex = -1
        
End Sub
Private Sub initFetus(i_intIndex As Integer)
    Dim intRow As Integer
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    '15126
    LoadComboFromArray Array("1.���", "2.����", "3.��̥", "4.����������"), cmb�������(i_intIndex)
    LoadComboFromArray Array("1.����ǰ", "2.����ǰ", "3.�����", "4.�����", "5.�����", "6.�����", "7.��λ", "8.����¶", "9.����¶", "10.ϥ��¶"), cmb����̥λ(i_intIndex)
    LoadComboFromArray Array("1.��", "2.��"), cmb����ȱ��(i_intIndex)
    
    strSql = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ���䷽ʽ Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    cmb���䷽ʽ(i_intIndex).Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cmb���䷽ʽ(i_intIndex).AddItem i & "." & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cmb���䷽ʽ(i_intIndex).ListIndex = cmb���䷽ʽ(i_intIndex).NewIndex
                cmb���䷽ʽ(i_intIndex).ItemData(cmb���䷽ʽ(i_intIndex).NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
    Else
        '15126
        LoadComboFromArray Array("1.��������", "2.��������", "3.�ʹ���", "4.���", "5.��ǯ", "6.�γ�", "7.����"), cmb���䷽ʽ(i_intIndex) '���˺�:�γ��Ϊ����,����:21778,by lesfeng 2009-9-23 �����γ����Ϸ������������·��������������������
    End If
    
    strSql = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    cmbӤ���Ա�(i_intIndex).Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cmbӤ���Ա�(i_intIndex).AddItem i & "." & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cmbӤ���Ա�(i_intIndex).ListIndex = cmbӤ���Ա�(i_intIndex).NewIndex
                cmbӤ���Ա�(i_intIndex).ItemData(cmbӤ���Ա�(i_intIndex).NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
    Else
        LoadComboFromArray Array("1.��", "2.Ů", "3.δ֪", "4.����"), cmbӤ���Ա�(i_intIndex)
    End If
    
    txtӤ������(i_intIndex).Text = ""
    txtApgar����(i_intIndex).Text = ""
    If mrs̥�����.RecordCount < i_intIndex + 1 Then
        If mbytModel = 1 Then
            dtp����ʱ��(i_intIndex).Value = CDate(txt��Ժ����.Text)
        Else
            dtp����ʱ��(i_intIndex).Value = dtp����ʱ��(0).Value
        End If
    End If
    '��ʼ���������������
    With bill����������(i_intIndex)
        .AllowAddRow = False
        .Font.Size = 10.5
        .Rows = 5
        .Cols = 3
        .MsfObj.FixedCols = 1
        .Clear
        .TextMatrix(0, col����) = "�������":  .ColData(0) = 5:  .ColWidth(0) = 1400:  .ColAlignment(0) = 1
        .TextMatrix(0, Col����) = "ICD����":   .ColData(1) = 1:  .ColWidth(1) = 1250:  .ColAlignment(1) = 1
        .TextMatrix(0, Col���) = "�������":  .ColData(2) = 5:  .ColWidth(2) = 3200:  .ColAlignment(2) = 1
        .PrimaryCol = 1
        .LocateCol = 1
        For intRow = 1 To 4
            .TextMatrix(intRow, col����) = "����������" & intRow
            .TextMatrix(intRow, Col����) = ""
            .TextMatrix(intRow, Col���) = ""
            .RowData(intRow) = 0
        Next
        .Active = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEditable And Not mblnOK Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
        mstrSumTime = ""
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate����1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate����1_LostFocus()
    If Trim(MaskDate����1.Text <> "__:__:__") Then
        Call GetSumTime
    End If
End Sub

Private Sub MaskDate����2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate����2_LostFocus()
    If Trim(MaskDate����2.Text <> "__:__:__") Then
        Call GetSumTime
    End If
End Sub

Private Sub MaskDate����3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate����3_LostFocus()
    If Trim(MaskDate����3.Text <> "__:__:__") Then
        Call GetSumTime
    End If
End Sub

Private Sub GetSumTime()
    Dim strDate1 As String
    Dim strDate2 As String
    Dim strDate3 As String
    Dim intH As Integer
    Dim intM As Integer
    Dim intS As Integer
    Dim arrDate() As String
    Dim i As Integer
    If Not IsDate(MaskDate����1.Text) And Trim(MaskDate����1.Text <> "__:__:__") Then
        MsgBox "����ʱ��1��ʱ���ʽ����ȷ,���飡", vbInformation, gstrSysName
        If CanFocus(MaskDate����1) Then MaskDate����1.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(MaskDate����2.Text) And Trim(MaskDate����2.Text <> "__:__:__") Then
        MsgBox "����ʱ��2��ʱ���ʽ����ȷ,���飡", vbInformation, gstrSysName
        If CanFocus(MaskDate����2) Then MaskDate����2.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(MaskDate����3.Text) And Trim(MaskDate����3.Text <> "__:__:__") Then
        MsgBox "����ʱ��3��ʱ���ʽ����ȷ,���飡", vbInformation, gstrSysName
        If CanFocus(MaskDate����3) Then MaskDate����3.SetFocus
        Exit Sub
    End If
    
    strDate1 = IIf(Format(MaskDate����1.Text, "HH:MM:SS") = "00:00:00", "", Format(MaskDate����1.Text, "HH:MM:SS"))
    strDate2 = IIf(Format(MaskDate����2.Text, "HH:MM:SS") = "00:00:00", "", Format(MaskDate����2.Text, "HH:MM:SS"))
    strDate3 = IIf(Format(MaskDate����3.Text, "HH:MM:SS") = "00:00:00", "", Format(MaskDate����3.Text, "HH:MM:SS"))
    For i = 1 To 3
       If i = 1 Then
            If IsDate(strDate1) Then
                arrDate = Split(strDate1, ":")
                If UBound(arrDate) <> -1 Then
                    intH = intH + arrDate(0)
                    intM = intM + arrDate(1)
                    intS = intS + arrDate(2)
                End If
            End If
        ElseIf i = 2 Then
            If IsDate(strDate2) Then
                arrDate = Split(strDate2, ":")
                If UBound(arrDate) <> -1 Then
                    intH = intH + arrDate(0)
                    intM = intM + arrDate(1)
                    intS = intS + arrDate(2)
                End If
            End If
        ElseIf i = 3 Then
            If IsDate(strDate3) Then
                arrDate = Split(strDate3, ":")
                If UBound(arrDate) <> -1 Then
                    intH = intH + arrDate(0)
                    intM = intM + arrDate(1)
                    intS = intS + arrDate(2)
                End If
            End If
        End If
    Next
    If intS >= 60 Then
        intM = intM + intS \ 60
        intS = intS Mod 60
    End If
    If intM >= 60 Then
        intH = intH + intM \ 60
        intM = intM Mod 60
    End If
    txt�ܲ���ʱ��.Text = Trim(IIf(Len(Trim(intH)) = 1, 0 & intH, intH)) & ":" & Trim(IIf(Len(Trim(intM)) = 1, 0 & intM, intM)) & ":" & IIf(Len(Trim(intS)) = 1, 0 & intS, intS)
    mstrSumTime = txt�ܲ���ʱ��.Text
End Sub

Private Sub tab̥��_Click()
    Dim intRow As Integer
    Dim intIndex As Integer
    
    intIndex = tab̥��.SelectedItem.Index - 1
    For intRow = frḁ����Ϣ.LBound To frḁ����Ϣ.UBound
        frḁ����Ϣ(intRow).Visible = False
    Next
    frḁ����Ϣ(intIndex).Visible = True
    Set frḁ����Ϣ(intIndex).Container = tab̥��.Container
End Sub

Private Function GetControlPos(ByVal ctl As Control) As POINTAPI
    Dim p As POINTAPI
    
    p.X = ctl.Left / Screen.TwipsPerPixelX
    p.Y = ctl.Top / Screen.TwipsPerPixelY
    ClientToScreen ctl.Container.hwnd, p
    
    p.X = p.X * Screen.TwipsPerPixelX
    p.Y = p.Y * Screen.TwipsPerPixelY
    
    GetControlPos = p
End Function

Public Function SetBillInput(ByVal intIndex As Integer, ByVal LngRow As Long, rsInput As ADODB.Recordset) As Boolean
'���ܣ����������Ŀ������
    If InStr(rsInput!���� & "", "*") > 0 Then
        MsgBox "�Ǻű��벻����Ϊ��Ҫ���롣", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(rsInput!���� & "", 1) = "R" Then
        If MsgBox("��������ʹ��R������Ϊ��Ҫ���룬�Ƿ�ȷ�ϣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    With bill����������(intIndex)
        .RowData(LngRow) = rsInput!��ĿID
        .TextMatrix(LngRow, Col����) = rsInput!����
        .Text = rsInput!����
        .TextMatrix(LngRow, Col���) = rsInput!����
    End With
    SetBillInput = True
End Function

Private Function Validate() As Boolean
    Dim ctlError As Control
    Dim strMessage As String

    '��������Ϣ��������Ч��
    If Not Validate������Ϣ(ctlError, strMessage) Then
        If strMessage <> "" Then MsgBox strMessage, vbInformation, gstrSysName
        If CanFocus(ctlError) Then ctlError.SetFocus
        Exit Function
    End If
    
    '���Ӥ����Ϣ��������Ч��
    If Not Validate̥����Ϣ(ctlError, strMessage) Then
        If strMessage <> "" Then MsgBox strMessage, vbInformation, gstrSysName
        '����22275���� by lesfeng 2009-09-24
'        If CanFocus(ctlError) Then ctlError.SetFocus
        If CanFocus(ctlError) Then
            If Not ctlError.Enabled Then ctlError.Enabled = True
            ctlError.SetFocus
        End If
        Exit Function
    End If
    
    Validate = True
End Function

Public Sub Save������Ϣ()
    Dim intIndex As Integer, intRow As Integer, int̥�� As Integer
    Dim strFilter As String
    Dim arrTmp As Variant, strData As String, blnNew As Boolean
    Dim rsNew̥����� As ADODB.Recordset
    Dim rsNew��������� As ADODB.Recordset
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim arrValues() As Variant
    Dim strOld As String
    Dim arrSQL() As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    '��������Ϣ����
    arrTmp = Split("�������,̥��,̥��,����ʱ��1,����ʱ��2,����ʱ��3,�ܲ���ʱ��,�����Ѫ��,���Ʋ���֢,�����������", ",")
    
    For intIndex = LBound(arrTmp) To UBound(arrTmp)
        mrs������Ϣ.Filter = "��Ϣ��='" & arrTmp(intIndex) & "'"
        blnNew = mrs������Ϣ.EOF
        If Not blnNew Then
            strOld = mrs������Ϣ!��Ϣֵ & ""
        End If
        Select Case arrTmp(intIndex)
            Case "�������"
                strData = Trim(txt�������.Text)
            Case "̥��"
                strData = Trim(txt̥��.Text)
            Case "̥��"
                strData = Get����FromCombo(cmb̥��)
            Case "����ʱ��1"
                strData = MaskDate����1.Text
            Case "����ʱ��2"
                strData = MaskDate����2.Text
            Case "����ʱ��3"
                strData = MaskDate����3.Text
            Case "�ܲ���ʱ��"
                strData = txt�ܲ���ʱ��.Text
            Case "�����Ѫ��"
                strData = Trim(txt�����Ѫ��.Text)
            Case "���Ʋ���֢"
                strData = Get����FromCombo(cbo����֢)
            Case "�����������"
                strData = Get����FromCombo(cbo����)
        End Select
        If blnNew Then
            mrs������Ϣ.AddNew Array("��Ϣ��", "��Ϣ��ֵ", "����", "��¼����"), Array(arrTmp(intIndex), strData, 0, IIf(strOld & "" = strData, 0, 1))
        Else
            mrs������Ϣ.Update Array("��Ϣ��ֵ", "��¼����"), Array(strData, 1)
        End If
    Next
    
    Set rsNew̥����� = zlDatabase.CopyNewRec(mrs̥�����, True)
    Set rsNew��������� = zlDatabase.CopyNewRec(mrs���������, True)
    arrMainFileds = Array("����ID", "��ҳID", "����ʱ��", "̥������", "���䷽ʽ", "����̥λ", "�������", "����ȱ��", "Ӥ���Ա�", "Ӥ������", "Apgar����", "��¼����")
    arrSecdFileds = Array("����ID", "��ҳID", "̥������", "��ϴ���", "����id", "����", "������Ϣ", "��¼����")
    
    For intIndex = 1 To tab̥��.Tabs.Count
        With rsNew̥�����
'            .AddNew
            .AddNew arrMainFileds, Array(mlng����ID, mlng��ҳID, dtp����ʱ��(intIndex - 1).Value, intIndex, Get����FromCombo(cmb���䷽ʽ(intIndex - 1)), Get����FromCombo(cmb����̥λ(intIndex - 1)), Get����FromCombo(cmb�������(intIndex - 1)), _
                        Val(cmb����ȱ��(intIndex - 1).ListIndex), Get����FromCombo(cmbӤ���Ա�(intIndex - 1)), Format(Val(txtӤ������(intIndex - 1).Text) / 1000, "0.000"), Trim(txtApgar����(intIndex - 1).Text), 0)
'            !����ID = mlng����ID
'            !��ҳID = ��ҳID
'            !̥������ = intIndex
'            !���䷽ʽ = Get����FromCombo(cmb���䷽ʽ(intIndex - 1))
'            !����̥λ = Get����FromCombo(cmb����̥λ(intIndex - 1))
'            !������� = Get����FromCombo(cmb�������(intIndex - 1))
'            !����ȱ�� = Val(Get����FromCombo(cmb����ȱ��(intIndex - 1)))
'            !Ӥ���Ա� = Get����FromCombo(cmbӤ���Ա�(intIndex - 1))
'            !Ӥ������ = Trim(txtӤ������(intIndex - 1).Text)
'            !Apgar���� = Trim(txtApgar����(intIndex - 1).Text)
'            !��¼���� = 0
            .Update
        End With
        
        For intRow = 1 To 4
            strFilter = bill����������(intIndex - 1).TextMatrix(intRow, Col����)
            If strFilter <> "" Then
                With rsNew���������
                    .AddNew arrSecdFileds, Array(mlng����ID, mlng��ҳID, intIndex, intRow, bill����������(intIndex - 1).RowData(intRow), bill����������(intIndex - 1).TextMatrix(intRow, Col����), Replace(bill����������(intIndex - 1).TextMatrix(intRow, Col���), "'", "��"), 0)
                    .Update
                End With
            End If
        Next
    Next
    
    '��¼���Ƚ�
    mrs̥�����.Filter = "": rsNew̥�����.Filter = ""
    mrs���������.Filter = "": rsNew���������.Filter = ""
    If Rec.Compare(mrs̥�����, rsNew̥�����) Then
        If Not Rec.Compare(mrs���������, rsNew���������) Then
            Set mrs��������� = rsNew���������
            Call Rec.Update(mrs���������, "", "��¼����", 1) '�����Ϣ�Ѿ��ı�
        End If
    Else
        Set mrs̥����� = rsNew̥�����
        Call Rec.Update(mrs̥�����, "", "��¼����", 1) '�����Ϣ�Ѿ��ı�
        Set mrs��������� = rsNew���������
        Call Rec.Update(mrs���������, "", "��¼����", 1) '�����Ϣ�Ѿ��ı�
    End If
    If mbytModel = 2 Then
    '�������Ǽ�
        arrSQL = Array()
        Set grsDeliceryInfo = mrs������Ϣ
        Set grsBabyInfo = mrs̥�����
        Set grsBabyDiag = mrs���������
        Call PopDelicerySQL(arrSQL)
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get����FromCombo(ctlTemp As ComboBox, Optional ByVal blnID As Boolean = False) As String
    If blnID = False Then
        If ctlTemp.Text = "" Or ctlTemp.Enabled = False Then
            Get����FromCombo = ""
        Else
            Get����FromCombo = Mid(ctlTemp.Text, InStr(ctlTemp.Text, ".") + 1)
        End If
    Else
        Get����FromCombo = ctlTemp.ItemData(ctlTemp.ListIndex)
    End If
End Function

Private Function CanFocus(ctlError As Control) As Boolean
    If TypeName(ctlError) = "BillEdit" Then
        CanFocus = True
    Else
        CanFocus = ctlError.Enabled
    End If
End Function

Private Function Validate������Ϣ(ctlError As Control, strMessage As String) As Boolean
    
    '56928:������,2013-04-25,���˿�������Ժ֮ǰ�Ϳ�ʼ������
'    If CDate(dtp����ʱ��.Value) < CDate(txt��Ժ����) Then
'        Set ctlError = dtp����ʱ��
'        strMessage = "����ʱ�䲻�ܱ���Ժ���ڻ���!"
'        Exit Function
'    End If
    '����2005-8-10�е������͵����Ժ�����
    'if dtp����ʱ��.value-1>txt��Ժ���� then
    
    If Not IsNumeric(txt�������) Then
        Set ctlError = txt�������
        strMessage = "���������д����ȷ,����!"
        Exit Function
    End If
    
    If Not IsNumeric(txt̥��) Then
        Set ctlError = txt̥��
        strMessage = "̥����д����ȷ,����!"
        Exit Function
    End If
    
    If Not IsNumeric(txt�����Ѫ��) Then
        Set ctlError = txt�����Ѫ��
        strMessage = "�����Ѫ����д����ȷ,����!"
        Exit Function
    End If
    
    If Not IsDate(MaskDate����1.Text) And MaskDate����1.Text <> "__:__:__" Then
        Set ctlError = MaskDate����1
        strMessage = "����ʱ��1��ʱ���ʽ����ȷ,���飡"
        Exit Function
    End If
    
    If Not IsDate(MaskDate����2.Text) And MaskDate����2.Text <> "__:__:__" Then
        Set ctlError = MaskDate����2
        strMessage = "����ʱ��2��ʱ���ʽ����ȷ,���飡"
        Exit Function
    End If
    
    If Not IsDate(MaskDate����3.Text) And MaskDate����3.Text <> "__:__:__" Then
        Set ctlError = MaskDate����3
        strMessage = "����ʱ��3��ʱ���ʽ����ȷ,���飡"
        Exit Function
    End If
    
    Validate������Ϣ = True
End Function

Private Function Validate̥����Ϣ(ctlError As Control, strMessage As String) As Boolean
    Dim intIndex As Integer
    Dim lng����ID As Long, j As Long
    
    
    For intIndex = 0 To tab̥��.Tabs.Count - 1
        If Val(Left(cmb�������(intIndex).Text, 1)) = 1 Then
            'ֻ�л���Ż�������������
            '����:22275
            If Not IsNumeric(txtӤ������(intIndex)) Then
                tab̥��.Tabs(intIndex + 1).Selected = True
                'frḁ����Ϣ(intIndex).Visible = True
                Set ctlError = txtӤ������(intIndex)
                strMessage = "Ӥ��������д����ȷ,����!"
                Exit Function
            End If
            
            If Not IsNumeric(txtApgar����(intIndex)) Then
                tab̥��.Tabs(intIndex + 1).Selected = True
                'frḁ����Ϣ(intIndex).Visible = True
                Set ctlError = txtApgar����(intIndex)
                strMessage = "Apgar������д����ȷ,����!"
                Exit Function
            End If
        End If
        '����:22275
        '���ַ�ʽ�Ƿ���ȷ
        If cmb���䷽ʽ(0).Text <> cmb���䷽ʽ(intIndex).Text Then
            Set ctlError = cmb���䷽ʽ(intIndex)
            tab̥��.Tabs(intIndex + 1).Selected = True
            If MsgBox("��" & intIndex + 1 & "̥�ķ��䷽ʽ���1̥���䷽ʽ��һ��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        If mbytModel = 1 And intIndex = 0 Then
            If CDate(dtp����ʱ��(intIndex).Value) > CDate(txt��Ժ����) Then
                If Not frḁ����Ϣ(intIndex).Visible Then
                    frḁ����Ϣ(intIndex).Visible = True
                    Set ctlError = dtp����ʱ��(intIndex)
                    strMessage = "����ʱ�䲻�ܱȳ�Ժ���ڻ���!"
                Else
                    Set ctlError = dtp����ʱ��(intIndex)
                    strMessage = "����ʱ�䲻�ܱȳ�Ժ���ڻ���!"
                End If
                Exit Function
            End If
        Else
            If intIndex > 0 Then
                If CDate(dtp����ʱ��(intIndex).Value) < CDate(dtp����ʱ��(intIndex - 1).Value) Then
                    If MsgBox("�ڡ�" & intIndex + 1 & "����̥���ķ���ʱ��ȵڡ�" & intIndex & "����̥���ķ���ʱ��С��ȷ��������", vbYesNo, gstrSysName) = vbNo Then
                        If Not frḁ����Ϣ(intIndex).Visible Then
                            frḁ����Ϣ(intIndex).Visible = True
                            Set ctlError = dtp����ʱ��(intIndex)
                        Else
                            Set ctlError = dtp����ʱ��(intIndex)
                        End If
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Validate̥����Ϣ = True
End Function

Private Sub LoadData()
    Dim intIndex As Integer, intRow As Integer
    Dim strData As String
    
    mrs������Ϣ.Filter = "����=0": mrs������Ϣ.Sort = "��Ϣ��"
    If mrs������Ϣ.RecordCount = 0 Then Exit Sub
    
    mintFlag = 1
    Do While Not mrs������Ϣ.EOF
        strData = mrs������Ϣ!��Ϣ��ֵ & ""
        Select Case mrs������Ϣ!��Ϣ�� & ""
            Case "�������"
                txt�������.Text = IIf(Val(strData) = 0, "", strData)
            Case "̥��"
                txt̥��.Text = IIf(Val(strData) = 0, "", strData)
            Case "̥��"
                AddComboItem cmb̥��, CStr(strData), 2
            Case "����ʱ��1"
                MaskDate����1.Text = Format(strData, "HH:MM:SS")
            Case "����ʱ��2"
                MaskDate����2.Text = Format(strData, "HH:MM:SS")
            Case "����ʱ��3"
                MaskDate����3.Text = Format(strData, "HH:MM:SS")
            Case "�ܲ���ʱ��"
                txt�ܲ���ʱ��.Text = Format(strData, "HH:MM:SS")
            Case "�����Ѫ��"
                txt�����Ѫ��.Text = IIf(Val(strData) = 0, "", strData)
            Case "���Ʋ���֢"
                AddComboItem cbo����֢, CStr(strData)
            Case "�����������"
                AddComboItem cbo����, CStr(strData)
        End Select
        mrs������Ϣ.MoveNext
    Loop
    cmb̥��_Click
    For intIndex = 1 To tab̥��.Tabs.Count
        With mrs̥�����
            .Filter = "̥������ = " & intIndex
            If Not .EOF Then
                If IsNull(!����ʱ��) Then
                    mrs������Ϣ.Filter = "��Ϣ��='����ʱ��'"
                    If Not mrs������Ϣ.EOF Then dtp����ʱ��(intIndex - 1).Value = CDate("" & mrs������Ϣ!��Ϣֵ)
                Else
                    dtp����ʱ��(intIndex - 1).Value = CDate(!����ʱ��)
                End If
                AddComboItem cmb���䷽ʽ(intIndex - 1), IIf(!���䷽ʽ = "����", "��������", !���䷽ʽ)
                AddComboItem cmb����̥λ(intIndex - 1), !����̥λ
                AddComboItem cmb�������(intIndex - 1), !�������
                AddComboItem cmbӤ���Ա�(intIndex - 1), !Ӥ���Ա�
                cmb����ȱ��(intIndex - 1).ListIndex = Val(!����ȱ�� & "")
                txtӤ������(intIndex - 1) = IIf(IsNull(!Ӥ������), "", Val(Nvl(!Ӥ������, "")) * 1000)
                txtApgar����(intIndex - 1) = IIf(IsNull(!Apgar����), "", !Apgar����)
            End If
        End With
        
        With mrs���������
            For intRow = 1 To 4
                .Filter = "̥������ = " & intIndex & " and ��ϴ��� = " & intRow
                If Not .EOF Then
                    bill����������(intIndex - 1).TextMatrix(intRow, Col����) = !����
                    bill����������(intIndex - 1).TextMatrix(intRow, Col���) = !������Ϣ
                    bill����������(intIndex - 1).RowData(intRow) = !����id
                End If
            Next
        End With
    Next
    
    mrs̥�����.Filter = ""
    mrs���������.Filter = ""
End Sub

Private Function AddComboItem(cmbTemp As Control, strItem As String, Optional ByVal cmbType As Integer = 1, Optional ByVal cmbData As Long) As Boolean
    '����cmbType  = 1ʱ��ʾ�����������ִ�ͷ��
    '             = 2ʱ��ʾȫ������
    Dim varTemp As Variant
    '�������б����
    If IsNull(strItem) Or Trim(strItem) = "" Then Exit Function
    For varTemp = 0 To cmbTemp.ListCount - 1
        If cmbType = 1 Then
            If strItem = Mid(cmbTemp.List(varTemp), InStr(cmbTemp.List(varTemp), ".") + 1) Then
                cmbTemp.ListIndex = varTemp
                Exit Function
            End If
        ElseIf cmbType = 2 Then
            If strItem = cmbTemp.List(varTemp) Then
                cmbTemp.ListIndex = varTemp
                Exit Function
            End If
        Else
            If cmbData = cmbTemp.ItemData(varTemp) Then
                cmbTemp.ListIndex = varTemp
                Exit Function
            End If
        End If
    Next
    
    If cmbType = 1 Then
        If cmbTemp.ListCount > 0 Then
            varTemp = cmbTemp.ListCount + 1
        Else
            varTemp = 1
        End If
        cmbTemp.AddItem IIf(Not mblnEditable, "", varTemp & ".") & strItem
        cmbTemp.ListIndex = cmbTemp.NewIndex
    ElseIf cmbType = 2 Then
        cmbTemp.AddItem strItem
        cmbTemp.ListIndex = cmbTemp.NewIndex
    End If
End Function

Private Sub tab̥��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtApgar����_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtApgar����(Index)
End Sub

Private Sub txtApgar����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtApgar����_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�����Ѫ��_GotFocus()
    zlControl.TxtSelAll txt�������
End Sub

Private Sub txt�����Ѫ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�����Ѫ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�������_GotFocus()
    zlControl.TxtSelAll txt�������
End Sub

Private Sub txt�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt̥��_GotFocus()
    zlControl.TxtSelAll txt̥��
End Sub

Private Sub txt̥��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt̥��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtӤ������_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtӤ������_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Function GetInDate() As Date
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim DatRet As Date
    
    On Error GoTo errH
    strSql = "Select ��Ժ���� From ������ҳ T Where t.����id = [1] And t.��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        DatRet = CDate(Nvl(rsTmp!��Ժ����, "0:00:00"))
    Else
        DatRet = CDate("0:00:00")
    End If
    GetInDate = DatRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt�ܲ���ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�ܲ���ʱ��_Validate(Cancel As Boolean)
    If Format(mstrSumTime, "HH:MM:SS") <> "" Then
        If Format(txt�ܲ���ʱ��.Text, "HH:MM:SS") <> "" Then
            If Format(txt�ܲ���ʱ��.Text, "HH:MM:SS") <> Format(mstrSumTime, "HH:MM:SS") Then
                If MsgBox("�ܲ���ʱ��Ϊ" & Format(txt�ܲ���ʱ��.Text, "HH:MM:SS") & "�����ڲ���ʱ��1+����ʱ��2+����ʱ��3֮��" & Format(mstrSumTime, "HH:MM:SS") & ",�Ƿ����?", vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                    Cancel = False
                Else
                    Cancel = True
                End If
            Else
                Cancel = False
            End If
        Else
            If MsgBox("�ܲ���ʱ��Ϊ��,����ʱ��1+����ʱ��2+����ʱ��3֮��Ϊ" & Format(mstrSumTime, "HH:MM:SS") & ",�Ƿ����?", vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                Cancel = False
            Else
                Cancel = True
            End If
        End If
    End If
End Sub
