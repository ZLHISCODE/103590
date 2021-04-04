VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeliceryInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分娩情况录入"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdGetDeliceryInfo 
      Caption         =   "提取新生儿信息"
      Height          =   300
      Left            =   120
      TabIndex        =   171
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   3
      Left            =   240
      TabIndex        =   93
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   3
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   111
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   3
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   105
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   3
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   3
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   3
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   3
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   3
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   107
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
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
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
         Height          =   180
         Index           =   3
         Left            =   6720
         TabIndex        =   104
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   3
         Left            =   3480
         TabIndex        =   102
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   100
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   3
         Left            =   6720
         TabIndex        =   98
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   3
         Left            =   3480
         TabIndex        =   96
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   94
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   2
      Left            =   240
      TabIndex        =   74
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   2
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   92
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   2
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   86
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   2
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   2
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   2
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   88
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
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
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
         Height          =   180
         Index           =   2
         Left            =   6720
         TabIndex        =   85
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   2
         Left            =   3480
         TabIndex        =   83
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   81
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   2
         Left            =   6720
         TabIndex        =   79
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   2
         Left            =   3480
         TabIndex        =   77
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   75
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   6
      Left            =   240
      TabIndex        =   150
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   6
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   168
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   6
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   162
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   6
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   160
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   6
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   158
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   6
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   6
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   154
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   6
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   164
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
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
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
         Height          =   180
         Index           =   6
         Left            =   6720
         TabIndex        =   161
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   159
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   157
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   6
         Left            =   6720
         TabIndex        =   155
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   153
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   151
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   5
      Left            =   240
      TabIndex        =   131
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   5
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   149
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   5
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   143
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   5
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   141
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   5
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   139
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   5
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   5
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   5
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   145
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
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
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
         Height          =   180
         Index           =   5
         Left            =   6720
         TabIndex        =   142
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   5
         Left            =   3480
         TabIndex        =   140
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   138
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   5
         Left            =   6720
         TabIndex        =   136
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   5
         Left            =   3480
         TabIndex        =   134
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   132
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   4
      Left            =   240
      TabIndex        =   112
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   4
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   130
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   4
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   124
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   4
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   4
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   4
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   4
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   4
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   360
         Width           =   1455
      End
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   126
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
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
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
         Height          =   180
         Index           =   4
         Left            =   6720
         TabIndex        =   123
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   4
         Left            =   3480
         TabIndex        =   121
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   119
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   4
         Left            =   6720
         TabIndex        =   117
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   4
         Left            =   3480
         TabIndex        =   115
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   113
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   1
      Left            =   240
      TabIndex        =   55
      Top             =   3840
      Width           =   9255
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   1
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   1
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   67
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   1
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   73
         Top             =   3435
         Width           =   975
      End
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   56
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   58
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   60
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   62
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   64
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
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
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   72
         Top             =   3480
         Width           =   810
      End
   End
   Begin VB.Frame fra胎儿信息 
      Caption         =   "胎儿信息"
      Height          =   3855
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   3840
      Width           =   9255
      Begin VB.TextBox txtApgar评分 
         Height          =   330
         Index           =   0
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   54
         Top             =   3435
         Width           =   975
      End
      Begin VB.TextBox txt婴儿体重 
         Height          =   330
         Index           =   0
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   48
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cmb婴儿性别 
         Height          =   300
         Index           =   0
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生缺陷 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   780
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩情况 
         Height          =   300
         Index           =   0
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb出生胎位 
         Height          =   300
         Index           =   0
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb分娩方式 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp分娩时间 
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
            Name            =   "宋体"
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
      Begin ZL9BillEdit.BillEdit bill新生儿疾病 
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
      Begin VB.Label lbl分娩时间 
         AutoSize        =   -1  'True
         Caption         =   "分娩时间"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblApgar评分 
         AutoSize        =   -1  'True
         Caption         =   "Apgar评分"
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
      Begin VB.Label lbl婴儿体重 
         AutoSize        =   -1  'True
         Caption         =   "婴儿体重"
         Height          =   180
         Index           =   0
         Left            =   6720
         TabIndex        =   47
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl婴儿性别 
         AutoSize        =   -1  'True
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   0
         Left            =   3480
         TabIndex        =   45
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl出生缺陷 
         AutoSize        =   -1  'True
         Caption         =   "出生缺陷"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl分娩情况 
         AutoSize        =   -1  'True
         Caption         =   "分娩情况"
         Height          =   180
         Index           =   0
         Left            =   6720
         TabIndex        =   41
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl出生胎位 
         AutoSize        =   -1  'True
         Caption         =   "出生胎位"
         Height          =   180
         Index           =   0
         Left            =   3480
         TabIndex        =   39
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl分娩方式 
         AutoSize        =   -1  'True
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   8520
      TabIndex        =   170
      Top             =   7920
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   7320
      TabIndex        =   169
      Top             =   7920
      Width           =   1100
   End
   Begin VB.Frame fra分娩信息 
      Caption         =   "分娩信息"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Tag             =   "##:##:##"
      Top             =   1320
      Width           =   9495
      Begin VB.TextBox txt总产程时间 
         Height          =   270
         Left            =   1200
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskDate产程1 
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
      Begin VB.TextBox txt产检次数 
         Height          =   330
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   15
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txt胎次 
         Height          =   330
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "1"
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txt产后出血量 
         Height          =   330
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmb胎数 
         Height          =   300
         ItemData        =   "frmDeliceryInfo.frx":000C
         Left            =   8055
         List            =   "frmDeliceryInfo.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   315
         Width           =   1215
      End
      Begin VB.ComboBox cbo并发症 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1515
         Width           =   4755
      End
      Begin VB.ComboBox cbo会阴 
         Height          =   300
         Left            =   8055
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1095
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskDate产程2 
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
      Begin MSMask.MaskEdBox MaskDate产程3 
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
         Caption         =   "会阴Ⅲ度裂伤"
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   31
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "产科并发症"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   33
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lbl总产程时间 
         AutoSize        =   -1  'True
         Caption         =   "总产程时间"
         Height          =   180
         Left            =   45
         TabIndex        =   26
         Top             =   1155
         Width           =   900
      End
      Begin VB.Label lbl产程时间1 
         AutoSize        =   -1  'True
         Caption         =   "产程时间1"
         Height          =   180
         Left            =   135
         TabIndex        =   20
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl产程时间2 
         AutoSize        =   -1  'True
         Caption         =   "产程时间2"
         Height          =   180
         Left            =   3810
         TabIndex        =   22
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl产后出血量 
         AutoSize        =   -1  'True
         Caption         =   "产后出血量"
         Height          =   180
         Left            =   3720
         TabIndex        =   28
         Top             =   1155
         Width           =   900
      End
      Begin VB.Label lbl产检次数 
         AutoSize        =   -1  'True
         Caption         =   "产检次数"
         Height          =   180
         Left            =   225
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl胎次 
         AutoSize        =   -1  'True
         Caption         =   "胎次"
         Height          =   180
         Left            =   4260
         TabIndex        =   16
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl胎数 
         AutoSize        =   -1  'True
         Caption         =   "胎数"
         Height          =   180
         Left            =   7425
         TabIndex        =   18
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl单位 
         AutoSize        =   -1  'True
         Caption         =   "Ml"
         Height          =   180
         Left            =   6060
         TabIndex        =   30
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label lbl产程时间3 
         AutoSize        =   -1  'True
         Caption         =   "产程时间3"
         Height          =   180
         Left            =   6975
         TabIndex        =   24
         Top             =   765
         Width           =   810
      End
   End
   Begin VB.Frame fra基本信息 
      Caption         =   "基本信息"
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txt入院日期 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txt出院日期 
         Enabled         =   0   'False
         Height          =   330
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txt出院主要诊断 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
      Begin VB.TextBox txt住院号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
      Begin VB.TextBox txt病案号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
      Begin VB.Label lbl入院日期 
         AutoSize        =   -1  'True
         Caption         =   "入院日期"
         Height          =   180
         Left            =   4830
         TabIndex        =   5
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl病案号 
         AutoSize        =   -1  'True
         Caption         =   "病案号"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   2790
         TabIndex        =   3
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   300
         TabIndex        =   9
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lbl出院日期 
         AutoSize        =   -1  'True
         Caption         =   "出院日期"
         Height          =   180
         Left            =   7200
         TabIndex        =   7
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl出院主要诊断 
         AutoSize        =   -1  'True
         Caption         =   "出院主要诊断"
         Height          =   180
         Left            =   2790
         TabIndex        =   11
         Top             =   690
         Width           =   1080
      End
   End
   Begin MSComctlLib.TabStrip tab胎数 
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
            Caption         =   "第1个胎儿"
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
'入口参数
Private mlng病人ID As Long                '病人ID
Private mlng主页ID As Long                '主页ID
Private mblnEditable As Boolean           '是否可以编辑
Private mlng疾病ID As Long                '病人出院主要诊断(西医)的疾病ID
Private mrs分娩信息 As ADODB.Recordset    '病人分娩信息
Private mrs胎儿情况 As ADODB.Recordset    '新生儿基本信息
Private mrs新生儿诊断 As ADODB.Recordset  '新生儿诊断信息
Private mbytModel As Byte                 'mbytModel=1 病案系统;=2新生儿登记
Private mstrSumTime As String

Private Enum BabyDiag
    col类型 = 0
    Col编码 = 1
    Col诊断 = 2
End Enum

Private mblnOK As Boolean
Private mintFlag As Integer

Public Function EditDelivery(ByRef frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng疾病ID As Long, ByVal blnEditable As Boolean, _
                        ByRef rs分娩信息 As ADODB.Recordset, ByRef rs胎儿信息 As ADODB.Recordset, ByRef rs新生儿疾病 As ADODB.Recordset, _
                        Optional ByRef blnOK As Boolean, Optional ByVal bytModel As Byte = 1) As Boolean
    Dim ctlTemp As Control

    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng疾病ID = lng疾病ID
    mblnEditable = blnEditable
    rs分娩信息.Filter = "": rs胎儿信息.Filter = "": rs新生儿疾病.Filter = ""
    Set mrs分娩信息 = zlDatabase.CopyNewRec(rs分娩信息)
    Set mrs胎儿情况 = zlDatabase.CopyNewRec(rs胎儿信息)
    Set mrs新生儿诊断 = zlDatabase.CopyNewRec(rs新生儿疾病)
    mbytModel = bytModel
    
    mblnOK = False
    Me.Show 1, frmParent
    rs分娩信息.Filter = "": rs胎儿信息.Filter = "": rs新生儿疾病.Filter = ""
    Set rs分娩信息 = mrs分娩信息
    Set rs胎儿信息 = mrs胎儿情况
    Set rs新生儿疾病 = mrs新生儿诊断
    blnOK = mblnOK
    EditDelivery = True
    
End Function

Private Sub bill新生儿疾病_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    Dim intIndex As Integer
    Dim intCol As Integer
    
    intIndex = tab胎数.SelectedItem.Index
    With bill新生儿疾病(intIndex - 1)
        For intCol = Col编码 To Col诊断
            .TextMatrix(.Row, intCol) = ""
        Next
        Cancel = True
    End With
End Sub

Private Sub bill新生儿疾病_CommandClick(Index As Integer)
    Dim intIndex As Integer
    Dim strSql As String, str性别 As String
    Dim rsTmp As ADODB.Recordset
    
    intIndex = tab胎数.SelectedItem.Index - 1
    str性别 = cmb婴儿性别(intIndex).Text
    If str性别 Like "*男*" Then
        str性别 = "男"
    ElseIf str性别 Like "*女*" Then
        str性别 = "女"
    Else
        str性别 = ""
    End If
    
    With bill新生儿疾病(intIndex)
        Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "D", gclsPros.出院科室ID, str性别, False)
        If Not rsTmp Is Nothing Then
            SetBillInput intIndex, bill新生儿疾病(intIndex).Row, rsTmp
        End If
'        .SetFocus
    End With
End Sub

Private Sub bill新生儿疾病_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim intIndex As Integer
    Dim strFilter As String
    Dim arrFilter As Variant, lngCount As Long
    Dim blnOK As Boolean
    Dim vPoint As POINTAPI
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    intIndex = tab胎数.SelectedItem.Index - 1
    With bill新生儿疾病(intIndex).MsfObj
        If .Col = Col编码 Then
            strFilter = UCase(Replace(Trim(bill新生儿疾病(intIndex).Text), "'", "''"))
            strFilter = Replace(strFilter, "。", ".")

            If strFilter = "%" Or strFilter = "_" Then
                strSql = "select A.ID,A.ID 项目ID,A.编码,附码,A.名称,A.简码,A.说明,A.性别限制,A.疗效限制 as 提醒疗效 " & _
                          " From 疾病编码目录 A  " & _
                          " Where A.类别='D' and Rownum<1000  and (a.撤档时间 is null or a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd'))"
            Else
                If InStr(strFilter, " ") = 0 Then
                    strSql = "select A.ID,A.ID 项目ID,A.编码,附码,A.名称,A.简码,A.说明,A.性别限制,A.疗效限制 as 提醒疗效 " & _
                              " From 疾病编码目录 A  " & _
                              " Where A.类别='D' and Rownum<1000     and (a.撤档时间 is null or a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd')) " & _
                              "         and (upper(A.编码) like '" & strFilter & "%' or upper(A.名称) like '%" & strFilter & "%' or upper(A.简码) like '%" & strFilter & "%')  "
                Else
                    arrFilter = Split(strFilter, " ")
                    strSql = ""
                    For lngCount = LBound(arrFilter) To UBound(arrFilter)
                        If Trim(arrFilter(lngCount)) <> "" Then
                            strSql = strSql & " and upper(A.名称) like '%" & Trim(arrFilter(lngCount)) & "%'"
                        End If
                    Next
                    strSql = Mid(strSql, 5) '去掉第一个and
                    If Trim(strSql) = "" Then
                        strSql = " upper(A.编码) like '" & strSql & "%'"
                    Else
                        strSql = "(" & strSql & ") or upper(A.编码) like '" & strFilter & "%'"
                    End If
                    strSql = "select A.ID, A.ID 项目ID,A.编码,附码,A.名称,A.简码,A.说明,A.性别限制,A.疗效限制 as 提醒疗效 " & _
                              " From 疾病编码目录 A  " & _
                              " Where A.类别='D' and Rownum<1000     and (a.撤档时间 is null or a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd')) and (" & strSql & ")"
                End If
            End If
            
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)

            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "疾病编码", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
            If blnCancel Or rsTmp Is Nothing Then
                Cancel = True
                Call Beep
                bill新生儿疾病(intIndex).TxtSetFocus
            Else
                blnOK = SetBillInput(intIndex, bill新生儿疾病(intIndex).Row, rsTmp)
                If blnOK = False Then
                    Cancel = True
                    bill新生儿疾病(intIndex).TxtSetFocus
                Else
                    Cancel = False
                    bill新生儿疾病(intIndex).TxtVisible = False
                End If
            End If
        End If
    End With
End Sub

Private Sub cbo并发症_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyCode = vbKeyDelete Then cbo并发症.ListIndex = -1
End Sub

 
Private Sub cbo会阴_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyCode = vbKeyDelete Then cbo会阴.ListIndex = -1
End Sub

Private Sub cmb出生缺陷_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb出生胎位_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb分娩方式_Click(Index As Integer)
    Dim intCount As Integer
    Dim i As Integer
    Dim str分娩 As String
    
    '问题22275完善 by lesfeng 2009-09-24
    intCount = Val(cmb胎数.Text) - 1
    If intCount < 1 Then Exit Sub
    If mintFlag = 1 Then Exit Sub
        
    If Index = 0 And mintFlag = 0 Then
        str分娩 = cmb分娩方式(0).Text
        For i = 1 To intCount
            cmb分娩方式(i).Text = str分娩
        Next
    End If
    
    If Not tab胎数.Tabs(1).Selected Then
        If Index > 0 Then mintFlag = 1
    End If
End Sub

Private Sub cmb分娩方式_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb分娩情况_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb胎数_Click()
    Dim intIndex As Integer
    Dim intRow As Integer
    Dim str分娩 As String
    
    str分娩 = cmb分娩方式(0).Text
    For intIndex = tab胎数.Tabs.Count To cmb胎数.Text - 1
        initFetus intIndex
        '问题22275完善 by lesfeng 2009-09-24
        cmb分娩方式(intIndex).Text = str分娩
    Next
    
    tab胎数.Tabs.Clear
    For intIndex = 1 To cmb胎数.Text
        tab胎数.Tabs.Add intIndex, , "第" & intIndex & "个胎儿"
    Next
    
    For intIndex = fra胎儿信息.LBound To fra胎儿信息.UBound
        fra胎儿信息(intIndex).Visible = False
    Next
    
    tab胎数.Tabs(1).Selected = True
    fra胎儿信息(0).Visible = True
    
End Sub

Private Sub cmb胎数_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb婴儿性别_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    Dim int胎数 As Integer
    Dim intIndex As Integer
    If mlng病人ID <> 0 And mlng主页ID <> 0 Then
        strSql = "Select 序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,身长,体重,血型,出生时间,死亡时间,备注说明 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] Order by 序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "新生儿信息", mlng病人ID, mlng主页ID)
        If rsTmp.EOF Then
            MsgBox "该病人没有新生儿记录！", vbInformation, gstrSysName
        Else
            If Not mrs分娩信息.RecordCount < 0 Then
                If MsgBox("该病人存在新生儿记录，是否覆盖当前分娩信息？", vbYesNo, Me.Caption) = vbYes Then
                    int胎数 = rsTmp.RecordCount
                    cmb胎数.ListIndex = Cbo.FindIndex(cmb胎数, CStr(int胎数))
                    cmb胎数_Click
                    For intIndex = 1 To tab胎数.Tabs.Count
                        With rsTmp
                            .Filter = "序号 = " & intIndex
                            If Not .EOF Then
                                If IsNull(!出生时间) Then
                                    dtp分娩时间(intIndex - 1).Value = zlDatabase.Currentdate
                                Else
                                    dtp分娩时间(intIndex - 1).Value = CDate(!出生时间)
                                End If
                                AddComboItem cmb分娩方式(intIndex - 1), IIf(!分娩方式 = "其他", "其他分娩", !分娩方式)
                                AddComboItem cmb婴儿性别(intIndex - 1), !婴儿性别
                                txt婴儿体重(intIndex - 1) = IIf(IsNull(!体重), "", Val(Nvl(!体重, "")))
                            End If
                        End With
                    Next
                End If
            Else
                int胎数 = rsTmp.RecordCount
                cmb胎数.ListIndex = Cbo.FindIndex(cmb胎数, CStr(int胎数))
                cmb胎数_Click
                For intIndex = 1 To tab胎数.Tabs.Count
                    With rsTmp
                        .Filter = "序号 = " & intIndex
                        If Not .EOF Then
                            If IsNull(!出生时间) Then
                                dtp分娩时间(intIndex - 1).Value = zlDatabase.Currentdate
                            Else
                                dtp分娩时间(intIndex - 1).Value = CDate(!出生时间)
                            End If
                            AddComboItem cmb分娩方式(intIndex - 1), IIf(!分娩方式 = "其他", "其他分娩", !分娩方式)
                            AddComboItem cmb婴儿性别(intIndex - 1), !婴儿性别
                            txt婴儿体重(intIndex - 1) = IIf(IsNull(!体重), "", Val(Nvl(!体重, "")))
                        End If
                    End With
                Next
            End If
        End If
    End If
End Sub

Private Sub cmdOk_Click()
    
    If Validate Then
        Save分娩信息
        mblnOK = True
        Unload Me
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp分娩时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    
    '初始化信息
    Call InitCombox
    
    If mbytModel = 1 Then
        mrs分娩信息.Filter = "类型=1": mrs分娩信息.Sort = "信息名"
        
        Do While Not mrs分娩信息.EOF
            Select Case mrs分娩信息!信息名
                Case "病案号"
                    txt病案号.Text = mrs分娩信息!信息值
                Case "住院号"
                    txt住院号.Text = mrs分娩信息!信息值
                Case "姓名"
                    txt姓名.Text = mrs分娩信息!信息值
                Case "入院日期"
                    txt入院日期.Text = mrs分娩信息!信息值
                Case "出院日期"
                    txt出院日期.Text = mrs分娩信息!信息值
                Case "主要诊断"
                    txt出院主要诊断.Text = mrs分娩信息!信息值
            End Select
            mrs分娩信息.MoveNext
        Loop
        If mrs胎儿情况.RecordCount <= 0 Then
             dtp分娩时间(0).Value = CDate(txt入院日期.Text)
        End If
    Else
        Me.Height = Me.Height - 1200
        fra基本信息.Visible = False
        fra分娩信息.Move 120, 120
        tab胎数.Move 120, fra分娩信息.Top + fra分娩信息.Height + 120
        For intIndex = fra胎儿信息.LBound To fra胎儿信息.UBound
            fra胎儿信息(intIndex).Move tab胎数.Left + 120, tab胎数.Top + 360
        Next
        cmdGetDeliceryInfo.Move cmdGetDeliceryInfo.Left, tab胎数.Top + tab胎数.Height + 120
        cmdOK.Move cmdOK.Left, Me.ScaleHeight - cmdOK.Height - 90
        cmdCancel.Move cmdCancel.Left, Me.ScaleHeight - cmdOK.Height - 90
        If mrs胎儿情况.RecordCount <= 0 Then
            dtp分娩时间(0).Value = GetInDate
        End If
    End If
    
    '调整控件字体大小
    For Each ctlTemp In Me.Controls
        If InStr("DTPickerTabStripBillEdit", TypeName(ctlTemp)) = 0 Then
            ctlTemp.FontSize = 10.5
        Else
            ctlTemp.Font.Size = 10.5
        End If
    Next
    
    '初始化胎儿数
    For intRow = 1 To 7
        arrFetus(intRow - 1) = intRow
    Next
    LoadComboFromArray arrFetus, cmb胎数
    
    '初始化表格
    initFetus 0
    For intIndex = fra胎儿信息.LBound To fra胎儿信息.UBound
        fra胎儿信息(intIndex).Visible = False
    Next
    fra胎儿信息(0).Visible = True
    
    LoadData
        
    If Not mblnEditable Then
        cmdOK.Visible = False
        cmdCancel.Caption = "关闭(&C)"
        
        For Each ctlTemp In Me.Controls
            If InStr("Label,TabStrip,CommandButton", TypeName(ctlTemp)) = 0 Then
                ctlTemp.Enabled = False
            End If
        Next
    End If
End Sub

Private Function LoadComboFromArray(ByVal varArray As Variant, cmbTemp As Variant) As Boolean
'本函数的功能是数组中读出列表值装到下拉框中
    Dim cmbArray As Variant
    Dim intArray As Long
    Dim intCount As Long
    
    On Error GoTo errHandle
    
    If IsArray(cmbTemp) Then
        cmbArray = cmbTemp
    Else
        '强行组成一个数组
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
    '功能:初始化combox控件信息,不包含胎儿信息
    '编制:刘兴洪
    '日期:2009-02-06 10:08:28
    '-----------------------------------------------------------------------------------------------------------
    '15126
    LoadComboFromArray Array("1.子痫", "2.产后流血", "3.产褥热"), cbo并发症
    LoadComboFromArray Array("1.有", "2.无"), cbo会阴
    '默认为空
    cbo会阴.ListIndex = -1: cbo并发症.ListIndex = -1
        
End Sub
Private Sub initFetus(i_intIndex As Integer)
    Dim intRow As Integer
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    '15126
    LoadComboFromArray Array("1.活产", "2.死产", "3.死胎", "4.新生儿死亡"), cmb分娩情况(i_intIndex)
    LoadComboFromArray Array("1.左枕前", "2.右枕前", "3.左枕后", "4.右枕后", "5.左枕横", "6.右枕横", "7.臀位", "8.肩先露", "9.足先露", "10.膝先露"), cmb出生胎位(i_intIndex)
    LoadComboFromArray Array("1.无", "2.有"), cmb出生缺陷(i_intIndex)
    
    strSql = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 分娩方式 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    cmb分娩方式(i_intIndex).Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cmb分娩方式(i_intIndex).AddItem i & "." & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cmb分娩方式(i_intIndex).ListIndex = cmb分娩方式(i_intIndex).NewIndex
                cmb分娩方式(i_intIndex).ItemData(cmb分娩方式(i_intIndex).NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
    Else
        '15126
        LoadComboFromArray Array("1.正常分娩", "2.其他分娩", "3.剖宫产", "4.早产", "5.产钳", "6.臀抽", "7.臀助"), cmb分娩方式(i_intIndex) '刘兴洪:臀抽改为臀助,问题:21778,by lesfeng 2009-9-23 由于臀抽是老方法、臀助是新方法，不是替代，是增加
    End If
    
    strSql = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    cmb婴儿性别(i_intIndex).Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cmb婴儿性别(i_intIndex).AddItem i & "." & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cmb婴儿性别(i_intIndex).ListIndex = cmb婴儿性别(i_intIndex).NewIndex
                cmb婴儿性别(i_intIndex).ItemData(cmb婴儿性别(i_intIndex).NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
    Else
        LoadComboFromArray Array("1.男", "2.女", "3.未知", "4.不明"), cmb婴儿性别(i_intIndex)
    End If
    
    txt婴儿体重(i_intIndex).Text = ""
    txtApgar评分(i_intIndex).Text = ""
    If mrs胎儿情况.RecordCount < i_intIndex + 1 Then
        If mbytModel = 1 Then
            dtp分娩时间(i_intIndex).Value = CDate(txt出院日期.Text)
        Else
            dtp分娩时间(i_intIndex).Value = dtp分娩时间(0).Value
        End If
    End If
    '初始化新生儿疾病表格
    With bill新生儿疾病(i_intIndex)
        .AllowAddRow = False
        .Font.Size = 10.5
        .Rows = 5
        .Cols = 3
        .MsfObj.FixedCols = 1
        .Clear
        .TextMatrix(0, col类型) = "诊断类型":  .ColData(0) = 5:  .ColWidth(0) = 1400:  .ColAlignment(0) = 1
        .TextMatrix(0, Col编码) = "ICD编码":   .ColData(1) = 1:  .ColWidth(1) = 1250:  .ColAlignment(1) = 1
        .TextMatrix(0, Col诊断) = "诊断描述":  .ColData(2) = 5:  .ColWidth(2) = 3200:  .ColAlignment(2) = 1
        .PrimaryCol = 1
        .LocateCol = 1
        For intRow = 1 To 4
            .TextMatrix(intRow, col类型) = "新生儿疾病" & intRow
            .TextMatrix(intRow, Col编码) = ""
            .TextMatrix(intRow, Col诊断) = ""
            .RowData(intRow) = 0
        Next
        .Active = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEditable And Not mblnOK Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
        mstrSumTime = ""
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate产程1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate产程1_LostFocus()
    If Trim(MaskDate产程1.Text <> "__:__:__") Then
        Call GetSumTime
    End If
End Sub

Private Sub MaskDate产程2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate产程2_LostFocus()
    If Trim(MaskDate产程2.Text <> "__:__:__") Then
        Call GetSumTime
    End If
End Sub

Private Sub MaskDate产程3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub MaskDate产程3_LostFocus()
    If Trim(MaskDate产程3.Text <> "__:__:__") Then
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
    If Not IsDate(MaskDate产程1.Text) And Trim(MaskDate产程1.Text <> "__:__:__") Then
        MsgBox "产程时间1的时间格式不正确,请检查！", vbInformation, gstrSysName
        If CanFocus(MaskDate产程1) Then MaskDate产程1.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(MaskDate产程2.Text) And Trim(MaskDate产程2.Text <> "__:__:__") Then
        MsgBox "产程时间2的时间格式不正确,请检查！", vbInformation, gstrSysName
        If CanFocus(MaskDate产程2) Then MaskDate产程2.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(MaskDate产程3.Text) And Trim(MaskDate产程3.Text <> "__:__:__") Then
        MsgBox "产程时间3的时间格式不正确,请检查！", vbInformation, gstrSysName
        If CanFocus(MaskDate产程3) Then MaskDate产程3.SetFocus
        Exit Sub
    End If
    
    strDate1 = IIf(Format(MaskDate产程1.Text, "HH:MM:SS") = "00:00:00", "", Format(MaskDate产程1.Text, "HH:MM:SS"))
    strDate2 = IIf(Format(MaskDate产程2.Text, "HH:MM:SS") = "00:00:00", "", Format(MaskDate产程2.Text, "HH:MM:SS"))
    strDate3 = IIf(Format(MaskDate产程3.Text, "HH:MM:SS") = "00:00:00", "", Format(MaskDate产程3.Text, "HH:MM:SS"))
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
    txt总产程时间.Text = Trim(IIf(Len(Trim(intH)) = 1, 0 & intH, intH)) & ":" & Trim(IIf(Len(Trim(intM)) = 1, 0 & intM, intM)) & ":" & IIf(Len(Trim(intS)) = 1, 0 & intS, intS)
    mstrSumTime = txt总产程时间.Text
End Sub

Private Sub tab胎数_Click()
    Dim intRow As Integer
    Dim intIndex As Integer
    
    intIndex = tab胎数.SelectedItem.Index - 1
    For intRow = fra胎儿信息.LBound To fra胎儿信息.UBound
        fra胎儿信息(intRow).Visible = False
    Next
    fra胎儿信息(intIndex).Visible = True
    Set fra胎儿信息(intIndex).Container = tab胎数.Container
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
'功能：处理诊断项目的输入
    If InStr(rsInput!编码 & "", "*") > 0 Then
        MsgBox "星号编码不能作为主要编码。", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(rsInput!编码 & "", 1) = "R" Then
        If MsgBox("你现在正使用R编码作为主要编码，是否确认？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    With bill新生儿疾病(intIndex)
        .RowData(LngRow) = rsInput!项目ID
        .TextMatrix(LngRow, Col编码) = rsInput!编码
        .Text = rsInput!编码
        .TextMatrix(LngRow, Col诊断) = rsInput!名称
    End With
    SetBillInput = True
End Function

Private Function Validate() As Boolean
    Dim ctlError As Control
    Dim strMessage As String

    '检查分娩信息的数据有效性
    If Not Validate分娩信息(ctlError, strMessage) Then
        If strMessage <> "" Then MsgBox strMessage, vbInformation, gstrSysName
        If CanFocus(ctlError) Then ctlError.SetFocus
        Exit Function
    End If
    
    '检查婴儿信息的数据有效性
    If Not Validate胎儿信息(ctlError, strMessage) Then
        If strMessage <> "" Then MsgBox strMessage, vbInformation, gstrSysName
        '问题22275完善 by lesfeng 2009-09-24
'        If CanFocus(ctlError) Then ctlError.SetFocus
        If CanFocus(ctlError) Then
            If Not ctlError.Enabled Then ctlError.Enabled = True
            ctlError.SetFocus
        End If
        Exit Function
    End If
    
    Validate = True
End Function

Public Sub Save分娩信息()
    Dim intIndex As Integer, intRow As Integer, int胎数 As Integer
    Dim strFilter As String
    Dim arrTmp As Variant, strData As String, blnNew As Boolean
    Dim rsNew胎儿情况 As ADODB.Recordset
    Dim rsNew新生儿诊断 As ADODB.Recordset
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim arrValues() As Variant
    Dim strOld As String
    Dim arrSQL() As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    '将界面信息保存
    arrTmp = Split("产检次数,胎次,胎数,产程时间1,产程时间2,产程时间3,总产程时间,产后出血量,产科并发症,会阴Ⅲ度裂伤", ",")
    
    For intIndex = LBound(arrTmp) To UBound(arrTmp)
        mrs分娩信息.Filter = "信息名='" & arrTmp(intIndex) & "'"
        blnNew = mrs分娩信息.EOF
        If Not blnNew Then
            strOld = mrs分娩信息!信息值 & ""
        End If
        Select Case arrTmp(intIndex)
            Case "产检次数"
                strData = Trim(txt产检次数.Text)
            Case "胎次"
                strData = Trim(txt胎次.Text)
            Case "胎数"
                strData = Get数据FromCombo(cmb胎数)
            Case "产程时间1"
                strData = MaskDate产程1.Text
            Case "产程时间2"
                strData = MaskDate产程2.Text
            Case "产程时间3"
                strData = MaskDate产程3.Text
            Case "总产程时间"
                strData = txt总产程时间.Text
            Case "产后出血量"
                strData = Trim(txt产后出血量.Text)
            Case "产科并发症"
                strData = Get数据FromCombo(cbo并发症)
            Case "会阴Ⅲ度裂伤"
                strData = Get数据FromCombo(cbo会阴)
        End Select
        If blnNew Then
            mrs分娩信息.AddNew Array("信息名", "信息现值", "类型", "记录性质"), Array(arrTmp(intIndex), strData, 0, IIf(strOld & "" = strData, 0, 1))
        Else
            mrs分娩信息.Update Array("信息现值", "记录性质"), Array(strData, 1)
        End If
    Next
    
    Set rsNew胎儿情况 = zlDatabase.CopyNewRec(mrs胎儿情况, True)
    Set rsNew新生儿诊断 = zlDatabase.CopyNewRec(mrs新生儿诊断, True)
    arrMainFileds = Array("病人ID", "主页ID", "分娩时间", "胎儿次序", "分娩方式", "出生胎位", "分娩情况", "出生缺陷", "婴儿性别", "婴儿体重", "Apgar评分", "记录性质")
    arrSecdFileds = Array("病人ID", "主页ID", "胎儿次序", "诊断次序", "疾病id", "编码", "描述信息", "记录性质")
    
    For intIndex = 1 To tab胎数.Tabs.Count
        With rsNew胎儿情况
'            .AddNew
            .AddNew arrMainFileds, Array(mlng病人ID, mlng主页ID, dtp分娩时间(intIndex - 1).Value, intIndex, Get数据FromCombo(cmb分娩方式(intIndex - 1)), Get数据FromCombo(cmb出生胎位(intIndex - 1)), Get数据FromCombo(cmb分娩情况(intIndex - 1)), _
                        Val(cmb出生缺陷(intIndex - 1).ListIndex), Get数据FromCombo(cmb婴儿性别(intIndex - 1)), Format(Val(txt婴儿体重(intIndex - 1).Text) / 1000, "0.000"), Trim(txtApgar评分(intIndex - 1).Text), 0)
'            !病人ID = mlng病人ID
'            !主页ID = 主页ID
'            !胎儿次序 = intIndex
'            !分娩方式 = Get数据FromCombo(cmb分娩方式(intIndex - 1))
'            !出生胎位 = Get数据FromCombo(cmb出生胎位(intIndex - 1))
'            !分娩情况 = Get数据FromCombo(cmb分娩情况(intIndex - 1))
'            !出生缺陷 = Val(Get数据FromCombo(cmb出生缺陷(intIndex - 1)))
'            !婴儿性别 = Get数据FromCombo(cmb婴儿性别(intIndex - 1))
'            !婴儿体重 = Trim(txt婴儿体重(intIndex - 1).Text)
'            !Apgar评分 = Trim(txtApgar评分(intIndex - 1).Text)
'            !记录性质 = 0
            .Update
        End With
        
        For intRow = 1 To 4
            strFilter = bill新生儿疾病(intIndex - 1).TextMatrix(intRow, Col编码)
            If strFilter <> "" Then
                With rsNew新生儿诊断
                    .AddNew arrSecdFileds, Array(mlng病人ID, mlng主页ID, intIndex, intRow, bill新生儿疾病(intIndex - 1).RowData(intRow), bill新生儿疾病(intIndex - 1).TextMatrix(intRow, Col编码), Replace(bill新生儿疾病(intIndex - 1).TextMatrix(intRow, Col诊断), "'", "’"), 0)
                    .Update
                End With
            End If
        Next
    Next
    
    '记录集比较
    mrs胎儿情况.Filter = "": rsNew胎儿情况.Filter = ""
    mrs新生儿诊断.Filter = "": rsNew新生儿诊断.Filter = ""
    If Rec.Compare(mrs胎儿情况, rsNew胎儿情况) Then
        If Not Rec.Compare(mrs新生儿诊断, rsNew新生儿诊断) Then
            Set mrs新生儿诊断 = rsNew新生儿诊断
            Call Rec.Update(mrs新生儿诊断, "", "记录性质", 1) '标记信息已经改变
        End If
    Else
        Set mrs胎儿情况 = rsNew胎儿情况
        Call Rec.Update(mrs胎儿情况, "", "记录性质", 1) '标记信息已经改变
        Set mrs新生儿诊断 = rsNew新生儿诊断
        Call Rec.Update(mrs新生儿诊断, "", "记录性质", 1) '标记信息已经改变
    End If
    If mbytModel = 2 Then
    '新生儿登记
        arrSQL = Array()
        Set grsDeliceryInfo = mrs分娩信息
        Set grsBabyInfo = mrs胎儿情况
        Set grsBabyDiag = mrs新生儿诊断
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

Private Function Get数据FromCombo(ctlTemp As ComboBox, Optional ByVal blnID As Boolean = False) As String
    If blnID = False Then
        If ctlTemp.Text = "" Or ctlTemp.Enabled = False Then
            Get数据FromCombo = ""
        Else
            Get数据FromCombo = Mid(ctlTemp.Text, InStr(ctlTemp.Text, ".") + 1)
        End If
    Else
        Get数据FromCombo = ctlTemp.ItemData(ctlTemp.ListIndex)
    End If
End Function

Private Function CanFocus(ctlError As Control) As Boolean
    If TypeName(ctlError) = "BillEdit" Then
        CanFocus = True
    Else
        CanFocus = ctlError.Enabled
    End If
End Function

Private Function Validate分娩信息(ctlError As Control, strMessage As String) As Boolean
    
    '56928:刘鹏飞,2013-04-25,病人可能在入院之前就开始分娩了
'    If CDate(dtp分娩时间.Value) < CDate(txt入院日期) Then
'        Set ctlError = dtp分娩时间
'        strMessage = "分娩时间不能比入院日期还早!"
'        Exit Function
'    End If
    '刘飞2005-8-10有当天分娩就当天出院的情况
    'if dtp分娩时间.value-1>txt出院日期 then
    
    If Not IsNumeric(txt产检次数) Then
        Set ctlError = txt产检次数
        strMessage = "产检次数填写不正确,请检查!"
        Exit Function
    End If
    
    If Not IsNumeric(txt胎次) Then
        Set ctlError = txt胎次
        strMessage = "胎次填写不正确,请检查!"
        Exit Function
    End If
    
    If Not IsNumeric(txt产后出血量) Then
        Set ctlError = txt产后出血量
        strMessage = "产后出血量填写不正确,请检查!"
        Exit Function
    End If
    
    If Not IsDate(MaskDate产程1.Text) And MaskDate产程1.Text <> "__:__:__" Then
        Set ctlError = MaskDate产程1
        strMessage = "产程时间1的时间格式不正确,请检查！"
        Exit Function
    End If
    
    If Not IsDate(MaskDate产程2.Text) And MaskDate产程2.Text <> "__:__:__" Then
        Set ctlError = MaskDate产程2
        strMessage = "产程时间2的时间格式不正确,请检查！"
        Exit Function
    End If
    
    If Not IsDate(MaskDate产程3.Text) And MaskDate产程3.Text <> "__:__:__" Then
        Set ctlError = MaskDate产程3
        strMessage = "产程时间3的时间格式不正确,请检查！"
        Exit Function
    End If
    
    Validate分娩信息 = True
End Function

Private Function Validate胎儿信息(ctlError As Control, strMessage As String) As Boolean
    Dim intIndex As Integer
    Dim lng疾病ID As Long, j As Long
    
    
    For intIndex = 0 To tab胎数.Tabs.Count - 1
        If Val(Left(cmb分娩情况(intIndex).Text, 1)) = 1 Then
            '只有活产才会输入以下内容
            '问题:22275
            If Not IsNumeric(txt婴儿体重(intIndex)) Then
                tab胎数.Tabs(intIndex + 1).Selected = True
                'fra胎儿信息(intIndex).Visible = True
                Set ctlError = txt婴儿体重(intIndex)
                strMessage = "婴儿体重填写不正确,请检查!"
                Exit Function
            End If
            
            If Not IsNumeric(txtApgar评分(intIndex)) Then
                tab胎数.Tabs(intIndex + 1).Selected = True
                'fra胎儿信息(intIndex).Visible = True
                Set ctlError = txtApgar评分(intIndex)
                strMessage = "Apgar评分填写不正确,请检查!"
                Exit Function
            End If
        End If
        '问题:22275
        '检查分方式是否正确
        If cmb分娩方式(0).Text <> cmb分娩方式(intIndex).Text Then
            Set ctlError = cmb分娩方式(intIndex)
            tab胎数.Tabs(intIndex + 1).Selected = True
            If MsgBox("第" & intIndex + 1 & "胎的分娩方式与第1胎分娩方式不一致,是否继续?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        If mbytModel = 1 And intIndex = 0 Then
            If CDate(dtp分娩时间(intIndex).Value) > CDate(txt出院日期) Then
                If Not fra胎儿信息(intIndex).Visible Then
                    fra胎儿信息(intIndex).Visible = True
                    Set ctlError = dtp分娩时间(intIndex)
                    strMessage = "分娩时间不能比出院日期还晚!"
                Else
                    Set ctlError = dtp分娩时间(intIndex)
                    strMessage = "分娩时间不能比出院日期还晚!"
                End If
                Exit Function
            End If
        Else
            If intIndex > 0 Then
                If CDate(dtp分娩时间(intIndex).Value) < CDate(dtp分娩时间(intIndex - 1).Value) Then
                    If MsgBox("第【" & intIndex + 1 & "】个胎儿的分娩时间比第【" & intIndex & "】个胎儿的分娩时间小，确定保存吗？", vbYesNo, gstrSysName) = vbNo Then
                        If Not fra胎儿信息(intIndex).Visible Then
                            fra胎儿信息(intIndex).Visible = True
                            Set ctlError = dtp分娩时间(intIndex)
                        Else
                            Set ctlError = dtp分娩时间(intIndex)
                        End If
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Validate胎儿信息 = True
End Function

Private Sub LoadData()
    Dim intIndex As Integer, intRow As Integer
    Dim strData As String
    
    mrs分娩信息.Filter = "类型=0": mrs分娩信息.Sort = "信息名"
    If mrs分娩信息.RecordCount = 0 Then Exit Sub
    
    mintFlag = 1
    Do While Not mrs分娩信息.EOF
        strData = mrs分娩信息!信息现值 & ""
        Select Case mrs分娩信息!信息名 & ""
            Case "产检次数"
                txt产检次数.Text = IIf(Val(strData) = 0, "", strData)
            Case "胎次"
                txt胎次.Text = IIf(Val(strData) = 0, "", strData)
            Case "胎数"
                AddComboItem cmb胎数, CStr(strData), 2
            Case "产程时间1"
                MaskDate产程1.Text = Format(strData, "HH:MM:SS")
            Case "产程时间2"
                MaskDate产程2.Text = Format(strData, "HH:MM:SS")
            Case "产程时间3"
                MaskDate产程3.Text = Format(strData, "HH:MM:SS")
            Case "总产程时间"
                txt总产程时间.Text = Format(strData, "HH:MM:SS")
            Case "产后出血量"
                txt产后出血量.Text = IIf(Val(strData) = 0, "", strData)
            Case "产科并发症"
                AddComboItem cbo并发症, CStr(strData)
            Case "会阴Ⅲ度裂伤"
                AddComboItem cbo会阴, CStr(strData)
        End Select
        mrs分娩信息.MoveNext
    Loop
    cmb胎数_Click
    For intIndex = 1 To tab胎数.Tabs.Count
        With mrs胎儿情况
            .Filter = "胎儿次序 = " & intIndex
            If Not .EOF Then
                If IsNull(!分娩时间) Then
                    mrs分娩信息.Filter = "信息名='分娩时间'"
                    If Not mrs分娩信息.EOF Then dtp分娩时间(intIndex - 1).Value = CDate("" & mrs分娩信息!信息值)
                Else
                    dtp分娩时间(intIndex - 1).Value = CDate(!分娩时间)
                End If
                AddComboItem cmb分娩方式(intIndex - 1), IIf(!分娩方式 = "其他", "其他分娩", !分娩方式)
                AddComboItem cmb出生胎位(intIndex - 1), !出生胎位
                AddComboItem cmb分娩情况(intIndex - 1), !分娩情况
                AddComboItem cmb婴儿性别(intIndex - 1), !婴儿性别
                cmb出生缺陷(intIndex - 1).ListIndex = Val(!出生缺陷 & "")
                txt婴儿体重(intIndex - 1) = IIf(IsNull(!婴儿体重), "", Val(Nvl(!婴儿体重, "")) * 1000)
                txtApgar评分(intIndex - 1) = IIf(IsNull(!Apgar评分), "", !Apgar评分)
            End If
        End With
        
        With mrs新生儿诊断
            For intRow = 1 To 4
                .Filter = "胎儿次序 = " & intIndex & " and 诊断次序 = " & intRow
                If Not .EOF Then
                    bill新生儿疾病(intIndex - 1).TextMatrix(intRow, Col编码) = !编码
                    bill新生儿疾病(intIndex - 1).TextMatrix(intRow, Col诊断) = !描述信息
                    bill新生儿疾病(intIndex - 1).RowData(intRow) = !疾病id
                End If
            Next
        End With
    Next
    
    mrs胎儿情况.Filter = ""
    mrs新生儿诊断.Filter = ""
End Sub

Private Function AddComboItem(cmbTemp As Control, strItem As String, Optional ByVal cmbType As Integer = 1, Optional ByVal cmbData As Long) As Boolean
    '参数cmbType  = 1时表示下拉框由数字打头，
    '             = 2时表示全是文字
    Dim varTemp As Variant
    '该项在列表框中
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

Private Sub tab胎数_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtApgar评分_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtApgar评分(Index)
End Sub

Private Sub txtApgar评分_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtApgar评分_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt产后出血量_GotFocus()
    zlControl.TxtSelAll txt产检次数
End Sub

Private Sub txt产后出血量_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt产后出血量_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt产检次数_GotFocus()
    zlControl.TxtSelAll txt产检次数
End Sub

Private Sub txt产检次数_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt产检次数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt胎次_GotFocus()
    zlControl.TxtSelAll txt胎次
End Sub

Private Sub txt胎次_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt胎次_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt婴儿体重_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt婴儿体重_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Function GetInDate() As Date
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim DatRet As Date
    
    On Error GoTo errH
    strSql = "Select 入院日期 From 病案主页 T Where t.病人id = [1] And t.主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        DatRet = CDate(Nvl(rsTmp!入院日期, "0:00:00"))
    Else
        DatRet = CDate("0:00:00")
    End If
    GetInDate = DatRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt总产程时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt总产程时间_Validate(Cancel As Boolean)
    If Format(mstrSumTime, "HH:MM:SS") <> "" Then
        If Format(txt总产程时间.Text, "HH:MM:SS") <> "" Then
            If Format(txt总产程时间.Text, "HH:MM:SS") <> Format(mstrSumTime, "HH:MM:SS") Then
                If MsgBox("总产程时间为" & Format(txt总产程时间.Text, "HH:MM:SS") & "不等于产程时间1+产程时间2+产程时间3之和" & Format(mstrSumTime, "HH:MM:SS") & ",是否继续?", vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                    Cancel = False
                Else
                    Cancel = True
                End If
            Else
                Cancel = False
            End If
        Else
            If MsgBox("总产程时间为空,产程时间1+产程时间2+产程时间3之和为" & Format(mstrSumTime, "HH:MM:SS") & ",是否继续?", vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                Cancel = False
            Else
                Cancel = True
            End If
        End If
    End If
End Sub
