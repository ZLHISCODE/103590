VERSION 5.00
Begin VB.Form frmReportUS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame framType 
      Caption         =   "��  ǻ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Index           =   15
      Left            =   0
      TabIndex        =   473
      Top             =   4080
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cbx��ǻ��Ϣ 
         Height          =   300
         Index           =   5
         ItemData        =   "frmReportUS.frx":0000
         Left            =   4200
         List            =   "frmReportUS.frx":000D
         Style           =   1  'Simple Combo
         TabIndex        =   281
         Tag             =   "�����Ƥ��:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx��ǻ��Ϣ 
         Height          =   300
         Index           =   4
         ItemData        =   "frmReportUS.frx":0021
         Left            =   2520
         List            =   "frmReportUS.frx":002E
         Style           =   1  'Simple Combo
         TabIndex        =   280
         Tag             =   "Һ�԰���:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx��ǻ��Ϣ 
         Height          =   300
         Index           =   2
         ItemData        =   "frmReportUS.frx":0042
         Left            =   4200
         List            =   "frmReportUS.frx":004F
         Style           =   1  'Simple Combo
         TabIndex        =   277
         Tag             =   "�����Ƥ��: [value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx��ǻ��Ϣ 
         Height          =   300
         Index           =   1
         ItemData        =   "frmReportUS.frx":0063
         Left            =   2520
         List            =   "frmReportUS.frx":0070
         Style           =   1  'Simple Combo
         TabIndex        =   276
         Tag             =   "Һ�԰���: [value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chk��ǻ��Ϣ 
         Caption         =   "�Ѷ�λ"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   282
         Tag             =   "�� �� λ[value]"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cbx��ǻ��Ϣ 
         Height          =   300
         Index           =   3
         ItemData        =   "frmReportUS.frx":0084
         Left            =   840
         List            =   "frmReportUS.frx":0091
         TabIndex        =   279
         Tag             =   "�����ǻ:[value]"
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chk��ǻ��Ϣ 
         Caption         =   "�Ѷ�λ"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   278
         Tag             =   "�Ѷ�λ[value]"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbx��ǻ��Ϣ 
         Height          =   300
         Index           =   0
         ItemData        =   "frmReportUS.frx":00A5
         Left            =   840
         List            =   "frmReportUS.frx":00B2
         TabIndex        =   275
         Tag             =   "�����ǻ: [value]"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lable1 
         Caption         =   "�Ҳ���ǻ          Һ�԰���      cm �����Ƥ��      cm"
         Height          =   255
         Index           =   145
         Left            =   120
         TabIndex        =   475
         Top             =   630
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "�����ǻ          Һ�԰���      cm �����Ƥ��      cm"
         Height          =   255
         Index           =   148
         Left            =   120
         TabIndex        =   474
         Top             =   270
         Width           =   4815
      End
   End
   Begin VB.Frame framType 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Index           =   14
      Left            =   0
      TabIndex        =   476
      Top             =   3840
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   11
         Left            =   3720
         TabIndex        =   274
         Tag             =   "[T7]AT:[value]s"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   10
         Left            =   2760
         TabIndex        =   273
         Tag             =   "[T7]RI:[value]"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   9
         Left            =   1485
         TabIndex        =   272
         Tag             =   "[T7]PS:[value]cm/s"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   8
         Left            =   3720
         TabIndex        =   271
         Tag             =   "[T6]AT: [value]s"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   7
         Left            =   2760
         TabIndex        =   270
         Tag             =   "[T6]RI: [value]"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   6
         Left            =   1485
         TabIndex        =   269
         Tag             =   "[T6]PS: [value]cm/s"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   5
         Left            =   3720
         TabIndex        =   268
         Tag             =   "[T7]AT:[value]s"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   4
         Left            =   2760
         TabIndex        =   267
         Tag             =   "[T7]RI:[value]"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   3
         Left            =   1485
         TabIndex        =   266
         Tag             =   "[T7]PS:[value]cm/s"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   2
         Left            =   3720
         TabIndex        =   265
         Tag             =   "[T6]AT: [value]s"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   264
         Tag             =   "[T6]RI: [value]"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt��������Ϣ 
         Height          =   270
         Index           =   0
         Left            =   1480
         TabIndex        =   263
         Tag             =   "[T6]PS: [value]cm/s"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "��������Զ��:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   150
         Left            =   120
         TabIndex        =   480
         Top             =   1350
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "������������:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   149
         Left            =   120
         TabIndex        =   479
         Top             =   990
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "��������Զ��:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   146
         Left            =   120
         TabIndex        =   478
         Top             =   630
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "������������:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   147
         Left            =   120
         TabIndex        =   477
         Top             =   270
         Width           =   4815
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��״��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Index           =   13
      Left            =   0
      TabIndex        =   456
      Top             =   3600
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   9
         Left            =   1320
         TabIndex        =   262
         Tag             =   "��״�������:[value]cm^3"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   8
         Left            =   4080
         TabIndex        =   261
         Tag             =   "Ͽ����״�ٿ�:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   7
         Left            =   2760
         TabIndex        =   260
         Tag             =   "Ͽ����״�ٺ�:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   6
         Left            =   1320
         TabIndex        =   259
         Tag             =   "Ͽ����״�ٳ�:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   5
         Left            =   4080
         TabIndex        =   258
         Tag             =   "�Ҳ��״�ٿ�:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   4
         Left            =   2760
         TabIndex        =   257
         Tag             =   "�Ҳ��״�ٺ�:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   3
         Left            =   1320
         TabIndex        =   256
         Tag             =   "�Ҳ��״�ٳ�:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   2
         Left            =   4080
         TabIndex        =   255
         Tag             =   "����״�ٿ�:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   254
         Tag             =   "����״�ٺ�:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt��״����Ϣ 
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   253
         Tag             =   "����״�ٳ�:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "��״�������       cm^3"
         Height          =   255
         Index           =   138
         Left            =   120
         TabIndex        =   460
         Top             =   1365
         Width           =   2175
      End
      Begin VB.Label lable1 
         Caption         =   "Ͽ����״��:��      cm      ��      cm    ��       cm"
         Height          =   255
         Index           =   137
         Left            =   120
         TabIndex        =   459
         Top             =   1005
         Width           =   4935
      End
      Begin VB.Label lable1 
         Caption         =   "�Ҳ��״��:��      cm      ��      cm    ��       cm"
         Height          =   255
         Index           =   134
         Left            =   120
         TabIndex        =   458
         Top             =   645
         Width           =   4935
      End
      Begin VB.Label lable1 
         Caption         =   "����״��:��      cm      ��      cm    ��       cm"
         Height          =   255
         Index           =   132
         Left            =   120
         TabIndex        =   457
         Top             =   280
         Width           =   4935
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Index           =   12
      Left            =   0
      TabIndex        =   454
      Top             =   3360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txt�۲���Ϣ 
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   248
         Tag             =   "�����᳤:[value]cm"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txt�۲���Ϣ 
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   252
         Tag             =   "�����᳤:[value]cm"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "�����᳤      cm      �����᳤      cm"
         Height          =   255
         Index           =   133
         Left            =   120
         TabIndex        =   455
         Top             =   390
         Width           =   4935
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��֫����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2415
      Index           =   11
      Left            =   0
      TabIndex        =   468
      Top             =   3120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "����֫����"
         Height          =   975
         Index           =   29
         Left            =   120
         TabIndex        =   470
         Top             =   240
         Width           =   5655
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "�㱳����"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   245
            Tag             =   "�㱳����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "�N����"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   244
            Tag             =   "�N����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "��ǳ����"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   243
            Tag             =   "��ǳ����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "�����"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   242
            Tag             =   "�����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "���ܾ���"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   241
            Tag             =   "���ܾ���[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "Ѫ��ͨ��,��ǻ��δ���쳣����,̽����ѹ���ǻ��ʧ"
            Height          =   255
            Index           =   143
            Left            =   120
            TabIndex        =   471
            Top             =   600
            Width           =   4455
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "����֫����"
         Height          =   975
         Index           =   28
         Left            =   120
         TabIndex        =   469
         Top             =   1320
         Width           =   5655
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "�㱳����"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   251
            Tag             =   "�㱳����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "�N����"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   250
            Tag             =   "�N����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "��ǳ����"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   249
            Tag             =   "��ǳ����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "�����"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   247
            Tag             =   "�����[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk����֫���� 
            Caption         =   "���ܾ���"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   246
            Tag             =   "���ܾ���[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "Ѫ��ͨ��,��ǻ��δ���쳣����,̽����ѹ���ǻ��ʧ"
            Height          =   255
            Index           =   144
            Left            =   120
            TabIndex        =   472
            Top             =   600
            Width           =   4455
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��ֳ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2415
      Index           =   10
      Left            =   0
      TabIndex        =   461
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "�Ҳ���ֳ��"
         Height          =   975
         Index           =   27
         Left            =   120
         TabIndex        =   463
         Top             =   1320
         Width           =   5655
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":00C6
            Left            =   3120
            List            =   "frmReportUS.frx":00D6
            Style           =   1  'Simple Combo
            TabIndex        =   239
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":00EE
            Left            =   1920
            List            =   "frmReportUS.frx":00FE
            Style           =   1  'Simple Combo
            TabIndex        =   238
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0116
            Left            =   720
            List            =   "frmReportUS.frx":0126
            Style           =   1  'Simple Combo
            TabIndex        =   237
            Tag             =   "��غ��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":013E
            Left            =   3120
            List            =   "frmReportUS.frx":014E
            Style           =   1  'Simple Combo
            TabIndex        =   235
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0166
            Left            =   1920
            List            =   "frmReportUS.frx":0176
            Style           =   1  'Simple Combo
            TabIndex        =   234
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":018E
            Left            =   720
            List            =   "frmReportUS.frx":019E
            Style           =   1  'Simple Combo
            TabIndex        =   233
            Tag             =   "غ�賤:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   7
            ItemData        =   "frmReportUS.frx":01B6
            Left            =   4440
            List            =   "frmReportUS.frx":01C6
            TabIndex        =   240
            Tag             =   "Ѫ��:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx�Ҳ���ֳ�� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":01DE
            Left            =   4440
            List            =   "frmReportUS.frx":01EE
            TabIndex        =   236
            Tag             =   "Ѫ��: [value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��غ��      cm    ��      cm  ��       cm  Ѫ��"
            Height          =   255
            Index           =   141
            Left            =   120
            TabIndex        =   467
            Top             =   645
            Width           =   4335
         End
         Begin VB.Label lable1 
            Caption         =   "غ�賤      cm    ��      cm  ��       cm  Ѫ��"
            Height          =   255
            Index           =   140
            Left            =   120
            TabIndex        =   466
            Top             =   285
            Width           =   4335
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�����ֳ��"
         Height          =   975
         Index           =   26
         Left            =   120
         TabIndex        =   462
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":0206
            Left            =   3120
            List            =   "frmReportUS.frx":0216
            Style           =   1  'Simple Combo
            TabIndex        =   231
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":022E
            Left            =   1920
            List            =   "frmReportUS.frx":023E
            Style           =   1  'Simple Combo
            TabIndex        =   230
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0256
            Left            =   720
            List            =   "frmReportUS.frx":0266
            Style           =   1  'Simple Combo
            TabIndex        =   229
            Tag             =   "��غ��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":027E
            Left            =   3120
            List            =   "frmReportUS.frx":028E
            Style           =   1  'Simple Combo
            TabIndex        =   227
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":02A6
            Left            =   1920
            List            =   "frmReportUS.frx":02B6
            Style           =   1  'Simple Combo
            TabIndex        =   226
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":02CE
            Left            =   720
            List            =   "frmReportUS.frx":02DE
            Style           =   1  'Simple Combo
            TabIndex        =   225
            Tag             =   "غ�賤:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   7
            ItemData        =   "frmReportUS.frx":02F6
            Left            =   4440
            List            =   "frmReportUS.frx":0306
            TabIndex        =   232
            Tag             =   "Ѫ��:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx�����ֳ�� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":031E
            Left            =   4440
            List            =   "frmReportUS.frx":032E
            TabIndex        =   228
            Tag             =   "Ѫ��: [value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��غ��      cm    ��      cm  ��       cm  Ѫ��"
            Height          =   255
            Index           =   139
            Left            =   120
            TabIndex        =   465
            Top             =   645
            Width           =   4335
         End
         Begin VB.Label lable1 
            Caption         =   "غ�賤      cm    ��      cm  ��       cm  Ѫ��"
            Height          =   255
            Index           =   142
            Left            =   120
            TabIndex        =   464
            Top             =   285
            Width           =   4335
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Index           =   9
      Left            =   0
      TabIndex        =   453
      Top             =   2640
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "�������"
         Height          =   615
         Index           =   31
         Left            =   120
         TabIndex        =   485
         Top             =   960
         Width           =   5655
         Begin VB.TextBox txt������Ϣ 
            Height          =   270
            Index           =   2
            Left            =   720
            TabIndex        =   487
            Tag             =   "�����:[value]cm "
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt������Ϣ 
            Height          =   270
            Index           =   3
            Left            =   3360
            TabIndex        =   486
            Tag             =   "���鵼���ھ�:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "�����       cm        ���鵼���ھ�       cm"
            Height          =   255
            Index           =   131
            Left            =   120
            TabIndex        =   488
            Top             =   270
            Width           =   4935
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�������"
         Height          =   615
         Index           =   30
         Left            =   120
         TabIndex        =   481
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txt������Ϣ 
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   483
            Tag             =   "�����:[value]cm "
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt������Ϣ 
            Height          =   270
            Index           =   1
            Left            =   3360
            TabIndex        =   482
            Tag             =   "���鵼���ھ�:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "�����       cm        ���鵼���ھ�       cm"
            Height          =   255
            Index           =   130
            Left            =   120
            TabIndex        =   484
            Top             =   270
            Width           =   4935
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2775
      Index           =   8
      Left            =   0
      TabIndex        =   445
      Top             =   2400
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   13
         Left            =   3120
         TabIndex        =   224
         Tag             =   "׵��������Ĥ��:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   12
         Left            =   1440
         TabIndex        =   223
         Tag             =   "׵�����ھ�:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   11
         Left            =   3480
         TabIndex        =   222
         Tag             =   "���⶯���ھ�:[value]cm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   10
         Left            =   1440
         TabIndex        =   221
         Tag             =   "�Ҿ��ڶ����ھ�:[value]cm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   7
         Left            =   3120
         TabIndex        =   218
         Tag             =   "�Ҿ��ܶ�������Ĥ��:[value]cm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   6
         Left            =   1440
         TabIndex        =   217
         Tag             =   "�Ҿ��ܶ����ھ�:[value]cm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   8
         Left            =   1920
         TabIndex        =   219
         Tag             =   "�Ҿ��ܶ�������ھ�:[value]cm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   9
         Left            =   3960
         TabIndex        =   220
         Tag             =   "�Ҿ��ܶ����������Ĥ��:[value]cm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   5
         Left            =   3480
         TabIndex        =   216
         Tag             =   "���⶯���ھ�: [value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   4
         Left            =   1440
         TabIndex        =   215
         Tag             =   "���ڶ����ھ�:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   1
         Left            =   3120
         TabIndex        =   212
         Tag             =   "���ܶ�������Ĥ��:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   211
         Tag             =   "���ܶ����ھ�:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   2
         Left            =   1920
         TabIndex        =   213
         Tag             =   "���ܶ�������ھ�[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   270
         Index           =   3
         Left            =   3960
         TabIndex        =   214
         Tag             =   "���ܶ����������Ĥ��:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "    ׵�����ھ�      cm  ����Ĥ��       cm"
         Height          =   255
         Index           =   129
         Left            =   120
         TabIndex        =   452
         Top             =   2430
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "�Ҿ��ڶ����ھ�      cm  ���⶯���ھ�       cm"
         Height          =   255
         Index           =   128
         Left            =   120
         TabIndex        =   451
         Top             =   2070
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "�Ҿ��ܶ����ھ�      cm  ����Ĥ��       cm"
         Height          =   255
         Index           =   127
         Left            =   120
         TabIndex        =   450
         Top             =   1350
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "�Ҿ��ܶ�������ھ�         cm  ����Ĥ��         cm"
         Height          =   255
         Index           =   126
         Left            =   120
         TabIndex        =   449
         Top             =   1710
         Width           =   5175
      End
      Begin VB.Label lable1 
         Caption         =   "���ڶ����ھ�      cm  ���⶯���ھ�       cm"
         Height          =   255
         Index           =   125
         Left            =   120
         TabIndex        =   448
         Top             =   990
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "���ܶ����ھ�      cm  ����Ĥ��       cm"
         Height          =   255
         Index           =   136
         Left            =   120
         TabIndex        =   447
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "���ܶ�������ھ�         cm  ����Ĥ��         cm"
         Height          =   255
         Index           =   135
         Left            =   120
         TabIndex        =   446
         Top             =   630
         Width           =   5175
      End
   End
   Begin VB.Frame framType 
      Caption         =   "����(Ů)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3615
      Index           =   7
      Left            =   0
      TabIndex        =   434
      Top             =   2160
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "����"
         Height          =   615
         Index           =   25
         Left            =   120
         TabIndex        =   443
         Top             =   2880
         Width           =   5655
         Begin VB.CheckBox chk����Ů 
            Caption         =   "�۲첻��"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   209
            Tag             =   "�۲첻��[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cbxŮ���� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0346
            Left            =   600
            List            =   "frmReportUS.frx":0356
            TabIndex        =   208
            Tag             =   "����:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chk����Ů 
            Caption         =   "δ�������쳣"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   210
            Tag             =   "δ�������쳣[value]"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lable1 
            Caption         =   "����"
            Height          =   255
            Index           =   124
            Left            =   120
            TabIndex        =   444
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�����"
         Height          =   1695
         Index           =   23
         Left            =   120
         TabIndex        =   435
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cbx����� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":0374
            Left            =   4220
            List            =   "frmReportUS.frx":0381
            Style           =   1  'Simple Combo
            TabIndex        =   207
            Tag             =   "�¶˰��ױ��ڶμ���ǿ�������:[value]cm����Ӱ"
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox cbx����� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":038F
            Left            =   3020
            List            =   "frmReportUS.frx":039C
            Style           =   1  'Simple Combo
            TabIndex        =   206
            Tag             =   "�϶��ھ�:[value]cm"
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cbx����� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":03AA
            Left            =   4220
            List            =   "frmReportUS.frx":03B7
            Style           =   1  'Simple Combo
            TabIndex        =   204
            Tag             =   "�¶˰��ױ��ڶμ���ǿ�������: [value]cm����Ӱ"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx����� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":03C5
            Left            =   3020
            List            =   "frmReportUS.frx":03D2
            Style           =   1  'Simple Combo
            TabIndex        =   203
            Tag             =   "�϶��ھ�: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx����� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":03E0
            Left            =   1080
            List            =   "frmReportUS.frx":03ED
            TabIndex        =   202
            Tag             =   "��������:[value]����"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cbx����� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":03FB
            Left            =   1080
            List            =   "frmReportUS.frx":0408
            TabIndex        =   205
            Tag             =   "�Ҳ������:[value]����"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "��������         ���� �϶��ھ�      cm  "
            Height          =   255
            Index           =   121
            Left            =   120
            TabIndex        =   439
            Top             =   285
            Width           =   4095
         End
         Begin VB.Label lable1 
            Caption         =   "�¶˰��ױ��ڶμ���ǿ�������      cm����Ӱ"
            Height          =   255
            Index           =   120
            Left            =   1680
            TabIndex        =   438
            Top             =   630
            Width           =   3855
         End
         Begin VB.Label lable1 
            Caption         =   "�Ҳ������         ���� �϶��ھ�      cm  "
            Height          =   255
            Index           =   119
            Left            =   120
            TabIndex        =   437
            Top             =   1005
            Width           =   4095
         End
         Begin VB.Label lable1 
            Caption         =   "�¶˰��ױ��ڶμ���ǿ�������      cm����Ӱ"
            Height          =   255
            Index           =   106
            Left            =   1680
            TabIndex        =   436
            Top             =   1350
            Width           =   3855
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�������"
         Height          =   975
         Index           =   24
         Left            =   120
         TabIndex        =   440
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txtŮ������� 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   198
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtŮ������� 
            Height          =   270
            Index           =   1
            Left            =   2040
            TabIndex        =   197
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtŮ������� 
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   196
            Tag             =   "������:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtŮ������� 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   201
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtŮ������� 
            Height          =   270
            Index           =   4
            Left            =   2040
            TabIndex        =   200
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtŮ������� 
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   199
            Tag             =   "������:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "������       cm    ��       cm    ��        cm"
            Height          =   255
            Index           =   123
            Left            =   120
            TabIndex        =   442
            Top             =   285
            Width           =   5415
         End
         Begin VB.Label lable1 
            Caption         =   "������       cm    ��       cm    ��        cm"
            Height          =   255
            Index           =   122
            Left            =   120
            TabIndex        =   441
            Top             =   645
            Width           =   5415
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "����(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4215
      Index           =   6
      Left            =   0
      TabIndex        =   421
      Top             =   1920
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "ǰ����"
         Height          =   615
         Index           =   19
         Left            =   120
         TabIndex        =   422
         Top             =   3480
         Width           =   5655
         Begin VB.TextBox txtǰ������Ϣ 
            Height          =   270
            Index           =   2
            Left            =   3120
            TabIndex        =   195
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtǰ������Ϣ 
            Height          =   270
            Index           =   1
            Left            =   1680
            TabIndex        =   194
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtǰ������Ϣ 
            Height          =   270
            Index           =   0
            Left            =   360
            TabIndex        =   193
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��       cm    ��       cm    ��       cm"
            Height          =   255
            Index           =   105
            Left            =   120
            TabIndex        =   423
            Top             =   285
            Width           =   4575
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "����"
         Height          =   615
         Index           =   20
         Left            =   120
         TabIndex        =   424
         Top             =   2880
         Width           =   5655
         Begin VB.CheckBox chk������ 
            Caption         =   "δ�������쳣"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   192
            Tag             =   "δ�������쳣[value]"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0416
            Left            =   600
            List            =   "frmReportUS.frx":0429
            TabIndex        =   190
            Tag             =   "����:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chk������ 
            Caption         =   "�۲첻��"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   191
            Tag             =   "�۲첻��[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "����"
            Height          =   255
            Index           =   107
            Left            =   120
            TabIndex        =   425
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�����"
         Height          =   1695
         Index           =   21
         Left            =   120
         TabIndex        =   426
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cbx������ 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":044F
            Left            =   4220
            List            =   "frmReportUS.frx":045C
            Style           =   1  'Simple Combo
            TabIndex        =   189
            Tag             =   "�¶˰��ױ��ڶμ���ǿ�������:[value]cm����Ӱ"
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox cbx������ 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":046A
            Left            =   3020
            List            =   "frmReportUS.frx":0477
            Style           =   1  'Simple Combo
            TabIndex        =   188
            Tag             =   "�϶��ھ�:[value]cm"
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cbx������ 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0485
            Left            =   4220
            List            =   "frmReportUS.frx":0492
            Style           =   1  'Simple Combo
            TabIndex        =   186
            Tag             =   "�¶˰��ױ��ڶμ���ǿ�������: [value]cm����Ӱ"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx������ 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":04A0
            Left            =   3020
            List            =   "frmReportUS.frx":04AD
            Style           =   1  'Simple Combo
            TabIndex        =   185
            Tag             =   "�϶��ھ�: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx������ 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":04BB
            Left            =   1080
            List            =   "frmReportUS.frx":04C8
            TabIndex        =   187
            Tag             =   "�Ҳ������:[value]����"
            Top             =   960
            Width           =   735
         End
         Begin VB.ComboBox cbx������ 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":04D6
            Left            =   1080
            List            =   "frmReportUS.frx":04E3
            TabIndex        =   184
            Tag             =   "��������:[value]����"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "�¶˰��ױ��ڶμ���ǿ�������      cm����Ӱ"
            Height          =   255
            Index           =   118
            Left            =   1680
            TabIndex        =   433
            Top             =   1350
            Width           =   3855
         End
         Begin VB.Label lable1 
            Caption         =   "�Ҳ������         ���� �϶��ھ�      cm  "
            Height          =   255
            Index           =   117
            Left            =   120
            TabIndex        =   432
            Top             =   1005
            Width           =   4095
         End
         Begin VB.Label lable1 
            Caption         =   "�¶˰��ױ��ڶμ���ǿ�������      cm����Ӱ"
            Height          =   255
            Index           =   116
            Left            =   1680
            TabIndex        =   431
            Top             =   630
            Width           =   3855
         End
         Begin VB.Label lable1 
            Caption         =   "��������         ���� �϶��ھ�      cm  "
            Height          =   255
            Index           =   108
            Left            =   120
            TabIndex        =   427
            Top             =   285
            Width           =   4095
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�������"
         Height          =   975
         Index           =   22
         Left            =   120
         TabIndex        =   428
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txt������� 
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   181
            Tag             =   "������:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt������� 
            Height          =   270
            Index           =   4
            Left            =   2040
            TabIndex        =   182
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt������� 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   183
            Tag             =   "��:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt������� 
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   178
            Tag             =   "������:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt������� 
            Height          =   270
            Index           =   1
            Left            =   2040
            TabIndex        =   179
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt������� 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   180
            Tag             =   "��: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "������       cm    ��       cm    ��        cm"
            Height          =   255
            Index           =   114
            Left            =   120
            TabIndex        =   430
            Top             =   645
            Width           =   5415
         End
         Begin VB.Label lable1 
            Caption         =   "������       cm    ��       cm    ��        cm"
            Height          =   255
            Index           =   115
            Left            =   120
            TabIndex        =   429
            Top             =   285
            Width           =   5415
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6975
      Index           =   5
      Left            =   0
      TabIndex        =   342
      Top             =   1680
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "Ƣ  ��"
         Height          =   975
         Index           =   14
         Left            =   120
         TabIndex        =   382
         Top             =   5880
         Width           =   5655
         Begin VB.ComboBox cbxƢ�� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":04F1
            Left            =   2880
            List            =   "frmReportUS.frx":0504
            Style           =   1  'Simple Combo
            TabIndex        =   174
            Tag             =   "Ƣ��-Ƣ�ⳤ��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbxƢ�� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":052E
            Left            =   600
            List            =   "frmReportUS.frx":0541
            Style           =   1  'Simple Combo
            TabIndex        =   173
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbxƢ�� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":056B
            Left            =   4320
            List            =   "frmReportUS.frx":0578
            TabIndex        =   177
            Tag             =   "��ɫѪ��:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbxƢ�� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":058A
            Left            =   600
            List            =   "frmReportUS.frx":059D
            TabIndex        =   176
            Tag             =   "����:[value]"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cbxƢ�� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":05C7
            Left            =   4320
            List            =   "frmReportUS.frx":05D4
            TabIndex        =   175
            Tag             =   "��̬:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��ɫѪ��"
            Height          =   255
            Index           =   47
            Left            =   3480
            TabIndex        =   387
            Top             =   645
            Width           =   855
         End
         Begin VB.Label lable1 
            Caption         =   "����"
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   386
            Top             =   645
            Width           =   405
         End
         Begin VB.Label lable1 
            Caption         =   "��̬"
            Height          =   255
            Index           =   46
            Left            =   3840
            TabIndex        =   385
            Top             =   285
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "��       cm"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   384
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "Ƣ�š�Ƣ�ⳤ��      cm"
            Height          =   255
            Index           =   48
            Left            =   1560
            TabIndex        =   383
            Top             =   285
            Width           =   2055
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "��  ��"
         Height          =   975
         Index           =   13
         Left            =   120
         TabIndex        =   374
         Top             =   4920
         Width           =   5655
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":05E6
            Left            =   4800
            List            =   "frmReportUS.frx":05F6
            Style           =   1  'Simple Combo
            TabIndex        =   169
            Tag             =   "�ȹ��ھ�:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0612
            Left            =   3360
            List            =   "frmReportUS.frx":0622
            Style           =   1  'Simple Combo
            TabIndex        =   168
            Tag             =   "��β��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":063E
            Left            =   2030
            List            =   "frmReportUS.frx":064E
            Style           =   1  'Simple Combo
            TabIndex        =   167
            Tag             =   "�����:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":066A
            Left            =   700
            List            =   "frmReportUS.frx":067A
            Style           =   1  'Simple Combo
            TabIndex        =   166
            Tag             =   "��ͷ��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":0696
            Left            =   4320
            List            =   "frmReportUS.frx":06A3
            TabIndex        =   172
            Tag             =   "��ɫѪ��:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":06B5
            Left            =   2280
            List            =   "frmReportUS.frx":06C8
            TabIndex        =   171
            Tag             =   "����:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx���� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":06F8
            Left            =   600
            List            =   "frmReportUS.frx":0708
            TabIndex        =   170
            Tag             =   "��Ĥ:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��ɫѪ��"
            Height          =   255
            Index           =   45
            Left            =   3480
            TabIndex        =   381
            Top             =   645
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "�ȹ��ھ�      cm"
            Height          =   255
            Index           =   44
            Left            =   4080
            TabIndex        =   380
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label lable1 
            Caption         =   "��β��      cm"
            Height          =   255
            Index           =   43
            Left            =   2760
            TabIndex        =   379
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "����"
            Height          =   255
            Index           =   42
            Left            =   1800
            TabIndex        =   378
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��Ĥ"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   377
            Top             =   645
            Width           =   405
         End
         Begin VB.Label lable1 
            Caption         =   "��ͷ��      cm"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   376
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "�����      cm"
            Height          =   255
            Index           =   41
            Left            =   1440
            TabIndex        =   375
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "���ܹ�"
         Height          =   975
         Index           =   12
         Left            =   120
         TabIndex        =   367
         Top             =   3960
         Width           =   5655
         Begin VB.ComboBox cbx���ܹ� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0724
            Left            =   2400
            List            =   "frmReportUS.frx":073D
            Style           =   1  'Simple Combo
            TabIndex        =   164
            Tag             =   "��С:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx���ܹ� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":076B
            Left            =   2400
            List            =   "frmReportUS.frx":0784
            Style           =   1  'Simple Combo
            TabIndex        =   161
            Tag             =   "�ɼ�����:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���ܹ� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":07B2
            Left            =   600
            List            =   "frmReportUS.frx":07CB
            Style           =   1  'Simple Combo
            TabIndex        =   160
            Tag             =   "�ھ�:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���ܹ� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":07F9
            Left            =   4320
            List            =   "frmReportUS.frx":0806
            TabIndex        =   162
            Tag             =   "���䲿λ:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cbx���ܹ� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":0818
            Left            =   600
            List            =   "frmReportUS.frx":0831
            TabIndex        =   163
            Tag             =   "ǻ:[value]"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cbx���ܹ� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":085F
            Left            =   4320
            List            =   "frmReportUS.frx":086F
            TabIndex        =   165
            Tag             =   "��Ӱ:[value]cm"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "�ɼ�����       cm"
            Height          =   255
            Index           =   35
            Left            =   1560
            TabIndex        =   373
            Top             =   285
            Width           =   1695
         End
         Begin VB.Label lable1 
            Caption         =   "�ھ�       cm"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   372
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lable1 
            Caption         =   "���䲿λ"
            Height          =   255
            Index           =   37
            Left            =   3480
            TabIndex        =   371
            Top             =   285
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "  ǻ"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   370
            Top             =   645
            Width           =   405
         End
         Begin VB.Label lable1 
            Caption         =   "��С       cm"
            Height          =   255
            Index           =   36
            Left            =   1920
            TabIndex        =   369
            Top             =   645
            Width           =   1215
         End
         Begin VB.Label lable1 
            Caption         =   "��Ӱ"
            Height          =   255
            Index           =   38
            Left            =   3840
            TabIndex        =   368
            Top             =   645
            Width           =   495
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "��  ��"
         Height          =   1695
         Index           =   11
         Left            =   120
         TabIndex        =   353
         Top             =   2280
         Width           =   5655
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   10
            ItemData        =   "frmReportUS.frx":0887
            Left            =   1680
            List            =   "frmReportUS.frx":0897
            Style           =   1  'Simple Combo
            TabIndex        =   157
            Tag             =   "��С:[value]cm"
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   7
            ItemData        =   "frmReportUS.frx":08B1
            Left            =   2880
            List            =   "frmReportUS.frx":08C1
            Style           =   1  'Simple Combo
            TabIndex        =   154
            Tag             =   "��ߴ�С:[value]cm"
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":08DB
            Left            =   4560
            List            =   "frmReportUS.frx":08EB
            Style           =   1  'Simple Combo
            TabIndex        =   152
            Tag             =   "����:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0905
            Left            =   2880
            List            =   "frmReportUS.frx":0915
            Style           =   1  'Simple Combo
            TabIndex        =   148
            Tag             =   "ǰ��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":092F
            Left            =   840
            List            =   "frmReportUS.frx":093F
            Style           =   1  'Simple Combo
            TabIndex        =   147
            Tag             =   "����:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   12
            ItemData        =   "frmReportUS.frx":0959
            Left            =   4800
            List            =   "frmReportUS.frx":0966
            TabIndex        =   159
            Tag             =   "��ɫѪ��:[value]"
            Top             =   1320
            Width           =   735
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   11
            ItemData        =   "frmReportUS.frx":0978
            Left            =   2890
            List            =   "frmReportUS.frx":0988
            TabIndex        =   158
            Tag             =   "��Ӱ:[value]"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   9
            ItemData        =   "frmReportUS.frx":09A0
            Left            =   360
            List            =   "frmReportUS.frx":09B6
            TabIndex        =   156
            Tag             =   "ǻ:[value]"
            Top             =   1320
            Width           =   855
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   8
            ItemData        =   "frmReportUS.frx":09DC
            Left            =   4560
            List            =   "frmReportUS.frx":09EC
            TabIndex        =   155
            Tag             =   "�����Ӱ:[value]"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":0A04
            Left            =   840
            List            =   "frmReportUS.frx":0A17
            TabIndex        =   153
            Tag             =   "�ұڹ��:[value]"
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0A51
            Left            =   2880
            List            =   "frmReportUS.frx":0A64
            TabIndex        =   151
            Tag             =   "��:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":0A86
            Left            =   840
            List            =   "frmReportUS.frx":0A96
            TabIndex        =   150
            Tag             =   "��������:[value]"
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0AB0
            Left            =   4560
            List            =   "frmReportUS.frx":0AC0
            TabIndex        =   149
            Tag             =   "��̬:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��ɫѪ��"
            Height          =   255
            Index           =   29
            Left            =   3960
            TabIndex        =   366
            Top             =   1365
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "��Ӱ"
            Height          =   255
            Index           =   30
            Left            =   2520
            TabIndex        =   365
            Top             =   1365
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��С      cm"
            Height          =   255
            Index           =   31
            Left            =   1320
            TabIndex        =   364
            Top             =   1365
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "ǻ"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   363
            Top             =   1365
            Width           =   285
         End
         Begin VB.Label lable1 
            Caption         =   "��Ӱ"
            Height          =   255
            Index           =   28
            Left            =   4080
            TabIndex        =   362
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��С      cm"
            Height          =   255
            Index           =   27
            Left            =   2520
            TabIndex        =   361
            Top             =   1005
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "�ұڹ��"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   360
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "����        cm"
            Height          =   255
            Index           =   25
            Left            =   4080
            TabIndex        =   359
            Top             =   645
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "��"
            Height          =   255
            Index           =   24
            Left            =   2640
            TabIndex        =   358
            Top             =   645
            Width           =   375
         End
         Begin VB.Label lable1 
            Caption         =   "��������"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   357
            Top             =   645
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "��̬"
            Height          =   255
            Index           =   22
            Left            =   4080
            TabIndex        =   356
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "����      cm"
            Height          =   255
            Index           =   20
            Left            =   480
            TabIndex        =   355
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "ǰ��      cm"
            Height          =   255
            Index           =   21
            Left            =   2280
            TabIndex        =   354
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "Ѫ  ��"
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   350
         Top             =   1680
         Width           =   5655
         Begin VB.ComboBox cbxѪ�� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0AD8
            Left            =   3840
            List            =   "frmReportUS.frx":0AE8
            Style           =   1  'Simple Combo
            TabIndex        =   146
            Tag             =   "Ƣ����:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbxѪ�� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0B02
            Left            =   1080
            List            =   "frmReportUS.frx":0B12
            Style           =   1  'Simple Combo
            TabIndex        =   145
            Tag             =   "�ž�������:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "Ƣ����       cm"
            Height          =   255
            Index           =   19
            Left            =   3240
            TabIndex        =   352
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label lable1 
            Caption         =   "�ž�������       cm"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   351
            Top             =   255
            Width           =   1815
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�������"
         Height          =   1455
         Index           =   9
         Left            =   120
         TabIndex        =   343
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0B2C
            Left            =   4560
            List            =   "frmReportUS.frx":0B3F
            Style           =   1  'Simple Combo
            TabIndex        =   136
            Tag             =   "�Ҹ�����б��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0B5F
            Left            =   2520
            List            =   "frmReportUS.frx":0B72
            Style           =   1  'Simple Combo
            TabIndex        =   135
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0B92
            Left            =   960
            List            =   "frmReportUS.frx":0BA5
            Style           =   1  'Simple Combo
            TabIndex        =   134
            Tag             =   "��γ���:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0BC5
            Left            =   2880
            List            =   "frmReportUS.frx":0BD5
            TabIndex        =   138
            Tag             =   "�иξ���:[value]"
            Top             =   660
            Width           =   1095
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����������"
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   139
            Tag             =   "����������[value]"
            Top             =   660
            Width           =   1215
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����"
            Height          =   255
            Index           =   5
            Left            =   4080
            TabIndex        =   144
            Tag             =   "[T5]����[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "������"
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   143
            Tag             =   "[T5]������[value]"
            Top             =   1050
            Width           =   855
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   142
            Tag             =   "[T5]����[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "��ǿ"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   141
            Tag             =   "[T5]��ǿ[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   140
            Tag             =   "[T5]����[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.ComboBox cbx������� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":0BF3
            Left            =   960
            List            =   "frmReportUS.frx":0C06
            TabIndex        =   137
            Tag             =   "��̬:[value]"
            Top             =   660
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "�иξ���"
            Height          =   255
            Index           =   16
            Left            =   2040
            TabIndex        =   349
            Top             =   720
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "��    ��(                                             )"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   348
            Top             =   1080
            Width           =   5295
         End
         Begin VB.Label lable1 
            Caption         =   "��    ̬"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   347
            Top             =   675
            Width           =   855
         End
         Begin VB.Label lable1 
            Caption         =   "�Ҹ�����б��      cm"
            Height          =   255
            Index           =   14
            Left            =   3480
            TabIndex        =   346
            Top             =   315
            Width           =   1815
         End
         Begin VB.Label lable1 
            Caption         =   "��       cm"
            Height          =   255
            Index           =   13
            Left            =   2040
            TabIndex        =   345
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label lable1 
            Caption         =   "��γ���       cm"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   344
            Top             =   315
            Width           =   1575
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3135
      Index           =   4
      Left            =   0
      TabIndex        =   412
      Top             =   1440
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   21
         ItemData        =   "frmReportUS.frx":0C26
         Left            =   5160
         List            =   "frmReportUS.frx":0C36
         Style           =   1  'Simple Combo
         TabIndex        =   133
         Tag             =   "[T4]A/B:[value]"
         Top             =   2760
         Width           =   390
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   20
         ItemData        =   "frmReportUS.frx":0C4E
         Left            =   4440
         List            =   "frmReportUS.frx":0C5E
         Style           =   1  'Simple Combo
         TabIndex        =   132
         Tag             =   "[T4]PI:[value]"
         Top             =   2760
         Width           =   390
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   19
         ItemData        =   "frmReportUS.frx":0C76
         Left            =   3840
         List            =   "frmReportUS.frx":0C86
         Style           =   1  'Simple Combo
         TabIndex        =   131
         Tag             =   "[T4]RI:[value]"
         Top             =   2760
         Width           =   390
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   18
         ItemData        =   "frmReportUS.frx":0C9E
         Left            =   2760
         List            =   "frmReportUS.frx":0CAE
         Style           =   1  'Simple Combo
         TabIndex        =   130
         Tag             =   "[T4]ED:[value]cm/s"
         Top             =   2760
         Width           =   390
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   17
         ItemData        =   "frmReportUS.frx":0CC6
         Left            =   1680
         List            =   "frmReportUS.frx":0CD6
         Style           =   1  'Simple Combo
         TabIndex        =   129
         Tag             =   "[T4]PS:[value]cm/s"
         Top             =   2760
         Width           =   390
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   16
         ItemData        =   "frmReportUS.frx":0CEE
         Left            =   4710
         List            =   "frmReportUS.frx":0CFE
         Style           =   1  'Simple Combo
         TabIndex        =   128
         Tag             =   "[T3]����:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   15
         ItemData        =   "frmReportUS.frx":0D16
         Left            =   3480
         List            =   "frmReportUS.frx":0D26
         Style           =   1  'Simple Combo
         TabIndex        =   127
         Tag             =   "[T3]����:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   14
         ItemData        =   "frmReportUS.frx":0D3E
         Left            =   2190
         List            =   "frmReportUS.frx":0D4E
         Style           =   1  'Simple Combo
         TabIndex        =   126
         Tag             =   "[T3]����:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   13
         ItemData        =   "frmReportUS.frx":0D66
         Left            =   960
         List            =   "frmReportUS.frx":0D76
         Style           =   1  'Simple Combo
         TabIndex        =   125
         Tag             =   "[T3]����:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   12
         ItemData        =   "frmReportUS.frx":0D8E
         Left            =   1320
         List            =   "frmReportUS.frx":0D9E
         Style           =   1  'Simple Combo
         TabIndex        =   124
         Tag             =   "̥ͷλ��:[value]cm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   9
         ItemData        =   "frmReportUS.frx":0DB6
         Left            =   2640
         List            =   "frmReportUS.frx":0DC6
         Style           =   1  'Simple Combo
         TabIndex        =   117
         Tag             =   "���:[value]cm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   5
         ItemData        =   "frmReportUS.frx":0DDE
         Left            =   4320
         List            =   "frmReportUS.frx":0DEE
         Style           =   1  'Simple Combo
         TabIndex        =   113
         Tag             =   "�ɹǳ�:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   4
         ItemData        =   "frmReportUS.frx":0E06
         Left            =   2640
         List            =   "frmReportUS.frx":0E16
         Style           =   1  'Simple Combo
         TabIndex        =   112
         Tag             =   "��Χ:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   2
         ItemData        =   "frmReportUS.frx":0E2E
         Left            =   4320
         List            =   "frmReportUS.frx":0E3E
         Style           =   1  'Simple Combo
         TabIndex        =   110
         Tag             =   "ͷΧ:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   1
         ItemData        =   "frmReportUS.frx":0E56
         Left            =   2640
         List            =   "frmReportUS.frx":0E66
         Style           =   1  'Simple Combo
         TabIndex        =   109
         Tag             =   "˫����:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "����δ��ӯ̥����Ե�۲첻��"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   123
         Tag             =   "����δ��ӯ̥����Ե�۲첻��[value]"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "������"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   122
         Tag             =   "������[value]"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "��Ե��"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   121
         Tag             =   "��Ե��[value]"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "������"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   120
         Tag             =   "������[value]"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   11
         ItemData        =   "frmReportUS.frx":0E7E
         Left            =   960
         List            =   "frmReportUS.frx":0E8B
         TabIndex        =   119
         Tag             =   "ǰ��̥��:[value]"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   10
         ItemData        =   "frmReportUS.frx":0E99
         Left            =   4320
         List            =   "frmReportUS.frx":0EB5
         TabIndex        =   118
         Tag             =   "̥�̼���:[value]"
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   8
         ItemData        =   "frmReportUS.frx":0EE6
         Left            =   960
         List            =   "frmReportUS.frx":0F05
         TabIndex        =   116
         Tag             =   "̥��λ��:[value]"
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   7
         ItemData        =   "frmReportUS.frx":0F43
         Left            =   4320
         List            =   "frmReportUS.frx":0F50
         TabIndex        =   115
         Tag             =   "����:[value]"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   6
         ItemData        =   "frmReportUS.frx":0F64
         Left            =   960
         List            =   "frmReportUS.frx":0F74
         TabIndex        =   114
         Tag             =   "̥��:[value]"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   0
         ItemData        =   "frmReportUS.frx":0F88
         Left            =   960
         List            =   "frmReportUS.frx":0F98
         TabIndex        =   108
         Tag             =   "̥ͷλ��:[value]"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cbx������� 
         Height          =   300
         Index           =   3
         ItemData        =   "frmReportUS.frx":0FB0
         Left            =   960
         List            =   "frmReportUS.frx":0FC3
         TabIndex        =   111
         Tag             =   "����λ��:[value]"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lable1 
         Caption         =   "��ˮ:����      cm  ����      cm  ����      cm  ����      cm"
         Height          =   255
         Index           =   104
         Left            =   120
         TabIndex        =   420
         Top             =   2420
         Width           =   5415
      End
      Begin VB.Label lable1 
         Caption         =   "���þ๬�ڿ�        cm"
         Height          =   255
         Index           =   103
         Left            =   120
         TabIndex        =   419
         Top             =   2080
         Width           =   2175
      End
      Begin VB.Label lable1 
         Caption         =   "�궯��Ѫ��ָ��:PS     cm/s ED     cm/S RI     PI     A/B"
         Height          =   255
         Index           =   102
         Left            =   120
         TabIndex        =   418
         Top             =   2790
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "ǰ��̥��"
         Height          =   255
         Index           =   109
         Left            =   120
         TabIndex        =   417
         Top             =   1720
         Width           =   855
      End
      Begin VB.Label lable1 
         Caption         =   "̥��λ��                ���      cm  ̥�̼���"
         Height          =   255
         Index           =   110
         Left            =   120
         TabIndex        =   416
         Top             =   1360
         Width           =   4335
      End
      Begin VB.Label lable1 
         Caption         =   "    ̥��                                  ����"
         Height          =   255
         Index           =   111
         Left            =   120
         TabIndex        =   415
         Top             =   1005
         Width           =   5295
      End
      Begin VB.Label lable1 
         Caption         =   "����λ��                ��Χ      cm    �ɹǳ�       cm"
         Height          =   255
         Index           =   112
         Left            =   120
         TabIndex        =   414
         Top             =   640
         Width           =   5175
      End
      Begin VB.Label lable1 
         Caption         =   "̥ͷλ��              ˫����      cm      ͷΧ       cm"
         Height          =   255
         Index           =   113
         Left            =   120
         TabIndex        =   413
         Top             =   270
         Width           =   5055
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3495
      Index           =   3
      Left            =   0
      TabIndex        =   401
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "̥  ��"
         Height          =   975
         Index           =   18
         Left            =   120
         TabIndex        =   409
         Top             =   2400
         Width           =   5655
         Begin VB.ComboBox cbx̥�� 
            Height          =   300
            Index           =   2
            Left            =   2760
            Style           =   1  'Simple Combo
            TabIndex        =   101
            Tag             =   "��ѿ:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx̥�� 
            Height          =   300
            Index           =   1
            Left            =   1480
            Style           =   1  'Simple Combo
            TabIndex        =   100
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx̥�� 
            Height          =   300
            Index           =   0
            Left            =   360
            Style           =   1  'Simple Combo
            TabIndex        =   99
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "��Բ"
            Height          =   255
            Index           =   6
            Left            =   2560
            TabIndex        =   106
            Tag             =   "[T2]��Բ[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "Բ��"
            Height          =   255
            Index           =   5
            Left            =   1900
            TabIndex        =   105
            Tag             =   "[T2]Բ��[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "����"
            Height          =   255
            Index           =   4
            Left            =   1230
            TabIndex        =   104
            Tag             =   "[T2]����[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�⻬"
            Height          =   255
            Index           =   3
            Left            =   580
            TabIndex        =   103
            Tag             =   "[T2]�⻬[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "��ѿδ��"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   102
            Tag             =   "��ѿδ��[value]"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cbx����ԭʼ�Ĺܲ��� 
            Height          =   300
            ItemData        =   "frmReportUS.frx":0FE1
            Left            =   4560
            List            =   "frmReportUS.frx":0FF1
            TabIndex        =   107
            Tag             =   "ԭʼ�Ĺܲ���:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��       cm  ��      cm  ��ѿ      cm"
            Height          =   255
            Index           =   101
            Left            =   120
            TabIndex        =   411
            Top             =   285
            Width           =   3495
         End
         Begin VB.Label lable1 
            Caption         =   "�ұ�                                 ԭʼ�Ĺܲ���"
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   410
            Top             =   645
            Width           =   5445
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "���ѳ�"
         Height          =   615
         Index           =   17
         Left            =   120
         TabIndex        =   407
         Top             =   1800
         Width           =   5655
         Begin VB.ComboBox cbx���ѳ� 
            Height          =   300
            Index           =   1
            Left            =   1440
            Style           =   1  'Simple Combo
            TabIndex        =   97
            Tag             =   "X [value]cm^2"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���ѳ� 
            Height          =   300
            Index           =   0
            Left            =   720
            Style           =   1  'Simple Combo
            TabIndex        =   96
            Tag             =   "�Ҳ�:[value]"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��ʾ"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   98
            Tag             =   "δ��ʾ[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "�Ҳ�         X       cm^2"
            Height          =   255
            Index           =   96
            Left            =   120
            TabIndex        =   408
            Top             =   285
            Width           =   2535
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "���ѳ�"
         Height          =   615
         Index           =   16
         Left            =   120
         TabIndex        =   405
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cbx���ѳ� 
            Height          =   300
            Index           =   1
            Left            =   1440
            Style           =   1  'Simple Combo
            TabIndex        =   94
            Tag             =   "X [value]cm^2"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx���ѳ� 
            Height          =   300
            Index           =   0
            Left            =   720
            Style           =   1  'Simple Combo
            TabIndex        =   93
            Tag             =   "���:[value]"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "δ��ʾ"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   95
            Tag             =   "δ��ʾ[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "���         X       cm^2"
            Height          =   255
            Index           =   98
            Left            =   120
            TabIndex        =   406
            Top             =   285
            Width           =   2415
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�ӹ����"
         Height          =   975
         Index           =   15
         Left            =   120
         TabIndex        =   402
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cbx�ӹ� 
            Height          =   300
            Index           =   2
            Left            =   4080
            Style           =   1  'Simple Combo
            TabIndex        =   89
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�ӹ� 
            Height          =   300
            Index           =   1
            Left            =   2160
            Style           =   1  'Simple Combo
            TabIndex        =   88
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�ӹ� 
            Height          =   300
            Index           =   0
            Left            =   600
            Style           =   1  'Simple Combo
            TabIndex        =   87
            Tag             =   "��:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx�ӹ� 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":1005
            Left            =   4080
            List            =   "frmReportUS.frx":1012
            TabIndex        =   92
            Tag             =   "��ǻ����:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx�ӹ� 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":1026
            Left            =   2160
            List            =   "frmReportUS.frx":1036
            TabIndex        =   91
            Tag             =   "λ��:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx�ӹ� 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":104E
            Left            =   600
            List            =   "frmReportUS.frx":105B
            TabIndex        =   90
            Tag             =   "��̬:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "��̬              λ��             ��ǻ����"
            Height          =   255
            Index           =   99
            Left            =   120
            TabIndex        =   404
            Top             =   645
            Width           =   4965
         End
         Begin VB.Label lable1 
            Caption         =   "  ��       cm       ��       cm          ��        cm"
            Height          =   255
            Index           =   97
            Left            =   120
            TabIndex        =   403
            Top             =   285
            Width           =   4935
         End
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4575
      Index           =   2
      Left            =   0
      TabIndex        =   388
      Top             =   960
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   32
         Left            =   3360
         Style           =   1  'Simple Combo
         TabIndex        =   85
         Tag             =   "HR:[value]��/��"
         Top             =   4200
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   31
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   84
         Tag             =   "[T1]CI:[value]L/ml/m^2"
         Top             =   4200
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   30
         Left            =   4200
         Style           =   1  'Simple Combo
         TabIndex        =   83
         Tag             =   "[T1]SI:[value]ml/m^2"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   29
         Left            =   3000
         Style           =   1  'Simple Combo
         TabIndex        =   82
         Tag             =   "[T1]SV:[value]ml"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   28
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   81
         Tag             =   "[T1]LVVD:[value]cm^3"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   27
         Left            =   3840
         Style           =   1  'Simple Combo
         TabIndex        =   80
         Tag             =   "[T1]CO:[value]L/Min"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   26
         Left            =   2760
         Style           =   1  'Simple Combo
         TabIndex        =   79
         Tag             =   "[T1]FS:[value]%"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   25
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   78
         Tag             =   "[T1]EF:[value]%"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   24
         Left            =   4920
         Style           =   1  'Simple Combo
         TabIndex        =   77
         Tag             =   "��С:[value]"
         Top             =   3120
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   22
         Left            =   480
         Style           =   1  'Simple Combo
         TabIndex        =   75
         Tag             =   "���:[value]cm^2"
         Top             =   3120
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   21
         Left            =   4920
         Style           =   1  'Simple Combo
         TabIndex        =   74
         Tag             =   "����������(����):[value]mm"
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   20
         Left            =   2400
         Style           =   1  'Simple Combo
         TabIndex        =   73
         Tag             =   "���:[value]cm^2"
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   19
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   72
         Tag             =   "����(����):[value]mm"
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   15
         Left            =   5010
         Style           =   1  'Simple Combo
         TabIndex        =   68
         Tag             =   "���Һ�ڶ���:[value]mm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   14
         Left            =   3840
         Style           =   1  'Simple Combo
         TabIndex        =   67
         Tag             =   "���Һ�ں��:[value]mm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   12
         Left            =   4560
         Style           =   1  'Simple Combo
         TabIndex        =   65
         Tag             =   "�Ҽ������:[value]mm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   11
         Left            =   3000
         Style           =   1  'Simple Combo
         TabIndex        =   64
         Tag             =   "�Ҽ�����:[value]mm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   10
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   63
         Tag             =   "����ǰ��:[value]mm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   9
         Left            =   4540
         Style           =   1  'Simple Combo
         TabIndex        =   62
         Tag             =   "���ҳ���:[value]mm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   8
         Left            =   2760
         Style           =   1  'Simple Combo
         TabIndex        =   61
         Tag             =   "���Һᾶ:[value]mm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   7
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   60
         Tag             =   "����ǰ��:[value]mm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   6
         Left            =   4540
         Style           =   1  'Simple Combo
         TabIndex        =   59
         Tag             =   "�ҷ�����:[value]mm"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   5
         Left            =   2760
         Style           =   1  'Simple Combo
         TabIndex        =   58
         Tag             =   "�ҷ��ᾶ:[value]mm"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   4
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   57
         Tag             =   "��ǰ��:[value]mm"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   3
         Left            =   4200
         Style           =   1  'Simple Combo
         TabIndex        =   56
         Tag             =   "������񼲿�ھ�:[value]mm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   2
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   55
         Tag             =   "������񼲿�ھ�:[value]mm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   1
         Left            =   4200
         Style           =   1  'Simple Combo
         TabIndex        =   54
         Tag             =   "�ζ������ھ�:[value]mm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   0
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   53
         Tag             =   "�����������ھ�:[value]mm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   33
         ItemData        =   "frmReportUS.frx":106F
         Left            =   4920
         List            =   "frmReportUS.frx":107C
         TabIndex        =   86
         Tag             =   "����:[value]"
         Top             =   4200
         Width           =   855
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   23
         ItemData        =   "frmReportUS.frx":108C
         Left            =   2040
         List            =   "frmReportUS.frx":109C
         TabIndex        =   76
         Tag             =   "������:[value]"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   13
         ItemData        =   "frmReportUS.frx":10C0
         Left            =   1800
         List            =   "frmReportUS.frx":10CD
         TabIndex        =   66
         Tag             =   "�Ҽ�������Һ�ڳ�:[value]"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   18
         ItemData        =   "frmReportUS.frx":10E7
         Left            =   4920
         List            =   "frmReportUS.frx":10F4
         TabIndex        =   71
         Tag             =   "��Ҷ��:[value]"
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   17
         ItemData        =   "frmReportUS.frx":1106
         Left            =   3360
         List            =   "frmReportUS.frx":1113
         TabIndex        =   70
         Tag             =   "�����ǰҶ����߳�:[value]"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cbx���� 
         Height          =   300
         Index           =   16
         ItemData        =   "frmReportUS.frx":1129
         Left            =   480
         List            =   "frmReportUS.frx":1139
         TabIndex        =   69
         Tag             =   "����:[value]"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lable1 
         Caption         =   "CI       L/ml/m^2  HR      ��/��  ����"
         Height          =   255
         Index           =   11
         Left            =   1440
         TabIndex        =   400
         Top             =   4230
         Width           =   4350
      End
      Begin VB.Label lable1 
         Caption         =   "LVVD      cm^3  SV       ml  SI       ml/m^2"
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   399
         Top             =   3880
         Width           =   4095
      End
      Begin VB.Label lable1 
         Caption         =   "�����������ܣ�EF       %  FS       %  CO       L/Min "
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   398
         Top             =   3530
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "���      cm^2 ������              ��ǿ�������  ��С       cm"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   397
         Top             =   3160
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "����(����)       mm  ���       cm^2 ����������(����)       mm"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   396
         Top             =   2780
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "�Ҽ�������Һ�ڳ�           ���Һ�ں��      mm ����      mm"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   395
         Top             =   2070
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "����            �����ǰҶ����߳�           ��Ҷ��"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   394
         Top             =   2430
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "����ǰ��       mm  �Ҽ�����       mm    ����       mm"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   393
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "����ǰ��       mm  ���Һᾶ       mm  ���ҳ���       mm"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   392
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "��ǰ��       mm  �ҷ��ᾶ       mm  �ҷ�����       mm"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   391
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "������񼲿�ھ�         mm       ���������ھ�        mm"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   390
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lable1 
         Caption         =   "�����������ھ�         mm       �ζ������ھ�        mm "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   389
         Top             =   270
         Width           =   5055
      End
   End
   Begin VB.ComboBox cbxType 
      Height          =   300
      ItemData        =   "frmReportUS.frx":1155
      Left            =   1200
      List            =   "frmReportUS.frx":1189
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame framType 
      Caption         =   "�����ճ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4455
      Index           =   1
      Left            =   0
      TabIndex        =   285
      Top             =   720
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "�ζ�����"
         Height          =   1935
         Index           =   8
         Left            =   3120
         TabIndex        =   315
         Top             =   2400
         Width           =   2655
         Begin VB.TextBox txt�ζ������ֵ���� 
            Height          =   270
            Left            =   1200
            TabIndex        =   52
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txt�ζ����귴��ѹ�� 
            Height          =   270
            Left            =   1200
            TabIndex        =   51
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txt�ζ�������Ѫʱ�� 
            Height          =   270
            Left            =   1200
            TabIndex        =   50
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txt�ζ��������ʱ�� 
            Height          =   270
            Left            =   1200
            TabIndex        =   49
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt�ζ�����������ѹ�� 
            Height          =   270
            Left            =   1680
            TabIndex        =   48
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt�ζ��������������� 
            Height          =   270
            Left            =   1200
            TabIndex        =   47
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt�ζ�����������ѹ�� 
            Height          =   270
            Left            =   1680
            TabIndex        =   46
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt�ζ��������������� 
            Height          =   270
            Left            =   1200
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��ֵ����          cm/s"
            Height          =   255
            Index           =   92
            Left            =   120
            TabIndex        =   340
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "����ѹ��          mmHg"
            Height          =   255
            Index           =   91
            Left            =   120
            TabIndex        =   326
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "��Ѫʱ��          ms"
            Height          =   255
            Index           =   90
            Left            =   120
            TabIndex        =   325
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "����ʱ��          ms"
            Height          =   255
            Index           =   89
            Left            =   120
            TabIndex        =   324
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "������"
            Height          =   255
            Index           =   88
            Left            =   120
            TabIndex        =   323
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������"
            Height          =   255
            Index           =   87
            Left            =   120
            TabIndex        =   316
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "��������"
         Height          =   1095
         Index           =   6
         Left            =   120
         TabIndex        =   313
         Top             =   3240
         Width           =   2775
         Begin VB.TextBox txt������������ 
            Height          =   270
            Left            =   1440
            TabIndex        =   37
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt������������������ 
            Height          =   270
            Left            =   1440
            TabIndex        =   35
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt��������������ѹ�� 
            Height          =   270
            Left            =   1920
            TabIndex        =   36
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt������������������ 
            Height          =   270
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt��������������ѹ�� 
            Height          =   270
            Left            =   1920
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "����"
            Height          =   255
            Index           =   82
            Left            =   120
            TabIndex        =   339
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������"
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   322
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������"
            Height          =   255
            Index           =   80
            Left            =   120
            TabIndex        =   314
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�����"
         Height          =   1575
         Index           =   7
         Left            =   3120
         TabIndex        =   311
         Top             =   720
         Width           =   2655
         Begin VB.TextBox txt����귴��ѹ�� 
            Height          =   270
            Left            =   1200
            TabIndex        =   44
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txt��������������� 
            Height          =   270
            Left            =   1200
            TabIndex        =   42
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt�����������ѹ�� 
            Height          =   270
            Left            =   1680
            TabIndex        =   43
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt�����������A������ 
            Height          =   270
            Left            =   1200
            TabIndex        =   40
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt�����������A��ѹ�� 
            Height          =   270
            Left            =   1680
            TabIndex        =   41
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt�����������E������ 
            Height          =   270
            Left            =   1200
            TabIndex        =   38
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt�����������E��ѹ�� 
            Height          =   270
            Left            =   1680
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "����ѹ��          mmHg"
            Height          =   255
            Index           =   86
            Left            =   120
            TabIndex        =   321
            Top             =   1110
            Width           =   2055
         End
         Begin VB.Label lable1 
            Caption         =   "������"
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   320
            Top             =   870
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������A��"
            Height          =   255
            Index           =   84
            Left            =   120
            TabIndex        =   319
            Top             =   630
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������E��"
            Height          =   255
            Index           =   83
            Left            =   120
            TabIndex        =   312
            Top             =   390
            Width           =   975
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�����"
         Height          =   2415
         Index           =   5
         Left            =   120
         TabIndex        =   309
         Top             =   720
         Width           =   2775
         Begin VB.TextBox txt�����E������ʱ�� 
            Height          =   270
            Left            =   1320
            TabIndex        =   32
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txt�����E�����ٶ� 
            Height          =   270
            Left            =   1320
            TabIndex        =   31
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txt������������ʱ�� 
            Height          =   270
            Left            =   1320
            TabIndex        =   30
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txt�����������ѹ�� 
            Height          =   270
            Left            =   1800
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt��������������� 
            Height          =   270
            Left            =   1320
            TabIndex        =   28
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt�����������A��ѹ�� 
            Height          =   270
            Left            =   1800
            TabIndex        =   27
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt�����������A������ 
            Height          =   270
            Left            =   1320
            TabIndex        =   26
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt�����������E��ѹ�� 
            Height          =   270
            Left            =   1800
            TabIndex        =   25
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt�����������E������ 
            Height          =   270
            Left            =   1320
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "E������ʱ��        196                       ��27ms"
            Height          =   375
            Index           =   79
            Left            =   120
            TabIndex        =   329
            Top             =   1920
            Width           =   2280
         End
         Begin VB.Label lable1 
            Caption         =   "E�����ٶ�          >150                     cm/s"
            Height          =   375
            Index           =   78
            Left            =   120
            TabIndex        =   328
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "��������ʱ��       <90ms"
            Height          =   255
            Index           =   77
            Left            =   120
            TabIndex        =   327
            Top             =   1110
            Width           =   2250
         End
         Begin VB.Label lable1 
            Caption         =   "������"
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   318
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������A��"
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   317
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "������E��"
            Height          =   255
            Index           =   74
            Left            =   120
            TabIndex        =   310
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label lable1 
         Caption         =   "              (cm/s) (mmHg)               (cm/s)(mmHg)"
         Height          =   255
         Index           =   95
         Left            =   120
         TabIndex        =   330
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "  ��Ŀ        ����  ѹ��              ��Ŀ    ����  ѹ��"
         Height          =   255
         Index           =   94
         Left            =   120
         TabIndex        =   308
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame framType 
      Caption         =   "��ά��M�ͳ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4455
      Index           =   0
      Left            =   0
      TabIndex        =   284
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame framGroup 
         Caption         =   "������"
         Height          =   2055
         Index           =   1
         Left            =   120
         TabIndex        =   331
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txt��������Ѫ���� 
            Height          =   270
            Left            =   1320
            TabIndex        =   11
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txt����������ĩ�ݻ� 
            Height          =   270
            Left            =   1320
            TabIndex        =   283
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txt����������ĩ�ݻ� 
            Height          =   270
            Left            =   1320
            TabIndex        =   10
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt�������˶����� 
            Height          =   270
            Left            =   1320
            TabIndex        =   9
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt�����Һ�ں�� 
            Height          =   270
            Left            =   1320
            TabIndex        =   8
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt����������ĩ�ھ� 
            Height          =   270
            Left            =   1320
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt������ÿ���� 
            Height          =   270
            Left            =   1320
            TabIndex        =   12
            Top             =   1690
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "ÿ����(ml)         35-90"
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   332
            Top             =   1710
            Width           =   2175
         End
         Begin VB.Label Label51 
            Caption         =   "����ĩ�ݻ� ml      60-125"
            Height          =   255
            Left            =   120
            TabIndex        =   338
            Top             =   990
            Width           =   2295
         End
         Begin VB.Label lable1 
            Caption         =   "��Ѫ����EF         >50%"
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   337
            Top             =   1470
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "����ĩ�ݻ� ml      30-40"
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   336
            Top             =   1230
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "�˶�����           8-15"
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   335
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "��ں��           8-12"
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   334
            Top             =   510
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "����ĩ�ھ�         <55/50"
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   333
            Top             =   270
            Width           =   2295
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "������"
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   289
         Top             =   480
         Width           =   2655
         Begin VB.TextBox txt��������Ѫ���� 
            Height          =   270
            Left            =   1320
            TabIndex        =   6
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txt����������ĩ�ݻ� 
            Height          =   270
            Left            =   1320
            TabIndex        =   5
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txt����������ĩ�ݻ� 
            Height          =   270
            Left            =   1320
            TabIndex        =   4
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt�����Һᾶ 
            Height          =   270
            Left            =   1320
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt������ǰ�ں�� 
            Height          =   270
            Left            =   1320
            TabIndex        =   2
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt������ǰ�� 
            Height          =   270
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��Ѫ����EF          >50%"
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   307
            Top             =   1470
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "����ĩ�ݻ� ml"
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   306
            Top             =   1230
            Width           =   2055
         End
         Begin VB.Label lable1 
            Caption         =   "����ĩ�ݻ� ml"
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   305
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lable1 
            Caption         =   "�ᾶ                <40"
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   304
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "ǰ�ں��            <5"
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   303
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lable1 
            Caption         =   "ǰ��              <25"
            Height          =   255
            Index           =   51
            Left            =   120
            TabIndex        =   302
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame framGroup 
         Height          =   1575
         Index           =   2
         Left            =   3000
         TabIndex        =   288
         Top             =   480
         Width           =   2775
         Begin VB.TextBox txt��ζ����ھ� 
            Height          =   270
            Left            =   1440
            TabIndex        =   17
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txt�ҷζ����ھ� 
            Height          =   270
            Left            =   1440
            TabIndex        =   16
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt���ζ����ھ� 
            Height          =   270
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt�������ھ� 
            Height          =   270
            Left            =   1440
            TabIndex        =   14
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt�����������ھ� 
            Height          =   270
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "��ζ����ھ�        <18"
            Height          =   255
            Index           =   67
            Left            =   120
            TabIndex        =   295
            Top             =   1230
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "�ҷζ����ھ�        <18"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   294
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "���ζ����ھ�        <25"
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   293
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "���������ھ�        <35"
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   292
            Top             =   510
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "�����������ھ�      <35"
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   291
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�Ҽ��"
         Height          =   1095
         Index           =   3
         Left            =   3000
         TabIndex        =   287
         Top             =   2070
         Width           =   2775
         Begin VB.TextBox txt�Ҽ�������Һ�ڷ��� 
            Height          =   270
            Left            =   1440
            TabIndex        =   20
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt�Ҽ���˶����� 
            Height          =   270
            Left            =   1440
            TabIndex        =   19
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt�Ҽ����� 
            Height          =   270
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "�����Һ�ڷ���       ����"
            Height          =   255
            Index           =   70
            Left            =   120
            TabIndex        =   301
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lable1 
            Caption         =   "�˶�����             5-8"
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   300
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "���                 8-12"
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   299
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "�ķ�"
         Height          =   1170
         Index           =   4
         Left            =   3000
         TabIndex        =   286
         Top             =   3170
         Width           =   2775
         Begin VB.TextBox txt�ķ��ҷ��ᾶ 
            Height          =   270
            Left            =   1440
            TabIndex        =   23
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt�ķ��ҷ����� 
            Height          =   270
            Left            =   1440
            TabIndex        =   22
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt�ķ���ǰ�� 
            Height          =   270
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "�ҷ��ᾶ            <40"
            Height          =   255
            Index           =   73
            Left            =   120
            TabIndex        =   298
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "�ҷ�����            <50"
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   297
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "��ǰ��          <37"
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   296
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label lable1 
         Caption         =   "  ��Ŀ      ��ֵ  ����(mm)         ��Ŀ      ��ֵ ����(mm)"
         Height          =   255
         Index           =   93
         Left            =   240
         TabIndex        =   290
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Label Label21 
      Caption         =   "ר����Ŀ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   341
      Top             =   140
      Width           =   1215
   End
End
Attribute VB_Name = "frmReportUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSingleWindow As Boolean     '�Ƿ�ʹ�ö���������ʾ����༭����True-����������ʾ��False-Ƕ��ʽ��ʾ
Private mlngAdviceID As Long    'ҽ��ID
Private mintEditType As Integer '����״̬ 0 ������1��д��2 �޶�
Private mReportID As Long       '����ID
Private mblnCheckModify As Boolean      '�Ƿ��������ݱ仯��¼
Private mblnEditable As Boolean         '�Ƿ���Ա༭����
Private mblnMoved As Boolean            '�Ƿ��Ѿ�ת��

Private Const Report_Element_������ǰ�� = "������ǰ��"
Private Const Report_Element_������ǰ�ں�� = "������ǰ�ں��"
Private Const Report_Element_�����Һᾶ = "�����Һᾶ"
Private Const Report_Element_����������ĩ�ݻ� = "����������ĩ�ݻ�"
Private Const Report_Element_����������ĩ�ݻ� = "����������ĩ�ݻ�"
Private Const Report_Element_��������Ѫ���� = "��������Ѫ����"

Private Const Report_Element_����������ĩ�ھ� = "����������ĩ�ھ�"
Private Const Report_Element_�����Һ�ں�� = "�����Һ�ں��"
Private Const Report_Element_�������˶����� = "�������˶�����"
Private Const Report_Element_����������ĩ�ݻ� = "����������ĩ�ݻ�"
Private Const Report_Element_����������ĩ�ݻ� = "����������ĩ�ݻ�"
Private Const Report_Element_��������Ѫ���� = "��������Ѫ����"
Private Const Report_Element_������ÿ���� = "������ÿ����"

Private Const Report_Element_�����������ھ� = "�����������ھ�"
Private Const Report_Element_�������ھ� = "���������ھ�"
Private Const Report_Element_���ζ����ھ� = "���ζ����ھ�"
Private Const Report_Element_�ҷζ����ھ� = "�ҷζ����ھ�"
Private Const Report_Element_��ζ����ھ� = "��ζ����ھ�"

Private Const Report_Element_�Ҽ����� = "�Ҽ�����"
Private Const Report_Element_�Ҽ���˶����� = "�Ҽ���˶�����"
Private Const Report_Element_�Ҽ�������Һ�ڷ��� = "�Ҽ�������Һ�ڷ���"

Private Const Report_Element_�ķ���ǰ�� = "�ķ���ǰ��"
Private Const Report_Element_�ķ��ҷ����� = "�ķ��ҷ�����"
Private Const Report_Element_�ķ��ҷ��ᾶ = "�ķ��ҷ��ᾶ"


Private Const Report_Element_�����������E������ = "�����������E������"
Private Const Report_Element_�����������E��ѹ�� = "�����������E��ѹ��"
Private Const Report_Element_�����������A������ = "�����������A������"
Private Const Report_Element_�����������A��ѹ�� = "�����������A��ѹ��"
Private Const Report_Element_��������������� = "���������������"
Private Const Report_Element_�����������ѹ�� = "�����������ѹ��"
Private Const Report_Element_������������ʱ�� = "������������ʱ��"
Private Const Report_Element_�����E�����ٶ� = "�����E�����ٶ�"
Private Const Report_Element_�����E������ʱ�� = "�����E������ʱ��"

Private Const Report_Element_������������������ = "������������������"
Private Const Report_Element_��������������ѹ�� = "��������������ѹ��"
Private Const Report_Element_������������������ = "������������������"
Private Const Report_Element_��������������ѹ�� = "��������������ѹ��"
Private Const Report_Element_������������ = "������������"

Private Const Report_Element_�����������E������ = "�����������E������"
Private Const Report_Element_�����������E��ѹ�� = "�����������E��ѹ��"
Private Const Report_Element_�����������A������ = "�����������A������"
Private Const Report_Element_�����������A��ѹ�� = "�����������A��ѹ��"
Private Const Report_Element_��������������� = "���������������"
Private Const Report_Element_�����������ѹ�� = "�����������ѹ��"
Private Const Report_Element_����귴��ѹ�� = "����귴��ѹ��"

Private Const Report_Element_�ζ��������������� = "�ζ���������������"
Private Const Report_Element_�ζ�����������ѹ�� = "�ζ�����������ѹ��"
Private Const Report_Element_�ζ��������������� = "�ζ���������������"
Private Const Report_Element_�ζ�����������ѹ�� = "�ζ�����������ѹ��"
Private Const Report_Element_�ζ��������ʱ�� = "�ζ��������ʱ��"
Private Const Report_Element_�ζ�������Ѫʱ�� = "�ζ�������Ѫʱ��"
Private Const Report_Element_�ζ����귴��ѹ�� = "�ζ����귴��ѹ��"
Private Const Report_Element_�ζ������ֵ���� = "�ζ������ֵ����"

Private Const Report_Element_ר�Ʊ��� = "ר�Ʊ���"
Private Const Report_CheckedValue = " "
Private Const Report_ProjectSplitChr = vbCrLf '"      "




Public pModified As Boolean     '��¼�Ƿ����޸�




Public Sub zlRefresh(frmParentReport As frmReport, ByVal lngAdviceID As Long, ReportID As Long, _
    blnSingleWindow As Boolean, blnEditable As Boolean, ByVal blnMoved As Boolean)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    mlngAdviceID = lngAdviceID
    mReportID = ReportID
    mblnSingleWindow = blnSingleWindow
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    
    mblnCheckModify = False         '�ر������޸ļ�¼
    '����޸ı��
    pModified = False
    
    '�����������
    txt������ǰ��.Text = ""
    txt������ǰ�ں��.Text = ""
    txt�����Һᾶ.Text = ""
    txt����������ĩ�ݻ�.Text = ""
    txt����������ĩ�ݻ�.Text = ""
    txt��������Ѫ����.Text = ""
    
    txt����������ĩ�ھ�.Text = ""
    txt�����Һ�ں��.Text = ""
    txt�������˶�����.Text = ""
    txt����������ĩ�ݻ�.Text = ""
    txt����������ĩ�ݻ�.Text = ""
    txt��������Ѫ����.Text = ""
    txt������ÿ����.Text = ""
    
    txt�����������ھ�.Text = ""
    txt�������ھ�.Text = ""
    txt���ζ����ھ�.Text = ""
    txt�ҷζ����ھ�.Text = ""
    txt��ζ����ھ�.Text = ""
    
    txt�Ҽ�����.Text = ""
    txt�Ҽ���˶�����.Text = ""
    txt�Ҽ�������Һ�ڷ���.Text = ""
    
    txt�ķ���ǰ��.Text = ""
    txt�ķ��ҷ�����.Text = ""
    txt�ķ��ҷ��ᾶ.Text = ""
    
    txt�����������E������.Text = ""
    txt�����������E��ѹ��.Text = ""
    txt�����������A������.Text = ""
    txt�����������A��ѹ��.Text = ""
    txt���������������.Text = ""
    txt�����������ѹ��.Text = ""
    txt������������ʱ��.Text = ""
    txt�����E�����ٶ�.Text = ""
    txt�����E������ʱ��.Text = ""
    
    txt������������������.Text = ""
    txt��������������ѹ��.Text = ""
    txt������������������.Text = ""
    txt��������������ѹ��.Text = ""
    txt������������.Text = ""
    
    txt�����������E������.Text = ""
    txt�����������E��ѹ��.Text = ""
    txt�����������A������.Text = ""
    txt�����������A��ѹ��.Text = ""
    txt���������������.Text = ""
    txt�����������ѹ��.Text = ""
    txt����귴��ѹ��.Text = ""
    
    txt�ζ���������������.Text = ""
    txt�ζ�����������ѹ��.Text = ""
    txt�ζ���������������.Text = ""
    txt�ζ�����������ѹ��.Text = ""
    txt�ζ��������ʱ��.Text = ""
    txt�ζ�������Ѫʱ��.Text = ""
    txt�ζ����귴��ѹ��.Text = ""
    txt�ζ������ֵ����.Text = ""
    
    '�����������
    For i = 0 To 33: cbx����(i).Text = "": Next i
    
    '��ո�������
    For i = 0 To 5: cbx�ӹ�(i).Text = "": Next i
    
    cbx���ѳ�(0).Text = ""
    cbx���ѳ�(1).Text = ""
    cbx���ѳ�(0).Text = ""
    cbx���ѳ�(1).Text = ""
    
    cbx����ԭʼ�Ĺܲ���.Text = ""
    
    For i = 0 To 6: chk����(i).value = 0: Next i
    
    '��ղ�������
    For i = 0 To 21: cbx�������(i).Text = "": Next i
    For i = 0 To 3: chk�������(i).value = 0: Next i
    
    '��ո�������
    For i = 0 To 4: cbx�������(i).Text = "": Next i
    For i = 0 To 5: chk�������(i).value = 0: Next i
    For i = 0 To 1: cbxѪ��(i).Text = "": Next i
    For i = 0 To 12: cbx�������(i).Text = "": Next i
    For i = 0 To 5: cbx���ܹ�(i).Text = "": Next i
    For i = 0 To 6: cbx����(i).Text = "": Next i
    For i = 0 To 4: cbxƢ��(i).Text = "": Next i
    
    '�������������
    For i = 0 To 5: txt�������(i).Text = "": Next i
    For i = 0 To 5: cbx������(i).Text = "":  Next i
    
    cbx����(0).Text = ""
    chk������(0).value = 0: chk������(1).value = 0
    
    For i = 0 To 2: txtǰ������Ϣ(i).Text = "": Next i
    
    '�������Ů����
    For i = 0 To 5: txtŮ�������(i).Text = "": Next i
    For i = 0 To 5: cbx�����(i).Text = "":  Next i
    
    cbxŮ����(0).Text = ""
    chk����Ů(0).value = 0: chk����Ů(1).value = 0
    
    '��վ�������
    For i = 0 To 13: txt������Ϣ(i).Text = "":  Next i
    
    '�����������
    For i = 0 To 3: txt������Ϣ(i).Text = "":  Next i
    
    '�����ֳ������
    For i = 0 To 7: cbx�����ֳ��(i).Text = "": Next i
    For i = 0 To 7: cbx�Ҳ���ֳ��(i).Text = "": Next i
    
    '�����֫��������
    For i = 0 To 4: chk����֫����(i).value = 0: Next i
    For i = 0 To 4: chk����֫����(i).value = 0: Next i
    
    '����۲�����
    For i = 0 To 1: txt�۲���Ϣ(i).Text = "": Next i
    
    '��ռ�״������
    For i = 0 To 9: txt��״����Ϣ(i).Text = "": Next i
    
    '�������������
    For i = 0 To 11: txt��������Ϣ(i).Text = "": Next i
    
    '�����ǻ����
    For i = 0 To 5: cbx��ǻ��Ϣ(i).Text = "": Next i
    For i = 0 To 1: chk��ǻ��Ϣ(i).value = 0: Next i
    
    
    strSql = "Select �����ı�,Ҫ������ From ���Ӳ������� Where �ļ�ID=[1] And ��������=4 And ��ֹ��=0 And �滻��=0"
    If mblnMoved = True Then
        strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
    
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!Ҫ������)
            Case Report_Element_������ǰ��
                txt������ǰ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_������ǰ�ں��
                txt������ǰ�ں��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����Һᾶ
                txt�����Һᾶ.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_����������ĩ�ݻ�
                txt����������ĩ�ݻ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_����������ĩ�ݻ�
                txt����������ĩ�ݻ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��������Ѫ����
                txt��������Ѫ����.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_����������ĩ�ھ�
                txt����������ĩ�ھ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����Һ�ں��
                txt�����Һ�ں��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�������˶�����
                txt�������˶�����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_����������ĩ�ݻ�
                txt����������ĩ�ݻ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_����������ĩ�ݻ�
                txt����������ĩ�ݻ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��������Ѫ����
                txt��������Ѫ����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_������ÿ����
                txt������ÿ����.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_�����������ھ�
                txt�����������ھ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�������ھ�
                txt�������ھ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_���ζ����ھ�
                txt���ζ����ھ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ҷζ����ھ�
                txt�ҷζ����ھ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��ζ����ھ�
                txt��ζ����ھ�.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_�Ҽ�����
                txt�Ҽ�����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�Ҽ���˶�����
                txt�Ҽ���˶�����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�Ҽ�������Һ�ڷ���
                txt�Ҽ�������Һ�ڷ���.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_�ķ���ǰ��
                txt�ķ���ǰ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ķ��ҷ�����
                txt�ķ��ҷ�����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ķ��ҷ��ᾶ
                txt�ķ��ҷ��ᾶ.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_�����������E������
                txt�����������E������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������E��ѹ��
                txt�����������E��ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������A������
                txt�����������A������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������A��ѹ��
                txt�����������A��ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_���������������
                txt���������������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������ѹ��
                txt�����������ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_������������ʱ��
                txt������������ʱ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����E�����ٶ�
                txt�����E�����ٶ�.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����E������ʱ��
                txt�����E������ʱ��.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_������������������
                txt������������������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��������������ѹ��
                txt��������������ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_������������������
                txt������������������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��������������ѹ��
                txt��������������ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_������������
                txt������������.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_�����������E������
                txt�����������E������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������E��ѹ��
                txt�����������E��ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������A������
                txt�����������A������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������A��ѹ��
                txt�����������A��ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_���������������
                txt���������������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�����������ѹ��
                txt�����������ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_����귴��ѹ��
                txt����귴��ѹ��.Text = Nvl(rsTemp!�����ı�)
                
            Case Report_Element_�ζ���������������
                txt�ζ���������������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ�����������ѹ��
                txt�ζ�����������ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ���������������
                txt�ζ���������������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ�����������ѹ��
                txt�ζ�����������ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ��������ʱ��
                txt�ζ��������ʱ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ�������Ѫʱ��
                txt�ζ�������Ѫʱ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ����귴��ѹ��
                txt�ζ����귴��ѹ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�ζ������ֵ����
                txt�ζ������ֵ����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_ר�Ʊ���
                '��ȡר�Ʊ����е�����
                Call LoadProfessionalReportData(Nvl(rsTemp!�����ı�))
        End Select
        rsTemp.MoveNext
    Wend
    
    '���ý���ؼ��Ƿ���Ա༭
    If cbxType.Text = "" Then
        cbxType.ListIndex = 0
        framType(0).Visible = True
    End If
    
    Call ConfigEditable(mblnEditable)
    
    mblnCheckModify = True         '�������޸ļ�¼
End Sub


Private Sub LoadProfessionalReportData(ByVal strProfessionalReport As String)
    Dim i As Long
    Dim strCurReport As String
    
    If strProfessionalReport = "" Then Exit Sub
    
    strCurReport = strProfessionalReport
    
    Select Case True
        Case InStr(strProfessionalReport, "�����ࡿ") >= 1
            For i = 0 To 33
                cbx����(i).Text = GetReportValue(strCurReport, "����", cbx����(i).Tag)
            Next i
            
            cbxType.ListIndex = 2
            framType(2).Visible = True
        Case InStr(strProfessionalReport, "���ӹ������") >= 1 Or InStr(strProfessionalReport, "�����ѳ���") >= 1 _
            Or InStr(strProfessionalReport, "�����ѳ���") >= 1 Or InStr(strProfessionalReport, "��̥�ҡ�") >= 1
            '�ӹ����
            For i = 0 To 5
                cbx�ӹ�(i).Text = GetReportValue(strCurReport, "�ӹ����", cbx�ӹ�(i).Tag)
            Next i
            
            '���ѳ�
            cbx���ѳ�(0).Text = GetReportValue(strCurReport, "���ѳ�", cbx���ѳ�(0).Tag)
            cbx���ѳ�(1).Text = GetReportValue(strCurReport, "���ѳ�", cbx���ѳ�(1).Tag)
            chk����(0).value = IIf(GetReportValue(strCurReport, "���ѳ�", chk����(0).Tag) = Report_CheckedValue, 1, 0)
            
            '���ѳ�
            cbx���ѳ�(0).Text = GetReportValue(strCurReport, "���ѳ�", cbx���ѳ�(0).Tag)
            cbx���ѳ�(1).Text = GetReportValue(strCurReport, "���ѳ�", cbx���ѳ�(1).Tag)
            chk����(1).value = IIf(GetReportValue(strCurReport, "���ѳ�", chk����(1).Tag) = Report_CheckedValue, 1, 0)
            
            '̥��
            For i = 0 To 2
                cbx̥��(i).Text = GetReportValue(strCurReport, "̥��", cbx̥��(i).Tag)
            Next i
            
            For i = 2 To 6
                chk����(i).value = IIf(GetReportValue(strCurReport, "̥��", chk����(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbx����ԭʼ�Ĺܲ���.Text = GetReportValue(strCurReport, "̥��", cbx����ԭʼ�Ĺܲ���.Tag)
            
            cbxType.ListIndex = 3
            framType(3).Visible = True
        Case InStr(strProfessionalReport, "�����������") >= 1
            For i = 0 To 21
                cbx�������(i).Text = GetReportValue(strCurReport, "�������", cbx�������(i).Tag)
            Next i
            
            For i = 0 To 3
                chk�������(i).value = IIf(GetReportValue(strCurReport, "�������", chk�������(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbxType.ListIndex = 4
            framType(4).Visible = True
        Case InStr(strProfessionalReport, "�����������") >= 1 Or InStr(strProfessionalReport, "��Ѫ�ܡ�") >= 1 _
            Or InStr(strProfessionalReport, "�����������") >= 1 Or InStr(strProfessionalReport, "�����ܹܡ�") >= 1 _
            Or InStr(strProfessionalReport, "�����١�") >= 1 Or InStr(strProfessionalReport, "��Ƣ�ࡿ") >= 1
            
            '�������
            For i = 0 To 4
                cbx�������(i).Text = GetReportValue(strCurReport, "�������", cbx�������(i).Tag)
            Next i
            
            For i = 0 To 5
                chk�������(i).value = IIf(GetReportValue(strCurReport, "�������", chk�������(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            'Ѫ��
            For i = 0 To 1
                cbxѪ��(i).Text = GetReportValue(strCurReport, "Ѫ��", cbxѪ��(i).Tag)
            Next i
            
            '�������
            For i = 0 To 12
                cbx�������(i).Text = GetReportValue(strCurReport, "����", cbx�������(i).Tag)
            Next i
            
            '���ܹ�
            For i = 0 To 5
                cbx���ܹ�(i).Text = GetReportValue(strCurReport, "���ܹ�", cbx���ܹ�(i).Tag)
            Next i
            
            '����
            For i = 0 To 6
                cbx����(i).Text = GetReportValue(strCurReport, "����", cbx����(i).Tag)
            Next i
            
            'Ƣ��
            For i = 0 To 4
                cbxƢ��(i).Text = GetReportValue(strCurReport, "Ƣ��", cbxƢ��(i).Tag)
            Next i
            
            cbxType.ListIndex = 5
            framType(5).Visible = True
        Case InStr(strProfessionalReport, "�����������") >= 1 Or InStr(strProfessionalReport, "������ܡ�") >= 1 _
            Or InStr(strProfessionalReport, "�����ס�") >= 1 Or InStr(strProfessionalReport, "��ǰ���١�") >= 1
            
            '�������
            For i = 0 To 5
                txt�������(i).Text = GetReportValue(strCurReport, "�������", txt�������(i).Tag)
            Next i
            
            '�����
            For i = 0 To 5
                cbx������(i).Text = GetReportValue(strCurReport, "�����", cbx������(i).Tag)
            Next i
            
            '����
            cbx����(0).Text = GetReportValue(strCurReport, "����", cbx����(0).Tag)
            
            For i = 0 To 1
                chk������(i).value = IIf(GetReportValue(strCurReport, "����", chk������(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            'ǰ����
            For i = 0 To 2
                txtǰ������Ϣ(i).Text = GetReportValue(strCurReport, "ǰ����", txtǰ������Ϣ(i).Tag)
            Next i
            
            cbxType.ListIndex = 6
            framType(6).Visible = True
        Case InStr(strProfessionalReport, "���� �� �� ����") >= 1 Or InStr(strProfessionalReport, "���� �� �ܡ�") >= 1 _
            Or InStr(strProfessionalReport, "���� �ס�") >= 1
            
            '�������
            For i = 0 To 5
                txtŮ�������(i).Text = GetReportValue(strCurReport, "�� �� �� ��", txtŮ�������(i).Tag)
            Next i
            
            '�����
            For i = 0 To 5
                cbx�����(i).Text = GetReportValue(strCurReport, "�� �� ��", cbx�����(i).Tag)
            Next i
            
            '����
            cbxŮ����(0).Text = GetReportValue(strCurReport, "�� ��", cbxŮ����(0).Tag)
            
            For i = 0 To 1
                chk����Ů(i).value = IIf(GetReportValue(strCurReport, "�� ��", chk����Ů(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbxType.ListIndex = 7
            framType(7).Visible = True
            
        Case InStr(strProfessionalReport, "��������") >= 1
            For i = 0 To 13
                txt������Ϣ(i).Text = GetReportValue(strCurReport, "����", txt������Ϣ(i).Tag)
            Next i
            
            cbxType.ListIndex = 8
            framType(8).Visible = True
            
        Case InStr(strProfessionalReport, "��������١�") >= 1 Or InStr(strProfessionalReport, "���Ҳ����١�") >= 1
            For i = 0 To 1
                txt������Ϣ(i).Text = GetReportValue(strCurReport, "�������", txt������Ϣ(i).Tag)
            Next i
            
            For i = 2 To 3
                txt������Ϣ(i).Text = GetReportValue(strCurReport, "�Ҳ�����", txt������Ϣ(i).Tag)
            Next i
            
            cbxType.ListIndex = 9
            framType(9).Visible = True
        Case InStr(strProfessionalReport, "�������ֳ����") >= 1 Or InStr(strProfessionalReport, "���Ҳ���ֳ����") >= 1
            For i = 0 To 7
                cbx�����ֳ��(i).Text = GetReportValue(strCurReport, "�����ֳ��", cbx�����ֳ��(i).Tag)
            Next i
            
            For i = 0 To 7
                cbx�Ҳ���ֳ��(i).Text = GetReportValue(strCurReport, "�Ҳ���ֳ��", cbx�Ҳ���ֳ��(i).Tag)
            Next i
            
            cbxType.ListIndex = 10
            framType(10).Visible = True
        Case InStr(strProfessionalReport, "������֫������") >= 1 Or InStr(strProfessionalReport, "������֫������") >= 1
            For i = 0 To 4
                chk����֫����(i).value = IIf(GetReportValue(strCurReport, "����֫����", chk����֫����(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            For i = 0 To 4
                chk����֫����(i).value = IIf(GetReportValue(strCurReport, "����֫����", chk����֫����(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbxType.ListIndex = 11
            framType(11).Visible = True
            
        Case InStr(strProfessionalReport, "���۲���") >= 1
            For i = 0 To 1
                txt�۲���Ϣ(i).Text = GetReportValue(strCurReport, "�۲�", txt�۲���Ϣ(i).Tag)
            Next i
        
            cbxType.ListIndex = 12
            framType(12).Visible = True
            
        Case InStr(strProfessionalReport, "����״�١�") >= 1
            For i = 0 To 9
                txt��״����Ϣ(i).Text = GetReportValue(strCurReport, "��״��", txt��״����Ϣ(i).Tag)
            Next i
        
            cbxType.ListIndex = 13
            framType(13).Visible = True
            
        Case InStr(strProfessionalReport, "������������") >= 1 Or InStr(strProfessionalReport, "������������") >= 1
            For i = 0 To 5
                txt��������Ϣ(i).Text = GetReportValue(strCurReport, "��������", txt��������Ϣ(i).Tag)
            Next i
            
            For i = 6 To 11
                txt��������Ϣ(i).Text = GetReportValue(strCurReport, "��������", txt��������Ϣ(i).Tag)
            Next i
        
            cbxType.ListIndex = 14
            framType(14).Visible = True
            
        Case InStr(strProfessionalReport, "����ǻ��") >= 1
            For i = 0 To 5
                cbx��ǻ��Ϣ(i).Text = GetReportValue(strCurReport, "��ǻ", cbx��ǻ��Ϣ(i).Tag)
            Next i
            
            For i = 0 To 1
                chk��ǻ��Ϣ(i).value = IIf(GetReportValue(strCurReport, "��ǻ", chk��ǻ��Ϣ(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
        
            cbxType.ListIndex = 15
            framType(15).Visible = True
        
    End Select
End Sub

Private Sub ConfigEditable(ByVal blnEditable As Boolean)
    Dim i As Long
    
    For i = 0 To 15
        framGroup(i).Enabled = blnEditable
    Next i
End Sub

Public Function getElementString() As String
    Dim strElements As String
    Dim i As Long
    
    
'0-��ά��M�ͳ���
'1-�����ճ���
'2-��    ��
'3-��    ��
'4-��    ��
'5-��    ��
'6-����(��)
'7-����(Ů)
'8-��    ��
'9-��    ��
'10-�� ֳ ��
'11-��֫����
'12-��    ��
'13-�� ״ ��
'14-�� �� ��
'15-��    ǻ
    
    Select Case Val(cbxType.Text)
        Case 0, 1
            strElements = SPLITER_REPORT & Report_Element_������ǰ�� & SPLITER_ELEMENT & txt������ǰ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_������ǰ�ں�� & SPLITER_ELEMENT & txt������ǰ�ں��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����Һᾶ & SPLITER_ELEMENT & txt�����Һᾶ.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_����������ĩ�ݻ� & SPLITER_ELEMENT & txt����������ĩ�ݻ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_����������ĩ�ݻ� & SPLITER_ELEMENT & txt����������ĩ�ݻ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��������Ѫ���� & SPLITER_ELEMENT & txt��������Ѫ����.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_����������ĩ�ھ� & SPLITER_ELEMENT & txt����������ĩ�ھ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����Һ�ں�� & SPLITER_ELEMENT & txt�����Һ�ں��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�������˶����� & SPLITER_ELEMENT & txt�������˶�����.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_����������ĩ�ݻ� & SPLITER_ELEMENT & txt����������ĩ�ݻ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_����������ĩ�ݻ� & SPLITER_ELEMENT & txt����������ĩ�ݻ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��������Ѫ���� & SPLITER_ELEMENT & txt��������Ѫ����.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_������ÿ���� & SPLITER_ELEMENT & txt������ÿ����.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_�����������ھ� & SPLITER_ELEMENT & txt�����������ھ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�������ھ� & SPLITER_ELEMENT & txt�������ھ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_���ζ����ھ� & SPLITER_ELEMENT & txt���ζ����ھ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ҷζ����ھ� & SPLITER_ELEMENT & txt�ҷζ����ھ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��ζ����ھ� & SPLITER_ELEMENT & txt��ζ����ھ�.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_�Ҽ����� & SPLITER_ELEMENT & txt�Ҽ�����.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�Ҽ���˶����� & SPLITER_ELEMENT & txt�Ҽ���˶�����.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�Ҽ�������Һ�ڷ��� & SPLITER_ELEMENT & txt�Ҽ�������Һ�ڷ���.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_�ķ���ǰ�� & SPLITER_ELEMENT & txt�ķ���ǰ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ķ��ҷ����� & SPLITER_ELEMENT & txt�ķ��ҷ�����.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ķ��ҷ��ᾶ & SPLITER_ELEMENT & txt�ķ��ҷ��ᾶ.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_�����������E������ & SPLITER_ELEMENT & txt�����������E������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������E��ѹ�� & SPLITER_ELEMENT & txt�����������E��ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������A������ & SPLITER_ELEMENT & txt�����������A������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������A��ѹ�� & SPLITER_ELEMENT & txt�����������A��ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��������������� & SPLITER_ELEMENT & txt���������������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������ѹ�� & SPLITER_ELEMENT & txt�����������ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_������������ʱ�� & SPLITER_ELEMENT & txt������������ʱ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����E�����ٶ� & SPLITER_ELEMENT & txt�����E�����ٶ�.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����E������ʱ�� & SPLITER_ELEMENT & txt�����E������ʱ��.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_������������������ & SPLITER_ELEMENT & txt������������������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��������������ѹ�� & SPLITER_ELEMENT & txt��������������ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_������������������ & SPLITER_ELEMENT & txt������������������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��������������ѹ�� & SPLITER_ELEMENT & txt��������������ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_������������ & SPLITER_ELEMENT & txt������������.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_�����������E������ & SPLITER_ELEMENT & txt�����������E������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������E��ѹ�� & SPLITER_ELEMENT & txt�����������E��ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������A������ & SPLITER_ELEMENT & txt�����������A������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������A��ѹ�� & SPLITER_ELEMENT & txt�����������A��ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_��������������� & SPLITER_ELEMENT & txt���������������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�����������ѹ�� & SPLITER_ELEMENT & txt�����������ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_����귴��ѹ�� & SPLITER_ELEMENT & txt����귴��ѹ��.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_�ζ��������������� & SPLITER_ELEMENT & txt�ζ���������������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ�����������ѹ�� & SPLITER_ELEMENT & txt�ζ�����������ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ��������������� & SPLITER_ELEMENT & txt�ζ���������������.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ�����������ѹ�� & SPLITER_ELEMENT & txt�ζ�����������ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ��������ʱ�� & SPLITER_ELEMENT & txt�ζ��������ʱ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ�������Ѫʱ�� & SPLITER_ELEMENT & txt�ζ�������Ѫʱ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ����귴��ѹ�� & SPLITER_ELEMENT & txt�ζ����귴��ѹ��.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_�ζ������ֵ���� & SPLITER_ELEMENT & txt�ζ������ֵ����.Text
        Case 2  '����ר��
            strElements = GetXinZangReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 3  '����ר��
            strElements = GetFuKeReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 4  '����ר��
            strElements = GetChanKeReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 5  '����ר��
            strElements = GetFuBuReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 6  '����(��)
            strElements = GetMiNiaoNanReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 7  '����(Ů)
            strElements = GetMiNiaoNvReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 8  '����
            strElements = GetJingBuReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 9  '����
            strElements = GetRuXianReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 10 '��ֳ��
            strElements = GetShengZhiQiReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 11 '��֫����
            strElements = GetXiaZhiJingMaiReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 12 '�۲�
            strElements = GetYanBuReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 13 '��״��
            strElements = GetJiaZhuangXianReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 14 '������
            strElements = GetShenDongMaiReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
        Case 15 '��ǻ
            strElements = GetXiongQiangReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_ר�Ʊ��� & SPLITER_ELEMENT & strElements
    End Select
         
    getElementString = strElements
End Function


Private Function GetReportValue(ByRef strReport As String, ByVal strSection As String, ByVal strTag As String) As String
    Dim aryFind() As String
    Dim strTemp As String
    Dim strFindTag As String
    Dim strSource As String
    Dim lngTemp As Long
    Dim strValue As String
    
    GetReportValue = ""
    
    If strReport = "" Or strTag = "" Then Exit Function
    If InStr(strTag, "[value]") <= 0 Then Exit Function
    
    strFindTag = Replace(strTag, "[T1]", "")
    strFindTag = Replace(strFindTag, "[T2]", "")
    strFindTag = Replace(strFindTag, "[T3]", "")
    strFindTag = Replace(strFindTag, "[T4]", "")
    strFindTag = Replace(strFindTag, "[T5]", "")
    strFindTag = Replace(strFindTag, "[T6]", "")
    strFindTag = Replace(strFindTag, "[T7]", "")
    
    aryFind = Split(strFindTag, "[value]")
    
    lngTemp = InStr(strReport, "��" & strSection & "��")
    
    '��ȡ���ݶ���"�����ѳ���  ��:12 X 20 cm^2 δ��ʾ �����ѳ���  ��13 X 22 cm^2
    If lngTemp <= 0 Then Exit Function
    
    strSource = Mid(strReport, lngTemp + Len("��" & strSection & "��"), 1000)
    lngTemp = InStr(strSource, "      ��")
    If lngTemp >= 1 Then strSource = Mid(strSource, 1, lngTemp - 1)
    
    
    '����ָ��Tag��Ӧ������ֵ
    If InStr(strSource, aryFind(0)) <= 0 Then Exit Function
    
    If strSource <> "" And strSource <> strReport Then strReport = Replace(strReport, strSource, "<@>")
    
    strValue = ""
    strTemp = Mid(strSource, InStr(strSource, aryFind(0)) + Len(aryFind(0)), 100) & "  "
    
    If UBound(aryFind) = 1 Then
        If Len(aryFind(1)) >= 1 Then
            strValue = Mid(strTemp, 1, InStr(strTemp, aryFind(1)) - 1)
        Else
            strValue = Mid(strTemp, 1, InStr(strTemp, "  ") - 1)
            If strValue = "" Then
                strValue = IIf(Mid(strTemp, InStr(strTemp, aryFind(0) & "  ") + 3, 1) = " ", " ", "")
            End If
        End If
    Else
        strValue = Mid(strTemp, 1, InStr(strTemp, "  ") - 1)
        If strValue = "" Then
            strValue = IIf(Mid(strTemp, InStr(strTemp, "  ") + 3, 1) = " ", " ", "")
        End If
    End If
    
    strFindTag = Replace(strFindTag, "[value]", strValue)
    strSource = Mid(strSource, 1, InStr(strSource, strFindTag) - 1) & Mid(strSource, InStr(strSource, strFindTag) + Len(strFindTag), 1000)
    
    strReport = Replace(strReport, "<@>", strSource)
    
'    If strTemp = "" Then strTemp = " "
    
    GetReportValue = strValue
End Function


'��ȡ���౨�棨��Ҫ���Զ��屨������֧�ֻس����ţ�
Private Function GetXinZangReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(cbx����, 0, 33)
    If strReport <> "" Then strReport = "�����ࡿ  " & strReport & "  "
    
    GetXinZangReport = strReport
End Function

'��ȡ���Ʊ���
Private Function GetFuKeReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""
    
    '�ӹ����
    strTemp = GetSectionReportWithCombobox(cbx�ӹ�, 0, 5)
    If strTemp <> "" Then strReport = "���ӹ������  " & strTemp & "  "
    
    
    '���ѳ�
    strTemp = GetSectionReportWithCombobox(cbx���ѳ�, 0, 1)
    
    If chk����(0).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk����(0).Tag, "[value]", Report_CheckedValue)
    End If
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "�����ѳ���  " & strTemp & "  "
        
    
    
    '���ѳ�
    strTemp = GetSectionReportWithCombobox(cbx���ѳ�, 0, 1)
    
    If chk����(1).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk����(1).Tag, "[value]", Report_CheckedValue)
    End If
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "�����ѳ���  " & strTemp & "  "
    
    
    '̥��
    strTemp = GetSectionReportWithCombobox(cbx̥��, 0, 2)
    
    If strTemp <> "" Then strTemp = strTemp & "  "
    strTemp = strTemp & GetSectionReportWithCombobox(chk����, 2, 6, True)
    
    If cbx����ԭʼ�Ĺܲ���.Text <> "" Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(cbx����ԭʼ�Ĺܲ���.Tag, "[value]", cbx����ԭʼ�Ĺܲ���.Text)
    End If
    
    
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "��̥�ҡ�  " & strTemp & "  "
    
    GetFuKeReport = strReport
End Function

'��ȡ���Ʊ���
Private Function GetChanKeReport() As String
    Dim strReport As String
    
    strReport = ""
    
    strReport = GetSectionReportWithCombobox(cbx�������, 0, 11)

    If strReport <> "" Then strReport = strReport & "  "
    strReport = strReport & GetSectionReportWithCombobox(chk�������, 0, 3)
    
    If strReport <> "" Then strReport = strReport & "  "
    strReport = strReport & GetSectionReportWithCombobox(cbx�������, 12, 21)
    
    If strReport <> "" Then strReport = "�����������  " & strReport & "  "
    
    GetChanKeReport = strReport
End Function

'��ȡ��������
Private Function GetFuBuReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '�������
    strTemp = GetSectionReportWithCombobox(cbx�������, 0, 4)
    
    If strTemp <> "" Then strTemp = strTemp & "  "
    strTemp = strTemp & GetSectionReportWithCombobox(chk�������, 0, 5, True)
    
    If strTemp <> "" Then strReport = "�����������  " & strTemp & "  "
    
    
    'Ѫ��
    strTemp = GetSectionReportWithCombobox(cbxѪ��, 0, 1)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "��Ѫ�ܡ�  " & strTemp & "  "
        

    '����
    strTemp = GetSectionReportWithCombobox(cbx�������, 0, 12)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "�����ҡ�  " & strTemp & "  "
    
    
    '���ܹ�
    strTemp = GetSectionReportWithCombobox(cbx���ܹ�, 0, 5)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "�����ܹܡ�  " & strTemp & "  "
    
    
    '����
    strTemp = GetSectionReportWithCombobox(cbx����, 0, 6)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "�����١�  " & strTemp & "  "
    
    
    'Ƣ��
    strTemp = GetSectionReportWithCombobox(cbxƢ��, 0, 4)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "��Ƣ�ࡿ  " & strTemp & "  "
    
    
    GetFuBuReport = strReport
End Function



Private Function GetSectionReportWithCombobox(aryControl As Variant, ByVal lngStartIndex As Long, _
    ByVal lngEndIndex As Long, Optional ByVal blnIsCheckBox As Boolean = False, _
    Optional ByVal blnAutoReplace As Boolean = True) As String
'���ݴ��ݵ������ȡ������Ϣ
    Dim i As Long
    Dim strResult As String
    Dim strCurElement As String
    Dim blnChange As Boolean
    
    Dim blnAddTag1 As Boolean
    Dim blnAddTag2 As Boolean
    Dim blnAddTag3 As Boolean
    Dim blnAddTag4 As Boolean
    Dim blnAddTag5 As Boolean
    Dim blnAddTag6 As Boolean
    Dim blnAddTag7 As Boolean
    
    strResult = ""
    blnChange = False
    
    blnAddTag1 = False: blnAddTag2 = False: blnAddTag3 = False: blnAddTag4 = False: blnAddTag5 = False: blnAddTag6 = False: blnAddTag7 = False
    
    For i = lngStartIndex To lngEndIndex
    
        If blnIsCheckBox Then
            blnChange = aryControl(i).value <> 0
        Else
            blnChange = aryControl(i).Text <> ""
        End If
        
        If blnChange Then
            If strResult <> "" Then strResult = strResult & "  "
            
'            If lngCrCount >= 3 Then
'                lngCrCount = 0
'                strReport = strReport & vbCrLf & "    "
'            End If
            
            
            If blnIsCheckBox Then
                strCurElement = Replace(aryControl(i).Tag, "[value]", Report_CheckedValue)
            Else
                strCurElement = Replace(aryControl(i).Tag, "[value]", aryControl(i).Text)
            End If
            
            
            If blnAutoReplace Then
                '[T1]�滻Ϊ��������������:��
                If InStr(strCurElement, "[T1]") >= 1 And Not blnAddTag1 Then
                    strCurElement = Replace(strCurElement, "[T1]", "������������:")
                    blnAddTag1 = True
                Else
                    strCurElement = Replace(strCurElement, "[T1]", "")
                End If
                
                '[T2]�滻Ϊ���ұ�:��
                If InStr(strCurElement, "[T2]") >= 1 And Not blnAddTag2 Then
                    strCurElement = Replace(strCurElement, "[T2]", "�ұ�:")
                    blnAddTag2 = True
                Else
                    strCurElement = Replace(strCurElement, "[T2]", "")
                End If
                
                '[T3]�滻Ϊ����ˮ:��
                If InStr(strCurElement, "[T3]") >= 1 And Not blnAddTag3 Then
                    strCurElement = Replace(strCurElement, "[T3]", "��ˮ:")
                    blnAddTag3 = True
                Else
                    strCurElement = Replace(strCurElement, "[T3]", "")
                End If
                
                '[T4]�滻Ϊ���궯��Ѫ��ָ��:��
                If InStr(strCurElement, "[T4]") >= 1 And Not blnAddTag4 Then
                    strCurElement = Replace(strCurElement, "[T4]", "�궯��Ѫ��ָ��:")
                    blnAddTag4 = True
                Else
                    strCurElement = Replace(strCurElement, "[T4]", "")
                End If
                
                '[T5]�滻Ϊ������:��
                If InStr(strCurElement, "[T5]") >= 1 And Not blnAddTag5 Then
                    strCurElement = Replace(strCurElement, "[T5]", "����:")
                    blnAddTag5 = True
                Else
                    strCurElement = Replace(strCurElement, "[T5]", "")
                End If
                
                '[T6]�滻Ϊ������:��
                If InStr(strCurElement, "[T6]") >= 1 And Not blnAddTag6 Then
                    strCurElement = Replace(strCurElement, "[T6]", "����:")
                    blnAddTag6 = True
                Else
                    strCurElement = Replace(strCurElement, "[T6]", "")
                End If
                
                '[T7]�滻Ϊ��Զ��:��
                If InStr(strCurElement, "[T7]") >= 1 And Not blnAddTag7 Then
                    strCurElement = Replace(strCurElement, "[T7]", "Զ��:")
                    blnAddTag7 = True
                Else
                    strCurElement = Replace(strCurElement, "[T7]", "")
                End If
            End If
            
            
            strResult = strResult & strCurElement
            
'            lngCrCount = lngCrCount + 1
        End If
    Next i
    
    GetSectionReportWithCombobox = strResult
End Function



'��ȡ�����б���
Private Function GetMiNiaoNanReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '�������
    strTemp = strTemp & GetSectionReportWithCombobox(txt�������, 0, 5)
    If strTemp <> "" Then strReport = "�����������  " & strTemp & "  "
    
    
    '����
    strTemp = GetSectionReportWithCombobox(cbx������, 0, 5)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "������ܡ�  " & strTemp & "  "
        
    '����
    strTemp = IIf(cbx����(0).Text = "", "", Replace(cbx����(0).Tag, "[value]", cbx����(0).Text))
    If strTemp <> "" Then strTemp = strTemp & "  "
    
    strTemp = strTemp & GetSectionReportWithCombobox(chk������, 0, 1, True)
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "�����ס�  " & strTemp & "  "
    
    
    'ǰ����
    strTemp = GetSectionReportWithCombobox(txtǰ������Ϣ, 0, 2)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "��ǰ���١�  " & strTemp & "  "
    
    
    GetMiNiaoNanReport = strReport
End Function

'��ȡ����Ů����
Private Function GetMiNiaoNvReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '�������
    strTemp = strTemp & GetSectionReportWithCombobox(txtŮ�������, 0, 5)
    If strTemp <> "" Then strReport = "���� �� �� ����  " & strTemp & "  "
    
    
    '����
    strTemp = GetSectionReportWithCombobox(cbx�����, 0, 5)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "���� �� �ܡ�  " & strTemp & "  "
        
    '����
    strTemp = IIf(cbxŮ����(0).Text = "", "", Replace(cbxŮ����(0).Tag, "[value]", cbxŮ����(0).Text))
    If strTemp <> "" Then strTemp = strTemp & "  "
    
    strTemp = strTemp & GetSectionReportWithCombobox(chk����Ů, 0, 1, True)
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "���� �ס�  " & strTemp & "  "
    
    
    GetMiNiaoNvReport = strReport
End Function

'��ȡ��������
Private Function GetJingBuReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(txt������Ϣ, 0, 13)
    If strReport <> "" Then strReport = "��������  " & strReport & "  "
    
    GetJingBuReport = strReport
End Function

'��ȡ���ٱ���
Private Function GetRuXianReport() As String
    Dim strReport As String
    Dim strTemp As String

    strTemp = ""
    strReport = ""
    
    strTemp = GetSectionReportWithCombobox(txt������Ϣ, 0, 1)
    If strTemp <> "" Then strReport = "��������١�  " & strTemp & "  "
    
    strTemp = GetSectionReportWithCombobox(txt������Ϣ, 2, 3)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "���Ҳ����١�  " & strTemp & "  "
    
    GetRuXianReport = strReport
End Function

'��ȡ��ֳ������
Private Function GetShengZhiQiReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '�����ֳ��
    strTemp = strTemp & GetSectionReportWithCombobox(cbx�����ֳ��, 0, 7)
    If strTemp <> "" Then strReport = "�������ֳ����  " & strTemp & "  "
    
    
    '�Ҳ���ֳ��
    strTemp = GetSectionReportWithCombobox(cbx�Ҳ���ֳ��, 0, 7)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "���Ҳ���ֳ����  " & strTemp & "  "
    
    GetShengZhiQiReport = strReport
End Function

'��ȡ��֫��������
Private Function GetXiaZhiJingMaiReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '�����ֳ��
    strTemp = strTemp & GetSectionReportWithCombobox(chk����֫����, 0, 4, True)
    If strTemp <> "" Then strReport = "������֫������  " & strTemp & "  Ѫ��ͨ��,��ǻ��δ���쳣����,̽����ѹ���ǻ��ʧ  "
    
    
    '�Ҳ���ֳ��
    strTemp = GetSectionReportWithCombobox(chk����֫����, 0, 4, True)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "������֫������  " & strTemp & "  Ѫ��ͨ��,��ǻ��δ���쳣����,̽����ѹ���ǻ��ʧ  "
    
    GetXiaZhiJingMaiReport = strReport
End Function

'��ȡ�۲�����
Private Function GetYanBuReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(txt�۲���Ϣ, 0, 1)
    If strReport <> "" Then strReport = "���۲���  " & strReport & "  "
    
    GetYanBuReport = strReport
End Function

'��ȡ��״�ٱ���
Private Function GetJiaZhuangXianReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(txt��״����Ϣ, 0, 9)
    If strReport <> "" Then strReport = "����״�١�  " & strReport & "  "
    
    GetJiaZhuangXianReport = strReport
End Function

'��ȡ����������
Private Function GetShenDongMaiReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""
    
    strTemp = GetSectionReportWithCombobox(txt��������Ϣ, 0, 5)
    If strTemp <> "" Then strReport = "������������  " & strTemp & "  "
    
    strTemp = GetSectionReportWithCombobox(txt��������Ϣ, 6, 11)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "������������  " & strTemp & "  "
    
    GetShenDongMaiReport = strReport
End Function

'��ȡ��ǻ����
Private Function GetXiongQiangReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""
    
    strTemp = GetSectionReportWithCombobox(cbx��ǻ��Ϣ, 0, 2)
    
    
    If chk��ǻ��Ϣ(0).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk��ǻ��Ϣ(0).Tag, "[value]", Report_CheckedValue)
    End If
    
    
    strReport = GetSectionReportWithCombobox(cbx��ǻ��Ϣ, 3, 5)
    If strReport <> "" Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & strReport
    End If
    
    strReport = ""
    If chk��ǻ��Ϣ(1).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk��ǻ��Ϣ(1).Tag, "[value]", Report_CheckedValue)
    End If
    
    If strTemp <> "" Then strReport = "����ǻ��  " & strTemp & "  "
    
    GetXiongQiangReport = strReport
End Function

        
        

Private Sub cbxType_Click()
    Dim i As Long
    
    If Not cbxType.Visible Then Exit Sub
    
    For i = 0 To 15
        framType(i).Visible = False
    Next i
    
    
'    framType(Val(cbxType.Text)).Refresh
    framType(Val(cbxType.Text)).Visible = True
    
'    Call Me.Refresh
End Sub

Private Sub cbx�������_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�������_Click(Index As Integer)
    If cbx�������(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�������_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�������_Click(Index As Integer)
    If cbx�������(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx���ܹ�_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx���ܹ�_Click(Index As Integer)
    If cbx���ܹ�(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����ԭʼ�Ĺܲ���_Click()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�������_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�������_Click(Index As Integer)
    If cbx�������(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx������_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx������_Click(Index As Integer)
    If cbx������(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����Ů_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbxƢ��_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbxƢ��_Click(Index As Integer)
    If cbxƢ��(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�����_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�����_Click(Index As Integer)
    If cbx�����(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx̥��_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����_Click(Index As Integer)
    If cbx����(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx��ǻ��Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx��ǻ��Ϣ_Click(Index As Integer)
    If cbx��ǻ��Ϣ(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbxѪ��_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx����_Click(Index As Integer)
    If cbx����(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�Ҳ���ֳ��_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�Ҳ���ֳ��_Click(Index As Integer)
    If cbx�Ҳ���ֳ��(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx���ѳ�_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�ӹ�_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�ӹ�_Click(Index As Integer)
    If cbx�ӹ�(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�����ֳ��_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx�����ֳ��_Click(Index As Integer)
    If cbx�����ֳ��(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx���ѳ�_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk�������_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk����_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk�������_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk������_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk����Ů_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk����֫����_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk����֫����_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim i As Long
    
    cbxType.Width = Me.ScaleWidth - 1200
    
    For i = 0 To 15
        framType(i).Left = IIf(Me.ScaleWidth > framType(i).Width, Fix((Me.ScaleWidth - framType(i).Width) / 2), 0)
        framType(i).Top = 480
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CY22", Me.Height
End Sub

Private Sub txt�����E�����ٶ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����E������ʱ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������������ʱ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt���������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������A������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������A��ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������E������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������E��ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ����귴��ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ������ֵ����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ��������ʱ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ�������Ѫʱ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ���������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ�����������ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ���������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ζ�����������ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��״����Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txtŮ�������_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txtǰ������Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt����귴��ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt���������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������A������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������A��ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������E������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������E��ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��������Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�������_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�������ھ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�Ҽ�����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�Ҽ�������Һ�ڷ���_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�Ҽ���˶�����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ķ��ҷ�����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ķ��ҷ��ᾶ_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ķ���ǰ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�۲���Ϣ_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�ҷζ����ھ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����Һᾶ_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������ǰ�ۺ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������ǰ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��������Ѫ����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt����������ĩ�ݻ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt����������ĩ�ݻ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��������������ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������������������_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��������������ѹ��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����������ھ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt���ζ����ھ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��ζ����ھ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�����Һ�ں��_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt������ÿ����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt��������Ѫ����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt����������ĩ�ݻ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt����������ĩ�ھ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt����������ĩ�ݻ�_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt�������˶�����_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub
