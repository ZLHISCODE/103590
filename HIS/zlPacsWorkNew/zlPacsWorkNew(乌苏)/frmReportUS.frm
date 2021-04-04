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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame framType 
      Caption         =   "胸  腔"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.ComboBox cbx胸腔信息 
         Height          =   300
         Index           =   5
         ItemData        =   "frmReportUS.frx":0000
         Left            =   4200
         List            =   "frmReportUS.frx":000D
         Style           =   1  'Simple Combo
         TabIndex        =   281
         Tag             =   "进针距皮肤:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx胸腔信息 
         Height          =   300
         Index           =   4
         ItemData        =   "frmReportUS.frx":0021
         Left            =   2520
         List            =   "frmReportUS.frx":002E
         Style           =   1  'Simple Combo
         TabIndex        =   280
         Tag             =   "液性暗区:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx胸腔信息 
         Height          =   300
         Index           =   2
         ItemData        =   "frmReportUS.frx":0042
         Left            =   4200
         List            =   "frmReportUS.frx":004F
         Style           =   1  'Simple Combo
         TabIndex        =   277
         Tag             =   "进针距皮肤: [value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx胸腔信息 
         Height          =   300
         Index           =   1
         ItemData        =   "frmReportUS.frx":0063
         Left            =   2520
         List            =   "frmReportUS.frx":0070
         Style           =   1  'Simple Combo
         TabIndex        =   276
         Tag             =   "液性暗区: [value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chk胸腔信息 
         Caption         =   "已定位"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   282
         Tag             =   "已 定 位[value]"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cbx胸腔信息 
         Height          =   300
         Index           =   3
         ItemData        =   "frmReportUS.frx":0084
         Left            =   840
         List            =   "frmReportUS.frx":0091
         TabIndex        =   279
         Tag             =   "左侧胸腔:[value]"
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chk胸腔信息 
         Caption         =   "已定位"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   278
         Tag             =   "已定位[value]"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbx胸腔信息 
         Height          =   300
         Index           =   0
         ItemData        =   "frmReportUS.frx":00A5
         Left            =   840
         List            =   "frmReportUS.frx":00B2
         TabIndex        =   275
         Tag             =   "左侧胸腔: [value]"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lable1 
         Caption         =   "右侧胸腔          液性暗区      cm 进针距皮肤      cm"
         Height          =   255
         Index           =   145
         Left            =   120
         TabIndex        =   475
         Top             =   630
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "左侧胸腔          液性暗区      cm 进针距皮肤      cm"
         Height          =   255
         Index           =   148
         Left            =   120
         TabIndex        =   474
         Top             =   270
         Width           =   4815
      End
   End
   Begin VB.Frame framType 
      Caption         =   "肾动脉"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   11
         Left            =   3720
         TabIndex        =   274
         Tag             =   "[T7]AT:[value]s"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   10
         Left            =   2760
         TabIndex        =   273
         Tag             =   "[T7]RI:[value]"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   9
         Left            =   1485
         TabIndex        =   272
         Tag             =   "[T7]PS:[value]cm/s"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   8
         Left            =   3720
         TabIndex        =   271
         Tag             =   "[T6]AT: [value]s"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   7
         Left            =   2760
         TabIndex        =   270
         Tag             =   "[T6]RI: [value]"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   6
         Left            =   1485
         TabIndex        =   269
         Tag             =   "[T6]PS: [value]cm/s"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   5
         Left            =   3720
         TabIndex        =   268
         Tag             =   "[T7]AT:[value]s"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   4
         Left            =   2760
         TabIndex        =   267
         Tag             =   "[T7]RI:[value]"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   3
         Left            =   1485
         TabIndex        =   266
         Tag             =   "[T7]PS:[value]cm/s"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   2
         Left            =   3720
         TabIndex        =   265
         Tag             =   "[T6]AT: [value]s"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   264
         Tag             =   "[T6]RI: [value]"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt肾动脉信息 
         Height          =   270
         Index           =   0
         Left            =   1480
         TabIndex        =   263
         Tag             =   "[T6]PS: [value]cm/s"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "右肾动脉远端:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   150
         Left            =   120
         TabIndex        =   480
         Top             =   1350
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "右肾动脉近端:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   149
         Left            =   120
         TabIndex        =   479
         Top             =   990
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "左肾动脉远端:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   146
         Left            =   120
         TabIndex        =   478
         Top             =   630
         Width           =   4815
      End
      Begin VB.Label lable1 
         Caption         =   "左肾动脉近端:PS      cm/s  RI        AT       s"
         Height          =   255
         Index           =   147
         Left            =   120
         TabIndex        =   477
         Top             =   270
         Width           =   4815
      End
   End
   Begin VB.Frame framType 
      Caption         =   "甲状腺"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   9
         Left            =   1320
         TabIndex        =   262
         Tag             =   "甲状腺总体积:[value]cm^3"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   8
         Left            =   4080
         TabIndex        =   261
         Tag             =   "峡部甲状腺宽:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   7
         Left            =   2760
         TabIndex        =   260
         Tag             =   "峡部甲状腺厚:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   6
         Left            =   1320
         TabIndex        =   259
         Tag             =   "峡部甲状腺长:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   5
         Left            =   4080
         TabIndex        =   258
         Tag             =   "右侧甲状腺宽:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   4
         Left            =   2760
         TabIndex        =   257
         Tag             =   "右侧甲状腺厚:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   3
         Left            =   1320
         TabIndex        =   256
         Tag             =   "右侧甲状腺长:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   2
         Left            =   4080
         TabIndex        =   255
         Tag             =   "左侧甲状腺宽:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   254
         Tag             =   "左侧甲状腺厚:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt甲状腺信息 
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   253
         Tag             =   "左侧甲状腺长:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "甲状腺总体积       cm^3"
         Height          =   255
         Index           =   138
         Left            =   120
         TabIndex        =   460
         Top             =   1365
         Width           =   2175
      End
      Begin VB.Label lable1 
         Caption         =   "峡部甲状腺:长      cm      厚      cm    宽       cm"
         Height          =   255
         Index           =   137
         Left            =   120
         TabIndex        =   459
         Top             =   1005
         Width           =   4935
      End
      Begin VB.Label lable1 
         Caption         =   "右侧甲状腺:长      cm      厚      cm    宽       cm"
         Height          =   255
         Index           =   134
         Left            =   120
         TabIndex        =   458
         Top             =   645
         Width           =   4935
      End
      Begin VB.Label lable1 
         Caption         =   "左侧甲状腺:长      cm      厚      cm    宽       cm"
         Height          =   255
         Index           =   132
         Left            =   120
         TabIndex        =   457
         Top             =   280
         Width           =   4935
      End
   End
   Begin VB.Frame framType 
      Caption         =   "眼  部"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.TextBox txt眼部信息 
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   248
         Tag             =   "左眼轴长:[value]cm"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txt眼部信息 
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   252
         Tag             =   "右眼轴长:[value]cm"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "左眼轴长      cm      右眼轴长      cm"
         Height          =   255
         Index           =   133
         Left            =   120
         TabIndex        =   455
         Top             =   390
         Width           =   4935
      End
   End
   Begin VB.Frame framType 
      Caption         =   "下肢静脉"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "左下肢静脉"
         Height          =   975
         Index           =   29
         Left            =   120
         TabIndex        =   470
         Top             =   240
         Width           =   5655
         Begin VB.CheckBox chk左下肢静脉 
            Caption         =   "足背静脉"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   245
            Tag             =   "足背静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk左下肢静脉 
            Caption         =   "腘静脉"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   244
            Tag             =   "腘静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk左下肢静脉 
            Caption         =   "股浅静脉"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   243
            Tag             =   "股浅静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk左下肢静脉 
            Caption         =   "股深静脉"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   242
            Tag             =   "股深静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk左下肢静脉 
            Caption         =   "股总静脉"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   241
            Tag             =   "股总静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "血流通畅,管腔内未见异常回声,探过加压后管腔消失"
            Height          =   255
            Index           =   143
            Left            =   120
            TabIndex        =   471
            Top             =   600
            Width           =   4455
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "右下肢静脉"
         Height          =   975
         Index           =   28
         Left            =   120
         TabIndex        =   469
         Top             =   1320
         Width           =   5655
         Begin VB.CheckBox chk右下肢静脉 
            Caption         =   "足背静脉"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   251
            Tag             =   "足背静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk右下肢静脉 
            Caption         =   "腘静脉"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   250
            Tag             =   "腘静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk右下肢静脉 
            Caption         =   "股浅静脉"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   249
            Tag             =   "股浅静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk右下肢静脉 
            Caption         =   "股深静脉"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   247
            Tag             =   "股深静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk右下肢静脉 
            Caption         =   "股总静脉"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   246
            Tag             =   "股总静脉[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "血流通畅,管腔内未见异常回声,探过加压后管腔消失"
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
      Caption         =   "生殖器"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "右侧生殖器"
         Height          =   975
         Index           =   27
         Left            =   120
         TabIndex        =   463
         Top             =   1320
         Width           =   5655
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":00C6
            Left            =   3120
            List            =   "frmReportUS.frx":00D6
            Style           =   1  'Simple Combo
            TabIndex        =   239
            Tag             =   "宽:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":00EE
            Left            =   1920
            List            =   "frmReportUS.frx":00FE
            Style           =   1  'Simple Combo
            TabIndex        =   238
            Tag             =   "厚:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0116
            Left            =   720
            List            =   "frmReportUS.frx":0126
            Style           =   1  'Simple Combo
            TabIndex        =   237
            Tag             =   "付睾长:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":013E
            Left            =   3120
            List            =   "frmReportUS.frx":014E
            Style           =   1  'Simple Combo
            TabIndex        =   235
            Tag             =   "宽: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0166
            Left            =   1920
            List            =   "frmReportUS.frx":0176
            Style           =   1  'Simple Combo
            TabIndex        =   234
            Tag             =   "厚: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":018E
            Left            =   720
            List            =   "frmReportUS.frx":019E
            Style           =   1  'Simple Combo
            TabIndex        =   233
            Tag             =   "睾丸长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   7
            ItemData        =   "frmReportUS.frx":01B6
            Left            =   4440
            List            =   "frmReportUS.frx":01C6
            TabIndex        =   240
            Tag             =   "血流:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx右侧生殖器 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":01DE
            Left            =   4440
            List            =   "frmReportUS.frx":01EE
            TabIndex        =   236
            Tag             =   "血流: [value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "付睾长      cm    厚      cm  宽       cm  血流"
            Height          =   255
            Index           =   141
            Left            =   120
            TabIndex        =   467
            Top             =   645
            Width           =   4335
         End
         Begin VB.Label lable1 
            Caption         =   "睾丸长      cm    厚      cm  宽       cm  血流"
            Height          =   255
            Index           =   140
            Left            =   120
            TabIndex        =   466
            Top             =   285
            Width           =   4335
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "左侧生殖器"
         Height          =   975
         Index           =   26
         Left            =   120
         TabIndex        =   462
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":0206
            Left            =   3120
            List            =   "frmReportUS.frx":0216
            Style           =   1  'Simple Combo
            TabIndex        =   231
            Tag             =   "宽:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":022E
            Left            =   1920
            List            =   "frmReportUS.frx":023E
            Style           =   1  'Simple Combo
            TabIndex        =   230
            Tag             =   "厚:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0256
            Left            =   720
            List            =   "frmReportUS.frx":0266
            Style           =   1  'Simple Combo
            TabIndex        =   229
            Tag             =   "付睾长:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":027E
            Left            =   3120
            List            =   "frmReportUS.frx":028E
            Style           =   1  'Simple Combo
            TabIndex        =   227
            Tag             =   "宽: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":02A6
            Left            =   1920
            List            =   "frmReportUS.frx":02B6
            Style           =   1  'Simple Combo
            TabIndex        =   226
            Tag             =   "厚: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":02CE
            Left            =   720
            List            =   "frmReportUS.frx":02DE
            Style           =   1  'Simple Combo
            TabIndex        =   225
            Tag             =   "睾丸长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   7
            ItemData        =   "frmReportUS.frx":02F6
            Left            =   4440
            List            =   "frmReportUS.frx":0306
            TabIndex        =   232
            Tag             =   "血流:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx左侧生殖器 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":031E
            Left            =   4440
            List            =   "frmReportUS.frx":032E
            TabIndex        =   228
            Tag             =   "血流: [value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "付睾长      cm    厚      cm  宽       cm  血流"
            Height          =   255
            Index           =   139
            Left            =   120
            TabIndex        =   465
            Top             =   645
            Width           =   4335
         End
         Begin VB.Label lable1 
            Caption         =   "睾丸长      cm    厚      cm  宽       cm  血流"
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
      Caption         =   "乳  腺"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "左侧乳腺"
         Height          =   615
         Index           =   31
         Left            =   120
         TabIndex        =   485
         Top             =   960
         Width           =   5655
         Begin VB.TextBox txt乳腺信息 
            Height          =   270
            Index           =   2
            Left            =   720
            TabIndex        =   487
            Tag             =   "右乳厚:[value]cm "
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt乳腺信息 
            Height          =   270
            Index           =   3
            Left            =   3360
            TabIndex        =   486
            Tag             =   "右乳导管内径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "右乳厚       cm        右乳导管内径       cm"
            Height          =   255
            Index           =   131
            Left            =   120
            TabIndex        =   488
            Top             =   270
            Width           =   4935
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "左侧乳腺"
         Height          =   615
         Index           =   30
         Left            =   120
         TabIndex        =   481
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txt乳腺信息 
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   483
            Tag             =   "左乳厚:[value]cm "
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt乳腺信息 
            Height          =   270
            Index           =   1
            Left            =   3360
            TabIndex        =   482
            Tag             =   "左乳导管内径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "左乳厚       cm        左乳导管内径       cm"
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
      Caption         =   "颈  部"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   13
         Left            =   3120
         TabIndex        =   224
         Tag             =   "椎动脉内中膜厚:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   12
         Left            =   1440
         TabIndex        =   223
         Tag             =   "椎动脉内径:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   11
         Left            =   3480
         TabIndex        =   222
         Tag             =   "颈外动脉内径:[value]cm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   10
         Left            =   1440
         TabIndex        =   221
         Tag             =   "右颈内动脉内径:[value]cm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   7
         Left            =   3120
         TabIndex        =   218
         Tag             =   "右颈总动脉内中膜厚:[value]cm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   6
         Left            =   1440
         TabIndex        =   217
         Tag             =   "右颈总动脉内径:[value]cm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   8
         Left            =   1920
         TabIndex        =   219
         Tag             =   "右颈总动脉膨大处内径:[value]cm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   9
         Left            =   3960
         TabIndex        =   220
         Tag             =   "右颈总动脉膨大处内中膜厚:[value]cm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   5
         Left            =   3480
         TabIndex        =   216
         Tag             =   "颈外动脉内径: [value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   4
         Left            =   1440
         TabIndex        =   215
         Tag             =   "左颈内动脉内径:[value]cm"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   1
         Left            =   3120
         TabIndex        =   212
         Tag             =   "左颈总动脉内中膜厚:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   211
         Tag             =   "左颈总动脉内径:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   2
         Left            =   1920
         TabIndex        =   213
         Tag             =   "左颈总动脉膨大处内径[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt颈部信息 
         Height          =   270
         Index           =   3
         Left            =   3960
         TabIndex        =   214
         Tag             =   "左颈总动脉膨大处内中膜厚:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lable1 
         Caption         =   "    椎动脉内径      cm  内中膜厚       cm"
         Height          =   255
         Index           =   129
         Left            =   120
         TabIndex        =   452
         Top             =   2430
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "右颈内动脉内径      cm  颈外动脉内径       cm"
         Height          =   255
         Index           =   128
         Left            =   120
         TabIndex        =   451
         Top             =   2070
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "右颈总动脉内径      cm  内中膜厚       cm"
         Height          =   255
         Index           =   127
         Left            =   120
         TabIndex        =   450
         Top             =   1350
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "右颈总动脉膨大处内径         cm  内中膜厚         cm"
         Height          =   255
         Index           =   126
         Left            =   120
         TabIndex        =   449
         Top             =   1710
         Width           =   5175
      End
      Begin VB.Label lable1 
         Caption         =   "左颈内动脉内径      cm  颈外动脉内径       cm"
         Height          =   255
         Index           =   125
         Left            =   120
         TabIndex        =   448
         Top             =   990
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "左颈总动脉内径      cm  内中膜厚       cm"
         Height          =   255
         Index           =   136
         Left            =   120
         TabIndex        =   447
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lable1 
         Caption         =   "左颈总动脉膨大处内径         cm  内中膜厚         cm"
         Height          =   255
         Index           =   135
         Left            =   120
         TabIndex        =   446
         Top             =   630
         Width           =   5175
      End
   End
   Begin VB.Frame framType 
      Caption         =   "泌尿(女)"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "膀胱"
         Height          =   615
         Index           =   25
         Left            =   120
         TabIndex        =   443
         Top             =   2880
         Width           =   5655
         Begin VB.CheckBox chk泌尿女 
            Caption         =   "观察不清"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   209
            Tag             =   "观察不清[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cbx女膀胱 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0346
            Left            =   600
            List            =   "frmReportUS.frx":0356
            TabIndex        =   208
            Tag             =   "膀胱:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chk泌尿女 
            Caption         =   "未见明显异常"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   210
            Tag             =   "未见明显异常[value]"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lable1 
            Caption         =   "膀胱"
            Height          =   255
            Index           =   124
            Left            =   120
            TabIndex        =   444
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "输尿管"
         Height          =   1695
         Index           =   23
         Left            =   120
         TabIndex        =   435
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cbx输尿管 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":0374
            Left            =   4220
            List            =   "frmReportUS.frx":0381
            Style           =   1  'Simple Combo
            TabIndex        =   207
            Tag             =   "下端膀胱壁内段见增强回声光斑:[value]cm伴声影"
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox cbx输尿管 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":038F
            Left            =   3020
            List            =   "frmReportUS.frx":039C
            Style           =   1  'Simple Combo
            TabIndex        =   206
            Tag             =   "上端内径:[value]cm"
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cbx输尿管 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":03AA
            Left            =   4220
            List            =   "frmReportUS.frx":03B7
            Style           =   1  'Simple Combo
            TabIndex        =   204
            Tag             =   "下端膀胱壁内段见增强回声光斑: [value]cm伴声影"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx输尿管 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":03C5
            Left            =   3020
            List            =   "frmReportUS.frx":03D2
            Style           =   1  'Simple Combo
            TabIndex        =   203
            Tag             =   "上端内径: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx输尿管 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":03E0
            Left            =   1080
            List            =   "frmReportUS.frx":03ED
            TabIndex        =   202
            Tag             =   "左侧输尿管:[value]扩张"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cbx输尿管 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":03FB
            Left            =   1080
            List            =   "frmReportUS.frx":0408
            TabIndex        =   205
            Tag             =   "右侧输尿管:[value]扩张"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "左侧输尿管         扩张 上端内径      cm  "
            Height          =   255
            Index           =   121
            Left            =   120
            TabIndex        =   439
            Top             =   285
            Width           =   4095
         End
         Begin VB.Label lable1 
            Caption         =   "下端膀胱壁内段见增强回声光斑      cm伴声影"
            Height          =   255
            Index           =   120
            Left            =   1680
            TabIndex        =   438
            Top             =   630
            Width           =   3855
         End
         Begin VB.Label lable1 
            Caption         =   "右侧输尿管         扩张 上端内径      cm  "
            Height          =   255
            Index           =   119
            Left            =   120
            TabIndex        =   437
            Top             =   1005
            Width           =   4095
         End
         Begin VB.Label lable1 
            Caption         =   "下端膀胱壁内段见增强回声光斑      cm伴声影"
            Height          =   255
            Index           =   106
            Left            =   1680
            TabIndex        =   436
            Top             =   1350
            Width           =   3855
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "肾脏情况"
         Height          =   975
         Index           =   24
         Left            =   120
         TabIndex        =   440
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txt女肾脏情况 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   198
            Tag             =   "厚: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt女肾脏情况 
            Height          =   270
            Index           =   1
            Left            =   2040
            TabIndex        =   197
            Tag             =   "宽: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt女肾脏情况 
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   196
            Tag             =   "左肾长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt女肾脏情况 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   201
            Tag             =   "厚:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt女肾脏情况 
            Height          =   270
            Index           =   4
            Left            =   2040
            TabIndex        =   200
            Tag             =   "宽:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt女肾脏情况 
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   199
            Tag             =   "右肾长:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "左肾长       cm    宽       cm    厚        cm"
            Height          =   255
            Index           =   123
            Left            =   120
            TabIndex        =   442
            Top             =   285
            Width           =   5415
         End
         Begin VB.Label lable1 
            Caption         =   "右肾长       cm    宽       cm    厚        cm"
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
      Caption         =   "泌尿(男)"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "前列腺"
         Height          =   615
         Index           =   19
         Left            =   120
         TabIndex        =   422
         Top             =   3480
         Width           =   5655
         Begin VB.TextBox txt前列腺信息 
            Height          =   270
            Index           =   2
            Left            =   3120
            TabIndex        =   195
            Tag             =   "厚:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt前列腺信息 
            Height          =   270
            Index           =   1
            Left            =   1680
            TabIndex        =   194
            Tag             =   "宽:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt前列腺信息 
            Height          =   270
            Index           =   0
            Left            =   360
            TabIndex        =   193
            Tag             =   "长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "长       cm    宽       cm    厚       cm"
            Height          =   255
            Index           =   105
            Left            =   120
            TabIndex        =   423
            Top             =   285
            Width           =   4575
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "膀胱"
         Height          =   615
         Index           =   20
         Left            =   120
         TabIndex        =   424
         Top             =   2880
         Width           =   5655
         Begin VB.CheckBox chk泌尿男 
            Caption         =   "未见明显异常"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   192
            Tag             =   "未见明显异常[value]"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cbx膀胱 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0416
            Left            =   600
            List            =   "frmReportUS.frx":0429
            TabIndex        =   190
            Tag             =   "膀胱:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chk泌尿男 
            Caption         =   "观察不清"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   191
            Tag             =   "观察不清[value]"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "膀胱"
            Height          =   255
            Index           =   107
            Left            =   120
            TabIndex        =   425
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "输尿管"
         Height          =   1695
         Index           =   21
         Left            =   120
         TabIndex        =   426
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cbx泌尿男 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":044F
            Left            =   4220
            List            =   "frmReportUS.frx":045C
            Style           =   1  'Simple Combo
            TabIndex        =   189
            Tag             =   "下端膀胱壁内段见增强回声光斑:[value]cm伴声影"
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox cbx泌尿男 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":046A
            Left            =   3020
            List            =   "frmReportUS.frx":0477
            Style           =   1  'Simple Combo
            TabIndex        =   188
            Tag             =   "上端内径:[value]cm"
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cbx泌尿男 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0485
            Left            =   4220
            List            =   "frmReportUS.frx":0492
            Style           =   1  'Simple Combo
            TabIndex        =   186
            Tag             =   "下端膀胱壁内段见增强回声光斑: [value]cm伴声影"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx泌尿男 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":04A0
            Left            =   3020
            List            =   "frmReportUS.frx":04AD
            Style           =   1  'Simple Combo
            TabIndex        =   185
            Tag             =   "上端内径: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx泌尿男 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":04BB
            Left            =   1080
            List            =   "frmReportUS.frx":04C8
            TabIndex        =   187
            Tag             =   "右侧输尿管:[value]扩张"
            Top             =   960
            Width           =   735
         End
         Begin VB.ComboBox cbx泌尿男 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":04D6
            Left            =   1080
            List            =   "frmReportUS.frx":04E3
            TabIndex        =   184
            Tag             =   "左侧输尿管:[value]扩张"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "下端膀胱壁内段见增强回声光斑      cm伴声影"
            Height          =   255
            Index           =   118
            Left            =   1680
            TabIndex        =   433
            Top             =   1350
            Width           =   3855
         End
         Begin VB.Label lable1 
            Caption         =   "右侧输尿管         扩张 上端内径      cm  "
            Height          =   255
            Index           =   117
            Left            =   120
            TabIndex        =   432
            Top             =   1005
            Width           =   4095
         End
         Begin VB.Label lable1 
            Caption         =   "下端膀胱壁内段见增强回声光斑      cm伴声影"
            Height          =   255
            Index           =   116
            Left            =   1680
            TabIndex        =   431
            Top             =   630
            Width           =   3855
         End
         Begin VB.Label lable1 
            Caption         =   "左侧输尿管         扩张 上端内径      cm  "
            Height          =   255
            Index           =   108
            Left            =   120
            TabIndex        =   427
            Top             =   285
            Width           =   4095
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "肾脏情况"
         Height          =   975
         Index           =   22
         Left            =   120
         TabIndex        =   428
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txt肾脏情况 
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   181
            Tag             =   "右肾长:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt肾脏情况 
            Height          =   270
            Index           =   4
            Left            =   2040
            TabIndex        =   182
            Tag             =   "宽:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt肾脏情况 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   183
            Tag             =   "厚:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt肾脏情况 
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   178
            Tag             =   "左肾长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt肾脏情况 
            Height          =   270
            Index           =   1
            Left            =   2040
            TabIndex        =   179
            Tag             =   "宽: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt肾脏情况 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   180
            Tag             =   "厚: [value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "右肾长       cm    宽       cm    厚        cm"
            Height          =   255
            Index           =   114
            Left            =   120
            TabIndex        =   430
            Top             =   645
            Width           =   5415
         End
         Begin VB.Label lable1 
            Caption         =   "左肾长       cm    宽       cm    厚        cm"
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
      Caption         =   "腹  部"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "脾  脏"
         Height          =   975
         Index           =   14
         Left            =   120
         TabIndex        =   382
         Top             =   5880
         Width           =   5655
         Begin VB.ComboBox cbx脾脏 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":04F1
            Left            =   2880
            List            =   "frmReportUS.frx":0504
            Style           =   1  'Simple Combo
            TabIndex        =   174
            Tag             =   "脾门-脾尖长径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx脾脏 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":052E
            Left            =   600
            List            =   "frmReportUS.frx":0541
            Style           =   1  'Simple Combo
            TabIndex        =   173
            Tag             =   "厚径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx脾脏 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":056B
            Left            =   4320
            List            =   "frmReportUS.frx":0578
            TabIndex        =   177
            Tag             =   "彩色血流:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx脾脏 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":058A
            Left            =   600
            List            =   "frmReportUS.frx":059D
            TabIndex        =   176
            Tag             =   "反射:[value]"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cbx脾脏 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":05C7
            Left            =   4320
            List            =   "frmReportUS.frx":05D4
            TabIndex        =   175
            Tag             =   "形态:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "彩色血流"
            Height          =   255
            Index           =   47
            Left            =   3480
            TabIndex        =   387
            Top             =   645
            Width           =   855
         End
         Begin VB.Label lable1 
            Caption         =   "反射"
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   386
            Top             =   645
            Width           =   405
         End
         Begin VB.Label lable1 
            Caption         =   "形态"
            Height          =   255
            Index           =   46
            Left            =   3840
            TabIndex        =   385
            Top             =   285
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "厚径       cm"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   384
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "脾门—脾尖长径      cm"
            Height          =   255
            Index           =   48
            Left            =   1560
            TabIndex        =   383
            Top             =   285
            Width           =   2055
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "胰  腺"
         Height          =   975
         Index           =   13
         Left            =   120
         TabIndex        =   374
         Top             =   4920
         Width           =   5655
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":05E6
            Left            =   4800
            List            =   "frmReportUS.frx":05F6
            Style           =   1  'Simple Combo
            TabIndex        =   169
            Tag             =   "胰管内径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0612
            Left            =   3360
            List            =   "frmReportUS.frx":0622
            Style           =   1  'Simple Combo
            TabIndex        =   168
            Tag             =   "胰尾厚:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":063E
            Left            =   2030
            List            =   "frmReportUS.frx":064E
            Style           =   1  'Simple Combo
            TabIndex        =   167
            Tag             =   "胰体厚:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":066A
            Left            =   700
            List            =   "frmReportUS.frx":067A
            Style           =   1  'Simple Combo
            TabIndex        =   166
            Tag             =   "胰头厚:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":0696
            Left            =   4320
            List            =   "frmReportUS.frx":06A3
            TabIndex        =   172
            Tag             =   "彩色血流:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":06B5
            Left            =   2280
            List            =   "frmReportUS.frx":06C8
            TabIndex        =   171
            Tag             =   "反射:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx胰腺 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":06F8
            Left            =   600
            List            =   "frmReportUS.frx":0708
            TabIndex        =   170
            Tag             =   "包膜:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "彩色血流"
            Height          =   255
            Index           =   45
            Left            =   3480
            TabIndex        =   381
            Top             =   645
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "胰管内径      cm"
            Height          =   255
            Index           =   44
            Left            =   4080
            TabIndex        =   380
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label lable1 
            Caption         =   "胰尾厚      cm"
            Height          =   255
            Index           =   43
            Left            =   2760
            TabIndex        =   379
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "反射"
            Height          =   255
            Index           =   42
            Left            =   1800
            TabIndex        =   378
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "包膜"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   377
            Top             =   645
            Width           =   405
         End
         Begin VB.Label lable1 
            Caption         =   "胰头厚      cm"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   376
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "胰体厚      cm"
            Height          =   255
            Index           =   41
            Left            =   1440
            TabIndex        =   375
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "胆总管"
         Height          =   975
         Index           =   12
         Left            =   120
         TabIndex        =   367
         Top             =   3960
         Width           =   5655
         Begin VB.ComboBox cbx胆总管 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0724
            Left            =   2400
            List            =   "frmReportUS.frx":073D
            Style           =   1  'Simple Combo
            TabIndex        =   164
            Tag             =   "大小:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx胆总管 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":076B
            Left            =   2400
            List            =   "frmReportUS.frx":0784
            Style           =   1  'Simple Combo
            TabIndex        =   161
            Tag             =   "可见长度:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胆总管 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":07B2
            Left            =   600
            List            =   "frmReportUS.frx":07CB
            Style           =   1  'Simple Combo
            TabIndex        =   160
            Tag             =   "内径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胆总管 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":07F9
            Left            =   4320
            List            =   "frmReportUS.frx":0806
            TabIndex        =   162
            Tag             =   "病变部位:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cbx胆总管 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":0818
            Left            =   600
            List            =   "frmReportUS.frx":0831
            TabIndex        =   163
            Tag             =   "腔:[value]"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cbx胆总管 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":085F
            Left            =   4320
            List            =   "frmReportUS.frx":086F
            TabIndex        =   165
            Tag             =   "声影:[value]cm"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "可见长度       cm"
            Height          =   255
            Index           =   35
            Left            =   1560
            TabIndex        =   373
            Top             =   285
            Width           =   1695
         End
         Begin VB.Label lable1 
            Caption         =   "内径       cm"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   372
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lable1 
            Caption         =   "病变部位"
            Height          =   255
            Index           =   37
            Left            =   3480
            TabIndex        =   371
            Top             =   285
            Width           =   735
         End
         Begin VB.Label lable1 
            Caption         =   "  腔"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   370
            Top             =   645
            Width           =   405
         End
         Begin VB.Label lable1 
            Caption         =   "大小       cm"
            Height          =   255
            Index           =   36
            Left            =   1920
            TabIndex        =   369
            Top             =   645
            Width           =   1215
         End
         Begin VB.Label lable1 
            Caption         =   "声影"
            Height          =   255
            Index           =   38
            Left            =   3840
            TabIndex        =   368
            Top             =   645
            Width           =   495
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "胆  囊"
         Height          =   1695
         Index           =   11
         Left            =   120
         TabIndex        =   353
         Top             =   2280
         Width           =   5655
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   10
            ItemData        =   "frmReportUS.frx":0887
            Left            =   1680
            List            =   "frmReportUS.frx":0897
            Style           =   1  'Simple Combo
            TabIndex        =   157
            Tag             =   "大小:[value]cm"
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   7
            ItemData        =   "frmReportUS.frx":08B1
            Left            =   2880
            List            =   "frmReportUS.frx":08C1
            Style           =   1  'Simple Combo
            TabIndex        =   154
            Tag             =   "光斑大小:[value]cm"
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":08DB
            Left            =   4560
            List            =   "frmReportUS.frx":08EB
            Style           =   1  'Simple Combo
            TabIndex        =   152
            Tag             =   "增厚:[value]cm"
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0905
            Left            =   2880
            List            =   "frmReportUS.frx":0915
            Style           =   1  'Simple Combo
            TabIndex        =   148
            Tag             =   "前后径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":092F
            Left            =   840
            List            =   "frmReportUS.frx":093F
            Style           =   1  'Simple Combo
            TabIndex        =   147
            Tag             =   "长径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   12
            ItemData        =   "frmReportUS.frx":0959
            Left            =   4800
            List            =   "frmReportUS.frx":0966
            TabIndex        =   159
            Tag             =   "彩色血流:[value]"
            Top             =   1320
            Width           =   735
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   11
            ItemData        =   "frmReportUS.frx":0978
            Left            =   2890
            List            =   "frmReportUS.frx":0988
            TabIndex        =   158
            Tag             =   "声影:[value]"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   9
            ItemData        =   "frmReportUS.frx":09A0
            Left            =   360
            List            =   "frmReportUS.frx":09B6
            TabIndex        =   156
            Tag             =   "腔:[value]"
            Top             =   1320
            Width           =   855
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   8
            ItemData        =   "frmReportUS.frx":09DC
            Left            =   4560
            List            =   "frmReportUS.frx":09EC
            TabIndex        =   155
            Tag             =   "光斑声影:[value]"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   6
            ItemData        =   "frmReportUS.frx":0A04
            Left            =   840
            List            =   "frmReportUS.frx":0A17
            TabIndex        =   153
            Tag             =   "囊壁光斑:[value]"
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0A51
            Left            =   2880
            List            =   "frmReportUS.frx":0A64
            TabIndex        =   151
            Tag             =   "壁:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":0A86
            Left            =   840
            List            =   "frmReportUS.frx":0A96
            TabIndex        =   150
            Tag             =   "胆囊轮廓:[value]"
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cbx胆囊情况 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0AB0
            Left            =   4560
            List            =   "frmReportUS.frx":0AC0
            TabIndex        =   149
            Tag             =   "形态:[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "彩色血流"
            Height          =   255
            Index           =   29
            Left            =   3960
            TabIndex        =   366
            Top             =   1365
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "声影"
            Height          =   255
            Index           =   30
            Left            =   2520
            TabIndex        =   365
            Top             =   1365
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "大小      cm"
            Height          =   255
            Index           =   31
            Left            =   1320
            TabIndex        =   364
            Top             =   1365
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "腔"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   363
            Top             =   1365
            Width           =   285
         End
         Begin VB.Label lable1 
            Caption         =   "声影"
            Height          =   255
            Index           =   28
            Left            =   4080
            TabIndex        =   362
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "大小      cm"
            Height          =   255
            Index           =   27
            Left            =   2520
            TabIndex        =   361
            Top             =   1005
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "囊壁光斑"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   360
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "增厚        cm"
            Height          =   255
            Index           =   25
            Left            =   4080
            TabIndex        =   359
            Top             =   645
            Width           =   1335
         End
         Begin VB.Label lable1 
            Caption         =   "壁"
            Height          =   255
            Index           =   24
            Left            =   2640
            TabIndex        =   358
            Top             =   645
            Width           =   375
         End
         Begin VB.Label lable1 
            Caption         =   "胆囊轮廓"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   357
            Top             =   645
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "形态"
            Height          =   255
            Index           =   22
            Left            =   4080
            TabIndex        =   356
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "长径      cm"
            Height          =   255
            Index           =   20
            Left            =   480
            TabIndex        =   355
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label lable1 
            Caption         =   "前后径      cm"
            Height          =   255
            Index           =   21
            Left            =   2280
            TabIndex        =   354
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "血  管"
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   350
         Top             =   1680
         Width           =   5655
         Begin VB.ComboBox cbx血管 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0AD8
            Left            =   3840
            List            =   "frmReportUS.frx":0AE8
            Style           =   1  'Simple Combo
            TabIndex        =   146
            Tag             =   "脾静脉:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx血管 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0B02
            Left            =   1080
            List            =   "frmReportUS.frx":0B12
            Style           =   1  'Simple Combo
            TabIndex        =   145
            Tag             =   "门静脉主干:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "脾静脉       cm"
            Height          =   255
            Index           =   19
            Left            =   3240
            TabIndex        =   352
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label lable1 
            Caption         =   "门静脉主干       cm"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   351
            Top             =   255
            Width           =   1815
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "肝脏情况"
         Height          =   1455
         Index           =   9
         Left            =   120
         TabIndex        =   343
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cbx肝脏情况 
            Height          =   300
            Index           =   2
            ItemData        =   "frmReportUS.frx":0B2C
            Left            =   4560
            List            =   "frmReportUS.frx":0B3F
            Style           =   1  'Simple Combo
            TabIndex        =   136
            Tag             =   "右肝上下斜径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx肝脏情况 
            Height          =   300
            Index           =   1
            ItemData        =   "frmReportUS.frx":0B5F
            Left            =   2520
            List            =   "frmReportUS.frx":0B72
            Style           =   1  'Simple Combo
            TabIndex        =   135
            Tag             =   "厚径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx肝脏情况 
            Height          =   300
            Index           =   0
            ItemData        =   "frmReportUS.frx":0B92
            Left            =   960
            List            =   "frmReportUS.frx":0BA5
            Style           =   1  'Simple Combo
            TabIndex        =   134
            Tag             =   "左肝长径:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx肝脏情况 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":0BC5
            Left            =   2880
            List            =   "frmReportUS.frx":0BD5
            TabIndex        =   138
            Tag             =   "中肝静脉:[value]"
            Top             =   660
            Width           =   1095
         End
         Begin VB.CheckBox chk肝脏情况 
            Caption         =   "胆管无扩张"
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   139
            Tag             =   "胆管无扩张[value]"
            Top             =   660
            Width           =   1215
         End
         Begin VB.CheckBox chk肝脏情况 
            Caption         =   "均质"
            Height          =   255
            Index           =   5
            Left            =   4080
            TabIndex        =   144
            Tag             =   "[T5]均质[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.CheckBox chk肝脏情况 
            Caption         =   "不均质"
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   143
            Tag             =   "[T5]不均质[value]"
            Top             =   1050
            Width           =   855
         End
         Begin VB.CheckBox chk肝脏情况 
            Caption         =   "增粗"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   142
            Tag             =   "[T5]增粗[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.CheckBox chk肝脏情况 
            Caption         =   "增强"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   141
            Tag             =   "[T5]增强[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.CheckBox chk肝脏情况 
            Caption         =   "增多"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   140
            Tag             =   "[T5]增多[value]"
            Top             =   1050
            Width           =   735
         End
         Begin VB.ComboBox cbx肝脏情况 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":0BF3
            Left            =   960
            List            =   "frmReportUS.frx":0C06
            TabIndex        =   137
            Tag             =   "形态:[value]"
            Top             =   660
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "中肝静脉"
            Height          =   255
            Index           =   16
            Left            =   2040
            TabIndex        =   349
            Top             =   720
            Width           =   885
         End
         Begin VB.Label lable1 
            Caption         =   "反    射(                                             )"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   348
            Top             =   1080
            Width           =   5295
         End
         Begin VB.Label lable1 
            Caption         =   "形    态"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   347
            Top             =   675
            Width           =   855
         End
         Begin VB.Label lable1 
            Caption         =   "右肝上下斜径      cm"
            Height          =   255
            Index           =   14
            Left            =   3480
            TabIndex        =   346
            Top             =   315
            Width           =   1815
         End
         Begin VB.Label lable1 
            Caption         =   "厚径       cm"
            Height          =   255
            Index           =   13
            Left            =   2040
            TabIndex        =   345
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label lable1 
            Caption         =   "左肝长径       cm"
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
      Caption         =   "产  科"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.ComboBox cbx产科情况 
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
      Begin VB.ComboBox cbx产科情况 
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
      Begin VB.ComboBox cbx产科情况 
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
      Begin VB.ComboBox cbx产科情况 
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
      Begin VB.ComboBox cbx产科情况 
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
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   16
         ItemData        =   "frmReportUS.frx":0CEE
         Left            =   4710
         List            =   "frmReportUS.frx":0CFE
         Style           =   1  'Simple Combo
         TabIndex        =   128
         Tag             =   "[T3]右下:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   15
         ItemData        =   "frmReportUS.frx":0D16
         Left            =   3480
         List            =   "frmReportUS.frx":0D26
         Style           =   1  'Simple Combo
         TabIndex        =   127
         Tag             =   "[T3]右上:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   14
         ItemData        =   "frmReportUS.frx":0D3E
         Left            =   2190
         List            =   "frmReportUS.frx":0D4E
         Style           =   1  'Simple Combo
         TabIndex        =   126
         Tag             =   "[T3]左下:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   13
         ItemData        =   "frmReportUS.frx":0D66
         Left            =   960
         List            =   "frmReportUS.frx":0D76
         Style           =   1  'Simple Combo
         TabIndex        =   125
         Tag             =   "[T3]左上:[value]cm"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   12
         ItemData        =   "frmReportUS.frx":0D8E
         Left            =   1320
         List            =   "frmReportUS.frx":0D9E
         Style           =   1  'Simple Combo
         TabIndex        =   124
         Tag             =   "胎头位置:[value]cm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   9
         ItemData        =   "frmReportUS.frx":0DB6
         Left            =   2640
         List            =   "frmReportUS.frx":0DC6
         Style           =   1  'Simple Combo
         TabIndex        =   117
         Tag             =   "厚度:[value]cm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   5
         ItemData        =   "frmReportUS.frx":0DDE
         Left            =   4320
         List            =   "frmReportUS.frx":0DEE
         Style           =   1  'Simple Combo
         TabIndex        =   113
         Tag             =   "股骨长:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   4
         ItemData        =   "frmReportUS.frx":0E06
         Left            =   2640
         List            =   "frmReportUS.frx":0E16
         Style           =   1  'Simple Combo
         TabIndex        =   112
         Tag             =   "腹围:[value]cm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   2
         ItemData        =   "frmReportUS.frx":0E2E
         Left            =   4320
         List            =   "frmReportUS.frx":0E3E
         Style           =   1  'Simple Combo
         TabIndex        =   110
         Tag             =   "头围:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   1
         ItemData        =   "frmReportUS.frx":0E56
         Left            =   2640
         List            =   "frmReportUS.frx":0E66
         Style           =   1  'Simple Combo
         TabIndex        =   109
         Tag             =   "双顶径:[value]cm"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chk产科情况 
         Caption         =   "膀胱未充盈胎盘下缘观察不清"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   123
         Tag             =   "膀胱未充盈胎盘下缘观察不清[value]"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chk产科情况 
         Caption         =   "部分型"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   122
         Tag             =   "部分型[value]"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chk产科情况 
         Caption         =   "边缘型"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   121
         Tag             =   "边缘型[value]"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chk产科情况 
         Caption         =   "中央型"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   120
         Tag             =   "中央型[value]"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   11
         ItemData        =   "frmReportUS.frx":0E7E
         Left            =   960
         List            =   "frmReportUS.frx":0E8B
         TabIndex        =   119
         Tag             =   "前置胎盘:[value]"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   10
         ItemData        =   "frmReportUS.frx":0E99
         Left            =   4320
         List            =   "frmReportUS.frx":0EB5
         TabIndex        =   118
         Tag             =   "胎盘级别:[value]"
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   8
         ItemData        =   "frmReportUS.frx":0EE6
         Left            =   960
         List            =   "frmReportUS.frx":0F05
         TabIndex        =   116
         Tag             =   "胎盘位置:[value]"
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   7
         ItemData        =   "frmReportUS.frx":0F43
         Left            =   4320
         List            =   "frmReportUS.frx":0F50
         TabIndex        =   115
         Tag             =   "节律:[value]"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   6
         ItemData        =   "frmReportUS.frx":0F64
         Left            =   960
         List            =   "frmReportUS.frx":0F74
         TabIndex        =   114
         Tag             =   "胎心:[value]"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   0
         ItemData        =   "frmReportUS.frx":0F88
         Left            =   960
         List            =   "frmReportUS.frx":0F98
         TabIndex        =   108
         Tag             =   "胎头位置:[value]"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cbx产科情况 
         Height          =   300
         Index           =   3
         ItemData        =   "frmReportUS.frx":0FB0
         Left            =   960
         List            =   "frmReportUS.frx":0FC3
         TabIndex        =   111
         Tag             =   "脊柱位置:[value]"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lable1 
         Caption         =   "羊水:左上      cm  左下      cm  右上      cm  右下      cm"
         Height          =   255
         Index           =   104
         Left            =   120
         TabIndex        =   420
         Top             =   2420
         Width           =   5415
      End
      Begin VB.Label lable1 
         Caption         =   "低置距宫内口        cm"
         Height          =   255
         Index           =   103
         Left            =   120
         TabIndex        =   419
         Top             =   2080
         Width           =   2175
      End
      Begin VB.Label lable1 
         Caption         =   "脐动脉血流指数:PS     cm/s ED     cm/S RI     PI     A/B"
         Height          =   255
         Index           =   102
         Left            =   120
         TabIndex        =   418
         Top             =   2790
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "前置胎盘"
         Height          =   255
         Index           =   109
         Left            =   120
         TabIndex        =   417
         Top             =   1720
         Width           =   855
      End
      Begin VB.Label lable1 
         Caption         =   "胎盘位置                厚度      cm  胎盘级别"
         Height          =   255
         Index           =   110
         Left            =   120
         TabIndex        =   416
         Top             =   1360
         Width           =   4335
      End
      Begin VB.Label lable1 
         Caption         =   "    胎心                                  节律"
         Height          =   255
         Index           =   111
         Left            =   120
         TabIndex        =   415
         Top             =   1005
         Width           =   5295
      End
      Begin VB.Label lable1 
         Caption         =   "脊柱位置                腹围      cm    股骨长       cm"
         Height          =   255
         Index           =   112
         Left            =   120
         TabIndex        =   414
         Top             =   640
         Width           =   5175
      End
      Begin VB.Label lable1 
         Caption         =   "胎头位置              双顶径      cm      头围       cm"
         Height          =   255
         Index           =   113
         Left            =   120
         TabIndex        =   413
         Top             =   270
         Width           =   5055
      End
   End
   Begin VB.Frame framType 
      Caption         =   "妇  科"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "胎  囊"
         Height          =   975
         Index           =   18
         Left            =   120
         TabIndex        =   409
         Top             =   2400
         Width           =   5655
         Begin VB.ComboBox cbx胎囊 
            Height          =   300
            Index           =   2
            Left            =   2760
            Style           =   1  'Simple Combo
            TabIndex        =   101
            Tag             =   "胚芽:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胎囊 
            Height          =   300
            Index           =   1
            Left            =   1480
            Style           =   1  'Simple Combo
            TabIndex        =   100
            Tag             =   "宽:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx胎囊 
            Height          =   300
            Index           =   0
            Left            =   360
            Style           =   1  'Simple Combo
            TabIndex        =   99
            Tag             =   "长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "椭圆"
            Height          =   255
            Index           =   6
            Left            =   2560
            TabIndex        =   106
            Tag             =   "[T2]椭圆[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "圆形"
            Height          =   255
            Index           =   5
            Left            =   1900
            TabIndex        =   105
            Tag             =   "[T2]圆形[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "变形"
            Height          =   255
            Index           =   4
            Left            =   1230
            TabIndex        =   104
            Tag             =   "[T2]变形[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "光滑"
            Height          =   255
            Index           =   3
            Left            =   580
            TabIndex        =   103
            Tag             =   "[T2]光滑[value]"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "胚芽未见"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   102
            Tag             =   "胚芽未见[value]"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cbx妇科原始心管搏动 
            Height          =   300
            ItemData        =   "frmReportUS.frx":0FE1
            Left            =   4560
            List            =   "frmReportUS.frx":0FF1
            TabIndex        =   107
            Tag             =   "原始心管搏动:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "长       cm  宽      cm  胚芽      cm"
            Height          =   255
            Index           =   101
            Left            =   120
            TabIndex        =   411
            Top             =   285
            Width           =   3495
         End
         Begin VB.Label lable1 
            Caption         =   "囊壁                                 原始心管搏动"
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   410
            Top             =   645
            Width           =   5445
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "右卵巢"
         Height          =   615
         Index           =   17
         Left            =   120
         TabIndex        =   407
         Top             =   1800
         Width           =   5655
         Begin VB.ComboBox cbx右卵巢 
            Height          =   300
            Index           =   1
            Left            =   1440
            Style           =   1  'Simple Combo
            TabIndex        =   97
            Tag             =   "X [value]cm^2"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx右卵巢 
            Height          =   300
            Index           =   0
            Left            =   720
            Style           =   1  'Simple Combo
            TabIndex        =   96
            Tag             =   "右侧:[value]"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "未显示"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   98
            Tag             =   "未显示[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "右侧         X       cm^2"
            Height          =   255
            Index           =   96
            Left            =   120
            TabIndex        =   408
            Top             =   285
            Width           =   2535
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "左卵巢"
         Height          =   615
         Index           =   16
         Left            =   120
         TabIndex        =   405
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cbx左卵巢 
            Height          =   300
            Index           =   1
            Left            =   1440
            Style           =   1  'Simple Combo
            TabIndex        =   94
            Tag             =   "X [value]cm^2"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx左卵巢 
            Height          =   300
            Index           =   0
            Left            =   720
            Style           =   1  'Simple Combo
            TabIndex        =   93
            Tag             =   "左侧:[value]"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chk妇科 
            Caption         =   "未显示"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   95
            Tag             =   "未显示[value]"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "左侧         X       cm^2"
            Height          =   255
            Index           =   98
            Left            =   120
            TabIndex        =   406
            Top             =   285
            Width           =   2415
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "子宫情况"
         Height          =   975
         Index           =   15
         Left            =   120
         TabIndex        =   402
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cbx子宫 
            Height          =   300
            Index           =   2
            Left            =   4080
            Style           =   1  'Simple Combo
            TabIndex        =   89
            Tag             =   "宽:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx子宫 
            Height          =   300
            Index           =   1
            Left            =   2160
            Style           =   1  'Simple Combo
            TabIndex        =   88
            Tag             =   "厚:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx子宫 
            Height          =   300
            Index           =   0
            Left            =   600
            Style           =   1  'Simple Combo
            TabIndex        =   87
            Tag             =   "长:[value]cm"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cbx子宫 
            Height          =   300
            Index           =   5
            ItemData        =   "frmReportUS.frx":1005
            Left            =   4080
            List            =   "frmReportUS.frx":1012
            TabIndex        =   92
            Tag             =   "宫腔反射:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx子宫 
            Height          =   300
            Index           =   4
            ItemData        =   "frmReportUS.frx":1026
            Left            =   2160
            List            =   "frmReportUS.frx":1036
            TabIndex        =   91
            Tag             =   "位置:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbx子宫 
            Height          =   300
            Index           =   3
            ItemData        =   "frmReportUS.frx":104E
            Left            =   600
            List            =   "frmReportUS.frx":105B
            TabIndex        =   90
            Tag             =   "形态:[value]"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "形态              位置             宫腔反射"
            Height          =   255
            Index           =   99
            Left            =   120
            TabIndex        =   404
            Top             =   645
            Width           =   4965
         End
         Begin VB.Label lable1 
            Caption         =   "  长       cm       厚       cm          宽        cm"
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
      Caption         =   "心  脏"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   32
         Left            =   3360
         Style           =   1  'Simple Combo
         TabIndex        =   85
         Tag             =   "HR:[value]次/分"
         Top             =   4200
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   31
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   84
         Tag             =   "[T1]CI:[value]L/ml/m^2"
         Top             =   4200
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   30
         Left            =   4200
         Style           =   1  'Simple Combo
         TabIndex        =   83
         Tag             =   "[T1]SI:[value]ml/m^2"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   29
         Left            =   3000
         Style           =   1  'Simple Combo
         TabIndex        =   82
         Tag             =   "[T1]SV:[value]ml"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   28
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   81
         Tag             =   "[T1]LVVD:[value]cm^3"
         Top             =   3840
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   27
         Left            =   3840
         Style           =   1  'Simple Combo
         TabIndex        =   80
         Tag             =   "[T1]CO:[value]L/Min"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   26
         Left            =   2760
         Style           =   1  'Simple Combo
         TabIndex        =   79
         Tag             =   "[T1]FS:[value]%"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   25
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   78
         Tag             =   "[T1]EF:[value]%"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   24
         Left            =   4920
         Style           =   1  'Simple Combo
         TabIndex        =   77
         Tag             =   "大小:[value]"
         Top             =   3120
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   22
         Left            =   480
         Style           =   1  'Simple Combo
         TabIndex        =   75
         Tag             =   "面积:[value]cm^2"
         Top             =   3120
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   21
         Left            =   4920
         Style           =   1  'Simple Combo
         TabIndex        =   74
         Tag             =   "主动脉开口(长轴):[value]mm"
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   20
         Left            =   2400
         Style           =   1  'Simple Combo
         TabIndex        =   73
         Tag             =   "面积:[value]cm^2"
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   19
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   72
         Tag             =   "开口(长轴):[value]mm"
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   15
         Left            =   5010
         Style           =   1  'Simple Combo
         TabIndex        =   68
         Tag             =   "左室后壁动度:[value]mm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   14
         Left            =   3840
         Style           =   1  'Simple Combo
         TabIndex        =   67
         Tag             =   "左室后壁厚度:[value]mm"
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   12
         Left            =   4560
         Style           =   1  'Simple Combo
         TabIndex        =   65
         Tag             =   "室间隔动度:[value]mm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   11
         Left            =   3000
         Style           =   1  'Simple Combo
         TabIndex        =   64
         Tag             =   "室间隔厚度:[value]mm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   10
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   63
         Tag             =   "左室前后径:[value]mm"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   9
         Left            =   4540
         Style           =   1  'Simple Combo
         TabIndex        =   62
         Tag             =   "右室长径:[value]mm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   8
         Left            =   2760
         Style           =   1  'Simple Combo
         TabIndex        =   61
         Tag             =   "右室横径:[value]mm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   7
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   60
         Tag             =   "右室前后径:[value]mm"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   6
         Left            =   4540
         Style           =   1  'Simple Combo
         TabIndex        =   59
         Tag             =   "右房长径:[value]mm"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   5
         Left            =   2760
         Style           =   1  'Simple Combo
         TabIndex        =   58
         Tag             =   "右房横径:[value]mm"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   4
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   57
         Tag             =   "左房前后径:[value]mm"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   3
         Left            =   4200
         Style           =   1  'Simple Combo
         TabIndex        =   56
         Tag             =   "主动脉窦部内径:[value]mm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   2
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   55
         Tag             =   "主动脉窦部内径:[value]mm"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   1
         Left            =   4200
         Style           =   1  'Simple Combo
         TabIndex        =   54
         Tag             =   "肺动脉干内径:[value]mm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   0
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   53
         Tag             =   "右室流出道内径:[value]mm"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   33
         ItemData        =   "frmReportUS.frx":106F
         Left            =   4920
         List            =   "frmReportUS.frx":107C
         TabIndex        =   86
         Tag             =   "心律:[value]"
         Top             =   4200
         Width           =   855
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   23
         ItemData        =   "frmReportUS.frx":108C
         Left            =   2040
         List            =   "frmReportUS.frx":109C
         TabIndex        =   76
         Tag             =   "主动脉:[value]"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   13
         ItemData        =   "frmReportUS.frx":10C0
         Left            =   1800
         List            =   "frmReportUS.frx":10CD
         TabIndex        =   66
         Tag             =   "室间隔与左室后壁呈:[value]"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   18
         ItemData        =   "frmReportUS.frx":10E7
         Left            =   4920
         List            =   "frmReportUS.frx":10F4
         TabIndex        =   71
         Tag             =   "后叶呈:[value]"
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   17
         ItemData        =   "frmReportUS.frx":1106
         Left            =   3360
         List            =   "frmReportUS.frx":1113
         TabIndex        =   70
         Tag             =   "二尖瓣前叶活动曲线呈:[value]"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cbx心脏 
         Height          =   300
         Index           =   16
         ItemData        =   "frmReportUS.frx":1129
         Left            =   480
         List            =   "frmReportUS.frx":1139
         TabIndex        =   69
         Tag             =   "搏动:[value]"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lable1 
         Caption         =   "CI       L/ml/m^2  HR      次/分  心律"
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
         Caption         =   "左心收缩功能：EF       %  FS       %  CO       L/Min "
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   398
         Top             =   3530
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "面积      cm^2 主动脉              增强回声光斑  大小       cm"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   397
         Top             =   3160
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "开口(长轴)       mm  面积       cm^2 主动脉开口(长轴)       mm"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   396
         Top             =   2780
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "室间隔与左室后壁呈           左室后壁厚度      mm 动度      mm"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   395
         Top             =   2070
         Width           =   5655
      End
      Begin VB.Label lable1 
         Caption         =   "搏动            二尖瓣前叶活动曲线呈           后叶呈"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   394
         Top             =   2430
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "左室前后径       mm  室间隔厚度       mm    动度       mm"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   393
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "右室前后径       mm  右室横径       mm  右室长径       mm"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   392
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "左房前后径       mm  右房横径       mm  右房长径       mm"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   391
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label lable1 
         Caption         =   "主动脉窦部内径         mm       升主动脉内径        mm"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   390
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lable1 
         Caption         =   "右室流出道内径         mm       肺动脉干内径        mm "
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
      Caption         =   "多普勒超声"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "肺动脉瓣"
         Height          =   1935
         Index           =   8
         Left            =   3120
         TabIndex        =   315
         Top             =   2400
         Width           =   2655
         Begin VB.TextBox txt肺动脉瓣峰值流速 
            Height          =   270
            Left            =   1200
            TabIndex        =   52
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣反流压差 
            Height          =   270
            Left            =   1200
            TabIndex        =   51
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣射血时间 
            Height          =   270
            Left            =   1200
            TabIndex        =   50
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣加速时间 
            Height          =   270
            Left            =   1200
            TabIndex        =   49
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣舒张期压差 
            Height          =   270
            Left            =   1680
            TabIndex        =   48
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣舒张期流速 
            Height          =   270
            Left            =   1200
            TabIndex        =   47
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣收缩期压差 
            Height          =   270
            Left            =   1680
            TabIndex        =   46
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt肺动脉瓣收缩期流速 
            Height          =   270
            Left            =   1200
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "峰值流速          cm/s"
            Height          =   255
            Index           =   92
            Left            =   120
            TabIndex        =   340
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "反流压差          mmHg"
            Height          =   255
            Index           =   91
            Left            =   120
            TabIndex        =   326
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "射血时间          ms"
            Height          =   255
            Index           =   90
            Left            =   120
            TabIndex        =   325
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "加速时间          ms"
            Height          =   255
            Index           =   89
            Left            =   120
            TabIndex        =   324
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "舒张期"
            Height          =   255
            Index           =   88
            Left            =   120
            TabIndex        =   323
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "收缩期"
            Height          =   255
            Index           =   87
            Left            =   120
            TabIndex        =   316
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "主动脉瓣"
         Height          =   1095
         Index           =   6
         Left            =   120
         TabIndex        =   313
         Top             =   3240
         Width           =   2775
         Begin VB.TextBox txt主动脉瓣流速 
            Height          =   270
            Left            =   1440
            TabIndex        =   37
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt主动脉瓣舒张期流速 
            Height          =   270
            Left            =   1440
            TabIndex        =   35
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt主动脉瓣舒张期压差 
            Height          =   270
            Left            =   1920
            TabIndex        =   36
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt主动脉瓣收缩期流速 
            Height          =   270
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt主动脉瓣收缩期压差 
            Height          =   270
            Left            =   1920
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "流速"
            Height          =   255
            Index           =   82
            Left            =   120
            TabIndex        =   339
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "舒张期"
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   322
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "收缩期"
            Height          =   255
            Index           =   80
            Left            =   120
            TabIndex        =   314
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "三尖瓣"
         Height          =   1575
         Index           =   7
         Left            =   3120
         TabIndex        =   311
         Top             =   720
         Width           =   2655
         Begin VB.TextBox txt三尖瓣反流压差 
            Height          =   270
            Left            =   1200
            TabIndex        =   44
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txt三尖瓣收缩期流速 
            Height          =   270
            Left            =   1200
            TabIndex        =   42
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt三尖瓣收缩期压差 
            Height          =   270
            Left            =   1680
            TabIndex        =   43
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt三尖瓣舒张期A峰流速 
            Height          =   270
            Left            =   1200
            TabIndex        =   40
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt三尖瓣舒张期A峰压差 
            Height          =   270
            Left            =   1680
            TabIndex        =   41
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt三尖瓣舒张期E峰流速 
            Height          =   270
            Left            =   1200
            TabIndex        =   38
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt三尖瓣舒张期E峰压差 
            Height          =   270
            Left            =   1680
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "反流压差          mmHg"
            Height          =   255
            Index           =   86
            Left            =   120
            TabIndex        =   321
            Top             =   1110
            Width           =   2055
         End
         Begin VB.Label lable1 
            Caption         =   "收缩期"
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   320
            Top             =   870
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "舒张期A峰"
            Height          =   255
            Index           =   84
            Left            =   120
            TabIndex        =   319
            Top             =   630
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "舒张期E峰"
            Height          =   255
            Index           =   83
            Left            =   120
            TabIndex        =   312
            Top             =   390
            Width           =   975
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "二尖瓣"
         Height          =   2415
         Index           =   5
         Left            =   120
         TabIndex        =   309
         Top             =   720
         Width           =   2775
         Begin VB.TextBox txt二尖瓣E波减速时间 
            Height          =   270
            Left            =   1320
            TabIndex        =   32
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣E波减速度 
            Height          =   270
            Left            =   1320
            TabIndex        =   31
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣等容舒张时间 
            Height          =   270
            Left            =   1320
            TabIndex        =   30
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣收缩期压差 
            Height          =   270
            Left            =   1800
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣收缩期流速 
            Height          =   270
            Left            =   1320
            TabIndex        =   28
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣舒张期A峰压差 
            Height          =   270
            Left            =   1800
            TabIndex        =   27
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣舒张期A峰流速 
            Height          =   270
            Left            =   1320
            TabIndex        =   26
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣舒张期E峰压差 
            Height          =   270
            Left            =   1800
            TabIndex        =   25
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt二尖瓣舒张期E峰流速 
            Height          =   270
            Left            =   1320
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "E波减速时间        196                       ±27ms"
            Height          =   375
            Index           =   79
            Left            =   120
            TabIndex        =   329
            Top             =   1920
            Width           =   2280
         End
         Begin VB.Label lable1 
            Caption         =   "E波减速度          >150                     cm/s"
            Height          =   375
            Index           =   78
            Left            =   120
            TabIndex        =   328
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "等容舒张时间       <90ms"
            Height          =   255
            Index           =   77
            Left            =   120
            TabIndex        =   327
            Top             =   1110
            Width           =   2250
         End
         Begin VB.Label lable1 
            Caption         =   "收缩期"
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   318
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "舒张期A峰"
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   317
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lable1 
            Caption         =   "舒张期E峰"
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
         Caption         =   "  项目        流速  压差              项目    流速  压差"
         Height          =   255
         Index           =   94
         Left            =   120
         TabIndex        =   308
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame framType 
      Caption         =   "二维及M型超声"
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "左心室"
         Height          =   2055
         Index           =   1
         Left            =   120
         TabIndex        =   331
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txt左心室射血分数 
            Height          =   270
            Left            =   1320
            TabIndex        =   11
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txt左心室收缩末容积 
            Height          =   270
            Left            =   1320
            TabIndex        =   283
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txt左心室舒张末容积 
            Height          =   270
            Left            =   1320
            TabIndex        =   10
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt左心室运动幅度 
            Height          =   270
            Left            =   1320
            TabIndex        =   9
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt左心室后壁厚度 
            Height          =   270
            Left            =   1320
            TabIndex        =   8
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt左心室舒张末期径 
            Height          =   270
            Left            =   1320
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt左心室每博量 
            Height          =   270
            Left            =   1320
            TabIndex        =   12
            Top             =   1690
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "每博量(ml)         35-90"
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   332
            Top             =   1710
            Width           =   2175
         End
         Begin VB.Label Label51 
            Caption         =   "舒张末容积 ml      60-125"
            Height          =   255
            Left            =   120
            TabIndex        =   338
            Top             =   990
            Width           =   2295
         End
         Begin VB.Label lable1 
            Caption         =   "射血分数EF         >50%"
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   337
            Top             =   1470
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "收缩末容积 ml      30-40"
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   336
            Top             =   1230
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "运动幅度           8-15"
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   335
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "后壁厚度           8-12"
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   334
            Top             =   510
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "舒张末期径         <55/50"
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   333
            Top             =   270
            Width           =   2295
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "右心室"
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   289
         Top             =   480
         Width           =   2655
         Begin VB.TextBox txt右心室射血分数 
            Height          =   270
            Left            =   1320
            TabIndex        =   6
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txt右心室收缩末容积 
            Height          =   270
            Left            =   1320
            TabIndex        =   5
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txt右心室舒张末容积 
            Height          =   270
            Left            =   1320
            TabIndex        =   4
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt右心室横径 
            Height          =   270
            Left            =   1320
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt右心室前壁厚度 
            Height          =   270
            Left            =   1320
            TabIndex        =   2
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt右心室前后径 
            Height          =   270
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "射血分数EF          >50%"
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   307
            Top             =   1470
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "收缩末容积 ml"
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   306
            Top             =   1230
            Width           =   2055
         End
         Begin VB.Label lable1 
            Caption         =   "舒张末容积 ml"
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   305
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lable1 
            Caption         =   "横径                <40"
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   304
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "前壁厚度            <5"
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   303
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lable1 
            Caption         =   "前后径              <25"
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
         Begin VB.TextBox txt左肺动脉内径 
            Height          =   270
            Left            =   1440
            TabIndex        =   17
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txt右肺动脉内径 
            Height          =   270
            Left            =   1440
            TabIndex        =   16
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txt主肺动脉内径 
            Height          =   270
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt升动脉内径 
            Height          =   270
            Left            =   1440
            TabIndex        =   14
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt主动脉根部内径 
            Height          =   270
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "左肺动脉内径        <18"
            Height          =   255
            Index           =   67
            Left            =   120
            TabIndex        =   295
            Top             =   1230
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "右肺动脉内径        <18"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   294
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "主肺动脉内径        <25"
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   293
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "升主动脉内径        <35"
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   292
            Top             =   510
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "主动脉根部内径      <35"
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   291
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "室间隔"
         Height          =   1095
         Index           =   3
         Left            =   3000
         TabIndex        =   287
         Top             =   2070
         Width           =   2775
         Begin VB.TextBox txt室间隔与左室后壁方向 
            Height          =   270
            Left            =   1440
            TabIndex        =   20
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt室间隔运动幅度 
            Height          =   270
            Left            =   1440
            TabIndex        =   19
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt室间隔厚度 
            Height          =   270
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "与左室后壁方向       逆向"
            Height          =   255
            Index           =   70
            Left            =   120
            TabIndex        =   301
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lable1 
            Caption         =   "运动幅度             5-8"
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   300
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "厚度                 8-12"
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   299
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame framGroup 
         Caption         =   "心房"
         Height          =   1170
         Index           =   4
         Left            =   3000
         TabIndex        =   286
         Top             =   3170
         Width           =   2775
         Begin VB.TextBox txt心房右房横径 
            Height          =   270
            Left            =   1440
            TabIndex        =   23
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt心房右房长径 
            Height          =   270
            Left            =   1440
            TabIndex        =   22
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt心房左房前后径 
            Height          =   270
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lable1 
            Caption         =   "右房横径            <40"
            Height          =   255
            Index           =   73
            Left            =   120
            TabIndex        =   298
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "右房长径            <50"
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   297
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lable1 
            Caption         =   "左房前后径          <37"
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   296
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label lable1 
         Caption         =   "  项目      测值  正常(mm)         项目      测值 正常(mm)"
         Height          =   255
         Index           =   93
         Left            =   240
         TabIndex        =   290
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Label Label21 
      Caption         =   "专科项目："
      BeginProperty Font 
         Name            =   "宋体"
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

Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mlngAdviceID As Long    '医嘱ID
Private mintEditType As Integer '病历状态 0 创建，1书写，2 修订
Private mReportID As Long       '报告ID
Private mblnCheckModify As Boolean      '是否启动内容变化记录
Private mblnEditable As Boolean         '是否可以编辑内容
Private mblnMoved As Boolean            '是否已经转储

Private Const Report_Element_右心室前后径 = "右心室前后径"
Private Const Report_Element_右心室前壁厚度 = "右心室前壁厚度"
Private Const Report_Element_右心室横径 = "右心室横径"
Private Const Report_Element_右心室舒张末容积 = "右心室舒张末容积"
Private Const Report_Element_右心室收缩末容积 = "右心室收缩末容积"
Private Const Report_Element_右心室射血分数 = "右心室射血分数"

Private Const Report_Element_左心室舒张末期径 = "左心室舒张末期径"
Private Const Report_Element_左心室后壁厚度 = "左心室后壁厚度"
Private Const Report_Element_左心室运动幅度 = "左心室运动幅度"
Private Const Report_Element_左心室舒张末容积 = "左心室舒张末容积"
Private Const Report_Element_左心室收缩末容积 = "左心室收缩末容积"
Private Const Report_Element_左心室射血分数 = "左心室射血分数"
Private Const Report_Element_左心室每博量 = "左心室每博量"

Private Const Report_Element_主动脉根部内径 = "主动脉根部内径"
Private Const Report_Element_升动脉内径 = "升主动脉内径"
Private Const Report_Element_主肺动脉内径 = "主肺动脉内径"
Private Const Report_Element_右肺动脉内径 = "右肺动脉内径"
Private Const Report_Element_左肺动脉内径 = "左肺动脉内径"

Private Const Report_Element_室间隔厚度 = "室间隔厚度"
Private Const Report_Element_室间隔运动幅度 = "室间隔运动幅度"
Private Const Report_Element_室间隔与左室后壁方向 = "室间隔与左室后壁方向"

Private Const Report_Element_心房左房前后径 = "心房左房前后径"
Private Const Report_Element_心房右房长径 = "心房右房长径"
Private Const Report_Element_心房右房横径 = "心房右房横径"


Private Const Report_Element_二尖瓣舒张期E峰流速 = "二尖瓣舒张期E峰流速"
Private Const Report_Element_二尖瓣舒张期E峰压差 = "二尖瓣舒张期E峰压差"
Private Const Report_Element_二尖瓣舒张期A峰流速 = "二尖瓣舒张期A峰流速"
Private Const Report_Element_二尖瓣舒张期A峰压差 = "二尖瓣舒张期A峰压差"
Private Const Report_Element_二尖瓣收缩期流速 = "二尖瓣收缩期流速"
Private Const Report_Element_二尖瓣收缩期压差 = "二尖瓣收缩期压差"
Private Const Report_Element_二尖瓣等容舒张时间 = "二尖瓣等容舒张时间"
Private Const Report_Element_二尖瓣E波减速度 = "二尖瓣E波减速度"
Private Const Report_Element_二尖瓣E波减速时间 = "二尖瓣E波减速时间"

Private Const Report_Element_主动脉瓣收缩期流速 = "主动脉瓣收缩期流速"
Private Const Report_Element_主动脉瓣收缩期压差 = "主动脉瓣收缩期压差"
Private Const Report_Element_主动脉瓣舒张期流速 = "主动脉瓣舒张期流速"
Private Const Report_Element_主动脉瓣舒张期压差 = "主动脉瓣舒张期压差"
Private Const Report_Element_主动脉瓣流速 = "主动脉瓣流速"

Private Const Report_Element_三尖瓣舒张期E峰流速 = "三尖瓣舒张期E峰流速"
Private Const Report_Element_三尖瓣舒张期E峰压差 = "三尖瓣舒张期E峰压差"
Private Const Report_Element_三尖瓣舒张期A峰流速 = "三尖瓣舒张期A峰流速"
Private Const Report_Element_三尖瓣舒张期A峰压差 = "三尖瓣舒张期A峰压差"
Private Const Report_Element_三尖瓣收缩期流速 = "三尖瓣收缩期流速"
Private Const Report_Element_三尖瓣收缩期压差 = "三尖瓣收缩期压差"
Private Const Report_Element_三尖瓣反流压差 = "三尖瓣反流压差"

Private Const Report_Element_肺动脉瓣收缩期流速 = "肺动脉瓣收缩期流速"
Private Const Report_Element_肺动脉瓣收缩期压差 = "肺动脉瓣收缩期压差"
Private Const Report_Element_肺动脉瓣舒张期流速 = "肺动脉瓣舒张期流速"
Private Const Report_Element_肺动脉瓣舒张期压差 = "肺动脉瓣舒张期压差"
Private Const Report_Element_肺动脉瓣加速时间 = "肺动脉瓣加速时间"
Private Const Report_Element_肺动脉瓣射血时间 = "肺动脉瓣射血时间"
Private Const Report_Element_肺动脉瓣反流压差 = "肺动脉瓣反流压差"
Private Const Report_Element_肺动脉瓣峰值流速 = "肺动脉瓣峰值流速"

Private Const Report_Element_专科报告 = "专科报告"
Private Const Report_CheckedValue = " "
Private Const Report_ProjectSplitChr = vbCrLf '"      "




Public pModified As Boolean     '记录是否有修改




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
    
    mblnCheckModify = False         '关闭内容修改记录
    '清除修改标记
    pModified = False
    
    '清空所有内容
    txt右心室前后径.Text = ""
    txt右心室前壁厚度.Text = ""
    txt右心室横径.Text = ""
    txt右心室舒张末容积.Text = ""
    txt右心室收缩末容积.Text = ""
    txt右心室射血分数.Text = ""
    
    txt左心室舒张末期径.Text = ""
    txt左心室后壁厚度.Text = ""
    txt左心室运动幅度.Text = ""
    txt左心室舒张末容积.Text = ""
    txt左心室收缩末容积.Text = ""
    txt左心室射血分数.Text = ""
    txt左心室每博量.Text = ""
    
    txt主动脉根部内径.Text = ""
    txt升动脉内径.Text = ""
    txt主肺动脉内径.Text = ""
    txt右肺动脉内径.Text = ""
    txt左肺动脉内径.Text = ""
    
    txt室间隔厚度.Text = ""
    txt室间隔运动幅度.Text = ""
    txt室间隔与左室后壁方向.Text = ""
    
    txt心房左房前后径.Text = ""
    txt心房右房长径.Text = ""
    txt心房右房横径.Text = ""
    
    txt二尖瓣舒张期E峰流速.Text = ""
    txt二尖瓣舒张期E峰压差.Text = ""
    txt二尖瓣舒张期A峰流速.Text = ""
    txt二尖瓣舒张期A峰压差.Text = ""
    txt二尖瓣收缩期流速.Text = ""
    txt二尖瓣收缩期压差.Text = ""
    txt二尖瓣等容舒张时间.Text = ""
    txt二尖瓣E波减速度.Text = ""
    txt二尖瓣E波减速时间.Text = ""
    
    txt主动脉瓣收缩期流速.Text = ""
    txt主动脉瓣收缩期压差.Text = ""
    txt主动脉瓣舒张期流速.Text = ""
    txt主动脉瓣舒张期压差.Text = ""
    txt主动脉瓣流速.Text = ""
    
    txt三尖瓣舒张期E峰流速.Text = ""
    txt三尖瓣舒张期E峰压差.Text = ""
    txt三尖瓣舒张期A峰流速.Text = ""
    txt三尖瓣舒张期A峰压差.Text = ""
    txt三尖瓣收缩期流速.Text = ""
    txt三尖瓣收缩期压差.Text = ""
    txt三尖瓣反流压差.Text = ""
    
    txt肺动脉瓣收缩期流速.Text = ""
    txt肺动脉瓣收缩期压差.Text = ""
    txt肺动脉瓣舒张期流速.Text = ""
    txt肺动脉瓣舒张期压差.Text = ""
    txt肺动脉瓣加速时间.Text = ""
    txt肺动脉瓣射血时间.Text = ""
    txt肺动脉瓣反流压差.Text = ""
    txt肺动脉瓣峰值流速.Text = ""
    
    '清空心脏数据
    For i = 0 To 33: cbx心脏(i).Text = "": Next i
    
    '清空妇科数据
    For i = 0 To 5: cbx子宫(i).Text = "": Next i
    
    cbx左卵巢(0).Text = ""
    cbx左卵巢(1).Text = ""
    cbx右卵巢(0).Text = ""
    cbx右卵巢(1).Text = ""
    
    cbx妇科原始心管搏动.Text = ""
    
    For i = 0 To 6: chk妇科(i).value = 0: Next i
    
    '清空产科数据
    For i = 0 To 21: cbx产科情况(i).Text = "": Next i
    For i = 0 To 3: chk产科情况(i).value = 0: Next i
    
    '清空腹部数据
    For i = 0 To 4: cbx肝脏情况(i).Text = "": Next i
    For i = 0 To 5: chk肝脏情况(i).value = 0: Next i
    For i = 0 To 1: cbx血管(i).Text = "": Next i
    For i = 0 To 12: cbx胆囊情况(i).Text = "": Next i
    For i = 0 To 5: cbx胆总管(i).Text = "": Next i
    For i = 0 To 6: cbx胰腺(i).Text = "": Next i
    For i = 0 To 4: cbx脾脏(i).Text = "": Next i
    
    '清空泌尿男数据
    For i = 0 To 5: txt肾脏情况(i).Text = "": Next i
    For i = 0 To 5: cbx泌尿男(i).Text = "":  Next i
    
    cbx膀胱(0).Text = ""
    chk泌尿男(0).value = 0: chk泌尿男(1).value = 0
    
    For i = 0 To 2: txt前列腺信息(i).Text = "": Next i
    
    '清空泌尿女数据
    For i = 0 To 5: txt女肾脏情况(i).Text = "": Next i
    For i = 0 To 5: cbx输尿管(i).Text = "":  Next i
    
    cbx女膀胱(0).Text = ""
    chk泌尿女(0).value = 0: chk泌尿女(1).value = 0
    
    '清空颈部数据
    For i = 0 To 13: txt颈部信息(i).Text = "":  Next i
    
    '清空乳腺数据
    For i = 0 To 3: txt乳腺信息(i).Text = "":  Next i
    
    '清空生殖器数据
    For i = 0 To 7: cbx左侧生殖器(i).Text = "": Next i
    For i = 0 To 7: cbx右侧生殖器(i).Text = "": Next i
    
    '清空下肢静脉数据
    For i = 0 To 4: chk左下肢静脉(i).value = 0: Next i
    For i = 0 To 4: chk右下肢静脉(i).value = 0: Next i
    
    '清除眼部数据
    For i = 0 To 1: txt眼部信息(i).Text = "": Next i
    
    '清空甲状腺数据
    For i = 0 To 9: txt甲状腺信息(i).Text = "": Next i
    
    '清空肾动脉数据
    For i = 0 To 11: txt肾动脉信息(i).Text = "": Next i
    
    '清空胸腔数据
    For i = 0 To 5: cbx胸腔信息(i).Text = "": Next i
    For i = 0 To 1: chk胸腔信息(i).value = 0: Next i
    
    
    strSql = "Select 内容文本,要素名称 From 电子病历内容 Where 文件ID=[1] And 对象类型=4 And 终止版=0 And 替换域=0"
    If mblnMoved = True Then
        strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
    
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!要素名称)
            Case Report_Element_右心室前后径
                txt右心室前后径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_右心室前壁厚度
                txt右心室前壁厚度.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_右心室横径
                txt右心室横径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_右心室舒张末容积
                txt右心室舒张末容积.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_右心室收缩末容积
                txt右心室收缩末容积.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_右心室射血分数
                txt右心室射血分数.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_左心室舒张末期径
                txt左心室舒张末期径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左心室后壁厚度
                txt左心室后壁厚度.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左心室运动幅度
                txt左心室运动幅度.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左心室舒张末容积
                txt左心室舒张末容积.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左心室收缩末容积
                txt左心室收缩末容积.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左心室射血分数
                txt左心室射血分数.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左心室每博量
                txt左心室每博量.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_主动脉根部内径
                txt主动脉根部内径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_升动脉内径
                txt升动脉内径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_主肺动脉内径
                txt主肺动脉内径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_右肺动脉内径
                txt右肺动脉内径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_左肺动脉内径
                txt左肺动脉内径.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_室间隔厚度
                txt室间隔厚度.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_室间隔运动幅度
                txt室间隔运动幅度.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_室间隔与左室后壁方向
                txt室间隔与左室后壁方向.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_心房左房前后径
                txt心房左房前后径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_心房右房长径
                txt心房右房长径.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_心房右房横径
                txt心房右房横径.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_二尖瓣舒张期E峰流速
                txt二尖瓣舒张期E峰流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣舒张期E峰压差
                txt二尖瓣舒张期E峰压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣舒张期A峰流速
                txt二尖瓣舒张期A峰流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣舒张期A峰压差
                txt二尖瓣舒张期A峰压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣收缩期流速
                txt二尖瓣收缩期流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣收缩期压差
                txt二尖瓣收缩期压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣等容舒张时间
                txt二尖瓣等容舒张时间.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣E波减速度
                txt二尖瓣E波减速度.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_二尖瓣E波减速时间
                txt二尖瓣E波减速时间.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_主动脉瓣收缩期流速
                txt主动脉瓣收缩期流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_主动脉瓣收缩期压差
                txt主动脉瓣收缩期压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_主动脉瓣舒张期流速
                txt主动脉瓣舒张期流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_主动脉瓣舒张期压差
                txt主动脉瓣舒张期压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_主动脉瓣流速
                txt主动脉瓣流速.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_三尖瓣舒张期E峰流速
                txt三尖瓣舒张期E峰流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_三尖瓣舒张期E峰压差
                txt三尖瓣舒张期E峰压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_三尖瓣舒张期A峰流速
                txt三尖瓣舒张期A峰流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_三尖瓣舒张期A峰压差
                txt三尖瓣舒张期A峰压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_三尖瓣收缩期流速
                txt三尖瓣收缩期流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_三尖瓣收缩期压差
                txt三尖瓣收缩期压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_三尖瓣反流压差
                txt三尖瓣反流压差.Text = Nvl(rsTemp!内容文本)
                
            Case Report_Element_肺动脉瓣收缩期流速
                txt肺动脉瓣收缩期流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣收缩期压差
                txt肺动脉瓣收缩期压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣舒张期流速
                txt肺动脉瓣舒张期流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣舒张期压差
                txt肺动脉瓣舒张期压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣加速时间
                txt肺动脉瓣加速时间.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣射血时间
                txt肺动脉瓣射血时间.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣反流压差
                txt肺动脉瓣反流压差.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_肺动脉瓣峰值流速
                txt肺动脉瓣峰值流速.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_专科报告
                '获取专科报告中的内容
                Call LoadProfessionalReportData(Nvl(rsTemp!内容文本))
        End Select
        rsTemp.MoveNext
    Wend
    
    '设置界面控件是否可以编辑
    If cbxType.Text = "" Then
        cbxType.ListIndex = 0
        framType(0).Visible = True
    End If
    
    Call ConfigEditable(mblnEditable)
    
    mblnCheckModify = True         '打开内容修改记录
End Sub


Private Sub LoadProfessionalReportData(ByVal strProfessionalReport As String)
    Dim i As Long
    Dim strCurReport As String
    
    If strProfessionalReport = "" Then Exit Sub
    
    strCurReport = strProfessionalReport
    
    Select Case True
        Case InStr(strProfessionalReport, "【心脏】") >= 1
            For i = 0 To 33
                cbx心脏(i).Text = GetReportValue(strCurReport, "心脏", cbx心脏(i).Tag)
            Next i
            
            cbxType.ListIndex = 2
            framType(2).Visible = True
        Case InStr(strProfessionalReport, "【子宫情况】") >= 1 Or InStr(strProfessionalReport, "【左卵巢】") >= 1 _
            Or InStr(strProfessionalReport, "【右卵巢】") >= 1 Or InStr(strProfessionalReport, "【胎囊】") >= 1
            '子宫情况
            For i = 0 To 5
                cbx子宫(i).Text = GetReportValue(strCurReport, "子宫情况", cbx子宫(i).Tag)
            Next i
            
            '左卵巢
            cbx左卵巢(0).Text = GetReportValue(strCurReport, "左卵巢", cbx左卵巢(0).Tag)
            cbx左卵巢(1).Text = GetReportValue(strCurReport, "左卵巢", cbx左卵巢(1).Tag)
            chk妇科(0).value = IIf(GetReportValue(strCurReport, "左卵巢", chk妇科(0).Tag) = Report_CheckedValue, 1, 0)
            
            '右卵巢
            cbx右卵巢(0).Text = GetReportValue(strCurReport, "右卵巢", cbx右卵巢(0).Tag)
            cbx右卵巢(1).Text = GetReportValue(strCurReport, "右卵巢", cbx右卵巢(1).Tag)
            chk妇科(1).value = IIf(GetReportValue(strCurReport, "右卵巢", chk妇科(1).Tag) = Report_CheckedValue, 1, 0)
            
            '胎囊
            For i = 0 To 2
                cbx胎囊(i).Text = GetReportValue(strCurReport, "胎囊", cbx胎囊(i).Tag)
            Next i
            
            For i = 2 To 6
                chk妇科(i).value = IIf(GetReportValue(strCurReport, "胎囊", chk妇科(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbx妇科原始心管搏动.Text = GetReportValue(strCurReport, "胎囊", cbx妇科原始心管搏动.Tag)
            
            cbxType.ListIndex = 3
            framType(3).Visible = True
        Case InStr(strProfessionalReport, "【产科情况】") >= 1
            For i = 0 To 21
                cbx产科情况(i).Text = GetReportValue(strCurReport, "产科情况", cbx产科情况(i).Tag)
            Next i
            
            For i = 0 To 3
                chk产科情况(i).value = IIf(GetReportValue(strCurReport, "产科情况", chk产科情况(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbxType.ListIndex = 4
            framType(4).Visible = True
        Case InStr(strProfessionalReport, "【肝脏情况】") >= 1 Or InStr(strProfessionalReport, "【血管】") >= 1 _
            Or InStr(strProfessionalReport, "【胆囊情况】") >= 1 Or InStr(strProfessionalReport, "【胆总管】") >= 1 _
            Or InStr(strProfessionalReport, "【胰腺】") >= 1 Or InStr(strProfessionalReport, "【脾脏】") >= 1
            
            '肝脏情况
            For i = 0 To 4
                cbx肝脏情况(i).Text = GetReportValue(strCurReport, "肝脏情况", cbx肝脏情况(i).Tag)
            Next i
            
            For i = 0 To 5
                chk肝脏情况(i).value = IIf(GetReportValue(strCurReport, "肝脏情况", chk肝脏情况(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            '血管
            For i = 0 To 1
                cbx血管(i).Text = GetReportValue(strCurReport, "血管", cbx血管(i).Tag)
            Next i
            
            '胆囊情况
            For i = 0 To 12
                cbx胆囊情况(i).Text = GetReportValue(strCurReport, "胆囊", cbx胆囊情况(i).Tag)
            Next i
            
            '胆总管
            For i = 0 To 5
                cbx胆总管(i).Text = GetReportValue(strCurReport, "胆总管", cbx胆总管(i).Tag)
            Next i
            
            '胰腺
            For i = 0 To 6
                cbx胰腺(i).Text = GetReportValue(strCurReport, "胰腺", cbx胰腺(i).Tag)
            Next i
            
            '脾脏
            For i = 0 To 4
                cbx脾脏(i).Text = GetReportValue(strCurReport, "脾脏", cbx脾脏(i).Tag)
            Next i
            
            cbxType.ListIndex = 5
            framType(5).Visible = True
        Case InStr(strProfessionalReport, "【肾脏情况】") >= 1 Or InStr(strProfessionalReport, "【输尿管】") >= 1 _
            Or InStr(strProfessionalReport, "【膀胱】") >= 1 Or InStr(strProfessionalReport, "【前列腺】") >= 1
            
            '肾脏情况
            For i = 0 To 5
                txt肾脏情况(i).Text = GetReportValue(strCurReport, "肾脏情况", txt肾脏情况(i).Tag)
            Next i
            
            '输尿管
            For i = 0 To 5
                cbx泌尿男(i).Text = GetReportValue(strCurReport, "输尿管", cbx泌尿男(i).Tag)
            Next i
            
            '膀胱
            cbx膀胱(0).Text = GetReportValue(strCurReport, "膀胱", cbx膀胱(0).Tag)
            
            For i = 0 To 1
                chk泌尿男(i).value = IIf(GetReportValue(strCurReport, "膀胱", chk泌尿男(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            '前列腺
            For i = 0 To 2
                txt前列腺信息(i).Text = GetReportValue(strCurReport, "前列腺", txt前列腺信息(i).Tag)
            Next i
            
            cbxType.ListIndex = 6
            framType(6).Visible = True
        Case InStr(strProfessionalReport, "【肾 脏 情 况】") >= 1 Or InStr(strProfessionalReport, "【输 尿 管】") >= 1 _
            Or InStr(strProfessionalReport, "【膀 胱】") >= 1
            
            '肾脏情况
            For i = 0 To 5
                txt女肾脏情况(i).Text = GetReportValue(strCurReport, "肾 脏 情 况", txt女肾脏情况(i).Tag)
            Next i
            
            '输尿管
            For i = 0 To 5
                cbx输尿管(i).Text = GetReportValue(strCurReport, "输 尿 管", cbx输尿管(i).Tag)
            Next i
            
            '膀胱
            cbx女膀胱(0).Text = GetReportValue(strCurReport, "膀 胱", cbx女膀胱(0).Tag)
            
            For i = 0 To 1
                chk泌尿女(i).value = IIf(GetReportValue(strCurReport, "膀 胱", chk泌尿女(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbxType.ListIndex = 7
            framType(7).Visible = True
            
        Case InStr(strProfessionalReport, "【颈部】") >= 1
            For i = 0 To 13
                txt颈部信息(i).Text = GetReportValue(strCurReport, "颈部", txt颈部信息(i).Tag)
            Next i
            
            cbxType.ListIndex = 8
            framType(8).Visible = True
            
        Case InStr(strProfessionalReport, "【左侧乳腺】") >= 1 Or InStr(strProfessionalReport, "【右侧乳腺】") >= 1
            For i = 0 To 1
                txt乳腺信息(i).Text = GetReportValue(strCurReport, "左侧乳腺", txt乳腺信息(i).Tag)
            Next i
            
            For i = 2 To 3
                txt乳腺信息(i).Text = GetReportValue(strCurReport, "右侧乳腺", txt乳腺信息(i).Tag)
            Next i
            
            cbxType.ListIndex = 9
            framType(9).Visible = True
        Case InStr(strProfessionalReport, "【左侧生殖器】") >= 1 Or InStr(strProfessionalReport, "【右侧生殖器】") >= 1
            For i = 0 To 7
                cbx左侧生殖器(i).Text = GetReportValue(strCurReport, "左侧生殖器", cbx左侧生殖器(i).Tag)
            Next i
            
            For i = 0 To 7
                cbx右侧生殖器(i).Text = GetReportValue(strCurReport, "右侧生殖器", cbx右侧生殖器(i).Tag)
            Next i
            
            cbxType.ListIndex = 10
            framType(10).Visible = True
        Case InStr(strProfessionalReport, "【左下肢静脉】") >= 1 Or InStr(strProfessionalReport, "【右下肢静脉】") >= 1
            For i = 0 To 4
                chk左下肢静脉(i).value = IIf(GetReportValue(strCurReport, "左下肢静脉", chk左下肢静脉(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            For i = 0 To 4
                chk右下肢静脉(i).value = IIf(GetReportValue(strCurReport, "右下肢静脉", chk右下肢静脉(i).Tag) = Report_CheckedValue, 1, 0)
            Next i
            
            cbxType.ListIndex = 11
            framType(11).Visible = True
            
        Case InStr(strProfessionalReport, "【眼部】") >= 1
            For i = 0 To 1
                txt眼部信息(i).Text = GetReportValue(strCurReport, "眼部", txt眼部信息(i).Tag)
            Next i
        
            cbxType.ListIndex = 12
            framType(12).Visible = True
            
        Case InStr(strProfessionalReport, "【甲状腺】") >= 1
            For i = 0 To 9
                txt甲状腺信息(i).Text = GetReportValue(strCurReport, "甲状腺", txt甲状腺信息(i).Tag)
            Next i
        
            cbxType.ListIndex = 13
            framType(13).Visible = True
            
        Case InStr(strProfessionalReport, "【左肾动脉】") >= 1 Or InStr(strProfessionalReport, "【右肾动脉】") >= 1
            For i = 0 To 5
                txt肾动脉信息(i).Text = GetReportValue(strCurReport, "左肾动脉", txt肾动脉信息(i).Tag)
            Next i
            
            For i = 6 To 11
                txt肾动脉信息(i).Text = GetReportValue(strCurReport, "右肾动脉", txt肾动脉信息(i).Tag)
            Next i
        
            cbxType.ListIndex = 14
            framType(14).Visible = True
            
        Case InStr(strProfessionalReport, "【胸腔】") >= 1
            For i = 0 To 5
                cbx胸腔信息(i).Text = GetReportValue(strCurReport, "胸腔", cbx胸腔信息(i).Tag)
            Next i
            
            For i = 0 To 1
                chk胸腔信息(i).value = IIf(GetReportValue(strCurReport, "胸腔", chk胸腔信息(i).Tag) = Report_CheckedValue, 1, 0)
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
    
    
'0-二维及M型超声
'1-多普勒超声
'2-心    脏
'3-妇    科
'4-产    科
'5-腹    部
'6-泌尿(男)
'7-泌尿(女)
'8-颈    部
'9-乳    腺
'10-生 殖 器
'11-下肢静脉
'12-眼    部
'13-甲 状 腺
'14-肾 动 脉
'15-胸    腔
    
    Select Case Val(cbxType.Text)
        Case 0, 1
            strElements = SPLITER_REPORT & Report_Element_右心室前后径 & SPLITER_ELEMENT & txt右心室前后径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_右心室前壁厚度 & SPLITER_ELEMENT & txt右心室前壁厚度.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_右心室横径 & SPLITER_ELEMENT & txt右心室横径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_右心室舒张末容积 & SPLITER_ELEMENT & txt右心室舒张末容积.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_右心室收缩末容积 & SPLITER_ELEMENT & txt右心室收缩末容积.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_右心室射血分数 & SPLITER_ELEMENT & txt右心室射血分数.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_左心室舒张末期径 & SPLITER_ELEMENT & txt左心室舒张末期径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左心室后壁厚度 & SPLITER_ELEMENT & txt左心室后壁厚度.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左心室运动幅度 & SPLITER_ELEMENT & txt左心室运动幅度.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左心室舒张末容积 & SPLITER_ELEMENT & txt左心室舒张末容积.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左心室收缩末容积 & SPLITER_ELEMENT & txt左心室收缩末容积.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左心室射血分数 & SPLITER_ELEMENT & txt左心室射血分数.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左心室每博量 & SPLITER_ELEMENT & txt左心室每博量.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_主动脉根部内径 & SPLITER_ELEMENT & txt主动脉根部内径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_升动脉内径 & SPLITER_ELEMENT & txt升动脉内径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_主肺动脉内径 & SPLITER_ELEMENT & txt主肺动脉内径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_右肺动脉内径 & SPLITER_ELEMENT & txt右肺动脉内径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_左肺动脉内径 & SPLITER_ELEMENT & txt左肺动脉内径.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_室间隔厚度 & SPLITER_ELEMENT & txt室间隔厚度.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_室间隔运动幅度 & SPLITER_ELEMENT & txt室间隔运动幅度.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_室间隔与左室后壁方向 & SPLITER_ELEMENT & txt室间隔与左室后壁方向.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_心房左房前后径 & SPLITER_ELEMENT & txt心房左房前后径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_心房右房长径 & SPLITER_ELEMENT & txt心房右房长径.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_心房右房横径 & SPLITER_ELEMENT & txt心房右房横径.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_二尖瓣舒张期E峰流速 & SPLITER_ELEMENT & txt二尖瓣舒张期E峰流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣舒张期E峰压差 & SPLITER_ELEMENT & txt二尖瓣舒张期E峰压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣舒张期A峰流速 & SPLITER_ELEMENT & txt二尖瓣舒张期A峰流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣舒张期A峰压差 & SPLITER_ELEMENT & txt二尖瓣舒张期A峰压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣收缩期流速 & SPLITER_ELEMENT & txt二尖瓣收缩期流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣收缩期压差 & SPLITER_ELEMENT & txt二尖瓣收缩期压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣等容舒张时间 & SPLITER_ELEMENT & txt二尖瓣等容舒张时间.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣E波减速度 & SPLITER_ELEMENT & txt二尖瓣E波减速度.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_二尖瓣E波减速时间 & SPLITER_ELEMENT & txt二尖瓣E波减速时间.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_主动脉瓣收缩期流速 & SPLITER_ELEMENT & txt主动脉瓣收缩期流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_主动脉瓣收缩期压差 & SPLITER_ELEMENT & txt主动脉瓣收缩期压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_主动脉瓣舒张期流速 & SPLITER_ELEMENT & txt主动脉瓣舒张期流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_主动脉瓣舒张期压差 & SPLITER_ELEMENT & txt主动脉瓣舒张期压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_主动脉瓣流速 & SPLITER_ELEMENT & txt主动脉瓣流速.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_三尖瓣舒张期E峰流速 & SPLITER_ELEMENT & txt三尖瓣舒张期E峰流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_三尖瓣舒张期E峰压差 & SPLITER_ELEMENT & txt三尖瓣舒张期E峰压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_三尖瓣舒张期A峰流速 & SPLITER_ELEMENT & txt三尖瓣舒张期A峰流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_三尖瓣舒张期A峰压差 & SPLITER_ELEMENT & txt三尖瓣舒张期A峰压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_三尖瓣收缩期流速 & SPLITER_ELEMENT & txt三尖瓣收缩期流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_三尖瓣收缩期压差 & SPLITER_ELEMENT & txt三尖瓣收缩期压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_三尖瓣反流压差 & SPLITER_ELEMENT & txt三尖瓣反流压差.Text & SPLITER_REPORT
            
            strElements = strElements & Report_Element_肺动脉瓣收缩期流速 & SPLITER_ELEMENT & txt肺动脉瓣收缩期流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣收缩期压差 & SPLITER_ELEMENT & txt肺动脉瓣收缩期压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣舒张期流速 & SPLITER_ELEMENT & txt肺动脉瓣舒张期流速.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣舒张期压差 & SPLITER_ELEMENT & txt肺动脉瓣舒张期压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣加速时间 & SPLITER_ELEMENT & txt肺动脉瓣加速时间.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣射血时间 & SPLITER_ELEMENT & txt肺动脉瓣射血时间.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣反流压差 & SPLITER_ELEMENT & txt肺动脉瓣反流压差.Text & SPLITER_REPORT
            strElements = strElements & Report_Element_肺动脉瓣峰值流速 & SPLITER_ELEMENT & txt肺动脉瓣峰值流速.Text
        Case 2  '心脏专科
            strElements = GetXinZangReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 3  '妇科专科
            strElements = GetFuKeReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 4  '产科专科
            strElements = GetChanKeReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 5  '腹部专科
            strElements = GetFuBuReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 6  '泌尿(男)
            strElements = GetMiNiaoNanReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 7  '泌尿(女)
            strElements = GetMiNiaoNvReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 8  '颈部
            strElements = GetJingBuReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 9  '乳腺
            strElements = GetRuXianReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 10 '生殖器
            strElements = GetShengZhiQiReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 11 '下肢静脉
            strElements = GetXiaZhiJingMaiReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 12 '眼部
            strElements = GetYanBuReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 13 '甲状腺
            strElements = GetJiaZhuangXianReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 14 '肾动脉
            strElements = GetShenDongMaiReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
        Case 15 '胸腔
            strElements = GetXiongQiangReport()
            
            If strElements <> "" Then strElements = SPLITER_REPORT & Report_Element_专科报告 & SPLITER_ELEMENT & strElements
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
    
    lngTemp = InStr(strReport, "【" & strSection & "】")
    
    '获取数据段如"【左卵巢】  长:12 X 20 cm^2 未显示 【右卵巢】  长13 X 22 cm^2
    If lngTemp <= 0 Then Exit Function
    
    strSource = Mid(strReport, lngTemp + Len("【" & strSection & "】"), 1000)
    lngTemp = InStr(strSource, "      【")
    If lngTemp >= 1 Then strSource = Mid(strSource, 1, lngTemp - 1)
    
    
    '查找指定Tag对应的数据值
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


'获取心脏报告（需要用自定义报表，否则不支持回车符号）
Private Function GetXinZangReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(cbx心脏, 0, 33)
    If strReport <> "" Then strReport = "【心脏】  " & strReport & "  "
    
    GetXinZangReport = strReport
End Function

'获取妇科报告
Private Function GetFuKeReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""
    
    '子宫情况
    strTemp = GetSectionReportWithCombobox(cbx子宫, 0, 5)
    If strTemp <> "" Then strReport = "【子宫情况】  " & strTemp & "  "
    
    
    '左卵巢
    strTemp = GetSectionReportWithCombobox(cbx左卵巢, 0, 1)
    
    If chk妇科(0).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk妇科(0).Tag, "[value]", Report_CheckedValue)
    End If
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【左卵巢】  " & strTemp & "  "
        
    
    
    '右卵巢
    strTemp = GetSectionReportWithCombobox(cbx右卵巢, 0, 1)
    
    If chk妇科(1).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk妇科(1).Tag, "[value]", Report_CheckedValue)
    End If
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【右卵巢】  " & strTemp & "  "
    
    
    '胎囊
    strTemp = GetSectionReportWithCombobox(cbx胎囊, 0, 2)
    
    If strTemp <> "" Then strTemp = strTemp & "  "
    strTemp = strTemp & GetSectionReportWithCombobox(chk妇科, 2, 6, True)
    
    If cbx妇科原始心管搏动.Text <> "" Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(cbx妇科原始心管搏动.Tag, "[value]", cbx妇科原始心管搏动.Text)
    End If
    
    
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【胎囊】  " & strTemp & "  "
    
    GetFuKeReport = strReport
End Function

'获取产科报告
Private Function GetChanKeReport() As String
    Dim strReport As String
    
    strReport = ""
    
    strReport = GetSectionReportWithCombobox(cbx产科情况, 0, 11)

    If strReport <> "" Then strReport = strReport & "  "
    strReport = strReport & GetSectionReportWithCombobox(chk产科情况, 0, 3)
    
    If strReport <> "" Then strReport = strReport & "  "
    strReport = strReport & GetSectionReportWithCombobox(cbx产科情况, 12, 21)
    
    If strReport <> "" Then strReport = "【产科情况】  " & strReport & "  "
    
    GetChanKeReport = strReport
End Function

'获取腹部报告
Private Function GetFuBuReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '肝脏情况
    strTemp = GetSectionReportWithCombobox(cbx肝脏情况, 0, 4)
    
    If strTemp <> "" Then strTemp = strTemp & "  "
    strTemp = strTemp & GetSectionReportWithCombobox(chk肝脏情况, 0, 5, True)
    
    If strTemp <> "" Then strReport = "【肝脏情况】  " & strTemp & "  "
    
    
    '血管
    strTemp = GetSectionReportWithCombobox(cbx血管, 0, 1)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【血管】  " & strTemp & "  "
        

    '胆囊
    strTemp = GetSectionReportWithCombobox(cbx胆囊情况, 0, 12)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【胆囊】  " & strTemp & "  "
    
    
    '胆总管
    strTemp = GetSectionReportWithCombobox(cbx胆总管, 0, 5)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【胆总管】  " & strTemp & "  "
    
    
    '胰腺
    strTemp = GetSectionReportWithCombobox(cbx胰腺, 0, 6)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【胰腺】  " & strTemp & "  "
    
    
    '脾脏
    strTemp = GetSectionReportWithCombobox(cbx脾脏, 0, 4)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【脾脏】  " & strTemp & "  "
    
    
    GetFuBuReport = strReport
End Function



Private Function GetSectionReportWithCombobox(aryControl As Variant, ByVal lngStartIndex As Long, _
    ByVal lngEndIndex As Long, Optional ByVal blnIsCheckBox As Boolean = False, _
    Optional ByVal blnAutoReplace As Boolean = True) As String
'根据传递的组件获取报告信息
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
                '[T1]替换为“左心收缩功能:”
                If InStr(strCurElement, "[T1]") >= 1 And Not blnAddTag1 Then
                    strCurElement = Replace(strCurElement, "[T1]", "左心收缩功能:")
                    blnAddTag1 = True
                Else
                    strCurElement = Replace(strCurElement, "[T1]", "")
                End If
                
                '[T2]替换为“囊壁:”
                If InStr(strCurElement, "[T2]") >= 1 And Not blnAddTag2 Then
                    strCurElement = Replace(strCurElement, "[T2]", "囊壁:")
                    blnAddTag2 = True
                Else
                    strCurElement = Replace(strCurElement, "[T2]", "")
                End If
                
                '[T3]替换为“羊水:”
                If InStr(strCurElement, "[T3]") >= 1 And Not blnAddTag3 Then
                    strCurElement = Replace(strCurElement, "[T3]", "羊水:")
                    blnAddTag3 = True
                Else
                    strCurElement = Replace(strCurElement, "[T3]", "")
                End If
                
                '[T4]替换为“脐动脉血流指数:”
                If InStr(strCurElement, "[T4]") >= 1 And Not blnAddTag4 Then
                    strCurElement = Replace(strCurElement, "[T4]", "脐动脉血流指数:")
                    blnAddTag4 = True
                Else
                    strCurElement = Replace(strCurElement, "[T4]", "")
                End If
                
                '[T5]替换为“反射:”
                If InStr(strCurElement, "[T5]") >= 1 And Not blnAddTag5 Then
                    strCurElement = Replace(strCurElement, "[T5]", "反射:")
                    blnAddTag5 = True
                Else
                    strCurElement = Replace(strCurElement, "[T5]", "")
                End If
                
                '[T6]替换为“近端:”
                If InStr(strCurElement, "[T6]") >= 1 And Not blnAddTag6 Then
                    strCurElement = Replace(strCurElement, "[T6]", "近端:")
                    blnAddTag6 = True
                Else
                    strCurElement = Replace(strCurElement, "[T6]", "")
                End If
                
                '[T7]替换为“远端:”
                If InStr(strCurElement, "[T7]") >= 1 And Not blnAddTag7 Then
                    strCurElement = Replace(strCurElement, "[T7]", "远端:")
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



'获取泌尿男报告
Private Function GetMiNiaoNanReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '肾脏情况
    strTemp = strTemp & GetSectionReportWithCombobox(txt肾脏情况, 0, 5)
    If strTemp <> "" Then strReport = "【肾脏情况】  " & strTemp & "  "
    
    
    '泌尿
    strTemp = GetSectionReportWithCombobox(cbx泌尿男, 0, 5)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【输尿管】  " & strTemp & "  "
        
    '膀胱
    strTemp = IIf(cbx膀胱(0).Text = "", "", Replace(cbx膀胱(0).Tag, "[value]", cbx膀胱(0).Text))
    If strTemp <> "" Then strTemp = strTemp & "  "
    
    strTemp = strTemp & GetSectionReportWithCombobox(chk泌尿男, 0, 1, True)
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【膀胱】  " & strTemp & "  "
    
    
    '前列腺
    strTemp = GetSectionReportWithCombobox(txt前列腺信息, 0, 2)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【前列腺】  " & strTemp & "  "
    
    
    GetMiNiaoNanReport = strReport
End Function

'获取泌尿女报告
Private Function GetMiNiaoNvReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '肾脏情况
    strTemp = strTemp & GetSectionReportWithCombobox(txt女肾脏情况, 0, 5)
    If strTemp <> "" Then strReport = "【肾 脏 情 况】  " & strTemp & "  "
    
    
    '泌尿
    strTemp = GetSectionReportWithCombobox(cbx输尿管, 0, 5)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【输 尿 管】  " & strTemp & "  "
        
    '膀胱
    strTemp = IIf(cbx女膀胱(0).Text = "", "", Replace(cbx女膀胱(0).Tag, "[value]", cbx女膀胱(0).Text))
    If strTemp <> "" Then strTemp = strTemp & "  "
    
    strTemp = strTemp & GetSectionReportWithCombobox(chk泌尿女, 0, 1, True)
    
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【膀 胱】  " & strTemp & "  "
    
    
    GetMiNiaoNvReport = strReport
End Function

'获取颈部报告
Private Function GetJingBuReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(txt颈部信息, 0, 13)
    If strReport <> "" Then strReport = "【颈部】  " & strReport & "  "
    
    GetJingBuReport = strReport
End Function

'获取乳腺报告
Private Function GetRuXianReport() As String
    Dim strReport As String
    Dim strTemp As String

    strTemp = ""
    strReport = ""
    
    strTemp = GetSectionReportWithCombobox(txt乳腺信息, 0, 1)
    If strTemp <> "" Then strReport = "【左侧乳腺】  " & strTemp & "  "
    
    strTemp = GetSectionReportWithCombobox(txt乳腺信息, 2, 3)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【右侧乳腺】  " & strTemp & "  "
    
    GetRuXianReport = strReport
End Function

'获取生殖器报告
Private Function GetShengZhiQiReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '左侧生殖器
    strTemp = strTemp & GetSectionReportWithCombobox(cbx左侧生殖器, 0, 7)
    If strTemp <> "" Then strReport = "【左侧生殖器】  " & strTemp & "  "
    
    
    '右侧生殖器
    strTemp = GetSectionReportWithCombobox(cbx右侧生殖器, 0, 7)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【右侧生殖器】  " & strTemp & "  "
    
    GetShengZhiQiReport = strReport
End Function

'获取下肢静脉报告
Private Function GetXiaZhiJingMaiReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""

    
    '左侧生殖器
    strTemp = strTemp & GetSectionReportWithCombobox(chk左下肢静脉, 0, 4, True)
    If strTemp <> "" Then strReport = "【左下肢静脉】  " & strTemp & "  血流通畅,管腔内未见异常回声,探过加压后管腔消失  "
    
    
    '右侧生殖器
    strTemp = GetSectionReportWithCombobox(chk右下肢静脉, 0, 4, True)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【右下肢静脉】  " & strTemp & "  血流通畅,管腔内未见异常回声,探过加压后管腔消失  "
    
    GetXiaZhiJingMaiReport = strReport
End Function

'获取眼部报告
Private Function GetYanBuReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(txt眼部信息, 0, 1)
    If strReport <> "" Then strReport = "【眼部】  " & strReport & "  "
    
    GetYanBuReport = strReport
End Function

'获取甲状腺报告
Private Function GetJiaZhuangXianReport() As String
    Dim strReport As String

    strReport = GetSectionReportWithCombobox(txt甲状腺信息, 0, 9)
    If strReport <> "" Then strReport = "【甲状腺】  " & strReport & "  "
    
    GetJiaZhuangXianReport = strReport
End Function

'获取肾动脉报告
Private Function GetShenDongMaiReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""
    
    strTemp = GetSectionReportWithCombobox(txt肾动脉信息, 0, 5)
    If strTemp <> "" Then strReport = "【左肾动脉】  " & strTemp & "  "
    
    strTemp = GetSectionReportWithCombobox(txt肾动脉信息, 6, 11)
    If strTemp <> "" Then strReport = strReport & IIf(strReport <> "", Report_ProjectSplitChr, "") & "【右肾动脉】  " & strTemp & "  "
    
    GetShenDongMaiReport = strReport
End Function

'获取胸腔报告
Private Function GetXiongQiangReport() As String
    Dim strReport As String
    Dim strTemp As String
    
    strReport = ""
    strTemp = ""
    
    strTemp = GetSectionReportWithCombobox(cbx胸腔信息, 0, 2)
    
    
    If chk胸腔信息(0).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk胸腔信息(0).Tag, "[value]", Report_CheckedValue)
    End If
    
    
    strReport = GetSectionReportWithCombobox(cbx胸腔信息, 3, 5)
    If strReport <> "" Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & strReport
    End If
    
    strReport = ""
    If chk胸腔信息(1).value <> 0 Then
        If strTemp <> "" Then strTemp = strTemp & "  "
        strTemp = strTemp & Replace(chk胸腔信息(1).Tag, "[value]", Report_CheckedValue)
    End If
    
    If strTemp <> "" Then strReport = "【胸腔】  " & strTemp & "  "
    
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

Private Sub cbx产科情况_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx产科情况_Click(Index As Integer)
    If cbx产科情况(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胆囊情况_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胆囊情况_Click(Index As Integer)
    If cbx胆囊情况(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胆总管_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胆总管_Click(Index As Integer)
    If cbx胆总管(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx妇科原始心管搏动_Click()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx肝脏情况_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx肝脏情况_Click(Index As Integer)
    If cbx肝脏情况(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx泌尿男_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx泌尿男_Click(Index As Integer)
    If cbx泌尿男(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx泌尿女_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx膀胱_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx脾脏_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx脾脏_Click(Index As Integer)
    If cbx脾脏(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx输尿管_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx输尿管_Click(Index As Integer)
    If cbx输尿管(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胎囊_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx心脏_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx心脏_Click(Index As Integer)
    If cbx心脏(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胸腔信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胸腔信息_Click(Index As Integer)
    If cbx胸腔信息(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx血管_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胰腺_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx胰腺_Click(Index As Integer)
    If cbx胰腺(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx右侧生殖器_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx右侧生殖器_Click(Index As Integer)
    If cbx右侧生殖器(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx右卵巢_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx子宫_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx子宫_Click(Index As Integer)
    If cbx子宫(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx左侧生殖器_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx左侧生殖器_Click(Index As Integer)
    If cbx左侧生殖器(Index).Style <> 0 Then Exit Sub
    
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub cbx左卵巢_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk产科情况_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk妇科_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk肝脏情况_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk泌尿男_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk泌尿女_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk右下肢静脉_Click(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub chk左下肢静脉_Click(Index As Integer)
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
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CY22", Me.Height
End Sub

Private Sub txt二尖瓣E波减速度_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣E波减速时间_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣等容舒张时间_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣收缩期流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣收缩期压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣舒张期A峰流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣舒张期A峰压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣舒张期E峰流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt二尖瓣舒张期E峰压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣反流压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣峰值流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣加速时间_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣射血时间_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣收缩期流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣收缩期压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣舒张期流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肺动脉瓣舒张期压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt甲状腺信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt颈部信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt女肾脏情况_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt前列腺信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt乳腺信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣反流压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣收缩期流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣收缩期压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣舒张期A峰流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣舒张期A峰压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣舒张期E峰流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt三尖瓣舒张期E峰压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肾动脉信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt肾脏情况_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt升动脉内径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt室间隔厚度_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt室间隔与左室后壁方向_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt室间隔运动幅度_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt心房右房长径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt心房右房横径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt心房左房前后径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt眼部信息_Change(Index As Integer)
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右肺动脉内径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右心室横径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右心室前臂厚度_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右心室前后径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右心室射血分数_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右心室收缩末容积_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt右心室舒张末容积_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主动脉瓣流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主动脉瓣收缩期流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主动脉瓣收缩期压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主动脉瓣舒张期流速_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主动脉瓣舒张期压差_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主动脉根部内径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt主肺动脉内径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左肺动脉内径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室后壁厚度_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室每搏量_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室射血分数_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室收缩末容积_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室舒张末期径_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室舒张末容积_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txt左心室运动幅度_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub
