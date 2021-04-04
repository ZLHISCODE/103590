VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMain_北京尚洋病案信息编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "长治病案信息编辑"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14685
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   Icon            =   "frmMain_北京尚洋病案信息编辑.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   13380
      TabIndex        =   217
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读取(&R)"
      Height          =   350
      Left            =   11040
      TabIndex        =   216
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   12225
      TabIndex        =   215
      Top             =   5430
      Width           =   1100
   End
   Begin VB.PictureBox pic诊断信息 
      BorderStyle     =   0  'None
      Height          =   4545
      Left            =   255
      ScaleHeight     =   4545
      ScaleWidth      =   14220
      TabIndex        =   109
      Top             =   345
      Width           =   14220
      Begin VB.ComboBox cmbTREAT_RESULT2 
         DataField       =   "TREAT_RESULT2"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   178
         Top             =   3600
         Width           =   3000
      End
      Begin VB.ComboBox cmbBLOOD_TRAN_REACT_FLAG 
         DataField       =   "BLOOD_TRAN_REACT_FLAG"
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   177
         Top             =   960
         Width           =   3000
      End
      Begin VB.ComboBox cmbTEACH_MR_FLAG 
         DataField       =   "TEACH_MR_FLAG"
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   176
         Top             =   510
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_DATE3 
         DataField       =   "OUT_DIAGNOSIS_DATE3"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   175
         Top             =   4110
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_CODE2 
         DataField       =   "OUT_DIAGNOSIS_CODE2"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   174
         Top             =   3165
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_NAME1 
         DataField       =   "OUT_DIAGNOSIS_NAME1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   173
         Top             =   2730
         Width           =   3000
      End
      Begin VB.TextBox txtIN_DIAGNOSIS_NAME 
         DataField       =   "IN_DIAGNOSIS_NAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1425
         TabIndex        =   172
         Top             =   2295
         Width           =   3000
      End
      Begin VB.TextBox txtHANDLE 
         DataField       =   "HANDLE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   171
         Top             =   1860
         Width           =   3000
      End
      Begin VB.TextBox txtPLASM 
         Alignment       =   1  'Right Justify
         DataField       =   "PLASM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   170
         Text            =   "0"
         Top             =   1425
         Width           =   3000
      End
      Begin VB.TextBox txtSickID3 
         BackColor       =   &H80000000&
         DataField       =   "STICKID"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   169
         Top             =   150
         Width           =   3000
      End
      Begin VB.ComboBox cmbTREAT_RESULT3 
         DataField       =   "TREAT_RESULT3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   168
         Top             =   4095
         Width           =   3000
      End
      Begin VB.ComboBox cmbTREAT_RESULT1 
         DataField       =   "TREAT_RESULT1"
         Height          =   315
         Left            =   11025
         Style           =   2  'Dropdown List
         TabIndex        =   167
         Top             =   2715
         Width           =   3000
      End
      Begin VB.ComboBox cmbRH 
         DataField       =   "RH"
         Height          =   315
         Left            =   11025
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   510
         Width           =   3000
      End
      Begin VB.ComboBox cmbBLOOD_TYPE 
         DataField       =   "BLOOD_TYPE"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   510
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_CODE3 
         DataField       =   "OUT_DIAGNOSIS_CODE3"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   164
         Top             =   3615
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_NAME2 
         DataField       =   "OUT_DIAGNOSIS_NAME2"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6150
         TabIndex        =   163
         Top             =   3165
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_DATE1 
         DataField       =   "OUT_DIAGNOSIS_DATE1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   162
         Top             =   2730
         Width           =   3000
      End
      Begin VB.TextBox txtIN_DIAGNOSIS_DATE 
         DataField       =   "IN_DIAGNOSIS_DATE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   161
         Top             =   2295
         Width           =   3000
      End
      Begin VB.TextBox txtHANDLE_DATE 
         DataField       =   "HANDLE_DATE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6135
         TabIndex        =   160
         Top             =   1860
         Width           =   3000
      End
      Begin VB.TextBox txtBLOOD 
         Alignment       =   1  'Right Justify
         DataField       =   "BLOOD"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   159
         Text            =   "0"
         Top             =   1425
         Width           =   3000
      End
      Begin VB.TextBox txtERYTHROCYTE 
         Alignment       =   1  'Right Justify
         DataField       =   "ERYTHROCYTE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   158
         Text            =   "0"
         Top             =   975
         Width           =   3000
      End
      Begin VB.TextBox txtCnName3 
         BackColor       =   &H80000000&
         DataField       =   "CNNAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   150
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_NAME3 
         DataField       =   "OUT_DIAGNOSIS_NAME3"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   11025
         TabIndex        =   130
         Top             =   3615
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_DATE2 
         DataField       =   "OUT_DIAGNOSIS_DATE2"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   11025
         TabIndex        =   129
         Top             =   3165
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DIAGNOSIS_CODE1 
         DataField       =   "OUT_DIAGNOSIS_CODE1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         TabIndex        =   128
         Top             =   2295
         Width           =   3000
      End
      Begin VB.TextBox txtIN_DIAGNOSIS_CODE 
         DataField       =   "IN_DIAGNOSIS_CODE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11025
         TabIndex        =   127
         Top             =   1860
         Width           =   3000
      End
      Begin VB.TextBox txtOTHER_BLOOD 
         Alignment       =   1  'Right Justify
         DataField       =   "OTHER_BLOOD"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11025
         TabIndex        =   126
         Text            =   "0"
         Top             =   1425
         Width           =   3000
      End
      Begin VB.TextBox txtHEMOBLAST 
         Alignment       =   1  'Right Justify
         DataField       =   "HEMOBLAST"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11025
         TabIndex        =   125
         Text            =   "0"
         Top             =   975
         Width           =   3000
      End
      Begin VB.TextBox txtSex3 
         BackColor       =   &H80000000&
         DataField       =   "SEX"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11025
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   150
         Width           =   3000
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   24
         Left            =   9180
         TabIndex        =   254
         Top             =   4140
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   23
         Left            =   4470
         TabIndex        =   253
         Top             =   4140
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   22
         Left            =   14055
         TabIndex        =   252
         Top             =   3660
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   21
         Left            =   4455
         TabIndex        =   251
         Top             =   3660
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   20
         Left            =   14040
         TabIndex        =   250
         Top             =   3225
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   19
         Left            =   9180
         TabIndex        =   249
         Top             =   3195
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   18
         Left            =   14055
         TabIndex        =   236
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   17
         Left            =   9180
         TabIndex        =   235
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   16
         Left            =   4440
         TabIndex        =   234
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   15
         Left            =   14040
         TabIndex        =   233
         Top             =   2355
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   14
         Left            =   9180
         TabIndex        =   232
         Top             =   2325
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   13
         Left            =   4455
         TabIndex        =   231
         Top             =   2325
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   12
         Left            =   14040
         TabIndex        =   230
         Top             =   1905
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   11
         Left            =   9165
         TabIndex        =   229
         Top             =   1890
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   4455
         TabIndex        =   228
         Top             =   1905
         Width           =   120
      End
      Begin VB.Label labTREAT_RESULT3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "治疗结果3"
         Height          =   195
         Left            =   5160
         TabIndex        =   156
         Top             =   4155
         Width           =   885
      End
      Begin VB.Label labOUT_DIAGNOSIS_DATE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断日期3"
         Height          =   195
         Left            =   90
         TabIndex        =   155
         Top             =   4155
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_NAME3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断名称3"
         Height          =   195
         Left            =   9630
         TabIndex        =   154
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_CODE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断代码3"
         Height          =   195
         Left            =   4770
         TabIndex        =   153
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label labTREAT_RESULT2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "治疗结果2"
         Height          =   195
         Left            =   480
         TabIndex        =   152
         Top             =   3660
         Width           =   885
      End
      Begin VB.Label labOUT_DIAGNOSIS_DATE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断日期2"
         Height          =   195
         Left            =   9630
         TabIndex        =   151
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_NAME2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断名称2"
         Height          =   195
         Left            =   4770
         TabIndex        =   150
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_CODE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断代码2"
         Height          =   195
         Left            =   90
         TabIndex        =   149
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label labTREAT_RESULT1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "治疗结果1"
         Height          =   195
         Left            =   10020
         TabIndex        =   148
         Top             =   2775
         Width           =   885
      End
      Begin VB.Label labOUT_DIAGNOSIS_DATE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断日期1"
         Height          =   195
         Left            =   4770
         TabIndex        =   147
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_NAME1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断名称1"
         Height          =   195
         Left            =   90
         TabIndex        =   146
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_CODE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院诊断代码1"
         Height          =   195
         Left            =   9615
         TabIndex        =   145
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label labIN_DIAGNOSIS_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "诊断日期"
         Height          =   195
         Left            =   5250
         TabIndex        =   144
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label labIN_DIAGNOSIS_NAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院诊断名称"
         Height          =   195
         Left            =   180
         TabIndex        =   143
         Top             =   2340
         Width           =   1170
      End
      Begin VB.Label labIN_DIAGNOSIS_CODE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院诊断代码"
         Height          =   195
         Left            =   9735
         TabIndex        =   142
         Top             =   1905
         Width           =   1170
      End
      Begin VB.Label labHANDLE_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "操作日期"
         Height          =   195
         Left            =   5265
         TabIndex        =   141
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label labHANDLE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "操作人"
         Height          =   195
         Left            =   780
         TabIndex        =   140
         Top             =   1905
         Width           =   585
      End
      Begin VB.Label labOTHER_BLOOD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "其他"
         Height          =   195
         Left            =   10515
         TabIndex        =   139
         Top             =   1470
         Width           =   390
      End
      Begin VB.Label labBLOOD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "全血"
         Height          =   195
         Left            =   5655
         TabIndex        =   138
         Top             =   1470
         Width           =   390
      End
      Begin VB.Label labPLASM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "输入血浆"
         Height          =   195
         Left            =   585
         TabIndex        =   137
         Top             =   1470
         Width           =   780
      End
      Begin VB.Label labHEMOBLAST 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "输入血小板"
         Height          =   195
         Left            =   9930
         TabIndex        =   136
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label labERYTHROCYTE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "输入红细胞"
         Height          =   195
         Left            =   5070
         TabIndex        =   135
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label labBLOOD_TRAN_REACT_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "输入血反应标志"
         Height          =   195
         Left            =   0
         TabIndex        =   134
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label labRH 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "RH"
         Height          =   195
         Left            =   10665
         TabIndex        =   133
         Top             =   570
         Width           =   240
      End
      Begin VB.Label labBLOOD_TYPE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "血型"
         Height          =   195
         Left            =   5655
         TabIndex        =   132
         Top             =   525
         Width           =   390
      End
      Begin VB.Label labTEACH_MR_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "示教病例"
         Height          =   195
         Left            =   585
         TabIndex        =   131
         Top             =   570
         Width           =   780
      End
      Begin VB.Label labSick1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人ID"
         Height          =   195
         Left            =   765
         TabIndex        =   124
         Top             =   195
         Width           =   600
      End
      Begin VB.Label labCnName3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   195
         Left            =   5655
         TabIndex        =   123
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSex3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   195
         Left            =   10515
         TabIndex        =   122
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.PictureBox pic基本信息 
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   232
      ScaleHeight     =   4425
      ScaleWidth      =   14220
      TabIndex        =   0
      Top             =   315
      Width           =   14220
      Begin VB.CommandButton cmdSick 
         Caption         =   "…"
         Height          =   285
         Left            =   4170
         TabIndex        =   247
         Top             =   150
         Width           =   255
      End
      Begin VB.TextBox txtDISCHARGE_DATE 
         DataField       =   "DISCHARGE_DATE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   218
         Top             =   3690
         Width           =   3000
      End
      Begin VB.ComboBox cmbRELATIONSHIP 
         DataField       =   "RELATIONSHIP"
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2760
         Width           =   3000
      End
      Begin VB.ComboBox cmbMARITAL_STATUS 
         DataField       =   "MARITAL_STATUS"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   960
         Width           =   3000
      End
      Begin VB.TextBox txtHOSPITAL_NUMBER 
         BackColor       =   &H80000000&
         DataField       =   "HOSPITAL_NUMBER"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   540
         Width           =   3000
      End
      Begin VB.TextBox txtRESIDENCE_NO 
         BackColor       =   &H80000000&
         DataField       =   "RESIDENCE_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   28
         Top             =   540
         Width           =   3000
      End
      Begin VB.TextBox txtIN_COUNT 
         BackColor       =   &H80000000&
         DataField       =   "IN_COUNT"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   27
         Top             =   540
         Width           =   3000
      End
      Begin VB.TextBox txtMEDICAL_RECORD_NO 
         DataField       =   "MEDICAL_RECORD_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   26
         Top             =   975
         Width           =   3000
      End
      Begin VB.TextBox txtSTATUS 
         DataField       =   "STATUS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   25
         Top             =   975
         Width           =   3000
      End
      Begin VB.TextBox txtBIRTH_ADDRESS 
         DataField       =   "BIRTH_ADDRESS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   24
         Top             =   1440
         Width           =   3000
      End
      Begin VB.TextBox txtIDENTITY_NUMBER 
         DataField       =   "IDENTITY_NUMBER"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   18
         TabIndex        =   23
         Top             =   1440
         Width           =   3000
      End
      Begin VB.TextBox txtUNIT_NAME 
         DataField       =   "UNIT_NAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1440
         Width           =   3000
      End
      Begin VB.TextBox txtUNIT_ADDRESS 
         DataField       =   "UNIT_ADDRESS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1875
         Width           =   3000
      End
      Begin VB.TextBox txtUNIT_PHONE 
         DataField       =   "UNIT_PHONE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1875
         Width           =   3000
      End
      Begin VB.TextBox txtUNIT_ZIPCODE 
         DataField       =   "UNIT_ZIPCODE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1875
         Width           =   3000
      End
      Begin VB.TextBox txtREGISTER_ADDRESS 
         DataField       =   "REGISTER_ADDRESS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtREGISTER_ZIPCODE 
         DataField       =   "REGISTER_ZIPCODE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtCONTACT_PERSON 
         DataField       =   "CONTACT_PERSON"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtCONTACT_ADDRESS 
         DataField       =   "CONTACT_ADDRESS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   60
         TabIndex        =   15
         Top             =   2775
         Width           =   3000
      End
      Begin VB.TextBox txtCONTACT_PHONE 
         DataField       =   "CONTACT_PHONE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2775
         Width           =   3000
      End
      Begin VB.TextBox txtADMISSION_DATE 
         DataField       =   "ADMISSION_DATE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   3240
         Width           =   3000
      End
      Begin VB.TextBox txtADMISSION_DEPT 
         DataField       =   "ADMISSION_DEPT"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   20
         TabIndex        =   12
         Top             =   3240
         Width           =   3000
      End
      Begin VB.TextBox txtIN_DEPT_ZONE 
         DataField       =   "IN_DEPT_ZONE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3240
         Width           =   3000
      End
      Begin VB.TextBox txtDEPT_TRANSFERED_TO 
         DataField       =   "DEPT_TRANSFERED_TO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3690
         Width           =   3000
      End
      Begin VB.TextBox txtDISCHARGE_DEPT 
         DataField       =   "DISCHARGE_DEPT"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   9
         Top             =   3690
         Width           =   3000
      End
      Begin VB.TextBox txtOUT_DEPT_ZONE 
         DataField       =   "OUT_DEPT_ZONE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   8
         Top             =   4095
         Width           =   3000
      End
      Begin VB.TextBox txtDIAGNOSIS_DATE 
         DataField       =   "DIAGNOSIS_DATE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         TabIndex        =   7
         Top             =   4095
         Width           =   3000
      End
      Begin VB.TextBox txtSickID1 
         DataField       =   "STICKID"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   150
         Width           =   2730
      End
      Begin VB.TextBox txtCnName1 
         BackColor       =   &H80000000&
         DataField       =   "CNNAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   150
         Width           =   3000
      End
      Begin VB.TextBox txtSex1 
         BackColor       =   &H80000000&
         DataField       =   "SEX"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   150
         Width           =   3000
      End
      Begin VB.ComboBox cmbPAT_ADM_CONDITION 
         DataField       =   "PAT_ADM_CONDITION"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   4080
         Width           =   3000
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   248
         Top             =   0
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   9165
         TabIndex        =   226
         Top             =   3765
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   7
         Left            =   14010
         TabIndex        =   225
         Top             =   3285
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   9150
         TabIndex        =   224
         Top             =   3285
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   4455
         TabIndex        =   223
         Top             =   3300
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   14010
         TabIndex        =   222
         Top             =   570
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   4455
         TabIndex        =   221
         Top             =   1020
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   9150
         TabIndex        =   220
         Top             =   585
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   4455
         TabIndex        =   219
         Top             =   570
         Width           =   120
      End
      Begin VB.Label labHOSPITAL_NUMBER 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医疗机构编码"
         Height          =   195
         Left            =   210
         TabIndex        =   56
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label labRESIDENCE_NO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   195
         Left            =   5460
         TabIndex        =   55
         Top             =   585
         Width           =   585
      End
      Begin VB.Label labIN_COUNT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "本次住院次数"
         Height          =   195
         Left            =   9735
         TabIndex        =   54
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label labMEDICAL_RECORD_NO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病案号"
         Height          =   195
         Left            =   795
         TabIndex        =   53
         Top             =   1020
         Width           =   585
      End
      Begin VB.Label labMARITAL_STATUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "婚姻状况"
         Height          =   195
         Left            =   5265
         TabIndex        =   52
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label labSTATUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "职业"
         Height          =   195
         Left            =   10515
         TabIndex        =   51
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label labBIRTH_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出生地"
         Height          =   195
         Left            =   795
         TabIndex        =   50
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label labIDENTITY_NUMBER 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "身份证号"
         Height          =   195
         Left            =   5265
         TabIndex        =   49
         Top             =   1485
         Width           =   780
      End
      Begin VB.Label labUNIT_NAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "工作单位"
         Height          =   195
         Left            =   10125
         TabIndex        =   48
         Top             =   1485
         Width           =   780
      End
      Begin VB.Label labUNIT_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "单位地址"
         Height          =   195
         Left            =   600
         TabIndex        =   47
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labUNIT_PHONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "单位电话"
         Height          =   195
         Left            =   5265
         TabIndex        =   46
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labUNIT_ZIPCODE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "单位邮编"
         Height          =   195
         Left            =   10125
         TabIndex        =   45
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labREGISTER_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "户口地址"
         Height          =   195
         Left            =   600
         TabIndex        =   44
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labREGISTER_ZIPCODE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "户口邮编"
         Height          =   195
         Left            =   5265
         TabIndex        =   43
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labCONTACT_PERSON 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "联系人"
         Height          =   195
         Left            =   10320
         TabIndex        =   42
         Top             =   2370
         Width           =   585
      End
      Begin VB.Label labRELATIONSHIP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "与病人关系"
         Height          =   195
         Left            =   405
         TabIndex        =   41
         Top             =   2805
         Width           =   975
      End
      Begin VB.Label labCONTACT_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "联系地址"
         Height          =   195
         Left            =   5265
         TabIndex        =   40
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labCONTACT_PHONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "联系电话"
         Height          =   195
         Left            =   10125
         TabIndex        =   39
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labADMISSION_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院日期"
         Height          =   195
         Left            =   600
         TabIndex        =   38
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labADMISSION_DEPT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院科室"
         Height          =   195
         Left            =   5265
         TabIndex        =   37
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labIN_DEPT_ZONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院病室"
         Height          =   195
         Left            =   10125
         TabIndex        =   36
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labDEPT_TRANSFERED_TO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "转科科别"
         Height          =   195
         Left            =   600
         TabIndex        =   35
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labDISCHARGE_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院日期"
         Height          =   195
         Left            =   5265
         TabIndex        =   34
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labDISCHARGE_DEPT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院科室"
         Height          =   195
         Left            =   10125
         TabIndex        =   33
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labOUT_DEPT_ZONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院病室"
         Height          =   195
         Left            =   600
         TabIndex        =   32
         Top             =   4140
         Width           =   780
      End
      Begin VB.Label labPAT_ADM_CONDITION 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院病情"
         Height          =   195
         Left            =   5265
         TabIndex        =   31
         Top             =   4140
         Width           =   780
      End
      Begin VB.Label labDIAGNOSIS_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院后确诊日期"
         Height          =   195
         Left            =   9540
         TabIndex        =   30
         Top             =   4140
         Width           =   1365
      End
      Begin VB.Label labSickID1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人ID"
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   195
         Width           =   600
      End
      Begin VB.Label labCnName1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   195
         Left            =   5655
         TabIndex        =   5
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSex1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   195
         Left            =   10515
         TabIndex        =   4
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.PictureBox pic医师信息 
      BorderStyle     =   0  'None
      Height          =   4545
      Left            =   232
      ScaleHeight     =   4545
      ScaleWidth      =   14220
      TabIndex        =   60
      Top             =   315
      Width           =   14220
      Begin VB.ComboBox cmbMEDICAL_RECORD_MASS 
         DataField       =   "MEDICAL_RECORD_MASS"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   237
         Top             =   3225
         Width           =   3000
      End
      Begin VB.ComboBox cmbFOLLOW_FLAG 
         DataField       =   "FOLLOW_FLAG"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   4095
         Width           =   3000
      End
      Begin VB.ComboBox cmbFIRST_FLAG 
         DataField       =   "FIRST_FLAG"
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   4095
         Width           =   3000
      End
      Begin VB.ComboBox cmbEMIT_PATHOLOGY 
         DataField       =   "EMIT_PATHOLOGY"
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   1440
         Width           =   3000
      End
      Begin VB.ComboBox cmbIN_OUT 
         DataField       =   "IN_OUT"
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   975
         Width           =   3000
      End
      Begin VB.ComboBox cmbCLINIC_PATHOLOGY 
         DataField       =   "CLINIC_PATHOLOGY"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   1440
         Width           =   3000
      End
      Begin VB.ComboBox cmbBEFORE_AFTER_TREATMENT 
         DataField       =   "BEFORE_AFTER_TREATMENT"
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   1440
         Width           =   3000
      End
      Begin VB.ComboBox cmbHBSAG 
         DataField       =   "HBSAG"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   525
         Width           =   3000
      End
      Begin VB.ComboBox cmdCLINIC_INHOSPITAL 
         DataField       =   "CLINIC_INHOSPITAL"
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   975
         Width           =   3000
      End
      Begin VB.ComboBox cmbHIV_AB 
         DataField       =   "HIV_AB"
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   975
         Width           =   3000
      End
      Begin VB.ComboBox cmbHCV_AB 
         DataField       =   "HCV_AB"
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   525
         Width           =   3000
      End
      Begin VB.TextBox txtSickID2 
         BackColor       =   &H80000000&
         DataField       =   "STICKID"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   135
         Width           =   3000
      End
      Begin VB.TextBox txtCnName2 
         BackColor       =   &H80000000&
         DataField       =   "CNNAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   150
         Width           =   3000
      End
      Begin VB.TextBox txtSex2 
         BackColor       =   &H80000000&
         DataField       =   "SEX"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11025
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   150
         Width           =   3000
      End
      Begin VB.TextBox txtFOLLOW_TERM 
         DataField       =   "FOLLOW_TERM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   4
         TabIndex        =   75
         Top             =   4110
         Width           =   3000
      End
      Begin VB.TextBox txtBAL_DATE 
         DataField       =   "BAL_DATE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         TabIndex        =   74
         Top             =   3690
         Width           =   3000
      End
      Begin VB.TextBox txtCONTROL_NURSE 
         DataField       =   "CONTROL_NURSE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   73
         Top             =   3690
         Width           =   3000
      End
      Begin VB.TextBox txtCONTROL_DOCTOR 
         DataField       =   "CONTROL_DOCTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   72
         Top             =   3240
         Width           =   3000
      End
      Begin VB.TextBox txtCODE_NAME 
         DataField       =   "CODE_NAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   71
         Top             =   3240
         Width           =   3000
      End
      Begin VB.TextBox txtINTERM 
         DataField       =   "INTERM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   70
         Top             =   2775
         Width           =   3000
      End
      Begin VB.TextBox txtGRADUATE_DOCTOR 
         DataField       =   "GRADUATE_DOCTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   20
         TabIndex        =   69
         Top             =   2775
         Width           =   3000
      End
      Begin VB.TextBox txtREFRESH_DOCTOR 
         DataField       =   "REFRESH_DOCTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   68
         Top             =   2775
         Width           =   3000
      End
      Begin VB.TextBox txtINHOSPITAL_DOCTOR 
         DataField       =   "INHOSPITAL_DOCTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   67
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtATTENDING_DOCTOR 
         DataField       =   "ATTENDING_DOCTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   20
         TabIndex        =   66
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtDIRECTOR_DOCTOR 
         DataField       =   "DIRECTOR_DOCTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   65
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtDIRECTOR 
         DataField       =   "DIRECTOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11010
         MaxLength       =   20
         TabIndex        =   64
         Top             =   1875
         Width           =   3000
      End
      Begin VB.TextBox txtESC_EMER_TIMES 
         DataField       =   "ESC_EMER_TIMES"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6150
         MaxLength       =   2
         TabIndex        =   63
         Top             =   1875
         Width           =   3000
      End
      Begin VB.TextBox txtEMER_TREAT_TIMES 
         DataField       =   "EMER_TREAT_TIMES"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   62
         Top             =   1875
         Width           =   3000
      End
      Begin VB.TextBox txtALERGY_DRUGS 
         DataField       =   "ALERGY_DRUGS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   61
         Top             =   540
         Width           =   3000
      End
      Begin VB.ComboBox cmbBODY_EXAMINE_FLAG 
         DataField       =   "BODY_EXAMINE_FLAG"
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   3675
         Width           =   3000
      End
      Begin VB.Label labSickID2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人ID"
         Height          =   195
         Left            =   765
         TabIndex        =   108
         Top             =   195
         Width           =   600
      End
      Begin VB.Label labCnName2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   195
         Left            =   5655
         TabIndex        =   107
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSex2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   195
         Left            =   10515
         TabIndex        =   106
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labFOLLOW_TERM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "随诊期限"
         Height          =   195
         Left            =   10125
         TabIndex        =   102
         Top             =   4155
         Width           =   780
      End
      Begin VB.Label labFOLLOW_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "随诊标志"
         Height          =   195
         Left            =   5265
         TabIndex        =   101
         Top             =   4155
         Width           =   780
      End
      Begin VB.Label labFIRST_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "本院第一例"
         Height          =   195
         Left            =   390
         TabIndex        =   100
         Top             =   4155
         Width           =   975
      End
      Begin VB.Label labBODY_EXAMINE_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "尸检标志"
         Height          =   195
         Left            =   10125
         TabIndex        =   99
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labBAL_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "结算日期"
         Height          =   195
         Left            =   5265
         TabIndex        =   98
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labCONTROL_NURSE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "质控护师"
         Height          =   195
         Left            =   585
         TabIndex        =   97
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labCONTROL_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "质控医师"
         Height          =   195
         Left            =   10125
         TabIndex        =   96
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labMEDICAL_RECORD_MASS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病案质量"
         Height          =   195
         Left            =   5265
         TabIndex        =   95
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labCODE_NAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "编码员"
         Height          =   195
         Left            =   780
         TabIndex        =   94
         Top             =   3285
         Width           =   585
      End
      Begin VB.Label labINTERM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "实习医师"
         Height          =   195
         Left            =   10125
         TabIndex        =   93
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labGRADUATE_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "研究生实习医师"
         Height          =   195
         Left            =   4680
         TabIndex        =   92
         Top             =   2820
         Width           =   1365
      End
      Begin VB.Label labREFRESH_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "进修医师"
         Height          =   195
         Left            =   585
         TabIndex        =   91
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labINHOSPITAL_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "住院医师"
         Height          =   195
         Left            =   10125
         TabIndex        =   90
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labATTENDING_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "主治医师"
         Height          =   195
         Left            =   5265
         TabIndex        =   89
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labDIRECTOR_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "主任医师"
         Height          =   195
         Left            =   585
         TabIndex        =   88
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labDIRECTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "科主任"
         Height          =   195
         Left            =   10320
         TabIndex        =   87
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label labESC_EMER_TIMES 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "抢救成功次数"
         Height          =   195
         Left            =   4875
         TabIndex        =   86
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label labEMER_TREAT_TIMES 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "抢救次数"
         Height          =   195
         Left            =   585
         TabIndex        =   85
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labEMIT_PATHOLOGY 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "放射与病理"
         Height          =   195
         Left            =   9930
         TabIndex        =   84
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label labCLINIC_PATHOLOGY 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "临床与病理"
         Height          =   195
         Left            =   5070
         TabIndex        =   83
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label labBEFORE_AFTER_TREATMENT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "术前与术后"
         Height          =   195
         Left            =   390
         TabIndex        =   82
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label labIN_OUT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入院与出院"
         Height          =   195
         Left            =   9930
         TabIndex        =   81
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label labCLINIC_INHOSPITAL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "门诊与出院"
         Height          =   195
         Left            =   5070
         TabIndex        =   80
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label labHIV_AB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HIV_AB"
         Height          =   195
         Left            =   735
         TabIndex        =   79
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label labHCV_AB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HCV_AB"
         Height          =   195
         Left            =   10275
         TabIndex        =   78
         Top             =   585
         Width           =   630
      End
      Begin VB.Label labHBSAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HBSAG"
         Height          =   195
         Left            =   5520
         TabIndex        =   77
         Top             =   585
         Width           =   525
      End
      Begin VB.Label labALERGY_DRUGS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "过敏药物"
         Height          =   195
         Left            =   585
         TabIndex        =   76
         Top             =   585
         Width           =   780
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   9
         Left            =   14025
         TabIndex        =   227
         Top             =   1035
         Width           =   120
      End
   End
   Begin VB.PictureBox pic手术信息 
      BorderStyle     =   0  'None
      Height          =   4350
      Left            =   255
      ScaleHeight     =   4350
      ScaleWidth      =   14220
      TabIndex        =   179
      Top             =   720
      Width           =   14220
      Begin VB.ComboBox cmbHEAL3 
         DataField       =   "HEAL3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   246
         Top             =   2640
         Width           =   3000
      End
      Begin VB.ComboBox cmbHEAL2 
         DataField       =   "HEAL2"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   245
         Top             =   1815
         Width           =   3000
      End
      Begin VB.ComboBox cmbHEAL1 
         DataField       =   "HEAL1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   244
         Top             =   990
         Width           =   3000
      End
      Begin VB.ComboBox cmbANAESTHESIA_METHOD3 
         DataField       =   "ANAESTHESIA_METHOD3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   243
         Top             =   2625
         Width           =   3000
      End
      Begin VB.ComboBox cmbANAESTHESIA_METHOD2 
         DataField       =   "ANAESTHESIA_METHOD2"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   242
         Top             =   1785
         Width           =   3000
      End
      Begin VB.ComboBox cmbANAESTHESIA_METHOD1 
         DataField       =   "ANAESTHESIA_METHOD1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11010
         Style           =   2  'Dropdown List
         TabIndex        =   241
         Top             =   960
         Width           =   3000
      End
      Begin VB.ComboBox cmbWOUND_GRADE3 
         DataField       =   "WOUND_GRADE3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11025
         Style           =   2  'Dropdown List
         TabIndex        =   240
         Top             =   2220
         Width           =   3000
      End
      Begin VB.ComboBox cmbWOUND_GRADE2 
         DataField       =   "WOUND_GRADE2"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11025
         Style           =   2  'Dropdown List
         TabIndex        =   239
         Top             =   1380
         Width           =   3000
      End
      Begin VB.ComboBox cmbWOUND_GRADE1 
         DataField       =   "WOUND_GRADE1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11025
         Style           =   2  'Dropdown List
         TabIndex        =   238
         Top             =   555
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATING_DATE3 
         DataField       =   "OPERATING_DATE3"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6165
         TabIndex        =   194
         Top             =   2670
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATION_NAME3 
         DataField       =   "OPERATION_NAME3"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6165
         TabIndex        =   193
         Top             =   2250
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATION_CODE3 
         DataField       =   "OPERATION_CODE3"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   192
         Top             =   2250
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATING_DATE2 
         DataField       =   "OPERATING_DATE2"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6165
         TabIndex        =   191
         Top             =   1830
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATION_NAME2 
         DataField       =   "OPERATION_NAME2"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6165
         TabIndex        =   190
         Top             =   1380
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATION_CODE2 
         DataField       =   "OPERATION_CODE2"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   189
         Top             =   1410
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATING_DATE1 
         DataField       =   "OPERATING_DATE1"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6165
         TabIndex        =   188
         Top             =   975
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATION_NAME1 
         DataField       =   "OPERATION_NAME1"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6165
         TabIndex        =   187
         Top             =   555
         Width           =   3000
      End
      Begin VB.TextBox txtOPERATION_CODE1 
         DataField       =   "OPERATION_CODE1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   186
         Top             =   555
         Width           =   3000
      End
      Begin VB.TextBox txtSex4 
         BackColor       =   &H80000000&
         DataField       =   "SEX"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   11025
         Locked          =   -1  'True
         TabIndex        =   182
         Top             =   150
         Width           =   3000
      End
      Begin VB.TextBox txtCNNAME4 
         BackColor       =   &H80000000&
         DataField       =   "CNNAME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   6165
         Locked          =   -1  'True
         TabIndex        =   181
         Top             =   150
         Width           =   3000
      End
      Begin VB.TextBox txtSickID4 
         BackColor       =   &H80000000&
         DataField       =   "STICKID"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   180
         Top             =   150
         Width           =   3000
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   39
         Left            =   14055
         TabIndex        =   269
         Top             =   2685
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   38
         Left            =   9210
         TabIndex        =   268
         Top             =   2730
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   37
         Left            =   4500
         TabIndex        =   267
         Top             =   2745
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   36
         Left            =   14025
         TabIndex        =   266
         Top             =   2280
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   35
         Left            =   9240
         TabIndex        =   265
         Top             =   2295
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   34
         Left            =   14025
         TabIndex        =   264
         Top             =   1800
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   33
         Left            =   9225
         TabIndex        =   263
         Top             =   1875
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   32
         Left            =   4470
         TabIndex        =   262
         Top             =   1905
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   31
         Left            =   14025
         TabIndex        =   261
         Top             =   1425
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   30
         Left            =   9225
         TabIndex        =   260
         Top             =   1440
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   29
         Left            =   14025
         TabIndex        =   259
         Top             =   1005
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   28
         Left            =   9225
         TabIndex        =   258
         Top             =   1020
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   27
         Left            =   4470
         TabIndex        =   257
         Top             =   1065
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   26
         Left            =   14025
         TabIndex        =   256
         Top             =   615
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labNot 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   25
         Left            =   9225
         TabIndex        =   255
         Top             =   600
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labANAESTHESIA_METHOD3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "麻醉方法3"
         Height          =   195
         Left            =   10020
         TabIndex        =   212
         Top             =   2715
         Width           =   885
      End
      Begin VB.Label labOPERATING_DATE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术日期3"
         Height          =   195
         Left            =   5130
         TabIndex        =   211
         Top             =   2715
         Width           =   885
      End
      Begin VB.Label labHEAL3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "切口愈合情况3"
         Height          =   195
         Left            =   60
         TabIndex        =   210
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label labWOUND_GRADE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "切口等级3"
         Height          =   195
         Left            =   10020
         TabIndex        =   209
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label labOPERATION_NAME3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术名称3"
         Height          =   195
         Left            =   5130
         TabIndex        =   208
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label labOPERATION_CODE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术编码3"
         Height          =   195
         Left            =   450
         TabIndex        =   207
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label labANAESTHESIA_METHOD2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "麻醉方法2"
         Height          =   195
         Left            =   10020
         TabIndex        =   206
         Top             =   1875
         Width           =   885
      End
      Begin VB.Label labOPERATING_DATE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术日期2"
         Height          =   195
         Left            =   5130
         TabIndex        =   205
         Top             =   1875
         Width           =   885
      End
      Begin VB.Label labHEAL2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "切口愈合情况2"
         Height          =   195
         Left            =   60
         TabIndex        =   204
         Top             =   1875
         Width           =   1275
      End
      Begin VB.Label labWOUND_GRADE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "切口等级2"
         Height          =   195
         Left            =   10020
         TabIndex        =   203
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labOPERATION_NAME2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术名称2"
         Height          =   195
         Left            =   5130
         TabIndex        =   202
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labOPERATION_CODE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术编码2"
         Height          =   195
         Left            =   450
         TabIndex        =   201
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labANAESTHESIA_METHOD1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "麻醉方法1"
         Height          =   195
         Left            =   10020
         TabIndex        =   200
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label labOPERATING_DATE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术日期1"
         Height          =   195
         Left            =   5130
         TabIndex        =   199
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label labHEAL1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "切口愈合情况1"
         Height          =   195
         Left            =   60
         TabIndex        =   198
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label labWOUND_GRADE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "切口等级1"
         Height          =   195
         Left            =   10020
         TabIndex        =   197
         Top             =   600
         Width           =   885
      End
      Begin VB.Label labOPERATION_NAME1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术名称1"
         Height          =   195
         Left            =   5130
         TabIndex        =   196
         Top             =   600
         Width           =   885
      End
      Begin VB.Label labOPERATION_CODE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "手术编码1"
         Height          =   195
         Left            =   450
         TabIndex        =   195
         Top             =   600
         Width           =   885
      End
      Begin VB.Label labSex4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   195
         Left            =   10515
         TabIndex        =   185
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labCNNAME4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   195
         Left            =   5625
         TabIndex        =   184
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSTICKID4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人ID"
         Height          =   195
         Left            =   735
         TabIndex        =   183
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Frame fra病案编辑 
      Caption         =   "病案编辑"
      Height          =   5145
      Left            =   75
      TabIndex        =   213
      Top             =   120
      Width           =   14565
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   4755
         Left            =   195
         TabIndex        =   214
         Top             =   225
         Width           =   14220
         _Version        =   589884
         _ExtentX        =   25082
         _ExtentY        =   8387
         _StockProps     =   64
      End
   End
End
Attribute VB_Name = "frmMain_北京尚洋病案信息编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
'只显示30天内出院且未上传病人的信息。
Const strSickFields = "select B.病人ID as ID,B.病人ID as 病人ID,B.住院次数 As 主页ID,a.医保号, a.卡号,a.人员身份 as 人员类别,b.姓名,b.性别,b.身份证号 " & vbNewLine & _
                      "from 保险帐户 a , 病人信息 b,病案主页 c where a.病人ID = b.病人id And b.病人ID = c.病人id and b.住院次数 = c.主页id And a.险类 = [1]" & vbNewLine & _
                      "AND C.出院日期 >= sysdate-120" & vbNewLine & _
                      "And not exists (select 1 from 长治病案信息 d where a.病人ID= d.STICKID And B.住院次数=d.In_Count)"

Private mstrHospitalNumber      As String
Private mlng病人ID              As Long
Private mlng主页ID              As Long
Private mblnUpdateCenter        As Boolean

Public Property Let HospitalNumber(ByVal vNewValue As String)
    mstrHospitalNumber = vNewValue
End Property

Public Property Get HospitalNumber() As String
    HospitalNumber = mstrHospitalNumber
End Property

Public Property Let UpdateCenter(ByVal vNewValue As Boolean)
    mblnUpdateCenter = vNewValue
End Property
'==============================================================================
'=功能： 初始Tab控件
'==============================================================================
Private Function InitTabControl() As Boolean
    
    On Error GoTo ErrH
    
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With
        Set .Icons = frmPubIcons.imgPublic.Icons
        .InsertItem 0, " 基本信息 ", pic基本信息.hwnd, 0
        .InsertItem 1, " 医师信息 ", pic医师信息.hwnd, 0
        .InsertItem 2, " 诊断信息 ", pic诊断信息.hwnd, 0
        .InsertItem 3, " 手术信息 ", pic手术信息.hwnd, 0
        .Item(0).Selected = True
    End With
    InitTabControl = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 初始Tab控件
'==============================================================================
Private Sub InitCmbControl()
    On Error GoTo ErrH
    '婚姻状况
    With cmbMARITAL_STATUS
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "未婚"
        .AddItem "2" & gstrSplitCmb & "已婚"
        .AddItem "3" & gstrSplitCmb & "离婚"
        .AddItem "4" & gstrSplitCmb & "丧偶"
    End With
    '与病人关系
    With cmbRELATIONSHIP
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "配偶"
        .AddItem "2" & gstrSplitCmb & "子女"
        .AddItem "3" & gstrSplitCmb & "父母"
        .AddItem "9" & gstrSplitCmb & "其他"
    End With
    '入院病情
    With cmbPAT_ADM_CONDITION
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "危"
        .AddItem "2" & gstrSplitCmb & "急"
        .AddItem "3" & gstrSplitCmb & "一般"
        .AddItem "4" & gstrSplitCmb & "其他"
    End With
    'HBSAG
    With cmbHBSAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    'HCV_AB
    With cmbHCV_AB
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    'HIV_AB
    With cmbHIV_AB
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    '门诊与出院
    With cmdCLINIC_INHOSPITAL
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    '入院与出院
    With cmbIN_OUT
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    '术前与术后
    With cmbBEFORE_AFTER_TREATMENT
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    '临床与病理
    With cmbCLINIC_PATHOLOGY
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    '放射与病理
    With cmbEMIT_PATHOLOGY
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "符合"
        .AddItem "2" & gstrSplitCmb & "不符"
        .AddItem "3" & gstrSplitCmb & "未定"
    End With
    '病案质量
    With cmbMEDICAL_RECORD_MASS
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "甲"
        .AddItem "2" & gstrSplitCmb & "乙"
        .AddItem "3" & gstrSplitCmb & "丙"
    End With
    '本院第一例
    With cmbFIRST_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "是"
        .AddItem "2" & gstrSplitCmb & "否"
    End With
    '尸检标志
    With cmbBODY_EXAMINE_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "是"
        .AddItem "2" & gstrSplitCmb & "否"
    End With
    '随诊标志
    With cmbFOLLOW_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "是"
        .AddItem "2" & gstrSplitCmb & "否"
    End With
    '示教病例
    With cmbTEACH_MR_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "是"
        .AddItem "2" & gstrSplitCmb & "否"
    End With
    '血型标志
    With cmbBLOOD_TYPE
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "A"
        .AddItem "2" & gstrSplitCmb & "B"
        .AddItem "3" & gstrSplitCmb & "AB"
        .AddItem "4" & gstrSplitCmb & "O"
        .AddItem "5" & gstrSplitCmb & "其它"
    End With
    'RH
    With cmbRH
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "阴"
        .AddItem "2" & gstrSplitCmb & "阳"
    End With
    '输入血反应标志
    With cmbBLOOD_TRAN_REACT_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "有"
        .AddItem "2" & gstrSplitCmb & "无"
    End With
    '治疗结果1
    With cmbTREAT_RESULT1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "治愈"
        .AddItem "2" & gstrSplitCmb & "好转"
        .AddItem "3" & gstrSplitCmb & "未愈"
        .AddItem "4" & gstrSplitCmb & "死亡"
        .AddItem "5" & gstrSplitCmb & "其它"
    End With
    '治疗结果2
    With cmbTREAT_RESULT2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "治愈"
        .AddItem "2" & gstrSplitCmb & "好转"
        .AddItem "3" & gstrSplitCmb & "未愈"
        .AddItem "4" & gstrSplitCmb & "死亡"
        .AddItem "5" & gstrSplitCmb & "其它"
    End With
    '治疗结果3
    With cmbTREAT_RESULT3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "治愈"
        .AddItem "2" & gstrSplitCmb & "好转"
        .AddItem "3" & gstrSplitCmb & "未愈"
        .AddItem "4" & gstrSplitCmb & "死亡"
        .AddItem "5" & gstrSplitCmb & "其它"
    End With
    '麻醉方式1
    With cmbANAESTHESIA_METHOD1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "全麻"
        .AddItem "2" & gstrSplitCmb & "硬膜外"
        .AddItem "3" & gstrSplitCmb & "局麻"
    End With
    '麻醉方式2
    With cmbANAESTHESIA_METHOD2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "全麻"
        .AddItem "2" & gstrSplitCmb & "硬膜外"
        .AddItem "3" & gstrSplitCmb & "局麻"
    End With
    '麻醉方式3
    With cmbANAESTHESIA_METHOD3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "全麻"
        .AddItem "2" & gstrSplitCmb & "硬膜外"
        .AddItem "3" & gstrSplitCmb & "局麻"
    End With
    '切口等级1
    With cmbWOUND_GRADE1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "I"
        .AddItem "2" & gstrSplitCmb & "II"
        .AddItem "3" & gstrSplitCmb & "III"
    End With
    '切口等级2
    With cmbWOUND_GRADE2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "I"
        .AddItem "2" & gstrSplitCmb & "II"
        .AddItem "3" & gstrSplitCmb & "III"
    End With
    '切口等级3
    With cmbWOUND_GRADE3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "I"
        .AddItem "2" & gstrSplitCmb & "II"
        .AddItem "3" & gstrSplitCmb & "III"
    End With
    '切口愈合情况1
    With cmbHEAL1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "甲"
        .AddItem "2" & gstrSplitCmb & "乙"
        .AddItem "3" & gstrSplitCmb & "丙"
    End With
    '切口愈合情况2
    With cmbHEAL2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "甲"
        .AddItem "2" & gstrSplitCmb & "乙"
        .AddItem "3" & gstrSplitCmb & "丙"
    End With
    '切口愈合情况3
    With cmbHEAL3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "甲"
        .AddItem "2" & gstrSplitCmb & "乙"
        .AddItem "3" & gstrSplitCmb & "丙"
    End With
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub cmdSick_Click()
    gstrSQL = strSickFields
    Call SickSelect(gstrSQL)
    MsgBox gstrSQL
End Sub

Private Sub txtSickID1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode <> 13 Then Exit Sub
    Dim strCode As String, strWhere As String
    strCode = txtSickID1.Text
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then
        '病人ID
        strWhere = " And A.病人ID=" & Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then
        '住院号
        strWhere = " And b.住院号='" & Mid(strCode, 2) & "'"
    Else
        '医保号
        strWhere = " And (b.姓名 Like '%" & strCode & "%' or A.医保号 like '%" & strCode & "%')"
    End If
    gstrSQL = strSickFields & vbCrLf & strWhere
    Call SickSelect(gstrSQL)
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub SickSelect(sSql As String)
    Dim vRect       As RECT
    Dim sngX        As Single
    Dim sngY        As Single
    Dim sngH        As Single
    Dim rsTmp       As ADODB.Recordset
    On Error GoTo ErrH
    vRect = GetControlRect(txtSickID1.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txtSickID1.Height
    Set rsTmp = zlDatabase.ShowSQLSelect( _
            Nothing, sSql, 0, "医保病种选择", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, TYPE_北京尚洋, txtSickID1.Text _
            )
    If Not ChkRsState(rsTmp) Then
        txtSickID1.Text = Nvl(rsTmp!病人ID)
        txtSickID1.Tag = Nvl(rsTmp!病人ID) & gstrSplitCmb & Nvl(rsTmp!主页ID)
        mlng病人ID = Nvl(rsTmp!病人ID)
        mlng主页ID = Nvl(rsTmp!主页ID)
        txtCnName1.Text = Nvl(rsTmp!姓名)
        txtSex1.Text = Nvl(rsTmp!性别)
        txtSickID2.Text = Nvl(rsTmp!病人ID)
        txtCnName2.Text = Nvl(rsTmp!姓名)
        txtSex2.Text = Nvl(rsTmp!性别)
        txtSickID3.Text = Nvl(rsTmp!病人ID)
        txtCnName3.Text = Nvl(rsTmp!姓名)
        txtSex3.Text = Nvl(rsTmp!性别)
        txtSickID4.Text = Nvl(rsTmp!病人ID)
        txtCNNAME4.Text = Nvl(rsTmp!姓名)
        txtSex4.Text = Nvl(rsTmp!性别)
    Else
        MsgBox "没有找到病人信息!", vbInformation, gstrSysName
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
Private Sub Form_Load()
    Dim rsTmp   As ADODB.Recordset
    Dim objCtr  As Control
    Dim strField As String
    On Error GoTo ErrH
    '初始化
    Call InitTabControl
    Call InitCmbControl
    If mstrHospitalNumber <> "" Then
        '修改
        gstrSQL = "Select * from 长治病案信息 Where RESIDENCE_NO=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrHospitalNumber)
        '文本框赋值
        If Not ChkRsState(rsTmp) Then
            With rsTmp
                For Each objCtr In Me.Controls
                    Select Case TypeName(objCtr)
                        Case "TextBox"
                            strField = objCtr.DataField
                            objCtr.Text = Nvl(.Fields(strField))
                        Case "ComboBox"
                            strField = objCtr.DataField
                            objCtr.ListIndex = Cmb_EditIndex(objCtr, Nvl(.Fields(strField)))
                    End Select
                Next
                txtSickID1.Tag = txtSickID1.Text
                mlng病人ID = txtSickID1.Tag
                mlng主页ID = Nvl(!IN_COUNT)
            End With
        End If
    End If
    cmdRead.Enabled = Not mblnUpdateCenter
    cmdOK.Enabled = Not mblnUpdateCenter
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRead_Click()
    Dim rsTmp           As ADODB.Recordset
    Dim str过敏药物      As String
On Error GoTo ErrH
    If txtSickID1.Tag = "" Then
        MsgBox "请选择病人！", vbCritical, gstrSysName
        Exit Sub
    End If
    '==============================================================================
    '=基本信息
    '==============================================================================
    txtHOSPITAL_NUMBER.Text = gstr医院编码
    #If gverControl < 6 Then
        '取病人基本信息
        gstrSQL = " SELECT A.住院号,B.医疗付款方式,B.主页ID,B.病案号,A.姓名,A.性别,A.出生日期,B.婚姻状况,B.职业,A.出生地点," & _
             "        H.编码 AS 民族,B.国籍,A.身份证号,A.工作单位,B.单位地址,B.单位电话,B.单位邮编,B.家庭地址,B.户口邮编," & _
             "        B.联系人姓名,B.联系人关系,B.联系人地址,B.联系人电话,B.入院日期,D.名称 AS 入院科室,E.名称 AS 入院病区," & _
             "        B.出院日期,F.名称 AS 出院科室,B.入院病况,B.确诊日期,B.抢救次数,B.成功次数,B.出院方式," & _
             "        B.编目员姓名,NVL(B.编目日期,SYSDATE) AS 编目日期,B.尸检标志,B.随诊标志,B.随诊期限,B.血型,B.住院医师" & _
             " FROM 病人信息 A,病案主页 B,合约单位 C,部门表 D,部门表 E,部门表 F,民族 H" & _
             " WHERE A.病人ID=B.病人ID AND A.住院次数=B.主页ID AND A.合同单位ID=C.ID(+)" & _
             " AND B.入院科室ID=D.ID(+) AND B.入院病区ID=E.ID(+) AND B.出院科室ID=F.ID(+) " & _
             " AND A.民族=H.名称 AND B.病人ID = [1] AND B.主页ID = [2]"
    #Else
        '取病人基本信息
        gstrSQL = " SELECT A.住院号,B.医疗付款方式,B.主页ID,B.病案号,A.姓名,A.性别,A.出生日期,B.婚姻状况,B.职业,A.出生地点," & _
             "        H.编码 AS 民族,B.国籍,A.身份证号,A.工作单位,B.单位地址,B.单位电话,B.单位邮编,B.家庭地址,B.家庭地址邮编 As 户口邮编," & _
             "        B.联系人姓名,B.联系人关系,B.联系人地址,B.联系人电话,B.入院日期,D.名称 AS 入院科室,E.名称 AS 入院病区," & _
             "        B.出院日期,F.名称 AS 出院科室,B.入院病况,B.确诊日期,B.抢救次数,B.成功次数,B.出院方式," & _
             "        B.编目员姓名,NVL(B.编目日期,SYSDATE) AS 编目日期,B.尸检标志,B.随诊标志,B.随诊期限,B.血型,B.住院医师" & _
             " FROM 病人信息 A,病案主页 B,合约单位 C,部门表 D,部门表 E,部门表 F,民族 H" & _
             " WHERE A.病人ID=B.病人ID AND A.住院次数=B.主页ID AND A.合同单位ID=C.ID(+)" & _
             " AND B.入院科室ID=D.ID(+) AND B.入院病区ID=E.ID(+) AND B.出院科室ID=F.ID(+) " & _
             " AND A.民族=H.名称 AND B.病人ID = [1] AND B.主页ID = [2]"
    #End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTmp
        'txtRESIDENCE_NO.Text = Nvl(!住院号) & "_" & Nvl(!主页ID)
        '小范更新于2011-06-29日住院号不提取主页ID不取 住院次数
              
        txtRESIDENCE_NO.Text = Nvl(!住院号)
        txtIN_COUNT.Text = !主页ID
        txtMEDICAL_RECORD_NO.Text = Nvl(!住院号)
        cmbMARITAL_STATUS.ListIndex = Cmb_EditIndex(cmbMARITAL_STATUS, TRANDATA("婚姻", Nvl(!婚姻状况, "未婚")))
        txtSTATUS.Text = ChkStrUniCode(Nvl(!职业), txtSTATUS.MaxLength)
        txtBIRTH_ADDRESS.Text = ChkStrUniCode(Nvl(!出生地点), txtBIRTH_ADDRESS.MaxLength)
        txtIDENTITY_NUMBER.Text = ChkStrUniCode(Nvl(!身份证号), txtIDENTITY_NUMBER.MaxLength)
        txtUNIT_NAME.Text = ChkStrUniCode(Nvl(!工作单位), txtUNIT_NAME.MaxLength)
        txtUNIT_ADDRESS.Text = ChkStrUniCode(Nvl(!单位地址), txtUNIT_ADDRESS.MaxLength)
        txtUNIT_PHONE.Text = ChkStrUniCode(Nvl(!单位电话), txtUNIT_PHONE.MaxLength)
        txtUNIT_ZIPCODE.Text = ChkStrUniCode(Nvl(!单位邮编), txtUNIT_ZIPCODE.MaxLength)
        txtREGISTER_ADDRESS.Text = ChkStrUniCode(Nvl(!家庭地址), txtREGISTER_ADDRESS.MaxLength)
        txtREGISTER_ZIPCODE.Text = ChkStrUniCode(Nvl(!户口邮编), txtREGISTER_ZIPCODE.MaxLength)
        txtCONTACT_PERSON.Text = ChkStrUniCode(Nvl(!联系人姓名), txtCONTACT_PERSON.MaxLength)
        cmbRELATIONSHIP.ListIndex = Cmb_EditIndex(cmbMARITAL_STATUS, TRANDATA("与病人关系", Nvl(!联系人关系)))
        txtCONTACT_ADDRESS.Text = ChkStrUniCode(Nvl(!联系人地址), txtCONTACT_ADDRESS.MaxLength)
        txtCONTACT_PHONE.Text = ChkStrUniCode(Nvl(!联系人电话), txtCONTACT_PHONE.MaxLength)
        txtADMISSION_DATE.Text = Format(!入院日期, "YYYY-MM-DD HH:MM:SS")
        txtADMISSION_DEPT.Text = ChkStrUniCode(Nvl(!入院科室), txtADMISSION_DEPT.MaxLength)
        txtIN_DEPT_ZONE.Text = ChkStrUniCode(Nvl(!入院病区), txtIN_DEPT_ZONE.MaxLength)
        txtDEPT_TRANSFERED_TO.Text = "" 'ChkStrUniCode("", txtDEPT_TRANSFERED_TO.MaxLength)
        txtDISCHARGE_DATE.Text = Format(!出院日期, "YYYY-MM-DD HH:MM:SS")
        txtDISCHARGE_DEPT.Text = ChkStrUniCode(Nvl(!出院科室), txtDISCHARGE_DEPT.MaxLength)
        txtOUT_DEPT_ZONE.Text = ChkStrUniCode(Nvl(!出院科室), txtDISCHARGE_DEPT.MaxLength)
        cmbPAT_ADM_CONDITION.ListIndex = Cmb_EditIndex(cmbPAT_ADM_CONDITION, TRANDATA("入院病情", Nvl(!入院病况)))
        txtDIAGNOSIS_DATE.Text = Format(!确诊日期, "YYYY-MM-DD HH:MM:SS")
        '医师信息
        txtINHOSPITAL_DOCTOR = ChkStrUniCode(Nvl(!住院医师), txtINHOSPITAL_DOCTOR.MaxLength)
        txtINTERM.Text = ChkStrUniCode(Nvl(!编目员姓名), txtINTERM.MaxLength)
        cmbBODY_EXAMINE_FLAG.ListIndex = Cmb_EditIndex(cmbBODY_EXAMINE_FLAG, TRANDATA("尸检标志", Nvl(!尸检标志, "否")))
        cmbFOLLOW_FLAG.ListIndex = Cmb_EditIndex(cmbFOLLOW_FLAG, TRANDATA("随诊标志", Nvl(!随诊标志)))
        txtFOLLOW_TERM.Text = ChkStrUniCode(Nvl(!随诊期限, 0), txtFOLLOW_TERM.MaxLength)
        '诊断信息
        cmbBLOOD_TYPE.ListIndex = Cmb_EditIndex(cmbBLOOD_TYPE, TRANDATA("血型", Nvl(!血型)))
        txtHANDLE.Text = ChkStrUniCode(Nvl(!编目员姓名, UserInfo.姓名), txtHANDLE.MaxLength)
        txtHANDLE_DATE.Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    End With
    '==============================================================================
    '=医师信息
    '==============================================================================
    str过敏药物 = ""
    '任选一种过敏药物
    gstrSQL = " SELECT 过敏药物 FROM 病人过敏药物 WHERE 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID)
    Do While Not ChkRsState(rsTmp)
        str过敏药物 = str过敏药物 & " ," & Trim(Nvl(rsTmp!过敏药物))
        rsTmp.MoveNext
    Loop
    txtALERGY_DRUGS.Text = ChkStrUniCode(Mid(str过敏药物, 2), txtALERGY_DRUGS.MaxLength)
    
    '取病案主页从表
    Dim STR信息值 As String
    gstrSQL = "SELECT UPPER(信息名) AS 信息名,信息值 FROM 病案主页从表 WHERE 病人ID=[1] AND 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTmp
        Do While Not .EOF
            STR信息值 = Nvl(!信息值)
            Select Case !信息名
                Case "HBSAG"
                    cmbHBSAG.ListIndex = Cmb_EditIndex(cmbHBSAG, TRANDATA("HBSAG", STR信息值))
                Case "HCV-AB"
                    cmbHCV_AB.ListIndex = Cmb_EditIndex(cmbHCV_AB, TRANDATA("HCV-AB", STR信息值))
                Case "HIV-AB"
                    cmbHIV_AB.ListIndex = Cmb_EditIndex(cmbHIV_AB, TRANDATA("HIV-AB", STR信息值))
                Case "科主任"
                    txtDIRECTOR.Text = ChkStrUniCode(STR信息值, txtDIRECTOR.MaxLength)
                Case "主任医师"
                    txtDIRECTOR_DOCTOR.Text = ChkStrUniCode(STR信息值, txtDIRECTOR_DOCTOR.MaxLength)
                Case "主治医师"
                    txtATTENDING_DOCTOR.Text = ChkStrUniCode(STR信息值, txtATTENDING_DOCTOR.MaxLength)
                Case "进修医师"
                    txtREFRESH_DOCTOR.Text = ChkStrUniCode(STR信息值, txtREFRESH_DOCTOR.MaxLength)
                Case "研究生实习医师"
                    txtGRADUATE_DOCTOR.Text = ChkStrUniCode(STR信息值, txtGRADUATE_DOCTOR.MaxLength)
                Case "实习医师"
                    txtINTERM.Text = ChkStrUniCode(STR信息值, txtINTERM.MaxLength)
'                Case "病案质量"
'                    cmbMEDICAL_RECORD_MASS.ListIndex = Cmb_EditIndex(cmbMEDICAL_RECORD_MASS, TRANDATA("病案质量", STR信息值))
                Case "质控医师"
                    txtCONTROL_DOCTOR.Text = ChkStrUniCode(STR信息值, txtCONTROL_DOCTOR.MaxLength)
                Case "质控护师"
                    txtCONTROL_NURSE.Text = ChkStrUniCode(STR信息值, txtCONTROL_NURSE.MaxLength)
                Case "首例"
                    cmbFIRST_FLAG.ListIndex = Cmb_EditIndex(cmbFIRST_FLAG, TRANDATA("首例", STR信息值))
                Case "示教病案"
                    cmbTEACH_MR_FLAG.ListIndex = Cmb_EditIndex(cmbTEACH_MR_FLAG, TRANDATA("示教病例", STR信息值))
                Case "RH"
                    cmbRH.ListIndex = Cmb_EditIndex(cmbRH, TRANDATA("RH", STR信息值))
                Case "输血反应"
                    cmbBLOOD_TRAN_REACT_FLAG.ListIndex = Cmb_EditIndex(cmbBLOOD_TRAN_REACT_FLAG, TRANDATA("输血反应", STR信息值))
                Case "输红细胞"
                    txtERYTHROCYTE.Text = IIf(Val(STR信息值) < 90009000.99 Or Val(STR信息值) > 0, Val(STR信息值), 0)
                Case "输血小板"
                    txtHEMOBLAST.Text = IIf(Val(STR信息值) < 90009000.99 Or Val(STR信息值) > 0, Val(STR信息值), 0)
                Case "输血浆"
                    txtPLASM.Text = IIf(Val(STR信息值) < 90009000.99 Or Val(STR信息值) > 0, Val(STR信息值), 0)
                Case "输全血"
                    txtBLOOD.Text = IIf(Val(STR信息值) < 90009000.99 Or Val(STR信息值) > 0, Val(STR信息值), 0)
                Case "输其他"
                    txtOTHER_BLOOD.Text = IIf(Val(STR信息值) < 90009000.99 Or Val(STR信息值) > 0, Val(STR信息值), 0)
            End Select
            .MoveNext
        Loop
    End With
    
    '取诊断情况
    gstrSQL = "SELECT 符合类型,NVL(符合情况,0) AS 符合情况 FROM 诊断符合情况 WHERE 病人ID=[1] AND 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTmp
        Do While Not ChkRsState(rsTmp)
            Select Case !符合类型
                Case 1  '门诊与出院
                    cmdCLINIC_INHOSPITAL.ListIndex = Cmb_EditIndex(cmdCLINIC_INHOSPITAL, Nvl(!符合情况))
                Case 2  '入院与出院
                    cmbIN_OUT.ListIndex = Cmb_EditIndex(cmbIN_OUT, Nvl(!符合情况))
                Case 3  '放射与病理
                    cmbEMIT_PATHOLOGY.ListIndex = Cmb_EditIndex(cmbEMIT_PATHOLOGY, Nvl(!符合情况))
                Case 4  '临床与病理
                    cmbCLINIC_PATHOLOGY.ListIndex = Cmb_EditIndex(cmbCLINIC_PATHOLOGY, Nvl(!符合情况))
                Case 6  '术前与术后
                    cmbBEFORE_AFTER_TREATMENT.ListIndex = Cmb_EditIndex(cmbBEFORE_AFTER_TREATMENT, Nvl(!符合情况))
            End Select
            .MoveNext
        Loop
    End With

    '取病案评分结果（从病案主页从表中读取，只要填写了病案的都会有数据）
    gstrSQL = "SELECT 等级 FROM 病案评分结果 WHERE 病人ID=[1] AND 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not ChkRsState(rsTmp) Then
       cmbMEDICAL_RECORD_MASS.ListIndex = Cmb_EditIndex(cmbMEDICAL_RECORD_MASS, rsTmp!等级)
    End If
    '取结算数据
    gstrSQL = "SELECT 操作员姓名,收费时间 FROM 病人结帐记录 WHERE ID = (SELECT MAX(ID) FROM 病人结帐记录 WHERE 病人ID=[1] AND 记录状态=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not ChkRsState(rsTmp) Then
        txtBAL_DATE.Text = Format(rsTmp!收费时间, "YYYY-MM-DD HH:MM:SS")
    End If
    '==============================================================================
    '=诊断数据
    '==============================================================================
    '入院诊断
'    gstrSQL = " SELECT A.诊断类型,A.诊断次序,A.编码序号,B.编码 AS 疾病编码,A.诊断描述,A.记录人,NVL(A.记录日期,SYSDATE) AS 记录日期" & _
'             " FROM 病人诊断记录 A,疾病编码目录 B" & _
'             " WHERE A.疾病ID=B.ID AND A.记录来源=3 AND A.诊断类型 In ('2','12') AND A.病人ID=[1] AND A.主页ID=[2]"
' 小范更新
gstrSQL = " SELECT A.诊断类型,A.诊断次序,A.编码序号,B.编码 AS 疾病编码,A.诊断描述,A.记录人,NVL(A.记录日期,SYSDATE) AS 记录日期" & _
             " FROM 病人诊断记录 A,疾病编码目录 B" & _
             " WHERE A.疾病ID=B.ID and A.记录来源=4 AND A.诊断类型 In ('2','12') and A.病人ID=[1] AND A.主页ID=[2]"
             
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not ChkRsState(rsTmp) Then
        txtIN_DIAGNOSIS_CODE.Text = ChkStrUniCode(Nvl(rsTmp!疾病编码), txtIN_DIAGNOSIS_CODE.MaxLength)
        txtIN_DIAGNOSIS_NAME.Text = ChkStrUniCode(Nvl(rsTmp!诊断描述), txtIN_DIAGNOSIS_NAME.MaxLength)
        txtIN_DIAGNOSIS_DATE.Text = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
    End If
    '出院诊断
    gstrSQL = " SELECT A.诊断类型,A.诊断次序,A.编码序号,B.编码 AS 疾病编码,A.诊断描述,A.记录人,NVL(A.记录日期,SYSDATE) AS 记录日期,出院情况" & _
             " FROM 病人诊断记录 A,疾病编码目录 B" & _
             " WHERE A.疾病ID=B.ID AND A.记录来源=4 AND A.诊断类型 In ('3','13') AND A.病人ID=[1] AND A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Do While Not ChkRsState(rsTmp)
        If rsTmp.Bookmark = 1 Then
            txtOUT_DIAGNOSIS_CODE1.Text = ChkStrUniCode(Nvl(rsTmp!疾病编码), txtOUT_DIAGNOSIS_CODE1.MaxLength)
            txtOUT_DIAGNOSIS_NAME1.Text = ChkStrUniCode(Nvl(rsTmp!诊断描述), txtOUT_DIAGNOSIS_NAME1.MaxLength)
            txtOUT_DIAGNOSIS_DATE1.Text = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
            cmbTREAT_RESULT1.ListIndex = Cmb_EditIndex(cmbTREAT_RESULT1, TRANDATA("出院情况", Nvl(rsTmp!出院情况)))
        ElseIf rsTmp.Bookmark = 2 Then
            txtOUT_DIAGNOSIS_CODE2.Text = ChkStrUniCode(Nvl(rsTmp!疾病编码), txtOUT_DIAGNOSIS_CODE2.MaxLength)
            txtOUT_DIAGNOSIS_NAME2.Text = ChkStrUniCode(Nvl(rsTmp!诊断描述), txtOUT_DIAGNOSIS_NAME2.MaxLength)
            txtOUT_DIAGNOSIS_DATE2.Text = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
            cmbTREAT_RESULT2.ListIndex = Cmb_EditIndex(cmbTREAT_RESULT2, TRANDATA("出院情况", Nvl(rsTmp!出院情况)))
        ElseIf rsTmp.Bookmark = 3 Then
            txtOUT_DIAGNOSIS_CODE3.Text = ChkStrUniCode(Nvl(rsTmp!疾病编码), txtOUT_DIAGNOSIS_CODE3.MaxLength)
            txtOUT_DIAGNOSIS_NAME3.Text = ChkStrUniCode(Nvl(rsTmp!诊断描述), txtOUT_DIAGNOSIS_NAME3.MaxLength)
            txtOUT_DIAGNOSIS_DATE3.Text = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
            cmbTREAT_RESULT3.ListIndex = Cmb_EditIndex(cmbTREAT_RESULT3, TRANDATA("出院情况", Nvl(rsTmp!出院情况)))
        End If
        rsTmp.MoveNext
    Loop
    '==============================================================================
    '=手术信息
    '==============================================================================
    gstrSQL = " SELECT B.编码,B.名称,A.切口,A.愈合,A.手术日期,A.麻醉类型,A.主刀医师,A.第一助手,A.第二助手,A.麻醉医师,A.记录人,NVL(A.记录日期,SYSDATE) AS 记录日期 " & _
             " FROM 病人手麻记录 A ,疾病编码目录 B " & _
             " WHERE A.手术操作ID=B.ID And A.病人ID=[1] AND A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Do While Not ChkRsState(rsTmp)
        If rsTmp.Bookmark = 1 Then
            txtOPERATION_CODE1.Text = ChkStrUniCode(Nvl(rsTmp!编码), txtOPERATION_CODE1.MaxLength)
            txtOPERATION_NAME1.Text = ChkStrUniCode(Nvl(rsTmp!名称), txtOPERATION_NAME1.MaxLength)
            cmbWOUND_GRADE1.ListIndex = Cmb_EditIndex(cmbWOUND_GRADE1, TRANDATA("切口", Nvl(rsTmp!切口)))
            cmbHEAL1.ListIndex = Cmb_EditIndex(cmbHEAL1, TRANDATA("愈合", Nvl(rsTmp!愈合)))
            txtOPERATING_DATE1 = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
            cmbANAESTHESIA_METHOD1.ListIndex = Cmb_EditIndex(cmbANAESTHESIA_METHOD1, TRANDATA("麻醉类型", Nvl(rsTmp!麻醉类型)))
            
        ElseIf rsTmp.Bookmark = 2 Then
            txtOPERATION_CODE2.Text = ChkStrUniCode(Nvl(rsTmp!编码), txtOPERATION_CODE2.MaxLength)
            txtOPERATION_NAME2.Text = ChkStrUniCode(Nvl(rsTmp!名称), txtOPERATION_NAME2.MaxLength)
            cmbWOUND_GRADE2.ListIndex = Cmb_EditIndex(cmbWOUND_GRADE2, TRANDATA("切口", Nvl(rsTmp!切口)))
            cmbHEAL2.ListIndex = Cmb_EditIndex(cmbHEAL2, TRANDATA("愈合", Nvl(rsTmp!愈合)))
            txtOPERATING_DATE2.Text = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
            cmbANAESTHESIA_METHOD2.ListIndex = Cmb_EditIndex(cmbANAESTHESIA_METHOD2, TRANDATA("麻醉类型", Nvl(rsTmp!麻醉类型)))
        ElseIf rsTmp.Bookmark = 3 Then
            txtOPERATION_CODE3.Text = ChkStrUniCode(Nvl(rsTmp!编码), txtOPERATION_CODE3.MaxLength)
            txtOPERATION_NAME3.Text = ChkStrUniCode(Nvl(rsTmp!名称), txtOPERATION_NAME3.MaxLength)
            cmbWOUND_GRADE3.ListIndex = Cmb_EditIndex(cmbWOUND_GRADE3, TRANDATA("切口", Nvl(rsTmp!切口)))
            cmbHEAL3.ListIndex = Cmb_EditIndex(cmbHEAL3, TRANDATA("愈合", Nvl(rsTmp!愈合)))
            txtOPERATING_DATE3.Text = Format(Nvl(rsTmp!记录日期), "yyyy-mm-dd hh:mm:ss")
            cmbANAESTHESIA_METHOD3.ListIndex = Cmb_EditIndex(cmbANAESTHESIA_METHOD3, TRANDATA("麻醉类型", Nvl(rsTmp!麻醉类型)))
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrH:
 MsgBox Err.Description, vbCritical, gstrSysName
    If ErrCenter() = 1 Then
                Resume
            End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

Private Sub SetNotText(labNot As Label, txtNot As TextBox)
On Error GoTo ErrH
    labNot.Caption = IIf(Len(txtNot.Text) > 0, "√", "*")
    labNot.ForeColor = IIf(Len(txtNot.Text) > 0, vbGreen, vbRed)
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub SetNotComb(labNot As Label, cmbNot As ComboBox)
On Error GoTo ErrH
    labNot.Caption = IIf(Len(cmbNot.Text) > 0, "√", "*")
    labNot.ForeColor = IIf(Len(cmbNot.Text) > 0, vbGreen, vbRed)
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHOSPITAL_NUMBER_Change()
    Call SetNotText(labNot(1), txtHOSPITAL_NUMBER)
End Sub

Private Sub txtRESIDENCE_NO_Change()
    Call SetNotText(labNot(2), txtRESIDENCE_NO)
End Sub

Private Sub txtIN_COUNT_Change()
    Call SetNotText(labNot(3), txtIN_COUNT)
End Sub

Private Sub txtMEDICAL_RECORD_NO_Change()
    Call SetNotText(labNot(4), txtMEDICAL_RECORD_NO)
End Sub

Private Sub txtADMISSION_DATE_Change()
    labNot(5).Caption = IIf(IsDate(txtADMISSION_DATE.Text), "√", "*")
    labNot(5).ForeColor = IIf(IsDate(txtADMISSION_DATE.Text), vbGreen, vbRed)
End Sub

Private Sub txtADMISSION_DEPT_Change()
    Call SetNotText(labNot(6), txtADMISSION_DEPT)
End Sub

Private Sub txtIN_DEPT_ZONE_Change()
    Call SetNotText(labNot(7), txtIN_DEPT_ZONE)
End Sub

Private Sub txtDISCHARGE_DATE_Change()
    Call SetNotText(labNot(8), txtDISCHARGE_DATE)
End Sub

Private Sub cmbIN_OUT_Click()
    Call SetNotComb(labNot(9), cmbIN_OUT)
End Sub

Private Sub txtHANDLE_Change()
    Call SetNotText(labNot(10), txtHANDLE)
End Sub

Private Sub txtHANDLE_DATE_Change()
    labNot(11).Caption = IIf(IsDate(txtHANDLE_DATE.Text), "√", "*")
    labNot(11).ForeColor = IIf(IsDate(txtHANDLE_DATE.Text), vbGreen, vbRed)
End Sub

Private Sub txtIN_DIAGNOSIS_CODE_Change()
    Call SetNotText(labNot(12), txtIN_DIAGNOSIS_CODE)
End Sub

Private Sub txtIN_DIAGNOSIS_NAME_Change()
    Call SetNotText(labNot(13), txtIN_DIAGNOSIS_NAME)
End Sub

Private Sub txtIN_DIAGNOSIS_DATE_Change()
    labNot(14).Caption = IIf(IsDate(txtIN_DIAGNOSIS_DATE.Text), "√", "*")
    labNot(14).ForeColor = IIf(IsDate(txtIN_DIAGNOSIS_DATE.Text), vbGreen, vbRed)
End Sub

Private Sub txtOUT_DIAGNOSIS_CODE1_Change()
    Call SetNotText(labNot(15), txtOUT_DIAGNOSIS_CODE1)
End Sub

Private Sub txtOUT_DIAGNOSIS_NAME1_Change()
    Call SetNotText(labNot(16), txtOUT_DIAGNOSIS_NAME1)
End Sub

Private Sub txtOUT_DIAGNOSIS_DATE1_Change()
    labNot(17).Caption = IIf(IsDate(txtOUT_DIAGNOSIS_DATE1.Text), "√", "*")
    labNot(17).ForeColor = IIf(IsDate(txtOUT_DIAGNOSIS_DATE1.Text), vbGreen, vbRed)
End Sub

Private Sub cmbTREAT_RESULT1_Click()
    Call SetNotComb(labNot(18), cmbTREAT_RESULT1)
End Sub

Private Sub txtOUT_DIAGNOSIS_CODE2_Change()
    If Len(Trim(txtOUT_DIAGNOSIS_CODE2.Text & txtOUT_DIAGNOSIS_CODE3.Text)) > 0 Then
        labNot(19).Visible = True
        labNot(20).Visible = True
        labNot(21).Visible = True
        txtOUT_DIAGNOSIS_NAME2.Enabled = True
        txtOUT_DIAGNOSIS_DATE2.Enabled = True
        cmbTREAT_RESULT2.Enabled = True
    Else
        labNot(19).Visible = False
        labNot(20).Visible = False
        labNot(21).Visible = False
        txtOUT_DIAGNOSIS_NAME2.Text = ""
        txtOUT_DIAGNOSIS_NAME2.Enabled = False
        txtOUT_DIAGNOSIS_DATE2.Text = ""
        txtOUT_DIAGNOSIS_DATE2.Enabled = False
        cmbTREAT_RESULT2.ListIndex = -1
        cmbTREAT_RESULT2.Enabled = False
    End If
End Sub

Private Sub txtOUT_DIAGNOSIS_NAME2_Change()
    Call SetNotText(labNot(19), txtOUT_DIAGNOSIS_NAME2)
End Sub

Private Sub txtOUT_DIAGNOSIS_DATE2_Change()
    labNot(20).Caption = IIf(IsDate(txtOUT_DIAGNOSIS_DATE2.Text), "√", "*")
    labNot(20).ForeColor = IIf(IsDate(txtOUT_DIAGNOSIS_DATE2.Text), vbGreen, vbRed)
End Sub

Private Sub cmbTREAT_RESULT2_Click()
    Call SetNotComb(labNot(21), cmbTREAT_RESULT2)
End Sub

Private Sub txtOUT_DIAGNOSIS_CODE3_Change()
    If Len(Trim(txtOUT_DIAGNOSIS_CODE3.Text)) > 0 Then
        labNot(22).Visible = True
        labNot(23).Visible = True
        labNot(24).Visible = True
        txtOUT_DIAGNOSIS_NAME3.Enabled = True
        txtOUT_DIAGNOSIS_DATE3.Enabled = True
        cmbTREAT_RESULT3.Enabled = True
    Else
        labNot(22).Visible = False
        labNot(23).Visible = False
        labNot(24).Visible = False
        txtOUT_DIAGNOSIS_NAME3.Text = ""
        txtOUT_DIAGNOSIS_NAME3.Enabled = False
        txtOUT_DIAGNOSIS_DATE3.Text = ""
        txtOUT_DIAGNOSIS_DATE3.Enabled = False
        cmbTREAT_RESULT3.ListIndex = -1
        cmbTREAT_RESULT3.Enabled = False
    End If
End Sub

Private Sub txtOUT_DIAGNOSIS_NAME3_Change()
    Call SetNotText(labNot(22), txtOUT_DIAGNOSIS_NAME3)
End Sub

Private Sub txtOUT_DIAGNOSIS_DATE3_Change()
    labNot(23).Caption = IIf(IsDate(txtOUT_DIAGNOSIS_DATE3.Text), "√", "*")
    labNot(23).ForeColor = IIf(IsDate(txtOUT_DIAGNOSIS_DATE3.Text), vbGreen, vbRed)
End Sub

Private Sub cmbTREAT_RESULT3_Click()
    Call SetNotComb(labNot(24), cmbTREAT_RESULT3)
End Sub

Private Sub txtOPERATION_CODE1_Change()
    If Len(Trim(txtOPERATION_CODE1.Text)) > 0 Then
        labNot(25).Visible = True
        labNot(26).Visible = True
        labNot(27).Visible = True
        labNot(28).Visible = True
        labNot(29).Visible = True
        txtOPERATION_NAME1.Enabled = True
        cmbWOUND_GRADE1.Enabled = True
        cmbHEAL1.Enabled = True
        txtOPERATING_DATE1.Enabled = True
        cmbANAESTHESIA_METHOD1.Enabled = True
    Else
        labNot(25).Visible = False
        labNot(26).Visible = False
        labNot(27).Visible = False
        labNot(28).Visible = False
        labNot(29).Visible = False
        txtOPERATION_NAME1.Text = ""
        cmbWOUND_GRADE1.ListIndex = -1
        cmbHEAL1.ListIndex = -1
        txtOPERATING_DATE1.Text = ""
        cmbANAESTHESIA_METHOD1.ListIndex = -1
        txtOPERATION_NAME1.Enabled = False
        cmbWOUND_GRADE1.Enabled = False
        cmbHEAL1.Enabled = False
        txtOPERATING_DATE1.Enabled = False
        cmbANAESTHESIA_METHOD1.Enabled = False
    End If
End Sub

Private Sub txtOPERATION_NAME1_Change()
    Call SetNotText(labNot(25), txtOPERATION_NAME1)
End Sub

Private Sub cmbWOUND_GRADE1_Click()
    Call SetNotComb(labNot(26), cmbWOUND_GRADE1)
End Sub

Private Sub cmbHEAL1_Click()
    Call SetNotComb(labNot(27), cmbHEAL1)
End Sub

Private Sub txtOPERATING_DATE1_Change()
    labNot(28).Caption = IIf(IsDate(txtOPERATING_DATE1.Text), "√", "*")
    labNot(28).ForeColor = IIf(IsDate(txtOPERATING_DATE1.Text), vbGreen, vbRed)
End Sub

Private Sub cmbANAESTHESIA_METHOD1_Click()
    Call SetNotComb(labNot(29), cmbANAESTHESIA_METHOD1)
End Sub

Private Sub txtOPERATION_CODE2_Change()
    If Len(Trim(txtOPERATION_CODE2.Text)) > 0 Then
        labNot(30).Visible = True
        labNot(31).Visible = True
        labNot(32).Visible = True
        labNot(33).Visible = True
        labNot(34).Visible = True
        txtOPERATION_NAME2.Enabled = True
        cmbWOUND_GRADE2.Enabled = True
        cmbHEAL2.Enabled = True
        txtOPERATING_DATE2.Enabled = True
        cmbANAESTHESIA_METHOD2.Enabled = True
    Else
        labNot(30).Visible = False
        labNot(31).Visible = False
        labNot(32).Visible = False
        labNot(33).Visible = False
        labNot(34).Visible = False
        txtOPERATION_NAME2.Text = ""
        cmbWOUND_GRADE2.ListIndex = -1
        cmbHEAL2.ListIndex = -1
        txtOPERATING_DATE2.Text = ""
        cmbANAESTHESIA_METHOD2.ListIndex = -1
        txtOPERATION_NAME2.Enabled = False
        cmbWOUND_GRADE2.Enabled = False
        cmbHEAL2.Enabled = False
        txtOPERATING_DATE2.Enabled = False
        cmbANAESTHESIA_METHOD2.Enabled = False
    End If
End Sub

Private Sub txtOPERATION_NAME2_Change()
    Call SetNotText(labNot(30), txtOPERATION_NAME2)
End Sub

Private Sub cmbWOUND_GRADE2_Click()
    Call SetNotComb(labNot(31), cmbWOUND_GRADE2)
End Sub

Private Sub cmbHEAL2_Click()
    Call SetNotComb(labNot(32), cmbHEAL2)
End Sub

Private Sub txtOPERATING_DATE2_Change()
    labNot(33).Caption = IIf(IsDate(txtOPERATING_DATE2.Text), "√", "*")
    labNot(33).ForeColor = IIf(IsDate(txtOPERATING_DATE2.Text), vbGreen, vbRed)
End Sub

Private Sub cmbANAESTHESIA_METHOD2_Click()
    Call SetNotComb(labNot(34), cmbANAESTHESIA_METHOD2)
End Sub

Private Sub txtOPERATION_CODE3_Change()
    If Len(Trim(txtOPERATION_CODE3.Text)) > 0 Then
        labNot(35).Visible = True
        labNot(36).Visible = True
        labNot(37).Visible = True
        labNot(38).Visible = True
        labNot(39).Visible = True
        txtOPERATION_NAME3.Enabled = True
        cmbWOUND_GRADE3.Enabled = True
        cmbHEAL3.Enabled = True
        txtOPERATING_DATE3.Enabled = True
        cmbANAESTHESIA_METHOD3.Enabled = True
    Else
        labNot(35).Visible = False
        labNot(36).Visible = False
        labNot(37).Visible = False
        labNot(38).Visible = False
        labNot(39).Visible = False
        txtOPERATION_NAME3.Text = ""
        cmbWOUND_GRADE3.ListIndex = -1
        cmbHEAL3.ListIndex = -1
        txtOPERATING_DATE3.Text = ""
        cmbANAESTHESIA_METHOD3.ListIndex = -1
        txtOPERATION_NAME3.Enabled = False
        cmbWOUND_GRADE3.Enabled = False
        cmbHEAL3.Enabled = False
        txtOPERATING_DATE3.Enabled = False
        cmbANAESTHESIA_METHOD3.Enabled = False
    End If
End Sub

Private Sub txtOPERATION_NAME3_Change()
    Call SetNotText(labNot(35), txtOPERATION_NAME3)
End Sub

Private Sub cmbWOUND_GRADE3_Click()
    Call SetNotComb(labNot(36), cmbWOUND_GRADE3)
End Sub

Private Sub cmbHEAL3_Click()
    Call SetNotComb(labNot(37), cmbHEAL3)
End Sub

Private Sub txtOPERATING_DATE3_Change()
    labNot(38).Caption = IIf(IsDate(txtOPERATING_DATE3.Text), "√", "*")
    labNot(38).ForeColor = IIf(IsDate(txtOPERATING_DATE3.Text), vbGreen, vbRed)
End Sub

Private Sub cmbANAESTHESIA_METHOD3_Click()
    Call SetNotComb(labNot(39), cmbANAESTHESIA_METHOD2)
End Sub

Private Function TRANDATA(ByVal STR信息名 As String, ByVal STR信息值 As String) As String
    '根据接口文档转换HIS中的值
    Select Case STR信息名
    Case "医疗付款方式"
        Select Case STR信息值
            Case "社会基本医疗保险"
                TRANDATA = 1
            Case "商业保险"
                TRANDATA = 2
            Case "自费医疗"
                TRANDATA = 3
            Case "公费医疗"
                TRANDATA = 4
            Case "大病统筹"
                TRANDATA = 5
            Case Else   '其他
                TRANDATA = 6
        End Select
    Case "性别"
        Select Case STR信息值
            Case "男"
                TRANDATA = 1
            Case Else   '女
                TRANDATA = 2
        End Select
    Case "婚姻"
        Select Case STR信息值
            Case "未婚"
                TRANDATA = 1
            Case "已婚"
                TRANDATA = 2
            Case "离婚"
                TRANDATA = 3
            Case Else   '丧
                TRANDATA = 4
        End Select
    Case "与病人关系"
        Select Case STR信息值
            Case "配偶"
                TRANDATA = 1
            Case "子", "女"
                TRANDATA = 2
            Case "父母"
                TRANDATA = 3
            Case Else   '孙子\孙女\祖父\祖母\本人等等,都归入其他
                TRANDATA = 9
        End Select
    Case "尸检标志", "首例", "随诊标志", "示教病例", "RH", "输血反应"
        Select Case STR信息值
            Case "是"
                TRANDATA = 1
            Case Else
                TRANDATA = 2
        End Select
    Case "血型"
        Select Case STR信息值
            Case "A"
                TRANDATA = 1
            Case "B"
                TRANDATA = 2
            Case "AB"
                TRANDATA = 3
            Case "O"
                TRANDATA = 4
            Case Else
                TRANDATA = 5
        End Select
    Case "麻醉类型"
        Select Case STR信息值
            Case "全麻"
                TRANDATA = 1
            Case "局麻"
                TRANDATA = 3
            Case Else
                TRANDATA = 2
        End Select
    Case "病案质量"
        Select Case STR信息值
            Case "甲"
                TRANDATA = 1
            Case "乙"
                TRANDATA = 2
            Case Else
                TRANDATA = 3
        End Select
    Case "诊疗结果", "出院情况"
        Select Case STR信息值
            Case "治愈", "正常"
                TRANDATA = 1
            Case "好转"
                TRANDATA = 2
            Case "未愈"
                TRANDATA = 3
            Case "死亡"
                TRANDATA = 4
            Case Else
                TRANDATA = 5
        End Select
    Case "HBSAG", "HCV_AB", "HIV_AB"
        Select Case STR信息值
            Case "阴性"
                TRANDATA = 1
            Case "阳性"
                TRANDATA = 2
            Case Else
                TRANDATA = 0
        End Select
    Case "入院病情"
        Select Case STR信息值
            Case "危"
                TRANDATA = 1
            Case "急"
                TRANDATA = 2
            Case Else
                TRANDATA = 3
        End Select
    Case "切口"
        Select Case STR信息值
            Case "Ⅰ"
                TRANDATA = "1"
            Case "Ⅱ"
                TRANDATA = "2"
            Case "Ⅲ"
                TRANDATA = "3"
        End Select
    Case "愈合"     '接口内是统一判断的
        Select Case STR信息值
            Case "甲"
                TRANDATA = "1"
            Case "乙"
                TRANDATA = "2"
            Case "丙"
                TRANDATA = "3"
        End Select
    Case "诊断类型"
        Select Case STR信息值
            Case 5, 6, 7
                TRANDATA = Val(STR信息值) - 1
            Case 1, 2, 3
                TRANDATA = Val(STR信息值)
        End Select
    End Select
End Function

Private Sub cmdCancel_Click()
On Error GoTo ErrH
    mstrHospitalNumber = ""
    Unload Me
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrH
    If Not chkData Then Exit Sub
    '保存数据
    gstrSQL = ""
    gstrSQL = gstrSQL & vbCrLf & "zl_长治病案信息_Update("
    '基本信息
    gstrSQL = gstrSQL & vbCrLf & txtSickID2.Text & "," & "'" & txtCnName1.Text & "','" & txtSex1.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtHOSPITAL_NUMBER.Text & "','" & txtRESIDENCE_NO.Text & "'," & txtIN_COUNT.Text & ","
    gstrSQL = gstrSQL & vbCrLf & "'" & txtMEDICAL_RECORD_NO.Text & "','" & Cmb_ID(cmbMARITAL_STATUS) & "','" & txtSTATUS.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtBIRTH_ADDRESS.Text & "','" & txtIDENTITY_NUMBER.Text & "','" & txtUNIT_NAME.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtUNIT_ADDRESS.Text & "','" & txtUNIT_PHONE.Text & "','" & txtUNIT_ZIPCODE.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtREGISTER_ADDRESS.Text & "','" & txtREGISTER_ZIPCODE.Text & "','" & txtCONTACT_PERSON.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbRELATIONSHIP) & "','" & txtCONTACT_ADDRESS.Text & "','" & txtCONTACT_PHONE.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "to_date('" & txtADMISSION_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & txtADMISSION_DEPT.Text & "','" & txtIN_DEPT_ZONE.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtDEPT_TRANSFERED_TO.Text & "',to_date('" & txtDISCHARGE_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & txtDISCHARGE_DEPT.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOUT_DEPT_ZONE.Text & "','" & Cmb_ID(cmbPAT_ADM_CONDITION) & "',to_date('" & txtDIAGNOSIS_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),"
    '医师信息
    gstrSQL = gstrSQL & vbCrLf & "'" & txtALERGY_DRUGS.Text & "','" & Cmb_ID(cmbHBSAG) & "','" & Cmb_ID(cmbHCV_AB) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbHIV_AB) & "','" & Cmb_ID(cmdCLINIC_INHOSPITAL) & "','" & Cmb_ID(cmbIN_OUT) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbBEFORE_AFTER_TREATMENT) & "','" & Cmb_ID(cmbCLINIC_PATHOLOGY) & "','" & Cmb_ID(cmbEMIT_PATHOLOGY) & "',"
    
    gstrSQL = gstrSQL & vbCrLf & "" & Val(txtEMER_TREAT_TIMES.Text) & "," & Val(txtESC_EMER_TIMES.Text) & ",'" & txtDIRECTOR.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtDIRECTOR_DOCTOR.Text & "','" & txtATTENDING_DOCTOR.Text & "','" & txtINHOSPITAL_DOCTOR.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtREFRESH_DOCTOR.Text & "','" & txtGRADUATE_DOCTOR.Text & "','" & txtINTERM.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtCODE_NAME.Text & "','" & Cmb_ID(cmbMEDICAL_RECORD_MASS) & "','" & txtCONTROL_DOCTOR.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtCONTROL_NURSE.Text & "',to_date('" & txtBAL_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbBODY_EXAMINE_FLAG) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbFIRST_FLAG) & "','" & Cmb_ID(cmbFOLLOW_FLAG) & "','" & txtFOLLOW_TERM.Text & "',"
    '诊断信息
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbTEACH_MR_FLAG) & "','" & Cmb_ID(cmbBLOOD_TYPE) & "','" & Cmb_ID(cmbRH) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbBLOOD_TRAN_REACT_FLAG) & "'," & Val(txtERYTHROCYTE) & "," & Val(txtHEMOBLAST) & ","
    gstrSQL = gstrSQL & vbCrLf & "" & Val(txtPLASM.Text) & "," & Val(txtBLOOD.Text) & ",'" & Val(txtOTHER_BLOOD.Text) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtHANDLE.Text & "',to_date('" & txtHANDLE_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & txtIN_DIAGNOSIS_CODE.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtIN_DIAGNOSIS_NAME.Text & "',to_date('" & txtIN_DIAGNOSIS_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & txtOUT_DIAGNOSIS_CODE1.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOUT_DIAGNOSIS_NAME1.Text & "',to_date('" & txtOUT_DIAGNOSIS_DATE1.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbTREAT_RESULT1) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOUT_DIAGNOSIS_CODE2.Text & "','" & txtOUT_DIAGNOSIS_NAME2.Text & "',to_date('" & txtOUT_DIAGNOSIS_DATE2.Text & "','yyyy-mm-dd hh24:mi:ss'),"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbTREAT_RESULT2) & "','" & txtOUT_DIAGNOSIS_CODE3.Text & "','" & txtOUT_DIAGNOSIS_NAME3.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "to_date('" & txtOUT_DIAGNOSIS_DATE3.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbTREAT_RESULT3) & "',"
    '手术信息
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOPERATION_CODE1.Text & "','" & txtOPERATION_NAME1.Text & "','" & Cmb_ID(cmbWOUND_GRADE1) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbHEAL1) & "',to_date('" & txtOPERATING_DATE1.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbANAESTHESIA_METHOD1) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOPERATION_CODE2.Text & "','" & txtOPERATION_NAME2.Text & "','" & Cmb_ID(cmbWOUND_GRADE2) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbHEAL2) & "',to_date('" & txtOPERATING_DATE2.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbANAESTHESIA_METHOD2) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOPERATION_CODE3.Text & "','" & txtOPERATION_NAME3.Text & "','" & Cmb_ID(cmbWOUND_GRADE3) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbHEAL3) & "',to_date('" & txtOPERATING_DATE3.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbANAESTHESIA_METHOD3) & "'"
    gstrSQL = gstrSQL & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Unload Me
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'检测必须录入数据
Private Function chkData() As Boolean
    Dim strMsg      As String
    Dim blnTbc      As Boolean
On Error GoTo ErrH
    strMsg = ""
    blnTbc = False
    If Len(Trim(txtHOSPITAL_NUMBER.Text)) = 0 Then
        strMsg = strMsg & "【基础信息—" & labHOSPITAL_NUMBER.Caption & "】不能为空！" & vbCrLf
    End If
    If Len(Trim(txtRESIDENCE_NO.Text)) = 0 Then
        strMsg = strMsg & "【基础信息—" & labRESIDENCE_NO.Caption & "】不能为空！" & vbCrLf
    End If
    If Val(Trim(txtIN_COUNT.Text)) <= 0 Then
        strMsg = strMsg & "【基础信息—" & labIN_COUNT.Caption & "】不能为空！且大于零！" & vbCrLf
    End If
    If Len(Trim(txtMEDICAL_RECORD_NO.Text)) = 0 Then
        strMsg = strMsg & "【基础信息—" & labMEDICAL_RECORD_NO.Caption & "】不能为空！" & vbCrLf
    End If
    If Not IsDate(Trim(txtADMISSION_DATE.Text)) Then
        strMsg = strMsg & "【基础信息—" & labADMISSION_DATE.Caption & "】不能为空！为日期类型！" & vbCrLf
    End If
    If Len(Trim(txtADMISSION_DEPT.Text)) = 0 Then
        strMsg = strMsg & "【基础信息—" & labADMISSION_DEPT.Caption & "】不能为空！" & vbCrLf
    End If
    If Len(Trim(txtIN_DEPT_ZONE.Text)) = 0 Then
        strMsg = strMsg & "【基础信息—" & labIN_DEPT_ZONE.Caption & "】不能为空！" & vbCrLf
    End If
    If Not IsDate(Trim(txtDISCHARGE_DATE.Text)) Then
        strMsg = strMsg & "【基础信息—" & labDISCHARGE_DATE.Caption & "】不能为空！为日期类型！" & vbCrLf
    End If
    If strMsg <> "" Then
        tbcPage.Item(0).Selected = True
        blnTbc = True
    End If
    '医师信息
    If Len(Trim(cmbIN_OUT.Text)) = 0 Then
        strMsg = strMsg & "【医师信息—" & labIN_OUT.Caption & "】不能为空！请选择！" & vbCrLf
    End If
    If strMsg <> "" And Not blnTbc Then
        tbcPage.Item(1).Selected = True
        blnTbc = True
    End If
    '诊断信息
    If Len(Trim(txtHANDLE.Text)) = 0 Then
        strMsg = strMsg & "【诊断信息—" & labHANDLE.Caption & "】不能为空！" & vbCrLf
    End If
    If Not IsDate(Trim(txtHANDLE_DATE.Text)) Then
        strMsg = strMsg & "【诊断信息—" & labHANDLE_DATE.Caption & "】不能为空！为日期类型！" & vbCrLf
    End If
    If Len(Trim(txtIN_DIAGNOSIS_CODE.Text)) = 0 Then
        strMsg = strMsg & "【诊断信息—" & labIN_DIAGNOSIS_CODE.Caption & "】不能为空！" & vbCrLf
    End If
    If Len(Trim(txtIN_DIAGNOSIS_NAME.Text)) = 0 Then
        strMsg = strMsg & "【诊断信息—" & labIN_DIAGNOSIS_NAME.Caption & "】不能为空！" & vbCrLf
    End If
    If Not IsDate(Trim(txtIN_DIAGNOSIS_DATE.Text)) Then
        strMsg = strMsg & "【诊断信息—" & labIN_DIAGNOSIS_DATE.Caption & "】不能为空！为日期类型！" & vbCrLf
    End If
    If Len(Trim(txtOUT_DIAGNOSIS_CODE1.Text)) = 0 Then
        strMsg = strMsg & "【诊断信息—" & labOUT_DIAGNOSIS_CODE1.Caption & "】不能为空！" & vbCrLf
    End If
    If Len(Trim(txtOUT_DIAGNOSIS_NAME1.Text)) = 0 Then
        strMsg = strMsg & "【诊断信息—" & labOUT_DIAGNOSIS_NAME1.Caption & "】不能为空！" & vbCrLf
    End If
    If Not IsDate(Trim(txtOUT_DIAGNOSIS_DATE1.Text)) Then
        strMsg = strMsg & "【诊断信息—" & labOUT_DIAGNOSIS_DATE1.Caption & "】不能为空！为日期类型！" & vbCrLf
    End If
    If Len(Trim(cmbTREAT_RESULT1.Text)) = 0 Then
        strMsg = strMsg & "【诊断信息—" & labTREAT_RESULT1.Caption & "】不能为空！请选择！" & vbCrLf
    End If
    '如果诊断2或诊断3不为空，那么所有诊断2的数据都必须填写
    If Len(Trim(txtOUT_DIAGNOSIS_CODE2.Text & txtOUT_DIAGNOSIS_CODE3.Text)) <> 0 Then
        '检测诊断2的信息
        
    End If
    '如果诊断3的数据不为空，那么必须填写完整。
    
    If strMsg <> "" And Not blnTbc Then
        tbcPage.Item(2).Selected = True
        blnTbc = True
    End If
    
    If strMsg <> "" And Not blnTbc Then
        tbcPage.Item(3).Selected = True
        blnTbc = True
    End If
    If strMsg = "" Then
        chkData = True
    Else
        MsgBox strMsg, vbCritical, gstrSysName
    End If
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Function
End Function
