VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMain_�������󲡰���Ϣ�༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���β�����Ϣ�༭"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14685
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   Icon            =   "frmMain_�������󲡰���Ϣ�༭.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   13380
      TabIndex        =   217
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "��ȡ(&R)"
      Height          =   350
      Left            =   11040
      TabIndex        =   216
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   12225
      TabIndex        =   215
      Top             =   5430
      Width           =   1100
   End
   Begin VB.PictureBox pic�����Ϣ 
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "���ƽ��3"
         Height          =   195
         Left            =   5160
         TabIndex        =   156
         Top             =   4155
         Width           =   885
      End
      Begin VB.Label labOUT_DIAGNOSIS_DATE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������3"
         Height          =   195
         Left            =   90
         TabIndex        =   155
         Top             =   4155
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_NAME3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������3"
         Height          =   195
         Left            =   9630
         TabIndex        =   154
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_CODE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��ϴ���3"
         Height          =   195
         Left            =   4770
         TabIndex        =   153
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label labTREAT_RESULT2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ƽ��2"
         Height          =   195
         Left            =   480
         TabIndex        =   152
         Top             =   3660
         Width           =   885
      End
      Begin VB.Label labOUT_DIAGNOSIS_DATE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������2"
         Height          =   195
         Left            =   9630
         TabIndex        =   151
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_NAME2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������2"
         Height          =   195
         Left            =   4770
         TabIndex        =   150
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_CODE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��ϴ���2"
         Height          =   195
         Left            =   90
         TabIndex        =   149
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label labTREAT_RESULT1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ƽ��1"
         Height          =   195
         Left            =   10020
         TabIndex        =   148
         Top             =   2775
         Width           =   885
      End
      Begin VB.Label labOUT_DIAGNOSIS_DATE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������1"
         Height          =   195
         Left            =   4770
         TabIndex        =   147
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_NAME1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������1"
         Height          =   195
         Left            =   90
         TabIndex        =   146
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label labOUT_DIAGNOSIS_CODE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��ϴ���1"
         Height          =   195
         Left            =   9615
         TabIndex        =   145
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label labIN_DIAGNOSIS_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   195
         Left            =   5250
         TabIndex        =   144
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label labIN_DIAGNOSIS_NAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ�������"
         Height          =   195
         Left            =   180
         TabIndex        =   143
         Top             =   2340
         Width           =   1170
      End
      Begin VB.Label labIN_DIAGNOSIS_CODE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��ϴ���"
         Height          =   195
         Left            =   9735
         TabIndex        =   142
         Top             =   1905
         Width           =   1170
      End
      Begin VB.Label labHANDLE_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   195
         Left            =   5265
         TabIndex        =   141
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label labHANDLE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   195
         Left            =   780
         TabIndex        =   140
         Top             =   1905
         Width           =   585
      End
      Begin VB.Label labOTHER_BLOOD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   195
         Left            =   10515
         TabIndex        =   139
         Top             =   1470
         Width           =   390
      End
      Begin VB.Label labBLOOD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ȫѪ"
         Height          =   195
         Left            =   5655
         TabIndex        =   138
         Top             =   1470
         Width           =   390
      End
      Begin VB.Label labPLASM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����Ѫ��"
         Height          =   195
         Left            =   585
         TabIndex        =   137
         Top             =   1470
         Width           =   780
      End
      Begin VB.Label labHEMOBLAST 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ѪС��"
         Height          =   195
         Left            =   9930
         TabIndex        =   136
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label labERYTHROCYTE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����ϸ��"
         Height          =   195
         Left            =   5070
         TabIndex        =   135
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label labBLOOD_TRAN_REACT_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����Ѫ��Ӧ��־"
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
         Caption         =   "Ѫ��"
         Height          =   195
         Left            =   5655
         TabIndex        =   132
         Top             =   525
         Width           =   390
      End
      Begin VB.Label labTEACH_MR_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ʾ�̲���"
         Height          =   195
         Left            =   585
         TabIndex        =   131
         Top             =   570
         Width           =   780
      End
      Begin VB.Label labSick1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ID"
         Height          =   195
         Left            =   765
         TabIndex        =   124
         Top             =   195
         Width           =   600
      End
      Begin VB.Label labCnName3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   195
         Left            =   5655
         TabIndex        =   123
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSex3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   195
         Left            =   10515
         TabIndex        =   122
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.PictureBox pic������Ϣ 
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   232
      ScaleHeight     =   4425
      ScaleWidth      =   14220
      TabIndex        =   0
      Top             =   315
      Width           =   14220
      Begin VB.CommandButton cmdSick 
         Caption         =   "��"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "ҽ�ƻ�������"
         Height          =   195
         Left            =   210
         TabIndex        =   56
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label labRESIDENCE_NO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   195
         Left            =   5460
         TabIndex        =   55
         Top             =   585
         Width           =   585
      End
      Begin VB.Label labIN_COUNT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����סԺ����"
         Height          =   195
         Left            =   9735
         TabIndex        =   54
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label labMEDICAL_RECORD_NO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   195
         Left            =   795
         TabIndex        =   53
         Top             =   1020
         Width           =   585
      End
      Begin VB.Label labMARITAL_STATUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����״��"
         Height          =   195
         Left            =   5265
         TabIndex        =   52
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label labSTATUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ְҵ"
         Height          =   195
         Left            =   10515
         TabIndex        =   51
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label labBIRTH_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   195
         Left            =   795
         TabIndex        =   50
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label labIDENTITY_NUMBER 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���֤��"
         Height          =   195
         Left            =   5265
         TabIndex        =   49
         Top             =   1485
         Width           =   780
      End
      Begin VB.Label labUNIT_NAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������λ"
         Height          =   195
         Left            =   10125
         TabIndex        =   48
         Top             =   1485
         Width           =   780
      End
      Begin VB.Label labUNIT_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��λ��ַ"
         Height          =   195
         Left            =   600
         TabIndex        =   47
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labUNIT_PHONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��λ�绰"
         Height          =   195
         Left            =   5265
         TabIndex        =   46
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labUNIT_ZIPCODE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��λ�ʱ�"
         Height          =   195
         Left            =   10125
         TabIndex        =   45
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labREGISTER_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ڵ�ַ"
         Height          =   195
         Left            =   600
         TabIndex        =   44
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labREGISTER_ZIPCODE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����ʱ�"
         Height          =   195
         Left            =   5265
         TabIndex        =   43
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labCONTACT_PERSON 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ϵ��"
         Height          =   195
         Left            =   10320
         TabIndex        =   42
         Top             =   2370
         Width           =   585
      End
      Begin VB.Label labRELATIONSHIP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�벡�˹�ϵ"
         Height          =   195
         Left            =   405
         TabIndex        =   41
         Top             =   2805
         Width           =   975
      End
      Begin VB.Label labCONTACT_ADDRESS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ϵ��ַ"
         Height          =   195
         Left            =   5265
         TabIndex        =   40
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labCONTACT_PHONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ϵ�绰"
         Height          =   195
         Left            =   10125
         TabIndex        =   39
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labADMISSION_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   600
         TabIndex        =   38
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labADMISSION_DEPT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   5265
         TabIndex        =   37
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labIN_DEPT_ZONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   10125
         TabIndex        =   36
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labDEPT_TRANSFERED_TO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ת�ƿƱ�"
         Height          =   195
         Left            =   600
         TabIndex        =   35
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labDISCHARGE_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   5265
         TabIndex        =   34
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labDISCHARGE_DEPT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   10125
         TabIndex        =   33
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labOUT_DEPT_ZONE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   600
         TabIndex        =   32
         Top             =   4140
         Width           =   780
      End
      Begin VB.Label labPAT_ADM_CONDITION 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����"
         Height          =   195
         Left            =   5265
         TabIndex        =   31
         Top             =   4140
         Width           =   780
      End
      Begin VB.Label labDIAGNOSIS_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��ȷ������"
         Height          =   195
         Left            =   9540
         TabIndex        =   30
         Top             =   4140
         Width           =   1365
      End
      Begin VB.Label labSickID1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ID"
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   195
         Width           =   600
      End
      Begin VB.Label labCnName1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   195
         Left            =   5655
         TabIndex        =   5
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSex1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   195
         Left            =   10515
         TabIndex        =   4
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.PictureBox picҽʦ��Ϣ 
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
         Caption         =   "����ID"
         Height          =   195
         Left            =   765
         TabIndex        =   108
         Top             =   195
         Width           =   600
      End
      Begin VB.Label labCnName2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   195
         Left            =   5655
         TabIndex        =   107
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSex2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   195
         Left            =   10515
         TabIndex        =   106
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labFOLLOW_TERM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   195
         Left            =   10125
         TabIndex        =   102
         Top             =   4155
         Width           =   780
      End
      Begin VB.Label labFOLLOW_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����־"
         Height          =   195
         Left            =   5265
         TabIndex        =   101
         Top             =   4155
         Width           =   780
      End
      Begin VB.Label labFIRST_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ��һ��"
         Height          =   195
         Left            =   390
         TabIndex        =   100
         Top             =   4155
         Width           =   975
      End
      Begin VB.Label labBODY_EXAMINE_FLAG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ʬ���־"
         Height          =   195
         Left            =   10125
         TabIndex        =   99
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labBAL_DATE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   195
         Left            =   5265
         TabIndex        =   98
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labCONTROL_NURSE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ʿػ�ʦ"
         Height          =   195
         Left            =   585
         TabIndex        =   97
         Top             =   3735
         Width           =   780
      End
      Begin VB.Label labCONTROL_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ʿ�ҽʦ"
         Height          =   195
         Left            =   10125
         TabIndex        =   96
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labMEDICAL_RECORD_MASS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   195
         Left            =   5265
         TabIndex        =   95
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label labCODE_NAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����Ա"
         Height          =   195
         Left            =   780
         TabIndex        =   94
         Top             =   3285
         Width           =   585
      End
      Begin VB.Label labINTERM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ʵϰҽʦ"
         Height          =   195
         Left            =   10125
         TabIndex        =   93
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labGRADUATE_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�о���ʵϰҽʦ"
         Height          =   195
         Left            =   4680
         TabIndex        =   92
         Top             =   2820
         Width           =   1365
      End
      Begin VB.Label labREFRESH_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ҽʦ"
         Height          =   195
         Left            =   585
         TabIndex        =   91
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label labINHOSPITAL_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "סԺҽʦ"
         Height          =   195
         Left            =   10125
         TabIndex        =   90
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labATTENDING_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ҽʦ"
         Height          =   195
         Left            =   5265
         TabIndex        =   89
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labDIRECTOR_DOCTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ҽʦ"
         Height          =   195
         Left            =   585
         TabIndex        =   88
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label labDIRECTOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   195
         Left            =   10320
         TabIndex        =   87
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label labESC_EMER_TIMES 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ȳɹ�����"
         Height          =   195
         Left            =   4875
         TabIndex        =   86
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label labEMER_TREAT_TIMES 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ȴ���"
         Height          =   195
         Left            =   585
         TabIndex        =   85
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label labEMIT_PATHOLOGY 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����벡��"
         Height          =   195
         Left            =   9930
         TabIndex        =   84
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label labCLINIC_PATHOLOGY 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ٴ��벡��"
         Height          =   195
         Left            =   5070
         TabIndex        =   83
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label labBEFORE_AFTER_TREATMENT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ǰ������"
         Height          =   195
         Left            =   390
         TabIndex        =   82
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label labIN_OUT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ���Ժ"
         Height          =   195
         Left            =   9930
         TabIndex        =   81
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label labCLINIC_INHOSPITAL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�������Ժ"
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
         Caption         =   "����ҩ��"
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
            Name            =   "����"
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
   Begin VB.PictureBox pic������Ϣ 
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "������3"
         Height          =   195
         Left            =   10020
         TabIndex        =   212
         Top             =   2715
         Width           =   885
      End
      Begin VB.Label labOPERATING_DATE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������3"
         Height          =   195
         Left            =   5130
         TabIndex        =   211
         Top             =   2715
         Width           =   885
      End
      Begin VB.Label labHEAL3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�п��������3"
         Height          =   195
         Left            =   60
         TabIndex        =   210
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label labWOUND_GRADE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�пڵȼ�3"
         Height          =   195
         Left            =   10020
         TabIndex        =   209
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label labOPERATION_NAME3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������3"
         Height          =   195
         Left            =   5130
         TabIndex        =   208
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label labOPERATION_CODE3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������3"
         Height          =   195
         Left            =   450
         TabIndex        =   207
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label labANAESTHESIA_METHOD2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������2"
         Height          =   195
         Left            =   10020
         TabIndex        =   206
         Top             =   1875
         Width           =   885
      End
      Begin VB.Label labOPERATING_DATE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������2"
         Height          =   195
         Left            =   5130
         TabIndex        =   205
         Top             =   1875
         Width           =   885
      End
      Begin VB.Label labHEAL2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�п��������2"
         Height          =   195
         Left            =   60
         TabIndex        =   204
         Top             =   1875
         Width           =   1275
      End
      Begin VB.Label labWOUND_GRADE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�пڵȼ�2"
         Height          =   195
         Left            =   10020
         TabIndex        =   203
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labOPERATION_NAME2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������2"
         Height          =   195
         Left            =   5130
         TabIndex        =   202
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labOPERATION_CODE2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������2"
         Height          =   195
         Left            =   450
         TabIndex        =   201
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labANAESTHESIA_METHOD1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������1"
         Height          =   195
         Left            =   10020
         TabIndex        =   200
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label labOPERATING_DATE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������1"
         Height          =   195
         Left            =   5130
         TabIndex        =   199
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label labHEAL1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�п��������1"
         Height          =   195
         Left            =   60
         TabIndex        =   198
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label labWOUND_GRADE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�пڵȼ�1"
         Height          =   195
         Left            =   10020
         TabIndex        =   197
         Top             =   600
         Width           =   885
      End
      Begin VB.Label labOPERATION_NAME1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������1"
         Height          =   195
         Left            =   5130
         TabIndex        =   196
         Top             =   600
         Width           =   885
      End
      Begin VB.Label labOPERATION_CODE1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������1"
         Height          =   195
         Left            =   450
         TabIndex        =   195
         Top             =   600
         Width           =   885
      End
      Begin VB.Label labSex4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   195
         Left            =   10515
         TabIndex        =   185
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labCNNAME4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   195
         Left            =   5625
         TabIndex        =   184
         Top             =   195
         Width           =   390
      End
      Begin VB.Label labSTICKID4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ID"
         Height          =   195
         Left            =   735
         TabIndex        =   183
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Frame fra�����༭ 
      Caption         =   "�����༭"
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
Attribute VB_Name = "frmMain_�������󲡰���Ϣ�༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
'ֻ��ʾ30���ڳ�Ժ��δ�ϴ����˵���Ϣ��
Const strSickFields = "select B.����ID as ID,B.����ID as ����ID,B.סԺ���� As ��ҳID,a.ҽ����, a.����,a.��Ա��� as ��Ա���,b.����,b.�Ա�,b.���֤�� " & vbNewLine & _
                      "from �����ʻ� a , ������Ϣ b,������ҳ c where a.����ID = b.����id And b.����ID = c.����id and b.סԺ���� = c.��ҳid And a.���� = [1]" & vbNewLine & _
                      "AND C.��Ժ���� >= sysdate-120" & vbNewLine & _
                      "And not exists (select 1 from ���β�����Ϣ d where a.����ID= d.STICKID And B.סԺ����=d.In_Count)"

Private mstrHospitalNumber      As String
Private mlng����ID              As Long
Private mlng��ҳID              As Long
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
'=���ܣ� ��ʼTab�ؼ�
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
        .InsertItem 0, " ������Ϣ ", pic������Ϣ.hwnd, 0
        .InsertItem 1, " ҽʦ��Ϣ ", picҽʦ��Ϣ.hwnd, 0
        .InsertItem 2, " �����Ϣ ", pic�����Ϣ.hwnd, 0
        .InsertItem 3, " ������Ϣ ", pic������Ϣ.hwnd, 0
        .Item(0).Selected = True
    End With
    InitTabControl = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� ��ʼTab�ؼ�
'==============================================================================
Private Sub InitCmbControl()
    On Error GoTo ErrH
    '����״��
    With cmbMARITAL_STATUS
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "δ��"
        .AddItem "2" & gstrSplitCmb & "�ѻ�"
        .AddItem "3" & gstrSplitCmb & "���"
        .AddItem "4" & gstrSplitCmb & "ɥż"
    End With
    '�벡�˹�ϵ
    With cmbRELATIONSHIP
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��ż"
        .AddItem "2" & gstrSplitCmb & "��Ů"
        .AddItem "3" & gstrSplitCmb & "��ĸ"
        .AddItem "9" & gstrSplitCmb & "����"
    End With
    '��Ժ����
    With cmbPAT_ADM_CONDITION
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "Σ"
        .AddItem "2" & gstrSplitCmb & "��"
        .AddItem "3" & gstrSplitCmb & "һ��"
        .AddItem "4" & gstrSplitCmb & "����"
    End With
    'HBSAG
    With cmbHBSAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    'HCV_AB
    With cmbHCV_AB
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    'HIV_AB
    With cmbHIV_AB
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    '�������Ժ
    With cmdCLINIC_INHOSPITAL
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    '��Ժ���Ժ
    With cmbIN_OUT
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    '��ǰ������
    With cmbBEFORE_AFTER_TREATMENT
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    '�ٴ��벡��
    With cmbCLINIC_PATHOLOGY
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    '�����벡��
    With cmbEMIT_PATHOLOGY
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "����"
        .AddItem "3" & gstrSplitCmb & "δ��"
    End With
    '��������
    With cmbMEDICAL_RECORD_MASS
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
        .AddItem "3" & gstrSplitCmb & "��"
    End With
    '��Ժ��һ��
    With cmbFIRST_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
    End With
    'ʬ���־
    With cmbBODY_EXAMINE_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
    End With
    '�����־
    With cmbFOLLOW_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
    End With
    'ʾ�̲���
    With cmbTEACH_MR_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
    End With
    'Ѫ�ͱ�־
    With cmbBLOOD_TYPE
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "A"
        .AddItem "2" & gstrSplitCmb & "B"
        .AddItem "3" & gstrSplitCmb & "AB"
        .AddItem "4" & gstrSplitCmb & "O"
        .AddItem "5" & gstrSplitCmb & "����"
    End With
    'RH
    With cmbRH
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
    End With
    '����Ѫ��Ӧ��־
    With cmbBLOOD_TRAN_REACT_FLAG
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
    End With
    '���ƽ��1
    With cmbTREAT_RESULT1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "��ת"
        .AddItem "3" & gstrSplitCmb & "δ��"
        .AddItem "4" & gstrSplitCmb & "����"
        .AddItem "5" & gstrSplitCmb & "����"
    End With
    '���ƽ��2
    With cmbTREAT_RESULT2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "��ת"
        .AddItem "3" & gstrSplitCmb & "δ��"
        .AddItem "4" & gstrSplitCmb & "����"
        .AddItem "5" & gstrSplitCmb & "����"
    End With
    '���ƽ��3
    With cmbTREAT_RESULT3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "����"
        .AddItem "2" & gstrSplitCmb & "��ת"
        .AddItem "3" & gstrSplitCmb & "δ��"
        .AddItem "4" & gstrSplitCmb & "����"
        .AddItem "5" & gstrSplitCmb & "����"
    End With
    '����ʽ1
    With cmbANAESTHESIA_METHOD1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "ȫ��"
        .AddItem "2" & gstrSplitCmb & "ӲĤ��"
        .AddItem "3" & gstrSplitCmb & "����"
    End With
    '����ʽ2
    With cmbANAESTHESIA_METHOD2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "ȫ��"
        .AddItem "2" & gstrSplitCmb & "ӲĤ��"
        .AddItem "3" & gstrSplitCmb & "����"
    End With
    '����ʽ3
    With cmbANAESTHESIA_METHOD3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "ȫ��"
        .AddItem "2" & gstrSplitCmb & "ӲĤ��"
        .AddItem "3" & gstrSplitCmb & "����"
    End With
    '�пڵȼ�1
    With cmbWOUND_GRADE1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "I"
        .AddItem "2" & gstrSplitCmb & "II"
        .AddItem "3" & gstrSplitCmb & "III"
    End With
    '�пڵȼ�2
    With cmbWOUND_GRADE2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "I"
        .AddItem "2" & gstrSplitCmb & "II"
        .AddItem "3" & gstrSplitCmb & "III"
    End With
    '�пڵȼ�3
    With cmbWOUND_GRADE3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "I"
        .AddItem "2" & gstrSplitCmb & "II"
        .AddItem "3" & gstrSplitCmb & "III"
    End With
    '�п��������1
    With cmbHEAL1
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
        .AddItem "3" & gstrSplitCmb & "��"
    End With
    '�п��������2
    With cmbHEAL2
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
        .AddItem "3" & gstrSplitCmb & "��"
    End With
    '�п��������3
    With cmbHEAL3
        .Clear
        .AddItem ""
        .AddItem "1" & gstrSplitCmb & "��"
        .AddItem "2" & gstrSplitCmb & "��"
        .AddItem "3" & gstrSplitCmb & "��"
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
        '����ID
        strWhere = " And A.����ID=" & Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then
        'סԺ��
        strWhere = " And b.סԺ��='" & Mid(strCode, 2) & "'"
    Else
        'ҽ����
        strWhere = " And (b.���� Like '%" & strCode & "%' or A.ҽ���� like '%" & strCode & "%')"
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
            Nothing, sSql, 0, "ҽ������ѡ��", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, TYPE_��������, txtSickID1.Text _
            )
    If Not ChkRsState(rsTmp) Then
        txtSickID1.Text = Nvl(rsTmp!����ID)
        txtSickID1.Tag = Nvl(rsTmp!����ID) & gstrSplitCmb & Nvl(rsTmp!��ҳID)
        mlng����ID = Nvl(rsTmp!����ID)
        mlng��ҳID = Nvl(rsTmp!��ҳID)
        txtCnName1.Text = Nvl(rsTmp!����)
        txtSex1.Text = Nvl(rsTmp!�Ա�)
        txtSickID2.Text = Nvl(rsTmp!����ID)
        txtCnName2.Text = Nvl(rsTmp!����)
        txtSex2.Text = Nvl(rsTmp!�Ա�)
        txtSickID3.Text = Nvl(rsTmp!����ID)
        txtCnName3.Text = Nvl(rsTmp!����)
        txtSex3.Text = Nvl(rsTmp!�Ա�)
        txtSickID4.Text = Nvl(rsTmp!����ID)
        txtCNNAME4.Text = Nvl(rsTmp!����)
        txtSex4.Text = Nvl(rsTmp!�Ա�)
    Else
        MsgBox "û���ҵ�������Ϣ!", vbInformation, gstrSysName
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
    '��ʼ��
    Call InitTabControl
    Call InitCmbControl
    If mstrHospitalNumber <> "" Then
        '�޸�
        gstrSQL = "Select * from ���β�����Ϣ Where RESIDENCE_NO=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrHospitalNumber)
        '�ı���ֵ
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
                mlng����ID = txtSickID1.Tag
                mlng��ҳID = Nvl(!IN_COUNT)
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
    Dim str����ҩ��      As String
On Error GoTo ErrH
    If txtSickID1.Tag = "" Then
        MsgBox "��ѡ���ˣ�", vbCritical, gstrSysName
        Exit Sub
    End If
    '==============================================================================
    '=������Ϣ
    '==============================================================================
    txtHOSPITAL_NUMBER.Text = gstrҽԺ����
    #If gverControl < 6 Then
        'ȡ���˻�����Ϣ
        gstrSQL = " SELECT A.סԺ��,B.ҽ�Ƹ��ʽ,B.��ҳID,B.������,A.����,A.�Ա�,A.��������,B.����״��,B.ְҵ,A.�����ص�," & _
             "        H.���� AS ����,B.����,A.���֤��,A.������λ,B.��λ��ַ,B.��λ�绰,B.��λ�ʱ�,B.��ͥ��ַ,B.�����ʱ�," & _
             "        B.��ϵ������,B.��ϵ�˹�ϵ,B.��ϵ�˵�ַ,B.��ϵ�˵绰,B.��Ժ����,D.���� AS ��Ժ����,E.���� AS ��Ժ����," & _
             "        B.��Ժ����,F.���� AS ��Ժ����,B.��Ժ����,B.ȷ������,B.���ȴ���,B.�ɹ�����,B.��Ժ��ʽ," & _
             "        B.��ĿԱ����,NVL(B.��Ŀ����,SYSDATE) AS ��Ŀ����,B.ʬ���־,B.�����־,B.��������,B.Ѫ��,B.סԺҽʦ" & _
             " FROM ������Ϣ A,������ҳ B,��Լ��λ C,���ű� D,���ű� E,���ű� F,���� H" & _
             " WHERE A.����ID=B.����ID AND A.סԺ����=B.��ҳID AND A.��ͬ��λID=C.ID(+)" & _
             " AND B.��Ժ����ID=D.ID(+) AND B.��Ժ����ID=E.ID(+) AND B.��Ժ����ID=F.ID(+) " & _
             " AND A.����=H.���� AND B.����ID = [1] AND B.��ҳID = [2]"
    #Else
        'ȡ���˻�����Ϣ
        gstrSQL = " SELECT A.סԺ��,B.ҽ�Ƹ��ʽ,B.��ҳID,B.������,A.����,A.�Ա�,A.��������,B.����״��,B.ְҵ,A.�����ص�," & _
             "        H.���� AS ����,B.����,A.���֤��,A.������λ,B.��λ��ַ,B.��λ�绰,B.��λ�ʱ�,B.��ͥ��ַ,B.��ͥ��ַ�ʱ� As �����ʱ�," & _
             "        B.��ϵ������,B.��ϵ�˹�ϵ,B.��ϵ�˵�ַ,B.��ϵ�˵绰,B.��Ժ����,D.���� AS ��Ժ����,E.���� AS ��Ժ����," & _
             "        B.��Ժ����,F.���� AS ��Ժ����,B.��Ժ����,B.ȷ������,B.���ȴ���,B.�ɹ�����,B.��Ժ��ʽ," & _
             "        B.��ĿԱ����,NVL(B.��Ŀ����,SYSDATE) AS ��Ŀ����,B.ʬ���־,B.�����־,B.��������,B.Ѫ��,B.סԺҽʦ" & _
             " FROM ������Ϣ A,������ҳ B,��Լ��λ C,���ű� D,���ű� E,���ű� F,���� H" & _
             " WHERE A.����ID=B.����ID AND A.סԺ����=B.��ҳID AND A.��ͬ��λID=C.ID(+)" & _
             " AND B.��Ժ����ID=D.ID(+) AND B.��Ժ����ID=E.ID(+) AND B.��Ժ����ID=F.ID(+) " & _
             " AND A.����=H.���� AND B.����ID = [1] AND B.��ҳID = [2]"
    #End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTmp
        'txtRESIDENCE_NO.Text = Nvl(!סԺ��) & "_" & Nvl(!��ҳID)
        'С��������2011-06-29��סԺ�Ų���ȡ��ҳID��ȡ סԺ����
              
        txtRESIDENCE_NO.Text = Nvl(!סԺ��)
        txtIN_COUNT.Text = !��ҳID
        txtMEDICAL_RECORD_NO.Text = Nvl(!סԺ��)
        cmbMARITAL_STATUS.ListIndex = Cmb_EditIndex(cmbMARITAL_STATUS, TRANDATA("����", Nvl(!����״��, "δ��")))
        txtSTATUS.Text = ChkStrUniCode(Nvl(!ְҵ), txtSTATUS.MaxLength)
        txtBIRTH_ADDRESS.Text = ChkStrUniCode(Nvl(!�����ص�), txtBIRTH_ADDRESS.MaxLength)
        txtIDENTITY_NUMBER.Text = ChkStrUniCode(Nvl(!���֤��), txtIDENTITY_NUMBER.MaxLength)
        txtUNIT_NAME.Text = ChkStrUniCode(Nvl(!������λ), txtUNIT_NAME.MaxLength)
        txtUNIT_ADDRESS.Text = ChkStrUniCode(Nvl(!��λ��ַ), txtUNIT_ADDRESS.MaxLength)
        txtUNIT_PHONE.Text = ChkStrUniCode(Nvl(!��λ�绰), txtUNIT_PHONE.MaxLength)
        txtUNIT_ZIPCODE.Text = ChkStrUniCode(Nvl(!��λ�ʱ�), txtUNIT_ZIPCODE.MaxLength)
        txtREGISTER_ADDRESS.Text = ChkStrUniCode(Nvl(!��ͥ��ַ), txtREGISTER_ADDRESS.MaxLength)
        txtREGISTER_ZIPCODE.Text = ChkStrUniCode(Nvl(!�����ʱ�), txtREGISTER_ZIPCODE.MaxLength)
        txtCONTACT_PERSON.Text = ChkStrUniCode(Nvl(!��ϵ������), txtCONTACT_PERSON.MaxLength)
        cmbRELATIONSHIP.ListIndex = Cmb_EditIndex(cmbMARITAL_STATUS, TRANDATA("�벡�˹�ϵ", Nvl(!��ϵ�˹�ϵ)))
        txtCONTACT_ADDRESS.Text = ChkStrUniCode(Nvl(!��ϵ�˵�ַ), txtCONTACT_ADDRESS.MaxLength)
        txtCONTACT_PHONE.Text = ChkStrUniCode(Nvl(!��ϵ�˵绰), txtCONTACT_PHONE.MaxLength)
        txtADMISSION_DATE.Text = Format(!��Ժ����, "YYYY-MM-DD HH:MM:SS")
        txtADMISSION_DEPT.Text = ChkStrUniCode(Nvl(!��Ժ����), txtADMISSION_DEPT.MaxLength)
        txtIN_DEPT_ZONE.Text = ChkStrUniCode(Nvl(!��Ժ����), txtIN_DEPT_ZONE.MaxLength)
        txtDEPT_TRANSFERED_TO.Text = "" 'ChkStrUniCode("", txtDEPT_TRANSFERED_TO.MaxLength)
        txtDISCHARGE_DATE.Text = Format(!��Ժ����, "YYYY-MM-DD HH:MM:SS")
        txtDISCHARGE_DEPT.Text = ChkStrUniCode(Nvl(!��Ժ����), txtDISCHARGE_DEPT.MaxLength)
        txtOUT_DEPT_ZONE.Text = ChkStrUniCode(Nvl(!��Ժ����), txtDISCHARGE_DEPT.MaxLength)
        cmbPAT_ADM_CONDITION.ListIndex = Cmb_EditIndex(cmbPAT_ADM_CONDITION, TRANDATA("��Ժ����", Nvl(!��Ժ����)))
        txtDIAGNOSIS_DATE.Text = Format(!ȷ������, "YYYY-MM-DD HH:MM:SS")
        'ҽʦ��Ϣ
        txtINHOSPITAL_DOCTOR = ChkStrUniCode(Nvl(!סԺҽʦ), txtINHOSPITAL_DOCTOR.MaxLength)
        txtINTERM.Text = ChkStrUniCode(Nvl(!��ĿԱ����), txtINTERM.MaxLength)
        cmbBODY_EXAMINE_FLAG.ListIndex = Cmb_EditIndex(cmbBODY_EXAMINE_FLAG, TRANDATA("ʬ���־", Nvl(!ʬ���־, "��")))
        cmbFOLLOW_FLAG.ListIndex = Cmb_EditIndex(cmbFOLLOW_FLAG, TRANDATA("�����־", Nvl(!�����־)))
        txtFOLLOW_TERM.Text = ChkStrUniCode(Nvl(!��������, 0), txtFOLLOW_TERM.MaxLength)
        '�����Ϣ
        cmbBLOOD_TYPE.ListIndex = Cmb_EditIndex(cmbBLOOD_TYPE, TRANDATA("Ѫ��", Nvl(!Ѫ��)))
        txtHANDLE.Text = ChkStrUniCode(Nvl(!��ĿԱ����, UserInfo.����), txtHANDLE.MaxLength)
        txtHANDLE_DATE.Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    End With
    '==============================================================================
    '=ҽʦ��Ϣ
    '==============================================================================
    str����ҩ�� = ""
    '��ѡһ�ֹ���ҩ��
    gstrSQL = " SELECT ����ҩ�� FROM ���˹���ҩ�� WHERE ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
    Do While Not ChkRsState(rsTmp)
        str����ҩ�� = str����ҩ�� & " ," & Trim(Nvl(rsTmp!����ҩ��))
        rsTmp.MoveNext
    Loop
    txtALERGY_DRUGS.Text = ChkStrUniCode(Mid(str����ҩ��, 2), txtALERGY_DRUGS.MaxLength)
    
    'ȡ������ҳ�ӱ�
    Dim STR��Ϣֵ As String
    gstrSQL = "SELECT UPPER(��Ϣ��) AS ��Ϣ��,��Ϣֵ FROM ������ҳ�ӱ� WHERE ����ID=[1] AND ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTmp
        Do While Not .EOF
            STR��Ϣֵ = Nvl(!��Ϣֵ)
            Select Case !��Ϣ��
                Case "HBSAG"
                    cmbHBSAG.ListIndex = Cmb_EditIndex(cmbHBSAG, TRANDATA("HBSAG", STR��Ϣֵ))
                Case "HCV-AB"
                    cmbHCV_AB.ListIndex = Cmb_EditIndex(cmbHCV_AB, TRANDATA("HCV-AB", STR��Ϣֵ))
                Case "HIV-AB"
                    cmbHIV_AB.ListIndex = Cmb_EditIndex(cmbHIV_AB, TRANDATA("HIV-AB", STR��Ϣֵ))
                Case "������"
                    txtDIRECTOR.Text = ChkStrUniCode(STR��Ϣֵ, txtDIRECTOR.MaxLength)
                Case "����ҽʦ"
                    txtDIRECTOR_DOCTOR.Text = ChkStrUniCode(STR��Ϣֵ, txtDIRECTOR_DOCTOR.MaxLength)
                Case "����ҽʦ"
                    txtATTENDING_DOCTOR.Text = ChkStrUniCode(STR��Ϣֵ, txtATTENDING_DOCTOR.MaxLength)
                Case "����ҽʦ"
                    txtREFRESH_DOCTOR.Text = ChkStrUniCode(STR��Ϣֵ, txtREFRESH_DOCTOR.MaxLength)
                Case "�о���ʵϰҽʦ"
                    txtGRADUATE_DOCTOR.Text = ChkStrUniCode(STR��Ϣֵ, txtGRADUATE_DOCTOR.MaxLength)
                Case "ʵϰҽʦ"
                    txtINTERM.Text = ChkStrUniCode(STR��Ϣֵ, txtINTERM.MaxLength)
'                Case "��������"
'                    cmbMEDICAL_RECORD_MASS.ListIndex = Cmb_EditIndex(cmbMEDICAL_RECORD_MASS, TRANDATA("��������", STR��Ϣֵ))
                Case "�ʿ�ҽʦ"
                    txtCONTROL_DOCTOR.Text = ChkStrUniCode(STR��Ϣֵ, txtCONTROL_DOCTOR.MaxLength)
                Case "�ʿػ�ʦ"
                    txtCONTROL_NURSE.Text = ChkStrUniCode(STR��Ϣֵ, txtCONTROL_NURSE.MaxLength)
                Case "����"
                    cmbFIRST_FLAG.ListIndex = Cmb_EditIndex(cmbFIRST_FLAG, TRANDATA("����", STR��Ϣֵ))
                Case "ʾ�̲���"
                    cmbTEACH_MR_FLAG.ListIndex = Cmb_EditIndex(cmbTEACH_MR_FLAG, TRANDATA("ʾ�̲���", STR��Ϣֵ))
                Case "RH"
                    cmbRH.ListIndex = Cmb_EditIndex(cmbRH, TRANDATA("RH", STR��Ϣֵ))
                Case "��Ѫ��Ӧ"
                    cmbBLOOD_TRAN_REACT_FLAG.ListIndex = Cmb_EditIndex(cmbBLOOD_TRAN_REACT_FLAG, TRANDATA("��Ѫ��Ӧ", STR��Ϣֵ))
                Case "���ϸ��"
                    txtERYTHROCYTE.Text = IIf(Val(STR��Ϣֵ) < 90009000.99 Or Val(STR��Ϣֵ) > 0, Val(STR��Ϣֵ), 0)
                Case "��ѪС��"
                    txtHEMOBLAST.Text = IIf(Val(STR��Ϣֵ) < 90009000.99 Or Val(STR��Ϣֵ) > 0, Val(STR��Ϣֵ), 0)
                Case "��Ѫ��"
                    txtPLASM.Text = IIf(Val(STR��Ϣֵ) < 90009000.99 Or Val(STR��Ϣֵ) > 0, Val(STR��Ϣֵ), 0)
                Case "��ȫѪ"
                    txtBLOOD.Text = IIf(Val(STR��Ϣֵ) < 90009000.99 Or Val(STR��Ϣֵ) > 0, Val(STR��Ϣֵ), 0)
                Case "������"
                    txtOTHER_BLOOD.Text = IIf(Val(STR��Ϣֵ) < 90009000.99 Or Val(STR��Ϣֵ) > 0, Val(STR��Ϣֵ), 0)
            End Select
            .MoveNext
        Loop
    End With
    
    'ȡ������
    gstrSQL = "SELECT ��������,NVL(�������,0) AS ������� FROM ��Ϸ������ WHERE ����ID=[1] AND ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTmp
        Do While Not ChkRsState(rsTmp)
            Select Case !��������
                Case 1  '�������Ժ
                    cmdCLINIC_INHOSPITAL.ListIndex = Cmb_EditIndex(cmdCLINIC_INHOSPITAL, Nvl(!�������))
                Case 2  '��Ժ���Ժ
                    cmbIN_OUT.ListIndex = Cmb_EditIndex(cmbIN_OUT, Nvl(!�������))
                Case 3  '�����벡��
                    cmbEMIT_PATHOLOGY.ListIndex = Cmb_EditIndex(cmbEMIT_PATHOLOGY, Nvl(!�������))
                Case 4  '�ٴ��벡��
                    cmbCLINIC_PATHOLOGY.ListIndex = Cmb_EditIndex(cmbCLINIC_PATHOLOGY, Nvl(!�������))
                Case 6  '��ǰ������
                    cmbBEFORE_AFTER_TREATMENT.ListIndex = Cmb_EditIndex(cmbBEFORE_AFTER_TREATMENT, Nvl(!�������))
            End Select
            .MoveNext
        Loop
    End With

    'ȡ�������ֽ�����Ӳ�����ҳ�ӱ��ж�ȡ��ֻҪ��д�˲����Ķ��������ݣ�
    gstrSQL = "SELECT �ȼ� FROM �������ֽ�� WHERE ����ID=[1] AND ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not ChkRsState(rsTmp) Then
       cmbMEDICAL_RECORD_MASS.ListIndex = Cmb_EditIndex(cmbMEDICAL_RECORD_MASS, rsTmp!�ȼ�)
    End If
    'ȡ��������
    gstrSQL = "SELECT ����Ա����,�շ�ʱ�� FROM ���˽��ʼ�¼ WHERE ID = (SELECT MAX(ID) FROM ���˽��ʼ�¼ WHERE ����ID=[1] AND ��¼״̬=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not ChkRsState(rsTmp) Then
        txtBAL_DATE.Text = Format(rsTmp!�շ�ʱ��, "YYYY-MM-DD HH:MM:SS")
    End If
    '==============================================================================
    '=�������
    '==============================================================================
    '��Ժ���
'    gstrSQL = " SELECT A.�������,A.��ϴ���,A.�������,B.���� AS ��������,A.�������,A.��¼��,NVL(A.��¼����,SYSDATE) AS ��¼����" & _
'             " FROM ������ϼ�¼ A,��������Ŀ¼ B" & _
'             " WHERE A.����ID=B.ID AND A.��¼��Դ=3 AND A.������� In ('2','12') AND A.����ID=[1] AND A.��ҳID=[2]"
' С������
gstrSQL = " SELECT A.�������,A.��ϴ���,A.�������,B.���� AS ��������,A.�������,A.��¼��,NVL(A.��¼����,SYSDATE) AS ��¼����" & _
             " FROM ������ϼ�¼ A,��������Ŀ¼ B" & _
             " WHERE A.����ID=B.ID and A.��¼��Դ=4 AND A.������� In ('2','12') and A.����ID=[1] AND A.��ҳID=[2]"
             
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not ChkRsState(rsTmp) Then
        txtIN_DIAGNOSIS_CODE.Text = ChkStrUniCode(Nvl(rsTmp!��������), txtIN_DIAGNOSIS_CODE.MaxLength)
        txtIN_DIAGNOSIS_NAME.Text = ChkStrUniCode(Nvl(rsTmp!�������), txtIN_DIAGNOSIS_NAME.MaxLength)
        txtIN_DIAGNOSIS_DATE.Text = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
    End If
    '��Ժ���
    gstrSQL = " SELECT A.�������,A.��ϴ���,A.�������,B.���� AS ��������,A.�������,A.��¼��,NVL(A.��¼����,SYSDATE) AS ��¼����,��Ժ���" & _
             " FROM ������ϼ�¼ A,��������Ŀ¼ B" & _
             " WHERE A.����ID=B.ID AND A.��¼��Դ=4 AND A.������� In ('3','13') AND A.����ID=[1] AND A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Do While Not ChkRsState(rsTmp)
        If rsTmp.Bookmark = 1 Then
            txtOUT_DIAGNOSIS_CODE1.Text = ChkStrUniCode(Nvl(rsTmp!��������), txtOUT_DIAGNOSIS_CODE1.MaxLength)
            txtOUT_DIAGNOSIS_NAME1.Text = ChkStrUniCode(Nvl(rsTmp!�������), txtOUT_DIAGNOSIS_NAME1.MaxLength)
            txtOUT_DIAGNOSIS_DATE1.Text = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
            cmbTREAT_RESULT1.ListIndex = Cmb_EditIndex(cmbTREAT_RESULT1, TRANDATA("��Ժ���", Nvl(rsTmp!��Ժ���)))
        ElseIf rsTmp.Bookmark = 2 Then
            txtOUT_DIAGNOSIS_CODE2.Text = ChkStrUniCode(Nvl(rsTmp!��������), txtOUT_DIAGNOSIS_CODE2.MaxLength)
            txtOUT_DIAGNOSIS_NAME2.Text = ChkStrUniCode(Nvl(rsTmp!�������), txtOUT_DIAGNOSIS_NAME2.MaxLength)
            txtOUT_DIAGNOSIS_DATE2.Text = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
            cmbTREAT_RESULT2.ListIndex = Cmb_EditIndex(cmbTREAT_RESULT2, TRANDATA("��Ժ���", Nvl(rsTmp!��Ժ���)))
        ElseIf rsTmp.Bookmark = 3 Then
            txtOUT_DIAGNOSIS_CODE3.Text = ChkStrUniCode(Nvl(rsTmp!��������), txtOUT_DIAGNOSIS_CODE3.MaxLength)
            txtOUT_DIAGNOSIS_NAME3.Text = ChkStrUniCode(Nvl(rsTmp!�������), txtOUT_DIAGNOSIS_NAME3.MaxLength)
            txtOUT_DIAGNOSIS_DATE3.Text = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
            cmbTREAT_RESULT3.ListIndex = Cmb_EditIndex(cmbTREAT_RESULT3, TRANDATA("��Ժ���", Nvl(rsTmp!��Ժ���)))
        End If
        rsTmp.MoveNext
    Loop
    '==============================================================================
    '=������Ϣ
    '==============================================================================
    gstrSQL = " SELECT B.����,B.����,A.�п�,A.����,A.��������,A.��������,A.����ҽʦ,A.��һ����,A.�ڶ�����,A.����ҽʦ,A.��¼��,NVL(A.��¼����,SYSDATE) AS ��¼���� " & _
             " FROM ���������¼ A ,��������Ŀ¼ B " & _
             " WHERE A.��������ID=B.ID And A.����ID=[1] AND A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Do While Not ChkRsState(rsTmp)
        If rsTmp.Bookmark = 1 Then
            txtOPERATION_CODE1.Text = ChkStrUniCode(Nvl(rsTmp!����), txtOPERATION_CODE1.MaxLength)
            txtOPERATION_NAME1.Text = ChkStrUniCode(Nvl(rsTmp!����), txtOPERATION_NAME1.MaxLength)
            cmbWOUND_GRADE1.ListIndex = Cmb_EditIndex(cmbWOUND_GRADE1, TRANDATA("�п�", Nvl(rsTmp!�п�)))
            cmbHEAL1.ListIndex = Cmb_EditIndex(cmbHEAL1, TRANDATA("����", Nvl(rsTmp!����)))
            txtOPERATING_DATE1 = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
            cmbANAESTHESIA_METHOD1.ListIndex = Cmb_EditIndex(cmbANAESTHESIA_METHOD1, TRANDATA("��������", Nvl(rsTmp!��������)))
            
        ElseIf rsTmp.Bookmark = 2 Then
            txtOPERATION_CODE2.Text = ChkStrUniCode(Nvl(rsTmp!����), txtOPERATION_CODE2.MaxLength)
            txtOPERATION_NAME2.Text = ChkStrUniCode(Nvl(rsTmp!����), txtOPERATION_NAME2.MaxLength)
            cmbWOUND_GRADE2.ListIndex = Cmb_EditIndex(cmbWOUND_GRADE2, TRANDATA("�п�", Nvl(rsTmp!�п�)))
            cmbHEAL2.ListIndex = Cmb_EditIndex(cmbHEAL2, TRANDATA("����", Nvl(rsTmp!����)))
            txtOPERATING_DATE2.Text = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
            cmbANAESTHESIA_METHOD2.ListIndex = Cmb_EditIndex(cmbANAESTHESIA_METHOD2, TRANDATA("��������", Nvl(rsTmp!��������)))
        ElseIf rsTmp.Bookmark = 3 Then
            txtOPERATION_CODE3.Text = ChkStrUniCode(Nvl(rsTmp!����), txtOPERATION_CODE3.MaxLength)
            txtOPERATION_NAME3.Text = ChkStrUniCode(Nvl(rsTmp!����), txtOPERATION_NAME3.MaxLength)
            cmbWOUND_GRADE3.ListIndex = Cmb_EditIndex(cmbWOUND_GRADE3, TRANDATA("�п�", Nvl(rsTmp!�п�)))
            cmbHEAL3.ListIndex = Cmb_EditIndex(cmbHEAL3, TRANDATA("����", Nvl(rsTmp!����)))
            txtOPERATING_DATE3.Text = Format(Nvl(rsTmp!��¼����), "yyyy-mm-dd hh:mm:ss")
            cmbANAESTHESIA_METHOD3.ListIndex = Cmb_EditIndex(cmbANAESTHESIA_METHOD3, TRANDATA("��������", Nvl(rsTmp!��������)))
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
    labNot.Caption = IIf(Len(txtNot.Text) > 0, "��", "*")
    labNot.ForeColor = IIf(Len(txtNot.Text) > 0, vbGreen, vbRed)
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub SetNotComb(labNot As Label, cmbNot As ComboBox)
On Error GoTo ErrH
    labNot.Caption = IIf(Len(cmbNot.Text) > 0, "��", "*")
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
    labNot(5).Caption = IIf(IsDate(txtADMISSION_DATE.Text), "��", "*")
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
    labNot(11).Caption = IIf(IsDate(txtHANDLE_DATE.Text), "��", "*")
    labNot(11).ForeColor = IIf(IsDate(txtHANDLE_DATE.Text), vbGreen, vbRed)
End Sub

Private Sub txtIN_DIAGNOSIS_CODE_Change()
    Call SetNotText(labNot(12), txtIN_DIAGNOSIS_CODE)
End Sub

Private Sub txtIN_DIAGNOSIS_NAME_Change()
    Call SetNotText(labNot(13), txtIN_DIAGNOSIS_NAME)
End Sub

Private Sub txtIN_DIAGNOSIS_DATE_Change()
    labNot(14).Caption = IIf(IsDate(txtIN_DIAGNOSIS_DATE.Text), "��", "*")
    labNot(14).ForeColor = IIf(IsDate(txtIN_DIAGNOSIS_DATE.Text), vbGreen, vbRed)
End Sub

Private Sub txtOUT_DIAGNOSIS_CODE1_Change()
    Call SetNotText(labNot(15), txtOUT_DIAGNOSIS_CODE1)
End Sub

Private Sub txtOUT_DIAGNOSIS_NAME1_Change()
    Call SetNotText(labNot(16), txtOUT_DIAGNOSIS_NAME1)
End Sub

Private Sub txtOUT_DIAGNOSIS_DATE1_Change()
    labNot(17).Caption = IIf(IsDate(txtOUT_DIAGNOSIS_DATE1.Text), "��", "*")
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
    labNot(20).Caption = IIf(IsDate(txtOUT_DIAGNOSIS_DATE2.Text), "��", "*")
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
    labNot(23).Caption = IIf(IsDate(txtOUT_DIAGNOSIS_DATE3.Text), "��", "*")
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
    labNot(28).Caption = IIf(IsDate(txtOPERATING_DATE1.Text), "��", "*")
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
    labNot(33).Caption = IIf(IsDate(txtOPERATING_DATE2.Text), "��", "*")
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
    labNot(38).Caption = IIf(IsDate(txtOPERATING_DATE3.Text), "��", "*")
    labNot(38).ForeColor = IIf(IsDate(txtOPERATING_DATE3.Text), vbGreen, vbRed)
End Sub

Private Sub cmbANAESTHESIA_METHOD3_Click()
    Call SetNotComb(labNot(39), cmbANAESTHESIA_METHOD2)
End Sub

Private Function TRANDATA(ByVal STR��Ϣ�� As String, ByVal STR��Ϣֵ As String) As String
    '���ݽӿ��ĵ�ת��HIS�е�ֵ
    Select Case STR��Ϣ��
    Case "ҽ�Ƹ��ʽ"
        Select Case STR��Ϣֵ
            Case "������ҽ�Ʊ���"
                TRANDATA = 1
            Case "��ҵ����"
                TRANDATA = 2
            Case "�Է�ҽ��"
                TRANDATA = 3
            Case "����ҽ��"
                TRANDATA = 4
            Case "��ͳ��"
                TRANDATA = 5
            Case Else   '����
                TRANDATA = 6
        End Select
    Case "�Ա�"
        Select Case STR��Ϣֵ
            Case "��"
                TRANDATA = 1
            Case Else   'Ů
                TRANDATA = 2
        End Select
    Case "����"
        Select Case STR��Ϣֵ
            Case "δ��"
                TRANDATA = 1
            Case "�ѻ�"
                TRANDATA = 2
            Case "���"
                TRANDATA = 3
            Case Else   'ɥ
                TRANDATA = 4
        End Select
    Case "�벡�˹�ϵ"
        Select Case STR��Ϣֵ
            Case "��ż"
                TRANDATA = 1
            Case "��", "Ů"
                TRANDATA = 2
            Case "��ĸ"
                TRANDATA = 3
            Case Else   '����\��Ů\�游\��ĸ\���˵ȵ�,����������
                TRANDATA = 9
        End Select
    Case "ʬ���־", "����", "�����־", "ʾ�̲���", "RH", "��Ѫ��Ӧ"
        Select Case STR��Ϣֵ
            Case "��"
                TRANDATA = 1
            Case Else
                TRANDATA = 2
        End Select
    Case "Ѫ��"
        Select Case STR��Ϣֵ
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
    Case "��������"
        Select Case STR��Ϣֵ
            Case "ȫ��"
                TRANDATA = 1
            Case "����"
                TRANDATA = 3
            Case Else
                TRANDATA = 2
        End Select
    Case "��������"
        Select Case STR��Ϣֵ
            Case "��"
                TRANDATA = 1
            Case "��"
                TRANDATA = 2
            Case Else
                TRANDATA = 3
        End Select
    Case "���ƽ��", "��Ժ���"
        Select Case STR��Ϣֵ
            Case "����", "����"
                TRANDATA = 1
            Case "��ת"
                TRANDATA = 2
            Case "δ��"
                TRANDATA = 3
            Case "����"
                TRANDATA = 4
            Case Else
                TRANDATA = 5
        End Select
    Case "HBSAG", "HCV_AB", "HIV_AB"
        Select Case STR��Ϣֵ
            Case "����"
                TRANDATA = 1
            Case "����"
                TRANDATA = 2
            Case Else
                TRANDATA = 0
        End Select
    Case "��Ժ����"
        Select Case STR��Ϣֵ
            Case "Σ"
                TRANDATA = 1
            Case "��"
                TRANDATA = 2
            Case Else
                TRANDATA = 3
        End Select
    Case "�п�"
        Select Case STR��Ϣֵ
            Case "��"
                TRANDATA = "1"
            Case "��"
                TRANDATA = "2"
            Case "��"
                TRANDATA = "3"
        End Select
    Case "����"     '�ӿ�����ͳһ�жϵ�
        Select Case STR��Ϣֵ
            Case "��"
                TRANDATA = "1"
            Case "��"
                TRANDATA = "2"
            Case "��"
                TRANDATA = "3"
        End Select
    Case "�������"
        Select Case STR��Ϣֵ
            Case 5, 6, 7
                TRANDATA = Val(STR��Ϣֵ) - 1
            Case 1, 2, 3
                TRANDATA = Val(STR��Ϣֵ)
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
    '��������
    gstrSQL = ""
    gstrSQL = gstrSQL & vbCrLf & "zl_���β�����Ϣ_Update("
    '������Ϣ
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
    'ҽʦ��Ϣ
    gstrSQL = gstrSQL & vbCrLf & "'" & txtALERGY_DRUGS.Text & "','" & Cmb_ID(cmbHBSAG) & "','" & Cmb_ID(cmbHCV_AB) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbHIV_AB) & "','" & Cmb_ID(cmdCLINIC_INHOSPITAL) & "','" & Cmb_ID(cmbIN_OUT) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbBEFORE_AFTER_TREATMENT) & "','" & Cmb_ID(cmbCLINIC_PATHOLOGY) & "','" & Cmb_ID(cmbEMIT_PATHOLOGY) & "',"
    
    gstrSQL = gstrSQL & vbCrLf & "" & Val(txtEMER_TREAT_TIMES.Text) & "," & Val(txtESC_EMER_TIMES.Text) & ",'" & txtDIRECTOR.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtDIRECTOR_DOCTOR.Text & "','" & txtATTENDING_DOCTOR.Text & "','" & txtINHOSPITAL_DOCTOR.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtREFRESH_DOCTOR.Text & "','" & txtGRADUATE_DOCTOR.Text & "','" & txtINTERM.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtCODE_NAME.Text & "','" & Cmb_ID(cmbMEDICAL_RECORD_MASS) & "','" & txtCONTROL_DOCTOR.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtCONTROL_NURSE.Text & "',to_date('" & txtBAL_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbBODY_EXAMINE_FLAG) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbFIRST_FLAG) & "','" & Cmb_ID(cmbFOLLOW_FLAG) & "','" & txtFOLLOW_TERM.Text & "',"
    '�����Ϣ
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbTEACH_MR_FLAG) & "','" & Cmb_ID(cmbBLOOD_TYPE) & "','" & Cmb_ID(cmbRH) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbBLOOD_TRAN_REACT_FLAG) & "'," & Val(txtERYTHROCYTE) & "," & Val(txtHEMOBLAST) & ","
    gstrSQL = gstrSQL & vbCrLf & "" & Val(txtPLASM.Text) & "," & Val(txtBLOOD.Text) & ",'" & Val(txtOTHER_BLOOD.Text) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtHANDLE.Text & "',to_date('" & txtHANDLE_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & txtIN_DIAGNOSIS_CODE.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtIN_DIAGNOSIS_NAME.Text & "',to_date('" & txtIN_DIAGNOSIS_DATE.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & txtOUT_DIAGNOSIS_CODE1.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOUT_DIAGNOSIS_NAME1.Text & "',to_date('" & txtOUT_DIAGNOSIS_DATE1.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbTREAT_RESULT1) & "',"
    gstrSQL = gstrSQL & vbCrLf & "'" & txtOUT_DIAGNOSIS_CODE2.Text & "','" & txtOUT_DIAGNOSIS_NAME2.Text & "',to_date('" & txtOUT_DIAGNOSIS_DATE2.Text & "','yyyy-mm-dd hh24:mi:ss'),"
    gstrSQL = gstrSQL & vbCrLf & "'" & Cmb_ID(cmbTREAT_RESULT2) & "','" & txtOUT_DIAGNOSIS_CODE3.Text & "','" & txtOUT_DIAGNOSIS_NAME3.Text & "',"
    gstrSQL = gstrSQL & vbCrLf & "to_date('" & txtOUT_DIAGNOSIS_DATE3.Text & "','yyyy-mm-dd hh24:mi:ss'),'" & Cmb_ID(cmbTREAT_RESULT3) & "',"
    '������Ϣ
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

'������¼������
Private Function chkData() As Boolean
    Dim strMsg      As String
    Dim blnTbc      As Boolean
On Error GoTo ErrH
    strMsg = ""
    blnTbc = False
    If Len(Trim(txtHOSPITAL_NUMBER.Text)) = 0 Then
        strMsg = strMsg & "��������Ϣ��" & labHOSPITAL_NUMBER.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Len(Trim(txtRESIDENCE_NO.Text)) = 0 Then
        strMsg = strMsg & "��������Ϣ��" & labRESIDENCE_NO.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Val(Trim(txtIN_COUNT.Text)) <= 0 Then
        strMsg = strMsg & "��������Ϣ��" & labIN_COUNT.Caption & "������Ϊ�գ��Ҵ����㣡" & vbCrLf
    End If
    If Len(Trim(txtMEDICAL_RECORD_NO.Text)) = 0 Then
        strMsg = strMsg & "��������Ϣ��" & labMEDICAL_RECORD_NO.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Not IsDate(Trim(txtADMISSION_DATE.Text)) Then
        strMsg = strMsg & "��������Ϣ��" & labADMISSION_DATE.Caption & "������Ϊ�գ�Ϊ�������ͣ�" & vbCrLf
    End If
    If Len(Trim(txtADMISSION_DEPT.Text)) = 0 Then
        strMsg = strMsg & "��������Ϣ��" & labADMISSION_DEPT.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Len(Trim(txtIN_DEPT_ZONE.Text)) = 0 Then
        strMsg = strMsg & "��������Ϣ��" & labIN_DEPT_ZONE.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Not IsDate(Trim(txtDISCHARGE_DATE.Text)) Then
        strMsg = strMsg & "��������Ϣ��" & labDISCHARGE_DATE.Caption & "������Ϊ�գ�Ϊ�������ͣ�" & vbCrLf
    End If
    If strMsg <> "" Then
        tbcPage.Item(0).Selected = True
        blnTbc = True
    End If
    'ҽʦ��Ϣ
    If Len(Trim(cmbIN_OUT.Text)) = 0 Then
        strMsg = strMsg & "��ҽʦ��Ϣ��" & labIN_OUT.Caption & "������Ϊ�գ���ѡ��" & vbCrLf
    End If
    If strMsg <> "" And Not blnTbc Then
        tbcPage.Item(1).Selected = True
        blnTbc = True
    End If
    '�����Ϣ
    If Len(Trim(txtHANDLE.Text)) = 0 Then
        strMsg = strMsg & "�������Ϣ��" & labHANDLE.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Not IsDate(Trim(txtHANDLE_DATE.Text)) Then
        strMsg = strMsg & "�������Ϣ��" & labHANDLE_DATE.Caption & "������Ϊ�գ�Ϊ�������ͣ�" & vbCrLf
    End If
    If Len(Trim(txtIN_DIAGNOSIS_CODE.Text)) = 0 Then
        strMsg = strMsg & "�������Ϣ��" & labIN_DIAGNOSIS_CODE.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Len(Trim(txtIN_DIAGNOSIS_NAME.Text)) = 0 Then
        strMsg = strMsg & "�������Ϣ��" & labIN_DIAGNOSIS_NAME.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Not IsDate(Trim(txtIN_DIAGNOSIS_DATE.Text)) Then
        strMsg = strMsg & "�������Ϣ��" & labIN_DIAGNOSIS_DATE.Caption & "������Ϊ�գ�Ϊ�������ͣ�" & vbCrLf
    End If
    If Len(Trim(txtOUT_DIAGNOSIS_CODE1.Text)) = 0 Then
        strMsg = strMsg & "�������Ϣ��" & labOUT_DIAGNOSIS_CODE1.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Len(Trim(txtOUT_DIAGNOSIS_NAME1.Text)) = 0 Then
        strMsg = strMsg & "�������Ϣ��" & labOUT_DIAGNOSIS_NAME1.Caption & "������Ϊ�գ�" & vbCrLf
    End If
    If Not IsDate(Trim(txtOUT_DIAGNOSIS_DATE1.Text)) Then
        strMsg = strMsg & "�������Ϣ��" & labOUT_DIAGNOSIS_DATE1.Caption & "������Ϊ�գ�Ϊ�������ͣ�" & vbCrLf
    End If
    If Len(Trim(cmbTREAT_RESULT1.Text)) = 0 Then
        strMsg = strMsg & "�������Ϣ��" & labTREAT_RESULT1.Caption & "������Ϊ�գ���ѡ��" & vbCrLf
    End If
    '������2�����3��Ϊ�գ���ô�������2�����ݶ�������д
    If Len(Trim(txtOUT_DIAGNOSIS_CODE2.Text & txtOUT_DIAGNOSIS_CODE3.Text)) <> 0 Then
        '������2����Ϣ
        
    End If
    '������3�����ݲ�Ϊ�գ���ô������д������
    
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
