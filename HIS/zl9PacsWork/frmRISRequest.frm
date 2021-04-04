VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.1#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRISRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "检查登记"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11506.24
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin DicomObjects.DicomViewer dcmTmpView 
      Height          =   375
      Left            =   5160
      TabIndex        =   93
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
      _Version        =   262147
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   35
      BackColor       =   -2147483639
   End
   Begin VB.CommandButton cmdPetitionCapture 
      Caption         =   "申请单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   35
      ToolTipText     =   "保存(F2)"
      Top             =   7170
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   3525
      Left            =   135
      TabIndex        =   49
      Top             =   375
      Width           =   11235
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   360
         Left            =   825
         TabIndex        =   0
         ToolTipText     =   """数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"""
         Top             =   225
         Width           =   2860
         _ExtentX        =   5054
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmRISRequest.frx":0000
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         ShowPropertySet =   -1  'True
         DefaultCardType =   "就诊卡"
         IDkindBorderStyle=   1
         IDKindWidth     =   600
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.CommandButton cmdSelectPinyinName 
         Caption         =   "…"
         Height          =   350
         Left            =   3450
         TabIndex        =   92
         Top             =   675
         Width           =   260
      End
      Begin VB.TextBox txt送检医生 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4995
         TabIndex        =   10
         Top             =   1485
         Width           =   2280
      End
      Begin VB.TextBox txt送检单位 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         TabIndex        =   9
         Top             =   1485
         Width           =   2280
      End
      Begin VB.ComboBox cbo医生2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         TabIndex        =   13
         Text            =   "cbo医生2"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.ComboBox cbo技师二 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3120
         Width           =   2325
      End
      Begin VB.ComboBox cbo执行科室 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "frmRISRequest.frx":00E8
         Left            =   4995
         List            =   "frmRISRequest.frx":00EA
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1095
         Width           =   2280
      End
      Begin VB.TextBox txt年龄 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         MaxLength       =   50
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
      Begin VB.ComboBox cboAge 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRISRequest.frx":00EC
         Left            =   10215
         List            =   "frmRISRequest.frx":00F9
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   915
      End
      Begin VB.ComboBox cbo技师一 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2700
         Width           =   2325
      End
      Begin VB.TextBox txt医嘱内容 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1410
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1905
         Width           =   5595
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   375
         Left            =   7020
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1905
         Width           =   260
      End
      Begin VB.TextBox Txt部位方法 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1145
         Left            =   1395
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   2280
         Width           =   5895
      End
      Begin VB.ComboBox cbo医生1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1485
         Width           =   2325
      End
      Begin VB.ComboBox cbo开单科室 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRISRequest.frx":0109
         Left            =   8820
         List            =   "frmRISRequest.frx":010B
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1080
         Width           =   2325
      End
      Begin VB.ComboBox cbo婚姻 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox Txt身份证号 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4995
         TabIndex        =   5
         Top             =   690
         Width           =   2280
      End
      Begin VB.TextBox Txt电话 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8820
         TabIndex        =   6
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox Txt英文名 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   675
         Width           =   2020
      End
      Begin VB.ComboBox cbo性别 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmRISRequest.frx":010D
         Left            =   4995
         List            =   "frmRISRequest.frx":011A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   2280
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   330
         Index           =   0
         Left            =   8820
         TabIndex        =   15
         Top             =   1905
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   3080195
         CurrentDate     =   38222
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   330
         Index           =   1
         Left            =   8820
         TabIndex        =   16
         Top             =   2310
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   3080195
         CurrentDate     =   38222
      End
      Begin VB.Label lbl送检单位 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "送检单位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   91
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label lbl送检医生 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "送检医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   90
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查技师二"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   87
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label lab执行科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3765
         TabIndex        =   86
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年   龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7335
         TabIndex        =   65
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查技师一"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   64
         Top             =   2730
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  申请时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7335
         TabIndex        =   63
         Top             =   1935
         Width           =   1440
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  检查时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   7335
         TabIndex        =   62
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Lbl部位方法 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "部位方法"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   61
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label lbl医嘱内容 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   135
         TabIndex        =   60
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  申请医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   57
         Top             =   1530
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  申请科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   56
         Top             =   1125
         Width           =   1440
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   55
         Top             =   1095
         Width           =   1155
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电   话"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   54
         Top             =   705
         Width           =   1425
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3765
         TabIndex        =   53
         Top             =   705
         Width           =   1140
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "英 文 名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   52
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性   别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3765
         TabIndex        =   51
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   90
         TabIndex        =   50
         Top             =   255
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   135
      TabIndex        =   66
      Top             =   3780
      Width           =   11235
      Begin VB.ComboBox cboRoom 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4995
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   195
         Width           =   2280
      End
      Begin VB.ComboBox cboDevice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   180
         Width           =   2310
      End
      Begin VB.TextBox txt检查号 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         MaxLength       =   64
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查设备"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   7515
         TabIndex        =   69
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label lblRoom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执 行 间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3750
         TabIndex        =   68
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "检 查 号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -180
         TabIndex        =   67
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   150
      TabIndex        =   41
      Top             =   0
      Width           =   11190
      Begin VB.CheckBox chk紧急 
         Caption         =   "紧急检查"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   9570
         TabIndex        =   48
         Top             =   75
         Width           =   1545
      End
      Begin VB.TextBox txtBed 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7350
         TabIndex        =   45
         Top             =   105
         Width           =   1890
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   105
         Width           =   1935
      End
      Begin VB.TextBox txtPatientDept 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1350
         TabIndex        =   42
         Top             =   105
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标 识 号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3345
         TabIndex        =   47
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   46
         Top             =   90
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   43
         Top             =   90
         Width           =   1140
      End
   End
   Begin VB.CheckBox chkRoom 
      Caption         =   "执行间情况(&R)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7170
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9015
      TabIndex        =   37
      ToolTipText     =   "保存(F2)"
      Top             =   7170
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10245
      TabIndex        =   38
      Top             =   7170
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwRoom 
      Height          =   1410
      Left            =   1470
      TabIndex        =   39
      Top             =   7710
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2487
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame frm其他信息 
      Height          =   2730
      Left            =   135
      TabIndex        =   70
      Top             =   4335
      Width           =   11235
      Begin VB.ComboBox cbo付款方式 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2265
         Width           =   1905
      End
      Begin VB.ComboBox cbo费别 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2280
         Width           =   1905
      End
      Begin VB.ComboBox cbo造影剂 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRISRequest.frx":012C
         Left            =   1335
         List            =   "frmRISRequest.frx":012E
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1875
         Width           =   1905
      End
      Begin VB.TextBox Txt造影用量 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4830
         TabIndex        =   31
         Top             =   1890
         Width           =   1890
      End
      Begin VB.TextBox Txt造影浓度 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8580
         TabIndex        =   32
         Top             =   1860
         Width           =   2190
      End
      Begin VB.TextBox txt附加主述 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1335
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1455
         Width           =   9435
      End
      Begin VB.TextBox Txt联系地址 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1335
         TabIndex        =   28
         Top             =   1035
         Width           =   9435
      End
      Begin VB.TextBox Txt邮编 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8565
         TabIndex        =   27
         Top             =   630
         Width           =   2205
      End
      Begin VB.ComboBox cbo职业 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   615
         Width           =   1830
      End
      Begin VB.ComboBox cbo民族 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   615
         Width           =   2140
      End
      Begin VB.TextBox Txt体重 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8565
         TabIndex        =   24
         Top             =   195
         Width           =   2205
      End
      Begin VB.TextBox Txt身高 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4770
         TabIndex        =   23
         Top             =   210
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtp出生日期 
         Height          =   294
         Left            =   1316
         TabIndex        =   22
         Top             =   238
         Width           =   2212
         _ExtentX        =   3916
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   3080195
         CurrentDate     =   38222
      End
      Begin VB.Label Label31 
         Caption         =   "KG"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10830
         TabIndex        =   89
         Top             =   255
         Width           =   240
      End
      Begin VB.Label Label24 
         Caption         =   "CM"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6645
         TabIndex        =   88
         Top             =   255
         Width           =   315
      End
      Begin VB.Label lblCash 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   8565
         TabIndex        =   85
         Top             =   2310
         Width           =   2160
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费    用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7215
         TabIndex        =   84
         Top             =   2295
         Width           =   1170
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   83
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费    别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   82
         Top             =   2295
         Width           =   1170
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "造 影 剂"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   75
         TabIndex        =   81
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "造影剂用量"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3270
         TabIndex        =   80
         Top             =   1905
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "造影剂浓度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6930
         TabIndex        =   79
         Top             =   1890
         Width           =   1560
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "附加主述"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   78
         Top             =   1470
         Width           =   1140
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   77
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮  编"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   76
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职  业"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3810
         TabIndex        =   75
         Top             =   645
         Width           =   870
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民    族"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   74
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体  重"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   73
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身  高"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3810
         TabIndex        =   72
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   71
         Top             =   270
         Width           =   1140
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   4440
      Top             =   7200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lbl执行间 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执 行 间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   40
      Top             =   7710
      Width           =   1395
   End
End
Attribute VB_Name = "frmRISRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'模块变量----以从值从外部传入
Public mstrPrivs As String          '调用者的权限
Public mlngModul As Long            '由谁调用
Public mlngAdviceID As Long         '医嘱ID
Public mstrPatientName As String
Public mlngSendNo As Long           '发送号
Public mintEditMode As Integer      '0－登记、1－登记后修改、2－报到、3－报到后修改
Public mlngCurDeptId As Long        '当前科室ID
Public mstrCur科室 As String        '科室编码和名称
Public mstrTechnicRoom As String    '当前报到执行间
Public mlngResultState As Long      '保存或取消,0-失败， 1-登记成功，2-报到成功，3-修改成功，4-处理成功（用于连续登记时返回）
'Public mlngQueueWay As Long        '排队方式
Public mblnIsAllDepartment As Boolean '是否所有部门

Public mintImgCount As Integer      '已扫描图像数量
Public mblnIsRelationImage As Boolean '判断是否进行了图像关联处理

Private frmPetitionCap As frmPetitionCapture      '扫描申请单窗体对象

'公共模块变量------以下值从参数表中取得
Private mblnUseModalityPretext As Boolean '检查号使用影像类别前缀
Private mblnUseAdviceNo As Boolean  '使用医嘱ID
Private mblnUsePatientNo As Boolean '使用患者号
Private mblnChangeNo As Boolean     '手工调整检查号
Private mblnCanOverWrite            '允许检查号重复
Private mblnLike As Boolean, mlngLike As Long    '姓名模糊查找,查找天数
Private mBeforeDays As Integer      '过滤天数
Private mlngTypeSuit As Long        '提前进行的检查，匹配检查图像方式  0-检查号 1-门诊/住院号  2-检查标识号
Private mlngGoOnReg As Long         '连续登记 0-非连续,1-连续
Private mlngUnicode As Long         '患者检查号保持不变,1-保持检查号不变；0-检查号流水递增
Private mlngUnicodeType As Long     '检查号保持不变类别,不变类别 0-按类别不变 1-按科室不变;
Private mlngBuildType As Long       '检查号生成方式,0-按类别递增 1-按科室递增
Private mblnRegToCheck As Boolean   '登记直接检查
Private mblnNoshowReagent As Boolean '不显示造影剂
Private mblnNoshowAddons As Boolean '不显示附加主述
Private mblnInputOutInfo As Boolean  '录入外院信息
Private mintCheckInMode As Integer  '登记模式 1--精简模式，2--正常模式
Private mblnUseReferencePatient     '使用关联病人模式
Private mintCapital As Integer      '拼音名大小写
Private mblnUseSplitter As Boolean  '拼音名分隔符
Private mblnAllPatientIsOutside As Boolean '所有登记病人标记为外来
Private mlngVideoStationMoneyExeModle As Long   '影像采集的费用执行模式 0-报到时执行，1-检查时执行，2-报告时执行
Private mlngPacsStationMoneyExeModle As Long    '影像医技的费用执行模式 0-报到时执行，1-报告时执行
Private mstrRoom As String
Private mblnNameColColorCfg As Boolean  '是否根据病人类型设置姓名颜色
Private mblnOrdinaryNameColColorCfg As Boolean '缺省类型的病人是否根据病人类型设置姓名颜色
Private mstrDefaultPatientType As String '缺省的病人类型

'公共模块变量------以下运行中赋值
Public mintSourceType As Integer   '病人来源 1-门诊 2-住院 3-外来 4-体检
Private mlngPatiId As Long, mlngPageID As Long  '病人ID,主页ID
Private mstrItemType As String      '影像类别
Public mlngClinicID As Long        '诊疗项目ID
'Private mstrItemIDS As String       '收费细目ID
Private mInputType As Integer       '提取病人方式　0-就诊卡 1-病人ID 2-住院号 3-门诊号 4-挂号单 5-收费单据号 6-姓名 7-医保号 8-身份证号 9-IC卡号
Private mstrExtData  As String      '登记的申请项目部位及方法 检查="部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
Private mstrAppend As String        '检查="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
Private mstrOutNo As String         '门诊号
Private mstrCardNo As String        '就诊卡号
Private mstrCardPass As String      '卡验证码
Private mstrChargeNo As String      '收费单据
Private mstrRegNo As String         '挂号单据
Private arrSQL() As Variant
Private mstrNextCheckNo As String     '记录本次获取到的下一个检查号

Private mobjSquareCard As Object    '一卡通，卡结算部件
Private oneSquardCard As TSquardCard

Private mlngBaby As Long            '是否婴儿，0--不是婴儿，1-9表示婴儿序号

Private mblnIsOutSideHosp As Boolean     '是否是外院科室
Private mblnIsPetitionScan As Boolean    '是否启用申请单扫描
Private mblnIsSamePatient As Boolean     '是否存在相同病人
Private mblnUsePacsQueue As Boolean          '是否启用排队叫号
Private mlngPacsQueueNoWay As Long          '排队叫号号码生成方式
Private mblnSelectDefRoom As Boolean        '是否选择默认执行间



Private mblnExamineDoctorVerify As Boolean '是否技师确认
Private mstrExamineDoctorName As String    '技师名字
Private mstrExamineDoctorFst As String     '检查技师一
Private mstrExamineDoctorSed As String     '检查技师二

Private mlngInsureCheckType As Long         '医保对码检查类型 0-不检查， 1-仅提示，2-禁止
Private mobjInsure As Object

Private mfrmParent As Form          '父窗体
Private mobjPublicPatient As Object

Private mstrOldAdvice As String  '登记后修改界面的原医嘱
Public Event HaveRegist() '完成登记

Public Function zlShowMe(frmParent As Form, ByVal strDefaultPatientType As String, ByVal blnBigFont As Boolean, Optional ByVal blnIsAllDepartment As Boolean = False, _
    Optional ByVal lngCopyAdviceId As Long, Optional ByVal lngCopySendNo As Long) As Boolean
    Set mfrmParent = frmParent
    
    mlngResultState = 0
    mblnIsRelationImage = False
    mstrTechnicRoom = ""
    mstrRoom = ""
    mstrDefaultPatientType = strDefaultPatientType
    mblnIsAllDepartment = blnIsAllDepartment
    
    Set mobjPublicPatient = VBA.Interaction.GetObject("", "zlPublicPatient.clsPublicPatient")
    If mobjPublicPatient Is Nothing Then Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
    
    If Not mobjPublicPatient Is Nothing Then Call mobjPublicPatient.zlInitCommon(gcnOracle, glngSys)
    
    Call ConfigExeDepartment(mblnIsAllDepartment)
    
    Call InitParameter
    Call InitEdit(False)  '初始化界面数据
    
    Call SetFontSize(blnBigFont)
    
    
    '读取病人信息
    If mintEditMode <> 0 And mlngAdviceID <> 0 Then
        If Not RefreshPatiInfor(mintEditMode = 2) Then Exit Function
    End If
    Call RefreshRoom
    '复制传递的登记信息
    If lngCopyAdviceId <> 0 And lngCopySendNo <> 0 Then Call CopyCheck(lngCopyAdviceId, lngCopySendNo)
    
    mstrOldAdvice = Trim(txt医嘱内容.Text) & Trim(Txt部位方法.Text)
    
    If (mblnUseAdviceNo Or mblnUsePatientNo) And (mblnRegToCheck And mintEditMode = 0) Then txt检查号.Enabled = False
    Me.Show 1, mfrmParent
End Function



Private Sub SetFontSize(ByVal blnIsBigFont As Boolean)
    Dim objControl As Object
    Dim lngLabFontSize As Long
    Dim lngTxtFontSize As Long
    
    lngLabFontSize = IIf(blnIsBigFont, 14, 12)
    lngTxtFontSize = IIf(blnIsBigFont, 12, 10.5)
    
    For Each objControl In Me.Controls
        If TypeName(objControl) = "Label" Then
            If objControl.Name <> "Label24" And objControl.Name <> "Label31" Then
                objControl.Font.Size = lngLabFontSize
            End If
        ElseIf TypeName(objControl) = "CommandBars" Then
            'objControl.Options.Font.Size
        Else
            If TypeName(objControl) <> "DicomViewer" Then objControl.Font.Size = lngTxtFontSize
        End If
    Next
    
    lblCash.FontSize = lngTxtFontSize
    chk紧急.FontSize = lngLabFontSize
End Sub




Private Sub SaveAdviceData()
'------------------------------------------------
'功能：保存医嘱
'参数： 无
'返回：无
'------------------------------------------------
    Dim str检查时间 As String, str申请时间 As String, curDate As String
    Dim strNO As String, lngAdviceID As Long, lngSendNO As Long
    Dim IntSeq As Integer   '病人医嘱记录.序号
    Dim str部位 As String, str方法 As String
    Dim i As Integer, j As Integer, strTmp方法 As String, str部位方法 As String
    Dim lng开嘱科室ID As Long, lng病人ID As Long, strDoctor As String
    Dim str执行科室ID As String, lngTmpID As Long, arrAppend
    Dim rsTemp As ADODB.Recordset
    Dim lngMasSeq As Long   '病人医嘱发送.记录序号，主医嘱中的
    Dim lngSonSeq As Long   '病人医嘱发送.记录序号，附加医嘱中的，要递增
    

    On Error GoTo errHand
    
    curDate = zlStr.To_Date(zlDatabase.Currentdate)
    str检查时间 = zlStr.To_Date(dtp(1))
    str申请时间 = zlStr.To_Date(dtp(0))
    
    str部位方法 = Split(mstrExtData, Chr(9))(0)
    lng开嘱科室ID = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    strDoctor = IIf(Me.cbo医生1.Visible, zlStr.NeedName(Me.cbo医生1.Text), zlStr.NeedName(Me.cbo医生2.Text))
    str执行科室ID = mlngCurDeptId
    
    '新病人，要添加病人信息
    If mlngPatiId <= 0 Then
        '提取新的病人ID
        mlngPatiId = zlDatabase.GetNextNo(1)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_挂号病人病案_INSERT(1," & mlngPatiId & ",''," & _
            "'',''," & _
            "'" & Trim(PatiIdentify.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, "") & "'," & _
            "'" & zlStr.NeedName(cbo费别.Text) & "','" & zlStr.NeedName(cbo付款方式.Text) & "'," & _
            "'','" & zlStr.NeedName(cbo民族.Text) & "','" & zlStr.NeedName(cbo婚姻.Text) & "'," & _
            "'" & zlStr.NeedName(cbo职业.Text) & "','" & zlCommFun.ToVarchar(Txt身份证号, 18) & "',''," & Val(Label22.tag) & ",'','','" & zlCommFun.ToVarchar(Txt联系地址.Text, 50) & _
            "','" & zlCommFun.ToVarchar(Txt电话, 20) & "','" & zlCommFun.ToVarchar(Txt邮编, 6) & "'," & curDate & ",'" & mstrRegNo & "'," & zlStr.To_Date(dtp出生日期.value) & ",NULL)"
    End If
    
    '保存医嘱并发送
    '收费单据为空，提取下一个收费单据号
    If mstrChargeNo = "" Then
        strNO = zlDatabase.GetNextNo(IIf(mintSourceType <> 2, 13, 14)) '门诊取收费单据号,住院取记帐单据号
        lngMasSeq = 1
        lngSonSeq = 1
    Else    '有收费单据号
        strNO = mstrChargeNo
        '已收费单据,根据NO提取当前最大序号+1开始,用于病人医嘱发送,附医嘱的序号需根据再次递加
        gstrSQL = "Select Max(记录序号) as 序号 From 病人医嘱发送 Where No=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前NO最大序号", CStr(mstrChargeNo))
        If rsTemp.EOF Then
            lngMasSeq = 1
            lngSonSeq = 1
        Else
            lngMasSeq = NVL(rsTemp!序号, 0) + 1
            lngSonSeq = lngMasSeq
        End If
    End If
    
    lngAdviceID = zlDatabase.GetNextId("病人医嘱记录")
    lngSendNO = zlDatabase.GetNextNo(10) '医嘱发送号
    
    '插入外院信息，主要是送检单位和送检医生
    If mblnInputOutInfo Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlngPatiId & ",'送检单位','" & Trim(NVL(txt送检单位.Text)) & "'," & lngAdviceID & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlngPatiId & ",'送检医生','" & Trim(NVL(txt送检医生.Text)) & "'," & lngAdviceID & ")"
    End If
    
    '插入主医嘱
    IntSeq = IntSeq + 1     '病人医嘱记录.序号，递增
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
                    IntSeq & "," & mintSourceType & "," & mlngPatiId & "," & IIf(mintSourceType = 2, mlngPageID, "NULL") & "," & mlngBaby & _
                    ",1,1,'D'," & mlngClinicID & ",NULL,NULL,NULL,1," & _
                    "'" & Me.txt医嘱内容 & "," & Decode(Txt部位方法.tag, 1, "床旁", 2, "术中", "常规") & "执行:" & _
                    get部位方法(mstrExtData) & "',Null,Null,'一次性',NULL,NULL,NULL,NULL,2," & _
                    str执行科室ID & ",3," & chk紧急.value & "," & str检查时间 & "," & str检查时间 & "," & _
                    IIf(Val(Me.txtPatientDept.tag) = 0, lng开嘱科室ID, Val(Me.txtPatientDept.tag)) & "," & lng开嘱科室ID & _
                    ",'" & strDoctor & "'," & curDate & ",'" & mstrRegNo & "',Null,Null," & Txt部位方法.tag & ",NULL,NULL,'" & UserInfo.姓名 & "')"
    
    '循环部位方法，插入附加医嘱
    For i = 0 To UBound(Split(str部位方法, "|")) '部位1;方法1,方法2,方法3|部位n;方法1,方法2,方法3---
        str部位 = Split(Split(str部位方法, "|")(i), ";")(0)
        strTmp方法 = Split(Split(str部位方法, "|")(i), ";")(1)
        For j = 0 To UBound(Split(strTmp方法, ","))
            IntSeq = IntSeq + 1     '病人医嘱记录.序号，递增
            str方法 = Split(strTmp方法, ",")(j)
            lngTmpID = zlDatabase.GetNextId("病人医嘱记录")
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                 IntSeq & "," & mintSourceType & "," & mlngPatiId & "," & IIf(mintSourceType = 2, mlngPageID, "NULL") & "," & mlngBaby & _
                 ",1,1,'D'," & mlngClinicID & ",NULL,NULL,NULL,1," & _
                 "'" & Replace(Me.txt医嘱内容, "'", "") & "',NULL," & _
                 "'" & str部位 & "','一次性',NULL,NULL,NULL,NULL,2," & _
                 str执行科室ID & ",3," & chk紧急.value & "," & str检查时间 & "," & str检查时间 & "," & _
                 IIf(Val(Me.txtPatientDept.tag) = 0, lng开嘱科室ID, Val(Me.txtPatientDept.tag)) & "," & lng开嘱科室ID & _
                 ",'" & strDoctor & "'," & curDate & ",'" & mstrRegNo & "',Null,'" & str方法 & "'," & Txt部位方法.tag & ",NULL,NULL,'" & UserInfo.姓名 & "')"
            
            '发送附加医嘱
            '有收费单据号的为已计费,无的为未计费
            lngSonSeq = lngSonSeq + 1       '病人医嘱发送.记录序号，附加医嘱中的，要递增
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            '发送医嘱的时候，不填写首次时间和末次时间，报到的时候才填写
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱发送_Insert(" & _
                lngTmpID & "," & lngSendNO & "," & IIf(mintSourceType = 2, 2, 1) & ",'" & strNO & "'," & _
                lngSonSeq & ",1,NULL,NULL," & str申请时间 & ",0," & str执行科室ID & "," & _
                IIf(mstrChargeNo = "", 0, 1) & ",0,Null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Next
    Next
    
    '发送主医嘱
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    '发送医嘱的时候，不填写首次时间和末次时间，报到的时候才填写
    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱发送_Insert(" & _
            lngAdviceID & "," & lngSendNO & "," & IIf(mintSourceType = 2, 2, 1) & ",'" & strNO & "'," & _
            lngMasSeq & ",1,NULL,NULL," & str申请时间 & ",0," & str执行科室ID & "," & _
            IIf(mstrChargeNo = "", 0, 1) & ",1,Null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    
    '插入病人医嘱附件 '     检查="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    If mstrAppend <> "" Then
        arrAppend = Split(mstrAppend, "<Split1>")
        For i = 0 To UBound(arrAppend)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lngAdviceID & _
                ",'" & Split(arrAppend(i), "<Split2>")(0) & "'," & Val(Split(arrAppend(i), "<Split2>")(1)) & "," & _
                i + 1 & "," & zlStr.ZVal(Split(arrAppend(i), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(i), "<Split2>")(3), "'", "''") & "'" & _
                            IIf(i = 0, ",1", "") & ")"
        Next
    End If
    
    '有收费单据号的，设置费用记录和医嘱的关联关系
    If mstrChargeNo <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人费用记录_医嘱('" & strNO & "',1," & lngAdviceID & ")"
    End If
    
    mlngAdviceID = lngAdviceID
    mstrPatientName = Trim(PatiIdentify.Text)
    mlngSendNo = lngSendNO
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ConfigExeDepartment(ByVal blnIsAllDepartment As Boolean)
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strFrom As String
    Dim lngDefaultDeptIndex As Long
    
    lab执行科室.Visible = blnIsAllDepartment
    cbo执行科室.Visible = blnIsAllDepartment
    
    Call cbo执行科室.Clear
    
    If Not blnIsAllDepartment Then Exit Sub
    
    strFrom = "1,2,3"
    strSql = " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
        " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
        " Order by A.编码"
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr("," & strFrom & ","))
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngDefaultDeptIndex = 0
    
    While Not rsData.EOF
        cbo执行科室.AddItem (NVL(rsData!编码) & "-" & NVL(rsData!名称))
        cbo执行科室.ItemData(cbo执行科室.ListCount - 1) = NVL(rsData!ID)
        
        If NVL(rsData!ID) = mlngCurDeptId Then
            lngDefaultDeptIndex = cbo执行科室.ListCount - 1
        End If
        
        rsData.MoveNext
    Wend
    
    If cbo执行科室.ListCount > 0 Then cbo执行科室.ListIndex = lngDefaultDeptIndex
End Sub


Private Sub cboAge_LostFocus()
    If Not CheckOldData(txt年龄, cboAge) Then Exit Sub
    If IsNumeric(txt年龄.Text) Then Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
End Sub


Private Sub cbo技师二_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo技师二.hWnd, zlControl.CboMatchIndex(cbo技师二.hWnd, KeyAscii))
End Sub

Private Sub cbo技师一_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo技师一.hWnd, zlControl.CboMatchIndex(cbo技师一.hWnd, KeyAscii))
End Sub




Private Sub cbo执行科室_Click()
On Error GoTo errHandle
    If cbo执行科室.ListCount <= 0 Then Exit Sub
    
    mlngCurDeptId = cbo执行科室.ItemData(cbo执行科室.ListIndex)
    mstrCur科室 = zlStr.NeedName(cbo执行科室.Text)
    
    txt医嘱内容.Text = ""
    Txt部位方法.Text = ""
    
    Call InitParameter
    Call InitEdit(True)  '初始化界面数据
    Call RefreshRoom
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Txt英文名.Text = Control.Caption
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkRoom_Click()
    If chkRoom.value = 1 Then
        Me.Height = Me.Height + lvwRoom.Height + 300
        InitRoomPati
    Else
        Me.Height = Me.Height - lvwRoom.Height - 300
    End If
End Sub

'从医嘱模块中，复制过来的检查函数
Public Function CheckAdviceInsure(ByVal int险类 As Integer, ByVal bln提醒对码 As Boolean, ByVal lng病人ID As Long, ByVal lng病人性质 As Long, _
   ByVal strIDs1 As String, ByVal strIDs2 As String, ByVal str医嘱内容 As String, Optional ByVal lng病人病区ID As Long) As String
'功能：医保病人下达医嘱时，医嘱录入后，对医嘱涉及的计价项目的保险对码情况进行检查
'参数：strIDs1:药品卫材的收费细目ID字符串（一组医嘱例如：青霉素+葡萄糖）:收费细目ID1,收费细目ID2,···
'      strIDs2 ：其他诊疗项目的诊疗项目ID（一组医嘱例如：输血项目+输血途径）:执行科室字符串 诊疗项目ID1:执行科室1,诊疗项目ID2:执行科室2,···
'      lng病人性质=1门诊，=2住院
'      str医嘱内容：用户提示时显示的医嘱内容
'      bln提醒对码=False 表示当前不继续检查，=True 继续检查
'返回：提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    If mlngInsureCheckType = 0 Or int险类 = 0 Or Not bln提醒对码 Then Exit Function
    If mobjInsure.GetCapability(12, lng病人ID, int险类) Then Exit Function '12:support允许不设置医保项目
    
    
    If strIDs1 = "" And strIDs2 = "" Then Exit Function
    
    If strIDs1 <> "" Then
        If Mid(strIDs1, 1, 1) = "," Then strIDs1 = Mid(strIDs1, 2)
        strSql = "Select Column_Value as 收费项目ID From Table(f_Num2list([1]))"
    End If
    If strIDs2 <> "" Then
        If Mid(strIDs2, 1, 1) = "," Then strIDs2 = Mid(strIDs2, 2)
        If strIDs1 <> "" Then strSql = strSql & " Union All "
        '由于没有加部位等条件，所以要用Distinct
        strSql = strSql & "Select 收费项目ID From (" & _
                "Select Distinct C.收费项目ID,C.适用科室id" & _
                " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                " From 诊疗收费关系 C,Table(f_Num2list2([2])) D Where C.诊疗项目ID=D.c1" & _
                "      And (C.适用科室ID is Null or C.适用科室ID = Nvl(D.c2,[4]) And C.病人来源 = " & IIf(lng病人性质 = 1, 1, 2) & ")" & _
                " ) Where Nvl(适用科室id, 0) = Top"
    End If
    
    strSql = "Select /*+ RULE */ Distinct C.名称,B.收费细目ID" & _
        " From (" & strSql & ") A,保险支付项目 B,收费项目目录 C" & _
        " Where A.收费项目ID=B.收费细目ID(+) And A.收费项目ID=C.ID" & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        " And B.险类(+)=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckAdviceInsure", strIDs1, strIDs2, int险类, lng病人病区ID)
    strSql = "": i = 0
    Do While Not rsTmp.EOF
        If IsNull(rsTmp!收费细目ID) Then
            If i = 8 Then
                strSql = strSql & vbCrLf & "… …"
                Exit Do
            End If
            strSql = strSql & vbCrLf & "●" & rsTmp!名称
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    If strSql <> "" Then
        CheckAdviceInsure = "当前病人是医保病人，但医嘱的以下计价项目没有设置对应的保险项目！" & vbCrLf & vbCrLf & _
            "医嘱内容：" & vbCrLf & str医嘱内容 & vbCrLf & vbCrLf & "计价项目：" & strSql
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub CmdOK_Click()
    Dim l As Long, blnTran As Boolean, rsTmp As New ADODB.Recordset
    Dim rsMother As New ADODB.Recordset
    Dim rsPatiInfo As New ADODB.Recordset
    Dim int记录性质 As Integer     '病人医嘱发送.记录性质，本次医嘱的记录性质，1-收费记录；2-记帐记录
    Dim int门诊记帐 As Integer     '病人医嘱发送.门诊记帐，门诊和住院医生站发送为门诊记帐时填为1,用于区分门诊记帐和住院记帐，其他的都填为空
    Dim str诊疗类别 As String
    Dim lng发送号 As Long
    Dim str单据号 As String
    Dim str医嘱IDs As String
    Dim strMsg As String
    Dim lngCurFromType As Long
    Dim lngMsgResult As Long
    Dim blnInpatientOut As Boolean  '住院患者已经出院，true已出院；false未出院
    
    On Error GoTo errHandle
    
    '检查数据输入是否合法，不合法则退出
    If ValidData = False Then Exit Sub
    
    mstrPatientName = Trim(PatiIdentify.Text)
    
    mstrTechnicRoom = ""
    If cboRoom.Text <> "呼叫时指派" Then mstrTechnicRoom = zlStr.NeedCode(cboRoom.Text)
    
    
    arrSQL = Array()
    
    lngCurFromType = mintSourceType
    If mblnAllPatientIsOutside Then mintSourceType = 3  '所有登记病人标记为外来
    
    '登记 ， 登记作为一个单独的数据库事务来处理
    If mintEditMode = 0 Then
    '        1)医保对码检查
    '        2)保存医嘱（新建病人，新建医嘱，发送医嘱）
        If (lngCurFromType = 1 Or lngCurFromType = 2) And mlngInsureCheckType <> 0 Then
            '只有从门诊或住院开过来的医保病人才进行医保对码检查
            gstrSQL = "select 险类 from 病人信息 Where 病人ID = [1]"
            Set rsPatiInfo = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人险类信息", mlngPatiId)
            
            '医保对码检查
            strMsg = CheckAdviceInsure(Val(NVL(rsPatiInfo!险类)), True, mlngPatiId, mintSourceType, _
                                        "", mlngClinicID & ":" & mlngCurDeptId, "当前项目")
                                        
            If strMsg <> "" Then
                If mlngInsureCheckType = 1 Then '只提示
                    lngMsgResult = MsgBoxD(Me, strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", vbYesNo, "提示信息")
                    If lngMsgResult = vbNo Then Exit Sub
                Else    '禁用
                    MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, "提示信息"
                    Exit Sub
                End If
            End If
        End If
        
        '组织新建患者和保存医嘱的SQL语句，存放到 arrSQL 中，如果是新病人，提取病人ID
        Call SaveAdviceData
        
        If mblnRegToCheck And mintEditMode = 0 Then

            If mblnUseAdviceNo Then
                txt检查号.Text = mlngAdviceID
            ElseIf mblnUsePatientNo Then
                txt检查号.Text = mlngPatiId
            Else
                If CheckNoValidate = False Then
                    If mintSourceType = 3 Then mlngPatiId = 0  '外来病人执行这个，问题103915
                    Exit Sub
                End If
            End If
        End If
        
        '--------------------------执行过程，写入数据
        gcnOracle.BeginTrans
        blnTran = True
        For l = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "检查登记")
        Next
        gcnOracle.CommitTrans
        blnTran = False
        
        '清空SQL语句数组，为后面的登记后直接报到做准备
        arrSQL = Array()
    End If
    
    '登记后修改，才能修改医嘱内容
    '诊疗项目有改动，作废原有的医嘱，重新创建新的医嘱，条件如下
    '1、登记后修改
    '2、且病人来源为外诊
    '3、医嘱内容和部位方法有改变
    If mintEditMode = 1 And mintSourceType = 3 And mstrOldAdvice <> (Trim(txt医嘱内容.Text) & Trim(Txt部位方法.Text)) Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人医嘱记录_作废(" & mlngAdviceID & ")"
        Call SaveAdviceData
    End If

    '修改影像表的病人信息条件如下
    '1、不是登记，需要修改病人的信息，外诊病人的信息比较多
    '实际条件是：登记后修改；报到；报到后修改
    If mintEditMode <> 0 Then
        '修改病人信息的存储过程中，只有外诊病人可以修改姓名性别年龄等基本信息；门诊住院体检病人，只能修改联系信息
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_影像病人信息_修改(" & mintSourceType & "," & mlngAdviceID & "," & mlngPatiId & "," & _
            "'" & Trim(PatiIdentify.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & txt年龄.Text & cboAge.Text & "'," & _
            "'" & zlStr.NeedName(cbo费别.Text) & "','" & zlStr.NeedName(cbo付款方式.Text) & "','" & zlStr.NeedName(cbo民族.Text) & "'," & _
            "'" & zlStr.NeedName(cbo婚姻.Text) & "','" & zlStr.NeedName(cbo职业.Text) & "','" & zlCommFun.ToVarchar(Txt身份证号, 18) & "'," & _
            "'" & zlCommFun.ToVarchar(Txt联系地址.Text, 50) & "','" & zlCommFun.ToVarchar(Txt电话, 20) & "','" & zlCommFun.ToVarchar(Txt邮编, 6) & _
            "'," & zlStr.To_Date(CDate(dtp出生日期.value)) & "," & mlngPageID & "," & mlngBaby & ")"
    End If
    
    '1、报到
    '2、登记，且登记后直接检查。其实也就是报到
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        '处理检查设备
        If cboRoom.ListCount > 0 Then   '如果有执行间
            If zlStr.NeedName(cboRoom.list(cboRoom.ListIndex)) = "" Then '执行间未对应检查设备 , 检查设备由影像类别确定
                InitDevice mstrItemType
            End If
        Else                          '无执行间, 检查设备由影像类别确定
            InitDevice mstrItemType
        End If
        
        '检查费用以及一卡通的处理
        '业务逻辑是：
        '1、总体逻辑没有收费的不能报到，但是如果有“未缴费报到”权限的，可以在不使用一卡通流程的情况下报到。
        '   在刷新信息的时候已经控制报到的确定按钮。
        '2、对公共基础参数的支持：
        '       参数号28--门诊一卡通，消费减少剩余款额时是否需要验证
        '       参数号81--执行后自动审核
        '       参数号163--门诊一卡通，项目执行前必须先收费或先记帐审核
        '3、先处理需要一卡通消费确认的，条件是以下之一
        '       （1）记录性质=1
        '       （2）执行后自动审核=False，记录性质=2，且 “来源<>住院”  或者 “来源=住院，门诊记帐”。
        '   如果一卡通消费确认成功，则可以报到。如果一卡通消费确认不成功，就算有“未缴费报到”权限，也不能报到。
        '4、再处理一卡通费用减少验证的，只处理记账的，条件是：
        '       （1）记录性质=2，执行后自动审核=True
        '       （2）有未审核费用
        gstrSQL = "Select A.记录性质,A.门诊记帐,A.发送号,A.NO,B.诊疗类别 from 病人医嘱发送 A,病人医嘱记录 B  where A.医嘱ID=B.ID and  B.ID =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS报到查找记录性质", mlngAdviceID)
        If rsTmp.EOF = False Then
            int记录性质 = NVL(rsTmp!记录性质, 0)
            int门诊记帐 = NVL(rsTmp!门诊记帐, 0)
            str诊疗类别 = NVL(rsTmp!诊疗类别)
            lng发送号 = rsTmp!发送号
            str单据号 = NVL(rsTmp!NO)
        End If
        
        'int记录性质  病人医嘱发送.记录性质，本次医嘱的记录性质，1-收费记录；2-记帐记录
        
        '判断患者是否有费用的前提：门诊收费或者门诊记账或者已出院患者
        '1、是收费记录
        '2、或者是记账记录且执行后不审核且（不是住院患者或是住院患者门诊记账）
        '3、已出院患者
        
        If mintSourceType = 2 Then blnInpatientOut = Not bln病人在院(mlngPatiId, mlngPageID)
        
        If int记录性质 = 1 _
            Or (gbln执行后审核 = False And int记录性质 = 2 And (mintSourceType <> 2 Or (mintSourceType = 2 And int门诊记帐 = 1))) _
            Or (mintSourceType = 2 And blnInpatientOut = True) Then
            
            '判断当前的执行医嘱是否已收费或记帐划价单是否已审核
            If Not ItemHaveCash(mintSourceType, False, mlngAdviceID, 0, lng发送号, str诊疗类别, str单据号, int记录性质, _
                int门诊记帐, 0) Then
                
                '先判断是否已经出院的住院患者
                If (mintSourceType = 2 And blnInpatientOut = True) Then
                    '提示患者已经出院，有未收费记录，不能报到
                    MsgBoxD Me, "该患者已经出院结算，本次检查不能报到，否则将产生额外费用。", vbOKOnly, "提示信息"
                    Exit Sub
                End If
                
                If gbln执行前先结算 Then
                    '门诊一卡通,项目执行前必须先收费或先记帐审核,不传单据号，根据医嘱ID读取所有未收费单据或未审核的记帐单
                    '读取医嘱ID串
                    str医嘱IDs = mlngAdviceID
                    gstrSQL = "Select Id  from 病人医嘱记录 where 相关ID = [1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医嘱ID串", mlngAdviceID)
                    While rsTmp.EOF = False
                        str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
                        rsTmp.MoveNext
                    Wend
                    
                    If mobjSquareCard.zlSquareAffirm(Me, mlngModul, mstrPrivs, mlngPatiId, 0, False, , , str医嘱IDs) = False Then
                        MsgBoxD Me, "缴费不成功，该病人还存在未收费的费用，无法报到，请检查。", vbOKOnly, "缴费失败"
                        Exit Sub
                    End If
                Else
                    '如果有“未缴费报到”权限，则提示是否确认未收费可以报到？
                    If CheckPopedom(mstrPrivs, "未缴费报到") Then
                        If MsgBoxD(Me, "该病人还存在未收费的费用，是否继续报到？", vbYesNo, "提示信息") = vbNo Then
                            Exit Sub
                        End If
                    Else
                        MsgBoxD Me, "该病人还存在未收费的费用，请检查。", vbOKOnly, "提示信息"
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        
        If int记录性质 = 2 Then
            '取出病人当前划价费用（当执行后自动审核划价单据有效时）
            Dim curMoney As Currency, str类别 As String, str类别名 As String
            
            curMoney = GetAdviceMoney(mlngAdviceID, mintSourceType, str类别, str类别名)
            '当费用不为0时，检查是否一卡通刷卡，是否需要记账报警
            If curMoney <> 0 Then
                '记账报警
                If Not FinishBillingWarn(Me, "", mlngPatiId, mlngPageID, Val(lblCash.tag), curMoney, str类别, str类别名) Then
                    Exit Sub
                End If
                
                '问题：34856
                '门诊一卡通消费身份验证
                '参数28--门诊一卡通消费减少剩余款额时是否需要验证
                '参数81--执行后自动审核
                If curMoney > 0 And mintSourceType = 1 Then
                    If Not zlPatiIdentify(mlngModul, Me, mlngPatiId, curMoney) Then Exit Sub
                End If
            End If
        End If
            
        
        '开始检查
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_影像检查_BEGIN('" & mstrTechnicRoom & "','" & txt检查号.Text & "'," & mlngAdviceID & "," & mlngSendNo & ",'" & mstrItemType & "','" & _
            Trim(Me.PatiIdentify.Text) & "','" & Trim(Txt英文名.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & _
            txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, "") & "'," & zlStr.To_Date(dtp出生日期.value) & ",'" & zlCommFun.ToVarchar(Txt身高, 16) & "','" & _
            zlCommFun.ToVarchar(Txt体重, 16) & "',Null,Null,'" & zlStr.NeedCode(cboDevice.Text) & "','" & zlStr.NeedName(cbo技师一.Text) & "','" & zlStr.NeedName(cbo技师二.Text) & "','" & txt附加主述.Text & "'," & zlStr.To_Date(CDate(dtp(1).value)) & "," & mlngCurDeptId & ",Null)"
        
        '设置影像检查记录--执行过程为-已报到，报到时处理记账的费用
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_影像检查_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
        
        '报到时执行费用
        If (mlngModul = G_LNG_VIDEOSTATION_MODULE And mlngVideoStationMoneyExeModle = 0) Or _
           (mlngModul = G_LNG_PACSSTATION_MODULE And mlngPacsStationMoneyExeModle = 0) Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_影像费用执行(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
        End If
        
        '填写服用造影剂
        If Trim(cbo造影剂.Text) <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_服用造影剂_INSERT(" & mlngAdviceID & ",'" & zlCommFun.ToVarchar(Trim(cbo造影剂.Text), 30) & "','" & zlCommFun.ToVarchar(Txt造影用量.Text, 30) & "','" & zlCommFun.ToVarchar(Txt造影浓度.Text, 30) & "')"
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_服用造影剂_DELETE(" & mlngAdviceID & ")"
        End If
        
        '发送安排
        If GetDeptPara(mlngCurDeptId, "报道时自动发送WorkList") = "1" And mlngModul = 1290 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_影像检查记录_发送安排(" & mlngAdviceID & "," & mlngSendNo & ",1," & "'" & zlStr.NeedName(cbo技师一.Text) & "','" & zlStr.NeedName(cbo技师二.Text) & "','" & zlStr.NeedCode(cboRoom.Text) & "')"
        End If
    End If  '报到的if
    
    
    '报到后修改
    If mintEditMode = 3 Then
         '处理检查设备
        If cboRoom.ListCount > 0 Then   '如果有执行间
            If zlStr.NeedName(cboRoom.list(cboRoom.ListIndex)) = "" Then '执行间未对应检查设备 , 检查设备由影像类别确定
                InitDevice mstrItemType
            End If
        Else                          '无执行间, 检查设备由影像类别确定
            InitDevice mstrItemType
        End If
        
        '修改病人信息
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_影像检查记录_UPDATE(" & mlngAdviceID & ", " & mlngSendNo & ",'" & txt检查号.Text & "','" & _
            Trim(Me.PatiIdentify.Text) & "','" & Trim(Txt英文名.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & _
            txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, "") & "'," & zlStr.To_Date(dtp出生日期.value) & ",'" & zlCommFun.ToVarchar(Txt身高, 16) & "','" & _
            zlCommFun.ToVarchar(Txt体重, 16) & "','" & zlStr.NeedCode(cboDevice.Text) & "','" & zlStr.NeedName(cbo技师一.Text) & "','" & zlStr.NeedName(cbo技师二.Text) & "','" & txt附加主述.Text & "','" & zlStr.NeedCode(cboRoom.Text) & "'," & zlStr.To_Date(dtp(1).value) & ",Null)"

        '填写服用造影剂
        If Trim(cbo造影剂.Text) <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_服用造影剂_INSERT(" & mlngAdviceID & ",'" & zlCommFun.ToVarchar(Trim(cbo造影剂.Text), 30) & "','" & zlCommFun.ToVarchar(Txt造影用量.Text, 30) & "','" & zlCommFun.ToVarchar(Txt造影浓度.Text, 30) & "')"
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_服用造影剂_DELETE(" & mlngAdviceID & ")"
        End If
    End If
    
    '执行数据写入前，先判断检查号是否重复,条件是
    '1、登记，且登记后直接报到，相当于报到
    '2、或者报到
    '3、检查号的tag不等于检查号（为什么要这个条件？）
    If mintEditMode = 2 Or txt检查号.tag <> txt检查号.Text Then
        If CheckNoValidate = False Then
            Exit Sub
        End If
    End If
        
    '--------------------------执行过程，写入数据
    gcnOracle.BeginTrans
    blnTran = True
    For l = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "写入数据")
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    '对病人信息同步,更新专业版PACS的病人信息，条件如下
    '1、使用专业版PACS
    '2、且报到后修改
    '3、且是“影像医技工作站”
    If mintEditMode = 3 And gblnUseXinWangView And mlngModul = G_LNG_PACSSTATION_MODULE Then
        Call XWStudyUpdate(mlngAdviceID, Me.PatiIdentify.Text, zlStr.NeedName(cbo性别.Text), txt年龄.Text)
    End If
    
    '报到的后续处理：回填检查号，匹配图像
    '1、报到
    '2、或者登记，且登记后直接检查，实际上就是报到
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        '查找提前进行的检查，按照规则匹配检查和图像
        gstrSQL = "Select A.检查UID As ID From 影像临时记录 a Where a.检查号=[1] And a.影像类别=[2]"
        Select Case mlngTypeSuit
            Case 0 '检查号
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt检查号.Text, mstrItemType)
            Case 1 '门诊/住院号
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.txtID.Text, mstrItemType)
            Case 2 '检查标识号（医嘱ID）
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str(mlngAdviceID), mstrItemType)
        End Select
        
        '找到匹配的临时图像记录，则将图像和检查自动匹配
        If rsTmp.RecordCount = 1 Then
            gstrSQL = "ZL_影像检查_SET(" & mlngAdviceID & "," & mlngSendNo & ",'" & rsTmp("ID") & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "提前检查匹配"
            
            mblnIsRelationImage = True
        End If
    End If
    
   '保存申请单图像   释放 窗口
   If mintEditMode = 0 Then
        If Not frmPetitionCap Is Nothing And dcmTmpView.Images.Count > 0 Then
            Call frmPetitionCap.subSaveImage(, mlngAdviceID, dcmTmpView)
            '卸载扫描申请单窗体对象
            Set frmPetitionCap = Nothing
            '保存申请单图像后，将其清空，避免连续申请时下一条检查还有图片
            dcmTmpView.Images.Clear
        End If
   End If

    'mlngResultState  '保存或取消,0-失败， 1-登记成功，2-报到成功，3-修改成功，4-处理成功（用于连续登记时返回）
    '设置返回状态
    Select Case mintEditMode
        Case 0
            If mblnRegToCheck Then
                mlngResultState = 2
            Else
                mlngResultState = 1
            End If
        Case 1, 3
            mlngResultState = 3
        Case 2
            mlngResultState = 2
    End Select
           
    
    '如果是连续登记，而且处于登记状态，则不关闭窗口。
    If mlngGoOnReg = 1 And mintEditMode = 0 Then
        RaiseEvent HaveRegist
        Call InitMvar '初始化模块变量
        Call ClearFaceData
        'InitEdit '初始化界面 '屏蔽次语句，不需要每次重新加载combobox数据
        Me.PatiIdentify.SetFocus
    Else
        '如果处于报到状态,或者登记后直接报到，则检查是否提示关联病人
        If (mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0)) And mblnUseReferencePatient = True Then
            frmReferencePatient.zlShowMe mlngAdviceID, Trim(PatiIdentify.Text), Me, False, mlngCurDeptId
        End If
        
        Unload Me
        
    End If
    
    Exit Sub
errHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Unload Me
End Sub


Private Sub cmdPetitionCapture_Click()
On Error GoTo errHand
    
    If frmPetitionCap Is Nothing Then
        Set frmPetitionCap = New frmPetitionCapture
    End If


     '打开扫描申请单窗口
    Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                            mlngCurDeptId, _
                                            NVL(Mid(cbo开单科室.Text, InStr(cbo开单科室.Text, "-") + 1, Len(cbo开单科室.Text))), _
                                            NVL(Trim(PatiIdentify.Text)), _
                                            NVL(txt年龄.Text), _
                                            NVL(Mid(cbo性别.Text, InStr(cbo性别.Text, "-") + 1, Len(cbo性别.Text))), _
                                            NVL(txt医嘱内容.Text), _
                                            NVL(Txt部位方法.Text), _
                                            Not CheckPopedom(mstrPrivs, "检查登记"), _
                                            IIf(mintEditMode = 0, True, False), _
                                            IIf(mintEditMode = 0, 0, mlngAdviceID), , dcmTmpView)

    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
Dim rsTmp As ADODB.Recordset
    
    With txt医嘱内容
        .Text = ""
        Set rsTmp = SelectDiagItem() '提取项目
        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            .Text = .tag
            zlControl.TxtSelAll txt医嘱内容
            .SetFocus
            Exit Sub
        Else
            If AdviceInput(rsTmp) Then '根据选择项目设置部位及方法
                .tag = .Text
            Else '取消部位及方法
                .Text = .tag
                zlControl.TxtSelAll txt医嘱内容
                .SetFocus
                Exit Sub
            End If
        End If
    End With
End Sub
Private Function SelectDiagItem() As ADODB.Recordset
'选择检查项目
    Dim objPoint As RECT
    gstrSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') as 计算单位,nvl(A.标本部位,' ') as 标本部位," & _
                "A.操作类型 As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID," & _
                "nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID," & _
                "nvl(执行科室,0) As 执行科室ID,B.影像类别" & _
              " From 诊疗项目目录 A,影像检查项目 B,诊疗项目别名 C,诊疗执行科室 D" & _
              " Where A.ID=B.诊疗项目ID AND A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID" & _
                    " And D.执行科室ID=" & mlngCurDeptId & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " & _
                    " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) " & _
                    " And A.服务对象 IN(" & IIf(mintSourceType = 3, "1,2,4", mintSourceType) & ",3) " & _
                    " And Nvl(A.单独应用,0)=1" & _
                    " And Nvl(A.适用性别,0) IN (" & IIf(cbo性别.Text Like "*男*", "1,0)", "2,0)") & _
                    " And Nvl(A.执行频率,0) IN(0,1)" & _
                    " And (" & zlCommFun.GetLike("A", "编码", txt医嘱内容) & _
                            " Or " & zlCommFun.GetLike("A", "名称", txt医嘱内容) & _
                            " Or " & zlCommFun.GetLike("C", "简码", txt医嘱内容) & ")"
    objPoint = zlControl.GetControlRect(txt医嘱内容.hWnd)
     Set SelectDiagItem = zlDatabase.ShowSelect(Me, gstrSQL, 0, "选择申请项目", True, Me.txt医嘱内容.Text, "", True, True, True, objPoint.Left, objPoint.Top, Me.txt医嘱内容.Height, True, True, True)
End Function

Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的部位及方法
'参数：rsInput=选择返回的记录集
'返回：mstrExtData "部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
    Dim rsTemp As ADODB.Recordset
    Dim strExtData As String, strAppend As String
    Dim blnOk As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    Dim lngHwnd As Long, int服务对象 As Integer
    Dim blnChangeModality As Boolean
    
    On Error GoTo errHandle
    
    If Not rsInput Is Nothing Then
        txt医嘱内容.Text = Replace(Replace(rsInput!名称, ",", ""), "'", "") '暂时显示
    End If
    
    With t_Pati
        .lng病人ID = mlngPatiId
        If mintSourceType = 2 Then  '住院，填写主页ID
            .lng主页ID = mlngPageID
        Else
            .str挂号单 = mstrRegNo
        End If
        .str性别 = zlStr.NeedName(cbo性别.Text)
    End With
    
    lngHwnd = IIf(mintCheckInMode = 1, Me.txt检查号.hWnd, Me.Txt联系地址.hWnd)
    int服务对象 = IIf(mintSourceType <> 2, 1, 2)
    strExtData = ""
    strAppend = mstrAppend
    
    On Error Resume Next
    '接口改造：int场合没有传入，现传入0，bytUseType以前没有传入现传0
    blnOk = frmAdviceEditEx.ShowMe(Me, lngHwnd, t_Pati, 0, 0, 0, 1, int服务对象, , , , rsInput!诊疗项目ID, strExtData, strAppend)

    If Not blnOk Or strExtData = "" Then Exit Function
    err.Clear
    On Error GoTo errHandle
    
    mstrExtData = strExtData        '返回 "部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
    mstrAppend = strAppend '     检查="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    mlngClinicID = rsInput!诊疗项目ID

    
    Txt部位方法.tag = Split(mstrExtData, Chr(9))(1) '执行标记
    Txt部位方法.Text = Replace(get部位方法(mstrExtData), "),", ")" & vbCrLf)
    Txt部位方法.Text = Txt部位方法.Text & vbCrLf & get附件项目(mstrAppend)
    
    blnChangeModality = IIf(mstrItemType = rsInput!影像类别, False, True)
    mstrItemType = rsInput!影像类别
    
    If mblnRegToCheck And Not mblnUseAdviceNo And Not mblnUsePatientNo _
        And ((Trim(txt检查号.Text) = "") Or (Trim(txt检查号.Text) <> "" And blnChangeModality = True And mblnUseModalityPretext = True)) Then
        txt检查号.Text = Next检查号: txt检查号.tag = txt检查号.Text '初始检查号
    End If
    
'    mstrItemIDS = "" '可能改变项目,所以得先赋0
'    gstrSQL = "select 收费项目ID FROM 诊疗收费关系　Where 诊疗项目id=[1]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费细目ID", CLng(mlngClinicID))
'    Do Until rsTemp.EOF
'        mstrItemIDS = mstrItemIDS & "," & rsTemp!收费项目ID
'        rsTemp.MoveNext
'    Loop
'    mstrItemIDS = Mid(mstrItemIDS, 2)

    AdviceInput = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function
Private Function get附件项目(ByVal strAppend As String) As String
Dim i As Integer, strReturn As String
    For i = 0 To UBound(Split(strAppend, "<Split1>"))
        strReturn = strReturn & Split(Split(strAppend, "<Split1>")(i), "<Split2>")(0) & ":" & Split(Split(strAppend, "<Split1>")(i), "<Split2>")(3) & vbCrLf
    Next
    get附件项目 = strReturn
End Function
Private Function get部位方法(ByVal strExtData As String) As String
'入:部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中
'出:部位名1(方法名1,方法名2),部位名2(方法名1,方法名2)-----
Dim i As Integer, strReturn As String, Arr部位
    Arr部位 = Split(Split(strExtData, Chr(9))(0), "|")
    For i = 0 To UBound(Arr部位)
        strReturn = strReturn & "," & Split(Arr部位(i), ";")(0) & "(" & Split(Arr部位(i), ";")(1) & ")"
    Next
    get部位方法 = Mid(strReturn, 2)
End Function

Private Sub cmdSelectPinyinName_Click()
    Dim i As Long
    Dim strPinyinName As String
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    
    On Error GoTo errHandle
    strPinyinName = GetPinyinName(PatiIdentify.Text, mintCapital, mblnUseSplitter)
    If strPinyinName = "" Then Exit Sub

    Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
    With objPopup.Controls
        For i = 0 To UBound(Split(strPinyinName, ","))
            Set objControl = .Add(xtpControlButton, i + 1, Split(strPinyinName, ",")(i))
        Next
    End With
    objPopup.ShowPopup
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtp出生日期_Change()
    txt年龄.Text = ReCalcOld(dtp出生日期.value, cboAge)
End Sub

Private Sub RefreshObjEnabled(Optional lngRegType As Long)
'mintEditMode '0－登记、1－登记后修改、2－报到、3－报到后修改
'lngRegType  '1-新登记、2-提取登记、3-复制登记
    Dim blnEditableState As Boolean
    
    Dim strRegType As String
    
    '全部状态下的统一设置
    txtPatientDept.Enabled = False
    txtID.Enabled = False
    txtBed.Enabled = False
    Txt部位方法.Locked = True
    
    '通过权限来控制病人基本信息是否能被修改
    blnEditableState = IIf(Not CheckPopedom(mstrPrivs, "强制修改住院门诊信息"), (mintSourceType = 3), True)
    
    
    '基本信息，只有mintSourceType = 3外诊的情况下可以修改
    PatiIdentify.objTxtInput.Locked = Not (mintSourceType = 3)
    
    cbo性别.Enabled = mintSourceType = 3: cboAge.Enabled = mintSourceType = 3
    Call sutSetTxtEnable(txt年龄, mintSourceType = 3)
    dtp出生日期.Enabled = mintSourceType = 3
    Call sutSetTxtEnable(Txt身份证号, mintSourceType = 3)
            
    cbo费别.Enabled = (mintSourceType = 3)
    cbo付款方式.Enabled = (mintSourceType = 3): cbo民族.Enabled = blnEditableState
    cbo职业.Enabled = blnEditableState: cbo婚姻.Enabled = blnEditableState
    
    '技师确认后将不能进行修改
    cbo技师一.Enabled = Not mblnExamineDoctorVerify
    cbo技师二.Enabled = Not mblnExamineDoctorVerify
    
    '这三个信息一直都可以修改
    Call sutSetTxtEnable(Txt电话, True)
    Call sutSetTxtEnable(Txt邮编, True)
    Call sutSetTxtEnable(Txt联系地址, True)
    
    Frame3.Visible = True
    frm其他信息.Top = Frame2.Top + Frame2.Height + 450
            
    Select Case mintEditMode
        Case 0          '0－登记
            If lngRegType = 1 Then
                strRegType = " （ 新病人 ）"
            ElseIf lngRegType = 2 Then
                strRegType = " （ 提取病人 ）"
            ElseIf lngRegType = 3 Then
                strRegType = " （ 复制病人 ）"
            End If
            
            Me.Caption = "检查登记" & strRegType
            
            cboRoom.Enabled = mblnRegToCheck: cbo技师一.Enabled = mblnRegToCheck: cbo技师二.Enabled = mblnRegToCheck:
            If Not mblnRegToCheck Then
                Frame3.Visible = False
                frm其他信息.Top = Frame2.Top + Frame2.Height - 100
            End If
            cbo造影剂.Enabled = mblnRegToCheck
            
            '登记的时候，姓名允许修改
            PatiIdentify.objTxtInput.Locked = False
            
            cmdSelectPinyinName.Enabled = False
            
            Call sutSetTxtEnable(Txt英文名, True)
            If mblnUseAdviceNo Or mblnUsePatientNo Then
                txt检查号.Enabled = False
            Else
                Call sutSetTxtEnable(txt检查号, mblnRegToCheck)
            End If
            Call sutSetTxtEnable(Txt造影用量, mblnRegToCheck)
            Call sutSetTxtEnable(Txt造影浓度, mblnRegToCheck)
            Call sutSetTxtEnable(Txt身高, mblnRegToCheck)
            Call sutSetTxtEnable(Txt体重, mblnRegToCheck)
            Call sutSetTxtEnable(txt附加主述, mblnRegToCheck)
        Case 1          '1－登记后修改
            Me.Caption = "修改信息"
            
            cmdSelectPinyinName.Enabled = False
            cboRoom.Enabled = False:  cbo技师一.Enabled = False: cbo技师二.Enabled = False
            Frame3.Visible = False
            frm其他信息.Top = Frame2.Top + Frame2.Height - 100
            cbo造影剂.Enabled = False: dtp(0).Enabled = False
            dtp(1).Enabled = False:  cmdSel.Enabled = mintSourceType = 3
            chk紧急.Enabled = False: cbo开单科室.Enabled = False
            cbo医生1.Enabled = False: cbo医生2.Enabled = False
            
            PatiIdentify.Enabled = (mintSourceType = 3)
            
            Call sutSetTxtEnable(txt送检单位, False)
            Call sutSetTxtEnable(txt送检医生, False)
            
            Call sutSetTxtEnable(txt医嘱内容, IIf(mintSourceType = 3, True, False))
            Call sutSetTxtEnable(Txt英文名, False)
            
            Call sutSetTxtEnable(txt检查号, False)
            Call sutSetTxtEnable(Txt造影用量, False)
            Call sutSetTxtEnable(Txt造影浓度, False)
            Call sutSetTxtEnable(Txt身高, False)
            Call sutSetTxtEnable(Txt体重, False)
            Call sutSetTxtEnable(txt附加主述, False)
        Case 2          '2－报到
            Me.Caption = "检查报到"
            
            cmdSelectPinyinName.Enabled = True
            cbo技师一.Enabled = True
            cbo技师二.Enabled = True
            cbo开单科室.Enabled = False: cbo医生1.Enabled = False: cbo医生2.Enabled = False
            chk紧急.Enabled = False: dtp(0).Enabled = False
            dtp(1).Enabled = True: cmdSel.Enabled = False
            
            Call sutSetTxtEnable(txt送检单位, False)
            Call sutSetTxtEnable(txt送检医生, False)
            
            Call sutSetTxtEnable(txt医嘱内容, False)
            
            Call sutSetTxtEnable(Txt英文名, False)
            Call sutSetTxtEnable(txt附加主述, True)
        Case 3          '3－报到后修改
            Me.Caption = "修改信息"
            
            cmdSelectPinyinName.Enabled = True
            cboRoom.Enabled = True
            cbo造影剂.Enabled = True: dtp(0).Enabled = False
            dtp(1).Enabled = True: cmdSel.Enabled = False
            chk紧急.Enabled = False: cbo开单科室.Enabled = False
            cbo医生1.Enabled = False: cbo医生2.Enabled = False
            
            PatiIdentify.Enabled = (mintSourceType = 3)
            
            Call sutSetTxtEnable(txt送检单位, False)
            Call sutSetTxtEnable(txt送检医生, False)
            
            Call sutSetTxtEnable(txt医嘱内容, False)
            
            Call sutSetTxtEnable(Txt英文名, False)
            Call sutSetTxtEnable(txt检查号, True)
            Call sutSetTxtEnable(Txt造影用量, True)
            Call sutSetTxtEnable(Txt造影浓度, True)
            Call sutSetTxtEnable(Txt身高, True)
            Call sutSetTxtEnable(Txt体重, True)
            Call sutSetTxtEnable(txt附加主述, True)
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    '创建卡结算部件
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '初始化卡结算部件
    mobjSquareCard.zlInitComponents Me, mlngModul, glngSys, gstrDBUser, gcnOracle
    
    '初始化PatiIdentify
    PatiIdentify.IDKindStr = "姓|姓名|0|0|0|0|0|0|0|0|0;医|医保号|0|0|0|0|0|0|0|0|0;身|二代身份证|1|2|18|0|0||0|1|0;IC|IC卡|1|3|8|0|0||0|1|0;门|门诊号|0|0|0|0|0|0|0|0|0;就|就诊卡|0|1|8|0|0||0|0|0;挂|挂号单号|0|0|0|0|0|0|0|0|0;据|收费单据号|0|0|0|0|0|0|0|0|0"
    PatiIdentify.zlInit Me, glngSys, mlngModul, gcnOracle, gstrDBUser, mobjSquareCard, PatiIdentify.IDKindStr
    
    '获取IDKindStr
    If Not mobjSquareCard Is Nothing Then
        'PatiIdentify.objIDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(PatiIdentify.objIDKind.IDKindStr)
        
        '取缺省的刷卡方式
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
        '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
        '第7位后,就只能用索引,不然取不到数
        'oneSquardCard.bln缺省卡号密文 = Trim(PatiIdentify.GetfaultCard.卡号密文规则) <> ""
        
        'oneSquardCard.lng缺省卡类别ID = PatiIdentify.GetDefaultCardTypeID
    End If
    
    
    '赋默认值
    mlngUnicode = 0
    mlngTypeSuit = 0
    mblnLike = False
    mlngLike = 0
    mblnChangeNo = False
    mBeforeDays = 2
    If mintEditMode = 0 Then mlngBaby = 0        '设置默认值，不是婴儿,只有登记模式才设置
    
    '从注册表取得检查技师一 二的值
    mstrExamineDoctorFst = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查技师一", "")
    mstrExamineDoctorSed = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查技师二", "")

    
    Call ClearFaceData
End Sub

Private Sub cbo医生1_DropDown()
On Error GoTo errHandle
    Call SendMessage(cbo医生1.hWnd, &H160, 300, 0)
errHandle:
End Sub

Private Sub cbo医生2_DropDown()
On Error GoTo errHandle
    Call SendMessage(cbo医生2.hWnd, &H160, 300, 0)
errHandle:
End Sub

Private Sub InitParameter()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select 是否技师确认,检查技师 from 影像检查记录 where 医嘱id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    '技师是否确认
    If rsTemp.RecordCount > 0 Then
        mblnExamineDoctorVerify = NVL(rsTemp!是否技师确认, 0) = 1
        mstrExamineDoctorName = NVL(rsTemp!检查技师)
    End If
    
    mlngGoOnReg = Val(zlDatabase.GetPara("连续登记申请", glngSys, mlngModul, 0)) '连续登记
    mblnRegToCheck = (Val(GetDeptPara(mlngCurDeptId, "登记后直接检查", 0)) = 1) '登记后直接检查
    mblnAllPatientIsOutside = IIf(Val(GetDeptPara(mlngCurDeptId, "所有登记病人标记为外来", 0)) = 0, False, True)
    mblnUsePacsQueue = IIf(Val(GetDeptPara(mlngCurDeptId, "启动排队叫号", 0)) = 0, False, True)
    mlngPacsQueueNoWay = Val(GetDeptPara(mlngCurDeptId, "排队叫号编码规则", 0))
    mblnSelectDefRoom = Val(GetDeptPara(mlngCurDeptId, "报到时分配默认执行间", 0))
    mblnNameColColorCfg = GetDeptPara(mlngCurDeptId, "姓名颜色区分", 0) = "1"     '姓名颜色区分
    mblnOrdinaryNameColColorCfg = GetDeptPara(mlngCurDeptId, "缺省类型病人姓名颜色区分", 0) = "1"     '缺省类型病人姓名颜色区分
    
    If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
    '影像采集系统才需要根据不同的费用执行模式进行处理
        mlngVideoStationMoneyExeModle = Val(zlDatabase.GetPara("采集费用执行模式", glngSys, mlngModul, 0))
    Else
        mlngPacsStationMoneyExeModle = Val(zlDatabase.GetPara("医技费用执行模式", glngSys, mlngModul, 0))
    End If
    
    mlngInsureCheckType = Val(zlDatabase.GetPara(59, glngSys))  '获取医保对码检查类型
    If mlngInsureCheckType <> 0 Then
        Set mobjInsure = CreateObject("zl9Insure.clsInsure")
    End If
    
    strSql = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "患者检查号保持不变"
                mlngUnicode = NVL(rsTemp!参数值, 0)
            Case "检查号保持不变类别"
                mlngUnicodeType = NVL(rsTemp!参数值, 0)
            Case "检查号生成方式"
                mlngBuildType = NVL(rsTemp!参数值, 0)
            Case "匹配数据库项目"
                mlngTypeSuit = NVL(rsTemp!参数值, 0)
            Case "登记时姓名模糊查找天数"
                mblnLike = IIf(NVL(rsTemp!参数值, 0) <> 0, True, False)
                mlngLike = Abs(NVL(rsTemp!参数值, 0))
            Case "手工调整检查号"
                mblnChangeNo = NVL(rsTemp!参数值, 0) = 1
            Case "使用患者号"
                mblnUsePatientNo = NVL(rsTemp!参数值, 0) = 1
            Case "使用医嘱号"
                mblnUseAdviceNo = NVL(rsTemp!参数值, 0) = 1
            Case "默认过滤天数"
                mBeforeDays = Val(NVL(rsTemp!参数值, 2))
                If mBeforeDays > 15 Or mBeforeDays <= 0 Then
                    mBeforeDays = 2
                End If
            Case "允许检查号重复"
                mblnCanOverWrite = NVL(rsTemp!参数值, 0) = 1
            Case "启动关联病人"
                mblnUseReferencePatient = NVL(rsTemp!参数值, 0) = 1
            Case "拼音名大小写"
                mintCapital = NVL(rsTemp!参数值, 0)
            Case "拼音名分隔符"
                mblnUseSplitter = NVL(rsTemp!参数值, 0) = 0
            Case "检查号前缀"
                mblnUseModalityPretext = (NVL(rsTemp!参数值, 0) = 1)
        End Select
        rsTemp.MoveNext
    Wend
    
    Call InitFaceScheme
End Sub

Public Sub InitMvar()
    mintSourceType = 3
    mlngPatiId = 0
    mlngPageID = 0
    mstrItemType = ""
    mInputType = 6
    mstrChargeNo = ""
    mstrRegNo = ""
    mstrExtData = ""
    mlngClinicID = 0
'    mstrItemIDS = ""
    mstrAppend = ""
    mstrOutNo = 0
    mstrCardNo = ""
    mstrCardPass = ""
End Sub

Private Function Next检查号(Optional ByVal lngCurNo As Long = 0) As String
'------------------------------------------------
'功能：获取下一个检查号
'参数： lngCurNo -- 当前号码
'返回：下一个可用的检查号
'------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim varResult As Variant
    
'mlngUnicode, mlngUnicodeType, mlngBuildType '患者检查号保持不变;不变类别 0-按类别不变 1-按科室不变;0-按类别递增 1-按科室递增
    
    On Error GoTo errH
    
    '使用病人ID
    If mblnUsePatientNo Then
        Next检查号 = mlngPatiId
        mstrNextCheckNo = Next检查号
        Exit Function
    End If
    
    '使用医嘱ID
    If mblnUseAdviceNo Then
        Next检查号 = mlngAdviceID
        mstrNextCheckNo = Next检查号
        Exit Function
    End If
    
    '检查过的病人保持不变，查找原来的检查号
    If mlngUnicode = 1 Then
        If mlngUnicodeType = 0 Then '0-按类别不变 1-按科室不变
            strSql = "Select Max(B.检查号) 最大号码" & vbNewLine & _
                        " From 病人医嘱记录 A, 影像检查记录 B" & vbNewLine & _
                        " Where A.病人id = [1] And A.相关id Is Null And A.ID = B.医嘱id And B.影像类别 = [2]"
        Else
            strSql = "Select Max(C.检查号) 最大号码" & vbNewLine & _
                        " From 病人医嘱记录 A, 影像检查记录 C" & vbNewLine & _
                        " Where A.病人id = [1] And A.相关id Is Null And A.执行科室id = [3] And A.ID = C.医嘱id"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查号提取", mlngPatiId, mstrItemType, mlngCurDeptId)
        
        If NVL(rsTemp!最大号码, 0) <> 0 Then
            Next检查号 = CStr(rsTemp!最大号码)
            mstrNextCheckNo = Next检查号
            Exit Function
        End If
    End If
    
    '按生成规则重取
    varResult = zlDatabase.GetNextNo(123, mlngCurDeptId, mstrItemType, lngCurNo)
    If IsNull(varResult) = False Then
        Next检查号 = varResult
    End If
    
    mstrNextCheckNo = Next检查号
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReCalcBirth(ByVal strOld As String, ByVal str年龄单位 As String) As String
'功能:根据年龄和年龄单位估算病人的出生日期,年龄单位为岁时,出年月日假定为1月1号,年龄单位为月时,出生日期假定为1号
'返回:出生日期
    Dim strTmp As String, strFormat As String, lngDays As Long
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    
    strTmp = "____-__-__"
    If str年龄单位 = "" Then
        strFormat = "YYYY-MM-DD"
        If strOld Like "*岁*月" Or strOld Like "*岁*个月" Then
            strFormat = "YYYY-MM-01"
            lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "岁") + 1))
        ElseIf strOld Like "*月*天" Or strOld Like "*个月*天" Then
            lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "月") + 1))
        ElseIf strOld Like "*岁" Or IsNumeric(strOld) Then
            strFormat = "YYYY-01-01"
            lngDays = 365 * Val(strOld)
        ElseIf strOld Like "*月" Or strOld Like "*个月" Then
            strFormat = "YYYY-MM-01"
            lngDays = 30 * Val(strOld)
        ElseIf strOld Like "*天" Then
            lngDays = Val(strOld)
        End If
        If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, curDate), strFormat)
    ElseIf strOld <> "" Then
        Select Case str年龄单位
            Case "岁"
                If Val(strOld) > 200 Then lngDays = -1
            Case "月"
                If Val(strOld) > 2400 Then lngDays = -1
            Case "天"
                If Val(strOld) > 73000 Then lngDays = -1
        End Select
        
        If lngDays = 0 Then
            strTmp = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
            strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, curDate), "YYYY-MM-DD")
            
            If str年龄单位 = "岁" Then
                strTmp = Format(strTmp, "YYYY-01-01")
            ElseIf str年龄单位 = "月" Then
                strTmp = Format(strTmp, "YYYY-MM-01")
            End If
        End If
    End If
    If strTmp = "____-__-__" Then strTmp = Format(curDate, "YYYY-MM-DD")
    ReCalcBirth = strTmp
End Function
Function CheckOldData(ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox) As Boolean
'功能：检查年龄输入值的有效性
'返回：
    If Not IsNumeric(txt年龄.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo年龄单位.Text
        Case "岁"
            If Val(txt年龄.Text) > 200 Then
                MsgBoxD Me, "年龄不能大于200岁!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "月"
            If Val(txt年龄.Text) > 2400 Then
                MsgBoxD Me, "年龄不能大于2400月!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "天"
            If Val(txt年龄.Text) > 73000 Then
                MsgBoxD Me, "年龄不能大于73000天!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function
Private Function ReCalcOld(ByVal DateBir As Date, ByRef cbo年龄单位 As ComboBox, _
    Optional ByVal lng病人ID As Long, Optional ByVal RequestDate As Date) As String
'功能:根据出生日期重新计算病人的年龄,重设年龄单位
'返回:年龄,年龄单位
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
 
 On Error GoTo errH
    
    If RequestDate = CDate("0") Then
        strSql = "Select Zl_Age_Calc([1],[2]) old From Dual"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, _
                                App.ProductName, _
                                lng病人ID, _
                                IIf(DateBir = CDate("0"), Null, DateBir) _
                                )
    Else
        strSql = "Select Zl_Age_Calc([1],[2], [3]) old From Dual"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, _
                                App.ProductName, _
                                lng病人ID, _
                                IIf(DateBir = CDate("0"), Null, DateBir), _
                                IIf(RequestDate = CDate("0"), Null, RequestDate) _
                                )
    End If
                            
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*岁" Or rsTmp!old Like "*月" Or rsTmp!old Like "*天" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call zlControl.CboLocate(cbo年龄单位, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo年龄单位.ListIndex = -1
            End If
        ElseIf rsTmp!old Like "*小时" Or rsTmp!old Like "*分钟" Then
            strTmp = rsTmp!old
            cbo年龄单位.ListIndex = -1
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo年龄单位.ListIndex = 0
            Else
                cbo年龄单位.ListIndex = -1
            End If
        End If
    End If
    If cbo年龄单位.ListIndex = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatient(strCode As String, blnCard As Boolean) As ADODB.Recordset
'功能：读取病人信息，并显示该病人存在的医嘱时间
    Dim strNO As String, strSeek As String
    Dim objRect As RECT, blnCancel As Boolean
    Dim lng卡类别ID As Long
    Dim lng病人ID As Long
    Dim rsTemp As ADODB.Recordset
    
    'mInputType   1-病人ID 2-住院号 3-门诊号 4-挂号单 5-收费单据号 6-姓名 7-医保号 8-身份证号 9-IC卡号
    '一卡通修改之后，mInputType中不存在就诊卡了，就诊卡算到所有动态卡之中，通过病人ID提取信息
    On Error GoTo errH

    mstrChargeNo = "": mstrRegNo = ""
    strSeek = strCode
    '判断当前输入模式
    Select Case PatiIdentify.IDKindIDX
        Case PatiIdentify.GetKindIndex(IDKind_医保号)
            mInputType = 7
            strSeek = strCode
        Case PatiIdentify.GetKindIndex(IDKind_身份证号)
            mInputType = 8
            strSeek = strCode
        Case PatiIdentify.GetKindIndex(IDKind_IC卡号)
            mInputType = 9
            strSeek = strCode
        Case PatiIdentify.GetKindIndex(IDKind_门诊号)
            mInputType = 3
            strSeek = Val(strCode)
        Case PatiIdentify.GetKindIndex(IDKind_住院号)
            mInputType = 2
            strSeek = Val(strCode)
        Case PatiIdentify.GetKindIndex(IDKind_挂号单)
            mInputType = 4
            strSeek = strCode
        Case PatiIdentify.GetKindIndex(IDKind_收费单据号)
            mInputType = 5
            strSeek = strCode
        Case Else       '使用姓名的时候，经常直接刷卡，所以姓名和刷卡的放在一起处理
            
            If PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(IDKind_姓名) And blnCard = False And InStr(",1,2,3,4,5,6,7,8,9,0,", Left(strCode, 1)) <= 1 Then
                '是姓名，但是不是刷卡的
                If Left(strCode, 1) = "-" And IsNumeric(Mid(strCode, 2)) Then    '病人ID
                    mInputType = 1
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "+" And IsNumeric(Mid(strCode, 2)) Then '住院号
                    mInputType = 2
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "*" And IsNumeric(Mid(strCode, 2)) Then '门诊号
                    mInputType = 3
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "." Then '挂号单
                    mInputType = 4
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "/" Then '收费单据号
                    mInputType = 5
                    strSeek = Mid(strCode, 2)
                ElseIf Not IsNumeric(Mid(strCode, 2)) Then '当作姓名
                    mInputType = 6
                    strSeek = strCode
                End If
            Else
                '处理动态部分的医疗卡
                '其他类别的，获取相关的病人ID
                '其他类别的,获取相关的病人ID
                '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
                '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
                '第7位后,就只能用索引,不然取不到数
                
                If PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(IDKind_姓名) And blnCard Then
                    lng卡类别ID = Val(PatiIdentify.GetDefaultCardTypeID)
                Else
                    lng卡类别ID = Val(PatiIdentify.GetCurCard.接口序号)
                End If
                
                If lng卡类别ID <> 0 Then
                    If mobjSquareCard.zlGetPatiID(lng卡类别ID, strCode, False, lng病人ID) = False Then
                        lng病人ID = 0
                    End If
                Else
                    If mobjSquareCard.zlGetPatiID(IIf(PatiIdentify.GetCurCard.名称 = "姓名", "就诊卡号", PatiIdentify.GetCurCard.名称), strCode, False, lng病人ID) = False Then
                        lng病人ID = 0
                    End If
                End If
                '标记查找方式使用病人ID
                mInputType = 1
                strSeek = lng病人ID
            End If
    End Select
    
    '病人ID 姓名 性别 年龄 来源 病人科室 主页id 病人科室ID 医生 住院号 门诊号 当前床号
    '    费别 医疗付款方式 身份证号 民族 职业 婚姻状况 电话 邮编 地址
    If mInputType = 1 Then '病人ID
        gstrSQL = "select 病人id,姓名,性别,年龄,出生日期,来源ID,主页ID,病人科室ID,当前科室ID,医生,门诊号,住院号,就诊卡号,就诊状态,卡验证码,当前床号,费别" & _
                        ",医疗付款方式,身份证号,民族,职业,婚姻状况,电话,邮编,地址,合同单位ID, 新病人" & _
                    " From(Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,nvl(B.执行人,'') As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人,B.登记时间" & _
                  " From 病人信息 A,病人挂号记录 B Where A.病人ID=[2] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and '%'='%' " & _
                  " order by B.登记时间 desc) where rownum=1" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 2 Then '住院号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.入院科室ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,B.住院医师 As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A,病案主页 B " & _
                  " Where A.住院号=[1] And A.病人ID=B.病人ID and A.出院时间 Is Null and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 3 Then '门诊号,输入门诊号的，认为是门诊病人
        gstrSQL = "select 病人id,姓名,性别,年龄,出生日期,来源ID,主页ID,病人科室ID,当前科室ID,医生,门诊号,住院号,就诊卡号,就诊状态,卡验证码,当前床号,费别" & _
                        ",医疗付款方式,身份证号,民族,职业,婚姻状况,电话,邮编,地址,合同单位ID, 新病人" & _
                    " From (Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,B.执行人 As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,B.登记时间,A.合同单位ID, 0 as 新病人" & _
                        " From 病人信息 A,病人挂号记录 B Where A.门诊号=[1] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and B.记录性质=1 and B.记录状态=1 Order By B.登记时间 Desc)" & _
                    " where Rownum=1 and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 4 Then '挂号单
        strNO = zlCommFun.GetFullNO(strSeek, 12)
        PatiIdentify.Text = strNO
'        mstrRegNo = strNO
        gstrSQL = "Select Distinct A.病人id, A.姓名, A.性别, A.年龄, To_Char(A.出生日期, 'yyyy-mm-dd') 出生日期, Decode(Nvl(A.在院, 0), 0, 1, 2) As 来源id," & vbNewLine & _
                    "                A.主页id, Nvl(B.执行部门id, B.转诊科室id) As 病人科室id, A.当前科室ID,B.执行人 As 医生, Nvl(A.门诊号, B.门诊号) 门诊号, A.住院号," & vbNewLine & _
                    "                A.就诊卡号, decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码, A.当前床号, A.费别, A.医疗付款方式, A.身份证号, A.民族, A.职业, A.婚姻状况, Nvl(A.家庭电话, A.联系人电话) 电话," & vbNewLine & _
                    "                Nvl(A.家庭地址邮编, A.单位邮编) 邮编, Nvl(A.家庭地址, A.工作单位) 地址, A.合同单位id, 0 as 新病人" & vbNewLine & _
                    "From 病人信息 A, 病人挂号记录 B" & vbNewLine & _
                    "Where B.NO = [3] And B.病人id = A.病人id and B.记录性质=1 and B.记录状态=1 and '%'='%'"  '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 5 Then '收费单据号，输入收费单据号的，认为是门诊病人
        strNO = zlCommFun.GetFullNO(strSeek, 13)
        PatiIdentify.Text = strNO
        mstrChargeNo = strNO
        
        '门诊费用记录的NO=病人挂号记录的NO，所以使用收费单据号提取病人的时候，同时记录挂号单。
        '如果没有挂号单为空，则通过收费单据号提取并登记的门诊病人，看不到医嘱内容。
'        mstrRegNo = strNO
        
        gstrSQL = "Select Distinct Nvl(A.病人id, 0) 病人id, Nvl(A.姓名, B.姓名) 姓名, Nvl(A.性别, B.性别) 性别, Nvl(A.年龄, B.年龄) 年龄," & vbNewLine & _
                    "                To_Char(A.出生日期, 'yyyy-mm-dd') 出生日期, Decode(Nvl(A.在院, 0), 0, 1, 2) As 来源id, A.主页id," & vbNewLine & _
                    "                Nvl(B.开单部门id, B.病人科室id) As 病人科室id,A.当前科室ID, Nvl(B.开单人, B.执行人) As 医生, Nvl(A.门诊号, B.标识号) 门诊号, A.住院号, A.就诊卡号, decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码," & vbNewLine & _
                    "                A.当前床号, A.费别, A.医疗付款方式, A.身份证号, A.民族, A.职业, A.婚姻状况, Nvl(A.家庭电话, A.联系人电话) 电话, Nvl(A.家庭地址邮编, A.单位邮编) 邮编," & vbNewLine & _
                    "                Nvl(A.家庭地址, A.工作单位) 地址, A.合同单位id, 0 as 新病人" & vbNewLine & _
                    "From 病人信息 A, 门诊费用记录 B" & vbNewLine & _
                    "Where B.NO = [3] And Mod(B.记录性质,10) = 1 And B.记录状态 = 1 And nvl(B.费用状态,0) <>1 And B.病人id = A.病人id(+) And '%' = '%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 6 Then '当作姓名
            gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "NVL(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                " From 病人信息 A where " & IIf(mblnLike = False, "A.姓名=[1]", IIf(mlngLike = 0, "instr(A.姓名,[1])>0", "A.登记时间 Between sysdate-" & mlngLike & " and sysdate and instr(A.姓名,[1])>0"))
    
    ElseIf mInputType = 7 Then '医保号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "NVL(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A Where A.医保号=[1] and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 8 Then '身份证号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "NVL(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A Where A.身份证号=[1] and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 9 Then 'IC卡号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "NVL(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "nvl(A.家庭地址邮编,A.单位邮编) 邮编,nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A Where A.IC卡号=[1] and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    End If
    
    gstrSQL = gstrSQL & " Union " & _
                "Select null 病人ID,'新病人' 姓名,'未知' 性别,'' 年龄,null 出生日期,3 As 来源ID,0 As 主页ID," & _
                        "0 As 病人科室ID,0 As 当前科室ID,'' As 医生,null as 门诊号,null as 住院号,'' as 就诊卡号,'' as 就诊状态, '' as 卡验证码,'' as 当前床号," & _
                        "'' as 费别,'' as 医疗付款方式,'' as 身份证号,'' as 民族,'' as  职业,'' as 婚姻状况,'' 电话,'' 邮编,'' 地址,0 合同单位ID, 1 as 新病人" & _
             " From dual where '%'='%'"
    gstrSQL = "select RowNum as ID,病人id,姓名,性别,年龄,出生日期,来源ID,主页ID,病人科室ID,当前科室ID,医生,门诊号," & _
                "住院号,就诊卡号, 就诊状态,卡验证码,当前床号,费别,医疗付款方式,身份证号,民族,职业,婚姻状况,电话,邮编,地址,合同单位ID" & _
                " From (" & gstrSQL & ") Order by 新病人 asc,病人ID desc"
    objRect = zlControl.GetControlRect(PatiIdentify.hWnd)
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在相同病人", CStr(strSeek), Val(strSeek), strNO)
    mblnIsSamePatient = IIf(rsTemp.RecordCount > 1, True, False)
    
    Set GetPatient = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "查病人信息", False, "病人ID", "", False, False, True, objRect.Left, objRect.Top, PatiIdentify.Height, blnCancel, True, False, CStr(strSeek), Val(strSeek), strNO)
        
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    strSql = "Select 编码,nvl(名称,'未知') as 名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "提取" & strDict)
    
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub InitDoctors(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    strSql = "Select /*+ RULE*/" & vbNewLine & _
                "Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b, 人员性质说明 c" & vbNewLine & _
                " Where a.人员id = b.Id And b.Id = c.人员id And c.人员性质 = '医生' And" & vbNewLine & _
                "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and a.部门id = [1] " & vbNewLine & _
                " Order By 简码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID)
    
    If mblnIsOutSideHosp Then
        cbo医生2.Clear
        If Not rsTmp.EOF Then
            Do Until rsTmp.EOF
                cbo医生2.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                If rsTmp!ID = UserInfo.ID Then cbo医生2.ListIndex = cbo医生2.NewIndex
                rsTmp.MoveNext
            Loop
            If cbo医生2.ListCount > 0 And cbo医生2.ListIndex = -1 Then cbo医生2.ListIndex = 0
            cbo医生2.Enabled = True
        End If
    Else
        cbo医生1.Clear
        If Not rsTmp.EOF Then
            Do Until rsTmp.EOF
                cbo医生1.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                If rsTmp!ID = UserInfo.ID Then cbo医生1.ListIndex = cbo医生1.NewIndex
                rsTmp.MoveNext
            Loop
            If cbo医生1.ListCount > 0 And cbo医生1.ListIndex = -1 Then cbo医生1.ListIndex = 0
            cbo医生1.Enabled = True
        End If
    End If
    
End Sub
Private Sub InitInput()
    Dim i As Integer, strInput As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select ID ,科室ID,参数值 from 影像流程参数 where 科室ID = [1] and 参数名 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId, CStr("输入控制"))
    If Not rsTemp.EOF Then
        strInput = NVL(rsTemp!参数值)
    End If
    
    For i = 0 To UBound(Split(strInput, "|"))
        Select Case Split(strInput, "|")(i)
            Case "英文名"
                Txt英文名.TabStop = False
            Case "性别"
                cbo性别.TabStop = False
            Case "年龄"
                txt年龄.TabStop = False
                cboAge.TabStop = False
            Case "出生日期"
                dtp出生日期.TabStop = False
            Case "身高"
                Txt身高.TabStop = False
            Case "体重"
                Txt体重.TabStop = False
            Case "费别"
                cbo费别.TabStop = False
            Case "付款方式"
                cbo付款方式.TabStop = False
            Case "身份证号"
                Txt身份证号.TabStop = False
            Case "民族"
                cbo民族.TabStop = False
            Case "职业"
                cbo职业.TabStop = False
            Case "婚姻"
                cbo婚姻.TabStop = False
            Case "电话"
                Txt电话.TabStop = False
            Case "邮编"
                Txt邮编.TabStop = False
            Case "地址"
                Txt联系地址.TabStop = False
            Case "执行间"
                cboRoom.TabStop = False
            Case "紧急"
                chk紧急.TabStop = False
            Case "检查设备"
                cboDevice.TabStop = False
            Case "检查号"
                txt检查号.TabStop = False
            Case "申请时间"
                dtp(0).TabStop = False
            Case "检查时间"
                dtp(1).TabStop = False
            Case "造影剂"
                cbo造影剂.TabStop = False
                Txt造影用量.TabStop = False
                Txt造影浓度.TabStop = False
            Case "检查技师"
                cbo技师一.TabStop = False
            Case "检查技师二"
                cbo技师二.TabStop = False
        End Select
    Next
End Sub
Public Sub InitRoomPati()
Dim rsTemp As ADODB.Recordset, i As Integer, lst As ListItem
    On Error GoTo errH:
    If cboRoom.ListCount < 1 Then '没有执行间
        Exit Sub
    End If
    With lvwRoom
        With .ColumnHeaders
            .Clear
            .Add , , "执行间", 2800
            .Add , , "病人总数", 1400, 1
            .Add , , "已报告", 1400, 1
            .Add , , "进行中", 1400, 1
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    gstrSQL = "Select Count(ID) 数量, 执行间, 状态" & vbNewLine & _
                "From (Select /*+ rule*/" & vbNewLine & _
                "        A.ID, Decode(Nvl(B.执行间, ''), '', '未定执行间', B.执行间) 执行间," & vbNewLine & _
                "        Decode(Nvl(D.病历id, 0), 0, '进行中', '已报告') 状态" & vbNewLine & _
                "       From 病人医嘱记录 A, 病人医嘱发送 B, 影像检查记录 C, 病人医嘱报告 D" & vbNewLine & _
                "       Where A.相关id Is Null And A.执行科室id = [1] And" & vbNewLine & _
                "             A.开始执行时间 Between To_Date(To_Char(Sysdate-" & (mBeforeDays - 1) & ", 'yyyy-mm-dd'), 'yyyy-mm-dd hh24:mi:ss') And Sysdate And" & vbNewLine & _
                "             A.ID = B.医嘱id And B.医嘱id = C.医嘱id And B.发送号 = C.发送号 And A.ID = D.医嘱id(+))" & vbNewLine & _
                "Group By 执行间, 状态" & vbNewLine & _
                "Order By 执行间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取执行间病人情况", mlngCurDeptId)

    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    For i = 0 To cboRoom.ListCount - 1
        Set lst = lvwRoom.ListItems.Add(, "_" & zlStr.NeedCode(cboRoom.list(i)), zlStr.NeedCode(cboRoom.list(i)))
        rsTemp.Filter = "执行间='" & zlStr.NeedCode(cboRoom.list(i)) & "'"
        Do Until rsTemp.EOF
            If rsTemp!状态 = "已报告" Then
                lst.SubItems(2) = rsTemp!数量
            Else
                lst.SubItems(3) = rsTemp!数量
            End If
            lst.SubItems(1) = Val(NVL(lst.SubItems(1), 0)) + rsTemp!数量
            rsTemp.MoveNext
        Loop
    Next
    
    rsTemp.Filter = "执行间='未定执行间'"
    If Not rsTemp.EOF Then Set lst = lvwRoom.ListItems.Add(, "_未定执行间", "未定执行间")
    Do Until rsTemp.EOF
        If rsTemp!状态 = "已报告" Then
            lst.SubItems(2) = rsTemp!数量
        Else
            lst.SubItems(3) = rsTemp!数量
        End If
        lst.SubItems(1) = Val(NVL(lst.SubItems(1), 0)) + rsTemp!数量
        rsTemp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitFaceScheme()
    '读取参数
    mblnNoshowReagent = Val(GetDeptPara(mlngCurDeptId, "显示造影剂", 1)) <> 1
    mblnNoshowAddons = Val(GetDeptPara(mlngCurDeptId, "显示附加主述", 1)) <> 1
    mblnInputOutInfo = Val(zlDatabase.GetPara("录入外院信息", glngSys, mlngModul, 0)) = 1
    mintCheckInMode = Val(zlDatabase.GetPara("登记模式", glngSys, mlngModul, 2))
    
    mblnIsPetitionScan = IIf(Val(GetDeptPara(mlngCurDeptId, "启用申请单扫描", 1)) = 1, True, False)   '读取启用申请单扫描参数
    Me.cmdPetitionCapture.Visible = mblnIsPetitionScan
    
    If mintCheckInMode <> 1 Then mintCheckInMode = 2
    
    If Not mblnInputOutInfo Then
        lbl送检单位.Visible = False
        txt送检单位.Visible = False
        lbl送检医生.Visible = False
        txt送检医生.Visible = False
        
        lbl医嘱内容.Top = 1530
        txt医嘱内容.Top = 1515
        cmdSel.Top = 1500
        Lbl部位方法.Top = 2040
        Txt部位方法.Top = 2010
        Txt部位方法.Height = 1400
    End If
    
    '因为附加主诉在造影剂的上方显示，所以先处理附加主诉
    If mblnNoshowAddons And Label29.Visible = True Then '不显示附加主诉，且附加主诉已经被显示，则关闭显示附加主诉
        Label29.Visible = False: txt附加主述.Visible = False: txt附加主述.Enabled = False
        '调整后面控件的位置
        Label26.Top = Label26.Top - 350: cbo造影剂.Top = cbo造影剂.Top - 370
        Label27.Top = Label27.Top - 350: Txt造影用量.Top = Txt造影用量.Top - 370
        Label28.Top = Label28.Top - 350: Txt造影浓度.Top = Txt造影浓度.Top - 370
        Label1.Top = Label1.Top - 370: cbo费别.Top = cbo费别.Top - 370
        Label13.Top = Label13.Top - 370: cbo付款方式.Top = cbo付款方式.Top - 370
        Label12.Top = Label12.Top - 370: lblCash.Top = lblCash.Top - 370
        frm其他信息.Height = frm其他信息.Height - 400
        cmdOK.Top = cmdOK.Top - 400: CmdCancle.Top = cmdOK.Top: chkRoom.Top = cmdOK.Top: cmdPetitionCapture.Top = cmdOK.Top
        lvwRoom.Top = lvwRoom.Top - 400: lbl执行间.Top = lvwRoom.Top
        Me.Height = Me.Height - 400
    End If
    
    If mblnNoshowReagent And Label26.Visible = True Then    '不显示造影剂，且造影剂已经被显示，则关闭造影剂的显示
        Label26.Visible = False: Label27.Visible = False: Label28.Visible = False
        cbo造影剂.Visible = False: cbo造影剂.Enabled = False
        Txt造影浓度.Visible = False: Txt造影浓度.Visible = False
        Txt造影用量.Visible = False: Txt造影用量.Visible = False
        '调整后面的控件位置
        Label1.Top = Label1.Top - 370: cbo费别.Top = cbo费别.Top - 370
        Label13.Top = Label13.Top - 370: cbo付款方式.Top = cbo付款方式.Top - 370
        Label12.Top = Label12.Top - 370: lblCash.Top = lblCash.Top - 370
        frm其他信息.Height = frm其他信息.Height - 400
        cmdOK.Top = cmdOK.Top - 400: CmdCancle.Top = cmdOK.Top: chkRoom.Top = cmdOK.Top: cmdPetitionCapture.Top = cmdOK.Top
        lvwRoom.Top = lvwRoom.Top - 400: lbl执行间.Top = lvwRoom.Top
        Me.Height = Me.Height - 400
    End If
    
    If mintCheckInMode = 1 Then     '精简模式
        frm其他信息.Visible = False
        cmdOK.Top = cmdOK.Top - frm其他信息.Height: CmdCancle.Top = cmdOK.Top: chkRoom.Top = cmdOK.Top: cmdPetitionCapture.Top = cmdOK.Top
        lvwRoom.Top = lvwRoom.Top - frm其他信息.Height: lbl执行间.Top = lvwRoom.Top
        Me.Height = Me.Height - frm其他信息.Height
    End If
End Sub

Private Sub ClearFaceData(Optional blnSaveName As Boolean)
    Dim curDate As Date
    
    If Not blnSaveName Then
    PatiIdentify.Text = ""
    End If
    PatiIdentify.tag = ""
    Txt英文名.Text = "":    Txt英文名.tag = ""
    txt年龄.Text = "":      cboAge.Visible = True: txt年龄.tag = ""
    Txt身高.Text = "":      Txt体重.Text = ""
    Txt身份证号.Text = "":  Txt电话.Text = ""
    Txt邮编.Text = "":      Txt联系地址 = ""
    txtPatientDept.Text = "":  txtID.Text = ""
    txtBed.Text = ""
    txt检查号.Text = "":    txt检查号.tag = ""
    Txt造影用量.Text = "":  Txt造影浓度.Text = ""
    txt医嘱内容.Text = "":  txt医嘱内容.tag = ""
    Txt部位方法.Text = "":  Txt部位方法.tag = ""
    
    curDate = zlDatabase.Currentdate
    
    dtp出生日期.value = Format(curDate, "yyyy-mm-dd")
    dtp(0).value = curDate
    dtp(1).value = Format(curDate, "yyyy-mm-dd HH:MM")
    
    cboAge.ListIndex = 0
    
End Sub

Private Sub InitEdit(ByVal blnIsChangeDept As Boolean)
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Integer
    
    On Error GoTo DBError
    
    If Not blnIsChangeDept Then
        cboAge.ListIndex = 0
        
        '性别
        Set rsTmp = GetDictData("性别")
        cbo性别.Clear
        If Not rsTmp Is Nothing Then
            For i = 1 To rsTmp.RecordCount
                cbo性别.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    cbo性别.ItemData(cbo性别.NewIndex) = 1
                    cbo性别.ListIndex = cbo性别.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
        
        '根据传入的图像数量来判断改变按钮的内容
        If mintEditMode > 0 Then cmdPetitionCapture.Caption = IIf(mintImgCount = 0, "申请单", "申请单(" & mintImgCount & "张)")
        
        '费别
        Set rsTmp = GetDictData("费别")
        cbo费别.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    cbo费别.ItemData(cbo费别.NewIndex) = 1
                    cbo费别.ListIndex = cbo费别.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '付款方式
        Set rsTmp = GetDictData("医疗付款方式")
        cbo付款方式.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo付款方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    cbo付款方式.ItemData(cbo付款方式.NewIndex) = 1
                    cbo付款方式.ListIndex = cbo付款方式.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '民族
        Set rsTmp = GetDictData("民族")
        cbo民族.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo民族.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    cbo民族.ItemData(cbo民族.NewIndex) = 1
                    cbo民族.ListIndex = cbo民族.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '职业
        Set rsTmp = GetDictData("职业")
        cbo职业.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo职业.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    cbo职业.ItemData(cbo职业.NewIndex) = 1
                    cbo职业.ListIndex = cbo职业.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '婚姻状况
        Set rsTmp = GetDictData("婚姻状况")
        cbo婚姻.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo婚姻.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    cbo婚姻.ItemData(cbo婚姻.NewIndex) = 1
                    cbo婚姻.ListIndex = cbo婚姻.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '开单科室
        strSql = " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                    " From 部门表 A,部门性质说明 B " & _
                    " Where B.部门ID = A.ID " & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                    " And (B.工作性质 IN('临床','体检','检查'))" & _
                    " Order by A.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        cbo开单科室.Clear
        Do Until rsTmp.EOF
            cbo开单科室.AddItem rsTmp!简码 & "-" & rsTmp!名称
            cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlngCurDeptId Then cbo开单科室.ListIndex = cbo开单科室.NewIndex
            rsTmp.MoveNext
        Loop
        If cbo开单科室.ListCount > 0 And Me.cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
        
        '造影剂
        strSql = "select 名称 from 造影剂"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        cbo造影剂.Clear
        cbo造影剂.AddItem "                 "
        Do Until rsTmp.EOF
            cbo造影剂.AddItem rsTmp!名称
            rsTmp.MoveNext
        Loop
    End If
    
    '检查技师
    strSql = "Select /*+ RULE*/" & vbNewLine & _
                "Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b " & vbNewLine & _
                " Where a.人员id = b.Id And " & vbNewLine & _
                "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and a.部门id = [1] " & vbNewLine & _
                " Order By 简码 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    '加载检查技师一
    cbo技师一.Clear
    Do Until rsTmp.EOF
        cbo技师一.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!ID = UserInfo.ID Then cbo技师一.ListIndex = cbo技师一.NewIndex
        rsTmp.MoveNext
    Loop
    'If cbo技师一.ListCount > 0 And cbo技师一.ListIndex = -1 And mintEditMode = 2 Then cbo技师一.ListIndex = 0
    
    '加载检查技师二
    cbo技师二.Clear
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            cbo技师二.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            If rsTmp!ID = UserInfo.ID Then cbo技师二.ListIndex = cbo技师二.NewIndex
            rsTmp.MoveNext
        Loop
    End If
    
    Call LoadDefaultCheckDoctor

    InitInput '光标经过位置
    
    '登记的情况，需要控制控件的可用性
    If mintEditMode = 0 Then Call RefreshObjEnabled(1)
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshRoom()
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    '初始化执行间
    If mlngCurDeptId = 0 Then
        strSql = "Select 执行间,检查设备 From 医技执行房间"
    Else
        strSql = "Select 执行间,检查设备 From 医技执行房间 Where 科室id = [1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    cboRoom.Clear
    Do While Not rsTmp.EOF
        cboRoom.AddItem rsTmp!执行间 & "-" & NVL(rsTmp!检查设备)
        rsTmp.MoveNext
    Loop
    
    If mblnUsePacsQueue Then cboRoom.AddItem "呼叫时指派"
    
    If cboRoom.ListCount <= 0 Then
        cboRoom.Enabled = False
    Else
        Call InitDevice
        
        strSql = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & mlngCurDeptId & "\" & Me.Name, "当前执行间", "") '提取上次登记时的执行间
        If Trim(strSql = "") Then '判断注册表执行间是否为空，若是，代表呼叫时指派
            cboRoom.ListIndex = cboRoom.ListCount - 1
        Else
            Call SeekIndex(cboRoom, strSql, True, , tNeedNo)
        End If
    End If
    
    If mintEditMode = 1 Or mintEditMode = 3 Then
    '这是修改的情况  根据数据库判断
        If Trim(NVL(mstrRoom)) = "" Then
            cboRoom.ListIndex = cboRoom.ListCount - 1
        Else
            Call SeekIndex(cboRoom, NVL(mstrRoom), True, , tNeedNo)
        End If
    Else
    '这是报到的情况，先判断mblnSelectDefRoom这个参数 ,若参数为否 不做修改 之前已经是注册表执行间
        If mblnSelectDefRoom Then
            Dim strRoomTmp As String
            
            strRoomTmp = GetDefaultRoom(mlngAdviceID, mlngCurDeptId)
            '根据医嘱ID，获取默认的执行间
            If Trim(strRoomTmp) <> "" Then
                Call SeekIndex(cboRoom, strRoomTmp, , , tNeedNo)
            End If
        End If
    End If
End Sub

Private Sub InitDevice(Optional ByVal CheckType As String)
'------------------------------------------------
'功能：初始化并填充影像设备
'参数： CheckType -影像类别
'返回：无
'------------------------------------------------
Dim rsTmp As ADODB.Recordset
    
    cboDevice.Clear
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where " & IIf(CheckType <> "", "影像类别=[1] AND ", "") & "类型=4 AND  状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CheckType)
    Do Until rsTmp.EOF
        cboDevice.AddItem rsTmp!设备号 & "-" & NVL(rsTmp!设备名)
        rsTmp.MoveNext
    Loop
End Sub
Private Sub LoadOldData(ByVal strOld As String, ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "岁") > 0 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    
    If strTmp = "" Then strTmp = 0
    txt年龄.Text = strTmp
    If cbo年龄单位.ListCount > 0 Then Call zlControl.CboSetIndex(cbo年龄单位.hWnd, lngIdx)
    If lngIdx = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
End Sub

Private Function CopyCheck(ByVal lngAdviceID As Long, ByVal lngSendNO As Long) As Boolean
'功能:用于复制登记，同一病人相同项目，不同部位
'返回： True--复制成功；False--复制信息不完整

    Dim rsTemp As New ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    
    Dim curDate As Date
    Dim lngPatientID As Long
    Dim lngPageID As Long
    Dim strSql As String

    On Error GoTo errHand
    CopyCheck = False
    
    curDate = zlDatabase.Currentdate
    
    strSql = "SELECT nvl(E.姓名,B.姓名) 姓名,nvl(E.性别,B.性别) 性别,nvl(E.年龄,B.年龄) 年龄,B.出生日期,B.费别,B.医疗付款方式,B.身份证号,B.民族,B.职业,Nvl(E.英文名,'') 英文名,E.身高,E.体重" & _
                    ",B.婚姻状况,Nvl(B.家庭电话,B.联系人电话) 电话,Nvl(B.家庭地址邮编,B.单位邮编) 邮编,nvl(B.家庭地址,B.工作单位) 地址,B.合同单位ID,B.门诊号,B.就诊卡号,B.卡验证码" & _
                    ",NVL(D.名称,'') AS 病人科室,A.病人科室ID,A.挂号单,Decode(A.病人来源,2,B.住院号,B.门诊号) As 病人号,Decode(B.住院号,NULL,NULL,B.当前床号) As 床号" & _
                    ",F.发送时间 开嘱时间,NVL(C.简码,0) 科室简码,NVL(C.名称,'未知') AS 开嘱科室,A.开嘱时间 As 申请时间, A.开嘱医生,A.紧急标志,F.首次时间,F.执行间,E.检查设备,A.医嘱内容,E.检查号,E.检查技师,E.检查技师二 " & _
                    ",DECODE(A.病人来源,2,2,1,1,4,4,3) AS 病人来源,Nvl(E.影像类别,G.影像类别) As 影像类别,B.病人id,A.主页id, Nvl(A.婴儿,0) As 婴儿,A.诊疗项目ID,E.附加主述" & _
                " FROM 病人医嘱发送 F,病人医嘱记录 A, 病人信息 B,部门表 C,部门表 D,影像检查记录 E,影像检查项目 G " & _
                " Where F.医嘱ID=[1] And F.发送号=[2] AND F.医嘱ID=A.ID" & _
                        " AND F.医嘱ID=E.医嘱ID(+) And F.发送号=E.发送号(+)  And A.病人ID=B.病人ID" & _
                        " And A.开嘱科室ID=C.ID And A.病人科室ID=D.ID And A.诊疗项目ID=G.诊疗项目ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取病人信息", lngAdviceID, lngSendNO)

    If rsTemp.EOF Then
        '检查病人信息不完整的原因，如果是没有“病人医嘱发送记录，则提示本次医嘱已被回退或作废
        gstrSQL = "Select 医嘱ID From 病人医嘱发送 Where 医嘱ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查医嘱状态", lngAdviceID)
        If rsTemp.EOF Then
            Call MsgBoxD(Me, "本次检查医嘱没有发送记录，可能是该医嘱已经被回退或者已作废，请刷新后检查医嘱状态！", vbInformation, gstrSysName)
        Else
            Call MsgBoxD(Me, "病人信息不完整，请与管理员联系！", vbInformation, gstrSysName)
        End If
        
        mlngResultState = 0
        cmdOK.Enabled = False
        Exit Function
    End If
    
    '如果是婴儿则应该显示婴儿的信息
    mlngBaby = rsTemp!婴儿
    
    If mlngBaby = 0 Then
Normal:
        PatiIdentify.Text = NVL(rsTemp!姓名)
        Call SeekIndex(cbo性别, NVL(rsTemp!性别), True)
        
        If NVL(rsTemp!年龄) <> "" Then
            LoadOldData rsTemp!年龄, txt年龄, cboAge
        Else
            '没有年龄情况下的处理方式
            Call ReCalcOld(Format(NVL(rsTemp!出生日期, curDate), "yyyy-mm-dd hh:mm:ss"), _
                        cboAge, _
                        0, _
                        Format(curDate, "yyyy-mm-dd hh:mm:ss"))
        End If
 
        dtp出生日期.CustomFormat = "yyyy-MM-dd"
    
        If Trim(NVL(rsTemp!出生日期)) = "" Then
            Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
        Else
            '判断出生日期是否满一年，不满一年按则可能是婴儿
            If zlDatabase.Currentdate - CDate(rsTemp!出生日期) >= 366 Then
                dtp出生日期.value = Format(NVL(rsTemp!出生日期), "yyyy-mm-dd")
            Else
                dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
                dtp出生日期.value = Format(NVL(rsTemp!出生日期), "yyyy-mm-dd hh:mm:ss")
            End If
        End If
        
        
    Else
        lngPageID = NVL(rsTemp!主页ID, 0)
        
        strSql = "Select A.开嘱时间 As 申请时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                 "  Where a.病人ID = b.病人ID And b.主页id = [2] And b.序号 = [3] And a.ID = [1]"

        Set rsBaby = zlDatabase.OpenSQLRecord(strSql, "提取婴儿信息", lngAdviceID, lngPageID, mlngBaby)
            
        If rsBaby.EOF Then
            GoTo Normal
        Else
            PatiIdentify.Text = NVL(rsBaby!婴儿姓名)
            Call SeekIndex(cbo性别, NVL(rsBaby!婴儿性别), True)

            If Trim(NVL(rsTemp!出生日期)) <> "" Then
                txt年龄.Text = ReCalcOld(Format(NVL(rsBaby!出生时间, curDate), "yyyy-mm-dd hh:mm:ss"), _
                                            cboAge, _
                                            0, _
                                            Format(curDate, "yyyy-mm-dd hh:mm:ss"))
            End If
            
            '如果是婴儿，出生日期需要显示时分秒
            dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
            dtp出生日期.value = Format(NVL(rsBaby!出生时间), "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    
    dtp(0).value = Format(curDate, "yyyy-mm-dd HH:MM")
    
    Txt英文名 = Decode(NVL(rsTemp!英文名), "", zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter), rsTemp!英文名)
    
    If Trim(txt年龄) = "" Then txt年龄 = 0
    Txt身高 = NVL(rsTemp!身高): Txt体重 = NVL(rsTemp!体重)
    
    Call SeekIndex(cbo费别, NVL(rsTemp!费别), True)
    Call SeekIndex(cbo付款方式, NVL(rsTemp!医疗付款方式), True)
    Txt身份证号 = NVL(rsTemp!身份证号)
    Call SeekIndex(cbo民族, NVL(rsTemp!民族), True)
    Call SeekIndex(cbo职业, NVL(rsTemp!职业), True)
    Call SeekIndex(cbo婚姻, NVL(rsTemp!婚姻状况), True)
    Txt电话 = NVL(rsTemp!电话): Txt邮编 = NVL(rsTemp!邮编)
    Txt联系地址 = NVL(rsTemp!地址)
    Label22.tag = NVL(rsTemp!合同单位ID, 0)
    
    txtPatientDept.Text = NVL(rsTemp!病人科室)
    txtPatientDept.tag = NVL(rsTemp!病人科室ID, 0)
    txtID = NVL(rsTemp!病人号): txtBed = NVL(rsTemp!床号)
    
'    Call SeekIndex(cbo开单科室, Nvl(rsTemp!科室简码), True, , True)
    Call SeekIndex(cbo开单科室, NVL(rsTemp!开嘱科室), True, , TNeedType.tNeedName)

    Call SeekIndex(cbo医生1, NVL(rsTemp!开嘱医生), True)
    Call SeekIndex(cbo医生2, NVL(rsTemp!开嘱医生), True)
    '查找不到开嘱医生，且开嘱医生不为空，则直接填写开嘱医生字段
    
    If NVL(rsTemp!开嘱医生) <> "" And cbo医生1.ListIndex = -1 Then
        Me.cbo医生1.Visible = False
        Me.cbo医生2.Visible = True
        cbo医生2.Text = NVL(rsTemp!开嘱医生)
    End If
    
    chk紧急.value = NVL(rsTemp!紧急标志, 0)
    dtp(1).value = Format(curDate, "yyyy-mm-dd HH:MM")
        
    Call SeekIndex(cboRoom, NVL(rsTemp!执行间), True, , tNeedNo) '匹配执行间
    mstrRoom = NVL(rsTemp!执行间)
    
    txt附加主述.Text = NVL(rsTemp!附加主述)
    '医嘱内容　诊疗名称,床旁/术中:部位1(方法1),部位1(方法2),部位2(方法1)---
    txt医嘱内容 = Split(Split(rsTemp!医嘱内容, ":")(0), ",")(0)
    Call SeekIndex(cbo技师一, NVL(rsTemp!检查技师), True, True)
    Call SeekIndex(cbo技师二, NVL(rsTemp!检查技师二), True, True)
    
    mstrOutNo = NVL(rsTemp!门诊号, 0)
    mstrCardNo = NVL(rsTemp!就诊卡号)
    mstrCardPass = NVL(rsTemp!卡验证码)
    mintSourceType = rsTemp!病人来源
    mstrRegNo = NVL(rsTemp!挂号单)
    
    If mblnAllPatientIsOutside Then mintSourceType = 3
    
    mlngPatiId = NVL(rsTemp!病人ID, 0)
    mlngPageID = NVL(rsTemp!主页ID, 0)
    mstrItemType = NVL(rsTemp!影像类别)
    mlngClinicID = NVL(rsTemp!诊疗项目ID)
    
    If mstrItemType = "" Then
        MsgBoxD Me, "本次检查项目未加入影像检查项目,请检查", vbInformation, gstrSysName
        mlngResultState = 0
        cmdOK.Enabled = False
        Exit Function
    End If
    
    '显示送检单位和送检医生信息
    If mblnInputOutInfo Then
        gstrSQL = "select 信息名,信息值 from 病人信息从表 where 病人ID=[1] and 就诊id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取外院病人信息", mlngPatiId, mlngAdviceID)
        Do Until rsTemp.EOF
            If NVL(rsTemp!信息名) = "送检单位" Then txt送检单位.Text = NVL(rsTemp!信息值)
            If NVL(rsTemp!信息名) = "送检医生" Then txt送检医生.Text = NVL(rsTemp!信息值)
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "select 造影剂,用量,浓度 from 服用造影剂 where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", mlngAdviceID)
    If Not rsTemp.EOF Then
        Call SeekIndex(cbo造影剂, NVL(rsTemp!造影剂), True, , tNeedAll)
        Txt造影用量.Text = NVL(rsTemp!用量)
        Txt造影浓度.Text = NVL(rsTemp!浓度)
    End If

    txt医嘱内容.TabIndex = 0
    
    '复制完之后，设置控件的可用性
    Call RefreshObjEnabled(3)
    
    CopyCheck = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function RefreshPatiInfor(bln报到 As Boolean) As Boolean
'功能:用于报到或修改时刷新病人
'bln报到=True，是报到，则部分信息可以直接使用默认信息
'bln报到=False,是修改，则信息应该全部使用数据库中的信息

Dim rsTemp As New ADODB.Recordset
Dim strSql As String
Dim rsBaby As New ADODB.Recordset
Dim lngPatientID As Long
Dim lngPageID As Long
Dim intChargeState As ChargeState
Dim intChargeType As Integer    '病人医嘱发送.记录性质---1-收费记录；2-记帐记录。
Dim curDate As Date
Dim intTMP As Integer
Dim strTmp As String

    On Error GoTo errHand
    
    RefreshPatiInfor = False
    
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "SELECT nvl(E.姓名,A.姓名) 姓名,nvl(E.性别,A.性别) 性别,nvl(E.年龄,A.年龄) 年龄,B.出生日期,B.入院时间,B.费别,B.医疗付款方式,B.身份证号,B.民族,B.职业,Nvl(E.英文名,'') 英文名,E.身高,E.体重" & _
                    ",B.婚姻状况,Nvl(B.家庭电话,B.联系人电话) 电话,Nvl(B.家庭地址邮编,B.单位邮编) 邮编,nvl(B.家庭地址,B.工作单位) 地址,B.合同单位ID,B.门诊号,B.就诊卡号,B.卡验证码" & _
                    ",NVL(D.名称,'') AS 病人科室,A.病人科室ID,Decode(A.病人来源,2,B.住院号,B.门诊号) As 病人号,Decode(B.住院号,NULL,NULL,B.当前床号) As 床号,B.当前病区ID" & _
                    ",F.发送时间 开嘱时间,NVL(C.简码,0) 科室简码,NVL(C.名称,'未知') AS 开嘱科室,A.开嘱时间 As 申请时间,A.开嘱医生,A.紧急标志,F.首次时间,F.执行间,E.检查设备,A.医嘱内容,E.检查号,E.检查技师,F.执行部门ID" & _
                    ",DECODE(A.病人来源,2,2,1,1,4,4,3) AS 病人来源,Nvl(E.影像类别,G.影像类别) As 影像类别,B.病人id,A.主页id,A.诊疗项目ID,E.附加主述,Nvl(A.婴儿, 0) As 婴儿" & _
                    ",F.记录性质 " & _
                " FROM 病人医嘱发送 F,病人医嘱记录 A, 病人信息 B,部门表 C,部门表 D,影像检查记录 E,影像检查项目 G " & _
                " Where F.医嘱ID=[1] And F.发送号=[2] AND F.医嘱ID=A.ID" & _
                        " AND F.医嘱ID=E.医嘱ID(+) And F.发送号=E.发送号(+)  And A.病人ID=B.病人ID" & _
                        " And A.开嘱科室ID=C.ID And A.病人科室ID=D.ID And A.诊疗项目ID=G.诊疗项目ID(+)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", mlngAdviceID, mlngSendNo)

    If rsTemp.EOF Then
        '检查病人信息不完整的原因，如果是没有“病人医嘱发送记录，则提示本次医嘱已被回退或作废
        gstrSQL = "Select 医嘱ID From 病人医嘱发送 Where 医嘱ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查医嘱状态", mlngAdviceID)
        If rsTemp.EOF Then
            Call MsgBoxD(Me, "本次检查医嘱没有发送记录，可能是该医嘱已经被回退或者已作废，请刷新后检查医嘱状态！", vbInformation, gstrSysName)
        Else
            Call MsgBoxD(Me, "病人信息不完整，请与管理员联系！", vbInformation, gstrSysName)
        End If
    
        mlngResultState = 0
        cmdOK.Enabled = False
        Exit Function
    End If
    
    '处理婴儿信息
    mlngBaby = rsTemp!婴儿
    If mlngBaby = 0 Then
Normal:
        PatiIdentify.Text = NVL(rsTemp!姓名)
        Call SeekIndex(cbo性别, NVL(rsTemp!性别), True)
        
        If bln报到 Or mintEditMode = 1 Then
            If bln报到 And IsNull(rsTemp!出生日期) Then
  
                intTMP = MsgBoxD(Me, "出生日期为空，是否根据年龄推算？", vbYesNo, gstrSysName)
                
                If intTMP = vbYes Then
                    txt年龄.Text = NVL(rsTemp!年龄, "")
                    If txt年龄.Text = "" Then
                        MsgBoxD Me, "无年龄和出生日期数据，不能进行报到，请联系对应人员处理。", vbInformation, gstrSysName
                        RefreshPatiInfor = False
                        Exit Function
                    End If
                Else
                    MsgBoxD Me, "出生日期为空，不能进行报到，请联系对应人员处理。", vbInformation, gstrSysName
                    RefreshPatiInfor = False
                    Exit Function
                End If
            End If
        End If
        
        If NVL(rsTemp!年龄) <> "" Then
            LoadOldData rsTemp!年龄, txt年龄, cboAge
        Else
            '没有年龄情况下的处理方式
            Call ReCalcOld(Format(NVL(rsTemp!出生日期, curDate), "yyyy-mm-dd hh:mm:ss"), _
                        cboAge, _
                        0, _
                        Format(NVL(rsTemp!申请时间, curDate), "yyyy-mm-dd hh:mm:ss"))
        End If
 
        
        dtp出生日期.CustomFormat = "yyyy-MM-dd"
        
        If Trim(NVL(rsTemp!出生日期)) = "" Then
            Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
        Else
            '判断出生日期是否满一年，不满一年按则可能是婴儿
            If zlDatabase.Currentdate - CDate(rsTemp!出生日期) >= 366 Then
                dtp出生日期.value = Format(NVL(rsTemp!出生日期), "yyyy-mm-dd")
            Else
                dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
                dtp出生日期.value = Format(NVL(rsTemp!出生日期), "yyyy-mm-dd hh:mm:ss")
            End If
        End If
    Else
        lngPatientID = rsTemp!病人ID
        lngPageID = NVL(rsTemp!主页ID, 0)
        
        strSql = "Select A.开嘱时间 As 申请时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间, B.死亡时间 " & vbNewLine & _
                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                 "  Where a.病人ID = b.病人ID And b.主页id = [2] And b.序号 = [3] And a.ID = [1]"

        Set rsBaby = zlDatabase.OpenSQLRecord(strSql, "提取婴儿信息", mlngAdviceID, lngPageID, mlngBaby)
            
        If rsBaby.EOF Then
            GoTo Normal
        Else
            PatiIdentify.Text = NVL(rsBaby!婴儿姓名)
            Call SeekIndex(cbo性别, NVL(rsBaby!婴儿性别), True)
            
            If bln报到 Or mintEditMode = 1 Then
                If bln报到 And IsNull(rsBaby!出生时间) Then
                    MsgBoxD Me, "婴儿无出生日期数据，不能进行报到，请联系对应人员处理。", vbInformation, gstrSysName
                    RefreshPatiInfor = False
                    Exit Function
                End If
            End If
            
            If mintEditMode > 2 Then
                '大于2表明是报到后修改,报到后修改，会优先查询影像检查记录中的年龄
                Call LoadOldData(NVL(rsTemp!年龄), txt年龄, cboAge)
            Else
                txt年龄.Text = ReCalcOld(Format(NVL(rsBaby!出生时间, curDate), "yyyy-mm-dd hh:mm:ss"), _
                                            cboAge, _
                                            0, _
                                            Format(NVL(rsBaby!死亡时间, NVL(rsBaby!申请时间, curDate)), "yyyy-mm-dd hh:mm:ss"))
            End If
            
            '如果是婴儿，出生日期需要显示时分秒
            dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
            dtp出生日期.value = Format(NVL(rsBaby!出生时间), "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    
    lblCash.tag = NVL(rsTemp!当前病区ID)
    Txt英文名 = Decode(NVL(rsTemp!英文名), "", zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter), rsTemp!英文名)
    If Trim(txt年龄) = "" Then txt年龄 = 0
    Txt身高 = NVL(rsTemp!身高): Txt体重 = NVL(rsTemp!体重)
    Call SeekIndex(cbo费别, NVL(rsTemp!费别), True)
    Call SeekIndex(cbo付款方式, NVL(rsTemp!医疗付款方式), True)
    Txt身份证号 = NVL(rsTemp!身份证号)
    Call SeekIndex(cbo民族, NVL(rsTemp!民族), True)
    Call SeekIndex(cbo职业, NVL(rsTemp!职业), True)
    Call SeekIndex(cbo婚姻, NVL(rsTemp!婚姻状况), True)
    Txt电话 = NVL(rsTemp!电话): Txt邮编 = NVL(rsTemp!邮编)
    Txt联系地址 = NVL(rsTemp!地址)
    Label22.tag = NVL(rsTemp!合同单位ID, 0)
    
    txtPatientDept.Text = NVL(rsTemp!病人科室)
    txtPatientDept.tag = NVL(rsTemp!病人科室ID, 0)
    txtID = NVL(rsTemp!病人号): txtBed = NVL(rsTemp!床号)
    dtp(0).value = Format(rsTemp!申请时间, "yyyy-mm-dd HH:MM")
'    Call SeekIndex(cbo开单科室, Nvl(rsTemp!科室简码), True, , True)
    Call SeekIndex(cbo开单科室, NVL(rsTemp!开嘱科室), True, , TNeedType.tNeedName)
    Call SeekIndex(cbo医生1, NVL(rsTemp!开嘱医生), True)
    Call SeekIndex(cbo医生2, NVL(rsTemp!开嘱医生), True)
    
    '查找不到开嘱医生，且开嘱医生不为空，则直接填写开嘱医生字段
    If NVL(rsTemp!开嘱医生) <> "" And cbo医生1.ListIndex = -1 Then
        Me.cbo医生1.Visible = False
        Me.cbo医生2.Visible = True
        cbo医生2.Text = NVL(rsTemp!开嘱医生)
    End If

    chk紧急.value = NVL(rsTemp!紧急标志, 0)
    dtp(1).value = Format(curDate, "yyyy-mm-dd HH:MM")
    
    strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & mlngCurDeptId & "\" & Me.Name, "当前执行间", "") '提取上次登记时的执行间
    mstrRoom = NVL(rsTemp!执行间)
    
    txt附加主述.Text = NVL(rsTemp!附加主述)
    '医嘱内容　诊疗名称,床旁/术中:部位1(方法1),部位1(方法2),部位2(方法1)---
    txt医嘱内容 = Split(Split(rsTemp!医嘱内容, ":")(0), ",")(0)
    txt医嘱内容.tag = txt医嘱内容.Text
    If InStr(NVL(rsTemp!医嘱内容, ""), ":") > 0 Then
        Txt部位方法 = Replace(Split(rsTemp!医嘱内容, ":")(1), "),", ")" & vbCrLf)
    Else
        Txt部位方法 = NVL(rsTemp!医嘱内容, "")
    End If
    txt检查号.Text = CStr(NVL(rsTemp!检查号)): txt检查号.tag = txt检查号.Text
    
    '如果是修改病人 则刷新变量的值
    If mintEditMode = 3 Then mstrNextCheckNo = CStr(NVL(rsTemp!检查号))
    
    Call SeekIndex(cbo技师一, NVL(rsTemp!检查技师), True, True)
    
    mstrOutNo = NVL(rsTemp!门诊号, 0)
    mstrCardNo = NVL(rsTemp!就诊卡号)
    mstrCardPass = NVL(rsTemp!卡验证码)
    mintSourceType = rsTemp!病人来源
    mlngPatiId = NVL(rsTemp!病人ID, 0)
    mlngPageID = NVL(rsTemp!主页ID, 0)
    mstrItemType = NVL(rsTemp!影像类别)
    mlngClinicID = NVL(rsTemp!诊疗项目ID)
    mlngCurDeptId = Val(NVL(rsTemp!执行部门ID))
    
    intChargeType = NVL(rsTemp!记录性质, 1)
    
    If mstrItemType = "" Then
        MsgBoxD Me, "本次检查项目未加入影像检查项目,请检查", vbInformation, gstrSysName
        mlngResultState = 0
        cmdOK.Enabled = False
        Exit Function
    End If
    
    '显示送检单位和送检医生信息
    If mblnInputOutInfo Then
        gstrSQL = "select 信息名,信息值 from 病人信息从表 where 病人ID=[1] and 就诊id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取外院病人信息", mlngPatiId, mlngAdviceID)
        Do Until rsTemp.EOF
            If NVL(rsTemp!信息名) = "送检单位" Then txt送检单位.Text = NVL(rsTemp!信息值)
            If NVL(rsTemp!信息名) = "送检医生" Then txt送检医生.Text = NVL(rsTemp!信息值)
            rsTemp.MoveNext
        Loop
    End If
    
    '根据病人类型配置姓名文本框颜色
    If mblnNameColColorCfg Then
        If mintSourceType = 2 Then
            gstrSQL = "select 病人类型 from 病案主页 where 病人id=[1] and 主页id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人类型", mlngPatiId, mlngPageID)
        Else
            gstrSQL = "select 病人类型 from 病人信息 where 病人id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人类型", mlngPatiId)
        End If
        
        If rsTemp.RecordCount > 0 Then
            If mstrDefaultPatientType = NVL(rsTemp!病人类型) Then
                If mblnOrdinaryNameColColorCfg Then
                    PatiIdentify.objTxtInput.ForeColor = zlDatabase.GetPatiColor(NVL(rsTemp!病人类型))
                End If
            Else
                PatiIdentify.objTxtInput.ForeColor = zlDatabase.GetPatiColor(NVL(rsTemp!病人类型))
            End If
        End If
    End If
    
    gstrSQL = "select 造影剂,用量,浓度 from 服用造影剂 where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", mlngAdviceID)
    If Not rsTemp.EOF Then
        Call SeekIndex(cbo造影剂, NVL(rsTemp!造影剂), True, , tNeedAll)
        Txt造影用量.Text = NVL(rsTemp!用量)
        Txt造影浓度.Text = NVL(rsTemp!浓度)
    End If
    
    gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人附件", mlngAdviceID)
    Txt部位方法 = Txt部位方法 & vbCrLf
    Do Until rsTemp.EOF
        Txt部位方法 = Txt部位方法 & rsTemp!项目 & ":" & NVL(rsTemp!内容) & vbCrLf
        rsTemp.MoveNext
    Loop
    
    If mintEditMode = 2 Then
        txt检查号.Text = Next检查号: txt检查号.tag = txt检查号.Text
    End If
    
    '报道时，"无"/"欠"费情况下，不允许报道，除非有"未缴费报道"权限，"退"/"销"费在任何情况下，都不允许报道，
    '"调"允许进行报道（目前可能不会出现该状态，需要后续支持分部位执行时才会出现）
    intChargeState = CheckChargeState(mlngAdviceID, mintSourceType)
    
    If intChargeState = ChargeState.未收费 Then
        lblCash.Caption = "未收"
    ElseIf intChargeState = ChargeState.已收费 Then
        lblCash.Caption = "已收"
    ElseIf intChargeState = ChargeState.无费用 Then
        lblCash.Caption = "无费"
    ElseIf intChargeState = ChargeState.已记账 Then
        lblCash.Caption = "记账"
    ElseIf intChargeState = ChargeState.已退费 Then
        lblCash.Caption = "退费"
    ElseIf intChargeState = ChargeState.已销账 Then
        lblCash.Caption = "销账"
    ElseIf intChargeState = ChargeState.已调整 Then
        lblCash.Caption = "调"
    Else
        lblCash.Caption = ""
    End If
    
    Call RefreshObjEnabled
    
    If bln报到 And Not CheckPopedom(mstrPrivs, "未缴费报到") And mintSourceType <> 3 Then  '24361 有权限不判断，自行登记不控制，急诊也进行判断
        If lblCash.Caption = "已收" Or lblCash.Caption = "无费" _
            Or (gbln执行后审核 And (intChargeType = ChargeState.无费用 Or intChargeState = ChargeState.已记账)) _
            Or gbln执行前先结算 Then
            ''需要根据系统参数判断， gbln执行后审核=81号参数是"执行后自动审核划价单",勾选这个参数后，没有未交费报告权限时，也应该可以对记账记录进行报到
            ''gbln执行前先结算 = 163--门诊一卡通，项目执行前必须先收费或先记帐审核,勾选这个参数后，没有未交费报告权限时，也应该可以进行报到，报到的时候会刷卡消费
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If

        If cmdOK.Enabled = False Then
            Me.Caption = Me.Caption & "(当前病人未收费，不能报到)"
        End If
    End If
    
    If lblCash.Caption = "退费" And bln报到 Then
        cmdOK.Enabled = False
        Me.Caption = Me.Caption & "(当前病人已退费，不能报到)"
    End If
    
    If lblCash.Caption = "销账" And bln报到 Then
        cmdOK.Enabled = False
        Me.Caption = Me.Caption & "(当前病人已销账，不能报到)"
    End If
    
    RefreshPatiInfor = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function GetDefaultRoom(ByVal lngAdviceID As Long, ByVal lngExecuteDeptId As Long) As String
'获取默认的检查执行间
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetDefaultRoom = ""
    
'    strSQL = "select distinct 执行间,检查设备 from 医技执行房间 a, 影像执行分组 b, 影像分组关联 c, 病人医嘱记录 d " & _
'            " Where a.分组id = b.ID And b.ID = c.分组id And c.诊疗项目id = d.诊疗项目id " & _
'            " and d.id=[1] and a.科室id=[2]"

    '查询检查项目对应的执行间，并根据执行间的当前检查数量排序
    strSql = "select /*+ RULE*/ 执行间,数量 " & _
            " from ( " & _
            " with tmp as( " & _
            "   select 1 as tag, x.执行间,x.执行过程 " & _
            "   from 病人医嘱发送 x, 病人医嘱记录 y " & _
            "   Where X.医嘱ID = Y.ID And Y.相关id Is Null And X.执行部门id = [2] " & _
            "     and x.首次时间 between sysdate-10 and sysdate and x.执行过程=2 " & _
            " ) " & _
            " select  a.执行间,a.检查设备, count( f.tag)  as 数量 " & _
            " from 医技执行房间 a, 影像执行分组 b, 影像分组关联 c, 病人医嘱记录 d , tmp f " & _
            " Where a.分组id = b.ID And b.ID = c.分组id And c.诊疗项目id = d.诊疗项目id " & _
            "   and a.执行间=f.执行间(+) and c.科室id=[2] " & _
            "   and a.科室id=[2] and d.id=[1] " & _
            " group by a.执行间,a.检查设备) order by 数量"
        
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询默认检查执行间", lngAdviceID, lngExecuteDeptId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetDefaultRoom = NVL(rsData!执行间)
End Function


Private Sub CmdCancle_Click()
    mlngResultState = IIf(mlngGoOnReg = 1, 4, 0)
    Unload Me
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'功能：检查输入数据的合法性
'参数： 无
'返回：True--数据输入合格，可以继续；False --有数据输入不合格，需要修改数据
'------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lngResult As Long
    
    ValidData = False
    
    gstrSQL = "select ID ,科室ID,参数值 from 影像流程参数 where 科室ID = [1] and 参数名 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurDeptId, CStr("必录控制"))
    If Not rsTemp.EOF Then
        If NVL(rsTemp!参数值) <> "" Then
            If InStr(rsTemp!参数值, "英文名") > 0 And Trim(Txt英文名) = "" And Txt英文名.Enabled = True Then
                MsgBoxD Me, "必须输入英文名，请检查！", vbInformation, gstrSysName: DoEvents
                Txt英文名.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "性别") > 0 And Trim(cbo性别.Text) = "" And cbo性别.Enabled = True Then
                MsgBoxD Me, "必须输入性别，请检查！", vbInformation, gstrSysName: DoEvents
                cbo性别.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "年龄") > 0 And Trim(txt年龄) = "" And txt年龄.Enabled = True Then
                MsgBoxD Me, "必须输入年龄，请检查！", vbInformation, gstrSysName: DoEvents
                txt年龄.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "出生日期") > 0 And Trim(dtp出生日期.value) = "" And dtp出生日期.Enabled = True Then
                MsgBoxD Me, "必须输入出生日期，请检查！", vbInformation, gstrSysName: DoEvents
                dtp出生日期.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "身高") > 0 And Trim(Txt身高) = "" And Txt身高.Enabled = True Then
                MsgBoxD Me, "必须输入身高，请检查！", vbInformation, gstrSysName: DoEvents
                Txt身高.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "体重") > 0 And Trim(Txt体重) = "" And Txt体重.Enabled = True Then
                MsgBoxD Me, "必须输入体重，请检查！", vbInformation, gstrSysName: DoEvents
                Txt体重.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "费别") > 0 And Trim(cbo费别.Text) = "" And cbo费别.Enabled = True Then
                MsgBoxD Me, "必须输入费别，请检查！", vbInformation, gstrSysName: DoEvents
                cbo费别.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "付款方式") > 0 And Trim(cbo付款方式.Text) = "" And cbo付款方式.Enabled = True Then
                MsgBoxD Me, "必须输入付款方式，请检查！", vbInformation, gstrSysName: DoEvents
                cbo付款方式.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "身份证号") > 0 And Trim(Txt身份证号) = "" And Txt身份证号.Enabled = True Then
                MsgBoxD Me, "必须输入身份证号，请检查！", vbInformation, gstrSysName: DoEvents
                Txt身份证号.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "民族") > 0 And Trim(cbo民族.Text) = "" And cbo民族.Enabled = True Then
                MsgBoxD Me, "必须输入民族，请检查！", vbInformation, gstrSysName: DoEvents
                cbo民族.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "职业") > 0 And Trim(cbo职业.Text) = "" And cbo职业.Enabled = True Then
                MsgBoxD Me, "必须输入职业，请检查！", vbInformation, gstrSysName: DoEvents
                cbo职业.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "婚姻") > 0 And Trim(cbo婚姻.Text) = "" And cbo婚姻.Enabled = True Then
                MsgBoxD Me, "必须输入婚姻，请检查！", vbInformation, gstrSysName: DoEvents
                cbo婚姻.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "电话") > 0 And Trim(Txt电话) = "" And Txt电话.Enabled = True Then
                MsgBoxD Me, "必须输入电话，请检查！", vbInformation, gstrSysName: DoEvents
                Txt电话.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "邮编") > 0 And Trim(Txt邮编) = "" And Txt邮编.Enabled = True Then
                MsgBoxD Me, "必须输入邮编，请检查！", vbInformation, gstrSysName: DoEvents
                Txt邮编.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "地址") > 0 And Trim(Txt联系地址) = "" And Txt联系地址.Enabled = True Then
                MsgBoxD Me, "必须输入联系地址，请检查！", vbInformation, gstrSysName: DoEvents
                Txt联系地址.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "执行间") > 0 And Trim(cboRoom.Text) = "" And cboRoom.Enabled = True Then
                MsgBoxD Me, "必须输入执行间，请检查！", vbInformation, gstrSysName: DoEvents
                cboRoom.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "造影剂") > 0 And Trim(cbo造影剂.Text) = "" And cbo造影剂.Enabled = True Then
                MsgBoxD Me, "必须输入造影剂，请检查！", vbInformation, gstrSysName: DoEvents
                cbo造影剂.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "检查技师") > 0 And Trim(cbo技师一.Text) = "" And cbo技师一.Enabled = True Then
                MsgBoxD Me, "必须输入检查技师，请检查！", vbInformation, gstrSysName: DoEvents
                cbo技师一.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "检查技师二") > 0 And Trim(cbo技师一.Text) = "" And cbo技师二.Enabled = True Then
                MsgBoxD Me, "必须输入检查技师二，请检查！", vbInformation, gstrSysName: DoEvents
                cbo技师二.SetFocus: Exit Function
            ElseIf InStr(rsTemp!参数值, "附加主述") > 0 And Trim(txt附加主述.Text) = "" And txt附加主述.Enabled = True Then
                MsgBoxD Me, "必须输入附加主述，请检查！", vbInformation, gstrSysName: DoEvents
                txt附加主述.SetFocus: Exit Function
            End If
        End If
    End If

    On Error Resume Next
    
    '检查输入的年龄是否有效

    If txt年龄.Enabled Then
        If mobjPublicPatient Is Nothing Then
            If MsgBoxD(Me, "未检测到部件zlPublicPatient.dll的有效注册信息，不能对年龄的有效性进行检查，是否继续？", vbYesNo + vbExclamation) = vbNo Then
                Exit Function
            End If
        Else
'            txt年龄.tag 保存查找时所对应的入院或登记时间
            If txt年龄.tag <> "" Then
                If Not mobjPublicPatient.CheckPatiAge( _
                                                    txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, ""), _
                                                    dtp出生日期.value, 0, _
                                                    txt年龄.tag) Then Exit Function
            Else
                If Not mobjPublicPatient.CheckPatiAge( _
                                                    txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, ""), _
                                                    dtp出生日期.value, 0, _
                                                    dtp(0).value) Then Exit Function
    
            End If
        End If
    End If
    
    If Len(Trim(Me.txt医嘱内容.tag)) = 0 Then
        MsgBoxD Me, "必须输入申请项目！", vbInformation, gstrSysName: DoEvents
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Me.cbo开单科室.ListIndex = -1 Then
        MsgBoxD Me, "请指定申请科室！", vbInformation, gstrSysName: DoEvents
        Me.cbo开单科室.SetFocus: Exit Function
    End If
    
    If cbo医生1.Visible Then
        If Len(Trim(Me.cbo医生1.Text)) = 0 Then
            MsgBoxD Me, "请指定申请医生！", vbInformation, gstrSysName: DoEvents
            Me.cbo医生1.SetFocus: Exit Function
        End If
    Else
        If Len(Trim(Me.cbo医生2.Text)) = 0 Then
            MsgBoxD Me, "请指定申请医生！", vbInformation, gstrSysName: DoEvents
            Me.cbo医生2.SetFocus: Exit Function
        End If
    End If
   
    '问题号：76509
'    If dtp(0).value > dtp(1).value Then
'        MsgBoxD Me, "申请时间不能大于检查时间！", vbInformation, gstrSysName: DoEvents
'        Me.dtp(0).SetFocus: Exit Function
'    End If
    
    If Len(Trim(Me.PatiIdentify.Text)) = 0 And PatiIdentify.objTxtInput.Enabled Then
        MsgBoxD Me, "请输入病人姓名！", vbInformation, gstrSysName: DoEvents
        Me.PatiIdentify.SetFocus
        Exit Function
    End If
    
    If Trim(Txt英文名) = "" And Txt英文名.TabStop And Txt英文名.Enabled Then
        MsgBoxD Me, "英文名不能为空！", vbInformation, gstrSysName: DoEvents
        Txt英文名.SetFocus
        Exit Function
    End If

    If mintEditMode >= 2 Or (mintEditMode = 0 And mblnRegToCheck) Then '报到,或报到后修改　或　登记后直接检查 (登记时或登记后修改不判断)
        If Len(Trim(Me.txt检查号)) = 0 And txt检查号.Enabled Then
            If MsgBoxD(Me, "检查号不能为空，是否自动生成检查号？", vbYesNo, gstrSysName) = vbYes Then
                txt检查号.Text = Next检查号
            End If
            
            txt检查号.SetFocus
            Exit Function
        End If
        
        '判断检查号长度
        If LenB(StrConv(txt检查号.Text, vbFromUnicode)) > 64 Then
            MsgBoxD Me, "检查号长度超过64个字符，请重新输入。", vbInformation, gstrSysName: DoEvents
            txt检查号.SetFocus
        End If
    
        '判断检查号的递增情况，递增超过10的则提示
        If Val(txt检查号.Text) > Val(mstrNextCheckNo) + 10 Then
            If MsgBoxD(Me, "检查号过大，比当前的最大号码大了" & (Val(txt检查号.Text) - Val(mstrNextCheckNo)) & "，请确认是否继续" & IIf(mintEditMode = 3, "修改", "报到") & "？", vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                txt检查号.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '报到或登记后直接检查的时候，必须要输入执行间
    If mintEditMode = 2 Or mintEditMode = 3 Or (mblnRegToCheck And mintEditMode = 0) Then
        If cboRoom.Text = "" And Not mblnUsePacsQueue Then
            MsgBoxD Me, "执行间不能为空！", vbInformation, gstrSysName: DoEvents
            cboRoom.SetFocus
            Exit Function
        End If
    End If
    
    If mblnUsePacsQueue Then
        '判断执行间是否允许进行变更
        If mstrRoom <> "" And mstrRoom <> zlStr.NeedCode(cboRoom.Text) Then
            lngResult = MsgBoxD(Me, "执行间修改后，将重新进行排队，是否继续？", vbYesNo, "提示")
            If lngResult = vbNo Then
                cboRoom.SetFocus
                Exit Function
            End If
        End If
        
        If cboRoom.Text = "呼叫时指派" Then
            '当执行间为呼叫时指派时，需要判断当前检查项目是否设置了对应的执行分组，如果没有设置，则不允许指派
            If Not IsConfigStudyGroup(mlngClinicID) Then
                MsgBoxD Me, "该检查项目为设置对应的执行分组，因此不能在呼叫时指派执行间。", vbInformation, gstrSysName: DoEvents
                cboRoom.SetFocus
                Exit Function
            End If
        End If
    End If
    
    ValidData = True
End Function


Private Function IsConfigStudyGroup(ByVal lng诊疗项目ID As Long) As Boolean
'判断该检查是否设置分组
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    IsConfigStudyGroup = False
    
    '如果排队叫号的号码生成方式是按照科室排队，则不需要判断项目的执行分组
    If mlngPacsQueueNoWay <> 0 Then
        strSql = "select 组名 from 影像执行分组 a, 影像分组关联 b " & _
                " where  a.id = b.分组Id and b.诊疗项目id=[1] and b.科室ID=[2]"
        
        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询检查分组", lng诊疗项目ID, mlngCurDeptId)
        If rsData.RecordCount <= 0 Then Exit Function
    End If
    
    IsConfigStudyGroup = True
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            zlCommFun.PressKey vbKeyTab
        Case vbKeyF2
            If mintEditMode <> 1 Then CmdOK_Click   '登记和修改都用F2
        Case vbKeyF4
            If mintEditMode = 1 Then CmdOK_Click   '报到用F4
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSquareCard = Nothing
    
    Set mobjInsure = Nothing
    Set mobjPublicPatient = Nothing
    
    If mintEditMode = 2 Or mintEditMode = 3 Or mblnRegToCheck Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & mlngCurDeptId & "\" & Me.Name, "当前执行间", zlStr.NeedCode(cboRoom)
    End If
    
    If mintEditMode > 1 Or mblnRegToCheck Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查技师一", zlStr.NeedName(cbo技师一.Text)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查技师二", zlStr.NeedName(cbo技师二.Text)
    End If
    
    
        '这是判断登记时扫描后 点击取消按钮 扫描窗体释放
    If Not frmPetitionCap Is Nothing Then
        frmPetitionCap.mblnIsLogin = False
        Call frmPetitionCap.Form_Unload(0)
        Set frmPetitionCap = Nothing
    End If
    
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    'IDKind控件内部还会自己查询一次，如果使用“收费单据号”和“挂号单号”会查询失败，
    '查询失败后，会将输入的信息填入text，因此需要在这里强制将PACS查询出来的姓名填入ShowName
    ' 避免在控件内部自动修改显示的名称
    ShowName = PatiIdentify.Text
        Call SendKeys("{tab}")
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    If mintEditMode = 0 Then
        Call FindPatient(blnCard)
    End If
    strShowText = PatiIdentify.Text
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If PatiIdentify.Text <> "" Then PatiIdentify.Text = ""
    If PatiIdentify.objTxtInput.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub

Private Sub PatiIdentify_LostFocus()
    Txt英文名.Text = zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter)
    
    Call zlCommFun.OpenIme
End Sub

Private Sub Txt电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt附加主述_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt附加主述_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt检查号_Change()
    Dim lngSet As Long
    '判断是否包含小写字符，如果包含，则提示用户，并自动改成大写
    If txt检查号.Text <> UCase(txt检查号.Text) Then
        lngSet = txt检查号.SelStart
        txt检查号.Text = UCase(txt检查号.Text)
        txt检查号.SelStart = lngSet
    End If
    
    If LenB(StrConv(txt检查号.Text, vbFromUnicode)) > 64 Then
        MsgBoxD Me, "检查号长度超过64个字符，请重新输入。", vbInformation, gstrSysName
        txt检查号.Text = Left(txt检查号.Text, Len(txt检查号.Text) - 1)
    End If
End Sub

Private Sub txt检查号_GotFocus()
    zlControl.TxtSelAll txt检查号
End Sub

Private Sub txt检查号_KeyDown(KeyCode As Integer, Shift As Integer)
    txt检查号.Locked = Not mblnChangeNo
End Sub

Private Sub txt年龄_Change()
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.Visible = True
    End If
End Sub

Private Sub txt年龄_GotFocus()
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            cboAge.SetFocus
        End If
        If Not IsNumeric(txt年龄.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt年龄_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not CheckOldData(txt年龄, cboAge) Then Exit Sub
    
    Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
End Sub


Public Function ReCalcBirthDay(ByVal strAge As String, ByVal strUnit As String) As String
'根据年龄算出出生日期
    Dim sreDateOfBirth As String
    
    On Error GoTo errHand
    
    If Not mobjPublicPatient Is Nothing Then
        Call mobjPublicPatient.ReCalcBirthDay(strAge & IIf(strUnit = "", "", strUnit), sreDateOfBirth)
    End If
    
    If Trim(sreDateOfBirth) <> "" Then dtp出生日期.value = sreDateOfBirth
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt年龄_Validate(Cancel As Boolean)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.ListIndex = 0: cboAge.Visible = True
    End If
End Sub

Private Sub Txt身高_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Call TxtInputControl(Txt身高, KeyAscii, 2)
End Sub

Private Sub Txt体重_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Call TxtInputControl(Txt体重, KeyAscii, 2)
End Sub

Private Sub FindPatient(blnCard As Boolean)
On Error GoTo err
    Dim rsTmp As ADODB.Recordset
    Dim rsAge As ADODB.Recordset
    
    Dim lngAge As Long
    Dim curDate As Date
    Dim strSql As String

    Set rsTmp = GetPatient(PatiIdentify.Text, blnCard) '根据输入提取病人信息
    
    txt年龄.tag = ""
    
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            If NVL(rsTmp!姓名) <> "新病人" Then
                curDate = zlDatabase.Currentdate
            
                PatiIdentify.tag = Trim(NVL(rsTmp!姓名))
                PatiIdentify.Text = Trim(NVL(rsTmp!姓名))
                Call SeekIndex(cbo性别, NVL(rsTmp!性别), True)
                
                dtp出生日期.CustomFormat = "yyyy-MM-dd"
                
                If NVL(rsTmp!出生日期) <> "" Then
                    '判断出生日期是否满一年，不满一年按则可能是婴儿
                    If zlDatabase.Currentdate - CDate(rsTmp!出生日期) >= 366 Then
                        dtp出生日期.value = Format(NVL(rsTmp!出生日期), "yyyy-mm-dd")
                    Else
                        dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
                        dtp出生日期.value = Format(NVL(rsTmp!出生日期), "yyyy-mm-dd hh:mm:ss")
                    End If
                Else
                    dtp出生日期.value = Format(NVL(rsTmp!出生日期, curDate), "yyyy-mm-dd")
                End If

                strSql = ""
                Set rsAge = Nothing
                
                If Val(NVL(rsTmp!来源ID, 0)) = 2 Then
                    '住院病人年龄提取
                    strSql = "select 年龄, 入院日期 As 时间 from 病案主页 where 病人ID=[1] and 主页ID=[2] and rownum <=1"
                    Set rsAge = zlDatabase.OpenSQLRecord(strSql, "提取患者年龄", Val(NVL(rsTmp!病人ID)), Val(NVL(rsTmp!主页ID)))
                Else
                    If Val(NVL(rsTmp!就诊状态, 0)) <> 0 Then
                        '门诊患者年龄提取
                        strSql = "select 年龄, 登记时间 As 时间 from 病人挂号记录 where 病人ID=[1] and 门诊号=[2] and rownum <=1"
                        Set rsAge = zlDatabase.OpenSQLRecord(strSql, "提取患者年龄", Val(NVL(rsTmp!病人ID)), Val(NVL(rsTmp!门诊号)))
                    End If
                End If
                
                If Not rsAge Is Nothing Then
                    If rsAge.RecordCount > 0 Then
                        LoadOldData NVL(rsAge!年龄), txt年龄, cboAge
                        txt年龄.tag = NVL(rsAge!时间)
                    End If
                Else
                    '根据当前日期计算年龄
                    LoadOldData ReCalcOld(dtp出生日期.value, cboAge, 0, dtp(0).value), txt年龄, cboAge
                End If
                
                If txt年龄.Text = "" Then
                    txt年龄 = 0
                    cboAge.Visible = True
                    cboAge.ListIndex = 0
                End If
                
                Call SeekIndex(cbo费别, NVL(rsTmp!费别, "普通"))
                Call SeekIndex(cbo付款方式, NVL(rsTmp!医疗付款方式, "自费医疗"))
                Txt身份证号 = NVL(rsTmp!身份证号)
                Call SeekIndex(cbo民族, NVL(rsTmp!民族, "汉族"))
                Call SeekIndex(cbo职业, NVL(rsTmp!职业, "工人"))
                Call SeekIndex(cbo婚姻, NVL(rsTmp!婚姻状况, "未婚"))
                Txt电话 = NVL(rsTmp!电话)
                Txt邮编 = NVL(rsTmp!邮编)
                Txt联系地址 = NVL(rsTmp!地址)
                Label22.tag = NVL(rsTmp!合同单位ID, 0)
                txtID = Decode(NVL(rsTmp!住院号), "", NVL(rsTmp!门诊号), NVL(rsTmp!住院号))
                txtBed = NVL(rsTmp!当前床号)
                
                If mInputType = 5 Then
                '根据收费单据提取信息
                    Call SeekIndex(cbo开单科室, getID_TO_简码(NVL(rsTmp!病人科室ID), "部门表"), True, , True)
                    Call SeekIndex(cbo医生1, NVL(rsTmp!医生))
                    Call SeekIndex(cbo医生2, NVL(rsTmp!医生))
                End If
                
                mlngPatiId = NVL(rsTmp!病人ID, 0)
                mintSourceType = NVL(rsTmp!来源ID, 1)
                
                '对于非住院病人，需区分是门诊还是外来
                If mintSourceType <> 2 Then mintSourceType = getSourceType(rsTmp!病人ID)
                
                mlngPageID = NVL(rsTmp!主页ID, 0)
                mstrOutNo = NVL(rsTmp!门诊号, 0)
                mstrCardNo = NVL(rsTmp!就诊卡号)
                mstrCardPass = NVL(rsTmp!卡验证码)
                
                '显示病人科室
                txtPatientDept.Text = zlStr.NeedName(cbo开单科室)
                txtPatientDept.tag = NVL(rsTmp!病人科室ID)
                If cbo性别.Enabled = True Then cbo性别.SetFocus
                
                Call RefreshObjEnabled(2)
                
                '提取病人信息完成后 自动反算病人出生日期
                If IsNumeric(txt年龄.Text) And NVL(rsTmp!出生日期, "") = "" Then Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
    
                Exit Sub
            Else
                If cbo性别.Enabled = True And mblnIsSamePatient Then cbo性别.SetFocus
            End If
        End If
    End If
    
    '没查到按新登记病人算
    Dim strTmp As String
    strTmp = Trim(PatiIdentify.Text)
    
    If PatiIdentify.IDKindIDX <> PatiIdentify.GetKindIndex(IDKind_身份证号) Then '身份证读取帖身份证触发函数填写姓名等信息
        If PatiIdentify.Text <> strTmp Then PatiIdentify.Text = strTmp
        PatiIdentify.tag = Trim(PatiIdentify.Text)
        Txt英文名.Text = zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter)
    End If
    mlngPatiId = 0
    mintSourceType = 3
    mlngPageID = 0
    '刷卡，而且没有提取到病人信息，依然选择txt姓名
    If blnCard Then PatiIdentify.SetFocus

    Call RefreshObjEnabled(1)
    Exit Sub
err:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Function getSourceType(ByVal lngPatiID As Long) As Integer
'功能:获取病人来源和挂号单
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If mInputType = 4 Then
        getSourceType = 1
        Exit Function '为挂号单时，确认为门诊病人
    End If
    
    '缺省为外院病人
    getSourceType = 3
    
    strSql = "select NO from 病人挂号记录 where 病人ID=[1] and 执行状态<>-1 and 执行状态<>1  order by 登记时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取病人来源和挂号单", lngPatiID)
    
    If rsTemp.RecordCount > 0 Then
        getSourceType = 1
        mstrRegNo = NVL(rsTemp!NO)
    End If
End Function

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
Dim rsTmp As ADODB.Recordset
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        With txt医嘱内容
            If .Text = "" Then Call cmdSel_Click
            If Trim(.Text) = .tag Then Exit Sub
            
            Set rsTmp = SelectDiagItem() '提取项目
            If rsTmp Is Nothing Then '取消或无数据
                '恢复原值
                .Text = .tag
                zlControl.TxtSelAll txt医嘱内容
                .SetFocus
                Exit Sub
            Else
                If AdviceInput(rsTmp) Then '根据选择项目设置部位及方法
                    .tag = .Text
                Else '取消部位及方法
                    .Text = .tag
                    zlControl.TxtSelAll txt医嘱内容
                    .SetFocus
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    '恢复人为的改变,回车时赋值
    If txt医嘱内容.Text <> txt医嘱内容.tag Then
        txt医嘱内容.Text = txt医嘱内容.tag
    End If
End Sub

Private Sub Txt英文名_LostFocus()
    zlControl.TxtSelAll Txt英文名
End Sub

Private Sub Txt邮编_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub cbo开单科室_Click()
    '判断选择科室 是否是外院科室
    If InStr(cbo开单科室.Text, "外院") > 0 Then
        mblnIsOutSideHosp = True
        
        cbo医生1.Visible = False
        cbo医生2.Visible = True
    Else
        mblnIsOutSideHosp = False
    
        cbo医生1.Visible = True
        cbo医生2.Visible = False
    End If

    If cbo开单科室.ListIndex > -1 Then InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub

Private Sub txt医嘱内容_GotFocus()
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub Txt联系地址_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Txt联系地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub PatiIdentify_Change()
    '只有登记的时候，提取了病人，再修改姓名，才会变更成新病人
    If mintEditMode = 0 And mlngPatiId <> 0 And PatiIdentify.Text <> "" Then
        MsgBoxD Me, "病人修改姓名后，就作为新病人处理了。", vbOKOnly, "提示信息"
        Call ClearFaceData(True)
        Call InitEdit(False)
        mlngPatiId = 0
        Call FindPatient(False)
    End If
End Sub

Private Sub PatiIdentify_GotFocus()
    Call zlCommFun.OpenIme(gstrIme <> "不自动开启")
End Sub

Private Sub PatiIdentify_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long
    Dim strExpand As String
    Dim strOutCardNO As String
    Dim strOutPatiInfoXML As String
    
    lng卡类别ID = Val(PatiIdentify.GetCurCard.接口序号)

    If lng卡类别ID = 0 Then Exit Sub
    If mobjSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInfoXML) = False Then
        Exit Sub
    End If
    PatiIdentify.Text = strOutCardNO
    If PatiIdentify.Text <> "" Then
        Call FindPatient(False)
    End If
End Sub

Private Sub PatiIdentify_Validate(Cancel As Boolean)
    Select Case PatiIdentify.IDKindIDX
        Case PatiIdentify.GetKindIndex(IDKind_IC卡号)
            PatiIdentify.objTxtInput.ToolTipText = "IC卡识别"
        Case PatiIdentify.GetKindIndex(IDKind_姓名)
            PatiIdentify.objTxtInput.ToolTipText = "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
        Case PatiIdentify.GetKindIndex(IDKind_医保号)
            PatiIdentify.objTxtInput.ToolTipText = "请录入医保号"
        Case PatiIdentify.GetKindIndex(IDKind_身份证号)
            PatiIdentify.objTxtInput.ToolTipText = "请将身份证置于读卡器上"
    End Select
End Sub

Private Sub cboRoom_Click()
    If zlStr.NeedName(cboRoom.list(cboRoom.ListIndex)) <> "" Then
        Call SeekIndex(cboDevice, zlStr.NeedName(cboRoom.list(cboRoom.ListIndex)), True, , tNeedNo)
    Else
        cboDevice.ListIndex = -1
    End If
End Sub


Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo费别.hWnd, zlControl.CboMatchIndex(cbo费别.hWnd, KeyAscii))
End Sub

Private Sub cbo付款方式_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo付款方式.hWnd, zlControl.CboMatchIndex(cbo付款方式.hWnd, KeyAscii))
End Sub

Private Sub cbo婚姻_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo婚姻.hWnd, zlControl.CboMatchIndex(cbo婚姻.hWnd, KeyAscii))
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo开单科室.hWnd, zlControl.CboMatchIndex(cbo开单科室.hWnd, KeyAscii))
    
    If KeyAscii = vbKeyReturn Then
        Call cbo开单科室_Click
    End If
End Sub

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo民族.hWnd, zlControl.CboMatchIndex(cbo民族.hWnd, KeyAscii))
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo性别.hWnd, zlControl.CboMatchIndex(cbo性别.hWnd, KeyAscii))
End Sub
Private Sub cbo医生1_KeyPress(KeyAscii As Integer)
    '如果开单科室选择的是 外院科室，那么跳过医生的简码查找功能，否则医生栏不能自由录入
    If Not mblnIsOutSideHosp Then
        Call zlControl.CboSetIndex(cbo医生1.hWnd, zlControl.CboMatchIndex(cbo医生1.hWnd, KeyAscii))
    End If
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo职业.hWnd, zlControl.CboMatchIndex(cbo职业.hWnd, KeyAscii))
End Sub

Private Function CheckNoValidate() As Boolean
'------------------------------------------------
'功能：判断检查号是否重复，如果重复了，是否可以继续，
'       其中使用了几个流程管理参数辅助判断
'       1、mlngBuildType---1-按科室递增，0-按影像类别递增
'       2、mlngUnicode --- 患者检查号保持不变,1-保持检查号不变；0-检查号流水递增
'       3、mblnCanOverWrite --- 允许检查号重复
'参数： 无
'返回：True--继续报到；False --停止报到
'------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    CheckNoValidate = True
    
    'mintEditMode >= 2---报到，报到后修改,或者 mblnRegToCheck --登记直接检查
    If mintEditMode >= 2 Or mblnRegToCheck Then '判断检查号是否重复
        
        If mlngBuildType = 1 Then
            '1-按科室递增,查询同一科室中是否有相同的检查号
            gstrSQL = "Select A.姓名,A.性别,A.年龄,B.病人ID From 影像检查记录 A,病人医嘱记录 B Where A.执行科室ID=[1] AND 检查号=[2] " _
                        & " AND B.ID=A.医嘱ID AND B.相关ID IS NULL"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurDeptId, txt检查号.Text)
        Else
            '0-按影像类别递增,查询同一影像类别是否有相同的检查号
            gstrSQL = "Select A.姓名,A.性别,A.年龄,B.病人ID From 影像检查记录 A,病人医嘱记录 B Where 影像类别=[1] AND 检查号=[2] " _
                        & " AND B.ID=A.医嘱ID AND B.相关ID IS NULL"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrItemType, txt检查号.Text)
        End If
        
        If Not rsTmp.EOF Then   ' 存在重复的检查号
            'mlngUnicode = 0--检查号流水递增;   rsTmp!病人ID <> mlngPatiId --检查号跟其他病人的检查号重复了
            If mlngUnicode = 0 Or rsTmp!病人ID <> mlngPatiId Then
            
                If mblnCanOverWrite Then    '允许检查号重复
                    '84537
                    Do While Not rsTmp.EOF
                        If NVL(rsTmp!姓名) <> PatiIdentify.Text Then
                            '如果患者姓名不同，则需要自动加1生成新的检查号后重新判断
                            txt检查号.Text = Next检查号
                            
                            MsgBoxD Me, "当前检查号与下列患者重复！" & vbCrLf & "患者信息：" & NVL(rsTmp!姓名) & " " _
                                & NVL(rsTmp!性别) & " " & NVL(rsTmp!年龄) & vbCrLf _
                                & "已经重新生成检查号：" & txt检查号.Text & "，请核实后再次确定。", vbInformation, gstrSysName
                            
                            CheckNoValidate = False
                            Exit Function
                        End If
                                            
                        rsTmp.MoveNext
                    Loop
                    
                    If Not mblnUsePatientNo Then
                        rsTmp.MoveFirst
                        If MsgBoxD(Me, "当前检查号与下列患者重复！是否继续！" & vbCrLf & "患者信息：" & NVL(rsTmp!姓名) & " " _
                                    & NVL(rsTmp!性别) & " " & NVL(rsTmp!年龄), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            txt检查号.Text = Next检查号
                            MsgBoxD Me, "已经重新生成检查号：" & txt检查号.Text & "，请核实后再次确定。", vbInformation, gstrSysName
                            txt检查号.SetFocus
                            CheckNoValidate = False
                            Exit Function
                        End If
                    End If
                    
                Else        '不允许检查号重复
                    '强制将检查号替换成新生成的号码
                    On Error GoTo errH
                    'Call Next检查号
errH:
                    '重新生成检查号
                    txt检查号.Text = Next检查号
                    
                    MsgBoxD Me, "当前检查号与下列患者重复！请检查！" & vbCrLf & "患者信息：" & NVL(rsTmp!姓名) & " " & NVL(rsTmp!性别) & " " & NVL(rsTmp!年龄) _
                        & vbCrLf & "已经重新生成检查号：" & txt检查号.Text & "，请核实后再次确定。", vbExclamation, gstrSysName
                    txt检查号.SetFocus
                    
                    CheckNoValidate = False
                    Exit Function
                End If
            End If
        End If
    End If
    Exit Function

errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub sutSetTxtEnable(thisBox As TextBox, blnEnable As Boolean)
    thisBox.Enabled = blnEnable
    If blnEnable = True Then
        thisBox.BackColor = vbWhite
    Else
        thisBox.BackColor = &H8000000B
    End If
End Sub

Private Sub LoadDefaultCheckDoctor()
'加载检查技师下拉框默认值
' mblnRegToCheck     mintEditMode :  '0－登记、1－登记后修改、2－报到、3－报到后修改
'规则：登记（不直接报到）、登记后修改  默认为空
'报到、登记（登记后直接报到）：加载注册表值
'报到后修改：加载数据库值
On Error GoTo errH
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer

    If (mintEditMode = 0 And Not mblnRegToCheck) Or mintEditMode = 1 Then
    '登记（不直接报到）、登记后修改  默认为空
        cbo技师一.ListIndex = -1
        cbo技师二.ListIndex = -1
    ElseIf (mintEditMode = 0 And mblnRegToCheck) Or mintEditMode = 2 Then
    '报到、登记（登记后直接报到）：加载注册表值
        For i = 0 To cbo技师一.ListCount - 1
            If zlStr.NeedName(cbo技师一.list(i)) = mstrExamineDoctorFst Then
                cbo技师一.ListIndex = i
                Exit For
            Else
                cbo技师一.ListIndex = -1
            End If
        Next i
            
        For i = 0 To cbo技师二.ListCount - 1
            If zlStr.NeedName(cbo技师二.list(i)) = mstrExamineDoctorSed Then
                cbo技师二.ListIndex = i
                Exit For
            Else
                cbo技师二.ListIndex = -1
            End If
        Next i
        
    ElseIf mintEditMode = 3 Then
    '报到后修改：加载数据库值
        strSql = "select 检查技师,检查技师二 from 影像检查记录 where 医嘱ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获得检查技师", mlngAdviceID)
            
        If rsTmp.RecordCount > 0 Then
            For i = 0 To cbo技师一.ListCount - 1
                    If zlStr.NeedName(cbo技师一.list(i)) = NVL(rsTmp!检查技师) Then
                        cbo技师一.ListIndex = i
                        Exit For
                    Else
                        cbo技师一.ListIndex = -1
                    End If
            Next i
        End If
        
        If rsTmp.RecordCount > 0 Then
            For i = 0 To cbo技师二.ListCount - 1
                If zlStr.NeedName(cbo技师二.list(i)) = NVL(rsTmp!检查技师二) Then
                    cbo技师二.ListIndex = i
                    Exit For
                Else
                    cbo技师二.ListIndex = -1
                End If
            Next i
        End If
    End If
    Exit Sub
    
errH:
    If ErrCenter = 1 Then Resume
End Sub

