VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmPatholRIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "检查登记"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin DicomObjects.DicomViewer dcmTmpView 
      Height          =   255
      Left            =   1320
      TabIndex        =   82
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
      _Version        =   262147
      _ExtentX        =   661
      _ExtentY        =   450
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
      Height          =   405
      Left            =   6315
      TabIndex        =   34
      ToolTipText     =   "保存(F2)"
      Top             =   7920
      Width           =   1245
   End
   Begin VB.Frame framPatholInf 
      Height          =   735
      Left            =   0
      TabIndex        =   73
      Top             =   4800
      Width           =   10350
      Begin VB.ComboBox cbxStudyType 
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
         Left            =   6375
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtPatholNum 
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
         Left            =   1305
         TabIndex        =   20
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label labStudyType 
         Caption         =   "号别名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5025
         TabIndex        =   75
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label labPatholNum 
         Caption         =   "病理号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   0
      TabIndex        =   45
      Top             =   360
      Width           =   10350
      Begin VB.ComboBox cbo待处理人 
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
         Left            =   8070
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2280
         Width           =   2070
      End
      Begin VB.CommandButton cmdSelectPinyinName 
         Caption         =   "…"
         Height          =   350
         Left            =   3080
         TabIndex        =   81
         Top             =   680
         Width           =   260
      End
      Begin VB.Frame framSongJian 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   2880
         Width           =   10095
         Begin VB.TextBox txtOldBarCode 
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
            Left            =   1200
            TabIndex        =   18
            Top             =   960
            Width           =   3285
         End
         Begin VB.ComboBox cboUnitName 
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
            Left            =   1200
            TabIndex        =   14
            Text            =   "cboUnitName"
            Top             =   0
            Width           =   3285
         End
         Begin VB.TextBox txtOldStudyNo 
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
            Left            =   6720
            TabIndex        =   17
            Top             =   480
            Width           =   3285
         End
         Begin VB.TextBox txtSendTag 
            Height          =   360
            Left            =   6720
            TabIndex        =   19
            Top             =   960
            Width           =   3285
         End
         Begin VB.TextBox txtSubmitDoctor 
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
            Left            =   1200
            TabIndex        =   16
            Top             =   480
            Width           =   3285
         End
         Begin VB.TextBox txtFormDepart 
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
            Left            =   6720
            TabIndex        =   15
            Top             =   0
            Width           =   3285
         End
         Begin VB.Label labOldBarCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "原条码号"
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
            Left            =   0
            TabIndex        =   86
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label labOldStudyNo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "原检查号"
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
            Left            =   5400
            TabIndex        =   85
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label lab备注 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备    注"
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
            Left            =   5400
            TabIndex        =   84
            Top             =   960
            Width           =   1170
         End
         Begin VB.Label labSendDoctor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "送 检 人"
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
            Left            =   0
            TabIndex        =   78
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label labSendRoom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "送检科室"
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
            Left            =   5400
            TabIndex        =   77
            Top             =   0
            Width           =   1140
         End
         Begin VB.Label labSendUnit 
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
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.TextBox txt年龄 
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
         Left            =   8070
         MaxLength       =   20
         TabIndex        =   2
         Top             =   210
         Width           =   1155
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
         ItemData        =   "frmPatholRIS.frx":0000
         Left            =   9255
         List            =   "frmPatholRIS.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   885
      End
      Begin VB.TextBox txt医嘱内容 
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
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1470
         Width           =   4980
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   360
         Left            =   6315
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1455
         Width           =   300
      End
      Begin VB.TextBox Txt部位方法 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1335
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   1905
         Width           =   5295
      End
      Begin VB.ComboBox cbo医生 
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
         Left            =   8070
         TabIndex        =   9
         Text            =   "cbo医生"
         Top             =   1005
         Width           =   2070
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
         ItemData        =   "frmPatholRIS.frx":001D
         Left            =   4605
         List            =   "frmPatholRIS.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1040
         Width           =   2025
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
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   2025
      End
      Begin VB.TextBox Txt身份证号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4605
         TabIndex        =   5
         Top             =   680
         Width           =   2025
      End
      Begin VB.TextBox Txt电话 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8070
         TabIndex        =   6
         Top             =   645
         Width           =   2070
      End
      Begin VB.TextBox Txt英文名 
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
         TabIndex        =   4
         Top             =   680
         Width           =   1750
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
         ItemData        =   "frmPatholRIS.frx":0021
         Left            =   4605
         List            =   "frmPatholRIS.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   360
         Index           =   0
         Left            =   8070
         TabIndex        =   11
         Top             =   1425
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   151781379
         CurrentDate     =   38222
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   375
         Index           =   1
         Left            =   8070
         TabIndex        =   12
         Top             =   1830
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   151781379
         CurrentDate     =   38222
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   360
         Left            =   720
         TabIndex        =   0
         ToolTipText     =   """数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"""
         Top             =   240
         Width           =   2625
         _ExtentX        =   4630
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
         IDKindStr       =   $"frmPatholRIS.frx":0037
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
      Begin VB.Label lab待处理人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "待处理人"
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
         Left            =   6840
         TabIndex        =   83
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查时间"
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
         Left            =   6855
         TabIndex        =   72
         Top             =   1860
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
         Left            =   6855
         TabIndex        =   58
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请时间"
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
         Left            =   6855
         TabIndex        =   57
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Lbl部位方法 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查标本"
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
         TabIndex        =   56
         Top             =   1905
         Width           =   1245
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
         Left            =   30
         TabIndex        =   55
         Top             =   1485
         Width           =   1245
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请医生"
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
         Left            =   6855
         TabIndex        =   53
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请科室"
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
         Left            =   3390
         TabIndex        =   52
         Top             =   1080
         Width           =   1140
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
         Left            =   45
         TabIndex        =   51
         Top             =   1095
         Width           =   1245
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电    话"
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
         Left            =   6855
         TabIndex        =   50
         Top             =   660
         Width           =   1140
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
         Left            =   3390
         TabIndex        =   49
         Top             =   720
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
         Left            =   45
         TabIndex        =   48
         Top             =   705
         Width           =   1245
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
         Left            =   3390
         TabIndex        =   47
         Top             =   315
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
         Left            =   45
         TabIndex        =   46
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   10350
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
         Height          =   255
         Left            =   8655
         TabIndex        =   44
         Top             =   60
         Width           =   1620
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
         Height          =   300
         Left            =   7125
         TabIndex        =   41
         Top             =   75
         Width           =   1290
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
         Height          =   300
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   75
         Width           =   1725
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
         Height          =   300
         Left            =   1365
         TabIndex        =   38
         Top             =   75
         Width           =   1590
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
         Left            =   3225
         TabIndex        =   43
         Top             =   60
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
         Left            =   6435
         TabIndex        =   42
         Top             =   60
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   39
         Top             =   60
         Width           =   1140
      End
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
      Height          =   400
      Left            =   7710
      TabIndex        =   33
      ToolTipText     =   "保存(F2)"
      Top             =   7920
      Width           =   1245
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
      Height          =   400
      Left            =   9075
      TabIndex        =   35
      Top             =   7920
      Width           =   1245
   End
   Begin VB.Frame frm其他信息 
      Height          =   2175
      Left            =   0
      TabIndex        =   59
      Top             =   5520
      Width           =   10350
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
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1770
         Width           =   1800
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
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1755
         Width           =   1800
      End
      Begin VB.TextBox txt附加主述 
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
         Left            =   1290
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1365
         Width           =   8715
      End
      Begin VB.TextBox Txt联系地址 
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
         Left            =   1290
         TabIndex        =   28
         Top             =   990
         Width           =   8715
      End
      Begin VB.TextBox Txt邮编 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8220
         TabIndex        =   27
         Top             =   630
         Width           =   1800
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
         Left            =   4900
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   615
         Width           =   1800
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
         Height          =   360
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2124
      End
      Begin VB.TextBox Txt体重 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8220
         TabIndex        =   24
         Top             =   255
         Width           =   1800
      End
      Begin VB.TextBox Txt身高 
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
         Height          =   360
         Left            =   4900
         TabIndex        =   23
         Top             =   240
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtp出生日期 
         Height          =   336
         Left            =   1288
         TabIndex        =   22
         Top             =   238
         Width           =   2184
         _ExtentX        =   3863
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   151781379
         CurrentDate     =   38222
      End
      Begin VB.Label Label27 
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
         Height          =   210
         Left            =   10050
         TabIndex        =   80
         Top             =   315
         Width           =   225
      End
      Begin VB.Label Label26 
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
         Height          =   210
         Left            =   6720
         TabIndex        =   79
         Top             =   294
         Width           =   224
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
         Height          =   255
         Left            =   8205
         TabIndex        =   71
         Top             =   1785
         Width           =   1800
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
         Left            =   7005
         TabIndex        =   70
         Top             =   1785
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
         Left            =   3465
         TabIndex        =   69
         Top             =   1800
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
         Left            =   60
         TabIndex        =   68
         Top             =   1800
         Width           =   1170
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
         Left            =   90
         TabIndex        =   67
         Top             =   1395
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
         Left            =   90
         TabIndex        =   66
         Top             =   1005
         Width           =   1140
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮   编"
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
         Left            =   7095
         TabIndex        =   65
         Top             =   645
         Width           =   1020
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职   业"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3766
         TabIndex        =   64
         Top             =   644
         Width           =   1022
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
         Left            =   60
         TabIndex        =   63
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体   重"
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
         Left            =   7095
         TabIndex        =   62
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身   高"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3752
         TabIndex        =   61
         Top             =   252
         Width           =   1022
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
         Left            =   90
         TabIndex        =   60
         Top             =   285
         Width           =   1140
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   3480
      Top             =   7920
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholRIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'模块变量----注册表参数
Private mstrUnitName As String '送检单位
Private mstrExamineDoctor As String '待处理人
'模块变量----以从值从外部传入
Public mstrPrivs As String          '调用者的权限
Public mlngModul As Long            '由谁调用
Public mlngAdviceID As Long         '医嘱ID
Public mlngSendNo As Long           '发送号
Public mintEditMode As Integer      '0－登记、1－登记后修改、2－报到、3－报到后修改
Public mlngCurDeptId As Long        '当前科室ID
Public mstrTechnicRoom As String    '报到执行间
Public mblnOK As Boolean            '保存或取消

Public mintImgCount As Integer      '已扫描图像数量

'扫描窗体对象
Private frmPetitionCap As frmPetitionCapture

'公共模块变量------以下值从参数表中取得
Private mblnChangeNo As Boolean     '手工调整检查号
Private mblnCanOverWrite            '允许检查号重复
Private mblnLike As Boolean, mlngLike As Long    '姓名模糊查找,查找天数
Private mBeforeDays As Integer      '过滤天数
Private mlngGoOnReg As Long         '连续登记 0-非连续,1-连续
Private mblnAutoPrint As Boolean    '报到后自动打印申请单
Private mlngUnicode As Long         '患者检查号保持不变,1-保持检查号不变；0-检查号流水递增
Private mlngUnicodeType As Long     '检查号保持不变类别,不变类别 0-按类别不变 1-按科室不变;
Private mlngBuildType As Long       '检查号生成方式,0-按类别递增 1-按科室递增
Private mblnRegToCheck As Boolean   '登记直接检查
Private mblnNoshowReagent As Boolean '不显示造影剂
Private mblnNoshowAddons As Boolean '不显示附加主述
Private mintCheckInMode As Integer  '登记模式 1--精简模式，2--正常模式
Private mblnUseReferencePatient     '使用关联病人模式
Private mintCapital As Integer      '拼音名大小写
Private mblnUseSplitter As Boolean  '拼音名分隔符
Private mblnAllPatientIsOutside As Boolean '所有登记病人标记为外来
Private mblnNameColColorCfg As Boolean  '是否根据病人类型设置姓名颜色
Private mblnOrdinaryNameColColorCfg As Boolean '缺省的病人是否根据病人类型设置姓名颜色
Private mstrDefaultPatientType As String '缺省的病人类型

'公共模块变量------以下运行中赋值
Private mintSourceType As Integer   '病人来源 1-门诊 2-住院 3-外来 4-体检
Private mlngPatiId As Long, mlngPageID As Long  '病人ID,主页ID
'Private mstrItemType As String      '影像类别
Private mlngClinicID As Long        '诊疗项目ID
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
Private mlngNextCheckNo As Long     '记录本次获取到的下一个检查号
Private mlngPatholStationMoneyExeModle As Long   '病理采集的费用执行模式 0-报到时执行，1-检查时执行，2-报告时执行

Private mobjSquareCard As Object    '一卡通，卡结算部件
Private oneSquardCard As TSquardCard

Public mlngPatholSerialNum As Long
Public mstrPatholInitNum As String


Private mblnShowSentInfo As Boolean    '是否启用显示送检信息
Private mStrOutSideCfg As String     '外院配置

Private mblnIsOutSideHosp As Boolean     '是否是外院科室
Private mblnIsPetitionScan As Boolean    '是否启用申请单扫描
Private mblnIsSamePatient As Boolean     '是否存在相同病人

Private mlngBaby As Long            '是否婴儿，0--不是婴儿，1-9表示婴儿序号

Private mlngInsureCheckType As Long         '医保对码检查类型 0-不检查， 1-仅提示，2-禁止
Private mobjInsure As Object

Private mfrmParent As Form          '父窗体
Private mobjPublicPatient As Object

Private Sub SaveAdviceData()
'------------------------------------------------
'功能：保存医嘱
'参数： 无
'返回：无
'------------------------------------------------
    Dim str检查时间 As String
    Dim str申请时间 As String, curDate As String
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
    strDoctor = zlStr.NeedName(Me.cbo医生.Text)
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
            "','" & zlCommFun.ToVarchar(Txt电话, 20) & "','" & zlCommFun.ToVarchar(Txt邮编, 6) & "'," & curDate & ",'','" & mstrRegNo & "'," & zlStr.To_Date(dtp出生日期.value) & ",NULL)"
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
            lngMasSeq = Nvl(rsTemp!序号, 0) + 1
            lngSonSeq = lngMasSeq
        End If
    End If
    
    lngAdviceID = zlDatabase.GetNextId("病人医嘱记录")
    lngSendNO = zlDatabase.GetNextNo(10) '医嘱发送号
    
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
                i + 1 & "," & ZVal(Split(arrAppend(i), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(i), "<Split2>")(3), "'", "''") & "'" & _
                            IIf(i = 0, ",1", "") & ")"
        Next
    End If

    '有收费单据号的，设置费用记录和医嘱的关联关系
    If mstrChargeNo <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人费用记录_医嘱('" & strNO & "',1," & lngAdviceID & ")"
    End If
    
    
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cboAge_LostFocus()
    If Not CheckOldData(txt年龄, cboAge) Then Exit Sub
    If IsNumeric(txt年龄.Text) Then Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
End Sub


Private Function GetPatholNum(ByVal lngID As Long) As String
'根据检查类型获取病理号
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetPatholNum = ""
    
    strSql = "select Zl_病理号码_序号获取([1]) as 病理序号 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    mlngPatholSerialNum = Val(Nvl(rsData!病理序号))
    
    strSql = "select Zl_病理号码_生成([1],[2]) as 病理号 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID, mlngPatholSerialNum)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    mstrPatholInitNum = Nvl(rsData!病理号)
    
    GetPatholNum = mstrPatholInitNum
End Function





Private Sub cboUnitName_Click()
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "送检单位", zlStr.NeedName(cboUnitName.Text)
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Txt英文名.Text = control.Caption
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxStudyType_Click()
On Error GoTo errHandle
    txtPatholNum.Text = GetPatholNum(Val(cbxStudyType.Text))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
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
    Dim l As Long
    Dim blnTran As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rsMother As New ADODB.Recordset
    Dim rsPatiInfo As New ADODB.Recordset
    Dim int记录性质 As Integer     '病人医嘱发送.记录性质，本次医嘱的记录性质，1-收费记录；2-记帐记录
    Dim int门诊记帐 As Integer     '病人医嘱发送.门诊记帐，门诊和住院医生站发送为门诊记帐时填为1,用于区分门诊记帐和住院记帐，其他的都填为空
    Dim str诊疗类别 As String
    Dim lng发送号 As Long
    Dim str单据号 As String
    Dim str医嘱IDs As String
    Dim lngCurFromType As Long
    Dim strMsg As String
    Dim lngMsgResult As Long
    
    On Error GoTo errHandle
    
    arrSQL = Array()
    
    lngCurFromType = mintSourceType
    If mblnAllPatientIsOutside Then mintSourceType = 3
    
    '如果没有检查登记权限，则只能修改病理号和检查类型(该信息为科室内部信息)
    If Not Frame2.Visible Then
        If Trim(txtPatholNum.Text) = "" Then
            If MsgBoxD(Me, "病理号不能为空，是否自动生成病理号？", vbYesNo, Me.Caption) = vbYes Then
                txtPatholNum.Text = GetPatholNum(Val(cbxStudyType.Text))
            End If
            txtPatholNum.SetFocus

            Exit Sub
        End If

        '如果有病理号，才对此检查信息进行更新
        If Not txtPatholNum.Enabled Then
            Call MsgBoxD(Me, "病理信息不允许编辑。", vbInformation, Me.Caption)

            Exit Sub
        End If

        ReDim Preserve arrSQL(UBound(arrSQL) + 1)

        arrSQL(UBound(arrSQL)) = "Zl_病理报到_号码更新(" & mlngAdviceID & ",'" & UCase(txtPatholNum.Text) & "'," & Val(cbxStudyType.Text) & ")"
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(UBound(arrSQL))), "更新病理数据")

        mblnOK = True

        If mstrPatholInitNum = UCase(Trim(txtPatholNum.Text)) Then
            '更新病理序号
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)

            arrSQL(UBound(arrSQL)) = "ZL_病理号码_序号更新(" & Val(cbxStudyType.Text) & "," & mlngPatholSerialNum & ")"
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(UBound(arrSQL))), Me.Caption)
        End If

        Unload Me
        Exit Sub
    End If
    
    '检查数据输入是否合法，不合法则退出
    If ValidData = False Then Exit Sub
    
     If framPatholInf.Visible Then
        If Trim(txtPatholNum.Text) = "" Then
            If MsgBoxD(Me, "病理号不能为空，是否自动生成病理号？", vbYesNo, Me.Caption) = vbYes Then
                txtPatholNum.Text = GetPatholNum(Val(cbxStudyType.Text))
            End If
            txtPatholNum.SetFocus

            Exit Sub
        End If
    End If
    
    '登记 ， 登记作为一个单独的数据库事务来处理
    ' 如果是登记，则保存医嘱
    If mintEditMode = 0 Then
    '        1)医保对码检查
    '        2)保存医嘱（新建病人，新建医嘱，发送医嘱）
        If (lngCurFromType = 1 Or lngCurFromType = 2) And mlngInsureCheckType <> 0 Then
            '只有从门诊或住院开过来的医保病人才进行医保对码检查
            gstrSQL = "select 险类 from 病人信息 Where 病人ID = [1]"
            Set rsPatiInfo = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人险类信息", mlngPatiId)
            
            '医保对码检查
            strMsg = CheckAdviceInsure(Val(Nvl(rsPatiInfo!险类)), True, mlngPatiId, mintSourceType, _
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
        
        '病理送检信息 登记
        If framSongJian.Visible Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病理报到_送检更新(" & mlngAdviceID & ",'" & cboUnitName.Text & "','" & _
                                                     txtFormDepart.Text & "','" & txtSubmitDoctor.Text & "','" & txtOldStudyNo & "|" & txtOldBarCode & "|" & txtSendTag.Text & "')"
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
    
    
    '修改影像表的病人信息条件如下
    '1、不是登记，需要修改病人的信息，外诊病人的信息比较多
    '实际条件是：登记后修改；报到；报到后修改
    If mintEditMode <> 0 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_影像病人信息_修改(" & mintSourceType & "," & mlngAdviceID & "," & mlngPatiId & "," & _
            "'" & Trim(PatiIdentify.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & txt年龄.Text & cboAge.Text & "'," & _
            "'" & zlStr.NeedName(cbo费别.Text) & "','" & zlStr.NeedName(cbo付款方式.Text) & "','" & zlStr.NeedName(cbo民族.Text) & "'," & _
            "'" & zlStr.NeedName(cbo婚姻.Text) & "','" & zlStr.NeedName(cbo职业.Text) & "','" & zlCommFun.ToVarchar(Txt身份证号, 18) & "'," & _
            "'" & zlCommFun.ToVarchar(Txt联系地址.Text, 50) & "','" & zlCommFun.ToVarchar(Txt电话, 20) & "','" & zlCommFun.ToVarchar(Txt邮编, 6) & _
            "'," & zlStr.To_Date(CDate(dtp出生日期.value)) & "," & mlngPageID & "," & mlngBaby & ")"
    End If
    
    If mintEditMode = 1 Then
        '病理送检信息  登记后修改
        If framSongJian.Visible Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病理报到_送检更新(" & mlngAdviceID & ",'" & cboUnitName.Text & "','" & _
                                                     txtFormDepart.Text & "','" & txtSubmitDoctor.Text & "','" & txtOldStudyNo & "|" & txtOldBarCode & "|" & txtSendTag.Text & "')"
        End If
    End If
    '1、报到
    '2、登记，且登记后直接检查。其实也就是报到
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        
        '检查费用以及一卡通的处理
        '业务逻辑是：
        '1、总体逻辑没有收费的不能报到，但是如果有“未缴费报到”权限的，可以在没有收费的情况下报到。
        '   在刷新信息的时候已经控制报到的确定按钮。
        '2、对公共基础参数的支持：
        '       参数号28--门诊一卡通，消费减少剩余款额时是否需要验证
        '       参数号81--执行后自动审核
        '       参数号163--门诊一卡通，项目执行前必须先收费或先记帐审核
        '3、先处理需要一卡通消费确认的，条件是以下之一
        '       （1）记录性质=1
        '       （2）执行后自动审核=False，记录性质=2，且 “来源<>住院”  或者 “来源=住院，门诊记帐”。
        '   如果一卡通消费确认成功，则可以报到。如果一卡通消费确认不成功，则根据权限“未缴费报到”提示是否继续报到。
        '4、再处理一卡通费用减少验证的，只处理记账的，条件是：
        '       （1）记录性质=2，执行后自动审核=True
        '       （2）有未审核费用
        '
        gstrSQL = "Select A.记录性质,A.门诊记帐,A.发送号,A.NO,B.诊疗类别 from 病人医嘱发送 A,病人医嘱记录 B  where A.医嘱ID=B.ID and  B.ID =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS报到查找记录性质", mlngAdviceID)
        If rsTmp.EOF = False Then
            int记录性质 = Nvl(rsTmp!记录性质, 0)
            int门诊记帐 = Nvl(rsTmp!门诊记帐, 0)
            str诊疗类别 = Nvl(rsTmp!诊疗类别)
            lng发送号 = rsTmp!发送号
            str单据号 = Nvl(rsTmp!NO)
        End If
        
        If int记录性质 = 1 Or _
            (gbln执行后审核 = False And int记录性质 = 2 And (mintSourceType <> 2 Or (mintSourceType = 2 And int门诊记帐 = 1))) Then
            
            If Not ItemHaveCash(mintSourceType, False, mlngAdviceID, 0, lng发送号, str诊疗类别, str单据号, int记录性质, _
                int门诊记帐, 0) Then
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
                        '如果有“未缴费报到”权限，则提示是否确认未收费可以报到？
                        If CheckPopedom(mstrPrivs, "未缴费报到") Then
                            If MsgBoxD(Me, "缴费不成功，该病人还存在未收费的费用，是否继续报到？", vbYesNo, "缴费失败") = vbNo Then
                                Exit Sub
                            End If
                        Else
                            MsgBoxD Me, "缴费不成功，该病人还存在未收费的费用，无法报到，请检查。", vbOKOnly, "缴费失败"
                            Exit Sub
                        End If
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
        
        If gbln执行后审核 And int记录性质 = 2 Then
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
                If glng消费验证 <> 0 And gbln执行后审核 _
                    And curMoney > 0 And mintSourceType = 1 Then
                    If Not zlDatabase.PatiIdentify(Me, glngSys, mlngPatiId, curMoney, , , , , , , IIf(glng消费验证 = 2, True, False)) Then Exit Sub
                End If
            End If
        End If
        
        '开始检查
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        
        '影像类别"DG"表示病理
        arrSQL(UBound(arrSQL)) = "ZL_影像检查_BEGIN(Null,Null," & mlngAdviceID & "," & mlngSendNo & ",'DG','" & _
            Trim(Me.PatiIdentify.Text) & "','" & Trim(Txt英文名.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & _
            txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, "") & "'," & zlStr.To_Date(dtp出生日期.value) & ",'" & zlCommFun.ToVarchar(Txt身高, 16) & "','" & _
            zlCommFun.ToVarchar(Txt体重, 16) & "',Null,Null,Null,Null,Null,'" & txt附加主述.Text & "',Null," & mlngCurDeptId & ",'" & zlStr.NeedName(cbo待处理人.Text) & "')"
        
        '设置影像检查记录--执行过程为-已报到，报到时处理记账的费用
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_影像检查_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
        
        
        '病理在报到时，需要执行费用
        If mlngPatholStationMoneyExeModle = 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_影像费用执行(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDeptId & ")"
        End If
        
        '病理检查直接报道
        If framPatholInf.Visible Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病理报到_号码更新(" & mlngAdviceID & ",'" & UCase(txtPatholNum.Text) & "'," & Val(cbxStudyType.Text) & ")"
            
            If mstrPatholInitNum = UCase(Trim(txtPatholNum.Text)) Then
                '更新病理序号
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病理号码_序号更新(" & Val(cbxStudyType.Text) & "," & mlngPatholSerialNum & ")"
            End If
        End If
        
        
        '病理送检信息  报到
        If framSongJian.Visible Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病理报到_送检更新(" & mlngAdviceID & ",'" & cboUnitName.Text & "','" & _
                                                     txtFormDepart.Text & "','" & txtSubmitDoctor.Text & "','" & txtOldStudyNo & "|" & txtOldBarCode & "|" & txtSendTag.Text & "')"
        End If
        
    End If   '报到的if
    
    
    
    '报到后修改
    If mintEditMode = 3 Then
        
        '修改病人信息
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_影像检查记录_UPDATE(" & mlngAdviceID & ", " & mlngSendNo & ",Null,'" & _
            Trim(Me.PatiIdentify.Text) & "','" & Trim(Txt英文名.Text) & "','" & zlStr.NeedName(cbo性别.Text) & "','" & _
            txt年龄.Text & IIf(cboAge.Visible, cboAge.Text, "") & "'," & zlStr.To_Date(dtp出生日期.value) & ",'" & zlCommFun.ToVarchar(Txt身高, 16) & "','" & _
            zlCommFun.ToVarchar(Txt体重, 16) & "',Null,Null,Null,'" & txt附加主述.Text & "',Null," & zlStr.To_Date(dtp(1).value) & ",'" & zlStr.NeedName(cbo待处理人.Text) & "')"
            
          '病理检查直接报道
        If framPatholInf.Visible Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病理报到_号码更新(" & mlngAdviceID & ",'" & UCase(txtPatholNum.Text) & "'," & Val(cbxStudyType.Text) & ")"

            If mstrPatholInitNum = UCase(Trim(txtPatholNum.Text)) Then
                '更新病理序号
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病理号码_序号更新(" & Val(cbxStudyType.Text) & "," & mlngPatholSerialNum & ")"
            End If
        End If

        '病理送检信息 报到后修改
        If framSongJian.Visible Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病理报到_送检更新(" & mlngAdviceID & ",'" & cboUnitName.Text & "','" & _
                                                     txtFormDepart.Text & "','" & txtSubmitDoctor.Text & "','" & txtOldStudyNo & "|" & txtOldBarCode & "|" & txtSendTag.Text & "')"
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
        
    '报到,或登记后直接检查， 的后续处理
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        '打印申请单
        AutoPrintApplication
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

    mblnOK = True
    '如果是连续登记，而且处于登记状态，则不关闭窗口。
    If mlngGoOnReg = 1 And mintEditMode = 0 Then
        InitMvar '初始化模块变量
        InitEdit '初始化界面
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
Private Sub AutoPrintApplication()
'功能:根据能数据自动打印申请单
Dim rsTemp As ADODB.Recordset, strBillNo As String, strExseNo As String, intExseKind As Integer

On Error GoTo errHand

    If Not mblnAutoPrint Then Exit Sub
    gstrSQL = "select NO,记录性质 from 病人医嘱发送 where 医嘱ID=[1] and 发送号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取NO", mlngAdviceID, mlngSendNo)
    If rsTemp.EOF Then Exit Sub
    strExseNo = rsTemp!NO: intExseKind = rsTemp!记录性质
    
    gstrSQL = "Select B.ID, B.编号" & vbNewLine & _
                "From 病历单据应用 A, 病历文件列表 B" & vbNewLine & _
                "Where A.诊疗项目id =[1] And A.应用场合 =[2] And A.病历文件id = B.ID And B.种类 = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取单据编号", mlngClinicID, CLng(Decode(mintSourceType, 1, 1, 2, 2, 1)))
    If rsTemp.EOF Then Exit Sub
    strBillNo = "ZLCISBILL" & Format(rsTemp!编号, "00000") & "-1"
    ReportOpen gcnOracle, glngSys, strBillNo, Me, "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & mlngAdviceID, 2
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdPetitionCapture_Click()
On Error GoTo errHand
    
    If frmPetitionCap Is Nothing Then
        Set frmPetitionCap = New frmPetitionCapture
    End If

     '打开扫描申请单窗口
    Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                            mlngCurDeptId, _
                                            Nvl(Mid(cbo开单科室.Text, InStr(cbo开单科室.Text, "-") + 1, Len(cbo开单科室.Text))), _
                                            Nvl(Trim(PatiIdentify.Text)), _
                                            Nvl(txt年龄.Text), _
                                            Nvl(Mid(cbo性别.Text, InStr(cbo性别.Text, "-") + 1, Len(cbo性别.Text))), _
                                            Nvl(txt医嘱内容.Text), _
                                            Nvl(Txt部位方法.Text), _
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
                
                Call LoadStudyType
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
    gstrSQL = "Select Distinct A.ID,A.编码,A.名称,Nvl(A.计算单位,'次') as 计算单位,Nvl(A.标本部位,' ') as 标本部位," & _
                "A.操作类型 As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,Nvl(执行频率,0) As 执行频率ID," & _
                "Nvl(计算方式,0) As 计算方式ID,Nvl(执行安排,0) As 执行安排ID,Nvl(计价性质,0) As 计价性质ID," & _
                "Nvl(执行科室,0) As 执行科室ID,B.影像类别" & _
              " From 诊疗项目目录 A,影像检查项目 B,诊疗项目别名 C,诊疗执行科室 D" & _
              " Where A.ID=B.诊疗项目ID AND A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID" & _
                    " And D.执行科室ID=" & mlngCurDeptId & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " & _
                    " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) " & _
                    " And A.服务对象 IN(" & IIf(mintSourceType = 3, "1,2,4", mintSourceType) & ",3) And Nvl(A.单独应用,0)=1 " & _
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
    Dim t_Pati As TYPE_PatiInfoEx
    Dim blnOk As Boolean
    Dim strExtData As String, strAppend As String
    Dim lngHwnd As Long, int服务对象 As Integer
    
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
  
    lngHwnd = IIf(mintCheckInMode = 1, Me.Txt部位方法.hWnd, Me.Txt联系地址.hWnd)
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
    
    Dim blnShowPatholNum As Boolean
    Dim blnShowOtherInf As Boolean
    Dim blnShowStandard As Boolean
    
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
    Call sutSetTxtEnable(txt年龄, mintSourceType = 3)
    
    cbo性别.Enabled = mintSourceType = 3: cboAge.Enabled = mintSourceType = 3
    dtp出生日期.Enabled = mintSourceType = 3:
    Call sutSetTxtEnable(Txt身份证号, mintSourceType = 3)
    
    cbo费别.Enabled = (mintSourceType = 3)
    cbo付款方式.Enabled = (mintSourceType = 3): cbo民族.Enabled = blnEditableState
    cbo职业.Enabled = blnEditableState: cbo婚姻.Enabled = blnEditableState
    
    '这三个信息一直都可以修改
    Call sutSetTxtEnable(Txt电话, True)
    Call sutSetTxtEnable(Txt邮编, True)
    Call sutSetTxtEnable(Txt联系地址, True)
    
    blnShowPatholNum = False
    blnShowStandard = True 'CheckPopedom(mstrPrivs, "检查登记")
    blnShowOtherInf = blnShowStandard And (mintCheckInMode <> 1)
    
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
            
            '登记后直接报道 则显示病理号
            blnShowPatholNum = mblnRegToCheck
            
            '登记的时候，姓名允许修改
            PatiIdentify.objTxtInput.Locked = False
            
            cmdSelectPinyinName.Enabled = False
            
            
            Call sutSetTxtEnable(Txt英文名, True)
            Call sutSetTxtEnable(Txt身高, mblnRegToCheck)
            Call sutSetTxtEnable(Txt体重, mblnRegToCheck)
            Call sutSetTxtEnable(txt附加主述, mblnRegToCheck)
        Case 1          '1－登记后修改
            Me.Caption = "修改信息"
            
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            cmdSel.Enabled = False
            chk紧急.Enabled = False: cbo开单科室.Enabled = False
            cbo医生.Enabled = False
            
            PatiIdentify.Enabled = (mintSourceType = 3)
            
            cmdSelectPinyinName.Enabled = False
            Call sutSetTxtEnable(txt医嘱内容, False)
            Call sutSetTxtEnable(Txt英文名, False)
            
            Call sutSetTxtEnable(Txt身高, False)
            Call sutSetTxtEnable(Txt体重, False)
            Call sutSetTxtEnable(txt附加主述, False)
        Case 2          '2－报到
            Me.Caption = "检查报到"
            
            blnShowPatholNum = True
            
            cmdSelectPinyinName.Enabled = True
            cbo开单科室.Enabled = False: cbo医生.Enabled = False
            chk紧急.Enabled = False
            dtp(0).Enabled = False
            dtp(1).Enabled = True
            cmdSel.Enabled = False
            
            Call sutSetTxtEnable(txt医嘱内容, False)
            
            Call sutSetTxtEnable(Txt英文名, False)
            Call sutSetTxtEnable(txt附加主述, True)
        Case 3          '3－报到后修改
            Me.Caption = "修改信息"
            
            blnShowPatholNum = True
            
            cmdSelectPinyinName.Enabled = True
            dtp(0).Enabled = False
            dtp(1).Enabled = True
            cmdSel.Enabled = False
            chk紧急.Enabled = False
            cbo开单科室.Enabled = False
            cbo医生.Enabled = False
            
            PatiIdentify.Enabled = (mintSourceType = 3)
            
            Call sutSetTxtEnable(txt医嘱内容, False)
            
            Call sutSetTxtEnable(Txt英文名, False)
            Call sutSetTxtEnable(Txt身高, True)
            Call sutSetTxtEnable(Txt体重, True)
            Call sutSetTxtEnable(txt附加主述, True)
    End Select
    
    framSongJian.Visible = mblnShowSentInfo
    Frame2.Height = IIf(mblnShowSentInfo, 4455, 2765)

    
    '显示病理号的三种情况
    '1.报到的时候且为使用标本核收的功能，需要在该窗口中显示病理号
    '2.修改病理检查信息的时候，需要在该窗口中显示病理号
    '3.登记后直接报到
    framPatholInf.Visible = blnShowPatholNum
    
    If blnShowPatholNum Then
        framPatholInf.Top = Frame2.Top + Frame2.Height
        
        frm其他信息.Top = framPatholInf.Top + framPatholInf.Height
    Else
        frm其他信息.Top = Frame2.Top + Frame2.Height
    End If
    
    '调整窗口高度
    Me.Height = IIf(blnShowStandard, Frame2.Top + 240, 0) + _
                IIf(blnShowStandard, Frame2.Height + 120, 0) + _
                IIf(blnShowPatholNum, framPatholInf.Height + 120, 120) + _
                IIf(blnShowOtherInf, frm其他信息.Height, 0) + 120 + cmdOK.Height
                
                
    '调整按钮位置
    cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 120
    CmdCancle.Top = cmdOK.Top
    cmdPetitionCapture.Top = cmdOK.Top
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
End Sub

Private Sub LoadStudyType()
    '载入检查类型
    On Error GoTo errH
    Dim i As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset

'    strSQL = "select ID,名称 from 病理号码规则"
    strSql = "select ID,名称 from 病理号码规则 order by ID"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "获得病理号码分类")

    If rsData.RecordCount > 0 Then
        With cbxStudyType
        .Clear
            rsData.MoveFirst
            Do While Not rsData.EOF
                If Nvl(rsData!名称, "  ") <> "  " Then
                    .AddItem Nvl(rsData!ID, 0) & "-" & rsData!名称
                End If
                rsData.MoveNext
            Loop

        End With
    End If

    '检查类型 分为 登记后直接报到 和 单独登记 情况加载。
    If mblnRegToCheck And mintEditMode = 0 Then
        strSql = "select 执行分类 from 诊疗项目目录 where 操作类型='病理' and 名称 =[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "获得检查项目对应的执行分类", txt医嘱内容.Text)
    Else
        strSql = "select 执行分类 from 诊疗项目目录 where ID= (select 诊疗项目ID from 病人医嘱记录 where id=[1])"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "提取医嘱中的执行分类", mlngAdviceID)
    End If

    If rsData.RecordCount > 0 Then
        For i = 0 To cbxStudyType.ListCount - 1
            If Val(Mid(cbxStudyType.list(i), 1, InStr(cbxStudyType.list(i), "-") - 1)) = Val(Nvl(rsData!执行分类)) Then
                cbxStudyType.ListIndex = i
                Exit Sub
            End If
            cbxStudyType.ListIndex = 0
        Next
    Else
        cbxStudyType.ListIndex = 0
    End If
    Exit Sub
errH:
    MsgBoxD Me, err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    mblnShowSentInfo = Val(zlDatabase.GetPara("录入外院信息", glngSys, mlngModul, 0)) '是否显示送检信息
    mStrOutSideCfg = zlDatabase.GetPara("外院单位结构分类", glngSys, mlngModul, "0") '是否显示送检信息
    
    mlngGoOnReg = Val(zlDatabase.GetPara("连续登记申请", glngSys, mlngModul, 0)) '连续登记
    mblnRegToCheck = (Val(GetDeptPara(mlngCurDeptId, "登记后直接检查", 0)) = 1) '登记后直接检查
    mblnAutoPrint = Val(zlDatabase.GetPara("报到后自动打印申请单", glngSys, mlngModul, 0)) '报到后自动打印申请单
    mblnAllPatientIsOutside = IIf(Val(GetDeptPara(mlngCurDeptId, "所有登记病人标记为外来", 0)) = 0, False, True)
    mlngPatholStationMoneyExeModle = Val(zlDatabase.GetPara("病理费用执行模式", glngSys, mlngModul, 0))
    
    mlngInsureCheckType = Val(zlDatabase.GetPara(59, glngSys))  '获取医保对码检查类型
    If mlngInsureCheckType <> 0 Then
        Set mobjInsure = CreateObject("zl9Insure.clsInsure")
    End If
    
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
'    mlngTypeSuit = 0
    mblnLike = False
    mlngLike = 0
    mblnChangeNo = False
    mBeforeDays = 2
'    If mintEditMode = 0 Then mlngBaby = 0        '设置默认值，不是婴儿,只有登记模式才设置
    
    strSql = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "患者检查号保持不变"
                mlngUnicode = Nvl(rsTemp!参数值, 0)
            Case "检查号保持不变类别"
                mlngUnicodeType = Nvl(rsTemp!参数值, 0)
            Case "检查号生成方式"
                mlngBuildType = Nvl(rsTemp!参数值, 0)
'            Case "匹配数据库项目"
'                mlngTypeSuit = Nvl(rsTemp!参数值, 0)
            Case "登记时姓名模糊查找天数"
                mblnLike = IIf(Nvl(rsTemp!参数值, 0) <> 0, True, False)
                mlngLike = Abs(Nvl(rsTemp!参数值, 0))
            Case "手工调整检查号"
                mblnChangeNo = Nvl(rsTemp!参数值, 0) = 1
            Case "默认过滤天数"
                mBeforeDays = Abs(Nvl(rsTemp!参数值, 2))
            Case "允许检查号重复"
                mblnCanOverWrite = Nvl(rsTemp!参数值, 0) = 1
            Case "启动关联病人"
                mblnUseReferencePatient = Nvl(rsTemp!参数值, 0) = 1
            Case "拼音名大小写"
                mintCapital = Nvl(rsTemp!参数值, 0)
            Case "拼音名分隔符"
                mblnUseSplitter = Nvl(rsTemp!参数值, 0) = 0
        End Select
        rsTemp.MoveNext
    Wend
    
    '载入病理检查类型
    Call LoadStudyType
    
    InitFaceScheme
    InitEdit  '初始化界面数据
End Sub
Public Sub InitMvar()
    mintSourceType = 3
    mlngPatiId = 0
    mlngPageID = 0
    mlngBaby = 0
'    mstrItemType = ""
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
    mblnNameColColorCfg = GetDeptPara(mlngCurDeptId, "姓名颜色区分", 0) = "1"     '姓名颜色区分
    mblnOrdinaryNameColColorCfg = GetDeptPara(mlngCurDeptId, "缺省类型病人姓名颜色区分", 0) = "1"   '缺省类型病人姓名颜色区分
End Sub

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

'mInputType  0-就诊卡 1-病人ID 2-住院号 3-门诊号 4-挂号单 5-收费单据号 6-姓名 7-医保号 8-身份证号 9-IC卡号
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
        Case Else
             '使用姓名的时候，经常直接刷卡，所以姓名和刷卡的放在一起处理
             
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
    If mInputType = 0 Then '刷卡
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,B.执行人 As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                        " From 病人信息 A,病人挂号记录 B Where A.就诊卡号=[1] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and B.记录性质=1 and B.记录状态=1 and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制

    ElseIf mInputType = 1 Then '病人ID
         gstrSQL = "select 病人id,姓名,性别,年龄,出生日期,来源ID,主页ID,病人科室ID,当前科室ID,医生,门诊号,住院号,就诊卡号,就诊状态,卡验证码,当前床号,费别" & _
                        ",医疗付款方式,身份证号,民族,职业,婚姻状况,电话,邮编,地址,合同单位ID, 新病人" & _
                    " From(Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,Nvl(B.执行人,'') As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人,B.登记时间" & _
                  " From 病人信息 A,病人挂号记录 B Where A.病人ID=[2] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and '%'='%' " & _
                  " order by B.登记时间 desc) where rownum=1" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 2 Then '住院号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.入院科室ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,B.住院医师 As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A,病案主页 B " & _
                  " Where A.住院号=[1] And A.病人ID=B.病人ID and A.出院时间 Is Null and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 3 Then '门诊号
        gstrSQL = "select 病人id,姓名,性别,年龄,出生日期,来源ID,主页ID,病人科室ID,当前科室ID,医生,门诊号,住院号,就诊卡号,就诊状态,卡验证码,当前床号,费别" & _
                        ",医疗付款方式,身份证号,民族,职业,婚姻状况,电话,邮编,地址,合同单位ID, 新病人" & _
                    " From (Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室ID,A.当前科室ID,B.执行人 As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,B.登记时间,A.合同单位ID, 0 as 新病人" & _
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
                    
    ElseIf mInputType = 5 Then '收费单据号
        strNO = zlCommFun.GetFullNO(strSeek, 13)
        PatiIdentify.Text = strNO
        mstrChargeNo = strNO
        
        '门诊费用记录的NO=病人挂号记录的NO，所以使用收费单据号提取病人的时候，同时记录挂号单。
        '如果没有挂号单为空，则通过收费单据号提取并登记的门诊病人，看不到医嘱内容。
'        mstrRegNo = strNO
        
        gstrSQL = "Select Distinct Nvl(A.病人id, 0) 病人id, Nvl(A.姓名, B.姓名) 姓名, Nvl(A.性别, B.性别) 性别, Nvl(A.年龄, B.年龄) 年龄," & vbNewLine & _
                    "                To_Char(A.出生日期, 'yyyy-mm-dd') 出生日期, Decode(Nvl(A.在院, 0), 0, 1, 2) As 来源id, A.主页id," & vbNewLine & _
                    "                Nvl(B.开单部门id, B.病人科室id) As 病人科室id, A.当前科室ID,Nvl(B.开单人, B.执行人) As 医生, Nvl(A.门诊号, B.标识号) 门诊号, A.住院号, A.就诊卡号, decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码," & vbNewLine & _
                    "                A.当前床号, A.费别, A.医疗付款方式, A.身份证号, A.民族, A.职业, A.婚姻状况, Nvl(A.家庭电话, A.联系人电话) 电话, Nvl(A.家庭地址邮编, A.单位邮编) 邮编," & vbNewLine & _
                    "                Nvl(A.家庭地址, A.工作单位) 地址, A.合同单位id, 0 as 新病人" & vbNewLine & _
                    "From 病人信息 A, 门诊费用记录 B" & vbNewLine & _
                    "Where B.NO = [3] And Mod(B.记录性质,10) = 1 And B.记录状态 = 1 And Nvl(B.费用状态,0) <>1 And B.病人id = A.病人id(+) And '%' = '%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 6 Then '当作姓名
            gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Nvl(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                " From 病人信息 A Where " & IIf(mblnLike = False, "A.姓名=[1]", IIf(mlngLike = 0, "instr(A.姓名,[1])>0", "A.登记时间 Between sysdate-" & mlngLike & " and sysdate and instr(A.姓名,[1])>0"))
    ElseIf mInputType = 7 Then '医保号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Nvl(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A Where A.医保号=[1] and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 8 Then '身份证号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Nvl(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A Where A.身份证号=[1] and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    ElseIf mInputType = 9 Then 'IC卡号
        gstrSQL = "Select distinct A.病人id,A.姓名,A.性别,A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,Decode(A.当前科室id,Null,1,2) As 来源ID,A.主页ID," & _
                        "Nvl(A.当前科室id,0) As 病人科室ID,A.当前科室ID,'' As 医生,A.门诊号,A.住院号,A.就诊卡号,decode(A.就诊状态,0,'正常',1,'等待就诊',2,'正在就诊','')as 就诊状态,A.卡验证码,A.当前床号," & _
                        "A.费别,A.医疗付款方式,A.身份证号,A.民族,A.职业,A.婚姻状况,Nvl(A.家庭电话,A.联系人电话) 电话," & _
                        "Nvl(A.家庭地址邮编,A.单位邮编) 邮编,Nvl(A.家庭地址,A.工作单位) 地址,A.合同单位ID, 0 as 新病人" & _
                  " From 病人信息 A Where A.IC卡号=[1] and '%'='%'" '为免避一行也弹出窗口所以用%,%在ShowSQLSelect在限制
    End If


    gstrSQL = gstrSQL & " Union " & _
                "Select null 病人ID,'新病人' 姓名,'未知' 性别,'' 年龄,null 出生日期,3 As 来源ID,0 As 主页ID," & _
                        "0 As 病人科室ID,0 As 当前科室ID,'' As 医生,null as 门诊号,null as 住院号,'' as 就诊卡号, '' as 就诊状态,'' 卡验证码,'' as 当前床号," & _
                        "'' as 费别,'' as 医疗付款方式,'' as 身份证号,'' as 民族,'' as  职业,'' as 婚姻状况,'' 电话,'' 邮编,'' 地址,0 合同单位ID, 1 as 新病人" & _
             " From dual where '%'='%'"
    gstrSQL = "select RowNum as ID,病人id,姓名,性别,年龄,出生日期,来源ID,主页ID,病人科室ID,当前科室ID,医生,门诊号," & _
                "住院号,就诊卡号,就诊状态,卡验证码,当前床号,费别,医疗付款方式,身份证号,民族,职业,婚姻状况,电话,邮编,地址,合同单位ID" & _
                " From (" & gstrSQL & ") Order by 新病人 asc, 病人ID desc"
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
        
    strSql = "Select 编码,Nvl(名称,'未知') as 名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
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
    
    strSql = "Select " & vbNewLine & _
                "Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b, 人员性质说明 c" & vbNewLine & _
                " Where a.人员id = b.Id And b.Id = c.人员id And c.人员性质 = '医生' And" & vbNewLine & _
                "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and a.部门id = [1] " & vbNewLine & _
                " Order By 简码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID)
    cbo医生.Clear
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cbo医生.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            If rsTmp!ID = UserInfo.ID Then cbo医生.ListIndex = cbo医生.NewIndex
            rsTmp.MoveNext
        Loop
        If cbo医生.ListCount > 0 And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = 0
        cbo医生.Enabled = True
    End If
End Sub
Private Sub InitInput()
    Dim i As Integer, strInput As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select ID ,科室ID,参数值 from 影像流程参数 where 科室ID = [1] and 参数名 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId, CStr("输入控制"))
    If Not rsTemp.EOF Then
        strInput = Nvl(rsTemp!参数值)
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
'            Case "执行间"
            Case "紧急"
                chk紧急.TabStop = False
            Case "申请时间"
                dtp(0).TabStop = False
        End Select
    Next
End Sub




Private Sub InitFaceScheme()
    '读取参数
    mblnNoshowReagent = Val(GetDeptPara(mlngCurDeptId, "显示造影剂", 1)) <> 1
    mblnNoshowAddons = Val(GetDeptPara(mlngCurDeptId, "显示附加主述", 1)) <> 1
    mintCheckInMode = Val(zlDatabase.GetPara("登记模式", glngSys, mlngModul, 2))
    
    mblnIsPetitionScan = IIf(Val(GetDeptPara(mlngCurDeptId, "启用申请单扫描", 1)) = 1, True, False)   '读取启用申请单扫描参数
    Me.cmdPetitionCapture.Visible = mblnIsPetitionScan
    
    If mintCheckInMode <> 1 Then mintCheckInMode = 2
    
    '因为附加主诉在造影剂的上方显示，所以先处理附加主诉
    If mblnNoshowAddons And Label29.Visible = True Then '不显示附加主诉，且附加主诉已经被显示，则关闭显示附加主诉
        Label29.Visible = False: txt附加主述.Visible = False: txt附加主述.Enabled = False
        '调整后面控件的位置
        Label1.Top = Label1.Top - 400: cbo费别.Top = cbo费别.Top - 400
        Label13.Top = Label13.Top - 400: cbo付款方式.Top = cbo付款方式.Top - 400
        Label12.Top = Label12.Top - 400: lblCash.Top = lblCash.Top - 400
        frm其他信息.Height = frm其他信息.Height - 400
        cmdOK.Top = cmdOK.Top - 400: CmdCancle.Top = cmdOK.Top: cmdPetitionCapture.Top = cmdOK.Top
        Me.Height = Me.Height - 400
    End If
    
    If mintCheckInMode = 1 Then     '精简模式
        frm其他信息.Visible = False
        cmdOK.Top = cmdOK.Top - frm其他信息.Height: CmdCancle.Top = cmdOK.Top: cmdPetitionCapture.Top = cmdOK.Top
        Me.Height = Me.Height - frm其他信息.Height
    End If
    
End Sub


Private Sub InitEdit(Optional blnSaveName As Boolean)
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Integer
    Dim curDate As Date
    
    On Error GoTo DBError
    
    If Not blnSaveName Then
        PatiIdentify.Text = ""
    End If
    PatiIdentify.tag = ""
    Txt英文名.Text = "":    Txt英文名.tag = ""
    txt年龄.Text = "":      cboAge.Visible = True
    Txt身高.Text = "":      Txt体重.Text = ""
    Txt身份证号.Text = "":  Txt电话.Text = ""
    Txt邮编.Text = "":      Txt联系地址 = ""
    txtPatientDept.Text = "":  txtID.Text = ""
    txtBed.Text = ""
    txt医嘱内容.Text = "":  txt医嘱内容.tag = ""
    Txt部位方法.Text = "":  Txt部位方法.tag = ""
    cboAge.ListIndex = 0
    
    txtPatholNum.Text = ""
    
    mstrExamineDoctor = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "待处理人", "")
    mstrUnitName = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "送检单位", "")
'    txtPatholNum.Enabled = False
'    cbxStudyType.Enabled = False
    
    '根据传入的图像数量来判断改变按钮的内容
    If mintEditMode > 0 Then cmdPetitionCapture.Caption = IIf(mintImgCount = 0, "申请单", "申请单(" & mintImgCount & "张)")
    
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
    
    curDate = zlDatabase.Currentdate
    
    dtp出生日期.value = Format(curDate, "yyyy-mm-dd")
    dtp(0).value = curDate
    dtp(1).value = Format(curDate, "yyyy-mm-dd HH:MM")

    InitInput '光标经过位置
    
    '登记的情况，需要控制控件的可用性
    If mintEditMode = 0 Then Call RefreshObjEnabled(1)
    
    '当无标本核收模块，且处于报到状态或者登记后直接报到且无标本核收模块时，自动生成病理号
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        '自动生成病理号
        txtPatholNum.Text = GetPatholNum(Val(cbxStudyType.Text))
    End If
    
    
    '若启用外院信息则加载送检科室
    If mblnShowSentInfo Then
        
        strSql = "Select B.名称 From 部门表 A , 部门表 B where  B.上级iD=A.Id and A.名称=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mStrOutSideCfg)
        cboUnitName.Clear
        
        If rsTmp.RecordCount > 0 Then
            While Not rsTmp.EOF
                cboUnitName.AddItem rsTmp!名称
                
                rsTmp.MoveNext
            Wend
        End If
        
        If mintEditMode = 0 Then
            cboUnitName.Text = mstrUnitName
        End If
        
        
    End If
    
Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
Public Function CopyCheck(ByVal lngAdviceID As Long, ByVal lngSendNO As Long) As Boolean
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
    
    gstrSQL = "SELECT Nvl(E.姓名,B.姓名) 姓名,Nvl(E.性别,B.性别) 性别,Nvl(E.年龄,B.年龄) 年龄,B.出生日期,B.费别,B.医疗付款方式,B.身份证号,B.民族,B.职业,Nvl(E.英文名,'') 英文名,E.身高,E.体重" & _
                    ",B.婚姻状况,Nvl(B.家庭电话,B.联系人电话) 电话,Nvl(B.家庭地址邮编,B.单位邮编) 邮编,Nvl(B.家庭地址,B.工作单位) 地址,B.合同单位ID,B.门诊号,B.就诊卡号,B.卡验证码" & _
                    ",Nvl(D.名称,'') AS 病人科室,A.病人科室ID,Decode(A.病人来源,2,B.住院号,B.门诊号) As 病人号,Decode(B.住院号,NULL,NULL,B.当前床号) As 床号" & _
                    ",F.发送时间 开嘱时间,Nvl(C.简码,0) 科室简码,Nvl(C.名称,'未知') AS 开嘱科室,A.开嘱时间 As 申请时间,A.开嘱医生,A.紧急标志,F.首次时间,F.执行间,E.检查设备,A.医嘱内容,E.检查号,E.检查技师" & _
                    ",DECODE(A.病人来源,2,2,1,1,4,4,3) AS 病人来源,Nvl(E.影像类别,G.影像类别) As 影像类别,B.病人id,A.主页id, Nvl(A.婴儿,0) As 婴儿,A.诊疗项目ID,E.附加主述" & _
                " FROM 病人医嘱发送 F,病人医嘱记录 A, 病人信息 B,部门表 C,部门表 D,影像检查记录 E,影像检查项目 G " & _
                " Where F.医嘱ID=[1] And F.发送号=[2] AND F.医嘱ID=A.ID" & _
                        " AND F.医嘱ID=E.医嘱ID(+) And F.发送号=E.发送号(+)  And A.病人ID=B.病人ID" & _
                        " And A.开嘱科室ID=C.ID And A.病人科室ID=D.ID And A.诊疗项目ID=G.诊疗项目ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", lngAdviceID, lngSendNO)

    If rsTemp.EOF Then
        '检查病人信息不完整的原因，如果是没有“病人医嘱发送记录，则提示本次医嘱已被回退或作废
        gstrSQL = "Select 医嘱ID From 病人医嘱发送 Where 医嘱ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查医嘱状态", lngAdviceID)
        If rsTemp.EOF Then
            Call MsgBoxD(Me, "本次检查医嘱没有发送记录，可能是该医嘱已经被回退或者已作废，请刷新后检查医嘱状态！", vbInformation, gstrSysName)
        Else
            Call MsgBoxD(Me, "病人信息不完整，请与管理员联系！", vbInformation, gstrSysName)
        End If
        
        mblnOK = False
        cmdOK.Enabled = False
        Exit Function
    End If
    
    curDate = zlDatabase.Currentdate
    
    '如果是婴儿则应该显示婴儿的信息
    mlngBaby = rsTemp!婴儿
    
    If mlngBaby = 0 Then
Normal:
        PatiIdentify.Text = Nvl(rsTemp!姓名)
        Call SeekIndex(cbo性别, Nvl(rsTemp!性别), True)
        
        If Nvl(rsTemp!年龄) <> "" Then
            LoadOldData rsTemp!年龄, txt年龄, cboAge
        Else
            '没有年龄情况下的处理方式
            Call ReCalcOld(Format(Nvl(rsTemp!出生日期, curDate), "yyyy-mm-dd hh:mm:ss"), _
                        cboAge, _
                        0, _
                        Format(curDate, "yyyy-mm-dd hh:mm:ss"))
        End If
 
        dtp出生日期.CustomFormat = "yyyy-MM-dd"
    
        If Trim(Nvl(rsTemp!出生日期)) = "" Then
            Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
        Else
            '判断出生日期是否满一年，不满一年按则可能是婴儿
            If zlDatabase.Currentdate - CDate(rsTemp!出生日期) >= 366 Then
                dtp出生日期.value = Format(Nvl(rsTemp!出生日期), "yyyy-mm-dd")
            Else
                dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
                dtp出生日期.value = Format(Nvl(rsTemp!出生日期), "yyyy-mm-dd hh:mm:ss")
            End If
        End If
        
        
    Else
        lngPageID = Nvl(rsTemp!主页ID, 0)
        
        strSql = "Select A.开嘱时间 As 申请时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                 "  Where a.病人ID = b.病人ID And b.主页id = [2] And b.序号 = [3] And a.ID = [1]"

        Set rsBaby = zlDatabase.OpenSQLRecord(strSql, "提取婴儿信息", lngAdviceID, lngPageID, mlngBaby)
            
        If rsBaby.EOF Then
            GoTo Normal
        Else
            PatiIdentify.Text = Nvl(rsBaby!婴儿姓名)
            Call SeekIndex(cbo性别, Nvl(rsBaby!婴儿性别), True)

            If Trim(Nvl(rsTemp!出生日期)) <> "" Then
                txt年龄.Text = ReCalcOld(Format(Nvl(rsBaby!出生时间, curDate), "yyyy-mm-dd hh:mm:ss"), _
                                            cboAge, _
                                            0, _
                                            Format(curDate, "yyyy-mm-dd hh:mm:ss"))
            End If
            
            '如果是婴儿，出生日期需要显示时分秒
            dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm"
            dtp出生日期.value = Format(Nvl(rsBaby!出生时间), "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    
    dtp(0).value = Format(curDate, "yyyy-mm-dd HH:MM")
    
    Txt英文名 = Decode(Nvl(rsTemp!英文名), "", zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter), rsTemp!英文名)
    
    If Trim(txt年龄) = "" Then txt年龄 = 0
    Txt身高 = Nvl(rsTemp!身高): Txt体重 = Nvl(rsTemp!体重)
    
    Call SeekIndex(cbo费别, Nvl(rsTemp!费别), True)
    Call SeekIndex(cbo付款方式, Nvl(rsTemp!医疗付款方式), True)
    Txt身份证号 = Nvl(rsTemp!身份证号)
    Call SeekIndex(cbo民族, Nvl(rsTemp!民族), True)
    Call SeekIndex(cbo职业, Nvl(rsTemp!职业), True)
    Call SeekIndex(cbo婚姻, Nvl(rsTemp!婚姻状况), True)
    Txt电话 = Nvl(rsTemp!电话): Txt邮编 = Nvl(rsTemp!邮编)
    Txt联系地址 = Nvl(rsTemp!地址)
    Label22.tag = Nvl(rsTemp!合同单位ID, 0)
    
    txtPatientDept.Text = Nvl(rsTemp!病人科室)
    txtPatientDept.tag = Nvl(rsTemp!病人科室ID, 0)
    txtID = Nvl(rsTemp!病人号): txtBed = Nvl(rsTemp!床号)
'    Call SeekIndex(cbo开单科室, Nvl(rsTemp!科室简码), True, , True)
    Call SeekIndex(cbo开单科室, Nvl(rsTemp!开嘱科室), True, , TNeedType.tNeedName)
    Call SeekIndex(cbo医生, Nvl(rsTemp!开嘱医生), True)
    '查找不到开嘱医生，且开嘱医生不为空，则直接填写开嘱医生字段
    If Nvl(rsTemp!开嘱医生) <> "" And cbo医生.ListIndex = -1 Then
        cbo医生.Text = Nvl(rsTemp!开嘱医生)
    End If
    
    chk紧急.value = Nvl(rsTemp!紧急标志, 0)
    dtp(1).value = Format(curDate, "yyyy-mm-dd HH:MM")
    
    txt附加主述.Text = Nvl(rsTemp!附加主述)
    '医嘱内容　诊疗名称,床旁/术中:部位1(方法1),部位1(方法2),部位2(方法1)---
    txt医嘱内容 = Split(Split(rsTemp!医嘱内容, ":")(0), ",")(0)
    
    mstrOutNo = Nvl(rsTemp!门诊号, 0)
    mstrCardNo = Nvl(rsTemp!就诊卡号)
    mstrCardPass = Nvl(rsTemp!卡验证码)
    mintSourceType = rsTemp!病人来源
    
    If mblnAllPatientIsOutside Then mintSourceType = 3
    
    mlngPatiId = Nvl(rsTemp!病人ID, 0)
    mlngPageID = Nvl(rsTemp!主页ID, 0)
    mlngClinicID = Nvl(rsTemp!诊疗项目ID)
    
    txt医嘱内容.TabIndex = 0
    
    '复制完之后，设置控件的可用性
    Call RefreshObjEnabled(3)
    
    CopyCheck = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function RefreshPatiInfor(bln报到 As Boolean) As Boolean
'功能:用于报到或修改时刷新病人
'bln报到=True，是报到，则部分信息可以直接使用默认信息
'bln报到=False,是修改，则信息应该全部使用数据库中的信息

Dim rsTemp As New ADODB.Recordset
Dim rsSongJian As ADODB.Recordset
Dim strSql As String
Dim rsBaby As New ADODB.Recordset
Dim lngPatientID As Long
Dim lngPageID As Long
Dim intChargeType As Integer    '病人医嘱发送.记录性质---1-收费记录；2-记帐记录。
Dim intChargeState As ChargeState
Dim curDate As Date
Dim intTMP As Integer
Dim i As Integer
Dim strTmp As String


    On Error GoTo errHand
    
    RefreshPatiInfor = False
    
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "SELECT I.名称 as 号别名称,H.病理号,H.检查类型,Nvl(E.姓名,A.姓名) 姓名,Nvl(E.性别,A.性别) 性别,Nvl(E.年龄,A.年龄) 年龄,B.出生日期,B.入院时间,B.费别,B.医疗付款方式,B.身份证号,B.民族,B.职业,Nvl(E.英文名,'') 英文名,E.身高,E.体重" & _
                    ",E.待处理人,B.婚姻状况,Nvl(B.家庭电话,B.联系人电话) 电话,Nvl(B.家庭地址邮编,B.单位邮编) 邮编,Nvl(B.家庭地址,B.工作单位) 地址,B.合同单位ID,B.门诊号,B.就诊卡号,B.卡验证码" & _
                    ",Nvl(D.名称,'') AS 病人科室,A.病人科室ID,Decode(A.病人来源,2,B.住院号,B.门诊号) As 病人号,Decode(B.住院号,NULL,NULL,B.当前床号) As 床号,B.当前病区ID" & _
                    ",F.发送时间 开嘱时间,Nvl(C.简码,0) 科室简码,Nvl(C.名称,'未知') AS 开嘱科室,A.开嘱时间 As 申请时间,A.开嘱医生,A.紧急标志,F.首次时间,F.执行间,E.检查设备,A.医嘱内容,E.报告人,E.检查号,E.检查技师" & _
                    ",DECODE(A.病人来源,2,2,1,1,4,4,3) AS 病人来源,Nvl(E.影像类别,G.影像类别) As 影像类别,B.病人id,A.主页id,A.诊疗项目ID,E.附加主述,Nvl(A.婴儿, 0) As 婴儿" & _
                    ",F.记录性质 " & _
                " FROM 病人医嘱发送 F,病人医嘱记录 A, 病人信息 B,部门表 C,部门表 D,影像检查记录 E,影像检查项目 G, 病理检查信息 H ,病理号码规则 I " & _
                " Where F.医嘱ID=[1] And F.发送号=[2] AND F.医嘱ID=A.ID And F.医嘱ID=H.医嘱ID(+) And H.号码规则ID=I.ID(+) " & _
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
    
        mblnOK = False
        cmdOK.Enabled = False
        Exit Function
    End If
    
    txt年龄.tag = ""
    
    '处理婴儿信息
    mlngBaby = rsTemp!婴儿
    If mlngBaby = 0 Then
Normal:
        PatiIdentify.Text = Nvl(rsTemp!姓名)
        Call SeekIndex(cbo性别, Nvl(rsTemp!性别), True)
        
        If bln报到 Or mintEditMode = 1 Then
            If bln报到 And IsNull(rsTemp!出生日期) Then
  
                intTMP = MsgBoxD(Me, "出生日期为空，是否根据年龄推算？", vbYesNo, gstrSysName)
                
                If intTMP = vbYes Then
                    txt年龄.Text = Nvl(rsTemp!年龄, "")
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
        
        
        If Nvl(rsTemp!年龄) <> "" Then
            LoadOldData rsTemp!年龄, txt年龄, cboAge
        Else
            '没有年龄情况下的处理方式
            Call ReCalcOld(Format(Nvl(rsTemp!出生日期, curDate), "yyyy-mm-dd hh:mm:ss"), _
                        cboAge, _
                        0, _
                        Format(Nvl(rsTemp!申请时间, curDate), "yyyy-mm-dd hh:mm:ss"))
        End If

 
        dtp出生日期.CustomFormat = "yyyy-MM-dd"
        
        If Trim(Nvl(rsTemp!出生日期)) = "" Then
            Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
        Else
            '判断出生日期是否满一年，不满一年按则可能是婴儿
            If zlDatabase.Currentdate - CDate(rsTemp!出生日期) >= 366 Then
                dtp出生日期.value = Format(Nvl(rsTemp!出生日期), "yyyy-mm-dd")
            Else
                dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
                dtp出生日期.value = Format(Nvl(rsTemp!出生日期), "yyyy-mm-dd hh:mm:ss")
            End If
        End If
    Else
        lngPatientID = rsTemp!病人ID
        lngPageID = Nvl(rsTemp!主页ID, 0)

        strSql = "Select A.开嘱时间 As 申请时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                 "  Where a.病人ID = b.病人ID And b.主页id = [2] And b.序号 = [3] And a.ID = [1]"

        Set rsBaby = zlDatabase.OpenSQLRecord(strSql, "提取婴儿信息", mlngAdviceID, lngPageID, mlngBaby)
        
        If rsBaby.EOF Then
            GoTo Normal
        Else
            PatiIdentify.Text = Nvl(rsBaby!婴儿姓名)
            Call SeekIndex(cbo性别, Nvl(rsBaby!婴儿性别), True)
            
            If bln报到 Or mintEditMode = 1 Then
                If bln报到 And IsNull(rsBaby!出生时间) Then
                    MsgBoxD Me, "婴儿无出生日期数据，不能进行报到，不能进行报到，请联系对应人员处理。", vbInformation, gstrSysName
                    RefreshPatiInfor = False
                    Exit Function
                End If
            End If
            
            If mintEditMode > 2 Then
                '大于2表明是报到后修改
                Call LoadOldData(Nvl(rsTemp!年龄), txt年龄, cboAge)
            Else
                txt年龄.Text = ReCalcOld(Format(Nvl(rsBaby!出生时间, curDate), "yyyy-mm-dd hh:mm:ss"), _
                                            cboAge, _
                                            0, _
                                            Format(Nvl(rsBaby!申请时间, curDate), "yyyy-mm-dd hh:mm:ss"))
            End If
            
            '如果是婴儿，出生日期需要显示时分秒
            dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm"

            dtp出生日期.value = Format(Nvl(rsBaby!出生时间), "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    
    lblCash.tag = Nvl(rsTemp!当前病区ID)
    Txt英文名 = Decode(Nvl(rsTemp!英文名), "", zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter), rsTemp!英文名)
    If Trim(txt年龄) = "" Then txt年龄 = 0
    Txt身高 = Nvl(rsTemp!身高): Txt体重 = Nvl(rsTemp!体重)
    Call SeekIndex(cbo费别, Nvl(rsTemp!费别), True)
    Call SeekIndex(cbo付款方式, Nvl(rsTemp!医疗付款方式), True)
    Txt身份证号 = Nvl(rsTemp!身份证号)
    Call SeekIndex(cbo民族, Nvl(rsTemp!民族), True)
    Call SeekIndex(cbo职业, Nvl(rsTemp!职业), True)
    Call SeekIndex(cbo婚姻, Nvl(rsTemp!婚姻状况), True)
    
    If Not bln报到 Then Call SeekIndex(cbo待处理人, Nvl(rsTemp!待处理人), True)
    Txt电话 = Nvl(rsTemp!电话): Txt邮编 = Nvl(rsTemp!邮编)
    Txt联系地址 = Nvl(rsTemp!地址)
    Label22.tag = Nvl(rsTemp!合同单位ID, 0)

    If mintEditMode = 3 Then    '只有报到后修改时，才从数据库读取病理号
        For i = 0 To cbxStudyType.ListCount - 1
            If InStr(cbxStudyType.list(i), Nvl(rsTemp!号别名称)) > 0 Then
                cbxStudyType.ListIndex = i
                Exit For
            End If
            cbxStudyType.ListIndex = 0
        Next
        txtPatholNum.Text = Nvl(rsTemp!病理号)
    End If
    
    If mblnShowSentInfo Then   '当显示送检信息时，才读取送检信息数据
        '启用了显示送检信息，则可以读取送检信息
        strSql = "select 送检单位, 送检科室,送检人,备注 from 病理送检信息 where  医嘱ID=[1] and 送检日期=to_date('1000/10/10 10:10:10','yyyy/mm/dd hh24:mi:ss')"
        Set rsSongJian = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        
        If rsSongJian.RecordCount > 0 Then
            
            Call SeekIndex(cboUnitName, Nvl(rsSongJian!送检单位), False, , tNeedAll) '送检单位
            If cboUnitName.Text <> Nvl(rsSongJian!送检单位) Then
                cboUnitName.AddItem Nvl(rsSongJian!送检单位)
                Call SeekIndex(cboUnitName, Nvl(rsSongJian!送检单位), False, , tNeedAll) '送检单位
            End If

            txtFormDepart.Text = Nvl(rsSongJian!送检科室)
            txtSubmitDoctor.Text = Nvl(rsSongJian!送检人)

            strTmp = Nvl(rsSongJian!备注)
            If UBound(Split(strTmp, "|")) < 2 Then
                txtSendTag.Text = strTmp
            Else
                txtOldStudyNo.Text = Split(strTmp, "|")(0)
                txtOldBarCode.Text = Split(strTmp, "|")(1)
                txtSendTag.Text = Split(strTmp, "|")(2)
            End If
            
        End If
        
    End If

    
    txtPatientDept.Text = Nvl(rsTemp!病人科室)
    txtPatientDept.tag = Nvl(rsTemp!病人科室ID, 0)
    txtID = Nvl(rsTemp!病人号): txtBed = Nvl(rsTemp!床号)
    dtp(0).value = Format(rsTemp!申请时间, "yyyy-mm-dd HH:MM")
'    Call SeekIndex(cbo开单科室, Nvl(rsTemp!科室简码), True, , True)
    Call SeekIndex(cbo开单科室, Nvl(rsTemp!开嘱科室), True, , TNeedType.tNeedName)
    Call SeekIndex(cbo医生, Nvl(rsTemp!开嘱医生), True)
    '查找不到开嘱医生，且开嘱医生不为空，则直接填写开嘱医生字段
    If Nvl(rsTemp!开嘱医生) <> "" And cbo医生.ListIndex = -1 Then
        cbo医生.Text = Nvl(rsTemp!开嘱医生)
    End If
    
    chk紧急.value = Nvl(rsTemp!紧急标志, 0)
    dtp(1).value = Format(curDate, "yyyy-mm-dd HH:MM")
    
    
    txt附加主述.Text = Nvl(rsTemp!附加主述)
    '医嘱内容　诊疗名称,床旁/术中:部位1(方法1),部位1(方法2),部位2(方法1)---
    txt医嘱内容 = Split(Split(rsTemp!医嘱内容, ":")(0), ",")(0)
    txt医嘱内容.tag = txt医嘱内容.Text
    If InStr(Nvl(rsTemp!医嘱内容, ""), ":") > 0 Then
        Txt部位方法 = Replace(Split(rsTemp!医嘱内容, ":")(1), "),", ")" & vbCrLf)
    Else
        Txt部位方法 = Nvl(rsTemp!医嘱内容, "")
    End If
    
    mstrOutNo = Nvl(rsTemp!门诊号, 0)
    mstrCardNo = Nvl(rsTemp!就诊卡号)
    mstrCardPass = Nvl(rsTemp!卡验证码)
    mintSourceType = rsTemp!病人来源
    mlngPatiId = Nvl(rsTemp!病人ID, 0)
    mlngPageID = Nvl(rsTemp!主页ID, 0)
'    mstrItemType = Nvl(rsTemp!影像类别)
    mlngClinicID = Nvl(rsTemp!诊疗项目ID)
    
    If mintSourceType = 2 And mlngBaby = 0 Then
        txt年龄.tag = Nvl(rsTemp!入院时间)  '独立身份的患者才使用入院日期
    End If
    
    intChargeType = Nvl(rsTemp!记录性质, 1)
    
    gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人附件", mlngAdviceID)
    Txt部位方法 = Txt部位方法 & vbCrLf
    Do Until rsTemp.EOF
        Txt部位方法 = Txt部位方法 & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
        rsTemp.MoveNext
    Loop
    
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
            If mstrDefaultPatientType = Nvl(rsTemp!病人类型) Then
                If mblnOrdinaryNameColColorCfg Then
                    PatiIdentify.objTxtInput.ForeColor = zlDatabase.GetPatiColor(Nvl(rsTemp!病人类型))
                End If
            Else
                PatiIdentify.objTxtInput.ForeColor = zlDatabase.GetPatiColor(Nvl(rsTemp!病人类型))
            End If
        End If
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
            Or (gbln执行后审核 And (intChargeState = ChargeState.无费用 Or intChargeState = ChargeState.已记账)) _
            Or gbln执行前先结算 Then
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

Private Sub CmdCancle_Click()
    mblnOK = IIf(mlngGoOnReg = 1, True, False)
    Unload Me
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'功能：检查输入数据的合法性
'参数： 无
'返回：True--数据输入合格，可以继续；False --有数据输入不合格，需要修改数据
'------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    ValidData = False
    
    gstrSQL = "select ID ,科室ID,参数值 from 影像流程参数 where 科室ID = [1] and 参数名 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurDeptId, CStr("必录控制"))
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!参数值) <> "" Then
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
            ElseIf InStr(rsTemp!参数值, "附加主述") > 0 And Trim(txt附加主述.Text) = "" And txt附加主述.Enabled = True Then
                MsgBoxD Me, "必须输入附加主述，请检查！", vbInformation, gstrSysName: DoEvents
                txt附加主述.SetFocus: Exit Function
            End If
        End If
    End If

    On Error Resume Next
    
    '检查输入的年龄是否有效
    If mobjPublicPatient Is Nothing Then
        If MsgBoxD(Me, "未检测到部件zlPublicPatient.dll的有效注册信息，不能对年龄的有效性进行检查，是否继续？", vbYesNo + vbExclamation) = vbNo Then
            Exit Function
        End If
    Else
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
    
    If Len(Trim(Me.txt医嘱内容.tag)) = 0 Then
        MsgBoxD Me, "必须输入申请项目！", vbInformation, gstrSysName: DoEvents
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Me.cbo开单科室.ListIndex = -1 Then
        MsgBoxD Me, "请指定申请科室！", vbInformation, gstrSysName: DoEvents
        Me.cbo开单科室.SetFocus: Exit Function
    End If
    If Len(Trim(Me.cbo医生.Text)) = 0 Then
        MsgBoxD Me, "请指定申请医生！", vbInformation, gstrSysName: DoEvents
        Me.cbo医生.SetFocus: Exit Function
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
    
    ValidData = True
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

    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "待处理人", zlStr.NeedName(cbo待处理人.Text)
    
    Set mobjSquareCard = Nothing
    Set mobjPublicPatient = Nothing
    Set mobjInsure = Nothing
    
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

Private Sub Txt电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
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
            If Nvl(rsTmp!姓名) <> "新病人" Then
                curDate = zlDatabase.Currentdate
                
                PatiIdentify.tag = Trim(Nvl(rsTmp!姓名))
                PatiIdentify.Text = Trim(Nvl(rsTmp!姓名))
                Call SeekIndex(cbo性别, Nvl(rsTmp!性别), True)
                
                dtp出生日期.CustomFormat = "yyyy-MM-dd"
                
                If Nvl(rsTmp!出生日期) <> "" Then
                    '判断出生日期是否满一年，不满一年按则可能是婴儿
                    If zlDatabase.Currentdate - CDate(rsTmp!出生日期) >= 366 Then
                        dtp出生日期.value = Format(Nvl(rsTmp!出生日期), "yyyy-mm-dd")
                    Else
                        dtp出生日期.CustomFormat = "yyyy-MM-dd HH:mm:ss"
                        dtp出生日期.value = Format(Nvl(rsTmp!出生日期), "yyyy-mm-dd hh:mm:ss")
                    End If
                Else
                    dtp出生日期.value = Format(Nvl(rsTmp!出生日期, curDate), "yyyy-mm-dd")
                End If
                
                strSql = ""
                If Val(Nvl(rsTmp!当前科室ID, 0)) <> 0 Then
                    '住院病人年龄提取
                    strSql = "select 年龄,入院日期 As 时间 from 病案主页 where 病人ID=[1] and 主页ID=[2] and rownum <=1"
                Else
                    If Val(Nvl(rsTmp!就诊状态, 0)) <> 0 Then
                        '门诊患者年龄提取
                        strSql = "select 年龄,登记时间 As 时间 from 病人挂号记录 where 病人ID=[1] and 门诊号=[3] and rownum <=1"
                    End If
                End If
                
                If Trim(strSql) <> "" Then
                    Set rsAge = zlDatabase.OpenSQLRecord(strSql, "提取患者年龄", Val(Nvl(rsTmp!病人ID)), Val(Nvl(rsTmp!主页ID)), Val(Nvl(rsTmp!门诊号)))
                    If rsAge.RecordCount > 0 Then
                        LoadOldData Nvl(rsAge!年龄), txt年龄, cboAge
                    End If
                    
                    If Nvl(rsTmp!来源ID, 0) = 2 Then
                        txt年龄.tag = Nvl(rsAge!时间)
                    End If
                Else
                    '根据当前日期计算年龄
                    txt年龄.Text = ReCalcOld(dtp出生日期.value, cboAge, 0, dtp(0).value) & cboAge.Text '加上单位，不然分别配置txt年龄和cboage时，会将单位改成年，bugNo 89538
                End If
                
                If txt年龄.Text = "" Then
                    txt年龄 = 0
                    cboAge.Visible = True
                    cboAge.ListIndex = 0
                End If
                
                Call SeekIndex(cbo费别, Nvl(rsTmp!费别, "普通"))
                Call SeekIndex(cbo付款方式, Nvl(rsTmp!医疗付款方式, "自费医疗"))
                Txt身份证号 = Nvl(rsTmp!身份证号)
                Call SeekIndex(cbo民族, Nvl(rsTmp!民族, "汉族"))
                Call SeekIndex(cbo职业, Nvl(rsTmp!职业, "工人"))
                Call SeekIndex(cbo婚姻, Nvl(rsTmp!婚姻状况, "未婚"))
                Txt电话 = Nvl(rsTmp!电话)
                Txt邮编 = Nvl(rsTmp!邮编)
                Txt联系地址 = Nvl(rsTmp!地址)
                Label22.tag = Nvl(rsTmp!合同单位ID, 0)
                txtID = Decode(Nvl(rsTmp!住院号), "", Nvl(rsTmp!门诊号), Nvl(rsTmp!住院号))
                txtBed = Nvl(rsTmp!当前床号)

                mlngPatiId = Nvl(rsTmp!病人ID, 0)
                mintSourceType = Nvl(rsTmp!来源ID, 1)
                
                '对于非住院病人，需区分是门诊还是外来
                If mintSourceType <> 2 Then mintSourceType = getSourceType(rsTmp!病人ID)
                
                mlngPageID = Nvl(rsTmp!主页ID, 0)
                mstrOutNo = Nvl(rsTmp!门诊号, 0)
                mstrCardNo = Nvl(rsTmp!就诊卡号)
                mstrCardPass = Nvl(rsTmp!卡验证码)
                
                '显示病人科室
                txtPatientDept.Text = zlStr.NeedName(cbo开单科室)
                txtPatientDept.tag = Nvl(rsTmp!病人科室ID)
                If cbo性别.Enabled = True Then cbo性别.SetFocus
                
                Call RefreshObjEnabled(2)
                
                '提取病人信息完成后 自动反算病人出生日期
                If IsNumeric(txt年龄.Text) And Nvl(rsTmp!出生日期, "") = "" Then Call ReCalcBirthDay(txt年龄.Text, cboAge.Text)
                
                Exit Sub
            Else
                If cbo性别.Enabled = True And mblnIsSamePatient Then cbo性别.SetFocus
            End If
        End If
    End If
    
    '没查到按新登记病人算
    Dim strTmp As String
    strTmp = Trim(PatiIdentify.Text)
    
'        InitEdit
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
    
    strSql = "select NO from 病人挂号记录 where 病人ID=[1] and 执行状态<>-1 order by 登记时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取病人来源和挂号单", lngPatiID)
    
    If rsTemp.RecordCount > 0 Then
        getSourceType = 1
        mstrRegNo = Nvl(rsTemp!NO)
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
    mblnIsOutSideHosp = IIf(InStr(cbo开单科室.Text, "外院") > 0, True, False)
    
    If cbo开单科室.ListIndex > -1 Then InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub
Private Sub PatiIdentify_LostFocus()
    Txt英文名.Text = zlCommFun.mGetFullPY(PatiIdentify.Text, mintCapital, mblnUseSplitter)
    
    Call zlCommFun.OpenIme
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
        Call InitEdit(True)
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
Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    '如果开单科室选择的是 外院科室，那么跳过医生的简码查找功能，否则医生栏不能自由录入
    If Not mblnIsOutSideHosp Then
        Call zlControl.CboSetIndex(cbo医生.hWnd, zlControl.CboMatchIndex(cbo医生.hWnd, KeyAscii))
    End If
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo职业.hWnd, zlControl.CboMatchIndex(cbo职业.hWnd, KeyAscii))
End Sub

Public Function zlShowMe(frmParent As Form, ByVal strDefaultPatientType As String, ByVal blnIsBigFont As Boolean) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lngRowIndex As Long
    
    Set mfrmParent = frmParent
    
    mstrDefaultPatientType = strDefaultPatientType
    
    Set mobjPublicPatient = VBA.Interaction.GetObject("", "zlPublicPatient.clsPublicPatient")
    If mobjPublicPatient Is Nothing Then Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
    
    If Not mobjPublicPatient Is Nothing Then Call mobjPublicPatient.zlInitCommon(gcnOracle, glngSys)
    
    Call ConfigPopedomFace
    
    strSql = "Select Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b " & vbNewLine & _
                " Where a.人员id = b.Id And " & vbNewLine & _
                " (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and a.部门id = [1] " & vbNewLine & _
                " Order By 简码 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)

    If Not IsNull(rsTmp) Then
        '加载待处理人
        cbo待处理人.Clear
        cbo待处理人.Enabled = True
        Do Until rsTmp.EOF
            cbo待处理人.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            If rsTmp!ID = UserInfo.ID Then cbo待处理人.ListIndex = cbo待处理人.NewIndex
            rsTmp.MoveNext
        Loop
            
        Call SeekIndex(cbo待处理人, mstrExamineDoctor, True)

        If mfrmParent.ufgStudyList.DataGrid.Rows > 1 Then
            lngRowIndex = mfrmParent.ufgStudyList.FindRowIndex(mlngAdviceID, "医嘱ID", True)
            If lngRowIndex > 0 Then
                If mfrmParent.ufgStudyList.Text(lngRowIndex, "检查过程") = "已报到" Or _
                mfrmParent.ufgStudyList.Text(lngRowIndex, "检查过程") = "已检查" Or _
                (mfrmParent.ufgStudyList.Text(lngRowIndex, "检查过程") = "已登记" And mintEditMode = 2) Then
                    cbo待处理人.Enabled = True
                Else
                    cbo待处理人.Enabled = False
                End If
            Else
                cbo待处理人.Enabled = False
            End If
        Else
            cbo待处理人.Enabled = False
        End If
        
    End If
    
    Call SetFontSize(blnIsBigFont)
    
    Me.Show 1, mfrmParent
End Function

Private Sub cbo医生_DropDown()
On Error GoTo errHandle
    Call SendMessage(cbo医生.hWnd, &H160, 300, 0)
errHandle:
End Sub

Private Sub SetFontSize(ByVal blnIsBigFont As Boolean)
    Dim objControl As Object
    Dim lngLabFontSize As Long
    Dim lngTxtFontSize As Long
    
    lngLabFontSize = IIf(blnIsBigFont, 14, 12)
    lngTxtFontSize = IIf(blnIsBigFont, 12, 10.5)
    
    Label3.FontSize = lngLabFontSize
    Label6.FontSize = lngLabFontSize
    Label7.FontSize = lngLabFontSize
    
    Label11.FontSize = lngLabFontSize
    Label5.FontSize = lngLabFontSize
    Label4.FontSize = lngLabFontSize
    
    Label10.FontSize = lngLabFontSize
    Label6.FontSize = lngLabFontSize
    Label20.FontSize = lngLabFontSize
    
    Label19.FontSize = lngLabFontSize
    Label2.FontSize = lngLabFontSize
    
    lbl医嘱内容.FontSize = lngLabFontSize
    Label8.FontSize = lngLabFontSize
    Lbl部位方法.FontSize = lngLabFontSize
    
    lbl(6).FontSize = lngLabFontSize
    lbl(0).FontSize = lngLabFontSize
    labSendRoom.FontSize = lngLabFontSize
    labSendDoctor.FontSize = lngLabFontSize
    labSendUnit.FontSize = lngLabFontSize
    labOldStudyNo.FontSize = lngLabFontSize
    labPatholNum.FontSize = lngLabFontSize
    labStudyType.FontSize = lngLabFontSize
    labOldBarCode.FontSize = lngLabFontSize
    lab备注.FontSize = lngLabFontSize
    
    Label25.FontSize = lngLabFontSize
    Label14.FontSize = lngLabFontSize
    Label15.FontSize = lngLabFontSize
    
    
    Label17.FontSize = lngLabFontSize
    Label18.FontSize = lngLabFontSize
    Label21.FontSize = lngLabFontSize
    
    Label22.FontSize = lngLabFontSize
    Label29.FontSize = lngLabFontSize
    
    Label16.FontSize = lngLabFontSize
    
    Label1.FontSize = lngLabFontSize
    Label13.FontSize = lngLabFontSize
    Label12.FontSize = lngLabFontSize
    
    
    
    lab待处理人.FontSize = lngLabFontSize
    cbo待处理人.FontSize = lngLabFontSize
    
    chk紧急.FontSize = lngLabFontSize
    
    
    txtPatientDept.FontSize = lngTxtFontSize
    txtID.FontSize = lngTxtFontSize
    txtBed.FontSize = lngTxtFontSize
    lblCash.FontSize = lngTxtFontSize
    
    For Each objControl In Me.Controls
        If TypeName(objControl) = "TextBox" Then
            objControl.FontSize = lngTxtFontSize
        End If
        
        If TypeName(objControl) = "ComboBox" Then
            objControl.FontSize = lngTxtFontSize
        End If
        
        If TypeName(objControl) = "DTPicker" Then
            objControl.Font.Size = lngTxtFontSize
        End If
    Next
    
    CmdCancle.FontSize = lngTxtFontSize
    cmdOK.FontSize = lngTxtFontSize
    cmdPetitionCapture.FontSize = lngTxtFontSize
End Sub


Private Sub ConfigPopedomFace()
'配置权限界面
    Dim blnEnregPopedom As Boolean
    Dim i As Long
    
    '如果没有登记权限，则只允许对病理科内部的信息进行修改
    blnEnregPopedom = True ' CheckPopedom(mstrPrivs, "检查登记")
    
    Frame1.Enabled = blnEnregPopedom
    Frame2.Visible = blnEnregPopedom
    
    If Not blnEnregPopedom Then
        '无检查登记权限，但可在报到后对病理号进行修改
        txtPatholNum.Enabled = IIf(mintEditMode = 3, True, False)
        cbxStudyType.Enabled = IIf(mintEditMode = 3, True, False)
    End If
    
    frm其他信息.Visible = blnEnregPopedom And Not (mintCheckInMode = 1) 'mintCheckInMode=1表示精简模式
    
    If Not blnEnregPopedom Then
        framPatholInf.Top = Frame1.Top + Frame1.Height + 240
        
        cmdOK.Top = framPatholInf.Top + framPatholInf.Height + 240
        CmdCancle.Top = cmdOK.Top
        
        cmdPetitionCapture.Top = cmdOK.Top
        
        Me.Height = Frame1.Height + framPatholInf.Height + cmdOK.Height + 1080
        
        For i = 0 To Me.Controls.Count - 1
            If UCase(Me.Controls(i).Name) <> UCase("txtPatholNum") And UCase(Me.Controls(i).Name) <> UCase("cbxStudyType") Then
                On Error Resume Next
                Me.Controls(i).BackColor = Me.BackColor
            End If
        Next i
    End If
    
End Sub


Private Sub sutSetTxtEnable(thisBox As TextBox, blnEnable As Boolean)
    thisBox.Enabled = blnEnable
    If blnEnable = True Then
        thisBox.BackColor = vbWhite
    Else
        thisBox.BackColor = &H8000000B
    End If
End Sub
