VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPacsCustomFilter 
   Caption         =   "自定义查询"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPacsCustomFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退 出(&Q)"
      Height          =   375
      Left            =   10680
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "确  定(&O)"
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox cboSchemaName 
      Height          =   330
      Left            =   1200
      TabIndex        =   16
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdDelSchema 
      Caption         =   "删除方案(&D)"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveSchema 
      Caption         =   "保存方案(&S)"
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame fraControl 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   11775
      Begin VB.ComboBox cboRightBracket 
         Height          =   330
         ItemData        =   "frmPacsCustomFilter.frx":000C
         Left            =   9000
         List            =   "frmPacsCustomFilter.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboLeftBracket 
         Height          =   330
         ItemData        =   "frmPacsCustomFilter.frx":003C
         Left            =   120
         List            =   "frmPacsCustomFilter.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删 除(&D)"
         Height          =   375
         Left            =   10440
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添 加(&A)"
         Height          =   375
         Left            =   10440
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboQueryLink 
         Height          =   330
         ItemData        =   "frmPacsCustomFilter.frx":006C
         Left            =   7680
         List            =   "frmPacsCustomFilter.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboQueryData 
         Height          =   330
         Left            =   5160
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cboQueryWhere 
         Height          =   330
         ItemData        =   "frmPacsCustomFilter.frx":0086
         Left            =   3840
         List            =   "frmPacsCustomFilter.frx":009F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboQueryField 
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lab 
         Caption         =   "右括号："
         Height          =   255
         Index           =   5
         Left            =   9000
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lab 
         Caption         =   "左括号："
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lab 
         Caption         =   "连接条件："
         Height          =   255
         Index           =   3
         Left            =   7680
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lab 
         Caption         =   "查询数据："
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lab 
         Caption         =   "查询条件："
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lab 
         Caption         =   "查询字段："
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lvwQueryConfig 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label labSchema 
      Caption         =   "查询方案："
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPacsCustomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

