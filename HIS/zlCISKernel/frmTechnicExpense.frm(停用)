VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmTechnicExpense 
   AutoRedraw      =   -1  'True
   Caption         =   "病人计费处理"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "宋体"
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
      Caption         =   "成套(&T)"
      Height          =   375
      Left            =   90
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   525
      Width           =   1080
   End
   Begin VB.CommandButton cmdSaveWholeSet 
      Caption         =   "保存为成套收费项目(&W)"
      Height          =   375
      Left            =   1215
      TabIndex        =   52
      Top             =   525
      Width           =   2715
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   7875
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmTechnicExpense.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "病人余额"
            Object.ToolTipText     =   "病人余额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5010
      Width           =   11850
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
               Picture         =   "frmTechnicExpense.frx":1A92
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   9780
         TabIndex        =   21
         ToolTipText     =   "热键:Esc"
         Top             =   1785
         Width           =   1680
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   7965
         TabIndex        =   20
         ToolTipText     =   "热键：F2"
         Top             =   1785
         Width           =   1680
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   36
         ToolTipText     =   "清除:F6"
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
         Begin VB.CheckBox chk急诊 
            Caption         =   "急诊费用"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk加班 
            Caption         =   "加班(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo开单人 
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
               Name            =   "宋体"
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
            Caption         =   "婴儿费(&B)"
            Height          =   240
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl开单人 
            AutoSize        =   -1  'True
            Caption         =   "开单人"
            Height          =   240
            Left            =   5790
            TabIndex        =   38
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   240
            Left            =   8820
            TabIndex        =   37
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraDrawDept 
         Height          =   720
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Width           =   13575
         Begin VB.TextBox txt病人备注 
            BackColor       =   &H00E0E0E0&
            Height          =   360
            Left            =   1095
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   2700
         End
         Begin VB.Label lbl病人备注 
            Caption         =   "病人备注"
            Height          =   225
            Left            =   105
            TabIndex        =   49
            Top             =   315
            Width           =   1005
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   42
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
         FormatString    =   "^         项目|^          金额"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         TabIndex        =   39
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
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1230
            Width           =   1845
         End
         Begin VB.TextBox txt实收 
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
         Begin VB.TextBox txt应收 
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
            Caption         =   "上张"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   51
            Top             =   1298
            Width           =   690
         End
         Begin VB.Label lbl实收 
            AutoSize        =   -1  'True
            Caption         =   "实收"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   41
            Top             =   818
            Width           =   690
         End
         Begin VB.Label lbl应收 
            AutoSize        =   -1  'True
            Caption         =   "应收"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   40
            Top             =   318
            Width           =   690
         End
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   30
      TabIndex        =   24
      ToolTipText     =   "清除:F6"
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
         Caption         =   "病人计费单"
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   28
         ToolTipText     =   "清除:F6"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9540
         TabIndex        =   25
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   9375
      TabIndex        =   23
      Top             =   855
      Width           =   2505
      Begin VB.ComboBox cbo开单科室 
         Height          =   360
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "cbo开单科室"
         Top             =   615
         Width           =   2265
      End
      Begin VB.Label lbl开单科室 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   22
      Top             =   855
      Width           =   9345
      Begin VB.TextBox txt住院号 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1290
      End
      Begin VB.TextBox txt费别 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txt付款方式 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   615
         Width           =   2085
      End
      Begin VB.TextBox txt性别 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox txt担保额 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   1290
      End
      Begin VB.TextBox txt担保人 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1110
      End
      Begin VB.TextBox txt床号 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt姓名 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   765
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   240
         Left            =   7140
         TabIndex        =   46
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl担保额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   240
         Left            =   7140
         TabIndex        =   45
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl担保人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   240
         Left            =   5145
         TabIndex        =   44
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl付款方式 
         Caption         =   "付款 方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   43
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   240
         Left            =   5385
         TabIndex        =   34
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   32
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Left            =   2370
         TabIndex        =   31
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Left            =   3705
         TabIndex        =   30
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3105
      Left            =   15
      TabIndex        =   11
      Top             =   1920
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5477
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "合计:"
      BeginProperty Font 
         Name            =   "宋体"
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
      TabIndex        =   33
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmTechnicExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入口参数
'————————————————————————————————————————————————————————————————————
Public mlng医嘱ID As Long '新增费用时用
Public mlng发送号 As Long '新增费用时用
Public mlng病人ID As Long '确定要计费的病人ID
Public mlng主页ID As Long '确定要计费的主页ID
Public mint病人来源 As Integer '1-门诊病人,2-住院病人
Public mint记录性质 As Integer '1-收费(划价),2-记帐(门/住)
Public mstrFeeTab As String
Public mbln费用登记 As Boolean '仅登记,不计实收金额
Public mlng开单科室ID As Long '为当前主界面医技科室
Public mlng病人科室id As Long '主要是用于确定门诊病人的科室ID

Public mlng开嘱科室ID As Long
Public mstr开嘱医生 As String

Public mbytInState As Byte '0-执行,1-查阅,2-调整(不支持),3-删费
Public mstrInNO As String '所操作的单据号(执行时为修改)
Public mstrOriginalNO As String '补充主费用时,医嘱发送中的单据号

Public mstrTime As String '操作单据内容的登记时间
Public mblnDelete As Boolean '是否处理退费单据(查阅)
Private mblnWarnCloseed As Boolean  '刘兴洪:因费用报敬发生的关闭
Private mblnSendMateria  As Boolean
Private mbytSendMateria As Byte '0-记帐后不发药,1-自动发药,2-提示发药
Private mbyt缺省科室 As Byte    '0-医技科室;1-病人科室
Private mobjBaseItem As Object
Private mstr住院医生 As String
Private mrsAll开单科室 As ADODB.Recordset
Private Enum BillColType       '单据控件的列类型
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum
Private Enum BillCol
    行 = 0
    类别 = 1
    项目 = 2
    商品名 = 3
    规格 = 4
    单位 = 5
    付数 = 6
    数次 = 7
    单价 = 8
    应收金额 = 9
    实收金额 = 10
    执行科室 = 11
    标志 = 12
    类型 = 13
End Enum

Private Enum Pan
    C2提示信息 = 2
End Enum

'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    实时监控 As Boolean
    医生确定处方类型 As Boolean '目前只有北京医保专用
End Type
Private MCPAR As TYPE_MedicarePAR
Private mrsDept As ADODB.Recordset
Public mstrPrivs As String
'医技工作站本地费用参数
'————————————————————————————————————————————————————————————————————
Private mstrLike As String '输入匹配方式
Private mblnPay As Boolean '中药是否输入付数
Private mblnTime As Boolean '变价是否输入付数
Private mbln其它药房 As Boolean '是否显示其它药房库存
Private mbln其它药库 As Boolean '是否显示其它药库库存
Private mstr收费类别 As String '可输入的收费类别
Private mbln药房单位 As Boolean '是否按照门诊单位或住院单位显示药品
Private mstr药房单位 As String '根据病人来源决定如"门诊单位"或"住院单位"
Private mstr药房包装 As String '根据病人来源决定如"门诊包装"或"住院包装"
Private mlngPreRow As Long '记录当前行,当仅改变列时
Private mlng西药房 As Long, mlng成药房 As Long, mlng中药房 As Long, mlng发料部门 As Long
'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsInfo As New ADODB.Recordset '病人信息
Private mrsMedAudit As ADODB.Recordset  '病人已审批的费用项目
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsClass As ADODB.Recordset '根据参数读取的当前可用的收费类别
Private mrsWork As New ADODB.Recordset '当天上班的药房
Private mblnWork As Boolean '当前是否有正在上班的药房
Private mlng药品类别ID As Long '当前单据操作的药品入出类别ID
Private mlng卫材类别ID As Long '当前单据操作的卫材入出类别ID
'程序对象
Private mobjBill As ExpenseBill '费用单据对象
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mobjDetail As Detail '单独的收费细目对象
Private mcolDetails As Details '单独的收费细目集合
Private mcolMoneys As BillInComes '收入项目汇总集合

'程序变量
Private mbytWarn As Byte '记帐报警的返回值
Private mintWarn As Integer '记帐报警提示的继续选项
Private mstrWarn As String '已经报过警并选择继续的类别
Private mrsWarn As New ADODB.Recordset '病区报警线
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库的出库检查方式

Private mcurModiMoney As Currency '修改单据时原单据的金额
Private mblnDrop As Boolean '在KeyDown中判断cbo开单人当前是否弹出
Private mblnNewRow As Boolean
Private mblnOne As Boolean '是否只有一个可用收费类别
Private marrColData() As Integer '当前单据编辑属性映象
Private mdblItemNum As Double '数据库中当前输入费目的数次
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器
Private marrDr() As String '记录医生的"ID|科室ID|编号|姓名|简码"
Private mblnEnterCell As Boolean '控制是否执行Entercell事件

Private Const STR_HEAD = "行,450,4;类别,750,1;项目,2175,1;商品名,2000,1;规格,1105,1;单位,520,4;付数,520,1;数次,570,1;单价,1055,7;" & "应收金额,1030,7;实收金额,1080,7;执行科室,1255,1;标志,520,4;类型,520,1"
Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    '刘兴洪 问题:27378 日期:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "执行科室"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "发药药店"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .收费类别) > 0 Then
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

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then
        mobjBill.开单部门ID = 0
        cbo开单科室.Text = ""
    End If
End Sub
Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln从项汇总折扣 As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Then Cancel = True: Exit Sub
    
    If mobjBill.Details.Count >= Row Then
        '带从属项目的项删除确认
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("该项目带有 " & bytsubs & " 个从属项目,删除该项目也将删除它的从属项目,继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).从属父号 <> 0 Then '从属项目删除确认
            If MsgBox("该项目是[" & mobjBill.Details(mobjBill.Details(Row).从属父号).Detail.名称 & "]的从属项目,确定要删除它吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln从项汇总折扣 = gbln从项汇总折扣
            End If
        ElseIf MsgBox("确实要删除该收费项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If

        If bln从项汇总折扣 Then lngMainRow = mobjBill.Details(Bill.Row).从属父号 '如果是从项,删除之前记下从项的从属父号,如果是主项,则级联删除,不用重算
        
        '删除处理
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).从属父号 = Row Then
                Call DeleteDetail(i) '反顺序删除其从属行
            End If
        Next
        Call DeleteDetail(Row) '删除该行
        
        '重新计算并刷新
        If bln从项汇总折扣 Then
            If CheckItemHaveSub(lngMainRow) Then
                Call Calc重算主项实收(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) '只有一个主项了,从项全部被删除时,当成普通独立项计算
            End If
        End If
        
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '不用控件来处理删除
        
        mlngPreRow = 0  '表示行改变了
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub ShowStock(str药品 As String, dbl库存 As Double)
'功能：显示药品或卫材的库存
    If InStr(1, mstrPrivs, "显示库存") > 0 Then
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & dbl库存
    Else
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]" & IIF(dbl库存 > 0, "有", "无") & "库存."
    End If
End Sub

Private Sub bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, i As Long
    
    '药品库存检查
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                    .执行部门ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID, 0)  '29680
                        If mbln药房单位 Then
                            dblStock = dblStock / .Detail.药房包装
                        End If
                        .Detail.库存 = dblStock  '记录当前行药品库存
                        Call ShowStock(.Detail.名称, .Detail.库存)
                        
                        '药房改变,实价药品重新计算价格
                        If .Detail.变价 Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                        .Detail.库存 = dblStock
                        Call ShowStock(.Detail.名称, .Detail.库存)
                        
                        '发料部门改变,时价卫材重新计算价格
                        If .Detail.变价 Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf InStr(",4,5,6,7,", .收费类别) = 0 Then
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub bill_CellCheck(Row As Long, Col As Long)
'说明：可以全部为主要手术,但不能全部为附加手术
    Dim i As Long, strCheck As String, bytTime As Byte
    
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    If mbytInState = 3 Then Exit Sub
    
    '新增的未处理行无效
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "F" And mobjBill.Details(i).附加标志 = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    If bytTime > 0 Then
        mobjBill.Details(Row).附加标志 = IIF(strCheck = "", 0, 1)
        Call CalcMoneys(Row)
        Call ShowDetails(Row)
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "单据中必然有一个手术不是附加手术！", vbInformation, gstrSysName
    End If
End Sub

Private Function SelectIsNurse() As Boolean
'功能：判断当前开单人是否护士
    Dim str性质 As String
    
    If cbo开单人.ListIndex <> -1 Then
        If cbo开单人.ItemData(cbo开单人.ListIndex) = 0 Then Exit Function
        
        If cbo开单人.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 6 Then
                str性质 = Split(marrDr(cbo开单人.ListIndex), "|")(6)
                SelectIsNurse = str性质 = "护士"
            End If
        End If
    End If
End Function

Private Sub bill_CommandClick()
    Dim lng项目ID As Long, blnCancel As Boolean
    Dim str类别 As String, str特准项目 As String, int险类 As Integer
    
    If gbln收费类别 Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str类别 = IIF(SelectIsNurse, "'E','M','4'", mstr收费类别)
        End If
    Else
        str类别 = IIF(SelectIsNurse, "'E','M','4'", mstr收费类别)
    End If
    If Not IsNull(mrsInfo!险类) Then
        int险类 = mrsInfo!险类
        '刘兴洪:24862
        'mint病人来源 As Integer '1-门诊病人,2-住院病人
        'mint记录性质 As Integer '1-收费(划价),2-记帐(门/住)
        If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), (mint记录性质 = 1 Or mint病人来源 = 1)) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
    End If
    
    lng项目ID = frmItemSelect.ShowSelect(Me, mstrPrivs, mint病人来源, int险类, True, str类别, , , str特准项目, zl获取中药形态(Bill.Row), IIF(mbln费用登记, True, False))
    If lng项目ID <> 0 Then
        Bill.Text = lng项目ID
        mblnSelect = True
        Call bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'功能：处理单据输入
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency
    Dim blnStock As Boolean, lngDoUnit As Long, str摘要 As String
    Dim lng项目ID As Long, str特准项目 As String, str类别 As String
    Dim blnInput As Boolean, cur余额 As Currency, lng病人科室ID As Long, int险类 As Integer, lngOld付数 As Long
    Dim colStock As Collection
    On Error GoTo errH
    
    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "执行科室" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = 0 Then Exit Sub
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "类别"
                If Bill.ListIndex <> -1 Then '如果不输入类别则不会进入类别列
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        '一旦改更收费类别,则清除(如有)原有该项目内容
                        For i = 2 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .收费细目ID = 0: .收费类别 = ""
                            End With
                            Call CalcMoneys
                            Call ShowMoney
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '暂时用RowData记录所选择的收费类别
                End If
            Case "项目"
                '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
                If Bill.Text <> "" Then
                    '如果在已输入的项目上按回车,或选择器选择
                    If mobjBill.Details.Count >= Bill.Row Then
                        '通过按钮选择是返回的ID,而输入则是文本,如果是一样的,则不改变
                        If Bill.TextMatrix(Bill.Row, BillCol.项目) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                
                    sta.Panels(2).Text = ""
                    sta.Panels(4).Text = ""
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '立即清除该标志
                        Set mobjDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        If gbln收费类别 Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "没有确定费用类别,请先输入类别！"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            str类别 = IIF(SelectIsNurse, "'E','M','4'", mstr收费类别)
                        End If
                        If Not IsNull(mrsInfo!险类) Then
                            int险类 = mrsInfo!险类
                           
                            '刘兴洪:24862
                            'mint病人来源 As Integer '1-门诊病人,2-住院病人
                            'mint记录性质 As Integer '1-收费(划价),2-记帐(门/住)
                            If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), (mint记录性质 = 1 Or mint病人来源 = 1)) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
                        End If
                        lng项目ID = frmItemSelect.ShowSelect(Me, mstrPrivs, mint病人来源, int险类, True, str类别, Bill.Text, Bill.TxtHwnd, str特准项目, zl获取中药形态(Bill.Row), IIF(mbln费用登记, True, False))
                        If lng项目ID <> 0 Then
                            Set mobjDetail = GetInputDetail(lng项目ID)
                            If int险类 <> 0 Then sta.Panels(4).Text = Get医保大类(lng项目ID, int险类)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Bill.TxtVisible = False '(不加不行)
                    
                    '医保费用项目是否审批检查
                    If mint病人来源 = 2 And mint记录性质 = 2 And Not IsNull(mrsInfo!险类) Then
                        If mobjDetail.要求审批 And Not mrsMedAudit Is Nothing Then
                            mrsMedAudit.Filter = "项目ID=" & mobjDetail.ID
                            If mrsMedAudit.RecordCount = 0 Then
                                MsgBox "当前病人未被批准使用该项目！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf Not IsNull(mrsMedAudit!可用数量) Then
                                If mrsMedAudit!可用数量 <= 0 Then
                                    MsgBox "当前病人使用[" & mobjDetail.名称 & "]已达到批准的使用限量" & FormatEx(mrsMedAudit!使用限量 / IIF(mbln药房单位, mobjDetail.药房包装, 1), 5) & "。", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    '主项适用病人病区科室
                    If mint病人来源 = 2 And mint记录性质 = 2 Then
                        If InStr(",5,6,7,", mobjDetail.类别) = 0 Then
                            If Not CheckFeeItemLimitDept(mobjDetail.ID) Then
                                MsgBox "该收费项目对当前病人病区和科室不适用！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    
                    '检查药品输入是否重复:分批及时价同一药房不允许重复(这里只提醒)
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 _
                        Or (mobjDetail.类别 = "4" And mobjDetail.跟踪在用) Then
                        If PhysicExist(mobjDetail, Bill.Row) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '检查处方职务
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        mobjDetail.处方职务 = Get处方职务(mobjDetail.ID)
                        '医保或公费病人
                        '问题:45605
                        If zlIsCheckMedicinePayMode(txt付款方式) Then
                            If CheckDuty(mobjDetail, False) > 0 Then
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                        '所有病人
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '病人科室ID
                    lng病人科室ID = mobjBill.科室ID
                    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    
                    '缺省执行科室
                    lngDoUnit = Val("" & mrsInfo!病区ID)
                    If mobjDetail.类别 = "4" And mlng发料部门 > 0 Then lngDoUnit = mlng发料部门
                    If lngDoUnit = 0 Then lngDoUnit = lng病人科室ID
                    
                    lngDoUnit = Get收费执行科室ID(mlng病人ID, mlng主页ID, mobjDetail.类别, mobjDetail.ID, _
                        mobjDetail.执行科室, lng病人科室ID, Get开单科室ID, mint病人来源, lngDoUnit, 1, 1)
                    
                    
                    '读取药品相关信息
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        '当前行药品库存
                        dblStock = GetStock(mobjDetail.ID, lngDoUnit, 0)  '29680
                        If mbln药房单位 Then
                            dblStock = dblStock / mobjDetail.药房包装
                        End If
                        mobjDetail.库存 = dblStock
                        Call ShowStock(mobjDetail.名称, mobjDetail.库存)

                        '处方限量
                        mobjDetail.处方限量 = Get处方限量(mobjDetail.ID)
                    ElseIf mobjDetail.类别 = "4" And mobjDetail.跟踪在用 Then
                        dblStock = GetStock(mobjDetail.ID, lngDoUnit, 0) ''29680
                        mobjDetail.库存 = dblStock
                        Call ShowStock(mobjDetail.名称, mobjDetail.库存)
                    End If
                    
                    '保险支付项目对应检查
                    If Not IsNull(mrsInfo!险类) Then
                        
                        If zlCheck定价零价格对码(mobjDetail.ID, Not mobjDetail.变价) Then
                            '问题:27286
                        Else
                            If Not ItemExistInsure(mrsInfo!病人ID, mobjDetail.ID, mrsInfo!险类) Then
                                If gint医保对码 = 1 Then
                                    If MsgBox("项目""" & mobjDetail.名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                    End If
                                ElseIf gint医保对码 = 2 Then
                                    MsgBox "项目""" & mobjDetail.名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    '输入摘要(取已有的行以便修改)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            str摘要 = mobjBill.Details(Bill.Row).摘要
                        End If
                    End If
                    
                    '加入或修改该收费细目行
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051
                    '输入摘要(根据新输入的行更改摘要)
                    If mobjBill.Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Details(Bill.Row).摘要 = str摘要
                        End If
                    ElseIf mint病人来源 = 2 And mrsInfo.State = 1 Then  '问题:主要是大连一院要求,由于BH不能登记,所以没有BugNo
                        str摘要 = gclsInsure.GetItemInfo(Val(Nvl(mrsInfo!险类)), mrsInfo!病人ID, mobjBill.Details(Bill.Row).收费细目ID, str摘要, 2)
                        mobjBill.Details(Bill.Row).摘要 = str摘要
                    End If
                    
                    Call CalcMoneys(Bill.Row)
                    
                    'Calcmoney中医保可能返回摘要
                    If mobjBill.Details(Bill.Row).摘要 <> "" Then str摘要 = mobjBill.Details(Bill.Row).摘要
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mint记录性质 = 2 And mrsWarn.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        '刘兴洪:30504: and mbln费用登记=False
                        If curTotal > 0 And mbln费用登记 = False Then
                            cur余额 = Val(txt实收.Tag)
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mint记录性质 = 2 Then
                        If Not IsNull(mrsInfo!险类) And mobjBill.Details(Bill.Row).数次 <> 0 And MCPAR.实时监控 Then
                            If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text), Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '费用类型检查
                    Call Check费用类型(Bill.Row)
                    

                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        '下一列的性质确定
                        If .收费类别 = "7" And mblnPay Then Bill.ColData(BillCol.付数) = 4 '付数
                        If .收费类别 = "F" Then Bill.ColData(BillCol.标志) = -1 '附加标志
                        
                        '变价允许输入数次
                        If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                            And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                            Bill.ColData(BillCol.数次) = IIF(mblnTime, 4, 5) '数次
                            Bill.ColData(BillCol.单价) = 4 '单价
                        Else
                            Bill.ColData(BillCol.数次) = 4 '数次
                            Bill.ColData(BillCol.单价) = 5 '单价
                        End If
                        
                        '执行科室
                        mblnEnterCell = False: Bill.Col = BillCol.执行科室: mblnEnterCell = True
                        Call FillBillComboBox(Bill.Row, BillCol.执行科室, Not blnInput) '直接回车时保持执行科室
                        mblnEnterCell = False: Bill.Col = BillCol.项目: mblnEnterCell = True
                        
                        blnSkip = Bill.ListCount = 1
                        If Not blnSkip And InStr(",4,5,6,7,", .收费类别) > 0 Then
                            '指定了固定药房时,不允许再选择
                            Select Case .收费类别
                                Case "4"
                                    blnSkip = mlng发料部门 > 0 And .执行部门ID = mlng发料部门
                                Case "5"
                                    blnSkip = mlng西药房 > 0 And .执行部门ID = mlng西药房
                                Case "6"
                                    blnSkip = mlng成药房 > 0 And .执行部门ID = mlng成药房
                                Case "7"
                                    blnSkip = mlng中药房 > 0 And .执行部门ID = mlng中药房
                            End Select
                        End If
                        If blnSkip Then
                            Bill.ColData(BillCol.执行科室) = 5: .Key = 1
                        Else
                            Bill.ColData(BillCol.执行科室) = 3: .Key = Bill.ListCount
                        End If
                        
                        '检查卫生材料的灭菌效期,在确定执行科室之后
                        If .收费类别 = "4" And .Detail.跟踪在用 Then
                            Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                        End If
                                                
                         '从属项目处理,仅该行收费项目有从属项目及尚未取才取,药品无需判断,药品不能设置主从项
                        If Bill.TextMatrix(0, Bill.Col) = "项目" And InStr(",5,6,7,", .收费类别) = 0 Then
                            If (gbln从项汇总折扣 And mobjBill.Details(Bill.Row).从属父号 = 0) Or Not gbln从项汇总折扣 Then  '(如果有级联,只取一级)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem
                                   mlngPreRow = 0 '通过行变化标志来重新确定列性质
                                End If
                            End If
                        End If
                        
                    End With
                End If
                
                '只输入一次付数
                If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Details(Bill.Row).收费类别 = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Details(i).收费类别 = "7" Then
                                '正常执行该过程：本身会定位下一个单元,先定位到付数,则下一个单元是数次
                                '选择调用该过程：调用后会送个回车，这里不能再回车，否则是三个回车的效果(控件原因)。
                                Bill.Col = BillCol.付数: Exit For
                            End If
                        Next
                    End If
                End If
            Case "付数"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "付数应该为正的整数！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                
                    '仅中药及非从属项目才可更改付数(主项付数改变,从属也变)
                    If mobjBill.Details(Bill.Row).收费类别 = "7" Then 'And mobjBill.Details(Bill.Row).从属父号 = 0 Then
                        '分批或时价药品不足禁止输入(没有分批的时价药品可以修改付数、数次)
                        If mobjBill.Details(Bill.Row).Detail.分批 Or mobjBill.Details(Bill.Row).Detail.变价 Then
                            If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 > mobjBill.Details(Bill.Row).Detail.库存 Then
                                MsgBox """" & mobjBill.Details(Bill.Row).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                            End If
                        End If
                              
                        '检查其它时价或分批中药更改付数后库存是否足够
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" _
                                And (mobjBill.Details(i).Detail.变价 Or mobjBill.Details(i).Detail.分批) Then
                                If Val(Bill.Text) * mobjBill.Details(i).数次 > mobjBill.Details(i).Detail.库存 Then
                                    MsgBox "第 " & i & " 行药品""" & mobjBill.Details(Bill.Row).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                        
                        '计算并刷新该行
                        lngOld付数 = mobjBill.Details(Bill.Row).付数
                        mobjBill.Details(Bill.Row).付数 = Bill.Text
                        Call CalcMoneys(Bill.Row)
                                                
                        If mint记录性质 = 2 Then
                            If Not IsNull(mrsInfo!险类) And mobjBill.Details(Bill.Row).数次 <> 0 And MCPAR.实时监控 Then
                                If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text), Bill.Row)) = False Then
                                    mobjBill.Details(Bill.Row).付数 = lngOld付数
                                    Bill.Text = lngOld付数
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                                               
                         '处理其它中药付数,如果是独立项,则修改其它非从项的,如果是从项,则修改同一主项的从项的.因为限定为中草药,不可能有主项
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" And mobjBill.Details(i).从属父号 = mobjBill.Details(Bill.Row).从属父号 Then
                                If mobjBill.Details(i).从属父号 = 0 Or (mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).Detail.固有从属 = 0) Then     '1和2固定和按比例的不改
                                    mobjBill.Details(i).付数 = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                                                
                        Call ShowMoney
                    Else
                        sta.Panels(2) = "从属项目的付数不能更改！"
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Beep '恢复原有付数值
                    End If
                End If
            Case "数次"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                        End If
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '负数合法性检查
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 < 0 Then
                        '权限
                        If Not ((InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 And InStr(mstrPrivs, "药品负数费用") > 0) _
                             Or (InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) = 0 And InStr(mstrPrivs, "诊疗负数费用") > 0)) Then
                            MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                        Else
                            If mobjBill.Details(Bill.Row).Detail.分批 Then
                                MsgBox "分批药品不允许输入负数。", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                            End If
                            If mrsInfo.State = 1 And mint记录性质 = 2 Then
                                If Not IsNull(mrsInfo!险类) Then
                                    If Not MCPAR.负数记帐 Then
                                        MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                                        Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '药品库存检查
                    With mobjBill.Details(Bill.Row)
                        If (.收费类别 = "4" And .Detail.跟踪在用) Or InStr(",5,6,7,", .收费类别) > 0 Then
                            If .Detail.分批 Or .Detail.变价 Then
                                '分批或时价药品不足禁止输入
                                If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                    If .收费类别 = "4" Then
                                        MsgBox """" & .Detail.名称 & """为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIF(.收费类别 = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .执行部门ID) <> 0 And Bill.ColData(BillCol.执行科室) = 5 Then
                                    '其它药品正常检查
                                    If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                        If colStock("_" & .执行部门ID) = 1 Then
                                            If MsgBox("""" & .Detail.名称 & """的当前可用库存不足输入数量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .数次: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .执行部门ID) = 2 Then
                                            MsgBox """" & .Detail.名称 & """的当前可用库存不足输入数量！", vbInformation, gstrSysName
                                            Bill.Text = .数次: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                        dblPreTime = .数次
                        .数次 = Bill.Text
                        
                        '处方限量检查
                        If Not CheckLimit(mobjBill, Bill.Row, mbln药房单位) Then
                            .数次 = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                        '单笔最大录入限量
                        If .Detail.录入限量 > 0 And .数次 * .付数 * IIF(mbln药房单位, .Detail.药房包装, 1) > .Detail.录入限量 Then
                            If MsgBox("输入的数次超过了录入限量" & .Detail.录入限量 & ",是否继续?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                                .数次 = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        '审批使用限量
                        If mint病人来源 = 2 And mint记录性质 = 2 And mrsInfo.State = 1 Then
                            If .Detail.要求审批 And Not IsNull(mrsInfo!险类) And Not mrsMedAudit Is Nothing Then
                                mrsMedAudit.Filter = "项目ID=" & .收费细目ID
                                If mrsMedAudit.RecordCount > 0 Then
                                    If Not IsNull(mrsMedAudit!可用数量) Then
                                        If .数次 * .付数 * IIF(mbln药房单位, .Detail.药房包装, 1) > mrsMedAudit!可用数量 Then
                                            MsgBox "输入的数次超过了批准的可用数量" & FormatEx(mrsMedAudit!可用数量 / IIF(mbln药房单位, .Detail.药房包装, 1), 5) & "。", vbInformation, gstrSysName
                                            .数次 = dblPreTime: Bill.Text = dblPreTime
                                            Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '固有从属不能更改数次(主项目数次改变,固有从属的数次也变)
                        If .从属父号 <> 0 And .Detail.固有从属 <> 0 Then
                            sta.Panels(2) = "该项目是固有从属项目,其数次不能够更改。"
                            .数次 = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                    End With
                
                    Call CalcMoneys(Bill.Row)
                    
                    '数据溢出检查(在已经算出该行费用但未显示前)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                        mobjBill.Details(Bill.Row).数次 = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mint记录性质 = 2 And mrsWarn.State = 1 Then
                        curTotal = GetBillTotal(mobjBill)
                        
                        '刘兴洪:2010-07-01 10:23:11:30504:and mbln费用登记=False
                        If curTotal > 0 And mbln费用登记 = False Then
                            cur余额 = Val(txt实收.Tag)
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details(Bill.Row).数次 = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                                     
                    If mint记录性质 = 2 Then
                        If Not IsNull(mrsInfo!险类) And mobjBill.Details(Bill.Row).数次 <> 0 And MCPAR.实时监控 Then
                            If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text), Bill.Row)) = False Then
                                mobjBill.Details(Bill.Row).数次 = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '更改其固有从属的数次
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).从属父号 = Bill.Row Then
                            '28136
                            '如果是输入的负数,需要将下级中的负数集中更新成负数
                            With mobjBill.Details(i)
                                If .Detail.固有从属 = 0 Then  '非固有从属
                                    If Abs(.数次) <> Abs(.Detail.从项数次) Then GoTo NotCalc:
                                    .数次 = IIF(Val(Bill.Text) < 0, -1, 1) * .Detail.从项数次
                                ElseIf .Detail.固有从属 = 1 Then '固定的固有从属
                                    .数次 = IIF(Val(Bill.Text) < 0, -1, 1) * IIF(.Detail.从项数次 = 0, 1, .Detail.从项数次)
                                ElseIf .Detail.固有从属 = 2 Then   '按比例的固有从属
                                    .数次 = Val(Bill.Text) * .Detail.从项数次
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
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
               End If
                If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
            Case "单价"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "项目价格不应该为负数，要删除费用，请输入负的数量来实现！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 * mobjBill.Details(Bill.Row).数次 > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If

                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '如果没有对应的收入项目,则无法计算
                    If mobjBill.Details(Bill.Row).Detail.变价 And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                        If Not (mobjBill.Details(Bill.Row).InComes(1).现价 = 0 And mobjBill.Details(Bill.Row).InComes(1).原价 = 0) Then
                            strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).原价, mobjBill.Details(Bill.Row).InComes(1).现价, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(2) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).标准单价
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        
                        dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).标准单价
                        
                        mobjBill.Details(Bill.Row).InComes(1).标准单价 = Bill.Text '这种收费细目只能对应一个收入项目
                        Call CalcMoneys(Bill.Row)
                        
                        '记帐分类报警(在已经算出该行费用但未显示前)
                        If mint记录性质 = 2 And mrsWarn.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            '30504:and mbln费用登记=False
                            If curTotal > 0 And mbln费用登记 = False Then
                                cur余额 = Val(txt实收.Tag)
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                    Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn)
                                If mbytWarn = 2 Or mbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).标准单价 = dblPreMoney
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
                        sta.Panels(2) = "该项目设有设置对应的费目，所以无法计算费用！"
                        Beep
                    End If
                End If
            Case "执行科室"
                If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Details(Bill.Row)
                            If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                                .执行部门ID = Bill.ItemData(Bill.ListIndex)
                                If CheckItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                            End If
                    
                            '药品库存检查:动态药房,分批或时价药品也要检查了
                            If (.收费类别 = "4" And .Detail.跟踪在用) Or InStr(",5,6,7,", .收费类别) > 0 Then
                                If .Detail.分批 Or .Detail.变价 Then '分批或时价药品库存不足禁止输入
                                    If .付数 * .数次 > .Detail.库存 Then
                                        If .收费类别 = "4" Then
                                            MsgBox "[" & .Detail.名称 & "]为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                        Else
                                            MsgBox "[" & .Detail.名称 & "]为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                        End If
                                        Cancel = True
                                    End If
                                Else
                                    Set colStock = IIF(.收费类别 = "4", mcolStock2, mcolStock1)
                                    If colStock("_" & .执行部门ID) <> 0 Then
                                        If .付数 * .数次 > .Detail.库存 Then
                                            If colStock("_" & .执行部门ID) = 1 Then
                                                If MsgBox("[" & .Detail.名称 & "]的当前可用库存不足输入数量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                    Cancel = True
                                                End If
                                            ElseIf colStock("_" & .执行部门ID) = 2 Then
                                                MsgBox "[" & .Detail.名称 & "]的当前可用库存不足输入数量！", vbInformation, gstrSysName
                                                Cancel = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                            '检查卫生材料的灭菌效期,在确定执行科室之后
                            If .收费类别 = "4" And .Detail.跟踪在用 Then
                                Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                            End If
                        
                            If CheckItemHaveSub(Bill.Row) Then
                                KeyCode = 0
                                Call LocateMainItemNextRow(Bill.Row)
                            End If
                    End With
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 = lngRow Then
            If mobjBill.Details(i).Detail.固有从属 = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.数次
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
        Bill.Col = BillCol.类别
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.类别
    End If
    '问题:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetSubItem()
'功能:输入收费项目后,加载当前收费项目的从属项目到费用集对象,并显示在单据控件中
'参数:
'调用者:Bill_KeyDown中输入项目后
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long               '执行科室ID
Dim bln从项汇总折扣 As Boolean
Dim str摘要 As String

lngMainRow = Bill.Row               '主项的行
If gbln从项汇总折扣 Then            '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
    bln从项汇总折扣 = Not mobjBill.Details(lngMainRow).Detail.屏蔽费别
End If

With mobjBill.Details(lngMainRow)
    Set mcolDetails = New Details
    Set mcolDetails = GetSubDetails(.收费细目ID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)
            mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = "" '有必要加上
        
        'a.从属项目为非药品项目的执行科室
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).类别) = 0 Then
             If mcolDetails(i).类别 = .收费类别 Then
                '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                lngDoUnit = .执行部门ID
             Else
                If mcolDetails(i).执行科室 = 0 Then
                    '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                    lngDoUnit = .执行部门ID
                Else
                    '其它非药项目的执行科室
                    lngDoUnit = Get收费执行科室ID(mlng病人ID, mlng主页ID, mcolDetails(i).类别, _
                        mcolDetails(i).ID, mcolDetails(i).执行科室, lngDoUnit, Get开单科室ID, mint病人来源, , 1, 1)
                End If
             End If
        End If
        
        'b.从属项目为药品,卫材的执行科室(如果主项的执行科室为空,也会执行到这里)
        If lngDoUnit = 0 Then
            lngDoUnit = mobjBill.科室ID
            If lngDoUnit = 0 And cbo开单科室.ListIndex <> -1 Then
                lngDoUnit = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            End If
            lngDoUnit = Get收费执行科室ID(mlng病人ID, mlng主页ID, mcolDetails(i).类别, mcolDetails(i).ID, _
                mcolDetails(i).执行科室, lngDoUnit, Get开单科室ID, mint病人来源, .执行部门ID, 1, 1) '卫材从项缺省与主项执行科室相同
        End If
            
        '保险支付项目对应检查
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!险类) Then
                If zlCheck定价零价格对码(mcolDetails(i).ID, Not mcolDetails(i).变价) Then
                    '问题:27286
                Else
                    If Not ItemExistInsure(mrsInfo!病人ID, mcolDetails(i).ID, mrsInfo!险类) Then
                        If gint医保对码 = 1 Then
                            If MsgBox("项目""" & mcolDetails(i).名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        ElseIf gint医保对码 = 2 Then
                            MsgBox "项目""" & mcolDetails(i).名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        Call SetDetailtStock(lngDoUnit, mcolDetails(i))
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln从项汇总折扣)
        Call ShowDetails(Bill.Rows - 1)
        
        If mrsInfo.State = 1 And mint病人来源 = 2 Then
                'mint病人来源 = 2:41136
                'CalcMoney中先调用GetuItemInsure可能返回摘要
                str摘要 = mobjBill.Details(Bill.Rows - 1).摘要
                str摘要 = gclsInsure.GetItemInfo(Val(Nvl(mrsInfo!险类)), mrsInfo!病人ID, mcolDetails(i).ID, str摘要, 2)
                mobjBill.Details(Bill.Rows - 1).摘要 = str摘要
        End If
    Next
    
    If bln从项汇总折扣 And Not mbln费用登记 Then
        Call CalcMoney(lngMainRow, bln从项汇总折扣) '先重算主项的应收与实收,因为在没有加入从项前可能是按单独打折算的.
        
        Call Calc重算主项实收(lngMainRow)
    End If
    
    Call ShowMoney
End With

End Sub

Private Function Get开单科室ID() As Long
    If cbo开单科室.ListIndex <> -1 Then
        Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        Get开单科室ID = UserInfo.部门ID
    End If
End Function

Private Sub Calc重算主项实收(ByVal lngMainRow As Long)
'功能:当从项汇总折扣时,根据指定的主项的行ID的第一个收入项目重算主项的实收金额
'参数:  lngMainRow-主项行ID

Dim i As Long, j As Long
Dim cur打折前应收合计 As Currency     '记录所有主从项的应收合计
Dim cur打折后实收 As Currency


With mobjBill
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).从属父号 <> lngMainRow Then Exit For    '虽然目前限制了不允许在从项中间插入别的主从项,但因一张单据行数不多,为了将来可能的需求,还是全部扫描
        
        If i = lngMainRow Or .Details(i).从属父号 = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur打折前应收合计 = cur打折前应收合计 + .Details(i).InComes(j).应收金额
            Next
        End If
    Next
   
    cur打折后实收 = CCur(Format(ActualMoney(.费别, .Details(lngMainRow).InComes(1).收入项目ID, cur打折前应收合计), gstrDec))
    
    cur打折后实收 = cur打折后实收 - cur打折前应收合计 + .Details(lngMainRow).InComes(1).应收金额
    
    .Details(lngMainRow).InComes(1).实收金额 = Format(cur打折后实收, gstrDec)
    .Details(lngMainRow).InComes(1).Key = "_" & Format(cur打折后实收, gstrDec)
    
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubDept(ByVal lngRow As Long)
Dim i As Long, j As Long
    With mobjBill
        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID) '必须现取
        
        For i = lngRow + 1 To .Details.Count
            If .Details(i).从属父号 = lngRow Then
                '从属项为药品和卫材的项目的执行科室不随主项变动
                If InStr(",4,5,6,7,", .Details(i).收费类别) = 0 Then
                    If .Details(i).收费类别 = .Details(lngRow).收费类别 Then
                        '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                        .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).执行科室 = 0 Then
                                '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                                 .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                            Else
                                '3.其它非药项目的执行科室
                                If cbo开单科室.ListIndex <> -1 Then
                                    .Details(i).执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                                End If
                                .Details(i).执行部门ID = Get收费执行科室ID(mlng病人ID, mlng主页ID, mcolDetails(j).类别, _
                                    mcolDetails(j).ID, mcolDetails(j).执行科室, .Details(i).执行部门ID, Get开单科室ID, mint病人来源, , 1, 1)
                            End If
                        End If
                    End If
                    
                    '显示从项执行科室
                    If .Details(i).执行部门ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                            Else
                                Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.Details(i).执行部门ID)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.Details(i).执行部门ID)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
'功能：判断当前行的项目是否具有从属项目
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = lngRow Then
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
        '针对列编辑性质设置颜色
        Bill.SetColColor BillCol.类别, &HE7CFBA '不然要成白色
        Exit Sub
    End If
    sta.Panels(2).Text = ""
     '--------------------------------------------------------------------------
    '1.行改变的相关数据处理和设置
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '显示库存
            If InStr(",5,6,7,", .收费类别) > 0 And .收费细目ID <> 0 Then
                If mbln其它药房 Or mbln其它药库 Then
                    strStock = GetStockInfo(.收费细目ID, mbln其它药房, mbln其它药库, mbln药房单位, mstr药房包装)
                    If strStock <> "" Then
                        If InStr(1, mstrPrivs, "显示库存") > 0 Then
                            sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行库存:" & strStock
                        Else
                            sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行有库存."
                        End If
                    End If
                End If
                If strStock = "" Then
                    '随时更新库存显示
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                    If mbln药房单位 Then
                        .Detail.库存 = .Detail.库存 / .Detail.药房包装
                    End If
                    Call ShowStock(.Detail.名称, .Detail.库存)
                End If
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                Call ShowStock(.Detail.名称, .Detail.库存)
            ElseIf .Detail.变价 And .InComes.Count > 0 And Bill.TextMatrix(0, Bill.Col) = "单价" Then
                sta.Panels(2) = "价格范围:" & FormatEx(.InComes(1).原价, 5) & "-" & FormatEx(.InComes(1).现价, 5)
            Else
                sta.Panels(2) = ""
            End If
            
            Bill.ColData(BillCol.类别) = IIF(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.项目) = BillColType.CommandButton
            
             '如果是从属项目的主项目或从项,则不允许更改类别和项目
            If CheckItemHaveSub(Row) Or .从属父号 > 0 Then
                Bill.ColData(BillCol.类别) = BillColType.Text_UnModify
                Bill.ColData(BillCol.项目) = BillColType.Text_UnModify
            End If
            
            '如果是非调整状态
            If mbytInState <> 2 Then
                If .收费类别 = "7" And mblnPay Then
                    Bill.ColData(BillCol.付数) = 4
                Else
                    Bill.ColData(BillCol.付数) = 5
                End If
                
                '变价允许输入数次
                If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                    And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                    Bill.ColData(BillCol.数次) = IIF(mblnTime, 4, 5) '数次
                    Bill.ColData(BillCol.单价) = 4 '金额
                Else
                    Bill.ColData(BillCol.数次) = 4
                    Bill.ColData(BillCol.单价) = 5
                End If
                
                If .Key = "1" Then    '指定了固定药房时,不允许再选择执行科室
                    Bill.ColData(BillCol.执行科室) = BillColType.UnFocus
                Else
                    Bill.ColData(BillCol.执行科室) = BillColType.ComboBox
                End If
                
                If .收费类别 = "F" Then
                    Bill.ColData(BillCol.标志) = -1
                Else
                    Bill.ColData(BillCol.标志) = 5
                End If
                
                 '只允许一个类别
                If mblnOne Then Bill.ColData(BillCol.类别) = 5
            End If
        End With
    End If
   
    '如果点击未保存的行,则恢复列的性质
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.类别) = IIF(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus) '类别列,当主从项时会被改变
        Bill.ColData(BillCol.项目) = BillColType.CommandButton  '项目列,当主从项时会被改变
    End If
    
    
    '-----------------------------------------------------------------
    '2.列改变的相关数据处理和显示设置
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then
        Call FillBillComboBox(Bill.Row, Bill.Col, True) '进入该列
    End If
    
    If gbln收费类别 And Bill.TextMatrix(Row, BillCol.类别) = "" And mblnOne Then
        mrsClass.Filter = "编码=" & mstr收费类别
        Bill.TextMatrix(Row, BillCol.类别) = mrsClass!类别
        Bill.RowData(Row) = Asc(mrsClass!编码)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "类别" '不输入收费类别时不会进入类别列
            Call zlControl.CboSetWidth(Bill.CboHwnd, 1000)
            '类别如果为空,则自动默认为上一收费细目的类别
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "编码=" & mstr收费类别
                    Bill.TextMatrix(Row, Col) = mrsClass!类别
                    Bill.RowData(Row) = Asc(mrsClass!编码)
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
        Case "执行科室"
            Call zlControl.CboSetWidth(Bill.CboHwnd, 2000)
        Case "付数"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "数次"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789" & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                '可否输入小数
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                    If InStr(mstrPrivs, "药品小数输入") > 0 Then
                        Bill.TextMask = "." & Bill.TextMask
                    End If
                Else
                    Bill.TextMask = "." & Bill.TextMask
                End If
                
                '可否输入负数
                If Not mobjBill.Details(Bill.Row).Detail.分批 Then
                    If InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                        If InStr(mstrPrivs, "药品负数费用") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivs, "诊疗负数费用") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                                    
                    If InStr(Bill.TextMask, "-") > 0 Then
                        If mrsInfo.State = 1 And mint记录性质 = 2 Then
                            If Not IsNull(mrsInfo!险类) Then
                                If Not MCPAR.负数记帐 Then
                                    Bill.TextMask = Replace(Bill.TextMask, "-", "")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case "单价"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
   
    
    '显示摘要
    If mobjBill.Details.Count >= Bill.Row Then
        If mobjBill.Details(Bill.Row).摘要 <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  摘要:" & mobjBill.Details(Bill.Row).摘要
        End If
    End If
    
    '新行,或更改已有行的类别时,看作换行还没有开始
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'bill.ToolTipText = bill.TextMatrix(bill.MouseRow, bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, strDoctor As String
    Dim rsReturn As ADODB.Recordset
    Dim intIndex As Integer
    If mbytInState <> 0 Then Exit Sub
    mrsAll开单科室.Filter = ""
    If cbo开单科室.ItemData(cbo开单科室.ListIndex) = 0 And cbo开单科室.Text Like "其他科室*" Then
        If zlDatabase.zlShowListSelect(Me, glngSys, 1150, cbo开单科室, mrsAll开单科室, True, "", "缺省,优先级", rsReturn) = False Then
            mobjBill.开单部门ID = 0
            Exit Sub
        End If
        If rsReturn Is Nothing Then Exit Sub
        If rsReturn.State <> 1 Then Exit Sub
        If rsReturn.RecordCount = 0 Then Exit Sub
        rsReturn.MoveFirst
        If zlControl.CboLocate(cbo开单科室, Val(rsReturn!ID), True) = False Then
            cbo开单科室.RemoveItem cbo开单科室.ListCount - 1
            cbo开单科室.AddItem IIF(zlIsShowDeptCode, rsReturn!编码 & "-", "") & rsReturn!名称
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = Val(Nvl(rsReturn!ID))
            intIndex = cbo开单科室.NewIndex
            cbo开单科室.AddItem "其他科室…"
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = 0
            cbo开单科室.ListIndex = intIndex
        End If
        Exit Sub
    End If
    '定位医生
    cbo开单人.Clear
    If cbo开单科室.ListIndex <> -1 Then
        Call Load开单人(cbo开单科室.ItemData(cbo开单科室.ListIndex))
    End If
    
    '数据对象
    If mbytInState = 0 Then
        If cbo开单科室.ListIndex = -1 Then
            mobjBill.开单部门ID = 0
        Else
            mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
    End If
    
    
    '重新设置相关项目的执行科室
    If mbytInState = 0 And cbo开单科室.ListIndex <> -1 And cbo开单科室.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '仅处理收费项目
                If InStr(",4,5,6,7,", .Detail.类别) = 0 And .Detail.执行科室 = 6 Then '6-开单人科室
                    .执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    '刷新显示从项执行科室
                    If i <= Bill.Rows - 1 And .执行部门ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                            Else
                                Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.执行部门ID, mrsUnit)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.执行部门ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                    '撤销8113的修改
'                ElseIf InStr(",4,5,6,7,", .Detail.类别) > 0 Then
'                '重设可用药房为存储库房中设置的服务于病人科室(开单科室)的执行科室
'                    If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
'                        Bill.ColData(BillCol.执行科室) = BillColType.ComboBox
'                    End If
'                    If .Key = "1" Then .Key = "0"        '1表示执行科室不可选择
'                    mlngPreRow = 0      '用于在Entercell事件中再重视执行科室的可选性
                End If
            End With
        Next
    End If
End Sub

Private Sub cbo开单人_Click()
    Dim arrDepts As Variant, i As Long, k As Long
    
    If mbytInState = 0 Then
        mobjBill.开单人 = IIF(cbo开单人.ListIndex = -1, "", NeedName(cbo开单人.Text))
                        
        '护士类别
        If Bill.Active Then
            If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
                And Bill.RowData(Bill.Rows - 1) <> 0 Then
                '清除无效输入
                Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = ""
                Bill.RowData(Bill.Rows - 1) = 0
            ElseIf Bill.Col = BillCol.类别 Then
                Call Bill_EnterCell(Bill.Row, Bill.Col) '刷新
            End If
        End If
        
        '护士类别:判断非法输入
'        If HaveStopClass > 0 Then
'            MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
'        End If
    End If
End Sub

Private Sub cbo开单人_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo开单人.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo开单人.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo开单人_Validate(Cancel As Boolean)
    If cbo开单人.Text <> "" Then
        Call GetCboIndex(cbo开单人, NeedName(cbo开单人.Text))
        If cbo开单人.ListIndex = -1 Then cbo开单人.Text = ""
    End If
    If cbo开单人.Text = "" Then Call cbo开单人_KeyPress(vbKeyReturn)
End Sub

Private Sub chk急诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk加班_Click()
    If mbytInState = 1 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk加班.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    blnAdd = OverTime
    If chk加班.value = Unchecked And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.value = Checked
        End If
    End If
    If chk加班.value = Checked And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.value = Unchecked
        End If
    End If
    mobjBill.加班标志 = IIF(chk加班.value = Checked, 1, 0)
    
    '重新计算价格
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count > 0 Or gblnOK Then
        If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Function CheckNegative() As Boolean
'功能：检查单据中输入的负数数量及退回科室是否正确
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim strItems As String, str部门 As String
    Dim str单位 As String, dbl数量 As Double
    Dim strValues(0 To 10) As String, intR As Long
    Dim strSubTable As String, dbl数次合计 As Double, dbl已结数量 As Double
    
    '问题:26951
    If InStr(1, mstrPrivs, ";负数记帐不检查发生项目;") > 0 Then
        '对于负数冲销时不检查本次住院发生的项目数量,有此权限,允许录入病人未曾发生的费用项目进行冲销,否则检查本次住院发生的项目数量才能冲销
        CheckNegative = True: Exit Function
    End If
    
    CheckNegative = True
    If mobjBill.病人ID = 0 Then Exit Function
    
    strItems = ""
    strSubTable = ""
    intR = 0
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And .执行部门ID <> 0 Then
                If Len(strItems) > 2000 Then
                    If intR <= 10 Then
                        strValues(intR) = Mid(strItems, 2)
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As 收费细目ID,  " & _
                        "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As 执行部门ID,0 as 数量,0 as 结帐数量 " & _
                        " From Table(Cast(f_str2list([" & intR + 3 & "]) As ZLTOOLS.t_strlist))"
                    Else
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As 收费细目ID,  " & _
                        "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As 执行部门ID,0 as 数量,0 as 结帐数量 " & _
                        " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
                    End If
                    strItems = "": intR = intR + 1
                End If
                strItems = strItems & "," & .收费细目ID & "_" & .执行部门ID & ""
'                strSQL = strSQL & " Union ALL Select " & .收费细目ID & "," & .执行部门ID & ",0 From Dual"
            End If
        End With
    Next
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
            " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As 收费细目ID,  " & _
            "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As 执行部门ID,0 as 数量,0 as 结帐数量 " & _
            " From Table(Cast(f_str2list([" & intR + 3 & "]) As ZLTOOLS.t_strlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
            " Select to_number(substr(Column_Value,1,instr(Column_Value,'_')-1)) As 收费细目ID,  " & _
            "           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As 执行部门ID,0 as 数量,0 as 结帐数量 " & _
            " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
        End If
    End If
    
    If strSubTable = "" Then Exit Function
    strSubTable = Mid(strSubTable, 11)
    
    StrSQL = " " & _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.收费细目ID,A.执行部门ID,  " & _
    "             Nvl(Sum(Decode(A.记录性质, 2, 1, 3, 1, 0) * Nvl(A.付数, 1) * A.数次), 0) As 数量, " & _
     "            Sum(Decode(nvL(Mod(M.记录状态 , 3),1),  0, 1, 1, 1, -1) * Decode(A.结帐id, Null, 0, 1) * Nvl(付数, 1) * 数次) As 结帐数量 " & _
     "     From " & mstrFeeTab & " A, 病人结帐记录 M " & _
     "     Where  A.结帐id = M.ID(+)  And A.记帐费用=1 And A.价格父号 Is Null  And A.记录状态<>0 " & _
     "             And A.病人ID=[1] " & IIF(mint病人来源 = 2, "  And Nvl(A.主页ID,0)=[2]", "") & _
     "             And (A.收费细目ID+0,执行部门ID,0,0) in (select * From C1) " & _
     "     Group By A.收费细目ID,A.执行部门ID" & _
     "     Union ALL Select * From C1 "
   ' strSQL = _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.收费细目ID,A.执行部门ID,Sum(Nvl(A.付数,1)*A.数次) as 数量," & _
    "           Sum(decode(A.结帐ID,NULL,0,1)* Nvl(A.付数,1)*A.数次) as 结帐数量 " & _
    " From  " & mstrFeeTab & " A " & _
    " Where A.记录状态<>0 And A.记帐费用=1 And A.价格父号 is NULL" & _
    "           And A.病人ID=[1] " & IIF(mint病人来源 = 2, "  And Nvl(A.主页ID,0)=[2]", "") & _
    "           And (A.收费细目ID+0,执行部门ID,0,0) in (select * From C1) " & _
    " Group by A.收费细目ID,A.执行部门ID" & _
    " Union ALL Select * From C1"
    
    StrSQL = "" & _
    "   Select 收费细目ID,执行部门ID,Sum(数量) as 数量,sum(结帐数量) as 结帐数量 " & _
    "   From (" & StrSQL & ") " & _
    "   Group by 收费细目ID,执行部门ID"
    
    On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mobjBill.病人ID, mobjBill.主页ID, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And .执行部门ID <> 0 Then
                rsTmp.Filter = "收费细目ID=" & .收费细目ID & " And 执行部门ID=" & .执行部门ID
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        str单位 = .Detail.药房单位
                        dbl数量 = Nvl(rsTmp!数量, 0) / .Detail.药房包装
                        dbl数次合计 = Abs(.数次) * .付数
                        dbl已结数量 = Val(Nvl(rsTmp!结帐数量)) / .Detail.药房包装
                    Else
                        str单位 = .Detail.计算单位
                        dbl数量 = Nvl(rsTmp!数量, 0)
                        dbl数次合计 = Abs(.数次) * .付数
                        dbl已结数量 = Val(Nvl(rsTmp!结帐数量))
                        '可能存在两条相同的记录
                        '问题:29412
                        For j = i + 1 To mobjBill.Details.Count
                             If .收费细目ID = mobjBill.Details(j).收费细目ID _
                                And mobjBill.Details(j).数次 < 0 And mobjBill.Details(j).执行部门ID = .执行部门ID Then
                                dbl数次合计 = dbl数次合计 + Abs(.数次) * .付数
                             End If
                        Next
                    End If
                    '问题:32106
                    If dbl数次合计 > dbl数量 - dbl已结数量 Then
                        Select Case gbytBillOpt '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
                        Case 0  '允许
                            If dbl数次合计 > dbl数量 Then
                                str部门 = Get部门名称(.执行部门ID)
                                MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                    " 多于已计费数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                            End If
                        Case 1   '提醒
                            str部门 = Get部门名称(.执行部门ID)
                            If dbl数次合计 > dbl数量 Then
                                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                        " 多于已计费数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                    CheckNegative = False: Exit Function
                            End If
                            
                            If MsgBox("第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                " 中包含了已结部分(未结:" & FormatEx(dbl数量 - dbl已结数量, 5) & str单位 & "; 已结:" & FormatEx(dbl已结数量, 5) & str单位 & ") 。" & vbCrLf & _
                                " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                CheckNegative = False: Exit Function
                            End If
                        Case 2   '禁止
                            str部门 = Get部门名称(.执行部门ID)
                            MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                " 多于已计费数量 " & FormatEx(dbl数量 - dbl已结数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                        End Select
                    End If
                Else
                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]可销帐数量为零，不允许冲销。", vbInformation, gstrSysName
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, StrSQL As String, strTmp As String
    Dim i As Long, j As Long, lng结帐ID As Long
    Dim curTotal As Currency, intInsure As Integer
    Dim dblTotal As Double, cur余额 As Currency, dbl数次 As Double
    Dim cur当日额 As Currency, colStock As Collection
    Dim blnTrans As Boolean, strNos As String
    
    If mbytInState = 3 Then
        If mint记录性质 <> 1 And (False Or mlng医嘱ID <> 0) Then '划价是全部删除
            For i = 1 To Bill.Rows - 1
                'If Bill.TextMatrix(i, Bill.Cols - 1) = "√" And Bill.RowData(i) > 0 Then
                If Bill.RowData(i) > 0 Then
                    StrSQL = StrSQL & "," & Bill.RowData(i)
                End If
            Next
            If StrSQL = "" Then
                MsgBox "请至少选择一行要删除的费用！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
            
            '所有行选择处理
            StrSQL = Mid(StrSQL, 2)
            i = GetBillRows(mstrInNO, mint记录性质, mint病人来源)
            If UBound(Split(StrSQL, ",")) + 1 = i Then StrSQL = ""
        Else
            '因为要处理为全退，如果结帐后不允许销帐，部份结帐后就要检查
            j = 0
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then j = j + 1
            Next
            i = GetBillRows(mstrInNO, mint记录性质, mint病人来源)
            If j < i Then
                MsgBox "单据中的部份项目当前已不允许销帐(比如已执行或已结帐的项目)。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '医保记帐作废上传(注意判断顺序)
        If mint病人来源 = 2 Then
            intInsure = BillExistInsure(mstrInNO) '判断是否医保病人记的帐
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) Then
                    '去掉了医保连接匹配检查
                    If StrSQL <> "" Then '不能部分销帐
                        MsgBox "因为医保处理需要,该单据中的项目必须全部删除！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If mint病人来源 = 2 Then
            StrSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "','" & StrSQL & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            If mint记录性质 = 2 Then
                StrSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "','" & StrSQL & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Else
                StrSQL = "zl_门诊划价记录_DELETE('" & mstrInNO & "')"
            End If
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        
            Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
                        
            '医保记帐作废上传
            If mint病人来源 = 2 And intInsure > 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保记帐作废上传
        If mint病人来源 = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的删费数据向医保传送失败，该单据已删除。", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0
        
        gblnOK = True: Unload Me: Exit Sub
    Else '正常输入单据状态
        If mobjBill.病人ID = 0 Or mrsInfo.State = 0 Then
            MsgBox "没有发现病人信息，单据不能保存。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mobjBill.Details.Count = 0 Then
            MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Check执行科室
        If i <> 0 Then
            MsgBox "单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.开单部门ID = 0 Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            cbo开单科室.SetFocus: Exit Sub
        End If
        
        If mobjBill.开单人 = "" Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        '护士类别:判断非法输入
'        If HaveStopClass > 0 Then
'            MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
'            Exit Sub
'        End If
                
        '发生时间检查
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '出院强制记帐权限检查
        If mint病人来源 = 2 Then
            If Not PatiCanBilling(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), mstrPrivs, p医嘱附费管理) Then Exit Sub
            If zlPatiIS病案已编目(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0)) Then Exit Sub
            '49501:住院
            If zlIsAllowFeeChange(mrsInfo!病人ID, Val(Nvl(mrsInfo!主页ID))) = False Then Exit Sub
        End If
        
        '刘兴洪 问题:?? 日期:2010-01-07 10:37:09
        If zlCheck北京医保(Val(Nvl(mrsInfo!险类))) = False Then Exit Sub
        
        
        
        '发生时间检查
        If Not IsNull(mrsInfo!出院日期) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!出院日期, txtDate.Format) Then
                MsgBox "强制对出院病人记帐时，费用时间不能大于病人出院时间:" & Format(mrsInfo!出院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!入院日期) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!入院日期, txtDate.Format) Then
                MsgBox "费用的发生时间不能小于病人的入院时间:" & Format(mrsInfo!入院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '非法行
        dbl数次 = 0
        For i = 1 To mobjBill.Details.Count
            '27467,52828
            If mobjBill.Details(i).数次 <> 0 And dbl数次 = 0 Then
                dbl数次 = mobjBill.Details(i).数次
            End If
            If mobjBill.Details(i).收费细目ID = 0 Then
                MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
        Next
        '27467,52828
        If mbytInState = 0 And Round(dbl数次, 7) = 0 Then
            MsgBox "单据中至少要有一条不为零的数次,请检查！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '处方职务检查
        '问题:45605
        If zlIsCheckMedicinePayMode(txt付款方式) Then
            i = CheckDuty(, False)
            If i > 0 Then
                Bill.Row = i: Bill.MsfObj.TopRow = i
                Bill.Col = BillCol.项目: Bill.SetFocus
                Exit Sub
            End If
        End If

        '所有病人项目
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.项目: Bill.SetFocus
            Exit Sub
        End If
        
        '费用类型检查
        If Not Check费用类型 Then Exit Sub
                
        '记帐分类报警
        If mint记录性质 = 2 And mrsWarn.State = 1 And mstrWarn <> "-" Then
            '单据费用
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                '刷新病人费用状况
                Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, IIF(mint病人来源 = 1, 0, mlng主页ID), mcurModiMoney)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!预交余额
                    cmdCancel.Tag = rsTmp!费用余额
                    txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
                End If
                sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
                sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + curTotal, gstrDec)
                sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - curTotal, "0.00")
                
                '重新读取当日额
                cur当日额 = GetPatiDayMoney(mrsInfo!病人ID)
                
                '是否医保病人
                cur余额 = Val(txt实收.Tag)
                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                
                If mbln费用登记 = False Then    '30504
                    For i = 1 To mobjBill.Details.Count
                        mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, cur当日额 - mcurModiMoney, curTotal, IIF(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Details(i).收费类别, mobjBill.Details(i).Detail.类别名称, mstrWarn, mintWarn)
                        If mbytWarn = 2 Or mbytWarn = 3 Then Exit Sub
                    Next
                End If
            End If
        End If
        
        If mint记录性质 = 2 And Not IsNull(mrsInfo!险类) And MCPAR.实时监控 Then
            If gclsInsure.CheckItem(mrsInfo!险类, 1, 2, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text))) = False Then
                Exit Sub
            End If
        End If
        
        
        '药品禁忌检查
        strInfo = CheckDisable(mobjBill)
        If strInfo <> "" Then
            If strInfo Like "*(互相禁用)*" Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            Else
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
                    
        '处方限量检查
        If Not CheckLimit(mobjBill, , mbln药房单位) Then Exit Sub
        
        '检查分批或时价药品同一药房是否有重复输入
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.分批 Or .Detail.变价) _
                    And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .收费细目ID = mobjBill.Details(j).收费细目ID And .执行部门ID = mobjBill.Details(j).执行部门ID Then
                            If .收费类别 = "4" Then
                                MsgBox "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入，请合并！", vbInformation, gstrSysName
                            Else
                                MsgBox "第 " & j & " 行的分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入，请合并！", vbInformation, gstrSysName
                            End If
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
        
        '药品库存检查(仅不足禁止时或分批时价药品)
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIF(.收费类别 = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                        If mbln药房单位 Then
                            .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        End If
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存" & IIF(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                        If mbln药房单位 Then
                            .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        End If
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行药品""" & .Detail.名称 & _
                                """的当前库存" & IIF(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIF(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, 0) '29680
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIF(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Next
        
        '检查卫生材料的灭菌效期
        '检查自动发药:25490
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .收费类别 = "4" And .Detail.跟踪在用 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    If Not CheckValidity(.收费细目ID, .执行部门ID, dblTotal) Then Exit Sub
                ElseIf InStr(1, ",5,6,7,", .收费类别) > 0 Then
                    '打印发药单,仅普通记帐,且划价单除外
                    If mbytSendMateria <> 0 And mint记录性质 = 2 And mint病人来源 = 2 Then
                        '全部药品都确定了药房的才自动发药(分离发药时,没有确定药房)
                        mblnSendMateria = .执行部门ID <> 0
                    End If
                End If
            End With
        Next
        If InStr(mstrPrivs, ";药品发药;") = 0 Then mblnSendMateria = False
        
        If mblnSendMateria And mbytSendMateria = 2 Then
            If MsgBox("记帐完成后自动执行发药吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        '负数退费检查
        If mint记录性质 = 2 Then
            If Not CheckNegative Then Exit Sub
        End If
        
        '刷卡消费验卡
        If mint病人来源 = 1 And mint记录性质 = 2 And gbln消费验证 Then
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.病人ID, curTotal) Then Exit Sub
            End If
        End If
        '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
        If gobjSquareCard Is Nothing Then
            If mint病人来源 = 1 And gbln开单后立即结算 Then
                If MsgBox("注意：" & vbCrLf & "      医疗卡部件（zl9CardSquare）未创建，在您开单后将不能进行收费或记帐审核，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not SaveBill(strNos) Then Exit Sub
        
        '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
        If mint病人来源 = 1 And gbln开单后立即结算 And strNos <> "" Then
            If Not gobjSquareCard Is Nothing Then
                Call gobjSquareCard.zlSquareAffirm(Me, p医嘱附费管理, mstrPrivs, mlng病人ID, , , mint记录性质, strNos)
            End If
        End If
        
        '刘兴洪:打印发药单:25490
        If mblnSendMateria Then
            If InStr(1, mstrPrivs, ";发药清单打印;") > 0 Then
                If MsgBox("单据""" & mobjBill.NO & """发药完成，要打印发药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "单据号=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
        End If
        
        If mstrInNO <> "" Or mstrOriginalNO <> "" Then
            gblnOK = True: Unload Me: Exit Sub
        Else
            txtPreNO.Text = mobjBill.NO
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            Call SetMoneyList
            Call NewBill
            
            '重新读取病人信息
            Call GetPatient(mlng病人ID, mlng主页ID)
            Bill.SetFocus
        End If
    End If
    gblnOK = True
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str开单人 As String, ByVal str开单科室 As String, _
    Optional ByVal lngRow As Long) As ADODB.Recordset
'功能：根据单据对象内容创建一个明细记录集信息(以售价单位)
'字段：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室
'参数：intPage=指定的单据,lngRow=指定的行，不指定时包含所有单据的所有行
    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl单价 As Double, cur实收 As Currency
    Dim rsTmp As New ADODB.Recordset
    
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    '69788:李南春,2014-6-5,调整开单人字段大小，由20改为100
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "开单科室", adVarChar, 50, adFldIsNullable
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
        dbl单价 = 0: cur实收 = 0
        With objBill.Details(i)
            If lngRow = 0 Then
                rsTmp.Filter = "收费细目ID=" & .收费细目ID
                blnNew = rsTmp.RecordCount = 0
            Else
                blnNew = True
            End If
                            
            If blnNew Then
                rsTmp.AddNew
                
                rsTmp!病人ID = objBill.病人ID
                rsTmp!主页ID = objBill.主页ID
                
                rsTmp!收费类别 = .收费类别
                rsTmp!收费细目ID = .收费细目ID
                
                
                For j = 1 To .InComes.Count
                    dbl单价 = dbl单价 + .InComes(j).标准单价
                    cur实收 = cur实收 + .InComes(j).实收金额
                Next
                If InStr(",5,6,7,", .收费类别) > 0 And mbln药房单位 Then
                    '从药房单位转换为售价单位
                    rsTmp!数量 = IIF(.付数 = 0, 1, .付数) * .数次 * .Detail.药房包装
                    rsTmp!单价 = Format(dbl单价 / .Detail.药房包装, gstrDecPrice)
                Else
                    rsTmp!数量 = IIF(.付数 = 0, 1, .付数) * .数次
                    rsTmp!单价 = Format(dbl单价, gstrDecPrice)
                End If
                rsTmp!实收金额 = Format(cur实收, gstrDec)
                
                rsTmp!开单人 = str开单人
                rsTmp!开单科室 = str开单科室
            Else
                For j = 1 To .InComes.Count
                    dbl单价 = dbl单价 + .InComes(j).标准单价
                    cur实收 = cur实收 + .InComes(j).实收金额
                Next
                If InStr(",5,6,7,", .收费类别) > 0 And mbln药房单位 Then
                    '从药房单位转换为售价单位
                    rsTmp!数量 = rsTmp!数量 + IIF(.付数 = 0, 1, .付数) * .数次 * .Detail.药房包装
                    rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价 / .Detail.药房包装, gstrDecPrice)) / 2, gstrDecPrice)
                Else
                    rsTmp!数量 = rsTmp!数量 + IIF(.付数 = 0, 1, .付数) * .数次
                    rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价, gstrDecPrice)) / 2, gstrDecPrice)
                End If
                rsTmp!实收金额 = rsTmp!实收金额 + Format(cur实收, gstrDec)
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

Private Sub Form_Activate()
    If mbytInState <> 0 Then
        If cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus
        ElseIf cmdCancel.Visible And cmdCancel.Enabled Then
            cmdCancel.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',;|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill, i As Long
    gstrLike = IIF(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gbytCode = Val(zlDatabase.GetPara("简码方式"))

    mblnWarnCloseed = False
    glngFormW = 12000: glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName, mbytInState)
    
    gblnOK = False
    mblnEnterCell = True
    mintWarn = -1: mstrWarn = ""
    mstrFeeTab = IIF(mint病人来源 = 2 And mint记录性质 = 2, "住院费用记录", "门诊费用记录")
    Call InitLocPar
    
    '初始化单据数据
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
        '读取该单据的内容
        If mstrInNO <> "" Then '修改单据
            Set mobjBill = ImportBill(mint病人来源, mstrInNO, mint记录性质, mbln费用登记)
            If mobjBill.NO = "" Then
                MsgBox "不能正确读取计费单据的内容！", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                
                '针对列编辑性质设置颜色
                Bill.SetColColor BillCol.类别, &HE7CFBA
                Bill.SetColColor BillCol.项目, &HE7CFBA
                Bill.SetColColor BillCol.数次, &HE7CFBA
                Bill.SetColColor BillCol.执行科室, &HE7CFBA
                Bill.SetColColor BillCol.付数, &HE0E0E0
                Bill.SetColColor BillCol.单价, &HE0E0E0
                Bill.SetColColor BillCol.标志, &HE0E0E0
                
                cboNO.Text = mobjBill.NO
                               
                
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                If mint记录性质 = 2 Then
                    mcurModiMoney = GetBillMoney(mobjBill.NO) '在读取病人前取
                End If
                
                '新单时读取病人,看单据时根据单据显示病人信息
                Call GetPatient(mlng病人ID, mlng主页ID)
                If mrsInfo.State = 0 Then
                    If Not mblnWarnCloseed Then
                        MsgBox "不能读取病人信息，可能是你不具有对该病人计费的权限。", vbInformation, gstrSysName
                    End If
                    Unload Me: Exit Sub
                End If
                
                Call FindCboIndex(cbo开单科室, mobjBill.开单部门ID, False)
                Call GetCboIndex(cbo开单人, mobjBill.开单人)
                Call zlControl.CboLocate(cboBaby, mobjBill.婴儿费, True)
                
                If gbln从项汇总折扣 Then CalcMoneys
                Call ShowDetails
                Call ShowMoney
                
                '调整库存:修改时加上将要退回的库存
                For i = 1 To mobjBill.Details.Count
                    With mobjBill.Details(i)
                        Bill.RowData(i) = Asc(.收费类别) '特殊处理
                        If InStr(",5,6,7,", .收费类别) > 0 Then
                            .Detail.库存 = .Detail.库存 + .付数 * .数次
                        ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                            .Detail.库存 = .Detail.库存 + .付数 * .数次
                        End If
                    End With
                Next
                
                Call SetColNum
            End If
        Else
            '新单时读取病人,看单据时根据单据显示病人信息
            Call GetPatient(mlng病人ID, mlng主页ID)
            If mrsInfo.State = 0 Then
                If Not mblnWarnCloseed Then
                    MsgBox "不能读取病人信息，可能是你不具有对该病人计费的权限。", vbInformation, gstrSysName
                End If
                Unload Me: Exit Sub
            End If
            If Not IsNull(mrsInfo!险类) Then
                MCPAR.负数记帐 = gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.记帐上传 = gclsInsure.GetCapability(support记帐上传, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.实时监控 = gclsInsure.GetCapability(support实时监控, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.医生确定处方类型 = gclsInsure.GetCapability(support医生确定处方类型, mrsInfo!病人ID, mrsInfo!险类)
            End If
        End If
        
        If mstrInNO <> "" And mint记录性质 = 2 And mint病人来源 = 2 Then
            Call ReCalcInsure '重新计算统筹金额
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - fraTitle.Height - fraInfo.Height + 230
    
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
        
    cbo开单人.Left = lblDate.Left - cbo开单人.Width - 200
    lbl开单人.Left = cbo开单人.Left - lbl开单人.Width - 45
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 500
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200

    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mbytInState)
    
    mlng医嘱ID = 0
    mlng发送号 = 0
    mlng病人ID = 0
    mlng主页ID = 0
    mint病人来源 = 0
    mint记录性质 = 0
    mbln费用登记 = False
    mlng开单科室ID = 0
    mlng病人科室id = 0
    
    mlng开嘱科室ID = 0
    mstr开嘱医生 = ""
    mstrOriginalNO = ""
    
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    
    mbytInState = 0
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mstrPrivs = ""
    
    Set mrsInfo = Nothing
    Set mrsUnit = Nothing
    Set mrsClass = Nothing
    Set mrsWork = Nothing
    Set mrsMedAudit = Nothing
    
    If Not InDesign Then
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, glngOld)
    End If
End Sub


Private Sub picAppend_Resize()
    err = 0: On Error Resume Next
    With picAppend
        txt病人备注.Width = .ScaleWidth - txt病人备注.Left - 100
    End With
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Not gbln简码匹配方式切换 Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("简码方式", IIF(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIF(sta.Panels("WB").Bevel = sbrInset, 1, 0)))
        
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long
    If mbytInState = 3 Then
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        If mbytInState <> 2 Then
            .ColData(BillCol.类别) = IIF(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus) '类别列,当主从项时会被改变
            .ColData(BillCol.项目) = BillColType.CommandButton  '项目列,当主从项时会被改变
            .ColData(BillCol.付数) = 5 '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = 5 '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.标志) = 5 '标志缺省跳过,当为手术时,设为复选(-1)
        End If
        '针对列编辑性质设置颜色
        .SetColColor BillCol.类别, &HE7CFBA
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.执行科室, &HE7CFBA
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        .SetColColor BillCol.标志, &HE0E0E0
        
        .TextMatrix(Row, BillCol.行) = Row
        
        '特殊地方手动调用不执行
        If Row > 0 And .ColData(BillCol.类别) <> 5 And Me.Visible And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
     '刘兴洪 问题:27378 日期:2010-01-27 16:20:02
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If cbo开单科室.Locked Then Exit Sub
    If mrsAll开单科室 Is Nothing Then Exit Sub
    
    If zlSelectDept(Me, 1150, cbo开单科室, mrsAll开单科室, cbo开单科室.Text, True, , , True) = False Then
        mobjBill.开单部门ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
     
'    Dim lngIdx As Long
'
'    If KeyAscii = 13 And cbo开单科室.ListIndex <> -1 Then
'        KeyAscii = 0
'        Call zlCommFun.PressKey(vbKeyTab)
'    ElseIf KeyAscii >= 32 And Not cbo开单科室.Locked Then
'        lngIdx = zlControl.CboMatchIndex(cbo开单科室.Hwnd, KeyAscii)
'        If lngIdx = -1 And cbo开单科室.ListCount > 0 Then lngIdx = 0
'        cbo开单科室.ListIndex = lngIdx
'    End If
End Sub

Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        If cbo开单人.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = cbo开单人.Text
        If cbo开单人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo开单人.List(cbo开单人.ListIndex) Then
                Call zlControl.CboSetIndex(cbo开单人.Hwnd, -1)
            Else
                zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        If strText = "" Then
            cbo开单人.ListIndex = -1
        ElseIf cbo开单人.ListIndex = -1 Then
            intIdx = -1
            If IsNumeric(strText) Then
                For i = 0 To cbo开单人.ListCount - 1
                    If i > UBound(marrDr) Then Exit For
                    If CStr(Split(marrDr(i), "|")(2)) = strText Then
                        If intIdx = -1 Then cbo开单人.ListIndex = i
                        intIdx = i
                    End If
                Next
                If intIdx = -1 Then
                    For i = 0 To cbo开单人.ListCount - 1
                        If i > UBound(marrDr) Then Exit For
                        If Val(Split(marrDr(i), "|")(2)) = Val(strText) Then
                            If intIdx = -1 Then cbo开单人.ListIndex = i
                            intIdx = i
                        End If
                    Next
                End If
            Else
                For i = 0 To cbo开单人.ListCount - 1
                    If UCase(cbo开单人.List(i)) Like UCase(strText) & "*" Then
                        If intIdx = -1 Then cbo开单人.ListIndex = i
                        intIdx = i
                    End If
                Next
            End If
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo开单人_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo开单人.ListIndex = -1 Then
            cbo开单人.Text = ""
            mobjBill.开单人 = ""
        Else
            mobjBill.开单人 = NeedName(cbo开单人.Text)
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo开单人_Click
            ElseIf intIdx <> cbo开单人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo开单人.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo开单人_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp "zl9InExse", Me.Hwnd, "frmCharge"
        Case vbKeyF2
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF6 '清除当前单据内容,进入新单状态
            If mbytInState = 0 Then
                txt实收.Text = gstrDec: txt应收.Text = gstrDec
                Call ClearRows: Call Bill.ClearBill
                Call SetColNum: Call ClearMoney
                Call NewBill
                Bill.SetFocus
            End If
        Case vbKeyF7 '切换输入法
            If Not gbln简码匹配方式切换 Then Exit Sub
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
'功能:根据当前收入项目行数调整各列宽
    Dim lngW As Long
    lngW = mshMoney.Width - 60
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    mshMoney.ColWidth(0) = lngW * 0.5
    mshMoney.ColWidth(1) = lngW * 0.5
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    
    mshMoney.TextMatrix(0, 0) = "项目"
    mshMoney.TextMatrix(0, 1) = "金额"
    mshMoney.Row = 0
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    Dim strOperDoc As String
    
    On Error GoTo errH
    
    '不同药房药品出库检查方式
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '开单科室
    StrSQL = "Select 开嘱科室ID,开嘱医生 From 病人医嘱记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng医嘱ID)
    If Not rsTmp.EOF Then
        mlng开嘱科室ID = Nvl(rsTmp!开嘱科室id, 0)
        mstr开嘱医生 = Nvl(rsTmp!开嘱医生)
    End If
    If mlng开单科室ID = 0 Or mstr开嘱医生 = "" Then
        MsgBox "没有发现源医嘱信息。", vbInformation, gstrSysName
        Exit Function
    End If
      
    StrSQL = _
    "   Select A.ID, A.编码, A.名称, A.简码, 0 As 缺省, B.工作性质, D.优先级" & vbNewLine & _
    "   From 部门表 A, 部门性质说明 B," & vbNewLine & _
    "       (Select 部门id, Max(Decode(服务对象, 2, 1, 2)) As 优先级 From 部门性质说明 Where 服务对象 <> 0 Group By 部门id) D" & vbNewLine & _
    "   Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And A.ID = B.部门id" & vbNewLine & _
    "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "       And B.部门id = D.部门id And (B.服务对象 IN(1,2,3) AND B.工作性质 IN('临床','手术') Or b.工作性质='产科')" & vbNewLine & _
    "Order By 优先级,编码"
    Set mrsAll开单科室 = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    
    '69912:刘尔旋,2014-02-12,开单科室下拉列表新增主刀医生科室
    strOperDoc = Get医嘱附项内容(mlng医嘱ID, "主刀医生科室")
    
    If mbln费用登记 Then
        '就为当前选择的医技科室
        StrSQL = "(Select ID,编码,名称,简码 From 部门表 Where ID=[1]"
    Else
        '就为当前选择的医技科室或开嘱科室
        StrSQL = "(Select ID,编码,名称,简码 From 部门表 Where ID IN([1],[2])"
    End If
    
    If strOperDoc <> "" Then
        StrSQL = StrSQL & " Union " & _
                "Select ID,编码,名称,简码 From 部门表 Where 名称=[3]"
    End If
    StrSQL = StrSQL & ") Order By 编码"
    Set mrsDept = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng开单科室ID, mlng开嘱科室ID, strOperDoc)
    
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cbo开单科室.AddItem IIF(zlIsShowDeptCode, mrsDept!编码 & "-", "") & mrsDept!名称
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = mrsDept!ID
            If mbyt缺省科室 = 0 Then    '缺省医技科室:36060
                If mrsDept!ID = mlng开单科室ID Then
                    cbo开单科室.ListIndex = cbo开单科室.NewIndex
                End If
            Else
                If mrsDept!ID = mlng病人科室id Then
                    cbo开单科室.ListIndex = cbo开单科室.NewIndex
                End If
            End If
            mrsDept.MoveNext
        Next
        cbo开单科室.AddItem "其他科室…"
        cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = 0
        If cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    Else
        MsgBox "不能确定开单科室，请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '可用收费类别:"'5','E','Z'"
    If mstr收费类别 = "" Then
        StrSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 序号"
    Else
        StrSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where Instr([1],编码)>0 Order by 序号"
    End If
    'Set mrsClass = New ADODB.Recordset
    Set mrsClass = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mstr收费类别)
    If mrsClass.EOF Then
        MsgBox "没有设置可用的收费类别,请先在本地参数中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    '当只有一种可选收费类别时,不用用户选择
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(mstr收费类别, "'5'") > 0 Or InStr(mstr收费类别, "'6'") > 0 Or InStr(mstr收费类别, "'7'") > 0 Or mstr收费类别 = "" Then
        mlng药品类别ID = ExistIOClass(IIF(mint记录性质 = 1, 8, 9))
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定药品单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(mstr收费类别, "'4'") > 0 Or mstr收费类别 = "" Then
        mlng卫材类别ID = ExistIOClass(IIF(mint记录性质 = 1, 40, 41))
        If mlng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '执行部门
    StrSQL = _
        "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID and B.服务对象 IN([1],3) " & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by B.服务对象,A.编码"
    'Set mrsUnit = New ADODB.Recordset
    Set mrsUnit = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mint病人来源)
    If mrsUnit.EOF Then
        MsgBox "没有初始化部门信息,单据无法处理执行部门。请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str类别 As String, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'功能：获取最近输入的相同类别项目的执行科室ID
    Dim i As Long
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).收费类别 = str类别 _
            And mobjBill.Details(i).执行部门ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).执行部门ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).执行部门ID
                Exit Function
            End If
        End If
    Next
    If str类别 = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).执行部门ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).执行部门ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).执行部门ID
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, Optional blnEnter As Boolean)
'功能：根据单据列设置下拉列表框内容
'参数：blnEnter=是否按进入该列处理,比如执行科室保持不变
    Dim rsTmp As New ADODB.Recordset
    Dim str人员性质 As String, strTmp As String
    Dim lng科室ID As Long, strIDs As String
    Dim StrSQL As String, i As Long
    
    Bill.Clear
    On Error GoTo errH
    Select Case Bill.TextMatrix(0, lngCol)
        Case "类别"
            Bill.cboStyle = DropOlnyDown
            If cbo开单人.ListIndex <> -1 Then
                If cbo开单人.ListIndex <= UBound(marrDr) Then
                    If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 6 Then
                        str人员性质 = Split(marrDr(cbo开单人.ListIndex), "|")(6)
                    End If
                End If
            End If
        
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                For i = 1 To mrsClass.RecordCount
                    '护士类别:限制
'                    If Not (str人员性质 = "护士" And InStr(",E,M,4,", mrsClass!编码) = 0) Then
                        Bill.AddItem Bill.ListCount + 1 & "-" & mrsClass!类别
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!编码)  '存放类别编码的ASCII码
'                    End If
                    mrsClass.MoveNext
                Next
            End If
        Case "执行科室"
            Bill.cboStyle = DropDownAndEdit
            
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .收费类别) > 0 Then
                        Call GetWorkUnit(.收费细目ID, .收费类别)
                        If mrsWork.RecordCount > 0 Then
                            '取上一个药的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '进入该列时保持已确定值不变
                                lng科室ID = GetLastDeptID(.收费类别, lngRow, Mid(strIDs, 2))
                            End If
                            If lng科室ID = 0 Then lng科室ID = .执行部门ID
                            
                            '确定当前行的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIF(zlIsShowDeptCode, mrsWork!编码 & "-", "") & mrsWork!名称
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        lng科室ID = Get开单科室ID
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        '0-不明确,1-病人科室,2-病人病区,3-开单人科室,4-指定科室
                        Select Case .Detail.执行科室
                            Case 0 '不明确
                                mrsUnit.Filter = 0
                            Case 1 '病人科室
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!科室ID, 0) & " Or ID=" & .执行部门ID
                            Case 2 '病人病区
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!病区ID, 0) & " Or ID=" & .执行部门ID
                            Case 3 '操作员科室
                                mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                            Case 4 '指定科室
                                StrSQL = "Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                                    " From 收费执行科室 A,部门表 B" & _
                                    " Where A.执行科室ID=B.ID And A.收费细目ID=[1]" & _
                                    " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                                    " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                                    " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
                                    " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
                                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, .收费细目ID, mint病人来源, Val(Nvl(mrsInfo!科室ID, 0)))
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!执行科室ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .执行部门ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                                End If
                            Case 6 '开单人科室
                                mrsUnit.Filter = "ID=" & lng科室ID & " Or ID=" & .执行部门ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                        If Not mrsUnit.EOF Then
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = IIF(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                                '刘兴洪:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.NewIndex) = mrsUnit!ID
                                    
                                    '设置缺省执行科室
                                    If Not blnEnter Then '进入该列时保持已确定值不变
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '与上一行非药品相同
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).执行部门ID And mobjBill.Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).收费类别) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                        End If
                            
                        If Not blnEnter And .Detail.执行科室 = 4 Then '执行科室为指定科室的,缺省为操作员所在科室
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.部门ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 Then '如果没有则取现有的执行科室
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .执行部门ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .执行部门ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .执行部门ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitFace()
'功能：根据表单要完成的功能设置界面布局
    Dim arrHead() As String, i As Long, arrBaby As Variant
    
    '公用单据表格式
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.项目
        .PrimaryCol = BillCol.项目
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.行) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 Then
            .ColData(BillCol.行) = 5
            
            .ColData(BillCol.类别) = IIF(gbln收费类别, 3, 5)
            If mblnOne Then .ColData(BillCol.类别) = 5
            
            .ColData(BillCol.项目) = 1 '项目输入,按扭可选
            .ColData(BillCol.数次) = 4 '数/次输入
            '刘兴洪:27990 2010-02-23 12:04:37
            .ColData(BillCol.商品名) = 5 '规格跳过
            .ColData(BillCol.规格) = 5 '规格跳过
            .ColData(BillCol.单位) = 5 '单位跳过
            .ColData(BillCol.付数) = 5 '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = 5 '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.应收金额) = 5 '应收金额跳过
            .ColData(BillCol.实收金额) = 5 '实收金额跳过
            .ColData(BillCol.执行科室) = 3 '默认取开单科室或上一科室
            .ColData(BillCol.标志) = 5 '标志缺省跳过,当为手术时,设为复选(-1)
            .ColData(BillCol.类型) = 5 '类型缺省跳过
        End If
        .SetColColor BillCol.类别, &HE7CFBA
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.执行科室, &HE7CFBA
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        .SetColColor BillCol.标志, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    If gbyt药品名称显示 <> 2 Then
        '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        Bill.ColWidth(BillCol.商品名) = 0
    Else
        If Bill.ColWidth(BillCol.商品名) = 0 Then
             Bill.ColWidth(BillCol.商品名) = GetOrigColWidth(BillCol.商品名)
        End If
    End If
    
    Call SetMoneyList

    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln简码匹配方式切换 '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln简码匹配方式切换
    If mbytInState = 0 Then
        '简码匹配方式：0-拼音,1-五笔
        i = Val(zlDatabase.GetPara("简码方式"))
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
    End If

    '标题
    If mbln费用登记 Then
        lblTitle.Caption = gstrUnitName & "零费耗用登记"
    Else
        If mint记录性质 = 1 Then
            lblTitle.Caption = gstrUnitName & "病人收费单"
        ElseIf mint记录性质 = 2 Then
            lblTitle.Caption = gstrUnitName & "病人记帐单"
        End If
    End If
    txt应收.Text = gstrDec: txt实收.Text = gstrDec
    
    Select Case mbytInState
        Case 0 '执行
            Call SetShowCol
            cmdSelWholeSet.Visible = True
            cmdSaveWholeSet.Visible = zlCheckPrivs(mstrPrivs, "增加成套项目")
        Case 1 '查阅
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
            cmdSelWholeSet.Visible = False
            cmdSaveWholeSet.Visible = zlCheckPrivs(mstrPrivs, "增加成套项目")
            cmdSaveWholeSet.Left = cmdSelWholeSet.Left
        Case 3 '销帐
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            cmdSelWholeSet.Visible = False
            cmdSaveWholeSet.Visible = zlCheckPrivs(mstrPrivs, "增加成套项目")
            cmdSaveWholeSet.Left = cmdSelWholeSet.Left
            '暂时不支持部份删除
            If mint记录性质 <> 1 And False Then
                Call ShowDeleteCol(True)
                Bill.Active = True
            Else
                Bill.Active = False
            End If
    End Select
    
    If mbytInState <> 0 Then
        lblPreNO.Visible = False: txtPreNO.Visible = False
        lbl应收.Top = lbl应收.Top + txtPreNO.Height / 2
        txt应收.Top = txt应收.Top + txtPreNO.Height / 2
        lbl实收.Top = lbl实收.Top + txtPreNO.Height * 0.75
        txt实收.Top = txt实收.Top + txtPreNO.Height * 0.75
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'界面设置为不可修改状态
    cboNO.Locked = Not bln
    txt姓名.Locked = Not bln
    cbo开单科室.Locked = Not bln
    cbo开单人.Locked = Not bln
    
    chk加班.Enabled = bln
    cboBaby.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
End Sub

Private Function GetPatient(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：获取病人信息
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    
    mblnWarnCloseed = False
    mintWarn = -1: mstrWarn = ""
    Set mrsWarn = New ADODB.Recordset
    
    txt姓名.ForeColor = Me.ForeColor
    Set mrsInfo = New ADODB.Recordset
    
    If mint病人来源 = 2 Then '对住院病人是否具有强制记帐权限
        If InStr(mstrPrivs, "出院未结强制记帐") > 0 And InStr(mstrPrivs, "出院结清强制记帐") > 0 Then
            StrSQL = ""
        ElseIf InStr(mstrPrivs, "出院未结强制记帐") > 0 Then
            StrSQL = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)<>0)"
        ElseIf InStr(mstrPrivs, "出院结清强制记帐") > 0 Then
            StrSQL = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)=0)"
        Else
            StrSQL = " And B.出院日期 is NULL And Nvl(B.状态,0)<>3"
        End If
    End If
    
    '字段中使用参数时，如果不明确类型(如Null值),则结果为adVarChar类型
    StrSQL = "Select" & _
        " A.病人ID,Nvl(B.主页ID,0) 主页ID,To_Number(Nvl(B.当前病区ID,[3])) as 病区ID," & _
        "       Nvl(B.出院科室ID,[3]) as 科室ID,B.入院日期,B.出院日期," & _
        "       A.门诊号,B.住院号,B.出院病床 as 床号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄 ,Nvl(B.费别,A.费别) as 费别," & _
        "       A.担保人," & IIF(mint病人来源 = 2 And mint记录性质 = 2, "Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,", "A.担保额,") & _
        "       Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Y.编码 as 付款码,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人," & _
        "       B.住院医师,zl_PatiDayCharge(A.病人ID) as 当日额,Nvl(B.险类,A.险类) as 险类,Nvl(B.病人性质,0) as 病人性质,B.审核标志,B.备注 as 病人备注" & _
        " From 病人信息 A,病案主页 B,病人余额 X,医疗付款方式 Y" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=X.病人ID(+)  And X.类型(+) = " & IIF(mint病人来源 = 1, 1, 2) & StrSQL & _
        " And A.病人ID=[1] And B.主页ID(+)=[2] And A.医疗付款方式=Y.名称(+)"
        
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng病人ID, lng主页ID, mlng病人科室id)
    If Not mrsInfo.EOF Then
        mstr住院医生 = Nvl(mrsInfo!住院医师)
        If Not IsNull(mrsInfo!险类) Then
            txt姓名.ForeColor = vbRed
        End If
        
        '除了门诊划价以外要处理的内容
        If mint记录性质 = 2 Then
            If mint病人来源 = 2 Then
                '49501:住院
                If zlIsAllowFeeChange(mrsInfo!病人ID, Val(Nvl(mrsInfo!主页ID)), Val(Nvl(mrsInfo!审核标志))) = False Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt姓名.Text = "":
                    mlng病人ID = 0
                    If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
            End If
            
            '刷新病人费用状况
            Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, IIF(mint病人来源 = 1, 0, mlng主页ID), mcurModiMoney)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!预交余额
                cmdCancel.Tag = rsTmp!费用余额
                txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
            End If
            sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag), gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag), "0.00")
            
            '刷新报警信息
            StrSQL = "Select Nvl(报警方法,1) as 报警方法," & _
                " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
                " Where 适用病人=[2] And " & IIF(mint病人来源 = 1, "Nvl(病区ID,0)=0", "病区ID=[1]")
            Set mrsWarn = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(Nvl(mrsInfo!病区ID, 0)), CStr(Nvl(mrsInfo!适用病人)))
            
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------
            '刘兴洪:26952
            Dim cur余额 As Currency, curItemMoney As Currency, curTotal As Double
            curItemMoney = 0
            curTotal = GetBillTotal(mobjBill)
            cur余额 = Val(txt实收.Tag)
            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)

            If mbln费用登记 = False Then    '30504
            
                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, Val(Nvl(mrsInfo!当日额)) - mcurModiMoney, curTotal, _
                     Nvl(mrsInfo!担保额, 0), "", "", mstrWarn, mintWarn, , True)
                '返回:0;没有报警,继续
                '     1:报警提示后用户选择继续
                '     2:报警提示后用户选择中断
                '     3:报警提示必须中断
                '     4:强制记帐报警,继续
                '     5.报警提示后用户选择继续,但只允许保存存为划价单
                If mbytWarn = 2 Or mbytWarn = 3 Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt姓名.Text = "":
                    mlng病人ID = 0
                    If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                If mrsWarn.EOF Then mrsWarn.Close '用于后面状态判断
            End If
        End If
                            
        '住院记帐才处理的内容
        If mint病人来源 = 2 Then
            '急诊费用
            If Not IsNull(mrsInfo!险类) Then
                chk急诊.value = 0: chk急诊.Visible = True
            Else
                chk急诊.value = 0: chk急诊.Visible = False
            End If
            
            '发生时间
            If Not IsNull(mrsInfo!出院日期) Then
                txtDate.Text = Format(mrsInfo!出院日期, "yyyy-MM-dd HH:mm:ss")
            Else
                txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        
        Call LoadPatientBaby(cboBaby, mrsInfo!病人ID, mrsInfo!主页ID)
        
        '显示病人信息
        txt姓名.Text = Nvl(mrsInfo!姓名)
        txt性别.Text = Nvl(mrsInfo!性别)
        txt年龄.Text = Nvl(mrsInfo!年龄)
        txt费别.Text = Nvl(mrsInfo!费别)
        txt付款方式.Text = Nvl(mrsInfo!医疗付款方式)
        txt付款方式.Tag = Nvl(mrsInfo!付款码, 0) '不要填写为空
        txt床号.Text = Nvl(mrsInfo!床号)
        
        '刘兴洪 问题:26953 日期:2009-12-25 15:21:47
        txt病人备注 = Nvl(mrsInfo!病人备注)
        If mint病人来源 = 1 Then
            lbl住院号.Caption = "门诊号"
            txt住院号.Text = Nvl(mrsInfo!门诊号)
        Else
            lbl住院号.Caption = "住院号"
            txt住院号.Text = Nvl(mrsInfo!住院号)
        End If
        
        txt担保人.Text = Nvl(mrsInfo!担保人)
        txt担保额.Text = Format(Nvl(mrsInfo!担保额), "0.00")
        
        With mobjBill
            .病人ID = Nvl(mrsInfo!病人ID, 0)
            .主页ID = Nvl(mrsInfo!主页ID, 0)
            .病区ID = Nvl(mrsInfo!病区ID, 0)
            .科室ID = Nvl(mrsInfo!科室ID, 0)
            .床号 = Nvl(mrsInfo!床号)
            .标识号 = IIF(mint病人来源 = 1, Nvl(mrsInfo!门诊号), Nvl(mrsInfo!住院号))
            .姓名 = Nvl(mrsInfo!姓名)
            .性别 = Nvl(mrsInfo!性别)
            .年龄 = Nvl(mrsInfo!年龄)
            .费别 = Nvl(mrsInfo!费别)
        End With
        
        '在第一次进入时读取病人审批费用项目信息
        If Not Visible And mint病人来源 = 2 And mint记录性质 = 2 And mbytInState = 0 Then Set mrsMedAudit = GetAuditRecord(mrsInfo!病人ID, mrsInfo!主页ID)
        
        GetPatient = True
    Else
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'功能：计算或重新计算指定行或所有行的金额
'参数：lngRow=指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long
    Dim strMainRows As String
    Dim bln从项汇总折扣 As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIF(lngRow = 0, 1, lngRow) To IIF(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln从项汇总折扣 = False
        If gbln从项汇总折扣 And Not mbln费用登记 Then                    '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
            If mobjBill.Details(i).从属父号 > 0 Then    '从项
                bln从项汇总折扣 = Not mobjBill.Details(mobjBill.Details(i).从属父号).Detail.屏蔽费别
                If bln从项汇总折扣 And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).从属父号      '单独计算一行的时候
            Else
                If CheckItemHaveSub(i) Then                          '主项或独立项
                     bln从项汇总折扣 = Not mobjBill.Details(i).Detail.屏蔽费别
                     If bln从项汇总折扣 Then strMainRows = strMainRows & "," & i  '一页可能有多个主从项,先记录主项行号,后面再重算主项折扣
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln从项汇总折扣)
    Next
    
    '重算所有主项,不能用bln从项汇总折扣变量,因为可能在遇到不是从项的行时已改变
    If gbln从项汇总折扣 And Not mbln费用登记 Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc重算主项实收(Split(strMainRows, ",")(i))
        Next
    End If
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln从项汇总折扣 As Boolean)
'功能：计算或重新计算指定行的金额
'参数：lngRow=指定行
'说明：1.ExpenseBill集合的索引对应单据的行号
'      2.变价只能对应一个收入项目:mobjBill.Details(lngRow).InComes(1)
'      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
'      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    Dim rsTmp As New ADODB.Recordset
    Dim strInfo As String, i As Long
    Dim dblMoney As Double '用户输入的变价金额
    Dim dbl加班加价率 As Double
        
    Dim dblAllTime As Double, dblPrice As Double, dblPriceSingle As Double
    
    On Error GoTo errH
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).收费细目ID)
    End If
    
    gstrSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,B.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL)) " & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID)
    If Not rsTmp.EOF Then
        With mobjBill.Details(lngRow)
            '先获取操作员以前输入的变价金额
            If .Detail.变价 Then
                If InStr(",5,6,7,", .收费类别) > 0 Or (.收费类别 = "4" And .Detail.跟踪在用) Then
                    '计算药品时价(分批或不分批)
                    '必然有记录(输入该项目时已判断)
                    dblAllTime = .付数 * .数次
                    If mbln药房单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                        dblAllTime = dblAllTime * .Detail.药房包装 '库存时价按售价数量进行计算
                    End If
                    If dblAllTime <> 0 Then
                        dblPrice = Get时价药品应收金额(.执行部门ID, .收费细目ID, dblAllTime, gstrDec, dblPriceSingle)
                        If dblAllTime <> 0 Then
                            '数量未分解完毕
                            If InStr(",5,6,7,", .收费类别) > 0 Then
                                MsgBox "第 " & lngRow & " 行时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            Else
                                MsgBox "第 " & lngRow & " 行时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            End If
                            dblMoney = 0
                        Else
                            '注意：货币型最多只能保留4位小数,且不四舍五入,所以需要手工舍入;而用其它型在计算精度上又有问题
                            dblAllTime = .付数 * .数次
                            If mbln药房单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                                dblAllTime = dblAllTime * .Detail.药房包装 '按售价数量计算实价
                            End If
                            dblMoney = IIF(dblPriceSingle = 0, Format(dblPrice / dblAllTime, gstrDecPrice), dblPriceSingle) '这里结果是按售价单位
                        End If
                    Else
                        dblMoney = 0
                    End If
                Else
                    If .InComes.Count = 0 Then  '第一次计算金额取缺省值
                        dblMoney = IIF(IsNull(rsTmp!缺省价格), 0, rsTmp!缺省价格)
                    Else                        '获取操作员以前输入的变价金额
                        dblMoney = .InComes(1).标准单价
                        '如果用户输入的变价不满足变价范围，则取缺省值
                        If Abs(dblMoney) > Abs(IIF(IsNull(rsTmp!现价), 0, rsTmp!现价)) Then
                            dblMoney = IIF(IsNull(rsTmp!缺省价格), 0, rsTmp!缺省价格)
                        End If
                    End If
                End If
            End If
        End With
        
        '再清除原有记录
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        
        '填写现有费用记录
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .收入项目ID = rsTmp!收入项目ID
                .收入项目 = rsTmp!名称
                .收据费目 = IIF(IsNull(rsTmp!收据费目), "", rsTmp!收据费目)
                .原价 = IIF(IsNull(rsTmp!原价), 0, rsTmp!原价)
                .现价 = IIF(IsNull(rsTmp!现价), 0, rsTmp!现价)
                If mobjBill.Details(lngRow).Detail.变价 Then
                    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And mbln药房单位 Then
                        .标准单价 = Format(dblMoney * mobjBill.Details(lngRow).Detail.药房包装, gstrDecPrice)
                    Else
                        .标准单价 = Format(dblMoney, gstrDecPrice)
                    End If
                Else
                    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And mbln药房单位 Then
                        .标准单价 = Format(Nvl(rsTmp!现价, 0) * mobjBill.Details(lngRow).Detail.药房包装, gstrDecPrice)
                    Else
                        .标准单价 = Format(Nvl(rsTmp!现价, 0), gstrDecPrice)
                    End If
                End If
                '应收金额=单价 * 付数 * 数次
                If mobjBill.Details(lngRow).Detail.变价 And (InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 _
                        Or mobjBill.Details(lngRow).收费类别 = "4" And mobjBill.Details(lngRow).Detail.跟踪在用) Then
                    .应收金额 = dblPrice '保证应收金额与零售金额没有误差
                Else
                    .应收金额 = .标准单价 * mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
                End If
                
                '附加手术费率用计算(所有收入项目)
                If mobjBill.Details(lngRow).附加标志 = 1 And mobjBill.Details(lngRow).收费类别 = "F" Then
                    .应收金额 = .应收金额 * IIF(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
                End If
                '加班费用率计算
                dbl加班加价率 = 0
                If mobjBill.加班标志 = 1 And mobjBill.Details(lngRow).Detail.加班加价 Then
                    dbl加班加价率 = Nvl(rsTmp!加班加价率, 0) / 100
                    .应收金额 = .应收金额 * (1 + dbl加班加价率)
                End If
                
                .应收金额 = CCur(Format(.应收金额, gstrDec))
                
                dblAllTime = mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
                If mbln药房单位 And InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
                    dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.药房包装
                End If
                
                If mbln费用登记 Or .应收金额 = 0 Then
                    .实收金额 = 0
                Else
                    If mobjBill.Details(lngRow).Detail.屏蔽费别 Or bln从项汇总折扣 Then
                        .实收金额 = .应收金额
                    Else
                        .实收金额 = CCur(Format(ActualMoney(mobjBill.费别, .收入项目ID, .应收金额, _
                            mobjBill.Details(lngRow).收费细目ID, mobjBill.Details(lngRow).执行部门ID, _
                            dblAllTime, dbl加班加价率), gstrDec))
                    End If
                End If
                
                '获取项目保险信息,医保病人才处理,不需要连接医保
                If Not IsNull(mrsInfo!险类) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(lngRow).收费细目ID, .实收金额, False, mrsInfo!险类, _
                        mobjBill.Details(lngRow).摘要 & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Details(lngRow).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(lngRow).保险大类ID = Val(Split(strInfo, ";")(1))
                        .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        mobjBill.Details(lngRow).保险编码 = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).摘要 = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.类型 = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
                
                '实收金额存入Key中,以处理分币问题(即Key中存放原始实收金额,不变)
                mobjBill.Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
            End With
            rsTmp.MoveNext
        Next
    Else
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'功能：刷新显示指定行或所有行的内容
'参数：lngRow=指定行,为0表示显示所有行
'说明：ExpenseBill集合的索引对应单据的行号
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
        sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + curTotal, gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - curTotal, "0.00")
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '清除单据行
    For i = 1 To Bill.Cols - 1
        '输入时收费类别不清除
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).收费类别 <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).收费类别)
    End If
    
    '刷新单据行
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "类别"
                '浏览单据或从属项目只(能)显示名称
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类别名称
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.名称
            Case "商品名"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.商品名
            Case "规格"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.规格
            Case "单位"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And mbln药房单位 Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.药房单位
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.计算单位
                End If
            Case "付数"
                Bill.TextMatrix(lngRow, i) = IIF(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数)
            Case "数次"
                '数次在第一次显示时已默认设置为1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).数次, 5)
            Case "单价"
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                dbl单价 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl单价 = dbl单价 + mobjBill.Details(lngRow).InComes(j).标准单价
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl单价, gstrDecPrice)
            Case "应收金额"
                '应收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Details(lngRow).InComes(j).应收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "实收金额"
                '实收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Details(lngRow).InComes(j).实收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "执行科室"
                If mobjBill.Details(lngRow).执行部门ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            Bill.TextMatrix(lngRow, i) = Get部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                        End If
                    Else
                        '浏览单据只(能)显示名称
                        Bill.TextMatrix(lngRow, i) = Get部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "标志"
                If mobjBill.Details(lngRow).收费类别 = "F" And mobjBill.Details(lngRow).附加标志 = 1 Then
                    Bill.TextMatrix(lngRow, i) = "√"
                End If
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类型
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
'功能：刷新显示收入项目费用区
    Dim blnExist As Boolean
    Dim cur实收 As Currency, cur应收 As Currency
    Dim i As Long, j As Long, k As Long
    
    mshMoney.Redraw = False
    
    '产生汇总费目
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '查找是否已经加入此类收入项目,如是则合计,否则新入
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).收入项目ID = mobjBill.Details(i).InComes(j).收入项目ID Then
                    blnExist = True: Exit For
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + mobjBill.Details(i).InComes(j).实收金额
                mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + mobjBill.Details(i).InComes(j).应收金额
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额
                End With
            End If
        Next
    Next
    
    '刷新显示
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).收入项目
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).实收金额, gstrDec)
        cur实收 = cur实收 + mcolMoneys(i).实收金额
        cur应收 = cur应收 + mcolMoneys(i).应收金额
    Next
    
    txt应收.Text = Format(cur应收, gstrDec)
    txt实收.Text = Format(cur实收, gstrDec)
    
    mshMoney.TopRow = mshMoney.Rows - 1
    mshMoney.Redraw = True
End Sub

Private Function GetCur应收() As Currency
'功能：获取病人当前单据合计金额(收费病人累加单据时用)
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCur应收 = GetCur应收 + mcolMoneys(i).应收金额
    Next
End Function

Private Function GetInputDetail(ByVal lng项目ID As Long) As Detail
'功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    
    If lngMediCareNO > 0 Then
        StrSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,A.费用类型,A.补充摘要,M.要求审批," & _
        "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
        "       Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
        "       Decode(A.类别,'4',1,C." & mstr药房包装 & ") as 药房包装," & _
        "       Decode(A.类别,'4',A.计算单位,C." & mstr药房单位 & ") as 药房单位,D.跟踪在用,A.录入限量,C.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,保险支付项目 M" & _
        " Where A.ID=C.药品ID(+) And A.ID=D.材料ID(+) And B.编码=A.类别" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[2] " & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        "       And A.ID=[1] And A.ID=M.收费细目ID(+) And M.险类(+)=[3]"

    Else
        StrSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,A.费用类型,A.补充摘要,0 as 要求审批," & _
        "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
        "        Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
        "       Decode(A.类别,'4',1,C." & mstr药房包装 & ") as 药房包装," & _
        "       Decode(A.类别,'4',A.计算单位,C." & mstr药房单位 & ") as 药房单位,D.跟踪在用,A.录入限量,C.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1" & _
        " Where A.ID=C.药品ID(+) And A.ID=D.材料ID(+) And B.编码=A.类别" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[2] " & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng项目ID, IIF(gbyt药品名称显示 = 1, 3, 1), lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .药名ID = Nvl(rsTmp!药名ID, 0)
        .编码 = rsTmp!编码
        .规格 = Nvl(rsTmp!规格)
        .药房包装 = Nvl(rsTmp!药房包装, 1)
        .药房单位 = Nvl(rsTmp!药房单位)
        .分批 = Nvl(rsTmp!分批, 0) = 1
        .变价 = Nvl(rsTmp!是否变价, 0) = 1
        .计算单位 = Nvl(rsTmp!计算单位)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .名称 = rsTmp!名称
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .服务对象 = Nvl(rsTmp!服务对象, 0)
        .类型 = Nvl(rsTmp!费用类型)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
        .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
        .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
        .录入限量 = Val("" & rsTmp!录入限量)
        .中药形态 = Val(Nvl(rsTmp!中药形态))
        .商品名 = Nvl(rsTmp!商品名)
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'功能：根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
'说明：
'      1.用于新输入或更改收费细目行！！！
'      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    '取其它中药的付数
    intPay = GetPay(lngRow)
    If Detail.类别 <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '如果该行对应的程序对象尚未初始,则加入
        With Detail
            '序号=行号,父号=0
            '次数=1,从属项目的次数由主项计算确定
            '执行部门ID:根据细目执行科室标志取
            '附加标志:以第一行为假,其它为真优先权
            '收入集=空
            If bytParent <> 0 Then
                '设置该行RowData
                Bill.RowData(lngRow) = Asc(Detail.类别)
                '初始数次
                If Detail.固有从属 = 0 Then '非固有从属
                    dblTime = Detail.从项数次
                ElseIf Detail.固有从属 = 1 Then '固定的固有从属
                    dblTime = IIF(Detail.从项数次 = 0, 1, Detail.从项数次)
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Details(bytParent).数次
                End If
            Else
                
                If InStr(",5,6,7,", Detail.类别) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Details.Add tmpIncomes, Detail, .ID, CByte(lngRow), CInt(bytParent), .类别, .计算单位, intPay, dblTime, 0, lngDoUnit, ""
        End With
    Else '如果该行已经存在,则修改
        
        If InStr(",5,6,7,", Detail.类别) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .付数 = intPay
            .附加标志 = 0
            .计算单位 = Detail.计算单位
            .收费类别 = Detail.类别
            .收费细目ID = Detail.ID
            .数次 = dblTime
            .序号 = lngRow
            .从属父号 = 0
            .执行部门ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'功能：判断该行是否应该取从属项目
'说明：仅该行收费项目有从属项目及尚未取才取。
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim StrSQL As String
    
    StrSQL = "Select count(从项ID) as NUM from 收费从属项目 where 主项ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID)
    
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).从属父号 = lngRow Then
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetSubDetails(ByVal lng项目ID As Long) As Details
'功能：返回一个收费细目的从属项目集
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
            
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    If lngMediCareNO > 0 Then
        StrSQL = _
        " Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
        "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别," & _
        "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
        "       Decode(A.类别,'4',1,D." & mstr药房包装 & ") as 药房包装,A.服务对象," & _
        "       Decode(A.类别,'4',A.计算单位,D." & mstr药房单位 & ") as 药房单位," & _
        "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,G.要求审批,D.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1,保险支付项目 G" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And A.ID=E.材料ID(+)" & _
        "       And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        "       And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[2] " & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        "       And C.主项ID=[1] And A.ID=G.收费细目ID(+) And G.险类(+)=[3] " & _
        " Order by 编码"
    Else
        StrSQL = _
        " Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
        "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别," & _
        "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
        "       Decode(A.类别,'4',1,D." & mstr药房包装 & ") as 药房包装,A.服务对象," & _
        "       Decode(A.类别,'4',A.计算单位,D." & mstr药房单位 & ") as 药房单位," & _
        "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,0 as 要求审批,D.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And A.ID=E.材料ID(+)" & _
        "       And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        "       And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[2] " & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        "       And C.主项ID=[1] " & _
        " Order by 编码"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng项目ID, IIF(gbyt药品名称显示 = 1, 3, 1), lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
                .ID = rsTmp!ID
                .药名ID = Nvl(rsTmp!药名ID, 0)
                .编码 = rsTmp!编码
                .变价 = Nvl(rsTmp!是否变价, 0) = 1
                .规格 = Nvl(rsTmp!规格)
                .药房包装 = Nvl(rsTmp!药房包装, 1)
                .药房单位 = Nvl(rsTmp!药房单位)
                .计算单位 = Nvl(rsTmp!计算单位)
                .分批 = Nvl(rsTmp!分批, 0) = 1
                .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
                .类别 = rsTmp!类别
                .类别名称 = rsTmp!类别名称
                .名称 = rsTmp!名称
                .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
                .执行科室 = Nvl(rsTmp!执行科室, 0)
                .服务对象 = Nvl(rsTmp!服务对象, 0)
                .固有从属 = Nvl(rsTmp!固有从属, 0)
                .从项数次 = Nvl(rsTmp!从项数次, 1)
                .类型 = Nvl(rsTmp!费用类型)
                .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
                .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
                .中药形态 = Val(Nvl(rsTmp!中药形态))
                .商品名 = Nvl(rsTmp!商品名)
                GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .别名, .规格, .计算单位, .说明, .屏蔽费别, _
                    .药房包装, .药房单位, .分批, .变价, .加班加价, .执行科室, .服务对象, .类型, .补充摘要, .固有从属, .从项数次, .跟踪在用, , , , , , .要求审批, , .中药形态, .商品名
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'功能：删除指定收费项目行
'说明：这时不处理从属行的删除,但要对其它单据行从属关系作相应的调整
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).从属父号 > lngRow Then
            mobjBill.Details(i).从属父号 = mobjBill.Details(i).从属父号 - 1
        End If
        mobjBill.Details(i).序号 = mobjBill.Details(i).序号 - 1 '序号与行号对应
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
    chk加班.value = IIF(OverTime, 1, 0)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                
    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo开单科室_Click
    With mobjBill
        .门诊标志 = mint病人来源
        .开单人 = NeedName(cbo开单人.Text)
        If cbo开单科室.ListIndex = -1 Then
            .开单部门ID = 0
        Else
            .开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        .发生时间 = CDate(txtDate.Text)
        .加班标志 = chk加班.value
        .划价人 = UserInfo.姓名
        .操作员编号 = UserInfo.编号
        .操作员姓名 = UserInfo.姓名
    End With
End Sub

Private Sub ClearMoney()
'功能：清除费用显示区
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
'功能:保存当前输入的记帐单据(适用住院记帐、划价、或对两者的修改)
'入口:mobjBill=单据对象
'出口:保存是否成功
    Dim int行号 As Integer, int序号 As Integer, int价格父号 As Integer
    Dim dbl数次 As Double, dbl单价 As Double
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim arrSQL As Variant, i As Long, j As Long
    Dim int划价 As Integer, bln上传 As Boolean
    Dim StrSQL As String, strStuffDept As String '记录卫料发料部门
    Dim strDeptIDs As String, str汇总号 As String
    Dim cllProExeute As New Collection, varTemp As Variant
    Dim rsTmp As ADODB.Recordset
    Dim lng医疗小组ID As Long
    Dim blnTrans As Boolean
    
    strNos = ""
    If mstrOriginalNO = "" Then
        If mint记录性质 = 1 Then
            mobjBill.NO = zlDatabase.GetNextNo(13)
        Else
            mobjBill.NO = zlDatabase.GetNextNo(14)
        End If
    Else
        mobjBill.NO = mstrOriginalNO
    End If
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    
    lng医疗小组ID = zlGet医疗小组ID
    int序号 = 0
    arrSQL = Array()
    Set cllProExeute = New Collection
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.数次 <> 0 Then
            For Each mobjBillIncome In mobjBillDetail.InComes
                int序号 = int序号 + 1 '当前记录序号
                '单据主体
                With mobjBill
                    If mint病人来源 = 2 Then
                        gstrSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.主页ID) & "," & _
                            IIF(.标识号 = "", "NULL", "'" & .标识号 & "'") & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .床号 & "','" & .费别 & "'," & _
                            ZVal(.病区ID) & "," & ZVal(.科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
                    Else
                        If mint记录性质 = 2 Then
                            gstrSQL = "zl_门诊记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & _
                                IIF(.标识号 = "", "NULL", "'" & .标识号 & "'") & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "'," & _
                                "'" & .费别 & "'," & .加班标志 & "," & .婴儿费 & "," & _
                                ZVal(.科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                        Else
                            gstrSQL = "zl_门诊划价记录_Insert('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.主页ID) & "," & _
                                IIF(.标识号 = "", "NULL", "'" & .标识号 & "'") & ",'" & IIF(Val(txt付款方式.Tag) = 0, "", txt付款方式.Tag) & "','" & .姓名 & "'," & _
                                "'" & .性别 & "','" & .年龄 & "','" & .费别 & "'," & .加班标志 & "," & _
                                  ZVal(.科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                        End If
                    End If
                End With
                
                '收费细目部份
                With mobjBillDetail
                    '处理从属父号
                    If .序号 <> int行号 Then
                        int行号 = .序号
                        int价格父号 = int序号
                        
                        '重新处理从属父号
                        If mobjBill.Details(.序号).从属父号 = 0 Then
                            For i = .序号 + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).从属父号 = .序号 Then
                                    mobjBill.Details(i).从属父号 = int序号
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                    
                    If mint病人来源 = 2 Then
                        gstrSQL = gstrSQL & IIF(.保险项目否, 1, 0) & "," & ZVal(.保险大类ID) & ",'" & .保险编码 & "',"
                    ElseIf mint记录性质 = 1 Then
                        gstrSQL = gstrSQL & "NULL,"
                    End If
                    
                    dbl数次 = .数次
                    If InStr(",5,6,7,", .收费类别) > 0 And mbln药房单位 Then
                        dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                    End If
                    gstrSQL = gstrSQL & IIF(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & "," & ZVal(.执行部门ID) & ","
                End With
                
                '收入项目部份
                With mobjBillIncome
                    dbl单价 = .标准单价
                    If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 And mbln药房单位 Then
                        dbl单价 = Format(.标准单价 / mobjBillDetail.Detail.药房包装, gstrDecPrice)
                    End If
                    gstrSQL = gstrSQL & IIF(int价格父号 = int序号, "NULL", int价格父号) & "," & .收入项目ID & "," & _
                        "'" & .收据费目 & "'," & dbl单价 & "," & .应收金额 & "," & .实收金额 & ","
                    If mint病人来源 = 2 Then
                        gstrSQL = gstrSQL & IIF(.统筹金额 = 0, "NULL", .统筹金额) & ","
                    End If
                End With
                                                
                '其它部分
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                
                '是否只生成划价单
    '                If mbln费用登记 Then
    '                    int划价 = 0 '零耗费用登记不必生成划价单
    '                Else
    '                    '本来应是诊疗类别,应按执行医嘱的类别来判断
    '                    If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 Then
    '                        int划价 = IIF(InStr(gstr发送划价单, "5") > 0, 1, 0)
    '                    Else
    '                        int划价 = IIF(InStr(gstr发送划价单, mobjBillDetail.收费类别) > 0, 1, 0)
    '                    End If
    '                End If
                If int划价 = 0 Then bln上传 = True '只要存在不是划价单就要上传
                
                '收集卫料发料部门,以便自动发料,门诊病人仅记帐时(发送为划价时不管),住院病人只有记帐
                'mint病人来源 :1-门诊病人,2-住院病人
                'mint记录性质 :1-收费(划价),2-记帐(门/住)
                With mobjBillDetail
                    Select Case .收费类别
                    Case "4"    '卫材
                        If (mint病人来源 = 1 And mint记录性质 = 2 And gbln门诊自动发料 Or mint病人来源 = 2 And gbln住院自动发料) And int划价 = 0 Then
                            If .执行部门ID <> 0 And .Detail.跟踪在用 Then
                                If InStr("," & strStuffDept, "," & .执行部门ID & ",") = 0 Then
                                    strStuffDept = strStuffDept & "," & .执行部门ID
                                End If
                            End If
                        End If
                    Case "5", "6", "7"  '药品
                            If gbln收费后自动发药 And mint病人来源 = 1 And int划价 = 0 Then
                                   If .执行部门ID <> 0 And Not gbln分离发药 Then
                                       If InStr(strDeptIDs & ",", "," & .执行部门ID & ",") = 0 Then
                                           strDeptIDs = strDeptIDs & "," & .执行部门ID
                                       End If
                                   End If
                               End If
                    End Select
                End With
                
                If mint病人来源 = 2 Then
                    gstrSQL = gstrSQL & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                        "0," & IIF(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                        "NULL,'" & mobjBillDetail.摘要 & "'," & chk急诊.value & "," & ZVal(mlng医嘱ID) & "," & _
                        "Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.类型 & "'," & _
                        IIF(mobjBill.开单部门ID = mlng开嘱科室ID, "1", "0") & "," & mlng开单科室ID & ",NULL" & IIF(lng医疗小组ID = 0, "", "," & lng医疗小组ID) & ")"
                Else
                    If mint记录性质 = 2 Then
                        gstrSQL = gstrSQL & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                            IIF(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                            "NULL,'" & mobjBillDetail.摘要 & "'," & ZVal(mlng医嘱ID) & ")"
                    Else
                        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                            IIF(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                            "'" & mobjBillDetail.摘要 & "'," & ZVal(mlng医嘱ID) & ")"
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
            Next
        End If
    Next
    
    '-----------------------------------------------------------------------------------------------------------------
    If mstrOriginalNO = "" Then
        '插入医嘱院加费用
        gstrSQL = "ZL_病人医嘱附费_Insert(" & mlng医嘱ID & "," & mlng发送号 & "," & mint记录性质 & ",'" & mobjBill.NO & "')"
    Else
        '补主费用
        gstrSQL = "ZL_病人医嘱发送_计费(" & mlng医嘱ID & "," & mlng发送号 & ")"
    End If
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
    
    '修改前退除原单据
    If mstrInNO <> "" Then
        '先判断是否医保病人记的帐,并作合法性检查(进入修改时已作了一次相关判断)
        If mint病人来源 = 2 Then
            '去掉了医保连接匹配检查
            intInsure = BillExistInsure(mstrInNO)
        End If
        If mint病人来源 = 2 Then
            gstrSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            If mint记录性质 = 2 Then
                gstrSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Else
                gstrSQL = "zl_门诊划价记录_DELETE('" & mstrInNO & "')"
            End If
        End If
        If gstrSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
        End If
    End If
    
    If UBound(arrSQL) >= 0 Then
        '对SQL序列按收费细目ID排序
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        '执行SQL语句
        strTmp = ""
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '-----------------------------------------------------------------------
            '执行自动发料
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                varTemp = Split(strStuffDept, ",")
                For i = 0 To UBound(varTemp)
                    '69902:刘尔旋,2014-02-09,只对同开单科室一致的执行科室项目进行自动发料
                    If Val(varTemp(i)) = Val(cbo开单科室.ItemData(cbo开单科室.ListIndex)) Then
                        StrSQL = "zl_材料收发记录_处方发料(" & Val(varTemp(i)) & ",25,'" & mobjBill.NO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                        zlAddArray cllProExeute, StrSQL
                    End If
                Next
            End If
             
            ''            '-----------------------------------------------------------------------
            ''            '收费后自动发药,记帐不自动发药,收费且不是保存为划价单,或者门诊记帐
            ''            '--刘兴洪:门诊暂不处理
            ''            If strDeptIDs <> "" Then
            ''                strDeptIDs = Mid(strDeptIDs, 2)
            ''                varTemp = Split(strDeptIDs, ",")
            ''                For i = 0 To UBound(varTemp)
            ''                    strSQL = "ZL_药品收发记录_处方发药(" & Val(varTemp(i)) & ",8,'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & mobjBill.Pages(P).开单人 & "')"
            ''                    zlAddArray cllProExeute, strSQL
            ''                Next
            ''            End If
            ''
            '准备自动发药(仅普通记帐),必须在事务中才能读到数据
            If mblnSendMateria Then
                Set rsTmp = Get待发药清单(mobjBill.NO, Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"))
                If rsTmp.RecordCount > 0 Then
                    str汇总号 = zlDatabase.GetNextNo(20)
                    For i = 0 To rsTmp.RecordCount - 1
                        StrSQL = "ZL_药品收发记录_部门发药(" & rsTmp!库房ID & "," & rsTmp!ID & ",'" & UserInfo.姓名 & "',to_date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null,Null,Null," & str汇总号 & ")"
                        zlAddArray cllProExeute, StrSQL
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            '执行发药和发料
            zlExecuteProcedureArrAy cllProExeute, Me.Caption, False, False
            '-----------------------------------------------------------------------
            
            
            '医保接口
            '1.医保记帐作废上传
            If mint病人来源 = 2 And mstrInNO <> "" And intInsure <> 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.记帐实时上传
            If mint病人来源 = 2 And bln上传 And Not IsNull(mrsInfo!险类) Then
                '医保传输费用明细
                If gclsInsure.GetCapability(support记帐上传, mlng病人ID, mrsInfo!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, mrsInfo!险类) Then
                    strTmp = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!险类) Then
                        gcnOracle.RollbackTrans
                        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保接口
        '1.医保记帐作废上传
        If mint病人来源 = 2 And mstrInNO <> "" And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """向医保传送失败,该单据的费用已删除！", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.记帐实时上传
        If mint病人来源 = 2 And bln上传 And Not IsNull(mrsInfo!险类) Then
            '医保传输费用明细
            If MCPAR.记帐上传 And MCPAR.记帐完成后上传 Then
                strTmp = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!险类) Then
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                    Else
                        MsgBox "单据""" & mobjBill.NO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '加入单据历史记录(所有类型单据)
        cboNO.AddItem mobjBill.NO, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i '只显示10个
        Next
        
        '医保接口
        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
    End If
    '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
    strNos = mobjBill.NO
    
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Integer
'功能：根据单据号读取一张单据并将其填入表格
'参数：strNO=单据号
'      blnDelete=是否读取要退费的单据
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String, intSign As Integer
    Dim curTotal As Currency, cur应收Total As Currency
    Dim StrSQL As String, i As Long
    Dim intInsure As Integer, blnDo As Boolean
    Dim blnNOMoved As Boolean
    
    If mbytInState = 1 Then
        If mint记录性质 = 1 Or (mint记录性质 = 2 And mint病人来源 = 1) Then
            blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNO, "记录性质=", mint记录性质)
        Else
            blnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, "记录性质=", mint记录性质)
        End If
    End If
    
    On Error GoTo errH
    
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    
    If mstrFeeTab = "住院费用记录" Then
        StrSQL = _
        " Select A.病人ID,Nvl(A.主页ID,0) 主页ID,A.姓名,A.性别,A.年龄,A.费别,A.床号,A.标识号," & _
        " A.病人病区ID,A.开单部门ID,A.加班标志,A.婴儿费,A.开单人,A.划价人,A.操作员姓名," & _
        " A.开单部门ID," & IIF(zlIsShowDeptCode, "C.编码||'-'||", "") & "C.名称 as 开单部门,A.发生时间," & _
        " B.医疗付款方式,B.担保人,B.担保额,A.是否急诊,B1.备注 as 病人备注" & _
        " From 住院费用记录 A,病人信息 B,部门表 C,病案主页 B1" & _
        " Where Rownum=1  And A.病人id=B1.病人id(+) and A.主页id=B1.主页ID(+) And NO=[1] And A.记录性质=[2]" & _
        " And A.病人ID=B.病人ID And Instr([3],A.记录状态)>0" & _
        IIF(mstrTime <> "", " And A.登记时间=[4]", "") & _
        " And A.开单部门ID=C.ID"
    Else
        StrSQL = _
        " Select A.病人ID,0 as 主页ID,A.姓名,A.性别,A.年龄,A.费别,A.付款方式 as 床号,A.标识号," & _
        " 0 as 病人病区ID,A.开单部门ID,A.加班标志,A.婴儿费,A.开单人,A.划价人,A.操作员姓名," & _
        " A.开单部门ID," & IIF(zlIsShowDeptCode, "C.编码||'-'||", "") & "C.名称 as 开单部门,A.发生时间," & _
        " B.医疗付款方式,B.担保人,B.担保额,A.是否急诊,Null as 病人备注" & _
        " From 门诊费用记录 A,病人信息 B,部门表 C" & _
        " Where Rownum=1  And NO=[1] And A.记录性质=[2]" & _
        " And A.病人ID=B.病人ID And Instr([3],A.记录状态)>0" & _
        IIF(mstrTime <> "", " And A.登记时间=[4]", "") & _
        " And A.开单部门ID=C.ID"
    End If
    If blnNOMoved Then
        StrSQL = Replace(StrSQL, mstrFeeTab, "H" & mstrFeeTab)
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strNO, mint记录性质, _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)))
    If rsTmp.EOF Then
        MsgBox "没有发现该单据。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mlng病人ID = 0 Then mlng病人ID = Nvl(rsTmp!病人ID, 0)
    
    cboNO.Text = strNO
    txt姓名.Text = Nvl(rsTmp!姓名)
    txt性别.Text = Nvl(rsTmp!性别)
    txt年龄.Text = Nvl(rsTmp!年龄)
    If Nvl(rsTmp!主页ID, 0) <> 0 Then
        txt床号.Text = Nvl(rsTmp!床号)
    End If
    
    '刘兴洪 问题:26953 日期:2009-12-25 15:23:48
    txt病人备注.Text = Nvl(rsTmp!病人备注)
    If mint病人来源 = 1 Then
        lbl住院号.Caption = "门诊号"
    Else
        lbl住院号.Caption = "住院号"
    End If
    txt住院号.Text = Nvl(rsTmp!标识号)
    
    txt费别.Text = Nvl(rsTmp!费别)
    txt担保人.Text = Nvl(rsTmp!担保人)
    txt担保额.Text = Format(Nvl(rsTmp!担保额), "0.00")
    txt付款方式.Text = Nvl(rsTmp!医疗付款方式)
    
    cbo开单科室.AddItem Nvl(rsTmp!开单部门)
    cbo开单科室.ItemData(cbo开单科室.NewIndex) = Nvl(rsTmp!开单部门ID, 0)
    cbo开单科室.ListIndex = cbo开单科室.NewIndex
    
    If Nvl(rsTmp!是否急诊, 0) = 1 Then
        chk急诊.value = 1: chk急诊.Visible = True
    End If
    
    chk加班.value = Nvl(rsTmp!加班标志, 0)
    Call LoadPatientBaby(cboBaby, rsTmp!病人ID, rsTmp!主页ID)
    Call zlControl.CboLocate(cboBaby, Nvl(rsTmp!婴儿费, 0), True)
    
    '开单人
    Call GetCboIndex(cbo开单人, Nvl(rsTmp!开单人))
    If cbo开单人.ListIndex = -1 And Not IsNull(rsTmp!开单人) Then
        cbo开单人.AddItem rsTmp!开单人
        cbo开单人.ListIndex = cbo开单人.NewIndex
    End If
    
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    
    If mint记录性质 = 2 Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!病人ID, IIF(mint病人来源 = 1, 0, mlng主页ID))
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "预交:" & Format(rsPatiMoney!预交余额, "0.00") & _
                "/费用:" & Format(rsPatiMoney!费用余额, gstrDec) & _
                "/剩余:" & Format(rsPatiMoney!预交余额 - rsPatiMoney!费用余额, "0.00")
        End If
    End If
    
    '------------------------------------------------------------------------------------
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        
        '读取单据中原始记录的费用ID
        strSQL1 = _
            " Select A.ID,A.序号,A.收费细目ID," & _
            " Nvl(A.付数,1)*A.数次" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & " as 原始数量" & _
            " From " & mstrFeeTab & " A,药品规格 B" & _
            " Where A.NO=[1] And A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
            " And A.收费细目ID=B.药品ID(+) And A.记录性质=[2]"
        
        '读取药品收发记录中的准退数
        strSQL2 = _
            " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & ") as 准退数量" & _
            " From 药品收发记录 A,药品规格 B" & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1" & _
            " And A.药品ID=B.药品ID(+) And A.审核人 is NULL" & _
            " And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
        
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
        StrSQL = "Select Nvl(价格父号,序号) From " & mstrFeeTab & _
            " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1]" & _
            " And Nvl(执行状态,0)<>1" & IIF(mlng医嘱ID <> 0, " And 医嘱序号+0=[8]", "")
        
        '如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
        If mint记录性质 = 2 Then
            If mint病人来源 = 2 Then intInsure = BillExistInsure(strNO)
            If intInsure <> 0 Then
                blnDo = Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, rsTmp!病人ID, intInsure)
            Else
                blnDo = gbytBillOpt = 2
            End If
            If blnDo Then
                StrSQL = StrSQL & " And Nvl(价格父号,序号) IN" & _
                    " (" & _
                    " Select Nvl(价格父号,序号) as 序号" & _
                    " From " & mstrFeeTab & _
                    " Where NO=[1] And 记录性质 IN(2,12)" & _
                    " Group by Nvl(价格父号,序号)" & _
                    " Having Sum(Nvl(结帐金额,0))=0" & _
                    " )"
            End If
        End If
        
        '因为是将要汇总求有剩余数量的，所以不能用直接用时间限制，用序号限制
        StrSQL = _
            " Select A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号," & _
            " C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIF(mbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & mstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg(A.数次" & IIF(mbln药房单位, "/Nvl(X." & mstr药房包装 & ",1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIF(mbln药房单位, "*Nvl(X." & mstr药房包装 & ",1)", "") & ") as 单价," & _
            " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From " & mstrFeeTab & " A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+)" & _
            " And A.收费细目ID=X.药品ID(+) And A.记录性质=[2]" & _
            " And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & StrSQL & ")" & _
            " Group by A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),C.编码,C.名称,A.收费细目ID,B.名称," & _
            " B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X." & mstr药房单位 & ",X." & mstr药房包装
            
        '最后计算结果
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        StrSQL = _
            " Select A.序号,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型,A.计算单位," & _
            " Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Avg(A.付数),1) as 准退付数," & _
            " Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
            " Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
            " A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,A.执行部门,A.附加标志" & _
            " From (" & StrSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.序号=B.序号 And B.ID=C.费用ID(+)" & _
            " Group by A.序号,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型," & _
            " A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行部门,A.附加标志" & _
            " Having Sum(A.付数*A.数次)<>0"
            
        StrSQL = _
            " Select A.序号,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格," & _
            "       A.费用类型,A.计算单位,A.准退付数 as 付数,A.准退数次 as 数次,A.单价," & _
            "       A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
            "       A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
            "       A.执行部门,A.附加标志" & _
            " From (" & StrSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[6]" & _
            "       And  A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
            " Order by A.序号"
    Else
        '读取单据原始内容
        intSign = IIF(mblnDelete, -1, 1) '数量,金额正负符号
        
        StrSQL = _
            "Select A.收费细目ID,A.收费类别,A.执行部门ID,Nvl(A.价格父号,A.序号) as 序号," & _
            " A.计算单位,A.付数,A.数次,A.标准单价,A.应收金额,A.实收金额,A.附加标志,A.费用类型" & _
            " From " & mstrFeeTab & " A Where A.记录性质=[2]" & _
            " And Instr([4],A.记录状态)>0 And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.登记时间=[5]", "")
        If blnNOMoved Then
            StrSQL = StrSQL & " Union ALL " & Replace(StrSQL, mstrFeeTab, "H" & mstrFeeTab)
        End If
        
        StrSQL = _
            " Select A.序号,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIF(mbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & mstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg([7]*A.数次" & IIF(mbln药房单位, "/Nvl(X." & mstr药房包装 & ",1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIF(mbln药房单位, "*Nvl(X." & mstr药房包装 & ",1)", "") & ") as 单价," & _
            " Sum([7]*A.应收金额) as 应收金额,Sum([7]*A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From (" & StrSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别" & _
            " And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
            " Group by A.序号,C.编码,C.名称,A.收费细目ID,B.名称,B.规格," & _
            " Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X." & mstr药房单位
            
        StrSQL = _
            " Select A.序号,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
            "       A.计算单位,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门,A.附加标志" & _
            " From (" & StrSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[6]" & _
            "       And  A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
            " Order by 序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strNO, mint记录性质, IIF(mint记录性质 = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), IIF(gbyt药品名称显示 = 1, 3, 1), intSign, mlng医嘱ID)
    If rsTmp.EOF Then
        If blnDelete Then
            MsgBox "单据中当前无可以操作的记录，可能单据中的项目已经全部执行。", vbInformation, gstrSysName
        Else
            MsgBox "单据中当前无可以操作的记录。", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        Bill.RowData(i) = rsTmp!序号 '用于记帐销帐
        
        Bill.TextMatrix(i, BillCol.类别) = rsTmp!类别
        Bill.TextMatrix(i, BillCol.项目) = rsTmp!名称
        Bill.TextMatrix(i, BillCol.商品名) = Nvl(rsTmp!商品名)
        Bill.TextMatrix(i, BillCol.规格) = Nvl(rsTmp!规格)
        Bill.TextMatrix(i, BillCol.单位) = Nvl(rsTmp!计算单位)
        Bill.TextMatrix(i, BillCol.付数) = Nvl(rsTmp!付数)
        Bill.TextMatrix(i, BillCol.数次) = FormatEx(rsTmp!数次, 5)
        Bill.TextMatrix(i, BillCol.单价) = Format(rsTmp!单价, gstrDecPrice)
        Bill.TextMatrix(i, BillCol.应收金额) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.实收金额) = Format(rsTmp!实收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.执行科室) = Nvl(rsTmp!执行部门)
        Bill.TextMatrix(i, BillCol.标志) = IIF(rsTmp!附加标志 = 1, "√", "")
        Bill.TextMatrix(i, BillCol.类型) = Nvl(rsTmp!费用类型)
        
        '设置销帐标志
        If Bill.TextMatrix(0, Bill.Cols - 1) = "删除" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "√"
        End If
        
        rsTmp.MoveNext
    Next
    '针对列编辑性质设置颜色
    Bill.SetColColor BillCol.类别, &HE7CFBA
    Bill.SetColColor BillCol.项目, &HE7CFBA
    Bill.SetColColor BillCol.数次, &HE7CFBA
    Bill.SetColColor BillCol.执行科室, &HE7CFBA
    Bill.SetColColor BillCol.付数, &HE0E0E0
    Bill.SetColColor BillCol.单价, &HE0E0E0
    Bill.SetColColor BillCol.标志, &HE0E0E0
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))

        '读取药品收发记录中的准退数
        strSQL1 = _
            " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & ") as 准退数量" & _
            " From 药品收发记录 A,药品规格 B" & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1" & _
            " And A.药品ID=B.药品ID(+) And A.审核人 is NULL" & _
            " And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
        
        '整张费用单据(明细到收入项目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        StrSQL = "Select Nvl(价格父号,序号) From " & mstrFeeTab & _
            " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1]" & _
            " And Nvl(执行状态,0)<>1" & IIF(mlng医嘱ID <> 0, " And 医嘱序号+0=[7]", "")
        If blnDo Then
            StrSQL = StrSQL & " And Nvl(价格父号,序号) IN" & _
                " (" & _
                " Select Nvl(价格父号,序号) as 序号" & _
                " From " & mstrFeeTab & _
                " Where NO=[1] And 记录性质 IN(2,12)" & _
                " Group by Nvl(价格父号,序号)" & _
                " Having Sum(Nvl(结帐金额,0))=0" & _
                " )"
        End If
        
        StrSQL = _
            " Select Sum(A.ID) as ID,A.序号,A.名称,A.收费类别," & _
            " Sum(A.数量) as 剩余数量,Sum(A.应收金额) as 剩余应收," & _
            " Sum(A.实收金额) as 剩余实收 From (" & _
            " Select Decode(A.记录状态,2,0,A.ID) as ID,A.序号,B.名称,A.收费类别," & _
            " Nvl(A.付数,1)*A.数次" & IIF(mbln药房单位, "/Nvl(X." & mstr药房包装 & ",1)", "") & " as 数量," & _
            " A.应收金额,A.实收金额" & _
            " From " & mstrFeeTab & " A,收入项目 B,药品规格 X" & _
            " Where A.记录性质=[2] And A.NO=[1]" & _
            " And A.收入项目ID=B.ID And Nvl(A.价格父号,A.序号) IN(" & StrSQL & ")" & _
            " And A.收费细目ID=X.药品ID(+)) A" & _
            " Group by A.序号,A.名称,A.收费类别" & _
            " Having Sum(A.数量)<>0"
                    
        '最后计算结果
        StrSQL = _
            " Select A.名称,Sum(A.剩余应收*(A.准退数量/A.剩余数量)) as 应收金额," & _
            " Sum(剩余实收*(A.准退数量/A.剩余数量)) as 实收金额 From (" & _
            " Select A.名称,A.剩余数量,A.剩余应收,A.剩余实收," & _
            " Decode(Instr(',4,5,6,7,',A.收费类别),0,A.剩余数量,Nvl(B.准退数量,A.剩余数量)) as 准退数量" & _
            " From (" & StrSQL & ") A,(" & strSQL1 & ") B" & _
            " Where A.ID=B.费用ID(+)" & _
            " ) A Group by A.名称"
    Else
        '读取单据原始内容
        intSign = IIF(mblnDelete, -1, 1) '数量,金额正负符号
        
        StrSQL = "Select A.收入项目ID,A.应收金额,A.实收金额 From " & mstrFeeTab & " A" & _
            " Where Instr([4],A.记录状态)>0 And A.记录性质=[2] And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.登记时间=[5]", "")
        If blnNOMoved Then
            StrSQL = StrSQL & " Union ALL " & Replace(StrSQL, mstrFeeTab, "H" & mstrFeeTab)
        End If
        
        StrSQL = _
            " Select B.名称,Sum([6]*A.应收金额) as 应收金额,Sum([6]*A.实收金额) as 实收金额 " & _
            " From (" & StrSQL & ") A,收入项目 B Where A.收入项目ID=B.ID Group By B.名称"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strNO, mint记录性质, IIF(mint记录性质 = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), intSign, mlng医嘱ID)
    If rsTmp.EOF Then Exit Function
    
    '刷新显示(收费要叠加)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    Call SetMoneyList
    
    For i = 1 To rsTmp.RecordCount
        mshMoney.TextMatrix(i, 0) = rsTmp!名称
        mshMoney.TextMatrix(i, 1) = Format(rsTmp!实收金额, gstrDec)
        curTotal = curTotal + rsTmp!实收金额
        cur应收Total = cur应收Total + rsTmp!应收金额
        rsTmp.MoveNext
    Next
    
    txt实收.Text = Format(curTotal, gstrDec)
    txt应收.Text = Format(cur应收Total, gstrDec)
    
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetShowCol()
'功能：付数列的控制(浏览时展开)
    mrsClass.Filter = "编码='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.付数) = 0
    ElseIf Bill.ColWidth(BillCol.付数) = 0 Then
        Bill.ColWidth(BillCol.付数) = 520
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
    '取其它中药的付数
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).付数
            Exit For
        End If
    Next
End Function

Private Function GetWorkUnit(ByVal lng药品ID As Long, ByVal str类别 As String) As Boolean
'功能：取所有可供选择的药房
    Dim StrSQL As String, bytDay As Byte
    Dim str药房 As String, lng开单科室ID As Long
    
    lng开单科室ID = mrsInfo!科室ID    '开单科室优先
    If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    
    If str类别 = "4" Then
        StrSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
            
        '以及SQL在卫材不支持存储库房设置之前用
'        strSQL = "Select A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象" & _
'            " From 部门表 A,部门性质说明 B" & _
'            " Where A.ID=B.部门ID And B.工作性质='发料部门' And B.服务对象 IN([1],3)" & _
'            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
'            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
'            " Order by B.服务对象,A.编码"
    Else
        '由药品材质确定药房性质
        Select Case str类别
            Case "5"
                str药房 = "西药房"
            Case "6"
                str药房 = "成药房"
            Case "7"
                str药房 = "中药房"
        End Select
        
        '药品从系统指定的储备药房中找
        If Not Check上班安排(True) Then
            StrSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
                " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[3]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            StrSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
                " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[3]" & _
                " Order by B.服务对象,C.编码"
        End If
    End If
    
    On Error GoTo errH
    'Set mrsWork = New ADODB.Recordset
    Set mrsWork = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mint病人来源, lng开单科室ID, lng药品ID, str药房, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load开单人(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    Dim lngOldID As Long
    
    cbo开单人.Clear
    
    '科室医生或护士
    StrSQL = _
        "   Select Distinct A.ID,B.部门ID,A.编号,A.姓名, Upper(A.简码) as 简码," & _
        "       C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
        "   From 人员表 A,部门人员 B,人员性质说明 C" & _
        "   Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        "       And C.人员性质 IN('医生','护士') And B.部门ID=[1]  " & _
        "   Order by 简码,人员性质 Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng科室ID)
    
    i = IIF(rsTmp.RecordCount = 0, 0, rsTmp.RecordCount - 1)
    ReDim marrDr(i)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If lngOldID <> rsTmp!ID Then
                cbo开单人.AddItem IIF(IsNull(rsTmp!简码), "", rsTmp!简码 & "-") & rsTmp!姓名
                cbo开单人.ItemData(cbo开单人.ListCount - 1) = rsTmp!部门ID
                marrDr(cbo开单人.ListCount - 1) = rsTmp!ID & "|" & rsTmp!部门ID & "|" & Nvl(rsTmp!编号) & "|" & rsTmp!姓名 & "|" & Nvl(rsTmp!简码) & "|" & rsTmp!职务 & "|" & Nvl(rsTmp!人员性质)
                
                If rsTmp!姓名 = mstr开嘱医生 Then cbo开单人.ListIndex = cbo开单人.NewIndex
                If lng科室ID = mlng病人科室id Then
                    '缺省为病人科室时,检查是否为住院医生
                    '问题:36862
                    If rsTmp!姓名 = mstr住院医生 Then cbo开单人.ListIndex = cbo开单人.NewIndex
                End If
                If rsTmp!ID = UserInfo.ID And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = cbo开单人.NewIndex
                lngOldID = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        
        If cbo开单人.ListCount > 0 Then ReDim Preserve marrDr(cbo开单人.ListCount - 1)
        If cbo开单人.ListCount = 1 And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CalcGridToTal(Optional bln应收 As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If bln应收 Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.应收金额
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.实收金额
                End If
            Next
        Next
    Else
        For i = 1 To Bill.Cols - 1
            If bln应收 Then
                If Bill.TextMatrix(0, i) = "应收金额" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "实收金额" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'功能：显示\隐藏销帐标志列
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "删除" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "删除"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = -1
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.类别) = GetOrigColWidth(BillCol.类别) - 120
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目) - 100
            Bill.ColWidth(BillCol.执行科室) = GetOrigColWidth(BillCol.执行科室) - 200
            
            Bill.ColWidth(BillCol.单价) = GetOrigColWidth(BillCol.单价) - 50
            Bill.ColWidth(BillCol.应收金额) = GetOrigColWidth(BillCol.应收金额) - 50
            Bill.ColWidth(BillCol.实收金额) = GetOrigColWidth(BillCol.实收金额) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "删除" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.类别) = GetOrigColWidth(BillCol.类别)
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目)
            Bill.ColWidth(BillCol.执行科室) = GetOrigColWidth(BillCol.执行科室)
            
            Bill.ColWidth(BillCol.单价) = GetOrigColWidth(BillCol.单价)
            Bill.ColWidth(BillCol.应收金额) = GetOrigColWidth(BillCol.应收金额)
            Bill.ColWidth(BillCol.实收金额) = GetOrigColWidth(BillCol.实收金额)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'功能：获取指定列的原始列宽
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub SetColNum(Optional intRow As Long = 1)
'功能：重新显示各行的行号
'参数：intRow=从该行开始
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.行) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True) As Integer
'功能：检查指定药品行的职务是否与当前医生的职务相匹配
'参数：tmpDetail=输入的项目,不传为所有行,blnCommon=是否正常的判断,否则为医保或公费病人的判断
'返回：不匹配的行,0为正确
'说明：职务：1=正高,2=副高,3=中级,4=助理/师级,5=员/士,9=待聘
    Dim i As Long, int职务A As Integer, int职务B As Integer
    Dim strTmp As String
    
    strTmp = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    
    If cbo开单人.ListIndex = -1 Then Exit Function
    If cbo开单人.ListIndex <= UBound(marrDr) Then
        If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 5 Then
            int职务A = Val(Split(marrDr(cbo开单人.ListIndex), "|")(5))
        End If
    End If
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                If Not blnCommon Then
                    int职务B = Val(Right(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            strTmp = "对医保或公费病人,第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            strTmp = "对医保或公费病人,第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    int职务B = Val(Left(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.类别) = 0 Then Exit Function
        If Not blnCommon Then
            int职务B = Val(Right(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "对医保或公费病人,药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "对医保或公费病人,药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        Else
            int职务B = Val(Left(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
'功能：判断指定药品在单据中是否已经存在
'参数：objDetail=项目,intRow=要判断的行
'说明：时价或分批药品在同一药房禁止重复输入(这里仅提示,保存时禁止)
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.分批 Or mobjBill.Details(i).Detail.变价) _
                    And (objDetail.分批 Or objDetail.变价) Then
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价药品,重复输入时必须保证它们的发料部门不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该药品为分批或时价药品,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function Check费用类型(Optional intRow As Integer) As Boolean
'功能：根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    Dim StrSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim bln医保 As Boolean, bln公费 As Boolean
    Check费用类型 = True
    
    '无法检查
    If txt付款方式.Tag = "" Then Exit Function
    '45605
    '只检查医保病人和公费病人
    If zlIsCheckMedicinePayMode(txt付款方式.Text, bln医保, bln公费) = False Then Exit Function
    '确定病人类型
    bytType = IIF(bln医保, 1, 2) ' Val(txt付款方式.Tag)
    '读取检查数据
    If bytType = 1 Then
        StrSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
    Else
        StrSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)

    If rsTmp.EOF Then Exit Function
    
    If intRow > 0 Then
        If mobjBill.Details(intRow).Detail.类型 = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型未设置！", vbInformation, gstrSysName
            Check费用类型 = False
        Else
            rsTmp.Filter = "名称='" & mobjBill.Details(intRow).Detail.类型 & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型为""" & _
                    mobjBill.Details(intRow).Detail.类型 & """,不是" & _
                    IIF(bytType = 1, "医保", "公费") & "费用类型！", vbInformation, gstrSysName
                Check费用类型 = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).Detail.类型 = "" Then
                If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """的费用类型未设置！" & vbCrLf & "确实要保存单据吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check费用类型 = False: Exit For
                End If
            Else
                rsTmp.Filter = "名称='" & mobjBill.Details(i).Detail.类型 & "'"
                If rsTmp.EOF Then
                    If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """的费用类型为""" & _
                        mobjBill.Details(i).Detail.类型 & """,不是" & _
                        IIF(bytType = 1, "医保", "公费") & "费用类型！" & vbCrLf & "确实要保存单据吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check费用类型 = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReCalcInsure()
'功能：修改单据时,重新计算统筹金额及更新相关信息
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    
    If Not IsNull(mrsInfo!险类) Then
        For i = 1 To mobjBill.Details.Count
            For j = 1 To mobjBill.Details(i).InComes.Count
                dblAllTime = mobjBill.Details(i).付数 * mobjBill.Details(i).数次
                If InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                    If mbln药房单位 Then dblAllTime = dblAllTime * mobjBill.Details(i).Detail.药房包装
                End If
                strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(i).收费细目ID, mobjBill.Details(i).InComes(j).实收金额, False, mrsInfo!险类, _
                    mobjBill.Details(i).摘要 & "||" & dblAllTime)
                    
                If strInfo <> "" Then
                    mobjBill.Details(i).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(i).保险大类ID = Val(Split(strInfo, ";")(1))
                    mobjBill.Details(i).InComes(j).统筹金额 = Val(Split(strInfo, ";")(2))
                    mobjBill.Details(i).保险编码 = CStr(Split(strInfo, ";")(3))
                    
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).摘要 = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.类型 = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            Next
        Next
    End If
End Sub

Private Function HaveStopClass() As Integer
'功能：判断当前单据中是否有护士禁止输入的内容
    Dim i As Long, str性质 As String
    
    If cbo开单人.ListIndex <> -1 Then
        If cbo开单人.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 6 Then
                str性质 = Split(marrDr(cbo开单人.ListIndex), "|")(6)
            End If
        End If
    End If
    
    For i = 1 To mobjBill.Details.Count
        If str性质 = "护士" And InStr(",E,M,4,", mobjBill.Details(i).收费类别) = 0 Then
            HaveStopClass = i: Exit Function
        End If
    Next
End Function

Private Function Check执行科室() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).执行部门ID = 0 Or Bill.TextMatrix(i, BillCol.执行科室) = "" Then
            Check执行科室 = i: Exit Function
        End If
    Next
End Function

Public Sub InitLocPar()
'功能：初始化费用本机参数
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mblnPay = Val(zlDatabase.GetPara("中药输入付数", glngSys, p医嘱附费管理)) <> 0
    mblnTime = Val(zlDatabase.GetPara("变价输入数次", glngSys, p医嘱附费管理)) <> 0
    mbln其它药房 = Val(zlDatabase.GetPara("显示其它药房库存", glngSys, p医嘱附费管理)) = 1
    mbln其它药库 = Val(zlDatabase.GetPara("显示其它药库库存", glngSys, p医嘱附费管理)) = 1
    mstr收费类别 = zlDatabase.GetPara("收费类别", glngSys, p医嘱附费管理, "")
    
    '药品单位
    mbln药房单位 = Val(zlDatabase.GetPara("药品单位", glngSys, p医嘱附费管理)) <> 0
    If mint病人来源 = 1 Then
        mstr药房单位 = "门诊单位": mstr药房包装 = "门诊包装"
    Else
        mstr药房单位 = "住院单位": mstr药房包装 = "住院包装"
    End If
    mbytSendMateria = Val(zlDatabase.GetPara("记帐后发药", glngSys, p医嘱附费管理))
    mbyt缺省科室 = Val(zlDatabase.GetPara("补费缺省科室", glngSys, p医嘱附费管理))
    '缺省药房
    mlng西药房 = Val(zlDatabase.GetPara(IIF(mint病人来源 = 2, "住院", "门诊") & "缺省西药房", glngSys, p医嘱附费管理))
    mlng成药房 = Val(zlDatabase.GetPara(IIF(mint病人来源 = 2, "住院", "门诊") & "缺省成药房", glngSys, p医嘱附费管理))
    mlng中药房 = Val(zlDatabase.GetPara(IIF(mint病人来源 = 2, "住院", "门诊") & "缺省中药房", glngSys, p医嘱附费管理))
    mlng发料部门 = Val(zlDatabase.GetPara(IIF(mint病人来源 = 2, "住院", "门诊") & "缺省发料部门", glngSys, p医嘱附费管理))
End Sub

Public Function zlCheck北京医保(ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对北京医保的一些检查
    '入参:intInsuer-险类
    '出参:
    '返回:检查成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-07 10:25:04
    '问题:27278
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    If intInsure = 0 Then zlCheck北京医保 = True: Exit Function
    
    err = 0: On Error GoTo Errhand:
    '刘兴洪:???
    'mint病人来源:1-门诊病人,2-住院病人
    'mint记录性质:1-收费(划价),2-记帐(门/住)
    'mbytInState :0-执行,1-查阅,2-调整(不支持),3-删费
    
    '只有划价才支持检查
    If (mint病人来源 = 2 Or mint记录性质 = 2) And mbytInState <> 0 Or MCPAR.医生确定处方类型 = False Then
        zlCheck北京医保 = True: Exit Function
    End If
    
    'showmsgbox
    '参数：strCaption=消息窗体标题
    '      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
    '      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
    '              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
    '              每个按钮文字最多支持4个汉字
    '      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
    '返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    strTemp = zlCommFun.ShowMsgBox("处方类型", "请确定当前医保病人本次要发送的药品处方的类型。", "!医保内(&A),医保外(&B),?取消(&C)", Me)
    If strTemp = "" Then Exit Function
    '如果是补门诊收费划价单，且是医保病人，则当医保参数”support医生确定处方类型”有效时，保存时提示该单据是”医保内，医保外”，如果是医保内费用记录摘要中存放1，医保外存放2。
    strTemp = IIF(strTemp = "医保内", 1, 2)
    For Each mobjBillDetail In mobjBill.Details
        mobjBillDetail.摘要 = strTemp
    Next
    zlCheck北京医保 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlCheck定价零价格对码(ByVal lng收费细目ID As Long, bln定价 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医保对码(定价为零的)
    '入参:
    '出参:
    '返回:如果定价项目为零,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-12 11:22:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
   '刘兴洪 问题:27286 定价的价格为零的不进行检查对码 日期:2010-01-07 15:13:45
   Dim StrSQL As String, rs价格 As ADODB.Recordset, dbl价格 As Double
    err = 0: On Error GoTo Errhand:
   zlCheck定价零价格对码 = False
    If bln定价 Then
        StrSQL = _
        " Select  B.现价 " & _
        " From 收费价目 B " & _
        " Where   ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        "       And B.收费细目ID=[1]"
        Set rs价格 = zlDatabase.OpenSQLRecord(StrSQL, "获取当前价格", lng收费细目ID)
        If rs价格.EOF = False Then
            dbl价格 = Val(Nvl(rs价格!现价))
        Else
            dbl价格 = 0
        End If
        If dbl价格 = 0 Then zlCheck定价零价格对码 = True: Exit Function
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zl获取中药形态(Optional ByVal lngRow As Long = -1, Optional blnOnly中成药 As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据是否录入了中草药的
    '入参:blnOnly中成药-仅判断是否有中成药(对配方时判断有效):原因是中划药在配方中已经存在,就不需要检查
    '     lngRow-当前操作的行
    '出参:
    '返回:录入了中草药的,则返回免煎属性(1-免煎,0-需要煎),否则返回-1 表示还没有录入免煎项目
    '编制:刘兴洪
    '日期:2010-02-02 11:44:17
    '问题:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl获取中药形态 = -1
    '如果未指定页,则用当前页
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIF(blnOnly中成药, ",6,", ",6,7,")
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).收费类别 & ",") > 0 And .Item(i).收费细目ID <> 0 And i <> lngRow Then
                zl获取中药形态 = .Item(i).Detail.中药形态
                Exit Function
            End If
        Next
    End With
End Function
Public Function Get待发药清单(strNO As String, strTime As String) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据费用单据号,登记时间,获取待发药品清单
    '入参：strNO-单据号
    '          strTime-登记时间
    '出参：
    '返回：待发药清单
    '编制：刘兴洪
    '日期：2010-03-19 18:59:27
    '说明：普通发药时为病人科室，急诊、医技则为开单科室。
    '------------------------------------------------------------------------------------------------------------------------
    Dim StrSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    StrSQL = _
        " Select A.ID,A.库房ID,A.对方部门ID" & _
        " From 药品收发记录 A,住院费用记录 B" & _
        " Where A.NO=[1] And A.单据=[2] And Mod(A.记录状态,3)=1 And A.审核人 is NULL" & _
        " And A.NO=B.NO And A.费用ID=B.ID And B.记录状态<>0 And B.登记时间+0=[3]" & _
        " Order by A.药品ID"
    If strTime <> "" Then
        Set Get待发药清单 = zlDatabase.OpenSQLRecord(StrSQL, "mdlInExse", strNO, 9, CDate(strTime))
    Else
        Set Get待发药清单 = zlDatabase.OpenSQLRecord(StrSQL, "mdlInExse", strNO, 9)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub SetDetailtStock(ByVal lng执行科室ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置明细的库存数据
    '编制：刘兴洪
    '日期：2010-07-12 14:27:51
    '说明：
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim str药房IDs As String, dblStock As Double
    
    '获取库存
    '不处理非药品或卫材
    If InStr(1, "5,6,7,4", objDetail.类别) = 0 Then Exit Sub
    If objDetail.类别 = "4" And objDetail.跟踪在用 = False Then Exit Sub
    If objDetail.类别 = "4" Then
        '卫材
        dblStock = GetStock(objDetail.ID, lng执行科室ID)
        objDetail.库存 = dblStock
        Exit Sub
    End If
    
    dblStock = GetStock(objDetail.ID, lng执行科室ID)
    If mbln药房单位 Then
        dblStock = dblStock / objDetail.药房包装
    End If
    objDetail.库存 = dblStock  '记录当前行药品库存
End Sub

Private Sub cmdSelWholeSet_Click()
    '选成套项目
    Dim rsSel As ADODB.Recordset, lng病人ID As Long, lng开单部门ID As Long
    Dim tmpBill As New ExpenseBill, byt婴儿费 As Byte, curDate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim lng病人科室ID As Long, str费别 As String, intInsure As Integer
    intInsure = 0
    If mobjBill Is Nothing Then
        If mrsInfo Is Nothing Then
            MsgBox "请先选择病人,请检查!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        intInsure = Val(Nvl(mrsInfo!险类))
        If cbo开单科室.ListIndex < 0 Then
            lng开单部门ID = 0
        Else
            lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        If cboBaby.ListIndex < 0 Then
            byt婴儿费 = 0
        Else
            byt婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
        End If
        lng病人科室ID = mlng病人科室id: str费别 = Nvl(mrsInfo!费别)
    Else
        lng病人ID = mobjBill.病人ID: lng开单部门ID = mobjBill.开单部门ID: byt婴儿费 = mobjBill.婴儿费
        lng病人科室ID = mobjBill.科室ID: str费别 = mobjBill.费别
        If mrsInfo Is Nothing Then
           If mrsInfo.State = 1 Then intInsure = Val(Nvl(mrsInfo!险类))
        End If
    End If
    
    If lng病人ID = 0 Then
        MsgBox "请先选择病人,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If frmWholeSelect.ShowSelect(Me, p医嘱附费管理, mstrPrivs, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    err = 0: On Error GoTo Errhand:
    Screen.MousePointer = 11
    
    Set tmpBill = ImportWholeSet(Me, intInsure, rsSel, lng病人ID, lng开单部门ID, byt婴儿费, IIF(mint病人来源 = 2 And mint记录性质 = 2, 2, 0), chk加班.value = 1, _
        0, mint病人来源, UserInfo.姓名, NeedName(cbo开单人.Text))
    '处理数据
    '清除导入的病人信息
    '问题:37500
    With tmpBill
        .病人ID = mobjBill.病人ID
        .主页ID = mobjBill.主页ID
        .病区ID = mobjBill.病区ID
        .科室ID = mobjBill.科室ID
        .床号 = mobjBill.床号
        .标识号 = mobjBill.标识号
        .姓名 = mobjBill.姓名
        .性别 = mobjBill.性别
        .年龄 = mobjBill.年龄
        .费别 = mobjBill.费别
    End With
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    Dim bln中药 As Boolean
    bln中药 = False
    With mobjBill
        For i = 1 To .Details.Count - 1
            If .Details(i).收费类别 = "7" Then
                bln中药 = True
                Exit For
            End If
            Exit For
        Next
    End With
    curDate = zlDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.登记时间 = curDate
    mobjBill.操作员编号 = UserInfo.编号
    mobjBill.操作员姓名 = UserInfo.姓名
    mobjBill.加班标志 = chk加班.value
    If mobjBill.费别 = "" Then mobjBill.费别 = str费别
    If mobjBill.科室ID = 0 Then mobjBill.科室ID = lng病人科室ID
    mobjBill.婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
    txtDate.Text = Format(curDate, "yyyy-MM-dd HH:mm:ss")
    Bill.Redraw = False
    Bill.ClearBill
    Bill.Rows = mobjBill.Details.Count + 1
    
   ' Call InitBillColumnColor
    '记帐分类报警
    mstrWarn = ""
        
   ' Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mobjBill.开单人, mobjBill.开单部门ID)
        
    '等上面的读病人后确定费别后,再计算价格
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    With Bill
        For i = 1 To .Rows - 1
            .TextMatrix(i, BillCol.行) = i
        Next
    End With
    
    Bill.Redraw = True
    '刷新病人费用信息
    If mrsInfo.State = 1 Then
        '刷新病人预交款信息
        curTotal = GetBillTotal(mobjBill)
        Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, IIF(mint病人来源 = 1, 0, mlng主页ID))
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!预交余额
            cmdCancel.Tag = rsTmp!费用余额
            txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
        End If
    End If
    '重新计算统筹金额
    Call ReCalcInsure
    '针对列编辑性质设置颜色
    Bill.SetColColor BillCol.类别, &HE7CFBA
    Bill.SetColColor BillCol.项目, &HE7CFBA
    Bill.SetColColor BillCol.数次, &HE7CFBA
    Bill.SetColColor BillCol.执行科室, &HE7CFBA
    Bill.SetColColor BillCol.付数, &HE0E0E0
    Bill.SetColColor BillCol.单价, &HE0E0E0
    Bill.SetColColor BillCol.标志, &HE0E0E0
    Screen.MousePointer = 0
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lng执行科室ID As Long
    Dim rsTemp As ADODB.Recordset, dbl数次 As Double, dbl价格 As Double
    Dim StrSQL As String, blnNOMoved As Boolean
    '保存为存套收费项目
    '问题:27327
    err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    If mint记录性质 = 1 Or (mint记录性质 = 2 And mint病人来源 = 1) Then
        blnNOMoved = zlDatabase.NOMoved("门诊费用记录", mstrInNO, "记录性质=", mint记录性质)
    Else
        blnNOMoved = zlDatabase.NOMoved("住院费用记录", mstrInNO, "记录性质=", mint记录性质)
    End If

    
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:序号,父号,收费细目ID,数量,单价,执行科室|序号,父号,收费细目ID,数量,单价,执行科室|…
    err = 0: On Error GoTo Errhand:
   If mbytInState = 1 Then
        '查看
        
         StrSQL = _
        " Select Nvl(A.价格父号,A.序号) as 序号,A.收费类别,A.从属父号,A.收费细目ID,A.执行部门ID," & _
        "       　   Avg(Nvl(A.付数,1)) as 付数, Avg(A.数次) 数次, Sum(A.标准单价) as 单价,B.执行科室, B.是否变价,M.跟踪在用" & _
        " From " & IIF(blnNOMoved, "H" & mstrFeeTab, mstrFeeTab & " A") & ",收费项目目录 B,材料特性 M" & _
        " Where  A.记录状态  IN(0,1,3)  And A.NO=[1]  And A.记录性质=[2] " & _
        "               And a.收费细目ID=b.ID And a.收费细目ID=M.材料ID(+) " & _
                        IIF(mstrTime <> "", " And A.登记时间=[3]", "") & _
        "  Group by Nvl(A.价格父号,A.序号),A.收费类别,A.收费细目ID,A.从属父号,A.执行部门id,B.执行科室,B.是否变价,M.跟踪在用" & _
        " Order by 序号"
        If mstrTime <> "" Then
            Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mstrInNO, mint记录性质, CDate(mstrTime))
        Else
            Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mstrInNO, mint记录性质)
        End If
        With rsTemp
            Do While Not .EOF
                 '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
                If InStr(1, ",4,5,6,7,", "," & Nvl(!收费类别)) > 0 Then
                    lng执行科室ID = 0
                ElseIf InStr(1, ",0,4", Val(Nvl(!执行科室))) > 0 Then
                    lng执行科室ID = Val(Nvl(!执行部门ID))
                Else
                    lng执行科室ID = 0
                End If
                dbl价格 = 0
                If Val(Nvl(!是否变价)) = 1 Then
                    If InStr(1, "5,6,7", Nvl(!收费类别)) > 0 Or (Nvl(!收费类别) = "4" And Val(Nvl(!跟踪在用)) = 1) Then
                        '药品,跟踪卫材因为有缺省价格,所以不处理(通过库存计算)
                        dbl价格 = 0
                    Else
                        dbl价格 = Val(Nvl(!单价))
                    End If
                End If
                strItems = strItems & "|" & Val(Nvl(!序号)) & "," & Val(Nvl(!从属父号)) & "," & Val(Nvl(!收费细目ID)) & "," & Val(Nvl(!付数)) & "," & Val(Nvl(!数次)) & "," & dbl价格 & "," & lng执行科室ID
                .MoveNext
            Loop
        End With
         If strItems = "" Then
            MsgBox "单据未输入任何信息,不能保存为成套收费项目,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Sub
        End If
        strItems = Mid(strItems, 2)
   Else
        With mobjBill
            strItems = ""
            For i = 1 To .Details.Count
                 '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
                If InStr(1, ",4,5,6,7,", "," & .Details(i).Detail.类别) > 0 Then
                    lng执行科室ID = 0
                ElseIf InStr(1, ",0,4", .Details(i).Detail.执行科室) > 0 Then
                    lng执行科室ID = .Details(i).执行部门ID
                Else
                    lng执行科室ID = 0
                End If
                '问题:52349
                dbl数次 = .Details(i).数次
                dbl价格 = IIF(.Details(i).Detail.变价, .Details(i).InComes(1).标准单价, 0)
                If InStr(",5,6,7,", .Details(i).收费类别) > 0 And mbln药房单位 Then
                    dbl数次 = Format(dbl数次 * .Details(i).Detail.药房包装, "0.00000")
                    dbl价格 = Format(dbl价格 / .Details(i).Detail.药房包装, gstrDecPrice)
                End If
                strItems = strItems & "|" & .Details(i).序号 & "," & .Details(i).从属父号 & "," & .Details(i).收费细目ID & "," & .Details(i).付数 & "," & dbl数次 & "," & dbl价格 & "," & lng执行科室ID
             Next
             
             If strItems = "" Then
                MsgBox "单据未输入任何信息,不能保存为成套收费项目,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
            strItems = Mid(strItems, 2)
        End With
    End If
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, 1150, mstrPrivs, strItems)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Function ImportWholeSet(frmParent As Object, ByVal intInsure As Integer, rsSel As ADODB.Recordset, Optional lng病人ID As Long = 0, _
     Optional lng开单部门ID As Long = 0, Optional byt婴儿费 As Byte, _
     Optional int门诊标志 As Integer, Optional bln加班加价 As Boolean = False, _
     Optional ByVal lngUnitID As Long, Optional int范围 As Integer, _
     Optional str划价人 As String = "", Optional str开单人 As String = "", _
     Optional lng主页ID As Long = 0) As ExpenseBill
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取费用单据到单据对象中
    '入参:rsSel-选中的成套项目
    '       lngUnitID    当前操作病区ID
    '      int范围=1.门诊,2-住院
    '      intInsure:险类
    '出参:
    '返回:存放单据信息的单据对象
    '编制:刘兴洪
    '日期:2010-09-02 16:17:54
    '说明:因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
    '       不包含已停用收费细目
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue(0 To 10) As String, strSubItem As String, str收费细目ID As String, j As Long
    Dim rsItems As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng病人科室ID As Long, str摘要 As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset, rsPrice As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim lngDoUnit As Long
    Dim i As Long, intCurNo As Integer
    Dim int序号 As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, StrSQL As String, str药房IDs As String, str停用项目序号 As String, strPrivs As String
    Dim curModiMoney As Currency
    Dim strWhere As String
    Dim dblAllTime As Double, dblCurTime As Double, dbl加班加价率 As Double, dblPriceSingle As Double, lngLastPati As Long
    Dim colSerial As New Collection, dblPrice As Double
    Dim bytType As Byte '0-门诊;1-住院;2-门诊或住院
    Dim strTable  As String
    
    On Error GoTo errH
    With rsSel
        str收费细目ID = "": j = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Len(str收费细目ID) > 1990 And j <= 10 Then
                strValue(j) = Mid(str收费细目ID, 2)
                strSubItem = strSubItem & " Union ALL " & _
                " Select Column_Value as 收费细目ID From Table(f_Num2List([" & j + 1 & "])) B "
                str收费细目ID = "": j = j + 1
            End If
            str收费细目ID = str收费细目ID & "," & Val(Nvl(!收费细目ID))
            .MoveNext
        Loop
    End With
    
    If str收费细目ID <> "" Then
        If j > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From 收费项目目录 Where id in (" & Mid(str收费细目ID, 2) & ")"
        Else
            strValue(j) = Mid(str收费细目ID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as 收费细目ID From Table(f_Num2List([" & j + 1 & "])) B "
        End If
    End If
    
    gstrSQL = "" & _
       "   Select A.主项id, A.从项id, A.固有从属, A.从项数次 " & _
       "   From 收费从属项目 A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.主项id = D.收费细目id "
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    strSubItem = Mid(strSubItem, 11)
    strTable = " Select [13] as 病人ID,收费细目ID From (" & strSubItem & ")"
    
    gstrSQL = "" & _
    " Select  X.药品ID,W.材料ID,W.跟踪在用," & _
    "       nvl(G.费别,F.费别) as 费别,NVL( G.姓名,F.姓名) 姓名,NVL(G.性别,F.性别) 性别,NVL(G.年龄,F.年龄) 年龄,F.担保额," & _
    "       G.出院病床 as 床号,F.住院号 as 标识号,F.病人ID,G.主页ID,G.当前病区ID as 病人病区ID,G.出院科室ID as 病人科室ID," & _
    "       G.病人性质,B.类别 as 收费类别,A.收费细目ID," & _
    "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(H.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
    "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象, B.费用类型  费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
    "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
    "       Decode(B.类别,'4',1,X." & mstr药房包装 & ") as 药房包装,Decode(B.类别,'4',B.计算单位,X." & mstr药房单位 & ") as 药房单位," & _
    "       Decode(b.类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,B.录入限量, " & _
    "       M1.编码 as 诊疗编码,M1.名称 as 诊疗名称,X.中药形态,x.剂量系数,M1.计算单位 as 剂量单位" & _
    "   From  (" & strTable & ") A ,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,病人信息 F, " & _
    "          病案主页 G ,收费项目别名 H,收费项目别名 E1,材料特性 W,药品规格 X,诊疗项目目录 M1" & _
    " Where  A.收费细目ID=D.收费细目ID And A.收费细目ID=B.ID " & _
    "       And b.类别=C.编码 And A.收费细目ID=X.药品ID(+) and X.药名ID=M1.ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
    "       And A.收费细目ID=H.收费细目ID(+) And H.码类(+)=1 And H.性质(+)=[12]" & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "       And A.病人ID=F.病人ID(+) And F.病人ID=G.病人ID(+)  And F.主页ID=G.主页ID(+)" & _
    "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) " & vbNewLine & _
    "       And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) "
    
    If Not gbln分离发药 Then
        gstrSQL = "Select * From (" & gstrSQL & ")"
    Else
        '分离发药时排开时价和分批药品或卫材
        gstrSQL = "Select * From (" & gstrSQL & ") Where Not( Instr(',5,6,7,',收费类别)>0 And (分批=1 Or 是否变价=1))"
    End If
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, "mdlExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10), IIF(gbyt药品名称显示 = 1, 3, 1), lng病人ID)
    '没有记录就是空单子
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    
    With rsSel
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
NextRecord: Do While Not .EOF
            '检查收费项目是否停用或服务于门诊病人
            '主项停用时,不导从项
            rsItems.Filter = "收费细目ID=" & Val(Nvl(!收费细目ID))
            If rsItems.EOF Then '未找到.不加入
                 .MoveNext
                GoTo NextRecord:
            End If
            If InStr(",5,6,7,", rsItems!收费类别) = 0 Then
                If InStr(1, str停用项目序号 & ",", "," & !从属父号 & ",") > 0 Then
                    .MoveNext
                    GoTo NextRecord
                Else
                    If Not CheckFeeItemAvailable(!收费细目ID, 2) Then
                        str停用项目序号 = str停用项目序号 & "," & !序号
                        MsgBox "成套收费项目中的第" & !序号 & "行收费项目:" & rsItems!名称 & "" & vbCrLf & _
                            "已停用或不再服务于病人,将不会被导入." & IIF(IsNull(!从属父号), "如果有从属项目,也不会被导入.", ""), vbInformation, gstrSysName
                        .MoveNext
                        GoTo NextRecord
                    End If
                End If
            End If
            
            If i = 1 Then
                objBill.NO = ""
                objBill.病人ID = Val(Nvl(rsItems!病人ID))
                objBill.主页ID = Val(Nvl(rsItems!主页ID))
                objBill.病区ID = Val(Nvl(rsItems!病人病区ID))
                objBill.科室ID = Val(Nvl(rsItems!病人科室id))
                objBill.姓名 = Nvl(rsItems!姓名)
                objBill.性别 = Nvl(rsItems!性别)
                objBill.年龄 = Nvl(rsItems!年龄)
                objBill.标识号 = Val(Nvl(rsItems!标识号))
                objBill.床号 = "" & rsItems!床号
                objBill.费别 = Nvl(rsItems!费别)
                objBill.门诊标志 = int门诊标志
                objBill.加班标志 = IIF(bln加班加价, 1, 0)
                objBill.婴儿费 = byt婴儿费
                objBill.开单部门ID = lng开单部门ID
                objBill.划价人 = str划价人
                objBill.开单人 = str开单人
                objBill.操作员编号 = UserInfo.编号
                objBill.操作员姓名 = UserInfo.姓名
                objBill.发生时间 = zlDatabase.Currentdate   ' !发生时间
                objBill.登记时间 = zlDatabase.Currentdate
                objBill.多病人单 = 0
            End If
            '处理收费细目=====================================================
            Set objBillDetail = New BillDetail
            Set objBillDetail.Detail = New Detail
                
            '处理序号和从属父号
            intCurNo = intCurNo + 1
            objBillDetail.序号 = intCurNo
            colSerial.Add Array(Val(Nvl(!收费细目ID)), intCurNo), "_" & !序号  '记录原序号现在的行号
            objBillDetail.从属父号 = Nvl(!从属父号, 0) '因为可能排序乱了,先记录原来的,后面再处理
            objBillDetail.收费类别 = Nvl(rsItems!收费类别)
            objBillDetail.收费细目ID = Val(Nvl(!收费细目ID))
            objBillDetail.计算单位 = Nvl(rsItems!计算单位)
            objBillDetail.付数 = IIF(Val(Nvl(!付数)) = 0, 1, Val(Nvl(!付数)))
            
            If InStr(",5,6,7,", rsItems!收费类别) > 0 And mbln药房单位 Then
                objBillDetail.数次 = Nvl(!数量, 0) / Nvl(rsItems!药房包装, 1)
            Else
                objBillDetail.数次 = Nvl(!数量, 0)
            End If
            objBillDetail.发药窗口 = ""
            
            objBillDetail.附加标志 = 0 ' IIf(IsNull(!附加标志), 0, !附加标志)
            objBillDetail.摘要 = "" ' IIf(IsNull(!摘要), "", !摘要)
            '卫材和药品部分
            '卫材执行科室缺省为病人病区,如果本地指定了,则为指定科室
            If objBillDetail.收费类别 = "4" Then
                lngDoUnit = IIF(mlng发料部门 > 0, mlng发料部门, objBill.病区ID)
                If lngDoUnit = 0 Then lngDoUnit = lng开单部门ID
            ElseIf InStr(1, ",5,6,7,", "," & objBillDetail.收费类别 & ",") > 0 Then
                '检查是否有缺省药房,存在缺省的,则取缺省药房,否则取上一条药房
                '问题:36736
                Select Case objBillDetail.收费类别
                    Case "5"
                        If mlng西药房 > 0 Then lngDoUnit = mlng西药房
                    Case "6"
                        If mlng成药房 > 0 Then lngDoUnit = mlng成药房
                    Case "7"
                        If mlng中药房 > 0 Then lngDoUnit = mlng中药房
                End Select
            Else
                If Val(Nvl(!执行科室ID)) <> 0 Then lngDoUnit = Val(Nvl(!执行科室ID))
            End If
            
            '病人科室ID
            lng病人科室ID = objBill.科室ID
            If lng病人科室ID = 0 Then lng病人科室ID = lng开单部门ID
            objBillDetail.Detail.执行科室 = IIF(IsNull(rsItems!执行科室), 0, rsItems!执行科室)
            objBillDetail.执行部门ID = Val(Nvl(!执行科室ID))
            
           lngDoUnit = Get收费执行科室ID(Val(Nvl(rsItems!病人ID)), Val(Nvl(rsItems!主页ID)), objBillDetail.收费类别, objBillDetail.收费细目ID, _
                        objBillDetail.Detail.执行科室, lng病人科室ID, lng开单部门ID, int范围, lngDoUnit, 1, 1, , objBillDetail.执行部门ID)          '0-医嘱程序调用,1-附费程序调用
            objBillDetail.执行部门ID = lngDoUnit

            If InStr(",5,6,7,", rsItems!收费类别) > 0 And gbln分离发药 Then
                objBillDetail.执行部门ID = 0
            End If
            objBillDetail.Detail.ID = !收费细目ID
            objBillDetail.Detail.编码 = Nvl(rsItems!编码)
            objBillDetail.Detail.变价 = (Val(Nvl(rsItems!是否变价)) = 1)
            objBillDetail.Detail.从项数次 = 0
            objBillDetail.Detail.固有从属 = 0
            If objBillDetail.从属父号 <> 0 Then
                'A.主项id, A.从项id, A.固有从属, A.从项数次 "
                rsOthers.Filter = "主项ID=" & colSerial("_" & !从属父号)(0) & " And 从项ID=" & objBillDetail.收费细目ID
                If Not rsOthers.EOF Then
                    objBillDetail.Detail.从项数次 = Val(Nvl(rsOthers!从项数次))
                    objBillDetail.Detail.固有从属 = Val(Nvl(rsOthers!固有从属))
                End If
            End If
            
            objBillDetail.Detail.规格 = Nvl(rsItems!规格)
            objBillDetail.Detail.计算单位 = Nvl(rsItems!计算单位)
            
            objBillDetail.Detail.药房单位 = Nvl(rsItems!药房单位)
            objBillDetail.Detail.药房包装 = Val(Nvl(rsItems!药房包装))
            
            objBillDetail.Detail.加班加价 = 0 ' (IIf(IsNull(!加班加价), 0, !加班加价) = 1)
            objBillDetail.Detail.类别 = Nvl(rsItems!类别)
            objBillDetail.Detail.类别名称 = Nvl(rsItems!类别名称)
            objBillDetail.Detail.名称 = Nvl(rsItems!名称)
            objBillDetail.Detail.商品名 = Nvl(rsItems!商品名)
            objBillDetail.Detail.屏蔽费别 = (Val(Nvl(rsItems!屏蔽费别)) = 1)
            objBillDetail.Detail.说明 = ""
            objBillDetail.Detail.服务对象 = IIF(IsNull(rsItems!服务对象), 0, rsItems!服务对象)
            objBillDetail.Detail.类型 = IIF(IsNull(rsItems!费用类型), "", rsItems!费用类型)
            
            If InStr(",5,6,7,", rsItems!收费类别) > 0 Then
                objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                objBillDetail.Detail.处方限量 = Get处方限量(objBillDetail.Detail.ID)
            End If
            objBillDetail.Detail.录入限量 = Val(Nvl(rsItems!录入限量))
            objBillDetail.Detail.药名ID = Val(Nvl(rsItems!药名ID))
            objBillDetail.Detail.变价 = Val(Nvl(rsItems!是否变价)) = 1
            objBillDetail.Detail.分批 = Val(Nvl(rsItems!分批)) = 1
            objBillDetail.Detail.跟踪在用 = Val(Nvl(rsItems!跟踪在用)) = 1
            objBillDetail.Detail.要求审批 = 0
            objBillDetail.Detail.中药形态 = Val(Nvl(rsItems!中药形态))
            '问题:41136
            str摘要 = objBillDetail.摘要
            If lng病人ID <> 0 And mint病人来源 = 2 Then
                str摘要 = gclsInsure.GetItemInfo(intInsure, lng病人ID, objBillDetail.收费细目ID, str摘要, 2, , "|1")
                objBillDetail.摘要 = str摘要
            Else
                objBillDetail.摘要 = ""
            End If
            
             '处理价格部份=====================================================
             rsItems.MoveFirst
            Set objBillDetail.InComes = New BillInComes
            Do While Not rsItems.EOF
                '按照现有的价格设置重新计算'***
                If Val(Nvl(rsItems!是否变价)) = 1 Then
                    If InStr(",5,6,7,", rsItems!收费类别) > 0 Or (rsItems!收费类别 = "4" And Nvl(rsItems!跟踪在用, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        '时价药品计算价格(分批可不分批)
                        dblAllTime = Val(Nvl(!数量))
                        If dblAllTime <> 0 Then
                            dblPrice = Get时价药品应收金额(objBillDetail.执行部门ID, CLng(Nvl(!收费细目ID)), dblAllTime, gstrDec, dblPriceSingle)
                            If dblAllTime <> 0 Then
                                If Val(Nvl(!单价)) = 0 Then
                                    '数量未分解完毕
                                    If rsItems!收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & Nvl(rsItems!名称) & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & Nvl(rsItems!名称) & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.标准单价 = 0
                                Else
                                    objBillIncome.标准单价 = Val(Nvl(!单价))
                                End If
                            Else
                                '注意：货币型最多只能保留4位小数,且不四舍五入,所以需要手工舍入;而用其它型在计算精度上又有问题
                                objBillIncome.标准单价 = IIF(dblPriceSingle = 0, Format(dblPrice / (Val(Nvl(!数量))), gstrDecPrice), dblPriceSingle)  '这里是售价价格
                            End If
                        Else
                            objBillIncome.标准单价 = 0
                        End If
                        '----------------------------------------------------------------------------------------------
                    Else
                        
                        If Abs(Val(Nvl(!单价))) > Val(Nvl(rsItems!现价)) Or Abs(Val(Nvl(!单价))) = 0 Then
                            objBillIncome.标准单价 = Val(Nvl(rsItems!缺省价格))
                        Else
                            objBillIncome.标准单价 = Val(Nvl(!单价))
                        End If
                    End If
                Else
                    objBillIncome.标准单价 = Val(Nvl(rsItems!现价))
                End If

                If InStr(",5,6,7,", rsItems!收费类别) > 0 And mbln药房单位 Then
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(rsItems!药房包装, 1), gstrDecPrice)
                Else
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrDecPrice)
                End If
                
                objBillIncome.现价 = Val(Nvl(rsItems!现价))  '现价原价对药品变价无用
                objBillIncome.原价 = Val(Nvl(rsItems!原价))
                objBillIncome.收入项目ID = Val(Nvl(rsItems!现收入ID))
                objBillIncome.收入项目 = Nvl(rsItems!收入项目)
                objBillIncome.收据费目 = Nvl(rsItems!现费目)
                
                '应收金额=单价*付次*数次
                If Val(Nvl(rsItems!是否变价)) = 1 And (InStr(",5,6,7,", rsItems!收费类别) > 0 Or rsItems!收费类别 = "4" And Nvl(rsItems!跟踪在用, 0) = 1) Then
                    objBillIncome.应收金额 = dblPrice '保证应收金额与零售金额没有误差
                Else
                    objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                End If
                
'                    '附加手术费率用计算(所有收入项目)
'                    If Val(Nvl(rsItems!附加标志)) = 1 And Nvl(rsItems!收费类别) = "F" Then
'                        objBillIncome.应收金额 = objBillIncome.应收金额 * IIf(IsNull(rsItems!附术收费率), 1, rsItems!附术收费率 / 100)
'                    End If
'
                '加班费用率计算
                dbl加班加价率 = 0
                If bln加班加价 And Val(Nvl(rsItems!加班加价)) = 1 Then
                    dbl加班加价率 = Val(Nvl(rsItems!加班加价)) / 100
                    objBillIncome.应收金额 = objBillIncome.应收金额 + objBillIncome.应收金额 * dbl加班加价率
                End If
                objBillIncome.应收金额 = Format(objBillIncome.应收金额, gstrDec)
                
                '计算实收金额
                If Val(Nvl(rsItems!屏蔽费别)) = 1 Then
                    objBillIncome.实收金额 = objBillIncome.应收金额
                Else
                    objBillIncome.实收金额 = ActualMoney(objBill.费别, Val(Nvl(rsItems!现收入ID)), objBillIncome.应收金额, _
                        objBillDetail.收费细目ID, objBillDetail.执行部门ID, objBillDetail.数次, dbl加班加价率)
                End If
                With objBillIncome
                    objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额
                End With
                '判断下一条记录是否属于当前行
                int序号 = !序号
                i = i + 1
                rsItems.MoveNext
            Loop
            With objBillDetail
                objBill.Details.Add .InComes, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .付数, .数次, .附加标志, .执行部门ID, .发药窗口, .保险项目否, .保险大类ID, .保险编码, .摘要, .Key
                '分离发药时,Key设置为1,表示编辑时执行科室列不可进入
                If InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药 Then
                    objBill.Details(objBill.Details.Count).Key = 1
                End If
            End With
            .MoveNext
        Loop
    End With
     '再重新处理从属父号
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).从属父号 <> 0 Then
            objBill.Details(i).从属父号 = colSerial("_" & objBill.Details(i).从属父号)(1)
        End If
    Next
    Set ImportWholeSet = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlGet医疗小组ID() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取补费时的医疗小组ID
    '返回:如果开单科室为病人科室且不是当前病人科室时，则取病人变动记录中的最后一次变动的医疗小组ID
    '        否则返回0,按其他规则进行处理(在存储过程中处理)
    '编制:刘兴洪
    '日期:2011-05-23 10:45:39
    '问题:37793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng开单部门ID As Long, rsTemp As ADODB.Recordset, StrSQL As String
    
    If cbo开单科室.ListIndex < 0 Then Exit Function
    lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If Not (mlng病人科室id = lng开单部门ID) Then
        Exit Function
    End If
    '只有住院才会存在
    If Not (mlng病人ID <> 0 And mlng主页ID <> 0) Then Exit Function
    StrSQL = "" & _
    "   Select 医疗小组ID From 病人变动记录 A,病人信息 B " & _
    "   Where  A.病人ID=B.病人ID  And nvl(A.终止原因,3)=3 " & _
    "               And A.科室ID<>B.当前科室ID And A.病人ID=[1] and A.主页ID=[2]  " & _
    "               And A.科室ID=[3] "
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID, lng开单部门ID)
    If rsTemp.EOF = False Then
        zlGet医疗小组ID = Val(Nvl(rsTemp!医疗小组ID))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 

