VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimpleCharge 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "简单收费处理"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimpleCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic误差 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7620
      ScaleHeight     =   600
      ScaleWidth      =   2310
      TabIndex        =   57
      Top             =   5340
      Visible         =   0   'False
      Width           =   2310
      Begin VB.Label lbl误差额 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0111"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   58
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label lbl误差 
         Caption         =   "本次误差"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   59
         Top             =   120
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8880
      Top             =   4110
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
            Picture         =   "frmSimpleCharge.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   7035
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmSimpleCharge.frx":09BC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11245
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   88
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleCharge.frx":1250
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleCharge.frx":188A
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3045
      Left            =   30
      TabIndex        =   9
      Top             =   1830
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   5371
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
   Begin VB.Frame fraTitle 
      Height          =   1035
      Left            =   30
      TabIndex        =   33
      ToolTipText     =   "清除:F6"
      Top             =   -120
      Width           =   9885
      Begin VB.TextBox txtRePrint 
         Height          =   360
         Left            =   1140
         MaxLength       =   8
         TabIndex        =   25
         Top             =   615
         Width           =   1320
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4875
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   615
         Width           =   1560
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   7860
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "单据号长度不足时自动补足长度"
         Top             =   615
         Width           =   1500
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9390
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   615
         Width           =   435
      End
      Begin VB.Label lblRePrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重打(&P)"
         Height          =   240
         Left            =   225
         TabIndex        =   24
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   240
         Left            =   4080
         TabIndex        =   10
         Top             =   675
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   11490
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   11460
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   9405
         TabIndex        =   45
         Top             =   630
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "病人收费单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   90
         TabIndex        =   38
         ToolTipText     =   "清除:F6"
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label lbl单据号 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "单据号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7080
         TabIndex        =   34
         Top             =   675
         Width           =   720
      End
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
      Height          =   555
      Left            =   15
      TabIndex        =   28
      ToolTipText     =   "清除:F6"
      Top             =   4785
      Width           =   9900
      Begin VB.ComboBox cbo结算方式 
         Height          =   360
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   150
         Width           =   1500
      End
      Begin VB.ComboBox cbo开单人 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4695
         TabIndex        =   13
         Text            =   "cbo开单人"
         Top             =   150
         Width           =   1710
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   7380
         TabIndex        =   14
         Top             =   150
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.CheckBox chk加班 
         Alignment       =   1  'Right Justify
         Caption         =   "加班(&W)"
         Height          =   270
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label lbl结算方式 
         AutoSize        =   -1  'True
         Caption         =   "结算"
         Height          =   240
         Left            =   1395
         TabIndex        =   47
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "时间"
         Height          =   240
         Left            =   6810
         TabIndex        =   36
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lbl开单人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "开单人"
         Height          =   240
         Left            =   3930
         TabIndex        =   29
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Frame fraMoney 
      Height          =   1815
      Left            =   15
      TabIndex        =   30
      Top             =   5220
      Width           =   2925
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   30
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   5
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   320
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "^序号|^项目      |^      金额"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.Frame fraStat 
      Height          =   1815
      Left            =   2940
      TabIndex        =   26
      Top             =   5220
      Width           =   4620
      Begin VB.TextBox txt应缴 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3300
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox txt预交冲款 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   3300
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   210
         Width           =   1260
      End
      Begin VB.TextBox txt累计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1290
         Width           =   1395
      End
      Begin VB.TextBox txt应收 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   315
         Width           =   1395
      End
      Begin VB.TextBox txt缴款 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3300
         MaxLength       =   12
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   990
         Width           =   1260
      End
      Begin VB.TextBox txt合计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   810
         Width           =   1395
      End
      Begin VB.TextBox txt找补 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lbl应缴 
         AutoSize        =   -1  'True
         Caption         =   "应缴金额"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   53
         Top             =   660
         Width           =   960
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         Caption         =   "预交冲款"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   2280
         TabIndex        =   51
         Top             =   270
         Width           =   960
      End
      Begin VB.Label lbl累计 
         AutoSize        =   -1  'True
         Caption         =   "累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   49
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label lbl应收 
         AutoSize        =   -1  'True
         Caption         =   "应收"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   48
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lbl缴款 
         AutoSize        =   -1  'True
         Caption         =   "缴款金额"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   43
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label lbl合计 
         AutoSize        =   -1  'True
         Caption         =   "合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   300
         Left            =   90
         TabIndex        =   37
         Top             =   855
         Width           =   630
      End
      Begin VB.Label lbl找补 
         AutoSize        =   -1  'True
         Caption         =   "找补金额"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   27
         Top             =   1440
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1020
      Left            =   30
      TabIndex        =   32
      Top             =   795
      Width           =   9885
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   645
         TabIndex        =   54
         Top             =   195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;就|就诊卡|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         NotContainFastKey=   "F1;F12;CTRL+F12;F6;F7;F8;F9;F12;CTRL+F12;ESC"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         MustSelectItems =   "姓名"
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   360
         Left            =   6105
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   580
      End
      Begin VB.ComboBox cbo医疗付款 
         Height          =   360
         Left            =   7845
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   1980
      End
      Begin VB.ComboBox cbo开单科室 
         Height          =   360
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   1890
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   2
         Top             =   195
         Width           =   1680
      End
      Begin VB.ComboBox cboSex 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   1035
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   5310
         MaxLength       =   20
         TabIndex        =   4
         Top             =   195
         Width           =   765
      End
      Begin VB.ComboBox cbo费别 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl动态费别 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5955
         TabIndex        =   52
         Top             =   630
         Width           =   3855
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         Caption         =   "付款方式"
         Height          =   240
         Left            =   6825
         TabIndex        =   50
         Top             =   255
         Width           =   960
      End
      Begin VB.Label lbl科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   240
         Left            =   135
         TabIndex        =   46
         Top             =   660
         Width           =   960
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   150
         TabIndex        =   42
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Left            =   3135
         TabIndex        =   41
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Left            =   4800
         TabIndex        =   40
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   240
         Left            =   3165
         TabIndex        =   39
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   390
      Left            =   8010
      TabIndex        =   22
      ToolTipText     =   "热键：F2"
      Top             =   6000
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   390
      Left            =   8010
      TabIndex        =   23
      ToolTipText     =   "热键:Esc"
      Top             =   6450
      Width           =   1500
   End
   Begin MSCommLib.MSComm com 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "本次误差"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0111"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   100
      TabIndex        =   55
      Top             =   310
      Width           =   1890
   End
End
Attribute VB_Name = "frmSimpleCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'————————————————————————————————————————————————————————————————————
'入口参数：
Public mbytInState As Byte '0执行(修改)收费,1-浏览收费单,2-调整单据
Public mstrInNO As String '当mbytInState=0时有效,等于单据号
Public mblnNOMoved As Boolean '操作的单据是否在后备数据表中
Public mstrDelete As String '查看第退费单据的登记时间,为""无效
Public mstrPrivs As String
Public mlngModul As Long
'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsInfo As New ADODB.Recordset '病人信息
Private mrs费别 As ADODB.Recordset      '所用费别及适用科室
Private mrs开单人 As ADODB.Recordset    '所用医生和护士列有
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室

'程序对象
Private mobjBill As ExpenseBill '费用单据对象
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mcolDetails As Details '单独的收费细目集合
Private mcolMoneys As BillInComes  '收入项目汇总集合(显示及打印时使用)

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
    项目 = 0
    应收金额 = 1
    实收金额 = 2
    执行科室 = 3
    类型 = 4
End Enum

'程序变量
Private mblnHotKey As Boolean '手工报价时,是否才按了报价热键
Private mbln报合计 As Boolean
Private mstrCardNO As String
Private mblnKeyPress As Boolean
Private mblnDo As Boolean
Private mblnDrop As Boolean         '在KeyDown中判断cbo开单人当前是否弹出
Private mblnCboClick As Boolean      '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
Private mobjICCard As Object
Private mblnNotClick As Boolean

Private mstrPreUnit As String
Private mblnValid As Boolean
Private mstr付款方式 As String
Private mlng领用ID As Long
Private mintBillNO As Integer
Private mintMoneyRow As Integer '当前显示到的费目行
Private mbln不重算价格 As Boolean

Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mintOldInvoiceFormat As Integer '旧票据打印格式
Private mintInvoicePrint As Integer '0-不打印;1-自动打印;2-提示打印
Private mblnStartFactUseType As Boolean

'收费处同一病人病人单据累计金额
Private mstrPrePati As String  '上一个收费病人
Private mcurBill应收 As Currency
Private mcurBill实收 As Currency
Private mcurBill应缴 As Currency

Private marrColData() As Integer '当前单据编辑属性映象
Private mblnPrint As Boolean
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器
Private Const STR_HEAD = "项目,3000,1;应收金额,1500,7;实收金额,1500,7;执行科室,1950,1;类型,1000,1"
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mstr药品价格等级 As String, mstr卫材价格等级 As String, mstr普通价格等级 As String

Private Sub Bill_BeforeAddRow(Row As Long)
    Dim dbl单价  As Double, curMoney As Currency, i As Integer
    'LED动态显示项目
    If gblnLED And mbytInState = 0 And mobjBill.Pages(1).Details.Count >= Row - 1 Then
        With mobjBill.Pages(1).Details(Row - 1)
            dbl单价 = 0: curMoney = 0
            For i = 1 To .InComes.Count
                curMoney = curMoney + .InComes(i).实收金额
                dbl单价 = dbl单价 + .InComes(i).标准单价
            Next
            'LED显示
            If curMoney <> 0 Then
                zl9LedVoice.Display .Detail.名称, .Detail.规格, .计算单位, dbl单价, .数次, curMoney
            End If
        End With
    End If
End Sub

Private Sub ShowGroupLED(ByVal lngMain As Long, ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：为加快速度，一次性调用套餐项目的LED显示
'参数：行号范围，lngMain=主项行号,lngBegin-lngEnd:从项行号
    Dim dbl数量 As Double, dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    
    If gblnLED Then
        With mobjBill.Pages(1)
            For j = 1 To .Details(lngMain).InComes.Count
                cur金额 = cur金额 + .Details(lngMain).InComes(j).实收金额
            Next
            For i = lngBegin To lngEnd
                For j = 1 To .Details(i).InComes.Count
                    cur金额 = cur金额 + .Details(i).InComes(j).实收金额
                Next
            Next
        End With
        With mobjBill.Pages(1).Details(lngMain)
            If cur金额 <> 0 Then
                dbl数量 = .数次
                If dbl数量 <> 0 Then
                    dbl单价 = cur金额 / dbl数量
                Else
                    dbl单价 = cur金额
                End If
                zl9LedVoice.Display .Detail.名称, .Detail.规格, .计算单位, dbl单价, dbl数量, cur金额
            End If
        End With
    End If
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytSubs As Byte
    
    If mobjBill.Pages(1).Details.Count >= Row Then
        If mobjBill.Pages(1).Details(Row).工本费 Then
            MsgBox "该行不能修改及删除！", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End If
    
    If mobjBill.Pages(1).Details.Count >= Row Then
        '带从属项目的项删除确认
        For i = Row + 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).从属父号 = Row Then bytSubs = bytSubs + 1
        Next
        If bytSubs > 0 Then
            If MsgBox("该项目带有 " & bytSubs & " 个从属项目,删除该项目也将删除它的从属项目,继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Pages(1).Details(Row).从属父号 <> 0 Then '从属项目删除确认
            If MsgBox("该项目是[" & mobjBill.Pages(1).Details(mobjBill.Pages(1).Details(Row).从属父号).Detail.名称 & "]的从属项目,确定要删除它吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf MsgBox("确实要删除该收费项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        '删除处理
        For i = mobjBill.Pages(1).Details.Count To Row + 1 Step -1
            If mobjBill.Pages(1).Details(i).从属父号 = Row Then
                Call DeleteDetail(i) '反顺序删除其从属行
            End If
        Next
        Call DeleteDetail(Row) '删除该行
        
        '重新计算并刷新
        'Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
        
        If mobjBill.Pages(1).Details.Count = 0 Then ClearMoney
        
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '不用控件来处理了
    End If
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim lng执行科室 As Long, str执行科室 As String
    If mobjBill.Pages(1).Details.Count >= Bill.Row Then
        If Bill.ListIndex <> -1 Then
            If mobjBill.Pages(1).Details(Bill.Row).执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                lng执行科室 = mobjBill.Pages(1).Details(Bill.Row).执行部门ID: str执行科室 = Bill.TextMatrix(Bill.Row, Bill.Col)
                mobjBill.Pages(1).Details(Bill.Row).执行部门ID = Bill.ItemData(Bill.ListIndex)
                If ItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row)
                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, 1, Bill.Row)) = False Then
                    Bill.Text = "": Bill.TxtVisible = False
                    Bill.cboObj.Text = str执行科室: mobjBill.Pages(1).Details(Bill.Row).执行部门ID = lng执行科室: Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub Bill_CommandClick()
    Dim lng项目id As Long, blnCancel As Boolean
        
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, gint病人来源, 0, False, "'Z'", , , _
        , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    If lng项目id <> 0 Then
        Bill.Text = lng项目id
        mblnSelect = True
        Call Bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'功能：处理单据输入
    Dim strScope As String, i As Long
    Dim objDetail As Detail, lng项目id As Long, lngDoUnit As Long
    
    If KeyCode = 13 And Not Bill.Active Then
        Cancel = True: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
        
    On Error GoTo errH
    
    If KeyCode = 13 Then
        '收费时,工本费不能修改
        If mobjBill.Pages(1).Details.Count >= Bill.Row Then
            If mobjBill.Pages(1).Details(Bill.Row).工本费 Then Exit Sub
        End If
        If Bill.ColData(Bill.Col) = 0 Then Exit Sub
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "项目"
                '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
                If Bill.Text <> "" Then
                    If mblnSelect Then
                        mblnSelect = False '立即清除该标志
                        Set objDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, gint病人来源, 0, False, "'Z'", Bill.Text, Bill.TxtHwnd, _
                                        , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                        If lng项目id <> 0 Then
                            Set objDetail = GetInputDetail(lng项目id)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    sta.Panels(2) = ""
                    Bill.TxtVisible = False '(不加不行)
                    '加入或修改该收费细目行
                    Call SetDetail(objDetail, Bill.Row)
                    
                    '输入摘要(根据新输入的行更改摘要)
                    Dim str摘要 As String '90304
                    If mobjBill.Pages(1).Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Pages(1).Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Pages(1).Details(Bill.Row).摘要 = str摘要
                        End If
                    Else
                         str摘要 = gclsInsure.GetItemInfo(0, mobjBill.病人ID, mobjBill.Pages(1).Details(Bill.Row).收费细目ID, str摘要, 1)
                         mobjBill.Pages(1).Details(Bill.Row).摘要 = str摘要
                    End If
                    
                    Call CalcMoneys(Bill.Row)
                    
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, 1, Bill.Row)) = False Then
                        mobjBill.Pages(1).Details.Remove Bill.Row '删除刚刚想要加入的费用行
                        Bill.Text = "": Bill.TxtVisible = False
                        Cancel = True: Exit Sub
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '费用类型检查
                    Call Check费用类型(Bill.Row)
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Pages(1).Details.Count >= Bill.Row Then
                    '下一列的性质确定
                    If mobjBill.Pages(1).Details(Bill.Row).Detail.变价 Then Bill.ColData(1) = 4 '应收金额
                    
                    '执行科室!!!
                    Call FillBillComboBox(Bill.Row, 3)
                    If Bill.ListCount = 1 Then
                        Bill.ColData(3) = 5
                        mobjBill.Pages(1).Details(Bill.Row).Key = 1
                    Else
                        Bill.ColData(3) = 3
                        mobjBill.Pages(1).Details(Bill.Row).Key = Bill.ListCount
                    End If
                    
                    '从属项目处理(在这里可以处理多级从属-从属的从属...)
                    If Bill.TextMatrix(0, Bill.Col) = "项目" Then
                        If ShouldDO(Bill.Row) Then
                            Set mcolDetails = New Details
                            Set mcolDetails = GetSubDetails(mobjBill.Pages(1).Details(Bill.Row).收费细目ID)
                            For i = 1 To mcolDetails.Count
                                If mobjBill.Pages(1).Details.Count >= Bill.Rows - 1 Then
                                    Bill.Rows = Bill.Rows + 1
                                    Call bill_AfterAddRow(Bill.Rows - 1)
                                End If
                                Bill.TextMatrix(Bill.Rows - 1, 0) = "" '有必要加上
                                
                                If mcolDetails(i).类别 = mobjBill.Pages(1).Details(Bill.Row).收费类别 Then
                                    '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                                    lngDoUnit = mobjBill.Pages(1).Details(Bill.Row).执行部门ID
                                Else
                                    If mcolDetails(i).执行科室 = 0 Then
                                        '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                                        lngDoUnit = mobjBill.Pages(1).Details(Bill.Row).执行部门ID
                                    End If
                                        '其余情况,取本身设置的执行科室
                                End If
                                            
                                Call SetDetail(mcolDetails(i), Bill.Rows - 1, Bill.Row, lngDoUnit)
                                Call CalcMoneys(Bill.Rows - 1)
                                Call ShowDetails(Bill.Rows - 1)
                                Call ShowMoney
                            Next
                            '一次性调用套餐项目LED显示
                            Call ShowGroupLED(Bill.Row, Bill.Rows - mcolDetails.Count, Bill.Rows - 1)
                        End If
                    End If
                End If
            Case "应收金额" '实际上是单价(因为数据次缺省为1,且不能更改)
                If mobjBill.Pages(1).Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '负数权限
                    If InStr(mstrPrivs, "负数费用") = 0 And CDbl(Bill.Text) < 0 Then
                        MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    Bill.Text = Format(Bill.Text, gstrDec)
                    
                    If mobjBill.Pages(1).Details.Count >= Bill.Row And Bill.Text <> "" Then
                        '如果没有对应的收入项目,则无法计算
                        If mobjBill.Pages(1).Details(Bill.Row).Detail.变价 And mobjBill.Pages(1).Details(Bill.Row).InComes.Count > 0 Then
                            If Not (mobjBill.Pages(1).Details(Bill.Row).InComes(1).现价 = 0 And mobjBill.Pages(1).Details(Bill.Row).InComes(1).原价 = 0) Then
                                strScope = CheckScope(mobjBill.Pages(1).Details(Bill.Row).InComes(1).原价, mobjBill.Pages(1).Details(Bill.Row).InComes(1).现价, CCur(Bill.Text))
                                If strScope <> "" Then
                                    sta.Panels(2) = strScope
                                    If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Pages(1).Details(Bill.Row).InComes(1).标准单价
                                    If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                    Cancel = True: Beep: Exit Sub
                                End If
                            End If
                            
                            '这种收费细目只能对应一个收入项目
                            mobjBill.Pages(1).Details(Bill.Row).数次 = Sgn(Val(Bill.Text))
                            mobjBill.Pages(1).Details(Bill.Row).InComes(1).标准单价 = Abs(Val(Bill.Text))
                            
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        Else
                            Bill.Text = "0"
                            sta.Panels(2) = "该项目设有设置对应的费目，所以无法计算费用！"
                            Beep
                        End If
                    End If
                End If
            Case "执行科室"
                If mobjBill.Pages(1).Details.Count >= Bill.Row Then
                   If Bill.ListIndex <> -1 Then
                        'If mobjBill.Pages(1).Details(Bill.Row).执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                            mobjBill.Pages(1).Details(Bill.Row).执行部门ID = Bill.ItemData(Bill.ListIndex)
                            If ItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row)
                        'End If
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, 1, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'功能:根据主项执行科室的变化,刷新非药从项的执行科室
    Dim i As Long, j As Long, lng病人科室ID As Long
    
    lng病人科室ID = mobjBill.科室ID
    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)

    With mobjBill.Pages(1)
        For i = lngRow + 1 To .Details.Count
            If .Details(i).从属父号 = lngRow Then
                '从属项为药品和卫材的项目的执行科室不随主项变动
                If InStr(",4,5,6,7,", .Details(i).收费类别) = 0 Then
                    If .Details(i).收费类别 = .Details(lngRow).收费类别 Then
                        '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                        .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                    Else
                        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID) '必须现取
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
                                .Details(i).执行部门ID = Get收费执行科室ID(mcolDetails(j).类别, mcolDetails(j).ID, _
                                    mcolDetails(j).执行科室, lng病人科室ID, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
                            End If
                        End If
                    End If
                    
                    If .Details(i).执行部门ID > 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, 3) = mrsUnit!编码 & "-" & mrsUnit!名称
                            Else
                                Bill.TextMatrix(i, 3) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, 3) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                        End If
                    End If
                End If
            End If
        Next
    End With

End Sub

Private Sub Set开单人开单科室(ByVal str开单人 As String, ByVal lng开单科室ID As Long)
'功能:根据开单人或开单科室ID设置开单科室及开单人,但不触发点击事件
       '利用公共函数CboSetIndex避免隐式调用cbo_click事件
    
    Dim str开单科室 As String, lng人员ID As Long
    
    'a.医生确定科室
    If gbyt科室医生 = 0 Then
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True)) '不触发click事件
        
        If cbo开单人.ListIndex = -1 And str开单人 <> "" Then
            lng人员ID = GetPersonnelID(str开单人, mrs开单人)
            cbo开单人.AddItem str开单人, 0
            cbo开单人.ItemData(cbo开单人.NewIndex) = lng人员ID
            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo开单人.NewIndex)
        End If
                
        If cbo开单人.ListIndex <> -1 Then
            cbo开单科室.Clear
            Call FillDept(cbo开单人.ItemData(cbo开单人.ListIndex))
        End If
        
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        If cbo开单科室.ListIndex = -1 And lng开单科室ID > 0 Then
            str开单科室 = GET部门名称(lng开单科室ID)
            If str开单科室 <> "" Then
                cbo开单科室.AddItem str开单科室, 0
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = lng开单科室ID
                Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo开单科室.NewIndex)
            End If
        End If
        
    'b.科室确定医生或独立输入
    Else
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        
        If cbo开单科室.ListIndex = -1 And lng开单科室ID > 0 Then
            str开单科室 = GET部门名称(lng开单科室ID)
            If str开单科室 <> "" Then
                cbo开单科室.AddItem str开单科室, 0
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = lng开单科室ID
                Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo开单科室.NewIndex)
            End If
        End If
        
        If gbyt科室医生 = 1 And cbo开单科室.ListIndex <> -1 Then
            cbo开单人.Clear
            Call FillDoctor(lng开单科室ID)
        End If
        
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True))
        If cbo开单人.ListIndex = -1 And str开单人 <> "" Then
            lng人员ID = GetPersonnelID(str开单人, mrs开单人)
            cbo开单人.AddItem str开单人, 0
            cbo开单人.ItemData(cbo开单人.NewIndex) = lng人员ID
            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo开单人.NewIndex)
        End If
    End If
End Sub

Private Sub Set开单人开单科室Click(ByVal str开单人 As String, ByVal lng开单科室ID As Long)
'功能:根据开单人或开单科室ID设置开单科室及开单人,并触发点击事件
'     当Listindex=x时,如果Listindex的值本身等于x,就不会触发点击事件,所以要用API+Click强制调用
    Dim i As Long
    
    If gbyt科室医生 = 0 Then
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True)) '不触发click事件
        Call cbo开单人_Click
        
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        Call cbo开单科室_Click
        
    Else
        '科室确定医生或各自独立输入
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        Call cbo开单科室_Click
        
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True)) '不触发click事件
        Call cbo开单人_Click
    End If
End Sub

Private Function ItemHaveSub(ByVal lngRow As Long) As Boolean
'功能：判断当前行的项目是否具有从属项目
    Dim i As Long
    
    If mobjBill.Pages(1).Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).从属父号 = lngRow Then
                ItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim i As Integer, bln工本费 As Boolean
    
    If Not Bill.Active Then Exit Sub
    
    '恢复列编辑属性
    If mbytInState = 0 Then
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
    End If
    
    '收费时,如果为工本费,则不能修改
    If mobjBill.Pages(1).Details.Count >= Row And mbytInState = 0 Then
        If mobjBill.Pages(1).Details(Row).工本费 Then
            bln工本费 = True
            For i = 0 To UBound(marrColData)
                Bill.ColData(i) = IIf(marrColData(i) = 5, 5, 0)
            Next
        End If
    End If
    
    '如果是从属项目的主项目或从项,则不允许更改类别和项目
    If mobjBill.Pages(1).Details.Count >= Row Then
        If ItemHaveSub(Row) Or mobjBill.Pages(1).Details(Row).从属父号 > 0 Then
            Bill.ColData(0) = BillColType.Text_UnModify
        End If
    End If
    
    '执行科室列
    If mobjBill.Pages(1).Details.Count >= Bill.Row And mbytInState <> 2 And Not bln工本费 Then
        If mobjBill.Pages(1).Details(Bill.Row).Key = "1" Then
            Bill.ColData(3) = 5
        Else
            Bill.ColData(3) = 3
        End If
    End If
    If Bill.ColData(Bill.Col) = 3 Then Call FillBillComboBox(Bill.Row, Bill.Col)
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "执行科室"
            SetWidth Bill.cboHwnd, 130
        Case "应收金额"
            Bill.TextLen = 10
            If InStr(mstrPrivs, "负数费用") = 0 Then
                Bill.TextMask = "0123456789." & Chr(8)
            Else
                Bill.TextMask = "-0123456789." & Chr(8)
            End If
    End Select

    '进入行时,重新设置该行的编辑性质
    If mobjBill.Pages(1).Details.Count >= Bill.Row And Not bln工本费 Then
        If mobjBill.Pages(1).Details(Bill.Row).Detail.变价 Then
            Bill.ColData(1) = 4
        Else
            Bill.ColData(1) = 5
        End If
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboSex_Click()
    If mbytInState = 0 Then
        mobjBill.性别 = zlStr.NeedName(cboSex.Text)
    End If
End Sub

Private Sub cbo费别_Click()
    If cbo费别.ListIndex <> -1 Then
        If mobjBill.费别 <> zlStr.NeedName(cbo费别.Text) And Not mbln不重算价格 Then
            mobjBill.费别 = zlStr.NeedName(cbo费别.Text)
            
            If mbytInState = 0 And mobjBill.Pages(1).Details.Count > 0 Then
                '重新计算价格
                Call CalcMoneys
                Call ShowDetails
                Call ShowMoney
            End If
        End If
    End If
End Sub

Private Sub cbo结算方式_Click()
'功能：在现金与非现金之间切换时，需要根据情况决定是否处理分币
    Dim dblTemp As Double
    
    If Not (Visible And gBytMoney <> 0) Then Exit Sub
    If Bill.Active Then
        Call ShowMoney
    ElseIf chkCancel.Value = 1 Then
        txt应缴.Text = Format(GetDelMoney, "0.00")
        
        '误差显示
        dblTemp = -1 * Format((Val(txt合计.Text) - Val(txt预交冲款.Text) - Val(txt应缴.Text)), gstrDec)
        If dblTemp <> 0 Then
            pic误差.Visible = True
            lbl误差额.Caption = Format(dblTemp, "0.00")
        Else
            pic误差.Visible = False
        End If
    Else
        Call ShowPrice
    End If
End Sub

Private Function GetDelMoney() As Currency
    Dim cur退费合计 As Currency
    Dim bln现金结算 As Boolean
    Dim i As Integer
    
    cur退费合计 = Format(Val(txt合计.Text) - Val(txt预交冲款.Text), "0.00")
    
    '现金结算时处理分币(因为部份退费时不调用医保接口,因此不管医保是否支持分币)
    bln现金结算 = False
    If cbo结算方式.ListIndex <> -1 Then
        If cbo结算方式.ItemData(cbo结算方式.ListIndex) = 1 Then
            bln现金结算 = True
        End If
    End If
    If bln现金结算 Then
        cur退费合计 = CentMoney(Val(txt合计.Text) - Val(txt预交冲款.Text))
    End If
    GetDelMoney = cur退费合计
End Function

Private Sub cbo结算方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii >= 32 Then
        If cbo结算方式.Locked Then Exit Sub
        
        lngIdx = zlControl.CboMatchIndex(cbo结算方式.hWnd, KeyAscii)
        If lngIdx = -1 And cbo结算方式.ListCount > 0 Then lngIdx = 0
        cbo结算方式.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, lng开单部门ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
        
    If cbo开单科室.ListIndex <> -1 Then lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If mobjBill.Pages(1).开单部门ID = lng开单部门ID Then Exit Sub
    mobjBill.Pages(1).开单部门ID = lng开单部门ID
    
    '定位医生
    If gbyt科室医生 = 1 Then
        If cbo开单科室.ListIndex <> -1 Then
            Call FillDoctor(lng开单部门ID)
            
            If cbo开单人.ListCount > 0 And Not gbln不缺省开单人 Then
                Call zlControl.CboSetIndex(cbo开单人.hWnd, 0)
            End If
        Else
            cbo开单人.Clear
        End If
        Call cbo开单人_Click
    End If
    
    
    '根据开单科室重新设置收费项目的执行科室
    If cbo开单科室.ListIndex <> -1 And Visible Then
        With mobjBill.Pages(1)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.类别) = 0 And _
             (.Details(i).Detail.执行科室 = 6 And gbyt科室医生 <> 2 Or InStr(",1,2,", "," & .Details(i).Detail.执行科室 & ",") > 0 And gint病人来源 = 1) Then '6-开单人科室
                
                .Details(i).执行部门ID = lng开单部门ID
                
                If i <= Bill.Rows - 1 And .Details(i).执行部门ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                        End If
                    Else
                        '浏览单据只(能)显示名称
                        Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(i, BillCol.执行科室) = ""
                End If
            End If
        Next
        End With
    End If
        
    '费别处理
    Call LoadAndSeek费别
End Sub


Private Sub cbo开单科室_Validate(Cancel As Boolean)
 '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件

    If Not mblnCboClick Then cbo开单科室_Click
    mblnCboClick = False
End Sub

Private Function SetDefaultDept(lng开单人ID As Long) As Boolean
'功能:设置缺省的开单科室,但不触发Click事件
'说明:缺省科室为"只服务于门诊,不具有医技性质"时，可以定位缺省
'     或者开单人的所有科室都为同一优先排序级别时(如都是即服务于门诊或住院的)，可以定位缺省
'     否则,不管人员的缺省科室，以GetDoctorDept中的医生顺序为准,第一个为缺省
'     该顺序为: 1.只服务于门诊,不具有医技性质(检查,检验,手术,治疗,营养)
'               2.只服务于门诊,具有医技性质(检查,检验,手术,治疗,营养)
'               3.不只服务于门诊的
    Dim i As Long, lng开单科室ID As Long, lng优先级 As Long, blnDo As Boolean
    
    mrs开单人.Filter = "缺省=1 And ID=" & lng开单人ID
    If mrs开单人.RecordCount > 0 Then lng开单科室ID = mrs开单人!部门ID
        
    If mrs开单科室.RecordCount > 1 And lng开单科室ID > 0 Then
        If gbln缺省科室优先 Then
            blnDo = True
        Else
            mrs开单科室.MoveFirst
            For i = 1 To mrs开单科室.RecordCount
                If lng开单科室ID = mrs开单科室!ID And mrs开单科室!优先级 = 1 Then blnDo = True: Exit For
                mrs开单科室.MoveNext
            Next
            
            If Not blnDo Then
                blnDo = True
                mrs开单科室.MoveFirst
                For i = 1 To mrs开单科室.RecordCount
                    If lng优先级 <> mrs开单科室!优先级 And lng优先级 <> 0 Then blnDo = False: Exit For
                    lng优先级 = mrs开单科室!优先级
                    mrs开单科室.MoveNext
                Next
            End If
        End If
        
        If blnDo Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
    End If
    
    If cbo开单科室.ListIndex = -1 Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, 0)
End Function

Private Sub cbo开单人_Click()
    Dim i As Long, lng开单人ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.Pages(1).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text)) Then Exit Sub
    
    mobjBill.Pages(1).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    '由医生确定科室
    If gbyt科室医生 = 0 Then
        If cbo开单人.ListIndex <> -1 Then
            lng开单人ID = cbo开单人.ItemData(cbo开单人.ListIndex)
            
            Call FillDept(lng开单人ID)
            Call SetDefaultDept(lng开单人ID)
        Else
            cbo开单科室.Clear
        End If
        Call cbo开单科室_Click
    End If
    
    '科室医生独立,因为开单人变了，所以,执行科室是由开单人科室决定时，需要重设执行科室
     '不独立时在Cbo开单科室_click中处理
    If cbo开单人.ListIndex <> -1 And Visible And gbyt科室医生 = 2 Then
        With mobjBill.Pages(1)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.类别) = 0 And .Details(i).Detail.执行科室 = 6 Then    '6-开单人科室
                
                mrs开单人.Filter = "缺省=1 And ID=" & lng开单人ID
                If mrs开单人.RecordCount = 0 Then mrs开单人.Filter = "ID=" & lng开单人ID
                If mrs开单人.RecordCount > 0 Then
                    .Details(i).执行部门ID = mrs开单人!部门ID
                Else
                    .Details(i).执行部门ID = 0
                End If
                
                If i <= Bill.Rows - 1 And .Details(i).执行部门ID > 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                        End If
                    Else
                        '浏览单据只(能)显示名称
                        Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(i, BillCol.执行科室) = ""
                End If
            End If
        Next
        End With
    End If
End Sub



Private Sub cbo开单人_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo开单人.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo开单人.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo开单人_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub cbo开单人_Validate(Cancel As Boolean)
    If cbo开单人.Text <> "" Then
        If cbo.FindIndex(cbo开单人, zlStr.NeedName(cbo开单人.Text), True) = -1 Then cbo开单人.ListIndex = -1: cbo开单人.Text = ""
    End If
    If cbo开单人.Text = "" Then Call cbo开单人_KeyPress(vbKeyReturn)
    If gbyt科室医生 = 0 And gbln必须输开单人 And cbo开单人.ListIndex = -1 Then Cancel = True
End Sub

Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo年龄单位_Validate(Cancel As Boolean)
    If mbytInState = 0 Then mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
End Sub

Private Sub cbo医疗付款_Click()
    On Error GoTo errHandler
    If mbytInState <> 0 Then Exit Sub
    If gintPriceGradeStartType < 2 Then Exit Sub
    
    If mrsInfo.State = adStateOpen Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), zlStr.NeedName(cbo医疗付款.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    Else
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cbo医疗付款.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    End If
    
    If mbln不重算价格 Then Exit Sub
    If mobjBill.Pages(1).Details.Count = 0 Then Exit Sub
    
    '重新计算价格
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo医疗付款_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii >= 32 Then
        If cbo医疗付款.Locked Then Exit Sub
    
        lngIdx = zlControl.CboMatchIndex(cbo医疗付款.hWnd, KeyAscii)
        If lngIdx = -1 And cbo医疗付款.ListCount > 0 Then lngIdx = 0
        cbo医疗付款.ListIndex = lngIdx
    
    ElseIf KeyAscii = 13 And cbo医疗付款.ListIndex <> -1 Then
        If Bill.Active Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Not Bill.Active Then
            If gbyt科室医生 = 0 Then
                If cbo结算方式.Enabled Then cbo结算方式.SetFocus
            Else
                If cbo开单科室.Enabled Then cbo开单科室.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chkCancel_Click()
    mstrInNO = ""
    txt找补.Text = "0.00": txt缴款.Text = "0.00": txt应缴.Text = "0.00"
    mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
    mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
    Call ClearPatientInfo
    txt合计.Text = gstrDec: txt应收.Text = gstrDec
    
    If chkCancel.Value = Checked Then
        chkCancel.ForeColor = &HFF&
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney
        Call NewBill(False)
        Call SetDisible
        
        cboNO.Text = "": cboNO.Locked = False
        txtInvoice.Text = "": txtInvoice.Locked = True
        txtRePrint.Locked = True
        
        lbl应缴.Caption = "应退金额"
        lbl应缴.ForeColor = vbRed
        txt应缴.ForeColor = vbRed
        txt应缴.Text = "0.00"
        
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        txtInvoice.Locked = Not (InStr(1, mstrPrivs, "修改票据号") > 0) And gblnStrictCtrl
        txtRePrint.Locked = False
        
        lbl应缴.Caption = "应缴金额"
        lbl应缴.ForeColor = 0
        txt应缴.ForeColor = &HFF0000
        txt应缴.Text = "0.00"
        
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney
        Call NewBill
        Call SetDisible(True)
        txtPatient.SetFocus
    End If
End Sub

Private Sub chk加班_Click()
    If Not mblnDo Then Exit Sub
    If mbytInState = 1 Or chkCancel.Value = 1 Then Exit Sub
    If Not chk加班.Visible Or Not Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk加班.Value = Unchecked And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Checked
        End If
    End If
    If chk加班.Value = Checked And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Unchecked
        End If
    End If
    mobjBill.加班标志 = IIf(chk加班.Value = Checked, 1, 0)
    
    '重新计算价格
    If Not mobjBill.Pages(1).Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If (mobjBill.Pages(1).Details.Count > 0 Or txtPatient.Text <> "") And Bill.Active And mbytInState = 0 And mstrInNO = "" Then
        If MsgBox("确实要清除当前单据中的内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    
        txt找补.Text = "0.00"
        txt缴款.Text = "0.00"
        txt应缴.Text = "0.00"
        If chkCancel.Value = Checked Then '退据单状态
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney
            chkCancel.Value = Unchecked
            Call NewBill
            Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Bill.Active Then '正常输入单据状态'(清除后当作是新病人单据)
            mcurBill实收 = 0:  mcurBill应收 = 0: mcurBill应缴 = 0
            mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
            Call ClearPatientInfo
            txt合计.Text = gstrDec: txt应收.Text = gstrDec
            If gbln累计 Then txt累计.Text = Format(GetChargeTotal, "0.00")
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney
            Call NewBill '保持原单据号
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Not Bill.Active Then '收取划价单费用状态
            Call ClearPatientInfo
            txt合计.Text = gstrDec: txt应收.Text = gstrDec
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney
            Call NewBill
            Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub ClearPatientInfo()
    txtPatient.Text = "": txtPatient.Tag = ""
    txt年龄.Text = ""
    Call zlControl.CboLocate(cbo年龄单位, "岁")
    Call txt年龄_Validate(False)
End Sub

Private Function GetCboIndexByCode(ByRef objCbo As ComboBox, ByVal strCode As String) As Integer
    Dim i As Integer
    
    GetCboIndexByCode = -1
    For i = 0 To objCbo.ListCount - 1
        If strCode = Mid(objCbo.List(i), 1, InStr(1, objCbo.List(i), "-") - 1) Then
            GetCboIndexByCode = i
            Exit For
        End If
    Next
End Function

Private Sub cmdOK_Click()
    Dim strInfo As String, strSQL As String
    Dim i As Long, j As Long, lng结帐ID As Long
    Dim curMoney As Currency, cur工本费 As Currency
    Dim str医疗付款 As String
    Dim str划价Nos As String, rsItems As ADODB.Recordset
    
    If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
    
    If mbytInState = 2 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入合法的费用时间！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If Not SaveModi() Then Exit Sub
        Unload Me
    ElseIf Bill.Active Then '正常输入单据状态
        Call txt缴款_GotFocus
        
        If txtPatient.Text = "" And mobjBill.姓名 = "" Then
            MsgBox "没有发现病人信息,请输入病人信息！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        Else
            If mobjBill.姓名 = "" Then
                mobjBill.姓名 = txtPatient.Text
            Else
                txtPatient.Text = mobjBill.姓名
            End If
        End If
        
        
        If CheckTextLength("姓名", txtPatient) = False Then Exit Sub
        If CheckTextLength("年龄", txt年龄) = False Then Exit Sub
        If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Sub
        
        If cbo费别.ListIndex = -1 Or mobjBill.费别 = "" Then
            MsgBox "请选择病人费别！", vbInformation, gstrSysName
            If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus: Exit Sub
        End If
        If mobjBill.Pages(1).Details.Count = 0 Then
            MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Check执行科室
        If i <> 0 Then
            MsgBox "单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If cbo开单科室.ListIndex = -1 Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            If gbyt科室医生 = 0 Then
                cbo开单人.SetFocus
            Else
                cbo开单科室.SetFocus
            End If
            Exit Sub
        End If
        
        '开单人
        If gbln必须输开单人 And cbo开单人.ListIndex = -1 Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        If cbo结算方式.ListIndex = -1 Then
            MsgBox "请确定收费的结算方式！", vbInformation, gstrSysName
            cbo结算方式.SetFocus: Exit Sub
        End If
    
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If Val(txt缴款.Text) <> 0 And txt缴款.Enabled Then
            If Val(txt缴款.Text) < Val(txt应缴.Text) Then
                MsgBox "病人缴款金额不足，请补足应缴金额！", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt缴款): txt缴款.SetFocus: Exit Sub
            End If
        End If
        '刘兴洪:22343,缴款金额控制
        Select Case gTy_Module_Para.byt缴款控制
        Case 1  '1-代表输入缴款后才结束病人累计
        Case 2  '2-收费时必须要输入缴款金额
            If Val(txt应缴.Text) > 0 And Val(txt缴款.Text) = 0 Then
                MsgBox "注意:" & vbCrLf & _
                "    该病人未输入缴款金额,不能进行收费!", vbInformation + vbDefaultButton1, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                Exit Sub
            End If
        Case Else   ',0-代表不进行缴款输入和累计控制
        End Select
                
                
        '非法行
        For i = 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).收费细目ID = 0 Then
                MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
        Next
        
        '费用类型检查
        If Not Check费用类型 Then Exit Sub
                
        '单张单据最高额
        If gcurMax <> 0 And CalcBillToTal > gcurMax Then
            MsgBox "单据金额超过最大限制金额:" & Format(gcurMax, "0.00") & " ,不允许保存！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0)) = False Then
            Exit Sub
        End If
        
        '票据号码检查,工本费打印检查
        mblnPrint = True
        '检查是否打印票据
        If mintInvoicePrint = 0 Then
            mblnPrint = False
        Else
            If mintInvoicePrint = 2 Then
                If MsgBox("是否打印票据？" & vbCrLf & "要取消此提示,请在本地参数中设置票据打印控制参数!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
        
        '检查零费用(只有工本费)是否打印,划价不产生工本费
        If mblnPrint Then
            If CalcBillToTal = Calc工本费 Then
                If MsgBox("当前单据实际没有收取费用,要打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
           
        If Not mblnPrint Then
            j = 0
            For i = 1 To mobjBill.Pages(1).Details.Count
                If mobjBill.Pages(1).Details(i).工本费 Then
                    If j = 0 Then MsgBox "因为不打印票据,系统将自动删除工本费！", vbInformation, gstrSysName
                    j = j + 1
                    Call DeleteDetail(i)
                    Call ShowDetails
                    Call ShowMoney
                    Bill.TxtVisible = False:  Bill.CmdVisible = False: Bill.CboVisible = False
                    Exit For
                End If
            Next
        Else
            If gblnStrictCtrl Then
                If Trim(txtInvoice.Text) = "" Then
                    MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
                If zlGetInvoiceGroupUseID(mlng领用ID, 1, txtInvoice.Text) = False Then
                    Exit Sub
                End If
                 
                '并发操作检查,票号是否已用
                If CheckBillRepeat(mlng领用ID, 1, txtInvoice.Text) Then
                    MsgBox "票据号""" & txtInvoice.Text & """已经被使用，请重新输入。", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            Else
                If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                    MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            End If
        End If
        
                
        If gblnLED And Val(txt应缴.Text) <> 0 And Not mbln报合计 And Not gbln手工报价 Then
            zl9LedVoice.Speak "#21 " & txt应缴.Text
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
        mobjBill.登记时间 = zlDatabase.Currentdate
        If zlGetSaveDataItems_Plugin(mobjBill, str划价Nos, rsItems) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(glngModul, 1, True, False, str划价Nos, rsItems) = False Then Exit Sub
        
        '保存单据
        If Not SaveBill Then Exit Sub
        
        Call zlChargeSaveAfter_Plugin(glngModul, mobjBill.病人ID, mobjBill.主页ID, True, 1, mobjBill.NO)
        
        '保存后的处理
        If mblnPrint Then '打印门诊收据
            Call frmPrint.ReportPrint(1, "'" & mobjBill.NO & "'", "", "", mlng领用ID, mlngShareUseID, txtInvoice.Text, _
                mobjBill.登记时间, , , , mintInvoiceFormat, , , mstrUseType, , , , mstr普通价格等级)
        End If
        
        '费用清单的打印
        If InStr(mstrPrivs, "打印清单") > 0 Then
            If gint收费清单 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mobjBill.NO & "'", "药品单位=0", 2)
            ElseIf gint收费清单 = 2 Then
                If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mobjBill.NO & "'", "药品单位=0", 2)
                End If
            End If
        End If
        
        If mbytInState = 0 And gbln累计 Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
        End If
        
        If mstrInNO = "" Then
            mstrInNO = ""
            sta.Panels(2) = "上一张单据:" & mobjBill.NO
            '是否可以连续收费：
            '使用预交款结算,当次收费结束(除非设置仅缴款结束参数)
            '如已缴款,则强行作为病人收费结束
            
            '刘兴洪:22343
            If CCur(txt缴款.Text) <> 0 Or (Val(txt预交冲款.Text) <> 0 And gTy_Module_Para.byt缴款控制 <> 1) Then
                mcurBill实收 = 0:  mcurBill应收 = 0: mcurBill应缴 = 0
                mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
                Call ClearPatientInfo
                txt合计.Text = gstrDec: txt应收.Text = gstrDec
            Else
                mstrPrePati = mobjBill.姓名 '记录当前病人
                
                '病人单据金额累加
                mcurBill应收 = mcurBill应收 + CalcBillToTal(True)
                mcurBill实收 = mcurBill实收 + CalcBillToTal
                mcurBill应缴 = mcurBill应缴 + mobjBill.Pages(1).应缴金额
                
                mintBillNO = mintBillNO + 1
                For i = 1 To mshMoney.Rows - 1
                    If mshMoney.TextMatrix(i, 0) = "" Then Exit For
                Next
                mintMoneyRow = i - 1
            End If
            
            Call ClearRows: Call Bill.ClearBill
            Call NewBill(, False) '不设置费别
            txtPatient.SetFocus
        Else '读取修改
            Unload Me
        End If
    ElseIf chkCancel.Value = 1 Then '退单据状态
        If mstrInNO = "" Then
            MsgBox "没有正确读取单据内容,不能执行该操作！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        If gblnBillPrint Then
            If gobjBillPrint.zlEraseBill("'" & mstrInNO & "'", 0) = False Then Exit Sub
        End If
        
        strSQL = "Zl_门诊简单收费_Delete('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, "'" & mstrInNO & "'")
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        
        If mbytInState = 0 And gbln累计 Then txt累计.Text = Format(GetChargeTotal, "0.00")
        
        mstrInNO = "": cboNO.Text = "": txtInvoice.Text = ""
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney
        chkCancel.Value = Unchecked
        Call ClearPatientInfo
        txt合计.Text = gstrDec: txt应收.Text = gstrDec
        Call NewBill
        Call SetDisible(True)
        txtPatient.SetFocus
    ElseIf Not Bill.Active Then '收取划价单费用状态
        If mstrInNO = "" Then
            MsgBox "没有正确读取单据！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        If txtPatient.Text = "" Then
            MsgBox "没有发现病人信息,请输入病人信息！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用时间！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    
        If cbo开单科室.ListIndex = -1 Or cbo开单科室.Text = "" Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            cbo开单科室.SetFocus: Exit Sub
        End If
        If cbo结算方式.ListIndex = -1 Then
            MsgBox "请确定收费的结算方式！", vbInformation, gstrSysName
            cbo结算方式.SetFocus: Exit Sub
        End If
        If Val(txt缴款.Text) <> 0 And txt缴款.Enabled Then
            If Val(txt缴款.Text) < Val(txt应缴.Text) Then
                MsgBox "病人缴款金额不足，请补足应缴金额！", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt缴款): txt缴款.SetFocus: Exit Sub
            End If
        End If
        '刘兴洪:22343,缴款金额控制
        Select Case gTy_Module_Para.byt缴款控制
        Case 1  '1-代表输入缴款后才结束病人累计
        Case 2  '2-收费时必须要输入缴款金额
            If Val(txt应缴.Text) > 0 And Val(txt缴款.Text) = 0 Then
                MsgBox "注意:" & vbCrLf & _
                "    该病人未输入缴款金额,不能进行收费!", vbInformation + vbDefaultButton1, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                Exit Sub
            End If
        Case Else   ',0-代表不进行缴款输入和累计控制
        End Select
        
                    
        '票据号码检查,工本费打印检查
        mblnPrint = True
        '检查是否打印票据
        If mintInvoicePrint = 0 Then
            mblnPrint = False
        Else
            If mintInvoicePrint = 2 Then
                If MsgBox("是否打印票据？" & vbCrLf & "要取消此提示,请在本地参数中设置票据打印控制参数!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
        
        If mblnPrint Then
            If gblnStrictCtrl Then
                If Trim(txtInvoice.Text) = "" Then
                    MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
                If zlGetInvoiceGroupUseID(mlng领用ID, 1, txtInvoice.Text) = False Then
                    Exit Sub
                End If
                '并发操作检查,票号是否已用
                If CheckBillRepeat(mlng领用ID, 1, txtInvoice.Text) Then
                    MsgBox "票据号""" & txtInvoice.Text & """已经被使用，请重新输入。", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            Else
                If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                    MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            End If
        End If
        
        If cbo医疗付款.ListIndex <> -1 Then
            str医疗付款 = Mid(cbo医疗付款.Text, 1, InStr(1, cbo医疗付款, "-") - 1)
        End If
        
        lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
        '缴款结算
        strSQL = "zl_划价收费记录_INSERT('" & mstrInNO & "'," & Val(txtPatient.Tag) & "," & gint病人来源 & ",'" & _
            str医疗付款 & "','" & txtPatient.Text & "'," & _
            "'" & zlStr.NeedName(cboSex.Text) & "','" & mobjBill.年龄 & "'," & _
             ZVal(mobjBill.科室ID, , cbo开单科室.ItemData(cbo开单科室.ListIndex)) & "," & _
            cbo开单科室.ItemData(cbo开单科室.ListIndex) & ",'" & zlStr.NeedName(cbo开单人.Text) & "'," & _
            "'" & zlStr.NeedName(cbo结算方式.Text) & "|" & mobjBill.Pages(1).应缴金额 & "| ',"
        '预交结算
        If Val(txt预交冲款.Text) <> 0 Then
            strSQL = strSQL & mobjBill.Pages(1).冲预交额 & ","
        Else
            strSQL = strSQL & "NULL,"
        End If
        strSQL = strSQL & "NULL," & lng结帐ID & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','Z',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1)"
        
        On Error GoTo errH
        gcnOracle.BeginTrans
        '收取划价单费用
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '误差处理部份
        If mobjBill.Pages(1).误差金额 <> 0 Then
            '  Zl_简单收费误差_Insert(
            '  No_In         门诊费用记录.No%Type,
            '  病人id_In     门诊费用记录.病人id%Type,
            '  结帐id_In     门诊费用记录.结帐id%Type,
            '  误差金额_In   门诊费用记录.实收金额%Type,
            '  登记时间_In   门诊费用记录.登记时间%Type,
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type
            strSQL = "Zl_简单收费误差_Insert('" & mstrInNO & "'," & Val(txtPatient.Tag) & "," & lng结帐ID & "," & _
                mobjBill.Pages(1).误差金额 & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
                "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        If mblnPrint Then '打印门诊收据
            Call frmPrint.ReportPrint(1, "'" & mstrInNO & "'", "", "", mlng领用ID, mlngShareUseID, txtInvoice.Text, _
                zlDatabase.Currentdate, , , , mintInvoiceFormat, , , mstrUseType, , , , mstr普通价格等级)
        End If
        '费用清单的打印
        If InStr(mstrPrivs, "打印清单") > 0 Then
            If gint收费清单 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mstrInNO & "'", "药品单位=0", 2)
            ElseIf gint收费清单 = 2 Then
                If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mstrInNO & "'", "药品单位=0", 2)
                End If
            End If
        End If
        
        '是否可以连续收费：
        '使用预交款结算,当次收费结束(除非设置仅缴款结束参数)
        '如已缴款,则强行作为病人收费结束
        
        '刘兴洪:22343
        If CCur(txt缴款.Text) <> 0 Or (Val(txt预交冲款.Text) <> 0 And gTy_Module_Para.byt缴款控制 <> 1) Then
            mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
            mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
            Call ClearPatientInfo
            txt合计.Text = gstrDec: txt应收.Text = gstrDec
        Else
            mstrPrePati = txtPatient.Text
            mcurBill应收 = mcurBill应收 + CalcBillToTal(True)
            mcurBill实收 = mcurBill实收 + CalcBillToTal
            mcurBill应缴 = mcurBill应缴 + mobjBill.Pages(1).应缴金额
            
            mintBillNO = mintBillNO + 1
            For i = 1 To mshMoney.Rows - 1
                If mshMoney.TextMatrix(i, 0) = "" Then Exit For
            Next
            mintMoneyRow = i - 1
        End If
        
        mstrInNO = ""
        
        Call SetDisible(True)
        Call ClearRows: Call Bill.ClearBill
        Call NewBill
        
        mstrPreUnit = ""
        Call cbo开单科室_Click '删除无效开单人
        
        txtPatient.SetFocus
    End If
    gblnOK = True
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mstrInNO <> "" Then
        Bill.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strPre As String, strTmp As String
    Dim lngPre As Long, i As Long
    mblnStartFactUseType = zlStartFactUseType(1)
    Call RestoreWinState(Me, App.ProductName)
    Me.Width = 10000: Me.Height = 7770
    Call initCardSquareData
    If IsCheck误差费() = False Then Unload Me: Exit Sub
    
    'LED初始化
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModul, gcnOracle
    End If
    
    gblnOK = False
    mstrPrePati = "": mcurBill实收 = 0: mcurBill应收 = 0
    mstr付款方式 = ""
    mstrPreUnit = ""
    mblnDo = True
    mbln不重算价格 = False
    txt应收.Text = gstrDec: txt合计.Text = gstrDec
    
    Set mobjBill = New ExpenseBill
    Set mrsInfo = New ADODB.Recordset
    
    '查看功能时，无需初始数据
    If mbytInState = 0 Or mbytInState = 2 Then
        If mbytInState = 0 Then
            mstr药品价格等级 = gstr药品价格等级
            mstr卫材价格等级 = gstr卫材价格等级
            mstr普通价格等级 = gstr普通价格等级
        End If
        If Not InitData Then Unload Me: Exit Sub
    Else
        '年龄单位
        cbo年龄单位.AddItem "岁"
        cbo年龄单位.AddItem "月"
        cbo年龄单位.AddItem "天"
        cbo年龄单位.ListIndex = 0
    End If
    
    Call InitFace
    
    If mbytInState = 1 Or mbytInState = 2 Then '浏览、调整
        If Not ReadBill(mstrInNO) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
    Else '新增单据
        If mbytInState = 0 And gbln累计 Then txt累计.Text = Format(GetChargeTotal, "0.00")
        
        If Not NewBill(IIf(mblnStartFactUseType, False, True)) Then Unload Me: Exit Sub
        
        '读取该单据的内容
        If mstrInNO <> "" Then '修改原单据
            Set mobjBill = ImportBill(mstrInNO, True, 0, , , True, mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
            If mobjBill.NO = "" Then
                MsgBox "读取单据失败。", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                txtPatient.BackColor = &HE0E0E0
                cboNO.Text = mobjBill.NO
                
                mbln不重算价格 = True               '在费别_click事件中不重算价格
                Call Set开单人开单科室(mobjBill.Pages(1).开单人, mobjBill.Pages(1).开单部门ID)
                                    
                If mobjBill.病人ID <> 0 Then
                    txtPatient.Text = "-" & mobjBill.病人ID
                    Call txtPatient_KeyPress(13)
                Else
                    txtPatient.Text = mobjBill.姓名
                    cboSex.ListIndex = cbo.FindIndex(cboSex, mobjBill.性别, True)
                    Call LoadOldData(mobjBill.年龄, txt年龄, cbo年龄单位)
                    
                    '如果整张单据费别相同,则定位
                    strTmp = GetBill费别(mobjBill)
                    If strTmp <> "" Then
                        cbo费别.ListIndex = cbo.FindIndex(cbo费别, strTmp, True)
                        If cbo费别.ListIndex = -1 Then
                           cbo费别.AddItem strTmp
                           cbo费别.ListIndex = cbo费别.NewIndex
                        End If
                    End If
                    If cbo费别.ListIndex = -1 And cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
                    
                    If gint病人来源 <> 2 Then cbo医疗付款.ListIndex = GetCboIndexByCode(cbo医疗付款, "" & mobjBill.床号)
                End If
                mbln不重算价格 = False
                
                Bill.ClearBill
                Bill.Rows = mobjBill.Pages(1).Details.Count + 1
                '针对列编辑性质设置颜色
                Bill.SetColColor 0, &HE7CFBA
                Bill.SetColColor 1, &HE7CFBA
                Bill.SetColColor 3, &HE7CFBA
                txtDate.Text = Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss")
                chk加班.Value = mobjBill.加班标志
                chkCancel.Enabled = False
                
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                '缺省为原单据的结算方式
                strTmp = GetBalanceName(mstrInNO)
                If strTmp <> "" Then
                    i = cbo.FindIndex(cbo结算方式, strTmp, True)
                    If i <> -1 Then cbo结算方式.ListIndex = i
                End If
                
                Call ShowDetails
                Call ShowMoney
                
                '修改单据时不允许修改病人信息
                txtPatient.Locked = True
                Call ReInitPatiInvoice
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    
    mbytInState = Empty
    mstrInNO = Empty
    mintBillNO = 0: mintMoneyRow = 0
    mlng领用ID = 0
    mstrCardNO = ""
    mstrDelete = ""
    zlCommFun.OpenIme False
    mblnNOMoved = False
    mintInvoicePrint = 0
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrs费别 = Nothing
    Call initCardSquareData
    'LED初始化
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
   If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call txtPatient_KeyPress(vbKeyReturn)
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long
    If txtPatient.Locked Then Exit Sub
    mblnNotClick = True
    lngPreIDKind = IDKind.IDKind
    IDKind.IDKind = IDKind.GetKindIndex(objCard.名称)
    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Not gbln简码切换 Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        zlDatabase.SetPara "简码方式", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
        gbytCode = Val(zlDatabase.GetPara("简码方式", , , True))
    End If
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        mobjBill.发生时间 = CDate(txtDate.Text)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    cboNO.SelStart = 0
    cboNO.SelLength = Len(cboNO.Text)
    If (mobjBill.Pages(1).Details.Count = 0 And mbytInState = 0) Or chkCancel.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, blnNull As Boolean
    Dim strOper As String, vDate As Date, i As Integer
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        txt找补.Text = "0.00"
        txt缴款.Text = "0.00"
        txt应缴.Text = "0.00"
    
        cboNO.Text = GetFullNO(cboNO.Text, 13)
        If chkCancel.Value = 1 Then
            '是否已转入后备数据表中
            If zlDatabase.NOMoved("门诊费用记录", cboNO.Text, , "1") Then
                If Not ReturnMovedExes(cboNO.Text, 1, Me.Caption) Then Exit Sub
                mblnNOMoved = False
            End If
        
            '单据退费权限判断
            If Not ReadBillInfo(1, cboNO.Text, 1, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If InStr(mstrPrivs, "所有操作员") <= 0 Then
                If UserInfo.姓名 <> strOper Then
                    MsgBox "你没有""所有操作员""权限,不能对" & strOper & "的单据进行退费!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(2, strOper, vDate, "退费", cboNO.Text, , 1) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '退费时,单据必须为简单收费的单据
            If Not isSimple(cboNO.Text) Then
                MsgBox "该单据不存在或不是简单收费单据！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
                        
            '是否已执行
            i = BillCanDelete(cboNO.Text, 1)
            If i <> 0 Then
                Select Case i
                    Case 1 '该单据不存在
                        MsgBox "指定的单据不存在！", vbInformation, gstrSysName
                    Case 2 '已经全部完全执行
                        MsgBox "该单据中的项目已经全部完全执行！", vbInformation, gstrSysName
                    Case 3 '未完全执行部分剩余数量为0
                        MsgBox "该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
                
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        If chkCancel.Value = 1 Then '读取退费单
            blnRead = ReadBill(cboNO.Text, False, True)
        ElseIf mobjBill.Pages(1).Details.Count = 0 Then '读取划价单
            blnRead = ReadBill(cboNO.Text, True, False, blnNull)
        End If
        
        lbl动态费别.Visible = blnRead
        If blnRead Then
            '显示动态费别:显示退费单或显示要收取的划价单时
            cbo费别.Locked = True
            cbo费别.Visible = False
            lbl动态费别.BorderStyle = 1
            lbl动态费别.Left = cbo费别.Left
            lbl动态费别.Width = cbo医疗付款.Left - cboSex.Left
        
            mstrInNO = cboNO.Text '确定时以mstrInNO为准
            If chkCancel.Value = 0 Then '划价单
                chk加班.Enabled = False
                Bill.Active = False
                
                If gint病人来源 = 1 And InStr(mstrPrivs, "允许非医保病人") = 0 Then
                     Call ClearPatientInfo
                End If
                                
                If txtPatient.Text = "" Or blnNull Then
                    txtPatient.SetFocus
                Else
                    If txt缴款.Visible Then
                        txt缴款.SetFocus
                    Else
                        cmdOK.SetFocus
                    End If
                End If
            Else '退
                Call SetDisible 'cboNO在获取焦点后unLock
                '部分退费只支持退费指定结算方式
                cbo结算方式.Locked = False
                cmdOK.SetFocus
            End If
        Else
            Call ClearPatientInfo: Call ClearMoney: Call ClearRows
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Enabled = False Or txtPatient.Locked Then Exit Sub
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)

    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txt年龄_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt年龄.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt年龄.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt年龄_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt年龄.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True
    End If
    
    If mbytInState = 0 Then mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
End Sub

Private Sub txt应缴_Change()
    If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00": txt找补.Text = "0.00": Exit Sub
    txt找补.Text = Format(Val(txt缴款.Text) - Val(txt应缴.Text), "0.00")
End Sub

Private Sub txt预交冲款_GotFocus()
    Call txt缴款_GotFocus '收费自动产生工本费
    zlControl.TxtSelAll txt预交冲款
End Sub

Private Sub txt预交冲款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt预交冲款_Validate(Cancel As Boolean)
    Dim curTotal As Currency
    
    curTotal = CalcBillToTal
    If txt预交冲款.Text = "" Then
        txt预交冲款.Text = "0.00"
    ElseIf Not IsNumeric(txt预交冲款.Text) And txt预交冲款.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) < 0 Then
        MsgBox "预交款冲款金额不能为负！", vbInformation, gstrSysName
        If curTotal < 0 Then
            txt预交冲款.Text = "0.00"
        Else
            txt预交冲款.Text = Format(IIf(curTotal > Val(sta.Panels(3).Tag), Val(sta.Panels(3).Tag), curTotal), "0.00")
        End If
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) > 0 And curTotal < 0 Then
        MsgBox "单据应付金额为负时不能使用预交款！", vbInformation, gstrSysName
        txt预交冲款.Text = "0.00"
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) > Val(sta.Panels(3).Tag) Then
        MsgBox "预交款冲款金额不能超过病人的预交余额:" & Format(Val(sta.Panels(3).Tag), "0.00") & " ！", vbInformation, gstrSysName
        If curTotal < 0 Then
            txt预交冲款.Text = "0.00"
        Else
            txt预交冲款.Text = Format(IIf(curTotal > Val(sta.Panels(3).Tag), Val(sta.Panels(3).Tag), curTotal), "0.00")
        End If
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) > Format(curTotal, "0.00") And Val(txt预交冲款.Text) <> 0 Then
        MsgBox "预交款冲款金额不能大于应付金额:" & Format(curTotal, "0.00") & " ！", vbInformation, gstrSysName
        If curTotal < 0 Then
            txt预交冲款.Text = "0.00"
        Else
            txt预交冲款.Text = Format(IIf(curTotal > Val(sta.Panels(3).Tag), Val(sta.Panels(3).Tag), curTotal), "0.00")
        End If
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    Else
        txt预交冲款.Text = Format(txt预交冲款.Text, "0.00")
        
        '重新计算应缴，误差(分币)等
        If Bill.Active Then
            Call ShowMoney
        Else
            Call ShowPrice
        End If
    End If
End Sub

Private Sub txtInvoice_GotFocus()
    zlControl.TxtSelAll txtInvoice
End Sub

Private Sub txtInvoice_LostFocus()
'    If Not txtInvoice.Locked And txtInvoice.Text <> "" Then
'        txtInvoice.Text = Format(Left(txtInvoice.Text, gbytFactLength), String(gbytFactLength, "0"))
'    End If
End Sub

Private Sub txt年龄_Gotfocus()
    txt年龄.SelStart = 0
    txt年龄.SelLength = Len(txt年龄.Text)
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            Call cbo年龄单位.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    zlCommFun.OpenIme True
    If txtPatient.Enabled = False Or txtPatient.Locked Then Exit Sub
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        .ColData(1) = 5 '应收缺省跳过,当项目变价时,设为输入(4)
        '针对列编辑性质设置颜色
        .SetColColor 0, &HE7CFBA
        .SetColColor 1, &HE7CFBA
        .SetColColor 3, &HE7CFBA
    End With
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboSex.ListIndex <> -1 Then mobjBill.性别 = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If cboSex.Locked Then Exit Sub
    If SendMessage(cboSex.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
   
    If cbo费别.Locked Then Exit Sub
    
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo费别.hWnd, KeyAscii)
        If lngIdx = -1 And cbo费别.ListCount > 0 Then lngIdx = 0
        cbo费别.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 And cbo费别.ListIndex <> -1 Then
        mobjBill.费别 = Mid(cbo费别.Text, InStr(cbo费别.Text, "-") + 1)
        If mbytInState = 0 And mstrInNO <> "" And mobjBill.Pages(1).Details.Count > 0 Then
            '重新计算价格
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        mblnCboClick = False    '先用鼠标在下拉列表选择一个并点击,不要移开,此时只触发click,再输入简码并且回车,不触发click,所以需要在此赋值,以便validate事件中强行调用click事件
        Call zlCommFun.PressKey(vbKeyTab)
        
    ElseIf KeyAscii >= 32 And Not cbo开单科室.Locked Then
        lngIdx = zlControl.CboMatchIndex(cbo开单科室.hWnd, KeyAscii)
        If lngIdx = -1 And cbo开单科室.ListCount > 0 Then lngIdx = 0
        cbo开单科室.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        If cbo开单人.Locked Then Exit Sub
        
        strText = cbo开单人.Text
        If cbo开单人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo开单人.List(cbo开单人.ListIndex) Then Call zlControl.CboSetIndex(cbo开单人.hWnd, -1)
        End If
        If strText = "" Then
            cbo开单人.ListIndex = -1
        ElseIf cbo开单人.ListIndex = -1 Then
            intIdx = -1
            For i = 0 To cbo开单人.ListCount - 1
                If UCase(cbo开单人.List(i)) Like UCase(strText) & "*" Then
                    If intIdx = -1 Then cbo开单人.ListIndex = i
                    intIdx = i
                End If
            Next
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo开单人_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo开单人.ListIndex = -1 Then
            cbo开单人.Text = ""
            mobjBill.Pages(1).开单人 = ""
            If gbyt科室医生 = 0 Or gbln必须输开单人 Then Exit Sub
        Else
            mobjBill.Pages(1).开单人 = zlStr.NeedName(cbo开单人.Text)
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
        Case vbKeyF1  '帮助
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is txtPatient Then
                Call txtPatient_Validate(False)
                Me.Refresh
            End If
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF6 '定位到病人输入框
            txtPatient.SetFocus
            Call zlControl.TxtSelAll(txtPatient)
        Case vbKeyF7 '切换输入法
            If Not gbln简码切换 Then Exit Sub
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF8 '退(自动激活事件)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyF9 '定位到单据号输入框
            cboNO.SetFocus
            Call zlControl.TxtSelAll(cboNO)
        Case vbKeyF12
            If Shift = 2 Then
                '强制性LED报价,(合计)
                If gblnLED And (Bill.Active Or (Not Bill.Active And chkCancel.Value = 0)) _
                    And txt缴款.Enabled And txt缴款.Visible And CCur(txt合计.Text) <> 0 Then
                    mblnHotKey = True: txt缴款.SetFocus
                    If ActiveControl Is txt缴款 Then txt缴款_GotFocus
                End If
            End If
        Case vbKeyEscape
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
    lngW = mshMoney.Width - 75
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    
    mshMoney.ColWidth(0) = 600
    
    lngW = lngW - mshMoney.ColWidth(0)
    
    mshMoney.ColWidth(1) = lngW * 0.45
    mshMoney.ColWidth(2) = lngW * 0.55
    
    mshMoney.ColAlignment(0) = 4
    mshMoney.ColAlignment(1) = 1
    mshMoney.ColAlignment(2) = 7
    
    mshMoney.TextMatrix(0, 0) = "序号"
    mshMoney.TextMatrix(0, 1) = "项目"
    mshMoney.TextMatrix(0, 2) = "金额"
    mshMoney.Row = 0
    mshMoney.Col = 0: mshMoney.CellAlignment = 4
    mshMoney.Col = 1: mshMoney.CellAlignment = 4
    mshMoney.Col = 2: mshMoney.CellAlignment = 4
    
    mshMoney.MergeCol(0) = True
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
        
    '自动识别加班
    If OverTime(zlDatabase.Currentdate) Then chk加班.Value = Checked
    
    '年龄单位
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    
    '可选性别
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboSex.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then cboSex.ListIndex = cboSex.NewIndex
            rsTmp.MoveNext
        Next
    End If
    
    '可选医疗付款方式
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 医疗付款方式 Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo医疗付款.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo医疗付款.ListIndex = cbo医疗付款.NewIndex
                mstr付款方式 = rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '结算方式
    Set rsTmp = Get结算方式("收费")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            '只加入非医保的结算方式供选择
            If InStr(",1,2,", rsTmp!性质) > 0 Then
                cbo结算方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
                cbo结算方式.ItemData(cbo结算方式.NewIndex) = rsTmp!性质
                
                If rsTmp!名称 = gstr结算方式 Then
                    cbo结算方式.ListIndex = cbo结算方式.NewIndex
                End If
                
                If rsTmp!缺省 = 1 And cbo结算方式.ListIndex = -1 Then
                    cbo结算方式.ListIndex = cbo结算方式.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cbo结算方式.ListCount = 0 Then
        MsgBox "收费场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
   
    
    '不缺省开单人和开单科室
    Call FillDept
    If cbo开单科室.ListCount = 0 Then
        MsgBox "没有可用的开单科室,可用的开单科室须满足以下规则:" & vbCrLf & _
               "    1.部门性质为产科" & vbCrLf & _
               "    2.或者部门性质为临床,并且部门服务于门诊和住院、仅服务于门诊(病人来源为门诊病人)或仅服务于住院(病人来源为住院病人).", vbInformation, gstrSysName
        Exit Function
    End If
    zlControl.CboSetWidth cbo开单科室.hWnd, 2500
    Call FillDoctor
    If cbo开单人.ListCount = 0 Then
        MsgBox "没有可用的开单人,可用的开单人须满足以下规则:" & vbCrLf & _
               "    1.人员性质为医生或护士," & vbCrLf & _
               "    2.并且,人员所在部门性质为临床" & vbCrLf & _
               "    3.并且,人员所在部门服务于门诊和住院、仅服务于门诊(病人来源为门诊病人)或仅服务于住院(病人来源为住院病人)." & vbCrLf & _
               "    4.护士是否允许做为可用开单人须满足以下规则:" & vbCrLf & _
               "      本地参数开单人允许为护士,并且本地参数的可用收费类别包含卫材,材料,治疗", vbInformation, gstrSysName
        Exit Function
    End If
    
    '执行部门
    Set mrsUnit = GetDepartments("", gint病人来源 & ",3")
    If mrsUnit.EOF Then
        MsgBox "没有初始化部门信息,单据无法处理执行部门。请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '开单日期
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    Set mrsInfo = New ADODB.Recordset
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFace()
'功能：根据表单要完成的功能设置界面布局
    Dim arrHead() As String, i As Integer
    
    lblTitle.Caption = gstrUnitName & "病人收费单"
    
    '公用单据表格式
    With Bill
        .LocateCol = 0
        .PrimaryCol = 0
        .Font.Size = 11
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        arrHead = Split(STR_HEAD, ";")
        .COLS = UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
        If mbytInState = 0 Then
            .ColData(0) = 1 '项目输入,按扭可选
            .ColData(1) = 5 '应收金额缺省跳过,当项目变价时,设为输入(4)
            .ColData(2) = 5 '实收金额跳过
            .ColData(3) = 3 '默认取开单科室或上一科室
            .ColData(4) = 5
            
            .SetColColor 0, &HE7CFBA
            .SetColColor 1, &HE7CFBA
            .SetColColor 3, &HE7CFBA
            
            ReDim marrColData(.COLS - 1)
            For i = 0 To .COLS - 1
                marrColData(i) = .ColData(i)
            Next
        End If
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name)
    
    '读取简码匹配方式
    sta.Panels("PY").Visible = mbytInState = 0 And gbln简码切换 '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln简码切换
    If mbytInState = 0 Then
        '简码匹配方式：0-拼音,1-五笔,2-两者
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If
    
    If mbytInState = 1 Or mbytInState = 2 Then
        cbo费别.Visible = False
        lbl动态费别.Left = cbo费别.Left
        lbl动态费别.Visible = True
    Else
        lbl动态费别.BorderStyle = 0
        lbl动态费别.AutoSize = True
    End If
    
    Call SetMoneyList
    
    '权限设置
    If mbytInState = 0 Then
        If InStr(mstrPrivs, "门诊退费") = 0 Then
            chkCancel.Visible = False
            lblFact.Left = lblFact.Left + chkCancel.Width
            txtInvoice.Left = txtInvoice.Left + chkCancel.Width
            lbl单据号.Left = lbl单据号.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
        End If
        
        If InStr(mstrPrivs, "重打票据") = 0 Then
            lblRePrint.Visible = False
            txtRePrint.Visible = False
        End If
        txtInvoice.Locked = Not (InStr(1, mstrPrivs, "修改票据号") > 0) And gblnStrictCtrl
            
        If Not gbln累计 Then
            lbl累计.Visible = False
            txt累计.Visible = False
            lbl应收.Top = lbl应收.Top + txt累计.Height / 3
            txt应收.Top = txt应收.Top + txt累计.Height / 3
            lbl合计.Top = lbl合计.Top + txt累计.Height / 1.5
            txt合计.Top = txt合计.Top + txt累计.Height / 1.5
        End If
    Else
        lbl应缴.Visible = False
        txt应缴.Visible = False
        lbl缴款.Visible = False
        lbl找补.Visible = False
        txt缴款.Visible = False
        txt找补.Visible = False
        
        lbl累计.Visible = False
        txt累计.Visible = False
        
        lbl应收.Top = lbl应收.Top + txt累计.Height / 3
        txt应收.Top = txt应收.Top + txt累计.Height / 3
        lbl合计.Top = lbl合计.Top + txt累计.Height / 1.5
        txt合计.Top = txt合计.Top + txt累计.Height / 1.5
        
        txt预交冲款.Top = txt应收.Top
        lblDeposit.Top = txt预交冲款.Top + (txt预交冲款.Height - lblDeposit.Height) / 2
        
        fraTitle.Enabled = False
        
        lblRePrint.Visible = False
        txtRePrint.Visible = False
        
        chkCancel.Visible = False
        If mstrDelete <> "" Then
            lblFlag.Visible = True
        Else
            lblFact.Left = lblFact.Left + chkCancel.Width
            txtInvoice.Left = txtInvoice.Left + chkCancel.Width
            lbl单据号.Left = lbl单据号.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
        End If
        
        Call SetDisible
        
        If mbytInState = 2 Then
            txtDate.Enabled = True
            cbo开单人.Locked = False
        Else
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
            cmdCancel.Top = cmdCancel.Top - cmdCancel.Height / 2
        End If
    End If
    
    '输入控制
    If Not gbln性别 Then cboSex.TabStop = False
    If Not gbln年龄 Then txt年龄.TabStop = False: cbo年龄单位.TabStop = False
    If Not gbln费别 Then cbo费别.TabStop = False
    If Not gbln加班 Then chk加班.TabStop = False
    If Not gbln开单日期 Then txtDate.TabStop = False
    If Not gbln开单人 Then cbo开单人.TabStop = False
    If Not gbln医疗付款 Then cbo医疗付款.TabStop = False
       
    If gbyt科室医生 = 0 Then
        Call ExChangeLocate(cbo开单科室, cbo开单人)
        lbl科室.Caption = "开单人"
        lbl开单人.Caption = "开单科室"
        cbo开单科室.TabStop = False
    End If
    
    '82801,冉俊明,2015-2-26
    txt年龄.MaxLength = zlGetPatiInforMaxLen.intPatiAge
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'界面设置为不可修改状态
    cboNO.Locked = Not bln
    txtPatient.Locked = Not bln
    cboSex.Locked = Not bln
    txt年龄.Locked = Not bln
    cbo年龄单位.Locked = Not bln
    
    cbo费别.Locked = Not bln
    cbo医疗付款.Locked = Not bln
    
    cbo开单科室.Locked = Not bln
    cbo开单人.Locked = Not bln
    cbo开单科室.Enabled = bln
    cbo开单人.Enabled = bln
    
    chk加班.Enabled = bln
    cbo结算方式.Locked = Not bln
    txtDate.Enabled = bln
    fraStat.Enabled = bln
    Bill.Active = bln
    
    If Not bln Then
        txt缴款.BackColor = &HE0E0E0
        txtPatient.BackColor = &HE0E0E0
        txt年龄.BackColor = &HE0E0E0
    Else
        txt缴款.BackColor = &HFFFFFF
        txtPatient.BackColor = &HFFFFFF
        txt年龄.BackColor = &HFFFFFF
    End If
End Sub

Private Function IsRegisterDept() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否通过挂号单读取的病人
    '返回:是返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-19 15:31:01
    '问题:34032
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    For i = mrsInfo.Fields.Count - 1 To 0 Step -1
        If UCase(mrsInfo.Fields(i).Name) = "执行部门ID" Then
            IsRegisterDept = True: Exit Function
        End If
    Next
End Function

Private Sub SetDeptDoctorByRegevent(ByVal lng病人ID As Long, _
    Optional strRegNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID或挂号单中病人的挂号科室和医生信息设置开单科室和开单人
    '入参:lng病人ID-病人ID
    '     strRegNO-挂号单号
    '编制:刘兴洪
    '日期:2014-06-06 17:38:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    strTmp = zlGetRegEventsCons("加班标志")
    If strRegNO <> "" Then
        strTmp = strTmp & " And NO=[2]"
    Else
        strTmp = strTmp & " And 病人ID=[1]"
    End If
    
    strSQL = "Select 执行部门id, 执行人" & vbNewLine & _
            "From (Select 执行部门id, 执行人, 登记时间" & vbNewLine & _
            "       From 门诊费用记录" & vbNewLine & _
            "       Where 记录性质 = 4 And 记录状态 = 1 " & strTmp & vbNewLine & _
            "       Order By 登记时间 Desc)" & vbNewLine & _
            "Where Rownum < 2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, strRegNO)
    If Not rsTmp.EOF Then
        Call Set开单人开单科室Click("" & rsTmp!执行人, Val("" & rsTmp!执行部门ID))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowWelcomeByLed()
'功能:显示欢迎信息
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModul, gcnOracle
    End If

End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim i As Integer, intNum As Long
    Dim rsTemp As ADODB.Recordset
    Dim lng病人ID As Long, strPati As String
    Dim objDetail As Detail, blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean, blnIDCard As Boolean
    
    On Error Resume Next
    
    If txtPatient.Locked Then Exit Sub
    
    '问题:51488
    If (IDKind.Cards.读卡快键 = "空格键" Or IDKind.Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub

    '特殊字符过滤在Form_KeyPress中进行
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
     Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    '按了回车或刷卡执行本过程后就不再执行Validate事件
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 Then
        mblnKeyPress = True
    Else
        mblnKeyPress = False   '刷就诊卡时不会调用validate事件来设置此变量,所以需要这里设置
    End If
    
    
    '正常输入病人(姓名各种标识)部份:住院病人收费时可弹出选择器'@
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And gint病人来源 = 2 And mbytInState = 0 And txtPatient.Text = "" And Not mblnValid Then
        frmPatiSelect.Show 1, Me
        If frmPatiSelect.mlngPatient = 0 Then Exit Sub
        txtPatient.Text = "-" & frmPatiSelect.mlngPatient
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If gint病人来源 = 1 And InStr(mstrPrivs, "允许非医保病人") = 0 Then
            Call ClearPatientInfo: Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '病人未改变退出
        If mrsInfo.State = 1 Then
            If txtPatient.Text = mrsInfo!姓名 Then
                If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                 Exit Sub
            End If
        End If
        
        '读取病人信息
        txt找补.Text = "0.00": txt缴款.Text = "0.00": sta.Panels(2) = ""
        
        '收费保持病人ID
        If Val(txtPatient.Tag) <> 0 And txtPatient.Text = mstrPrePati Then
            strPati = "-" & Val(txtPatient.Tag)
        Else
            strPati = txtPatient.Text
        End If
        
        If IDKind.GetCurCard.名称 Like "IC卡*" And IDKind.GetCurCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        If IDKind.GetCurCard.名称 Like "*身份证*" And IDKind.GetCurCard.系统 Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        If Not GetPatient(strPati, blnCancel, blnCard) Then
            Call ReInitPatiInvoice
            If blnCancel Then '取消输入
                If Visible Then txtPatient.SetFocus
                txtPatient.Text = ""
                Exit Sub
            End If
            If blnCard Then
                MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
                Call ClearPatientInfo
                 Exit Sub
            ElseIf gint病人来源 = 1 And gblnInputName Then
                If mstrInNO = "" Then
                    If Not CheckRegisted(0) Then
                       txtPatient.Text = "": Exit Sub
                    End If
                End If
                
                sta.Panels(2) = "输入的标识不能读取病人信息，将默认为新病人姓名！"
                mobjBill.病人ID = 0: mobjBill.标识号 = 0: mobjBill.主页ID = 0
                txtPatient.PasswordChar = ""
                '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
                txtPatient.IMEMode = 0
                If mstrInNO <> "" And Not Bill.Active Then
                    '划价时改姓名也不处理费别
                    cbo费别.Locked = True '实际此时已不可见
                Else
                    cbo费别.Locked = False
                    If Not mblnValid Then '同一个病人不重置费别
                        If Not (Bill.Active And txtPatient.Text = mstrPrePati) Then Call LoadAndSeek费别
                    End If
                End If
                
                cbo医疗付款.Locked = False
                
                '预交信息初始
                lblDeposit.ForeColor = &H808080
                txt预交冲款.Enabled = False: txt预交冲款.ForeColor = &H808080: txt预交冲款.Text = "0.00"
                sta.Panels(3).Tag = "": sta.Panels(3).Text = "": sta.Panels(3).Visible = False
                txtPatient.Tag = ""
                If Bill.Active Then
                    If txtPatient.Text = mstrPrePati Then
                        mobjBill.姓名 = txtPatient.Text
                        mobjBill.性别 = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
                        mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
                        mobjBill.费别 = Mid(cbo费别.Text, InStr(cbo费别.Text, "-") + 1)
                        If Not mblnValid Then Bill.SetFocus
                         Exit Sub
                    Else
                        '不同的收费病人
                        '清除医生科室
                        If gbyt科室医生 = 0 And mstrInNO = "" Then
                            cbo开单人.ListIndex = -1: cbo开单科室.ListIndex = -1
                            mobjBill.Pages(1).开单人 = "":  mobjBill.Pages(1).开单部门ID = 0
                        End If
                        
                        sta.Panels(3).Tag = "": sta.Panels(3).Text = "": sta.Panels(3).Visible = False
                        mobjBill.姓名 = txtPatient.Text
                        txt年龄.Text = ""
                        
                        '仅以缴款作为结束时,即使不同的病人也保留收费
                        '除非刚好缴款结束(mstrPrePati = "")
                        '刘兴洪:22343
                        If gTy_Module_Para.byt缴款控制 <> 1 Or mstrPrePati = "" Then
                            Call ClearMoney
                            mintBillNO = 0: mintMoneyRow = 0
                            mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
                            txt找补.Text = "0.00": txt缴款.Text = "0.00": txt应缴.Text = "0.00"
                            If mobjBill.Pages(1).Details.Count = 0 Then
                                mstrPrePati = ""
                                If mstrInNO = "" Then
                                    txt合计.Text = gstrDec: txt应收.Text = gstrDec
                                End If
                            Else
                                Call ShowMoney
                            End If
                        End If
                        If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                        
                        'LED初始化
                        If Not mblnValid Then ShowWelcomeByLed
                        
                        Exit Sub
                    End If
                Else
                    '空姓名划价单输入不同姓名
                    '刘兴洪:22343
                    If txtPatient.Text <> mstrPrePati And mstrPrePati <> "" And gTy_Module_Para.byt缴款控制 <> 1 Then
                        mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
                        
                        txt应收.Text = Format(CalcBillToTal(True), gstrDec)
                        txt合计.Text = Format(CalcBillToTal, gstrDec)
                        txt应缴.Text = Format(mobjBill.Pages(1).应缴金额, "0.00")
                        
                        '调整费用列表
                        For i = mshMoney.Rows - 1 To 1 Step -1
                            If mshMoney.TextMatrix(i, 0) <> "" Then
                                intNum = Val(mshMoney.TextMatrix(i, 0))
                                Exit For
                            End If
                        Next
                        If intNum > 1 Then
                            mintBillNO = 0
                            
                            mshMoney.Redraw = False
                            For i = mshMoney.Rows - 1 To 1 Step -1
                                If Val(mshMoney.TextMatrix(i, 0)) <> intNum Then
                                    mshMoney.RemoveItem i
                                Else
                                    mshMoney.TextMatrix(i, 0) = 1
                                End If
                            Next
                            If mshMoney.Rows < 5 Then mshMoney.Rows = 5
                            mshMoney.Redraw = True
                        End If
                    End If
                    If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                     Exit Sub
                End If
            Else
                MsgBox "不能确定病人信息！", vbInformation, gstrSysName
                Call ClearPatientInfo
                If Not mblnValid Then txtPatient.SetFocus
                 Exit Sub
            End If
        '正确读出了病人信息
        Else
            lng病人ID = Val("" & mrsInfo!病人ID)
            If mbytInState = 0 And mstrInNO = "" And gint病人来源 = 1 Then
                If Not CheckRegisted(lng病人ID) Then
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                End If
            End If
            '就诊卡密码检查
            If mbytInState = 0 And (blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.接口序号 <> 0) And mstrPassWord <> "" Then
                If Mid(gstrCardPass, 3, 1) = "1" Then
                    If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
                        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                    End If
                End If
            End If
            
            '102234,调用外挂部件接口
            If PatiValiedCheckByPlugIn(mlngModul, lng病人ID) = False Then
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
            End If
        
            '仅新增单据时,开单人和开单科室的处理
            '-----------------------------------------------------------------
            If mbytInState = 0 And mstrInNO = "" Then
                '不是同一个病人时清除医生
                If Not (Nvl(mrsInfo!姓名) = mstrPrePati And Nvl(mrsInfo!姓名) <> "") Then
                    If gbyt科室医生 = 0 And mstrInNO = "" Then
                        cbo开单人.ListIndex = -1
                        cbo开单科室.ListIndex = -1
                        mobjBill.Pages(1).开单人 = ""
                        mobjBill.Pages(1).开单部门ID = 0
                    End If
                End If
            
                '由挂号单得来时有执行部门
                If IsRegisterDept Then
                    If IsNull(mrsInfo!姓名) Then '没有建档,但挂了号,根据挂号单读开单人和开单科室
                        Call SetDeptDoctorByRegevent(0, txtPatient.Text)
                        sta.Panels(2) = "该病人挂号时没有登记档案,请输入病人姓名！"
                        Call ClearPatientInfo
                        
                        Set mrsInfo = New ADODB.Recordset
                        If Not mblnValid And Visible Then txtPatient.SetFocus
                        Exit Sub
                    Else
                        Call Set开单人开单科室Click(mrsInfo!执行人 & "", Val("" & mrsInfo!执行部门ID))
                    End If
                ElseIf gint病人来源 = 2 Then
                    If gbyt科室医生 <> 0 Then
                        '取住院病人的开单部门:科室确定医生或各自独立输入
                        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, Val("" & mrsInfo!当前科室id)))
                        Call cbo开单科室_Click
                        
                    End If
                ElseIf gint病人来源 = 1 Then
                    Call SetDeptDoctorByRegevent(lng病人ID) '试图查找病人的挂号科室和医生
                End If
            End If
            
            '预交信息
            Set rsTemp = GetMoneyInfo(lng病人ID, 0, False, 1, False, 0, True)
            Dim dbl病人余额 As Double, dbl家属余额 As Double, dbl可用预交 As Double
            Do While Not rsTemp.EOF
                If Nvl(rsTemp!家属, 0) = 0 Then
                    dbl病人余额 = Val(Nvl(rsTemp!预交余额)) - Val(Nvl(rsTemp!费用余额))
                Else
                    dbl家属余额 = Val(Nvl(rsTemp!预交余额)) - Val(Nvl(rsTemp!费用余额))
                End If
                dbl可用预交 = dbl可用预交 + (Val(Nvl(rsTemp!预交余额)) - Val(Nvl(rsTemp!费用余额)))
                rsTemp.MoveNext
            Loop
            sta.Panels(3).Tag = dbl可用预交
            sta.Panels(3).Text = "预交:" & Format(dbl病人余额 + dbl家属余额, "0.00") & _
                IIf(dbl家属余额 > 0, "(含家属:" & Format(dbl家属余额, "0.00") & ")", "")
                
            '预交信息初始
            If Val(sta.Panels(3).Tag) > 0 Then
                lblDeposit.ForeColor = 0
                txt预交冲款.Enabled = True
                txt预交冲款.ForeColor = 0
                txt预交冲款.Text = "0.00"
                sta.Panels(3).Visible = True
            Else
                lblDeposit.ForeColor = &H808080
                txt预交冲款.Enabled = False
                txt预交冲款.ForeColor = &H808080
                txt预交冲款.Text = "0.00"
                sta.Panels(3).Tag = ""
                sta.Panels(3).Text = ""
                sta.Panels(3).Visible = False
            End If
            
            txtPatient.Text = IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名)
            cboSex.ListIndex = cbo.FindIndex(cboSex, IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别), True)
            Call LoadOldData("" & mrsInfo!年龄, txt年龄, cbo年龄单位)
            If Not IsNull(mrsInfo!出生日期) Then
                 txt年龄.Text = ReCalcOld(mrsInfo!出生日期, cbo年龄单位, lng病人ID)
            End If
            
            If Not mblnValid Then
                If Not (mrsInfo.Fields(mrsInfo.Fields.Count - 1).Name = "执行部门ID" _
                    And cbo开单科室.ListIndex <> -1) Then
                    If mstrInNO <> "" And Not Bill.Active Then
                        '划价时改姓名也不处理费别
                        cbo费别.Locked = True '实际此时已不可见
                    Else
                        Call LoadAndSeek费别
                    End If
                Else
                    '挂号时确定的费别
                    cbo费别.ListIndex = cbo.FindIndex(cbo费别, IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别), True)
                End If
            End If
            
            cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, Nvl(mrsInfo!医疗付款方式), True)
            cbo医疗付款.Locked = gint病人来源 = 2 'Or (cbo医疗付款.ListIndex <> -1)

            If gstr费别 <> "" And cbo费别.ListIndex = -1 Then cbo费别.ListIndex = cbo.FindIndex(cbo费别, gstr费别, True)
            If mstr付款方式 <> "" And cbo医疗付款.ListIndex = -1 Then cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, mstr付款方式, True)

            txtPatient.PasswordChar = ""
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
            txtPatient.Tag = lng病人ID
            
            '填写对象中的病人信息
            With mobjBill
                .病人ID = lng病人ID
                .主页ID = Nvl(mrsInfo!主页ID, 0)
                .标识号 = IIf(gint病人来源 = 2, Nvl(mrsInfo!住院号, 0), Nvl(mrsInfo!门诊号, 0))
                .病区ID = Nvl(mrsInfo!当前病区ID, 0)
                .科室ID = Nvl(mrsInfo!当前科室id, 0)
                .床号 = "" & mrsInfo!当前床号
                .姓名 = txtPatient.Text
                .性别 = Nvl(mrsInfo!性别)
                .年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
                .费别 = zlStr.NeedName(cbo费别.Text) '以当前有效为准
            End With
            Call ReInitPatiInvoice
            If Bill.Active Then
                If txtPatient.Text = mstrPrePati And txtPatient.Text <> "" Then
                    '同一个病人
                    If Not mblnValid Then Bill.SetFocus
                     Exit Sub
                Else
                    '不同的病人
                    '不同的收费病人:如果仅以缴款作为结束,不同病人也保留收费
                    '除非刚好缴款结束(mstrPrePati = "")
                    '刘兴洪:22343
                    If gTy_Module_Para.byt缴款控制 <> 1 Or mstrPrePati = "" Then
                        Call ClearMoney
                        mintBillNO = 0: mintMoneyRow = 0
                        mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
                        If mobjBill.Pages(1).Details.Count = 0 Then
                            mstrPrePati = ""
                            If mstrInNO = "" Then
                                txt合计.Text = gstrDec: txt应收.Text = gstrDec
                            End If
                        Else
                            Call ShowMoney
                        End If
                    End If
                    
                    '产生就诊卡费用行
                    If mstrCardNO = "" Then
                        Set objDetail = ReadPatiCardObj(mobjBill.病人ID, mstrCardNO)
                        If mstrCardNO <> "" And Not objDetail Is Nothing Then
                            If Not ItemExist(objDetail.ID) Then
                                If mobjBill.Pages(1).Details.Count >= Bill.Rows - 1 Then
                                    Bill.Rows = Bill.Rows + 1
                                    Call bill_AfterAddRow(Bill.Rows - 1)
                                End If
                                Bill.TextMatrix(Bill.Rows - 1, 1) = "" '有必要加上
                                Call SetDetail(objDetail, Bill.Rows - 1)
                                Call CalcMoneys(Bill.Rows - 1)
                                Call ShowDetails(Bill.Rows - 1)
                                Call ShowMoney
                            End If
                        End If
                    End If
                End If
            Else
                '空姓名划价单输入不同姓名
                '刘兴洪:22343
                If txtPatient.Text <> mstrPrePati And mstrPrePati <> "" And gTy_Module_Para.byt缴款控制 <> 1 Then
                    mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
                    
                    txt应收.Text = Format(CalcBillToTal(True), gstrDec)
                    txt合计.Text = Format(CalcBillToTal, gstrDec)
                    txt应缴.Text = Format(mobjBill.Pages(1).应缴金额, "0.00")
                    
                    '调整费用列表
                    For i = mshMoney.Rows - 1 To 1 Step -1
                        If mshMoney.TextMatrix(i, 0) <> "" Then
                            intNum = Val(mshMoney.TextMatrix(i, 0))
                            Exit For
                        End If
                    Next
                    If intNum > 1 Then
                        mintBillNO = 0
                        
                        mshMoney.Redraw = False
                        For i = mshMoney.Rows - 1 To 1 Step -1
                            If Val(mshMoney.TextMatrix(i, 0)) <> intNum Then
                                mshMoney.RemoveItem i
                            Else
                                mshMoney.TextMatrix(i, 0) = 1
                            End If
                        Next
                        If mshMoney.Rows < 5 Then mshMoney.Rows = 5
                        mshMoney.Redraw = True
                    End If
                End If
            End If
            
            If Not mblnValid Then
                If cbo医疗付款.ListIndex = -1 And gbln医疗付款 Then
                    cbo医疗付款.SetFocus
                Else
                    If gbyt科室医生 = 0 Then
                        cbo开单人.SetFocus
                    Else
                        cbo开单科室.SetFocus
                    End If
                End If
            End If
            
            'LED初始化
            If Not mblnValid Then ShowWelcomeByLed
        End If
    End If
    mblnValid = False
End Sub

Private Function GetPatient(ByVal strInput As String, Optional blnCancel As Boolean, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '返回:
    '编制:刘兴洪
    '日期:2011-08-03 16:50:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strWhere As String
    Dim rsTmp As ADODB.Recordset, strPati As String
    Dim vRect As RECT
    
    blnCancel = False
    
    '病人输入的权限
    If gint病人来源 = 1 Then
        'strWhere = " And Nvl(A.当前科室ID,0)=0"
        strWhere = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    ElseIf gint病人来源 = 2 Then
        strWhere = " And Nvl(A.当前科室ID,0)<>0"
    End If
    
    '读取病人信息
    '76451,冉俊明,2014-8-19
    strSQL = "Select Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型,A.险类," & _
        IIf(gint病人来源 = 1, "NULL", "Decode(A.当前科室ID,NULL,NULL,A.主页ID)") & " as 主页ID,A.就诊卡号,A.卡验证码,A.门诊号," & _
        " A.住院号,A.姓名,A.性别,A.年龄,A.出生日期,A.费别,A.担保额," & _
        " A.医疗付款方式,A.当前病区ID,A.当前科室ID,A.当前床号" & _
        " From 病人信息 A Where A.停用时间 is NULL"
    
    '不允许发卡刷输入, 退出
    If blnCard And gint病人来源 = 1 And Not gblnInputCard Then Set mrsInfo = New ADODB.Recordset: Exit Function
    
    If blnCard = True And IDKind.GetCurCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then '103563
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & strWhere & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        If gint病人来源 = 1 And Not gblnInputID And Not (mstrInNO <> "" And mbytInState = 0) Then
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        If gint病人来源 = 1 And Not gblnInputID Then Set mrsInfo = New ADODB.Recordset: Exit Function
        
        strSQL = strSQL & strWhere & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        If gint病人来源 = 1 And Not gblnInputID Then Set mrsInfo = New ADODB.Recordset: Exit Function
        
        strSQL = strSQL & strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "." Then '挂号单号(最后为执行部门ID以区分)
        If gint病人来源 = 1 And Not gblnInputNO Then Set mrsInfo = New ADODB.Recordset: Exit Function
        '按日或年顺序编号规则
        strInput = GetFullNO(Mid(strInput, 2), 12)
        txtPatient.Text = strInput
        
        '76451,冉俊明,2014-8-19
        strSQL = "" & _
        "Select Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型,A.险类," & _
                IIf(gint病人来源 = 1, "NULL", "Decode(A.当前科室ID,NULL,NULL,A.主页ID)") & " as 主页ID,A.就诊卡号,A.卡验证码,Nvl(B.标识号,A.门诊号) as 门诊号," & _
        "       A.住院号,B.姓名,B.性别,B.年龄,A.出生日期,B.费别,A.担保额,A.医疗付款方式,A.当前病区ID,A.当前科室ID,A.当前床号,B.执行人,B.执行部门ID" & _
        " From 病人信息 A,门诊费用记录 B" & _
        " Where B.病人ID=A.病人ID(+) And B.记录性质=4 And B.记录状态=1" & _
            zlGetRegEventsCons("加班标志", "B") & _
                strWhere & " And B.NO=[2]"
    Else
    
        Select Case IDKind.GetCurCard.名称
        Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If mrsInfo!姓名 = strInput Then GetPatient = True: Exit Function
                End If
                '通过姓名模糊查找病人(允许输入病人标识时)
                If Not mblnValid And gblnSeekName And gblnInputID Then
                    strWhere = " And A.姓名 Like '" & strInput & "%' " & strWhere
                    strPati = _
                    " Select /*+Rule */1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                    IIf(gint病人来源 = 2, "A.住院号,B.名称 as 科室,A.当前床号 as 床号,", "A.门诊号,") & _
                    " A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                    " From 病人信息 A,部门表 B" & _
                    " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And Rownum<101 " & strWhere & _
                    IIf(gintNameDays = 0, "", " And (A.就诊时间>Trunc(Sysdate-" & gintNameDays & ") Or A.登记时间>Trunc(Sysdate-" & gintNameDays & "))")
                    
                    '门诊病人收费时可以不对应病人档案
                    If gint病人来源 = 1 Then
                        strPati = strPati & " Union ALL " & _
                            "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                    End If
                    strPati = strPati & " Order by 排序ID,姓名"
                        
                    vRect = zlControl.GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSelect(Me, strPati, 0, "病人0" & gint病人来源, , , , , , True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, , True, 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp!ID = 0 Then '当作新病人
                            strSQL = ""
                        Else '以病人ID读取
                            strInput = rsTmp!病人ID
                            strSQL = strSQL & strWhere & " And A.病人ID=[2]"
                        End If
                    Else '取消选择
                        strSQL = ""
                    End If
                Else
                    strSQL = ""
                End If
        Case "医保号"
            strInput = UCase(strInput)
             strSQL = strSQL & strWhere & "  And A.医保号=[2]"
        Case "身份证号", "二代身份证", "身份证"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            strInput = "-" & lng病人ID
            strSQL = strSQL & strWhere & " And A.病人ID=[2]"
        Case "IC卡号", "IC卡"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            strInput = "-" & lng病人ID
            strSQL = strSQL & strWhere & " And A.病人ID=[2]"
        Case "门诊号"
            If gint病人来源 = 1 And Not gblnInputID Then
                Set mrsInfo = New ADODB.Recordset
                Exit Function
            End If
            If Not IsNumeric(strInput) Then strInput = "0"
            If gint病人来源 = 1 Then strWhere = ""
            strSQL = strSQL & strWhere & " And A.门诊号=[2]"
            '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
            strInput = zlCommFun.GetFullNO(strInput, 3)
        Case "住院号"
            If gint病人来源 = 1 And Not gblnInputID Then
                Set mrsInfo = New ADODB.Recordset
                Exit Function
            End If
            If Not IsNumeric(strInput) Then strInput = "0"
            If gint病人来源 = 1 Then strWhere = ""
            strSQL = strSQL & strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
        Case Else
            '其他类别的,获取相关的病人ID
            If IDKind.GetCurCard.接口序号 > 0 Then
                lng卡类别ID = IDKind.GetCurCard.接口序号
                If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng病人ID = 0 Then GoTo NotFoundPati:
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, _
                    strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            End If
            If lng病人ID <= 0 Then GoTo NotFoundPati:
            strSQL = strSQL & strWhere & " And A.病人ID=[1]"
            strInput = "-" & lng病人ID
            blnHavePassWord = True
        End Select
    End If
    On Error GoTo errH
    '75259:李南春,2014-7-10，病人姓名颜色处理
    If strSQL = "" Then GoTo NotFoundPati:
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then GoTo NotFoundPati:
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), &HC00000, vbRed))
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
    GetPatient = True
    Exit Function
NotFoundPati:
    txtPatient.ForeColor = &HC00000
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub txtPatient_LostFocus()
    If mbytInState = 0 And Trim(txtPatient.Text) <> "" Then
        mobjBill.姓名 = txtPatient.Text
        mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
        mobjBill.性别 = zlStr.NeedName(cboSex.Text)
    End If
    zlCommFun.OpenIme False
    mblnKeyPress = False
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If gint病人来源 = 1 And InStr(mstrPrivs, "允许非医保病人") = 0 And txtPatient.Text <> "" And Not txtPatient.Locked Then
        Call ClearPatientInfo:  Exit Sub
    End If
    
    If Not mblnKeyPress Then
        mblnValid = True: Call txtPatient_KeyPress(13): mblnValid = False
    End If
End Sub

Private Sub txt缴款_Change()
    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00": Exit Sub
    txt找补.Text = Format(Val(txt缴款.Text) - Val(txt应缴.Text), "0.00")
End Sub

Private Sub txt缴款_GotFocus()
    '修改时不管工本费
    If mbytInState = 0 And mobjBill.Pages(1).Details.Count <> 0 And gTy_Module_Para.bln工本费 Then 'And mstrInNO = "" Then
        Call SetFactMoney
    End If
    
    '只以缴款作为收费结束条件时,必须输入缴款或0
    '刘兴洪:22343
    If mbytInState = 0 And (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 2) Then
        If Val(txt缴款.Text) = 0 And Me.ActiveControl Is txt缴款 Then
            txt缴款.Text = ""
        End If
    End If
    
    Call zlControl.TxtSelAll(txt缴款)
    
    'LED显示
    If mbytInState = 0 And gblnLED Then
        '自动报价或手工报价时由热键激活
        If (Not gbln手工报价 And ActiveControl Is txt缴款) Or (gbln手工报价 And mblnHotKey) Then
            mblnHotKey = False
            zl9LedVoice.Speak "#21 " & txt应缴.Text
            mbln报合计 = True
        End If
    End If
End Sub

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '只以缴款作为收费结束条件时,必须输入缴款或0
        '刘兴洪:22343
        If mbytInState = 0 And (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 2) Then
            If txt缴款.Text = "" Then Exit Sub
        End If
        
        If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
        If Val(txt缴款.Text) <> 0 Then
            If CSng(txt找补.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                If gblnLED And CCur(txt合计.Text) <> 0 And mbytInState = 0 Then 'LED显示
                    mblnHotKey = False
                    If Val(txt预交冲款.Text) = 0 Then
                        zl9LedVoice.DispCharge txt应缴.Text, txt缴款.Text, txt找补.Text
                    Else '部分支付现金时的处理
                        Call zl9LedVoice.DisplayBank( _
                            "合计:" & txt合计.Text & "元,应付" & txt应缴.Text & "元", _
                            "收您:" & txt缴款.Text & "元" & IIf(Val(txt找补.Text) = 0, "", ",找您:" & txt找补.Text & "元"))
                    End If
                    
                    zl9LedVoice.Speak "#22 " & txt缴款.Text
                    zl9LedVoice.Speak "#23 " & txt找补.Text
                    zl9LedVoice.Speak "#3"
                End If
            Else
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt缴款): txt缴款.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '病人累加缴款
        End If
    End If
    If KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
End Sub

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'功能：计算或重新计算指定行或所有行的金额
'参数：lngRow=指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long
    If mobjBill.Pages(1).Details.Count = 0 Then Exit Sub
    If lngRow = 0 Then
        For i = 1 To mobjBill.Pages(1).Details.Count
            CalcMoney i
        Next
    Else
        CalcMoney lngRow
    End If
End Sub

Private Sub CalcMoney(lngRow As Long)
'功能：计算或重新计算指定行的金额
'参数：lngRow=指定行
'说明：1.ExpenseBill集合的索引对应单据的行号
'      2.变价只能对应一个收入项目:mobjBill.Details(lngRow).InComes(1)
'      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
'      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    Dim i As Long, strInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim dblMoney As Double '用户输入的变价金额
    Dim str费别 As String
    Dim dbl加班加价率 As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    If mstr普通价格等级 <> "" Then
        strWherePriceGrade = _
            "       And (b.价格等级 = [2]" & vbNewLine & _
            "            Or (b.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From 收费价目" & vbNewLine & _
            "                               Where b.收费细目Id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                     And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    gstrSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,b.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID=A.ID And C.ID = B.收入项目ID " & _
        "       And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        "       And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Pages(1).Details(lngRow).收费细目ID, mstr普通价格等级)
    
    If rsTmp.RecordCount > 0 Then
        With mobjBill.Pages(1).Details(lngRow)
            If .Detail.变价 Then
                If .InComes.Count = 0 Then '第一次计算金额取缺省值
                    dblMoney = Val(Nvl(rsTmp!缺省价格))
                Else                        '获取操作员以前输入的变价金额
                    dblMoney = .InComes(1).标准单价
                    '如果用户输入的变价不满足变价范围，则取缺省值
                    If CheckScope(Val(Nvl(rsTmp!原价)), Val(Nvl(rsTmp!现价)), dblMoney) <> "" Then
                        dblMoney = Val(Nvl(rsTmp!缺省价格))
                    End If
                End If
            End If
        End With
        
        '再清除原有记录
        Set mobjBill.Pages(1).Details(lngRow).InComes = New BillInComes
        
        '填写现有费用记录
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .收入项目ID = rsTmp!收入项目ID
                .收入项目 = rsTmp!名称
                .收据费目 = Nvl(rsTmp!收据费目)
                .原价 = Val(Nvl(rsTmp!原价))
                .现价 = Val(Nvl(rsTmp!现价))
                If mobjBill.Pages(1).Details(lngRow).Detail.变价 Then
                    .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .标准单价 = Format(Val(Nvl(rsTmp!现价)), gstrFeePrecisionFmt)
                End If
                
                '应收金额=单价 * 付数 * 数次
                .应收金额 = .标准单价 * IIf(mobjBill.Pages(1).Details(lngRow).付数 = 0, 1, mobjBill.Pages(1).Details(lngRow).付数) * mobjBill.Pages(1).Details(lngRow).数次
                
                '附加手术费率用计算(所有收入项目)
                If mobjBill.Pages(1).Details(lngRow).附加标志 = 1 And mobjBill.Pages(1).Details(lngRow).收费类别 = "F" Then
                    .应收金额 = .应收金额 * IIf(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
                End If
                
                '加班费用率计算
                dbl加班加价率 = 0
                If mobjBill.加班标志 = 1 And mobjBill.Pages(1).Details(lngRow).Detail.加班加价 Then
                    dbl加班加价率 = IIf(IsNull(rsTmp!加班加价率), 0, rsTmp!加班加价率 / 100)             '传入根据费别计算实收金额函数
                    .应收金额 = .应收金额 + .应收金额 * dbl加班加价率
                End If
                
                .应收金额 = CCur(Format(.应收金额, gstrDec))
                
                If mobjBill.Pages(1).Details(lngRow).Detail.屏蔽费别 Then
                    .实收金额 = .应收金额
                    mobjBill.Pages(1).Details(lngRow).费别 = mobjBill.费别
                Else
                    If .应收金额 = 0 Then
                        .实收金额 = 0
                        mobjBill.Pages(1).Details(lngRow).费别 = mobjBill.费别
                    Else
                        str费别 = IIf(glngSys Like "8??", mobjBill.费别, zlStr.TrimEx(mobjBill.费别 & "," & lbl动态费别.Tag, ","))
                        
                        .实收金额 = CCur(Format(ActualMoney(str费别, .收入项目ID, .应收金额, _
                            mobjBill.Pages(1).Details(lngRow).收费细目ID, 0, 0, dbl加班加价率), gstrDec))
                        mobjBill.Pages(1).Details(lngRow).费别 = str费别
                    End If
                End If
                
                '实收金额存入Key中,以处理分币问题(即Key中存放原始实收金额,不变)
                mobjBill.Pages(1).Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额
            End With
            rsTmp.MoveNext
        Next
    Else
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Pages(1).Details(lngRow).InComes = New BillInComes
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
    Dim i As Long
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Pages(1).Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
End Sub

Private Sub ShowDetail(lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, j As Long, curMoney As Currency
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '清除单据行
    For i = 0 To Bill.COLS - 1
        '输入时收费类别不清除
        If Not (i = 0 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    '刷新单据行
    For i = 0 To Bill.COLS - 1
        Select Case Bill.TextMatrix(0, i)
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(1).Details(lngRow).Detail.名称
            Case "应收金额" '实际上是单价
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                curMoney = 0
                If mobjBill.Pages(1).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(1).Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Pages(1).Details(lngRow).InComes(j).应收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "实收金额"
                '实收金额是该收费细目所有收入项目的合计
                curMoney = 0
                If mobjBill.Pages(1).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(1).Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Pages(1).Details(lngRow).InComes(j).实收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "执行科室"
                If mbytInState = 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Pages(1).Details(lngRow).执行部门ID
                    If mrsUnit.RecordCount <> 0 Then
                        Bill.TextMatrix(lngRow, i) = mrsUnit!编码 & "-" & mrsUnit!名称
                    Else
                        Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Pages(1).Details(lngRow).执行部门ID, mrsUnit)
                    End If
                Else
                    '浏览单据只(能)显示名称
                    Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Pages(1).Details(lngRow).执行部门ID, mrsUnit)
                End If
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(1).Details(lngRow).Detail.类型
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
'功能：刷新显示收入项目费用区
    Dim cur实收合计 As Currency, cur应收合计 As Currency, cur冲款合计 As Currency
    Dim blnExist As Boolean, i As Integer, j As Integer, k As Integer
    
    '产生汇总费目
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Pages(1).Details.Count
        For j = 1 To mobjBill.Pages(1).Details(i).InComes.Count
            '查找是否已经加入此类收据费目,如是则合计,否则新入
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If gint分类合计 = 0 Then
                    If mcolMoneys(k).收据费目 = mobjBill.Pages(1).Details(i).InComes(j).收据费目 Then
                        blnExist = True: Exit For
                    End If
                Else
                    If mcolMoneys(k).收入项目 = mobjBill.Pages(1).Details(i).InComes(j).收入项目 Then
                        blnExist = True: Exit For
                    End If
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + mobjBill.Pages(1).Details(i).InComes(j).实收金额
                mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + mobjBill.Pages(1).Details(i).InComes(j).应收金额
            Else
                With mobjBill.Pages(1).Details(i).InComes(j)
                    mcolMoneys.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额
                End With
            End If
        Next
    Next
    
    '刷新显示(收费要叠加)
    mshMoney.Redraw = False
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1 + mintMoneyRow
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    
    Call SetMoneyList
    
    For i = mintMoneyRow + 1 To mcolMoneys.Count + mintMoneyRow
        mshMoney.TextMatrix(i, 0) = mintBillNO + 1
        If gint分类合计 = 0 Then
            mshMoney.TextMatrix(i, 1) = mcolMoneys(i - mintMoneyRow).收据费目
        Else
            mshMoney.TextMatrix(i, 1) = mcolMoneys(i - mintMoneyRow).收入项目
        End If
        mshMoney.TextMatrix(i, 2) = Format(mcolMoneys(i - mintMoneyRow).实收金额, gstrDec)
        cur实收合计 = cur实收合计 + mcolMoneys(i - mintMoneyRow).实收金额
        cur应收合计 = cur应收合计 + mcolMoneys(i - mintMoneyRow).应收金额
    Next
    For i = 1 To mshMoney.Rows - 1
        If Val(mshMoney.TextMatrix(i, 0)) = mintBillNO + 1 Then
            mshMoney.TopRow = i
        End If
    Next
    mshMoney.Redraw = True
    
    '当前单据的相关汇总金额计算
    '----------------------------------------
    With mobjBill.Pages(1)
        cur冲款合计 = Format(Val(txt预交冲款.Text), "0.00")
        
        .应收金额 = cur应收合计
        .实收金额 = cur实收合计
        
        '计算当前单据应分解冲款的金额,为了计算应缴(多单据时先冲预交)
        If cur冲款合计 <> 0 Then
            If cur冲款合计 <= Format(.实收金额, "0.00") Then
                .冲预交额 = cur冲款合计
            Else
                .冲预交额 = Format(.实收金额, "0.00")
            End If
            cur冲款合计 = cur冲款合计 - .冲预交额
        End If
        
        '计算当前单据应缴金额，分币处理，误差等
        .应缴金额 = Format(.实收金额 - .冲预交额, "0.00")
        
        '现金方式时才处理分币
        If cbo结算方式.ListIndex <> -1 Then
            If cbo结算方式.ItemData(cbo结算方式.ListIndex) = 1 Then
                .应缴金额 = CentMoney(.实收金额 - .冲预交额)
            End If
        End If
    
        .误差金额 = Format((.实收金额 - .冲预交额) - .应缴金额, gstrDec)
    End With
    
    txt应收.Text = Format(mcurBill应收 + cur应收合计, gstrDec)
    txt合计.Text = Format(mcurBill实收 + cur实收合计, gstrDec)
    txt应缴.Text = Format(mcurBill应缴 + mobjBill.Pages(1).应缴金额, "0.00")
    
    '误差显示
    If mobjBill.Pages(1).误差金额 <> 0 Then
        pic误差.Visible = True
        lbl误差额.Caption = Format(mobjBill.Pages(1).误差金额, "0.00")
    Else
        pic误差.Visible = False
    End If
End Sub

Private Sub ShowPrice()
'功能：在收费收取划价单费用时，计算并显示划价单据各类汇总信息
    Dim cur冲款合计 As Currency
    
    With mobjBill.Pages(1)
        '计算当前单据应分解冲款的金额,为了计算应缴(多单据时先冲预交)
        cur冲款合计 = Val(txt预交冲款.Text)
        If cur冲款合计 <> 0 Then
            If cur冲款合计 <= Format(.实收金额, "0.00") Then
                .冲预交额 = cur冲款合计
            Else
                .冲预交额 = Format(.实收金额, "0.00")
            End If
            cur冲款合计 = cur冲款合计 - .冲预交额
        End If
        
        '计算当前单据应缴金额，分币处理，误差等
        .应缴金额 = Format(.实收金额 - .冲预交额, "0.00")
        
        '现金方式时才处理分币
        If cbo结算方式.ListIndex <> -1 Then
            If cbo结算方式.ItemData(cbo结算方式.ListIndex) = 1 Then
                .应缴金额 = CentMoney(.实收金额 - .冲预交额)
            End If
        End If
    
        .误差金额 = Format((.实收金额 - .冲预交额) - .应缴金额, gstrDec)
        
        '显示合计
        txt应收.Text = Format(.应收金额 + mcurBill应收, gstrDec)
        txt合计.Text = Format(.实收金额 + mcurBill实收, gstrDec)
        txt应缴.Text = Format(.应缴金额 + mcurBill应缴, "0.00")
        
        '误差显示
        If .误差金额 <> 0 Then
            pic误差.Visible = True
            lbl误差额.Caption = Format(.误差金额, "0.00")
        Else
            pic误差.Visible = False
        End If
    End With
End Sub

Private Function GetInputDetail(ByVal lng项目id As Long) As Detail
'功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称,A.规格,A.计算单位," & _
        " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要" & _
        " From 收费项目目录 A,收费项目类别 B" & _
        " Where A.类别=B.编码 And A.ID=[1]"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id)
    With objDetail
        .ID = rsTmp!ID
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .编码 = rsTmp!编码
        .名称 = rsTmp!名称
        .规格 = Nvl(rsTmp!规格)
        .计算单位 = Nvl(rsTmp!计算单位)
        .变价 = Nvl(rsTmp!是否变价, 0) = 1 '对药品表明是否时价
        .类型 = Nvl(rsTmp!费用类型)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, Optional bytParent As Byte = 0, Optional ByVal lngDoUnit As Long)
'功能：根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
'说明：
'      1.用于新输入或更改收费细目行！！！
'      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在
    Dim tmpIncomes As New BillInComes
    Dim dblTime As Double, i As Long
    
        
    '执行科室
    If bytParent <> 0 Then
        '从属项目的执行科室,如果类别与主项相同,或设为无明确执行科室,则取主项执行科室,否则取本身的
        If lngDoUnit <> 0 Then
            lngDoUnit = mobjBill.Pages(1).Details(bytParent).执行部门ID
        Else
            If cbo开单科室.ListIndex <> -1 Then
                lngDoUnit = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            End If
            lngDoUnit = Get收费执行科室ID("Z", Detail.ID, Detail.执行科室, lngDoUnit, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
        End If
    Else
        lngDoUnit = mobjBill.科室ID
        If lngDoUnit = 0 And cbo开单科室.ListIndex <> -1 Then
            lngDoUnit = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        lngDoUnit = Get收费执行科室ID("Z", Detail.ID, Detail.执行科室, lngDoUnit, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
    End If
    
    If mobjBill.Pages(1).Details.Count < lngRow Then
        '如果该行对应的程序对象尚未初始,则加入
        With Detail
            '序号=行号,父号=0
            '付数=1
            '次数=1,从属项目的次数由主项计算确定
            '执行部门ID:根据细目执行科室标志取
            '附加标志:以第一行为假,其它为真优先权
            '收入集=空
            If bytParent <> 0 Then
                '初始数次
                If Detail.固有从属 = 0 Then '非固有从属
                    dblTime = mobjBill.Pages(1).Details(bytParent).数次
                ElseIf Detail.固有从属 = 1 Then '固定的固有从属
                    dblTime = Detail.从项数次
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Pages(1).Details(bytParent).数次
                End If
            Else
                dblTime = 1
            End If
            
            mobjBill.Pages(1).Details.Add mobjBill.费别, Detail, .ID, CByte(lngRow), CInt(bytParent), .类别, .计算单位, "", 1, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else
        '如果该行已经存在,则修改
        With mobjBill.Pages(1).Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .费别 = mobjBill.费别
            .付数 = 1
            .附加标志 = 0
            .计算单位 = Detail.计算单位
            .收费类别 = Detail.类别
            .收费细目ID = Detail.ID
            .数次 = 1
            .序号 = lngRow
            .从属父号 = 0
            .执行部门ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'功能：判断该行是否应该取从属项目
'说明：仅该行收费项目有从属项目及尚未取才取。
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle

    strSQL = "Select count(从项ID) as NUM From 收费从属项目 Where 主项ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(1).Details(lngRow).收费细目ID)
    If rsTmp.RecordCount <> 0 Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Pages(1).Details.Count
                If mobjBill.Pages(1).Details(i).从属父号 = lngRow Then
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(lng项目id As Long) As Details
'功能：返回一个收费细目的从属项目集
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim objDetail As New Detail
    
    Set GetSubDetails = New Details
    
    strSQL = _
        "Select A.ID,A.类别,B.名称 as 类别名称," & _
        " A.费用类型,A.编码,A.名称,A.规格,A.计算单位,A.屏蔽费别,A.是否变价," & _
        " A.加班加价,A.执行科室,C.固有从属,C.从项数次 " & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.类别='Z' And C.主项ID=[1]" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .编码 = rsTmp!编码
            .变价 = Nvl(rsTmp!是否变价, 0) = 1
            .规格 = Nvl(rsTmp!规格)
            .计算单位 = Nvl(rsTmp!计算单位)
            .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
            .类别 = rsTmp!类别
            .类别名称 = rsTmp!类别名称
            .名称 = rsTmp!名称
            .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
            .执行科室 = Nvl(rsTmp!执行科室, 0) '缺省为无明确科室(用户选)
            .固有从属 = Nvl(rsTmp!固有从属, 0) '缺省为非固定,用户可以随意更改数次
            .从项数次 = Nvl(rsTmp!从项数次, 1)
            .类型 = Nvl(rsTmp!费用类型)
            
            GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .规格, .计算单位, .说明, .屏蔽费别, _
                1, .计算单位, .分批, .变价, .加班加价, .执行科室, .类型, .补充摘要, .固有从属, .从项数次
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
    For i = lngRow + 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).从属父号 <> 0 And _
            mobjBill.Pages(1).Details(i).从属父号 > lngRow Then
            mobjBill.Pages(1).Details(i).从属父号 = mobjBill.Pages(1).Details(i).从属父号 - 1
        End If
        mobjBill.Pages(1).Details(i).序号 = mobjBill.Pages(1).Details(i).序号 - 1 '序号与行号对应
    Next
    mobjBill.Pages(1).Details.Remove lngRow
    If lngRow = 1 And mobjBill.Pages(1).Details.Count = 0 And Bill.Rows = 2 Then
        For i = 0 To Bill.COLS - 1
            Bill.TextMatrix(lngRow, i) = ""
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Function NewBill(Optional blnFact As Boolean = True, Optional bln费别 As Boolean = True) As Boolean
'功能：初始化一张新的单据(程序对象)
'参数：blnFact=是否取票号
'      bln费别=是否重新初始化费别
    Dim i As Long
    
    mblnKeyPress = False
    mblnHotKey = False
    mbln报合计 = False
    
    cbo费别.Locked = False
    cbo医疗付款.Locked = False
    
    Set mobjBill = New ExpenseBill
    
    mstrCardNO = ""
    
    '预交信息初始
    txt预交冲款.Text = "0.00"
    lblDeposit.ForeColor = &H808080
    txt预交冲款.Enabled = False
    txt预交冲款.ForeColor = &H808080
    sta.Panels(3).Tag = ""
    sta.Panels(3).Text = ""
    sta.Panels(3).Visible = False
    '隐藏误差
    pic误差.Visible = False
    
    txtPatient.Locked = False
    cboSex.Locked = False
    txt年龄.Locked = False
    cbo年龄单位.Locked = False
    
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Set mrsInfo = New ADODB.Recordset
    
    If mbytInState = 0 Then
        cboNO.Text = ""
        
        '实际号码
        Call ReInitPatiInvoice(blnFact)
        
        '仅为执行状态才初始
        chk加班.Value = IIf(OverTime(zlDatabase.Currentdate), Checked, Unchecked)
        
        '结算方式
        i = cbo.FindIndex(cbo结算方式, gstr结算方式, True)
        If i = -1 And cbo结算方式.ListCount > 0 Then i = 0
        Call zlControl.CboSetIndex(cbo结算方式.hWnd, i)
        
        '其它
        With mobjBill
            .NO = cboNO.Text
            .操作员编号 = UserInfo.编号
            .操作员姓名 = UserInfo.姓名
            .发生时间 = CDate(txtDate.Text)
            .费别 = IIf(cbo费别.ListIndex = -1, "", Mid(cbo费别.Text, InStr(cbo费别.Text, "-") + 1))
            .加班标志 = IIf(chk加班.Value = Checked, 1, 0)
            If cbo开单科室.ListIndex = -1 Then
                .Pages(1).开单部门ID = 0
            Else
                .Pages(1).开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            End If
            .门诊标志 = gint病人来源
            .Pages(1).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
        End With
        
        '费别处理:收费
        cbo费别.Locked = False
        cbo费别.Visible = True
        lbl动态费别.AutoSize = True
        lbl动态费别.BorderStyle = 0
        lbl动态费别.Left = cbo费别.Left + cbo费别.Width + 60
        
        If bln费别 Then
            Call LoadAndSeek费别
        End If
    End If
    NewBill = True
End Function

Private Sub ClearMoney()
'功能：清除费用显示区
    Dim i As Integer, j As Integer
    mshMoney.Redraw = False
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.COLS - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mintMoneyRow = 0
    mshMoney.Rows = 5
    mshMoney.Redraw = True
End Sub

Private Function SaveBill() As Boolean
'功能:保存当前输入的单据
'入口:mobjBill=单据对象
    Dim i As Integer, j As Integer, str医疗付款 As String
    Dim int序号 As Integer, int行号 As Integer, strNo As String, strTmp As String
    Dim intParent As Integer, intParentNO As Integer
    Dim arrSQL As Variant, strDelBill As String, lng结帐ID As Long
    
    If cbo医疗付款.ListIndex <> -1 Then
        str医疗付款 = Mid(cbo医疗付款.Text, 1, InStr(1, cbo医疗付款, "-") - 1)
    End If
    Err = 0: On Error GoTo Errhand:
    mobjBill.NO = zlDatabase.GetNextNo(13)
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    
    For Each mobjBillDetail In mobjBill.Pages(1).Details
        intParent = 0: intParentNO = int序号
        For Each mobjBillIncome In mobjBillDetail.InComes
            int序号 = int序号 + 1 '当前记录序号
            '单据主体
            '76451,冉俊明,2014-8-19
            With mobjBill
                gstrSQL = "zl_门诊收费记录_INSERT('" & .NO & "'," & int序号 & "," & ZVal(.病人ID) & "," & IIf(.主页ID = 0, 1, ZVal(.主页ID)) & "," & _
                    ZVal(.标识号) & ",'" & IIf(gint病人来源 = 2, .床号, str医疗付款) & "','" & .姓名 & "'," & _
                    "'" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                    .加班标志 & "," & ZVal(.科室ID, , .Pages(1).开单部门ID) & "," & ZVal(.Pages(1).开单部门ID) & ",'" & .Pages(1).开单人 & "',"
            End With
            
            '收费细目部份
            With mobjBillDetail
                '处理从属父号
                If .序号 <> int行号 Then
                    int行号 = .序号
                    '重新处理从属父号
                    If mobjBill.Pages(1).Details(.序号).从属父号 = 0 Then
                        For i = .序号 + 1 To mobjBill.Pages(1).Details.Count
                            If mobjBill.Pages(1).Details(i).从属父号 = .序号 Then
                                mobjBill.Pages(1).Details(i).从属父号 = int序号 '当父项目有多个收入项目(多个序号)时,取第一个序号
                            End If
                        Next
                    End If
                End If
                
                gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                gstrSQL = gstrSQL & "NULL,NULL,'" & .收费类别 & "',"
                gstrSQL = gstrSQL & IIf(.付数 = 0, 1, .付数) & "," & .数次 & "," & _
                    IIf(.工本费, 8, .附加标志) & "," & .执行部门ID & ","
            End With
            
            '收入项目部份
            With mobjBillIncome
                intParent = intParent + 1
                gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .收入项目ID & "," & _
                    "'" & .收据费目 & "'," & .标准单价 & "," & .应收金额 & "," & .实收金额 & ","
                gstrSQL = gstrSQL & "NULL,"
            End With
                                            
            '其它部分
            '缴款结算
            gstrSQL = gstrSQL & _
                "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "'" & mstrInNO & "'," & lng结帐ID & ",'" & zlStr.NeedName(cbo结算方式.Text) & "|" & mobjBill.Pages(1).应缴金额 & "| ',"
            '预交结算
            If Val(txt预交冲款.Text) <> 0 Then
                gstrSQL = gstrSQL & mobjBill.Pages(1).冲预交额 & ","
            Else
                gstrSQL = gstrSQL & "NULL,"
            End If
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        Next
    Next
    
    '修改前退除原单据
    If mstrInNO <> "" Then
        strDelBill = "zl_门诊简单收费_DELETE('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    End If
    If UBound(arrSQL) >= 0 Then
        '执行SQL语句
        On Error GoTo errH
        gcnOracle.BeginTrans
            '删除就诊卡划价单
            If mstrCardNO <> "" Then
                gstrSQL = "zl_门诊划价记录_Delete('" & mstrCardNO & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
    
            '修改前作废被修改单据
            If strDelBill <> "" Then
                Call zlDatabase.ExecuteProcedure(strDelBill, Me.Caption)
            End If
            '产生新费用
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            '处理单据误差
            If mobjBill.Pages(1).误差金额 <> 0 Then
                '  Zl_简单收费误差_Insert(
                '  No_In         门诊费用记录.No%Type,
                '  病人id_In     门诊费用记录.病人id%Type,
                '  结帐id_In     门诊费用记录.结帐id%Type,
                '  误差金额_In   门诊费用记录.实收金额%Type,
                '  登记时间_In   门诊费用记录.登记时间%Type,
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type
                gstrSQL = "Zl_简单收费误差_Insert('" & mobjBill.NO & "'," & Val(mobjBill.病人ID) & "," & lng结帐ID & "," & _
                    mobjBill.Pages(1).误差金额 & ",To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        gcnOracle.CommitTrans
        
        '加入单据历史记录(所有类型单据)
        cboNO.AddItem mobjBill.NO, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i '只显示10个
        Next
    End If
    SaveBill = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function ReadBill(ByVal strNo As String, Optional ByVal bln划价 As Boolean, _
    Optional blnDelete As Boolean, Optional blnNull As Boolean) As Boolean
'功能：根据单据号读取一张单据并将其填入表格
'参数：blnDelete=是否读取要退费的单据,按退费方式读
'说明：为统一，退费时都不显示误差费用(虽然不可部份退费)
    Dim rsTmp As ADODB.Recordset, rs结算 As ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset, strSQL As String
    Dim i As Long, curBill实收 As Currency, curBill应收 As Currency
    Dim blnSame As Boolean, str费别 As String, intSign As Integer
    Dim str费用费别 As String, blnHaveNoOne As Boolean
    
    On Error GoTo errH
    
    strNo = GetFullNO(strNo, 13)
    Call ClearRows: Call Bill.ClearBill
    
    '读取单据主体
    strSQL = _
    " Select A.结帐ID,A.实际票号 as 票据号,A.病人ID,0 as 主页ID,A.标识号,A.姓名,A.性别,A.年龄,A.费别,A.付款方式 ,B.病人类型,B.险类," & _
    "        0 as 病人病区ID,A.病人科室ID,A.开单部门ID,Nvl(A.加班标志,0) as 加班标志,A.开单人,A.划价人,A.发生时间,B.医疗付款方式,A.门诊标志" & _
    " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,病人信息 B,人员表 C" & _
    " Where Rownum=1 And A.病人ID=B.病人ID(+)" & _
    "       And A.记录状态" & IIf(mstrDelete <> "", "=2", IIf(bln划价, "=0", " IN(1,3)")) & _
    "       And A.记录性质=1 And A.NO=[1] And Nvl(A.操作员姓名,A.划价人)=C.姓名" & _
    "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
            IIf(mstrDelete <> "", " And A.登记时间=[2]", "") & _
            IIf(bln划价, " And A.操作员姓名 is Null And 划价人 is Not NULL", "")
    If mstrDelete <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrDelete))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    If rsTmp.EOF Then
        MsgBox "没有发现该单据,该单据可能已经作废！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '单据号
    cboNO.Text = strNo
    If Trim(Nvl(rsTmp!票据号)) <> "" Then txtInvoice.Text = Nvl(rsTmp!票据号)
    
    cboNO.Tag = IIf(IsNull(rsTmp!结帐ID), "", rsTmp!结帐ID) '用于医保门诊退费
    
    '病人ID
    txtPatient.Tag = Nvl(rsTmp!病人ID)
    
    '病人相关信息提取:可能用于划价单收费
    mobjBill.姓名 = Nvl(rsTmp!姓名)
    mobjBill.性别 = Nvl(rsTmp!性别)
    mobjBill.年龄 = Nvl(rsTmp!年龄)
    mobjBill.病人ID = Nvl(rsTmp!病人ID, 0)
    mobjBill.主页ID = Nvl(rsTmp!主页ID, 0)
    mobjBill.标识号 = Nvl(rsTmp!标识号, 0)
    mobjBill.床号 = IIf(gint病人来源 = 2, "" & rsTmp!付款方式, "") '床号暂存付款方式
    mobjBill.病区ID = Nvl(rsTmp!病人病区ID, 0)
    mobjBill.科室ID = Nvl(rsTmp!病人科室ID, 0)
    mobjBill.Pages(1).开单人 = Nvl(rsTmp!开单人)
    mobjBill.Pages(1).开单部门ID = Nvl(rsTmp!开单部门ID, 0)
    
    Call ReInitPatiInvoice
    '姓名
    If (IsNull(rsTmp!姓名) Or Nvl(rsTmp!姓名) = mstrPrePati) And chkCancel.Value = 0 Then
        If IsNull(rsTmp!姓名) Then
            blnNull = True
            txtPatient.Text = mstrPrePati '缺省为上一个病人姓名
        Else
            txtPatient.Text = rsTmp!姓名
            '75259:李南春,2014-7-10，病人姓名显示颜色处理
            Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), &HC00000, vbRed))
        End If
        blnSame = True
    Else
        '不同的病人
        txtPatient.Text = Nvl(rsTmp!姓名)
        '75259:李南春,2014-7-10，病人姓名显示颜色处理
        Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), &HC00000, vbRed))
        '如果仅以缴款作为结束,即使不同的病人也保留收费
        '除非刚好缴款结束(mstrPrePati = "")
        '刘兴洪:22343
        If gTy_Module_Para.byt缴款控制 <> 1 Or mstrPrePati = "" Then
            mcurBill实收 = 0: mcurBill应收 = 0: mcurBill应缴 = 0
            mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
            txt合计.Text = gstrDec: txt应收.Text = gstrDec
            Call ClearMoney
        End If
    End If
    
    '性别
    cboSex.ListIndex = cbo.FindIndex(cboSex, Nvl(rsTmp!性别), True)
    If cboSex.ListIndex = -1 Then
        If Not IsNull(rsTmp!性别) Then
            cboSex.AddItem rsTmp!性别, 0
            cboSex.ListIndex = 0
        ElseIf cboSex.ListCount > 0 Then
            cboSex.ListIndex = 0
        End If
    End If
    
    '年龄
    Call LoadOldData("" & rsTmp!年龄, txt年龄, cbo年龄单位)
    '费别
    cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(rsTmp!费别), True)
    If cbo费别.ListIndex = -1 And Not IsNull(rsTmp!费别) Then
        cbo费别.AddItem rsTmp!费别, 0
        cbo费别.ListIndex = 0
    End If
    
    '医疗付款方式
    If Nvl(rsTmp!门诊标志, 0) = 2 Or Not IsNull(rsTmp!医疗付款方式) Then
        cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, rsTmp!医疗付款方式, True)
        If cbo医疗付款.ListIndex = -1 Then
            cbo医疗付款.AddItem "0-" & rsTmp!医疗付款方式, 0
            cbo医疗付款.ListIndex = 0
        End If
    Else
        cbo医疗付款.ListIndex = GetCboIndexByCode(cbo医疗付款, "" & rsTmp!付款方式)
        If cbo医疗付款.ListIndex = -1 And Not IsNull(rsTmp!付款方式) Then
            cbo医疗付款.AddItem rsTmp!付款方式 & "-" & GetMedPayModeName(rsTmp!付款方式), 0
            cbo医疗付款.ListIndex = 0
        ElseIf cbo医疗付款.ListIndex = -1 Then
            cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, mstr付款方式, True)
        End If
    End If
    
    Call Set开单人开单科室(mobjBill.Pages(1).开单人, mobjBill.Pages(1).开单部门ID)
    
    '结算方式:不会存在医保结算方式
    If Not bln划价 Then
        '读取单据原始内容时,显示各种金额
        '(部份)退费时,显示原始单据的结算金额
        intSign = IIf(mstrDelete <> "", -1, 1) '数量,金额正负符号
        strSQL = _
            " Select 1 As 方式, 结算方式, Sum(1 * 冲预交) As 金额" & _
            " From 病人预交记录 A, 结算方式 B" & _
            " Where a.结算方式 = b.名称 And b.性质 <> 9 And 记录性质 = 3 And 结帐id = [1]" & _
            " Group By 结算方式" & _
            " Having Nvl(Sum(1 * 冲预交), 0) <> 0"
        '预交款
        strSQL = strSQL & _
            " Union All" & _
            " Select 2 As 方式, Null, Sum(1 * 冲预交) As 金额" & _
            " From 病人预交记录" & _
            " Where 记录性质 In (1, 11) And 结帐id = [1] Having Nvl(Sum(1 * 冲预交), 0) <> 0"
        '误差费
        strSQL = strSQL & _
            " Union All" & _
            " Select 3 As 方式, 结算方式, Sum(1 * 冲预交) As 金额" & _
            " From 病人预交记录 A, 结算方式 B" & _
            " Where a.结算方式 = b.名称 And b.性质 = 9 And 记录性质 = 3 And 结帐id = [1]" & _
            " Group By 结算方式" & _
            " Having Nvl(Sum(1 * 冲预交), 0) <> 0"

            
        Set rs结算 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp!结帐ID))
        
        For i = 1 To rs结算.RecordCount
            If rs结算!方式 = 2 Then
                txt预交冲款.Text = Format(rs结算!金额, "0.00")
            ElseIf rs结算!方式 = 3 Then
                pic误差.Visible = True: lbl误差额.Caption = Format(rs结算!金额, "0.00")
            Else
                cbo结算方式.ListIndex = cbo.FindIndex(cbo结算方式, rs结算!结算方式, True)
                If cbo结算方式.ListIndex = -1 Then
                    cbo结算方式.AddItem rs结算!结算方式, 0
                    cbo结算方式.ListIndex = 0
                End If
            End If
            rs结算.MoveNext
        Next
    End If
    
    '加班状态
    chk加班.Value = IIf(IsNull(rsTmp!加班标志), 0, rsTmp!加班标志)
    
    '发生时间
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    
    '读取单据收费细目部份
    '---------------------------------------------------------------------------------------------
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        strSQL = "" & _
        " Select Nvl(价格父号,序号) " & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录" & _
        " Where 记录性质=1 And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1 And Nvl(附加标志,0)<>9"
        
        strSQL = _
        " Select A.记录状态,Nvl(A.价格父号,A.序号) as 序号," & _
        "        A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型,A.计算单位," & _
        "        Avg(Nvl(A.付数,1)*A.数次) as 数量,Sum(A.标准单价) as 单价," & _
        "        Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
        "        A.执行部门ID,D.名称 as 执行部门,A.附加标志" & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D " & _
        " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID " & _
        "       And A.记录性质=1 And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
        "       And Nvl(A.附加标志,0)<>9" & _
        " Group by A.记录状态,Nvl(A.价格父号,A.序号),A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
        "          B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志"
            
        '最后计算结果(剩余数量即为准退数量,不必计算)
        '排开已经全部退费的行(执行状态=0的一种可能)
        strSQL = _
        " Select A.序号,A.费别,A.编码,A.类别,A.收费细目ID,A.名称,A.规格," & _
        "        A.费用类型,A.计算单位,A.执行部门ID,A.执行部门,A.附加标志," & _
        "        Sum(A.数量) as 数量,A.单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额" & _
        " From (" & strSQL & ") A" & _
        " Group by A.序号,A.费别,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型," & _
        "          A.计算单位,A.单价,A.执行部门ID,A.执行部门,A.附加标志" & _
        " Having Sum(A.数量)<>0" & _
        " Order by A.序号"
    Else
        '读取单据原始内容
        intSign = IIf(mstrDelete <> "", -1, 1) '数量,金额正负符号
        strSQL = _
        " Select Nvl(A.价格父号,A.序号) as 序号," & _
        "        A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型,A.计算单位," & _
        "        Avg(" & intSign & "*Nvl(A.付数,1)*A.数次) as 数量," & _
        "        Sum(A.标准单价) as 单价,Sum(" & intSign & "*A.应收金额) as 应收金额, " & _
        "        Sum(" & intSign & "*A.实收金额) as 实收金额, " & _
        "        A.执行部门ID,D.名称 as 执行部门,A.附加标志" & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D " & _
        " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID " & _
        "       And A.记录性质=1 And A.NO=[1]" & _
        "       And A.记录状态" & IIf(mstrDelete <> "", "=2", IIf(bln划价, "=0", " IN(1,3)")) & _
                IIf(mstrDelete <> "", " And A.登记时间=[2]", "") & _
                IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
        " Group by Nvl(A.价格父号,A.序号),A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
        "           B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志" & _
        " Order by 序号"
    End If
    
    If mstrDelete <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrDelete))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    If rsTmp.EOF Then Exit Function
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    str费用费别 = "": blnHaveNoOne = False
    For i = 1 To rsTmp.RecordCount
        '费别
        If Not IsNull(rsTmp!费别) Then
            If InStr(str费别 & ",", "," & rsTmp!费别 & ",") = 0 Then
                str费别 = str费别 & "," & rsTmp!费别
            End If
        End If
    
        Bill.TextMatrix(i, 0) = rsTmp!名称
        Bill.TextMatrix(i, 1) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, 2) = Format(rsTmp!实收金额, gstrDec)
        Bill.TextMatrix(i, 3) = rsTmp!执行部门
        Bill.TextMatrix(i, 4) = IIf(IsNull(rsTmp!费用类型), "", rsTmp!费用类型)
        
        If Nvl(rsTmp!类别) <> "其他" Then
            If InStr(1, "," & str费用费别 & ",", "," & Nvl(rsTmp!类别) & ",") = 0 Then
                str费用费别 = str费用费别 & "," & Nvl(rsTmp!类别)
            End If
        End If
        If Val(Nvl(rsTmp!数量)) <> 1 Then blnHaveNoOne = True
        rsTmp.MoveNext
    Next
    
    If str费用费别 <> "" Then
        str费用费别 = Mid(str费用费别, 2)
        str费用费别 = Replace(str费用费别, ",", "，")
        MsgBox "单据 [" & strNo & "] 中存在如下类别的收费项目，不能进行简单收费！" & vbCrLf & vbCrLf & _
            "        " & str费用费别, vbInformation, gstrSysName
        Exit Function
    ElseIf blnHaveNoOne Then
        MsgBox "单据 [" & strNo & "] 中存在数量不为1的收费项目，不能进行简单收费！", vbInformation, gstrSysName
        Exit Function
    End If

    
    '费别
    lbl动态费别.Caption = Mid(str费别, 2)
    
    '针对列编辑性质设置颜色
    Bill.SetColColor 0, &HE7CFBA
    Bill.SetColColor 1, &HE7CFBA
    Bill.SetColColor 3, &HE7CFBA
    Bill.Redraw = True
    
    '读取单据收据费目汇总
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        '整张费用单据(明细到收入项目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        strSQL = "" & _
        " Select Nvl(价格父号,序号)  " & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录" & _
        " Where 记录性质=1 And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1 And Nvl(附加标志,0)<>9"
        
        strSQL = _
        " Select A.序号,A.名称," & _
        "       Sum(A.数量) as 剩余数量,Sum(A.应收金额) as 剩余应收," & _
        "       Sum(A.实收金额) as 剩余实收 " & _
        " From ( Select A.记录状态,A.序号," & IIf(gint分类合计 = 0, "A.收据费目", "B.名称") & " as 名称," & _
        "               Nvl(A.付数,1)*A.数次 as 数量,A.应收金额,A.实收金额" & _
        "        From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收入项目 B" & _
        "       Where A.记录性质=1 And A.NO=[1] And Nvl(A.附加标志,0)<>9" & _
        "            And A.收入项目ID=B.ID And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
        "        ) A" & _
        " Group by A.序号,A.名称" & _
        " Having Sum(数量)<>0"
                    
        '最后计算结果(准退数量即剩余数量,不必真正计算)
        strSQL = _
        " Select A.名称,Sum(A.剩余应收) as 应收金额," & _
        "       Sum(A.剩余实收) as 实收金额" & _
        " From (" & strSQL & ") A" & _
        " Group by A.名称"
    Else
        '读取单据原始内容
        intSign = IIf(mstrDelete <> "", -1, 1) '数量,金额正负符号
        strSQL = _
        " Select " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称") & " as 名称," & _
        "        Sum(" & intSign & "*A.应收金额) as 应收金额," & _
        "        Sum(" & intSign & "*A.实收金额) as 实收金额 " & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收入项目 B" & _
        " Where A.收入项目ID=B.ID And A.记录状态" & IIf(mstrDelete <> "", "=2", IIf(bln划价, "=0", " IN(1,3)")) & _
        "       AND A.记录性质=1 And A.NO=[1]" & _
                IIf(mstrDelete <> "", " And A.登记时间=[2]", "") & _
                IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
        " Group By " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称")
    End If
    
    If mstrDelete <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrDelete))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    If rsTmp.EOF Then Exit Function
    
    '刷新显示(收费要叠加)
    mshMoney.Rows = rsTmp.RecordCount + 1 + mintMoneyRow
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = mintMoneyRow + 1 To mshMoney.Rows - 1
        mshMoney.TextMatrix(i, 0) = ""
        mshMoney.TextMatrix(i, 1) = ""
        mshMoney.TextMatrix(i, 2) = ""
    Next
    For i = mintMoneyRow + 1 To rsTmp.RecordCount + mintMoneyRow
        mshMoney.TextMatrix(i, 0) = mintBillNO + 1
        mshMoney.TextMatrix(i, 1) = rsTmp!名称
        mshMoney.TextMatrix(i, 2) = Format(rsTmp!实收金额, gstrDec)
        curBill应收 = curBill应收 + rsTmp!应收金额
        curBill实收 = curBill实收 + rsTmp!实收金额
        rsTmp.MoveNext
    Next
    
    '汇总处理
    With mobjBill.Pages(1)
        .NO = strNo
        .应收金额 = curBill应收
        .实收金额 = curBill实收
        '收费时收取划价单时
        If bln划价 Then Call ShowPrice
    End With
    
    txt应收.Text = Format(mcurBill应收 + curBill应收, gstrDec)
    txt合计.Text = Format(mcurBill实收 + curBill实收, gstrDec)
    
    '收费显示退款合计
    If blnDelete Then
        lbl应缴.Caption = "应退金额"
        lbl应缴.ForeColor = vbRed
        txt应缴.ForeColor = vbRed
        txt应缴.Text = Format(GetDelMoney, "0.00")
    End If
    
    '刷新收费累计
    If chkCancel.Value = 0 And gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    
    On Error Resume Next
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.TextMatrix(i, 0) = mintBillNO + 1 Then
            mshMoney.TopRow = i
        End If
    Next
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get开单科室ID() As Long
    Dim lng开单人ID As Long
    Dim rs开单人 As ADODB.Recordset
    
    If gbyt科室医生 = 2 Then
        If cbo开单人.ListIndex <> -1 Then
            lng开单人ID = cbo开单人.ItemData(cbo开单人.ListIndex)
            Set rs开单人 = mrs开单人 '避免影响外部调用的记录集
            
            rs开单人.Filter = "缺省=1 And ID=" & lng开单人ID
            If rs开单人.RecordCount = 0 Then rs开单人.Filter = "ID=" & lng开单人ID
            If rs开单人.RecordCount > 0 Then Get开单科室ID = rs开单人!部门ID
        End If
    End If
    
    If Get开单科室ID = 0 Then
        If cbo开单科室.ListIndex <> -1 Then
            Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        Else
            Get开单科室ID = UserInfo.部门ID
        End If
    End If
End Function

Private Function GetBillCount() As Integer
'功能：计算当前收费需要打印多少张票据
    Dim strItems As String
    Dim i As Integer, j As Integer
    
    If gTy_Module_Para.bln一张票据 Then GetBillCount = 1: Exit Function
    
    '多张单据按费目汇总计算
    For i = 1 To mobjBill.Pages(1).Details.Count
        If Not mobjBill.Pages(1).Details(i).工本费 Then '排开工本费
            For j = 1 To mobjBill.Pages(1).Details(i).InComes.Count
                If mobjBill.Pages(1).Details(i).InComes(j).实收金额 <> 0 Then '金额不为零
                    If InStr(strItems & ",", "," & mobjBill.Pages(1).Details(i).InComes(j).收据费目 & ",") = 0 Then
                        strItems = strItems & "," & mobjBill.Pages(1).Details(i).InComes(j).收据费目
                    End If
                End If
            Next
        End If
    Next
    GetBillCount = IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt门诊收据行次)
End Function

Private Sub DelFactMoney()
'功能：删除单据中的工本费用(当不需要工本费时)
    Dim i As Long
    
    '先判断是否已经加入了工本费
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).工本费 Then
            Call DeleteDetail(i)
            Call ShowMoney
            If mobjBill.Pages(1).Details.Count = 0 Then ClearMoney
            Exit Sub
        End If
    Next
End Sub

Private Sub SetFactMoney()
'功能：收费时设置、显示、计算工本费
'说明：工本费自动加在当前显示的单据中
    Dim objDetail As Detail
    Dim colIncomes As New BillInComes
    Dim blnExist As Boolean, i As Integer
    Dim lngRow As Long, lngDoUnit As Long
    Dim int张数 As Integer
    
    int张数 = GetBillCount
    If int张数 = 0 Then Call DelFactMoney: Exit Sub '删除工本费
    
    '先判断是否已经加入了工本费
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).工本费 Then
            lngRow = i: blnExist = True: Exit For
        End If
    Next

    If Not blnExist Then
        Set objDetail = Get工本费
        If objDetail Is Nothing Then Exit Sub '找不到工本费,不设置
        
        If mobjBill.Pages(1).Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
        Else
            For i = 1 To Bill.COLS - 1
                Bill.TextMatrix(Bill.Rows - 1, i) = ""
            Next
        End If
        lngRow = mobjBill.Pages(1).Details.Count + 1
        
        lngDoUnit = mobjBill.科室ID '病人科室
        If lngDoUnit = 0 Then lngDoUnit = Get开单科室ID
        lngDoUnit = Get收费执行科室ID(objDetail.类别, objDetail.ID, objDetail.执行科室, lngDoUnit, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
        With objDetail
            mobjBill.Pages(1).Details.Add "", objDetail, .ID, CInt(lngRow), 0, .类别, .计算单位, .类别, 1, 1, 0, lngDoUnit, colIncomes
        End With
        mobjBill.Pages(1).Details(lngRow).工本费 = True
    End If
    
    '重新根据当前费用内容设置工本费数次(固定为1)
    mobjBill.Pages(1).Details(lngRow).数次 = int张数
    Call CalcMoney(lngRow)
    
    Call ShowDetails(lngRow)
    Call ShowMoney
End Sub

Private Sub ClearRows()
    Dim i As Integer
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Sub FillBillComboBox(lngRow As Long, lngCol As Long)
'功能：根据单据列设置下拉列表框内容
    Dim rsTmp As ADODB.Recordset
    Dim str人员性质 As String, strTmp As String
    Dim strSQL As String, i As Long
    Dim lng病区ID As Long, lng科室ID As Long
    Dim rsUnit As ADODB.Recordset
    On Error GoTo errHandle
    

    Bill.Clear
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "执行科室"
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Pages(1).Details.Count >= lngRow Then
                With mobjBill.Pages(1).Details(lngRow)
                    Bill.TextMatrix(lngRow, lngCol) = ""
                    
                    lng科室ID = mobjBill.科室ID
                    If lng科室ID = 0 Then lng科室ID = Get开单科室ID
                                        
                    If gint病人来源 = 2 Then
                        lng病区ID = mobjBill.病区ID
                        If lng病区ID = 0 Then lng病区ID = Get病区ID(lng科室ID)
                    End If
                    If lng病区ID = 0 Then lng病区ID = lng科室ID
                    
                    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
                    Select Case .Detail.执行科室
                        Case 0 '不明确
                            mrsUnit.Filter = 0
                        Case 1 '病人科室
                            mrsUnit.Filter = "ID=" & lng科室ID & " Or ID=" & .执行部门ID
                        Case 2 '病人病区
                            mrsUnit.Filter = "ID=" & lng病区ID & " Or ID=" & .执行部门ID
                        Case 3 '操作员科室
                            mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                        Case 4 '指定科室
                            strSQL = "" & _
                            "   Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                            "   From 收费执行科室 A,部门表 C" & _
                            "   Where A.收费细目ID=[1]　And A.执行科室ID+0=C.ID " & _
                            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                            "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                            "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                            " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .收费细目ID, gint病人来源, lng科室ID)
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
                        Case 5 '院外执行(预留,程序暂未用)
                        Case 6 '开单人科室
                           mrsUnit.Filter = "ID=" & Get开单科室ID & " Or ID=" & .执行部门ID
                    End Select
                    If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                    Set rsUnit = Rec.CopyNew(mrsUnit)
                    If Not rsUnit.EOF Then
                        For i = 1 To rsUnit.RecordCount
                            strTmp = rsUnit!编码 & "-" & rsUnit!名称
                            If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                Bill.AddItem strTmp
                                Bill.ItemData(Bill.ListCount - 1) = rsUnit!ID
                                                                
                                '设置缺省执行科室
                                If lngRow = 1 Then
                                    If rsUnit!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                ElseIf lngRow > 1 Then
                                    '与上一行非药品相同
                                    If rsUnit!ID = mobjBill.Pages(1).Details(lngRow - 1).执行部门ID And _
                                        mobjBill.Pages(1).Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 Then
                                        Bill.ListIndex = Bill.NewIndex
                                    ElseIf rsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                        Bill.ListIndex = Bill.NewIndex
                                    End If
                                End If
                            End If
                            rsUnit.MoveNext
                        Next
                    End If
                        
                    If .Detail.执行科室 = 4 Then     '执行科室为指定科室的,缺省为操作员所在科室
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FillDoctor(Optional lng科室ID As Long)
'功能：根据指定的开单科室ID读取并填写医生列表,但不缺省医生
    Dim lngOldID As Long
    
    cbo开单人.Clear
    Call GetDoctor(lng科室ID, mrs开单人)
    
    Do While Not mrs开单人.EOF
        If lngOldID <> mrs开单人!ID Then
            If gbyt开单人显示 = 1 Then
                cbo开单人.AddItem mrs开单人!简码 & "-" & mrs开单人!姓名
            Else
                cbo开单人.AddItem mrs开单人!编号 & "-" & mrs开单人!姓名
            End If
            cbo开单人.ItemData(cbo开单人.NewIndex) = mrs开单人!ID
            lngOldID = mrs开单人!ID
        End If
        mrs开单人.MoveNext
    Loop
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtInvoice.Text) = txtInvoice.MaxLength And KeyAscii <> 8 And txtInvoice.SelLength <> Len(txtInvoice) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtRePrint_GotFocus()
    zlControl.TxtSelAll txtRePrint
End Sub

Private Sub txtRePrint_KeyPress(KeyAscii As Integer)
    Dim strNo As String, strOper As String, vDate As Date
    Dim strReclaimIvoice As String  '回收票据
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtRePrint, KeyAscii)
    Else
        '重打
        txtRePrint.Text = GetFullNO(txtRePrint.Text, 13)
        zlControl.TxtSelAll txtRePrint
       
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("门诊费用记录", txtRePrint.Text, , "1") Then
            If Not ReturnMovedExes(txtRePrint.Text, 1, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        If Not ReadBillInfo(1, txtRePrint.Text, 1, strOper, vDate) Then
            txtRePrint.SetFocus: Exit Sub
        End If
        If InStr(mstrPrivs, "所有操作员") <= 0 Then
            If UserInfo.姓名 <> strOper Then
                MsgBox "你没有""所有操作员""权限,不能重打" & strOper & "的单据！", vbInformation, gstrSysName
                txtRePrint.Text = "": Exit Sub
            End If
        End If
        If Not BillOperCheck(2, strOper, vDate, "重打", txtRePrint.Text, , 1) Then
            txtRePrint.SetFocus: Exit Sub
        End If
        
        
        strNo = "'" & txtRePrint.Text & "'"
        '单据有剩余数量的才可以重打
        If Not BillExistMoney(strNo, 1, True) Then
            MsgBox "单据不存在或已经全部退费,不能重打！", vbInformation, gstrSysName
            txtRePrint.Text = "": Exit Sub
        End If
        
        '56963
        strReclaimIvoice = zlGetReclaimInvoice(strNo)
        If strReclaimIvoice <> "" Then
            Call MsgBox("注意:" & vbCrLf & " 请注意回收如下发票:" & vbCrLf & strReclaimIvoice, vbOKOnly + vbInformation, gstrSysName)
        End If
        Dim intInvoiceFormat As Integer
        intInvoiceFormat = IIf(strReclaimIvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
        
        Dim strPriceGrade As String
        If gintPriceGradeStartType >= 2 Then
            strPriceGrade = GetPriceGradeFromNos(strNo)
        Else
            strPriceGrade = mstr普通价格等级
        End If
        If Not RePrintCharge(1, strNo, Me, mlng领用ID, strReclaimIvoice, , , intInvoiceFormat, , , mlngShareUseID, _
            mstrUseType, , strPriceGrade) Then
            txtRePrint.SetFocus
        Else
            Call RefreshFact
            txtRePrint.Text = ""
            txtPatient.SetFocus
        End If
    End If
End Sub

Private Sub RefreshFact()
'功能：刷新收费票据号
    If gblnStrictCtrl Then
        If zlGetInvoiceGroupUseID(mlng领用ID) = False Then
            txtInvoice.Text = "": Exit Sub
        End If
        '严格：取下一个号码
        txtInvoice.Text = GetNextBill(mlng领用ID)
    Else
        '松散：取下一个号码
        txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModul)))
    End If
End Sub

Private Function CalcBillToTal(Optional bln应收 As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Integer, intCol As Integer

    If mobjBill.Pages(1).Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Pages(1).Details
            For Each objTmpIncome In objTmpDetail.InComes
                If bln应收 Then
                    CalcBillToTal = CalcBillToTal + objTmpIncome.应收金额
                Else
                    CalcBillToTal = CalcBillToTal + objTmpIncome.实收金额
                End If
            Next
        Next
    Else
        For i = 0 To Bill.COLS - 1
            If bln应收 Then
                If Bill.TextMatrix(0, i) = "应收金额" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "实收金额" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcBillToTal = CalcBillToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
    CalcBillToTal = Format(CalcBillToTal, gstrDec)
End Function

Private Function Calc工本费() As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome

    For Each objTmpDetail In mobjBill.Pages(1).Details
        If objTmpDetail.工本费 Then
            For Each objTmpIncome In objTmpDetail.InComes
                Calc工本费 = Calc工本费 + objTmpIncome.实收金额
            Next
        End If
    Next
End Function

Private Sub txt缴款_LostFocus()
    mblnHotKey = False
    If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
End Sub

Private Function SaveModi() As Boolean
'功能：保存当前修改的费用单据
    Dim strSQL As String
    
    strSQL = "zl_病人费用记录_Update('" & cboNO.Text & "',1," & _
        "'" & zlStr.NeedName(cbo开单人.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'))"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FillDept(Optional lng人员ID As Long)
'功能：读取并加载科室列表,但不缺省科室
'参数：lng人员ID=只读取指定人员所在科室(包含非缺省的)
'返回：科室个数
    
    Dim strSQL As String, i As Long, lngOldDepID As Long
    Dim strDepts As String  '指定人员所属的多个部门
        
    cbo开单科室.Clear
    If mrs开单科室 Is Nothing Then Call GetDoctorDept(mrs开单科室)
   
    If lng人员ID <> 0 Then
        If Not mrs开单人 Is Nothing Then
            mrs开单人.Filter = "ID=" & lng人员ID
            For i = 1 To mrs开单人.RecordCount
                strDepts = strDepts & " OR ID=" & mrs开单人!部门ID      'filter不支持in
                mrs开单人.MoveNext
            Next
        End If
        If strDepts <> "" Then
            mrs开单科室.Filter = Mid(strDepts, 4)
        Else
            mrs开单科室.Filter = "ID=0" '人员没有设置部门,不显示开单科室
        End If
    Else
        mrs开单科室.Filter = ""
    End If
    
    If mrs开单科室.RecordCount > 0 Then
        For i = 1 To mrs开单科室.RecordCount
            If lngOldDepID <> mrs开单科室!ID Then   '一个部门可能同时属于手术和临床,不加载相同的
                cbo开单科室.AddItem mrs开单科室!编码 & "-" & mrs开单科室!名称
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = mrs开单科室!ID
                lngOldDepID = mrs开单科室!ID
            End If
            mrs开单科室.MoveNext
        Next
    End If
End Sub

Private Function Check费用类型(Optional intRow As Integer) As Boolean
'功能：根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    Dim strSQL As String
    Dim i As Integer, strType As String
    Dim rsTmp As New ADODB.Recordset
    
    Check费用类型 = True
    
    On Error GoTo errHandle
    
    '无法检查
    If cbo医疗付款.ListIndex = -1 Then Exit Function
    
    '确定病人类型
    strType = cbo医疗付款.Text
    
    '只检查医保病人和公费病人
    If strType <> "1" And strType <> "2" Then Exit Function
    
    '读取检查数据
    If strType = "1" Then
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.EOF Then Exit Function
    
    If intRow > 0 Then
        If mobjBill.Pages(1).Details(intRow).Detail.类型 = "" Then
            MsgBox """" & mobjBill.Pages(1).Details(intRow).Detail.名称 & """的费用类型未设置！", vbInformation, gstrSysName
            Check费用类型 = False
        Else
            rsTmp.Filter = "名称='" & mobjBill.Pages(1).Details(intRow).Detail.类型 & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Pages(1).Details(intRow).Detail.名称 & """的类型为""" & _
                    mobjBill.Pages(1).Details(intRow).Detail.类型 & """,不是" & _
                    IIf(strType = "1", "医保", "公费") & "费用类型！", vbInformation, gstrSysName
                Check费用类型 = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).Detail.类型 = "" Then
                If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Pages(1).Details(i).Detail.名称 & """的费用类型未设置！" & vbCrLf & "确实要保存单据吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check费用类型 = False: Exit For
                End If
            Else
                rsTmp.Filter = "名称='" & mobjBill.Pages(1).Details(i).Detail.类型 & "'"
                If rsTmp.EOF Then
                    If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Pages(1).Details(i).Detail.名称 & """的费用类型为""" & _
                        mobjBill.Pages(1).Details(i).Detail.类型 & """,不是" & _
                        IIf(strType = "1", "医保", "公费") & "费用类型！" & vbCrLf & "确实要保存单据吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check费用类型 = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt缴款_Validate(Cancel As Boolean)
    If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
End Sub

Private Sub LoadAndSeek费别()
    Dim lngDeptID As Long
     
    '费别处理
    If cbo开单科室.ListIndex <> -1 Then lngDeptID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    
    '读取唯一性费别(都含仅限初诊,以便定位)及动态费别
    Call Load费别(cbo费别, lngDeptID, True, mrs费别)
    
    '读取动态费别,默认为可见
    If lbl动态费别.Visible Then     '窗体默认为True
        lbl动态费别.Caption = Load动态费别(lngDeptID)
        lbl动态费别.Tag = lbl动态费别.Caption
        lbl动态费别.Visible = lbl动态费别.Caption <> ""
        If lbl动态费别.Caption <> "" Then lbl动态费别.Caption = "(" & lbl动态费别.Caption & ")"
    End If
    
    If mrsInfo.State = 0 Then
        '输入姓名病人可以自由选择
        cbo费别.Locked = False
        If cbo费别.ListIndex = -1 And cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
    ElseIf mrsInfo.State = 1 Then
        '定位有档案病人的费别
        cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(mrsInfo!费别), True)
        If cbo费别.ListIndex <> -1 Then
            '再判断初诊是否合适
            If cbo费别.ItemData(cbo费别.ListIndex) = 2 And mrsInfo!初诊 = 0 Then
                '使用缺省费别(不含仅限初诊费别)
                Call Load费别(cbo费别, lngDeptID, False, mrs费别)
                If cbo费别.ListIndex <> -1 Then
                    If Visible Then MsgBox "病人使用仅限初诊的费别:" & mrsInfo!费别 & ",但病人不是第一次就诊,将使用缺省费别！", vbInformation, gstrSysName
                Else
                    cbo费别.Locked = False '无法确定,让其自由选择
                    If Visible Then MsgBox "病人使用仅限初诊的费别:" & mrsInfo!费别 & ",但病人不是第一次就诊,请选择一种费别！", vbInformation, gstrSysName
                    If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus
                End If
            Else
                'cbo费别.Locked = True '定位到了对应费别,不可修改
            End If
        Else
            '使用缺省费别(不含仅限初诊费别)
            Call Load费别(cbo费别, lngDeptID, False, mrs费别)
            If cbo费别.ListIndex <> -1 Then
            
                If Visible Then MsgBox "没有发现病人费别:" & mrsInfo!费别 & ",将使用缺省费别！", vbInformation, gstrSysName
            Else
                cbo费别.Locked = False '无法确定,让其自由选择
                If Visible Then MsgBox "没有发现病人费别:" & mrsInfo!费别 & "和缺省费别,请选择一种费别！", vbInformation, gstrSysName
                If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus
            End If
        End If
    End If
End Sub

Private Function ItemExist(lng收费细目ID As Long) As Boolean
    Dim i As Long
    
    If mobjBill Is Nothing Then Exit Function
    
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).收费细目ID = lng收费细目ID Then
            ItemExist = True: Exit Function
        End If
    Next
End Function

Private Function Check执行科室() As Integer
    Dim i As Integer
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).执行部门ID = 0 Or Bill.TextMatrix(i, 3) = "" Then
            Check执行科室 = i: Exit Function
        End If
    Next
End Function
Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(mobjBill.病人ID, 0, 0)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    If blnFact Then RefreshFact
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng领用ID = GetInvoiceGroupID(1, intNum, lng领用ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mstrUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mstrUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
                If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 
 
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If mbytInState = 1 Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    
    Dim objCard As Card
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKind.Cards.按缺省卡查找
End Sub

Private Function IsCheck误差费() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查误差费是否正常设置
    '返回:正常设置,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 15:17:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gstr误差费名称 <> "" Then IsCheck误差费 = True: Exit Function
    Select Case mbytInState
        Case 0
            MsgBox "系统中尚未设置有效的误差处理,请在[结算方式管理]中设置。", vbInformation, gstrSysName
            Exit Function
        Case Else
            IsCheck误差费 = True
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
