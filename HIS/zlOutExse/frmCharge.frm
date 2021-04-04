VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "病人收费处理"
   ClientHeight    =   8148
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8148
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTemp 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   0
      TabIndex        =   94
      TabStop         =   0   'False
      Text            =   "cbo开单人"
      ToolTipText     =   "支持输入简码和编号自动匹配"
      Top             =   20000
      Width           =   2145
   End
   Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
      Height          =   375
      Left            =   -135
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7650
      Width           =   11310
      _cx             =   19950
      _cy             =   661
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   12632256
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   360
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCharge.frx":08CA
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fra退费摘要 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   15
      TabIndex        =   80
      Top             =   5160
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txt退费摘要 
         Height          =   360
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   16
         Top             =   0
         Width           =   5820
      End
      Begin VB.Label lbl摘要 
         Caption         =   "退费摘要"
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.Frame fraSubBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   15
      TabIndex        =   69
      Top             =   5160
      Visible         =   0   'False
      Width           =   11865
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "配方合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4440
         TabIndex        =   79
         Top             =   45
         Width           =   1155
      End
      Begin VB.Label lblDuty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开单人专业职务:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   74
         Top             =   45
         Width           =   1800
      End
      Begin VB.Label lblSub应收 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收:0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7245
         TabIndex        =   71
         Top             =   45
         Width           =   1185
      End
      Begin VB.Label lblSub实收 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收:0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9345
         TabIndex        =   70
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.Frame fraBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   15
      TabIndex        =   66
      Top             =   1830
      Width           =   11820
      Begin VB.CommandButton cmdDelBill 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10850
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "删除当前单据(ALT+D)"
         Top             =   30
         Width           =   960
      End
      Begin VB.CommandButton cmdAddBill 
         Caption         =   "增加(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9870
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "增加一张单据(F12)"
         Top             =   30
         Width           =   960
      End
      Begin MSComctlLib.TabStrip tbsBill 
         Height          =   705
         Left            =   30
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   1249
         TabWidthStyle   =   2
         TabFixedWidth   =   2117
         TabFixedHeight  =   616
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "单据&1"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox cbo开单科室 
      Height          =   360
      Left            =   1200
      TabIndex        =   7
      Text            =   "cbo开单科室"
      Top             =   1410
      Width           =   2010
   End
   Begin VB.Frame fraTitle 
      Height          =   1080
      Left            =   0
      TabIndex        =   45
      ToolTipText     =   "清除:F6"
      Top             =   -120
      Width           =   11880
      Begin VB.CommandButton cmdSaveWholeSet 
         Caption         =   "保存为成套收费项目(&W)"
         Height          =   375
         Left            =   6630
         TabIndex        =   85
         Top             =   195
         Width           =   2715
      End
      Begin VB.CommandButton cmdSelWholeSet 
         Caption         =   "成套(&T)"
         Height          =   375
         Left            =   5520
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "医保"
         Height          =   375
         Left            =   1080
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "热键：F6"
         Top             =   660
         Width           =   720
      End
      Begin VB.CommandButton cmdIDCard 
         Caption         =   "医疗卡(&K)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   72
         ToolTipText     =   "热键：F10"
         Top             =   195
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdRegist 
         Caption         =   "挂号(&E)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10725
         TabIndex        =   42
         ToolTipText     =   "热键：F3"
         Top             =   195
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtRePrint 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2500
         MaxLength       =   8
         TabIndex        =   35
         Top             =   667
         Width           =   1065
      End
      Begin VB.TextBox txtModi 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4250
         MaxLength       =   8
         TabIndex        =   37
         Top             =   667
         Width           =   1065
      End
      Begin VB.CommandButton cmd配方 
         Caption         =   "配方(&R)"
         Height          =   375
         Left            =   80
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   660
         Width           =   1000
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   1545
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H80000007&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9975
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "定位:F9,单据号长度不足时自动补足长度"
         Top             =   660
         Width           =   1350
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   390
         Left            =   11370
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   645
         Visible         =   0   'False
         Width           =   435
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
         Height          =   390
         Left            =   11370
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   645
         Width           =   435
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5985
         MaxLength       =   8
         TabIndex        =   39
         ToolTipText     =   "从已有的单据中复制信息,不影响已有单据"
         Top             =   667
         Width           =   1065
      End
      Begin VB.TextBox txtMCInvoice 
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   7110
         X2              =   7110
         Y1              =   630
         Y2              =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   7100
         X2              =   7100
         Y1              =   630
         Y2              =   1050
      End
      Begin VB.Label lblRePrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打(&P)"
         Height          =   240
         Left            =   1900
         TabIndex        =   34
         Top             =   727
         Width           =   600
      End
      Begin VB.Label lblIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导(&I)"
         Height          =   240
         Left            =   5350
         TabIndex        =   38
         Top             =   727
         Width           =   600
      End
      Begin VB.Label lblModi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "改(&M)"
         Height          =   240
         Left            =   3650
         TabIndex        =   36
         Top             =   727
         Width           =   600
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   9360
         TabIndex        =   67
         Top             =   255
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票号"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7200
         TabIndex        =   43
         Top             =   720
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   38015
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   38000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   11370
         TabIndex        =   53
         Top             =   645
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "病人收费单"
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
         Left            =   135
         TabIndex        =   48
         ToolTipText     =   "清除:F6"
         Top             =   195
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "单据号"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   9250
         TabIndex        =   46
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblCorp 
         Caption         =   "工作单位:"
         Height          =   255
         Left            =   5280
         TabIndex        =   73
         Top             =   960
         Visible         =   0   'False
         Width           =   5895
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   47
      Top             =   7785
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2625
            MinWidth        =   882
            Picture         =   "frmCharge.frx":0994
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9991
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   88
            Object.Tag             =   "用于记帐或收费个人帐户显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   370
            MinWidth        =   88
            Object.Tag             =   "用于收费预交显示"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   360
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Picture         =   "frmCharge.frx":1228
            Key             =   "Drugstore"
            Object.ToolTipText     =   "药房设置"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Key             =   "PatiSource"
            Object.ToolTipText     =   "病人来源"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmCharge.frx":1542
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":1C1C
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":2256
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1101
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1101
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraInfo 
      Height          =   990
      Left            =   0
      TabIndex        =   44
      Top             =   840
      Width           =   11880
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   390
         Left            =   555
         TabIndex        =   92
         Top             =   180
         Width           =   630
         _ExtentX        =   1101
         _ExtentY        =   699
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
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
         MustSelectItems =   "姓名,就诊卡"
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   360
         Left            =   5750
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   580
      End
      Begin VB.TextBox txt门诊号 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   9650
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   180
         Width           =   2145
      End
      Begin VB.CheckBox chk急诊 
         Alignment       =   1  'Right Justify
         Caption         =   "急诊"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8040
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   790
      End
      Begin VB.ComboBox cbo医疗付款 
         Height          =   360
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   2505
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "定位:F6,输入:-病人ID,*门诊号,+住院号,.挂号单号,例如:*2536表示按门诊号查找"
         Top             =   180
         Width           =   1470
      End
      Begin VB.ComboBox cboSex 
         Height          =   360
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.ComboBox cbo费别 
         Height          =   360
         Left            =   3765
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   570
         Width           =   1575
      End
      Begin VB.Label lbl险类 
         Alignment       =   1  'Right Justify
         Caption         =   "险类"
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   8910
         TabIndex        =   76
         Top             =   630
         Width           =   2880
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   240
         Left            =   8910
         TabIndex        =   68
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl动态费别 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5520
         TabIndex        =   64
         Top             =   600
         Width           =   2370
      End
      Begin VB.Label lbl科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   240
         Left            =   100
         TabIndex        =   10
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   75
         TabIndex        =   52
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Left            =   2680
         TabIndex        =   51
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Left            =   4395
         TabIndex        =   50
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   240
         Left            =   3240
         TabIndex        =   49
         Top             =   630
         Width           =   480
      End
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   0
      ScaleHeight     =   2280
      ScaleWidth      =   11280
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5505
      Width           =   11280
      Begin VB.CommandButton cmd预结算 
         Caption         =   "预结算(&V)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10305
         TabIndex        =   29
         ToolTipText     =   "热键：F5"
         Top             =   540
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10305
         TabIndex        =   30
         ToolTipText     =   "热键F2,右键弹出保存为划价单(或按CTRL+S)"
         Top             =   975
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1770
         Left            =   5415
         TabIndex        =   91
         Top             =   495
         Width           =   2445
         _cx             =   4313
         _cy             =   3122
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483630
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   400
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCharge.frx":2890
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "完成收费(&F)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   31
         ToolTipText     =   "热键：Alt+F"
         Top             =   1860
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   32
         ToolTipText     =   "热键:Esc"
         Top             =   1410
         Width           =   1440
      End
      Begin VB.TextBox txtTmp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   6510
         MaxLength       =   10
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   570
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   0
         TabIndex        =   56
         ToolTipText     =   "清除:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   165
            Width           =   1800
         End
         Begin VB.ComboBox cbo结算方式 
            Height          =   360
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   165
            Width           =   1680
         End
         Begin VB.CheckBox chk加班 
            Caption         =   "加班(&L)"
            Height          =   270
            Left            =   80
            TabIndex        =   17
            Top             =   210
            Width           =   1170
         End
         Begin VB.ComboBox cbo开单人 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6615
            TabIndex        =   23
            Text            =   "cbo开单人"
            ToolTipText     =   "支持输入简码和编号自动匹配"
            Top             =   165
            Width           =   2145
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9390
            TabIndex        =   24
            Top             =   165
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
            Left            =   1320
            TabIndex        =   18
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label lbl开单人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "开单人(&W)"
            Height          =   240
            Left            =   5505
            TabIndex        =   22
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   240
            Left            =   8880
            TabIndex        =   57
            Top             =   225
            Width           =   480
         End
         Begin VB.Label lbl结算方式 
            AutoSize        =   -1  'True
            Caption         =   "结算方式(&X)"
            Height          =   240
            Left            =   2400
            TabIndex        =   20
            Top             =   225
            Width           =   1320
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1770
         Left            =   15
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   510
         Width           =   2775
         _ExtentX        =   4890
         _ExtentY        =   3112
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "^序号|^项目     |^    金额|^     合计"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Frame fraStat 
         Height          =   1905
         Left            =   2865
         TabIndex        =   58
         Top             =   375
         Width           =   2490
         Begin VB.TextBox txt合计 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   735
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "0.00"
            ToolTipText     =   "连续收费时未缴款单据的实收金额合计"
            Top             =   810
            Width           =   1650
         End
         Begin VB.TextBox txt应收 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "0.00"
            ToolTipText     =   "连续收费时未缴款单据的应收金额合计"
            Top             =   285
            Width           =   1650
         End
         Begin VB.TextBox txt累计 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1350
            Width           =   1650
         End
         Begin VB.Label lbl合计 
            AutoSize        =   -1  'True
            Caption         =   "实收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.6
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   885
            Width           =   660
         End
         Begin VB.Label lbl应收 
            AutoSize        =   -1  'True
            Caption         =   "应收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.6
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   60
            TabIndex        =   60
            Top             =   345
            Width           =   690
         End
         Begin VB.Label lbl累计 
            AutoSize        =   -1  'True
            Caption         =   "累计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.6
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   60
            TabIndex        =   59
            Top             =   1410
            Width           =   690
         End
      End
      Begin MSComctlLib.ImageList imgPati 
         Left            =   4875
         Top             =   1875
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCharge.frx":28DE
               Key             =   "InPati"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCharge.frx":31B8
               Key             =   "OutPati"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraUpBillShow 
         Caption         =   "上张单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5325
         TabIndex        =   81
         Top             =   540
         Visible         =   0   'False
         Width           =   1920
         Begin VB.TextBox txtPreMoney 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "上张单据金额"
            Top             =   1005
            Width           =   1710
         End
         Begin VB.TextBox txtPreNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "上张单据号"
            Top             =   495
            Width           =   1710
         End
      End
      Begin VB.Frame fra缴款 
         Height          =   1905
         Left            =   7860
         TabIndex        =   86
         Top             =   375
         Width           =   2325
         Begin VB.TextBox txt预交冲款 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   450
            Left            =   795
            TabIndex        =   89
            Text            =   "0.00"
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txt应缴 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   87
            TabStop         =   0   'False
            Text            =   "0.00"
            ToolTipText     =   "连续收费时,指应缴合计"
            Top             =   765
            Width           =   1395
         End
         Begin VB.Label lbl预交冲款 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "预存款"
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   60
            TabIndex        =   90
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl应缴 
            AutoSize        =   -1  'True
            Caption         =   "退  款"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   45
            TabIndex        =   88
            Top             =   870
            Width           =   720
         End
      End
      Begin VB.Label lblSeek 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用于按钮定位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   10200
         TabIndex        =   65
         Top             =   585
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblTotal 
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
         Left            =   225
         TabIndex        =   62
         Top             =   585
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   120
      Top             =   75
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2925
      Left            =   0
      TabIndex        =   14
      Top             =   2220
      Width           =   11865
      _ExtentX        =   20934
      _ExtentY        =   5165
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
      ColWidth0       =   1008
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
      cboStyle        =   0
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
      TabIndex        =   54
      Top             =   0
      Width           =   945
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnuFileSavePrice 
         Caption         =   "保存为划价单(&S)"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private Const M_MONEY_ROWS = 6 '左下角项目列表可显示行数
'————————————————————————————————————————————————————————————————————
'入口参数：
Public mbytInFun As Byte '0-收费,1-划价,2-门诊记帐
Public mbytInState As Byte '0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费),4-重新收费;5-异常单据作废
Public mstrInNO As String '操作的单据号(查看，调整，修改，退费，销帐,重新收费时)
Public mblnCopyBill As Boolean '是否自动复制产生的单据
Public mblnNOMoved As Boolean '操作的单据是否在后备数据表中
Public mstrTime As String '操作单据内容的登记时间
Public mblnDelete As Boolean '是否处理退费单据(查阅)
Public mbytBilling As Byte 'mbytInFun=2时：0-正常记帐,1-记帐划价,2-记帐审核
Public mstrPrivs As String
Public mlngModul As Long
Public mlngFirstID As Long '记录被修改单据第一药品行的执行部门ID,用于收费
Public mstrFirstWin As String '记录被修改单据第一药品行的发药窗口,用于收费
Public mbln退费异常 As Boolean '异常冲销单据

'门诊留观门诊记账补费相关变量
Public mlng病人ID As Long
Public mlng主页ID As Long
Public mlngUnitID As Long '当前记帐病区,为0时表示所有病区
Public mlngDeptID As Long '当前记帐科室,为0时表示所有科室
Public mbln补费 As Boolean '33744
Public mlng关联医嘱 As Long
Public mstr最后转科时间 As String

'消息相关对象变量
Public WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------------------------------------------
Private mstr应付款结算方式 As String    '33722
Private mint退费回单打印 As Integer '退费回单打印方式 0-不打印,1-自动打印,2-选择是否打印
Private mblnSaveAsPrice As Boolean '联合医保：收费时是否保存为划价单
Private mintReturnMode As Integer   '用于退费时,全退禁用结算方式时恢复初始的结算方式
Private mblnNotValied As Boolean '不处理效点失效问题
Private mblnNotClick As Boolean
Private mstrBalance As String
Private mblnHaveExcuteData As Boolean '是否医嘱计价中存在数据:60735
'————————————————————————————————————————————————————————————————————
Private mblnErrBill As Boolean
'数据对象
Private mrs结算方式 As ADODB.Recordset
Private mrsWork As ADODB.Recordset      '当天上班的药房
Private mrsClass As ADODB.Recordset     '根据参数读取的当前可用的收费类别
Private mrsUnit As ADODB.Recordset      '可选择的执行科室
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室
Private mrs开单人 As ADODB.Recordset    '所用医生和护士列有
Private mrsInfo As ADODB.Recordset      '病人信息
Private mrsWarn As ADODB.Recordset      '病区报警线记录集
Private mrs费别 As ADODB.Recordset      '所有费别及适用科室
Private mrs费用类型 As ADODB.Recordset  '所有费用类型
Private mrs发药窗口 As ADODB.Recordset  '发药窗口清单,用于判断药房是否指定了发药窗口
'程序对象
Private mobjBill As ExpenseBill '费用单据对象
Private mcolMoneys As BillInComes  '所有单据的收入项目汇总集合
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mobjDetail As Detail '单独的收费细目对象
Private mcolDetails As Details '单独的收费细目集合
Private mrs收费对照 As ADODB.Recordset '收费对照 :问题:33634

Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mintOldInvoiceFormat As Integer '旧发票格式打印
Private mblnStartFactUseType As Boolean   '是否启用了使用类别
Private mintInvoicePrint As Integer  '0-不打印;1-自动打印;2-提示打印
Private mblnFirst As Boolean
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
    从属父号 = 4
    规格 = 5
    单位 = 6
    付数 = 7
    数次 = 8
    单价 = 9
    应收金额 = 10
    实收金额 = 11
    执行科室 = 12
    标志 = 13
    医嘱序号 = 14
    类型 = 15
    执行科室ID = 16
End Enum

'程序变量
Private mintPage As Integer '当前是第几张单据
Private mstrWarn As String '已经报过警并选择继续的类别
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库的出库检查方式

Private mlngPreRow As Long '当前行号,用于列改变时判断
Private mlng药品类别ID As Long '当前单据操作的药品入出类别ID
Private mlng卫材类别ID As Long '当前单据操作的卫材入出类别ID

Private mbln处方职务检查 As Boolean     '是否进行处方职务检查
Private mbln处方限量检查 As Boolean     '是否进行处方限量检查
Private mbln储备限额检查 As Boolean     '是否进行储备限额检查

Private mcolBalance As Collection '记录各张单据保险结算原始值及修改值
Private mcolRquareBalance As Collection '刘兴洪:增加了消费卡的结算内容

Private mblnHotKey As Boolean '手工报价时,是否才按了报价热键
Private mbln报合计 As Boolean
Private mstrCardNO As String '就诊卡划价单据号
Private mstr付款方式 As String '缺省医疗付款方式
Private mbytBillSource As Byte   '单据来源:1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检

Private mstrPrePati As String  '上一个收费病人
Private mlngPrePati As Long     '上一个收费病人ID
Private mstrPreDoctor As String  '记录前一开单人

Private mstr西窗 As String, mstr成窗 As String, mstr中窗 As String '记录门诊病人连续收费的窗口分配
Private mlng西药房 As Long, mlng成药房 As Long, mlng中药房 As Long '记录门诊病人连续收费的药房分配
Private mblnNewRow As Boolean '表示是否人为加行
Private mlng领用ID As Long '收费票据的领用批次
Private mbln不重算价格 As Boolean     '在修改和导入单据时,设置费别时不重算价格,读入时会算,后面也会重算

Private mblnF2Save As Boolean   '是否按F2保存
Private mblnValid As Boolean '是否因为焦点丢失
Private mblnDo As Boolean           '控制加班_click事件是否激活
Private mblnDoing As Boolean        '控制是否正在读单据信息
Private mblnEnterCell As Boolean    '控制是否激活EnterCell事件
Private mblnDrop As Boolean         '在KeyDown中判断cbo开单人当前是否弹出
Private mblnCboClick As Boolean      '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
'收费处同一病人病人单据累计金额
Private mcurBill应收 As Currency
Private mcurBill实收 As Currency
Private mcurBill应缴 As Currency
Private mdbl缴款 As Double, mdbl找补 As Double
Private mbln连续输入 As Boolean     '确定当前单据是否连续输入:44944
Private mintBillNO As Integer '病人当前连续收了几张单子
Private mintMoneyRow As Integer '当前显示到的费目行
Private mblnLoad As Boolean
Private mblnOne As Boolean '是否只有一个可用收费类别
Private marrColData() As Integer '当前单据编辑属性映象
Private mblnPrint As Boolean '收费时是否打印票据,有两种:本地参数设置是否打印,费用为0是否打印
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器

Private Const STR_HEAD = "行,450,4;类别,750,1;项目,2175,1;商品名,2000,1;从属父号,0,0;规格,1105,1;单位,520,4;付数,520,1;数次,570,1;单价,1055,7;" & _
    "应收金额,1030,7;实收金额,1080,7;执行科室,1255,1;标志,520,4;医嘱序号,0,0;类型,520,1;执行科室ID,0,1"

'医保相关
Private mintInsure As Integer
Private mstrYBPati As String 'New:空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
Private mstr个人帐户 As String '是否将个人帐户设置到收费可用
Private mcur个帐余额 As Currency '当前病人个人帐户余额
Private mcur个帐透支 As Currency '个人帐户允许透支金额
Private mblnYB结算作废 As Boolean '医保是否支持结算作废,用于退费时判断
Private mstrYBBill As String '医保病人连续收费的单据号
Private mlng结算序号  As Long '重新收费时有效
Private mblnOneCard As Boolean      '是否启用了一卡通接口
Private mrsOneCard As ADODB.Recordset
Private mrsDelInvoice As ADODB.Recordset

'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    允许不设置医保项目 As Boolean
    门诊收费存为划价单 As Boolean
    不提醒缴款金额不足 As Boolean    '27536
    门诊必须传递明细 As Boolean
    医保接口打印票据 As Boolean
    医生确定处方类型 As Boolean
    多单据一次结算 As Boolean
    门诊连续收费 As Boolean
    门诊预结算 As Boolean
    多单据收费 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    先自付 As Boolean
    全自付 As Boolean
    blnOnlyBjYb As Boolean '本地仅支持北京医保:刘兴洪
    退费后打印回单 As Boolean '
    多单据调一次交易 As Boolean
    医保不走票号  As Boolean        '预结算时有效
End Type
Private MCPAR As TYPE_MedicarePAR

Private Type TYPE_Original
    实收合计 As Currency    '门诊记帐,记录修改单据时的原单据实收金额合计
    应缴金额 As Currency    '收费,记录修改单据时的应缴金额
    冲预交款 As Currency    '收费,记录修改单据时的原始预交冲款金额
    结帐ID As Long          '退费,记录原单据结帐ID
End Type
Private Original As TYPE_Original
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mblnAutoChangePati As Boolean '当前的在院模式是自动切换到在院模式的

Private Type Ty_ModulePara
    bln住院病人门诊收费 As Boolean    '住院病人完全门诊收费
    '以后扩展
End Type
Private mTy_Para As Ty_ModulePara
Private mobjBaseItem As Object
Private Enum Pan
    C2提示信息 = 2
    C3个人帐户 = 3
    C4预交信息 = 4
    C5医保大类 = 5
End Enum
Private mblnSaveData As Boolean  '是否数据保存成功
Private mblnKeyReturn As Boolean '是否按了回车的
Private mrsErrBlance As ADODB.Recordset  '异常单据的结算信息
Private Type Ty_DelFee  '退费相关
      strNos  As String          '当前退费单所涉及的单据号,用逗号分离
      rsBlance As ADODB.Recordset
      blnSingleBalance As Boolean '单结算方式
      dblCurDelMoney As Double '当前退款合计
      bln三方卡全退 As Boolean
End Type
Private mTyDelFee As Ty_DelFee
Private mblnNotClearLedDisplay As Boolean   '不清除显示
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
Private mlngPreBrushCard As Long  '上次刷卡的卡类别ID
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private WithEvents mFrmBalanceWin   As frmChargePayMentWin
Attribute mFrmBalanceWin.VB_VarHelpID = -1
'-----------------------------------------------------------------------------------
'数据保存相关
Private mstrModiNOs As String, mstrSaveNos As String
Private mcllPayDrugAndStuff As Collection   '保存自动发料和发药
Private mCllWindows As Collection
Private mobjDrugPacker  As Object ' 自动发药机(更新发药窗口)
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object '自动发药机(新）
Private mblnDrugMachine As Boolean
Private mblnClearBlance As Boolean '是否清除结算信息
Private mlngCardTypeID As Long   '当前提取病人信息刷的卡类别ID 56615
Private mblnOlny预交 As Boolean '仅使用预交68177
Private mstr药品价格等级 As String, mstr卫材价格等级 As String
Public mstr普通价格等级 As String
Private mblnSetControl As Boolean
Private WithEvents mobjBrushCheck As clsBrushCardInput
Attribute mobjBrushCheck.VB_VarHelpID = -1
Private mobjCard As New Card
Private mbln条码刷卡 As Boolean

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
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
    
    If mobjBill.Pages(mintPage).Details.Count < Bill.Row Then
        Exit Sub
     End If
    With mobjBill.Pages(mintPage).Details(Bill.Row)
        If InStr(",4,5,6,7,", .收费类别) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModul, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModul, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub

Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
       
End Sub

Private Sub cbo医疗付款_Click()
    On Error GoTo errHandler
    If mbytInState <> 0 Then Exit Sub
    If gintPriceGradeStartType < 2 Then Exit Sub
    
    If mrsInfo.State = adStateOpen Then
        If gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), zlStr.NeedName(cbo医疗付款.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级) = False Then Exit Sub
    Else
        If gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cbo医疗付款.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级) = False Then Exit Sub
    End If
    
    If mbln不重算价格 Then Exit Sub
    If CheckBillsEmpty() Then Exit Sub
    
    '需要重新预结算
    If cmd预结算.Visible Then
        Call InitBalanceGrid
        cmd预结算.TabStop = True
        cmdOK.Enabled = False
    End If
    
    '全部重新计算价格
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

Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lng执行科室ID As Long
    Dim rsTemp As ADODB.Recordset, dbl价格 As Double
    Dim strSQL As String
    '保存为存套收费项目
    '问题:27327
    Err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:序号,父号,收费细目ID,数量,单价,执行科室|序号,父号,收费细目ID,数量,单价,执行科室|…
    Err = 0: On Error GoTo Errhand:
   If mbytInState = 1 Then
        '查看
         strSQL = _
        " Select Nvl(A.价格父号,A.序号) as 序号,A.收费类别,A.从属父号,A.收费细目ID,A.执行部门ID," & _
        "       　   Avg(Nvl(A.付数,1)) as 付数, Avg(A.数次) 数次, Sum(A.标准单价) as 单价,B.执行科室, B.是否变价,M.跟踪在用" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录  A") & ",收费项目目录 B,材料特性 M" & _
        " Where  A.记录状态  IN(0,1,3)  And A.NO=[1]  And A.记录性质=[2]   " & _
        "               And a.收费细目ID=b.ID And a.收费细目ID=M.材料ID(+) " & _
                        IIf(mstrTime <> "", " And A.登记时间=[3]", "") & _
        "  Group by Nvl(A.价格父号,A.序号),A.收费类别,A.收费细目ID,A.从属父号,A.执行部门id,B.执行科室,B.是否变价,M.跟踪在用" & _
        " Order by 序号"
        If mstrTime <> "" Then
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, 2, CDate(mstrTime))
        Else
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, 2)
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
        Dim dbl数次 As Double, dbl单价 As Double
        
        With mobjBill.Pages(mintPage)
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
                dbl数次 = .Details(i).数次: dbl单价 = IIf(.Details(i).Detail.变价, .Details(i).InComes(1).标准单价, 0)
                If InStr(",5,6,7,", "," & .Details(i).Detail.类别) > 0 And gbln药房单位 Then
                     dbl数次 = Format(.Details(i).数次 * .Details(i).Detail.药房包装, "0.00000")
                    dbl单价 = Format(dbl单价, gstrFeePrecisionFmt)
                End If
                
                strItems = strItems & "|" & .Details(i).序号 & "," & .Details(i).从属父号 & "," & .Details(i).收费细目ID & "," & .Details(i).付数 & ","
                strItems = strItems & dbl数次 & "," & dbl单价 & "," & lng执行科室ID
             Next
             If strItems = "" Then
                MsgBox "单据未输入任何信息,不能保存为成套收费项目,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
            strItems = Mid(strItems, 2)
        End With
    End If
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, mlngModul, mstrPrivs, strItems)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelWholeSet_Click()
    '选成套项目
    '问题:34465
    Dim rsSel As ADODB.Recordset, lng病人ID As Long, lng开单部门ID As Long
    Dim tmpBill As New ExpenseBill, byt婴儿费 As Byte, Curdate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim j As Long
    
    Dim bln中药 As Boolean
    
    If mobjBill Is Nothing Then
        If mrsInfo Is Nothing Then
            MsgBox "请先选择病人,请检查!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        Else
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
        
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
    Else
        lng病人ID = mobjBill.病人ID: lng开单部门ID = mobjBill.Pages(mintPage).开单部门ID: byt婴儿费 = mobjBill.婴儿费
    End If
     
    If zlSelectWholeItems(Me, mlngModul, mstrPrivs, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    Err = 0: On Error GoTo Errhand:
    Screen.MousePointer = 11
                         
    Set tmpBill = ImportWholeSet(Me, IIf(mstrYBPati <> "", mintInsure, 0), rsSel, mlng西药房, mlng成药房, mlng中药房, _
        lng病人ID, 0, gbln药房单位, lng开单部门ID, byt婴儿费, 2, chk加班.Value = 1, _
        0, gint病人来源, UserInfo.姓名, zlStr.NeedName(cbo开单人.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级, _
        IIf(mbln补费 And mlng主页ID <> 0, mlng主页ID, 0), IIf(mbln补费 And mstr最后转科时间 <> "", mlngDeptID, 0), _
        IIf(mbln补费 And mstr最后转科时间 <> "", mlngUnitID, 0))
    
     'a.单张单据模式,清除当前单据对象及病人信息
    If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then
        Dim rsTemp As ADODB.Recordset '95473
        Set rsTemp = mrsInfo
        Call ClearFullBill(False, False, True)
        Set mrsInfo = rsTemp
        
        '问题:36764
        '支持预结算时就不固定显示个人帐户,否则显示
        If MCPAR.门诊预结算 And mintInsure <> 0 Then
            '显示预结算按钮
            cmd预结算.Enabled = True
            Call SetButton(1) '预结算,确定,取消
            cmdOK.Enabled = False
        ElseIf mstr个人帐户 <> "" Then '只有使用个人帐户才用
            Call SetButton(2) '确定,取消
            vsBalance.TextMatrix(0, 0) = mstr个人帐户
            vsBalance.TextMatrix(0, 1) = "0.00"
            vsBalance.RowData(0) = 0
        End If
        
        Set mobjBill = tmpBill
        mobjBill.费别 = zlStr.NeedName(cbo费别.Text)
        If mobjBill.Pages(1).Details.Count > 0 Then
           If mobjBill.Pages(1).Details(1).收费类别 = "7" Then
                    bln中药 = True
           End If
        End If
        If InStr(mstrPrivs, "显示开单人") = 0 Then mobjBill.Pages(mintPage).开单人 = ""
        '清除病人信息
       ' Call ClearmobjBill
    Else
        'b.多张单据模块,新增单据,保留当前单据内容及病人相关信息,
        If i > 0 Or mobjBill.Pages(mintPage).Details.Count > 0 Then
            Call AddNewBill
        End If
        mintPage = tbsBill.Tabs.Count
        
        '不需要导入病人相关信息
        With mobjBill.Pages(mintPage)
            .NO = "" '要清空以便修改时表明是直接输入的费用
            .Key = tmpBill.Pages(1).Key
            .保险金额 = tmpBill.Pages(1).保险金额
            .冲预交额 = tmpBill.Pages(1).冲预交额
            .煎法 = tmpBill.Pages(1).煎法
            .进入统筹 = tmpBill.Pages(1).进入统筹
            .开单部门ID = tmpBill.Pages(1).开单部门ID
            If InStr(mstrPrivs, "显示开单人") > 0 Then .开单人 = tmpBill.Pages(1).开单人
            .全自付 = tmpBill.Pages(1).全自付
            .实收金额 = tmpBill.Pages(1).实收金额
            .收费结算 = tmpBill.Pages(1).收费结算
            .误差金额 = tmpBill.Pages(1).误差金额
            .先自付 = tmpBill.Pages(1).先自付
            .应缴金额 = tmpBill.Pages(1).应缴金额
            .应收金额 = tmpBill.Pages(1).应收金额

        End With
        bln中药 = False
        
        For j = 1 To tmpBill.Pages(1).Details.Count
            With tmpBill.Pages(1).Details(j)
                mobjBill.Pages(mintPage).Details.Add .费别, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, .InComes, , .保险项目否, .保险大类ID, .保险编码, .摘要
                If .收费类别 = "7" Then
                    bln中药 = True
                End If
            End With
        Next
         tbsBill.Tabs(mintPage).Selected = True '不会触发Click事件
    End If
    Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
    'Call LoadAndSeek费别
    '取第一药品行
    For i = 1 To mobjBill.Pages(1).Details.Count
        If InStr(",5,6,7,", mobjBill.Pages(1).Details(i).收费类别) > 0 Then
            mlngFirstID = mobjBill.Pages(1).Details(i).执行部门ID
            mstrFirstWin = mobjBill.Pages(1).Details(i).发药窗口
            Exit For
        End If
    Next
    Bill.Active = False
    Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
    Call InitBillColumnColor
    
    If IIf(mlngPrePati = 0, mstrPrePati <> mobjBill.姓名, mlngPrePati <> mobjBill.病人ID) Then
        '新病人
        mcurBill实收 = 0:  mcurBill应收 = 0: mcurBill应缴 = 0
        mintBillNO = 0: mintMoneyRow = 0
    End If
    '修改时应保存当前操作员的名字
    mobjBill.操作员编号 = UserInfo.编号
    mobjBill.操作员姓名 = UserInfo.姓名
    Call CalcMoneys     '因为不导入病人信息,所以需要根据当前的费别重算价格
    Call ShowDetails
    Call ShowMoney
    txtIn.Text = ""
    If mbytInState = 0 And mstrInNO <> "" Then txtModi.Text = "": mstrInNO = ""
        
    '要放在mstrInNO之后,因为以此来判断是否修改单据,以加回原库存
    Call CalcDrugStock
    Bill.Active = True
    ''设置列号
    Call SetColNum
    Screen.MousePointer = 0
    If bln中药 Then
        Call cmd配方_Click
    Else
        If mstrYBPati <> "" Then
            If cmd预结算.Enabled And cmd预结算.Visible Then
                cmd预结算.SetFocus
            ElseIf cmdOK.Enabled And cmdOK.Visible Then
                cmdOK.SetFocus
            End If
        Else
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        End If
    End If
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdYB_Click()
    txtPatient.SetFocus
    Call zlCommFun.PressKey(vbKeyF6)
End Sub



Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand  As String
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        
        '问题:27364 日期:2010-01-13 15:27:50
        If mblnAutoChangePati And gint病人来源 = 2 Then
            '需要切找到病人来源1中
            gint病人来源 = 1: zlChangePatiSource (gint病人来源)
        End If
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call txtPatient_KeyPress(vbKeyReturn)
            Call SetOneCardBalance
        End If
        Exit Sub
    End If
    If objCard.接口序号 <= 0 Then Exit Sub
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, objCard.接口序号, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        If mblnAutoChangePati And gint病人来源 = 2 Then
            '需要切找到病人来源1中
            gint病人来源 = 1: zlChangePatiSource (gint病人来源)
        End If
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub SetOneCardBalance()
    Dim CurOneCard As Currency, strName As String
    
    If mblnOneCard And Not mobjICCard Is Nothing Then
        CurOneCard = mobjICCard.GetSpare(strName)
        If CurOneCard <> 0 Then
           mrsOneCard.Filter = "名称='" & strName & "'"
           If mrsOneCard.RecordCount > 0 Then
                strName = mrsOneCard!结算方式
                If zlStr.NeedName(cbo结算方式) <> strName Then zlControl.CboLocate cbo结算方式, strName
           End If
        End If
        sta.Panels(Pan.C3个人帐户).Text = "卡余额:" & Format(CurOneCard, "0.00") & "元"
        sta.Panels(Pan.C3个人帐户).Visible = True
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, _
    objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
    'Or Not Me.ActiveControl Is txtPatient : Or txtPatient.Text <> ""
    '问题:60010
    
    If txtPatient.Locked Then Exit Sub
    mblnNotClick = True
    
    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    
    txtPatient.Text = objPatiInfor.卡号
    
    Call txtPatient_KeyPress(vbKeyReturn)
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
    '当成新病人
     If (txtPatient.Text = "" Or blnNew) And objPatiInfor.姓名 <> "" Then
        txtPatient.Text = objPatiInfor.姓名
        intIndex = IDKind.GetKindIndex("姓名")
        If intIndex > 0 Then IDKind.IDKind = IDKind.GetKindIndex("姓名")
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text <> "" Then
        Call zlControl.CboLocate(cboSex, objPatiInfor.年龄)
            If IsDate(objPatiInfor.出生日期) = False Then
                 txt年龄.Text = ReCalcOld(CDate(objPatiInfor.出生日期), cbo年龄单位, mobjBill.病人ID)
            End If
        End If
    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub mFrmBalanceWin_zlSaveData(lng结算序号 As Long, str结帐IDs As String, strSaveNos As String, blnNotCommit As Boolean, blnCancel As Boolean)
    '保存数据
    Dim blnSaveBill As Boolean
    mstrModiNOs = "": mstrSaveNos = "": strSaveNos = ""
    '先检查费用是否正确:主要是进行相关的费用检查(避免并发原因造成错误)
    '问题:62981
    If CheckChargeDataValied = False Then
        blnCancel = True
        Exit Sub
    End If
    Call SaveChargeBill(lng结算序号, str结帐IDs, strSaveNos, mstrModiNOs, blnSaveBill, blnNotCommit)
    '73960,冉俊明,2014-6-17,一卡通写卡时,在病人收费管理界面完成收费后调zlMzInforWriteToCard接口传入的参数lngBalanceID为0没有传入病人的结帐id
    mlng结算序号 = lng结算序号
    blnCancel = Not blnSaveBill
    mstrSaveNos = strSaveNos
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        
        '当成新病人
        If txtPatient.Text = "" Then
            txtPatient.Text = strName
            IDKind.IDKind = IDKind.GetKindIndex("姓名")
            Call txtPatient_KeyPress(vbKeyReturn)
            If txtPatient.Text <> "" Then
                Call zlControl.CboLocate(cboSex, strSex)
                txt年龄.Text = ReCalcOld(datBirthday, cbo年龄单位, mobjBill.病人ID)
            End If
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub
Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    Dim lngPreIDKind As Long
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient And strNo <> "" Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        Dim intIndex As Integer
        intIndex = IDKind.GetKindIndex("IC卡号")
        If intIndex <= 0 Then mblnNotClick = False: Exit Sub
        IDKind.IDKind = intIndex
        txtPatient.Text = strNo
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text = "" Then Call mobjICCard.SetEnabled(False)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Sub Bill_BeforeAddRow(Row As Long)
'说明：Row为将要新增的行号,当前行号为Row-1
    Dim dbl单价 As Double, cur金额 As Currency, i As Integer
    
    'LED动态显示项目
    If mbytInFun = 0 And gblnLED And mbytInState = 0 And gblnLedDispDetail Then
        If mobjBill.Pages(mintPage).Details.Count >= Row - 1 Then
            With mobjBill.Pages(mintPage).Details(Row - 1)
                dbl单价 = 0: cur金额 = 0
                For i = 1 To .InComes.Count
                    cur金额 = cur金额 + .InComes(i).实收金额
                    dbl单价 = dbl单价 + .InComes(i).标准单价
                Next
                'LED显示
                If cur金额 <> 0 Then
                    If InStr(",5,6,7,", .Detail.类别) > 0 And gbln药房单位 Then
                        '按药房单位显示单位
                        zl9LedVoice.Display .Detail.名称, .Detail.规格, .Detail.药房单位, dbl单价, IIf(.付数 = 0, 1, .付数) * .数次, cur金额
                    Else
                        zl9LedVoice.Display .Detail.名称, .Detail.规格, .计算单位, dbl单价, IIf(.付数 = 0, 1, .付数) * .数次, cur金额
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub ShowGroupLED(ByVal lngMain As Long, ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：为加快速度，一次性调用套餐项目的LED显示
'参数：行号范围，lngMain=主项行号,lngBegin-lngEnd:从项行号
    Dim dbl数量 As Double, dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    
    If mbytInFun = 0 And mbytInState = 0 And gblnLED And gblnLedDispDetail Then
        With mobjBill.Pages(mintPage)
            For j = 1 To .Details(lngMain).InComes.Count
                cur金额 = cur金额 + .Details(lngMain).InComes(j).实收金额
            Next
            For i = lngBegin To lngEnd
                For j = 1 To .Details(i).InComes.Count
                    cur金额 = cur金额 + .Details(i).InComes(j).实收金额
                Next
            Next
        End With
        With mobjBill.Pages(mintPage).Details(lngMain)
            If cur金额 <> 0 Then
                dbl数量 = IIf(.付数 = 0, 1, .付数) * .数次
                If dbl数量 <> 0 Then
                    dbl单价 = cur金额 / dbl数量
                Else
                    dbl单价 = cur金额
                End If
                If InStr(",5,6,7,", .Detail.类别) > 0 And gbln药房单位 Then
                    zl9LedVoice.Display .Detail.名称, .Detail.规格, .Detail.药房单位, dbl单价, dbl数量, cur金额
                Else
                    zl9LedVoice.Display .Detail.名称, .Detail.规格, .计算单位, dbl单价, dbl数量, cur金额
                End If
            End If
        End With
    End If
End Sub


Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytSubs As Byte
    Dim bln从项汇总折扣 As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Or chkCancel.Value = 1 Then Cancel = True: Exit Sub
    
    With mobjBill.Pages(mintPage)
        If .Details.Count >= Row Then
            If .Details(Row).工本费 Then
                MsgBox "该行不能修改及删除！", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        End If
        
        If .Details.Count >= Row Then
            '带从属项目的项删除确认
            For i = Row + 1 To .Details.Count
                If .Details(i).从属父号 = Row Then bytSubs = bytSubs + 1
            Next
            If bytSubs > 0 Then
                If MsgBox("该项目带有 " & bytSubs & " 个从属项目,删除该项目也将删除它的从属项目,继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: Exit Sub
                End If
            ElseIf .Details(Row).从属父号 <> 0 Then '从属项目删除确认
                If MsgBox("该项目是[" & .Details(.Details(Row).从属父号).Detail.名称 & "]的从属项目,确定要删除它吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: Exit Sub
                Else
                    bln从项汇总折扣 = gbln从项汇总折扣
                End If
            ElseIf MsgBox("确实要删除该收费项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
            
            If bln从项汇总折扣 Then lngMainRow = .Details(Bill.Row).从属父号 '如果是从项,删除之前记下从项的从属父号,如果是主项,则级联删除,不用重算
            
            '删除其从属行(反顺序)
            For i = .Details.Count To Row + 1 Step -1
                If .Details(i).从属父号 = Row Then
                    Call DeleteDetail(i)
                End If
            Next
 
            Call DeleteDetail(Row) '删除该行
            
            
            If bln从项汇总折扣 Then
                If CheckMainItem(lngMainRow) Or lngMainRow > 0 Then
                    Call CalcPItemActualIncome(lngMainRow)
                Else
                    Call CalcMoney(mintPage, lngMainRow, False)  '只有一个主项了,从项全部被删除时,当成普通独立项计算
                End If
            End If
                        
            '重新计算所有行并刷新
            Call ShowDetails
            Call ShowMoney(mintPage)
            
            '需要重新预结算
            If cmd预结算.Visible Then
                Call InitBalanceGrid
                cmd预结算.TabStop = True
                cmdOK.Enabled = False
            End If
            
'''            Call zlClear结算卡
            
            If CheckBillsEmpty Then ClearMoney
                                   
            Bill.TxtVisible = False
            Bill.CmdVisible = False
            Bill.CboVisible = False
            
            Cancel = True '不用控件来处理删除
            
            mlngPreRow = 0  '表示行改变了
            Call Bill_EnterCell(Bill.Row, Bill.Col)
        ElseIf Row = 1 Then
            For i = 1 To Bill.COLS - 1
                Bill.TextMatrix(Row, i) = ""
            Next
            Call SetBillRowForeColor(Row, Bill.ForeColor)
            Cancel = True
        End If
    End With
    
    Call SetColNum(Row)
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, strStock As String
    Dim blnComboxDown As Boolean
    Dim lng执行科室 As Long, str执行科室 As String
    '药品库存检查
    If ListIndex <> -1 And (Bill.TextMatrix(0, Bill.Col) = "执行科室" Or Bill.TextMatrix(0, Bill.Col) = "发药药店") Then
        If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
            With mobjBill.Pages(mintPage).Details(Bill.Row)
                blnComboxDown = SendMessage(Bill.cboHwnd, &H157, 0, 0) = 1
                If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                    lng执行科室 = .执行部门ID: str执行科室 = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .执行部门ID = Bill.ItemData(Bill.ListIndex)
                    .发药窗口 = ""
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                     
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        '取库存,如果是修改功能,此时重取库存后,需加上当前单据在该库房的库存,因比较麻烦,暂时不管
                        dblStock = GetStock(.收费细目ID, .执行部门ID)
                        If gbln药房单位 Then
                            dblStock = dblStock / .Detail.药房包装
                        End If
                        .Detail.库存 = dblStock  '记录当前行药品库存
                        Call ShowStock(.执行部门ID, .Detail.名称, .Detail.库存)
                        Call ShowStatusCargoSpace(.收费细目ID, .执行部门ID)    '显示货位
                        
                        '药房改变,时价药品重新计算价格
                        'If .Detail.变价 Then    '如果费别的计算方式是成本价加收法,则需要重算价格,这里简化不作判断
                            Call CalcMoneys(mintPage, Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney(mintPage)
                        'End If
                        '储备限额提示:
                        Call SetItemRowColor(mintPage, Bill.Row)
                        If blnComboxDown Then '显示出弹出菜单:问题:25238
                            DoEvents
                             SendMessageLong Bill.cboHwnd, &H14F, True, 0
                        End If
                    
                    ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                        .Detail.库存 = dblStock
                        Call ShowStock(.执行部门ID, .Detail.名称, .Detail.库存)
                        
                        '发料部门改变,时价卫材重新计算价格
                        If .Detail.变价 Then
                            Call CalcMoneys(mintPage, Bill.Row) '如果需要汇总计算,会重算主项实收
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney(mintPage)
                        End If
                        '储备限额提示:
                        Call SetItemRowColor(mintPage, Bill.Row)
                        If blnComboxDown Then '显示出弹出菜单:问题:25238
                            DoEvents
                             SendMessageLong Bill.cboHwnd, &H14F, True, 0
                        End If
                    '收费项目
                    ElseIf InStr(",4,5,6,7,", .收费类别) = 0 Then
                        If CheckMainItem(Bill.Row) Then Call SetSubItemDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                    End If
                    If Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
                        If mbytInFun = 0 And mintInsure <> 0 And MCPAR.实时监控 And mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                            If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = str执行科室: .执行部门ID = lng执行科室
                                Exit Sub
                            End If
                        End If
                        
                        If mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = str执行科室: .执行部门ID = lng执行科室
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To Bill.COLS - 1
        If Bill.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Public Function GetOriginalTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long, _
    Optional ByVal intPage As Integer) As Double
'功能：获取单据中指定药品在同一药房多行的原始数量和
'参数： lng药房ID-0表示分离发药时,不限定药房检查
    Dim i As Integer, p As Integer, dblCount As Double
    
    For p = 1 To objBill.Pages.Count
        If intPage = 0 Or p = intPage Then
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).收费细目ID = lng药品ID Then
                    If IIf(lng药房ID <> 0, objBill.Pages(p).Details(i).原始执行部门ID = lng药房ID, 1 = 1) Then
                        dblCount = dblCount + objBill.Pages(p).Details(i).原始数量
                    End If
                End If
            Next
        End If
    Next
    GetOriginalTotal = dblCount
End Function

Private Function GetDelMoney(Optional curError As Currency) As Currency
'功能：获取部份退费时当前已选择项的退款金额
'返回：curError=部份退费时产生的误差金额
'说明：退费时处理误差和分币；部份退费时才处理误差
    Dim cur单据合计 As Currency
    Dim cur选择合计 As Currency
    Dim cur退费合计 As Currency
    Dim bln完全退费 As Boolean, bln现金结算 As Boolean
    Dim intCOL_选择 As Integer, intCOL_金额 As Integer
    Dim i As Integer, strTempBalance As String, bln原样退 As Boolean
    Dim bln三方退现 As Boolean, bln三方部分退不退现 As Boolean
    
    curError = 0
    intCOL_选择 = GetColNum("退费")
    intCOL_金额 = GetColNum("实收金额")
    
    '单据合计金额
    For i = 1 To Bill.Rows - 1
        cur单据合计 = cur单据合计 + Val(Bill.TextMatrix(i, intCOL_金额))
        If Bill.TextMatrix(i, intCOL_选择) <> "" Then
            cur选择合计 = cur选择合计 + Val(Bill.TextMatrix(i, intCOL_金额))
        End If
    Next
    mTyDelFee.dblCurDelMoney = cur选择合计
    
    '完全退费时排开其它结算金额
    bln完全退费 = Not BillExistDelete(mstrInNO, 1) _
        And BillDeleteAll(mstrInNO, 1, mblnHaveExcuteData) And (cur单据合计 = cur选择合计)
    bln原样退 = bln完全退费
    bln三方退现 = False
    vsBalance.Tag = ""
    If bln完全退费 Then
        For i = 0 To vsBalance.Rows - 1
            '退费时可能显示了多种非医保结算,要排开,-1表示非医保结算
            If vsBalance.TextMatrix(i, 0) <> "" Then
                If IsNumeric(vsBalance.TextMatrix(i, 1)) And vsBalance.RowData(i) <> -1 Then
                    strTempBalance = vsBalance.TextMatrix(i, 0)
                    '如果这种结算方式不支持回退,要退为现金,则不用减去
                    If mblnYB结算作废 Then
                        If gclsInsure.GetCapability(support门诊结算作废, , mintInsure, strTempBalance) Then
                            cur选择合计 = cur选择合计 - Val(vsBalance.TextMatrix(i, 1))
                            vsBalance.Cell(flexcpBackColor, i, 0, i, 1) = &HE7CFBA
                        Else
                            bln原样退 = False
                        End If
                    Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                        If strTempBalance <> mstr个人帐户 Then
                            cur选择合计 = cur选择合计 - Val(vsBalance.TextMatrix(i, 1))
                            vsBalance.Cell(flexcpBackColor, i, 0, i, 1) = &HE7CFBA
                        End If
                        If InStr("3,4,5", vsBalance.Cell(flexcpData, i, 0)) = 0 Then
                            bln原样退 = False
                        Else
                            cur选择合计 = cur选择合计 - (vsBalance.Cell(flexcpData, i, 1) - Val(vsBalance.TextMatrix(i, 1)))
                            vsBalance.TextMatrix(i, 1) = FormatEx(vsBalance.Cell(flexcpData, i, 1), 2)
                        End If
                    End If
             
                Else
                    If InStr("3,4,5", vsBalance.Cell(flexcpData, i, 0)) > 0 Then
                        '缺省排除结算卡;医疗卡;一卡通的数据
                        'vsBalance.Cell(flexcpData, lngRow, 0)=性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
                        If Val(vsBalance.TextMatrix(i, 1)) <> 0 Then vsBalance.TextMatrix(i, 1) = FormatEx(vsBalance.Cell(flexcpData, i, 1), 2)
                        cur选择合计 = cur选择合计 - Val(vsBalance.TextMatrix(i, 1))
                        vsBalance.Cell(flexcpBackColor, i, 0, i, 1) = &HE7CFBA
                        If Not bln三方退现 Then
                            bln三方退现 = Val(vsBalance.TextMatrix(i, 1)) = 0 And Val(vsBalance.Cell(flexcpData, i, 1)) <> 0
                        End If
                    End If
                End If
            End If
        Next
        cur选择合计 = cur选择合计 - Val(txt预交冲款.Text)
    Else
        '部分退:55064
         vsBalance.Tag = "1"
        With vsBalance
            For i = 0 To .Rows - 1
                If mTyDelFee.blnSingleBalance And Val(.Cell(flexcpData, i, 0)) = 3 Then
                    If Val(.RowData(i)) = -1 And mTyDelFee.bln三方卡全退 = False Then
                        If Val(vsBalance.TextMatrix(i, 1)) <> 0 Then
                            .TextMatrix(i, 1) = FormatEx(cur选择合计, 2)
                        End If
                        cur选择合计 = cur选择合计 - Val(.TextMatrix(i, 1))
                    ElseIf Val(.RowData(i)) = -1 And mTyDelFee.bln三方卡全退 Then
                        .TextMatrix(i, 1) = ""
                    ElseIf Val(vsBalance.Cell(flexcpData, i, 0)) = 3 Then '3-医疗卡
                        .TextMatrix(i, 1) = FormatEx(IIf(vsBalance.Cell(flexcpData, i, 1) > cur选择合计, _
                                        cur选择合计, vsBalance.Cell(flexcpData, i, 1)), 2)
                        cur选择合计 = cur选择合计 - Val(.TextMatrix(i, 1))
                        bln三方部分退不退现 = True '不退现
                    End If
                Else
                    If Val(.RowData(i)) = -1 Then
                        .TextMatrix(i, 1) = ""
                    End If
                End If
            Next
        End With
    End If
    

    If Not bln完全退费 Then '68177
        '收费时全部用预交,退费时,不允许指定退费方式
        '部分退费时，如果预交金额正好与部分退的金额一致时，不能退为预交款:Val(txt预交冲款.Text) = GetBillSum And Val(txt预交冲款.Text) <> 0
        If mblnOlny预交 Then
             bln原样退 = True  '可能是部分退
            txt预交冲款.Visible = True
            lbl预交冲款.Visible = True
            lbl应缴.Caption = "退预交"
        Else
            txt预交冲款.Visible = False
            lbl预交冲款.Visible = False
            lbl应缴.Caption = "退款"
        End If
    Else
        '收费时全部用预交,退费时,不允许指定退费方式
        If Val(txt预交冲款.Text) = GetBillSum Then bln原样退 = True  '可能是部分退
        txt预交冲款.Visible = Val(txt预交冲款.Text) <> 0
        lbl预交冲款.Visible = Val(txt预交冲款.Text) <> 0
        lbl应缴.Caption = "退款"
    End If
    
    If bln原样退 And Not bln三方退现 Then
        zlControl.CboSetIndex cbo结算方式.hWnd, mintReturnMode
    End If
    cbo结算方式.Enabled = (Not bln原样退 Or bln三方退现) And mblnOlny预交 = False
    cbo结算方式.Locked = (bln原样退 And Not bln三方退现) And mblnOlny预交 = False
    fraAppend.Enabled = (Not bln原样退 Or bln三方退现) And mblnOlny预交 = False
    If mTyDelFee.blnSingleBalance And bln三方部分退不退现 Then
        cbo结算方式.Enabled = False
        cbo结算方式.Locked = True
        fraAppend.Enabled = False
    End If
    If cbo结算方式.Enabled And cbo结算方式.Locked = False Then
        If cbo结算方式.ListCount > 0 And cbo结算方式.ListIndex = -1 Then
            cbo结算方式.ListIndex = 0
        End If
    End If
    
    '现金结算时处理分币
    bln现金结算 = False
    If cbo结算方式.ListIndex <> -1 And mblnOlny预交 = False Then
        If cbo结算方式.ItemData(cbo结算方式.ListIndex) = 1 Then
            bln现金结算 = True
        End If
    End If
    
    '费用金额保留位数,及现金结算时处理分币
    If bln现金结算 Then
        If mintInsure > 0 Then
            If gclsInsure.GetCapability(support分币处理, , mintInsure) Then
                cur退费合计 = CentMoney(cur选择合计)
            Else
                cur退费合计 = Format(cur选择合计, "0.00")
            End If
        Else
            cur退费合计 = CentMoney(cur选择合计)
        End If
    Else
        cur退费合计 = Format(cur选择合计, "0.00")
    End If
    
    '误差金额,部分退,或医保全退时因为结算方式不支持回退而退为现金,可能产生误差
    '非现金结算时,也可能有误差,这个误差是费用金额保留位数引起的
    '60974
    curError = cur退费合计 - cur选择合计
    If curError <> 0 Then
        vsBalance.ToolTipText = "退费产生的误差金额:" & curError
    Else
        vsBalance.ToolTipText = "本次操作没有误差金额!"
    End If
'    If Not bln原样退 Then
'        curError = cur退费合计 - cur选择合计
'        vsBalance.ToolTipText = "退费产生的误差金额:" & curError
'    Else
'        vsBalance.ToolTipText = "本次操作没有误差金额!"
'    End If
'
    GetDelMoney = cur退费合计
End Function

Private Sub SelALLRow()
    '功能：实现退费时的全选
    Dim i As Long
    If InStr(",销帐,退费,", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
        For i = 1 To Bill.Rows - 1
            If Bill.TextMatrix(i, BillCol.项目) <> "" Then
                Bill.TextMatrix(i, Bill.COLS - 1) = "√"
            End If
        Next
    End If
    If mbytInFun = 0 And Bill.TextMatrix(0, Bill.COLS - 1) = "退费" Then
        Call ReCalce退款
    End If
End Sub

Private Sub ClearALLRow()
'功能：实现退费时的全清
    Dim i As Long
    
    If InStr(",销帐,退费,", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
        '有险类时必然是收费
        If mintInsure <> 0 Then sta.Panels(Pan.C2提示信息).Text = "医保病人的收费单据不允许部份退费!": Exit Sub
'        '刘兴洪:??
'        If mtySquareCard.bln卡结算 Then
'            sta.Panels(Pan.C2提示信息).Text = "刷结算卡病人的收费单据不允许部份退费!": Exit Sub
'        End If
'
        For i = 1 To Bill.Rows - 1
            Bill.TextMatrix(i, Bill.COLS - 1) = ""
        Next
    End If
    If mbytInFun = 0 And Bill.TextMatrix(0, Bill.COLS - 1) = "退费" Then
        Call ReCalce退款
    End If
End Sub
Private Sub zlSet诊疗固定关系(ByVal lngRow As Long, ByVal Col As Long, Optional lngNotCheckRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院费用记录
    '编制:刘兴洪
    '日期:2010-12-31 15:49:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, bln固定 As Boolean, i As Long, j As Long
    If Bill.TextMatrix(lngRow, BillCol.医嘱序号) = "" Then Exit Sub
    If mrs收费对照 Is Nothing Then Exit Sub
     '问题:33634:如果是固定的项目(诊疗收费关系):即医嘱产生的才判断
    varData = Split(Bill.TextMatrix(lngRow, BillCol.医嘱序号) & ",", ",")
    If Val(varData(0)) = 0 Then Exit Sub
    
    mrs收费对照.Filter = "医嘱序号=" & Val(varData(0)) & " And 收费细目ID=" & Val(varData(1))
    If Not mrs收费对照.EOF Then
        bln固定 = Val(Nvl(mrs收费对照!固有对照)) = 1
    Else
        bln固定 = False
    End If
    mrs收费对照.Filter = 0
    If bln固定 = False Then Exit Sub
    
    For i = 1 To Bill.Rows - 1
        If i <> lngRow And lngNotCheckRow <> i Then
            varTemp = Split(Bill.TextMatrix(i, BillCol.医嘱序号) & ",", ",")
            If varData(0) = varTemp(0) Then    '是相同的医嘱序号
                 mrs收费对照.Filter = "医嘱序号=" & Val(varTemp(0)) & " And 收费细目ID=" & Val(varTemp(1))
                If Not mrs收费对照.EOF Then
                    bln固定 = Val(Nvl(mrs收费对照!固有对照)) = 1
                Else
                    bln固定 = False
                End If
                If bln固定 Then
                    Bill.TextMatrix(i, Col) = Bill.TextMatrix(lngRow, Col)
                    '如果是主项,需要检查重项
                    If Val(Bill.TextMatrix(i, BillCol.从属父号)) = 0 Then '肯定为父项,因此,需要找从项内容
                        For j = i + 1 To Bill.Rows - 1
                             If Bill.RowData(i) = Val(Bill.TextMatrix(j, BillCol.从属父号)) Then
                                    Bill.TextMatrix(j, Col) = Bill.TextMatrix(i, Col)
                             End If
                        Next
                    End If
                End If
             End If
        End If
    Next
End Sub


Private Sub Bill_CellCheck(Row As Long, Col As Long)
'说明：可以全部为主要手术,但不能全部为附加手术
    Dim i As Long, strCheck As String, bytTime As Byte
    Dim blnReSet As Boolean '重新设置
    Dim bln固定 As Boolean, strErrMsg As String, varData As Variant ' (0-医嘱序号;1-收费细目ID)
    Dim varTemp As Variant
    Dim bln固定1 As Boolean
    Dim j As Long
    '退费,双击退费列
    
    If mbytInFun = 0 And Bill.TextMatrix(0, Col) = "退费" Then
        If mintInsure <> 0 Then sta.Panels(Pan.C2提示信息).Text = "医保病人的收费单据不允许部份退费!": Exit Sub
        '刘兴洪:??
        sta.Panels(2).Text = ""
        If Bill.TextMatrix(Row, Col) = "" Then
            With mTyDelFee.rsBlance
                 .Filter = 0
                 '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
                .Filter = "(是否全退=1 And 是否退现=0) Or (性质=4 And 是否全退=0 And 是否退现=0) " & _
                        "Or (性质=5 And 是否全退=0 And 是否退现=0)"
                If .RecordCount <> 0 Then .MoveFirst
                strErrMsg = ""
                If Not .EOF Then
                    strErrMsg = Nvl(!名称) & ":" & Format(Val(Nvl(!结算金额, 0)), "0.00")
                End If
                If strErrMsg <> "" Then
                    sta.Panels(Pan.C2提示信息).Text = "存在第三方交易(" & strErrMsg & ")，不允许部分退费": Exit Sub
                End If
                
                If Not mTyDelFee.blnSingleBalance Then
                    .Filter = "性质=3 And 是否全退=0 And 是否退现=0"
                    If .RecordCount <> 0 Then .MoveFirst
                    strErrMsg = ""
                    If Not .EOF Then
                        strErrMsg = Nvl(!名称) & ":" & Format(Val(Nvl(!结算金额, 0)), "0.00")
                    End If
                    If strErrMsg <> "" Then
                        sta.Panels(Pan.C2提示信息).Text = "存在第三方交易(" & strErrMsg & ")，不允许部分退费": Exit Sub
                    End If
                End If
            End With
        End If
        
        '问题:29201
        '级联更新
        If Val(Bill.TextMatrix(Row, BillCol.从属父号)) = 0 Then '肯定为父项,因此,需要找从项内容
            For i = Row + 1 To Bill.Rows - 1
                 If Bill.RowData(Row) = Val(Bill.TextMatrix(i, BillCol.从属父号)) Then
                        Bill.TextMatrix(i, Col) = Bill.TextMatrix(Row, Col)
                 End If
            Next
            Call zlSet诊疗固定关系(Row, Col)
        Else
            Call zlSet诊疗固定关系(Row, Col)
            '需要检查主项是否已经被
                For i = Row - 1 To 1 Step -1
                    If Bill.RowData(i) = Val(Bill.TextMatrix(Row, BillCol.从属父号)) Then
                        If Bill.TextMatrix(i, Col) <> "" Then
                            Bill.TextMatrix(i, Col) = Bill.TextMatrix(Row, Col)
                        End If
                        Call zlSet诊疗固定关系(i, Col, Row)
                         Exit For
                    End If
                Next
        End If
        Call ReCalce退款
        Call LoadInvoiceData(mTyDelFee.strNos)
        Call ShowInvoiceInfor
    End If
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '问题:29201
        '级联更新
        If mbytInFun <> 0 Then  '上面已经处理
            If Val(Bill.TextMatrix(Row, BillCol.从属父号)) = 0 Then '肯定为父项,因此,需要找从项内容
                    For i = Row + 1 To Bill.Rows - 1
                         If Bill.RowData(Row) = Val(Bill.TextMatrix(i, BillCol.从属父号)) Then
                                Bill.TextMatrix(i, Col) = Bill.TextMatrix(Row, Col)
                         End If
                    Next
            End If
        End If
        Exit Sub
    End If
    
    '新增的未处理行无效
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    If mobjBill.Pages(mintPage).Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = ""
        Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).收费类别 = "F" _
            And mobjBill.Pages(mintPage).Details(i).附加标志 = 0 And i <> Row Then
            bytTime = bytTime + 1
        End If
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '可能只存在附加手术后又改成了主手术,需要重新计处理
        blnReSet = (strCheck = "" And mobjBill.Pages(mintPage).Details(Row).收费类别 = "F" And mobjBill.Pages(mintPage).Details(Row).附加标志 = 1)
    End If
    If blnReSet Then
        With mobjBill.Pages(mintPage).Details(Row)
            
            .附加标志 = IIf(strCheck = "", 0, 1)
            Call CalcMoneys(mintPage, Row)
            Call ShowDetails(Row)
        End With
        
        Call ShowMoney(mintPage)
        
        '需要重新预结算
        If cmd预结算.Visible Then
            Call InitBalanceGrid
            cmd预结算.TabStop = True
            cmdOK.Enabled = False
        End If
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "单据中必然有一个手术不是附加手术！", vbInformation, gstrSysName
    End If
    
End Sub

Private Sub Bill_CommandClick()
    Dim lng项目id As Long, blnCancel As Boolean, bln护士 As Boolean
    Dim str类别 As String, str特准项目 As String
    Dim str排除类别 As String, lng批次 As Long
    
    Call GetOperatorInfo(mobjBill.Pages(mintPage).开单人, bln护士)
    If gbln收费类别 Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
        End If
    Else
        str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
    End If
    If mbytInFun = 0 And mstrYBPati <> "" Then
        '刘兴洪:24862
        If zl_Check特准项目(gclsInsure, mintInsure, mobjBill.病人ID, True) Then str特准项目 = Get保险特准项目(mobjBill.病人ID, "A.ID")
    End If
    If zlCheckBill存在非散装草药(mintPage) = True Then
        mblnSelect = False: Exit Sub
    End If
    lng批次 = -1
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, gint病人来源, mintInsure, gbln药房单位, str类别, _
        , , str特准项目, 0, , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级, mbln条码刷卡, lng批次)
    If lng项目id <> 0 Then
        Bill.Text = lng项目id
        Bill.Tag = lng批次
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

Private Sub ShowStock(ByVal lng库房ID As Long, str药品 As String, dbl库存 As Double)
    '功能：显示药品或卫材的库存
    '31936
    Call zlInit缺省部门
    If InStr(1, mstrPrivs, "显示库存") > 0 Then
        If InStr(1, gstr所属部门ID & ",", "," & lng库房ID & ",") > 0 Or gbyt库存显示方式 <= 0 Then   '31936
                sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & dbl库存
        Else
                sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & IIf(dbl库存 > 0, "有", "无") & "库存."
        End If
    Else
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]" & IIf(dbl库存 > 0, "有", "无") & "库存."
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'功能：处理单据输入
    Dim lng项目id As Long, str类别 As String, str特准项目 As String, bln护士 As Boolean
    Dim dblStock As Double, strScope As String
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency, cur余额 As Currency
    Dim blnInput As Boolean, str摘要 As String, lngOld付数 As Long
    Dim lngDoUnit As Long, lng病人科室ID As String, str药房IDs As String, i As Long, j As Long
    Dim colStock As Collection, str排除类别 As String
    Dim dblNum As Double, strPriceGrade As String, lng批次 As Long
    
    If KeyCode = 13 And Not Bill.Active Then
        Cancel = True: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
        
    On Error GoTo errH
    
    
    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.COLS - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "执行科室" And Bill.TextMatrix(0, Bill.Col) <> "发药药店" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
        
        '收费时,工本费不能修改
        If (mbytInFun = 0 Or mbytInFun = 1) And mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
            If mobjBill.Pages(mintPage).Details(Bill.Row).工本费 Then Exit Sub
        End If
        
 
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "类别"
                Call Clear连续累计
                If Bill.ListIndex <> -1 Then '不输入类别时不会定位到类别列
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        '一旦改更收费类别,则清除(如有)原有该项目内容
                        For i = 2 To Bill.COLS - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                            Set mobjBill.Pages(mintPage).Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Pages(mintPage).Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Pages(mintPage).Details(Bill.Row)
                                .收费细目ID = 0: .收费类别 = ""
                            End With
                            Call SetItemRowColor(mintPage, Bill.Row)
                            Call CalcMoneys(mintPage)
                            Call ShowMoney(mintPage)
                        End If
                    End If
                    Bill.TextMatrix(Bill.Row, BillCol.类别) = Bill.CboText
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '暂时用RowData记录所选择的收费类别
                End If
            Case "项目"
            
                '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
                If Bill.Text <> "" Then
                    '如果在已输入的项目上按回车,或选择器选择
                    If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                        '通过按钮选择是返回的ID,而输入则是文本,如果是一样的,则不改变
                        If Bill.TextMatrix(Bill.Row, BillCol.项目) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                    Call Clear连续累计
                    sta.Panels(Pan.C2提示信息) = ""
                    sta.Panels("MedicareType").Text = ""
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '立即清除该标志
                        Set mobjDetail = GetInputDetail(Val(Bill.Text), Val(Bill.Tag))
                    Else
                        If gbln收费类别 Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(Pan.C2提示信息) = "没有确定费用类别,请先输入类别！"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            Call GetOperatorInfo(mobjBill.Pages(mintPage).开单人, bln护士)
                            str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
                        End If
                        
                        If mbytInFun = 0 And mstrYBPati <> "" Then
                            '刘兴洪:24862
                            If zl_Check特准项目(gclsInsure, mintInsure, mobjBill.病人ID, True) Then str特准项目 = Get保险特准项目(mobjBill.病人ID, "A.ID")
                        End If
                        If zlCheckBill存在非散装草药(mintPage) Then
                            '存在非散装的,界面中就不能进行录入
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                        lng批次 = -1
                        lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, gint病人来源, mintInsure, gbln药房单位, _
                            str类别, Bill.Text, Bill.TxtHwnd, str特准项目, 0, str排除类别, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级, mbln条码刷卡, lng批次)
                        If lng项目id <> 0 Then
                            Set mobjDetail = GetInputDetail(lng项目id, lng批次)
                            If mintInsure <> 0 Then sta.Panels("MedicareType").Text = Get医保大类(lng项目id, mintInsure)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If

                    '确定了收费细目
                    Bill.TxtVisible = False '(不加不行)
                                            
                    '检查药品输入是否重复:分批及时价同一药房不允许重复(这里只提醒)
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 _
                        Or (mobjDetail.类别 = "4" And mobjDetail.跟踪在用) Then
                        If CheckDrugExist(mobjDetail) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                
                    '检查处方职务
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 And mbln处方职务检查 Then
                        mobjDetail.处方职务 = Get处方职务(mobjDetail.ID)
                        '医保或公费病人
                        If cbo医疗付款.ListIndex <> -1 Then
                            '医保或公费病人
                            '问题:45605
                            If zlIsCheckMedicinePayMode(zlStr.NeedName(cbo医疗付款)) Then
                                If CheckDuty(mobjDetail, False) > 0 Then
                                    Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        '所有病人
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '读取药品和卫材相关信息,卫材执行科室缺省为病人,如果本地指定了,则为指定科室
                    If mobjDetail.类别 = "4" Then
                        lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, mobjBill.科室ID)
                    Else
                        lngDoUnit = mobjBill.科室ID      '病人科室ID
                    End If
                    If lngDoUnit = 0 Then lngDoUnit = Get开单科室ID
                                         
                    '病人科室ID
                    lng病人科室ID = mobjBill.科室ID
                    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    
                    lngDoUnit = Get收费执行科室ID(mobjDetail.类别, mobjDetail.ID, _
                        mobjDetail.执行科室, lng病人科室ID, Get开单科室ID, gint病人来源, _
                        IIf(mlng西药房 = 0, glng西药房, mlng西药房), _
                        IIf(mlng成药房 = 0, glng成药房, mlng成药房), _
                        IIf(mlng中药房 = 0, glng中药房, mlng中药房), _
                        lngDoUnit, mobjBill.病区ID)
                                        
                    '读取药品及卫材库存
                    Call ReadDrugAndStuffStock(lngDoUnit, mobjDetail)
                    
                   
                    
                    '处方限量
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 And mbln处方限量检查 Then
                        mobjDetail.处方限量 = Get处方限量(mobjDetail.ID)
                    End If
                                        
                    '保险支付项目对应检查
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        strPriceGrade = mstr药品价格等级
                    ElseIf mobjDetail.类别 = "4" Then
                        strPriceGrade = mstr卫材价格等级
                    Else
                        strPriceGrade = mstr普通价格等级
                    End If
                    If mstrYBPati <> "" And Not MCPAR.允许不设置医保项目 Then
                        If Not CheckMediCareItem(mobjDetail.ID, mintInsure, mobjDetail.名称, mobjDetail.变价 = False, strPriceGrade) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '输入摘要(取已有的行以便修改)
                    If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                        If mobjBill.Pages(mintPage).Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            str摘要 = mobjBill.Pages(mintPage).Details(Bill.Row).摘要
                        End If
                    End If
                    '清除结算卡数据:
'''                    Call zlClear结算卡
                    
                    '加入或修改该收费细目行
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051
                    '输入摘要(根据新输入的行更改摘要)
                    If mobjBill.Pages(mintPage).Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Pages(mintPage).Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Pages(mintPage).Details(Bill.Row).摘要 = str摘要
                        End If
                    Else 'If mstrYBPati <> "" Then'90304
                         str摘要 = gclsInsure.GetItemInfo(mintInsure, mobjBill.病人ID, mobjBill.Pages(mintPage).Details(Bill.Row).收费细目ID, str摘要, 1)
                         mobjBill.Pages(mintPage).Details(Bill.Row).摘要 = str摘要
                    End If
                    
                    Call CalcMoney(mintPage, Bill.Row)      '此时还没有取从属项目
                    
                    'Calcmoney中医保可能返回摘要
                    If mobjBill.Pages(mintPage).Details(Bill.Row).摘要 <> "" Then str摘要 = mobjBill.Pages(mintPage).Details(Bill.Row).摘要
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mbytInFun = 2 Then
                        If mrsInfo.State = 1 And Not mrsWarn Is Nothing Then
                            curTotal = GetBillSum
                            If mobjBill.Pages(mintPage).Details.Count = Bill.Row And curTotal > 0 Then
                                cur余额 = Val(cmdPrint.Tag)
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(0, mrsInfo!病人ID) + IIf(mbytBilling = 1, Original.实收合计, 0)
                                gbytWarn = BillingWarn(mstrPrivs, mrsInfo!姓名, mrsInfo!适用病人, mrsWarn, cur余额, mrsInfo!当日额 - Original.实收合计, curTotal, _
                                            IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Pages(mintPage).Details(Bill.Row).收费类别, mobjBill.Pages(mintPage).Details(Bill.Row).Detail.类别名称, mstrWarn)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Pages(mintPage).Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                    Bill.Text = "": Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End If
                                        
                    If mbytInFun = 0 And mintInsure <> 0 And MCPAR.实时监控 And mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                        If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            mobjBill.Pages(mintPage).Details.Remove Bill.Row '删除刚刚想要加入的费用行
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                            mobjBill.Pages(mintPage).Details.Remove Bill.Row  '删除刚刚想要加入的费用行
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                          
                    '储备限额提示,不是药品也要执行,用于恢复单元格颜色
                    Call SetItemRowColor(mintPage, Bill.Row)
                          
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney(mintPage)
                    
                    '费用类型检查
                    Call CheckFeeType(Bill.Row)
                    
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If Bill.TextMatrix(Bill.Row, BillCol.付数) * Bill.TextMatrix(Bill.Row, BillCol.数次) * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Call DeleteDetail(Bill.Row, mintPage): Exit Sub
                            End If
                        End If
                    End If
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                    mlngPreRow = 0  '修改已有列时,恢复此值,以便显示库存
                    With mobjBill.Pages(mintPage).Details(Bill.Row)
                        '下一列的性质确定
                        If .收费类别 = "7" And gblnPay Then Bill.ColData(BillCol.付数) = BillColType.Text  '付数
                        If .收费类别 = "F" Then Bill.ColData(BillCol.标志) = BillColType.CheckBox  '附加标志
                        
                        '变价允许输入数次
                        If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                            And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                            Bill.ColData(BillCol.数次) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus)   '数次
                            Bill.ColData(BillCol.单价) = BillColType.Text '单价
                        Else
                            Bill.ColData(BillCol.数次) = BillColType.Text '数次
                            Bill.ColData(BillCol.单价) = BillColType.UnFocus '单价
                        End If
                        
                        '执行科室
                        '在FillBillComboBox中设置ListIndex时调用CboClick事件
                        mblnEnterCell = False: Bill.Col = BillCol.执行科室: mblnEnterCell = True
                        Call FillBillComboBox(Bill.Row, BillCol.执行科室, Not blnInput) '直接回车时保持执行科室
                        mblnEnterCell = False: Bill.Col = BillCol.项目: mblnEnterCell = True
                        
                        blnSkip = Bill.ListCount = 1
                        If Not blnSkip And InStr(",4,5,6,7,", .收费类别) > 0 Then
                            Select Case .收费类别 '指定了固定药房或发料部门时,不允许再选择
                                Case "4"
                                    blnSkip = glng发料部门 > 0 And .执行部门ID = glng发料部门
                                Case "5"
                                    blnSkip = glng西药房 > 0 And .执行部门ID = glng西药房
                                Case "6"
                                    blnSkip = glng成药房 > 0 And .执行部门ID = glng成药房
                                Case "7"
                                    blnSkip = glng中药房 > 0 And .执行部门ID = glng中药房
                            End Select
                        End If
                        If blnSkip Then
                            Bill.ColData(BillCol.执行科室) = BillColType.UnFocus: .Key = 1
                        Else
                            Bill.ColData(BillCol.执行科室) = BillColType.ComboBox: .Key = Bill.ListCount
                        End If
                        
                        If lngDoUnit <> .执行部门ID Then
                            '读取药品及卫材库存
                            Call ReadDrugAndStuffStock(.执行部门ID, .Detail)
                        End If
                        '检查卫生材料的灭菌效期,在确定执行科室之后
                        If .收费类别 = "4" And .Detail.跟踪在用 Then
                            Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                        End If
                        
                        '从属项目处理,仅该行收费项目有从属项目及尚未取才取,药品无需判断,药品不能设置主从项
                        If Bill.TextMatrix(0, Bill.Col) = "项目" And InStr(",5,6,7,", .收费类别) = 0 Then
                            If (gbln从项汇总折扣 And mobjBill.Pages(mintPage).Details(Bill.Row).从属父号 = 0) Or Not gbln从项汇总折扣 Then  '(如果有级联,只取一级)
                                If CheckHaveChildren(Bill.Row) Then
                                   Call SetSubItem
                                   mlngPreRow = 0 '通过行变化标志来重新确定列性质
                                End If
                            End If
                        End If
                    End With
                End If
'                With mobjBill.Pages(mintPage)
'                    If .Details.Count <> 0 And .Details.Count >= Bill.Row And Bill.Active And Visible Then
'                        If .Details(Bill.Row).收费类别 = "7" Then
'                             Call cmd配方_Click
'                             Exit Sub
'                        End If
'                    End If
'                End With
'
                '中药,默认只输入一次付数
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Pages(mintPage).Details(Bill.Row).收费类别 = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Pages(mintPage).Details(i).收费类别 = "7" Then
                                '正常执行该过程：本身会定位下一个单元,先定位到付数,则下一个单元是数次
                                '选择调用该过程：调用后会送个回车，这里不能再回车，否则是三个回车的效果(控件原因)。
                                Bill.Col = BillCol.付数: Exit For
                            End If
                        Next
                    End If
                End If
            Case "付数"
                With mobjBill.Pages(mintPage)
                    If .Details.Count >= Bill.Row And Bill.Text <> "" Then
                        '数字合法性
                        If Not IsNumeric(Bill.Text) Then
                            MsgBox "非法数值！", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).付数: Cancel = True: Exit Sub
                        End If
                        If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                            MsgBox "付数应该为正的整数！", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).付数: Cancel = True: Exit Sub
                        End If
    
                        '仅中草药才可更改付数(一项付数改变,其余也变,中药不能设置主从关系)
                        If mobjBill.Pages(mintPage).Details(Bill.Row).收费类别 = "7" Then
                            '分批或时价药品不足禁止输入(没有分批的时价药品可以修改付数、数次)
                            If .Details(Bill.Row).Detail.分批 Or .Details(Bill.Row).Detail.变价 Then
                                If CSng(Bill.Text) * .Details(Bill.Row).数次 > .Details(Bill.Row).Detail.库存 Then
                                    MsgBox """" & .Details(Bill.Row).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Bill.Text = .Details(Bill.Row).付数: Cancel = True: Exit Sub
                                End If
                            End If
                                  
                            '检查其它时价或分批中药更改付数后库存是否足够
                            For i = 1 To .Details.Count
                                If i <> Bill.Row And .Details(i).收费类别 = "7" And (.Details(i).Detail.变价 Or .Details(i).Detail.分批) Then
                                    If Val(Bill.Text) * .Details(i).数次 > .Details(i).Detail.库存 Then
                                        MsgBox "第 " & i & " 行药品""" & .Details(i).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                        Bill.Text = .Details(Bill.Row).付数: Cancel = True: Exit Sub
                                    End If
                                End If
                            Next
                            '最大金额检查
                            If gcurMaxMoney > 0 Then
                                If CSng(Bill.Text) * .Details(Bill.Row).数次 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                                    If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                        Bill.Text = .Details(Bill.Row).付数: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                            lngOld付数 = .Details(Bill.Row).付数
                            '清除结算卡数据:
'''                            Call zlClear结算卡
                            '计算并刷新该行
                            .Details(Bill.Row).付数 = Bill.Text
                            Call CalcMoneys(mintPage, Bill.Row)

                            '输了数量再改付数的，在这里重新检查，先输付数，再输数量的，在输数量后检查
                            If mbytInFun = 0 And mintInsure <> 0 And MCPAR.实时监控 And .Details(Bill.Row).数次 <> 0 Then
                                If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    .Details(Bill.Row).付数 = lngOld付数
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            If .Details(Bill.Row).数次 <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                                    .Details(Bill.Row).付数 = lngOld付数
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            Call ShowDetails(Bill.Row)
                            
                            '处理其它中药付数,如果是独立项,则修改其它非从项的,如果是从项,则修改同一主项的从项的.因为限定为中草药,不可能有主项
                            For i = 1 To .Details.Count
                                If i <> Bill.Row And .Details(i).收费类别 = "7" And .Details(i).从属父号 = .Details(Bill.Row).从属父号 Then
                                    If .Details(i).从属父号 = 0 Or (.Details(i).从属父号 <> 0 And .Details(i).Detail.固有从属 = 0) Then     '1和2固定和按比例的不改
                                        .Details(i).付数 = Bill.Text
                                        Call CalcMoneys(mintPage, i)
                                        Call ShowDetails(i)
                                    End If
                                End If
                            Next
                                                        
                            Call ShowMoney(mintPage)
                        Else
                            sta.Panels(Pan.C2提示信息) = "从属项目的付数不能更改！"
                            Bill.Text = .Details(Bill.Row).付数 '恢复原有付数值
                        End If
                    End If
                End With
            Case "数次"
                With mobjBill.Pages(mintPage)
                If .Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '快捷输入转换
                    If InStr(",7,", .Details(Bill.Row).收费类别) > 0 Then Bill.Text = ConvertABCtoNUM(Bill.Text)
                
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                        End If
                    End If
                    '药品输入小数
                    If InStr(",5,6,7,", .Details(Bill.Row).收费类别) > 0 Then
                        If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivs, "药品输入小数") = 0 Then
                            MsgBox "你没有权限输入小数！", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If InStr(",5,6,7,", .Details(Bill.Row).收费类别) > 0 And gbln药房单位 Then
                        dblNum = Val(Bill.Text) * .Details(Bill.Row).付数 * .Details(Bill.Row).Detail.药房包装
                    Else
                        dblNum = Val(Bill.Text) * .Details(Bill.Row).付数
                    End If
                                            
                    '负数合法性检查
                    If CSng(Bill.Text) * .Details(Bill.Row).付数 < 0 Then
                        '权限
                        If (mbytInFun < 2 And InStr(mstrPrivs, "负数费用") = 0 Or mbytInFun = 2 And InStr(mstrPrivs, "负数记帐") = 0) Then
                            MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                        ElseIf .Details(Bill.Row).Detail.分批 Then
                            MsgBox "分批药品不允许输入负数。", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                        End If
                        If mbytInFun = 2 Then
                            '负数冲销数量检查:只有门诊留观病人才会进行检查是否充足
                            '问题:36558
                             If Not mrsInfo Is Nothing Then
                                If mrsInfo.State = 1 Then
                                     If Nvl(mrsInfo!留观, 0) = 1 Then
                                        If Not CheckNegative(mobjBill.病人ID, mobjBill.主页ID, .Details(Bill.Row).收费细目ID, .Details(Bill.Row).执行部门ID, dblNum, .Details(Bill.Row).Detail.药房包装, mstrPrivs, Format(mrsInfo!入院日期, "yyyy-mm-dd HH:MM:SS")) Then
                                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * .Details(Bill.Row).付数 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                          
                    Bill.Text = FormatEx(Bill.Text, 5)
                          
                    '药品库存检查
                    With .Details(Bill.Row)
                        If (.收费类别 = "4" And .Detail.跟踪在用) Or InStr(",5,6,7,", .收费类别) > 0 Then
                            If .Detail.分批 Or .Detail.变价 Then
                                If .付数 * CSng(Bill.Text) > .Detail.库存 Then '分批或时价药品不足禁止输入
                                    If .收费类别 = "4" Then
                                        MsgBox """" & .Detail.名称 & """为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
                                
                                If colStock("_" & .执行部门ID) <> 0 And InStr(mstrPrivs, "不检查库存") = 0 And Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                                    If .付数 * CSng(Bill.Text) > .Detail.库存 Then '其它药品正常检查
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
                    End With
                    
                    dblPreTime = .Details(Bill.Row).数次
                    .Details(Bill.Row).数次 = Bill.Text
                    
                    '处方限量检查
                    If mbln处方限量检查 And Not gbln处方限量 Then
                        If Not CheckLimit(mobjBill, mintPage, Bill.Row) Then
                            .Details(Bill.Row).数次 = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    If .Details(Bill.Row).Detail.录入限量 > 0 And .Details(Bill.Row).数次 > .Details(Bill.Row).Detail.录入限量 Then
                        If MsgBox("输入的数次超过了录入限量" & .Details(Bill.Row).Detail.录入限量 & ",是否继续?", vbDefaultButton2 + vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                            .Details(Bill.Row).数次 = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '固有从属不能更改数次(主项目数次改变,固有从属的数次也变)
                    If .Details(Bill.Row).从属父号 <> 0 And .Details(Bill.Row).Detail.固有从属 <> 0 Then
                        sta.Panels(Pan.C2提示信息) = "该项目是固有从属项目,其数次不能够更改。"
                        .Details(Bill.Row).数次 = dblPreTime: Bill.Text = dblPreTime
                        Exit Sub
                    End If
                    '清除结算卡数据:
'''                    Call zlClear结算卡
                    
                    Call CalcMoneys(mintPage, Bill.Row)
                    
                    '数据溢出检查(在已经算出该行费用但未显示前)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                        .Details(Bill.Row).数次 = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(mintPage, Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    If mbytInFun = 0 And mintInsure <> 0 And MCPAR.实时监控 And .Details(Bill.Row).数次 <> 0 Then
                        If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            .Details(Bill.Row).数次 = dblPreTime
                            Call CalcMoneys(mintPage, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True::  Exit Sub
                        End If
                    End If
                    
                    If .Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                            .Details(Bill.Row).数次 = dblPreTime
                            Call CalcMoneys(mintPage, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mbytInFun = 2 Then
                        If mrsInfo.State = 1 And Not mrsWarn Is Nothing Then
                            curTotal = GetBillSum
                            If curTotal > 0 Then
                                cur余额 = Val(cmdPrint.Tag)
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(0, mrsInfo!病人ID) + IIf(mbytBilling = 1, Original.实收合计, 0)
                                gbytWarn = BillingWarn(mstrPrivs, mrsInfo!姓名, mrsInfo!适用病人, mrsWarn, cur余额, mrsInfo!当日额 - Original.实收合计, _
                                            curTotal, IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), .Details(Bill.Row).收费类别, .Details(Bill.Row).Detail.类别名称, mstrWarn)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    .Details(Bill.Row).数次 = dblPreTime
                                    Bill.Text = ""
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '更改其固有从属的数次(药品没有从属项目)
                    If .Details(Bill.Row).从属父号 = 0 Then
                        For i = Bill.Row + 1 To .Details.Count
                            If .Details(i).从属父号 = Bill.Row Then
                                '28136
                                '如果是输入的负数,需要将下级中的负数集中更新成负数
                                With .Details(i)
                                    If .Detail.固有从属 = 0 Then  '非固有从属
                                        If Abs(.数次) <> Abs(.Detail.从项数次) Then GoTo NotCalc:
                                        .数次 = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.从项数次
                                    ElseIf .Detail.固有从属 = 1 Then '固定的固有从属
                                        .数次 = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.从项数次 = 0, 1, .Detail.从项数次)
                                    ElseIf .Detail.固有从属 = 2 Then   '按比例的固有从属
                                        .数次 = Val(Bill.Text) * .Detail.从项数次
                                    Else
                                         GoTo NotCalc:
                                    End If
                                End With
                                Call CalcMoneys(mintPage, i)
                                Call ShowDetails(i)
NotCalc:
                            End If
                        Next
                    End If
                    
                    Call ShowMoney(mintPage)
                    
                ElseIf .Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
                If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                    If CheckMainItem(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
                End With
            Case "单价"
                With mobjBill.Pages(mintPage)
                If .Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "项目价格不应该为负数，要退费可以输入负的数量来实现！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * .Details(Bill.Row).付数 * .Details(Bill.Row).数次 > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '如果没有对应的收入项目,则无法计算
                    If .Details(Bill.Row).Detail.变价 And .Details(Bill.Row).InComes.Count > 0 Then
                        If Not (.Details(Bill.Row).InComes(1).现价 = 0 And .Details(Bill.Row).InComes(1).原价 = 0) Then
                            strScope = CheckScope(.Details(Bill.Row).InComes(1).原价, .Details(Bill.Row).InComes(1).现价, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(Pan.C2提示信息) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = .Details(Bill.Row).InComes(1).标准单价
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        '清除结算卡数据:
'''                        Call zlClear结算卡
                        
                        dblPreMoney = .Details(Bill.Row).InComes(1).标准单价
                                                
                        .Details(Bill.Row).InComes(1).标准单价 = Bill.Text '这种收费细目只能对应一个收入项目
                        Call CalcMoneys(mintPage, Bill.Row)
                        
                        '记帐分类报警(在已经算出该行费用但未显示前)
                        If mbytInFun = 2 Then
                            If mrsInfo.State = 1 And Not mrsWarn Is Nothing Then
                                curTotal = GetBillSum
                                If curTotal > 0 Then
                                    cur余额 = Val(cmdPrint.Tag)
                                    If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(0, mrsInfo!病人ID) + IIf(mbytBilling = 1, Original.实收合计, 0)
                                    gbytWarn = BillingWarn(mstrPrivs, mrsInfo!姓名, mrsInfo!适用病人, mrsWarn, cur余额, mrsInfo!当日额 - Original.实收合计, curTotal, _
                                                IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), .Details(Bill.Row).收费类别, .Details(Bill.Row).Detail.类别名称, mstrWarn)
                                    If gbytWarn = 2 Or gbytWarn = 3 Then
                                        .Details(Bill.Row).InComes(1).标准单价 = dblPreMoney
                                        Bill.Text = ""
                                        Call CalcMoneys(mintPage, Bill.Row)
                                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        Call ShowMoney(mintPage)
                    Else
                        Bill.Text = "0"
                        sta.Panels(Pan.C2提示信息) = "该项目设有设置对应的费目，所以无法计算费用！"
                    End If
                End If
                End With
            Case "执行科室", "发药药店"
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Pages(mintPage).Details(Bill.Row)
                        If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then    'cbo_click中有可能会执行一次
                             .执行部门ID = Bill.ItemData(Bill.ListIndex)
                            If CheckMainItem(Bill.Row) Then Call SetSubItemDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
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
                                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
                                
                                If colStock("_" & .执行部门ID) <> 0 And InStr(mstrPrivs, "不检查库存") = 0 Then
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
                        
                        If CheckMainItem(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                        If Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
                            If mbytInFun = 0 And mintInsure <> 0 And MCPAR.实时监控 And mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                                If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            If mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End With
                End If
        End Select
        
        '需要重新预结算
        If InStr(",类别,项目,付数,数次,单价,", "," & Bill.TextMatrix(0, Bill.Col) & ",") > 0 Then
            If cmd预结算.Visible Then
                Call InitBalanceGrid
                cmd预结算.TabStop = True
                cmdOK.Enabled = False
            End If
          '  Call zlClear结算卡
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).从属父号 = lngRow Then
            If mobjBill.Pages(mintPage).Details(i).Detail.固有从属 = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Pages(mintPage).Details.Count Then
        Bill.Col = BillCol.数次
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Pages(mintPage).Details.Count >= Bill.Rows - 1 Then
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
Dim lngDoUnit As Long, lng病人科室ID As Long
Dim bln从项汇总折扣 As Boolean
Dim str摘要 As String, strPriceGrade As String

lngMainRow = Bill.Row               '主项的行
If gbln从项汇总折扣 Then            '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
    bln从项汇总折扣 = Not mobjBill.Pages(mintPage).Details(lngMainRow).Detail.屏蔽费别
End If

lng病人科室ID = mobjBill.科室ID
If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)


With mobjBill.Pages(mintPage).Details(lngMainRow)
    Set mcolDetails = GetSubDetails(.收费细目ID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Pages(mintPage).Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)    '增加新行
             mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = ""  '有必要加上
        
        'a.从属项目为非药品项目的执行科室
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).类别) = 0 Then
             If mcolDetails(i).类别 = .收费类别 Or mcolDetails(i).执行科室 = 0 Then
                '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                lngDoUnit = .执行部门ID
             Else
                '其它非药项目的执行科室
                lngDoUnit = Get收费执行科室ID(mcolDetails(i).类别, mcolDetails(i).ID, _
                    mcolDetails(i).执行科室, lng病人科室ID, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
             End If
        'b.从属项目为药品,卫材的执行科室
        Else
            lngDoUnit = Get收费执行科室ID(mcolDetails(i).类别, mcolDetails(i).ID, mcolDetails(i).执行科室, lng病人科室ID, Get开单科室ID, gint病人来源, _
                IIf(mlng西药房 = 0, glng西药房, mlng西药房), IIf(mlng成药房 = 0, glng成药房, mlng成药房), _
                IIf(mlng中药房 = 0, glng中药房, mlng中药房), .执行部门ID, mobjBill.病区ID)  '卫材从项缺省与主项执行科室相同
        End If
        '保险支付项目对应检查
        If InStr(",5,6,7,", mcolDetails(i).类别) > 0 Then
            strPriceGrade = mstr药品价格等级
        ElseIf mcolDetails(i).类别 = "4" Then
            strPriceGrade = mstr卫材价格等级
        Else
            strPriceGrade = mstr普通价格等级
        End If
        If mstrYBPati <> "" And Not MCPAR.允许不设置医保项目 Then
            If Not CheckMediCareItem(mcolDetails(i).ID, mintInsure, mcolDetails(i).名称, mcolDetails(i).变价 = False, strPriceGrade) Then
                Exit Sub
            End If
        End If
        
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
                
        Call CalcMoney(mintPage, Bill.Rows - 1, bln从项汇总折扣)
        Call ShowDetails(Bill.Rows - 1, i, mcolDetails.Count)
                
'        If mstrYBPati <> "" Then'90304
            'CalcMoney中先调用GetuItemInsure可能返回摘要
            str摘要 = mobjBill.Pages(mintPage).Details(Bill.Rows - 1).摘要
             
            str摘要 = gclsInsure.GetItemInfo(mintInsure, mobjBill.病人ID, mcolDetails(i).ID, str摘要, 1)
            mobjBill.Pages(mintPage).Details(Bill.Rows - 1).摘要 = str摘要
'        End If
    Next
        
    If bln从项汇总折扣 Then
        Call CalcMoney(mintPage, lngMainRow, bln从项汇总折扣) '先重算主项的应收与实收,因为在没有加入从项前不能确定算不算
        
        Call CalcPItemActualIncome(lngMainRow)
    End If
    
    Call ShowMoney(mintPage)
    
    '一次性调用套餐项目LED显示
    Call ShowGroupLED(Bill.Row, Bill.Rows - mcolDetails.Count, Bill.Rows - 1)
    
End With

End Sub

Private Sub CalcPItemActualIncome(ByVal lngMainRow As Long, Optional intPage As Integer)
'功能:当从项汇总折扣时,根据指定的主项的行ID的第一个收入项目重算主项的实收金额
'参数:  lngMainRow-主项行ID
'       intpage -页号,默认为当前页mintpage

Dim i As Long, j As Long
Dim cur打折前应收合计 As Currency     '记录所有主从项的应收合计
Dim cur打折后实收 As Currency
Dim str费别 As String               '记录根据应收等确定的最优惠的费别
If intPage = 0 Then intPage = mintPage

With mobjBill.Pages(intPage)
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).从属父号 <> lngMainRow Then Exit For    '虽然目前限制了不允许在从项中间插入别的主从项,但因一张单据行数不多,为了将来可能的需求,还是全部扫描
        
        If i = lngMainRow Or .Details(i).从属父号 = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur打折前应收合计 = cur打折前应收合计 + .Details(i).InComes(j).应收金额
            Next
        End If
    Next
    '药品不支持主从项，所以无需传加班加价率等
    '打折后的实收金额仅算到主项的第一个收入项目上
    str费别 = IIf(glngSys Like "8??" Or mbytInFun = 2, mobjBill.费别, zlStr.TrimEx(mobjBill.费别 & "," & lbl动态费别.Tag, ","))
    
    cur打折后实收 = CCur(Format(ActualMoney(str费别, .Details(lngMainRow).InComes(1).收入项目ID, cur打折前应收合计, 0, 0, 0, 0), gstrDec))
    cur打折后实收 = cur打折后实收 - cur打折前应收合计 + .Details(lngMainRow).InComes(1).应收金额
    
    .Details(lngMainRow).InComes(1).实收金额 = Format(cur打折后实收, gstrDec)
    .Details(lngMainRow).费别 = str费别
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'功能:根据主项执行科室的变化,刷新非药从项的执行科室

    Dim i As Long, j As Long, lng病人科室ID As Long
    
    With mobjBill.Pages(mintPage)
        '获取所有从项及其执行科室类型,必须现取(因为界面上的从项信息可能是修改过的)
        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID)
        
        lng病人科室ID = mobjBill.科室ID
        If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        
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
                                .Details(i).执行部门ID = Get收费执行科室ID(mcolDetails(j).类别, mcolDetails(j).ID, _
                                    mcolDetails(j).执行科室, lng病人科室ID, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
                            End If
                        End If
                    End If
                    
                    '刷新显示从项执行科室
                    If .Details(i).执行部门ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
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
            End If
        Next
    End With
End Sub

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim i As Integer, bln工本费 As Boolean
    
    Dim strStock As String, strTmp As String
    Dim str药房IDs As String
    
    If Not mblnEnterCell Then Exit Sub
    
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '针对列编辑性质设置颜色
        Bill.SetColColor BillCol.类别, &HE7CFBA  '不然要成白色
        Exit Sub
    End If
    
    If Not Bill.Active Then
        '显示划价单摘要:医嘱内容
        If mbytInFun = 0 And mbytInState = 0 Then
            If mobjBill.Pages(mintPage).NO <> "" And Bill.RowData(Bill.Row) <> 0 Then
                strTmp = Get费用摘要(mobjBill.Pages(mintPage).NO, 1, Bill.RowData(Bill.Row))
                If strTmp <> "" Then sta.Panels(Pan.C2提示信息) = "摘要:" & strTmp
            End If
        End If
        Exit Sub
    End If
    If zlCheckBill存在非散装草药(mintPage) = True Then
        '如果单据中存在非散装的,则不能输入
        Call SetBill中草药EditEnabled
         Exit Sub
    End If
    
     '--------------------------------------------------------------------------
    '1.行改变的相关数据处理和设置
    If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And mlngPreRow <> Row Then
        '收费时,如果为工本费,则不能修改
        If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then
            If mobjBill.Pages(mintPage).Details(Row).工本费 Then
                bln工本费 = True
                For i = 0 To UBound(marrColData)
                    Bill.ColData(i) = IIf(marrColData(i) = BillColType.UnFocus, BillColType.UnFocus, BillColType.Text_UnModify)
                Next
            End If
        End If
        
        If Not bln工本费 Then
            '显示库存
            With mobjBill.Pages(mintPage).Details(Bill.Row)
                If InStr(",5,6,7,", .收费类别) > 0 And .收费细目ID <> 0 Then
                    If gbln其它药房 Or gbln其它药库 Then
                        strStock = GetStockInfo(.收费细目ID, gbln其它药房, gbln其它药库)
                        If strStock <> "" Then
                            If InStr(1, mstrPrivs, "显示库存") > 0 Then
                                sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行库存:" & strStock
                            Else
                                sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行有库存."
                            End If
                           ' Call ShowStatusCargoSpace(.收费细目ID, .执行部门ID)     '显示货位
                        End If
                        
                    End If
                    If strStock = "" Then
                        '更新库存显示
                        If Not (mbytInState = 0 And mstrInNO <> "") Then
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        End If
                        Call ShowStock(.执行部门ID, .Detail.名称, .Detail.库存)
                        Call ShowStatusCargoSpace(.收费细目ID, .执行部门ID)     '显示货位
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                    If Not (mbytInState = 0 And mstrInNO <> "") Then
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                    End If
                    Call ShowStock(.执行部门ID, .Detail.名称, .Detail.库存)
                Else
                    sta.Panels(Pan.C2提示信息) = ""
                End If
                   
                Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
                Bill.ColData(BillCol.项目) = BillColType.CommandButton
                 '如果是从属项目的主项目或从项,则不允许更改类别和项目
                If CheckMainItem(Row) Or mobjBill.Pages(mintPage).Details(Row).从属父号 > 0 Then
                    Bill.ColData(BillCol.类别) = BillColType.Text_UnModify
                    Bill.ColData(BillCol.项目) = BillColType.Text_UnModify
                End If
            
                '如果是非调整状态
                If mbytInState <> 2 Then
                    If .收费类别 = "7" And gblnPay Then
                        Bill.ColData(BillCol.付数) = BillColType.Text
                    Else
                        Bill.ColData(BillCol.付数) = BillColType.UnFocus
                    End If
                    
                    '变价允许输入数次
                    If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                        And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                        Bill.ColData(BillCol.数次) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus)   '数次
                        Bill.ColData(BillCol.单价) = BillColType.Text  '金额
                    Else
                        Bill.ColData(BillCol.数次) = BillColType.Text
                        Bill.ColData(BillCol.单价) = BillColType.UnFocus
                    End If
                    
                    If .Key = "1" Then    '指定了固定药房时,不允许再选择执行科室
                        Bill.ColData(BillCol.执行科室) = BillColType.UnFocus
                    Else
                        Bill.ColData(BillCol.执行科室) = BillColType.ComboBox
                    End If
                    
                    If .收费类别 = "F" Then
                        Bill.ColData(BillCol.标志) = BillColType.CheckBox
                    Else
                        Bill.ColData(BillCol.标志) = BillColType.UnFocus
                    End If
                    
                    '只允许一个类别,不允许选择类别
                    If mblnOne Then Bill.ColData(BillCol.类别) = BillColType.UnFocus
                End If
                
                '显示输入的摘要
                If .摘要 <> "" Then
                    sta.Panels(Pan.C2提示信息) = sta.Panels(Pan.C2提示信息) & "  摘要:" & .摘要
                End If
            End With
        End If
    End If
    
    '如果点击未保存的行,则恢复列的性质
    If mobjBill.Pages(mintPage).Details.Count < Bill.Row Then
        Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus) '类别列,当主从项时会被改变
        Bill.ColData(BillCol.项目) = BillColType.CommandButton  '项目列,当主从项时会被改变
    End If
    
    
    '-----------------------------------------------------------------
    '2.列改变的相关数据处理和显示设置
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then   '加载当前列的下拉项数据
        Call FillBillComboBox(Bill.Row, Bill.Col, True)
    End If
    
    If gbln收费类别 And Bill.TextMatrix(Row, BillCol.类别) = "" And mblnOne Then
        mrsClass.Filter = "编码=" & gstr收费类别
        Bill.TextMatrix(Row, BillCol.类别) = mrsClass!类别
        Bill.RowData(Row) = Asc(mrsClass!编码)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "类别" '不输入类别时不会定位到类别列
            SetWidth Bill.cboHwnd, 70
            '类别如果为空,则自动默认为上一收费细目的类别
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "编码=" & gstr收费类别
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
                    Bill.ListIndex = SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "执行科室", "发药药店"
            SetWidth Bill.cboHwnd, 130
        Case "付数"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "数次"
            Bill.TextLen = 8
            If (mbytInFun < 2 And InStr(mstrPrivs, "负数费用") = 0 Or mbytInFun = 2 And InStr(mstrPrivs, "负数记帐") = 0) Then
                Bill.TextMask = "0123456789." & Chr(8)
            Else
                Bill.TextMask = "-0123456789." & Chr(8)
            End If
            
            If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And InStr(Bill.TextMask, "-") > 0 Then
                If mobjBill.Pages(mintPage).Details(Bill.Row).Detail.分批 Then
                    Bill.TextMask = Replace(Bill.TextMask, "-", "")
                End If
            End If
            
            If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Pages(mintPage).Details(Bill.Row).收费类别) > 0 Then
                    If InStr(mstrPrivs, "药品输入小数") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                    '中药快捷输入
                    If mobjBill.Pages(mintPage).Details(Bill.Row).收费类别 = "7" Then
                        Bill.TextMask = Bill.TextMask & gstrABC & LCase(gstrABC)
                    End If
                End If
            End If
        Case "单价"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
            
    '新行,或更改已有行的类别时,看作换行还没有开始
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Pages(mintPage).Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_LostFocus()
    Bill.TxtVisible = False
    Bill.CmdVisible = False
    Bill.CboVisible = False
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.婴儿费 = cboBaby.ListIndex
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboSex_Click()
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then
        mobjBill.性别 = zlStr.NeedName(cboSex.Text)
    End If
End Sub

Private Sub cbo费别_Click()
    If mbytInState = 0 Then
        If cbo费别.ListIndex = -1 Then
            mobjBill.费别 = ""
        Else
            '即使费用相同也要重算,因为医保验卡后必须重算,预结算才正确
            If (mstrYBPati <> "" Or mobjBill.费别 <> zlStr.NeedName(cbo费别.Text)) And Not mbln不重算价格 Then
                mobjBill.费别 = zlStr.NeedName(cbo费别.Text)
                If mbytInState = 0 And Not CheckBillsEmpty Then
                    '需要重新预结算
                    If cmd预结算.Visible Then
                        Call InitBalanceGrid
                        cmd预结算.TabStop = True
                        cmdOK.Enabled = False
                    End If
'''                    Call zlClear结算卡
                    
                    '全部重新计算价格
                    Call CalcMoneys
                    Call ShowDetails
                    Call ShowMoney
                End If
            End If
        End If
    End If
End Sub

Private Sub cbo结算方式_Click()
'功能：在现金与非现金之间切换时，需要根据情况决定是否处理分币
    If cbo结算方式.ListIndex = -1 Then Exit Sub
    If mblnNotClick Then Exit Sub
    If Not (Visible And mbytInFun = 0 And gBytMoney <> 0) Then Exit Sub
    
    If Bill.TextMatrix(0, Bill.COLS - 1) = "退费" Then
        Call ReCalce退款
    Else
        Call ShowMoney(-1) '单据内容未变,全部不用重新计算
    End If
    
'''    Call zlCheck支票结算
End Sub

Private Sub cbo结算方式_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.COLS - 1
    End If
End Sub

Private Sub cbo结算方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii >= 32 Then
        If cbo结算方式.Locked Then Exit Sub
        
        lngIdx = zlControl.CboMatchIndex(cbo结算方式.hWnd, KeyAscii)
        If lngIdx = -1 And cbo结算方式.ListCount > 0 Then lngIdx = 0
        cbo结算方式.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, lng开单部门ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
        
    If cbo开单科室.ListIndex <> -1 Then lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If mobjBill.Pages(mintPage).开单部门ID = lng开单部门ID Then Exit Sub
    mobjBill.Pages(mintPage).开单部门ID = lng开单部门ID
        
    '定位医生
    If gbyt科室医生 = 1 Then
        If cbo开单科室.ListIndex <> -1 Then
            Call FillDoctor(lng开单部门ID)
            
            If cbo开单人.ListCount > 0 And (Not (gbln不缺省开单人 And mbytInFun <> 2)) Then
                Call zlControl.CboSetIndex(cbo开单人.hWnd, 0)
            End If
        Else
            cbo开单人.Clear
        End If
        Call cbo开单人_Click
    End If
    
    
    '根据开单科室重新设置收费项目的执行科室
    If cbo开单科室.ListIndex <> -1 And Visible Then
        With mobjBill.Pages(mintPage)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.类别) = 0 And _
             (.Details(i).Detail.执行科室 = 6 And gbyt科室医生 <> 2 Or InStr(",1,2,", "," & .Details(i).Detail.执行科室 & ",") > 0 And gint病人来源 = 1) Then '6-开单人科室
                
                .Details(i).执行部门ID = lng开单部门ID
                
                If i <= Bill.Rows - 1 And .Details(i).执行部门ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.执行科室) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
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
            ElseIf .Details(i).Detail.类别 = "4" Then
                Call ReSet卫材执行科室(i) '113644
            End If
        Next
        End With
    End If
    
    
    '婴儿费的处理,门诊记帐,新增或修改时
    If mbytInFun = 2 And mbytBilling <> 2 Then
        If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0 '触发click事件
        cboBaby.Enabled = False
        If cbo开单科室.ListIndex <> -1 Then cboBaby.Enabled = is产科(cbo开单科室.ItemData(cbo开单科室.ListIndex), mrs开单科室)
    End If
    
    '费别处理
    Call LoadAndSeek费别
    
End Sub

Private Sub ReSet卫材执行科室(ByVal lngRow As Long)
    '开单科室改变后，重新设置卫生材料的执行科室
    '问题号：113644
    '说明：卫材执行科室缺省顺序:
    '    一、门诊病人:
    '    1:指定发料部门(参数“缺省发料部门”)
    '    2: 开单科室
    '    3: 第一个可用的执行科室
    '    二、住院病人:
    '    1:指定发料部门(参数“缺省发料部门”),不管是否服务于病人科室
    '    2: 其它可服务于病人科室的执行科室
    '     2.1:病人科室
    '     2.3:病人病区
    '     2.4:可服务于病人科室的第一个执行科室
    '    3: 第一个可用的执行科室
    Dim lngDoUnit As Long, lng病人科室ID As Long
    
    On Error GoTo errHandler
    If Not (mbytInFun = 1 And mbytInState = 0) Then Exit Sub
    With mobjBill.Pages(mintPage)
        If .Details(lngRow).Detail.类别 <> "4" Then Exit Sub
        
        '卫材执行科室缺省为病人科室,如果本地指定了,则为指定科室
        lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, mobjBill.科室ID)
        If lngDoUnit = 0 Then lngDoUnit = Get开单科室ID
                             
        '病人科室ID
        lng病人科室ID = mobjBill.科室ID
        If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        
        lngDoUnit = Get收费执行科室ID(.Details(lngRow).Detail.类别, .Details(lngRow).Detail.ID, _
            .Details(lngRow).Detail.执行科室, lng病人科室ID, Get开单科室ID, gint病人来源, _
            IIf(mlng西药房 = 0, glng西药房, mlng西药房), _
            IIf(mlng成药房 = 0, glng成药房, mlng成药房), _
            IIf(mlng中药房 = 0, glng中药房, mlng中药房), _
            lngDoUnit, mobjBill.病区ID)
        
        .Details(lngRow).执行部门ID = lngDoUnit
        
        If lngRow <= Bill.Rows - 1 And .Details(lngRow).执行部门ID <> 0 Then
            mrsUnit.Filter = "ID=" & .Details(lngRow).执行部门ID
            If mrsUnit.RecordCount <> 0 Then
                Bill.TextMatrix(lngRow, BillCol.执行科室) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
            Else
                Bill.TextMatrix(lngRow, BillCol.执行科室) = GET部门名称(.Details(lngRow).执行部门ID, mrsUnit)
            End If
        Else
            Bill.TextMatrix(lngRow, BillCol.执行科室) = ""
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadAndSeek费别(Optional blnNew As Boolean)
'功能:加载普通费别与动态费别,定位缺省费别或病人费别
'参数:blnNew 是否新单据初始
'说明:门诊记帐不使用动态费别
    Dim lngDeptID As Long, blnDo As Boolean, strInfo As String
    
    If glngSys Like "8??" Or mbytInFun = 2 Then Exit Sub
    
    If cbo开单科室.ListIndex <> -1 Then lngDeptID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Call Load费别(cbo费别, lngDeptID, True, mrs费别)
                
    '显示可用动态费别：当前不是划价单时,窗体默认为可见
    If Bill.Active Or blnNew Then
        lbl动态费别.Caption = Load动态费别(lngDeptID)
        lbl动态费别.Tag = lbl动态费别.Caption
        lbl动态费别.Visible = lbl动态费别.Caption <> ""
        If lbl动态费别.Caption <> "" Then lbl动态费别.Caption = "(" & lbl动态费别.Caption & ")"
    End If
    
    cbo费别.Locked = (Not Bill.Active) Or (mbytInFun = 0 And mrsInfo.State = 1 And InStr(1, mstrPrivs, "调整病人费别") = 0) Or mbytInFun = 2: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
    If mrsInfo.State = 0 Then
         '未建档案的病人可以自由选择
         cbo费别.Locked = False: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
        If cbo费别.ListIndex = -1 And cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
    Else
        '定位有档案病人的费别
        cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(mrsInfo!费别), True)
        If cbo费别.ListIndex <> -1 Then
            '费别为初诊但病人不是初诊
            If cbo费别.ItemData(cbo费别.ListIndex) = 2 And mrsInfo!初诊 = 0 Then
                blnDo = True
                strInfo = "病人费别""" & mrsInfo!费别 & """仅限初诊时使用,但该病人不是第一次就诊"
            End If
        Else
            blnDo = True
            strInfo = "病人费别" & mrsInfo!费别 & "不可用，可能已失效"
        End If
        
        If blnDo Then
            Call Load费别(cbo费别, lngDeptID, False, mrs费别)
            If cbo费别.ListIndex <> -1 Then
                If Visible And Not mblnDoing Then MsgBox strInfo & ",将使用缺省费别！", vbInformation, gstrSysName
            Else
                cbo费别.Locked = False: cbo费别.TabStop = Not cbo费别.Locked And gbln费别 '无法确定,让其自由选择
                If cbo费别.Visible And Not mblnDoing Then
                    MsgBox strInfo & ",请选择一种费别！", vbInformation, gstrSysName
                    If cbo费别.Enabled Then cbo费别.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
 '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件

    If Not mblnCboClick Then cbo开单科室_Click
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then cbo开单科室.Text = ""
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
    If mobjBill.Pages(mintPage).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text)) Then Exit Sub
    
    mobjBill.Pages(mintPage).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    If cbo开单人.ListIndex <> -1 Then
        lng开单人ID = cbo开单人.ItemData(cbo开单人.ListIndex)
        mrs开单人.Filter = "ID=" & lng开单人ID
        If mrs开单人.RecordCount > 0 Then
            lblDuty.Caption = IIf(IsNull(mrs开单人!专业技术职务), "", mobjBill.Pages(mintPage).开单人 & "专业职务:" & mrs开单人!专业技术职务)
        Else
            lblDuty.Caption = ""
        End If
    Else
        lblDuty.Caption = ""
    End If
    
    
    '由医生确定科室
    If gbyt科室医生 = 0 Then
        If cbo开单人.ListIndex <> -1 Then
            Call FillDept(mlngDeptID, lng开单人ID)
            Call SetDefaultDept(lng开单人ID)
        Else
            cbo开单科室.Clear
        End If
        Call cbo开单科室_Click
    End If
    
    '科室医生独立,因为开单人变了，所以,执行科室是由开单人科室决定时，需要重设执行科室
     '不独立时在Cbo开单科室_click中处理
    If cbo开单人.ListIndex <> -1 And Visible And gbyt科室医生 = 2 Then
        With mobjBill.Pages(mintPage)
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
                            Bill.TextMatrix(i, BillCol.执行科室) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
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
            ElseIf .Details(i).Detail.类别 = "4" Then
                Call ReSet卫材执行科室(i) '113644
            End If
        Next
        End With
    End If
    
    '护士类别
    If Bill.Active And Visible Then
        If mobjBill.Pages(mintPage).Details.Count < Bill.Rows - 1 _
            And Bill.Row = Bill.Rows - 1 And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '清除无效输入
            Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.类别 Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) '刷新
        End If
    End If
    
    '护士类别:判断非法输入
    If Not mblnDoing Then
        If CheckInhibitiveByNurse(mintPage) Then
            MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
        End If
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

Private Sub cbo年龄单位_Validate(Cancel As Boolean)
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
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
                '问题:42886
                If txtDate.Enabled And txtDate.Visible Then
                    txtDate.SetFocus
                ElseIf cmdOK.Enabled And cmdOK.Visible Then
                    cmdOK.SetFocus
                End If
            Else
                If cbo开单科室.Enabled Then cbo开单科室.SetFocus
            End If
        End If
    End If
End Sub


Private Sub chkCancel_Click()
    Dim i As Integer
    
    mstrInNO = "": txtModi.Text = ""
    mlngFirstID = 0: mstrFirstWin = ""
    Call ClearPayInfo
        
    Call ClearPatientInfo(True)
    Call ClearTotalInfo
        
    Call InitCommVariable
    
    Call ClearBillRows: Call ClearMoney
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    IDKind.Enabled = (chkCancel.Value = 0)
    
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        If cboBaby.Visible Then cboBaby.Enabled = False
        
        Call NewBill(False)
        Set mobjBill = New ExpenseBill
        If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
        
        cboNO.Text = ""

        Call SetDisible
        If InStr(mstrPrivs, "显示开单人") = 0 Then
            cbo开单人.Visible = False
            If gbyt科室医生 = 0 Then
                lbl科室.Visible = False
            Else
                lbl开单人.Visible = False
            End If
        End If
        
        fraAppend.Enabled = False
        cbo结算方式.Enabled = False
        cboNO.Locked = False
        cmd配方.Enabled = False
        cmdYB.Enabled = False
        
        txtModi.Enabled = False
        txtIn.Text = ""
        txtIn.Enabled = False
        txtRePrint.Enabled = False
        
        txtInvoice.Text = ""
        txtInvoice.Locked = True
                
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = BillColType.Text_UnModify
        Next
        Call ShowDeleteCol(True)
        Bill.SetColColor BillCol.类别, &HE7CFBA  '不然要成白色
        
        If mbytInFun = 0 Then
            lbl应缴.Caption = "退款"
            lbl应缴.ForeColor = vbRed
            txt应缴.ForeColor = vbRed
            txt应缴.Text = "0.00"
        End If
        
        cboNO.SetFocus
    Else
        
        If InStr(mstrPrivs, "显示开单人") = 0 Then
            cbo开单人.Visible = True
            If gbyt科室医生 = 0 Then
                lbl科室.Visible = True
            Else
                lbl开单人.Visible = True
            End If
        End If
        
        txtRePrint.Enabled = True
        txtModi.Enabled = True
        txtIn.Text = ""
        txtIn.Enabled = True
        
        chkCancel.ForeColor = 0
        If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
        txtInvoice.Locked = Not (InStr(1, mstrPrivs, "修改票据号") > 0) And gblnStrictCtrl
        If mbytBilling <> 2 Then Call SetDisible(True)
        cmd配方.Enabled = True
        cmdYB.Enabled = True
        
        Call NewBill(IIf(Not mblnStartFactUseType, False, True), False)
        Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
        Call LoadAndSeek费别
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        Bill.SetColColor BillCol.类别, &HE7CFBA  '不然要成白色
        
        If mbytInFun = 0 Then
            lbl应缴.Caption = "应缴"
            lbl应缴.ForeColor = 0
            txt应缴.ForeColor = &HFF0000
            txt应缴.Text = "0.00"
        End If
        
        If mbytBilling = 2 Then
            fraInfo.Enabled = False
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            cbo开单科室.Enabled = True
            cbo开单人.Enabled = True
            
            fraAppend.Enabled = True
            If cbo结算方式.Visible Then cbo结算方式.Enabled = (mbytInState = 0)
            
            If mlng病人ID > 0 Then
                txtPatient.Text = "-" & mlng病人ID
                Call txtPatient_KeyPress(13)
                Bill.SetFocus
            Else
                txtPatient.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chk急诊_Click()
    If chk急诊.Visible And Visible And mbytInFun = 0 Then
        '需要重新预结算
        If cmd预结算.Visible Then
            Call InitBalanceGrid
            cmd预结算.TabStop = True
            cmdOK.Enabled = False
        End If
'''        Call zlClear结算卡
    End If
End Sub

Private Sub chk急诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk加班_Click()
    Dim blnAdd As Boolean
    
    If Not mblnDo Then Exit Sub
    If mbytInState = 1 Or chkCancel.Value = 1 Or mbytBilling = 2 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk加班.Visible Or Not Visible Then Exit Sub
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk加班.Value = 0 And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = 1
        End If
    End If
    If chk加班.Value = 1 And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = 0
        End If
    End If
    mobjBill.加班标志 = chk加班.Value
    
    '需要重新预结算
    If cmd预结算.Visible Then
        Call InitBalanceGrid
        cmd预结算.TabStop = True
        cmdOK.Enabled = False
    End If
'''    Call zlClear结算卡
    
    '全部重新计算价格
    If Not CheckBillsEmpty Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub AutoSplitBill()
    '功能:自动将所有单据按收费类别进行单据分组
    '     暂不处理医保,收取工本费模式下,引起的工本费变化,暂不处理

    Dim i As Integer, j As Integer, strFKind As String, strFeeKind As String
    Dim intMinPage As Integer, intMaxPage As Integer, intPage As Integer, intRows As Integer
    Dim intOrder As Integer, intMainItem_New As Integer, intMainItem_Old As Integer, strMainKind As String, curError As Currency
    Dim blnMainItem As Boolean
    
    If cmdAddBill.Enabled = False Then Exit Sub
        
    If mobjBill.Pages.Count = 1 Then
        For i = 1 To mobjBill.Pages(1).Details.Count
            If i = 1 Then
                strFeeKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(1).Details(i).收费类别, mobjBill.Pages(1).Details(i).执行部门ID)
            ElseIf strFeeKind <> IIf(gbytAutoSplitBill = 1, mobjBill.Pages(1).Details(i).收费类别, mobjBill.Pages(1).Details(i).执行部门ID) Then
                Exit For
            End If
        Next
        If i > mobjBill.Pages(1).Details.Count Then Exit Sub
    End If
        
    '序号最小的非划价单据
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO = "" Then Exit For
    Next
    If i > mobjBill.Pages.Count Then Exit Sub '单据全为空,或全是划价单
    intMinPage = i
    intMaxPage = mobjBill.Pages.Count
    If mobjBill.收费结算 <> "" Then curError = mobjBill.Pages(intMaxPage).误差金额 '多种结算方式的误差是存在最后一张单据上的
    
    For i = intMinPage To intMaxPage
        intMainItem_Old = 0
        intMainItem_New = 0
        strMainKind = ""
        If i <> intMinPage Then
            '1.移走与前面单据中类别相同的行
            j = 1
            intRows = mobjBill.Pages(i).Details.Count
            Do While j <= intRows
                If mobjBill.Pages(i).Details(j).从属父号 = 0 Then
                    blnMainItem = CheckMainItem(j, i)
                Else
                    blnMainItem = False
                End If
                If blnMainItem Then
                    intMainItem_Old = mobjBill.Pages(i).Details(j).序号
                    intMainItem_New = intMainItem_Old
                    strMainKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID)
                End If
            
                '从项的父号处理
                If mobjBill.Pages(i).Details(j).从属父号 = intMainItem_Old And intMainItem_Old <> 0 Then
                    If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID) = strMainKind Then
                        mobjBill.Pages(i).Details(j).从属父号 = intMainItem_New
                    Else
                        mobjBill.Pages(i).Details(j).从属父号 = 0
                    End If
                End If
                
                intPage = CheckKindInOtherPage(IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID), i, 1) '向前检查
                If intPage > 0 Then
                    intOrder = AddRowByOtherPageRow(mobjBill.Pages(i).Details(j), intPage)
                    If blnMainItem Then intMainItem_New = intOrder
                                            
                    Call DeleteDetail(j, i) '当前总行数已变化
                    j = j - 1
                    intRows = intRows - 1
                Else
                    If mobjBill.Pages(i).Details(j).从属父号 = intMainItem_Old And intMainItem_Old <> intMainItem_New Then  '主项移走了,从项没有动
                        mobjBill.Pages(i).Details(j).从属父号 = 0
                    End If
                End If
                j = j + 1
            Loop
        End If
        
        '2.移走与本单据中第一行类别不同的行.
        If mobjBill.Pages(i).Details.Count > 0 Then '可能因前面的移动,全部被移走了,单据为空
            strFKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(1).收费类别, mobjBill.Pages(i).Details(1).执行部门ID)
            If mobjBill.Pages(i).Details(1).从属父号 = 0 Then
                blnMainItem = CheckMainItem(1, i)
            Else
                blnMainItem = False
            End If
            If blnMainItem Then
                intMainItem_Old = 1
                intMainItem_New = 1
                strMainKind = strFKind
            End If
        End If
        j = 2
        intRows = mobjBill.Pages(i).Details.Count
        Do While j <= intRows
            If mobjBill.Pages(i).Details(j).从属父号 = 0 Then
                blnMainItem = CheckMainItem(j, i)
            Else
                blnMainItem = False
            End If
            If blnMainItem Then
                intMainItem_Old = mobjBill.Pages(i).Details(j).序号
                intMainItem_New = intMainItem_Old
                strMainKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID)
            End If
            
            If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID) <> strFKind Then
                
                '从项的父号处理
                If mobjBill.Pages(i).Details(j).从属父号 = intMainItem_Old And intMainItem_Old <> 0 Then
                    If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID) = strMainKind Then
                        mobjBill.Pages(i).Details(j).从属父号 = intMainItem_New
                    Else
                        mobjBill.Pages(i).Details(j).从属父号 = 0
                    End If
                End If
            
                intPage = CheckKindInOtherPage(IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID), i, 0) '向后检查
                If intPage = 0 Then
                    Call AddNewBill
                    intPage = mobjBill.Pages.Count
                End If
                intOrder = AddRowByOtherPageRow(mobjBill.Pages(i).Details(j), intPage)
                If blnMainItem Then intMainItem_New = intOrder
                
                Call DeleteDetail(j, i)
                j = j - 1
                intRows = intRows - 1
            Else
                If mobjBill.Pages(i).Details(j).从属父号 = intMainItem_Old And intMainItem_Old <> intMainItem_New Then '主项移走了,从项没有动
                    mobjBill.Pages(i).Details(j).从属父号 = 0
                End If
            End If
            j = j + 1
        Loop
    Next
    
    '3.删除那些因移走而产生的空单据
    i = 1
    intMaxPage = mobjBill.Pages.Count
    Do While i <= intMaxPage
        If CheckBillsEmpty(i) Then
            Call DelOneBill(i)
            i = i - 1
            intMaxPage = intMaxPage - 1
        End If
        i = i + 1
    Loop
    
    '刷新界面显示
    Call ShowDetails
    Call ShowMoney
    
    If mobjBill.收费结算 <> "" Then
        If mobjBill.Pages.Count = 1 Then
            mobjBill.Pages(1).收费结算 = mobjBill.收费结算
        Else
            Call ApportionMultiBalance(mobjBill.收费结算, curError)
        End If
    End If
End Sub

Private Function AddRowByOtherPageRow(tmpBillDetail As BillDetail, intPage As Integer) As Integer
'功能:将某单据行对象增加到指定的单据页中
    Dim int序号 As Integer
    
    With tmpBillDetail
        int序号 = mobjBill.Pages(intPage).Details.Count + 1
        Call mobjBill.Pages(intPage).Details.Add(.费别, .Detail, .收费细目ID, int序号, .从属父号, _
            .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, _
            .InComes, "", .保险项目否, .保险大类ID, .保险编码, .摘要, .原始数量, .原始执行部门ID)
    End With
    AddRowByOtherPageRow = int序号
End Function


Private Function CheckKindInOtherPage(ByVal strKind As String, ByVal intCurrentPage As Integer, bytWay As Byte) As Integer
'功能:检查非当前单据(并且不是划价单)中是否存在指定的收费类别或执行部门
'参数:bytWay-检查其它单据的方向,0-向后检查,1-向前检查
'返回:如果不存在则返回0,存在则返回第一个存在的单据序号
    Dim intBegin As Integer, intEnd As Integer, i As Integer, j As Integer

    If mobjBill.Pages.Count < 2 Then Exit Function
    If bytWay = 0 Then
        intBegin = intCurrentPage + 1
        intEnd = mobjBill.Pages.Count
    Else
        intBegin = 1
        intEnd = intCurrentPage - 1
    End If
    
    For i = intBegin To intEnd
        If mobjBill.Pages(i).NO = "" Then
            For j = 1 To mobjBill.Pages(i).Details.Count
                If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).收费类别, mobjBill.Pages(i).Details(j).执行部门ID) = strKind Then
                    CheckKindInOtherPage = i: Exit Function
                End If
            Next
        End If
    Next
End Function

Private Sub AddNewBill()
'功能：增加一张单据
    Dim objPage As New BillPage
    Dim i As Long

    '加入单据页标签
    If tbsBill.Tabs.Count >= 10 Then
        Call tbsBill.Tabs.Add(, , "单据" & tbsBill.Tabs.Count + 1)
    Else
        If tbsBill.Tabs.Count + 1 = 10 Then
            Call tbsBill.Tabs.Add(, , "单据1&0")
        Else
            Call tbsBill.Tabs.Add(, , "单据&" & tbsBill.Tabs.Count + 1)
        End If
    End If
    cmdDelBill.Enabled = True
    
    '加入单据页对象:即使是划价收费也保持一致
    mobjBill.Pages.Add objPage.Details
    
    '单据缺省的开单科室,开单人与当前相同
    i = mobjBill.Pages.Count
    If cbo开单科室.ListIndex <> -1 Then
        mobjBill.Pages(i).开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        mobjBill.Pages(i).开单部门ID = 0
    End If
    mobjBill.Pages(i).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    
    '加入结算集合:划价收费也要保持一致
    mcolBalance.Add Array()
        
    '多张单据时禁止修改,导入,退费功能
    txtModi.Text = ""
    txtModi.Enabled = False
    chkCancel.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub cmdAddBill_Click()
    Dim i As Long
    Dim strFirst费别 As String
    
    '不应有多余的空单据
    For i = 1 To mobjBill.Pages.Count
        If CheckBillsEmpty(i) Then
            MsgBox "第 " & i & " 张单据内容为空，请先在该单据中输入。", vbInformation, gstrSysName
            tbsBill.Tabs(i).Selected = True
            Bill.SetFocus: Exit Sub '缺省为直接输入费用
        End If
    Next
    
    If tbsBill.Tabs.Count >= 200 Then
        MsgBox "单据数量太多，请分成多次收费。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strFirst费别 = mobjBill.费别
            
    '增加单据
    Call AddNewBill
    
    '激活Click,显示新增加单据的内容(空白)
    tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
    
    If mobjBill.Pages(1).NO <> "" Then cbo费别.ListIndex = cbo.FindIndex(cbo费别, strFirst费别, True)
    
    Bill.SetFocus '缺省为直接输入费用
End Sub

Private Sub DelOneBill(ByVal intPage As Integer)
'功能：删除指定的单据
    Dim blnCurEmpty As Boolean, i As Integer
    
    blnCurEmpty = CheckBillsEmpty(intPage)
    
    '删除单据集合中的内容
    mobjBill.Pages.Remove intPage
    
    '删除结算集合
    mcolBalance.Remove intPage
    
    '删除页卡之后自动重新定位,并且不会激活Click
    tbsBill.Tabs.Remove intPage
    For i = 1 To tbsBill.Tabs.Count
        If i = 10 Then
            tbsBill.Tabs(i).Caption = "单据1&0"
        ElseIf i < 10 Then
            tbsBill.Tabs(i).Caption = "单据&" & i
        Else
            tbsBill.Tabs(i).Caption = "单据" & i
        End If
    Next
    If tbsBill.Tabs.Count = 1 Then cmdDelBill.Enabled = False
        
    '需要重新预结算
    If Not blnCurEmpty And cmd预结算.Visible Then
        Call InitBalanceGrid
        cmd预结算.TabStop = True
        cmdOK.Enabled = False
    End If
            
'''    '刘兴洪:??
'''    If Not blnCurEmpty Then
'''        Call zlClear结算卡
'''    End If
'''
    '打开修改及退费功能
    If tbsBill.Tabs.Count = 1 Then
        txtModi.Text = ""
        txtModi.Enabled = True
        chkCancel.Enabled = True
        cmdDelete.Enabled = True
    End If
    
    '激活Click,显示新定位单据的内容
    mintPage = 0 '强行激活
    Call tbsBill_Click
End Sub

Private Sub cmdDelBill_Click()
'功能：删除当前单据
    Dim i As Long
    
    If MsgBox("确实要删除第 " & mintPage & " 张单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If mobjBill.Pages(mintPage).NO = "" Then
            Bill.SetFocus
        Else
            If txtPatient.Text = "" Then
                txtPatient.SetFocus
            Else
                'If txt缴款.Enabled And txt缴款.Visible Then
                '    txt缴款.SetFocus
                If cmd预结算.Enabled And cmd预结算.Visible Then
                    cmd预结算.SetFocus
                ElseIf cmdOK.Enabled And cmdOK.Visible Then
                    cmdOK.SetFocus
                End If
            End If
        End If
        Exit Sub
    End If
            
    '删除单据
    Call DelOneBill(mintPage)
    
    '重新计算
    Call ShowMoney(-1)  '其它单据费用未变
    
    '重新设置工本费(包含了重新计算)
    If gTy_Module_Para.bln工本费 Then
        If Not CheckBillsEmpty Then Call SetFactMoney
    End If
End Sub

Private Function LoadMultiBills(ByVal lng病人ID As Long, ByVal bln不允许多单据 As Boolean, _
    ByVal lng挂号科室 As Long, Optional blnCard As Boolean) As Boolean
'功能：一次性读取病人的多张划价单,该过程在病人读取成功之后调用
'参数：bln不允许多单据，医保连续收费或不支持多单据收费时，不允许返回多张划价单收费
'      lng挂号科室,当通过挂号单输入时,传入病人当前挂号单的挂号科室

    Dim objPage As New BillPage
    Dim arrBills As Variant, strBills As String
    Dim blnRead As Boolean, i As Long, k As Long
    
    If Not (gblnMulti And gblnSeekBill) Then Exit Function
    
    If lng病人ID = 0 Then Exit Function
    i = SeekPatiBill(lng病人ID)
    If i = 0 Then Exit Function
    If gblnUnPopPriceBill Then
        strBills = frmPatiPrice.GetPriceBillString(lng病人ID, bln不允许多单据, lng挂号科室, mTy_Para.bln住院病人门诊收费, blnCard)
    Else
        strBills = frmPatiPrice.FindBill(Me, mstrPrivs, lng病人ID, bln不允许多单据, lng挂号科室, mTy_Para.bln住院病人门诊收费, blnCard)
    End If
     
    If strBills = "" Then Exit Function
    
    
    LoadMultiBills = True
    '清除现有单据的内容
    '---------------------------------------------------------------------
    mstrInNO = "": txtModi.Text = ""
    Call ClearTotalInfo
    Call ClearPayInfo
    Call ClearBillRows
        
    '预结算支持时才清除,否则会自动算
    If cmd预结算.Visible Then
        Call InitBalanceGrid
    End If
    

    '读取划价单重新计算时,需要累计显示在表格中
    '刘兴洪,问题:22343;只有在输入缴款金额后,才存在累计的问题
    '  Not gbln缴款结束 取掉
    '51670: 分单病人累计和多病人累计
    If gTy_Module_Para.byt缴款控制 <> 1 And gTy_Module_Para.byt缴款控制 <> 3 Or mstrPrePati = "" Then
        Call ClearMoney
    End If
    
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
        
    '多单据收费:只保留一页对象
    For i = mobjBill.Pages.Count To 1 Step -1
        mobjBill.Pages.Remove i
    Next
    mobjBill.Pages.Add objPage.Details
    
    '多单据收费:恢复缺省单据页卡
    mintPage = 1
    For i = tbsBill.Tabs.Count To 1 Step -1
        tbsBill.Tabs(i).Tag = ""
        If i <> 1 Then tbsBill.Tabs.Remove i
    Next
        
    '读取显示每张划价单
    '---------------------------------------------------------------------
    mblnNOMoved = False '划价单读取不从后备表中读
    k = 1
    mblnDoing = True '表明正在自动读
    arrBills = Split(strBills, ",")
    For i = 0 To UBound(arrBills)
        Me.Refresh
        '增加单据页标签(同cmdAdd_Click内容)
        '-----------------------------------------------------------------------
        If k > 1 And mobjBill.Pages(mobjBill.Pages.Count).NO <> "" Then
            If tbsBill.Tabs.Count >= 10 Then
                Call tbsBill.Tabs.Add(, , "单据" & tbsBill.Tabs.Count + 1)
            Else
                If tbsBill.Tabs.Count + 1 = 10 Then
                    Call tbsBill.Tabs.Add(, , "单据1&0")
                Else
                    Call tbsBill.Tabs.Add(, , "单据&" & tbsBill.Tabs.Count + 1)
                End If
            End If
            
            '加入单据页对象:即使是划价收费也保持一致
            mobjBill.Pages.Add objPage.Details
            
            '加入结算集合:划价收费也要保持一致
            mcolBalance.Add Array()
                
            '多张单据时禁止修改及退费功能
            txtModi.Enabled = False
            chkCancel.Enabled = False
            cmdDelete.Enabled = False
                
            '激活Click,显示新增加单据的内容(空白)
            tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
        End If
                
        '读取划价单据内容(同cboNO_KeyPress)
        '----------------------------------------------------------------------
        blnRead = ReadBill(arrBills(i), 1, False)
        If blnRead Then k = k + 1: cboNO.Text = arrBills(i)
    Next
    Bill.Active = False
    chk加班.Enabled = False
    If mbln补费 And mstr最后转科时间 <> "" Then
        txtDate.Text = Format(CDate(mstr最后转科时间) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
    Else
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    cmdDelBill.Enabled = tbsBill.Tabs.Count > 1
    
    mblnDoing = False '表明自动读取完毕
    
    
    '显示摘要
    Call Bill_EnterCell(1, BillCol.项目)
    '计算票据是否充足
    If gTy_Module_Para.byt票据分配规则 <> 0 Then
        Dim str发票号 As String, int张数 As Integer
        If mintInvoicePrint <> 0 Then
            If zlExeCuteBillNoSplit(True, 1, mlng领用ID, strBills, 0, txtInvoice.Text, Now, 1, str发票号, int张数) Then
                Call zlCheckFactIsEnough(int张数)
            End If
        End If
    End If
    
    If mstrYBPati = "" And gbln划价立即缴款 Then
       If cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus
        End If
    End If
End Function



Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0, Optional ByVal lng病人ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '入参:blnFact-是否重新取发票号
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng病人ID As Long
    Dim intInsure As Integer
    lng病人ID = IIf(lng病人ID_In <> 0, lng病人ID_In, mobjBill.病人ID)
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    
    If lng病人ID = 0 Then
        '上次病人ID
        If (mbytInFun = 0 Or mbytInFun = 1) And txtPatient.Text = mstrPrePati And mlngPrePati <> 0 Then
            lng病人ID = mlngPrePati
        End If
    End If
    If lng病人ID = 0 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    
    Call ZlShowBillFormat(mlngModul, lblFormat, mintInvoiceFormat)
    If blnFact Then Call RefreshFact
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

Private Sub RefreshFact()
    '功能：刷新收费票据号
    If mintInvoicePrint = 0 Then Exit Sub
    If gblnStrictCtrl Then
        'lblFact.tag主要是检查发票号是否手工输入的.手工输入的,发票号为空,否则是自动产生的发票号
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            If zlGetInvoiceGroupUseID(mlng领用ID) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            '严格：取下一个号码
            txtInvoice.Text = GetNextBill(mlng领用ID)
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            If mblnStartFactUseType Then Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '松散：取下一个号码
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModul)))
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub cmdDelete_Click()

    If (mbytInFun = 0 And Not gblnMulti) Or mbytInFun = 2 Then
        cmd配方.Enabled = Not cmd配方.Enabled
        cmdYB.Enabled = Not cmdYB.Enabled
    End If
    If frmMultiBills.ShowMe(Me, 1, mstrPrivs, "", "", , mlng领用ID, mblnOneCard) Then
        Call RefreshFact
        If gbln累计 Then txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub cmdIDCard_Click()
    Dim strCommon As String, intAtom As Integer
            
    On Error Resume Next
    If gobjPatient Is Nothing Then
        Set gobjPatient = CreateObject("zl9Patient.clsPatient")
        If gobjPatient Is Nothing Then Exit Sub
    End If
    
    Err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    Call gobjPatient.IDCard(Me, gcnOracle, glngSys, gstrDBUser)
    Call GlobalDeleteAtom(intAtom)
    
    If txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub cmdRegist_Click()
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
            
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Sub
    End If
    
    Err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    
    blnOK = gobjRegist.Register(Me, gcnOracle, glngSys, gstrDBUser, gblnSharedInvoice, IIf(gblnSharedInvoice, mlngShareUseID, 0))
    Call GlobalDeleteAtom(intAtom)
    '完成挂号
    '刷新票据号
    If gblnSharedInvoice And blnOK Then
        If txtInvoice.Enabled Then Call RefreshFact
    End If
    If txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub cmd配方_Click()
    Call ShowCHRecipe
End Sub

Private Sub zlChangePatiSource(ByVal int病人来源 As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:改变病人来源状态
    '编制:刘兴洪
    '日期:2010-01-13 11:23:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Panel As MSComctlLib.Panel
    
    Set Panel = sta.Panels("PatiSource")
    Select Case int病人来源
    Case 1 '门诊
        Set Panel.Picture = imgPati.ListImages("OutPati").Picture
        Panel.ToolTipText = "门诊病人"
        gstr药房单位 = "门诊单位": gstr药房包装 = "门诊包装"
    Case Else    '住院
        Set Panel.Picture = imgPati.ListImages("InPati").Picture
        Panel.ToolTipText = "住院病人"
        gstr药房单位 = "住院单位": gstr药房包装 = "住院包装"
    End Select
    sta.Panels(Pan.C2提示信息).Text = "已将病人来源设置为" & IIf(int病人来源 = 1, "门诊病人", "住院病人")
    Set mrsUnit = GetDepartments("", gint病人来源 & ",3")
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Call FillDept(mlngDeptID)
    Call FillDoctor
    Call ClearFullBill(False)    '主要是设置mobjBill.门诊标志
End Sub
Private Sub picAppend_Resize()
    Dim sngLeft As Single
    Err = 0: On Error Resume Next
    sngLeft = vsBalance.Left + vsBalance.Width + 100
    If fra缴款.Visible Then sngLeft = fra缴款.Left + fra缴款.Width + 100
    If mbytInFun = 0 Then
        cmdOK.Left = sngLeft + (ScaleWidth - sngLeft - cmdOK.Width) \ 2 '  ScaleWidth - cmdOK.Width - 100
        cmdCancel.Left = cmdOK.Left
        cmdPrint.Left = cmdOK.Left
        cmd预结算.Left = cmdOK.Left
    End If
    If mbytInFun <> 0 Then Exit Sub
    If mbytInState <> 3 Then
        lbl预交冲款.Left = vsBalance.Left
        txt预交冲款.Left = lbl预交冲款.Left + lbl预交冲款.Width + 10
        txt预交冲款.Top = picAppend.ScaleHeight - txt预交冲款.Height - 10
        lbl预交冲款.Top = txt预交冲款.Top + (txt预交冲款.Height - lbl预交冲款.Height) \ 2
        txt预交冲款.Width = vsBalance.Left + vsBalance.Width - txt预交冲款.Left
    End If
    If mbytInState = 0 Then
        vsBalance.Height = picAppend.ScaleHeight - vsBalance.Top - 20
    ElseIf mbytInState = 3 Then
        '退款
        vsBalance.Height = picAppend.ScaleHeight - vsBalance.Top - 20
        fra缴款.Left = vsBalance.Left + vsBalance.Width + 20
    Else
        vsBalance.Height = IIf(lbl预交冲款.Visible = False, picAppend.ScaleHeight, txt预交冲款.Top) - vsBalance.Top - 20
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Dim lngR As Long
    If Panel.Key = "Calc" Then
        lngR = FindWindow("SciCalc", "计算器")
        If lngR <> 0 Then
            BringWindowToTop lngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
    ElseIf Panel.Key = "Drugstore" Then
        With frmSetExpence
            .mlngModul = mlngModul
            .mstrPrivs = mstrPrivs
            .mbytInFun = mbytInFun
            .mblnSetDrugStore = True
            .Show 1, Me
        End With
    ElseIf Panel.Key = "PatiSource" Then
        If gbln病人来源受权限控制 And InStr(1, mstrPrivs, ";参数设置;") = 0 Or mbln补费 Then
            '授权限控制,不能更改
            Exit Sub
        End If
        If Not CheckBillsEmpty Or txtPatient.Text <> "" Then
            If MsgBox("如果切换病人来源,将清空当前单据和病人信息" & vbCrLf & "你确定要继续吗?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If gint病人来源 = 1 Then    '门诊
           ' Set Panel.Picture = imgPati.ListImages("InPati").Picture
            'Panel.ToolTipText = "住院病人"
            gint病人来源 = 2
            'gstr药房单位 = "住院单位": gstr药房包装 = "住院包装"
        Else
            'Set Panel.Picture = imgPati.ListImages("OutPati").Picture
            'Panel.ToolTipText = "门诊病人"
            gint病人来源 = 1
            'gstr药房单位 = "门诊单位": gstr药房包装 = "门诊包装"
        End If
        
        zlDatabase.SetPara "病人来源", gint病人来源, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        Call zlChangePatiSource(gint病人来源)
        mblnAutoChangePati = False
    ElseIf Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        If Not gbln简码切换 Then Exit Sub     '35242
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

Private Sub ShowDeposit(ByVal lngPatientID As Long)
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Sum(金额), 0) 预交款总额, Nvl(Sum(冲预交), 0) 冲预交总额 From 病人预交记录 Where 病人id = [1] And 记录性质 In(1,11) and nvl(预交类别,2)=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID)
    
    If rsTmp.RecordCount > 0 Then
        MsgBox "预交款总额:" & Format(rsTmp!预交款总额, "0.00") & vbCrLf & "冲预交总额:" & Format((rsTmp!冲预交总额 - Original.冲预交款), "0.00") & vbCrLf & _
               "未 结 费用:" & Format(Val(cmdCancel.Tag), "0.00") & vbCrLf & _
               "可用预交款:" & Format((rsTmp!预交款总额 - (rsTmp!冲预交总额 - Original.冲预交款 + Val(cmdCancel.Tag))), "0.00"), vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub sta_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel Is sta.Panels(Pan.C4预交信息) And mrsInfo.State = 1 Then
        Call ShowDeposit(mrsInfo!病人ID)
    End If
End Sub

Private Sub tbsBill_Click()
'功能：显示选定页卡的页单据内容
'说明：目前只有收费时才可能会进入
    Dim i As Integer, str费别 As String, blnLock As Boolean
    
    '相同点击时退出(只有一张时相当于不处理)
    If tbsBill.SelectedItem.Index = mintPage Then Exit Sub
    mintPage = tbsBill.SelectedItem.Index
    
    '清除表格显示的内容
    Call ClearBillRows
    If mobjBill.Pages(mintPage).Details.Count > 0 Then
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
    Else
        Bill.Rows = 2
    End If
    
    Call InitBillColumnColor
    
    '设置列号
    Call SetColNum
    
    If Not mblnDoing Then
        '重新显示单据刷新,并确定单据可否编辑
        mblnDoing = True
        cboNO.Text = mobjBill.Pages(tbsBill.SelectedItem.Index).NO
        If mobjBill.Pages(tbsBill.SelectedItem.Index).NO = "" Then
            Bill.Active = True
            mbln不重算价格 = True
            
            '显示开单部门,开单人
            If mbytInFun = 0 Then '因收费存在医嘱划价单混合收费时锁住,所以解锁
                cbo开单科室.Locked = False
                cbo开单人.Locked = False
            End If
            Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
                        
            '动态费别的显示,要在科室显示之后
            If cbo费别.Visible Then
                str费别 = zlStr.NeedName(cbo费别.Text)
                blnLock = cbo费别.Locked
            End If
            
            cbo费别.Visible = True
            lbl动态费别.Visible = True
            lbl动态费别.BorderStyle = 0
            lbl动态费别.Left = cbo费别.Left + cbo费别.Width + 60
            Call LoadAndSeek费别
            
            If str费别 <> "" Then Call zlControl.CboLocate(cbo费别, str费别)
            If cbo费别.ListIndex <> -1 Then cbo费别.Locked = blnLock
            cbo费别.TabStop = Not cbo费别.Locked And gbln费别
            
            mbln不重算价格 = False
            Call ShowDetails
        Else
            Bill.Active = False
            Call ReadBill(mobjBill.Pages(mintPage).NO, 1, False, , True)
        End If
        mblnDoing = False
        
        '缺省定位单元
        If mobjBill.Pages(tbsBill.SelectedItem.Index).NO = "" Then
            If mobjBill.Pages(mintPage).Details.Count = 0 Then
                Bill.Col = Bill.MsfObj.FixedCols
            Else
                Bill.Col = Bill.PrimaryCol
                mlngPreRow = 0
            End If
            Bill.Row = 1
        ElseIf Visible Then
            sta.Panels(Pan.C2提示信息).Text = ""
        End If
        If Visible Then Bill.SetFocus
    End If
End Sub

Private Function CheckBillsEmpty(Optional ByVal intPage As Integer) As Boolean
'功能：判断是否多单据的内容都为空
'参数：intPage=是否检查指定页,否则检查所有页
    Dim i As Integer
    
    If intPage = 0 Then
        For i = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(i).NO <> "" Then
                Exit Function
            ElseIf mobjBill.Pages(i).Details.Count > 0 Then
                Exit Function
            End If
        Next
    Else
        If mobjBill.Pages(intPage).NO <> "" Then
            Exit Function
        ElseIf mobjBill.Pages(intPage).Details.Count > 0 Then
            Exit Function
        End If
    End If
    CheckBillsEmpty = True
End Function

Private Function ClearFullBill(ByVal bln提示 As Boolean, _
    Optional blnClearPatiInfor As Boolean = True, _
    Optional blnNotClearYb As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除单据信息
    '入参:blnNotClearYb-不清除医保病人
    '返回:清除成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-26 11:55:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strYBPati As String, intInsure As Integer
    Dim blnAdd As Boolean, strYBBill As String
    Dim cur个帐余额 As Currency, cur个帐透支 As Currency, blnYB结算作废 As Boolean
    
    strYBPati = mstrYBPati: intInsure = mintInsure: strYBBill = mstrYBBill
    cur个帐余额 = mcur个帐余额: cur个帐透支 = mcur个帐透支: blnYB结算作废 = mblnYB结算作废
    blnAdd = cmdAddBill.Enabled
    '不清除医保信息
    If bln提示 Then
        If MsgBox("确实要清除当前单据中的内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    If Not blnNotClearYb Then
        If YBIdentifyCancel = False Then '取消医保病人身份验证
            Exit Function                '返回假时，不清除
        End If
    End If
    
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    Call ClearDisplaySHow
    Call ClearPayInfo
    
    If chkCancel.Value = 1 Then '退据单状态
        chkCancel.Value = 0
    Else
        mstrInNO = "": txtModi.Text = ""
        mlngFirstID = 0: mstrFirstWin = ""
        
        If blnClearPatiInfor Then Call ClearPatientInfo(blnClearPatiInfor)
        Call ClearTotalInfo(True)
        
        Call InitCommVariable
        
        If mbytInFun = 0 And gbln累计 Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
        End If
    End If
    
    Call ClearBillRows
    Call ClearMoney
    Call SetDisible(True)
    Call NewBill(IIf(mblnStartFactUseType, False, True), blnClearPatiInfor, Not mbln补费)
    If mbln补费 Then
        With mobjBill
            .病人ID = IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID)
            .主页ID = IIf(mbln补费 And mlng主页ID <> 0, mlng主页ID, Nvl(mrsInfo!主页ID, 0))
            .标识号 = IIf(gint病人来源 = 2, Nvl(mrsInfo!住院号, 0), Nvl(mrsInfo!门诊号, 0))
            .姓名 = "" & mrsInfo!姓名
            .性别 = "" & mrsInfo!性别
            .年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
            .床号 = "" & mrsInfo!当前床号
            .病区ID = IIf(mbln补费 And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!当前病区ID)))
            .科室ID = IIf(mbln补费 And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!当前科室id)))
            .费别 = zlStr.NeedName(cbo费别.Text) '以当前有效为准
        End With
        Bill.SetFocus
    End If
    If blnNotClearYb And intInsure <> 0 Then
        mintInsure = intInsure: mstrYBBill = strYBBill: mstrYBPati = strYBPati
        mcur个帐余额 = cur个帐余额: mcur个帐透支 = cur个帐透支: mblnYB结算作废 = blnYB结算作废
        sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
        sta.Panels(Pan.C3个人帐户).Visible = True
        Call SetPatientEnableModi(False)
        '75259：李南春,2014-7-10，病人姓名显示颜色处理
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), vbRed)
            Else
                txtPatient.ForeColor = vbRed
            End If
        Else
            txtPatient.ForeColor = vbRed
        End If
        cmdAddBill.Enabled = blnAdd
    End If
    sta.Panels(Pan.C2提示信息).Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    ClearFullBill = True
End Function


Private Function GetOneCardMoney(Optional ByVal str收费结算 As String) As Currency
'功能：获取所有单据或某张单据使用一卡通结算的金额
'参数：str收费结算-当前单据的结算串
    Dim arrTmp As Variant, strTmp As String, i As Long
    
    If mblnOneCard = False Then Exit Function
    
    If str收费结算 <> "" Then
        arrTmp = Split(str收费结算, "||")
        For i = 0 To UBound(arrTmp)
            mrsOneCard.Filter = "结算方式='" & Split(arrTmp(i), "|")(0) & "'"
            If mrsOneCard.RecordCount > 0 Then
                GetOneCardMoney = Val(Split(arrTmp(i), "|")(1))
                Exit Function
            End If
        Next
    Else
        If mobjBill.Pages(1).收费结算 = "" Then
            mrsOneCard.Filter = "结算方式='" & zlStr.NeedName(cbo结算方式) & "'"
            If mrsOneCard.RecordCount > 0 Then GetOneCardMoney = GetMustPaySum
        Else
            arrTmp = Split(mobjBill.收费结算, "||")
            For i = 0 To UBound(arrTmp)
                mrsOneCard.Filter = "结算方式='" & Split(arrTmp(i), "|")(0) & "'"
                If mrsOneCard.RecordCount > 0 Then
                    GetOneCardMoney = Val(Split(arrTmp(i), "|")(1))
                    Exit Function
                End If
            Next
        End If
    End If
End Function
Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是手术输入情况(如果不存在主要手术,但存在附加手术,则禁止
    '入参:
    '出参:lngRow-返回附加手术的行
    '返回:存在主手术或没有输入附加手术,返回true,否则返回False
    '编制:
    '修改:刘兴洪(退号时,增加定位功能),增加参数;strBackNo
    '日期:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   '指定行
    Dim i As Long, p As Long
    lngCount = 0
    For p = 1 To mobjBill.Pages.Count
         For i = 1 To mobjBill.Pages(p).Details.Count
            lngCount = 0
            If mobjBill.Pages(p).Details(i).收费类别 = "F" Then
               If mobjBill.Pages(p).Details(i).附加标志 = 0 Then lngCount = 0: Exit For  '存在主要手术,则不检查,直接返回true
               lngCount = lngCount + 1  '表示附加手术
               If lngRow <= 0 Then lngRow = i
            End If
        Next
        If lngCount > 0 Then Exit For
    Next
    If lngCount <> 0 Then
          MsgBox "单据中不存主要手术,但存在附加手术,请检查！", vbInformation, gstrSysName
          Err = 0: On Error GoTo Errhand:
          If p <= tbsBill.Tabs.Count Then tbsBill.Tabs(p).Selected = True
          '定位行:
          Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
Errhand:
End Function
Private Function WriteOff() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐销帐处理
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 10:04:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, i As Long
    Dim cllPro As Collection, cur消费金额 As Currency
    
    If Not (mbytInState = 3 Or (mbytInState = 0 And chkCancel.Value = 1 And chkCancel.Visible)) Then Exit Function
    If mbytInState = 0 And mstrInNO = "" Then
        MsgBox "没有正确读取单据内容,不能执行该操作！", vbInformation, gstrSysName
        cboNO.SetFocus: Exit Function
    End If
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, Bill.COLS - 1) = "√" And Bill.RowData(i) > 0 Then
            strSQL = strSQL & "," & Bill.RowData(i)
        End If
    Next
    If strSQL = "" Then
        MsgBox "请至少选择一个要退费的项目。", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    Set cllPro = New Collection
    '所有行选择处理
    strSQL = Mid(strSQL, 2)
    i = GetBillRows(mstrInNO, IIf(mbytInFun = 2, 2, 1))
    
    If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
    
    On Error GoTo errHandle
      If zlCheckIsMzToZY(mstrInNO, 2) Then
        MsgBox "注意:" & vbCrLf & _
                      "    该单据已经被门诊费用转住院费用 " & vbCrLf & _
                      "    或已经审核了门诊费用转住院费用,不能再销帐", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '问题:37307
    If (mbytBilling = 0 And Val(txt合计.Text) <> 0) And gbyt预存款退费验卡 <> 0 Then
        cur消费金额 = Val(txt合计.Text)
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.病人ID, cur消费金额, mlngModul, 1, , , , , , (gbyt预存款退费验卡 = 2)) Then Exit Function
    End If
    strSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    WriteOff = True
    
    '110319
    If mblnDrugMachine Then
        Dim rsTemp As ADODB.Recordset
        Dim strReturn As String, strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
        strSQL = "Select Id As 费用id, -1 * Nvl(付数, 1) * 数次 As 退药数量" & vbNewLine & _
                " From 门诊费用记录" & vbNewLine & _
                " Where 记录性质 = 2 And 记录状态 = 2 And NO = [1] And 收费类别 In ('5', '6', '7')" & vbNewLine & _
                "       And 登记时间 + 0 = (Select Max(登记时间)" & vbNewLine & _
                "                       From 门诊费用记录" & vbNewLine & _
                "                       Where 记录性质 = 2 And 记录状态 = 2 And NO = [1])"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询本次退费项目", mstrInNO)
        Do While Not rsTemp.EOF
            strData = strData & ";" & Nvl(rsTemp!费用id) & "," & Nvl(rsTemp!退药数量)
            rsTemp.MoveNext
        Loop
        If strData <> "" Then
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub SetAllDelSelAll()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选中所有退费记录
    '编制:刘兴洪
    '日期:2011-08-30 14:55:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.TextMatrix(i, Bill.COLS - 1) = "√"
    Next
    Call ReCalce退款
End Sub



Private Function DelChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费处理
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 10:07:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, i As Long, cur消费金额 As Currency, lng领用ID As Long
    Dim rsOneCard As ADODB.Recordset, rsThree As ADODB.Recordset
    Dim cllPro As Collection, bln退现 As Boolean
    Dim cur金额 As Currency, cur误差金额 As Currency, cur余额 As Currency
    Dim blnAll部份退费 As Boolean, blnCur部份退费 As Boolean, strInvoices As String
    Dim strBalance As String, strTmp As String, objICCard As Object, strCardNo As String
    Dim strInvoice As String, dtDateDel As Date, lng结帐ID As Long
    Dim blnTrans As Boolean, blnTransMedicare As Boolean, strAdvance As String
    Dim strErrMsg As String '错误信息
    Dim blnCommited As Boolean '已调接口
    Dim dblErrMoney As Double   '未成功交易的金额
    Dim cllBalance As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim blnExistThreeSwap As Boolean '存在第三方交易
    Dim blnExistOneCard As Boolean '存在一卡通交易
    Dim str退结算方式 As String, dbl退金额 As Double
    Dim str保险结算 As String, strOtherBalance As String
    Dim strThreeBalance As String, bln药品 As Boolean, intCol As Integer
    Dim lng冲销ID As Long, lng病人ID As Long, intCol类别 As Integer
    Dim strReclaimInvoice As String, intInvoiceFormat As Integer '回收票据:票据分配规则为1和2时有效25187
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    Dim bln完全退费 As Boolean
    
    If Not (mbytInState = 3 Or (mbytInState = 0 And chkCancel.Value = 1 And chkCancel.Visible)) Then Exit Function
    
    If mbytInState = 0 And mstrInNO = "" Then
        MsgBox "没有正确读取单据内容,不能执行该操作！", vbInformation, gstrSysName
        cboNO.SetFocus: Exit Function
    End If
    If CheckBillExistReplenishData(1, , mstrInNO) = True Then
        MsgBox "选择的记录进行了医保补充结算，不允许进行重打或补打票据操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '47400
    intCol类别 = -1
    For intCol = 0 To Bill.COLS - 1
        Select Case Bill.TextMatrix(0, intCol)
        Case "类别"
            intCol类别 = intCol: Exit For
        Case Else
        End Select
    Next
    bln药品 = False
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, Bill.COLS - 1) = "√" And Bill.RowData(i) > 0 Then
            strSQL = strSQL & "," & Bill.RowData(i)
            If intCol类别 <> -1 Then     '47400
                If Bill.TextMatrix(i, intCol类别) Like "*西*药*" _
                    Or Bill.TextMatrix(i, intCol类别) Like "*中*药*" _
                    Or Bill.TextMatrix(i, intCol类别) Like "*卫材*" Then
                    bln药品 = True
                    '81190,冉俊明,退费业务向发药机上传退费信息
                    If Not Bill.TextMatrix(i, intCol类别) Like "*卫材*" Then
                        If InStr(strReturnRecipt & "|", "|" & mstrInNO & "," & Bill.TextMatrix(i, BillCol.执行科室ID) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & mstrInNO & "," & Bill.TextMatrix(i, BillCol.执行科室ID)
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If strSQL = "" Then
        MsgBox "请至少选择一个要退费的项目。", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    Set cllPro = New Collection
    '所有行选择处理
    strSQL = Mid(strSQL, 2)
    '47400
    If bln药品 Then
        If zlCheckDrugIsPutDrug(mstrInNO) = False Then Exit Function
    End If
    '获取本次回收票据
    '单据号1:序号1(1..n);单据号2:序号2(1..n
    strReclaimInvoice = zlGetReclaimInvoice(mstrInNO & ":" & strSQL)
    
    i = GetBillRows(mstrInNO, IIf(mbytInFun = 2, 2, 1))
    
    If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
    
    On Error GoTo errHandle
    '刘兴洪:28947
    If mintInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mintInsure) = False Then Exit Function
    End If
    
    '可能是多单据收费中的一张,获取多个NO,收费部份退费时收回以前的单据,全部重打印票据
    '问题:51080,53145
    If gTy_Module_Para.byt票据分配规则 <> 0 Then
        strNos = GetMultiNOs(mstrInNO, , , True)
    Else
        strNos = GetMultiNOs(mstrInNO, , , False)
    End If
    
    If zlCheckIsMzToZY(mstrInNO, 1) Then
        MsgBox "注意:" & vbCrLf & _
                      "    该单据已经被门诊费用转住院费用 " & vbCrLf & _
                      "    或已经审核了门诊费用转住院费用,不能再退费", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '当前单据本次是否部份退费
    blnCur部份退费 = Not (BillDeleteAll(mstrInNO, 1, mblnHaveExcuteData) And strSQL = "")
    If blnCur部份退费 Then blnAll部份退费 = True '这张单据为部份退费,则所有单据为部份退费
    If Not blnCur部份退费 Then bln完全退费 = True
    bln完全退费 = bln完全退费 And Not BillExistDelete(Replace(strNos, "'", ""), 1)

    If Not blnCur部份退费 Then
        strTmp = ""
        If UBound(Split(strNos, ",")) + 1 > 1 Then
            strTmp = Replace("," & strNos, ",'" & mstrInNO & "'", "")
            If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
            '只有当多张单据中的其它单据已全退,且当前单据也是全退时,才不算是部分退
            If BillExistMoney(strTmp, 1) Then blnAll部份退费 = True
        End If
    End If
    If blnAll部份退费 Then
        If InStr(mstrPrivs, "部份退费") = 0 Then
            MsgBox "你没有权限执行部份退费操作！", vbInformation, gstrSysName
            Exit Function
        End If
        If mintInsure > 0 And blnCur部份退费 Then
            If strSQL = "" Then
                MsgBox "该张单据包含保险结算费用，而其中一些项目可能已经执行，不允许部份退费。", vbInformation, gstrSysName
            Else
                MsgBox "该张单据包含保险结算费用，不允许部份退费。", vbInformation, gstrSysName
            End If
            Call SetAllDelSelAll: Exit Function
        End If
        '多张单据部份退费时的工本费检查
        If gTy_Module_Para.bln工本费 Then
            MsgBox "自动收取工本费时不允许部份退费。", vbInformation, gstrSysName: Exit Function
        End If
    End If
    strBalance = ""      '记录医保不允许退回的结算方式
    strThreeBalance = ""
    Dim dblDelMoney As Double '当前结算方式的退款金额
    With vsBalance
        For i = 0 To .Rows - 1
            strTmp = vsBalance.TextMatrix(i, 0)
            dblDelMoney = Val(vsBalance.TextMatrix(i, 1))
            If strTmp <> "" Then
                '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
                Select Case Val(.Cell(flexcpData, i, 0))
                Case 1
                Case 2  '医保类
                    str保险结算 = str保险结算 & "," & strTmp
                    If mintInsure <> 0 Then
                           If mblnYB结算作废 Then
                                If Not gclsInsure.GetCapability(support门诊结算作废, , mintInsure, strTmp) Then
                                    strBalance = strBalance & "," & strTmp
                                End If
                            Else    '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                                If strTmp = mstr个人帐户 Then strBalance = "," & strTmp
                            End If
                    End If
                Case 3   '3-医疗卡
                    '如果支持退现,则直接退现；否则，只退当前应退金额
                    If Val(.RowData(i)) = -1 And Val(.TextMatrix(i, 1)) = 0 Then
                        strThreeBalance = strThreeBalance & "," & strTmp
                    Else
                        blnExistThreeSwap = True
                    End If
                    If blnExistThreeSwap Then
                        If blnCur部份退费 Or Not bln完全退费 Then
                            If blnCur部份退费 And Not mTyDelFee.blnSingleBalance Or mTyDelFee.bln三方卡全退 Then
                                MsgBox "当前单据使用了第三方结算交易,不能进行部分退费！", vbInformation, gstrSysName
                                Call SetAllDelSelAll: Exit Function
                            End If
                            strThreeBalance = strThreeBalance & "," & strTmp & "|" & dblDelMoney
                        End If
                        mTyDelFee.rsBlance.Filter = 0
                        mTyDelFee.rsBlance.Filter = "性质=" & Val(.Cell(flexcpData, i, 0)) & " And 结算方式='" & strTmp & "'"
                        If mTyDelFee.rsBlance.EOF Then
                            MsgBox "不存在第三方交易数据,请检查!", vbOKOnly + vbInformation, gstrSysName
                            Exit Function
                        End If
                        With mTyDelFee.rsBlance
                            If blnAll部份退费 And Val(Nvl(!是否全退)) = 1 And InStr(1, strNos, ",") > 0 Then
                                If Not mTyDelFee.blnSingleBalance Or Val(Nvl(!是否全退)) = 1 Then
                                    MsgBox "当前单据使用了第三方结算交易,所有单据必须全退！", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                            '不退现才处理
                            If zlCheckDelValied(Val(Nvl(!卡类别ID)), Nvl(!名称), Val(Nvl(!性质)) = 4, Nvl(!卡号), Nvl(!交易流水号), Nvl(!交易说明), Original.结帐ID, dblDelMoney) = False Then Exit Function
                        End With
                    End If
                Case 4   '4-结算卡
                    '如果支持退现,则直接退现
                    If Val(.RowData(i)) = -1 And Val(.TextMatrix(i, 1)) = 0 Then
                        strThreeBalance = strThreeBalance & "," & strTmp
                    Else
                        blnExistThreeSwap = True:
                    End If
                    If blnExistThreeSwap Then
                        If blnCur部份退费 Then
                            MsgBox "当前单据使用了第三方结算交易,不能进行部分退费！", vbInformation, gstrSysName
                            Call SetAllDelSelAll: Exit Function
                        End If
                        mTyDelFee.rsBlance.Filter = 0
                        mTyDelFee.rsBlance.Filter = "性质=" & Val(.Cell(flexcpData, i, 0)) & " And 结算方式='" & strTmp & "'"
                        If mTyDelFee.rsBlance.EOF Then
                            MsgBox "不存在第三方交易数据,请检查!", vbOKOnly + vbInformation, gstrSysName
                            Exit Function
                        End If
                        With mTyDelFee.rsBlance
                            If blnAll部份退费 And Val(Nvl(!是否全退)) = 1 And InStr(1, strNos, ",") > 0 Then
                                MsgBox "当前单据使用了第三方结算交易,所有单据必须全退！", vbInformation, gstrSysName
                                Exit Function
                            End If
                            '不退现才处理
                            If zlCheckDelValied(Val(Nvl(!卡类别ID)), Nvl(!名称), Val(Nvl(!性质)) = 4, Nvl(!卡号), Nvl(!交易流水号), Nvl(!交易说明), Original.结帐ID, Val(Nvl(!结算金额))) = False Then Exit Function
                        End With
                    End If
                Case 5 '一卡通
                    If Val(.RowData(i)) <> -1 And Val(.TextMatrix(i, 1)) = 0 Then
                        strThreeBalance = strThreeBalance & "," & strTmp
                    Else
                        blnExistOneCard = True
                    End If
                    
                    If blnExistOneCard Then
                        If blnCur部份退费 Then
                            MsgBox "当前单据使用了一卡通结算,不能进行部分退费！", vbInformation, gstrSysName
                            Call SetAllDelSelAll: Exit Function
                        End If
                        mTyDelFee.rsBlance.Filter = 0
                        mTyDelFee.rsBlance.Filter = "性质=5"
                        If mTyDelFee.rsBlance.RecordCount = 0 Then
                            MsgBox "未找到一卡通的结算数据,请检查!", vbInformation + vbOKOnly, gstrSysName
                            Exit Function
                        End If
                        '检查一卡通是否合法
                        On Error Resume Next
                        Set objICCard = CreateObject("zlICCard.clsICCard")
                        On Error GoTo 0
                        If objICCard Is Nothing Then
                            MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
                            Exit Function
                        End If
                        'gobjSquare.objSquareCard
                        'strCardNo = objICCard.Read_Card(Me)
                        '弹出刷卡界面
                        'zlBrushCard(frmMain As Object, _
                        'ByVal lngModule As Long, _
                        'ByVal rsClassMoney As ADODB.Recordset, _
                        'ByVal lngCardTypeID As Long, _
                        'ByVal bln消费卡 As Boolean, _
                        'ByVal strPatiName As String, ByVal strSex As String, _
                        'ByVal strOld As String, ByVal dbl金额 As Double, _
                        'Optional ByRef strCardNo As String, _
                        'Optional ByRef strPassWord As String) As Boolean
                        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, 0, False, _
                          mobjBill.姓名, mobjBill.性别, mobjBill.年龄, 0, strCardNo, "") = False Then Exit Function
                        If strCardNo = "" Then Exit Function
                        If strCardNo <> Nvl(mTyDelFee.rsBlance!卡号) Then
                            MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                Case Else
                        strOtherBalance = strOtherBalance & "," & strTmp
                End Select
            End If
        Next
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    If strThreeBalance <> "" Then strThreeBalance = Mid(strThreeBalance, 2)
    If strOtherBalance <> "" Then strOtherBalance = Mid(strOtherBalance, 2)
    
    '问题:37307
    If Val(txt预交冲款.Text) <> 0 And gbyt预存款退费验卡 <> 0 Then
            cur消费金额 = Val(txt预交冲款.Text)
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.病人ID, cur消费金额, mlngModul, 1, , , , , , (gbyt预存款退费验卡 = 2)) Then Exit Function
    End If
    If strReclaimInvoice <> "" Then
        '需要显示出本次需要回收的发票
        If InStr(1, mstrPrivs, "退费核收发票") > 0 Then
            If MsgBox("注意:" & vbCrLf & " 当前退费需要回收以下发票:" & vbCrLf & strReclaimInvoice, vbQuestion + vbDefaultButton1 + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        strInvoices = ""
    Else
        If blnAll部份退费 Then
            If InStr(1, mstrPrivs, "退费核收发票") > 0 Then
                If frmReInvoice.ShowMe(Me, mstrInNO, Val(txt合计.Text), Val(txt应缴.Text), strInvoices) = False Then Exit Function
            End If
        End If
        If Not blnAll部份退费 And strBalance = "" And strThreeBalance = "" Then strOtherBalance = ""
    End If
    
    '如果是医保,当前单据必是全退,如果有不支持退回的医保结算,则要处理误差
    '部份退费产生的误差金额:如果是第一次退费且全部退费,则不处理误差
    '60974
'    If mintInsure <> 0 Then
'        If strBalance <> "" Then
'            Call GetDelMoney(cur误差金额)
'        End If
'    ElseIf BillExistDelete(mstrInNO, 1) Or blnCur部份退费 Then
'        Call GetDelMoney(cur误差金额)
'    End If
    Call GetDelMoney(cur误差金额)
    If mintInsure <> 0 And MCPAR.医保接口打印票据 Then
        If zlGetInvoiceGroupUseID(lng领用ID) = False Then Exit Function
        strInvoice = GetNextBill(lng领用ID)
    End If
    dtDateDel = zlDatabase.Currentdate
    lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
    
    
    '问题:43403
    '产生SQL语句,如果不是部分退费，则要收回票据，是部分退费则在调用打印时收回并重打。
    strSQL = "zl_门诊收费记录_DELETE('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
        "'" & strBalance & "','" & strSQL & "','" & zlStr.NeedName(cbo结算方式.Text) & "'," & cur误差金额 & _
        ",To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(blnAll部份退费, "1", "0") & _
        IIf(Trim(txt退费摘要.Text) <> "", ",'" & Trim(txt退费摘要.Text) & "'", ",NULL") & ","
    '     校对标志_In: 0-不需要较对;1-需较对(不处理人员缴款余额,不回收票据)
    strSQL = strSQL & "1," & lng冲销ID & "," & lng冲销ID & ",'" & strThreeBalance & IIf(strThreeBalance <> "", ",", "") & strOtherBalance & "')"
    zlAddArray cllPro, strSQL
    
    '先产生票据，医保接口才能取到
    If mintInsure <> 0 And MCPAR.医保接口打印票据 And (gTy_Module_Para.byt票据分配规则 = 0 Or _
        gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice = "") Then
        '25187:在退费完成后,再进行收回重打(已经在frmPrint中处理了),所以不再处理:strReclaimInvoice=""
        strSQL = "zl_门诊收费记录_RePrint('" & mstrInNO & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
       zlAddArray cllPro, strSQL
    End If
    
'    '处理误差
'    '   部份退费的误差处理
'60974
'    If cur误差金额 <> 0 Then
'        strSql = "zl_门诊收费误差_Insert('" & mstrInNO & "'," & cur误差金额 & ",1)"
'       zlAddArray cllPro, strSql
'        vsBalance.ToolTipText = "医保结算方式" '恢复信息,之前记录了误差金额
'    End If
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strNos, 0) = False Then Exit Function
    End If
    cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保延时
    On Error GoTo errH
    bln退现 = False
    If cbo结算方式.ListIndex >= 0 Then
        bln退现 = cbo结算方式.ItemData(cbo结算方式.ListIndex) = 1
        str退结算方式 = zlStr.NeedName(cbo结算方式.Text)
    Else
        bln退现 = True
        If mrs结算方式 Is Nothing Then
            str退结算方式 = "现金"
        Else
            mrs结算方式.Filter = "性质=1"
            If mrs结算方式.EOF Then
                str退结算方式 = "现金"
            Else
                str退结算方式 = Nvl(mrs结算方式!名称, "现金")
            End If
        End If
    End If
    strErrMsg = "": dblErrMoney = 0
    '1.执行退费
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    '医保处理
    If mintInsure <> 0 And mblnYB结算作废 Then
        '医保退费处理
        If Not DelInsure(blnExistThreeSwap, mintInsure, str保险结算, Original.结帐ID, mstrInNO, _
            mobjBill.Pages(1).实收金额, Val(txt预交冲款.Text), str退结算方式, bln退现, dbl退金额) Then gcnOracle.RollbackTrans: Exit Function
        gcnOracle.CommitTrans: blnTrans = False: blnCommited = True
        
        '不存在第三方交易的，直接退出
        If Not blnExistThreeSwap Then
            If dbl退金额 <> 0 Then
                MsgBox "应退金额" & vbCrLf & str退结算方式 & "：" & Format(dbl退金额, "0.00") & "元", vbInformation, gstrSysName
            End If
            GoTo PrintBill:
        End If
        gcnOracle.BeginTrans: blnTrans = True
    End If
    
    '退一卡通
    blnCommited = False  '54949
    If DelOneCardSwap(objICCard, blnCommited) = False Then Exit Function
    
    '退第三方接口交易
    blnCommited = False '54949
    If DelTreeSwap(blnCommited, lng冲销ID) = False Then Exit Function
    gcnOracle.CommitTrans: blnTrans = False
    If strErrMsg <> "" Then
        MsgBox "第三方交易失败总额为:" & Format(dblErrMoney, "0.00") & " 请补调结算交易接口." & vbCrLf & strErrMsg, vbInformation, gstrSysName
        cmdOK.Enabled = True: Exit Function
    End If
    
PrintBill:
    If OverFeeDel(lng冲销ID, mobjBill.病人ID) = False Then Exit Function
    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo errHandle
    
    lng病人ID = mobjBill.病人ID
    Call PrintDelBill(strNos, dtDateDel, lng病人ID, blnAll部份退费, strInvoices, strReclaimInvoice)
    '56615
    Call WriteMzInforToCard(lng病人ID, lng冲销ID, True)
    cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
    DelChargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    cmdOK.Enabled = True
    Call SaveErrLog
End Function
Private Sub PrintDelBill(ByVal strNos As String, ByVal dtDateDel As Date, ByVal lng病人ID As Long, ByVal blnAll部份退费 As Boolean, _
    ByVal strInvoices As String, ByVal strReclaimInvoice As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印退费单据
    '入参:strNos-退费单据号
    '编制:刘兴洪
    '日期:2013-05-27 11:47:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, strSQL As String
    Dim strPriceGrade As String
    
    On Error GoTo errHandle
    '部分退费时收回并重打
    If Not blnAll部份退费 Then
         '税控部件全退时收回处理(全退时，zl_门诊收费记录_DELETE中已收回票据)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strNos)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        '打印回单
        GoTo PrintList:
        Exit Sub
    End If
    
    If mblnPrint Then
        If gintPriceGradeStartType >= 2 Then
            strPriceGrade = GetPriceGradeFromNos(strNos)
        Else
            strPriceGrade = mstr普通价格等级
        End If
    End If
    
    If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "" Then
        '按新票据分配规则打印
        '先预算,看票据是否充足
        Dim str发票号 As String, int票据张数 As Integer
        str发票号 = strReclaimInvoice
        If zlExeCuteBillNoSplit(True, 4, mlng领用ID, strNos, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数) = False Then GoTo PrintList:
        If int票据张数 = 0 Then
            '只回收票据,但不打印
            str发票号 = strReclaimInvoice
            Call zlExeCuteBillNoSplit(False, 4, mlng领用ID, strNos, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数)
            GoTo PrintList:
        End If
        
        mblnPrint = True
        ''0-不打印;1-自动打印;2-提示打印
        If mintInvoicePrint = 0 Then mblnPrint = False   '自动打印
        If mintInvoicePrint = 2 Then
            If MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then mblnPrint = False
        End If
        
        '重打收回发票
        If mblnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            Call RePrintCharge(1, strNos, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    If strInvoices = "" Then  '收回并重新打印门诊收据
       '0-不打印;1-自动打印;2-提示打印
        mblnPrint = True
        ''0-不打印;1-自动打印;2-提示打印
        If mintInvoicePrint = 0 Then mblnPrint = False   '自动打印
        If mintInvoicePrint = 2 Then
            If MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then mblnPrint = False
        End If
        If mblnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            Call RePrintCharge(1, strNos, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    'b.收费或上一次退时没有打印票据
    If strInvoices <> "无可退票据" Then
        'c.只收回票据
        strSQL = "zl_门诊收费记录_RePrint('" & mstrInNO & "',Null,0,'" & UserInfo.姓名 & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
 
PrintList:
    '打印费用清单
    If blnAll部份退费 Then
        If InStr(mstrPrivs, "打印清单") > 0 Then
            If gint收费清单 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            ElseIf gint收费清单 = 2 Then
                If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
                End If
            End If
        End If
    End If
    If mintInsure <> 0 And MCPAR.退费后打印回单 And InStr(1, mstrPrivs, "医保退费回单") > 0 Then
        '问题:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strNos, 2)
    End If
    If mint退费回单打印 = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strNos, 2)
    ElseIf mint退费回单打印 = 2 Then
        If MsgBox("是否打印退费回单？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strNos, 2)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DelOneCardSwap(ByRef objICCard As Object, blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通调用成功
    '入参:
    '出参:blnCommited-提交过一次成功的,返回true
    '返回:调用成功, 返回true,否则返回False(False时,被回退了的)
    '编制:刘兴洪
    '日期:2011-08-30 12:22:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, dblErrMoney As Double
    Dim strSuccess As String
    
    On Error GoTo errHandle
    With mTyDelFee.rsBlance
        .Filter = "性质=5 "
        '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
        If .RecordCount = 0 Then DelOneCardSwap = True: Exit Function
        Do While Not .EOF
            If Not DelOneCardMoney(objICCard, Nvl(!NO), Nvl(!卡号), Nvl(!交易流水号), Nvl(!医院编码), Val(Nvl(!结算金额))) Then
                gcnOracle.RollbackTrans
                If blnCommited Then
                    dblErrMoney = dblErrMoney + Val(Nvl(!结算金额))
                    strErrMsg = strErrMsg & vbCrLf & "    " & Nvl(!名称, "一卡通") & ":" & Val(Nvl(!结算金额))
                Else
                    MsgBox "一卡通退费交易调用失败,退费操作失败！", vbExclamation, gstrSysName
                    cmdOK.Enabled = True: Exit Function
                End If
             Else
                gcnOracle.CommitTrans: gcnOracle.BeginTrans: blnCommited = True
                strSuccess = strSuccess & vbCrLf & "    " & Nvl(!名称, "一卡通") & ":" & Val(Nvl(!结算金额))
             End If
        Loop
     End With
    If strErrMsg <> "" Then '54949
       gcnOracle.RollbackTrans
        MsgBox "    向一卡通退费时失败(总额为:" & Format(dblErrMoney, "0.00") & "), " & vbCrLf & _
                      "请重新对异常单据退费,失败接口如下:" & vbCrLf & _
                      strErrMsg & vbCrLf & _
                      "   向一卡通退费时,成功的交易如下:" & vbCrLf & _
                      strSuccess, vbExclamation, gstrSysName
        cmdOK.Enabled = True: Exit Function
    End If
    DelOneCardSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    SaveErrLog
End Function
Private Function DelTreeSwap(ByRef blnCommited As Boolean, ByVal lng冲销ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退第三方交易
    '入参;blnCommited-是否已经被提交过
    '返回:退成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-30 11:36:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance  As New Collection, cllUpdate As New Collection, cllThreeSwap As New Collection
    Dim lngRow As Long, str结算方式 As String, dblMoney As Double, strSQL As String
    Dim blnYes As Boolean   '是否调用接口
    Dim strErrMsg As String, dblErrMoney As Double
    Dim strSucces As String
    
    On Error GoTo errHandle
    With mTyDelFee.rsBlance
        .Filter = "性质=3 or 性质=4"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
           '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
           blnYes = True
           str结算方式 = Nvl(!结算方式)
'            If Val(Nvl(!是否退现)) = 1 Then
                blnYes = False
                With vsBalance
                    For lngRow = 0 To .Rows - 1
                        If str结算方式 = .TextMatrix(lngRow, 0) And Val(.TextMatrix(lngRow, 1)) <> 0 Then
                            '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
                            Select Case Val(.Cell(flexcpData, lngRow, 0))
                            Case 3, 4
'                                If .RowData(lngRow) = -1 And .RowHidden(lngRow) = False Then
                                     blnYes = True
                                     dblMoney = Val(.TextMatrix(lngRow, 1))
                                     Exit For
'                                End If
                            End Select
                        End If
                    Next
                End With
'            End If
            If blnYes Then
                Set cllBalance = New Collection
                Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
                'cllBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
                '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
                cllBalance.Add Array(Val(Nvl(!卡类别ID)), Nvl(!卡号), IIf(Val(Nvl(!性质)) = 4, 1, 0), Nvl(!交易流水号), Nvl(!交易说明)), mstrInNO
                
                If CallBackBalanceInterface(cllBalance, Original.结帐ID, lng冲销ID, mstrInNO, dblMoney, cllUpdate, cllThreeSwap, strErrMsg) = False Then
                    gcnOracle.RollbackTrans
                    If blnCommited Then
                        dblErrMoney = dblErrMoney + dblMoney
                        strErrMsg = strErrMsg & vbCrLf & "    " & Nvl(!名称) & ":" & dblMoney
                    Else
                        MsgBox "调用三方交易接口时,退费交易调用失败!" & vbCrLf & "    " & Nvl(!名称) & ":" & Format(dblMoney, "0.00"), vbExclamation, gstrSysName
                        cmdOK.Enabled = True: Exit Function
                    End If
                Else
                        '更新数据
                        'Zl_门诊收费_完成校对
                        strSQL = "Zl_门诊收费_完成校对("
                        '  No_In       门诊费用记录.NO%Type,
                        strSQL = strSQL & "'" & mstrInNO & "',"
                        '  操作类型_In Number, 0-一卡通;1-消费卡;2-医疗卡
                        strSQL = strSQL & "" & IIf(Val(Nvl(!性质)) = 4, 1, 2) & ","
                        '  卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & "" & Val(Nvl(!卡类别ID)) & ","
                        '  卡号_In     病人预交记录.卡号%Type
                        strSQL = strSQL & "'" & Nvl(!卡号) & "')"
                        zlDatabase.ExecuteProcedure strSQL, Me.Caption
                        'cllUpdate:已经在Delete中执行,不能再更新
                         gcnOracle.CommitTrans
                        Call SaveThreeData(cllThreeSwap)
                         gcnOracle.BeginTrans: blnCommited = True
                         strSucces = strSucces & vbCrLf & "    " & Nvl(!名称) & ":" & dblMoney
                         '调用相关的退费检查
                End If
            End If
            .MoveNext
        Loop
    End With
    If strErrMsg <> "" Then '54949
       gcnOracle.RollbackTrans
        MsgBox "    向三方交易退费时失败(总额为:" & Format(dblErrMoney, "0.00") & "), " & vbCrLf & _
                      "请重新对异常单据退费,失败接口如下:" & vbCrLf & _
                      strErrMsg & vbCrLf & _
                      "   向三方交易退费时,成功的交易如下:" & vbCrLf & _
                      strSucces, vbExclamation, gstrSysName
        cmdOK.Enabled = True: Exit Function
    End If
    DelTreeSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function
Private Function OverFeeDel(ByVal str冲销IDs As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:完成退费收费
    '入参:strNos-完成收费的单据(可以为多张,但目前只有一张单据)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    If Left(str冲销IDs, 1) = "," Then str冲销IDs = Mid(str冲销IDs, 2)
    ' Zl_门诊收费结算_完成退费
    strSQL = "Zl_门诊收费结算_完成退费("
    '  病人id_In       门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  退费结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "NULL,"
    '  冲销ids_In      Varchar2,
    strSQL = strSQL & "'" & str冲销IDs & "',"
    '  操作员姓名_In   病人预交记录.操作员姓名%Type := Null
    strSQL = strSQL & "'" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    OverFeeDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function DelOneCardMoney(ByVal objICCard As Object, ByVal strNo As String, ByVal strCardNo As String, _
    ByVal strSwapNO As String, ByVal strHsptName As String, _
    ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通退费
    '入参:strCardNo-卡号
    '        strSwapNo-交易流水号
    '        strHsptName-医院编码
    '        dblMoney-退费金额
    '出参:
    '返回:交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-29 00:07:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    
     If Not objICCard.ReturnSwap(strCardNo, strHsptName, strSwapNO, dblMoney) Then Exit Function
    '更新一卡通的校对标志
    'Zl_门诊收费_完成较对
    strSQL = "Zl_门诊收费_完成校对("
    '  No_In       门诊费用记录.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  操作类型_In Number, 0-一卡通;1-消费卡;2-医疗卡
    strSQL = strSQL & "1,"
    '  卡类别id_In 病人预交记录.卡类别id%Type,
    strSQL = strSQL & "NULL,"
    '  卡号_In     病人预交记录.卡号%Type
    strSQL = strSQL & "'" & strCardNo & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    DelOneCardMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function DelInsure(ByVal blnExistThreeSwap As Boolean, _
    ByVal intInsure As Integer, ByVal str医保结算 As String, _
    ByVal lng结帐ID As Long, ByVal strNo As String, _
    ByVal dbl实收金额 As Double, ByVal dbl本次冲预交 As Double, _
    ByVal str退结算方式 As String, ByVal bln退现 As Boolean, ByRef cur金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据号调用接口
    '入参:str退结算方式-退回的指定结算方式
    '       bln退现-是否退回现金
    '       blnExistThreeSwap-是否存在第三方卡结算
    '       str医保结算-医保结算,以逗号分离(个人帐户,医保帐户)
    '出参:cur金额－退费金额
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 12:21:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur余额 As Currency, strSQL As String, strAdvance As String
    Dim lng冲销ID As Long, i As Long, cur误差金额 As Double, str收费结算 As String, blnTransMedicare As Boolean
    Err = 0: On Error GoTo Errhand:
    If blnExistThreeSwap Then
        ' Zl_门诊结算_较对标志_Update
        strSQL = "Zl_门诊结算_较对标志_Update("
        '  结帐id_In     门诊费用记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算序号id_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "NULL,"
        '  收费结算_In   Varchar2,
        strSQL = strSQL & "'" & str医保结算 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  消费卡_In     Integer := 0,
        strSQL = strSQL & "0,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  校对标志_In   病人预交记录.校对标志%Type := 0
        strSQL = strSQL & "2)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    strAdvance = "1|1"
    If Not gclsInsure.ClinicDelSwap(lng结帐ID, , intInsure, strAdvance) Then Exit Function
    blnTransMedicare = True
    If strAdvance = "1|1" Or strAdvance = "" Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
        DelInsure = True: Exit Function
     End If
    '根据返回的结算信息，修正预交记录
    '   strAdvance返回格式:结算方式1|金额||结算方式2:金额...
    cur余额 = 0
    For i = 0 To UBound(Split(strAdvance, "||"))
        cur余额 = cur余额 + -1 * Split(Split(strAdvance, "||")(i), "|")(1)
    Next
    cur余额 = dbl实收金额 - cur余额 - Val(txt预交冲款.Text)
    cur金额 = cur余额: cur误差金额 = 0
    '退为指定的结算方式，如果是现金，可能产生新的误差金额
    If bln退现 Then
        cur金额 = Format(CentMoney(cur余额), "0.00")
        cur误差金额 = cur金额 - cur余额
    End If
    str收费结算 = str退结算方式 & "|" & -1 * cur金额 & "| "
    lng冲销ID = GetDelBalanceID(strNo)
    If Not blnExistThreeSwap Then
        strSQL = "zl_门诊收费结算_Update(" & lng冲销ID & ",'" & str收费结算 & "'," & -1 * dbl本次冲预交 & ",'" & _
            strAdvance & "'," & -1 * cur误差金额 & ",NULL,NULL,NULL,1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
        DelInsure = True: Exit Function
    End If
    'Zl_医保结算校对_Update
    strSQL = "Zl_医保结算校对_Update("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  保险结算_In Varchar2
    strSQL = strSQL & "'" & strAdvance & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
    
    DelInsure = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
End Function

Private Function DelCharge() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费或销帐
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 09:35:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, i As Long
    Dim cllPro As Collection
     
    If mbytInFun = 2 Then '销帐
        If WriteOff = False Then Exit Function
    ElseIf mbytInFun = 0 Then '退费
        If DelChargeFee = False Then Exit Function
    End If
    '完成后的界面处理
    DelCharge = True
    If mbytInState <> 0 Then Unload Me: Exit Function
    If mbytInFun = 0 And gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    mstrInNO = "": cboNO.Text = "": txtInvoice.Text = ""
    Call ClearBillRows: Call ClearMoney
    chkCancel.Value = 0
    Call ClearPatientInfo(True)
    Call ClearTotalInfo
    Call SetDisible(True)
    Call NewBill(IIf(mblnStartFactUseType, False, True))
    If mbytBilling = 2 Then
        cboNO.SetFocus
    Else
        txtPatient.SetFocus
    End If
    mblnSaveData = True
End Function
Private Function SaveVerify() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐划价单审核操作
    '返回:审核成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 10:17:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long
    
    On Error GoTo errHandle
    If Not (mbytInFun = 2 And mbytBilling = 2) Then Exit Function
    If mstrInNO = "" Then
        MsgBox "没有记帐划价单据,请先输入！", vbInformation, gstrSysName
        cboNO.SetFocus: Exit Function
    End If
    '取本次审核的行序号
    strSQL = ""
    For i = 1 To Bill.Rows - 1
        If Bill.RowData(i) > 0 Then
            strSQL = strSQL & "," & Bill.RowData(i)
        End If
    Next
    strSQL = Mid(strSQL, 2)
    i = GetBillRows(mstrInNO, 2)
    If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
    If Val(txt合计.Text) <> 0 And gdbl预存款消费验卡 <> 0 Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.病人ID, Val(txt合计.Text), mlngModul, 1, , IIf(-1 * gdbl预存款消费验卡 >= Val(txt合计.Text), False, True), , , , (gdbl预存款消费验卡 = 2)) Then Exit Function
    End If
    '费用报警
    If Not AuditingWarn(mstrPrivs, mstrInNO, strSQL) Then Exit Function
    strSQL = "zl_门诊记帐记录_Verify('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & strSQL & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '发送消息
    Call SendMsgModule
    SaveVerify = True
    If gbln审核打印 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & mstrInNO, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
    End If
    
    '110319
    If mblnDrugMachine Then
        '门诊格式：1|单据1,处方号1;单据2,处方号2
        Dim strData As String, strReturn As String
        strData = "1|" & "9," & mstrInNO
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    End If
    
    mstrInNO = "": cboNO.Text = ""
    Call ClearPatientInfo(True)
    Call ClearTotalInfo
    Call ClearBillRows: Call ClearMoney
    Call NewBill: Call SetMoneyList
    cboNO.Locked = False: cboNO.SetFocus
    mblnSaveData = True
    SaveVerify = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isValiedCargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据的合法性
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 14:05:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Long, j As Long, strTmp As String
    Dim p As Integer, dblNum As Double, strInfo As String
    Dim cur合计 As Currency, cur当日额 As Currency, cur余额 As Currency
    Dim blnMerge As Boolean, k As Integer, bln检查库存 As Boolean, colStock As Collection
    Dim dblToTal As Double, lng药房ID As Long
    Dim blnExistValidItem As Boolean
    
    On Error GoTo errHandle
    If mbytInFun = 2 Then
        If mrsInfo.State = adStateClosed Then
            MsgBox "没有发现" & gstrCustomerAppellation & "信息,请确定" & gstrCustomerAppellation & "信息！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Function
        End If
    ElseIf mbytInFun = 0 Then '收费和记帐必须输入姓名
        If txtPatient.Text = "" Then
            MsgBox "没有发现" & gstrCustomerAppellation & "信息,请输入" & gstrCustomerAppellation & "信息！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Function
        ElseIf mobjBill.姓名 = "" Then
            mobjBill.姓名 = txtPatient.Text
        End If
    End If
    If mbytInFun = 1 And gint病人来源 = 2 And Trim(txtPatient.Text) = "" Then
         MsgBox "病人来源为住院病人时，必须输入病人信息！", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Function
    End If
    If CheckTextLength("姓名", txtPatient) = False Then Exit Function
    If CheckTextLength("年龄", txt年龄) = False Then Exit Function
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Function
    '刘兴洪 问题:?? 日期:2010-01-07 11:26:40
    '北京医保检查
    If mobjBill.病人ID <> 0 And mbytInFun = 1 And mbytInState = 0 Then
        gstrSQL = "Select 险类 from 病人信息 where 病人id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.病人ID)
        If Not rsTmp.EOF Then
            MCPAR.医生确定处方类型 = gclsInsure.GetCapability(support医生确定处方类型, mobjBill.病人ID, Val(Nvl(rsTmp!险类)))
            If zlCheck北京医保(Val(Nvl(rsTmp!险类))) = False Then Exit Function
        End If
    End If
    If mobjBill.费别 = "" Then
        MsgBox "请选择" & gstrCustomerAppellation & "费别！", vbInformation, gstrSysName
        If cbo费别.Visible And cbo费别.Enabled Then cbo费别.SetFocus
        Exit Function
    End If

    If CheckBillsEmpty Then
        MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    ElseIf mobjBill.Pages.Count > 1 Then
        For i = 1 To mobjBill.Pages.Count
            If CheckBillsEmpty(i) Then
                MsgBox "第 " & i & " 张单据没有输入任何内容！", vbInformation, gstrSysName
                tbsBill.Tabs(i).Selected = True
                Bill.SetFocus: Exit Function
            End If
        Next
    End If
    '是否全部输入了执行科室
    i = CheckExecuteDept(j)
    If i > 0 And j > 0 Then
        If mobjBill.Pages.Count > 1 Then
            MsgBox "第 " & j & " 张单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
            tbsBill.Tabs(j).Selected = True
        Else
            MsgBox "单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
        End If
        Bill.SetFocus: Exit Function
    End If
    If Not glngSys Like "8??" Then
        For i = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(i).开单部门ID = 0 Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "第 " & i & " 张单据没有指定开单科室！", vbInformation, gstrSysName
                    tbsBill.Tabs(i).Selected = True
                Else
                    MsgBox "没有指定开单科室！", vbInformation, gstrSysName
                End If
                If gbyt科室医生 = 0 Then
                    cbo开单人.SetFocus
                Else
                    cbo开单科室.SetFocus
                End If
                Exit Function
            End If
        Next
    End If
    
    '开单人
    If mbytInFun <> 2 And gbln必须输开单人 Then
        For i = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(i).开单人 = "" Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "第 " & i & " 张单据没有指定开单人！", vbInformation, gstrSysName
                    tbsBill.Tabs(i).Selected = True
                Else
                    MsgBox "没有指定开单人！", vbInformation, gstrSysName
                End If
                cbo开单人.SetFocus: Exit Function
            End If
        Next
    End If
    '检查开单人与开单科室对应关系
    If mbytInFun <> 2 And mbytInState = 0 And (gbyt科室医生 = 0 Or gbyt科室医生 = 1) Then
        If Not (cbo开单人.Locked And cbo开单科室.Locked) Then
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).开单人 <> "" And mobjBill.Pages(i).NO = "" Then        '25618:And mobjBill.Pages(i).NO = "":刘兴洪加入,主要是挂号产生的划价单时,不一定开单科室时临床的,因此不能检查
                    mrs开单人.Filter = "姓名='" & mobjBill.Pages(i).开单人 & "' And 部门ID=" & mobjBill.Pages(i).开单部门ID
                    If mrs开单人.RecordCount = 0 Then
                        MsgBox "开单人""" & mobjBill.Pages(i).开单人 & """不属于开单科室""" & zlStr.NeedName(cbo开单科室.Text) & """,请检查！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
        End If
    End If
    '检查病人挂号科室,是否允许对没有挂号的病人收费
    If mbytInFun = 0 And gblnCheckRegeventDept And gint病人来源 = 1 _
        And (gTy_System_Para.Sy_Reg.bytNODaysGeneral > 0 Or gTy_System_Para.Sy_Reg.bytNoDayseMergency > 0) And mobjBill.病人ID > 0 Then
        Set rsTmp = GetDeptByRegevent(mobjBill.病人ID)
        If rsTmp.RecordCount > 0 Then '未挂号的根据本地参数已检查
            For i = 1 To mobjBill.Pages.Count
                If Not CheckDeptIsMedTech(mobjBill.Pages(i).开单部门ID) Then
                    rsTmp.Filter = "执行部门ID=" & mobjBill.Pages(i).开单部门ID
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "当前病人没有在第" & i & "张单据的开单科室挂过号,不允许收费!", vbInformation, gstrSysName
                        tbsBill.Tabs(i).Selected = True
                        Exit Function
                    End If
                End If
            Next
        End If
    End If
    '护士类别:判断非法输入
    For i = 1 To mobjBill.Pages.Count
        If CheckInhibitiveByNurse(i) Then
            If mobjBill.Pages.Count > 1 Then
                MsgBox "护士只能输入治疗及材料项目,而第 " & i & " 张单据中存在其它类型的项目。", vbInformation, gstrSysName
                If tbsBill.SelectedItem.Index <> i Then
                    tbsBill.Tabs(i).Selected = True
                End If
            Else
                MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
            End If
            Bill.SetFocus: Exit Function
        End If
    Next
 

    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Function
    End If
    
    If mbln补费 Then
        If txtDate.Text > mstr最后转科时间 And mstr最后转科时间 <> "" Then
            MsgBox "该病人补录的费用时间超过了最后转出的时间(" & mstr最后转科时间 & ")，不能进行补费操作！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Function
        End If
        If cbo开单科室.ItemData(cbo开单科室.ListIndex) <> mlngDeptID And mlngDeptID <> 0 Then
            MsgBox "开单科室不是病人转科的科室，不能进行补费操作！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
 
    '导入划价单收费时,如果是医嘱生成的,可能已作废
    For i = 1 To mobjBill.Pages.Count
        '针对每张单据判断(因为可能划价和收费混用),是否是导入医嘱生成的划价单收费
        If mobjBill.Pages(i).NO <> "" And mobjBill.Pages(i).医嘱序号 <> 0 Then
            If mobjBill.Pages(i).实收金额 <> GetBillSumByDB(mobjBill.Pages(i).NO) Then
                MsgBox "单据[" & mobjBill.Pages(i).NO & "]的部分收费记录已被他人修改或作废,请重新读取单据后再收费！", vbInformation, gstrSysName
                tbsBill.Tabs(i).Selected = True
                Exit Function
            End If
        End If
    Next
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If CheckServeRange(0, .收费细目ID) = False Then Exit Function
            End With
        Next i
    Next p

    '输入了无效项目的行
    strTmp = ""
    For p = 1 To mobjBill.Pages.Count
        blnExistValidItem = False
        For i = 1 To mobjBill.Pages(p).Details.Count
            '27467,106490
            If mobjBill.Pages(p).Details(i).数次 <> 0 Then blnExistValidItem = True
            If mobjBill.Pages(p).Details(i).收费细目ID = 0 Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "第 " & p & " 张单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                    tbsBill.Tabs(p).Selected = True
                Else
                    MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                End If
                Bill.SetFocus: Exit Function
            ElseIf InStr(1, ",5,6,7,", mobjBill.Pages(p).Details(i).收费类别) > 0 Then
                '收集药品的发药药房对应的服务科室
                strTmp = strTmp & "," & mobjBill.Pages(p).Details(i).收费细目ID
            End If
            
            If mbytInFun = 2 Then
                '负数冲销数量检查:只有门诊留观病人才会进行检查是否充足
                '问题:36558
                 If Not mrsInfo Is Nothing Then
                    If mrsInfo.State = 1 Then
                        If Nvl(mrsInfo!留观, 0) = 1 Then
                            If InStr(",5,6,7,", mobjBill.Pages(p).Details(i).收费类别) > 0 And gbln药房单位 Then
                                 dblNum = mobjBill.Pages(p).Details(i).数次 * mobjBill.Pages(p).Details(i).付数 * mobjBill.Pages(p).Details(i).Detail.药房包装
                             Else
                                 dblNum = mobjBill.Pages(p).Details(i).数次 * mobjBill.Pages(p).Details(i).付数
                             End If
                             If dblNum < 0 Then
                                If Not CheckNegative(mobjBill.病人ID, mobjBill.主页ID, mobjBill.Pages(p).Details(i).收费细目ID, mobjBill.Pages(p).Details(i).执行部门ID, dblNum, mobjBill.Pages(p).Details(i).Detail.药房包装, mstrPrivs, Format(mrsInfo!入院日期, "yyyy-mm-dd HH:MM:SS")) Then
                                    tbsBill.Tabs(p).Selected = True
                                    Bill.SetFocus: Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
            
        '问题:41668,106490
        If mobjBill.Pages(p).NO = "" And mbytInState = 0 And blnExistValidItem = False Then
            If mobjBill.Pages.Count > 1 Then
                MsgBox "第 " & p & " 张单据中至少要有一条数次不为零的项目，请检查！", vbInformation, gstrSysName
                tbsBill.Tabs(p).Selected = True
            Else
                MsgBox "单据中至少要有一条数次不为零的项目，请检查！", vbInformation, gstrSysName
            End If
            Bill.SetFocus: Exit Function
        End If
    Next
            
    '检查药品的发药药房对应的服务科室
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
        Set rsTmp = GetServiceDept(strTmp)
        If Not rsTmp Is Nothing Then
            For p = 1 To mobjBill.Pages.Count
                strTmp = ""
                For i = 1 To mobjBill.Pages(p).Details.Count
                    If InStr(1, ",5,6,7,", mobjBill.Pages(p).Details(i).收费类别) > 0 Then
                        strInfo = mobjBill.Pages(p).Details(i).收费细目ID
                        '先检查是否是允许的存储库房
                        rsTmp.Filter = "收费细目ID=" & strInfo & " And 执行科室id=" & mobjBill.Pages(p).Details(i).执行部门ID
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        Else
                            '再检查是否是允许的服务科室(没有设置服务科室的,开单科室ID为零)
                            rsTmp.Filter = "(" & rsTmp.Filter & " And 开单科室ID=" & _
                                IIf(mobjBill.科室ID = 0, mobjBill.Pages(p).开单部门ID, mobjBill.科室ID) & ") Or (" & rsTmp.Filter & " And 开单科室ID=0)"
                            If rsTmp.RecordCount = 0 Then
                                strTmp = strTmp & "," & i
                            End If
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                    MsgBox "请检查,第" & p & "张单据,第" & strTmp & "行药品是否违反以下规则:" & vbCrLf & vbCrLf & _
                        "A.选择的执行科室不是药品的存储库房" & vbCrLf & _
                        "B.病人科室[" & GET部门名称(IIf(mobjBill.科室ID = 0, mobjBill.Pages(p).开单部门ID, mobjBill.科室ID), mrs开单科室) & "]不属于药品在此存储库房的服务科室.", _
                        vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        End If
    End If

    '处方职务检查
    '1.公费或医保病人
    If cbo医疗付款.ListIndex <> -1 And mbln处方职务检查 Then
        '医保或公费病人
        '问题:45605
        If zlIsCheckMedicinePayMode(zlStr.NeedName(cbo医疗付款)) Then
            i = CheckDuty(, False, j)
            If i > 0 And j > 0 Then
                If mobjBill.Pages.Count > 1 Then tbsBill.Tabs(j).Selected = True
                Bill.Row = i: Bill.MsfObj.TopRow = i
                Bill.Col = BillCol.项目: Bill.SetFocus: Exit Function
            End If
        End If
    End If
    '2.所有病人项目
    If mbln处方职务检查 Then
        i = CheckDuty(, True, j)
        If i > 0 And j > 0 Then
            If mobjBill.Pages.Count > 1 Then tbsBill.Tabs(j).Selected = True
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.项目: Bill.SetFocus: Exit Function
        End If
    End If
    
    '费用类型检查
    If Not CheckFeeType Then Exit Function
    
    '记帐分类报警
    If mbytInFun = 2 And Not mrsWarn Is Nothing Then
        '单据费用
        cur合计 = GetBillSum
        If cur合计 > 0 Then
           Call LoadFeeInfor(mrsInfo!病人ID)
           
            '重新读取当日额
            cur当日额 = GetPatiDayMoney(mrsInfo!病人ID)
                           
            cur余额 = Val(cmdPrint.Tag)
            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(0, mrsInfo!病人ID) + IIf(mbytBilling = 1, Original.实收合计, 0)
                    
            For i = 1 To mobjBill.Pages.Count
                For j = 1 To mobjBill.Pages(i).Details.Count
                    gbytWarn = BillingWarn(mstrPrivs, mrsInfo!姓名, mrsInfo!适用病人, mrsWarn, cur余额, cur当日额 - Original.实收合计, cur合计, _
                        IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Pages(i).Details(j).收费类别, _
                        mobjBill.Pages(i).Details(j).Detail.类别名称, mstrWarn)
                    If gbytWarn = 2 Or gbytWarn = 3 Then Exit Function
                Next
            Next
        End If
    End If
    '药品禁忌检查
    strInfo = CheckDisable(mobjBill)
    If strInfo <> "" Then
        If strInfo Like "*(互相禁用)*" Then
            MsgBox strInfo, vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
                
    '处方限量检查
    If mbln处方限量检查 Then
        If Not gbln处方限量 Then
            If Not CheckLimit(mobjBill) Then Exit Function
        End If
    End If
    
    '单张单据最高额
    If gcurMax <> 0 And (mbytInFun = 0 Or mbytInFun = 1) Then
        For i = 1 To mobjBill.Pages.Count
            If GetBillSum(, i) > gcurMax Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "第 " & i & " 张单据金额超过最大限制金额:" & Format(gcurMax, "0.00") & " ,不允许保存！", vbInformation, gstrSysName: Exit Function
                Else
                    MsgBox "单据金额超过最大限制金额:" & Format(gcurMax, "0.00") & " ,不允许保存！", vbInformation, gstrSysName: Exit Function
                End If
            End If
        Next
    End If
    '检查分批或时价药品同一药房是否有重复输入
    blnMerge = False
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If (.Detail.分批 Or .Detail.变价) And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                    For k = 1 To mobjBill.Pages.Count
                        For j = 1 To mobjBill.Pages(k).Details.Count
                            If Not (p = k And i = j) And .收费细目ID = mobjBill.Pages(k).Details(j).收费细目ID And .执行部门ID = mobjBill.Pages(k).Details(j).执行部门ID Then
                                '多张单据的情况
                                If mobjBill.Pages.Count > 1 Then
                                    '非时价的分批药品，在不同的单据上有相同的，允许不合并，不提醒
                                    If .Detail.变价 Or (Not .Detail.变价 And .Detail.分批 And p = k) Then
                                        If .收费类别 = "4" Then
                                            If Not blnMerge Then
                                                If .Detail.批次 = mobjBill.Pages(k).Details(j).Detail.批次 Then
                                                    If MsgBox("第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                        vbCrLf & "分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入。" & _
                                                        vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？", _
                                                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                        blnMerge = True     '不应退出循环，因为还要检查是否有不同付数的中草药,如果有的话，不能自动合并
                                                    Else
                                                        tbsBill.Tabs(k).Selected = True: Exit Function
                                                    End If
                                                End If
                                            End If
                                        Else
                                            '两个不同单据中的中药付数不同时，应是不同的配方，无法自动合并
                                            If .收费类别 = "7" And .付数 <> mobjBill.Pages(k).Details(j).付数 Then
                                                MsgBox "第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                    vbCrLf & "分批或时价中草药""" & .Detail.名称 & """(不同付数)在同一个药房被重复输入。", vbInformation, gstrSysName
                                                tbsBill.Tabs(k).Selected = True: Exit Function
                                            ElseIf Not blnMerge Then
                                                If MsgBox("第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                    vbCrLf & "分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？", _
                                                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    blnMerge = True
                                                Else
                                                    tbsBill.Tabs(k).Selected = True: Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                '单张单据的情况
                                ElseIf Not blnMerge Then
                                    If .收费类别 = "4" Then
                                        If .Detail.批次 = mobjBill.Pages(k).Details(j).Detail.批次 Then
                                            strInfo = "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入。" & _
                                                        vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？"
                                        End If
                                    Else
                                        strInfo = "第 " & j & " 行的分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？"
                                    End If
                                    
                                    If strInfo <> "" Then
                                        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                            blnMerge = True     '可以退出循环
                                        Else
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End With
        Next
    Next
    '自动合并,只要有合并,都应重新保存,因为如果是收费并打印,行数的变化,可能影响工本费的张数
    If blnMerge Then
        Call MergeRepeatItem
        MsgBox "自动合并已完成，合并后费用金额或行数已发生变化，请检查后保存。", vbInformation, gstrSysName
        Exit Function
    End If
   '药品库存检查(仅不足禁止时或分批时价药品)
    bln检查库存 = (InStr(mstrPrivs, "不检查库存") = 0)    '是否有权限不检查库存(分批和时价必须检查)
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
            
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & _
                                dblToTal & """。", vbInformation, gstrSysName
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行药品""" & .Detail.名称 & _
                                    """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & _
                                    dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                                tbsBill.Tabs(p).Selected = True
                                Bill.SetFocus: Exit Function
                            End If
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblToTal & """。", vbInformation, gstrSysName
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                    """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                                tbsBill.Tabs(p).Selected = True
                                Bill.SetFocus: Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next
    '检查卫生材料的灭菌效期
    For i = 1 To mobjBill.Pages.Count
        For j = 1 To mobjBill.Pages(i).Details.Count
            With mobjBill.Pages(i).Details(j)
                If .收费类别 = "4" And .Detail.跟踪在用 Then
                    dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID, i)
                    If Not CheckValidity(.收费细目ID, .执行部门ID, dblToTal) Then Exit Function
                End If
            End With
        Next
    Next
    '发药窗口检查(仅划价单)
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO <> "" And tbsBill.Tabs(i).Tag = "" Then
            lng药房ID = BillExistDrug(mobjBill.Pages(i).NO, 1)
            If lng药房ID <> 0 Then
                If ExistWindow(lng药房ID, mrs发药窗口) Then
                    MsgBox "无法分配" & GET部门名称(lng药房ID, mrsUnit) & "的发药窗口，请确定是否正常安排窗口上班。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    '零差价检查,105872
    If Not gobjPublicDrug Is Nothing Then
        'Private Function zlCheckPriceAdjustBySell(ByVal lng药品id As Long, ByVal lng药房id As Long) As Boolean
        '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
        '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
        '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
        '销售出库时只判断药房
        '返回：True-正常进行销售出库；false-不能进行销售出库
        For p = 1 To mobjBill.Pages.Count
            For i = 1 To mobjBill.Pages(p).Details.Count
                With mobjBill.Pages(p).Details(i)
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        If gobjPublicDrug.zlCheckPriceAdjustBySell(.收费细目ID, .执行部门ID) = False Then
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    End If
                End With
            Next
        Next
    End If
    
    If mstrInNO <> "" Then
        If HaveExecute(1, mstrInNO, IIf(mbytInFun = 2, 2, 1)) Then
            MsgBox "该单据包含完全执行或部分执行的项目,不允许修改。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '刘兴洪:检查是否只有附加手术,如果只有附加手术,直接退出:
    '22441
    If CheckMainOperation = False Then Exit Function
    
    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 1, _
        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)))) = False Then
        Exit Function
    End If
    
    isValiedCargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckServeRange(intType As Integer, lng收费细目ID As Long, Optional intRow As Integer = 0) As Boolean
'功能:检查收费项目的服务对象,intType:0-门诊调用;1-住院调用
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 名称,Nvl(服务对象,0) As 服务对象 From 收费项目目录 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckServeRange", lng收费细目ID)
    If rsTmp.EOF Then
        MsgBox "不能确定" & IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目的服务对象,请检查项目是否正确录入!"
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!服务对象) = 2 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于门诊,请检查!"
                Exit Function
            End If
        Case 1
            If Val(rsTmp!服务对象) = 1 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于住院,请检查!"
                Exit Function
            End If
        Case Else
            If Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于病人,请检查!"
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function

Private Function CheckBillNOAndBookeFee(Optional blnReCharge As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据号码检查,工本费打印检查
    '入参:blnReCharge-是否重新收费的检查
    '返回:数据合法,返回tru,否则返回false
    '编制:刘兴洪
    '日期:2011-08-16 14:25:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl金额 As Double, j As Long, p As Long, i As Long
    
    On Error GoTo errHandle
    '票据号码检查,工本费打印检查
    If Not blnReCharge Then
        If Not (mbytInFun = 0 And Not mblnSaveAsPrice) Then CheckBillNOAndBookeFee = True: Exit Function
    End If
    mblnPrint = True
    '检查是否打印票据
    If mintInvoicePrint = 0 Then
        mblnPrint = False
    Else
        If mintInvoicePrint = 2 Then
            If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                mblnPrint = False
            End If
        End If
    End If
    If Not blnReCharge Then
        '检查零费用(只有工本费)是否打印,划价不产生工本费,多张中的某一张只有工本费时，在打印调用时判断不打印
        If mblnPrint And gTy_Module_Para.bln工本费 Then
            If GetBillSum = Calc工本费 Then
                If MsgBox("当前单据实际没有收取费用,要打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
    End If
    If Not mblnPrint Then
        If blnReCharge Then
                CheckBillNOAndBookeFee = True: Exit Function
        End If
        If gTy_Module_Para.bln工本费 Then
            j = 0
            For p = 1 To mobjBill.Pages.Count
                For i = 1 To mobjBill.Pages(p).Details.Count
                    If mobjBill.Pages(p).Details(i).工本费 Then
                        If j = 0 Then MsgBox "因为不打印票据,系统将自动删除工本费！", vbInformation, gstrSysName
                        j = j + 1
                        Call DeleteDetail(i, p)
                        Call ShowDetails
                        Call ShowMoney(p)
                        Bill.TxtVisible = False: Bill.CmdVisible = False: Bill.CboVisible = False
                        Exit For
                    End If
                Next
            Next
        End If
    Else
        If gblnStrictCtrl Then
            If Trim(txtInvoice.Text) = "" Then
                MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Function
            End If

InvoiceHandle:
            If zlGetInvoiceGroupUseID(mlng领用ID, IIf(gTy_Module_Para.bln分别打印 And mbytBillSource <> 4, mobjBill.Pages.Count, 1), txtInvoice.Text) = False Then
                Exit Function
            End If
            '并发操作检查,票号是否已用
            If CheckBillRepeat(mlng领用ID, 1, txtInvoice.Text) Then
                'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
                If txtInvoice.Locked = False And txtInvoice.Tag <> Trim(txtInvoice.Text) Then
                    MsgBox "票据号""" & txtInvoice.Text & """已经被使用，请重新输入。", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Function
                Else
                    Call RefreshFact
                    If txtInvoice.Text = "" Then
                        txtInvoice.SetFocus: Exit Function
                    Else
                        MsgBox "当前票据号已经被使用，已重新获取票据号:" & txtInvoice.Text, vbInformation, gstrSysName
                        GoTo InvoiceHandle
                    End If
                End If
            End If
        Else
            If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Function
            End If
        End If
    End If
    If blnReCharge Then
        CheckBillNOAndBookeFee = True: Exit Function
    End If
    '明细必须等于汇总
    dbl金额 = GetBillSum - GetMedicareSum
    For j = 1 To mobjBill.Pages.Count
        dbl金额 = RoundEx(dbl金额 + Val(mobjBill.Pages(j).误差金额) - Val(mobjBill.Pages(j).应缴金额) - Val(mobjBill.Pages(j).冲预交额), 7)
    Next
    If dbl金额 <> 0 Then
        MsgBox "实收金额合计与支付金额合计不符,不允许保存!" & vbCrLf & vbCrLf & _
            "单据明细实收金额合计+误差金额-(保险支付合计+应缴合计+冲预交金额)=" & dbl金额, vbInformation, gstrSysName
        Exit Function
    End If
    CheckBillNOAndBookeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckInsure() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保相关检查
    '返回:成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-08-16 16:48:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNone As String
    On Error GoTo errHandle
    If mstrYBPati = "" Then CheckInsure = True: Exit Function
    If mbytInFun <> 0 Then CheckInsure = True: Exit Function
    If mintInsure = 61 Then '咸阳医保
        If Not 门诊预结算(strNone) Then
            If strNone <> "" Then
                MsgBox "当前保险结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                    "在门诊未设置，请先到结算方式管理中设置这些结算方式！", vbInformation, gstrSysName
            End If
            If cmd预结算.Visible Then
                cmd预结算.TabStop = True
                cmdOK.Enabled = False
                cmd预结算.SetFocus
            End If
            Exit Function
        End If
    End If
    CheckInsure = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function SaveChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、划价、记帐
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 10:22:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur消费金额 As Currency, strNos As String, cur已缴合计 As Currency
    Dim strModiNos As String, dbl缴款 As Double, dbl找补 As Double
    Dim bln连续  As Boolean, blnGetFact As Boolean, i As Long, j As Long
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim bytReturnMode As ExitMode
    Dim str划价Nos As String, rsItems As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not (mstrYBPati <> "" And MCPAR.门诊连续收费) And Not mblnSaveAsPrice Then
        'If txt缴款.Enabled And txt缴款.Visible Then
            Call AutoBultBookFee '收费时自动产生工本费项目
        'End If
    End If
    
    If isValiedCargeFee = False Then Exit Function
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    If zlGetSaveDataItems_Plugin(mobjBill, str划价Nos, rsItems) = False Then Exit Function
    If zlChargeSaveValied_Plugin(glngModul, IIf(mbytInFun = 2, 2, 1), True, _
                                 mbytInFun = 1 Or (mbytInFun = 2 And mbytBilling = 1), str划价Nos, rsItems) = False Then Exit Function
    '只有记账刷卡验证(冲预交在结算窗口处理)
    If (mbytInFun = 2 And mbytBilling = 0 And Val(txt合计.Text) <> 0) And gdbl预存款消费验卡 <> 0 Then
        cur消费金额 = Val(txt合计.Text)
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.病人ID, cur消费金额, mlngModul, 1, , IIf(-1 * gdbl预存款消费验卡 >= cur消费金额, False, True), , , , (gdbl预存款消费验卡 = 2)) Then Exit Function
    End If
    
    '票据号及工本费及汇总金额相关检查
    If CheckBillNOAndBookeFee = False Then Exit Function
    If CheckInsure = False Then Exit Function
        
    On Error GoTo errH
    cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保结算延时
    cmdCancel.Enabled = False: cmdAddBill.Enabled = False: cmdDelBill.Enabled = False
    If cmd预结算.Visible And cmd预结算.Enabled Then cmd预结算.Enabled = False
    Dim blnSaveBill As Boolean '单据是否保存成功
    '保存单据
    '---------------------------------------------------------------------------------------------
    strNos = "": bytReturnMode = 0
    If Not SaveBill(strNos, strModiNos, blnSaveBill, False, bytReturnMode, bln连续) Then
        '收费,保存单据失败后的处理
         If blnSaveBill And bytReturnMode <> EM_本次作废 Then
            If bytReturnMode <> EM_退出收费 Then Call ShowBillChargeFee(mlng结算序号)
         End If
        mlng结算序号 = 0
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        If mintInsure <> 0 Then
            cmdAddBill.Enabled = Not MCPAR.门诊连续收费 And MCPAR.多单据收费 And InStr(1, mstrPrivs, "医保病人多单据收费") > 0
        Else
            cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
        End If
        
        If cmdDelBill.Visible And tbsBill.Tabs.Count > 1 Then cmdDelBill.Enabled = True
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        If bytReturnMode = EM_本次作废 Then
                If mblnAutoChangePati And gint病人来源 = 2 Then
                    '需要切找到病人来源1中
                    gint病人来源 = 1: zlChangePatiSource (gint病人来源)
                End If
                Call ClearFullBill(False)
        End If
        Exit Function
    End If
    Call zlChargeSaveAfter_Plugin(glngModul, mobjBill.病人ID, mobjBill.主页ID, True, IIf(mbytInFun = 2, 2, 1), strNos)
    mlng结算序号 = 0
    '显示Led相关信息
     If mbytInFun = 0 And Not mblnSaveAsPrice Then
        'LED显示:(合计,)发药窗口
        If gblnLED And CCur(txt合计.Text) <> 0 And (mstr西窗 <> "" Or mstr中窗 <> "" Or mstr成窗 <> "") Then
            zl9LedVoice.DisplayBank "费用合计:" & txt合计.Text, _
                "取药窗口:" & IIf(mstr西窗 <> "", " " & mstr西窗, "") & _
                IIf(mstr成窗 <> "", " " & mstr成窗, "") & IIf(mstr中窗 <> "", " " & mstr中窗, "")
        End If
     End If
     
    Call SendMsgModule
     '打印票据
    Call PrintBill(strNos, strModiNos)
    
    If mbytInFun = 2 And mbytBilling = 0 Then
        '110319
        If mblnDrugMachine Then
            Dim strData As String, strReturn As String
            If mstrInNO <> "" Then
                '修改单据，删除了原来的单据的
                Dim rsTemp As ADODB.Recordset, strSQL As String
                '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
                strSQL = "Select Id As 费用id, -1 * Nvl(付数, 1) * 数次 As 退药数量" & vbNewLine & _
                        " From 门诊费用记录" & vbNewLine & _
                        " Where 记录性质 = 2 And 记录状态 = 2 And NO = [1] And 收费类别 In ('5', '6', '7')" & vbNewLine & _
                        "       And 登记时间 + 0 = (Select Max(登记时间)" & vbNewLine & _
                        "                       From 门诊费用记录" & vbNewLine & _
                        "                       Where 记录性质 = 2 And 记录状态 = 2 And NO = [1])"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询本次退费项目", mstrInNO)
                Do While Not rsTemp.EOF
                    strData = strData & ";" & Nvl(rsTemp!费用id) & "," & Nvl(rsTemp!退药数量)
                    rsTemp.MoveNext
                Loop
                If strData <> "" Then
                    strData = Mid(strData, 2)
                    Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
                End If
            End If
        
            '门诊格式：1|单据1,处方号1;单据2,处方号2
            strData = "1|" & "9," & Replace(Replace(strNos, "'", ""), ",", ";9,")
            Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
        End If
    End If
    
    cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
    cmdCancel.Enabled = True
    If cmd预结算.Visible Then cmd预结算.Enabled = True
    If mbytInFun = 0 And mbytInState = 0 And gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
        
    If mstrInNO = "" And Not mblnCopyBill Or (txtModi.Visible And mbytInState = 0 And mstrInNO <> "") Then   '新增,或新增界面通过输入单据号修改单据
         If fraUpBillShow.Visible Then
            txtPreNO.Text = mobjBill.NO
            '27505
             txtPreMoney.Text = Format(GetBillSum(False, 1), "0.00")
        Else
            sta.Panels(Pan.C2提示信息) = "上一张单据:" & mobjBill.NO '多单据时为第一张
        End If
        
        '如果是修改，则退出修改状态
        If txtModi.Visible And mbytInState = 0 And mstrInNO <> "" Then
            txtModi.Text = "": txtModi.Enabled = True
            If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
        End If
        
        mstrInNO = "":  mlngFirstID = 0: mstrFirstWin = ""
        
        If mbytInFun = 0 Or mbytInFun = 1 Then
            '以下情况终止连续收费：
            '1.医保病人每次刷卡,当次收费结束(除非设置仅缴款结束参数)
            '2.使用预交款结算,当次收费结束(除非设置仅缴款结束参数)
            '3.一次多张单据,当次收费结束(除非设置仅缴款结束参数)
            '3.如已缴款,则强行作为病人收费结束
            '4.划价时没有输入病人姓名
            '5.使用多种结算方式结算
            '6.收费时保存为划价单
            
            '刘兴洪:22343:gTy_Module_Para.byt缴款控制:0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计(改变病人除外)，2-收费时必须要输入缴款金额
            'bln连续 = Not ((mstrYBPati <> "" And Not gbln缴款结束) _
                        Or (Val(txt预交冲款.Text) <> 0 And Not gbln缴款结束) _
                        Or mobjBill.Pages.Count > 1 And Not gbln缴款结束 _
                        Or Val(txt缴款.Text) <> 0 _
                        Or mobjBill.姓名 = "" And mbytInFun = 1 _
                        Or mobjBill.Pages(mintPage).收费结算 <> "") '多种结算方式
            
            '缴款控制:0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计
            '       2-收费时必须要输入缴款金额
    
               '         Or Val(txt缴款.Text) <> 0
            If mbytInFun = 0 Then
            '    bln连续 = bln连续
            Else
                bln连续 = Not ((mstrYBPati <> "" And gTy_Module_Para.byt缴款控制 <> 1 And gTy_Module_Para.byt缴款控制 <> 1 And gTy_Module_Para.byt缴款控制 <> 3) _
                        Or (Val(txt预交冲款.Text) <> 0 And gTy_Module_Para.byt缴款控制 <> 1 And gTy_Module_Para.byt缴款控制 <> 3) _
                        Or (mobjBill.Pages.Count > 1 And gTy_Module_Para.byt缴款控制 <> 1 And gTy_Module_Para.byt缴款控制 <> 3) _
                        Or (mobjBill.姓名 = "" And mbytInFun = 1) _
                        Or mobjBill.Pages(mintPage).收费结算 <> "") '多种结算方式
                bln连续 = bln连续 Or (mstrYBPati <> "" And MCPAR.门诊连续收费)
             End If
            If Not bln连续 Or mblnSaveAsPrice Then
                If gint病人来源 = 2 And mblnAutoChangePati Then
                
                    '自动切换的,要换回来
                    gint病人来源 = 1
                    Call zlChangePatiSource(gint病人来源)
                End If
                Call ClearPatientInfo(True)
                Call ClearTotalInfo(True)
                Call InitCommVariable
                blnGetFact = IIf(mblnStartFactUseType, False, True)
            Else
                '虽然连续,但医保病人清除姓名以便再次验证
                If mstrYBPati <> "" Then Call ClearPatientInfo(True)
                blnGetFact = True
                mstrPrePati = mobjBill.姓名 '记录当前病人
                mlngPrePati = mobjBill.病人ID
                mstrPreDoctor = zlStr.NeedName(cbo开单人.Text)
                
                '病人单据金额累加
                mcurBill应收 = mcurBill应收 + GetBillSum(True)
                mcurBill实收 = mcurBill实收 + GetBillSum
                mcurBill应缴 = GetMustPaySum
                
                mintBillNO = mintBillNO + 1
                For i = 1 To mshMoney.Rows - 1
                    If mshMoney.TextMatrix(i, 0) = "" Then Exit For
                Next
                mintMoneyRow = i - 1
                
                Call SaveDrugID(mobjBill.Pages.Count)
            End If
        End If
        
         '门诊划价,门诊记帐划价保留单据内容
        If mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1 Then
            Bill.Active = False
        Else
            Call ClearBillRows
        End If
        If Not (mbytInFun = 0 Or mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) Then
            Call ClearMoney
        End If
        
        If mbytInFun = 0 And (mstrYBPati <> "" And MCPAR.门诊连续收费) Then
            Call NewYBBill
            mobjBill.病人ID = CLng(Split(mstrYBPati, ";")(8))
            
            '重新读取预交余额
            If txt预交冲款.Enabled Then Call LoadFeeInfor(mobjBill.病人ID)
            
            '重新读取个帐余额
            mcur个帐余额 = gclsInsure.SelfBalance(mobjBill.病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
            sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
            sta.Panels(Pan.C3个人帐户).Visible = True

            mstrYBPati = ""
        Else
            Call NewBill(blnGetFact, Not Bill.Active And mbytInFun <> 1, Not mbln补费)      '划价单时不更改费别
            If mbln补费 Then
                With mobjBill
                    .病人ID = IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID)
                    .主页ID = IIf(mbln补费 And mlng主页ID <> 0, mlng主页ID, Nvl(mrsInfo!主页ID, 0))
                    .标识号 = IIf(gint病人来源 = 2, Nvl(mrsInfo!住院号, 0), Nvl(mrsInfo!门诊号, 0))
                    .姓名 = "" & mrsInfo!姓名
                    .性别 = "" & mrsInfo!性别
                    .年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
                    .床号 = "" & mrsInfo!当前床号
                    .病区ID = IIf(mbln补费 And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!当前病区ID)))
                    .科室ID = IIf(mbln补费 And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!当前科室id)))
                    .费别 = zlStr.NeedName(cbo费别.Text) '以当前有效为准
                End With
                Bill.SetFocus
            End If
            If Not (mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) Then Call SetDisible(True) 'Active=False
        End If
        
        '提醒票据是否充ss足
        If Not mblnStartFactUseType Then Call zlCheckFactIsEnough
        
        If Not txtPatient.Locked And txtPatient.Enabled Then
            txtPatient.SetFocus
        Else
            Bill.SetFocus
        End If
        mblnSaveData = True
    Else '从主界面选择修改单据
        '问题:44196
        mlng结算序号 = 0: SaveChargeFee = True
        Unload Me: Exit Function
    End If
    If mbytInFun = 0 Then
        Call LoadCurBalance
    End If
    mlng结算序号 = 0
    SaveChargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
    End If
    cmdOK.Enabled = True
    Call SaveErrLog
End Function
Private Function zlAutoPayDrugAndStuff(ByRef cllDrugAndStuff As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行自动发料
    '返回:发料成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-06 14:55:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If cllDrugAndStuff Is Nothing Then zlAutoPayDrugAndStuff = True: Exit Function
    
    zlExecuteProcedureArrAy cllDrugAndStuff, Me.Caption
    zlAutoPayDrugAndStuff = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Set应缴累计(ByVal bln连续 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置应缴累计
    '编制:刘兴洪
    '日期:2012-02-06 14:59:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl本次应缴 As Double
    
    mblnNotClearLedDisplay = True
    mbln连续输入 = False
    If Not (mstrYBPati <> "" And bln连续 Or mstrYBPati = "" And bln连续) Then Exit Sub
    mbln连续输入 = True
    For i = 1 To mobjBill.Pages.Count
        mobjBill.Pages(i).应缴金额 = 0
    Next
    If grsTotal.RecordCount <> 0 Then grsTotal.MoveFirst
    dbl本次应缴 = 0
    Do While Not grsTotal.EOF
        '性质:0-缴款;1-找补,2-冲预交;其他(mod 10:0-普通结算;1-医保结算;2-三方接品;3-一卡通)
        If Val(Nvl(grsTotal!性质)) <> 11 Then
            '非医保的累计
            dbl本次应缴 = dbl本次应缴 + Val(Nvl(grsTotal!结算金额))
        End If
        grsTotal.MoveNext
    Loop
    mobjBill.Pages(1).应缴金额 = dbl本次应缴
End Sub
Public Sub zlGetClassMoney(ByRef rsClass As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取分类汇总金额
    '编制:刘兴洪
    '日期:2011-12-26 13:19:04
    '问题:44944
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim p As Integer, strNos As String, dbl实收金额 As Double
    Dim i As Integer, j As Integer, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strNos = ""
    Set rsClass = New ADODB.Recordset
    rsClass.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsClass.Fields.Append "金额", adDouble, , adFldIsNullable
    rsClass.CursorLocation = adUseClient
    rsClass.LockType = adLockOptimistic
    rsClass.CursorType = adOpenStatic
    rsClass.Open
    With mobjBill
        For p = 1 To .Pages.Count
             If .Pages(p).NO <> "" Then        '提取的是划价单
                  strNos = strNos & "," & .Pages(p).NO & ""
             Else
                For i = 1 To .Pages(p).Details.Count
                    dbl实收金额 = 0
                    With .Pages(p).Details(i)
                        For j = 1 To .InComes.Count
                            dbl实收金额 = dbl实收金额 + .InComes(j).实收金额
                        Next
                        rsClass.Find "收费类别='" & .收费类别 & "'", , adSearchForward, 1
                        If rsClass.EOF Then rsClass.AddNew
                        rsClass!收费类别 = .收费类别
                        rsClass!金额 = Val(Nvl(rsClass!金额)) + dbl实收金额
                        rsClass.Update
                    End With
                Next
            End If
        Next
    End With
    If strNos = "" Then Exit Sub
    strNos = Mid(strNos, 2)
    strSQL = _
    "  Select /*+ RULE */ A.收费类别,  Sum(实收金额) As 实收金额 " & _
    "  From 门诊费用记录 A, Table( f_Str2list([1])) J" & _
    "  Where A.NO=J.Column_Value and A.记录性质=1 And A.记录状态=0  " & _
    " Group By  收费类别 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取收费类别的划价信息", strNos)
    If rsTemp.RecordCount = 0 Then Exit Sub
    Do While Not rsTemp.EOF
        rsClass.Find "收费类别='" & Nvl(rsTemp!收费类别) & "'", , adSearchForward, 1
        If rsClass.EOF Then rsClass.AddNew
        rsClass!收费类别 = Nvl(rsTemp!收费类别)
        rsClass!金额 = Val(Nvl(rsClass!金额)) + Val(Nvl(rsTemp!实收金额))
        rsClass.Update
        rsTemp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlChargeFeeWin() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入收费界面
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-05 16:20:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytReturnMode As ExitMode, bln连续 As Boolean, dbl本次应缴 As Double
    Dim blnGetFact As Boolean, i As Integer, p As Integer
    Dim strReturn As String, lng结算序号 As Long, lng病人ID As Long
    
    On Error GoTo Errhand
    If Not (mstrYBPati <> "" And MCPAR.门诊连续收费) And Not mblnSaveAsPrice Then
            Call AutoBultBookFee '收费时自动产生工本费项目
    End If
    If isValiedCargeFee = False Then Exit Function
    
    '票据号及工本费及汇总金额相关检查
    If CheckBillNOAndBookeFee = False Then Exit Function
    If CheckInsure = False Then Exit Function
          
    
    Set mcllPayDrugAndStuff = New Collection
    Set mFrmBalanceWin = New frmChargePayMentWin
    
    lng病人ID = mobjBill.病人ID
    If mFrmBalanceWin.zlChargeWin(Me, EM_正常收费, mlngModul, mstrPrivs, _
        mlngShareUseID, mstrUseType, 0, "", "", mobjBill.病人ID, mintInsure, _
        mobjBill.姓名, mobjBill.性别, mobjBill.年龄, mobjBill.费别, mdbl缴款, mdbl找补, _
        bytReturnMode, CDbl(mcurBill应缴), bln连续, mlngPreBrushCard, dbl本次应缴, mstrBalance) = False Then
        If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
        '收费,保存单据失败后的处理
         If bytReturnMode <> EM_本次作废 And bytReturnMode <> EM_退出收费 Then
             Call ShowBillChargeFee(mlng结算序号)
         End If
         mlng结算序号 = 0
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        If mintInsure <> 0 Then
            cmdAddBill.Enabled = Not MCPAR.门诊连续收费 And MCPAR.多单据收费 And InStr(1, mstrPrivs, "医保病人多单据收费") > 0
        Else
            cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
        End If
        If cmdDelBill.Visible And tbsBill.Tabs.Count > 1 Then cmdDelBill.Enabled = True
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        If bytReturnMode = EM_本次作废 Then
                If mblnAutoChangePati And gint病人来源 = 2 Then
                    '需要切找到病人来源1中
                    gint病人来源 = 1: zlChangePatiSource (gint病人来源)
                End If
                Call ClearFullBill(False)
        End If
        Exit Function
    End If
    If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
    lng结算序号 = mlng结算序号
    '设置应缴累计
    Call Set应缴累计(bln连续)
    If mblnDrugPacker Then
        '51510
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.编号, UserInfo.姓名, 0, "8," & Replace(Replace(mstrSaveNos, "'", ""), ",", "|8,"), strReturn)
    End If
    '自动发药和发料处理
    Call zlAutoPayDrugAndStuff(mcllPayDrugAndStuff)
    
    '消息发送
    Call SendMsgModule
    
    mlng结算序号 = 0
    '显示Led:发药窗口及费用合计金额
    Call ShowLedWinAndSum
    '票据打印,打印票据
    Call PrintBill(mstrSaveNos, mstrModiNOs)
    '设置其他相关内容
    '防止设置打印机弹出的非模态窗体,以及医保延时
    '写卡:56615
    Call WriteMzInforToCard(lng病人ID, lng结算序号)
    
    cmdOK.Enabled = True: cmdCancel.Enabled = True
    If cmd预结算.Visible Then cmd预结算.Enabled = True
    If mbytInFun = 0 And mbytInState = 0 And gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    If Not (mstrInNO = "" Or (txtModi.Visible And mbytInState = 0 And mstrInNO <> "")) Then
        '从主界面选择修改单据
        '问题:44196
        zlChargeFeeWin = True
        Unload Me: Exit Function
    End If
    
    '新增,或新增界面通过输入单据号修改单据
     If fraUpBillShow.Visible Then
        txtPreNO.Text = mobjBill.NO
         txtPreMoney.Text = Format(GetBillSum(False, 1), "0.00")
    Else
        sta.Panels(Pan.C2提示信息) = "上一张单据:" & mobjBill.NO '多单据时为第一张
    End If
    i = UBound(Split(mstrSaveNos, ",")) + 1
    If i <> mobjBill.Pages.Count Then
        If MsgBox("目前病人只收费了" & i & "张单据,是否对未收费单据进行重收费!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            '删除已经收费的单据
            '删除单据
            For p = 1 To i
                Call DelOneBill(1)
            Next
            '重新计算
            Call ShowMoney(0)   '其它单据费用未变
            '重新设置工本费(包含了重新计算)
            If gTy_Module_Para.bln工本费 Then
                If Not CheckBillsEmpty Then Call SetFactMoney
            End If
            Exit Function
        End If
    End If
    '如果是修改，则退出修改状态
    If txtModi.Visible And mbytInState = 0 And mstrInNO <> "" Then
        txtModi.Text = "": txtModi.Enabled = True
        If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
    End If
    mstrInNO = "":  mlngFirstID = 0: mstrFirstWin = ""
    
    If mbytInFun = 0 Or mbytInFun = 1 Then
        '以下情况终止连续收费：
        '1.医保病人每次刷卡,当次收费结束(除非设置仅缴款结束参数)
        '2.使用预交款结算,当次收费结束(除非设置仅缴款结束参数)
        '3.一次多张单据,当次收费结束(除非设置仅缴款结束参数)
        '3.如已缴款,则强行作为病人收费结束
        '4.划价时没有输入病人姓名
        '5.使用多种结算方式结算
        '6.收费时保存为划价单
        
        '刘兴洪:22343:gTy_Module_Para.byt缴款控制:0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计(改变病人除外)，2-收费时必须要输入缴款金额
        'bln连续 = Not ((mstrYBPati <> "" And Not gbln缴款结束) _
                    Or (Val(txt预交冲款.Text) <> 0 And Not gbln缴款结束) _
                    Or mobjBill.Pages.Count > 1 And Not gbln缴款结束 _
                    Or Val(txt缴款.Text) <> 0 _
                    Or mobjBill.姓名 = "" And mbytInFun = 1 _
                    Or mobjBill.Pages(mintPage).收费结算 <> "") '多种结算方式
        
        '缴款控制:0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计
        '       2-收费时必须要输入缴款金额

           '         Or Val(txt缴款.Text) <> 0
        If Not bln连续 Then
            If gint病人来源 = 2 And mblnAutoChangePati Then
                '自动切换的,要换回来
                gint病人来源 = 1
                Call zlChangePatiSource(gint病人来源)
            End If
            Call ClearPatientInfo(True)
            Call ClearTotalInfo(True)
            Call InitCommVariable
            blnGetFact = IIf(mblnStartFactUseType, False, True)
        Else
            '虽然连续,但医保病人清除姓名以便再次验证
            If mstrYBPati <> "" Then Call ClearPatientInfo(True)
            blnGetFact = True
            mstrPrePati = mobjBill.姓名 '记录当前病人
            mlngPrePati = mobjBill.病人ID
            mstrPreDoctor = zlStr.NeedName(cbo开单人.Text)
            
            '病人单据金额累加
            mcurBill应收 = mcurBill应收 + GetBillSum(True)
            mcurBill实收 = mcurBill实收 + GetBillSum
            mcurBill应缴 = GetMustPaySum
            mintBillNO = mintBillNO + 1
            For i = 1 To mshMoney.Rows - 1
                If mshMoney.TextMatrix(i, 0) = "" Then Exit For
            Next
            mintMoneyRow = i - 1
            Call SaveDrugID(mobjBill.Pages.Count)
        End If
    End If
    Call ClearBillRows
    If (mstrYBPati <> "" And MCPAR.门诊连续收费) Then
        Call NewYBBill
        mobjBill.病人ID = CLng(Split(mstrYBPati, ";")(8))
        
        '重新读取预交余额
        If txt预交冲款.Enabled Then Call LoadFeeInfor(mobjBill.病人ID)
        '重新读取个帐余额
        mcur个帐余额 = gclsInsure.SelfBalance(mobjBill.病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
        sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
        sta.Panels(Pan.C3个人帐户).Visible = True

        mstrYBPati = ""
    Else
        Call NewBill(blnGetFact, Not Bill.Active And mbytInFun <> 1)      '划价单时不更改费别
        If Not (mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) Then Call SetDisible(True) 'Active=False
    End If
    '提醒票据是否充ss足
    If Not mblnStartFactUseType Then Call zlCheckFactIsEnough
    If Not txtPatient.Locked Then
       If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Else
        Bill.SetFocus
    End If
    mblnSaveData = True
    Call LoadCurBalance
    zlChargeFeeWin = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub ShowLedWinAndSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示发药窗口及相关合计数据
    '编制:刘兴洪
    '日期:2012-02-06 14:31:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
    If Not (mbytInFun = 0 And Not mblnSaveAsPrice) Then Exit Sub
    If Not (mstr西窗 <> "" Or mstr中窗 <> "" Or mstr成窗 <> "") _
        Or CCur(txt合计.Text) = 0 Then Exit Sub
    zl9LedVoice.DisplayBank "费用合计:" & txt合计.Text, _
        "取药窗口:" & IIf(mstr西窗 <> "", " " & mstr西窗, "") & _
        IIf(mstr成窗 <> "", " " & mstr成窗, "") & IIf(mstr中窗 <> "", " " & mstr中窗, "")
End Sub
 


Private Sub cmdOK_Click()
     mblnSaveData = False
    '成都妇幼新增(门诊划价保留单据内容，整段)
    If (mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) And mbytInState = 0 And Not Bill.Active Then
        Call ClearBillRows: Call ClearMoney
        Call ClearTotalInfo(True)
        Bill.Active = True
        MsgBox "请输入新的单据内容！", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    
    'If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
    '设置按钮不可用,必免重复执行
    
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Value = 1 And chkCancel.Visible) Then
        '========================================================================================================
        '退费或销帐(主界面调用或正常界面使用退功能)
        If DelCharge = False Then Exit Sub
    ElseIf mbytInState = 2 Then '调整单据
        '========================================================================================================
        If Not SaveModi() Then Exit Sub
        mblnSaveData = True
        Unload Me
    ElseIf mbytInFun = 2 And mbytBilling = 2 Then '记帐划价单审核操作
        If SaveVerify = False Then Exit Sub
    ElseIf (mbytInState = 0) And chkCancel.Value = 0 Then
        '收费，记帐，划价：正常输入单据状态,收费可能是划价单混合
        '包含异常单据的重新收费
        Call GetAsyncKeyState(VK_RETURN)
        If mbytInFun = 0 And Not mblnSaveAsPrice Then
            If zlChargeFeeWin = False Then Exit Sub
        Else
            If SaveChargeFee = False Then Exit Sub
            
        End If
    
    ElseIf mbytInState = 4 Then
        cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保延时
        cmdCancel.Enabled = False: cmdAddBill.Enabled = False:: cmdDelBill.Enabled = False
        If ReChargeFee = False Then
            '61688
            cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
            cmdCancel.Enabled = True
            Exit Sub
        End If
    ElseIf mbytInState = 5 Then
        '作废异常单据
        cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保延时
        cmdCancel.Enabled = False: cmdAddBill.Enabled = False:: cmdDelBill.Enabled = False
        If DelErrBillFee = False Then
            cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
            cmdCancel.Enabled = True
            Exit Sub
        End If
    End If
    mblnSaveData = True
    gblnOK = True
    Exit Sub
End Sub

Private Sub LoadFeeInfor(ByVal lngPatientID As Long, Optional ByVal blnDelete As Boolean)
'功能:读取并显示病人预交,及费用余额信息
'参数:blnDel-是否退费功能
    Dim rsTmp As ADODB.Recordset
    Dim cur实收合计 As Currency
 
    If mbytInFun = 0 Then
        Set rsTmp = GetMoneyInfo(lngPatientID, , , 1)
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!预交余额
            cmdCancel.Tag = rsTmp!费用余额
            cmdPrint.Tag = Val(cmdOK.Tag) - Val(cmdCancel.Tag)
            If mbytInState = 0 And mstrInNO <> "" Then cmdPrint.Tag = Val(cmdPrint.Tag) + Original.冲预交款
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: cmdPrint.Tag = 0
        End If
        sta.Panels(Pan.C4预交信息).Tag = cmdPrint.Tag
        sta.Panels(Pan.C4预交信息).Text = "预交:" & Format(Val(cmdPrint.Tag), "0.00")
        Call ShowPrePayInfo(Val(cmdPrint.Tag) > 0 And Not blnDelete)
        
    ElseIf mbytInFun = 2 Then
        '门诊记帐及划价不会使用预交,所以不用考虑Original.冲预交款
        
        '记录报警所需数据
        '修改记帐单时,费用余额不含当前单据金额(记帐单划价时未计入病人余额的费用余额)
        Set rsTmp = GetMoneyInfo(lngPatientID, IIf(mbytBilling = 0, Original.实收合计, 0), , 1)
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!预交余额
            cmdCancel.Tag = rsTmp!费用余额
            cmdPrint.Tag = Val(cmdOK.Tag) - Val(cmdCancel.Tag)
            sta.Panels(Pan.C4预交信息).Visible = True
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: cmdPrint.Tag = 0
            sta.Panels(Pan.C4预交信息).Visible = False
        End If
                
        
        '划价显示时不含当前单据费用(因为未审核),但划价报警要算
        If mbytBilling = 0 Then
            If mbytInState = 0 And mstrInNO <> "" Then
                cur实收合计 = Original.实收合计
            Else
                cur实收合计 = GetBillSum
            End If
        End If
        sta.Panels(Pan.C4预交信息).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(Pan.C4预交信息).Text = sta.Panels(Pan.C4预交信息).Text & "/费用:" & Format(Val(cmdCancel.Tag) + cur实收合计, "0.00")
        sta.Panels(Pan.C4预交信息).Text = sta.Panels(Pan.C4预交信息).Text & "/剩余:" & Format(Val(cmdPrint.Tag) - cur实收合计, "0.00")
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim blnBillsEmpty As Boolean
    mbln连续输入 = False
    blnBillsEmpty = CheckBillsEmpty()
    If mbln补费 And blnBillsEmpty Then
        Unload Me: Exit Sub
    End If
    
    If (Not blnBillsEmpty Or txtPatient.Text <> "") And mbytInState = 0 And mstrInNO = "" And Not mblnCopyBill Then
        If ClearFullBill(True, Not mbln补费) = False Then Exit Sub
        '问题:27364 日期:2010-01-13 15:27:50
        If mblnAutoChangePati And gint病人来源 = 2 Then
            '需要切找到病人来源1中
            gint病人来源 = 1: zlChangePatiSource (gint病人来源)
        End If
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub SaveDrugID(intPage As Integer)
'功能:保存当前指定单据序号以前中最后一个含有药品的单据的第一行药品的部门ID
    Dim i As Long, j As Long
    
    '记录该病人本次收费分配的各药房(单据内容为输入时)
    For i = 1 To intPage
        If mobjBill.Pages(i).NO = "" Then
            j = GetFirstRow(mobjBill, i)
            If j > 0 Then
                Select Case mobjBill.Pages(i).Details(j).收费类别
                    Case "5"
                        mlng西药房 = mobjBill.Pages(i).Details(j).执行部门ID
                    Case "6"
                        mlng成药房 = mobjBill.Pages(i).Details(j).执行部门ID
                    Case "7"
                        mlng中药房 = mobjBill.Pages(i).Details(j).执行部门ID
                End Select
            End If
        Else
            Call BillDrugDept(mobjBill.Pages(i).NO, mlng西药房, mlng成药房, mlng中药房)
        End If
    Next
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.COLS - 1
    End If
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '保存为划价单
    If Button = 2 Then
        If CheckSaveMultiPrice Then
            PopupMenu mnuFile, 2, cmdOK.Left + picAppend.Left - 800, cmdOK.Top + cmdOK.Height + picAppend.Top
        End If
    End If
End Sub
Private Sub cmdPrint_Click()
    Dim i As Integer, j As Integer
    Dim strPrintNO As String, strInfo As String
    Dim blnPrintList As Boolean
    
    If mstrYBBill = "" Then
        MsgBox "该医保病人本次还没有收取费用！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnPrint Then
        If gclsInsure.GetCapability(support门诊收费完成后验证, mobjBill.病人ID, mintInsure) Then
            If gclsInsure.Identify(id门诊确认, , mintInsure) = "" Then
                MsgBox "病人身份验证失败，不能完成收费打印操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.Refresh
        Else
            If MsgBox("确实要完成收费操作并打印票据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    
    Screen.MousePointer = 11
    
    blnPrintList = False
    If InStr(mstrPrivs, "打印清单") > 0 Then
        If gint收费清单 = 1 Then
            blnPrintList = True
        ElseIf gint收费清单 = 2 Then
            If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
            End If
        End If
    End If
    For i = 0 To UBound(Split(mstrYBBill, ","))
        strPrintNO = CStr(Split(mstrYBBill, ",")(i))
        If strPrintNO <> "" Then
            If mblnPrint Then
                If Not gobjTax Is Nothing And gblnTax Then
                    If Not gobjTax Is Nothing And gblnTax Then
                        gstrTax = gobjTax.zlTaxOutPrint(gcnOracle, "'" & strPrintNO & "'")
                        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                    End If
                Else
                    If gblnBillPrint Then
                        If gobjBillPrint.zlPrintBill("'" & strPrintNO & "'", 0) = False Then Exit Sub
                    End If
                    
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me, _
                        "NO='" & strPrintNO & "'", "价格等级=" & IIf(mstr普通价格等级 = "", "-", mstr普通价格等级), _
                        IIf(glngFactMediCare = 0, "", "ReportFormat=" & glngFactMediCare), 2)
                End If
            End If
            
            If blnPrintList Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & strPrintNO & "'", "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            End If
        End If
    Next
    
    mintInsure = 0: mstrYBPati = ""
    cmdPrint.SetFocus
        
    Call ClearFullBill(False)
    txtPatient.SetFocus
    Set grsTotal = Nothing
    Screen.MousePointer = 0
End Sub


Private Function zlSquareCardFeeList(ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡明细信息
    '入参:
    '出参:rsFreeList-返回明细数据
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-05 16:02:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, p As Long, strSQL As String, strDate As String, strInvoice As String
    strInvoice = ""
    If zlCreateFeeListStruc(rsFeeList) = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Err = 0: On Error GoTo Errhand:
    For p = 1 To tbsBill.Tabs.Count
          If mobjBill.Pages(p).NO = "" Then
              '直接输入的费用
              If zlBuldingFeeListdata(mobjBill, chk急诊.Value = 1, p, strDate, cbo费别.Text, strInvoice, rsFeeList) = False Then Exit Function
          Else
              '提取的划价单(售价单位)
              strSQL = _
                "Select '" & strInvoice & "' as 实际票号,NO,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as 结算时间," & _
                        mobjBill.病人ID & " As 病人ID,'" & cbo费别.Text & "' As 费别,收费类别,收据费目,计算单位,开单人," & _
                "       收费细目ID,保险大类ID As 保险支付大类ID,Nvl(保险项目否,0) As 是否医保,保险编码," & _
                "       Avg(Nvl(付数,0)*数次) As 数量,Avg(标准单价) As 单价," & _
                "       Sum(实收金额) As 实收金额,Sum(统筹金额) As 统筹金额,摘要," & _
                        chk急诊.Value & " as 是否急诊,开单部门ID,执行部门ID,保险大类ID From 门诊费用记录" & _
                " Where 记录性质=1 And 记录状态=0 And NO=[1]" & _
                " Group By Nvl(价格父号,序号),收费类别,收据费目,计算单位,开单人," & _
                "       收费细目ID,保险大类ID,Nvl(保险项目否,0),保险编码,摘要,开单部门ID,执行部门ID,NO"
              
              Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
              Do While Not rsTemp.EOF
                    rsFeeList.AddNew
                    rsFeeList!单据序号 = p
                    rsFeeList!费别 = Nvl(rsTemp!费别)
                    rsFeeList!NO = Nvl(rsTemp!NO)   '仅提取划价单时才有值
                    rsFeeList!实际票号 = strInvoice
                    rsFeeList!结算时间 = CDate(strDate)
                    rsFeeList!病人ID = IIf(mobjBill.病人ID = 0, Null, mobjBill.病人ID)
                    rsFeeList!收费类别 = Nvl(rsTemp!收费类别)
                    
                    If Nvl(rsTemp!收据费目) <> "" Then
                        rsFeeList!收据费目 = Nvl(rsTemp!收据费目)
                    Else
                        rsFeeList!收据费目 = Null
                    End If
                    rsFeeList!开单人 = Nvl(rsTemp!开单人)
                    rsFeeList!收费细目ID = Val(Nvl(rsTemp!收费细目ID))
                    rsFeeList!计算单位 = Nvl(rsTemp!计算单位)
                    rsFeeList!数量 = Val(Nvl(rsTemp!数量))
                    rsFeeList!单价 = Format(Val(Nvl(rsTemp!单价)), gstrFeePrecisionFmt)
                    rsFeeList!实收金额 = Format(Val(Nvl(rsTemp!实收金额)), gstrDec)
                    rsFeeList!统筹金额 = Format(Val(Nvl(rsTemp!统筹金额)), gstrDec)
                    rsFeeList!保险支付大类ID = IIf(Val(Nvl(rsTemp!保险大类ID)) = 0, Null, Val(Nvl(rsTemp!保险大类ID)))
                    rsFeeList!是否医保 = Val(Nvl(rsTemp!是否医保))
                    rsFeeList!保险编码 = Nvl(rsTemp!保险编码)
                    rsFeeList!摘要 = Nvl(rsTemp!摘要)
                    rsFeeList!是否急诊 = Val(Nvl(rsTemp!是否急诊))
                    rsFeeList!开单部门ID = Val(Nvl(rsTemp!开单部门ID))
                    rsFeeList!执行部门ID = Val(Nvl(rsTemp!执行部门ID))
                    rsFeeList!本次结算 = 0
                    rsFeeList.Update
                    rsTemp.MoveNext
              Loop
          End If
     Next
    zlSquareCardFeeList = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 门诊预结算不区分单据(ByVal strDate As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊预结算时,不区分单据进行结算(多单据只调用一次接口)
    '编制:刘兴洪
    '日期:2011-08-15 17:30:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, arrBalance() As String, strInvoice As String
    Dim str结算方式 As String, dbl结算金额 As Double, dbl可分配额 As Double
    Dim rsTemp  As ADODB.Recordset, i As Long, k As Long, j As Long, p As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    If MCPAR.多单据调一次交易 = False Then 门诊预结算不区分单据 = True: Exit Function
    Set rsTemp = MakeBillRecord(mobjBill, chk急诊.Value = 1, 0, strDate, cbo费别.Text, strInvoice)
    strBalance = ""
    strAdvance = tbsBill.Tabs.Count
    
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
        If tbsBill.Tabs.Count > 1 Then
            sta.Panels(Pan.C2提示信息).Text = "单据预结算失败。"
        End If
        If mstr个人帐户 <> "" And Not MCPAR.门诊预结算 Then  '只有使用个人帐户才用
            vsBalance.TextMatrix(0, 0) = mstr个人帐户
            vsBalance.TextMatrix(0, 1) = "0.00"
            vsBalance.RowData(0) = 0
            Call ShowMoney(-1, Not (cmd预结算.Visible And cmdOK.Enabled))
        End If
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
            txtMCInvoice.Text = strAdvance
            txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
            txtMCInvoice.Visible = True
    End If
    
     '根据预结算结果设置结算集
    arrPage = Array()
    If strBalance <> "" Then
        Set rsTmp = GetBalanceSet
        arrBalance = Split(strBalance, "|")
        For i = 0 To UBound(arrBalance)
            str结算方式 = Split(arrBalance(i), ";")(0)
            dbl结算金额 = Val(Split(arrBalance(i), ";")(1))
            '必须已设置该结算方式,且为医保类的结算方式
            mrs结算方式.Filter = "名称='" & str结算方式 & "' And 性质<>1 And 性质<>2"
            If mrs结算方式.EOF Then
                '记录医保有但本地没有的结算方式
                If InStr(strNone & ",", "," & str结算方式 & ",") = 0 Then
                    strNone = strNone & "," & str结算方式
                End If
            End If
            If Not mrs结算方式.EOF Then
                '只有最后呈张单据调接口时才返回结算信息，依次分摊到各单据中
                For k = 1 To tbsBill.Tabs.Count
                    dbl可分配额 = mobjBill.Pages(k).实收金额
                    rsTmp.Filter = "单据序号=" & k
                    For j = 1 To rsTmp.RecordCount
                        dbl可分配额 = dbl可分配额 - rsTmp!结算金额
                        rsTmp.MoveNext
                    Next
                    If dbl可分配额 > 0 Then
                        If dbl可分配额 <= dbl结算金额 Then
                            dbl结算金额 = dbl结算金额 - dbl可分配额
                        Else
                            dbl可分配额 = dbl结算金额
                            dbl结算金额 = 0
                        End If
                        
                        rsTmp.AddNew
                        rsTmp!单据序号 = k
                        rsTmp!结算方式 = str结算方式
                        rsTmp!结算金额 = dbl可分配额
                        rsTmp.Update
                        If dbl结算金额 = 0 Then Exit For
                    End If
                Next
                If dbl结算金额 <> 0 Then
                    '可能存在医保结算大于单据费用总额的情况,直接放在最后一张单据中
                    rsTmp.Filter = "单据序号=" & tbsBill.Tabs.Count & " and 结算方式='" & str结算方式 & "'"
                    If rsTmp.EOF Then
                        rsTmp.AddNew
                        rsTmp!单据序号 = tbsBill.Tabs.Count
                        rsTmp!结算方式 = str结算方式
                    End If
                    rsTmp!结算金额 = Val(Nvl(rsTmp!结算金额)) + dbl结算金额
                    rsTmp.Update
                    rsTmp.Filter = 0
                End If
            End If
        Next
        For p = 1 To tbsBill.Tabs.Count
            arrPage = Array()
            rsTmp.Filter = "单据序号=" & p
            For k = 1 To rsTmp.RecordCount
                ReDim Preserve arrPage(UBound(arrPage) + 1)
                arrPage(UBound(arrPage)) = rsTmp!结算方式 & ";" & rsTmp!结算金额 & ";0;" & rsTmp!结算金额
                rsTmp.MoveNext
            Next
            mcolBalance.Remove p '集合元素不能直接修改
            If mcolBalance.Count >= p Then
                mcolBalance.Add arrPage, , p
            Else
                mcolBalance.Add arrPage
            End If
        Next
    End If
    门诊预结算不区分单据 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function 门诊预结区分单据(ByVal strDate As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊预结区分单据
    '返回:成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-08-15 18:20:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, arrBalance() As String
    Dim dbl个帐合计 As Double, strInvoice As String, dbl结算金额 As Double, dbl可分配额 As Double
    Dim str结算方式 As String, p As Long, strSQL As String, i As Long, k As Long, j As Long
    Dim cur个帐 As Double
    
    strInvoice = Trim(txtInvoice.Text)
    If MCPAR.多单据调一次交易 Then 门诊预结区分单据 = True: Exit Function
    On Error GoTo errHandle
    '对多张单据循环预结算
    dbl个帐合计 = 0
    For p = 1 To tbsBill.Tabs.Count
        If mobjBill.Pages(p).NO = "" Then
            '直接输入的费用
            Set rsTmp = MakeBillRecord(mobjBill, chk急诊.Value = 1, p, strDate, cbo费别.Text, strInvoice)
        Else
            '提取的划价单(售价单位):增加序号:42961
            strSQL = _
            "   Select '" & strInvoice & "' as 实际票号,NO,  Nvl(价格父号, 序号) as 序号,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as 结算时间," & _
                    mobjBill.病人ID & " As 病人ID,'" & cbo费别.Text & "' As 费别,收费类别,收据费目,计算单位,开单人," & _
            "       收费细目ID,保险大类ID As 保险支付大类ID,Nvl(保险项目否,0) As 是否医保,保险编码," & _
            "       Avg(Nvl(付数,0)*数次) As 数量,Avg(标准单价) As 单价," & _
            "       Sum(实收金额) As 实收金额,Sum(统筹金额) As 统筹金额,摘要," & _
                    chk急诊.Value & " as 是否急诊,开单部门ID,执行部门ID " & _
            "   From 门诊费用记录" & _
            "   Where 记录性质=1 And 记录状态=0 And NO=[1]" & _
            "   Group By Nvl(价格父号,序号),收费类别,收据费目,计算单位,开单人," & _
            "       收费细目ID,保险大类ID,Nvl(保险项目否,0),保险编码,摘要,开单部门ID,执行部门ID,NO" & _
            " Order by  序号 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
        End If
        
        strBalance = ""
        strAdvance = tbsBill.Tabs.Count & "|" & p
        If Not gclsInsure.ClinicPreSwap(rsTmp, strBalance, mintInsure, strAdvance) Then
             '38821:strAdvance:发票号;是否不走票据号
            If tbsBill.Tabs.Count > 1 Then
                sta.Panels(Pan.C2提示信息).Text = "第 " & p & " 张单据预结算失败。"
            End If
            
            If mstr个人帐户 <> "" And Not MCPAR.门诊预结算 Then  '只有使用个人帐户才用
                vsBalance.TextMatrix(0, 0) = mstr个人帐户
                vsBalance.TextMatrix(0, 1) = "0.00"
                vsBalance.RowData(0) = 0
                
                Call ShowMoney(-1, Not (cmd预结算.Visible And cmdOK.Enabled))
            End If
            
            Screen.MousePointer = 0
            Exit Function
        End If
        If MCPAR.多单据一次结算 And InStr(1, strAdvance, ";") > 0 Then
              '38821:strAdvance:发票号;是否不走票据号
              MCPAR.医保不走票号 = Val(Split(strAdvance & ";", ";")(1)) = 1
        End If
        If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
             '38821:strAdvance:发票号;是否不走票据号
            txtMCInvoice.Text = Trim(Split(strAdvance & ";", ";")(0))
            txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
            txtMCInvoice.Visible = True
        End If
        '根据预结算结果设置结算集
        arrPage = Array()
        If strBalance <> "" Then
            If MCPAR.多单据一次结算 Then Set rsTmp = GetBalanceSet
            
            arrBalance = Split(strBalance, "|")
            For i = 0 To UBound(arrBalance)
                str结算方式 = Split(arrBalance(i), ";")(0)
                dbl结算金额 = Val(Split(arrBalance(i), ";")(1))
                
                '必须已设置该结算方式,且为医保类的结算方式
                mrs结算方式.Filter = "名称='" & str结算方式 & "' And 性质<>1 And 性质<>2"
                If Not mrs结算方式.EOF Then
                    If MCPAR.多单据一次结算 Then
                        '只有最后呈张单据调接口时才返回结算信息，依次分摊到各单据中
                        For k = 1 To tbsBill.Tabs.Count
                            dbl可分配额 = mobjBill.Pages(k).实收金额
                            rsTmp.Filter = "单据序号=" & k
                            For j = 1 To rsTmp.RecordCount
                                dbl可分配额 = dbl可分配额 - rsTmp!结算金额
                                rsTmp.MoveNext
                            Next
                            
                            If dbl可分配额 > 0 Then
                                If dbl可分配额 <= dbl结算金额 Then
                                    dbl结算金额 = dbl结算金额 - dbl可分配额
                                Else
                                    dbl可分配额 = dbl结算金额
                                    dbl结算金额 = 0
                                End If
                                rsTmp.AddNew
                                rsTmp!单据序号 = k
                                rsTmp!结算方式 = str结算方式
                                rsTmp!结算金额 = dbl可分配额
                                rsTmp.Update
                                
                                If dbl结算金额 = 0 Then Exit For
                            End If
                        Next
                        If dbl结算金额 <> 0 Then
                            '可能存在医保结算大于单据费用总额的情况,直接放在最后一张单据中
                            rsTmp.Filter = "单据序号=" & tbsBill.Tabs.Count & " and 结算方式='" & str结算方式 & "'"
                            If rsTmp.EOF Then
                                rsTmp.AddNew
                                rsTmp!单据序号 = tbsBill.Tabs.Count
                                rsTmp!结算方式 = str结算方式
                            End If
                            rsTmp!结算金额 = Val(Nvl(rsTmp!结算金额)) + dbl结算金额
                            rsTmp.Update
                            rsTmp.Filter = 0
                        End If
                    Else
                        If dbl结算金额 <> 0 Or str结算方式 = mstr个人帐户 Then
                            If str结算方式 = mstr个人帐户 Then
                                '咸阳医保无法返回余额
                                If (mcur个帐余额 > -1 * mcur个帐透支 Or mintInsure = 61) And CCur(txt合计.Text) > 0 Then
                                    cur个帐 = dbl结算金额
                                    If mintInsure <> 61 Then
                                        '计算个人帐户支付金额
                                        If mcur个帐余额 - dbl个帐合计 - cur个帐 >= -1 * mcur个帐透支 Then
                                            cur个帐 = cur个帐 '在允许透支范围内足够(允许透支0为特例)
                                        Else
                                            If mcur个帐透支 = 0 And mcur个帐余额 - dbl个帐合计 > 0 Then
                                                cur个帐 = mcur个帐余额 - dbl个帐合计 '不允许透支且有余额
                                            Else
                                                '超过允许透支范围或不允许透支时无余额
                                                If mcur个帐透支 <> 0 Then
                                                    cur个帐 = mcur个帐余额 - dbl个帐合计 + mcur个帐透支 '在允许透支范围内支付
                                                Else
                                                    cur个帐 = 0
                                                End If
                                            End If
                                        End If
                                    End If
                                    dbl个帐合计 = dbl个帐合计 + cur个帐
                                    cur个帐 = Format(cur个帐, "0.00")
                                    
                                    ReDim Preserve arrPage(UBound(arrPage) + 1) '结算方式;原始(最大)金额;可否修改;改后金额
                                    arrPage(UBound(arrPage)) = mstr个人帐户 & ";" & cur个帐 & ";" & Split(arrBalance(i), ";")(2) & ";" & cur个帐
                                End If
                            Else
                                ReDim Preserve arrPage(UBound(arrPage) + 1)
                                arrPage(UBound(arrPage)) = arrBalance(i) & ";" & Format(dbl结算金额, "0.00")
                            End If
                        End If
                    End If
                Else
                '记录医保有但本地没有的结算方式
                    If InStr(strNone & ",", "," & str结算方式 & ",") = 0 Then
                        strNone = strNone & "," & str结算方式
                    End If
                End If
            Next
        End If
        
        If Not MCPAR.多单据一次结算 Then
            '每个单据对应一个数组,数组可能没有元素
            mcolBalance.Remove p '集合元素不能直接修改
            If mcolBalance.Count >= p Then
                mcolBalance.Add arrPage, , p
            Else
                mcolBalance.Add arrPage
            End If
        End If
    Next
    
    If MCPAR.多单据一次结算 And strBalance <> "" Then
        For p = 1 To tbsBill.Tabs.Count
            arrPage = Array()
            rsTmp.Filter = "单据序号=" & p
            For k = 1 To rsTmp.RecordCount
                ReDim Preserve arrPage(UBound(arrPage) + 1)
                arrPage(UBound(arrPage)) = rsTmp!结算方式 & ";" & rsTmp!结算金额 & ";0;" & rsTmp!结算金额
                rsTmp.MoveNext
            Next
            mcolBalance.Remove p '集合元素不能直接修改
            If mcolBalance.Count >= p Then
                mcolBalance.Add arrPage, , p
            Else
                mcolBalance.Add arrPage
            End If
        Next
    End If

    门诊预结区分单据 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function 门诊预结算(ByRef strNone As String) As Boolean
    '功能：门诊预结算
    Dim arrBalance() As String, dbl个帐合计 As Double
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strDate As String, str结算方式 As String
    Dim dbl合计 As Double
    
    strNone = ""
    Screen.MousePointer = 11
    On Error GoTo errH
    '初始化结算结果表格
    Call InitBalanceGrid
    '获取结算时间
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If MCPAR.多单据调一次交易 = True Then
        If 门诊预结算不区分单据(strDate, strNone) = False Then Exit Function
    Else
        If 门诊预结区分单据(strDate, strNone) = False Then Exit Function
    End If
        
    '全部预结完后的处理
    '-----------------------------------------------------------
    '显示预结的表格结果
    For p = 1 To mcolBalance.Count
        For i = 0 To UBound(mcolBalance(p))
            '结算方式;原始(最大)金额;可否修改;改后金额
            arrBalance = Split(mcolBalance(p)(i), ";")
            
            '定位到匹配行或空行
            k = -1
            For j = 0 To vsBalance.Rows - 1
                If vsBalance.TextMatrix(j, 0) = arrBalance(0) Then
                    k = j: Exit For '记录已填写的匹配行
                ElseIf vsBalance.TextMatrix(j, 0) = "" Then
                    If k = -1 Then k = j '记录第一可用空行
                End If
            Next
            If j > vsBalance.Rows - 1 And k = -1 Then
                vsBalance.Rows = vsBalance.Rows + 1
                k = vsBalance.Rows - 1
            End If
            
            '汇总该种结算方式的金额
            vsBalance.TextMatrix(k, 0) = arrBalance(0)
            vsBalance.TextMatrix(k, 1) = Format(Val(vsBalance.TextMatrix(k, 1)) + Val(arrBalance(1)), "0.00")
            dbl合计 = dbl合计 + Val(Format(Val(arrBalance(1)), "0.00"))
            If vsBalance.RowData(k) = 0 Then
                '多张单据中,只要有一张允许修改,则汇总的允许修改
                vsBalance.RowData(k) = arrBalance(2)
            End If
        Next
    Next
    
    For i = 0 To vsBalance.Rows - 1
        If vsBalance.RowData(i) <> 0 Then
            vsBalance.Row = i: vsBalance.Col = 1
            vsBalance.TabStop = True
            Exit For
        End If
    Next
    
    
    '要先设置以便其它地方识别
    If cmd预结算.Visible Then
        cmd预结算.TabStop = False
        cmdOK.Enabled = True
    End If
    '重新计算应缴，误差(分币)等
    Call ShowMoney(-1, Not (cmd预结算.Visible And cmdOK.Enabled))
    With vsBalance
        For i = 0 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) = "" Then Exit For
        Next
        If i > .Rows - 1 Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = "自付合计": .TextMatrix(i, 1) = txt应缴.Text
        
        .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
        .Cell(flexcpFontBold, i, 0, i, .COLS - 1) = vbRed
        .RowPosition(i) = 0
    End With
    Call zl9InsureLedSpeak
'    dbl个帐合计 = GetMedicareSum(mstr个人帐户)
'    If gblnLED Then zl9LedVoice.DisplayBank "医保结算:", "帐户余额" & Format(mcur个帐余额, "0.00"), "帐户支付" & Format(dbl个帐合计, "0.00"), "统筹支付" & Format(GetMedicareSum - dbl个帐合计, "0.00")
    strNone = Mid(strNone, 2)
    If strNone = "" Then 门诊预结算 = True
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub zl9InsureLedSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结Led报价
    '编制:刘兴洪
    '日期:2011-12-15 13:40:46
    '问题:44425
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double
    If Not gblnLED Then Exit Sub
    dbl个帐合计 = GetMedicareSum(mstr个人帐户)
    zl9LedVoice.DisplayBank "医保结算:", "帐户余额" & Format(mcur个帐余额, "0.00"), "帐户支付" & Format(dbl个帐合计, "0.00"), "统筹支付" & Format(GetMedicareSum - dbl个帐合计, "0.00")
    zl9LedVoice.Speak "#21 " & Format(Val(txt应缴.Text), "0.00")
End Sub

Private Sub cmd预结算_Click()
    Dim strNone As String
    Call AutoBultBookFee '收费时自动产生工本费
    
    If CheckBillsEmpty Then Exit Sub
    If gbytAutoSplitBill > 0 Then Call AutoSplitBill
                  
    If mbytInFun = 0 And mintInsure <> 0 And MCPAR.实时监控 Then
        '本来对于划价单才传2进行明细和汇总的检查，但是，由于以下原因，数量和实收金额在输入检查通过后可能改变，所以须再次检查明细
        '1.导入单据，2.修改单据，3.输入中药配方，4.修改中药付数后，其它行的付数同时变化，5.输入主项，自动产生从项，以及从项汇总计算折扣
        '6.修改单价，7.调整执行科室，药品价格重算，8.调整费别，实收金额重算,9.先输费用再验证医保身份,其它等等
        If gclsInsure.CheckItem(mintInsure, 0, 2, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0)) = False Then
            Exit Sub
        End If
    End If
    
    '预结算
    If Not 门诊预结算(strNone) Then
        If strNone <> "" Then
            MsgBox "当前保险结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                "在门诊未设置，请先到结算方式管理中设置这些结算方式！", vbInformation, gstrSysName
        End If
        cmd预结算.TabStop = True
        cmdOK.Enabled = False
        cmd预结算.SetFocus
        Exit Sub
    Else
    
    End If
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
End Sub

Private Sub Form_Activate()
    Dim objTemp As Object
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False: mblnNotClearLedDisplay = False
    If LoadBill = False Then Unload Me: Exit Sub
    If mbytInState = 5 Then cmdOK_Click: Exit Sub
    
    On Error Resume Next
    If mblnCopyBill Then
        cmdOK.SetFocus
    ElseIf mbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 0 And mstrInNO <> "" And Bill.Active Then
        Bill.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    End If
    
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If mbytInFun = 0 And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    DoEvents
    If mblnSetControl Then
        mblnSetControl = False
        Set objTemp = Me.ActiveControl
        If cboTemp.Visible And cboTemp.Enabled Then cboTemp.SetFocus
        If objTemp.Visible And objTemp.Enabled Then objTemp.SetFocus
    End If
    
    If mbln补费 Then
        Call Set病人补费编辑属性
        If Bill.Active Then Bill.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is Bill Then Exit Sub
    If InStr("',|~:：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr("`｀", Chr(KeyAscii)) > 0 Then
        '报请出示就诊卡
         KeyAscii = 0
        If gblnLED Then zl9LedVoice.Speak "#30"  '`为语音报价:有点奇怪:本来应该是192,但不知怎么会成229:32663
    End If
End Sub
Private Sub InitCommVariable()
    If Not mbln连续输入 Then
        mcurBill应缴 = 0
        mcurBill实收 = 0: mcurBill应收 = 0:
    End If
    mlng西药房 = 0: mlng成药房 = 0: mlng中药房 = 0
    mstr西窗 = "": mstr中窗 = "": mstr成窗 = ""
    mintBillNO = 0: mintMoneyRow = 0
    
    With mTyDelFee
        .strNos = ""
        Set .rsBlance = Nothing
        .blnSingleBalance = False
        .dblCurDelMoney = 0
        .bln三方卡全退 = False
    End With
End Sub

Private Sub InitBillColumnColor()
        Bill.SetColColor BillCol.类别, &HE7CFBA
        Bill.SetColColor BillCol.项目, &HE7CFBA
        Bill.SetColColor BillCol.数次, &HE7CFBA
        Bill.SetColColor BillCol.执行科室, &HE7CFBA
        Bill.SetColColor BillCol.付数, &HE0E0E0
        Bill.SetColColor BillCol.单价, &HE0E0E0
        Bill.SetColColor BillCol.标志, &HE0E0E0
End Sub

Private Sub ClearPayInfo()
    txt应缴.Text = "0.00"
End Sub

Private Sub ClearTotalInfo(Optional ByVal bln清除累计 As Boolean = False)
'默认bln为false,不清除累计,(划价时累计txtbox作为应缴显示)
    txt合计.Text = gstrDec: txt应收.Text = gstrDec
    If bln清除累计 Then
        If mbytInFun = 1 Then txt累计.Text = "0.00"
    End If
End Sub

Private Sub ClearPatientInfo(Optional ByVal bln清除病人 As Boolean = False)
'默认bln为false不清除病人txtbox
    If bln清除病人 Then
        mstrPrePati = ""
        mlngPrePati = 0
        mstrPreDoctor = ""
        txtPatient.Text = ""
        txtPatient.Locked = False
        txtPatient.BackColor = &HFFFFFF
        If mbytInFun = 2 Then lblCorp.Visible = False: lblCorp.Caption = ""
    End If
    txt年龄.Text = "": txt门诊号.Text = ""
    Call zlControl.CboLocate(cbo年龄单位, "岁")
    Call txt年龄_Validate(False)
    lbl险类.Caption = ""
End Sub

Private Sub ClearmobjBill()
    With mobjBill
        .姓名 = ""
        .性别 = ""
        .年龄 = ""
        .病人ID = 0
        .主页ID = 0
        .标识号 = 0
        .床号 = ""
        
        .病区ID = 0
        .科室ID = 0
        .婴儿费 = 0
        .费别 = zlStr.NeedName(cbo费别.Text)
        .门诊标志 = gint病人来源
        .加班标志 = chk加班.Value
    End With
End Sub
Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在关联数据
    '返回:如果不存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 16:49:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    

    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    mblnFirst = True: mbln连续输入 = False
    mblnHaveExcuteData = False
    mblnSetControl = True
    mblnStartFactUseType = zlStartFactUseType(1)
    '----------------------------界面初始-------------------------------------
    If glngSys Like "8??" Then
        lblPatient.Caption = "客户姓名"
        lbl费别.Caption = "会员等级"
        lbl门诊号.Caption = "客户号"
        lbl科室.Visible = False
        cbo开单科室.Visible = False
        lbl费别.Left = lblPatient.Left
        cbo费别.Left = txtPatient.Left
        cbo费别.Width = txtPatient.Width
        mshMoney.Visible = False
        fraStat.Left = mshMoney.Left
        vsBalance.Left = fraStat.Left + fraStat.Width + 30
        fra缴款.Left = vsBalance.Left + vsBalance.Width + 30
    End If
    
    '最小窗体尺寸
    glngFormW = 12000: glngFormH = 7710
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
        Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    '应该放在限制尺寸之后
    RestoreWinState Me, App.ProductName, mbytInFun & mbytInState
    sta.Visible = True
     
    If mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 3 Or mbytInState = 4 Or mbytInState = 5) Then
        If glng误差细目ID = 0 Then
            MsgBox "系统中尚未设置有效的误差处理项目。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    '----------------------------变量及对象初始化------------------------------
    'LED
    If mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 4 Or mbytInState = 5) And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModul, gcnOracle
    End If
    
    '问题:51510
    Call CreateDrugPacker '创建自动化药房部件
    mblnDrugPacker = False
    If mobjDrugPacker Is Nothing And mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Or mbytInState = 4) Then
        Err = 0: On Error Resume Next
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err <> 0 Then
            mblnDrugPacker = False
        Else
            mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
        End If
    End If
    
    Call ClearTotalInfo(True)
    lblSub应收.Caption = "应收:" & gstrDec
    lblSub实收.Caption = "实收:" & gstrDec
    lblAmount.Caption = ""
    
    '模块变量
    Call InitCommVariable
    
    gbln处方限量 = False
    gblnOK = False:         mblnLoad = False:           mblnDoing = False
    mblnDo = True:          mblnEnterCell = True:       mbln不重算价格 = False
    mblnCboClick = False
    mstrPrePati = "":       mlngPrePati = 0:            mstr付款方式 = ""
    mstr个人帐户 = "":      mblnValid = False:          mstrPreDoctor = ""
    mblnF2Save = False:     mblnAutoChangePati = False
    
    '单据对象
    mintPage = 1
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection    '该集合用于预结算,与单据标签保持一致
    mcolBalance.Add Array()
    Set mrsInfo = New ADODB.Recordset
    
    '-------------------------数据初始及加载------------------------------------
    '查看功能时，无需初始数据
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Or mbytInState = 4 Or mbytInState = 5 Then
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
    Call InitFace   'InitData需要在此之前
End Sub
Private Function LoadBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载单据数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 16:41:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Select Case mbytInState
    Case 0 'b.新增,修改
        If mbytInFun = 0 And mbytInState = 0 And gbln累计 Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
            txt累计.ToolTipText = "当前操作员今日收费累计额"
        End If
        '1.新增单据
        If Not NewBill(Not mblnStartFactUseType, False) Then Exit Function           '参数false表示不用再读取可用费别,因为前面InitData已做此操作
        '2.修改单据,多单据收费时，修改的是当前选中的那一张单据
        If mstrInNO <> "" Then
            Call LoadModifyNO(mstrInNO, IIf(mbytInFun = 2, 2, 1))
        Else
            If mlng病人ID <> 0 Then
                txtPatient.Text = "-" & mlng病人ID
                Call txtPatient_KeyPress(13)
            End If
        End If
        LoadBill = True: Exit Function
    Case 4, 5     '异常单据的处理
        If mstrInNO = "" Then Exit Function
        If LoadErrBillCharge(mstrInNO) = False Then Exit Function
        LoadBill = True: Exit Function
    Case 1, 2, 3    'a.显示、调整单据,记帐退费
        If mbytInState = 3 Then
            If Not ReadBill(mstrInNO, mbytInFun, True) Then Exit Function
        Else
            If Not ReadBill(mstrInNO, mbytInFun) Then Exit Function
        End If
        If InStr(mstrPrivs, "显示开单人") = 0 Then
            cbo开单人.Visible = False
            If gbyt科室医生 = 0 Then
                lbl科室.Visible = False
            Else
                lbl开单人.Visible = False
            End If
        End If
        cboNO.Text = mstrInNO
        LoadBill = True: Exit Function
    End Select
    LoadBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub Form_Resize()
    Dim lngCancelW As Long
    Dim lngLeft As Long
    On Error Resume Next
    
    fraTitle.Left = 0
    fraTitle.Width = Me.ScaleWidth
    
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    fraBill.Left = 0
    fraBill.Top = fraInfo.Top + fraInfo.Height
    fraBill.Width = Me.ScaleWidth
    cmdDelBill.Left = fraBill.Width - cmdDelBill.Width - 60
    cmdAddBill.Left = cmdDelBill.Left - cmdAddBill.Width
    tbsBill.Width = cmdAddBill.Left - tbsBill.Left - 300
    
    If fraBill.Visible Then
        Bill.Top = fraBill.Top + fraBill.Height
    Else
        Bill.Top = fraInfo.Top + fraInfo.Height
    End If
    Bill.Width = Me.ScaleWidth - Bill.Left
    If vsInvoice.Visible Then
        '25187
        With vsInvoice
            .Left = Bill.Left
            .Width = Bill.Width
        End With
        Call SetInvoceSizeAndShowTittle
    End If
    Bill.Height = Me.ScaleHeight - Bill.Top - sta.Height - picAppend.Height - IIf(fraSubBill.Visible, fraSubBill.Height + 30, 0) _
        - IIf(fra退费摘要.Visible, fra退费摘要.Height + 30, 0) _
        - IIf(vsInvoice.Visible, vsInvoice.Height + 30, 0)
    If fraSubBill.Visible Then
        fraSubBill.Left = Bill.Left
        fraSubBill.Width = Bill.Width
        fraSubBill.Top = Bill.Top + Bill.Height + 15
        lblSub实收.Left = fraSubBill.Width - 2250
        lblSub应收.Left = lblSub实收.Left - 2250
        lblAmount.Left = lblSub应收.Left - 2250
    End If
    If fra退费摘要.Visible Then
        With fra退费摘要
             .Left = Bill.Left
             .Width = Bill.Width
             .Top = Bill.Top + Bill.Height + 15
             txt退费摘要.Width = .Left + .Width - txt退费摘要.Left - 50
        End With
    End If
    '25187
    With vsInvoice
         .Top = IIf(fra退费摘要.Visible, fra退费摘要.Height + fra退费摘要.Top + 15, Bill.Top + Bill.Height + 15)
    End With
    
    cbo结算方式.Left = fraAppend.Left + lbl结算方式.Left + lbl结算方式.Width + 30
    cbo结算方式.Top = fraAppend.Top + lbl结算方式.Top - (cbo结算方式.Height - lbl结算方式) / 2
    
    cmdRegist.Left = fraTitle.Width - cmdRegist.Width - 90
    cmdIDCard.Left = fraTitle.Width - IIf(cmdRegist.Visible, cmdRegist.Width + 90, 0) - cmdIDCard.Width - 90
    
    lngLeft = fraTitle.Width - 90
    lngLeft = IIf(cmdRegist.Visible, cmdRegist.Left - 50, lngLeft)
    lngLeft = IIf(cmdIDCard.Visible, cmdIDCard.Left - 50, lngLeft)
    cmdSaveWholeSet.Left = lngLeft - cmdSaveWholeSet.Width
    lngLeft = IIf(cmdSaveWholeSet.Visible, cmdSaveWholeSet.Left - 50, lngLeft)
    cmdSelWholeSet.Left = lngLeft - cmdSelWholeSet.Width
    
    lngLeft = IIf(cmdSelWholeSet.Visible, cmdSelWholeSet.Left - 50, lngLeft)
    
    lblFormat.Left = lngLeft - lblFormat.Width
    'fraTitle.Width - IIf(cmdRegist.Visible, cmdRegist.Width + 90, 0) - IIf(cmdIDCard.Visible, cmdIDCard.Width + 90, 0) - lblFormat.Width - 90
    If cmdDelete.Visible Or chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    cmdDelete.Left = chkCancel.Left
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 30
    
    txtInvoice.Left = lblNO.Left - txtInvoice.Width - 40
    lblFact.Left = txtInvoice.Left - lblFact.Width - 40
    txtMCInvoice.Left = txtInvoice.Left
    
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
    If mbytInFun <> 0 Then
        cmdOK.Left = ScaleWidth - cmdOK.Width - 100
        cmdCancel.Left = cmdOK.Left
        cmdPrint.Left = cmdOK.Left
        cmd预结算.Left = cmdOK.Left
    End If
    If mbytInFun <> 2 Then
        If TypeName(cbo结算方式.Container) = TypeName(cbo开单人.Container) Then
            lbl开单人.Left = IIf(cbo结算方式.Visible, cbo结算方式.Left + cbo结算方式.Width + 100, lbl结算方式.Left)
            cbo开单人.Left = lbl开单人.Left + lbl开单人.Width + 20
        Else
             lbl开单人.Left = IIf(cbo结算方式.Visible, cbo结算方式.Left + cbo结算方式.Width + 100, lbl结算方式.Left)
            cbo开单科室.Left = lbl开单人.Left + lbl开单人.Width + 20
        End If
    End If
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytInFun = 0 And mbytInState = 0 And mstrYBPati <> "" And mstrInNO = "" Then
        If MsgBox("当前正在对医保病人收费，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        '取消医保病人身份验证,返回假时不退出
            Cancel = 1: Exit Sub
        End If
    End If
    
    SaveWinState Me, App.ProductName, mbytInFun & mbytInState
    If mbytInState = 0 Then
        Call SaveRegisterItem(g私有模块, Me.Name, "idkind", IDKind.IDKind)
    End If
    
    zlCommFun.OpenIme False
        
    mbytInFun = 0
    mbytInState = 0
    mblnCopyBill = False
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mbytBilling = 0
    mstrCardNO = ""
    mblnNOMoved = False   '查看时,可能传入true,
    mblnYB结算作废 = False
    
    mintBillNO = 0: mintMoneyRow = 0
    mlngFirstID = 0: mstrFirstWin = ""
    mlng领用ID = 0
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    
    mlng病人ID = 0
    mlng主页ID = 0
    mlngUnitID = 0
    mlngDeptID = 0
    mbln补费 = False
    mlng关联医嘱 = 0
    mstr最后转科时间 = ""
    
    '清空数据对象
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrs费别 = Nothing
    Set mrs费用类型 = Nothing
    Set mrs发药窗口 = Nothing
    Set mrsWarn = Nothing
    Set mobjCard = Nothing
    Set mobjBrushCheck = Nothing
    
    'LED初始化
    If mbytInFun = 0 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    mintInvoicePrint = 0
    If Not OS.IsDesinMode Then
        Call SetWindowLong(hWnd, GWL_WNDPROC, glngOld)
    End If
    
    If Not mobjDrugPacker Is Nothing Then
        '51510
        Set mobjDrugPacker = Nothing
    End If
    mblnHaveExcuteData = False
End Sub

Private Sub mobjBrushCheck_ReadCardNoed(ByVal strCardNo As String, ByVal blnBrushCard As Boolean)
    If blnBrushCard Then
        mbln条码刷卡 = True
    Else
        mbln条码刷卡 = False
    End If
End Sub

Private Sub mnuFileSavePrice_Click()
    '保存为划价单
    mnuFileSavePrice.Checked = True
    mblnSaveAsPrice = True
    
    Call DelFactMoney  '删除工本费
    Call cmdOK_Click
    If mnuFileSavePrice.Checked Then '检查中退出
        mnuFileSavePrice.Checked = False
        mblnSaveAsPrice = False
    End If
End Sub
Private Sub ReCalce退款()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算退款
    '编制:刘兴洪
    '日期:2011-11-21 17:27:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt应缴.Text = Format(GetDelMoney, "0.00")
End Sub
Private Sub ModiyVsBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改结算数据(目前只能修改(结算卡和消费卡)数据)
    '编制:刘兴洪
    '日期:2011-11-21 17:23:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not (mbytInState = 3 And mbytInFun = 0) Then Exit Sub
    With vsBalance
        '1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
        If InStr("3,4,5", .Cell(flexcpData, .Row, 0)) = 0 Then Exit Sub
        If .RowData(.Row) <> -1 Then Exit Sub
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Cell(flexcpForeColor, .Row, 0, .Row, .COLS - 1) = vbRed
        Else
            .Cell(flexcpForeColor, .Row, 0, .Row, .COLS - 1) = Me.ForeColor
        End If
    End With
    Call ReCalce退款
End Sub
Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Call ReCalce退款
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   If Not (mbytInState = 3 And mbytInFun = 0) Then Cancel = True: Exit Sub
    With vsBalance
        '1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
        If InStr("3,4,5", .Cell(flexcpData, .Row, 0)) = 0 Then Cancel = True: Exit Sub
'        If .RowData(.Row) <> -1 Or vsBalance.Tag = "1" Then Cancel = True: Exit Sub
'        .ColComboList(Col) = " ||" & Val(.Cell(flexcpData, Row, Col))
        If .RowData(.Row) <> -1 Then Cancel = True: Exit Sub '不退现
        If mTyDelFee.blnSingleBalance And mTyDelFee.bln三方卡全退 = False And .Cell(flexcpData, Row, 0) = 3 Then
            .ColComboList(Col) = " ||" & FormatEx(mTyDelFee.dblCurDelMoney, 2): Exit Sub
        End If
        If vsBalance.Tag = "1" Then Cancel = True: Exit Sub
        .ColComboList(Col) = " ||" & FormatEx(Val(.Cell(flexcpData, Row, Col)), 2)
    End With
End Sub

Private Sub vsBalance_DblClick()
    Dim lngRow As Long
    'Call ModiyVsBalance
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    
    If vsBalance.MouseCol <> 1 Then Exit Sub
    lngRow = vsBalance.MouseRow
    If vsBalance.RowData(lngRow) <> 0 And vsBalance.TextMatrix(lngRow, 0) <> "" Then
        txtTmp.Text = vsBalance.TextMatrix(lngRow, 1)
        txtTmp.SelStart = 0
        txtTmp.SelLength = Len(txtTmp.Text)
        txtTmp.ZOrder
        txtTmp.Left = vsBalance.Left + vsBalance.CellLeft
        txtTmp.Top = vsBalance.Top + vsBalance.CellTop + (vsBalance.CellHeight - txtTmp.Height) / 2 - 15
        txtTmp.Width = vsBalance.CellWidth - 30
        
        txtTmp.Visible = True
        txtTmp.SetFocus
    End If
End Sub

Private Sub vsBalance_EnterCell()
    If vsBalance.Col = 0 Then vsBalance.Col = 1
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    
    If vsBalance.RowData(vsBalance.Row) = 0 Then
        vsBalance.FocusRect = flexFocusLight
    Else
        vsBalance.FocusRect = flexFocusHeavy
    End If
End Sub

Private Sub vsBalance_GotFocus()
    vsBalance_EnterCell
End Sub

Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789", Chr(KeyAscii)) > 0 _
        And vsBalance.RowData(vsBalance.Row) <> 0 _
        And vsBalance.TextMatrix(vsBalance.Row, 0) <> "" _
        And (mbytInState = 0 And chkCancel.Value = 0) Then
        
        txtTmp.Text = Chr(KeyAscii)
        txtTmp.SelStart = Len(txtTmp.Text)
        txtTmp.SelLength = 0
        txtTmp.ZOrder
                    
        txtTmp.Left = vsBalance.Left + vsBalance.CellLeft
        txtTmp.Top = vsBalance.Top + vsBalance.CellTop + (vsBalance.CellHeight - txtTmp.Height) / 2 - 15
        txtTmp.Width = vsBalance.CellWidth - 30
        
        KeyAscii = 0
        
        txtTmp.Visible = True
        txtTmp.SetFocus
    End If
End Sub

Private Sub tbsBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtIn_GotFocus()
    Call zlControl.TxtSelAll(txtIn)
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
'功能:收费或划价时导入单据
    Dim lngPre As Long, strPre As String, strNo As String, strNos As String
    Dim intInsure As Integer, i As Long, j As Long
    Dim lng病人ID As Long, lng结帐ID As Long, bln急诊 As Boolean
    Dim strTmp As String
    Dim objBill As ExpenseBill
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
    Else
        KeyAscii = 0
        '导入单据
        txtIn.Text = GetFullNO(txtIn.Text, 13)
        Call zlControl.TxtSelAll(txtIn)
        strNo = txtIn.Text
               
        'a.单张单据模式,清除当前单据对象及病人信息
        If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then
            Call ClearFullBill(False)
            
            Set mobjBill = ImportBill(strNo, False, mbytInFun, , False, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
            If mobjBill.NO = "" Then
                MsgBox "读取单据失败。", vbInformation, gstrSysName
                txtIn.SetFocus: Exit Sub
            End If
            
            If InStr(mstrPrivs, "显示开单人") = 0 Then mobjBill.Pages(mintPage).开单人 = ""
            '清除病人信息
            Call ClearmobjBill
        Else
        'b.多张单据模块,新增单据,保留当前单据内容及病人相关信息,
        '不提供从后备表中导入的功能
            strNos = GetMultiNOs(strNo, , , True)
            For i = 0 To UBound(Split(strNos, ","))
                strNo = Replace(Split(strNos, ",")(i), "'", "")
                
                Set objBill = ImportBill(strNo, False, mbytInFun, , False, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                
                If objBill.NO = "" Then
                    MsgBox "读取单据失败。", vbInformation, gstrSysName
                    txtIn.SetFocus: Exit Sub
                End If
                
                If i > 0 Or mobjBill.Pages(mintPage).Details.Count > 0 Then
                    Call AddNewBill
                End If
                mintPage = tbsBill.Tabs.Count
                
                '不需要导入病人相关信息
                With mobjBill.Pages(mintPage)
                    .NO = "" '要清空以便修改时表明是直接输入的费用
                    .Key = objBill.Pages(1).Key
                    .保险金额 = objBill.Pages(1).保险金额
                    .冲预交额 = objBill.Pages(1).冲预交额
                    .煎法 = objBill.Pages(1).煎法
                    .进入统筹 = objBill.Pages(1).进入统筹
                    .开单部门ID = objBill.Pages(1).开单部门ID
                    If InStr(mstrPrivs, "显示开单人") > 0 Then .开单人 = objBill.Pages(1).开单人
                    .全自付 = objBill.Pages(1).全自付
                    .实收金额 = objBill.Pages(1).实收金额
                    .收费结算 = objBill.Pages(1).收费结算
                    .误差金额 = objBill.Pages(1).误差金额
                    .先自付 = objBill.Pages(1).先自付
                    .应缴金额 = objBill.Pages(1).应缴金额
                    .应收金额 = objBill.Pages(1).应收金额
                End With
                
                For j = 1 To objBill.Pages(1).Details.Count
                    With objBill.Pages(1).Details(j)
                        mobjBill.Pages(mintPage).Details.Add .费别, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, .InComes, , .保险项目否, .保险大类ID, .保险编码, .摘要
                    End With
                Next
            Next
            tbsBill.Tabs(mintPage).Selected = True  '不会引发click事件,因为mintpage=index
        End If
        
        Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
        Call LoadAndSeek费别
        
        '取第一药品行
        For i = 1 To mobjBill.Pages(1).Details.Count
            If InStr(",5,6,7,", mobjBill.Pages(1).Details(i).收费类别) > 0 Then
                mlngFirstID = mobjBill.Pages(1).Details(i).执行部门ID
                mstrFirstWin = mobjBill.Pages(1).Details(i).发药窗口
                Exit For
            End If
        Next
        
        Bill.Active = False
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
        Call InitBillColumnColor
        
        If IIf(mlngPrePati = 0, mstrPrePati <> mobjBill.姓名, mlngPrePati <> mobjBill.病人ID) Then
            '新病人
            mcurBill实收 = 0:  mcurBill应收 = 0: mcurBill应缴 = 0
            mintBillNO = 0: mintMoneyRow = 0
        End If
        
        '修改时应保存当前操作员的名字
        mobjBill.操作员编号 = UserInfo.编号
        mobjBill.操作员姓名 = UserInfo.姓名
        
        Call CalcMoneys     '因为不导入病人信息,所以需要根据当前的费别重算价格
        Call ShowDetails
        Call ShowMoney
                        
        txtIn.Text = ""
        'txt本次应缴.Visible = False:
        If mbytInState = 0 And mstrInNO <> "" Then txtModi.Text = "": mstrInNO = "": lbl应缴.Caption = "应缴"
        
        '要放在mstrInNO之后,因为以此来判断是否修改单据,以加回原库存
        Call CalcDrugStock
                    
        Bill.Active = True
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        
    End If
End Sub

Private Sub CalcDrugStock(Optional intPage As Integer)
    Dim i As Integer
    Dim str药房IDs As String
    '重新计算每行药品库存

     If intPage = 0 Then intPage = mintPage
     
     For i = 1 To mobjBill.Pages(intPage).Details.Count
        With mobjBill.Pages(intPage).Details(i)
            Bill.RowData(i) = Asc(.收费类别) '特殊处理
            
            If InStr(",5,6,7,", .收费类别) > 0 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + .原始数量
                
                Call SetItemRowColor(1, i)  '储备限额提示
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + .原始数量
                
                Call SetItemRowColor(1, i) '储备限额提示
            End If
        End With
    Next
End Sub

Private Sub txtInvoice_Change()
    lblFact.Tag = ""
End Sub

Private Sub txtInvoice_LostFocus()
    If Not (mbytInFun = 0 And mbytInState = 0) Then Exit Sub
    If txtInvoice.Text = "" Then
        Call RefreshFact
    End If
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
    
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
End Sub

Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTmp_Validate(Cancel As Boolean)
    Dim curOrig As Currency, curValue As Currency
    Dim curTotal As Currency, arrValue As Variant
    Dim p As Integer, i As Integer
    
    With vsBalance
        If Not IsNumeric(txtTmp.Text) Then
            Cancel = True
            MsgBox "输入了非法的""" & .TextMatrix(.Row, 0) & """结算金额！", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtTmp): Exit Sub
        Else
            '结算金额不允许超过返回的原始金额(个人帐户允许透支时再判断)
            curOrig = GetMedicareSum(.TextMatrix(.Row, 0), , True) '该结算方式所有原始返回金额和
            If (.TextMatrix(.Row, 0) <> mstr个人帐户 Or mcur个帐透支 = 0) _
                And Val(txtTmp.Text) > curOrig And Val(txtTmp.Text) <> 0 And curOrig <> 0 Then
                Cancel = True
                MsgBox "输入的""" & .TextMatrix(.Row, 0) & """结算金额不能超过 " & Format(curOrig, "0.00") & " ！", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtTmp): Exit Sub
            End If
            
            '个人帐户检查
            If .TextMatrix(.Row, 0) = mstr个人帐户 Then
                '不允许超过允许透支金额
                If mcur个帐余额 - Val(txtTmp.Text) < -1 * mcur个帐透支 Then
                    Cancel = True
                    MsgBox "帐户余额:" & Format(mcur个帐余额, "0.00") & _
                        IIf(mcur个帐透支 = 0, "", "(" & "允许透支:" & Format(mcur个帐透支, "0.00") & ")") & _
                        "不足要结算的金额。", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtTmp): Exit Sub
                End If
            End If
            
            '不允许超出单据剩余可结算金额
            curTotal = GetBillSum - Val(txt预交冲款.Text)
            For p = 1 To mcolBalance.Count
                For i = 0 To UBound(mcolBalance(p))
                    arrValue = Split(mcolBalance(p)(i), ";")
                    If arrValue(0) <> .TextMatrix(.Row, 0) Then
                        curTotal = curTotal - CCur(arrValue(3))
                    End If
                Next
            Next
            If Val(txtTmp.Text) > curTotal Then
                Cancel = True
                MsgBox "结算金额过大，超过单据允许结算金额:" & Format(curTotal, "0.00") & "。", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtTmp): Exit Sub
            End If
            
            .Text = Format(Val(txtTmp.Text), "0.00")
            txtTmp.Text = "": txtTmp.Visible = False
            
            '将修改后的金额分配到各张单据中
            '分配原则：从单据1开始,以不超过原始金额循环分配
            curValue = CCur(.Text)
            For p = 1 To mcolBalance.Count
                If .TextMatrix(.Row, 0) = mstr个人帐户 And mcur个帐透支 <> 0 Then
                    '允许透支的个人帐户,以不超过单据剩余可结算金额为准(不计冲预交,因为是后分配)
                    curOrig = GetBillSum(, CLng(p))
                    For i = 0 To UBound(mcolBalance(p))
                        arrValue = Split(mcolBalance(p)(i), ";")
                        If arrValue(0) <> .TextMatrix(.Row, 0) Then
                            curOrig = curOrig - CCur(arrValue(3))
                        End If
                    Next
                Else
                    curOrig = GetMedicareSum(.TextMatrix(.Row, 0), p, True)
                End If
                If curOrig <= curValue Then
                    Call SetBalanceVal(p, .TextMatrix(.Row, 0), curOrig)
                    curValue = curValue - curOrig
                Else
                    Call SetBalanceVal(p, .TextMatrix(.Row, 0), curValue)
                    curValue = 0
                End If
            Next
            
            '重新计算应缴，误差(分币)等:费用明细未变,全部不用重新计算
            Call ShowMoney(-1, Not (cmd预结算.Visible And cmdOK.Enabled))
        End If
    End With
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        mobjBill.发生时间 = CDate(txtDate.Text)
        If cmd预结算.Visible And cmd预结算.Enabled Then
            cmd预结算.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    Call zlControl.TxtSelAll(cboNO)

    If (mbytInFun = 0 And mbytInState = 0 And mobjBill.Pages(mintPage).Details.Count = 0) _
        Or chkCancel.Value = 1 Or mbytBilling = 2 Then
        cboNO.Locked = False '收费时，空单据可以提划价单，也可重复提取
    Else
        cboNO.Locked = True
    End If
    
    '收费时如果已验证医保病人身份,则禁止再读取划价单
    If mbytInFun = 0 And mbytInState = 0 And mstrYBPati <> "" Then cboNO.Locked = True
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, blnNull As Boolean, rsTmp As ADODB.Recordset
    Dim strOper As String, strNos As String, vDate As Date, intTmp As Integer
    Dim intInsure As Integer, blnHaveExe As Boolean, blnFlagPrint As Boolean
    Dim i As Integer, strErrMsg As String
    
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        If chkCancel.Value = 1 Then
            If mbytInFun = 0 Then '收费退费
                cboNO.Text = GetFullNO(cboNO.Text, 13)
            ElseIf mbytInFun = 2 Then '记帐销帐
                cboNO.Text = GetFullNO(cboNO.Text, 14)
            End If
        ElseIf mbytInFun = 2 Then
            '审核记帐
            cboNO.Text = GetFullNO(cboNO.Text, 14)
        ElseIf mbytInFun = 0 Then
            '划价收费
            cboNO.Text = GetFullNO(cboNO.Text, 13)
        End If
        
        If chkCancel.Value = 1 Then
            '1.收费时提划价单不会进入
            '2.药房划价不会进入
            '3.如果是门诊记帐划价新开进入、审核划价单则要排开，不必检查
            If mbytInFun <> 2 Or (mbytInFun = 2 And mbytBilling = 0) Then
                '是否已转入后备数据表中
                If zlDatabase.NOMoved("门诊费用记录", cboNO.Text, , IIf(mbytInFun = 2, "2", "1"), Me.Caption) Then
                    If Not ReturnMovedExes(cboNO.Text, IIf(mbytInFun = 2, "2", "1"), Me.Caption) Then cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    mblnNOMoved = False
                End If
            End If
        
            '多次审核或不完全审核的不允许销帐
            If mbytInFun = 2 Then
                If Not BillIdentical(cboNO.Text) Then
                    MsgBox "单据中包含部份不全完审核或分多次审核的内容，不允许在这里销帐。" & _
                        vbCrLf & "请退回管理界面过滤出相应的单据内容，然后再销帐。", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
        
            '单据退费权限判断
            If mbytInFun = 0 Then '收费
                If Not ReadBillInfo(1, cboNO.Text, 1, strOper, vDate) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
                If InStr(mstrPrivs, "所有操作员") <= 0 Then
                    If UserInfo.姓名 <> strOper Then
                        MsgBox "你没有""所有操作员""权限,不能对" & strOper & "的单据进行退费！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                If Not BillOperCheck(2, strOper, vDate, "退费", cboNO.Text, , 1) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            ElseIf mbytInFun = 2 Then '记帐
                If Not ReadBillInfo(1, cboNO.Text, 2, strOper, vDate) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
                If InStr(mstrPrivs, "所有操作员") <= 0 Then
                    If UserInfo.姓名 <> strOper Then
                        MsgBox "你没有""所有操作员""权限,不能对" & strOper & "的单据进行销帐！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                
                If Not BillOperCheck(4, strOper, vDate, "销帐", cboNO.Text, , 2) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            
            If mbytInFun = 0 Then '收费退费
                '检查是否异常单据
                If ChargeIsErrBill(cboNO.Text) Then
                    If MsgBox("单据：" & cboNO.Text & "的收费单据为异常收费单据,该单据只能作废或重新收费，是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                
                If gbln退费申请模式 Then
                    Set rsTmp = GetApply(cboNO.Text, 1)
                    rsTmp.Filter = "状态<>2"
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "请先对该单据进行退费申请！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    If IsNull(rsTmp!审核人) Then
                        MsgBox "该单据未进行退费审核，不能进行退费！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                
                strNos = GetMultiNOs(cboNO.Text, , mblnNOMoved, True)
                
                If gblnMultiBalance And InStr(strNos, ",") > 0 Then
                    If CheckSingleBalance(strNos) = False Then
                        MsgBox "多张单据使用多种结算方式模式下不允许对其中一张单据退费！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
            
                '医保类型匹配判断(确定时会再重复判断一次,因为还要获取其它医保参数)
                intInsure = ChargeExistInsure(strNos)
                If intInsure > 0 Then
                    '保险退费权限检查
                    If InStr(mstrPrivs, "保险收费") = 0 Then
                        MsgBox "你没有权限对医保病人的单据退费！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    
                    If InStr(strNos, ",") > 0 Then
                        If gclsInsure.GetCapability(support多单据收费必须全退, , intInsure) Then
                            MsgBox "医保要求一起收费的多张单据必须整体退费,请使用多单据退费模式！", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                        If gclsInsure.GetCapability(support多单据一次结算, , intInsure) Then
                            MsgBox " 多张单据一次交易必须整体退费,请使用多单据退费模式！", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                    End If
                Else
                    '是否有非医保病人的退费权限
                    If InStr(mstrPrivs, "允许非医保病人") = 0 Then
                        MsgBox "你没有权限对非医保病人进行退费操作！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                '退费
                With mTyDelFee
                    .strNos = strNos
                    Set .rsBlance = GetChargeBalance(strNos)
                End With
                '检查三方交易
                If InStr(1, strNos, ",") > 0 Then
                    With mTyDelFee.rsBlance
                        .Filter = "是否全退=1 And 是否退现=0"
                        If .RecordCount <> 0 Then .MoveFirst
                        strErrMsg = ""
                        Do While Not .EOF
                            strErrMsg = strErrMsg & vbCrLf & Nvl(!名称) & ":" & Format(Val(Nvl(!结算金额, 0)), "0.00")
                            .MoveNext
                        Loop
                        '问题:43734
                        If strErrMsg <> "" Then
                            MsgBox "存在以下三方交易不能进行部分退费，请检查！" & vbCrLf & strErrMsg, vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                    End With
                End If
            End If
            
            '是否已执行
            intTmp = BillCanDelete(cboNO.Text, IIf(mbytInFun = 2, 2, 1), blnHaveExe, , blnFlagPrint)
            If intTmp <> 0 Then
                Select Case intTmp
                    Case 1 '该单据不存在
                        MsgBox "指定的单据不存在！", vbInformation, gstrSysName
                    Case 2 '已经全部完全执行
                        '收费不考虑退费自动退药
                        MsgBox "该单据中的项目已经全部完全执行！", vbInformation, gstrSysName
                    Case 3 '未完全执行部分剩余数量为0
                        MsgBox "该单据中未完全执行部分项目剩余数量为零,没有可以" & IIf(mbytInFun = 2, "销帐", "退费") & "的费用！", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            ElseIf mbytInFun = 0 And intInsure > 0 And blnHaveExe Then '收费医保退费
                MsgBox "该医保收费单据中包含已经执行的项目,不能退费！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If blnHaveExe Then
                MsgBox "注意:该单据由于存在已执行的项目，当前将执行的是部分退费。", vbInformation, gstrSysName
            End If
            If blnFlagPrint Then
                If MsgBox("注意:检验医嘱的条码已打印，是否继续退费？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            '是否已结帐
            If mbytInFun = 2 Then
                If HaveBilling(1, cboNO.Text) <> 0 Then
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("该单据已经结帐,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            End If
                        Case 2
                            MsgBox "该单据已经结帐,不能销帐！", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End Select
                End If
            End If
        ElseIf mbytInFun = 2 And mobjBill.Pages(1).Details.Count = 0 Then
            '记帐划价单(记帐审核)
            If Not BillExistMoney("'" & cboNO.Text & "'", 2) Then
                MsgBox "单据费用已经全部销帐或单据不存在！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        ElseIf mbytInFun = 0 And chkCancel.Value = 0 Then
            '提取划价单收费
            If gblnCheckTest Then
                If Not CheckTest(cboNO.Text) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            
            '检查是否已提取该划价单
            For i = 1 To tbsBill.Tabs.Count
                If mobjBill.Pages(i).NO = cboNO.Text And i <> mintPage Then
                    MsgBox "该张划价单已经在第 " & i & " 张单据中输入。", vbInformation, gstrSysName
                    cboNO.Text = mobjBill.Pages(mintPage).NO: cboNO.SetFocus: Exit Sub
                End If
            Next
        End If
        
        
        Call ClearPayInfo
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        '不是修改时,mstrInNO为提取的退费单,审核单,但不含划价单
        If Not (mbytInFun = 0 And chkCancel.Value = 0) Then mstrInNO = UCase(cboNO.Text)
        
        If chkCancel.Value = 1 Then '读取退费单
            blnRead = ReadBill(cboNO.Text, mbytInFun, True)
        ElseIf mbytInFun = 2 Then '读取记帐划价单
            blnRead = ReadBill(cboNO.Text, 2, False, blnNull)
        ElseIf mbytInFun = 0 Then '读取收费划价单
            blnRead = ReadBill(cboNO.Text, 1, False, blnNull)
        End If
        If blnRead Then
            If chkCancel.Value = 0 Then '收费或记帐的划价单
                Bill.Active = False
                chk加班.Enabled = False
                
                '如果没有权限，提取划价单后,只能输入医保病人
                If gint病人来源 = 1 And mbytInFun = 0 And InStr(mstrPrivs, "允许非医保病人") = 0 Then
                     ClearPatientInfo (True)
                End If
                
                '如果是挂号产生临时病人姓名模式,则读取病人身份信息,以便修改
                If mbytInFun = 0 And txtPatient.Text = "新病人" Then
                    Call GetPatient("-" & mobjBill.病人ID)
                End If
                
                '显示摘要
                Call Bill_EnterCell(1, BillCol.项目)
                
                If mbytInFun = 0 And txtPatient.Text <> "新病人" Then
                    If Not CheckRegisted(mobjBill.病人ID) Then
                        Call ClearFullBill(False)
                        Exit Sub
                    End If
            
                    '自动加收挂号费
                    Call LoadAddedItem(mobjBill.病人ID, mobjBill.姓名)
                    
                    '划价单收费时报LED
                    If tbsBill.Tabs.Count = 1 Then Call ShowWelcomeByLed
                End If
                Call ReInitPatiInvoice '97160
                
                '光标定位
                If mbytInFun = 2 Then
                    cmdOK.SetFocus
                ElseIf txtPatient.Text = "" Or blnNull Then
                    txtPatient.SetFocus
                Else
'                    If txt缴款.Enabled And txt缴款.Visible Then
'                        txt缴款.SetFocus
'                    Else
                    If cmd预结算.Enabled And cmd预结算.Visible Then
                        cmd预结算.SetFocus
                    ElseIf cmdOK.Enabled And cmdOK.Visible Then
                        cmdOK.SetFocus
                    End If
                End If
            Else '退
                Call SetDisible 'cboNO在获取焦点后unLock
                If mbytInFun = 0 Then
                    '部份退费只支持退费指定结算方式
                    cbo结算方式.Enabled = True
                    cbo结算方式.Locked = False
                End If
                Bill.Active = True
                cmdOK.SetFocus
            End If
        Else
            If Not (mbytInFun = 0 And chkCancel.Value = 0) Then mstrInNO = ""
            cboNO.Text = ""
            If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        End If
    End If
End Sub

Private Sub txt门诊号_GotFocus()
    zlControl.TxtSelAll txt门诊号
End Sub

Private Sub txt退费摘要_Change()
    txt退费摘要.Tag = ""
End Sub


Private Sub txt退费摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    '选择退费原因
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    If Trim(txt退费摘要.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt退费摘要.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt退费摘要, Trim(txt退费摘要.Text), "常用退费原因", "常用退费原因选择", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txt退费摘要.Text)) = False Then
            Exit Sub
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
End Sub
Private Sub txt退费摘要_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt退费摘要
End Sub
Private Sub txt退费摘要_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt应缴_Change()
    If mbytInFun <> 0 Then Exit Sub
'        If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00": txt找补.Text = "0.00": Exit Sub
'        txt找补.Text = Format(Val(txt缴款.Text) - Val(txt应缴.Text), "0.00")
End Sub

Private Sub txt预交冲款_GotFocus()
    Call AutoBultBookFee '收费自动产生工本费
    zlControl.TxtSelAll txt预交冲款
    txt预交冲款.Tag = txt预交冲款.Text
End Sub

Private Sub txt预交冲款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt预交冲款_Validate(Cancel As Boolean)
    Dim curTotal As Currency
        
    curTotal = GetBillSum
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
            txt预交冲款.Text = Format(IIf(curTotal - GetMedicareSum > Val(sta.Panels(Pan.C4预交信息).Tag), Val(sta.Panels(Pan.C4预交信息).Tag), curTotal - GetMedicareSum), "0.00")
        End If
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) > 0 And curTotal < 0 Then
        MsgBox "单据应付金额为负时不能使用预交款！", vbInformation, gstrSysName
        txt预交冲款.Text = "0.00"
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) > Val(sta.Panels(Pan.C4预交信息).Tag) Then
        MsgBox "预交款冲款金额不能超过病人的预交余额:" & Format(Val(sta.Panels(Pan.C4预交信息).Tag), "0.00") & " ！", vbInformation, gstrSysName
        If curTotal < 0 Then
            txt预交冲款.Text = "0.00"
        Else
            txt预交冲款.Text = Format(IIf(curTotal - GetMedicareSum > Val(sta.Panels(Pan.C4预交信息).Tag), Val(sta.Panels(Pan.C4预交信息).Tag), curTotal - GetMedicareSum), "0.00")
        End If
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    ElseIf Val(txt预交冲款.Text) > Format(curTotal - GetMedicareSum, "0.00") And Val(txt预交冲款.Text) <> 0 Then
        MsgBox "预交款冲款金额不能大于应付金额:" & Format(curTotal - GetMedicareSum, "0.00") & " ！", vbInformation, gstrSysName
        If curTotal < 0 Then
            txt预交冲款.Text = "0.00"
        Else
            txt预交冲款.Text = Format(IIf(curTotal - GetMedicareSum > Val(sta.Panels(Pan.C4预交信息).Tag), Val(sta.Panels(Pan.C4预交信息).Tag), curTotal - GetMedicareSum), "0.00")
        End If
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    Else
        txt预交冲款.Text = Format(txt预交冲款.Text, "0.00")
    End If
    
    If Val(txt预交冲款.Tag) = Val(txt预交冲款.Text) Then Exit Sub
    
    '重新计算应缴，误差(分币)等:费用明细未变,全部不用重新计算
    Call ShowMoney(-1, Not (cmd预结算.Visible And cmdOK.Enabled))
End Sub

Private Sub txtInvoice_GotFocus()
    zlControl.TxtSelAll txtInvoice
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

Private Sub txtModi_LostFocus()
    If mstrInNO <> "" And txtModi.Text <> mstrInNO Then txtModi.Text = mstrInNO
End Sub

Private Sub txt年龄_Gotfocus()
    Call zlCommFun.OpenIme
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
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPatient_GotFocus()
    If mbytInFun = 0 Or mbytInFun = 1 Then
        zlControl.TxtSelAll txtPatient
        zlCommFun.OpenIme True
    Else
        zlControl.TxtSelAll txtPatient
    End If
    
    'LED语音报价
    If mbytInFun = 0 And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.Speak "#51" '请问你的姓名
    End If
    
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard (True)
    End If
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long

    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        If mbytInState <> 2 Then
            .ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus) '类别列,当主从项时会被改变
            .ColData(BillCol.项目) = BillColType.CommandButton  '项目列,当主从项时会被改变
            .ColData(BillCol.付数) = BillColType.UnFocus  '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = BillColType.UnFocus  '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.标志) = BillColType.UnFocus  '标志缺省跳过,当为手术时,设为复选(-1)
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
        If Visible And Bill.Active And Row > 0 And .ColData(BillCol.类别) <> BillColType.UnFocus And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboSex.ListIndex <> -1 Then mobjBill.性别 = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If cboSex.Locked Then Exit Sub
    If SendMessage(cboSex.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 17 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo费别.Locked Then Exit Sub
    
    If KeyAscii >= 32 Then
        If cbo费别.Locked Then Exit Sub
    
        lngIdx = zlControl.CboMatchIndex(cbo费别.hWnd, KeyAscii)
        If lngIdx = -1 And cbo费别.ListCount > 0 Then lngIdx = 0
        cbo费别.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 Then
        If cbo费别.ListIndex = -1 Then
            mobjBill.费别 = ""
        Else
             '即使费用相同也要重算,因为医保验卡后必须重算,预结算才正确
            If (mstrYBPati <> "" Or mobjBill.费别 <> zlStr.NeedName(cbo费别.Text)) Then
                mobjBill.费别 = zlStr.NeedName(cbo费别.Text)
                If mbytInState = 0 And Not CheckBillsEmpty Then
                    '需要重新预结算
                    If cmd预结算.Visible Then
                        Call InitBalanceGrid
                        cmd预结算.TabStop = True
                        cmdOK.Enabled = False
                    End If
                    
                    '全部重新计算价格
                    Call CalcMoneys
                    Call ShowDetails
                    Call ShowMoney
                End If
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    '刘兴洪 问题:27378 日期:2010-01-27 13:35:37
    If KeyAscii <> 13 Then Exit Sub
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo开单人.ListIndex >= 0 Then lng医生ID = cbo开单人.ItemData(cbo开单人.ListIndex)
    If mrs开单科室 Is Nothing Then Call FillDept(mlngDeptID, lng医生ID)
    If zlSelectDept(Me, mlngModul, cbo开单科室, mrs开单科室, cbo开单科室.Text) = False Then KeyAscii = 0: Exit Sub
    Exit Sub
'
'    If KeyAscii = 13 Then
'
'        mblnCboClick = False    '先用鼠标在下拉列表选择一个并点击,不要移开,此时只触发click,再输入简码并且回车,不触发click,所以需要在此赋值,以便validate事件中强行调用click事件
'        Call zlCommfun.PressKey(vbKeyTab)
'    ElseIf KeyAscii >= 32 And Not cbo开单科室.Locked Then
'        lngIdx = zlControl.CboMatchIndex(cbo开单科室.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo开单科室.ListCount > 0 Then lngIdx = 0
'        cbo开单科室.ListIndex = lngIdx
'    End If
End Sub

Private Function isCheck开单人Exists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo开单人.ListCount - 1
        If zlStr.NeedName(cbo开单人.List(i)) = str姓名 Then
            If blnLocateItem Then cbo开单人.ListIndex = i
            isCheck开单人Exists = True
            Exit Function
        End If
    Next
End Function

Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    Dim strAdded As String
    If KeyAscii = 13 Then
        If cbo开单人.Locked Then
            If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        strText = UCase(cbo开单人.Text)
        If cbo开单人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> UCase(cbo开单人.List(cbo开单人.ListIndex)) Then Call zlControl.CboSetIndex(cbo开单人.hWnd, -1)
        End If
        If strText = "" Then
            cbo开单人.ListIndex = -1
        ElseIf cbo开单人.ListIndex = -1 Then
            intIdx = -1
            strFilter = IIf(gbln护士, "人员性质<>''", "人员性质<>'护士'")
            
            '刘兴洪:22383
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs开单人)
            Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
            Dim strCompents As String '匹配串
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs开单人.Filter = strFilter: iCount = 0
            With mrs开单人
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs开单人.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '如果输入的数字,需要检查:
                        '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                        If Nvl(!编号) = strText Then strResult = Nvl(!姓名): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(Nvl(!编号)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!姓名)
                            iCount = iCount + 1
                        End If
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(mrs开单人!编号) Like strText & "*" Then
                            If isCheck开单人Exists(Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                                strAdded = strAdded & "," & Nvl(!编号) & ","
                            End If
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(Nvl(!简码)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同简码
                            iCount = iCount + 1
                        End If
                        '2.根据参数来匹配相同数据
                        If Trim(Nvl(!简码)) Like strCompents Then
                            If isCheck开单人Exists(Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                                strAdded = strAdded & "," & Nvl(!编号) & ","
                            End If
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                            If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!编号) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!姓名)) Like strCompents Then
                            If isCheck开单人Exists(Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                                strAdded = strAdded & "," & Nvl(!编号) & ","
                            End If
                        End If
                    End Select
                    mrs开单人.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!姓名)
            '刘兴洪:直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck开单人Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                Select Case intInputType
                Case 0 '输入全数字
                    rsTemp.Sort = "编号"
                Case 1 '输入全拼音
                    rsTemp.Sort = "简码"
                Case Else
                    '根据选择来定
                    If gbyt开单人显示 = 1 Then '简码
                        rsTemp.Sort = "简码"
                    Else
                        rsTemp.Sort = "编号"
                    End If
                End Select
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, mlngModul, cbo开单人, rsTemp, True, "", "缺省,职务,优先级别", rsReturn) Then
                    If cbo开单人.Enabled Then cbo开单人.SetFocus
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheck开单人Exists(Nvl(rsReturn!姓名), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo开单人: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
            
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo开单人_Click
            If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo开单人.ListIndex = -1 Then
            cbo开单人.Text = ""
            mobjBill.Pages(mintPage).开单人 = ""
            lblDuty.Caption = ""
            If gbyt科室医生 = 0 Or gbln必须输开单人 Then Exit Sub
        Else
            mobjBill.Pages(mintPage).开单人 = zlStr.NeedName(cbo开单人.Text)
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo开单人_Click
            ElseIf intIdx <> cbo开单人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo开单人.SetFocus
                If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo开单人_Click
            End If
        End If
        If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub ShowCHRecipe()
'功能：调用中药配方输入功能
    Dim objDetails As BillDetails
    Dim str动态费别 As String, lng病人科室ID As Long
    Dim int序号 As Integer, i As Long
    
    If Not (Bill.Active And mbytInState = 0) Then Exit Sub
    
    '检查是否有非中药
    For i = 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).收费类别 <> "7" _
            And Not mobjBill.Pages(mintPage).Details(i).工本费 Then
            Call MsgBox("在当前单据中存在不是中草药的收费项目，请删除非中草药收费项目后,再进入配方!", vbInformation + vbDefaultButton1, gstrSysName)
             
            If cmd配方.Enabled And cmd配方.Visible Then cmd配方.SetFocus
            Exit Sub
        End If
    Next
    
    '病人科室或开单科室ID
    lng病人科室ID = mobjBill.科室ID
    If lng病人科室ID = 0 Then lng病人科室ID = Get开单科室ID
        
    '动态费别串
    If glngSys Like "8??" Or mbytInFun = 2 Then
        str动态费别 = zlStr.NeedName(cbo费别.Text)
    Else
        str动态费别 = zlStr.TrimEx(zlStr.NeedName(cbo费别.Text) & "," & lbl动态费别.Tag, ",")
    End If
    
    '调用窗口
    Set objDetails = frmCHRecipe.ShowMe(Me, mstrPrivs, mlngModul, mbytInFun, mbytBilling, Original.实收合计, mobjBill.病人ID, lng病人科室ID, Get开单科室ID, _
        IIf(mlng中药房 = 0, glng中药房, mlng中药房), mobjBill.Pages(mintPage).Details, zlStr.NeedName(cbo费别.Text), str动态费别, _
         IIf(mstrYBPati <> "", mintInsure, 0), chk加班.Value = 1, mobjBill.Pages(mintPage).煎法, mrsWarn, mcolStock1, zl获取中药形态(mintPage, Bill.Row, True))
    If Not objDetails Is Nothing Then
        '清除原单据中的中草药
        For i = mobjBill.Pages(mintPage).Details.Count To 1 Step -1
            If mobjBill.Pages(mintPage).Details(i).收费类别 = "7" Then
                mobjBill.Pages(mintPage).Details.Remove i
            End If
        Next
        '添加编辑后的中草药
        For i = 1 To objDetails.Count
            With objDetails(i)
                int序号 = mobjBill.Pages(mintPage).Details.Count + 1
                Call mobjBill.Pages(mintPage).Details.Add(.费别, .Detail, .收费细目ID, int序号, .从属父号, _
                    .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, _
                    .InComes, "", .保险项目否, .保险大类ID, .保险编码, .摘要, .原始数量, .原始执行部门ID)
            End With
        Next
        
        '更新中药煎法
        mobjBill.Pages(mintPage).煎法 = frmCHRecipe.mstr煎法
        '刷新当前单据中的显示
        Call ClearBillRows
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
        
        Call InitBillColumnColor
        
        '在重新计算之前清除
        If cmd预结算.Visible Then
            Call InitBalanceGrid
            cmd预结算.TabStop = True
            cmdOK.Enabled = False
        End If

        Call ShowDetails
        Call ShowMoney(mintPage)
        Call SetColNum
                
        Call CalcDrugStock
        Call SetBill中草药EditEnabled
        
        Bill.Col = BillCol.项目: Bill.CmdVisible = False  '不然定位不起
        If cmd预结算.Enabled And cmd预结算.Visible Then
            cmd预结算.SetFocus
'        ElseIf txt缴款.Enabled And txt缴款.Visible Then
'            txt缴款.SetFocus
        ElseIf cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus
        End If
    Else
        Bill.SetFocus
    End If
End Sub

Private Sub ApportionMultiBalance(ByVal strBalance As String, ByVal curError As Currency)
    Dim i As Long, j As Long
    Dim cur当前结算 As Currency, cur当前未缴 As Currency, arrPay As Variant
    Dim varData As Variant, str退支票(0 To 3) As String, bln应付款 As Boolean
    Dim bln存在退支票 As Boolean
    
    With mobjBill
        For i = 1 To mobjBill.Pages.Count
            .Pages(i).应缴金额 = 0      '这种模式下的应缴实际上没有使用,仅在下面的程序中用于判断,最后,累计的应缴可能有小数
            .Pages(i).收费结算 = ""
            If i = .Pages.Count Then
                .Pages(i).误差金额 = curError
            Else
                .Pages(i).误差金额 = 0
            End If
        Next
        '按单据顺序分摊(医保和预交款已在之前分摊,如果重输预交款,会清除分摊信息)
        arrPay = Split(strBalance, "||")
        ':33722
        bln存在退支票 = False
        For j = 0 To UBound(arrPay)
            varData = Split(arrPay(j), "|")
            mrs结算方式.Filter = "名称='" & varData(0) & "'"
            If Not mrs结算方式.EOF Then
                bln应付款 = Val(Nvl(mrs结算方式!应付款)) = 1
                If bln应付款 Then
                    str退支票(0) = varData(0)
                    str退支票(1) = varData(1)
                    str退支票(2) = varData(2)
                    str退支票(3) = varData(3)
                    bln存在退支票 = True
                    Exit For
                End If
            End If
        Next
        
        For j = 0 To UBound(arrPay)
            varData = Split(arrPay(j), "|")
            cur当前结算 = varData(1) '结算方式|结算金额|结算号码|摘要||......
            ':33722
            If str退支票(0) = varData(0) And bln存在退支票 Then
                   bln应付款 = True
            Else
                bln应付款 = False
            End If
            
            If bln应付款 = False Then
                For i = 1 To mobjBill.Pages.Count
                    cur当前未缴 = .Pages(i).实收金额 + .Pages(i).误差金额 - .Pages(i).保险金额 - .Pages(i).冲预交额 - .Pages(i).应缴金额 - .Pages(i).消费卡刷卡额
                                            
                    If cur当前未缴 > 0 Then
                        If cur当前未缴 <= cur当前结算 Then
                            '支票的处理:33722
                            '可能存在应退支票这种情况,可能到了最后一张单据,支票都还未分配完的情况
                            '这时,将余下部分直接分配给最后一张单据
                            If i = mobjBill.Pages.Count And varData(0) Like "*支票*" And bln存在退支票 Then
                                  .Pages(i).收费结算 = .Pages(i).收费结算 & "||" & varData(0) & "|" & cur当前结算 & "|" & varData(2) & "|" & varData(3)
                                  .Pages(i).应缴金额 = .Pages(i).应缴金额 + cur当前结算
                                  cur当前结算 = 0: Exit For
                            Else
                            .Pages(i).收费结算 = .Pages(i).收费结算 & "||" & _
                                               varData(0) & "|" & cur当前未缴 & "|" & varData(2) & "|" & varData(3)
                            .Pages(i).应缴金额 = .Pages(i).应缴金额 + cur当前未缴
                            cur当前结算 = cur当前结算 - cur当前未缴
                            End If
                        Else
                            .Pages(i).收费结算 = .Pages(i).收费结算 & "||" & _
                                               varData(0) & "|" & cur当前结算 & "|" & varData(2) & "|" & varData(3)
                            .Pages(i).应缴金额 = .Pages(i).应缴金额 + cur当前结算
                            cur当前结算 = 0
                        End If
                        If cur当前结算 = 0 Then Exit For
                    End If
                Next
            End If
        Next
        
        If str退支票(0) <> "" And bln存在退支票 Then
            '退支票部分,只能放在最后一张
            .Pages(mobjBill.Pages.Count).收费结算 = .Pages(mobjBill.Pages.Count).收费结算 & "||" & _
                                               str退支票(0) & "|" & Val(str退支票(1)) & "|" & str退支票(2) & "|" & str退支票(3)
             .Pages(mobjBill.Pages.Count).应缴金额 = .Pages(mobjBill.Pages.Count).应缴金额 + RoundEx(Val(str退支票(1)), 2)
        End If
        For i = 1 To mobjBill.Pages.Count
            If Mid(.Pages(i).收费结算, 1, 2) = "||" Then .Pages(i).收费结算 = Mid(.Pages(i).收费结算, 3)
        Next
    End With
    mrs结算方式.Filter = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '参数：Shift=-1：表示是程序强行在调用
    Select Case KeyCode
        Case vbKeyF1  '帮助
            Select Case mbytInFun
                Case 0
                    ShowHelp App.ProductName, Me.hWnd, Me.Name & "2"
                Case 1
                    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
                Case 2
                    ShowHelp App.ProductName, Me.hWnd, Me.Name & "3"
            End Select
        Case vbKeyF2
            If Shift = vbCtrlMask Then
                If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gbytAutoSplitBill > 0 Then
                    Call AutoSplitBill
                End If
            Else
                mblnF2Save = True
                    If ActiveControl Is txtPatient And mbytInFun <> 1 Then
                        Call txtPatient_LostFocus
                        Call txtPatient_Validate(False)
                        Me.Refresh
                    End If
                    If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
                mblnF2Save = False
                If cmdOK.Enabled And cmdOK.Visible Then
                    Call cmdOK.SetFocus
                    Call cmdOK_Click
                End If
            End If
        Case vbKeyF3 '挂号
            If cmdRegist.Visible And cmdRegist.Enabled Then
                cmdRegist.SetFocus
                Call cmdRegist_Click
            End If
        Case vbKeyUp
'            '刘兴洪:27498 20100117
'            If Me.ActiveControl Is txtPatient Then
'                Call IDKind.Locale(-1)
'                'IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
'            End If
        Case vbKeyDown
'            '刘兴洪:27498 20100117
'            If Me.ActiveControl Is txtPatient Then
'                Call IDKind.Locale
'                'IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
'            End If
        Case vbKeyF4 '多种方式结算
            If Shift = vbCtrlMask Then
                If IDKind.Enabled And txtPatient.Locked = False And txtPatient.Enabled Then
                    Dim intIndex As Integer
                    intIndex = IDKind.GetKindIndex("IC卡号")
                    If intIndex <= 0 Then Exit Sub
                    IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
                End If
            End If
        Case vbKeyF5
            If cmd预结算.Visible And cmd预结算.Enabled Then cmd预结算.SetFocus: cmd预结算_Click
        Case vbKeyF6 '定位到病人输入框
            If Me.ActiveControl Is txtPatient And txtPatient.Enabled And mstrYBPati = "" Then   '提取划价单后，姓名输入框是锁定的
                '70143:刘尔旋,2014-3-3,住院病人医保验证
                If mbytInFun = 0 And mbytInState = 0 And (gint病人来源 = 1 Or gint病人来源 = 2) Then
                    If chkCancel.Value = 0 And InStr(mstrPrivs, "保险收费") > 0 Then
                        Dim lngCur病人ID As Long
                        If mrsInfo.State = 1 Then
                            If txtPatient.Text = mrsInfo!姓名 Then lngCur病人ID = mrsInfo!病人ID
                        Else
                            If txtPatient.Text = mobjBill.姓名 Then lngCur病人ID = mobjBill.病人ID  '问题:25486
                        End If
                        Call MCPatientProcess(lngCur病人ID)
                    End If
                End If
            Else
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            End If
        Case vbKeyF7 '切换输入法
            If Shift = vbCtrlMask Then
                If sta.Panels("PatiSource").Visible Then
                    Call sta_PanelClick(sta.Panels("PatiSource"))
                End If
            Else
                If Not gbln简码切换 Then Exit Sub
                If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                    If sta.Panels("WB").Bevel = sbrRaised Then
                        Call sta_PanelClick(sta.Panels("WB"))
                    Else
                        Call sta_PanelClick(sta.Panels("PY"))
                    End If
                End If
            End If
        Case vbKeyF8 '退(自动激活事件)
            If mbytInFun = 1 Then
                cmdCancel.SetFocus
                Call cmdCancel_Click
            Else
                If chkCancel.Visible And chkCancel.Enabled Then
                    chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
                ElseIf cmdDelete.Visible And cmdDelete.Enabled Then
                    cmdDelete.SetFocus: Call cmdDelete_Click
                End If
            End If
        Case vbKeyF9 '定位到单据号输入框
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Case vbKeyF10 '就诊卡发放
            If cmdIDCard.Visible And cmdIDCard.Enabled Then cmdIDCard.SetFocus: cmdIDCard_Click
        Case vbKeyF11
            If cmd配方.Enabled And cmd配方.Visible Then cmd配方.SetFocus: Call cmd配方_Click
        Case vbKeyF12
            If Shift = vbCtrlMask Then
''                '强制性LED报价,(合计)
''                If mbytInFun = 0 And gblnLED And mbytInState = 0 _
''                    And txt缴款.Enabled And txt缴款.Visible And CCur(txt合计.Text) <> 0 Then
''                    mblnHotKey = True: txt缴款.SetFocus
''                    If ActiveControl Is txt缴款 Then AutoBultBookFee
''                End If
            ElseIf Shift = vbAltMask Then
                Call sta_PanelClick(sta.Panels("Drugstore"))
            Else
                '问题:27939
                If Me.ActiveControl Is txtPatient Then
                    Call txtPatient_Validate(False)
                End If
                '增加单据
                If cmdAddBill.Enabled And cmdAddBill.Visible Then cmdAddBill.SetFocus: Call cmdAddBill_Click
            End If
        Case vbKeyS
            '保存为划价单
            If Shift = vbCtrlMask Then
                If CheckSaveMultiPrice Then
                    Call mnuFileSavePrice_Click
                Else
                    MsgBox "仅在收费时允许保存为划价单." & vbCrLf & "如果是多张单据收费,要求不含导入的单据", vbInformation, gstrSysName
                End If
            End If
        Case vbKeyA, vbKeyR
            '全选，全清
            If Shift = vbCtrlMask Then
                If KeyCode = vbKeyA Then
                    Call SelALLRow
                ElseIf KeyCode = vbKeyR Then
                    Call ClearALLRow
                End If
            End If
        Case vbKeyD
            If Shift = vbCtrlMask Then
                If sta.Panels(Pan.C4预交信息).Visible And mrsInfo.State = 1 Then
                    Call ShowDeposit(mrsInfo!病人ID)
                End If
            End If
        Case vbKeyF 'Ctrl+F定位缴款输入框
'            If Shift = vbCtrlMask Then
'                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
'            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False
                Bill.SetFocus
            ElseIf txtTmp.Visible Then
                txtTmp.Visible = False
                If vsBalance.Enabled Then vsBalance.SetFocus
            Else
                cmdCancel.SetFocus: Call cmdCancel_Click
            End If
        Case 191 '"?"计算器
            If Shift = vbAltMask Then
                Call sta_PanelClick(sta.Panels("Calc"))
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
    
    If mbytInFun = 1 Then
        mshMoney.ColWidth(1) = lngW * 0.4
        mshMoney.ColWidth(2) = lngW * 0.3
        mshMoney.ColWidth(3) = lngW * 0.3
    Else
        mshMoney.ColWidth(1) = lngW * 0.45
        mshMoney.ColWidth(2) = lngW * 0.55
        mshMoney.ColWidth(3) = 0
    End If
    
    mshMoney.ColAlignment(0) = 4
    mshMoney.ColAlignment(1) = 1
    mshMoney.ColAlignment(2) = 7
    mshMoney.ColAlignment(3) = 7
    
    mshMoney.TextMatrix(0, 0) = "序号"
    mshMoney.TextMatrix(0, 1) = "项目"
    mshMoney.TextMatrix(0, 2) = "金额"
    mshMoney.TextMatrix(0, 3) = "合计"
    mshMoney.Row = 0
    mshMoney.Col = 0: mshMoney.CellAlignment = 4
    mshMoney.Col = 1: mshMoney.CellAlignment = 4
    mshMoney.Col = 2: mshMoney.CellAlignment = 4
    mshMoney.Col = 3: mshMoney.CellAlignment = 4
    
    mshMoney.MergeCol(0) = True
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim Curdate As Date     '服务器当前时间
    
    On Error GoTo errH
        
    '初始化病人信息对象
    Set mrsInfo = New ADODB.Recordset
    '查看时,不支持身份证识别,修改时要支持,因为修改后可能继续新单收费
    If mbytInState = 0 Then
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.hWnd)
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    
    '刘兴洪:结算卡的一些处理
    Call initCardSquareData
    
    '年龄单位
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    
    
    '------------------批量读取------------------
    
    '可选性别,医疗付款方式,结算方式
    strSQL = " Select '性别' as 类别,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Union All " & _
             " Select '医疗付款方式' as 类别,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 医疗付款方式 "
    
    strSQL = strSQL & " Order by 类别,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '1.性别
    rsTmp.Filter = "类别='性别'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboSex.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then cboSex.ListIndex = cboSex.NewIndex
            rsTmp.MoveNext
        Next
    End If
    '2.医疗付款方式
    rsTmp.Filter = "类别='医疗付款方式'"
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
    
    
    strSQL = " Select '处方职务' As 类别,count(药名ID) As num From 药品特性 Where 处方职务<>'00' Union All " & _
             " Select '处方限量' As 类别,count(药名ID) As num From 药品特性 Where 处方限量>0     Union All " & _
             " Select '储备限额' As 类别,Count(库房ID) As num From 药品储备限额 Where 上限>0 Or 下限>0"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    rsTmp.Filter = "类别='处方职务'"
    If Not rsTmp.EOF Then mbln处方职务检查 = (rsTmp!Num > 0)
    
    rsTmp.Filter = "类别='处方限量'"
    If Not rsTmp.EOF Then mbln处方限量检查 = (rsTmp!Num > 0)
    
    rsTmp.Filter = "类别='储备限额'"
    If Not rsTmp.EOF Then mbln储备限额检查 = (rsTmp!Num > 0)
    
    '------------------批量读取------------------
    
    
    
    '读取中药输入快捷
    Call ReadABCNum(mstrPrivs)
    
    '不同药房药品出库检查方式(包含所有药房,因为可以录入住院病人)
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '记帐分类报警
    If mbytInFun = 2 And mbytInState = 0 Then Set mrsWarn = GetUnitWarn("", "0")
    
        
    '结算方式:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹
    If mbytInFun = 0 Then
        Set mrs结算方式 = Get结算方式("收费")
        If Not mrs结算方式.EOF Then
            For i = 1 To mrs结算方式.RecordCount
                If Val(Nvl(mrs结算方式!应付款)) = 1 Then
                    '问题:33722,不加入应付款性质
                    mstr应付款结算方式 = Nvl(mrs结算方式!名称)
                Else
                    '是否有个人帐户
                    If mrs结算方式!性质 = 3 And mstr个人帐户 = "" Then
                        mstr个人帐户 = mrs结算方式!名称
                    End If
                    '只加入非医保和代收款的结算方式供选择
                    If InStr(",1,2,7,", "," & mrs结算方式!性质 & ",") > 0 Then
                        cbo结算方式.AddItem mrs结算方式!编码 & "-" & mrs结算方式!名称
                        cbo结算方式.ItemData(cbo结算方式.NewIndex) = mrs结算方式!性质
                        
                        If mrs结算方式!名称 = gstr结算方式 Then
                            cbo结算方式.ListIndex = cbo结算方式.NewIndex
                        End If
                        
                        If mrs结算方式!缺省 = 1 And cbo结算方式.ListIndex = -1 Then
                            cbo结算方式.ListIndex = cbo结算方式.NewIndex
                        End If
                    End If
                End If
                mrs结算方式.MoveNext
            Next
        End If
        If cbo结算方式.ListCount = 0 Then   '缺省值会在NewBill中再次设定
            MsgBox "收费场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If mbytInState = 0 Or mbytInState = 3 Then
            Set mrsOneCard = GetOneCard
            mblnOneCard = mrsOneCard.RecordCount > 0
        End If
    End If
    
    
    '费别,默认显示适用于所有科室的
    Call Load费别(cbo费别, 0, mbytInFun = 2, mrs费别)
    mrs费别.Filter = ""
    If mrs费别.RecordCount = 0 Then
        MsgBox "没有有效费别设置，请先到费别管理中进行设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '不缺省开单人和开单科室
    Call FillDept(mlngDeptID)
    If cbo开单科室.ListCount = 0 Then
        MsgBox "没有可用的开单科室,可用的开单科室须满足以下规则:" & vbCrLf & _
               "    1.部门性质为产科" & IIf(mbytInFun = 1, "、手术、治疗、检查、检验", "") & vbCrLf & _
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
    
    
    '可用收费类别:按序号排序
    If gstr收费类别 = "" Then
        strSQL = "Select 编码,名称 as 类别 from 收费项目类别 Where 编码<>'1' Order by 序号"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.编码,A.名称 as 类别 " & _
        "   From 收费项目类别 A, Table( f_Str2list([1])) J " & _
        "   Where A.编码=J. Column_Value " & _
        "   Order by 序号"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr收费类别, "'", ""))
    If mrsClass.EOF Then
        MsgBox "没有设置可用的收费类别,请先在本地参数中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    '当只有一种可选收费类别时,不用用户选择
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(gstr收费类别, "'5'") > 0 Or InStr(gstr收费类别, "'6'") > 0 _
        Or InStr(gstr收费类别, "'7'") > 0 Or gstr收费类别 = "" Then
        mlng药品类别ID = ExistIOClass(IIf(mbytInFun = 2, 9, 8))
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定处方单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr收费类别, "'4'") > 0 Or gstr收费类别 = "" Then
        mlng卫材类别ID = ExistIOClass(IIf(mbytInFun = 2, 41, 40))
        If mlng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    '费用类型
    strSQL = " Select '医保' As 类别,编码,名称 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Union All " & _
                 " Select '公费' As 类别,编码,名称 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") "
    Set mrs费用类型 = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrs费用类型, strSQL, Me.Caption)
       
        
    '开单日期
    If mbln补费 And mstr最后转科时间 <> "" Then
        Curdate = CDate(mstr最后转科时间) - 1 / 24 / 60
    Else
        Curdate = zlDatabase.Currentdate
    End If
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    '自动识别加班
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk加班.Value = 1
    End If
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str类别 As String, ByVal intPage As Integer, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'功能：获取最近输入的相同类别项目的执行科室ID
    Dim i As Long, j As Long, k As Long
    
    For i = intPage To 1 Step -1
        If i = intPage Then
            k = lngRow - 1
        Else
            k = mobjBill.Pages(i).Details.Count
        End If
        For j = k To 1 Step -1
            If mobjBill.Pages(i).Details(j).收费类别 = str类别 _
                And mobjBill.Pages(i).Details(j).执行部门ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Pages(i).Details(j).执行部门ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Pages(i).Details(j).执行部门ID
                    Exit Function
                End If
            End If
        Next
    Next
    
    '如果是卫生材料,再取与最近其它类别相匹配的执行科室
    If str类别 = "4" Then
        For i = intPage To 1 Step -1
            If i = intPage Then
                k = lngRow - 1
            Else
                k = mobjBill.Pages(i).Details.Count
            End If
            For j = k To 1 Step -1
                If mobjBill.Pages(i).Details(j).执行部门ID <> 0 Then
                    If InStr("," & strDeptIDs & ",", "," & mobjBill.Pages(i).Details(j).执行部门ID & ",") > 0 Then
                        GetLastDeptID = mobjBill.Pages(i).Details(j).执行部门ID
                        Exit Function
                    End If
                End If
            Next
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, Optional blnEnter As Boolean)
'功能：根据单据列设置下拉列表框内容
'参数：blnEnter=是否按光标进入该列处理,这时显示的内容保持不变
    Dim rsTmp As ADODB.Recordset
    Dim bln护士 As Boolean, strTmp As String
    Dim strSQL As String, strIDs As String, i As Long
    Dim lng病区ID As Long, lng科室ID As Long, j As Long
    Dim bln草药类别 As Boolean '是否允许输入草药类别
    Dim rsUnit As ADODB.Recordset
    Bill.Clear
    Err = 0: On Error GoTo Errhand:
    Select Case Bill.TextMatrix(0, lngCol)
        Case "类别"
            Call GetOperatorInfo(mobjBill.Pages(mintPage).开单人, bln护士)
            
                    
            mrsClass.Filter = 0: j = 1
            For i = 1 To mrsClass.RecordCount
                '护士类别:限制
                If Not (bln护士 And InStr(",E,M,4,", mrsClass!编码) = 0) Then
                    Bill.AddItem j & "-" & mrsClass!类别
                    Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!编码)  '存放类别编码的ASCII码
                    j = j + 1
                End If
                mrsClass.MoveNext
            Next
            Bill.cboStyle = DropOlnyDown
            
        Case "执行科室", "发药药店"
            Bill.cboStyle = DropDownAndEdit
            'Bill.ToolTipText = "执行科室当前项目的执行科室性质,科室本身的性质,病人来源等相关,如果是药品,与存储库房,材质对应的部门工作性质等相关"
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Pages(mintPage).Details.Count >= lngRow Then
                With mobjBill.Pages(mintPage).Details(lngRow)
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
                                lng科室ID = GetLastDeptID(.收费类别, mintPage, lngRow, Mid(strIDs, 2))
                            End If
                            If lng科室ID = 0 Then lng科室ID = .执行部门ID
                            
                            '确定当前行的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!编码 & "-", "") & mrsWork!名称
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                            
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        lng科室ID = mobjBill.科室ID     '病人科室
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
                                strTmp = IIf(zlIsShowDeptCode, rsUnit!编码 & "-", "") & rsUnit!名称
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(rsUnit!ID))) = False Then
                                '刘兴洪:28947
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.NewIndex) = rsUnit!ID
                                    
                                   '设置缺省执行科室
                                    If Not blnEnter Then '进入该列时保持已确定值不变
                                        If lngRow = 1 Then
                                            If rsUnit!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '与上一行非药品相同
                                            If rsUnit!ID = mobjBill.Pages(mintPage).Details(lngRow - 1).执行部门ID And mobjBill.Pages(mintPage).Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 _
                                                And InStr(",5,6,7,", mobjBill.Pages(mintPage).Details(lngRow - 1).收费类别) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf rsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                rsUnit.MoveNext
                            Next
                        End If
                            
                        If Not blnEnter And .Detail.执行科室 = 4 Then    '执行科室为指定科室的,缺省为操作员所在科室
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
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数
    '编制:刘兴洪
    '日期:2010-01-27 10:17:11
    '问题:27663
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mTy_Para
        .bln住院病人门诊收费 = IIf(Val(zlDatabase.GetPara("住院病人按门诊收费", glngSys, mlngModul, "0")) = 1, True, False)
        If mbytInFun = 2 Then .bln住院病人门诊收费 = True '住院病人记帐时默认按门诊收费
    End With
    mint退费回单打印 = Val(zlDatabase.GetPara("退费回单打印方式", glngSys, mlngModul, "0"))
End Sub


Private Sub InitFace()
'功能：根据表单要完成的功能设置界面布局
    Dim arrHead() As String, i As Integer, arrBaby As Variant, strTmp As String
    Dim blnStatusIn As Boolean
    
    '刘兴洪 问题:27331 日期:2010-01-12 09:48:43
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then
       '只有划价才会有此判断
       MCPAR.blnOnlyBjYb = zlIsOnly北京医保
    Else
        MCPAR.blnOnlyBjYb = False
    End If
    
    '刘兴洪 问题:27663 日期:2010-01-27 10:18:39
    Call InitPara
    
    
    '公用单据表格式
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .COLS = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.项目
        .PrimaryCol = BillCol.项目
        .MsfObj.ColAlignmentFixed(BillCol.行) = 4
        .TextMatrix(1, BillCol.行) = 1
        
        For i = 0 To UBound(arrHead)
            If glngSys Like "8??" And Split(arrHead(i), ",")(0) = "执行科室" Then
                .TextMatrix(0, i) = "发药药店"
            Else
                .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            End If
            If glngSys Like "8??" And .TextMatrix(0, i) = "标志" Then
                .ColWidth(i) = 0 '不要手术标志
            ElseIf glngSys Like "8??" And .TextMatrix(0, i) = "规格" Then
                .ColWidth(i) = Split(arrHead(i), ",")(1) + 270
            ElseIf glngSys Like "8??" And .TextMatrix(0, i) = "类型" Then
                .ColWidth(i) = Split(arrHead(i), ",")(1) + 250
            Else
                .ColWidth(i) = Split(arrHead(i), ",")(1)
            End If
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
        
           
        
        If mbytInState = 0 And mbytBilling <> 2 Then
            .ColData(BillCol.行) = BillColType.UnFocus
            .ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
            If mblnOne Then .ColData(BillCol.类别) = BillColType.UnFocus
            
            .ColData(BillCol.项目) = BillColType.CommandButton    '项目输入,按扭可选
            .ColData(BillCol.数次) = BillColType.Text             '数/次输入
            .ColData(BillCol.规格) = BillColType.UnFocus          '规格跳过
            .ColData(BillCol.商品名) = BillColType.UnFocus          '商品名跳过
            .ColData(BillCol.单位) = BillColType.UnFocus          '单位跳过
            .ColData(BillCol.付数) = BillColType.UnFocus          '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = BillColType.UnFocus          '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.应收金额) = BillColType.UnFocus          '应收金额跳过
            .ColData(BillCol.实收金额) = BillColType.UnFocus          '实收金额跳过
            .ColData(BillCol.执行科室) = BillColType.ComboBox        '默认取开单科室或上一科室
            .ColData(BillCol.标志) = BillColType.UnFocus         '标志缺省跳过,当为手术时,设为复选(-1)
            .ColData(BillCol.类型) = BillColType.UnFocus         '类型缺省跳过
        End If
        If mbytInState = 0 Or mbytInState = 2 Then '编辑界面
            .SetColColor BillCol.类别, &HE7CFBA
            .SetColColor BillCol.项目, &HE7CFBA
            .SetColColor BillCol.数次, &HE7CFBA
            .SetColColor BillCol.执行科室, &HE7CFBA
            .SetColColor BillCol.付数, &HE0E0E0
            .SetColColor BillCol.单价, &HE0E0E0
            .SetColColor BillCol.标志, &HE0E0E0
        End If
        
        ReDim marrColData(.COLS - 1)
        For i = 0 To .COLS - 1
            marrColData(i) = .ColData(i)
        Next

        If mbytInState = 3 Then .AllowAddRow = False
    End With
    '恢复注册表保存宽度
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInFun & mbytInState)
    If gTy_System_Para.byt药品名称显示 <> 2 Then
        '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        Bill.ColWidth(BillCol.商品名) = 0
    Else
        If Bill.ColWidth(BillCol.商品名) = 0 Then
             Bill.ColWidth(BillCol.商品名) = GetOrigColWidth(BillCol.商品名)
        End If
    End If
    
    Me.KeyPreview = True
    Set mobjBrushCheck = New clsBrushCardInput
    mobjBrushCheck.OnlyLegalCardNo = False
    Call mobjBrushCheck.InitCompents(Me, Bill, mobjCard)
        
    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = mbytInState = 0
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
    
    IDKind.Enabled = mbytInState = 0
    If mbytInState = 0 Then
        Call GetRegisterItem(g私有模块, Me.Name, "idkind", strTmp)
        IDKind.IDKind = Val(strTmp)
    End If
    
    '多单据收费:目录仅支持收费界面
    fraBill.Visible = mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gblnMulti
    lblDuty.Caption = ""
    fraSubBill.Visible = mbytInFun = 0 And mbytInState = 0    '该栏上还要显示开单人的专业技术职务
    
    '刘兴洪 问题:26949 日期:2009-12-28 13:52:50
    fra退费摘要.Visible = (mbytInFun = 0 And mbytInState = 3 Or mblnDelete)
    '25187
    vsInvoice.Visible = (mbytInFun = 0 And mbytInState = 3 Or mblnDelete) And gTy_Module_Para.byt票据分配规则 <> 0
    
    If Not (mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And InStr(mstrPrivs, "保险收费") > 0 And gint病人来源 = 1) Then
        cmdYB.Visible = False
        lblRePrint.Left = lblRePrint.Left - cmdYB.Width
        txtRePrint.Left = txtRePrint.Left - cmdYB.Width
        lblModi.Left = lblModi.Left - cmdYB.Width
        txtModi.Left = txtModi.Left - cmdYB.Width
        lblIn.Left = lblIn.Left - cmdYB.Width
        txtIn.Left = txtIn.Left - cmdYB.Width
    End If
    cmdSelWholeSet.Visible = mbytInState = 0
    cmdSaveWholeSet.Visible = InStr(mstrPrivs, ";增加成套项目;") > 0
    
    '中药配方:新单或修改时有效
    If Not (mbytInState = 0) Or mbytBilling = 2 Then
        cmd配方.Visible = False
        lblRePrint.Left = lblRePrint.Left - cmd配方.Width
        txtRePrint.Left = txtRePrint.Left - cmd配方.Width
        lblModi.Left = lblModi.Left - cmd配方.Width
        txtModi.Left = txtModi.Left - cmd配方.Width
        lblIn.Left = lblIn.Left - cmd配方.Width
        txtIn.Left = txtIn.Left - cmd配方.Width
    End If
                    
    '重打(仅收费有效)
    If Not (mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And InStr(mstrPrivs, "重打票据") > 0) Then
        lblRePrint.Visible = False
        txtRePrint.Visible = False
        
        lblModi.Left = lblModi.Left - lblRePrint.Width - txtRePrint.Width
        txtModi.Left = txtModi.Left - lblRePrint.Width - txtRePrint.Width
        lblIn.Left = lblIn.Left - lblRePrint.Width - txtRePrint.Width
        txtIn.Left = txtIn.Left - lblRePrint.Width - txtRePrint.Width
    End If
    
    '修改(仅收费,划价有效)
    If Not (((mbytInFun = 0 And InStr(";" & mstrPrivs & ";", ";记录修改;") > 0) Or (mbytInFun = 1 And InStr(";" & mstrPrivs & ";", ";修改;") > 0)) And mbytInState = 0 And mstrInNO = "") Then
        lblModi.Visible = False
        txtModi.Visible = False
        
        lblIn.Left = lblIn.Left - lblModi.Width - txtModi.Width
        txtIn.Left = txtIn.Left - lblModi.Width - txtModi.Width
    End If

    '导入(仅新增时有效)
    If Not (mbytInState = 0 And mstrInNO = "") Or mbytBilling = 2 Then
        lblIn.Visible = False
        txtIn.Visible = False
    End If
    Line3.Visible = mbytInFun = 0 And mbytInState = 0 And mstrInNO = ""
    Line4.Visible = mbytInFun = 0 And mbytInState = 0 And mstrInNO = ""
        
    
    '保险帐户和预交款
    If mbytInFun <> 0 Then
        vsBalance.Visible = False
        lbl预交冲款.Visible = False
        txt预交冲款.Visible = False
    End If
    
    '婴儿费
    cboBaby.Visible = (mbytInFun = 2)
    lblBaby.Visible = (mbytInFun = 2)
    If mbytInFun = 2 Then
        arrBaby = Array("0-病人本人", "1-第1个婴儿", "2-第2个婴儿", "3-第3个婴儿", "4-第4个婴儿", "5-第5个婴儿")
        For i = 0 To UBound(arrBaby)
            cboBaby.AddItem arrBaby(i)
        Next
        cboBaby.ListIndex = 0
    End If
    '结算方式
    '仅新增,修改,退费可用
    cbo结算方式.Enabled = (mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 3))
    
    '浏览和调整时不可见
    '只有退费才显示
    lbl结算方式.Visible = (mbytInFun = 0 And Not (mbytInState = 1 Or mbytInState = 2 Or mbytInState = 0))
    cbo结算方式.Visible = (mbytInFun = 0 And Not (mbytInState = 1 Or mbytInState = 2 Or mbytInState = 0 Or mbytInState = 4 Or mbytInState = 5))
    fra缴款.Visible = (mbytInFun = 0 And (mbytInState = 3))
    If mbytInFun = 0 And mbytInState = 1 Then
        Set lbl预交冲款.Container = picAppend
        Set txt预交冲款.Container = picAppend
         vsBalance.Width = vsBalance.Width + 100
    End If
    '票据号
    lblFact.Visible = (mbytInFun = 0)
    txtInvoice.Visible = (mbytInFun = 0)
    txtMCInvoice.Top = txtInvoice.Top   '在预结算后才会显示
    txtMCInvoice.Left = txtInvoice.Left
    
    '动态费别
    If glngSys Like "8??" Or mbytInFun = 2 Then
        lbl动态费别.Visible = False
        If mbytInFun = 2 Then cbo费别.Locked = True: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
    Else
        If mbytInState = 1 Or mbytInState = 2 Or mbytInState = 3 Then
            cbo费别.Locked = True: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
            lbl动态费别.Left = cbo费别.Left
            lbl动态费别.Visible = True
        Else
            lbl动态费别.BorderStyle = 0
        End If
    End If
    lbl险类.Caption = ""
    
    '收费时是否允许挂号
    Call ShowRegist
    
    '收费时是否允许就诊卡
    Call ShowIDCard
    
    '收费票据打印格式:收费,修改,退费时显示
    If mbytInFun = 0 And InStr(",0,3,", mbytInState) > 0 Then
        Call ZlShowBillFormat(mlngModul, lblFormat, mintInvoiceFormat)
    End If
    
    '退费销帐按钮
    If mbytInFun = 0 And mstrInNO = "" And gblnMulti Then
        cmdDelete.Visible = True '收费支持多单据时使用多单据退费
        chkCancel.Visible = False
    End If
    If Not (mbytInState = 0 And mstrInNO = "") Or mbytInFun = 1 Then
        chkCancel.Visible = False
    End If
    
    Select Case mbytInFun
        Case 0 '收费
            If glngSys Like "8??" Then
                Caption = "药店收费处理"
                lblTitle.Caption = gstrUnitName & "药店收费单"
            Else
                Caption = "病人收费处理"
                lblTitle.Caption = gstrUnitName & "病人收费单"
            End If
            
            Call SetMoneyList
            
            Call InitBalanceGrid
            
            If mbytInState <> 0 Then
                '非收费状态
                lbl累计.Visible = False
                txt累计.Visible = False
                lbl应收.Top = lbl应收.Top + txt累计.Height / 3
                txt应收.Top = txt应收.Top + txt累计.Height / 3
                lbl合计.Top = lbl合计.Top + txt累计.Height / 1.5
                txt合计.Top = txt合计.Top + txt累计.Height / 1.5
            Else
                If Not gbln累计 Then
                    lbl累计.Visible = False
                    txt累计.Visible = False
                    lbl应收.Top = lbl应收.Top + txt累计.Height / 3
                    txt应收.Top = txt应收.Top + txt累计.Height / 3
                    lbl合计.Top = lbl合计.Top + txt累计.Height / 1.5
                    txt合计.Top = txt合计.Top + txt累计.Height / 1.5
                End If
            End If
            
            '输入控制
            Call SetInputItem
            
            '权限设置
            If InStr(mstrPrivs, "门诊退费") = 0 Then
                chkCancel.Visible = False
                cmdDelete.Visible = False
            End If
            txtInvoice.Locked = Not (InStr(1, mstrPrivs, "修改票据号") > 0) And gblnStrictCtrl
        Case 1 '划价
            Caption = "门诊划价处理"
            
            lblTitle.Caption = gstrUnitName & "门诊划价单(" & UserInfo.姓名 & ")"
            mshMoney.Width = mshMoney.Width * 2
            Call SetMoneyList
            Call SetStatPosition
'
'            lbl缴款.Visible = False: txt缴款.Visible = False
'            lbl找补.Visible = False: txt找补.Visible = False
                        
            '输入控制
            Call SetInputItem
        Case 2 '门诊记帐
            Caption = "门诊记帐处理"
            
            Select Case mbytBilling
                Case 0
                    lblTitle.Caption = gstrUnitName & "门诊记帐单"
                Case 1
                    lblTitle.Caption = gstrUnitName & "门诊记帐单(划价)"
                Case 2
                    lblTitle.Caption = gstrUnitName & "门诊记帐单(审核)"
                    
                    cboNO.Locked = False
                    fraInfo.Enabled = False
                    fraAppend.Enabled = False
                    Bill.Active = False
                    
                    Call SetPatientEnableModi(False)
            End Select
            
            lblCorp.Visible = (mbytInState = 0)
            lblCorp.Left = txtIn.Left + txtIn.Width + 100
            lblCorp.Top = lblIn.Top
            
            chkCancel.Caption = "销"
            lblFlag.Caption = "销"
                        
            cbo医疗付款.Locked = True
            
            Call SetMoneyList
            Call SetStatPosition
            
            
'''            lbl缴款.Visible = False: txt缴款.Visible = False
'''            lbl找补.Visible = False: txt找补.Visible = False
                       
            
            '权限设置
            If InStr(mstrPrivs, "门诊销帐") = 0 Then
                chkCancel.Visible = False
            End If
            
            lblTotal.Visible = True
            lblTotal.Top = cmdOK.Top
            
            cboSex.Enabled = False
            txt年龄.Enabled = False
            cbo年龄单位.Enabled = False
    End Select
    
    If mbln补费 Then
        If mlng主页ID <> 0 Then
            Dim strSQL As String, rsTemp As ADODB.Recordset
            strSQL = "Select 当前病区ID,出院科室ID,出院日期 From 病案主页 Where 病人ID = [1] And 主页ID = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
            If Not rsTemp.EOF Then
                If mlngDeptID = 0 Then
                    mlngDeptID = Val(Nvl(rsTemp!出院科室ID))
                End If
                If mlngUnitID = 0 Then
                    mlngUnitID = Val(Nvl(rsTemp!当前病区ID))
                End If
                blnStatusIn = IsNull(rsTemp!出院日期)
            End If
        End If
        If blnStatusIn Or mlng主页ID = 0 Or rsTemp.EOF Then
            lblTitle.Caption = lblTitle.Caption & "(" & "补费" & ")"
        Else
            lblTitle.Caption = lblTitle.Caption & "(第" & mlng主页ID & "次补费" & ")"
        End If
    End If
        
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 4 Or mbytInState = 5 Then
        '执行或调整状态
        
        If mbytInFun = 2 Then
            If mbytInState <> 0 Or mbytBilling = 2 Then Call SetDisible       '记帐审核
        End If
        
        If mbytInState = 0 Then
            If mstrInNO <> "" Then txtPatient.BackColor = &HE0E0E0           '修改
            If mbytBilling = 0 Or mbytBilling = 1 Then Call SetShowCol       '记帐、划价
                        
        ElseIf mbytInState = 2 Then  '调整开单人和时间
            Call SetDisible
            
            txtInvoice.Enabled = False
            fraInfo.Enabled = False
                            
            cbo开单人.Locked = False
            txtDate.Enabled = True
            Call SetShowCol
        End If
        
        Call SetButton(2) '确定,取消
    Else
        '查阅 或退费,销帐
        Call SetDisible
        
        fraAppend.Enabled = False
        
        fraTitle.Enabled = False
        fraInfo.Enabled = False
        
        If mbytInState = 3 Then  '退费
            If mbytInFun = 0 Then
                '部份退费只支持指定结算方式
                fraAppend.Enabled = True
                cbo结算方式.Locked = False
            End If
            Call SetButton(2) '确定,取消
            Call ShowDeleteCol(True)
            Bill.Active = True
            fra退费摘要.Enabled = True
        Else
            Call SetButton(3) '取消
            fra退费摘要.Enabled = False
            
        End If
        
        If mblnDelete Then lblFlag.Visible = True
    End If
    
    If gbyt科室医生 = 0 Then
        Call ExChangeLocate(cbo开单科室, cbo开单人)
        lbl科室.Caption = "开单人(&W)"
        lbl科室.Left = lblPatient.Left
        lbl开单人.Caption = "开单科室"
        cbo开单科室.TabStop = False
    End If
    
    If Not (mbytInState = 0 And (mbytInFun = 0 Or mbytInFun = 1)) Then
        sta.Panels("Drugstore").Visible = False
    End If
    
    If mbytInState = 0 And mstrInNO = "" Then
        sta.Panels("PatiSource").Visible = True
        Set sta.Panels("PatiSource").Picture = imgPati.ListImages(IIf(gint病人来源 = 1, "OutPati", "InPati")).Picture
    Else
        sta.Panels("PatiSource").Visible = False
    End If
    Bill.ColWidth(BillCol.从属父号) = 0
    Bill.ColWidth(BillCol.医嘱序号) = 0
    Bill.ColWidth(BillCol.执行科室ID) = 0
    
    '82801,冉俊明,2015-2-26
    txt年龄.MaxLength = zlGetPatiInforMaxLen.intPatiAge
End Sub

Private Sub SetStatPosition()
'功能：门诊划价和门诊记帐时设置合计信息区内的控件位置
    Dim blnVisible As Boolean
    
    If mbytInState = 0 And mstrInNO = "" And (mbytInFun = 1 Or mbytInFun = 2) Then
        fraUpBillShow.Visible = True
        blnVisible = True
    Else
        blnVisible = False: fraUpBillShow.Visible = False
    End If
    
    fraStat.Width = lbl合计.Width + txt合计.Width + 600
    fraStat.Left = mshMoney.Left + mshMoney.Width
    lbl应收.Left = 200: txt应收.Left = lbl应收.Left + lbl应收.Width + 200
    lbl合计.Left = lbl应收.Left
    txt合计.Left = txt应收.Left
          
    
    If mbytInFun = 1 Then
        lbl累计.Left = lbl合计.Left: txt累计.Left = txt合计.Left
        '累计用来作分币处理后的金额
        lbl累计.Visible = True: txt累计.Visible = True: lbl累计.Caption = "应缴"
        
        If blnVisible Then
            fraUpBillShow.Left = fraStat.Left + fraStat.Width + 50
        End If
    ElseIf mbytInFun = 2 Then
        lbl累计.Visible = False: txt累计.Visible = False
         fraUpBillShow.Visible = False
        If blnVisible Then
            fraUpBillShow.Visible = True
            fraUpBillShow.Left = fraStat.Left + fraStat.Width + 50
'            txtPreNO.Width = txt合计.Width
'            Set lblPreNO.Container = fraStat
'            Set txtPreNO.Container = fraStat
'            lblPreNO.Left = lbl合计.Left: txtPreNO.Left = txt合计.Left
'            lblPreNO.Top = lbl累计.Top: txtPreNO.Top = txt累计.Top
        Else
            lbl应收.Top = lbl应收.Top + txtPreNO.Height / 2
            txt应收.Top = txt应收.Top + txtPreNO.Height / 2
            lbl合计.Top = lbl合计.Top + txtPreNO.Height * 0.75
            txt合计.Top = txt合计.Top + txtPreNO.Height * 0.75
        End If
    End If
End Sub

Private Sub SetButton(bytType As Byte)
'功能：设置功能按钮状态和位置
'参数：bytType=1:预结算,确定,取消
'              2:确定,取消
'              3:取消
'              4:预结算,确定,完成收费,取消
'说明：该函数为初始时调用,不可重复调用
    Dim lngTmp As Long
    
    Const H_间隔 = 45
    
    LockWindowUpdate picAppend
    
    '恢复缺省状态，且不可见
    cmd预结算.Visible = False
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdPrint.Visible = False
    
    cmd预结算.Top = lblSeek.Top
    If mbytInFun = 1 Or mbytInFun = 2 Then
        cmdOK.Top = lblSeek.Top
    Else
        cmdOK.Top = cmd预结算.Top + cmd预结算.Height + H_间隔
    End If
    cmdCancel.Top = cmdOK.Top + cmdOK.Height + H_间隔
    cmdPrint.Top = cmdCancel.Top + cmdCancel.Height + H_间隔
            
    cmdCancel.Caption = "取消(&C)"
    cmdOK.Enabled = True
    
    Select Case bytType
        Case 1 '预结算,确定,取消
            cmd预结算.Visible = True
            cmdOK.Visible = True
            cmdCancel.Visible = True
            
            cmd预结算.Top = cmd预结算.Top + cmdPrint.Height / 2 + H_间隔
            cmdOK.Top = cmdOK.Top + cmdPrint.Height / 2 + H_间隔
            cmdCancel.Top = cmdCancel.Top + cmdPrint.Height / 2 + H_间隔
            
            cmd预结算.TabStop = True
        Case 2 '确定,取消
            cmdOK.Visible = True
            cmdCancel.Visible = True
        Case 3 '取消
            cmdCancel.Visible = True
            cmdCancel.Caption = "退出(&X)"
            cmdCancel.Top = cmdCancel.Top - cmdPrint.Height / 2 - H_间隔
        Case 4 '预结算,确定,打印,取消
            cmd预结算.Visible = True
            cmdOK.Visible = True
            cmdCancel.Visible = True
            cmdPrint.Visible = True
            
            lngTmp = cmdPrint.Top
            cmdPrint.Top = cmdCancel.Top
            cmdCancel.Top = lngTmp
    End Select
    LockWindowUpdate 0
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'功能:界面设置为不可修改状态
'参数:bln为True表示设置为可以修改的状态

    cboNO.Locked = Not bln
    
    cbo费别.Locked = Not bln Or mbytInFun = 2: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
    cbo医疗付款.Locked = Not bln Or mbytInFun = 2
    
    cbo开单科室.Locked = Not bln Or mbytBilling = 2
    cbo开单人.Locked = Not bln Or mbytBilling = 2
    cbo开单科室.Enabled = bln And mbytBilling <> 2
    cbo开单人.Enabled = bln And mbytBilling <> 2
    
    chk加班.Enabled = bln And mbytBilling <> 2
    
    If mbytInFun = 2 And mbytInState = 0 Then
        If bln And mbytBilling <> 2 And cbo开单科室.ListIndex <> -1 Then
            cboBaby.Enabled = is产科(cbo开单科室.ItemData(cbo开单科室.ListIndex), mrs开单科室)
        Else
            cboBaby.Enabled = False
        End If
    End If
    
    cbo结算方式.Locked = Not bln
    txtDate.Enabled = bln
    fraStat.Enabled = bln
    fra缴款.Enabled = bln
    Bill.Active = bln And mbytBilling <> 2
    
'''    If Not bln Then
'''        txt缴款.BackColor = &HE0E0E0
'''    Else
'''        txt缴款.BackColor = &HFFFFFF
'''    End If
    
    SetPatientEnableModi (bln)
End Sub

Private Sub SetDeptDoctorByRegevent(ByVal lng病人ID As Long, Optional strRegNO As String)
'功能：根据病人ID或挂号单中病人的挂号科室和医生信息设置开单科室和开单人
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


Private Function GetDeptByRegevent(ByVal lng病人ID As Long) As ADODB.Recordset
'功能：根据病人ID返回有效挂号单的科室ID
    Dim strSQL As String, strWhere As String
    strWhere = zlGetRegEventsCons(, , True)
    On Error GoTo errH
    strSQL = "Select 执行部门ID From 病人挂号记录" & _
            " Where 病人ID=[1] and 记录性质=1 and 记录状态=1  " & strWhere
    Set GetDeptByRegevent = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadAddedItem(ByVal lng病人ID As Long, Optional ByVal str病人姓名 As String)
'功能:自动加收挂号费
    Dim i As Long, j As Long, objThis As Control
    
    '检查现有单据中是否已加收
    For i = 1 To mobjBill.Pages.Count
        For j = 1 To mobjBill.Pages(i).Details.Count
            If mobjBill.Pages(i).Details(j).收费细目ID = glngAddedItem Then
                Exit Sub
            End If
        Next
    Next
    
    If CheckAddedItem(lng病人ID, str病人姓名) Then
        Set objThis = Me.ActiveControl
        '如果当前单据是划价单，则新增一张单据
        If Not Bill.Active Then
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO = "" Then Exit For
            Next
            If i <= mobjBill.Pages.Count Then
                tbsBill.Tabs(i).Selected = True
            Else
                If cmdAddBill.Enabled And cmdAddBill.Visible Then
                    Call cmdAddBill_Click
                Else
                    Exit Sub '不允许多张单据收费时，不进行加收
                End If
            End If
        End If
        
        Call LocateNewRow
        If gbln收费类别 Then
            Bill.Col = BillCol.类别 '自动调用call Bill_EnterCell
            For i = 0 To Bill.ListCount - 1
                If Bill.ItemData(i) = Asc("Z") Then Bill.ListIndex = i: Exit For
            Next
            If i > Bill.ListCount - 1 Then Exit Sub '如果开单人是护士，可能不能输入该类别，则不进行加收
            
            Call Bill_KeyDown(vbKeyReturn, 0, False)
        End If
        
        Bill.Col = BillCol.项目
        Bill.TxtVisible = True
        Bill.Text = glngAddedItem
        mblnSelect = True
        Call Bill_KeyDown(vbKeyReturn, 0, False)
        
        On Error Resume Next
        If objThis.Visible And objThis.Enabled Then objThis.SetFocus
        On Error GoTo 0
    End If
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
Private Sub initInsurePara(ByVal lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2011-08-27 12:25:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    MCPAR.允许不设置医保项目 = gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, mintInsure)
    MCPAR.门诊收费存为划价单 = gclsInsure.GetCapability(support门诊收费存为划价单, lng病人ID, mintInsure)
    MCPAR.门诊必须传递明细 = gclsInsure.GetCapability(support门诊必须传递明细, lng病人ID, mintInsure)
    MCPAR.医生确定处方类型 = gclsInsure.GetCapability(support医生确定处方类型, lng病人ID, mintInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    MCPAR.多单据一次结算 = gclsInsure.GetCapability(support多单据一次结算, lng病人ID, mintInsure)
    MCPAR.门诊连续收费 = gclsInsure.GetCapability(support门诊连续收费, lng病人ID, mintInsure)
    MCPAR.多单据收费 = gclsInsure.GetCapability(support多单据收费, lng病人ID, mintInsure)
    MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, mintInsure)
    MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, mintInsure)
    MCPAR.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, lng病人ID, mintInsure)
    MCPAR.全自付 = gclsInsure.GetCapability(support收费帐户全自费, lng病人ID, mintInsure)
    MCPAR.实时监控 = gclsInsure.GetCapability(support实时监控, lng病人ID, mintInsure)
    MCPAR.退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, lng病人ID, mintInsure)
    MCPAR.多单据调一次交易 = gclsInsure.GetCapability(support门诊_不分单据结算, lng病人ID, mintInsure)
    MCPAR.医保不走票号 = False
    '刘兴洪:27536 20100119
    MCPAR.不提醒缴款金额不足 = gclsInsure.GetCapability(support不提醒缴款金额不足, lng病人ID, mintInsure)
End Sub


Private Sub MCPatientProcess(Optional ByVal lngCur病人ID As Long, Optional blnErrBill As Boolean)
    Dim i As Long, blnTran As Boolean
    Dim lng病人ID As Long, lng病人IDOut As Long
    Dim lng挂号科室 As Long, str开单科室 As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, strTemp As String, intInsure As Integer
    Dim blnPriceBill As Boolean
    
    On Error GoTo errH
    If gblnLED Then zl9LedVoice.Speak "#50"
    lng病人IDOut = lngCur病人ID '避免Identify接口中修改该变量后返回新值
    
    '返回：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,24就诊类型(1=急诊门诊),25开单科室名称
    
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard (False)
    mstrYBPati = gclsInsure.Identify(id门诊收费, lng病人IDOut, mintInsure)
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
    
    blnPriceBill = False
    If mstrYBPati <> "" Then
        '获取病人信息
        If UBound(Split(mstrYBPati, ";")) >= 8 Then
            If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
                lng病人ID = Val(CLng(Split(mstrYBPati, ";")(8)))
                If lng病人ID <> lngCur病人ID And lngCur病人ID <> 0 And lng病人ID <> 0 Then
                    MsgBox "医保验证的病人与之前提取的病人不是同一个病人!", vbInformation, gstrSysName
                    Call YBIdentifyCancel
                    mintInsure = 0: mstrYBPati = ""
                    Exit Sub
                End If
            End If
        End If
        '问题:29283
        '  -- 参数:调用场合-1-挂号;2-收费
        '  --        病人id_In-病人ID(未建档的,传入零)
        '  --        卡号_In: 刷卡卡号;未刷卡时,为空
        '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
        If zlPatiCardCheck(2, lng病人ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
            Call YBIdentifyCancel
            mintInsure = 0: mstrYBPati = ""
            Exit Sub
        End If
        Call initInsurePara(lng病人ID)   '初始化医保参数
        If (MCPAR.门诊连续收费 Or Not MCPAR.多单据收费) And tbsBill.Tabs.Count > 1 Then
            If MCPAR.门诊连续收费 Then
                MsgBox "在医保连续收费模式下不支持多张单据收费。", vbInformation, gstrSysName
            ElseIf Not MCPAR.多单据收费 Then
                MsgBox "当前险类不支持多张单据收费。", vbInformation, gstrSysName
            End If
            Call YBIdentifyCancel
            If Visible Then txtPatient.SetFocus
            mintInsure = 0: mstrYBPati = ""
            Exit Sub
        End If
        If MCPAR.多单据一次结算 Then
            If gTy_Module_Para.bln分别打印 Then
                MsgBox "多单据模式下，医保一次结算时，不允许分别打印。", vbInformation, gstrSysName
                Call YBIdentifyCancel
                If Visible Then txtPatient.SetFocus
                mintInsure = 0: mstrYBPati = ""
                Exit Sub
            End If
        End If
            
        '问题:28240
        strTemp = mstrYBPati: intInsure = mintInsure
            
        If GetPatient("-" & lng病人ID, , , True) Then
            mstrYBPati = strTemp: mintInsure = intInsure
            If Not CheckRegisted(lng病人ID) Then
                Call YBIdentifyCancel
                Set mrsInfo = New ADODB.Recordset
                mintInsure = 0: mstrYBPati = ""
                Exit Sub
            End If
            With mobjBill
                .病人ID = Nvl(mrsInfo!病人ID, 0)
                .主页ID = Nvl(mrsInfo!主页ID, 0)
                .标识号 = Nvl(mrsInfo!门诊号, 0)
                .病区ID = Nvl(mrsInfo!当前病区ID, 0)
                .科室ID = Nvl(mrsInfo!当前科室id, 0)
                .床号 = "" & mrsInfo!当前床号
                .姓名 = "" & mrsInfo!姓名
                .性别 = "" & mrsInfo!性别
                .年龄 = "" & mrsInfo!年龄
                '费别在后面调用LoadAndSeek费别时赋值
            End With
            txt门诊号.Text = Nvl(mrsInfo!门诊号)
            Call InitBalanceGrid(True)
        Else
            Call YBIdentifyCancel
            mintInsure = 0: mstrYBPati = ""
            Exit Sub
        End If
        
        
        If fraBill.Visible Then
            cmdAddBill.Enabled = Not MCPAR.门诊连续收费 And MCPAR.多单据收费 And InStr(1, mstrPrivs, "医保病人多单据收费") > 0
        End If
        '75259:李南春,2014-7-10，病人姓名显示颜色处理
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), vbRed)
            Else
                txtPatient.ForeColor = vbRed
            End If
        Else
            txtPatient.ForeColor = vbRed
        End If
        txtPatient.Text = Split(mstrYBPati, ";")(3)
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        cboSex.ListIndex = cbo.FindIndex(cboSex, CStr(Split(mstrYBPati, ";")(4)), True)
        If IsDate(Split(mstrYBPati, ";")(5)) Then
            txt年龄.Text = ReCalcOld(CDate(Split(mstrYBPati, ";")(5)), cbo年龄单位, lng病人ID)
        Else
            Call LoadOldData("" & mrsInfo!年龄, txt年龄, cbo年龄单位)
            If Not IsNull(mrsInfo!出生日期) Then txt年龄.Text = ReCalcOld(mrsInfo!出生日期, cbo年龄单位, lng病人ID)
            
        End If
        lbl险类.Caption = "" & mrsInfo!险类名称
        
        mobjBill.病人ID = lng病人ID
        mobjBill.姓名 = Split(mstrYBPati, ";")(3)
        mobjBill.性别 = Split(mstrYBPati, ";")(4)
        mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
        
        
        '开单科室名称
        If UBound(Split(mstrYBPati, ";")) >= 25 And mobjBill.Pages(mintPage).NO = "" Then   '划价单的开单人开单科室优先
            str开单科室 = CStr(Split(mstrYBPati, ";")(25))
            If str开单科室 <> "" Then
                Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, str开单科室, True)) '不触发click事件
                Call cbo开单科室_Click
            End If
        End If
        '根据病人挂号信息设置开单科室和医生
        If mobjBill.Pages(mintPage).NO = "" Then    '划价单的开单人开单科室优先
            Call SetDeptDoctorByRegevent(lng病人ID)
        End If
        
        '显示急诊标记
        If UBound(Split(mstrYBPati, ";")) >= 24 Then
            chk急诊.Visible = Val(Split(mstrYBPati, ";")(24)) = 1
        End If
        
        '个人帐户
        mcur个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
        sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
        sta.Panels(Pan.C3个人帐户).Visible = True
        
        '支持预结算时就不固定显示个人帐户,否则显示
        If MCPAR.门诊预结算 Then
            '显示预结算按钮
            cmd预结算.Enabled = True
            Call SetButton(1) '预结算,确定,取消
            cmdOK.Enabled = False
        ElseIf mstr个人帐户 <> "" Then '只有使用个人帐户才用
            Call SetButton(2) '确定,取消
            vsBalance.TextMatrix(0, 0) = mstr个人帐户
            vsBalance.TextMatrix(0, 1) = "0.00"
            vsBalance.RowData(0) = 0
        End If
        
        sta.Panels(Pan.C2提示信息) = ""
        SetPatientEnableModi (False)
        
        txtRePrint.Enabled = False
        txtModi.Enabled = False '不能清除
        txtIn.Enabled = False
        cboNO.Enabled = False
        chkCancel.Enabled = False
        cmdDelete.Enabled = False
        
        '一个交易未完成,不允许另一个交易(挂号)
        If cmdIDCard.Visible Then cmdIDCard.Enabled = False
        If cmdRegist.Visible Then cmdRegist.Enabled = False
        
        If MCPAR.门诊连续收费 Then Call SetButton(4)  '预结算,确定,完收收费,取消
        
        '医疗付款方式
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!医疗付款方式) Then
                cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, mrsInfo!医疗付款方式, True)
            End If
        End If
        If cbo医疗付款.ListIndex = -1 Then cbo医疗付款.ListIndex = GetCboIndexByCode(cbo医疗付款, "1")
        
        cbo医疗付款.Locked = True
        
        '读取病人的多张划价单,之前提取过并且支持多单据收费时不再提取
        If mbytInFun = 0 And mbytInState = 0 And Visible And mstrInNO = "" And txtIn.Text = "" And mrsInfo.State = 1 And _
            Not (lngCur病人ID > 0 And Not MCPAR.门诊连续收费 And MCPAR.多单据收费 And InStr(1, mstrPrivs, "医保病人多单据收费") > 0) Then
            If gblnCheckRegeventDept And gint病人来源 = 1 And IsRegisterDept Then lng挂号科室 = Val("" & mrsInfo!执行部门ID)
            blnPriceBill = LoadMultiBills(lng病人ID, MCPAR.门诊连续收费 Or Not MCPAR.多单据收费 Or InStr(1, mstrPrivs, "医保病人多单据收费") = 0, lng挂号科室)
        End If
        
        '自动加收挂号费
        Call LoadAddedItem(lng病人ID)
                    
        '现有单据输入内容的处理
        '--------------------------------------------------------------------
        
        '医保病人当新单子处理,不管缴款结束以及是否是相同病人
        '刘兴洪:22343
        If (gTy_Module_Para.byt缴款控制 <> 1 And gTy_Module_Para.byt缴款控制 <> 3) Or mstrPrePati = "" Then    '仅以缴款作为结束时,即使不同的病人也保留收费,'除非缴款结束
            Call ClearPayInfo
            mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
            Call InitCommVariable
            Call ClearTotalInfo(True)
            Call ClearMoney
        End If
        
        '计算已提取的划价单的相关保险数据
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To tbsBill.Tabs.Count
            If mobjBill.Pages(i).NO <> "" Then
                strSQL = "zl_门诊划价记录_Update(" & mintInsure & "," & lng病人ID & ",'" & mobjBill.Pages(i).NO & "',0)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        Next
        gcnOracle.CommitTrans: blnTran = False
        
        '全部重新计算并显示
        Call ShowMoney
        
        '其它特殊处理
        '---------------------------------------------------------------------------------------
        '联合医保
        mblnSaveAsPrice = MCPAR.门诊收费存为划价单
        If mblnSaveAsPrice Then
            Call SetButton(2) '确定,取消
            sta.Panels(Pan.C3个人帐户).Text = ""
            sta.Panels(Pan.C3个人帐户).Visible = False
            Call ShowPayInfo(False)
        End If

        '处理预交结算
        '联合医保不使用预交冲款(划价模式)
        '咸阳医保不使用预交冲款
        If Not mblnSaveAsPrice And mintInsure <> 61 Then Call LoadFeeInfor(lng病人ID)
        
        '咸阳医保不缴款
        If mintInsure = 61 Then Call ShowPayInfo(False)
                
        If mstrInNO = "" Then
            Call LoadAndSeek费别
            '49573
            If cmdOK.Enabled And cmdOK.Visible Then
                cmdOK.SetFocus
            ElseIf cbo开单科室.Enabled And cbo开单科室.Visible And gbyt科室医生 <> 0 Then
                cbo开单科室.SetFocus
            ElseIf cbo开单人.Enabled And cbo开单人.Visible Then
                cbo开单人.SetFocus
            ElseIf cboSex.Enabled And cboSex.Visible Then
                cboSex.SetFocus
            ElseIf Bill.Enabled Then
                Bill.SetFocus
            End If
            
            '问题:39253
            If gbln划价立即缴款 = False And blnPriceBill Or mstrYBPati <> "" Then
                If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
            End If
           
            
            If gbln划价立即缴款 And blnPriceBill And mstrYBPati <> "" Then
                If cmd预结算.Visible And cmd预结算.Enabled Then
                    cmd预结算.SetFocus
                End If
            End If
            
            If gbyt科室医生 <> 0 Then
                If blnPriceBill Then
                    If cbo开单科室.Enabled And cbo开单科室.Visible And cbo开单科室.ListIndex < 0 Then cbo开单科室.SetFocus
                Else
                    If cbo开单科室.Enabled And cbo开单科室.Visible Then cbo开单科室.SetFocus
                End If
            Else
                If blnPriceBill Then
                    If cbo开单人.Enabled And cbo开单人.Visible And cbo开单人.ListIndex < 0 Then cbo开单人.SetFocus
                Else
                    If cbo开单人.Enabled And cbo开单人.Visible Then cbo开单人.SetFocus
                End If
            End If
            
            Call ShowWelcomeByLed
            Call ReInitPatiInvoice
        End If
    Else
        mintInsure = 0: mcur个帐余额 = 0: mcur个帐透支 = 0
        Call InitBalanceGrid
        sta.Panels(Pan.C3个人帐户).Text = ""
        sta.Panels(Pan.C3个人帐户).Visible = False
        
        sta.Panels(Pan.C2提示信息) = "身份验证不成功！"
        If Visible Then
            txtPatient.SetFocus
            Call txtPatient_GotFocus
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Set病人补费编辑属性()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置病人补费时的编辑属性
    '编制:刘兴洪
    '日期:2010-12-10 14:54:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln补费 = False Then Exit Sub
    txtPatient.Enabled = False
    cbo开单科室.Enabled = False
    cboSex.Enabled = False
    IDKind.Enabled = False
    chkCancel.Visible = False
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim i As Long, lng病人ID As Long, lng挂号科室 As Long
    Dim strPati As String, blnIDCard As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean, blnCancel As Boolean
    
    Dim int上次病人来源 As Integer
    Dim blnHavePriceBill As Boolean '当前是否提取的划价单(划价单时,直接缴款)
    Dim blnCheckReg As Boolean
    
    On Error GoTo errH
    blnHavePriceBill = False
    If KeyAscii = 13 And mblnValid = False Then
        mblnKeyReturn = True
    Else
        mblnKeyReturn = False
    End If
    
    '1.医保身份验证部份:仅门诊病人收费时使用
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And mbytInFun = 0 And mbytInState = 0 And gint病人来源 = 1 And Not mblnValid Then
        If txtPatient.Text = "" And chkCancel.Value = 0 And InStr(mstrPrivs, "保险收费") > 0 Then
            Call MCPatientProcess
            Exit Sub
        End If
    End If
    If txtPatient.Locked Then Exit Sub '锁定状态只允许医保验卡
   
   '问题:51488
    If (IDKind.Cards.读卡快键 = "空格键" Or IDKind.Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
   
    blnCheckReg = False
    '问题:27364 日期:2010-01-13 15:27:50
    If mblnAutoChangePati And gint病人来源 = 2 And (KeyAscii <> 13) Then
        '需要切找到病人来源1中
        gint病人来源 = 1: zlChangePatiSource (gint病人来源)
    End If
    
    
    '2.划价和记帐不输入病人直接回车(允许不输入姓名、性别、年龄):住院病人划价暂不提供选择器
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And Trim(txtPatient.Text) = "" And (mbytInFun = 1 Or mbytInFun = 2) Then
        '门诊划价保存后保留了单据信息,清除保留信息(成都妇幼新增)
        If CheckBillsEmpty Then Bill.Active = True:  Call ClearBillRows
               
        Call ClearmobjBill '清除对象中的病人信息
        Call ClearPatientInfo
        Call InitCommVariable
        Call ClearMoney
        If CheckBillsEmpty Then
            mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
            Call ClearTotalInfo(True)
        Else
            Call ShowMoney
        End If
                    
        If mbytInFun = 1 Then
            If Not mblnValid And Visible Then
                If gint病人来源 = 2 Then Exit Sub
                If gbln医疗付款 Then
                    If cbo医疗付款.Enabled And cbo医疗付款.Visible Then cbo医疗付款.SetFocus
                Else
                    If gbyt科室医生 = 1 Then
                        If cbo开单科室.Enabled And cbo开单科室.Visible Then cbo开单科室.SetFocus
                    Else
                        If cbo开单人.Enabled And cbo开单人.Visible Then cbo开单人.SetFocus
                    End If
                End If
            End If
            
            KeyAscii = 0
            Exit Sub
        End If
    End If
        
       
    '3.正常输入病人(姓名各种标识)部份:住院病人收费时可弹出选择器
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And mbytInState = 0 And Trim(txtPatient.Text) = "" And Not mblnValid Then
        If mbytInFun = 2 Then
            lng病人ID = SelectPatient
            If lng病人ID = 0 Then Exit Sub
            txtPatient.Text = "-" & lng病人ID
        ElseIf mbytInFun = 0 And gint病人来源 = 2 Then
            frmPatiSelect.Show 1, Me
            If frmPatiSelect.mlngPatient = 0 Then Exit Sub
            txtPatient.Text = "-" & frmPatiSelect.mlngPatient
        End If
    End If
    
     
    If IDKind.GetCurCard.名称 Like "姓名*" And Not mblnValid Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText And Not mblnValid, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If gint病人来源 = 1 And mbytInFun = 0 And InStr(mstrPrivs, "允许非医保病人") = 0 Then
            txtPatient.Text = "":  Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '病人未改变退出(指未保存前,不是指连续收费，因为连续收费时mrsInfo是在newbill中初始了的)
        If mrsInfo.State = 1 Then
            
            If txtPatient.Text = mrsInfo!姓名 Then
                If mblnValid Then Exit Sub
                mblnNotValied = True
                Call zlCommFun.PressKey(vbKeyTab): mblnNotValied = False: Exit Sub
            
            End If
            If mrsInfo!姓名 = "新病人" Then
                mobjBill.姓名 = txtPatient.Text
                mblnNotValied = True
                Call zlCommFun.PressKey(vbKeyTab): mblnNotValied = False: Exit Sub
            End If
        End If
        
        '门诊划价清除保留信息(保存后保留了单据和病人信息),此处仅清除单据信息,累计信息在后面输入确认后再处理
        If mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1 Then
            If CheckBillsEmpty Then Bill.Active = True:  Call ClearBillRows
        End If
 
        sta.Panels(Pan.C2提示信息) = ""
        lblTotal.Caption = "合计:"
        
        '收费保持病人ID
        If (mbytInFun = 0 Or mbytInFun = 1) And txtPatient.Text = mstrPrePati And mlngPrePati <> 0 Then
            strPati = "-" & mlngPrePati
        Else
            strPati = txtPatient.Text
        End If
        
        If IDKind.GetCurCard.名称 Like "IC卡*" And IDKind.GetCurCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        If IDKind.GetCurCard.名称 Like "*身份证*" And IDKind.GetCurCard.系统 Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        int上次病人来源 = gint病人来源
        
        '50200(防止窗口找开过长,发生时间与登记时间拉得过长)
        If mbln补费 And mstr最后转科时间 <> "" Then
            txtDate.Text = Format(CDate(mstr最后转科时间) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
        Else
            txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        If Not mobjBill Is Nothing Then mobjBill.发生时间 = CDate(txtDate.Text)
                
        'a.根据输入读取病人信息失败
        If Not GetPatient(strPati, blnCancel, blnCard) Then
        
            Call InitBalanceGrid(True)
            If blnCancel Then '取消输入
                If Visible Then txtPatient.SetFocus
                txtPatient.Text = ""
                Exit Sub
            End If
            
            If blnCard Then
                MsgBox "不能确定" & gstrCustomerAppellation & "信息，请检查是否正确刷卡！", vbInformation, gstrSysName
                Call ClearPatientInfo(True)
                Exit Sub
            Else
                '门诊收费、划价可以手动输入病人信息(允许时)。
                If gint病人来源 = 1 And gblnInputName And (mbytInFun = 0 Or mbytInFun = 1) And IDKind.IDKind = IDKind.GetKindIndex("姓名") And txtPatient.Text <> "" Then
                    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" Then
                        If Not CheckRegisted(0) Then
                           Call ClearPatientInfo(True): Exit Sub
                        End If
                    End If
                    If mbytInFun = 0 And mbytInState = 0 Then
                        '问题:29283
                         '  -- 参数:调用场合-1-挂号;2-收费
                         '  --        病人id_In-病人ID(未建档的,传入零)
                         '  --        卡号_In: 刷卡卡号;未刷卡时,为空
                         '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
                         If zlPatiCardCheck(2, 0, IIf(blnCard Or blnICCard, txtPatient.Text, ""), 1) = False Then
                               Call ClearPatientInfo(True): Exit Sub
                         End If
                    End If
                    sta.Panels(Pan.C2提示信息) = "输入的标识不能读取" & gstrCustomerAppellation & "信息，将默认为新" & gstrCustomerAppellation & "姓名！"
                    Call ClearmobjBill
                    
                    If mbytInFun = 0 And mbytInState = 0 And Not mblnValid And Visible And mstrInNO = "" And txtIn.Text = "" Then
                        Call LoadAddedItem(0, txtPatient.Text)
                    End If
                    
                    If (mbytInFun = 0 Or mbytInFun = 1) And mobjBill.Pages(mintPage).NO = "" Then
                        cbo费别.Locked = False: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
                        If Not mblnValid And Not (Bill.Active And txtPatient.Text = mstrPrePati And txtPatient.Text <> "") Then '同一个病人不处理
                            Call LoadAndSeek费别
                        End If
                    End If
                    cbo医疗付款.Locked = False
                    Call ShowPrePayInfo(False) '预交信息初始
                    mobjBill.姓名 = txtPatient.Text
                    Call Set连续收费操作(True)
                    
                    If txtPatient.Text = mstrPrePati And txtPatient.Text <> "" Then '同一个收费病人,此时没有病人ID
                        mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
                        mobjBill.性别 = zlStr.NeedName(cboSex.Text)
                        mobjBill.费别 = zlStr.NeedName(cbo费别.Text)
                                                
                        If Bill.Active Then
                            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, mstrPreDoctor, True)) '不触发click事件
                            Call cbo开单人_Click
                        End If
                        If Not mblnValid And Visible Then Bill.SetFocus
                        
                        Exit Sub
                    Else
                        '清除医生
                        If gbyt科室医生 = 0 And CheckBillsEmpty Then
                            For i = 1 To mobjBill.Pages.Count
                                mobjBill.Pages(i).开单部门ID = 0: mobjBill.Pages(i).开单人 = ""
                            Next
                            cbo开单人.ListIndex = -1: cbo开单科室.ListIndex = -1: lblDuty.Caption = ""
                        End If
                        
                        '取消了医保信息初始,因为NewBill中已初始过了
                                                           
                        Call ClearPatientInfo   '清除年龄,门诊号,初始年龄单位
                        '刘兴洪:22343 gbln缴款结束改为gTy_Module_Para.byt缴款控制 = 1
                        If Not (mbytInFun = 0 And gTy_Module_Para.byt缴款控制 = 1) _
                            Or mstrPrePati = "" Then
                            Call ClearPayInfo
                            Call InitCommVariable
                            Call ClearMoney
                            If CheckBillsEmpty Then
                                mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
                                Call ClearTotalInfo(True)
                            Else
                                Call ShowMoney
                            End If
                        End If
                        Call ReInitPatiInvoice
                        mblnNotValied = True
                        If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                    mblnNotValied = False
                        If Not mblnValid Then Call ShowWelcomeByLed
                        Exit Sub
                    End If   '同一个收费病人
                    
                Else '记帐必须有病人信息
                    MsgBox "请检查输入内容,不能读取" & gstrCustomerAppellation & "信息！", vbInformation, gstrSysName
                    Call ClearPatientInfo(True)
                    Exit Sub
                End If
            End If
            
        Else 'b.根据输入读取病人信息成功
            lng病人ID = Val("" & mrsInfo!病人ID)
            Call InitBalanceGrid(True)
            Call Set连续收费操作
            
            If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gint病人来源 = 1 Then
                If Not CheckRegisted(lng病人ID) Then
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                End If
            End If
            If mbytInFun = 0 And mbytInState = 0 Then
                '问题:29283
                 '  -- 参数:调用场合-1-挂号;2-收费
                 '  --        病人id_In-病人ID(未建档的,传入零)
                 '  --        卡号_In: 刷卡卡号;未刷卡时,为空
                 '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
                 If zlPatiCardCheck(2, lng病人ID, IIf(blnCard Or blnICCard, txtPatient.Text, ""), 1) = False Then
                    '恢复上次病人来源
                    If int上次病人来源 <> gint病人来源 And mTy_Para.bln住院病人门诊收费 = False Then
                        '问题:27364 日期:2010-01-13 15:27:50
                        Call zlChangePatiSource(int上次病人来源)
                    End If
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                     Exit Sub
                 End If
            End If
            
            '就诊卡密码检查
            If mbytInState = 0 And ((blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.接口序号 <> 0) And mstrPassWord <> "") Then
                i = Nvl(Choose(mbytInFun + 1, 3, 2, 4), 99)
                If Mid(gstrCardPass, i, 1) = "1" Then
                    If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
                        '恢复上次病人来源
                        If int上次病人来源 <> gint病人来源 And mTy_Para.bln住院病人门诊收费 = False Then
                            '问题:27364 日期:2010-01-13 15:27:50
                            Call zlChangePatiSource(gint病人来源)
                        End If
                        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                    End If
                End If
            
            End If
                
            '连续划价或收费时,不是同一个病人时，记帐没有保留病人信息
            If Not IIf(mlngPrePati = 0, mstrPrePati = "" & mrsInfo!姓名, mlngPrePati = lng病人ID) Then
                '清除医生
                If mbytInState = 0 And mstrInNO = "" Then
                    If gbyt科室医生 = 0 And CheckBillsEmpty Then
                        For i = 1 To mobjBill.Pages.Count
                            mobjBill.Pages(i).开单部门ID = 0: mobjBill.Pages(i).开单人 = ""
                        Next
                        cbo开单人.ListIndex = -1: cbo开单科室.ListIndex = -1: lblDuty.Caption = ""
                    End If
                End If
                
                Call ClearPatientInfo
                
                '刘兴洪:22343
                If Not (mbytInFun = 0 And gTy_Module_Para.byt缴款控制 = 1) _
                    Or mstrPrePati = "" Then
                    Call ClearPayInfo
                    Call InitCommVariable
                    Call ClearMoney
                    If CheckBillsEmpty Then
                        mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
                        Call ClearTotalInfo(True)
                    Else
                        Call ShowMoney
                    End If
                End If
            End If
                
            '开单人与开单科室
            '    新增单据并且根据挂号单输入时才有执行部门ID
            If IsRegisterDept Then
                If IsNull(mrsInfo!姓名) Then '没有建档,但挂了号,根据挂号单读开单人和开单科室
                    Call SetDeptDoctorByRegevent(0, txtPatient.Text)
                    sta.Panels(Pan.C2提示信息) = "该病人挂号时没有登记档案,请输入病人姓名！"
                    Call ClearPatientInfo(True)
                    
                    Set mrsInfo = New ADODB.Recordset
                    If Not mblnValid And Visible Then txtPatient.SetFocus
                    Exit Sub
                Else
                    Call Set开单人开单科室Click(mrsInfo!执行人 & "", Val("" & mrsInfo!执行部门ID))
                End If
            ElseIf gint病人来源 = 2 Then
                If gbyt科室医生 <> 0 And mbytInState = 0 And mstrInNO = "" Then
                    '仅新增单据时,取住院病人的开单部门:科室确定医生或各自独立输入
                    If mlngDeptID = 0 Then
                        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, Val(Nvl(mrsInfo!当前科室id))))
                    Else
                        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, mlngDeptID))
                    End If
                    Call cbo开单科室_Click
                End If
            ElseIf gint病人来源 = 1 Then
                If mbytInState = 0 And mstrInNO = "" Then
                    If gbyt科室医生 <> 0 And mlng主页ID <> 0 Then
                        If mlngDeptID = 0 Then
                            Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, Val(Nvl(mrsInfo!当前科室id))))
                        Else
                            Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, mlngDeptID))
                        End If
                        Call cbo开单科室_Click
                    Else
                        Call SetDeptDoctorByRegevent(lng病人ID) '根据病人挂号信息设置开单科室和医生
                    End If
                End If
            End If
            
            If mbytInFun = 2 Then
                If Not IsNull(mrsInfo!工作单位) Then
                    lblCorp.Visible = True: lblCorp.Caption = "工作单位:" & mrsInfo!工作单位
                Else
                    lblCorp.Visible = False: lblCorp.Caption = ""
                End If
            End If
             
            '病人预交款信息
            If lng病人ID <> 0 Then Call LoadFeeInfor(lng病人ID)
            
            lbl险类.Caption = "" & mrsInfo!险类名称
            txtPatient.Text = "" & mrsInfo!姓名
            txtPatient.PasswordChar = ""
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
            cboSex.ListIndex = cbo.FindIndex(cboSex, Nvl(mrsInfo!性别), True)
            txt门诊号.Text = "" & mrsInfo!门诊号
            
            Call LoadOldData("" & mrsInfo!年龄, txt年龄, cbo年龄单位)
            If Not IsNull(mrsInfo!出生日期) Then
                 txt年龄.Text = ReCalcOld(mrsInfo!出生日期, cbo年龄单位, lng病人ID)
            End If
            
            If glngSys Like "8??" Or mbytInFun = 2 Then
                cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(mrsInfo!费别), True)
                cbo费别.Locked = mbytInFun = 2: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
            ElseIf Not mblnValid Then
                If IsRegisterDept And cbo开单科室.ListIndex <> -1 Then
                    cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(mrsInfo!费别), True) '挂号时确定的费别
                Else
                    If mstrInNO = "" Then Call LoadAndSeek费别
                End If
            End If
            If gstr费别 <> "" And cbo费别.ListIndex = -1 Then cbo费别.ListIndex = cbo.FindIndex(cbo费别, gstr费别, True)
            
            cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, Nvl(mrsInfo!医疗付款方式), True)
            If mstr付款方式 <> "" And cbo医疗付款.ListIndex = -1 Then cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, mstr付款方式, True)
            cbo医疗付款.Locked = mbytInFun = 2 Or gint病人来源 = 2
            

            '设置对象中的病人信息
            With mobjBill
                .病人ID = lng病人ID
                .主页ID = IIf(mbln补费 And mlng主页ID <> 0, mlng主页ID, Nvl(mrsInfo!主页ID, 0))
                .标识号 = IIf(gint病人来源 = 2, Nvl(mrsInfo!住院号, 0), Nvl(mrsInfo!门诊号, 0))
                .姓名 = "" & mrsInfo!姓名
                .性别 = "" & mrsInfo!性别
                .年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
                .床号 = "" & mrsInfo!当前床号
                .病区ID = IIf(mbln补费 And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!当前病区ID)))
                .科室ID = IIf(mbln补费 And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!当前科室id)))
                .费别 = zlStr.NeedName(cbo费别.Text) '以当前有效为准
            End With
            Call ReInitPatiInvoice
            
            '关联操作处理
            If Not mblnValid And Visible Then
                If mbytInFun = 0 Or mbytInFun = 1 Then
                    '不是同一个病人时
                    If Not (IIf(mlngPrePati = 0, mstrPrePati = mobjBill.姓名, mlngPrePati = mobjBill.病人ID) And txtPatient.Text <> "") Then
                         Call AddCardFee '产生就诊卡费用行
                    End If
                    
                    '读取病人的多张划价单
                    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And txtIn.Text = "" Then
                        If mobjBill.病人ID <> 0 Then
                            If gblnCheckRegeventDept And gint病人来源 = 1 And IsRegisterDept Then lng挂号科室 = Val("" & mrsInfo!执行部门ID)
                           blnHavePriceBill = LoadMultiBills(mobjBill.病人ID, InStr(1, mstrPrivs, "普通病人多单据收费") = 0, lng挂号科室, blnCard)
                        End If
                        Call LoadAddedItem(mobjBill.病人ID, mobjBill.姓名)
                    End If
                End If
                '光标定位
                If mstrInNO = "" Then
                    If mbytInFun = 0 And mbytInState = 0 And txtPatient.Text = "新病人" Then
                        txtPatient.SetFocus
                        Call txtPatient_GotFocus
                    Else
                        If cbo医疗付款.ListIndex = -1 And gbln医疗付款 And mbytInFun <> 2 Then '记帐不允许更改费别、付款方式
                            If cbo医疗付款.Enabled And cbo医疗付款.Visible Then cbo医疗付款.SetFocus
                        Else
                            '问题:39253
                            If gbln划价立即缴款 = False And blnHavePriceBill Then
                                If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
                            End If
                            
                            If gbyt科室医生 = 0 Then
                                If blnHavePriceBill Then
                                    If cbo开单人.Enabled And cbo开单人.Visible And cbo开单人.ListIndex < 0 Then cbo开单人.SetFocus
                                Else
                                    If cbo开单人.Enabled And cbo开单人.Visible Then cbo开单人.SetFocus
                                End If
                            ElseIf glngSys Like "8??" Then
                                Bill.SetFocus
                            Else
                                If blnHavePriceBill Then
                                    If cbo开单科室.Enabled And cbo开单科室.Visible And cbo开单科室.ListIndex < 0 Then cbo开单科室.SetFocus
                                Else
                                    If cbo开单科室.Enabled And cbo开单科室.Visible Then cbo开单科室.SetFocus
                                End If
                            End If
                        End If
                    End If
                    
                    Call ShowWelcomeByLed
                End If
            End If
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddCardFee()
'功能:产生就诊卡费用行
    Dim objDetail As Detail, lngDoUnit As Long
        
    If mbytInFun = 0 And mstrCardNO = "" And Bill.Active Then
        Set objDetail = ReadPatiCardObj(mobjBill.病人ID, mstrCardNO)
        
        If mstrCardNO <> "" And Not objDetail Is Nothing Then
            If Not ItemExist(objDetail.ID) Then
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Rows - 1 Then
                    Bill.Rows = Bill.Rows + 1
                    mblnNewRow = True: Call bill_AfterAddRow(Bill.Rows - 1): mblnNewRow = False
                End If
                Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = "" '有必要加上
                
                lngDoUnit = mobjBill.科室ID
                If lngDoUnit = 0 Then lngDoUnit = Get开单科室ID
                
                lngDoUnit = Get收费执行科室ID(objDetail.类别, objDetail.ID, objDetail.执行科室, lngDoUnit, Get开单科室ID, _
                            gint病人来源, , , , , mobjBill.病区ID)
                
                Call SetDetail(objDetail, Bill.Rows - 1, lngDoUnit)
                Call CalcMoneys(mintPage, Bill.Rows - 1)
                Call ShowDetails(Bill.Rows - 1)
                Call ShowMoney
            End If
        End If
    End If
End Sub


Private Sub ShowWelcomeByLed()
'功能:显示欢迎信息和病人信息
    Dim strInfo As String, lngPatient As Long

    If mbytInFun = 0 And mbytInState = 0 And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModul, gcnOracle
        End If
        
        strInfo = Trim(txtPatient.Text)
        If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val("" & mrsInfo!病人ID)
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub
Private Function GetPatient(ByVal strInput As String, Optional blnCancel As Boolean, Optional ByVal blnCard As Boolean, Optional blnYbCheckCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCancel=用于表示输入取消
    '       blnCard=表示是否就诊卡刷卡
    '       blnYbCheckCard-医保生身验卡(24689)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 16:43:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strMoney As String, strWhere As String, strPati As String
    Dim rsTmp As ADODB.Recordset, strTemp As String, strTempYb As String
    Dim bln挂号 As Boolean
    Dim vRect As RECT
    
    bln挂号 = False: mblnNotClearLedDisplay = False
    mlngPreBrushCard = 0: mlngCardTypeID = 0
    
ReDO:
    blnCancel = False
    mstrWarn = "" '记帐分类报警
    If mbytInFun = 2 Then
        strMoney = "zl_PatiDayCharge(A.病人ID) as 当日额, Zl_Patiwarnscheme(A.病人id) As 适用病人,"
    End If
    
    '输入病人的权限
    If mbytInState = 0 And mstrInNO = "" Then  '新增
        If gint病人来源 = 2 Then
            strWhere = " And Nvl(A.当前科室ID,0)<>0"
        End If
    End If
    
    '读取病人信息
    If mbln补费 And mlng主页ID <> 0 Then
        strSQL = _
        " Select " & strMoney & "Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID," & _
        "        nvl(b1.病人类型,A.病人类型) as 病人类型,A.险类," & _
        "        Nvl(b1.主页ID,A.主页ID) as 主页ID," & _
        "        A.IC卡号,A.就诊卡号,A.卡验证码,A.门诊号,Nvl(B1.住院号,A.住院号) As 住院号,nvl(B1.姓名,A.姓名) as 姓名," & _
        "        nvl(b1.性别,A.性别) as 性别,nvl(b1.年龄,A.年龄) as 年龄,C.名称 As 险类名称, A.出生日期," & _
        "        nvl(b1.费别,A.费别) as 费别,A.担保额,nvl(b1.医疗付款方式,A.医疗付款方式) as 医疗付款方式," & _
        "        A.工作单位,nvl(b1.当前病区ID,A.当前病区ID) as 当前病区ID,nvl(b1.出院科室ID,A.当前科室ID) as 当前科室ID," & _
        "        nvl(b1.出院病床,A.当前床号) as 当前床号,A.在院," & _
        "        Decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.入院日期" & _
        " From 病人信息 A,病案主页 B1,保险类别 C  " & _
        " Where A.险类 = C.序号(+) And A.病人ID=B1.病人ID(+) And B1.主页ID = [4] And A.停用时间 is NULL"
    Else
        strSQL = _
        " Select " & strMoney & "Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型,A.险类," & _
                 IIf(gint病人来源 = 1, "NULL", "Decode(A.当前科室ID,NULL,NULL,A.主页ID)") & " as 主页ID,A.IC卡号,A.就诊卡号,A.卡验证码,A.门诊号,A.住院号,A.姓名," & _
        "        A.性别,A.年龄,C.名称 险类名称, A.出生日期,A.费别,A.担保额,A.医疗付款方式,A.工作单位,A.当前病区ID,A.当前科室ID,A.当前床号,A.在院," & _
        "        decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.入院日期" & _
        " From 病人信息 A,病案主页 B1,保险类别 C  " & _
        " Where A.险类 = C.序号(+) And A.病人ID=B1.病人ID(+) And A.主页ID=B1.主页ID(+) And A.停用时间 is NULL"
    End If

    If blnYbCheckCard = False And blnCard And IDKind.GetCurCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then   '103563
        If gint病人来源 = 1 And Not gblnInputCard Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
      
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        mlngCardTypeID = lng卡类别ID
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & strWhere & " And A.病人ID=[1] "
        mlngPreBrushCard = lng卡类别ID
        
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Or blnYbCheckCard Then '病人ID
        If gint病人来源 = 1 And (Not gblnInputID And mstrYBPati = "") _
            And Not (mstrInNO <> "" And mbytInState = 0) Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        If gint病人来源 = 1 And Not gblnInputID Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-15,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        If gint病人来源 = 1 And Not gblnInputID Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "." Then '挂号单号(最后为执行部门ID以区分)
        If gint病人来源 = 1 And (mbytInFun = 0 Or mbytInFun = 1) And Not gblnInputNO Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        bln挂号 = True
        '按日或年顺序编号规则
        strInput = UCase(GetFullNO(Mid(strInput, 2), 12))
        txtPatient.Text = strInput
        
        '门诊记帐时必须要挂号建档
        '如果是出院病人,则通过设置主页ID为0来为将来保存是通过它来识别是门诊费用,注意:最后一个字段执行部门ID在patient_keypress中会用到
        '76451,冉俊明,2014-8-19
        strSQL = "" & _
            "   Select " & strMoney & "Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型,A.险类," & _
                                IIf(gint病人来源 = 1, "NULL", "Decode(A.当前科室ID,NULL,NULL,A.主页ID)") & " as 主页ID,A.就诊卡号,A.卡验证码,Nvl(B.标识号,A.门诊号) as 门诊号," & _
            "               A.住院号,B.姓名,B.性别,B.年龄,C.名称 险类名称, A.出生日期,B.费别,A.担保额,A.医疗付款方式,A.工作单位,A.当前病区ID,A.当前科室ID,A.当前床号,B.执行人,B.执行部门ID,A.在院," & _
            "               decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.入院日期" & _
            " From 病人信息 A,病案主页 B1,门诊费用记录 B,保险类别 C " & _
            " Where B.病人ID=A.病人ID" & IIf(mbytInFun = 2, "", "(+)") & _
            "            And A.病人ID=B1.病人ID(+) And A.主页ID=B1.主页ID(+)  " & _
            "           And A.险类 = C.序号(+) And B.记录性质=4 And B.记录状态=1 " & _
            zlGetRegEventsCons("加班标志", "B") & _
            strWhere & " And B.NO=[2] And Rownum<2"
    Else
        If mrsInfo.State = 1 Then
            If mrsInfo!姓名 = strInput Then GetPatient = True: Exit Function
        End If
        mlngCardTypeID = IDKind.GetCurCard.接口序号
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '通过姓名模糊查找病人(允许输入病人标识时)
                If Not mblnValid And gblnSeekName And (mbytInFun <> 2 And gblnInputID Or mbytInFun = 2) Then
                    strPati = _
                        " Select /*+Rule */1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                                    IIf(gint病人来源 = 2 And mbytInFun <> 2, "A.住院号,B.名称 as 科室,A.当前床号 as 床号,", "A.门诊号,") & _
                        "           A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                        " From 病人信息 A,部门表 B" & _
                        " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And Rownum <101 " & strWhere & " And A.姓名 Like [1]" & _
                        IIf(gintNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                    
                    If mbytInFun = 2 And gblnOnlyUnitPatient Then
                        strPati = strPati & " And A.合同单位ID Is Not Null"
                    End If
                    
                    '门诊病人收费时可以不对应病人档案
                    If gint病人来源 = 1 And mbytInFun <> 2 Then
                        strPati = strPati & " Union ALL " & _
                            "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                    End If
                    strPati = strPati & " Order by 排序ID,姓名"
                        
                    vRect = zlControl.GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人" & mbytInFun & gint病人来源, 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays, "bytSize=1")
                    If Not rsTmp Is Nothing Then
                        If rsTmp!ID = 0 Then '当作新病人
                            strSQL = ""
                        Else '以病人ID读取
                            '85187,冉俊明,2015-05-27,在院病人门诊收费时进行模糊查找找不到病人信息（病人来源设置的是"门诊病人"）
                            strInput = "-" & rsTmp!病人ID
                            strSQL = strSQL & strWhere & " And A.病人ID=[1]"
                        End If
                    Else '取消选择
                        strSQL = ""
                    End If
                Else
                    strSQL = ""
                End If
            Case "医保号"
                strInput = UCase(strInput)

                If MCPAR.blnOnlyBjYb And zlCommFun.ActualLen(strInput) >= 9 Then
                    '仅北京医保才有效:见问题:问题:27331
                    strSQL = strSQL & strWhere & "  And A.医保号 like [3] "
                    strTemp = Left(strInput, 9) & "%"
                Else
                     strSQL = strSQL & strWhere & "  And A.医保号=[2]"
                End If
                
                'strSQL = strSQL & strWhere & " And A.医保号=[2]"
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                 If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                 strInput = "-" & lng病人ID
                 blnHavePassWord = True
                strSQL = strSQL & strWhere & " And   A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & strWhere & " And A.病人ID=[1]"
               blnHavePassWord = True
            Case "门诊号"
                If gint病人来源 = 1 And Not gblnInputID Then
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & strWhere & " And A.门诊号=[2]"
                '75087,冉俊明,2014-7-15,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If gint病人来源 = 1 And Not gblnInputID Then
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If IDKind.GetCurCard.接口序号 > 0 Then
                    lng卡类别ID = IDKind.GetCurCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                    mlngPreBrushCard = lng卡类别ID
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
    If strSQL <> "" Then
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, strTemp, mlng主页ID)
        If Not mrsInfo.EOF Then
            '75259：李南春,2014-7-10，病人姓名的显示颜色处理
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), Me.ForeColor, vbRed))
            If gint病人来源 = 1 And mTy_Para.bln住院病人门诊收费 = False Then
                '需要检查是否为在院病人
                '问题:27364 日期:2010-01-13 15:27:50
                If Val(Nvl(mrsInfo!在院)) = 1 Then
                        If gbln病人来源受权限控制 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then
                            '29720
                            '不能转换病人
                            Call MsgBox("该病人是在院病人,不能进行收费(划价或记帐)操作!)", vbOKCancel + vbInformation + vbDefaultButton1, gstrSysName)
                            Set mrsInfo = New ADODB.Recordset
                            Exit Function
                        End If
                    '此为在院病人,自动到在院状态
                    mblnAutoChangePati = True
                    gint病人来源 = 2: Call zlChangePatiSource(gint病人来源)
                    Set mrsInfo = New ADODB.Recordset
                     GoTo ReDO:
                End If
                strWhere = ""
            End If
            '对异常单据进行收费
            If PatiErrBillPay(Val(Nvl(mrsInfo!病人ID))) Then
                Call ClearBillRows: Call ClearMoney
                Call ClearTotalInfo(True)
                NewBill True
                blnCancel = True
                Exit Function
            End If
            If mlng病人ID <> mrsInfo!病人ID Then mlng关联医嘱 = 0
            
            GetPatient = True
            mstrPassWord = strPassWord
            If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        Else
            Set mrsInfo = New ADODB.Recordset
            If bln挂号 Then
                 txtPatient.Text = "": GetPatient = False
            End If
        End If
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    '问题:60010
    IDKind.SetAutoReadCard (False)
    zlCommFun.OpenIme False
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 And Trim(txtPatient.Text) <> "" Then
        mobjBill.姓名 = txtPatient.Text
        mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
        mobjBill.性别 = zlStr.NeedName(cboSex.Text)
    End If
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
    If mblnKeyReturn = False Then
        mblnValid = True: Call txtPatient_KeyPress(13): mblnValid = False
    Else
        mblnKeyReturn = False
    End If
End Sub

Private Sub txtRePrint_GotFocus()
    Call zlControl.TxtSelAll(txtRePrint)
End Sub

Private Sub txtRePrint_KeyPress(KeyAscii As Integer)
    Dim strNos As String, strOper As String, vDate As Date, intInsure As Integer, blnVirtualPrint As Boolean
    Dim lng结帐ID As Long, lng病人ID As Long
    Dim strReclaimInvoice As String, intInvoiceFormat As Integer '回收的票据
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtRePrint, KeyAscii)
    Else
        '重打
        txtRePrint.Text = GetFullNO(txtRePrint.Text, 13)
        zlControl.TxtSelAll txtRePrint
        
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("门诊费用记录", txtRePrint.Text, , "1", Me.Caption) Then
            If Not ReturnMovedExes(txtRePrint.Text, 1, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        If Not ReadBillInfo(1, txtRePrint.Text, 1, strOper, vDate, lng病人ID) Then
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
        
        '可能是多单据收费中的一张
        If gTy_Module_Para.byt票据分配规则 <> 0 Then
            strNos = GetMultiNOs(txtRePrint.Text, , , True)
        Else
            strNos = GetMultiNOs(txtRePrint.Text)
        End If
        
        '单据有剩余数量的才可以重打
        If Not BillExistMoney(strNos, 1, True) Then
            MsgBox "单据不存在或已经全部退费,不能重打！", vbInformation, gstrSysName
            txtRePrint.Text = "": Exit Sub
        End If
        '进行了医保补充结算，不允许重打和补打
        If CheckBillExistReplenishData(1, , Replace(strNos, "'", "")) = True Then
            MsgBox "当前单据进行了医保补充结算，不允许重打票据！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '调出重打的单据显示
        If frmMultiBills.ShowMe(Me, 0, mstrPrivs, txtRePrint.Text, "", True) = False Then Exit Sub
        intInsure = ChargeExistInsure(txtRePrint.Text, , lng结帐ID)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
            '此处只提供了收费票据的重打
        End If
        
        Call ReInitPatiInvoice(True, intInsure, lng病人ID)
        strReclaimInvoice = zlGetReclaimInvoice(txtRePrint.Text)
        If strReclaimInvoice <> "" Then
            '需要显示出本次需要回收的发票
            If MsgBox("注意:" & vbCrLf & " 请注意回收以下发票:" & vbCrLf & strReclaimInvoice, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Call RefreshFact '刷新票据号
                txtRePrint.Text = ""
                txtPatient.SetFocus
                Exit Sub
            End If
        End If
        intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
        Dim strPriceGrade As String
        If gintPriceGradeStartType >= 2 Then
            strPriceGrade = GetPriceGradeFromNos(strNos)
        Else
            strPriceGrade = mstr普通价格等级
        End If
        If Not RePrintCharge(1, strNos, Me, mlng领用ID, strReclaimInvoice, , , _
            intInvoiceFormat, blnVirtualPrint, , mlngShareUseID, mstrUseType, , strPriceGrade) Then
            txtRePrint.SetFocus
        Else
        
            '银医一卡通写卡，85950
            Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, 0, strNos)
            
            Call RefreshFact '刷新票据号
            txtRePrint.Text = ""
            txtPatient.SetFocus
        End If
    End If
End Sub

Private Sub txtRePrint_LostFocus()
    txtRePrint.BackColor = vbWhite
End Sub
Public Function GetMustPaySum() As Currency
'功能：求本次收费的应缴合计，主要用于多单据收费模式
    Dim curMoney As Currency, i As Integer
    For i = 1 To mobjBill.Pages.Count
        curMoney = curMoney + mobjBill.Pages(i).应缴金额
    Next
    GetMustPaySum = curMoney
End Function

Private Function Get中药数量(ByRef str计算单位 As String) As Long
'功能：取当前单据中中药的数量，如果存在不同单位的药品，则返回为0
    Dim i As Integer, str单位 As String
    
    Get中药数量 = 0
    With mobjBill.Pages(mintPage)
        For i = 1 To .Details.Count
            If .Details(i).收费类别 = "7" Then
                If gbln药房单位 Then
                    If str单位 <> "" And str单位 <> .Details(i).Detail.药房单位 Then
                        str单位 = "不同单位"
                        Exit For
                    Else
                        If str单位 = "" Then str单位 = .Details(i).Detail.药房单位
                    End If
                Else
                    If str单位 <> "" And str单位 <> .Details(i).计算单位 Then
                        str单位 = "不同单位"
                        Exit For
                    Else
                        If str单位 = "" Then str单位 = .Details(i).计算单位
                    End If
                End If
                
                Get中药数量 = Get中药数量 + .Details(i).付数 * .Details(i).数次
            End If
        Next
    End With
    If str单位 = "不同单位" Then
        Get中药数量 = 0
    Else
        str计算单位 = str单位
    End If
End Function

Private Sub AutoBultBookFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动生成工作费或自动分单据
    '编制:刘兴洪
    '日期:2011-08-16 10:34:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurOneCard As Currency
   ' If txt缴款.Tag = "退出" Then txt缴款.Tag = "": Exit Sub
    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gbytAutoSplitBill > 0 And Not (mstrYBPati <> "" And MCPAR.门诊预结算) Then
        Call AutoSplitBill
    End If
    '收费时自动产生工本费项目:修改时不管工本费
    If mbytInFun = 0 And mbytInState = 0 And gTy_Module_Para.bln工本费 Then
        If Not CheckBillsEmpty Then Call SetFactMoney
    End If
End Sub
 

Private Sub CalcMoneys(Optional intPage As Integer, Optional lngRow As Long)
'功能：计算或重新计算指定行或所有行的金额
'参数：intPage,lngRow=指定单据页指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, p As Integer
    Dim strMainRows As String
    Dim bln从项汇总折扣 As Boolean
        
    'If CheckBillsEmpty Then Exit Sub   '此处不必再判断,在改变费别(包括记帐病人变动)和改变加班状态,调用时判断
    
    Screen.MousePointer = 11
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, mobjBill.Pages.Count, intPage)
        strMainRows = ""
        If mobjBill.Pages.Count >= p Then
            For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Pages(p).Details.Count, lngRow)
                If mobjBill.Pages(p).Details.Count >= i Then
                    
                    bln从项汇总折扣 = False
                    If gbln从项汇总折扣 Then                    '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
                        If mobjBill.Pages(p).Details(i).从属父号 > 0 Then    '从项
                            bln从项汇总折扣 = Not mobjBill.Pages(p).Details(mobjBill.Pages(p).Details(i).从属父号).Detail.屏蔽费别
                            If bln从项汇总折扣 And lngRow <> 0 Then strMainRows = strMainRows & "," & mobjBill.Pages(p).Details(i).从属父号      '单独计算一行的时候
                        Else
                            If CheckMainItem(i, p) Then                          '主项或独立项
                                 bln从项汇总折扣 = Not mobjBill.Pages(p).Details(i).Detail.屏蔽费别
                                 If bln从项汇总折扣 Then strMainRows = strMainRows & "," & i  '一页可能有多个主从项,先记录主项行号,后面再重算主项折扣
                            End If
                        End If
                    End If
                            
                    Call CalcMoney(p, i, bln从项汇总折扣)
                End If
            Next
        
            '重算所有主项
            If gbln从项汇总折扣 Then
                For i = 1 To UBound(Split(strMainRows, ","))
                    Call CalcPItemActualIncome(Split(strMainRows, ",")(i), p)
                Next
            End If
        End If
    Next
    
    Screen.MousePointer = 0
End Sub

Private Sub CalcMoney(intPage As Integer, lngRow As Long, Optional bln从项汇总折扣 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算或重新计算指定行的金额
    '入参:intPage=指定页单据,lngRow=指定行
    '编制:刘兴洪
    '日期:2014-06-06 18:02:30
    '说明：1.ExpenseBill集合的索引对应单据的行号
    '      2.变价只能对应一个收入项目:mobjBill.Pages(intPage).Details(lngRow).InComes(1)
    '      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
    '      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strInfo As String, strAdvance As String
    Dim rsTmp As ADODB.Recordset
    Dim dblMoney As Double '用户输入的变价金额
    Dim str费别 As String
    Dim dblAllTime As Double, dbl加班加价率 As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    
    If mobjBill.Pages.Count < intPage Then Exit Sub
    If mobjBill.Pages(intPage).Details.Count < lngRow Then Exit Sub
    
    If InStr(",5,6,7,", mobjBill.Pages(intPage).Details(lngRow).收费类别) > 0 Then
        strPriceGrade = mstr药品价格等级
    ElseIf mobjBill.Pages(intPage).Details(lngRow).收费类别 = "4" Then
        strPriceGrade = mstr卫材价格等级
    Else
        strPriceGrade = mstr普通价格等级
    End If
    
    If InStr(",4,5,6,7,", mobjBill.Pages(intPage).Details(lngRow).收费类别) > 0 Then
        Call AdjustCpt(mobjBill.Pages(intPage).Details(lngRow).收费细目ID)
    End If
    
    If strPriceGrade <> "" Then
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
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,B.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID=A.ID And C.ID=B.收入项目ID " & _
        " And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        "       And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Pages(intPage).Details(lngRow).收费细目ID, strPriceGrade)
    If Not rsTmp.EOF Then
        With mobjBill.Pages(intPage).Details(lngRow)
            If InStr(",5,6,7,", .收费类别) > 0 Or (.收费类别 = "4" And .Detail.跟踪在用) Then
                '计算药品时价(分批或不分批),必然有记录(输入该项目时已判断)
                dblAllTime = .付数 * .数次
                If gbln药房单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                    dblAllTime = dblAllTime * .Detail.药房包装 '库存时价按售价数量进行计算
                End If
                
                If dblAllTime <> 0 Or Not .Detail.变价 Then
                    If .Detail.批次 <= 0 Then
                        gstrSQL = "Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual"
                    Else
                        gstrSQL = "Select Zl_Fun_Getprice([1],[2],[3],[4],[5]) As Price From Dual"
                    End If
                    Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .收费细目ID, .执行部门ID, dblAllTime, 0, .Detail.批次)
                    If rsPrice.EOF Then
                        '获取价格失败
                        If InStr(",5,6,7,", .收费类别) > 0 Then
                            MsgBox "第 " & lngRow & " 行药品""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                        Else
                            MsgBox "第 " & lngRow & " 行卫生材料""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                        End If
                    Else
                        strPrice = Nvl(rsPrice!Price) & "|||"
                        varPrice = Split(strPrice, "|")
                        dblMoney = Val(varPrice(0))
                        dbl剩余数量 = Val(varPrice(2))
                        
                        If dbl剩余数量 <> 0 And .Detail.变价 Then
                            '数量未分解完毕
                            If InStr(",5,6,7,", .收费类别) > 0 Then
                                MsgBox "第 " & lngRow & " 行时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            Else
                                MsgBox "第 " & lngRow & " 行时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            End If
                            dblMoney = 0
                        End If
                    End If
                Else
                    dblMoney = 0
                End If
            Else
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
            End If
        End With
        
        '再清除原有记录
        Set mobjBill.Pages(intPage).Details(lngRow).InComes = New BillInComes
        
        '填写现有费用记录
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .收入项目ID = rsTmp!收入项目ID
                .收入项目 = rsTmp!名称
                .收据费目 = Nvl(rsTmp!收据费目)
                .原价 = Val(Nvl(rsTmp!原价))
                .现价 = Val(Nvl(rsTmp!现价))
                
                If InStr(",5,6,7,", mobjBill.Pages(intPage).Details(lngRow).收费类别) > 0 Then
                    If gbln药房单位 Then
                        .标准单价 = Format(dblMoney * mobjBill.Pages(intPage).Details(lngRow).Detail.药房包装, gstrFeePrecisionFmt)
                    Else
                        .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                    End If
                Else
                    If mobjBill.Pages(intPage).Details(lngRow).Detail.变价 Then
                        .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                    Else
                        .标准单价 = Format(Nvl(rsTmp!现价, 0), gstrFeePrecisionFmt)
                    End If
                End If
                
                '应收金额=单价 * 付数 * 数次
                .应收金额 = .标准单价 * mobjBill.Pages(intPage).Details(lngRow).付数 * mobjBill.Pages(intPage).Details(lngRow).数次
                
                '附加手术费率用计算(所有收入项目)
                If mobjBill.Pages(intPage).Details(lngRow).附加标志 = 1 And mobjBill.Pages(intPage).Details(lngRow).收费类别 = "F" Then
                    .应收金额 = .应收金额 * IIf(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
                End If
                
                '加班费用率计算
                dbl加班加价率 = 0
                If mobjBill.加班标志 = 1 And mobjBill.Pages(intPage).Details(lngRow).Detail.加班加价 Then
                    dbl加班加价率 = IIf(IsNull(rsTmp!加班加价率), 0, rsTmp!加班加价率 / 100)             '传入根据费别计算实收金额函数
                    .应收金额 = .应收金额 + .应收金额 * dbl加班加价率
                End If
                
                .应收金额 = CCur(Format(.应收金额, gstrDec))
                
                dblAllTime = mobjBill.Pages(intPage).Details(lngRow).付数 * mobjBill.Pages(intPage).Details(lngRow).数次
                If InStr(",5,6,7,", mobjBill.Pages(intPage).Details(lngRow).收费类别) > 0 Then
                    If gbln药房单位 Then dblAllTime = dblAllTime * mobjBill.Pages(intPage).Details(lngRow).Detail.药房包装
                End If
                
                If mobjBill.Pages(intPage).Details(lngRow).Detail.屏蔽费别 Or bln从项汇总折扣 Then
                    .实收金额 = .应收金额
                    mobjBill.Pages(intPage).Details(lngRow).费别 = mobjBill.费别
                Else
                    If .应收金额 = 0 Then
                        .实收金额 = 0
                        mobjBill.Pages(intPage).Details(lngRow).费别 = mobjBill.费别
                    Else
                        '药品按成本价加收,传入数量
                        str费别 = IIf(glngSys Like "8??", mobjBill.费别, zlStr.TrimEx(mobjBill.费别 & "," & lbl动态费别.Tag, ","))
                        
                        .实收金额 = CCur(Format(ActualMoney(str费别, .收入项目ID, .应收金额, _
                            mobjBill.Pages(intPage).Details(lngRow).收费细目ID, mobjBill.Pages(intPage).Details(lngRow).执行部门ID, dblAllTime, dbl加班加价率), gstrDec))
                        mobjBill.Pages(intPage).Details(lngRow).费别 = str费别
                    End If
                End If
                
                '获取项目保险信息,门诊只有医保病人才算
                If mstrYBPati <> "" Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Pages(intPage).Details(lngRow).收费细目ID, .实收金额, True, mintInsure, _
                        mobjBill.Pages(intPage).Details(lngRow).摘要 & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Pages(intPage).Details(lngRow).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Pages(intPage).Details(lngRow).保险大类ID = Val(Split(strInfo, ";")(1))
                        .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        mobjBill.Pages(intPage).Details(lngRow).保险编码 = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Pages(intPage).Details(lngRow).摘要 = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Pages(intPage).Details(lngRow).Detail.类型 = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
                '实收金额存入Key中,以处理分币问题(即Key中存放原始实收金额,不变)
                mobjBill.Pages(intPage).Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
            End With
            rsTmp.MoveNext
        Next
    Else
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Pages(intPage).Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long, Optional intCurSubItem As Integer = 0, Optional intSubItemCount As Integer = 0)
'功能：刷新显示当前单据指定行或所有行的内容
'参数：lngRow=指定行,为0表示显示所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim curTotal As Currency, i As Long, str计算单位 As String
    Dim intCount As Integer
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Pages(mintPage).Details.Count
            Call ShowDetail(i)
        Next
    ElseIf mobjBill.Pages(mintPage).Details.Count > 0 Then
        Call ShowDetail(lngRow, intCurSubItem, intSubItemCount)
    End If
    Bill.Redraw = True
    
    '显示单据小计
    lblSub应收.Caption = "应收:" & Format(GetBillSum(True, CLng(mintPage)), gstrDec)
    lblSub实收.Caption = "实收:" & Format(GetBillSum(False, CLng(mintPage)), gstrDec)
    
    i = Get中药数量(str计算单位)
    If i = 0 Then
        lblAmount.Caption = ""
    Else
        lblAmount.Caption = "中药共:" & i & str计算单位
    End If
    
    
    If mbytInFun = 2 Or mbytInState = 2 Then
        curTotal = GetBillSum
        lblTotal.Caption = "合计:" & Format(curTotal, gstrDec)
        If mbytInFun = 2 And IsNumeric(cmdOK.Tag) Then
            '划价时显示不算当前单据费用,但划价报警要算
            sta.Panels(Pan.C4预交信息).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(Pan.C4预交信息).Text = sta.Panels(Pan.C4预交信息).Text & "/费用:" & Format(Val(cmdCancel.Tag) + IIf(mbytBilling = 0, curTotal, 0), "0.00")
            sta.Panels(Pan.C4预交信息).Text = sta.Panels(Pan.C4预交信息).Text & "/剩余:" & Format(Val(cmdPrint.Tag) - IIf(mbytBilling = 0, curTotal, 0), "0.00")
        End If
    End If
End Sub

Private Sub ShowDetail(lngRow As Long, Optional intCurSubItem As Integer = 0, Optional intSubItemCount As Integer = 0)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行
'         intCurSubItem-加载的当前套餐
'         intSubItemCount- 主要是针对套餐来说的,总共套餐项目数(是否为最后一笔)
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, j As Long, strTemp As String
    Dim cur金额 As Currency, dbl单价 As Double
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    If lngRow > mobjBill.Pages(mintPage).Details.Count Then Exit Sub
    
    '清除单据行
    For i = 1 To Bill.COLS - 1
        '输入时收费类别不清除
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Pages(mintPage).Details(lngRow).收费类别 <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Pages(mintPage).Details(lngRow).收费类别)
    End If
    
    '刷新单据行
    '问题:29201
    strTemp = ""
    If mobjBill.Pages(mintPage).Details(lngRow).从属父号 <> 0 Then
         strTemp = "┣"
         If intSubItemCount > 0 Then
            If intCurSubItem = intSubItemCount Then
                    strTemp = "┗"
            End If
         Else
                If lngRow < mobjBill.Pages(mintPage).Details.Count Then
                    If mobjBill.Pages(mintPage).Details(lngRow).从属父号 <> mobjBill.Pages(mintPage).Details(lngRow + 1).从属父号 Then
                         strTemp = "┗"
                    End If
                ElseIf lngRow = mobjBill.Pages(mintPage).Details.Count Then
                         strTemp = "┗"
                End If
          End If
        strTemp = "  " & strTemp & " "
    End If
    
    For i = 1 To Bill.COLS - 1
        Select Case Bill.TextMatrix(0, i)
            Case "类别"
                '浏览单据或从属项目只(能)显示名称
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.类别名称
            Case "从属父号"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).从属父号
            Case "项目"
                Bill.TextMatrix(lngRow, i) = strTemp & mobjBill.Pages(mintPage).Details(lngRow).Detail.名称
            Case "规格"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.规格
            Case "商品名"
                Bill.TextMatrix(lngRow, i) = strTemp & mobjBill.Pages(mintPage).Details(lngRow).Detail.商品名
            Case "单位"
                If InStr(",5,6,7,", mobjBill.Pages(mintPage).Details(lngRow).收费类别) > 0 And gbln药房单位 Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.药房单位
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.计算单位
                End If
            Case "付数"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Pages(mintPage).Details(lngRow).付数 = 0, 1, mobjBill.Pages(mintPage).Details(lngRow).付数)
            Case "数次"
                '数次在第一次显示时已默认设置为1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Pages(mintPage).Details(lngRow).数次, 5)
            Case "单价"
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                dbl单价 = 0
                If mobjBill.Pages(mintPage).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(mintPage).Details(lngRow).InComes.Count
                        dbl单价 = dbl单价 + mobjBill.Pages(mintPage).Details(lngRow).InComes(j).标准单价
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl单价, gstrFeePrecisionFmt)
            Case "应收金额"
                '应收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Pages(mintPage).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(mintPage).Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Pages(mintPage).Details(lngRow).InComes(j).应收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "实收金额"
                '实收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Pages(mintPage).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(mintPage).Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Pages(mintPage).Details(lngRow).InComes(j).实收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "执行科室", "发药药店"
                '可能无执行科室'200402
                If mobjBill.Pages(mintPage).Details(lngRow).执行部门ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Pages(mintPage).Details(lngRow).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                        Else
                            Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Pages(mintPage).Details(lngRow).执行部门ID, mrsUnit)
                        End If
                    Else
                        '浏览单据只(能)显示名称
                        Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Pages(mintPage).Details(lngRow).执行部门ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "标志"
                If mobjBill.Pages(mintPage).Details(lngRow).收费类别 = "F" And mobjBill.Pages(mintPage).Details(lngRow).附加标志 = 1 Then
                    Bill.TextMatrix(lngRow, i) = "√"
                End If
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.类型
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney(Optional ByVal intPage As Integer, Optional bln个帐 As Boolean = True)
'功能：刷新显示收入项目费用区，不支持预结算时的保险结算区，单据合计等
'参数：bln个帐=是否处理个人帐户显示
'      intPage=是否只重新计算指定单据(加快速度)，0-全部计算,-1,全不计算,x-计算指定单据
    Dim rsTmp As New ADODB.Recordset, arrDetail As Variant
    Dim cur冲款合计 As Currency, cur实收金额 As Currency, cur可用个帐 As Currency
    Dim cur个帐 As Currency, curTotal As Currency
    Dim cur全自付 As Currency, cur先自付 As Currency, cur进入统筹 As Currency
    Dim cur实收合计 As Currency, cur应收合计 As Currency, strTmp As String
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim blnExist As Boolean, blnDo As Boolean, strSQL As String

    '产生汇总费目,并统计保险相关金额
    '-------------------------------------------------------------------------
    
''    '优先使用预交款缴费
''    If mbytInFun = 0 And gblnPrePayPriority And Val(sta.Panels(Pan.C4预交信息).Tag) > 0 And txt预交冲款.Enabled Then
''        If Not Me.ActiveControl Is txt预交冲款 Then
''            curTotal = GetBillSum - GetMedicareSum
''            If curTotal > 0 Then
''                txt预交冲款.Text = Format(IIf(curTotal > Val(sta.Panels(Pan.C4预交信息).Tag), Val(sta.Panels(Pan.C4预交信息).Tag), curTotal), "0.00")
''            End If
''        End If
''    End If
    
    cur冲款合计 = Format(Val(txt预交冲款.Text), "0.00")
        
    Set mcolMoneys = New BillInComes
    
    For p = 1 To mobjBill.Pages.Count
        arrDetail = Array()
        cur应收合计 = 0: cur实收合计 = 0
        cur进入统筹 = 0: cur全自付 = 0: cur先自付 = 0
        If intPage = 0 Or p = intPage Then
            If mobjBill.Pages(p).NO = "" Then
                '该张单据内容是直接输入的
                For i = 1 To mobjBill.Pages(p).Details.Count
                    For j = 1 To mobjBill.Pages(p).Details(i).InComes.Count
                        With mobjBill.Pages(p).Details(i).InComes(j)
                            '合并到所有单据的项目汇总
                            blnExist = False
                            For k = 1 To mcolMoneys.Count
                                strTmp = IIf(gint分类合计 = 0, .收据费目, IIf(gint分类合计 = 2, "第" & p & "张", .收入项目)) '31479
                                If mcolMoneys(k).收据费目 = strTmp Then
                                    blnExist = True: Exit For
                                End If
                            Next
                            If blnExist Then
                                mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + .应收金额
                                mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + .实收金额
                            Else
                                strTmp = IIf(gint分类合计 = 0, .收据费目, IIf(gint分类合计 = 2, "第" & p & "张", .收入项目)) '31479
                                mcolMoneys.Add 0, strTmp, strTmp, 0, .应收金额, .实收金额
                            End If
                            
                            '合并到当前单据的项目汇总
                            blnExist = False
                            For k = 0 To UBound(arrDetail)
                                strTmp = IIf(gint分类合计 = 0, .收据费目, IIf(gint分类合计 = 2, "第" & p & "张", .收入项目)) '31479
                                If CStr(Split(arrDetail(k), ",")(0)) = strTmp Then
                                    blnExist = True: Exit For
                                End If
                            Next
                            If blnExist Then
                                arrDetail(k) = Split(arrDetail(k), ",")(0) & "," & _
                                    Val(Split(arrDetail(k), ",")(1)) + .应收金额 & "," & _
                                    Val(Split(arrDetail(k), ",")(2)) + .实收金额
                            Else
                                strTmp = IIf(gint分类合计 = 0, .收据费目, IIf(gint分类合计 = 2, "第" & p & "张", .收入项目)) '31479
                                ReDim Preserve arrDetail(UBound(arrDetail) + 1)
                                arrDetail(UBound(arrDetail)) = strTmp & "," & .应收金额 & "," & .实收金额
                            End If
                                 
                            '--
                            cur应收合计 = cur应收合计 + .应收金额
                            cur实收合计 = cur实收合计 + .实收金额
                            
                            '统计保险金额
                            cur实收金额 = .实收金额
                            If .统筹金额 = 0 Or Not mobjBill.Pages(p).Details(i).保险项目否 Then
                                '以原始金额为准,不管分币处理
                                cur全自付 = cur全自付 + cur实收金额
                            Else
                                cur进入统筹 = cur进入统筹 + .统筹金额
                                '以原始金额为准,不管分币处理
                                cur先自付 = cur先自付 + cur实收金额 - .统筹金额
                            End If
                        End With
                    Next
                Next
            Else
                '该单据是提取的划价单内容
                strSQL = "Select A.收据费目,B.名称 as 收入项目," & _
                    " A.应收金额,A.实收金额,A.统筹金额,A.保险项目否" & _
                    " From 门诊费用记录 A,收入项目 B" & _
                    " Where A.记录性质=1 And A.记录状态 IN(0,1,3) And A.收入项目ID=B.ID And A.NO=[1]" & _
                    " Order by 序号"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
                For i = 1 To rsTmp.RecordCount
                    '合并到所有单据的项目汇总
                    blnExist = False
                    For k = 1 To mcolMoneys.Count
                        strTmp = IIf(gint分类合计 = 0, rsTmp!收据费目, IIf(gint分类合计 = 2, "第" & p & "张", rsTmp!收入项目)) '31479
                        If mcolMoneys(k).收据费目 = strTmp Then
                            blnExist = True: Exit For
                        End If
                    Next
                    If blnExist Then
                        mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + Nvl(rsTmp!应收金额, 0)
                        mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + Nvl(rsTmp!实收金额, 0)
                    Else
                        strTmp = IIf(gint分类合计 = 0, rsTmp!收据费目, IIf(gint分类合计 = 2, "第" & p & "张", rsTmp!收入项目))  '31479
                        mcolMoneys.Add 0, strTmp, strTmp, 0, Nvl(rsTmp!应收金额, 0), Nvl(rsTmp!实收金额, 0)
                    End If
                    
                    '合并到当前单据的项目汇总
                    blnExist = False
                    For k = 0 To UBound(arrDetail)
                        strTmp = IIf(gint分类合计 = 0, rsTmp!收据费目, IIf(gint分类合计 = 2, "第" & p & "张", rsTmp!收入项目)) '31479
                        If CStr(Split(arrDetail(k), ",")(0)) = strTmp Then
                            blnExist = True: Exit For
                        End If
                    Next
                    If blnExist Then
                        arrDetail(k) = Split(arrDetail(k), ",")(0) & "," & _
                            Val(Split(arrDetail(k), ",")(1)) + Nvl(rsTmp!应收金额, 0) & "," & _
                            Val(Split(arrDetail(k), ",")(2)) + Nvl(rsTmp!实收金额, 0)
                    Else
                        strTmp = IIf(gint分类合计 = 0, rsTmp!收据费目, IIf(gint分类合计 = 2, "第" & p & "张", rsTmp!收入项目)) '31479
                        ReDim Preserve arrDetail(UBound(arrDetail) + 1)
                        arrDetail(UBound(arrDetail)) = strTmp & "," & Nvl(rsTmp!应收金额, 0) & "," & Nvl(rsTmp!实收金额, 0)
                    End If
                                        
                    '--
                    cur应收合计 = cur应收合计 + Nvl(rsTmp!应收金额, 0)
                    cur实收合计 = cur实收合计 + Nvl(rsTmp!实收金额, 0)
                    
                    '统计保险金额
                    cur实收金额 = Nvl(rsTmp!实收金额, 0)
                    If Nvl(rsTmp!统筹金额, 0) = 0 Or Nvl(rsTmp!保险项目否, 0) = 0 Then
                        '以原始金额为准,不管分币处理
                        cur全自付 = cur全自付 + cur实收金额
                    Else
                        cur进入统筹 = cur进入统筹 + Nvl(rsTmp!统筹金额, 0)
                        '以原始金额为准,不管分币处理
                        cur先自付 = cur先自付 + cur实收金额 - Nvl(rsTmp!统筹金额, 0)
                    End If
                    
                    rsTmp.MoveNext
                Next
            End If
        Else
            With mobjBill.Pages(p)
                cur应收合计 = mobjBill.Pages(p).应收金额
                cur实收合计 = mobjBill.Pages(p).实收金额
                cur进入统筹 = mobjBill.Pages(p).进入统筹
                cur全自付 = mobjBill.Pages(p).全自付
                cur先自付 = mobjBill.Pages(p).先自付
                
                '直接取Key值：项目名称,应收金额,实收金额;
                arrDetail = Split(.Key, ";")
                For i = 0 To UBound(arrDetail)
                    '合并到所有单据的项目汇总
                    blnExist = False
                    For k = 1 To mcolMoneys.Count
                        If mcolMoneys(k).收据费目 = CStr(Split(arrDetail(i), ",")(0)) Then
                            blnExist = True: Exit For
                        End If
                    Next
                    If blnExist Then
                        mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + Val(Split(arrDetail(i), ",")(1))
                        mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + Val(Split(arrDetail(i), ",")(2))
                    Else
                        strTmp = CStr(Split(arrDetail(i), ",")(0))
                        mcolMoneys.Add 0, strTmp, strTmp, 0, Val(Split(arrDetail(i), ",")(1)), Val(Split(arrDetail(i), ",")(2))
                    End If
                Next
            End With
        End If
        
        '更新当前单据个人帐户支付金额:不支持预结算时
        '医保病人且满足相应条件才处理,合计为负不能退到个人帐户
        If Not MCPAR.门诊预结算 Then
            If mstrYBPati <> "" And bln个帐 And mstr个人帐户 <> "" And mcur个帐余额 > -1 * mcur个帐透支 Then
                If cur实收合计 >= 0 Then
                    cur个帐 = cur进入统筹 + IIf(MCPAR.先自付, cur先自付, 0) + IIf(MCPAR.全自付, cur全自付, 0)
                    
                    '统计除开之前单据个帐支付后的个帐余额
                    cur可用个帐 = 0
                    For i = 1 To p - 1
                        cur可用个帐 = cur可用个帐 + GetMedicareSum(mstr个人帐户, i)
                    Next
                    cur可用个帐 = mcur个帐余额 - cur可用个帐
                                        
                    '计算个人帐户支付金额
                    If cur可用个帐 - cur个帐 >= -1 * mcur个帐透支 Then
                        Call SetBalanceVal(p, mstr个人帐户, Format(cur个帐, "0.00")) '在允许透支范围内足够(允许透支0为特例)
                    Else
                        If mcur个帐透支 = 0 And cur可用个帐 > 0 Then
                            Call SetBalanceVal(p, mstr个人帐户, Format(cur可用个帐, "0.00")) '不允许透支且有余额
                        Else
                            '超过允许透支范围或不允许透支时无余额
                            If mcur个帐透支 <> 0 Then
                                Call SetBalanceVal(p, mstr个人帐户, cur可用个帐 + mcur个帐透支) '在允许透支范围内支付
                            Else
                                Call SetBalanceVal(p, mstr个人帐户, 0)
                            End If
                        End If
                    End If
                Else
                    Call SetBalanceVal(p, mstr个人帐户, 0)
                End If
            End If
        End If
        
        '当前单据的相关汇总金额计算
        '----------------------------------------
        With mobjBill.Pages(p)
            .应收金额 = cur应收合计
            .实收金额 = cur实收合计
            
            If mbytInFun = 0 Then
                .进入统筹 = cur进入统筹
                .全自付 = cur全自付
                .先自付 = cur先自付
            
                '医保支付的所有金额,可能为预结算返回的,也可能是该过程计算的
                .保险金额 = GetMedicareSum(, p)
                                
                '计算当前单据应分解冲款的金额,为了计算应缴(多单据时先冲预交)
                If cur冲款合计 <> 0 Then
                    If cur冲款合计 <= Format(.实收金额 - .保险金额 - .消费卡刷卡额, "0.00") Then
                        .冲预交额 = cur冲款合计
                    Else
                        .冲预交额 = Format(.实收金额 - .保险金额 - .消费卡刷卡额, "0.00")
                    End If
                    cur冲款合计 = cur冲款合计 - .冲预交额
                Else
                    .冲预交额 = cur冲款合计
                End If
                
                
                
                '计算当前单据应缴金额，分币处理，误差等
                '恢复结算方式
                If cbo结算方式.ListIndex = -1 And .收费结算 <> "" Then
                    Call zlControl.CboSetIndex(cbo结算方式.hWnd, cbo.FindIndex(cbo结算方式, gstr结算方式, True))
                    If cbo结算方式.ListIndex = -1 And cbo结算方式.ListCount <> 0 Then
                        Call zlControl.CboSetIndex(cbo结算方式.hWnd, 0)
                    End If
                End If
                
                blnDo = False '现金方式时才处理分币,医保要求时才处理
                If cbo结算方式.ListIndex <> -1 And cbo结算方式.Visible Then
                    If cbo结算方式.ItemData(cbo结算方式.ListIndex) = 1 Then
                        blnDo = True
                    End If
                End If
                
                If blnDo And mstrYBPati <> "" Then
                    If MCPAR.门诊预结算 Then
                        If Not MCPAR.分币处理 Then
                            blnDo = False
                        End If
                    End If
                End If
                
                If blnDo Then
                    .应缴金额 = CentMoney(.实收金额 - .保险金额 - .冲预交额 - .消费卡刷卡额)
                Else
                     .应缴金额 = Format(.实收金额 - .保险金额 - .冲预交额 - .消费卡刷卡额, "0.00")
                End If
            
                .误差金额 = Format(.应缴金额 - (.实收金额 - .保险金额 - .冲预交额 - .消费卡刷卡额), gstrDec)
                
                .收费结算 = "" '两种方式相冲突
            End If
            
            'Key值的保存,用于快速计算
            strTmp = ""
            For i = 0 To UBound(arrDetail)
                strTmp = strTmp & ";" & Split(arrDetail(i), ",")(0) & "," & _
                    Split(arrDetail(i), ",")(1) & "," & Split(arrDetail(i), ",")(2)
            Next
            .Key = Mid(strTmp, 2)
        End With
    Next
    
    '刷新显示所有单据的个人帐户支付情况
    '-------------------------------------------------------------------------
    If mstrYBPati <> "" And bln个帐 And mstr个人帐户 <> "" And mcur个帐余额 > 0 Then
        If Not MCPAR.门诊预结算 Then
            With vsBalance
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, 0) = mstr个人帐户 Then Exit For
                Next
                If i <= .Rows - 1 Then
                    .TextMatrix(i, 1) = Format(GetMedicareSum(mstr个人帐户), "0.00")
                End If
            End With
        End If
    End If
    
    '刷新显示所有单据的分类金额(收费要按操作次数叠加)
    '-------------------------------------------------------------------------
    mshMoney.Redraw = False
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1 + mintMoneyRow
    End If
    If mshMoney.Rows < M_MONEY_ROWS Then mshMoney.Rows = M_MONEY_ROWS

    Call SetMoneyList
    
    cur应收合计 = 0: cur实收合计 = 0
    For i = mintMoneyRow + 1 To mcolMoneys.Count + mintMoneyRow
        mshMoney.TextMatrix(i, 0) = mintBillNO + 1
        mshMoney.TextMatrix(i, 1) = mcolMoneys(i - mintMoneyRow).收据费目
        mshMoney.TextMatrix(i, 2) = Format(mcolMoneys(i - mintMoneyRow).实收金额, gstrDec)
        cur应收合计 = cur应收合计 + mcolMoneys(i - mintMoneyRow).应收金额
        cur实收合计 = cur实收合计 + mcolMoneys(i - mintMoneyRow).实收金额
        
        '单据小计
        If i = mcolMoneys.Count + mintMoneyRow Then
            mshMoney.TextMatrix(i, 3) = Format(cur实收合计, gstrDec)
        Else
            mshMoney.TextMatrix(i, 3) = ""
        End If
    Next
    On Error Resume Next
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.TextMatrix(i, 0) = mintBillNO + 1 Then
            mshMoney.TopRow = i
        End If
    Next
    On Error GoTo 0
    mshMoney.Redraw = True
        
    '更新合计金额显示
    '----------------------------------------------------------
    txt应收.Text = Format(mcurBill应收 + cur应收合计, gstrDec)
    txt合计.Text = Format(mcurBill实收 + cur实收合计, gstrDec)
    txt应缴.Text = Format(GetMustPaySum + mcurBill应缴, "0.00")
   
    
    '划价时,txt累计用来表示应缴,即分币处理后的金额
    If mbytInFun = 1 Then txt累计.Text = Format(CentMoney(txt合计.Text), "0.00")
End Sub

Private Function GetInputDetail(ByVal lng项目id As Long, Optional ByVal lng批次 As Long) As Detail
'功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    '考虑卫生材料部份
    If mintInsure = 0 Then
        strSQL = _
            " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名," & _
            " A.规格,A.计算单位,A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要," & _
            " Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
            " Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
            " Decode(A.类别,'4',1,C." & gstr药房包装 & ") as 药房包装," & _
            " Decode(A.类别,'4',A.计算单位,C." & gstr药房单位 & ") as 药房单位,D.跟踪在用,A.录入限量,C.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,C.剂量系数" & _
            " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,诊疗项目目录 M1" & _
            " Where A.类别=B.编码 And A.ID=C.药品ID(+) And C.药名ID=M1.ID(+) And A.ID=D.材料ID(+)" & _
            " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            " And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
            " And A.ID=[1]"
    Else
        strSQL = _
            " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名," & _
            " A.规格,A.计算单位,A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要," & _
            " Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
            " Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
            " Decode(A.类别,'4',1,C." & gstr药房包装 & ") as 药房包装," & _
            " Decode(A.类别,'4',A.计算单位,C." & gstr药房单位 & ") as 药房单位,D.跟踪在用,A.录入限量,C.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,C.剂量系数" & _
            " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,保险支付项目 M,诊疗项目目录 M1" & _
            " Where A.类别=B.编码 And A.ID=C.药品ID(+) And C.药名ID=M1.ID(+)  And A.ID=D.材料ID(+)" & _
            " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            " And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
            " And A.ID=M.收费细目ID(+) And M.险类(+)=[2]" & vbNewLine & _
            " And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, mintInsure)
    With objDetail
        .ID = rsTmp!ID
        .药名ID = Nvl(rsTmp!药名ID, 0) '用于判断输入重复
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .编码 = rsTmp!编码
        .名称 = rsTmp!名称
        .商品名 = Nvl(rsTmp!商品名)
        .规格 = Nvl(rsTmp!规格)
        .计算单位 = Nvl(rsTmp!计算单位)
        .药房单位 = Nvl(rsTmp!药房单位)
        .药房包装 = Nvl(rsTmp!药房包装, 1)
        .分批 = Nvl(rsTmp!分批, 0) = 1 '是否药房分批
        .变价 = Nvl(rsTmp!是否变价, 0) = 1 '对药品表明是否时价
        .类型 = Nvl(rsTmp!费用类型)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
        .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
        .录入限量 = Val("" & rsTmp!录入限量)
        .中药形态 = Val(Nvl(rsTmp!中药形态))
        .诊疗名称 = Nvl(rsTmp!诊疗名称)
        .剂量单位 = Nvl(rsTmp!剂量单位)
        .剂量系数 = Val(Nvl(rsTmp!剂量系数))
        .批次 = lng批次
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
    intPay = GetOtherCTMGroups(lngRow)
    If Detail.类别 <> "7" Then intPay = 1
    
    If mobjBill.Pages(mintPage).Details.Count < lngRow Then
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
                    dblTime = IIf(Detail.从项数次 = 0, 1, Detail.从项数次)
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Pages(mintPage).Details(bytParent).数次
                End If
            Else
                
                If InStr(",5,6,7,", Detail.类别) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Pages(mintPage).Details.Add mobjBill.费别, Detail, .ID, CInt(lngRow), CInt(bytParent), .类别, .计算单位, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else '如果该行已经存在,则修改
        
        If InStr(",5,6,7,", Detail.类别) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Pages(mintPage).Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .费别 = mobjBill.费别
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

Private Function CheckHaveChildren(lngRow As Long) As Boolean
'功能：判断该行是否应该取从属项目
'说明：仅该行收费项目有从属项目及尚未取才取。
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select Count(从项ID) as NUM From 收费从属项目 Where 主项ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(mintPage).Details(lngRow).收费细目ID)
    If rsTmp.RecordCount <> 0 Then
        If IsNull(rsTmp!Num) Then
            CheckHaveChildren = False
        ElseIf rsTmp!Num = 0 Then
            CheckHaveChildren = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Pages(mintPage).Details.Count
                If mobjBill.Pages(mintPage).Details(i).从属父号 = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                CheckHaveChildren = True
            Else
                CheckHaveChildren = False
            End If
        End If
    Else
        CheckHaveChildren = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckMainItem(ByVal lngRow As Long, Optional ByVal intPage As Long) As Boolean
'功能：判断当前行的项目是否具有从属项目
    Dim i As Long
    
    If intPage = 0 Then intPage = mintPage
    
    If mobjBill.Pages(intPage).Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Pages(intPage).Details.Count
            If mobjBill.Pages(intPage).Details(i).从属父号 = lngRow Then
                CheckMainItem = True: Exit Function
            End If
        Next
    End If
End Function

Private Function GetSubDetails(ByVal lng项目id As Long) As Details
'功能：返回一个收费细目的从属项目集
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim objDetail As New Detail
        
    Set GetSubDetails = New Details
    
    '考虑卫材部份
    strSQL = _
    "Select A.ID,Decode(A.类别,'4',E.材料ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
    "       A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.计算单位,A.规格,A.屏蔽费别," & _
    "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.费用类型," & _
    "       Decode(A.类别,'4',1,D." & gstr药房包装 & ") as 药房包装," & _
    "       Decode(A.类别,'4',A.计算单位,D." & gstr药房单位 & ") as 药房单位," & _
    "       A.是否变价,A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,D.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,D.剂量系数" & _
    " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1,诊疗项目目录 M1" & _
    " Where A.类别=B.编码 And C.从项ID=A.ID And A.ID=D.药品ID(+) and D.药名ID=M1.ID(+) And A.ID=E.材料ID(+)" & _
    "   And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
    "   And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "   And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "   And C.主项ID=[1] Order by 编码"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .药名ID = Nvl(rsTmp!药名ID, 0)
            .编码 = rsTmp!编码
            .变价 = rsTmp!是否变价 = 1
            .规格 = Nvl(rsTmp!规格)
            .药房包装 = Nvl(rsTmp!药房包装, 1)
            .药房单位 = Nvl(rsTmp!药房单位)
            .计算单位 = Nvl(rsTmp!计算单位)
            .分批 = Nvl(rsTmp!分批, 0) = 1
            .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
            .类别 = rsTmp!类别
            .类别名称 = rsTmp!类别名称
            .名称 = rsTmp!名称
            .商品名 = Nvl(rsTmp!商品名)
            .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
            .执行科室 = Nvl(rsTmp!执行科室, 0) '缺省为无明确科室(用户选)
            .固有从属 = Nvl(rsTmp!固有从属, 0) '缺省为非固定,用户可以随意更改数次
            .从项数次 = Nvl(rsTmp!从项数次, 1)
            .类型 = Nvl(rsTmp!费用类型)
            .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
            .中药形态 = Val(Nvl(rsTmp!中药形态))
            .诊疗名称 = Nvl(rsTmp!诊疗名称)
            .剂量单位 = Nvl(rsTmp!剂量单位)
            .剂量系数 = Val(Nvl(rsTmp!剂量系数))
            GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .规格, .计算单位, .说明, .屏蔽费别, _
                .药房包装, .药房单位, .分批, .变价, .加班加价, .执行科室, .类型, .补充摘要, .固有从属, .从项数次, .跟踪在用, , , , , , , .中药形态, .商品名, .诊疗名称, .剂量单位, .剂量系数
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(ByVal lngRow As Long, Optional ByVal intPage As Integer)
'功能：删除指定收费项目行
'说明：这时不处理从属行的删除,但要对其它单据行从属关系作相应的调整
    Dim i As Long
    
    '如果未指定页,则用当前页
    If intPage = 0 Then intPage = mintPage
    
    For i = lngRow + 1 To mobjBill.Pages(intPage).Details.Count
        If mobjBill.Pages(intPage).Details(i).从属父号 <> 0 And _
            mobjBill.Pages(intPage).Details(i).从属父号 > lngRow Then
            mobjBill.Pages(intPage).Details(i).从属父号 = mobjBill.Pages(intPage).Details(i).从属父号 - 1
        End If
        mobjBill.Pages(intPage).Details(i).序号 = mobjBill.Pages(intPage).Details(i).序号 - 1 '序号与行号对应
    Next
    mobjBill.Pages(intPage).Details.Remove lngRow
    
    '删除当前显示单据页的指定行
    If tbsBill.SelectedItem.Index = intPage Then
        If lngRow = 1 And mobjBill.Pages(intPage).Details.Count = 0 And Bill.Rows = 2 Then
            For i = 1 To Bill.COLS - 1
                Bill.TextMatrix(lngRow, i) = ""
                Bill.RowData(lngRow) = 0
            Next
            Call SetBillRowForeColor(lngRow, Bill.ForeColor)
        Else
            Bill.RemoveMSFItem lngRow
        End If
    End If
End Sub

Private Sub NewYBBill()
'功能：用于医保连续收费时调用,连续收费模式下不能使用多单据收费
    Dim i As Integer
    
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '多单据收费:恢复缺省单据页卡
    mintPage = 1
    If fraBill.Visible Then
        cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
        cmdDelBill.Enabled = False
        tbsBill.TabStop = False
        For i = tbsBill.Tabs.Count To 1 Step -1
            tbsBill.Tabs(i).Tag = ""
            If i <> 1 Then tbsBill.Tabs.Remove i
        Next
    End If
    
    mlngPreRow = 0
    mblnHotKey = False
    mstrCardNO = ""
    If mbln补费 And mstr最后转科时间 <> "" Then
        txtDate.Text = Format(CDate(mstr最后转科时间) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
    Else
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    Call InitBalanceGrid
    txt预交冲款.Text = "0.00"
    sta.Panels(Pan.C4预交信息).Tag = ""
    Original.冲预交款 = 0
    Original.实收合计 = 0
    Original.应缴金额 = 0
    ''txt本次应缴.Visible = False: lbl应缴.Caption = "应缴"
      
    cboNO.Text = ""
    
    '刷新票据号,只有自用的时，在打印后已刷新
    If mbytInFun = 0 Then Call RefreshFact
        
    With mobjBill
        .发生时间 = CDate(txtDate.Text)
        .费别 = IIf(cbo费别.ListIndex = -1, "", Mid(cbo费别.Text, InStr(cbo费别.Text, "-") + 1))
        .加班标志 = chk加班.Value
        If cbo开单科室.ListIndex = -1 Then
            .Pages(mintPage).开单部门ID = 0
        Else
            .Pages(mintPage).开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        .Pages(mintPage).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
        .门诊标志 = gint病人来源
        .划价人 = UserInfo.姓名
        .操作员编号 = UserInfo.编号
        .操作员姓名 = UserInfo.姓名
    End With
End Sub

Private Function NewBill(Optional blnFact As Boolean = True, Optional bln费别 As Boolean = True, _
    Optional ByVal blnClearPatiInfor As Boolean) As Boolean
'功能：初始化一张新的单据(程序对象)
'参数：blnFact=是否取票号
'      bln费别=是否重新初始化费别
    Dim i As Long
    Dim Curdate As Date '服务器当前时间
    
    If blnClearPatiInfor Then Set mrsInfo = New ADODB.Recordset
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '多单据收费:恢复缺省单据页卡
    mintPage = 1
    
    Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
    If cmdIDCard.Visible Then cmdIDCard.Enabled = True
    If cmdRegist.Visible Then cmdRegist.Enabled = True
    
    cmdAddBill.Enabled = InStr(1, mstrPrivs, "普通病人多单据收费") > 0
    cmdDelBill.Enabled = False
    tbsBill.TabStop = False
    If fraBill.Visible Then
        For i = tbsBill.Tabs.Count To 1 Step -1
            tbsBill.Tabs(i).Tag = ""
            If i <> 1 Then tbsBill.Tabs.Remove i
        Next
    End If
    mdbl缴款 = 0: mdbl找补 = 0
    
    mstrYBBill = "": mstrYBPati = "": mintInsure = 0
    mcur个帐余额 = 0: mcur个帐透支 = 0
    mblnYB结算作废 = False  '不同的病人可能险类不同而医保作废支持不同,所以要清除
    mbytBillSource = 1
    If txtMCInvoice.Visible Then
        txtMCInvoice.Visible = False
        txtMCInvoice.Text = ""
    End If
    
    mblnSaveAsPrice = False
    mblnHotKey = False
    mbln报合计 = False
    Original.实收合计 = 0: Original.冲预交款 = 0: mlngPreRow = 0
    Original.应缴金额 = 0
'''    txt本次应缴.Visible = False: lbl应缴.Caption = "应缴"
    
    mstrCardNO = ""
    txtPatient.ForeColor = Me.ForeColor
    mnuFileSavePrice.Checked = False
    chk急诊.Value = 0: chk急诊.Visible = False
    
    If mstr付款方式 <> "" Then
        cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, mstr付款方式, True)
        If cbo医疗付款.ListIndex = -1 And cbo医疗付款.ListCount > 0 Then cbo医疗付款.ListIndex = 0
    ElseIf cbo医疗付款.ListCount > 0 Then
        cbo医疗付款.ListIndex = 0
    End If
    cbo医疗付款.Locked = False Or mbytInFun = 2
    If mbytInFun = 2 And mbytInState = 0 Then cboBaby.ListIndex = 0
    sta.Panels(Pan.C3个人帐户).Tag = "": sta.Panels(Pan.C3个人帐户).Text = "": sta.Panels(Pan.C3个人帐户).Visible = False
            
    Call InitBalanceGrid
    Call SetButton(2) '确定,取消
    Call ShowPrePayInfo(False) '预交信息初始
    Call ShowPayInfo(True) '联合医保
    
    If mbytInFun = 0 And blnClearPatiInfor Then
        Call SetPatientEnableModi(True)
        txtRePrint.Enabled = True: txtModi.Enabled = True: txtIn.Enabled = True
        cboNO.Enabled = True: chkCancel.Enabled = True: cmdDelete.Enabled = True
    End If
        
    If gbyt科室医生 = 0 And mstrPrePati <> txtPatient.Text Then
        cbo开单人.ListIndex = -1: cbo开单科室.ListIndex = -1: lblDuty.Caption = ""
    End If
    
    If mbln补费 And mstr最后转科时间 <> "" Then
        Curdate = CDate(mstr最后转科时间) - 1 / 24 / 60
    Else
        Curdate = zlDatabase.Currentdate
    End If
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbytInState = 0 Then
        cboNO.Text = ""
        mstrWarn = ""
        cmdOK.Tag = "": cmdCancel.Tag = "": cmdPrint.Tag = ""
        If blnFact Then
            txtInvoice.Text = ""
            Call ReInitPatiInvoice(blnFact)
        End If
        
'        If mbytInFun = 2 Then Call ClearPatientInfo(True)   '仅保留合计信息
        chk加班.Value = IIf(OverTime(Curdate), 1, 0)
        
        '结算方式
        If mbytInFun = 0 Then
            i = cbo.FindIndex(cbo结算方式, gstr结算方式, True)
            If i = -1 And cbo结算方式.ListCount > 0 Then i = 0
            Call zlControl.CboSetIndex(cbo结算方式.hWnd, i)
        End If
        
        '费别处理：收费或划价
        If Not (glngSys Like "8??" Or mbytInFun = 2) Then
            cbo费别.Locked = False: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
            cbo费别.Visible = True
            lbl动态费别.BorderStyle = 0
            lbl动态费别.Left = cbo费别.Left + cbo费别.Width + 60
            
            If bln费别 Then Call LoadAndSeek费别(True)
        End If
        
        '其它
        With mobjBill
            .发生时间 = CDate(txtDate.Text)
            .费别 = IIf(cbo费别.ListIndex = -1, "", Mid(cbo费别.Text, InStr(cbo费别.Text, "-") + 1))
            .加班标志 = chk加班.Value
            If cbo开单科室.ListIndex = -1 Then
                .Pages(mintPage).开单部门ID = 0
            Else
                .Pages(mintPage).开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            End If
            .Pages(mintPage).开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
            .门诊标志 = gint病人来源
            .划价人 = UserInfo.姓名
            .操作员编号 = UserInfo.编号
            .操作员姓名 = UserInfo.姓名
        End With
        
    End If
    
    NewBill = True
End Function

Private Sub ClearMoney()
'功能：清除费用显示区
    Dim i As Integer, j As Integer
    mshMoney.Redraw = False
    mintMoneyRow = 0
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.COLS - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mshMoney.Rows = M_MONEY_ROWS
    mshMoney.Redraw = True
End Sub

Private Function GetDrugWindow(ByVal lng药房ID As Long, ByVal str类别 As String, ByVal intPage As Integer) As String
'功能：获取缺省的发药窗口,如果参数指定了缺省,则以指定为准,否则,如果是划价单,则以第一药品行的窗口为准,否则以已输入相同药品的窗口为准
'参数：intPage=搜录到的单据编号
'说明：主要用于多单据收费时，不同类别的药品可能动态分配到同一药房，这样他们的窗口也应相同，但强行指定的除外
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Dim strPayWin As String
    
    Err = 0: On Error GoTo errH:
    strPayWin = ""
    For p = 1 To intPage
         If mobjBill.Pages(p).NO <> "" Then
             If tbsBill.Tabs(p).Tag <> "" Then
                 '问题:47489
                 '取划价单的第一药品行的药房进行比较
                 ''执行部门ID|发药窗口;...
                 varData = Split(tbsBill.Tabs(p).Tag, ";")
                 For i = 0 To UBound(varData)
                     varTemp = Split(varData(i) & "|", "|")
                     If varTemp(0) = lng药房ID Then
                          strPayWin = varTemp(1)
                          GoTo GoFind:
                     End If
                 Next
             End If
         Else
             For i = 1 To mobjBill.Pages(p).Details.Count
                 If mobjBill.Pages(p).Details(i).执行部门ID = lng药房ID _
                     And InStr(",5,6,7,", mobjBill.Pages(p).Details(i).收费类别) > 0 _
                     And mobjBill.Pages(p).Details(i).发药窗口 <> "" Then
                     strPayWin = mobjBill.Pages(p).Details(i).发药窗口
                     GoTo GoFind:
                 End If
             Next
         End If
     Next
GoFind:
    If strPayWin <> "" Then GetDrugWindow = strPayWin: Exit Function
    
    strPayWin = GetDefaultWindow(str类别, lng药房ID)
    '检查是否上班
    strSQL = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药房ID, GetDrugWindow)
    If rsTmp.EOF Then strPayWin = ""
    GetDrugWindow = strPayWin
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Function ReChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新收取费用
    '返回:重新收取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 18:18:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPageInfor As Collection, strInvoice As String
    Dim p As Integer, strBalanceIDs As String, lng结帐ID As Long
    Dim rsErrBlance As ADODB.Recordset
    Dim strSaveNos As String, strSaveSuessNos As String, blnAffair As Boolean
    Dim blnNotCallInsure As Boolean '不调用医保接口
    Dim blnPatiIndentify As Boolean '是否已经医保身份验证
    If mbytInFun <> 0 Then ReChargeFee = True: Exit Function
   Dim strSQL As String, strDate As String
   Dim blnCommit As Boolean
   
    '并发检查
    If zlIsCheckExistErrBill(mlng结算序号) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng结算序号) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    strInvoice = Trim(txtInvoice.Text)
    If Not CheckBillNOAndBookeFee Then Exit Function

    Set cllPageInfor = New Collection
    On Error GoTo ErrRoll:
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    mobjBill.登记时间 = CDate(strDate)
    gcnOracle.BeginTrans: blnCommit = True
    For p = 1 To mobjBill.Pages.Count
        lng结帐ID = mobjBill.Pages(p).结帐ID
        strBalanceIDs = strBalanceIDs & "," & lng结帐ID
        cllPageInfor.Add Array(lng结帐ID, mobjBill.Pages(p).NO), "K" & p
        strSaveNos = strSaveNos & "," & "'" & mobjBill.Pages(p).NO & "'"
    
        '44507
        strSQL = "Zl_门诊收费异常_Update('" & mobjBill.Pages(p).NO & "',to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        strSQL = "Zl_票据起始号_Update('" & mobjBill.Pages(p).NO & "','" & strInvoice & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans: blnCommit = False
    On Error GoTo errHandle
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
    If strSaveNos <> "" Then strSaveNos = Mid(strSaveNos, 2)
    '医保补调
    If zlInsure补调交易(strSaveNos) = False Then Exit Function
    If mbytInFun <> 0 Then Exit Function
    
    '弹出付款方式
    If frmChargePayMentWin.zlChargeWin(Me, EM_重新收费, mlngModul, mstrPrivs, mlngShareUseID, mstrUseType, mlng结算序号, strBalanceIDs, strSaveNos, mobjBill.病人ID, mintInsure, mobjBill.姓名, mobjBill.性别, mobjBill.年龄, mobjBill.费别, mdbl缴款, mdbl找补) = False Then
        If Not mblnErrBill Then
            Unload Me
        End If
        Exit Function
    End If
    
    '显示Led相关信息
    'LED显示:(合计,)发药窗口
    If gblnLED And CCur(txt合计.Text) <> 0 And (mstr西窗 <> "" Or mstr中窗 <> "" Or mstr成窗 <> "") Then
        zl9LedVoice.DisplayBank "费用合计:" & txt合计.Text, _
            "取药窗口:" & IIf(mstr西窗 <> "", " " & mstr西窗, "") & _
            IIf(mstr成窗 <> "", " " & mstr成窗, "") & IIf(mstr中窗 <> "", " " & mstr中窗, "")
    End If
    
    Call CheckBillNOAndBookeFee(True)
     '打印票据
     Call PrintBill(strSaveNos, "")
     If Not mblnErrBill Then
        gblnOK = True: Unload Me
     End If
    ReChargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrRoll:
    If blnCommit Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Function DelErrBillFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:异常单据作废
    '返回:异常单据作废成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 18:18:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPageInfor As Collection, strInvoice As String
    Dim p As Integer, strBalanceIDs As String, lng结帐ID As Long
    Dim rsErrBlance As ADODB.Recordset
    Dim strSaveNos As String, strSaveSuessNos As String, blnAffair As Boolean
    Dim blnNotCallInsure As Boolean '不调用医保接口
    Dim blnPatiIndentify As Boolean '是否已经医保身份验证
    If mbytInFun <> 0 Then DelErrBillFee = True: Exit Function
   
    strInvoice = Trim(txtInvoice.Text)
    Set cllPageInfor = New Collection
    On Error GoTo errHandle
    '并发检查
    If zlIsCheckExistErrBill(mlng结算序号) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng结算序号) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    For p = 1 To mobjBill.Pages.Count
        lng结帐ID = mobjBill.Pages(p).结帐ID
        strBalanceIDs = strBalanceIDs & "," & lng结帐ID
        cllPageInfor.Add Array(lng结帐ID, mobjBill.Pages(p).NO), "K" & p
        strSaveNos = strSaveNos & "," & "'" & mobjBill.Pages(p).NO & "'"
    Next
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
    If strSaveNos <> "" Then strSaveNos = Mid(strSaveNos, 2)
    mbln连续输入 = False
    '医保补调
    '弹出付款方式
    If frmChargePayMentWin.zlChargeWin(Me, EM_异常作废, mlngModul, mstrPrivs, mlngShareUseID, mstrUseType, _
        mlng结算序号, strBalanceIDs, strSaveNos, mobjBill.病人ID, mintInsure, _
        mobjBill.姓名, mobjBill.性别, mobjBill.年龄, mobjBill.费别, mdbl缴款, mdbl找补, _
        , , , , , , mbln退费异常) = False Then
        mlng结算序号 = 0
        Unload Me
        Exit Function
    End If
    Dim lng结算序号 As Long
    lng结算序号 = Get作废结算序号(strSaveNos)
    Call WriteMzInforToCard(mobjBill.病人ID, lng结算序号, True)
    mlng结算序号 = 0:
    gblnOK = True: Unload Me
    DelErrBillFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get作废结算序号(ByVal strNos As String) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取作废的结算序号
    '返回:返回作废的结算序号
    '编制:刘兴洪
    '日期:2012-12-14 18:52:31
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select a.结算序号" & _
            " From 病人预交记录 A, 门诊费用记录 B" & _
            " Where a.结帐id = b.结帐id And b.No In (Select Column_Value From Table(f_Str2list([1]))) " & _
            "       And a.记录状态 = 2 And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
    If rsTemp.EOF Then Exit Function
    Get作废结算序号 = Val(Nvl(rsTemp!结算序号))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlInsure补调交易(ByVal strSaveNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保补调
    '返回:补调成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, str保险结算 As String, strAdvance As String, blnMedicareCheck As Boolean
    Dim strTmp As String, i As Long, blnNotCallInsure As Boolean, blnPatiIndentify As Boolean
    Dim strSQL As String, strsuccesNOs As String, strNotSucces As String
    Dim lng病人ID  As Long, lng结帐ID As Long
    Dim dbl个人帐户 As Double, dbl医保基金 As Double
   
    lng病人ID = mobjBill.病人ID
    strSQL = "" & _
    "   Select A.NO,A.结帐ID,decode(A.记录性质,1,'预存款',11,'预存款',A.结算方式) as 结算方式,A.冲预交, " & _
    "           Decode(B.名称,NULL,0,1) as 医保, " & _
    "           Decode(C.结算方式,NULL,0,1) as 一卡通, " & _
    "           Decode(nvl(A.卡类别ID,0),0,0,1) as 医疗卡, " & _
    "           Decode(nvl(A.结算卡序号,0),0,0,1) as 消费卡, " & _
    "           nvl(A.校对标志,0) as 校对标志  " & _
    "   From 病人预交记录 A, " & _
    "           (Select 名称 From 结算方式 where 性质 in (3,4)) B," & _
    "           (Select 结算方式 From 一卡通目录 Where 启用=1 ) C " & _
    "   where A.结算序号=[1]  and a.记录性质=3 and A.结算方式 is not null  " & _
    "               And A.结算方式=B.名称(+)  And A.结算方式=C.结算方式(+)"
    
    strSQL = "" & _
    "   Select NO,结算方式,nvl(sum(冲预交),0) as 结算金额, " & _
    "               nvl(Max(医保),0) as 医保, nvl(Max(一卡通),0) as 一卡通, " & _
    "               nvl(max(医疗卡),0) as 医疗卡, " & _
    "               nvl(Max(消费卡),0) as 消费卡,nvl(Max(校对标志),0) as 校对标志 " & _
    "   From (" & strSQL & ")" & _
    "   Group by  NO,结帐ID,结算方式"
    Err = 0: On Error GoTo errH:
    '异常单据的结算方式(不含预交款)
    Set mrsErrBlance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结算序号)
    
    Err = 0: On Error GoTo errHandle:
    mrsErrBlance.Filter = "医保=0 "
    '如果存在其他的结算方式,则肯定不会有医保,直接返回
    If mrsErrBlance.RecordCount > 0 Then zlInsure补调交易 = True: Exit Function
    '检查是否存在医保
    mrsErrBlance.Filter = 0
    mrsErrBlance.Filter = "医保=1 and 校对标志=2  "
    mintInsure = 0
    blnNotCallInsure = False
    If Not mrsErrBlance.EOF Then
        mintInsure = ChargeExistInsure(Nvl(mrsErrBlance!NO), , lng结帐ID)
        Call initInsurePara(lng病人ID)
        '以下情况不调用医保接口:
        '1.多单据一次结算(因为是一个事务,所以肯定医保成功,单据也保存成功)
        '2.多单据调用一次接口(因为是一个事务,所以肯定医保成功,单据也保存成功)
        '3.所有单据,都存在医保,都不存在较对的情况
        '4.存在预结算时,每张单据都应该有医保,只是调用不成功的情况
        If MCPAR.多单据一次结算 Then zlInsure补调交易 = True: Exit Function
        If MCPAR.多单据调一次交易 Then zlInsure补调交易 = True:  Exit Function
    End If
    '补调医保交易
    blnPatiIndentify = False: strsuccesNOs = ""
    For p = 1 To mobjBill.Pages.Count
         
        blnNotCallInsure = True
        '肯定在单据中存在数据
        mrsErrBlance.Filter = "  NO='" & mobjBill.Pages(p).NO & "' and 医保=1 and  校对标志=1  "
        If Not mrsErrBlance.EOF Then blnNotCallInsure = False
        If Not MCPAR.门诊预结算 And blnNotCallInsure Then    '不存在虚拟结算,那么可能存在调用医保的情况
             mrsErrBlance.Filter = "  NO='" & mobjBill.Pages(p).NO & "' and 医保=1"
             If mrsErrBlance.EOF Then blnNotCallInsure = True
        End If
        '需要补调交易,需要验证身份
        If blnPatiIndentify = False And Not blnNotCallInsure Then
            '进行医保刷卡验证
            If MsgBox("注意:" & vbCrLf & _
                "    单据号为" & mobjBill.Pages(p).NO & " 的是医保结算单据, " & _
                "     目前只预结算了,但还未进行医保正式结算,是否重新医保结算?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                If VerifyPati(mobjBill.病人ID) = False Then Exit Function
                blnPatiIndentify = True
            End If
        End If
        '不调交易
        If Not blnNotCallInsure Then
            mrsErrBlance.Filter = "  NO='" & mobjBill.Pages(p).NO & "' and 医保=1 and  校对标志=1  "
            With mrsErrBlance
                str保险结算 = ""    '结算方式|金额||
                dbl个人帐户 = 0: dbl医保基金 = 0
                Do While Not .EOF
                    str保险结算 = str保险结算 & "||" & Nvl(!结算方式) & "|" & Val(Nvl(!结算金额))
                    If Nvl(!结算方式) = mstr个人帐户 Then
                        dbl个人帐户 = dbl个人帐户 + Val(Nvl(!结算金额))
                    End If
                    If Nvl(!结算方式) = "医保基金 " Then
                        dbl医保基金 = dbl医保基金 + Val(Nvl(!结算金额))
                    End If
                    .MoveNext
                Loop
                If str保险结算 <> "" Then str保险结算 = Mid(str保险结算, 3)
            End With
            gcnOracle.BeginTrans: blnTrans = True: blnTransMedicare = False
            strAdvance = mobjBill.Pages.Count & "|" & p
            If Not gclsInsure.ClinicSwap(mobjBill.Pages(p).结帐ID, CCur(dbl个人帐户), _
                CCur(dbl医保基金), mobjBill.Pages(p).全自付, mobjBill.Pages(p).先自付, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans:
                If strsuccesNOs <> "" Then
                    strsuccesNOs = Mid(strsuccesNOs, 2)
                    strSaveNos = "'" & Replace(strsuccesNOs, ",", "','") & "'"
                Else
                    strSaveNos = ""
                End If
                strNotSucces = ""
                For i = p To mobjBill.Pages.Count
                    strNotSucces = strNotSucces & "," & mobjBill.Pages(i).NO
                Next
                If strNotSucces <> "" Then strNotSucces = Mid(strNotSucces, 2)
                If ModifyNotInsureNOs(strNotSucces, strsuccesNOs) = False Then
                    Exit Function
                End If
                zlInsure补调交易 = True
                Exit Function
            Else
                blnTransMedicare = True
            End If
            blnMedicareCheck = zlInsureCheck(str保险结算, strAdvance)
            '问题:
            ' Zl_病人门诊收费_医保更新
            gstrSQL = "Zl_病人门诊收费_医保更新("
            '  结帐id_In   门诊费用记录.结帐id%Type,
            gstrSQL = gstrSQL & "" & "NULL" & ","
            '  结算序号_In 病人预交记录.结算序号%Type,
            gstrSQL = gstrSQL & "" & mlng结算序号 & ","
            '  保险结算_In Varchar2
            gstrSQL = gstrSQL & IIf(blnMedicareCheck, "'" & strAdvance & "'", "NULL") & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            strsuccesNOs = strsuccesNOs & "," & mobjBill.Pages(p).NO
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
            gcnOracle.CommitTrans: blnTrans = False: blnTransMedicare = False
        End If
    Next
    zlInsure补调交易 = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mintInsure)
    End If
    Call SaveErrLog
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
End Function


Private Function VerifyPati(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保身份证验正
    '编制:刘兴洪
    '日期:2011-08-27 12:58:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID1 As Long
    lng病人ID1 = lng病人ID '避免Identify接口中修改该变量后返回新值
    '返回：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,24就诊类型(1=急诊门诊),25开单科室名称
    mstrYBPati = gclsInsure.Identify(id门诊收费, lng病人ID, mintInsure)
    If mstrYBPati = "" Then Exit Function
    If UBound(Split(mstrYBPati, ";")) < 8 Then GoTo ErrCancelYb:
    If Val(Split(mstrYBPati, ";")(8)) = 0 Then GoTo ErrCancelYb:
    
    '获取病人信息
    lng病人ID1 = Val(CLng(Split(mstrYBPati, ";")(8)))
    If lng病人ID <> lng病人ID1 And lng病人ID1 <> 0 And lng病人ID <> 0 Then
        MsgBox "医保验证的病人与之前提取的病人不是同一个病人!", vbInformation, gstrSysName
        GoTo ErrCancelYb: Exit Function
    End If
    '初始医保参数
    Call initInsurePara(lng病人ID)
    VerifyPati = True
    Exit Function
ErrCancelYb:
    Call YBIdentifyCancel: mintInsure = 0: mstrYBPati = ""
End Function
Private Function SaveBill(ByRef strSaveNos As String, _
    Optional ByRef strModiNos As String, _
    Optional ByRef blnSaveBill As Boolean, _
    Optional ByRef blnNotPayWin As Boolean, _
    Optional bytReturnMode As ExitMode = EM_收费完成, _
    Optional bln连续 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前输入的单据(适用于收费、划价、门诊记帐)
    '出参:strSaveNos-返回已成功保存的单据号，格式为"'AAA','BBB',..."
    '       cur已缴合计-配合strSaveNOs，返回已保存成功的单据实际已缴的现金
    '       strModiNOs -修改的是多单据收费中的一张时，返回该多张单据的所有NO，格式如"'AAA','BBB',..."
    '       blnSaveBill-是否单据已经保存成功
    '       blnNotPayWin-不弹出收费界面
    '       bytReturnMode As Byte ' '0-正常收费完成,1-暂停收费;2-本次作废收费;3-继续输入
    '返回:收费成功或单据保存存功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-26 17:28:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '     *** 医保收费时,先临时保存为划价单,在结算前再转为收费单,以避免更新药品库存时因等待同一事务的医保结算操作而锁表 ***
    Dim lng打印ID As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim int序号 As Integer, int价格父号 As Integer, int行号 As Integer
    Dim lng结帐ID As Long, int药品行次 As Integer, str医疗付款 As String
    Dim dbl数次 As Double, dbl单价 As Double, cur缴款 As Currency
    Dim strDeptIDs As String, strTmp As String, strDelBill As String, strBillNO As String
    Dim str收费结算 As String, str保险结算 As String, str收费结算校对 As String
    Dim arrSQL As Variant, arrPut As Variant, arrOTMSQL As Variant
    Dim bln直接收费 As Boolean, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim CurOneCard As Currency, dblOneCardBalance As Double
    Dim strCardNo  As String, intCardType As Integer, strTransFlow As String
    Dim str中药形态 As String
    Dim strStuffDept As String          '自动发料的部门
    Dim strAdvance As String            '医保结算返回的信息:"结算方式|结算金额||....."
    Dim blnPriceSaved As Boolean        '医保收费时是否已存为划价单,用于在转为收费单及医保结算事务失败回退后删除划价单
    Dim blnMedicareCheck As Boolean     '是否执行医保结算校对
    Dim strBalanceIDs As String         '所有单据的结帐ID，用来传给医保接口
    Dim strInvoice As String            '当前单据使用的票据号，用于医保一张单据只打一张票的情况
    Dim cllRqure As Collection
    Dim rsSqure As ADODB.Recordset
    Dim str结帐IDs As String
    Dim bln应付款 As Boolean
    Dim dbl应缴额 As Double, lng结帐序号 As Long
    Dim cllPutout As Collection '自动发料
    Dim cllPro As Collection, cllDelete As Collection, cllPageInfor As Collection
    Dim cur已缴合计 As Currency
    
    Set mCllWindows = New Collection
    
    strSaveNos = "": cur已缴合计 = 0: strModiNos = ""
    Err = 0: On Error GoTo Errhand:
    If cbo医疗付款.ListIndex <> -1 Then
        str医疗付款 = Mid(cbo医疗付款.Text, 1, InStr(1, cbo医疗付款, "-") - 1)
    End If
    strInvoice = Trim(txtInvoice.Text)
    
    arrOTMSQL = Array()
    '修改功能时,是否修改医嘱附费
    If mstrInNO <> "" Then
        Call BillisAdviceMoney(mstrInNO, IIf(mbytInFun = 2, 2, 1), lng医嘱ID, lng发送号)
    End If
    If mlng关联医嘱 <> 0 And lng医嘱ID = 0 Then lng医嘱ID = mlng关联医嘱
    
    blnSaveBill = False
    dbl应缴额 = 0: lng结帐序号 = 0
    Set cllPutout = New Collection: Set cllPro = New Collection
    Set cllPageInfor = New Collection
    '对每张单据独立执行保存
    For p = 1 To mobjBill.Pages.Count
        int序号 = 0: int行号 = 0: blnPriceSaved = False
        int药品行次 = 0: strDeptIDs = "": strStuffDept = ""
        '当前收费单据的各类结算
        If mbytInFun = 0 And Not mblnSaveAsPrice Then
            str保险结算 = GetMedicareStr(p)
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            If lng结帐序号 = 0 Then lng结帐序号 = lng结帐ID
            'If mlng结算序号 <> 0 Then lng结帐序号 = mlng结算序号
            str结帐IDs = str结帐IDs & "," & lng结帐ID
        End If
        '产生每张收费单据的单据号
        If mobjBill.Pages(p).NO = "" Then
            '为保存失败后仍能识别,不改对象NO
            Select Case mbytInFun
                Case 0, 1 '收费单、划价单
                    strBillNO = zlDatabase.GetNextNo(13)
                Case 2  '门诊记帐单
                    strBillNO = zlDatabase.GetNextNo(14)
            End Select
            bln直接收费 = True
        Else
            bln直接收费 = False
            strBillNO = mobjBill.Pages(p).NO
        End If
        '主要为消息发送用,为每页保存的单据号
        mobjBill.Pages(p).收费单号 = strBillNO
        If p = 1 Then
            mobjBill.NO = strBillNO
            gstrModiNO = strBillNO
        End If
        
        arrSQL = Array() '多单据时,逐张单据提交
        If Not bln直接收费 Then
            '1.收费新单据功能时,提取的划价单收费
            '虽然Zl_病人划价收费_Insert没有更新医保信息,但在根据病人提取的划价单时执行了zl_门诊划价记录_Update,已更新
            '---------------------------------------------------------------
            If Not mblnSaveAsPrice Then
                'Zl_病人划价收费_Insert
                strSQL = "Zl_病人划价收费_Insert("
                '  No_In         门诊费用记录.NO%Type,
                strSQL = strSQL & "'" & strBillNO & "',"
                '  病人id_In     门诊费用记录.病人id%Type,
                strSQL = strSQL & "" & ZVal(mobjBill.病人ID) & ","
                '  病人来源_In   Number,
                strSQL = strSQL & "" & gint病人来源 & ","
                '  付款方式_In   门诊费用记录.付款方式%Type,
                strSQL = strSQL & "'" & str医疗付款 & "',"
                '  姓名_In       门诊费用记录.姓名%Type,
                strSQL = strSQL & "'" & mobjBill.姓名 & "',"
                '  性别_In       门诊费用记录.性别%Type,
                strSQL = strSQL & "'" & mobjBill.性别 & "',"
                '  年龄_In       门诊费用记录.年龄%Type,
                strSQL = strSQL & "'" & mobjBill.年龄 & "',"
                '  病人科室id_In 门诊费用记录.病人科室id%Type,
                strSQL = strSQL & "" & IIf(mobjBill.Pages(p).医嘱序号 > 0, "Null", ZVal(mobjBill.科室ID, , mobjBill.Pages(p).开单部门ID)) & ","
                '  开单部门id_In 门诊费用记录.开单部门id%Type,
                strSQL = strSQL & "" & ZVal(mobjBill.Pages(p).开单部门ID) & ","
                '  开单人_In     门诊费用记录.开单人%Type,
                strSQL = strSQL & "'" & mobjBill.Pages(p).开单人 & "',"
                '  保险结算_In   Varchar2,
                If mstrYBPati <> "" And str保险结算 <> "" Then
                    strSQL = strSQL & "'" & str保险结算 & "',"
                Else
                    strSQL = strSQL & "NULL,"
                End If
                '  结帐id_In     门诊费用记录.结帐id%Type,
                strSQL = strSQL & "" & lng结帐ID & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                strSQL = strSQL & "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                strSQL = strSQL & "'" & UserInfo.编号 & "',"
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
                strSQL = strSQL & "'" & tbsBill.Tabs(p).Tag & "',"
                '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
                strSQL = strSQL & "" & chk急诊.Value & ","
                '  登记时间_In   门诊费用记录.登记时间%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  结算序号_In   病人预交记录.结算序号%Type := Null
                strSQL = strSQL & "" & lng结帐序号 & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "0;" & strSQL
                            
                '获取自动发药的多个药房
                If gbln收费后自动发药 Then
                    strDeptIDs = strDeptIDs & "," & Get发药部门IDs(strBillNO)
                End If
                '针对每张单据收集卫料发料部门,以便自动发料,是否是跟踪在用材料在SQL中判断
                If gbln门诊自动发料 Then
                    strStuffDept = strStuffDept & "," & Get发药部门IDs(strBillNO, "'4'")
                End If
                
                '通过划价单收费的方式收取了挂号发卡的费用,则不用删除该费用
                If strBillNO = mstrCardNO Then mstrCardNO = ""
            '提取划价单收费,但仍保存为划价单,或联合医保的保存
            ElseIf mstrYBPati <> "" And mobjBill.病人ID <> 0 Then
                '更新划价单病人信息
                gstrSQL = "zl_门诊划价记录_Update(" & mintInsure & "," & mobjBill.病人ID & ",'" & strBillNO & "',1)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
            End If
            
        ElseIf bln直接收费 Then
            '2.直接输入的单据内容,包括新增和修改,可能是收费(或收费界面保存为划价单),记帐,划价
            '---------------------------------------------------------------
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                If mobjBillDetail.数次 <> 0 Then
                    For Each mobjBillIncome In mobjBillDetail.InComes
                        int序号 = int序号 + 1 '当前记录序号
                        '1.单据主体---------------------------------------------------------------
                        With mobjBill                              '医保收费时,先临时保存为划价单,在结算前再转为收费单
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = "zl_病人门诊收费_INSERT('" & strBillNO & "'," & int序号 & "," & ZVal(.病人ID) & "," & _
                                    IIf(gint病人来源 = 2, 2, 1) & "," & ZVal(.标识号) & ",'" & str医疗付款 & "'," & _
                                    "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                                    .加班标志 & "," & ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                    ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
                            ElseIf mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                                gstrSQL = "zl_门诊划价记录_INSERT('" & strBillNO & "'," & int序号 & "," & ZVal(.病人ID) & "," & _
                                    ZVal(.主页ID) & "," & ZVal(.标识号) & ",'" & str医疗付款 & "'," & _
                                    "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                                    .加班标志 & "," & ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                    ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
                            ElseIf mbytInFun = 2 Then
                                gstrSQL = "zl_门诊记帐记录_INSERT('" & strBillNO & "'," & int序号 & "," & _
                                    .病人ID & "," & .标识号 & ",'" & .姓名 & "','" & .性别 & "','" & .年龄 & "'," & _
                                    "'" & .费别 & "'," & .加班标志 & "," & .婴儿费 & "," & _
                                      ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                    ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
                            End If
                        End With
        
                        '2.收费细目部份---------------------------------------------------------------
                        With mobjBillDetail
                            If .序号 <> int行号 Then     '处理从属父号
                                int行号 = .序号
                                int价格父号 = int序号
                                '重新处理从属父号
                                If mobjBill.Pages(p).Details(.序号).从属父号 = 0 Then
                                    For i = .序号 + 1 To mobjBill.Pages(p).Details.Count
                                        If mobjBill.Pages(p).Details(i).从属父号 = .序号 Then
                                            '当父项目有多个收入项目(多个序号)时,取第一个序号
                                            mobjBill.Pages(p).Details(i).从属父号 = int序号
                                        End If
                                    Next
                                End If
                            End If
        
                            '收费、划价的药品行,处理发药窗口
                            If (mbytInFun = 0 Or mbytInFun = 1) And InStr(",5,6,7,", .收费类别) > 0 Then
                                If Set发药窗口(p, mobjBillDetail) = False Then Exit Function
                            End If
                            
                            '医保直接收费时,因为先暂存为划价单,收费时需要取发药窗口
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati <> "" Then tbsBill.Tabs(p).Tag = .发药窗口
                            
                            dbl数次 = .数次
                            If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                                dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                            End If
                            
                            gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = gstrSQL & IIf(.保险项目否, 1, 0) & "," & ZVal(.保险大类ID) & ",'" & .发药窗口 & "'," & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & IIf(.工本费, 8, .附加标志) & ","
                            ElseIf mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                                gstrSQL = gstrSQL & "'" & .发药窗口 & "'," & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & ","
                            ElseIf mbytInFun = 2 Then
                                gstrSQL = gstrSQL & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & ","
                            End If
                            gstrSQL = gstrSQL & IIf(.执行部门ID = 0, "NULL", .执行部门ID) & ","
                           
                        End With
        
                        '3.收入项目部份---------------------------------------------------------------
                        With mobjBillIncome
                            dbl单价 = .标准单价
                            If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 And gbln药房单位 Then
                                dbl单价 = Format(.标准单价 / mobjBillDetail.Detail.药房包装, gstrFeePrecisionFmt)
                            End If
                            gstrSQL = gstrSQL & IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & .收入项目ID & "," & _
                                    "'" & .收据费目 & "'," & dbl单价 & "," & .应收金额 & "," & .实收金额 & ","
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = gstrSQL & "NULL,"
                            End If
                        End With
        
                        '4.其它部分
                        '---------------------------------------------------------------
                        gstrSQL = gstrSQL & _
                                "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                        If mobjBillDetail.收费类别 = "7" Then
                            str中药形态 = "'" & mobjBillDetail.Detail.中药形态 & "'"
                        Else
                            str中药形态 = "NULL"
                        End If
                        '中药形态_In       住院费用记录.结论%Type := Null
                        
                        If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                            '非医保收费,并且不是划价
                            gstrSQL = gstrSQL & lng结帐ID & "," & lng结帐序号 & ","
                            '卫材类别ID
                            gstrSQL = gstrSQL & "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                "'" & mobjBillDetail.摘要 & "'," & chk急诊.Value & ",'|" & mobjBill.Pages(mintPage).煎法 & "'" & _
                                "," & str中药形态 & ")"
                                '只在第一张单据的第一条记录时传入
                        ElseIf mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                            '门诊划价,收费功能划价,医保收费
                            gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                                "'" & mobjBillDetail.摘要 & "'," & ZVal(lng医嘱ID) & ",NULL,NULL,'|" & mobjBill.Pages(mintPage).煎法 & _
                                "',NULL,NULL," & gint病人来源 & ",'" & mobjBillDetail.保险编码 & "'," & _
                                "'" & mobjBillDetail.Detail.类型 & "'," & IIf(mobjBillDetail.保险项目否, 1, 0) & "," & ZVal(mobjBillDetail.保险大类ID) & "," & _
                                str中药形态 & ",0," & IIf(mobjBillDetail.Detail.批次 = -1 Or mobjBillDetail.Detail.批次 = 0, "Null", mobjBillDetail.Detail.批次) & "," & _
                                "NULL," & ZVal(mobjBill.病区ID) & ")"
                        ElseIf mbytInFun = 2 Then
                            '门诊记帐
                            gstrSQL = gstrSQL & IIf(mbytBilling = 1, 1, 0) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                "NULL,'" & mobjBillDetail.摘要 & "'," & ZVal(lng医嘱ID) & ",Null,Null,'|" & mobjBill.Pages(mintPage).煎法 & "'," & _
                                "NULL,NULL,1," & str中药形态 & ",0," & IIf(mobjBillDetail.Detail.批次 = -1 Or mobjBillDetail.Detail.批次 = 0, "Null", mobjBillDetail.Detail.批次) & "," & _
                                ZVal(mobjBill.主页ID) & "," & ZVal(mobjBill.病区ID) & ")"
                        End If
        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
                    Next    '每一条收入项目
                    
                    '对每一行收费记录收集药品执行部门,门诊记帐划价单的审核操作,在Oracle过程中处理:zl_门诊记帐记录_Verify
                    '----------------------------------------------------------------------------------------------------------------
                    '自动发药,仅收费时且不是分离发药时                    '
                    With mobjBillDetail
                        If gbln收费后自动发药 And mbytInFun = 0 And Not mblnSaveAsPrice Then
                            If .执行部门ID <> 0 And InStr("5,6,7", .收费类别) > 0 Then
                                If InStr(strDeptIDs & ",", "," & .执行部门ID & ",") = 0 Then
                                    strDeptIDs = strDeptIDs & "," & .执行部门ID
                                End If
                            End If
                        End If
                        '自动发料,收费且不是保存为划价单或者门诊记帐,分离发药参数不影响卫材
                        If gbln门诊自动发料 And ((mbytInFun = 0 And Not mblnSaveAsPrice) Or (mbytInFun = 2 And mbytBilling = 0)) Then
                                If .执行部门ID <> 0 And .收费类别 = "4" And .Detail.跟踪在用 Then
                                    If InStr(strStuffDept & ",", "," & .执行部门ID & ",") = 0 Then
                                        strStuffDept = strStuffDept & "," & .执行部门ID
                                    End If
                                End If
                        End If
                    End With
                End If
            Next            '每一行收费项目
            
            '保存前一张单据的药房ID,以便多张单据时确定发药窗口
            If mobjBill.Pages.Count > 1 Then Call SaveDrugID(p)
                
        
            '修改后退除原单据(修改多收费单中的一张时需要后退费以统一重打)
            '--------------------------------------------------------------------------------------------------------
            If mstrInNO <> "" Then
                strDelBill = ""
                If mbytInFun = 0 And Not mblnSaveAsPrice Then
                    '修改医保收费单,必然为单张内全退,因为修改调用时已判断了如果不是全退,则不允许修改
                    strDelBill = "zl_门诊收费记录_DELETE('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                        "NULL,NULL,'" & zlStr.NeedName(cbo结算方式.Text) & "',0,To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                        
                    '如果是多单据收费中的一张则将新单关联到与原单据的打印ID上,以便一起重打
                    strTmp = GetMultiNOs(mstrInNO, lng打印ID)
                    If UBound(Split(strTmp, ",")) = 0 Then
                        lng打印ID = 0: strModiNos = ""
                    ElseIf lng打印ID <> 0 Then
                        strModiNos = strTmp
                    End If
                ElseIf mbytInFun = 1 Or (mbytInFun = 0 And mblnSaveAsPrice) Then
                    strDelBill = "zl_门诊划价记录_DELETE('" & mstrInNO & "')"
                ElseIf mbytInFun = 2 Then
                    strDelBill = "zl_门诊记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                End If
            
                '如果是修改医嘱的附费,则将新的NO放在附费中
                If lng医嘱ID <> 0 And lng发送号 <> 0 Then
                    gstrSQL = "ZL_病人医嘱附费_Insert(" & lng医嘱ID & "," & lng发送号 & "," & IIf(mbytInFun = 2, 2, 1) & ",'" & strBillNO & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
                End If
            End If
        End If
        '收费后自动发药,记帐不自动发药,收费且不是保存为划价单,或者门诊记帐
        '-----------------------------------------------------------------------
        If strDeptIDs <> "" Then
            arrPut = Array()
            strDeptIDs = Mid(strDeptIDs, 2)
            For i = 0 To UBound(Split(strDeptIDs, ","))
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "ZL_药品收发记录_处方发药(" & Val(Split(strDeptIDs, ",")(i)) & ",8,'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & mobjBill.Pages(p).开单人 & "')"
            Next
        End If
        '收费后自动发料,在收费(直接收费,划价单导入收费),门诊记帐时执行
        If strStuffDept <> "" Then
            If strDeptIDs = "" Then arrPut = Array()
            strStuffDept = Mid(strStuffDept, 2)
            For i = 0 To UBound(Split(strStuffDept, ","))          '24-收费处方发料；25-记帐单处方发料
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "zl_材料收发记录_处方发料(" & Split(strStuffDept, ",")(i) & "," & IIf(mbytInFun = 0, 24, 25) & ",'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
            Next
        End If
        
    
        '执行相关SQL语句及提交医保结算,多张单据时,每张单据在独立事务中提交
        '--------------------------------------------------------------------------------------------------------------------------------
        If UBound(arrSQL) >= 0 Then
            '对SQL序列按收费细目ID排序
            For i = 0 To UBound(arrSQL) - 1
                For j = i + 1 To UBound(arrSQL)
                    If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                        strTmp = CStr(arrSQL(j)): arrSQL(j) = arrSQL(i): arrSQL(i) = strTmp
                    End If
                Next
            Next
            
            '医保直接收费时,先保存为划价单,再转为收费单
            '-------------------------------------------------------------------
            If bln直接收费 And mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati <> "" Then
                '1.先保存划价单,先提交库存更新以便不锁表
                On Error GoTo errH
                For i = 0 To UBound(arrSQL)
                    zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                Next
                blnPriceSaved = True
                '更新划价单的保险信息(保险项目否,医保大类ID,统筹金额)
                gstrSQL = "zl_门诊划价记录_Update(" & mintInsure & "," & mobjBill.病人ID & ",'" & strBillNO & "',0)"
                zlAddArray cllPro, gstrSQL
                    
                '划价单转为收费单
                 'Zl_病人划价收费_Insert
                gstrSQL = "Zl_病人划价收费_Insert("
                '  No_In         门诊费用记录.NO%Type,
                gstrSQL = gstrSQL & "'" & strBillNO & "',"
                '  病人id_In     门诊费用记录.病人id%Type,
                gstrSQL = gstrSQL & "" & mobjBill.病人ID & ","
                '  病人来源_In   Number,
                gstrSQL = gstrSQL & "" & gint病人来源 & ","
                '  付款方式_In   门诊费用记录.付款方式%Type,
                gstrSQL = gstrSQL & "'" & str医疗付款 & "',"
                '  姓名_In       门诊费用记录.姓名%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.姓名 & "',"
                '  性别_In       门诊费用记录.性别%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.性别 & "',"
                '  年龄_In       门诊费用记录.年龄%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.年龄 & "',"
                '  病人科室id_In 门诊费用记录.病人科室id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.科室ID, , mobjBill.Pages(p).开单部门ID) & ","
                '  开单部门id_In 门诊费用记录.开单部门id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).开单部门ID) & ","
                '  开单人_In     门诊费用记录.开单人%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).开单人 & "',"
                '  保险结算_In   Varchar2,
                gstrSQL = gstrSQL & "" & IIf(str保险结算 <> "", "'" & str保险结算 & "'", "NULL") & ","
                '  结帐id_In     门诊费用记录.结帐id%Type,
                gstrSQL = gstrSQL & "" & lng结帐ID & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.编号 & "',"
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
                gstrSQL = gstrSQL & "" & chk急诊.Value & ","
                '  登记时间_In   门诊费用记录.登记时间%Type := Null,
                gstrSQL = gstrSQL & "" & "NULL" & ","
                '  结算序号_In   病人预交记录.结算序号%Type := Null
                gstrSQL = gstrSQL & "" & lng结帐序号 & ")"
            End If
            '医保多单据一次结算时，所有单据做为一个事务提交
            If MCPAR.多单据一次结算 And mstrYBPati <> "" And strDelBill = "" And Not mblnSaveAsPrice Then
                '1.划价单转收费
                zlAddArray cllPro, gstrSQL
                '2.误差费用
                If mobjBill.Pages(p).误差金额 <> 0 Then '44657
                    gstrSQL = "zl_门诊收费误差_Insert('" & strBillNO & "'," & mobjBill.Pages(p).误差金额 & ",0,1)"
                    zlAddArray cllPro, gstrSQL
                End If
                '3.收费后自动发药,自动发料
                If strDeptIDs <> "" Or strStuffDept <> "" Then
                    For i = 0 To UBound(arrPut)
                        zlAddArray cllPutout, arrPut(i)
                    Next
                End If
                'strBalanceIDs = IIf(strBalanceIDs = "", "", strBalanceIDs & ",") & lng结帐ID
            Else
                On Error GoTo errH
                    '修改功能相关处理
                    '先删除原单据,因为库存和预交款需要先还原
                    If strDelBill <> "" Then zlAddArray cllPro, strDelBill
                    
                    'a.非医保直接收费
                    If Not (bln直接收费 And mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati <> "") Then
                        '删除就诊卡划价单:多张单据时只删除一次(因为通过就诊卡号读病人时,就诊卡划价单已生成收费细目行,所以要删除)
                        If mbytInFun = 0 And mstrCardNO <> "" And strSaveNos = "" Then
                            gstrSQL = "zl_门诊划价记录_Delete('" & mstrCardNO & "')"
                            zlAddArray cllPro, gstrSQL
                        End If
                        '执行主体的SQL语句
                        For i = 0 To UBound(arrSQL)
                            'Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
                            zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                        Next
                        'b.医保直接收费
                    Else
                       ' Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        zlAddArray cllPro, gstrSQL
                    End If
                    '收费完成后的处理
                    '-----------------------------------------------------
                    If mbytInFun = 0 And Not mblnSaveAsPrice Then
                        '先填写开始票据号以便医保调用时上传,多张分别打印时,填写相同的,打印调用时将重写,取消打印或打印失败将清除
                        '修改时,只填写新单据的开始票据号,因为医保只对新单据上传
                        If strInvoice <> "" And mblnPrint Then
                            gstrSQL = "Zl_票据起始号_Update('" & strBillNO & "','" & strInvoice & "',1)"
                            zlAddArray cllPro, gstrSQL
                        End If
                    
                        '每张单据处理误差,该结帐ID与刚生成的收费记录相同
                        If mobjBill.Pages(p).误差金额 <> 0 Then '44657
                            gstrSQL = "zl_门诊收费误差_Insert('" & strBillNO & "'," & mobjBill.Pages(p).误差金额 & ",0,1)"
                            zlAddArray cllPro, gstrSQL
                        End If
                    End If
                    
                    '收费后自动发药,自动发料
                    If strDeptIDs <> "" Or strStuffDept <> "" Then
                        For i = 0 To UBound(arrPut)
                            'Call zlDatabase.ExecuteProcedure(CStr(arrPut(i)), Me.Caption)
                            zlAddArray cllPutout, CStr(arrPut(i))
                        Next
                    End If
                    
                    '修改功能相关处理
                    If strDelBill <> "" Then
                        '收费：新单据关联到原单据的打印ID上,以便一起重打,此时并未产生票据
                        If lng打印ID <> 0 And mblnPrint Then
                            gstrSQL = "zl_门诊收费票据_Insert('" & strBillNO & "','',Null,'',Null," & lng打印ID & ",0)"
                            zlAddArray cllPro, gstrSQL
                            'Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        End If
                    End If
                    '需要处理银行卡刷卡或第三方刷卡部分(暂不支持,改费方式,在程允入口判断)
                On Error GoTo 0
            End If
            strBalanceIDs = IIf(strBalanceIDs = "", "", strBalanceIDs & ",") & lng结帐ID
            cllPageInfor.Add Array(lng结帐ID, strBillNO), "K" & p
            
            '提交成功后再累加
            If mbytInFun = 0 And Not mblnSaveAsPrice Then
                cur已缴合计 = cur已缴合计 + mobjBill.Pages(p).应缴金额
            End If
            strSaveNos = strSaveNos & ",'" & strBillNO & "'"
            If Left(strSaveNos, 1) = "," Then strSaveNos = Mid(strSaveNos, 2)
            '加入单据历史记录(所有类型单据)
            cboNO.AddItem strBillNO, 0
            For i = cboNO.ListCount - 1 To 10 Step -1
                cboNO.RemoveItem i '只显示10个
            Next
        End If
    Next  '下一张单据
    On Error GoTo errH:
    '先保存单据
    Dim blnAffair As Boolean, strSaveCuessNos As String
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
   
    If zlInsureClinicSwap(cllPageInfor, lng结帐序号, strInvoice, strDelBill <> "", _
        strBalanceIDs, strSaveNos, strSaveCuessNos, blnAffair) = False Then
        If Not blnAffair Then gcnOracle.RollbackTrans
        If strSaveCuessNos <> "" Then blnSaveBill = True:
        If strSaveNos <> "" Then bytReturnMode = 2
        Exit Function
    End If
    If blnAffair = False Then gcnOracle.CommitTrans
    
    blnSaveBill = True
    
    '记账单自动发料处理,112720
    If mbytInFun = 2 And mbytBilling = 0 And cllPutout.Count > 0 Then
        Err = 0: On Error GoTo ErrPutOut:
        zlExecuteProcedureArrAy cllPutout, Me.Caption
        SaveBill = True: Exit Function
    End If
    
    If mblnSaveAsPrice Then SaveBill = True: Exit Function
    blnTrans = False
    If mbytInFun <> 0 Then SaveBill = True: Exit Function
    '弹出付款方式
    If blnNotPayWin Then SaveBill = True: Exit Function
    Dim dbl本次应缴 As Double
    mlng结算序号 = lng结帐序号
    Dim frmNew As frmChargePayMentWin
    Set frmNew = New frmChargePayMentWin
    If frmNew.zlChargeWin(Me, 0, mlngModul, mstrPrivs, mlngShareUseID, mstrUseType, lng结帐序号, strBalanceIDs, strSaveNos, mobjBill.病人ID, mintInsure, mobjBill.姓名, mobjBill.性别, mobjBill.年龄, mobjBill.费别, mdbl缴款, mdbl找补, bytReturnMode, CDbl(mcurBill应缴), bln连续, mlngPreBrushCard, dbl本次应缴, mstrBalance) = False Then
        If Not frmNew Is Nothing Then Unload frmNew
        Exit Function
    End If
    If Not frmNew Is Nothing Then Unload frmNew
    
    mblnNotClearLedDisplay = True
    mbln连续输入 = False
    If mstrYBPati <> "" And bln连续 Or mstrYBPati = "" And bln连续 Then
        mbln连续输入 = True
        For i = 1 To mobjBill.Pages.Count
            mobjBill.Pages(i).应缴金额 = 0
        Next
        If grsTotal.RecordCount <> 0 Then grsTotal.MoveFirst
        dbl本次应缴 = 0
        Do While Not grsTotal.EOF
            '性质:0-缴款;1-找补,2-冲预交;其他(mod 10:0-普通结算;1-医保结算;2-三方接品;3-一卡通)
            If Val(Nvl(grsTotal!性质)) <> 11 Then
                '非医保的累计
                dbl本次应缴 = dbl本次应缴 + Val(Nvl(grsTotal!结算金额))
            End If
            grsTotal.MoveNext
        Loop
        mobjBill.Pages(1).应缴金额 = dbl本次应缴
    End If
    
    '自动发药和发料处理
    Err = 0: On Error GoTo ErrPutOut:
    zlExecuteProcedureArrAy cllPutout, Me.Caption
    SaveBill = True
    Exit Function
errH:
    If Err.Description Like "*当前计算单价不一致*" Then
        If blnTrans Then gcnOracle.RollbackTrans
        If MsgBox("某些分批药品价格已发生变化，要自动重算价格吗？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
            Exit Function
        End If
     Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        If blnTransMedicare = False Then    '如果医保成功了，不删除划价单，费用失败可以重收
            Call DelMedicareTempNO(blnPriceSaved, strBillNO)
        End If
        Call SaveErrLog
    End If
    
    Exit Function
ErrPutOut:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function zlInsureClinicSwap(ByVal cllPageInfor As Collection, _
    ByVal lng结算序号 As Long, _
    ByVal strInvoice As String, _
    ByVal blnModifyBill As Boolean, _
    ByVal strBalanceIDs As String, _
    ByRef strSaveNos As String, _
    ByRef strSaveSucessNos As String, _
    Optional ByRef blnAffair As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保调用
    '入参:blnModifyBill-是否修改单据
    '       strBalanceIDs:本次结帐的ID,分别用逗号分离
    '       strSaveNos-保存的单据号
    '出参:strSaveNos-返回已经结算成功的单据号
    '       blnAffair-是否事务处理
    '       strSaveSucessNos-保存成功的票据(对划价有效)
    '返回:医保调用成功或非医保,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, str保险结算 As String, strAdvance As String, blnMedicareCheck As Boolean
    Dim strTmp As String, i As Long
    Dim strSuccesInsureNOs As String   '医保成功的单据
    Dim strNotSuccesInsureNOs As String   '医保成功的单据
    On Error GoTo errHandle
    '非医保，返回true,否则返加
    blnAffair = False
    If mstrYBPati = "" Or mbytInFun <> 0 Then zlInsureClinicSwap = True: Exit Function
    
    '1. 保存为划价单
    If mblnSaveAsPrice Then
        For p = 1 To mobjBill.Pages.Count
            strBillNO = cllPageInfor("K" & p)(1)
            If blnAffair Then gcnOracle.BeginTrans
            '保存为划价单
            '如果是联合医保,收费确定时实际却保存为划价单:传划价单明细,不在Oracle事务中执行
            If mbytInFun = 0 And Not mnuFileSavePrice.Checked Then
                If Not gclsInsure.TranChargeDetail(1, strBillNO, 1, 0, "", , mintInsure) Then
                    '删除划价单(继续处理)
                    Call DelMedicareTempNO(True, strBillNO)
                Else
                    strSaveSucessNos = strSaveSucessNos & "," & strBillNO
                End If
            End If
            gcnOracle.CommitTrans
            blnAffair = True
        Next
        zlInsureClinicSwap = True
        Exit Function
    End If
    
    '2.医保多单据一次结算时，所有单据做为一个事务提交
    If MCPAR.多单据一次结算 And Not blnModifyBill And Not mblnSaveAsPrice Then
        If blnAffair Then gcnOracle.BeginTrans
        
        If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
            '38821
            '票据数据生成(因为不调HIS的打印，医保接口打印，所以先填票据数据)
            gstrSQL = "zl_门诊收费票据_Insert('" & Replace(strSaveNos, "'", "") & "','" & strInvoice & "'," & ZVal(mlng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                      "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),0,1)"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        End If
        
        '不严格控制票据时保存当前票号
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "当前收费票据号", strInvoice, glngSys, 1121, InStr(1, mstrPrivs, ";参数设置;") > 0
        End If
        strAdvance = strBalanceIDs
        If Not gclsInsure.ClinicSwap(Val(Split(strBalanceIDs, ",")(0)), 0, 0, 0, 0, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            For i = 0 To UBound(Split(strSaveNos, ","))
                strBillNO = Replace(Split(strSaveNos, ",")(i), "'", "")
                For p = 1 To mobjBill.Pages.Count
                    If mobjBill.Pages(p).NO = strBillNO Then strBillNO = "": Exit For
                Next
                If strBillNO <> "" Then Call DelMedicareTempNO(True, strBillNO)
            Next
            blnAffair = True
            strSaveNos = ""
            Exit Function
        Else
            blnTransMedicare = True
        End If
        If strAdvance = strBalanceIDs Then strBalanceIDs = ""
         
        '根据返回的结算方式进行分摊
        ' Zl_病人门诊收费_医保更新
        gstrSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        gstrSQL = gstrSQL & "" & "NULL" & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        gstrSQL = gstrSQL & "" & lng结算序号 & ","
        '  保险结算_In Varchar2
        '问题:47409
        gstrSQL = gstrSQL & "" & IIf(strAdvance = "", "NULL", "'" & strAdvance & "'") & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        gcnOracle.CommitTrans: blnTrans = False
        '问题:47123
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
        blnAffair = True
        zlInsureClinicSwap = True: Exit Function
   End If
   
    '---------------------------------------------------
    '3.修改时,先退除原收费单据(改费方式)
    blnTransMedicare = False
    strAdvance = ""
    If Not mblnSaveAsPrice And blnModifyBill Then
        strAdvance = mobjBill.Pages.Count & "|" & p
        If Not gclsInsure.ClinicDelSwap(Original.结帐ID, False, mintInsure, strAdvance) Then
            blnAffair = True
            gcnOracle.RollbackTrans:   Exit Function
        End If
        blnTransMedicare = True
        gcnOracle.CommitTrans: blnTrans = False: blnAffair = True
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
    End If
    
    '4.多单据调用一次交易
    If MCPAR.多单据调一次交易 Then
        strAdvance = strBalanceIDs & "|" & lng结算序号
        If Not gclsInsure.ClinicSwap(lng结算序号, 0, 0, 0, 0, mintInsure, strAdvance) Then
            Exit Function
        End If
        If strAdvance = strBalanceIDs & "|" & lng结算序号 Then strAdvance = ""
        '根据返回的结算方式进行分摊
        ' Zl_病人门诊收费_医保更新
        gstrSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        gstrSQL = gstrSQL & "" & "NULL" & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        gstrSQL = gstrSQL & "" & lng结算序号 & ","
        '  保险结算_In Varchar2
        gstrSQL = gstrSQL & IIf(strAdvance <> "", "'" & strAdvance & "'", "NULL") & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        gcnOracle.CommitTrans: blnAffair = True
        '47123
         Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
         zlInsureClinicSwap = True: Exit Function
    End If
    
    '5.分单据调医保交易
    '       因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
    strSuccesInsureNOs = ""
    For p = 1 To mobjBill.Pages.Count
           If blnAffair Then gcnOracle.BeginTrans
           blnTrans = True: blnTransMedicare = False
            If (GetMedicareSum(, p) <> 0 Or MCPAR.门诊必须传递明细) Then
                strAdvance = mobjBill.Pages.Count & "|" & p
                If Not gclsInsure.ClinicSwap(Val(cllPageInfor("K" & p)(0)), GetMedicareSum(mstr个人帐户, p), _
                    GetMedicareSum("医保基金", p), mobjBill.Pages(p).全自付, mobjBill.Pages(p).先自付, mintInsure, strAdvance) Then
                    blnAffair = True: gcnOracle.RollbackTrans:
                    strSaveNos = ""
                    If strSuccesInsureNOs <> "" Then
                        strSuccesInsureNOs = Mid(strSuccesInsureNOs, 2)
                        strSaveNos = "'" & Replace(strSuccesInsureNOs, ",", "','") & "'"
                        strNotSuccesInsureNOs = ""
                        For i = p To mobjBill.Pages.Count
                            strNotSuccesInsureNOs = strNotSuccesInsureNOs & "," & cllPageInfor("K" & i)(1)
                        Next
                        If strNotSuccesInsureNOs <> "" Then strNotSuccesInsureNOs = Mid(strNotSuccesInsureNOs, 2)
                        If ModifyNotInsureNOs(strNotSuccesInsureNOs, strSuccesInsureNOs) = False Then
                            Exit Function
                        End If
                        zlInsureClinicSwap = True
                    End If
                    Exit Function
                Else
                    blnTransMedicare = True
                End If
            End If
            str保险结算 = GetMedicareStr(p)
            blnMedicareCheck = zlInsureCheck(str保险结算, strAdvance)
            ' Zl_病人门诊收费_医保更新
            gstrSQL = "Zl_病人门诊收费_医保更新("
            '  结帐id_In   门诊费用记录.结帐id%Type,
            gstrSQL = gstrSQL & "" & Val(cllPageInfor("K" & p)(0)) & ","
            '  结算序号_In 病人预交记录.结算序号%Type,
            gstrSQL = gstrSQL & "NULL,"
            '  保险结算_In Varchar2
            gstrSQL = gstrSQL & IIf(blnMedicareCheck, "'" & strAdvance & "'", "NULL") & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            strSuccesInsureNOs = strSuccesInsureNOs & "," & cllPageInfor("K" & p)(1)
            gcnOracle.CommitTrans: blnTrans = False
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
            blnAffair = True
    Next
    zlInsureClinicSwap = True
    Exit Function
Errhand:
    
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTrans Then
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mintInsure)
    End If
    If blnTransMedicare = False Then    '如果医保成功了，不删除划价单，费用失败可以重收
        Call DelMedicareTempNO(False, strBillNO)
    End If
    Call SaveErrLog
End Function

Private Sub DelMedicareTempNO(ByVal blnPriceSaved As Boolean, ByVal strBillNO As String)
'医保直接收费时,删除前一个事务提交的划价单
    If blnPriceSaved Then
        gstrSQL = "zl_门诊划价记录_DELETE('" & strBillNO & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 

Private Sub ShowBillChargeFee(ByVal lng结算序号 As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示收费成功的异常数据
    '编制:刘兴洪
    '日期:2011-08-26 18:59:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl结算金额 As Double, dbl未结金额 As Double
    Dim strInfor As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH:

    gstrSQL = "" & _
    "   Select decode(a.记录性质,1,'预存款',11,'预存款',结算方式) as 结算方式,  " & _
    "             nvl(sum(decode(nvl(校对标志,0),1, 1,0)* 冲预交),0) as 未结金额," & _
    "             nvl(sum(decode(nvl(校对标志,0),0,1,2,1,0)* 冲预交),0) as 结算金额" & _
    "   From 病人预交记录 A " & _
    "   Where 结算序号=[1]" & _
    "   Group by  decode(a.记录性质,1,'预存款',11,'预存款',结算方式) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng结算序号)
    strInfor = ""
    With rsTemp
        dbl结算金额 = 0
        Do While Not .EOF
            If Val(Nvl(rsTemp!结算金额)) <> 0 Then
                strInfor = strInfor & vbCrLf & "    " & Nvl(rsTemp!结算方式) & ":" & Format(rsTemp!结算金额, "0.00")
            End If
            dbl未结金额 = dbl未结金额 + Val(Nvl(rsTemp!未结金额))
            dbl结算金额 = dbl结算金额 + Val(Nvl(rsTemp!结算金额))
            .MoveNext
        Loop
    End With
    strInfor = "" & _
    "异常收费(请注意重新收取):" & vbCrLf & _
    "    当前已收取病人:" & Format(dbl结算金额, "0.00") & "元" & vbCrLf & _
    "    当前还未取病人:" & Format(dbl未结金额, "0.00") & "元" & vbCrLf & _
    "    收取成功的各项数据如下:" & strInfor
    MsgBox strInfor, vbExclamation, gstrSysName
    '清除界面所有显示
    Call ClearPayInfo
    mstrInNO = "": txtModi.Text = ""
    mlngFirstID = 0: mstrFirstWin = ""
    
    mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
            
    Call ClearPatientInfo(True)
    Call InitCommVariable
    Call ClearTotalInfo
    
    Call ClearBillRows: Call ClearMoney
    Call SetDisible(True): Call NewBill
    If txtPatient.Enabled Then txtPatient.SetFocus

    If gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub



Private Sub Set开单人开单科室(ByVal str开单人 As String, ByVal lng开单科室ID As Long)
'功能:根据开单人或开单科室ID设置开单科室及开单人,但不触发点击事件
       '利用公共函数CboSetIndex避免隐式调用cbo_click事件
    
    Dim str开单科室 As String, lng人员ID As Long
    
    'a.医生确定科室
    If gbyt科室医生 = 0 Then
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True))  '不触发click事件
        
        If cbo开单人.ListIndex = -1 And str开单人 <> "" Then
            lng人员ID = GetPersonnelID(str开单人, mrs开单人)
            cbo开单人.AddItem str开单人, 0
            cbo开单人.ItemData(cbo开单人.NewIndex) = lng人员ID
            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo开单人.NewIndex)
        End If
                
        If cbo开单人.ListIndex <> -1 Then
            cbo开单科室.Clear
            Call FillDept(mlngDeptID, cbo开单人.ItemData(cbo开单人.ListIndex))
        End If
        
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        If cbo开单科室.ListIndex = -1 And lng开单科室ID > 0 Then
            str开单科室 = GET部门名称(lng开单科室ID, mrs开单科室)
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
            str开单科室 = GET部门名称(lng开单科室ID, mrs开单科室)
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
    
    '开单人的专业技术职务
    If cbo开单人.ListIndex <> -1 And mobjBill.Pages(mintPage).开单人 <> "" And Not mrs开单人 Is Nothing Then
        mrs开单人.Filter = "ID=" & cbo开单人.ItemData(cbo开单人.ListIndex)
        If mrs开单人.RecordCount > 0 Then
            lblDuty.Caption = IIf(IsNull(mrs开单人!专业技术职务), "", mobjBill.Pages(mintPage).开单人 & "专业职务:" & mrs开单人!专业技术职务)
        Else
            lblDuty.Caption = ""
        End If
    Else
        lblDuty.Caption = ""
    End If
End Sub


Private Sub Set开单人开单科室Click(ByVal str开单人 As String, ByVal lng开单科室ID As Long)
'功能:根据开单人或开单科室ID设置开单科室及开单人,并触发点击事件
'     当Listindex=x时,如果Listindex的值本身等于x,就不会触发点击事件,所以要用API+Click强制调用
    Dim i As Long
    
    If gbyt科室医生 = 0 Then
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True)) '不触发click事件
        Call cbo开单人_Click
        
        '没有传入 开单科室ID 的时候以上面 cbo开单人_Click 缺省的为准
        If lng开单科室ID <> 0 Then
            Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
            Call cbo开单科室_Click
        End If
        
    Else
        '科室确定医生或各自独立输入
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        Call cbo开单科室_Click
        
        '没有传入 开单人 的时候，以上面 cbo开单科室_Click 缺省的为准
        If str开单人 <> "" Then
            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True)) '不触发click事件
            Call cbo开单人_Click
        End If
    End If
End Sub

Private Function ReadBill(ByVal strNo As String, ByVal bytFun As Byte, _
    Optional ByVal blnDelete As Boolean, Optional blnNoName As Boolean, _
    Optional blnShow As Boolean, Optional blnErrBill As Boolean) As Boolean
'功能：1.读取主界面原始单据或已退单据,2.读取划价收费,记帐审核单据,3.读取要部份退的单据
'调用：目前供以下操作调用
'      1.提划价单收费或记帐，包括输单据号提划价单收费，确定病人身份后自动提取划价单收费，多张收费时切换到单据页时重新读划价单
'      2.查看，调整，退费，销帐单据时读单据，包括读收费单，划价单，记帐单，记帐划价单
'参数：strNo=单据号
'      bytFun=0:收费单,1:划价单,2:门诊记帐单
'      blnDelete=是否进行退费或销帐(数量处理为准退数,金额计算处理)
'      blnShow=是否是因为切换单据读取(仅显示内容)
'      blnErrBill-显示异常单据
'返回：blnNoName=病人姓名是否为空
'说明：读取要退费的单据时(收费),排开误差处理费用,否则根据参数决定是否显示
'      因为多次部份退费时,每次都可能产生误差,原始的误差始终退不完。
    Dim rsTmp As ADODB.Recordset
    Dim rs结算 As ADODB.Recordset
    Dim i As Long, j As Long, k  As Long, intSign As Integer
    Dim strSQL As String, strSQL1 As String, strSQL2 As String
    Dim curBill实收 As Currency, curBill应收 As Currency
    Dim str费别 As String, str发药窗口 As String, lng结帐ID As Long
    Dim lng病人ID As Long
    Dim strPayDrugWins As String '执行部门ID|发药窗口;执行部门IDn|发药窗口n
    Dim strTemp As String, str医嘱序号 As String '退费时有效:分号分隔
    On Error GoTo errH
    strPayDrugWins = ""
    '收费时,要么在后备表中,要么在在线表中
    '记帐时,可能一张单据既在后备表中又在在线表中,是因为中途结帐允许只结一部分
    '因一张单据的数据少,下面简化为不区分,两张表联接查询
    
    '读取单据主体
    '----------------------------------------------------------------------------------------------------
    str医嘱序号 = ""
    If Not blnShow Then
        '收费单据多张票据只读取一个票号
        strSQL = _
        " Select A.结帐ID,A.实际票号 as 票据号,A.病人ID,0 as 主页ID,A.标识号,B.病人类型,B.险类," & _
        "       A.姓名,A.性别,A.年龄,A.费别,A.付款方式 ,0 as 病人病区ID,A.病人科室ID," & _
        "       A.开单部门ID,Nvl(A.加班标志,0) as 加班标志,a.开单部门ID," & _
        "       Nvl(A.婴儿费,0) as 婴儿费,A.开单人,A.划价人,A.操作员姓名,A.发生时间,A.登记时间," & _
        "       B.医疗付款方式,Nvl(A.是否急诊,0) as 是否急诊,A.门诊标志,Nvl(A.医嘱序号,0) as 医嘱序号,A.摘要,A.记录状态" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A") & " ,病人信息 B,人员表 C" & _
        " Where Rownum=1 And Nvl(A.操作员姓名,A.划价人)=C.姓名" & _
        "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
        "       And A.记录性质=" & IIf(bytFun = 2, 2, 1) & _
        "       And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
                IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
        "       And NO=[1] And A.病人ID=B.病人ID(+)" & _
                IIf(bytFun = 0, " And A.操作员姓名 is Not Null", "") & _
                IIf(bytFun = 1, " And A.操作员姓名 is Null And A.划价人 is Not NULL", "") & _
                IIf(bytFun = 2 And mbytInState = 0 And mbytBilling = 0, " And A.操作员姓名 is Not Null", "") & _
                IIf(bytFun = 2 And mbytInState = 0 And mbytBilling = 1, " And A.操作员姓名 is Null And A.划价人 is Not NULL", "") & _
                IIf(bytFun = 2 And mbytInState = 0 And mbytBilling = 2, " And A.操作员姓名 is Null And A.划价人 is Not NULL", "")
        '对住院病人提取划价单时，允许提取门诊发生的单据
        If bytFun = 1 And blnDelete = False And blnShow = False Then strSQL = strSQL & IIf(gint病人来源 = 2, " ", " And 门诊标志 In(1,3,4)")
        
        If mstrTime <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime))
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        End If
        If rsTmp.EOF Then
            MsgBox "没有发现指定的单据！", vbInformation, gstrSysName
            Exit Function
        ElseIf bytFun = 1 And Not mblnDoing And Not IsNull(rsTmp!姓名) And txtPatient.Text <> "" Then
            '判断是否相同病人，及要使用的病人信息
            If txtPatient.Text <> rsTmp!姓名 Then
                If MsgBox("单据中病人为""" & rsTmp!姓名 & """，与当前病人不符，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False And mlngDeptID > 0 Then
            If Val(Nvl(rsTmp!开单部门ID)) <> mlngDeptID Then
                MsgBox "你没有权限读取其它科室开单的单据！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Original.结帐ID = Nvl(rsTmp!结帐ID, 0) '用于医保门诊退费,一卡通单据修改
        If mbytBillSource <> 4 Then mbytBillSource = Val("" & rsTmp!门诊标志)   '只要有一张是体检,则认为全部是体检单据
        
    
        '病人相关信息提取:可能用于划价单收费,自动提取多张单据时不管
        '问题:30717
        If Not IsNull(rsTmp!登记时间) Then
            mobjBill.登记时间 = CDate(Format(rsTmp!登记时间, "yyyy-mm-dd HH:MM:SS"))
        End If
        If Val(Nvl(rsTmp!记录状态)) = 0 And gbytUnRegevent <> 0 Then
            mobjBill.病人ID = Val(Nvl(rsTmp!病人ID, mobjBill.病人ID))
            mobjBill.主页ID = Val(Nvl(rsTmp!主页ID, mobjBill.主页ID))
            mobjBill.标识号 = Nvl(rsTmp!标识号, mobjBill.标识号)
        Else
            mobjBill.病人ID = Val("" & rsTmp!病人ID)
            mobjBill.主页ID = Val("" & rsTmp!主页ID)
            mobjBill.标识号 = Nvl(rsTmp!标识号, 0)
        End If
        lng病人ID = mobjBill.病人ID
        mobjBill.床号 = ""            'IIf(gint病人来源 = 2, "" & rsTmp!床号, "")
        mobjBill.病区ID = Val("" & rsTmp!病人病区ID)
        mobjBill.科室ID = Val("" & rsTmp!病人科室ID)
        If mobjBill.费别 = "" Then
            mobjBill.费别 = Nvl(rsTmp!费别)
        End If
        mobjBill.Pages(mintPage).开单部门ID = Val("" & rsTmp!开单部门ID)
        mobjBill.Pages(mintPage).开单人 = "" & rsTmp!开单人
        mobjBill.Pages(mintPage).医嘱序号 = Val("" & rsTmp!医嘱序号)
        txtPatient.Locked = (mobjBill.病人ID <> 0 And "" & rsTmp!姓名 <> "新病人")    '为便于医保验卡,文本框不变为禁用状态颜色
        cboSex.Locked = txtPatient.Locked
        txt年龄.Locked = txtPatient.Locked
        cbo年龄单位.Locked = txtPatient.Locked
        txt退费摘要.Text = Nvl(rsTmp!摘要)
        
        If Not mblnDoing Then
            If Not IsNull(rsTmp!票据号) Then txtInvoice.Text = rsTmp!票据号: txtInvoice.SelStart = Len(txtInvoice.Text) '有才显示,划价单是没有的
            
            
            mobjBill.姓名 = Nvl(rsTmp!姓名)
            '75259：李南春,2014-7-10，病人姓名颜色处理
            Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), txtPatient.ForeColor, vbRed))
            mobjBill.性别 = Nvl(rsTmp!性别)
            'mobjBill.年龄 = Nvl(rsTmp!年龄)
            
            '病人姓名
            If mbytInFun = 0 And chkCancel.Value = 0 And (IsNull(rsTmp!姓名) Or IIf(mlngPrePati = 0, mstrPrePati = mobjBill.姓名, mlngPrePati = mobjBill.病人ID)) Then
                '同一个病人:空姓名或相同姓名
                
                If IsNull(rsTmp!姓名) Then
                    blnNoName = True
                    If Val(Nvl(rsTmp!记录状态)) = 0 And mstrPrePati = "" Then
                            
                    Else
                        txtPatient.Text = mstrPrePati '缺省为上一个病人姓名
                    End If
                Else
                    txtPatient.Text = Nvl(rsTmp!姓名)
                End If
            Else
                '不同的病人
                txtPatient.Text = Nvl(rsTmp!姓名)
                '刘兴洪:22343,51670
                If Not (mbytInFun = 0 And gTy_Module_Para.byt缴款控制 = 1) _
                    Or mstrPrePati = "" Then
                    mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
                    Call ClearPatientInfo
                    Call ClearTotalInfo
                    Call InitCommVariable
                    Call ClearMoney
                End If
            End If
            
            Call zlControl.CboSetText(cboSex, "" & rsTmp!性别)
            Call LoadOldData("" & rsTmp!年龄, txt年龄, cbo年龄单位)
            '刘兴洪:24348,由于在执行ClearPatientInfo清掉了年龄,因此应该将上面mobjBill.年龄 = Nvl(rsTmp!年龄),移置在下面才对.
            mobjBill.年龄 = Nvl(rsTmp!年龄)
            
            txt门诊号.Text = Nvl(rsTmp!标识号)
            
            If Nvl(rsTmp!门诊标志, 0) = 2 Or bytFun = 2 Or Not IsNull(rsTmp!医疗付款方式) Then
                cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, Nvl(rsTmp!医疗付款方式), True)
                If cbo医疗付款.ListIndex = -1 And Not IsNull(rsTmp!医疗付款方式) Then
                    cbo医疗付款.AddItem "0-" & rsTmp!医疗付款方式, 0
                    cbo医疗付款.ListIndex = cbo医疗付款.NewIndex
                End If
            Else
                cbo医疗付款.ListIndex = GetCboIndexByCode(cbo医疗付款, "" & rsTmp!付款方式)
                If cbo医疗付款.ListIndex = -1 And Not IsNull(rsTmp!付款方式) Then
                    cbo医疗付款.AddItem rsTmp!付款方式 & "-" & GetMedPayModeName(rsTmp!付款方式), 0
                    cbo医疗付款.ListIndex = cbo医疗付款.NewIndex
                ElseIf cbo医疗付款.ListIndex = -1 Then
                    cbo医疗付款.ListIndex = cbo.FindIndex(cbo医疗付款, mstr付款方式, True)
                End If
            End If
            
            If bytFun = 2 Then
                cboBaby.ListIndex = IIf(Val("" & rsTmp!婴儿费) > cboBaby.ListCount - 1, 0, Val("" & rsTmp!婴儿费))
                cboBaby.Enabled = mbytInState = 0 And mbytBilling <> 2
            End If
            
            txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
                        
            If Not rsTmp!病人ID Is Nothing Then Call LoadFeeInfor(Val("" & rsTmp!病人ID), blnDelete)
            
            If Nvl(rsTmp!是否急诊, 0) = 1 Then chk急诊.Value = 1: chk急诊.Visible = True
            mblnDo = False: chk加班.Value = Nvl(rsTmp!加班标志, 0): mblnDo = True
        End If
    End If
    
    '开单部门,开单人
    Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
    
    '收费读划价单时，目前允许修改开单人和开单科室,除非是医嘱发送过来的。
    If mbytInFun = 0 And mbytInState = 0 And chkCancel.Value = 0 Then
        cbo开单人.Locked = False
        cbo开单科室.Locked = False
        
        If mobjBill.Pages(mintPage).医嘱序号 <> 0 Then
            If cbo开单人.ListIndex <> -1 Then cbo开单人.Locked = True
            If cbo开单科室.ListIndex <> -1 Then cbo开单科室.Locked = True
        End If
    End If
    '读取结算方式
    '----------------------------------------------------------------------------------------------------
    If bytFun = 0 And Not blnShow Then
        '读取结算方式
        Call ReadBalance(CLng(rsTmp!结帐ID), blnDelete)
    End If
    
    '读取单据收费细目部份:分离发药时没有药房
    '---------------------------------------------------------------------------------------------
    If blnDelete Then
        '退费时不考虑后备表,前面的操作已禁用
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        '读取单据中原始记录的费用ID
        Dim strTableNo As String
        mblnHaveExcuteData = zlCheckIsExcuteData(strNo, IIf(bytFun = 2, 2, 1))     '60735
        
        '刘兴宏45685,58077
        strTableNo = " With 准退数  as ( " & _
        "            Select  A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIf(gbln药房单位, "/Nvl(B." & gstr药房包装 & ",1)", "") & ") as 准退数量" & _
        "            From 药品收发记录 A,药品规格 B " & _
        "           Where A.NO=[1]    And Mod(A.记录状态,3)=1  " & _
        "                       And (A.单据=[4] or A.单据=[5]) And A.审核人 is NULL  " & _
        "                       And A.药品ID=B.药品ID(+)  " & _
        "           Group by A.费用ID" & _
        "           Union ALL    "
        '取诊疗的部分退费
       If mblnHaveExcuteData Then
            '60735:在医嘱执行计价中存在数据时,则按医嘱执行计价中取数
            '77686,李南春,2014/9/18,单据类别限制
            strTableNo = strTableNo & _
            " Select Max(ID) As 费用id, Decode(Sign(Sum(数量)), -1, 0, Sum(数量)) As 准退数" & vbNewLine & _
            " From ( Select Decode(a.记录状态, 2, 0, a.Id) As ID, a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * Nvl(a.数次, 1) As 数量," & vbNewLine & _
            "              Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 原始数量" & vbNewLine & _
            "       From 门诊费用记录 A, 病人医嘱记录 M" & vbNewLine & _
            "       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And" & vbNewLine & _
            "             a.No = [1] And a.记录性质 = [3] And a.记录状态 In (1, 2, 3)　and A.价格父号 is null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From 病人医嘱附费 Where a.医嘱序号 = 医嘱id and a.No = NO and Mod(a.记录性质, 10) = 记录性质)" & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量" & vbNewLine & _
            "       From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M" & vbNewLine & _
            "       Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0" & vbNewLine & _
            "           And Instr('5,6,7', a.收费类别) = 0" & vbNewLine & _
            "           And (Exists (Select 1  From 病人医嘱执行  Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1)" & vbNewLine & _
            "                Or Exists (Select 1 From 病人医嘱发送 Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1))" & vbNewLine & _
            "          And a.No = [1] And a.记录性质 = [3] And a.记录状态 In (1, 3)　and A.价格父号 is null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From 病人医嘱附费 Where a.医嘱序号 = 医嘱id and a.No = NO and Mod(a.记录性质, 10) = 记录性质)" & _
            "       ) Q1" & vbNewLine & _
            " Where Not Exists (Select 1 From 药品收发记录 Where 费用id = Q1.Id And instr( ',8,9,10,21,24,25,26,',','||单据||',')>0) " & vbNewLine & _
            " Group by 医嘱ID,收费细目ID  Having Max(ID)<>0 )"
       Else
            'And A.费用性质=0 :61879,经与张永康确认,费用性质在门诊只有0-基础费用
            
            strTableNo = strTableNo & "" & _
             " Select Max(ID) as 费用ID,decode(sign(Sum(数量)),-1,0,Sum(数量))  as 准退数 " & _
             " From (  Select decode(J.记录状态,2,0,J.ID) as ID,J.医嘱序号 as 医嘱ID,J.收费细目ID, " & _
             "                       nvl(J.付数,1)*nvl(J.数次,1) as 数量,  decode(J.记录状态,2,0,nvl(J.付数,1)*nvl(J.数次,1)) as  原始数量" & _
             "              From  门诊费用记录 J,病人医嘱记录 M" & _
             "              where   J.医嘱序号=M.ID  " & _
             "                       And J.No=[1] and J.记录性质=[3] And J.记录状态 in (1,2,3) and J.价格父号 is null   " & _
             "                       And Exists(Select 1 From   病人医嘱发送 A Where   A.医嘱ID=J.医嘱序号 and  Nvl( A.执行状态, 0) <> 1 And A.No||''=[1]  ) " & _
             "                       And Exists(Select 1 From   病人医嘱计价 A Where   A.医嘱ID=J.医嘱序号 and A.收费细目ID=J.收费细目ID And A.费用性质=0 And  Nvl( A.收费方式, 0) =0 ) " & _
             "                       And Instr('5,6,7', j.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
             "                       And  instr(',C,D,F,G,K,',','||M.诊疗类别||',')=0  " & _
             "           Union all  " & _
             "           Select j.id, A.医嘱ID,a.收费细目ID,-1*nvl(a.数量,1)*nvl(C.本次数次,1) as 数量,0 as 原始数量 " & _
             "           From 病人医嘱计价 A,病人医嘱发送 B,病人医嘱执行 C,门诊费用记录 J,病人医嘱记录 M " & _
             "           where  A.医嘱ID=b.医嘱id  and b.医嘱id=c.医嘱id and b.发送号=c.发送号 And a.医嘱id=M.ID " & _
             "                       And Nvl(C.执行结果, 1) =1 And A.费用性质=0 And  Nvl( A.收费方式, 0) =0  And Nvl(b.执行状态, 0) <> 1 and  Nvl( B.执行状态, 0) <> 1 And B.No||''=[1]  " & _
             "                       And a.医嘱id=J.医嘱序号 and a.收费细目id=j.收费细目id  " & _
             "                       And Instr('5,6,7', j.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
             "                       And J.No=[1] and J.记录性质=[3] And J.记录状态 in (1,3) and J.价格父号 is null   " & _
             "                       And  instr(',C,D,F,G,K,',','||M.诊疗类别||',')=0)  " & _
             "   Group by 医嘱ID,收费细目ID  Having Max(ID)<>0 )"
        End If
        strSQL1 = _
            " Select A.ID,A.序号,A.收费细目ID," & _
            "       Nvl(A.付数,1)*A.数次" & IIf(gbln药房单位, "/Nvl(B." & gstr药房包装 & ",1)", "") & " as 原始数量" & _
            " From 门诊费用记录 A,药品规格 B" & _
            " Where A.NO=[1] And A.记录状态 IN(0,1,3) And A.价格父号 is null" & _
            "           And A.收费细目ID=B.药品ID(+) And A.记录性质=" & IIf(bytFun = 2, 2, 1) & _
                        IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
            "           And Nvl(A.附加标志,0)<>9"
         
        
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部份退费的记录)
        '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
        
        '需要排开医嘱计划中不为正常收取的费用:
        '   0-正常收取，1-检验试管费用；2-一次发送只收取一次；3-当天只收取一次；4-当天未执行收取一次；5-当天只收取一次，排斥其他项目；6-当天未执行收取一次，排斥其他项目；7-每天首次不收取
        
        strSQL = "" & _
            "  Select Nvl(价格父号,序号) as 序号 From 门诊费用记录  " & _
            "  Where 记录性质=[3] And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1     " & _
                        IIf(mstrTime <> "", " And 登记时间=[2]", "") & _
            "            And Nvl(附加标志,0)<>9 "
         If mblnHaveExcuteData = False Then
             '60735
           strSQL = strSQL & _
            "   Minus " & _
            "  Select Nvl(价格父号,序号) as 序号 " & _
            "  From 门诊费用记录 A1,病人医嘱计价 B1 " & _
            "  Where A1.医嘱序号=B1.医嘱id And A1.收费细目ID=B1.收费细目ID And B1.费用性质=0 And Nvl( B1.收费方式, 0) <>0  " & _
            "           And A1.记录性质=[3] And A1.记录状态 IN(0,1,3) And A1.NO=[1] And Nvl(A1.执行状态,0)=2 " & _
            "           And Instr('5,6,7', a1.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = a1.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
            "           And Not Exists (Select 1 From 药品收发记录 Where 费用id =a1.Id) " & _
                        IIf(mstrTime <> "", " And A1.登记时间=[2]", "") & _
            "           And Nvl(A1.附加标志,0)<>9 "
        End If
        '因为是将要汇总求有剩余数量的，所以不能用直接用时间限制，用序号限制
        strSQL = _
            " Select A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号 ," & _
            "       A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
                    IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位,max(A.医嘱序号) as 医嘱序号," & _
            "       Avg(Nvl(A.付数,1)) as 付数," & _
            "       Avg(A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
            "       Sum(A.标准单价" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
            "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            "       A.执行部门ID,D.名称 as 执行部门,A.附加标志,A.发药窗口" & _
            " From 门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+) " & _
            "       And A.收费细目ID=X.药品ID(+) And A.记录性质=[3]" & _
            "       And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            "       And Nvl(A.附加标志,0)<>9" & _
            " Group by A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),A.从属父号,A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
            "       B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志,A.发药窗口,X.药品ID,X." & gstr药房单位
            
        '最后计算结果：
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行,即剩余数量=0.(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        strSQL = strTableNo & vbCrLf & _
            " Select A.序号,A.从属父号,A.费别,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型,A.计算单位, " & _
            "       max(A.医嘱序号) as 医嘱序号," & _
            "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Avg(A.付数),1) as 准退付数," & _
            "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
            "        Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
            "        A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收," & _
            "        A.执行部门ID,A.执行部门,A.附加标志,A.发药窗口" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,准退数 C" & _
            " Where A.序号=B.序号 And A.收费细目ID=B.收费细目ID+0 And B.ID=C.费用ID(+)" & _
            " Group by A.序号,A.从属父号,A.费别,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型," & _
            "       A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行部门ID,A.执行部门,A.附加标志,A.发药窗口" & _
            " Having Sum(A.付数*A.数次)<>0"
            
        strSQL = _
            " Select A.序号,A.从属父号,A.费别,A.编码,A.类别,A.收费细目ID,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名," & _
            "           A.规格,A.费用类型,A.计算单位,A.医嘱序号,A.准退付数 as 付数,A.准退数次 as 数次,A.单价," & _
            "           A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
            "           A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
            "           A.执行部门ID,A.执行部门,A.附加标志,A.发药窗口" & _
            " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1 " & _
            " Where     A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
            " Order by A.序号"
    ElseIf bytFun = 2 And mbytInState = 0 And mbytBilling = 2 Then
        '划价时,仅从在线表读数据
        '读取记帐划价单内容(记帐审核),只读取未审核部份
        strSQL = _
            " Select Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
            "       A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
                    IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位,max(A.医嘱序号) as 医嘱序号," & _
            "       Avg(Nvl(A.付数,1)) as 付数," & _
            "       Avg(A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
            "       Sum(A.标准单价" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
            "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
            "       A.执行部门ID,D.名称 as 执行部门,A.附加标志,A.发药窗口" & _
            " From 门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.记录状态=0 And A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+) " & _
                " And A.收费细目ID=X.药品ID(+) And A.NO=[1] And A.记录性质=2" & _
            " Group by Nvl(A.价格父号,A.序号),A.从属父号,A.记录状态,A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
                " B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志,A.发药窗口,X.药品ID,X." & gstr药房单位
            
        strSQL = "Select" & _
            " A.序号,A.从属父号,A.费别,A.编码,A.类别,A.收费细目ID,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
            " A.计算单位,A.医嘱序号,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门ID,A.执行部门,A.附加标志,A.发药窗口" & _
            " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
            " Where     A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
            " Order by A.序号"
    Else
        '读取单据原始内容
        intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
        strSQL = _
            " Select Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
            "       A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
                    IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位,max(A.医嘱序号) as 医嘱序号," & _
            "       Avg(Nvl(A.付数,1)) as 付数," & _
            "       Avg(" & intSign & "*A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
            "       Sum(A.标准单价" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
            "       Sum(" & intSign & "*A.应收金额) as 应收金额,Sum(" & intSign & "*A.实收金额) as 实收金额," & _
            "       A.执行部门ID,D.名称 as 执行部门,A.附加标志,A.发药窗口" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录  A") & ",收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费类别 IN('5','6','7') And A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+) " & _
            "       And A.收费细目ID=X.药品ID And A.记录性质=" & IIf(bytFun = 2, 2, 1) & _
            "       And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
                    IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
                    IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
            " Group by Nvl(A.价格父号,A.序号),A.从属父号,A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
            "   B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志,A.发药窗口,X.药品ID,X." & gstr药房单位
        
        strSQL = strSQL & " Union ALL " & _
            " Select Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
            " A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            " A.计算单位,max(A.医嘱序号) as 医嘱序号,Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg(" & intSign & "*A.数次) as 数次,Sum(A.标准单价) as 单价," & _
            " Sum(" & intSign & "*A.应收金额) as 应收金额,Sum(" & intSign & "*A.实收金额) as 实收金额," & _
            " A.执行部门ID,D.名称 as 执行部门,A.附加标志,A.发药窗口" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录  A") & ",收费项目目录 B,收费项目类别 C,部门表 D" & _
            " Where A.收费类别 Not IN('5','6','7') And A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+) " & _
            " And A.记录性质=" & IIf(bytFun = 2, 2, 1) & _
            " And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
            " Group by Nvl(A.价格父号,A.序号),A.从属父号,A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
            " B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志,A.发药窗口"
            
        strSQL = "Select" & _
            " A.序号,A.从属父号,A.费别,A.编码,A.类别,A.收费细目ID,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
            " A.计算单位,A.医嘱序号,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门ID,A.执行部门,A.附加标志,A.发药窗口" & _
            " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
            " Order by A.序号"
    End If
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime), IIf(bytFun = 2, 2, 1), IIf(bytFun = 2, 9, 8), IIf(bytFun = 2, 25, 24))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, "", IIf(bytFun = 2, 2, 1), IIf(bytFun = 2, 9, 8), IIf(bytFun = 2, 25, 24))
    End If
    If rsTmp.EOF Then Exit Function
    
 
    j = 0
    Bill.Redraw = False
    Call ClearBillRows
    Bill.Rows = rsTmp.RecordCount + 1
    curBill应收 = 0: curBill实收 = 0
    For i = 1 To rsTmp.RecordCount
        '费别
        If Not IsNull(rsTmp!费别) Then
            If InStr(str费别 & ",", "," & rsTmp!费别 & ",") = 0 Then
                str费别 = str费别 & "," & rsTmp!费别
            End If
        End If
        
        '划价收费时重新处理发药窗口
        If mbytInFun = 0 And bytFun = 1 And InStr(",5,6,7,", rsTmp!编码) > 0 Then
            j = j + 1
            'If j = 1 Then '只有未分配发药窗口时才重新分配,以第一药品行为准
                If IsNull(rsTmp!发药窗口) Then
                    '分配窗口时，如果发现药房与上张单据不同，则清除缺省窗口,以免药房不同分配到相同窗口
                    If rsTmp!编码 = "5" Then
                        If rsTmp!执行部门ID <> mlng西药房 And mlng西药房 <> 0 Then mstr西窗 = ""
                        mlng西药房 = rsTmp!执行部门ID '记录该病人使用的药房(划价已定)
                    ElseIf rsTmp!编码 = "6" Then
                        If rsTmp!执行部门ID <> mlng成药房 And mlng成药房 <> 0 Then mstr成窗 = ""
                        mlng成药房 = rsTmp!执行部门ID
                    ElseIf rsTmp!编码 = "7" Then
                        If rsTmp!执行部门ID <> mlng中药房 And mlng中药房 <> 0 Then mstr中窗 = ""
                        mlng中药房 = rsTmp!执行部门ID
                    End If
                    
                    '71902,冉俊明,2014-04-09,同一个人病人不同时间段多张单据收费，分配同一个发药窗口，方便病人取药
                    '判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
                    str发药窗口 = Get未发药品发药窗口(lng病人ID, rsTmp!执行部门ID)
        
                    '不同类别的药品可能使用相同的药房,因此寻找以分配相同窗口
                    If str发药窗口 = "" Then
                        str发药窗口 = GetDrugWindow(rsTmp!执行部门ID, rsTmp!编码, tbsBill.SelectedItem.Index)
                    End If
                    If str发药窗口 = "" Then
                        str发药窗口 = Get发药窗口(zlDatabase.Currentdate, rsTmp!执行部门ID, rsTmp!编码, mstr西窗, mstr成窗, mstr中窗)
                    End If
                Else
                    str发药窗口 = rsTmp!发药窗口
                End If
                '问题:47489
                If InStr(1, strPayDrugWins & ";", ";" & rsTmp!执行部门ID & "|") = 0 Then
                    strPayDrugWins = strPayDrugWins & ";" & rsTmp!执行部门ID & "|" & str发药窗口
                End If
            'End If
        End If
        
        Bill.RowData(i) = rsTmp!序号 '价格父号(用于部份退费或销帐)
        Bill.TextMatrix(i, BillCol.类别) = rsTmp!类别
        Bill.TextMatrix(i, BillCol.从属父号) = Nvl(rsTmp!从属父号)
        Bill.TextMatrix(i, BillCol.医嘱序号) = Nvl(rsTmp!医嘱序号) & "," & Nvl(rsTmp!收费细目ID)
        If Val(Nvl(rsTmp!医嘱序号)) <> 0 And InStr(str医嘱序号 & ",", "," & Val(Nvl(rsTmp!医嘱序号)) & ",") = 0 Then
            str医嘱序号 = str医嘱序号 & "," & Val(Nvl(rsTmp!医嘱序号))
        End If
        '问题:29201
        strTemp = ""
        If Val(Nvl(rsTmp!从属父号)) <> 0 Then
            rsTmp.MoveNext
            strTemp = "┣"
            If rsTmp.EOF Then
                strTemp = "┗"
            ElseIf Bill.TextMatrix(i, BillCol.从属父号) <> Nvl(rsTmp!从属父号) Then
                strTemp = "┗"
            End If
            rsTmp.MovePrevious
            strTemp = "  " & strTemp & " "
        End If
        Bill.TextMatrix(i, BillCol.项目) = strTemp & rsTmp!名称
        Bill.TextMatrix(i, BillCol.商品名) = strTemp & Nvl(rsTmp!商品名)
        Bill.TextMatrix(i, BillCol.规格) = Nvl(rsTmp!规格)
        Bill.TextMatrix(i, BillCol.单位) = Nvl(rsTmp!计算单位)
        Bill.TextMatrix(i, BillCol.付数) = Nvl(rsTmp!付数)
        Bill.TextMatrix(i, BillCol.数次) = FormatEx(rsTmp!数次, 5)
        Bill.TextMatrix(i, BillCol.单价) = Format(rsTmp!单价, gstrFeePrecisionFmt)
        Bill.TextMatrix(i, BillCol.应收金额) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.实收金额) = Format(rsTmp!实收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.执行科室) = Nvl(rsTmp!执行部门)
        Bill.TextMatrix(i, BillCol.标志) = IIf(rsTmp!附加标志 = 1, "√", "")
        Bill.TextMatrix(i, BillCol.类型) = Nvl(rsTmp!费用类型)
        Bill.TextMatrix(i, BillCol.执行科室ID) = Nvl(rsTmp!执行部门ID)
        
        curBill应收 = curBill应收 + rsTmp!应收金额
        curBill实收 = curBill实收 + rsTmp!实收金额
        
        '设置销帐标志
        If InStr("销帐,退费", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
            Bill.TextMatrix(i, Bill.COLS - 1) = "√"
        End If
        
        rsTmp.MoveNext
    Next
 
    If str医嘱序号 <> "" And Bill.TextMatrix(0, Bill.COLS - 1) = "退费" Then
        str医嘱序号 = Mid(str医嘱序号, 2)
        Set mrs收费对照 = zlGet诊疗收费对照(str医嘱序号)
    Else
        Set mrs收费对照 = Nothing
    End If
    
    Set mrsDelInvoice = Nothing
    '25187
    Call LoadInvoiceData(strNo)
    Call ShowInvoiceInfor
    '显示单据小计
    lblSub应收.Caption = "应收:" & Format(curBill应收, gstrDec)
    lblSub实收.Caption = "实收:" & Format(curBill实收, gstrDec)
    lblAmount.Caption = ""
    
    '显示费别(包括一张单据中动态费别产生的多种费别)
    str费别 = Mid(str费别, 2)
    i = UBound(Split(str费别, ","))
    lbl动态费别.Visible = (i <> 0 And mbytInFun <> 2)
    cbo费别.Visible = Not (i <> 0 And mbytInFun <> 2)
    If i <> 0 And mbytInFun <> 2 Then
        lbl动态费别.Caption = str费别
        lbl动态费别.BorderStyle = 1
        lbl动态费别.Left = cbo费别.Left
    Else
        cbo费别.ListIndex = cbo.FindIndex(cbo费别, str费别, True)
        If cbo费别.ListIndex = -1 Then
            cbo费别.AddItem str费别, 0
            cbo费别.ListIndex = cbo费别.NewIndex
        End If
        cbo费别.Locked = bytFun <> 0    '收费提划价单时不允许修改费别,因为费用不能变
        cbo医疗付款.Locked = bytFun <> 0 And gintPriceGradeStartType >= 2 '收费提划价单时若医疗付款方式启用了价格等级则不允许修改费别,因为费用不能变
    End If
    cbo费别.TabStop = Not cbo费别.Locked And gbln费别
    
    '收费显示退款合计
    If bytFun = 0 And blnDelete Then
        lbl应缴.Caption = "退款"
        lbl应缴.ForeColor = vbRed
        txt应缴.ForeColor = vbRed
        
        mblnYB结算作废 = False
        MCPAR.医保接口打印票据 = False
        mintInsure = ChargeExistInsure(strNo, , lng结帐ID)
        If mintInsure <> 0 Then
            mblnYB结算作废 = gclsInsure.GetCapability(support门诊结算作废, , mintInsure)
            MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure, CStr(lng结帐ID))
            MCPAR.退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, , mintInsure)
        End If
        Call ReCalce退款
        ReInitPatiInvoice (False)
    ElseIf bytFun = 0 And blnErrBill Then
        '异常单据的处理
        If mintInsure = 0 Then
            mintInsure = ChargeExistInsure(strNo, , lng结帐ID)
        End If
    End If
    
    Call InitBillColumnColor
    Call SetColNum
    Bill.Redraw = True
    '读取单据收据费目汇总
    If Not blnShow Then
        If blnDelete Then
            '退费时不考虑后备表,前面的操作已禁止
            '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
    
            '读取药品收发记录中的准退数
            strSQL1 = _
                " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIf(gbln药房单位, "/Nvl(B." & gstr药房包装 & ",1)", "") & ") as 准退数量" & _
                " From 药品收发记录 A,药品规格 B" & _
                " Where A.NO=[1] And Mod(A.记录状态,3)=1 And A.审核人 is NULL" & _
                " And A.药品ID=B.药品ID(+) And A.单据 IN(" & IIf(bytFun = 2, "9,25", "8,24") & ")" & _
                " Group by A.费用ID"
            
            '整张费用单据(明细到收入项目)
            '执行状态应该在原始记录上判断(部分退药且部份退费的记录)
            strSQL = "Select Nvl(价格父号,序号) From 门诊费用记录" & _
                " Where 记录性质=" & IIf(bytFun = 2, 2, 1) & _
                " And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1" & _
                IIf(mstrTime <> "", " And 登记时间=[2]", "") & _
                " And Nvl(附加标志,0)<>9"
            strSQL = _
                " Select Sum(A.ID) as ID,A.序号,A.名称,A.收费类别," & _
                    " Sum(A.数量) as 剩余数量,Sum(A.应收金额) as 剩余应收," & _
                    " Sum(A.实收金额) as 剩余实收" & _
                " From (" & _
                    " Select Decode(A.记录状态,2,0,A.ID) as ID,A.序号," & _
                        IIf(gint分类合计 = 0, "A.收据费目", IIf(gint分类合计 = 2, "'单据合计'", "B.名称")) & " as 名称,A.收费类别," & _
                        " Nvl(A.付数,1)*A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & " as 数量," & _
                        " A.应收金额,A.实收金额" & _
                    " From 门诊费用记录 A,收入项目 B,药品规格 X" & _
                    " Where A.记录性质=" & IIf(bytFun = 2, 2, 1) & _
                        " And A.收费细目ID=X.药品ID(+) And A.NO=[1] And A.收入项目ID=B.ID" & _
                        " And Nvl(A.价格父号,A.序号) IN(" & strSQL & ") And Nvl(A.附加标志,0)<>9" & _
                        IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
                    " ) A" & _
                " Group by A.序号,A.名称,A.收费类别" & _
                " Having Sum(A.数量)<>0"
                        
            '最后计算结果,卫材也用准退数量:Nvl(B.准退数量,A.剩余数量)
            '有剩余数量无准退数量的有两种情况：
                '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
                '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
            strSQL = _
                " Select A.名称,Sum(A.剩余应收*(A.准退数量/A.剩余数量)) as 应收金额," & _
                " Sum(剩余实收*(A.准退数量/A.剩余数量)) as 实收金额 From (" & _
                " Select A.名称,A.剩余数量,A.剩余应收,A.剩余实收," & _
                " Decode(Instr(',4,5,6,7,',A.收费类别),0,A.剩余数量,Nvl(B.准退数量,A.剩余数量)) as 准退数量" & _
                " From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
                " Where A.ID=B.费用ID(+)" & _
                " ) A Group by A.名称"
        ElseIf bytFun = 2 And mbytInState = 0 And mbytBilling = 2 Then
            '读取记帐划价单内容(记帐审核),只读取未审核部份
            strSQL = _
                "Select " & IIf(gint分类合计 = 0, "A.收据费目", IIf(gint分类合计 = 2, "'单据合计'", "B.名称")) & " as 名称," & _
                " Sum(A.应收金额) as 应收金额," & _
                " Sum(A.实收金额) as 实收金额 " & _
                " From 门诊费用记录 A,收入项目 B" & _
                " Where A.记录状态=0 And A.记录性质=2" & _
                " And A.收入项目ID=B.ID And A.NO=[1]" & _
                IIf(gint分类合计 = 2, "", " Group By " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称"))
        Else
            '读取单据原始内容
            intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
            strSQL = _
                "Select " & IIf(gint分类合计 = 0, "A.收据费目", IIf(gint分类合计 = 2, "'单据合计'", "B.名称")) & " as 名称," & _
                " Sum(" & intSign & "*A.应收金额) as 应收金额," & _
                " Sum(" & intSign & "*A.实收金额) as 实收金额 " & _
                " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录  A") & " ,收入项目 B" & _
                " Where A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
                " And A.记录性质=" & IIf(bytFun = 2, 2, 1) & _
                IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
                " And A.NO=[1] And A.收入项目ID=B.ID" & _
                IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9 ", "") & _
                IIf(gint分类合计 = 2, "", " Group By " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称"))
        End If
        If mstrTime <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime))
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        End If
        If rsTmp.EOF Then Exit Function
        
        '刷新显示(收费要叠加)
        mshMoney.Rows = rsTmp.RecordCount + 1 + mintMoneyRow
        If mshMoney.Rows < M_MONEY_ROWS Then mshMoney.Rows = M_MONEY_ROWS
        Call SetMoneyList
        For i = mintMoneyRow + 1 To mshMoney.Rows - 1
            mshMoney.TextMatrix(i, 0) = ""
            mshMoney.TextMatrix(i, 1) = ""
            mshMoney.TextMatrix(i, 2) = ""
        Next
        curBill应收 = 0: curBill实收 = 0
        For i = mintMoneyRow + 1 To rsTmp.RecordCount + mintMoneyRow
            mshMoney.TextMatrix(i, 0) = mintBillNO + 1
            mshMoney.TextMatrix(i, 1) = rsTmp!名称
            mshMoney.TextMatrix(i, 2) = Format(rsTmp!实收金额, gstrDec)
            curBill应收 = curBill应收 + rsTmp!应收金额
            curBill实收 = curBill实收 + rsTmp!实收金额
            rsTmp.MoveNext
        Next
        On Error Resume Next
        For i = 1 To mshMoney.Rows - 1
            If mshMoney.TextMatrix(i, 0) = mintBillNO + 1 Then
                mshMoney.TopRow = i
            End If
        Next
        On Error GoTo errH
        
        '各类单据显示合计
        txt应收.Text = Format(mcurBill应收 + curBill应收, gstrDec)
        txt合计.Text = Format(mcurBill实收 + curBill实收, gstrDec)
        
        '划价时用来表示应缴,即分币处理后的金额
        If mbytInFun = 1 Then txt累计.Text = Format(CentMoney(txt合计.Text), "0.00")
        
        '用于记帐单据显示合计
        lblTotal.Caption = "合计:" & Format(curBill实收, gstrDec)
        
        '刷新收费累计
        If chkCancel.Value = 0 And mbytInFun = 0 And gbln累计 And Not mblnDoing Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
            txt累计.ToolTipText = "当前操作员今日收费累计额"
        End If
        
        '多单据收费支持:共用于各种单据
        With mobjBill.Pages(tbsBill.SelectedItem.Index)
            .NO = strNo
            .应收金额 = curBill应收
            .实收金额 = curBill实收
            
            '仅收费时收取划价单用
            If mbytInFun = 0 And bytFun = 1 Then
                '47489
                If strPayDrugWins <> "" Then strPayDrugWins = Mid(strPayDrugWins, 2)
                tbsBill.SelectedItem.Tag = strPayDrugWins ' str发药窗口
                Call ShowMoney(mintPage) '只需要计算当前单据
            End If
        End With
    End If
    If mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 3 Or mbytInState = 2) Then
        '退费
        With mTyDelFee
            .strNos = strNo
        End With
    End If
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReadBalance(ByVal lng结帐ID As Long, _
    ByVal blnDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算算方式
    '入参:lng结帐ID- 结帐ID
    '       blnDel-是否退费
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-30 08:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Long, i As Long, j As Long, dbl冲预交 As Double, str结算方式 As String
    Dim blnFind As Boolean, lngRow As Long, intSign As Integer
    On Error GoTo errHandle
     '退费
    With mTyDelFee
        Set .rsBlance = GetChargeBalance("", 0, lng结帐ID, mblnNOMoved)
    End With
    '读取单据原始内容时,显示各种金额
    '(部份)退费时,显示原始单据的结算金额
    intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
    
    '返回:收费相关的结算方式(性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类)
    '       字段:A.结帐ID,A.NO,A.性质,A.结算性质,A.结算方式,A.结算金额,
    '               A.卡类别ID,A.名称,A.是否全退,A.是否退现,A.结算号码,A.卡号,A.交易流水号,
    '               A.交易说明,A.结算序号,A.校对标志
    With mTyDelFee.rsBlance
        .Filter = 0: k = 0: j = 0
        Do While Not .EOF
            '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
            ',7,8:59673
            If Val(Nvl(!性质)) <> 1 And InStr(",1,2,", Val(Nvl(!结算性质))) > 0 Then
                k = k + 1
            End If
            If InStr(str结算方式 & ",", "," & Nvl(!结算方式) & ",") = 0 Then
                str结算方式 = str结算方式 & "," & Nvl(!结算方式)
            End If
            .MoveNext
        Loop
        If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
        mTyDelFee.blnSingleBalance = InStr(str结算方式, ",") = 0
        str结算方式 = ""
        
        If .RecordCount <> 0 Then .MoveFirst
        dbl冲预交 = 0
        vsBalance.Rows = 1
         mblnOlny预交 = True
        Do While Not .EOF
            If Val(Nvl(!性质)) = 1 Then
                dbl冲预交 = dbl冲预交 + intSign * Val(Nvl(!结算金额))
            Else    '查看收费单时全部显示在列表中
                mblnOlny预交 = False
                str结算方式 = Nvl(!结算方式, " ")
                ''1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
                
                If InStr("1,2,3,4,8,7", !结算性质) > 0 Or k > 1 Or Not blnDel Then
                    '医保结算方式,或单张单据多种结算方式
                    With vsBalance
                        blnFind = False: lngRow = 0
                        For i = 0 To .Rows - 1
                            If .TextMatrix(i, 0) = str结算方式 Then
                                lngRow = i
                                blnFind = True: Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, 0) = "" And lngRow = 0 Then blnFind = True
                        If Not blnFind Then
                              .Rows = .Rows + 1: lngRow = .Rows - 1
                        End If
                    End With
                    vsBalance.TextMatrix(lngRow, 0) = str结算方式
                    vsBalance.Cell(flexcpData, lngRow, 0) = Val(Nvl(!性质))
                    vsBalance.TextMatrix(lngRow, 1) = Val(vsBalance.TextMatrix(lngRow, 1)) + intSign * Val(Nvl(!结算金额))
                    vsBalance.Cell(flexcpData, lngRow, 1) = vsBalance.TextMatrix(lngRow, 1)
                    
                    If blnDel Then
                     '   vsBalance.Cell(flexcpForeColor, lngRow, 0, lngRow, vsBalance.COLS - 1) = vbRed
                    End If
                    If blnDel Then
                        '退费时，非医保的方式作特殊标注,以计算退费金额
                        '1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
                        Select Case Val(Nvl(!性质))
                        Case 3, 4 '3-医疗卡,4-结算卡,
                            If Val(Nvl(!是否退现)) = 1 Then vsBalance.RowData(lngRow) = -1
                            mTyDelFee.bln三方卡全退 = Val(Nvl(!是否全退)) = 1
                        Case Else
                            If InStr(",1,2,", !结算性质) > 0 Then vsBalance.RowData(lngRow) = -1
                        End Select
                    End If
                End If
                If InStr(",1,2,", !结算性质) > 0 And k = 1 Then
                    '非医保结算方式
                    mblnNotClick = True
                    zlControl.CboSetText cbo结算方式, str结算方式
                    mblnNotClick = False
                    If cbo结算方式.ListIndex = -1 Then
                        cbo结算方式.AddItem str结算方式
                        zlControl.CboSetIndex cbo结算方式.hWnd, cbo结算方式.NewIndex
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    If mblnOlny预交 = True And dbl冲预交 = 0 Then mblnOlny预交 = False
    txt预交冲款.Text = Format(intSign * dbl冲预交, "0.00")
    If dbl冲预交 <> 0 Then
        If vsBalance.Rows < 4 Then vsBalance.Rows = 4
        txt预交冲款.Visible = True: lbl预交冲款.Visible = True
    Else
        txt预交冲款.Visible = False: lbl预交冲款.Visible = False
        If vsBalance.Rows < 6 Then vsBalance.Rows = 6
    End If
    
    With vsBalance
        For i = 0 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                .TextMatrix(i, 1) = Trim(Format(Val(.TextMatrix(i, 1)), "##0.00###"))
            End If
        Next
    End With
    vsBalance.ToolTipText = "结算方式列表"
    mintReturnMode = cbo结算方式.ListIndex  '用于退费时,全退禁用结算方式时恢复初始的结算方式
    Call picAppend_Resize
    ReadBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub SetShowCol()
'功能：自动确定是否隐藏付数列
    mrsClass.Filter = "编码='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.付数) = 0
    ElseIf Bill.ColWidth(BillCol.付数) = 0 Then
        Bill.ColWidth(BillCol.付数) = 520 '强行显示
    End If
End Sub

Private Sub DelFactMoney()
'功能：删除单据中的工本费用(当不需要工本费时)
    Dim i As Long, p As Integer
    
    '先判断是否已经加入了工本费
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).工本费 Then
                Call DeleteDetail(i, p)
                
                '更新行号(当前单据)
                If mintPage = p Then
                    Call SetColNum(i)
                End If
                
                '只有工本费时，同时删除单据
                If mobjBill.Pages(p).Details.Count = 0 And fraBill.Visible Then
                    If tbsBill.Tabs.Count > 1 Then Call DelOneBill(p)
                End If
                
                Call ShowMoney(p)
                
                If CheckBillsEmpty Then ClearMoney
                Exit Sub
            End If
        Next
    Next
End Sub

Private Sub SetFactMoney()
'功能：收费时设置、显示、计算工本费
'说明：工本费自动加在当前显示的单据中
    Dim objDetail As Detail
    Dim colIncomes As New BillInComes
    Dim lngDoUnit As Long, blnExist As Boolean
    Dim intPage As Integer, lngRow As Long
    Dim i As Integer, p As Integer
    Dim int张数 As Integer, blnReCalc As Boolean
    
    int张数 = GetInvoiceCount '打印张数(不包含工本费)
    If int张数 = 0 Then Call DelFactMoney: Exit Sub '删除工本费
    
    '先判断是否已经加入了工本费
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).工本费 Then
                intPage = p: lngRow = i '存在的行
                blnExist = True: Exit For
            End If
        Next
        If blnExist Then Exit For
    Next
    
    '不存在则添加工本费
    If Not blnExist Then
        blnReCalc = True
        Set objDetail = Get工本费
        If objDetail Is Nothing Then Exit Sub '找不到工本费,不设置
        
        '寻找可以添加工本费的单据
        For p = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(p).NO = "" Then
                intPage = p: lngRow = mobjBill.Pages(p).Details.Count + 1
                Exit For
            End If
        Next
        If intPage = 0 Then
            '无可以编辑单据,新增一张单据
            If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then Exit Sub '不支持多单据
            Call AddNewBill
            intPage = mobjBill.Pages.Count: lngRow = 1
        ElseIf intPage = mintPage Then
            '是当前单据,处理界面
            If mobjBill.Pages(intPage).Details.Count >= Bill.Rows - 1 Then
                Bill.Rows = Bill.Rows + 1
            Else
                For i = 1 To Bill.COLS - 1
                    Bill.TextMatrix(Bill.Rows - 1, i) = ""
                Next
            End If
        End If
        
        With objDetail
            lngDoUnit = mobjBill.科室ID
            If lngDoUnit = 0 Then lngDoUnit = mobjBill.Pages(intPage).开单部门ID
            lngDoUnit = Get收费执行科室ID(.类别, .ID, .执行科室, lngDoUnit, Get开单科室ID, gint病人来源, , , , , mobjBill.病区ID)
            mobjBill.Pages(intPage).Details.Add "", objDetail, .ID, CInt(lngRow), 0, .类别, .计算单位, "", 1, 1, 0, lngDoUnit, colIncomes
        End With
        mobjBill.Pages(intPage).Details(lngRow).工本费 = True
    Else
        '如果存在且张数未变,则不用重算
        If mobjBill.Pages(intPage).Details(lngRow).数次 <> int张数 Then blnReCalc = True
    End If
    
    If blnReCalc Then
        '重新根据当前费用内容设置工本费数次
        mobjBill.Pages(intPage).Details(lngRow).数次 = int张数
        Call CalcMoney(intPage, lngRow)
        
        If mintPage = intPage Then
            Call ShowDetails(lngRow)
        End If
        Call ShowMoney(intPage)
    End If
End Sub

Private Sub ClearBillRows()
'功能：清除单据表格显示内容
    Dim i As Integer
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
        Call SetBillRowForeColor(i, Bill.ForeColor)
    Next
    Bill.ClearBill
    Call SetColNum
    
    lblSub应收.Caption = "应收:" & gstrDec
    lblSub实收.Caption = "实收:" & gstrDec
    lblAmount.Caption = ""
End Sub

Private Function GetOtherCTMGroups(lngRow As Long) As Integer
'功能：取当前单据中其它中药的付数
    Dim i As Integer
    
    GetOtherCTMGroups = 1
    For i = 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).收费类别 = "7" And i <> lngRow Then
            GetOtherCTMGroups = mobjBill.Pages(mintPage).Details(i).付数
            Exit For
        End If
    Next
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
Private Function GetWorkUnit(ByVal lng药品ID As Long, ByVal str类别 As String) As Boolean
'功能：取所有可供选择的药房
    Dim strSQL As String, bytDay As Byte
    Dim str药房 As String, lng开单科室ID As Long
    
    lng开单科室ID = mobjBill.科室ID     '开单科室优先
    If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    
    If str类别 = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.编码, c.简码, c.名称, b.工作性质, b.服务对象" & vbNewLine & _
        "From 收费执行科室 A, 部门性质说明 B, 部门表 C" & vbNewLine & _
        "Where a.执行科室id + 0 = b.部门id And b.工作性质 = '发料部门' And b.服务对象 IN(" & gint病人来源 & ",3) And b.部门id = c.Id And" & vbNewLine & _
        "      (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And (c.站点 = '" & gstrNodeNo & "' Or c.站点 Is Null) And" & vbNewLine & _
        "      (a.病人来源 Is Null Or A.病人来源=" & gint病人来源 & ") And" & vbNewLine & _
        "      (a.开单科室id Is Null Or a.开单科室id = [1] Or Exists (Select 1 From 病区科室对应 Where 科室id = [1] And a.开单科室id = 病区id)) And a.收费细目id = [2]" & vbNewLine & _
        "Order By b.服务对象, c.编码"
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
        If Not gbln药房上班安排 Then
            strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='" & str药房 & "'" & _
            "       And B.服务对象 IN(" & gint病人来源 & ",3) And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And (A.病人来源 is NULL Or A.病人来源=" & gint病人来源 & ")" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[1])" & _
            "       And A.收费细目ID=[2]" & _
            " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='" & str药房 & "'" & _
            "       And B.服务对象 IN(" & gint病人来源 & ",3) And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And D.部门ID=C.ID And D.星期=" & bytDay & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
            "       And (A.病人来源 is NULL Or A.病人来源=" & gint病人来源 & ")" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[1])" & _
            "       And A.收费细目ID=[2]" & _
            " Order by B.服务对象,C.编码"
        End If
    End If
    
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng开单科室ID, lng药品ID)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPrePaySum(Optional intPage As Integer) As Currency
    Dim curTotal As Currency, i As Long
    
    For i = 1 To mobjBill.Pages.Count
        If intPage = 0 Or i = intPage Then
            curTotal = curTotal + mobjBill.Pages(i).冲预交额
        End If
    Next
    GetPrePaySum = curTotal
End Function
Public Function GetBillSum(Optional bln应收 As Boolean, Optional ByVal intPage As Integer) As Currency
'功能：获取单据合计金额
'参数：intPage=指定单据,否则为所有单据
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim curTotal As Currency, intCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    For i = 1 To mobjBill.Pages.Count
        If intPage = 0 Or i = intPage Then
            If mobjBill.Pages(i).Details.Count > 0 Then
                For j = 1 To mobjBill.Pages(i).Details.Count
                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                        If bln应收 Then
                            curTotal = curTotal + mobjBill.Pages(i).Details(j).InComes(k).应收金额
                        Else
                            curTotal = curTotal + mobjBill.Pages(i).Details(j).InComes(k).实收金额
                        End If
                    Next
                Next
            Else    '提取划价单收费时没有明细费用
                If bln应收 Then
                    curTotal = curTotal + mobjBill.Pages(i).应收金额
                Else
                    curTotal = curTotal + mobjBill.Pages(i).实收金额
                End If
            End If
        End If
    Next
    
    '如果没有,再尝试从表格中取(仅一张单据时)
    If curTotal = 0 And tbsBill.Tabs.Count = 1 And Bill.Rows > 1 Then
        If Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.项目) = "") Then
            intCol = IIf(bln应收, BillCol.应收金额, BillCol.实收金额)
            For i = 1 To Bill.Rows - 1
                If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                    curTotal = curTotal + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
                End If
            Next
        End If
    End If
    GetBillSum = Format(curTotal, gstrDec)
End Function

Private Function Calc工本费(Optional ByVal intPage As Integer) As Currency
Dim i As Integer, j As Integer, k As Integer

    For i = 1 To mobjBill.Pages.Count
        If intPage = 0 Or i = intPage Then
            For j = 1 To mobjBill.Pages(i).Details.Count
                If mobjBill.Pages(i).Details(j).工本费 Then
                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                        Calc工本费 = Calc工本费 + mobjBill.Pages(i).Details(j).InComes(k).实收金额
                    Next
                End If
            Next
        End If
    Next
End Function

Private Sub txtModi_GotFocus()
    Call zlControl.TxtSelAll(txtModi)
End Sub

Private Sub txtModi_KeyPress(KeyAscii As Integer)
    '收费和划价功能可以在窗口调入单据修改
    Dim strNo As String
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtModi, KeyAscii)
    Else
        KeyAscii = 0
        
        '读取修改单据
        txtModi.Text = GetFullNO(txtModi.Text, 13)
        Call zlControl.TxtSelAll(txtModi)
        strNo = txtModi.Text
        
        Call ClearFullBill(False)
        
        mstrInNO = strNo
        Call LoadModifyNO(strNo, IIf(mbytInFun = 2, 2, 1))
    End If
End Sub

Private Function ExistOneCardSwap(strNo As String) As Boolean
'功能：检查指定的单据是否存在一卡通结算
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 1" & vbNewLine & _
            " From 门诊费用记录 a, 病人预交记录 b, 结算方式 c" & vbNewLine & _
            " Where a.结帐id = b.结帐id And b.结算方式 = c.名称 And a.记录性质 = 1 And a.No = [1] And c.性质=7" & vbNewLine & _
            "       And Rownum < 2"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    ExistOneCardSwap = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Private Sub LoadModifyNO(strNo As String, bytFlag As Byte)
'功能:划价,收费,记帐修改单据加载数据及相关检查
'     划价和收费除了在清单管理界面修改外,还可以在窗口界面输入单据号来修改
'参数:strNO-当前修改的单据号
'     bytFlag-记录性质，1-收费或划价，2-记帐
    Dim lng病人ID As Long, lng结帐ID As Long, bln急诊 As Boolean
    Dim strMessage As String, strNos As String
    Dim strTmp As String, i As Long
                
    'a.规则检查
    '-------------------------------------------------------------------------------------
    '是否已转入后备数据表中
    If mbytInFun = 0 Or mbytInFun = 2 And mblnCopyBill = False Then
        If zlDatabase.NOMoved("门诊费用记录", strNo, , bytFlag, Me.Caption) Then
            If Not ReturnMovedExes(strNo, bytFlag, Me.Caption) Then GoTo ExitHandle
        End If
        
        '已经退过费(部分)或销帐的单据不允许修改
        If BillExistDelete(strNo, bytFlag) Then
            strMessage = "该单据包含已" & IIf(mbytInFun = 2, "销帐", "退费") & "内容,不允许修改。": GoTo ExitHandle
        End If
    End If
    
    If mblnCopyBill = False Then
        '包含分批或时价药品的单据不允许修改
        If Not BillCanModi(strNo, bytFlag) Then
            strMessage = "该张单据包含分批或时价药品,可能库存已发生变化,不允许修改。": GoTo ExitHandle
        End If
        
        '要检查划价单,因为系统参数可能允许未收费的处方发药
        '如果包含部分执行或全部执行的项目,则退费后可能需要打印票据,不允许修改
        '                                   如果是记帐,则不一定可以全部冲销,不允许修改
        If HaveExecute(1, strNo, bytFlag) Then
            strMessage = "该单据包含完全执行或部分执行的项目,不允许修改。": GoTo ExitHandle
        End If
    End If
    
    '在读取病人信息前先读
    If mbytInFun = 2 And mblnCopyBill = False Then
        Original.实收合计 = GetBillMoney(1, strNo, , IIf(mbytInFun = 2, 2, 1))
    ElseIf mbytInFun = 0 Then
        Call GetBillPay(strNo, Original.冲预交款, Original.应缴金额)
    End If
    
    '收费功能相关检查
    If mbytInFun = 0 Then
        '未收费的划价单不允许修改
        If Bill未收费(strNo, 1) Then
            strMessage = "该划价单据尚未收费,不允许修改。": GoTo ExitHandle
        End If
        If gblnMultiBalance Or gTy_Module_Para.bln工本费 Then strNos = GetMultiNOs(strNo, , , True)
        If gblnMultiBalance And InStr(1, strNos, ",") > 0 Then
            If CheckSingleBalance(strNos) = False Then
                strMessage = "多张单据使用多种结算方式模式下不允许修改多张单据。": GoTo ExitHandle
            End If
        End If
        If gTy_Module_Para.bln工本费 Then
            If InStr(1, strNos, ",") > 0 Then
                If BillExistFact(strNo) Then
                    strMessage = "与该单据一起收费的多张单据中包含工本费，不能进行修改。": GoTo ExitHandle
                End If
            End If
        End If
        If mbytInFun = 0 Then
            '进行了医保补充结算，不允许修改
            If CheckBillExistReplenishData(1, , Replace(strNos, "'", "")) = True Then
                MsgBox "当前单据进行了医保补充结算，不允许修改！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        mintInsure = ChargeExistInsure(strNo, lng病人ID, lng结帐ID, bln急诊)
        If mintInsure > 0 Then
            If InStr(mstrPrivs, "保险收费") = 0 Then
                strMessage = "你没有权限对医保病人的单据进行修改操作。": GoTo ExitHandle
            End If
        
            '修改医保单据先验证身份
            Call txtPatient_KeyPress(13)
            If mstrYBPati = "" Then GoTo ExitHandle
            
            '验证的病人身份必须相符
            If lng病人ID <> mobjBill.病人ID Then
                strMessage = "验证的病人身份与单据中的病人身份不符，不能修改单据。"
                mstrYBPati = "": mintInsure = 0: GoTo ExitHandle
            End If
            
            '判断每一种结算方式是否支持退费,只要其中一种不支持,则不允许修改
            If Not Check门诊结算作废(lng结帐ID, mintInsure) Then
                mstrYBPati = "": mintInsure = 0: GoTo ExitHandle
            End If
            
            Original.结帐ID = lng结帐ID '记录修改时要退费的单据的结帐ID
            If bln急诊 Then chk急诊.Value = 1
            
            '因为是修改,将个人帐户余额加上
            mcur个帐余额 = mcur个帐余额 + Read个人帐户结算(lng结帐ID)
            sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
            
        Else
            '是否有非医保病人的退费权限
            If InStr(mstrPrivs, "允许非医保病人") = 0 Then
                strMessage = "你没有权限对非医保病人的单据进行修改操作。": GoTo ExitHandle
            End If
           
            If mblnOneCard Then
                If ExistOneCardSwap(strNo) Then
                    strMessage = "该单据采用了一卡通结算,不允许修改。": GoTo ExitHandle
                End If
            End If
           
        End If
        
    ElseIf mbytInFun = 2 And mblnCopyBill = False Then
        '未全部审核或多次审核的不允许修改
        If Not BillIdentical(strNo) Then
            strMessage = "单据中包含部份不全完审核或分多次审核的内容，不允许修改。": GoTo ExitHandle
        End If
                    
        '如果已经结帐,根据参数决定是否允许修改
        If HaveBilling(1, strNo) Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("该记帐单已经结帐,要修改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo ExitHandle
                Case 2
                    strMessage = "该记帐单已经结帐,不允许修改。": GoTo ExitHandle
            End Select
        End If
    End If
    
    
    'b.读取费用明细到单据对象
    '---------------------------------------------------------------------------------------------------
    Set mobjBill = ImportBill(strNo, IIf(mblnCopyBill, False, True), mbytInFun, mintInsure, , , _
        mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    If mobjBill.NO = "" Then
        strMessage = "读取单据失败。": GoTo ExitHandle
    End If
    
    If mblnCopyBill = False Then
        If mbytInFun = 0 And InStr(mstrPrivs, "所有操作员") <= 0 Then
            If UserInfo.姓名 <> mobjBill.操作员姓名 Then
                strMessage = "你没有""所有操作员""权限,不能修改" & mobjBill.操作员姓名 & "的单据。": GoTo ExitHandle
            End If
        End If
        
        If Not BillOperCheck(Choose(mbytInFun + 1, 2, 3, 4), mobjBill.操作员姓名, mobjBill.登记时间, "修改", strNo, , IIf(mbytInFun = 2, 2, 1)) Then
            GoTo ExitHandle
        End If
        
        '医嘱生成的划价单,或已收费的医嘱生成的划价单,不允许修改
        If mobjBill.Pages(1).医嘱序号 <> 0 Then strMessage = "由医嘱产生的单据不允许修改！": GoTo ExitHandle
    End If
    
    'c.显示信息
    '------------------------------------------------------------------------------------------------------
    mbln不重算价格 = True
        Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
        Call LoadAndSeek费别        '加载费别和动态费别
        
        'a.已建档的病人
        If mobjBill.病人ID <> 0 Then
            If mstrYBPati = "" Then '医保病人在前面已验证身份
                txtPatient.Text = "-" & mobjBill.病人ID
                Call txtPatient_KeyPress(13)
            End If
        Else
        'b.划价或收费时，未建档的病人
            txtPatient.Text = mobjBill.姓名
            cboSex.ListIndex = cbo.FindIndex(cboSex, mobjBill.性别, True)
            Call LoadOldData(mobjBill.年龄, txt年龄, cbo年龄单位)
            txt门诊号.Text = IIf(mobjBill.标识号 = 0, "", mobjBill.标识号)
            cbo医疗付款.ListIndex = GetCboIndexByCode(cbo医疗付款, mobjBill.床号)
            
            strTmp = GetBill费别(mobjBill)
            If strTmp <> "" Then cbo费别.ListIndex = cbo.FindIndex(cbo费别, strTmp, True)
        End If
        If cbo费别.ListIndex = -1 And cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
    mbln不重算价格 = False
    
    
    '取第一药品行
    For i = 1 To mobjBill.Pages(1).Details.Count
        If InStr(",5,6,7,", mobjBill.Pages(1).Details(i).收费类别) > 0 Then
            mlngFirstID = mobjBill.Pages(1).Details(i).执行部门ID
            mstrFirstWin = mobjBill.Pages(1).Details(i).发药窗口
            Exit For
        End If
    Next
    
        
    If mblnCopyBill Then
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        mstrInNO = "" '复制内容后清除,以区别修改
    Else
        '显示的是原单据号,保存的是新单据号
        cboNO.Text = strNo
        txtDate.Text = Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss")
    End If
    mblnDo = False: chk加班.Value = mobjBill.加班标志: mblnDo = True
    Bill.Rows = mobjBill.Pages(1).Details.Count + 1
    
    If mbytInFun = 2 Then
        Call zlControl.CboSetIndex(cboBaby.hWnd, mobjBill.婴儿费)
        If cbo开单科室.ListIndex <> -1 Then cboBaby.Enabled = is产科(cbo开单科室.ItemData(cbo开单科室.ListIndex), mrs开单科室)
    End If
    
    '修改时应保存当前操作员的名字
    mobjBill.操作员编号 = UserInfo.编号: mobjBill.操作员姓名 = UserInfo.姓名
    
    '缺省为原单据的结算方式
    If mbytInFun = 0 Then
        strTmp = GetBalanceName(strNo)
        If strTmp <> "" Then
            i = cbo.FindIndex(cbo结算方式, strTmp, True)
            If i <> -1 Then cbo结算方式.ListIndex = i
        End If
    End If
    
    '新病人
    If IIf(mlngPrePati = 0, mstrPrePati = mobjBill.姓名, mlngPrePati = mobjBill.病人ID) Then
        mcurBill实收 = 0:  mcurBill应收 = 0: mcurBill应缴 = 0
        mintBillNO = 0: mintMoneyRow = 0
    End If
    
''    '如果不是优先使用预交款,则保持原先的冲款额
''    If Not gblnPrePayPriority And Original.冲预交款 > 0 And txt预交冲款.Enabled Then
''       txt预交冲款.Text = Format(IIf(Original.冲预交款 > Val(sta.Panels(Pan.C4预交信息).Tag), Val(sta.Panels(Pan.C4预交信息).Tag), Original.冲预交款), "0.00")
''    End If
''
    If gintPriceGradeStartType < 2 Then
        If gbln从项汇总折扣 Then Call CalcMoneys
    Else
        Call CalcMoneys
    End If
    Call ShowDetails
    Call ShowMoney
              
                
    'd.界面控制
    '------------------------------------------------------------------------------------------
'    txt本次应缴.Visible = True: lbl应缴.Caption = "再缴"
    
    Call InitBillColumnColor
    Call SetColNum
    Call SetPatientEnableModi(mobjBill.病人ID = 0)
           
    chkCancel.Enabled = False: cmdDelete.Enabled = False
    cmdAddBill.Enabled = False
    
    If Me.Visible Then
        Bill.Active = True: Bill.SetFocus
        If mintInsure > 0 Then txtModi.Enabled = False  '医保单据必须按"取消"按钮来调用取消身份验证
    End If
    
    Exit Sub
ExitHandle:
    If strMessage <> "" Then
        
        MsgBox strMessage, vbInformation, gstrSysName
    End If
    If Me.Visible Then
        Set mobjBill = New ExpenseBill
        txtModi.Text = ""
        If txtModi.Visible And txtModi.Enabled Then txtModi.SetFocus
    Else
        Unload Me
    End If
End Sub
'''
'''
'''Private Sub txt缴款_LostFocus()
'''    mblnHotKey = False
'''    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
'''End Sub

Private Sub SetPatientEnableModi(blnModi As Boolean)
    
    txtPatient.Locked = Not blnModi
    
    If blnModi Then
        txtPatient.BackColor = &HFFFFFF
    Else
        txtPatient.BackColor = &HE0E0E0
    End If

    cboSex.Locked = txtPatient.Locked
    txt年龄.Locked = txtPatient.Locked
    txt年龄.BackColor = txtPatient.BackColor
    cbo年龄单位.Locked = txtPatient.Locked
End Sub

Private Sub SetInputItem()
    '输入项目
    If Not gbln性别 Then cboSex.TabStop = False
    If Not gbln年龄 Then txt年龄.TabStop = False: cbo年龄单位.TabStop = False
    If Not gbln费别 Then cbo费别.TabStop = False
    If Not gbln医疗付款 Then cbo医疗付款.TabStop = False
    If Not gbln加班 Then chk加班.TabStop = False
    If Not gbln开单日期 Then txtDate.TabStop = False
    If Not gbln开单人 Then cbo开单人.TabStop = False
End Sub

Private Function SaveModi() As Boolean
    '功能：保存当前修改的费用单据
    Dim strSQL As String
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入合法的费用时间！", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    strSQL = "zl_病人费用记录_Update('" & cboNO.Text & "'," & IIf(mbytInFun = 2, 2, 1) & "," & _
        "'" & zlStr.NeedName(cbo开单人.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'))"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'功能：显示\隐藏销帐标志列
    Dim i As Integer, blnACT As Boolean
    If blnShow Then
        If InStr("销帐,退费", Bill.TextMatrix(0, Bill.COLS - 1)) = 0 Then
            Bill.Redraw = False
            Bill.COLS = Bill.COLS + 1
            If mbytInFun = 2 Then
                Bill.TextMatrix(0, Bill.COLS - 1) = "销帐"
            Else
                Bill.TextMatrix(0, Bill.COLS - 1) = "退费"
            End If
            Bill.ColAlignment(Bill.COLS - 1) = 4
            Bill.ColWidth(Bill.COLS - 1) = 550
            Bill.ColData(Bill.COLS - 1) = BillColType.CheckBox
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.COLS - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.COLS - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.类别) = GetOrigColWidth(BillCol.类别) - 100
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目) - 100
            Bill.ColWidth(BillCol.执行科室) = GetOrigColWidth(BillCol.执行科室) - 200
            
            Bill.ColWidth(BillCol.单价) = GetOrigColWidth(BillCol.单价) - 50
            Bill.ColWidth(BillCol.应收金额) = GetOrigColWidth(BillCol.应收金额) - 50
            Bill.ColWidth(BillCol.实收金额) = GetOrigColWidth(BillCol.实收金额) - 50
            Bill.Redraw = True
        End If
    Else
        If InStr("销帐,退费", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
            Bill.Redraw = False
            Bill.COLS = Bill.COLS - 1
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
    Dim bln As Boolean, i As Integer
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.行) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True, Optional intPage As Long) As Integer
'功能：检查指定药品行的职务是否与当前医生的职务相匹配
'参数：tmpDetail=正在输入的项目,不传为所有单据所有行,blnCommon=是否正常的判断,否则为医保或公费病人的判断
'返回：不匹配的行,0为正确,intPage=单据页号
'说明：职务：1=正高,2=副高,3=中级,4=助理/师级,5=员/士,9=待聘
    Dim i As Integer, p As Integer, strTmp As String
    Dim int职务A As Integer, int职务B As Integer
    Dim strMsg As String
    
    strTmp = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    intPage = 0
    
    If tmpDetail Is Nothing Then
        For p = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(p).开单人 <> "" Then
                '每张单据开单人不同,当前单据的开单的人职务
                Call GetOperatorInfo(mobjBill.Pages(p).开单人, , int职务A)
                
                For i = 1 To mobjBill.Pages(p).Details.Count
                    If InStr(",5,6,7,", mobjBill.Pages(p).Details(i).收费类别) > 0 Then
                        If mobjBill.Pages.Count > 1 Then strMsg = "在单据 " & p & "中"
                        If Not blnCommon Then
                            int职务B = Val(Right(mobjBill.Pages(p).Details(i).Detail.处方职务, 1))
                            If int职务B > 0 Then
                                If int职务A = 0 Then
                                    strMsg = "对医保或公费" & gstrCustomerAppellation & "," & strMsg & _
                                        "第 " & p & " 页 " & i & " 行药品""" & mobjBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                                        "而""" & mobjBill.Pages(p).开单人 & """未设置职务！"
                                    CheckDuty = 1: intPage = p
                                ElseIf int职务B < int职务A Then
                                    strMsg = "对医保或公费" & gstrCustomerAppellation & "," & strMsg & _
                                        "第 " & p & " 页 " & i & " 行药品""" & mobjBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                                        "而""" & mobjBill.Pages(p).开单人 & """职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                                    CheckDuty = i: intPage = p: Exit For
                                End If
                            End If
                        Else
                            int职务B = Val(Left(mobjBill.Pages(p).Details(i).Detail.处方职务, 1))
                            If int职务B > 0 Then
                                If int职务A = 0 Then
                                    strMsg = strMsg & "第 " & p & " 页 " & i & " 行药品""" & mobjBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                                        "而""" & mobjBill.Pages(p).开单人 & """未设置职务！"
                                    CheckDuty = 1: intPage = p
                                ElseIf int职务B < int职务A Then
                                    strMsg = strMsg & "第 " & p & " 页 " & i & " 行药品""" & mobjBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                                        "而""" & mobjBill.Pages(p).开单人 & """职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                                    CheckDuty = i: intPage = p: Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    ElseIf mobjBill.Pages(mintPage).开单人 <> "" Then
        If InStr(",5,6,7,", tmpDetail.类别) = 0 Then Exit Function
        Call GetOperatorInfo(mobjBill.Pages(mintPage).开单人, , int职务A)
        
        If Not blnCommon Then
            int职务B = Val(Right(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strMsg = "对医保或公费" & gstrCustomerAppellation & ",药品""" & tmpDetail.名称 & _
                        """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                        "而当前开单人未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strMsg = "对医保或公费" & gstrCustomerAppellation & ",药品""" & tmpDetail.名称 & _
                        """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                        "而当前开单人职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        Else
            int职务B = Val(Left(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strMsg = "药品""" & tmpDetail.名称 & """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                        "而当前开单人未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strMsg = "药品""" & tmpDetail.名称 & """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                        "而当前开单人职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInhibitiveByNurse(ByVal intPage As Integer) As Boolean
'功能：判断指定单据中是否有护士禁止输入的内容
    Dim rsTmp As New ADODB.Recordset
    Dim bln护士 As Boolean, strSQL As String
    Dim i As Integer
    
    CheckInhibitiveByNurse = False
    If mobjBill.Pages(intPage).开单人 <> "" Then
        Call GetOperatorInfo(mobjBill.Pages(intPage).开单人, bln护士)
        If Not bln护士 Then Exit Function
        
        If mobjBill.Pages(intPage).NO = "" Then
            For i = 1 To mobjBill.Pages(intPage).Details.Count
                If InStr(",E,M,4,", mobjBill.Pages(intPage).Details(i).收费类别) = 0 Then
                    CheckInhibitiveByNurse = True: Exit Function
                End If
            Next
'            '划价单不再检查
        End If
    End If
End Function

Private Sub FillDoctor(Optional lng科室ID As Long)
'功能：根据指定的开单科室ID读取并填写医生列表,但不缺省医生
    Dim strOldID As String
    Dim bln仅操作员部门 As Boolean, str部门性质 As String
    
    cbo开单人.Clear
    If mbytInFun = 1 And mbytInState = 0 Then '113577
        bln仅操作员部门 = zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False And gblnUserIsClinic
    End If
    If mbytInFun = 1 Then
        str部门性质 = "'临床','手术','治疗','检查','检验','产科'"
        Call GetDoctor(lng科室ID, mrs开单人, bln仅操作员部门, str部门性质)
    Else
        Call GetDoctor(lng科室ID, mrs开单人, bln仅操作员部门)
    End If
    
    Do While Not mrs开单人.EOF
    '70857:刘尔旋,2014-03-07,开单人简码一致时存在加载重复的问题
        If InStr("," & strOldID & ",", "," & mrs开单人!ID & ",") = 0 Then
            If gbyt开单人显示 = 1 Then
                cbo开单人.AddItem mrs开单人!简码 & "-" & mrs开单人!姓名
            Else
                cbo开单人.AddItem mrs开单人!编号 & "-" & mrs开单人!姓名
            End If
            cbo开单人.ItemData(cbo开单人.NewIndex) = mrs开单人!ID
            strOldID = strOldID & mrs开单人!ID & ","
        End If
        mrs开单人.MoveNext
    Loop
End Sub



Private Sub FillDept(ByVal lngDeptID As Long, Optional lng人员ID As Long)
'功能：读取并加载科室列表,但不缺省科室
'参数：
'   lngDeptID-当前操作的病区
'   lng人员ID=只读取指定人员所在科室(包含非缺省的)
'返回：科室个数
    
    Dim strSQL As String, i As Long, lngOldDepID As Long
    Dim strDepts As String  '指定人员所属的多个部门
    Dim bln仅操作员部门 As Boolean, str性质 As String
        
    cbo开单科室.Clear
    If mbytInFun = 1 And mbytInState = 0 Then '113577
        bln仅操作员部门 = zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False And gblnUserIsClinic
    End If
    If mrs开单科室 Is Nothing Then
        If mbytInFun = 1 Then
            str性质 = "'临床','手术','治疗','检查','检验','产科'"
        Else
            str性质 = "'临床','产科'"
        End If
        Call GetDoctorDept(mrs开单科室, bln仅操作员部门, str性质, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), 0, lngDeptID))
    End If
   
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
            If lngOldDepID <> mrs开单科室!ID Then   '一个部门可能同时属于产科和临床,不加载相同的
                cbo开单科室.AddItem IIf(zlIsShowDeptCode, mrs开单科室!编码 & "-", "") & mrs开单科室!名称     '见问题:27658
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = mrs开单科室!ID
                lngOldDepID = mrs开单科室!ID
            End If
            mrs开单科室.MoveNext
        Next
    End If
End Sub

Private Function CheckDrugExist(objDetail As Detail) As Boolean
'功能：判断指定药品或卫材(跟踪在用)在单据中是否已经存在
'参数：objDetail=项目,intRow=要判断的行
'说明：时价或分批在同一执行科室禁止重复输入(这里仅提示,保存时禁止)
'      非时价的分批药品，在不同的单据上有相同的，允许不合并，不提醒
    Dim i As Integer, p As Integer
    Dim strTmp As String
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If Not (p = mintPage And i = Bill.Row) And InStr(",4,5,6,7,", mobjBill.Pages(p).Details(i).收费类别) > 0 Then
                If mobjBill.Pages(p).Details(i).Detail.ID = objDetail.ID Then
                    If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                    If (mobjBill.Pages(p).Details(i).Detail.分批 Or mobjBill.Pages(p).Details(i).Detail.变价) _
                        And (objDetail.分批 Or objDetail.变价) Then
                        
                        '非时价的分批药品，在不同的单据上有相同的，允许不合并，不提醒
                        If objDetail.变价 Or (Not objDetail.变价 And objDetail.分批 And mintPage = p) Then
                            If objDetail.类别 = "4" Then
                                If MsgBox("卫生材料""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？" & _
                                    vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价材料,重复输入时必须保证它们的发料部门不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    CheckDrugExist = True
                                End If
                            Else
                                If MsgBox("药品""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？" & _
                                    vbCrLf & vbCrLf & "注意：该药品为分批或时价药品,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    CheckDrugExist = True
                                End If
                            End If
                            Exit Function
                        End If
                    Else
                        If objDetail.类别 = "4" Then
                            If MsgBox("卫生材料""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                CheckDrugExist = True
                            End If
                        Else
                            If MsgBox("药品""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                CheckDrugExist = True
                            End If
                        End If
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
End Function

Private Function CheckFeeType(Optional intRow As Integer) As Boolean
'功能：根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    Dim strSQL As String, strType As String
    Dim i As Integer, p As Integer
    Dim strTmp As String, bln医保 As Boolean, bln公费 As Boolean
    
    On Error GoTo errHandle
    
    CheckFeeType = True
    
    '无法检查
    If cbo医疗付款.ListIndex = -1 Then Exit Function
    '医保或公费病人
    '问题:45605
    If zlIsCheckMedicinePayMode(zlStr.NeedName(cbo医疗付款), bln医保, bln公费) = False Then Exit Function
    '只检查医保病人和公费病人
    strType = IIf(bln医保, 1, 2)
    
    '读取检查数据
    If mrs费用类型 Is Nothing Then
        strSQL = " Select '医保' As 类别,编码,名称 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Union All " & _
                 " Select '公费' As 类别,编码,名称 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") "
        Set mrs费用类型 = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs费用类型, strSQL, Me.Caption)
    End If
    mrs费用类型.Filter = ""
    If mrs费用类型.RecordCount = 0 Then Exit Function
        
    If strType = "1" Then
        strSQL = " And 类别='医保'"
    Else
        strSQL = " And 类别='公费'"
    End If
    
    
    If intRow > 0 Then
        If mobjBill.Pages(mintPage).Details(intRow).Detail.类型 = "" Then
            MsgBox """" & mobjBill.Pages(mintPage).Details(intRow).Detail.名称 & """的费用类型未设置！", vbInformation, gstrSysName
            CheckFeeType = False
        Else
            mrs费用类型.Filter = "名称='" & mobjBill.Pages(mintPage).Details(intRow).Detail.类型 & "'" & strSQL
            If mrs费用类型.EOF Then
                MsgBox """" & mobjBill.Pages(mintPage).Details(intRow).Detail.名称 & """的费用类型为""" & _
                    mobjBill.Pages(mintPage).Details(intRow).Detail.类型 & """,不是" & _
                    IIf(strType = "1", "医保", "公费") & "费用类型！", vbInformation, gstrSysName
                CheckFeeType = False
            End If
        End If
    Else
        For p = 1 To mobjBill.Pages.Count
            For i = 1 To mobjBill.Pages(p).Details.Count
                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张"
                If mobjBill.Pages(p).Details(i).Detail.类型 = "" Then
                    If MsgBox(strTmp & "单据中第 " & i & " 行项目""" & mobjBill.Pages(p).Details(i).Detail.名称 & """的费用类型未设置！" & vbCrLf & "确实要保存单据吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckFeeType = False: Exit For
                    End If
                Else
                    mrs费用类型.Filter = "名称='" & mobjBill.Pages(p).Details(i).Detail.类型 & "'" & strSQL
                    If mrs费用类型.EOF Then
                        If MsgBox(strTmp & "单据中第 " & i & " 行项目""" & mobjBill.Pages(p).Details(i).Detail.名称 & """的费用类型为""" & _
                            mobjBill.Pages(p).Details(i).Detail.类型 & """,不是" & _
                            IIf(strType = "1", "医保", "公费") & "费用类型！" & vbCrLf & "确实要保存单据吗？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            CheckFeeType = False: Exit For
                        End If
                    End If
                End If
            Next
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function ItemExist(lng收费细目ID As Long) As Boolean
    Dim i As Integer, p As Integer
    
    If CheckBillsEmpty Then Exit Function
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).收费细目ID = lng收费细目ID Then
                ItemExist = True: Exit Function
            End If
        Next
    Next
End Function

Private Function CheckExecuteDept(intPage As Long) As Integer
'功能：检查单据中是否有行未输入执行科室
    Dim i As Integer, p As Integer
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).执行部门ID = 0 Then
                intPage = p: CheckExecuteDept = i: Exit Function
            End If
        Next
    Next
End Function
Private Sub InitBalanceGrid(Optional blnOnlyClearBalace As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化保险结算表格
    '入参:blnOnlyBalace-仅清除结算算信息
    '编制:刘兴洪
    '日期:2011-11-02 13:53:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    vsBalance.Clear
    vsBalance.Rows = 4
    If mbytInFun = 0 And InStr("01245", mbytInState) > 0 Then
        ''问题:44425
        'mbytInState As Byte '0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费),4-重新收费;5-异常单据作废
        vsBalance.Width = 2415 * 1.4
        Call picAppend_Resize
    Else
        vsBalance.Width = 2415 * 1.2
        Call picAppend_Resize
    End If
    vsBalance.ColWidth(0) = (vsBalance.Width - 300) * 0.6
    vsBalance.ColWidth(1) = (vsBalance.Width - 300) * 0.4
    vsBalance.ColAlignment(0) = 1
    vsBalance.ColAlignment(1) = 7
    If mbytInState = 3 And mbytInFun = 0 Then vsBalance.Editable = flexEDKbdMouse
    vsBalance.Row = 0
    vsBalance.Col = 1
    vsBalance.TabStop = False
    With vsBalance
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .COLS - 1) = False
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .COLS - 1) = Me.ForeColor
    End With
    For i = 0 To vsBalance.Rows - 1
        vsBalance.RowData(i) = 0
    Next
    If blnOnlyClearBalace Then Exit Sub
    '清除结算集内容
    Set mcolBalance = New Collection
    For i = 1 To tbsBill.Tabs.Count
        mcolBalance.Add Array()
    Next
End Sub
Private Sub ShowPrePayInfo(Optional blnShow As Boolean)
      '65748:取消预交款冲缴权限
    txt预交冲款.Enabled = blnShow And mbytInState = 0  'And InStr(1, mstrPrivs, "预交款冲缴") > 0
    sta.Panels(Pan.C4预交信息).Visible = blnShow
    
    If blnShow Then
        lbl预交冲款.ForeColor = 0
        txt预交冲款.ForeColor = 0
        txt预交冲款.Text = "0.00"
    Else
        lbl预交冲款.ForeColor = &H808080
        txt预交冲款.ForeColor = &H808080
        txt预交冲款.Text = "0.00"
        sta.Panels(Pan.C4预交信息).Tag = ""
        sta.Panels(Pan.C4预交信息).Text = ""
    End If
End Sub

Private Sub ShowPayInfo(Optional blnShow As Boolean)
    txt应缴.Enabled = blnShow
    If blnShow Then
        lbl应缴.ForeColor = 0
        txt应缴.ForeColor = &HFF0000
    Else
        lbl应缴.ForeColor = &H808080
        txt应缴.ForeColor = &H808080
    End If
End Sub

Public Function GetMedicareSum(Optional strItem As String, Optional intPage As Integer, Optional blnOrig As Boolean) As Currency
'功能：获取保险结算的金额
'参数：strItem=是否指定结算方式,否则为所有结算方式
'      intPage=是否指定单据,否则为所有单据
'      blnOrig=是否取原始(最大)结算金额,否则取现在(修改后)有效金额
'说明：该函数以mcolBalance为准计算,对于医保划价收费也是
    Dim arrValue As Variant, curMoney As Currency
    Dim i As Integer, p As Integer
    
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, mcolBalance.Count, intPage)
        For i = 0 To UBound(mcolBalance(p))
            '结算方式;原始(最大)金额;可否修改;有效金额
            arrValue = Split(mcolBalance(p)(i), ";")
            If strItem = "" Or (strItem <> "" And arrValue(0) = strItem) Then
                If blnOrig Then
                    curMoney = curMoney + CCur(arrValue(1))
                Else
                    curMoney = curMoney + CCur(arrValue(3))
                End If
            End If
        Next
    Next
    GetMedicareSum = Format(curMoney, "0.00")
End Function

Private Function GetMedicareStr(intPage As Integer) As String
'功能：返回保险结算方式串,"结算方式|金额||...."
'参数：intPage=指定单据编号
'说明：该函数以mcolBalance为准计算,对于医保划价收费也是
    Dim i As Integer, strTmp As String
    Dim arrValue As Variant
    
    For i = 0 To UBound(mcolBalance(intPage))
        '结算方式;原始(最大)金额;可否修改;有效金额
        arrValue = Split(mcolBalance(intPage)(i), ";")
        strTmp = strTmp & "||" & arrValue(0) & "|" & Format(arrValue(3), "0.00")
    Next
    GetMedicareStr = Mid(strTmp, 3)
End Function

Private Sub SetBalanceVal(intPage As Integer, strItem As String, curVal As Currency)
'功能：设置指定编号单据指定保险结算方式的有效值
'说明：该函数以mcolBalance为准计算,对于医保划价收费也是
'说明：用于正常医保收费修改保险结算金额；及划价单医保收费设置个人帐户等结算金额
    Dim arrValue As Variant, arrPage As Variant
    Dim blnDo As Boolean, strTmp As String, i As Long
        
    arrPage = Array()
    
    If UBound(mcolBalance(intPage)) >= 0 Then
        For i = 0 To UBound(mcolBalance(intPage))
            '结算方式;原始(最大)金额;可否修改;有效金额
            arrValue = Split(mcolBalance(intPage)(i), ";")
            If arrValue(0) = strItem And arrValue(3) <> curVal Then
                blnDo = True
                strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & Format(curVal, "0.00")
            Else
                strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & arrValue(3)
            End If
            ReDim Preserve arrPage(UBound(arrPage) + 1)
            arrPage(UBound(arrPage)) = strTmp
        Next
    Else
        '无内容时强行增加:不支持预结算或医保划价收费时用
        ReDim Preserve arrPage(UBound(arrPage) + 1)
        arrPage(UBound(arrPage)) = strItem & ";" & Format(curVal, "0.00") & ";0;" & Format(curVal, "0.00")
        blnDo = True
    End If
    
    '更新单据结算集(集合不能直接更新)
    If blnDo Then
        mcolBalance.Remove intPage
        If mcolBalance.Count >= intPage Then
            mcolBalance.Add arrPage, , intPage
        Else
            mcolBalance.Add arrPage
        End If
    End If
End Sub

Private Function GetExecDepts(Optional ByVal i As Integer) As String
'功能:获取某单张单据所有的执行部门,不含划价单收费
'参数:i-单据序号,如果i=0,则获取所有单据
    Dim j As Integer, p As Integer, strTmp As String
    
    For p = IIf(i = 0, 1, i) To IIf(i = 0, mobjBill.Pages.Count, i)
        For j = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).NO = "" Then
                If InStr(1, "," & strTmp & ",", "," & mobjBill.Pages(p).Details(j).执行部门ID & ",") <= 0 Then
                    strTmp = strTmp & "," & mobjBill.Pages(p).Details(j).执行部门ID
                End If
            End If
        Next
        If gTy_Module_Para.bln误差占用票据 Then
            If mobjBill.Pages(p).误差金额 <> 0 Then '误差项的执行部门固定为操作员缺省科室,见zl_门诊收费误差_Insert
                If InStr(1, "," & strTmp & ",", "," & UserInfo.部门ID & ",") <= 0 Then
                    strTmp = strTmp & "," & UserInfo.部门ID
                End If
            End If
        End If
    Next
    GetExecDepts = Mid(strTmp, 2)
End Function
Private Function GetInvoiceCount() As Integer
    '功能：计算当前收费需要打印多少张票据
    '说明：共有三级结构
    '   多张单据分别打印--按执行科室分别打印--按收费细目或收据费目打印
    '   如果误差项占用票据行，必须要在这里考虑,因为如果涉及工本费,会因此影工作费的张数
                    
    Dim rsTmp As ADODB.Recordset
    Dim strItems As String, strSQL As String, strNos As String, strTmp As String
    Dim i As Integer, j As Integer, k As Integer, X As Integer, intid As Integer, cur费用行金额 As Currency
    Dim str执行部门IDs As String, lng执行部门ID As Long
    Dim str发票号 As String, int张数 As Integer
    On Error GoTo errH
    
    '计算票据是否充足
    '25187
    If gTy_Module_Para.byt票据分配规则 <> 0 Then
        strNos = ""
        For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO <> "" Then
                    strNos = strNos & "," & mobjBill.Pages(i).NO
                End If
        Next
        If strNos <> "" Then strNos = Mid(strNos, 2)
        If strNos = "" Then GetInvoiceCount = 1: Exit Function
        Call zlExeCuteBillNoSplit(True, 1, mlng领用ID, strNos, 0, txtInvoice.Text, Now, 1, str发票号, int张数)
        If mintInvoicePrint <> 0 Then
            '不打印,不检查
            Call zlCheckFactIsEnough(int张数)
        End If
        GetInvoiceCount = int张数
        Exit Function
    End If
    
    
    If gTy_Module_Para.bln一张票据 Then
        If mobjBill.Pages.Count > 1 And gTy_Module_Para.bln分别打印 And mbytBillSource <> 4 Then
            GetInvoiceCount = mobjBill.Pages.Count
        Else
            GetInvoiceCount = 1
        End If
        Exit Function
    End If
    
    
    If mobjBill.Pages.Count > 1 And gTy_Module_Para.bln分别打印 And mbytBillSource <> 4 Then
        'a.多张分别打印(每张独立)
        For i = 1 To mobjBill.Pages.Count
            'a.a每张按执行科室分别打印
            '------------------------------------------------
            If gTy_Module_Para.byt票据生成方式 >= 10 Then
                'a.a.a直接收费单
                If mobjBill.Pages(i).NO = "" Then
                    str执行部门IDs = GetExecDepts(i)
                    For intid = 0 To UBound(Split(str执行部门IDs, ","))
                        lng执行部门ID = Val(Split(str执行部门IDs, ",")(intid))
                        For j = 1 To mobjBill.Pages(i).Details.Count
                            If Not mobjBill.Pages(i).Details(j).工本费 And mobjBill.Pages(i).Details(j).执行部门ID = lng执行部门ID Then '排开工本费
                                If gTy_Module_Para.byt票据生成方式 = 10 Then
                                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                        If mobjBill.Pages(i).Details(j).InComes(k).实收金额 <> 0 Then '金额不为零
                                            strTmp = mobjBill.Pages(i).Details(j).InComes(k).收据费目
                                            If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                        End If
                                    Next
                                Else
                                    k = k + 1
                                End If
                            End If
                        Next
                        If gTy_Module_Para.bln误差占用票据 And mobjBill.Pages(i).误差金额 <> 0 And lng执行部门ID = UserInfo.部门ID Then
                            If gTy_Module_Para.byt票据生成方式 = 10 Then
                                If InStr("," & strItems & ",", "," & gstr误差收据费目 & ",") = 0 Then strItems = strItems & "," & gstr误差收据费目
                            Else
                                k = k + 1
                            End If
                        End If
                        
                        If gTy_Module_Para.byt票据生成方式 = 10 Then
                            If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt门诊收据行次)
                            strItems = ""
                        Else
                            X = X + IntEx(k / gTy_Module_Para.byt门诊收据行次)
                            k = 0
                        End If
                    Next
                'a.a.b划价单收费
                Else
                    '如果有误差
                    If gTy_Module_Para.bln误差占用票据 And mobjBill.Pages(i).误差金额 <> 0 Then
                        strSQL = "Select count(项目) AS num" & vbNewLine & _
                                "From (Select " & IIf(gTy_Module_Para.byt票据生成方式 = 10, "收据费目", "收费细目id") & " as 项目, 执行部门id" & vbNewLine & _
                                "            From 门诊费用记录" & vbNewLine & _
                                "            Where 记录性质 = 1 And 记录状态 = 0 And Nvl(实收金额, 0) <> 0 And No = [1]" & vbNewLine & _
                                "            Union" & vbNewLine & _
                                "            Select " & IIf(gTy_Module_Para.byt票据生成方式 = 10, "'" & gstr误差收据费目 & "'", glng误差细目ID) & " as 项目," & UserInfo.部门ID & vbNewLine & _
                                "            From Dual)" & vbNewLine & _
                                "Group By 执行部门id"
                    Else
                        strSQL = "Select Count(" & IIf(gTy_Module_Para.byt票据生成方式 = 10, "Distinct 收据费目", "ID") & ") AS num From 门诊费用记录" & _
                            " Where 记录性质=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And NO=[1]" & _
                            " Group by 执行部门id"
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(i).NO)
                    Do While Not rsTmp.EOF
                        X = X + IntEx(rsTmp!Num / gTy_Module_Para.byt门诊收据行次)
                        rsTmp.MoveNext
                    Loop
                End If
                
            'a.b不按执行科室分别打印
            '---------------------------------------------
            Else
                If mobjBill.Pages(i).NO = "" Then
                    For j = 1 To mobjBill.Pages(i).Details.Count
                        If Not mobjBill.Pages(i).Details(j).工本费 Then '排开工本费
                            If gTy_Module_Para.byt票据生成方式 = 0 Then
                                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                    If mobjBill.Pages(i).Details(j).InComes(k).实收金额 <> 0 Then '金额不为零
                                        strTmp = mobjBill.Pages(i).Details(j).InComes(k).收据费目
                                        If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                    End If
                                Next
                            Else
                                k = k + 1
                            End If
                        End If
                        If gTy_Module_Para.bln误差占用票据 And mobjBill.Pages(i).误差金额 <> 0 Then
                            If gTy_Module_Para.byt票据生成方式 = 0 Then
                                If InStr("," & strItems & ",", "," & gstr误差收据费目 & ",") = 0 Then strItems = strItems & "," & gstr误差收据费目
                            Else
                                k = k + 1
                            End If
                        End If
                    Next
                    If gTy_Module_Para.byt票据生成方式 = 0 Then
                        If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt门诊收据行次)
                        strItems = ""
                    Else
                        X = X + IntEx(k / gTy_Module_Para.byt门诊收据行次)
                        k = 0
                    End If
                Else
                    If gTy_Module_Para.bln误差占用票据 And mobjBill.Pages(i).误差金额 <> 0 Then
                        strSQL = "Select count(项目) AS num" & vbNewLine & _
                                "From (Select " & IIf(gTy_Module_Para.byt票据生成方式 = 0, "收据费目", "收费细目id") & " as 项目" & vbNewLine & _
                                "            From 门诊费用记录" & vbNewLine & _
                                "            Where 记录性质 = 1 And 记录状态 = 0 And Nvl(实收金额, 0) <> 0 And No = [1]" & vbNewLine & _
                                "            Union" & vbNewLine & _
                                "            Select " & IIf(gTy_Module_Para.byt票据生成方式 = 0, "'" & gstr误差收据费目 & "'", glng误差细目ID) & " as 项目" & vbNewLine & _
                                "            From Dual)"
                    Else
                        strSQL = "Select Count(" & IIf(gTy_Module_Para.byt票据生成方式 = 0, "Distinct 收据费目", "ID") & ") AS num From 门诊费用记录" & _
                            " Where 记录性质=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And NO=[1]"
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(i).NO)
                    X = X + IntEx(rsTmp!Num / gTy_Module_Para.byt门诊收据行次)
                End If
            End If
        Next
        
    'b.只有一张,或有多张单据一起打印
    '---------------------------------------------------------------------------
    Else
        'b.a按执行科室分别打印
        '----------------------------------------------
        If gTy_Module_Para.byt票据生成方式 >= 10 Then
            str执行部门IDs = GetExecDepts()   '所有单据的执行部门,包含划价单的误差费
            
            '先收集所有的划价单,多张一起打
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO <> "" Then strNos = strNos & ",'" & mobjBill.Pages(i).NO & "'"
            Next
            If strNos <> "" Then
                strNos = Mid(strNos, 2)
                strSQL = "Select Distinct " & IIf(gTy_Module_Para.byt票据生成方式 = 10, "收据费目", "收费细目ID") & " AS 项目,执行部门id From 门诊费用记录" & _
                    " Where 记录性质=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And " & IIf(InStr(1, strNos, ",") > 0, "NO IN(" & strNos & ")", " NO = [1]")
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
                
                Do While Not rsTmp.EOF
                    If InStr(1, "," & str执行部门IDs & ",", "," & rsTmp!执行部门ID & ",") = 0 Then str执行部门IDs = str执行部门IDs & "," & rsTmp!执行部门ID
                    rsTmp.MoveNext
                Loop
                If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst   '后面还要用到
            End If
            
            If InStr(1, str执行部门IDs, ",") = 1 Then str执行部门IDs = Mid(str执行部门IDs, 2)
            
            '再和直接收费单一起处理
            For intid = 0 To UBound(Split(str执行部门IDs, ","))
                lng执行部门ID = Val(Split(str执行部门IDs, ",")(intid))
                For i = 1 To mobjBill.Pages.Count
                    If mobjBill.Pages(i).NO = "" Then
                        For j = 1 To mobjBill.Pages(i).Details.Count
                            If Not mobjBill.Pages(i).Details(j).工本费 And mobjBill.Pages(i).Details(j).执行部门ID = lng执行部门ID Then '排开工本费
                                If gTy_Module_Para.byt票据生成方式 = 10 Then
                                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                        If mobjBill.Pages(i).Details(j).InComes(k).实收金额 <> 0 Then '金额不为零
                                            strTmp = mobjBill.Pages(i).Details(j).InComes(k).收据费目
                                            If strTmp = gstr误差收据费目 Then strTmp = i & "-" & strTmp '加i是为了便于后面的误差费的执行科室处理
                                            If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                        End If
                                    Next
                                Else    '数量为零的行在保存前已检查禁止继续
                                    strTmp = mobjBill.Pages(i).Details(j).收费细目ID
                                    If strTmp = glng误差细目ID Then strTmp = i & "-" & strTmp
                                    If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                End If
                            End If
                        Next
                    End If
                    
                    '划价单和直接收费的误差一起处理
                    '多张一起打,如果有误差,误差项是每张单据一个,所以要加i以区别
                    If gTy_Module_Para.bln误差占用票据 And mobjBill.Pages(i).误差金额 <> 0 And lng执行部门ID = UserInfo.部门ID Then
                        If gTy_Module_Para.byt票据生成方式 = 10 Then
                            If InStr("," & strItems & ",", "," & i & "-" & gstr误差收据费目 & ",") = 0 Then strItems = strItems & "," & i & "-" & gstr误差收据费目
                        Else
                            If InStr("," & strItems & ",", "," & i & "-" & glng误差细目ID & ",") = 0 Then strItems = strItems & "," & i & "-" & glng误差细目ID
                        End If
                    End If
                Next
                
                '再处理所有的收费划价单
                If strNos <> "" And Not rsTmp Is Nothing Then
                    rsTmp.Filter = "执行部门id=" & lng执行部门ID
                    For k = 1 To rsTmp.RecordCount
                        If InStr("," & strItems & ",", "," & rsTmp!项目 & ",") = 0 Then strItems = strItems & "," & rsTmp!项目
                        rsTmp.MoveNext
                    Next
                End If
                
                '划价收费单与直接收费单可能混用,所以需要,最后处理
                If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt门诊收据行次)
                strItems = ""
            Next
            
            
        'b.b不按执行科室分别打印
        '-----------------------------------------------------
        Else
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO = "" Then
                    For j = 1 To mobjBill.Pages(i).Details.Count
                        If Not mobjBill.Pages(i).Details(j).工本费 Then '排开工本费
                            If gTy_Module_Para.byt票据生成方式 = 0 Then
                                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                    If mobjBill.Pages(i).Details(j).InComes(k).实收金额 <> 0 Then '金额不为零
                                        strTmp = mobjBill.Pages(i).Details(j).InComes(k).收据费目
                                        If strTmp = gstr误差收据费目 Then strTmp = i & "-" & strTmp
                                        If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                    End If
                                Next
                            Else    '数量为零的行在保存前已检查禁止继续
                                strTmp = mobjBill.Pages(i).Details(j).收费细目ID
                                If strTmp = glng误差细目ID Then strTmp = i & "-" & strTmp
                                If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                            End If
                        End If
                    Next
                Else
                    strNos = strNos & ",'" & mobjBill.Pages(i).NO & "'"
                End If
                '划价单和直接收费的误差一起处理
                '多张一起打,如果有误差,误差项是每张单据一个,所以要加i以区别
                If gTy_Module_Para.bln误差占用票据 And mobjBill.Pages(i).误差金额 <> 0 Then
                    If gTy_Module_Para.byt票据生成方式 = 0 Then
                        If InStr("," & strItems & ",", "," & i & "-" & gstr误差收据费目 & ",") = 0 Then strItems = strItems & "," & i & "-" & gstr误差收据费目
                    Else
                        If InStr("," & strItems & ",", "," & i & "-" & glng误差细目ID & ",") = 0 Then strItems = strItems & "," & i & "-" & glng误差细目ID
                    End If
                End If
            Next
            If strNos <> "" Then
                strNos = Mid(strNos, 2)
                strSQL = "Select Distinct " & IIf(gTy_Module_Para.byt票据生成方式 = 0, "收据费目", "收费细目ID") & " AS 项目,执行部门id From 门诊费用记录" & _
                    " Where 记录性质=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And " & IIf(InStr(1, strNos, ",") > 0, "NO IN(" & strNos & ")", " NO = [1]")
                
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
                For k = 1 To rsTmp.RecordCount
                    If InStr("," & strItems & ",", "," & rsTmp!项目 & ",") = 0 Then strItems = strItems & "," & rsTmp!项目
                    rsTmp.MoveNext
                Next
            End If
            ''划价收费单与直接收费单可能混用,所以需要,最后处理
            X = IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt门诊收据行次)
        End If
    End If
    GetInvoiceCount = X
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetBillSumByDB(strNo As String) As Currency
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
        strSQL = "Select SUM(实收金额) AS 实收金额 From 门诊费用记录 " & _
                " Where 记录性质=1 And 记录状态=0 And NO=[1] And 操作员姓名 is Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        If Not rsTmp.EOF Then
            GetBillSumByDB = Nvl(rsTmp!实收金额, 0)
        Else
            GetBillSumByDB = 0
        End If
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
End Function

Private Function Check门诊结算作废(ByVal lng结帐ID As Long, ByVal intInsure As Long) As Boolean
'功能：根据指定结帐ID的医保结算方式名称判断是否全部支持门诊结算作废
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    strSQL = "Select B.结算方式 From 病人预交记录 B,结算方式 C" & _
        " Where B.记录性质=3 And B.结算方式=C.名称 And Nvl(C.性质,1) IN(3,4)  And B.结帐ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
        
    If rsTmp.RecordCount > 0 Then
       Check门诊结算作废 = True
       For i = 1 To rsTmp.RecordCount
           If Not gclsInsure.GetCapability(support门诊结算作废, , intInsure, rsTmp!结算方式) Then
                MsgBox "医保结算方式:" & rsTmp!结算方式 & ",不支持门诊结算作废!" & vbCrLf & _
                "修改单据要求所有结算方式全部作废后重新产生新单据,所以不能修改此单据!", vbInformation, gstrSysName
                Check门诊结算作废 = False
                Exit For
           End If
       Next
    Else
        MsgBox "数据发生异常,读取单据所使用的医保结算方式失败,不允修改此单据！", vbInformation, gstrSysName
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowRegist()
'功能：检查是否可以显示挂号按钮
    Dim strPrivs As String
    On Error GoTo errH
    
    If mbytInFun = 0 And mbytInState = 0 Then
        strPrivs = GetPrivFunc(glngSys, 1111)
        If InStr(";" & strPrivs & ";", ";挂号;") > 0 Then '功能是否授权
            cmdRegist.Visible = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ShowIDCard()
'功能：检查是否可以显示就诊卡按钮
    Dim strPrivs As String
    On Error GoTo errH
    
    If mbytInFun = 0 And mbytInState = 0 Then
        strPrivs = GetPrivFunc(glngSys, 1107)
        If InStr(";" & strPrivs & ";", ";发卡;") > 0 Then '功能是否授权
            cmdIDCard.Visible = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetOperatorInfo(ByVal str姓名 As String, Optional bln护士 As Boolean, Optional int职务 As Integer) As Boolean
'功能：获取指定姓名开单人(医生或护士)的性质或职务
'返回：int职务:0-未设置；bln护士:是否只是护士
'说明：以前是直接读取marrDr中的内容,改为多单据多开单人后一些地方不行
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    bln护士 = False: int职务 = 0
    If Not mrs开单人 Is Nothing Then
        mrs开单人.Filter = "姓名='" & str姓名 & "' " & IIf(gbln护士, "", " And 人员性质<>'护士'")
        If mrs开单人.RecordCount > 0 Then
            int职务 = mrs开单人!职务
            strSQL = mrs开单人!人员性质
            If strSQL = "护士" Then bln护士 = True
            If strSQL = "医生" Then bln护士 = False
        End If
    Else
        strSQL = _
            " Select Nvl(A.聘任技术职务,0) as 职务,B.人员性质 From 人员表 A,人员性质说明 B" & _
            " Where A.ID=B.人员ID And B.人员性质 IN('医生','护士') And A.姓名=[1] And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str姓名)
        If Not rsTmp.EOF Then
            int职务 = rsTmp!职务
            Do While Not rsTmp.EOF
                If rsTmp!人员性质 = "护士" Then bln护士 = True
                If rsTmp!人员性质 = "医生" Then bln护士 = False: Exit Do
                rsTmp.MoveNext
            Loop
        End If
    End If
    GetOperatorInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function YBIdentifyCancel() As Boolean
'功能：取消医保病人身份验证
'返回：返回假时不退出界面或清除操作
    Dim lng病人ID As Long
    YBIdentifyCancel = True
    If mbytInFun = 0 And mbytInState = 0 Then
        If mstrYBPati <> "" And txtPatient.Text <> "" Then
            If UBound(Split(mstrYBPati, ";")) >= 8 Then
                If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
                    lng病人ID = Val(CLng(Split(mstrYBPati, ";")(8)))
                End If
            End If
            If lng病人ID <> 0 Then
                YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng病人ID, mintInsure)
            End If
        End If
    End If
End Function

Private Function SelectPatient() As Long
'功能：显示合约单位病人并选择
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    
    strSQL = _
        "Select 0 As 末级, ID, 上级id, -null As 门诊号, '[' || 编码 || ']' || 名称 As 姓名, Null As 性别, Null As 年龄," & vbNewLine & _
        "       Null As 费别, Null As 付款方式, Null As 出生日期, Null As 家庭地址" & vbNewLine & _
        "From 合约单位" & vbNewLine & _
        "Where 撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')" & vbNewLine & _
        "Start With 上级id Is Null" & vbNewLine & _
        "Connect By Prior ID = 上级id" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select 1 As 末级, A.病人id As ID, A.合同单位id As 上级id, A.门诊号, A.姓名, A.性别, A.年龄, A.费别," & vbNewLine & _
        "       A.医疗付款方式 As 付款方式, To_Char(A.出生日期, 'YYYY-MM-DD') As 出生日期, A.家庭地址" & vbNewLine & _
        "From 病人信息 A, 合约单位 B" & vbNewLine & _
        "Where B.ID = A.合同单位id And A.停用时间 Is Null And A.当前科室id Is Null And A.合同单位id Is Not Null And B.末级 = 1"
            
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "合约单位病人", , , "请先选择合约单位，再选择病人", , , , , , , blnCancel, , , 1)
    If rsTmp Is Nothing Then Exit Function
    SelectPatient = rsTmp!ID
End Function

Private Sub SetBillRowForeColor(ByVal lngRow As Long, ByVal lngColor As Long)
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnActive As Boolean, blnRedraw As Boolean
    Dim i As Long
    
    '备份属性
    mblnEnterCell = False
    lngPreRow = Bill.Row: lngPreCol = Bill.Col
    blnActive = Bill.Active: blnRedraw = Bill.MsfObj.Redraw
        
    '开始设置
    Bill.Active = False: Bill.Redraw = False
    Bill.Row = lngRow
    For i = Bill.MsfObj.FixedCols To Bill.COLS - 1
        Bill.Col = i: Bill.MsfObj.CellForeColor = lngColor
    Next
    
    '恢复属性
    Bill.Row = lngPreRow: Bill.Col = lngPreCol
    Bill.Active = blnActive: Bill.Redraw = blnRedraw
    mblnEnterCell = True
End Sub

Private Sub SetItemRowColor(ByVal intPage As Integer, ByVal lngRow As Long)
'功能：根据药品/材料的储备限额设置行颜色提示
    If mobjBill.Pages(intPage).Details.Count >= lngRow And InStr(",0,1,", mbytInFun) > 0 And mbytInState = 0 Then
        With mobjBill.Pages(intPage).Details(lngRow)
            If mbln储备限额检查 And (InStr(",5,6,7,", .收费类别) > 0 Or (.收费类别 = "4" And .Detail.跟踪在用)) Then
                If ItemUnderSet(.收费类别, .收费细目ID, .执行部门ID, IIf(gbln药房单位, .Detail.药房包装, 1) * .Detail.库存) Then
                    Call SetBillRowForeColor(lngRow, &HC000C0)
                Else
                    Call SetBillRowForeColor(lngRow, Bill.ForeColor)
                End If
            Else
                Call SetBillRowForeColor(lngRow, Bill.ForeColor)
            End If
        End With
    End If
End Sub

Private Function CheckSaveMultiPrice() As Boolean
    Dim p As Integer
    
    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And chkCancel.Value = 0 Then
        For p = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(p).NO <> "" Then
                Exit Function
            End If
        Next
        CheckSaveMultiPrice = True  '允许保存为划价单
    Else
        Exit Function
    End If
End Function

Private Sub MergeRepeatItem()
'功能：合并单据中所有重复输入的分批/时价药品/卫生材料(项目，执行科室相同)
'说明：调用之前应已确保不存在中药需要合并但付数不同的情况
    Dim i As Integer, j As Integer
    Dim m As Integer, n As Integer
    Dim objDetail As BillDetail
    Dim rsItem As New ADODB.Recordset
    Dim blnRefresh As Boolean
    
    rsItem.Fields.Append "Type", adBigInt
    rsItem.Fields.Append "Page", adBigInt
    rsItem.Fields.Append "Row", adBigInt
    rsItem.CursorLocation = adUseClient
    rsItem.LockType = adLockOptimistic
    rsItem.CursorType = adOpenStatic
    rsItem.Open
        
    For i = 1 To mobjBill.Pages.Count
        For j = 1 To mobjBill.Pages(i).Details.Count
            With mobjBill.Pages(i).Details(j)
                If (.Detail.分批 Or .Detail.变价) And .数次 * .付数 <> 0 _
                    And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                    For m = i To mobjBill.Pages.Count
                        For n = IIf(m = i, j + 1, 1) To mobjBill.Pages(m).Details.Count
                            Set objDetail = mobjBill.Pages(m).Details(n)
                            If objDetail.收费细目ID = .收费细目ID _
                                And objDetail.执行部门ID = .执行部门ID And objDetail.付数 * objDetail.数次 <> 0 Then
                                .数次 = .数次 + objDetail.数次
                                objDetail.数次 = 0
                                                                
                                rsItem.AddNew
                                rsItem!Type = 1 '合并到哪行
                                rsItem!Page = i
                                rsItem!Row = j
                                rsItem.Update
                                                                
                                rsItem.AddNew
                                rsItem!Type = 2 '被合并的行
                                rsItem!Page = m
                                rsItem!Row = n
                                rsItem.Update
                            End If
                        Next
                    Next
                End If
            End With
        Next
    Next
    
    If rsItem.RecordCount > 0 Then
        '删除被合并的行(反序)
        rsItem.Sort = "Page,Row Desc"
        rsItem.Filter = "Type=2"
        Do While Not rsItem.EOF
            Call DeleteDetail(rsItem!Row, rsItem!Page)
            If rsItem!Page = mintPage Then blnRefresh = True
            rsItem.MoveNext
        Loop
        
        '重算合并到的行
        For i = 1 To mobjBill.Pages.Count
            rsItem.Filter = "Type=1 And Page=" & i
            If rsItem.RecordCount > 1 Then          '一张单据有几组合并时,删除行号后,之前记录的合并到的行号可能变了
                Call CalcMoneys(i)
            ElseIf rsItem.RecordCount = 1 Then
                Call CalcMoneys(rsItem!Page, rsItem!Row)
            End If
            If i = mintPage Then blnRefresh = True
        Next
    End If
    
    If blnRefresh Then
        Call ShowDetails
    End If
    Call ShowMoney
    
    '需要重新预结算
    If cmd预结算.Visible Then
        Call InitBalanceGrid
        cmd预结算.TabStop = True
        cmdOK.Enabled = False
    End If
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
    Dim strTemp As String, p As Long
    If intInsure = 0 Then zlCheck北京医保 = True: Exit Function
    
    Err = 0: On Error GoTo Errhand:
    '刘兴洪:???
    '入口参数：
    'mbytInFun As Byte '0-收费,1-划价,2-门诊记帐
    'mbytInState As Byte '0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费)
    'mstrInNO As String '操作的单据号(查看，调整，修改，退费，销帐)
    'mbytBilling As Byte 'mbytInFun=2时：0-正常记帐,1-记帐划价,2-记帐审核

    
    '只有划价才支持检查
    If mbytInFun <> 1 And mbytInState <> 0 Or MCPAR.医生确定处方类型 = False Then
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
    strTemp = zlCommFun.ShowMsgbox("处方类型", "请确定当前医保病人本次要发送的药品处方的类型。", "!医保内(&A),医保外(&B),?取消(&C)", Me)
    If strTemp = "" Or strTemp = "取消" Then Exit Function
    '如果是补门诊收费划价单，且是医保病人，则当医保参数”support医生确定处方类型”有效时，保存时提示该单据是”医保内，医保外”，如果是医保内费用记录摘要中存放1，医保外存放2。
    strTemp = IIf(strTemp = "医保内", 1, 2)
    
    '--更新摘要
    '对每张单据独立执行保存
    For p = 1 To mobjBill.Pages.Count
        '产生每张收费单据的单据号
        If mobjBill.Pages(p).NO = "" Then
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                mobjBillDetail.摘要 = strTemp
            Next
        End If
    Next
    zlCheck北京医保 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function Get可刷金额() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡的可刷金额
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-01-08 14:53:45
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, j As Long, intCol As Integer
    Dim dbl可刷金额 As Double, dbl总额 As Double
    
    With mobjBill
        For i = 1 To mobjBill.Pages.Count
            dbl可刷金额 = dbl可刷金额 + .Pages(i).实收金额 - .Pages(i).保险金额 - .Pages(i).冲预交额            ' + .Pages(i).误差金额:
            dbl总额 = dbl总额 + 0 + .Pages(i).实收金额
        Next
    End With
  
    '如果没有,再尝试从表格中取(仅一张单据时)
    If dbl总额 = 0 And tbsBill.Tabs.Count = 1 _
        And Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.项目) = "") Then
        
        intCol = BillCol.实收金额
        For i = 1 To Bill.Rows - 1
            If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                dbl总额 = dbl总额 + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
            End If
        Next
        dbl可刷金额 = dbl总额 - Val(txt预交冲款.Text)
    End If
    Get可刷金额 = Format(dbl可刷金额, gstrDec)
End Function
  
Private Sub ShowStatusCargoSpace(ByVal lng收费细目ID As Long, lng执行库房ID As Long, _
    Optional bln卫材 As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：显示库房货位
    '编制：刘兴洪
    '日期：2010-04-13 14:30:20
    '说明：27505(随后调整成公用函数)
    '         目前只针对划价单
    '------------------------------------------------------------------------------------------------------------------------
    Static lngPre收费细目ID As Long
    Static lngPre执行库房ID As Long
    Static strCargo_Space As String  '上次货位
    Dim strTemp As String
    Err = 0: On Error GoTo Errhand:
    '划价时要显示库房货位
    If mbytInFun <> 1 Then Exit Sub
    If Not (lngPre收费细目ID = lng收费细目ID And lng执行库房ID = lngPre执行库房ID) Then
        lngPre收费细目ID = lng收费细目ID: lngPre执行库房ID = lng执行库房ID
        strCargo_Space = GetPlace(lng收费细目ID, lng执行库房ID, bln卫材)     '重新获取库房货位
    End If
    If strCargo_Space <> "" And InStr(1, strCargo_Space, "货位:") = 0 Then strCargo_Space = "货位:" & strCargo_Space
    strTemp = Split(sta.Panels(Pan.C2提示信息), ",货位:")(0)
    strTemp = Split(strTemp, "货位:")(0)
    If strTemp <> "" And strCargo_Space <> "" Then strTemp = strTemp & ","
    strTemp = strTemp & strCargo_Space
    sta.Panels(Pan.C2提示信息) = strTemp    '显示出货位
Errhand:
End Sub

Public Function zl获取中药形态(Optional ByVal intPage As Integer = 0, Optional ByVal lngRow As Long = -1, Optional blnOnly中成药 As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据是否录入了中草药的
    '入参:intPage-当前第几页
    '     blnOnly中成药-仅判断是否有中成药(对配方时判断有效):原因是中划药在配方中已经存在,就不需要检查
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
    If intPage = 0 Then intPage = mintPage
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly中成药, ",6,", ",6,7,")
    With mobjBill.Pages(intPage).Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).收费类别 & ",") > 0 And .Item(i).收费细目ID <> 0 And i <> lngRow Then
                zl获取中药形态 = .Item(i).Detail.中药形态
                Exit Function
            End If
        Next
    End With
End Function
Private Sub SetBill中草药EditEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置中草药的编辑状态
    '编制：刘兴洪
    '日期：2010-08-06 10:58:45
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With Bill
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "项目" Then
                .ColData(i) = 0
            Else
                .ColData(i) = 5
            End If
        Next
    End With
End Sub
'''Private Sub txt找补_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If lbl找补.Caption <> "找补" Then
'''        zlCommFun.ShowTipInfo txt找补.hWnd, mstr应付款结算方式, False
'''    Else
'''        zlCommFun.ShowTipInfo txt找补.hWnd, "", False
'''    End If
'''End Sub
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
Private Sub zlCheckFactIsEnough(Optional ByVal intInvoicePages As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    ' 入参:intInvoicePages-需要的发票张数,如果为0,按系统参数提醒
    '编制:刘兴洪
    '日期:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long, lngNums As Long
    
    If Not (mbytInFun = 0 And mbytInState = 0) Then Exit Sub
    '刘兴洪 问题:26948 日期:2009-12-28 17:43:00
    '需要检查剩余数量是否充足:
    If intInvoicePages <> 0 Then
        If zlCheckInvoiceOverplusEnough(1, intInvoicePages, lng剩余数量, mlng领用ID, mstrUseType) = False Then
            MsgBox "注意:" & vbCrLf & _
                   "    当前剩余票据不足(" & lng剩余数量 & ") ,当前需要" & intInvoicePages & "张票据,请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    Else
        If zlCheckInvoiceOverplusEnough(1, gTy_Module_Para.int提醒剩余票据张数, lng剩余数量, mlng领用ID, mstrUseType) = False Then
            MsgBox "注意:" & vbCrLf & _
                   "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & gTy_Module_Para.int提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    End If
End Sub
Private Function zlCheckBill存在非散装草药(intPage As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中存在非散装草药形态
    '入参:intPage-指定的页
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-26 10:19:46
    '问题:38328
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mobjBill Is Nothing Then Exit Function
    With mobjBill.Pages(intPage)
        If .Details.Count = 0 Then Exit Function
        For i = 1 To .Details.Count
            If .Details(i).收费类别 = "7" Then
                If .Details(i).Detail.中药形态 <> 0 Then    '0-散装;1-中药饮片;2-免煎剂
                    zlCheckBill存在非散装草药 = True: Exit Function
                End If
            End If
        Next
    End With
End Function

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytInState = 1 Then Exit Sub
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, _
        gobjSquare.objSquareCard, "", txtPatient)
        
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
Private Function zlCheckDelValied(ByVal lng卡类别ID As Long, _
     ByVal strName As String, _
     ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
     ByVal strSwapNO As String, strSwapMemo As String, _
     ByRef lng结帐ID As Long, _
    ByVal dbl退款金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费交易接口
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    If lng卡类别ID = 0 Then zlCheckDelValied = True: Exit Function
    If Not gobjSquare.objSquareCard Is Nothing Then
        Call CreateSquareCardObject(gfrmMain, mlngModul)
        Call initCardSquareData
    End If
    
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "注意:" & vbCrLf & _
                     "      当前收费是按" & strName & " 收费的,但不存在操作的相关部件,不能退款,请与系统管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户回退交易前的检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退款合法,返回true,否则返回Flase
      If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng卡类别ID, bln消费卡, strCardNo, _
        "3|" & lng结帐ID, dbl退款金额, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function
Private Function CheckBrushCard(ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal dbl退费额 As Double, ByRef strBrushCardNo As String, ByRef strbrPassWord As String, Optional ByRef bln退现 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset
    On Error GoTo errHandle
    Dim dblMoney As Double
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean) As Boolean
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng卡类别ID, bln消费卡, Trim(txtPatient.Text), cboSex.Text, txt年龄.Text & cbo年龄单位.Text, dbl退费额, strBrushCardNo, strbrPassWord, _
        True, True, bln退现) = False Then Exit Function
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function CallBackBalanceInterface(ByVal cllBalance As Collection, _
    ByVal lng结帐ID As Long, ByVal lng冲销ID As Long, _
    ByVal strNo As String, _
    ByVal dblMoney As Double, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, strSwapGlideNO As String, strSwapMemo As String, str结算信息 As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim bln消费卡 As Boolean, lng卡类别ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
    If cllBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    
    bln消费卡 = Val(cllBalance(1)(2)) = 1
    lng卡类别ID = cllBalance(1)(0)
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str卡号 = cllBalance(1)(1)
    strSwapGlideNO = cllBalance(1)(3)
    strSwapMemo = cllBalance(1)(4)
    If lng结帐ID <> 0 Then str结算信息 = str结算信息 & "||3|" & lng结帐ID
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
'    strSQL = "Select 结帐ID From 门诊费用记录  Where 记录性质=1 And NO =[1] and 记录状态=2"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
'    If rsTemp.EOF Then
'        strErrMsg = "未找第三方结算交易信息，退费失败": Exit Function
'    End If
'    lng冲销ID = Val(Nvl(rsTemp!结帐ID))
    
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    strSwapExtendInfor = "3|" & lng冲销ID: strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-传入，本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strSwapExtendInfor-传出，交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, lng卡类别ID, bln消费卡, str卡号, str结算信息, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, lng冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, lng冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
    
End Function

Private Sub SaveThreeData(ByVal cllThree As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方数据
    '编制:刘兴洪
    '日期:2011-08-18 17:53:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllThree, Me.Caption
Errhand:
    Err = 0: On Error GoTo 0
End Sub

Private Function LoadErrBillCharge(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载错误的收费票据,进入重新收费
    '入参:strNo-错误的收费单据号
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 16:14:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsNos As ADODB.Recordset, strSQL As String
    Dim objPage As New BillPage
    Dim arrBills As Variant, strBills As String
    Dim blnRead As Boolean, i As Long, k As Long
    Dim lng结帐ID As Long, blnMulitNos As Boolean '多单据
    Dim lng病人ID As Long, lngRow As Long
    
    If Not (mbytInFun = 0 And (mbytInState = 4 Or mbytInState = 5 Or mblnErrBill)) Then LoadErrBillCharge = True: Exit Function
     
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "" & _
    "   Select    A.NO, A.病人ID,A.结帐ID,max(B.结算序号) as 结算序号   " & _
    "   From 门诊费用记录 A,病人预交记录 B , " & _
    "           (Select max(j.结算序号) as 结算序号 From 门诊费用记录 m,病人预交记录 j  Where m.记录性质=1 and m.记录状态=[2] and m.结帐ID=j.结帐ID And   m.NO=[1]) I" & _
    "   Where  A.结帐ID=B.结帐ID  " & _
    "           And B.结算序号=I.结算序号 And A.记录状态=[2]  " & _
    "   Group by A.NO, A.病人ID,A.结帐ID" & _
    "   Order by A.结帐ID " & IIf(mbln退费异常, " Desc ", "") & ",A.NO"
    
    Set rsNos = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, IIf(mbln退费异常, 2, 1))
    If rsNos.RecordCount = 0 Then Exit Function
    With rsNos
        '检查是否存在未结的医保数据
        blnMulitNos = .RecordCount > 1
        mlng结算序号 = Val(Nvl(!结算序号))
    End With
    mblnDelete = mbln退费异常
    '57682
    strSQL = "" & _
    "   Select  decode(A.记录性质,1,'预存款',11,'预存款',A.结算方式) as 结算方式, " & _
    "            sum(nvl(A.冲预交,0)) as 结算金额 " & _
    "   From 病人预交记录 A" & _
    "   where A.结算序号=[1]  And  A.记录状态=[2]  " & _
    "   Group by decode(A.记录性质,1,'预存款',11,'预存款',A.结算方式)" & _
    "   Order by 结算方式"
    
    '异常单据的结算方式(不含预交款)
    Set mrsErrBlance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结算序号, IIf(mbln退费异常, 2, 1))
    LoadErrBillCharge = True
    
    '清除现有单据的内容
    '---------------------------------------------------------------------
     txtModi.Text = ""
    Call ClearTotalInfo
    Call ClearPayInfo
    Call ClearBillRows
        
    '预结算支持时才清除,否则会自动算
    If cmd预结算.Visible Then Call InitBalanceGrid
    Call ClearMoney
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '多单据收费:只保留一页对象
    For i = mobjBill.Pages.Count To 1 Step -1
        mobjBill.Pages.Remove i
    Next
    mobjBill.Pages.Add objPage.Details
    
    '多单据收费:恢复缺省单据页卡
    mintPage = 1
    For i = tbsBill.Tabs.Count To 1 Step -1
        tbsBill.Tabs(i).Tag = ""
        If i <> 1 Then tbsBill.Tabs.Remove i
    Next
        
    '读取显示每张划价单
    '---------------------------------------------------------------------
    mblnNOMoved = False  '不从后备表中读取
    k = 1: i = 0
    mblnDoing = False '表明正在自动读
    tbsBill.Visible = blnMulitNos
    cmdAddBill.Visible = blnMulitNos
    cmdDelBill.Visible = blnMulitNos
    cmdAddBill.Enabled = True
    fraBill.Visible = blnMulitNos
    Form_Resize
    mintInsure = 0
    Do While Not rsNos.EOF
        If mintInsure = 0 Then
                mintInsure = ChargeExistInsure(Nvl(rsNos!NO), lng病人ID, lng结帐ID)
                If mintInsure <> 0 Then Call initInsurePara(lng病人ID)
        End If
        Me.Refresh
        '增加单据页标签(同cmdAdd_Click内容)
        '-----------------------------------------------------------------------
        If k > 1 And mobjBill.Pages(mobjBill.Pages.Count).NO <> "" Then
            If tbsBill.Tabs.Count >= 10 Then
                Call tbsBill.Tabs.Add(, , "单据" & tbsBill.Tabs.Count + 1)
            Else
                If tbsBill.Tabs.Count + 1 = 10 Then
                    Call tbsBill.Tabs.Add(, , "单据1&0")
                Else
                    Call tbsBill.Tabs.Add(, , "单据&" & tbsBill.Tabs.Count + 1)
                End If
            End If
            
            '加入单据页对象:即使是划价收费也保持一致
            mobjBill.Pages.Add objPage.Details
            '加入结算集合:划价收费也要保持一致
            mcolBalance.Add Array()
            '多张单据时禁止修改及退费功能
            txtModi.Enabled = False
            chkCancel.Enabled = False
            cmdDelete.Enabled = False
            '激活Click,显示新增加单据的内容(空白)
            tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
        End If
                
        '读取划价单据内容(同cboNO_KeyPress)
        '----------------------------------------------------------------------
        strNo = Nvl(rsNos!NO)
        blnRead = ReadBill(strNo, 0, False, , , True)
        mobjBill.Pages(k).结帐ID = Val(Nvl(rsNos!结帐ID))
        If blnRead Then k = k + 1: cboNO.Text = strNo
        i = i + 1
        rsNos.MoveNext
    Loop
    Dim blnFind As Boolean
    '加载结算方式
    mrsErrBlance.Filter = 0
    With mrsErrBlance
        If mrsErrBlance.RecordCount <> 0 Then mrsErrBlance.MoveFirst
        vsBalance.Clear
        vsBalance.Rows = 1
        i = 1
        Do While Not .EOF
            lngRow = 0
            blnFind = False
            For i = 0 To vsBalance.Rows - 1
                If vsBalance.TextMatrix(i, 0) = Nvl(!结算方式, " ") Then
                    blnFind = True
                    lngRow = i: Exit For
                End If
            Next
            If Not blnFind And vsBalance.TextMatrix(lngRow, 0) <> "" Then
                vsBalance.Rows = vsBalance.Rows + 1
                lngRow = vsBalance.Rows - 1
            End If
             vsBalance.TextMatrix(lngRow, 0) = Nvl(!结算方式, " ")
             vsBalance.TextMatrix(lngRow, 1) = Format(Val(Nvl(!结算金额)) + Val(vsBalance.TextMatrix(lngRow, 1)), "0.00")
            .MoveNext
        Loop
    End With
    txtInvoice.Text = ""
    Call ReInitPatiInvoice(True, mintInsure, lng病人ID)
    Bill.Active = False
    chk加班.Enabled = False
    'txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    cmdDelBill.Enabled = False
    cmdAddBill.Enabled = False
    mblnDoing = False '表明自动读取完毕
    Call ShowMoney
    '显示摘要
    Call Bill_EnterCell(1, BillCol.项目)
    cmdOK.Enabled = True: cmdOK.Visible = True
    If cmdOK.Enabled Then cmdOK.SetFocus
    LoadErrBillCharge = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub PrintBill(ByVal strNos, strModiNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印
    '编制:刘兴洪
    '日期:2011-08-26 18:38:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    Dim strReclaimInvoice As String '回收的发票号
    
    If mbytInFun = 1 Or mblnSaveAsPrice Then  '打印划价通知单
        If gint划价通知单 = 1 Then
           Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & mobjBill.NO, 2)
        ElseIf gint划价通知单 = 2 Then
            If MsgBox("要打印划价通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & mobjBill.NO, 2)
            End If
        End If
        Exit Sub
    End If
    
    If mbytInFun = 2 Then   '记帐单打印
        If mbytBilling = 0 And gbln记帐打印 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & mobjBill.NO, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
        ElseIf mbytBilling = 1 And gbln划价打印 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & mobjBill.NO, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
        End If
        Exit Sub
    End If
    If mbytInFun <> 0 Then Exit Sub
    If mstrYBPati <> "" And MCPAR.门诊连续收费 Then
        '医保连续收费模式时，确定时不打印，等同一病人的几张单据确定完后，按[完成收费]按钮一起打印。
        '医保连续收费时不支持多单据,取一个就行了
        mstrYBBill = mstrYBBill & "," & mobjBill.NO
        Exit Sub
    End If
    
   '打印门诊收据
     '问题:34941:And Not (MCPAR.多单据一次结算 And mstrYBPati <> "")
     Dim blnPrintBillEmpty As Boolean   '55052
    If mblnPrint And Not (MCPAR.医保接口打印票据 And mstrYBPati <> "") Then
        '问题:42708
        If Format(mobjBill.登记时间, "yyyy") < 2000 Then mobjBill.登记时间 = zlDatabase.Currentdate
        '问题:44322
RePrint:
        strReclaimInvoice = ""
        Call frmPrint.ReportPrint(1, strNos, strModiNos, strReclaimInvoice, mlng领用ID, mlngShareUseID, txtInvoice.Text, mobjBill.登记时间, CStr(mdbl缴款), CStr(mdbl找补), _
            gTy_Module_Para.bln分别打印 And mbytBillSource <> 4, mintInvoiceFormat, , , mstrUseType, blnPrintBillEmpty, , , mstr普通价格等级)
        If gblnStrictCtrl And blnPrintBillEmpty = False Then
            If zlIsNotSucceedPrintBill(1, strNos, strNotValiedNos) = True Then
                    If MsgBox("单据[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
            End If
        End If
    End If
    '打印费用清单:固定不分别打印
    If InStr(mstrPrivs, "打印清单") > 0 Then
        If gint收费清单 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
        ElseIf gint收费清单 = 2 Then
            If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            End If
        End If
    End If
End Sub

Public Function ChargeIsErrBill(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否异常单据
    '返回:是异常单据,返回升True,否则返回False
    '编制:刘兴洪
    '日期:2011-08-28 11:32:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select /*+cardinality(j,10)*/ 1" & vbNewLine & _
            " From 门诊费用记录 A, Table(f_Str2list([1])) J" & vbNewLine & _
            " Where a.记录性质 = 1 And Nvl(a.费用状态, 0) = 1 And a.No = j.Column_Value And Rownum < 2"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否异常单据", Replace(strNos, "'", ""))
    ChargeIsErrBill = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function PatiErrBillPay(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人,对异常单据进行收费
    '返回:存在异常单据,并进行收费,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-10-23 21:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNo As String, lng结帐ID As Long, lng结算序号 As Long
    Dim str操作员姓名 As String
    '问题:44271
    If (mbytInState = 1 Or mbytInState = 2 Or mbytInState = 3) Or mbytInFun <> 0 Then Exit Function
    If mbytInFun = 0 And mbytInState = 0 And mstrInNO <> "" Then PatiErrBillPay = False: Exit Function
   
    On Error GoTo errHandle
    
    mblnErrBill = False
    strSQL = "" & _
    "   Select  A.NO,A.结帐ID,A.操作员姓名" & _
    "   From 门诊费用记录 A" & _
    "   Where A.病人ID=[1] and nvl(A.费用状态,0)=1  " & _
    "               And A.记录状态=1 and A.记录性质=1 " & _
    "               And Rownum=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsTemp.EOF Then Exit Function
    strNo = Nvl(rsTemp!NO): lng结帐ID = Val(Nvl(rsTemp!结帐ID))
    
    If isCheckExiseSingularity(strNo) Then         '是作废的
        Exit Function
    End If
    
    str操作员姓名 = Nvl(rsTemp!操作员姓名)
    If str操作员姓名 <> UserInfo.姓名 Then
        strSQL = "" & _
        "   Select  结算序号 From 病人预交记录  A" & _
        "   Where 结帐ID=[1]  " & _
        "               And exists(Select 1 From 结算方式 B Where nvl(A.结算方式,'-')=b.名称 and 性质 not in ('3','4') )"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
        If Not rsTemp.EOF Then

            '存在其他结算方式
           Call MsgBox("注意:" & vbCrLf & _
            "       该病人存在异常的收费单据,但操作员 " & vbCrLf & _
            "       [" & str操作员姓名 & "]收取了一部分, " & _
            "       请到该操作员处进行收费!", vbOKOnly + vbInformation, gstrSysName)
            PatiErrBillPay = True
            Exit Function
        End If


    End If
    If MsgBox("注意:" & vbCrLf & _
                        "       该病人存在异常的收费单据" & IIf(str操作员姓名 <> UserInfo.姓名, ",该单据是操作员[" & str操作员姓名 & "]收取的," & vbCrLf, "") & " ,是否重新对该单据进行收费?" & vbCrLf & vbCrLf & _
                        "『是』代表重新对异常单据收费 " & vbCrLf & _
                        "『否』代表不对异常单据进行处理,继续新收费.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Function
    End If
    PatiErrBillPay = True
    mblnErrBill = True
    If LoadErrBillCharge(strNo) = False Then Exit Function
    
    '并发检查
    If zlIsCheckExistErrBill(mlng结算序号) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng结算序号) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '将这部分单据更新为当前操作员
    'Zl_门诊异常收费_更新操作员
    gstrSQL = "Zl_门诊异常收费_更新操作员("
    '病人id_In     门诊费用记录.病人id%Type,
    gstrSQL = gstrSQL & "" & lng病人ID & ","
    '操作员编号_In 门诊费用记录.操作员编号%Type,
    gstrSQL = gstrSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In 门诊费用记录.操作员姓名%Type,
    gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
    '结算序号_In   病人预交记录.结算序号%Type
    gstrSQL = gstrSQL & IIf(lng结算序号 = 0, mlng结算序号, lng结算序号) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    If ReChargeFee = False Then Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadCurBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载当前结算信息
    '入参:strBalance:本次收费的结算方式,格式如下:
    '        金额:缴款标志(1-缴款;2-找补)|结算方式1:金额1:缴款标志(1-缴款;2-找补)|...
    '返回：本次连续收费的总额
    '编制:刘兴洪
    '日期:2011-11-02 13:27:04
    '问题:42791
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim int性质 As Integer
    
    Call InitBalanceGrid
    If grsTotal Is Nothing Then Exit Function
    If grsTotal.State <> 1 Then Exit Function
    
    With vsBalance
        '性质:0-缴款;1-找补,2-冲预交;其他(mod 10:0-普通结算;1-医保结算;2-三方接品;3-一卡通)
        grsTotal.Sort = "性质"
        .Rows = IIf(.Rows >= grsTotal.RecordCount, .Rows, grsTotal.RecordCount)
        lngRow = 0
        Do While Not grsTotal.EOF
            '性质 ,结算方式  结算金额
            '从frmChargePayMentWin-传入,主要是一些累计数
            .TextMatrix(lngRow, 0) = Nvl(grsTotal!结算方式)
            .TextMatrix(lngRow, 1) = Format(Val(Nvl(grsTotal!结算金额)), "###0.00;-###0.00;0.00;0.00")
             int性质 = Val(Nvl(grsTotal!性质))
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .COLS - 1) = Me.ForeColor
            If int性质 = 0 Then
                .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True
            ElseIf int性质 = 1 Then
                .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True
                'If Val(.TextMatrix(lngRow, 1)) < 0 Then '45416
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .COLS - 1) = vbRed
               ' End If
            End If
            lngRow = lngRow + 1
            grsTotal.MoveNext
        Loop
    End With
End Function

Private Function ModifyNotInsureNOs(ByVal strNotSucceedNo As String, _
    ByVal strSucceedNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改未调用成功的医保单据
    '入参:strNotSucceedNo-医保结算不成功的单据
    '        strSucceedNos-医保结算成功的单据
    '        blnErrReChager-异常单据重新收费
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-12-17 22:37:04
    '问题:44535
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInfor As String
    Dim varNos As Variant, varNotNOs As Variant
    Dim intNum As Integer, intNotNum As Integer
    Dim intType As Integer
    If strNotSucceedNo = "" Then Exit Function
    varNos = Split(strSucceedNos, ","): varNotNOs = Split(strNotSucceedNo, ",")
    If strSucceedNos <> "" Then intNum = UBound(varNos) + 1
    If strNotSucceedNo <> "" Then intNotNum = UBound(varNotNOs) + 1
    intType = 0
    If intNum <> 0 Then
        strInfor = "医保成功调用" & intNum & "张" & vbCrLf & _
        "    " & strSucceedNos & vbCrLf
    End If
    strInfor = strInfor & "" & _
    "医保非成功调用" & intNotNum & "张" & vbCrLf & _
    "    " & strNotSucceedNo & vbCrLf

    If intNum = 0 Then
        strInfor = strInfor & "" & _
        "不能进行医保结算!"
        Call MsgBox(strInfor, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
        Exit Function
        intType = 1
    Else
       strInfor = strInfor & "" & _
        "目前只能对成功交易部分进行收费!"
    End If
    Call MsgBox(strInfor, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
       
    On Error GoTo errHandle
    'Zl_医保收费异常_Update
    strSQL = "Zl_医保收费异常_Update("
    '  Nos_In          Varchar2,
    strSQL = strSQL & "'" & strNotSucceedNo & "',"
    '  更新结算方式_In Integer:=0
    strSQL = strSQL & "" & intType & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    ModifyNotInsureNOs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearDisplaySHow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除双屏显示
    '编制:刘兴洪
    '日期:2011-12-29 09:54:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If Not gblnLED Then Exit Sub
    If Not (mbytInFun = 0 And mbytInState = 0) Then Exit Sub
    If mblnNotClearLedDisplay Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub
Private Function SaveChargeBill(ByRef lng结算序号 As Long, ByRef strBalanceIDs As String, ByRef strSaveNos As String, _
    Optional ByRef strModiNos As String, _
    Optional ByRef blnSaveBill As Boolean, _
    Optional ByRef blnNotCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前输入的单据(适用于收费 )
    '出参:lng结算序号-返回本次保存单据的结算序号
    '       strBalanceIDs-返回成功结帐的结帐IDs( 用来传给医保接口或第三方数据的修正)
    '       strSaveNos-返回已成功保存的单据号，格式为"'AAA','BBB',..."
    '       strModiNOs -修改的是多单据收费中的一张时，返回该多张单据的所有NO，格式如"'AAA','BBB',..."
    '       blnSaveBill-是否单据已经保存成功
    '       blnNotCommit-不进行事务提交（主要是处理普通病人收费(不用一卡通结算，不是医保结算的情况，减少出现异常的出现)
    '返回:收费成功或单据保存存功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-26 17:28:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '     *** 医保收费时,先临时保存为划价单,在结算前再转为收费单,以避免更新药品库存时因等待同一事务的医保结算操作而锁表 ***
    Dim lng打印ID As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim int序号 As Integer, int价格父号 As Integer, int行号 As Integer
    Dim lng结帐ID As Long, int药品行次 As Integer, str医疗付款 As String
    Dim dbl数次 As Double, dbl单价 As Double, cur缴款 As Currency
    Dim strDeptIDs As String, strTmp As String, strDelBill As String, strBillNO As String
    Dim str收费结算 As String, str保险结算 As String, str收费结算校对 As String
    Dim arrSQL As Variant, arrPut As Variant, arrOTMSQL As Variant
    Dim bln直接收费 As Boolean, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim CurOneCard As Currency, dblOneCardBalance As Double
    Dim strCardNo  As String, intCardType As Integer, strTransFlow As String
    Dim str中药形态 As String
    Dim strStuffDept As String          '自动发料的部门
    Dim strAdvance As String            '医保结算返回的信息:"结算方式|结算金额||....."
    Dim blnPriceSaved As Boolean        '医保收费时是否已存为划价单,用于在转为收费单及医保结算事务失败回退后删除划价单
    Dim blnMedicareCheck As Boolean     '是否执行医保结算校对
    Dim strInvoice As String            '当前单据使用的票据号，用于医保一张单据只打一张票的情况
    Dim cllRqure As Collection
    Dim rsSqure As ADODB.Recordset
    Dim str结帐IDs As String
    Dim bln应付款 As Boolean
    Dim dbl应缴额 As Double, lng结帐序号 As Long
    Dim cllPutout As Collection '自动发料
    Dim cllYBCommit As Collection   'SQL,key(单据号)
    Dim cllPro As Collection, cllDelete As Collection, cllPageInfor As Collection
    Dim cur已缴合计 As Currency
    Dim strSaveCuessNos As String
    
    '只处理收费单
    If mblnSaveAsPrice Then Exit Function
    If mbytInFun <> 0 Then Exit Function
    '新的发药窗品集(目前只针手工录入有效)
    Set mCllWindows = New Collection
    
    strBalanceIDs = ""
    strSaveNos = "": cur已缴合计 = 0: strModiNos = ""
    Err = 0: On Error GoTo Errhand:
    If cbo医疗付款.ListIndex <> -1 Then
        str医疗付款 = Mid(cbo医疗付款.Text, 1, InStr(1, cbo医疗付款, "-") - 1)
    End If
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    strInvoice = Trim(txtInvoice.Text)
    arrOTMSQL = Array()
    
    '修改功能时,是否修改医嘱附费
    If mstrInNO <> "" Then
        Call BillisAdviceMoney(mstrInNO, 1, lng医嘱ID, lng发送号)
    End If
    
    blnSaveBill = False
    dbl应缴额 = 0: lng结算序号 = 0
     Set cllPro = New Collection
    Set cllPageInfor = New Collection
    Set mcllPayDrugAndStuff = New Collection
     lng结算序号 = 0
    
    '对每张单据独立执行保存
    For p = 1 To mobjBill.Pages.Count
        int序号 = 0: int行号 = 0: blnPriceSaved = False
        int药品行次 = 0: strDeptIDs = "": strStuffDept = ""
        '当前收费单据的各类结算
        If mbytInFun = 0 And Not mblnSaveAsPrice Then
            str保险结算 = GetMedicareStr(p)
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            If lng结算序号 = 0 Then lng结算序号 = lng结帐ID
            str结帐IDs = str结帐IDs & "," & lng结帐ID
        End If
        
        '产生每张收费单据的单据号
        bln直接收费 = False
        strBillNO = mobjBill.Pages(p).NO
        If mobjBill.Pages(p).NO = "" Then
            '为保存失败后仍能识别,不改对象NO
            strBillNO = zlDatabase.GetNextNo(13)    '收费单
            bln直接收费 = True
        End If
        If p = 1 Then
            mobjBill.NO = strBillNO: gstrModiNO = strBillNO
        End If
        
        '主要为消息发送用,为每页保存的单据号
        mobjBill.Pages(p).收费单号 = strBillNO
        
        
        arrSQL = Array() '多单据时,逐张单据提交
        If Not bln直接收费 Then
            '1.收费新单据功能时,提取的划价单收费
            '虽然Zl_病人划价收费_Insert没有更新医保信息,但在根据病人提取的划价单时执行了zl_门诊划价记录_Update,已更新
            '---------------------------------------------------------------
            'Zl_病人划价收费_Insert
           gstrSQL = "Zl_病人划价收费_Insert("
            '  No_In         门诊费用记录.NO%Type,
            gstrSQL = gstrSQL & "'" & strBillNO & "',"
            '  病人id_In     门诊费用记录.病人id%Type,
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.病人ID) & ","
            '  病人来源_In   Number,
            gstrSQL = gstrSQL & "" & gint病人来源 & ","
            '  付款方式_In   门诊费用记录.付款方式%Type,
            gstrSQL = gstrSQL & "'" & str医疗付款 & "',"
            '  姓名_In       门诊费用记录.姓名%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.姓名 & "',"
            '  性别_In       门诊费用记录.性别%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.性别 & "',"
            '  年龄_In       门诊费用记录.年龄%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.年龄 & "',"
            '  病人科室id_In 门诊费用记录.病人科室id%Type,
            gstrSQL = gstrSQL & "" & IIf(mobjBill.Pages(p).医嘱序号 > 0, "Null", ZVal(mobjBill.科室ID, , mobjBill.Pages(p).开单部门ID)) & ","
            '  开单部门id_In 门诊费用记录.开单部门id%Type,
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).开单部门ID) & ","
            '  开单人_In     门诊费用记录.开单人%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).开单人 & "',"
            '  保险结算_In   Varchar2,
            If mstrYBPati <> "" And str保险结算 <> "" Then
                gstrSQL = gstrSQL & "'" & str保险结算 & "',"
            Else
                gstrSQL = gstrSQL & "NULL,"
            End If
            '  结帐id_In     门诊费用记录.结帐id%Type,
            gstrSQL = gstrSQL & "" & lng结帐ID & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            gstrSQL = gstrSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
            '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
            gstrSQL = gstrSQL & "'" & tbsBill.Tabs(p).Tag & "',"
            '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
            gstrSQL = gstrSQL & "" & chk急诊.Value & ","
            '  登记时间_In   门诊费用记录.登记时间%Type := Null,
            gstrSQL = gstrSQL & "" & "NULL" & ","
            '  结算序号_In   病人预交记录.结算序号%Type := Null
            gstrSQL = gstrSQL & "" & lng结算序号 & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
                        
            '获取自动发药的多个药房
            If gbln收费后自动发药 Then
                strDeptIDs = strDeptIDs & "," & Get发药部门IDs(strBillNO)
            End If
            '针对每张单据收集卫料发料部门,以便自动发料,是否是跟踪在用材料在SQL中判断
            If gbln门诊自动发料 Then
                strStuffDept = strStuffDept & "," & Get发药部门IDs(strBillNO, "'4'")
            End If
            '通过划价单收费的方式收取了挂号发卡的费用,则不用删除该费用
            If strBillNO = mstrCardNO Then mstrCardNO = ""
        ElseIf bln直接收费 Then
            '2.直接输入的单据内容,包括新增和修改,可能是收费(或收费界面保存为划价单),记帐,划价
            '---------------------------------------------------------------
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                If mobjBillDetail.数次 <> 0 Then
                    For Each mobjBillIncome In mobjBillDetail.InComes
                        int序号 = int序号 + 1 '当前记录序号
                        '1.单据主体---------------------------------------------------------------
                        With mobjBill                              '医保收费时,先临时保存为划价单,在结算前再转为收费单
                            If mstrYBPati = "" Then
                                gstrSQL = "zl_病人门诊收费_INSERT('" & strBillNO & "'," & int序号 & "," & ZVal(.病人ID) & "," & _
                                    IIf(gint病人来源 = 2, 2, 1) & "," & ZVal(.标识号) & ",'" & str医疗付款 & "'," & _
                                    "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                                    .加班标志 & "," & ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                    ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
                            Else
                                gstrSQL = "zl_门诊划价记录_INSERT('" & strBillNO & "'," & int序号 & "," & ZVal(.病人ID) & "," & _
                                    ZVal(.主页ID) & "," & ZVal(.标识号) & ",'" & str医疗付款 & "'," & _
                                    "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                                    .加班标志 & "," & ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                    ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
                            End If
                            
                        End With
        
                        '2.收费细目部份---------------------------------------------------------------
                        With mobjBillDetail
                            If .序号 <> int行号 Then     '处理从属父号
                                int行号 = .序号
                                int价格父号 = int序号
                                '重新处理从属父号
                                If mobjBill.Pages(p).Details(.序号).从属父号 = 0 Then
                                    For i = .序号 + 1 To mobjBill.Pages(p).Details.Count
                                        If mobjBill.Pages(p).Details(i).从属父号 = .序号 Then
                                            '当父项目有多个收入项目(多个序号)时,取第一个序号
                                            mobjBill.Pages(p).Details(i).从属父号 = int序号
                                        End If
                                    Next
                                End If
                            End If
        
                            If Not Set发药窗口(p, mobjBillDetail) Then
                                Exit Function
                            End If
                            
                            '医保直接收费时,因为先暂存为划价单,收费时需要取发药窗口
                            If mstrYBPati <> "" Then tbsBill.Tabs(p).Tag = .发药窗口
                            dbl数次 = .数次
                            If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                                dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                            End If
                            
                            gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                            If mstrYBPati = "" Then
                                gstrSQL = gstrSQL & IIf(.保险项目否, 1, 0) & "," & ZVal(.保险大类ID) & ",'" & .发药窗口 & "'," & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & IIf(.工本费, 8, .附加标志) & ","
                            Else
                                gstrSQL = gstrSQL & "'" & .发药窗口 & "'," & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & ","
                            End If
                            gstrSQL = gstrSQL & IIf(.执行部门ID = 0, "NULL", .执行部门ID) & ","
                        End With
        
                        '3.收入项目部份---------------------------------------------------------------
                        With mobjBillIncome
                            dbl单价 = .标准单价
                            If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 And gbln药房单位 Then
                                dbl单价 = Format(.标准单价 / mobjBillDetail.Detail.药房包装, gstrFeePrecisionFmt)
                            End If
                            gstrSQL = gstrSQL & IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & .收入项目ID & "," & _
                                    "'" & .收据费目 & "'," & dbl单价 & "," & .应收金额 & "," & .实收金额 & ","
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = gstrSQL & "NULL,"
                            End If
                        End With
        
                        '4.其它部分
                        '---------------------------------------------------------------
                        gstrSQL = gstrSQL & _
                                "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                        If mobjBillDetail.收费类别 = "7" Then
                            str中药形态 = "'" & mobjBillDetail.Detail.中药形态 & "'"
                        Else
                            str中药形态 = "NULL"
                        End If
                        '中药形态_In       住院费用记录.结论%Type := Null
                        
                        If mstrYBPati = "" Then
                            '非医保收费,并且不是划价
                            gstrSQL = gstrSQL & lng结帐ID & "," & lng结算序号 & ","
                            '卫材类别ID
                            gstrSQL = gstrSQL & "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                "'" & mobjBillDetail.摘要 & "'," & chk急诊.Value & ",'|" & mobjBill.Pages(mintPage).煎法 & "'" & _
                                "," & str中药形态 & ")"
                                '只在第一张单据的第一条记录时传入
                        Else
                            '门诊划价,收费功能划价,医保收费
                            gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                                "'" & mobjBillDetail.摘要 & "'," & ZVal(lng医嘱ID) & ",NULL,NULL,'|" & mobjBill.Pages(mintPage).煎法 & _
                                "',NULL,NULL," & gint病人来源 & ",'" & mobjBillDetail.保险编码 & "'," & _
                                "'" & mobjBillDetail.Detail.类型 & "'," & IIf(mobjBillDetail.保险项目否, 1, 0) & "," & ZVal(mobjBillDetail.保险大类ID) & "," & _
                                str中药形态 & ")"
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
                    Next    '每一条收入项目
                    
                    '对每一行收费记录收集药品执行部门,门诊记帐划价单的审核操作,在Oracle过程中处理:zl_门诊记帐记录_Verify
                    '----------------------------------------------------------------------------------------------------------------
                    '自动发药,仅收费时且不是分离发药时                    '
                    With mobjBillDetail
                        If gbln收费后自动发药 And mbytInFun = 0 And Not mblnSaveAsPrice Then
                            If .执行部门ID <> 0 And InStr("5,6,7", .收费类别) > 0 Then
                                If InStr(strDeptIDs & ",", "," & .执行部门ID & ",") = 0 Then
                                    strDeptIDs = strDeptIDs & "," & .执行部门ID
                                End If
                            End If
                        End If
                        '自动发料,收费且不是保存为划价单或者门诊记帐,分离发药参数不影响卫材
                        If gbln门诊自动发料 And ((mbytInFun = 0 And Not mblnSaveAsPrice) Or (mbytInFun = 2 And mbytBilling = 0)) Then
                                If .执行部门ID <> 0 And .收费类别 = "4" And .Detail.跟踪在用 Then
                                    If InStr(strStuffDept & ",", "," & .执行部门ID & ",") = 0 Then
                                        strStuffDept = strStuffDept & "," & .执行部门ID
                                    End If
                                End If
                        End If
                    End With
                End If
            Next            '每一行收费项目
            '保存前一张单据的药房ID,以便多张单据时确定发药窗口
            If mobjBill.Pages.Count > 1 Then Call SaveDrugID(p)
            '修改后退除原单据(修改多收费单中的一张时需要后退费以统一重打)
            '--------------------------------------------------------------------------------------------------------
            If mstrInNO <> "" Then
                strDelBill = ""
                '修改医保收费单,必然为单张内全退,因为修改调用时已判断了如果不是全退,则不允许修改
                strDelBill = "zl_门诊收费记录_DELETE('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                    "NULL,NULL,'" & zlStr.NeedName(cbo结算方式.Text) & "',0,To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                '如果是多单据收费中的一张则将新单关联到与原单据的打印ID上,以便一起重打
                strTmp = GetMultiNOs(mstrInNO, lng打印ID)
                If UBound(Split(strTmp, ",")) = 0 Then
                    lng打印ID = 0: strModiNos = ""
                ElseIf lng打印ID <> 0 Then
                    strModiNos = strTmp
                End If
                '如果是修改医嘱的附费,则将新的NO放在附费中
                If lng医嘱ID <> 0 And lng发送号 <> 0 Then
                    gstrSQL = "ZL_病人医嘱附费_Insert(" & lng医嘱ID & "," & lng发送号 & "," & IIf(mbytInFun = 2, 2, 1) & ",'" & strBillNO & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
                End If
            End If
        End If
        
        '收费后自动发药,记帐不自动发药,收费且不是保存为划价单,或者门诊记帐
        '-----------------------------------------------------------------------
        If strDeptIDs <> "" Then
            arrPut = Array()
            strDeptIDs = Mid(strDeptIDs, 2)
            For i = 0 To UBound(Split(strDeptIDs, ","))
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "ZL_药品收发记录_处方发药(" & Val(Split(strDeptIDs, ",")(i)) & ",8,'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & mobjBill.Pages(p).开单人 & "')"
            Next
        End If
        '收费后自动发料,在收费(直接收费,划价单导入收费),门诊记帐时执行
        If strStuffDept <> "" Then
            If strDeptIDs = "" Then arrPut = Array()
            strStuffDept = Mid(strStuffDept, 2)
            For i = 0 To UBound(Split(strStuffDept, ","))          '24-收费处方发料；25-记帐单处方发料
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "zl_材料收发记录_处方发料(" & Split(strStuffDept, ",")(i) & "," & IIf(mbytInFun = 0, 24, 25) & ",'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
            Next
        End If
        '执行相关SQL语句及提交医保结算,多张单据时,每张单据在独立事务中提交
        '--------------------------------------------------------------------------------------------------------------------------------
        If UBound(arrSQL) >= 0 Then
            '对SQL序列按收费细目ID排序
            For i = 0 To UBound(arrSQL) - 1
                For j = i + 1 To UBound(arrSQL)
                    If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                        strTmp = CStr(arrSQL(j)): arrSQL(j) = arrSQL(i): arrSQL(i) = strTmp
                    End If
                Next
            Next
            
            '医保直接收费时,先保存为划价单,再转为收费单
            '-------------------------------------------------------------------
            If bln直接收费 And mstrYBPati <> "" Then
                '1.先保存划价单,先提交库存更新以便不锁表
                If MCPAR.多单据调一次交易 Or MCPAR.多单据一次结算 Then
                    For i = 0 To UBound(arrSQL)
                        zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                    Next
                Else
                    On Error GoTo errH
                    gcnOracle.BeginTrans
                        blnTrans = True
                        For i = 0 To UBound(arrSQL)
                            Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
                        Next
                    gcnOracle.CommitTrans
                    blnTrans = False: blnPriceSaved = True
                End If
                
                '更新划价单的保险信息(保险项目否,医保大类ID,统筹金额)
                gstrSQL = "zl_门诊划价记录_Update(" & mintInsure & "," & mobjBill.病人ID & ",'" & strBillNO & "',0)"
                If MCPAR.多单据调一次交易 Or MCPAR.多单据一次结算 Then
                    zlAddArray cllPro, gstrSQL
                Else
                   Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
                
                '划价单转为收费单
                 'Zl_病人划价收费_Insert
                gstrSQL = "Zl_病人划价收费_Insert("
                '  No_In         门诊费用记录.NO%Type,
                gstrSQL = gstrSQL & "'" & strBillNO & "',"
                '  病人id_In     门诊费用记录.病人id%Type,
                gstrSQL = gstrSQL & "" & mobjBill.病人ID & ","
                '  病人来源_In   Number,
                gstrSQL = gstrSQL & "" & gint病人来源 & ","
                '  付款方式_In   门诊费用记录.付款方式%Type,
                gstrSQL = gstrSQL & "'" & str医疗付款 & "',"
                '  姓名_In       门诊费用记录.姓名%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.姓名 & "',"
                '  性别_In       门诊费用记录.性别%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.性别 & "',"
                '  年龄_In       门诊费用记录.年龄%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.年龄 & "',"
                '  病人科室id_In 门诊费用记录.病人科室id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.科室ID, , mobjBill.Pages(p).开单部门ID) & ","
                '  开单部门id_In 门诊费用记录.开单部门id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).开单部门ID) & ","
                '  开单人_In     门诊费用记录.开单人%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).开单人 & "',"
                '  保险结算_In   Varchar2,
                gstrSQL = gstrSQL & "" & IIf(str保险结算 <> "", "'" & str保险结算 & "'", "NULL") & ","
                '  结帐id_In     门诊费用记录.结帐id%Type,
                gstrSQL = gstrSQL & "" & lng结帐ID & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.编号 & "',"
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
                gstrSQL = gstrSQL & "'" & tbsBill.Tabs(p).Tag & "',"
                '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
                gstrSQL = gstrSQL & "" & chk急诊.Value & ","
                '  登记时间_In   门诊费用记录.登记时间%Type := Null,
                gstrSQL = gstrSQL & "" & "NULL" & ","
                '  结算序号_In   病人预交记录.结算序号%Type := Null
                gstrSQL = gstrSQL & "" & lng结算序号 & ")"
            End If
            
            '医保多单据一次结算时，所有单据做为一个事务提交
            If (MCPAR.多单据一次结算 Or MCPAR.多单据调一次交易) And mstrYBPati <> "" And strDelBill = "" Then
                '1.划价单转收费
                zlAddArray cllPro, gstrSQL
                '2.误差费用
                If mobjBill.Pages(p).误差金额 <> 0 Then '44657
                    gstrSQL = "zl_门诊收费误差_Insert('" & strBillNO & "'," & mobjBill.Pages(p).误差金额 & ",0,1)"
                    zlAddArray cllPro, gstrSQL
                End If
                '3.收费后自动发药,自动发料
                If strDeptIDs <> "" Or strStuffDept <> "" Then
                    For i = 0 To UBound(arrPut)
                        zlAddArray mcllPayDrugAndStuff, arrPut(i)
                    Next
                End If
            Else
                On Error GoTo errH
                    '修改功能相关处理
                    If mstrYBPati <> "" Then
                        gcnOracle.BeginTrans: blnTrans = True
                    End If
                    '先删除原单据,因为库存和预交款需要先还原
                    If strDelBill <> "" Then
                        If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(strDelBill, Me.Caption)
                        Else
                            zlAddArray cllPro, strDelBill
                        End If
                    End If
                    'a.非医保直接收费
                    If Not (bln直接收费 And mstrYBPati <> "") Then
                        '删除就诊卡划价单:多张单据时只删除一次(因为通过就诊卡号读病人时,就诊卡划价单已生成收费细目行,所以要删除)
                        If mstrCardNO <> "" And strSaveNos = "" Then
                            gstrSQL = "zl_门诊划价记录_Delete('" & mstrCardNO & "')"
                            If mstrYBPati <> "" Then
                                 Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                            Else
                                    zlAddArray cllPro, gstrSQL
                            End If
                        End If
                        '执行主体的SQL语句
                        For i = 0 To UBound(arrSQL)
                            If mstrYBPati <> "" Then
                                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
                            Else
                                zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                            End If
                        Next
                        'b.医保直接收费
                    Else
                         If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                         Else
                            zlAddArray cllPro, gstrSQL
                         End If
                    End If
                    
                    '收费完成后的处理
                    '-----------------------------------------------------
                    '先填写开始票据号以便医保调用时上传,多张分别打印时,填写相同的,打印调用时将重写,取消打印或打印失败将清除
                    '修改时,只填写新单据的开始票据号,因为医保只对新单据上传
                    If strInvoice <> "" And mblnPrint Then
                        gstrSQL = "Zl_票据起始号_Update('" & strBillNO & "','" & strInvoice & "',1)"
                        If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        Else
                            zlAddArray cllPro, gstrSQL
                        End If
                    End If
                
                    '每张单据处理误差,该结帐ID与刚生成的收费记录相同
                    If mobjBill.Pages(p).误差金额 <> 0 Then '44657
                        gstrSQL = "zl_门诊收费误差_Insert('" & strBillNO & "'," & mobjBill.Pages(p).误差金额 & ",0,1)"
                        If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        Else
                            zlAddArray cllPro, gstrSQL
                        End If
                    End If
                    
                    '收费后自动发药,自动发料
                    If strDeptIDs <> "" Or strStuffDept <> "" Then
                        For i = 0 To UBound(arrPut)
                            zlAddArray mcllPayDrugAndStuff, CStr(arrPut(i))
                        Next
                    End If
                    
                    '修改功能相关处理
                    If strDelBill <> "" Then
                        '收费：新单据关联到原单据的打印ID上,以便一起重打,此时并未产生票据
                        If lng打印ID <> 0 And mblnPrint Then
                            gstrSQL = "zl_门诊收费票据_Insert('" & strBillNO & "','',Null,'',Null," & lng打印ID & ",0)"
                            If mstrYBPati <> "" Then
                                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                            Else
                                zlAddArray cllPro, gstrSQL
                            End If
                        End If
                    End If
                    '医保处理
                    If mstrYBPati <> "" Then
                        If zlInsureOneBillClinicSwap(lng结帐ID, strBillNO, strInvoice, strDelBill <> "", p, Original.结帐ID, blnPriceSaved) = False Then
                           'If strSaveNos <> "" Then strSaveNos = Mid(strSaveNos, 2)
                           If strSaveCuessNos <> "" Then
                                 strSaveCuessNos = Mid(strSaveCuessNos, 2)
                                 i = UBound(Split(strSaveCuessNos, ",")) + 1
                                Call MsgBox("注意:" & vbCrLf & _
                                                      "      医保已经成功收费了" & i & "张单据,但还存在" & mobjBill.Pages.Count - i & "张" & vbCrLf & _
                                                      "单据未收费成功,现只对成功单据进行收费!" & vbCrLf & _
                                                      "医保成功收费单据如下: " & vbCrLf & _
                                                       strSaveCuessNos, vbDefaultButton1 + vbInformation + vbOKOnly, gstrSysName)
                               blnSaveBill = True: SaveChargeBill = True
                            End If
                            
                            Exit Function
                        End If
                        strSaveCuessNos = strSaveCuessNos & "," & strBillNO
                    End If
                    
                On Error GoTo 0
            End If
            strBalanceIDs = IIf(strBalanceIDs = "", "", strBalanceIDs & ",") & lng结帐ID
            cllPageInfor.Add Array(lng结帐ID, strBillNO), "K" & p
            
            '提交成功后再累加
            If mbytInFun = 0 And Not mblnSaveAsPrice Then
                cur已缴合计 = cur已缴合计 + mobjBill.Pages(p).应缴金额
            End If
            
            strSaveNos = strSaveNos & ",'" & strBillNO & "'"
            If Left(strSaveNos, 1) = "," Then strSaveNos = Mid(strSaveNos, 2)
            '加入单据历史记录(所有类型单据)
            cboNO.AddItem strBillNO, 0
            For i = cboNO.ListCount - 1 To 10 Step -1
                cboNO.RemoveItem i '只显示10个
            Next
        End If
    Next  '下一张单据
    
    On Error GoTo errH:
    '先保存单据
    Dim blnAffair As Boolean
    blnTrans = True
    If mstrYBPati = "" Or (mstrYBPati <> "" And (MCPAR.多单据调一次交易 Or MCPAR.多单据一次结算)) Then
        If mcllPayDrugAndStuff.Count <> 0 And blnNotCommit Then
            '自动发料(只一次提交数据时,同时将发料,放在同一事务中)
            For i = 1 To mcllPayDrugAndStuff.Count
                zlAddArray cllPro, mcllPayDrugAndStuff(i)
            Next
            Set mcllPayDrugAndStuff = New Collection
        End If
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If mstrYBPati <> "" Then
            If zlInsureClinicSwap(cllPageInfor, lng结算序号, strInvoice, strDelBill <> "", _
                strBalanceIDs, strSaveNos, strSaveCuessNos, blnAffair) = False Then
                If Not blnAffair Then gcnOracle.RollbackTrans
                If strSaveCuessNos <> "" Then blnSaveBill = True:
                Exit Function
            End If
        End If
        If blnAffair = False And blnNotCommit = False Then
            gcnOracle.CommitTrans
        Else
            '进行自动发料(同一事务)
            If blnNotCommit Then
                zlExecuteProcedureArrAy mcllPayDrugAndStuff, Me.Caption, True, True
                Set mcllPayDrugAndStuff = Nothing
            End If
        End If
    End If
    blnSaveBill = True: blnTrans = False: mblnNotClearLedDisplay = True
    SaveChargeBill = True
    Exit Function
errH:
    If Err.Description Like "*当前计算单价不一致*" Then
        If blnTrans Then gcnOracle.RollbackTrans
        If MsgBox("某些分批药品价格已发生变化，要自动重算价格吗？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
            Exit Function
        End If
     Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Exit Function
ErrPutOut:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlInsureOneBillClinicSwap(ByVal lng结帐ID As Long, _
    ByVal strBillNO As String, _
    ByVal strInvoice As String, _
    ByVal blnModifyBill As Boolean, _
    ByVal intPage As Integer, _
    ByVal lng原结帐ID As Long, _
    ByVal blnPriceSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保调用(单张单据)
    '入参:blnModifyBill-是否修改单据
    '       strBalanceIDs:本次结帐的ID,分别用逗号分离
    '       strSaveNos-保存的单据号
    '       lng原结帐ID-原被修改的单据的结帐ID
    '       blnPriceSaved-是否保存了划价单的
    '返回:医保调用成功或非医保,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTransMedicare As Boolean, str保险结算 As String, strAdvance As String, blnMedicareCheck As Boolean
    Dim strTmp As String, i As Long
    Dim intExeCount As Integer
    
    On Error GoTo errHandle
    '更新标志
    ' Zl_病人门诊收费_医保更新
    gstrSQL = "Zl_病人门诊收费_医保更新("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    gstrSQL = gstrSQL & "" & lng结帐ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    gstrSQL = gstrSQL & "NULL,"
    '  保险结算_In Varchar2
    gstrSQL = gstrSQL & "NULL)"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    '---------------------------------------------------
    '1.修改时,先退除原收费单据(改费方式)
    blnTransMedicare = False
    strAdvance = ""
    If blnModifyBill Then
        strAdvance = mobjBill.Pages.Count & "|" & intPage
        If Not gclsInsure.ClinicDelSwap(lng原结帐ID, False, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans: Call DelMedicareTempNO(blnPriceSaved, strBillNO): Exit Function
        Else
            blnTransMedicare = True
        End If
    End If
    
    '2.调医保交易
    '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
    If (GetMedicareSum(, intPage) <> 0 Or MCPAR.门诊必须传递明细) Then
        strAdvance = mobjBill.Pages.Count & "|" & intPage
        If Not gclsInsure.ClinicSwap(lng结帐ID, GetMedicareSum(mstr个人帐户, intPage), _
            GetMedicareSum("医保基金", intPage), mobjBill.Pages(intPage).全自付, mobjBill.Pages(intPage).先自付, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans:  Call DelMedicareTempNO(blnPriceSaved, strBillNO)
            If blnModifyBill Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
            Exit Function
        Else
            blnTransMedicare = True
        End If
    End If
    gcnOracle.CommitTrans
    If blnModifyBill Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
    str保险结算 = GetMedicareStr(intPage)
    blnMedicareCheck = zlInsureCheck(str保险结算, strAdvance)
    If blnMedicareCheck Then
        ' Zl_病人门诊收费_医保更新
        gstrSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        gstrSQL = gstrSQL & "" & lng结帐ID & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        gstrSQL = gstrSQL & "NULL,"
        '  保险结算_In Varchar2
        gstrSQL = gstrSQL & IIf(blnMedicareCheck, "'" & strAdvance & "'", "NULL") & ")"
        Err = 0: On Error GoTo ErrModifyTag:
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
    Err = 0: On Error GoTo errHandle:
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
    zlInsureOneBillClinicSwap = True
    Exit Function
ErrModifyTag:
    If intExeCount > 3 Then
        MsgBox "单据[" & strBillNO & "]进行医保结算校对3次以上失败,请与系统管理员联系,较对数据如下:" & vbCrLf & strAdvance & vbCrLf & "错误原因如下:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrSysName
        intExeCount = 3
    Else
        MsgBox "单据[" & strBillNO & "]进行医保结算校对失败,结算金额不正确,点击确定后重新较对, 较对数据如下:" & vbCrLf & _
         strAdvance & vbCrLf & "错误原因如下:" & vbCrLf & Err.Description, vbInformation, gstrSysName
    End If
    intExeCount = intExeCount + 1
    Resume
    Exit Function
errHandle:
     gcnOracle.RollbackTrans
     Call ErrCenter
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mintInsure)
    Else     '如果医保成功了，不删除划价单，费用失败可以重收
        Call DelMedicareTempNO(False, strBillNO)
    End If
    Call SaveErrLog
End Function
Public Function zlGetToTatal() As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次缴款总额
    '编制:刘兴洪
    '日期:2012-02-17 15:25:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim curTotal As Currency, intCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).Details.Count > 0 Then
            curTotal = curTotal + mobjBill.Pages(i).误差金额
            For j = 1 To mobjBill.Pages(i).Details.Count
                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                    curTotal = curTotal + mobjBill.Pages(i).Details(j).InComes(k).实收金额
                Next
            Next
        Else    '提取划价单收费时没有明细费用
            curTotal = curTotal + mobjBill.Pages(i).误差金额
            curTotal = curTotal + mobjBill.Pages(i).实收金额
        End If
    Next
    
    '如果没有,再尝试从表格中取(仅一张单据时)
    If curTotal = 0 And tbsBill.Tabs.Count = 1 _
        And Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.项目) = "") Then
        intCol = BillCol.实收金额
        For i = 1 To Bill.Rows - 1
            If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                curTotal = curTotal + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
            End If
        Next
    End If
    zlGetToTatal = Format(curTotal, gstrDec)
End Function


Private Function Get未发药品发药窗口(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As String
    '-------------------------------------------------------------------------
    '功能：判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
    '返回：若存在相同执行部门的未发药品，则返回未发药品的发药窗口，否则返回空
    '编制：冉俊明
    '日期：2014-04-09
    '问题：71902
    '说明：
    '   同一个人病人不同时间段多张单据收费，分配同一个发药窗口，方便病人取药
    '-------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select 发药窗口" & vbNewLine & _
            "From 未发药品记录" & vbNewLine & _
            "Where 单据 = 8 And 发药窗口 Is Not Null And 病人id = [1] And 库房id = [2]" & vbNewLine & _
            "Order By 已收费 Desc, 填制日期 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人未发药品发药窗口", lng病人ID, lng执行部门ID)
    
    If Not rsTemp.EOF Then
        Get未发药品发药窗口 = Nvl(rsTemp!发药窗口)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Set发药窗口(ByVal p As Integer, ByRef objBillDetail As BillDetail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置发药窗口
    '返回:  设置成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-07-03 09:53:33
    '问题:45172
    '说明:
    '   根据药房ID来确定,相同的药房ID分配相同的发药窗口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, i As Long, strSendWindows As String
    Dim blnFind As Boolean
    Dim strTemp As String
    
    Err = 0:     On Error GoTo errHandle:
    
    With objBillDetail
        '收费、划价的药品行,处理发药窗口
        If Not InStr(",5,6,7,", .收费类别) > 0 Then Set发药窗口 = True: Exit Function
        
        '非修改单据
        'array(药房ID,窗口),检查是否存在该窗口，保证相同药房有同一个窗口
        strSendWindows = ""
        blnFind = False
        For i = 1 To mCllWindows.Count
            If mCllWindows(i)(0) = .执行部门ID Then
                strSendWindows = mCllWindows(i)(1): blnFind = True
            End If
        Next
        
        If mstrInNO <> "" Then
            '修改单据
            .发药窗口 = IIf(strSendWindows <> "", strSendWindows, .发药窗口) '修改时保持原有发药窗口
            Set发药窗口 = True
            Exit Function
        End If
        
        '71902,冉俊明,2014-04-09,同一个人病人不同时间段多张单据收费，分配同一个发药窗口，方便病人取药
        '判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
        strTemp = Get未发药品发药窗口(mobjBill.病人ID, .执行部门ID)
        If strTemp <> "" Then
            .发药窗口 = strTemp
            Set发药窗口 = True: Exit Function
        End If
        
        If strSendWindows <> "" Then    '存在发药窗口，以第一个为准
            .发药窗口 = strSendWindows: Set发药窗口 = True: Exit Function
        End If
        
        .发药窗口 = GetDrugWindow(.执行部门ID, .收费类别, p)
        If .发药窗口 = "" Then
           .发药窗口 = Get发药窗口(mobjBill.登记时间, .执行部门ID, .收费类别, _
                       IIf(.执行部门ID <> mlng西药房, "", mstr西窗), IIf(.执行部门ID <> mlng成药房, "", mstr成窗), IIf(.执行部门ID <> mlng中药房, "", mstr中窗))
        End If
        If .发药窗口 <> "" Then
            Select Case .收费类别
                Case "5"
                    mstr西窗 = .发药窗口
                Case "6"
                    mstr成窗 = .发药窗口
                Case "7"
                    mstr中窗 = .发药窗口
            End Select
        ElseIf ExistWindow(.执行部门ID, mrs发药窗口) Then
            MsgBox "无法分配" & GET部门名称(.执行部门ID, mrsUnit) & "的发药窗口，请检查是否正常安排窗口上班。", vbInformation, gstrSysName
            Exit Function
        End If
        If Not blnFind Then
            mCllWindows.Add Array(.执行部门ID, .发药窗口), "K" & .执行部门ID
        End If
    End With
    Set发药窗口 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Clear连续累计()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:信息改变时,清除连续累计的显示
    '编制:刘兴洪
    '日期:2012-08-01 10:28:35
    '问题:51670
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnClearBlance Then Exit Sub
    '清除结算信息
    Call InitBalanceGrid(True)
    mblnClearBlance = False
End Sub
Private Sub Set连续收费操作(Optional bln未建档 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置连续收费操作的相关信息
    '入参:bln未建档-病人未建档时
    '编制:刘兴洪
    '日期:2012-08-01 10:37:22
    '说明:
    '问题:51670
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mbln连续输入 Or mbytInFun <> 0 Or mbytInState <> 0 Then Exit Sub
    
    '只有收费才有连续收费
    With gTy_Module_Para
        If Not (.byt缴款控制 = 1 Or .byt缴款控制 = 3) Then Exit Sub
    End With
    
    '显示连续收费的情况给表格
    Call LoadCurBalance: sta.Panels(2).Text = IIf(mstrPrePati = "", "", "上一病人:" & mstrPrePati)
    If gTy_Module_Para.byt缴款控制 <> 3 Then Exit Sub
    If bln未建档 Then
        If mstrPrePati = Trim(txtPatient.Text) Or Trim(txtPatient.Text) = "" Then Exit Sub
    Else
        If mrsInfo Is Nothing Then Exit Sub
        If mrsInfo.State <> 1 Then Exit Sub
        '同一病人,允许连继输入
        If mstrPrePati = mrsInfo!姓名 Or mlngPrePati = Val(mrsInfo!病人ID) Then Exit Sub
    End If
    '不同病人时,结束连续输入
    mblnClearBlance = True
    mbln连续输入 = False: Set grsTotal = Nothing
End Sub

Private Sub WriteMzInforToCard(ByVal lng病人ID As Long, ByVal lng结算序号 As Long, Optional blnDelete As Boolean = False)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将门诊信息写入卡中
    '入参:blnDelete-是否退费
    '编制:刘兴洪
    '日期:2012-12-14 17:06:27
    '说明:
    '问题:56615
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strExpend As String
    '未确定刷卡类别,直接退出
    If InStr(1, mstrPrivs, ";门诊信息写卡;") = 0 Then Exit Sub
    If lng病人ID = 0 Then Exit Sub
    If mlngCardTypeID = 0 Then
        If blnDelete Then GoTo goDelete:
        Exit Sub
    End If
    Dim objCard As Card
    If IDKind.GetCurCard.接口序号 = mlngCardTypeID Then
        Set objCard = IDKind.GetCurCard
    Else
        Set objCard = IDKind.GetIDKindCard(mlngCardTypeID, CardTypeID)
    End If
    If objCard Is Nothing Then Exit Sub
    If objCard.是否写卡 = False Or objCard.接口序号 <= 0 Then Exit Sub '不准写卡的,不调用接口
    lngCardTypeID = objCard.接口序号
goDelete:
   Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, _
    lng病人ID, lng结算序号, strExpend)
End Sub
Private Sub SetInvoceSizeAndShowTittle()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整发票显示控件的大小和显示
    '编制:刘兴洪
    '日期:2013-05-07 16:14:02
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoice As New Collection
    Dim r As Long, c As Long
    Dim bytSel As Byte '1-选择;2-不选择,3-不能取消的选择(关联发票)
    Dim strInvoice As String '发票号
    Dim sngColWidth As Single
    Dim i As Long
    Err = 0: On Error GoTo Errhand:
    Set cllInvoice = New Collection
    With vsInvoice
        If .Rows = 1 And .Cell(flexcpLeft, 0, .COLS - 1) + .ColWidth(.COLS - 1) <= .Width Then Exit Sub
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                bytSel = .Cell(flexcpChecked, r, c)
                strInvoice = Trim(.Cell(flexcpData, r, c))
                sngColWidth = .ColWidth(c)
                If strInvoice <> "" Then
                    cllInvoice.Add Array(bytSel, strInvoice, sngColWidth)
                End If
            Next
        Next
        .Redraw = flexRDNone
        .Rows = 1
        .COLS = 1
        .Clear
        .TextMatrix(0, 0) = "回收发票"
        sngColWidth = .ColWidth(0)
        For i = 1 To cllInvoice.Count
            If sngColWidth + cllInvoice(i)(2) * 0.5 > .Width Then
                If .COLS <= 1 Then
                    .COLS = .COLS + 1
                    .ColWidth(.COLS - 1) = cllInvoice(i)(2)
                End If
                Exit For
            End If
            .COLS = .COLS + 1
            .ColWidth(.COLS - 1) = cllInvoice(i)(2)
            sngColWidth = sngColWidth + .ColWidth(.COLS - 1)
        Next
        .Cell(flexcpChecked, 0, .COLS - 1, .Rows - 1, .COLS - 1) = 0
        c = 0: r = 0
        For i = 1 To cllInvoice.Count
            If c >= .COLS - 1 Then
                .Rows = .Rows + 1
                r = r + 1
                c = 1
            Else
                c = c + 1
            End If
            .TextMatrix(r, c) = cllInvoice(i)(1)
            .Cell(flexcpData, r, c) = cllInvoice(i)(1)
            .Cell(flexcpChecked, r, c) = cllInvoice(i)(0)
            .ColWidth(c) = cllInvoice(i)(2)
        Next
        .Height = (.RowHeight(0) + 90) * (.Rows)
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    vsInvoice.Redraw = flexRDBuffered
End Sub
Private Sub ShowInvoiceInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示发票信息
    '编制:刘兴洪
    '日期:2013-05-27 11:36:16
    '25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    If Not (mbytInFun = 0 And mbytInState = 3) Then Exit Sub
    
    If mrsDelInvoice Is Nothing Then
        vsInvoice.Visible = False:
        Call Form_Resize
    End If
    If mrsDelInvoice.RecordCount = 0 Then
        vsInvoice.Visible = False:
        Call Form_Resize
        Exit Sub
    End If
    vsInvoice.Visible = True
    Call Form_Resize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadInvoiceData(ByVal strNo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载发票信息
    '编制:刘兴洪
    '日期:2013-05-07 17:07:38
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str序号 As String, varTemp As Variant
    Dim i As Long, str发票号 As String
    If gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    If Not (mbytInFun = 0 And mbytInState = 3) Then Exit Sub
    
    If mrsDelInvoice Is Nothing Then
        Set mrsDelInvoice = zlGetFromNoTOInvoice(strNo)
    End If
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    With Bill
        For i = 1 To .Rows - 1
            If InStr("销帐,退费", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
                If Bill.TextMatrix(i, Bill.COLS - 1) = "√" Then
                    str序号 = str序号 & "," & Bill.RowData(i)
                End If
            End If
        Next
    End With
    If str序号 <> "" Then str序号 = Mid(str序号, 2)
    str发票号 = GetFromNumToInvoiceNo(strNo, str序号)
    '加载发票号
    varTemp = Split(str发票号, ",")
    With vsInvoice
        .Clear
        .Rows = 1: .COLS = 1
        .FixedCols = 1
        .TextMatrix(0, 0) = "回收票据"
        .Redraw = flexRDNone
        .COLS = .COLS + UBound(varTemp) + 1
        For i = 0 To UBound(varTemp)
            If i + 1 > .COLS - 1 Then
                .COLS = .COLS + 1
            End If
            .TextMatrix(0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpData, 0, i + 1) = CStr(varTemp(i))
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        Call Form_Resize
        .Redraw = flexRDBuffered
    End With
End Sub
Private Function GetFromNumToInvoiceNo(ByVal strNo As String, ByVal str序号 As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据序号获取对应的发票号
    '入参:strNO-单据号
    '       str序号-序号,可以为多个,多个用逗号分离
    '       strNotInvoice-不包含的发票号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-05-07 17:38:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, str关联序号 As String
    Dim varTemp As Variant, i As Long, strTemp As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Function
    If mrsDelInvoice.State <> 1 Then Exit Function
    If mrsDelInvoice.RecordCount = 0 Then Exit Function
    With mrsDelInvoice
        str关联序号 = "": str发票号 = ""
        varTemp = Split(str序号, ",")
        .Filter = "NO='" & strNo & "'"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                strTemp = "," & Nvl(!序号) & ","
                For i = 0 To UBound(varTemp)
                    If InStr(1, strTemp, "," & varTemp(i) & ",") > 0 _
                        And InStr(str发票号 & ",", "," & Nvl(!票号) & ",") = 0 Then
                        str发票号 = str发票号 & "," & Nvl(!票号)
                        If Val(Nvl(!关联票号序号)) <> 0 Then
                            str关联序号 = str关联序号 & "," & Val(Nvl(!关联票号序号))
                        End If
                    End If
                Next
            .MoveNext
        Loop
        .Filter = 0: .MoveFirst
        If str关联序号 = "" Then GoTo GoSort:
        '需要查找关联票号
       varTemp = Split(Mid(str关联序号, 2), ",")
        Do While Not .EOF
                For i = 0 To UBound(varTemp)
                    If Val(varTemp(i)) = Val(Nvl(!关联票号序号)) _
                        And InStr(str发票号 & ",", "," & Nvl(!票号) & ",") = 0 Then
                        str发票号 = str发票号 & "," & Nvl(!票号)
                    End If
                Next
            .MoveNext
        Loop
    End With
    '进行排序处理
GoSort:
    If str发票号 = "" Then Exit Function
    str发票号 = Mid(str发票号, 2)
    GetFromNumToInvoiceNo = zlStringSort(str发票号)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckChargeDataValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查收费数据是否合法
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-06-25 16:34:58
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, bln检查库存 As Boolean, p As Integer
    Dim dblToTal As Double, strTmp As String, strInfo As String
    Dim lng药房ID As Long
    Dim colStock As Collection
    
    On Error GoTo errHandle

    '导入划价单收费时,如果是医嘱生成的,可能已作废
    For i = 1 To mobjBill.Pages.Count
        '针对每张单据判断(因为可能划价和收费混用),是否是导入医嘱生成的划价单收费
        If mobjBill.Pages(i).NO <> "" And mobjBill.Pages(i).医嘱序号 <> 0 Then
            If mobjBill.Pages(i).实收金额 <> GetBillSumByDB(mobjBill.Pages(i).NO) Then
                MsgBox "单据[" & mobjBill.Pages(i).NO & "]的部分收费记录已被他人修改或作废,请重新读取单据后再收费！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If CheckServeRange(0, .收费细目ID) = False Then Exit Function
            End With
        Next i
    Next p
    
   '药品库存检查(仅不足禁止时或分批时价药品)
    bln检查库存 = (InStr(mstrPrivs, "不检查库存") = 0)    '是否有权限不检查库存(分批和时价必须检查)
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
            
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & _
                                dblToTal & """。", vbInformation, gstrSysName
                            'tbsBill.Tabs(p).Selected = True
                            Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行药品""" & .Detail.名称 & _
                                    """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & _
                                    dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                                'tbsBill.Tabs(p).Selected = True
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblToTal & """。", vbInformation, gstrSysName
                            'tbsBill.Tabs(p).Selected = True:
                            Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID, .Detail.批次)
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                    """的当前库存" & IIf(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                                'tbsBill.Tabs(p).Selected = True:
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next
    
    '发药窗口检查(仅划价单)
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO <> "" And tbsBill.Tabs(i).Tag = "" Then
            lng药房ID = BillExistDrug(mobjBill.Pages(i).NO, 1)
            If lng药房ID <> 0 Then
                If ExistWindow(lng药房ID, mrs发药窗口) Then
                    MsgBox "无法分配" & GET部门名称(lng药房ID, mrsUnit) & "的发药窗口，请确定是否正常安排窗口上班。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    If mstrInNO <> "" Then
        If HaveExecute(1, mstrInNO, IIf(mbytInFun = 2, 2, 1)) Then
            MsgBox "该单据包含完全执行或部分执行的项目,不允许修改。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckChargeDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 Public Sub SendMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消息发送处理
    '入参: 0-收费划价单;1-门诊收费单;2-记帐划价单;3-记帐单
    '     strNO-单据号
    '编制:刘兴洪
    '日期:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, objDrugXML As New clsXML, objCheckXML As New clsXML
    Dim objTemp As clsXML, str收费时间 As String
    Dim rsTemp As ADODB.Recordset, int性质 As Integer
    Dim bln直接收费 As Boolean, p As Long
    Dim lngDrug As Long, lngCheck As Long, blnAddBill As Boolean, blnHaveCheck As Boolean, blnHaveDrug As Boolean
    
    'mbytInFun:0-收费,1-划价,2-门诊记帐
    '  mbytInState  :0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费),4-重新收费;5-异常单据作废
    On Error GoTo errHandle
    
    
    If Not (mbytInState = 0 Or mbytInState = 5) Then Exit Sub
    If mobjMsgModule Is Nothing Then Exit Sub
    If mobjMsgModule.IsConnect = False Then Exit Sub
    
    If Format(mobjBill.登记时间, "yyyy") < 2000 Then mobjBill.登记时间 = zlDatabase.Currentdate
    str收费时间 = mobjBill.登记时间
    
    
    'ZLHIS_CHARGE_003 门诊费用单据
    '节点名称    属性    含义    重复    类型    缺省值  值域描述
    'patient_info        病人信息    1
    '   patient_id      病人id  1   N
    '   patient_name        姓名    1   S
    '   patient_sex     性别    1   S
    '   patient_age     年龄    1   S
    '   identity_card       身份证号    0..1    S
    '   in_number       住院号  0..1    S
    '   out_number      门诊号  0..1    S
    'charge_bill         1..*
    '   bill_no     单据号码    1   S
    '   bill_kind       单据性质    1   N       1-收费单;2-记帐单
    '   charge_state        收费状态    1   N       1-未收费;2-已收费
    '   charge_time     收费时间    1   S
    '   charge_person       收费人员    1   S
    '   bill_item           1..*
    '       charge_item_id      收费项目id  1   N
    '       charge_item_kind        收费类别    1   S
    '       execute_dept_id     执行部门id  1   N
    '       drug_window     发药窗口    0..1    S
    objDrugXML.ClearXmlText
    objCheckXML.ClearXmlText
    blnHaveCheck = False: blnHaveDrug = False
    For p = 1 To mobjBill.Pages.Count
    
        If mobjBill.Pages(p).NO = "" Then
            bln直接收费 = True
        Else
            bln直接收费 = False
        End If
        
        If p = 1 Then
            '药品
            Call objDrugXML.AppendNode("patient_info")
                Call objDrugXML.appendData("patient_id", mobjBill.病人ID)
                Call objDrugXML.appendData("patient_name", mobjBill.姓名)
                Call objDrugXML.appendData("patient_sex", mobjBill.性别)
                Call objDrugXML.appendData("patient_age", mobjBill.年龄)
                '身份证号和住院号暂不传(意义不大)
                Call objDrugXML.appendData("out_number", mobjBill.标识号)
            Call objDrugXML.AppendNode("patient_info", True)
            '检查
            Call objCheckXML.AppendNode("patient_info")
                Call objCheckXML.appendData("patient_id", mobjBill.病人ID)
                Call objCheckXML.appendData("patient_name", mobjBill.姓名)
                Call objCheckXML.appendData("patient_sex", mobjBill.性别)
                Call objCheckXML.appendData("patient_age", mobjBill.年龄)
                '身份证号和住院号暂不传(意义不大)
                Call objCheckXML.appendData("out_number", mobjBill.标识号)
            Call objCheckXML.AppendNode("patient_info", True)
        End If
        
        If bln直接收费 Then
          '针对划价单进行收费的
          lngDrug = 1: lngCheck = 1
          
          For Each mobjBillDetail In mobjBill.Pages(p).Details
            
              blnAddBill = False
              If InStr(1, ",5,6,7,", "," & mobjBillDetail.收费类别 & ",") > 0 _
                And Not gbln收费后自动发药 Then
                '不含自动发药
                  '药品
                  Set objTemp = objDrugXML
                  If lngDrug = 1 Then blnAddBill = True
                  blnHaveDrug = True
                  lngDrug = lngDrug + 1
                  
              ElseIf InStr(1, ",D,", "," & mobjBillDetail.收费类别 & ",") > 0 Then
                  '检查
                  Set objTemp = objCheckXML
                  If lngCheck = 1 Then blnAddBill = True
                  lngCheck = lngCheck + 1
                  blnHaveCheck = True
              Else
                  Set objTemp = Nothing
              End If
              
              If Not objTemp Is Nothing Then
                If blnAddBill Then
                    Call objTemp.AppendNode("charge_bill")
                    Call objTemp.appendData("bill_no", mobjBill.Pages(p).收费单号)
                    If mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                        '门诊划价(收费)
                        Call objTemp.appendData("bill_kind", 1)  '1-收费单;2-记帐单
                        Call objTemp.appendData("charge_state", 1)   '1-未收费;2-已收费
                    ElseIf mbytInFun = 2 Then
                        '门诊记帐
                        Call objTemp.appendData("bill_kind", 2)  '1-收费单;2-记帐单
                        Call objTemp.appendData("charge_state", IIf(mbytBilling = 1, 1, 2))  '1-未收费;2-已收费
                    Else
                        Call objTemp.appendData("bill_kind", 1)  '1-收费单;2-记帐单
                        Call objTemp.appendData("charge_state", 2)   '1-未收费;2-已收费
                    End If
                    Call objTemp.appendData("charge_time", str收费时间)
                    Call objTemp.appendData("charge_person", UserInfo.姓名)
                End If
                '----------------------------------------------------------------------------
                '明细项
                objTemp.AppendNode ("bill_item")
                '       charge_item_id      收费项目id  1   N
                    Call objTemp.appendData("charge_item_id", mobjBillDetail.收费细目ID)
                '       charge_item_kind        收费类别    1   S
                    Call objTemp.appendData("charge_item_kind", mobjBillDetail.收费类别)
                '       execute_dept_id     执行部门id  1   N
                    Call objTemp.appendData("execute_dept_id", mobjBillDetail.执行部门ID)
                '       drug_window     发药窗口    0..1    S
                    Call objTemp.appendData("drug_window", mobjBillDetail.发药窗口)
                Call objTemp.AppendNode("bill_item", True)
              End If
          Next
        End If
        If Not bln直接收费 Then
            '划价单,审核单
            strSQL = "" & _
            "   Select NO,收费类别,收费细目ID,执行部门ID,发药窗口,登记时间,操作员姓名" & _
            "   From 门诊费用记录 " & _
            "   Where NO=[1] And 记录性质=[2] And  记录状态=1 " & _
            "   Order by 收费类别"
            If mbytInFun = 2 Then
                int性质 = 2
            Else
                int性质 = 1
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO, int性质)
            If rsTemp.EOF Then Exit Sub
            
            lngDrug = 1: lngCheck = 1
            Do While Not rsTemp.EOF
                 blnAddBill = False
                If InStr(1, ",5,6,7,", "," & rsTemp!收费类别 & ",") > 0 Then
                    '药品
                    Set objTemp = objDrugXML
                    If lngDrug = 1 Then blnAddBill = True
                    blnHaveDrug = True
                    lngDrug = lngDrug + 1
                ElseIf InStr(1, ",D,", "," & rsTemp!收费类别 & ",") > 0 Then
                    '检查
                    Set objTemp = objCheckXML
                    If lngCheck = 1 Then blnAddBill = True
                    lngCheck = lngCheck + 1
                    blnHaveCheck = True
                Else
                    Set objTemp = Nothing
                End If
                
                If Not objTemp Is Nothing Then
                  If blnAddBill Then
                        Call objTemp.AppendNode("charge_bill")
                        Call objTemp.appendData("bill_no", Nvl(rsTemp!NO))
                        If mbytInFun = 2 Then
                            '门诊记帐
                            mbytBilling = 1
                            Call objTemp.appendData("bill_kind", 2)  '1-收费单;2-记帐单
                            Call objTemp.appendData("charge_state", 2)   '1-未收费;2-已收费
                        Else
                            Call objTemp.appendData("bill_kind", 1)  '1-收费单;2-记帐单
                            Call objTemp.appendData("charge_state", 2)   '1-未收费;2-已收费
                        End If
                      
                        Call objTemp.appendData("charge_time", Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS"))
                        Call objTemp.appendData("charge_person", Nvl(rsTemp!操作员姓名))
                  End If
                  '----------------------------------------------------------------------------
                  '明细项
                  Call objTemp.AppendNode("bill_item")
                  '       charge_item_id      收费项目id  1   N
                      Call objTemp.appendData("charge_item_id", Val(Nvl(rsTemp!收费细目ID)))
                  '       charge_item_kind        收费类别    1   S
                      Call objTemp.appendData("charge_item_kind", Nvl(rsTemp!收费类别))
                  '       execute_dept_id     执行部门id  1   N
                      Call objTemp.appendData("execute_dept_id", Nvl(rsTemp!执行部门ID))
                  '       drug_window     发药窗口    0..1    S
                      Call objTemp.appendData("drug_window", Nvl(rsTemp!发药窗口))
                  Call objTemp.AppendNode("bill_item", True)
               End If
            rsTemp.MoveNext
          Loop
        End If
        If lngDrug > 1 Then Call objDrugXML.AppendNode("charge_bill", True)
        If lngCheck > 1 Then Call objCheckXML.AppendNode("charge_bill", True)
    
    Next
     
    If blnHaveDrug = True _
        And Not gbln收费后自动发药 Then
        '不含自动发药
        '发药品消息
        Call zlDebugWriteFile(objDrugXML.XmlText)
        Call mobjMsgModule.CommitMessage("ZLHIS_CHARGE_003", objDrugXML.XmlText)
    End If
    If blnHaveCheck Then
        '发检查消息
        Call zlDebugWriteFile(objCheckXML.XmlText)
        Call mobjMsgModule.CommitMessage("ZLHIS_CHARGE_003", objCheckXML.XmlText)
    End If
    objDrugXML.ClearXmlText: objCheckXML.ClearXmlText
    Set objDrugXML = Nothing: Set objCheckXML = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Private Sub CreateDrugPacker()
    '功能:创建自助发药机(自动化药房)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugPacker = False: mblnDrugMachine = False

    If Not (mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 _
        Or mbytInState = 4) Or mbytInFun = 2) Then Exit Sub

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    If mblnDrugMachine = False Then
        '旧部件
        Err = 0
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err = 0 Then mblnDrugPacker = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        '权限检查
        strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
        If InStr(";" & strPrivs & ";", ";基本;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    ElseIf mblnDrugPacker Then

        mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
    End If
End Sub
Private Function ReadDrugAndStuffStock(ByVal lng库房ID As Long, ByRef objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取药品和卫材料的库存信息
    '入参:lng库房ID-库房ID
    '出参:objDetail-Detail对象
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-10 09:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblStock As Double, str药房IDs As String
    
    On Error GoTo errHandle
    If objDetail Is Nothing Then Exit Function
    If InStr(",5,6,7,4,", objDetail.类别) = 0 Then ReadDrugAndStuffStock = True: Exit Function
    If objDetail.类别 = "4" And objDetail.跟踪在用 = False Then ReadDrugAndStuffStock = True: Exit Function
   
    If objDetail.类别 = "4" And objDetail.跟踪在用 Then
        dblStock = GetStock(objDetail.ID, lng库房ID, objDetail.批次)
        objDetail.库存 = dblStock
        Call ShowStock(lng库房ID, objDetail.名称, objDetail.库存)
        ReadDrugAndStuffStock = True: Exit Function
    End If
    
    '当前行药品库存
    If InStr(",5,6,7,", objDetail.类别) > 0 Then
        dblStock = GetStock(objDetail.ID, lng库房ID)
        If gbln药房单位 Then dblStock = dblStock / objDetail.药房包装
        objDetail.库存 = dblStock
        Call ShowStock(lng库房ID, objDetail.名称, objDetail.库存)
    End If
    ReadDrugAndStuffStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
