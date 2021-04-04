VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "病人收费管理"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmClinicCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra退费摘要 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   15
      TabIndex        =   74
      Top             =   5160
      Visible         =   0   'False
      Width           =   7035
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
      TabIndex        =   64
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
         TabIndex        =   73
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
         TabIndex        =   68
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
         TabIndex        =   66
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
         TabIndex        =   65
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
      TabIndex        =   61
      Top             =   1830
      Width           =   11820
      Begin VB.CommandButton cmdDelBill 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
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
            Size            =   10.5
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
         _ExtentY        =   1244
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
      TabIndex        =   41
      ToolTipText     =   "清除:F6"
      Top             =   -120
      Width           =   11880
      Begin VB.TextBox txtIn 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4212
         MaxLength       =   8
         TabIndex        =   35
         ToolTipText     =   "从已有的单据中复制信息,不影响已有单据"
         Top             =   660
         Width           =   1065
      End
      Begin VB.CommandButton cmdSaveWholeSet 
         Caption         =   "保存为成套收费项目(&W)"
         Height          =   375
         Left            =   6630
         TabIndex        =   76
         Top             =   195
         Width           =   2715
      End
      Begin VB.CommandButton cmdSelWholeSet 
         Caption         =   "成套(&T)"
         Height          =   375
         Left            =   5505
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "医保"
         Height          =   375
         Left            =   1080
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "热键：F6"
         Top             =   660
         Width           =   720
      End
      Begin VB.CommandButton cmdIDCard 
         Caption         =   "医疗卡(&K)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   67
         ToolTipText     =   "热键：F10"
         Top             =   195
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdRegist 
         Caption         =   "挂号(&E)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10725
         TabIndex        =   38
         ToolTipText     =   "热键：F3"
         Top             =   195
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmd配方 
         Caption         =   "配方(&R)"
         Height          =   375
         Left            =   80
         TabIndex        =   31
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
         TabIndex        =   36
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
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   645
         Width           =   435
      End
      Begin VB.TextBox txtMCInvoice 
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtRePrint 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2500
         MaxLength       =   8
         TabIndex        =   33
         Top             =   667
         Width           =   1065
      End
      Begin VB.Label lblIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导(&I)"
         Height          =   240
         Left            =   3588
         TabIndex        =   34
         Top             =   732
         Width           =   600
      End
      Begin VB.Label lblRePrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打(&P)"
         Height          =   240
         Left            =   1900
         TabIndex        =   32
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
         TabIndex        =   62
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
         TabIndex        =   39
         Top             =   720
         Width           =   480
      End
      Begin VB.Line linTopSplitW 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   38015
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line linTopSplitG 
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
         TabIndex        =   49
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
         TabIndex        =   44
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
         TabIndex        =   42
         Top             =   720
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   43
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
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmClinicCharge.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   2
            Object.Tag             =   "用于记帐或收费个人帐户显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   2
            Object.Tag             =   "用于收费预交显示"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2
            MinWidth        =   2
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Picture         =   "frmClinicCharge.frx":115E
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
            Picture         =   "frmClinicCharge.frx":1478
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmClinicCharge.frx":1B52
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmClinicCharge.frx":218C
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
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
   Begin VB.Frame fraInfo 
      Height          =   990
      Left            =   0
      TabIndex        =   40
      Top             =   840
      Width           =   11880
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   390
         Left            =   555
         TabIndex        =   78
         Top             =   180
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   688
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
         TabIndex        =   69
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
         TabIndex        =   70
         Top             =   630
         Width           =   2880
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   240
         Left            =   8910
         TabIndex        =   63
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl动态费别 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5520
         TabIndex        =   59
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
         TabIndex        =   48
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Left            =   2680
         TabIndex        =   47
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Left            =   4395
         TabIndex        =   46
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   240
         Left            =   3240
         TabIndex        =   45
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
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5505
      Width           =   11280
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1770
         Left            =   5415
         TabIndex        =   77
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
         FormatString    =   $"frmClinicCharge.frx":27C6
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
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   29
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
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   30
         ToolTipText     =   "热键:Esc"
         Top             =   1410
         Width           =   1440
      End
      Begin VB.CommandButton cmd预结算 
         Caption         =   "预结算(&V)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10305
         TabIndex        =   27
         ToolTipText     =   "热键：F5"
         Top             =   540
         Visible         =   0   'False
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   570
         Visible         =   0   'False
         Width           =   720
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
         Height          =   585
         Left            =   0
         TabIndex        =   51
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
            TabIndex        =   21
            Text            =   "cbo开单人"
            ToolTipText     =   "支持输入简码和编号自动匹配"
            Top             =   165
            Width           =   2145
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9390
            TabIndex        =   22
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
            TabIndex        =   20
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   240
            Left            =   8880
            TabIndex        =   52
            Top             =   225
            Width           =   480
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1770
         Left            =   15
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   510
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3122
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
            Size            =   10.5
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
         TabIndex        =   53
         Top             =   375
         Width           =   2490
         Begin VB.TextBox txt合计 
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
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   735
            Locked          =   -1  'True
            TabIndex        =   24
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
               Size            =   14.25
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
            TabIndex        =   23
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
               Size            =   14.25
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
            Top             =   1350
            Width           =   1650
         End
         Begin VB.Label lbl合计 
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
            Left            =   60
            TabIndex        =   56
            Top             =   885
            Width           =   660
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
            Left            =   60
            TabIndex        =   55
            Top             =   345
            Width           =   690
         End
         Begin VB.Label lbl累计 
            AutoSize        =   -1  'True
            Caption         =   "累计"
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
            Left            =   60
            TabIndex        =   54
            Top             =   1410
            Width           =   690
         End
      End
      Begin MSComctlLib.ImageList imgPati 
         Left            =   4875
         Top             =   1875
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicCharge.frx":2814
               Key             =   "InPati"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicCharge.frx":30EE
               Key             =   "OutPati"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   28
         ToolTipText     =   "热键F2,右键弹出保存为划价单(或按CTRL+S)"
         Top             =   975
         Width           =   1440
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
         TabIndex        =   60
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
         TabIndex        =   57
         Top             =   585
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   120
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2925
      Left            =   -15
      TabIndex        =   14
      Top             =   2220
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5159
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
      cboStyle        =   0
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
Attribute VB_Name = "frmClinicCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private Const M_MONEY_ROWS = 6 '左下角项目列表可显示行数
Public Enum gEM_ChargeEditType
    EM_ED_收费 = 0
    EM_ED_浏览 = 1
    EM_ED_调整 = 2
    EM_ED_异常重收 = 4
    EM_ED_异常作废 = 5
End Enum
'————————————————————————————————————————————————————————————————————
'入口参数：
Private mfrmMain As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mbytInState As gEM_ChargeEditType '0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费),4-重新收费;5-异常单据作废
Private mstrInNO As String '操作的单据号(查看，调整，修改，退费，销帐,重新收费时)(暂未用)
Private mlng结帐ID  As Long '一次结算:操作的单据号

Private mblnNOMoved As Boolean '操作的单据是否在后备数据表中
Private mstrTime As String '操作单据内容的登记时间
Private mblnDelete As Boolean '是否处理退费单据(查阅)
Private mlngFirstID As Long '记录被修改单据第一药品行的执行部门ID,用于收费
Private mstrFirstWin As String '记录被修改单据第一药品行的发药窗口,用于收费
Private mbln作废异常 As Boolean '异常冲销单据
'消息相关对象变量
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1
Private mblnErrBill As Boolean  '收费界面时，是否提取的是异常单据
Private mblnElsePersonErrBill As Boolean '是否是他人的异常单据
'----------------------------------------------------------------------------------------------------------------------------------------
Private mrs结算方式 As ADODB.Recordset
Private mrs缺省结算方式 As ADODB.Recordset
Private mobjChargeInfor As clsClinicChargeInfor
Private mstr应付款结算方式 As String    '33722
Private mblnSaveAsPrice As Boolean '联合医保：收费时是否保存为划价单
Private mintReturnMode As Integer   '用于退费时,全退禁用结算方式时恢复初始的结算方式
Private mblnNotValied As Boolean '不处理效点失效问题
Private mblnNotClick As Boolean
Private mstrBalance As String
Private mblnHaveExcuteData As Boolean '是否医嘱计价中存在数据:60735
'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsWork As ADODB.Recordset      '当天上班的药房
Private mrsClass As ADODB.Recordset     '根据参数读取的当前可用的收费类别
Private mrsUnit As ADODB.Recordset      '可选择的执行科室
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室
Private mrs开单人 As ADODB.Recordset    '所用医生和护士列有
Private mrsInfo As ADODB.Recordset      '病人信息
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
Private mobjDrugPacker As Object '自动发药机
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object '自动发药机(新）
Private mblnDrugMachine As Boolean

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
    "应收金额,1030,7;实收金额,1080,7;执行科室,1255,1;标志,520,4;医嘱序号,0,0;类型,520,1"

'医保相关
Private mintInsure As Integer
Private mstrYBPati As String 'New:空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
Private mstr个人帐户 As String '是否将个人帐户设置到收费可用
Private mdbl个帐余额 As Double   '当前病人个人帐户余额
Private mdbl个帐透支 As Double   '个人帐户允许透支金额

Private mblnYB结算作废 As Boolean '医保是否支持结算作废,用于退费时判断
Private mstrYBBill As String '医保病人连续收费的单据号
Private mlng结算序号  As Long '重新收费时有效
Private mrsDelInvoice As ADODB.Recordset
Private mblnOneCard As Boolean      '是否启用了一卡通接口
Private mrsOneCard As ADODB.Recordset

'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    允许不设置医保项目 As Boolean
    门诊收费存为划价单 As Boolean
    不提醒缴款金额不足 As Boolean    '27536
    医保接口打印票据 As Boolean
    门诊连续收费 As Boolean
    门诊预结算 As Boolean
    多单据收费 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    先自付 As Boolean
    全自付 As Boolean
    blnOnlyBjYb As Boolean '本地仅支持北京医保:刘兴洪
    医保不走票号  As Boolean        '预结算时有效
    多单据分单据结算 As Boolean '86321
    门诊结算作废 As Boolean
    一次结算分单据退费 As Boolean '91602
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
Private mblnNotClearLedDisplay As Boolean   '不清除显示
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
Private mlngPreBrushCardID As Long  '上次刷卡的卡类别ID
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'-----------------------------------------------------------------------------------
'数据保存相关
Private mstrModiNOs As String, mstrSaveNos As String
Private mCllWindows As Collection
Private mblnClearBlance As Boolean '是否清除结算信息
Private mlngCardTypeID As Long   '当前提取病人信息刷的卡类别ID 56615
Private mblnOlny预交 As Boolean '仅使用预交68177

Private mintSucces As Integer '收费成功次数
Private mdbl应缴合计 As Double

'结算窗口
Private mFrmBalanceWin   As frmClinicChargeBalance
Attribute mFrmBalanceWin.VB_VarHelpID = -1
Private mblnPeisPriceBill As Boolean '102660,当前病人是否存在体检单据
Private mstrTittle As String '窗体标题
Private mstr药品价格等级 As String, mstr卫材价格等级 As String, mstr普通价格等级 As String

Public Function zlEditBill(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal bytInState As gEM_ChargeEditType, _
    Optional ByVal lng结帐ID As Long, Optional ByVal lng结算序号 As Long, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal strTime As String, Optional ByVal blnDelete As Boolean, _
    Optional objMsgModule As clsMipModule, Optional strInNO As String, _
    Optional ByVal bln作废异常 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊收费的程序入口(收费,查看,异常重收,异常作废)
    '入参:frmMain-调用者主窗体
    '     lngModule-模块号
    '     strPrivs-权限串
    '     bytInState-操作功能(0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费),4-重新收费;5-异常单据作废)
    '     strInNO-操作的单据号( 调整时传入)
    '     blnNoMoved-操作的单据是否在后备数据表中
    '     strTime-操作单据内容的登记时间
    '     blnDelete-是否处理退费单据(查阅)
    '     objMsgModule-消息相关对象变量
    '     bln作废异常-异常收费单作废后的异常单据
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 11:06:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    'Load 事件存在数据判断,可能窗体显示时出错(避免窗体闪动后立即关闭)
    Set mfrmMain = frmMain: mlngModul = lngModule: mbytInState = bytInState
    mlng结帐ID = lng结帐ID: mblnNOMoved = blnNOMoved: mstrTime = strTime
    mlng结算序号 = lng结算序号: mstrInNO = strInNO
    mbln作废异常 = bln作废异常
    mblnDelete = blnDelete: Set mobjMsgModule = objMsgModule
    mintSucces = 0: mstrPrivs = strPrivs
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    zlEditBill = mintSucces > 0
End Function

Private Sub cbo医疗付款_Click()
    On Error GoTo errHandler
    If mbytInState <> EM_ED_收费 Then Exit Sub
    If gintPriceGradeStartType < 2 Then Exit Sub
    
    If mrsInfo.State = adStateOpen Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), zlStr.NeedName(cbo医疗付款.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    Else
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cbo医疗付款.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    End If
    
    If mbln不重算价格 Then Exit Sub
    If CheckBillsEmpty Then Exit Sub
    
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

Private Sub Form_Load()
    mstrTittle = "病人收费管理"
    
    mblnFirst = True: mbln连续输入 = False
    mblnHaveExcuteData = False
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
    End If
    
    '最小窗体尺寸
    glngFormW = 12000: glngFormH = 7710
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
        Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    '应该放在限制尺寸之后
    RestoreWinState Me, App.ProductName, mstrTittle & "_" & mbytInState
    sta.Visible = True
    
    '----------------------------变量及对象初始化------------------------------
    Call InitLed    '初始化Led
    Call CreateDrugPacker '创建自动化药房部件
    Call ClearTotalInfo(True)
    
    lblSub应收.Caption = "应收:" & gstrDec
    lblSub实收.Caption = "实收:" & gstrDec
    lblAmount.Caption = ""
    
    '模块变量
    Call InitCommVariable
    
    gbln处方限量 = False
    mblnLoad = False:           mblnDoing = False
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
    
    If CheckDepend = False Then Unload Me: Exit Sub
    
    '-------------------------数据初始及加载------------------------------------
    '查看功能时，无需初始数据
    Select Case mbytInState
    Case EM_ED_收费, EM_ED_调整, EM_ED_异常重收, EM_ED_异常作废
        If mbytInState = EM_ED_收费 Then
            mstr药品价格等级 = gstr药品价格等级
            mstr卫材价格等级 = gstr卫材价格等级
            mstr普通价格等级 = gstr普通价格等级
        End If
        If Not InitData Then Unload Me: Exit Sub
    Case Else
        '年龄单位
        cbo年龄单位.AddItem "岁"
        cbo年龄单位.AddItem "月"
        cbo年龄单位.AddItem "天"
        cbo年龄单位.ListIndex = 0
    End Select
    Call InitFace   'InitData需要在此之前
End Sub
Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "执行科室"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "发药药店"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    
    If mobjBill.Pages(mintPage).Details.Count < Bill.Row Then Exit Sub
     
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
End Sub

Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lng执行科室ID As Long
    Dim rsTemp As ADODB.Recordset, dbl价格 As Double
    Dim strSQL As String
    Dim dbl数次 As Double, dbl单价 As Double
    
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
   If mbytInState = EM_ED_浏览 Then Exit Sub
 
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
    Dim tmpBill As New ExpenseBill, byt婴儿费 As Byte, dtCurdate As Date
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
                         
    Set tmpBill = ImportWholeSet(Me, IIf(mstrYBPati <> "", mintInsure, 0), rsSel, mlng西药房, mlng成药房, mlng中药房, lng病人ID, 0, gbln药房单位, lng开单部门ID, byt婴儿费, 2, chk加班.Value = 1, _
        0, gint病人来源, UserInfo.姓名, zlStr.NeedName(cbo开单人.Text), mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    
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
        If Not zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then mobjBill.Pages(mintPage).开单人 = ""
        '清除病人信息
       ' Call ClearmobjBill
    Else
        'b.多张单据模块,新增单据,保留当前单据内容及病人相关信息,
        '78566,冉俊明,2014-10-13,最后一张单据为划价单时也要新增单据
        If i > 0 Or mobjBill.Pages(mintPage).Details.Count > 0 Or mobjBill.Pages(mintPage).NO <> "" Then
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
            If zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then .开单人 = tmpBill.Pages(1).开单人
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
    If mbytInState = EM_ED_收费 And mstrInNO <> "" Then mstrInNO = ""
        
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
          '  Call SetOneCardBalance
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

Public Sub zlExeBalanceWinRefrshData(ByVal blnSaveOK As Boolean, ByVal bytExitMode As gExitMode, _
    ByVal bln继续输入 As Boolean, _
    ByRef objChargeInfor As clsClinicChargeInfor)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行结算操作后的刷新操作
    '入参:blnSaveOK-是否保存成功
    '     bytExitMode-当前退出模式
    '     bln继续输入-继续输入
    '     objChargeInfor-结算信息
    '编制:刘兴洪
    '日期:2014-06-17 10:50:41
    '说明:之所要独立出来,主要原因是解决医保调试的问题(模态窗体不好调试)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln连续 As Boolean, i As Long, p As Long
    Dim blnGetFact As Boolean, strReturn As String
    Dim strData As String
    
    On Error GoTo errHandle
    If mbytInState = EM_ED_异常作废 Or bytExitMode = EM_EX_作废 Then
        If blnSaveOK Then
            mblnSaveData = True: mintSucces = mintSucces + 1
        End If
        mlng结算序号 = 0: Unload Me
        Exit Sub
    End If
    If mbytInState = EM_ED_异常重收 Or mblnErrBill Then
        If Not blnSaveOK Then Unload Me: Exit Sub
        '显示Led相关信息
        'LED显示:(合计,)发药窗口
        If gblnLED And CCur(txt合计.Text) <> 0 And (mstr西窗 <> "" Or mstr中窗 <> "" Or mstr成窗 <> "") Then
            zl9LedVoice.DisplayBank "费用合计:" & txt合计.Text, _
                "取药窗口:" & IIf(mstr西窗 <> "", " " & mstr西窗, "") & _
                IIf(mstr成窗 <> "", " " & mstr成窗, "") & IIf(mstr中窗 <> "", " " & mstr中窗, "")
        End If
        Call CheckBillNOAndBookeFee(True)
        '打印票据
        Call PrintBill(objChargeInfor.Nos, "")
        
        If mblnDrugMachine Then
            '门诊格式：1|单据1,处方号1;单据2,处方号2
            strData = "1|" & "8," & Replace(Replace(objChargeInfor.Nos, "'", ""), ",", ";8,")
            Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
        ElseIf mblnDrugPacker Then
            '51510
            strData = "8," & Replace(Replace(objChargeInfor.Nos, "'", ""), ",", "|8,")
            Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.编号, UserInfo.姓名, 0, strData, strReturn)
        End If
        
        '81688:李南春,2015/5/18,评价器
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.OutPatiExseAfter(objChargeInfor.病人ID, objChargeInfor.结帐ID)
            Err.Clear
        End If
        If Not mblnErrBill Then Unload Me
        mblnSaveData = True
        mintSucces = mintSucces + 1
        Exit Sub
    End If
    
    
    If Not blnSaveOK Then
        '保存不成功,收费,保存单据失败后的处理
         If bytExitMode <> EM_EX_作废 And bytExitMode <> EM_EX_退出 Then
             Call ShowBillChargeFee(objChargeInfor.结帐ID)
         End If
         
        mlng结算序号 = 0
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        If mintInsure <> 0 Then
            cmdAddBill.Enabled = Not MCPAR.门诊连续收费 And MCPAR.多单据收费 _
                And zlStr.IsHavePrivs(mstrPrivs, "医保病人多单据收费")
        Else
            cmdAddBill.Enabled = zlStr.IsHavePrivs(mstrPrivs, "普通病人多单据收费")
        End If
        
        If cmdDelBill.Visible And tbsBill.Tabs.Count > 1 Then cmdDelBill.Enabled = True
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        If bytExitMode = EM_EX_作废 Then
            If mblnAutoChangePati And gint病人来源 = 2 Then
                '需要切找到病人来源1中
                gint病人来源 = 1: zlChangePatiSource (gint病人来源)
            End If
            Call ClearFullBill(False)
        End If
        If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
        If gfrmMain Is Nothing Then Me.Enabled = True
        Exit Sub
    End If
    mblnSaveData = True
    mintSucces = mintSucces + 1
    
    bln连续 = bln继续输入
    '收费操作成功
    If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
    If gfrmMain Is Nothing Then Me.Enabled = True
    
    
    '设置应缴累计
    Call Set应缴累计(bln连续)
    If mblnDrugMachine Then
        '门诊格式：1|单据1,处方号1;单据2,处方号2
        strData = "1|" & "8," & Replace(Replace(objChargeInfor.Nos, "'", ""), ",", ";8,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    ElseIf mblnDrugPacker Then
        '51510
        strData = "8," & Replace(Replace(objChargeInfor.Nos, "'", ""), ",", "|8,")
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.编号, UserInfo.姓名, 0, strData, strReturn)
    End If
    
    '消息发送
    Call SendMsgModule
    
    mlng结算序号 = 0
    '显示Led:发药窗口及费用合计金额
    Call ShowLedWinAndSum
    
    Call zlChargeSaveAfter_Plugin(glngModul, mobjBill.病人ID, mobjBill.主页ID, True, 1, mobjChargeInfor.Nos)
    '票据打印,打印票据
    Call PrintBill(objChargeInfor.Nos, "")
    '81688:李南春,2015/5/18,评价器
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.OutPatiExseAfter(objChargeInfor.病人ID, objChargeInfor.结帐ID)
        Err.Clear
    End If
    
    '设置其他相关内容
    '防止设置打印机弹出的非模态窗体,以及医保延时
    '写卡:56615
    Call WriteMzInforToCard(objChargeInfor.病人ID, objChargeInfor.结帐ID)
    
    cmdOK.Enabled = True: cmdCancel.Enabled = True
    If cmd预结算.Visible Then cmd预结算.Enabled = True
    If mbytInState = EM_ED_收费 And gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    
    '新增,或新增界面通过输入单据号修改单据
    sta.Panels(Pan.C2提示信息) = "上一张单据:" & mobjBill.NO '多单据时为第一张
    
    i = UBound(Split(objChargeInfor.Nos, ",")) + 1
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
            Exit Sub
        End If
    End If
 
    mstrInNO = "":  mlngFirstID = 0: mstrFirstWin = ""
    
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
    Call ClearBillRows
    If (mstrYBPati <> "" And MCPAR.门诊连续收费) Then
        Call NewYBBill
        mobjBill.病人ID = CLng(Split(mstrYBPati, ";")(8))
        
        '重新读取预交余额
        Call LoadFeeInfor(mobjBill.病人ID)
        '重新读取个帐余额
        Dim cur个帐透支 As Currency
        cur个帐透支 = RoundEx(mdbl个帐透支, 2)
        mdbl个帐余额 = gclsInsure.SelfBalance(mobjBill.病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, cur个帐透支, mintInsure)
        mdbl个帐透支 = cur个帐透支
        sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mdbl个帐余额, "0.00")
        sta.Panels(Pan.C3个人帐户).Visible = True
        
        mstrYBPati = ""
    Else
        Call NewBill(blnGetFact, Not Bill.Active)        '划价单时不更改费别
        Call SetDisible(True)
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
    

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    
    If gblnLED = False Or gblnLedDispDetail = False Then Exit Sub
    If Not mbytInState = EM_ED_收费 Then Exit Sub
    
    'LED动态显示项目
    If mobjBill.Pages(mintPage).Details.Count >= Row - 1 Then
        With mobjBill.Pages(mintPage).Details(Row - 1)
            dbl单价 = 0: cur金额 = 0
            For i = 1 To .InComes.Count
                cur金额 = cur金额 + .InComes(i).实收金额
                dbl单价 = dbl单价 + .InComes(i).标准单价
            Next
            dbl单价 = RoundEx(dbl单价, 6)
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
End Sub

Private Sub ShowGroupLED(ByVal lngMain As Long, ByVal lngBegin As Long, ByVal lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:为加快速度，一次性调用套餐项目的LED显示
    '入参:行号范围:
    '     lngMain=主项行号,
    '     lngBegin-lngEnd:从项行号
    '     lngEnd-结束行号
    '编制:刘兴洪
    '日期:2014-06-05 15:55:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl数量 As Double, dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    If gblnLED = False Or gblnLedDispDetail = False Then Exit Sub
    If Not mbytInState = EM_ED_收费 Then Exit Sub
      
    
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
End Sub


Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytSubs As Byte
    Dim bln从项汇总折扣 As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> EM_ED_收费 Or chkCancel.Value = 1 Then Cancel = True: Exit Sub
    
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
                        dblStock = GetStock(.收费细目ID, .执行部门ID)
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
                        If mintInsure <> 0 And MCPAR.实时监控 And mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                            If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = str执行科室: .执行部门ID = lng执行科室
                                Exit Sub
                            End If
                        End If
                        
                        If mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
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

Public Function GetOriginalTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long, _
    Optional ByVal intPage As Integer) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据中指定药品在同一药房多行的原始数量和
    '入参: lng药房ID-0表示分离发药时,不限定药房检查
    '出参:
    '返回:成功,返回原始数量和,否则返回0
    '编制:刘兴洪
    '日期:2014-06-05 15:59:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    GetOriginalTotal = RoundEx(dblCount, 6)
End Function

Private Sub Bill_CellCheck(Row As Long, Col As Long)
    Dim i As Long, strCheck As String, bytTime As Byte
    Dim blnReSet As Boolean '重新设置
    Dim bln固定 As Boolean, strErrMsg As String, varData As Variant ' (0-医嘱序号;1-收费细目ID)
    Dim varTemp As Variant
    Dim bln固定1 As Boolean
    Dim j As Long
    
    If chkCancel.Visible And chkCancel.Value = 1 Then Exit Sub
     
    
    '说明:可以全部为主要手术,但不能全部为附加手术
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
    Dim str排除类别 As String
    
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
    If mstrYBPati <> "" Then
        '刘兴洪:24862
        If zl_Check特准项目(gclsInsure, mintInsure, mobjBill.病人ID, True) Then str特准项目 = Get保险特准项目(mobjBill.病人ID, "A.ID")
    End If
    If zlCheckBill存在非散装草药(mintPage) = True Then
        mblnSelect = False: Exit Sub
    End If
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, gint病人来源, mintInsure, gbln药房单位, str类别, , , str特准项目, 0, _
        , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
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
Private Sub ShowStock(ByVal lng库房ID As Long, str药品 As String, dbl库存 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示药品或卫材的库存
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 16:09:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    Call zlInit缺省部门
    If zlStr.IsHavePrivs(mstrPrivs, "显示库存") Then
        If InStr(1, gstr所属部门ID & ",", "," & lng库房ID & ",") > 0 Or gbyt库存显示方式 <= 0 Then   '31936
                sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & dbl库存
        Else
                sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & IIf(dbl库存 > 0, "有", "无") & "库存."
        End If
    Else
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]" & IIf(dbl库存 > 0, "有", "无") & "库存."
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:处理单据输入
    '编制:刘兴洪
    '日期:2014-06-05 16:10:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng项目id As Long, str类别 As String, str特准项目 As String, bln护士 As Boolean
    Dim dblStock As Double, strScope As String
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency, cur余额 As Currency
    Dim blnInput As Boolean, str摘要 As String, lngOld付数 As Long
    Dim lngDoUnit As Long, lng病人科室ID As String, str药房IDs As String, i As Long, j As Long
    Dim colStock As Collection, str排除类别 As String
    Dim dblNum As Double, strPriceGrade As String
    
    
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
        If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
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
                        Set mobjDetail = GetInputDetail(Val(Bill.Text))
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
                        
                        If mstrYBPati <> "" Then
                            '刘兴洪:24862
                            If zl_Check特准项目(gclsInsure, mintInsure, mobjBill.病人ID, True) Then str特准项目 = Get保险特准项目(mobjBill.病人ID, "A.ID")
                        End If
                        If zlCheckBill存在非散装草药(mintPage) Then
                            '存在非散装的,界面中就不能进行录入
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                        lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, gint病人来源, mintInsure, gbln药房单位, _
                            str类别, Bill.Text, Bill.TxtHwnd, str特准项目, 0, str排除类别, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                        If lng项目id <> 0 Then
                            Set mobjDetail = GetInputDetail(lng项目id)
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
                                                                
                    '当前行药品或卫材库存
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
                    
                    If mintInsure <> 0 And MCPAR.实时监控 And mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                        If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            mobjBill.Pages(mintPage).Details.Remove Bill.Row '删除刚刚想要加入的费用行
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            mobjBill.Pages(mintPage).Details.Remove Bill.Row '删除刚刚想要加入的费用行
                            Bill.Text = "": Cancel = True: Exit Sub
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
                            '当前行药品或卫材库存
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
                            '计算并刷新该行
                            .Details(Bill.Row).付数 = Bill.Text
                            Call CalcMoneys(mintPage, Bill.Row)

                            '输了数量再改付数的，在这里重新检查，先输付数，再输数量的，在输数量后检查
                            If mintInsure <> 0 And MCPAR.实时监控 And .Details(Bill.Row).数次 <> 0 Then
                                If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    .Details(Bill.Row).付数 = lngOld付数
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            If .Details(Bill.Row).数次 <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
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
                        If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And zlStr.IsHavePrivs(mstrPrivs, "药品输入小数") = False Then
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
                        If zlStr.IsHavePrivs(mstrPrivs, "负数费用") = False Then
                            MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
                        ElseIf .Details(Bill.Row).Detail.分批 Then
                            MsgBox "分批药品不允许输入负数。", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).数次: Cancel = True: Exit Sub
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
                                
                                If colStock("_" & .执行部门ID) <> 0 And zlStr.IsHavePrivs(mstrPrivs, "不检查库存") = False And Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
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
                    Call CalcMoneys(mintPage, Bill.Row)
                    
                    '数据溢出检查(在已经算出该行费用但未显示前)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                        .Details(Bill.Row).数次 = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(mintPage, Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    If mintInsure <> 0 And MCPAR.实时监控 And .Details(Bill.Row).数次 <> 0 Then
                        If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            .Details(Bill.Row).数次 = dblPreTime
                            Call CalcMoneys(mintPage, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If .Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            .Details(Bill.Row).数次 = dblPreTime
                            Call CalcMoneys(mintPage, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
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
                        dblPreMoney = .Details(Bill.Row).InComes(1).标准单价
                                                
                        .Details(Bill.Row).InComes(1).标准单价 = Bill.Text '这种收费细目只能对应一个收入项目
                        Call CalcMoneys(mintPage, Bill.Row)

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
                                
                                If colStock("_" & .执行部门ID) <> 0 And zlStr.IsHavePrivs(mstrPrivs, "不检查库存") = False Then
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
                        If Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
                            If mintInsure <> 0 And MCPAR.实时监控 And mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                                If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            If mobjBill.Pages(mintPage).Details(Bill.Row).数次 <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        If CheckMainItem(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入收费项目后,加载当前收费项目的从属项目到费用集对象,并显示在单据控件中
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 16:13:04
    '调用者:Bill_KeyDown中输入项目后
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
                    
'            If mstrYBPati <> "" Then'90304
                'CalcMoney中先调用GetuItemInsure可能返回摘要
                str摘要 = mobjBill.Pages(mintPage).Details(Bill.Rows - 1).摘要
                 
                str摘要 = gclsInsure.GetItemInfo(mintInsure, mobjBill.病人ID, mcolDetails(i).ID, str摘要, 1)
                mobjBill.Pages(mintPage).Details(Bill.Rows - 1).摘要 = str摘要
'            End If
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当从项汇总折扣时,根据指定的主项的行ID的第一个收入项目重算主项的实收金额
    '入参: lngMainRow-主项行ID
    '     intpage -页号,默认为当前页mintpage
    '编制:刘兴洪
    '日期:2014-06-05 16:19:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim i As Long, j As Long
    Dim cur打折前应收合计 As Currency     '记录所有主从项的应收合计
    Dim cur打折后实收 As Currency
    Dim str费别 As String               '记录根据应收等确定的最优惠的费别
    
    If intPage = 0 Then intPage = mintPage

    With mobjBill.Pages(intPage)
        For i = lngMainRow To .Details.Count
            If i = lngMainRow Or .Details(i).从属父号 = lngMainRow Then
                For j = 1 To .Details(i).InComes.Count
                    cur打折前应收合计 = cur打折前应收合计 + .Details(i).InComes(j).应收金额
                Next
            End If
        Next
        '药品不支持主从项，所以无需传加班加价率等
        '打折后的实收金额仅算到主项的第一个收入项目上
        str费别 = IIf(glngSys Like "8??", mobjBill.费别, zlStr.TrimEx(mobjBill.费别 & "," & lbl动态费别.Tag, ","))
        
        cur打折后实收 = CCur(Format(ActualMoney(str费别, .Details(lngMainRow).InComes(1).收入项目ID, cur打折前应收合计, 0, 0, 0, 0), gstrDec))
        cur打折后实收 = cur打折后实收 - cur打折前应收合计 + .Details(lngMainRow).InComes(1).应收金额
        
        .Details(lngMainRow).InComes(1).实收金额 = Format(cur打折后实收, gstrDec)
        .Details(lngMainRow).费别 = str费别
        
        Call ShowDetails(lngMainRow)
    End With
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据主项执行科室的变化,刷新非药从项的执行科室
    '入参:lngRow-指定的行号
    '编制:刘兴洪
    '日期:2014-06-05 16:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
                        If mbytInState = EM_ED_收费 Then
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
    
    If Not Bill.Active Then
        '显示划价单摘要:医嘱内容
        If Not mbytInState = EM_ED_收费 Then Exit Sub
        
        If mobjBill.Pages(mintPage).NO <> "" And Bill.RowData(Bill.Row) <> 0 Then
            strTmp = Get费用摘要(mobjBill.Pages(mintPage).NO, 1, Bill.RowData(Bill.Row))
            If strTmp <> "" Then sta.Panels(Pan.C2提示信息) = "摘要:" & strTmp
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
        If mbytInState = EM_ED_收费 Then
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
                            If zlStr.IsHavePrivs(mstrPrivs, "显示库存") Then
                                sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行库存:" & strStock
                            Else
                                sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行有库存."
                            End If
                        End If
                        
                    End If
                    If strStock = "" Then
                        '更新库存显示
                        If Not (mbytInState = EM_ED_收费 And mstrInNO <> "") Then
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        End If
                        Call ShowStock(.执行部门ID, .Detail.名称, .Detail.库存)
                        Call ShowStatusCargoSpace(.收费细目ID, .执行部门ID)     '显示货位
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                    If Not (mbytInState = EM_ED_收费 And mstrInNO <> "") Then
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
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
            If zlStr.IsHavePrivs(mstrPrivs, "负数费用") = False Then
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
                    If zlStr.IsHavePrivs(mstrPrivs, "药品输入小数") = False Then
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
    If Not mbytInState = EM_ED_收费 Then Exit Sub
    mobjBill.性别 = zlStr.NeedName(cboSex.Text)

End Sub

Private Sub cbo费别_Click()
    If Not mbytInState = EM_ED_收费 Then Exit Sub
    If cbo费别.ListIndex = -1 Then
        mobjBill.费别 = "": Exit Sub
    End If
    If mbln不重算价格 Then Exit Sub
    If Not (mstrYBPati <> "" Or mobjBill.费别 <> zlStr.NeedName(cbo费别.Text)) Then Exit Sub
    '即使费用相同也要重算,因为医保验卡后必须重算,预结算才正确
    mobjBill.费别 = zlStr.NeedName(cbo费别.Text)
    If mbytInState <> EM_ED_收费 Then Exit Sub
    If CheckBillsEmpty Then Exit Sub
    
    
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
End Sub
Private Sub cbo开单科室_Click()
    Dim i As Long, lng开单部门ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then Exit Sub
        
    If cbo开单科室.ListIndex <> -1 Then lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If mobjBill.Pages(mintPage).开单部门ID = lng开单部门ID Then Exit Sub
    mobjBill.Pages(mintPage).开单部门ID = lng开单部门ID
        
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
        With mobjBill.Pages(mintPage)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.类别) = 0 And _
             (.Details(i).Detail.执行科室 = 6 And gbyt科室医生 <> 2 Or InStr(",1,2,", "," & .Details(i).Detail.执行科室 & ",") > 0 And gint病人来源 = 1) Then '6-开单人科室
                
                .Details(i).执行部门ID = lng开单部门ID
                
                If i <= Bill.Rows - 1 And .Details(i).执行部门ID <> 0 Then
                    If mbytInState = EM_ED_收费 Then
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
        Next
        End With
    End If
    
    '费别处理
    Call LoadAndSeek费别
    
End Sub

Private Sub LoadAndSeek费别(Optional blnNew As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载普通费别与动态费别,定位缺省费别或病人费别
    '入参:blnNew 是否新单据初始
    '编制:刘兴洪
    '日期:2014-06-05 16:30:25
    '说明:门诊记帐不使用动态费别
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long, blnDo As Boolean, strInfo As String
    
    If glngSys Like "8??" Then Exit Sub
    
    If cbo开单科室.ListIndex <> -1 Then lngDeptID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Call Load费别(cbo费别, lngDeptID, True, mrs费别)
                
    '显示可用动态费别：当前不是划价单时,窗体默认为可见
    If Bill.Active Or blnNew Then
        lbl动态费别.Caption = Load动态费别(lngDeptID)
        lbl动态费别.Tag = lbl动态费别.Caption
        lbl动态费别.Visible = lbl动态费别.Caption <> ""
        If lbl动态费别.Caption <> "" Then lbl动态费别.Caption = "(" & lbl动态费别.Caption & ")"
    End If
    
    
    cbo费别.Locked = (Not Bill.Active) _
            Or (mrsInfo.State = 1 And Not zlStr.IsHavePrivs(mstrPrivs, "调整病人费别"))
    
    cbo费别.TabStop = Not cbo费别.Locked And gbln费别
    
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
    '如果在cbo的keypress事件中用了弹出列表的API函数:
    '  sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
    '   cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
    If Not mblnCboClick Then cbo开单科室_Click
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then cbo开单科室.Text = ""
    mblnCboClick = False
End Sub

Private Function SetDefaultDept(lng开单人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的开单科室,但不触发Click事件
    '入参:lng开单人ID-开单人的ID
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 16:37:03
    '说明:缺省科室为"只服务于门诊,不具有医技性质"时，可以定位缺省
    '     或者开单人的所有科室都为同一优先排序级别时(如都是即服务于门诊或住院的)，可以定位缺省
    '     否则,不管人员的缺省科室，以GetDoctorDept中的医生顺序为准,第一个为缺省
    '     该顺序为: 1.只服务于门诊,不具有医技性质(检查,检验,手术,治疗,营养)
    '               2.只服务于门诊,具有医技性质(检查,检验,手术,治疗,营养)
    '               3.不只服务于门诊的
    '---------------------------------------------------------------------------------------------------------------------------------------------

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
    If Not (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then Exit Sub
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
                    If mbytInState = EM_ED_收费 Then
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
    If Not mbytInState = EM_ED_收费 Then Exit Sub
    mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
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
    
    mstrInNO = ""
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
        If fraBill.Visible Then cmdAddBill.Enabled = zlStr.IsHavePrivs(mstrPrivs, "普通病人多单据收费")
        
        cboNO.Text = ""

        Call SetDisible
        If Not zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then
            cbo开单人.Visible = False
            If gbyt科室医生 = 0 Then
                lbl科室.Visible = False
            Else
                lbl开单人.Visible = False
            End If
        End If
        
        fraAppend.Enabled = False
        cboNO.Locked = False
        cmd配方.Enabled = False
        cmdYB.Enabled = False
        
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
        cboNO.SetFocus
    Else
        
        If Not zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then
            cbo开单人.Visible = True
            If gbyt科室医生 = 0 Then
                lbl科室.Visible = True
            Else
                lbl开单人.Visible = True
            End If
        End If
        
        txtRePrint.Enabled = True
        txtIn.Text = ""
        txtIn.Enabled = True
        
        chkCancel.ForeColor = 0
        If fraBill.Visible Then cmdAddBill.Enabled = zlStr.IsHavePrivs(mstrPrivs, "普通病人多单据收费")
        txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnStrictCtrl
        Call SetDisible(True)
        
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
  
        cbo开单科室.Enabled = True
        cbo开单人.Enabled = True
        
        fraAppend.Enabled = True
        txtPatient.SetFocus
    End If
End Sub

Private Sub chk急诊_Click()

    If Not (chk急诊.Visible And Visible) Then Exit Sub
    '需要重新预结算
    If cmd预结算.Visible Then
        Call InitBalanceGrid
        cmd预结算.TabStop = True
        cmdOK.Enabled = False
    End If
End Sub

Private Sub chk急诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk加班_Click()
    Dim blnAdd As Boolean
    
    If Not mblnDo Then Exit Sub
    
    If mbytInState = EM_ED_浏览 Or chkCancel.Value = 1 Then Exit Sub
    If mbytInState = EM_ED_调整 Then Exit Sub
    If mbytInState = EM_ED_异常重收 Or mbytInState = EM_ED_异常作废 Then Exit Sub
    
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动将所有单据按收费类别进行单据分组
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 16:56:13
    '说明:暂不处理医保,收取工本费模式下,引起的工本费变化,暂不处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
End Sub

Private Function AddRowByOtherPageRow(tmpBillDetail As BillDetail, intPage As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将某单据行对象增加到指定的单据页中
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 16:56:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查非当前单据(并且不是划价单)中是否存在指定的收费类别或执行部门
    '入参:strKind-按收费类别区分时,为收费类别,按执行科室分单据时,执行部门ID
    '     bytWay-检查其它单据的方向,0-向后检查,1-向前检查
    '返回:如果不存在则返回0,存在则返回第一个存在的单据序号
    '编制:刘兴洪
    '日期:2014-06-05 16:57:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加一张单据
    '编制:刘兴洪
    '日期:2014-06-05 16:59:36
    '---------------------------------------------------------------------------------------------------------------------------------------------

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
    '多张单据时禁止导入,退费功能
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除指定的单据
    '入参:intPage-指定单据
    '编制:刘兴洪
    '日期:2014-06-05 17:00:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCurEmpty As Boolean, i As Integer
    
    blnCurEmpty = CheckBillsEmpty(intPage)
    
    '删除单据集合中的内容
    mobjBill.Pages.Remove intPage
    
    If intPage >= mcolBalance.Count Then mcolBalance.Remove intPage
    
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
    '打开退费功能
    If tbsBill.Tabs.Count = 1 Then
        chkCancel.Enabled = True
        cmdDelete.Enabled = True
    End If
    
    '激活Click,显示新定位单据的内容
    mintPage = 0 '强行激活
    Call tbsBill_Click
    
    '93450,多个病人科室的划价单，删除某一张后将病人科室设置为开单科室
    mobjBill.科室ID = 0
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

Private Function PriceBillShowing() As Boolean
    '当前界面中是否显示了划价单
    Dim i As Integer
    
    On Error GoTo errHandle
    If mobjBill Is Nothing Then Exit Function
    If mobjBill.Pages.Count = 0 Then Exit Function
    
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO <> "" Then '单据号不为空即为划价单
            PriceBillShowing = True
            Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadMultiBills(ByVal lng病人ID As Long, ByVal bln不允许多单据 As Boolean, _
    ByVal lng挂号科室 As Long, Optional blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一次性读取病人的多张划价单,该过程在病人读取成功之后调用
    '入参:bln不允许多单据，医保连续收费或不支持多单据收费时，不允许返回多张划价单收费
    '     lng挂号科室,当通过挂号单输入时,传入病人当前挂号单的挂号科室
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 17:00:54
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim objPage As New BillPage
    Dim arrBills As Variant, strBills As String
    Dim blnRead As Boolean, i As Long, k As Long
    
    If Not (gblnMulti And gblnSeekBill) Then Exit Function
    '108208,如果界面中显示了划价单则不再提取划价单
    If PriceBillShowing() = True Then Exit Function
    
    If lng病人ID = 0 Then Exit Function
    i = SeekPatiBill(lng病人ID)
    
    Call GetAsyncKeyState(VK_RETURN)
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
    mstrInNO = ""
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
    
            '多张单据时禁止退费功能
            chkCancel.Enabled = False
            cmdDelete.Enabled = False
                
            '激活Click,显示新增加单据的内容(空白)
            tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
        End If
                
        '读取划价单据内容(同cboNO_KeyPress)
        '----------------------------------------------------------------------
        blnRead = ReadBill(arrBills(i), 1)
        If blnRead Then k = k + 1: cboNO.Text = arrBills(i)
    Next
    Bill.Active = False
    chk加班.Enabled = False
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
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
        If txtPatient.Text = mstrPrePati And mlngPrePati <> 0 Then
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

Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新收费票据号
    '编制:刘兴洪
    '日期:2014-06-06 14:21:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mintInvoicePrint = 0 Then Exit Sub
    If gblnStrictCtrl Then
        'lblFact.tag主要是检查发票号是否手工输入的.手工输入的,发票号为空,否则是自动产生的发票号
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            If zlCheckInvoiceValied(mlng领用ID, 1, , mlngShareUseID, mstrUseType) = False Then
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

    If Not gblnMulti Then
        cmd配方.Enabled = Not cmd配方.Enabled
        cmdYB.Enabled = Not cmdYB.Enabled
    End If
    If frmClinicDelAndView.ShowMe(Me, EM_MULTI_退费, mstrPrivs, 0, False, mlng领用ID) Then
        Call RefreshFact
        If gbln累计 Then txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub cmdIDCard_Click()
    Dim strCommon As String, intAtom As Integer
    Dim strExpend As String, blnCreate As Boolean
    
    On Error GoTo errHandle
    '医疗卡发放管理
    If gobjSquare.objSquareCard Is Nothing Then
        Call CreateSquareCardObject(Me, mlngModul)
        blnCreate = True
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "医疗卡部件不存在,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If blnCreate Then
        If gobjSquare.objSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then Exit Sub
    End If
    Err.Clear: On Error GoTo 0
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    Call gobjSquare.objSquareCard.zlSendCard(Me, mlngModul, 0, 0)
    Call GlobalDeleteAtom(intAtom)
    If txtPatient.Enabled Then txtPatient.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Call FillDept
    Call FillDoctor
    Call ClearFullBill(False)    '主要是设置mobjBill.门诊标志
End Sub

Private Sub picAppend_Resize()
    Dim sngLeft As Single
    Err = 0: On Error Resume Next
    sngLeft = vsBalance.Left + vsBalance.Width + 100
    cmdOK.Left = sngLeft + (ScaleWidth - sngLeft - cmdOK.Width) \ 2 '  ScaleWidth - cmdOK.Width - 100
    cmdCancel.Left = cmdOK.Left
    cmdPrint.Left = cmdOK.Left
    cmd预结算.Left = cmdOK.Left
    If Not mbytInState = EM_ED_收费 Then Exit Sub
    vsBalance.Height = picAppend.ScaleHeight - vsBalance.Top - 20
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
            .mbytInFun = 0
            .mblnSetDrugStore = True
            .Show 1, Me
        End With
    ElseIf Panel.Key = "PatiSource" Then
        If gbln病人来源受权限控制 And zlStr.IsHavePrivs(mstrPrivs, "参数设置") = False Then
            '授权限控制,不能更改
            Exit Sub
        End If
        If Not CheckBillsEmpty Or txtPatient.Text <> "" Then
            If MsgBox("如果切换病人来源,将清空当前单据和病人信息" & vbCrLf & "你确定要继续吗?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If gint病人来源 = 1 Then    '门诊
            gint病人来源 = 2
        Else
            gint病人来源 = 1
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示预交及未结相关费用
    '入参:lngPatientID-病人ID
    '编制:刘兴洪
    '日期:2014-06-05 17:18:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    strSQL = "Select Nvl(Sum(金额), 0) 预交款总额, Nvl(Sum(冲预交), 0) 冲预交总额 From 病人预交记录 Where 病人id = [1] And 记录性质 In(1,11) and nvl(预交类别,2)=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID)
    
    If rsTmp.RecordCount > 0 Then
        MsgBox "预交款总额:" & Format(rsTmp!预交款总额, "0.00") & vbCrLf & "冲预交总额:" & Format((rsTmp!冲预交总额), "0.00") & vbCrLf & _
               "未 结 费用:" & Format(Val(cmdCancel.Tag), "0.00") & vbCrLf & _
               "可用预交款:" & Format((rsTmp!预交款总额 - (rsTmp!冲预交总额 + Val(cmdCancel.Tag))), "0.00") & vbCrLf & _
               "家 属 余额:" & Format(Val(cmd预结算.Tag), "0.00"), vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub sta_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Not Panel Is sta.Panels(Pan.C4预交信息) Then Exit Sub
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    '显示预交及未结详细信息
    Call ShowDeposit(mrsInfo!病人ID)
End Sub

Private Sub tbsBill_Click()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示选定页卡的页单据内容
    '编制:刘兴洪
    '日期:2014-06-05 17:21:19
    '说明:目前只有收费时才可能会进入
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
            cbo开单科室.Locked = False
            cbo开单人.Locked = False
            
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
            Call ReadBill(mobjBill.Pages(mintPage).NO, 1, , True)
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否多单据的内容都为空
    '入参:intPage=是否检查指定页,否则检查所有页
    '出参:
    '返回:为空返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 17:21:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    
    strYBPati = mstrYBPati: intInsure = mintInsure: strYBBill = mstrYBBill
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
        mstrInNO = ""
        mlngFirstID = 0: mstrFirstWin = ""
        
        If blnClearPatiInfor Then Call ClearPatientInfo(blnClearPatiInfor)
        Call ClearTotalInfo(True)
        
        Call InitCommVariable
        If gbln累计 Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
        End If
    End If
    
    Call ClearBillRows
    Call ClearMoney
    Call SetDisible(True)
    Call NewBill(IIf(mblnStartFactUseType, False, True), IIf(blnClearPatiInfor, True, False))
    If blnNotClearYb And intInsure <> 0 Then
        mintInsure = intInsure: mstrYBBill = strYBBill: mstrYBPati = strYBPati
        Call SetPatientEnableModi(False)
        txtPatient.ForeColor = vbRed
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), vbRed)
            End If
        End If
        cmdAddBill.Enabled = blnAdd
    End If
    sta.Panels(Pan.C2提示信息).Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    ClearFullBill = True
End Function
 
Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是手术输入情况(如果不存在主要手术,但存在附加手术,则禁止
    '出参:lngRow-返回附加手术的行
    '返回:存在主手术或没有输入附加手术,返回true,否则返回False
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
    Exit Function
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
    If txtPatient.Text = "" Then
        MsgBox "没有发现" & gstrCustomerAppellation & "信息,请输入" & gstrCustomerAppellation & "信息！", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Function
    ElseIf mobjBill.姓名 = "" Then
        mobjBill.姓名 = txtPatient.Text
    End If
    
    If CheckTextLength("姓名", txtPatient) = False Then Exit Function
    If CheckTextLength("年龄", txt年龄) = False Then Exit Function
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Function
    
    If mobjBill.费别 = "" Then
        MsgBox "请选择" & gstrCustomerAppellation & "费别！", vbInformation, gstrSysName
        If cbo费别.Visible And cbo费别.Enabled Then cbo费别.SetFocus
        Exit Function
    End If

    If CheckBillsEmpty Then
        MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    If mobjBill.Pages.Count > 1 Then
        For i = 1 To mobjBill.Pages.Count
            If CheckBillsEmpty(i) Then
                MsgBox "第 " & i & " 张单据没有输入任何内容！", vbInformation, gstrSysName
                tbsBill.Tabs(i).Selected = True
                Bill.SetFocus: Exit Function
            End If
        Next
    End If
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If CheckServeRange(0, .收费细目ID) = False Then Exit Function
            End With
        Next i
    Next p
    
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
    If CheckExecuteDeptCanDo() = False Then Exit Function
    
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
    If gbln必须输开单人 Then
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
    If mbytInState = EM_ED_收费 And (gbyt科室医生 = 0 Or gbyt科室医生 = 1) Then
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
    If gblnCheckRegeventDept And gint病人来源 = 1 _
        And (gTy_System_Para.Sy_Reg.bytNODaysGeneral > 0 Or gTy_System_Para.Sy_Reg.bytNoDayseMergency > 0) And mobjBill.病人ID > 0 Then
        Set rsTmp = GetDeptByRegevent(mobjBill.病人ID)
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
        Next
            
        '106490
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
    If gcurMax <> 0 Then
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
                                                If MsgBox("第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                    vbCrLf & "分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？", _
                                                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    blnMerge = True     '不应退出循环，因为还要检查是否有不同付数的中草药,如果有的话，不能自动合并
                                                Else
                                                    tbsBill.Tabs(k).Selected = True: Exit Function
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
                                        strInfo = "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？"
                                    Else
                                        strInfo = "第 " & j & " 行的分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？"
                                    End If
                                    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                        blnMerge = True     '可以退出循环
                                    Else
                                        Exit Function
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
    bln检查库存 = Not zlStr.IsHavePrivs(mstrPrivs, "不检查库存")      '是否有权限不检查库存(分批和时价必须检查)
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
            
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        
                        If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & _
                                dblToTal & """。", vbInformation, gstrSysName
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                            
                            If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行药品""" & .Detail.名称 & _
                                    """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & _
                                    dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                                tbsBill.Tabs(p).Selected = True
                                Bill.SetFocus: Exit Function
                            End If
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        
                        If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & dblToTal & """。", vbInformation, gstrSysName
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            
                            If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                    """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
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
        If HaveExecute(1, mstrInNO, 1) Then
            MsgBox "该单据包含完全执行或部分执行的项目,不允许修改。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '102660
    If mblnPeisPriceBill And mbytInState = EM_ED_收费 And gint病人来源 = 1 Then
        If CheckRegistedPeisBill() = False Then Exit Function
    End If
    
    '刘兴洪:检查是否只有附加手术,如果只有附加手术,直接退出:
    '22441
    If CheckMainOperation = False Then Exit Function
    
    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 1, _
        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0)) = False Then
        Exit Function
    End If
    
    isValiedCargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckRegistedPeisBill() As Boolean
    '体检病人挂号检查
    '102660，当前选中的费用中是否包含非体检费用，如果包含则需要检查是否挂号，如果只是体检费用，则不用检查是否挂号
    Dim blnExistCheckBill As Boolean, strNos As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO = "" Then '划价单
            blnExistCheckBill = True: Exit For
        Else
            strNos = strNos & "," & mobjBill.Pages(i).NO
        End If
    Next
    If blnExistCheckBill = False And strNos <> "" Then
        strSQL = "Select /*+cardinality(b, 10)*/ 1" & vbNewLine & _
                " From 门诊费用记录 A, Table(f_Str2list([1])) B" & vbNewLine & _
                " Where a.No = b.Column_Value And a.记录性质 = 1 And a.记录状态 = 0 And Nvl(a.门诊标志, 0) <> 4 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "是否体检单据检查", Mid(strNos, 2))
        blnExistCheckBill = Not rsTemp.EOF
    End If
    If blnExistCheckBill Then
        CheckRegistedPeisBill = CheckRegisted(mobjBill.病人ID, , True)
    Else
        CheckRegistedPeisBill = True
    End If
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
        If mblnSaveAsPrice Then CheckBillNOAndBookeFee = True: Exit Function
    End If
    
    mblnPrint = True
    '检查是否打印票据
    If mintInvoicePrint = 0 Then
        mblnPrint = False
    Else
        If (mintInvoicePrint = 2 And mbytInState <> EM_ED_异常重收) Or blnReCharge Then
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
            If zlCheckInvoiceValied(mlng领用ID, IIf(IsSplitPrintByNO, mobjBill.Pages.Count, 1), _
                                    txtInvoice.Text, mlngShareUseID, mstrUseType) = False Then
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
    dbl金额 = GetBillSum
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
        '-99-缴款;-98-找补;0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        Select Case Val(Nvl(grsTotal!性质))
        Case -98, -99, 1, 2
        Case Else
            '非医保的累计
            dbl本次应缴 = dbl本次应缴 + Val(Nvl(grsTotal!结算金额))
        End Select
        grsTotal.MoveNext
    Loop
    
    mobjBill.Pages(1).应缴金额 = RoundEx(dbl本次应缴, 6)
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
                        rsClass!金额 = RoundEx(Val(Nvl(rsClass!金额)) + dbl实收金额, 6)
                        rsClass.Update
                    End With
                Next
            End If
        Next
    End With
    If strNos = "" Then Exit Sub
    strNos = Mid(strNos, 2)
    strSQL = _
    "  Select  A.收费类别,  Sum(实收金额) As 实收金额 " & _
    "  From 门诊费用记录 A" & _
    "  Where A.NO in (Select Column_Value From  Table( f_Str2list([1])))  " & _
    "        And A.记录性质=1 And A.记录状态=0  " & _
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
    Dim frmBalance   As frmClinicChargeBalance
    Dim bytReturnMode As gExitMode, bln连续 As Boolean, dbl本次应缴 As Double
    Dim blnGetFact As Boolean, i As Integer, p As Integer
    Dim strReturn As String, lng结算序号 As Long, lng病人ID As Long
    Dim str划价Nos As String, rsItems As ADODB.Recordset
    
    If Not (mstrYBPati <> "" And MCPAR.门诊连续收费) And Not mblnSaveAsPrice Then
            Call AutoBultBookFee '收费时自动产生工本费项目
    End If
    
    If isValiedCargeFee = False Then Exit Function
    If zlGetSaveDataItems_Plugin(mobjBill, str划价Nos, rsItems) = False Then Exit Function
    If zlChargeSaveValied_Plugin(glngModul, 1, True, False, str划价Nos, rsItems) = False Then Exit Function
    
    '票据号及工本费及汇总金额相关检查
    If CheckBillNOAndBookeFee = False Then Exit Function
    If CheckInsure = False Then Exit Function
    
    '获取结算信息
    Set mobjChargeInfor = Nothing
    If GetChargeInfor(mobjChargeInfor) = False Then Exit Function
    mobjChargeInfor.应缴累计 = mcurBill应缴
    
    Set mFrmBalanceWin = New frmClinicChargeBalance
    If mFrmBalanceWin.zlChargeWin(Me, EM_FUN_收费, mlngModul, mstrPrivs, mobjChargeInfor, bytReturnMode, bln连续, mlngPreBrushCardID) = False Then
       If Not gfrmMain Is Nothing Then
             Call zlExeBalanceWinRefrshData(False, bytReturnMode, bln连续, mobjChargeInfor)
       End If
       Exit Function
    End If
    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, bytReturnMode, bln连续, mobjChargeInfor)
        mblnSaveData = True
        mintSucces = mintSucces + 1
        zlChargeFeeWin = True
    End If
End Function

Private Sub ShowLedWinAndSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示发药窗口及相关合计数据
    '编制:刘兴洪
    '日期:2012-02-06 14:31:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
    If mblnSaveAsPrice Then Exit Sub
    
    If Not (mstr西窗 <> "" Or mstr中窗 <> "" Or mstr成窗 <> "") _
        Or CCur(txt合计.Text) = 0 Then Exit Sub
    zl9LedVoice.DisplayBank "费用合计:" & txt合计.Text, _
        "取药窗口:" & IIf(mstr西窗 <> "", " " & mstr西窗, "") & _
        IIf(mstr成窗 <> "", " " & mstr成窗, "") & IIf(mstr中窗 <> "", " " & mstr中窗, "")
End Sub
 


Private Sub cmdOK_Click()
     mblnSaveData = False
    
    If mbytInState = EM_ED_收费 And chkCancel.Value = 0 Then
        '收费:包含异常单据的重新收费
        Call GetAsyncKeyState(VK_RETURN)
        If Not mblnSaveAsPrice Then
            If gfrmMain Is Nothing Then Me.Enabled = False
            If zlChargeFeeWin = False Then Exit Sub
        Else
            If SaveChargePriceBill = False Then Exit Sub
        End If
    ElseIf mbytInState = EM_ED_调整 Then '调整单据
        '========================================================================================================
        If Not SaveModi() Then Exit Sub
        mblnSaveData = True
        Unload Me
        
    ElseIf mbytInState = EM_ED_异常重收 Then
        cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保延时
        cmdCancel.Enabled = False: cmdAddBill.Enabled = False:: cmdDelBill.Enabled = False
        If ReChargeFee = False Then
            '61688
            cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
            cmdCancel.Enabled = True
            Exit Sub
        End If
    ElseIf mbytInState = EM_ED_异常作废 Then
        '作废异常单据
        cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保延时
        cmdCancel.Enabled = False: cmdAddBill.Enabled = False:: cmdDelBill.Enabled = False
        If DelErrBillFee = False Then
            cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
            cmdCancel.Enabled = True
            Exit Sub
        End If
    End If
    cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
    cmdCancel.Enabled = True
    Exit Sub
End Sub

Private Sub LoadFeeInfor(ByVal lngPatientID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取并显示病人预交,及费用余额信息
    '编制:刘兴洪
    '日期:2014-06-05 17:46:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim cur实收合计 As Currency
 
    Set rsTmp = GetMoneyInfo(lngPatientID, 0, False, 1, False, 0, True)
    Do While Not rsTmp.EOF
        If Nvl(rsTmp!家属, 0) = 0 Then
            cmdOK.Tag = Val(Nvl(rsTmp!预交余额))
            cmdCancel.Tag = Val(Nvl(rsTmp!费用余额))
            cmdPrint.Tag = Val(cmdOK.Tag) - Val(cmdCancel.Tag)
        Else
            cmd预结算.Tag = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
        End If
        rsTmp.MoveNext
    Loop
    sta.Panels(Pan.C4预交信息).Text = "预交:" & Format(Val(cmdPrint.Tag) + Val(Val(cmd预结算.Tag)), "0.00") & _
            IIf(Val(cmd预结算.Tag) > 0, "(含家属:" & Format(Val(cmd预结算.Tag), "0.00") & ")", "")
    Call ShowPrePayInfo(Val(cmdPrint.Tag) > 0 Or Val(cmd预结算.Tag) > 0)
End Sub

Private Sub cmdCancel_Click()
    mbln连续输入 = False
    If Not mbytInState = EM_ED_收费 Then Unload Me: Exit Sub
    If Not CheckBillsEmpty Or txtPatient.Text <> "" Then
        If ClearFullBill(True) = False Then Exit Sub
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前指定单据序号以前中最后一个含有药品的单据的第一行药品的部门ID
    '编制:刘兴洪
    '日期:2014-06-05 17:48:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    If chkCancel.Visible And chkCancel.Value = 1 Then
        Bill.Row = 1: Bill.Col = Bill.COLS - 1
    End If
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '保存为划价单
    If Button <> 2 Then Exit Sub
    If CheckSaveMultiPrice Then
        PopupMenu mnuFile, 2, cmdOK.Left + picAppend.Left - 800, cmdOK.Top + cmdOK.Height + picAppend.Top
    End If
End Sub
Private Sub cmdPrint_Click()
    Dim i As Integer, j As Integer
    Dim strPrintNO As String, strInfo As String
    Dim blnPrintList As Boolean, blnPrintExe As Boolean
    Dim int收费执行单 As Integer
    
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
    If zlStr.IsHavePrivs(mstrPrivs, "打印清单") Then
        If gint收费清单 = 1 Then
            blnPrintList = True
        ElseIf gint收费清单 = 2 Then
            If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
            End If
        End If
    End If
    
    '62982:李南春,2015/5/19,收费执行单
    int收费执行单 = Val(zlDatabase.GetPara("收费执行单打印方式", glngSys, mlngModul))
    If zlStr.IsHavePrivs(mstrPrivs, "收费执行单") Then
        If int收费执行单 = 1 Then
            blnPrintExe = True
        ElseIf int收费执行单 = 2 Then
            If MsgBox("要打印收费执行单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintExe = True
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
                    '77570,冉俊明,2014-9-5,医保支持连续收费，在病人完成收费后点击“完成收费”票据打印失败
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me, _
                        "发票号=NO", "NO='" & strPrintNO & "'", "价格等级=" & IIf(mstr普通价格等级 = "", "-", mstr普通价格等级), _
                        IIf(mintInvoiceFormat = 0, "", "ReportFormat=" & mintInvoiceFormat), 2)
                End If
            End If
            
            If blnPrintList Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & strPrintNO & "'", "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            End If
            
            If blnPrintExe Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me, "NO='" & strPrintNO & "'", 2)
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

Private Function InsurePreSwapAll(ByVal strDate As String, _
    ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊预结算(多单据一次结算)
    '编制:刘兴洪
    '日期:2011-08-15 17:30:29
    '说明:预结算信息保存在第一张单据中
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, strInvoice As String
    Dim str结算方式 As String, dbl结算金额 As Double
    Dim i As Long, p As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    
    Set rsTemp = MakeBillRecord(mobjBill, chk急诊.Value = 1, 0, strDate, cbo费别.Text, strInvoice)
    
    strBalance = "": strAdvance = ""
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
        If tbsBill.Tabs.Count > 1 Then
            sta.Panels(Pan.C2提示信息).Text = "单据预结算失败。"
        End If
        
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
        txtMCInvoice.Text = strAdvance
        txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
        txtMCInvoice.Visible = True
    End If
    
    MCPAR.医保不走票号 = False
    If InStr(1, strAdvance, ";") > 0 Then
          '38821:strAdvance:发票号;是否不走票据号
          MCPAR.医保不走票号 = Val(Split(strAdvance & ";", ";")(1)) = 1
    End If

    '根据预结算结果设置结算集
    p = 1: arrPage = Array()
    mcolBalance.Add Array()
    If strBalance <> "" Then
        '报销方式;金额;是否允许修改|....
        strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
        Call SetBalanceVal(mcolBalance, p, strBalance, strNone)
    End If
    InsurePreSwapAll = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetYBActualMoeny(ByVal str结算方式 As String, ByVal dbl结算金额 As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保结算的实际使用金额
    '入参:str结算方式-医保的结算方式
    '     dbl结算金额-医保的结算金额
    '返回:实际金额,否则返回False
    '编制:刘兴洪
    '日期:2014-06-06 16:12:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐 As Double, dbl个帐合计 As Double
    
    On Error GoTo errHandle
    
    If dbl结算金额 = 0 Then Exit Function
    If str结算方式 <> mstr个人帐户 Then GetYBActualMoeny = dbl结算金额: Exit Function
    '咸阳医保无法返回余额
     If (mdbl个帐余额 > -1 * mdbl个帐透支 Or mintInsure = 61) _
        And CCur(txt合计.Text) > 0 Then
        dbl个帐 = dbl结算金额
        If mintInsure <> 61 Then
            '计算个人帐户支付金额
            If RoundEx(mdbl个帐余额 - dbl个帐合计 - dbl个帐, 6) >= -1 * mdbl个帐透支 Then
                dbl个帐 = dbl个帐 '在允许透支范围内足够(允许透支0为特例)
            Else
                If mdbl个帐透支 = 0 And RoundEx(mdbl个帐余额 - dbl个帐合计, 6) > 0 Then
                    dbl个帐 = mdbl个帐余额 - dbl个帐合计 '不允许透支且有余额
                Else
                    '超过允许透支范围或不允许透支时无余额
                    If mdbl个帐透支 <> 0 Then
                        dbl个帐 = mdbl个帐余额 - dbl个帐合计 + mdbl个帐透支 '在允许透支范围内支付
                    Else
                        dbl个帐 = 0
                    End If
                End If
            End If
        End If
        dbl个帐合计 = dbl个帐合计 + dbl个帐
        dbl个帐 = Format(dbl个帐, "0.00")
        GetYBActualMoeny = dbl个帐
    Else
        GetYBActualMoeny = dbl结算金额
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    GetYBActualMoeny = dbl结算金额
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
    If MCPAR.多单据分单据结算 Then
        If InsurePreSwapNo(strDate, strNone) = False Then Exit Function
    ElseIf MCPAR.一次结算分单据退费 Then
        If InsurePreSwapDelNo(strDate, strNone) = False Then Exit Function
    Else
        If InsurePreSwapAll(strDate, strNone) = False Then Exit Function
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
        .TextMatrix(i, 0) = "自付合计": .TextMatrix(i, 1) = Format(mdbl应缴合计, "0.00")
        .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
        .Cell(flexcpFontBold, i, 0, i, .COLS - 1) = vbRed
        .RowPosition(i) = 0
    End With
    
    Call zl9InsureLedSpeak
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

Private Function InsurePreSwapNo(ByVal strDate As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊预结区分单据
    '返回:成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-08-15 18:20:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, strInvoice As String
    Dim p As Long, i As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    '对多张单据循环预结算
    MCPAR.医保不走票号 = False
    For p = 1 To tbsBill.Tabs.Count
        '直接输入的费用
        Set rsTemp = MakeBillRecord(mobjBill, chk急诊.Value = 1, p, strDate, cbo费别.Text, strInvoice)
        
        strBalance = "": strAdvance = ""
        If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
            If tbsBill.Tabs.Count > 1 Then
                sta.Panels(Pan.C2提示信息).Text = "第 " & p & " 张单据预结算失败。"
            End If
            
            Screen.MousePointer = 0
            Exit Function
        End If
        
        If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
             '38821:strAdvance:发票号;是否不走票据号
            txtMCInvoice.Text = Trim(Split(strAdvance & ";", ";")(0))
            txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
            txtMCInvoice.Visible = True
        End If
        
        '只要有一张单据要走票号，都要走票号
        If InStr(1, strAdvance, ";") > 0 Then
              '38821:strAdvance:发票号;是否不走票据号
              MCPAR.医保不走票号 = MCPAR.医保不走票号 Or Val(Split(strAdvance & ";", ";")(1)) = 1
        End If
        
        '根据预结算结果设置结算集
        arrPage = Array()
        '报销方式;金额;是否允许修改|....
        If strBalance <> "" Then
            strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
            Call SetBalanceVal(mcolBalance, p, strBalance, strNone)
        End If
    Next

    InsurePreSwapNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsurePreSwapDelNo(ByVal strDate As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊预结，一次结算分单据退费
    '返回:成功,返回true,否则返回false
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, strInvoice As String
    Dim varAdvance As Variant, varItem As Variant, str结算方式 As String
    Dim p As Long, i As Long, j As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    
    MCPAR.医保不走票号 = False
    
    Set rsTemp = MakeBillRecord(mobjBill, chk急诊.Value = 1, 0, strDate, cbo费别.Text, strInvoice)

    strBalance = "": strAdvance = ""
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
        If tbsBill.Tabs.Count > 1 Then
            sta.Panels(Pan.C2提示信息).Text = "单据预结算失败。"
        End If

        Screen.MousePointer = 0
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
        txtMCInvoice.Text = strAdvance
        txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
        txtMCInvoice.Visible = True
    End If
    
    MCPAR.医保不走票号 = False
    If InStr(1, strAdvance, ";") > 0 Then
        '38821:strAdvance:发票号;是否不走票据号
        MCPAR.医保不走票号 = Val(Split(strAdvance & ";", ";")(1)) = 1
    End If
    
    '单据序号:结算方式;金额;是否允许修改|...||单据序号:结算方式;金额;是否允许修改|...||...
    varAdvance = Split(strBalance, "||")
    For i = 0 To UBound(varAdvance)
        If InStr(varAdvance(i), ":") = 0 Then
            Screen.MousePointer = 0
            MsgBox "医保预结算返回结算结果格式不正确！", vbInformation, gstrSysName
            Exit Function
        End If
        varItem = Split(varAdvance(i), ":")
        p = Val(varItem(0)): str结算方式 = varItem(1)
        
        If p = 0 Then
            Screen.MousePointer = 0
            MsgBox "医保预结算返回结算结果格式不正确！", vbInformation, gstrSysName
            Exit Function
        End If
        
        str结算方式 = Replace(Replace(str结算方式, "|", "||"), ";", "|")
        '报销方式;金额;是否允许修改|....
        SetBalanceVal mcolBalance, p, str结算方式, strNone
    Next

    InsurePreSwapDelNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    dbl个帐合计 = GetMedicareSum(mcolBalance, mstr个人帐户)
    zl9LedVoice.DisplayBank "医保结算:", "帐户余额" & Format(mdbl个帐余额, "0.00"), _
        "帐户支付" & Format(dbl个帐合计, "0.00"), "统筹支付" & Format(GetMedicareSum(mcolBalance) - dbl个帐合计, "0.00")
    zl9LedVoice.Speak "#21 " & Format(mdbl应缴合计, "0.00")
End Sub

Private Sub cmd预结算_Click()
    Dim strNone As String
    Call AutoBultBookFee '收费时自动产生工本费
    
    If CheckBillsEmpty Then Exit Sub
    If gbytAutoSplitBill > 0 Then Call AutoSplitBill
                  
    If mintInsure <> 0 And MCPAR.实时监控 Then
        '本来对于划价单才传2进行明细和汇总的检查，但是，由于以下原因，数量和实收金额在输入检查通过后可能改变，所以须再次检查明细
        '1.导入单据，2.修改单据，3.输入中药配方，4.修改中药付数后，其它行的付数同时变化，5.输入主项，自动产生从项，以及从项汇总计算折扣
        '6.修改单价，7.调整执行科室，药品价格重算，8.调整费别，实收金额重算,9.先输费用再验证医保身份,其它等等
        If gclsInsure.CheckItem(mintInsure, 0, 2, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 1, 0)) = False Then Exit Sub
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

    If mblnFirst = False Then Exit Sub
    mblnFirst = False: mblnNotClearLedDisplay = False
    If LoadBill = False Then Unload Me: Exit Sub
    If mbytInState = EM_ED_异常作废 Then cmdOK_Click: Exit Sub
    
    On Error Resume Next
    If mbytInState = EM_ED_浏览 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = EM_ED_调整 Then
        txtDate.SetFocus
    ElseIf mbytInState = EM_ED_收费 And mstrInNO <> "" And Bill.Active Then
        Bill.SetFocus
    End If
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If mbytInState = EM_ED_收费 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    DoEvents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr("`｀", Chr(KeyAscii)) > 0 Then
        '报请出示就诊卡
        KeyAscii = 0
        If gblnLED Then zl9LedVoice.Speak "#30"
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
    mdbl应缴合计 = 0
End Sub

Private Sub ClearTotalInfo(Optional ByVal bln清除累计 As Boolean = False)
    '默认bln为false,不清除累计,(划价时累计txtbox作为应缴显示)
    txt合计.Text = gstrDec: txt应收.Text = gstrDec
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
    End If
    txt年龄.Text = "": txt门诊号.Text = ""
    Call zlControl.CboLocate(cbo年龄单位, "岁")
    Call txt年龄_Validate(False)
    lbl险类.Caption = ""
    cmdOK.Tag = "": cmdCancel.Tag = "": cmdPrint.Tag = "": cmd预结算.Tag = ""
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
    
    '是否存在误差费的处理
    If IsCheck误差费 = False Then Exit Function
    
    '结算方式检查
    Set mrs结算方式 = Get结算方式("收费")
    Set mrs缺省结算方式 = Get结算方式("收费", "", True)
    If mrs结算方式.RecordCount = 0 Then
        MsgBox "收费场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    If mstr个人帐户 = "" Then
        mrs结算方式.Filter = "性质=3"
        If Not mrs结算方式.EOF Then mstr个人帐户 = mrs结算方式!名称
    End If
    If mstr应付款结算方式 = "" Then
        mrs结算方式.Filter = "应付款=1"
        If Not mrs结算方式.EOF Then mstr应付款结算方式 = Nvl(mrs结算方式!名称)
    End If
    mrs结算方式.Filter = 0
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function LoadBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载单据数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 16:41:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Select Case mbytInState
    Case EM_ED_收费 'b.新增,修改
        If mbytInState = EM_ED_收费 And gbln累计 Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
            txt累计.ToolTipText = "当前操作员今日收费累计额"
        End If
        '1.新增单据
        If Not NewBill(Not mblnStartFactUseType, False) Then Exit Function           '参数false表示不用再读取可用费别,因为前面InitData已做此操作
        LoadBill = True: Exit Function
    Case EM_ED_异常重收, EM_ED_异常作废 '异常单据的处理
        If mlng结帐ID = 0 And mlng结算序号 = 0 Then Exit Function
        If mlng结帐ID = 0 Then mlng结帐ID = Abs(mlng结算序号)
        If LoadErrBillCharge(mlng结帐ID) = False Then Exit Function
        LoadBill = True: Exit Function
    Case EM_ED_调整, EM_ED_浏览   'a.显示、调整单据
        If Not ReadBill(mstrInNO, 0) Then Exit Function
        If Not zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then
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
    Bill.Height = Me.ScaleHeight - Bill.Top - sta.Height - picAppend.Height - IIf(fraSubBill.Visible, fraSubBill.Height + 30, 0) _
        - IIf(fra退费摘要.Visible, fra退费摘要.Height + 30, 0)
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
    If TypeName(cbo开单人.Container) = TypeName(fraAppend) Then
       ' lbl开单人.Left = fraAppend.Left + cboBaby.Left + cboBaby.Width + 1000
        cbo开单人.Left = lbl开单人.Left + lbl开单人.Width + 20
    Else
        cbo开单科室.Left = lbl开单人.Left + lbl开单人.Width + 20
    End If
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If mbytInState = EM_ED_收费 And mstrYBPati <> "" And mstrInNO = "" Then
        If MsgBox("当前正在对医保病人收费，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        '取消医保病人身份验证,返回假时不退出
            Cancel = 1: Exit Sub
        End If
    End If
    
    SaveWinState Me, App.ProductName, mstrTittle & "_" & mbytInState
    If mbytInState = EM_ED_收费 Then
        Call SaveRegisterItem(g私有模块, Me.Name, "idkind", IDKind.IDKind)
    End If
    
    zlCommFun.OpenIme False
    mbytInState = EM_ED_收费
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mstrCardNO = ""
    mblnNOMoved = False   '查看时,可能传入true,
    mblnYB结算作废 = False
    
    mintBillNO = 0: mintMoneyRow = 0
    mlngFirstID = 0: mstrFirstWin = ""
    mlng领用ID = 0
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    
    '清空数据对象
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrs费别 = Nothing
    Set mrs费用类型 = Nothing
    Set mrs发药窗口 = Nothing
    
    'LED初始化
    If mbytInState = EM_ED_收费 And gblnLED Then
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
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
    mblnHaveExcuteData = False
    
    Set mrs结算方式 = Nothing
    Set mrs缺省结算方式 = Nothing
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

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Not (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then Cancel = True: Exit Sub
  With vsBalance
    '不允许修改的医保项目
    If Val(.RowData(Row)) = 0 Or Col <> 1 Then Cancel = True: Exit Sub
    If MCPAR.多单据分单据结算 Then Cancel = True: Exit Sub
  End With
End Sub
 
Private Sub vsBalance_DblClick()
  If Not (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then Exit Sub
  With vsBalance
    '不允许修改的医保项目
    If Val(.RowData(.Row)) = 0 Or .Col <> 1 Then Exit Sub
    If MCPAR.多单据分单据结算 Then Exit Sub
    .EditCell
    .EditSelStart = 0
    .EditSelLength = zlCommFun.ActualLen(.EditText)
  End With
End Sub

Private Sub vsBalance_EnterCell()
    With vsBalance
        If .Col < 0 Then Exit Sub
        If .Col = 0 Then .Col = 1
    End With
    If Not (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then Exit Sub
    
    With vsBalance
        If .Row < 0 Then Exit Sub
        If .RowData(.Row) = 0 Then
             .FocusRect = flexFocusLight
        Else
             .FocusRect = flexFocusHeavy
        End If
    End With
End Sub

Private Sub vsBalance_GotFocus()
    vsBalance_EnterCell
End Sub

Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If Not (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then Exit Sub
    If vsBalance.Col <> 1 Then Exit Sub
    Call VsFlxGridCheckKeyPress(vsBalance, vsBalance.Row, vsBalance.Col, KeyAscii, m金额式)
End Sub
Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col <> 1 Then Exit Sub
    Call VsFlxGridCheckKeyPress(vsBalance, Row, Col, KeyAscii, m金额式)
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费时导入单据
    '编制:刘兴洪
    '日期:2014-06-06 15:38:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngPre As Long, strPre As String, strNo As String, strNos As String
    Dim intInsure As Integer, i As Long, j As Long
    Dim lng病人ID As Long, lng结帐ID As Long, bln急诊 As Boolean
    Dim strTmp As String, blnNOMoved As Boolean
    Dim objBill As ExpenseBill
    Dim varNos As Variant
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii): Exit Sub
    End If

    KeyAscii = 0
    '导入单据
    txtIn.Text = GetFullNO(txtIn.Text, 13)
    Call zlControl.TxtSelAll(txtIn)
    strNo = txtIn.Text
           
    'a.单张单据模式,清除当前单据对象及病人信息
    If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then
        Call ClearFullBill(False)
        
        Set mobjBill = ImportBill(strNo, False, 0, , False, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
        If mobjBill.NO = "" Then
            '78502,冉俊明,2014-10-10
            MsgBox "单据 " & strNo & " 读取失败。", vbInformation, gstrSysName
            txtIn.SetFocus: Exit Sub
        End If
        
        If Not zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then mobjBill.Pages(mintPage).开单人 = ""
        '清除病人信息
        Call ClearmobjBill
    Else
    'b.多张单据模块,新增单据,保留当前单据内容及病人相关信息,
    '不提供从后备表中导入的功能
        blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "1,11")
        strNos = zlGetBalanceNos(0, strNo, blnNOMoved)
        '77841,冉俊明,2014-9-15,门诊收费多张单据模式时不能导入划价单
        If strNos = "" Then strNos = strNo
        varNos = Split(strNos, ",")
        For i = 0 To UBound(varNos)
            strNo = Replace(varNos(i), "'", "")
            Set objBill = ImportBill(strNo, False, 0, , False, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
            
            If objBill.NO = "" Then
                '78502,冉俊明,2014-10-10
                MsgBox "单据 " & strNo & " 读取失败。", vbInformation, gstrSysName
                '使其触发tbsBill_Click事件
                mintPage = tbsBill.Tabs.Count + 1
                tbsBill.Tabs(mintPage - 1).Selected = True
                txtIn.SetFocus: Exit Sub
            End If
            
            '78566,冉俊明,2014-10-13,最后一张单据为划价单时也要新增单据
            If i > 0 Or mobjBill.Pages(mintPage).Details.Count > 0 Or mobjBill.Pages(mintPage).NO <> "" Then
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
                If zlStr.IsHavePrivs(mstrPrivs, "显示开单人") Then .开单人 = objBill.Pages(1).开单人
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
    If mobjBill.Pages(mintPage).Details.Count = 0 Then
        Bill.Rows = 2
    Else
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
    End If
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
    If mbytInState = EM_ED_收费 And mstrInNO <> "" Then mstrInNO = ""
    
    '要放在mstrInNO之后,因为以此来判断是否修改单据,以加回原库存
    Call CalcDrugStock
                
    Bill.Active = True
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub CalcDrugStock(Optional intPage As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算每行药品库存
    '入参:intPage-指定页面(0时为当前页面)
    '编制:刘兴洪
    '日期:2014-06-06 15:39:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim str药房IDs As String

    If intPage = 0 Then intPage = mintPage
    
    For i = 1 To mobjBill.Pages(intPage).Details.Count
        With mobjBill.Pages(intPage).Details(i)
            Bill.RowData(i) = Asc(.收费类别) '特殊处理
            
            If InStr(",5,6,7,", .收费类别) > 0 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + .原始数量
                
                Call SetItemRowColor(1, i)  '储备限额提示
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + .原始数量
                
                Call SetItemRowColor(1, i) '储备限额提示
            End If
        End With
    Next
End Sub

Private Sub txtInvoice_Change()
    lblFact.Tag = ""
End Sub

Private Sub txtInvoice_LostFocus()
    If Not (mbytInState = EM_ED_收费) Then Exit Sub
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
    
    If mbytInState = EM_ED_收费 Then mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
End Sub
 
 
Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim curOrig As Currency
    Dim curTotal As Currency, arrValue As Variant
    Dim i As Integer, p As Integer, str结算方式 As String
    
    With vsBalance
        If Row < 0 Then Exit Sub
        If Col <> 1 Or Col < 0 Then Exit Sub
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If Not IsNumeric(strKey) Then
            MsgBox "输入了非法的""" & strKey & """结算金额！", vbInformation, gstrSysName
            .EditCell
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True
            Exit Sub
        End If
        
        str结算方式 = Trim(.TextMatrix(.Row, 0))
        If str结算方式 = "" Then Exit Sub        '结算金额不允许超过返回的原始金额(个人帐户允许透支时再判断)
        curOrig = GetMedicareSum(mcolBalance, .TextMatrix(.Row, 0), , True) '该结算方式所有原始返回金额和
        If (.TextMatrix(Row, 0) <> mstr个人帐户 Or mdbl个帐透支 = 0) _
            And Val(strKey) > curOrig And Val(strKey) <> 0 And curOrig <> 0 Then
            MsgBox "输入的""" & .TextMatrix(Row, 0) & """结算金额不能超过 " & Format(curOrig, "0.00") & " ！", vbInformation, gstrSysName
            .EditCell
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True
            Exit Sub
        End If
        '个人帐户检查
        If .TextMatrix(Row, 0) = mstr个人帐户 Then
            '不允许超过允许透支金额
            If mdbl个帐余额 - Val(strKey) < -1 * mdbl个帐透支 Then
                MsgBox "帐户余额:" & Format(mdbl个帐余额, "0.00") & _
                    IIf(mdbl个帐透支 = 0, "", "(" & "允许透支:" & Format(mdbl个帐透支, "0.00") & ")") & _
                    "不足要结算的金额。", vbInformation, gstrSysName
                .EditCell
                .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
        End If
        
        '不允许超出单据剩余可结算金额
        curTotal = GetBillSum
        For p = 1 To mcolBalance.Count
            For i = 0 To UBound(mcolBalance(p))
                '结算方式;原始(最大)金额;可否修改;改后金额
                arrValue = Split(mcolBalance(p)(i), ";")
                If arrValue(0) <> .TextMatrix(.Row, 0) Then
                    curTotal = curTotal - CCur(arrValue(3))
                End If
            Next
        Next

        If Val(strKey) > curTotal And RoundEx(Val(strKey), 6) <> 0 Then
            MsgBox "结算金额过大，超过单据允许结算金额:" & Format(curTotal, "0.00") & "。", vbInformation, gstrSysName
            .EditCell
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True
            Exit Sub
        End If
                
        
        If zlDblIsValid(strKey, 5, False, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        strKey = Format(Val(strKey), "0.00")
        .EditText = strKey
        .TextMatrix(Row, Col) = strKey
        
        Call SetBalanceVal(mcolBalance, 1, str结算方式 & "|" & CCur(Val(strKey)))
        '重新计算应缴，误差(分币)等:费用明细未变,全部不用重新计算
        Call ShowMoney(-1, Not (cmd预结算.Visible And cmdOK.Enabled))
        vsBalance.TextMatrix(0, 1) = Format(mdbl应缴合计, "0.00") '更新自付金额显示
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

    If (mbytInState = EM_ED_收费 And mobjBill.Pages(mintPage).Details.Count = 0) _
        Or chkCancel.Value = 1 Then
        cboNO.Locked = False '收费时，空单据可以提划价单，也可重复提取
    Else
        cboNO.Locked = True
    End If
    '收费时如果已验证医保病人身份,则禁止再读取划价单
    If mbytInState = EM_ED_收费 And mstrYBPati <> "" Then cboNO.Locked = True
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
        '划价收费
        cboNO.Text = GetFullNO(cboNO.Text, 13)
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
        
        Call ClearPayInfo
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        '不是修改时,mstrInNO为提取的退费单,审核单,但不含划价单
        If Not (chkCancel.Value = 0) Then mstrInNO = UCase(cboNO.Text)
        
        blnRead = ReadBill(cboNO.Text, 1, blnNull)
        
        If blnRead Then
           Bill.Active = False
            chk加班.Enabled = False
            
            '如果没有权限，提取划价单后,只能输入医保病人
            If gint病人来源 = 1 And zlStr.IsHavePrivs(mstrPrivs, "允许非医保病人") = False Then
                 ClearPatientInfo (True)
            End If
            
            '如果是挂号产生临时病人姓名模式,则读取病人身份信息,以便修改
            If txtPatient.Text = "新病人" Then
                Call GetPatient("-" & mobjBill.病人ID)
            End If
            
            '显示摘要
            Call Bill_EnterCell(1, BillCol.项目)
            
            If txtPatient.Text <> "新病人" Then
                If Not CheckRegisted(mobjBill.病人ID, mblnPeisPriceBill) Then
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
            If txtPatient.Text = "" Or blnNull Then
                txtPatient.SetFocus
            Else
                If cmd预结算.Enabled And cmd预结算.Visible Then
                    cmd预结算.SetFocus
                ElseIf cmdOK.Enabled And cmdOK.Visible Then
                    cmdOK.SetFocus
                End If
            End If
        Else
            If Not (chkCancel.Value = 0) Then mstrInNO = ""
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
    zlControl.TxtSelAll txtPatient
    zlCommFun.OpenIme True
    
    'LED语音报价
    If mbytInState = EM_ED_收费 And gblnLED And Trim(txtPatient.Text) = "" Then
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
                If mbytInState = EM_ED_收费 And Not CheckBillsEmpty Then
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
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo开单人.ListIndex >= 0 Then lng医生ID = cbo开单人.ItemData(cbo开单人.ListIndex)
    If mrs开单科室 Is Nothing Then FillDept (lng医生ID)
    If zlSelectDept(Me, mlngModul, cbo开单科室, mrs开单科室, cbo开单科室.Text) = False Then KeyAscii = 0: Exit Sub
End Sub

Private Function isCheck开单人Exists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '返回:存在返回true,否则返回False
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
    Dim rsTemp As ADODB.Recordset, strAdded As String
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用中药配方输入功能
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-06 16:43:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDetails As BillDetails
    Dim str动态费别 As String, lng病人科室ID As Long
    Dim int序号 As Integer, i As Long
    
    If Not (Bill.Active And mbytInState = EM_ED_收费) Then Exit Sub
    
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
    If glngSys Like "8??" Then
        str动态费别 = zlStr.NeedName(cbo费别.Text)
    Else
        str动态费别 = zlStr.TrimEx(zlStr.NeedName(cbo费别.Text) & "," & lbl动态费别.Tag, ",")
    End If
    
    '调用窗口
    Set objDetails = frmCHRecipe.ShowMe(Me, mstrPrivs, mlngModul, 0, 0, Original.实收合计, mobjBill.病人ID, lng病人科室ID, Get开单科室ID, _
        IIf(mlng中药房 = 0, glng中药房, mlng中药房), mobjBill.Pages(mintPage).Details, zlStr.NeedName(cbo费别.Text), str动态费别, _
         IIf(mstrYBPati <> "", mintInsure, 0), chk加班.Value = 1, mobjBill.Pages(mintPage).煎法, Nothing, mcolStock1, zl获取中药形态(mintPage, Bill.Row, True))
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
        ElseIf cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus
        End If
    Else
        Bill.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '参数：Shift=-1：表示是程序强行在调用
    Select Case KeyCode
        Case vbKeyF1  '帮助
            ShowHelp App.ProductName, Me.hWnd, Me.Name & "2"
        Case vbKeyF2
            If Shift = vbCtrlMask Then
                If mbytInState = EM_ED_收费 And mstrInNO = "" And gbytAutoSplitBill > 0 Then
                    Call AutoSplitBill
                End If
            Else
                mblnF2Save = True
                    If ActiveControl Is txtPatient Then
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
                If mbytInState = EM_ED_收费 And (gint病人来源 = 1 Or gint病人来源 = 2) Then
                    If chkCancel.Value = 0 And zlStr.IsHavePrivs(mstrPrivs, "保险收费") Then
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
            If cmdDelete.Visible And cmdDelete.Enabled Then
                cmdDelete.SetFocus: Call cmdDelete_Click
            End If
        Case vbKeyF9 '定位到单据号输入框
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Case vbKeyF10 '就诊卡发放
            If cmdIDCard.Visible And cmdIDCard.Enabled Then cmdIDCard.SetFocus: cmdIDCard_Click
        Case vbKeyF11
            If cmd配方.Enabled And cmd配方.Visible Then cmd配方.SetFocus: Call cmd配方_Click
        Case vbKeyF12
            If Shift = vbAltMask Then
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
            If Shift = vbCtrlMask Then
                If CheckSaveMultiPrice Then
                    Call mnuFileSavePrice_Click
                Else
                    MsgBox "仅在收费时允许保存为划价单." & vbCrLf & "如果是多张单据收费,要求不含导入的单据", vbInformation, gstrSysName
                End If
            End If
        Case vbKeyD
            If Shift = vbCtrlMask Then
                If sta.Panels(Pan.C4预交信息).Visible And mrsInfo.State = 1 Then
                    Call ShowDeposit(mrsInfo!病人ID)
                End If
            End If
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前收入项目行数调整各列宽
    '编制:刘兴洪
    '日期:2014-06-06 16:47:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngW As Long
    lngW = mshMoney.Width - 75
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    
    mshMoney.ColWidth(0) = 600
    
    lngW = lngW - mshMoney.ColWidth(0)
    mshMoney.ColWidth(1) = lngW * 0.45
    mshMoney.ColWidth(2) = lngW * 0.55
    mshMoney.ColWidth(3) = 0
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '返回:数据初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-06 16:48:21
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim dtCurdate As Date     '服务器当前时间
    
    On Error GoTo errH
        
    '初始化病人信息对象
    Set mrsInfo = New ADODB.Recordset
    '查看时,不支持身份证识别,修改时要支持,因为修改后可能继续新单收费
    If mbytInState = EM_ED_收费 Then
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.hWnd)
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    
    '刘兴洪:结算卡的一些处理
    Call initCardSquareData
    
    If mbytInState = EM_ED_收费 Then
        Set mrsOneCard = GetOneCard
        mblnOneCard = mrsOneCard.RecordCount > 0
    End If
        
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
    
 
    '费别,默认显示适用于所有科室的
    Call Load费别(cbo费别, 0, False, mrs费别)
    mrs费别.Filter = ""
    If mrs费别.RecordCount = 0 Then
        MsgBox "没有有效费别设置，请先到费别管理中进行设置！", vbInformation, gstrSysName
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
    
    
    '可用收费类别:按序号排序
    If gstr收费类别 = "" Then
        strSQL = "Select 编码,名称 as 类别 from 收费项目类别 Where 编码<>'1' Order by 序号"
    Else
        strSQL = "" & _
        "   Select A.编码,A.名称 as 类别 " & _
        "   From 收费项目类别 A  " & _
        "   Where A.编码 in (select Column_Value From Table( f_Str2list([1]))) " & _
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
        mlng药品类别ID = ExistIOClass(8)
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定处方单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr收费类别, "'4'") > 0 Or gstr收费类别 = "" Then
        mlng卫材类别ID = ExistIOClass(40)
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
    dtCurdate = zlDatabase.Currentdate
    txtDate.Text = Format(dtCurdate, "yyyy-MM-dd HH:mm:ss")
    '自动识别加班
    If mbytInState <> EM_ED_调整 And mstrInNO = "" Then
        If OverTime(dtCurdate) Then chk加班.Value = 1
    End If
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str类别 As String, _
    ByVal intPage As Integer, ByVal lngRow As Long, _
    ByVal strDeptIDs As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最近输入的相同类别项目的执行科室ID
    '入参:str类别-收费类别
    '     intPage-指定页面
    '返回:成功返回执行部门ID ,否则返回0
    '编制:刘兴洪
    '日期:2014-06-06 16:54:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional blnEnter As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据列设置下拉列表框内容
    '入参:blnEnter=是否按光标进入该列处理,这时显示的内容保持不变
    '编制:刘兴洪
    '日期:2014-06-06 16:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数
    '编制:刘兴洪
    '日期:2010-01-27 10:17:11
    '问题:27663
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mTy_Para
        .bln住院病人门诊收费 = IIf(Val(zlDatabase.GetPara("住院病人按门诊收费", glngSys, mlngModul, "0")) = 1, True, False)
    End With
End Sub


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据表单要完成的功能设置界面布局
    '编制:刘兴洪
    '日期:2014-06-06 16:56:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead() As String, i As Integer, arrBaby As Variant, strTmp As String
    
    '刘兴洪 问题:27331 日期:2010-01-12 09:48:43
    If mbytInState = EM_ED_收费 Then
        '只有划价才会有此判断
        MCPAR.blnOnlyBjYb = zlIsOnly北京医保
    Else
        MCPAR.blnOnlyBjYb = False
    End If
    Call InitModulePara
    
    
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
         If mbytInState = EM_ED_收费 Then
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
        If mbytInState = EM_ED_收费 Or mbytInState = EM_ED_调整 Then '编辑界面
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
    End With
    
    '恢复注册表保存宽度
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & 0 & mbytInState)
    If gTy_System_Para.byt药品名称显示 <> 2 Then
        '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        Bill.ColWidth(BillCol.商品名) = 0
    Else
        If Bill.ColWidth(BillCol.商品名) = 0 Then
             Bill.ColWidth(BillCol.商品名) = GetOrigColWidth(BillCol.商品名)
        End If
    End If
        
    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = mbytInState = EM_ED_收费
    sta.Panels("PY").Visible = mbytInState = EM_ED_收费 And gbln简码切换 '35242
    sta.Panels("WB").Visible = mbytInState = EM_ED_收费 And gbln简码切换
    If mbytInState = EM_ED_收费 Then
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
    
    IDKind.Enabled = mbytInState = EM_ED_收费
    If mbytInState = EM_ED_收费 Then
        Call GetRegisterItem(g私有模块, Me.Name, "idkind", strTmp)
        IDKind.IDKind = Val(strTmp)
    End If
    
    '多单据收费:目录仅支持收费界面
    fraBill.Visible = mbytInState = EM_ED_收费 And mstrInNO = "" And gblnMulti
    lblDuty.Caption = ""
    fraSubBill.Visible = mbytInState = EM_ED_收费      '该栏上还要显示开单人的专业技术职务
    
    '刘兴洪 问题:26949 日期:2009-12-28 13:52:50
    fra退费摘要.Visible = mblnDelete
    If Not (mbytInState = EM_ED_收费 And mstrInNO = "" _
        And zlStr.IsHavePrivs(mstrPrivs, "保险收费") _
        And gint病人来源 = 1) Then
        
        cmdYB.Visible = False
        lblRePrint.Left = lblRePrint.Left - cmdYB.Width
        txtRePrint.Left = txtRePrint.Left - cmdYB.Width
        lblIn.Left = lblIn.Left - cmdYB.Width
        txtIn.Left = txtIn.Left - cmdYB.Width
    End If
    cmdSelWholeSet.Visible = mbytInState = EM_ED_收费
    cmdSaveWholeSet.Visible = zlStr.IsHavePrivs(mstrPrivs, "增加成套项目")
    
    '中药配方:新单时有效
    If Not (mbytInState = EM_ED_收费) Then
        cmd配方.Visible = False
        lblRePrint.Left = lblRePrint.Left - cmd配方.Width
        txtRePrint.Left = txtRePrint.Left - cmd配方.Width
        lblIn.Left = lblIn.Left - cmd配方.Width
        txtIn.Left = txtIn.Left - cmd配方.Width
    End If
                    
    '重打(仅收费有效)
    If Not (mbytInState = EM_ED_收费 And mstrInNO = "" _
            And zlStr.IsHavePrivs(mstrPrivs, "重打票据") And zlStr.IsHavePrivs(mstrPrivs, "收据打印")) Then
        lblRePrint.Visible = False
        txtRePrint.Visible = False
        
        lblIn.Left = lblIn.Left - lblRePrint.Width - txtRePrint.Width
        txtIn.Left = txtIn.Left - lblRePrint.Width - txtRePrint.Width
    End If

    '导入(仅新增时有效)
    If Not (mbytInState = EM_ED_收费 And mstrInNO = "") Then
        lblIn.Visible = False
        txtIn.Visible = False
    End If
   
    If mbytInState = EM_ED_浏览 Then
         vsBalance.Width = vsBalance.Width + 100
    End If
    
    '票据号
    lblFact.Visible = True
    txtInvoice.Visible = True
    txtMCInvoice.Top = txtInvoice.Top   '在预结算后才会显示
    txtMCInvoice.Left = txtInvoice.Left
    
    '动态费别
    If glngSys Like "8??" Then
        lbl动态费别.Visible = False
    Else
        If mbytInState = EM_ED_浏览 Or mbytInState = EM_ED_调整 Then
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
    
    '收费票据打印格式:收费,退费时显示
    If mbytInState = EM_ED_收费 Then
        Call ZlShowBillFormat(mlngModul, lblFormat, mintInvoiceFormat)
    End If
    
    '退费销帐按钮
    If mstrInNO = "" Then
        cmdDelete.Visible = True '收费支持多单据时使用多单据退费
        chkCancel.Visible = False
    End If
    
    If Not (mbytInState = EM_ED_收费 And mstrInNO = "") Then
        chkCancel.Visible = False
    End If

    If glngSys Like "8??" Then
        Caption = "药店收费处理"
        lblTitle.Caption = gstrUnitName & "药店收费单"
    Else
        Caption = "病人收费处理"
        lblTitle.Caption = gstrUnitName & "病人收费单"
    End If
        
    Call SetMoneyList
    
    Call InitBalanceGrid
    
    If mbytInState <> EM_ED_收费 Then
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
    If Not zlStr.IsHavePrivs(mstrPrivs, "门诊退费") Then
        chkCancel.Visible = False
        cmdDelete.Visible = False
    End If
    txtInvoice.Locked = Not (zlStr.IsHavePrivs(mstrPrivs, "修改票据号")) And gblnStrictCtrl
     
        
    If mbytInState = EM_ED_收费 Or mbytInState = EM_ED_调整 _
        Or mbytInState = EM_ED_异常重收 Or mbytInState = EM_ED_异常作废 Then
        '执行或调整状态
        If mbytInState = EM_ED_收费 Then
            If mstrInNO <> "" Then txtPatient.BackColor = &HE0E0E0           '修改
        ElseIf mbytInState = EM_ED_调整 Then '调整开单人和时间
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
        Call SetButton(3) '取消
        fra退费摘要.Enabled = False
        If mblnDelete Then lblFlag.Visible = True
    End If
    
    If gbyt科室医生 = 0 Then
        Call ExChangeLocate(cbo开单科室, cbo开单人)
        lbl科室.Caption = "开单人(&W)"
        lbl科室.Left = lblPatient.Left
        lbl开单人.Caption = "开单科室"
        cbo开单科室.TabStop = False
    End If
    
    If Not mbytInState = EM_ED_收费 Then
        sta.Panels("Drugstore").Visible = False
    End If
    
    If mbytInState = EM_ED_收费 And mstrInNO = "" Then
        sta.Panels("PatiSource").Visible = True
        Set sta.Panels("PatiSource").Picture = imgPati.ListImages(IIf(gint病人来源 = 1, "OutPati", "InPati")).Picture
    Else
        sta.Panels("PatiSource").Visible = False
    End If
    Bill.ColWidth(BillCol.从属父号) = 0
    Bill.ColWidth(BillCol.医嘱序号) = 0
    
    '82801,冉俊明,2015-2-26
    txt年龄.MaxLength = zlGetPatiInforMaxLen.intPatiAge
End Sub

Private Sub SetButton(bytType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置功能按钮状态和位置
    '入参:：bytType=1:预结算,确定,取消
    '              2:确定,取消
    '              3:取消
    '              4:预结算,确定,完成收费,取消
    '编制:刘兴洪
    '日期:2014-06-06 17:36:02
    '说明：该函数为初始时调用,不可重复调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Const H_间隔 = 45
    
    LockWindowUpdate picAppend
    
    '恢复缺省状态，且不可见
    cmd预结算.Visible = False
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdPrint.Visible = False
    
    cmd预结算.Top = lblSeek.Top
    cmdOK.Top = cmd预结算.Top + cmd预结算.Height + H_间隔
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

Private Sub SetDisible(Optional blnEditSta As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:界面设置为不可修改状态
    '入参:blnEditSta为True表示设置为可以修改的状态
    '编制:刘兴洪
    '日期:2014-06-06 17:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cboNO.Locked = Not blnEditSta
    
    cbo费别.Locked = Not blnEditSta: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
    cbo医疗付款.Locked = Not blnEditSta
    
    cbo开单科室.Locked = Not blnEditSta
    cbo开单人.Locked = Not blnEditSta
    cbo开单科室.Enabled = blnEditSta
    cbo开单人.Enabled = blnEditSta
    
    chk加班.Enabled = blnEditSta
    
    txtDate.Enabled = blnEditSta
    fraStat.Enabled = blnEditSta
    Bill.Active = blnEditSta
    SetPatientEnableModi (blnEditSta)
End Sub

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


Private Function GetDeptByRegevent(ByVal lng病人ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID返回有效挂号单的科室ID集
    '入参:lng病人ID-病人ID
    '返回:返回有效挂号单科室的数据集
    '编制:刘兴洪
    '日期:2014-06-06 17:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动加收挂号费
    '入参:lng病人ID-病人ID
    '     str病人姓名-病人姓名
    '编制:刘兴洪
    '日期:2014-06-06 17:41:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取Combox的索引值
    '返回:返回索引值,未找到时,返回-1
    '编制:刘兴洪
    '日期:2014-06-06 17:42:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    MCPAR.门诊连续收费 = gclsInsure.GetCapability(support门诊连续收费, lng病人ID, mintInsure)
    MCPAR.多单据收费 = gclsInsure.GetCapability(support多单据收费, lng病人ID, mintInsure)
    MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, mintInsure)
    MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, mintInsure)
    MCPAR.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, lng病人ID, mintInsure)
    MCPAR.全自付 = gclsInsure.GetCapability(support收费帐户全自费, lng病人ID, mintInsure)
    MCPAR.实时监控 = gclsInsure.GetCapability(support实时监控, lng病人ID, mintInsure)
    MCPAR.医保不走票号 = False
    '刘兴洪:27536 20100119
    MCPAR.不提醒缴款金额不足 = gclsInsure.GetCapability(support不提醒缴款金额不足, lng病人ID, mintInsure)
    MCPAR.多单据分单据结算 = gclsInsure.GetCapability(support多单据分单据结算, lng病人ID, mintInsure)
    MCPAR.门诊结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, mintInsure)
    MCPAR.一次结算分单据退费 = gclsInsure.GetCapability(support一次结算分单据退费, lng病人ID, mintInsure)
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
        
        '调用外挂部件接口
        If PatiValiedCheckByPlugIn(mlngModul, lng病人ID) = False Then
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
        
        '问题:28240
        strTemp = mstrYBPati: intInsure = mintInsure
            
        If GetPatient("-" & lng病人ID, , , True) Then
            mstrYBPati = strTemp: mintInsure = intInsure
            If Not CheckRegisted(lng病人ID, mblnPeisPriceBill) Then
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
            cmdAddBill.Enabled = Not MCPAR.门诊连续收费 And MCPAR.多单据收费 And zlStr.IsHavePrivs(mstrPrivs, "医保病人多单据收费")
        End If
        txtPatient.ForeColor = vbRed
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), vbRed)
            End If
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
        Dim cur透支额 As Currency
        cur透支额 = RoundEx(mdbl个帐透支, 2)
        
        mdbl个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, cur透支额, mintInsure)
        sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mdbl个帐余额, "0.00")
        sta.Panels(Pan.C3个人帐户).Visible = True
        mdbl个帐透支 = cur透支额
        
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
        If mbytInState = EM_ED_收费 And Visible And mstrInNO = "" And txtIn.Text = "" And mrsInfo.State = 1 And _
            Not (lngCur病人ID > 0 And Not MCPAR.门诊连续收费 And MCPAR.多单据收费 And InStr(1, mstrPrivs, "医保病人多单据收费") > 0) Then
            If gblnCheckRegeventDept And gint病人来源 = 1 And IsRegisterDept Then lng挂号科室 = Val("" & mrsInfo!执行部门ID)
            blnPriceBill = LoadMultiBills(lng病人ID, MCPAR.门诊连续收费 Or Not MCPAR.多单据收费 Or zlStr.IsHavePrivs(mstrPrivs, "医保病人多单据收费") = False, lng挂号科室)
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
        End If

        '处理预交结算
        '联合医保不使用预交冲款(划价模式)
        '咸阳医保不使用预交冲款
        If Not mblnSaveAsPrice And mintInsure <> 61 Then Call LoadFeeInfor(lng病人ID)
        
        '咸阳医保不缴款
'        If mintInsure = 61 Then Call ShowPayInfo(False)
                
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
            
            If gbln划价立即缴款 And blnPriceBill And mstrYBPati <> "" Then
                If cmd预结算.Visible And cmd预结算.Enabled Then
                    cmd预结算.SetFocus
                End If
            End If
            
            If gbyt科室医生 <> 0 Then
                If blnPriceBill Then
                    If cbo开单科室.Enabled And cbo开单科室.Visible And cbo开单科室.ListIndex < 0 Then
                        cbo开单科室.SetFocus
                    Else
                        If cmd预结算.Visible And cmd预结算.Enabled Then
                            cmd预结算.SetFocus
                        Else
                            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                        End If
                    End If
                Else
                    If cbo开单科室.Enabled And cbo开单科室.Visible Then cbo开单科室.SetFocus
                End If
            Else
                If blnPriceBill Then
                    If cbo开单人.Enabled And cbo开单人.Visible And cbo开单人.ListIndex < 0 Then
                        cbo开单人.SetFocus
                    Else
                        If cmd预结算.Visible And cmd预结算.Enabled Then
                            cmd预结算.SetFocus
                        Else
                            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                        End If
                    End If
                Else
                    If cbo开单人.Enabled And cbo开单人.Visible Then cbo开单人.SetFocus
                End If
            End If
            
            Call ShowWelcomeByLed
            Call ReInitPatiInvoice
        End If
    Else
        mintInsure = 0: mdbl个帐余额 = 0: mdbl个帐透支 = 0
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
    If KeyAscii = 13 And mbytInState = EM_ED_收费 And gint病人来源 = 1 And Not mblnValid Then
        If txtPatient.Text = "" And chkCancel.Value = 0 And zlStr.IsHavePrivs(mstrPrivs, "保险收费") Then
            Call MCPatientProcess
            Exit Sub
        End If
    End If
    If txtPatient.Locked Then Exit Sub '锁定状态只允许医保验卡
   
   '问题:51488
    If (IDKind.Cards.读卡快键 = "空格键" Or IDKind.Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
   
    blnCheckReg = False
    
    If mblnAutoChangePati And gint病人来源 = 2 And (KeyAscii <> 13) Then
        '需要切找到病人来源1中
        gint病人来源 = 1: zlChangePatiSource (gint病人来源)
    End If
    
 
       
    '3.正常输入病人(姓名各种标识)部份:住院病人收费时可弹出选择器
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And mbytInState = EM_ED_收费 And Trim(txtPatient.Text) = "" _
        And Not mblnValid Then
        If gint病人来源 = 2 Then
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
        If gint病人来源 = 1 And zlStr.IsHavePrivs(mstrPrivs, "允许非医保病人") = False Then
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
 
        sta.Panels(Pan.C2提示信息) = ""
        lblTotal.Caption = "合计:"
        
        '收费保持病人ID
        If txtPatient.Text = mstrPrePati And mlngPrePati <> 0 Then
            strPati = "-" & mlngPrePati
        Else
            strPati = txtPatient.Text
        End If
        
        If IDKind.GetCurCard.名称 Like "IC卡*" And IDKind.GetCurCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        If IDKind.GetCurCard.名称 Like "*身份证*" And IDKind.GetCurCard.系统 Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        int上次病人来源 = gint病人来源
        
        '50200(防止窗口找开过长,发生时间与登记时间拉得过长)
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
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
                If gint病人来源 = 1 And gblnInputName And IDKind.IDKind = IDKind.GetKindIndex("姓名") And txtPatient.Text <> "" Then
                    If mbytInState = EM_ED_收费 And mstrInNO = "" Then
                        If Not CheckRegisted(0, mblnPeisPriceBill) Then
                           Call ClearPatientInfo(True): Exit Sub
                        End If
                    End If
                    If mbytInState = EM_ED_收费 Then
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
                    
                    If mbytInState = EM_ED_收费 And Not mblnValid And Visible And mstrInNO = "" And txtIn.Text = "" Then
                        Call LoadAddedItem(0, txtPatient.Text)
                    End If
                    
                    If mobjBill.Pages(mintPage).NO = "" Then
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
                        If gTy_Module_Para.byt缴款控制 <> 1 _
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
                    
                Else
                    MsgBox "请检查输入内容,不能读取" & gstrCustomerAppellation & "信息！", vbInformation, gstrSysName
                    Call ClearPatientInfo(True)
                    Exit Sub
                End If
            End If
            
        Else 'b.根据输入读取病人信息成功
            lng病人ID = Val("" & mrsInfo!病人ID)
            Call InitBalanceGrid(True)
            Call Set连续收费操作
            
            If mbytInState = EM_ED_收费 And mstrInNO = "" And gint病人来源 = 1 Then
                If Not CheckRegisted(lng病人ID, mblnPeisPriceBill) Then
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                End If
            End If
            If mbytInState = EM_ED_收费 Then
                '问题:29283
                 '  -- 参数:调用场合-1-挂号;2-收费
                 '  --        病人id_In-病人ID(未建档的,传入零)
                 '  --        卡号_In: 刷卡卡号;未刷卡时,为空
                 '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
                 If zlPatiCardCheck(2, lng病人ID, IIf(blnCard Or blnICCard, txtPatient.Text, ""), 1) = False Then
                    '恢复上次病人来源
                    If int上次病人来源 <> gint病人来源 And mTy_Para.bln住院病人门诊收费 = False Then
                        Call zlChangePatiSource(int上次病人来源)
                    End If
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                     Exit Sub
                 End If
            End If
            
            
            '就诊卡密码检查
            If mbytInState = EM_ED_收费 And (blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.接口序号 <> 0) And mstrPassWord <> "" Then
                If Mid(gstrCardPass, 3, 1) = "1" Then
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
            
            '102234,调用外挂部件接口
            If PatiValiedCheckByPlugIn(mlngModul, lng病人ID) = False Then
                '恢复上次病人来源
                If int上次病人来源 <> gint病人来源 And mTy_Para.bln住院病人门诊收费 = False Then
                    Call zlChangePatiSource(int上次病人来源)
                End If
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                Exit Sub
            End If
                
            '连续划价或收费时,不是同一个病人时，记帐没有保留病人信息
            If Not IIf(mlngPrePati = 0, mstrPrePati = "" & mrsInfo!姓名, mlngPrePati = lng病人ID) Then
                '清除医生
                If mbytInState = EM_ED_收费 And mstrInNO = "" Then
                    If gbyt科室医生 = 0 And CheckBillsEmpty Then
                        For i = 1 To mobjBill.Pages.Count
                            mobjBill.Pages(i).开单部门ID = 0: mobjBill.Pages(i).开单人 = ""
                        Next
                        cbo开单人.ListIndex = -1: cbo开单科室.ListIndex = -1: lblDuty.Caption = ""
                    End If
                End If
                
                Call ClearPatientInfo
                
                '刘兴洪:22343
                If Not gTy_Module_Para.byt缴款控制 = 1 _
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
                If gbyt科室医生 <> 0 And mbytInState = EM_ED_收费 And mstrInNO = "" Then
                    '仅新增单据时,取住院病人的开单部门:科室确定医生或各自独立输入
                    Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, Val("" & mrsInfo!当前科室id)))
                    Call cbo开单科室_Click
                End If
            ElseIf gint病人来源 = 1 Then
                If mbytInState = EM_ED_收费 And mstrInNO = "" Then
                    Call SetDeptDoctorByRegevent(lng病人ID) '根据病人挂号信息设置开单科室和医生
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
            
            If glngSys Like "8??" Then
                cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(mrsInfo!费别), True)
                cbo费别.Locked = False: cbo费别.TabStop = Not cbo费别.Locked And gbln费别
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
            cbo医疗付款.Locked = gint病人来源 = 2
            

            '设置对象中的病人信息
            With mobjBill
                .病人ID = lng病人ID
                .主页ID = Nvl(mrsInfo!主页ID, 0)
                .标识号 = IIf(gint病人来源 = 2, Nvl(mrsInfo!住院号, 0), Nvl(mrsInfo!门诊号, 0))
                .姓名 = "" & mrsInfo!姓名
                .性别 = "" & mrsInfo!性别
                .年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
                .床号 = "" & mrsInfo!当前床号
                .病区ID = Nvl(mrsInfo!当前病区ID, 0)
                .科室ID = Nvl(mrsInfo!当前科室id, 0)
                .费别 = zlStr.NeedName(cbo费别.Text) '以当前有效为准
            End With
            Call ReInitPatiInvoice
            
            '关联操作处理
            If Not mblnValid And Visible Then
                '不是同一个病人时
                If Not (IIf(mlngPrePati = 0, mstrPrePati = mobjBill.姓名, mlngPrePati = mobjBill.病人ID) And txtPatient.Text <> "") Then
                     Call AddCardFee '产生就诊卡费用行
                End If
                
                '读取病人的多张划价单
                If mbytInState = EM_ED_收费 And mstrInNO = "" And txtIn.Text = "" Then
                    If mobjBill.病人ID <> 0 Then
                        If gblnCheckRegeventDept And gint病人来源 = 1 And IsRegisterDept Then lng挂号科室 = Val("" & mrsInfo!执行部门ID)
                       blnHavePriceBill = LoadMultiBills(mobjBill.病人ID, InStr(1, mstrPrivs, "普通病人多单据收费") = 0, lng挂号科室, blnCard)
                    End If
                    Call LoadAddedItem(mobjBill.病人ID, mobjBill.姓名)
                End If
                '光标定位
                If mstrInNO = "" Then
                    If mbytInState = EM_ED_收费 And txtPatient.Text = "新病人" Then
                        txtPatient.SetFocus
                        Call txtPatient_GotFocus
                    Else
                        If cbo医疗付款.ListIndex = -1 And gbln医疗付款 Then
                            If cbo医疗付款.Enabled And cbo医疗付款.Visible Then cbo医疗付款.SetFocus
                        Else
                            If gbln划价立即缴款 And blnHavePriceBill Then
                                If mstrYBPati <> "" And cmd预结算.Enabled And cmd预结算.Visible Then
                                    Call cmd预结算.SetFocus
                                Else
                                    Call ShowWelcomeByLed '显示欢迎信息和病人信息
                                    Call cmdOK_Click: Exit Sub
                                End If
                            End If
                            
                            If gbyt科室医生 = 0 Then
                                If blnHavePriceBill Then
                                    If cbo开单人.Enabled And cbo开单人.Visible And cbo开单人.ListIndex < 0 Then
                                        cbo开单人.SetFocus
                                    Else
                                        If cmd预结算.Visible And cmd预结算.Enabled Then
                                            cmd预结算.SetFocus
                                        Else
                                            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                                        End If
                                    End If
                                Else
                                    If cbo开单人.Enabled And cbo开单人.Visible Then cbo开单人.SetFocus
                                End If
                            ElseIf glngSys Like "8??" Then
                                Bill.SetFocus
                            Else
                                If blnHavePriceBill Then
                                    If cbo开单科室.Enabled And cbo开单科室.Visible And cbo开单科室.ListIndex < 0 Then
                                        cbo开单科室.SetFocus
                                    Else
                                        If cmd预结算.Visible And cmd预结算.Enabled Then
                                            cmd预结算.SetFocus
                                        Else
                                            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                                        End If
                                    End If
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
        
    If mstrCardNO = "" And Bill.Active Then
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示欢迎信息和病人信息
    '编制:刘兴洪
    '日期:2014-06-06 17:56:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strInfo As String, lngPatient As Long
    If gblnLED = False Then Exit Sub
    If mbytInState <> EM_ED_收费 Then Exit Sub
    If gblnLedWelcome Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModul, gcnOracle
    End If
    strInfo = Trim(txtPatient.Text)
    If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val("" & mrsInfo!病人ID)
    zl9LedVoice.DisplayPatient strInfo, lngPatient
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
    mlngPreBrushCardID = 0: mlngCardTypeID = 0
    
ReDO:
    blnCancel = False
    
    If mstrYBPati = "" Then
        If gint病人来源 = 1 Then
            'strWhere = " And Nvl(A.当前科室ID,0)=0"
             strWhere = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
        ElseIf gint病人来源 = 2 Then
            strWhere = " And Nvl(A.当前科室ID,0)<>0"
        End If
    End If
    
    '读取病人信息
    '76451,冉俊明,2014-8-19
    strSQL = "" & _
        "   Select " & strMoney & "Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型," & _
                        IIf(gint病人来源 = 1, "NULL", "Decode(A.当前科室ID,NULL,NULL,A.主页ID)") & " as 主页ID,A.IC卡号,A.就诊卡号,A.卡验证码,A.门诊号,A.住院号,A.姓名," & _
        "               A.性别,A.年龄,C.名称 险类名称, A.出生日期,A.费别,A.担保额,A.医疗付款方式,A.工作单位,A.当前病区ID,A.当前科室ID,A.当前床号,A.在院," & _
        "               decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.入院日期" & _
        "   From 病人信息 A,病案主页 B1,保险类别 C  " & _
        "   Where A.险类 = C.序号(+) And A.病人ID=B1.病人ID(+) And A.主页ID=B1.主页ID(+) And A.停用时间 is NULL"
    
    If blnYbCheckCard = False And blnCard And IDKind.GetCurCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        If gint病人来源 = 1 And Not gblnInputCard Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        
        '见问题:27364
        If gint病人来源 = 1 Then strWhere = ""
        
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
        mlngPreBrushCardID = lng卡类别ID
        
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Or blnYbCheckCard Then '病人ID
        If gint病人来源 = 1 And (Not gblnInputID And mstrYBPati = "") _
            And Not (mstrInNO <> "" And mbytInState = EM_ED_收费) Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        If gint病人来源 = 1 Then strWhere = ""
        strSQL = strSQL & strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        If gint病人来源 = 1 And Not gblnInputID Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        If gint病人来源 = 1 Then strWhere = ""
        strSQL = strSQL & strWhere & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        If gint病人来源 = 1 And Not gblnInputID Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        If gint病人来源 = 1 Then strWhere = ""
        strSQL = strSQL & strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "." Then '挂号单号(最后为执行部门ID以区分)
        If gint病人来源 = 1 And Not gblnInputNO Then
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
            "   Select " & strMoney & "Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型," & _
                                IIf(gint病人来源 = 1, "NULL", "Decode(A.当前科室ID,NULL,NULL,A.主页ID)") & " as 主页ID,A.就诊卡号,A.卡验证码,Nvl(B.标识号,A.门诊号) as 门诊号," & _
            "               A.住院号,B.姓名,B.性别,B.年龄,C.名称 险类名称, A.出生日期,B.费别,A.担保额,A.医疗付款方式,A.工作单位,A.当前病区ID,A.当前科室ID,A.当前床号,B.执行人,B.执行部门ID,A.在院," & _
            "               decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.入院日期" & _
            " From 病人信息 A,病案主页 B1,门诊费用记录 B,保险类别 C " & _
            " Where B.病人ID=A.病人ID (+) " & _
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
                If Not mblnValid And gblnSeekName And gblnInputID Then
                    strPati = _
                        " Select /*+Rule */1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                                    IIf(gint病人来源 = 2, "A.住院号,B.名称 as 科室,A.当前床号 as 床号,", "A.门诊号,") & _
                        "           A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                        " From 病人信息 A,部门表 B" & _
                        " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And Rownum <101 " & strWhere & " And A.姓名 Like [1]" & _
                        IIf(gintNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                    
                    '门诊病人收费时可以不对应病人档案
                    If gint病人来源 = 1 Then
                        strPati = strPati & " Union ALL " & _
                            "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                    End If
                    strPati = strPati & " Order by 排序ID,姓名"
                        
                    vRect = zlControl.GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人" & 0 & gint病人来源, 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays, "bytSize=1")
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
                
                If gint病人来源 = 1 Then strWhere = ""
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
                If gint病人来源 = 1 Then strWhere = ""
                strSQL = strSQL & strWhere & " And   A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                If gint病人来源 = 1 Then strWhere = ""
                strSQL = strSQL & strWhere & " And A.病人ID=[1]"
               blnHavePassWord = True
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
                    mlngPreBrushCardID = lng卡类别ID
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                If gint病人来源 = 1 Then strWhere = ""
                strSQL = strSQL & strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
        
    On Error GoTo errH
    If strSQL <> "" Then
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, strTemp)
        If Not mrsInfo.EOF Then
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类名称), Me.ForeColor, vbRed))
            If gint病人来源 = 1 And mTy_Para.bln住院病人门诊收费 = False Then
                '需要检查是否为在院病人
                '问题:27364 日期:2010-01-13 15:27:50
                If Val(Nvl(mrsInfo!在院)) = 1 Then
                        If gbln病人来源受权限控制 And zlStr.IsHavePrivs(mstrPrivs, "参数设置") = False Then
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
    If mbytInState = EM_ED_收费 And Trim(txtPatient.Text) <> "" Then
        mobjBill.姓名 = txtPatient.Text
        mobjBill.年龄 = Trim(txt年龄.Text) & IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
        mobjBill.性别 = zlStr.NeedName(cboSex.Text)
    End If
    
    '===========================
    '82864,冉俊明,2015-3-2
    '将该段代码由txtPatient_Validate中调整到这里，因为在代码中如果使用SetFocus方法设置了焦点，则不会触发Validate事件
    '同时，该段代码也不是检查，所以也可以不用放在txtPatient_Validate中
    If mblnKeyReturn = False Then
        mblnValid = True: Call txtPatient_KeyPress(13): mblnValid = False
    Else
        mblnKeyReturn = False
    End If
    '===========================
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
'    If mblnKeyReturn = False Then
'        mblnValid = True: Call txtPatient_KeyPress(13): mblnValid = False
'    Else
'        mblnKeyReturn = False
'    End If
End Sub

Private Sub txtRePrint_GotFocus()
    Call zlControl.TxtSelAll(txtRePrint)
End Sub

Private Sub txtRePrint_KeyPress(KeyAscii As Integer)
    Dim strNos As String, strNo As String
    Dim strOper As String, vDate As Date, intInsure As Integer, blnVirtualPrint As Boolean
    Dim lng结帐ID As Long, lng病人ID As Long, lng结算序号 As Long
    Dim strReclaimInvoice As String, intInvoiceFormat As Integer '回收的票据
    Dim blnNOMoved As Boolean
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtRePrint, KeyAscii)
        Exit Sub
    End If
    '重打
    strNo = txtRePrint.Text
    strNo = GetFullNO(strNo, 13)
    txtRePrint.Text = strNo: zlControl.TxtSelAll txtRePrint
    blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "1", Me.Caption)
        
    '是否已转入后备数据表中
    If blnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False
    End If
    If Not ReadBillInfo(1, strNo, 1, strOper, vDate, lng病人ID) Then txtRePrint.SetFocus: Exit Sub
        
    If zlStr.IsHavePrivs(mstrPrivs, "所有操作员") = False Then
        If UserInfo.姓名 <> strOper Then
            MsgBox "你没有""所有操作员""权限,不能重打" & strOper & "的单据！", vbInformation, gstrSysName
            txtRePrint.Text = "": Exit Sub
        End If
    End If

    If Not BillOperCheck(2, strOper, vDate, "重打", txtRePrint.Text, , 1) Then
        txtRePrint.SetFocus: Exit Sub
    End If
    
    lng结帐ID = zlGetFirstBalanceID(strNo, blnNOMoved, False, lng结算序号)
    '可能是多单据收费中的一张
    If lng结算序号 >= 0 Then
        '针对老版本(10.34.0以前的数据)的数据进行重打
        Call FromBillNoReprintBill(strNo, blnNOMoved)
        Exit Sub
    End If
    
    strNos = zlGetBalanceNos(0, txtRePrint.Text, blnNOMoved)
    '单据有剩余数量的才可以重打
    If Not BillExistMoney(strNos, 1, True) Then
        MsgBox "单据不存在或已经全部退费,不能重打！", vbInformation, gstrSysName
        txtRePrint.Text = "": Exit Sub
    End If
    '调出重打的单据显示
    If frmClinicDelAndView.ShowMe(Me, EM_MULTI_查看, mstrPrivs, lng结算序号, True) = False Then Exit Sub
    intInsure = zlGetBillChargeExistInsure(lng结帐ID, lng病人ID)
    
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
        '此处只提供了收费票据的重打
    End If
    Call ReInitPatiInvoice(True, intInsure, lng病人ID)
    strReclaimInvoice = zlGetReclaimInvoice(strNo)
    If strReclaimInvoice <> "" Then
        '需要显示出本次需要回收的发票
        If MsgBox("注意:" & vbCrLf & " 请注意回收以下发票:" & vbCrLf & strReclaimInvoice, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Call RefreshFact '刷新票据号
            txtRePrint.Text = ""
            txtPatient.SetFocus
            Exit Sub
        End If
    End If
    If InStr(1, strNos, "'") = 0 Then
        strNos = "'" & Replace(strNos, ",", "','") & "'"
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
End Sub

Private Sub txtRePrint_LostFocus()
    txtRePrint.BackColor = vbWhite
End Sub
Public Function GetMustPaySum() As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:求本次收费的应缴合计，主要用于多单据收费模式
    '返回:成功,返回应缴合计
    '编制:刘兴洪
    '日期:2014-06-06 18:00:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curMoney As Currency, i As Integer
    For i = 1 To mobjBill.Pages.Count
        curMoney = curMoney + mobjBill.Pages(i).应缴金额
    Next
    GetMustPaySum = curMoney
End Function

Private Function Get中药数量(ByRef str计算单位 As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取当前单据中中药的数量，如果存在不同单位的药品，则返回为0
    '返回:返回中药数量
    '编制:刘兴洪
    '日期:2014-06-06 18:00:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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
    If mbytInState = EM_ED_收费 And mstrInNO = "" And gbytAutoSplitBill > 0 And Not (mstrYBPati <> "" And MCPAR.门诊预结算) Then
        Call AutoSplitBill
    End If
    '收费时自动产生工本费项目:修改时不管工本费
    If mbytInState = EM_ED_收费 And gTy_Module_Para.bln工本费 Then
        If Not CheckBillsEmpty Then Call SetFactMoney
    End If
End Sub
 

Private Sub CalcMoneys(Optional intPage As Integer, Optional lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算或重新计算指定行或所有行的金额
    '入参:intPage,lngRow=指定单据页指定行,为0表示计算所有行
    '编制:刘兴洪
    '日期:2014-06-06 18:01:56
    '说明：ExpenseBill集合的索引对应单据的行号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, p As Integer
    Dim strMainRows As String
    Dim bln从项汇总折扣 As Boolean
        
    
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
                    Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                Me.Caption, .收费细目ID, .执行部门ID, dblAllTime)
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新显示当前单据指定行或所有行的内容
    '入参:lngRow=指定行,为0表示显示所有行
    '编制:刘兴洪
    '日期:2014-06-06 18:03:12
    '说明：ExpenseBill集合的索引对应单据的行号
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    
    If mbytInState = EM_ED_调整 Then
        curTotal = GetBillSum
        lblTotal.Caption = "合计:" & Format(curTotal, gstrDec)
    End If
End Sub

Private Sub ShowDetail(lngRow As Long, Optional intCurSubItem As Integer = 0, Optional intSubItemCount As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新显示指定行的内容
    '入参:lngRow=指定行
    '     intCurSubItem-加载的当前套餐
    '     intSubItemCount- 主要是针对套餐来说的,总共套餐项目数(是否为最后一笔)
    '编制:刘兴洪
    '日期:2014-06-06 18:04:03
    '说明：ExpenseBill集合的索引对应单据的行号
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
                    If mbytInState = EM_ED_收费 Then
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:：刷新显示收入项目费用区，不支持预结算时的保险结算区，单据合计等
    '入参:bln个帐=是否处理个人帐户显示
    '      intPage=是否只重新计算指定单据(加快速度)，0-全部计算,-1,全不计算,x-计算指定单据
    '编制:刘兴洪
    '日期:2014-06-06 18:04:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, arrDetail As Variant
    Dim cur冲款合计 As Currency, cur实收金额 As Currency, cur可用个帐 As Currency
    Dim cur个帐 As Currency, curTotal As Currency
    Dim cur全自付 As Currency, cur先自付 As Currency, cur进入统筹 As Currency
    Dim cur实收合计 As Currency, cur应收合计 As Currency, strTmp As String
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim blnExist As Boolean, blnDo As Boolean, strSQL As String

    '产生汇总费目,并统计保险相关金额
    '-------------------------------------------------------------------------
        
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
        
        If MCPAR.多单据分单据结算 And Not MCPAR.门诊预结算 Then
            '更新当前单据个人帐户支付金额:不支持预结算时
            '医保病人且满足相应条件才处理,合计为负不能退到个人帐户
            If mstrYBPati <> "" And bln个帐 And mstr个人帐户 <> "" Then
                If mdbl个帐余额 > -1 * mdbl个帐透支 Then
                    If cur实收合计 >= 0 Then
                        cur个帐 = cur进入统筹 + IIf(MCPAR.先自付, cur先自付, 0) + IIf(MCPAR.全自付, cur全自付, 0)
                        
                        '统计除开之前单据个帐支付后的个帐余额
                        cur可用个帐 = 0
                        For i = 1 To p - 1
                            cur可用个帐 = cur可用个帐 + GetMedicareSum(mcolBalance, mstr个人帐户, i)
                        Next
                        cur可用个帐 = mdbl个帐余额 - cur可用个帐
                                            
                        '计算个人帐户支付金额
                        If cur可用个帐 - cur个帐 >= -1 * mdbl个帐透支 Then
                            Call SetBalanceVal(mcolBalance, p, mstr个人帐户 & "|" & Format(cur个帐, "0.00"))  '在允许透支范围内足够(允许透支0为特例)
                        Else
                            If mdbl个帐透支 = 0 And cur可用个帐 > 0 Then
                                Call SetBalanceVal(mcolBalance, p, mstr个人帐户 & "|" & Format(cur可用个帐, "0.00"))  '不允许透支且有余额
                            Else
                                '超过允许透支范围或不允许透支时无余额
                                If mdbl个帐透支 <> 0 Then
                                    Call SetBalanceVal(mcolBalance, p, mstr个人帐户 & "|" & cur可用个帐 + mdbl个帐透支)   '在允许透支范围内支付
                                Else
                                    Call SetBalanceVal(mcolBalance, p, mstr个人帐户 & "|" & 0)
                                End If
                            End If
                        End If
                    Else
                        Call SetBalanceVal(mcolBalance, p, mstr个人帐户 & "|" & 0)
                    End If
                Else
                    Call SetBalanceVal(mcolBalance, p, mstr个人帐户 & "|" & 0)
                End If
            End If
        End If
        
        '当前单据的相关汇总金额计算
        '----------------------------------------
        With mobjBill.Pages(p)
            .应收金额 = cur应收合计
            .实收金额 = cur实收合计
            
            .进入统筹 = cur进入统筹
            .全自付 = cur全自付
            .先自付 = cur先自付
            
            '医保支付的所有金额,可能为预结算返回的,也可能是该过程计算的
            .保险金额 = 0
                        
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
            .应缴金额 = RoundEx(.实收金额 - .保险金额 - .冲预交额 - .消费卡刷卡额, 7)
            
            'Key值的保存,用于快速计算
            strTmp = ""
            For i = 0 To UBound(arrDetail)
                strTmp = strTmp & ";" & Split(arrDetail(i), ",")(0) & "," & _
                    Split(arrDetail(i), ",")(1) & "," & Split(arrDetail(i), ",")(2)
            Next
            .Key = Mid(strTmp, 2)
        End With
    Next
    
    If Not MCPAR.多单据分单据结算 And Not MCPAR.门诊预结算 Then
        '更新当前单据个人帐户支付金额:不支持预结算时
        '医保病人且满足相应条件才处理,合计为负不能退到个人帐户
        If mstrYBPati <> "" And bln个帐 And mstr个人帐户 <> "" Then
            If mdbl个帐余额 > -1 * mdbl个帐透支 Then
                If cur实收合计 >= 0 Then
                    For i = 1 To mobjBill.Pages.Count
                        cur个帐 = cur个帐 + mobjBill.Pages(i).进入统筹 + IIf(MCPAR.先自付, mobjBill.Pages(i).先自付, 0) + IIf(MCPAR.全自付, mobjBill.Pages(i).全自付, 0)
                    Next
                    cur可用个帐 = mdbl个帐余额
                    '计算个人帐户支付金额
                    If cur可用个帐 - cur个帐 >= -1 * mdbl个帐透支 Then
                        Call SetBalanceVal(mcolBalance, 1, mstr个人帐户 & "|" & Format(cur个帐, "0.00"))    '在允许透支范围内足够(允许透支0为特例)
                    Else
                        If mdbl个帐透支 = 0 And cur可用个帐 > 0 Then
                            Call SetBalanceVal(mcolBalance, 1, mstr个人帐户 & "|" & Format(cur可用个帐, "0.00"))   '不允许透支且有余额
                        Else
                            '超过允许透支范围或不允许透支时无余额
                            If mdbl个帐透支 <> 0 Then
                                Call SetBalanceVal(mcolBalance, 1, mstr个人帐户 & "|" & cur可用个帐 + mdbl个帐透支)   '在允许透支范围内支付
                            Else
                                Call SetBalanceVal(mcolBalance, 1, mstr个人帐户 & "|" & 0)
                            End If
                        End If
                    End If
                Else
                    Call SetBalanceVal(mcolBalance, 1, mstr个人帐户 & "|" & 0)
                End If
            Else
                Call SetBalanceVal(mcolBalance, 1, mstr个人帐户 & "|" & 0)
            End If
        End If
    End If
    
    '刷新显示所有单据的个人帐户支付情况
    '-------------------------------------------------------------------------
    If mstrYBPati <> "" And bln个帐 And mstr个人帐户 <> "" And mdbl个帐余额 > -1 * mdbl个帐透支 Then
        If Not MCPAR.门诊预结算 Then
            With vsBalance
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, 0) = mstr个人帐户 Then Exit For
                Next
                If i <= .Rows - 1 Then
                    .TextMatrix(i, 1) = Format(GetMedicareSum(mcolBalance, mstr个人帐户), "0.00")
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
    mdbl应缴合计 = RoundEx(GetMustPaySum + mcurBill应缴 - GetMedicareSum(mcolBalance), 6)
End Sub

Private Function GetInputDetail(ByVal lng项目id As Long) As Detail
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取收费项目信息
    '编制:刘兴洪
    '日期:2014-06-06 18:07:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
    '编制:刘兴洪
    '日期:2014-06-06 18:08:04
    '说明:
    '      1.用于新输入或更改收费细目行！！！
    '      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在
    '---------------------------------------------------------------------------------------------------------------------------------------------

 
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断该行是否应该取从属项目
    '返回:是从属项目返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-06 18:08:30
    '说明：仅该行收费项目有从属项目及尚未取才取。
    '---------------------------------------------------------------------------------------------------------------------------------------------

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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前行的项目是否具有主项目
    '入参:lngRow-当前行号
    '     intPage-指定页
    '返回:主项目返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-06 18:09:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回一个收费细目的从属项目集
    '编制:刘兴洪
    '日期:2014-06-06 18:10:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除指定收费项目行
    '编制:刘兴洪
    '日期:2014-06-06 18:10:25
    '说明：这时不处理从属行的删除,但要对其它单据行从属关系作相应的调整
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:用于医保连续收费时调用,连续收费模式下不能使用多单据收费
    '编制:刘兴洪
    '日期:2014-06-06 18:10:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim i As Integer
    
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '多单据收费:恢复缺省单据页卡
    mintPage = 1
    If fraBill.Visible Then
        cmdAddBill.Enabled = zlStr.IsHavePrivs(mstrPrivs, "普通病人多单据收费")
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
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Call InitBalanceGrid
    Original.冲预交款 = 0
    Original.实收合计 = 0
    Original.应缴金额 = 0
    ''txt本次应缴.Visible = False: lbl应缴.Caption = "应缴"
      
    cboNO.Text = ""
    
    '刷新票据号,只有自用的时，在打印后已刷新
    Call RefreshFact
        
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

Private Function NewBill(Optional blnFact As Boolean = True, Optional bln费别 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化一张新的单据(程序对象)
    '入参:blnFact=是否取票号
    '      bln费别=是否重新初始化费别
    '编制:刘兴洪
    '日期:2014-06-06 18:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long
    Dim dtCurdate As Date '服务器当前时间
    
    Set mrsInfo = New ADODB.Recordset
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '多单据收费:恢复缺省单据页卡
    mintPage = 1
    
    Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
    If cmdIDCard.Visible Then cmdIDCard.Enabled = True
    If cmdRegist.Visible Then cmdRegist.Enabled = True
    
    cmdAddBill.Enabled = zlStr.IsHavePrivs(mstrPrivs, "普通病人多单据收费")
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
    mdbl个帐余额 = 0: mdbl个帐透支 = 0
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
    cbo医疗付款.Locked = False
    sta.Panels(Pan.C3个人帐户).Tag = "": sta.Panels(Pan.C3个人帐户).Text = "": sta.Panels(Pan.C3个人帐户).Visible = False
            
    Call InitBalanceGrid
    Call SetButton(2) '确定,取消
    Call ShowPrePayInfo(False) '预交信息初始
'    Call ShowPayInfo( True) '联合医保
    
    SetPatientEnableModi (True)
    txtRePrint.Enabled = True: txtIn.Enabled = True
    cboNO.Enabled = True: chkCancel.Enabled = True: cmdDelete.Enabled = True
        
    If gbyt科室医生 = 0 And mstrPrePati <> txtPatient.Text Then
        cbo开单人.ListIndex = -1: cbo开单科室.ListIndex = -1: lblDuty.Caption = ""
    End If
    
    dtCurdate = zlDatabase.Currentdate
    txtDate.Text = Format(dtCurdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbytInState = EM_ED_收费 Then
        cboNO.Text = ""
        mstrWarn = ""
        cmdOK.Tag = "": cmdCancel.Tag = "": cmdPrint.Tag = "": cmd预结算.Tag = ""
        txtInvoice.Text = ""
        Call ReInitPatiInvoice(blnFact)
        
        chk加班.Value = IIf(OverTime(dtCurdate), 1, 0)
        
        
        '费别处理：收费或划价
        If Not (glngSys Like "8??") Then
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除费用显示区
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-06 18:12:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
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

Private Function GetDrugWindow(ByVal lng药房ID As Long, ByVal str类别 As String, _
    ByVal intPage As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的发药窗口,如果参数指定了缺省,则以指定为准,否则,如果是划价单,则以第一药品行的窗口为准,否则以已输入相同药品的窗口为准
    '入参:intPage=搜录到的单据编号
    '返回:返回发药窗口
    '编制:刘兴洪
    '日期:2014-06-06 18:12:20
    '说明：主要用于多单据收费时，不同类别的药品可能动态分配到同一药房，这样他们的窗口也应相同，但强行指定的除外
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    If strPayWin = "" Then strPayWin = GetDefaultWindow(str类别, lng药房ID)
    
    If strPayWin <> "" Then
        '检查是否上班
        strSQL = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药房ID, strPayWin)
        If rsTmp.EOF Then strPayWin = ""
    End If
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
    Dim strInvoice As String, strDate As String, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim intCheckInsure As Integer
    
    On Error GoTo errHandle
    '并发检查
    If zlIsCheckExistErrBill(mlng结算序号) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng结算序号) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInvoice = Trim(txtInvoice.Text)
    If Not CheckBillNOAndBookeFee Then Exit Function

    Set mobjChargeInfor = Nothing
    If GetChargeInfor(mobjChargeInfor) = False Then Exit Function
    mobjChargeInfor.结帐ID = mlng结帐ID
    mobjChargeInfor.结算序号 = mlng结算序号
    mobjChargeInfor.Nos = zlGetBalanceNos(1, mobjChargeInfor.结帐ID, False)
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    mobjBill.登记时间 = CDate(strDate)
    
    '若为部分医保结算成功单据，则会进行重新收费，所以要先进行医保验证接口(Identifiy)
    '校对标志等于2则已成功结算
    '91914,多单据分单据结算不支持预结算时病人预交记录中有可能没有医保结算信息
    If mintInsure <> 0 And mstrYBPati = "" Then '已进行过医保验证的，不再验证
        intCheckInsure = mintInsure
        strSQL = "Select 1" & _
                " From 病人预交记录 A, 结算方式 B" & _
                " Where a.结算方式 = b.名称 And b.性质 In (3, 4) And Nvl(a.校对标志, 0) = 1" & _
                "       And a.结帐id = [1] And Rownum < 2"
        strSQL = strSQL & "Union All" & _
                " Select 1" & _
                " From 保险结算记录" & _
                " Where 记录id = [1] " & _
                "       And Not Exists(Select 1 From 病人预交记录 A, 结算方式 B" & _
                "                       Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结帐id = 记录id)" & _
                "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID)
        If Not rsTemp.EOF Then
            'strAdvace:医保部分退时:传入1,表示医保部分退后再重新收费的身份验证;其他传入: 空
            mstrYBPati = gclsInsure.Identify(0, mobjBill.病人ID, intCheckInsure)
            
            If mstrYBPati = "" Then
                MsgBox "医保身份验证失败，不允许继续处理异常重收！", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
                Exit Function
            End If
            
            If Val(CLng(Split(mstrYBPati, ";")(8))) <> mobjBill.病人ID Then
                MsgBox "医保验证的病人与退费的病人不是同一个病人!", vbInformation, gstrSysName
                Call gclsInsure.IdentifyCancel(0, mobjBill.病人ID, intCheckInsure)
                Exit Function
            End If
        End If
    End If
    
    '重新收费时，将收费的登记时按新时间进行登记处理
    'Zl_门诊收费异常_Update
    strSQL = "Zl_门诊收费异常_Update("
    '  No_In       门诊费用记录.No%Type,
    strSQL = strSQL & "NULL,"
    '  登记时间_In 门诊费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  结帐id_In   门诊费用记录.结帐id%Type := Null
    strSQL = strSQL & "" & mobjChargeInfor.结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If Not frmClinicChargeBalance.zlChargeWin(Me, EM_FUN_重收, mlngModul, mstrPrivs, mobjChargeInfor _
        , , , , , mblnElsePersonErrBill) Then
        If Not gfrmMain Is Nothing And Not mblnErrBill Then Unload Me
        Exit Function
    End If
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, EM_EX_完成, False, mobjChargeInfor)
    End If
    ReChargeFee = True
    Exit Function
errHandle:
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
    Err = 0: On Error GoTo errHandle:
    Dim lng原结算序号 As Long, rsBalance As ADODB.Recordset, str结算方式 As String
    
    '并发检查
    If zlIsCheckExistErrBill(mlng结算序号) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng结算序号) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mbln连续输入 = False
    '获取结算信息
    Set mobjChargeInfor = Nothing
    If GetChargeInfor(mobjChargeInfor) = False Then Exit Function
    
    '医保结算方式不允许作废时，单据不允许作废
    If mobjChargeInfor.intInsure <> 0 Then
        If MCPAR.门诊结算作废 Then
            Set rsBalance = zlFromIDGetChargeBalance(0, mlng结帐ID)
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            rsBalance.Filter = "类型=2"
            With rsBalance
                Do While Not .EOF
                    If Not gclsInsure.GetCapability(support门诊结算作废, mobjChargeInfor.病人ID, _
                                        mobjChargeInfor.intInsure, Nvl(!结算方式)) Then
                        str结算方式 = str结算方式 & "," & Nvl(!结算方式)
                    End If
                    .MoveNext
                Loop
            End With
            If str结算方式 <> "" Then
                MsgBox "医保结算方式【" & str结算方式 & "】不支持作废，不能作废该单据！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "医保不支持门诊结算作废，不能作废该单据！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    mobjChargeInfor.Nos = zlGetBalanceNos(1, mlng结帐ID)
    If mbln作废异常 Then
        mobjChargeInfor.冲销ID = mlng结帐ID
        mobjChargeInfor.结算序号 = mlng结算序号
        mobjChargeInfor.结帐ID = zlGetFirstBalanceID(mobjChargeInfor.Nos, , , lng原结算序号)
    Else
        mobjChargeInfor.结帐ID = mlng结帐ID
        mobjChargeInfor.结算序号 = mlng结算序号
    End If
    If Not frmClinicChargeBalance.zlChargeWin(Me, EM_FUN_作废, mlngModul, mstrPrivs, mobjChargeInfor, , , , mbln作废异常) Then
        If Not gfrmMain Is Nothing Then
            mlng结算序号 = 0: Unload Me
        End If
        Exit Function
    End If
    
    Call WriteMzInforToCard(mobjBill.病人ID, mobjChargeInfor.结算序号, True)
    
    If Not gfrmMain Is Nothing Then
        mintSucces = mintSucces + 1
        mlng结算序号 = 0: Unload Me
    End If
    DelErrBillFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Unload Me
End Function

Private Function zlInsureClinicSwapPrice(ByVal strSaveNos As String, _
    ByRef strSaveSucessNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保调用
    '入参 strSaveNos-保存的单据号
    '出参:strSaveSucessNos-返回已经结算成功的单据号
    '返回:医保调用成功或非医保,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim p As Integer, blnTrans As Boolean, blnAffair As Boolean
    Dim varNos As Variant, blnTransMedicare As Boolean
    Dim strNo As String
   
   On Error GoTo errHandle
    '非医保，返回true,否则返加
    blnAffair = False
    If mstrYBPati = "" Or Not mblnSaveAsPrice Then zlInsureClinicSwapPrice = True: Exit Function
    blnTrans = True
    '1. 保存为划价单
    varNos = Split(Replace(strSaveNos, "'", ""), ",")
    For p = 0 To UBound(varNos)
        strNo = varNos(p)
        '保存为划价单
        '如果是联合医保,收费确定时实际却保存为划价单:传划价单明细,不在Oracle事务中执行
        If Not mnuFileSavePrice.Checked Then
            If Not gclsInsure.TranChargeDetail(1, strNo, 1, 0, "", , mintInsure) Then
                '删除划价单(继续处理)
                Call DelMedicareTempNO(True, strNo)
            Else
                strSaveSucessNos = strSaveSucessNos & "," & strNo
            End If
        End If
        gcnOracle.CommitTrans
        gcnOracle.BeginTrans: blnTrans = True
    Next
    zlInsureClinicSwapPrice = True
    Exit Function
errHandle:
    If blnTrans Then
         gcnOracle.RollbackTrans
        Call ErrCenter
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mintInsure)
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTransMedicare = False Then    '如果医保成功了，不删除划价单，费用失败可以重收
        Call DelMedicareTempNO(False, strNo)
    End If
    Call SaveErrLog
End Function

Private Sub DelMedicareTempNO(ByVal blnPriceSaved As Boolean, ByVal strBillNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:直接收费时,删除前一个事务提交的划价单
    '编制:刘兴洪
    '日期:2014-06-06 18:20:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not blnPriceSaved Then Exit Sub
    
    gstrSQL = "zl_门诊划价记录_DELETE('" & strBillNO & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Exit Sub
errHandle:
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
    Dim strInfor As String, i As Integer
    Dim rsTemp As ADODB.Recordset
    Dim varBalance As Variant, strTemp As String
    
    On Error GoTo errH:

    gstrSQL = "" & _
    "   Select decode(a.记录性质,1,'预存款',11,'预存款',结算方式) as 结算方式,  " & _
    "             nvl(sum(decode(nvl(校对标志,0),1, 1,0)* 冲预交),0) as 未结金额," & _
    "             nvl(sum(decode(nvl(校对标志,0),0,1,2,1,0)* 冲预交),0) as 结算金额" & _
    "   From 病人预交记录 A " & _
    "   Where 结帐ID=[1]" & _
    "   Group by  decode(a.记录性质,1,'预存款',11,'预存款',结算方式) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng结算序号)
    strInfor = ""
    With rsTemp
        dbl结算金额 = 0: dbl未结金额 = 0
        Do While Not .EOF
            If Val(Nvl(rsTemp!结算金额)) <> 0 Then
                strInfor = strInfor & vbCrLf & "    " & Nvl(rsTemp!结算方式) & ":" & Format(rsTemp!结算金额, "0.00")
            End If
            dbl未结金额 = dbl未结金额 + Val(Nvl(rsTemp!未结金额))
            dbl结算金额 = dbl结算金额 + Val(Nvl(rsTemp!结算金额))
            .MoveNext
        Loop
    End With
    If strInfor <> "" Then strInfor = Mid(strInfor, 3)
    
    '多单据分单据结算时，可能只有部分结算成功
    '医保先结算，所以只要strInfor不为空则表示医保已全部结算成功
    If MCPAR.多单据分单据结算 And strInfor = "" Then
        '返回结算信息,格式:结算方式|结算金额||...
        strTemp = zlGetYBBalanceNo(lng结算序号)
        varBalance = Split(strTemp, "||")
        For i = 0 To UBound(varBalance)
            dbl未结金额 = dbl未结金额 - Val(Split(varBalance(i), "|")(1))
            dbl结算金额 = dbl结算金额 + Val(Split(varBalance(i), "|")(1))
        Next
        strInfor = strInfor & "    " & Replace(Replace(strTemp, "||", vbCrLf & "    "), "|", ":")
    End If
    
    strInfor = "" & _
        "异常收费(请注意重新收取):" & vbCrLf & _
        "    当前已收取病人:" & Format(dbl结算金额, "0.00") & "元" & vbCrLf & _
        "    当前还未收取病人:" & Format(dbl未结金额, "0.00") & "元" & vbCrLf & _
        "收取成功的各项数据如下:" & vbCrLf & strInfor
    MsgBox strInfor, vbExclamation, gstrSysName
    '清除界面所有显示
    Call ClearPayInfo
    mstrInNO = ""
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据开单人或开单科室ID设置开单科室及开单人,但不触发点击事件
    '编制:刘兴洪
    '日期:2014-06-06 18:21:03
    '说明:利用公共函数CboSetIndex避免隐式调用cbo_click事件
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
            Call FillDept(cbo开单人.ItemData(cbo开单人.ListIndex))
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据开单人或开单科室ID设置开单科室及开单人,并触发点击事件
    '编制:刘兴洪
    '日期:2014-06-06 18:21:31
    '说明:当Listindex=x时,如果Listindex的值本身等于x,就不会触发点击事件,所以要用API+Click强制调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    Optional blnNoName As Boolean, _
    Optional blnShow As Boolean, Optional blnErrBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：1.读取主界面原始单据或已退单据,2.读取划价收费,记帐审核单据,3.读取要部份退的单据
    '调用：目前供以下操作调用
    '      1.提划价单收费或记帐，包括输单据号提划价单收费，确定病人身份后自动提取划价单收费，多张收费时切换到单据页时重新读划价单
    '      2.查看，调整，退费，销帐单据时读单据，包括读收费单，划价单，记帐单，记帐划价单
    '参数：strNo=单据号
    '      bytFun=0:收费单,1:划价单
    '      blnShow=是否是因为切换单据读取(仅显示内容)
    '      blnErrBill-显示异常单据
    '返回：blnNoName=病人姓名是否为空
    '说明：读取要退费的单据时(收费),排开误差处理费用,否则根据参数决定是否显示
    '      因为多次部份退费时,每次都可能产生误差,原始的误差始终退不完。
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
        "       A.开单部门ID,Nvl(A.加班标志,0) as 加班标志," & _
        "       Nvl(A.婴儿费,0) as 婴儿费,A.开单人,A.划价人,A.操作员姓名,A.发生时间,A.登记时间," & _
        "       B.医疗付款方式,Nvl(A.是否急诊,0) as 是否急诊,A.门诊标志,Nvl(A.医嘱序号,0) as 医嘱序号,A.摘要,A.记录状态" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A") & " ,病人信息 B,人员表 C" & _
        " Where Rownum=1 And Nvl(A.操作员姓名,A.划价人)=C.姓名" & _
        "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
        "       And mod(A.记录性质,10)=1" & _
        "       And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
        "       And NO=[1] And A.病人ID=B.病人ID(+)" & _
        IIf(bytFun = 1, " And A.操作员姓名 is Null And A.划价人 is Not NULL", "") & _
        IIf(mstrTime <> "", " And A.登记时间=[2]", "")
        
        If mstrTime <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime))
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        End If
        If rsTmp.EOF Then
            MsgBox "没有发现指定的单据！", vbInformation, gstrSysName
            Exit Function
        End If
        If bytFun = 1 And Not mblnDoing Then
            '病人ID不同，不能一起收费
            If mobjBill.Pages.Count > 1 _
                And Val(Nvl(rsTmp!病人ID)) <> 0 And mobjBill.病人ID <> 0 _
                And Val(Nvl(rsTmp!病人ID)) <> mobjBill.病人ID Then
                MsgBox "单据【" & strNo & "】的病人""" & rsTmp!姓名 & """与当前病人不是同一个病人，不能一起收费！", vbInformation, gstrSysName
                Exit Function
            End If
            If Not IsNull(rsTmp!姓名) And txtPatient.Text <> "" Then
                '判断是否相同病人，及要使用的病人信息
                If txtPatient.Text <> rsTmp!姓名 Then
                    If MsgBox("单据中病人为""" & rsTmp!姓名 & """，与当前病人不符，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
        Original.结帐ID = Nvl(rsTmp!结帐ID, 0) '用于医保门诊退费,一卡通单据修改
        If mbytBillSource <> 4 Then mbytBillSource = Val("" & rsTmp!门诊标志)   '只要有一张是体检,则认为全部是体检单据
        
    
        '病人相关信息提取:可能用于划价单收费,自动提取多张单据时不管
        '问题:30717,123609
        If Not IsNull(rsTmp!登记时间) Then
            mobjBill.登记时间 = CDate(Format(rsTmp!登记时间, "yyyy-mm-dd HH:MM:SS"))
        End If
        If Val(Nvl(rsTmp!记录状态)) = 0 Then
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
            mobjBill.性别 = Nvl(rsTmp!性别)
            'mobjBill.年龄 = Nvl(rsTmp!年龄)
            
            '病人姓名
            If chkCancel.Value = 0 And (IsNull(rsTmp!姓名) Or IIf(mlngPrePati = 0, mstrPrePati = mobjBill.姓名, mlngPrePati = mobjBill.病人ID)) Then
                '同一个病人:空姓名或相同姓名
                
                If IsNull(rsTmp!姓名) Then
                    blnNoName = True
                    If Val(Nvl(rsTmp!记录状态)) = 0 And mstrPrePati = "" Then
                            
                    Else
                        txtPatient.Text = mstrPrePati '缺省为上一个病人姓名
                    End If
                Else
                    txtPatient.Text = Nvl(rsTmp!姓名)
                    Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), txtPatient.ForeColor, vbRed))
                End If
            Else
                '不同的病人
                txtPatient.Text = Nvl(rsTmp!姓名)
                Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), txtPatient.ForeColor, vbRed))
                '刘兴洪:22343,51670
                If Not (gTy_Module_Para.byt缴款控制 = 1) _
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

            txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
                        
            If Not rsTmp!病人ID Is Nothing Then Call LoadFeeInfor(Val("" & rsTmp!病人ID))
            
            If Nvl(rsTmp!是否急诊, 0) = 1 Then chk急诊.Value = 1: chk急诊.Visible = True
            mblnDo = False: chk加班.Value = Nvl(rsTmp!加班标志, 0): mblnDo = True
        End If
    End If
    
    '开单部门,开单人
    Call Set开单人开单科室(mobjBill.Pages(mintPage).开单人, mobjBill.Pages(mintPage).开单部门ID)
    
    '收费读划价单时，目前允许修改开单人和开单科室,除非是医嘱发送过来的。
    If mbytInState = EM_ED_收费 And chkCancel.Value = 0 Then
        cbo开单人.Locked = False
        cbo开单科室.Locked = False
        
        If mobjBill.Pages(mintPage).医嘱序号 <> 0 Then
            If cbo开单人.ListIndex <> -1 Then cbo开单人.Locked = True
            If cbo开单科室.ListIndex <> -1 Then cbo开单科室.Locked = True
        End If
    End If
    
    '读取单据收费细目部份:分离发药时没有药房
    '---------------------------------------------------------------------------------------------
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
    "       And A.收费细目ID=X.药品ID And mod(A.记录性质,10)=1" & _
    "       And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
    " Group by Nvl(A.价格父号,A.序号),A.从属父号,A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
    "   B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志,A.发药窗口,X.药品ID,X." & gstr药房单位
    
    strSQL = strSQL & " Union ALL " & _
    " Select Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
    "       A.费别,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
    "       A.计算单位,max(A.医嘱序号) as 医嘱序号,Avg(Nvl(A.付数,1)) as 付数," & _
    "       Avg(" & intSign & "*A.数次) as 数次,Sum(A.标准单价) as 单价," & _
    "       Sum(" & intSign & "*A.应收金额) as 应收金额,Sum(" & intSign & "*A.实收金额) as 实收金额," & _
    "       A.执行部门ID,D.名称 as 执行部门,A.附加标志,A.发药窗口" & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录  A") & ",收费项目目录 B,收费项目类别 C,部门表 D" & _
    " Where A.收费类别 Not IN('5','6','7') And A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+) " & _
    "       And mod(A.记录性质,10)=1  " & _
    "       And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
    " Group by Nvl(A.价格父号,A.序号),A.从属父号,A.费别,C.编码,C.名称,A.收费细目ID,B.名称," & _
    "       B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.执行部门ID,D.名称,A.附加标志,A.发药窗口"
        
    strSQL = "Select" & _
        " A.序号,A.从属父号,A.费别,A.编码,A.类别,A.收费细目ID,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
        " A.计算单位,A.医嘱序号,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门ID,A.执行部门,A.附加标志,A.发药窗口" & _
        " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
        " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        " Order by A.序号"
        
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime), 1, 8, 24)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, "", 1, 8, 24)
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
        If bytFun = 1 And InStr(",5,6,7,", rsTmp!编码) > 0 Then
            j = j + 1
            '只有未分配发药窗口时才重新分配,以第一药品行为准
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
        Bill.TextMatrix(i, BillCol.付数) = Val(Nvl(rsTmp!付数))
        Bill.TextMatrix(i, BillCol.数次) = FormatEx(Val(Nvl(rsTmp!数次)), 5)
        Bill.TextMatrix(i, BillCol.单价) = Format(Val(Nvl(rsTmp!单价)), gstrFeePrecisionFmt)
        Bill.TextMatrix(i, BillCol.应收金额) = Format(Val(Nvl(rsTmp!应收金额)), gstrDec)
        Bill.TextMatrix(i, BillCol.实收金额) = Format(Val(Nvl(rsTmp!实收金额)), gstrDec)
        Bill.TextMatrix(i, BillCol.执行科室) = Nvl(rsTmp!执行部门)
        Bill.TextMatrix(i, BillCol.标志) = IIf(rsTmp!附加标志 = 1, "√", "")
        Bill.TextMatrix(i, BillCol.类型) = Nvl(rsTmp!费用类型)
        
        curBill应收 = curBill应收 + Val(Nvl(rsTmp!应收金额))
        curBill实收 = curBill实收 + Val(Nvl(rsTmp!实收金额))
        
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
    '显示单据小计
    lblSub应收.Caption = "应收:" & Format(curBill应收, gstrDec)
    lblSub实收.Caption = "实收:" & Format(curBill实收, gstrDec)
    lblAmount.Caption = ""
    
    '显示费别(包括一张单据中动态费别产生的多种费别)
    str费别 = Mid(str费别, 2)
    i = UBound(Split(str费别, ","))
    lbl动态费别.Visible = i = 0
    cbo费别.Visible = i = 0
    If i <> 0 Then
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
    
    If bytFun = 0 And blnErrBill Then
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
        '读取单据原始内容
        intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
        strSQL = _
        "   Select " & IIf(gint分类合计 = 0, "A.收据费目", IIf(gint分类合计 = 2, "'单据合计'", "B.名称")) & " as 名称," & _
        "       Sum(" & intSign & "*A.应收金额) as 应收金额," & _
        "       Sum(" & intSign & "*A.实收金额) as 实收金额 " & _
        "   From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录  A") & " ,收入项目 B" & _
        "   Where A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
        "       And MOD(A.记录性质,10)=1" & IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
        "       And A.NO=[1] And A.收入项目ID=B.ID" & _
                IIf(gint分类合计 = 2, "", " Group By " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称"))
        
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
            mshMoney.TextMatrix(i, 2) = Format(Val(Nvl(rsTmp!实收金额)), gstrDec)
            curBill应收 = curBill应收 + Val(Nvl(rsTmp!应收金额))
            curBill实收 = curBill实收 + Val(Nvl(rsTmp!实收金额))
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
        
        lblTotal.Caption = "合计:" & Format(curBill实收, gstrDec)
        
        '刷新收费累计
        If chkCancel.Value = 0 And gbln累计 And Not mblnDoing Then
            txt累计.Text = Format(GetChargeTotal, "0.00")
            txt累计.ToolTipText = "当前操作员今日收费累计额"
        End If
        
        '多单据收费支持:共用于各种单据
        With mobjBill.Pages(tbsBill.SelectedItem.Index)
            .NO = strNo
            .应收金额 = curBill应收
            .实收金额 = curBill实收
            
            '仅收费时收取划价单用
            If bytFun = 1 Then
                '47489
                If strPayDrugWins <> "" Then strPayDrugWins = Mid(strPayDrugWins, 2)
                tbsBill.SelectedItem.Tag = strPayDrugWins ' str发药窗口
                Call ShowMoney(mintPage) '只需要计算当前单据
            End If
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除单据中的工本费用(当不需要工本费时)
    '编制:刘兴洪
    '日期:2014-06-06 18:26:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费时设置、显示、计算工本费
    '     工本费自动加在当前显示的单据中
    '编制:刘兴洪
    '日期:2014-06-06 18:26:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除单据表格显示内容
    '编制:刘兴洪
    '日期:2014-06-06 18:26:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取当前单据中其它中药的付数
    '编制:刘兴洪
    '日期:2014-06-06 18:27:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取开单科室ID
    '编制:刘兴洪
    '日期:2014-06-06 18:27:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取所有可供选择的药房
    '编制:刘兴洪
    '日期:2014-06-06 18:27:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, bytDay As Byte
    Dim str药房 As String, lng开单科室ID As Long
    
    lng开单科室ID = mobjBill.科室ID     '开单科室优先
    If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    
    If str类别 = "4" Then
        strSQL = _
        " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
        "       And B.服务对象 IN(" & gint病人来源 & ",3) And B.部门ID=C.ID" & _
        "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
        "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
        "       And (A.病人来源 is NULL Or A.病人来源=" & gint病人来源 & ")" & _
        "       And (A.开单科室ID is NULL Or A.开单科室ID=[1] Or Exists (Select 1 From 病区科室对应 Where 科室id = [1] And a.开单科室id = 病区id))" & _
        "       And A.收费细目ID=[2]" & _
        " Order by B.服务对象,C.编码"
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

Public Function GetBillSum(Optional bln应收 As Boolean, Optional ByVal intPage As Integer) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据合计金额
    '入参:intPage=指定单据,否则为所有单据
    '编制:刘兴洪
    '日期:2014-06-06 18:28:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    If curTotal = 0 And tbsBill.Tabs.Count = 1 _
        And Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.项目) = "") Then
        intCol = IIf(bln应收, BillCol.应收金额, BillCol.实收金额)
        For i = 1 To Bill.Rows - 1
            If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                curTotal = curTotal + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
            End If
        Next
    End If
    GetBillSum = Format(curTotal, gstrDec)
End Function

Private Function Calc工本费(Optional ByVal intPage As Integer) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算工本费
    '返回:返回工本费
    '编制:刘兴洪
    '日期:2014-06-06 18:28:54
    '---------------------------------------------------------------------------------------------------------------------------------------------


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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前修改的费用单据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-20 16:31:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入合法的费用时间！", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    strSQL = "zl_病人费用记录_Update('" & cboNO.Text & "'," & 1 & "," & _
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
        If InStr("退费", Bill.TextMatrix(0, Bill.COLS - 1)) = 0 Then
            Bill.Redraw = False
            Bill.COLS = Bill.COLS + 1
            Bill.TextMatrix(0, Bill.COLS - 1) = "退费"
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
        If InStr("退费", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
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
    
    cbo开单人.Clear
    Call GetDoctor(lng科室ID, mrs开单人)
    
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

Private Function CheckExecuteDeptCanDo() As Boolean
'功能：检查单据中输入执行科室是否在已设置执行科室范围内
    Dim i As Integer, p As Integer, k As Integer
    Dim blnNotExist As Boolean, varDept As Variant
    Dim blnFind As Boolean
    Dim blnExistNotPrice As Boolean
    
    Err = 0: On Error GoTo errHandler
    '96357
    If gTy_Module_Para.str本机收费执行科室 <> "" Then
        blnNotExist = False
        varDept = Split(gTy_Module_Para.str本机收费执行科室, ",")
    ElseIf gTy_Module_Para.str已设置收费执行科室 <> "" Then
        blnNotExist = True
        varDept = Split(gTy_Module_Para.str已设置收费执行科室, ",")
    Else
        CheckExecuteDeptCanDo = True
        Exit Function
    End If
    
    For p = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(p).NO = "" Then blnExistNotPrice = True '非划价单
        If blnFind Then Exit For
        For i = 1 To mobjBill.Pages(p).Details.Count
            If blnFind Then Exit For
            For k = 0 To UBound(varDept)
                If mobjBill.Pages(p).Details(i).执行部门ID = Val(varDept(k)) Then
                    blnFind = True: Exit For
                End If
            Next
        Next
    Next
    
    '如果全部都是划价单则不用检查，在提取划价单时已检查
    If blnExistNotPrice = False Then
        CheckExecuteDeptCanDo = True
        Exit Function
    End If
    
    If blnNotExist And blnFind Then
        If mobjBill.Pages.Count > 1 Then
            MsgBox "第 " & p - 1 & " 张单据中第 " & i - 1 & " 行的项目的执行科室为本机不允许收费的执行科室！", vbInformation, gstrSysName
            tbsBill.Tabs(p - 1).Selected = True
        Else
            MsgBox "单据中第 " & i - 1 & " 行的项目的执行科室为本机不允许收费的执行科室！", vbInformation, gstrSysName
        End If
        Bill.SetFocus: Exit Function
    ElseIf blnNotExist = False And blnFind = False Then
        If mobjBill.Pages.Count > 1 Then
            MsgBox "第 " & p - 1 & " 张单据中的项目不存在执行科室为本机允许收费的执行科室！", vbInformation, gstrSysName
            tbsBill.Tabs(p - 1).Selected = True
        Else
            MsgBox "单据中的项目不存在执行科室为本机允许收费的执行科室！", vbInformation, gstrSysName
        End If
        Bill.SetFocus: Exit Function
    End If
    CheckExecuteDeptCanDo = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    If InStr("01245", mbytInState) > 0 Then
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
    ',冉俊明
    If (mbytInState = EM_ED_收费 And chkCancel.Value = 0) Then vsBalance.Editable = flexEDKbdMouse
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
    sta.Panels(Pan.C4预交信息).Visible = blnShow
    
    If Not blnShow Then
        sta.Panels(Pan.C4预交信息).Text = ""
    End If
End Sub

Public Function GetMedicareSum(colBalance As Collection, Optional ByVal strItem As String, Optional ByVal intPage As Integer, _
    Optional ByVal blnOrig As Boolean, Optional ByVal intBeforePage As Integer) As Currency
    '功能：获取保险结算的金额
    '参数：strItem=是否指定结算方式,否则为所有结算方式
    '      blnOrig=是否取原始(最大)结算金额,否则取现在(修改后)有效金额
    '      intPage=是否指定单据,否则为所有单据
    '      intBeforePage=计算该单据及以前的单据
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    Dim arrValue As Variant, curMoney As Currency
    Dim i As Integer, p As Integer
    
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, IIf(intBeforePage = 0, colBalance.Count, intBeforePage), intPage)
        For i = 0 To UBound(colBalance(p))
            '结算方式;原始(最大)金额;可否修改;有效金额
            arrValue = Split(colBalance(p)(i), ";")
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
    Next
    GetExecDepts = Mid(strTmp, 2)
End Function
Private Function GetInvoiceCount() As Integer
    '功能：计算当前收费需要打印多少张票据
    '说明：共有三级结构
    '   多张单据分别打印--按执行科室分别打印--按收费细目或收据费目打印
                    
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
        If mobjBill.Pages.Count > 1 And IsSplitPrintByNO Then
            GetInvoiceCount = mobjBill.Pages.Count
        Else
            GetInvoiceCount = 1
        End If
        Exit Function
    End If
    
    
    If mobjBill.Pages.Count > 1 And IsSplitPrintByNO Then
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
                    strSQL = "Select Count(" & IIf(gTy_Module_Para.byt票据生成方式 = 10, "Distinct 收据费目", "ID") & ") AS num From 门诊费用记录" & _
                        " Where 记录性质=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And NO=[1]" & _
                        " Group by 执行部门id"
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
                    Next
                    If gTy_Module_Para.byt票据生成方式 = 0 Then
                        If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt门诊收据行次)
                        strItems = ""
                    Else
                        X = X + IntEx(k / gTy_Module_Para.byt门诊收据行次)
                        k = 0
                    End If
                Else
                    strSQL = "Select Count(" & IIf(gTy_Module_Para.byt票据生成方式 = 0, "Distinct 收据费目", "ID") & ") AS num From 门诊费用记录" & _
                        " Where 记录性质=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And NO=[1]"
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
            str执行部门IDs = GetExecDepts()   '所有单据的执行部门
            
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
                                            If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                        End If
                                    Next
                                Else    '数量为零的行在保存前已检查禁止继续
                                    strTmp = mobjBill.Pages(i).Details(j).收费细目ID
                                    If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                End If
                            End If
                        Next
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
                                        If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                    End If
                                Next
                            Else    '数量为零的行在保存前已检查禁止继续
                                strTmp = mobjBill.Pages(i).Details(j).收费细目ID
                                If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                            End If
                        End If
                    Next
                Else
                    strNos = strNos & ",'" & mobjBill.Pages(i).NO & "'"
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

Private Sub ShowRegist()
'功能：检查是否可以显示挂号按钮
    Dim strPrivs As String
    On Error GoTo errH
    If mbytInState <> EM_ED_收费 Then Exit Sub
    strPrivs = GetPrivFunc(glngSys, 1111)
    '功能是否授权
    cmdRegist.Visible = zlStr.IsHavePrivs(strPrivs, "挂免费号") Or zlStr.IsHavePrivs(strPrivs, "挂收费号")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ShowIDCard()
'功能：检查是否可以显示就诊卡按钮
    Dim strPrivs As String
    On Error GoTo errH
    If mbytInState <> EM_ED_收费 Then Exit Sub
    strPrivs = GetPrivFunc(glngSys, 1107)
    cmdIDCard.Visible = zlStr.IsHavePrivs(strPrivs, "发卡")
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetOperatorInfo(ByVal str姓名 As String, Optional bln护士 As Boolean, Optional int职务 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定姓名开单人(医生或护士)的性质或职务
    '入参:int职务:0-未设置；bln护士:是否只是护士
    '编制:刘兴洪
    '日期:2014-06-09 14:35:52
    '说明：以前是直接读取marrDr中的内容,改为多单据多开单人后一些地方不行
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消医保病人身份验证
    '返回:返回假时不退出界面或清除操作
    '编制:刘兴洪
    '日期:2014-06-09 14:37:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    YBIdentifyCancel = True
    If mbytInState <> EM_ED_收费 Then Exit Function
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
End Function

Private Sub SetBillRowForeColor(ByVal lngRow As Long, ByVal lngColor As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置单据行颜色
    '编制:刘兴洪
    '日期:2014-06-09 14:39:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据药品/材料的储备限额设置行颜色提示
    '编制:刘兴洪
    '日期:2014-06-09 14:39:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjBill.Pages(intPage).Details.Count >= lngRow And mbytInState = EM_ED_收费 Then
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许保存为划价单
    '返回:允许保存划价单返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 17:52:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim p As Integer
    If Not (mbytInState = EM_ED_收费 And mstrInNO = "" And chkCancel.Value = 0) Then Exit Function
    
    For p = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(p).NO <> "" Then Exit Function
    Next
    CheckSaveMultiPrice = True  '允许保存为划价单
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
    
    If Not (mbytInState = EM_ED_收费) Then Exit Sub
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

Private Function LoadErrBillCharge(ByVal lng结帐ID As Long) As Boolean
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
    Dim lng病人ID As Long, lngRow As Long, strNo As String
    Dim blnMulitNos As Boolean
    
    If Not (mbytInState = EM_ED_异常重收 Or mbytInState = EM_ED_异常作废 Or mblnErrBill) Then LoadErrBillCharge = True: Exit Function
     
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "" & _
    "   Select A.NO, A.病人ID  " & _
    "   From 门诊费用记录 A" & _
    "   Where  A.结帐ID=[1]  " & _
    "   Group by A.NO,A.病人ID" & _
    "   Order by A.NO"
    
    Set rsNos = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    If rsNos.RecordCount = 0 Then Exit Function
    '检查是否存在未结的医保数据
    blnMulitNos = rsNos.RecordCount > 1
    
    mblnDelete = mbln作废异常
    '57682
    strSQL = "" & _
    "   Select decode(B.性质,NULL,-1,b.性质) as 序号,  decode(A.记录性质,1,'预存款',11,'预存款',A.结算方式) as 结算方式, " & _
    "          sum(nvl(A.冲预交,0)) as 结算金额 " & _
    "   From 病人预交记录 A,结算方式 B" & _
    "   where A.结帐ID=[1] And A.结算方式=B.名称(+) " & _
    "   Group by decode(B.性质,NULL,-1,b.性质),decode(A.记录性质,1,'预存款',11,'预存款',A.结算方式)" & _
    "   Order by 序号,结算方式"
    
    '异常单据的结算方式(不含预交款)
    Set mrsErrBlance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID)
    If mrsErrBlance.RecordCount = 0 Then Exit Function
    
    LoadErrBillCharge = True
    
    '清除现有单据的内容
    '---------------------------------------------------------------------
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
    
    mintInsure = zlGetBillChargeExistInsure(lng结帐ID, lng病人ID)
    If mintInsure <> 0 Then Call initInsurePara(lng病人ID)
    
    Do While Not rsNos.EOF
        
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
            '多张单据时禁止退费功能
            chkCancel.Enabled = False
            cmdDelete.Enabled = False
            '激活Click,显示新增加单据的内容(空白)
            tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
        End If
                
        '读取划价单据内容(同cboNO_KeyPress)
        '----------------------------------------------------------------------
        strNo = Nvl(rsNos!NO)
        blnRead = ReadBill(strNo, 0, , , True)
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
                If vsBalance.TextMatrix(i, 0) = Nvl(!结算方式, "未结金额") Then
                    blnFind = True
                    lngRow = i: Exit For
                End If
            Next
            If Not blnFind And vsBalance.TextMatrix(lngRow, 0) <> "" Then
                vsBalance.Rows = vsBalance.Rows + 1
                lngRow = vsBalance.Rows - 1
            End If
            vsBalance.TextMatrix(lngRow, 0) = Nvl(!结算方式, "未结金额")
            vsBalance.TextMatrix(lngRow, 1) = Format(Val(Nvl(!结算金额)) + Val(vsBalance.TextMatrix(lngRow, 1)), "0.00")
            If vsBalance.TextMatrix(lngRow, 0) = "未结金额" Then
                vsBalance.Cell(flexcpForeColor, lngRow, 0, lngRow, vsBalance.COLS - 1) = vbRed
                vsBalance.Cell(flexcpFontBold, lngRow, 0, lngRow, vsBalance.COLS - 1) = True
            Else
                vsBalance.Cell(flexcpForeColor, lngRow, 0, lngRow, vsBalance.COLS - 1) = Bill.ForeColor
                vsBalance.Cell(flexcpFontBold, lngRow, 0, lngRow, vsBalance.COLS - 1) = False
            End If
            .MoveNext
        Loop
    End With
    
    txtInvoice.Text = ""
    Call ReInitPatiInvoice(True, mintInsure, lng病人ID)
    Bill.Active = False
    chk加班.Enabled = False
    
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

Private Sub PrintBill(ByVal strNos As String, strModiNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印
    '编制:刘兴洪
    '日期:2011-08-26 18:38:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    Dim strReclaimInvoice As String '回收的发票号
    Dim int收费执行单 As Integer
    If InStr(1, strNos, "'") = 0 Then
        strNos = Replace(strNos, " ", "")
        strNos = Replace("'" & strNos & "'", ",", "','")
    End If
    
    If mblnSaveAsPrice Then   '打印划价通知单
        If gint划价通知单 = 1 Then
           Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & mobjBill.NO, 2)
        ElseIf gint划价通知单 = 2 Then
            If MsgBox("要打印划价通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & mobjBill.NO, 2)
            End If
        End If
        Exit Sub
    End If
     
     
    If mstrYBPati <> "" And MCPAR.门诊连续收费 Then
        '医保连续收费模式时，确定时不打印，等同一病人的几张单据确定完后，按[完成收费]按钮一起打印。
        '医保连续收费时不支持多单据,取一个就行了
        mstrYBBill = mstrYBBill & "," & mobjBill.NO
        Exit Sub
    End If
    
   '打印门诊收据
    '问题:34941
    Dim blnPrintBillEmpty As Boolean   '55052
    If mblnPrint And Not (MCPAR.医保接口打印票据 And mstrYBPati <> "") Then
        '问题:42708
        If Format(mobjBill.登记时间, "yyyy") < 2000 Then mobjBill.登记时间 = zlDatabase.Currentdate
        '问题:44322
RePrint:
        strReclaimInvoice = ""
        Call frmPrint.ReportPrint(1, strNos, strModiNos, strReclaimInvoice, mlng领用ID, mlngShareUseID, txtInvoice.Text, mobjBill.登记时间, CStr(mdbl缴款), CStr(mdbl找补), _
            IsSplitPrintByNO, mintInvoiceFormat, , , mstrUseType, blnPrintBillEmpty, , , mstr普通价格等级)
        If gblnStrictCtrl And blnPrintBillEmpty = False Then
            If zlIsNotSucceedPrintBill(1, strNos, strNotValiedNos) = True Then
                    If MsgBox("单据[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
            End If
        End If
    End If
    '打印费用清单:固定不分别打印
    If zlStr.IsHavePrivs(mstrPrivs, "打印清单") Then
        If gint收费清单 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
        ElseIf gint收费清单 = 2 Then
            If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            End If
        End If
    End If
    '62982:李南春,2015/5/19,收费执行单
    int收费执行单 = Val(zlDatabase.GetPara("收费执行单打印方式", glngSys, mlngModul))
    If zlStr.IsHavePrivs(mstrPrivs, "收费执行单") Then
        If int收费执行单 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), 2)
        ElseIf int收费执行单 = 2 Then
            If MsgBox("要打印收费执行单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), 2)
            End If
        End If
    End If
End Sub

Private Function PatiErrBillPay(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人,对异常单据进行收费
    '入参:lng病人ID-指定的病人ID
    '返回:存在异常单据,并进行重新收费或重新退费或重新作废,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-29 14:43:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNo As String, lng结帐ID As Long, lng结算序号 As Long
    Dim str操作员姓名 As String, strTittle As String
    Dim blnDel As Boolean, blnErrCancel As Boolean '异常作废
    Dim strDelTime As String, blnReplenishDel As Boolean
    Dim strPrivsReplenish As String, blnDoElsePersonErr As Boolean
    
    mblnErrBill = False
    mblnElsePersonErrBill = False
    If (mbytInState = EM_ED_浏览 Or mbytInState = EM_ED_调整) Then Exit Function
    If mbytInState = EM_ED_收费 And mstrInNO <> "" Then PatiErrBillPay = False: Exit Function
   
    On Error GoTo errHandle
    strSQL = " " & _
    "    Select  a.No, a.结帐id, a.操作员姓名, 1 As 异常类型,A.登记时间, a.记录状态 " & _
    "    From 门诊费用记录 A" & _
    "    Where nvl(费用状态,0) = 1 And 记录性质 = 1 And 病人id =[1] And 记录状态 = 1  " & _
    "          And Not Exists (Select 1 From 门诊费用记录 B Where a.No = b.No And Mod(b.记录性质, 10) = 1 And b.记录状态 = 2)" & _
    "    Union All " & _
    "    Select a.No, a.结帐id, a.操作员姓名, 2 As 异常类型,A.登记时间, a.记录状态 " & _
    "    From 门诊费用记录 A " & _
    "    Where nvl(费用状态,0) = 1 And 记录性质 = 1 And 病人id = [1] And 记录状态 = 2  " & _
    "          And Not Exists (Select 1 From 病人预交记录 B Where a.结帐id = b.结帐id And Nvl(b.校对标志, 0) = 0)"
    
    '异常单据处理顺序：操作员自己的单据优先，其次是收费异常单据优先
    strSQL = "" & _
    " Select distinct A.NO,A.结帐ID,A.操作员姓名,A.异常类型,A.登记时间,B.结算序号,a.记录状态" & _
    " From (" & strSQL & ") A,病人预交记录 B " & _
    " Where a.结帐ID=B.结帐ID(+)" & _
    " Order By Decode(a.操作员姓名,[2],0,1),a.记录状态"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, UserInfo.姓名)
    If rsTemp.EOF Then Exit Function
    
    strNo = Nvl(rsTemp!NO): lng结帐ID = Val(Nvl(rsTemp!结帐ID))
    blnDel = Val(Nvl(rsTemp!异常类型)) = 2
    strTittle = IIf(Not blnDel, "收费", "退费")
    lng结算序号 = Val(Nvl(rsTemp!结算序号))
    strDelTime = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
    str操作员姓名 = Nvl(rsTemp!操作员姓名)
    
    If str操作员姓名 <> UserInfo.姓名 Then
        If blnDel = False Then
            '判断是否能够对他人的收费异常单据进行重收
            strSQL = "Select 结算序号" & vbNewLine & _
                    " From 病人预交记录 A, 结算方式 B" & vbNewLine & _
                    " Where Nvl(a.结算方式, '-') = b.名称 And b.性质 Not In ('3', '4') And a.结帐id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
            If rsTemp.EOF Then
                '107905，具有“重收他人异常单据”权限时，可以对只进行了医保结算的他人的异常收费单据进行重收
                blnDoElsePersonErr = zlStr.IsHavePrivs(mstrPrivs, "重收他人异常单据")
            Else
                '存在其他非医保结算方式，其它操作员就不能处理了
                blnDoElsePersonErr = False
            End If
        End If
        
        If blnDoElsePersonErr = False Then
            If MsgBox("注意:" & vbCrLf & _
                "       该病人存在异常的" & strTittle & "单据，操作员[" & str操作员姓名 & "]收取了一部分，" & _
                "注意到操作员[" & str操作员姓名 & "]处对异常单据进行" & strTittle & "！" & vbCrLf & vbCrLf & _
                "       是否继续对该病人进行收费？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                PatiErrBillPay = True
            End If
            Exit Function
        End If
    End If
    
    '检查单据是否为补充结算单据，若为补充结算，只会是退费异常
    blnReplenishDel = CheckBillExistReplenishData(0, lng结算序号)
    If Not blnReplenishDel Then
        If MsgBox("注意:" & vbCrLf & _
                "       该病人存在异常的" & strTittle & "单据" & IIf(str操作员姓名 <> UserInfo.姓名, _
                ",该单据是操作员[" & str操作员姓名 & "]收取的", "") & _
                " ,是否重新对该单据进行" & strTittle & "?" & vbCrLf & vbCrLf & _
                "『是』代表重新对异常单据 " & strTittle & vbCrLf & _
                "『否』代表不对异常单据进行处理,继续进行收费操作.", _
                vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    If blnDel Then
       blnErrCancel = zlIsErrChargeCancel(strNo)
       If lng结算序号 < 0 Then
            If Not blnErrCancel Then
                '调用异常单据退
                '83271
                If blnReplenishDel Then
                    strPrivsReplenish = ";" & GetPrivFunc(glngSys, 1124) & ";"
                    If InStr(strPrivsReplenish, ";结算退费;") > 0 Then
                        If MsgBox("注意:" & vbCrLf & _
                                "       该病人存在异常的【保险补充结算】" & strTittle & "单据" & _
                                IIf(str操作员姓名 <> UserInfo.姓名, ",该单据是操作员[" & str操作员姓名 & "]收取的", "") & _
                                " ,是否重新对该单据进行" & strTittle & "?" & vbCrLf & vbCrLf & _
                                "『是』代表重新对异常单据 " & strTittle & vbCrLf & _
                                "『否』代表不对异常单据进行处理,继续进行收费操作.", _
                                vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                            PatiErrBillPay = frmReplenishTheBalanceDel.zlShowMe(Me, 1124, strPrivsReplenish, _
                                EM_RBDTY_异常重退, lng结算序号, False, 0, False, strDelTime) = False
                        End If
                    Else
                        If MsgBox("注意:" & vbCrLf & _
                                "       该病人存在异常的【保险补充结算】" & strTittle & "单据" & _
                                IIf(str操作员姓名 <> UserInfo.姓名, ",该单据是操作员[" & str操作员姓名 & "]收取的", "") & _
                                " ，你不具备操作该异常记录的权限！" & vbCrLf & vbCrLf & _
                                "       是否继续对该病人进行收费？", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PatiErrBillPay = True
                        End If
                    End If
                    Exit Function
                End If
                PatiErrBillPay = frmClinicDelAndView.ShowMe(Me, EM_MULTI_异常重退, mstrPrivs, lng结算序号, _
                    False, 0, False, strDelTime) = False
                Exit Function
            End If
            '重新对异常收费作废单据进行作废
            mbytInState = EM_ED_异常作废
            mlng结帐ID = lng结帐ID
            mlng结算序号 = lng结算序号
            mbln作废异常 = True
            mblnErrBill = True
            If LoadBill() = False Then Exit Function
            
            PatiErrBillPay = True
            Call cmdOK_Click
            If Not gfrmMain Is Nothing Then
                mlng结帐ID = 0: mbytInState = EM_ED_收费
            End If
            Exit Function
       Else
            If Not blnErrCancel Then
                PatiErrBillPay = frmMultiBills.ShowMe(gfrmMain, 2, mstrPrivs, strNo, strDelTime, , , False)
                Exit Function
            End If
            '重新对异常作废的进行重新作废
            frmCharge.mlngModul = mlngModul
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 5
            frmCharge.mstrInNO = strNo
            frmCharge.mbln退费异常 = True
            Set frmCharge.mobjMsgModule = mobjMsgModule
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
            PatiErrBillPay = gblnOK
            Exit Function
        End If
    End If
    If lng结算序号 >= 0 Then
        '针对34.0以前的版本,重新对异常作废的进行重新作废
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInState = 4
        frmCharge.mstrInNO = strNo
        frmCharge.mbln退费异常 = True
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        PatiErrBillPay = gblnOK
    End If
    '重新对异常单据进行重收费
    mbytInState = EM_ED_异常重收
    mlng结帐ID = lng结帐ID
    mlng结算序号 = lng结算序号
    mblnErrBill = True
    If LoadBill() = False Then Exit Function
    
    mblnElsePersonErrBill = blnDoElsePersonErr
    PatiErrBillPay = True
    Call cmdOK_Click
    If Not gfrmMain Is Nothing Then
        mlng结帐ID = 0: mbytInState = EM_ED_收费
    End If
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
        '性质:-99-缴款;-98-找补,0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        grsTotal.Sort = "性质"
        .Rows = IIf(.Rows >= grsTotal.RecordCount, .Rows, grsTotal.RecordCount)
        lngRow = 0
        Do While Not grsTotal.EOF
            '性质 ,结算方式  结算金额
            '从frmClinicChargePayMentWin-传入,主要是一些累计数
            .TextMatrix(lngRow, 0) = Nvl(grsTotal!结算方式)
            .TextMatrix(lngRow, 1) = Format(Val(Nvl(grsTotal!结算金额)), "###0.00;-###0.00;0.00;0.00")
             int性质 = Val(Nvl(grsTotal!性质))
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .COLS - 1) = Me.ForeColor
            If int性质 = -99 Then
                .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True
            ElseIf int性质 = -98 Then
                .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .COLS - 1) = vbRed
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
    If Not (mbytInState = EM_ED_收费) Then Exit Sub
    If mblnNotClearLedDisplay Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub

Private Function SaveChargeBill(ByRef lng结帐ID As Long, _
      ByRef cllSavePriceSQL As Collection, ByRef cllSaveSQL As Collection, _
      ByRef cllChargeOverAfterPro As Collection, _
      Optional ByRef strSaveNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前输入的单据(适用于收费 )
    '出参:lng结帐ID-返回本次保存单据的结算ID
    '     cllSaveSQL-保存的单据SQL,该集合的元素为集合,Key值为单据号
    '     cllChargeOverAfterPro-完成收费后,执行的其他过程(主要是发料和发药),该集合的元素为集合,Key值为单据号
    '     strSaveNos-返回的单据号
    '返回:收费成功或单据保存存功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-26 17:28:24
    '说明:
    '     *** 医保收费时,先临时保存为划价单,在结算前再转为收费单,以避免更新药品库存时因等待同一事务的医保结算操作而锁表 ***
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, p As Integer
    Dim bln直接收费 As Boolean, arrSQL As Variant, strSQL As String
    Dim strBillNO As String, strInvoice As String '当前单据使用的票据号，用于医保一张单据只打一张票的情况
    Dim str医疗付款 As String, str结算方式 As String
    Dim int序号 As Integer, int价格父号 As Integer, int行号 As Integer
    Dim strDeptIDs As String, strStuffDept As String '自动发药和发料的部门
    Dim dbl数次 As Double, dbl单价 As Double, lng医嘱ID As Long
    Dim str中药形态 As String
    Dim varTemp As Variant, strTmp As String
    
    Dim cllFeeBillItem As Collection, cllChargeOverItem As Collection
    Dim lng打印ID As Long
    Dim cllPriceBillItem As Collection
    
    '只处理收费单
    If mblnSaveAsPrice Then Exit Function
    
    '新的发药窗品集(目前只针手工录入有效)
    Set mCllWindows = New Collection
    
    Err = 0: On Error GoTo errHander
    If cbo医疗付款.ListIndex <> -1 Then
        str医疗付款 = Mid(cbo医疗付款.Text, 1, InStr(1, cbo医疗付款, "-") - 1)
    End If
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    strInvoice = Trim(txtInvoice.Text)
    str结算方式 = GetMedicareStr(mcolBalance) '预结算结果
    
    strSaveNos = ""
    Set cllSavePriceSQL = New Collection
    Set cllSaveSQL = New Collection
    Set cllChargeOverAfterPro = New Collection
    '=================================================================================
    '处理规则：
    '1.对直接收费单据先保存为划价单先提交以便不锁表(药品库存)，再对划价单收费
    '2.单据号作为集合的Key值
    'cllSavePriceSQL - 划价单SQL集合
    'cllSaveSQL - 划价单收费SQL集合
    '=================================================================================
    
    '对每张单据独立执行保存
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    lng打印ID = zlDatabase.GetNextId("票据打印内容")
    
    For p = 1 To mobjBill.Pages.Count
        int序号 = 0: int行号 = 0
        strDeptIDs = "": strStuffDept = ""
        Set cllPriceBillItem = New Collection
        Set cllFeeBillItem = New Collection
        Set cllChargeOverItem = New Collection
        
        '产生每张收费单据的单据号
        bln直接收费 = False
        strBillNO = mobjBill.Pages(p).NO
        If mobjBill.Pages(p).NO = "" Then
            '为保存失败后仍能识别,不改对象NO
            strBillNO = zlDatabase.GetNextNo(13)    '收费单
            bln直接收费 = True
        End If
        
        '主要为消息发送用,为每页保存的单据号
        mobjBill.Pages(p).收费单号 = strBillNO
        If p = 1 Then mobjBill.NO = strBillNO
        
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
            gstrSQL = gstrSQL & "" & "NULL" & ")"
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
            '2.直接输入的单据内容
            '---------------------------------------------------------------
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                If mobjBillDetail.数次 <> 0 Then
                    For Each mobjBillIncome In mobjBillDetail.InComes
                        int序号 = int序号 + 1 '当前记录序号
                        '1.单据主体---------------------------------------------------------------
                        With mobjBill                              '先临时保存为划价单,在结算前再转为收费单
                            gstrSQL = "zl_门诊划价记录_INSERT('" & strBillNO & "'," & int序号 & "," & ZVal(.病人ID) & "," & _
                                ZVal(.主页ID) & "," & ZVal(.标识号) & ",'" & str医疗付款 & "'," & _
                                "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                                .加班标志 & "," & ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
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
                            
                            '直接收费时,因为先暂存为划价单,收费时需要取发药窗口
                            tbsBill.Tabs(p).Tag = .发药窗口
                            
                            dbl数次 = .数次
                            If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                                dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                            End If
                            
                            gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                            gstrSQL = gstrSQL & "'" & .发药窗口 & "'," & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & ","
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
                        End With
        
                        '4.其它部分
                        '---------------------------------------------------------------
                        gstrSQL = gstrSQL & _
                                "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        gstrSQL = gstrSQL & "'" & mstrInNO & "',"
                        
                        If mobjBillDetail.收费类别 = "7" Then
                            str中药形态 = "'" & mobjBillDetail.Detail.中药形态 & "'"
                        Else
                            str中药形态 = "NULL"
                        End If
                        '中药形态_In       门诊费用记录.结论%Type := Null
                        
                        '门诊划价,收费功能划价
                        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                            "'" & mobjBillDetail.摘要 & "'," & ZVal(lng医嘱ID) & ",NULL,NULL,'|" & mobjBill.Pages(mintPage).煎法 & _
                            "',NULL,NULL," & gint病人来源 & ",'" & mobjBillDetail.保险编码 & "'," & _
                            "'" & mobjBillDetail.Detail.类型 & "'," & IIf(mobjBillDetail.保险项目否, 1, 0) & "," & ZVal(mobjBillDetail.保险大类ID) & "," & _
                            str中药形态 & ")"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
                    Next    '每一条收入项目
                    
                    
                    '对每一行收费记录收集药品执行部门
                    '----------------------------------------------------------------------------------------------------------------
                    '自动发药                   '
                    With mobjBillDetail
                        If gbln收费后自动发药 Then
                            If .执行部门ID <> 0 And InStr("5,6,7", .收费类别) > 0 Then
                                If InStr(strDeptIDs & ",", "," & .执行部门ID & ",") = 0 Then
                                    strDeptIDs = strDeptIDs & "," & .执行部门ID
                                End If
                            End If
                        End If
                        '自动发料
                        If gbln门诊自动发料 Then
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
        End If
        
        '收费后自动发药, 收费且不是保存为划价单
        '-----------------------------------------------------------------------
        If strDeptIDs <> "" Then
            strDeptIDs = Mid(strDeptIDs, 2)
            varTemp = Split(strDeptIDs, ",")
            For i = 0 To UBound(varTemp)
                gstrSQL = "ZL_药品收发记录_处方发药(" & Val(varTemp(i)) & ",8,'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & mobjBill.Pages(p).开单人 & "')"
                zlAddArray cllChargeOverItem, gstrSQL
            Next
        End If
        
        '收费后自动发料,在收费(直接收费,划价单导入收费),门诊记帐时执行
        '-----------------------------------------------------------------------
        If strStuffDept <> "" Then
            strStuffDept = Mid(strStuffDept, 2)
            varTemp = Split(strStuffDept, ",")
            For i = 0 To UBound(varTemp)          '24-收费处方发料；25-记帐单处方发料
               gstrSQL = "zl_材料收发记录_处方发料(" & varTemp(i) & "," & 24 & ",'" & strBillNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                zlAddArray cllChargeOverItem, gstrSQL
            Next
        End If
        
        '执行相关SQL语句及提交医保结算
        '--------------------------------------------------------------------------------------------------------------------------------
        '对SQL序列按收费细目ID排序
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j)): arrSQL(j) = arrSQL(i): arrSQL(i) = strTmp
                End If
            Next
        Next
        
        '直接收费时,先保存为划价单,再转为收费单
        '-------------------------------------------------------------------
        If bln直接收费 Then
            '1.先保存划价单,先提交库存更新以便不锁表
            For i = 0 To UBound(arrSQL)
                zlAddArray cllPriceBillItem, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
            Next
            '更新划价单的保险信息(保险项目否,医保大类ID,统筹金额)
            gstrSQL = "zl_门诊划价记录_Update(" & mintInsure & "," & mobjBill.病人ID & ",'" & strBillNO & "',0)"
            zlAddArray cllFeeBillItem, gstrSQL
            
            '划价单转为收费单
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
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.科室ID, , mobjBill.Pages(p).开单部门ID) & ","
            '  开单部门id_In 门诊费用记录.开单部门id%Type,
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).开单部门ID) & ","
            '  开单人_In     门诊费用记录.开单人%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).开单人 & "',"
            '  结帐id_In     门诊费用记录.结帐id%Type,
            gstrSQL = gstrSQL & "" & lng结帐ID & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            gstrSQL = gstrSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
            '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
            'gstrSQL = gstrSQL & "'" & tbsBill.Tabs(p).Tag & "',"
            gstrSQL = gstrSQL & "NULL," '前面划价单阶段已经保存，这里不用修改，冉俊明，2015-1-20
            '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
            gstrSQL = gstrSQL & "" & chk急诊.Value & ","
            '  登记时间_In   门诊费用记录.登记时间%Type := Null,
            gstrSQL = gstrSQL & "" & "NULL" & ")"
            zlAddArray cllFeeBillItem, gstrSQL
        Else
            For i = 0 To UBound(arrSQL)
                zlAddArray cllFeeBillItem, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
            Next
        End If
        
        '收费完成后的处理
        '-----------------------------------------------------
        '先填写开始票据号以便医保调用时上传,多张分别打印时,填写相同的,打印调用时将重写,取消打印或打印失败将清除
        If strInvoice <> "" And mblnPrint Then
            gstrSQL = "Zl_票据起始号_Update('" & strBillNO & "','" & strInvoice & "',1)"
            zlAddArray cllFeeBillItem, gstrSQL
        End If
        
        '81579,冉俊明,2015-1-9,医保接口打印票据时,在票据使用明细中无记录,导致打印不出来内容
        If mintInsure <> 0 And _
            MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
            '38821
            '票据数据生成(因为不调HIS的打印，医保接口打印，所以先填票据数据)
            '只有第一张单据时需要插入票据使用记录，后面的单据只需要插入票据打印内容（打印ID相同）
            gstrSQL = "zl_门诊收费票据_Insert('" & strBillNO & "','" & strInvoice & "'," & ZVal(mlng领用ID) & "," & _
                "'" & UserInfo.姓名 & "',To_Date('" & mobjBill.登记时间 & "','YYYY-MM-DD HH24:MI:SS')," & lng打印ID & ",1,0,NULL," & IIf(p = 1, "1", "0") & ")"
            zlAddArray cllFeeBillItem, gstrSQL
        End If
        
        '预结算结果，“只对医保结算成功的单据收费”时第一张单据提交时就保存到病人预交记录中，结算完成后再进行校对
        '其它情况，都放在最后一张单据
        If p = IIf(mintInsure <> 0 And MCPAR.多单据分单据结算 And gTy_Module_Para.bln只对医保结算成功单据收费, 1, mobjBill.Pages.Count) Then
           'Zl_门诊收费结算_Modify
            gstrSQL = "Zl_门诊收费结算_Modify("
            '  操作类型_In   Number,
            gstrSQL = gstrSQL & "" & 2 & ","
            '  病人id_In     门诊费用记录.病人id%Type,
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.病人ID) & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            gstrSQL = gstrSQL & "" & lng结帐ID & ","
            '  结算方式_In   Varchar2,
            gstrSQL = gstrSQL & IIf(str结算方式 = "", "NULL", "'" & str结算方式 & "'") & ")"
            '  退预交_In     病人预交记录.冲预交%Type := Null,
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            '  卡号_In       病人预交记录.卡号%Type := Null,
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            '  缴款_In       病人预交记录.缴款%Type := Null,
            '  找补_In       病人预交记录.找补%Type := Null,
            '  误差金额_In   门诊费用记录.实收金额%Type := Null,
            '  完成退费_In Number:=0
             zlAddArray cllFeeBillItem, gstrSQL
        End If
        
        If bln直接收费 Then cllSavePriceSQL.Add cllPriceBillItem, strBillNO '以单据号作为集合的Key值
        cllSaveSQL.Add cllFeeBillItem, strBillNO '以单据号作为集合的Key值
        cllChargeOverAfterPro.Add cllChargeOverItem, strBillNO
        strSaveNos = strSaveNos & "," & strBillNO
        
        '加入单据历史记录(所有类型单据)
        cboNO.AddItem strBillNO, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i '只显示10个
        Next
    Next  '下一张单据
    If strSaveNos = "" Then Exit Function
    
    strSaveNos = Mid(strSaveNos, 2)
    
    SaveChargeBill = True
    Exit Function
errHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetMoneyToTal(Optional ByVal intBeforePage As Integer) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次缴款总额
    '入参:
    '      intBeforePage=计算该单据及以前的单据
    '编制:刘兴洪
    '日期:2012-02-17 15:25:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim dblToTal As Double, intCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    For i = 1 To IIf(intBeforePage = 0, mobjBill.Pages.Count, intBeforePage)
        If mobjBill.Pages(i).Details.Count > 0 Then
            dblToTal = dblToTal + mobjBill.Pages(i).误差金额
            For j = 1 To mobjBill.Pages(i).Details.Count
                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                    dblToTal = dblToTal + mobjBill.Pages(i).Details(j).InComes(k).实收金额
                Next
            Next
        Else    '提取划价单收费时没有明细费用
            dblToTal = dblToTal + mobjBill.Pages(i).误差金额
            dblToTal = dblToTal + mobjBill.Pages(i).实收金额
        End If
    Next
    dblToTal = RoundEx(dblToTal, 6)
    
    '如果没有,再尝试从表格中取(仅一张单据时)
    If dblToTal = 0 And tbsBill.Tabs.Count = 1 _
        And Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.项目) = "") Then
        intCol = BillCol.实收金额
        For i = 1 To Bill.Rows - 1
            If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                dblToTal = dblToTal + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
            End If
        Next
    End If
    GetMoneyToTal = Format(dblToTal, gstrDec)
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
    If Not mbln连续输入 Or mbytInState <> EM_ED_收费 Then Exit Sub
    
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
    bln检查库存 = (zlStr.IsHavePrivs(mstrPrivs, "不检查库存") = False)    '是否有权限不检查库存(分批和时价必须检查)
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
            
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        
                        If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & _
                                dblToTal & """。", vbInformation, gstrSysName
                            'tbsBill.Tabs(p).Selected = True
                            Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            If gbln药房单位 Then .Detail.库存 = .Detail.库存 / .Detail.药房包装
                            
                            If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行药品""" & .Detail.名称 & _
                                    """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & _
                                    dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                                'tbsBill.Tabs(p).Selected = True
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        
                        If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblToTal > .Detail.库存 Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                            MsgBox strTmp & "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & dblToTal & """。", vbInformation, gstrSysName
                            'tbsBill.Tabs(p).Selected = True:
                            Exit Function
                        End If
                    Else
                        If colStock("_" & .执行部门ID) = 2 And bln检查库存 Then
                            dblToTal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                            .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                            
                            If mbytInState = EM_ED_收费 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                            If dblToTal > .Detail.库存 Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                                MsgBox strTmp & "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                    """的当前库存" & IIf(zlStr.IsHavePrivs(mstrPrivs, "显示库存"), .Detail.库存, "") & "不足输入数量""" & dblToTal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
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
        If HaveExecute(1, mstrInNO, 1) Then
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
    
    
    If Not (mbytInState = EM_ED_收费 Or mbytInState = 5) Then Exit Sub
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
                    If mblnSaveAsPrice Or mstrYBPati <> "" Then
                        '门诊划价(收费)
                        Call objTemp.appendData("bill_kind", 1)
                        Call objTemp.appendData("charge_state", 1)
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
            "   Where NO=[1] And mod(记录性质,10)=1 And  记录状态=1 " & _
            "   Order by 收费类别"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
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
                        Call objTemp.appendData("bill_kind", 1)  '1-收费单;2-记帐单
                        Call objTemp.appendData("charge_state", 2)   '1-未收费;2-已收费
                      
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
    Case EM_ED_收费, EM_ED_异常重收, EM_ED_异常作废
        MsgBox "系统中尚未设置有效的误差处理,请在[结算方式管理]中设置。", vbInformation, gstrSysName
        Exit Function
    Case Else
        IsCheck误差费 = True
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitLed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Led
    '编制:刘兴洪
    '日期:2014-06-05 15:27:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If gblnLED = False Then Exit Sub
    If Not (mbytInState = EM_ED_收费 Or mbytInState = EM_ED_异常重收 Or mbytInState = EM_ED_异常作废) Then Exit Sub
    zl9LedVoice.Reset com
    zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModul, gcnOracle
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

    If Not (mbytInState = EM_ED_收费 Or mbytInState = EM_ED_调整 Or mbytInState = EM_ED_异常重收) Then Exit Sub

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
Private Function GetChargeInfor(ByRef objCharge As clsClinicChargeInfor, _
    Optional ByVal intBeforePage As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算相关信息
    '入参:
    '      intBeforePage=计算该单据及以前的单据
    '出参:objCharge-获取结算信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 14:49:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If objCharge Is Nothing Then
        Set objCharge = New clsClinicChargeInfor
    End If
    
    With objCharge
        .intInsure = mintInsure
        .PatiUseType = mstrUseType
        .ShareUserID = mlngShareUseID
        .病人ID = mobjBill.病人ID
        .姓名 = mobjBill.姓名
        .性别 = mobjBill.性别
        .年龄 = mobjBill.年龄
        .费别 = mobjBill.费别
        .缴款 = mdbl缴款
        .实收金额 = GetMoneyToTal(intBeforePage)
        .消费合计 = .实收金额
        .医保预结金额 = GetMedicareSum(mcolBalance, , , , intBeforePage)
        .医保结算金额 = .医保预结金额
        .预结结算 = GetMedicareStr(mcolBalance, , intBeforePage)
        .当前发票号 = Trim(txtInvoice.Text)
        .医保不走票号 = MCPAR.医保不走票号
        .缺省结算方式 = Get缺省结算方式(zlStr.NeedName(cbo医疗付款.Text))
        .费用来源 = GetFeeFromType()
    End With
    GetChargeInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function Get缺省结算方式(ByVal str医疗付款方式 As String) As String
    '根据医疗付款方式获取缺省的结算方式
    On Error GoTo errHandler
    If mrs缺省结算方式 Is Nothing Then
        Set mrs缺省结算方式 = Get结算方式("收费", "", True)
    End If
    mrs缺省结算方式.Filter = "付款方式='" & str医疗付款方式 & "'"
    If mrs缺省结算方式.EOF Then Exit Function
    Get缺省结算方式 = Nvl(mrs缺省结算方式!名称)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'--------------------------------------------------------------------------------------------------------
'相关接口
Public Function zlReCalcMoney(ByRef objChargeInfor As clsClinicChargeInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算费用
    '出参:objChargeInfor-重新返回结算信息
    '返回:重算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 14:42:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    If mstrYBPati = "" Or MCPAR.门诊预结算 = False Then zlReCalcMoney = True: Exit Function
    Call MsgBox("注意:" & vbCrLf & "  费用价格发生变化,需要重新进行医保预拟结算,请确认医保卡是否插入!", vbInformation + vbOKOnly, gstrSysName)
    Call cmdYB_Click
    zlReCalcMoney = GetChargeInfor(objChargeInfor)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetSaveBillSQL(ByRef lng结帐ID As Long, _
    ByRef cllSavePriceSQL As Collection, ByRef strSaveNos As String, _
    ByRef cllSavePro As Collection, _
    ByRef cllChargeOverAfterPro As Collection, _
    Optional blnSavePrice As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取可以保存单据的相关SQL
    '出参:lng结帐ID-返回结帐ID
    '    cllSavePro-保存的单据的相关过程集
    '    strSaveNos-返回要保存的单据号
    '    cllChargeOverAfterPro-收费完成后执行的过程
    '    blnSavePrice-是否保存为划价单(联合医保使用)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 15:12:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If CheckChargeDataValied = False Then Exit Function
    zlGetSaveBillSQL = SaveChargeBill(lng结帐ID, cllSavePriceSQL, cllSavePro, cllChargeOverAfterPro, strSaveNos)
    blnSavePrice = mblnSaveAsPrice And Not mnuFileSavePrice.Checked
End Function

Private Function SaveChargePriceBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费保存为划价单
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-16 10:22:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, bytReturnMode As String
    Dim blnSaveBill As Boolean, bln连续 As Boolean, blnGetFact As Boolean
    Dim cur个帐透支 As Currency
    Dim str划价Nos As String, rsItems As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If Not mblnSaveAsPrice Then Exit Function
    If isValiedCargeFee = False Then Exit Function
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    If zlGetSaveDataItems_Plugin(mobjBill, str划价Nos, rsItems) = False Then Exit Function
    If zlChargeSaveValied_Plugin(glngModul, 1, True, True, str划价Nos, rsItems) = False Then Exit Function
    '票据号及工本费及汇总金额相关检查
    If CheckBillNOAndBookeFee = False Then Exit Function
    If CheckInsure = False Then Exit Function
        
     
    cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保结算延时
    cmdCancel.Enabled = False: cmdAddBill.Enabled = False: cmdDelBill.Enabled = False
    
    If cmd预结算.Visible And cmd预结算.Enabled Then cmd预结算.Enabled = False
    '保存单据
    '---------------------------------------------------------------------------------------------
    strNos = "": bytReturnMode = 0
    If Not SaveClinicPriceBill(strNos, blnSaveBill, bln连续) Then
        '收费,保存单据失败后的处理
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        If mintInsure <> 0 Then
            cmdAddBill.Enabled = Not MCPAR.门诊连续收费 And _
                MCPAR.多单据收费 And zlStr.IsHavePrivs(mstrPrivs, "医保病人多单据收费")
        Else
            cmdAddBill.Enabled = zlStr.IsHavePrivs(mstrPrivs, "普通病人多单据收费")
        End If
        
        If cmdDelBill.Visible And tbsBill.Tabs.Count > 1 Then cmdDelBill.Enabled = True
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        If mblnAutoChangePati And gint病人来源 = 2 Then
            '需要切找到病人来源1中
            gint病人来源 = 1: zlChangePatiSource (gint病人来源)
        End If
        Call ClearFullBill(False)
         Exit Function
    End If
    Call zlChargeSaveAfter_Plugin(glngModul, mobjBill.病人ID, mobjBill.主页ID, True, 1, strNos)
    
    mlng结算序号 = 0
    Call SendMsgModule
     
     '打印票据
    Call PrintBill(strNos, "")
    cmdOK.Enabled = True   '防止设置打印机弹出的非模态窗体,以及医保延时
    cmdCancel.Enabled = True
    If cmd预结算.Visible Then cmd预结算.Enabled = True
    If mbytInState = EM_ED_收费 And gbln累计 Then
        txt累计.Text = Format(GetChargeTotal, "0.00")
    End If
        
    sta.Panels(Pan.C2提示信息) = "上一张单据:" & mobjBill.NO '多单据时为第一张
    mstrInNO = "":  mlngFirstID = 0: mstrFirstWin = ""
    If gint病人来源 = 2 And mblnAutoChangePati Then
    
        '自动切换的,要换回来
        gint病人来源 = 1
        Call zlChangePatiSource(gint病人来源)
    End If
    Call ClearPatientInfo(True)
    Call ClearTotalInfo(True)
    Call InitCommVariable
    blnGetFact = IIf(mblnStartFactUseType, False, True)
    Call ClearBillRows
    
    If mstrYBPati <> "" And MCPAR.门诊连续收费 Then
        Call NewYBBill
        mobjBill.病人ID = CLng(Split(mstrYBPati, ";")(8))
        '重新读取个帐余额
        cur个帐透支 = mdbl个帐透支
        mdbl个帐余额 = gclsInsure.SelfBalance(mobjBill.病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, cur个帐透支, mintInsure)
        mdbl个帐透支 = cur个帐透支
        sta.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mdbl个帐余额, "0.00")
        sta.Panels(Pan.C3个人帐户).Visible = True

        mstrYBPati = ""
    End If
    
    '提醒票据是否充ss足
    If Not mblnStartFactUseType Then Call zlCheckFactIsEnough
    
    If Not txtPatient.Locked Then
        txtPatient.SetFocus
    Else
        Bill.SetFocus
    End If
    mblnSaveData = True
    mlng结算序号 = 0
    SaveChargePriceBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    cmdOK.Enabled = True
    Call SaveErrLog
End Function

Private Function SaveClinicPriceBill(ByRef strSaveNos As String, _
    Optional ByRef blnSaveClinicPriceBill As Boolean, _
    Optional bln连续 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前输入的单据(适用于收费、划价、门诊记帐)
    '出参:strSaveNos-返回已成功保存的单据号，格式为"'AAA','BBB',..."
    '       cur已缴合计-配合strSaveNOs，返回已保存成功的单据实际已缴的现金
    '       blnSaveClinicPriceBill-是否单据已经保存成功
    '返回:收费成功或单据保存存功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-26 17:28:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str医疗付款 As String, strInvoice As String, strBillNO As String
    Dim str中药形态 As String, strTmp As String, strSaveCuessNos As String
    Dim arrSQL As Variant, arrPut As Variant, arrOTMSQL As Variant
    Dim cllPro As Collection, cllPageInfor As Collection, blnTransMedicare As Boolean
    Dim p As Integer, i As Long, j As Long, int药品行次 As Integer, bln直接收费 As Boolean
    Dim int序号 As Integer, int价格父号 As Integer, int行号 As Integer
    Dim dbl数次 As Double, dbl单价 As Double, blnTrans As Boolean
    
    Set mCllWindows = New Collection
    
    strSaveNos = ""
    Err = 0: On Error GoTo Errhand:
    If cbo医疗付款.ListIndex <> -1 Then
        str医疗付款 = Mid(cbo医疗付款.Text, 1, InStr(1, cbo医疗付款, "-") - 1)
    End If
    strInvoice = Trim(txtInvoice.Text)
    
    arrOTMSQL = Array()
    
    blnSaveClinicPriceBill = False
    Set cllPro = New Collection
    Set cllPageInfor = New Collection
    '对每张单据独立执行保存
    For p = 1 To mobjBill.Pages.Count
        int序号 = 0: int行号 = 0: int药品行次 = 0
        '产生每张收费单据的单据号
        If mobjBill.Pages(p).NO = "" Then
            '为保存失败后仍能识别,不改对象NO
            strBillNO = zlDatabase.GetNextNo(13)
            bln直接收费 = True
        Else
            bln直接收费 = False
            strBillNO = mobjBill.Pages(p).NO
        End If
        
        '主要为消息发送用,为每页保存的单据号
        mobjBill.Pages(p).收费单号 = strBillNO
        If p = 1 Then mobjBill.NO = strBillNO
        
        arrSQL = Array() '多单据时,逐张单据提交
        If Not bln直接收费 Then
            '1.收费新单据功能时,提取的划价单收费
            '提取划价单收费,但仍保存为划价单,或联合医保的保存
            If mstrYBPati <> "" And mobjBill.病人ID <> 0 Then
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
                            gstrSQL = "zl_门诊划价记录_INSERT('" & strBillNO & "'," & int序号 & "," & ZVal(.病人ID) & "," & _
                                ZVal(.主页ID) & "," & ZVal(.标识号) & ",'" & str医疗付款 & "'," & _
                                "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & IIf(mobjBillDetail.费别 = "", .费别, mobjBillDetail.费别) & "'," & _
                                .加班标志 & "," & ZVal(.科室ID, , .Pages(p).开单部门ID) & "," & _
                                ZVal(.Pages(p).开单部门ID) & ",'" & .Pages(p).开单人 & "',"
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
                            If InStr(",5,6,7,", .收费类别) > 0 Then
                                If Set发药窗口(p, mobjBillDetail) = False Then Exit Function
                            End If
                            '医保直接收费时,因为先暂存为划价单,收费时需要取发药窗口
                            dbl数次 = .数次
                            If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                                dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                            End If
                            
                            gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                            gstrSQL = gstrSQL & "'" & .发药窗口 & "'," & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & ","
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
                        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                            "'" & mobjBillDetail.摘要 & "',NULL,NULL,NULL,'|" & mobjBill.Pages(mintPage).煎法 & _
                            "',NULL,NULL," & gint病人来源 & ",'" & mobjBillDetail.保险编码 & "'," & _
                            "'" & mobjBillDetail.Detail.类型 & "'," & IIf(mobjBillDetail.保险项目否, 1, 0) & "," & ZVal(mobjBillDetail.保险大类ID) & "," & _
                            str中药形态 & ")"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
                    Next    '每一条收入项目
                End If
            Next            '每一行收费项目
            
            '保存前一张单据的药房ID,以便多张单据时确定发药窗口
            If mobjBill.Pages.Count > 1 Then Call SaveDrugID(p)
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

            '删除就诊卡划价单:多张单据时只删除一次(因为通过就诊卡号读病人时,就诊卡划价单已生成收费细目行,所以要删除)
            If mstrCardNO <> "" And strSaveNos = "" Then
                gstrSQL = "zl_门诊划价记录_Delete('" & mstrCardNO & "')"
                zlAddArray cllPro, gstrSQL
            End If
            '执行主体的SQL语句
            For i = 0 To UBound(arrSQL)
                zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
            Next
            
            cllPageInfor.Add Array(0, strBillNO), "K" & p
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
    blnTrans = True:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If zlInsureClinicSwapPrice(strSaveNos, strSaveCuessNos) = False Then
        If strSaveCuessNos <> "" Then blnSaveClinicPriceBill = True:
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnSaveClinicPriceBill = True: blnTrans = False
    SaveClinicPriceBill = True
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
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub FromBillNoReprintBill(ByVal strNo As String, ByVal blnNOMoved As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据重打票据(一般是针对一张单据一次结算的情况,即:10.34版本以前数据)
    '入参:lng病人ID-病人ID
    '     strNO-指定重打的单据
    '     blnNOMoved-是否转储到后备表
    '出参:
    '编制:刘兴洪
    '日期:2014-08-07 10:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, lng结帐ID As Long, lng病人ID As Long
    Dim intInsure As Integer, blnVirtualPrint As Boolean
    Dim strReclaimInvoice As String, intInvoiceFormat As Integer
    
    
    On Error GoTo errHandle
    
    strNos = zlGetBalanceNos(0, strNo, blnNOMoved)
    '单据有剩余数量的才可以重打
    If Not BillExistMoney(strNos, 1, True) Then
        MsgBox "单据不存在或已经全部退费,不能重打！", vbInformation, gstrSysName
        txtRePrint.Text = "": Exit Sub
    End If
    '调出重打的单据显示
    If frmMultiBills.ShowMe(Me, 0, mstrPrivs, strNo, "", True) = False Then Exit Sub
    intInsure = ChargeExistInsure(strNo, lng病人ID, lng结帐ID)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
        '此处只提供了收费票据的重打
    End If
    Call ReInitPatiInvoice(True, intInsure, lng病人ID)
    strReclaimInvoice = zlGetReclaimInvoice(strNo)
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
        Call RefreshFact '刷新票据号
        txtRePrint.Text = ""
        txtPatient.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlInsureClinicSwap(lng结帐ID As Long, Optional ByVal intInsure As Integer = 0, _
    Optional ByRef strAdvance As String = "", Optional ByRef blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保结算
    '参数：bytSucceed - 返回失败时，0:一张都未执行成功，1:部分成功
    '返回:医保结算成功或非医保,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur全自付 As Currency, cur先自付 As Currency
    Dim i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '记录各张单据保险结算
    Dim strSQL As String
    Dim str结算方式 As String, strNo As String
    Dim str预结算 As String
    Dim varAdvance As Variant, varItem As Variant
    Dim blnFind As Boolean
    
    On Error GoTo errHandle
    If intInsure = 0 Then zlInsureClinicSwap = True: Exit Function
    
    If MCPAR.多单据分单据结算 Then
        Set colBalance = New Collection
        strAdvanceOld = strAdvance
        
        For p = 1 To mobjBill.Pages.Count
            str结算方式 = "": colBalance.Add Array()
            '收费时划价单的对象属性NO没有存NO号
            strNo = IIf(mbytInState = EM_ED_收费, mobjBill.Pages(p).收费单号, mobjBill.Pages(p).NO)
            
            '检查该张单据是否已成功医保结算
            str结算方式 = zlGetYBBalanceNo(lng结帐ID, strNo)
            Call SetBalanceVal(colBalance, p, str结算方式)
            
            '没调用医保接口或为调用成功的单据重新进行医保结算
            If str结算方式 = "" Then
                strAdvance = strAdvanceOld & "|" & strNo
                str预结算 = GetMedicareStr(mcolBalance, p)
                '保存预结算结果
                '    Zl_医保结算明细_Insert(
                strSQL = "Zl_医保结算明细_Insert("
                '      结帐id_In   医保结算明细.结帐id%Type,
                strSQL = strSQL & "" & lng结帐ID & ","
                '      No_In       医保结算明细.No%Type,
                strSQL = strSQL & "'" & strNo & "',"
                '      结算方式_In Varchar2,
                strSQL = strSQL & "'" & str预结算 & "')"
                '      备注_In     医保结算明细.备注%Type := Null
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                cur全自付 = mobjBill.Pages(p).全自付
                cur先自付 = mobjBill.Pages(p).先自付
                '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
                If Not gclsInsure.ClinicSwap(lng结帐ID, GetMedicareSum(mcolBalance, mstr个人帐户, p), _
                                    GetMedicareSum(mcolBalance, "医保基金", p), cur全自付, cur先自付, _
                                    intInsure, strAdvance) Then Exit Function
                If strAdvance = strAdvanceOld & "|" & strNo Then strAdvance = ""
                
                If zlInsureCheck(str预结算, strAdvance) Then
                    str预结算 = strAdvance
                    '    Zl_医保结算明细_Insert(
                    strSQL = "Zl_医保结算明细_Insert("
                    '      结帐id_In   医保结算明细.结帐id%Type,
                    strSQL = strSQL & "" & lng结帐ID & ","
                    '      No_In       医保结算明细.No%Type,
                    strSQL = strSQL & "'" & strNo & "',"
                    '      结算方式_In Varchar2,
                    strSQL = strSQL & "'" & strAdvance & "')"
                    '      备注_In     医保结算明细.备注%Type := Null
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
                gcnOracle.CommitTrans '先提交，防止后续单据失败
                blnCommit = True
                
                Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
                Call SetBalanceVal(colBalance, p, str预结算)
                
                gcnOracle.BeginTrans
            End If
        Next
        
        '全部成功，返回总的结算方式
        strAdvance = GetMedicareStr(colBalance)
    ElseIf MCPAR.一次结算分单据退费 Then
        strAdvanceOld = strAdvance
        
        For p = 1 To mobjBill.Pages.Count
            strNo = mobjBill.Pages(p).收费单号

            '保存预结算结果
            str预结算 = GetMedicareStr(mcolBalance, p)
            '保存预结算结果
            '    Zl_医保结算明细_Insert(
            strSQL = "Zl_医保结算明细_Insert("
            '      结帐id_In   医保结算明细.结帐id%Type,
            strSQL = strSQL & "" & lng结帐ID & ","
            '      No_In       医保结算明细.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '      结算方式_In Varchar2,
            strSQL = strSQL & "'" & str预结算 & "')"
            '      备注_In     医保结算明细.备注%Type := Null
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            
            cur全自付 = cur全自付 + mobjBill.Pages(p).全自付
            cur先自付 = cur先自付 + mobjBill.Pages(p).先自付
        Next
            
        '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
        If Not gclsInsure.ClinicSwap(lng结帐ID, GetMedicareSum(mcolBalance, mstr个人帐户), _
                            GetMedicareSum(mcolBalance, "医保基金"), cur全自付, cur先自付, _
                            intInsure, strAdvance) Then Exit Function
        If strAdvance = strAdvanceOld Then strAdvance = ""
        If strAdvance = "" Then zlInsureClinicSwap = True: Exit Function
        
        'NO:结算方式,金额|结算方式,金额|...||NO:结算方式,金额|结算方式,金额|...||...
        Set colBalance = New Collection
        varAdvance = Split(strAdvance, "||")
        
        For p = 1 To mobjBill.Pages.Count
            '如果其中某一张单据不报销，没有返回对应结算信息，就按预结算结果保存
            blnFind = False
            For i = 0 To UBound(varAdvance)
                If InStr(varAdvance(i), ":") = 0 Then MsgBox "医保返回结算结果格式不正确！", vbInformation, gstrSysName: Exit Function
                
                varItem = Split(varAdvance(i), ":")
                strNo = varItem(0): str结算方式 = varItem(1)
                
                If strNo = mobjBill.Pages(p).收费单号 Then
                    str结算方式 = Replace(Replace(str结算方式, "|", "||"), ",", "|")
                    blnFind = True
                    Exit For
                End If
            Next
            
            If blnFind Then
                '直接修正医保结果，不检查是否需要校对
                '    Zl_医保结算明细_Insert(
                strSQL = "Zl_医保结算明细_Insert("
                '      结帐id_In   医保结算明细.结帐id%Type,
                strSQL = strSQL & "" & lng结帐ID & ","
                '      No_In       医保结算明细.No%Type,
                strSQL = strSQL & "'" & strNo & "',"
                '      结算方式_In Varchar2,
                strSQL = strSQL & "'" & str结算方式 & "')"
                '      备注_In     医保结算明细.备注%Type := Null
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            Else
                str结算方式 = GetMedicareStr(mcolBalance, p)
            End If
                
            colBalance.Add Array()
            SetBalanceVal colBalance, p, str结算方式
        Next
        strAdvance = GetMedicareStr(colBalance)
    Else
        '统计全自付和先自付金额
        For i = 1 To mobjBill.Pages.Count
            cur全自付 = cur全自付 + mobjBill.Pages(i).全自付
            cur先自付 = cur先自付 + mobjBill.Pages(i).先自付
        Next
        '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
        If Not gclsInsure.ClinicSwap(lng结帐ID, GetMedicareSum(mcolBalance, mstr个人帐户, 1), _
                            GetMedicareSum(mcolBalance, "医保基金", 1), cur全自付, cur先自付, _
                            intInsure, strAdvance) Then Exit Function
    End If
    zlInsureClinicSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSaveBillAndClinicSwapByNo(ByRef lng结帐ID As Long, ByRef strSavedNos As String, _
    ByRef cllChargeOverAfterPro As Collection, ByRef objChargeInfo As clsClinicChargeInfor, Optional ByRef blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保多单据分单据结算，按单据提交数据
    '参数：
    '   lng结帐ID
    '   strSaveNos - 结算成功的单据号：A001,A002,...
    '   cllChargeOverAfterPro - 完成收费后,执行的其他过程(主要是发料和发药)
    '返回:医保结算成功,返回true,否则返回False
    '说明：
    '    调用此过程时,不需要开始事务,异常时,数据回退,保存成功时,未提交数据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, p As Integer
    Dim cur全自付 As Currency, cur先自付 As Currency
    Dim colBalance As New Collection   '记录各张单据保险结算
    Dim strSQL As String, str预结算 As String, strNo As String
    Dim strAdvanceIn As String, strAdvance As String
    Dim cllSaveBillPro As Collection, blnTransMedicare As Boolean
    Dim cllPriceSQL As Collection, blnCommitPrice As Boolean '划价单是否已提交
    
    Err = 0: On Error GoTo errHandler
    If mintInsure = 0 Then Exit Function
    
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '不严格控制票据时保存当前票号
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "当前收费票据号", Trim(txtInvoice.Text), glngSys, 1121, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        End If
    End If
    
    If zlGetYBBalanceNo(lng结帐ID) <> "" Then '这种情况肯定是HIS保存数据出错，需要校对结算信息
        zlSaveBillAndClinicSwapByNo = True: Exit Function
    End If
    
    '1.检查收费数据是否合法
    If CheckChargeDataValied = False Then Exit Function
    '2.获取保存单据数据的SQL集合
    If SaveChargeBill(lng结帐ID, cllPriceSQL, cllSaveBillPro, cllChargeOverAfterPro) = False Then Exit Function
    '3.分单据进行结算
    strSavedNos = ""
    
    Err = 0: On Error GoTo errYBHandler
    gcnOracle.BeginTrans
    For p = 1 To mobjBill.Pages.Count
        str预结算 = "": colBalance.Add Array()
        strNo = mobjBill.Pages(p).收费单号
        blnCommitPrice = False
        '3.0无需判断单据是否结算成功，因为这种模式下，所有单据必然是一张都还未结算
        
        '3.1保存费用数据
        '先提交划价单，以便不锁表（药品库存）
        If CollectionExitsValue(cllPriceSQL, strNo) Then
            ExecuteProcedureArrAy cllPriceSQL(strNo), Me.Caption, False, True
            blnCommitPrice = True
            gcnOracle.BeginTrans
        End If
        If CollectionExitsValue(cllSaveBillPro, strNo) = False Then GoTo errYBHandler
        ExecuteProcedureArrAy cllSaveBillPro(strNo), Me.Caption, True, True
        
        '3.2保存预结算结果
        str预结算 = GetMedicareStr(mcolBalance, p)
        '    Zl_医保结算明细_Insert(
        strSQL = "Zl_医保结算明细_Insert("
        '      结帐id_In   医保结算明细.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '      No_In       医保结算明细.No%Type,
        strSQL = strSQL & "'" & strNo & "',"
        '      结算方式_In Varchar2,
        strSQL = strSQL & "'" & str预结算 & "')"
        '      备注_In     医保结算明细.备注%Type := Null
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '3.3调用医保接口
        strAdvance = CStr(-1 * lng结帐ID) & "|" & strNo '传入结算序号
        strAdvanceIn = strAdvance
        cur全自付 = mobjBill.Pages(p).全自付
        cur先自付 = mobjBill.Pages(p).先自付
        blnTransMedicare = False
        '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
        If Not gclsInsure.ClinicSwap(lng结帐ID, GetMedicareSum(mcolBalance, mstr个人帐户, p), _
                            GetMedicareSum(mcolBalance, "医保基金", p), cur全自付, cur先自付, _
                            mintInsure, strAdvance) Then GoTo errYBHandler
        blnTransMedicare = True '标记调用接口成功
        If strAdvance = strAdvanceIn Then strAdvance = ""
        
        '3.4校对医保结算结果
        If zlInsureCheck(str预结算, strAdvance) Then
            '    Zl_医保结算明细_Insert(
            strSQL = "Zl_医保结算明细_Insert("
            '      结帐id_In   医保结算明细.结帐id%Type,
            strSQL = strSQL & "" & lng结帐ID & ","
            '      No_In       医保结算明细.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '      结算方式_In Varchar2,
            strSQL = strSQL & "'" & strAdvance & "')"
            '      备注_In     医保结算明细.备注%Type := Null
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        
        Call SetBalanceVal(colBalance, p, strAdvance)
        strSavedNos = strSavedNos & "," & strNo
        gcnOracle.CommitTrans '先提交，防止后续单据失败
        blnCommit = True
        blnCommitPrice = False
        
        '交易确认
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
        blnTransMedicare = False
        
        '开启新事务
        gcnOracle.BeginTrans
    Next '下一张
    
    '全部成功，获取总的结算方式
    strAdvance = GetMedicareStr(colBalance)
    If strSavedNos <> "" Then strSavedNos = Mid(strSavedNos, 2)
    
    '105338，只要单据数量大于1就必须校对病人预交记录，因为病人预交记录中只有第一张单据的金额
    Call 医保数据更正(mobjBill.病人ID, lng结帐ID, GetMedicareStr(mcolBalance), strAdvance, True)
    gcnOracle.CommitTrans
    
    zlSaveBillAndClinicSwapByNo = True
    Exit Function
errYBHandler:
    gcnOracle.RollbackTrans
    Err = 0: On Error GoTo errHandler
    If blnCommitPrice Then
        '直接收费时,删除前一个事务提交的划价单
        Call DelMedicareTempNO(True, strNo)
    End If
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mintInsure)
    If blnCommit Then
        If Err.Description Like "*当前计算单价不一致*" Then
            MsgBox "某些分批药品价格已发生变化，结算中断！但部分单据已医保结算成功，现在将只对医保结算成功的这部分单据进行收费。", _
                vbInformation, gstrSysName
        End If
        
        '部分结算成功，只对结算成功这部分单据收费
        If strSavedNos <> "" Then strSavedNos = Mid(strSavedNos, 2)
        strAdvance = GetMedicareStr(colBalance)
        
        '105338，只要单据数量大于1就必须校对病人预交记录，因为病人预交记录中只有第一张单据的金额
        Call 医保数据更正(mobjBill.病人ID, lng结帐ID, GetMedicareStr(mcolBalance), strAdvance, True)
        
        '对未结算成功的单据进行处理
        For i = mobjBill.Pages.Count To p Step -1
            strNo = mobjBill.Pages(p).收费单号
            If CollectionExitsValue(cllChargeOverAfterPro, strNo) Then
                cllChargeOverAfterPro.Remove strNo
            End If
        Next
        Call GetChargeInfor(objChargeInfo, p - 1) '重新获取结算数据
        
        zlSaveBillAndClinicSwapByNo = True
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 医保数据更正(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal str预结算 As String, ByVal str医保结算 As String, _
    Optional ByVal blnMustCheckAdvance As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保数据校对更正
    '入参：
    '   blnMustCheckAdvance - 是否必须校对结算结果
    '返回:校对成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    If blnMustCheckAdvance = False Then
        If Not zlInsureCheck(str预结算, str医保结算) Then
            '修改校对标志
            ' Zl_病人门诊收费_医保更新
            strSQL = "Zl_病人门诊收费_医保更新("
            '  结帐id_In   门诊费用记录.结帐id%Type,
            strSQL = strSQL & lng结帐ID & ","
            '  结算序号_In 病人预交记录.结算序号%Type,
            strSQL = strSQL & "Null,"
            '  保险结算_In Varchar2
            strSQL = strSQL & "Null)"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            医保数据更正 = True: Exit Function
        End If
    End If
    
    'Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '  操作类型_In   Number,
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str医保结算 & "')"
    '  冲预交_In     病人预交记录.冲预交%Type,
    '  退支票额_In   病人预交记录.冲预交%Type,
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成结算_In Number:=0
    ') As
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --功能:收费结算时,修改结算的相关信息
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
    '  --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   3-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  -- 误差金额_In:存在误差费时,传入
    '  -- 完成结算_In:1-完成收费;0-未完成收费
    '  ------------------------------------------------------------------------------------------------------------------------------
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '修改校对标志
    ' Zl_病人门诊收费_医保更新
    strSQL = "Zl_病人门诊收费_医保更新("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSQL = strSQL & lng结帐ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "Null,"
    '  保险结算_In Varchar2
    strSQL = strSQL & "Null)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    医保数据更正 = True
End Function

Private Function IsSplitPrintByNO() As Boolean
    '是否按单据分别打印
    If mbytBillSource = 4 Then
        IsSplitPrintByNO = gTy_Module_Para.bln分别打印 And gTy_Module_Para.bln体检分别打印
    Else
        IsSplitPrintByNO = gTy_Module_Para.bln分别打印
    End If
End Function

Private Function GetFeeFromType() As String
    '获取收费单据来源类型
    '返回：1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    '说明：
    '   1.只要存在体检的费用单据(门诊标志=4)，则认为是体检费用
    '   2.只要存在住院的费用单据(门诊标志=2)，则认为是住院费用
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String, p As Integer
    Dim str费用来源 As String
    
    On Error GoTo errHandler
    For p = 1 To mobjBill.Pages.Count
         If mobjBill.Pages(p).NO = "" Then
            If gint病人来源 = 2 And InStr(str费用来源, "2") = 0 Then '住院
                str费用来源 = IIf(mTy_Para.bln住院病人门诊收费, "1", "2")
            ElseIf InStr(str费用来源, "1") = 0 Then  '门诊
                str费用来源 = "1"
            End If
         Else '提取的是划价单
            strNos = strNos & "," & mobjBill.Pages(p).NO
         End If
    Next
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
        strSQL = _
            "Select /*+cardinality(b, 10)*/ Nvl(Max(a.门诊标志), 0) As 门诊标志" & vbNewLine & _
            "From 门诊费用记录 A, Table(f_Str2list([1])) B" & vbNewLine & _
            "Where a.No = b.Column_Value And a.记录性质 = 1 And a.记录状态 = 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "单据来源类型", strNos)
        If rsTemp.EOF = False Then
            If InStr(str费用来源, Decode(Val(Nvl(rsTemp!门诊标志)), 4, 3, 2, 2, 1)) = 0 Then
                str费用来源 = str费用来源 & "," & Decode(Val(Nvl(rsTemp!门诊标志)), 4, 3, 2, 2, 1)
            End If
        End If
    End If
    If Left(str费用来源, 1) = "," Then str费用来源 = Mid(str费用来源, 2)
    GetFeeFromType = str费用来源
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



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
