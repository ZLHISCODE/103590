VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "ZLIDKIND.OCX"
Begin VB.Form frmLabSampleRegister 
   Caption         =   "检验标本登记"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   Icon            =   "frmLabSampleRegister.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12285
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   0
      Left            =   2970
      TabIndex        =   38
      Top             =   5250
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   1
      Left            =   1770
      TabIndex        =   39
      Top             =   5220
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   2
      Left            =   2970
      TabIndex        =   40
      Top             =   6240
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   3
      Left            =   1860
      TabIndex        =   41
      Top             =   6300
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.Timer Timer 
      Interval        =   60000
      Left            =   11715
      Top             =   105
   End
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   9780
      ScaleHeight     =   585
      ScaleWidth      =   1125
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox PicCuvetteCount 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   4710
      ScaleHeight     =   3105
      ScaleWidth      =   2805
      TabIndex        =   35
      Top             =   4890
      Width           =   2805
      Begin XtremeReportControl.ReportControl rptCuvetteCount 
         Height          =   1125
         Left            =   300
         TabIndex        =   36
         Top             =   1230
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption srtCuvetteCount 
         Height          =   285
         Left            =   150
         TabIndex        =   37
         Top             =   240
         Width           =   2235
         _Version        =   589884
         _ExtentX        =   3942
         _ExtentY        =   503
         _StockProps     =   6
         Caption         =   "当前已登记试管"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox PicCount 
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   7755
      ScaleHeight     =   2835
      ScaleWidth      =   3885
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4845
      Width           =   3885
      Begin XtremeReportControl.ReportControl rptCount 
         Height          =   1125
         Left            =   90
         TabIndex        =   33
         Top             =   450
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption SrtCount 
         Height          =   285
         Left            =   150
         TabIndex        =   34
         Top             =   60
         Width           =   2235
         _Version        =   589884
         _ExtentX        =   3942
         _ExtentY        =   503
         _StockProps     =   6
         Caption         =   "当前已登记的医嘱"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picBarCodeWork 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   60
      ScaleHeight     =   4770
      ScaleWidth      =   7995
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   7995
      Begin VB.CheckBox chkDept 
         BackColor       =   &H00FDD6C6&
         Caption         =   "按执行科室过滤"
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         ToolTipText     =   "若登记的标本是急诊标本则给出提示"
         Top             =   4455
         Width           =   1635
      End
      Begin VB.CheckBox chkUrgent 
         BackColor       =   &H00FDD6C6&
         Caption         =   "紧急标本提示"
         Height          =   255
         Left            =   6075
         TabIndex        =   46
         ToolTipText     =   "若登记的标本是急诊标本则给出提示"
         Top             =   4215
         Width           =   1515
      End
      Begin VB.CheckBox ChkBarCodeRegister 
         BackColor       =   &H00FDD6C6&
         Caption         =   "扫描条码后直接登记"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   4215
         Width           =   1965
      End
      Begin VB.CheckBox chkComRequest 
         BackColor       =   &H00FDD6C6&
         Caption         =   "检验标本送检后方可进行检验标本签收"
         Height          =   225
         Left            =   60
         TabIndex        =   44
         Top             =   4215
         Width           =   3390
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   405
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   714
         IDKindStr       =   "姓|姓名|0;医|医保号|1;身|身份证号|2;IC|IC卡号|3;门|门诊号|4;就|就诊卡|5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FDD6C6&
         Caption         =   "接收信息"
         Height          =   2385
         Left            =   60
         TabIndex        =   16
         Top             =   1800
         Width           =   7875
         Begin XtremeReportControl.ReportControl rptCuvette 
            Height          =   1635
            Left            =   120
            TabIndex        =   22
            Top             =   630
            Width           =   7605
            _Version        =   589884
            _ExtentX        =   13414
            _ExtentY        =   2884
            _StockProps     =   0
            AllowColumnRemove=   0   'False
            MultipleSelection=   0   'False
            SkipGroupsFocus =   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txt送检人 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5700
            MaxLength       =   20
            TabIndex        =   27
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "登记(&G)"
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
            Left            =   6765
            TabIndex        =   21
            Top             =   225
            Width           =   1065
         End
         Begin VB.TextBox txt接收时间 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1095
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   20
            Top             =   240
            Width           =   2010
         End
         Begin VB.TextBox txt接收人 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3870
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   18
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "送检人"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   4950
            TabIndex        =   26
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "接收时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "接收人"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   3150
            TabIndex        =   17
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.TextBox txtGoto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   780
         TabIndex        =   14
         ToolTipText     =   "数字为条码、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
         Top             =   315
         Width           =   7140
      End
      Begin VB.Frame FraPatientInfo 
         BackColor       =   &H00FDD6C6&
         Caption         =   "病人信息"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   60
         TabIndex        =   3
         Top             =   720
         Width           =   7875
         Begin VB.TextBox txt年龄 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5010
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   24
            Top             =   210
            Width           =   1155
         End
         Begin VB.TextBox txt性别 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   23
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox txt姓名 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Top             =   210
            Width           =   1635
         End
         Begin VB.TextBox txtBed 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6825
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox txtPatientDept 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   2445
         End
         Begin VB.TextBox txtID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   870
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   4515
            TabIndex        =   13
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2730
            TabIndex        =   12
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "所在科室"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   2730
            TabIndex        =   11
            Top             =   645
            Width           =   960
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   6330
            TabIndex        =   10
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标识号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓  名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   255
            Width           =   720
         End
      End
      Begin VB.CheckBox chkRemberPer 
         BackColor       =   &H00FDD6C6&
         Caption         =   "登记送检人后不清空"
         Height          =   255
         Left            =   90
         TabIndex        =   48
         ToolTipText     =   "若登记的标本是急诊标本则给出提示"
         Top             =   4470
         Width           =   1935
      End
      Begin VB.Label lbl显示费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1170
         TabIndex        =   29
         Top             =   30
         Width           =   60
      End
      Begin VB.Label lblGoto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人查找"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   30
         Width           =   900
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   8250
      ScaleHeight     =   2835
      ScaleWidth      =   3885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1470
      Width           =   3885
      Begin XtremeReportControl.ReportControl rptPlist 
         Height          =   1125
         Left            =   780
         TabIndex        =   30
         Top             =   630
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption srtPatient 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   60
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   6
         Caption         =   "病人信息"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8025
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabSampleRegister.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16589
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8340
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":0E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":0E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":19BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      WordWrap        =   -1  'True
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
   Begin XtremeSuiteControls.TabControl TabCtr 
      Height          =   1245
      Left            =   0
      TabIndex        =   42
      Top             =   5610
      Width           =   1965
      _Version        =   589884
      _ExtentX        =   3466
      _ExtentY        =   2196
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   8370
      Top             =   1050
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabSampleRegister.frx":1F58
      Left            =   8400
      Top             =   750
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabSampleRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mDkp                                   '窗格ID
    标本登记 = 0
    医嘱列表
    病人列表
    登记计数
    登记试管
End Enum
Private Enum mPcol                                  '病人列表
    病人ID = 1
    紧急
    来源
    病人姓名
    性别
    年龄
    标识号
    床号
    病人科室
    未登记
    已登记
    拒收
    已执行
    重采
    婴儿号
End Enum

Private Enum mAcol                                  '医嘱列表
    ID
    选择
    图标
    已执行
    采集方式
    医嘱内容
    条码
    执行科室
    开嘱医生
    开嘱时间
    发送人
    发送时间
    标本
    采样时间
    试管颜色
    合并医嘱
    试管编码
    采样人
    送检人
    采血量
    试管名称
    紧急
    病人来源
    申请科室
    婴儿
    别名
    相关ID
    医嘱id
    病人ID
    姓名
    性别
    年龄
    标识号
    床号
    病人科室
    接收人
    接收时间
    诊疗项目ID
    执行状态
    费用
    送检时间
End Enum
Private Enum mCuvette                               '试管
    选择
    编码
    名称
    添加剂
    采血量
    规格
    颜色
End Enum
Private Enum mCuvetteCount                          '试管计数
    编码
    名称
    合计
End Enum
Private Enum mFilter                                '过滤条件
    标识号 = 0
    就诊卡
    姓名
    单据号
    标本
    采集方式
    门诊
    住院
    体检
    病人科室
    间隔时间
    开始时间
    结束时间
End Enum
Private mlngKey As Long                             '病人ID
Private mlngDeptID As Long                          '科室ID
Private mstrPrivs As String
Private mlngBatch As Long                           '批号
Private mblnUse As Boolean                          '当前批次是否使用
Private mlngSelectBatch As Long                     '当前选择的批号
Private Enum IDKinds
    C0姓名 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
    C4门诊号 = 4
    C5就诊卡 = 5
End Enum
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object, mbln身份证 As Boolean
Private Const conMenu_IDkind_Change  As Integer = 12345
Private mobjSquareCard As Object                                        '取卡类型
Private mobjLisInsideComm As Object                                     'LIS内部接口
Private mblnShowPwd As Boolean                                          '是否显示密文
Private mintBabyNo  As Integer                                          '婴儿号
'条码类型
Private Enum BarCodeType
    Code39 = 1
    Code128 = 2
End Enum
Private mstrFirstBarCode   As String        '第一次扫条码
Private mintCodeType As Integer             '39码或128码
Private mrsSendPerson As Recordset          '送检人员记录集
Private mstrSendPerson As String            '送检人
    
Private Sub CreateCbs(Optional ByVal blnSecond As Boolean)
    '功能创建工具条
    
    '创建菜单
    Dim Control As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim ControlComboBox As CommandBarComboBox
    Dim intShowType As Integer
    
    '去掉扩展按钮
    cbrthis.VisualTheme = xtpThemeOffice2003
    Set cbrthis.Icons = zlCommFun.GetPubIcons
    With cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbrthis.EnableCustomization False
    cbrthis.ActiveMenuBar.Controls.DeleteAll
    '-----------------------------------------------------
    
    '==文件菜单
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"
        .Add xtpControlButton, conMenu_File_Preview, "预览(&V)"
        .Add xtpControlButton, conMenu_File_Print, "打印"
        Set Control = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): Control.BeginGroup = True
    End With
    
    '==编辑
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    With ControlFile.CommandBar.Controls
        Set Control = .Add(xtpControlButton, conMenu_Manage_Request, "登记(&G)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Insert, "开始计数(&I)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Import, "批量核对(&P)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Delete, "拒收(&R)")
    End With
    
    '==查看菜单
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    With ControlFile.CommandBar.Controls
        Set ControlSelect = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        ControlSelect.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        ControlSelect.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        ControlSelect.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set Control = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): Control.Checked = True
        Set Control = .Add(xtpControlButton, conMenu_View_Filter, "过滤(&F)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
    End With
   
    '==帮助菜单
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, conMenu_Help_Help, "帮助主题(&H)"
        .Add xtpControlButton, conMenu_Help_Web, "&WEB上的" & gstrProductName
        .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        .Add xtpControlButton, conMenu_Help_Web_Mail, "发送返馈(&S)", -1, False
        Set Control = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)")
        Control.BeginGroup = True
    End With
    
    '==列表科室
    If chkDept.Value = 0 Then
        Set ControlFile = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "采集科室")
    Else
        Set ControlFile = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "执行科室")
    End If
    With ControlFile.CommandBar.Controls
        intShowType = zlDatabase.GetPara("是否按病区显示", 100, 1212, 0)
        Set Control = .Add(xtpControlButton, conMenu_File_MedRecPreview, "按科室查看"): Control.Checked = (intShowType = 0)
        Set Control = .Add(xtpControlButton, conMenu_File_MedRecPrint, "按病区查看"): Control.Checked = (intShowType = 1)
    End With
    ControlFile.Flags = xtpFlagRightAlign
    Set ControlComboBox = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlComboBox, conMenu_View_Busy, "科室")
    ControlComboBox.ShortcutText = "科室"
    ControlComboBox.Width = 130
    ControlComboBox.Flags = xtpFlagRightAlign
    ControlComboBox.Style = xtpButtonIconAndCaption
    ControlComboBox.DropDownListStyle = True
    
    '批次
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "接收批次")
    Control.Flags = xtpFlagRightAlign
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlComboBox, conMenu_File_RoomSet, "批次")
    Control.ShortcutText = "批次"
    Control.Width = 130
    Control.Flags = xtpFlagRightAlign
    Control.Style = xtpButtonIconAndCaption
    
    If Not blnSecond Then
        '创建工具条
        Dim Toolbar As CommandBar
        Dim ControlPopup As CommandBarPopup
        
        Set Toolbar = cbrthis.Add("工具栏", xtpBarTop)
        Toolbar.ShowTextBelowIcons = False
        Toolbar.EnableDocking xtpFlagStretched
        With Toolbar.Controls
            .Add xtpControlButton, conMenu_File_Preview, "预览"
            .Add xtpControlButton, conMenu_File_Print, "打印"
            Set Control = .Add(xtpControlButton, conMenu_Manage_Request, "登记"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
            Set Control = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "重打条码")
            Control.Enabled = False
            Set Control = .Add(xtpControlButton, conMenu_Edit_Insert, "开始计数")
            Set Control = .Add(xtpControlButton, conMenu_Edit_Import, "批量核对")
            Set Control = .Add(xtpControlButton, conMenu_Edit_Delete, "拒收"): Control.BeginGroup = True
            
            Set Control = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            Set Control = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, conMenu_File_Exit, "退出"): Control.BeginGroup = True
        End With
        
        For Each Control In Toolbar.Controls
            Control.Style = xtpButtonIconAndCaption
        Next
    End If
    
    '快键绑定
    With cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add 0, VK_F2, conMenu_Manage_Request
        .Add 0, VK_F4, conMenu_Edit_Untread
        .Add 0, VK_F10, conMenu_IDkind_Change
        .Add 0, VK_F6, conMenu_Manage_Plan
    End With
    '设置不常用菜单
    With cbrthis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs)
End Sub

Private Function FindDept(ByVal cboCtrol As CommandBarComboBox, ByVal strTemp As String) As String
          '按编码、名称、简码查询部门
          Dim i As Integer
          Dim intShowType As Integer
          Dim rsTmp As Recordset
          Dim strSQL As String
          
1     On Error GoTo FindDept_Error

2         For i = 1 To cboCtrol.ListCount
3             If cboCtrol.List(i) = strTemp Then
4                 FindDept = strTemp
5                 Exit Function
6             End If
7         Next
          
8         intShowType = IIf(Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked, 0, 1)
          
9         If intShowType = 0 Then
10            strSQL = _
                      " Select Distinct A.ID,A.编码,A.名称" & _
                      " From 部门表 A,部门性质说明 B,部门人员 C " & _
                      " Where B.部门ID = A.ID And A.ID=C.部门ID " & _
                      " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                      " And B.服务对象 IN(1,2,3,4) And B.工作性质 IN('检验','护理','临床')"
11        Else
12            strSQL = _
                      " Select Distinct A.ID, A.编码, A.名称" & vbNewLine & _
                      " From 部门表 A, 病区科室对应 B" & vbNewLine & _
                      " Where B.病区id = A.ID And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)"
13        End If
          
14        strSQL = strSQL & " And (a.编码=[1] Or a.名称=[1] Or Upper(a.简码)=Upper([1]))"
          
15        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp)
          
16        If rsTmp.EOF Then
17            FindDept = ""
18            cboCtrol.ListIndex = 0
19        Else
20            For i = 1 To cboCtrol.ListCount
21                If cboCtrol.List(i) = rsTmp("编码") & "-" & rsTmp("名称") Then
22                    cboCtrol.ListIndex = i
23                    cboCtrol.Text = rsTmp("编码") & "-" & rsTmp("名称")
24                    FindDept = rsTmp("编码") & "-" & rsTmp("名称")
25                    Exit Function
26                End If
27            Next
28        End If


29        Exit Function
FindDept_Error:
30        MsgBox "zl9LisWork, frmLabSampleRegister, 执行(FindDept)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
31        Err.Clear
End Function

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strFilter As String                             '过滤字串
    Dim cboCtrol As CommandBarComboBox                  '科室
    Dim Controlcbo As CommandBarComboBox                '下拉框
    Dim cbrControl As CommandBarControl                 '文本标签
    Dim strText As String
    
    Select Case Control.ID
        Case conMenu_File_PrintSet                                                  '打印设置
            ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1212_1", Me
        
        Case conMenu_File_Preview                                                   '预览
'            Call zlRptPrint(2)
            RegisterLisPrint (1)
            
        Case conMenu_File_Print                                                     '打印
'            Call zlRptPrint(1)
            RegisterLisPrint (2)
       
        Case conMenu_File_Excel                                                     '输出到Excel
            Call zlRptPrint(4)
        
        Case conMenu_File_Exit                                                      '退出
            Unload Me
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Request                                                 '登记
            Call cmdOK_Click
            
        Case conMenu_Edit_Untread                                                   '取消
            Call cmdOK_Click
        Case conMenu_Edit_ReprintReceipt                                            '重打条码
            RePrintBarCode False
        Case conMenu_Edit_Insert                                                    '计数
            BeginRegister   '开始计数
                    
        Case conMenu_Edit_Import                                                    '批量核对
            frmLabSampleCheck.ShowMe Me
                    
        Case conMenu_Edit_Delete                                                    '拒收
            frmLabSampleRegisterRefuse.ShowMe Me, rptAlist(Me.TabCtr.Selected.Index).Records
            RefreshPatientData
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_ToolBar_Button                                            '标准按钮
            Me.cbrthis(2).Visible = Not Me.cbrthis(2).Visible
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                              '文本标签
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        Case conMenu_View_ToolBar_Size                                              '大图标
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
            
        Case conMenu_View_StatusBar                                                 '状态栏
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Control.Checked
            Me.cbrthis.RecalcLayout
            
        Case conMenu_View_Filter                                                    '过滤
            frmLabSampleRegisterFilter.ShowMe Me, strFilter
            Me.rptPlist.Tag = strFilter
            If strFilter <> "" Then RefreshPatientData
            
        Case conMenu_View_Refresh                                                   '刷新
            RefreshPatientData
        
        Case conMenu_IDkind_Change
            Call IdKindChange
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Help_Help                                                      '帮助主题
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web                                                       'Web上的中联
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Home                                                  '主页
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                                                  '发送反馈
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                                     '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_Busy                                                      '科室选择
            Set cboCtrol = Control
            If Trim(cboCtrol.Text) = "" Then Exit Sub
            strText = FindDept(cboCtrol, cboCtrol.Text)
            mlngDeptID = cboCtrol.ItemData(cboCtrol.ListIndex)
            RefreshPatientData
            If strText = "" Then cboCtrol.SetFocus
        Case conMenu_File_RoomSet                                                   '批次选择
            Set cboCtrol = Control
            mlngSelectBatch = cboCtrol.ItemData(cboCtrol.ListIndex)
            RefreshPatientData
        Case conMenu_File_MedRecPreview                                             '按科室查看
            Control.Checked = Not Control.Checked
            Me.cbrthis.FindControl(, conMenu_File_MedRecPrint, , True).Checked = Not Control.Checked
            Call GetDept
        Case conMenu_File_MedRecPrint                                               '按病区查看
            Control.Checked = Not Control.Checked
            Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked = Not Control.Checked
            Call GetDept
        Case Else

            If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            
    End Select
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.Visible = False Then Exit Sub
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
        
    Err = 0: On Error Resume Next
    Select Case Control.ID
        
        Case conMenu_File_Preview, conMenu_File_Print
            Control.Enabled = (mlngSelectBatch <> 0)
        
        Case conMenu_View_ToolBar_Button:                                                                   '按钮
            Control.Checked = Me.cbrthis(2).Visible
        
        Case conMenu_View_ToolBar_Text:                                                                     '按钮文字
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        
        Case conMenu_View_ToolBar_Size:                                                                     '大图标
            Control.Checked = Me.cbrthis.Options.LargeIcons
        
        Case conMenu_View_StatusBar:                                                                        '状态栏
            Control.Checked = Me.stbThis.Visible
                    
        Case conMenu_Manage_Request                                                                         '登记
            Control.Enabled = (Me.TabCtr.Selected.Index = 0 Or Me.TabCtr.Selected.Index = 3)
            
        Case conMenu_Edit_Untread                                                                           '取消
            Control.Enabled = (Me.TabCtr.Selected.Index = 1)
            
        Case conMenu_Edit_Delete                                                                            '拒收
            Control.Enabled = (Me.TabCtr.Selected.Index = 0 Or Me.TabCtr.Selected.Index = 1)
        
        Case conMenu_File_MedRecPreview
            Control.Checked = Control.Checked
            
        Case conMenu_File_MedRecPrint
            Control.Checked = Control.Checked
        Case conMenu_Edit_ReprintReceipt
            RePrintBarCode True
        Case conMenu_Edit_Import                                                                            '批量核对
            Control.Visible = InStr(";" & mstrPrivs & ";", ";批量核对;") > 0
    End Select
End Sub

Private Sub Timer_Timer()
    Me.txt接收时间.Text = zlDatabase.Currentdate
End Sub

Private Sub chkDept_Click()
    Call CreateCbs(True)              '创建工具栏
    Call GetDept                      '读入科室
End Sub

Private Sub cmdOK_Click()
    '登记和取消登记
    If Me.TabCtr.Selected.Index <= 1 Or Me.TabCtr.Selected.Index = 4 Then
        SaveRegister Me.TabCtr.Selected.Index
        If Me.TabCtr.Selected.Index = 4 Then Me.cmdOK.Enabled = False
        Me.txtGoto.SetFocus
        txtGoto.Text = ""
    End If
    RefreshPatientData 1, mintBabyNo
End Sub

Private Sub Form_Load()

    On Error GoTo errH

    mstrPrivs = gstrPrivs       '初使化权限
    mintCodeType = zlDatabase.GetPara("使用条码", "100", "1211", 2)
    ChkBarCodeRegister.Value = zlDatabase.GetPara("扫描条码后直接登记", "100", "1212", 0)
    chkUrgent.Value = zlDatabase.GetPara("紧急标本提示", "100", "1212", 0)
    chkDept.Value = zlDatabase.GetPara("按执行科室过滤", "100", "1212", 0)
    Call CreateCbs              '创建工具栏
    Call CreateDkp              '创建窗格
    Call CreateTab              '创建Tab列表
    Call CreateListHead         '创建表头
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    mbln身份证 = False
    chkComRequest = zlDatabase.GetPara("检验标本送检后方可进行检验标本签收", 100, 1212, 0)
    chkRemberPer.Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记送检人后不清空", 1)
    '数据读入
    Call GetDept                                                '读入科室
    Call GetPerson                                              '读入送检人
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind初始化失败!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If
        
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
    RefreshPatientData
    
    Me.txt接收人.Text = UserInfo.姓名
    Me.txt接收时间.Text = zlDatabase.Currentdate
    
    If mobjLisInsideComm Is Nothing Then
        Dim strErr As String
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        '初始化LIS接口部件
        If Not mobjLisInsideComm Is Nothing Then
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "初始化LIS接口失败！" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If
   
    Exit Sub
errH:
    MsgBox "初始化窗体时发生错误，请检查部件后重试！", vbInformation, "初始化"
End Sub

Private Sub GetPerson()
    Dim strSQL As String
    strSQL = "Select Distinct d.Id, d.编号, d.姓名, d.简码" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B, 部门人员 C, 人员表 D" & vbNewLine & _
            "Where a.Id = b.部门id And a.Id = c.部门id And c.人员id = d.Id And a.撤档时间 Is Not Null And" & vbNewLine & _
            "      a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd') And (b.工作性质 = '临床' Or b.工作性质 = '治疗' Or b.工作性质 = '护理' and 1=0)" & vbNewLine & _
            "Order By d.编号"
    Set mrsSendPerson = zlDatabase.OpenSQLRecord(strSQL, "查询送检人")
End Sub

Private Sub CreateDkp()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    
    dkpMan.SetCommandBars Me.cbrthis
    dkpMan.Options.DefaultPaneOptions = PaneNoCloseable
    dkpMan.Options.HideClient = True
    
    Set Pane1 = dkpMan.CreatePane(mDkp.标本登记, 400, 700, DockLeftOf, Nothing)
    Pane1.Title = "标本登记"
    Pane1.Handle = Me.picBarCodeWork.hWnd
    Pane1.Options = PaneNoCaption
    
    Set Pane2 = dkpMan.CreatePane(mDkp.医嘱列表, 400, 300, DockBottomOf, Pane1)
    Pane2.Title = "医嘱信息"
    Pane2.Handle = Me.TabCtr.hWnd
    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMan.CreatePane(mDkp.病人列表, 600, 300, DockRightOf, Nothing)
    Pane3.Title = "病人采集清单"
    Pane3.Handle = Me.picTab.hWnd
    Pane3.Options = PaneNoCaption
    
    Set Pane4 = dkpMan.CreatePane(mDkp.登记计数, 600, 150, DockBottomOf, Pane3)
    Pane4.Title = "登记计数"
    Pane4.Handle = Me.PicCount.hWnd
    Pane4.Options = PaneNoCaption
    
    Set Pane5 = dkpMan.CreatePane(mDkp.登记试管, 600, 150, DockBottomOf, Pane4)
    Pane5.Title = "登记试管"
    Pane5.Handle = Me.PicCuvetteCount.hWnd
    Pane5.Options = PaneNoCaption
    
    Pane1.Select
    
End Sub
Private Sub CreateTab()
    Dim Item As TabControlItem
    
    With Me.TabCtr
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 0, "未登记", Me.rptAlist(0).hWnd, 0
        .InsertItem 1, "已登记", Me.rptAlist(1).hWnd, 0
        .InsertItem 2, "执行中", Me.rptAlist(2).hWnd, 0
        .InsertItem 3, "拒收", Me.rptAlist(3).hWnd, 0
'
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane
    If Me.Visible = False Then Exit Sub
    Set Pane1 = Me.dkpMan.FindPane(mDkp.标本登记)
    Pane1.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane1.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    
    Me.cbrthis.RecalcLayout
    
    Set Pane2 = Me.dkpMan.FindPane(mDkp.病人列表)
    Pane2.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane2.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    
    Set Pane3 = Me.dkpMan.FindPane(mDkp.登记计数)
    Pane3.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane3.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
'
    Set Pane4 = Me.dkpMan.FindPane(mDkp.登记试管)
    Pane4.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane4.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    
    
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    
    Pane1.MinTrackSize.SetSize 100, 100
    Pane2.MinTrackSize.SetSize 100, 100
    Pane3.MinTrackSize.SetSize 100, 100
    Pane4.MinTrackSize.SetSize 100, 100

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngSelectBatch = 0
    mlngBatch = 0
    Set mobjSquareCard = Nothing
    Call zlDatabase.SetPara("是否按病区显示", IIf(Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked, 0, 1), 100, 1212)
    zlDatabase.SetPara "检验标本送检后方可进行检验标本签收", chkComRequest, 100, 1212
    Call zlDatabase.SetPara("扫描条码后直接登记", IIf(ChkBarCodeRegister.Value = 1, 1, 0), 100, 1212)
    Call zlDatabase.SetPara("紧急标本提示", IIf(chkUrgent.Value = 1, 1, 0), 100, 1212)
    Call zlDatabase.SetPara("按执行科室过滤", IIf(chkDept.Value = 1, 1, 0), 100, 1212)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记送检人后不清空", chkRemberPer.Value)
End Sub

Private Sub IDKind_Click()
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String, strOutPatiInforXML As String
    If IDKind.IDKind = IDKinds.C3IC卡号 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtGoto.Text = mobjICCard.Read_Card()
            If txtGoto.Text <> "" Then Call txtGoto_KeyPress(vbKeyReturn)
        End If
    End If
    lng卡类别ID = Val(IDKind.GetKindItem("卡类别ID"))
    If lng卡类别ID = 0 Then Exit Sub
    
    If mobjSquareCard.zlReadCard(Me, glngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtGoto.Text = strOutCardNO
    If txtGoto.Text <> "" Then Call txtGoto_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    mblnShowPwd = Trim(IDKind.GetKindItem(7)) <> ""
    Me.txtGoto = ""
    If mblnShowPwd = True Then
        Me.txtGoto.PasswordChar = "*"
    Else
        Me.txtGoto.PasswordChar = ""
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    mbln身份证 = False
    If Not txtGoto.Locked And txtGoto.Text = "" And Me.ActiveControl Is txtGoto Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C2身份证号
        txtGoto.Text = strID
        mbln身份证 = True
        Call txtGoto_KeyPress(vbKeyReturn)
        mbln身份证 = False
        IDKind.IDKind = lngPreIDKind
    End If
End Sub

Private Sub PicCount_Resize()
    On Error Resume Next
    
    With SrtCount
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
    End With
    
    With rptCount
        .Top = Me.srtPatient.Top + Me.srtPatient.Height + 10
        .Left = 0
        .Width = Me.PicCount.ScaleWidth
        .Height = Me.PicCount.ScaleHeight - .Top
    End With
    

End Sub

Private Sub PicCuvetteCount_Resize()
    On Error Resume Next
    
    With srtCuvetteCount
        .Top = 0
        .Left = 0
        .Width = Me.PicCuvetteCount.ScaleWidth
    End With
    
    With rptCuvetteCount
        .Top = Me.srtCuvetteCount.Top + Me.srtCuvetteCount.Height + 10
        .Left = 0
        .Width = Me.PicCuvetteCount.ScaleWidth
        .Height = Me.PicCuvetteCount.ScaleHeight - .Top
    End With
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    
    With srtPatient
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
    End With
    
    With rptPlist
        .Top = Me.srtPatient.Top + Me.srtPatient.Height + 10
        .Left = 0
        .Width = Me.picTab.ScaleWidth
        .Height = Me.picTab.ScaleHeight - .Top
    End With
End Sub
Private Sub CreateListHead()
    '创建列表头
    Dim Column As ReportColumn
    Dim intLoop As Integer
    
    '==医嘱列表头
    
    rptPlist.AllowColumnRemove = False
    rptPlist.ShowItemsInGroups = False
    
    With rptPlist.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
        
    End With
    rptPlist.SetImageList ImgList
    
    With Me.rptPlist.Columns
        Set Column = .Add(mPcol.病人ID, "病人ID", 0, False)
        
        Set Column = .Add(mPcol.来源, "来源", 45, True)
        Set Column = .Add(mPcol.病人姓名, "病人姓名", 75, True)
        Set Column = .Add(mPcol.性别, "性别", 60, True)
        Set Column = .Add(mPcol.年龄, "年龄", 60, True)
        Set Column = .Add(mPcol.病人科室, "病人科室", 75, True)
        Set Column = .Add(mPcol.标识号, "标识号", 60, True)
        Set Column = .Add(mPcol.床号, "床号", 60, True)
        
        Set Column = .Add(mPcol.未登记, "未登记", 45, True)
        Set Column = .Add(mPcol.已登记, "已登记", 45, True)
        Set Column = .Add(mPcol.拒收, "拒收", 30, True)
        Set Column = .Add(mPcol.已执行, "已执行", 45, True)
        Set Column = .Add(mPcol.重采, "重采", 30, True)
        Set Column = .Add(mPcol.婴儿号, "婴儿号", 0, False)
    End With
    
    For intLoop = 0 To 3
        '==病人列表头
        rptAlist(intLoop).AllowColumnRemove = False
        rptAlist(intLoop).ShowItemsInGroups = False
        With rptAlist(intLoop).PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
            .HideSelection = True
        End With
        rptAlist(intLoop).SetImageList ImgList
        With Me.rptAlist(intLoop).Columns
            Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
            Set Column = .Add(mAcol.选择, "Check", 18, False): Column.Icon = 0
            Set Column = .Add(mAcol.图标, "", 18, False): Column.Icon = 3
            Set Column = .Add(mAcol.已执行, "已执行", 45, False): Column.Visible = False: Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.费用, "费用", 30, False): Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.采集方式, "采集方式", 75, True)
            Set Column = .Add(mAcol.标本, "标本", 55, True)
            Set Column = .Add(mAcol.医嘱内容, "医嘱内容", 75, True)
            Set Column = .Add(mAcol.条码, "条码", 75, True)
            Set Column = .Add(mAcol.执行科室, "执行科室", 75, True)
            Set Column = .Add(mAcol.开嘱医生, "开嘱医生", 75, True)
            Set Column = .Add(mAcol.开嘱时间, "开嘱时间", 75, True)
            Set Column = .Add(mAcol.发送人, "发送人", 65, True)
            Set Column = .Add(mAcol.送检人, "送检人", 65, True)
            Set Column = .Add(mAcol.接收人, "接收人", 65, True)
            Set Column = .Add(mAcol.发送时间, "发送时间", 75, True)
            Set Column = .Add(mAcol.采样时间, "采样时间", 75, True)
            Set Column = .Add(mAcol.接收时间, "接收时间", 75, True)
            Set Column = .Add(mAcol.试管颜色, "颜色编码", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.试管编码, "试管编码", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.采样人, "采样人", 60, True)
            Set Column = .Add(mAcol.采血量, "采血量", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.试管名称, "试管名称", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.紧急, "紧急", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.病人来源, "病人来源", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.申请科室, "申请科室", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.婴儿, "婴儿", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.别名, "别名", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.相关ID, "相关ID", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.病人ID, "病人ID", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.姓名, "姓名", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.性别, "性别", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.年龄, "年龄", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.标识号, "标识号", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.床号, "床号", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.病人科室, "病人科室", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.诊疗项目ID, "诊疗项目Id", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.执行状态, "执行状态", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.医嘱id, "医嘱ID", 50, False): Column.Visible = False
            Set Column = .Add(mAcol.送检时间, "送检时间", 75, True)
            
        End With
    Next
    
    '==病人列表头
    rptCount.AllowColumnRemove = False
    rptCount.ShowItemsInGroups = False
    With rptCount.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
        .HideSelection = True
    End With
    rptCount.SetImageList ImgList
    With rptCount.Columns
        Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
        Set Column = .Add(mAcol.选择, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mAcol.图标, "", 18, False): Column.Icon = 3
        Set Column = .Add(mAcol.已执行, "已执行", 45, False): Column.Visible = False: Column.Alignment = xtpAlignmentCenter
        Set Column = .Add(mAcol.费用, "费用", 30, False): Column.Alignment = xtpAlignmentCenter
        Set Column = .Add(mAcol.病人来源, "来源", 35, True)
        Set Column = .Add(mAcol.姓名, "姓名", 70, True)
        Set Column = .Add(mAcol.性别, "性别", 35, True)
        Set Column = .Add(mAcol.标识号, "标识号", 70, True)
        Set Column = .Add(mAcol.采集方式, "采集方式", 75, True)
        Set Column = .Add(mAcol.标本, "标本", 55, True)
        Set Column = .Add(mAcol.医嘱内容, "医嘱内容", 75, True)
        Set Column = .Add(mAcol.条码, "条码", 75, True)
        Set Column = .Add(mAcol.执行科室, "执行科室", 75, True)
        Set Column = .Add(mAcol.开嘱医生, "开嘱医生", 75, True)
        Set Column = .Add(mAcol.开嘱时间, "开嘱时间", 80, True)
        Set Column = .Add(mAcol.发送人, "发送人", 65, True)
        Set Column = .Add(mAcol.送检人, "送检人", 65, True)
        Set Column = .Add(mAcol.接收人, "接收人", 65, True)
        Set Column = .Add(mAcol.发送时间, "发送时间", 80, True)
        Set Column = .Add(mAcol.采样时间, "采样时间", 80, True)
        Set Column = .Add(mAcol.接收时间, "接收时间", 80, True)
        Set Column = .Add(mAcol.试管颜色, "颜色编码", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.试管编码, "试管编码", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.采样人, "采样人", 60, True)
        Set Column = .Add(mAcol.采血量, "采血量", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.试管名称, "试管名称", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.紧急, "紧急", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.申请科室, "申请科室", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.婴儿, "婴儿", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.别名, "别名", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.相关ID, "相关ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.病人ID, "病人ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.年龄, "年龄", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.床号, "床号", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.病人科室, "病人科室", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.诊疗项目ID, "诊疗项目Id", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.执行状态, "执行状态", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.医嘱id, "医嘱ID", 50, False): Column.Visible = False
        Set Column = .Add(mAcol.送检时间, "送检时间", 80, True)
    End With
    
    '==试管
    rptCuvette.AllowColumnRemove = False
    rptCuvette.ShowItemsInGroups = False
    
    With rptCuvette.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
        
    End With
    rptCuvette.SetImageList ImgList
    With Me.rptCuvette.Columns
        Set Column = .Add(mCuvette.选择, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mCuvette.编码, "编码", 75, True)
        Set Column = .Add(mCuvette.名称, "名称", 120, True)
        Set Column = .Add(mCuvette.添加剂, "添加剂", 110, True)
        Set Column = .Add(mCuvette.采血量, "采血量", 80, True)
        Set Column = .Add(mCuvette.规格, "规格", 80, True)
        Set Column = .Add(mCuvette.颜色, "", 18, True): Column.Icon = 3
    End With
    
    '==试管计数
    rptCuvetteCount.AllowColumnRemove = False
    rptCuvetteCount.ShowItemsInGroups = False
    
    With rptCuvetteCount.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
        
    End With
    rptCuvetteCount.SetImageList ImgList
    With Me.rptCuvetteCount.Columns
        Set Column = .Add(mCuvetteCount.编码, "编码", 75, True)
        Set Column = .Add(mCuvetteCount.名称, "名称", 120, True)
        Set Column = .Add(mCuvetteCount.合计, "合计", 45, True)
    End With
End Sub

Private Sub RefreshPatientData(Optional lngPatientType As Long = 0, Optional ByVal intBabyNo As Integer = 0)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String                                '用于SQL语句存放
    Dim strSQL1 As String                               '用于H表的查找
    Dim Record As ReportRecord
    Dim Item As ReportColumn
    Dim intLoop As Integer
    Dim strTmp As String                                '临时字串变量
    Dim varFilter As Variant                            '过滤字串
    Dim strDateBegin As Date                            '开始时间
    Dim strDateEnd As Date                              '结束时间
    Dim blnDateMoved As Boolean                         '是否被转出
    Dim lngPatientID As Long                            '病人ID
    Dim strState As String                              '状态
    Dim strDeptIDs As String                            '病区下ID
    Dim int主页ID As Integer                            '主页ID
    Dim str挂号单 As String                             '挂号单
    Dim intPatientType As Integer                       '病人来源
    
    On Error GoTo errH
    
    '更新批次
    If TabCtr.Selected.Index > 0 Then
        GetBatch
    End If
    
    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("标本登记过滤", 100, 1212, "")
    
    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    zlCommFun.ShowFlash "正在更新数据,请稍候...", Me
    
    strSQL = "Select 病人id,病人来源,病人姓名,病人科室,性别,年龄,就诊卡号,标识号,当前床号," & vbNewLine & _
                "       Sum(decode(状态,'未登记',1,0)) As 未登记,Sum(decode(状态,'已登记',1,0)) As 已登记," & vbNewLine & _
                "       Sum(decode(状态,'拒收',1,0)) As 拒收,Sum(decode(状态,'已执行',1,0)) As 已执行," & vbNewLine & _
                "       Sum(decode(紧急,'紧急',1,0)) As 紧急,Sum(重采) As 重采,Nvl(sum(婴儿)/count(婴儿),0) as 婴儿号 From ("
    
    If chkDept.Value = 1 Then
        '按执行科室过滤
        strSQL = strSQL & "Select Distinct a.相关id, a.病人id, Decode(a.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检') As 病人来源," & vbNewLine & _
                        "                Decode(Nvl(a.婴儿, 0), 0, c.姓名, t.婴儿姓名) As 病人姓名, e.名称 As 病人科室, Decode(Nvl(a.婴儿, 0), 0, c.性别, t.婴儿性别) As 性别," & vbNewLine & _
                        "                Decode(Nvl(a.婴儿, 0), 0, c.年龄, Nvl(Round(Nvl(t.死亡时间, Sysdate) - t.出生时间), 0) || '天') As 年龄," & vbNewLine & _
                        "                c.就诊卡号, b.样本条码, Decode(b.执行状态, 1, '已执行', 2, '拒收', 3, '已执行', Decode(b.接收人, Null, '未登记', '已登记')) As 状态," & vbNewLine & _
                        "                a.婴儿, Decode(a.病人来源, 1, c.门诊号, 2, c.住院号) As 标识号," & vbNewLine & _
                        "                Decode(c.当前床号, Null, Decode(l.出院病床, Null, l.入院病床, l.出院病床), c.当前床号) As 当前床号," & vbNewLine & _
                        "                Decode(a.紧急标志, 1, '紧急', Decode(g.急诊, 1, '紧急')) As 紧急, Decode(b.执行状态, 0, '', 2, '拒收') As 拒收," & vbNewLine & _
                        "                Nvl(b.重采标本, 0) As 重采, k.试管编码" & vbNewLine & _
                        "From 病人医嘱记录 H, 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 部门表 E, 诊疗项目目录 F, 病人挂号记录 G, 诊疗项目目录 K, 病案主页 L, 病人新生儿记录 T" & vbNewLine & _
                        "Where h.Id = a.相关id And a.Id = b.医嘱id And a.病人id = c.病人id And a.病人科室id = e.Id And a.诊疗项目id = k.Id And h.诊疗项目id = f.Id And" & vbNewLine & _
                        "      a.病人id = t.病人id(+) And a.主页id = t.主页id(+) And a.婴儿 = t.序号(+) And a.挂号单 = g.No(+) And" & vbNewLine & _
                        "      (g.病人id Is Null Or (g.记录状态 = 1 And g.记录性质 = 1)) And a.诊疗类别 = 'C' And h.诊疗类别 = 'E' And f.操作类型 = '6' And" & vbNewLine & _
                        "      a.病人id = l.病人id(+) And b.执行部门id  In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And" & vbNewLine & _
                        "      k.试管编码 Is Not Null And b.执行状态 In (0, 1, 2, 3) And b.发送数次 = 1 "
    Else
        '按采集科室过滤
        strSQL = strSQL & "Select distinct h.相关id,a.病人id,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源, " & vbCrLf & _
                        " decode(nvl(a.婴儿,0),0,c.姓名,t.婴儿姓名)  as 病人姓名,e.名称 as 病人科室,decode(nvl(a.婴儿,0),0,c.性别,t.婴儿性别) as 性别," & vbCrLf & _
                        " decode(nvl(a.婴儿,0),0,c.年龄,Nvl(Round(Nvl(t.死亡时间, Sysdate) - t.出生时间), 0) ||'天') as 年龄,c.就诊卡号,b.样本条码, " & vbCrLf & _
                        " decode(b.执行状态,1, '已执行',2,'拒收',3,'已执行', decode(b.接收人,null,'未登记','已登记')) as 状态,a.婴儿, " & vbCrLf & _
                        " decode(A.病人来源, 1, C.门诊号, 2, C.住院号) As 标识号, " & vbCrLf & _
                        " decode(c.当前床号,null,decode(l.出院病床,null,l.入院病床,l.出院病床),c.当前床号) as 当前床号 , " & vbCrLf & _
                        " decode(a.紧急标志,1,'紧急',decode(g.急诊,1,'紧急')) as 紧急 , " & vbCrLf & _
                        " decode(b.执行状态,0,'',2,'拒收') as 拒收,nvl(b.重采标本,0) as 重采, k.试管编码 " & vbCrLf & _
                        " From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 部门表 E, 诊疗项目目录 F,病人挂号记录 G,病人医嘱记录 H, " & vbCrLf & _
                        "      诊疗项目目录 K ,病案主页 L,病人医嘱发送 M,病人新生儿记录 T" & vbCrLf & _
                        " Where A.ID = H.相关ID And H.id = B.医嘱id And A.病人id = C.病人id And A.病人科室id = E.ID And A.诊疗项目id = f.ID " & vbCrLf & _
                        "      And h.诊疗项目ID = k.id " & vbCrLf & _
                        "and A.病人ID=T.病人ID(+) and A.主页ID=T.主页ID(+) and A.婴儿=T.序号(+)" & vbCrLf & _
                        " And A.挂号单 = G.No(+) and (g.病人ID is null or (g.记录状态=1 and g.记录性质 =1) ) and a.诊疗类别 = 'E' and F.操作类型 = '6' and a.病人id = l.病人ID(+) " & vbCrLf & _
                        " and m.执行部门id  in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & vbCrLf & _
                        " And A.ID = M.医嘱ID And k.试管编码 is not null and B.执行状态 in (0,1,2,3) and m.发送数次 = 1 "
    End If
    
    '查找病区ID
    gstrSql = "select 科室ID from 病区科室对应 where 病区ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
    If rsTmp.EOF = True Then
        strDeptIDs = mlngDeptID
    Else
        Do Until rsTmp.EOF
            strDeptIDs = strDeptIDs & "," & rsTmp("科室ID")
            rsTmp.MoveNext
        Loop
        If strDeptIDs <> "" Then strDeptIDs = Mid(strDeptIDs, 2) & "," & mlngDeptID
    End If
    
    '批次
    If mlngSelectBatch <> 0 And TabCtr.Selected.Index > 0 Then
        strSQL = strSQL & " and b.接收批次 = [11] "
    End If
    
    '如果没有权限，就不能看到未采样标本
    If InStr(mstrPrivs, "强制登记未采样标本") = 0 Then
        strSQL = strSQL & " and b.采样人 is not null "
    End If
    
    If Me.rptPlist.Tag <> "" Then
        strSQL = strSQL & " And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3,", "") & _
                 Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "
    
        If varFilter(mFilter.标识号) <> "" Then
            strSQL = strSQL & " And decode(a.病人来源,2,c.住院号,c.门诊号) = [2] "
        End If
        
        If varFilter(mFilter.就诊卡) <> "" Then
            strSQL = strSQL & " And C.就诊卡号 = [3] "
        End If
        
        If varFilter(mFilter.姓名) <> "" Then
            strSQL = strSQL & " And C.姓名 like [4] "
        End If
        
        If varFilter(mFilter.单据号) <> "" Then
            strSQL = strSQL & " and B.NO = [5]"
        End If
        
        If varFilter(mFilter.标本) <> "所有标本" Then
            strSQL = strSQL & " and A.标本部位 = [6] "
        End If
        
        If varFilter(mFilter.采集方式) <> 0 Then
            strSQL = strSQL & " and f.ID +0 = [7] "
        End If
        
        If varFilter(mFilter.病人科室) <> 0 Then
            strSQL = strSQL & " And a.病人科室ID = [8] "
        End If
        
        If lngPatientType = 1 Then
            strSQL = strSQL & " And c.病人id=[12]"
            '使用病人ID时不使用
            strSQL = strSQL & " and b.发送时间+0 Between [9] and [10]"
        Else
            strSQL = strSQL & " and b.发送时间 Between [9] and [10]"
        End If
        
        If varFilter(mFilter.开始时间) = "" Then
            strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.间隔时间))
            strDateEnd = zlDatabase.Currentdate
        Else
            strDateBegin = varFilter(mFilter.开始时间)
            strDateEnd = varFilter(mFilter.结束时间)
        End If
    Else
        If strTmp <> "" Then
            strSQL = strSQL & " And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3,", "") & _
                 Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "
    
            If varFilter(mFilter.标本) <> "所有标本" Then
                strSQL = strSQL & " and A.标本部位 = [6] "
            End If
            
            If varFilter(mFilter.采集方式) <> 0 Then
                strSQL = strSQL & " and f.ID +0= [7] "
            End If
            
            If varFilter(mFilter.病人科室) <> 0 Then
                strSQL = strSQL & " And a.病人科室ID = [8] "
            End If
            
            If lngPatientType = 1 Then
                strSQL = strSQL & " And c.病人id=[12]"
                '使用病人ID时不使用
                strSQL = strSQL & " and b.发送时间+0 Between [9] and [10]"
            Else
                strSQL = strSQL & " and b.发送时间 Between [9] and [10]"
            End If
            
            If Val(varFilter(mFilter.间隔时间)) >= 0 Then
                strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.间隔时间))
                strDateEnd = zlDatabase.Currentdate
            Else
                strDateBegin = varFilter(mFilter.开始时间)
                strDateEnd = varFilter(mFilter.结束时间)
            End If
        Else
            If lngPatientType = 1 Then
                strSQL = strSQL & " And c.病人id=[12]"
                '使用病人ID时不使用
                strSQL = strSQL & " and b.发送时间+0 Between [9] and [10]"
            Else
                strSQL = strSQL & " and b.发送时间 Between [9] and [10]"
            End If
            strDateBegin = zlDatabase.Currentdate - 3
            strDateEnd = zlDatabase.Currentdate
        End If
    End If
    
    strSQL = strSQL & ") Group By  病人id,病人来源,病人姓名,病人科室,性别,年龄,就诊卡号,标识号,当前床号 "
    
    blnDateMoved = MovedByDate(CDate(strDateBegin)) '按时间看是否可能已转出
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    strSQL = strSQL & " Order by 病人科室 "
    
    If strTmp = "" And Me.rptPlist.Tag = "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strDeptIDs, "", "", "", "", "", "", "", _
                    CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), mlngSelectBatch, mlngKey)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strDeptIDs, Val(varFilter(mFilter.标识号)), CStr(varFilter(mFilter.就诊卡)) _
                    , CStr(varFilter(mFilter.姓名)) & "%", CStr(varFilter(mFilter.单据号)), CStr(varFilter(mFilter.标本)), CLng(varFilter(mFilter.采集方式)) _
                    , mlngDeptID, CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), _
                    CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), mlngSelectBatch, mlngKey)
    End If
    If lngPatientType <> 1 Then
        '清除记录
        Me.rptPlist.Records.DeleteAll
        Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
        Me.rptCuvette.Records.DeleteAll
        
        Do Until rsTmp.EOF
            Set Record = Me.rptPlist.Records.Add
                                    
            For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                Record.AddItem ""
            Next
            
            Record(mPcol.病人ID).Value = Nvl(rsTmp("病人ID"))
            Record(mPcol.紧急).Value = Nvl(rsTmp("紧急"))
            If Nvl(rsTmp("紧急")) = "紧急" Then Record(mPcol.紧急).Icon = 2
            Record(mPcol.来源).Value = Nvl(rsTmp("病人来源"))
            Record(mPcol.病人姓名).Value = Nvl(rsTmp("病人姓名"))
            Record(mPcol.病人科室).Value = Nvl(rsTmp("病人科室"))
            Record(mPcol.性别).Value = Nvl(rsTmp("性别"))
            Record(mPcol.年龄).Value = Nvl(rsTmp("年龄"))
            Record(mPcol.标识号).Value = Nvl(rsTmp("标识号"))
            Record(mPcol.床号).Value = Nvl(rsTmp("当前床号"))
            
            Record(mPcol.未登记).Value = Nvl(rsTmp("未登记"))
            Record(mPcol.已登记).Value = Nvl(rsTmp("已登记"))
            Record(mPcol.拒收).Value = Nvl(rsTmp("拒收"))
            Record(mPcol.已执行).Value = Nvl(rsTmp("已执行"))
            Record(mPcol.重采).Value = Nvl(rsTmp("重采"))
            Record(mPcol.婴儿号).Value = Nvl(rsTmp("婴儿号"))
            
            If Nvl(rsTmp("拒收"), 0) > 0 Then
                For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                    Record(intLoop).ForeColor = vbRed
                Next
            End If
            
            If Nvl(rsTmp("重采"), 0) > 0 Then
                For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                    Record(intLoop).Bold = True
                    Record(intLoop).ForeColor = vbBlue
                Next
            End If
            rsTmp.MoveNext
        Loop
    Else
        rsTmp.filter = "婴儿号=" & "'" & intBabyNo & "'"
        Do Until rsTmp.EOF
            For intLoop = 0 To Me.rptPlist.Rows.Count - 1
                If Me.rptPlist.Rows(intLoop).Record(mPcol.病人ID).Value = mlngKey And Me.rptPlist.Rows(intLoop).Record(mPcol.婴儿号).Value = intBabyNo And Me.rptPlist.Rows(intLoop).Record(mPcol.来源).Value = Nvl(rsTmp("病人来源")) Then
                    Me.rptPlist.Rows(intLoop).Record(mPcol.未登记).Value = Nvl(rsTmp("未登记"))
                    Me.rptPlist.Rows(intLoop).Record(mPcol.已登记).Value = Nvl(rsTmp("已登记"))
                End If
            Next
            rsTmp.MoveNext
        Loop
    End If
    
    '更新
    If Me.Visible = True Then
        
    Me.rptPlist.Populate
        
    End If
    If Me.Visible = True Then
        Me.rptAlist(TabCtr.Selected.Index).Populate
        Me.rptCuvette.Populate
    End If
    Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptPlist.Rows.Count & "个病人！"
    
    '定位到上次选中的病人
    With Me.rptPlist

        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mPcol.病人ID).Value = mlngKey And .Rows(intLoop).Record(mPcol.婴儿号).Value = intBabyNo Then
                Set .FocusedRow = .Rows(intLoop)
                mlngKey = .Rows(intLoop).Record(mPcol.病人ID).Value
                If .Rows(intLoop).Record(mPcol.来源).Value = "门诊" Then
                    intPatientType = 1
                ElseIf .Rows(intLoop).Record(mPcol.来源).Value = "住院" Then
                    intPatientType = 2
                ElseIf .Rows(intLoop).Record(mPcol.来源).Value = "院外" Then
                    intPatientType = 3
                ElseIf .Rows(intLoop).Record(mPcol.来源).Value = "体检" Then
                    intPatientType = 4
                Else
                    intPatientType = 1
                End If
                .Populate
                Me.rptPlist.Tag = ""
                Exit For
            End If
        Next
        
        If .FocusedRow Is Nothing And .Rows.Count > 0 Then
            Set .FocusedRow = .Rows(0)
            mlngKey = .Rows(0).Record(mPcol.病人ID).Value
            If .Rows(0).Record(mPcol.来源).Value = "门诊" Then
                intPatientType = 1
            ElseIf .Rows(0).Record(mPcol.来源).Value = "住院" Then
                intPatientType = 2
            ElseIf .Rows(0).Record(mPcol.来源).Value = "院外" Then
                intPatientType = 3
            ElseIf .Rows(0).Record(mPcol.来源).Value = "体检" Then
                intPatientType = 4
            Else
                intPatientType = 1
            End If
            .Populate
        End If
        
        If Not .FocusedRow Is Nothing Then
            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index, intPatientType, False, intBabyNo
        End If
        
    End With
    
    '过滤中条件只执行一次
    Me.rptPlist.Tag = ""
    
    If Me.rptPlist.Rows.Count = 0 Then
        txt姓名 = ""
        txt姓名.Tag = ""
        txt性别.Text = ""
        txt年龄 = ""
        txtBed = ""
        txtID = ""
        txtPatientDept = ""
    End If
    
    zlCommFun.StopFlash
    
    Exit Sub
errH:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDept()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim Controlcbo As CommandBarComboBox
    Dim lngDept As Long
    Dim intLoop As Integer
    Dim intShowType As Integer
    
    mlngDeptID = zlDatabase.GetPara("科室", 100, 1212, 0)
    intShowType = IIf(Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked, 0, 1)
    
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    
    On Error GoTo errH
    
    If intShowType = 0 Then
        strSQL = _
                " Select Distinct A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B,部门人员 C " & _
                " Where B.部门ID = A.ID And A.ID=C.部门ID " & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                " And B.服务对象 IN(1,2,3,4) And B.工作性质 IN('检验','护理','临床')"
    Else
        strSQL = _
                " Select Distinct A.ID, A.编码, A.名称" & vbNewLine & _
                " From 部门表 A, 病区科室对应 B" & vbNewLine & _
                " Where B.病区id = A.ID And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)"
    End If
'    If InStr(1, mstrPrivs, "所有科室") <= 0 Then
'        strSQL = strSQL & " And C.人员ID = [1] "
'    End If
    
    strSQL = strSQL & " Order by A.编码"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Controlcbo.Clear
    Controlcbo.ListIndex = -1
    Do Until rsTmp.EOF
        Controlcbo.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        Controlcbo.ItemData(Controlcbo.ListCount) = rsTmp("ID")
        If rsTmp("id") = IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID) Then
            Controlcbo.ListIndex = Controlcbo.ListCount
            mlngDeptID = IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID)
        End If
        rsTmp.MoveNext
    Loop
    If Controlcbo.Text = "" Then
        Controlcbo.ListIndex = 1
        mlngDeptID = Controlcbo.ItemData(Controlcbo.ListIndex)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function RefreshAdviceData(lngPatientID As Long, intState As Integer, Optional intPatientType As Integer = 0, Optional blOnlyWhere As Boolean = False, Optional intBabyNo As Integer = 0) As Boolean
    '功能：                         刷新采集医嘱记录
    '参数：                         lngpatientId = 病人ID ,
    '                               intPatientType = 病人来源
    '                               intState = 当前状态 0=未登记 1=已登记 2=已执行
    '                               blOnlyWhere = true 只使用病人ID进行查找
    '                               intBabyNo  婴儿号
    
    Dim blnDateMoved As Boolean                     '是否转出
    Dim strSQL As String                            'SQL语句
    Dim strSQL1 As String                           '用于查询H表
    Dim strTmp As String                            '临时字串变量
    Dim varFilter As Variant                        '过滤字串
    Dim strDateBegin As String                      '开始发送时间
    Dim strDateEnd As String                        '结束发送时间
    Dim rsTmp As New ADODB.Recordset                '数据集
    Dim intLoop As Integer                          '循环变量
    Dim Record As ReportRecord                      '列表数据集
    Dim strOldAdvice As String                      '记录上次医嘱
    Dim strCuvetteNumber As String                  '用于记录试管编码
    Dim intShowButtom As Integer                    '用于在查找时显示那些按钮,条码 not null = 1 采样人 not null = 2
    Dim strDeptIDs As String                        '病区下ID
    Dim int主页ID As Integer                        '主页ID
    Dim str挂号单 As String                         '挂号单
    Dim str已收费医嘱ID  As String, str所有医嘱ID As String
    Dim strSQLbak As String
    Dim strGetSql As String, intMainID As Integer
    Dim rsTest As ADODB.Recordset
    Dim strPatiDept As String
    Dim strOldNO As String
    Dim strOldCodeBar As String

    On Error GoTo errH
    
    blnDateMoved = MovedByDate(Date) '按时间看是否可能已转出
    
    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("标本登记过滤", 100, 1212, "")
    
    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    If chkDept.Value = 1 Then
        strSQL = "Select Distinct /*+ rule */ b.Id As 医嘱id, a.主页id, b.相关id, g.颜色 As 试管颜色, d.名称 As 采集方式, b.医嘱内容, c.样本条码, c.采样时间," & vbNewLine & _
                "                c.标本送出时间 As 送检时间, h.名称 As 执行科室, b.开嘱医生, b.开嘱时间, c.发送人, c.发送时间, g.编码 As 试管编码, b.标本部位 As 标本, i.姓名 As 病人姓名," & vbNewLine & _
                "                i.性别," & vbNewLine & _
                "                Decode(Nvl(a.婴儿, 0), 0, i.年龄, Decode(Sysdate - k.出生时间, '', '', Round(Sysdate - k.出生时间) || '天')) As 年龄," & vbNewLine & _
                "                i.当前床号 As 床号, Decode(b.病人来源, 1, i.门诊号, 2, i.住院号) As 标识号, k.婴儿姓名, k.婴儿性别, l.名称 As 病人所在科室," & vbNewLine & _
                "                Decode(c.执行状态, 2, '拒收') As 拒收, i.病人id, c.采样人, c.送检人, g.采血量, g.名称 As 试管名称," & vbNewLine & _
                "                Decode(b.紧急标志, 1, '紧急', '') As 紧急, Decode(b.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检') As 病人来源, b.婴儿," & vbNewLine & _
                "                n.名称 As 别名, j.名称 As 病人科室, c.接收时间, c.接收人, b.诊疗项目id, c.执行状态, o.记录性质, o.记录状态" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, 采血管类型 G, 部门表 H, 病人信息 I, 部门表 L, 部门表 J, 病人新生儿记录 K," & vbNewLine & _
                "     (Select 诊疗项目id, 名称 From 诊疗项目别名 Where 性质 = 9 And 码类 = 1) N, 门诊费用记录 O" & vbNewLine & _
                "Where a.Id = b.相关id And b.Id = c.医嘱id And a.诊疗项目id = d.Id And b.诊疗项目id = e.Id And e.类别 = 'C' And e.试管编码 = g.编码 And" & vbNewLine & _
                "      b.执行科室id = h.Id And d.类别 = 'E' And d.操作类型 = '6' And a.病人id = [1] And c.发送时间 + 0 Between [3] And [4] And" & vbNewLine & _
                "      b.病人id = i.病人id And i.当前科室id = l.Id(+) And a.病人id = k.病人id(+) And a.主页id = k.主页id(+) And a.婴儿 = k.序号(+) And" & vbNewLine & _
                "      e.Id = n.诊疗项目id(+) And b.病人科室id = j.Id(+) And c.医嘱id = o.医嘱序号(+) And c.记录性质 = Mod(o.记录性质(+), 10) And" & vbNewLine & _
                "      Nvl(o.记录状态, 0) In (0, 1) And b.病人来源 = [11] "
    Else
        strSQL = " Select distinct /*+ rule */ B.ID as 医嘱ID,a.主页id, B.相关id, G.颜色 As 试管颜色, D.名称 As 采集方式, B.医嘱内容, C.样本条码,C.采样时间,c.标本送出时间 as 送检时间, " & vbCrLf & _
                 " H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & _
                 " I.姓名 as 病人姓名,I.性别,decode(nvl(a.婴儿,0),0,i.年龄,decode(sysdate-k.出生时间,'','',round(sysdate-k.出生时间)||'天')) as 年龄," & vbCrLf & _
                 "i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号) as 标识号,K.婴儿姓名,K.婴儿性别, " & vbCrLf & _
                 " L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人ID,c.采样人,c.送检人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
                 " DECODE(B.紧急标志,1,'紧急','') as 紧急,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源, " & vbCrLf & _
                 " b.婴儿,N.名称 as 别名,J.名称 as 病人科室,C.接收时间,C.接收人,b.诊疗项目ID,C.执行状态,O.记录性质,O.记录状态 " & vbCrLf & _
                 " From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
                 " 采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,部门表 J,病人新生儿记录 K , " & vbCrLf & _
                 " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 O " & vbCrLf & _
                 " Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
                 " And E.类别 = 'C' And E.试管编码 = G.编码 And B.执行科室id = H.ID " & vbCrLf & _
                 " And D.类别 = 'E' And D.操作类型 = '6' And A.病人id = [1] And c.发送时间+0 Between [3] and [4] " & vbCrLf & _
                 " And B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & vbCrLf & _
                 " and A.病人ID=K.病人ID(+) and A.主页ID=K.主页ID(+) and A.婴儿=k.序号(+)" & vbCrLf & _
                 " and a.id = m.医嘱id And E.id = N.诊疗项目ID(+) And b.病人科室id = J.id(+)  " & vbCrLf & _
                 " and c.医嘱id = O.医嘱序号(+) and c.记录性质 =Mod(O.记录性质(+),10) and nvl(O.记录状态,0) in (0,1) And b.病人来源 = [11] "
    End If
    
    '查找病区Id
    gstrSql = "select 科室ID from 病区科室对应 where 病区ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
    If rsTmp.EOF = True Then
        strDeptIDs = mlngDeptID
    Else
        Do Until rsTmp.EOF
            strDeptIDs = strDeptIDs & "," & rsTmp("科室ID")
            rsTmp.MoveNext
        Loop
        If strDeptIDs <> "" Then strDeptIDs = Mid(strDeptIDs, 2) & "," & mlngDeptID
    End If
    
    '批次
    If mlngSelectBatch <> 0 And TabCtr.Selected.Index > 0 Then
        strSQL = strSQL & " and c.接收批次 = [7] "
    End If
    
    '执行科室
    If chkDept.Value = 1 Then
        strSQL = strSQL & " and b.执行科室ID + 0 in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) "
    Else
        strSQL = strSQL & " and a.执行科室ID + 0 in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) "
    End If
    
    '病人科室
    If rptPlist.FocusedRow Is Nothing Then
    Else
        strPatiDept = rptPlist.FocusedRow.Record(mPcol.病人科室).Value
        strSQL = strSQL & " and J.名称 = [12] "
    End If
    
    '如果没有权限，就不能看到未采样标本
    If InStr(mstrPrivs, "强制登记未采样标本") = 0 Then
        strSQL = strSQL & " and c.采样人 is not null "
    End If
    
    '处理三种不同的状态
    If intState = 0 Then
        strSQL = strSQL & " And C.执行状态 in(0) and C.接收人 is null " & vbCrLf
    ElseIf intState = 1 Or intState = 4 Then
        strSQL = strSQL & " And C.执行状态 in(0) and C.接收人 is not null " & vbCrLf
    ElseIf intState = 2 Then
        strSQL = strSQL & " And C.执行状态 in (1,3) " & vbCrLf
    ElseIf intState = 3 Then
        strSQL = strSQL & " And C.执行状态 in (2) " & vbCrLf
    End If
    
    '过滤
    If Me.rptPlist.Tag <> "" Or strTmp <> "" Then
        If varFilter(mFilter.标本) <> "所有标本" Then
            strSQL = strSQL & " and A.标本部位 = [5] "
        End If
        
        If varFilter(mFilter.采集方式) <> 0 Then
            strSQL = strSQL & " and d.ID + 0 = [6] "
        End If
        
        If Me.rptPlist.Tag <> "" Then
            strDateBegin = varFilter(mFilter.开始时间)
            strDateEnd = varFilter(mFilter.结束时间)
        Else
            strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.间隔时间))
            strDateEnd = zlDatabase.Currentdate
        End If
    Else
        strDateBegin = zlDatabase.Currentdate - 3
        strDateEnd = zlDatabase.Currentdate
    End If
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    If blOnlyWhere = True Then
        If chkDept.Value = 1 Then
            strSQL = "Select Distinct /*+ rule */ b.Id As 医嘱id, a.主页id, b.相关id, g.颜色 As 试管颜色, d.名称 As 采集方式, b.医嘱内容, c.样本条码, c.采样时间," & vbNewLine & _
                    "                c.标本送出时间 As 送检时间, h.名称 As 执行科室, b.开嘱医生, b.开嘱时间, c.发送人, c.发送时间, g.编码 As 试管编码, b.标本部位 As 标本, i.姓名 As 病人姓名," & vbNewLine & _
                    "                i.性别," & vbNewLine & _
                    "                Decode(Nvl(a.婴儿, 0), 0, i.年龄, Decode(Sysdate - k.出生时间, '', '', Round(Sysdate - k.出生时间) || '天')) As 年龄," & vbNewLine & _
                    "                i.当前床号 As 床号, Decode(b.病人来源, 1, i.门诊号, 2, i.住院号) As 标识号, k.婴儿姓名, k.婴儿性别, l.名称 As 病人所在科室," & vbNewLine & _
                    "                Decode(c.执行状态, 2, '拒收') As 拒收, i.病人id, c.采样人, c.送检人, g.采血量, g.名称 As 试管名称," & vbNewLine & _
                    "                Decode(b.紧急标志, 1, '紧急', '') As 紧急, Decode(b.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检') As 病人来源, b.婴儿," & vbNewLine & _
                    "                n.名称 As 别名, j.名称 As 病人科室, c.接收时间, c.接收人, b.诊疗项目id, c.执行状态, o.记录性质, o.记录状态, a.挂号单" & vbNewLine & _
                    "From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, 采血管类型 G, 部门表 H, 病人信息 I, 部门表 L, 部门表 J, 病人新生儿记录 K," & vbNewLine & _
                    "     (Select 诊疗项目id, 名称 From 诊疗项目别名 Where 性质 = 9 And 码类 = 1) N, 门诊费用记录 O" & vbNewLine & _
                    "Where a.Id = b.相关id And b.Id = c.医嘱id And a.诊疗项目id = d.Id And b.诊疗项目id = e.Id And e.类别 = 'C' And e.试管编码 = g.编码 And" & vbNewLine & _
                    "      b.执行科室id = h.Id And d.类别 = 'E' And d.操作类型 = '6' And a.病人id = [1] And" & vbNewLine & _
                    "      b.病人id = i.病人id And i.当前科室id = l.Id(+) And a.病人id = k.病人id(+) And a.主页id = k.主页id(+) And a.婴儿 = k.序号(+) And" & vbNewLine & _
                    "      e.Id = n.诊疗项目id(+) And i.当前科室id = j.Id(+) And c.医嘱id = o.医嘱序号(+) And c.记录性质 = Mod(o.记录性质(+), 10) And" & vbNewLine & _
                    "      Nvl(o.记录状态, 0) In (0, 1) "
        Else
            strSQL = " Select distinct /*+ rule */ B.ID as 医嘱ID,a.主页id, B.相关id, G.颜色 As 试管颜色, D.名称 As 采集方式, B.医嘱内容, C.样本条码,C.采样时间,c.标本送出时间 as 送检时间, " & vbCrLf & _
                 " H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & _
                 " I.姓名 as 病人姓名,I.性别,decode(nvl(a.婴儿,0),0,i.年龄,decode(sysdate-k.出生时间,'','',round(sysdate-k.出生时间)||'天')) as 年龄," & vbCrLf & _
                 " i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号) as 标识号, K.婴儿姓名,K.婴儿性别, " & vbCrLf & _
                 " L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人Id,c.采样人,c.送检人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
                 " DECODE(B.紧急标志,1,'紧急','') as 紧急,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源,b.婴儿,N.名称 as 别名,J.名称 as 病人科室,C.接收时间,C.接收人, " & vbCrLf & _
                 " b.诊疗项目ID,C.执行状态,O.记录性质,O.记录状态,a.挂号单 " & vbCrLf & _
                 " From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
                 " 采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,部门表 J,病人新生儿记录 K , " & vbCrLf & _
                 " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 O " & vbCrLf & _
                 " Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
                 " And E.类别 = 'C' And E.试管编码 = G.编码 And m.执行部门id = H.ID " & vbCrLf & _
                 " And D.类别 = 'E' And D.操作类型 = '6' And A.病人id = [1] And " & vbCrLf & _
                 " B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & _
                 " and A.病人ID=K.病人ID(+) and A.主页ID=K.主页ID(+) and A.婴儿=k.序号(+)" & vbCrLf & _
                 " And a.id  = m.医嘱id And E.id = N.诊疗项目ID(+) And I.当前科室id = J.id(+)  " & vbCrLf & _
                 " and c.医嘱id = O.医嘱序号(+) and c.记录性质 =mod(O.记录性质(+),10) and nvl(O.记录状态,0) in (0,1) "
        End If
        
        '处理三种不同的状态
        If intState = 0 Or intState = 4 Then
            strSQL = strSQL & " And C.执行状态 in(0) and C.接收人 is null " & vbCrLf
        ElseIf intState = 1 Then
            strSQL = strSQL & " And C.执行状态 in(1,0,3) and C.接收人 is not null " & vbCrLf
        ElseIf intState = 2 Then
            strSQL = strSQL & " And C.执行状态 in(1,3) and C.接收人 is not null " & vbCrLf
        ElseIf intState = 3 Then
            strSQL = strSQL & " And C.执行状态 in(2) " & vbCrLf
        End If
        If IDKind.IDKind = IDKinds.C0姓名 And BlnIsNumber(txtGoto) Then
            strSQL = strSQL & " And c.样本条码 = [10]   "
        Else
            '用于判断病人是住院还是门诊
            gstrSql = "Select 主页id, 出院日期 From 病案主页 Where 病人id = [1] Order By 主页id Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngPatientID)
            If rsTmp.EOF = False Then
                If Nvl(rsTmp("出院日期")) = "" Then
                    int主页ID = Nvl(rsTmp("主页Id"), 0)
                    strSQL = strSQL & " And a.主页id = [8] "
                Else
                    gstrSql = "Select NO " & vbNewLine & _
                            " From 病人挂号记录 A, 病人医嘱记录　b " & vbNewLine & _
                            " Where A.病人id = B.病人id And B.病人来源 = 1 And a.记录状态=1 and a.记录性质 =1 and A.病人id = [1] " & vbNewLine & _
                            " Order By A.ID Desc "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngPatientID)
                    If rsTmp.EOF = False Then
                        str挂号单 = Nvl(rsTmp("NO"))
                        'strSQL = strSQL & " And A.挂号单 = [9] "
                    Else
                        strSQL = strSQL & " And nvl(c.样本条码,0) <> 0   "
                        'strSQL = strSQL & " And a.病人来源<> 2 "
                    End If
                End If
            Else
                gstrSql = "Select NO " & vbNewLine & _
                        " From 病人挂号记录 A, 病人医嘱记录　b " & vbNewLine & _
                        " Where A.病人id = B.病人id And B.病人来源 = 1 And a.记录状态 =1 and a.记录性质 =1  and  A.病人id = [1] " & vbNewLine & _
                        " Order By A.ID Desc "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngPatientID)
                If rsTmp.EOF = False Then
                    str挂号单 = Nvl(rsTmp("NO"))
                    'strSQL = strSQL & " And A.挂号单 = [9] "
                Else
                    strSQL = strSQL & " And a.病人来源<> 2 "
                End If
            End If
            If IDKind.IDKind = IDKinds.C0姓名 And BlnIsNumber(txtGoto) Then
                strSQL = strSQL & " And c.样本条码 = [10]   "
            End If
        End If
    End If
    
    If intPatientType <> 2 Then
        strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    End If
    
'    strSQLbak = strSQL
'    strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
'    strSQL = strSQL & " union all " & strSQLbak
    
    strSQL = strSQL & IIf(intBabyNo <> 0, " and nvl(a.婴儿,0)= " & intBabyNo, IIf(blOnlyWhere, "", "and nvl(a.婴儿,0) = 0 ")) & " Order By 样本条码, 试管编码, 相关id, 医嘱id, 标本, 开嘱时间 "
    
    If strTmp <> "" Or rptPlist.Tag <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, strDeptIDs, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), varFilter(mFilter.标本), CLng(Val(varFilter(mFilter.采集方式))), mlngSelectBatch, _
                int主页ID, str挂号单, txtGoto, intPatientType, strPatiDept)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, strDeptIDs, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), "", 0, mlngSelectBatch, int主页ID, str挂号单, txtGoto, intPatientType, strPatiDept)
    End If
    
    Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll

    Me.rptCuvette.Records.DeleteAll
    Me.rptAlist(TabCtr.Selected.Index).Columns(mAcol.已执行).Visible = False
    If rsTmp.RecordCount < 1 Then
        Me.txt送检人 = ""
    End If
    Do Until rsTmp.EOF
        If intState = 0 Then
            If intPatientType = 2 Then
                If intMainID = 0 Then
                    intMainID = Val(rsTmp("主页ID") & "")
                    If intMainID <> 0 Then
                        strGetSql = "Select 主页id, 出院日期 From 病案主页 Where 病人id = [1] and 主页id=[2] "
                        Set rsTest = zlDatabase.OpenSQLRecord(strGetSql, Me.Caption, lngPatientID, intMainID)
                        If rsTest.RecordCount > 0 Then
                            If Nvl(rsTest("出院日期")) <> "" Then
                                If MsgBox("该病人已出院，是否继续执行！", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                                    Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
                                    Me.rptCuvette.Records.DeleteAll
                                    Me.rptAlist(TabCtr.Selected.Index).Populate
                                    Me.rptCuvette.Populate
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        '没有对应颜色编码的采集不写入
        If IsNull(rsTmp("试管颜色")) = False Then
            If strOldAdvice <> rsTmp("相关ID") & "" Or strOldNO <> rsTmp("试管编码") & "" Or strOldCodeBar <> rsTmp("样本条码") & "" Then
                Set Record = Me.rptAlist(TabCtr.Selected.Index).Records.Add
                For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                    Record.AddItem ""
                Next
                
                If blOnlyWhere = True And (Nvl(rsTmp("执行状态")) = 1 Or Nvl(rsTmp("执行状态")) = 3) Then
                    Me.rptAlist(TabCtr.Selected.Index).Columns(mAcol.已执行).Visible = True
                End If
                
                If Nvl(rsTmp("执行状态")) = 1 Or Nvl(rsTmp("执行状态")) = 3 Then
                    Record(mAcol.已执行).Value = "√"
                Else
                    Record(mAcol.已执行).Value = ""
                End If
                
                Record(mAcol.选择).HasCheckbox = True
                
                If str挂号单 = "" Then
                    Record(mAcol.选择).Checked = IIf(Nvl(rsTmp("执行状态")) = 0, True, False)
                Else
                    Record(mAcol.选择).Checked = IIf(Nvl(rsTmp("执行状态")) = 0 And str挂号单 = Nvl(rsTmp("挂号单")), True, False)
                End If
                
                Record(mAcol.费用).Value = IIf(rsTmp("记录状态") = 1, "√", "×")
                
                
                Record(mAcol.ID).Value = Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                
                
                Record(mAcol.图标).BackColor = Val(Nvl(rsTmp("试管颜色")))
                Record(mAcol.采集方式).Value = Nvl(rsTmp("采集方式"))
                Record(mAcol.医嘱内容).Value = Nvl(rsTmp("医嘱内容"))
                Record(mAcol.条码).Value = Nvl(rsTmp("样本条码"))
                Record(mAcol.执行科室).Value = Nvl(rsTmp("执行科室"))
                Record(mAcol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
                Record(mAcol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
                Record(mAcol.发送人).Value = Nvl(rsTmp("发送人"))
                Record(mAcol.发送时间).Value = Nvl(rsTmp("发送时间"))
                Record(mAcol.试管颜色).Value = Nvl(rsTmp("试管颜色"))
                Record(mAcol.试管编码).Value = Nvl(rsTmp("试管编码"))
                Record(mAcol.标本).Value = Nvl(rsTmp("标本")) & IIf(Nvl(rsTmp("婴儿")) = 0, "", "(婴儿)")
                Record(mAcol.采样时间).Value = Nvl(rsTmp("采样时间"))
                Record(mAcol.采样人).Value = Nvl(rsTmp("采样人"))
                Record(mAcol.送检人).Value = Nvl(rsTmp("送检人"))
                Record(mAcol.采血量).Value = Nvl(rsTmp("采血量"))
                Record(mAcol.试管名称).Value = Nvl(rsTmp("试管名称"))
                Record(mAcol.紧急).Value = Nvl(rsTmp("紧急"))
                Record(mAcol.病人来源).Value = Nvl(rsTmp("病人来源"))
                Record(mAcol.婴儿).Value = Nvl(rsTmp("婴儿"))
                Record(mAcol.别名).Value = Nvl(rsTmp("别名"))
                Record(mAcol.相关ID).Value = Nvl(rsTmp("相关ID"))
                
                Record(mAcol.病人ID).Value = Nvl(rsTmp("病人ID"))
                Record(mAcol.姓名).Value = Nvl(rsTmp("病人姓名"))
                Record(mAcol.性别).Value = Nvl(rsTmp("性别"))
                Record(mAcol.年龄).Value = Nvl(rsTmp("年龄"))
                Record(mAcol.标识号).Value = Nvl(rsTmp("标识号"))
                Record(mAcol.床号).Value = Nvl(rsTmp("床号"))
                Record(mAcol.病人科室).Value = Nvl(rsTmp("病人科室"))
                Record(mAcol.接收人).Value = Nvl(rsTmp("接收人"))
                Record(mAcol.接收时间).Value = Nvl(rsTmp("接收时间"))
                Record(mAcol.诊疗项目ID).Value = Nvl(rsTmp("诊疗项目ID"))
                Record(mAcol.执行状态).Value = Nvl(rsTmp("执行状态"))
                Record(mAcol.医嘱id).Value = Nvl(rsTmp("医嘱id"))
                Record(mAcol.送检时间).Value = Nvl(rsTmp("送检时间"))
                
                
                For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                    Record(intLoop).ForeColor = Val(Nvl(rsTmp("试管颜色")))
                Next
                
                If blOnlyWhere = True Then
                    If Record(mAcol.条码).Value <> "" And intShowButtom <> 2 Then
                        intShowButtom = 1
                    End If
                    If Record(mAcol.采样人).Value <> "" Then
                        intShowButtom = 2
                    End If
                End If
            Else
                Record(mAcol.医嘱内容).Value = Record(mAcol.医嘱内容).Value & " " & Nvl(rsTmp("医嘱内容"))
                Record(mAcol.合并医嘱).Value = Record(mAcol.合并医嘱).Value & ";" & _
                                               Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                Record(mAcol.别名).Value = Record(mAcol.别名).Value & " " & Nvl(rsTmp("别名"))
            End If
            strOldAdvice = rsTmp("相关ID") & ""
            strOldNO = rsTmp("试管编码") & ""
            strOldCodeBar = rsTmp("样本条码") & ""
            If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("试管编码")) & ",") <= 0 Then
                strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("试管编码"))
            End If
        End If
        If chkRemberPer.Value = 1 Then
            If Nvl(rsTmp("送检人") & "") <> "" Then
                '已有送检人。显示送检人
                txt送检人 = Nvl(rsTmp("送检人") & "")
            Else
                txt送检人 = mstrSendPerson
            End If
        Else
            txt送检人 = Nvl(rsTmp("送检人"))
        End If
        rsTmp.MoveNext
    Loop
    If Me.Visible = True Then
        Me.rptAlist(TabCtr.Selected.Index).Populate
    End If
    
    '有记录时表示写入成功
    If rptAlist(TabCtr.Selected.Index).Records.Count > 0 Then
        RefreshAdviceData = True
    End If
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        mlngKey = Nvl(rsTmp("病人ID"))
    End If
    '当使用病人查找时填写病人信息
    If blOnlyWhere = True Then
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            txt姓名 = IIf(IsNull(rsTmp("婴儿姓名")), Nvl(rsTmp("病人姓名")), Nvl(rsTmp("婴儿姓名")))
            txt姓名.Tag = IIf(IsNull(rsTmp("婴儿姓名")), Nvl(rsTmp("病人姓名")), Nvl(rsTmp("婴儿姓名")))
            On Error Resume Next
            txt性别 = IIf(IsNull(rsTmp("婴儿姓名")), Nvl(rsTmp("性别")), Nvl(rsTmp("婴儿性别")))
            txt年龄 = Nvl(rsTmp("年龄"))
            On Error GoTo 0
            txtBed = Nvl(rsTmp("床号"))
            txtID = Nvl(rsTmp("标识号"))
            txtPatientDept = Nvl(rsTmp("病人所在科室"))
            mlngKey = Nvl(rsTmp("病人ID"))
        End If
    Else
        '设置定义
        Select Case Me.TabCtr.Selected.Index
            Case 0
                
            Case 1
                
            Case 2
                
        End Select
    End If
    
    
    If strCuvetteNumber <> "" Then
        With Me.rptCuvette
            strSQL = "select 编码,名称,添加剂,采血量,规格,颜色 from 采血管类型 where 编码 in  " & _
                      "(Select * From Table(Cast(f_str2list([1]) As Zltools.t_strlist)))"
                       
                        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strCuvetteNumber, 2))
            Do Until rsTmp.EOF
                Set Record = .Records.Add
                For intLoop = 0 To .Columns.Count + 1
                    Record.AddItem ""
                Next
                Record(mCuvette.选择).HasCheckbox = True
                Record(mCuvette.选择).Checked = True
                Record(mCuvette.编码).Value = Nvl(rsTmp("编码"))
                Record(mCuvette.名称).Value = Nvl(rsTmp("名称"))
                Record(mCuvette.添加剂).Value = Nvl(rsTmp("添加剂"))
                Record(mCuvette.采血量).Value = Nvl(rsTmp("采血量"))
                Record(mCuvette.规格).Value = Nvl(rsTmp("规格"))
                Record(mCuvette.颜色).BackColor = Nvl(rsTmp("颜色"))
                
                For intLoop = 0 To .Columns.Count + 1
                    Record(intLoop).ForeColor = Nvl(rsTmp("颜色"))
                Next
                rsTmp.MoveNext
            Loop
            
            If InStr(1, Mid(strCuvetteNumber, 2), ",") <= 0 Then
                .Records(0).Item(mCuvette.选择).Checked = True
            End If
        End With
    End If
    
    Me.rptCuvette.Populate
    
    '显示当前查询的费用
    If Me.rptAlist(TabCtr.Selected.Index).Records.Count > 0 Then
        With Me.rptAlist(TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.费用).Value = "√" Then
                    str已收费医嘱ID = str已收费医嘱ID & "," & .Records(intLoop).Item(mAcol.ID).Value
                End If
                str所有医嘱ID = str所有医嘱ID & "," & .Records(intLoop).Item(mAcol.ID).Value
            Next
            
            If Mid(str所有医嘱ID, 2) <> "" Then
                gstrSql = "Select /*+ rule */ Sum(实收金额) As 已收金额" & vbNewLine & _
                            "From 住院费用记录" & vbNewLine & _
                            "Where 医嘱序号 In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist))) "
                            
                If intPatientType <> 2 Then
                    gstrSql = Replace(gstrSql, "住院费用记录", "门诊费用记录")
                End If
                                            
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str所有医嘱ID)
                lbl显示费用 = "(合计:应收<" & rsTmp("已收金额") & "元> "
            Else
                lbl显示费用 = ""
            End If
            If Mid(str已收费医嘱ID, 2) <> "" Then
                gstrSql = "Select /*+ rule */ Sum(实收金额) As 已收金额" & vbNewLine & _
                            "From 住院费用记录" & vbNewLine & _
                            "Where 医嘱序号 In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist))) "
                            
                If intPatientType <> 2 Then
                    gstrSql = Replace(gstrSql, "住院费用记录", "门诊费用记录")
                End If
                                            
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str已收费医嘱ID)
                
                lbl显示费用 = lbl显示费用 & "实收<" & rsTmp("已收金额") & "元>)"
            Else
                If lbl显示费用 <> "" Then
                    lbl显示费用 = lbl显示费用 & "实收<0元>)"
                End If
            End If
        End With
    Else
        lbl显示费用 = ""
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub rptAlist_ItemCheck(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim i  As Integer
    If Item.Record(mAcol.执行状态).Value = 1 Or Item.Record(mAcol.执行状态).Value = 3 Then
        Item.Record(mAcol.选择).Checked = False
        Me.rptAlist(TabCtr.Selected.Index).Populate
        MsgBox "已执行的标本不能选中!", vbInformation, Me.Caption
    End If
    If Index = 1 Or Index = 0 Then
        '选择，或取消选择同一个条码的都同时起作用
        If Item.Record(mAcol.条码).Value <> "" Then
            For i = 0 To rptAlist(Index).Rows.Count - 1
                If rptAlist(Index).Records(i).Item(i).Record(mAcol.条码).Value = Item.Record(mAcol.条码).Value Then
                    rptAlist(Index).Records(i).Item(i).Record(mAcol.选择).Checked = Item.Record(mAcol.选择).Checked
                End If
            Next
            rptAlist(Index).Redraw
        End If
    End If
End Sub

Private Sub rptAlist_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptAlist(TabCtr.Selected.Index)
        Set hitColumn = .HitTest(X, Y).Column
        
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                hitColumn.AutoSize = True
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mAcol.选择).Checked
                For Each Record In .Records
                    If Record.Item(mAcol.执行状态).Value = 0 Then
                        Record.Item(mAcol.选择).Checked = blSelect
                    Else
                        Record.Item(mAcol.选择).Checked = False
                    End If
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub rptAlist_SelectionChanged(Index As Integer)
    With Me.rptAlist(TabCtr.Selected.Index)
        .PaintManager.HighlightBackColor = .FocusedRow.Record(mAcol.试管颜色).Value
        .Populate
        If chkRemberPer.Value = 1 Then
            If .FocusedRow.Record(mAcol.送检人).Value <> "" Then
                '已经有送检人后，显示送检人
                txt送检人.Text = .FocusedRow.Record(mAcol.送检人).Value
            Else
                txt送检人.Text = mstrSendPerson
            End If
        Else
            txt送检人.Text = .FocusedRow.Record(mAcol.送检人).Value
        End If
    End With
    RePrintBarCode True
End Sub

Private Sub rptCount_SelectionChanged()
    With Me.rptCount
        txt送检人.Text = .FocusedRow.Record(mAcol.送检人).Value
    End With
End Sub

Private Sub rptCuvette_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
     '选中管码
    Call SelectCuvette
End Sub

Private Sub rptCuvette_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptCuvette
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mCuvette.选择).Checked
                For Each Record In .Records
                    Record.Item(mCuvette.选择).Checked = blSelect
                Next
            End If
        End If
        .Populate
        '选中管码
        Call SelectCuvette
    End With
End Sub

Private Sub rptCuvette_SelectionChanged()
    With Me.rptCuvette
        .PaintManager.HighlightBackColor = .FocusedRow.Record(mCuvette.颜色).ForeColor
        .Populate
    End With
End Sub

Private Sub rptPlist_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   '登记和取消登记
    If Me.TabCtr.Selected.Index <= 1 Then
        SaveRegister Me.TabCtr.Selected.Index
    End If
End Sub

Private Sub rptPlist_SelectionChanged()
    Dim intPatientType As Integer           '病人来源
    
    With Me.rptPlist.FocusedRow
        mlngKey = .Record(mPcol.病人ID).Value
        If .Record(mPcol.来源).Value = "门诊" Then
            intPatientType = 1
        ElseIf .Record(mPcol.来源).Value = "住院" Then
            intPatientType = 2
        ElseIf .Record(mPcol.来源).Value = "院外" Then
            intPatientType = 3
        ElseIf .Record(mPcol.来源).Value = "体检" Then
            intPatientType = 4
        Else
            intPatientType = 1
        End If
        mintBabyNo = Val(.Record(mPcol.婴儿号).Value)
    End With
'    lbl显示费用.Caption = ""
    '使用病人ID刷新医嘱
    RefreshAdviceData mlngKey, TabCtr.Selected.Index, intPatientType, False, mintBabyNo
    '刷新显示信息
    ShowPatientInfo
     Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
End Sub
Private Sub SelectCuvette()
    '功能               选择被选中的试管
    
    Dim RecordC As ReportRecord
    Dim RecordA As ReportRecord
    
    For Each RecordC In Me.rptCuvette.Records
        For Each RecordA In Me.rptAlist(TabCtr.Selected.Index).Records
            If RecordA(mAcol.试管编码).Value = RecordC(mCuvette.编码).Value And RecordA(mAcol.执行状态).Value = 0 Then
                RecordA(mAcol.选择).Checked = RecordC(mCuvette.选择).Checked
            End If
        Next
    Next

    Me.rptAlist(TabCtr.Selected.Index).Populate
End Sub
Private Function SelectBarCode(strBarCode As String) As Boolean
    Dim RowA As ReportRow
    
    For Each RowA In Me.rptAlist(TabCtr.Selected.Index).Rows
        If RowA.Record(mAcol.条码).Value = strBarCode Then
            RowA.Record(mAcol.选择).Checked = True
            SelectBarCode = True
        Else
            RowA.Record(mAcol.选择).Checked = False
        End If
    Next
    Me.rptAlist(TabCtr.Selected.Index).Populate
End Function

Private Sub ShowPatientInfo()
    
    '没有焦点行时退出
    If Me.rptPlist.FocusedRow Is Nothing Then Exit Sub
    On Error Resume Next
    With Me.rptPlist.FocusedRow
    
        
        txt姓名 = .Record(mPcol.病人姓名).Value
        txt姓名.Tag = .Record(mPcol.病人姓名).Value
        txt性别 = .Record(mPcol.性别).Value
        txt年龄 = .Record(mPcol.年龄).Value
        
        txtBed = .Record(mPcol.床号).Value
        txtID = .Record(mPcol.标识号).Value
        txtPatientDept = .Record(mPcol.病人科室).Value
    End With
End Sub

Private Function SaveRegister(intState As Integer) As Boolean
    '功能:              登记或取消登记
    '参数:              intState = 0 未登记 = 1 已登记 = 2 已执行
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strAdvice As String
    Dim intLoop As Integer
    Dim cbrControl As CommandBarControl
    Dim Record As ReportRecord
    Dim intTimeLimit As Integer         '送检时限单位分钟
    Dim blnTimeLimit As Boolean         '是否超过送检时限 true = 超过
    Dim str医嘱ID As String
    Dim strMsg As String
    Dim strUrgent As String
    Dim strTemp As String
    Dim strValue As String
    Dim intLen As Integer
    
    On Error GoTo errH
        
    Set cbrControl = Me.cbrthis.FindControl(, conMenu_Edit_Insert, True, True)
    
    If cbrControl.Caption = "开始计数" Then
        '没有开始计数时强行开始
        BeginRegister
    End If
    
    If Me.rptAlist(TabCtr.Selected.Index).Rows.Count <= 0 Then
        MsgBox "没有找到可以登记的医嘱记录!", vbQuestion, gstrSysName
        Exit Function
    End If
    
    If intState = 0 Then
        With Me.rptAlist(TabCtr.Selected.Index)
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mAcol.选择).Checked = True And .Rows(intLoop).Record(mAcol.执行状态).Value = 0 Then
                    str医嘱ID = str医嘱ID & "," & .Rows(intLoop).Record(mAcol.ID).Value & "," & .Rows(intLoop).Record(mAcol.合并医嘱).Value
                    
                    '紧急标本信息
                    If .Rows(intLoop).Record(mAcol.紧急).Value = "紧急" And chkUrgent.Value = 1 Then
                        strUrgent = strUrgent & "," & .Rows(intLoop).Record(mAcol.医嘱内容).Value
                    End If
                End If
            Next
        End With
        str医嘱ID = Mid(str医嘱ID, 2)
        If Chk划价费用(Me, str医嘱ID, 0) = False Then
            Exit Function
        End If
    End If
    With Me.rptAlist(TabCtr.Selected.Index)
        
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mAcol.选择).Checked = True And .Rows(intLoop).Record(mAcol.执行状态).Value = 0 Then
                '处理是否超过采集时限
                gstrSql = "select 送检时限 from 检验项目选项 where 诊疗项目id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(.Rows(intLoop).Record(mAcol.诊疗项目ID).Value))
                If rsTmp.EOF = True Then
                    intTimeLimit = 0
                Else
                    intTimeLimit = Val(Nvl(rsTmp("送检时限")))
                End If
                
                If IsDate(.Rows(intLoop).Record(mAcol.采样时间).Value) = False And intTimeLimit > 0 Then
                    blnTimeLimit = True
                Else
                    If IsDate(.Rows(intLoop).Record(mAcol.采样时间).Value) = True Then
                        If DateDiff("n", .Rows(intLoop).Record(mAcol.采样时间).Value, zlDatabase.Currentdate) > intTimeLimit _
                            And intTimeLimit > 0 Then
                            '超过送检时限
                            blnTimeLimit = True
                        End If
                    Else
                        If intTimeLimit > 0 Then
                            blnTimeLimit = True
                        End If
                    End If
                End If
                If chkComRequest.Value = 1 And .Rows(intLoop).Record(mAcol.送检时间).Value = "" Then
                    strMsg = strMsg & "当前《" & .Rows(intLoop).Record(mAcol.医嘱内容).Value & "》未送检,不允许登记！" & vbNewLine
                Else
                    If blnTimeLimit = True And (intState = 0 Or intState = 4) Then
                        '超时处理，查看是否有权限，有权限时只提示
                        If InStr(mstrPrivs, "强制通过送检时限") > 0 Then
                            '提示
                            MsgBox ("当前标本的采样时间为《" & .Rows(intLoop).Record(mAcol.采样时间).Value & "》" & vbCrLf & _
                                    "已超过采样时限" & intTimeLimit & "分钟,送检延迟！")
                            strAdvice = strAdvice & "|" & .Rows(intLoop).Record(mAcol.ID).Value & _
                                Replace(.Rows(intLoop).Record(mAcol.合并医嘱).Value, ";", "|")
                        Else
                            '拒绝登记
                            MsgBox ("当前标本的采样时间为《" & .Rows(intLoop).Record(mAcol.采样时间).Value & "》" & vbCrLf & _
                                    "已超过采样时限" & intTimeLimit & "分钟,不允许登记！")
                        End If
                        
                    ElseIf .Rows(intLoop).Record(mAcol.采样时间).Value = "" And intState = 0 Then
                        '处理强制登记未采样标本
                        If InStr(mstrPrivs, "强制登记未采样标本") > 0 Then
                            strAdvice = strAdvice & "|" & .Rows(intLoop).Record(mAcol.ID).Value & _
                                Replace(.Rows(intLoop).Record(mAcol.合并医嘱).Value, ";", "|")
                        Else
                            '拒绝登记
                            MsgBox "当前《" & .Rows(intLoop).Record(mAcol.医嘱内容).Value & "》未采样,不允许登记！", vbInformation
                        End If
                    Else
                        strAdvice = strAdvice & "|" & .Rows(intLoop).Record(mAcol.ID).Value & _
                                Replace(.Rows(intLoop).Record(mAcol.合并医嘱).Value, ";", "|")
                    End If
                End If
            End If
        Next
        If strMsg <> "" Then MsgBox strMsg, vbInformation, "标本签收"
    End With
    If strAdvice <> "" Then
    
        If intState = 0 Then
            '检查tat超时
            If getTATTime(strAdvice) = False Then
                Exit Function
            End If
            If strAdvice = "" Then
                Exit Function
            End If
 
        End If
        
        '提示紧急标本
        If strUrgent <> "" And chkUrgent.Value = 1 Then
            If UBound(Split(strUrgent, ",")) > 2 Then
                MsgBox "【" & Split(strUrgent, ",")(1) & "," & Split(strUrgent, ",")(2) & ",......】登记标本为紧急标本！", vbInformation, "标本登记"
            Else
                MsgBox "【" & Mid(strUrgent, 2) & "】登记标本为紧急标本！", vbInformation, "标本登记"
            End If
        End If
        
        '登记和取消登记
        If Len(Mid(strAdvice, 2)) > 2000 Then
            strTemp = Mid(strAdvice, 2)
            Do While Len(strTemp) > 2000
                strValue = Mid(strTemp, 1, 2000)
                intLen = InStrRev(strValue, "|")
                strTemp = Mid(strValue, intLen + 1) & Mid(strTemp, 2001)
                strValue = Mid(strValue, 1, intLen - 1)
                
                strSQL = "Zl_病人医嘱发送_SampleInput('" & strValue
                If intState = 0 Or intState = 3 Or intState = 4 Then
                    strSQL = strSQL & "','" & UserInfo.姓名 & "'," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & Trim(txt送检人.Text) & "')"
                ElseIf intState = 1 Then
                    strSQL = strSQL & "',NULL," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                End If
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
            Loop
            
            If strTemp <> "" Then
                strSQL = "Zl_病人医嘱发送_SampleInput('" & strTemp
                If intState = 0 Or intState = 3 Or intState = 4 Then
                    strSQL = strSQL & "','" & UserInfo.姓名 & "'," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & Trim(txt送检人.Text) & "')"
                ElseIf intState = 1 Then
                    strSQL = strSQL & "',NULL," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                End If
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
            End If
        Else
            strSQL = "Zl_病人医嘱发送_SampleInput('" & Mid(strAdvice, 2)
            If intState = 0 Or intState = 3 Or intState = 4 Then
                strSQL = strSQL & "','" & UserInfo.姓名 & "'," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & Trim(txt送检人.Text) & "')"
            ElseIf intState = 1 Then
                strSQL = strSQL & "',NULL," & mlngBatch & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            End If
            zlDatabase.ExecuteProcedure strSQL, gstrSysName
        End If
        
        SaveRegister = True
        mblnUse = True
        
        If intState = 0 Or intState = 3 Or intState = 4 Then
            Call WriterCheckSampleToLIS(Mid(strAdvice, 2), UserInfo.姓名, mlngBatch, Trim(Me.txt送检人.Text))
        ElseIf intState = 1 Then
            Call WriterCheckSampleToLIS(Mid(strAdvice, 2), "", 0)
        End If
        
        '没有医嘱记录时退出
        If Me.rptAlist(TabCtr.Selected.Index).Rows.Count = 0 Then Exit Function
        
        If intState = 0 Or intState = 4 Then
            '增加
            Call InsrOrDelAdvice(1, Replace(Mid(strAdvice, 2), "|", ","))
        Else
            '取消
            Call InsrOrDelAdvice(0, Replace(Mid(strAdvice, 2), "|", ","))
        End If
    End If
    
    Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
    Me.rptCuvette.Records.DeleteAll
    Me.rptAlist(TabCtr.Selected.Index).Populate
    Me.rptCuvette.Populate
    txt姓名.Text = ""
    txt性别.Text = ""
    txt年龄.Text = ""
    txtBed.Text = ""
    txtID.Text = ""
    txtPatientDept.Text = ""
    
    SaveRegister = True
    
    Exit Function
errH:
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub TabCtr_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intPatientType As Integer           '病人来源
    
    If Me.Visible = False Then Exit Sub
    Dim Controlcbo As CommandBarComboBox                '批次控件
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_File_RoomSet, True, True)
    Select Case Item.Index
        Case 0
            Me.cmdOK.Enabled = True
            Me.cmdOK.Caption = "登记(&G)"
            Controlcbo.Clear
            Controlcbo.AddItem "所有批次"
            Controlcbo.ItemData(Controlcbo.ListCount) = 0
            Controlcbo.ListIndex = 1
            mlngSelectBatch = 0
        Case 1
            Me.cmdOK.Enabled = True
            Me.cmdOK.Caption = "取消(&G)"
        Case 2
            Me.cmdOK.Enabled = False
            Me.cmdOK.Caption = "取消(&G)"
        Case 3
            Me.cmdOK.Enabled = False
            Me.cmdOK.Caption = "取消(&G)"
        Case 4
            Me.cmdOK.Enabled = False
            Me.cmdOK.Caption = "登记(&G)"
    End Select
    RefreshPatientData 1, mintBabyNo
    Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
    Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptPlist.Rows.Count & "个病人！"
End Sub

Private Sub txtGoto_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
End Sub

Private Sub txtGoto_GotFocus()
    txtGoto.SelStart = 0
    txtGoto.SelLength = Len(txtGoto.Text)
    If Not mobjIDCard Is Nothing And txtGoto.Text = "" And Not txtGoto.Locked Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtGoto_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blFind As Boolean                           '是否查找成功
    Dim Row As ReportRow                            '列表行对象
    Dim strFind As String                           '临进查找字串
    Dim blnCard As Boolean
    Dim str已收费医嘱ID As String                   '已收费的医嘱ID已","分隔
    Dim str所有医嘱ID As String                     '所有的医嘱ID已","分隔
    Dim intLoop As Integer
    Dim lng卡类别ID As Long
    Dim lng病人ID As Long
    
    On Error GoTo errH
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'‘’;；:：?？|,，。""") = True Then KeyAscii = 0
    
    If IDKind.IDKind = IDKinds.C0姓名 Then
'        blnCard = zlCommFun.InputIsCard(txtGoto, KeyAscii, False)
    End If
    
    blnCard = False
    
    If IDKind.IDKind = IDKinds.C5就诊卡 Then
'        Call zlCommFun.InputIsCard(txtGoto, KeyAscii, True)
        gbytCardNOLen = Val(IDKind.GetKindItem("卡号长度", IDKind.IDKind))
        blnCard = KeyAscii <> 8 And Len(txtGoto.Text) = gbytCardNOLen - 1 And txtGoto.SelLength <> Len(txtGoto.Text)
        If blnCard = True Then
            If KeyAscii <> 13 Then
                Me.txtGoto = Me.txtGoto & Chr(KeyAscii)
            End If
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = 13 Or (IDKind.IDKind = IDKinds.C5就诊卡 And blnCard = True) Then
    
        '先清空再定入
        
        If mstrFirstBarCode <> txtGoto.Text Then
            txt姓名 = ""
            txt姓名.Tag = ""
            txt性别 = ""
            txt年龄 = ""
            txtBed = ""
            txtID = ""
            lbl显示费用.Caption = ""
            txtPatientDept = ""
            Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
            Me.rptCuvette.Records.DeleteAll
            Me.rptAlist(TabCtr.Selected.Index).Populate
            Me.rptCuvette.Populate
        End If
        
        Select Case Mid(txtGoto, 1, 1)
            Case "-"                                '病人ID
                blFind = RefreshAdviceData(Mid(txtGoto, 2), Me.TabCtr.Selected.Index, 0, True)
                strFind = Val(Mid(txtGoto, 2))
            Case "+"                                '住院号
                strSQL = "select 病人ID from 病人信息 where 住院号 = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 2, True)
                strFind = Mid(txtGoto, 2)
            Case "*"                                '门诊号
                strSQL = "select 病人ID from 病人信息 where 门诊号 = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Mid(txtGoto, 2)
            Case "."                                '挂号单号
                strSQL = "select 病人ID from 病人医嘱记录　where 挂号单 = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Mid(txtGoto, 2))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Mid(txtGoto, 2)
            Case "/"                                '收费单据号
                strSQL = "select distinct 病人ID from 门诊费用记录 where No = [1] and 病人id is not null and 门诊标志 = 1 order by 病人ID desc"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, zlCommFun.GetFullNO(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Mid(txtGoto, 2)
            Case Else                               '就诊卡和姓名
                strFind = txtGoto
                If IDKind.IDKind = IDKinds.C0姓名 And BlnIsNumber(txtGoto) Then
                    strSQL = "select a.病人id,a.病人来源 from 病人医嘱记录 a , 病人医嘱发送 b " & _
                         " Where a.ID = b.医嘱id And b.样本条码 = [1]  order by a.开嘱时间 desc     "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, txtGoto)
                    If rsTmp.EOF = False Then
                        If mstrFirstBarCode = txtGoto.Text Then
                            mstrFirstBarCode = ""
                            '调用登记
                            Call LocationObj(txtGoto)
                            Call cmdOK_Click
                            
                            Exit Sub
                        Else
                            mstrFirstBarCode = txtGoto.Text
                            
                        End If
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, Nvl(rsTmp(1), 0), True)
                    End If
                Else
                    If blnCard Or IDKind.IDKind = IDKinds.C5就诊卡 Then
                        strSQL = "select 病人ID from 病人信息 where 就诊卡号 = [1] "
                        strFind = UCase(txtGoto)
                    ElseIf IDKind.IDKind = IDKinds.C0姓名 Then
                        strSQL = "select 病人ID from 病人信息 where 姓名 = [1] "
                    ElseIf IDKind.IDKind = IDKinds.C1医保号 Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    ElseIf IDKind.IDKind = IDKinds.C2身份证号 Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    ElseIf IDKind.IDKind = IDKinds.C3IC卡号 Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    ElseIf IDKind.IDKind = IDKinds.C4门诊号 Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    Else
                        If Val(IDKind.GetKindItem("卡类别ID")) <> 0 Then
                            lng卡类别ID = Val(IDKind.GetKindItem("卡类别ID"))
                            If mobjSquareCard.zlGetPatiID(lng卡类别ID, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                            If lng病人ID = 0 Then lng病人ID = 0
                        Else
                            If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        End If
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        strFind = lng病人ID
                    End If
                  
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind)
                    If rsTmp.EOF = False Then
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                    End If
                End If
        End Select
        '不是条码则情况第一次条码
        If BlnIsNumber(txtGoto) = False Then
            mstrFirstBarCode = ""
        End If
        
        '没有找到条码时给提示信息
        If blFind = False Then
            If IDKind.IDKind = IDKinds.C0姓名 And BlnIsNumber(txtGoto) Then
                '有条码时判断一下条码的状态
                gstrSql = " Select b.执行状态, b.采样人, b.采样时间, b.接收人, b.接收时间, b.标本送出时间 From 病人医嘱记录 a, 病人医嘱发送 b " & _
                         " Where a.id = b.医嘱id and a.相关id is not null and  b.样本条码 = [1] order by  a.开嘱时间 desc  "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.txtGoto)
                If rsTmp.EOF = True Then
                    MsgBox "没有找到条码<" & Me.txtGoto & ">!" & vbCrLf & _
                        "条码可能被取消绑定或者医嘱被作废!", vbInformation, Me.Caption
                Else
                    If rsTmp("执行状态") = 1 Or rsTmp("执行状态") = 3 Then
                        MsgBox "样本条码为<" & Me.txtGoto & ">已被核收!"
                    Else
                        If rsTmp("接收人") <> "" Then
                            MsgBox "样本条码<" & Me.txtGoto & ">已登记  " & vbCrLf & _
                                  "登记时间<" & rsTmp("接收时间") & ">" & vbCrLf & _
                                  "登记人<" & rsTmp("接收人") & ">"
                        End If
                    End If
                End If
                
            End If
        Else
            If Me.TabCtr.Selected.Index = 4 Then
                Me.cmdOK.Enabled = True
            End If

        End If
        
        If Me.rptAlist(TabCtr.Selected.Index).Rows.Count > 0 And Me.cmdOK.Enabled = True Then
            If ChkBarCodeRegister.Value = 1 Then
                mstrFirstBarCode = ""
                '调用登记
                Call LocationObj(txtGoto)
                Call cmdOK_Click
            Else
                cmdOK.SetFocus
            End If
        Else
            Me.txtGoto.Text = ""
            Me.txtGoto.SetFocus
        End If
        If mstrFirstBarCode <> "" Then
            Call LocationObj(txtGoto)
        End If
        On Error Resume Next
        '下面用于定位,如出错就忽略
        
'        If blFind = True Then
'            If Mid(txtGoto, 1, 1) <> "-" Then
'                rsTmp.MoveFirst
'                For Each Row In Me.rptPlist(0).Rows
'                    If Row.Record(mPcol.病人ID).Value = Nvl(rsTmp(0)) Then
'                        Me.rptPlist(0).FocusedRow = Row
'                        Me.rptPlist(0).Populate
'                    End If
'                Next
'            Else
'                For Each Row In Me.rptPlist(0).Rows
'                    If Row.Record(mPcol.病人ID).Value = Mid(txtGoto, 2) Then
'                        Me.rptPlist(0).FocusedRow = Row
'                        Me.rptPlist(0).Populate
'                    End If
'                Next
'            End If
'            Me.txtGoto.Text = ""
'            Me.txtGoto.SetFocus
'        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptPlist.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptPlist) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = "病人接收清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & zlDatabase.Currentdate)
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub BeginRegister()
    '功能           开始登记标本，并生成新的批号
   
    Dim rsTmp As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = Me.cbrthis.FindControl(, conMenu_Edit_Insert, True, True)
    
    If cbrControl.Caption = "开始计数" Then
        If Me.rptCount.Records.Count > 0 Then
            If MsgBox("是否停止当前计数开始新的计数？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
        cbrControl.Caption = "结束计数"
        
        '没有使用批次时不增加
        If mblnUse = True Or mlngBatch = 0 Then
            '得到一个新的批次
            gstrSql = "select 病人医嘱发送_接收批次.nextval from dual "
            zlDatabase.OpenRecordset rsTmp, gstrSql, Me.Caption
            mlngBatch = rsTmp(0)
            mblnUse = False
        End If
        Me.txt送检人.Enabled = True
    Else
        cbrControl.Caption = "开始计数"
        Me.txt送检人.Enabled = False
        If Me.rptCount.Records.Count > 0 Then
            If MsgBox("是否打印当前计数完成的清单?", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                RegisterLisPrint (1)
            End If
        End If
        
    End If
    If chkRemberPer.Value = 1 Then
        If Nvl(mstrSendPerson) <> "" Then
            txt送检人 = mstrSendPerson
        Else
            txt送检人 = ""
        End If
    Else
        Me.txt送检人.Text = ""
    End If
    Me.cbrthis.RecalcLayout
    Me.rptCount.Records.DeleteAll
    Me.rptCount.Populate

End Sub
Private Sub GetBatch()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String                                '用于SQL语句存放
    Dim strSQL1 As String                               '用于H表的查找
    Dim Record As ReportRecord
    Dim Item As ReportColumn
    Dim intLoop As Integer
    Dim strTmp As String                                '临时字串变量
    Dim varFilter As Variant                            '过滤字串
    Dim strDateBegin As Date                            '开始时间
    Dim strDateEnd As Date                              '结束时间
    Dim blnDateMoved As Boolean                         '是否被转出
    Dim lngPatientID As Long                            '病人ID
    Dim strState As String                              '状态
    Dim Controlcbo As CommandBarComboBox                '批次控件
    
    On Error GoTo errH
    
    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("标本登记过滤", 100, 1212, "")
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_File_RoomSet, True, True)
    
    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    strSQL = "Select Distinct b.接收批次" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 诊疗项目目录 F" & vbNewLine & _
            "Where a.Id = b.医嘱id And a.病人id = c.病人id And a.诊疗项目id = f.Id And a.诊疗类别 = 'E' And f.操作类型 = '6' And" & vbNewLine & _
            "      b.接收人 Is Not Null And b.执行部门id = [1]"
    
    If Me.rptPlist.Tag <> "" Then
        If Val(varFilter(mFilter.住院)) = 0 Then
            strSQL = strSQL & " And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3,", "") & _
                     Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "
        Else
            strSQL = strSQL & " and c.出院时间 Is Null And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3,", "") & _
                     Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "
        End If
        
        If varFilter(mFilter.就诊卡) <> "" Then
            strSQL = strSQL & " And C.就诊卡号 = [3] "
        End If
        
        If varFilter(mFilter.姓名) <> "" Then
            strSQL = strSQL & " And C.姓名 like [4] "
        End If
        
        If varFilter(mFilter.单据号) <> "" Then
            strSQL = strSQL & " and B.NO = [5]"
        End If
        
        If varFilter(mFilter.标本) <> "所有标本" Then
            strSQL = strSQL & " and A.标本部位 = [6] "
        End If
        
        If varFilter(mFilter.采集方式) <> 0 Then
            strSQL = strSQL & " and f.ID = [7] "
        End If
        
        If varFilter(mFilter.病人科室) <> 0 Then
            strSQL = strSQL & " And a.病人科室ID = [8] "
        End If
        
        strSQL = strSQL & " and b.发送时间 Between [9] and [10]"
        
        If varFilter(mFilter.开始时间) = "" Then
            strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.间隔时间))
            strDateEnd = zlDatabase.Currentdate
        Else
            strDateBegin = varFilter(mFilter.开始时间)
            strDateEnd = varFilter(mFilter.结束时间)
        End If
    Else
        If strTmp <> "" Then
            strSQL = strSQL & " And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3,", "") & _
                 Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "
            
            If varFilter(mFilter.标本) <> "所有标本" Then
                strSQL = strSQL & " and A.标本部位 = [6] "
            End If
            
            If varFilter(mFilter.采集方式) <> 0 Then
                strSQL = strSQL & " and f.ID = [7] "
            End If
            
            If varFilter(mFilter.病人科室) <> 0 Then
                strSQL = strSQL & " And a.病人科室ID = [8] "
            End If
            
            strSQL = strSQL & " and b.发送时间 Between [9] and [10]"
            
            If varFilter(mFilter.开始时间) = "" Then
                strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.间隔时间))
                strDateEnd = zlDatabase.Currentdate
            Else
                strDateBegin = varFilter(mFilter.开始时间)
                strDateEnd = varFilter(mFilter.结束时间)
            End If
        Else
            strSQL = strSQL & " and b.发送时间 Between [9] and [10]"
            strDateBegin = zlDatabase.Currentdate - 3
            strDateEnd = zlDatabase.Currentdate
        End If
    End If
    
    blnDateMoved = MovedByDate(CDate(strDateBegin)) '按时间看是否可能已转出
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    strSQL = strSQL & " Order by b.接收批次 "
    
    If strTmp = "" And Me.rptPlist.Tag = "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, "", "", "", "", "", "", "", _
                    CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, Val(varFilter(mFilter.标识号)), CStr(varFilter(mFilter.就诊卡)) _
                    , CStr(varFilter(mFilter.姓名)) & "%", CStr(varFilter(mFilter.单据号)), CStr(varFilter(mFilter.标本)), CLng(varFilter(mFilter.采集方式)) _
                    , varFilter(mFilter.病人科室), CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), _
                    CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")))
    End If
    
    intLoop = 1
    Controlcbo.Clear
    Controlcbo.AddItem "所有批次"
    Controlcbo.ItemData(Controlcbo.ListCount) = 0
    
    Do While Not rsTmp.EOF
        Controlcbo.AddItem "第" & intLoop & "批次"
        Controlcbo.ItemData(Controlcbo.ListCount) = Val(Nvl(rsTmp("接收批次")))
        If Val(Nvl(rsTmp("接收批次"))) = mlngSelectBatch Then
            Controlcbo.ListIndex = Controlcbo.ListCount
        End If
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    If Controlcbo.ListIndex < 1 Then
        Controlcbo.ListIndex = 1
        mlngSelectBatch = 0
    End If
    Exit Sub
errH:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RegisterLisPrint(Mode As Integer)
    '功能       打印登记清单
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1212_1", Me, "接收批次=" & IIf(mlngSelectBatch = 0, mlngBatch, mlngSelectBatch), Mode)
End Sub

Private Sub IdKindChange()
    If Me.ActiveControl Is txtGoto Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5就诊卡, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub InsrOrDelAdvice(intType As Integer, strAdvice As String)
    '功能       增加和删除记数的医嘱
    '参数       intType = 1 增加医嘱 0 = 删除当前医嘱
    '           医嘱ID字串
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim aItem() As String
    Dim RecordA As ReportRecord
    Dim strOldAdvice As String
    Dim Record As ReportRecord
    Dim strCuvetteNumber As String
    Dim strBarCode As String
    Dim intRow As Integer
    Dim strSQLbak As String
    Dim intPatientType As Integer                                             '病人来源
    Dim strTemp As String
    Dim strValue As String
    Dim intLen As Integer
    Dim strSqlAll As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        '删除
        For intLoop = Me.rptCount.Records.Count - 1 To 0 Step -1
            With Me.rptCount.Records(intLoop)
                If InStr("," & strAdvice & ",", "," & .Item(mAcol.医嘱id).Value & ",") > 0 Then
                    Call Me.rptCount.Records.RemoveAt(intLoop)
                End If
            End With
        Next
        Me.rptCount.Populate
    End If
    
    If intType = 1 Then
        '增加
        If Len(strAdvice) > 4000 Then
            strTemp = strAdvice
            Do While Len(strTemp) > 4000
                strValue = Mid(strTemp, 1, 4000)
                intLen = InStrRev(strValue, ",")
                strTemp = Mid(strValue, intLen + 1) & Mid(strTemp, 4001)
                strValue = Mid(strValue, 1, intLen - 1)
                
                strSQL = " Select /*+ rule */ B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色, D.名称 As 采集方式, B.医嘱内容, C.样本条码,C.采样时间,c.标本送出时间 as 送检时间, " & vbCrLf & _
                         " H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
                         " I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号) as 标识号, " & vbCrLf & _
                         " L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人Id,c.采样人,c.送检人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
                         " DECODE(B.紧急标志,1,'紧急','') as 紧急,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源,b.婴儿,N.名称 as 别名,J.名称 as 病人科室,C.接收时间,C.接收人, " & vbCrLf & _
                         " b.诊疗项目ID,C.执行状态,O.记录性质,O.记录状态 " & vbCrLf & _
                         " From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
                         " 采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,部门表 J, " & vbCrLf & _
                         " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 O " & vbCrLf & _
                         " Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
                         " And E.类别 = 'C' And E.试管编码 = G.编码 And m.执行部门id = H.ID " & vbCrLf & _
                         " And D.类别 = 'E' And D.操作类型 = '6' And  " & vbCrLf & _
                         " B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & _
                         " And a.id  = m.医嘱id And E.id = N.诊疗项目ID(+) And I.当前科室id = J.id(+) And 出院时间 is null " & vbCrLf & _
                         " and c.医嘱id = O.医嘱序号(+) and c.记录性质 = mod(O.记录性质(+),10) and nvl(O.记录状态,0) in (0,1) " & vbCrLf & _
                         " and b.Id in (Select * From Table(Cast(f_Num2list('" & strValue & "') As zlTools.t_Numlist))) "
                strSqlAll = strSqlAll & strSQL & " union "
            Loop
            
            strSQL = " select /*+ rule */ 病人来源 from 病人医嘱记录 where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue)
            If rsTmp.EOF = True Then Exit Sub
            intPatientType = Nvl(rsTmp("病人来源"), 0)
            
            strSQL = " Select /*+ rule */ B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色, D.名称 As 采集方式, B.医嘱内容, C.样本条码,C.采样时间,c.标本送出时间 as 送检时间, " & vbCrLf & _
                     " H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
                     " I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号) as 标识号, " & vbCrLf & _
                     " L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人Id,c.采样人,c.送检人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
                     " DECODE(B.紧急标志,1,'紧急','') as 紧急,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源,b.婴儿,N.名称 as 别名,J.名称 as 病人科室,C.接收时间,C.接收人, " & vbCrLf & _
                     " b.诊疗项目ID,C.执行状态,O.记录性质,O.记录状态 " & vbCrLf & _
                     " From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
                     " 采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,部门表 J, " & vbCrLf & _
                     " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 O " & vbCrLf & _
                     " Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
                     " And E.类别 = 'C' And E.试管编码 = G.编码 And m.执行部门id = H.ID " & vbCrLf & _
                     " And D.类别 = 'E' And D.操作类型 = '6' And  " & vbCrLf & _
                     " B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & _
                     " And a.id  = m.医嘱id And E.id = N.诊疗项目ID(+) And I.当前科室id = J.id(+) And 出院时间 is null " & vbCrLf & _
                     " and c.医嘱id = O.医嘱序号(+) and c.记录性质 = mod(O.记录性质(+),10) and nvl(O.记录状态,0) in (0,1) " & vbCrLf & _
                     " and b.Id in (Select * From Table(Cast(f_Num2list('" & strTemp & "') As zlTools.t_Numlist))) "
            strSqlAll = strSqlAll & strSQL
            
            If intPatientType <> 2 Then
                strSqlAll = Replace(strSqlAll, "住院费用记录", "门诊费用记录")
            End If
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSqlAll, Me.Caption)
            
            Do Until rsTmp.EOF
                '没有对应颜色编码的采集不写入
                If IsNull(rsTmp("试管颜色")) = False Then
                    
                    If strOldAdvice <> rsTmp("相关ID") Then
                    
                        Set Record = Me.rptCount.Records.Add
                        For intLoop = 0 To rptCount.Columns.Count + 1
                            Record.AddItem ""
                        Next
                        
                        If (Nvl(rsTmp("执行状态")) = 1 Or Nvl(rsTmp("执行状态")) = 3) Then
                            Me.rptCount.Columns(mAcol.已执行).Visible = True
                        End If
                        
                        If Nvl(rsTmp("执行状态")) = 1 Or Nvl(rsTmp("执行状态")) = 3 Then
                            Record(mAcol.已执行).Value = "√"
                        Else
                            Record(mAcol.已执行).Value = ""
                        End If
                        Record(mAcol.费用).Value = IIf(rsTmp("记录状态") = 1, "√", "×")
                        Record(mAcol.ID).Value = Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                        Record(mAcol.选择).HasCheckbox = True
                        Record(mAcol.选择).Checked = IIf(Nvl(rsTmp("执行状态")) = 0, True, False)
                        Record(mAcol.图标).BackColor = Val(Nvl(rsTmp("试管颜色")))
                        Record(mAcol.采集方式).Value = Nvl(rsTmp("采集方式"))
                        Record(mAcol.医嘱内容).Value = Nvl(rsTmp("医嘱内容"))
                        Record(mAcol.条码).Value = Nvl(rsTmp("样本条码"))
                        Record(mAcol.执行科室).Value = Nvl(rsTmp("执行科室"))
                        Record(mAcol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
                        Record(mAcol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
                        Record(mAcol.发送人).Value = Nvl(rsTmp("发送人"))
                        Record(mAcol.发送时间).Value = Nvl(rsTmp("发送时间"))
                        Record(mAcol.试管颜色).Value = Nvl(rsTmp("试管颜色"))
                        Record(mAcol.试管编码).Value = Nvl(rsTmp("试管编码"))
                        Record(mAcol.标本).Value = Nvl(rsTmp("标本")) & IIf(Nvl(rsTmp("婴儿")) = 0, "", "(婴儿)")
                        Record(mAcol.采样时间).Value = Nvl(rsTmp("采样时间"))
                        Record(mAcol.采样人).Value = Nvl(rsTmp("采样人"))
                        Record(mAcol.送检人).Value = Nvl(rsTmp("送检人"))
                        Record(mAcol.采血量).Value = Nvl(rsTmp("采血量"))
                        Record(mAcol.试管名称).Value = Nvl(rsTmp("试管名称"))
                        Record(mAcol.紧急).Value = Nvl(rsTmp("紧急"))
                        Record(mAcol.病人来源).Value = Nvl(rsTmp("病人来源"))
                        Record(mAcol.婴儿).Value = Nvl(rsTmp("婴儿"))
                        Record(mAcol.别名).Value = Nvl(rsTmp("别名"))
                        Record(mAcol.相关ID).Value = Nvl(rsTmp("相关ID"))
                        
                        Record(mAcol.病人ID).Value = Nvl(rsTmp("病人ID"))
                        Record(mAcol.姓名).Value = Nvl(rsTmp("病人姓名"))
                        Record(mAcol.性别).Value = Nvl(rsTmp("性别"))
                        Record(mAcol.年龄).Value = Nvl(rsTmp("年龄"))
                        Record(mAcol.标识号).Value = Nvl(rsTmp("标识号"))
                        Record(mAcol.床号).Value = Nvl(rsTmp("床号"))
                        Record(mAcol.病人科室).Value = Nvl(rsTmp("病人科室"))
                        Record(mAcol.接收人).Value = Nvl(rsTmp("接收人"))
                        Record(mAcol.接收时间).Value = Nvl(rsTmp("接收时间"))
                        Record(mAcol.诊疗项目ID).Value = Nvl(rsTmp("诊疗项目ID"))
                        Record(mAcol.执行状态).Value = Nvl(rsTmp("执行状态"))
                        Record(mAcol.医嘱id).Value = Nvl(rsTmp("医嘱id"))
                        Record(mAcol.送检时间).Value = Nvl(rsTmp("送检时间"))
                        
                        For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                            Record(intLoop).ForeColor = Val(Nvl(rsTmp("试管颜色")))
                        Next
                    Else
                        If InStr(Record(mAcol.医嘱内容).Value, Nvl(rsTmp("医嘱内容"))) <= 0 Then
                            Record(mAcol.医嘱内容).Value = Record(mAcol.医嘱内容).Value & " " & Nvl(rsTmp("医嘱内容"))
                        End If
                        Record(mAcol.合并医嘱).Value = Record(mAcol.合并医嘱).Value & ";" & _
                                                       Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                        Record(mAcol.别名).Value = Record(mAcol.别名).Value & " " & Nvl(rsTmp("别名"))
                    End If
                    strOldAdvice = rsTmp("相关ID")
                    If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("试管编码")) & ",") <= 0 Then
                        strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("试管编码"))
                    End If
                End If
                If chkRemberPer.Value = 1 Then
                    If Nvl(rsTmp("送检人") & "") <> "" Then
                        txt送检人 = Nvl(rsTmp("送检人") & "")
                    Else
                        txt送检人 = mstrSendPerson
                    End If
                Else
                    txt送检人 = Nvl(rsTmp("送检人"))
                End If
                rsTmp.MoveNext
            Loop
        Else
            strSQL = " select /*+ rule */ 病人来源 from 病人医嘱记录 where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAdvice)
            If rsTmp.EOF = True Then Exit Sub
            intPatientType = Nvl(rsTmp("病人来源"), 0)
            
            strSQL = " Select /*+ rule */ B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色, D.名称 As 采集方式, B.医嘱内容, C.样本条码,C.采样时间,c.标本送出时间 as 送检时间, " & vbCrLf & _
                     " H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
                     " I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号) as 标识号, " & vbCrLf & _
                     " L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人Id,c.采样人,c.送检人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
                     " DECODE(B.紧急标志,1,'紧急','') as 紧急,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源,b.婴儿,N.名称 as 别名,J.名称 as 病人科室,C.接收时间,C.接收人, " & vbCrLf & _
                     " b.诊疗项目ID,C.执行状态,O.记录性质,O.记录状态 " & vbCrLf & _
                     " From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
                     " 采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,部门表 J, " & vbCrLf & _
                     " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 O " & vbCrLf & _
                     " Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
                     " And E.类别 = 'C' And E.试管编码 = G.编码 And m.执行部门id = H.ID " & vbCrLf & _
                     " And D.类别 = 'E' And D.操作类型 = '6' And  " & vbCrLf & _
                     " B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & _
                     " And a.id  = m.医嘱id And E.id = N.诊疗项目ID(+) And I.当前科室id = J.id(+) And 出院时间 is null " & vbCrLf & _
                     " and c.医嘱id = O.医嘱序号(+) and c.记录性质 = mod(O.记录性质(+),10) and nvl(O.记录状态,0) in (0,1) " & vbCrLf & _
                     " and b.Id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))  "
                     
            If intPatientType <> 2 Then
                strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
            End If
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAdvice)
            
            Do Until rsTmp.EOF
                '没有对应颜色编码的采集不写入
                If IsNull(rsTmp("试管颜色")) = False Then
                    
                    If strOldAdvice <> rsTmp("相关ID") Then
                    
                        Set Record = Me.rptCount.Records.Add
                        For intLoop = 0 To rptCount.Columns.Count + 1
                            Record.AddItem ""
                        Next
                        
                        If (Nvl(rsTmp("执行状态")) = 1 Or Nvl(rsTmp("执行状态")) = 3) Then
                            Me.rptCount.Columns(mAcol.已执行).Visible = True
                        End If
                        
                        If Nvl(rsTmp("执行状态")) = 1 Or Nvl(rsTmp("执行状态")) = 3 Then
                            Record(mAcol.已执行).Value = "√"
                        Else
                            Record(mAcol.已执行).Value = ""
                        End If
                        Record(mAcol.费用).Value = IIf(rsTmp("记录状态") = 1, "√", "×")
                        Record(mAcol.ID).Value = Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                        Record(mAcol.选择).HasCheckbox = True
                        Record(mAcol.选择).Checked = IIf(Nvl(rsTmp("执行状态")) = 0, True, False)
                        Record(mAcol.图标).BackColor = Val(Nvl(rsTmp("试管颜色")))
                        Record(mAcol.采集方式).Value = Nvl(rsTmp("采集方式"))
                        Record(mAcol.医嘱内容).Value = Nvl(rsTmp("医嘱内容"))
                        Record(mAcol.条码).Value = Nvl(rsTmp("样本条码"))
                        Record(mAcol.执行科室).Value = Nvl(rsTmp("执行科室"))
                        Record(mAcol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
                        Record(mAcol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
                        Record(mAcol.发送人).Value = Nvl(rsTmp("发送人"))
                        Record(mAcol.发送时间).Value = Nvl(rsTmp("发送时间"))
                        Record(mAcol.试管颜色).Value = Nvl(rsTmp("试管颜色"))
                        Record(mAcol.试管编码).Value = Nvl(rsTmp("试管编码"))
                        Record(mAcol.标本).Value = Nvl(rsTmp("标本")) & IIf(Nvl(rsTmp("婴儿")) = 0, "", "(婴儿)")
                        Record(mAcol.采样时间).Value = Nvl(rsTmp("采样时间"))
                        Record(mAcol.采样人).Value = Nvl(rsTmp("采样人"))
                        Record(mAcol.送检人).Value = Nvl(rsTmp("送检人"))
                        Record(mAcol.采血量).Value = Nvl(rsTmp("采血量"))
                        Record(mAcol.试管名称).Value = Nvl(rsTmp("试管名称"))
                        Record(mAcol.紧急).Value = Nvl(rsTmp("紧急"))
                        Record(mAcol.病人来源).Value = Nvl(rsTmp("病人来源"))
                        Record(mAcol.婴儿).Value = Nvl(rsTmp("婴儿"))
                        Record(mAcol.别名).Value = Nvl(rsTmp("别名"))
                        Record(mAcol.相关ID).Value = Nvl(rsTmp("相关ID"))
                        
                        Record(mAcol.病人ID).Value = Nvl(rsTmp("病人ID"))
                        Record(mAcol.姓名).Value = Nvl(rsTmp("病人姓名"))
                        Record(mAcol.性别).Value = Nvl(rsTmp("性别"))
                        Record(mAcol.年龄).Value = Nvl(rsTmp("年龄"))
                        Record(mAcol.标识号).Value = Nvl(rsTmp("标识号"))
                        Record(mAcol.床号).Value = Nvl(rsTmp("床号"))
                        Record(mAcol.病人科室).Value = Nvl(rsTmp("病人科室"))
                        Record(mAcol.接收人).Value = Nvl(rsTmp("接收人"))
                        Record(mAcol.接收时间).Value = Nvl(rsTmp("接收时间"))
                        Record(mAcol.诊疗项目ID).Value = Nvl(rsTmp("诊疗项目ID"))
                        Record(mAcol.执行状态).Value = Nvl(rsTmp("执行状态"))
                        Record(mAcol.医嘱id).Value = Nvl(rsTmp("医嘱id"))
                        Record(mAcol.送检时间).Value = Nvl(rsTmp("送检时间"))
                        
                        For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                            Record(intLoop).ForeColor = Val(Nvl(rsTmp("试管颜色")))
                        Next
                    Else
                        If InStr(Record(mAcol.医嘱内容).Value, Nvl(rsTmp("医嘱内容"))) <= 0 Then
                            Record(mAcol.医嘱内容).Value = Record(mAcol.医嘱内容).Value & " " & Nvl(rsTmp("医嘱内容"))
                        End If
                        Record(mAcol.合并医嘱).Value = Record(mAcol.合并医嘱).Value & ";" & _
                                                       Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                        Record(mAcol.别名).Value = Record(mAcol.别名).Value & " " & Nvl(rsTmp("别名"))
                    End If
                    strOldAdvice = rsTmp("相关ID")
                    If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("试管编码")) & ",") <= 0 Then
                        strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("试管编码"))
                    End If
                End If
                txt送检人 = Nvl(rsTmp("送检人"))
                rsTmp.MoveNext
            Loop
        End If
        Me.rptCount.Populate
    End If
    '---------------------------------重新计算当前登记的数量---------------------------------------
    strCuvetteNumber = ""
    strBarCode = ""
    Me.rptCuvetteCount.Records.DeleteAll
    Me.rptCuvetteCount.Populate
    For intLoop = 0 To Me.rptCount.Rows.Count - 1
        With Me.rptCount.Rows(intLoop)
        
            If chkNO(.Record(mAcol.试管编码).Value) = False Then
                Set Record = Me.rptCuvetteCount.Records.Add
                For intRow = 0 To rptCuvetteCount.Columns.Count + 1
                    Record.AddItem ""
                Next
                Record(mCuvetteCount.编码).Value = .Record(mAcol.试管编码).Value
                Record(mCuvetteCount.名称).Value = .Record(mAcol.试管名称).Value
                For intRow = 0 To Me.rptCuvetteCount.Columns.Count - 1
                    Record(intRow).ForeColor = .Record(mAcol.试管颜色).Value
                Next
                Me.rptCuvetteCount.Populate
            End If
        End With
    Next
    
    Me.rptCuvetteCount.Populate
    For intLoop = 0 To Me.rptCount.Rows.Count - 1
        With Me.rptCount.Rows(intLoop)
            If ChkBarCode(.Record(mAcol.条码).Value, intLoop) = False Then
                For intRow = 0 To Me.rptCuvetteCount.Rows.Count - 1
                    If Me.rptCuvetteCount.Rows(intRow).Record(mCuvetteCount.编码).Value = .Record(mAcol.试管编码).Value Then
                        Me.rptCuvetteCount.Rows(intRow).Record(mCuvetteCount.合计).Value = _
                            Val(Me.rptCuvetteCount.Rows(intRow).Record(mCuvetteCount.合计).Value) + 1
                        Exit For
                    End If
                Next
            End If

        End With
    Next
    
    Me.rptCuvetteCount.Populate
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function chkNO(strNO As String) As Boolean
    '检验试管编码是否重复
    Dim intLoop As Integer
    For intLoop = 0 To Me.rptCuvetteCount.Rows.Count - 1
        With Me.rptCuvetteCount.Rows(intLoop).Record
            If .Item(mCuvetteCount.编码).Value = strNO Then
                chkNO = True
                Exit For
            End If
        End With
    Next
End Function

Private Function ChkBarCode(strBarCode As String, intIndex As Integer) As Boolean
    '检验条码是否重复
    Dim intLoop As Integer
    For intLoop = 0 To intIndex - 1
        With Me.rptCount.Rows(intLoop).Record
            If .Item(mAcol.条码).Value = strBarCode Then
                ChkBarCode = True
                Exit For
            End If
        End With
    Next

End Function
Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function
Private Sub WriterCheckSampleToLIS(strAdvices As String, strName As String, strBatchNO As Long, Optional ByVal strSentName As String)
    '功能   把签收信息写入LIS
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.SampleCheckinInfoWrite(strAdvices, strName, strBatchNO, strErr, strSentName) = False Then
            MsgBox "写入签收信息到LIS申请单出错!" & vbCrLf & strErr
        End If
    End If
End Sub


Private Sub RePrintBarCode(ByVal blnNotRePrint As Boolean)
    Dim strSQL  As String
    Dim intFrom  As Integer
    With rptAlist(TabCtr.Selected.Index)
        If blnNotRePrint Then
            If .Records.Count = 0 Then
                Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
                Exit Sub
            End If
            If Trim(.FocusedRow.Record(mAcol.条码).Value) <> "" Then
                Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = True
            Else
                Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
            End If
        Else
            If .Records.Count = 0 Then
                MsgBox "未选中医嘱！", vbInformation, "提示"
                Exit Sub
            End If
            If mintCodeType = BarCodeType.Code39 Then
                Bar39 Me.picBarCode, 3, Trim(.FocusedRow.Record(mAcol.条码).Value), False, True
            Else
                Bar128 Me.picBarCode, 3, Trim(.FocusedRow.Record(mAcol.条码).Value), True
            End If
            
            SavePicture Me.picBarCode.Image, App.path & "\BarCode.bmp"
            Select Case Trim(.FocusedRow.Record(mAcol.病人来源).Value)
                Case "门诊"
                    intFrom = 1
                Case "住院"
                    intFrom = 2
                Case "院外"
                    intFrom = 3
                Case "体检"
                    intFrom = 4
            End Select
            
            
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1212_2", Me, "样本条码=" & Trim(.FocusedRow.Record(mAcol.条码).Value), _
            "项目 = " & Trim(.FocusedRow.Record(mAcol.医嘱内容).Value), _
            "病人姓名 = " & IIf(Trim(.FocusedRow.Record(mAcol.姓名).Value) <> "", Trim(.FocusedRow.Record(mAcol.姓名).Value), "无"), _
            "性别 = " & IIf(Trim(.FocusedRow.Record(mAcol.性别).Value) <> "", Trim(.FocusedRow.Record(mAcol.性别).Value), "无"), _
            "年龄 = " & IIf(Trim(.FocusedRow.Record(mAcol.年龄).Value) <> "", Trim(.FocusedRow.Record(mAcol.年龄).Value), "无"), _
            "床号 = " & IIf(Trim(.FocusedRow.Record(mAcol.床号).Value) <> "", Trim(.FocusedRow.Record(mAcol.床号).Value), "无"), _
            "标识号 = " & IIf(Trim(.FocusedRow.Record(mAcol.标识号).Value) <> "", Trim(.FocusedRow.Record(mAcol.标识号).Value), "无"), _
            "所在科室 = " & IIf(Trim(.FocusedRow.Record(mAcol.病人科室).Value) <> "", Trim(.FocusedRow.Record(mAcol.病人科室).Value), "无"), _
            "采集方式 = " & IIf(Trim(.FocusedRow.Record(mAcol.采集方式).Value) <> "", Trim(.FocusedRow.Record(mAcol.采集方式).Value), "无"), _
            "标本 = " & IIf(Trim(.FocusedRow.Record(mAcol.标本).Value) <> "", Trim(.FocusedRow.Record(mAcol.标本).Value), "无"), _
            "执行科室 = " & IIf(Trim(.FocusedRow.Record(mAcol.执行科室).Value) <> "", Trim(.FocusedRow.Record(mAcol.执行科室).Value), "无"), _
            "开嘱医生 = " & IIf(Trim(.FocusedRow.Record(mAcol.开嘱医生).Value) <> "", Trim(.FocusedRow.Record(mAcol.开嘱医生).Value), "无"), _
            "开嘱时间 = " & IIf(Trim(.FocusedRow.Record(mAcol.开嘱时间).Value) <> "", Trim(.FocusedRow.Record(mAcol.开嘱时间).Value), "无"), _
            "采样人 = " & IIf(Trim(.FocusedRow.Record(mAcol.采样人).Value) <> "", Trim(.FocusedRow.Record(mAcol.采样人).Value), "无"), _
            "采样时间 = " & IIf(Trim(.FocusedRow.Record(mAcol.采样时间).Value) <> "", Trim(.FocusedRow.Record(mAcol.采样时间).Value), "无"), _
            "管码 = " & IIf(Trim(.FocusedRow.Record(mAcol.试管编码).Value) <> "", Trim(.FocusedRow.Record(mAcol.试管编码).Value), "无"), _
            "采血量 = " & IIf(Trim(.FocusedRow.Record(mAcol.采血量).Value) <> "", Trim(.FocusedRow.Record(mAcol.采血量).Value), "无"), _
            "试管名称 = " & IIf(Trim(.FocusedRow.Record(mAcol.试管名称).Value) <> "", Trim(.FocusedRow.Record(mAcol.试管名称).Value), "无"), _
            "紧急 = " & IIf(Trim(.FocusedRow.Record(mAcol.紧急).Value) <> "", Trim(.FocusedRow.Record(mAcol.紧急).Value), 0), _
            "病人来源 = " & intFrom, _
            "条码图像1=" & App.path & "\BarCode.Bmp", 2)
            Kill App.path & "\BarCode.Bmp"
            
        End If
    End With
End Sub

Private Function getTATTime(strIDs As String) As Boolean
    '检查TAT限时,返回可以送检的医嘱ID
    Dim strSex As String    '性别
    Dim strDept As String   '申请科室
    Dim strItem As String   '申请项目   项目ID1,项目名称1,采样时间1,急诊1;项目ID2,项目名称12,采样时间2,急诊2........
    Dim Record As ReportRecord
    Dim intMsg As Integer
    Dim strShowBef As String
    Dim strMsgShow As String
    Dim strMsgShowStop As String
    Dim strItemCode As String
    Dim strMsgNoTime As String '没有上一个时间节点的项目
    
    Dim strTATItems As String
    Dim var_Tmp As Variant
    Dim var_Tmp1 As Variant
    Dim var_Item As Variant

    Dim strErr As String
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrHand
    
    If mobjLisInsideComm Is Nothing Then
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not mobjLisInsideComm Is Nothing Then
            '初始化LIS接口部件
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "初始化LIS接口失败！" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If
    
'    If Me.rptPlist.FocusedRow Is Nothing Then
'        getTATTime = False
'        Exit Function
'    End If
    
    '获取病人性别和申请科室
'    With Me.rptPlist.FocusedRow
'        strSex = .Record(mPcol.性别).Value
'        strDept = .Record(mPcol.病人科室).Value
'    End With
    
    '获取项目ID,项目名称,采样时间,急诊
    strItem = ""
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.选择).Checked = True Then
'            If Record(mAcol.送检时间).Value <> "" Then
                var_Item = Split(Mid(strIDs, 2), "|")
                For i = LBound(var_Item) To UBound(var_Item)
                    strItem = strItem & ";" & Record(mAcol.诊疗项目ID).Value & "," & Record(mAcol.医嘱内容).Value & _
                                            "," & Record(mAcol.送检时间).Value & "," & IIf(Record(mAcol.紧急).Value = "紧急", 1, 0) & _
                                             "," & var_Item(i) & "," & Record(mAcol.条码).Value
                Next
'            Else
'                strMsgNoTime = strMsgNoTime & Record(mAcol.医嘱内容).Value & vbCrLf
'            End If
            '获取病人性别和申请科室
            strSex = Record(mAcol.性别).Value
            strDept = Record(mAcol.申请科室).Value
        End If
    Next
    
    If strMsgNoTime <> "" Then MsgBox strMsgNoTime & "未送检,不能签收   ", vbInformation, Me.Caption
    If strItem <> "" Then
        strItem = Mid(strItem, 2)
    Else
        strIDs = ""
        getTATTime = False
        Exit Function
    End If
    
    '检查TAT是否超时
    On Error GoTo errold
    strTATItems = mobjLisInsideComm.GetTatTimeShow(2, strItem, strDept, "", "", strSex, intMsg, strShowBef, , UserInfo.姓名)
    If strTATItems <> "" Then
        var_Tmp = Split(strTATItems, ";")
        Do While UBound(Split(var_Tmp(0), ",")) < 9
            '不足9个元素的在最后拼凑一个0
            strTATItems = ""
            For i = LBound(var_Tmp) To UBound(var_Tmp)
                strTATItems = strTATItems & ";" & var_Tmp(i) & ",0"
            Next
            If strTATItems <> "" Then strTATItems = Mid(strTATItems, 2)
            var_Tmp = Split(strTATItems, ";")
        Loop
        
        '获取所有超时且禁止项目的条码
        For i = LBound(var_Tmp) To UBound(var_Tmp)
            If Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 Then
                strItemCode = strItemCode & "," & Split(var_Tmp(i), ",")(6)
            End If
        Next
        
        strIDs = ""
        
        For i = LBound(var_Tmp) To UBound(var_Tmp)
            If Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 1 And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                '已超时只提示
                strMsgShow = strMsgShow & Replace(Replace(Split(var_Tmp(i), ",")(8), "[项目]", Split(var_Tmp(i), ",")(1)), "[超时]", Split(var_Tmp(i), ",")(7) & "分钟") & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 1 And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) > 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                '用相同条码项目的
                strMsgShow = strMsgShow & Replace(Replace(Split(var_Tmp(i), ",")(8), "[项目]", Split(var_Tmp(i), ",")(1)), "[超时]", "") & "存在同条码禁止项目,不能继续" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(8) <> "0" And Split(var_Tmp(i), ",")(2) = "" Then
                '没有前一个时间节点的
                strMsgShowStop = strMsgShowStop & Split(var_Tmp(i), ",")(1) & "未送检,不能签收" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 And Split(var_Tmp(i), ",")(2) <> "" Then
                '超时并禁止的
                strMsgShowStop = strMsgShowStop & Replace(Replace(Split(var_Tmp(i), ",")(8), "[项目]", Split(var_Tmp(i), ",")(1)), "[超时]", Split(var_Tmp(i), ",")(7) & "分钟") & vbCrLf
            Else
                '不同项目同条码的时候,当有一个项目超时,则所有该条码的项目均不能送检
                If InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 Then
                    strIDs = strIDs & "|" & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
                End If
            End If
        Next
        
        '当设置为提示时,如果点了时,则送检所有勾选的项目,点了否,则只送检为超时的标本
        If strMsgShow <> "" Then
            If MsgBox(strMsgShow & "是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If strIDs <> "" Then
                    getTATTime = True
                Else
                    getTATTime = False
                End If
                Exit Function
            Else
                '点击是,则从新组合所有勾选的项目
                strIDs = ""
                For i = LBound(var_Tmp) To UBound(var_Tmp)
                    If (Split(var_Tmp(i), ",")(7) = 0 Or Split(var_Tmp(i), ",")(9) = 1) And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                        strIDs = strIDs & "|" & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
                    End If
                Next
            End If
        End If
        If strMsgShowStop <> "" Then
            MsgBox strMsgShowStop, vbInformation, Me.Caption
            getTATTime = True
            Exit Function
        End If
        
    End If
    getTATTime = True
    
    Exit Function
errold:
    getTATTime = True
    
    Exit Function
ErrHand:
    MsgBox Err.Description, vbInformation, Me.Caption
    Err.Clear
End Function

Private Sub txt送检人_DblClick()
    Dim strVal As String
    Dim rsTmp As Recordset
    If Not mrsSendPerson Is Nothing Then
        mrsSendPerson.filter = ""
        If mrsSendPerson.RecordCount > 0 Then
            Set rsTmp = mrsSendPerson
            strVal = frmSelectPub.ShowMe(Me, rsTmp, "")
            If strVal <> "" Then
                txt送检人.Text = Split(strVal, ",")(2)
                mstrSendPerson = txt送检人.Text
                cmdOK.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txt送检人_KeyPress(KeyAscii As Integer)
    Dim strVal As String
    Dim rsTmp As Recordset
    If KeyAscii = 13 Then
        If Not mrsSendPerson Is Nothing Then
            mrsSendPerson.filter = ""
            If mrsSendPerson.RecordCount > 0 Then
                Set rsTmp = mrsSendPerson
                strVal = frmSelectPub.ShowMe(Me, rsTmp, Trim(txt送检人.Text))
                If strVal <> "" Then
                    txt送检人.Text = Split(strVal, ",")(2)
                    mstrSendPerson = txt送检人.Text
                    cmdOK.SetFocus
                End If
            End If
        End If
    End If
End Sub

