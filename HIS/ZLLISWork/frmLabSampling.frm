VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "*\A..\ZLIDKIND\zlIDKind.vbp"
Begin VB.Form frmLabSampling 
   AutoRedraw      =   -1  'True
   Caption         =   "检验采集工作站"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15360
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabSampling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   15360
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   0
      Left            =   13620
      TabIndex        =   39
      Top             =   3840
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   1
      Left            =   13980
      TabIndex        =   40
      Top             =   4320
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   2
      Left            =   13980
      TabIndex        =   41
      Top             =   4680
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   3
      Left            =   13980
      TabIndex        =   42
      Top             =   4800
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   5
      Left            =   13380
      TabIndex        =   43
      Top             =   4230
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   4
      Left            =   12540
      TabIndex        =   44
      Top             =   3720
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2004
      Left            =   8760
      ScaleHeight     =   2010
      ScaleWidth      =   3375
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   3372
      Begin XtremeSuiteControls.TabControl TabCtr 
         Height          =   1245
         Left            =   720
         TabIndex        =   50
         Top             =   420
         Width           =   1965
         _Version        =   589884
         _ExtentX        =   3466
         _ExtentY        =   2196
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBarCodePrint 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8220
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   8
      Top             =   2130
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   8250
      ScaleHeight     =   2985
      ScaleWidth      =   7935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   450
      Width           =   7935
      Begin XtremeReportControl.ReportControl rptPlist 
         Height          =   1125
         Left            =   1260
         TabIndex        =   9
         Top             =   1380
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   6555
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   6555
         Begin VB.OptionButton optFilter 
            Caption         =   "拒收"
            Height          =   225
            Index           =   6
            Left            =   5580
            TabIndex        =   36
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "已执行"
            Height          =   225
            Index           =   5
            Left            =   4650
            TabIndex        =   35
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "已送检"
            Height          =   225
            Index           =   4
            Left            =   3720
            TabIndex        =   34
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "已采样"
            Height          =   225
            Index           =   3
            Left            =   2730
            TabIndex        =   33
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "已绑定"
            Height          =   225
            Index           =   2
            Left            =   1770
            TabIndex        =   32
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "未绑定"
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   31
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "全部"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   30
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
      End
      Begin XtremeSuiteControls.ShortcutCaption srtPatient 
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2235
         _Version        =   589884
         _ExtentX        =   3942
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
      Height          =   6570
      Left            =   30
      ScaleHeight     =   6570
      ScaleWidth      =   8145
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   8145
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
         Height          =   1785
         Left            =   60
         TabIndex        =   51
         Top             =   750
         Width           =   8025
         Begin VB.TextBox txt年龄1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5550
            MaxLength       =   5
            TabIndex        =   63
            Top             =   210
            Width           =   555
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "…"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5820
            TabIndex        =   62
            Top             =   1350
            Width           =   285
         End
         Begin VB.ComboBox cbo医生 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3570
            TabIndex        =   61
            Top             =   960
            Width           =   1275
         End
         Begin VB.ComboBox cbo开单科室 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmLabSampling.frx":058A
            Left            =   900
            List            =   "frmLabSampling.frx":058C
            TabIndex        =   60
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txt医嘱内容 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   900
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   1350
            Width           =   4935
         End
         Begin VB.TextBox txtID 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txtPatientDept 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3570
            TabIndex        =   57
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtBed 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5310
            TabIndex        =   56
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox cboAge 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmLabSampling.frx":058E
            Left            =   4770
            List            =   "frmLabSampling.frx":05A1
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   210
            Width           =   750
         End
         Begin VB.TextBox txt年龄 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   54
            Top             =   210
            Width           =   435
         End
         Begin VB.ComboBox cbo性别 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmLabSampling.frx":05BD
            Left            =   3210
            List            =   "frmLabSampling.frx":05BF
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   210
            Width           =   675
         End
         Begin VB.TextBox txt姓名 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   900
            MaxLength       =   20
            TabIndex        =   52
            ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
            Top             =   210
            Width           =   1635
         End
         Begin VB.Image imgPatient 
            Height          =   975
            Left            =   6390
            Picture         =   "frmLabSampling.frx":05C1
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1125
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "记"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   510
            Index           =   10
            Left            =   7320
            TabIndex        =   74
            Top             =   1200
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请医生"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   2790
            TabIndex        =   73
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请科室"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   150
            TabIndex        =   72
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请项目"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   150
            TabIndex        =   71
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓       名"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   70
            Top             =   255
            Width           =   675
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "拒收"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   510
            Index           =   6
            Left            =   6330
            TabIndex        =   69
            Top             =   1200
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标  识 号"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   68
            Top             =   645
            Width           =   675
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床号"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   4920
            TabIndex        =   67
            Top             =   1005
            Width           =   360
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "所在科室"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   2790
            TabIndex        =   66
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2790
            TabIndex        =   65
            Top             =   255
            Width           =   360
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   3915
            TabIndex        =   64
            Top             =   255
            Width           =   360
         End
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   405
         Left            =   150
         TabIndex        =   47
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
         IDKindStr       =   $"frmLabSampling.frx":148B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         DefaultCardType =   "0"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.CheckBox chkFindMove 
         BackColor       =   &H00FDD6C6&
         Caption         =   "查找到病人后焦点移动到条码输入"
         Height          =   225
         Left            =   5100
         TabIndex        =   7
         Top             =   30
         Width           =   3135
      End
      Begin VB.Frame fraBarCode 
         BackColor       =   &H00FDD6C6&
         Caption         =   "条码绑定和生成"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   60
         TabIndex        =   3
         Top             =   2550
         Width           =   8025
         Begin VB.Frame Frame2 
            BackColor       =   &H00FDD6C6&
            BorderStyle     =   0  'None
            Height          =   3090
            Left            =   60
            TabIndex        =   18
            Top             =   840
            Width           =   7875
            Begin XtremeReportControl.ReportControl rptCuvette 
               Height          =   1935
               Left            =   90
               TabIndex        =   19
               Top             =   90
               Width           =   6375
               _Version        =   589884
               _ExtentX        =   11245
               _ExtentY        =   3413
               _StockProps     =   0
               AllowColumnRemove=   0   'False
               MultipleSelection=   0   'False
               SkipGroupsFocus =   0   'False
               EditOnClick     =   0   'False
               AutoColumnSizing=   0   'False
            End
            Begin VB.CheckBox chkMaterial 
               BackColor       =   &H00FDD6C6&
               Caption         =   "自动发料退料"
               Height          =   225
               Left            =   90
               TabIndex        =   49
               Top             =   2820
               Width           =   2835
            End
            Begin VB.CheckBox chkApplyDept 
               BackColor       =   &H00FDD6C6&
               Caption         =   "生成条码时区分申请科室"
               Height          =   225
               Left            =   4620
               TabIndex        =   48
               Top             =   2580
               Width           =   2835
            End
            Begin VB.CheckBox chkBindPage 
               BackColor       =   &H00FDD6C6&
               Caption         =   "生成或绑定条码后跳转到已绑定页"
               Height          =   195
               Left            =   90
               TabIndex        =   46
               Top             =   2595
               Width           =   3015
            End
            Begin VB.CheckBox chkSendPrint 
               BackColor       =   &H00FDD6C6&
               Caption         =   "取消送检单打印"
               Height          =   195
               Left            =   90
               TabIndex        =   38
               Top             =   2370
               Width           =   1755
            End
            Begin VB.CheckBox chkDeptShow 
               BackColor       =   &H00FDD6C6&
               Caption         =   "只显示当前采集科室试管"
               Height          =   225
               Left            =   1890
               TabIndex        =   37
               Top             =   2130
               Width           =   2505
            End
            Begin VB.CommandButton cmdBindBarCode 
               Caption         =   "绑定条码(&B)"
               Height          =   345
               Left            =   6600
               TabIndex        =   28
               Top             =   75
               Width           =   1185
            End
            Begin VB.CommandButton cmdNewBarcode 
               Caption         =   "生成条码(&N)"
               Height          =   345
               Left            =   6600
               TabIndex        =   27
               Top             =   465
               Width           =   1185
            End
            Begin VB.CheckBox chkBackBill 
               BackColor       =   &H00FDD6C6&
               Caption         =   "已完成打印回执单"
               Height          =   225
               Left            =   4620
               TabIndex        =   26
               Top             =   2130
               Width           =   1785
            End
            Begin VB.CheckBox chkComPlete 
               BackColor       =   &H00FDD6C6&
               Caption         =   "生成或绑定条码后标志为已采集"
               Height          =   225
               Left            =   4620
               TabIndex        =   25
               Top             =   2370
               Width           =   2835
            End
            Begin VB.CommandButton cmdBarcodePrint 
               Caption         =   "条码打印(&B)"
               Height          =   345
               Left            =   6600
               Picture         =   "frmLabSampling.frx":1530
               TabIndex        =   24
               Top             =   1260
               Width           =   1185
            End
            Begin VB.CommandButton cmdComplete 
               Caption         =   "完成采集(&P)"
               Height          =   345
               Left            =   6600
               Picture         =   "frmLabSampling.frx":167A
               TabIndex        =   23
               Top             =   870
               Width           =   1185
            End
            Begin VB.CommandButton cmdBakBillPrint 
               Caption         =   "回执单打印"
               Height          =   345
               Left            =   6600
               Picture         =   "frmLabSampling.frx":17C4
               TabIndex        =   22
               Top             =   1665
               Width           =   1185
            End
            Begin VB.CheckBox ChkBarCodePrint 
               BackColor       =   &H00FDD6C6&
               Caption         =   "生成或绑定条码后打印条码"
               Height          =   225
               Left            =   1890
               TabIndex        =   21
               Top             =   2370
               Width           =   2505
            End
            Begin VB.CheckBox chkPrintBarCode 
               BackColor       =   &H00FDD6C6&
               Caption         =   "已完成打印条码"
               Height          =   225
               Left            =   90
               TabIndex        =   20
               Top             =   2130
               Width           =   1575
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FDD6C6&
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   60
            TabIndex        =   11
            Top             =   210
            Width           =   7875
            Begin VB.TextBox TxtBarCode 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   990
               TabIndex        =   15
               Top             =   75
               Width           =   2145
            End
            Begin VB.TextBox TxtBarCodeCheck 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   4200
               TabIndex        =   14
               Top             =   75
               Width           =   2145
            End
            Begin VB.Frame fraSpace 
               BackColor       =   &H00FDD6C6&
               Height          =   45
               Left            =   60
               TabIndex        =   13
               Top             =   540
               Width           =   7785
            End
            Begin VB.CheckBox ChkContinuous 
               BackColor       =   &H00FDD6C6&
               Caption         =   "连续输入"
               Height          =   225
               Left            =   6480
               TabIndex        =   12
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label LabCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "条码输入"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   17
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "条码确认"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3300
               TabIndex        =   16
               Top             =   165
               Width           =   840
            End
         End
      End
      Begin VB.TextBox txtGoto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   795
         TabIndex        =   0
         ToolTipText     =   "“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
         Top             =   285
         Width           =   7275
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
         TabIndex        =   2
         Top             =   30
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   9270
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":190E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":197A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":1F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":24AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":2A48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7230
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabSampling.frx":2FE2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22013
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
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   0
      TabIndex        =   6
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
   Begin VB.Image imgLoad 
      Height          =   975
      Left            =   9390
      Picture         =   "frmLabSampling.frx":3876
      Stretch         =   -1  'True
      Top             =   3390
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgDefual 
      Height          =   975
      Left            =   8100
      Picture         =   "frmLabSampling.frx":4740
      Stretch         =   -1  'True
      Top             =   3270
      Visible         =   0   'False
      Width           =   1125
   End
   Begin XtremeSuiteControls.PopupControl PopupControl 
      Left            =   10020
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      Width           =   300
      ShowDelay       =   6000
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   8370
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabSampling.frx":560A
      Left            =   8880
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabSampling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPcol                                  '病人列表
    病人ID = 1
    紧急
    临床路径病人
    来源
    病人姓名
    性别
    年龄
    标识号
    床号
    病人科室
    状态
    拒收
    就诊卡
    条码
    挂号单
    已执行
    未绑定
    已绑定
    已采样
    重采标本
    已送检
    合计
    发送时间
    主页ID
End Enum

Private Enum mAcol                                  '医嘱列表
    类别
    ID
    选择
    急诊
    图标
    重采
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
    采血量
    试管名称
    紧急
    病人来源
    申请科室
    婴儿
    别名
    相关ID
    执行状态
    NO
    审核时间
    接收时间
    记录状态
    接收人
    条码打印
    拒收理由
    送出时间
    婴儿姓名
    采集科室ID
    采集执行科室
    诊疗项目ID
    诊疗项目组合
    检验执行科室ID
    计费状态
    记录性质
    婴儿性别
    病人所在科室
End Enum
Private Enum mDkp                                   '窗格ID
    条码操作 = 0
    医嘱列表
    病人列表
End Enum

Private Enum mFilter
    标识号 = 0
    就诊卡
    姓名
    单据号
    标本
    采集方式
    门诊
    住院
    体检
    间隔时间
    发送或审核时间          '=0 发送时间 = 1 审核时间
    开始时间
    结束时间
    检验类型
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

Private mlngDeptID As Long                                              '科室ID
Private mlngKey As Long                                                 'ID
Private mstrPrivs As String                                             '权限
Private mblnSaveAdvice As Boolean                                       '是否需要保存医嘱，用于修改在院病人标本信息
Private PatientType As Integer, mlng病人ID As Long, mstrNO As String    '门诊收费单据号
Private mblnBarCode As Boolean                                          '条码
Private mlngReqDept As Long, mstrReqDoctor As String                    '默认的登记科室和医生
Private mstrKeys As String                                              '当前核收的申请医嘱ID
Private mintEditMode As Integer                                         '0－核收、1－登记、2－重新核收、3－补填申请
Private rsRelativeAdvice As ADODB.Recordset                             '登记的相关医嘱
Private mlngCapID As Long                                               '采集项目ID
Private mstrExtData  As String                                           '登记的申请项目信息
Private ItemDeptID As Long, mlngDefaultDevice As Long
Private mbln微生物项目 As Boolean
Private mblFind As Boolean                                              '是否查找到的病人
Private mstrOldTime As String                                           '记录旧的时间用于定时提醒
Private iInputType As Integer
Private mintTop As Integer                                              '工具条高
Private mintHeight As Integer                                           '区域高
Private mobjSquareCard As Object                                        '取卡类型
Private mblnNowConsumption As Boolean                                   '是否立即付款

Private mblnShowPwd As Boolean                                          '是否显示密文
    
Private mstrIndex As String                                             '病人查找方式
'病人姓名当前输入状态，如果一直以该状态可以不输入前导符
'0：就诊卡
'1：病人ID
'2：住院号
'3：门诊号
'4：挂号单
'5：收费单据号
'6：姓名
'-------------------------------------------- 2007-08-17 加入一卡通支持
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private Enum IDKinds
    C0姓名 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
    C4门诊号 = 4
    C5就诊卡 = 5
End Enum
Private mbln身份证 As Boolean
Private Const conMenu_IDkind_Change  As Integer = 12345
Private mstrBarCodes As String                                          '选择当前的条码串使用逗号分隔多个条码


Private Function ReadPatPricture(ByVal lng病人ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '参数：lng病人ID=读取指定病人的照片
    '           imgPatient=照片加载位置
    '           strFile=照片的本地路径
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = Sys.ReadLob(glngSys, 27, lng病人ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CreateCbs()
    '功能创建工具条
    
    '创建菜单
    Dim Control As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim intBarCode As Integer                       '使那条码 1=39Code , 2=128Code(默认)
    Dim intExecDept As Integer                      '不区分执行科室打印
    Dim intHideBarCode As Integer                   '隐藏预置条码
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    
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
    '-----------------------------------------------------
   intBarCode = zlDatabase.GetPara("使用条码", "100", "1211", 2)
   intExecDept = zlDatabase.GetPara("不区分执行科室打印", "100", "1211", 1)
   intHideBarCode = zlDatabase.GetPara("隐藏预置条码", "100", "1211", 0)
    '==文件菜单
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With ControlFile.CommandBar.Controls
        Set Control = .Add(xtpControlPopup, conMenu_File_PrintSet, "打印设置(&S)…")
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_MedRecSetup, "条码打印设置", -1, False
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_MedRecPreview, "回执单打印设置", -1, False
            
        .Add xtpControlButton, conMenu_File_Preview, "清单预览(&V)"
        .Add xtpControlButton, conMenu_File_Print, "清单打印(&P)"
        Set Control = .Add(xtpControlPopup, conMenu_File_MedRecPrint, "打印")
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_RowPrint, "打印条码(&C)", -1, False
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_BatPrint, "打印回执单(&B)", -1, False
        Set Control = .Add(xtpControlPopup, conMenu_Edit_Send, "条码设置")
            Set cbrControl = Control.CommandBar.Controls.Add(xtpControlButton, conMenu_Tool_SignNew, "使用39Code", -1, False)
            If intBarCode = 1 Then cbrControl.Checked = True
            Set cbrControl = Control.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendOther, "使用128Code", -1, False)
            If intBarCode = 2 Then cbrControl.Checked = True
            Set cbrControl = Control.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Transfer_Force, "不区分执行科室打印", -1, False)
            If intExecDept = 0 Then cbrControl.Checked = True
        .Add xtpControlButton, conMenu_File_Excel, "输出到&Excel…"
        Set Control = .Add(xtpControlButton, conMenu_File_Parameter, "设备配置", -1, False)

        Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        Control.BeginGroup = True
    End With
    
    '==编辑
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    With ControlFile.CommandBar.Controls
        Set Control = .Add(xtpControlButton, conMenu_Manage_RequestView, "条码绑定(&B)")
        Set Control = .Add(xtpControlButton, conMenu_Manage_RequestPrint, "条码生成(&N)")
        Set Control = .Add(xtpControlButton, conMenu_Manage_RequestBatPrint, "完成采集(&P)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_Tool_MedRec, "病区条码打印(&A)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_Manage_Plan, "直接登记(&G)"): Control.BeginGroup = True
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
        Set Control = .Add(xtpControlButton, conMenu_View_PriceTable, "隐藏预置条码(&H)"): Control.BeginGroup = True
        Control.Checked = (intHideBarCode = 1): Call HideBarCode
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
    
    If Not mobjZLIHISPlugIn Is Nothing Then
        Dim astrPlug() As String
        Dim intLoop As Integer
        '==插件菜单
        astrPlug = Split(mobjZLIHISPlugIn.GetFuncNames(glngSys, glngModul), ",")
        Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PlugIn, "扩展功能(&P)", -1, False)
        With ControlFile.CommandBar.Controls
            For intLoop = 0 To UBound(astrPlug)
                .Add xtpControlButton, conMenu_PlugIn_Menu + intLoop + 1, astrPlug(intLoop)
            Next
            Control.BeginGroup = True
        End With
    End If
    
    '==列表科室
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "采样科室")
    Control.Flags = xtpFlagRightAlign
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlComboBox, conMenu_View_Busy, "科室")
    Control.ShortcutText = "科室"
    Control.Width = 130
    Control.Flags = xtpFlagRightAlign
    Control.Style = xtpButtonIconAndCaption
    
    '创建工具条
    Dim Toolbar As CommandBar
    Dim ControlPopup As CommandBarPopup
    
    Set Toolbar = cbrthis.Add("工具栏", xtpBarTop)
    Toolbar.ShowTextBelowIcons = False
    Toolbar.EnableDocking xtpFlagStretched
    With Toolbar.Controls
        .Add xtpControlButton, conMenu_File_Preview, "预览"
        .Add xtpControlButton, conMenu_File_Print, "打印"
        Set Control = .Add(xtpControlButton, conMenu_Tool_MedRec, "病区条码打印"): Control.BeginGroup = True
        
        Set Control = .Add(xtpControlButton, conMenu_View_Show, "送检核对")
        If InStr(GetPrivFunc(2500, 2001), "送检核对") = 0 Then
            Control.Visible = False
        End If
        
        Set Control = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): Control.BeginGroup = True
        .Add xtpControlButton, conMenu_View_Refresh, "刷新"
        Set Control = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): Control.BeginGroup = True
        .Add xtpControlButton, conMenu_File_Exit, "退出"
    End With
    
    For Each Control In Toolbar.Controls
        Control.Style = xtpButtonIconAndCaption
    Next
    
    '快键绑定
    With cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add FCONTROL, Asc("A"), conMenu_Tool_MedRec
'        .Add FCONTROL, Asc("C"), conMenu_File_RowPrint
        .Add FCONTROL, Asc("B"), conMenu_File_BatPrint
        .Add 0, VK_F2, conMenu_Manage_RequestView
        .Add 0, VK_F3, conMenu_Manage_RequestPrint
        .Add 0, VK_F4, conMenu_Manage_RequestBatPrint
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F6, conMenu_Manage_Plan
        .Add 0, VK_F10, conMenu_IDkind_Change
        
    End With
    '设置不常用菜单
    With cbrthis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs, "ZL1_INSIDE_1211_1", "ZL1_INSIDE_1211_2", "ZL1_INSIDE_1211_3")
    Me.cbrthis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    mintTop = lngTop
End Sub
Private Sub CreateDkp()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane

    dkpMan.SetCommandBars Me.cbrthis
    dkpMan.Options.DefaultPaneOptions = PaneNoCloseable
    dkpMan.Options.HideClient = True
    
    Set Pane1 = dkpMan.CreatePane(mDkp.条码操作, 400, 700, DockLeftOf, Nothing)
    Pane1.Title = "条码操作"
'    Pane1.Handle = Me.picBarCodeWork.hWnd
    Pane1.Options = PaneNoCaption
    
    Set Pane2 = dkpMan.CreatePane(mDkp.医嘱列表, 400, 300, DockBottomOf, Pane1)
    Pane2.Title = "医嘱信息"
'    Pane2.Handle = TabCtr.hWnd
    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMan.CreatePane(mDkp.病人列表, 600, 300, DockRightOf, Nothing)
    Pane3.Title = "病人采集清单"
'    Pane3.Handle = Me.picTab.hWnd
    Pane3.Options = PaneNoCaption
    
    Pane1.Select
End Sub

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_Click()
    If cbo开单科室.ListIndex > -1 Then InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub

Private Sub cbo开单科室_GotFocus()
    Call zlControl.TxtSelAll(cbo开单科室)
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    '</CSCustomCode> 1
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo开单科室.ListIndex <> -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex): Exit Sub '已选中
    If cbo开单科室.Text = "" Then '无输入
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo开单科室.Text))
    '全院临床科室
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检'))" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    
    On Error GoTo errH
    vRect = GetControlRect(cbo医生.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱科室", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo开单科室.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        If Not zlControl.CboLocate(cbo开单科室, rsTmp!名称) Then
            cbo开单科室.Text = ""
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的科室。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Me.cbo开单科室.ListIndex > -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
End Sub

Private Sub cbo医生_GotFocus()
    Call zlControl.TxtSelAll(cbo医生)
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo医生.ListIndex <> -1 Then mstrReqDoctor = Me.cbo医生.Text: Exit Sub '已选中
    If cbo医生.Text = "" Then '无输入
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo医生.Text))
    '全院医生
    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And B.部门ID IN(" & strSQL & ")" & _
        " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        " Order by A.简码"
    
    On Error GoTo errH
    vRect = GetControlRect(cbo医生.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱医生", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cbo医生.Text = rsTmp!姓名
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Len(Trim(Me.cbo医生.Text)) > 0 Then mstrReqDoctor = Me.cbo医生.Text
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strFilter As String                             '过滤字串
    Dim cboCtrol As CommandBarComboBox                  '科室
    Dim Controlcbo As CommandBarComboBox                '下拉框
    Dim cbrControl As CommandBarControl                 '文本标签
    Dim intExecDept As Integer                          '不同执行科室的是否一起打印
    Dim intDay As Integer
    Dim strTmp As String
    
    Select Case Control.ID
    
        Case conMenu_File_MedRecSetup                                               '条码打印设置
            ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me
        
        Case conMenu_File_MedRecPreview                                             '回执单打印设置
            ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1211_2", Me
            
        Case conMenu_File_Preview                                                   '清单预览
            Call zlRptPrint(2)
        
        Case conMenu_File_Print                                                     '清单打印
            Call zlRptPrint(1)
            
        Case conMenu_File_RowPrint                                                  '条码打印
            Call CmdBarCodePrint_Click
        
        Case conMenu_File_BatPrint                                                  '回执单打印
            Call cmdBakBillPrint_Click
                    
        Case conMenu_Tool_SignNew                                                   '使用39码
            Control.Checked = True
            Set cbrControl = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_SendOther, True, True)
            cbrControl.Checked = False
            
        Case conMenu_Edit_SendOther                                                 '使用128码
            Control.Checked = True
            Set cbrControl = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Tool_SignNew, True, True)
            cbrControl.Checked = False
        
        Case conMenu_Manage_Transfer_Force                                          '不区分科室打印
            Control.Checked = Not Control.Checked
        
        Case conMenu_File_Excel                                                     '输出到Excel
            Call zlRptPrint(3)
        
        Case conMenu_File_Parameter                                                 '设备配置
            'frmLabSampleSetup.Show vbModal, Me
            Call zlCommFun.DeviceSetup(Me, 100, 1101)
        Case conMenu_File_Exit                                                      '退出
            Unload Me
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_RequestView                                             '条码绑定
            Call cmdBindBarCode_Click
        
        Case conMenu_Manage_RequestPrint                                            '条码生成
            Call cmdNewBarcode_Click
        
        Case conMenu_Manage_RequestBatPrint                                         '完成采集
            Call cmdComplete_Click
        
        Case conMenu_Tool_MedRec                                                    '条码批量打印
            Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
            intExecDept = IIf(Control.Checked, 0, 1)
            Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
            intExecDept = IIf(Control.Checked, 0, 1)
            Set cbrControl = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_SendOther, True, True)
            frmLabBarCodeBatPrint.ShowMe Me, mstrPrivs, IIf(cbrControl.Checked, 2, 1), intExecDept, mblnNowConsumption
        Case conMenu_Manage_Plan                                                    '直接登记
            If InStr(mstrPrivs, "直接登记") > 0 Then
                frmLabSamplingRegister.ShowMe Me
            End If
        Case conMenu_View_Show                                                      '送检核对
            strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")
            If Me.rptPlist.Tag <> "" Then
                intDay = Val(Split(Me.rptPlist.Tag, ";")(mFilter.间隔时间))
            Else
                If strTmp <> "" Then
                    intDay = Val(Split(strTmp, ";")(mFilter.间隔时间))
                End If
            End If
            
            If Not mobjLisInsideComm Is Nothing Then
                Call mobjLisInsideComm.ShowFrmSampleSendCheck(Me, 1, intDay)
            End If
            
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
            frmLabSamplingFilter.ShowMe Me, strFilter
            Me.rptPlist.Tag = strFilter
            If strFilter <> "" Then RefreshPatientData
            
        Case conMenu_View_PriceTable                                                '隐藏预置条码
            Control.Checked = Not Control.Checked
            HideBarCode
                    
        Case conMenu_View_Refresh                                                   '刷新
            Me.rptPlist.Tag = ""
            RefreshPatientData
        
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
            mlngDeptID = cboCtrol.ItemData(cboCtrol.ListIndex)
            RefreshPatientData
        Case conMenu_IDkind_Change
            Call IdKindChange
        Case Else

            If Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99 Then

                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            Else
                Call mobjZLIHISPlugIn.ExecuteFunc(glngSys, glngModul, Control.Caption, mlngKey, txtID, mstrBarCodes)
            End If
        
    End Select
End Sub

Private Sub IdKindChange()
    If Me.ActiveControl Is txtGoto Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5就诊卡, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.Visible = False Then Exit Sub
    If Me.stbThis.Visible = True Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    If Me.Visible = False Then Exit Sub
        
    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_ExportToXML         '清单打印相关
            Control.Enabled = (Me.rptPlist.Records.Count <> 0)
        
        Case conMenu_View_ToolBar_Button:                                                                   '按钮
            Control.Checked = Me.cbrthis(2).Visible
        
        Case conMenu_View_ToolBar_Text:                                                                     '按钮文字
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        
        Case conMenu_View_ToolBar_Size:                                                                     '大图标
            Control.Checked = Me.cbrthis.Options.LargeIcons
        
        Case conMenu_View_StatusBar                                                                         '状态栏
            Control.Checked = Me.stbThis.Visible
            
        Case conMenu_File_RowPrint, conMenu_File_BatPrint                                                   '条码和回执单打印
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index > 0 And Me.TabCtr.Selected.Index <> 4)
        
        Case conMenu_Manage_RequestView                                                                     '条码绑定
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index <= 1)
            Select Case Me.TabCtr.Selected.Index
                Case 0
                    Control.Caption = "绑定条码(&B)"
                Case 1, 2
                    Control.Caption = "清除绑定(&B)"
            End Select
        Case conMenu_Manage_Transfer_Force
            Control.Enabled = InStr(mstrPrivs, "参数设置") > 0
            
        Case conMenu_Manage_RequestPrint                                                                    '生成条码
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index <= 1)
            Select Case Me.TabCtr.Selected.Index
                Case 0
                    Control.Caption = "生成条码(&N)"
                Case 1, 2
                    Control.Caption = "取消条码(&N)"
            End Select
            
        Case conMenu_Manage_RequestBatPrint                                                                 '完成采集
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index >= 1)
            Select Case Me.TabCtr.Selected.Index
                Case 0, 1
                    Control.Caption = "完成采集(&P)"
                Case 2
                    Control.Caption = "取消完成(&P)"
                Case 3
                    Control.Caption = "取消完成(&P)"
                    Control.Enabled = False
            End Select
        Case conMenu_Tool_SignNew, conMenu_Edit_SendOther                                                   '使用39码或128码
            Control.Checked = Control.Checked
            Control.Enabled = InStr(mstrPrivs, "设置条码打印格式") > 0
            
        Case conMenu_Manage_Transfer_Force                                                                  '不区分科室打印
            Control.Checked = Control.Checked
            
        Case conMenu_View_PriceTable                                                                        '隐藏预置条码
            Control.Checked = Control.Checked
    End Select
    
    '提醒拒收标本
    On Error Resume Next
    If mstrOldTime = "" Then
        mstrOldTime = Now
    End If
    If DateDiff("n", mstrOldTime, Now) >= 1 Then
        showJuShouPait (1)
    End If
End Sub

Private Sub showJuShouPait(ByVal intType As Integer)
    '提醒拒收标本
    'intType 显示/关闭提示框  1=显示,0=关闭
    On Error Resume Next
    Dim PopupItem As PopupControlItem
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strJuShou As String
    Dim var_Tmp As Variant
    Dim lngItemTop As Long
            
    For intLoop = 0 To Me.rptPlist.Rows.Count - 1
        If Val(Me.rptPlist.Rows(intLoop).Record(mPcol.拒收).Value) > 0 Then
            intCount = intCount + Val(Me.rptPlist.Rows(intLoop).Record(mPcol.拒收).Value)
            strJuShou = strJuShou & "|" & Me.rptPlist.Rows(intLoop).Record(mPcol.病人姓名).Value & "," & Val(Me.rptPlist.Rows(intLoop).Record(mPcol.拒收).Value)
        End If
    Next
    If strJuShou <> "" Then strJuShou = Mid(strJuShou, 2)
    var_Tmp = Split(strJuShou, "|")
    If intCount > 0 Then
        If mstrOldTime = "" Then
            mstrOldTime = Now
        End If
        If intType = 1 Then
            With Me.PopupControl
                .RemoveAllItems
                '加载各个项目
                For intLoop = 0 To UBound(var_Tmp)
                    lngItemTop = (.Height / intCount) + (intLoop * 17)
                    Set PopupItem = .AddItem(45, lngItemTop, 300, 200, Split(var_Tmp(intLoop), ",")(0) & " 有" & Split(var_Tmp(intLoop), ",")(1) & "个标本被拒收,点击查看")
                    PopupItem.TextColor = vbRed
                Next
                .VisualTheme = xtpPopupThemeOffice2003
                .SetSize 300, 120
                .Font.Bold = True
                .Animation = xtpPopupAnimationSlide
                .ShowDelay = 9999999
                .Show
            End With
            mstrOldTime = ""
            If intCount = 0 Then
                Me.Caption = "检验采集工作站"
            Else
                Me.Caption = "检验采集工作站(你有" & intCount & "个拒收标本，请查收！)"
            End If
        End If
    Else
        With Me.PopupControl
            .Close
        End With
        Me.Caption = "检验采集工作站"
    End If
End Sub

Private Sub chkfilter_Click(Index As Integer)
    Call RefreshPatientData
End Sub

Private Sub cmdBakBillPrint_Click()
    '只打印回执单
    If Me.cmdBakBillPrint.Caption = "回执单打印" Then
        WriterBarCode 4, False, False, True
    Else
        WriterBarCode 3, True, True, False
        Call showJuShouPait(1)
    End If
End Sub

Private Sub CmdBarCodePrint_Click()
    '只打印条码
    WriterBarCode 4, False, True, False
    
End Sub

Private Sub cmdBindBarCode_Click()
        
    If Me.cbo性别.Tag = "新增" Then
        If Not ValidAdvice Then Exit Sub
        
        mlngKey = SaveAdviceData
        If mlngKey = 0 Then
            MsgBox "更新医嘱失败!", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If mlng病人ID <> 0 Then
            Me.txtGoto.Text = "-" & mlng病人ID
            Call txtGoto_KeyPress(13)
        End If
    End If
    
    
    
    
    If Me.cmdBindBarCode.Caption = "绑定条码(&B)" Then
        '费用确认
        If CheckMoeny = False Then
    '        MsgBox "确认费用不成功！", gstrSysName
            Exit Sub
        End If
        WriterBarCode 0, IIf(chkComPlete.Value = 1, True, False), IIf(ChkBarCodePrint.Value = 1, True, False)
    Else
        WriterBarCode 2, False
    End If
    ' 刷新病人信息
    If Not rptPlist.FocusedRow Is Nothing And rptAlist(TabCtr.Selected.Index).Rows.Count = 0 And optFilter(TabCtr.Selected.Index + 1).Value = True Then
        rptPlist.Records(rptPlist.FocusedRow.Record.Index).DeleteAll
        rptPlist.Rows(rptPlist.FocusedRow.Index).Record.Visible = False
        rptPlist.Populate
    End If
End Sub

Private Sub cmdComplete_Click()
    Dim strItem As String
    Dim intLoop As Integer
    If Me.cmdComplete.Caption = "完成采集(&P)" Then
        WriterBarCode 3, True, IIf(chkPrintBarCode.Value = 1, True, False), IIf(chkBackBill.Value = 1, True, False)
    Else
        '提示
        With Me.rptAlist(Me.TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.选择).Checked = True Then
                    strItem = strItem & vbCrLf & .Records(intLoop).Item(mAcol.医嘱内容).Value
                End If
            Next
        End With
        If strItem <> "" Then
            If MsgBox("是否确定要取消下面医嘱的条码绑定?" & strItem, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            WriterBarCode 2, False
        Else
            MsgBox "没有找到可以操作的医嘱!", vbInformation, Me.Caption
        End If
    End If
    ' 刷新病人信息
    If Not rptPlist.FocusedRow Is Nothing And rptAlist(TabCtr.Selected.Index).Rows.Count = 0 And optFilter(TabCtr.Selected.Index + 1).Value = True Then
        rptPlist.Records(rptPlist.FocusedRow.Record.Index).DeleteAll
        rptPlist.Rows(rptPlist.FocusedRow.Index).Record.Visible = False
        rptPlist.Populate
    End If
End Sub

Private Sub cmdNewBarcode_Click()
    Dim strItem As String
    Dim intLoop As Integer
    Dim strIDs As String
    Dim strName As String
    Dim blnPrint As Boolean
    
    Dim rsSampleCode As Recordset
    Dim strSampleCode As String
    Dim strSQL As String
    
    If Me.cbo性别.Tag = "新增" Then
        If Not ValidAdvice Then Exit Sub
        
        mlngKey = SaveAdviceData
        If mlngKey = 0 Then
            MsgBox "更新医嘱失败!", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If mlng病人ID <> 0 Then
            Me.txtGoto.Text = "-" & mlng病人ID
            Call txtGoto_KeyPress(13)
        End If
    End If
    
    If Me.cmdNewBarcode.Caption = "生成条码(&N)" Then
        '费用确认
        If CheckMoeny = False Then
    '        MsgBox "确认费用不成功！", gstrSysName
            Exit Sub
        End If
        WriterBarCode 1, IIf(chkComPlete.Value = 1, True, False), _
                         IIf(ChkBarCodePrint.Value = 1, True, False), _
                         IIf(chkBackBill.Value = 1, True, False)
    ElseIf Me.cmdNewBarcode.Caption = "送检标本(&C)" Or Me.cmdNewBarcode.Caption = "取消送检(&C)" Then

        With Me.rptAlist(Me.TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.选择).Checked = True Then
                    strIDs = strIDs & .Records(intLoop).Item(mAcol.ID).Value & "," & .Records(intLoop).Item(mAcol.合并医嘱).Value
                End If
            Next
            strIDs = Replace(Replace(strIDs, ";", ","), "|", ",")
            
            If Me.cmdNewBarcode.Caption = "送检标本(&C)" Then
                '检查tat超时
                If getTATTime(strIDs) = False Then
                    Exit Sub
                End If
                If strIDs = "" Then
                    Exit Sub
                End If
                  
            End If
            
            If strIDs = "" Then
                MsgBox "没有找到可以操作的医嘱记录!", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '保存送出时间
            If strIDs <> "" Then
                
                If Me.cmdNewBarcode.Caption = "送检标本(&C)" And chkSendPrint.Value <> 1 Then
                    If frmLabSamplingSendInfo.ShowMe(Me, strName, blnPrint) = False Then
                        Exit Sub
                    End If
                End If
                If strName = "" Then
                    strName = UserInfo.姓名
                End If
                

                '生成发送批号
                strSQL = "select 病人医嘱发送_标本发送批号.NEXTVAL  from dual"
                Set rsSampleCode = zlDatabase.OpenSQLRecord(strSQL, "标本发送批号", "")
                
                gstrSql = "Zl_Lis预置条码_标本送出('" & strIDs & "'" & IIf(Me.cmdNewBarcode.Caption = "取消送检(&C)", ",1", ",0") & _
                          ",'" & strName & "','" & rsSampleCode(0) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                '写入送检时间到检验申请单中
                Call WriterSampleSendDateToLIS(strIDs, IIf(Me.cmdNewBarcode.Caption = "取消送检(&C)", "1", "0"), strName)
            End If
            
            If Me.cmdNewBarcode.Caption = "送检标本(&C)" And chkSendPrint.Value <> 1 Then
            
'                If MsgBox("是否打印送出清单?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                 If blnPrint = True Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me, "医嘱字串=" & strIDs, 2)
                 End If
            End If
            RefreshPatientData
    
            Me.TxtBarCode.Text = ""
            Me.TxtBarCode.Tag = ""
            Me.TxtBarCodeCheck.Text = ""
            Me.txtGoto.SetFocus
        End With
    ElseIf Me.cmdNewBarcode.Caption = "取消条码(&N)" Then
        With Me.rptAlist(Me.TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.选择).Checked = True Then
                    strItem = strItem & vbCrLf & .Records(intLoop).Item(mAcol.医嘱内容).Value
                End If
            Next
        End With
        If MsgBox("是否确定要取消下面医嘱内容的条码?" & strItem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            WriterBarCode 2, False
        End If
    ElseIf Me.cmdNewBarcode.Caption = "让步检验(&N)" Then
        WriterBarCode 3, True, True, False, 1
    End If
    
    If Me.cbo性别.Tag = "新增" Then
        Me.txt姓名.SetFocus
        Me.cbo性别.Tag = ""
    End If
    ' 刷新病人信息
    If Not rptPlist.FocusedRow Is Nothing And rptAlist(TabCtr.Selected.Index).Rows.Count = 0 And optFilter(TabCtr.Selected.Index + 1).Value = True Then
        rptPlist.Records(rptPlist.FocusedRow.Record.Index).DeleteAll
        rptPlist.Rows(rptPlist.FocusedRow.Index).Record.Visible = False
        rptPlist.Populate
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim strExtData As String
    Dim rsTmp As New ADODB.Recordset
    
    strExtData = frmLabSamplingSelect.ShowMe(Me, mlngDeptID)
    If strExtData <> "" Then
        '获取采集方式
        Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
        If rsTmp Is Nothing Then
            MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
            Exit Sub
        End If
        mlngCapID = rsTmp("ID")
        Call AdviceSet检查手术(3, strExtData)
        txt医嘱内容.Text = Get检查手术名称(2, "")
        txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(strExtData, ";")(1) & ")"
    End If
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case mDkp.条码操作
            Item.Handle = Me.picBarCodeWork.hWnd
        Case mDkp.医嘱列表
            Item.Handle = picAdvice.hWnd
        Case mDkp.病人列表
            Item.Handle = Me.picTab.hWnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    IDKind.ActiveFastKey
End Sub

Private Sub Form_Load()
    Dim intItem As Integer
    Dim bln参数设置 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strIndex As String
    
     On Error Resume Next
    '增加插件菜单
    If mobjZLIHISPlugIn Is Nothing Then
        Set mobjZLIHISPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not mobjZLIHISPlugIn Is Nothing Then
            Call mobjZLIHISPlugIn.Initialize(gcnOracle, glngSys, glngModul)
        End If
    End If
     On Error GoTo errH
    
    '---------------------------------------
    '界面初始化
    mstrPrivs = gstrPrivs                                       '初使化权限
    CreateCbs                                                   '创建工具条
    CreateDkp                                                   '创建窗格
    CreateListHead                                              '创建表头
    CreateTab                                                   '创建列表
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
    '---------------------------------------
    '---------------------------------------
    bln参数设置 = InStr(";" & mstrPrivs & ";", ";参数设置;")
    ChkBarCodePrint = zlDatabase.GetPara("生成条码后打印", 100, 1211, 1, Array(ChkBarCodePrint), bln参数设置)
    chkComPlete = zlDatabase.GetPara("生成后标记为已完成", 100, 1211, 0, Array(chkComPlete), bln参数设置)
    chkBackBill = zlDatabase.GetPara("已完成后打印回执单", 100, 1211, 0, Array(chkBackBill), bln参数设置)
    chkDeptShow = zlDatabase.GetPara("只显示当前采集科室试管", 100, 1211, 0, Array(chkDeptShow), bln参数设置)
    ChkContinuous = zlDatabase.GetPara("连续输入", 100, 1211, 0, Array(ChkContinuous), bln参数设置)
    chkFindMove = zlDatabase.GetPara("查找病人后光标移动", 100, 1211, 0, Array(chkFindMove), bln参数设置)
    chkPrintBarCode = zlDatabase.GetPara("已完成后打印条码", 100, 1211, 0, Array(chkPrintBarCode), bln参数设置)
    chkSendPrint = zlDatabase.GetPara("取消送检单打印", 100, 1211, 0, Array(chkSendPrint), bln参数设置)
    chkBindPage = zlDatabase.GetPara("跳转到已绑定页", 100, 1211, 0, Array(chkBindPage), bln参数设置)
    chkApplyDept = zlDatabase.GetPara("生成条码时区分申请科室", 100, 1211, 0, Array(chkApplyDept), bln参数设置)
    chkMaterial = zlDatabase.GetPara("自动发料退料", 100, 1211, 0, Array(chkMaterial), bln参数设置)
    mblnNowConsumption = zlDatabase.GetPara("项目执行前必须先收费或先记帐审核", 100, , False)
    intItem = zlDatabase.GetPara("病人信息过滤", 100, 1211, 0)
    Me.optFilter(intItem).Value = True
    
    '---------------------------------------
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    
    mbln身份证 = False
    
    '数据读入
    Call GetDept                                                '读入科室
    Call InitDepts                                              '读入申请科室
    
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind初始化失败!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If
    
    RefreshPatientData                                          '读入数据
    
    
        If mobjLisInsideComm Is Nothing Then
            Dim strErr As String
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

    
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, , txtGoto)
    '62760;保存病人查找方式
    strIndex = zlDatabase.GetPara("病人查找方式", 100, 1211, 0)
    IDKind.IDKind = CInt(Val(strIndex))
    Exit Sub
errH:
    MsgBox "初始化窗体时发生错误，请检查部件后重试！", vbInformation, "初始化"
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane


    If Me.Visible = False Then Exit Sub



    Set Pane1 = Me.dkpMan.FindPane(mDkp.条码操作)
    Dim Control As CommandBarControl
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_PriceTable, True, True)

    On Error Resume Next

    Pane1.MaxTrackSize.SetSize 8145 / Screen.TwipsPerPixelX, IIf(Control.Checked, 5500, 6300) / Screen.TwipsPerPixelY + 15
    Pane1.MinTrackSize.SetSize 8145 / Screen.TwipsPerPixelX, IIf(Control.Checked, 5500, 6300) / Screen.TwipsPerPixelY

    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

    Pane1.MinTrackSize.SetSize 100, 100
    '当点击窗体最大化按钮时,让文本框获得焦点
    Select Case Me.WindowState
        Case vbMaximized    '最大化按钮
            Me.txtGoto.SetFocus
    End Select
End Sub
Private Sub CreateListHead()
    '创建列表头
    Dim Column As ReportColumn
    Dim intLoop As Integer
    
    '==病人列表头
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
        Set Column = .Add(mPcol.临床路径病人, "临床路径病人", 18, False): Column.Icon = 4
        Set Column = .Add(mPcol.来源, "来源", 30, True)
        Set Column = .Add(mPcol.病人姓名, "病人姓名", 75, True)
        Set Column = .Add(mPcol.性别, "性别", 30, True)
        Set Column = .Add(mPcol.年龄, "年龄", 40, True)
        Set Column = .Add(mPcol.病人科室, "病人科室", 75, True)
        Set Column = .Add(mPcol.标识号, "标识号", 60, True)
        Set Column = .Add(mPcol.床号, "床号", 60, True)
        Set Column = .Add(mPcol.挂号单, "挂号单", 60, True): Column.Visible = False
        Set Column = .Add(mPcol.就诊卡, "就诊卡", 60, True): Column.Visible = False
        Set Column = .Add(mPcol.紧急, "紧急", 30, True)
        Set Column = .Add(mPcol.未绑定, "未绑定", 45, True)
        Set Column = .Add(mPcol.已绑定, "已绑定", 45, True)
        Set Column = .Add(mPcol.已采样, "已采样", 45, True)
        Set Column = .Add(mPcol.已送检, "已送检", 45, True)
        Set Column = .Add(mPcol.拒收, "拒收", 30, True)
        Set Column = .Add(mPcol.重采标本, "重采", 30, True)
        Set Column = .Add(mPcol.已执行, "已执行", 45, True)
        Set Column = .Add(mPcol.合计, "合计", 45, True)
        Set Column = .Add(mPcol.发送时间, "发送时间", 75, True)
        Set Column = .Add(mPcol.状态, "状态", 30, False): Column.Visible = False
        Set Column = .Add(mPcol.条码, "条码", 30, False): Column.Visible = False
    End With
    
    '==医嘱列表头
    For intLoop = 0 To 5
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
            Set Column = .Add(mAcol.类别, "类别", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
            Set Column = .Add(mAcol.选择, "Check", 18, False): Column.Icon = 0
            Set Column = .Add(mAcol.急诊, "", 18, False): Column.Icon = 2
            Set Column = .Add(mAcol.图标, "", 18, False): Column.Icon = 3
            Set Column = .Add(mAcol.记录状态, "收费", 30, False): Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.重采, "重采", 30, False): Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.采集方式, "采集方式", 75, True)
            Set Column = .Add(mAcol.标本, "标本", 55, True)
            Set Column = .Add(mAcol.医嘱内容, "医嘱内容", 75, True)
            Set Column = .Add(mAcol.条码, "条码", 75, True)
            Set Column = .Add(mAcol.婴儿姓名, "婴儿", 75, True)
            Set Column = .Add(mAcol.执行科室, "执行科室", 75, True)
            Set Column = .Add(mAcol.开嘱医生, "开嘱医生", 75, True)
            Set Column = .Add(mAcol.开嘱时间, "开嘱时间", 75, True)
            Set Column = .Add(mAcol.发送人, "发送人", 65, True)
            Set Column = .Add(mAcol.发送时间, "发送时间", 75, True)
            Set Column = .Add(mAcol.采样时间, "采样时间", 75, True)
            Set Column = .Add(mAcol.试管颜色, "颜色编码", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.试管编码, "试管编码", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.采样人, "采样人", 60, True)
            Set Column = .Add(mAcol.NO, "单据号", 60, True)
            Set Column = .Add(mAcol.审核时间, "审核时间", 75, True)
            Set Column = .Add(mAcol.接收时间, "接收时间", 75, True)
            Set Column = .Add(mAcol.采血量, "采血量", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.试管名称, "试管名称", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.紧急, "紧急", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.病人来源, "病人来源", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.申请科室, "申请科室", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.婴儿, "婴儿", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.别名, "别名", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.相关ID, "相关ID", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.执行状态, "执行状态", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.接收人, "接收人", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.条码打印, "打印", 65, True)
            Set Column = .Add(mAcol.拒收理由, "拒收理由", 400, True)
            Set Column = .Add(mAcol.送出时间, "送出时间", 120, True)
            Set Column = .Add(mAcol.采集科室ID, "采集科室ID", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.采集执行科室, "采集执行科室", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.诊疗项目ID, "诊疗项目ID", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.诊疗项目组合, "诊疗项目组合", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.检验执行科室ID, "检验执行科室ID", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.计费状态, "计费状态", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.记录性质, "记录性质", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.病人所在科室, "病人所在科室", 120, False): Column.Visible = False
        End With
    Next
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
        Set Column = .Add(mCuvette.编码, "编码", 55, True)
        Set Column = .Add(mCuvette.名称, "名称", 80, True)
        Set Column = .Add(mCuvette.添加剂, "添加剂", 90, True)
        Set Column = .Add(mCuvette.采血量, "采血量", 60, True)
        Set Column = .Add(mCuvette.规格, "规格", 60, True)
        Set Column = .Add(mCuvette.颜色, "", 18, True): Column.Icon = 3
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Controlcbo As CommandBarComboBox
    Dim Control As CommandBarControl
    Dim strTmp As String
    Dim aItem() As String
    Dim intLoop As Integer
    
    Call SaveWinState(Me, App.ProductName)
'    Me.Visible = False


    zlDatabase.SetPara "生成条码后打印", ChkBarCodePrint, 100, 1211
    zlDatabase.SetPara "生成后标记为已完成", chkComPlete, 100, 1211
    zlDatabase.SetPara "已完成后打印回执单", chkBackBill, 100, 1211
    zlDatabase.SetPara "连续输入", ChkContinuous, 100, 1211
    zlDatabase.SetPara "查找病人后光标移动", chkFindMove, 100, 1211
    zlDatabase.SetPara "已完成后打印条码", chkPrintBarCode, 100, 1211
    zlDatabase.SetPara "只显示当前采集科室试管", chkDeptShow, 100, 1211
    zlDatabase.SetPara "取消送检单打印", chkSendPrint, 100, 1211
    zlDatabase.SetPara "跳转到已绑定页", chkBindPage, 100, 1211
    zlDatabase.SetPara "生成条码时区分申请科室", chkApplyDept, 100, 1211
    zlDatabase.SetPara "自动发料退料", chkMaterial, 100, 1211
    '保存科室ID已便下次使用
    'Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    zlDatabase.SetPara "科室", Controlcbo.ItemData(Controlcbo.ListIndex), 100, 1211
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Tool_SignNew, True, True)
    zlDatabase.SetPara "使用条码", IIf(Control.Checked, 1, 2), 100, 1211
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
    zlDatabase.SetPara "不区分执行科室打印", IIf(Control.Checked, 0, 1), 100, 1211
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_PriceTable, True, True)
    zlDatabase.SetPara "隐藏预置条码", IIf(Control.Checked, 1, 0), 100, 1211

    For intLoop = 0 To Me.optFilter.Count - 1
        If Me.optFilter(intLoop).Value = True Then
            zlDatabase.SetPara "病人信息过滤", intLoop, 100, 1211
            Exit For
        End If
    Next
    '62760;保存病人查找方式
    zlDatabase.SetPara "病人查找方式", mstrIndex, 100, 1211
    
    '把间隔时间恢复为最近3天的
    strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")
    If strTmp <> "" Then
        aItem = Split(strTmp, ";")
        strTmp = ""
        For intLoop = 0 To UBound(aItem)
            If intLoop = mFilter.间隔时间 Then
                strTmp = strTmp & ";" & "3"
            ElseIf intLoop = mFilter.开始时间 Then
                strTmp = strTmp & ";"
            ElseIf intLoop = mFilter.结束时间 Then
                strTmp = strTmp & ";"
            Else
                strTmp = strTmp & ";" & aItem(intLoop)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        zlDatabase.SetPara "采集工作站过滤", strTmp, 100, 1211
    End If
    Set mobjSquareCard = Nothing
    Set mobjLisInsideComm = Nothing
    imgPatient.Picture = Nothing
'    If Not mobjIDCard Is Nothing Then
'        Call mobjIDCard.SetEnabled(False)
'    End If
'    If Not mobjICCard Is Nothing Then
'        Call mobjICCard.SetEnabled(False)
'    End If
    
    

'    Set mobjIDCard = Nothing
'    Set mobjICCard.gcnOracle = Nothing
'    Set mobjICCard = Nothing
'    Me.TabCtr.RemoveAll
'    Me.cbrthis.DeleteAll
'    Me.dkpMan.CloseAll
'    Me.dkpMan.DestroyAll

End Sub






Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
'    objPatiInfor.病人ID = 216178
    If objPatiInfor.病人ID <> 0 Then
        txtGoto.Text = "-" & objPatiInfor.病人ID
    ElseIf objPatiInfor.病人ID = 0 Then
        txtGoto.Text = objPatiInfor.卡号
    End If
    Call txtGoto_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
' 2007-08-17 增加一卡通支持
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

Private Sub optFilter_Click(Index As Integer)
    If Index <> 0 Then
        TabCtr.Item(Index - 1).Selected = True
    End If
    RefreshPatientData
End Sub



Private Sub CreateTab()
    Dim Item As TabControlItem
    
    With Me.TabCtr
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 0, "未绑定", Me.rptAlist(0).hWnd, 0
        .InsertItem 1, "已绑定", Me.rptAlist(1).hWnd, 0
        .InsertItem 2, "已采样", Me.rptAlist(2).hWnd, 0
        .InsertItem 3, "已送检", Me.rptAlist(3).hWnd, 0
        .InsertItem 4, "已执行", Me.rptAlist(4).hWnd, 0
        .InsertItem 5, "拒收", Me.rptAlist(5).hWnd, 0
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub picAdvice_Resize()
    With TabCtr
        .Top = 0
        .Left = 0
        .Width = Me.picAdvice.ScaleWidth
        .Height = Me.picAdvice.ScaleHeight
    End With
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
'    With srtFilter
'        .Top = 0
'        .Left = 0
'        .Width = Me.picFilter.Width
'        .Height = Me.picFilter.Height
'    End With
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Me.srtPatient.Top = 0
    Me.srtPatient.Left = 0
    Me.srtPatient.Width = Me.picTab.ScaleWidth
    
    Me.picFilter.Top = Me.srtPatient.Top + Me.srtPatient.Height + 5
    Me.picFilter.Left = 0
    Me.picFilter.Width = Me.ScaleWidth
    
    Me.rptPlist.Top = Me.picFilter.Top + Me.picFilter.Height + 10
    Me.rptPlist.Left = 0
    Me.rptPlist.Width = Me.picTab.ScaleWidth
    Me.rptPlist.Height = Me.picTab.ScaleHeight - picFilter.Top - picFilter.Height
End Sub


Private Sub PopupControl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    '点击右下角弹出提示框中的项目时,选中病人列表中的病人
    Dim strPaitName As String
    Dim rptRow As ReportRow
    
    Me.TabCtr.Item(5).Selected = True
    strPaitName = Mid(Item.Caption, 1, InStr(Item.Caption, " ") - 1)
    With Me.rptPlist
        For Each rptRow In .Rows
            If rptRow.Record(mPcol.病人姓名).Value = strPaitName Then
                .FocusedRow = rptRow
                Exit For
            End If
        Next
    End With
End Sub

Private Sub rptAlist_ItemCheck(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim RecordC As ReportRecord
    Dim RecordA As ReportRecord
    Dim strResivePeo
    If Me.TabCtr.Selected.Index = 0 Then Exit Sub
    For Each RecordC In Me.rptAlist(Index).Records
        If RecordC(mAcol.条码).Value = Item.Record(mAcol.条码).Value Then
            RecordC(mAcol.选择).Checked = Item.Checked
        End If
        If RecordC(mAcol.选择).Checked = True Then
            If RecordC(mAcol.接收人).Value <> "" Then
                strResivePeo = RecordC(mAcol.接收人).Value
            End If
        End If
    Next
    If strResivePeo <> "" Then
        cmdComplete.Enabled = False
    Else
        cmdComplete.Enabled = True
    End If
    Me.rptAlist(Index).Populate
End Sub

Private Sub rptAlist_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptAlist(Index)
        Set hitColumn = .HitTest(X, Y).Column
        
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                hitColumn.AutoSize = True
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mAcol.选择).Checked
                For Each Record In .Records
                    Record.Item(mAcol.选择).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub


Private Sub rptAlist_SelectionChanged(Index As Integer)
    With Me.rptAlist(Me.TabCtr.Selected.Index)
        If Not .FocusedRow Is Nothing And .FocusedRow.GroupRow = False Then
            .PaintManager.HighlightBackColor = Val(.FocusedRow.Record(mAcol.试管颜色).Value)
            .Populate
            '下面语句出错出时不处理
            On Error Resume Next
            Me.cbo开单科室.Text = .FocusedRow.Record(mAcol.申请科室).Value
            Me.cbo医生.Text = .FocusedRow.Record(mAcol.开嘱医生).Value
            Me.txt医嘱内容.Text = .FocusedRow.Record(mAcol.医嘱内容).Value
            txtPatientDept.Text = .FocusedRow.Record(mAcol.病人所在科室).Value
            On Error GoTo 0
        Else
            .PaintManager.HighlightBackColor = vbWhite
        End If
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

Private Sub rptPlist_GotFocus()
    Me.dkpMan.RecalcLayout
End Sub

Private Sub RefreshPatientData()
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
    Dim str检验类型 As String, strSample As String, strCapture As String
    Dim NowDate As Date                                 '当前时间
    Dim intPatientType As Integer                       '病人来源
    Dim blnPathPatient As Boolean                       '临床路径病人
    Dim intPatientPage As Integer
    Dim strBloodSQL As String                           '输血查询语句
    On Error GoTo errH
    
    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")
    
    NowDate = zlDatabase.Currentdate
    
    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    lblGoto.Caption = "病人查找"
    
    zlCommFun.ShowFlash "正在更新数据,请稍候...", Me
    
    strSQL = "Select 病人ID,病人来源,病人姓名,病人科室,性别,年龄,就诊卡号,标识号,当前床号,主页ID,Sum(decode(状态,'已完成',1,'执行中',1,0)) As 已执行," & vbNewLine & _
                "       Sum(decode(状态,'拒收',1,0)) As 拒收,Sum(decode(状态,'未绑定',1,0)) As 未绑定," & vbNewLine & _
                "       Sum(decode(状态,'已绑定',1,0)) As 已绑定,Sum(decode(状态,'已采样',1,0)) As 已采样, " & vbNewLine & _
                "       Sum(decode(紧急,'紧急',1,0)) as 紧急,Sum(重采标本) as 重采标本,sum(decode(状态,'已送检',1,0)) as 已送检,max(发送时间) as 发送时间,临床路径病人" & vbNewLine & _
                "From ("
    strSQL = strSQL & "Select    distinct a.病人id,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源, " & vbCrLf & _
             " c.姓名 as 病人姓名,e.名称 as 病人科室,c.性别,c.年龄,c.就诊卡号,b.样本条码, " & vbCrLf & _
             " decode(b.执行状态,1,'已完成',2,'拒收',3,'执行中', " & vbCrLf & _
             "         Decode(b.样本条码, Null, '未绑定', Decode(b.采样人, Null, '已绑定', decode(b.标本送出时间,null,'已采样','已送检'))) )as 状态, " & vbCrLf & _
             " decode(A.病人来源, 1, C.门诊号, 2, C.住院号,4,c.门诊号) As 标识号, " & vbCrLf & _
             " decode(c.当前床号,null,decode(l.出院病床,null,l.入院病床,l.出院病床),c.当前床号) as 当前床号 , " & vbCrLf & _
             " decode(a.紧急标志,1,'紧急',decode(g.急诊,1,'紧急')) as 紧急 ,decode(a.病人来源 , 2,a.主页ID,0) 主页ID, " & vbCrLf & _
             " decode(b.执行状态,0,'',2,'拒收') as 拒收,b.执行状态,nvl(b.重采标本,0) as 重采标本,b.发送时间,nvl(s.路径状态,0) as 临床路径病人,a.医嘱内容 " & vbCrLf & _
             " From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 部门表 E, 诊疗项目目录 F,病人挂号记录 G,病人医嘱记录 H, " & vbCrLf & _
             "      诊疗项目目录 K ,病案主页 L,病人医嘱发送 M,检验标本记录 J,病案主页 S " & vbCrLf & _
             " Where A.ID = H.相关ID And H.id = B.医嘱id And A.病人id = C.病人id And A.病人科室id = E.ID And A.诊疗项目id+0 = f.ID  " & vbCrLf & _
             "      And h.诊疗项目ID = k.id and a.id = j.医嘱id(+)  " & vbCrLf & _
             " And A.挂号单 = G.No(+) and a.病人id = g.病人id(+)  and a.病人id = g.病人id(+)  and (g.病人ID is null or (g.记录状态 =1 and g.记录性质 =1) ) And  f.类别 = 'E' And f.操作类型 = '6' and a.病人id = l.病人ID(+) and a.主页ID = l.主页ID(+) and m.执行部门id + 0 = [1] " & vbCrLf & _
             " And A.ID = M.医嘱ID And k.试管编码 is not null and a.病人ID = S.病人ID(+) and a.主页ID = s.主页ID(+) " & IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "")
                 
    
    If Me.rptPlist.Tag <> "" Then
        strSQL = strSQL & " And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3", "0") & "," & _
                 Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "

        If varFilter(mFilter.标识号) <> "" Then
            strSQL = strSQL & " And decode(a.病人来源,2,c.住院号,c.门诊号) = [2] "
        End If
        
        If varFilter(mFilter.就诊卡) <> "" Then
            strSQL = strSQL & " And c.就诊卡号 = [3] "
            
        End If
        
        If varFilter(mFilter.姓名) <> "" Then
            strSQL = strSQL & " And C.姓名 like [4] "
            
        End If
        
        If varFilter(mFilter.单据号) <> "" Then
            strSQL = strSQL & " and B.NO = [5]"
            
        End If
        
        If UBound(varFilter) >= mFilter.标本 Then
            If Trim(varFilter(mFilter.标本)) <> "" Then
                strSQL = strSQL & " And instr([6],','||H.标本部位||',') > 0 "
                
            End If
        End If
        
        If UBound(varFilter) >= mFilter.采集方式 Then
            If Trim(varFilter(mFilter.采集方式)) <> "" Then
                strSQL = strSQL & " And instr([7],','|| f.ID ||',') > 0 "
                
            End If
        End If
        
        If varFilter(mFilter.发送或审核时间) = 0 Then
            strSQL = strSQL & " and m.发送时间 Between [8] and [9]"
            
        Else
            strSQL = strSQL & " and m.发送时间 Between [8] and [9]"
        End If
        
        If varFilter(mFilter.开始时间) = "" Then
            strDateBegin = NowDate - Val(varFilter(mFilter.间隔时间))
            strDateEnd = NowDate
        Else
            strDateBegin = varFilter(mFilter.开始时间)
            strDateEnd = varFilter(mFilter.结束时间)
        End If
        If UBound(varFilter) >= mFilter.检验类型 Then
            If Trim(varFilter(mFilter.检验类型)) <> "" Then
                lblGoto.Caption = "病人查找(" & Mid(Trim(varFilter(mFilter.检验类型)), 2) & ")"
                strSQL = strSQL & " And instr([10],','|| k.操作类型 ||',') > 0 "
                
            End If
        End If
    Else
        If strTmp <> "" Then
            strSQL = strSQL & " And instr('" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3", "0") & "," & _
                 Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & "',A.病人来源)>0 "

    
            If UBound(varFilter) >= mFilter.标本 Then
                If Trim(varFilter(mFilter.标本)) <> "" Then
                    strSQL = strSQL & " And instr([6],','||H.标本部位||',') > 0 "
                    
                End If
            End If
            
            If UBound(varFilter) >= mFilter.采集方式 Then
                If Trim(varFilter(mFilter.采集方式)) <> "" Then
                    strSQL = strSQL & " And instr([7],','|| f.ID ||',') > 0 "
                    
                End If
            End If
            
            If varFilter(mFilter.发送或审核时间) = 0 Then
                strSQL = strSQL & " and m.发送时间 Between [8] and [9]"
                
            Else
                strSQL = strSQL & " and m.发送时间 Between [8] and [9]"
                
            End If
            
            If Val(varFilter(mFilter.间隔时间)) >= 0 Then
                strDateBegin = NowDate - Val(varFilter(mFilter.间隔时间))
                strDateEnd = NowDate
            Else
                strDateBegin = varFilter(mFilter.开始时间)
                strDateEnd = varFilter(mFilter.结束时间)
            End If
            If UBound(varFilter) >= mFilter.检验类型 Then
                If Trim(varFilter(mFilter.检验类型)) <> "" Then
                    lblGoto.Caption = "病人查找(" & Mid(Trim(varFilter(mFilter.检验类型)), 2) & ")"
                    strSQL = strSQL & " And instr([10],','|| k.操作类型 ||',') > 0 "
                    
                End If
            End If
        Else
            strSQL = strSQL & " and m.发送时间 Between [8] and [9]"
            
            strDateBegin = NowDate - 3
            strDateEnd = NowDate
        End If
    End If
    
    strBloodSQL = GetBooldPatientDataSql
    strSQL = strSQL & " union all " & strBloodSQL
    strSQL = strSQL & ") a group by 病人Id,病人来源,病人姓名,病人科室,性别,年龄,就诊卡号,标识号,当前床号,主页ID,临床路径病人 "
    
    blnDateMoved = MovedByDate(CDate(strDateBegin)) '按时间看是否可能已转出
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    strSQL = strSQL & " Order by 病人科室 "
    blnPathPatient = False
    If strTmp = "" And Me.rptPlist.Tag = "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, "", "", "", "", "", "", _
                    CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), "")
    Else
        If UBound(varFilter) >= mFilter.检验类型 Then
            str检验类型 = varFilter(mFilter.检验类型) & ","
        End If
        If UBound(varFilter) >= mFilter.标本 Then
            strSample = varFilter(mFilter.标本) & ","
        End If
        If UBound(varFilter) >= mFilter.采集方式 Then
            strCapture = varFilter(mFilter.采集方式) & ","
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, Val(varFilter(mFilter.标识号)), CStr(varFilter(mFilter.就诊卡)) _
                    , CStr(varFilter(mFilter.姓名)) & "%", CStr(varFilter(mFilter.单据号)), strSample, strCapture _
                    , CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), _
                    CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), str检验类型)
    End If
    
    '清除记录
    Me.rptPlist.Records.DeleteAll
    Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
    Me.rptCuvette.Records.DeleteAll
    
    Do Until rsTmp.EOF
    
        Set Record = rptPlist.Records.Add
        For intLoop = 0 To Me.rptPlist.Columns.Count + 1
            Record.AddItem ""
        Next
        
        If Val(Nvl(rsTmp("临床路径病人"), 0)) = 1 Then
            blnPathPatient = True
            Record(mPcol.临床路径病人).Icon = 4
        Else
            Record(mPcol.临床路径病人).Icon = -1
        End If
        Record(mPcol.病人ID).Value = Nvl(rsTmp("病人ID"))
        Record(mPcol.来源).Value = Nvl(rsTmp("病人来源"))
        Record(mPcol.病人姓名).Value = Nvl(rsTmp("病人姓名"))
        Record(mPcol.病人科室).Value = Nvl(rsTmp("病人科室"))
        Record(mPcol.性别).Value = Nvl(rsTmp("性别"))
        Record(mPcol.年龄).Value = Nvl(rsTmp("年龄"))
        Record(mPcol.标识号).Value = Nvl(rsTmp("标识号"))
        Record(mPcol.床号).Value = Nvl(rsTmp("当前床号"))
        Record(mPcol.拒收).Value = Nvl(rsTmp("拒收"))
        Record(mPcol.就诊卡).Value = Nvl(rsTmp("就诊卡号"))
        
        Record(mPcol.紧急).Value = Nvl(rsTmp("紧急"))
        Record(mPcol.未绑定).Value = Nvl(rsTmp("未绑定"))
        Record(mPcol.已绑定).Value = Nvl(rsTmp("已绑定"))
        Record(mPcol.已采样).Value = Nvl(rsTmp("已采样"))
        Record(mPcol.已送检).Value = Nvl(rsTmp("已送检"))
        Record(mPcol.拒收).Value = Nvl(rsTmp("拒收"))
        Record(mPcol.重采标本).Value = Nvl(rsTmp("重采标本"))
        Record(mPcol.已执行).Value = Nvl(rsTmp("已执行"))
        Record(mPcol.发送时间).Value = Nvl(rsTmp("发送时间"))
        Record(mPcol.主页ID).Value = Nvl(rsTmp("主页ID"), 0)
        
        Record(mPcol.合计).Value = Val(rsTmp("未绑定")) + Val(rsTmp("已绑定")) + Val(rsTmp("已采样")) + Val(rsTmp("拒收")) + Val(rsTmp("已执行"))
        
        lngPatientID = Nvl(rsTmp("病人ID"))
        
        If Nvl(rsTmp("紧急")) > 0 Then
            For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                Record(intLoop).ForeColor = vbRed
            Next
        End If
        
        rsTmp.MoveNext
    Loop
    
    '更新
    Me.rptPlist.Populate
    Me.rptPlist.Columns(1).Visible = blnPathPatient

    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
    Me.rptCuvette.Populate
    
    If strTmp <> "" Then
    
        Me.stbThis.Panels(2).Text = "当前范围<" & IIf(varFilter(mFilter.发送或审核时间) = 0, "发送时间 ", "发送时间 ") & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> 下共有:" & _
                                Me.rptPlist.Rows.Count & "个病人."
    Else
        Me.stbThis.Panels(2).Text = "当前范围<" & "发送时间 " & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> 下共有:" & _
                                Me.rptPlist.Rows.Count & "个病人."
    End If
                                
    '定位到上次选中的病人
    If Me.Visible = True Then
        With Me.rptPlist
            
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mPcol.病人ID).Value = mlngKey Then
                    Set .FocusedRow = .Rows(intLoop)
                    mlngKey = .Rows(intLoop).Record(mPcol.病人ID).Value
                    intPatientPage = .Rows(intLoop).Record(mPcol.主页ID).Value
                    .Populate
    '                Me.rptPlist.Tag = ""
                    Exit For
                End If
            Next
            
            If .FocusedRow Is Nothing And .Rows.Count > 0 Then
                Set .FocusedRow = .Rows(0)
                intPatientType = IIf(.Rows(0).Record(mPcol.来源).Value = "住院", 2, 1)
                mlngKey = .Rows(0).Record(mPcol.病人ID).Value
                intPatientPage = .Rows(0).Record(mPcol.主页ID).Value
                .Populate
            End If
            
            If Not .FocusedRow Is Nothing Then
                RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index, intPatientType, False, intPatientPage
            End If
            
        End With
    End If
    '过滤中条件只执行一次
'    Me.rptPlist.Tag = ""
    
    If Me.rptPlist.Rows.Count = 0 Then
        txt姓名 = ""
        txt姓名.Tag = ""
        cbo性别.ListIndex = -1
        txt年龄 = ""
        txt年龄1 = ""
        txtBed = ""
        txtID = ""
        txtPatientDept = ""
        cbo开单科室.ListIndex = -1
        cbo医生.ListIndex = -1
        txt医嘱内容.Text = ""
        txt医嘱内容.Tag = ""
        Me.lblCap(6).Visible = False
    End If
    
    '过滤病人信息列表
    Call FilterPatient
    
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
    
    mlngDeptID = zlDatabase.GetPara("科室", 100, 1211, 0)
    
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    
    On Error GoTo errH
    
    strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(1,2,3,4) And B.工作性质 IN('检验','护理','治疗')"
            
    If InStr(1, mstrPrivs, "所有科室") <= 0 Then
        strSQL = strSQL & " And C.人员ID = [1] "
    End If
    
    strSQL = strSQL & " Order by A.编码"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Controlcbo.Clear
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
    
    '性别
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("性别")
    cbo性别.Clear
    If Not rsTmp Is Nothing Then
        For intLoop = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function RefreshAdviceData(lngPatientID As Long, intState As Integer, intPatientType As Integer, Optional blOnlyWhere As Boolean = False, Optional intPatientPage As Integer) As Boolean
    '功能：                         刷新采集医嘱记录
    '参数：                         lngpatientId = 病人ID ,
    '                               intPatientType = 病人来源
    '                               intState = 当前状态 0=未绑定 1=已绑定 2=已采样
    '                               blOnlyWhere = true 只使用病人ID进行查找
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
    Dim intShowButtom As Integer                    '用于在查找时显示那些按钮,条码 not null = 1 采样人 not null = 2 已执行 = 3
    Dim strNO As String                             'NO号
    Dim blnShowExec As Boolean                      '是否显示未收费门诊病人
    Dim strAge As String                            '年龄计算
    Dim aAge() As String
    Dim rsBaby As New ADODB.Recordset               '婴儿姓名查询
    Dim NowDate As Date                             '当前时间
    Dim strSQLbak As String
    Dim str医嘱内容 As String                       '医嘱内容
    Dim strAdvice  As String                        '记录在医嘱信息栏中已经勾选的医嘱
    Dim strReceivePeople As String                 '记录接收人
    Dim strSample As String, strCapture As String
    Dim blnFL As Boolean                            '是否分类显示,存在输血医嘱时,分类显示
    Dim strBooldSql As String                       '查询输血sql
    blnDateMoved = MovedByDate(Date) '按时间看是否可能已转出
                 
    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")
    
    NowDate = zlDatabase.Currentdate
    
    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    On Error GoTo errH
                 
    strSQL = " select /*+ rule */ distinct a.类别,a.医嘱id,a.相关id,a.试管颜色,a.采集方式,a.医嘱内容,a.样本条码,a.采样时间,a.执行科室,a.开嘱医生,a.开嘱时间,a.发送人,a.发送时间,a.试管编码," & vbCrLf & _
             "        a.标本,a.病人姓名,a.性别,a.年龄,a.床号,a.标识号,Decode(A.病人来源, 2, nvl(A.病人所在科室,A.申请科室), A.申请科室) As 病人所在科室,a.拒收,a.病人ID,a.采样人,a.采血量,a.试管名称,a.紧急,a.病人来源,a.婴儿,a.别名,a.执行状态,a.NO," & vbCrLf & _
             "        a.审核时间,a.申请科室,a.接收时间,a.组合项目,b.记帐费用,b.记录状态,a.接收人,b.记录性质,a.采集项目ID,a.就诊卡号,a.挂号单,a.条码打印,a.执行说明,a.重采, " & vbCrLf & _
             "        a.标本送出时间,a.主页ID,a.执行科室ID,a.采集执行科室,a.诊疗项目ID,a.检验执行科室ID,b.门诊标志,计费状态,结算模式" & vbCrLf & _
             " from "
    strSQL = strSQL & " ( Select decode(d.类别,'K','输血','检验') 类别, B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色, decode(d.类别,'K',b.医嘱内容 ,d.名称) As 采集方式, decode(d.类别,'K',d.名称,b.医嘱内容) as 医嘱内容, C.样本条码,C.采样时间, " & vbCrLf & _
             "   H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
             "   I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号,4,i.门诊号) as 标识号, " & vbCrLf & _
             "   L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人ID,c.采样人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
             "   DECODE(B.紧急标志,1,'紧急','') as 紧急,b.病人来源,nvl(b.婴儿,0) as 婴儿,N.名称 as 别名,decode(d.类别, 'K', M.执行状态,C.执行状态) 执行状态,C.NO,j.审核时间,o.名称 as 申请科室,m.接收时间, " & vbCrLf & _
             "   E.组合项目,C.接收人,c.记录性质,decode(d.类别,'K',e.id ,d.id)  as 采集项目ID,i.就诊卡号,b.挂号单,C.条码打印,C.执行说明,nvl(c.重采标本,0) as 重采,c.标本送出时间, " & vbCrLf & _
             "   a.主页ID,Decode(d.类别, 'K', b.执行科室ID, a.执行科室ID) 执行科室ID,P.名称 as 采集执行科室,b.诊疗项目ID,Decode(d.类别, 'K', a.执行科室ID, b.执行科室ID) as 检验执行科室ID,c.计费状态,i.结算模式 " & vbCrLf & _
             "   From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
             "   采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,检验标本记录 J ,部门表 O ,部门表 P, " & vbCrLf & _
             "   (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N " & vbCrLf & _
             "  Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID  " & vbCrLf & _
             "    And (e.类别 = 'E' Or e.类别 = 'C') And E.试管编码 = G.编码 And B.执行科室id = H.ID(+) and a.执行科室ID = P.id(+)  " & vbCrLf & _
             "    And  d.类别 = 'E'  And d.操作类型 = '6'  And A.病人id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.发送时间+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
             "    and  m.执行部门id + 0 = [2] And B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & vbCrLf & _
             "    and a.id = m.医嘱id And E.id = N.诊疗项目ID(+) and a.id = j.医嘱id(+) and b.开嘱科室id = o.id  " & vbCrLf & _
             "    ) a , (Select 医嘱序号,记录性质,记录状态,记帐费用,门诊标志 From 住院费用记录 Where  病人ID=[1]) b " & vbCrLf & _
             "where a.医嘱id = b.医嘱序号(+) and a.记录性质 = mod(b.记录性质(+),10)  "
    
               '  IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 between  to_date( '" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 00:00:00' ,'yyyy-mm-dd hh24:mi:ss' ) and to_date('" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 23:59:59','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _

    '处理三种不同的状态,
    If intState = 0 Then
        strSQL = strSQL & " And a.样本条码 is null And a.执行状态 in (0) " & vbCrLf
    ElseIf intState = 1 Then
        strSQL = strSQL & " And a.样本条码 is not null And a.采样时间 is  null And a.执行状态 in (0) " & vbCrLf
    ElseIf intState = 2 Then
        strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0) and a.标本送出时间 is null " & vbCrLf
    ElseIf intState = 3 Then
        strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0) and a.标本送出时间 is not null  " & vbCrLf
    ElseIf intState = 4 Then
        strSQL = strSQL & " And a.执行状态 in (1,3) " & vbCrLf
    ElseIf intState = 5 Then
        strSQL = strSQL & " And a.执行状态 in (2) " & vbCrLf
    End If
    
    '过滤
    If Me.rptPlist.Tag <> "" Or strTmp <> "" Then
       
        If UBound(varFilter) >= mFilter.标本 Then
            If Trim(varFilter(mFilter.标本)) <> "" Then
                strSQL = strSQL & " And instr([5],','||a.标本||',') > 0 "
            End If
        End If
        
        If UBound(varFilter) >= mFilter.采集方式 Then
            If Trim(varFilter(mFilter.采集方式)) <> "" Then
                strSQL = strSQL & " And instr([6],','||a.采集项目ID||',') > 0 "
            End If
        End If
        
        If Me.rptPlist.Tag <> "" Then
            strDateBegin = varFilter(mFilter.开始时间)
            strDateEnd = varFilter(mFilter.结束时间)
        Else
            strDateBegin = NowDate - Val(varFilter(mFilter.间隔时间))
            strDateEnd = NowDate
        End If
    Else
        strDateBegin = NowDate - 3
        strDateEnd = NowDate
    End If
    
    If intPatientPage <> 0 Then
        strSQL = strSQL & " and a.主页id = [9] "
    End If
    
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
                 
    If blOnlyWhere = True Then
        strSQL = " select /*+ rule */ distinct a.类别,a.医嘱id,a.相关id,a.试管颜色,a.采集方式,a.医嘱内容,a.样本条码,a.采样时间,a.执行科室,a.开嘱医生,a.开嘱时间,a.发送人,a.发送时间,a.试管编码," & vbCrLf & _
             "        a.标本,a.病人姓名,a.性别,a.年龄,a.床号,a.标识号,Decode(A.病人来源, 2, nvl(A.病人所在科室,A.申请科室), A.申请科室) As 病人所在科室,a.拒收,a.病人ID,a.采样人,a.采血量,a.试管名称,a.紧急,a.病人来源,a.婴儿,a.别名,a.执行状态,a.NO," & vbCrLf & _
             "        a.审核时间,a.申请科室,a.接收时间,a.组合项目,b.记帐费用,b.记录状态,a.接收人,b.记录性质,a.采集项目ID,a.就诊卡号,a.挂号单,a.条码打印,a.执行说明,a.重采, " & vbCrLf & _
             "        a.标本送出时间,a.主页ID,a.执行科室ID,a.采集执行科室,a.诊疗项目ID,a.检验执行科室ID,b.门诊标志,a.计费状态 , 结算模式" & vbCrLf & _
             " from "
        strSQL = strSQL & " ( Select decode(d.类别,'K','输血','检验') 类别, B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色,decode(d.类别,'K',b.医嘱内容 ,d.名称) As 采集方式, decode(d.类别,'K',d.名称,b.医嘱内容) as 医嘱内容, C.样本条码,C.采样时间, " & vbCrLf & _
             "   H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
             "   I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号,4,i.门诊号) as 标识号, " & vbCrLf & _
             "   L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人ID,c.采样人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
             "   DECODE(B.紧急标志,1,'紧急','') as 紧急,b.病人来源,nvl(b.婴儿,0) as 婴儿,N.名称 as 别名,C.执行状态 执行状态,C.NO,j.审核时间,o.名称 as 申请科室,m.接收时间, " & vbCrLf & _
             "   E.组合项目,C.接收人,c.记录性质,decode(d.类别,'K',e.id ,d.id) as 采集项目ID,i.就诊卡号,a.挂号单,c.条码打印,C.执行说明,nvl(c.重采标本,0) as 重采,c.标本送出时间, " & vbCrLf & _
             "   A.主页ID,Decode(d.类别, 'K', b.执行科室ID, a.执行科室ID) 执行科室ID,P.名称 as 采集执行科室,b.诊疗项目ID, b.执行科室ID as 检验执行科室ID,c.计费状态 ,i.结算模式 " & vbCrLf & _
             "   From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
             "   采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,检验标本记录 J ,部门表 O ,部门表 P, " & vbCrLf & _
             "   (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N " & vbCrLf & _
             "  Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
             "    And (e.类别 = 'E' Or e.类别 = 'C') And E.试管编码 = G.编码 And B.执行科室id = H.ID(+) and a.执行科室ID = P.id(+) " & vbCrLf & _
             "    And  d.类别 = 'E'  And d.操作类型 = '6' And A.病人id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.发送时间+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
             "    and  m.执行部门id + 0 = [2] And B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & vbCrLf & _
             "    and a.id = m.医嘱id And E.id = N.诊疗项目ID(+) and a.id = j.医嘱id(+) and b.开嘱科室id = o.id  " & vbCrLf & _
             "    ) a , (Select 医嘱序号,记录性质,记录状态,记帐费用,门诊标志 From 住院费用记录 Where    病人ID=[1]) b " & vbCrLf & _
             "where a.医嘱id = b.医嘱序号(+) and a.记录性质 = b.记录性质(+) "
            
            '处理三种不同的状态
            If intState = 0 Then
                strSQL = strSQL & " And a.样本条码 is null And a.执行状态 in (0,2) " & vbCrLf
            ElseIf intState = 1 Then
                strSQL = strSQL & " And a.样本条码 is not null And a.采样时间 is  null And a.执行状态 in (0,2) " & vbCrLf
            ElseIf intState = 2 Then
                strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0,2) and 标本送出时间 is null " & vbCrLf
            ElseIf intState = 3 Then
                strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0,2) and 标本送出时间 is not null  " & vbCrLf
            ElseIf intState = 4 Then
                strSQL = strSQL & " And a.执行状态 in (1,3) " & vbCrLf
            ElseIf intState = 5 Then
                strSQL = strSQL & " And a.执行状态 in (2) " & vbCrLf
            End If
        
        '单据号
'        If Mid(Me.txtGoto.Text, 1, 1) = "/" Then
'            strNO = Mid(Me.txtGoto, 2)
'            If IsNumeric(strNO) = True Then
'                strsql = strsql & " And a.NO = [7] "
'            End If
'        End If
        '根据参数来判断是否按采集科室来显示
        If chkDeptShow.Value <> 1 Then
            strSQL = Replace(strSQL, " and  m.执行部门id + 0 = [2] ", "")
        End If
        
        If Mid(Me.txtGoto.Text, 1, 1) = "*" Or Mid(Me.txtGoto.Text, 1, 1) = "." Then
            strSQL = strSQL & " And a.病人来源 in (1,3,4) "
        End If
        
        If Mid(Me.txtGoto.Text, 1, 1) = "+" Then
            strSQL = strSQL & " And a.病人来源 in ( 2,4) "
        End If
        
        
        '条码
        If BlnIsNumber(txtGoto) Then
            strSQL = strSQL & " And (a.样本条码 = [8] or a.就诊卡号 = [8]) "
        End If
        
    End If
    
    If intPatientType <> 2 Then
        strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    End If
    strBooldSql = GetBloodAdviceSql(intState, intPatientType, blOnlyWhere, intPatientPage)
    strSQL = strSQL & " Union ALL " & strBooldSql & " order by 类别,试管编码,相关ID,执行科室,标本,样本条码,医嘱内容,开嘱时间,组合项目 desc "
             
    If strTmp <> "" Or rptPlist.Tag <> "" Then
        If UBound(varFilter) >= mFilter.标本 Then
            strSample = varFilter(mFilter.标本) & ","
        End If
        If UBound(varFilter) >= mFilter.采集方式 Then
            strCapture = varFilter(mFilter.采集方式) & ","
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, mlngDeptID, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), strSample, strCapture, zlCommFun.GetFullNO(strNO), txtGoto, _
                intPatientPage)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, mlngDeptID, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), "", 0, zlCommFun.GetFullNO(strNO), txtGoto, intPatientPage)
    End If
    
    Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
    Me.rptCuvette.Records.DeleteAll
    
    Do Until rsTmp.EOF
        blnShowExec = True
        '跟据权限来判断是否显示未收费的门诊记录
        If InStr(mstrPrivs, "显示划价记录") <= 0 And Nvl(rsTmp("记录状态"), "NULL") <> "NULL" Then
            If Nvl(rsTmp("门诊标志"), 1) = 1 Then
                '只处理门诊病人
                If Nvl(rsTmp("记录状态"), 0) <> 1 Then blnShowExec = False
            End If
        Else
            '退费的项目也不显示
            If Nvl(rsTmp("记录状态"), 0) > 1 Then blnShowExec = False
        End If
        
        '没有对应颜色编码的采集不写入
        If IsNull(rsTmp("试管颜色")) = False And blnShowExec = True Then
            
            If strOldAdvice <> rsTmp("相关ID") Then
                
                Set Record = Me.rptAlist(Me.TabCtr.Selected.Index).Records.Add
                For intLoop = 0 To Me.rptAlist(Me.TabCtr.Selected.Index).Columns.Count + 1
                    Record.AddItem ""
                Next
                
                Record(mAcol.ID).Value = Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                Record(mAcol.选择).HasCheckbox = True
                
                '按输入单据内容来选择要绑定的内容
                If InStr(strAdvice, ";" & Nvl(rsTmp("医嘱内容")) & Nvl(rsTmp("标本")) & Nvl(rsTmp("婴儿"))) <= 0 Then
                    If blOnlyWhere = True Then
                        Select Case Mid(txtGoto, 1, 1)
                            Case "+", "*"                           '住院号,门诊号
    '                            If Nvl(rsTmp("标识号")) = Mid(txtGoto, 2) Then
                                    Record(mAcol.选择).Checked = True
    '                            End If
                            Case "."                                '挂号单号
    '                            If Nvl(rsTmp("挂号单")) = Mid(txtGoto, 2) Then
                                    Record(mAcol.选择).Checked = True
    '                            End If
                            Case "/"                                '收费单据号
                                If Nvl(rsTmp("NO")) = zlCommFun.GetFullNO(Mid(txtGoto, 2)) Then
                                    Record(mAcol.选择).Checked = True
                                End If
                            Case Else
                                Record(mAcol.选择).Checked = True
                        End Select
                        strAdvice = strAdvice & ";" & Nvl(rsTmp("医嘱内容")) & Nvl(rsTmp("标本")) & Nvl(rsTmp("婴儿"))
                    Else
                        Record(mAcol.选择).Checked = True
                        strAdvice = strAdvice & ";" & Nvl(rsTmp("医嘱内容")) & Nvl(rsTmp("标本")) & Nvl(rsTmp("婴儿"))
                    End If
                End If
                If Nvl(rsTmp("婴儿"), 0) > 0 Then
                
                    If rsTmp("病人来源") = 2 Then
                        gstrSql = "select 婴儿姓名,婴儿性别 from 病人新生儿记录 where 病人ID = [1] and 主页ID = [2] and 序号 = [3] "
                        Set rsBaby = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Nvl(rsTmp("病人ID"), 0)), CLng(Nvl(rsTmp("主页ID"), 0)), CInt(rsTmp("婴儿")))
                        If rsBaby.EOF = False Then
                            Record(mAcol.婴儿姓名).Value = "婴儿(" & rsBaby("婴儿姓名") & ")"
                            Record(mAcol.婴儿性别).Value = "婴儿(" & rsBaby("婴儿性别") & ")"
                        Else
                            Record(mAcol.婴儿姓名).Value = "婴儿(" & rsTmp("婴儿") & ")"
                            Record(mAcol.婴儿性别).Value = "未知"
                        End If
                    Else
                        Record(mAcol.婴儿姓名).Value = "婴儿(" & rsTmp("婴儿") & ")"
                        Record(mAcol.婴儿性别).Value = "未知"
                    End If
                End If
                
'                Record(mAcol.选择).Checked = True
                '没产生费用时也显是为“√”
                If IsNull(rsTmp("记录状态")) = True Then
                    Record(mAcol.记录状态).Value = "√"
                Else
                    Record(mAcol.记录状态).Value = IIf(rsTmp("记录状态") = 1, "√", "×")
                End If
                Record(mAcol.重采).Value = IIf(rsTmp("重采") = 1, "√", "")
                Record(mAcol.急诊).Icon = IIf(rsTmp("紧急") = "紧急", 2, 10)
                Record(mAcol.图标).BackColor = Val(Nvl(rsTmp("试管颜色")))
                Record(mAcol.采集方式).Value = Nvl(rsTmp("采集方式"))
                Record(mAcol.医嘱内容).Value = Nvl(rsTmp("医嘱内容"))
                Record(mAcol.条码).Value = Nvl(rsTmp("样本条码"))
                Record(mAcol.执行科室).Value = Nvl(rsTmp("执行科室"))
                Record(mAcol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
                Record(mAcol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
                Record(mAcol.发送人).Value = Nvl(rsTmp("发送人"))
                Record(mAcol.发送时间).Value = Nvl(rsTmp("发送时间"))
                Record(mAcol.试管颜色).Value = Val(Nvl(rsTmp("试管颜色")))
                Record(mAcol.试管编码).Value = Nvl(rsTmp("试管编码"))
                Record(mAcol.标本).Value = Nvl(rsTmp("标本")) & IIf(Nvl(rsTmp("婴儿")) = 0, "", "(婴儿" & rsTmp("婴儿") & " )")
                Record(mAcol.采样时间).Value = Nvl(rsTmp("采样时间"))
                Record(mAcol.采样人).Value = Nvl(rsTmp("采样人"))
                Record(mAcol.采血量).Value = Nvl(rsTmp("采血量"))
                Record(mAcol.试管名称).Value = Nvl(rsTmp("试管名称"))
                Record(mAcol.紧急).Value = Nvl(rsTmp("紧急"))
                Record(mAcol.病人来源).Value = Nvl(rsTmp("病人来源"))
                Record(mAcol.婴儿).Value = Nvl(rsTmp("婴儿"))
                Record(mAcol.别名).Value = Nvl(rsTmp("别名"))
                Record(mAcol.相关ID).Value = Nvl(rsTmp("相关ID"))
                Record(mAcol.执行状态).Value = Nvl(rsTmp("执行状态"))
                Record(mAcol.NO).Value = Nvl(rsTmp("NO"))
                Record(mAcol.审核时间).Value = Nvl(rsTmp("审核时间"))
                Record(mAcol.申请科室).Value = Nvl(rsTmp("申请科室"))
                Record(mAcol.接收时间).Value = Nvl(rsTmp("接收时间"))
                Record(mAcol.接收人).Value = Nvl(rsTmp("接收人"))
                Record(mAcol.送出时间).Value = Nvl(rsTmp("标本送出时间"))
                Record(mAcol.条码打印).Value = IIf(Val(Nvl(rsTmp("条码打印"))) = 0, "未打印", "已打印")
                Record(mAcol.拒收理由).Value = Nvl(rsTmp("执行说明"))
                Record(mAcol.采集科室ID).Value = Nvl(rsTmp("执行科室ID"))
                Record(mAcol.采集执行科室).Value = Nvl(rsTmp("采集执行科室"))
                Record(mAcol.诊疗项目ID).Value = Val(Nvl(rsTmp("诊疗项目ID")))
                Record(mAcol.检验执行科室ID).Value = Val(Nvl(rsTmp("检验执行科室ID")))
                Record(mAcol.计费状态).Value = Val(Nvl(rsTmp("计费状态")))
                Record(mAcol.记录性质).Value = Val(Nvl(rsTmp("记录性质")))
                Record(mAcol.类别).Value = Nvl(rsTmp("类别"))
                Record(mAcol.病人所在科室).Value = Nvl(rsTmp("病人所在科室"))
                If Nvl(rsTmp("类别")) = "输血" Then blnFL = True    '当存在输血医嘱时,需要分类显示,用于提示技师这是输血医嘱,不要当做检验医嘱处理
                For intLoop = 0 To Me.rptAlist(Me.TabCtr.Selected.Index).Columns.Count + 1
                    Record(intLoop).ForeColor = Val(Nvl(rsTmp("试管颜色")))
                Next
                Record(mAcol.别名).Value = IIf(Trim(Nvl(rsTmp("别名"))) = "", Nvl(rsTmp("医嘱内容")), Nvl(rsTmp("别名")))
                If blOnlyWhere = True Then
                    If Record(mAcol.条码).Value <> "" And intShowButtom <> 2 Then
                        intShowButtom = 1
                    End If
                    If Record(mAcol.采样人).Value <> "" Then
                        If Nvl(rsTmp("标本送出时间")) = "" Then
                            intShowButtom = 2
                        Else
                            intShowButtom = 3
                        End If
                    End If
                    
                    If Record(mAcol.执行状态).Value = 1 Then
                        intShowButtom = 4
                    End If
                End If
                If strReceivePeople = "" Then
                    If Record(mAcol.接收人).Value <> "" Then
                        strReceivePeople = Record(mAcol.接收人).Value
                    End If
                End If
                If Record(mAcol.重采).Value = 1 Then
                    For intLoop = 0 To Me.rptAlist(Me.TabCtr.Selected.Index).Columns.Count + 1
                        Record(intLoop).Bold = True
                    Next
                End If
            Else
                str医嘱内容 = Nvl(rsTmp("医嘱内容"))
                If InStr(";" & Record(mAcol.医嘱内容).Value & ";", ";" & str医嘱内容 & ";") <= 0 Then
                    Record(mAcol.医嘱内容).Value = Record(mAcol.医嘱内容).Value & ";" & Nvl(rsTmp("医嘱内容"))
                End If
                
                Record(mAcol.合并医嘱).Value = Record(mAcol.合并医嘱).Value & "," & _
                                               Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                
                str医嘱内容 = IIf(Trim(Nvl(rsTmp("别名"))) = "", Nvl(rsTmp("医嘱内容")), Nvl(rsTmp("别名")))
                If InStr(";" & Record(mAcol.别名).Value & ";", ";" & str医嘱内容 & ";") <= 0 Then
                    Record(mAcol.别名).Value = Record(mAcol.别名).Value & ";" & str医嘱内容
                End If
                Record(mAcol.诊疗项目组合).Value = Record(mAcol.诊疗项目组合).Value & ";" & Val(Nvl(rsTmp("诊疗项目ID")))
                
            End If
            strOldAdvice = rsTmp("相关ID")
            If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("试管编码")) & ",") <= 0 Then
                strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("试管编码"))
            End If
            
            If InStr(1, mstrBarCodes & ",", "," & Nvl(rsTmp("样本条码")) & ",") <= 0 Then
                mstrBarCodes = mstrBarCodes & "," & Nvl(rsTmp("样本条码"))
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If mstrBarCodes <> "" Then
        mstrBarCodes = Mid$(mstrBarCodes, 2)
    End If
    
    If Me.Visible = True Then
        Me.rptAlist(Me.TabCtr.Selected.Index).Populate
    End If
    
    '有记录时表示写入成功
    If rptAlist(Me.TabCtr.Selected.Index).Records.Count > 0 Then
        RefreshAdviceData = True
    End If
    
    '当使用病人查找时填写病人信息
    Me.txt年龄.Text = ""
'    Me.cboAge.Text = ""
    Me.txt年龄1.Text = ""
    cbo开单科室.Text = ""
    cbo医生.Text = ""
    txt医嘱内容.Text = ""
    txt医嘱内容.Tag = ""
    
    If blOnlyWhere = True Then
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            txt姓名 = Nvl(rsTmp("病人姓名"))
            txt姓名.Tag = Nvl(rsTmp("病人姓名"))
            On Error Resume Next
            cbo性别 = Nvl(rsTmp("性别"))
            cbo性别.Tag = ""
            strAge = Nvl(rsTmp("年龄"))
            
            strAge = Replace(strAge, "小时", "时")
            strAge = Replace(strAge, "分钟", "分")
        
            If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
                If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                    Me.txt年龄.Text = ""
                    Me.cboAge.Text = Trim(strAge)
                Else
                    strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                    aAge = Split(strAge, ";")
                    If UBound(aAge) = 1 Then
                        Me.txt年龄.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Else
                        Me.txt年龄.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                        Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                    End If
                End If
            Else
                Me.txt年龄.Text = ""
                Me.cboAge.ListIndex = 0
            End If
'            txt年龄 = Val(Nvl(rsTmp("年龄")))
'            If IsNumeric(Nvl(rsTmp("年龄"))) = False And Len(Nvl(rsTmp("年龄"))) > 0 Then
'                Me.cboAge = Mid(Nvl(rsTmp("年龄")), Len(rsTmp("年龄")))
'            End If

            On Error GoTo 0
            txtBed = Nvl(rsTmp("床号"))
            txtID = Nvl(rsTmp("标识号"))
            If Nvl(rsTmp("病人来源")) = 2 Then
                lblCap(0).Caption = "住  院 号"
            Else
                If Nvl(rsTmp("标识号")) = "" Then
                    lblCap(0).Caption = "单  据 号"
                    txtID = Nvl(rsTmp("NO"))
                Else
                    lblCap(0).Caption = "门  诊 号"
                End If
            End If
            txtPatientDept = Nvl(rsTmp("病人所在科室"))
'            cbo开单科室.ListIndex = -1
'            cbo医生.ListIndex = -1
'            txt医嘱内容.Text = ""
'            txt医嘱内容.Tag = ""
            
            cbo开单科室.Text = Nvl(rsTmp("申请科室"))
            cbo医生.Text = Nvl(rsTmp("开嘱医生"))
            txt医嘱内容.Text = Nvl(rsTmp("医嘱内容"))
            
            If Nvl(rsTmp("拒收")) <> "" Then
                lblCap(6).Visible = True
            Else
                lblCap(6).Visible = False
            End If
            If Val(Nvl(rsTmp("结算模式"))) = 1 Then
                lblCap(10).Visible = True
            Else
                lblCap(10).Visible = False
            End If
            mlngKey = Nvl(rsTmp("病人ID"))
            
            Select Case intState
                Case 1
                    Me.cmdBindBarCode.Enabled = True
                    Me.cmdNewBarcode.Enabled = True
                    Me.cmdComplete.Enabled = True
                    Me.cmdBarcodePrint.Enabled = True
                    Me.cmdBakBillPrint.Enabled = True
                    Me.cmdBindBarCode.Caption = "清除绑定(&B)"
                    Me.cmdNewBarcode.Caption = "取消条码(&N)"
                    Me.cmdComplete.Caption = "完成采集(&P)"
                Case 2
                    Me.cmdBindBarCode.Enabled = False
                    Me.cmdNewBarcode.Enabled = True
                    If strReceivePeople = "" Then
                        Me.cmdComplete.Enabled = True
                    Else
                        Me.cmdComplete.Enabled = False
                    End If
                    Me.cmdBarcodePrint.Enabled = True
                    Me.cmdBakBillPrint.Enabled = True
                    Me.cmdBindBarCode.Caption = "清除绑定(&B)"
                    Me.cmdNewBarcode.Caption = "送检标本(&C)"
                    Me.cmdComplete.Caption = "取消完成(&P)"
                Case 3
                    Me.cmdBindBarCode.Enabled = False
                    Me.cmdNewBarcode.Enabled = True
                    Me.cmdComplete.Enabled = False
                    Me.cmdBarcodePrint.Enabled = True
                    Me.cmdBakBillPrint.Enabled = True
                    Me.cmdBindBarCode.Caption = "清除绑定(&B)"
                    Me.cmdNewBarcode.Caption = "取消送检(&C)"
                    Me.cmdComplete.Caption = "取消完成(&P)"
                Case 4
                    Me.cmdBindBarCode.Enabled = False
                    Me.cmdNewBarcode.Enabled = False
                    Me.cmdComplete.Enabled = False
                    Me.cmdBarcodePrint.Enabled = False
                    Me.cmdBakBillPrint.Enabled = False
                    Me.cmdBindBarCode.Caption = "清除绑定(&B)"
                    Me.cmdNewBarcode.Caption = "取消条码(&N)"
                    Me.cmdComplete.Caption = "取消完成(&P)"
                Case Else
                    Me.cmdBindBarCode.Enabled = True
                    Me.cmdNewBarcode.Enabled = True
                    Me.cmdComplete.Enabled = False
                    Me.cmdBarcodePrint.Enabled = False
                    Me.cmdBakBillPrint.Enabled = False
                    Me.cmdBindBarCode.Caption = "绑定条码(&B)"
                    Me.cmdNewBarcode.Caption = "生成条码(&N)"
                    Me.cmdComplete.Caption = "完成采集(&P)"
                            
            End Select
        End If
    Else
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            txt姓名 = Nvl(rsTmp("病人姓名"))
            txt姓名.Tag = Nvl(rsTmp("病人姓名"))
            On Error Resume Next
            cbo性别 = Nvl(rsTmp("性别"))
            cbo性别.Tag = ""
            
            strAge = Nvl(rsTmp("年龄"))
            
            
            strAge = Replace(strAge, "小时", "时")
            strAge = Replace(strAge, "分钟", "分")
            
            
            If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
                If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                    Me.txt年龄.Text = ""
                    Me.cboAge.Text = Trim(strAge)
                Else
                    strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                    aAge = Split(strAge, ";")
                    If UBound(aAge) = 1 Then
                        Me.txt年龄.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Else
                        Me.txt年龄.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                        Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                    End If
                End If
            Else
                Me.txt年龄.Text = ""
                Me.cboAge.ListIndex = 0
            End If
'            txt年龄 = Val(Nvl(rsTmp("年龄")))
'            If IsNumeric(Nvl(rsTmp("年龄"))) = False And Len(Nvl(rsTmp("年龄"))) > 0 Then
'                Me.cboAge = Mid(Nvl(rsTmp("年龄")), Len(rsTmp("年龄")))
'            End If
            On Error GoTo 0
            txtBed = Nvl(rsTmp("床号"))
            txtID = Nvl(rsTmp("标识号"))
            If Nvl(rsTmp("病人来源")) = 2 Then
                lblCap(0).Caption = "住  院 号"
            Else
                If Nvl(rsTmp("标识号")) = "" Then
                    lblCap(0).Caption = "单  据 号"
                    txtID = Nvl(rsTmp("NO"))
                Else
                    lblCap(0).Caption = "门  诊 号"
                End If
            End If
            
            
            txtPatientDept = Nvl(rsTmp("病人所在科室"))
'            cbo开单科室.ListIndex = -1
'            cbo医生.ListIndex = -1
'            txt医嘱内容.Text = ""
'            txt医嘱内容.Tag = ""
            
            cbo开单科室.Text = Nvl(rsTmp("申请科室"))
            cbo医生.Text = Nvl(rsTmp("开嘱医生"))
            txt医嘱内容.Text = Nvl(rsTmp("医嘱内容"))
            
            If Nvl(rsTmp("拒收")) <> "" Then
                lblCap(6).Visible = True
            Else
                lblCap(6).Visible = False
            End If
            If Val(Nvl(rsTmp("结算模式"))) = 1 Then
                lblCap(10).Visible = True
            Else
                lblCap(10).Visible = False
            End If
            mlngKey = Nvl(rsTmp("病人ID"))
        End If
        '设置定义
        Select Case Me.TabCtr.Selected.Index
            Case 0
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "绑定条码(&B)"
            Me.cmdNewBarcode.Caption = "生成条码(&N)"
            Me.cmdComplete.Caption = "完成采集(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 1
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "取消条码(&N)"
            Me.cmdComplete.Caption = "完成采集(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 2
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            If strReceivePeople = "" Then
                Me.cmdComplete.Enabled = True
            Else
                Me.cmdComplete.Enabled = False
            End If
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "送检标本(&C)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 3
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "取消送检(&C)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 4
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = False
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "取消条码(&N)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 5
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "让步检验(&N)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "重采样本(&R)"
        End Select
    End If
    
    
    If strCuvetteNumber <> "" Then
        With Me.rptCuvette
            strSQL = "select 编码,名称,添加剂,采血量,规格,颜色 from 采血管类型 where 编码 in " & _
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
            
            If InStr(1, Mid(strCuvetteNumber, 2), ",") <= 0 And .Records.Count > 0 Then
                .Records(0).Item(mCuvette.选择).Checked = True
            End If
        End With
    End If
    If strTmp <> "" Then
    
        Me.stbThis.Panels(2).Text = "当前范围<" & IIf(varFilter(mFilter.发送或审核时间) = 0, "发送时间 ", "发送时间 ") & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> 下共有:" & _
                                Me.rptPlist.Rows.Count & "个病人."
    Else
        Me.stbThis.Panels(2).Text = "当前范围<" & "发送时间 " & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> 下共有:" & _
                                Me.rptPlist.Rows.Count & "个病人."
    End If
    Me.rptCuvette.Populate
    
    Me.rptAlist(Me.TabCtr.Selected.Index).GroupsOrder.DeleteAll
    If blnFL = True Then
        With Me.rptAlist(Me.TabCtr.Selected.Index)
            Call .GroupsOrder.Add(.Columns.Column(mAcol.类别))
            .Populate
        End With
    End If
    If mlngKey <> 0 Then
        Call ReadPatPricture(mlngKey, imgLoad)
        If imgLoad.Picture = 0 Then
            imgPatient.Picture = imgDefual.Picture
        Else
            imgPatient.Picture = imgLoad.Picture
        End If
    End If
'    SelectCuvette
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetBooldPatientDataSql() As String
    Dim strBloodSQL
    Dim strTmp As String
    Dim strBloodType As String
    Dim varFilter As Variant
    Dim NowDate As Date

    On Error GoTo errH

    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")
    strBloodType = zlDatabase.GetPara(273, 100)
    NowDate = zlDatabase.Currentdate
    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If

    strBloodSQL = strBloodSQL & "Select    distinct a.病人id,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') as 病人来源, " & vbCrLf & _
                " c.姓名 as 病人姓名,e.名称 as 病人科室,c.性别,c.年龄,c.就诊卡号,b.样本条码, " & vbCrLf & _
                " decode(b.执行状态,1,'已完成',2,'拒收',3,'执行中', " & vbCrLf & _
                "         Decode(b.样本条码, Null, '未绑定', Decode(b.采样人, Null, '已绑定', decode(b.标本送出时间,null,'已采样','已送检'))) )as 状态, " & vbCrLf & _
                " decode(A.病人来源, 1, C.门诊号, 2, C.住院号,4,c.门诊号) As 标识号, " & vbCrLf & _
                " decode(c.当前床号,null,decode(l.出院病床,null,l.入院病床,l.出院病床),c.当前床号) as 当前床号 , " & vbCrLf & _
                " decode(a.紧急标志,1,'紧急',decode(g.急诊,1,'紧急')) as 紧急 ,decode(a.病人来源 , 2,a.主页ID,0) 主页ID, " & vbCrLf & _
                " decode(b.执行状态,0,'',2,'拒收') as 拒收,b.执行状态,nvl(b.重采标本,0) as 重采标本,b.发送时间,nvl(s.路径状态,0) as 临床路径病人,a.医嘱内容 " & vbCrLf & _
                " From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 部门表 E, 诊疗项目目录 F,病人挂号记录 G,病人医嘱记录 H, " & vbCrLf & _
                "      诊疗项目目录 K ,病案主页 L,病人医嘱发送 M,检验标本记录 J,病案主页 S " & vbCrLf & _
                " Where A.ID = H.相关ID And H.id = B.医嘱id And A.病人id = C.病人id And A.病人科室id = E.ID And A.诊疗项目id+0 = f.ID  " & vbCrLf & _
                "      And h.诊疗项目ID = k.id and a.id = j.医嘱id(+)  " & vbCrLf & _
                " And A.挂号单 = G.No(+) and a.病人id = g.病人id(+)  and a.病人id = g.病人id(+)  and (g.病人ID is null or (g.记录状态 =1 and g.记录性质 =1) ) And f.类别 = 'K'  and Decode(f.类别, 'K', '9') = k.操作类型 and a.病人id = l.病人ID(+) and a.主页ID = l.主页ID(+) and b.执行部门id + 0 = [1] " & vbCrLf & _
                " And A.ID = M.医嘱ID And k.试管编码 is not null and a.病人ID = S.病人ID(+) and a.主页ID = s.主页ID(+) " & IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "")


    If Me.rptPlist.Tag <> "" Then
        strBloodSQL = strBloodSQL & " And A.病人来源 in (" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3", "0") & "," & _
                      Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & ") "
        If varFilter(mFilter.标识号) <> "" Then
            strBloodSQL = strBloodSQL & " And decode(a.病人来源,2,c.住院号,c.门诊号) = [2] "
        End If

        If varFilter(mFilter.就诊卡) <> "" Then
            strBloodSQL = strBloodSQL & " And c.就诊卡号 = [3] "

        End If

        If varFilter(mFilter.姓名) <> "" Then
            strBloodSQL = strBloodSQL & " And C.姓名 like [4] "

        End If

        If varFilter(mFilter.单据号) <> "" Then
            strBloodSQL = strBloodSQL & " and B.NO = [5]"

        End If

        If UBound(varFilter) >= mFilter.标本 Then
            If Trim(varFilter(mFilter.标本)) <> "" Then
                strBloodSQL = strBloodSQL & " And decode(f.类别,'K',1,instr([6],','||H.标本部位||',')) > 0 "

            End If
        End If

        If UBound(varFilter) >= mFilter.采集方式 Then
            If Trim(varFilter(mFilter.采集方式)) <> "" Then
                strBloodSQL = strBloodSQL & " And instr([7],','|| decode(f.类别,'K',K.id,f.ID) ||',') > 0 "

            End If
        End If

        If varFilter(mFilter.发送或审核时间) = 0 Then
            strBloodSQL = strBloodSQL & " and b.发送时间 Between [8] and [9]"

        Else
            strBloodSQL = strBloodSQL & " and b.发送时间 Between [8] and [9]"
        End If

        If UBound(varFilter) >= mFilter.检验类型 Then
            If Trim(varFilter(mFilter.检验类型)) <> "" Then
                strBloodSQL = strBloodSQL & " And instr([10],','|| decode(f.类别,'K','" & strBloodType & "',k.操作类型) ||',') > 0 "

            End If
        End If
    Else
        If strTmp <> "" Then
            strBloodSQL = strBloodSQL & " And instr('" & IIf(Val(varFilter(mFilter.门诊)) = 1, "1,3", "0") & "," & _
                          Val(varFilter(mFilter.住院)) & "," & Val(varFilter(mFilter.体检)) & "',A.病人来源)>0 "

            If UBound(varFilter) >= mFilter.标本 Then
                If Trim(varFilter(mFilter.标本)) <> "" Then
                    strBloodSQL = strBloodSQL & " And decode(f.类别,'K',1,instr([6],','||H.标本部位||',')) > 0 "

                End If
            End If

            If UBound(varFilter) >= mFilter.采集方式 Then
                If Trim(varFilter(mFilter.采集方式)) <> "" Then
                    strBloodSQL = strBloodSQL & " And instr([7],','|| decode(f.类别,'K',K.id,f.ID) ||',') > 0 "

                End If
            End If

            If varFilter(mFilter.发送或审核时间) = 0 Then
                strBloodSQL = strBloodSQL & " and b.发送时间 Between [8] and [9]"
            Else
                strBloodSQL = strBloodSQL & " and b.发送时间 Between [8] and [9]"

            End If

            If UBound(varFilter) >= mFilter.检验类型 Then
                If Trim(varFilter(mFilter.检验类型)) <> "" Then
                    strBloodSQL = strBloodSQL & " And instr([10],','|| decode(f.类别,'K','" & strBloodType & "',k.操作类型) ||',') > 0 "

                End If
            End If
        Else
            strBloodSQL = strBloodSQL & " and m.发送时间 Between [8] and [9]"
        End If
    End If
    GetBooldPatientDataSql = strBloodSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function GetBloodAdviceSql(intState As Integer, intPatientType As Integer, Optional blOnlyWhere As Boolean = False, Optional intPatientPage As Integer) As String
    Dim blnDateMoved As Boolean                     '是否转出
    Dim strSQL As String                            'SQL语句
    Dim strTmp As String                            '临时字串变量
    Dim varFilter As Variant                        '过滤字串
    Dim strDateBegin As String                      '开始发送时间
    Dim strDateEnd As String                        '结束发送时间
    Dim rsTmp As New ADODB.Recordset                '数据集
    Dim intLoop As Integer                          '循环变量
    Dim Record As ReportRecord                      '列表数据集
    Dim strOldAdvice As String                      '记录上次医嘱
    Dim strCuvetteNumber As String                  '用于记录试管编码
    Dim NowDate As Date                             '当前时间
    Dim blnFL As Boolean                            '是否分类显示,存在输血医嘱时,分类显示
    Dim strSQL1 As String


    blnDateMoved = MovedByDate(Date)    '按时间看是否可能已转出

    '从注册表中读取过滤条件
    strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")

    NowDate = zlDatabase.Currentdate

    '从过滤窗体过来有条件时优先
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If

    On Error GoTo errH
    strSQL = " select /*+ rule */ distinct a.类别,a.医嘱id,a.相关id,a.试管颜色,a.采集方式,a.医嘱内容,a.样本条码,a.采样时间,a.执行科室,a.开嘱医生,a.开嘱时间,a.发送人,a.发送时间,a.试管编码," & vbCrLf & _
           "        a.标本,a.病人姓名,a.性别,a.年龄,a.床号,a.标识号,Decode(A.病人来源, 2, nvl(A.病人所在科室,A.申请科室), A.申请科室) As 病人所在科室,a.拒收,a.病人ID,a.采样人,a.采血量,a.试管名称,a.紧急,a.病人来源,a.婴儿,a.别名,a.执行状态,a.NO," & vbCrLf & _
           "        a.审核时间,a.申请科室,a.接收时间,a.组合项目,b.记帐费用,b.记录状态,a.接收人,b.记录性质,a.采集项目ID,a.就诊卡号,a.挂号单,a.条码打印,a.执行说明,a.重采, " & vbCrLf & _
           "        a.标本送出时间,a.主页ID,a.执行科室ID,a.采集执行科室,a.诊疗项目ID,a.检验执行科室ID,b.门诊标志,a.计费状态 , 结算模式" & vbCrLf & _
           " from "

    strSQL = strSQL & "(  Select decode(d.类别,'K','输血','检验') 类别, B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色, decode(d.类别,'K',b.医嘱内容 ,d.名称) As 采集方式, decode(d.类别,'K',d.名称,b.医嘱内容) as 医嘱内容, C.样本条码,C.采样时间, " & vbCrLf & _
           "   H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
           "   I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号,4,i.门诊号) as 标识号, " & vbCrLf & _
           "   L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人ID,c.采样人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
           "   DECODE(B.紧急标志,1,'紧急','') as 紧急,b.病人来源,nvl(b.婴儿,0) as 婴儿,N.名称 as 别名,decode(d.类别, 'K', M.执行状态,C.执行状态) 执行状态,C.NO,j.审核时间,o.名称 as 申请科室,m.接收时间, " & vbCrLf & _
           "   E.组合项目,C.接收人,c.记录性质,decode(d.类别,'K',e.id ,d.id)  as 采集项目ID,i.就诊卡号,b.挂号单,C.条码打印,C.执行说明,nvl(c.重采标本,0) as 重采,c.标本送出时间, " & vbCrLf & _
           "   a.主页ID,Decode(d.类别, 'K', b.执行科室ID, a.执行科室ID) 执行科室ID,P.名称 as 采集执行科室,b.诊疗项目ID,Decode(d.类别, 'K', a.执行科室ID, b.执行科室ID) as 检验执行科室ID,c.计费状态,i.结算模式 " & vbCrLf & _
           "   From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
           "   采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,检验标本记录 J ,部门表 O ,部门表 P, " & vbCrLf & _
           "   (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N " & vbCrLf & _
           "  Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID  " & vbCrLf & _
           "    And (e.类别 = 'E' Or e.类别 = 'C') And E.试管编码 = G.编码 And B.执行科室id = H.ID(+) and a.执行科室ID = P.id(+)  " & vbCrLf & _
           "    and   d.类别 = 'K' And  e.操作类型= '9' And A.病人id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.发送时间+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
           "    and c.执行部门id + 0 = [2] And B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & vbCrLf & _
           "    and a.id = m.医嘱id And E.id = N.诊疗项目ID(+) and a.id = j.医嘱id(+) and b.开嘱科室id = o.id  " & vbCrLf & _
           "    ) a , (Select 医嘱序号,记录性质,记录状态,记帐费用,门诊标志 From 住院费用记录 Where  病人ID=[1]) b " & vbCrLf & _
             "where a.医嘱id = b.医嘱序号(+) and a.记录性质 = mod(b.记录性质(+),10)  "

    '  IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 between  to_date( '" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 00:00:00' ,'yyyy-mm-dd hh24:mi:ss' ) and to_date('" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 23:59:59','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _

       '处理三种不同的状态,
    If intState = 0 Then
        strSQL = strSQL & " And a.样本条码 is null And a.执行状态 in (0) " & vbCrLf
    ElseIf intState = 1 Then
        strSQL = strSQL & " And a.样本条码 is not null And a.采样时间 is  null And a.执行状态 in (0) " & vbCrLf
    ElseIf intState = 2 Then
        strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0) and a.标本送出时间 is null " & vbCrLf
    ElseIf intState = 3 Then
        strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0) and a.标本送出时间 is not null  " & vbCrLf
    ElseIf intState = 4 Then
        strSQL = strSQL & " And a.执行状态 in (1,3) " & vbCrLf
    ElseIf intState = 5 Then
        strSQL = strSQL & " And a.执行状态 in (2) " & vbCrLf
    End If

    '过滤
    If Me.rptPlist.Tag <> "" Or strTmp <> "" Then

        If UBound(varFilter) >= mFilter.标本 Then
            If Trim(varFilter(mFilter.标本)) <> "" Then
                strSQL = strSQL & " And decode(a.类别,'输血',1,instr([5],','||a.标本||',')) > 0 "
            End If
        End If

        If UBound(varFilter) >= mFilter.采集方式 Then
            If Trim(varFilter(mFilter.采集方式)) <> "" Then
                strSQL = strSQL & " And instr([6],','||a.采集项目ID||',') > 0 "
            End If
        End If

        If Me.rptPlist.Tag <> "" Then
            strDateBegin = varFilter(mFilter.开始时间)
            strDateEnd = varFilter(mFilter.结束时间)
        Else
            strDateBegin = NowDate - Val(varFilter(mFilter.间隔时间))
            strDateEnd = NowDate
        End If
    Else
        strDateBegin = NowDate - 3
        strDateEnd = NowDate
    End If

    If intPatientPage <> 0 Then
        strSQL = strSQL & " and a.主页id = [9] "
    End If


    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If

    If blOnlyWhere = True Then
        strSQL = " select /*+ rule */ distinct a.类别,a.医嘱id,a.相关id,a.试管颜色,a.采集方式,a.医嘱内容,a.样本条码,a.采样时间,a.执行科室,a.开嘱医生,a.开嘱时间,a.发送人,a.发送时间,a.试管编码," & vbCrLf & _
               "        a.标本,a.病人姓名,a.性别,a.年龄,a.床号,a.标识号,Decode(A.病人来源, 2, nvl(A.病人所在科室,A.申请科室), A.申请科室) As 病人所在科室,a.拒收,a.病人ID,a.采样人,a.采血量,a.试管名称,a.紧急,a.病人来源,a.婴儿,a.别名,a.执行状态,a.NO," & vbCrLf & _
               "        a.审核时间,a.申请科室,a.接收时间,a.组合项目,b.记帐费用,b.记录状态,a.接收人,b.记录性质,a.采集项目ID,a.就诊卡号,a.挂号单,a.条码打印,a.执行说明,a.重采, " & vbCrLf & _
               "        a.标本送出时间,a.主页ID,a.执行科室ID,a.采集执行科室,a.诊疗项目ID,a.检验执行科室ID,b.门诊标志,a.计费状态 , 结算模式" & vbCrLf & _
               " from "

        strSQL = strSQL & "(   Select decode(d.类别,'K','输血','检验') 类别, B.ID as 医嘱ID, B.相关id, G.颜色 As 试管颜色,decode(d.类别,'K',b.医嘱内容 ,d.名称) As 采集方式, decode(d.类别,'K',d.名称,b.医嘱内容) as 医嘱内容, C.样本条码,C.采样时间, " & vbCrLf & _
               "   H.名称 As 执行科室, B.开嘱医生,B.开嘱时间, C.发送人, C.发送时间, G.编码 as 试管编码,b.标本部位 as 标本, " & vbCrLf & vbCrLf & _
               "   I.姓名 as 病人姓名,I.性别,i.年龄,i.当前床号 as 床号,decode(b.病人来源,1,I.门诊号,2,i.住院号,4,i.门诊号) as 标识号, " & vbCrLf & _
               "   L.名称 as 病人所在科室,Decode(C.执行状态,2,'拒收') as 拒收,I.病人ID,c.采样人,G.采血量,G.名称 as 试管名称, " & vbCrLf & _
               "   DECODE(B.紧急标志,1,'紧急','') as 紧急,b.病人来源,nvl(b.婴儿,0) as 婴儿,N.名称 as 别名,decode(d.类别, 'K', M.执行状态,C.执行状态) 执行状态,C.NO,j.审核时间,o.名称 as 申请科室,m.接收时间, " & vbCrLf & _
               "   E.组合项目,C.接收人,c.记录性质,decode(d.类别,'K',e.id ,d.id) as 采集项目ID,i.就诊卡号,a.挂号单,c.条码打印,C.执行说明,nvl(c.重采标本,0) as 重采,c.标本送出时间, " & vbCrLf & _
               "   A.主页ID,Decode(d.类别, 'K', b.执行科室ID, a.执行科室ID) 执行科室ID,P.名称 as 采集执行科室,b.诊疗项目ID,Decode(d.类别, 'K', a.执行科室ID, b.执行科室ID) as 检验执行科室ID,c.计费状态 ,i.结算模式 " & vbCrLf & _
               "   From 病人医嘱记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
               "   采血管类型 G,部门表 H, 病人信息 I,部门表 L,病人医嘱发送 M,检验标本记录 J ,部门表 O ,部门表 P, " & vbCrLf & _
               "   (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N " & vbCrLf & _
               "  Where A.ID = B.相关id And B.ID = C.医嘱id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID " & vbCrLf & _
               "    And (e.类别 = 'E' Or e.类别 = 'C') And E.试管编码 = G.编码 And B.执行科室id = H.ID(+) and a.执行科室ID = P.id(+) " & vbCrLf & _
               "    and d.类别 = 'K'  And  e.操作类型 = '9' And A.病人id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.发送时间+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.开始执行时间 < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
               "    and c.执行部门id + 0 = [2] And B.病人ID = I.病人ID and I.当前科室ID = L.ID(+) " & vbCrLf & _
               "    and a.id = m.医嘱id And E.id = N.诊疗项目ID(+) and a.id = j.医嘱id(+) and b.开嘱科室id = o.id  " & vbCrLf & _
               "    ) a , (Select 医嘱序号,记录性质,记录状态,记帐费用,门诊标志 From 住院费用记录 Where    病人ID=[1]) b " & vbCrLf & _
                 "where a.医嘱id = b.医嘱序号(+) and a.记录性质 = b.记录性质(+) "

        '处理三种不同的状态
        If intState = 0 Then
            strSQL = strSQL & " And a.样本条码 is null And a.执行状态 in (0,2) " & vbCrLf
        ElseIf intState = 1 Then
            strSQL = strSQL & " And a.样本条码 is not null And a.采样时间 is  null And a.执行状态 in (0,2) " & vbCrLf
        ElseIf intState = 2 Then
            strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0,2) and 标本送出时间 is null " & vbCrLf
        ElseIf intState = 3 Then
            strSQL = strSQL & " And a.样本条码 is not null and  a.采样时间 is not null And a.执行状态 in (0,2) and 标本送出时间 is not null  " & vbCrLf
        ElseIf intState = 4 Then
            strSQL = strSQL & " And a.执行状态 in (1,3) " & vbCrLf
        ElseIf intState = 5 Then
            strSQL = strSQL & " And a.执行状态 in (2) " & vbCrLf
        End If

        '单据号
        '        If Mid(Me.txtGoto.Text, 1, 1) = "/" Then
        '            strNO = Mid(Me.txtGoto, 2)
        '            If IsNumeric(strNO) = True Then
        '                strsql = strsql & " And a.NO = [7] "
        '            End If
        '        End If
        '根据参数来判断是否按采集科室来显示
        If chkDeptShow.Value <> 1 Then
            strSQL = Replace(strSQL, " and c.执行部门id + 0 = [2] ", "")
        End If

        If Mid(Me.txtGoto.Text, 1, 1) = "*" Or Mid(Me.txtGoto.Text, 1, 1) = "." Then
            strSQL = strSQL & " And a.病人来源 in (1,3,4) "
        End If

        If Mid(Me.txtGoto.Text, 1, 1) = "+" Then
            strSQL = strSQL & " And a.病人来源 in ( 2,4) "
        End If


        '条码
        If BlnIsNumber(txtGoto) Then
            strSQL = strSQL & " And (a.样本条码 = [8] or a.就诊卡号 = [8]) "
        End If

    End If

    If intPatientType <> 2 Then
        strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    End If
    GetBloodAdviceSql = strSQL

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub rptPlist_SelectionChanged()
    Dim intPatientType As Integer                   '病人来源
    Dim intPatientPage As Integer                   '主页ID
    
    If Me.rptPlist.FocusedRow Is Nothing Then Exit Sub
    
    With Me.rptPlist.FocusedRow
        mlngKey = .Record(mPcol.病人ID).Value
        intPatientType = IIf(.Record(mPcol.来源).Value = "住院", 2, 1)
        intPatientPage = .Record(mPcol.主页ID).Value
    End With
    '使用病人ID刷新医嘱
    RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index, intPatientType, False, intPatientPage
    '刷新显示信息
    ShowPatientInfo
    '定位焦点到条件输入
'    If Me.Visible = True Then Me.TxtBarCode.SetFocus
    mblFind = False
End Sub

Private Sub ShowPatientInfo()
    Dim strAge As String
    Dim aAge() As String
    
    '没有焦点行时退出
    If Me.rptPlist.FocusedRow Is Nothing Then Exit Sub
    On Error Resume Next
    With Me.rptPlist.FocusedRow
    
        Call AdjustEditState(True)
        
        txt姓名 = .Record(mPcol.病人姓名).Value
        txt姓名.Tag = .Record(mPcol.病人姓名).Value
        cbo性别 = .Record(mPcol.性别).Value
        cbo性别.Tag = ""
        strAge = .Record(mPcol.年龄).Value
        
        strAge = Replace(strAge, "小时", "时")
        strAge = Replace(strAge, "分钟", "分")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
            If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                Me.txt年龄.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                Else
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                End If
            End If
        Else
            Me.txt年龄.Text = ""
            Me.cboAge.ListIndex = 0
        End If
'        txt年龄 = Val(.Record(mPcol.年龄).Value)
'        If IsNumeric(.Record(mPcol.年龄).Value) = False And Len(.Record(mPcol.年龄).Value) > 0 Then
'                Me.cboAge = Mid(Nvl(.Record(mPcol.年龄).Value), Len(.Record(mPcol.年龄).Value))
'            End If
        txtBed = .Record(mPcol.床号).Value
        txtID = .Record(mPcol.标识号).Value
        txtPatientDept = .Record(mPcol.病人科室).Value
        cbo开单科室.ListIndex = -1
        cbo医生.ListIndex = -1
'        Me.txt医嘱内容.Text = ""
'        Me.txt医嘱内容.Tag = ""
        
        Call AdjustEditState(False)
        
        If .Record(mPcol.拒收).Value <> 0 Then
            lblCap(6).Visible = True
        Else
            lblCap(6).Visible = False
        End If
    End With
    
End Sub

Private Sub SelectCuvette()
    '功能               选择被选中的试管
    
    Dim RecordC As ReportRecord
    Dim RecordA As ReportRecord
    
    For Each RecordC In Me.rptCuvette.Records
        For Each RecordA In Me.rptAlist(Me.TabCtr.Selected.Index).Records
            If RecordA(mAcol.试管编码).Value = RecordC(mCuvette.编码).Value Then
                RecordA(mAcol.选择).Checked = RecordC(mCuvette.选择).Checked
            End If
        Next
    Next

    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
End Sub

Private Sub TabCtr_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '设置定义
    Select Case Item.Index
        Case 0
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "绑定条码(&B)"
            Me.cmdNewBarcode.Caption = "生成条码(&N)"
            Me.cmdComplete.Caption = "完成采集(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 1
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "取消条码(&N)"
            Me.cmdComplete.Caption = "完成采集(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 2
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "送检标本(&C)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 3
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "取消送检(&C)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 4
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = False
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "取消条码(&N)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "回执单打印"
        Case 5
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "清除绑定(&B)"
            Me.cmdNewBarcode.Caption = "让步检验(&N)"
            Me.cmdComplete.Caption = "取消完成(&P)"
            Me.cmdBakBillPrint.Caption = "重采样本(&R)"
    End Select
    
    If Me.Visible = True Then
        Call RefreshPatientData
    End If
    
    If InStr(1, mstrPrivs, "标本采集") <= 0 Then
        Me.cmdBindBarCode.Enabled = False
        Me.cmdNewBarcode.Enabled = False
        Me.cmdComplete.Enabled = False
        Me.cmdBarcodePrint.Enabled = False
        Me.cmdBakBillPrint.Enabled = False
    End If
    
    With Me.rptPlist
        If .Rows.Count > 0 Then
'            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index
            If .FocusedRow Is Nothing Then
                .FocusedRow = Me.rptPlist.Rows(0)
                .Populate
            Else
                RefreshAdviceData .FocusedRow.Record(mPcol.病人ID).Value, Me.TabCtr.Selected.Index, IIf(.FocusedRow.Record(mPcol.来源).Value = "住院", 2, 1), False, .FocusedRow.Record(mPcol.主页ID).Value
            End If
        Else
            Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
            Me.rptCuvette.Records.DeleteAll
            Me.rptAlist(Me.TabCtr.Selected.Index).Populate
            Me.rptCuvette.Populate
            txt姓名 = ""
            txt姓名.Tag = ""
            cbo性别.ListIndex = -1
            txt年龄 = ""
            txt年龄1 = ""
            txtBed = ""
            txtID = ""
            txtPatientDept = ""
            cbo开单科室.ListIndex = -1
            cbo医生.ListIndex = -1
            txt医嘱内容.Text = ""
            txt医嘱内容.Tag = ""
            Me.lblCap(6).Visible = False
        End If
    End With
    
    Me.stbThis.Panels(2).Text = Mid(Me.stbThis.Panels(2).Text, 1, InStr(1, Me.stbThis.Panels(2).Text, "有:") + 1) & _
                                Me.rptPlist.Rows.Count & "个病人."
End Sub

Private Sub txtBarCode_GotFocus()
    Me.TxtBarCode.SelStart = 0
    Me.TxtBarCode.SelLength = Len(Me.TxtBarCode)
End Sub

Private Sub TxtBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtBarCodeCheck.SetFocus
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub TxtBarCode_LostFocus()
    Dim Record As ReportRecord
    Dim strItems As String
    If Len(Trim(Me.TxtBarCode.Text)) < 2 Then Exit Sub
    
    If Me.TxtBarCode.Text = Me.TxtBarCode.Tag Then Exit Sub
    
    strItems = ""
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Val(Mid(Me.TxtBarCode.Text, 1, Len(Record(mAcol.试管编码).Value))) = Record(mAcol.试管编码).Value Then
            If InStr("," & strItems & ",", "," & Record(mAcol.诊疗项目ID).Value & ",") <= 0 Then
                Record(mAcol.选择).Checked = True
                Me.TxtBarCode.Tag = Me.TxtBarCode.Text
            Else
                Record(mAcol.选择).Checked = False
            End If
            strItems = strItems & "," & Record(mAcol.诊疗项目ID).Value
        Else
            Record(mAcol.选择).Checked = False
        End If
    Next
    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
End Sub

Private Sub TxtBarCodeCheck_GotFocus()
    TxtBarCodeCheck.SelStart = 0
    TxtBarCodeCheck.SelLength = Len(TxtBarCodeCheck.Text)
End Sub

Private Sub TxtBarCodeCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.TxtBarCodeCheck.Text <> "" Then
            '绑定条码
            cmdBindBarCode_Click
        Else
            If Me.TxtBarCode.Text = "" And Me.TxtBarCodeCheck.Text = "" Then
                '生成条码
                cmdNewBarcode_Click
            End If
        End If
    End If
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    
End Sub

Private Sub WriterBarCode(Mode As Integer, Optional WriterUser As Boolean = False, _
                          Optional PrintBarCode As Boolean = False, Optional PrintBackBill As Boolean = False, _
                          Optional ByVal intContinue As Integer)
    '功能                       写入和条码
    '参数                       Mode =0 绑定条码 =1 生成条码 =2 清除条码 = 3 完成采集 = 4 打印条码或回执单
    '                           WriterUser 是否写入采样人(写入采样人表示采集完成)
    '                           PrintBarCode 是否打印条码
    '                           PrintBackBill 是否打印回执单
    '                           intContinue   1=让步检验
    
    Dim blSelect As Boolean                         '查看是否有选中的行
    Dim Record As ReportRecord                      '列有记录集对像
    Dim strSQL As String                            'SQL语句
    Dim rsTmp As New ADODB.Recordset                '数据记录集
    Dim intLoop As Integer                          '循环变量
    Dim strTmp As String                            '临时字串变量
    Dim varAdvice As Variant                        '一条医嘱里有多个项目时使用
    Dim strCuvetteNumber As String                  '试管编码
    Dim strAdvice As String                         '医嘱ID和相关ID
    Dim varItem As Variant                          '用于分解医嘱ID和相关ID
    Dim strUnion As String                          '合并医嘱ID 相关ID 条码字串并以"|"分隔
    Dim strBarCodePrint As String                   '条码和项目字串用于打印
    Dim varBarcodePrint As Variant                  '条码打印
    Dim strBarCode As String                        '条码字串(分解后)
    Dim strItem As String                           '条码分解后项目
    Dim strBackBill As String                       '回执单医嘱字串
    Dim Control As CommandBarControl                '工具栏控件用于判断使用那种条码
    Dim intBaby As Integer                          '是否是婴儿,>0 表示婴儿数量
    Dim strSample As String                         '标本
    Dim strAdviceContent As String                  '医嘱内容
    Dim lngConnectID As Long                        '相关ID
    Dim varFilter As Variant                        '过滤相同的项目
    Dim blnMsgBox As Boolean                        '是否已显示提示
    Dim strDept As String                           '执行科室
    Dim intExecDept As Integer                      '不区分执行科室打印
    Dim str紧急 As String                           '紧急
    Dim strInfo As String                           '提示信息
    Dim str医嘱ID As String                         '多个医嘱ID用","分隔
    Dim str条码 As String                           '记录当前生成的条码
    Dim rsNumber As ADODB.Recordset                 '生成试管编码数据集
    Dim astrSQL() As String                         'SQL字串
    Dim blnRollBak As Boolean                       '是否是回退事务
    Dim str医嘱ID串 As String                       '医嘱ID串
    Dim blnPrint As Boolean
    
    ReDim astrSQL(0)
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.选择).Checked = True Then
            If chkDept(Record(mAcol.采集科室ID).Value) = False Then
                '记录提示信息在最后进行提示
                Record(mAcol.选择).Checked = False
                strInfo = strInfo & "项目:<" & Record(mAcol.医嘱内容).Value & ">的采集执行科室<" & Record(mAcol.采集执行科室).Value & ">不在你可以操作的科室范围内,不能绑定条码!"
            Else
                str医嘱ID = str医嘱ID & "," & Record(mAcol.ID).Value & "," & Record(mAcol.合并医嘱).Value
            End If
        End If
    Next
    
    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
    
    str医嘱ID = Mid(str医嘱ID, 2)
    str医嘱ID = Replace(Replace(str医嘱ID, ";", ","), "|", ",")
    
    
    '检验划价单费用(生成时才检查)
    If Mode = 0 Or Mode = 1 Then
        If Chk划价费用(Me, str医嘱ID, 0, "E") = False Then
            If strInfo <> "" Then
                MsgBox strInfo
            End If
            Exit Sub
        End If
    End If
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.选择).Checked = True Then
            blSelect = True
            Exit For
        End If
    Next
    

        
    '没有记录时退出
    If blSelect = False Then
        If strInfo <> "" Then
            MsgBox strInfo
        Else
            MsgBox "没有找到可以操作的医嘱内容！", vbInformation, gstrSysName
        End If
        If Me.ChkContinuous.Value = 1 Then
            Me.TxtBarCode.SetFocus
        Else
            Me.txtGoto.SetFocus
        End If
        Exit Sub
    End If
        
    '绑定时查看是否有条码
    If Mode = 0 Then
        If Trim(Me.TxtBarCode.Text) = "" Or Trim(Me.TxtBarCodeCheck.Text) = "" Then
            MsgBox "请扫入条码后再试!", vbInformation, gstrSysName
            Me.TxtBarCode.SetFocus
            Exit Sub
        End If
        
        If Me.TxtBarCode <> Me.TxtBarCodeCheck Then
            MsgBox "两次扫入条码不一致!请重新扫入!", vbInformation, gstrSysName
            Me.TxtBarCode.SetFocus
            Exit Sub
        End If
        
        strSQL = "select 样本条码 from 病人医嘱发送 where 样本条码 = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Me.TxtBarCode)
        If rsTmp.EOF = False Then
            If MsgBox("条码已存在是否确定清除?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Me.TxtBarCode.SetFocus
                Exit Sub
            End If
        End If
        
    End If
    
    InitRecordSet rsNumber
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.选择).Checked = True And chkDept(Record(mAcol.采集科室ID).Value) = True Then
            
            
            '是否区分执行科室打印条码
            Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
            intExecDept = IIf(Control.Checked, 0, 1)
            
            
            Select Case Mode
                Case 0                          '绑定
                    MakeBarCode rsNumber, Record, Mode, intExecDept, Me.TxtBarCode.Text
                Case 1                          '生成
                    MakeBarCode rsNumber, Record, Mode, intExecDept
                Case 2                          '取消
                    MakeBarCode rsNumber, Record, Mode
                Case 3, 4                       '完成、打印
                    MakeBarCode rsNumber, Record, Mode
                
            End Select
            
        End If
    Next
    
    On Error GoTo errH
    
    If rsNumber.RecordCount = 0 Then Exit Sub
    rsNumber.MoveFirst
    Select Case Mode
        Case 0, 1                                   '绑定或生成条码
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_条码生成('" & rsNumber("医嘱ID串") & "','" & rsNumber("样本条码") & "')"
                If WriterUser = True Then
                    '执行完成
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("类别") & "" = "输血" Then
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                    End If
                End If
                rsNumber.MoveNext
            Loop
        Case 2                                     '取消完成或条码绑定
            Do Until rsNumber.EOF
                If TabCtr.Selected.Index = 2 Then
                    '取消采集
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("类别") & "" = "输血" Then
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',1,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',1)"
                    End If
                    If chkComPlete.Value = 1 Then
                        '根据参数看是否取消绑定
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_条码生成('" & rsNumber("医嘱ID串") & "')"
                    End If
                Else
                    '取消绑定
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_条码生成('" & rsNumber("医嘱ID串") & "')"
                    If TabCtr.Selected.Index = 5 Then
                        intContinue = 2
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        If rsNumber("类别") & "" = "输血" Then
                            astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',1,1)"
                        Else
                            astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',1)"
                        End If
                    End If
                End If
                rsNumber.MoveNext
            Loop
        Case 3                                      '完成采集和取消完成、重采标本
            Do Until rsNumber.EOF
                If TabCtr.Selected.Index = 4 Then
                    '重采时先取消标本送检
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_Lis预置条码_标本送出('" & rsNumber("医嘱ID串") & "',1)"
                    '重新采集
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("类别") & "" = "输血" Then
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                    End If
                Else
                    '让步检验时送检标本
                    If intContinue = 1 Then
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Zl_Lis预置条码_标本送出('" & rsNumber("医嘱ID串") & "',0)"
                    ElseIf TabCtr.Selected.Index = 5 And rsNumber("样本条码") <> "" Then
                        intContinue = 3
                    End If
                    
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("类别") & "" = "输血" Then
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                    End If
                End If
                rsNumber.MoveNext
            Loop
    End Select
    
    gcnOracle.BeginTrans
    blnRollBak = True
    
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans
    
    If ((Mode = 0 Or Mode = 1) And WriterUser = True) Or Mode = 3 Or Mode = 2 Then
        Call WriterBarCodeToLIS(rsNumber, IIf(Mode = 2, 2, 3), intContinue)
    ElseIf (Mode = 0 Or Mode = 1) Then
        Call WriterBarCodeToLIS(rsNumber, 3, 4)
    End If
    
    '打印条码
    If PrintBarCode = True And intContinue <> 1 Then
        
        blnPrint = CheckPlugIn(glngSys, glngModul, rsNumber)
        If blnPrint = True Then
            rsNumber.MoveFirst
            Do Until rsNumber.EOF
                '成生条码到PIC
                Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Tool_SignNew, True, True)
                If Control.Checked = True Then
                    Bar39 Me.picBarCodePrint, 3, Nvl(rsNumber("样本条码")), False, True
                Else
                    Bar128 Me.picBarCodePrint, 3, Nvl(rsNumber("样本条码")), True
                End If
                SavePicture Me.picBarCodePrint.Image, App.path & "\BarCode.Bmp"
                '开始打印
                Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me, "样本条码=" & Nvl(rsNumber("样本条码")), _
                "项目=" & Replace(Nvl(rsNumber("医嘱内容")), ",", " "), _
                "病人姓名 = " & IIf(Nvl(rsNumber("婴儿"), 0) = 0, IIf(txt姓名 <> "", txt姓名, "无"), rsNumber("婴儿姓名")), _
                "性别 = " & IIf(Nvl(rsNumber("婴儿"), 0) = 0, IIf(cbo性别 <> "", cbo性别, "无"), rsNumber("婴儿性别")), _
                "年龄 = " & IIf(Nvl(rsNumber("婴儿"), 0) = 0, IIf(txt年龄 & cboAge <> "", txt年龄 & cboAge & txt年龄1, "无"), "婴"), _
                "床号 = " & IIf(txtBed <> "", txtBed, "无"), _
                "标识号 = " & IIf(txtID <> "", txtID, "无"), _
                "所在科室 = " & IIf(Nvl(rsNumber("所在科室")) <> "", Nvl(rsNumber("所在科室")), "无"), _
                "采集方式 = " & IIf(Nvl(rsNumber("采集方式")) <> "", Nvl(rsNumber("采集方式")), "无"), _
                "标本 = " & IIf(Nvl(rsNumber("标本")) <> "", Nvl(rsNumber("标本")), "无"), _
                "执行科室 = " & IIf(Nvl(rsNumber("执行科室")) <> "", Nvl(rsNumber("执行科室")), "无"), _
                "开嘱医生 = " & IIf(Nvl(rsNumber("开嘱医生")) <> "", Nvl(rsNumber("开嘱医生")), "无"), _
                "开嘱时间 = " & IIf(Nvl(rsNumber("开嘱时间")) <> "", Nvl(rsNumber("开嘱时间")), "无"), _
                "采样人 = " & IIf(Nvl(rsNumber("采样人")) <> "", Nvl(rsNumber("采样人")), "无"), _
                "采样时间 = " & IIf(Nvl(rsNumber("采样时间")) <> "", Nvl(rsNumber("采样时间")), "无"), _
                "管码 = " & IIf(Nvl(rsNumber("管码")) <> "", Nvl(rsNumber("管码")), "无"), _
                "采血量 = " & IIf(Nvl(rsNumber("采血量")) <> "", Nvl(rsNumber("采血量")), "无"), _
                "试管名称 = " & IIf(Nvl(rsNumber("试管名称")) <> "", Nvl(rsNumber("试管名称")), "无"), _
                "紧急 = " & IIf(Nvl(rsNumber("紧急标志")) <> "", Nvl(rsNumber("紧急标志")), "无"), _
                "病人来源 = " & IIf(Nvl(rsNumber("病人来源")) <> "", Nvl(rsNumber("病人来源")), "无"), _
                "条码图像1=" & App.path & "\BarCode.Bmp", 2)
                '删除条码图像
                Kill App.path & "\BarCode.Bmp"
                strSQL = "Zl_Lis预置条码_条码打印('" & Replace(rsNumber("医嘱ID串"), ",,", ",") & "')"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
                rsNumber.MoveNext
            Loop
        End If
    End If
    
    '打印回执单
    If PrintBackBill = True Then
        rsNumber.MoveFirst
        Do Until rsNumber.EOF
            str医嘱ID串 = str医嘱ID串 & "," & rsNumber("医嘱ID串")
            rsNumber.MoveNext
        Loop
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_2", Me, "病人ID=" & mlngKey, "医嘱ID串=" & Mid(str医嘱ID串, 2), 2)
    End If
    
    '刷新数据
    If mblFind = False Then
        '非查找的病人
        With Me.rptPlist
            If .Rows.Count > 0 Then
    '            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index
                If .FocusedRow Is Nothing Then
                    .FocusedRow = Me.rptPlist.Rows(0)
                    .Populate
                Else
                    RefreshAdviceData .FocusedRow.Record(mPcol.病人ID).Value, Me.TabCtr.Selected.Index, IIf(.FocusedRow.Record(mPcol.来源).Value = "住院", 2, 1), False, .FocusedRow.Record(mPcol.主页ID).Value
                End If
            Else
                Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
                Me.rptCuvette.Records.DeleteAll
                Me.rptAlist(Me.TabCtr.Selected.Index).Populate
                Me.rptCuvette.Populate
                txt姓名 = ""
                txt姓名.Tag = ""
                cbo性别.ListIndex = -1
                txt年龄 = ""
                txt年龄1 = ""
                txtBed = ""
                txtID = ""
                txtPatientDept = ""
                cbo开单科室.ListIndex = -1
                cbo医生.ListIndex = -1
                txt医嘱内容.Text = ""
                txt医嘱内容.Tag = ""
                Me.lblCap(6).Visible = False
            End If
        End With
    
        Me.TxtBarCode.Text = ""
        Me.TxtBarCode.Tag = ""
        Me.TxtBarCodeCheck.Text = ""
        
        If Me.ChkContinuous.Value = 1 Then
            Me.TxtBarCode.SetFocus
        Else
            Me.txtGoto.SetFocus
        End If
    Else
        '查找病人
        With Me.rptPlist
            If .Rows.Count > 0 Then
    '            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index
                If .FocusedRow Is Nothing Then
                    .FocusedRow = Me.rptPlist.Rows(0)
                    .Populate
                Else
                    RefreshAdviceData .FocusedRow.Record(mPcol.病人ID).Value, Me.TabCtr.Selected.Index, IIf(.FocusedRow.Record(mPcol.来源).Value = "住院", 2, 1), False, .FocusedRow.Record(mPcol.主页ID).Value
                End If
            Else
                Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
                Me.rptCuvette.Records.DeleteAll
                Me.rptAlist(Me.TabCtr.Selected.Index).Populate
                Me.rptCuvette.Populate
                txt姓名 = ""
                txt姓名.Tag = ""
                cbo性别.ListIndex = -1
                txt年龄 = ""
                txt年龄1 = ""
                txtBed = ""
                txtID = ""
                txtPatientDept = ""
                cbo开单科室.ListIndex = -1
                cbo医生.ListIndex = -1
                txt医嘱内容.Text = ""
                txt医嘱内容.Tag = ""
                Me.lblCap(6).Visible = False
            End If
        End With
        
        Me.TxtBarCode.Text = ""
        Me.TxtBarCode.Tag = ""
        Me.TxtBarCodeCheck.Text = ""
        
        If Mode = 1 Then
            If chkComPlete.Value = 1 Then
                Me.cmdBindBarCode.Enabled = False
                Me.cmdNewBarcode.Enabled = False
                Me.cmdComplete.Enabled = True
                Me.cmdBarcodePrint.Enabled = True
                Me.cmdBakBillPrint.Enabled = True
                Me.cmdBindBarCode.Caption = "清除绑定(&B)"
                Me.cmdNewBarcode.Caption = "取消条码(&N)"
                Me.cmdComplete.Caption = "取消完成(&P)"
            Else
                Me.cmdBindBarCode.Enabled = True
                Me.cmdNewBarcode.Enabled = True
                Me.cmdComplete.Enabled = True
                Me.cmdBarcodePrint.Enabled = True
                Me.cmdBakBillPrint.Enabled = True
                Me.cmdBindBarCode.Caption = "清除绑定(&B)"
                Me.cmdNewBarcode.Caption = "取消条码(&N)"
                Me.cmdComplete.Caption = "完成采集(&P)"
            End If
        End If
        
        If strInfo <> "" Then
            MsgBox strInfo
        End If
        
        If Me.ChkContinuous.Value = 1 Then
            Me.TxtBarCode.SetFocus
        Else
            Me.txtGoto.SetFocus
        End If
    End If
    
    '生成或绑定条码后跳转到已绑定页
    If chkBindPage.Value = 1 Then
        If Mode = 0 Or Mode = 1 Then TabCtr.Item(1).Selected = True
    End If
    
    Exit Sub
errH:
    If blnRollBak = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckPlugIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal rsMoneyNow As ADODB.Recordset) As Boolean
'    rsNumber.Fields.Append "类别", adVarChar, 20
'    rsNumber.Fields.Append "管码", adVarChar, 18
'    rsNumber.Fields.Append "相关ID", adBigInt
'    rsNumber.Fields.Append "样本条码", adVarChar, 18
'    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
'    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
'    rsNumber.Fields.Append "婴儿", adBigInt
'    rsNumber.Fields.Append "紧急标志", adBigInt
'    rsNumber.Fields.Append "标本", adVarChar, 30
'    rsNumber.Fields.Append "医嘱内容", adVarChar, 500
'    rsNumber.Fields.Append "采集方式", adVarChar, 100
'    rsNumber.Fields.Append "开嘱医生", adVarChar, 50
'    rsNumber.Fields.Append "开嘱时间", adDate
'    rsNumber.Fields.Append "采样人", adVarChar, 50
'    rsNumber.Fields.Append "采样时间", adDate
'    rsNumber.Fields.Append "采血量", adVarChar, 20
'    rsNumber.Fields.Append "试管名称", adVarChar, 50
'    rsNumber.Fields.Append "病人来源", adInteger
'    rsNumber.Fields.Append "医嘱ID串", adVarChar, 500
'    rsNumber.Fields.Append "执行科室", adVarChar, 50
'    rsNumber.Fields.Append "婴儿姓名", adVarChar, 50
'    rsNumber.Fields.Append "婴儿性别", adVarChar, 50
'    rsNumber.Fields.Append "申请科室", adVarChar, 50
    
    Dim blnTmp As Boolean
        On Error Resume Next
        CheckPlugIn = True
        If Not mobjZLIHISPlugIn Is Nothing Then
            blnTmp = mobjZLIHISPlugIn.LisPrintCodeBefore(lngSys, lngModual, rsMoneyNow)
            Call zlPlugInErrH(Err, "LisPrintCodeBefore")
            If Err.Number <> 0 Then
                '接口出错了,继续打印
                blnTmp = True
            End If
        Else
            blnTmp = True
        End If
        CheckPlugIn = blnTmp
    Err.Clear: On Error GoTo 0

End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdNewBarcode_GotFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (Me.ActiveControl Is cmdNewBarcode And txtGoto.Tag <> "")
End Sub

Private Sub cmdNewBarcode_LostFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub txtGoto_Change()
'    If Me.ActiveControl Is txtGoto Then
'        If IDKind.IDKind = IDKinds.C0姓名 Then
'            If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
'        ElseIf IDKind.IDKind = IDKinds.C3IC卡号 Then
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
'        End If
'    End If
End Sub

Private Sub txtGoto_GotFocus()
    If Me.Visible = False Then Exit Sub
    txtGoto.SelStart = 0
    txtGoto.SelLength = Len(txtGoto.Text)
    If txtGoto.Text = "" And Not txtGoto.Locked Then
'        If IDKind.IDKind = IDKinds.C0姓名 Then
'            If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (True)
'        ElseIf IDKind.IDKind = IDKinds.C3IC卡号 Then
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (True)
'        ElseIf IDKind.IDKind = IDKinds.C2身份证号 Then
'            If Not mobjIDCard Is Nothing And txtGoto.Text = "" And Not txtGoto.Locked Then mobjIDCard.SetEnabled (True)
'        End If
    End If
End Sub
Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blFind As Boolean                           '是否查找成功
    Dim Row As ReportRow                            '列表行对象
    Dim strFind As String                           '临进查找字串
    Dim blnBarCode As Boolean                       '是否条码
    Dim blnNo As Boolean                            '是否单据号
    Dim blnCard As Boolean
    Dim lng卡类别ID As Long
    Dim lng病人ID As Long
    Dim strTmp As String
    Dim NowDate As Date
    Dim strDateBegin As Date
    Dim strDateEnd As Date
    
    On Error GoTo errH
    
    
    If Trim(Me.txtGoto.Text) = "" Then Me.txtGoto.SetFocus: Exit Sub
'    blnCard = zlCommFun.InputIsCard(txtGoto, KeyAscii, mblnShowPwd)
    mstrIndex = IDKind.IDKind
    If IDKind.IDKind = IDKind.GetKindIndex("姓名") Then
'        blnCard = zlCommFun.InputIsCard(txtGoto, KeyAscii, False)
    End If
    If IDKind.IDKind = IDKind.GetKindIndex("就诊卡") Then
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If CheckIsInclude(UCase(Chr(KeyAscii)), "'‘’;；:：?？|,，。""") = True Then KeyAscii = 0
        'gbytCardNOLen = Val(IDKind.GetKindItem("卡号长度", IDKind.IDKind))
        gbytCardNOLen = IDKind.GetCardNoLen
'        Call zlCommFun.InputIsCard(txtGoto, KeyAscii, True)
        txtGoto.Text = ReplaseSpecial(txtGoto.Text)
        blnCard = KeyAscii <> 8 And Len(txtGoto.Text) = gbytCardNOLen - 1 And txtGoto.SelLength <> Len(txtGoto.Text)
        If blnCard = True And KeyAscii <> 0 Then
            If KeyAscii <> 13 Then
                Me.txtGoto = Me.txtGoto & Chr(KeyAscii)
            End If
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Or (IDKind.IDKind = IDKind.GetKindIndex("就诊卡") And blnCard = True) Then
        '刷新后再查找
'        RefreshPatientData
        '先清空再定入
        txt姓名 = ""
        txt姓名.Tag = ""
        cbo性别.ListIndex = -1
        txt年龄 = ""
        txt年龄1 = ""
        txtBed = ""
        txtID = ""
        lblCap(0).Caption = "标  识 号"
        txtPatientDept = ""
        cbo开单科室.ListIndex = -1
        cbo医生.ListIndex = -1
        txt医嘱内容.Text = ""
        txt医嘱内容.Tag = ""
        Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
        Me.rptCuvette.Records.DeleteAll
        Me.rptAlist(Me.TabCtr.Selected.Index).Populate
        Me.rptCuvette.Populate
        
        If mbln身份证 Or IDKind.IDKind = IDKind.GetKindIndex("身份证号") Then
'            strsql = "select 病人ID from 病人信息 where 身份证号 = [1] "
'            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, txtGoto)
'            If Not rsTmp.EOF Then
'                txtGoto = "-" & rsTmp.Fields("病人ID")
'            End If
            If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
            If lng病人ID <> 0 Then
                txtGoto = "-" & lng病人ID
            End If
        ElseIf IDKind.IDKind = IDKinds.C1医保号 Then
            
        End If
    
    
        Select Case Mid(txtGoto, 1, 1)
            Case "-"                                '病人ID
                blFind = RefreshAdviceData(Val(Mid(txtGoto, 2)), Me.TabCtr.Selected.Index, 1, True)
                strFind = Val(Mid(txtGoto, 2))
            Case "+"                                '住院号
                strSQL = "select 病人ID from 病人信息 where 住院号 = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 2, True)
                strFind = Val(Mid(txtGoto, 2))
            Case "*"                                '门诊号
                strSQL = "select 病人ID from 病人信息 where 门诊号 = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Val(Mid(txtGoto, 2))
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
                blnNo = True
            
            Case Else                               '就诊卡和姓名
                strFind = txtGoto
                If IDKind.IDKind = IDKind.GetKindIndex("姓名") And BlnIsNumber(txtGoto) Then
                    strSQL = "select a.病人id,a.病人来源 from 病人医嘱记录 a , 病人医嘱发送 b " & _
                         " Where a.ID = b.医嘱id And b.样本条码 = [1] order by a.开嘱时间 desc    "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, txtGoto)
                    If rsTmp.EOF = False Then
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, Nvl(rsTmp(1), 0), True)
                        blnBarCode = True
                    Else
                        strSQL = "select 病人ID from 病人信息 where 就诊卡号 = [1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, txtGoto)
                        If rsTmp.EOF = False Then
                            blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                        End If
                    End If
                    
                Else
'                    MsgBox IDKind.GetKindIndex("就诊卡")
                    If blnCard Or IDKind.IDKind = IDKind.GetKindIndex("就诊卡") Then
                        strSQL = "select 病人ID from 病人信息 where 就诊卡号 = [1] "
                        strFind = UCase(txtGoto)
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("医保号") Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("身份证号") Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("IC卡号") Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Then
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        strFind = lng病人ID
                        'strFind = Val(txtGoto)
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("姓名") Then '通过姓名检索时，应注意当前时间范围
'                        strSQL = "select 病人ID from 病人信息 where 姓名 = [1] "
'                        strSQL = strSQL & " or 就诊卡号 = [1] "
                        
                        If rptPlist.Tag = "" Then
                            strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "") '从注册表中读取过滤条件
                            NowDate = zlDatabase.Currentdate
                            strDateBegin = CDate(Format(NowDate - Val(Split(strTmp, ";")(9)), "yyyy-mm-dd 00:00:00"))
                            strDateEnd = CDate(Format(NowDate, "yyyy-mm-dd 23:59:59"))
                        Else
                            strDateBegin = CDate(Format(Split(rptPlist.Tag, ";")(11), "yyyy-mm-dd 00:00:00"))
                            strDateEnd = CDate(Format(Split(rptPlist.Tag, ";")(12), "yyyy-mm-dd 23:59:59"))
                        End If
                        
                        strSQL = "Select Distinct a.病人id" & vbNewLine & _
                            "From 病人信息 A, 病人医嘱记录 B, 病人医嘱发送 C" & vbNewLine & _
                            "Where a.病人id = b.病人id And b.Id = c.医嘱id And c.发送时间+0 Between To_Date('" & strDateBegin & "', 'yyyy-mm-dd hh24:mi:ss') And" & vbNewLine & _
                            "      To_Date('" & strDateEnd & "', 'yyyy-mm-dd hh24:mi:ss') And (a.姓名 = [1] Or a.就诊卡号 = [1]) "
                    Else
                        If IDKind.GetCurCard.接口序号 <> 0 Then
                            lng卡类别ID = IDKind.GetCurCard.接口序号
                            If mobjSquareCard.zlGetPatiID(lng卡类别ID, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                            If lng病人ID = 0 Then lng病人ID = 0
                        Else
                            If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, txtGoto, False, lng病人ID) = False Then lng病人ID = 0
                        End If
                        strSQL = "select 病人ID from 病人信息 where 病人ID = [1] "
                        strFind = lng病人ID
                    End If
                  
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind)
                    If rsTmp.RecordCount = 0 Then
                        If IDKind.IDKind = IDKind.GetKindIndex("姓名") Then
                            strFind = Trim(strFind)
                            strFind = Replace(strFind, Chr(&HD), "")
                            If IDKind.GetKindIndex("就诊卡") <> -1 Then lng卡类别ID = IDKind.GetIDKindCard("就诊卡").接口序号
                            strSQL = "select 病人ID from 病人医疗卡信息 where 卡号 = [1] and 卡类别id =[2] "
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind, lng卡类别ID)
                        End If
                    End If
                    If rsTmp.EOF = False Then
                        strFind = rsTmp(0)
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                    End If
                    
                End If
                    
                
        End Select
        
        On Error Resume Next
        '下面用于定位,如出错就忽略
        mblFind = blFind                '用于判断是否查找到的病人
        If blFind = True Then
            '查找成功
            If Me.chkFindMove.Value = 1 Then
                Me.TxtBarCode.SetFocus
            Else
                If cmdNewBarcode.Enabled = True And cmdNewBarcode.Caption = "生成条码(&N)" Then
                    cmdNewBarcode.SetFocus
                ElseIf cmdComplete.Enabled = True And cmdComplete.Caption = "完成采集(&P)" Then
                    Me.cmdComplete.SetFocus
                Else
                    cmdNewBarcode.SetFocus
                End If
            End If
            FindPatient txtGoto.Text
        Else
            '没有找到病人时再在列表中查找一次
            If FindPatient(txtGoto.Text) = False Then
                '没有查找到病人
                Me.txtGoto.SelStart = 0
                Me.txtGoto.SelLength = Len(Me.txtGoto.Text)
                Me.txtGoto.SetFocus
            Else
                '条码和单据时只查出对应的记录
                If blnNo = True Or blnBarCode = True Then
                     If rsTmp.EOF = False Then
                        If blnNo = True Then
                            Call RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                        Else
                            Call RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, Nvl(rsTmp(1), 1), True)
                        End If
                     End If
                End If
                
                If Me.chkFindMove.Value = 1 Then
                    Me.TxtBarCode.SetFocus
                Else
                    If cmdNewBarcode.Enabled = True And cmdNewBarcode.Caption = "生成条码(&N)" Then
                        cmdNewBarcode.SetFocus
                    Else
                        cmdComplete.SetFocus
                    End If
                End If
            End If
        End If
        Me.txtGoto.Text = ""
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FindPatient(strFind As String) As Boolean
    '查找病人
    '返回                   '成功返回True 失败返回False
    Dim Row  As ReportRow
    Dim intLoop As Integer
    
    For intLoop = 0 To 2
        For Each Row In Me.rptPlist.Rows '门诊号、住院号、病人id、姓名、就诊卡号、条码
            If (Trim(strFind) = "*" & Trim(Row.Record(mPcol.标识号).Value) And InStr("门诊,体检", Trim(Row.Record(mPcol.来源).Value)) > 0) Or _
                (Trim(IDKind.IDKind) = Trim(IDKind.GetKindIndex("门诊号")) And Trim(strFind) = Trim(Row.Record(mPcol.标识号).Value) And InStr("门诊,体检", Trim(Row.Record(mPcol.来源).Value)) > 0) Or _
                (Trim(strFind) = "+" & Trim(Row.Record(mPcol.标识号).Value) And "住院" = Trim(Row.Record(mPcol.来源).Value)) Or _
                Trim(strFind) = "-" & Trim(Row.Record(mPcol.病人ID).Value) Or Trim(Row.Record(mPcol.病人姓名).Value) Like Trim(strFind) & "*" Or _
                Trim(strFind) = Trim(Row.Record(mPcol.就诊卡).Value) Or InStr(1, "," & Trim(Row.Record(mPcol.条码).Value) & ",", "," & Trim(strFind) & ",") > 0 Then
                
                Set Me.rptPlist.FocusedRow = Row
                Me.rptPlist.Populate
                FindPatient = True
                
                Exit Function
            End If
        Next
    Next
End Function

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
    objPrint.Title.Text = "病历文件清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub txtGoto_LostFocus()
    If IDKind.IDKind = IDKinds.C0姓名 Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    ElseIf IDKind.IDKind = IDKinds.C3IC卡号 Then
        If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    ElseIf IDKind.IDKind = IDKinds.C2身份证号 Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    End If
End Sub


Private Sub txt年龄_GotFocus()
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txt医嘱内容)) > 0 Then
            If Me.cmdNewBarcode.Enabled = True Then
                Me.cmdNewBarcode.SetFocus
            End If
        Else
            zlCommFun.PressKey vbKeyTab
        End If
        Exit Sub
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
    Else
        Call AdjustEditState(True)
'        Me.cbo性别.SetFocus
        zlCommFun.PressKey vbKeyTab
    End If
End Sub






Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldText As String
    
    On Error GoTo errH
    strOldText = Me.cbo开单科室.Text
    Me.cbo开单科室.Clear
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检'))" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cbo开单科室.AddItem rsTmp!名称
        cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    
    On Error Resume Next
    Me.cbo开单科室.Text = strOldText
    If cbo开单科室.ListCount > 0 And Me.cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDoctors(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cbo医生.Clear
    
    '科室医生或护士
    strSQL = _
        "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
        " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        " And C.人员性质 IN('医生') And B.部门ID=[1] " & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
        
    strSQL = strSQL & " Order by 简码,人员性质 Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo医生.AddItem rsTmp!姓名
            cbo医生.ItemData(cbo医生.ListCount - 1) = rsTmp!部门ID
            
            If rsTmp!ID = UserInfo.ID And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = cbo医生.NewIndex
            rsTmp.MoveNext
        Next
        
        If cbo医生.ListCount = 1 And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = 0
    End If
End Sub


Private Function Get检查手术名称(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int类型 = 1 Then Get检查手术名称 = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            strTmp = strTmp & "," & rsRelativeAdvice("名称")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get检查手术名称 = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " 及 ") & Mid(strTmp, 2)
    Else
        Get检查手术名称 = txtMainAdvice
    End If
End Function

Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(strNO, "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=" & intNum
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic")
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(strNO, "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt医嘱内容_GotFocus()
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt医嘱内容.Text = txt医嘱内容.Tag Then
'            zlcommfun.PressKey vbKeyTab
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            If chkFindMove.Value = 0 Then
                cmdNewBarcode.SetFocus
            Else
                Me.TxtBarCode.SetFocus
            End If
            Exit Sub
        End If
        
        With txt医嘱内容
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
        '新项目的录入
        '根据选择项目设置缺省医嘱信息
        If AdviceInput(rsTmp) Then
            DoEvents
            '显示已缺省设置的值
            txt医嘱内容.Tag = txt医嘱内容.Text
        Else
            DoEvents
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容

            txt医嘱内容.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub
Private Function SelectDiagItem() As ADODB.Recordset
'选择检验项目
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID "
    strSQL = strSQL + "From 诊疗项目目录 A,诊疗项目别名 C,诊疗执行科室 D Where A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID And A.类别='C' "       'And D.执行科室ID=" & mlngDeptID
    strSQL = strSQL + " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And (A.编码 Like '" + txt医嘱内容 + "%' Or Upper(A.名称) Like '" + txt医嘱内容 + "%' Or Upper(C.简码) Like '" + UCase(txt医嘱内容) + "%')"
            
    Call ClientToScreen(txt医嘱内容.hWnd, objPoint)
    Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "选择申请项目", True, Me.txt医嘱内容.Text, "", True, True, True, objPoint.X * 15, objPoint.Y * 15, Me.txt医嘱内容.Height, False, True)
End Function

Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集
'返回：本次录入是否有效
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim strSQL As String
    Dim strExtData As String
    Dim blnOk As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    
    
    On Error GoTo errH
    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    If Not rsInput Is Nothing Then txt医嘱内容.Text = rsInput!名称    '暂时显示

    '需要输入更多数据的一些项目
    '---------------------------------------------------------------------------------------------------------------
    '检验项目选择检验标本
    strHelpText = "检验项目"
    
    If Not rsInput Is Nothing Then
        strExtData = rsInput!诊疗项目ID & ";" & rsInput!标本部位    '新输入项目
    Else
        strExtData = mstrExtData   '新输入项目
    End If
    
    On Error Resume Next
    '接口改造：int场合没有传，现在传为0， bytUseType 以前没传，现在传为0
    blnOk = frmAdviceEditEx.ShowMe(Me, Me.txt医嘱内容.hWnd, t_Pati, 0, 4, 0, 1, PatientType, , , , 0, strExtData, , , , , True)
    On Error GoTo errH

    If Not blnOk Then Exit Function
    If strExtData = "" Or Mid(strExtData, 1, 1) = ";" Then Exit Function
    
    '获取采集方式
    Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    mlngCapID = rsTmp("ID")
    
    strSQL = "Select C.项目类别 From 诊疗项目目录 A,检验报告项目 B,检验项目 C " & _
        "Where A.ID=B.诊疗项目ID And B.报告项目ID=C.诊治项目ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp.EOF Then
        mbln微生物项目 = False
    Else
        mbln微生物项目 = IIf(Nvl(rsTmp("项目类别"), 0) = 2, True, False)
    End If
    
    mstrExtData = strExtData
    
    
    Call AdviceSet检查手术(3, mstrExtData)
    txt医嘱内容.Text = Get检查手术名称(2, "")
    txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    
    '开嘱医生
    On Error Resume Next
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function SelectCap(Optional ByVal lngItemID As Long = 0) As ADODB.Recordset
'获取采集方式
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
        
    strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
        "From 诊疗项目目录 A,诊疗用法用量 D Where A.ID=D.用法ID" + _
        " And A.类别='E' And A.操作类型='6'" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
        IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And D.项目ID=" & lngItemID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A Where " + _
            " A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
            " And Nvl(A.执行频率,0) IN(0,1)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set SelectCap = rsTmp
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '处理检验项目
    strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
        "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室,操作类型 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        Set rsRelativeAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function
'检查医嘱内容的合法性
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
    If txt姓名.Text = "" Then
        ValidAdvice = False
        MsgBox "请输入病人的姓名！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.姓名
        txt姓名.SetFocus: Exit Function
    End If
    
    If Len(Trim(Me.txt医嘱内容)) = 0 Then
        ValidAdvice = False
        MsgBox "必须输入申请项目！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.医嘱内容
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Me.cbo开单科室.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "请指定开单科室！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.开单科室
        Me.cbo开单科室.SetFocus: Exit Function
    End If
    If Len(Trim(Me.cbo医生.Text)) = 0 Then
        ValidAdvice = False
        MsgBox "请指定开单医生！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.医生
        Me.cbo医生.SetFocus: Exit Function
    End If
End Function


Private Function SaveAdviceData() As Long
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim iMaxSeq As Integer, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng开嘱科室ID As Long, lng病人ID As Long, strDoctor As String, i As Integer
    Dim str执行科室ID As String, str执行科室ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr类别 As String, tmplngClinicID As Long, tmpint计价特性 As Integer, tmpint执行性质 As Integer
    Dim rsDept As ADODB.Recordset
    Dim intPatientSource As Integer                     '病人来源
    Dim astrSQL() As String
    Dim blnRollBack As Boolean
    Dim intLoop As Integer
    Dim strCostType As String, lngJ As Long
    On Error GoTo ErrHand
    ReDim astrSQL(0)

    On Error GoTo ErrHand
    
    
    '保存病人信息
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If PatientType = 1 Then '门诊病人
        If mlng病人ID > 0 Then '已有的病人
'            strsql = _
'                "zl_挂号病人病案_INSERT(3," & mlng病人ID & ",Null," & _
'                "'',''," & _
'                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Me.cboAge.Text & Me.txt年龄1.Text & "'," & _
'                "'自费','自费'," & _
'                "'','',''," & _
'                "'','','',0,'','','','',''," & strDate & ",NULL)"
        Else '新病人
            '添加获取默认费别
            strSQL = "select 名称,缺省标志 from 费别 order by 编码 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork")
            Do While Not rsTmp.EOF
                lngJ = lngJ + 1
                If lngJ = 1 Then
                    strCostType = rsTmp("名称")
                End If
                If rsTmp("缺省标志") = 1 Then
                    strCostType = rsTmp("名称")
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If strCostType = "" Then strCostType = "自费"
        
            mlng病人ID = zlDatabase.GetNextNo(1)
            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
            astrSQL(UBound(astrSQL)) = _
                "zl_挂号病人病案_INSERT(1," & mlng病人ID & ",Null," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Me.cboAge.Text & Me.txt年龄1.Text & "'," & _
                "'" & strCostType & "','" & strCostType & "'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
        End If
    End If
    '保存医嘱并发送
    lngAdviceID = zlDatabase.GetNextId("病人医嘱记录")
    iMaxSeq = 0
    
    lng开嘱科室ID = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    strDoctor = NeedName(Me.cbo医生.Text)
    
    If rsRelativeAdvice.RecordCount = 0 Then
        str执行科室ID = mlngDeptID
    Else
        'PatientType
        If mlng病人ID > 0 Then
            strSQL = "select  执行科室ID from  诊疗执行科室 where 病人来源 = [1] and 诊疗项目ID = [2] "
        Else
            strSQL = "select 执行科室id from 诊疗执行科室 where 诊疗项目id = [2]"
        End If
        rsRelativeAdvice.MoveFirst
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, PatientType, CLng(rsRelativeAdvice("Id")))
        str执行科室ID = Val(Nvl(rsTmp("执行科室ID")))
    End If
    
    iSendSeq = 1
    '检验项目将采集方式作为主医嘱
    tmplngClinicID = mlngCapID
    '取采集方式的执行部门
    str执行科室ID1 = "NULL"
    
    lngSendNO = zlDatabase.GetNextNo(10)
    strNO = zlDatabase.GetNextNo(IIf(PatientType = 2, 14, 13))
    
    '保存相关医嘱
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("病人医嘱记录")
            With rsRelativeAdvice
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    (iMaxSeq + i) & ",3," & mlng病人ID & ",NULL," & _
                    "0,1," & _
                    "1,'" & .Fields("类别") & "'," & _
                    .Fields("ID") & ",NULL,NULL,NULL,NULL," & _
                    "'" & Replace(.Fields("名称"), "'", "''") & "',''," & _
                    "'" & .Fields("标本部位") & "','一次性',NULL,NULL,'',NULL," & _
                    .Fields("计价性质") & "," & _
                    str执行科室ID & "," & _
                    .Fields("执行科室") & ",0," & strDate & ",NULL," & _
                    IIf(Me.txtPatientDept.Tag = 0, lng开嘱科室ID, Me.txtPatientDept.Tag) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                    "Sysdate,'',Null)"
                iSendSeq = iSendSeq + 1
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_病人医嘱发送_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & ",NULL,NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & str执行科室ID & ",0,0)"
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    '检验申请的采集方式放到最后
    iMaxSeq = iMaxSeq + 1
    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
    astrSQL(UBound(astrSQL)) = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
        iMaxSeq & ",3," & mlng病人ID & ",NULL," & _
        "0,1," & _
        "1,'E'," & mlngCapID & ",NULL,NULL,NULL,NULL," & _
        "'" & Replace(Me.txt医嘱内容, "'", "''") & "',''," & _
        "'','一次性',NULL,NULL,'',NULL,2," & _
        str执行科室ID1 & ",3,0," & strDate & ",NULL," & _
        IIf(Me.txtPatientDept.Tag = 0, lng开嘱科室ID, Me.txtPatientDept.Tag) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
        "Sysdate,'',Null)"
    iSendSeq = iSendSeq + 1
    '发送主医嘱
    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
    astrSQL(UBound(astrSQL)) = "ZL_病人医嘱发送_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & ",NULL,NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & str执行科室ID & ",0,1)"
    
    gcnOracle.BeginTrans
    blnRollBack = True
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), "写入医嘱"
        End If
    Next
    gcnOracle.CommitTrans
    SaveAdviceData = mlng病人ID
    Exit Function
ErrHand:
    mlng病人ID = 0
    If blnRollBack = True Then
        gcnOracle.RollbackTrans
    End If
'    Err.Raise Err.Number, "标本核收"
    Exit Function
End Function

Private Sub AdjustEditState(blEnable As Boolean)
    '功能:              调整编辑状态
    'Me.txt姓名.Enabled = blEnable
    cbo性别.Enabled = blEnable
    txt年龄.Enabled = blEnable
    txt年龄1.Enabled = blEnable
    cboAge.Enabled = blEnable
    cbo开单科室.Enabled = blEnable
    cbo医生.Enabled = blEnable
    txt医嘱内容.Enabled = blEnable
    cmdSelect.Enabled = blEnable
End Sub

Private Sub HideBarCode()
    '隐藏预置条码
    Dim Control As CommandBarControl
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_PriceTable, True, True)
    
    If Control.Checked = True Then
        Me.Frame1.Visible = False
        Me.Frame2.Top = Me.Frame1.Top - 20
        Me.fraBarCode.Height = Me.Frame2.Top + Me.Frame2.Height + 120
    Else
        Me.Frame1.Visible = True
        Me.Frame2.Top = Me.Frame1.Top + Me.Frame1.Height + 20
        Me.fraBarCode.Height = Me.Frame2.Top + Me.Frame2.Height + 120
    End If
    
    Call Form_Resize
End Sub

Private Function chkDept(lngDept As Long) As Boolean
    '检验执行科室是否在当前可操作的科室内
    
    Dim cboCtrol As CommandBarComboBox              '科室
    Dim intLoop As Integer
    
    Set cboCtrol = Me.cbrthis.FindControl(, conMenu_View_Busy, True, True)
    
    For intLoop = 1 To cboCtrol.ListCount
        If cboCtrol.ItemData(intLoop) = lngDept Then
            chkDept = True
            Exit Function
        End If
    Next

End Function
Private Sub InitRecordSet(rsNumber As ADODB.Recordset)
    '初始化记录集
    
    '计录试管编码
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "类别", adVarChar, 20
    rsNumber.Fields.Append "管码", adVarChar, 18
    rsNumber.Fields.Append "相关ID", adBigInt
    rsNumber.Fields.Append "样本条码", adVarChar, 18
    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
    rsNumber.Fields.Append "婴儿", adBigInt
    rsNumber.Fields.Append "紧急标志", adBigInt
    rsNumber.Fields.Append "标本", adVarChar, 30
    rsNumber.Fields.Append "医嘱内容", adVarChar, 500
    rsNumber.Fields.Append "采集方式", adVarChar, 100
    rsNumber.Fields.Append "开嘱医生", adVarChar, 100
    rsNumber.Fields.Append "开嘱时间", adDate
    rsNumber.Fields.Append "采样人", adVarChar, 50
    rsNumber.Fields.Append "采样时间", adDate
    rsNumber.Fields.Append "采血量", adVarChar, 20
    rsNumber.Fields.Append "试管名称", adVarChar, 50
    rsNumber.Fields.Append "病人来源", adInteger
    rsNumber.Fields.Append "医嘱ID串", adVarChar, 500
    rsNumber.Fields.Append "执行科室", adVarChar, 200
    rsNumber.Fields.Append "婴儿姓名", adVarChar, 50
    rsNumber.Fields.Append "婴儿性别", adVarChar, 50
    rsNumber.Fields.Append "申请科室", adVarChar, 200
    rsNumber.Fields.Append "所在科室", adVarChar, 200
    
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
End Sub

Public Function MakeBarCode(rsNumber As ADODB.Recordset, RowRecord As ReportRecord, intMode As Integer, Optional intExecDept As Integer, Optional strBarCode As String) As Boolean
'功能                   生成条码并记录方便后面保存到数据或打印
'参数                   用于记录的记录庥
'                       RowRecord行数据
'                       '执行科室是否要区分
'                       Mode =0 绑定条码 =1 生成条码 =2 清除条码 = 3 完成采集 = 4 打印条码或回执单
'                       strBarCode <> ""时表示使用绑定条码
    Dim strFilter As String
    Dim blnNew As Boolean
    Dim str医嘱内容 As String
    Dim strErr As String

    blnNew = False
    Select Case intMode
    Case 0                              '绑定
        If rsNumber.RecordCount = 0 Then blnNew = True
    Case 1                              '生成
        strFilter = "诊疗项目ID=" & RowRecord.Item(mAcol.诊疗项目ID).Value
        rsNumber.filter = strFilter
        If rsNumber.EOF = False Then
            '当诊疗项目相同时新增一个条码
            blnNew = True
        Else
            On Error GoTo GoOn  '兼容LIS低于10.35.130没有IsToleranceItem借口的情况
            If mobjLisInsideComm.IsToleranceItem(Val(RowRecord.Item(mAcol.诊疗项目ID).Value), strErr) Then
                '耐受试验重新生成条码
                If strErr <> "" Then
                    MsgBox strErr, vbInformation, Me.Caption
                End If
                blnNew = True
            End If
GoOn:
            Err.Clear
            On Error GoTo 0
            strFilter = "管码='" & RowRecord.Item(mAcol.试管编码).Value & _
                        "' And 婴儿=" & RowRecord.Item(mAcol.婴儿).Value & _
                      " And 紧急标志=" & IIf(RowRecord.Item(mAcol.紧急).Value = "紧急", 1, 0) & _
                      " And 标本='" & RowRecord.Item(mAcol.标本).Value & "'" & _
                        IIf(Me.chkApplyDept.Value = 1, " and 申请科室='" & RowRecord.Item(mAcol.申请科室).Value & "'", "")
            If intExecDept = 1 Then strFilter = strFilter & " And 执行科室id=" & RowRecord.Item(mAcol.检验执行科室ID).Value
            rsNumber.filter = strFilter
            If rsNumber.EOF = True Then
                '生成新条码
                blnNew = True
            End If
        End If
    Case 2                              '取消条码
        If rsNumber.RecordCount = 0 Then blnNew = True

    Case 3, 4                           '用于条码打印
        strFilter = "样本条码='" & RowRecord.Item(mAcol.条码).Value & "'"
        rsNumber.filter = strFilter
        If rsNumber.EOF = True Then
            blnNew = True
        End If
    End Select
    If blnNew = True Then
        rsNumber.AddNew
        '绑定和生成条码
        rsNumber!类别 = RowRecord.Item(mAcol.类别).Value
        If strBarCode <> "" Then
            rsNumber!样本条码 = strBarCode
        Else
            If intMode = 3 Or intMode = 4 Then
                rsNumber!样本条码 = RowRecord.Item(mAcol.条码).Value
            Else
                rsNumber!样本条码 = zlDatabase.GetNextNo(125, Split(RowRecord.Item(mAcol.ID).Value, ",")(0))
            End If
        End If
        rsNumber!申请科室 = RowRecord.Item(mAcol.申请科室).Value
        rsNumber!采集方式 = RowRecord.Item(mAcol.采集方式).Value
        rsNumber!标本 = RowRecord.Item(mAcol.标本).Value
        rsNumber!执行科室ID = RowRecord.Item(mAcol.检验执行科室ID).Value
        rsNumber!开嘱医生 = RowRecord.Item(mAcol.开嘱医生).Value
        rsNumber!开嘱时间 = RowRecord.Item(mAcol.开嘱时间).Value
        rsNumber!采样人 = RowRecord.Item(mAcol.采样人).Value
        If RowRecord.Item(mAcol.采样时间).Value <> "" Then
            rsNumber!采样时间 = RowRecord.Item(mAcol.采样时间).Value
        End If
        rsNumber!管码 = RowRecord.Item(mAcol.试管编码).Value
        rsNumber!采血量 = RowRecord.Item(mAcol.采血量).Value
        rsNumber!试管名称 = RowRecord.Item(mAcol.试管名称).Value
        rsNumber!紧急标志 = IIf(RowRecord.Item(mAcol.紧急).Value = "紧急", 1, 0)
        rsNumber!病人来源 = RowRecord.Item(mAcol.病人来源).Value
        rsNumber!婴儿 = RowRecord.Item(mAcol.婴儿).Value
        rsNumber!执行科室 = RowRecord.Item(mAcol.执行科室).Value
        rsNumber!医嘱内容 = RowRecord.Item(mAcol.别名).Value
        rsNumber!诊疗项目ID = RowRecord.Item(mAcol.诊疗项目ID).Value
        rsNumber!婴儿姓名 = RowRecord.Item(mAcol.婴儿姓名).Value
        rsNumber!婴儿性别 = RowRecord.Item(mAcol.婴儿性别).Value
        rsNumber!所在科室 = RowRecord.Item(mAcol.病人所在科室).Value
        rsNumber!医嘱ID串 = Replace(Replace(RowRecord.Item(mAcol.ID).Value & "," & RowRecord.Item(mAcol.合并医嘱).Value, ";", ","), ",,", ",")
        If Left(rsNumber!医嘱ID串, 1) = "," Then rsNumber!医嘱ID串 = Mid(rsNumber!医嘱ID串, 2)
        If Right(rsNumber!医嘱ID串, 1) = "," Then rsNumber!医嘱ID串 = Mid(rsNumber!医嘱ID串, 1, Len(rsNumber!医嘱ID串) - 1)
        rsNumber.Update
    Else
        If rsNumber.RecordCount > 0 Then
            rsNumber.MoveLast
            str医嘱内容 = IIf(Trim(RowRecord.Item(mAcol.别名).Value) = "", RowRecord.Item(mAcol.医嘱内容).Value, RowRecord.Item(mAcol.别名).Value)
            If InStr(";" & rsNumber!医嘱内容 & ";", ";" & str医嘱内容 & ";") <= 0 Then
                rsNumber!医嘱内容 = rsNumber!医嘱内容 & ";" & str医嘱内容
            End If

            rsNumber!医嘱ID串 = rsNumber!医嘱ID串 & "," & Replace(Replace(RowRecord.Item(mAcol.ID).Value & RowRecord.Item(mAcol.合并医嘱).Value, ";", ","), ",,", ",")
            If Left(rsNumber!医嘱ID串, 1) = "," Then rsNumber!医嘱ID串 = Mid(rsNumber!医嘱ID串, 2)
            If Right(rsNumber!医嘱ID串, 1) = "," Then rsNumber!医嘱ID串 = Mid(rsNumber!医嘱ID串, 1, Len(rsNumber!医嘱ID串) - 1)
            rsNumber.Update
        End If
    End If
    rsNumber.filter = ""
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    
    If Not txtGoto.Locked And txtGoto.Text = "" And Me.ActiveControl Is txtGoto And strNO <> "" Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C3IC卡号
        txtGoto.Text = strNO
        txtGoto.Tag = strNO
        Call txtGoto_KeyPress(vbKeyReturn)
        If txtGoto.Text = "" Then Call mobjICCard.SetEnabled(False)
       
        IDKind.IDKind = lngPreIDKind
    ElseIf Me.ActiveControl Is cmdNewBarcode And txtGoto.Tag = strNO Then
        Call cmdNewBarcode_Click
        txtGoto.Tag = ""
        Call mobjICCard.SetEnabled(False)
    End If
End Sub

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
Private Sub FilterPatient()
    '功能过滤病人列表
    Dim intLoop As Integer
    
    For intLoop = 0 To Me.rptPlist.Rows.Count - 1
        '未绑定
        If Me.optFilter(1).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.未绑定).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '已绑定
        If Me.optFilter(2).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.已绑定).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '已采样
        If Me.optFilter(3).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.已采样).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '已送检
        If Me.optFilter(4).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.已送检).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '已执行
        If Me.optFilter(5).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.已执行).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '拒收
        If Me.optFilter(6).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.拒收).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
    Next
    Me.rptPlist.Populate
End Sub

Private Function CheckMoeny() As Boolean
    '功能           检查是否收费如果收费打开收费确认窗体
    Dim strAdvice As String
    Dim lngLoop As Long
    Dim intProperties As Integer
    Dim intPatientType As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '根据参数来判断是否检查
    If mblnNowConsumption = False Then: CheckMoeny = True: Exit Function
    If mobjSquareCard Is Nothing Then Exit Function
    
    With Me.rptAlist(Me.TabCtr.Selected.Index)
        For lngLoop = 0 To .Records.Count - 1
            If .Records(lngLoop).Item(mAcol.选择).Checked = True Then
                If .Records(lngLoop).Visible = True And .Records(lngLoop).Item(mAcol.计费状态).Value <> 3 Then
                    strAdvice = strAdvice & "," & .Records(lngLoop).Item(mAcol.ID).Value & "," & .Records(lngLoop).Item(mAcol.相关ID).Value & "," & _
                        .Records(lngLoop).Item(mAcol.合并医嘱).Value
                    intProperties = Val(.Records(lngLoop).Item(mAcol.记录性质).Value)
                    intPatientType = Val(.Records(lngLoop).Item(mAcol.病人来源).Value)
                End If
            End If
        Next
        '只有门诊病人才处理
        If intPatientType = 1 Then
            If strAdvice <> "" Then strAdvice = Mid(strAdvice, 2)
            If mobjSquareCard.zlSquareAffirm(Me, glngModul, "", mlngKey, 0, False, , , strAdvice) = False Then
                Exit Function
            End If
        Else
            '处理住院发送到门诊收费的情况
            strSQL = "Select Count(ID) Count" & vbNewLine & _
                    "From 门诊费用记录" & vbNewLine & _
                    "Where 病人id = [1] And 医嘱序号 In (Select * From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))) And 记录状态 = 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "费用确认", mlngKey, strAdvice)
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp("Count") & "") > 0 Then
                    If mobjSquareCard.zlSquareAffirm(Me, glngModul, "", mlngKey, 0, False, , , strAdvice) = False Then
                        Exit Function
                    End If
                End If
            End If

        End If
    End With
    CheckMoeny = True
End Function

Private Function ReplaseSpecial(strTmp As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能               替换特殊字符
    '参数
    '                   需替换的字符
    '返回               需替换了特殊字符后的字串
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strSpecial As String
    Dim astrTmp() As String
    strSpecial = "'^‘^’^;^；^:^：^?^？^|^,^，^.^。^"""
    astrTmp = Split(strSpecial, "^")
    For intLoop = 0 To UBound(astrTmp)
        strTmp = Replace$(strTmp, astrTmp(intLoop), "")
    Next
    ReplaseSpecial = strTmp
    
End Function

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
    
    If Me.rptPlist.FocusedRow Is Nothing Then
        getTATTime = False
        Exit Function
    End If
    
    '获取病人性别和申请科室
    With Me.rptPlist.FocusedRow
        strSex = .Record(mPcol.性别).Value
        strDept = .Record(mPcol.病人科室).Value
    End With
    
    '获取项目ID,项目名称,采样时间,急诊
    strItem = ""
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.选择).Checked = True Then
'            If Record(mAcol.采样时间).Value <> "" Then
                var_Item = Split(Record(mAcol.医嘱内容).Value, ";")
                For i = LBound(var_Item) To UBound(var_Item)
                    strItem = strItem & ";" & Record(mAcol.诊疗项目ID).Value & "," & var_Item(i) & _
                                            "," & Record(mAcol.采样时间).Value & "," & IIf(Record(mAcol.紧急).Value = "紧急", 1, 0) & _
                                             "," & Record(mAcol.ID).Value & "," & Record(mAcol.条码).Value
                Next
'            Else
'                strMsgNoTime = strMsgNoTime & Record(mAcol.医嘱内容).Value & vbCrLf
'            End If
        End If
    Next
    
    If strMsgNoTime <> "" Then MsgBox strMsgNoTime & "未采样,不能送检", vbInformation, Me.Caption
    If strItem <> "" Then
        strItem = Mid(strItem, 2)
    Else
        getTATTime = False
        strIDs = ""
        Exit Function
    End If
    
    '检查TAT是否超时
    On Error GoTo errold
    strTATItems = mobjLisInsideComm.GetTatTimeShow(1, strItem, strDept, "", "", strSex, intMsg, strShowBef, , UserInfo.姓名)
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
        
        '获取所有项目的条码
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
                strMsgShowStop = strMsgShowStop & Split(var_Tmp(i), ",")(1) & "未采样,不能送检" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 And Split(var_Tmp(i), ",")(2) <> "" Then
                '超时并禁止的
                strMsgShowStop = strMsgShowStop & Replace(Replace(Split(var_Tmp(i), ",")(8), "[项目]", Split(var_Tmp(i), ",")(1)), "[超时]", Split(var_Tmp(i), ",")(7) & "分钟") & vbCrLf
            Else
                '不同项目同条码的时候,当有一个项目超时,则所有该条码的项目均不能送检
                If InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 Then
                    strIDs = strIDs & "," & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
                End If
            End If
        Next
        
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
        End If
        
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
                        strIDs = strIDs & "," & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
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



