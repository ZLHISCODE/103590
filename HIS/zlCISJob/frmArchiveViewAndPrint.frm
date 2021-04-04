VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmArchiveViewAndPrint 
   BackColor       =   &H80000005&
   Caption         =   "病案查询打印"
   ClientHeight    =   11460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16485
   Icon            =   "frmArchiveViewAndPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11460
   ScaleWidth      =   16485
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ProgressBar psb 
      Height          =   225
      Left            =   10680
      TabIndex        =   70
      Top             =   11160
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   480
      ScaleHeight     =   8055
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   1080
      Width           =   3975
      Begin VB.CheckBox chkFilter 
         Height          =   255
         Left            =   3360
         Picture         =   "frmArchiveViewAndPrint.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "按照查找条件对病人进行过滤显示"
         Top             =   1133
         Width           =   270
      End
      Begin VB.Frame fraType 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4815
         Begin VB.OptionButton optType 
            Caption         =   "住院"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   22
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optType 
            Caption         =   "门诊"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   21
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblType 
            BackStyle       =   0  'Transparent
            Caption         =   "类型"
            Height          =   180
            Left            =   60
            TabIndex        =   23
            Top             =   150
            Width           =   450
         End
      End
      Begin VB.ComboBox cboDept 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   19
         Text            =   "cboDept"
         Top             =   600
         Width           =   2655
      End
      Begin VB.PictureBox picPatiIn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   4005
         TabIndex        =   10
         Top             =   3840
         Width           =   4005
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   1620
            Left            =   480
            TabIndex        =   11
            Top             =   1800
            Width           =   2280
            _Version        =   589884
            _ExtentX        =   4022
            _ExtentY        =   2857
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   1
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   3855
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   1
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   0
               Width           =   1230
            End
            Begin VB.Label lblSeeTim 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "就诊时间"
               Height          =   180
               Left            =   120
               TabIndex        =   26
               Top             =   45
               Width           =   720
            End
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   0
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   3855
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   0
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   0
               Width           =   1230
            End
            Begin VB.Label lbl出院时间 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               Height          =   180
               Left            =   0
               TabIndex        =   14
               Top             =   60
               Width           =   720
            End
         End
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   300
         Left            =   960
         TabIndex        =   15
         Top             =   1110
         Width           =   2295
         _ExtentX        =   4683
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmArchiveViewAndPrint.frx":D0A4
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   0
         InputAppearance =   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   555
         FindPatiShowName=   0   'False
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
      End
      Begin XtremeSuiteControls.TabControl tbcPati 
         Height          =   1335
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   2355
         _StockProps     =   64
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室(&D)↓"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   660
         Width           =   810
      End
      Begin VB.Label lblFind 
         BackStyle       =   0  'Transparent
         Caption         =   "查找(F3)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   735
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   4800
      ScaleHeight     =   7935
      ScaleWidth      =   4935
      TabIndex        =   6
      Top             =   2640
      Width           =   4935
      Begin VB.ComboBox cboVisit 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   0
         Width           =   3765
      End
      Begin MSComctlLib.TreeView tvwArchive 
         Height          =   1785
         Left            =   960
         TabIndex        =   8
         Top             =   3600
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   3149
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ils16"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.TabControl tbcHistory 
         Height          =   2895
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3255
         _Version        =   589884
         _ExtentX        =   5741
         _ExtentY        =   5106
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11790
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   2670
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picRpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   11625
      ScaleHeight     =   780
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   3255
      Width           =   915
      Begin SHDocVwCtl.WebBrowser webRpt 
         Height          =   450
         Left            =   135
         TabIndex        =   3
         Top             =   150
         Width           =   450
         ExtentX         =   794
         ExtentY         =   794
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame fraPati 
      BackColor       =   &H80000005&
      Caption         =   "病人信息"
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   11565
      Begin VB.Frame fraOutPati 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1560
         TabIndex        =   48
         Top             =   360
         Width           =   9375
         Begin VB.Label lbl急 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "急"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   8880
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "姓名:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   2580
            TabIndex        =   66
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "性别:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   4860
            TabIndex        =   65
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "门诊号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   13
            Left            =   180
            TabIndex        =   64
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "年龄:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   14
            Left            =   7140
            TabIndex        =   63
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "身份证号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   15
            Left            =   2220
            TabIndex        =   62
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "就诊日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   17
            Left            =   4500
            TabIndex        =   61
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "降央卓玛"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   3060
            TabIndex        =   60
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "28岁"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   7620
            TabIndex        =   59
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "女"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   13
            Left            =   5340
            TabIndex        =   58
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "20150101"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   14
            Left            =   825
            TabIndex        =   57
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "500101198810121245"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   15
            Left            =   3060
            TabIndex        =   56
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   17
            Left            =   5340
            TabIndex        =   55
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "地址:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   18
            Left            =   4860
            TabIndex        =   54
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "出生日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   19
            Left            =   0
            TabIndex        =   53
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "门诊医师:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   0
            TabIndex        =   52
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "重庆市两江新区"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   16
            Left            =   5340
            TabIndex        =   51
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "赵丽颖"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   825
            TabIndex        =   50
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   18
            Left            =   825
            TabIndex        =   49
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fraInPati 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   9375
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "姓名:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   2580
            TabIndex        =   47
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "性别:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   4860
            TabIndex        =   46
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "住院号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   45
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "年龄:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   7140
            TabIndex        =   44
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "身份证号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   2220
            TabIndex        =   43
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "病况:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   2580
            TabIndex        =   42
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "入院:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   4860
            TabIndex        =   41
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "降央卓玛"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   3060
            TabIndex        =   40
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "28岁"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   7620
            TabIndex        =   39
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "女"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   5340
            TabIndex        =   38
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "20150101"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   825
            TabIndex        =   37
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   3060
            TabIndex        =   36
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   3060
            TabIndex        =   35
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   5340
            TabIndex        =   34
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "地址:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   4860
            TabIndex        =   33
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "出生日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   0
            TabIndex        =   32
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "住院医师:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   0
            TabIndex        =   31
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "重庆市两江新区"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   5340
            TabIndex        =   30
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "赵丽颖"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   825
            TabIndex        =   29
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   825
            TabIndex        =   28
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Image imgPatient 
         Height          =   1185
         Left            =   120
         Picture         =   "frmArchiveViewAndPrint.frx":D187
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   11100
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16060
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
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
   Begin XtremeSuiteControls.TabControl tbcArchive 
      Height          =   2595
      Left            =   10920
      TabIndex        =   5
      Top             =   3720
      Width           =   4125
      _Version        =   589884
      _ExtentX        =   7276
      _ExtentY        =   4577
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":E051
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":148B3
            Key             =   "门诊"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":14E4D
            Key             =   "home"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1B6AF
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1BC49
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1C1E3
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1C77D
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1CD17
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1D2B1
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1D84B
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":240AD
            Key             =   "Path"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   3000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":24647
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":2AEA9
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":3170B
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":31CA5
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":3223F
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":38AA1
            Key             =   "首页正面"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":3BADB
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":4233D
            Key             =   "检查报告"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":47DFF
            Key             =   "检验报告"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":4D8C1
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":54123
            Key             =   "Patient"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":54FFD
            Key             =   "unCheckAll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":55597
            Key             =   "CheckAll"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":55B31
            Key             =   "住院病历"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":58B6B
            Key             =   "其他报表"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":5BBA5
            Key             =   "疾病证明"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":5EBDF
            Key             =   "首页附页一"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":5FAB9
            Key             =   "临床路径"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":62AF3
            Key             =   "首页附页二"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":65B2D
            Key             =   "护理病历"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":68B67
            Key             =   "住院医嘱"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":6BBA1
            Key             =   "护理记录"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":6EBDB
            Key             =   "知情文件"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":71C15
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":7224F
            Key             =   "CheckFill"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":72889
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":72EC3
            Key             =   "首页反面"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":75EFD
            Key             =   "down"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":7C75F
            Key             =   "up"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":82FC1
            Key             =   "住院证"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmArchiveViewAndPrint.frx":86453
      Left            =   600
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmArchiveViewAndPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API声明
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Enum INPATIREPORT_COLUMN
    col_选择 = 0
    col_图标 = 1
    col_打印图标 = 2
    col_是否编目 = 3
    col_编目日期 = 4
    col_住院号 = 5
    COL_NO = 6             '门诊
    COL_门诊号 = 7          '门诊
    col_床号 = 8
    col_姓名 = 9
    col_性别 = 10
    col_年龄 = 11
    col_身份证号 = 12
    col_出生日期 = 13
    COL_执行时间 = 14        '门诊
    col_入院日期 = 15
    col_出院日期 = 16
    col_住院医师 = 17       '门诊就诊医生
    col_家庭地址 = 18
    col_就诊卡号 = 19
    col_留观号 = 20
    '隐藏列
    col_病人类型 = 21
    col_病人Id = col_病人类型 + 1        '隐藏
    col_主页ID = col_病人类型 + 2         '隐藏 门诊为挂号ID
    col_科室ID = col_病人类型 + 3       '隐藏
    COL_数据转出 = col_病人类型 + 4     '数据转出
    col_打印记录 = col_病人类型 + 5
End Enum

Private Enum PATI_INFO
    lbl_姓名 = 0
    lbl_性别 = 1
    lbl_年龄 = 2
    lbl_身份证号 = 3
    lbl_住院号 = 4
    lbl_病况 = 5
    lbl_入院日期 = 6
    lbl_家庭地址 = 8
    lbl_住院医师 = 9
    lbl_出生日期 = 10
    
    lblOUT_姓名 = 11
    lblOUT_年龄 = 12
    lblOUT_性别 = 13
    lblOUT_门诊号 = 14
    lblOUT_身份证号 = 15
    lblOUT_家庭地址 = 16
    lblOUT_就诊日期 = 17
    lblOUT_出生日期 = 18
    lblOUT_门诊医师 = 7
End Enum

Private Type PatiInfo
    状态 As Integer '病案主页.状态
    婴儿 As Integer
    住院号 As String
    床号 As String
    主页ID As Long
    病人ID As Long
    病区ID As Long
    科室ID As Long
    入院日期 As Date
    出院日期 As Date
    编目日期 As Date
    住院次数 As Long
    数据转出 As Boolean
End Type

Private Enum E_TYPE
    E_住院 = 0
    E_门诊 = 1
End Enum
'常量
Private Const M_CON_CATE As String = "R11,R12,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10"
'门诊卡结算对象返回的可用的医疗卡
Private Const mstrCardKindOut         As String = "就|就诊卡|0|0|8|0|0|0;门|门诊号|0|0|0|0|0|0;挂|挂号单|0|0|0|0|0|0;姓|姓名|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0;ＩＣ|ＩＣ卡|1|0|0|0|0|0"
'住院卡结算对象返回的可用的医疗卡
Private Const mstrCardKindIN          As String = "就|就诊卡|0|0|8|0|0|0;住|住院号|0|0|0|0|0|0;床|床号|0|0|0|0|0|0;姓|姓名|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0;留|留观号|0|0|0|0|0|0"
'直接查找卡结算对象返回的可用的医疗卡
Private Const mstrCardKindFind        As String = "就|就诊卡|0|0|8|0|0|0;门|门诊号|0|0|0|0|0|0;住|住院号|0|0|0|0|0|0;单|单据号|0|0|0|0|0|0;姓|姓名|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0;ＩＣ|ＩＣ卡|1|0|0|0|0|0;医|医保号|0|0|0|0|0|0"

'事件
Private WithEvents mclsTendsNew     As zl9TendFile.clsTendFile    '新版护士工作站
Attribute mclsTendsNew.VB_VarHelpID = -1
Private WithEvents mobjReport       As zl9Report.clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mclsDockAduits   As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
'
Private mclsOutAdvices          As zlPublicAdvice.clsDockOutAdvices
Private mclsInAdvices           As zlPublicAdvice.clsDockInAdvices
Private mclsPath                As zlPublicPath.clsDockPath
Private mclsArchive             As zlMedRecPage.clsArchive '电子病案查阅窗体类
Private mobjRichEMR             As Object
Private mobjPublicPACS          As Object
Private mobjSquareCard          As Object      '卡结算对象
Private mobjReportForm          As Object      '报表预览对象
Private mobjPatient             As Object
Private mobjInfection           As Object      '传染病报告卡,模拟预览功能
'
Private mstr挂号单              As String
Private mstrPrivs               As String

Private mstr检验报告打印        As String        '0-老版LIS报表或病历;1-新版LIS报表方式
Private mstr检验对应报表        As String
Private mstr检查对应报表        As String
Private mstrFindType           As String       '用来存储当前查找类型的名称
Private mstrPrintDocIDs        As String       '共享病历的子文档只打印一次
Private mstrTempDel            As String        '删除临时文件
Private mstrPrintMedRec        As String        '记录已经打印病案

'PDF打印
'Public gstrInputSeverName As String
'Public gstrInputUser As String
'Public gstrInputPwd As String
'
''
Private mlng病人ID      As Long
Private mlng就诊ID      As Long '病人当前或者最后的就诊ID，门诊为挂号ID,住院号主页ID
Private mlng科室ID      As Long
Private mlng病区ID      As Long
Private mlngPreDept     As Long

Private mintPatiCount   As Integer   '勾选病人数目
Private mintPreDept     As Integer
Private mintDeptView    As Integer '0-按科室显示，1-按病区显示
Private mintDeptViewBed As Integer '0，1-只显示有床位的病区或者科室
Private mintMecStandard As Integer  '病案首页格式 0-卫生部标准，1-四川省标准，2-云南省标准,3-湖南省标准
Private mintFindType    As Integer '0-住院号,1-床号,2-就诊卡,3-姓名
Private mintOutPreTime  As Integer

'
Private mbytType        As Byte             '0-住院;1-门诊
Private mbytPDFStatu    As Byte             '0-未初始化;1-初始化成功
Private mbytPrintType   As Byte             '1-打印首页
'
Private mblnLIS         As Boolean         '是否按照新版LIS
Private mblnOutDept     As Boolean '是否仅服务于门诊的科室（门诊留观病人显示门诊号）
Private mblnMoved       As Boolean '当前病人数据是否转出
Private mblnNewTends    As Boolean 'T-新版护理记录;F-老版护理记录
Private mblnICU         As Boolean '是否非本科的ICU室
Private mblnUndo        As Boolean
Private mblnTabTmp      As Boolean
Private mblnTvwTmp      As Boolean
Private mblnSeePic      As Boolean           'T-显示观片;F-隐藏观片
Private mblnPrint        As Boolean           'T-允许记录病案打印
'
Private mcolSubForm     As Collection
Private mcolReport      As Collection
Private mcolPrint       As Collection       '缓存打印机
'


Private mdatOutBegin As Date, mdatOutEnd As Date    '出院指定时间
Private mDatBegin As Date, mDatEnd As Date          '已诊指定时间

Private mrsPati         As ADODB.Recordset '病人信息集合，包含同一身份证号的所有病人
Private mrsData         As ADODB.Recordset
Private mrsMedRec       As ADODB.Recordset
Private mblnReturn      As Boolean

Public Sub ShowArchive(ByVal frmParent As Object, ByVal strPrivs As String)
'功能：公共接口方法，类似 ShowMe方法
'参数:frmParent-父窗体
'     strPrivs-模块权限
    mstrPrivs = strPrivs

    Me.Show 0, frmParent  ' 必须为无模式;有模式的情况下 临床路径、诊疗报告、体温单子窗体加载时卡死。
End Sub

Private Sub InitBasicData()
'功能：初始化一些基本数据，如下拉列表加载等
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim strSQL As String
    Dim objTab As TabControlItem
    Dim strTmp As String
    Dim str病人IDs As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    Dim blnTmp As Boolean
    Dim str身份证号 As String
    Dim strOrder As String
    
    Screen.MousePointer = 11
    LockWindowUpdate Me.hwnd
    mstr挂号单 = "": mlngPreDept = -1

    Call tbcHistory.RemoveAll
    Call cboVisit.Clear
    
    If mlng病人ID = 0 Then
        Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, "", tvwArchive.hwnd, 0)
        If mbytType = E_门诊 Then
            Call ShowOutPatiInfo
        Else
            Call ShowInPatiInfo
        End If
        Call ShowArchiveTree
    Else
        On Error GoTo errH
        strSQL = "select a.身份证号 from 病人信息 a where a.病人id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        strTmp = rsTmp!身份证号 & ""
        If strTmp <> "" Then
            '验证身份证号的合法性
            If mobjPatient Is Nothing Then
                On Error Resume Next
                Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                err.Clear: On Error GoTo 0
            End If
            If mobjPatient Is Nothing Then
                MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
            Else
                Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
                If mobjPatient.CheckPatiIdcard(strTmp) Then
                    str身份证号 = strTmp
                End If
            End If
        End If
        
        On Error GoTo errH
        If chkFilter.Value = vbChecked Then
            strOrder = "开始时间 Desc"
        Else
            If mbytType = E_住院 Then
                strOrder = "类型 Desc,开始时间 Desc"
            Else
                strOrder = "类型 ASC,开始时间 Desc"
            End If
        End If
        If str身份证号 <> "" Then
            strSQL = "select a.病人id from 病人信息 a where a.病人id<>[1] and a.身份证号=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, str身份证号)
            Do While Not rsTmp.EOF
                str病人IDs = str病人IDs & "," & rsTmp!病人ID
                rsTmp.MoveNext
            Loop
        End If
        If str病人IDs = "" Then
            strSQL = " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出,-1 as 病人性质,null as 就诊号,1 as 类型 From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                " Union ALL" & _
                " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出,-1 as 病人性质,null as 就诊号,1 as 类型 From H病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                " Union ALL" & _
                " Select 病人id,主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出,NVL(病人性质,0) as 病人性质,null as 就诊号,2 as 类型 From 病案主页 Where 病人ID=[1] And Nvl(主页ID,0)<>0"
            strSQL = "Select Rownum As 序号,a.病人ID,A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 ,A.病人性质,a.就诊号 From (" & strSQL & ") A,部门表 B Where A.科室ID=B.ID Order by " & strOrder
            Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        Else
            str病人IDs = mlng病人ID & str病人IDs
            strTmp = " 病人ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X) "
            strSQL = " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出,-1 as 病人性质,null as 就诊号,1 as 类型 From 病人挂号记录 Where " & strTmp & " And 记录性质=1 And 记录状态=1 and NO is not null" & _
                " Union ALL" & _
                " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出,-1 as 病人性质,null as 就诊号,1 as 类型 From H病人挂号记录 Where " & strTmp & " And 记录性质=1 And 记录状态=1 and NO is not null" & _
                " Union ALL" & _
                " Select 病人id,主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出,NVL(病人性质,0) as 病人性质,住院号 as 就诊号,2 as 类型 From 病案主页 Where " & strTmp & " And Nvl(主页ID,0)<>0"
            strSQL = "Select Rownum As 序号,a.病人ID,A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 ,A.病人性质,a.就诊号 From (" & strSQL & ") A,部门表 B Where A.科室ID=B.ID  Order by " & strOrder
            Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str病人IDs)
        End If
        Do While Not mrsData.EOF
            strTmp = IIf(IsNull(mrsData!NO), "第" & mrsData!就诊id & "次" & IIf(mrsData!病人性质 = 1, "门诊留观", IIf(mrsData!病人性质 = 2, "住院留观", "住院")), "门诊就诊") & ":" & mrsData!科室 & "," & Format(mrsData!开始时间, "yyyy-MM-dd HH:mm") & _
            IIf(Not IsNull(mrsData!结束时间), "～" & Format(mrsData!结束时间, "yyyy-MM-dd HH:mm"), "")
        
            If mrsData.AbsolutePosition = 1 Then
                Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, strTmp, tvwArchive.hwnd, IIf(IsNull(mrsData!NO), 0, 1))
            End If
             
            cboVisit.AddItem strTmp
            cboVisit.ItemData(cboVisit.NewIndex) = Val(mrsData!序号)
            mrsData.MoveNext
        Loop
        If cboVisit.ListCount > 0 Then
            Call Cbo.SetIndex(cboVisit.hwnd, 0)
            Call cboVisit_Click
        End If
    End If
    LockWindowUpdate 0
    Screen.MousePointer = 0
    Exit Sub
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim arrTmp  As Variant
    Dim strFunc As String
    Dim strTmp  As String
    Dim i As Long
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Dept, "部门显示(&D)") '固有
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 1, "按科室显示(&D)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 2, "按病区显示(&U)", -1, False)
        End With
    End With
        
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_FilePopup, "文件")
        objPopup.IconId = conMenu_File_Open
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet * 10# + 1, "打印设置")
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet * 10# + 2, "参数设置")
                objControl.IconId = conMenu_File_Parameter
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet * 10# + 3, "PDF位置")
                objControl.ToolTipText = "设置PDF输出位置"
                objControl.IconId = conMenu_File_PDF
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_PDF, "PDF")
        Set objControl = .Add(xtpControlButton, conMenu_Img_Look, "观片")
        objControl.IconId = conMenu_Edit_MarkMap: objControl.BeginGroup = True
        '扩展功能
        If CreatePlugInOK(P病案查阅打印) Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Tool_PlugIn, "扩展功能", objControl.Index + 1)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                On Error Resume Next
                strFunc = gobjPlugIn.GetFuncNames(glngSys, P病案查阅打印)
                Call zlPlugInErrH(err, "GetFuncNames")
                err.Clear: On Error GoTo 0
                If strFunc <> "" Then
                    arrTmp = Split(strFunc, ",")
                    strTmp = Replace(strFunc, "Auto:", "")
                    arrTmp = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp)
                        Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = arrTmp(i)
                    Next
                Else
                    objPopup.Visible = False
                End If
            End With
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
    End With
    
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub InitReportColumn()
'参数:bytFunc=0 住院;bytFunc=1 门诊
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptPati
        .Columns.DeleteAll
        
        Set objCol = .Columns.Add(col_选择, "", 20, False)
            objCol.Icon = imgPati.ListImages("UnCheck").Index - 1
            objCol.EditOptions.AllowEdit = True
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_图标, "", 20, False)  '图标
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_打印图标, "", 20, False)  'col_打印图标
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_是否编目, "是否编目", 60, True)
        Set objCol = .Columns.Add(col_编目日期, "编目日期", 80, True)
        Set objCol = .Columns.Add(col_住院号, "住院号", 80, True)
        Set objCol = .Columns.Add(COL_NO, "NO", 80, True)
        Set objCol = .Columns.Add(COL_门诊号, "门诊号", 80, True)
        Set objCol = .Columns.Add(col_床号, "床号", 50, True)
        Set objCol = .Columns.Add(col_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(col_性别, "性别", 45, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 45, True)
        Set objCol = .Columns.Add(col_身份证号, "身份证号", 150, True)
        Set objCol = .Columns.Add(col_出生日期, "出生日期", 80, True)
        Set objCol = .Columns.Add(COL_执行时间, "执行时间", 80, True)
        Set objCol = .Columns.Add(col_入院日期, "入院日期", 80, True)
        Set objCol = .Columns.Add(col_出院日期, "出院日期", 80, True)
        Set objCol = .Columns.Add(col_住院医师, "住院医师", 80, True)
        Set objCol = .Columns.Add(col_家庭地址, "地址", 150, True)
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_就诊卡号, "就诊卡号", 0, False)
        Else
            Set objCol = .Columns.Add(col_就诊卡号, "就诊卡号", 70, True)
        End If
        Set objCol = .Columns.Add(col_留观号, "留观号", 62, True)
        '隐藏列
        Set objCol = .Columns.Add(col_病人类型, "病人类型", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_科室ID, "科室ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_数据转出, "数据转出", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_打印记录, "打印记录", 0, False): objCol.Visible = False
        Call ShowReportColumn
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        mblnUndo = True
        .MultipleSelection = False '会引发SelectionChanged事件
         mblnUndo = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
        If mbytType = 0 Then
            .SortOrder.Add .Columns(col_入院日期)
            .SortOrder(0).SortAscending = True
        Else
            .SortOrder.Add .Columns(COL_执行时间)
            .SortOrder(0).SortAscending = True
        End If
    End With
End Sub

Private Sub ShowReportColumn()
    With rptPati.Columns
    If tbcPati.Selected.Tag = "出院" Then
        .Find(col_是否编目).Visible = True
        .Find(col_编目日期).Visible = True
        .Find(col_打印图标).Visible = True
        .Find(col_住院号).Visible = True
        .Find(col_入院日期).Visible = True
        .Find(col_出院日期).Visible = True
        .Find(col_床号).Visible = True
        .Find(col_留观号).Visible = True
        
        .Find(COL_NO).Visible = False
        .Find(COL_门诊号).Visible = False
        .Find(COL_执行时间).Visible = False
        .Find(col_住院医师).Caption = "住院医师"
    ElseIf tbcPati.Selected.Tag = "在院" Then
        .Find(col_是否编目).Visible = False
        .Find(col_编目日期).Visible = False
        .Find(col_打印图标).Visible = False
        .Find(col_住院号).Visible = True
        .Find(col_入院日期).Visible = True
        .Find(col_出院日期).Visible = True
        .Find(col_床号).Visible = True
        .Find(col_留观号).Visible = True
        
        .Find(COL_NO).Visible = False
        .Find(COL_门诊号).Visible = False
        .Find(COL_执行时间).Visible = False
        .Find(col_住院医师).Caption = "住院医师"
    ElseIf tbcPati.Selected.Tag = "在诊" Then
        .Find(col_是否编目).Visible = False
        .Find(col_编目日期).Visible = False
        .Find(col_打印图标).Visible = False
        .Find(col_住院号).Visible = False
        .Find(col_入院日期).Visible = False
        .Find(col_出院日期).Visible = False
        .Find(col_床号).Visible = False
        .Find(col_留观号).Visible = False
         
        .Find(COL_NO).Visible = True
        .Find(COL_门诊号).Visible = True
        .Find(COL_执行时间).Visible = True
        .Find(col_住院医师).Caption = "接诊医生"
    ElseIf tbcPati.Selected.Tag = "已诊" Then
        .Find(col_是否编目).Visible = False
        .Find(col_编目日期).Visible = False
        .Find(col_打印图标).Visible = False
        .Find(col_住院号).Visible = False
        .Find(col_入院日期).Visible = False
        .Find(col_出院日期).Visible = False
        .Find(col_床号).Visible = False
        .Find(col_留观号).Visible = False
        
        .Find(COL_NO).Visible = True
        .Find(COL_门诊号).Visible = True
        .Find(COL_执行时间).Visible = True
        .Find(col_住院医师).Caption = "接诊医生"
    End If
    End With
End Sub

Private Sub cboDept_Click()
'功能：刷新界面数据
'说明：从该事件开始会不重复引发相关的数据读取
    Dim lng部门ID As Long, i As Long, lngidx As Long
    Dim blnIn病区 As Boolean, rsTmp As Recordset, str科室IDs As String
    
    If cboDept.ListIndex = -1 Then
        Call ClearPatiInfo
        Exit Sub
    End If
    cboDept.Tag = cboDept.ListIndex
    mintPreDept = cboDept.ListIndex

    '重新读取病人
    Call LoadPatients
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strDeptIDs As String
    
    mblnReturn = False
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        If cboDept.Text <> "" Then
            Set rsTmp = GetDataToDepts(cboDept.Text)
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cboDept, rsTmp!ID)
            Else
                cboDept.ListIndex = Val(cboDept.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
        End If
    End If
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    End If
End Sub

Private Sub cboSelectTime_Click(Index As Integer)
'功能:Index 0-出院 ;1-已诊
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    If Index = 0 Then
        intDateCount = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
        datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
        If cboSelectTime(Index).ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdatOutBegin, mdatOutEnd, cboSelectTime(Index)) Then
                '取消时恢复原来的选择
                Call Cbo.SetIndex(cboSelectTime(Index).hwnd, mintOutPreTime)
                Exit Sub
            End If
        Else
            mdatOutEnd = datCurr
            mdatOutBegin = mdatOutEnd - intDateCount
        End If
        If mdatOutBegin = CDate(0) Or mdatOutEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "范围：" & Format(mdatOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdatOutEnd, "yyyy-MM-dd")
        End If
        mintOutPreTime = cboSelectTime(Index).ListIndex
    Else
        intDateCount = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
        datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        If Me.Visible Then
            If intDateCount = -1 Then
                If Not frmSelectTime.ShowMe(Me, mDatBegin, mDatEnd, cboSelectTime(Index)) Then
                    '取消时恢复原来的选择
                    Call Cbo.SetIndex(cboSelectTime(Index).hwnd, mintOutPreTime)
                    Exit Sub
                End If
            ElseIf intDateCount = 0 Then
                '今天  86114
                mDatBegin = Format(datCurr, "yyyy-MM-dd 00:00:00")
                mDatEnd = Format(datCurr, "yyyy-MM-dd 23:59:59")
            Else
                mDatEnd = Format(datCurr, "yyyy-MM-dd 23:59:59")
                mDatBegin = Format(mDatEnd - intDateCount, "yyyy-MM-dd 00:00:00")
            End If
        End If
        '选择了时间之后，清除挂号单条件
        cboSelectTime(Index).ToolTipText = Format(mDatBegin, "yyyy-MM-dd") & " - " & Format(mDatEnd, "yyyy-MM-dd")
        lblSeeTim.ToolTipText = cboSelectTime(Index).ToolTipText
        mintOutPreTime = cboSelectTime(Index).ListIndex
    End If
    If Me.Visible = True Then Call LoadPatients
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID

    Case conMenu_File_PrintSet * 10# + 1 '打印设置
        frmPrintSet.Show 1
    Case conMenu_File_PrintSet * 10# + 2 '参数设置
         Call SetPrintPara
    Case conMenu_File_PrintSet * 10# + 3
        Call SetPDFPath
    Case conMenu_File_Preview
        Call FuncPrintOrView(1) '预览
    Case conMenu_File_Print
        If Control.Parameter = "DO" Then Exit Sub
        Control.Parameter = "DO"
        Control.Enabled = False
        Call FuncPrintOrView(2) '打印
        Control.Parameter = ""
    Case conMenu_File_PDF
        If Control.Parameter = "DO" Then Exit Sub
        Control.Parameter = "DO"
        Control.Enabled = False
        Call FuncPrintOrView(3) 'PDF
        Control.Parameter = ""
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '按科室/病区显示
        If mintDeptView <> Control.ID - conMenu_View_Dept * 10# - 1 Then
            mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
            Call LoadDept
        End If
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(glngSys, P病案查阅打印, Control.Parameter, mlng病人ID, mlng就诊ID, 0)
            Call zlPlugInErrH(err, "ExecuteFunc")
            err.Clear: On Error GoTo 0
        End If
    Case conMenu_Img_Look
        Call FuncLookPicture
    Case Else
    
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With Me.fraPati
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With
    With Me.tbcArchive
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop + fraPati.Height: .Height = lngBottom - lngTop - fraPati.Height
    End With
    Me.fraInPati.Width = fraPati.Width - imgPatient.Width - 500
    Me.fraOutPati.Width = fraPati.Width - imgPatient.Width - 500
    
    psb.Top = stbThis.Top + Screen.TwipsPerPixelY * 4
    psb.Left = stbThis.Panels(3).Left + Screen.TwipsPerPixelX * 2
    psb.Width = stbThis.Panels(3).Width - Screen.TwipsPerPixelX * 7
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    Case conMenu_FilePopup
        If InStr(";" & mstrPrivs & ";", ";打印;") = 0 And InStr(";" & mstrPrivs & ";", ";参数设置;") = 0 And InStr(";" & mstrPrivs & ";", ";PDF;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_File_PrintSet * 10# + 1, conMenu_File_Print
        If InStr(";" & mstrPrivs & ";", ";打印;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_File_PrintSet * 10# + 2
        If InStr(";" & mstrPrivs & ";", ";参数设置;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_File_PrintSet * 10# + 3, conMenu_File_PDF   'PDF位置  'PDF输出
        If InStr(";" & mstrPrivs & ";", ";PDF;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Img_Look
        Control.Visible = mblnSeePic
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '按科室/病区显示
        Control.Checked = mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
    End Select
    If Control.Visible Then
        Select Case Control.ID
        Case conMenu_File_Print, conMenu_File_PDF, conMenu_File_Preview
            If cboVisit.ListIndex = -1 Or Control.Parameter = "DO" Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End Select
    End If
     
    DeleteLISTempFile
End Sub

Private Sub chkFilter_Click()
    Dim i As Long
    
    PatiIdentify.Text = ""
    If PatiIdentify.Visible And PatiIdentify.Enabled Then PatiIdentify.SetFocus
    If chkFilter.Value = vbChecked Then
        cboDept.Enabled = False
        optType(0).Enabled = False
        optType(1).Enabled = False
        chkFilter.Tag = mbytType
        rptPati.Tag = ""
        rptPati.Records.DeleteAll
        rptPati.Populate
        cboSelectTime(0).Enabled = False
        cboSelectTime(1).Enabled = False
        
        On Error Resume Next
        If mobjSquareCard Is Nothing Then Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        If mobjSquareCard.zlInitComponents(Me, P病案查阅打印, glngSys, UserInfo.用户名, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
        End If
        Call PatiIdentify.zlInit(Me, glngSys, P病案查阅打印, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKindFind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
        chkFilter.ToolTipText = "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。"
        For i = 0 To tbcPati.ItemCount - 1
            If Not tbcPati.Item(i).Selected And tbcPati.Item(i).Visible Then
                tbcPati.Item(i).Visible = False
            End If
        Next
        Call ClearPatiInfo
        Call InitBasicData
    Else
        cboDept.Enabled = True
        optType(0).Enabled = True
        optType(1).Enabled = True
        cboSelectTime(0).Enabled = True
        cboSelectTime(1).Enabled = True
        chkFilter.ToolTipText = "按照查找条件对病人进行过滤显示"
        rptPati.Tag = ""
        rptPati.Records.DeleteAll
        Call optType_Click(Val(chkFilter.Tag))
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picLeft.hwnd
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    PatiIdentify.ActiveFastKey '读卡
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And Not Me.ActiveControl Is PatiIdentify And mstrFindType = "就诊卡" Then
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        
        PatiIdentify.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
    End If
End Sub

Private Sub Form_Load()
    Dim objTab As TabControlItem
    Dim frmTendBody As Object
    Dim intIdx As Integer
    Dim intType As Integer
    Dim objPane As XtremeDockingPane.Pane
    
    '界面恢复：默认搜索类型读取
    '-----------------------------------------------------
    mintPatiCount = 0
    DeleteLISTempFile
    mintFindType = Val(zlDatabase.GetPara("病人查找方式", glngSys, P病案查阅打印, , , , intType))
    mintDeptViewBed = Val(zlDatabase.GetPara("不显示无床位的病区科室", glngSys, P病案查阅打印, , , , intType))
    mintMecStandard = Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
    chkFilter.Value = Val(zlDatabase.GetPara("过滤显示模式", glngSys, P病案查阅打印, , , , intType))
    mblnLIS = Sys.IsSysSetUp(2500)
    mstrPrintDocIDs = ""
    Call InitCommandBar
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 320, 400, DockLeftOf, Nothing)
    objPane.Title = "病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 360, 400, DockRightOf, objPane)
    objPane.Title = "文件列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    '避免与电子病历部件RichEditor冲突,产生多个PDF配置文件,解决电子病历每次输出时都选择路径的问题
    If zlCommFun.PDFInitialize() Then mbytPDFStatu = 1
    If Not gobjEmr Is Nothing Then
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            Set gobjEmr = Nothing
        Else
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    If mclsArchive Is Nothing Then
        Set mclsArchive = New zlMedRecPage.clsArchive
        Call mclsArchive.InitArchiveMedRec(gcnOracle, glngSys)
    End If
    Set mclsOutAdvices = New clsDockOutAdvices
    Set mclsInAdvices = New clsDockInAdvices
    Set mclsDockAduits = New clsDockAduits
    Set mclsPath = New clsDockPath
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    Set frmTendBody = mclsDockAduits.zlGetFormTendBody
    Call zlControl.FormSetCaption(frmTendBody, False, False)
    Call CreateObjectPacs(mobjPublicPACS)
    If mobjReport Is Nothing Then Set mobjReport = New zl9Report.clsReport
    '子窗体
    '-----------------------------------------------------
    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsArchive.zlGetForm(0), "_门诊首页"
    mcolSubForm.Add mclsArchive.zlGetForm(1), "_住院首页"
    mcolSubForm.Add mclsDockAduits.zlGetFormEPR, "_病历信息"
    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
    mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
    mcolSubForm.Add frmTendBody, "_体温记录单"
    mcolSubForm.Add mclsDockAduits.zlGetFormTendFile, "_护理记录单"
    mcolSubForm.Add mclsPath.zlGetForm, "_临床路径"
    mcolSubForm.Add mclsTendsNew.zlGetfrmInTendFile, "_新版护理"
    If Not mobjRichEMR Is Nothing Then mcolSubForm.Add mobjRichEMR.zlGetForm, "_电子病历"
    If Not mobjPublicPACS Is Nothing Then mcolSubForm.Add mobjPublicPACS.zlDocGetForm, "_检查报告"
    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        '隐式出发Form_Load采取添加一个图片方式，切换的时候再依次重新加载
        Set objTab = .InsertItem(intIdx, "门诊首页", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "住院首页", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "门诊医嘱", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "住院医嘱", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "体温记录单", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "护理记录单", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "新版护理", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        End If
        Set objTab = .InsertItem(intIdx, "检查报告", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "报表", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "三方报告", picRpt.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        
    End With
    
    Call ClearPatiInfo
        '---------------------------------------------------
    'tbcPati病人列表
    With Me.tbcPati
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "在院", picPatiIn.hwnd, 0).Tag = "在院"
        .InsertItem(1, "出院", picPatiIn.hwnd, 0).Tag = "出院"
        .InsertItem(2, "在诊", picPatiIn.hwnd, 0).Tag = "在诊"
        .InsertItem(3, "已诊", picPatiIn.hwnd, 0).Tag = "已诊"
        .Item(1).Selected = True
        .Item(0).Selected = True
        '定位病人选项卡
        tbcPati.Item(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", 1)).Selected = True
    End With
    
    '就诊历史
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With
    'RIS接口创建
    Call InitReportColumn
    Call HaveRIS(True)
    If mblnLIS Then Call InitObjLis(P病案查阅打印)
    Call FuncLoadReport
    Call InitSelectTime
    '缺省定位住院
    Call optType_Click(0)
    Call RestoreWinState(Me, App.ProductName, , True)
    Call ShowReportColumn
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Terminate()
     
    Call DeleteLISTempFile
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    If picRpt.Tag <> "" Then mstrTempDel = picRpt.Tag
    Set mclsArchive = Nothing
    Set mclsDockAduits = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsPath = Nothing
    Set mclsTendsNew = Nothing
    Set mobjRichEMR = Nothing
    Set mrsData = Nothing
    Set mobjPublicPACS = Nothing
    Set mobjPatient = Nothing
    Set mobjSquareCard = Nothing
    
    If Not mobjReportForm Is Nothing Then Unload mobjReportForm
    Set mobjReportForm = Nothing
    Set mobjReport = Nothing
    Set mobjInfection = Nothing
End Sub

Private Sub cboVisit_Click()

    If cboVisit.Text = "" Then Exit Sub
    
    If mlngPreDept = cboVisit.ItemData(cboVisit.ListIndex) Then Exit Sub
    
    mlngPreDept = cboVisit.ItemData(cboVisit.ListIndex)
    
    mrsData.Filter = "序号=" & mlngPreDept
    
    mlng就诊ID = mrsData!就诊id
    mlng病人ID = mrsData!病人ID
    mblnMoved = False
    If Not mrsData.EOF Then
        mstr挂号单 = NVL(mrsData!NO, "")
        mblnMoved = Val(NVL(mrsData!数据转出, "")) = 1
    End If
    '显示基本信息
    If mstr挂号单 <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    
    '显示档案目录
    Me.tbcHistory(0).Caption = cboVisit.Text
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
End Sub

Private Sub FuncLookPicture()
'观片功能
    Dim lng医嘱ID As Long
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        lng医嘱ID = Val(Split(tvwArchive.SelectedItem.Tag, ";")(1) & "")
        If lng医嘱ID <> 0 Then
            If CreateObjectPacs(gobjPublicPacs) Then
                Call gobjPublicPacs.ShowImage(lng医嘱ID, Me, mblnMoved)
            End If
        End If
    End If
End Sub

Private Sub lblDept_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim vPoint As POINTAPI
    If mbytType = E_门诊 Then Exit Sub
    If chkFilter.Value = vbChecked Then Exit Sub
    Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_View_Dept, , True)
    If Not objPopup Is Nothing Then
        vPoint.X = lblDept.Left / Screen.TwipsPerPixelX
        vPoint.Y = (lblDept.Top + lblDept.Height + 30) / Screen.TwipsPerPixelY
        ClientToScreen picPati.hwnd, vPoint
        objPopup.CommandBar.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'功能：结束打印事件，写入首页打印数据
    Dim strSQL As String
    
    If mblnPrint And mbytPrintType = 1 Then
        If InStr("," & mstrPrintMedRec & ",", ",0_9,") = 0 Then
            strSQL = "Zl_电子病历打印_Insert(Null,9," & mlng病人ID & "," & mlng就诊ID & ",'" & UserInfo.姓名 & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            '电子病历打印唯一键：(文件ID, 种类, 打印时间) 病案首页:文件ID为NULL,种类:9 打印时间Oracle过程自动提取
            '批量打印首页正反面时,只记录一次
            mstrPrintMedRec = mstrPrintMedRec & "," & "0_9"
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optType_Click(Index As Integer)
    Dim i As Integer
    Dim blnVisible As Boolean
    Dim strCardKind  As String
    mbytType = Index
    fraInPati.Visible = False: fraOutPati.Visible = False
    
    If Index = 0 Then '住院
        fraInPati.Visible = True
        On Error Resume Next
        If mobjSquareCard Is Nothing Then Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        strCardKind = mstrCardKindIN
        If mobjSquareCard.zlInitComponents(Me, P病案查阅打印, glngSys, UserInfo.用户名, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
        End If
        Call PatiIdentify.zlInit(Me, glngSys, P病案查阅打印, gcnOracle, gstrDBUser, mobjSquareCard, strCardKind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
        For i = 0 To tbcPati.ItemCount - 1
            If InStr(",在院,出院,", tbcPati.Item(i).Tag) > 0 Then
                blnVisible = True
            Else
                blnVisible = False
            End If
            tbcPati.Item(i).Visible = blnVisible
        Next
        mblnUndo = True
        tbcPati.Item(0).Selected = True '缺省选中在院
        mblnUndo = False
    ElseIf Index = 1 Then '门诊
        fraOutPati.Visible = True
        '加载门诊科室
        On Error Resume Next
        If mobjSquareCard Is Nothing Then Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        strCardKind = mstrCardKindOut
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlInitComponents(Me, P病案查阅打印, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
            Else
                strCardKind = mobjSquareCard.zlGetIDKindStr(strCardKind)
            End If
        End If
        Call PatiIdentify.zlInit(Me, glngSys, P病案查阅打印, gcnOracle, gstrDBUser, mobjSquareCard, strCardKind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
    
        For i = 0 To tbcPati.ItemCount - 1
            If InStr(",在诊,已诊,", tbcPati.Item(i).Tag) > 0 Then
                blnVisible = True
            Else
                blnVisible = False
            End If
            tbcPati.Item(i).Visible = blnVisible
        Next
        mblnUndo = True
        tbcPati.Item(2).Selected = True '缺省选中在诊
        mblnUndo = False
    End If
    Call LoadDept
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人ID
    End If
    Call ExecuteFindPati(False, lngPatiID)
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strPati As String, vRect As RECT, strName As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPatiID As Long
    
    If chkFilter.Value = vbUnchecked Then Exit Sub
    
    strName = Trim(PatiIdentify.Text)
'    "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。"
    On Error GoTo ErrHand
    blnCancel = False
            
    If objCard.名称 Like "*姓*名*" And blnCard = False And strName <> "" And InStr("-*+/", Left(Trim(PatiIdentify.Text), 1)) = 0 Then

        strPati = "Select 1 As 排序id, a.病人id As ID, a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.住院次数, Trunc(a.入院时间, 'dd') As 入院日期, a.出生日期," & vbNewLine & _
                "       a.身份证号, a.家庭地址, a.工作单位, a.病人类型" & vbNewLine & _
                "From 病人信息 A" & vbNewLine & _
                "Where 姓名 = [1]"
        strPati = strPati & " Order by 排序ID,入院日期 Desc"
        
        vRect = zlControl.GetControlRect(PatiIdentify.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, _
            vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strName)
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!ID) = 0 Then
                blnCancel = True: Exit Sub
            Else '以病人ID读取
                lngPatiID = NVL(rsTmp!ID)
            End If
        Else '取消选择
            If blnCancel = False Then
                MsgBox "没有找到符合条件的病人【" & strName & "】！", vbInformation, gstrSysName
            End If
            blnCancel = True: Exit Sub
        End If
        Call ExecuteFindPati(False, lngPatiID)
        blnCancel = True: Exit Sub
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mintFindType = Index - 1: mstrFindType = objCard.名称
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "住院号", "门诊号"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "床号"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case "就诊卡"
            If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        Case "姓名"
    End Select
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With cboVisit
        .Top = 60
        .Left = 60
        .Width = picLeft.ScaleWidth - 120
        .Height = 300
    End With
    With tbcHistory
        .Left = 60
        .Top = cboVisit.Top + cboVisit.Height
        .Width = picLeft.ScaleWidth - 120
        .Height = picLeft.ScaleHeight - (cboVisit.Height + 60)
    End With
End Sub

Private Sub picPati_GotFocus()
    On Error Resume Next
'    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub picPati_Resize()
    Dim lngSplit As Long
    
    On Error Resume Next
    fraType.Move 0, 0
    lngSplit = 75
     
    lblDept.Top = fraType.Height + lngSplit + (cboDept.Height - lblDept.Height) / 2
    lblDept.Left = lngSplit
    cboDept.Top = fraType.Top + fraType.Height + lngSplit
    cboDept.Left = lblDept.Left + lblDept.Width + lngSplit / 2
    cboDept.Width = picPati.ScaleWidth - cboDept.Left - lblDept.Left
    
    PatiIdentify.Left = cboDept.Left
    PatiIdentify.Top = cboDept.Top + cboDept.Height + lngSplit
    PatiIdentify.Width = cboDept.Width - chkFilter.Width - 60
    
    chkFilter.Left = PatiIdentify.Left + PatiIdentify.Width + 60
    chkFilter.Top = PatiIdentify.Top + (PatiIdentify.Height - chkFilter.Height) / 2
    lblFind.Left = lblDept.Left
    lblFind.Top = lngSplit + PatiIdentify.Top + (PatiIdentify.Height - lblFind.Height) / 2
    lblFind.Width = lblDept.Width

    tbcPati.Left = 0
    tbcPati.Top = PatiIdentify.Top + PatiIdentify.Height + lngSplit
    tbcPati.Width = picPati.ScaleWidth
    tbcPati.Height = picPati.ScaleHeight - tbcPati.Top
    
    picPatiIn.Width = picPati.ScaleWidth
End Sub

Private Sub picPatiIn_Resize()
    Dim lngTop As Long
    Dim i As Long
    On Error Resume Next
    For i = 0 To picPara.Count - 1
        If picPara(i).Visible Then
            picPara(i).Width = picPatiIn.ScaleWidth
            picPara(i).Top = lngTop
            lngTop = picPara(i).Top + picPara(i).Height
        End If
    Next
    rptPati.Move 60, lngTop, picPatiIn.ScaleWidth - 60, picPatiIn.ScaleHeight - lngTop
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Item.Checked = True Then
        mintPatiCount = mintPatiCount + 1
    Else
        mintPatiCount = mintPatiCount - 1
    End If
    If mintPatiCount = rptPati.Records.Count Then
        rptPati.Columns(col_选择).Icon = imgPati.ListImages("Check").Index - 1
    Else
        rptPati.Columns(col_选择).Icon = imgPati.ListImages("UnCheck").Index - 1
    End If
    stbThis.Panels(2).Text = IIf(mintPatiCount = 0, "", "勾选了" & mintPatiCount & "个病人！")
End Sub

Private Sub rptPati_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim lngHit As Long
    
    If Button = 1 Then
        Set hitColumn = rptPati.HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.ItemIndex = col_选择 Then
                lngHit = rptPati.HitTest(X, Y).ht
                If xtpHitTestHeader = lngHit Then
                    If rptPati.Records.Count = 0 Then Exit Sub  '无数据时禁止切换
                    If hitColumn.Icon = imgPati.ListImages("Check").Index - 1 Then
                        hitColumn.Icon = imgPati.ListImages("UnCheck").Index - 1
                        SelectItems 2
                    Else
                        hitColumn.Icon = imgPati.ListImages("Check").Index - 1
                        SelectItems 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Item As ReportRecordItem
    Dim strTipInfo As String
    Dim vPos As POINTAPI
    Dim lngHwnd As Long
    
    On Error Resume Next
    Set hitColumn = rptPati.HitTest(X, Y).Column
    If Not hitColumn Is Nothing Then
        If hitColumn.Index = col_打印图标 Then
            Set Item = rptPati.HitTest(X, Y).Item
            If Not Item Is Nothing Then
                If Item.Record(col_打印图标).Icon <> -1 Then
                    strTipInfo = Item.Record(col_打印记录).Value
                    If strTipInfo = "" Then '如果没有获取过，则立即获取并记录在列表中
                        strTipInfo = GetPrintLog(Item.Record(col_病人Id).Value, Item.Record(col_主页ID).Value) '提取打印记录
                        Item(col_打印记录).Value = strTipInfo
                    End If
                    GetCursorPos vPos
                    lngHwnd = WindowFromPoint(vPos.X, vPos.Y)
                    Call zlCommFun.ShowTipInfo(lngHwnd, strTipInfo, True)
                End If
            End If
        Else
            Call zlCommFun.ShowTipInfo(lngHwnd, "")
        End If
    End If
End Sub

Private Sub rptPati_SelectionChanged()
'功能:
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If mblnUndo Then Exit Sub       '切换门诊住院会触发
    If Not Me.Visible Then Exit Sub
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '非正常情况
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        If .GroupRow Then
            Call ClearPatiInfo
        Else
            '病人照片
            If rptPati.Tag = Val(.Record(col_病人Id).Value & "") & "_" & Val(.Record(col_主页ID).Value & "") Then Exit Sub
            If Not ReadPatPricture(Val(.Record(col_病人Id).Value & ""), imgPatient) Then
               Set imgPatient.Picture = imgList.ListImages("Patient").Picture
            End If
            mlng病人ID = Val(.Record(col_病人Id).Value & "")
            mlng就诊ID = Val(.Record(col_主页ID).Value & "")
            rptPati.Tag = Val(.Record(col_病人Id).Value & "") & "_" & Val(.Record(col_主页ID).Value & "")
            Call InitBasicData
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbcArchive_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    If Item.Handle = picTmp.hwnd Then
        Screen.MousePointer = 11
        Index = Item.Index
        mblnTabTmp = True
        On Error GoTo errH
        Select Case Item.Tag
            Case "门诊首页"
                Set objItem = tbcArchive.InsertItem(Index, "门诊首页", mcolSubForm("_门诊首页").hwnd, 0)
                objItem.Tag = "门诊首页"
            Case "住院首页"
                Set objItem = tbcArchive.InsertItem(Index, "住院首页", mcolSubForm("_住院首页").hwnd, 0)
                objItem.Tag = "住院首页"
            Case "病历信息"
                Set objItem = tbcArchive.InsertItem(Index, "病历信息", mcolSubForm("_病历信息").hwnd, 0)
                objItem.Tag = "病历信息"
            Case "门诊医嘱"
                Set objItem = tbcArchive.InsertItem(Index, "门诊医嘱", mcolSubForm("_门诊医嘱").hwnd, 0)
                objItem.Tag = "门诊医嘱"
            Case "住院医嘱"
                Set objItem = tbcArchive.InsertItem(Index, "住院医嘱", mcolSubForm("_住院医嘱").hwnd, 0)
                objItem.Tag = "住院医嘱"
            Case "体温记录单"
                Set objItem = tbcArchive.InsertItem(Index, "体温记录单", mcolSubForm("_体温记录单").hwnd, 0)
                objItem.Tag = "体温记录单"
            Case "护理记录单"
                Set objItem = tbcArchive.InsertItem(Index, "护理记录单", mcolSubForm("_护理记录单").hwnd, 0)
                objItem.Tag = "护理记录单"
            Case "临床路径"
                Set objItem = tbcArchive.InsertItem(Index, "临床路径", mcolSubForm("_临床路径").hwnd, 0)
                objItem.Tag = "临床路径"
            Case "新版护理"
                Set objItem = tbcArchive.InsertItem(Index, "新版护理", mcolSubForm("_新版护理").hwnd, 0)
                objItem.Tag = "新版护理"
            Case "电子病历"
                Set objItem = tbcArchive.InsertItem(Index, "电子病历", mcolSubForm("_电子病历").hwnd, 0)
                objItem.Tag = "电子病历"
            Case "检查报告"
                Set objItem = tbcArchive.InsertItem(Index, "检查报告", mcolSubForm("_检查报告").hwnd, 0)
                objItem.Tag = "检查报告"
            Case "报表"
                If mobjReportForm Is Nothing Then
                    Set objItem = tbcArchive.InsertItem(Index, "报表", picTmp.hwnd, 0)
                Else
                    Set objItem = tbcArchive.InsertItem(Index, "报表", mobjReportForm.hwnd, 0)
                End If
                objItem.Tag = "报表"
        End Select
        Call tbcArchive.RemoveItem(Index + 1)
        objItem.Selected = True
        mblnTabTmp = False
        Screen.MousePointer = 0
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'功能：切换显示不同的档案页面，或者清空界面
    Dim i As Long
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag = strShow Then
            '默认的卡片跟当前界面要展示的一样时，可能窗体还未绑定上去，这里通过条件判断一下手动绑一次。不会出现多重复执行
            If tbcArchive.Item(i).Handle = picTmp.hwnd Then Call tbcArchive_SelectedChanged(tbcArchive.Item(i))
            tbcArchive(i).Caption = strCaption
            If Not tbcArchive(i).Visible Then
                tbcArchive(i).Visible = True
                tbcArchive(i).Selected = True
                Exit For
            End If
        End If
    Next
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag <> strShow Then
            If tbcArchive(i).Visible Then tbcArchive(i).Visible = False
        End If
    Next
End Sub

Private Sub tbcPati_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    If chkFilter.Value = Checked Then Exit Sub
    For i = 0 To picPara.Count - 1
        picPara(i).Visible = False
    Next
    
    If Item.Tag = "出院" Then
        picPara(0).Visible = True
    ElseIf Item.Tag = "已诊" Then
        picPara(1).Visible = True
    End If
    
    Call picPatiIn_Resize
    
    If Me.Visible And Not mblnUndo Then
        LoadPatients
    End If
End Sub

Private Sub FuncPrintOrView(ByVal bytFunc As Byte)
'功能:预览打印
'参数:bytFunc=2（打印），=1（预览）0=设置;3-PDF
'说明:单个病人打印单个文件;单个病人打印所有选择文件;
'     批量打印单个文件;批量打印多个文件
    Dim objFSO As New Scripting.FileSystemObject    'FSO对象
    Dim rsTemp As ADODB.Recordset
    
    Dim strKey As String
    Dim strPath As String
    Dim strSQL As String
    Dim strRegRange As String
    Dim strFile   As String
    Dim str挂号单   As String
    Dim strPatiName As String
    Dim strPatiNo  As String
    Dim strErr      As String
    
    Dim i As Long, j As Long
    Dim lngSel As Long
    Dim lngNo  As Long
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim lng科室ID As Long
    
    Dim blnMoveData As Boolean
    Dim blnPath     As Boolean
    Dim strDeviceName   As String
    
    
    If mlng病人ID = 0 Then Exit Sub
    On Error GoTo errH
    If bytFunc = 3 Then
        strPath = GetRegister(私有模块, "打印档案", "PDF位置", App.Path)
        If strPath = "" Then SetPDFPath
        If mbytPDFStatu = 0 Then
            If Not zlCommFun.PDFInitialize(strErr) Then
                MsgBox "PDF设备初始化失败！" & strErr, vbExclamation, gstrSysName
                Exit Sub
            Else
                mbytPDFStatu = 1
            End If
        End If
        '检测是否存在TinyPDF(32位系统) Foxit Reader PDF Printer (64位系统)打印机
        strDeviceName = zlCommFun.PDFPrinterDeviceName()
    Else
        If Not LoadPrint Then Exit Sub
    End If
    With tvwArchive
        If bytFunc = 2 Or bytFunc = 3 Then
            Me.MousePointer = 11
            '打印选中项目
            If mintPatiCount = 0 Then
                lngSel = 0
                mstrPrintDocIDs = ""
                mstrPrintMedRec = ""
                If mstr挂号单 = "" Then
                    strPatiName = lblShow(lbl_姓名).Caption
                    strPatiNo = lblShow(lbl_住院号).Caption
                    If bytFunc = 2 Then
                        lngNo = 1
                        strSQL = "Select Nvl(Max(打印次数),0)+1 As 打印次数 From 病案打印记录 Where 病人id=[1] And 主页id=[2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
                        If rsTemp.BOF = False Then
                            lngNo = rsTemp!打印次数
                        End If
                        mblnPrint = True
                    End If
                Else
                    lngNo = -1
                    strPatiName = lblShow(lbl_姓名).Caption
                    strPatiNo = ""
                    mblnPrint = False
                End If
                mrsMedRec.Filter = ""
                Do While Not mrsMedRec.EOF
                    mrsMedRec!选择 = 0
                    mrsMedRec.MoveNext
                Loop
                For i = 1 To .Nodes.Count
                    strKey = .Nodes.Item(i).Key
                    If .Nodes.Item(i).Checked And Not InStr(",R0,R1,R2,R3,R4,R5,R6,R7,R9,R10,R11,R12,", strKey) > 0 Then
                        mrsMedRec.Filter = "ID='" & strKey & "'"
                        mrsMedRec!选择 = 1
                        lngSel = lngSel + 1
                    End If
                Next
                mrsMedRec.Filter = "选择=1"
                If mrsMedRec.RecordCount > 0 Then psb.Visible = True: psb.Max = mrsMedRec.RecordCount
                For i = 1 To mrsMedRec.RecordCount
                    psb.Value = i
                    Call PrintOrView(bytFunc, mlng病人ID, mlng就诊ID, mlng科室ID, mrsMedRec!ID & "", mrsMedRec!参数 & "", strPath, strPatiName, strPatiNo, lngNo, mblnMoved, strDeviceName)
                    mrsMedRec.MoveNext
                Next
                If psb.Visible Then psb.Visible = False
                If lngSel = 0 Then
                    strKey = .SelectedItem.Key
                    Call PrintOrView(bytFunc, mlng病人ID, mlng就诊ID, mlng科室ID, strKey, .SelectedItem.Tag, strPath, strPatiName, strPatiNo, lngNo, mblnMoved, strDeviceName)
                End If
            Else
                '批量打印
                '统计并记录打印类别
                strRegRange = ""
                For i = 1 To .Nodes.Count
                    If .Nodes.Item(i).Checked And InStr("R1,R2,R3,R4,R5,R6,R7,R8,R9,R10,R11,R12,", .Nodes.Item(i).Key) > 0 Then
                        If .Nodes.Item(i).Key = "R8" Then
                            strRegRange = strRegRange & " OR ID ='R8'"
                        Else
                            strRegRange = strRegRange & " OR 上级ID ='" & .Nodes.Item(i).Key & "'"
                        End If
                    End If
                Next
                If strRegRange <> "" Then strRegRange = Mid(strRegRange, 5)
                If strRegRange = "" Then
                    MsgBox "请在文件列表中选择需要输出的文件类型。", vbInformation, Me.Caption
                    GoTo errMsg
                End If
                For i = 0 To rptPati.Records.Count - 1
                    If rptPati.Records(i).Item(col_选择).Checked = True Then
                        With rptPati.Records(i)
                            lng病人ID = Val(.Item(col_病人Id).Value & "")
                            lng主页ID = Val(.Item(col_主页ID).Value & "")
                            lng科室ID = Val(.Item(col_科室ID).Value & "")
                            
                            strPatiName = .Item(col_姓名).Value & ""
                            mstrPrintDocIDs = ""
                            mstrPrintMedRec = ""
                            If mbytType = E_住院 Then
                                str挂号单 = ""
                                blnMoveData = Val(.Item(COL_数据转出).Value & "") = 1
                                strPatiNo = .Item(col_住院号).Value & ""
                                blnPath = False
                                If GetInsidePrivs(p临床路径应用) <> "" Then
                                    blnPath = HavePath(lng科室ID)
                                End If
                                If bytFunc = 2 Then
                                    lngNo = 1
                                    strSQL = "Select Nvl(Max(打印次数),0)+1 As 打印次数 From 病案打印记录 Where 病人id=[1] And 主页id=[2]"
                                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
                                    If rsTemp.BOF = False Then
                                        lngNo = rsTemp!打印次数
                                    End If
                                    mblnPrint = True
                                End If
                            Else
                                lngNo = -1
                                str挂号单 = .Item(COL_NO).Value
                                blnMoveData = False
                                blnPath = False
                                mblnPrint = False
                            End If

                            Set rsTemp = GetCISStruct(lng病人ID, lng主页ID, str挂号单, blnPath, blnMoveData)
                            If Not rsTemp Is Nothing Then
                                rsTemp.Filter = strRegRange
                                If rsTemp.RecordCount > 0 Then psb.Visible = True: psb.Max = rsTemp.RecordCount
                                For j = 1 To rsTemp.RecordCount
                                    psb.Visible = j
                                    Call PrintOrView(bytFunc, lng病人ID, lng主页ID, lng科室ID, rsTemp!ID & "", rsTemp!参数 & "", strPath, strPatiName, strPatiNo, lngNo, blnMoveData, strDeviceName)
                                    rsTemp.MoveNext
                                Next
                                If psb.Visible Then psb.Visible = False
                            End If
                        End With
                    End If
                Next
            End If
            mstrPrintMedRec = ""
            Me.MousePointer = 0
        ElseIf bytFunc = 1 Then
            If InStr(",R0,R1,R2,R3,R4,R5,R6,R7,R9,R10,R11,R12,", "," & .SelectedItem.Key & ",") > 0 Then Exit Sub
            Call PrintOrView(1, mlng病人ID, mlng就诊ID, mlng科室ID, .SelectedItem.Key, .SelectedItem.Tag, , , , , mblnMoved)
        ElseIf bytFunc = 0 Then
            Call PrintInMedRec(Nothing, 0, mlng病人ID, mlng就诊ID, mobjReport, mlng科室ID, Me)
        End If
    End With
    Exit Sub
errMsg:
    If bytFunc > 1 Then Me.MousePointer = 0
    Exit Sub
errH:
    If bytFunc > 1 Then Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintOrView(ByVal bytFunc As Byte, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, _
    ByVal strKey As String, Optional ByVal strPara As String, Optional ByVal strPath As String, Optional ByVal strPatiName As String, _
    Optional ByVal strHosNo As String, Optional ByVal lngNo As Long, Optional ByVal blnMoveData As Boolean, Optional ByVal strDeviceName As String)
'功能:预览打印
'参数:bytFunc=2（打印），=1（预览）0=设置;3-PDF
    Dim intPage As Integer
    Dim intMode As Integer
    Dim intSel  As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim varParam As Variant
    Dim blnMod  As Boolean
    Dim blnFoxitPDF As Boolean
    
    Dim strSQL          As String
    Dim strReportNO     As String
    Dim strInfo          As String
    Dim strMsg          As String
    Dim strFileName     As String
    Dim strType         As String
    Dim strPrint        As String
    Dim rsTemp          As ADODB.Recordset
    Dim objPrint        As Object
    Dim strPDFPath      As String

    On Error GoTo errH

    If InStr(",R0,R1,R2,R3,R4,R5,R6,R7,R9,R10,R11,R12,", "," & strKey & ",") > 0 Then Exit Sub
    '获取缺省打印机
    varParam = Split(strPara, ";")
    If bytFunc = 1 Or bytFunc = 2 Then
        If InStr(strPara, ";EMR;") > 0 Then
            strPrint = mcolPrint(varParam(4))
        ElseIf strKey Like "R7P*" Or strKey Like "R7L*" Then
            strPrint = mcolPrint("R7")
        Else
            On Error Resume Next
            strPrint = mcolPrint(Split(strKey, "K")(0))
            If err.Number <> 0 Then strPrint = mcolPrint("R11")
            err.Clear: On Error GoTo errH
        End If
    End If
    If strKey Like "R11K*" Then
        intPage = Val(Replace(strKey, "R11K", ""))
        If bytFunc = 3 Then
            If intPage = 1 Then
                strType = "首页正面"
            ElseIf intPage = 2 Then
                strType = "首页反面"
            ElseIf intPage = 3 Then
                strType = "首页附页一"
            ElseIf intPage = 4 Then
                strType = "首页附页二"
            End If
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & strType & ".PDF"
            Call zlCommFun.PDFFile(strFileName)
            Call PrintInMedRec(Nothing, 4, lng病人ID, lng主页ID, mobjReport, lng科室ID, Me, intPage, strFileName)
        Else
            If bytFunc = 2 Then mbytPrintType = 1 '标记首页打印
            Call PrintInMedRec(Nothing, bytFunc, lng病人ID, lng主页ID, mobjReport, lng科室ID, Me, intPage, IIf(bytFunc = 2, strPrint, ""))
            If bytFunc = 2 Then mbytPrintType = 0 '取消标记
        End If
    ElseIf strKey Like "R12K*" Then
        If strKey = "R12K1" Then
            If bytFunc = 3 Then
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_临时医嘱.PDF"
                Call zlCommFun.PDFFile(strFileName)
            End If
             '再打印临嘱
            Call gobjKernel.zlPrintAdvice(Me, lng病人ID, lng主页ID, 0, 1, IIf(bytFunc = 3, strFileName, strPrint), IIf(bytFunc = 3, 4, bytFunc))
        ElseIf strKey = "R12K2" Then
            If bytFunc = 3 Then
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_长期医嘱.PDF"
                Call zlCommFun.PDFFile(strFileName)
            End If
            '先打印长嘱
            Call gobjKernel.zlPrintAdvice(Me, lng病人ID, lng主页ID, 0, 0, IIf(bytFunc = 3, strFileName, strPrint), IIf(bytFunc = 3, 4, bytFunc))
        End If
    ElseIf strKey Like "R1K*" Or strKey Like "R2K*" Or strKey Like "R5K*" Or strKey Like "R6K*" Or strKey Like "R4K*" Then  '1-门诊病历,2-住院病历;4-护理病历;5-疾病证明;6-知情文件
        If bytFunc = 1 Then
            If strKey Like "R5K*" And Val(varParam(5)) = 2 Then '预览疾病传染报告卡
                Call FuncViewDisReportCard(Val(varParam(0)))
            Else
                Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrint, blnMoveData)
            End If
        Else
            '住院病历和护理病历 存在共享文档时只打印输出一次
            If (strKey Like "R2K*" Or strKey Like "R4K*") And InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") > 0 Then Exit Sub      '
            If bytFunc = 3 Then
                If strKey Like "R1K*" Then
                    strType = "门诊病历"
                ElseIf strKey Like "R2K*" Then
                    strType = "住院病历"
                ElseIf strKey Like "R4K*" Then
                    strType = "护理病历"
                ElseIf strKey Like "R5K*" Then
                    strType = "疾病证明"
                ElseIf strKey Like "R6K*" Then
                    strType = "知情文件"
                End If
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & strType & "_" & varParam(0) & ".PDF"
                Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
            End If
            Call mclsDockAduits.zlPrintDocument(3, 2, Val(varParam(0)), IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
            If bytFunc = 2 Then
                Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
            End If
        End If
    ElseIf strKey Like "R3K*" Then   '3-护理记录
        If mblnNewTends = False Then
            If UBound(varParam) >= 1 Then
                If Val(varParam(1)) = -1 Then '体温单
                    If bytFunc = 3 Then
                        strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_护理记录_" & Val(varParam(3)) & ".PDF"
                        Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                        strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    End If
                    Call mclsDockAduits.zlRefreshTendBody(lng病人ID, lng主页ID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)), blnMoveData)
                    Call mclsDockAduits.zlPrintDocument(1, IIf(bytFunc = 3, 2, bytFunc), , IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                    
                    If bytFunc = 2 Then
                        Call RecordEprPrintInfo(2, "体温单", lngNo, lng病人ID, lng主页ID)
                    End If
                Else '护理记录
                    If bytFunc = 3 Then
                        strFileName = strPath & "\" & strFileName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_护理记录_" & Val(varParam(3)) & ".PDF"
                        Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                        strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    End If
                    Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), lng病人ID, lng主页ID, Val(varParam(0)), CStr(varParam(2)), , Val(varParam(4)), blnMoveData)
                    Call mclsDockAduits.zlPrintDocument(2, IIf(bytFunc = 3, 2, bytFunc), , IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                    
                    If bytFunc = 2 Then
                        Call RecordEprPrintInfo(3, Val(varParam(3)), lngNo, lng病人ID, lng主页ID)
                    End If
                End If
            End If
        Else
            If UBound(varParam) >= 1 Then
                Select Case Val(varParam(1))
                    Case -1 '体温单
                        intSel = 1
                    Case 1  '产程图
                        intSel = 3
                    Case Else '记录单
                        intSel = 2
                End Select
                strInfo = "开始输出" & strPatiName & Decode(intSel, 1, "体温单", 2, "护理记录", "产程图") & "_" & Val(varParam(3))
                If bytFunc = 3 Then
                    strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & Decode(intSel, 1, "体温单", 2, "护理记录", "产程图") & "_" & Val(varParam(3)) & ".PDF"
                    Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                End If
                Call mclsTendsNew.zlPrintDocument(lng病人ID, lng主页ID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), intSel, IIf(bytFunc = 3, strDeviceName, strPrint), IIf(bytFunc = 1, False, True))
                If bytFunc = 2 Then
                    Call RecordEprPrintInfo(2, Decode(intSel, 1, "体温单", 2, "护理记录", "产程图"), lngNo, lng病人ID, lng主页ID)
                End If
            End If
        End If
    ElseIf strKey Like "R7K*" Then   '7-诊疗报告
        If bytFunc = 3 Then
            If Val(varParam(0)) = 0 Then
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & varParam(4) & "_" & varParam(2) & ".PDF"
            Else
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & varParam(4) & "_" & Val(varParam(0)) & ".PDF"
            End If
            Call zlCommFun.PDFFile(strFileName)
            strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
        End If
        If varParam(2) = "E" Then
            If mstr检验对应报表 <> "" Then
                strReportNO = Split(mstr检验对应报表, ",")(2)
                If bytFunc = 3 Then
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & lng病人ID, "主页id=" & lng主页ID, "医嘱ID=" & varParam(1), "PDF=" & strFileName, 4)
                Else
                    Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, strReportNO, "printer", strPrint)  '设置指定打印机
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & lng病人ID, "主页id=" & lng主页ID, "医嘱ID=" & varParam(1), bytFunc)
                End If
            ElseIf mblnLIS And Val(mstr检验报告打印) = 1 And Not gobjLIS Is Nothing Then
                blnMod = gobjLIS.PrintLisReport(Me, Val(varParam(1)), IIf(bytFunc = 3, 4, bytFunc), , IIf(bytFunc = 3, strFileName, ""), strPrint, strMsg)     '1-预览;2-打印;4-PDF
            Else
                Call mclsDockAduits.zlPrintDocument(4, IIf(bytFunc = 3, 2, bytFunc), Val(varParam(0)), IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                If bytFunc = 3 Then blnFoxitPDF = True
            End If
        ElseIf varParam(2) = "D" Then
            If mstr检查对应报表 <> "" Then
                strReportNO = Split(mstr检查对应报表, ",")(2)
                If bytFunc = 3 Then
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & lng病人ID, "主页id=" & lng主页ID, "医嘱ID=" & varParam(1), "PDF=" & strFileName, 4)
                Else
                    Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, strReportNO, "printer", strPrint)  '设置指定打印机
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "病人id=" & lng病人ID, "主页id=" & lng主页ID, "医嘱ID=" & varParam(1), bytFunc)
                End If
            Else
                If Val(varParam(3)) <> 0 Then
                    'RIS
                    If Not gobjRis Is Nothing Then
                        If bytFunc = 1 Then
                            Call gobjRis.ShowViewReport(Me.hwnd, Val(varParam(1)), True, Val(varParam(3)))
                        ElseIf bytFunc = 2 Then
                            Call gobjRis.HISPrintReportByAppno(Me.hwnd, Val(varParam(1)), Val(varParam(3)), strPrint)
                        ElseIf bytFunc = 3 Then
                            strPDFPath = gobjRis.HISGetPrintFile(Me.hwnd, 1, Val(varParam(1)), Val(varParam(3)))
                            If strPDFPath <> "" Then
                                If objFile.FileExists(strPDFPath) And strPath <> "" Then
                                    objFile.MoveFile strPDFPath, strPath & "\"
                                End If
                            End If
                        End If
                    End If
                ElseIf Val(varParam(0)) <> 0 Then
                    Call mclsDockAduits.zlPrintDocument(4, IIf(bytFunc = 3, 2, bytFunc), Val(varParam(0)), IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                    If bytFunc = 3 Then blnFoxitPDF = True
                End If
            End If
        End If
    ElseIf strKey Like "R7P*" Then  '检查报告
        If Not mobjPublicPACS Is Nothing Then
            If bytFunc = 3 Then strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & varParam(1) & ".PDF"
            Call mobjPublicPACS.PrintReport(varParam(0), IIf(bytFunc = 3, strFileName, strPrint), IIf(bytFunc = 1, True, False))  'True预览
            If bytFunc = 3 Then blnFoxitPDF = True
        End If
    ElseIf strKey Like "R7L*" Then  '三方报告
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & Split(varParam(2), "<sTab>")(1) & "_" & varParam(0) & ".PDF"
            Call zlCommFun.PDFFile(strFileName)
            strFileName = GetLisRptFile(strPara, strFileName)
        End If
    ElseIf strKey Like "R8*" Then  '8-临床路径
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_临床路径.PDF"
            Call mclsPath.zlFuncPathTableOutPut(4, True, strFileName, lng病人ID, lng主页ID, strDeviceName)
            blnFoxitPDF = True
        Else
            Call mclsPath.zlRefreshReadOnly(lng病人ID, lng主页ID)
            Call mclsPath.zlFuncPathTableOutPut(IIf(bytFunc = 1, 2, 1), True, "", 0, 0, strPrint) '2-预览;1-打印
            Call RecordEprPrintInfo(2, "临床路径", lngNo, lng病人ID, lng主页ID)
        End If
    ElseIf strKey Like "R9K*" Then   '9-住院证
        strReportNO = "ZLCISBILL" & Format(varParam(0), "00000") & "-1"
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_住院证.PDF"
            Call zlCommFun.PDFFile(strFileName)
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "NO=" & varParam(1), "性质=" & varParam(2), "医嘱ID=0", "PDF=" & strFileName, 4)
        Else
            Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, strReportNO, "printer", strPrint)  '设置指定打印机
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "NO=" & varParam(1), "性质=" & varParam(2), "医嘱ID=0", bytFunc)
        End If
    ElseIf strKey Like "R10K*" Then   '10-其他报表
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_其他报表_" & varParam(0) & ".PDF"
            Call zlCommFun.PDFFile(strFileName)
            Call mobjReport.ReportOpen(gcnOracle, 0, varParam(2), Me, "病人id=" & lng病人ID, "主页id=" & lng主页ID, "PDF=" & strFileName, 4)
        Else
            Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, varParam(2), "printer", strPrint)  '设置指定打印机
            Call mobjReport.ReportOpen(gcnOracle, 0, varParam(2), Me, "病人id=" & lng病人ID, "主页id=" & lng主页ID, bytFunc)
        End If
    ElseIf InStr(strKey, "R") = 0 And Len(strPara) >= 32 Then
        'EMR病历预览
        If Not mobjRichEMR Is Nothing Then
            If varParam(1) <> "" Then
                Call mobjRichEMR.zlShowDoc(varParam(0), varParam(1))
            Else
                Call mobjRichEMR.zlShowDoc(varParam(0), "")
            End If
            If bytFunc = 1 Then
                Call mobjRichEMR.zlPrintDoc(True)
            ElseIf bytFunc = 2 Or bytFunc = 3 Then
                '存在共享文档时只打印输出一次
                If InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") > 0 Then Exit Sub
                If bytFunc = 3 Then
                    strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng主页ID & "_" & varParam(2) & varParam(0) & ".PDF"
                    Call zlCommFun.PDFFile(strFileName)
                    strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                End If
                err.Number = 0: On Error Resume Next
                Call mobjRichEMR.zlPrintDoc(False, IIf(bytFunc = 3, strDeviceName, strPrint))
                If err.Number = 450 Then
                    err.Number = 0
                    Call mobjRichEMR.zlPrintDoc(False)
                End If
                err.Clear: On Error GoTo 0
                mstrPrintDocIDs = mstrPrintDocIDs & varParam(0) & ","  '
                If bytFunc = 3 Then blnFoxitPDF = True
            End If
        End If
    End If
    If blnFoxitPDF Then
        Call zlCommFun.PDFFileSuccess
    End If
    Exit Sub
errH:
    mstrPrintDocIDs = ""
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwArchive_NodeCheck(ByVal node As MSComctlLib.node)
    CheckNode node
End Sub

Private Sub tvwArchive_NodeClick(ByVal node As MSComctlLib.node)
    Dim arrPar As Variant
    Dim intSel As Integer
    Dim strFile As String
    Dim i As Long
    If mblnTvwTmp Then Exit Sub
    If tvwArchive.Tag = node.Key Then Exit Sub
    mblnTvwTmp = True
    LockWindowUpdate Me.hwnd
    
    arrPar = Split(node.Tag, ";")
     
    If node.Key Like "R1K*" Or node.Key Like "R2K*" Or node.Key Like "R4K*" Or node.Key Like "R5K*" Or node.Key Like "R6K*" Or node.Key Like "R7K*" Then
        Call ShowArchiveTab("病历信息", node.Text)
    End If
    mblnSeePic = False
    If node.Key = "R11" Or node.Key = "R0" Then
        Call ShowArchiveTab(IIf(mstr挂号单 <> "", "门诊首页", "住院首页"), tbcHistory.Selected.Caption)
        Call mclsArchive.zlRefresh(IIf(mstr挂号单 <> "", 0, 1), mlng病人ID, mlng就诊ID, mblnMoved)
    ElseIf node.Key = "R12" Then  '医嘱记录
        If mstr挂号单 <> "" Then
            Call ShowArchiveTab("门诊医嘱", tbcHistory.Selected.Caption)
            Call mclsOutAdvices.zlRefresh(mlng病人ID, mstr挂号单, False, mblnMoved)
        Else
            Call ShowArchiveTab("住院医嘱", tbcHistory.Selected.Caption)
            Call mclsInAdvices.zlRefresh(mlng病人ID, mlng就诊ID, mlng病区ID, mlng科室ID, 0, mblnMoved)
        End If
    ElseIf node.Key Like "R1K*" Then '门诊病历
        Call mclsDockAduits.zlRefresh(1, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R2K*" Then '住院病历
        Call mclsDockAduits.zlRefresh(2, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R3K*" Then '护理记录
        If UBound(arrPar) >= 1 Then
            If mblnNewTends = False Then
                If Val(arrPar(1)) = -1 Then
                    Call ShowArchiveTab("体温记录单", node.Text)
                    Call mclsDockAduits.zlRefreshTendBody(mlng病人ID, mlng就诊ID, Val(arrPar(0)), 0, mblnMoved)
                Else
                    Call ShowArchiveTab("护理记录单", node.Text)
                    Call mclsDockAduits.zlRefresh(3, Val(arrPar(3)), mlng病人ID, mlng就诊ID, Val(arrPar(0)), CStr(arrPar(2)), , , mblnMoved)
                End If
            Else
                Select Case Val(arrPar(1))
                    Case -1
                        intSel = 0
                    Case 1
                        intSel = 2
                    Case Else
                        intSel = 1
                End Select
                Call ShowArchiveTab("新版护理", node.Text)
                Call mclsTendsNew.zlRefreshTendFile(mlng病人ID, mlng就诊ID, Val(arrPar(4)), Val(arrPar(0)), False, IIf(glngModul = p住院医生站, True, False), intSel, Val(arrPar(3)), 1)
            End If
        End If
    ElseIf node.Key Like "R4K*" Then '护理病历
        Call mclsDockAduits.zlRefresh(4, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R5K*" Then '疾病证明
        Call mclsDockAduits.zlRefresh(5, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R6K*" Then '知情文件
        Call mclsDockAduits.zlRefresh(6, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R7K*" Then '诊疗报告
        Call mclsDockAduits.zlRefresh(7, Val(arrPar(0)), , , , , , , mblnMoved)
        mblnSeePic = arrPar(2) = "D"
    ElseIf node.Key = "R8" Then
        If mstr挂号单 = "" Then
            Call ShowArchiveTab("临床路径", node.Text)
            Call mclsPath.zlRefreshReadOnly(mlng病人ID, mlng就诊ID)
        End If
    ElseIf node.Key Like "R7P*" Then  '检查报告
        mblnSeePic = True
        Call ShowArchiveTab("检查报告", node.Text)
        If Not mobjPublicPACS Is Nothing Then Call mobjPublicPACS.zlDocRefresh(Split(node.Tag, ";")(0))
    ElseIf node.Key Like "R7L*" Then  '三方报告
        strFile = GetLisRptFile(node.Tag)
        If strFile <> "" Then
            If picRpt.Tag <> "" And picRpt.Tag <> mstrTempDel And picRpt.Tag <> strFile Then mstrTempDel = picRpt.Tag
            webRpt.Navigate strFile
            picRpt.Tag = strFile
        End If
        Call ShowArchiveTab("三方报告", node.Text)
    ElseIf node.Key Like "R9K*" Or node.Key Like "R10K*" Or node.Key Like "R11K*" Or node.Key Like "R12K*" Then   '住院证;其他报表;首页；医嘱
        Call FuncShowReport(node)
    ElseIf InStr(node.Key, "R") = 0 And Len(node.Tag) >= 32 Then
        'EMR病历预览
        If Not mobjRichEMR Is Nothing Then
            Call ShowArchiveTab("电子病历", node.Text)
            If arrPar(1) <> "" Then
                Call mobjRichEMR.zlShowDoc(arrPar(0), arrPar(1))
            Else
                Call mobjRichEMR.zlShowDoc(arrPar(0), "")
            End If
        End If
    Else
        LockWindowUpdate 0
        mblnTvwTmp = False
        Exit Sub
    End If
    tvwArchive.Tag = node.Key
    mblnTvwTmp = False
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
    LockWindowUpdate 0
End Sub

Private Function ShowArchiveTree() As Boolean
'功能：显示病人档案树形目录
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, objNode As node, strSQL1 As String
    Dim blnPath As Boolean
    Dim strSel As String
    Dim strKey As String
    Dim i As Long
    
    Screen.MousePointer = 11
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        strKey = tvwArchive.SelectedItem.Key
        If strKey Like "R11K*" Or strKey = "R11" Or strKey Like "R12K*" Or strKey = "R12" Then
            strSel = Split(strKey, "K")(0)
        End If
    End If
    
    '病人科室存在可用的临床路径时，显示临床路径记录
    If mstr挂号单 = "" Then
        If GetInsidePrivs(p临床路径应用) <> "" Then
            blnPath = HavePath(mlng科室ID)
        End If
    End If
    Set rsTmp = GetCISStruct(mlng病人ID, mlng就诊ID, mstr挂号单, blnPath, mblnMoved)
    Set mrsMedRec = rsTmp

    tvwArchive.Tag = ""
    tvwArchive.Nodes.Clear

    Do While Not rsTmp.EOF
        If NVL(rsTmp!上级ID) = "" Then
            Set objNode = tvwArchive.Nodes.Add(, , CStr(rsTmp!ID), rsTmp!名称, NVL(rsTmp!图标))
        Else
            Set objNode = tvwArchive.Nodes.Add(CStr(rsTmp!上级ID), tvwChild, CStr(rsTmp!ID), rsTmp!名称, NVL(rsTmp!图标))
        End If

        objNode.Tag = NVL(rsTmp!参数)
        objNode.Expanded = True

        If tvwArchive.Nodes.Count = 1 Then
            objNode.Selected = True
        ElseIf objNode.Key = strSel Then
            objNode.Selected = True
        End If

        rsTmp.MoveNext
    Loop
    If Not tvwArchive.SelectedItem Is Nothing Then
        tvwArchive.SelectedItem.EnsureVisible
        Call tvwArchive_NodeClick(tvwArchive.SelectedItem)
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadDept()
'功能：门诊按科室方式读取可用部门;住院 按科室/病区方式读取可用部门
'注意:此方法需要
    If mbytType = E_住院 Then
        lblDept.Caption = IIf(mintDeptView = 0, "科室(&D)↓", "病区(&D)↓")
    Else
        lblDept.Caption = "科室(&D)"
    End If
    mintPreDept = -1
    Call InitDepts
    Call cboDept_Click
    
    If cboDept.ListIndex = -1 Then
        If mbytType = E_住院 Then
            If InStr(mstrPrivs, "全院病人") > 0 Then
                MsgBox "没有发现住院" & IIf(mintDeptView = 0, "科室", "病区") & "信息,请先到部门管理中设置！", vbInformation, gstrSysName
            Else
                MsgBox "没有发现你所属" & IIf(mintDeptView = 0, "科室", "病区") & ",不能使用住院医生工作站！", vbInformation, gstrSysName
            End If
        Else
            If InStr(mstrPrivs, "全院病人") > 0 Then
                MsgBox "没有发现门诊科室信息,请先到部门管理中设置！", vbInformation, gstrSysName
            Else
                MsgBox "没有发现你所属科室,不能使用病案查询打印！", vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function LoadPatients(Optional ByVal lngPatiID As Long) As Boolean
'功能:读取病人信息
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngColor As Long
    Dim strFilter As String
    Dim strCaption As String
    Dim strTemp As String
    Dim intBedLen As Integer
    
    On Error GoTo errH
    '床位长度固定为10
    intBedLen = 10
    
    If lngPatiID <> 0 Then
        strSQL = "Select 就诊id,场合,科室,在院,执行状态 " & vbNewLine & _
                "From (Select a.主页id As 就诊id, a.入院日期 As 时间, 1 As 场合, 出院科室id As 科室, Decode(a.出院日期, Null, 1, 0) As 在院, NULL as 执行状态 " & vbNewLine & _
                "       From 病案主页 A" & vbNewLine & _
                "       Where a.病人id = [1] And Nvl(主页id, 0) <> 0" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select ID As 就诊id, b.执行时间 As 时间, 0 As 场合, b.执行部门id As 科室, 0 As 在院,执行状态 " & vbNewLine & _
                "       From 病人挂号记录 B" & vbNewLine & _
                "       Where b.病人id = [1] And 记录性质 = 1 And 记录状态 = 1 And 执行状态 In (1,2)) A" & vbNewLine & _
                "Order By a.时间 Desc"

        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID)
        If Not rsPati.EOF Then
            If Val(rsPati!场合 & "") = 1 Then
                mbytType = E_住院
                If (rsPati!在院 & "") = 1 Then
                    strCaption = "在院"
                    strSQL = "Select Distinct b.病人id, b.主页id, b.编目日期, b.住院号,b.留观号,b.姓名, b.性别, a.年龄, b.家庭地址, a.身份证号, a.出生日期, b.入院日期, b.出院日期, b.住院医师,B.数据转出, b.病人类型, b.病人性质," & vbNewLine & _
                            "       Decode(d.病人id || '_' || d.主页id, '_', 0, 1) As 是否打印, b.出院科室id as 科室ID, B.住院医师,a.就诊卡号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号 " & vbNewLine & _
                            "From 病人信息 A, 病案主页 B, 病案打印记录 D" & vbNewLine & _
                            "Where a.病人id = b.病人id And b.病人id = d.病人id(+) And b.主页id = d.主页id(+) And b.病人id = [1] And b.主页id = [2]"
                    
                Else
                    strCaption = "出院"
                    strSQL = "Select Distinct b.病人id, b.主页id, b.编目日期,b.住院号,b.留观号,b.姓名, b.性别, a.年龄, b.家庭地址, a.身份证号, a.出生日期, b.入院日期, b.出院日期, b.住院医师,B.数据转出, b.病人类型, b.病人性质," & vbNewLine & _
                            "  Decode(d.病人id || '_' || d.主页id, '_', 0, 1) As 是否打印,b.出院科室id as 科室ID, b.住院医师,a.就诊卡号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号 " & vbNewLine & _
                            "From 病人信息 A, 病案主页 B, 病案打印记录 D" & vbNewLine & _
                            "Where a.病人id = b.病人id And b.病人id = d.病人id(+) And b.主页id = d.主页id(+) And b.病人id = [1] And b.主页id = [2]"
                    
                End If
                
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, Val(rsPati!就诊id & ""))
                If rsPati.RecordCount > 0 Then
                    If InStr(mstrPrivs, "本科病人") = 0 And InStr(mstrPrivs, "全院病人") = 0 Then
                        rsPati.Filter = "住院医师 ='" & UserInfo.姓名 & "'"
                        If rsPati.RecordCount = 0 Then
                            MsgBox "用户【" & UserInfo.姓名 & "】权限不足,不允许操作该病人【" & rsPati!姓名 & "】。", vbInformation, gstrSysName
                            Exit Function
                        End If
                    ElseIf InStr(mstrPrivs, "全院病人") = 0 Then
                        strSQL = "Select 1 From 部门人员 Where 部门id = [1] And 人员id = [2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NVL(rsPati!出院科室ID), UserInfo.ID)
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "用户【" & UserInfo.姓名 & "】权限不足,不允许操作该病人【" & rsPati!姓名 & "】。", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            Else
                mbytType = E_门诊
                If Val(rsPati!执行状态 & "") = 1 Then '已诊
                    strCaption = "已诊"
                ElseIf Val(rsPati!执行状态 & "") = 2 Then '正在就诊
                    strCaption = "在诊"
                End If
                strSQL = "Select Distinct b.Id, b.No, b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.执行时间, b.执行人, b.执行部门id, a.家庭地址, a.病人类型, a.身份证号, a.出生日期,a.就诊卡号" & vbNewLine & _
                        "From 病人信息 A, 病人挂号记录 B" & vbNewLine & _
                        "Where b.病人id = a.病人id And b.Id = [1]"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsPati!就诊id & ""))
                If rsPati.RecordCount > 0 Then
                    If InStr(mstrPrivs, "本科病人") = 0 And InStr(mstrPrivs, "全院病人") = 0 Then
                        rsPati.Filter = "执行人 ='" & UserInfo.姓名 & "'"
                        If rsPati.RecordCount = 0 Then
                            MsgBox "用户【" & UserInfo.姓名 & "】权限不足,不允许操作该病人【" & rsPati!姓名 & "】。", vbInformation, gstrSysName
                            Exit Function
                        End If
                    ElseIf InStr(mstrPrivs, "全院病人") = 0 Then
                        strSQL = "Select 1 From 部门人员 Where 部门id = [1] And 人员id = [2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NVL(rsPati!执行部门ID), UserInfo.ID)
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "用户【" & UserInfo.姓名 & "】权限不足,不允许操作该病人【" & rsPati!姓名 & "】。", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
 
            For i = 0 To tbcPati.ItemCount - 1
                If strCaption = tbcPati.Item(i).Tag Then
                    tbcPati.Item(i).Visible = True
                    mblnUndo = True
                    tbcPati.Item(i).Selected = True '缺省选中在诊
                    mblnUndo = False
                Else
                    tbcPati.Item(i).Visible = False
                End If
            Next
        Else
            Exit Function
        End If
    Else
        If mbytType = E_住院 Then
            '病案主页.状态：0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院
            If mintDeptView = 0 Then
                '在院病人
                If tbcPati.Selected.Tag = "在院" Then
                    strSQL = "Select Distinct b.病人id, b.主页id, b.编目日期, b.住院号,b.留观号, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, NVL(b.年龄,a.年龄) As 年龄, b.家庭地址, a.身份证号, a.出生日期, b.入院日期," & vbNewLine & _
                                "       b.出院日期, b.住院医师,b.病人类型, B.数据转出, b.病人性质, Decode(d.病人id, Null, 0, 1) As 是否打印, R.科室ID,a.就诊卡号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号  " & vbNewLine & _
                                "From 病人信息 A,病案主页 B,病案打印记录 D, 在院病人 R" & vbNewLine & _
                                "Where r.病人id = d.病人id(+) And r.主页id = d.主页id(+) And r.病人id = a.病人id And r.病人id = b.病人id " & vbNewLine & _
                                " And r.主页id = b.主页id And (R.科室ID=[1] Or b.婴儿科室ID=[1]) And b.状态<>1" & _
                                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And b.住院医师=[2]")
                ElseIf tbcPati.Selected.Tag = "出院" Then
                    strFilter = " And B.出院日期 Between to_date('" & Format(mdatOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdatOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
                    
                    strSQL = "Select Distinct b.病人id, b.主页id, b.编目日期, b.住院号,b.留观号, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, NVL(b.年龄,a.年龄) As 年龄, b.家庭地址, a.身份证号, a.出生日期, b.入院日期," & vbNewLine & _
                            "       b.出院日期, b.住院医师, b.病人类型, B.数据转出,b.病人性质, Decode(d.病人id, Null, 0, 1) As 是否打印, B.出院科室ID As 科室ID,a.就诊卡号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号  " & vbNewLine & _
                            "From 病人信息 A, 病案主页 B, 病案打印记录 D" & vbNewLine & _
                            "Where a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And b.出院科室id + 0 = [1]" & vbNewLine & _
                            IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And b.住院医师=[2]") & vbNewLine & _
                            " And b.封存时间 Is Null " & strFilter
                    
                End If
            Else
                '按病区查看
                '在院病人
                If tbcPati.Selected.Tag = "在院" Then
                    strSQL = "Select Distinct b.病人id, b.主页id, b.编目日期, b.住院号,b.留观号, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, NVL(b.年龄,a.年龄) As 年龄, b.家庭地址, a.身份证号, a.出生日期, b.入院日期," & vbNewLine & _
                                "       b.出院日期, b.住院医师,b.病人类型,b.数据转出,b.病人性质, Decode(d.病人id, Null, 0, 1) As 是否打印, R.科室ID,a.就诊卡号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号  " & vbNewLine & _
                                "From 病人信息 A,病案主页 B,部门表 C,病案打印记录 D, 在院病人 R" & vbNewLine & _
                                "Where r.病人id = d.病人id(+) And r.主页id = d.主页id(+) And r.病人id = a.病人id And r.病人id = b.病人id " & vbNewLine & _
                                " And r.主页id = b.主页id And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And b.状态<>1" & _
                                " And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & vbNewLine & _
                                IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And b.住院医师=[2]") & vbNewLine & _
                                " And b.封存时间 Is Null " & strFilter
                     
                Else
                    strFilter = " And B.出院日期 Between to_date('" & Format(mdatOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdatOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
                    strSQL = "Select Distinct b.病人id, b.主页id, b.编目日期, b.住院号,b.留观号,Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, NVL(b.年龄,a.年龄) As 年龄, b.家庭地址, a.身份证号, a.出生日期, b.入院日期," & vbNewLine & _
                            "       b.出院日期, b.住院医师, b.病人类型, b.病人性质,b.数据转出, Decode(d.病人id, Null, 0, 1) As 是否打印, B.出院科室ID As 科室ID,a.就诊卡号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号 " & vbNewLine & _
                            "From 病人信息 A, 病案主页 B,部门表 C, 病案打印记录 D" & vbNewLine & _
                            "Where a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And B.当前病区ID+0=[1] " & vbNewLine & _
                            " And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & vbNewLine & _
                             IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And b.住院医师=[2]") & vbNewLine & _
                            " And b.封存时间 Is Null " & strFilter
        
                End If
            End If
        Else
            '门诊病人
            If tbcPati.Selected.Tag = "在诊" Then
                strSQL = "Select B.ID, b.No, b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.执行时间, b.执行人, b.执行部门ID, a.家庭地址, a.病人类型, a.身份证号, a.出生日期,a.就诊卡号 " & vbNewLine & _
                    "From 病人信息 A, 病人挂号记录 B" & vbNewLine & _
                    "Where b.病人id = a.病人id And b.记录性质 = 1 And b.记录状态 = 1 And b.执行状态 = 2 And b.执行部门ID = [1]" & vbNewLine & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And b.执行人=[2]")
            ElseIf tbcPati.Selected.Tag = "已诊" Then
                strSQL = "Select B.ID, b.No, b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.执行时间, b.执行人, b.执行部门ID, a.家庭地址, a.病人类型, a.身份证号, a.出生日期,a.就诊卡号 " & vbNewLine & _
                    "From 病人信息 A, 病人挂号记录 B" & vbNewLine & _
                    "Where b.病人id = a.病人id And b.记录性质 = 1 And b.记录状态 = 1 And b.执行状态 = 1 And b.执行部门ID = [1]" & vbNewLine & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And b.执行人=[2]")
                strSQL = strSQL & " And B.执行时间 Between To_Date('" & Format(mDatBegin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mDatEnd, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            End If
        End If
        If strSQL = "" Then
            rptPati.Tag = ""
            rptPati.Records.DeleteAll
            rptPati.Populate
            Screen.MousePointer = 0
            LoadPatients = True
            Exit Function
        End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.姓名)
    End If

    '加载病人列表
    Call ClearPatiInfo
    rptPati.Tag = ""
    rptPati.Records.DeleteAll
    mintPatiCount = 0
    stbThis.Panels(2).Text = ""
     
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati.Records.Add()
        Set objItem = objRecord.AddItem("")    '图标
        objItem.HasCheckbox = True
        Set objItem = objRecord.AddItem("")  '图1标
        If InStr(rsPati!性别 & "", "男") > 0 Then
            objItem.Icon = imgPati.ListImages("Boy").Index - 1
        ElseIf InStr(rsPati!性别 & "", "女") > 0 Then
            objItem.Icon = imgPati.ListImages("Girl").Index - 1
        End If
        

        If tbcPati.Selected.Tag = "出院" Then
            Set objItem = objRecord.AddItem("")  '图标
            If Val(rsPati!是否打印 & "") = 1 Then
                objItem.Icon = imgPati.ListImages("print").Index - 1
            End If
            objRecord.AddItem IIf(NVL(rsPati!编目日期) <> "", "已编目", "未编目")
            objRecord.AddItem Format(rsPati!编目日期 & "", "YYYY-MM-dd")
        Else
            objRecord.AddItem ""   '图标
            objRecord.AddItem ""
            objRecord.AddItem ""
        End If
        
        If mbytType = E_住院 Then
            objRecord.AddItem rsPati!住院号 & ""
            objRecord.AddItem ""
            objRecord.AddItem ""
            Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(rsPati!床号), 10)) 'Value用于排序
            objItem.Caption = CStr(Trim(NVL(rsPati!床号, " "))) '为空时会被Value替代
        Else
            objRecord.AddItem ""
            objRecord.AddItem rsPati!NO & ""
            objRecord.AddItem rsPati!门诊号 & ""
            objRecord.AddItem ""
        End If
        
        objRecord.AddItem rsPati!姓名 & ""
        objRecord.AddItem rsPati!性别 & ""
        objRecord.AddItem rsPati!年龄 & ""
        objRecord.AddItem rsPati!身份证号 & ""
        objRecord.AddItem Format(rsPati!出生日期 & "", "YYYY-MM-DD")
        If mbytType = E_住院 Then
            objRecord.AddItem ""
        Else
            objRecord.AddItem Format(rsPati!执行时间 & "", "YYYY-MM-DD")
        End If
       
        If mbytType = E_住院 Then
            objRecord.AddItem Format(rsPati!入院日期 & "", "YYYY-MM-DD")
            objRecord.AddItem Format(rsPati!出院日期 & "", "YYYY-MM-DD")
            objRecord.AddItem rsPati!住院医师 & ""
        Else
            objRecord.AddItem ""
            objRecord.AddItem ""
            objRecord.AddItem rsPati!执行人 & ""
        End If
        objRecord.AddItem rsPati!家庭地址 & ""
        objRecord.AddItem rsPati!就诊卡号 & ""
        If mbytType = E_住院 Then
            objRecord.AddItem rsPati!留观号 & ""
        Else
            objRecord.AddItem ""
        End If
        '隐藏列
        objRecord.AddItem rsPati!病人类型 & ""
        objRecord.AddItem CLng(rsPati!病人ID)

        If mbytType = E_住院 Then
            objRecord.AddItem NVL(rsPati!主页ID)
            objRecord.AddItem rsPati!科室ID & ""
            objRecord.AddItem rsPati!数据转出 & ""
        Else
            objRecord.AddItem NVL(rsPati!ID)
            objRecord.AddItem rsPati!执行部门ID & ""
            objRecord.AddItem ""
        End If
         '显示病人颜色
        lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
        objRecord.Item(col_姓名).ForeColor = lngColor

        rsPati.MoveNext
    Next
    Call ShowReportColumn
    rptPati.Populate
    If chkFilter.Value = vbChecked Then
        If rptPati.Records.Count > 0 Then Set rptPati.FocusedRow = rptPati.Rows(0)
    Else
        Call InitBasicData '清空文件列表
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowOutPatiInfo() As Boolean
'功能：选择门诊病人某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    fraOutPati.Visible = True: fraInPati.Visible = False
    If mlng病人ID <> 0 Then
        strSQL = "Select B.Id,B.NO,B.门诊号,B.姓名,B.性别,B.年龄,A.医疗付款方式,A.身份证号,A.家庭地址," & _
            " A.费别,A.险类,A.出生日期,A.医保号,B.急诊,B.发生时间,B.执行人,B.执行状态,B.执行时间," & _
            " B.执行部门ID as 科室ID,B.诊室,B.社区,D.社区号,C.名称 as 科室" & _
            " From 病人信息 A,病人挂号记录 B,部门表 C,病人社区信息 D" & _
            " Where A.病人ID=B.病人ID And B.ID=[1] And B.执行部门ID=C.ID" & _
            " And B.病人ID=D.病人ID(+) And B.社区=D.社区(+) And B.记录性质=1 And B.记录状态=1"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng就诊ID)
        With rsTmp
            '保险病人姓名红色显示
            lblShow(lblOUT_姓名).Caption = NVL(!姓名)
            
            If Not IsNull(!险类) Then
                lblShow(lblOUT_姓名).ForeColor = vbRed
            Else
                lblShow(lblOUT_姓名).ForeColor = lblShow(lblOUT_门诊号).ForeColor
            End If
            lblShow(lblOUT_门诊号).Caption = NVL(!门诊号)
            lblShow(lblOUT_年龄).Caption = NVL(!年龄)
            lblShow(lblOUT_性别).Caption = NVL(!性别)
            lblShow(lblOUT_身份证号).Caption = NVL(!身份证号)
            lblShow(lblOUT_就诊日期).Caption = NVL(!执行时间)
            lblShow(lblOUT_家庭地址).Caption = NVL(!家庭地址)
            lblShow(lblOUT_出生日期).Caption = NVL(!出生日期)
            lblShow(lblOUT_门诊医师).Caption = NVL(!执行人)
            lbl急.Visible = NVL(!急诊, 0) <> 0
            mlng科室ID = NVL(!科室ID, 0)
            mlng病区ID = 0
        End With
    Else
        Call ClearPatiInfo
    End If
    ShowOutPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'功能:清除病人信息显示框
    Dim i As Long
    mlng病人ID = 0
    mlng就诊ID = 0
    mlng科室ID = 0
    mlng病区ID = 0
    Set imgPatient.Picture = imgList.ListImages("Patient").Picture
    For i = lblShow.LBound To lblShow.UBound
        lblShow(i).Caption = ""
    Next

End Sub

Private Function ShowInPatiInfo() As Boolean
'功能：选择某次住院记录时，读取相关的病人信息
'返回：blnMoved=本次住院病案是否转出了
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    fraInPati.Visible = True: fraOutPati.Visible = False
    If mlng病人ID <> 0 Then
        strSQL = "Select NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄," & _
            " NVL(B.家庭地址,A.家庭地址) As 家庭地址, B.住院号,B.出院病床,B.医疗付款方式," & _
            " D.信息值 as 医保号,B.险类,B.当前病况,C.名称 as 护理等级,B.入院日期," & _
            " B.出院日期,B.病人类型,B.状态,B.出院科室ID,B.当前病区ID,A.住院次数,A.身份证号 " & _
            " From 病人信息 A,病案主页 B,收费项目目录 C,病案主页从表 D" & _
            " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2] And B.护理等级ID=C.ID(+)" & _
            " And B.病人ID=D.病人ID(+) And B.主页ID=D.主页ID(+) And D.信息名(+)='医保号'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    
        With rsTmp
            '保险病人颜色特殊显示
            lblShow(lbl_姓名).Caption = NVL(!姓名)
            lblShow(lbl_姓名).ForeColor = zlDatabase.GetPatiColor(NVL(!病人类型))

            lblShow(lbl_住院号).Caption = NVL(!住院号)
            lblShow(lbl_性别).Caption = NVL(!性别)
            lblShow(lbl_年龄).Caption = NVL(!年龄)
            lblShow(lbl_身份证号).Caption = NVL(!身份证号)
            lblShow(lbl_家庭地址).Caption = NVL(!身份证号)
            '危重病人病况红色显示
            lblShow(lbl_病况).Caption = NVL(!当前病况)
            If NVL(!当前病况) = "危" Or NVL(!当前病况) = "重" Or NVL(!当前病况) = "急" Then
                lblShow(lbl_病况).ForeColor = vbRed
            Else
                lblShow(lbl_病况).ForeColor = lblList(lbl_住院号).ForeColor
            End If
            lblShow(lbl_入院日期).Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
            If Not IsNull(!出院日期) Then
                lblShow(lbl_入院日期).Caption = lblShow(lbl_入院日期).Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")
            End If
            mlng科室ID = NVL(!出院科室ID, 0)
            mlng病区ID = NVL(!当前病区ID, 0)
        End With
    Else
        Call ClearPatiInfo
    End If
    ShowInPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strExtendTag As String, strReturn As String, strSQL As String, strSQLNew As String
    
    On Error GoTo errH
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '上级ID，ID，名称，参数，图标
    strSQL = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') 上级id," & vbNewLine & _
            "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
            "       e.Title ||" & vbNewLine & _
            "        Decode(d.Completor, Null, ''," & vbNewLine & _
            "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
            "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As 参数, 'object_case' As 图标" & vbNewLine & _
            "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
            "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
            "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
            "             c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
            "Where d.Antetype_Id = e.Id  And e.Title = Decode(e.Type, 3, d.Subdoc_Title, e.Title)" & vbNewLine & _
            "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Complete_Time"
            
    strSQLNew = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') 上级id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
                "       e.Title ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
                "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As 参数, 'object_case' As 图标" & vbNewLine & _
                "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor, c.Order_No" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
                "             c.Real_Doc_Id Is Not Null And Nvl(c.Intead, 0) = 0) D, Antetype_List E" & vbNewLine & _
                "Where d.Antetype_Id = e.Id " & vbNewLine & _
                "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Order_No"
    
    err.Clear
    On Error Resume Next
    strReturn = gobjEmr.OpenSQLRecordset(strSQLNew, strExtendTag & "^16^etag", rsTemp)
    If err.Number <> 0 Or strReturn <> "" Then
        err.Clear
        strReturn = gobjEmr.OpenSQLRecordset(strSQL, strExtendTag & "^16^etag", rsTemp)
    End If
    
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    If mbytType = E_门诊 Then
        GetEMRIn_Tag = "MZ_" & mlng就诊ID
    Else
        strSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                    "From (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And Nvl(附加床位, 0) = 0) A," & vbNewLine & _
                    "     (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 1 And Nvl(附加床位, 0) = 0) B"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病人入院ID", lngPatiID, lngPageID)
        
        If rsTmp Is Nothing Then Exit Function
        If NVL(rsTmp!ID) = "" Then Exit Function
        GetEMRIn_Tag = "BD_" & rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetLisRptFile(ByVal strTag As String, Optional ByVal strFile As String) As String
'功能：打开LIS报告文件查看，获取临时文件路径
'
    Dim lngRetu As Long, strInfo As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng报告ID As String
    Dim str报告名 As String
    Dim lng类型 As String
    Dim varTmp As Variant
    Dim strSuffix As String '文件后缀名
    
    Screen.MousePointer = 11
    
    varTmp = Split(strTag, ";")
    lng报告ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng类型 = varTmp(0)
    If lng类型 = 0 Then
        strSuffix = "pdf"
    ElseIf lng类型 = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str报告名 = varTmp(1) & "_" & lng报告ID
    If strFile = "" Then
        strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng报告ID & "." & strSuffix
    End If
    If Not objFile.FileExists(strFile) Then
        strFile = Sys.ReadLob(glngSys, 22, lng报告ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Private Sub picRpt_Resize()
    On Error Resume Next
    webRpt.Move 0, 0, picRpt.Width, picRpt.Height
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strDeptIDs As String, lngPreDept As Long
    
    If cboDept.ListIndex <> -1 Then
        lngPreDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    cboDept.Clear
    
    On Error GoTo errH
    Set rsTmp = GetDataToDepts()
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPreDept Then '保留原有定位
            Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
        ElseIf InStr(mstrPrivs, "全院病人") > 0 Then
            If UserInfo.部门ID = rsTmp!ID And (lngPreDept = 0 Or cboDept.ListIndex = -1) Then '直接所属优先
                Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
            If InStr("," & strDeptIDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then
                Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        Else
            '所属缺省病区包含的可能有多个
            If rsTmp!缺省 = 1 And cboDept.ListIndex = -1 Then
                Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDataToDepts(Optional ByVal strIn As String = "") As ADODB.Recordset
'功能：获取科室病区列表数据记录集
'参数：strIn 过滤条件
'      bytFunc=0 住院;1-门诊
    Dim strSQL As String
    Dim blnYN As Boolean
    Dim strDeptIDs As String
    
    If strIn <> "" Then blnYN = True

    If mbytType = E_住院 Then
        If mintDeptView = 0 Then
            '按科室读取显示
            '包含门急诊观察室的病人还没有上床，不加只显床上有病人的科室的限制
            If InStr(mstrPrivs, "全院病人") > 0 Then
                strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where B.部门ID=A.ID And B.工作性质='临床'" & _
                    " And ((B.服务对象 IN(2,3) " & _
                    IIf(mintDeptViewBed = 1, " And Exists (Select 1 From 床位状况记录 C,  病区科室对应 D Where D.病区ID = c.病区id and A.ID = D.科室ID) ", "") & _
                    ")Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
            Else
                '求有权限的科室：本身所在科室+所属病区包含的科室
                strSQL = _
                    " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                    " From 部门表 A,部门性质说明 B,部门人员 C" & _
                    " Where B.部门ID=A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                    " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                    " And B.工作性质='临床'"
                strSQL = strSQL & " Union " & _
                    " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) As 缺省" & _
                    " From 部门人员 A,病区科室对应 B,部门表 C" & _
                    " Where A.部门ID=B.病区ID And B.科室ID=C.ID And A.人员ID=[1]" & _
                    " And Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.病区ID)" & _
                    " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.病区ID)" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    IIf(blnYN, " And (C.编码 Like [2] Or C.简码 Like [3] Or C.名称 Like [3])", "") & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
                If InStr(mstrPrivs, "ICU病人") > 0 Then
                    strSQL = strSQL & " Union " & _
                        " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                        " From 部门表 A" & _
                        " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                        " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='临床')" & _
                        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                End If
                strSQL = "Select ID,编码,名称,Max(缺省) As 缺省 From (" & strSQL & ") Group By ID,编码,名称 Order by 编码"
            End If
        Else
            '按病区读取显示
            If InStr(mstrPrivs, "全院病人") > 0 Then
                strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 B " & _
                    " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                    IIf(mintDeptViewBed = 1, " And Exists (Select 1 From 床位状况记录 C Where A.ID = c.病区id) ", "") & _
                    " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                    " Order by A.编码"
            Else
                '求有权病区：直接所在病区+所在科室所属病区
                strSQL = _
                    " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                    " From 部门表 A,部门性质说明 B,部门人员 C" & _
                    " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                    " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                    " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                strSQL = strSQL & " Union " & _
                    " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) as 缺省" & _
                    " From 部门人员 A,病区科室对应 B,部门表 C" & _
                    " Where A.部门ID=B.科室ID And B.病区ID=C.ID And A.人员ID=[1]" & _
                    " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.科室ID)" & _
                    " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.科室ID)" & _
                    " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    IIf(blnYN, " And (C.编码 Like [2] Or C.简码 Like [3] Or C.名称 Like [3])", "") & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
                If InStr(mstrPrivs, "ICU病人") > 0 Then
                    strSQL = strSQL & " Union " & _
                        " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                        " From 部门表 A" & _
                        " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                        " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='护理')" & _
                        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                End If
                strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
            End If
        End If
    ElseIf mbytType = E_门诊 Then
        If InStr(mstrPrivs, "全院病人") > 0 Then
            strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where B.部门ID=A.ID And B.工作性质='临床' And B.服务对象 In(1,3) " & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
                    " Order by A.编码"
        Else
            strSQL = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
                " From 部门人员 A,部门表 B,部门性质说明 C" & _
                " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
                " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1]" & _
                IIf(blnYN, " And (B.编码 Like [2] Or B.简码 Like [3] Or B.名称 Like [3])", "") & _
                " Order by B.编码"
        End If
    End If
 
    On Error GoTo errH
    If blnYN Then
        Set GetDataToDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%")
    Else
        Set GetDataToDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitSelectTime()
    Dim datCurr As Date
     
   datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
   mdatOutEnd = datCurr
   mdatOutBegin = mdatOutEnd - 1
   mDatEnd = datCurr
   mDatBegin = mDatEnd - 1
   
   cboSelectTime(0).Clear '出院
   With cboSelectTime(0)
       .AddItem "今天内"
       .ItemData(.NewIndex) = 0
       .AddItem "昨天内"
       .ItemData(.NewIndex) = 1
       .AddItem "前天内"
       .ItemData(.NewIndex) = 2
       .AddItem "一周内"
       .ItemData(.NewIndex) = 7
       .AddItem "30天内"
       .ItemData(.NewIndex) = 30
       .AddItem "60天内"
       .ItemData(.NewIndex) = 60
       .AddItem "[指定...]"
       .ItemData(.NewIndex) = -1
   End With
   If cboSelectTime(0).ListCount > 0 Then cboSelectTime(0).ListIndex = 0

   cboSelectTime(1).Clear
   With cboSelectTime(1)
       .AddItem "今天"
       .ItemData(.NewIndex) = 0
       .AddItem "昨天(含今天)"
       .ItemData(.NewIndex) = 1
       .AddItem "一周内"
       .ItemData(.NewIndex) = 7
       .AddItem "[指定...]"
       .ItemData(.NewIndex) = -1
   End With
   If cboSelectTime(1).ListCount > 0 Then cboSelectTime(1).ListIndex = 0

End Sub

Private Sub CheckNode(ByVal node As MSComctlLib.node, Optional ByVal bytFunc As Byte = 0)
'功能:父节点勾选|不选则子节点对应勾选|不选；
'   子节点所有都不勾选，则父节点也不选。
'   子节点只要一个不勾选，父节点也默认不勾选。
'参数:Node-当前结点
    Dim objNode As MSComctlLib.node
    If bytFunc = 0 Then
        If node.Children > 0 Then
            If Not node.Child Is Nothing Then
                Set objNode = node.Child
                objNode.Checked = node.Checked
                If objNode.Children > 0 Then Call CheckNode(objNode)
                Do While Not objNode.Next Is Nothing
                    Set objNode = objNode.Next
                    objNode.Checked = node.Checked
                    If objNode.Children > 0 Then Call CheckNode(objNode)
                Loop
            End If
        End If

        If Not node.Parent Is Nothing Then Call CheckNode(node, 1)
    Else
        '子节点存在一个不选,父节点不勾选;当前节点如果存在父节点，且所有叶子节点都选择,那么对应的父节点也选中
        If node.Checked = False Then
            If node.Parent.Checked = True Then
                node.Parent.Checked = False
                If Not node.Parent.Parent Is Nothing Then Call CheckNode(node.Parent, 1)
            End If
        Else
            If node.Parent.Checked = False Then
                Set objNode = node.FirstSibling
                Do While Not objNode Is Nothing
                    If objNode.Checked = False Then Exit Sub
                    Set objNode = objNode.Next
                Loop
                node.Parent.Checked = True
                If Not node.Parent.Parent Is Nothing Then Call CheckNode(node.Parent, 1)
            End If
        End If
    End If
End Sub

Private Sub FuncLoadReport()
    Dim objControl As CommandBarControl
    Dim objPop As Object
    Dim strHide As String
    Dim i As Long
    
    strHide = ",ZL1_INSIDE_1254_1,ZL1_INSIDE_1254_2,ZL1_INSIDE_1261_1,ZL1_INSIDE_1261_4,ZL1_INSIDE_1261_5,ZL1_INSIDE_1261_6,ZL1_INSIDE_1261_7,ZL1_INSIDE_1261_8,ZL1_INSIDE_1261_9,ZL1_INSIDE_1261_10,"
    mstr检查对应报表 = zlDatabase.GetPara("检查对应报表", glngSys, P病案查阅打印)
    mstr检验对应报表 = zlDatabase.GetPara("检验对应报表", glngSys, P病案查阅打印)
    If mstr检查对应报表 <> "" Then strHide = strHide & "," & Split(mstr检查对应报表, ",")(2) & ","
    If mstr检验对应报表 <> "" Then strHide = strHide & "," & Split(mstr检验对应报表, ",")(2) & ","
    mstr检验报告打印 = zlDatabase.GetPara("检验报告打印", glngSys, P病案查阅打印)
    '清空缓存
    Set mcolReport = New Collection
    For i = 1 To cbsMain.ActiveMenuBar.Controls.Count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "报表*" Then
                cbsMain.ActiveMenuBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    Call zlDatabase.ShowReportMenu(cbsMain, glngSys, P病案查阅打印, mstrPrivs, strHide)
    
    For i = 1 To cbsMain.ActiveMenuBar.Controls.Count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "报表*" Then
            Set objControl = cbsMain.ActiveMenuBar.Controls.Item(i)
            Exit For
        End If
    Next
    
    If Not objControl Is Nothing Then
        With objControl.CommandBar.Controls
            For i = 1 To .Count
                Set objPop = .Item(i)
                mcolReport.Add Split(objPop.Caption, "(&")(0) & "," & objPop.Parameter, "_" & i     '报表名称,系统号,报表编号
            Next
        End With
    End If
End Sub

Private Sub SetPrintPara()
    Dim i As Long
    Dim objFrm As New frmParaSet
    Dim lngRow As Long
    
    Call objFrm.ShowMe(Me, glngSys, P病案查阅打印, mstrPrivs)
    '重新加载
    Call FuncLoadReport
End Sub

Private Function RecordEprPrintInfo(ByVal bytMode As Byte, ByVal strRecordKey As String, ByVal lngNo As Long, Optional ByVal lngPatientKey As Long, Optional ByVal lngPatientPageKey As Long) As Boolean
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    If Not mblnPrint Then Exit Function
    
    If lngNo = 0 Then
        lngNo = 1
        strSQL = "Select Nvl(Max(打印次数),0)+1 As 打印次数 From 病案打印记录 Where 病人id=[1] And 主页id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", lngPatientKey, lngPatientPageKey)
        If rsTmp.BOF = False Then
            lngNo = rsTmp("打印次数").Value
        End If
    End If
    
    Select Case bytMode
    Case 1
        strSQL = "Select 病人id,主页id,病历名称 From　电子病历记录 a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_病案打印记录_Insert(" & Val(rs("病人id").Value) & "," & Val(rs("主页id").Value) & "," & lngNo & ",'" & rs("病历名称").Value & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    Case 2
        strSQL = "Zl_病案打印记录_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & strRecordKey & "','" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
    Case 3
        strSQL = "Select 名称 From　病历文件列表 a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_病案打印记录_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & rs("名称").Value & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    End Select
    
    RecordEprPrintInfo = True
    
End Function

Private Sub SelectItems(ByVal bytFunc As Byte)
'参数:
'   bytFunc=1 全选,=2取消全选
    Dim i As Long
    
    With rptPati
        For i = 0 To .Records.Count - 1
            If bytFunc = 1 Then
                .Records(i).Item(col_选择).Checked = True
            Else
                .Records.Record(i).Item(0).Checked = False
            End If
        Next
        mintPatiCount = IIf(bytFunc = 1, .Records.Count, 0)
    End With
    stbThis.Panels(2).Text = IIf(mintPatiCount = 0, "", "勾选了" & mintPatiCount & "个病人！")
End Sub

Private Function GetPrintLog(ByVal lngPatient As Long, ByVal lngPageID As Long) As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 打印次数 As 打印次, 打印内容, 打印人, 打印时间 From 病案打印记录 Where 病人id = [1] And 主页id = [2] Order By 打印时间, 打印序号"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, lngPageID)
    Do Until rs.EOF
        
        GetPrintLog = GetPrintLog & vbCrLf & Rpad(rs!打印人, 10) & Rpad(Format(rs!打印时间, "yyyy-mm-dd hh:MM"), 20) & Rpad(rs!打印内容, 40)
        rs.MoveNext
    Loop
    GetPrintLog = Rpad("打印人", 10) & Rpad("打印时间", 20) & Rpad("打印内容", 40) & GetPrintLog
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    Rpad = zl9ComLib.zlStr.Rpad(strCode, lngLen, strChar, True)
End Function

Private Function GetCISStruct(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal blnPath As Boolean, ByVal blnDataMove As Boolean) As ADODB.Recordset
'参数:lng主页ID 住院病人为主页id,门诊病人为挂号id
    Dim strSQL As String, strSQL1 As String
    Dim rsTmp As ADODB.Recordset
    Dim rsMedRec As ADODB.Recordset
    Dim strRptIDs As String
    
    Dim i As Long
    
    On Error GoTo errH
    '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明;6-知情文件;7-诊疗报告,11-首页信息,12-医嘱记录,8-临床路径;9-住院证;10-其他报表
    strSQL = strSQL & _
        " Select 'R0' As ID, '' As 上级id, '所有文件' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'00' As 排序 From Dual Union All" & _
        " Select 'R11' As ID, 'R0' As 上级id, '首页信息' As 名称, '' As 参数,0 As 末级,'home' As 图标,'01' As 排序 From Dual Union All" & _
        " Select 'R12' As ID, 'R0' As 上级id, '医嘱记录' As 名称, '' As 参数,0 As 末级,'object_advice' As 图标,'02' As 排序 From Dual Union All" & _
        " Select 'R1' As ID, 'R0' As 上级id, '门诊病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'03' As 排序 From Dual Where [3]=0 Union All" & _
        " Select 'R2' As ID, 'R0' As 上级id, '住院病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'04' As 排序 From Dual Where [3]=1 Union All" & _
        " Select 'R3' As ID, 'R0' As 上级id, '护理记录' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'05' As 排序 From Dual Where [3]=1 Union All" & _
        " Select 'R4' As ID, 'R0' As 上级id, '护理病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'06' As 排序 From Dual Where [3]=1 Union All" & _
        " Select 'R7' As ID, 'R0' As 上级id, '诊疗报告' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'07' As 排序 From Dual Union All" & _
        " Select 'R5' As ID, 'R0' As 上级id, '疾病证明' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'08' As 排序 From Dual Union All" & _
        " Select 'R6' As ID, 'R0' As 上级id, '知情文件' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'09' As 排序 From Dual "
    strSQL = strSQL & _
            IIf(blnPath, " Union All Select 'R8' As ID, 'R0' As 上级id, '临床路径' As 名称, '' As 参数,0 As 末级,'Path' As 图标,'10' As 排序 From Dual", "") & _
            IIf(str挂号单 = "", " Union All Select 'R9' As ID, 'R0' As 上级id, '住院证' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'11' As 排序 From Dual", "") & _
            IIf(str挂号单 = "", " Union All Select 'R10' As ID, 'R0' As 上级id, '其他报表' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'12' As 排序 From Dual", "")
    If lng病人ID = 0 Then
        strSQL = " Select * From (" & strSQL & ") Order By Decode(上级id,Null,' ',上级id),排序"
        Set GetCISStruct = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 0, 0, IIf(mbytType = E_住院, 1, 0))
        Exit Function
    End If
    '首页子项部分:首页正面,首页反面,首页附页一,首页附页二
    If str挂号单 = "" Then
        strSQL = strSQL & _
                " Union All Select 'R11K1' As ID, 'R11' As 上级id, '首页正面' As 名称, '' As 参数,1 As 末级,'home' As 图标,'1' As 排序 From Dual Where [3]=1 " & _
                " Union All Select 'R11K2' As ID, 'R11' As 上级id, '首页反面' As 名称, '' As 参数,1 As 末级,'home' As 图标,'2' As 排序 From Dual Where [3]=1 "
        If mintMecStandard = 1 Or mintMecStandard = 2 Then
            strSQL = strSQL & _
                  " Union All Select 'R11K3' As ID, 'R11' As 上级id, '首页附页一' As 名称, '' As 参数,1 As 末级,'home' As 图标,'3' As 排序 From Dual Where [3]=1 " & _
                  " Union All Select 'R11K4' As ID, 'R11' As 上级id, '首页附页二' As 名称, '' As 参数,1 As 末级,'home' As 图标,'4' As 排序 From Dual Where [3]=1 "
        End If
    End If
    '医嘱部分
    If str挂号单 = "" Then
        strSQL = strSQL & " Union All Select 'R12K1' As ID, 'R12' As 上级id, '临时医嘱' As 名称, '' As 参数,1 As 末级,'object_advice' As 图标,'1' As 排序 From Dual"
        strSQL = strSQL & " Union All Select 'R12K2' As ID, 'R12' As 上级id, '长期医嘱' As 名称, '' As 参数,1 As 末级,'object_advice' As 图标,'2' As 排序 From Dual"
        '住院证
        strSQL = strSQL & " Union All" & _
            " Select * From (Select 上级id||'K'||ID as ID, 上级id, 名称, 参数, 末级, 图标, 排序" & vbNewLine & _
                "From (Select a.Id, 'R9' As 上级id, c.名称||'【操作时间：'|| To_Char(D.发送时间, 'yyyy-mm-dd hh24:mi')||' 操作人：'||D.发送人||'】' as 名称," & vbNewLine & _
                "  c.编号|| ';' ||d.No || ';' || d.记录性质 As 参数, 1 As 末级, 'Folder' As 图标, To_Char(d.发送时间, 'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                "       From 病人医嘱记录 A, 病人医嘱发送 D, 病历单据应用 B, 病历文件列表 C" & vbNewLine & _
                "       Where a.Id = d.医嘱id And a.病人id = [1] And a.诊疗项目id = b.诊疗项目id And b.病历文件id = c.Id And c.种类 = 7 And" & vbNewLine & _
                "             b.应用场合 = 1 And c.名称 Like '%住院证%'" & vbNewLine & _
                "       Order By D.发送时间 Desc)) Where Rownum<10 "
    End If
    
    '病历部分
    'ID=上级ID+K病历ID,医嘱ID,0
    '参数=病历ID;医嘱ID
    strSQL = strSQL & " Union All" & _
        " Select A.上级id||'K'||Trim(To_Char(A.ID))||','||Trim(To_Char(Nvl(A.医嘱id,0)))||',0' As ID,A.上级id," & _
        "       Decode(A.医嘱id,Null,A.名称||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')',A.名称||'：'||B.医嘱内容||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')') As 名称," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.医嘱id,Null,'0',Trim(To_Char(A.医嘱id))) || ';'|| B.诊疗类别 || ';'|| A.RISID||';'|| A.名称||';'||A.编辑方式 As 参数," & _
        "       1 As 末级,Decode(病历种类,1,'object_case',2,'object_case',4,'object_case',7,'object_report','object_file') As 图标,排序 " & _
        " From (Select A.ID, 'R'||A.病历种类 As 上级id, A.病历名称 As 名称,C.医嘱id,C.RISID,A.病历种类,A.编辑方式,A.创建时间,To_Char(A.创建时间,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
        "       From 电子病历记录 A,病人医嘱报告 C " & _
        "       Where A.病人id = [1] And A.主页id = [2] And (A.病人来源=2 And [3]=1 Or Nvl(A.病人来源,0)<>2 And [3]=0)" & _
        "           And C.病历id(+)=A.ID And A.病历种类 In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,病人医嘱记录 B Where A.医嘱id=B.Id(+)"
    '护理部分
    'ID=上级ID+K文件ID,0,科室ID
    '参数=科室ID;保留;开始～截止;文件ID
    '检查本次病人是使用的是老板还是新版
    strSQL1 = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL1, "检查是否存在老板数据", lng病人ID, mlng就诊ID)
    If rsTmp.RecordCount > 0 Then
        mblnNewTends = False
        strSQL = strSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R3' As 上级id," & _
            "       A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) ||';'||'0' As 参数," & _
            "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
            " From (" & _
            "   Select F.ID, F.编号, F.名称, R.开始, R.截止, R.科室id, 保留" & _
            "   From (" & _
            "       Select ID, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id, 保留" & _
            "          From 病历文件列表 Where 种类 = 3 And 保留 < 0" & _
            "       Union All" & _
            "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, A.科室id, L.保留" & _
            "          From 病历页面格式 F, 病历文件列表 L, 病历应用科室 A" & _
            "          Where L.种类 = 3 And L.保留 = 0 And L.种类 = F.种类 And L.编号 = F.编号 And L.ID = A.文件id(+)" & _
            "       ) F,(" & _
            "       Select R.科室id, Nvl(Min(R.护理级别), 3) As 护理级别, Min(R.发生时间) As 开始, Max(R.发生时间) As 截止" & _
            "          From 病人护理记录 R" & _
            "          Where R.病人来源 = 2 And R.病人id = [1] And Nvl(R.主页id, 0) = [2] And Nvl(R.婴儿, 0) = 0" & _
            "          Group By R.科室id" & _
            "       ) R" & _
            "       Where (F.通用 = 1 Or F.通用 = 2 And F.科室id = R.科室id) And F.护理级别 >= R.护理级别" & _
            "   ) A, 部门表 B Where A.科室id = B.ID "
    Else
        mblnNewTends = True
        strSQL = strSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R3' As 上级id," & vbNewLine & _
                "     A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.婴儿)) As 参数," & vbNewLine & _
                "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.编号, R.名称,R.婴儿, R.开始, NVL(R.截止,nvl(R.时间,R.开始)) 截止, R.科室id, 保留" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, L.保留" & vbNewLine & _
                "          From 病历页面格式 F, 病历文件列表 L" & vbNewLine & _
                "          Where L.种类 = 3 And L.种类 = F.种类 And L.编号 = F.编号 And (L.通用=1 OR L.通用=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.科室id,R.文件名称 名称,R.格式ID,nvl(R.婴儿,0) 婴儿,Min(R.开始时间) As 开始, Max(R.结束时间) As 截止,MAX(T.发生时间) 时间" & vbNewLine & _
                "          From 病人护理文件 R,病人护理数据 T" & vbNewLine & _
                "          Where R.ID=T.文件ID(+) And R.病人id = [1] And Nvl(R.主页id, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.文件名称,R.科室id,R.格式ID,R.婴儿" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.格式ID" & vbNewLine & _
                "   ) A, 部门表 B Where A.科室id = B.ID And DECODE(A.保留,-1,0,A.婴儿)=A.婴儿"
    End If
    
    strSQL = " Select * From (" & strSQL & ") Order By Decode(上级id,Null,' ',上级id),排序"
    
    If blnDataMove And mlng病人ID <> 0 Then
        strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
        strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, IIf(str挂号单 = "", 1, 0))
    Set rsMedRec = zlDatabase.CopyNewRec(rsTmp, False, "", Array("选择", adInteger, 2, Empty))
    'EMR
    Set rsTmp = Nothing
    Set rsTmp = GetEmrCISStruct(lng病人ID, lng主页ID)
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF

	If InStr("," & strRptIDs & ",", "," & rsTmp!ID & ",") = 0 Then
            rsMedRec.AddNew
            rsMedRec!ID = rsTmp!ID
            rsMedRec!上级ID = rsTmp!上级ID
            rsMedRec!名称 = rsTmp!名称
            rsMedRec!参数 = NVL(rsTmp!参数) & ";EMR;" & rsTmp!上级ID
            rsMedRec!图标 = NVL(rsTmp!图标)
            rsMedRec!末级 = 1
            rsMedRec.Update
	     strRptIDs = strRptIDs & "," & rsTmp!ID
            End If

            rsTmp.MoveNext
        Loop
    End If
    '新版PACS
    Set rsTmp = Nothing
    If Not mobjPublicPACS Is Nothing Then
        Set rsTmp = mobjPublicPACS.zlDocGetList(lng病人ID, lng主页ID, str挂号单)
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                rsMedRec.AddNew
                rsMedRec!ID = "R7P" & rsTmp!报告ID
                rsMedRec!上级ID = "R7"
                rsMedRec!名称 = rsTmp!文档标题 & ""
                rsMedRec!参数 = rsTmp!报告ID & ";" & rsTmp!医嘱ID
                rsMedRec!图标 = "object_report"
                rsMedRec!末级 = 1
                rsMedRec.Update
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    '三方LIS报告
    If str挂号单 = "" Then
        strSQL = "select b.id as 报告ID,b.报告名 as 文档标题,c.医嘱ID,b.类型 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and a.病人id=[1] and a.主页id=[2]"
    Else
        strSQL = "select b.id as 报告ID,b.报告名 as 文档标题,c.医嘱ID,b.类型 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and a.挂号单=[3]"
    End If
    If blnDataMove Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
  
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, str挂号单)
    If Not rsTmp Is Nothing Then
    strRptIDs = ""
        Do While Not rsTmp.EOF
	If InStr("," & strRptIDs & ",", "," & rsTmp!报告ID & ",") = 0 Then
            rsMedRec.AddNew
            rsMedRec!ID = "R7L" & rsTmp!报告ID
            rsMedRec!上级ID = "R7"
            rsMedRec!名称 = rsTmp!文档标题 & ""
            rsMedRec!参数 = rsTmp!报告ID & ";" & rsTmp!医嘱ID & ";" & rsTmp!类型 & "<sTab>" & rsTmp!文档标题
            rsMedRec!图标 = "object_report"
            rsMedRec!末级 = 1
            rsMedRec.Update

	     strRptIDs = strRptIDs & "," & rsTmp!报告ID
            End If

            rsTmp.MoveNext
        Loop
    End If
    '追加其他报表
    If str挂号单 = "" And lng病人ID <> 0 Then
        For i = 1 To mcolReport.Count
            rsMedRec.AddNew
            rsMedRec!ID = "R10K" & i
            rsMedRec!上级ID = "R10"
            rsMedRec!名称 = Split(mcolReport(i), ",")(0)
            rsMedRec!参数 = Split(mcolReport(i), ",")(0) & ";" & Split(mcolReport(i), ",")(1) & ";" & Split(mcolReport(i), ",")(2)  '报表名称,系统号,报表编号
            rsMedRec!图标 = "object_report"
            rsMedRec!末级 = 1
            rsMedRec.Update
        Next
    End If
    rsMedRec.Filter = ""
    Set GetCISStruct = rsMedRec
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetPDFPath()
    Dim objFSO As New Scripting.FileSystemObject
    Dim strPath As String
    
    strPath = GetRegister(私有模块, "打印档案", "PDF位置", App.Path)
    strPath = OS.OpenDir(Me.hwnd, "请选择导出PDF文件位置", strPath)
    If strPath = "" Then Exit Sub
    If Not objFSO.FolderExists(strPath) Then
        Call objFSO.CreateFolder(strPath)
    End If
    err.Clear: On Error Resume Next
    Call SetRegister(私有模块, "打印档案", "PDF位置", strPath)
End Sub

Private Sub FuncShowReport(ByVal node As MSComctlLib.node)
    Dim objItem As TabControlItem
    Dim lngIndex As Long
    Dim i As Long
    Dim intPage As Integer
    Dim strCaption As String
    Dim strMsg     As String
    Dim arrPar As Variant
    
    If mlng病人ID = 0 Then Exit Sub
    If Not LoadPrint Then Exit Sub  '加载打印机
    lngIndex = -1
    arrPar = Split(node.Tag, ";")
    strCaption = node.Text
    If Not mobjReportForm Is Nothing Then Unload mobjReportForm
    Set mobjReportForm = Nothing
    If node.Key Like "R11K*" Then
        intPage = Val(Replace(node.Key, "R11K", ""))
        Call PrintInMedRec(Nothing, 5, mlng病人ID, mlng就诊ID, mobjReport, mlng科室ID, Me, intPage, mcolPrint("R11"), mobjReportForm)
    ElseIf node.Key Like "R12K*" Then
        If node.Key = "R12K1" Then
            Call gobjKernel.zlPrintAdvice(Me, mlng病人ID, mlng就诊ID, 0, 1, mcolPrint("R12"), 5, mobjReportForm, strMsg)
        ElseIf node.Key = "R12K2" Then
            Call gobjKernel.zlPrintAdvice(Me, mlng病人ID, mlng就诊ID, 0, 0, mcolPrint("R12"), 5, mobjReportForm, strMsg)
        End If
    ElseIf node.Key Like "R9K*" Then
        strCaption = "住院证"
        Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, "ZLCISBILL" & Format(arrPar(0), "00000") & "-1", "printer", mcolPrint("R9"))  '设置指定打印机
        Call mobjReport.LoadReport(gcnOracle, glngSys, "ZLCISBILL" & Format(arrPar(0), "00000") & "-1", Me, mobjReportForm, Nothing, "NO=" & arrPar(1), "性质=" & arrPar(2), "医嘱ID=0", 1)
    ElseIf node.Key Like "R10K*" Then
        strCaption = "其他报表"
        Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, arrPar(2), "printer", mcolPrint("R10"))
        Call mobjReport.LoadReport(gcnOracle, 0, arrPar(2), Me, mobjReportForm, Nothing, "病人id=" & mlng病人ID, "主页id=" & mlng就诊ID, 1)
    End If
    
    If mobjReportForm Is Nothing Then MsgBox "报表未读取成功，请联系管理员！" & IIf(strMsg <> "", vbCrLf & "提示:" & strMsg, ""), vbInformation, Me.Caption
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive.Item(i).Tag = "报表" Then
            If mobjReportForm Is Nothing Then
                Set objItem = tbcArchive.InsertItem(i, strCaption, picTmp.hwnd, 0)
            Else
                Set objItem = tbcArchive.InsertItem(i, strCaption, mobjReportForm.hwnd, 0)
            End If
            objItem.Tag = "报表"
            Call tbcArchive.RemoveItem(i + 1)
            objItem.Selected = True
            Exit For
        End If
    Next
    Call ShowArchiveTab("报表", strCaption)
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal lngPatiID As Long)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    '直接查找病人
    If chkFilter.Value = vbChecked Then
        If lngPatiID = 0 Then
            MsgBox "找不到符合条件的病人。", vbInformation, gstrSysName
        Else
            Call LoadPatients(lngPatiID)
        End If
        Exit Sub
    End If
    '开始查找行
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            If Val(rptPati.SelectedRows(0).Record(col_病人Id).Value) <> 0 Then blnHave = True
        End If
    End If
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl的索引从是0开始
    Else
        i = rptPati.SelectedRows(0).Index + 1
    End If
    
    '查找病人
    For i = i To rptPati.Rows.Count - 1
        With rptPati.Rows(i)
            If Not .GroupRow Then
                If Val(.Record(col_病人Id).Value) = lngPatiID And lngPatiID <> 0 Then Exit For
                If mstrFindType = "住院号" Then '住院号
                    If .Record(col_住院号).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "门诊号" Then '门诊号
                    If .Record(COL_门诊号).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "挂号单" Then '挂号单
                    If UCase(.Record(COL_NO).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "床号" Then '床号
                    If UCase(Trim(.Record(col_床号).Value)) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "就诊卡" Then '就诊卡
                    If UCase(.Record(col_就诊卡号).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "留观号" Then '留观号
                    If UCase(.Record(col_留观号).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "姓名" Then '姓名
                    If .Record(col_姓名).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                ElseIf mstrFindType = "二代身份证" Then '二代身份证
                    If .Record(col_身份证号).Value = UCase(PatiIdentify.Text) Then Exit For
                End If
            End If
        End With
    Next
    
    If i <= rptPati.Rows.Count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptPati.FocusedRow = rptPati.Rows(i)
        If rptPati.Visible Then rptPati.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
    End If
End Sub

Private Sub FuncViewDisReportCard(ByVal lngEPRid As Long)
    Dim objFrm As New frmChildScale
    '如果传染病报告卡未初始化成功，则重新初始化
    If mobjInfection Is Nothing Then
        Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
        If Not mobjInfection Is Nothing Then
            mobjInfection.Init gcnOracle, glngSys
            
        End If
    End If
    
    If Not mobjInfection Is Nothing Then
        objFrm.zlInitData mobjInfection.zlGetForm
        mobjInfection.zlRefresh mlng病人ID, mlng就诊ID, lngEPRid, mblnMoved
        objFrm.Show 1, Me
    End If
End Sub

Private Function LoadPrint() As Boolean
    Dim varTemp As Variant
    Dim strPrint As String
    Dim i As Long
    
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set mcolPrint = New Collection
    varTemp = Split(M_CON_CATE, ",")
    For i = LBound(varTemp) To UBound(varTemp)
        strPrint = GetRegister(私有模块, "打印档案", "打印机" & varTemp(i), Printer.DeviceName)
        mcolPrint.Add strPrint, varTemp(i)
    Next
    LoadPrint = True
End Function

Private Function DeleteLISTempFile() As Boolean
    Dim objFile As New FileSystemObject
    Dim i As Long
    If mstrTempDel = "" Then Exit Function
    If objFile.FileExists(mstrTempDel) Then
        Do While i < 1000
            On Error Resume Next
            objFile.DeleteFile mstrTempDel, True
            If err.Number = 0 Then
                mstrTempDel = ""
                Exit Do
            End If
            err.Clear: On Error GoTo 0
        Loop
    End If
End Function


