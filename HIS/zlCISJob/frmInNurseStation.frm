VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmInNurseStation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "住院护士工作站"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   Icon            =   "frmInNurseStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleMode       =   0  'User
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   2910
      ScaleHeight     =   960
      ScaleWidth      =   11310
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   255
      Width           =   11310
      Begin VB.Frame fraPageId 
         Height          =   945
         Left            =   15
         TabIndex        =   66
         Top             =   -45
         Width           =   1275
         Begin VB.ComboBox cboPages 
            Height          =   300
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lblPages 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院次数"
            Height          =   180
            Left            =   60
            TabIndex        =   70
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblPatiName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   69
            Top             =   120
            Width           =   435
         End
         Begin VB.Label lblPatiName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "张三"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   68
            Top             =   135
            Width           =   390
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   960
         Left            =   1320
         TabIndex        =   41
         Top             =   -60
         Width           =   9495
         Begin VB.ComboBox cbo过敏 
            Height          =   300
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   390
            Width           =   3130
         End
         Begin VB.Label lblFee 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用:"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   72
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblFee 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   600
            TabIndex        =   71
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "诊断:"
            Height          =   180
            Index           =   0
            Left            =   7060
            TabIndex        =   65
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   7545
            TabIndex        =   64
            Top             =   450
            Width           =   90
         End
         Begin VB.Label lbl过敏 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过敏药物:"
            Height          =   180
            Left            =   75
            TabIndex        =   63
            Top             =   450
            Width           =   810
         End
         Begin VB.Label lbl病室 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   3900
            TabIndex        =   62
            Top             =   165
            Width           =   105
         End
         Begin VB.Label lbl病室 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床位:"
            Height          =   180
            Index           =   0
            Left            =   3465
            TabIndex        =   61
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl类型 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型:"
            Height          =   180
            Index           =   0
            Left            =   4140
            TabIndex        =   60
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lbl类型 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   5580
            TabIndex        =   59
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号:"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   58
            Top             =   150
            Width           =   630
         End
         Begin VB.Label lbl付款 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款:"
            Height          =   180
            Index           =   0
            Left            =   7060
            TabIndex        =   57
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl医保号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医保号:"
            Height          =   180
            Index           =   0
            Left            =   8595
            TabIndex        =   56
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lbl入院 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院:"
            Height          =   180
            Index           =   0
            Left            =   5115
            TabIndex        =   55
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl病况 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病况:"
            Height          =   180
            Index           =   0
            Left            =   4065
            TabIndex        =   54
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl护理 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "护理:"
            Height          =   180
            Index           =   0
            Left            =   2055
            TabIndex        =   53
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl护理 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   2505
            TabIndex        =   52
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl病况 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   4500
            TabIndex        =   51
            Top             =   450
            Width           =   90
         End
         Begin VB.Label lbl入院 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   5580
            TabIndex        =   50
            Top             =   450
            Width           =   90
         End
         Begin VB.Label lbl医保号 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   1
            Left            =   9240
            TabIndex        =   49
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl付款 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   7500
            TabIndex        =   48
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   750
            TabIndex        =   47
            Top             =   165
            Width           =   105
         End
         Begin VB.Label lblFluid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输液量:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   930
            TabIndex        =   46
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblFluid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   1980
            TabIndex        =   45
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblPrint 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医嘱:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   6750
            TabIndex        =   44
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblPrint 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   8010
            TabIndex        =   43
            Top             =   720
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7320
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   37
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox picNotify 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   3615
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   3615
      Begin XtremeReportControl.ReportControl rptNotify 
         Height          =   1515
         Left            =   30
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   330
         Width           =   3510
         _Version        =   589884
         _ExtentX        =   6191
         _ExtentY        =   2672
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.OptionButton optNotify 
         Appearance      =   0  'Flat
         Caption         =   "全病区"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2145
         TabIndex        =   75
         Top             =   90
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optNotify 
         Appearance      =   0  'Flat
         Caption         =   "本人负责"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   74
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label lblNotify 
         AutoSize        =   -1  'True
         Caption         =   "提醒范围:"
         Height          =   180
         Left            =   30
         TabIndex        =   73
         Top             =   90
         Width           =   810
      End
   End
   Begin VB.Timer timNotify 
      Interval        =   500
      Left            =   315
      Top             =   6240
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7485
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInNurseStation.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21114
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
            Key             =   "病人颜色"
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
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   120
      ScaleHeight     =   5145
      ScaleWidth      =   3495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   3495
      Begin VB.Frame fra审查 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   4755
         Visible         =   0   'False
         Width           =   3360
         Begin VB.Label lbl审查 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "共有 XXX 条未处理的病案审查反馈..."
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   450
            MouseIcon       =   "frmInNurseStation.frx":0E1C
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   75
            Width           =   3060
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   105
            Picture         =   "frmInNurseStation.frx":0F6E
            Top             =   45
            Width           =   240
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPati 
         Height          =   2055
         Left            =   1320
         TabIndex        =   35
         Top             =   2400
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   3625
         _StockProps     =   64
      End
      Begin VB.PictureBox picPatiIn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   0
         ScaleHeight     =   4335
         ScaleWidth      =   3525
         TabIndex        =   14
         Top             =   600
         Width           =   3525
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   2415
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   3360
            _Version        =   589884
            _ExtentX        =   5927
            _ExtentY        =   4260
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
            Index           =   4
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   28
            Top             =   1300
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CommandButton cmdRef 
               Caption         =   "刷新"
               Height          =   255
               Left            =   2505
               TabIndex        =   31
               Top             =   0
               Width           =   615
            End
            Begin VB.Frame fraChange 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   750
               TabIndex        =   30
               Top             =   210
               Width           =   300
            End
            Begin VB.TextBox txtChange 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Left            =   780
               MaxLength       =   3
               TabIndex        =   29
               Text            =   "7"
               Top             =   0
               Width           =   285
            End
            Begin VB.Label lbl转出 
               AutoSize        =   -1  'True
               Caption         =   "显示最近    天的转出病人"
               Height          =   180
               Left            =   0
               TabIndex        =   32
               Top             =   30
               Width           =   2160
            End
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   0
            Left            =   90
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CommandButton cmd护理条件 
               Height          =   240
               Left            =   2925
               Picture         =   "frmInNurseStation.frx":14F8
               Style           =   1  'Graphical
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "选择项目(F4)"
               Top             =   30
               Width           =   270
            End
            Begin VB.TextBox txt护理条件 
               Height          =   300
               Left            =   1060
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   0
               Width           =   2160
            End
            Begin VB.Label lbl护理条件 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "护理等级(&N)"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   30
               TabIndex        =   26
               Top             =   60
               Width           =   990
            End
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   1
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   18
            Top             =   340
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CheckBox chk病况条件 
               Caption         =   "重"
               Height          =   195
               Index           =   2
               Left            =   2355
               TabIndex        =   20
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   30
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox chk病况条件 
               Caption         =   "危"
               Height          =   195
               Index           =   1
               Left            =   1785
               TabIndex        =   21
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   30
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox chk病况条件 
               Caption         =   "一般"
               Height          =   195
               Index           =   0
               Left            =   1035
               TabIndex        =   22
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   30
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.Label lbl病况条件 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "当前病况(&S)"
               Height          =   180
               Left            =   0
               TabIndex        =   23
               Top             =   30
               Width           =   990
            End
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   2
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   16
            Top             =   660
            Visible         =   0   'False
            Width           =   3855
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   0
               Width           =   1230
            End
            Begin VB.Label lbl出院时间 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               Height          =   180
               Left            =   0
               TabIndex        =   17
               Top             =   60
               Width           =   720
            End
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   3
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   15
            Top             =   980
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CheckBox chk结清 
               Caption         =   "未结清"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   1845
               TabIndex        =   34
               Top             =   50
               Value           =   1  'Checked
               Width           =   915
            End
            Begin VB.CheckBox chk结清 
               Caption         =   "已结清"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   840
               TabIndex        =   33
               Top             =   50
               Value           =   1  'Checked
               Width           =   915
            End
         End
      End
      Begin VB.PictureBox picPatiFilter 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   45
         ScaleHeight     =   855
         ScaleWidth      =   3390
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   3390
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   1140
            TabIndex        =   2
            Text            =   "cboUnit"
            Top             =   45
            Width           =   2160
         End
         Begin zlIDKind.PatiIdentify PatiIdentify 
            Height          =   270
            Left            =   990
            TabIndex        =   38
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IDKindStr       =   $"frmInNurseStation.frx":15EE
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
         Begin VB.Label lblFind 
            Caption         =   "查找(F3)"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院病区(&U)"
            Height          =   180
            Left            =   105
            TabIndex        =   1
            Top             =   105
            Width           =   990
         End
      End
      Begin MSComctlLib.ImageList imgPati 
         Left            =   195
         Top             =   4635
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":16B0
               Key             =   "Pati"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":1C4A
               Key             =   "Notify"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":21E4
               Key             =   "等待审查"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":277E
               Key             =   "拒绝审查"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":2D18
               Key             =   "正在审查"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":32B2
               Key             =   "正在抽查"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":3CC4
               Key             =   "审查反馈"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":46D6
               Key             =   "抽查反馈"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":4C70
               Key             =   "审查整改"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":5682
               Key             =   "抽查整改"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":6094
               Key             =   "未导入"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":662E
               Key             =   "执行中"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":6BC8
               Key             =   "不符合"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":75DA
               Key             =   "正常结束"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":7B74
               Key             =   "变异结束"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":810E
               Key             =   "Child"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":86A8
               Key             =   "单病种"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":EF0A
               Key             =   "Out"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":F4A4
               Key             =   "紧急"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":FA3E
               Key             =   "Fbaby"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pic护理条件 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   1185
         ScaleHeight     =   1695
         ScaleWidth      =   2115
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   765
         Visible         =   0   'False
         Width           =   2145
         Begin VB.ListBox lst护理条件 
            Appearance      =   0  'Flat
            Height          =   1290
            Left            =   0
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   0
            Width           =   2145
         End
         Begin VB.CommandButton cmdFilterOK 
            Height          =   315
            Left            =   990
            Picture         =   "frmInNurseStation.frx":FFD8
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "确认"
            Top             =   1320
            Width           =   450
         End
         Begin VB.CommandButton cmdFilterCancel 
            Cancel          =   -1  'True
            Height          =   315
            Left            =   1530
            Picture         =   "frmInNurseStation.frx":10562
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "取消"
            Top             =   1320
            Width           =   450
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5265
      Left            =   3705
      TabIndex        =   0
      Top             =   2115
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9287
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmInNurseStation.frx":10AEC
      Left            =   705
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmInNurseStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum PATIREPORT_COLUMN
    col_类型 = 0
    col_审查 = 1
    col_图标 = 2
    col_路径状态 = 3
    col_病人Id = 4
    col_主页ID = 5
    col_姓名 = 6
    col_住院号 = 7
    col_床号 = 8
    col_监护 = 9
    col_性别 = 10
    col_年龄 = 11
    col_护理等级 = 12
    col_费别 = 13
    col_科室 = 14
    col_住院医师 = 15
    col_入院日期 = 16
    col_出院日期 = 17
    col_病人类型 = 18
    col_就诊卡 = 19
    col_科室ID = 20
    col_住院天数 = 21
    col_单病种 = 22
    COL_婴儿科室ID = 23
    COL_婴儿病区ID = 24
    col_西医诊断 = 25
    col_中医诊断 = 26
    col_责任护士 = 27
    col_床位编制 = 28
    col_留观号 = 29
    col_顺序号 = 30
End Enum
Private Enum NOTIFYREPORT_COLUMN
    c_图标 = 0
    C_病人ID = 1
    C_主页ID = 2
    c_姓名 = 3
    c_住院号 = 4
    c_床号 = 5
    C_状态 = 6
    
    '隐藏列
    C_消息 = 7
    C_序号 = 8
    C_日期 = 9
    C_业务 = 10
    C_就诊病区 = 11
End Enum
Private Enum PATI_TYPE
    pt入院待入住 = 0
    pt转科待入住 = 1
    pt转病区待入住 = 2
    pt在院 = 3
    pt预出 = 4
    pt出院 = 5
    pt死亡 = 6
    pt最近转出 = 7
End Enum

Private Type PatiInfo
    状态 As Integer '病案主页.状态  0-正常住院；1-尚未入住；2-正在转科或正在转病区；3-已预出院
    性质 As Integer '0-普通住院病人,1-门诊留观病人,2-住院留观病人
    婴儿 As Integer
    住院号 As String
    床号 As String
    主页ID As Long
    最大主页ID As Long   '病人信息.主页ID
    病区ID As Long
    科室ID As Long  '转科病人，是当前科室ID
    产科 As Boolean
    入院日期 As Date
    出院日期 As Date
    住院次数 As Long
    数据转出 As Boolean
    险类 As Integer
    结清 As Boolean
End Type
Public Enum EFun
    E入住 = 0
    E转科 = 1
    E换床 = 2
    E包房 = 3
    E出院 = 4
    E转为住院 = 5
    E更改床位等级 = 6
    E调整病人信息 = 7
    E新生儿登记 = 8
    E重算费用 = 9
    E医保病种选择 = 10
    E撤销 = 11
    E修改出院时间 = 12
    E床位对换 = 13
    E转医疗小组 = 14
    E转病区 = 15
    E转病区入住 = 16
    E病人备注编辑 = 17
End Enum

Public Enum tbcPatiEnu
    E待入住 = 0
    E在院 = 1
    E出院 = 2
    E转出 = 3
End Enum

'子窗体对象定义
Private mclsEMR As Object  '新版病历zlRichEMR.clsDockEMR
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockInEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zlRichEPR.cDockInTends
Attribute mclsTends.VB_VarHelpID = -1
Private WithEvents mclsFeeQuery As zl9InExse.clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
Private WithEvents mfrmResponse As frmAuditResponse '审查反馈窗口
Attribute mfrmResponse.VB_VarHelpID = -1
Private mclsInPatient As zl9InPatient.clsInPatient
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mclsPath As zlPublicPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private mclsWardMonitor As clsWardMonitor     '监护仪接口
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mobjKernel As zlPublicAdvice.clsPublicAdvice         '临床核心部件
Private mobjSquareCard As Object      '卡结算对象
Private mstrCardKind As String        '卡结算对象返回的可用的医疗卡
Private mblnIsInit As Boolean

Private mcolSubForm As Collection
Private mfrmActive As Form

'54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
Private mclsInOutMedRec As zlMedRecPage.clsInOutMedRec
'参数设置变量
Private mintChange As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintNotify As Integer '医嘱提醒自动刷新间隔(分钟)
Private mintNotifyDay As Integer '提醒多少天内的医嘱
Private mstrNotifyAdvice As String '提醒的医嘱类型
Private mlngMedRedDay As Long     '病案审查反馈天数

'其它窗体变量
Private mstrPrivs As String
Private mlngModul As Long
Private mPatiInfo As PatiInfo '历史住院记录中的,不一定为当前的
Private mlng病人ID As Long, mlng主页ID As Long, mlng病区ID As Long '病人清单中的
Private mblnOutDept As Boolean '是否仅服务于门诊的科室（门诊留观病人显示门诊号）

Private mintFindType As Integer '0-床号,1-住院号,2-就诊卡,3-姓名
Private mstrFindType As String '存储当前查找类型的名称
Private mblnFindTypeEnabled As Boolean
Private mintPreDept As Integer
Private mstrPrePati As String
Private mstrPreNotify As String
Private mintPrePage As Integer
Private mstrUnits As String '操作员所属的病区集
Private mblnNoCheck As Boolean
Private mblnUnRefresh As Boolean
Private mblnEditState As Boolean '子窗体是否处于编辑状态
Private mblnReturn As Boolean        'cboUnit回车按键
Private mintOutPreTime As Integer

Private mblnMonitor As Boolean '监护仪程序是否存在
Private mstrMonitor As String '监护仪程序路径
Private mblnIsFindAgain As Boolean
Private mbytSize As Byte '字体大小 0-小字体（9号字体) 1-大字体（12 号字体）
Private mblnNoRefNotify As Boolean '不刷新医嘱提醒
Private mintMecStandard As Integer  '病案首页格式 0-卫生部标准，1-四川省标准，2-云南省标准
Private mblnTabTmp As Boolean
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln消息语音 As Boolean
Private mblnCardOrder As Boolean '参数：true－床位；false－编制+床位

Private Sub chk病况条件_Click(Index As Integer)
    Dim i As Integer, k As Integer
    
    If Not Visible Or mblnNoCheck Then Exit Sub
    
    If (GetKeyState(vbKeyControl) And &H8000) <> 0 Then
        'Ctrl：排它选择
        mblnNoCheck = True
        For i = 0 To chk病况条件.UBound
            chk病况条件(i).Value = IIf(i = Index, 1, 0)
        Next
        mblnNoCheck = False
    Else
        '至少选择一个
        For i = 0 To chk病况条件.UBound
            If chk病况条件(i).Value = 1 Then k = k + 1
        Next
        If k = 0 Then chk病况条件(Index).Value = 1
    End If
    
    '重新读取病人
    Call LoadPatients
End Sub

Private Sub chk结清_Click(Index As Integer)
    If chk结清(0).Value = 0 And chk结清(1).Value = 0 Then
        chk结清((Index + 1) Mod 2).Value = 1
    End If
    If Me.Visible Then Call LoadPatients
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtOutEnd = datCurr
    mdtOutBegin = mdtOutEnd - 1
    
    cboSelectTime.Clear '出院
    With cboSelectTime
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
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
End Sub

Private Sub cboSelectTime_Click()
'功能：当时间范围是指定是，弹出时间选择窗体
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime) Then
            '取消时恢复原来的选择
            Call Cbo.SetIndex(cboSelectTime.hwnd, mintOutPreTime)
            Exit Sub
        End If
    Else
        mdtOutEnd = datCurr
        mdtOutBegin = mdtOutEnd - intDateCount
    End If
    If mdtOutBegin = CDate(0) Or mdtOutEnd = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "范围：" & Format(mdtOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdtOutEnd, "yyyy-MM-dd")
    End If
    '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
    Call zlDatabase.SetPara("出院病人结束间隔", DateDiff("d", datCurr, mdtOutEnd), glngSys, p住院护士站)
    Call zlDatabase.SetPara("出院病人开始间隔", DateDiff("d", mdtOutBegin, datCurr), glngSys, p住院护士站)
    mintOutPreTime = cboSelectTime.ListIndex
    Call LoadPatients
End Sub

Private Sub cmdRef_Click()
'输入转出天数后刷新
    Call txtChange_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdFilterCancel_Click()
    txt护理条件.SetFocus
    pic护理条件.Visible = False
End Sub

Private Sub cmdFilterCancel_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst护理条件 _
        And Not Me.ActiveControl Is pic护理条件 Then pic护理条件.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer
    
    If lst护理条件.SelCount = 0 Then
        MsgBox "请至少选择一种护理等级。", vbInformation, gstrSysName
        lst护理条件.SetFocus
    End If
    
    If lst护理条件.Selected(0) Then
        txt护理条件.Text = "全部"
        txt护理条件.Tag = ""
    Else
        txt护理条件.Text = ""
        txt护理条件.Tag = ""
        For i = 1 To lst护理条件.ListCount - 1
            If lst护理条件.Selected(i) Then
                txt护理条件.Text = txt护理条件.Text & "," & lst护理条件.List(i)
                txt护理条件.Tag = txt护理条件.Tag & "," & lst护理条件.ItemData(i)
            End If
        Next
        txt护理条件.Text = Mid(txt护理条件.Text, 2)
        txt护理条件.Tag = Mid(txt护理条件.Tag, 2)
    End If
    
    txt护理条件.SetFocus
    pic护理条件.Visible = False
    
    '重新读取病人
    Call LoadPatients
End Sub

Private Sub cmdFilterOK_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst护理条件 _
        And Not Me.ActiveControl Is pic护理条件 Then pic护理条件.Visible = False
End Sub

Private Sub cmd护理条件_Click()
    Dim i As Integer
    
    For i = 0 To lst护理条件.ListCount - 1
        If txt护理条件.Tag = "" Then
            lst护理条件.Selected(i) = True
        ElseIf InStr("," & txt护理条件.Tag & ",", "," & lst护理条件.ItemData(i) & ",") > 0 Then
            lst护理条件.Selected(i) = True
        Else
            lst护理条件.Selected(i) = False
        End If
    Next
    lst护理条件.ListIndex = 0
    pic护理条件.Top = cmd护理条件.Top + cmd护理条件.Height + 30 + picPatiFilter.Top
    pic护理条件.Left = txt护理条件.Left
    pic护理条件.Width = txt护理条件.Width
    pic护理条件.Visible = True
    pic护理条件.ZOrder
    lst护理条件.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '读卡
    PatiIdentify.ActiveFastKey
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
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim blnCol As Boolean, strTmp As String, i As Long, bln路径状态 As Boolean
    Dim intType As Integer
    Dim arrTmp As Variant
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mblnNoCheck = False
    mblnNoRefNotify = False
    
    '读取界面数据
    '-----------------------------------------------------
    mintPreDept = -1
    mstrPrePati = ""
    mstrPreNotify = ""
    mintPrePage = -1
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, p住院护士站, GetInsidePrivs(p住院护士站))
    Call AddMipModule(mclsMipModule)
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    Call GetLocalSetting '本地参数
    
    '界面恢复:默认查找类型读取
    '-----------------------------------------------------
    mintFindType = Val(zlDatabase.GetPara("病人查找方式", glngSys, p住院护士站, , , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    
    mstrMonitor = ""
    mblnMonitor = Dir(App.Path & "\..\gdhs\AC2005.exe") <> ""
    If mblnMonitor Then mstrMonitor = App.Path & "\..\gdhs\AC2005.exe"
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    mstrCardKind = "床|床号|0|0|0|0|0|0;住|住院号|0|0|0|0|0|0;就|就诊卡|0|0|8|0|0|0;姓|姓名|0|0|0|0|0|0;留|留观号|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, IIf(mbytSize = 0, 280, 320), 300, DockLeftOf, Nothing)
    objPane.Title = "住院病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 280, 100, DockBottomOf, objPane)
    objPane.Title = "消息提醒"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p新版住院病历, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    Set mclsAdvices = New zlPublicAdvice.clsDockInAdvices
    Set mclsEPRs = New zlRichEPR.cDockInEPRs
    Set mclsTends = New zlRichEPR.cDockInTends
    Set mclsFeeQuery = New zl9InExse.clsFeeQuery
    Call mclsFeeQuery.InitCallByNurse(gfrmMain, gcnOracle, gstrDBUser, glngSys)
        
    Set mclsInPatient = New zl9InPatient.clsInPatient
    Call mclsInPatient.InitCallByNurse(gfrmMain, gcnOracle, gstrDBUser, glngSys)
    
    Set mclsPath = New zlPublicPath.clsDockPath
    Call mclsAdvices.zlInitPath(mclsPath)
    
    Set mclsWardMonitor = New clsWardMonitor
    
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_新病历"
    End If
    mcolSubForm.Add mclsPath.zlGetForm, "_路径"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_费用"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_病历"
    mcolSubForm.Add mclsTends.zlGetForm, "_护理"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_监护"
    End If
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If GetInsidePrivs(p临床路径应用, True) <> "" Then
            .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0).Tag = "路径": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院医嘱下达, True) <> "" Or GetInsidePrivs(p住院医嘱发送, True) <> "" Then
            .InsertItem(intIdx, "医嘱信息", picTmp.hwnd, 0).Tag = "医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p费用查询, True) <> "" Then
            .InsertItem(intIdx, "费用信息", picTmp.hwnd, 0).Tag = "费用": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院病历管理, True) <> "" Then
            .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0).Tag = "病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p新版住院病历, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0).Tag = "新病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p护理记录管理, True) <> "" Then
            .InsertItem(intIdx, "护理信息", picTmp.hwnd, 0).Tag = "护理": intIdx = intIdx + 1
        End If
        If mclsWardMonitor.Enabled Then
            If InStr(GetInsidePrivs(p住院护士站), "护理监护") > 0 Then
                .InsertItem(intIdx, "护理监护", picTmp.hwnd, 0).Tag = "监护": intIdx = intIdx + 1
            End If
        End If
        
        '外挂部件中的卡片
        Call CreatePlugInOK(p住院护士站)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p住院护士站)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p住院护士站, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "你没有使用住院护士工作站的权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '恢复上次选择的卡片
        strTab = zlDatabase.GetPara("医护功能", glngSys, p住院护士站)
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '避免激活事件
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
        End If
        '只加载选择的子窗体
        Call tbcSub_SelectedChanged(.Selected)
    End With
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
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "待入住", picPatiIn.hwnd, 0).Tag = "待入住"
        .InsertItem(1, "在院", picPatiIn.hwnd, 0).Tag = "在院"
        .InsertItem(2, "出院", picPatiIn.hwnd, 0).Tag = "出院"
        .InsertItem(3, "转出", picPatiIn.hwnd, 0).Tag = "转出"
        
        .Item(3).Selected = True
        .Item(1).Selected = True
        '定位病人选项卡
        tbcPati.Item(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", 1)).Selected = True
    End With
    
    '其它界面设置
    Call InitReportColumn
    picPatiFilter.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    chk病况条件(0).BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    chk病况条件(1).BackColor = chk病况条件(0).BackColor
    chk病况条件(2).BackColor = chk病况条件(0).BackColor
    picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picInfo.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    Call Cbo.SetListWidth(cbo过敏.hwnd, cbo过敏.Width * 2)
    Call Cbo.SetListWidth(cboPages.hwnd, cboPages.Width * 1.5)
    
    '初始化病人过滤条件
    strTmp = zlDatabase.GetPara("当前病况过滤", glngSys, p住院护士站, "111", _
        Array(lbl病况条件, chk病况条件(0), chk病况条件(1), chk病况条件(2)), InStr(mstrPrivs, "参数设置") > 0)
    For i = 0 To chk病况条件.UBound
        chk病况条件(i).Value = IIf(Mid(strTmp, i + 1, 1) = "1", 1, 0)
    Next
    If Not InitNurselevel Then Unload Me: Exit Sub
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        If InStr(mstrPrivs, "全院病人") > 0 Then
            MsgBox "没有发现住院病区信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Else
            MsgBox "没有发现你所属病区,不能使用住院护士工作站！", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If
    Call cboUnit_Click
    Call HaveRIS(True)
    
    '转出病人天数
    txtChange.Text = mintChange
    mintOutPreTime = -1
    Call InitSelectTime
    
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    End If
    If Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "NotifyRang", "0")) = 0 Then
        optNotify(1).Value = True
    Else
        optNotify(0).Value = True
    End If
    blnCol = rptPati.Columns(col_审查).Visible
    bln路径状态 = rptPati.Columns(col_路径状态).Visible
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    rptPati.Columns(col_审查).Visible = blnCol
    
    rptPati.Columns(col_路径状态).Visible = bln路径状态
    If bln路径状态 And rptPati.Columns(col_路径状态).Width = 0 Then rptPati.Columns(col_路径状态).Width = 18
    
    If tbcSub.Selected.Tag = "医嘱" Then '以防特殊情况出现
        If dkpMain.Panes(2).Closed Then dkpMain.Panes(2).Closed = False
    End If
    Call LoadNotify
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long, strPrivs As String, strTmp As String
    Dim byt入住方式 As Byte, blnRefresh As Boolean
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
    Select Case Control.ID
    '管理菜单，病人入出转
    Case conMenu_Manage_Change_In
        If CheckBabyInOut Then Exit Sub
        If rptPati.SelectedRows(0).Record(col_类型).Value = pt转病区待入住 Then
            blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转病区入住, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID, "", 0)
        Else
            If rptPati.SelectedRows(0).Record(col_类型).Value = pt转科待入住 Then
                byt入住方式 = 1 '0-入院入住，1-转科入住
            Else
                byt入住方式 = 0
            End If
            blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E入住, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID, "", _
                Val(rptPati.SelectedRows(0).Record(col_科室ID).Value), byt入住方式)
        End If
    Case conMenu_Manage_Change_Turn
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转科, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID)
        
    Case conMenu_Manage_Change_TurnUnit
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转病区, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID)
        
    Case conMenu_Manage_Change_TurnTeam
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转医疗小组, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID)
        
    Case conMenu_Manage_Change_Bed
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E换床, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID, 0, "", "")
        
    Case conMenu_Manage_Change_TransposeBed
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E床位对换, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID, mPatiInfo.床号, "")
        
        
    Case conMenu_Manage_Change_House    '目前包房调用换床接口
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E换床, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID, 1, "", "")
        
    Case conMenu_Manage_Change_Out
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E出院, Me, strPrivs, mlng病人ID, mlng主页ID)
        
    Case conMenu_Manage_Change_InPati
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转为住院, Me, strPrivs, mlng病人ID, mlng主页ID, _
                Val(rptPati.SelectedRows(0).Record(col_住院号).Value), CStr(rptPati.SelectedRows(0).Record(col_姓名).Value))
        
        
    Case conMenu_Manage_Change_BedGrid
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E更改床位等级, Me, strPrivs, mlng病人ID, mlng主页ID, _
            Trim(CStr(rptPati.SelectedRows(0).Record(col_床号).Value)))
    Case conMenu_Manage_Change_PatiInfo
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E调整病人信息, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID)
       
    Case conMenu_Manage_Change_PaitNote
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E病人备注编辑, Me, strPrivs, mlng病人ID, mlng主页ID)
        
    Case conMenu_Manage_Change_Baby
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E新生儿登记, Me, strPrivs, mlng病人ID, mlng主页ID)
    Case conMenu_Manage_Change_ReCalcFee
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E重算费用, Me, strPrivs, mlng病人ID, mlng主页ID, _
        CStr(rptPati.SelectedRows(0).Record(col_姓名).Value))
    Case conMenu_Manage_Change_InsureSel
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E医保病种选择, Me, strPrivs, mlng病人ID, mlng主页ID, Val(mPatiInfo.险类))
        
    Case conMenu_Manage_Change_Undo * 10 + 1    '回退
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E撤销, Me, strPrivs, mlng病区ID, mlng病人ID, mlng主页ID, Val(mPatiInfo.险类), Control.Caption)
        
    Case conMenu_Manage_Monitor '监护仪
        Call ExecuteMonitor
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                    objControl.Style = xtpButtonIcon
                Else
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S '小字体
        If mbytSize <> 0 Then
            mbytSize = 0
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '大字体
        If mbytSize <> 1 Then
            mbytSize = 1
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_Tool_Archive '电子病案查阅
        mblnUnRefresh = True
        Call frmArchiveView.ShowArchive(Me, mlng病人ID, mPatiInfo.主页ID)
        mblnUnRefresh = False
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        mblnUnRefresh = True
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        mblnUnRefresh = True
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Manage_FeeItemSet  '诊疗项目费用设置
        Call Set诊疗项目费用设置
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    Case conMenu_Tool_MedRec '首页整理
        mblnUnRefresh = True
        Call ExecuteEditMediRec
        mblnUnRefresh = False
    Case conMenu_File_MedRecSetup '首页打印设置
        Call PrintInMedRec(mclsInOutMedRec, 0, mlng病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case conMenu_File_MedRecPreview '首页预览
        Call PrintInMedRec(mclsInOutMedRec, 1, mlng病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case conMenu_File_MedRecPrint '首页打印
        Call PrintInMedRec(mclsInOutMedRec, 2, mlng病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case conMenu_Manage_Print_Label '腕带打印
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, 2)
        End If
    Case conMenu_Tool_MedRecAuditResponse '审查反馈
        '都可以调用，至少可以查看(当前或历史)
        Call lbl审查_Click

    Case conMenu_View_Find '查找
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '有时需要定位一下
            If PatiIdentify.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            PatiIdentify.SetFocus
        End If
    Case conMenu_View_FindNext '查找下一个
        If PatiIdentify.Text = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True)
        End If
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                rptPati.SelectedRows(0).Expanded = False
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    rptPati.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '因折叠定位到分组上,不会自动激活该事件
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_CurExpend '展开当前组
        If rptPati.SelectedRows.Count > 0 Then
            rptPati.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '折叠所有组
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '因折叠定位到分组上,不会自动激活该事件
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_AllExpend '展开所有组
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_Notify And tbcSub.Selected.Tag <> "费用" '医嘱提醒(和一日清单的ID相同了)
        If rptNotify.Visible Then Call LoadNotify
    Case conMenu_View_Refresh '刷新
        blnRefresh = True
         
    Case conMenu_File_Parameter '参数设置
        mblnUnRefresh = True
        frmInStationSetup.mbln护士站 = True
        frmInStationSetup.mstrPrivs = mstrPrivs
        frmInStationSetup.Show 1, Me
        If gblnOK Then
            Call GetLocalSetting
            blnRefresh = True
        End If
        mblnUnRefresh = False
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    Case Else
        mblnUnRefresh = True
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            With mPatiInfo
                strTmp = Split(Control.Parameter, ",")(1)
                If strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then '住院科室日报
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                             "病区=" & mlng病区ID, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID)
                ElseIf strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Or strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then    '病人帐页和催款表
                    Call mclsFeeQuery.zlExecuteCommandBars(Control)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                        "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "住院号=" & .住院号, "病人病区=" & .病区ID, _
                        "病人科室=" & .科室ID, "床号=" & .床号)
                End If
            End With
        ElseIf Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 1, conMenu_File_MedRecPreview * 100# + 4) Then
            Call PrintInMedRec(mclsInOutMedRec, IIf(Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6), 2, 1), mlng病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me, Val(Mid(Control.ID & "", Len(Control.ID & ""))))
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "路径"
                If CheckBabyInOut Then Exit Sub
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "医嘱"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "费用"
                Call mclsFeeQuery.zlExecuteCommandBars(Control)
            Case "病历"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "护理"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "新病历"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p住院护士站, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng病人ID, mlng主页ID, "")
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        mblnUnRefresh = False
    End Select
    
    If blnRefresh Then Call LoadPatients
    
    If Control.ID = conMenu_View_Refresh Then Call LoadNotify '刷新医嘱提醒

End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim rsPatiLog As ADODB.Recordset
    Dim i As Long, j As Long, strPrivs As String
    Dim objControl As CommandBarControl
    
    If CommandBar.Parent Is Nothing Then Exit Sub
        
    'Call CommandBar.Controls.DeleteAll
        
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "床  号(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "住院号(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "就诊卡(&3)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 4, "姓  名(&4)"
            End If
        End With
    Case conMenu_File_MedRecPrint
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 1, "正面(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 2, "反面(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 3, "附页1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 4, "附页2(&4)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 5, "正面+附页1(&5)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 6, "反面+附页2(&6)"
            End If
        End With
    Case conMenu_File_MedRecPreview
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 1, "正面(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 2, "反面(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 3, "附页1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 4, "附页2(&4)"
            End If
        End With
    Case conMenu_Manage_Change_Undo
        With CommandBar.Controls
            .DeleteAll
            If mlng病人ID = 0 Then Exit Sub
            
            Set rsPatiLog = GetPatiLog(mlng病人ID, mlng主页ID)
            If rsPatiLog.RecordCount > 0 Then '动态子菜单,扩1位
                
                strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
                rsPatiLog.MoveFirst
                For i = 1 To rsPatiLog.RecordCount
                    If Not IsNull(rsPatiLog!终止时间) And rsPatiLog!终止原因 = 1 Then
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, "出院")
                        j = j + 1
                        If InStr(";" & strPrivs & ";", ";撤消出院;") = 0 Or j > 1 Then objControl.Enabled = False
                    Else
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, rsPatiLog!操作)
                        If rsPatiLog.RecordCount > 1 And rsPatiLog!开始原因 = 1 Then objControl.Visible = False
                        j = j + 1
                        If j > 1 Then
                            objControl.Enabled = False
                        Else
                            If (objControl.Caption Like "*入住" Or objControl.Caption = "转病区入住") Then
                                If InStr(strPrivs, "撤消入科") = 0 Then objControl.Enabled = False
                            End If
                            If objControl.Caption = "转为住院病人" Then
                                If InStr(strPrivs, "住院留观转住院") = 0 Then objControl.Enabled = False
                            ElseIf objControl.Caption = "预出院" Then
                                If InStr(strPrivs, "撤销预出院") = 0 Then objControl.Enabled = False
                                
                            ElseIf objControl.Caption = "换床" Then
                                If InStr(strPrivs, "换床") = 0 Then objControl.Enabled = False
                            End If
                        End If
                    End If
                    objControl.Category = "撤销"
                    If i <> 1 Then objControl.Enabled = False
                    rsPatiLog.MoveNext
                Next
            End If
        End With
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "路径"
            Call mclsPath.zlPopupCommandBars(CommandBar)
       Case "医嘱"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "费用"
           Call mclsFeeQuery.zlPopupCommandBars(CommandBar)
       Case "病历"
       
       Case "护理"
       
       End Select
    End Select
    
End Sub


Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置病人相关的菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strPrivs As String

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Parameter = "已判断" Then Exit Sub

    blnVisible = True
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
    
    Select Case Control.ID
        Case conMenu_Manage_Change_In
            blnVisible = strPrivs <> ""
        Case conMenu_Manage_Change_Out
            blnVisible = InStr(strPrivs, "病人出院") > 0
        Case conMenu_Manage_Change_Turn
            blnVisible = InStr(strPrivs, "病人转科") > 0
        Case conMenu_Manage_Change_Bed, conMenu_Manage_Change_TransposeBed, conMenu_Manage_Change_House
            blnVisible = InStr(strPrivs, "换床") > 0
        Case conMenu_Manage_Change_TurnUnit
            blnVisible = InStr(strPrivs, "转病区") > 0
        Case conMenu_Manage_Change_PatiInfo
            blnVisible = InStr(strPrivs, "调整病人信息") > 0
        Case conMenu_Manage_Change_Baby
            blnVisible = InStr(strPrivs, "新生儿登记") > 0
        Case conMenu_Manage_Change_ReCalcFee
            blnVisible = InStr(strPrivs, "重算费用") > 0
        Case conMenu_Manage_Change_BedGrid
            blnVisible = InStr(strPrivs, "调整床位等级") > 0
        Case conMenu_Manage_Change_InPati
            blnVisible = InStr(strPrivs, "住院留观转住院") > 0
    End Select

    Control.Visible = blnVisible
    Control.Parameter = "已判断"
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnSelect As Boolean, blnWaitIn As Boolean, blnOutTo As Boolean
    Dim blnOut As Boolean, blnPreOut As Boolean, lngType As Long, strPrivs As String
    Dim blnWriteMedRec As Boolean
    Dim i As Long
        
    If Not mblnIsInit Then
        mblnIsInit = True
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
            End If
            Call PatiIdentify.zlInit(Me, glngSys, p住院护士站, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
            PatiIdentify.objIDKind.AllowAutoICCard = True
            PatiIdentify.objIDKind.AllowAutoIDCard = True
            If Not PatiIdentify.objIDKind.Cards Is Nothing Then
            For i = 0 To PatiIdentify.objIDKind.Cards.Count - 1
                If i = mintFindType Then
                    PatiIdentify.objIDKind.IDKind = i + 1
                    mstrFindType = PatiIdentify.objIDKind.Cards(i + 1).名称
                    Exit For
                End If
            Next
            End If
        End If
    End If
    
    If Control.Category = "撤销" Then Exit Sub '在cbsMain_InitCommandsPopup已设置,退出避免子窗体设置其可见性
    
    If mblnEditState Then
        Select Case Control.ID
        Case conMenu_Help_Help, conMenu_File_Exit
            '不处理
        Case Else
            If Control.DescriptionText = "" Then Control.DescriptionText = IIf(Control.Enabled = True, 1, 0)
            Control.Enabled = False
            Exit Sub
        End Select
    Else
        If Control.DescriptionText <> "" Then
            Control.Enabled = IIf(Val(Control.DescriptionText) = 1, True, False)
                Control.DescriptionText = ""
        End If
    End If
    '其他情况下，Control.Enabled=false时，还要在业务部件中判断，所以不能退出
        
    If rptPati.SelectedRows.Count > 0 Then blnSelect = Not rptPati.SelectedRows(0).GroupRow
    If blnSelect Then
        lngType = Val(rptPati.SelectedRows(0).Record(col_类型).Value)
        blnWaitIn = lngType = pt转科待入住 Or lngType = pt入院待入住 Or lngType = pt转病区待入住
        blnOut = lngType = pt出院
        blnPreOut = lngType = pt预出
        blnOutTo = lngType = pt最近转出
    End If
    
    If Control.Category = "病人" Then
        Call SetControlVisible(Control)
        If Not Control.Visible Then Exit Sub
        
        strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
        If InStr(strPrivs, "所有病区") = 0 Then
            If InStr("," & mstrUnits & ",", "," & mlng病区ID & ",") = 0 Then Control.Enabled = False: Exit Sub
        End If
        If blnSelect = False Then Control.Enabled = False: Exit Sub
    End If
            
    Select Case Control.ID
    Case conMenu_Manage_Change_Undo
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOutTo And mlng主页ID = mPatiInfo.最大主页ID
    Case conMenu_Manage_FeeItemSet  '诊疗项目费用设置,没有权限时可查看
        
    Case conMenu_Manage_Change_In   '入住
        Control.Enabled = blnWaitIn
        
    Case conMenu_Manage_Change_InPati   '转为住院
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.性质 = 2
        End If
        
    '转科，换床，包房，调整病人信息，重算费用,转病区，转小组,床位对换
    Case conMenu_Manage_Change_Turn, conMenu_Manage_Change_Bed, conMenu_Manage_Change_House, _
         conMenu_Manage_Change_PatiInfo, conMenu_Manage_Change_ReCalcFee, conMenu_Manage_Change_TurnUnit, _
         conMenu_Manage_Change_TurnTeam, conMenu_Manage_Change_TransposeBed
         
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.状态 <> 2
            
            If Control.ID = conMenu_Manage_Change_TransposeBed Then '床位对换
                Control.Enabled = Trim(CStr(rptPati.SelectedRows(0).Record(col_床号).Value)) <> ""
            End If
        End If
    Case conMenu_Manage_Change_InsureSel
        Control.Enabled = Not blnWaitIn And Not blnPreOut
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.险类 <> 0
        End If
    Case conMenu_Manage_Change_BedGrid
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = Trim(CStr(rptPati.SelectedRows(0).Record(col_床号).Value)) <> "" And mPatiInfo.状态 <> 2
        End If
    Case conMenu_Manage_Change_Out
        Control.Enabled = Not blnWaitIn And Not blnOut
        If Control.Enabled Then
            Control.Enabled = (rptPati.SelectedRows(0).Record(col_类型).Value = pt在院 Or blnPreOut) And mPatiInfo.状态 <> 2
        End If
    Case conMenu_Manage_Change_Baby
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.产科 And rptPati.SelectedRows(0).Record(col_性别).Value = "女"
        End If
    Case conMenu_Manage_Monitor '监护仪
        Control.Visible = mblnMonitor
        
    Case conMenu_Manage_Change_PaitNote
        Control.Enabled = Not blnOutTo
        
        
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_FontSize_S '小字体
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_View_FontSize_L '大字体
        Control.Checked = (mbytSize = 1)
    Case conMenu_View_Expend_CurExpend '展开当前组
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPati.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = rptPati.SelectedRows(0).Expanded
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '折叠/展开组
        Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
    Case conMenu_Tool_Archive '电子病案查阅
        If GetInsidePrivs(p电子病案查阅) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        If GetInsidePrivs(p疾病诊断参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 '药品及诊疗参考
        If GetInsidePrivs(p药品诊疗参考) = "" Then Control.Visible = False
    Case conMenu_Tool_MedRecAuditResponse '审查反馈
        '都可以调用，至少可以查看(当前或历史)
        Control.Enabled = rptPati.Rows.Count > 0
    Case conMenu_View_Notify And tbcSub.Selected.Tag <> "费用" '医嘱提醒
        Control.Enabled = rptNotify.Visible
    Case conMenu_File_MedRec '首页打印
        If InStr(mstrPrivs, "打印首页") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = cboPages.ListIndex <> -1
        End If
    Case conMenu_File_MedRecPreview, conMenu_File_MedRecPrint
        If InStr(mstrPrivs, "打印首页") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = cboPages.ListIndex <> -1
        End If
    Case conMenu_Manage_Print_Label '腕带打印
        If InStr(mstrPrivs, ";腕带打印;") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = Not blnOut And blnSelect
        End If
         
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    Case conMenu_Tool_MedRec '首页整理
        blnWriteMedRec = Val(zlDatabase.GetPara("医生和护士分别填写病案首页", glngSys, p住院医生站, "0")) = 1
        Control.Enabled = cboPages.ListIndex <> -1 And mPatiInfo.主页ID > 0 And blnWriteMedRec = True
        Control.Visible = blnWriteMedRec
    Case conMenu_File_Parameter '参数设置
        'If InStr(mstrPrivs, "参数设置") = 0 Then Control.Visible = False
    Case Else
        '60075:刘鹏飞,2013-04-03,将外部对医嘱打印、预览菜单的处理，移植到此处,以前的方式导致无法调用虚拟模块的更新事件
        If (Control.ID = conMenu_File_Print Or Control.ID = conMenu_File_Preview Or Control.ID = conMenu_Help_Help) Then
            If tbcSub.Selected.Tag = "医嘱" Then
                Control.Visible = False
                Exit Sub
            Else
                Control.Visible = True
            End If
        End If
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then
                Control.Visible = tbcSub.Selected.Tag = "费用"  '催款表
                Exit Sub
            End If
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then
                Control.Visible = tbcSub.Selected.Tag = "费用"  '病人帐页
                Exit Sub
            End If
        End If
        '首页报表
        If Between(Control.ID, conMenu_File_MedRecPrint * 100# + 3, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 3, conMenu_File_MedRecPreview * 100# + 4) Then
            If mintMecStandard = 0 Or mintMecStandard = 3  Or mintMecStandard = 1 Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            Exit Sub
        End If
    
        Select Case tbcSub.Selected.Tag
        Case "路径"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "医嘱"
            Call mclsAdvices.zlUpdateCommandBars(Control)
         Case "费用"
            Call mclsFeeQuery.zlUpdateCommandBars(Control)
        Case "病历"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "护理"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "新病历"
            Call mclsEMR.zlUpdateCommandBars(Control)
        End Select
    End Select
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    
    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If
    
    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hwnd)
        
    Me.Caption = "住院护士工作站 - " & objItem.Caption & "(当前用户：" & UserInfo.姓名 & ")"
    
    If mstrNotifyAdvice = "000000000000" Then
        dkpMain.Panes(2).Tag = IIf(dkpMain.Panes(2).Hidden, 1, 0)
        dkpMain.Panes(2).Close
    Else
        dkpMain.Panes(2).Closed = False
        dkpMain.Panes(2).Hidden = Val(dkpMain.Panes(2).Tag) = 1
        dkpMain.Panes(2).Title = "消息提醒"
    End If
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '主窗口重新加入
    Call MainDefCommandBar
    
    '子窗口重新加入
    Select Case objItem.Tag
    Case "路径"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "医嘱"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "费用"
        Call mclsFeeQuery.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "病历"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "护理"
        Call mclsTends.zlDefCommandBars(Me.cbsMain)
    Case "新病历"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p住院护士站, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(err, "GetButtomName")
            '构建菜单
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            err.Clear: On Error GoTo 0
        End If
    End Select
    
    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                objControl.Style = xtpButtonIcon
            Else
                objControl.Style = bytStyle
            End If
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next
    
    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'功能：刷新子窗体数据及状态
    Dim blnEdit As Boolean, strInPatiNO As String, lng路径状态 As Long
    Dim lngType As PATI_TYPE, lng病区ID As Long, lng科室ID As Long
    Dim lngState As TYPE_PATI_State
    
    If mlng病人ID = 0 Then
        '要求子窗体按无数据处理界面
        Select Case objItem.Tag
        Case "路径"
            Call mclsPath.zlRefresh(0, 0, 0, 0, 0, False)
        Case "医嘱"
            Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "费用"
            Call mclsFeeQuery.zlRefresh(0, 0, 0, 0, 0, False, False, False)
        Case "病历"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "护理"
            Call mclsTends.zlRefresh(0, 0, 0, False, False)
        Case "监护"
            Call mclsWardMonitor.HideWindow
        Case "新病历"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 3)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p住院护士站, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            lngType = Val(rptPati.SelectedRows(0).Record(col_类型).Value)
            lng科室ID = Val("" & rptPati.SelectedRows(0).Record(col_科室ID).Value)    '最近转出病人为原科室ID
            
            If InStr("," & pt入院待入住 & "," & pt最近转出 & "," & pt转科待入住 & "," & pt转病区待入住 & ",", "," & lngType & ",") > 0 Then
                '待入住病人，转出病人，传当前界面的病区
                lng病区ID = mlng病区ID
            Else
                lng病区ID = .病区ID
            End If
            
            If lngType = pt最近转出 Then
                lngState = ps最近转出
            ElseIf lngType = pt转科待入住 Or lngType = pt转病区待入住 Then
                lngState = ps待转入
            Else
                lngState = IIf(.出院日期 = CDate(0), IIf(.状态 = 3, ps预出, ps在院), ps出院)
            End If
                        
        
            Select Case objItem.Tag
            Case "路径"
                Call mclsPath.zlRefresh(mlng病人ID, .主页ID, lng病区ID, lng科室ID, .状态, .数据转出, True, , , mclsMipModule)
            Case "医嘱"
                lng路径状态 = Val(rptPati.SelectedRows(0).Record(col_路径状态).Value)
                If .状态 = 1 Then '入院待入住
                   If Val(zlDatabase.GetPara("允许给待入住病人下达医嘱", glngSys, p住院医嘱下达, 1)) = 0 Then
                        lngState = ps待转入 'lngState=ps待转入时新开医嘱等功能不可用
                   End If
                End If
                Call mclsAdvices.zlRefresh(mlng病人ID, .主页ID, lng病区ID, lng科室ID, lngState, .数据转出, , , , lng路径状态, mlng病区ID, mclsMipModule, .婴儿)
                
            Case "费用"
                Call mclsFeeQuery.zlRefresh(mlng病人ID, mlng主页ID, Val(.住院号), lng病区ID, .险类, .数据转出, .出院日期 <> CDate("0:00:00"), .结清, False, _
                    lngType = pt最近转出 Or lngType = pt预出 Or lngType = pt出院, lng科室ID)
               
            Case "病历"
                Call mclsEPRs.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, False, .数据转出, 0, tbcSub.Tag = "路径", lng病区ID, lngState)
            Case "护理"
                blnEdit = True
                With rptPati.SelectedRows(0)
                    If lngType = pt出院 Or lngType = pt死亡 Then
                        If Not (.Record(col_审查).Value = 0 Or .Record(col_审查).Value = 2 Or .Record(col_审查).Value = 999) Then
                            '可能是在院抽查反馈状态，出院后并未提交审查
                            If .Record(col_图标).Value = 1 Then blnEdit = False
                        End If
                    ElseIf lngType = pt转科待入住 Or lngType = pt转病区待入住 Then
                        blnEdit = False
                    End If
                End With
                blnEdit = blnEdit And (mlng病区ID = .病区ID Or lngType = pt最近转出)
                Call mclsTends.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, blnEdit, False, lng病区ID, lngState)
            Case "监护"
                strInPatiNO = Trim(rptPati.SelectedRows(0).Record(col_住院号).Value)
                If strInPatiNO = "" Then
                    Call mclsWardMonitor.HideWindow
                Else
                    Call mclsWardMonitor.ShowInfor(strInPatiNO)
                End If
            Case "新病历"
                Call mclsEMR.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, lngState, 3)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p住院护士站, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng病人ID, "", .主页ID, .数据转出, , , _
                                    lng病区ID, lng科室ID, , lngState, , lng路径状态)
                    Call zlPlugInErrH(err, "RefreshForm")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End With
    End If
    Call SetFontSize(Not Me.Visible)
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_File_MedRec, "首页打印(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_File_MedRecSetup, "打印设置(&S)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "打印预览(&V)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "打印首页(&P)", -1, False
        End With
        '49854:刘鹏飞,2013-10-31,病人腕带打印
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print_Label, "打印腕带(&W)…")  '打印腕带
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "病人(&P)", -1, False) '固有
    objMenu.ID = conMenu_ManagePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "入住(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Turn, "转科(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnUnit, "转病区(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnTeam, "转小组(&T)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Bed, "换床(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TransposeBed, "床位对换(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_House, "包房(&H)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_BedGrid, "更改床位等级(&G)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PatiInfo, "调整住院信息(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PaitNote, "病人备注信息(&F)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Out, "出院(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InPati, "转为住院病人(&Z)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Baby, "新生儿登记(&N)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_ReCalcFee, "按费别重算费用(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "医保病种选择(&M)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Change_Undo, "撤销(&U)"): objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Edit_Untread
                        
        '监护仪
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Monitor, "监护仪(&N)")
        objControl.BeginGroup = True
    End With
    For Each objControl In objMenu.CommandBar.Controls
        objControl.Category = "病人"
    Next


    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "字体大小(&N)") '固有
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_FontSize_S, "小字体(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_FontSize_L, "大字体(&L)", -1, False '固有
        End With

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "查找方式(&Y)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
        
        '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "首页整理(&M)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "审查反馈(&S)")
            objControl.IconId = 3814
            objControl.BeginGroup = True
            objControl.ToolTipText = "处理或查看病案审查反馈"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "诊疗项目费用设置(&C)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With


    '工具栏定义
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop) '固有
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
        
        '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "首页"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "预览首页")
        objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "打印首页")
        objControl.IconId = conMenu_File_Print
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Monitor, "监护仪"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找病人
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add 0, vbKeyF12, conMenu_File_Parameter '参数设置
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF6, conMenu_View_Jump '跳转
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
    '-----------------------------------------------------
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '打印设置
'        .AddHiddenCommand conMenu_File_Excel '输出到Excel
'        .AddHiddenCommand conMenu_View_Jump '跳转
'    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8", _
            "ZL1_INSIDE_1261_9", "ZL1_INSIDE_1261_10")
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub lbl审查_Click()
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    '非模态显示审查反馈窗体
    If mfrmResponse Is Nothing Then
        Set mfrmResponse = New frmAuditResponse
    End If
    mblnUnRefresh = True
    Call mfrmResponse.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex), 1, False, 1, mstrPrivs)
    mblnUnRefresh = False
End Sub

Private Sub lst护理条件_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lst护理条件.ListCount - 1
            lst护理条件.Selected(i) = lst护理条件.Selected(0)
        Next
    ElseIf Not lst护理条件.Selected(Item) Then
        lst护理条件.Selected(0) = False
    ElseIf lst护理条件.SelCount = lst护理条件.ListCount - 1 Then
        lst护理条件.Selected(0) = True
    End If
End Sub

Private Sub lst护理条件_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst护理条件 _
        And Not Me.ActiveControl Is pic护理条件 Then pic护理条件.Visible = False
End Sub

Private Sub mclsAdvices_SetEditState(ByVal blnEditState As Boolean)
'功能：子窗体处于编辑状态时，禁止使用菜单及转移焦点的所有功能
    picPati.Enabled = Not blnEditState
    picInfo.Enabled = Not blnEditState
    picNotify.Enabled = Not blnEditState
    
    mblnEditState = blnEditState
    mblnUnRefresh = blnEditState
End Sub

Private Sub mclsAdvices_ExecLogModi(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, ByVal 执行时间 As String, 完成 As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    完成 = frmTechnicLog.ShowMe(Me, p住院医嘱发送, 科室ID, 医嘱ID, 发送号, False, 执行时间)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_ExecLogNew(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, 完成 As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    完成 = frmTechnicLog.ShowMe(Me, p住院医嘱发送, 科室ID, 医嘱ID, 发送号, False)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'功能：医嘱子窗体要求刷新
    If Not RefreshNotify Then
        Call LoadPatients
    ElseIf rptNotify.Visible Then
        '仅刷新医嘱提醒区域
        Call LoadNotify
    End If
End Sub

Private Sub mclsMipModule_OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
'功能：点击冒泡消息后定位病人
    Dim int列表 As Integer
    
    int列表 = -1
    If InStr(",ZLHIS_PATIENT_002,ZLHIS_PATIENT_012,ZLHIS_PATIENT_009,ZLHIS_PATIENT_006,ZLHIS_PATIENT_010,", "," & strMsgKey & ",") > 0 Then
        int列表 = E在院
    ElseIf InStr(",ZLHIS_PATIENT_003,", "," & strMsgKey & ",") > 0 Then
        int列表 = E待入住
    End If
    
    If int列表 <> -1 Then
        If tbcPati.Item(int列表).Selected = False Then
            tbcPati.Item(int列表).Selected = True '选项卡切换时会刷新病人列表
            Call LocatePati(strLinkPara)
        Else
            If Not LocatePati(strLinkPara) Then
                Call LoadPatients
                Call LocatePati(strLinkPara)
            End If
        End If
    End If
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'功能：消息接收
    Dim blnRecToLis As Boolean '是否加载到提醒列表中
    Dim rsMsg As ADODB.Recordset
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    If strMsgItemIdentity = "ZLHIS_TRANSFUSION_001" And Mid(mstrNotifyAdvice, 6, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_001" And Mid(mstrNotifyAdvice, 1, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_002" And Mid(mstrNotifyAdvice, 2, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_003" And Mid(mstrNotifyAdvice, 3, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CHARGE_001" And Mid(mstrNotifyAdvice, 7, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_BLOOD_001" And Mid(mstrNotifyAdvice, 12, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_BLOOD_003" And Mid(mstrNotifyAdvice, 10, 1) = "1" Then
        blnRecToLis = True
    ElseIf InStr(",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,", "," & strMsgItemIdentity & ",") > 0 And Mid(mstrNotifyAdvice, 4, 1) = "1" Then
        blnRecToLis = True
    ElseIf InStr(",ZLHIS_LIS_003,ZLHIS_PACS_005,", "," & strMsgItemIdentity & ",") > 0 And Mid(mstrNotifyAdvice, 5, 1) = "1" Then
        blnRecToLis = True
    End If
    
    If blnRecToLis Then
        Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
        If rsMsg Is Nothing Then Exit Sub
        Call AddMsgToLis(rsMsg)
    Else
        Call RecMsgToBub(mclsMipModule, cboUnit.ItemData(cboUnit.ListIndex), 3, strMsgItemIdentity, strMsgContent)
    End If
End Sub

Private Sub mclspath_RequestRefresh(ByVal lngPathState As Long)
'功能：临床路径中刷新病人信息列表中的状态,-1表示未导入状态
    With rptPati.SelectedRows(0)
        .Record(col_路径状态).Value = lngPathState
        .Record(col_路径状态).Caption = " "
        .Record(col_路径状态).Icon = -1 + Choose(lngPathState + 2, imgPati.ListImages("未导入").Index, imgPati.ListImages("不符合").Index, _
                imgPati.ListImages("执行中").Index, imgPati.ListImages("正常结束").Index, imgPati.ListImages("变异结束").Index)
    End With

    If rptPati.Columns(col_路径状态).Visible = False Then
        rptPati.Columns(col_路径状态).Visible = True
    End If
    rptPati.Populate
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Dim strTmp As String
    Dim intTmp As String
    If Text = "" And rptPati.SelectedRows.Count > 0 Then
        With rptPati.SelectedRows(0)
            If Not .GroupRow Then
                If Val(.Record(col_病人Id).Value) <> 0 Then intTmp = 1
            End If
            If intTmp = 1 Then
                stbThis.Panels(2).Text = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag)
                lblFee(1).Caption = GetPati费用信息(mlng病人ID, mlng主页ID)
                
                If mPatiInfo.出院日期 = CDate(0) Then
                    lblFluid(0).Visible = True
                    lblFluid(1).Visible = True
                    strTmp = Get病人输液量(mlng病人ID, mlng主页ID)
                    lblFluid(1).Caption = "今天" & Split(strTmp, ",")(0) & "ml,明天" & Split(strTmp, ",")(1) & "ml"
                Else
                    lblFluid(0).Visible = False
                    lblFluid(1).Visible = False
                End If
                
                intTmp = Get病人医嘱打印(mlng病人ID, mlng主页ID)
                lblPrint(1).Caption = IIf(intTmp = 0, "未打印", IIf(intTmp = 1, "部分打印", "全部打印"))
                If Visible And rptPati.Visible Then rptPati.SetFocus
            Else
                stbThis.Panels(2).Text = stbThis.Panels(2).Tag
                lblFee(1).Caption = ""
                lblFluid(1).Caption = ""
                lblPrint(1).Caption = ""
            End If
        End With
    Else
        Me.stbThis.Panels(2).Text = Text
    End If
End Sub

Private Sub cboPages_Click()
'功能：选择某次住院记录时，读取相关的病人信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If cboPages.ListIndex = -1 Then Exit Sub
    If cboPages.ListIndex = mintPrePage Then Exit Sub
    mintPrePage = cboPages.ListIndex
    
    On Error GoTo errH
    strSQL = "Select a.主页ID As 最大主页ID,NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄, b.住院号, b.出院病床, b.医疗付款方式, d.信息值 As 医保号, b.险类, b.当前病况, c.名称 As 护理等级, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) as 入院日期, b.出院日期, b.编目日期," & vbNewLine & _
            "       b.病人类型, b.状态, b.数据转出, b.出院科室id, b.当前病区id, a.住院次数, e.房间号" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 收费项目目录 C, 病案主页从表 D, 床位状况记录 E" & vbNewLine & _
            "Where a.病人id = b.病人id And a.病人id = [1] And b.主页id = [2] And b.护理等级id = c.Id(+) And b.病人id = d.病人id(+) And" & vbNewLine & _
            "      b.主页id = d.主页id(+) And d.信息名(+) = '医保号' And b.出院科室id = e.科室id(+) And b.病人id = e.病人id(+) And b.出院病床 = e.床号(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, cboPages.ItemData(cboPages.ListIndex))
    With rsTmp
        '保险病人姓名红色显示
        lbl姓名(1).Caption = "" & !住院号
        lbl姓名(1).ForeColor = zlDatabase.GetPatiColor(Nvl(!病人类型))
        lblPatiName(1).Caption = "" & !姓名
        lblPatiName(1).ToolTipText = lblPatiName(1).Caption
        lblPatiName(1).ForeColor = lbl姓名(1).ForeColor
        
        lbl医保号(1).Caption = Nvl(!医保号)
        lbl护理(1).Caption = Nvl(!护理等级)
        lbl付款(1).Caption = Nvl(!医疗付款方式)
        
        '危重病人病况红色显示
        lbl病况(1).Caption = Nvl(!当前病况)
        If Nvl(!当前病况) = "危" Or Nvl(!当前病况) = "重" Or Nvl(!当前病况) = "急" Then
            lbl病况(1).ForeColor = &HC0&
        Else
            lbl病况(1).ForeColor = lbl医保号(1).ForeColor
        End If
        
        lbl入院(1).Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
        If Not IsNull(!出院日期) Then
            lbl入院(1).Caption = lbl入院(1).Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")

        End If
        
        lbl类型(1).Caption = Nvl(!病人类型)
        lbl病室(1).Caption = IIf(IsNull(!房间号), "", "(" & !房间号 & ")") & !出院病床
        
        '诊断
        lblDiag(1).Caption = GetPatiDiagnose(mlng病人ID, cboPages.ItemData(cboPages.ListIndex), 2)
        
        '病人信息
        mPatiInfo.状态 = Nvl(!状态, 0)
        mPatiInfo.住院号 = Nvl(!住院号)
        mPatiInfo.床号 = Nvl(!出院病床)
        mPatiInfo.主页ID = cboPages.ItemData(cboPages.ListIndex)
        mPatiInfo.最大主页ID = Nvl(!最大主页ID, 0)
        mPatiInfo.病区ID = Nvl(!当前病区ID, 0)
        mPatiInfo.科室ID = Nvl(!出院科室ID, 0)
        mPatiInfo.入院日期 = !入院日期
        If Not IsNull(!出院日期) Then
            mPatiInfo.出院日期 = !出院日期
        Else
            mPatiInfo.出院日期 = CDate(0)
        End If
        mPatiInfo.住院次数 = Nvl(!住院次数, 0)
        mPatiInfo.数据转出 = Nvl(!数据转出, 0) <> 0
        
        Call SetPatiInfoCtlPos

    End With
    
        
    '以下信息取当前住院次数的
    strSQL = "Select B.住院号,B.病人性质,B.险类,B.出院科室ID,B.当前病区ID,Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清" & _
        " From 病案主页 B,病人余额 X" & _
        " Where B.病人ID=[1] And B.主页ID=[2] And B.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+) = 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTmp
        mPatiInfo.险类 = Val("" & !险类)
        mPatiInfo.结清 = Not IsNull(!结清)
        mPatiInfo.性质 = Nvl(!病人性质, 0)
        mPatiInfo.产科 = Sys.DeptHaveProperty(Val(!出院科室ID & ""), "产科")
    End With
        
    '刷新子窗体数据
    Call SubWinRefreshData(tbcSub.Selected)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclspath_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：临床路径中查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mfrmResponse_Closed(ByVal DataChange As Boolean)
    If DataChange Then Call LoadPatients
End Sub

Private Sub mfrmResponse_OpenObject(ByVal PatiID As Long, ByVal PageID As Long, ByVal ObjectType As Integer, ByVal ObjectID As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Dim objRow As ReportRow
    Dim blnEnabled As Boolean, blnSeek As Boolean
    Dim strTab As String, strPrivs As String
    Dim objDoc As cEPRDocument
    Dim objEmr As Object, strReturn As String, strDocID As String, strSubdocID As String, rsEmr As ADODB.Recordset
        
    '当前病人为当前要定位的病人
    blnSeek = False
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            If rptPati.SelectedRows(0).Record(col_病人Id).Value = PatiID _
                And rptPati.SelectedRows(0).Record(col_主页ID).Value = PageID Then blnSeek = True
        End If
    End If
    '自动寻找并切换显示当前要定位的病人
    If Not blnSeek Then
        For Each objRow In rptPati.Rows
            If Not objRow.GroupRow Then
                If objRow.Record(col_病人Id).Value = PatiID And objRow.Record(col_主页ID).Value = PageID Then
                    blnEnabled = timNotify.Enabled
                    timNotify.Enabled = False '避免连锁引起刷新提醒内容
                    Set rptPati.FocusedRow = objRow '选中,显示,[激活Change事件]
                    timNotify.Enabled = blnEnabled
                    blnSeek = True: Exit For
                End If
            End If
        Next
    End If
    If Not blnSeek Then
        MsgBox "当前病人清单中没有找到该病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '定位到对应的数据页面
    strTab = Decode(ObjectType, 1, "医嘱", 2, "病历", 3, "护理", 4, "护理", 5, "", 6, "医嘱", 7, "病历", 8, "病历")
    If strTab <> "" And tbcSub.Selected.Tag <> strTab Then
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub(i).Tag = strTab Then
                tbcSub(i).Selected = True
                Me.Refresh: Exit For
            End If
        Next
        If tbcSub.Selected.Tag <> strTab Then
            MsgBox "不能定位到" & strTab & "数据，可能是你没有相应的权限。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If ObjectType = 3 Or ObjectType = 4 Then '定位到具体的护理数据界面
        Call mclsTends.zlLocateData(IIf(ObjectType = 3, 1, 0))
    End If
    
    '打开对应的对象
    Select Case ObjectType
    Case 1 '住院医嘱
    Case 2, 3, 7, 8 '住院病历,护理病历,疾病证明,知情文件
        If ObjectID = "0" Or ObjectID = "" Then Exit Sub
        If IsNumeric(ObjectID) Then
            Call gobjRichEPR.EditDocument(p住院护士站, Me, cboUnit.ItemData(cboUnit.ListIndex), ObjectID)
        Else '新版病历
            If gobjEmr Is Nothing Then Exit Sub
            If InStr(ObjectID, "|") = 0 Then
                strDocID = ObjectID
                strSubdocID = ""
            Else
                strDocID = Split(ObjectID, "|")(0)
                strSubdocID = Split(ObjectID, "|")(1)
            End If
            strSQL = "Select RAWTOHEX(c.Master_Id) Masterid, RAWTOHEX(c.Id) Actlogid, RAWTOHEX(c.Basiclog_Id) Basiclogid," & vbNewLine & _
                        "       RAWTOHEX(c.Action_Id) Actionid, RAWTOHEX(b.Id) Taskid, RAWTOHEX(b.Antetype_Id) Antetypeid, d.Type Doctype," & vbNewLine & _
                        "       RAWTOHEX(a.Id) Docid, 2 Occasion, a.Sealed Besealed, nvl(e.Code,5) Docsecret, b.Subdoc_Id Subdocid,b.completor" & vbNewLine & _
                        "From Bz_Doc_Log A, Bz_Doc_Tasks B, Bz_Act_Log C, Antetype_List D, Secret_Grades E" & vbNewLine & _
                        "Where a.Actlog_Id = c.Id And a.Id = Hextoraw(:docid) And a.Id = b.Real_Doc_Id And " & IIf(strSubdocID = "", "", "b.Subdoc_Id = :subdocid And") & vbNewLine & _
                        "      b.Antetype_Id = d.Id And Decode(b.Subdoc_Id, Null, b.Antetype_Id, a.Antetype_Id) = a.Antetype_Id And" & vbNewLine & _
                        "      a.Secret = e.Code(+) And Rownum=1"
            strReturn = gobjEmr.OpenSQLRecordset(strSQL, strDocID & "^16^docid" & IIf(strSubdocID = "", "", "|" & strSubdocID & "^16^subdocid"), rsEmr)
            If strReturn <> "" Then Exit Sub
            If rsEmr.EOF Then
                                MsgBox "原始病历已不存在，无法查看。", vbInformation, gstrSysName
                                Exit Sub
                        End If
            
            strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, p电子病历管理) & ";"
            If Nvl(rsEmr!completor) = "" Then
                If InStr(strPrivs, ";文档书写;") > 0 Then '有书写权限
                    Call gobjEmr.OpenFormForModifyDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, Nvl(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), 2, strPrivs)
                Else '无权限只能查看
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "显示病历", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "查阅病历", strSubdocID)
                    End If
                End If
            Else
                If InStr(strPrivs, ";文档审订;") > 0 Then '有书写权限
                    Call gobjEmr.OpenFormForAuditDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, Nvl(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), 2, strPrivs)
                Else '无权限只能查看
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "显示病历", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "查阅病历", strSubdocID)
                    End If
                End If
            End If
        End If
    Case 4 '护理记录
    Case 5 '首页记录
        Call PrintInMedRec(mclsInOutMedRec, 1, mlng病人ID, mlng主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case 6 '医嘱报告
        If CLng(ObjectID) = 0 Then Exit Sub
        Call mclsAdvices.zlSeekAndViewEPRReport(ObjectID)
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optNotify_Click(Index As Integer)
    Call LoadNotify
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

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "住院号"
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

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index - 1: mstrFindType = objCard.名称
End Sub

Private Sub picNotify_GotFocus()
    If rptNotify.Visible Then rptNotify.SetFocus
End Sub

Private Sub picNotify_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    lblNotify.Left = 100
    lblNotify.Top = 100
    optNotify(0).Left = lblNotify.Width + lblNotify.Left
    optNotify(0).Top = lblNotify.Top
    
    optNotify(1).Top = lblNotify.Top
    optNotify(1).Left = optNotify(0).Left + optNotify(0).Width + 20
    
    rptNotify.Top = optNotify(0).Top + optNotify(0).Height + 120
    rptNotify.Left = 0
    rptNotify.Width = picNotify.Width    
    

    lngTmp = picNotify.Height - rptNotify.Top
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    rptNotify.Height = lngTmp
End Sub

Private Sub picPatiFilter_GotFocus()
    If cboUnit.Enabled And cboUnit.Visible Then
        cboUnit.SetFocus
    ElseIf rptPati.Visible Then
        rptPati.SetFocus
    End If
End Sub

Private Sub picPatiFilter_Resize()
    On Error Resume Next
    cboUnit.Width = picPatiFilter.ScaleWidth - cboUnit.Left - lblDept.Left
    txt护理条件.Width = cboUnit.Width
    cmd护理条件.Left = txt护理条件.Left + txt护理条件.Width - cmd护理条件.Width - 30
    picPara(0).Width = lbl护理条件.Width + txt护理条件.Width + 200
End Sub

Private Sub pic护理条件_GotFocus()
    lst护理条件.SetFocus
End Sub

Private Sub pic护理条件_Resize()
    On Error Resume Next
    
    lst护理条件.Left = -15
    lst护理条件.Top = -15
    lst护理条件.Width = pic护理条件.Width
    
    cmdFilterCancel.Left = pic护理条件.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lst护理条件.Height + (pic护理条件.ScaleHeight - lst护理条件.Height - cmdFilterOK.Height) / 2
    cmdFilterCancel.Top = cmdFilterOK.Top
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Childs.Count > 0 Then
        Row.Expanded = Not Row.Expanded
    End If
End Sub

Private Sub rptPati_SortOrderChanged()
    Dim objCol As ReportColumn
        
    '排序时，强行先按审查状态排序
    '子项排序功能无效，它随主项一起排序
    If rptPati.SortOrder.Count = 1 Then
        If rptPati.SortOrder(0).Index <> col_审查 Then
            Set objCol = rptPati.SortOrder(0)
            rptPati.SortOrder.DeleteAll
            rptPati.SortOrder.Add rptPati.Columns(col_审查)
            rptPati.SortOrder.Add objCol
        End If
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "病人颜色" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcPati_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    
    For i = 0 To picPara.Count - 1
        picPara(i).Visible = False
    Next
    If Item.Tag = "在院" Then
        picPara(0).Visible = True
        picPara(1).Visible = True
    ElseIf Item.Tag = "出院" Then
        picPara(2).Visible = True
        picPara(3).Visible = True
    ElseIf Item.Tag = "转出" Then
        picPara(4).Visible = True
    End If
    Call picPatiIn_Resize
    
    If Me.Visible Then
        Call LoadPatients
        If mblnNoRefNotify = False Then
            Call LoadNotify '刷新医嘱提醒
        End If
    End If
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "路径"
                Set objItem = tbcSub.InsertItem(Index, "临床路径", mcolSubForm("_路径").hwnd, 0)
                objItem.Tag = "路径"
            Case "医嘱"
                Set objItem = tbcSub.InsertItem(Index, "医嘱信息", mcolSubForm("_医嘱").hwnd, 0)
                objItem.Tag = "医嘱"
            Case "费用"
                Set objItem = tbcSub.InsertItem(Index, "费用信息", mcolSubForm("_费用").hwnd, 0)
                objItem.Tag = "费用"
            Case "病历"
                Set objItem = tbcSub.InsertItem(Index, "病历信息", mcolSubForm("_病历").hwnd, 0)
                objItem.Tag = "病历"
            Case "新病历"
                Set objItem = tbcSub.InsertItem(Index, "电子病历", mcolSubForm("_新病历").hwnd, 0)
                objItem.Tag = "新病历"
            Case "护理"
                Set objItem = tbcSub.InsertItem(Index, "护理信息", mcolSubForm("_护理").hwnd, 0)
                objItem.Tag = "护理"
            Case "监护"
                Set objItem = tbcSub.InsertItem(Index, "护理监护", mcolSubForm("_监护").hwnd, 0)
                objItem.Tag = "监护"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
     
    '刷新子窗体对应的CommandBar
    Call SubWinDefCommandBar(Item)
    
    '刷新子窗体数据
    Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    tbcSub.Tag = Item.Tag   '记录上一次选择的卡片
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboUnit.hwnd, mintPreDept)
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    mblnReturn = False
    If cboUnit.ListIndex <> -1 Then mintPreDept = cboUnit.ListIndex
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        If cboUnit.Text <> "" Then
            Set rsTmp = GetDataToUnits(cboUnit.Text)
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cboUnit, rsTmp!ID)
            Else
                cboUnit.ListIndex = mintPreDept
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            cboUnit.ListIndex = mintPreDept
        End If
    End If
End Sub

Private Sub cboUnit_Click()
'功能：刷新界面数据
'说明：从该事件开始会不重复引发相关的数据读取,包括医嘱提醒
    Dim i As Long, lngidx As Long
    
    mblnReturn = True
    
    If cboUnit.ListIndex = mintPreDept Then Exit Sub
    
    mintPreDept = cboUnit.ListIndex
        
    mlng病区ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
   
    
    '关闭业务窗体
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
    End If
    
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
    End If
    
    Call Sys.DeptHaveProperty(mlng病区ID, "护理", mblnOutDept)
    '重新读取病人
    Call LoadPatients
    
    '显示临床路径卡片
    lngidx = -1
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub(i).Tag = "路径" Then
            lngidx = i
            Exit For
        End If
    Next
    If lngidx >= 0 Then
        If HavePath(mlng病区ID) = False Then
            tbcSub(lngidx).Visible = False
            rptPati.Columns(col_路径状态).Visible = False
            rptPati.Columns(col_路径状态).Width = 0
            rptPati.Populate
            If tbcSub.Tag = "路径" Or tbcSub.Tag = "" Then tbcSub.Item(lngidx + 1).Selected = True
        Else
            If tbcSub(lngidx).Visible = False Then
                tbcSub(lngidx).Visible = True
                rptPati.Columns(col_路径状态).Visible = True
                rptPati.Columns(col_路径状态).Width = 18
                rptPati.Populate
                If tbcSub.Tag = "路径" Or tbcSub.Tag = "" Then tbcSub.Item(lngidx).Selected = True
            End If
        End If
    End If
    If Me.Visible Then Call LoadNotify
    'If Visible And rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptPati
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(col_类型, "类型", 0, False)
            objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_审查, "", 16, False)
            objCol.TreeColumn = True: objCol.Visible = False
            objCol.Sortable = False: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_图标, "", 18, False)
            objCol.Sortable = False: objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentCenter
                    
        lngidx = -1
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub(i).Tag = "路径" Then
                lngidx = i
                Exit For
            End If
        Next
        If lngidx >= 0 Then
            Set objCol = .Columns.Add(col_路径状态, "路径状态", 18, True)
        Else
            Set objCol = .Columns.Add(col_路径状态, "路径状态", 0, False): objCol.Visible = False
        End If
            
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_姓名, "姓名", 55, True)
        Set objCol = .Columns.Add(col_住院号, "住院号", 62, True)
        Set objCol = .Columns.Add(col_床号, "床号", 50, True)
        Set objCol = .Columns.Add(col_监护, "监护", 30, True)
        If mclsWardMonitor.Enabled = False Or InStr(GetInsidePrivs(p住院护士站), "护理监护") = 0 Then
            objCol.Visible = False
        End If
        Set objCol = .Columns.Add(col_性别, "性别", 30, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 30, True)
        Set objCol = .Columns.Add(col_护理等级, "护理等级", 56, True)
        Set objCol = .Columns.Add(col_费别, "费别", 55, True)
        Set objCol = .Columns.Add(col_科室, "科室", 70, True)
        Set objCol = .Columns.Add(col_住院医师, "住院医师", 55, True)
        Set objCol = .Columns.Add(col_入院日期, "入院日期", 106, True)
        Set objCol = .Columns.Add(col_出院日期, "出院日期", 106, True)
        Set objCol = .Columns.Add(col_病人类型, "病人类型", 106, True)
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_就诊卡, "就诊卡", 0, False)
        Else
            Set objCol = .Columns.Add(col_就诊卡, "就诊卡", 70, True)
        End If
        Set objCol = .Columns.Add(col_科室ID, "科室ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_住院天数, "住院天数", 56, True)
        Set objCol = .Columns.Add(col_单病种, "单病种", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_婴儿科室ID, "婴儿科室ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_婴儿病区ID, "婴儿病区ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_西医诊断, "西医诊断", 106, True)
        Set objCol = .Columns.Add(col_中医诊断, "中医诊断", 106, True)
        Set objCol = .Columns.Add(col_责任护士, "责任护士", 55, True)
        Set objCol = .Columns.Add(col_床位编制, "床位编制", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_留观号, "留观号", 62, True)
        Set objCol = .Columns.Add(col_顺序号, "顺序号", 50, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = col_类型
        Next
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
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        .GroupsOrder.Add .Columns(col_类型)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_顺序号)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_床位编制)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns(col_床号)
        .SortOrder(2).SortAscending = True
        .SortOrder.Add .Columns(col_审查)
        .SortOrder(3).SortAscending = True

    End With
    
    With rptNotify
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(c_住院号, "住院号", 62, True)
        Set objCol = .Columns.Add(c_床号, "床号", 40, True)
        Set objCol = .Columns.Add(C_状态, "状态", 150, True)
        
        Set objCol = .Columns.Add(C_消息, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_序号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_日期, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_业务, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_就诊病区, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_序号 Or objCol.Index <> C_日期 Then objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有提醒内容..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '排序 降序
        .SortOrder.Add .Columns(C_序号)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns(C_日期)
        .SortOrder(1).SortAscending = False
        
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picNotify.hwnd
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.picInfo
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With
    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop + picInfo.Height: .Height = lngBottom - lngTop - picInfo.Height
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    Dim curDate As Date
    Dim blnSetup As Boolean
        
    mlng病人ID = 0
    mlng主页ID = 0
    mlng病区ID = 0
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("病人查找方式", mintFindType, glngSys, p住院护士站, blnSetup)
    Call zlDatabase.SetPara("字体", mbytSize, glngSys, p住院护士站, blnSetup)

    strTmp = ""
    For i = 0 To chk病况条件.UBound
        strTmp = strTmp & IIf(chk病况条件(i).Value = 1, "1", "0")
    Next
    Call zlDatabase.SetPara("当前病况过滤", strTmp, glngSys, p住院护士站, blnSetup)
    Call zlDatabase.SetPara("护理等级过滤", txt护理条件.Tag, glngSys, p住院护士站, blnSetup)
    '病人范围
    curDate = zlDatabase.Currentdate
    Call zlDatabase.SetPara("最近转出天数", Val(txtChange.Text), glngSys, mlngModul, blnSetup)
    
    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("医护功能", tbcSub.Selected.Tag, glngSys, p住院护士站, blnSetup)
    End If
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End If
    If Me.Visible Then
        '公共部件固定按第一个控件的样式保存，工作站部件如果第一个是打印，则固定是图标样式,所以需恢复为其它按钮的样式
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
        If Not tbcPati.Selected Is Nothing Then
            Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", tbcPati.Selected.Index)
        End If
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "NotifyRang", IIf(optNotify(0).Value = True, 1, 0))
    End If
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsEMR = Nothing
    Set mclsAdvices = Nothing
    Set mclsEPRs = Nothing
    Set mclsTends = Nothing
    Set mclsFeeQuery = Nothing
    Set mclsInPatient = Nothing
    Set mclsWardMonitor = Nothing
    Set mclsPath = Nothing
    Set mobjReport = Nothing
    Set mobjSquareCard = Nothing
    mblnIsInit = False
    
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
        Set mfrmResponse = Nothing
    End If
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
        Set mclsInOutMedRec = Nothing
    End If
    Set mfrmActive = Nothing
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Set mobjKernel = Nothing
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
End Sub

Private Sub picInfo_GotFocus()
    If cboPages.Enabled Then cboPages.SetFocus
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraPageId.Left = 0
    fraPageId.Top = -75
    fraInfo.Top = -75
    
    fraInfo.Left = fraPageId.Left + fraPageId.Width + IIf(mbytSize = 0, 10, 30)
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left
    
    
    tbcSub.Top = picInfo.Top + picInfo.Height
End Sub

Private Sub picPati_GotFocus()
    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    picPatiFilter.Top = 0
    picPatiFilter.Left = 0
    picPatiFilter.Width = picPati.ScaleWidth
    
    tbcPati.Left = 0
    tbcPati.Top = picPatiFilter.Top + picPatiFilter.Height
    tbcPati.Width = picPati.ScaleWidth
    tbcPati.Height = picPati.ScaleHeight - tbcPati.Top - IIf(fra审查.Visible, fra审查.Height, 0)
    
    fra审查.Left = 0
    fra审查.Top = tbcPati.Top + tbcPati.Height
    fra审查.Width = picPati.ScaleWidth
    
    picPatiIn.Width = picPati.ScaleWidth
    lblFind.Width = lblDept.Width
    lblFind.Left = lblDept.Left
    lblFind.Top = lblDept.Top + lblDept.Height + 130
    PatiIdentify.Left = cboUnit.Left
    PatiIdentify.Top = lblFind.Top - 50
    PatiIdentify.Width = cboUnit.Width
    picPatiFilter.Height = PatiIdentify.Top + PatiIdentify.Height + 30
End Sub

Private Sub picPatiIn_Resize()
    Dim i As Long, Y As Long
    On Error Resume Next
    
    For i = 0 To picPara.Count - 1
        picPara(i).Width = picPatiIn.ScaleWidth
        '由于在大字体的时候picPara的高度不够因此需指定
        picPara(i).Height = IIf(mbytSize = 0, 320, 380)
        If picPara(i).Visible Then Y = Y + 1
        If i = 0 Then
            picPara(i).Top = 30
        Else
            picPara(i).Top = IIf(picPara(i - 1).Visible, picPara(i - 1).Top + picPara(i - 1).Height, 30)
        End If
    Next
    
    rptPati.Top = 30 + Y * IIf(mbytSize = 0, 320, 380)
    rptPati.Left = 0
    rptPati.Width = picPatiIn.Width
    rptPati.Height = picPatiIn.Height - rptPati.Top
End Sub

Private Sub rptNotify_KeyDown(KeyCode As Integer, Shift As Integer)
    'Panne中的Report控件需要强行处理光标顺序
    '无数据时不能捕获到vbKeyTab
    If KeyCode = vbKeyTab Then
        If Shift = vbShiftMask Then
            If rptPati.Visible Then
                rptPati.SetFocus
            Else
                If cboPages.Enabled Then cboPages.SetFocus
            End If
        Else
            If cboPages.Enabled Then cboPages.SetFocus
        End If
    End If
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'功能：自动进入医嘱校对、确认停止的执行界面
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng病人ID As Long, lng主页ID As Long
    Dim lng医嘱ID As Long, strSQLRead As String, strSQL As String
    Dim str业务 As String, str姓名 As String, str住院号 As String, str床号 As String
    Dim blnFinded As Boolean
    Dim strTmp As String
    Dim strNO As String
    Dim i As Long
    Dim strPatis As String
    Dim blnOnePati As Boolean
    Dim blnTmp As Boolean
    Dim rsTmp As ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_消息).Value
                str业务 = .Item(C_业务).Value
                lng病人ID = Val(.Item(C_病人ID).Value)
                lng主页ID = Val(.Item(C_主页ID).Value)
                str姓名 = .Item(c_姓名).Value
                str住院号 = .Item(c_住院号).Value
                str床号 = .Item(c_床号).Value
                lngIndex = .Index
            End With
            If strNO = "ZLHIS_PACS_006" Or strNO = "ZLHIS_PACS_007" Then
                strSQLRead = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "',3,'" & UserInfo.姓名 & "'," & mlng病区ID & ",null,null,'" & str业务 & "')"
            ElseIf strNO = "ZLHIS_BLOOD_007" And gbln血库系统 Then     '未回收前不允许设为已读
                If gobjPublicBlood Is Nothing And gbln血库系统 Then InitObjBlood
                If gobjPublicBlood.zlIsBloodMessageDone(1, lng病人ID, lng主页ID, 3, mlng病区ID) Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                    Call rptNotify.Populate
                End If
                Exit Sub
            Else
                strSQLRead = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "',3,'" & UserInfo.姓名 & "'," & mlng病区ID & ")"
            End If
            
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    '是不是已经定位
                    blnFinded = InStr("_" & rptPati.SelectedRows(0).Record.Tag & "_", "_" & rptNotify.SelectedRows(0).Record.Tag & "_") > 0
                End If
            End If
            '如果没找到病人，且界面没选择在院病人列表时切换列表再查找一次
            If tbcPati.Item(tbcPatiEnu.E在院).Selected = False And Not blnFinded Then
                mblnNoRefNotify = True
                tbcPati.Item(tbcPatiEnu.E在院).Selected = True
                mblnNoRefNotify = False
                blnFinded = LocatePati(rptNotify.SelectedRows(0).Record.Tag)
            End If
            
            If blnFinded And str业务 <> "" And tbcSub.Tag = "医嘱" Then   '找到病人后，再决定是否定位医嘱
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then '这两个消息的业务列存的是 医嘱id，病人来源
                    lng医嘱ID = Val(Split(str业务, ",")(0))
                ElseIf strNO = "ZLHIS_BLOOD_007" Then
                    lng医嘱ID = Val(Split(str业务, ":")(0))
                Else
                    lng医嘱ID = Val(str业务)
                End If
                If lng医嘱ID <> 0 Then
                    Call mclsAdvices.LocatedAdviceRow(lng医嘱ID)
                End If
            End If
            
            If strNO = "ZLHIS_CIS_001" Or strNO = "ZLHIS_CIS_002" Then
                '双击新开和新停医嘱消息提醒。只有在校对时才有可能操作多个病人，发送和确认停止只操作单个病人。
                '1.如果是新开可直接发送模式时，查询 病人医嘱记录
                If blnFinded Then
                    '医嘱判断权限
                    strTmp = GetInsidePrivs(p住院医嘱发送)
                    If strNO = "ZLHIS_CIS_001" Then
                        If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
                            Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                            Call rptNotify.Records.RemoveAt(lngIndex)
                            Call rptNotify.Populate
                        Else
                            If Val(zlDatabase.GetPara("发送前自动校对", glngSys, p住院医嘱发送, 0)) = 1 Then
                                If InStr(strTmp, ";发送药疗临嘱;") > 0 Or InStr(strTmp, ";发送药疗长嘱;") > 0 _
                                    Or InStr(strTmp, ";发送其他临嘱;") > 0 Or InStr(strTmp, ";发送其他长嘱;") > 0 Then
                                    
                                    blnTmp = mobjKernel.AdviceSend(Me, mlng病区ID, lng病人ID, lng主页ID, gstrPrivs, mclsMipModule)
                                    
                                    If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
                                        Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                        Call rptNotify.Records.RemoveAt(lngIndex)
                                        Call rptNotify.Populate
                                    End If
                                    
                                    If blnTmp Then
                                        If tbcSub.Selected.Tag = "医嘱" Then Call SubWinRefreshData(tbcSub.Selected)
                                    End If
                                End If
                            Else
                                If InStr(strTmp, ";医嘱校对处理;") > 0 Then
                                    blnOnePati = Val(zlDatabase.GetPara("批量医嘱校对", glngSys, p住院医嘱发送)) = 0
                                    blnTmp = mobjKernel.AdviceOperate(Me, gstrPrivs, 3, lng病人ID, lng主页ID, mlng病区ID, lng医嘱ID, mclsMipModule, strPatis, blnOnePati)
                                    If strPatis <> "" And blnTmp Then Call ReadMsg新开(strPatis)
                                    If blnTmp Then
                                        If tbcSub.Selected.Tag = "医嘱" Then Call SubWinRefreshData(tbcSub.Selected)
                                    Else
                                        If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
                                            Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                            Call rptNotify.Records.RemoveAt(lngIndex)
                                            Call rptNotify.Populate
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    ElseIf strNO = "ZLHIS_CIS_002" Then
                        If InStr(strTmp, ";医嘱确认停止;") > 0 Then
                            If Not HaveOperateAdvice(lng病人ID, lng主页ID, 1) Then
                                Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                Call rptNotify.Records.RemoveAt(lngIndex)
                                Call rptNotify.Populate
                            Else
                                blnTmp = mobjKernel.AdviceOperate(Me, gstrPrivs, 2, lng病人ID, lng主页ID, mlng病区ID, lng医嘱ID, mclsMipModule, strPatis, True)
                                If strPatis <> "" And blnTmp Then
                                    Call rptNotify.Records.RemoveAt(lngIndex)
                                    Call rptNotify.Populate
                                End If
                                If blnTmp Then
                                    If tbcSub.Selected.Tag = "医嘱" Then Call SubWinRefreshData(tbcSub.Selected)
                                Else
                                    If Not HaveOperateAdvice(lng病人ID, lng主页ID, 1) Then
                                        Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                        Call rptNotify.Records.RemoveAt(lngIndex)
                                        Call rptNotify.Populate
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                '护士站不能处理危急值消息
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then Exit Sub
                If strSQLRead = "" Then Exit Sub

                Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                Call rptNotify.Records.RemoveAt(lngIndex)
                Call rptNotify.Populate
            End If
        End If
    End If
End Sub

Private Sub ReadMsg新开(ByVal strPatis As String)
'功能：消息处理医嘱新开消息。
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim strIndexs As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            If InStr(";" & strPatis & ";", ";" & objRow.Record.Tag & ";") > 0 And objRow.Record(C_消息).Value = "ZLHIS_CIS_001" Then
                strIndexs = strIndexs & "," & objRow.Index
            End If
        End If
    Next
    If strIndexs <> "" Then
        strIndexs = Mid(strIndexs, 2)
        arrTmp = Split(strIndexs, ",")
        For i = UBound(arrTmp) To 0 Step -1
            Call rptNotify.Records.RemoveAt(Val(arrTmp(i)))
        Next
        Call rptNotify.Populate
    End If
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim strCurPati As String
    Dim lngIndex As Long
    Dim lng医嘱ID As Long
    Dim strNO As String
    Dim lng就诊病区ID As Long
    Dim lng病人ID As Long
    Dim lng主页ID As Long

    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    With rptNotify.SelectedRows(0)
        
        lngIndex = rptNotify.FocusedRow.Record.Index
        strNO = CStr(rptNotify.Rows(lngIndex).Record(C_消息).Value)
        lng病人ID = Val(rptNotify.Rows(lngIndex).Record(C_病人ID).Value)
        lng主页ID = Val(rptNotify.Rows(lngIndex).Record(C_主页ID).Value)
        lng就诊病区ID = Val(rptNotify.Rows(lngIndex).Record(C_就诊病区).Value)
                        
        If rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_LIS_003" Or rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_PACS_005" Then '这两个消息的业务列存的是 医嘱id，病人来源
            lng医嘱ID = Val(Split(rptNotify.Rows(lngIndex).Record(C_业务).Value, ",")(0))
        ElseIf rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_BLOOD_007" Then    '这个消息的业务列存的是 医嘱ID:收发ID
            lng医嘱ID = Val(Split(rptNotify.Rows(lngIndex).Record(C_业务).Value, ":")(0))
        Else
            lng医嘱ID = Val(rptNotify.Rows(lngIndex).Record(C_业务).Value)
        End If
        
        '当重复当点击同一个项目,且当前病人为当前点击的项目,则不管
        If .Record.Tag = mstrPreNotify Then
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    If Val(rptPati.SelectedRows(0).Record(col_病人Id).Value) <> 0 Then strCurPati = rptPati.SelectedRows(0).Record.Tag
                End If
            End If
        End If
        
        If .Record.Tag <> strCurPati Then
            mstrPreNotify = .Record.Tag
            '自动寻找并切换显示当前提醒的病人
            If Not LocatePati(.Record.Tag) And tbcSub.Tag = "医嘱" Then
                Call LoadPatients
                If Not LocatePati(.Record.Tag) Then
                    Call ReadAndSendMsg(strNO, lng病人ID, lng主页ID, lng就诊病区ID)
                    Call LoadNotify
                    Exit Sub
                End If
            End If
        End If
        
        If lng医嘱ID <> 0 And tbcSub.Tag = "医嘱" Then
            Call mclsAdvices.LocatedAdviceRow(lng医嘱ID)
        End If
        
    End With
    rptNotify.SetFocus
End Sub

Private Function LocatePati(ByVal strTag As String) As Boolean
'功能：通过reportControl的Record.Tag值定位病人
'参数   strTag   reportControl的Record.Tag，其内容格式为"病人ID,主页ID"

    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    
    For Each objRow In rptPati.Rows
        If objRow.GroupRow Then objRow.Expanded = True
            
        If Not objRow.GroupRow Then
            If InStr("_" & objRow.Record.Tag & "_", "_" & strTag & "_") > 0 Then
                blnEnabled = timNotify.Enabled
                timNotify.Enabled = False '避免连锁引起刷新提醒内容
                Set rptPati.FocusedRow = objRow '选中,显示,[激活Change事件]
                timNotify.Enabled = blnEnabled
                LocatePati = True: Exit Function
            End If
        End If
    Next
End Function

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    'Panne中的Report控件需要强行处理光标顺序
    '无数据时不能捕获到vbKeyTab
    If KeyCode = vbKeyTab Then
        If Shift = vbShiftMask Then
            If cboUnit.Enabled Then cboUnit.SetFocus
        Else
            If rptNotify.Visible And rptNotify.TabStop Then
                On Error Resume Next
                rptNotify.SetFocus
            Else
                If cboPages.Enabled Then cboPages.SetFocus
            End If
        End If
    Else
        cboUnit.SetFocus
        rptPati.SetFocus
        Form_KeyPress KeyCode
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup
        
    If Button = 2 Then
        Set objHitTest = rptPati.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.FindControl(, conMenu_View_Expend, , True)
            ElseIf objHitTest.Row.Childs.Count = 0 Or Val(objHitTest.Row.Record(col_病人Id).Value) <> 0 Then
                Set objPopup = cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_ManagePopup, True, False)
            End If
        End If
        
        rptPati.SetFocus
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    Dim rsTmp As New ADODB.Recordset
    Dim strCurPati As String, strSQL As String
    Dim intTmp As Integer
    Dim strTag As String
    Dim objRow As ReportRow
    Dim blnPopulate As Boolean
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub         '非正常情况
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then
            If InStr(.Record.Tag, ",") > 0 Then strCurPati = .Record.Tag
        End If
        If strCurPati = mstrPrePati Then Exit Sub
        strTag = mstrPrePati
        mstrPrePati = strCurPati
        If InStr(strCurPati, "_0") > 0 Then
            .Expanded = True
        End If
        If InStr(strTag, "_") > 0 Then
            For Each objRow In rptPati.Rows
                If Not objRow.GroupRow Then
                    If InStr(objRow.Record.Tag, "_0") > 0 And Split(objRow.Record.Tag, "_")(0) <> Split(mstrPrePati & "_", "_")(0) Then
                        objRow.Expanded = False
                        blnPopulate = True
                    End If
                End If
            Next
        End If
        
        If Not .GroupRow And strCurPati <> "" Then
            mlng病人ID = Val(.Record(col_病人Id).Value)
            mlng主页ID = Val(.Record(col_主页ID).Value)
            If InStr(strCurPati, "_") > 0 Then
                mPatiInfo.婴儿 = Val(Split(strCurPati, "_")(1))
            Else
                mPatiInfo.婴儿 = -1
            End If
            
            LockWindowUpdate Me.hwnd
            
            On Error GoTo errH
            strSQL = "Select 主页ID,NVL(病人性质,0) 病人性质 From 病案主页 Where 主页ID<>0 And 病人ID=[1] Order by 主页ID Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
            cboPages.Clear
            Do While Not rsTmp.EOF
                cboPages.AddItem "第 " & rsTmp!主页ID & " 次" & Decode(rsTmp!病人性质, 1, "(门诊留观)", 2, "(住院留观)", "")
                cboPages.ItemData(cboPages.NewIndex) = rsTmp!主页ID
                If rsTmp!主页ID = mlng主页ID Then
                    Call Cbo.SetIndex(cboPages.hwnd, cboPages.NewIndex)
                End If
                rsTmp.MoveNext
            Loop
            If cboPages.ListIndex = -1 Then
                Call Cbo.SetIndex(cboPages.hwnd, 0)
            End If
            
            mintPrePage = -1
            Call cboPages_Click
            
            Call LoadPatiAllergy(mlng病人ID, cbo过敏)
            
            '出院病人读取是否已提交审查
            If .Record(col_类型).Value = pt出院 Or .Record(col_类型).Value = pt死亡 Then
                If .Record(col_图标).Value = -1 Then
                    '1-等待审查;2-拒绝审查;3-正在审查;4-审查反馈;5-审查归档
                    If .Record(col_审查).Value = 0 Or .Record(col_审查).Value = 999 Then
                        .Record(col_图标).Value = 0
                    ElseIf .Record(col_审查).Value = 1 Or .Record(col_审查).Value = 2 Then
                        .Record(col_图标).Value = 1
                    Else
                        .Record(col_图标).Value = IIf(PatiMedRecHaveSubmit(mlng病人ID, mlng主页ID), 1, 0)
                    End If
                End If
            End If
            
            LockWindowUpdate 0
            
            stbThis.Panels(2).Text = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag)
            lblFee(1).Caption = GetPati费用信息(mlng病人ID, mlng主页ID)
            '出院病人不显示输液量
            If mPatiInfo.出院日期 = CDate(0) Then
                lblFluid(0).Visible = True
                lblFluid(1).Visible = True
                strSQL = Get病人输液量(mlng病人ID, mlng主页ID)
                lblFluid(1).Caption = "今天" & Split(strSQL, ",")(0) & "ml,明天" & Split(strSQL, ",")(1) & "ml"
            Else
                lblFluid(0).Visible = False
                lblFluid(1).Visible = False
            End If

            lblPrint(0).Visible = True
            lblPrint(1).Visible = True
            intTmp = Get病人医嘱打印(mlng病人ID, mlng主页ID)
            lblPrint(1).Caption = IIf(intTmp = 0, "未打印", IIf(intTmp = 1, "部分打印", "全部打印"))
            If Visible And rptPati.Visible Then rptPati.SetFocus
        Else
            Call ClearPatiInfo
            '按无数据刷新子窗体
            Call SubWinRefreshData(tbcSub.Selected)
            
            stbThis.Panels(2).Text = stbThis.Panels(2).Tag
        End If
    End With
    Call SetPatiInfoCtlPos
    If blnPopulate Then rptPati.Populate
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    Dim curDate As Date, intDay As Integer
    
    '病人显示范围
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院护士站, 7))
    '如果大于30天就取缺省值
    If mintChange > 30 Then mintChange = 7
    
    '出院病人时间范围，固定为过去3天
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    mdtOutBegin = Format(mdtOutEnd - 3, "yyyy-MM-dd 00:00:00")
        
    '医嘱提醒刷新设置
    mstrNotifyAdvice = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, p住院护士站, "000000000000")
    mintNotifyDay = Val(zlDatabase.GetPara("自动刷新医嘱天数", glngSys, p住院护士站, 1))
    mintNotify = Val(zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, p住院护士站))
    mbln消息语音 = Val(zlDatabase.GetPara("启用语音提示", glngSys, p住院护士站)) = 1
    '病案审查反馈天数
    mlngMedRedDay = Val(zlDatabase.GetPara("病案审查反馈天数", glngSys, p住院护士站))
    '字体设置
    mbytSize = zlDatabase.GetPara("字体", glngSys, p住院护士站, "0")
    
    '病案首页标准
    mintMecStandard = Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
    
    mblnCardOrder = (Val(zlDatabase.GetPara("床位卡片排序方式", glngSys, P新版护士站, 0)) = 0)
    
End Sub

Private Function InitNurselevel() As Boolean
'功能：初始化住院护理等级条件
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSel As String
    Dim blnSelAll As Boolean
    
    txt护理条件.Text = ""
    txt护理条件.Tag = ""
    
    lst护理条件.AddItem "全部"
    strSel = zlDatabase.GetPara("护理等级过滤", glngSys, p住院护士站, "", Array(lbl护理条件, txt护理条件, cmd护理条件), InStr(mstrPrivs, "参数设置") > 0)
    blnSelAll = True
    
    strSQL = _
        " Select ID,编码,名称 From 收费项目目录 Where 类别='H' And 项目特性>=1" & _
        " And (撤档时间 is NULL Or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
        " Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitNurselevel")
    Do While Not rsTmp.EOF
        lst护理条件.AddItem rsTmp!名称
        lst护理条件.ItemData(lst护理条件.NewIndex) = rsTmp!ID
        If strSel = "" Or InStr("," & strSel & ",", "," & rsTmp!ID & ",") > 0 Then
            txt护理条件.Text = txt护理条件.Text & "," & rsTmp!名称
            txt护理条件.Tag = txt护理条件.Tag & "," & rsTmp!ID
        Else
            blnSelAll = False
        End If
        rsTmp.MoveNext
    Loop
    
    If blnSelAll Then
        txt护理条件.Text = "全部"
        txt护理条件.Tag = ""
    Else
        txt护理条件.Text = Mid(txt护理条件.Text, 2)
        txt护理条件.Tag = Mid(txt护理条件.Tag, 2)
    End If
    
    '设置条件大小
    lst护理条件.Height = lst护理条件.ListCount * 210 + 30
    pic护理条件.Height = lst护理条件.Height + cmdFilterOK.Height + 120
    
    InitNurselevel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    mstrUnits = GetUser病区IDs
    
    cboUnit.Clear
    Set rsTmp = GetDataToUnits
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "全院病人") > 0 Then
                If rsTmp!ID = UserInfo.部门ID Then '直接所属优先
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & mstrUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '所属缺省病区包含的可能有多个
                If rsTmp!缺省 = 1 And cboUnit.ListIndex = -1 Then
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call Cbo.SetIndex(cboUnit.hwnd, 0)
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Get病案图标序号(ByVal lng状态 As Long) As Long
    Dim i As Long
    
    i = imgPati.ListImages("Pati").Index
    Select Case lng状态
        Case 1
            i = imgPati.ListImages("等待审查").Index
        Case 2
            i = imgPati.ListImages("拒绝审查").Index
        Case 13
            i = imgPati.ListImages("正在抽查").Index
        Case 3
            i = imgPati.ListImages("正在审查").Index
        Case 14
            i = imgPati.ListImages("抽查反馈").Index
        Case 4
            i = imgPati.ListImages("审查反馈").Index
        Case 16
            i = imgPati.ListImages("抽查整改").Index
        Case 6
            i = imgPati.ListImages("审查整改").Index
    End Select
    Get病案图标序号 = i - 1 '编号是从0开始的
End Function

Private Function LoadPatients() As Boolean
'功能：读取病人列表
    Dim rsPati As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim objParent As ReportRecord
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Dim objRow As ReportRow
    Dim strPatiRow As String, lngPatiRow As Long, intBedLen As Integer
    
    Dim strFilter As String, strMonitor As String
    Dim strSQL As String, i As Long, j As Long
    Dim lngCount(0 To 7) As Long, strState As String    '用于显示病人分类统计数目
    Dim strTmpDate As String                            '转出病人查询时间范围条件
    Dim blnIsFind As Boolean                            '判断是否是查找住院号及住院号是否为空
    Dim strTmpOut As String                             '查询出院病人
    Dim int结清 As Integer                              '出院病人的结清状态
    
    Dim rsBaby As ADODB.Recordset
    Dim strSQLBaby As String
    Dim objBabyParent As ReportRecord
    Dim lngCol As Long
    Dim strPre病人类型 As String, str字段诊断 As String
    Dim bln床位编制 As Boolean
    
    '当页面下拉框清空，F5刷新，应该恢复上一个的值
    If cboUnit.ListIndex = -1 Then Call Cbo.SetIndex(cboUnit.hwnd, mintPreDept)
     
    '床位长度固定为10
    intBedLen = 10
    mblnUnRefresh = True
    '病人过滤条件
    strFilter = ""
    strTmpDate = ""
    If txt护理条件.Tag <> "" Then
        strFilter = strFilter & " And Instr(','||[3]||',',','||B.护理等级ID||',')>0"
    End If
    strState = ""
    For i = 0 To chk病况条件.UBound
        If chk病况条件(i).Value = 1 Then
            strState = strState & "," & chk病况条件(i).Caption
        End If
    Next
    strState = Mid(strState, 2)
    If Not (UBound(Split(strState, ",")) = chk病况条件.UBound Or strState = "") Then
        strFilter = strFilter & " And Instr(','||[4]||',',','||B.当前病况||',')>0"
    End If
    
    If mintChange = 0 Then
        strTmpDate = ""
    Else
        strTmpDate = " And C.终止时间 Between Sysdate-[2] And Sysdate "
    End If
    
    '入院待入住和转科待入住病人(病人科室所属的病区都可接收),转病区待入住
    'c.科室id + 0,说明：通过H表的索引连接过滤后，记录数量很少，再连接B表则更快
    '“排序”作为“类型”的value
    If tbcPati.Selected.Tag = "待入住" Then
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.状态,1,0,Decode(c.开始原因,3,1,2)) As 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2," & _
            " Decode(B.状态,1,'入院待入住病人',Decode(c.开始原因,3,'转科待入住病人','转病区待入住病人')) As 类型," & _
            " a.病人id, b.主页id, A.门诊号,B.住院号, NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄," & vbNewLine & _
            " d.名称 As 科室, c.科室id, c.经治医师 As 住院医师, b.病案状态, LPAD(c.床号," & intBedLen & ",' ') as 床号," & _
            " e.名称 As 护理等级, b.费别, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) as 入院日期, b.出院日期, b.病人类型, b.状态, b.险类, a.就诊卡号,b.留观号," & vbNewLine & _
            " Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,q.序号 as 新生儿序号,b.单病种,B.婴儿科室ID,B.婴儿病区ID,'' AS 西医诊断,'' AS 中医诊断 " & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D, 收费项目目录 E,病人新生儿记录 Q" & vbNewLine & _
            "Where a.在院 = 1 And a.病人id = b.病人id  And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And (C.病区ID=[1] or C.病区ID is null) And c.科室id = d.Id  And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null) " & vbNewLine & _
            "      And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            "      And b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null" & vbNewLine & _
            "      And (c.开始原因 in(1,3) And Exists(Select 1 From 病区科室对应 H Where c.科室id = h.科室id And h.病区id = [1]) or c.开始原因=15 And c.病区id = [1])" & vbNewLine & _
            "      And ((c.开始原因 = 1 And b.状态 = 1) Or (c.开始原因 in (3,15) And c.开始时间 Is Null And b.状态 = 2)) "
        
        strSQLBaby = "Select q.病人id,q.主页id,q.序号,q.婴儿姓名,q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
            " From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D, 收费项目目录 E,病人新生儿记录 Q" & vbNewLine & _
            " Where a.在院 = 1 And a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And (C.病区ID=[1] or C.病区ID is null) And c.科室id = d.Id  And b.病人id=q.病人ID And b.主页ID=q.主页ID " & vbNewLine & _
            "      And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            "      And b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null" & vbNewLine & _
            "      And (c.开始原因 in(1,3) And Exists(Select 1 From 病区科室对应 H Where c.科室id = h.科室id And h.病区id = [1]) or c.开始原因=15 And c.病区id = [1])" & vbNewLine & _
            "      And ((c.开始原因 = 1 And b.状态 = 1) Or (c.开始原因 in (3,15) And c.开始时间 Is Null And b.状态 = 2)) "
    End If
    '在院病人
    If tbcPati.Selected.Tag = "在院" Then
        str字段诊断 = ",first_value(Decode(Sign(h.诊断类型-10),-1,h.诊断描述,'')) " & _
            " Over(partition By h.病人id,H.主页ID Order By sign(h.诊断类型-10),decode(h.记录来源,4,0,h.记录来源) desc,Decode(h.诊断类型,1,1,2,2,3,3,0) DESC,h.诊断次序) As 西医诊断"
        If Sys.DeptHaveProperty(cboUnit.ItemData(cboUnit.ListIndex), "中医科") Then
            str字段诊断 = str字段诊断 & ",first_value(Decode(Sign(h.诊断类型-10),1,h.诊断描述,'')) " & _
            " Over(partition By h.病人id,H.主页ID Order By sign(h.诊断类型-10) desc,decode(h.记录来源,4,0,h.记录来源) desc,Decode(h.诊断类型,11,1,12,2,13,3,0) DESC,h.诊断次序) As 中医诊断"
        Else
            str字段诊断 = str字段诊断 & ",null as 中医诊断"
        End If
        strSQL = _
            "Select /*+ RULE */ Distinct Decode(B.状态,3,4,3) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.状态,3,'预出院病人','在院病人') as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.病案状态," & _
            " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,E.名称 as 护理等级,B.费别,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) as 入院日期,B.出院日期,B.病人类型," & IIf(mblnCardOrder, "f.顺序号,", "d.编码 as 床位编制,f.顺序号,") & _
            " B.状态,B.险类,A.就诊卡号,b.留观号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,q.序号 as 新生儿序号,b.单病种,B.婴儿科室ID,B.婴儿病区ID " & str字段诊断 & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人新生儿记录 Q,在院病人 R,病人诊断记录 H" & IIf(mblnCardOrder, ",床位状况记录 F", ",床位编制分类 D,床位状况记录 F") & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And Nvl(B.状态,0)<>1 And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id " & _
            " And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null)" & IIf(mblnCardOrder, "and b.病人id=f.病人id(+)", " and b.病人id=f.病人id(+) and f.床位编制=D.名称(+)") & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And a.病人ID=R.病人ID And A.当前病区ID+0=R.病区ID And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL" & strFilter
        
        strSQLBaby = "Select q.病人id,q.主页id,q.序号,q.婴儿姓名,q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人新生儿记录 Q,在院病人 R" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And Nvl(B.状态,0)<>1" & _
            " And b.病人id=q.病人ID And b.主页ID=q.主页ID " & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And a.病人ID=R.病人ID And A.当前病区ID+0=R.病区ID And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL" & strFilter
        bln床位编制 = Not mblnCardOrder
    End If
    '出院病人:出院病人可能已有多次住院
    If tbcPati.Selected.Tag = "出院" Then
        '结清未结清的
        If chk结清(0).Value = 1 And chk结清(1).Value = 1 Then
            int结清 = 0               '都显示
        ElseIf chk结清(0).Value = 0 And chk结清(1).Value = 1 Then
            int结清 = 1               '只显示未结清的
        ElseIf chk结清(0).Value = 1 And chk结清(1).Value = 0 Then
            int结清 = 2              '只显示已结清的
        End If
        
        '判断是否是查找住院号及住院号是否为空
        If mstrFindType = "住院号" And Trim(PatiIdentify.Text) <> "" Then blnIsFind = True
        
        '由于取消了显示出院病人参数，所以查找时显示查找的人员和时间范围的人员
        If blnIsFind Then
            '由于取消了显示出院病人参数，所以查找时显示查找的人员和时间范围的人员
'            strTmpOut = " And (B.出院日期 Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
'                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')  Or (b.住院号=[5] And B.出院日期 is Not Null)) "

            strTmpOut = " And (B.出院日期 Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')  " & _
                        IIf(int结清 = 0, "", " And " & IIf(int结清 = 1, "", "Not") & " Exists(Select 1 From 病人余额 Where a.病人id = 病人id And 性质 = 1 And Nvl(费用余额, 0) <> 0)") & _
                        " Or (b.住院号=[5] And B.出院日期 is Not Null)) "
        Else
            '用户不是查找，显示参数对应时间内的病人
            strTmpOut = " And B.出院日期 Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
            strTmpOut = strTmpOut & IIf(int结清 = 0, "", " And " & IIf(int结清 = 1, "", "Not") & " Exists(Select 1 From 病人余额 Where a.病人id = 病人id And 性质 = 1 And Nvl(费用余额, 0) <> 0)")
        End If
    
        strSQL = _
            "Select /*+ RULE */ Decode(B.出院方式,'死亡',6,5) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.病案状态," & _
            " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,E.名称 as 护理等级,B.费别,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) as 入院日期,B.出院日期,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,b.留观号,Nvl(b.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,q.序号 as 新生儿序号,b.单病种,B.婴儿科室ID,B.婴儿病区ID,'' AS 西医诊断,'' AS 中医诊断 ,B.责任护士 " & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人新生儿记录 Q" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
            " And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null)" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID+0=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL" & strTmpOut
        
        strSQLBaby = "Select q.病人id,q.主页id,q.序号,q.婴儿姓名,q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人新生儿记录 Q" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
            " And b.病人id=q.病人ID And b.主页ID=q.主页ID " & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID+0=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL" & strTmpOut
    End If
    '转出病人:在院,医生和床号显示本科转出前的
    If tbcPati.Selected.Tag = "转出" Then
        strSQL = _
            "Select /*+ RULE */ Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄,D.名称 as 科室,C.科室ID,C.经治医师 as 住院医师,B.病案状态," & _
            " LPAD(c.床号," & intBedLen & ",' ') as 床号,E.名称 as 护理等级,B.费别,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) as 入院日期,B.出院日期,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,b.留观号,Nvl(b.路径状态,-1) 路径状态,trunc(Nvl(b.出院日期, Sysdate))-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,q.序号 as 新生儿序号,b.单病种,B.婴儿科室ID,B.婴儿病区ID,'' AS 西医诊断,'' AS 中医诊断 " & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人新生儿记录 Q" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
            " And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null)" & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
            " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID" & _
            " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) " & strTmpDate & _
            " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL "
            
        strSQLBaby = "Select q.病人id,q.主页id,q.序号,q.婴儿姓名,q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人新生儿记录 Q" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
            " And b.病人id=q.病人ID And b.主页ID=q.主页ID " & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
            " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID" & _
            " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) " & strTmpDate & _
            " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL "
    End If
    
    If strSQL = "" Then
        rptPati.Records.DeleteAll
        rptPati.Populate
        mblnUnRefresh = False
        Screen.MousePointer = 0
        LoadPatients = True
        Exit Function
    End If
    
    strSQL = strSQL & " Order by" & IIf(tbcPati.Selected.Tag = "在院", " 顺序号,", "") & IIf(bln床位编制, " 床位编制,", "") & " 床号,排序,排序2,主页ID Desc"
 
    Screen.MousePointer = 11
    On Error GoTo errH
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), _
        mintChange, txt护理条件.Tag, strState, Val(Trim(PatiIdentify.Text)))
    If strSQLBaby <> "" Then
        Set rsBaby = zlDatabase.OpenSQLRecord(strSQLBaby, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), _
            mintChange, txt护理条件.Tag, strState, Val(Trim(PatiIdentify.Text)))
    End If
      
    '添加诊断预处理
    If tbcPati.Selected.Tag = "在院" Then
        rptPati.Columns(col_西医诊断).Visible = True
        rptPati.Columns(col_中医诊断).Visible = Sys.DeptHaveProperty(cboUnit.ItemData(cboUnit.ListIndex), "中医科")
    Else
        rptPati.Columns(col_西医诊断).Visible = False
        rptPati.Columns(col_中医诊断).Visible = False
    End If
        If tbcPati.Selected.Tag = "出院" Then
        rptPati.Columns(col_责任护士).Visible = True
    Else
        rptPati.Columns(col_责任护士).Visible = False
    End If
    
    '记录现在选中的病人
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            If rptPati.SelectedRows(0).Record.Tag <> "" Then
                lngPatiRow = rptPati.SelectedRows(0).Index '用于快速重新定位
                strPatiRow = rptPati.SelectedRows(0).Record.Tag
            End If
        End If
    End If
    rptPati.Records.DeleteAll
    rptPati.Columns(col_审查).Visible = False
    
    If mclsWardMonitor.Enabled And InStr(GetInsidePrivs(p住院护士站), "护理监护") > 0 Then
        strMonitor = mclsWardMonitor.GetListPati
    End If
    
    
    '刷新后分组自动展开
    For i = 1 To rsPati.RecordCount
        '根据提交审查情况添加父行
        If Nvl(rsPati!病案状态, 0) <> 0 Then
            rptPati.Columns(col_审查).Visible = True
            
            '当病人类型发生变化时要另外开一个分支，否则会导致分组不对
            If strPre病人类型 <> rsPati!类型 & "" Then
                Set objParent = Nothing
                strPre病人类型 = rsPati!类型 & ""
            End If
            
            If objParent Is Nothing Then
                Set objParent = Me.rptPati.Records.Add()
            ElseIf objParent.Tag <> CStr(rsPati!病案状态) Then
                Set objParent = Me.rptPati.Records.Add()
            End If
            If objParent.Tag <> CStr(rsPati!病案状态) Then
                objParent.Tag = CStr(rsPati!病案状态)
                objParent.Expanded = True
                For j = 0 To rptPati.Columns.Count - 1
                    If j = col_类型 Then
                        Set objItem = objParent.AddItem(Val(rsPati!排序))
                        objItem.Caption = rsPati!类型
                    ElseIf j = col_审查 Then
                        Set objItem = objParent.AddItem(Val(rsPati!病案状态))
                        objItem.Caption = " "
                    ElseIf j = col_姓名 Then
                        Set objItem = objParent.AddItem(CStr(Decode(rsPati!病案状态, 1, "等待审查", 2, "拒绝审查", 3, "正在审查", 4, "审查反馈")))
                        objItem.ForeColor = rptPati.PaintManager.GroupForeColor
                    Else
                        Set objItem = objParent.AddItem("")
                        If j = col_图标 Then objItem.Icon = Get病案图标序号(rsPati!病案状态) 'rsPati!病案状态 + imgPati.ListImages("等待审查").Index - 2
                    End If
                    objItem.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
                Next
            End If
        Else
            Set objParent = Nothing
        End If
            
        '添加具体的病人数据行(子行或独立行)
        If Not objParent Is Nothing Then
            Set objRecord = objParent.Childs.Add()
        Else
            Set objRecord = Me.rptPati.Records.Add()
        End If
        
        objRecord.Tag = CStr(rsPati!病人ID & "," & rsPati!主页ID) '用于病人定位
        
        Set objItem = objRecord.AddItem(Val(rsPati!排序)) '分组以Value进行排序
        objItem.Caption = rsPati!类型
        
        Set objItem = objRecord.AddItem(Val(Decode(Nvl(rsPati!病案状态, 0), 0, 999, rsPati!病案状态)))
        objItem.Caption = " "
        If Nvl(rsPati!病案状态, 0) = 2 Then
            objRecord.PreviewText = "  理由:" & GetRefuseReason(rsPati!病人ID, rsPati!主页ID)
        End If
        
        '图标:注意用在这里是从0开始编号。
        '     图标Value用于存放是否已提交审查，点击才读取
        Set objItem = objRecord.AddItem(-1)
        objItem.Caption = " "
        If Nvl(rsPati!病案状态, 0) <> 0 Then
            objItem.Icon = Get病案图标序号(rsPati!病案状态)
        ElseIf "" & rsPati!单病种 <> "" Then
            objItem.Icon = imgPati.ListImages("单病种").Index - 1
        End If
        
        '      lng路径状态=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
        Set objItem = objRecord.AddItem(Val("" & rsPati!路径状态))
        objItem.Caption = " "
        objItem.Icon = -1 + Choose(rsPati!路径状态 + 2, imgPati.ListImages("未导入").Index, imgPati.ListImages("不符合").Index, _
            imgPati.ListImages("执行中").Index, imgPati.ListImages("正常结束").Index, imgPati.ListImages("变异结束").Index)
            
        
        objRecord.AddItem Val(rsPati!病人ID)
        objRecord.AddItem Val(rsPati!主页ID)
        objRecord.AddItem CStr(Nvl(rsPati!姓名))
        
        If mblnOutDept Then
            Set objItem = objRecord.AddItem("" & rsPati!门诊号)
            objItem.Caption = Nvl(rsPati!门诊号, " ")
        Else
            Set objItem = objRecord.AddItem("" & rsPati!住院号)
            objItem.Caption = Nvl(rsPati!住院号, " ")
        End If
        
        
        Set objItem = objRecord.AddItem(zlStr.Lpad(Nvl(rsPati!床号), 10)) 'Value用于排序
        objItem.Caption = CStr(Trim(Nvl(rsPati!床号, " "))) '为空时会被Value替代
        '新生儿图标
        If rsPati!新生儿序号 & "" <> "" Then
            objItem.Icon = imgPati.ListImages("Child").Index - 1
            If rsPati!婴儿科室ID & "" <> "" Then objItem.Icon = imgPati.ListImages("Out").Index - 1
        End If
            
        Set objItem = objRecord.AddItem(" ")    '护理监护
        If strMonitor <> "" And Not IsNull(rsPati!住院号) Then
            If InStr("," & strMonitor & ",", "," & rsPati!住院号 & ",") > 0 Then
                objItem.Caption = "★"
            End If
        End If
        
        objRecord.AddItem CStr(Nvl(rsPati!性别))
        objRecord.AddItem CStr(Nvl(rsPati!年龄))
        objRecord.AddItem CStr(Nvl(rsPati!护理等级))
        objRecord.AddItem CStr(Nvl(rsPati!费别))
        objRecord.AddItem CStr(Nvl(rsPati!科室))
        objRecord.AddItem CStr(Nvl(rsPati!住院医师))
        objRecord.AddItem Format(rsPati!入院日期, "yyyy-MM-dd HH:mm")
        objRecord.AddItem Format(Nvl(rsPati!出院日期), "yyyy-MM-dd HH:mm")
        objRecord.AddItem CStr(Nvl(rsPati!病人类型))
        objRecord.AddItem CStr(Nvl(rsPati!就诊卡号))
        objRecord.AddItem Val("" & rsPati!科室ID)
        objRecord.AddItem Val(Trim(IIf(CStr("" & rsPati!住院天数) = "0", "1", CStr("" & rsPati!住院天数))))     '待入住病人为空
        objRecord.AddItem "" & rsPati!单病种
        objRecord.AddItem Val("" & rsPati!婴儿科室ID)
        objRecord.AddItem Val("" & rsPati!婴儿病区ID)
        '添加诊断
        objRecord.AddItem CStr(Nvl(rsPati!西医诊断))
        objRecord.AddItem CStr(Nvl(rsPati!中医诊断))
        
        If tbcPati.Selected.Tag = "出院" Then
            objRecord.AddItem rsPati!责任护士 & ""
        Else
            '填充列数据
            objRecord.AddItem ""
        End If
        
        If bln床位编制 Then
            objRecord.AddItem rsPati!床位编制 & ""
        Else
            objRecord.AddItem ""
        End If
        '留观号
        objRecord.AddItem "" & rsPati!留观号
        
        If tbcPati.Selected.Tag = "在院" Then
            Set objItem = objRecord.AddItem("" & Nvl(rsPati!顺序号)) 'Value用于排序
            objItem.Caption = Nvl(rsPati!顺序号, " ") '为空时会被Value替代
        Else
            objRecord.AddItem ""
        End If
        '显示病人颜色
        objRecord.Item(col_姓名).ForeColor = zlDatabase.GetPatiColor(Nvl(rsPati!病人类型))
        For j = 0 To rptPati.Columns.Count - 1
            If j <> col_类型 And j <> col_审查 And j <> col_图标 And j <> col_顺序号 Then
                objRecord.Item(j).ForeColor = objRecord.Item(col_姓名).ForeColor
            End If
        Next
        
        '统计病人数目
        lngCount(Val(rsPati!排序)) = lngCount(Val(rsPati!排序)) + 1
        '根据是否有婴儿添加婴儿行
        If Not rsBaby Is Nothing Then
            Set objBabyParent = objRecord
            rsBaby.Filter = "病人ID=" & objBabyParent(col_病人Id).Value & " and 主页ID=" & objBabyParent(col_主页ID).Value
            If Not rsBaby.EOF Then
                rptPati.Columns(col_审查).Visible = True
                rsBaby.Sort = "序号"
                objBabyParent.Expanded = False
                For lngCol = 1 To rsBaby.RecordCount
                    Set objRecord = objBabyParent.Childs.Add()
                    objRecord.Tag = objBabyParent.Tag & "_" & rsBaby!序号
                    For j = 0 To rptPati.Columns.Count - 1
                        Set objItem = objRecord.AddItem(objBabyParent(j).Value)
                            objItem.Caption = " "
                            objItem.ForeColor = objBabyParent(j).ForeColor
                        Select Case j
                        Case col_姓名
                            objItem.Caption = "   " & rsBaby!婴儿姓名
                        Case col_性别
                            objItem.Caption = "" & rsBaby!婴儿性别
                        Case col_住院号
                            objItem.Caption = objBabyParent(j).Value & "-" & rsBaby!序号
                        Case col_床号
                            objItem.Caption = " "
                            If "" & rsBaby!婴儿性别 = "男" Then
                                objItem.Icon = imgPati.ListImages("Child").Index - 1
                            Else
                                objItem.Icon = imgPati.ListImages("Fbaby").Index - 1
                            End If
                            If lngCol = 1 And objBabyParent(j).Icon = imgPati.ListImages("Child").Index - 1 Then
                                objBabyParent(j).Icon = objItem.Icon
                            End If
                        Case col_年龄
                            objItem.Caption = "" & rsBaby!年龄
                        End Select
                    Next
                    rsBaby.MoveNext
                Next
                objBabyParent.Tag = objBabyParent.Tag & "_0"
            End If
        End If
        rsPati.MoveNext
    Next
    If mblnOutDept Then
        rptPati.Columns.Find(col_住院号).Caption = "门诊号"
    Else
        rptPati.Columns.Find(col_住院号).Caption = "住院号"
    End If
    rptPati.Populate
    '根据江陵医院需求加入病人数目统计信息
    strState = "共 " & rsPati.RecordCount & " 个病人"
    For i = LBound(lngCount) To UBound(lngCount)
        If lngCount(i) > 0 Then
            Select Case i
            Case 0
                strState = strState & "，入院待入住:"
            Case 1
                strState = strState & "，转科待入住:"
            Case 2
                strState = strState & "，转病区待入住:"
            Case 3
                strState = strState & "，在院:"
            Case 4
                strState = strState & "，预出院:"
            Case 5
                strState = strState & "，出院:"
            Case 6
                strState = strState & "，死亡:"
            Case 7
                strState = strState & "，转出:"
            End Select
            strState = strState & lngCount(i) & "人"
        End If
    Next
    stbThis.Panels(2).Text = strState
    stbThis.Panels(2).Tag = strState
    
    '定位病人行:在Populate之后
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Or rsPati.RecordCount > 1 And lngPatiRow = 0 Then
        Call ClearPatiInfo
        '按无数据刷新子窗体
        Call SubWinRefreshData(tbcSub.Selected)
        
        If tbcPati.Selected.Tag = "在院" And mblnNoRefNotify = False Then
            mstrPreNotify = ""
            rptNotify.Records.DeleteAll
            rptNotify.Populate
            rptNotify.TabStop = False
        End If
    Else
        '取指定病人行
        If strPatiRow <> "" Then
            '先快速定位
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To rptPati.Rows.Count - 1
                    If Not rptPati.Rows(i).GroupRow Then
                        If rptPati.Rows(i).Record.Tag = strPatiRow Then
                            Set objRow = rptPati.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        '取第一个非分组行
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow And rptPati.Rows(i).Childs.Count = 0 Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        
        Set rptPati.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
        
    End If
    
    Screen.MousePointer = 0
    LoadPatients = True
    
    '同步刷新审查反馈信息
    Call LoadResponse
    
    mblnUnRefresh = False
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnUnRefresh = False
End Function

Private Sub ClearPatiInfo()
'功能：清除单个病人相关的显示信息
    mlng病人ID = 0
    mlng主页ID = 0
    
    mPatiInfo.状态 = 0
    mPatiInfo.婴儿 = -1
    mPatiInfo.住院号 = ""
    mPatiInfo.床号 = ""
    mPatiInfo.主页ID = 0
    mPatiInfo.最大主页ID = 0
    mPatiInfo.病区ID = 0
    mPatiInfo.科室ID = 0
    mPatiInfo.入院日期 = CDate(0)
    mPatiInfo.出院日期 = CDate(0)
    mPatiInfo.住院次数 = 0
    mPatiInfo.数据转出 = False
    mPatiInfo.产科 = False
    mPatiInfo.结清 = False
    mPatiInfo.险类 = 0
    mPatiInfo.性质 = 0
        
    cboPages.Clear
    cbo过敏.Clear
    
    lbl类型(1).Caption = ""
    lbl姓名(1).Caption = ""
    lblPatiName(1).Caption = ""
    lblPatiName(1).ToolTipText = ""
    lbl医保号(1).Caption = ""
    lbl付款(1).Caption = ""
    lbl护理(1).Caption = ""
    lbl病况(1).Caption = ""
    lbl入院(1).Caption = ""
    lblDiag(1).Caption = ""
    lbl病室(1).Caption = ""
    lblFee(1).Caption = ""
    lblFluid(1).Caption = ""
    lblPrint(1).Caption = ""
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal lngPatiID As Long)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Dim strTmp As String
            
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
                If mstrFindType = "床号" Then
                    If UCase(Trim(.Record(col_床号).Value)) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "住院号" Then
                    If .Record(col_住院号).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "留观号" Then
                    If .Record(col_留观号).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "就诊卡" Then
                    If UCase(.Record(col_就诊卡).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "姓名" Then
                    If .Record(col_姓名).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptPati.Rows.Count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptPati.FocusedRow = rptPati.Rows(i)
        
'        If Not rptPati.Visible Then
'            For i = 1 To dkpMain.PanesCount
'                If dkpMain.Panes(i).Handle = picPati.hwnd Then
'                    dkpMain.Panes(i).Select
'                End If
'            Next
'        End If
        If rptPati.Visible Then rptPati.SetFocus
    Else
        If mstrFindType = "住院号" And Not mblnIsFindAgain Then '住院号
            mblnIsFindAgain = True
            Call LoadPatients
            Call ExecuteFindPati
            mblnIsFindAgain = False
        Else
            blnReStart = True
            If tbcPati.Selected.Tag = "出院" And mstrFindType = "住院号" Then
                strTmp = GetNoPatiWhy(PatiIdentify.Text)
            Else
                strTmp = IIf(blnNext, "后面已", "") & "找不到符合条件的病人。"
            End If
            MsgBox strTmp, vbInformation, gstrSysName
        End If
    End If
End Sub

Function ExecuteMonitor() As Boolean
'功能：调用监护仪
    Dim strUser As String, strPass As String, strServer As String
    Dim arrInfo As Variant, i As Long
    
    'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=ORCL";Persist Security Info=True;User ID=zlhis;Password=HIS;Data Provider=MSDASQL
    'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source=ORCL;Extended Properties="PLSQLRSet=1;DistribTx=0"
    arrInfo = Split(gcnOracle.ConnectionString, ";")
    For i = 0 To UBound(arrInfo)
        If UCase(arrInfo(i)) Like UCase("User ID=*") Then
            strUser = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Password=*") Then
            strPass = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Data Source=*") Then
            strServer = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Server=*") Then
            strServer = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
            strServer = Replace(strServer, """", "")
        End If
    Next
    
    On Error GoTo errH
    
    Shell mstrMonitor & " " & strUser & " " & strPass & " " & strServer, vbNormalFocus
    
    ExecuteMonitor = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadResponse() As Boolean
'功能：读取病案审查反馈
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngCount As Long
    Dim curDate As Date
    
    If cboUnit.ListIndex = -1 Then
        fra审查.Visible = False: LoadResponse = True: Exit Function
    End If
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    Screen.MousePointer = 11
    
    '读取当前病区的在院、出院病人，以"病案反馈记录"为准全部扫描
    strSQL = "Select Count(*) as 数量 From 病案主页 B,病案反馈记录 A" & _
        " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And A.记录状态=1" & _
        " And A.反馈对象 IN(3,4) And B.当前病区ID + 0 =[1]" & _
        " And a.反馈时间 Between [2] And [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadResponse", cboUnit.ItemData(cboUnit.ListIndex), CDate(Format(curDate - mlngMedRedDay, "yyyy-MM-dd")), CDate(Format(curDate, "yyyy-MM-dd HH:mm:ss")))
    If Not rsTmp.EOF Then lngCount = Nvl(rsTmp!数量, 0)
    
    lbl审查.Caption = mlngMedRedDay & "天内共有 " & lngCount & " 条未处理的病案审查反馈..."
    fra审查.Visible = lngCount > 0
    If Decode(lngCount, 0, 0, 1) <> Decode(Val(lbl审查.Tag), 0, 0, 1) Then
        Call picPati_Resize
    End If
    lbl审查.Tag = lngCount
    
    Screen.MousePointer = 0
    LoadResponse = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadNotify() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim blnTmp As Boolean
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    
    If cboUnit.ListIndex = -1 Then LoadNotify = True: Exit Function
        
    Screen.MousePointer = 11
    On Error GoTo errH
    blnTmp = mobjKernel.GetAdviceRemind(rsTmp, cboUnit.ItemData(cboUnit.ListIndex), IIf(optNotify(0).Value = True, UserInfo.姓名, ""))
    Screen.MousePointer = 0
    If blnTmp = False Then Exit Function
    If rsTmp Is Nothing Then Exit Function
    If rsTmp.State = adStateClosed Then Exit Function
    strTmp = ","
    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!类型编码 & ""
        Case "ZLHIS_PACS_006", "ZLHIS_PACS_007"
            'ZLHIS_PACS_006 ZLHIS_PACS_007 消息为一条消息（按医嘱项目为单位）显示一行
            If InStr(strTmp, "," & rsTmp!类型编码 & "," & rsTmp!业务标识 & ",") = 0 Then
                strTmp = strTmp & rsTmp!类型编码 & "," & rsTmp!业务标识 & ","
                Call AddReportRow(rsTmp!病人ID & "," & rsTmp!主页ID, rsTmp!病人ID, rsTmp!主页ID, Nvl(rsTmp!姓名), Nvl(rsTmp!住院号), Nvl(rsTmp!床号), Nvl(rsTmp!消息内容), _
                    rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!业务标识 & "", rsTmp!病人来源 & "", Nvl(rsTmp!险类, 0), Nvl(rsTmp!就诊病区id, 0))
            End If
        Case Else
            If InStr(strTmp, "," & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & ",") = 0 Then
                strTmp = strTmp & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & ","
                Call AddReportRow(rsTmp!病人ID & "," & rsTmp!主页ID, rsTmp!病人ID, rsTmp!主页ID, Nvl(rsTmp!姓名), Nvl(rsTmp!住院号), Nvl(rsTmp!床号), Nvl(rsTmp!消息内容), _
                    rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!业务标识 & "", rsTmp!病人来源 & "", Nvl(rsTmp!险类, 0), Nvl(rsTmp!就诊病区id, 0))
            End If
        End Select
        rsTmp.MoveNext
    Next
    rptNotify.Populate '缺省不选中任何行
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln消息语音 Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(2)
        End If
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Set mrsMsg = rsTmp
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ExecuteEditMediRec(Optional ByVal blnEditable As Boolean)
'功能：进行病案首页整理
'参数：blnEditable=是否允许编辑(有权限及签名允许的情况下)
    Dim blnReadOnly As Boolean
    
    If mlng病人ID = 0 Then Exit Sub
    
    If mPatiInfo.数据转出 Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '病案编目之后不可以整理
    If Not (CheckMecRed(mlng病人ID, mlng主页ID, Me.Caption) Or blnEditable) Then
        blnReadOnly = True
    End If
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p住院护士站, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    '非模态显示首页整理
    If Not mclsInOutMedRec.IsOpen Then
        If mclsInOutMedRec.ShowInMedRecEdit(Me, mlng病人ID, mPatiInfo.主页ID, mPatiInfo.科室ID, rptPati.SelectedRows(0).Record(col_路径状态).Value, , mstrPrivs, IIf(blnReadOnly, 1, 0), False) Then
            mstrPrePati = "": Call rptPati_SelectionChanged
        End If
    End If
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Static strPreTime2 As String
    Dim curTime As Date
    
    curTime = Now
    If gbln启用影像信息系统预约 Then
        If strPreTime2 = "" Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime2), curTime) > 300 Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            If mobjKernel.GetMsgRISReady(mlng病区ID) Then
                Call LoadNotify
            End If
        End If
    End If
    If mblnUnRefresh Then Exit Sub
    If mbln消息语音 Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                timNotify.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                timNotify.Enabled = True
            End If
        End If
    End If
    
    
    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then '使用了消息平台不用自动刷新
            Exit Sub
        End If
    End If

    '刷新病历审查提醒
    If mintNotify > 0 And rptNotify.Visible Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call LoadNotify
        End If
    End If

End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mintChange = Val(txtChange.Text)
    Call LoadPatients
End Sub

Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub

Private Sub Set诊疗项目费用设置()
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallSetClinicCharge(mlng病区ID, 1, Me, gcnOracle, glngSys, gstrDBUser, E住院调用, InStr(GetInsidePrivs(p住院护士站), ";诊疗项目费用设置;") = 0)
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'功能：结束打印事件，写入首页打印数据
    Dim strSQL As String
    
    strSQL = _
            "Zl_电子病历打印_Insert(Null,9," & mlng病人ID & "," & mPatiInfo.主页ID & ",'" & UserInfo.姓名 & "')"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'功能：进行界面字体的统一设置
'参数：blnSetMainFont  是否设置主界面字体 （用以区分子界面切换）
    If blnSetMainFont Then
        Call zlControl.SetPubFontSize(Me, mbytSize)
        Call SetpicPatiPosition
        Call SetPatiInfoCtlPos
    End If
    Select Case tbcSub.Selected.Tag
        Case "路径"
            Call mclsPath.SetFontSize(mbytSize)
        Case "医嘱"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "费用"
            Call mclsFeeQuery.SetFontSize(mbytSize)
        Case "病历"
            Call mclsEPRs.SetFontSize(mbytSize)
        Case "护理"
            Call mclsTends.SetFontSize(mbytSize)
                Case "新病历"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
        End Select
End Sub
Private Sub SetpicPatiPosition()
'功能：病人列表以及筛选条件的相关控件的位置与大小调整

    Dim i As Long
    Dim lngDistance As Long
    
    Call picPatiIn_Resize
    Call zlControl.SetPubCtrlPos(False, 0, lblDept, 10, cboUnit)
    cboUnit.Width = picPatiFilter.ScaleWidth - cboUnit.Left
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl护理条件, 10, txt护理条件)
    txt护理条件.Width = cboUnit.Width
    cmd护理条件.Width = 270
    cmd护理条件.Height = txt护理条件.Height - IIf(mbytSize = 0, 45, 60)
    cmd护理条件.Top = txt护理条件.Top + 30
    cmd护理条件.Left = txt护理条件.Left + txt护理条件.Width - cmd护理条件.Width - 30
    
    'checkBox选择框在字体增大时不会改变宽度，因此需要减去100
    lngDistance = IIf(mbytSize = 0, 10, -50)
    Call zlControl.SetPubCtrlPos(False, 0, lbl病况条件, 10, chk病况条件(0), lngDistance, chk病况条件(1), lngDistance, chk病况条件(2))
    Call zlControl.SetPubCtrlPos(False, 0, lbl出院时间, 50, cboSelectTime)
    chk结清(0).Left = cboSelectTime.Left + 0.5 * cboSelectTime.Width
    Call zlControl.SetPubCtrlPos(False, 0, chk结清(0), 20, chk结清(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl转出, 50, cmdRef)

    txtChange.Left = lbl转出.Left + Me.TextWidth("显示最近 ")
    fraChange.Left = txtChange.Left
    fraChange.Top = txtChange.Top + txtChange.Height
        
End Sub

Private Sub SetPatiInfoCtlPos()
'功能：对病人的详细信息界面的控件位置调整－－－－picInfo中的控件
    Dim lngDistance1 As Long, lngDistance2 As Long
    
    Dim lngTmp As Long
    
    lngDistance2 = 180: lngDistance1 = 10
    lngTmp = IIf(mbytSize = 0, 1080, 1270)
    
    lblPatiName(0).Top = IIf(mbytSize = 0, 190, 210)
    lbl姓名(0).Top = lblPatiName(0).Top
    
    '1.住院次数
    lblPatiName(0).Left = IIf(mbytSize = 0, 90, 110)
    lblPages.Left = lblPatiName(0).Left
    cboPages.Left = lblPatiName(0).Left
    
    Call zlControl.SetPubCtrlPos(False, 0, lblPatiName(0), lngDistance1, lblPatiName(1))
    lblPages.Top = lblPatiName(0).Top + lblPatiName(0).Height + 70
    lblPages.Width = cboPages.Width
    cboPages.Top = lblPages.Top + lblPages.Height + 15
    fraPageId.Width = cboPages.Left + cboPages.Width + 60
    fraPageId.Height = lngTmp
    
    fraInfo.Left = fraPageId.Width + fraPageId.Left + IIf(mbytSize = 0, 10, 30)
    fraInfo.Height = lngTmp
    picInfo.Height = lngTmp
    
    '2.基本信息
    lbl姓名(0).Left = lblPatiName(0).Left
    lbl过敏.Left = lbl姓名(0).Left
    lblFee(0).Left = lbl姓名(0).Left
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl姓名(0), lngDistance1, lbl姓名(1), lngDistance2, lbl护理(0), lngDistance1, lbl护理(1), lngDistance2, lbl病室(0), lngDistance1, lbl病室(1), _
            lngDistance2, lbl类型(0), lngDistance1, lbl类型(1), lngDistance2, lbl付款(0), lngDistance1, lbl付款(1), lngDistance2, lbl医保号(0), lngDistance1, lbl医保号(1))
    
    lbl过敏.Top = lbl姓名(0).Height + lbl姓名(0).Top + 90
    Call zlControl.SetPubCtrlPos(False, 0, lbl过敏, lngDistance1, cbo过敏, lngDistance2, lbl病况(0), lngDistance1, lbl病况(1), lngDistance2, lbl入院(0), lngDistance2, lblDiag(0), lngDistance1, lblDiag(1))
    lbl姓名(0).Left = lbl过敏.Left
    lblFee(0).Top = lbl过敏.Height + lbl过敏.Top + 90
    Call zlControl.SetPubCtrlPos(False, 0, lblFee(0), lngDistance1, lblFee(1), lngDistance2, lblFluid(0), lngDistance1, lblFluid(0))
    
  
    If lbl病室(0).Left <= cbo过敏.Left + cbo过敏.Width + lngDistance2 Then
        lbl病室(0).Left = cbo过敏.Left + cbo过敏.Width + lngDistance2
    End If
    lbl病况(0).Left = lbl病室(0).Left
    Call zlControl.SetPubCtrlPos(False, 0, lbl病室(0), lngDistance1, lbl病室(1), lngDistance2, lbl类型(0), lngDistance1, lbl类型(1), lngDistance2, lbl付款(0), lngDistance1, lbl付款(1), lngDistance2, lbl医保号(0), lngDistance1, lbl医保号(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl病况(0), lngDistance1, lbl病况(1), lngDistance2, lbl入院(0), lngDistance2, lblDiag(0), lngDistance1, lblDiag(1))
   
    If lbl类型(0).Left <= lbl入院(0).Left Then
        lbl类型(0).Left = lbl入院(0).Left
    Else
        lbl入院(0).Left = lbl类型(0).Left
    End If
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl类型(0), lngDistance1, lbl类型(1), lngDistance2, lbl付款(0), lngDistance1, lbl付款(1), lngDistance2, lbl医保号(0), lngDistance1, lbl医保号(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl入院(0), lngDistance1, lbl入院(1), lngDistance2, lblDiag(0), lngDistance1, lblDiag(1))
    
    If lbl医保号(0).Left >= lblDiag(0).Left Then
        If lbl付款(0).Left >= lblDiag(0).Left Then
            lblDiag(0).Left = lbl付款(0).Left
        Else
            lblDiag(0).Left = lbl医保号(0).Left
        End If
    Else
        lbl医保号(0).Left = lblDiag(0).Left
    End If
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl付款(0), lngDistance1, lbl付款(1), lngDistance2, lbl医保号(0), lngDistance1, lbl医保号(1))
    Call zlControl.SetPubCtrlPos(False, 0, lblDiag(0), lngDistance1, lblDiag(1))
    
    lblFluid(0).Left = lbl付款(0).Left
    Call zlControl.SetPubCtrlPos(False, 0, lblFluid(0), lngDistance1, lblFluid(1), lngDistance2, lblPrint(0), lngDistance1, lblPrint(1))
    
    If lblFee(1).Left + lblFee(1).Width > lblFluid(0).Left Then
        lblFluid(0).Left = lblFee(1).Left + lblFee(1).Width + lngDistance2
        Call zlControl.SetPubCtrlPos(False, 0, lblFluid(0), lngDistance1, lblFluid(1), lngDistance2, lblPrint(0), lngDistance1, lblPrint(1))
    End If
    
    lblFee(1).Top = lblFee(0).Top
    lblFluid(0).Top = lblFee(0).Top
    lblFluid(1).Top = lblFee(0).Top
    lblPrint(0).Top = lblFee(0).Top
    lblPrint(1).Top = lblFee(0).Top
    If Not lblFluid(0).Visible Then
        lblPrint(0).Left = lblFluid(0).Left
        lblPrint(1).Left = lblPrint(0).Left + lblPrint(0).Width + lngDistance1
    End If
End Sub

Private Function GetDataToUnits(Optional ByVal strIn As String = "") As ADODB.Recordset
'功能：获取科室列表数据记录集
'参数：strIn 过滤条件
    Dim strSQL As String
    Dim blnYN As Boolean
    
    If strIn <> "" Then blnYN = True
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
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
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            IIf(blnYN, " And (C.编码 Like [2] Or C.简码 Like [3] Or C.名称 Like [3])", "") & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    On Error GoTo errH
    If blnYN Then
        Set GetDataToUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%")
    Else
        Set GetDataToUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckBabyInOut() As Boolean
'功能：检查婴儿和母亲是否分离，切当前在婴儿科室,如果是在院页面才判断
    If rptPati.SelectedRows(0).GroupRow = False And tbcPati.Selected.Tag = "在院" Then
        If rptPati.SelectedRows(0).Record(COL_婴儿科室ID).Value <> 0 Then
            If rptPati.SelectedRows(0).Record(COL_婴儿科室ID).Value = cboUnit.ItemData(cboUnit.ListIndex) Or rptPati.SelectedRows(0).Record(COL_婴儿病区ID).Value = cboUnit.ItemData(cboUnit.ListIndex) Then
                MsgBox "该病人已经转出本科室了，只有婴儿留在本科室，不允许操作病人。", vbInformation, Me.Caption
                CheckBabyInOut = True
            End If
        End If
    End If
End Function

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'功能：将接收到的消息加入提醒列表中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    Dim blnAdd As Boolean '是否要添加一行
    
    On Error GoTo errH
    
    If Mid(rsMsg!提醒场合, 3, 1) <> "1" Then Exit Sub
    
    If InStr("," & rsMsg!部门IDs & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Or _
        InStr("," & rsMsg!提醒人员 & ",", "," & UserInfo.姓名 & ",") > 0 Then
        
        '判断列表是否已经有这类消息了，不放 AddReportRow 中判断，这样可能会减少一次SQL查询
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_消息).Value = rsMsg!类型编码 And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!病人ID & "," & rsMsg!就诊id) Then
                    blnAdd = True
                End If
            End If
        Next
        
        '如果是紧急的新开和新停医嘱要替换以前的消息
        If blnAdd Then
            If InStr(",ZLHIS_CIS_001,ZLHIS_CIS_002,", "," & rsMsg!类型编码 & ",") > 0 And Val(rsMsg!优先程度 & "") = 2 Then
                blnAdd = False
                rptNotify.Records.RemoveAt i
            End If
        End If
        
        If blnAdd Then Exit Sub
        
        strSQL = "Select a.住院号, a.姓名, a.性别, a.年龄, a.当前床号 As 床号, a.险类 From 病人信息 A Where a.病人id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!病人ID))
        
        Call AddReportRow(rsMsg!病人ID & "," & rsMsg!就诊id, rsMsg!病人ID, rsMsg!就诊id, Nvl(rsTmp!姓名), Nvl(rsTmp!住院号), Nvl(rsTmp!床号), Nvl(rsMsg!消息内容), _
             rsMsg!类型编码 & "", rsMsg!优先程度 & "", Format(rsMsg!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!业务标识 & "", rsMsg!病人来源 & "", Nvl(rsTmp!险类, 0))
        
        rptNotify.Populate
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ParamArray arrInput() As Variant)
'功能：向消息提配列表中增加一行
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objItemIcon As ReportRecordItem
    Dim strRowID As String '提醒列表行的唯一标识，"病人id,主页id,消息编码"
    Dim strNO As String
    Dim str业务 As String
    Dim str病人来源 As String
    Dim int优先级 As Integer
    Dim int险类 As Integer
    Dim Index As Integer
    
    On Error GoTo errH

    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tag值
    Set objItem = objRecord.AddItem(""): objItem.Icon = 1
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '病人id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '就诊id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '姓名
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '住院号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '床号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '状态，内容
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '消息编号
    objRecord.AddItem strNO: Index = Index + 1
    
    int优先级 = Val(arrInput(Index))                     '序号
    objRecord.AddItem int优先级: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '日期
    
    str业务 = arrInput(Index): Index = Index + 1              '业务标识
    str病人来源 = arrInput(Index): Index = Index + 1          '病人来源
    int险类 = arrInput(Index)
    
    If InStr(",ZLHIS_PACS_005,ZLHIS_LIS_003,", "," & strNO & ",") > 0 Then '危机值消息特殊处理，阅读时触发消息
        objRecord.AddItem str业务 & "," & Val(str病人来源)
    Else
        objRecord.AddItem str业务
    End If
    
    Index = Index + 1
    objRecord.AddItem Val(arrInput(Index))    '病区ID
    
    If (strNO = "ZLHIS_CIS_001" Or strNO = "ZLHIS_CIS_002") And int优先级 = 2 Then objItemIcon.Icon = 18
    
    If int优先级 > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int优先级 = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
    End If
    '保险病人用红色显示
    If int险类 > 0 And int优先级 <> 3 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            objRecord.Item(Index).ForeColor = &HC0&
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReadAndSendMsg(ByVal strNO As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng就诊病区ID As Long)
    '功能：新开消息时，该消息的病人已经不再当前病区，则先将消息设为已读，再重新发送消息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL() As String
    Dim lng当前科室ID As Long
    Dim lng当前病区ID As Long
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    
    If strNO <> "ZLHIS_CIS_001" Then Exit Sub
    
    strSQL = "select nvl(A.当前科室ID,0) as 当前科室ID, nvl(A.当前病区ID,0) as 当前病区ID from 病人信息 A where A.病人ID = [1] and 主页ID = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)

    If rsTmp.EOF Then Exit Sub
    
    lng当前科室ID = Val(rsTmp!当前科室id)
    lng当前病区ID = Val(rsTmp!当前病区ID)
    
    If lng就诊病区ID <> lng当前病区ID And lng当前病区ID <> 0 Then
        
        If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
            '设置消息为已读
            strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "','0010','" & _
            UserInfo.姓名 & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            gcnOracle.CommitTrans: blnTrans = False
        Else
            strSQL = "select A.消息内容, A.提醒场合,A.类型编码,A.业务标识,A.优先程度 From 业务消息清单 A Where a.病人id=[1] And a.就诊id=[2] And a.类型编码 =[3] and a.就诊病区ID =[4]  And a.是否已阅=0 And Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, strNO, lng就诊病区ID)
            If rsTmp.RecordCount > 0 Then
                For i = 0 To rsTmp.RecordCount - 1
                    ReDim Preserve arrSQL(i)
                    arrSQL(UBound(arrSQL)) = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng主页ID & "," & lng当前科室ID & "," & lng当前病区ID & ",2,'" & rsTmp!消息内容 & "','" & rsTmp!提醒场合 & "','" & rsTmp!类型编码 & "','" & rsTmp!业务标识 & "'," & rsTmp!优先程度 & ",0,null," & lng当前病区ID & ",null)"
                Next
            End If
            
            '设置消息为已读
            strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "','0010','" & _
            UserInfo.姓名 & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
            
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '重新发送消息
            If UBound(arrSQL) <> -1 Then
                For i = 0 To UBound(arrSQL)
                    zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
                Next
            End If
            gcnOracle.CommitTrans: blnTrans = False
        End If
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mclsAdvices_DoByAdvice(ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
'功能：对医嘱记帐  lngWayID＝conMenu_Edit_AdvicePrice
    Dim lngTmp As Long, lng科室ID As Long
    lngTmp = IIf(lng相关ID = 0, lng医嘱ID, lng相关ID)
    lng科室ID = Val("" & rptPati.SelectedRows(0).Record(col_科室ID).Value)
    Call mclsFeeQuery.zlPatiBilling(Me, mlng病人ID, mlng病区ID, mlng主页ID, lng科室ID, False, lngTmp)
End Sub

Private Function GetNoPatiWhy(ByVal str住院号 As String) As String
'功能：分析出院病人未显示的原因，根住院号过滤出院病人
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim strMsg As String
    Dim str费用情况 As String
    
    On Error GoTo errH
    
    strSQL = "Select b.病人id,b.主页id,b.姓名,b.出院日期,b.当前病区id,b.封存时间,Nvl(b.病案状态,0) as 病案状态,c.编码||'-'||c.名称 as 病区" & vbNewLine & _
        "From 病人信息 A, 病案主页 B,部门表 C Where a.病人id = b.病人id and b.当前病区id=c.id And b.住院号 =[1] order by b.主页ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str住院号, cboUnit.ItemData(cboUnit.ListIndex))
    
    If rsTmp.EOF Then
        strMsg = "根据当前输入的住院号未找到任何病人，原因：输入了一个错误的住院号。"
    Else
        strMsg = "根据当前输入的住院号找到以下病人(ID:" & rsTmp!病人ID & ")，未显示原因：" & vbCrLf
        For i = 1 To rsTmp.RecordCount
            strTmp = "病人""" & rsTmp!姓名 & """第" & rsTmp!主页ID & "次住院""" & rsTmp!病区 & """"
            If IsNull(rsTmp!出院日期) Then
                strTmp = strTmp & "未出院，当前正处于住院状态。"
            ElseIf Val(rsTmp!当前病区ID & "") <> cboUnit.ItemData(cboUnit.ListIndex) Then
                strTmp = strTmp & "已出院，不属于当前病区。"
            ElseIf Not IsNull(rsTmp!封存时间) Or Val(rsTmp!病案状态 & "") = 5 Then
                strTmp = strTmp & "已出院，病案已封存归档。"
            End If
            strMsg = strMsg & strTmp & vbCrLf
            rsTmp.MoveNext
        Next
    End If
    GetNoPatiWhy = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
