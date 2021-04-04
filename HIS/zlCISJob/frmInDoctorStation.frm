VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmInDoctorStation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "住院医生工作站"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15765
   Icon            =   "frmInDoctorStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleMode       =   0  'User
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTBPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   9840
      ScaleHeight     =   5925
      ScaleWidth      =   5145
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1710
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptTBPati 
         Height          =   5475
         Left            =   0
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   0
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9657
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   4530
         Picture         =   "frmInDoctorStation.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "取消"
         Top             =   5550
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   3990
         Picture         =   "frmInDoctorStation.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "确认"
         Top             =   5550
         Width           =   450
      End
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1230
      ScaleHeight     =   975
      ScaleWidth      =   1350
      TabIndex        =   68
      Top             =   6225
      Visible         =   0   'False
      Width           =   1350
      Begin XtremeReportControl.ReportControl rptNotify 
         Height          =   630
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
         _Version        =   589884
         _ExtentX        =   1085
         _ExtentY        =   1111
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.Timer timNotify 
      Interval        =   500
      Left            =   675
      Top             =   6585
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6480
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   52
      Top             =   4920
      Width           =   855
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5220
      Left            =   4245
      TabIndex        =   2
      Top             =   2415
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9208
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8070
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInDoctorStation.frx":109E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23283
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
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   3615
      ScaleHeight     =   1320
      ScaleWidth      =   11550
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   345
      Width           =   11550
      Begin VB.Frame fraInfo 
         Height          =   1335
         Left            =   1320
         TabIndex        =   11
         Top             =   -60
         Width           =   9495
         Begin VB.ComboBox cbo过敏 
            Height          =   300
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   12
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
            TabIndex        =   64
            Top             =   735
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblFluid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   1980
            TabIndex        =   59
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblFluid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输液量:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   930
            TabIndex        =   58
            Top             =   735
            Width           =   630
         End
         Begin VB.Label lblFee 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   600
            TabIndex        =   57
            Top             =   750
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
            TabIndex        =   31
            Top             =   165
            Width           =   105
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
            TabIndex        =   30
            Top             =   165
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
            TabIndex        =   29
            Top             =   165
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
            TabIndex        =   28
            Top             =   450
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
            TabIndex        =   27
            Top             =   450
            Width           =   90
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
            TabIndex        =   26
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl护理 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "护理:"
            Height          =   180
            Index           =   0
            Left            =   2055
            TabIndex        =   25
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl病况 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病况:"
            Height          =   180
            Index           =   0
            Left            =   4065
            TabIndex        =   24
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl入院 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院:"
            Height          =   180
            Index           =   0
            Left            =   5115
            TabIndex        =   23
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl医保号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医保号:"
            Height          =   180
            Index           =   0
            Left            =   8595
            TabIndex        =   22
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lbl付款 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款:"
            Height          =   180
            Index           =   0
            Left            =   7060
            TabIndex        =   21
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号:"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   150
            Width           =   630
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
            TabIndex        =   19
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl类型 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型:"
            Height          =   180
            Index           =   0
            Left            =   4140
            TabIndex        =   18
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lbl病室 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床位:"
            Height          =   180
            Index           =   0
            Left            =   3465
            TabIndex        =   17
            Top             =   165
            Width           =   450
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
            TabIndex        =   16
            Top             =   165
            Width           =   105
         End
         Begin VB.Label lbl过敏 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过敏药物:"
            Height          =   180
            Left            =   75
            TabIndex        =   15
            Top             =   450
            Width           =   810
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   7545
            TabIndex        =   14
            Top             =   450
            Width           =   90
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "诊断:"
            Height          =   180
            Index           =   0
            Left            =   7060
            TabIndex        =   13
            Top             =   450
            Width           =   450
         End
      End
      Begin VB.Frame fraPageId 
         Height          =   1335
         Left            =   15
         TabIndex        =   8
         Top             =   -45
         Width           =   1275
         Begin VB.ComboBox cboPages 
            Height          =   300
            Left            =   45
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   660
            Width           =   1155
         End
         Begin VB.Image imgCurPati 
            Height          =   240
            Index           =   2
            Left            =   1080
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgCurPati 
            Height          =   240
            Index           =   1
            Left            =   765
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgCurPati 
            Height          =   240
            Index           =   0
            Left            =   465
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblCurPati 
            AutoSize        =   -1  'True
            Caption         =   "图标:"
            Height          =   180
            Left            =   -45
            TabIndex        =   77
            Top             =   1110
            Width           =   450
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
            TabIndex        =   63
            Top             =   135
            Width           =   390
         End
         Begin VB.Label lblPatiName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   62
            Top             =   120
            Width           =   435
         End
         Begin VB.Label lblPages 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "次数:"
            Height          =   180
            Left            =   60
            TabIndex        =   0
            Top             =   465
            Width           =   450
         End
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5505
      Left            =   105
      ScaleHeight     =   5505
      ScaleWidth      =   4050
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   675
      Width           =   4050
      Begin VB.CheckBox chkFilter 
         Height          =   255
         Left            =   3240
         Picture         =   "frmInDoctorStation.frx":1930
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "按照住院号对病人进行精确过滤显示"
         Top             =   960
         Width           =   270
      End
      Begin VB.PictureBox picIconPati 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   150
         ScaleHeight     =   300
         ScaleWidth      =   1845
         TabIndex        =   70
         Top             =   855
         Width           =   1845
         Begin VB.Label lblBJ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "标记"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   76
            Top             =   60
            Width           =   360
         End
         Begin VB.Label lblCountThis 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(8)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   71
            Top             =   105
            Width           =   270
         End
         Begin VB.Image imgIconPati 
            Height          =   240
            Index           =   0
            Left            =   540
            Picture         =   "frmInDoctorStation.frx":8182
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.TextBox txtTestBug 
         Height          =   270
         Left            =   -550
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   120
         Width           =   180
      End
      Begin VB.PictureBox picPatiIn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3855
         ScaleWidth      =   4365
         TabIndex        =   33
         Top             =   1380
         Width           =   4365
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   2580
            Left            =   0
            TabIndex        =   34
            Top             =   1680
            Width           =   3360
            _Version        =   589884
            _ExtentX        =   5927
            _ExtentY        =   4551
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
            Height          =   345
            Index           =   0
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   4215
            TabIndex        =   35
            Top             =   10
            Visible         =   0   'False
            Width           =   4215
            Begin VB.CheckBox chkByTeam 
               Caption         =   "按小组显示"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2640
               TabIndex        =   49
               ToolTipText     =   "是否按医疗小组模式显示病人列表"
               Top             =   23
               Width           =   1280
            End
            Begin VB.CheckBox chk病况条件 
               Caption         =   "重"
               Height          =   195
               Index           =   2
               Left            =   2070
               TabIndex        =   36
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   23
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox chk病况条件 
               Caption         =   "危"
               Height          =   195
               Index           =   1
               Left            =   1500
               TabIndex        =   38
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   23
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox chk病况条件 
               Caption         =   "一般"
               Height          =   195
               Index           =   0
               Left            =   750
               TabIndex        =   37
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   23
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.Label lbl病况条件 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病 况(&S)"
               Height          =   180
               Left            =   0
               TabIndex        =   39
               Top             =   30
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
            TabIndex        =   44
            Top             =   970
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CommandButton cmdRef 
               Caption         =   "刷新"
               Height          =   255
               Left            =   2520
               TabIndex        =   48
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
               TabIndex        =   46
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
               TabIndex        =   45
               Text            =   "7"
               Top             =   30
               Width           =   285
            End
            Begin VB.Label lbl转出 
               AutoSize        =   -1  'True
               Caption         =   "显示最近    天的转出病人"
               Height          =   180
               Left            =   0
               TabIndex        =   47
               Top             =   30
               Width           =   2160
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
            TabIndex        =   42
            Top             =   650
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CheckBox chkHZ 
               Caption         =   "已会诊"
               Height          =   180
               Index           =   1
               Left            =   3225
               TabIndex        =   67
               Top             =   120
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox chkHZ 
               Caption         =   "未会诊"
               Height          =   180
               Index           =   0
               Left            =   2160
               TabIndex        =   66
               Top             =   90
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox chkOut 
               Caption         =   "包含出院病人"
               Height          =   195
               Left            =   2760
               TabIndex        =   53
               ToolTipText     =   "Ctrl+勾选：单独选择"
               Top             =   60
               Width           =   1500
            End
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   2
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   20
               Width           =   1230
            End
            Begin VB.Label lbl开始时间 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "开始时间"
               Height          =   180
               Left            =   0
               TabIndex        =   43
               Top             =   60
               Width           =   720
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
            TabIndex        =   40
            Top             =   330
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CheckBox chkOutByTeam 
               Caption         =   "按小组显示"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2340
               TabIndex        =   65
               ToolTipText     =   "是否按医疗小组模式显示病人列表"
               Top             =   60
               Width           =   1280
            End
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   1
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   20
               Width           =   1230
            End
            Begin VB.Label lbl出院时间 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               Height          =   180
               Left            =   0
               TabIndex        =   41
               Top             =   60
               Width           =   720
            End
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPati 
         Height          =   600
         Left            =   135
         TabIndex        =   32
         Top             =   930
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   1058
         _StockProps     =   64
      End
      Begin VB.Frame fra审查 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   5190
         Visible         =   0   'False
         Width           =   3360
         Begin VB.Image Image1 
            Height          =   240
            Left            =   105
            Picture         =   "frmInDoctorStation.frx":E9D4
            Top             =   45
            Width           =   240
         End
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
            MouseIcon       =   "frmInDoctorStation.frx":EF5E
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   75
            Width           =   3060
         End
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   900
         TabIndex        =   4
         Text            =   "cboDept"
         Top             =   120
         Width           =   2655
      End
      Begin MSComctlLib.ImageList imgPati 
         Left            =   255
         Top             =   345
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   21
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":F0B0
               Key             =   "Pati"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":F64A
               Key             =   "Meet"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":FBE4
               Key             =   "MeetFinish"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":1017E
               Key             =   "Notify"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":10718
               Key             =   "等待审查"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":10CB2
               Key             =   "拒绝审查"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":1124C
               Key             =   "正在审查"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":117E6
               Key             =   "正在抽查"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":121F8
               Key             =   "审查反馈"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":12C0A
               Key             =   "抽查反馈"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":131A4
               Key             =   "审查整改"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":13BB6
               Key             =   "抽查整改"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":145C8
               Key             =   "未导入"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":14B62
               Key             =   "变异结束"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":150FC
               Key             =   "正常结束"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":15696
               Key             =   "不符合"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":160A8
               Key             =   "执行中"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":16642
               Key             =   "Child"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":16BDC
               Key             =   "单病种"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":1D43E
               Key             =   "Out"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":1D9D8
               Key             =   "Fbaby"
            EndProperty
         EndProperty
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   870
         TabIndex        =   54
         Top             =   480
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
         IDKindStr       =   $"frmInDoctorStation.frx":1DF72
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
         Left            =   360
         TabIndex        =   55
         Top             =   525
         Width           =   735
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室(&D)↓"
         Height          =   180
         Left            =   135
         TabIndex        =   3
         Top             =   180
         Width           =   810
      End
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
      Bindings        =   "frmInDoctorStation.frx":1E039
      Left            =   705
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmInDoctorStation"
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
    col_费别 = 12
    col_科室 = 13
    col_病区 = 14
    col_住院医师 = 15
    col_入院日期 = 16  '取入科时间 兼容以前数据:入科时间 为空时取入院日期
    col_出院日期 = 17
    col_病人类型 = 18
    
    col_医嘱ID = 19
    col_发送号 = 20
    col_执行状态 = 21
    col_执行科室ID = 22
    col_就诊卡 = 23
    col_住院天数 = 24
    col_单病种 = 25
    COL_婴儿科室ID = 26
    COL_婴儿病区ID = 27
    col_西医诊断 = 28
    col_中医诊断 = 29
    COL_申请序号 = 30
    COL_传染病 = 31
    col_责任护士 = 32
    col_留观号 = 33
    col_身份证号 = 34
    col_是否急诊 = 35

End Enum

Private Enum NOTIFYREPORT_COLUMN
    c_图标 = 0
    C_病人Id = 1
    C_主页Id = 2
    c_姓名 = 3
    c_住院号 = 4
    c_床号 = 5
    C_状态 = 6
    '隐藏列
    C_消息 = 7
    C_序号 = 8
    C_日期 = 9
    C_业务 = 10
    C_Id = 11
End Enum

' 点图标后弹出病人选择器
Private Enum PATI_COLUMN
    CI_图标1 = 0
    CI_图标2
    CI_图标3
    CI_床号
    CI_病人ID
    CI_主页ID
    CI_姓名
    CI_住院号
    CI_入院日期
    CI_出院日期
    CI_病人类型
End Enum

Private Enum PATI_TYPE
    pt我的 = 1
    pt在院 = 2
    pt预出 = 3
    pt出院 = 4
    pt死亡 = 5
    pt会诊 = 6
    pt最近转出 = 7
End Enum

Private Enum Msg_Type '消息提醒类别
    m病历审阅 = 1
    m医嘱安排 = 2
    m危机值 = 3
    m报告撤消 = 4
    m医嘱审核 = 5
    m处方审查 = 6
    m传染病 = 7
    m病历质控 = 8
    m输血完成 = 9
    m校对疑问 = 10
    m用血审核 = 11
    m输血反应 = 12
End Enum

Private Type PatiInfo
    状态 As Integer '病案主页.状态
    婴儿 As Integer
    住院号 As String
    床号 As String
    病人ID As Long
    主页ID As Long
    病区ID As Long
    科室ID As Long
    入院日期 As Date
    出院日期 As Date
    编目日期 As Date
    住院次数 As Long
    rs图标 As ADODB.Recordset
    数据转出 As Boolean
End Type

'子窗体对象定义
Private mclsEMR As Object  '新版病历zlRichEMR.clsDockEMR
Private mclsDisease As zlRichEPR.cDockDisease
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockInEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zlRichEPR.cDockInTends
Attribute mclsTends.VB_VarHelpID = -1
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '新版护士工作站
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mclsTendEPRs As zlRichEPR.cDockInTendEPRs
Private WithEvents mclsPath As zlPublicPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mclsWardMonitor As clsWardMonitor     '监护仪接口
Private mobjEPRDoc As zlRichEPR.cEPRDocument
Private mclsChildQuestion As zlRichEPR.clsChildQuestion '病案审查接口
Private mobjKernel As zlPublicAdvice.clsPublicAdvice          '临床核心部件
Private WithEvents mclsDis As zl9Disease.clsDisease
Attribute mclsDis.VB_VarHelpID = -1

Private WithEvents mFrmConsultation As Form
Attribute mFrmConsultation.VB_VarHelpID = -1

Private WithEvents mclsInOutMedRec As zlMedRecPage.clsInOutMedRec  '首页窗体
Attribute mclsInOutMedRec.VB_VarHelpID = -1
Private WithEvents mfrmResponse As frmAuditResponse '审查反馈窗口
Attribute mfrmResponse.VB_VarHelpID = -1
Private WithEvents mfrmInView As frmInDoctorView    '住院一览
Attribute mfrmInView.VB_VarHelpID = -1
Private mcolSubForm As Collection
Private mfrmActive As Form
Private mobjSquareCard As Object      '卡结算对象
Private mstrCardKind As String        '卡结算对象返回的可用的医疗卡
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsReg As zlPublicExpense.clsRegist

'参数设置变量
Private mintChange As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mdtMeetBegin As Date, mdtMeetEnd As Date
Private mintNotify As Integer '审阅提醒自动刷新间隔(分钟)
Private mstrNotify As String  '提醒区域内容
Private mintNotifyDay As Integer '提醒多少天内完成的病历
Private mintDeptView As Integer '0-按科室显示，1-按病区显示
Private mintDeptViewBed As Integer '0，1-只显示有床位的病区或者科室
Private mblnDeptViewEnabled As Boolean
Private mlngMedRedDay As Long   '病案审查反馈天数
Private mintMecStandard As Integer  '病案首页格式 0-卫生部标准，1-四川省标准，2-云南省标准,3-湖南省标准
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln消息语音 As Boolean
Private mbln危急值弹窗 As Boolean

Private mstrAllPatis As String '当前列表中的病人信息，格式："病人ID:主页ID,病人ID:主页ID,..."
Private mrsNotes As ADODB.Recordset '病人个性图标，可设置图标
Private mstrList主题 As String
Private mrsPatiNotes As ADODB.Recordset '某一批病人的图标
Private mrsPati汇总 As ADODB.Recordset '图标汇总
Private mlngSource As Long '大字体，小字体

Private Const conMenu_图标 = 990050                     '标注所使用的图标ID从990050开始,最多150个图标
Private Const conMenu_标注1 = 990200
Private Const conMenu_标注2 = 990300
Private Const conMenu_标注3 = 990400
Private Const conMenu_标注结束 = 990500

Private Const conIconAll = 50 '所有图标数,上限暂定为50个

'其它窗体变量
Private mstrPrivs As String
Private mlngModul As Long
Private mPatiInfo As PatiInfo '历史住院记录中的,不一定为当前的
Private mlng病人ID As Long, mlng主页ID As Long '病人清单中的
Private mrsPati As ADODB.Recordset '病人信息集合，包含同一身份证号的所有病人
Private mobjPatient As Object '病人信息公共部件，验证身份证号
Private mintFindType As Integer '0-住院号,1-床号,2-就诊卡,3-姓名
Private mstrFindType As String '用来存储当前查找类型的名称
Private mblnFindTypeEnabled As Boolean
Private mblnICU As Boolean '是否非本科的ICU室
Private mblnOutDept As Boolean '是否仅服务于门诊的科室（门诊留观病人显示门诊号）
Private mstrDiagInfo As String  '首页整理时新填入的第一诊断信息
Private mstr疾病ID As String   '用于从子窗体获取疾病ID
Private mstr诊断ID As String   '用于从子窗体获取诊断ID
Private mblnReturn As Boolean       'cboDept回车按键
Private mintOutPreTime As Integer
Private mintMeetPreTime As Integer

Private mintPreDept As Integer
Private mstrPrePati As String
Private mintPrePage As Integer
Private mstrPreNotify As String
Private mblnUnRefresh As Boolean
Private mblnNoCheck  As Boolean '病况选择
Private mfrmParent As Object
Private mblnIsFindAgain As Boolean
Private mstrUserDeps As String '操作员所属科室字符串
Private mlngNewIndex As Long
Private mlngOldIndex As Long
Private mblnIsNot As Boolean
Private mbytSize As Byte '字体大小 0-小字体（9号）1-大字体（12号）
Private mblnTabTmp As Boolean
Private mblnIsInit As Boolean
Private mblnInView As Boolean
Private mbln接受会诊 As Boolean
Private mbln危急值 As Boolean '处危急值的权限
Private mlng会诊医嘱ID As Long
Private mbln单个病人 As Boolean

Public Sub ShowMe(frmParent As Object)
    
    Set mfrmParent = frmParent
    Me.Show , frmParent
End Sub

Private Sub chkByTeam_Click()
    Call LoadPatients
End Sub

Private Sub chkOutByTeam_Click()
    Call LoadPatients
End Sub

Private Sub chkOut_Click()
    '重新读取病人
    Call LoadPatients
End Sub

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

Private Sub chkHZ_Click(Index As Integer)
    Dim i As Integer, k As Integer
    
    If Not Visible Or mblnNoCheck Then Exit Sub
    
    If (GetKeyState(vbKeyControl) And &H8000) <> 0 Then
        'Ctrl：排它选择
        mblnNoCheck = True
        For i = 0 To chkHZ.UBound
            chkHZ(i).Value = IIf(i = Index, 1, 0)
        Next
        mblnNoCheck = False
    Else
        '至少选择一个
        For i = 0 To chkHZ.UBound
            If chkHZ(i).Value = 1 Then k = k + 1
        Next
        If k = 0 Then chkHZ(Index).Value = 1
    End If
    
    '重新读取病人
    Call LoadPatients
End Sub

Private Sub chkFilter_Click()
    PatiIdentify.Text = ""
    If PatiIdentify.Visible And PatiIdentify.Enabled Then PatiIdentify.SetFocus
    Call LoadPatients
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtOutEnd = datCurr
    mdtOutBegin = mdtOutEnd - 1
    mdtMeetEnd = datCurr
    mdtMeetBegin = mdtMeetEnd - 1
    
    cboSelectTime(1).Clear '出院
    With cboSelectTime(1)
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
    If cboSelectTime(1).ListCount > 0 Then cboSelectTime(1).ListIndex = 0
    
    cboSelectTime(2).Clear '会诊
    With cboSelectTime(2)
        .AddItem "今天内"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天内"
        .ItemData(.NewIndex) = 1
        .AddItem "前天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "15天内"
        .ItemData(.NewIndex) = 15
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime(2).ListCount > 0 Then cboSelectTime(2).ListIndex = 1
End Sub

Private Sub cboSelectTime_Click(Index As Integer)
'Index 1出院 2会诊
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Index = 1 Then
        If cboSelectTime(Index).ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime(1)) Then
                '取消时恢复原来的选择
                Call Cbo.SetIndex(cboSelectTime(Index).hwnd, mintOutPreTime)
                Exit Sub
            End If
        Else
            mdtOutEnd = datCurr
            mdtOutBegin = mdtOutEnd - intDateCount
        End If
        If mdtOutBegin = CDate(0) Or mdtOutEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "范围：" & Format(mdtOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdtOutEnd, "yyyy-MM-dd")
        End If
        '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
        Call zlDatabase.SetPara("出院病人结束间隔", DateDiff("d", datCurr, mdtOutEnd), glngSys, p住院医生站)
        Call zlDatabase.SetPara("出院病人开始间隔", DateDiff("d", mdtOutBegin, datCurr), glngSys, p住院医生站)
        mintOutPreTime = cboSelectTime(Index).ListIndex
    ElseIf Index = 2 Then
        If cboSelectTime(Index).ListIndex = mintMeetPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdtMeetBegin, mdtMeetEnd, cboSelectTime(2)) Then
                '取消时恢复原来的选择
                Call Cbo.SetIndex(cboSelectTime(Index).hwnd, mintMeetPreTime)
                Exit Sub
            End If
        Else
            mdtMeetEnd = datCurr
            mdtMeetBegin = mdtMeetEnd - intDateCount
        End If
        If mdtMeetBegin = CDate(0) Or mdtMeetEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "范围：" & Format(mdtMeetBegin, "yyyy-MM-dd") & " 至 " & Format(mdtMeetEnd, "yyyy-MM-dd")
        End If
        mintMeetPreTime = cboSelectTime(Index).ListIndex
    End If
    If Me.Visible = True Then Call LoadPatients
End Sub

Private Sub cmdRef_Click()
'输入转出天数后刷新
    Call txtChange_KeyPress(vbKeyReturn)
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
    Dim blnCol As Boolean, intType As Integer, bln路径状态 As Boolean
    Dim strTmp As String, i As Integer
    Dim arrTmp As Variant, objTabItem As TabControlItem
    Dim objTimeLine As Object
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mblnNoCheck = False '病况选择
    
    mblnICU = False
    mintPreDept = -1
    mstrPrePati = ""
    mintPrePage = -1
    mstrPreNotify = ""
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, p住院医生站, GetInsidePrivs(p住院医生站))
    Call AddMipModule(mclsMipModule)
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    Set mclsDis = New zl9Disease.clsDisease
    Call mclsDis.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, mclsMipModule)
    
    Call GetLocalSetting '本地参数
    
    '图标
    Call SetAllPati图标
    
    '界面恢复：默认搜索类型读取
    '-----------------------------------------------------
    mintFindType = Val(zlDatabase.GetPara("病人查找方式", glngSys, p住院医生站, , , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    mintDeptViewBed = Val(zlDatabase.GetPara("不显示无床位的病区科室", glngSys, p住院医生站, , , , intType))
    mbln危急值 = InStr(GetInsidePrivs(p住院医生站), ";危急值处理;") > 0
    
    Set mclsReg = New zlPublicExpense.clsRegist
    Call mclsReg.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Call mclsReg.zlInitData(2)
    
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    mstrCardKind = "住|住院号|0|0|0|0|0|0;床|床号|0|0|0|0|0|0;就|就诊卡|0|0|8|0|0|0;姓|姓名|0|0|0|0|0|0;留|留观号|0|0|0|0|0|0;诊|病人诊断|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
        Set mobjSquareCard = Nothing
        MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
    End If
    
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
    Set objPane = Me.dkpMain.CreatePane(1, IIf(mbytSize = 0, 310, 320), 400, DockLeftOf, Nothing)
    objPane.Title = "住院病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 310, 100, DockBottomOf, objPane)
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
    Set mclsDisease = New zlRichEPR.cDockDisease
    Set mclsTends = New zlRichEPR.cDockInTends
    Set mclsWardMonitor = New clsWardMonitor
    Set mclsPath = New zlPublicPath.clsDockPath
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    Call mclsAdvices.zlInitPath(mclsPath)
    Set mclsTendEPRs = New zlRichEPR.cDockInTendEPRs
    
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_新病历"
    End If
    mcolSubForm.Add mclsPath.zlGetForm, "_路径"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_病历"
    mcolSubForm.Add mclsTends.zlGetForm, "_护理"
    mcolSubForm.Add mclsTendEPRs.zlGetForm, "_护理病历"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_监护"
    End If
    mcolSubForm.Add mclsTendsNew.zlGetForm, "_新版护理"
    mcolSubForm.Add mclsDisease.zlGetForm, "_疾病报告"
    
    If InStr(GetInsidePrivs(p住院医生站), "住院一览") > 0 Then
        Set objTimeLine = DynamicCreate("ZLSoft.BusinessHome.ClientControl.TimeLineBase.Control.TimeLineControl", "时间轴", False)
        If Not objTimeLine Is Nothing Then
            Set mfrmInView = New frmInDoctorView
             mcolSubForm.Add mfrmInView, "_住院一览"
        End If
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
        If Not mfrmInView Is Nothing Then
            .InsertItem(intIdx, "住院一览", picTmp.hwnd, 0).Tag = "住院一览": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p临床路径应用, True) <> "" Then
            .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0).Tag = "路径": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院医嘱下达, True) <> "" Then
            .InsertItem(intIdx, "医嘱信息", picTmp.hwnd, 0).Tag = "医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院病历管理, True) <> "" Then
            .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0).Tag = "病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p新版住院病历, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0).Tag = "新病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p护理记录管理, True) <> "" Then
            .InsertItem(intIdx, "护理信息", picTmp.hwnd, 0).Tag = "护理": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngOldIndex = intIdx - 1
            .InsertItem(intIdx, "护理信息", picTmp.hwnd, 0).Tag = "新版护理": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNewIndex = intIdx - 1
            .InsertItem(intIdx, "护理病历", picTmp.hwnd, 0).Tag = "护理病历": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
        End If
        If GetInsidePrivs(p疾病报告填写, True) <> "" Then
            Set objTabItem = .InsertItem(intIdx, "疾病报告", picTmp.hwnd, 0): objTabItem.Tag = "疾病报告": objTabItem.Visible = False: intIdx = intIdx + 1
        End If
        If mclsWardMonitor.Enabled Then
            If InStr(GetInsidePrivs(p住院医生站), "护理监护") > 0 Then
                .InsertItem(intIdx, "护理监护", picTmp.hwnd, 0).Tag = "监护": intIdx = intIdx + 1
            End If
        End If
        If gbln启用整体护理接口 Then
            If InitNurseIntegrate = True Then
                If Not gobjNurseIntegrate Is Nothing Then
                    mcolSubForm.Add gobjNurseIntegrate.GetDocForm, "_护理评估评分"
                    .InsertItem(intIdx, "护理评估评分", mcolSubForm("_护理评估评分").hwnd, 0).Tag = "护理评估评分": intIdx = intIdx + 1
                    .Item(intIdx - 1).Visible = False
                End If
            End If
        End If
                
        '外挂提供的卡片
        Call CreatePlugInOK(p住院医生站)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p住院医生站)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p住院医生站, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "你没有使用住院医生工作站的权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '恢复上次选择的卡片
        strTab = zlDatabase.GetPara("医护功能", glngSys, p住院医生站)
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
        .InsertItem(4, "会诊", picPatiIn.hwnd, 0).Tag = "会诊"
        
        .Item(4).Selected = True
        .Item(1).Selected = True
        '定位病人选项卡
        tbcPati.Item(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", 1)).Selected = True
    End With
    
    
    '其它界面设置
    Call InitReportColumn
    picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picInfo.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '病况选择
    chk病况条件(0).BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    chk病况条件(1).BackColor = chk病况条件(0).BackColor
    chk病况条件(2).BackColor = chk病况条件(0).BackColor
    Call Cbo.SetListWidth(cbo过敏.hwnd, cbo过敏.Width * 2)
    
    '读取界面数据
    Call Set按小组条件显示
    
     '转出病人天数
    txtChange.Text = mintChange
    mintOutPreTime = -1
    mintMeetPreTime = -1
    Call InitSelectTime
    
    '初始化住院科室/病区
    Call ReLoadDept
    
    '取操作员所属科室
    mstrUserDeps = GetUser科室IDs(False)
    
    '初始化病人过滤条件
    strTmp = zlDatabase.GetPara("当前病况过滤", glngSys, p住院医生站, "111", _
        Array(lbl病况条件, chk病况条件(0), chk病况条件(1), chk病况条件(2)), InStr(mstrPrivs, "参数设置") > 0)
    For i = 0 To chk病况条件.UBound
        chk病况条件(i).Value = IIf(Mid(strTmp, i + 1, 1) = "1", 1, 0)
    Next
    
    strTmp = zlDatabase.GetPara("会诊病人过滤", glngSys, p住院医生站, "011", Array(chkOut, chkHZ(0), chkHZ(1)), InStr(mstrPrivs, "参数设置") > 0)
    chkOut.Value = IIf(Mid(strTmp, 1, 1) = "1", 1, 0)
    chkHZ(0).Value = IIf(Mid(strTmp, 2, 1) = "1", 1, 0)
    chkHZ(1).Value = IIf(Mid(strTmp, 3, 1) = "1", 1, 0)
    If chkHZ(0).Value = 0 And chkHZ(1).Value = 0 Then
        chkHZ(0).Value = 1
        chkHZ(1).Value = 1
    End If
    
    '按小组显示
    strTmp = zlDatabase.GetPara("按小组显示", glngSys, p住院医生站, "0", Array(chkByTeam), InStr(mstrPrivs, "参数设置") > 0)
    chkByTeam.Value = IIf(strTmp = "1", 1, 0)
        
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(chkOutByTeam), "chkOutByTeam", "0")
    chkOutByTeam.Value = IIf(strTmp = "1", 1, 0)
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
    blnCol = rptPati.Columns(col_审查).Visible
    bln路径状态 = rptPati.Columns(col_路径状态).Visible
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
  
    rptPati.Columns(col_审查).Visible = blnCol
    rptPati.Columns(col_路径状态).Visible = bln路径状态
    If bln路径状态 And rptPati.Columns(col_路径状态).Width = 0 Then rptPati.Columns(col_路径状态).Width = 18

    '命令行调用，定位到指定的病人
    If Not mfrmParent Is Nothing Then
        If mfrmParent.frmHide Then Call LocatePati
    End If
    Call LoadNotify
End Sub

Private Function LocatePati(Optional ByVal strTag As String) As Boolean
'功能：定位到指定的病人
    Dim varCmd As Variant, i As Integer
    Dim lng病人ID As Long, lng主页ID As Long, lng部门ID As Long
    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    Dim lngKey As Long
    
    If strTag <> "" Then
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
        Exit Function
    End If
    
    '获取命令行参数中的病人ID，主页ID
    varCmd = Split(mfrmParent.GetCommand, " ")
    For i = LBound(varCmd) To UBound(varCmd)
        If UCase(varCmd(i)) Like "病人ID=*" Then
            lng病人ID = Val(Split(varCmd(i), "=")(1))
        ElseIf UCase(varCmd(i)) Like "主页ID=*" Then
            lng主页ID = Val(Split(varCmd(i), "=")(1))
        ElseIf UCase(varCmd(i)) Like "SINGLEPATI=*" Then
            lngKey = Val(Split(varCmd(i), "=")(1))
        End If
    Next
    mbln单个病人 = False
    If lng病人ID <> 0 Then
        lng部门ID = GetPatiDept(lng病人ID, lng主页ID, mintDeptView)
        If lng部门ID <> 0 Then Call Cbo.Locate(cboDept, lng部门ID, True)   '如果是与之前的科室相同，不会触发click事件
    
        For i = 0 To rptPati.Rows.Count - 1
            With rptPati.Rows(i)
                If Not .GroupRow Then
                    If .Record(col_病人Id).Value = lng病人ID And .Record(col_主页ID).Value = lng主页ID Then Exit For
                End If
            End With
        Next
    
        If i <= rptPati.Rows.Count - 1 Then
            '该行选中且显示在可见区域,并引发SelectionChanged事件
            Set rptPati.FocusedRow = rptPati.Rows(i)
            If rptPati.Visible Then rptPati.SetFocus
            If lngKey = 1 Then
            mbln单个病人 = True
            dkpMain.Panes(1).Closed = True
            dkpMain.Panes(2).Closed = True
            End If
        End If
        
        '定位到医嘱信息页
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub(i).Visible And tbcSub(i).Tag = "医嘱" Then
                tbcSub.Item(i).Selected = True
            End If
        Next
    End If
End Function

Private Sub Set按小组条件显示()
'功能：设置是否显示按小组显示的条件
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    If InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = "Select 1 From 临床医疗小组 Where rownum=1"
    Else
        strSQL = "Select 1 From 医疗小组人员 Where 人员id = [1] And Rownum = 1"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    chkByTeam.Visible = rsTmp.RecordCount > 0
    chkOutByTeam.Visible = rsTmp.RecordCount > 0
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReLoadDept()
'功能：按科室/病区方式读取可用部门
    lblDept.Caption = IIf(mintDeptView = 0, "科室(&D)↓", "病区(&D)↓")
   
    mintPreDept = -1
    Call InitDepts
    Call cboDept_Click
    
    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "全院病人") > 0 Then
            MsgBox "没有发现住院" & IIf(mintDeptView = 0, "科室", "病区") & "信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Else
            MsgBox "没有发现你所属" & IIf(mintDeptView = 0, "科室", "病区") & ",不能使用住院医生工作站！", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    '如果是标注菜单,执行完即退出
    If Control.ID > conMenu_标注1 And Control.ID < conMenu_标注结束 Then
        Call SetPatiIcon(Control.Parameter)
        Exit Sub
    End If
    
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
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
            mlngSource = 999
            mbytSize = 0
            Call zlDatabase.SetPara("字体", mbytSize, glngSys, p住院医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '大字体
        If mbytSize <> 1 Then
            mlngSource = 0
            mbytSize = 1
            Call zlDatabase.SetPara("字体", mbytSize, glngSys, p住院医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_Manage_Bespeak '预约挂号
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng病人ID)
        Control.Enabled = True
    Case conMenu_Edit_AppRequestManage, conMenu_Edit_AppRequest
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng病人ID)
        Control.Enabled = True
    Case conMenu_View_Option '"挂号选项设置"
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "")
        Control.Enabled = True
    Case conMenu_Tool_KssAudit '抗菌用药审核
        Call frmExamineKSS.ShowMe(Me, mclsMipModule)

    Case conMenu_Tool_OPSAudit '手术审核管理
        Call frmExamineOPS.ShowMe(Me, mclsMipModule)

    Case conMenu_Tool_OPSEmpower '手术授权管理
        On Error Resume Next
        If gobjCISBase Is Nothing Then
            Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
            If gobjCISBase Is Nothing Then
                MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        err.Clear: On Error GoTo 0
        Call gobjCISBase.CallOPSEmpower(Me, gcnOracle, glngSys, gstrDBUser)
    Case conMenu_Tool_TransAudit '输血审核管理
        On Error Resume Next
        Call frmExamineTransfuse.ShowMe(Me, 2, mclsMipModule)
    Case conMenu_Tool_CISMed  '临床自管药
        Call Set临床自管药(Me)
    Case conMenu_Tool_Archive '电子病案查阅
        mblnUnRefresh = True
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.病人ID, mPatiInfo.主页ID)
        mblnUnRefresh = False
    Case conMenu_Tool_ExaReport
        '调用陈福荣那边提供的接口
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        mblnUnRefresh = True
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        mblnUnRefresh = True
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Tool_MedRatio
        On Error Resume Next
        Call frmMedRatio.ShowMe(Me, mstrPrivs)
    Case conMenu_Edit_TraReactionRecord '输血反应
        Call FuncTraReactionRecord(Me, 1, p住院医嘱下达)
    Case conMenu_Tool_MedRec '首页整理
        mblnUnRefresh = True
        Call ExecuteEditMediRec
        mblnUnRefresh = False
    Case conMenu_File_MedRecSetup '首页打印设置
        Call PrintInMedRec(mclsInOutMedRec, 0, mPatiInfo.病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case conMenu_File_MedRecPreview '首页预览
        Call PrintInMedRec(mclsInOutMedRec, 1, mPatiInfo.病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case conMenu_File_MedRecPrint '首页打印
        Call PrintInMedRec(mclsInOutMedRec, 2, mPatiInfo.病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me)
    Case conMenu_Tool_MeetIdea '填写/查看会诊意见
        Call ExecuteMeetIdea(IIf(Control.Caption = "填写会诊意见(&W)", 0, 1))
    Case conMenu_Tool_MeetOpen '接受会诊
        Call Execute接受会诊(IIf(Control.Caption = "接受会诊(&O)", False, True))
    Case conMenu_Tool_MeetFinish '完成会诊
        Call ExecuteMeetFinish
    Case conMenu_Tool_MeetCancel '取消完成
        Call ExecuteMeetCancel
    Case conMenu_Tool_MedRecAuditSubmit '提交审查
        '出院病人，尚未提交或拒绝审查状态
        Call ExecuteMedRecAuditSubmit
    Case conMenu_Tool_MedRecAuditCancel '取消提交
        '出院病人，已经提交状态
        Call ExecuteMedRecAuditCancel
    Case conMenu_Tool_MedRecAuditResponse '审查反馈
        '都可以调用，至少可以查看(当前或历史)
        Call lbl审查_Click
    Case conMenu_Tool_MedRecAuditWriteResponse '书写审查意见
        If mclsChildQuestion Is Nothing Then
            Set mclsChildQuestion = New zlRichEPR.clsChildQuestion
        End If
        If Not mclsChildQuestion Is Nothing Then
            Call mclsChildQuestion.zlOpenQuestion(Me, mlng病人ID, mlng主页ID)
        End If
    Case conMenu_View_Find '查找
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '有时需要定位一下
            If PatiIdentify.Text <> "" Then
                If chkFilter.Value = 1 And chkFilter.Visible = True Then
                    Call LoadPatients
                Else
                    Call ExecuteFindPati
                End If
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
    Case conMenu_View_Notify '审阅提醒
        If rptNotify.Visible Then Call LoadNotify
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '按科室/病区显示
        If mintDeptView <> Control.ID - conMenu_View_Dept * 10# - 1 Then
            mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
            Call zlDatabase.SetPara("部门显示方式", mintDeptView, glngSys, p住院医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            
            Call ReLoadDept
        End If
    Case conMenu_View_Refresh '刷新
        Call LoadPatients
        Call LoadNotify '刷新医嘱提醒
         
    Case conMenu_File_Parameter '参数设置
        mblnUnRefresh = True
        frmInStationSetup.mstrPrivs = mstrPrivs
        frmInStationSetup.Show 1, Me
        If gblnOK Then
            Call GetLocalSetting
            Call LoadPatients
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
    Case conMenu_Tool_HealthCard  '居民健康卡
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlHealthArchivesShow(Me, p门诊医生站, mlng病人ID, "")
        End If
    Case conMenu_Tool_Positive '阳性结果查看
        Call mclsDis.ShowRegistByPati(Me, 1, mlng病人ID, mlng主页ID)
    Case conMenu_Tool_Critical
        Call ExecuteCritical
    Case Else
        mblnUnRefresh = True
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            With mPatiInfo
                If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    If cboDept.ListIndex = -1 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                    Else
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "开嘱科室=" & Split(cboDept.List(cboDept.ListIndex), "-")(1) & "|=" & CLng(cboDept.ItemData(cboDept.ListIndex)))
                    End If
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "病人ID=" & .病人ID, "主页ID=" & .主页ID, "住院号=" & .住院号, "病人科室=" & .科室ID)
                End If
            End With
        ElseIf Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 1, conMenu_File_MedRecPreview * 100# + 4) Then
            Call PrintInMedRec(mclsInOutMedRec, IIf(Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6), 2, 1), mPatiInfo.病人ID, mPatiInfo.主页ID, mobjReport, mPatiInfo.科室ID, Me, Val(Mid(Control.ID & "", Len(Control.ID & ""))))
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "路径"
               If rptPati.SelectedRows.Count >= 1 Then '当选中行才能执行rptPati.SelectedRows(0).GroupRow的判断,否则会报错
                    If rptPati.SelectedRows(0).GroupRow = False Then
                        If rptPati.SelectedRows(0).Record(COL_婴儿科室ID).Value <> 0 Then
                            If rptPati.SelectedRows(0).Record(COL_婴儿科室ID).Value = cboDept.ItemData(cboDept.ListIndex) Or rptPati.SelectedRows(0).Record(COL_婴儿病区ID).Value = cboDept.ItemData(cboDept.ListIndex) Then
                                MsgBox "该病人已经转出本科室了，只有婴儿留在本科室，不允许操作路径。", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "医嘱"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "病历"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "护理"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "新版护理"
                Call mclsTendsNew.zlExecuteCommandBars(Control)
            Case "护理病历"
                Call mclsTendEPRs.zlExecuteCommandBars(Control)
            Case "新病历"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case "疾病报告"
                Call mclsDisease.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p住院医生站, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng病人ID, mlng主页ID, "")
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        mblnUnRefresh = False
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "住院号(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "床  号(&2)"
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
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "路径"
            Call mclsPath.zlPopupCommandBars(CommandBar)
       Case "医嘱"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "病历"
    
       Case "护理"
       
       End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim i As Long
        
    If Not mblnIsInit Then
        mblnIsInit = True
        If Not mobjSquareCard Is Nothing Then
            Call PatiIdentify.zlInit(Me, glngSys, p住院医生站, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
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
    
    Select Case Control.ID
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
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '按科室/病区显示
        Control.Checked = mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
        Control.Enabled = mblnDeptViewEnabled
    Case conMenu_Tool_KssAudit  '抗菌用药审核
        If GetInsidePrivs(p抗菌用药审核) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_OPSAudit  '手术审核管理
        If GetInsidePrivs(p手术审核管理) = "" Or Not gbln手术分级管理 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_OPSEmpower  '手术授权管理
        If GetInsidePrivs(p手术授权管理) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_TransAudit '输血分级管理
        If GetInsidePrivs(p输血审核管理) = "" Or Not gbln输血分级管理 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_CISMed  '临床自管药
        If InStr(GetInsidePrivs(p住院医生站), ";临床自管药;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_Archive '电子病案查阅
        If GetInsidePrivs(p电子病案查阅) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    Case conMenu_Tool_ExaReport
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Tool_HealthCard  '居民健康卡
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        If GetInsidePrivs(p疾病诊断参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 '药品及诊疗参考
        If GetInsidePrivs(p药品诊疗参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Meet '会诊病人
        If InStr(mstrPrivs, "会诊病人") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
                    blnEnabled = Val(Mid(rptPati.SelectedRows(0).Record(col_类型).Value, 1, 1)) = pt会诊
                End If
            End If
            Control.Enabled = blnEnabled
            If Me.Visible Then Control.Visible = tbcPati.Selected.Tag = "会诊"
        End If
    Case conMenu_Tool_MeetOpen '接受会诊
        blnEnabled = False
        Control.Caption = IIf(mbln接受会诊, "取消接受会诊(&X)", "接受会诊(&O)")
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
                blnEnabled = True
            End If
        End If
        Control.Enabled = blnEnabled
        If Me.Visible Then Control.Visible = rptPati.SelectedRows(0).Record(col_执行状态).Value = 0
    Case conMenu_Tool_MeetFinish, conMenu_Tool_MeetCancel '完成会诊,取消完成
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
                blnEnabled = rptPati.SelectedRows(0).Record(col_执行状态).Value = IIf(Control.ID = conMenu_Tool_MeetFinish, 0, 1)
            End If
        End If
        If Control.ID = conMenu_Tool_MeetFinish Then
            blnEnabled = blnEnabled And mbln接受会诊
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Tool_MeetIdea
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow Then
                If Val(rptPati.SelectedRows(0).Record(COL_申请序号).Value) > 0 Then
                    blnEnabled = True
                    If Val(rptPati.SelectedRows(0).Record(col_执行状态).Value) = 1 Then
                        Control.Caption = "查看会诊意见(&V)"
                    Else
                        Control.Caption = "填写会诊意见(&W)"
                    End If
                End If
            End If
        End If
        If Control.Caption = "填写会诊意见(&W)" Then
            blnEnabled = blnEnabled And mbln接受会诊
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Tool_MedRecAuditSubmit '提交审查
        '出院病人，尚未提交或拒绝审查状态
        If InStr(mstrPrivs, "病案审查提交") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = False
            If Control.Caption = "重新提交(&S)" Then Control.Caption = "提交审查(&S)"
            If rptPati.SelectedRows.Count > 0 Then
                With rptPati.SelectedRows(0)
                    If Not .GroupRow Then
                        If (Int(Val(.Record(col_类型).Value)) = pt出院 Or Int(Val(.Record(col_类型).Value)) = pt死亡) Or (tbcPati.Selected.Tag = "出院" And Val(.Record(col_病人Id).Value) <> 0) Then
                            '可能是在院抽查反馈状态，出院后并未提交审查。被拒绝的允许重新提交审查
                            If .Record(col_图标).Value <> 1 Or .Record(col_审查).Value = 2 Then blnEnabled = True
                            If .Record(col_审查).Value = 2 And Control.Caption = "提交审查(&S)" Then Control.Caption = "重新提交(&S)"
                        End If
                    End If
                End With
            End If
            Control.Enabled = blnEnabled
            If Me.Visible Then Control.Visible = tbcPati.Selected.Tag = "出院"
        End If
    Case conMenu_Tool_MedRecAuditCancel '取消提交
        '出院病人，已经提交状态
        If InStr(mstrPrivs, "病案审查提交") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                With rptPati.SelectedRows(0)
                    If Not .GroupRow Then
                        If (Int(Val(.Record(col_类型).Value)) = pt出院 Or Int(Val(.Record(col_类型).Value)) = pt死亡) Or (tbcPati.Selected.Tag = "出院" And Val(.Record(col_病人Id).Value) <> 0) Then
                            If .Record(col_审查).Value = 1 Then blnEnabled = True
                        End If
                    End If
                End With
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Tool_MedRecAuditResponse '审查反馈
        '都可以调用，至少可以查看(当前或历史)
        Control.Enabled = rptPati.Rows.Count > 0
    Case conMenu_Tool_MedRecAuditWriteResponse '书写审查意见
        Control.Enabled = rptPati.Rows.Count > 0 And InStr(GetInsidePrivs(p电子病案审查), ";审查病案;") <> 0
    Case conMenu_Tool_MedRatio
        If InStr(mstrPrivs, "药占比查询") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_MedRec '首页整理
        If InStr(mstrPrivs, "首页整理") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = cboPages.ListIndex <> -1 And mPatiInfo.主页ID > 0
        End If
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
    Case conMenu_Edit_TraReactionRecord '输血反应
        Control.Visible = InStr(1, GetInsidePrivs(9005, , 2200), "输血反应登记") <> 0
        Control.Enabled = Control.Visible And gbln血库系统
    Case conMenu_View_Notify '审阅提醒
        Control.Enabled = rptNotify.Visible
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
        '抗菌用药报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                If gblnKSSStrict Then
                    Control.Visible = True
                Else
                    Control.Visible = False
                End If
                Exit Sub
            End If
        End If
        '首页报表
        If Between(Control.ID, conMenu_File_MedRecPrint * 100# + 3, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 3, conMenu_File_MedRecPreview * 100# + 4) Then
            If mintMecStandard = 0 Or mintMecStandard = 3 Or mintMecStandard = 1 Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            Exit Sub
        End If
        mclsReg.zlUpdateCommandBars Control
        Select Case tbcSub.Selected.Tag
        Case "路径"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "医嘱"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "病历"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "护理"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "新版护理"
            Call mclsTendsNew.zlUpdateCommandBars(Control)
        Case "护理病历"
            Call mclsTendEPRs.zlUpdateCommandBars(Control)
        Case "新病历"
            Call mclsEMR.zlUpdateCommandBars(Control)
        Case "疾病报告"
            Call mclsDisease.zlUpdateCommandBars(Control)
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
        
    Me.Caption = "住院医生工作站 - " & objItem.Caption & "(当前用户：" & UserInfo.姓名 & ")"
    If Not mbln单个病人 Then
    If InStr(mstrNotify, "1") > 0 Or Not Me.Visible Then
        dkpMain.Panes(2).Closed = False
        dkpMain.Panes(2).Hidden = Val(dkpMain.Panes(2).Tag) = 1
        dkpMain.Panes(2).Title = "消息提醒"
    Else
        dkpMain.Panes(2).Tag = IIf(dkpMain.Panes(2).Hidden, 1, 0)
        dkpMain.Panes(2).Close
    End If
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
    
    If Not mclsReg Is Nothing And (InStr(GetInsidePrivs(9000), ";预约;") > 0 Or InStr(GetInsidePrivs(9000), ";预约登记;") > 0) Then
        Call mclsReg.zlDefCommandBars(Me, Me.cbsMain, True)
    End If
    
    '子窗口重新加入
    Select Case objItem.Tag
    Case "路径"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain, 0)
    Case "医嘱"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0)
    Case "病历"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "护理"
        Call mclsTends.zlDefCommandBars(Me.cbsMain)
    Case "新版护理"
        Call mclsTendsNew.zlDefCommandBars(Me.cbsMain)
    Case "护理病历"
        Call mclsTendEPRs.zlDefCommandBars(Me.cbsMain)
    Case "新病历"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "疾病报告"
        Call mclsDisease.zlDefCommandBars(Me.cbsMain)
    Case "护理评估评分"
        
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p住院医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag)
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
        cbsMain(lngCount).EnableDocking xtpFlagStretched
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
    Dim strInPatiNO As String, i As Integer
    Dim blnEdit As Boolean, lng路径状态 As Long, bln补费 As Boolean
    Dim lngType As Long, lng病区ID As Long, lng科室ID As Long
    Dim lng界面科室ID As Long, lng执行科室ID As Long
    Dim lngState As TYPE_PATI_State, blnDis As Boolean

    If mlng病人ID = 0 Or cboDept.ListIndex = -1 Then
        
        For i = 0 To tbcSub.ItemCount - 1 '默认情况，传染病报告卡不显示
            If tbcSub.Item(i).Tag = "疾病报告" Then
                blnDis = tbcSub.Item(i).Selected
                tbcSub.Item(i).Visible = False
                If blnDis Then '如果此前选中的是传染病报告卡则先隐藏再选中第0个TAB
                    tbcSub.Item(0).Selected = True: Exit Sub
                End If
                Exit For
            End If
        Next
        
        For i = 0 To tbcSub.ItemCount - 1 '默认情况，护理评估评分不显示
            If tbcSub.Item(i).Tag = "护理评估评分" Then
                tbcSub.Item(i).Visible = False
                Exit For
            End If
        Next
        
        
        '要求子窗体按无数据处理界面
        Select Case objItem.Tag
        Case "住院一览"
            Call mfrmInView.zlRefresh(0, 0, 0, 0)
        Case "路径"
            Call mclsPath.zlRefresh(0, 0, 0, 0, 0, False)
        Case "医嘱"
            Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "病历"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "护理"
            Call mclsTends.zlRefresh(0, 0, 0, False, True)
        Case "新版护理"
            Call mclsTendsNew.zlRefresh(0, 0, 0, False, True, 0, 0, 1)
        Case "护理病历"
            Call mclsTendEPRs.zlRefresh(0, 0, 0, False, False, False, True)
        Case "监护"
            Call mclsWardMonitor.HideWindow
        Case "新病历"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 2)
        Case "疾病报告"
            Call mclsDisease.zlRefresh(0, 0, 2, 0, False, False)
        Case "护理评估评分"
        
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p住院医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            lngType = Val(Mid(rptPati.SelectedRows(0).Record(col_类型).Value, 1, 1))
            lng界面科室ID = cboDept.ItemData(cboDept.ListIndex)
            lng路径状态 = Val(rptPati.SelectedRows(0).Record(col_路径状态).Value)
            
            If lngType = pt最近转出 Then
                '获取病人原来的病区和科室
                Call GetPatiLastChange(mlng病人ID, mlng主页ID, lng病区ID, lng科室ID)
                '如果是以科室查看则取当前科室;
                If mintDeptView = 0 Then
                    If cboDept.ListIndex <> -1 Then lng科室ID = cboDept.ItemData(cboDept.ListIndex)
                End If
                lngState = ps最近转出
            Else
                lng病区ID = .病区ID
                lng科室ID = .科室ID
                
                If lngType = pt会诊 Then
                    lng执行科室ID = rptPati.SelectedRows(0).Record(col_执行科室ID).Value
                    lngState = IIf(rptPati.SelectedRows(0).Record(col_执行状态).Value = 0, ps待诊, ps已诊)
                Else
                    lngState = IIf(.出院日期 = CDate(0), IIf(.状态 = 3, ps预出, ps在院), ps出院)
                End If
            End If
            
            
            For i = 0 To tbcSub.ItemCount - 1 '默认情况，传染病报告卡不显示
                If tbcSub.Item(i).Tag = "疾病报告" Then
                    blnDis = tbcSub.Item(i).Selected
                    tbcSub.Item(i).Visible = True
                    If tbcSub.Item(i).Visible = False And blnDis Then
                        tbcSub.Item(0).Selected = True: Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next

            For i = 0 To tbcSub.ItemCount - 1 '默认情况，护理评估评分不显示
                If tbcSub.Item(i).Tag = "护理评估评分" Then
                    tbcSub.Item(i).Visible = True
                    Exit For
                End If
            Next
            Select Case objItem.Tag
            Case "住院一览"
                Call mfrmInView.zlRefresh(.病人ID, .主页ID, lng科室ID, .婴儿)
            Case "路径"
                Call mclsPath.zlRefresh(.病人ID, .主页ID, lng病区ID, lng科室ID, .状态, .数据转出, True, lngState, , mclsMipModule)
            Case "医嘱"
                If .状态 = 1 Then '入院待入住
                   If Val(zlDatabase.GetPara("允许给待入住病人下达医嘱", glngSys, p住院医嘱下达, 1)) = 0 Then
                        lngState = ps待转入 'lngState=ps待转入时新开医嘱等功能不可用
                   End If
                End If
                Call mclsAdvices.zlRefresh(.病人ID, .主页ID, lng病区ID, lng科室ID, lngState, .数据转出, , , lng执行科室ID, lng路径状态, _
                    cboDept.ItemData(cboDept.ListIndex), mclsMipModule, .婴儿, 0, mlng会诊医嘱ID)
            Case "病历"
                blnEdit = True
                With rptPati.SelectedRows(0)
                    If Int(lngType) = pt出院 Or Int(lngType) = pt死亡 Then
                        If Not (.Record(col_审查).Value = 0 Or .Record(col_审查).Value = 2 Or .Record(col_审查).Value = 999) Then
                            '可能是在院抽查反馈状态，出院后并未提交审查
                            If .Record(col_图标).Value = 1 Then blnEdit = False
                        End If
                    End If
                End With
                blnEdit = blnEdit And (lng界面科室ID = IIf(mintDeptView = 0, lng科室ID, lng病区ID) Or lngType = pt会诊 Or lngType = pt最近转出)
                '待入住的病人不允许编辑病历
                blnEdit = blnEdit And tbcPati.Selected.Tag <> "待入住" And .病人ID = mlng病人ID
                '医嘱和临床路径都可能删除对应的病历文件，所以强制刷新
                Call mclsEPRs.zlRefresh(.病人ID, .主页ID, IIf(mintDeptView = 0, lng界面科室ID, lng科室ID), blnEdit, .数据转出, 0, True, lng病区ID, lngState)
            Case "护理"
                Call mclsTends.zlRefresh(.病人ID, .主页ID, lng科室ID, False, True, lng病区ID, lngState)
            Case "新版护理"
                Call mclsTendsNew.zlRefresh(.病人ID, .主页ID, lng科室ID, False, True, lng病区ID, lngState, 1)
            Case "护理病历"
                Call mclsTendEPRs.zlRefresh(.病人ID, .主页ID, lng科室ID, False, False, .数据转出, True)
            Case "监护"
                strInPatiNO = Trim(rptPati.SelectedRows(0).Record(col_住院号).Value)
                If strInPatiNO = "" Then
                    Call mclsWardMonitor.HideWindow
                Else
                    Call mclsWardMonitor.ShowInfor(strInPatiNO)
                End If
            Case "新病历"
                Call mclsEMR.zlRefresh(.病人ID, .主页ID, IIf(mintDeptView = 0 Or lngType = pt会诊, cboDept.ItemData(cboDept.ListIndex), lng科室ID), lngState, 2)
            Case "疾病报告"
                If objItem.Visible Then
                    '待入住的病人不允许编辑病历
                    Call mclsDisease.zlRefresh(.病人ID, .主页ID, 2, IIf(mintDeptView = 0, cboDept.ItemData(cboDept.ListIndex), lng科室ID), .数据转出, .病人ID = mlng病人ID And tbcPati.Selected.Tag <> "待入住", lngState)
                End If
            Case "护理评估评分"
                If Not gobjNurseIntegrate Is Nothing Then
                    If objItem.Visible Then
                        Call mcolSubForm("_护理评估评分").zlRefresh(.病人ID, .主页ID)
                    End If
                End If
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p住院医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag, .病人ID, "", .主页ID, .数据转出, _
                        lng执行科室ID, cboDept.ItemData(cboDept.ListIndex), lng病区ID, lng科室ID, , lngState, , lng路径状态)
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
    Dim strFunName As String

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
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "预览首页(&V)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "打印首页(&P)", -1, False
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Dept, "部门显示(&D)") '固有
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 1, "按科室显示(&D)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 2, "按病区显示(&U)", -1, False)
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
        Set objControl = .Add(xtpControlButton, conMenu_Tool_KssAudit, "抗菌用药审核(&K)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSEmpower, "手术授权管理(&N)")
        objControl.IconId = 3553
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSAudit, "手术审核管理(&L)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "输血审核管理(&M)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_CISMed, "临床自管药(&J)")
        objControl.IconId = 3901
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ExaReport, "查阅体检总检报告")
            objControl.IconId = conMenu_File_Preview
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRatio, "药占比查询")
            objControl.IconId = 813: objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "首页整理(&M)"): objControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_MedRecAudit, "病案审查(&Q)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditWriteResponse, "书写审查意见", -1, False)
                objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditSubmit, "提交审查(&S)", -1, False)
                objControl.IconId = conMenu_Manage_Complete
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditCancel, "取消提交(&C)", -1, False)
                objControl.IconId = conMenu_Edit_Untread
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "审查反馈(&R)", -1, False)
                objControl.IconId = 3814
                objControl.BeginGroup = True
                objControl.ToolTipText = "处理或查看病案审查反馈"
        End With
        
        If gbln血库系统 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReactionRecord, "输血反应记录"): objControl.BeginGroup = True
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Positive, "阳性结果")
            objControl.IconId = 3551
        If mbln危急值 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Critical, "危急值")
                objControl.IconId = 4113
        End If
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Meet, "病人会诊(&E)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_MeetOpen, "接受会诊(&O)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetIdea, "填写会诊意见(&W)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetFinish, "完成会诊(&F)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetCancel, "取消完成(&C)", -1, False
        End With
        If Not mobjSquareCard Is Nothing Then
            On Error Resume Next
            If mobjSquareCard.zlHealthArchiveIsSHow(Me, p住院医生站, strFunName, "") Then
                If err.Number = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                    objControl.BeginGroup = True
                    objControl.IconId = 3208
                End If
            End If
            On Error GoTo 0
        End If
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


    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
            
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Tool_MedRec, "首页", objControl.Index + 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "预览首页")
                objControl.IconId = conMenu_File_Preview
            Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "打印首页")
                objControl.IconId = conMenu_File_Print
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "病案查阅")
            objControl.ToolTipText = "电子病案查阅"

        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditSubmit, "提交")
            objControl.IconId = conMenu_Manage_Complete
            objControl.ToolTipText = "提交病案审查"
        
        Set objPopup = .Add(xtpControlPopup, conMenu_Tool_Meet, "会诊")
        objPopup.ID = conMenu_Tool_Meet
        objPopup.IconId = conMenu_Tool_Meet
        objPopup.Style = xtpButtonIconAndCaption
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_MeetOpen, "接受会诊", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetIdea, "填写会诊意见(&W)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetFinish, "完成会诊(&F)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetCancel, "取消完成(&C)", -1, False
        End With

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助") '固有
        objControl.BeginGroup = True
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
    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '打印设置
'        .AddHiddenCommand conMenu_File_Excel '输出到Excel
'        .AddHiddenCommand conMenu_View_Jump '跳转
    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8", _
                "ZL1_INSIDE_1261_9", "ZL1_INSIDE_1261_10")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnInView And UnloadMode = vbFormControlMenu Then
        Cancel = 1 '住院一览放大状态,取消窗体关闭
        mfrmInView.zlExecuteCommandBars
    End If
End Sub

Private Sub mFrmConsultation_Unload(Cancel As Integer)
    '界面数据刷新
    If Cancel = 0 And (Not mFrmConsultation Is Nothing) Then
        If mFrmConsultation.Tag = "1" Then
            mFrmConsultation.Tag = ""
            Call LoadPatients
            If rptPati.Visible Then rptPati.SetFocus
        End If
    End If
End Sub

Private Sub mfrmInView_ResizeForm(ByVal bytFunc As Long)
    Dim varItem As Variant
    On Error Resume Next
    With tbcSub
        If bytFunc = 1 Then
         '放大
            SetParent mfrmInView.hwnd, Me.hwnd
            With mfrmInView
                .Tag = .Left & "," & .Top & "," & .Width & "," & .Height
                .Left = 0: .Width = Me.ScaleWidth
                .Top = 0: .Height = Me.ScaleHeight - Me.stbThis.Height
            End With
            mblnInView = True
        Else
        '缩小
            SetParent mfrmInView.hwnd, .hwnd
            With mfrmInView
                 varItem = Split(.Tag, ",")
                 .Left = varItem(0): .Top = varItem(1)
                 .Width = varItem(2): .Height = varItem(3)
            End With
            mblnInView = False
        End If
    End With
End Sub

Private Sub mfrmInView_ViewPACSImage(ByVal 医嘱ID As Long)
'功能：PACS观片处理
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(医嘱ID, Me, mPatiInfo.数据转出)
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人ID
    End If
    
    If chkFilter.Value = 1 And chkFilter.Visible = True Then
        Call LoadPatients
    Else
        Call ExecuteFindPati(False, lngPatiID)
    End If
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index - 1: mstrFindType = objCard.名称
    If tbcPati.ItemCount <> 0 Then
        chkFilter.Visible = IIf(mstrFindType = "住院号" And tbcPati.Selected.Tag = "出院", True, False)
        Call picPati_Resize
    End If
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub lbl审查_Click()
    If cboDept.ListIndex = -1 Then Exit Sub
    
    '非模态显示审查反馈窗体
    If mfrmResponse Is Nothing Then
        Set mfrmResponse = New frmAuditResponse
    End If
    mblnUnRefresh = True
    Call mfrmResponse.ShowMe(Me, cboDept.ItemData(cboDept.ListIndex), mintDeptView, mblnICU, 0, mstrPrivs)
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_EditDiagnose(ParentForm As Object, ByVal 病人ID As Long, ByVal 主页ID As Long, ByVal 科室ID As Long, ByVal str类型 As String, Succeed As Boolean)
'功能：录入诊断
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p住院医生站, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    mblnUnRefresh = True
    
    If Not mclsInOutMedRec.ShowInMedRecEdit(ParentForm, 病人ID, 主页ID, 科室ID, rptPati.SelectedRows(0).Record(col_路径状态).Value, str类型, mstrPrivs, , False) Then
        Succeed = False
    Else
        Succeed = mclsInOutMedRec.IsDiagInput
    End If
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'功能：医嘱子窗体要求刷新
    If Not RefreshNotify Then '注意要判断
        Call LoadPatients
    ElseIf rptNotify.Visible Then
        '仅刷新医嘱提醒区域
        Call LoadNotify
    End If
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Dim strTmp As String
    Dim intTmp As Long
    If Text = "" And rptPati.SelectedRows.Count > 0 Then
        With rptPati.SelectedRows(0)
            If Not .GroupRow Then
                If Val(.Record(col_病人Id).Value) <> 0 Then intTmp = 1
            End If
            If intTmp = 1 Then
                stbThis.Panels(2).Text = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag)
                lblFee(1).Caption = GetPati费用信息(mlng病人ID, mlng主页ID) & IIf(InStr(mstrPrivs, "药占比查询") = 0, "", Get住院费用药占比(mlng病人ID, mlng主页ID))
                '出院病人不显示输液量
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

Private Sub mclsAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclsEPRs_Activate()
    If timNotify.Enabled And rptNotify.Visible Then
        '仅刷新医嘱提醒区域(自动刷新时)
        Call LoadNotify
    End If
End Sub

Private Sub mclsInOutMedRec_Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)
'功能: 完成首页整理后刷新数据
'   EditCancel=取消退出首页
' strTag=附加信息，现在存储门诊病人照片文件的路径，以后扩展时，以|分割
    Dim strSQL As String
    Dim lng路径ID As Long, lng路径状态 As Long, lng疾病ID As Long, lng诊断ID As Long
    Dim rsPath As ADODB.Recordset, rsTmp As ADODB.Recordset, rsNext As ADODB.Recordset
    Dim bln中医 As Boolean
    Dim i As Long, blnNo As Boolean
    Dim str疾病IDs As String
    Dim str诊断IDs As String
    Dim objControl As CommandBarControl
    Dim blnNotView As Boolean
    
    On Error GoTo errH
    
    If Not blnEditCancel Then
        If mPatiInfo.主页ID <> mlng主页ID Then Exit Sub
        If InStr(";" & GetPrivFunc(glngSys, p疾病报告填写) & ";", ";病历书写;") > 0 Then
            '查看一年以内是否填写过传染病报告卡
            '检查诊断是否需要填写传染病报告卡，不需要就退出
            '如果需要填写，检查一年内是否有重复填写的，没有就填写
            '有的话就弹出一年内重复填写过的
            '不填写的话直接退出，填写的话就填写
            mclsInOutMedRec.Hide
            Set rsTmp = mclsDisease.SatisfyEditDiseaseDoc(mlng病人ID, mlng主页ID, mPatiInfo.科室ID, str疾病ID, str诊断ID)
            If Not rsTmp Is Nothing Then
                If rsTmp.RecordCount > 0 Then
                    If Not mclsDis.ShowDiseaseStation(Me, mlng病人ID, mlng主页ID, 2, mPatiInfo.科室ID, str疾病ID, str诊断ID, blnNotView) Then
                        Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng病人ID, mlng主页ID, 2, mPatiInfo.科室ID, blnNo)
                        If blnNo Then
                            Call mclsDis.EditNotFillReason(Me, mlng病人ID, mlng主页ID, 2)
                        End If
                    ElseIf blnNotView Then
                        Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng病人ID, mlng主页ID, 2, mPatiInfo.科室ID, blnNo)
                        If blnNo Then
                            Call mclsDis.EditNotFillReason(Me, mlng病人ID, mlng主页ID, 2)
                        End If
                    End If
                End If
            End If
        End If
        
        Call LoadPatients
        '导入路径
        lng路径状态 = rptPati.SelectedRows(0).Record(col_路径状态).Value
        '1.路径状态-未导入
        '2.科室-启用临床路径的科室
        '3.病人诊断信息发生变化
        '4.病人状态-正常在院
        '5.当前用户拥有导入路径的权限
        '满足以上5个条件时才弹出路径导入窗体
        If lng路径状态 = -1 And mclsInOutMedRec.IsDiagChange _
            And mPatiInfo.状态 = 0 And InStr(GetInsidePrivs(p临床路径应用), ";导入路径;") <> 0 Then
            
            If HavePath(mPatiInfo.科室ID) Then
                Set rsTmp = Get病种ID(mlng病人ID, mlng主页ID, mPatiInfo.科室ID, bln中医)
                If bln中医 Then
                    rsTmp.Filter = "诊断类型 = 12 or 诊断类型 = 2 " '取入院诊断
                    If rsTmp.RecordCount = 0 Then Exit Sub
                    For i = 1 To rsTmp.RecordCount
                        lng疾病ID = Val("" & rsTmp!疾病id)
                        lng诊断ID = Val("" & rsTmp!诊断id)
                        Set rsPath = GetPathTable(lng疾病ID, lng诊断ID, mPatiInfo.科室ID)
                        If rsPath.RecordCount > 0 Then Exit For
                        rsTmp.MoveNext
                    Next
                Else
                    If rsTmp.RecordCount > 0 Then
                        lng疾病ID = Val("" & rsTmp!疾病id)
                        lng诊断ID = Val("" & rsTmp!诊断id)
                    End If
                    Set rsPath = GetPathTable(lng疾病ID, lng诊断ID, mPatiInfo.科室ID)
                End If
                
                If rsPath.RecordCount = 0 Then
                    Set rsNext = Get病种ID(mlng病人ID, mlng主页ID, mPatiInfo.科室ID, , 1)
                    If rsNext.RecordCount = 0 Then Exit Sub
                End If
                
                Call mclsInOutMedRec.Hide '隐藏首页
                Call mclsPath.zlRefresh(mlng病人ID, mlng主页ID, mPatiInfo.病区ID, mPatiInfo.科室ID, mPatiInfo.状态, False, True)
                Call mclsPath.zlImportPath
                Call LoadPatients '界面数据刷新
                If rptPati.Visible Then rptPati.SetFocus
            End If
        '有合并路径的话就导入合并路径。mPatiInfo.状态 =0:正常住院；lng路径状态 = 1:路径正在执行中
        '找出已经导入的路径。
        ElseIf lng路径状态 = 1 And mclsInOutMedRec.IsDiagChange And mPatiInfo.状态 = 0 And InStr(GetInsidePrivs(p临床路径应用), ";导入路径;") <> 0 Then
            If HavePath(mPatiInfo.科室ID) Then
                '已经导入的路径
                strSQL = "Select a.ID,a.路径ID,A.合并路径个数,c.路径ID as 原路径ID,a.版本号,a.状态,a.当前阶段ID,a.当前天数,b.名称 as 未导入原因,c.父ID,c.分支ID,d.分支ID as 前一阶段分支ID,e.结束路径控制,a.合并路径个数,e.名称 as 路径名称,a.导入人,a.导入时间,a.结束时间" & _
                        " From 病人临床路径 A,变异常见原因 B,临床路径阶段 C,临床路径阶段 D,临床路径目录 E" & _
                        " Where a.病人ID = [1] And a.主页ID = [2] And a.路径ID=e.id And a.未导入原因 = b.编码(+) And a.当前阶段ID = c.ID(+) And a.前一阶段ID=d.id(+)" & _
                        " Order By a.导入时间 Desc"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mPatiInfo.科室ID)
                '还没有导入路径的话就直接退出
                If rsTmp.RecordCount = 0 Then
                    Exit Sub
                Else
                    lng路径ID = NVL(rsTmp!ID)
                    '合并路径已经5个了的话直接退出
                    If Val(rsTmp!合并路径个数 & "") >= 5 Then Exit Sub
                End If

                strSQL = "Select 1 From 病人路径评估 A, 病人路径执行 B" & vbNewLine & _
                         "Where a.路径记录id = b.路径记录id And a.路径记录id = [1] And A.阶段id = [2] And b.天数 = [3] And a.日期 = b.日期"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng路径ID, rsTmp!当前阶段ID & "", rsTmp!当前天数 & "")
                '还没有评估的话直接退出
                If rsTmp.RecordCount = 0 Then
                    Exit Sub
                End If

                Set rsTmp = Get病种ID(mlng病人ID, mlng主页ID, mPatiInfo.科室ID, , 2)
                If rsTmp.RecordCount = 0 Then Exit Sub

                '如果是合并路径，则传入所有非导入病种的其他诊断或并发症
                Do While Not rsTmp.EOF
                    If Val(rsTmp!疾病id & "") <> 0 Then
                        str疾病IDs = str疾病IDs & "," & rsTmp!疾病id
                    End If
                    If Val(rsTmp!诊断id & "") <> 0 Then
                        str诊断IDs = str诊断IDs & "," & rsTmp!诊断id
                    End If
                    rsTmp.MoveNext
                Loop
                rsTmp.MoveFirst
                str疾病IDs = Mid(str疾病IDs, 2)
                str诊断IDs = Mid(str诊断IDs, 2)

                '这里加Distinct是因为，诊断id和疾病id做了绑定对应，所以，查出来会有重复值，排开已经导入了的合并路径
                strSQL = "Select Distinct a.Id, a.分类, a.编码, a.名称, a.说明, Nvl(a.适用病情,'通用') 适用病情, a.适用性别, a.适用年龄, a.最新版本, c.标准住院日,Nvl(a.病例分型,'无') as 病例分型,Nvl(a.确诊天数,0) as 确诊天数,b.疾病ID,b.诊断ID" & vbNewLine & _
                        "From 临床路径目录 A, 临床路径病种 B,临床路径版本 C" & vbNewLine & _
                        "Where a.Id = b.路径id And (instr(',' || [2] || ',',',' || b.疾病ID || ',')>0 and [2] is not null Or instr(',' || [4] || ',',',' || b.诊断ID || ',')>0 and [4] is not null)  And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
                        "And a.Id = c.路径id And a.性质=1 And b.性质=0 And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [1]))" & _
                        " And Not Exists(Select 1 From 病人合并路径 D Where a.id=d.路径ID  and d.首要路径记录ID=[3])"

                Set rsPath = zlDatabase.OpenSQLRecord(strSQL, "读取路径目录", mPatiInfo.科室ID, str疾病IDs, lng路径ID, str诊断IDs)
                If rsPath.RecordCount = 0 Then Exit Sub

                Call mclsInOutMedRec.Hide     '隐藏首页

                Set objControl = cbsMain.FindControl(, conMenu_Edit_ImportMerge, True, True)
                If Not objControl Is Nothing Then
                     Call mclsPath.zlExecuteCommandBars(objControl)
                End If

                Call mclsPath.zlRefresh(mlng病人ID, mlng主页ID, mPatiInfo.病区ID, mPatiInfo.科室ID, mPatiInfo.状态, False, True)
                Call LoadPatients '界面数据刷新
                If rptPati.Visible Then rptPati.SetFocus
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

Private Sub mclsDis_Closed(ByVal lngFunID As Long, ByVal strTag As String)
    Dim lng路径状态 As Long, lng疾病ID As Long, lng诊断ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim rsPath As ADODB.Recordset, rsNext As ADODB.Recordset
    Dim bln中医 As Boolean
    Dim i As Long
    
    lng路径状态 = rptPati.SelectedRows(0).Record(col_路径状态).Value
    '1.路径状态-未导入
    '2.科室-启用临床路径的科室
    '3.病人诊断信息发生变化
    '4.病人状态-正常在院
    '5.当前用户拥有导入路径的权限
    '满足以上5个条件时才弹出路径导入窗体
    If lng路径状态 = -1 And mPatiInfo.状态 = 0 And InStr(GetInsidePrivs(p临床路径应用), ";导入路径;") <> 0 Then
        If HavePath(mPatiInfo.科室ID) Then
            Set rsTmp = Get病种ID(mlng病人ID, mlng主页ID, mPatiInfo.科室ID, bln中医)
            If bln中医 Then
                rsTmp.Filter = "诊断类型 = 2 OR 诊断类型 = 12 "   '取入院诊断
                If rsTmp.RecordCount = 0 Then Exit Sub
                For i = 1 To rsTmp.RecordCount
                    lng疾病ID = Val("" & rsTmp!疾病id)
                    lng诊断ID = Val("" & rsTmp!诊断id)
                    Set rsPath = GetPathTable(lng疾病ID, lng诊断ID, mPatiInfo.科室ID)
                    If rsPath.RecordCount > 0 Then Exit For
                    rsTmp.MoveNext
                Next
            Else
                If rsTmp.RecordCount > 0 Then
                    lng疾病ID = Val("" & rsTmp!疾病id)
                    lng诊断ID = Val("" & rsTmp!诊断id)
                End If
                Set rsPath = GetPathTable(lng疾病ID, lng诊断ID, mPatiInfo.科室ID)
            End If

            If rsPath.RecordCount = 0 Then
                Set rsNext = Get病种ID(mlng病人ID, mlng主页ID, mPatiInfo.科室ID, , 1)
                If rsNext.RecordCount = 0 Then Exit Sub
            End If
            Call mclsDis.HideFrm(0) '隐藏传染病阳性结果窗体
            Call mclsPath.zlRefresh(mlng病人ID, mlng主页ID, mPatiInfo.病区ID, mPatiInfo.科室ID, mPatiInfo.状态, False, True)
            Call mclsPath.zlImportPath
            Call LoadPatients '界面数据刷新
            If rptPati.Visible Then rptPati.SetFocus
        End If
    End If
End Sub

Private Sub mclsMipModule_OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
'功能：点击冒泡消息后定位病人
    If InStr(",ZLHIS_PATIENT_002,ZLHIS_PATIENT_012,ZLHIS_PATIENT_009,ZLHIS_PATIENT_006,ZLHIS_PATIENT_010,", "," & strMsgKey & ",") > 0 Then
        If tbcPati.Item(1).Selected = False Then
            tbcPati.Item(1).Selected = True '选项卡切换时会刷新病人列表
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
    Dim blnRecToLis As Boolean '是否加载到提醒列表中
    Dim rsMsg As ADODB.Recordset
    
    If cboDept.ListIndex = -1 Then Exit Sub
    
    If Mid(mstrNotify, 1, 1) = "1" And strMsgItemIdentity = "ZLHIS_EMR_021" Then
        blnRecToLis = True
    ElseIf Mid(mstrNotify, 2, 1) = "1" And InStr(",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,", "," & strMsgItemIdentity & ",") > 0 Then
        blnRecToLis = True
    ElseIf Mid(mstrNotify, 3, 1) = "1" And InStr(",ZLHIS_LIS_003,ZLHIS_PACS_005,", "," & strMsgItemIdentity & ",") > 0 Then
        blnRecToLis = True
    ElseIf Mid(mstrNotify, 4, 1) = "1" And InStr(",ZLHIS_LIS_002,ZLHIS_PACS_003,", "," & strMsgItemIdentity & ",") > 0 Then
        blnRecToLis = True
    ElseIf Mid(mstrNotify, 5, 1) = "1" And InStr(",ZLHIS_CIS_026,ZLHIS_CIS_027,ZLHIS_CIS_028,ZLHIS_CIS_029,ZLHIS_CIS_030,", "," & strMsgItemIdentity & ",") > 0 Then
        blnRecToLis = True
    End If
    
    If blnRecToLis Then
        Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
        If rsMsg Is Nothing Then Exit Sub
        Call AddMsgToLis(rsMsg)
    Else
        Call RecMsgToBub(mclsMipModule, cboDept.ItemData(cboDept.ListIndex), 2, strMsgItemIdentity, strMsgContent, mintDeptView)
    End If
    
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

Private Sub picMsg_Resize()
'
    Dim lngTmp As Long
   
    On Error Resume Next
    
    lngTmp = picMsg.Height
    
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    
    rptNotify.Top = 0
    rptNotify.Left = 0
    rptNotify.Width = picMsg.Width
    rptNotify.Height = lngTmp
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'功能：阅读消息后删除消息，（双击消息或者选中消息后再按回车键）
'如果是阅读危机值消息则触发消息
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng病人ID As Long, lng主页ID As Long
    Dim lng医嘱ID As Long, str姓名 As String, str住院号 As String
    Dim str业务 As String, lng消息ID As Long
    Dim blnFinded As Boolean, blnOk As Boolean
    Dim strNO As String, str床号 As String
    Dim i As Long
    Dim str来源 As String
    Dim strTmp As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_消息).Value
                str业务 = .Item(C_业务).Value
                lng病人ID = Val(.Item(C_病人Id).Value)
                lng主页ID = Val(.Item(C_主页Id).Value)
                lng消息ID = Val(.Item(C_Id).Value)
                str姓名 = .Item(c_姓名).Value
                str住院号 = .Item(c_住院号).Value
                str床号 = .Item(c_床号).Value
                lngIndex = .Index
            End With
            
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    '是不是已经定位
                    blnFinded = InStr("_" & rptPati.SelectedRows(0).Record.Tag & "_", "_" & rptNotify.SelectedRows(0).Record.Tag & "_") > 0
                End If
            End If
            '如果没找到病人，且界面没选择在院病人列表时切换列表再查找一次
            If tbcPati.Item(tbcPatiEnu.E在院).Selected = False And Not blnFinded Then
                tbcPati.Item(tbcPatiEnu.E在院).Selected = True
                blnFinded = LocatePati(rptNotify.SelectedRows(0).Record.Tag)
            End If
            
            If blnFinded And tbcSub.Tag = "医嘱" And str业务 <> "" Then  '找到病人后，再决定是否定位医嘱
                lng医嘱ID = Val(str业务)
                If lng医嘱ID <> 0 Then
                    Call mclsAdvices.LocatedAdviceRow(lng医嘱ID)
                End If
            End If
          
            strTmp = ""
            If strNO = "ZLHIS_LIS_003" Then '检验
                strTmp = "ZLHIS_CIS_014"
            ElseIf strNO = "ZLHIS_PACS_005" Then '检查
                strTmp = "ZLHIS_CIS_025"
            End If
            If strTmp <> "" Then
                If Not (mclsMipModule Is Nothing) Then
                    If mclsMipModule.IsConnect Then
                        strSQL = "select 出院科室ID,当前病区ID from 病案主页 where 病人ID=[1] and 主页ID=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
                        Call ZLHIS_CIS_MsgReadAfter(mclsMipModule, strTmp, lng病人ID, str姓名, str住院号, , 2, _
                                lng主页ID, Val(rsTmp!当前病区ID & ""), Val(rsTmp!出院科室ID & ""), str床号, lng医嘱ID)
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_RECIPEAUDIT_002" Then
                If str业务 = "合理用药审方" Then
                    '合理用外部的审方消息   '调用接口判断当前病人的医嘱是不是都是通过的
                    blnOk = CheckZLPass(Me, lng病人ID, lng主页ID)
                    If blnOk Then
                        '未通过医嘱编辑窗体弹出来 修改
                        If tbcSub.Tag <> "医嘱" Then
                            For i = 0 To tbcSub.ItemCount - 1
                                If tbcSub.Item(i).Visible Then
                                    If tbcSub.Item(i).Tag = "医嘱" Then
                                        tbcSub.Item(i).Selected = True
                                        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                        
                        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                        If Not objControl Is Nothing Then
                            If objControl.Enabled Then objControl.Execute
                        End If
                    End If
                Else
                    '住院的处方审查只有不合格的然后调用作废
                    If tbcSub.Tag <> "医嘱" Then
                        For i = 0 To tbcSub.ItemCount - 1
                            If tbcSub.Item(i).Visible Then
                                If tbcSub.Item(i).Tag = "医嘱" Then
                                    tbcSub.Item(i).Selected = True
                                    cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_Blankoff, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
            
            If blnFinded And strNO = "ZLHIS_CIS_032" Then
                Call mclsDis.ShowDisRegist(Me, 1, Val(str业务), lng病人ID, lng主页ID)
            End If
            
            '病历质控消息
            If strNO = "ZLHIS_EMR_025" Then
                If tbcSub.Tag <> "新病历" Then
                    For i = 0 To tbcSub.ItemCount - 1
                        If tbcSub.Item(i).Visible Then
                            If tbcSub.Item(i).Tag = "新病历" Then
                                tbcSub.Item(i).Selected = True
                                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                                Exit For
                            End If
                        End If
                    Next
                End If
                If blnFinded Then
                    Set objControl = cbsMain.FindControl(, 3309, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then
                            objControl.Parameter = str业务
                            objControl.Execute
                            Call ReadMsg(lng病人ID, lng主页ID, strNO, str业务, lng消息ID)
                            Call ReadMsg批量(lng病人ID, lng主页ID, strNO)
                        End If
                    End If
                End If
                Exit Sub
            End If
            
            If blnFinded And strNO = "ZLHIS_CIS_033" Then
            '传染病报告反修改消息阅读
                blnOk = ReadMsgCIS033(lng病人ID, lng主页ID, str业务, lng消息ID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If blnFinded And strNO = "ZLHIS_BLOOD_006" Then
                If gobjPublicBlood Is Nothing And gbln血库系统 Then InitObjBlood
                blnOk = gobjPublicBlood.zlIsBloodMessageDone(2, lng病人ID, lng主页ID, 2, cboDept.ItemData(cboDept.ListIndex))
                If blnOk Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                Else
                    If FuncTraReaction(Val(str业务), mlngModul, False, IIf(InStr(1, str业务, ":") > 0, Val(Split(str业务, ":")(1)), 0)) Then
                        If gobjPublicBlood.zlIsBloodMessageDone(2, lng病人ID, lng主页ID, 2, cboDept.ItemData(cboDept.ListIndex)) Then
                            Call rptNotify.Records.RemoveAt(lngIndex)
            End If
                    End If
                End If
            End If
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOk = ReadMsg(lng病人ID, lng主页ID, strNO, str业务, lng消息ID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            Call rptNotify.Populate
        End If
    End If
End Sub

Private Function ReadMsgCIS033(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str标识 As String, ByVal lng消息ID As Long) As Boolean
'功能：传染病报告反修改消息阅读
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng文件ID As Long
    Dim objControl As CommandBarControl
    Dim i As Long
    
    On Error GoTo errH
    'conMenu_Edit_Modify 3003 修改按钮。
    lng文件ID = Val(Split(str标识, ",")(0))
    
    strSQL = "Select 1 From 疾病申报记录 where 文件ID=[1] and 处理状态=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng文件ID, 4)
    If rsTmp.RecordCount = 0 Then
    '把消息标记为已读
        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'ZLHIS_CIS_033',2,'" & UserInfo.姓名 & "'," & cboDept.ItemData(cboDept.ListIndex) & ",null," & lng消息ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    
    If "中华人民共和国传染病报告卡" = Sys.RowValue("电子病历记录", lng文件ID, "病历名称") Then
        '弹出来修改报告
        '先将卡片切换到医嘱卡片方便查找菜单
        If tbcSub.Tag <> "疾病报告" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "疾病报告" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                        Exit For
                    End If
                End If
            Next
        End If
        
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        If tbcSub.Selected.Tag = "疾病报告" And tbcSub.Selected.Visible = True Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    Else
        '弹出来修改报告
        Call mclsDis.ModifyDiseaseDoc(Me, lng文件ID, mlng病人ID, mlng主页ID, 2, mPatiInfo.科室ID)
    End If
    
    strSQL = "Select 1 From 疾病申报记录 where 文件ID=[1] and 处理状态=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng文件ID, 4)
    If rsTmp.RecordCount = 0 Then
    '把消息标记为已读
        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'ZLHIS_CIS_033',2,'" & UserInfo.姓名 & "'," & cboDept.ItemData(cboDept.ListIndex) & ",null," & lng消息ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mclspath_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：临床路径中查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
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

Private Sub mclsAdvices_PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean)
'功能：按编辑格式打印报告
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr诊疗报告, 报告ID, Not Preview, True)
End Sub

Private Sub mclsAdvices_ViewPACSImage(ByVal 医嘱ID As Long)
'功能：PACS观片处理
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(医嘱ID, Me, mPatiInfo.数据转出)
    End If
End Sub

Private Sub cboPages_Click()
'功能：选择某次住院记录时，读取相关的病人信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lng标识 As Long
    
    If cboPages.ListIndex = -1 Then Exit Sub
    If cboPages.ListIndex = mintPrePage Then Exit Sub
    mintPrePage = cboPages.ListIndex
    
    On Error GoTo errH
    
    mrsPati.Filter = "序号=" & cboPages.ItemData(cboPages.ListIndex)
    
    mPatiInfo.病人ID = mrsPati!病人ID
    mPatiInfo.主页ID = mrsPati!主页ID
    
    strSQL = "Select NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄, b.住院号, b.出院病床, b.医疗付款方式, d.信息值 As 医保号, b.险类, b.当前病况, c.名称 As 护理等级, Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期, b.出院日期, b.编目日期," & vbNewLine & _
            "       b.病人类型, b.状态, b.数据转出, b.出院科室id, b.当前病区id, a.住院次数, e.房间号" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 收费项目目录 C, 病案主页从表 D, 床位状况记录 E" & vbNewLine & _
            "Where a.病人id = b.病人id And a.病人id = [1] And b.主页id = [2] And b.护理等级id = c.Id(+) And b.病人id = d.病人id(+) And" & vbNewLine & _
            "      b.主页id = d.主页id(+) And d.信息名(+) = '医保号' And b.出院科室id = e.科室id(+) And b.病人id = e.病人id(+) And b.出院病床 = e.床号(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.病人ID, mPatiInfo.主页ID)
    
    With rsTmp
        '保险病人姓名红色显示
        lbl姓名(1).Caption = "" & !住院号
        lbl姓名(1).ForeColor = zlDatabase.GetPatiColor(NVL(!病人类型))
        lblPatiName(1).Caption = "" & !姓名
        lblPatiName(1).ToolTipText = lblPatiName(1).Caption
        lblPatiName(1).ForeColor = lbl姓名(1).ForeColor
        
        lbl医保号(1).Caption = NVL(!医保号)
        lbl护理(1).Caption = NVL(!护理等级)
        lbl付款(1).Caption = NVL(!医疗付款方式)
        
        '危重病人病况红色显示
        lbl病况(1).Caption = NVL(!当前病况)
        If NVL(!当前病况) = "危" Or NVL(!当前病况) = "重" Or NVL(!当前病况) = "急" Then
            lbl病况(1).ForeColor = &HC0&
        Else
            lbl病况(1).ForeColor = lbl医保号(1).ForeColor
        End If
        
        lbl入院(1).Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
        If Not IsNull(!出院日期) Then
            lbl入院(1).Caption = lbl入院(1).Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")
        
        End If
        
        lbl类型(1).Caption = NVL(!病人类型)
        lbl病室(1).Caption = IIf(IsNull(!房间号), "", "(" & !房间号 & ")") & !出院病床
        
        '诊断
        lblDiag(1).Caption = GetPatiDiagnose(mPatiInfo.病人ID, mPatiInfo.主页ID, 2)
        
        '病人信息
        mPatiInfo.状态 = NVL(!状态, 0)
        mPatiInfo.住院号 = NVL(!住院号)
        mPatiInfo.床号 = NVL(!出院病床)
                
        mPatiInfo.病区ID = NVL(!当前病区ID, 0)
        mPatiInfo.科室ID = NVL(!出院科室ID, 0)
                
        mPatiInfo.入院日期 = !入院日期
        If Not IsNull(!出院日期) Then
            mPatiInfo.出院日期 = !出院日期
        Else
            mPatiInfo.出院日期 = CDate(0)
        End If
        If Not IsNull(!编目日期) Then
            mPatiInfo.编目日期 = !编目日期
        Else
            mPatiInfo.编目日期 = CDate(0)
        End If
        mPatiInfo.住院次数 = NVL(!住院次数, 0)
        mPatiInfo.数据转出 = NVL(!数据转出, 0) <> 0
            
        Call SetPatiInfoCtlPos
    End With
    
    '显示病人图标
    Call ShowPati图标(mPatiInfo.病人ID, mPatiInfo.主页ID)
     
    '刷新子窗体数据
    Call SubWinRefreshData(tbcSub.Selected)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mclsEPRs_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    Call gobjKernel.ShowDiagHelp(Modal, Me, DiagnosisID)
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
    strTab = Decode(ObjectType, 1, "医嘱", 2, "病历", 3, "护理", 4, "护理", 5, "", 6, "医嘱", 7, "病历", 8, "病历", 9, "路径")
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
        If Me.tbcSub.Item(mlngOldIndex).Visible Then
            Call mclsTends.zlLocateData(IIf(ObjectType = 3, 1, 0))
        ElseIf Me.tbcSub.Item(mlngNewIndex).Visible Then
            Call mclsTendsNew.zlLocateData(IIf(ObjectType = 3, 1, 0))
        End If
    End If
    On Error GoTo errH
    '打开对应的对象
    Select Case ObjectType
    Case 1 '住院医嘱
    Case 2, 3, 7, 8 '住院病历,护理病历,疾病证明,知情文件
        If ObjectID = "0" Or ObjectID = "" Then Exit Sub
        If IsNumeric(ObjectID) Then
            Call gobjRichEPR.EditDocument(p住院医生站, Me, cboDept.ItemData(cboDept.ListIndex), ObjectID)
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
                        "       RAWTOHEX(a.Id) Docid, 2 Occasion, a.Sealed Besealed, nvl(e.Code,99) Docsecret, b.Subdoc_Id Subdocid,b.completor" & vbNewLine & _
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
            If NVL(rsEmr!completor) = "" Then
                If InStr(strPrivs, ";文档书写;") > 0 Then '有书写权限
                    Call gobjEmr.OpenFormForModifyDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, NVL(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, Val(rsEmr!Occasion), Val(rsEmr!besealed), Val(rsEmr!docsecret), NVL(rsEmr!subdocid), "02", strPrivs)
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
                    Call gobjEmr.OpenFormForAuditDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, NVL(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, Val(rsEmr!Occasion), Val(rsEmr!besealed), Val(rsEmr!docsecret), NVL(rsEmr!subdocid), "02", strPrivs)
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
        If InStr(mstrPrivs, "首页整理") > 0 Then
            Call ExecuteEditMediRec(True)
        Else
            MsgBox "你没有【住院医生工作站】的首页整理权限，不能查看编辑首页!", vbInformation, gstrSysName
        End If
    Case 6 '医嘱报告
        If CLng(ObjectID) = 0 Then Exit Sub
        Call mclsAdvices.zlSeekAndViewEPRReport(ObjectID)
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPatiIn_Resize()
    Dim i As Long, Y As Long
    On Error Resume Next
    
    For i = 0 To picPara.Count - 1
        picPara(i).Width = picPatiIn.ScaleWidth
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

Private Sub rptNotify_SelectionChanged()
    Dim strCurPati As String
    Dim lngIndex As Long
    Dim lng医嘱ID As Long
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    With rptNotify.SelectedRows(0)
        
        lngIndex = rptNotify.FocusedRow.Record.Index
        If rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_LIS_003" Or rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_PACS_005" Then '这两个消息的业务列存的是 医嘱id，病人来源
            lng医嘱ID = Val(Split(rptNotify.Rows(lngIndex).Record(C_业务).Value, ",")(0))
        ElseIf rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_BLOOD_006" Then
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
            If rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_CIS_020" Then
                If tbcPati.Selected.Index <> 4 Then
                    tbcPati.Item(4).Selected = True
                End If
            End If
            '自动寻找并切换显示当前提醒的病人
            If Not LocatePati(.Record.Tag) Then
                Call LoadPatients
                Call LocatePati(.Record.Tag)
            End If
        End If
        
        If lng医嘱ID <> 0 And tbcSub.Tag = "医嘱" Then
            Call mclsAdvices.LocatedAdviceRow(lng医嘱ID)
        End If
        
    End With
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
    ElseIf Item.Tag = "出院" Then
        picPara(1).Visible = True
    ElseIf Item.Tag = "转出" Then
        picPara(3).Visible = True
    ElseIf Item.Tag = "会诊" Then
        picPara(2).Visible = True
    End If
    
    chkFilter.Visible = IIf(mstrFindType = "住院号" And tbcPati.Selected.Tag = "出院", True, False)
    Call picPati_Resize
    
    Call picPatiIn_Resize
    If Me.Visible Then
        Call LoadPatients
        Call LoadNotify '刷新医嘱提醒
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
            Case "住院一览"
                Set objItem = tbcSub.InsertItem(Index, "住院一览", mcolSubForm("_住院一览").hwnd, 0)
                objItem.Tag = "住院一览"
            Case "路径"
                Set objItem = tbcSub.InsertItem(Index, "临床路径", mcolSubForm("_路径").hwnd, 0)
                objItem.Tag = "路径"
            Case "医嘱"
                Set objItem = tbcSub.InsertItem(Index, "医嘱信息", mcolSubForm("_医嘱").hwnd, 0)
                objItem.Tag = "医嘱"
            Case "病历"
                Set objItem = tbcSub.InsertItem(Index, "病历信息", mcolSubForm("_病历").hwnd, 0)
                objItem.Tag = "病历"
            Case "新病历"
                Set objItem = tbcSub.InsertItem(Index, "电子病历", mcolSubForm("_新病历").hwnd, 0)
                objItem.Tag = "新病历"
            Case "护理"
                Set objItem = tbcSub.InsertItem(Index, "护理信息", mcolSubForm("_护理").hwnd, 0)
                objItem.Tag = "护理"
            Case "新版护理"
                Set objItem = tbcSub.InsertItem(Index, "护理信息", mcolSubForm("_新版护理").hwnd, 0)
                objItem.Tag = "新版护理"
            Case "护理病历"
                Set objItem = tbcSub.InsertItem(Index, "护理病历", mcolSubForm("_护理病历").hwnd, 0)
                objItem.Tag = "护理病历"
            Case "监护"
                Set objItem = tbcSub.InsertItem(Index, "护理监护", mcolSubForm("_监护").hwnd, 0)
                objItem.Tag = "监护"
            Case "疾病报告"
                Set objItem = tbcSub.InsertItem(Index, "疾病报告", mcolSubForm("_疾病报告").hwnd, 0)
                objItem.Tag = "疾病报告"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
    
    If Item.Tag = "护理评估评分" Then
        Me.dkpMain.Options.UseSplitterTracker = True '处理网页控件不能实时拖动
    Else
        Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    End If
     
    '刷新子窗体对应的CommandBar
    Call SubWinDefCommandBar(Item)
    '91136:将此语句放到SubWinDefCommandBar之后，以前在前面导致先新后老时，老板菜单无法刷新。
    If mblnIsNot Then mblnIsNot = False: Exit Sub
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

Private Sub cboDept_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    End If
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

Private Sub cboDept_Click()
'功能：刷新界面数据
'说明：从该事件开始会不重复引发相关的数据读取
    Dim lng部门ID As Long, i As Long, lngidx As Long
    Dim blnIn病区 As Boolean, rsTmp As Recordset, str科室IDs As String
    
    If cboDept.ListIndex = -1 Then
        Call ClearPatiInfo
        Call SubWinRefreshData(tbcSub.Selected)
        
        mblnICU = False
        mstrPreNotify = ""
        rptNotify.Records.DeleteAll
        rptNotify.Populate
        rptNotify.TabStop = False
        Exit Sub
    End If
    cboDept.Tag = cboDept.ListIndex
    mblnReturn = True
    If cboDept.ListIndex = mintPreDept Then Exit Sub
    mintPreDept = cboDept.ListIndex
    lng部门ID = cboDept.ItemData(cboDept.ListIndex)
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    rptNotify.Populate
    rptNotify.TabStop = False
    
    
    '是否非本科的ICU室
    str科室IDs = "," & GetUser科室IDs(True) & ","
    mblnICU = Sys.DeptHaveProperty(lng部门ID, "ICU")
    If mblnICU = True Then
        If mintDeptView = 0 Then
            mblnICU = InStr(str科室IDs, "," & lng部门ID & ",") = 0
        Else
            '病区显示，判断病区对应的所有科室都是操作员的科室时，才成立
            blnIn病区 = True
            Set rsTmp = Sys.RowValue("病区科室对应", lng部门ID, , "病区ID")
            Do While Not rsTmp.EOF
                If InStr(str科室IDs, "," & rsTmp!科室ID & ",") = 0 Then
                    blnIn病区 = False
                End If
                rsTmp.MoveNext
            Loop
            mblnICU = Not blnIn病区
        End If
    End If
    
    Call Sys.DeptHaveProperty(lng部门ID, IIf(mintDeptView = 0, "临床", "护理"), mblnOutDept)
        
    '关闭业务窗体
    Set mclsInOutMedRec = Nothing
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
    End If
    
    '重新读取病人
    Call LoadPatients
    
    '初始化病区标记
    Call Init图标信息(lng部门ID)
    
    '显示临床路径卡片
    lngidx = -1
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub(i).Tag = "路径" Then
            lngidx = i
            Exit For
        End If
    Next
    If lngidx >= 0 Then
        If HavePath(lng部门ID) = False Then
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
    If Visible And rptPati.Visible Then rptPati.SetFocus
 
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
        If mclsWardMonitor.Enabled = False Or InStr(GetInsidePrivs(p住院医生站), "护理监护") = 0 Then
            objCol.Visible = False
        End If
        Set objCol = .Columns.Add(col_性别, "性别", 30, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 30, True)
        Set objCol = .Columns.Add(col_费别, "费别", 55, True)
        Set objCol = .Columns.Add(col_科室, "科室", 70, True)
        Set objCol = .Columns.Add(col_病区, "病区", 70, True): objCol.Visible = False
        Set objCol = .Columns.Add(col_住院医师, "住院医师", 55, True)
        Set objCol = .Columns.Add(col_入院日期, "入院日期", 106, True)
        Set objCol = .Columns.Add(col_出院日期, "出院日期", 106, True)
        Set objCol = .Columns.Add(col_病人类型, "病人类型", 106, True)
        
        Set objCol = .Columns.Add(col_医嘱ID, "医嘱ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_发送号, "发送号", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_执行状态, "执行状态", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_执行科室ID, "执行科室ID", 0, False): objCol.Visible = False
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_就诊卡, "就诊卡", 0, False)
        Else
            Set objCol = .Columns.Add(col_就诊卡, "就诊卡", 70, True)
        End If
        Set objCol = .Columns.Add(col_住院天数, "住院天数", 56, True)
        Set objCol = .Columns.Add(col_单病种, "单病种", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_婴儿科室ID, "婴儿科室ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_婴儿病区ID, "婴儿病区ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_西医诊断, "西医诊断", 106, True)
        Set objCol = .Columns.Add(col_中医诊断, "中医诊断", 106, True)
        Set objCol = .Columns.Add(COL_申请序号, "申请序号", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_传染病, "传染病", 106, True)
        Set objCol = .Columns.Add(col_责任护士, "责任护士", 55, True)
        Set objCol = .Columns.Add(col_留观号, "留观号", 62, True)
        Set objCol = .Columns.Add(col_身份证号, "身份证号", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_是否急诊, "急", 30, False): objCol.Visible = False
        
        
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
        .SortOrder.Add .Columns(col_是否急诊)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_审查)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns(col_床号)
        .SortOrder(2).SortAscending = True
    End With
    
    With rptNotify
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_病人Id, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_主页Id, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(c_住院号, "住院号", 62, True)
        Set objCol = .Columns.Add(c_床号, "床号", 40, True)
        Set objCol = .Columns.Add(C_状态, "状态", 150, True)
         
        Set objCol = .Columns.Add(C_消息, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_序号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_日期, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_业务, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_Id, "", 0, False): objCol.Visible = False
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
    
    '------点击图标后弹出来的列表
    With rptTBPati
        Set objCol = .Columns.Add(CI_图标1, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(CI_图标2, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(CI_图标3, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(CI_床号, "床号", 40, True)
        Set objCol = .Columns.Add(CI_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(CI_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(CI_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(CI_住院号, "住院号", 60, True)
        Set objCol = .Columns.Add(CI_入院日期, "入院日期", 70, True)
        Set objCol = .Columns.Add(CI_出院日期, "出院日期", 70, True)
        Set objCol = .Columns.Add(CI_病人类型, "病人类型", 100, True)
        
                
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有病人..."
        End With
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList zlCommFun.GetPaitSignImageList(0)
        
        .SortOrder.Add .Columns.Find(CI_床号)
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picMsg.hwnd
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
    If mblnInView Then
        With mfrmInView
            .Left = 0: .Top = 0
            .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - stbThis.Height
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim strTmp As String
    Dim curDate As Date
    Dim blnSetup As Boolean
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("病人查找方式", mintFindType, glngSys, p住院医生站, blnSetup)
    strTmp = ""
    For i = 0 To chk病况条件.UBound
        strTmp = strTmp & IIf(chk病况条件(i).Value = 1, "1", "0")
    Next
    Call zlDatabase.SetPara("当前病况过滤", strTmp, glngSys, p住院医生站, blnSetup)
    Call zlDatabase.SetPara("按小组显示", Val(chkByTeam.Value), glngSys, p住院医生站, blnSetup)
    
    strTmp = chkHZ(0).Value & chkHZ(1).Value
    If strTmp = "00" Then
        strTmp = "11"
    End If
    strTmp = chkOut.Value & strTmp
    Call zlDatabase.SetPara("会诊病人过滤", strTmp, glngSys, p住院医生站, blnSetup)
    
    '病人范围
    curDate = zlDatabase.Currentdate
    Call zlDatabase.SetPara("最近转出天数", Val(txtChange.Text), glngSys, mlngModul, blnSetup)
    
     If Not mclsInOutMedRec Is Nothing Then
        If Not mclsInOutMedRec.FormUnLoad Then
            Cancel = True
            Exit Sub
        Else
            Cancel = False
            Set mclsInOutMedRec = Nothing
        End If
    End If
    
    Call SetAllPati图标(1)
        
    If Me.Visible Then
        If Not tbcSub.Selected Is Nothing Then
            Call zlDatabase.SetPara("医护功能", tbcSub.Selected.Tag, glngSys, p住院医生站, blnSetup)
        End If
        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 And Not mbln单个病人 Then
            Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        End If
        If Not tbcPati.Selected Is Nothing Then
            Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", tbcPati.Selected.Index)
        End If
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(chkOutByTeam), "chkOutByTeam", chkOutByTeam.Value)
        '公共部件固定按第一个控件的样式保存，工作站部件如果第一个是打印，则固定是图标样式,所以需恢复为其它按钮的样式
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsEMR = Nothing
    Set mclsAdvices = Nothing
    Set mclsEPRs = Nothing
    Set mclsTends = Nothing
    Set mclsTendsNew = Nothing
    Set mclsTendEPRs = Nothing
    Set mclsWardMonitor = Nothing
    Set mclsPath = Nothing
    Set mobjEPRDoc = Nothing
    Set mobjSquareCard = Nothing
    Set mclsReg = Nothing
    mblnIsInit = False

    If Not mclsChildQuestion Is Nothing Then
        Set mclsChildQuestion = Nothing
    End If
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
        Set mfrmResponse = Nothing
    End If
    
    If Not mfrmActive Is Nothing Then
        Unload mfrmActive
        Set mfrmActive = Nothing
    End If
    If Not mfrmInView Is Nothing Then
        Unload mfrmInView
        Set mfrmInView = Nothing
    End If
    
    If Not mfrmParent Is Nothing Then
        If mfrmParent.frmHide Then mfrmParent.UnloadForm
    End If
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Set mobjKernel = Nothing
    Set mclsDis = Nothing
    Set gobjPublicPacs = Nothing
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
    Set mobjPatient = Nothing
    Set mrsPati = Nothing
    mstrAllPatis = ""
    Set mrsNotes = Nothing
    mstrList主题 = ""
    Set mrsPatiNotes = Nothing
    Set mrsPati汇总 = Nothing
    mbln危急值 = False
    Set mclsDisease = Nothing
    Set mobjReport = Nothing
    If Not mFrmConsultation Is Nothing Then
        Unload mFrmConsultation
    End If
    Set mFrmConsultation = Nothing
    mlng会诊医嘱ID = 0
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
    lblDept.Top = (cboDept.Height - lblDept.Height) / 2 + 30
    lblDept.Left = lblDept.Top
    cboDept.Top = 30
    cboDept.Left = lblDept.Left + lblDept.Width + 30
    cboDept.Width = picPati.ScaleWidth - cboDept.Left - lblDept.Left
    
    lblFind.Left = lblDept.Left
    lblFind.Top = lblDept.Top + lblDept.Height + 120
    lblFind.Width = lblDept.Width
    PatiIdentify.Left = cboDept.Left
    PatiIdentify.Top = lblFind.Top - 40
    PatiIdentify.Width = IIf(chkFilter.Visible, cboDept.Width - 60 - chkFilter.Width, cboDept.Width)
    
    chkFilter.Left = PatiIdentify.Left + PatiIdentify.Width + 30
    chkFilter.Top = PatiIdentify.Top
    
    picIconPati.Top = PatiIdentify.Top + PatiIdentify.Height + 10
    picIconPati.Left = 0
    picIconPati.Width = cboDept.Width + cboDept.Left
    picIconPati.Height = IIf(mbytSize = 0, 350, 450)
    
    tbcPati.Left = 0
    If picIconPati.Visible Then
        tbcPati.Top = picIconPati.Top + picIconPati.Height + 30
    Else
        tbcPati.Top = PatiIdentify.Top + PatiIdentify.Height + 30
    End If
    tbcPati.Width = picPati.ScaleWidth
    tbcPati.Height = picPati.ScaleHeight - tbcPati.Top - IIf(fra审查.Visible, fra审查.Height, 0)
    
    fra审查.Left = 0
    fra审查.Top = tbcPati.Top + tbcPati.Height
    fra审查.Width = picPati.ScaleWidth
    
    picPatiIn.Width = picPati.ScaleWidth
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    'Panne中的Report控件需要强行处理光标顺序
    '无数据时不能捕获到vbKeyTab
    If KeyCode = vbKeyTab Then
        If Shift = vbShiftMask Then
            If cboDept.Enabled Then cboDept.SetFocus
        Else
            If rptNotify.Visible And rptNotify.TabStop Then
                On Error Resume Next
                rptNotify.SetFocus
            Else
                If cboPages.Enabled Then cboPages.SetFocus
            End If
        End If
    Else
        cboDept.SetFocus
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
                Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
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
    Dim str身份证号 As String
    Dim strTmp As String
    Dim str病人IDs As String
    
    mblnIsNot = False
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '非正常情况
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
                        objRow.Expanded = False '
                        blnPopulate = True
                    End If
                End If
            Next
        End If
        
        If Not .GroupRow And strCurPati <> "" Then
            mlng病人ID = Val(.Record(col_病人Id).Value)
            mlng主页ID = Val(.Record(col_主页ID).Value)
            str身份证号 = .Record(col_身份证号).Value
            
            If InStr(strCurPati, "_") > 0 Then
                mPatiInfo.婴儿 = Val(Split(strCurPati, "_")(1))
            Else
                mPatiInfo.婴儿 = -1
            End If

            mbln接受会诊 = False
            mlng会诊医嘱ID = 0
            If tbcPati.Selected.Tag = "会诊" Then
                If Val(.Record(col_医嘱ID).Value) <> 0 Then
                    mbln接受会诊 = Get接受会诊
                End If
                mlng会诊医嘱ID = Val(.Record(col_医嘱ID).Value)
            End If
            
            LockWindowUpdate Me.hwnd
            
            '验证身份证号
            If str身份证号 <> "" Then
                If mobjPatient Is Nothing Then
                    On Error Resume Next
                    Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                    err.Clear: On Error GoTo 0
                    If mobjPatient Is Nothing Then
                        MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
                    Else
                        Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
                    End If
                End If
                strTmp = ""
                If Not mobjPatient Is Nothing Then
                    If mobjPatient.CheckPatiIdcard(str身份证号) Then
                        strTmp = str身份证号
                    End If
                End If
                str身份证号 = strTmp
            End If
            
            On Error GoTo errH
            
            If str身份证号 <> "" Then
                strSQL = "select a.病人id from 病人信息 a where a.病人id<>[1] and a.身份证号=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, str身份证号)
                Do While Not rsTmp.EOF
                    str病人IDs = str病人IDs & "," & rsTmp!病人ID
                    rsTmp.MoveNext
                Loop
                If str病人IDs <> "" Then
                    str病人IDs = mlng病人ID & str病人IDs
                End If
            End If
            
            If str病人IDs = "" Then
                strSQL = "Select rownum as 序号,病人id,主页ID,NVL(病人性质,0) 病人性质,住院号,To_Char(入院日期,'YYYY-MM-DD HH24:MI') as 入院日期 From 病案主页 Where 主页ID<>0 And 病人ID=[1] Order by 主页ID Desc,序号 Desc"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
                Call Cbo.SetListWidth(cboPages.hwnd, cboPages.Width * 1.5)
            Else
                strSQL = "Select rownum as 序号,a.病人id,a.主页ID,NVL(a.病人性质,0) 病人性质,a.住院号,To_Char(a.入院日期,'YYYY-MM-DD HH24:MI') as 入院日期 From 病案主页 a" & _
                " where a.主页ID<>0 And A.病人ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
                " order by a.病人id desc,a.主页id desc,序号 Desc"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str病人IDs)
                Call Cbo.SetListWidth(cboPages.hwnd, cboPages.Width * 4)
            End If
            Set mrsPati = zlDatabase.CopyNewRec(rsTmp)
            cboPages.Clear
            Do While Not rsTmp.EOF
            
                If str病人IDs = "" Then
                    strTmp = "第 " & rsTmp!主页ID & " 次" & Decode(rsTmp!病人性质, 1, "(门诊留观)", 2, "(住院留观)", "")
                Else
                    strTmp = "第 " & rsTmp!主页ID & IIf(IsNull(rsTmp!住院号), "", "_" & rsTmp!住院号) & " 次" & Decode(rsTmp!病人性质, 1, "(门诊留观)", 2, "(住院留观)", "") & ":" & rsTmp!入院日期
                End If
                
                cboPages.AddItem strTmp
                cboPages.ItemData(cboPages.NewIndex) = rsTmp!序号
                If rsTmp!主页ID = mlng主页ID And rsTmp!病人ID = mlng病人ID Then
                    Call Cbo.SetIndex(cboPages.hwnd, cboPages.NewIndex)
                End If
                rsTmp.MoveNext
            Loop
            If cboPages.ListIndex = -1 Then
                Call Cbo.SetIndex(cboPages.hwnd, 0)
            End If
                        Call cboPages_Click
            
            mintPrePage = -1
            On Error GoTo errH
            If GetInsidePrivs(p护理记录管理, True) <> "" Then
                strSQL = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
                If mPatiInfo.数据转出 Then
                    strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
                End If
                
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
                If rsTmp.RecordCount > 0 Then
                    Me.tbcSub.Item(mlngOldIndex).Visible = True
                    Me.tbcSub.Item(mlngNewIndex).Visible = False
                    Me.tbcSub.Item(mlngNewIndex + 1).Visible = False '新版的同时隐藏护理病历
                    If tbcSub.Item(mlngOldIndex).Selected Or tbcSub.Item(mlngNewIndex).Selected Or tbcSub.Item(mlngNewIndex + 1).Selected Then
                        mblnIsNot = Not tbcSub.Item(mlngOldIndex).Selected
                        Me.tbcSub.Item(mlngOldIndex).Selected = True
                    End If
                Else
                    Me.tbcSub.Item(mlngNewIndex).Visible = True
                    Me.tbcSub.Item(mlngOldIndex).Visible = False
                    Me.tbcSub.Item(mlngNewIndex + 1).Visible = True '新版的同时显示护理病历
                    If tbcSub.Item(mlngOldIndex).Selected Or tbcSub.Item(mlngNewIndex).Selected Then
                        mblnIsNot = Not tbcSub.Item(mlngNewIndex).Selected
                        Me.tbcSub.Item(mlngNewIndex).Selected = True
                    End If
                End If
            End If
            
            
            Call LoadPatiAllergy(mlng病人ID, cbo过敏)
                        
            '出院病人读取是否已提交审查
            If (Int(Val(.Record(col_类型).Value)) = pt出院 Or Int(Val(.Record(col_类型).Value)) = pt死亡) Or (tbcPati.Selected.Tag = "出院" And Val(.Record(col_病人Id).Value) <> 0) Then
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
            lblFee(1).Caption = GetPati费用信息(mlng病人ID, mlng主页ID) & IIf(InStr(mstrPrivs, "药占比查询") = 0, "", Get住院费用药占比(mlng病人ID, mlng主页ID))
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
            intTmp = Get病人医嘱打印(mlng病人ID, mlng主页ID)
            lblPrint(1).Caption = IIf(intTmp = 0, "未打印", IIf(intTmp = 1, "部分打印", "全部打印"))
            On Error Resume Next
            If Visible And rptPati.Visible Then rptPati.SetFocus
            If err.Number <> 0 Then err.Clear
            On Error GoTo errH
        Else
            Call ClearPatiInfo
            '按无数据刷新子窗体
            Call SubWinRefreshData(tbcSub.Selected)
            
            stbThis.Panels(2).Text = stbThis.Panels(2).Tag
            lblFee(1).Caption = ""
            lblFluid(1).Caption = ""
            lblPrint(1).Caption = ""
        End If
    End With
    '三方控件得不到焦点，所以通过设置其他控件得到焦点的方式解决这个BUG（74488）
    On Error Resume Next
    If Me.Visible Then
        txtTestBug.SetFocus
        rptPati.SetFocus
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
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

Private Function Get接受会诊() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "SELECT A.接收时间,A.接收人 FROM 病人医嘱发送 A where 医嘱ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rptPati.SelectedRows(0).Record(col_医嘱ID).Value))
    If Not rsTmp.EOF Then
        Get接受会诊 = IIf(rsTmp!接收人 & "" = "", False, True)
    Else
        Get接受会诊 = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    Dim curDate As Date, intDay As Integer
    Dim intType As Integer
    
    '病人显示范围
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院医生站, 7))
    '如果大于30天就取缺省值
    If mintChange > 30 Then mintChange = 7
    
    '出院病人时间范围，固定为过去3天
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    mdtOutBegin = Format(mdtOutEnd - 3, "yyyy-MM-dd 00:00:00")
    
    '会诊病人时间范围，固定为过去3天
    mdtMeetEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    mdtMeetBegin = Format(mdtMeetEnd - 3, "yyyy-MM-dd 00:00:00")
    
    '自动刷新病历审阅间隔
    mintNotify = Val(zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, p住院医生站))
    mintNotifyDay = Val(zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, p住院医生站, 1))
    mstrNotify = zlDatabase.GetPara("自动刷新内容", glngSys, p住院医生站, "0000")
    mbln消息语音 = Val(zlDatabase.GetPara("启用语音提示", glngSys, p住院医生站)) = 1
    mbln危急值弹窗 = Val(zlDatabase.GetPara("住院危急值弹窗提醒", glngSys, p住院医生站, 1)) = 1
    
    '部门显示方式
    mintDeptView = Val(zlDatabase.GetPara("部门显示方式", glngSys, p住院医生站, , , , intType))
    mblnDeptViewEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    
    '病案审查反馈天数
    mlngMedRedDay = Val(zlDatabase.GetPara("病案审查反馈天数", glngSys, p住院医生站))
    '字体大小
    mbytSize = Val(zlDatabase.GetPara("字体", glngSys, p住院医生站, "0"))
    
    '病案首页标准
    mintMecStandard = Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
    
    mlngSource = IIf(mbytSize = 1, 0, 999)

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
    Set rsTmp = GetDataToDepts
    
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
    Dim objPt As ReportRecord '分组父结点
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objRow As ReportRow, intBedLen As Integer
    Dim strPatiRow As String, lngPatiRow As Long, blnTeam As Boolean, blnOutTeam As Boolean
    
    Dim strSQL As String, strMonitor As String
    Dim i As Long, j As Long
    Dim lngCount(0 To 7) As Long, strState As String    '用于显示病人分类统计数目
    Dim str病况条件 As String, strFilter As String
    Dim strTmpDate As String                            '转出病人查询时间范围条件
    Dim blnIsFind As Boolean                            '判断是否是查找住院号及住院号是否为空
    Dim strTmpOut As String                             '查询出院病人
    Dim strICUSQL As String
    Dim strICUOutSQL As String                          '转出病人用的经治医师
    Dim lng申请序号 As Long
    
    Dim rs传染病状态 As ADODB.Recordset
    Dim blnDo传染病状态 As Boolean, blnTeamGroup As Boolean, blnVisible审查 As Boolean
    Dim rsBaby As ADODB.Recordset
    Dim strSQLBaby As String
    Dim objBabyParent As ReportRecord
    Dim lngCol As Long
    Dim strPre病人类型 As String
    Dim str字段诊断 As String, strPre小组名 As String
    Dim strTab病人 As String, str小组名 As String, str类型 As String
    Dim str会诊状态 As String
    
    mblnUnRefresh = True
    Screen.MousePointer = 11
    On Error GoTo errH
    
    '当页面下拉框清空，F5刷新，应该恢复上一个的值
    If cboDept.ListIndex = -1 Then Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    '判断是否是查找住院号及住院号是否为空
    If mstrFindType = "住院号" And Trim(PatiIdentify.Text) <> "" Then blnIsFind = True
    
    If blnIsFind Then
        '由于取消了显示出院病人参数，所以查找时显示查找的人员和时间范围的人员
        If chkFilter.Value = 1 And chkFilter.Visible = True Then
            strTmpOut = " And (B.出院日期 Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') And b.住院号=[6] And B.出院日期 is Not Null) "
        Else
            strTmpOut = " And (B.出院日期 Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')  Or (b.住院号=[6]  And B.出院日期 is Not Null )) "
        End If
    Else
        '用户不是查找，显示参数对应时间内的病人
        strTmpOut = " And B.出院日期 Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
    End If
    Set mrsPatiNotes = Nothing
    mstrAllPatis = ""
    blnTeam = chkByTeam.Value = 1
    blnOutTeam = chkOutByTeam.Value = 1
    '床位长度固定为10
    intBedLen = 10
    
    'Start 病况条件
    str病况条件 = ""
    strFilter = ""
    strTmpDate = ""
    For i = 0 To chk病况条件.UBound
        If chk病况条件(i).Value = 1 Then
            str病况条件 = str病况条件 & "," & chk病况条件(i).Caption
        End If
    Next
    str病况条件 = Mid(str病况条件, 2)
    If Not (UBound(Split(str病况条件, ",")) = chk病况条件.UBound Or str病况条件 = "") Then
        strFilter = " And Instr(','||[4]||',',','||B.当前病况||',')>0"
    End If
    
    '会诊状态
    str会诊状态 = chkHZ(0).Value & chkHZ(1).Value
    
    If str会诊状态 = "01" Then
        str会诊状态 = " and d.执行状态=1"
    ElseIf str会诊状态 = "10" Then
        str会诊状态 = " and nvl(d.执行状态,0)<>1"
    Else
        str会诊状态 = ""
    End If
    
    If mintChange = 0 Then
        strTmpDate = ""
    Else
        strTmpDate = " And C.终止时间 Between Sysdate-[3] And Sysdate "
    End If
    
    If mblnICU Then
        If InStr(mstrPrivs, "全院病人") = 0 And InStr(mstrPrivs, "本科病人") = 0 Then
            strICUSQL = " And B.住院医师=[2] "
            strICUOutSQL = " And C.经治医师=[2] "

        ElseIf InStr(mstrPrivs, "全院病人") > 0 Then
            strICUSQL = ""
            strICUOutSQL = strICUSQL
        Else
            strICUSQL = " And exists(select 1 from 病人变动记录 x where " & _
                " b.病人id=x.病人id and b.主页id=x.主页id and x.终止原因 in(2,3,15) and instr(','||[7]||',' , ','||x.科室id||',')>0) "
            strICUOutSQL = strICUSQL
        End If
        
    End If
    
    If cboDept.ListIndex <> -1 Then
        If tbcPati.Selected.Tag = "在院" Or tbcPati.Selected.Tag = "出院" Then
            str字段诊断 = ",first_value(Decode(Sign(h.诊断类型-10),-1,h.诊断描述,'')) " & _
                " Over(partition By h.病人id,H.主页ID Order By sign(h.诊断类型-10),decode(h.记录来源,4,0,h.记录来源) desc,Decode(h.诊断类型," & Decode(tbcPati.Selected.Tag, "在院", "1,1,2,2,3,3,0", "出院", "3,3,0") & ") DESC,h.诊断次序) As 西医诊断"
            If Sys.DeptHaveProperty(cboDept.ItemData(cboDept.ListIndex), "中医科") Then
                str字段诊断 = str字段诊断 & ",first_value(Decode(Sign(h.诊断类型-10),1,h.诊断描述,'')) " & _
                " Over(partition By h.病人id,H.主页ID Order By sign(h.诊断类型-10) desc,decode(h.记录来源,4,0,h.记录来源) desc,Decode(h.诊断类型," & Decode(tbcPati.Selected.Tag, "在院", "11,1,12,2,13,3,0", "出院", "13,3,0") & ") DESC,h.诊断次序) As 中医诊断"
            Else
                str字段诊断 = str字段诊断 & ",null as 中医诊断"
            End If
        Else
            str字段诊断 = ",Null As 西医诊断,Null As 中医诊断"
        End If
        If mintDeptView = 0 Then
            '在院病人
            If tbcPati.Selected.Tag = "在院" Or tbcPati.Selected.Tag = "待入住" Then
                If blnTeam And tbcPati.Selected.Tag = "在院" Then    '按小组模式
                    strSQL = _
                        "Select Distinct Decode(B.状态,1,0,3,3,Decode(G.ID,Null,2,1)) as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,G.Id 组ID,G.名称 as 小组名," & _
                        " Decode(B.状态,1,'待入住病人',3,'预出院病人',Decode(G.名称,Null,'在院病人',G.名称)) as 类型,A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号," & _
                        " NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,NULL as 科室,B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号," & _
                        " B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态,-Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容," & _
                        " Nvl(b.路径状态,-1) 路径状态,A.身份证号,trunc(sysdate)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID" & str字段诊断 & _
                        " From 病人信息 A,病案主页 B,临床医疗小组 G,病人新生儿记录 Q,在院病人 R,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null)" & _
                        " And (R.科室ID=[1] Or b.婴儿科室ID=[1]) And a.病人ID=R.病人ID And  A.当前科室ID=R.科室ID And B.医疗小组ID=G.ID(+)" & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "And (B.医疗小组ID is Null or b.医疗小组id in(select id from 临床医疗小组 where 科室id =[1]))", _
                            " And (B.医疗小组ID is Null And B.住院医师=[2] or B.医疗小组ID in (Select 小组id From 医疗小组人员 Where 人员id = [5]))") & _
                        strICUSQL & " And  B.封存时间 is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "待入住", " And Nvl(B.状态,0)=1", " And Nvl(B.状态,0)<>1")
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.状态,1,0,3,3,Decode(B.住院医师,[2],1,2)) as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,Null as 组ID," & _
                        " Decode(B.状态,1,'待入住病人',3,'预出院病人',Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的在院病人','在院病人')) as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,A.身份证号,B.住院号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,NULL as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID" & _
                        str字段诊断 & _
                        " From 病人信息 A,病案主页 B,病人新生儿记录 Q,在院病人 R,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID  And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null)" & _
                        " And (R.科室ID=[1] Or b.婴儿科室ID=[1]) And a.病人ID=R.病人ID And A.当前科室ID=R.科室ID " & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                        strICUSQL & _
                        " And B.封存时间 is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "待入住", " And B.状态=1", " And B.状态<>1")
                End If
                strSQLBaby = "Select q.病人id, q.主页id, q.序号, q.婴儿姓名, q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
                    " From 病人信息 A,病案主页 B,病人新生儿记录 Q,在院病人 R" & _
                    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID  And b.病人id=q.病人ID And b.主页ID=q.主页ID" & _
                    " And (R.科室ID=[1] Or b.婴儿科室ID=[1]) And a.病人ID=R.病人ID And A.当前科室ID=R.科室ID " & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                    strICUSQL & " And  B.封存时间 is NULL" & strFilter & _
                    IIf(tbcPati.Selected.Tag = "待入住", " And B.状态=1", " And B.状态<>1")
            ElseIf tbcPati.Selected.Tag = "出院" Then
                '出院病人:出院病人可能已有多次住院
                If blnOutTeam Then
                    strSQL = _
                        "Select Distinct Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],5.1,5.2), Decode(B.住院医师,[2],4.1,4.2)) as 排序," & _
                        " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,G.Id as 组ID,G.名称 as 小组名," & _
                        " Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的死亡病人','其他死亡病人'),Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的出院病人',Decode(g.名称, Null, '其他出院病人', g.名称))) as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,NULL as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(B.出院日期)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数," & _
                        " q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID,B.责任护士" & str字段诊断 & _
                        " From 病人信息 A,病案主页 B,临床医疗小组 G,病人新生儿记录 Q,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID+0=[1] And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "And (B.医疗小组ID is Null or b.医疗小组id in(select id from 临床医疗小组 where 科室id =[1]))", _
                            " And (B.医疗小组ID is Null And B.住院医师=[2] or B.医疗小组ID in(Select 小组id From 医疗小组人员 Where 人员id = [5]))") & _
                        strICUSQL & " And B.封存时间 is NULL And b.医疗小组id = g.Id(+)"
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],5.1,5.2),Decode(B.住院医师,[2],4.1,4.2)) as 排序," & _
                        " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,Null as 组ID," & _
                        " Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的死亡病人','其他死亡病人'),Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的出院病人','其他出院病人')) as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,NULL as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(B.出院日期)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数," & _
                        " q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID,B.责任护士" & str字段诊断 & _
                        " From 病人信息 A,病案主页 B,病人新生儿记录 Q,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID+0=[1] And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                        strICUSQL & " And B.封存时间 is NULL"
                End If
                strSQLBaby = "Select q.病人id, q.主页id, q.序号, q.婴儿姓名, q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
                    " From 病案主页 B,病人新生儿记录 Q" & _
                    " Where Nvl(B.主页ID,0)<>0 And B.出院科室ID+0=[1] And b.病人id=q.病人ID  And b.主页ID=q.主页ID " & strTmpOut & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                    strICUSQL & " And B.封存时间 is NULL"
            ElseIf tbcPati.Selected.Tag = "会诊" Then
                '会诊病人:在院
                If InStr(mstrPrivs, "会诊病人") > 0 And Not mblnICU Then
                    strSQL = _
                        "Select 6||Decode(D.执行状态,1,decode(d.完成人,[2],1,2),0) as 排序,Decode(D.执行状态,1,1,0) as 排序2,Null as 组ID,decode(d.完成人,[2],d.完成人,null,null,'其它医生')||Decode(d.执行状态,1,'已完成会诊','未完成会诊') as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,G.名称 as 科室,P.名称 as 病区,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,A.身份证号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID," & _
                        " B.婴儿病区ID" & str字段诊断 & ",D.医嘱ID,D.发送号,Decode(D.执行状态,1,1,0) As 执行状态,D.执行部门ID as 执行科室ID,e.申请序号," & _
                        " G.名称||'病人,'||Decode(D.执行状态,1,'已完成'||E.医嘱内容,'请于'||To_Char(E.开始执行时间,'MM.DD HH24:MI')||'进行'||E.医嘱内容||Decode(E.医生嘱托,NULL,NULL,'('||E.医生嘱托||')')) As 会诊内容," & _
                        " Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,Decode(E.紧急标志,1,'急','普') as 紧急标志" & _
                        " From 病人信息 A,病案主页 B,病人医嘱发送 D,病人医嘱记录 E,诊疗项目目录 F,部门表 G,病人新生儿记录 Q,部门表 P" & _
                        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And E.病人科室ID=G.ID" & _
                        IIf(chkOut.Value = 1, "", " And B.出院日期 is NULL") & " And Nvl(B.状态,0)<>3  And b.病人id=q.病人ID(+) AND B.当前病区ID=P.ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null)" & _
                        " And B.病人ID=E.病人ID And B.主页ID=E.主页ID And D.医嘱ID=E.ID And E.诊疗项目ID=F.ID" & _
                        " And E.诊疗类别='Z' And F.操作类型='7' And E.执行科室id+0=[1] And E.开始执行时间 Between to_date('" & Format(mdtMeetBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And to_date('" & Format(mdtMeetEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And B.封存时间 is NULL" & str会诊状态
                End If
            ElseIf tbcPati.Selected.Tag = "转出" Then
                '转出病人:医生和床号显示本科转出前的，包含已出院的
                strSQL = _
                    "Select /*+ RULE */Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,Null as 组ID,'转出病人' as 类型," & _
                    " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,NULL as 科室,C.经治医师 as 住院医师," & _
                    " LPAD(C.床号," & intBedLen & ",' ') as 床号,A.身份证号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                    " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(Nvl(b.出院日期, Sysdate))-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种," & _
                    " b.婴儿科室ID,B.婴儿病区ID" & str字段诊断 & _
                    " From 病人信息 A,病案主页 B,病人变动记录 C,病人新生儿记录 Q" & _
                    " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null)" & _
                    " And Nvl(B.状态,0)<>2 And B.出院科室ID<>[1] And Nvl(C.附加床位,0)=0" & _
                    " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.科室ID+0=[1]" & _
                    " And C.终止原因 =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And C.经治医师=[2]") & _
                    strICUOutSQL & _
                    " And B.封存时间 is NULL "
                
                strSQLBaby = "Select q.病人id, q.主页id, q.序号, q.婴儿姓名, q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
                    " From 病案主页 B,病人变动记录 C,病人新生儿记录 Q" & _
                    " Where Nvl(B.主页ID,0)<>0 And b.病人id=q.病人ID And b.主页ID=q.主页ID " & _
                    " And Nvl(B.状态,0)<>2 And B.出院科室ID<>[1] And Nvl(C.附加床位,0)=0" & _
                    " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.科室ID+0=[1]" & _
                    " And C.终止原因 =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And C.经治医师=[2]") & _
                    strICUOutSQL & _
                    " And B.封存时间 is NULL "
            End If
        Else
            '按病区查看
            '在院病人
            If tbcPati.Selected.Tag = "在院" Or tbcPati.Selected.Tag = "待入住" Then
                If blnTeam And tbcPati.Selected.Tag = "在院" Then
                    strSQL = _
                        "Select Distinct Decode(B.状态,1,0,3,3,Decode(G.名称,Null,2,1)) as 排序,Decode(Nvl(B.病案状态,0),2,'正在转科',0,999,B.病案状态) as 排序2,G.Id 组ID," & _
                        " Decode(B.状态,1,'待入住病人',3,'预出院病人',Decode(G.名称,Null,'在院病人',G.名称)) as 类型,G.名称 as 小组名," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,C.名称 as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种," & _
                        " b.婴儿科室ID,B.婴儿病区ID" & str字段诊断 & _
                        " From 病人信息 A,病案主页 B,部门表 C,临床医疗小组 G,病人新生儿记录 Q,在院病人 R,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null)" & _
                        " And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID And B.医疗小组ID=G.ID(+)" & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "And (B.医疗小组ID is Null or b.医疗小组id in(Select g.Id From 临床医疗小组 G, 病区科室对应 I Where g.科室id = i.科室id And i.病区id = [1]))", _
                            " And (B.医疗小组ID is Null And B.住院医师=[2] or B.医疗小组ID in(Select 小组id From 医疗小组人员 Where 人员id = [5]))") & _
                        strICUSQL & _
                        " And B.封存时间 is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "待入住", " And B.状态=1", " And B.状态<>1")
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.状态,1,0,3,3,Decode(B.住院医师,[2],1,2)) as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,Null as 组ID," & _
                        " Decode(B.状态,1,'待入住病人',3,'预出院病人',Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的在院病人','在院病人')) as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,C.名称 as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间) ) as 住院天数,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID" & _
                        str字段诊断 & _
                        " From 病人信息 A,病案主页 B,部门表 C,病人新生儿记录 Q,在院病人 R,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null)" & _
                        " And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID " & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                        strICUSQL & _
                        " And B.封存时间 is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "待入住", " And B.状态=1", " And B.状态<>1")
                End If
                strSQLBaby = "Select q.病人id, q.主页id, q.序号, q.婴儿姓名, q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
                    " From 病人信息 A,病案主页 B,部门表 C,病人新生儿记录 Q,在院病人 R" & _
                    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And b.病人id=q.病人ID And b.主页ID=q.主页ID " & _
                    " And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID " & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                    strICUSQL & " And B.封存时间 is NULL" & strFilter & _
                    IIf(tbcPati.Selected.Tag = "待入住", " And B.状态=1", " And B.状态<>1")
            ElseIf tbcPati.Selected.Tag = "出院" Then
                '出院病人:出院病人可能已有多次住院
                If blnOutTeam Then
                    strSQL = _
                        "Select Distinct Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],5.1,5.2),Decode(B.住院医师,[2],4.1,4.2)) as 排序," & _
                        " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,G.Id as 组ID,G.名称 as 小组名," & _
                        " Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的死亡病人','其他死亡病人'),Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的出院病人',Decode(g.名称, Null, '其他出院病人', g.名称))) as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,C.名称 as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(B.出院日期)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID,B.责任护士" & _
                        str字段诊断 & _
                        " From 病人信息 A,病案主页 B,部门表 C,临床医疗小组 G,病人新生儿记录 Q,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null)" & _
                        " And B.当前病区ID+0=[1] And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "And (B.医疗小组ID is Null or b.医疗小组id in(select id from 临床医疗小组 where 科室id =[1]))", _
                            " And (B.医疗小组ID is Null And B.住院医师=[2] or B.医疗小组ID in(Select 小组id From 医疗小组人员 Where 人员id = [5]))") & _
                        strICUSQL & " And B.封存时间 is NULL And b.医疗小组id = g.Id(+)"
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],5.1,5.2),Decode(B.住院医师,[2],4.1,4.2)) as 排序," & _
                        " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,Null as 组ID," & _
                        " Decode(B.出院方式,'死亡',Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的死亡病人','其他死亡病人'),Decode(B.住院医师,[2],'" & UserInfo.姓名 & "的出院病人','其他出院病人')) as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,C.名称 as 科室,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                        " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(B.出院日期)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID,B.责任护士" & _
                        str字段诊断 & _
                        " From 病人信息 A,病案主页 B,部门表 C,病人新生儿记录 Q,病人诊断记录 H" & _
                        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And H.病人id(+)=b.病人id And h.主页id(+)=b.主页id And (q.序号=1 Or q.序号 is Null)" & _
                        " And B.当前病区ID+0=[1] And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                        strICUSQL & " And B.封存时间 is NULL"
                End If
                strSQLBaby = "Select q.病人id, q.主页id, q.序号, q.婴儿姓名, q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
                    " From 病案主页 B,部门表 C,病人新生儿记录 Q" & _
                    " Where Nvl(B.主页ID,0)<>0 And b.病人id=q.病人ID And b.主页ID=q.主页ID " & _
                    " And B.当前病区ID+0=[1] And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & strTmpOut & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                    strICUSQL & " And B.封存时间 is NULL"
            ElseIf tbcPati.Selected.Tag = "会诊" Then
                '会诊病人:在院(不加站点限制)
                If InStr(mstrPrivs, "会诊病人") > 0 And Not mblnICU Then
                    strSQL = _
                        "Select 6||Decode(D.执行状态,1,decode(d.完成人,[2],1,2),0) as 排序,Decode(D.执行状态,1,1,0) as 排序2,Null as 组ID,decode(d.完成人,[2],d.完成人,null,null,'其它医生')||Decode(d.执行状态,1,'已完成会诊','未完成会诊') as 类型," & _
                        " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,G.名称 as 科室,P.名称 as 病区,B.住院医师," & _
                        " LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID ,B.婴儿病区ID" & str字段诊断 & "," & _
                        " D.医嘱ID,D.发送号,Decode(D.执行状态,1,1,0) As 执行状态,D.执行部门ID as 执行科室ID,e.申请序号," & _
                        " G.名称||'病人,'||Decode(D.执行状态,1,'已完成'||E.医嘱内容,'请于'||To_Char(E.开始执行时间,'MM.DD HH24:MI')||'进行'||E.医嘱内容||Decode(E.医生嘱托,NULL,NULL,'('||E.医生嘱托||')')) As 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,Decode(E.紧急标志,1,'急','普') as 紧急标志" & _
                        " From 病人信息 A,病案主页 B,病人医嘱发送 D,病人医嘱记录 E,诊疗项目目录 F,部门表 G,病人新生儿记录 Q,部门表 P" & _
                        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And E.病人科室ID=G.ID  AND B.当前病区ID=P.ID(+) And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null)" & _
                        IIf(chkOut.Value = 1, "", " And B.出院日期 is NULL") & " And Nvl(B.状态,0)<>3 " & _
                        " And B.病人ID=E.病人ID And B.主页ID=E.主页ID And D.医嘱ID=E.ID And E.诊疗项目ID=F.ID" & _
                        " And E.诊疗类别='Z' And F.操作类型='7' And E.开始执行时间 Between to_date('" & Format(mdtMeetBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And to_date('" & Format(mdtMeetEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And E.执行科室id+0 IN (Select 科室ID From 病区科室对应 Where 病区ID=[1])" & _
                        " And B.封存时间 is NULL" & str会诊状态
                End If
            ElseIf tbcPati.Selected.Tag = "转出" Then
                '转出病人:医生和床号显示本科转出前的，包含已出院的(不加站点限制)
                strSQL = _
                    "Select /*+ RULE */Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,Null as 组ID,'转出病人' as 类型," & _
                    " A.病人ID,B.主页ID,A.就诊卡号,B.留观号,A.门诊号,B.住院号,A.身份证号,NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄,D.名称 as 科室,C.经治医师 as 住院医师," & _
                    " LPAD(C.床号," & intBedLen & ",' ') as 床号,B.费别,Decode(B.入科时间,NULL,B.入院日期,B.入科时间) As 入院日期,B.出院日期,B.病人类型,B.状态,B.险类,B.病案状态," & _
                    " -Null as 医嘱ID,-Null as 发送号,-Null as 执行状态,-Null as 执行科室ID,Null as 会诊内容,Nvl(b.路径状态,-1) 路径状态,trunc(Nvl(b.出院日期, Sysdate))-trunc(Decode(B.入科时间,NULL,B.入院日期,B.入科时间)) as 住院天数,q.序号 As 新生儿序号,b.单病种,b.婴儿科室ID,B.婴儿病区ID" & _
                    str字段诊断 & _
                    " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,病人新生儿记录 Q" & _
                    " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID=D.ID And b.病人id=q.病人ID(+) And b.主页ID=q.主页ID(+) And (q.序号=1 Or q.序号 is Null) " & _
                    " And Nvl(B.状态,0)<>2 And B.当前病区ID<>[1] And Nvl(C.附加床位,0)=0" & _
                    " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.病区ID+0=[1]" & _
                    " And C.终止原因 =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And C.经治医师=[2]") & _
                    strICUOutSQL & _
                    " And B.封存时间 is NULL "
                    
                strSQLBaby = "Select q.病人id, q.主页id, q.序号, q.婴儿姓名, q.婴儿性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间)||'天' As 年龄" & _
                    " From 病案主页 B,病人变动记录 C,部门表 D,病人新生儿记录 Q" & _
                    " Where Nvl(B.主页ID,0)<>0 And B.出院科室ID=D.ID And b.病人id=q.病人ID And b.主页ID=q.主页ID" & _
                    " And Nvl(B.状态,0)<>2 And B.当前病区ID<>[1] And Nvl(C.附加床位,0)=0" & _
                    " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.病区ID+0=[1]" & _
                    " And C.终止原因 =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And C.经治医师=[2]") & _
                    strICUOutSQL & _
                    " And B.封存时间 is NULL "
            End If
        End If
        If strSQL = "" Then
            rptPati.Records.DeleteAll
            rptPati.Populate
            mblnUnRefresh = False
            Screen.MousePointer = 0
            LoadPatients = True
            Exit Function
        End If
        If blnTeamGroup Then
            strSQL = strSQL & " Order by 小组名,排序,排序2,床号,主页ID Desc"
        Else
            strSQL = strSQL & " Order by 排序,排序2,床号,主页ID Desc"
        End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.姓名, _
            mintChange, str病况条件, UserInfo.ID, Val(Trim(PatiIdentify.Text)), mstrUserDeps)
        If strSQLBaby <> "" Then
            Set rsBaby = zlDatabase.OpenSQLRecord(strSQLBaby, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.姓名, _
                mintChange, str病况条件, UserInfo.ID, Val(Trim(PatiIdentify.Text)), mstrUserDeps)
        End If
        
        strSQL = ""
        If tbcPati.Selected.Tag = "在院" Or tbcPati.Selected.Tag = "出院" Then
            If mintDeptView = 0 Then
                If tbcPati.Selected.Tag = "在院" Then
                    strTab病人 = "Select b.病人ID,B.主页ID From 病案主页 B,在院病人 R " & _
                        " Where (R.科室ID=[1] Or b.婴儿科室ID=[1]) And b.病人ID=R.病人ID And b.出院科室ID+0=R.科室ID and b.出院日期 is null " & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                        strICUSQL & " And B.封存时间 is NULL" & strFilter & " And B.状态<>1"
                Else
                    strTab病人 = "Select b.病人id, b.主页id" & vbNewLine & _
                                "From 病案主页 B" & vbNewLine & _
                                "Where (b.出院科室ID = [1] Or b.婴儿科室ID = [1]) And b.出院日期 Is Not Null And b.封存时间 Is Null " & strTmpOut
                End If
            Else
                If tbcPati.Selected.Tag = "在院" Then
                    strTab病人 = "Select b.病人ID,B.主页ID From 病案主页 B,在院病人 R " & _
                        " Where (R.病区ID=[1] Or b.婴儿病区ID=[1]) And b.病人ID=R.病人ID And b.出院科室ID+0=R.科室ID and b.出院日期 is null " & _
                        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
                        strICUSQL & " And B.封存时间 is NULL" & strFilter & " And B.状态<>1"
                Else
                    strTab病人 = "Select b.病人id, b.主页id" & vbNewLine & _
                                "From 病案主页 B" & vbNewLine & _
                                "Where (b.当前病区id = [1] Or b.婴儿病区id = [1]) And b.出院日期 Is Not Null And b.封存时间 Is Null " & strTmpOut
                End If
            End If
            strSQL = "select  m.病人id,m.主页id,max(m.记录) as 记录,max(m.填写) as 填写,max(m.状态) as 状态 from " & vbNewLine & _
                "( " & _
                "select a.病人id,a.主页id,1 as 记录,0 as 填写,0 as 状态 from ( " & strTab病人 & ") a " & vbNewLine & _
                "where exists(select 1 from 疾病阳性记录 b where a.病人id=b.病人id and a.主页id=b.主页id) " & vbNewLine & _
                "union all " & vbNewLine & _
                "Select  a.病人id,a.主页id,0 as 记录,1 as 填写,0 as 状态 From ( " & strTab病人 & ") a " & vbNewLine & _
                "Where exists( select 1 from  电子病历记录 C where c.病历种类 = 5 And a.病人id = c.病人id And a.主页id = c.主页id  and c.病历名称 like '%传染病%') " & vbNewLine & _
                "union all " & vbNewLine & _
                "Select  a.病人id,a.主页id,0 as 记录,1 as 填写,e.处理状态 as 状态 From ( " & strTab病人 & ") a ,疾病申报记录 E " & vbNewLine & _
                "Where a.病人ID=E.病人ID and A.主页ID=e.主页ID ) M " & vbNewLine & _
                "group by m.病人id,m.主页id "
        End If
        
        If strSQL <> "" Then
            Set rs传染病状态 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.姓名, _
                mintChange, str病况条件, UserInfo.ID, Val(Trim(PatiIdentify.Text)), mstrUserDeps)
            If rs传染病状态.RecordCount > 0 Then blnDo传染病状态 = True
        End If
    End If
    
    If Not rsPati.EOF Then
        '记录现在选中的病人
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow Then
                If rptPati.SelectedRows(0).Record.Tag <> "" Then
                    lngPatiRow = rptPati.SelectedRows(0).Index '用于快速重新定位
                    strPatiRow = rptPati.SelectedRows(0).Record.Tag
                End If
            End If
        End If
        If mclsWardMonitor.Enabled And InStr(GetInsidePrivs(p住院医生站), "护理监护") > 0 Then
            strMonitor = mclsWardMonitor.GetListPati
        End If
    End If
    
    rptPati.Records.DeleteAll
    
    If tbcPati.Selected.Tag = "出院" Then
        rptPati.Columns(col_责任护士).Visible = True
    Else
        rptPati.Columns(col_责任护士).Visible = False
    End If
    
    '刷新后分组自动展开
    For i = 1 To rsPati.RecordCount
        str类型 = rsPati!类型 & ""
        If blnTeamGroup Then
            str小组名 = rsPati!小组名 & ""
        Else
            str小组名 = ""
        End If
        '小组分组添加父行
        If str类型 <> str小组名 And str小组名 <> "" Then
            blnVisible审查 = True
            str类型 = str小组名
            If str小组名 <> strPre小组名 Then
                Set objPt = Nothing
                strPre小组名 = str小组名
            End If
            If objPt Is Nothing Then
                Set objPt = Me.rptPati.Records.Add()
            ElseIf objPt.Tag <> CStr(rsPati!组ID & "组") Then
                Set objPt = Me.rptPati.Records.Add()
            End If
            If objPt.Tag <> CStr(rsPati!组ID & "组") Then
                objPt.Tag = CStr(rsPati!组ID & "组")
                objPt.Expanded = True
                For j = 0 To rptPati.Columns.Count - 1
                    If j = col_类型 Then
                        If IsNull(rsPati!组ID) Then
                            Set objItem = objPt.AddItem(Val(rsPati!排序))
                        Else
                            Set objItem = objPt.AddItem(-1 * Val(rsPati!组ID))
                        End If
                        objItem.Caption = str类型
                    ElseIf j = col_姓名 Then
                        Set objItem = objPt.AddItem(rsPati!类型 & "")
                        objItem.ForeColor = rptPati.PaintManager.GroupForeColor
                    Else
                        Set objItem = objPt.AddItem("")
                    End If
                    objItem.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
                Next
            End If
        Else
            Set objPt = Nothing
        End If
        
        '根据提交审查情况添加父行
        If NVL(rsPati!病案状态, 0) <> 0 And Val(Mid(rsPati!排序, 1, 1)) <> pt会诊 Then
            blnVisible审查 = True
            '当病人类型发生变化时要另外开一个分支，否则会导致分组不对
            If strPre病人类型 <> rsPati!类型 & "" Then
                Set objParent = Nothing
                strPre病人类型 = rsPati!类型 & ""
            End If
            
            If objParent Is Nothing Then
                If objPt Is Nothing Then
                    Set objParent = Me.rptPati.Records.Add()
                Else
                    Set objParent = objPt.Childs.Add()
                End If
            ElseIf objParent.Tag <> CStr(rsPati!病案状态) Then
                Set objParent = Me.rptPati.Records.Add()
            End If
            If objParent.Tag <> CStr(rsPati!病案状态) Then
                objParent.Tag = CStr(rsPati!病案状态)
                objParent.Expanded = True
                For j = 0 To rptPati.Columns.Count - 1
                    If j = col_类型 Then
                        If IsNull(rsPati!组ID) Then
                            Set objItem = objParent.AddItem(Val(rsPati!排序))
                        Else
                            Set objItem = objParent.AddItem(-1 * Val(rsPati!组ID))
                        End If
                        objItem.Caption = str类型
                    ElseIf j = col_审查 Then
                        Set objItem = objParent.AddItem(Val(rsPati!病案状态))
                        objItem.Caption = " "
                    ElseIf j = col_姓名 Then
                        Set objItem = objParent.AddItem(CStr(Decode(rsPati!病案状态, 1, "等待审查", 2, "拒绝审查", 13, "正在抽查", 3, "正在审查", 14, "抽查反馈", 4, "审查反馈", 16, "抽查整改中", 6, "审查整改中")))
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
        
        If Not objParent Is Nothing Then
            Set objRecord = objParent.Childs.Add()
        ElseIf Not objPt Is Nothing Then
            Set objRecord = objPt.Childs.Add()
        Else
            Set objRecord = Me.rptPati.Records.Add()
        End If
        
        objRecord.Tag = CStr(rsPati!病人ID & "," & rsPati!主页ID) '用于病人定位
        
        If IsNull(rsPati!组ID) Then
            Set objItem = objRecord.AddItem(Val(rsPati!排序)) '分组以Value进行排序
        Else
            Set objItem = objRecord.AddItem(-1 * Val(rsPati!组ID))  '分组以Value进行排序
        End If
        objItem.Caption = str类型

        Set objItem = objRecord.AddItem(Val(Decode(NVL(rsPati!病案状态, 0), 0, 999, rsPati!病案状态)))
        objItem.Caption = " "
        If NVL(rsPati!病案状态, 0) = 2 Then
            objRecord.PreviewText = "  理由:" & GetRefuseReason(rsPati!病人ID, rsPati!主页ID)
        End If

        '图标:注意用在这里是从0开始编号。
        '     图标Value用于存放是否已提交审查，点击才读取
        Set objItem = objRecord.AddItem(-1)
        objItem.Caption = " "
        If NVL(rsPati!病案状态, 0) <> 0 Then
            objItem.Icon = Get病案图标序号(rsPati!病案状态)
        ElseIf "" & rsPati!单病种 <> "" Then
            objItem.Icon = imgPati.ListImages("单病种").Index - 1
        End If

        'lng路径状态=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
        Set objItem = objRecord.AddItem(Val("" & rsPati!路径状态))
        objItem.Caption = " "
        objItem.Icon = -1 + Choose(rsPati!路径状态 + 2, imgPati.ListImages("未导入").Index, imgPati.ListImages("不符合").Index, _
            imgPati.ListImages("执行中").Index, imgPati.ListImages("正常结束").Index, imgPati.ListImages("变异结束").Index)
        
        objRecord.AddItem Val(rsPati!病人ID)
        objRecord.AddItem Val(rsPati!主页ID)
        objRecord.AddItem CStr(NVL(rsPati!姓名))

        If mblnOutDept Then
            Set objItem = objRecord.AddItem("" & rsPati!门诊号)
            objItem.Caption = NVL(rsPati!门诊号, " ")
        Else
            Set objItem = objRecord.AddItem("" & rsPati!住院号)
            objItem.Caption = NVL(rsPati!住院号, " ")
        End If
        
        Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(rsPati!床号), 10)) 'Value用于排序
        objItem.Caption = CStr(Trim(NVL(rsPati!床号, " "))) '为空时会被Value替代
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
        
        objRecord.AddItem CStr(NVL(rsPati!性别))
        objRecord.AddItem CStr(NVL(rsPati!年龄))
        objRecord.AddItem CStr(NVL(rsPati!费别))
        objRecord.AddItem CStr(NVL(rsPati!科室))
        If tbcPati.Selected.Tag = "会诊" Then
            objRecord.AddItem CStr(NVL(rsPati!病区))
        Else
            objRecord.AddItem ""
        End If
        objRecord.AddItem CStr(NVL(rsPati!住院医师))
        objRecord.AddItem Format(rsPati!入院日期, "yyyy-MM-dd HH:mm")
        objRecord.AddItem Format(NVL(rsPati!出院日期), "yyyy-MM-dd HH:mm")
        objRecord.AddItem NVL(rsPati!病人类型)
        
        '用于会诊病人
        objRecord.AddItem Val(NVL(rsPati!医嘱ID, 0))
        objRecord.AddItem Val(NVL(rsPati!发送号, 0))
        objRecord.AddItem Val(NVL(rsPati!执行状态, 0))
        objRecord.AddItem Val(NVL(rsPati!执行科室ID, 0))
        objRecord.AddItem CStr(NVL(rsPati!就诊卡号))
        objRecord.AddItem Val(Trim(IIf(CStr("" & rsPati!住院天数) = "0", "1", CStr("" & rsPati!住院天数))))
        objRecord.AddItem "" & rsPati!单病种
        objRecord.AddItem Val("" & rsPati!婴儿科室ID)
        objRecord.AddItem Val("" & rsPati!婴儿病区ID)
        
        '添加诊断
        objRecord.AddItem CStr(NVL(rsPati!西医诊断))
        objRecord.AddItem CStr(NVL(rsPati!中医诊断))
        
        If tbcPati.Selected.Tag = "会诊" Then
            lng申请序号 = Val("" & rsPati!申请序号)
        Else
            lng申请序号 = 0
        End If
        objRecord.AddItem lng申请序号
        
        '添加传染病状态
        strSQL = ""
        If blnDo传染病状态 Then
            rs传染病状态.Filter = "病人ID=" & Val(rsPati!病人ID) & " and 主页ID=" & Val(rsPati!主页ID)
            If Not rs传染病状态.EOF Then strSQL = Get传染病状态(Val(rs传染病状态!记录 & ""), Val(rs传染病状态!填写 & ""), Val(rs传染病状态!状态 & ""))
        End If
        objRecord.AddItem strSQL
        If tbcPati.Selected.Tag = "出院" Then
            objRecord.AddItem rsPati!责任护士 & ""
        Else
            '填充列数据
            objRecord.AddItem ""
        End If
        
        objRecord.AddItem "" & rsPati!留观号
        objRecord.AddItem "" & rsPati!身份证号
        If tbcPati.Selected.Tag = "会诊" Then
            objRecord.AddItem "" & rsPati!紧急标志
        Else
            objRecord.AddItem ""
        End If
        
        '显示病人颜色
        objRecord.Item(col_姓名).ForeColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
        For j = 0 To rptPati.Columns.Count - 1
            If j <> col_类型 And j <> col_审查 And j <> col_图标 Then
                objRecord.Item(j).ForeColor = objRecord.Item(col_姓名).ForeColor
            End If
        Next
        
        '已完成的会诊病人用灰色显示
        If Val(Mid(rsPati!排序, 1, 1)) = pt会诊 Then
            objRecord.PreviewText = "  " & rsPati!会诊内容
            If NVL(rsPati!执行状态, 0) = 1 Then
                For j = 0 To rptPati.Columns.Count - 1
                    objRecord.Item(j).ForeColor = &H808080
                Next
                objRecord.Item(col_图标).Icon = 2
            Else
                objRecord.Item(col_图标).Icon = 1
            End If
        End If
        '统计病人数目
        lngCount(Val(Mid(rsPati!排序, 1, 1))) = lngCount(Val(Mid(rsPati!排序, 1, 1))) + 1
        
        '根据是否有婴儿添加婴儿行
        If Not rsBaby Is Nothing Then
            Set objBabyParent = objRecord
            rsBaby.Filter = "病人ID=" & objBabyParent(col_病人Id).Value & " and 主页ID=" & objBabyParent(col_主页ID).Value
            If Not rsBaby.EOF Then
                blnVisible审查 = True
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
        
        '收信病人关键信息
        mstrAllPatis = mstrAllPatis & "," & rsPati!病人ID & ":" & rsPati!主页ID
        rsPati.MoveNext
    Next
     
    Call ShowAllPati图标(mstrAllPatis)

    If tbcPati.Selected.Tag = "在院" Then
        rptPati.Columns.Find(col_西医诊断).Visible = True
        rptPati.Columns.Find(col_西医诊断).Caption = "西医诊断"
        rptPati.Columns.Find(col_中医诊断).Visible = Sys.DeptHaveProperty(cboDept.ItemData(cboDept.ListIndex), "中医科")
        rptPati.Columns.Find(COL_传染病).Visible = True
    ElseIf tbcPati.Selected.Tag = "出院" Then
        rptPati.Columns.Find(col_西医诊断).Visible = True
        rptPati.Columns.Find(col_西医诊断).Caption = "出院诊断"
        rptPati.Columns.Find(col_中医诊断).Visible = False
        rptPati.Columns.Find(COL_传染病).Visible = True
    Else
        rptPati.Columns.Find(col_西医诊断).Visible = False
        rptPati.Columns.Find(col_中医诊断).Visible = False
        rptPati.Columns.Find(COL_传染病).Visible = False
    End If
    
    
    rptPati.Columns.Find(col_是否急诊).Visible = IIf(tbcPati.Selected.Tag = "会诊", True, False)
    rptPati.Columns.Find(col_病区).Visible = IIf(tbcPati.Selected.Tag = "会诊", True, False)
    
    
    rptPati.Columns(col_审查).Visible = blnVisible审查
    If tbcPati.Selected.Tag = "会诊" Then
        rptPati.Columns.Find(col_科室).Visible = True
    Else
        rptPati.Columns.Find(col_科室).Visible = mintDeptView = 1
    End If
    If mblnOutDept Then
        rptPati.Columns.Find(col_住院号).Caption = "门诊号"
    Else
        rptPati.Columns.Find(col_住院号).Caption = "住院号"
    End If
    rptPati.Populate
    '根据江陵医院需求加入病人数目统计信息
    strState = " 共 " & rsPati.RecordCount & " 个病人"
    For i = LBound(lngCount) To UBound(lngCount)
        If lngCount(i) > 0 Then
            Select Case i
            Case 0
                strState = strState & "，待入住:"
            Case 1
                If blnTeam Then
                    strState = strState & "，医疗小组:"
                Else
                    strState = strState & IIf(tbcPati.Selected.Tag = "在院", "，" & UserInfo.姓名 & "的在院:", "，" & UserInfo.姓名 & "的出院:")
                End If
            Case 2
                strState = strState & "，本科在院:"
            Case 3
                strState = strState & "，本科预出院:"
            Case 4
                strState = strState & "，本科出院:"
            Case 5
                strState = strState & "，本科死亡:"
            Case 6
                strState = strState & "，会诊:"
            Case 7
                strState = strState & "，转出:"
            End Select
            strState = strState & lngCount(i) & "人"
        End If
    Next
    stbThis.Panels(2).Text = strState
    stbThis.Panels(2).Tag = strState
    lblFee(1).Caption = ""
    
    '定位病人行:在Populate之后
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Or rsPati.RecordCount > 1 And lngPatiRow = 0 Then
        Call ClearPatiInfo
        '按无数据刷新子窗体
        Call SubWinRefreshData(tbcSub.Selected)
        
        mstrPreNotify = ""
        rptNotify.Records.DeleteAll
        rptNotify.Populate
        rptNotify.TabStop = False
        
        strState = " 没有病人"
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
    
    Dim i As Long
    
    mlng病人ID = 0
    mlng主页ID = 0
    
    mPatiInfo.状态 = 0
    mPatiInfo.住院号 = ""
    mPatiInfo.床号 = ""
    mPatiInfo.婴儿 = -1
    mPatiInfo.主页ID = 0
    mPatiInfo.病区ID = 0
    mPatiInfo.科室ID = 0
    mPatiInfo.入院日期 = CDate(0)
    mPatiInfo.出院日期 = CDate(0)
    mPatiInfo.编目日期 = CDate(0)
    mPatiInfo.住院次数 = 0
    mPatiInfo.数据转出 = False
    Set mPatiInfo.rs图标 = Nothing
        
    cboPages.Clear
    cbo过敏.Clear
    lbl类型(1).Caption = ""
    lblPatiName(1).Caption = ""
    lblPatiName(1).ToolTipText = ""
    lbl姓名(1).Caption = ""
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
                ElseIf mstrFindType = "留观号" Then
                    If .Record(col_留观号).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "床号" Then '床号
                    If UCase(Trim(.Record(col_床号).Value)) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "就诊卡" Then '就诊卡
                    If UCase(.Record(col_就诊卡).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "姓名" Then '姓名
                    If .Record(col_姓名).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                ElseIf mstrFindType = "病人诊断" Then '病人诊断
                    If tbcPati.Selected.Tag = "在院" Or tbcPati.Selected.Tag = "出院" Then
                        If .Record(col_西医诊断).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                    End If
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
            MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function ExecuteMeetIdea(ByVal intType As Integer) As Boolean
'功能：查看或填写会诊意见
'参数：intType 0-填写，1－查看
    Dim blnOk As Boolean
    Dim objFrm As Object
    With rptPati.SelectedRows(0)
        blnOk = mobjKernel.ShowConsultationApply(Me, Val(.Record(col_医嘱ID).Value), IIf(intType = 0, 3, 4), objFrm)
        Set mFrmConsultation = objFrm
    End With
    If blnOk Then
        Call LoadPatients '界面数据刷新
        If rptPati.Visible Then rptPati.SetFocus
    End If
End Function

Private Function ExecuteMeetFinish() As Boolean
'功能：完成对当前病的人会诊
    Dim strSQL As String
    Dim strTmp As String
    Dim str会诊意见 As String
    
    Dim lng部门ID As Long
    Dim lng医嘱ID As Long
    Dim lng附项序号 As Long
    Dim lngTmp As Long
    Dim i As Long
    Dim lngType As Long '处理方式 1－内部处理，2－弹出窗体处理
    Dim lng判断 As Long '检查判断 1－先查会诊意见，再处理，2－直接处理
    
    Dim blnOk As Boolean
    Dim bln意见 As Boolean
    Dim blnTrans As Boolean
    Dim blnDo As Boolean
    
    Dim rsTmp As ADODB.Recordset
    Dim rs要素 As ADODB.Recordset
    Dim arrSQL As Variant
    
    If mlng病人ID = 0 Then Exit Function
    
    lng医嘱ID = Val(rptPati.SelectedRows(0).Record(col_医嘱ID).Value)
    
    
    
    If Val(rptPati.SelectedRows(0).Record(COL_申请序号).Value) <> 0 Then
        '判断是否已经写了会诊意见查医嘱附项
        strSQL = "select a.内容 from 病人医嘱附件 a where a.医嘱ID=[1] and a.项目='会诊意见'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
        If Not rsTmp.EOF Then
            str会诊意见 = rsTmp!内容 & ""
            str会诊意见 = Trim(str会诊意见)
        End If
    
        '取参数多科会诊
        If Val(zlDatabase.GetPara(237, glngSys)) = 1 Then
            If Is代表科室(lng医嘱ID) Then
                lng判断 = 1
            Else
                lng判断 = 2
            End If
        Else
            lng判断 = 1
        End If
        
        If str会诊意见 = "" And lng判断 = 1 Then
            If MsgBox("未填写会诊意见，是否填写会诊意见？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                lngType = 1
            Else
                lngType = 2
            End If
        Else
            lngType = 1
        End If
    
        
        If lngType = 2 Then
            blnOk = mobjKernel.ShowConsultationApply(Me, lng医嘱ID, 3)
            If blnOk Then
                '判断内部是否点了完成
                strSQL = "select 1 from 病人医嘱发送 a where a.医嘱ID=[1] and a.执行状态 = 1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                
                If rsTmp.EOF Then
                
                    If cboDept.ListIndex <> -1 Then lng部门ID = cboDept.ItemData(cboDept.ListIndex)
                    
                    With rptPati.SelectedRows(0)
                        strSQL = "ZL_病人医嘱执行_Finish(" & .Record(col_医嘱ID).Value & "," & .Record(col_发送号).Value & ",NULL,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lng部门ID & ")"
                    End With
                    On Error GoTo errH
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    On Error GoTo 0
                End If
            End If
        End If
        
        If lngType = 1 Then
            arrSQL = Array()
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Delete(" & lng医嘱ID & ",'会诊意见')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Delete(" & lng医嘱ID & ",'会诊完成时间')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Delete(" & lng医嘱ID & ",'会诊完成科室')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Delete(" & lng医嘱ID & ",'会诊医生')"
            
            strSQL = "select max(a.排列) as 序号 from 病人医嘱附件 a where a.医嘱ID=[1] and nvl(a.项目,'空') not in ('会诊意见','会诊完成时间','会诊完成科室','会诊医生')"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then
                lng附项序号 = Val(rsTmp!序号 & "") + 1
            Else
                lng附项序号 = 1
            End If
            
            '1.插入附项
            strSQL = "select a.id,a.中文名 from 诊治所见项目 a where nvl(a.中文名,'空') in ('会诊意见','会诊完成时间','会诊完成科室','会诊医生')"
            Set rs要素 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            For i = 1 To rs要素.RecordCount
                If rs要素!中文名 & "" = "会诊意见" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊意见',0," & lng附项序号 & "," & rs要素!ID & ",'" & str会诊意见 & "')"
                    lng附项序号 = lng附项序号 + 1
                ElseIf rs要素!中文名 & "" = "会诊完成时间" Then
                    strTmp = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊完成时间',0," & lng附项序号 & "," & rs要素!ID & ",'" & strTmp & "')"
                    lng附项序号 = lng附项序号 + 1
                ElseIf rs要素!中文名 & "" = "会诊完成科室" Then
                    lngTmp = Val(rptPati.SelectedRows(0).Record(col_执行科室ID).Value)
                    strSQL = "select 名称 from 部门表 where id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngTmp)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊完成科室',0," & lng附项序号 & "," & rs要素!ID & ",'" & rsTmp!名称 & "')"
                    lng附项序号 = lng附项序号 + 1
                ElseIf rs要素!中文名 & "" = "会诊医生" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊医生',0," & lng附项序号 & "," & rs要素!ID & ",'" & UserInfo.姓名 & "')"
                    lng附项序号 = lng附项序号 + 1
                End If
                rs要素.MoveNext
            Next
            
            If cboDept.ListIndex <> -1 Then lng部门ID = cboDept.ItemData(cboDept.ListIndex)
            With rptPati.SelectedRows(0)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱执行_Finish(" & .Record(col_医嘱ID).Value & "," & .Record(col_发送号).Value & ",NULL,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lng部门ID & ")"
            End With
            gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans: blnTrans = False
            blnOk = True
        End If
    Else
        If MsgBox("确实要完成对该""" & lbl姓名(1).Caption & """的会诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        If cboDept.ListIndex <> -1 Then lng部门ID = cboDept.ItemData(cboDept.ListIndex)
        
        With rptPati.SelectedRows(0)
            strSQL = "ZL_病人医嘱执行_Finish(" & .Record(col_医嘱ID).Value & "," & .Record(col_发送号).Value & ",NULL,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lng部门ID & ")"
        End With
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        blnOk = True
    End If
    
    If blnOk Then
        Call LoadPatients '界面数据刷新
        If rptPati.Visible Then rptPati.SetFocus
    End If
    
    ExecuteMeetFinish = blnOk
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteMeetCancel() As Boolean
'功能：取消完成对当前病的人会诊
    Dim strSQL As String, lng部门ID As Long
    
    If mlng病人ID = 0 Then Exit Function
    If MsgBox("确实要取消完成对该""" & lbl姓名(1).Caption & """的会诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If cboDept.ListIndex <> -1 Then lng部门ID = cboDept.ItemData(cboDept.ListIndex)
    
    With rptPati.SelectedRows(0)
        strSQL = "ZL_病人医嘱执行_Cancel(" & .Record(col_医嘱ID).Value & "," & .Record(col_发送号).Value & ",Null,0," & lng部门ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    End With
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '界面数据刷新
    If rptPati.Visible Then rptPati.SetFocus
    ExecuteMeetCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Execute接受会诊(blnCancel As Boolean) As Boolean
'功能：取消完成对当前病的人会诊
    Dim strSQL As String, lng部门ID As Long
    
    If mlng病人ID = 0 Then Exit Function
    If MsgBox("确认要" & IIf(blnCancel, "取消", "") & "接受对该""" & lblPatiName(1).Caption & """的会诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If cboDept.ListIndex <> -1 Then lng部门ID = cboDept.ItemData(cboDept.ListIndex)
    
    With rptPati.SelectedRows(0)
        strSQL = "Zl_病人医嘱发送_会诊处理(" & .Record(col_医嘱ID).Value & "," & .Record(col_发送号).Value & "," & IIf(blnCancel = True, "1", "0") & ",'" & UserInfo.姓名 & "')"
    End With
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '界面数据刷新
    If rptPati.Visible Then rptPati.SetFocus
    Execute接受会诊 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ExecuteMedRecAuditSubmit() As Boolean
'功能：提交病人病案审查
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long, i As Long, lng路径状态 As Long
    Dim strMsg As String
    
    If mlng病人ID = 0 Then Exit Function
    On Error GoTo errH
    With rptPati.SelectedRows(0)
        lng病人ID = .Record(col_病人Id).Value
        lng主页ID = .Record(col_主页ID).Value
    End With

    lng路径状态 = rptPati.SelectedRows(0).Record(col_路径状态).Value
    If lng路径状态 = 1 Then
        strMsg = "该病人还存在未完成的临床路径，不能提交病案！"
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Select 病历名称 From 电子病历记录 Where 病人id = [1] And 主页id = [2] And 打印时间 Is Null Order By 创建时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病历打印", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If strMsg = "" Then
                strMsg = rsTmp!病历名称
            Else
                strMsg = strMsg & IIf((i Mod 2) = 0, "," & vbTab, vbCrLf) & rsTmp!病历名称
            End If
            If Len(strMsg) > 1000 Then
                strMsg = strMsg & "......"
                Exit For
            End If
            rsTmp.MoveNext
        Next
        strMsg = "以下病历未打印：" & vbCrLf & strMsg & vbCrLf & "你确定要继续吗？"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    
    strSQL = "Zl_病案提交记录_Insert(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '界面数据刷新
    If rptPati.Visible Then rptPati.SetFocus
    ExecuteMedRecAuditSubmit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteMedRecAuditCancel() As Boolean
'功能：取消提交病人病案审查
    Dim strSQL As String
    
    If mlng病人ID = 0 Then Exit Function
    If MsgBox("确实要将""" & lbl姓名(1).Caption & """的病案取消提交审查吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    With rptPati.SelectedRows(0)
        strSQL = "Zl_病案提交记录_Delete(" & .Record(col_病人Id).Value & "," & .Record(col_主页ID).Value & ")"
    End With
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '界面数据刷新
    If rptPati.Visible Then rptPati.SetFocus
    ExecuteMedRecAuditCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
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
    
    If rptPati.SelectedRows(0).GroupRow = False Then
        If rptPati.SelectedRows(0).Record(COL_婴儿科室ID).Value <> 0 Then
            If rptPati.SelectedRows(0).Record(COL_婴儿科室ID).Value = cboDept.ItemData(cboDept.ListIndex) Or rptPati.SelectedRows(0).Record(COL_婴儿病区ID).Value = cboDept.ItemData(cboDept.ListIndex) Then
                MsgBox "该病人已经转出本科室了，只有婴儿留在本科室，不允许调整首页。", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    
    '病案编目之后不可以整理
    If Not (CheckMecRed(mlng病人ID, mlng主页ID, Me.Caption) Or blnEditable) Then
        blnReadOnly = True
    End If
    
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p住院医生站, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    '非模态显示首页整理
    If Not mclsInOutMedRec.IsOpen Then
        If mclsInOutMedRec.ShowInMedRecEdit(Me, mlng病人ID, mPatiInfo.主页ID, mPatiInfo.科室ID, rptPati.SelectedRows(0).Record(col_路径状态).Value, , mstrPrivs, IIf(blnReadOnly, 1, 0), False) Then
            mstrPrePati = ""
        End If
    End If
End Sub

Private Sub ExecuteCritical()
'功能：危急值相关处理
    Dim lng危急值ID As Long '本次处理的危急值记录ID
    
    Call mobjKernel.ShowDealCritical(Me, mlng病人ID, mlng主页ID, "", lng危急值ID)
    
    Call SetCriticalAdvice(lng危急值ID)
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Dim curTime As Date
    
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
    
    curTime = Now

    '刷新病历审查提醒
    If mintNotify > 0 And rptNotify.Visible Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call LoadNotify
            If mbln危急值弹窗 Then Call ReadMsgAuto
        End If
    End If
End Sub

Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mintChange = Val(txtChange.Text)
    Call LoadPatients
End Sub

Private Function LoadResponse() As Boolean
'功能：读取病案审查反馈
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngCount As Long
    Dim curDate As Date
    
    If cboDept.ListIndex = -1 Then
        fra审查.Visible = False: LoadResponse = True: Exit Function
    End If
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    Screen.MousePointer = 11
    
    '读取当前科室/病区的在院、出院病人，以"病案反馈记录"为准全部扫描
    strSQL = "Select Count(*) as 数量 From 病案主页 B,病案反馈记录 A" & _
        " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And A.记录状态=1 And A.反馈对象 IN(1,2,5,6,7,8,9)" & _
        IIf(mintDeptView = 0, " And B.出院科室ID + 0=[1]", " And B.当前病区ID + 0=[1]") & _
        IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]") & _
        IIf(mblnICU And InStr(mstrPrivs, "全院病人") = 0, " And B.住院医师=[2]", "") & _
        " And a.反馈时间 Between [3] And [4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadResponse", cboDept.ItemData(cboDept.ListIndex), UserInfo.姓名, CDate(Format(curDate - mlngMedRedDay, "yyyy-MM-dd")), CDate(Format(curDate, "yyyy-MM-dd HH:mm:ss")))
    If Not rsTmp.EOF Then lngCount = NVL(rsTmp!数量, 0)
    
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
    Dim rsTmp As ADODB.Recordset, rsOld As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String, i As Long, j As Long
    Dim strTmp As String, strMsgType As String
    Dim blnDo As Boolean
    Dim strTag As String
    
    On Error GoTo errH
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    
    If cboDept.ListIndex = -1 Then LoadNotify = True: Exit Function
    
    If Mid(mstrNotify, m病历审阅, 1) = "1" Then strTmp = strTmp & ",ZLHIS_EMR_021"
    If Mid(mstrNotify, m医嘱安排, 1) = "1" Then strTmp = strTmp & ",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,ZLHIS_CIS_020"
    If Mid(mstrNotify, m危机值, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotify, m报告撤消, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_002,ZLHIS_PACS_003"
    If Mid(mstrNotify, m医嘱审核, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_026,ZLHIS_CIS_027,ZLHIS_CIS_028,ZLHIS_CIS_029,ZLHIS_CIS_030"
    If Mid(mstrNotify, m处方审查, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_002"
    If Mid(mstrNotify, m传染病, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotify, m病历质控, 1) = "1" Then strTmp = strTmp & ",ZLHIS_EMR_025"
    If Mid(mstrNotify, m输血完成, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_001"  '启用血库才有此消息和参数
    If Mid(mstrNotify, m校对疑问, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_035"
    If Mid(mstrNotify, m用血审核, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_004"  '启用血库才有此消息和参数
    If Mid(mstrNotify, m输血反应, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_006"
    
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
    
    strSQL = "Select b.id,b.病人id, b.就诊id as 主页id,a.住院号,a.姓名,a.当前床号 As 床号, Nvl(b.就诊科室id, a.当前科室id) As 就诊科室id," & _
        " Nvl(b.就诊病区id, a.当前病区id) As 就诊病区id, b.病人来源, b.消息内容, b.类型编码, b.业务标识, b.优先程度, b.登记时间,a.险类" & _
        " From 病人信息 A, 业务消息清单 B, 业务消息提醒部门 C, 业务消息提醒人员 D,病案主页 E" & _
        " Where a.病人id = b.病人id And b.Id = c.消息id And b.Id = d.消息id(+) And b.病人id=e.病人id and b.就诊id=e.主页id and e.主页id is not null And b.登记时间 >=Trunc(Sysdate-" & (mintNotifyDay - 1) & ") and substr(b.提醒场合,[4],1)='1'" & _
        " And Nvl(b.是否已阅, 0) = 0  And instr(','||[5]||',',','||b.类型编码||',')>0 and (c.部门id = [1] Or d.提醒人员 = [3])" & _
        " Order By b.优先程度, b.登记时间 Desc"
        
    If strSQL = "" Then Exit Function
    Screen.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, cboDept.ItemData(cboDept.ListIndex), , UserInfo.姓名, 2, strTmp)

    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!类型编码
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033", "ZLHIS_EMR_025"
            strTag = strTag & "<TB>" & rsTmp!类型编码 & "," & rsTmp!ID
            blnDo = True
        Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!业务标识 & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!业务标识
                blnDo = True
            End If
        Case "ZLHIS_BLOOD_006"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!类型编码 & ":" & rsTmp!病人ID & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!类型编码 & ":" & rsTmp!病人ID
                blnDo = True
            End If
        Case Else
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码
                blnDo = True
            End If
        End Select
        
        If blnDo Then
            Call AddReportRow(rsTmp!病人ID & "," & rsTmp!主页ID, rsTmp!病人ID, rsTmp!主页ID, NVL(rsTmp!姓名), NVL(rsTmp!住院号), NVL(rsTmp!床号), NVL(rsTmp!消息内容), _
                rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!业务标识 & "", rsTmp!病人来源 & "", rsTmp!ID)
                        blnDo = False
        End If
        rsTmp.MoveNext
    Next
    
    '老版病历提醒添加
    If Mid(mstrNotify, m病历审阅, 1) = "1" Then
        strSQL = " Select A.病人ID,A.主页ID,A.病历种类,A.病历名称,A.签名级别,A.完成时间,B.保留" & _
                " From 电子病历记录 A,病历文件列表 B" & _
                " Where A.病人来源 = 2 And A.病历种类 In (2,5,6) And Nvl(A.处理状态,0)<=0 And A.归档人 Is Null" & _
                " And A.文件ID=B.ID(+) And A.完成时间>=Trunc(Sysdate-[2])" & _
                IIf(mintDeptView = 0, " And A.科室ID=[1]", " And A.科室ID IN(Select 科室ID From 病区科室对应 Where 病区ID=[1])")
        strSQL = "Select A.病人ID,A.主页ID,B.住院号,C.床号,NVL(D.姓名,B.姓名) 姓名,Min(A.完成时间) as 时间, -1 As 医嘱状态,'' as 状态" & _
            " From (" & strSQL & ") A,病人信息 B,病人变动记录 C ,病案主页 D" & _
            " Where A.病人ID=B.病人ID And (A.病历种类<>2 Or Nvl(A.保留,0)>=0)" & _
            " And A.病人ID=C.病人ID And A.主页ID=C.主页ID And A.病人ID=D.病人ID And A.主页ID=D.主页ID" & _
            " And C.开始时间 Is Not Null And Nvl(C.附加床位,0)=0 And (C.终止时间 Is Null Or C.终止原因=1)" & _
            " And A.签名级别<Decode([3],C.主任医师,4,C.主治医师,2,C.经治医师,1,0)" & _
            " Group by A.病人ID,A.主页ID,B.住院号,C.床号,NVL(D.姓名,B.姓名) ,NVL(D.性别,B.性别),NVL(D.年龄, B.年龄)  Order by 时间"
        Set rsOld = zlDatabase.OpenSQLRecord(strSQL, Me.Name, cboDept.ItemData(cboDept.ListIndex), mintNotifyDay - 1, UserInfo.姓名)
    
        For i = 1 To rsOld.RecordCount
            Call AddReportRow(rsOld!病人ID & "," & rsOld!主页ID, rsOld!病人ID, rsOld!主页ID, NVL(rsOld!姓名), NVL(rsOld!住院号), NVL(rsOld!床号), "有需要审核的病历。", _
               "", 1, Format(rsOld!时间 & "", "yyyy-MM-dd HH:mm:ss"), "", "", 0)
            rsOld.MoveNext
        Next
    End If
    
    rptNotify.Populate '缺省不选中任何行
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln消息语音 Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(1)
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
        Call SetPatiIconScale
        Call SetpicPatiPosition
        Call SetPatiInfoCtlPos
    End If
    Select Case tbcSub.Selected.Tag
        Case "住院一览"
            Call mfrmInView.SetFontSize(mbytSize)
        Case "路径"
            Call mclsPath.SetFontSize(mbytSize)
        Case "医嘱"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "病历"
            Call mclsEPRs.SetFontSize(mbytSize)
        Case "护理"
            Call mclsTends.SetFontSize(mbytSize)
        Case "新版护理"
            Call mclsTendsNew.SetFontSize(mbytSize)
        Case "护理病历"
            Call mclsTendEPRs.SetFontSize(mbytSize)
        Case "疾病报告"
            Call mclsDisease.SetFontSize(mbytSize)
        Case "新病历"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
     
End Sub

Private Sub SetpicPatiPosition()
'功能：病人列表和过滤条件的相关控件的位置与大小调整
    Dim i As Long
    Dim lngDistance As Long
        
    'checkBox选择框在字体增大时不会改变宽度，因此需要减去100
    lngDistance = IIf(mbytSize = 0, 10, -50)
    Call zlControl.SetPubCtrlPos(False, 0, lbl病况条件, 10, chk病况条件(0), lngDistance, chk病况条件(1), lngDistance, chk病况条件(2), lngDistance, chkByTeam)
    Call zlControl.SetPubCtrlPos(False, 0, lbl出院时间, 50, cboSelectTime(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl开始时间, 50, cboSelectTime(2), 50, chkOut, 50, chkHZ(0), 50, chkHZ(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl转出, 50, cmdRef)
        
    For i = 0 To picPara.Count - 1
        If i = 0 Then picPara(i).Height = IIf(mbytSize = 0, 320, 280)
        If i = 3 Then picPara(i).Height = IIf(mbytSize = 0, 320, 420)
    Next
    
    txtChange.Left = lbl转出.Left + Me.TextWidth("显示最近 ")
    fraChange.Left = txtChange.Left
    fraChange.Top = txtChange.Top + txtChange.Height
    chkOutByTeam.Top = chkByTeam.Top
    chkOutByTeam.Left = chkByTeam.Left
    chkFilter.Height = PatiIdentify.Height
End Sub

Private Sub SetPatiInfoCtlPos()
'功能：病人的详细信息界面的控件位置调整－－－－picInfo中的控件
    Dim lngDistance1 As Long, lngDistance2 As Long
    Dim lngTmp As Long
    
    lngDistance2 = 180: lngDistance1 = 10
    lngTmp = IIf(mbytSize = 0, 1080, 1300)
    
    lblPatiName(0).Top = IIf(mbytSize = 0, 190, 210)
    lbl姓名(0).Top = lblPatiName(0).Top
    cboPages.Width = 1600 ' IIf(mbytSize = 0, 1500, 1600)
    
    '1.住院次数
    lblPatiName(0).Left = IIf(mbytSize = 0, 90, 110)
    lblPages.Left = lblPatiName(0).Left
    cboPages.Left = lblPages.Left + lblPages.Width
    
    Call zlControl.SetPubCtrlPos(False, 0, lblPatiName(0), lngDistance1, lblPatiName(1))
        
    lblPages.Top = lblPatiName(0).Top + lblPatiName(0).Height + 90
    Call zlControl.SetPubCtrlPos(False, 0, lblPages, lngDistance1, cboPages)

    fraPageId.Width = cboPages.Left + cboPages.Width + 60
    fraPageId.Height = lngTmp
    fraInfo.Height = lngTmp
    picInfo.Height = lngTmp
    
    lblCurPati.Left = lblPages.Left
    lblCurPati.Top = lblPages.Top + lblPages.Height + 90
    
    Call zlControl.SetPubCtrlPos(False, 0, lblCurPati, lngDistance1, imgCurPati(0), 80, imgCurPati(1), 80, imgCurPati(2))
    
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

Private Function GetDataToDepts(Optional ByVal strIn As String = "") As ADODB.Recordset
'功能：获取科室病区列表数据记录集
'参数：strIn 过滤条件
    Dim strSQL As String
    Dim blnYN As Boolean
    Dim strDeptIDs As String
    
    If strIn <> "" Then blnYN = True
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

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'功能：将接收到的消息加入提醒列表中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnTmp As Boolean
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    If Mid(rsMsg!提醒场合, 2, 1) <> "1" Then Exit Sub
    
    strSQL = "select 部门id as id,工作性质 as 性质 from 部门性质说明 where 部门id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(rsMsg!部门IDs & ""))
    rsTmp.Filter = "性质=" & IIf(mintDeptView = 0, "'临床'", "'护理'")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!ID & "") = cboDept.ItemData(cboDept.ListIndex) Then
                blnTmp = True: Exit For
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If blnTmp Or InStr("," & rsMsg!提醒人员 & ",", "," & UserInfo.姓名 & ",") > 0 Then
        
        '判断列表是否已经有这类消息了
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_消息).Value = rsMsg!类型编码 And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!病人ID & "," & rsMsg!就诊id) Then
                    Exit Sub
                End If
            End If
        Next
        
        strSQL = "Select a.住院号, a.姓名, a.性别, a.年龄, a.当前床号 As 床号, a.险类 From 病人信息 A Where a.病人id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!病人ID))
        
        Call AddReportRow(rsMsg!病人ID & "," & rsMsg!就诊id, rsMsg!病人ID, rsMsg!就诊id, rsTmp!姓名, NVL(rsTmp!住院号), NVL(rsTmp!床号), NVL(rsMsg!消息内容), _
             rsMsg!类型编码 & "", rsMsg!优先程度 & "", Format(rsMsg!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!业务标识 & "", rsMsg!病人来源 & "", 0)
        
        rptNotify.Populate
         
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ParamArray arrInput() As Variant)
'功能：向消息提配列表中增加一行
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strNO As String
    Dim str业务 As String
    Dim str病人来源 As String
    Dim int优先级 As Integer
    Dim Index As Integer
    
    On Error GoTo errH
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tag值
    Set objItem = objRecord.AddItem(""): objItem.Icon = 3
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '病人id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '就诊id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '姓名
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))      '住院号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))      '床号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))      '状态，内容
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                                     '消息编号
    objRecord.AddItem strNO: Index = Index + 1
    
    int优先级 = Val(arrInput(Index))                            '序号
    objRecord.AddItem int优先级: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '日期
    
    str业务 = arrInput(Index): Index = Index + 1              '业务标识
    str病人来源 = arrInput(Index)                             '病人来源
    objRecord.AddItem str业务
    Index = Index + 1
    objRecord.AddItem Val(arrInput(Index)) '消息ID：业务消息清单.ID
    
    If int优先级 > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int优先级 = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadMsg(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strNO As String, ByVal str业务 As String, ByVal lng消息ID As Long) As Boolean
'功能：阅读消息
'说明：消息阅读方式目前有3种：按消息编译码阅读，消息ID阅读，按业务标识阅读
    Dim strSQL As String
    Dim str医嘱ID As String
    Dim blnDo As Boolean
    Dim lng危急值ID As Long  '本次处理的危急值记录ID
    Dim blnHis危急值 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQLReadMsg As String
    Dim objControl As Object
    Dim i As Long
    On Error GoTo errH
    blnDo = True
    
    strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "',2,'" & UserInfo.姓名 & "'," & cboDept.ItemData(cboDept.ListIndex)
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
        strSQL = strSQL & ",null,null,'" & str业务 & "'"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSQL = strSQL & ",null," & lng消息ID
    End Select
    strSQL = strSQL & ")"
    
    strSQLReadMsg = strSQL
     
    If strNO = "ZLHIS_CIS_035" Or strNO = "ZLHIS_BLOOD_004" Then
        If strNO = "ZLHIS_CIS_035" Then
            '校对疑问消息的处理在医嘱编辑界面，医嘱处理后自动处理消息
            strSQL = "select a.ID,a.相关ID,a.诊疗类别 from 病人医嘱记录 a where A.医嘱状态=2 and a.病人id=[1] and a.主页id=[2] order by a.序号"
        Else
            strSQL = "select 1 from 病人医嘱记录 a where a.病人id=[1] and a.主页id=[2] and a.医嘱状态=1 and a.诊疗类别='K' and a.检查方法='1' and a.审核状态=1 and rownum<2"
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.EOF Then '无数据则将消息设置为已阅
             Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
             ReadMsg = True
             Exit Function
        End If
        
        If strNO = "ZLHIS_CIS_035" Then
            '定位一个有效医嘱行
            For i = 1 To rsTmp.RecordCount
                If InStr(",5,6,", rsTmp!诊疗类别 & "") > 0 Then
                    Call mclsAdvices.LocatedAdviceRow(Val(rsTmp!ID & ""))
                ElseIf "7" = rsTmp!诊疗类别 & "" Then
                    Call mclsAdvices.LocatedAdviceRow(Val(rsTmp!相关ID & ""))
                Else
                    Call mclsAdvices.LocatedAdviceRow(Val(rsTmp!ID & ""))
                End If
                Exit For
                rsTmp.MoveNext
            Next
        End If
  
        If tbcSub.Tag <> "医嘱" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "医嘱" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                        Exit For
                    End If
                End If
            Next
        End If
        
        '医嘱行定位失败修改菜单可能不可用
        Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then
                objControl.Execute
            End If
        End If
        Exit Function
    End If
    
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mbln危急值 Then
            '危急值消息相关处理
            Call mobjKernel.ShowDealCritical(Me, lng病人ID, lng主页ID, "", lng危急值ID)
            
            If lng危急值ID <> 0 Then
                strSQL = "select a.标本id,a.处理情况,a.确认人 from 病人危急值记录 a where a.id=[1] and a.确认人 is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng危急值ID)
                If Not rsTmp.EOF Then
                    '将消息设置为已阅
                    Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
                    '如果是LIS危急值调用LIS接口
                    If strNO = "ZLHIS_LIS_003" Then
                        Call InitObjLis(p住院医生站)
                        If Not gobjLIS Is Nothing Then
                            Call gobjLIS.WriteNotifyToLis(Val(rsTmp!标本ID & ""), rsTmp!确认人 & "", rsTmp!处理情况 & "")
                        End If
                    End If
                End If
            End If
            Call SetCriticalAdvice(lng危急值ID)
            blnHis危急值 = True
        End If
    End If
    
    If Not blnHis危急值 Then
        If strNO = "ZLHIS_LIS_003" Then
            If str业务 <> "" Then
                str医嘱ID = str业务
                Call InitObjLis(p住院医生站)
                If Not gobjLIS Is Nothing Then
                    blnDo = gobjLIS.GetReadNotify(Me, str医嘱ID, UserInfo.姓名)
                End If
            End If
        End If
        If blnDo Then
            Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
        End If
    End If
    ReadMsg = blnDo
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Is代表科室(ByVal lng医嘱ID As Long) As Boolean
'功能：书写会诊议见时判断当前的科室是不是代表科室
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "select 1 from 病人医嘱记录 a,部门表 b where a.执行科室id=b.id and a.id=[1] and" & vbNewLine & _
        "exists (select 1 from 病人医嘱附件 c where c.医嘱id =[1] and c.项目='会诊代表科室'and b.名称=c.内容)"
        
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
    If Not rsTmp.EOF Then
        Is代表科室 = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lblDept_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim vPoint As POINTAPI
    
    Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_View_Dept, , True)
    If Not objPopup Is Nothing Then
        vPoint.X = lblDept.Left / Screen.TwipsPerPixelX
        vPoint.Y = (lblDept.Top + lblDept.Height + 30) / Screen.TwipsPerPixelY
        ClientToScreen picPati.hwnd, vPoint
        objPopup.CommandBar.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub ReadMsg批量(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strNO As String)
'功能：消息处理表格中的一类消息
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim strIndexs As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            If Val(objRow.Record(C_病人Id).Value) = lng病人ID And Val(objRow.Record(C_主页Id).Value) = lng主页ID And objRow.Record(C_消息).Value = strNO Then
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

Private Sub imgCurPati_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'修改图标
    If Button = 2 Then
        Call Init图标菜单(1, Index + 1)
    End If
End Sub

Private Sub fraPageId_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'空白处弹出图标设置菜单
    If Button = 2 Then
        Call Init图标菜单(0)
    End If
End Sub

Private Sub lblCountThis_Click(Index As Integer)
'弹出选择器
    Dim strTmp As String
    Dim strIcon As String
    
    strTmp = lblCountThis(Index).Tag
    If strTmp = "" Then Exit Sub
    mrsPati汇总.Filter = "说明='" & Split(strTmp, "<Tab>")(0) & "' and 图形索引=" & Split(strTmp, "<Tab>")(1)
    If Not mrsPati汇总.EOF Then
        strTmp = mrsPati汇总!病人
        If strTmp <> "" Then
            Call LoadIconSelect(strTmp)
        End If
    End If
End Sub

Private Sub imgIconPati_Click(Index As Integer)
'弹出选择器
    Call lblCountThis_Click(Index)
End Sub

Private Sub ShowPati图标(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
'功能：显示当前病人的图标
'说明：每个病人要最多标记5类图标，查询出来后只需设置可见性即可
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim lngIconIdx As Long
    
    On Error GoTo errH
    
    '隐藏图标
    For i = 0 To imgCurPati.Count - 1
        imgCurPati(i).Visible = False
        imgCurPati(i).Tag = ""
    Next
    Set mPatiInfo.rs图标 = Nothing
    
    '先读缓存
    If Not mrsPatiNotes Is Nothing Then
        mrsPatiNotes.Filter = "病人id=" & lng病人ID & " and 主页id=" & lng主页ID
        If Not mrsPatiNotes.EOF Then
            Set rsTmp = zlDatabase.CopyNewRec(mrsPatiNotes)
        End If
    End If
    
    If rsTmp Is Nothing Then
        strSQL = "Select a.病人id, a.主页id, a.标记顺序,a.主题序号, nvl(a.主题病区ID,0) as 主题病区ID, a.标记序号, a.日期, Replace(b.说明, '|', '') as 说明, b.图形索引, b.有效天数, Floor(Sysdate - a.日期) As 实际天数" & _
            " From 病区标记记录 A, 病区标记内容 B,病案主页 c Where a.主题序号 = b.主题序号 And a.标记序号 = b.标记序号  And nvl(a.主题病区ID,0) = nvl(b.病区id,0) And (b.有效天数 = 0 Or (b.有效天数 > Floor(Sysdate - a.日期))) " & _
            " and a.病人id=c.病人id and a.主页id=c.主页id and a.病区id=c.当前病区id And a.病人id = [1] And a.主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    End If
    
    If Not rsTmp.EOF Then
        Set mPatiInfo.rs图标 = zlDatabase.CopyNewRec(rsTmp)
        For i = 1 To rsTmp.RecordCount
            If InStr(",1,2,3,", rsTmp!标记顺序) > 0 Then
                j = Val(rsTmp!标记顺序 & "") - 1
                imgCurPati(j).Visible = True
                lngIconIdx = Val(rsTmp!图形索引 & "") + 1
                If lngIconIdx > 0 And lngIconIdx <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                    Set imgCurPati(j).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(lngIconIdx).Picture
                End If
                imgCurPati(j).ToolTipText = rsTmp!说明 & ""
                imgCurPati(j).Tag = lngIconIdx
            End If
            rsTmp.MoveNext
        Next
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetAllPati图标(Optional ByVal intFun As Integer)
'功能：所有图标的初始化和卸载
'参数：0-加载，1-卸载
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To conIconAll - 1
        If intFun = 0 Then
            Load imgIconPati(i)
            Load lblCountThis(i)
            Set imgIconPati(i).Container = picIconPati
            Set lblCountThis(i).Container = picIconPati
            imgIconPati(i).ZOrder 1
            lblCountThis(i).ZOrder 1
            imgIconPati(i).Visible = False
            lblCountThis(i).Visible = False
        Else
            Unload imgIconPati(i)
            Unload lblCountThis(i)
            mstrList主题 = ""
        End If
    Next
    
End Sub

Private Sub ShowAllPati图标(ByVal strPatis As String)
'功能：显示当前所有病人的汇总图标,加载当前界面列表病人的图标
'参数：strPatis 格式："病人ID:主页ID,病人ID:主页ID,..."
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    Dim strParTable  As String
    Dim strTable As String
    Dim varArr As Variant
    Dim strPar As String
    Dim lngCnt As Long
    Dim str病人 As String
    Dim lngIconIdx As Long
    
    
    On Error GoTo errH
    
    If strPatis = "" Then
        '隐藏图标
        For i = 0 To conIconAll - 1
            imgIconPati(i).Visible = False
            lblCountThis(i).Visible = False
            imgIconPati(i).Tag = ""
            lblCountThis(i).Tag = ""
        Next
        picIconPati.Visible = False
        Call picPati_Resize
        Exit Sub
    End If
 
    strPatis = Mid(strPatis, 2)
    
    strParTable = "Select /*+cardinality(D,10)*/ d.C1, d.C2 From Table(f_Num2list2([1])) D"
    strTable = strParTable
    
    If Len(strPatis) >= 4000 Then
        varArr = Array()
        varArr = GetParTable(strPatis, strParTable, strTable)
    End If
    
    strSQL = "Select a.病人id, a.主页id, a.主题序号,a.标记顺序, nvl(a.主题病区ID,0) as 主题病区ID,a.标记序号, a.日期, Replace(b.说明, '|', '') as 说明, b.图形索引, b.有效天数, Floor(Sysdate - a.日期) As 实际天数" & _
        " From 病区标记记录 A, 病区标记内容 B,病案主页 c Where a.主题序号 = b.主题序号 And a.标记序号 = b.标记序号  And nvl(a.主题病区ID,0) = nvl(b.病区id,0) And (b.有效天数 = 0 Or (b.有效天数 > Floor(Sysdate - a.日期))) " & _
        " and a.病人id=c.病人id and a.主页id=c.主页id and a.病区id=c.当前病区id and (a.病人id,a.主页id) In (" & strTable & ")"
                
    If mrsPatiNotes Is Nothing Then
        If Len(strPatis) >= 4000 Then
            Set mrsPatiNotes = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
                CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        Else
            Set mrsPatiNotes = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatis)
        End If
    Else
        mrsPatiNotes.Filter = 0
    End If
        
    '初始化汇总图标记录集
    Set mrsPati汇总 = New ADODB.Recordset

    mrsPati汇总.Fields.Append "说明", adVarChar, 40000
    mrsPati汇总.Fields.Append "图形索引", adBigInt
    mrsPati汇总.Fields.Append "数量", adBigInt
    mrsPati汇总.Fields.Append "病人", adVarChar, 40000
    mrsPati汇总.CursorLocation = adUseClient
    mrsPati汇总.LockType = adLockOptimistic
    mrsPati汇总.CursorType = adOpenStatic
    mrsPati汇总.Open
    
    For i = 1 To mrsPatiNotes.RecordCount
         
        mrsPati汇总.Filter = "说明='" & mrsPatiNotes!说明 & "' and 图形索引=" & mrsPatiNotes!图形索引
        
        If Not mrsPati汇总.EOF Then
            If InStr(str病人, "," & mrsPatiNotes!病人ID & ":" & mrsPatiNotes!主页ID & ",") = 0 Then
                mrsPati汇总!病人 = mrsPati汇总!病人 & "," & mrsPatiNotes!病人ID & ":" & mrsPatiNotes!主页ID
                mrsPati汇总!数量 = Val(mrsPati汇总!数量 & "") + 1
                str病人 = str病人 & mrsPatiNotes!病人ID & ":" & mrsPatiNotes!主页ID & ","
            End If
        Else
            mrsPati汇总.AddNew
            mrsPati汇总!说明 = mrsPatiNotes!说明
            mrsPati汇总!图形索引 = mrsPatiNotes!图形索引
            mrsPati汇总!数量 = 1
            mrsPati汇总!病人 = mrsPatiNotes!病人ID & ":" & mrsPatiNotes!主页ID
            str病人 = "," & mrsPati汇总!病人 & ","
        End If
        mrsPati汇总.Update
        
        mrsPatiNotes.MoveNext
    Next
    mrsPati汇总.Filter = 0
    mrsPati汇总.Sort = "数量 desc,图形索引"
    
    '限制图标个数
    For i = 0 To mrsPati汇总.RecordCount - 1
        lngIconIdx = Val(mrsPati汇总!图形索引 & "") + 1
        If lngIconIdx > 0 And lngIconIdx <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
            Set imgIconPati(i).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(lngIconIdx).Picture
        End If
        lblCountThis(i).Caption = "(" & mrsPati汇总!数量 & ")"
        imgIconPati(i).Visible = True
        lblCountThis(i).Visible = True
        imgIconPati(i).ToolTipText = mrsPati汇总!说明
        lblCountThis(i).ToolTipText = mrsPati汇总!说明
        imgIconPati(i).Tag = lngIconIdx
        lblCountThis(i).Tag = mrsPati汇总!说明 & "<Tab>" & mrsPati汇总!图形索引
        lngCnt = lngCnt + 1
        mrsPati汇总.MoveNext
    Next
    
    For i = lngCnt To conIconAll - 1
        imgIconPati(i).Visible = False
        lblCountThis(i).Visible = False
    Next
    
    '设置位置
    Call picIconPati_Resize
    picIconPati.Visible = imgIconPati(0).Visible
    Call picPati_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Init图标信息(ByVal lng病区ID As Long)
'功能：初始化病人个性化图标标记
'提取当前病区设定的标注主题
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    mstrList主题 = ""
    strSQL = "Select Nvl(a.病区id,0) as 病区id, a.主题序号, a.标记序号, Replace(a.说明, '|', '') 说明, a.图形索引, a.有效天数,a.是否特殊" & vbNewLine & _
        "From 病区标记内容 a Where a.病区id Is Null Or a.病区id =[1]" & vbNewLine & _
        "Order By Nvl(a.病区id, 0), a.主题序号, a.标记序号"

    Set mrsNotes = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病区ID)
    
    For i = 1 To mrsNotes.RecordCount
        If Val(mrsNotes!标记序号 & "") = 0 Then
            mstrList主题 = mstrList主题 & "<TabB>" & mrsNotes!病区ID & "<TabA>" & mrsNotes!主题序号 & "<TabA>" & mrsNotes!说明
        End If
        mrsNotes.MoveNext
    Next
    mstrList主题 = Mid(mstrList主题, 7)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Init图标菜单(ByVal intType As Integer, Optional ByVal lng顺序号 As Long)
'功能：弹出图标设置菜单
'参数：intType 0-点空白处，1-修改某一个
'      lng顺序号 修改的第几个
'显示出所有标注主题并提供选择，注意排它性
    Dim int个性1 As Integer
    Dim int个性2 As Integer
    Dim int个性3 As Integer
    
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopup As CommandBarPopup
    Dim var主题 As Variant
    Dim varTmp As Variant
    Dim str已有标记录 As String
    Dim strBarName As String
    Dim i As Long
    Dim rsPati图标 As ADODB.Recordset
    Dim intFun As Integer '1-全显示，2-排它
    Dim int组数 As Integer
    Dim blnDo As Boolean
    Dim blnDel As Boolean
    Dim strIdxChk As String
    Dim strDelInfo As String
    Dim lngNew序号 As Long
    
    If mPatiInfo.病区ID = 0 Then Exit Sub
    If mintDeptView = 0 Then
        Call Init图标信息(mPatiInfo.病区ID)
    End If
    If mrsNotes Is Nothing Then Exit Sub
    If mrsNotes.RecordCount = 0 Then Exit Sub
    
    '只允许标记与列表相同的一次住院
    If mPatiInfo.病人ID & "," & mPatiInfo.主页ID <> mlng病人ID & "," & mlng主页ID Then Exit Sub
    
    If Not mPatiInfo.rs图标 Is Nothing Then
        Set rsPati图标 = mPatiInfo.rs图标
    End If
    
    If intType = 0 Then
        If rsPati图标 Is Nothing Then
            '全显示
            intFun = 1
            lngNew序号 = 1
        Else
            If rsPati图标.RecordCount < 3 Then
                intFun = 2
                lngNew序号 = 1
                rsPati图标.Sort = "标记顺序"
                For i = 1 To rsPati图标.RecordCount
                    If Val(rsPati图标!标记顺序 & "") = lngNew序号 Then
                        lngNew序号 = lngNew序号 + 1
                    End If
                    rsPati图标.MoveNext
                Next
            Else
                Exit Sub
            End If
        End If
    ElseIf intType = 1 Then
        intFun = 2
    End If
    
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    cbrPopupBar.Title = "标注设定"
    If mlngSource = 999 Then
        Call cbrPopupBar.SetIconSize(16, 16)
    Else
        Call cbrPopupBar.SetIconSize(24, 24)
    End If
    
    var主题 = Split(mstrList主题, "<TabB>")
    For i = 0 To UBound(var主题)
        varTmp = Split(var主题(i), "<TabA>")
        strBarName = varTmp(2)
        If intFun = 1 Or intFun = 2 Then
            blnDo = True
            If intFun = 2 Then
                rsPati图标.Filter = "主题病区ID=" & varTmp(0) & "  and 主题序号=" & varTmp(1)
                blnDo = rsPati图标.EOF
                If intType = 1 And Not blnDo Then
                    If lng顺序号 = Val(rsPati图标!标记顺序 & "") Then
                        strIdxChk = varTmp(0) & "," & varTmp(1) & "," & rsPati图标!图形索引 & "," & rsPati图标!说明
                        
                        strDelInfo = varTmp(1) & ",0,0," & varTmp(0)
                        blnDel = True
                        blnDo = True
                    End If
                End If
            End If
            
            If blnDo Then
                mrsNotes.Filter = "病区ID=" & varTmp(0) & "  and 主题序号=" & varTmp(1) & " And 标记序号>0"
                If mrsNotes.RecordCount <> 0 Then
                    int组数 = int组数 + 1
                    Set cbrPopup = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_标注1, strBarName)
                    If mlngSource = 999 Then
                        Call cbrPopup.CommandBar.SetIconSize(16, 16)
                    Else
                        Call cbrPopup.CommandBar.SetIconSize(24, 24)
                    End If
                    Do While Not mrsNotes.EOF
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_标注1 + mrsNotes.RecordCount, mrsNotes!说明)
                        cbrPopupItem.IconId = conMenu_图标 + mrsNotes!图形索引
                        
                        
                        If intType = 1 Then
                            cbrPopupItem.Parameter = "1," & mrsNotes!主题序号 & "," & mrsNotes!标记序号 & "," & lng顺序号 & "," & mrsNotes!病区ID
                            cbrPopupItem.Checked = (strIdxChk = mrsNotes!病区ID & "," & mrsNotes!主题序号 & "," & mrsNotes!图形索引 & "," & mrsNotes!说明)
                        Else
                            cbrPopupItem.Parameter = "0," & mrsNotes!主题序号 & "," & mrsNotes!标记序号 & "," & lngNew序号 & "," & mrsNotes!病区ID
                        End If
                        
                        
                        mrsNotes.MoveNext
                    Loop
                    If blnDel Then
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_标注1 + mrsNotes.RecordCount + 1, "清除标注")
                            cbrPopupItem.BeginGroup = True
                            cbrPopupItem.IconId = 3014
                        cbrPopupItem.Parameter = "2," & strDelInfo    '设置为-1,是因为image的索引从1开始,而ImageManager是从0开始
                    End If
                    blnDel = False
                End If
            End If
        End If
    Next
    mrsNotes.Filter = 0
    
    If int组数 <> 0 Then
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub picIconPati_Resize()
    Dim i As Long
    
    On Error Resume Next
    
    lblBJ.Left = 80
    lblBJ.Top = (picIconPati.Height - lblBJ.Height) / 2
    Call zlControl.SetPubCtrlPos(False, 0, lblBJ, 50, imgIconPati(0), 10, lblCountThis(0))
    For i = 1 To conIconAll - 1
        imgIconPati(i).Left = lblCountThis(i - 1).Left + lblCountThis(i - 1).Width + IIf(mbytSize = 0, 100, 150)
        imgIconPati(i).Top = imgIconPati(i - 1).Top
        lblCountThis(i).Left = imgIconPati(i).Left + imgIconPati(i).Width
        lblCountThis(i).Top = lblCountThis(i - 1).Top
    Next
End Sub

Private Sub SetPatiIconScale()
'功能：病人图标，大字体与小字体的对应
    Dim i As Long
    Dim strTmp As String
    Dim lngIconIdx As Long
    
    For i = 0 To 2
        lngIconIdx = Val(imgCurPati(i).Tag)
        If lngIconIdx <> 0 Then
            Set imgCurPati(i).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(lngIconIdx).Picture
        End If
    Next
    
    For i = 0 To conIconAll - 1
        lngIconIdx = Val(imgIconPati(i).Tag)
        If lngIconIdx <> 0 Then
            Set imgIconPati(i).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(lngIconIdx).Picture
        End If
    Next
    
    Call picIconPati_Resize
End Sub

Private Sub cmdFilterCancel_Click()
    picTBPati.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Call rptTBPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub rptTBPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptTBPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub rptTBPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdFilterCancel_Click
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptTBPati.Records.Count = 0 Then Exit Sub
    If rptTBPati.FocusedRow Is Nothing Then Exit Sub
    If rptTBPati.FocusedRow.Record Is Nothing Then Exit Sub
    
    picTBPati.Visible = False
    '定位病人
    Call LocatePati(rptTBPati.FocusedRow.Record.Tag)
End Sub

Private Sub LoadIconSelect(ByVal strPatis As String)
'功能：加载点击图标后的选择器列表
    Dim rsTmp As ADODB.Recordset
    Dim rsIconOther As ADODB.Recordset
    Dim strSQL As String
    Dim strTable As String
    Dim lngColor As Long, j As Long
    Dim lngloop As Long
    Dim objRow As ReportRow, blnSelect As Boolean
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    Dim i As Long
    Dim intIcon1 As Integer
    Dim intIcon2 As Integer
    Dim intIcon3 As Integer

    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    
    On Error GoTo errH
    
    strTable = "Select /*+cardinality(D,10)*/ d.C1, d.C2 From Table(f_Num2list2([1])) D"
    
    strSQL = "select a.病人id,a.主页id,a.姓名,lpad(a.出院病床,10,' ') AS 床号,a.住院号," & vbNewLine & _
        " Decode(a.入科时间,NULL,a.入院日期,a.入科时间) AS 入院日期 ,a.出院日期,a.病人类型 from 病案主页 a" & vbNewLine & _
        " where (a.病人id,a.主页id) In (" & strTable & ")" & _
        " order by 床号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatis)
    
    rptTBPati.Records.DeleteAll
 
    With rsTmp
        Do While Not .EOF
            Set objRecord = Me.rptTBPati.Records.Add()
            objRecord.Tag = CStr(!病人ID & "," & !主页ID)
            
            intIcon1 = -1
            intIcon2 = -1
            intIcon3 = -1
            
            If Not mrsPatiNotes Is Nothing Then
                mrsPatiNotes.Filter = "病人ID=" & !病人ID & " and 主页ID=" & !主页ID
                If Not mrsPatiNotes.EOF Then
                    For i = 1 To mrsPatiNotes.RecordCount
                        If mrsPatiNotes!标记顺序 = 1 Then
                            intIcon1 = mrsPatiNotes!图形索引
                        ElseIf mrsPatiNotes!标记顺序 = 2 Then
                            intIcon2 = mrsPatiNotes!图形索引
                        ElseIf mrsPatiNotes!标记顺序 = 3 Then
                            intIcon3 = mrsPatiNotes!图形索引
                        End If
                        mrsPatiNotes.MoveNext
                    Next
                End If
            End If
            
            '图标1
            Set objItem = objRecord.AddItem("")
            If intIcon1 > -1 Then
                objItem.Icon = intIcon1
            End If
  
            '图标2
            Set objItem = objRecord.AddItem("")
            If intIcon2 > -1 Then
                objItem.Icon = intIcon2
            End If
            
            '图标3
            Set objItem = objRecord.AddItem("")
            If intIcon3 > -1 Then
                objItem.Icon = intIcon3
            End If
            
            Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(!床号), 10))
            objItem.Caption = Trim(NVL(!床号, " "))
            objRecord.AddItem Val(!病人ID)
            objRecord.AddItem Val(!主页ID)
            objRecord.AddItem CStr(NVL(!姓名))
            Set objItem = objRecord.AddItem(CStr(NVL(!住院号)))
            objItem.Caption = NVL(!住院号, " ")
            
            Set objItem = objRecord.AddItem(Format(!入院日期, "yyyy-MM-dd"))
            objItem.Caption = Format(!入院日期, "yyyy-MM-dd")
            Set objItem = objRecord.AddItem(Format(!出院日期, "yyyy-MM-dd"))
            objItem.Caption = Format(!出院日期, "yyyy-MM-dd")
            
            Set objItem = objRecord.AddItem(NVL(!病人类型))
            objItem.Caption = NVL(!病人类型)
            
            .MoveNext
        Loop
    End With

    'picTBPati 调整坐标
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    rptTBPati.Populate '缺省不选中任何行
    picTBPati.Left = lngLeft - picIconPati.Width + 500
    picTBPati.Top = lngTop + picIconPati.Top + picIconPati.Height + 350
    If mbytSize = 0 Then
        picTBPati.Height = 5955
    Else
        picTBPati.Height = 6050
    End If
    picTBPati.Visible = True
    If rptTBPati.Visible Then rptTBPati.SetFocus
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetPatiIcon(ByVal strIconInfo As String)
'功能：更新病人图标
'参数： strIconInfo 过程参数信息
    '保存数据
    Dim strSQL As String
    Dim strSQLOther As String
    Dim varTmp As Variant
    Dim intType As Integer '0-增，1-改，2-删
    Dim lng顺序号 As Long
    Dim lng主题序号 As Long
    Dim lng主题病区 As Long
    Dim lng标记序号 As Long
    Dim blnTrans As Boolean
    
    On Error GoTo errH
       
    varTmp = Split(strIconInfo, ",")
    intType = varTmp(0)
    lng主题序号 = varTmp(1)
    lng标记序号 = varTmp(2)
    lng顺序号 = varTmp(3)
    lng主题病区 = varTmp(4)
    
    If intType = 1 Then
        mPatiInfo.rs图标.Filter = "标记顺序=" & lng顺序号
        If mPatiInfo.rs图标!主题病区ID & "," & mPatiInfo.rs图标!主题序号 <> lng主题病区 & "," & lng主题序号 Then
            strSQLOther = "ZL_病区标记记录_UPDATE(" & mPatiInfo.病区ID & "," & mPatiInfo.病人ID & "," & mPatiInfo.主页ID & "," & mPatiInfo.rs图标!主题序号 & ",0,0," & mPatiInfo.rs图标!主题病区ID & ")"
        End If
        strSQL = "ZL_病区标记记录_UPDATE(" & mPatiInfo.病区ID & "," & mPatiInfo.病人ID & "," & mPatiInfo.主页ID & "," & lng主题序号 & "," & lng标记序号 & "," & lng顺序号 & "," & IIf(0 = lng主题病区, "null", lng主题病区) & ")"
        
        gcnOracle.BeginTrans: blnTrans = True
        If strSQLOther <> "" Then Call zlDatabase.ExecuteProcedure(strSQLOther, Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTrans = False
    Else
        strSQL = "ZL_病区标记记录_UPDATE(" & mPatiInfo.病区ID & "," & mPatiInfo.病人ID & "," & mPatiInfo.主页ID & "," & lng主题序号 & "," & lng标记序号 & "," & lng顺序号 & "," & IIf(0 = lng主题病区, "null", lng主题病区) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    '更新卡片
    Set mrsPatiNotes = Nothing
    Call ShowAllPati图标(mstrAllPatis)
    Call ShowPati图标(mPatiInfo.病人ID, mPatiInfo.主页ID)
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetCriticalAdvice(ByVal lng记录ID As Long)
'功能：确认是危急值后弹出医嘱下达界面，刚才当前保存的医嘱与本次的记录进关联
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim objControl As Object
    
    On Error GoTo errH
    If lng记录ID = 0 Then Exit Sub
    strSQL = "select 1 from 病人危急值记录 a where a.id=[1] and a.是否危急值=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    
    If Not rsTmp.EOF Then
        '弹出下达医嘱的窗口
        If tbcSub.Tag <> "医嘱" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "医嘱" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                        Exit For
                    End If
                End If
            Next
        End If
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then
                objControl.Parameter = lng记录ID
                objControl.Execute
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

Private Sub ReadMsgAuto()
'功能：危急值消息处理自动弹出
    Dim i As Long
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim strNO As String
    Dim str业务 As String
    Dim lng消息ID As Long
    Dim blnRs As Boolean
    
    On Error GoTo errH
    
    For i = i To rptNotify.Rows.Count - 1
        With rptNotify.Rows(i)
            If Not .GroupRow Then
                strNO = .Record(C_消息).Value
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
                    lng病人ID = Val(.Record(C_病人Id).Value)
                    lng主页ID = Val(.Record(C_主页Id).Value)
                    str业务 = .Record(C_业务).Value
                    lng消息ID = Val(.Record(C_Id).Value)
                    blnRs = ReadMsg(lng病人ID, lng主页ID, strNO, str业务, lng消息ID)
                End If
            End If
        End With
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
